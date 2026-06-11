import ast
import io
import json
import multiprocessing as mp
import os
import queue
import re
import secrets
import shutil
import threading
import time
import traceback
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import load_workbook

STATIC_DIR = Path(__file__).resolve().parent / "static"

ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}
MAX_FILE_SIZE_BYTES = 20 * 1024 * 1024
JOB_TTL_SECONDS = 3600
LLM_MAX_RETRY = 3
SUMMARY_FORMULA_MAX_ROWS = 200
SUMMARY_FORMULA_MAX_COLS = 50
PREVIEW_MAX_CHANGED_CELLS = 500
PREVIEW_FORMULA_NOTE = "数式セルは文字列比較です。実値はExcelで再計算されます。"


class JobStatus:
    QUEUED = "queued"
    ANALYZING = "analyzing"
    GENERATING = "generating"
    CHECKING = "checking"
    EXECUTING = "executing"
    PREVIEW_READY = "preview_ready"
    APPROVED = "approved"
    DONE = "done"
    ERROR = "error"


class JobError(Exception):
    def __init__(self, error_code: str, message: str, retryable: bool = False):
        self.error_code = error_code
        self.message = message
        self.retryable = retryable
        super().__init__(message)


@dataclass
class Job:
    job_id: str
    instruction: str
    sheet_name: str | None
    created_at: float
    work_dir: Path
    source_path: Path
    status: str = JobStatus.QUEUED
    error_code: str | None = None
    message: str | None = None
    retryable: bool = False
    preview: dict[str, Any] | None = None
    result_path: Path | None = None
    download_token: str | None = None


class LLMClient:
    def __init__(self) -> None:
        self.endpoint = os.getenv(
            "OLLAMA_ENDPOINT", "http://localhost:11434/api/generate"
        )
        self.model = os.getenv("OLLAMA_MODEL", "gemma4:e4b")
        self.timeout = int(os.getenv("LLM_TIMEOUT_SECONDS", "60"))
        self._resolved_model: str | None = None

    @property
    def _tags_endpoint(self) -> str:
        base = self.endpoint.rsplit("/api/", 1)[0]
        return f"{base}/api/tags"

    def _list_models(self) -> list[str]:
        req = Request(self._tags_endpoint, method="GET")
        with urlopen(req, timeout=self.timeout) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
        return [m.get("name", "") for m in payload.get("models", []) if m.get("name")]

    def resolve_model(self) -> str:
        """設定モデル名を、実際にOllamaへ入っているタグ名に解決する。

        `ollama pull gemma4` は `gemma4:latest`、`ollama pull gemma4:e4b` は
        `gemma4:e4b` という名前で保存されるため、設定値とインストール名が
        食い違うと404になる。ベース名(gemma4)が一致するタグへ自動で寄せる。
        """
        if self._resolved_model:
            return self._resolved_model
        try:
            installed = self._list_models()
        except (URLError, OSError, ValueError):
            return self.model  # 取得失敗時は設定値のまま試す
        if not installed or self.model in installed:
            self._resolved_model = self.model
            return self.model
        base = self.model.split(":", 1)[0]
        for candidate in (f"{base}:e4b", f"{base}:latest", base):
            if candidate in installed:
                self._resolved_model = candidate
                return candidate
        same_base = [m for m in installed if m.split(":", 1)[0] == base]
        self._resolved_model = same_base[0] if same_base else self.model
        return self._resolved_model

    def generate_code(
        self, summary: dict[str, Any], instruction: str, feedback: str = ""
    ) -> str:
        prompt = (
            "あなたはExcel編集用Pythonコード生成器です。"
            "出力は```python ... ```のコードブロック1つだけにしてください。"
            "説明文、思考過程、Markdown本文、箇条書きは禁止です。"
            "使用可能な変数は wb, ws, helpers のみです。"
            "import / open / exec / eval / ファイルI/O / ネットワーク / OS操作は禁止です。\n"
            f"sheet_summary={json.dumps(summary, ensure_ascii=False)}\n"
            f"instruction={instruction}\n"
            f"feedback={feedback}\n"
        )
        model = self.resolve_model()
        body = {
            "model": model,
            "stream": False,
            "prompt": prompt,
            "options": {"temperature": 0.1},
        }
        data = json.dumps(body).encode("utf-8")
        req = Request(
            self.endpoint,
            data=data,
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        try:
            with urlopen(req, timeout=self.timeout) as resp:
                payload = json.loads(resp.read().decode("utf-8"))
        except HTTPError as exc:
            detail = ""
            try:
                detail = exc.read().decode("utf-8")[:300]
            except Exception:  # pragma: no cover - 補助的な詳細取得のみ
                pass
            if exc.code == 404:
                raise JobError(
                    "LLM_MODEL_NOT_FOUND",
                    f"Ollamaにモデル '{model}' が見つかりません。"
                    f"サーバーで `ollama pull {self.model}` を実行してください。{detail}",
                ) from exc
            raise JobError(
                "LLM_HTTP_ERROR",
                f"LLM呼び出しがHTTP {exc.code} を返しました: {detail}",
                retryable=True,
            ) from exc
        except URLError as exc:
            raise JobError(
                "LLM_TIMEOUT", f"LLM呼び出しに失敗しました: {exc}", retryable=True
            ) from exc
        text = payload.get("response", "")
        match = re.search(r"```(?:python)?\s*(.*?)```", text, flags=re.DOTALL)
        if match:
            return match.group(1).strip()
        return text.strip()


class CodeChecker:
    FORBIDDEN_NAMES = {
        "__import__",
        "eval",
        "exec",
        "compile",
        "open",
        "os",
        "sys",
        "subprocess",
        "socket",
        "shutil",
        "input",
        "breakpoint",
        "globals",
        "locals",
        "vars",
        "dir",
        "getattr",
        "setattr",
        "delattr",
        "type",
        "object",
        "super",
        "memoryview",
    }
    FORBIDDEN_NODES = (
        ast.FunctionDef,
        ast.AsyncFunctionDef,
        ast.ClassDef,
        ast.Lambda,
        ast.With,
        ast.AsyncWith,
        ast.Try,
        ast.Global,
        ast.Nonlocal,
        ast.Delete,
        ast.Raise,
        ast.Yield,
        ast.YieldFrom,
        ast.Await,
    )

    def validate(self, code: str, non_anchor_cells: set[str]) -> None:
        try:
            tree = ast.parse(code)
        except SyntaxError as exc:
            raise JobError(
                "CODE_CHECK_FAILED", f"生成コードの構文エラー: {exc}"
            ) from exc

        for node in ast.walk(tree):
            if isinstance(node, (ast.Import, ast.ImportFrom)):
                raise JobError(
                    "CODE_CHECK_FAILED",
                    "生成コードに禁止された操作（import）が含まれていました",
                )
            if isinstance(node, self.FORBIDDEN_NODES):
                raise JobError(
                    "CODE_CHECK_FAILED",
                    f"生成コードに禁止された構文が含まれていました: {type(node).__name__}",
                )
            if isinstance(node, ast.Name) and node.id in self.FORBIDDEN_NAMES:
                raise JobError(
                    "CODE_CHECK_FAILED",
                    f"禁止された名前参照が含まれていました: {node.id}",
                )
            if isinstance(node, ast.Attribute) and node.attr.startswith("__"):
                raise JobError(
                    "CODE_CHECK_FAILED", "ダンダー属性アクセスは禁止されています"
                )
            if isinstance(node, ast.Call):
                if (
                    isinstance(node.func, ast.Name)
                    and node.func.id in self.FORBIDDEN_NAMES
                ):
                    raise JobError(
                        "CODE_CHECK_FAILED",
                        f"禁止された関数呼び出しが含まれていました: {node.func.id}",
                    )
            if isinstance(node, (ast.Assign, ast.AugAssign)):
                targets = (
                    node.targets if isinstance(node, ast.Assign) else [node.target]
                )
                for target in targets:
                    addr = self._extract_ws_address(target)
                    if addr and addr in non_anchor_cells:
                        raise JobError(
                            "MERGED_CELL_CONFLICT",
                            f"結合セル非アンカーへの書き込みは禁止です: {addr}",
                        )

    def _extract_ws_address(self, node: ast.AST) -> str | None:
        if (
            isinstance(node, ast.Subscript)
            and isinstance(node.value, ast.Name)
            and node.value.id == "ws"
        ):
            if isinstance(node.slice, ast.Constant) and isinstance(
                node.slice.value, str
            ):
                return node.slice.value.upper()
        if (
            isinstance(node, ast.Call)
            and isinstance(node.func, ast.Attribute)
            and node.func.attr == "cell"
        ):
            if isinstance(node.func.value, ast.Name) and node.func.value.id == "ws":
                row = None
                column = None
                if len(node.args) >= 2:
                    row = self._const_int(node.args[0])
                    column = self._const_int(node.args[1])
                for kw in node.keywords:
                    if kw.arg == "row":
                        row = self._const_int(kw.value)
                    if kw.arg in {"column", "col"}:
                        column = self._const_int(kw.value)
                if row and column:
                    return f"{self._col_letters(column)}{row}"
        return None

    @staticmethod
    def _const_int(node: ast.AST) -> int | None:
        if isinstance(node, ast.Constant) and isinstance(node.value, int):
            return node.value
        return None

    @staticmethod
    def _col_letters(column: int) -> str:
        out = ""
        while column:
            column, rem = divmod(column - 1, 26)
            out = chr(65 + rem) + out
        return out


def _close_workbook(wb) -> None:
    vba_archive = getattr(wb, "vba_archive", None)
    if vba_archive is not None:
        vba_archive.close()
        wb.vba_archive = None
    wb.close()


def _safe_set_merged_value(ws, merged_range: str, value: Any) -> None:
    ws.unmerge_cells(merged_range)
    anchor = merged_range.split(":", 1)[0]
    ws[anchor] = value
    ws.merge_cells(merged_range)


def _summary_value(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return str(value)


def _row_values(ws, row: int, min_col: int, max_col: int) -> list[Any]:
    return [
        _summary_value(ws.cell(row=row, column=col).value)
        for col in range(min_col, max_col + 1)
    ]


def _non_empty_count(values: list[Any]) -> int:
    return sum(value not in (None, "") for value in values)


def _summarize_sheet(ws) -> dict[str, Any]:
    sample_max_row = min(max(ws.max_row, 1), 30)
    sample_max_col = 20
    sample_values = [
        {"row": row, "values": _row_values(ws, row, 1, sample_max_col)}
        for row in range(1, sample_max_row + 1)
    ]

    merged_cells = []
    for rng in ws.merged_cells.ranges:
        min_col, min_row, _max_col, _max_row = rng.bounds
        anchor = f"{CodeChecker._col_letters(min_col)}{min_row}"
        merged_cells.append(
            {
                "range": str(rng),
                "anchor": anchor,
                "value": _summary_value(ws[anchor].value),
            }
        )

    hidden_rows = [
        idx for idx, dimension in ws.row_dimensions.items() if dimension.hidden
    ]
    hidden_columns = [
        col for col, dimension in ws.column_dimensions.items() if dimension.hidden
    ]

    formulas: list[str] = []
    for row in ws.iter_rows(
        min_row=1,
        max_row=min(ws.max_row, SUMMARY_FORMULA_MAX_ROWS),
        min_col=1,
        max_col=min(ws.max_column, SUMMARY_FORMULA_MAX_COLS),
    ):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                formulas.append(cell.coordinate)

    header_candidates = []
    for sample_row in sample_values:
        values = sample_row["values"]
        string_count = sum(
            1 for value in values if isinstance(value, str) and value.strip()
        )
        if string_count >= 2 and _non_empty_count(values) >= 2:
            header_candidates.append({"row": sample_row["row"], "values": values})

    table_like_ranges = []
    for candidate in header_candidates[:5]:
        row = candidate["row"]
        values = candidate["values"]
        non_empty_cols = [
            idx + 1 for idx, value in enumerate(values) if value not in (None, "")
        ]
        if not non_empty_cols:
            continue
        min_col = min(non_empty_cols)
        max_col = max(non_empty_cols)
        end_row = row
        for scan_row in range(row + 1, min(ws.max_row, row + 200) + 1):
            row_values = _row_values(ws, scan_row, min_col, max_col)
            if _non_empty_count(row_values) == 0:
                break
            end_row = scan_row
        table_like_ranges.append(
            {
                "range": f"{CodeChecker._col_letters(min_col)}{row}:{CodeChecker._col_letters(max_col)}{end_row}",
                "header_row": row,
                "columns": values[min_col - 1 : max_col],
            }
        )

    return {
        "sheet_name": ws.title,
        "max_row": ws.max_row,
        "max_col": ws.max_column,
        "merged_ranges": [str(rng) for rng in ws.merged_cells.ranges],
        "merged_cells": merged_cells,
        "hidden_rows": hidden_rows,
        "hidden_columns": hidden_columns,
        "protected": bool(ws.protection.sheet),
        "sample_values": sample_values,
        "header_candidates": header_candidates,
        "table_like_ranges": table_like_ranges,
        "formula_cells": formulas,
    }


def _non_anchor_cells(ws) -> set[str]:
    blocked: set[str] = set()
    for rng in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = rng.bounds
        anchor = f"{CodeChecker._col_letters(min_col)}{min_row}"
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                address = f"{CodeChecker._col_letters(col)}{row}"
                if address != anchor:
                    blocked.add(address)
    return blocked


def _sandbox_runner(path: str, sheet_name: str, code: str, result_queue: Any) -> None:
    try:
        wb = load_workbook(path, keep_vba=True, data_only=False)
        ws = wb[sheet_name]
        safe_globals = {
            "__builtins__": {
                "range": range,
                "len": len,
                "min": min,
                "max": max,
                "sum": sum,
                "enumerate": enumerate,
                "int": int,
                "float": float,
                "str": str,
                "bool": bool,
                "list": list,
                "dict": dict,
                "set": set,
                "tuple": tuple,
                "abs": abs,
                "any": any,
                "all": all,
                "zip": zip,
            }
        }
        safe_locals = {
            "wb": wb,
            "ws": ws,
            "helpers": {
                "safe_set_merged_value": lambda merged_range, value: (
                    _safe_set_merged_value(ws, merged_range, value)
                )
            },
        }
        exec(compile(code, "<generated>", "exec"), safe_globals, safe_locals)
        wb.save(path)
        _close_workbook(wb)
        result_queue.put({"ok": True})
    except Exception:
        result_queue.put({"ok": False, "error": traceback.format_exc()})


def _exec_in_sandbox(
    path: Path, sheet_name: str, code: str, timeout_sec: int = 30
) -> None:
    ctx = mp.get_context("spawn")
    q = ctx.Queue()
    proc = ctx.Process(target=_sandbox_runner, args=(str(path), sheet_name, code, q))
    proc.start()
    proc.join(timeout=timeout_sec)
    if proc.is_alive():
        proc.terminate()
        proc.join(5)
        raise JobError(
            "EXEC_TIMEOUT", "生成コードの実行がタイムアウトしました", retryable=True
        )
    result = q.get() if not q.empty() else {"ok": False, "error": "no result"}
    if not result.get("ok"):
        raise JobError(
            "EXEC_RUNTIME_ERROR",
            f"実行エラー: {result.get('error', '')}",
            retryable=True,
        )


def _create_preview(
    before_path: Path, after_path: Path, sheet_name: str
) -> dict[str, Any]:
    before = load_workbook(before_path, keep_vba=True, data_only=False)
    after = load_workbook(after_path, keep_vba=True, data_only=False)
    ws_before = before[sheet_name]
    ws_after = after[sheet_name]
    max_row = max(ws_before.max_row, ws_after.max_row)
    max_col = max(ws_before.max_column, ws_after.max_column)
    changed_cells = []
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            c1 = ws_before.cell(row=row, column=col)
            c2 = ws_after.cell(row=row, column=col)
            if c1.value != c2.value or c1.style_id != c2.style_id:
                coord = f"{CodeChecker._col_letters(col)}{row}"
                changed_cells.append(
                    {
                        "cell": coord,
                        "before": c1.value,
                        "after": c2.value,
                        "before_formula": c1.value
                        if isinstance(c1.value, str) and c1.value.startswith("=")
                        else None,
                        "after_formula": c2.value
                        if isinstance(c2.value, str) and c2.value.startswith("=")
                        else None,
                        "style_changed": c1.style_id != c2.style_id,
                    }
                )
    _close_workbook(before)
    _close_workbook(after)
    return {
        "sheet_name": sheet_name,
        "changed_cell_count": len(changed_cells),
        "changed_cells": changed_cells[:PREVIEW_MAX_CHANGED_CELLS],
        "notes": [PREVIEW_FORMULA_NOTE],
    }


class JobService:
    def __init__(self, root_dir: Path, llm: LLMClient | None = None) -> None:
        self.root_dir = root_dir
        self.root_dir.mkdir(parents=True, exist_ok=True)
        self.llm = llm or LLMClient()
        self.checker = CodeChecker()
        self.jobs: dict[str, Job] = {}
        self.download_tokens: dict[str, Path] = {}
        self.lock = threading.Lock()
        self.q: queue.Queue[str] = queue.Queue()
        self.worker = threading.Thread(target=self._worker_loop, daemon=True)
        self.worker.start()
        self.cleaner = threading.Thread(target=self._cleanup_loop, daemon=True)
        self.cleaner.start()

    def create_job(
        self, upload: UploadFile, instruction: str, sheet_name: str | None
    ) -> str:
        ext = Path(upload.filename or "").suffix.lower()
        if ext not in ALLOWED_EXTENSIONS:
            raise JobError("UNSUPPORTED_FORMAT", "対応形式は .xlsx / .xlsm のみです")
        payload = upload.file.read(MAX_FILE_SIZE_BYTES + 1)
        if len(payload) > MAX_FILE_SIZE_BYTES:
            raise JobError("FILE_TOO_LARGE", "ファイルサイズが上限を超えています")
        try:
            uploaded_wb = load_workbook(
                io.BytesIO(payload), keep_vba=True, data_only=False
            )
            _close_workbook(uploaded_wb)
        except Exception as exc:
            raise JobError(
                "UNSUPPORTED_FORMAT", f"Excelファイルを開けませんでした: {exc}"
            ) from exc

        job_id = uuid.uuid4().hex
        work_dir = self.root_dir / job_id
        work_dir.mkdir(parents=True, exist_ok=True)
        source_path = work_dir / f"input{ext}"
        source_path.write_bytes(payload)

        with self.lock:
            self.jobs[job_id] = Job(
                job_id=job_id,
                instruction=instruction,
                sheet_name=sheet_name,
                created_at=time.time(),
                work_dir=work_dir,
                source_path=source_path,
            )
        self.q.put(job_id)
        return job_id

    def get_job(self, job_id: str) -> Job:
        with self.lock:
            job = self.jobs.get(job_id)
        if not job:
            raise KeyError(job_id)
        return job

    def approve_job(self, job_id: str) -> str:
        with self.lock:
            job = self.jobs.get(job_id)
            if not job:
                raise KeyError(job_id)
            if job.status != JobStatus.PREVIEW_READY or not job.result_path:
                raise JobError("INVALID_STATE", "プレビュー準備完了後のみ承認できます")
            job.status = JobStatus.APPROVED
            token = secrets.token_urlsafe(24)
            self.download_tokens[token] = job.result_path
            job.download_token = token
            job.status = JobStatus.DONE
            return token

    def pop_download_path(self, token: str) -> Path:
        with self.lock:
            path = self.download_tokens.pop(token, None)
        if path is None or not path.exists():
            raise KeyError(token)
        return path

    def _worker_loop(self) -> None:
        while True:
            job_id = self.q.get()
            try:
                self._process(job_id)
            finally:
                self.q.task_done()

    def _process(self, job_id: str) -> None:
        job = self.get_job(job_id)
        try:
            job.status = JobStatus.ANALYZING
            wb = load_workbook(job.source_path, keep_vba=True, data_only=False)
            ws = (
                wb[job.sheet_name]
                if job.sheet_name and job.sheet_name in wb.sheetnames
                else wb.active
            )
            sheet_name = ws.title
            summary = _summarize_sheet(ws)
            blocked_cells = _non_anchor_cells(ws)
            _close_workbook(wb)

            generated_code = ""
            feedback = ""
            validation_errors: list[str] = []
            for _ in range(LLM_MAX_RETRY):
                job.status = JobStatus.GENERATING
                generated_code = self.llm.generate_code(
                    summary, job.instruction, feedback=feedback
                )
                job.status = JobStatus.CHECKING
                try:
                    self.checker.validate(generated_code, blocked_cells)
                    break
                except JobError as err:
                    feedback = f"{err.error_code}: {err.message}"
                    validation_errors.append(feedback)
            else:
                reason = (
                    "; ".join(validation_errors)
                    if validation_errors
                    else "コード生成に失敗しました（検証詳細なし）"
                )
                raise JobError(
                    "CODE_CHECK_FAILED", f"コード再生成上限に達しました: {reason}"
                )

            job.status = JobStatus.EXECUTING
            result_path = job.work_dir / f"result{job.source_path.suffix}"
            shutil.copy2(job.source_path, result_path)
            _exec_in_sandbox(result_path, sheet_name, generated_code)
            job.preview = _create_preview(job.source_path, result_path, sheet_name)
            job.result_path = result_path
            job.status = JobStatus.PREVIEW_READY
        except JobError as err:
            job.status = JobStatus.ERROR
            job.error_code = err.error_code
            job.message = err.message
            job.retryable = err.retryable
        except Exception as err:  # pragma: no cover
            job.status = JobStatus.ERROR
            job.error_code = "INTERNAL_ERROR"
            job.message = str(err)
            job.retryable = False

    def _cleanup_loop(self) -> None:
        while True:
            time.sleep(60)
            now = time.time()
            with self.lock:
                expired = [
                    job_id
                    for job_id, job in self.jobs.items()
                    if now - job.created_at > JOB_TTL_SECONDS
                ]
            for job_id in expired:
                with self.lock:
                    job = self.jobs.pop(job_id, None)
                    if job and job.download_token:
                        self.download_tokens.pop(job.download_token, None)
                if job and job.work_dir.exists():
                    shutil.rmtree(job.work_dir, ignore_errors=True)


def _to_job_response(job: Job) -> dict[str, Any]:
    data = {
        "job_id": job.job_id,
        "status": job.status,
    }
    if job.status == JobStatus.ERROR:
        data.update(
            {
                "error_code": job.error_code,
                "message": job.message,
                "retryable": job.retryable,
            }
        )
    if job.preview:
        data["preview"] = job.preview
    return data


def create_app(job_service: JobService | None = None) -> FastAPI:
    app = FastAPI(title="xlsx-agent", version="0.1")
    service = job_service or JobService(Path(os.getenv("JOB_ROOT", "./data/jobs")))

    # ブラウザ（別PC）からのアクセスを許可する。
    # 既定は全許可。社内LAN限定にしたい場合は CORS_ORIGINS にカンマ区切りで指定。
    cors_origins = [
        origin.strip()
        for origin in os.getenv("CORS_ORIGINS", "*").split(",")
        if origin.strip()
    ]
    app.add_middleware(
        CORSMiddleware,
        allow_origins=cors_origins or ["*"],
        allow_credentials=False,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    @app.get("/healthz", include_in_schema=False)
    async def healthz() -> dict[str, str]:
        return {"status": "ok", "model": service.llm.model}

    @app.get("/", include_in_schema=False)
    async def index() -> FileResponse:
        return FileResponse(STATIC_DIR / "index.html")

    @app.post("/jobs")
    async def create_job(
        file: UploadFile = File(...),
        instruction: str = Form(...),
        sheet_name: str | None = Form(default=None),
    ) -> dict[str, str]:
        try:
            job_id = service.create_job(file, instruction, sheet_name)
            return {"job_id": job_id, "status": JobStatus.QUEUED}
        except JobError as err:
            raise HTTPException(
                status_code=400,
                detail={
                    "status": JobStatus.ERROR,
                    "error_code": err.error_code,
                    "message": err.message,
                    "retryable": err.retryable,
                },
            ) from err

    @app.get("/jobs/{job_id}")
    async def get_job(job_id: str) -> dict[str, Any]:
        try:
            return _to_job_response(service.get_job(job_id))
        except KeyError as err:
            raise HTTPException(status_code=404, detail="job not found") from err

    @app.post("/jobs/{job_id}/approve")
    async def approve(job_id: str) -> dict[str, str]:
        try:
            token = service.approve_job(job_id)
            return {"job_id": job_id, "download_url": f"/download/{token}"}
        except KeyError as err:
            raise HTTPException(status_code=404, detail="job not found") from err
        except JobError as err:
            raise HTTPException(status_code=400, detail=err.message) from err

    @app.get("/download/{token}")
    async def download(token: str):
        try:
            path = service.pop_download_path(token)
            return FileResponse(
                path,
                filename=f"edited{path.suffix}",
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except KeyError as err:
            raise HTTPException(status_code=404, detail="token not found") from err

    if STATIC_DIR.is_dir():
        app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

    return app


app = create_app()
