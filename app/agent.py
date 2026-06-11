"""ReActスタイルのExcel編集エージェント。

LLMが提案するopenpyxlコードを「永続サンドボックス子プロセス」で1ステップずつ実行し、
print()や例外を観測としてLLMへ戻して次の手を促す。複数ステップ・自己修正に対応する。

権限方針:
    元データはユーザーPC側にあり、サーバーが触るのはアップロードされた複製のため、
    ファイル破損リスクは低い。一方でサーバー自体のRCEは防ぐ必要があるため、
    - import はホワイトリスト(_SAFE_IMPORT_ROOTS)のみ許可（datetime/re/openpyxl.styles等）
    - os/sys/subprocess/socket/open/eval/exec 等はビルトインから除去
    という形で「Excel編集に必要な範囲だけ」開放している。
"""

import ast
import io
import multiprocessing as mp
import traceback
from contextlib import redirect_stdout
from pathlib import Path
from typing import Any

# import 可能なトップレベルモジュール（これ以外の import は ImportError にする）
_SAFE_IMPORT_ROOTS = {
    "datetime",
    "re",
    "math",
    "decimal",
    "fractions",
    "statistics",
    "calendar",
    "collections",
    "string",
    "itertools",
    "functools",
    "json",
    "copy",
    "random",
    "textwrap",
    "unicodedata",
    "openpyxl",
}

# ビルトインから除去する危険な名前（システムアクセス・脱獄経路）
_DENY_BUILTINS = {
    "open",
    "eval",
    "exec",
    "compile",
    "input",
    "breakpoint",
    "memoryview",
    "globals",
    "locals",
    "vars",
    "getattr",
    "setattr",
    "delattr",
}

# 事前ASTチェックで弾く名前参照（ビルトイン除去と二重防御）
_FORBIDDEN_NAMES = {
    "eval",
    "exec",
    "compile",
    "open",
    "input",
    "breakpoint",
    "getattr",
    "setattr",
    "delattr",
    "globals",
    "locals",
    "vars",
    "__import__",
}

MAX_OBSERVATION_CHARS = 1500


def _safe_import(name, globals=None, locals=None, fromlist=(), level=0):
    import builtins as _b

    root = name.split(".", 1)[0]
    if level == 0 and root in _SAFE_IMPORT_ROOTS:
        return _b.__import__(name, globals, locals, fromlist, level)
    raise ImportError(f"このサンドボックスでは '{name}' の import は許可されていません")


def _make_safe_builtins() -> dict:
    import builtins as _b

    safe = {
        name: getattr(_b, name)
        for name in dir(_b)
        if not name.startswith("_") and name not in _DENY_BUILTINS and hasattr(_b, name)
    }
    safe["__import__"] = _safe_import
    # class 文を許可するために必要（脱獄経路ではない）
    safe["__build_class__"] = _b.__build_class__
    safe["__name__"] = "__agent_sandbox__"
    return safe


def precheck_step_code(code: str) -> None:
    """exec前の軽量ASTチェック。明白な脱獄（ダンダー属性・禁止名）を弾く。

    import 自体はここでは弾かず、実行時の _safe_import がホワイトリスト判定する
    （モデルが import エラーを観測して自己修正できるようにするため）。
    """
    from app.main import JobError

    try:
        tree = ast.parse(code)
    except SyntaxError as exc:
        raise JobError("AGENT_CODE_SYNTAX", f"生成コードの構文エラー: {exc}") from exc
    for node in ast.walk(tree):
        if isinstance(node, ast.Attribute) and node.attr.startswith("__"):
            raise JobError(
                "AGENT_CODE_REJECTED",
                f"ダンダー属性アクセスは禁止されています: {node.attr}",
            )
        if isinstance(node, ast.Name) and node.id in _FORBIDDEN_NAMES:
            raise JobError(
                "AGENT_CODE_REJECTED", f"禁止された名前の参照: {node.id}"
            )


def _agent_worker(input_path: str, sheet_name: str, conn) -> None:
    """spawnされる子プロセス本体。wb/wsを保持し、コマンドを逐次処理する。"""
    try:
        from openpyxl import load_workbook

        from app.main import _keep_vba_for, _safe_set_merged_value

        wb = load_workbook(
            input_path, keep_vba=_keep_vba_for(input_path), data_only=False
        )
        ws = wb[sheet_name]
        namespace: dict[str, Any] = {
            "__builtins__": _make_safe_builtins(),
            "wb": wb,
            "ws": ws,
            "helpers": {
                "safe_set_merged_value": lambda merged_range, value: (
                    _safe_set_merged_value(ws, merged_range, value)
                )
            },
        }
        conn.send({"ok": True})
    except Exception:
        try:
            conn.send({"ok": False, "error": traceback.format_exc()[-MAX_OBSERVATION_CHARS:]})
        except Exception:
            pass
        return

    while True:
        try:
            msg = conn.recv()
        except EOFError:
            break
        cmd = msg.get("cmd")
        if cmd == "exec":
            buf = io.StringIO()
            try:
                with redirect_stdout(buf):
                    exec(compile(msg["code"], "<agent-step>", "exec"), namespace)
                conn.send({"ok": True, "stdout": buf.getvalue()[-MAX_OBSERVATION_CHARS:]})
            except Exception:
                conn.send(
                    {
                        "ok": False,
                        "stdout": buf.getvalue()[-MAX_OBSERVATION_CHARS:],
                        "error": traceback.format_exc()[-MAX_OBSERVATION_CHARS:],
                    }
                )
        elif cmd == "save":
            try:
                from app.main import validate_excel_file

                target = Path(msg["path"])
                tmp_path = target.with_name(f"{target.stem}.tmp{target.suffix}")
                wb.save(tmp_path)
                validate_excel_file(tmp_path)
                tmp_path.replace(target)
                conn.send({"ok": True})
            except Exception as exc:
                error_code = getattr(exc, "error_code", "EXCEL_SAVE_VALIDATION_FAILED")
                message = getattr(exc, "message", None) or traceback.format_exc()[
                    -MAX_OBSERVATION_CHARS:
                ]
                conn.send({"ok": False, "error_code": error_code, "error": message})
        elif cmd == "close":
            break
    try:
        conn.close()
    except Exception:
        pass


class AgentSandbox:
    """親プロセス側のサンドボックス制御。1つの子プロセスにコマンドを送る。"""

    def __init__(self, input_path: Path, sheet_name: str, init_timeout: int = 30) -> None:
        from app.main import JobError

        ctx = mp.get_context("spawn")
        self._conn, child_conn = ctx.Pipe()
        self._proc = ctx.Process(
            target=_agent_worker,
            args=(str(input_path), sheet_name, child_conn),
            daemon=True,
        )
        self._proc.start()
        child_conn.close()
        if not self._conn.poll(init_timeout):
            self.close()
            raise JobError("AGENT_INIT_FAILED", "サンドボックス初期化がタイムアウトしました")
        ready = self._conn.recv()
        if not ready.get("ok"):
            self.close()
            raise JobError(
                "AGENT_INIT_FAILED",
                f"サンドボックス初期化に失敗しました: {ready.get('error', '')}",
            )

    def run(self, code: str, timeout: int) -> dict:
        from app.main import JobError

        self._conn.send({"cmd": "exec", "code": code})
        if not self._conn.poll(timeout):
            self.close()
            raise JobError(
                "EXEC_TIMEOUT", "ステップ実行がタイムアウトしました", retryable=True
            )
        return self._conn.recv()

    def save(self, path: str, timeout: int) -> dict:
        from app.main import JobError

        self._conn.send({"cmd": "save", "path": path})
        if not self._conn.poll(timeout):
            self.close()
            raise JobError(
                "EXEC_TIMEOUT", "保存がタイムアウトしました", retryable=True
            )
        return self._conn.recv()

    def close(self) -> None:
        try:
            if self._proc.is_alive():
                try:
                    self._conn.send({"cmd": "close"})
                except Exception:
                    pass
                self._proc.join(timeout=3)
            if self._proc.is_alive():
                self._proc.terminate()
                self._proc.join(timeout=3)
        except Exception:
            pass
        finally:
            try:
                self._conn.close()
            except Exception:
                pass


def parse_action(text: str) -> tuple[str, str]:
    """LLM応答を (種別, コード) に解釈する。種別は 'code' か 'done'。"""
    import re

    match = re.search(r"```(?:python)?\s*(.*?)```", text, flags=re.DOTALL)
    if match and match.group(1).strip():
        return "code", match.group(1).strip()
    if "DONE" in text.upper():
        return "done", ""
    # コードもDONEも無い場合は、無限ループを避けるため done 扱いにする
    return "done", ""


def run_agent(
    working_path: Path,
    sheet_name: str,
    instruction: str,
    summary: dict[str, Any],
    llm,
    max_steps: int = 6,
    step_timeout: int = 30,
) -> dict:
    """ReActループ本体。working_path を直接編集する。失敗時 JobError。"""
    from app.main import JobError

    sandbox = AgentSandbox(working_path, sheet_name)
    transcript: list[dict[str, Any]] = []
    applied = 0
    try:
        for step in range(1, max_steps + 1):
            text = llm.agent_step(summary, instruction, transcript)
            kind, code = parse_action(text)
            if kind == "done":
                break
            try:
                precheck_step_code(code)
            except JobError as rejected:
                transcript.append(
                    {
                        "step": step,
                        "code": code,
                        "observation": f"REJECTED: {rejected.message}",
                    }
                )
                continue
            result = sandbox.run(code, timeout=step_timeout)
            if result.get("ok"):
                applied += 1
                stdout = (result.get("stdout") or "").strip()
                observation = "OK" + (f"\n{stdout}" if stdout else "")
            else:
                observation = "ERROR:\n" + (
                    result.get("error") or result.get("stdout") or "unknown error"
                )
            transcript.append(
                {
                    "step": step,
                    "code": code,
                    "observation": observation[:MAX_OBSERVATION_CHARS],
                }
            )

        if applied == 0:
            raise JobError(
                "AGENT_NO_CHANGES",
                "エージェントが有効な編集を生成できませんでした",
                retryable=True,
            )
        save_result = sandbox.save(str(working_path), timeout=step_timeout)
        if not save_result.get("ok"):
            raise JobError(
                save_result.get("error_code", "EXCEL_SAVE_VALIDATION_FAILED"),
                f"保存に失敗しました: {save_result.get('error', '')}",
            )
    finally:
        sandbox.close()
    return {"steps": len(transcript), "applied": applied, "transcript": transcript}
