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

# import 可能な標準ライブラリのルート（サブモジュールも許可）。
# I/O・システムアクセスを伴うモジュールは含めない。
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
}

# openpyxl は「丸ごと」は許可しない（`from openpyxl import load_workbook` で
# サーバー上の任意ブックを読み出し print() 経由で内容を漏洩できてしまうため）。
# ファイルI/Oを伴わない書式・ユーティリティ系サブモジュールだけを許可する。
_SAFE_OPENPYXL_MODULES = {
    "openpyxl.styles",
    "openpyxl.utils",
    "openpyxl.utils.cell",
    "openpyxl.comments",
    "openpyxl.formatting",
    "openpyxl.formatting.rule",
    "openpyxl.worksheet.table",
    "openpyxl.chart",
}

# 安全なモジュール経由でも到達させてはいけない危険モジュール／属性名。
# precheck で ast.Name / ast.Attribute の両方として遮断する
# （例: `import random; random.os.system(...)` のような属性チェーン脱獄を防ぐ）。
_DANGEROUS_NAMES = {
    "os",
    "sys",
    "subprocess",
    "socket",
    "shutil",
    "ctypes",
    "importlib",
    "posix",
    "nt",
    "builtins",
    "platform",
    "multiprocessing",
    "threading",
    "signal",
    "pickle",
    "marshal",
    "inspect",
    "pty",
    "system",
    "popen",
    "spawn",
}

# サンドボックスで公開するビルトインの「明示的許可リスト」。
# dir(builtins) からの除外方式だと site が注入する license/copyright/help 等が
# 残り、`license._Printer__filenames` 経由で任意ファイルを開示できてしまうため、
# 必要なものだけを列挙する方式にする（type/object/super/getattr 等は意図的に除外）。
_ALLOWED_BUILTINS = {
    "abs", "all", "any", "ascii", "bin", "bool", "bytearray", "bytes",
    "callable", "chr", "dict", "divmod", "enumerate", "filter", "float",
    "format", "frozenset", "hasattr", "hash", "hex", "int", "isinstance",
    "issubclass", "iter", "len", "list", "map", "max", "min", "next", "oct",
    "ord", "pow", "print", "range", "repr", "reversed", "round", "set",
    "slice", "sorted", "str", "sum", "tuple", "zip",
}
# try/except で使う例外クラス
_ALLOWED_EXCEPTIONS = {
    "BaseException", "Exception", "ArithmeticError", "AssertionError",
    "AttributeError", "FloatingPointError", "IndexError", "KeyError",
    "LookupError", "NameError", "NotImplementedError", "OverflowError",
    "RuntimeError", "StopIteration", "TypeError", "ValueError",
    "ZeroDivisionError",
}

# 事前ASTチェックで弾く名前参照（ビルトイン除去と二重防御）。
# 危険モジュール名(_DANGEROUS_NAMES)も含め、ast.Name / ast.Attribute の両方で遮断する。
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
} | _DANGEROUS_NAMES

MAX_OBSERVATION_CHARS = 1500


def _is_allowed_import(name: str) -> bool:
    root = name.split(".", 1)[0]
    if root in _SAFE_IMPORT_ROOTS:
        return True
    if name in _SAFE_OPENPYXL_MODULES:
        return True
    return any(name.startswith(f"{mod}.") for mod in _SAFE_OPENPYXL_MODULES)


def _safe_import(name, globals=None, locals=None, fromlist=(), level=0):
    import builtins as _b

    if level != 0 or not _is_allowed_import(name):
        raise ImportError(f"このサンドボックスでは '{name}' の import は許可されていません")
    # fromlist 経由のエイリアス脱獄を遮断する。
    # 例: `from random import _os as x` は許可モジュール(random)に見えるが、
    # 危険モジュール os を別名 x に束縛してしまい AST 検査をすり抜ける。
    for item in fromlist or ():
        if isinstance(item, str) and item.lstrip("_") in _DANGEROUS_NAMES:
            raise ImportError(
                f"禁止されたモジュール/名前の import は許可されていません: {item}"
            )
    return _b.__import__(name, globals, locals, fromlist, level)


def _make_safe_builtins() -> dict:
    import builtins as _b

    safe = {
        name: getattr(_b, name)
        for name in (_ALLOWED_BUILTINS | _ALLOWED_EXCEPTIONS)
        if hasattr(_b, name)
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
    from app.common import JobError

    try:
        tree = ast.parse(code)
    except SyntaxError as exc:
        raise JobError("AGENT_CODE_SYNTAX", f"生成コードの構文エラー: {exc}") from exc

    def _is_forbidden(identifier: str) -> bool:
        # ダンダー(__builtins__ 等)は一括禁止。先頭の "_" を剥がした名前も照合し、
        # `re._sys` / `random._os` のようなプライベート別名経由の脱獄を防ぐ。
        if identifier.startswith("__"):
            return True
        if identifier in _FORBIDDEN_NAMES:
            return True
        return identifier.lstrip("_") in _FORBIDDEN_NAMES

    for node in ast.walk(tree):
        if isinstance(node, ast.Attribute):
            # `.compile` 属性は re.compile 等の安全用途なので許可（ビルトイン
            # compile の Name 参照は引き続き禁止）。
            if node.attr != "compile" and _is_forbidden(node.attr):
                raise JobError(
                    "AGENT_CODE_REJECTED",
                    f"禁止された属性へのアクセスです: {node.attr}",
                )
        if isinstance(node, ast.Name) and _is_forbidden(node.id):
            raise JobError(
                "AGENT_CODE_REJECTED", f"禁止された名前の参照: {node.id}"
            )


def _agent_worker(input_path: str, sheet_name: str, conn) -> None:
    """spawnされる子プロセス本体。wb/wsを保持し、コマンドを逐次処理する。

    ※ `app.main` ではなく `app.common`（副作用なし）から部品を import する。
      `app.main` を子プロセスで読むと `create_app()` が走り JobService の
      スレッドが余計に起動してしまうため。

    各ステップはトランザクション的に扱う：実行が失敗したら、その途中変更を
    破棄して直近の成功状態（committed スナップショット）へロールバックする。
    これにより失敗ステップの中途半端な変更が保存されない。
    """
    try:
        from openpyxl import load_workbook

        from app.common import (
            _keep_vba_for,
            _safe_set_merged_value,
            validate_excel_file,
        )

        keep_vba = _keep_vba_for(input_path)

        def _build_namespace(workbook):
            worksheet = workbook[sheet_name]
            return {
                "__builtins__": _make_safe_builtins(),
                "wb": workbook,
                "ws": worksheet,
                "helpers": {
                    "safe_set_merged_value": lambda merged_range, value: (
                        _safe_set_merged_value(worksheet, merged_range, value)
                    )
                },
            }

        def _snapshot(workbook) -> bytes:
            buffer = io.BytesIO()
            workbook.save(buffer)
            return buffer.getvalue()

        wb = load_workbook(input_path, keep_vba=keep_vba, data_only=False)
        namespace: dict[str, Any] = _build_namespace(wb)
        committed = _snapshot(wb)  # 直近の「成功状態」のスナップショット
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
                committed = _snapshot(namespace["wb"])  # 成功 → コミット
                conn.send({"ok": True, "stdout": buf.getvalue()[-MAX_OBSERVATION_CHARS:]})
            except Exception:
                error_text = traceback.format_exc()[-MAX_OBSERVATION_CHARS:]
                # 失敗ステップの途中変更を破棄し、直近コミット状態へロールバック
                try:
                    wb = load_workbook(
                        io.BytesIO(committed), keep_vba=keep_vba, data_only=False
                    )
                    namespace = _build_namespace(wb)
                except Exception:
                    pass
                conn.send(
                    {
                        "ok": False,
                        "stdout": buf.getvalue()[-MAX_OBSERVATION_CHARS:],
                        "error": error_text,
                    }
                )
        elif cmd == "save":
            try:
                target = Path(msg["path"])
                tmp_path = target.with_name(f"{target.stem}.tmp{target.suffix}")
                namespace["wb"].save(tmp_path)
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
        from app.common import JobError

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
        ready = self._recv("AGENT_INIT_FAILED", "サンドボックス初期化中にプロセスが予期せず終了しました")
        if not ready.get("ok"):
            self.close()
            raise JobError(
                "AGENT_INIT_FAILED",
                f"サンドボックス初期化に失敗しました: {ready.get('error', '')}",
            )

    def _send(self, message: dict, fail_message: str) -> None:
        from app.common import JobError

        try:
            self._conn.send(message)
        except OSError as exc:
            self.close()
            raise JobError("AGENT_SANDBOX_CRASHED", fail_message, retryable=True) from exc

    def _recv(self, error_code: str, fail_message: str) -> dict:
        from app.common import JobError

        try:
            return self._conn.recv()
        except (EOFError, OSError) as exc:
            self.close()
            raise JobError(error_code, fail_message, retryable=True) from exc

    def run(self, code: str, timeout: int) -> dict:
        from app.common import JobError

        self._send({"cmd": "exec", "code": code}, "サンドボックスへのコマンド送信に失敗しました")
        if not self._conn.poll(timeout):
            self.close()
            raise JobError(
                "EXEC_TIMEOUT", "ステップ実行がタイムアウトしました", retryable=True
            )
        return self._recv(
            "AGENT_SANDBOX_CRASHED",
            "ステップ実行中にサンドボックスプロセスが予期せず終了しました",
        )

    def save(self, path: str, timeout: int) -> dict:
        from app.common import JobError

        self._send({"cmd": "save", "path": path}, "サンドボックスへの保存コマンド送信に失敗しました")
        if not self._conn.poll(timeout):
            self.close()
            raise JobError(
                "EXEC_TIMEOUT", "保存がタイムアウトしました", retryable=True
            )
        return self._recv(
            "AGENT_SANDBOX_CRASHED", "保存中にサンドボックスプロセスが予期せず終了しました"
        )

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
    from app.common import JobError

    sandbox = AgentSandbox(working_path, sheet_name)
    transcript: list[dict[str, Any]] = []
    applied = 0
    completed = False
    try:
        for step in range(1, max_steps + 1):
            text = llm.agent_step(summary, instruction, transcript)
            kind, code = parse_action(text)
            if kind == "done":
                completed = True
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

        # DONE に到達せずステップ上限で打ち切った場合は、未完成の部分編集を
        # そのまま保存せず、再試行可能なエラーにする。
        if not completed:
            raise JobError(
                "AGENT_STEP_LIMIT",
                f"ステップ上限({max_steps})に達しました。指示を分割するか "
                "XLSX_AGENT_MAX_STEPS を増やして再実行してください。",
                retryable=True,
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
