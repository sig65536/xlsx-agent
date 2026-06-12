# AGENTS.md

Guidance for AI agents (e.g. Codex code review and coding tasks) in this repository.

## Project
A local-LLM-powered spreadsheet editing service:
- `app/main.py` — FastAPI backend (job queue, summary, preview, approval, download).
- `app/agent.py` — a ReAct editing agent that runs LLM-proposed openpyxl code step by step
  in a sandboxed subprocess.
- `app/static/` — the browser UI.

A user uploads a spreadsheet and an instruction; the server edits a copy with openpyxl and
returns it after a diff preview is approved. The original upload is never modified in place.

## How to test
```bash
pip install -e .[test]
pytest
```
Keep the suite green and add tests for new behavior.

## Review focus (flag P0/P1 only)
- **Sandbox safety**: generated code runs with a restricted import whitelist
  (`_SAFE_IMPORT_ROOTS`) and a builtins denylist, plus an AST pre-check. Never allow access to
  the filesystem, network, `os`/`sys`, `subprocess`, `eval`/`exec`/`open`, or dunder-attribute
  escapes from generated code.
- **File integrity**: `keep_vba` must stay tied to the extension (`.xlsm` only); the saved
  workbook must be validated (zip parts + content-type vs extension + openpyxl reload) before
  it replaces the output, so the result is never a file Excel refuses to open.
- **Data safety**: avoid silent data loss; the original upload stays untouched until approval.

## Conventions
- Python 3.10+, PEP 8, type hints. Keep changes small and consistent with the surrounding code.
- The default LLM model id (`gemma4-e4b:latest`) is intentional — do not change it.
- Prefer focused, high-signal review comments over style nitpicks.
