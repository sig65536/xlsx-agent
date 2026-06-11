# Review Style Guide

A Python service (FastAPI + openpyxl) that edits spreadsheets through a local LLM.
A request is turned into openpyxl operations, applied in a sandbox, previewed as a diff,
and only written out after explicit approval.

When reviewing, prioritize the following.

## Sandbox safety (highest priority)
LLM-generated code runs in a separate process with deliberately restricted capabilities.
Do not let these weaken:
- Imports are limited to a whitelist (`_SAFE_IMPORT_ROOTS` in `app/agent.py`).
- Dangerous builtins (`open`, `eval`, `exec`, `compile`, `getattr`, `setattr`,
  `__import__`, `globals`, ...) are removed.
- An AST pre-check rejects dunder-attribute access and forbidden names.

Flag any change that could let generated code reach the filesystem, network, `os`/`sys`,
`subprocess`, or otherwise escape the sandbox.

## Spreadsheet integrity (high priority)
- `keep_vba` must follow the file extension (`.xlsm` only). Using it on `.xlsx` produces a
  macro-enabled content-type that Excel rejects as corrupt.
- The saved workbook must be validated (zip parts present, content-type matches the
  extension, openpyxl can reload it) before it replaces the result file.
- The original upload must stay untouched until the user approves.

## Correctness
Watch for data loss, off-by-one errors in row/column handling, merged-cell writes to a
non-anchor cell, and formula-vs-value confusion.

## Tests
New behavior should come with tests. The suite runs with `pytest`.

## Style and severity
- Follow PEP 8 and match the surrounding code (type hints, small helpers, comment language).
- Treat sandbox-escape and file-corruption risks as high/critical; treat pure style nits as low.
