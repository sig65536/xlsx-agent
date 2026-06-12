@echo off
REM xlsx-agent server launcher (Windows)
REM ASCII-only on purpose: Japanese text in a .bat garbles under the console
REM codepage (CP932 etc.). Keep this file ASCII so it shows correctly anywhere.
REM
REM Prerequisites:
REM   - Python 3.10+ installed (python on PATH)
REM   - Ollama running, with the model pulled:
REM       ollama pull gemma4-e4b:latest
REM
REM Edit config.env to change the model name etc. in ONE place.
REM
REM Usage:  start.bat   then open http://<server-ip>:8000 in a browser
setlocal
cd /d "%~dp0"

REM Load config.env (KEY=VALUE lines; ignore blanks and lines starting with #)
if exist "config.env" (
  for /f "usebackq eol=# tokens=1,* delims==" %%a in ("config.env") do (
    if not "%%a"=="" set "%%a=%%b"
  )
)

if "%PORT%"=="" set PORT=8000
if "%HOST%"=="" set HOST=0.0.0.0
if "%OLLAMA_MODEL%"=="" set OLLAMA_MODEL=gemma4-e4b:latest

if not exist ".venv" (
  echo [setup] Creating virtual environment (.venv)
  python -m venv .venv
)
call .venv\Scripts\activate.bat

echo [setup] Installing/updating dependencies
python -m pip install --quiet --upgrade pip
python -m pip install --quiet -e .

echo [run] http://%HOST%:%PORT%  (open in a browser) / model=%OLLAMA_MODEL%
echo       Make sure Ollama is running and "ollama pull %OLLAMA_MODEL%" was done.
uvicorn app.main:app --host %HOST% --port %PORT%
endlocal
