@echo off
REM xlsx-agent サーバー起動スクリプト（会社のサーバーPC / Windows 用）
REM
REM 前提:
REM   - Python 3.10+ がインストール済み（py または python が使えること）
REM   - Ollama が起動済みで、モデル gemma4:e4b が pull 済み
REM       ollama pull gemma4:e4b   （初回のみ・数GBのDL）
REM
REM 使い方:  start.bat   → http://<サーバーIP>:8000 にブラウザでアクセス
setlocal
cd /d "%~dp0"

if "%PORT%"=="" set PORT=8000
if "%HOST%"=="" set HOST=0.0.0.0
if "%OLLAMA_MODEL%"=="" set OLLAMA_MODEL=gemma4:e4b

if not exist ".venv" (
  echo [setup] 仮想環境を作成します (.venv)
  python -m venv .venv
)
call .venv\Scripts\activate.bat

echo [setup] 依存パッケージをインストール/更新します
python -m pip install --quiet --upgrade pip
python -m pip install --quiet -e .

echo [run] http://%HOST%:%PORT%  (ブラウザからアクセス) / model=%OLLAMA_MODEL%
echo       Ollama が起動済みで `ollama pull %OLLAMA_MODEL%` 済みであることを確認してください。
uvicorn app.main:app --host %HOST% --port %PORT%
endlocal
