#!/usr/bin/env bash
# xlsx-agent サーバー起動スクリプト（サーバーPC / Linux 用）
#
# 前提:
#   - Python 3.10+ がインストール済み
#   - Ollama が起動済みで、使用モデルが pull 済み
#       ollama serve            # 別ターミナル or systemd で常駐
#       ollama pull gemma4-e4b:latest  # 初回のみ（数GBのDL）
#
# モデル名などの設定は config.env を編集（ここ一箇所で切替）。
#
# 使い方:
#   ./start.sh            # 0.0.0.0:8000 で起動。ユーザーは http://<サーバーIP>:8000 にアクセス
#   PORT=9000 ./start.sh  # ポート変更
set -euo pipefail

DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$DIR"

# 設定は config.env 一箇所で切り替え（モデル名など）。実環境変数があれば優先。
if [ -f config.env ]; then
  set -a
  # shellcheck disable=SC1091
  . ./config.env
  set +a
fi

PORT="${PORT:-8000}"
HOST="${HOST:-0.0.0.0}"
OLLAMA_ENDPOINT="${OLLAMA_ENDPOINT:-http://localhost:11434/api/generate}"
OLLAMA_BASE="${OLLAMA_ENDPOINT%/api/generate}"
OLLAMA_MODEL="${OLLAMA_MODEL:-gemma4-e4b:latest}"

# --- venv 準備 ---
if [ ! -d ".venv" ]; then
  echo "[setup] 仮想環境を作成します (.venv)"
  python3 -m venv .venv
fi
# shellcheck disable=SC1091
source .venv/bin/activate
echo "[setup] 依存パッケージをインストール/更新します"
pip install --quiet --upgrade pip
pip install --quiet -e .

# --- Ollama 疎通チェック（警告のみ・自動DLはしない） ---
if command -v curl >/dev/null 2>&1; then
  if curl -sf "${OLLAMA_BASE}/api/tags" >/dev/null 2>&1; then
    if curl -sf "${OLLAMA_BASE}/api/tags" | grep -q "${OLLAMA_MODEL%%:*}"; then
      echo "[ok] Ollama 稼働中・モデル ${OLLAMA_MODEL} を確認"
    else
      echo "[warn] Ollama は稼働中ですが ${OLLAMA_MODEL} が見つかりません。先に実行してください:"
      echo "         ollama pull ${OLLAMA_MODEL}"
    fi
  else
    echo "[warn] Ollama (${OLLAMA_BASE}) に接続できません。別ターミナルで 'ollama serve' を起動してください。"
  fi
fi

# --- サーバー起動 ---
export OLLAMA_ENDPOINT OLLAMA_MODEL
echo "[run] http://${HOST}:${PORT}  (ブラウザからアクセス) / model=${OLLAMA_MODEL}"
exec uvicorn app.main:app --host "${HOST}" --port "${PORT}"
