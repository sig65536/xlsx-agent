# xlsx-agent

ローカルLLM（Ollama / gemma4:e4b）でExcelを自然言語編集するツール。

ユーザーは**ブラウザからアクセスするだけ**。アップロードされたファイルは**サーバーPC**で
解析・編集され、差分プレビューを確認・承認してからダウンロードします。元ファイルは
承認するまで書き換わりません。

> 👉 操作手順の詳細は **[USAGE.md](USAGE.md)** を参照してください。

## 構成

```
[ユーザーPC: ブラウザ]  ──HTTP──▶  [サーバーPC]
   ・ファイルをアップロード              ・FastAPI (このアプリ)
   ・指示を入力                          ・Ollama + gemma4:e4b （ローカルLLM）
   ・差分プレビューを確認・承認          ・openpyxl でExcelを編集
   ・編集済みファイルをDL                ・サンドボックスでコード実行
```

## サーバーPCでのセットアップ

前提: Python 3.10+ と [Ollama](https://ollama.com) がインストール済み。

```bash
# 1) モデルを取得（初回のみ・約9.6GBのDL）
ollama pull gemma4:e4b
ollama serve            # 常駐していなければ起動（systemd等でもOK）

# 2) このリポジトリを取得して起動
git clone https://github.com/sig65536/xlsx-agent.git
cd xlsx-agent
./start.sh              # Windowsは start.bat をダブルクリック
```

`start.sh` / `start.bat` が venv 作成・依存インストール・サーバー起動（`0.0.0.0:8000`）まで行います。
Ollama の稼働とモデルの有無もチェックして警告します。

## ユーザーの使い方

サーバーと同じLAN（またはTailscale等）から、ブラウザで以下にアクセスするだけ:

```
http://<サーバーPCのIP>:8000
```

ファイルを選び、指示文（例:「C列の売上を合計して最終行の下に合計行を追加」）を入力して
「編集を依頼」→ 差分プレビューを確認 →「承認してダウンロード」。

## 環境変数

| 変数 | 既定値 | 説明 |
|---|---|---|
| `OLLAMA_ENDPOINT` | `http://localhost:11434/api/generate` | Ollama の generate API |
| `OLLAMA_MODEL` | `gemma4:e4b` | 使用モデル名 |
| `LLM_TIMEOUT_SECONDS` | `60` | LLM応答のタイムアウト |
| `JOB_ROOT` | `./data/jobs` | ジョブ作業ディレクトリ |
| `CORS_ORIGINS` | `*` | 許可オリジン（カンマ区切り。特定オリジンに限定する場合に指定） |
| `PORT` / `HOST` | `8000` / `0.0.0.0` | 待ち受けポート・ホスト |
| `XLSX_AGENT_MODE` | `agent` | 編集方式。`agent`（ReActループ）/ `oneshot`（単発生成） |
| `XLSX_AGENT_MAX_STEPS` | `6` | エージェントの最大ステップ数 |
| `XLSX_AGENT_STEP_TIMEOUT` | `30` | 1ステップの実行タイムアウト（秒） |
| `XLSX_AGENT_DISABLE_NETWORK` | `1` | worker のネットワーク発信を遮断（`0`で無効化） |
| `XLSX_AGENT_WORKER_CPU_SEC` | `300` | worker のCPU時間上限（POSIXのみ） |
| `XLSX_AGENT_WORKER_FSIZE_MB` | `100` | worker のファイル書込サイズ上限（POSIXのみ） |
| `XLSX_AGENT_WORKER_NOFILE` | `256` | worker のファイルディスクリプタ上限（POSIXのみ） |
| `XLSX_AGENT_WORKER_MEM_MB` | （無制限） | worker のメモリ上限MB（POSIXのみ・指定時のみ適用） |
| `XLSX_AGENT_WORKER_UID` / `_GID` | （無効） | root実行時に降格する低権限UID/GID（POSIXのみ） |

### サンドボックスのOSレベル隔離

エージェントの worker 子プロセスには、AST/import/builtins のガードに加えて
OSレベルの隔離を適用しています（万一それらを抜けても実害を出さないため）。

- **ネットワーク遮断**: worker からのアウトバウンド通信を全面禁止（全OS対応）。
- **環境変数スクラブ**: OS標準以外の環境変数（APIキー等）を worker から削除（全OS対応）。
- **リソース上限**: CPU時間・ファイル書込・FD・プロセス数の上限（POSIXのみ）。
- **権限ドロップ**: root実行時に低権限ユーザーへ降格（POSIX・`*_UID/_GID`指定時）。

> **Windows サーバの場合**: ネットワーク遮断と環境変数スクラブは有効です。CPU/メモリ等の
> リソース上限は POSIX 専用のため Windows では無効（代わりにステップ単位の実行タイムアウトで
> 暴走を抑止）。より厳格な隔離が必要なら **WSL2 / コンテナ上での実行**を推奨します。

### 編集方式（エージェント）

既定の `agent` モードでは、LLMが **「コードを書く→実行→結果(print/エラー)を観測→次の手」**
を繰り返すReActループで編集します。途中のエラーを見て自己修正でき、複数ステップの作業や
シート間の操作にも対応します。

サンドボックスは安全のため `os` / `sys` / ファイル / ネットワークへのアクセスを遮断しますが、
**Excel編集に必要な範囲は開放**しており、`datetime` / `re` / `openpyxl.styles` 等の import が
可能です。これにより**太字・色・罫線・日付**などの書式設定も行えます。

従来の単発方式に戻したい場合は `XLSX_AGENT_MODE=oneshot` を指定してください。

### モデル名について

`gemma4:e4b` は Ollama 上では「latest」エイリアスでもあります。`ollama pull gemma4`
で取得すると `gemma4:latest` という名前で保存されるため、設定値と食い違って 404
（model not found）になることがあります。本アプリは Ollama の `/api/tags` を参照し、
**ベース名 `gemma4` が一致する実際のタグ（`gemma4:e4b` / `gemma4:latest` 等）へ自動で
寄せる**ため、どちらの名前で pull していても動作します。見つからない場合は
`LLM_MODEL_NOT_FOUND` エラーで pull コマンドを案内します。

## API

- `POST /jobs` : Excelファイル + 指示文を受け付け、`job_id`を返す
- `GET /jobs/{job_id}` : ジョブ状態とプレビューを返す
- `POST /jobs/{job_id}/approve` : 承認してダウンロードURLを返す
- `GET /download/{token}` : ワンタイムトークンで結果をダウンロード
- `GET /` : ブラウザUI（`app/static/index.html`）
- `GET /healthz` : 稼働確認（解決済みモデル名を含む）

対応形式は `.xlsx` / `.xlsm`（`keep_vba=True` でVBAは保持のみ）です。マクロは実行せず、内容も変更しません。

## テスト

```bash
pip install -e .[test]
pytest
```

## セキュリティ上の注意

本MVPのLLM生成コード実行は、別プロセス + ASTチェック + restricted builtins による制限付き実行です。
Docker / VM / 別OSユーザーによる完全なサンドボックスではありません。

外部公開環境では使用しないでください。
利用範囲はローカルPCまたは信頼できるネットワーク内に限定してください。
機密ファイルを扱う場合は、必ず原本バックアップを保持してください。
