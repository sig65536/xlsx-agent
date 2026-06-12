# xlsx-agent

ローカルLLM（Ollama / gemma4-e4b:latest）でExcelを自然言語編集するツール。

ユーザーは**ブラウザからアクセスするだけ**。アップロードされたファイルは**サーバーPC**で
解析・編集され、差分プレビューを確認・承認してからダウンロードします。元ファイルは
承認するまで書き換わりません。

> 👉 操作手順の詳細は **[USAGE.md](USAGE.md)** を参照してください。

## 構成

```
[ユーザーPC: ブラウザ]  ──HTTP──▶  [サーバーPC]
   ・ファイルをアップロード              ・FastAPI (このアプリ)
   ・指示を入力                          ・Ollama + gemma4-e4b:latest （ローカルLLM）
   ・差分プレビューを確認・承認          ・openpyxl でExcelを編集
   ・編集済みファイルをDL                ・サンドボックスでコード実行
```

## サーバーPCでのセットアップ

前提: Python 3.10+ と [Ollama](https://ollama.com) がインストール済み。

```bash
# 1) モデルを取得（初回のみ・約9.6GBのDL）
ollama pull gemma4-e4b:latest
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

**チャット形式（既定・`/`）**: ファイルをアップロードし、チャットで指示を重ねて編集します。
各ターンの差分がその場で表示され、「↶ 1手戻す」でUndo、いつでもダウンロード可能。
「さっきの合計を赤字にして」のように**前の文脈を引き継いだ追加指示**ができます。
ヘッダの **thinking** トグルで推論モードを切替（精度↑/速度↓）。元ファイルは変更されません。

**単発フォーム（`/classic`）**: ファイルを選び、指示文を入力 →「編集を依頼」→
差分プレビューを確認 →「承認してダウンロード」。

## 環境変数

以下は環境変数で指定できますが、**`config.env`（リポジトリ直下）に書けば一箇所でまとめて
設定**できます（→「モデル名の設定」参照）。

| 変数 | 既定値 | 説明 |
|---|---|---|
| `OLLAMA_ENDPOINT` | `http://127.0.0.1:11434/api/generate` | Ollama の generate API |
| `OLLAMA_MODEL` | `gemma4-e4b:latest` | 使用モデル名 |
| `LLM_TIMEOUT_SECONDS` | `300` | LLM応答のタイムアウト |
| `JOB_ROOT` | `./data/jobs` | ジョブ作業ディレクトリ |
| `CORS_ORIGINS` | `*` | 許可オリジン（カンマ区切り。特定オリジンに限定する場合に指定） |
| `PORT` / `HOST` | `8000` / `0.0.0.0` | 待ち受けポート・ホスト |
| `XLSX_AGENT_THINK` | `0` | thinking(推論)モード。`1`で有効（精度↑/速度↓、対応モデルのみ） |
| `SESSION_ROOT` | `./data/sessions` | チャットセッションの作業ディレクトリ |
| `XLSX_AGENT_MODE` | `agent` | 編集方式。`agent`（ReActループ）/ `oneshot`（単発生成） |
| `XLSX_AGENT_MAX_STEPS` | `6` | エージェントの最大ステップ数 |
| `XLSX_AGENT_STEP_TIMEOUT` | `30` | 1ステップの実行タイムアウト（秒） |
| `XLSX_AGENT_DISABLE_NETWORK` | `1` | worker のネットワーク発信を遮断（`0`で無効化） |
| `XLSX_AGENT_WORKER_CPU_SEC` | `300` | worker のCPU時間上限（POSIXのみ） |
| `XLSX_AGENT_WORKER_FSIZE_MB` | `100` | worker のファイル書込サイズ上限（POSIXのみ） |
| `XLSX_AGENT_WORKER_NOFILE` | `256` | worker のファイルディスクリプタ上限（POSIXのみ） |
| `XLSX_AGENT_WORKER_MEM_MB` | （無制限） | worker のメモリ上限MB（POSIXのみ・指定時のみ適用） |

### サンドボックスのOSレベル隔離

エージェントの worker 子プロセスには、AST/import/builtins のガードに加えて
OSレベルの隔離を適用しています（万一それらを抜けても実害を出さないため）。

- **ネットワーク遮断**: worker からのアウトバウンド通信（DNS含む）を全面禁止（全OS対応）。
- **環境変数スクラブ**: OS標準以外の環境変数（APIキー等）を worker から削除（全OS対応）。
- **リソース上限**: CPU時間・ファイル書込・FD・プロセス数の上限（POSIXのみ）。

> **Windows サーバの場合**: ネットワーク遮断と環境変数スクラブは有効です。CPU/メモリ等の
> リソース上限は POSIX 専用のため Windows では無効（代わりにステップ単位の実行タイムアウトで
> 暴走を抑止）。より厳格な隔離（権限分離・FS制限）が必要なら、サービス自体を**専用の
> 低権限ユーザー**で起動し、**WSL2 / コンテナ上で実行**することを推奨します。

### 編集方式（エージェント）

既定の `agent` モードでは、LLMが **「コードを書く→実行→結果(print/エラー)を観測→次の手」**
を繰り返すReActループで編集します。途中のエラーを見て自己修正でき、複数ステップの作業や
シート間の操作にも対応します。

サンドボックスは安全のため `os` / `sys` / ファイル / ネットワークへのアクセスを遮断しますが、
**Excel編集に必要な範囲は開放**しており、`datetime` / `re` / `openpyxl.styles` 等の import が
可能です。これにより**太字・色・罫線・日付**などの書式設定も行えます。

従来の単発方式に戻したい場合は `XLSX_AGENT_MODE=oneshot` を指定してください。

### モデル名の設定（一箇所で切替）

モデル名などの設定は、リポジトリ直下の **`config.env` を編集すれば一括で切り替わります**
（`start.sh` / `start.bat` / アプリ本体すべてがこの値を参照）。サーバーで `ollama list` に
表示される名前に合わせてください。

```ini
# config.env
OLLAMA_MODEL=gemma4-e4b:latest
```

優先順位は **実環境変数 > config.env > コード内の既定値**。さらに本アプリは Ollama の
`/api/tags` を参照し、**ベース名が一致する実際のタグ（`...:latest` など）へ自動で寄せる**
ため、多少の表記揺れがあっても動作します。見つからない場合は `LLM_MODEL_NOT_FOUND`
エラーで pull コマンドを案内します。

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
