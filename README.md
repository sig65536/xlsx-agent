# xlsx-agent
ローカルllmでエクセル操作したい

## ローカル起動

```bash
pip install -e .[test]
uvicorn app.main:app --reload
```

## API

- `POST /jobs` : Excelファイル + 指示文を受け付け、`job_id`を返す
- `GET /jobs/{job_id}` : ジョブ状態とプレビューを返す
- `POST /jobs/{job_id}/approve` : 承認してダウンロードURLを返す
- `GET /download/{token}` : ワンタイムトークンで結果をダウンロード

対応形式は `.xlsx` / `.xlsm`（`keep_vba=True` でVBAは保持のみ）です。マクロは実行せず、内容も変更しません。

## セキュリティ上の注意

本MVPのLLM生成コード実行は、別プロセス + ASTチェック + restricted builtins による制限付き実行です。
Docker / VM / 別OSユーザーによる完全なサンドボックスではありません。

外部公開環境では使用しないでください。
利用範囲はローカルPCまたは信頼できる社内LANに限定してください。
機密ファイルを扱う場合は、必ず原本バックアップを保持してください。
