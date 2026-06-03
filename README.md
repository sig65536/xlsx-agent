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
