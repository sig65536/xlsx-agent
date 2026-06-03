import io
import time
from pathlib import Path

from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook

from app.main import JobService, create_app


class StubLLM:
    def generate_code(self, summary, instruction, feedback=""):
        return "ws['A1'] = 'edited'"


def _xlsx_bytes() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "before"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_job_lifecycle(tmp_path: Path) -> None:
    service = JobService(tmp_path / "jobs", llm=StubLLM())
    client = TestClient(create_app(service))

    resp = client.post(
        "/jobs",
        files={"file": ("sample.xlsx", _xlsx_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        data={"instruction": "A1をeditedに変更"},
    )
    assert resp.status_code == 200
    job_id = resp.json()["job_id"]

    status = "queued"
    for _ in range(40):
        job = client.get(f"/jobs/{job_id}").json()
        status = job["status"]
        if status == "preview_ready":
            break
        time.sleep(0.1)
    assert status == "preview_ready"
    assert job["preview"]["changed_cell_count"] >= 1

    approve = client.post(f"/jobs/{job_id}/approve")
    assert approve.status_code == 200
    download_url = approve.json()["download_url"]

    download = client.get(download_url)
    assert download.status_code == 200
    out = load_workbook(io.BytesIO(download.content), data_only=False)
    assert out.active["A1"].value == "edited"

    second = client.get(download_url)
    assert second.status_code == 404
