import asyncio
import io
import time
from pathlib import Path

from fastapi import UploadFile
from openpyxl import Workbook, load_workbook

from app.main import JobService, create_app


class StubLLM:
    def generate_code(self, summary, instruction, feedback=""):
        return "ws['A1'] = 'edited'"


def _workbook_bytes() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "before"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _upload(filename: str) -> UploadFile:
    return UploadFile(file=io.BytesIO(_workbook_bytes()), filename=filename)


def _download_endpoint(app):
    for route in app.routes:
        if getattr(route, "path", None) == "/download/{token}":
            return route.endpoint
    raise AssertionError("download endpoint not found")


def _run_lifecycle(tmp_path: Path, filename: str) -> tuple[JobService, str, bytes, str]:
    service = JobService(tmp_path / f"jobs-{Path(filename).suffix[1:]}", llm=StubLLM())
    job_id = service.create_job(_upload(filename), "A1をeditedに変更", None)

    status = "queued"
    job = service.get_job(job_id)
    for _ in range(40):
        job = service.get_job(job_id)
        status = job.status
        if status == "preview_ready":
            break
        time.sleep(0.1)
    assert status == "preview_ready"
    assert job.preview is not None
    assert job.preview["changed_cell_count"] >= 1

    token = service.approve_job(job_id)
    app = create_app(service)
    response = asyncio.run(_download_endpoint(app)(token))
    assert response.status_code == 200
    with open(response.path, "rb") as result_file:
        content = result_file.read()
    out = load_workbook(io.BytesIO(content), data_only=False)
    assert out.active["A1"].value == "edited"
    out.close()

    try:
        service.pop_download_path(token)
        assert False, "Expected one-time download token to be consumed"
    except KeyError:
        pass
    return service, job_id, content, response.headers["content-disposition"]


def test_job_lifecycle(tmp_path: Path) -> None:
    service, job_id, _content, content_disposition = _run_lifecycle(
        tmp_path, "sample.xlsx"
    )

    job = service.get_job(job_id)
    assert job.source_path.suffix == ".xlsx"
    assert job.result_path is not None
    assert job.result_path.name == "result.xlsx"
    assert 'filename="edited.xlsx"' in content_disposition


def test_xlsm_job_lifecycle_preserves_result_and_download_extensions(
    tmp_path: Path,
) -> None:
    service, job_id, _content, content_disposition = _run_lifecycle(
        tmp_path, "sample.xlsm"
    )

    job = service.get_job(job_id)
    assert job.source_path.suffix == ".xlsm"
    assert job.result_path is not None
    assert job.result_path.name == "result.xlsm"
    assert 'filename="edited.xlsm"' in content_disposition


def test_load_workbook_calls_keep_vba_true_policy_is_present() -> None:
    source = Path("app/main.py").read_text()
    assert source.count("keep_vba=True, data_only=False") >= 5
    assert "uploaded_wb = load_workbook(" in source
    assert "load_workbook(job.source_path, keep_vba=True, data_only=False)" in source
    assert "load_workbook(path, keep_vba=True, data_only=False)" in source
