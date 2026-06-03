import pytest

from app.main import CodeChecker, JobError


@pytest.mark.parametrize(
    "code",
    [
        "import os\nws['A1'] = 1",
        "def f():\n    pass",
        "class X:\n    pass",
        "lambda x: x",
        "with something:\n    pass",
        "globals()",
        "getattr(ws, 'A1')",
    ],
)
def test_checker_rejects_forbidden_code(code: str) -> None:
    checker = CodeChecker()
    with pytest.raises(JobError) as err:
        checker.validate(code, set())
    assert err.value.error_code == "CODE_CHECK_FAILED"


def test_checker_rejects_dunder() -> None:
    checker = CodeChecker()
    with pytest.raises(JobError) as err:
        checker.validate("x = ws.__class__", set())
    assert err.value.error_code == "CODE_CHECK_FAILED"


def test_checker_rejects_non_anchor_write() -> None:
    checker = CodeChecker()
    with pytest.raises(JobError) as err:
        checker.validate("ws['B1'] = 'x'", {"B1"})
    assert err.value.error_code == "MERGED_CELL_CONFLICT"
