from app.main import CodeChecker, JobError


def test_checker_rejects_import() -> None:
    checker = CodeChecker()
    try:
        checker.validate("import os\nws['A1'] = 1", set())
        assert False, "Expected JobError"
    except JobError as err:
        assert err.error_code == "CODE_CHECK_FAILED"


def test_checker_rejects_dunder() -> None:
    checker = CodeChecker()
    try:
        checker.validate("x = ws.__class__", set())
        assert False, "Expected JobError"
    except JobError as err:
        assert err.error_code == "CODE_CHECK_FAILED"


def test_checker_rejects_non_anchor_write() -> None:
    checker = CodeChecker()
    try:
        checker.validate("ws['B1'] = 'x'", {"B1"})
        assert False, "Expected JobError"
    except JobError as err:
        assert err.error_code == "MERGED_CELL_CONFLICT"
