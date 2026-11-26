
from datetime import date, datetime

import pytest
from openpyxl import Workbook, load_workbook

from tools.common_func import (
    save_workbook_safely,
    validate_expense_date,
    validate_expense_amount,
    validate_text,
    validate_expense_headers,
)


def test_save_workbook_safely(tmp_path):
    file_path = tmp_path / "safe_save.xlsx"

    workbook = Workbook()
    workbook.active["A1"] = "Test Data"

    save_workbook_safely(workbook, file_path)

    assert file_path.exists()

    saved_workbook = load_workbook(file_path)

    assert saved_workbook.active["A1"].value == "Test Data"

    saved_workbook.close()
    workbook.close()


def test_validate_expense_date_with_date():
    result = validate_expense_date(date(2026, 11, 25))

    assert result == date(2026, 11, 25)


def test_validate_expense_date_with_datetime():
    result = validate_expense_date(
        datetime(2026, 11, 25, 10, 30)
    )

    assert isinstance(result, date)
    assert result.year == 2026
    assert result.month == 11
    assert result.day == 25


def test_validate_expense_date_with_iso_string():
    result = validate_expense_date("2026-11-25")

    assert result == date(2026, 11, 25)


@pytest.mark.parametrize(
    "invalid_date",
    [
        "invalid-date",
        "1999-12-31",
        "2101-01-01",
        None,
        12345,
    ],
)
def test_reject_invalid_expense_date(invalid_date):
    with pytest.raises(ValueError):
        validate_expense_date(invalid_date)


def test_validate_expense_amount():
    result = validate_expense_amount(500.50)

    assert result == 500.50
    assert isinstance(result, float)


@pytest.mark.parametrize(
    "invalid_amount",
    [
        -100,
        None,
        True,
        "invalid",
        float("nan"),
        float("inf"),
        float("-inf"),
    ],
)
def test_reject_invalid_expense_amount(invalid_amount):
    with pytest.raises(ValueError):
        validate_expense_amount(invalid_amount)


def test_validate_text():
    result = validate_text(
        "  Travel expense  ",
        "Description",
    )

    assert result == "Travel expense"


@pytest.mark.parametrize(
    "invalid_text",
    [
        None,
        "",
        "   ",
        "=SUM(A1:A5)",
        "+123",
        "-123",
        "@command",
    ],
)
def test_reject_invalid_text(invalid_text):
    with pytest.raises(ValueError):
        validate_text(invalid_text, "Description")


def test_validate_text_rejects_non_string():
    with pytest.raises(ValueError):
        validate_text(123, "Description")


def test_validate_expense_headers():
    workbook = Workbook()
    worksheet = workbook.active

    expected_headers = [
        "Expense ID",
        "Date",
        "Category",
        "Description",
        "Amount (INR)",
        "Receipt Reference",
        "Date Check",
    ]

    for column, header in enumerate(expected_headers, start=1):
        worksheet.cell(row=4, column=column).value = header

    # Should not raise an error for valid headers.
    validate_expense_headers(worksheet)

    workbook.close()


def test_reject_invalid_expense_headers():
    workbook = Workbook()
    worksheet = workbook.active

    worksheet["A4"] = "Wrong Header"
    worksheet["B4"] = "Wrong Date"

    with pytest.raises(ValueError):
        validate_expense_headers(worksheet)

    workbook.close()


def test_save_workbook_safely_retries_transient_permission_error(tmp_path, monkeypatch):
    import tools.common_func as common_func

    file_path = tmp_path / "retry_save.xlsx"
    workbook = Workbook()
    workbook.active["A1"] = "Retry test"
    original_replace = common_func.os.replace
    attempts = {"count": 0}

    def flaky_replace(source, destination):
        attempts["count"] += 1
        if attempts["count"] == 1:
            raise PermissionError("temporary file lock")
        return original_replace(source, destination)

    monkeypatch.setattr(common_func.os, "replace", flaky_replace)
    monkeypatch.setattr(common_func.time, "sleep", lambda _seconds: None)

    save_workbook_safely(workbook, file_path)

    assert attempts["count"] == 2
    saved_workbook = load_workbook(file_path)
    assert saved_workbook.active["A1"].value == "Retry test"
    saved_workbook.close()
    workbook.close()


def test_save_workbook_safely_reports_persistent_windows_lock(tmp_path, monkeypatch):
    import tools.common_func as common_func

    file_path = tmp_path / "locked.xlsx"
    workbook = Workbook()
    monkeypatch.setattr(
        common_func.os, "replace",
        lambda *_args: (_ for _ in ()).throw(PermissionError("access denied")),
    )
    monkeypatch.setattr(common_func.time, "sleep", lambda _seconds: None)

    with pytest.raises(PermissionError, match="Close the Excel file"):
        save_workbook_safely(workbook, file_path)

    assert not file_path.exists()
    assert not list(tmp_path.glob(".reimbursement_*.xlsx"))
    workbook.close()
