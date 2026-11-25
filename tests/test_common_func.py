
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
