
import pytest
from datetime import date
from openpyxl import load_workbook

import tools.create_excel as create_excel_module
import tools.update_excel as update_module

from templates.reimbursement import (
    EMPLOYEE_SHEET,
    EMPLOYEE_FIELDS,
    DECLARATION_CELL,
    EXPENSE_SHEET,
    EXPENSE_FIRST_ROW,
    EXPENSE_COLUMNS,
    SUMMARY_SHEET,
)

def test_update_expense_record(reimbursement_file_with_expense):
    result = update_module.update_excel(
        filename="test.xlsx",
        expense_id="EXP001",
        updates={
            "description": "Airport taxi",
            "amount": 750,
        },
    )

    assert result["status"] == "success"
    assert result["updated_fields"] == ["description", "amount"]

    workbook = load_workbook(reimbursement_file_with_expense)
    worksheet = workbook[EXPENSE_SHEET]

    assert worksheet.cell(
        EXPENSE_FIRST_ROW,
        EXPENSE_COLUMNS["description"],
    ).value == "Airport taxi"

    assert worksheet.cell(
        EXPENSE_FIRST_ROW,
        EXPENSE_COLUMNS["amount"],
    ).value == 750

    workbook.close()


def test_update_expense_date(reimbursement_file_with_expense):
    result = update_module.update_excel(
        filename="test.xlsx",
        expense_id="EXP001",
        updates={"date": "2026-11-21"},
    )

    assert result["status"] == "success"

    workbook = load_workbook(reimbursement_file_with_expense)
    worksheet = workbook[EXPENSE_SHEET]

    assert worksheet.cell(
        EXPENSE_FIRST_ROW,
        EXPENSE_COLUMNS["date"],
    ).value.date() == date(2026, 11, 21)

    workbook.close()


def test_update_expense_category(reimbursement_file_with_expense):
    result = update_module.update_excel(
        filename="test.xlsx",
        expense_id="EXP001",
        updates={"category": "Food"},
    )

    assert result["status"] == "success"

    workbook = load_workbook(reimbursement_file_with_expense)
    worksheet = workbook[EXPENSE_SHEET]

    assert worksheet.cell(
        EXPENSE_FIRST_ROW,
        EXPENSE_COLUMNS["category"],
    ).value == "Food"

    workbook.close()


def test_clear_receipt_reference(reimbursement_file_with_expense):
    result = update_module.update_excel(
        filename="test.xlsx",
        expense_id="EXP001",
        updates={"receipt_reference": ""},
    )

    assert result["status"] == "success"

    workbook = load_workbook(reimbursement_file_with_expense)
    worksheet = workbook[EXPENSE_SHEET]

    assert worksheet.cell(
        EXPENSE_FIRST_ROW,
        EXPENSE_COLUMNS["receipt_reference"],
    ).value is None

    workbook.close()


def test_reject_missing_expense_id(reimbursement_file_with_expense):
    result = update_module.update_excel(
        filename="test.xlsx",
        updates={"amount": 700},
    )

    assert result["status"] == "error"
    assert "Expense ID" in result["message"]


def test_reject_missing_updates(reimbursement_file_with_expense):
    result = update_module.update_excel(
        filename="test.xlsx",
        expense_id="EXP001",
    )

    assert result["status"] == "error"


def test_reject_unsupported_update_field(reimbursement_file_with_expense):
    result = update_module.update_excel(
        filename="test.xlsx",
        expense_id="EXP001",
        updates={"employee_name": "Test Employee"},
    )

    assert result["status"] == "error"
    assert "Unsupported update fields" in result["message"]


def test_reject_invalid_expense_category(reimbursement_file_with_expense):
    result = update_module.update_excel(
        filename="test.xlsx",
        expense_id="EXP001",
        updates={"category": "Shopping"},
    )

    assert result["status"] == "error"
    assert "Invalid category" in result["message"]


def test_reject_negative_expense_amount(reimbursement_file_with_expense):
    result = update_module.update_excel(
        filename="test.xlsx",
        expense_id="EXP001",
        updates={"amount": -100},
    )

    assert result["status"] == "error"


def test_reject_nonexistent_expense_id(reimbursement_file_with_expense):
    result = update_module.update_excel(
        filename="test.xlsx",
        expense_id="EXP999",
        updates={"amount": 700},
    )

    assert result["status"] == "error"
    assert "not found" in result["message"]


def test_update_employee_information(reimbursement_file_with_expense):
    result = update_module.update_excel(
        filename="test.xlsx",
        employee_updates={
            "employee_name": "Test Employee",
            "department": "Engineering",
        },
    )

    assert result["status"] == "success"
    assert "employee_name" in result["updated_fields"]

    workbook = load_workbook(reimbursement_file_with_expense)
    worksheet = workbook[EMPLOYEE_SHEET]

    assert worksheet[
        EMPLOYEE_FIELDS["employee_name"]["cell"]
    ].value == "Test Employee"

    assert worksheet[
        EMPLOYEE_FIELDS["department"]["cell"]
    ].value == "Engineering"

    workbook.close()


def test_update_declaration_only(reimbursement_file_with_expense):
    result = update_module.update_excel(
        filename="test.xlsx",
        declaration="Yes",
    )

    assert result["status"] == "success"
    assert result["declaration"] == "Yes"

    workbook = load_workbook(reimbursement_file_with_expense)

    assert workbook[SUMMARY_SHEET][DECLARATION_CELL].value == "Yes"

    workbook.close()


def test_update_employee_information_and_declaration(reimbursement_file_with_expense):
    result = update_module.update_excel(
        filename="test.xlsx",
        employee_updates={
            "employee_name": "Test Employee",
        },
        declaration="Yes",
    )

    assert result["status"] == "success"
    assert result["declaration"] == "Yes"

    workbook = load_workbook(reimbursement_file_with_expense)

    assert workbook[EMPLOYEE_SHEET][
        EMPLOYEE_FIELDS["employee_name"]["cell"]
    ].value == "Test Employee"

    assert workbook[SUMMARY_SHEET][DECLARATION_CELL].value == "Yes"

    workbook.close()


def test_reject_invalid_employee_field(reimbursement_file_with_expense):
    result = update_module.update_excel(
        filename="test.xlsx",
        employee_updates={
            "invalid_field": "Some value",
        },
    )

    assert result["status"] == "error"
    assert "Unsupported employee fields" in result["message"]


@pytest.mark.parametrize(
    "travel_days",
    [0, -1, 1.5, True, "invalid"],
)
def test_reject_invalid_travel_days(reimbursement_file_with_expense, travel_days):
    result = update_module.update_excel(
        filename="test.xlsx",
        employee_updates={
            "travel_days": travel_days,
        },
    )

    assert result["status"] == "error"
    assert "Travel days" in result["message"]


def test_reject_end_date_before_start_date(reimbursement_file_with_expense):
    result = update_module.update_excel(
        filename="test.xlsx",
        employee_updates={
            "travel_start_date": "2026-11-25",
            "travel_end_date": "2026-11-20",
        },
    )

    assert result["status"] == "error"
    assert "Travel end date" in result["message"]


def test_reject_invalid_declaration(reimbursement_file_with_expense):
    result = update_module.update_excel(
        filename="test.xlsx",
        declaration="Maybe",
    )

    assert result["status"] == "error"


def test_reject_missing_workbook(tmp_path, monkeypatch):
    monkeypatch.setattr(
        update_module,
        "OUTPUT_DIR",
        tmp_path,
    )

    result = update_module.update_excel(
        filename="missing.xlsx",
        expense_id="EXP001",
        updates={"amount": 700},
    )

    assert result["status"] == "error"
    assert "not found" in result["message"].lower()


def test_reject_invalid_filename(tmp_path, monkeypatch):
    monkeypatch.setattr(
        update_module,
        "OUTPUT_DIR",
        tmp_path,
    )

    result = update_module.update_excel(
        filename="../outside.xlsx",
        expense_id="EXP001",
        updates={"amount": 700},
    )

    assert result["status"] == "error"
