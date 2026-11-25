
import pytest
from openpyxl import load_workbook

import tools.create_excel as create_excel_module
import tools.write_excel as write_excel_module

from templates.reimbursement import (
    EMPLOYEE_SHEET,
    EMPLOYEE_FIELDS,
    DECLARATION_CELL,
    EXPENSE_SHEET,
    EXPENSE_FIRST_ROW,
    EXPENSE_COLUMNS,
    SUMMARY_SHEET,
)


@pytest.fixture
def reimbursement_file(tmp_path, monkeypatch):
    monkeypatch.setattr(
        create_excel_module,
        "OUTPUT_DIR",
        tmp_path,
    )

    monkeypatch.setattr(
        write_excel_module,
        "OUTPUT_DIR",
        tmp_path,
    )

    create_excel_module.create_excel("test.xlsx")

    return tmp_path / "test.xlsx"


def test_write_expense_record(reimbursement_file):
    result = write_excel_module.write_excel(
        filename="test.xlsx",
        expense_id="EXP001",
        expense_date="2026-11-25",
        category="Transport",
        description="Taxi to airport",
        amount=500,
        receipt_reference="REC001",
    )

    assert result["status"] == "success"
    assert result["expense_id"] == "EXP001"

    workbook = load_workbook(reimbursement_file)
    worksheet = workbook[EXPENSE_SHEET]

    row = EXPENSE_FIRST_ROW

    assert worksheet.cell(
        row=row,
        column=EXPENSE_COLUMNS["expense_id"],
    ).value == "EXP001"

    assert worksheet.cell(
        row=row,
        column=EXPENSE_COLUMNS["category"],
    ).value == "Transport"

    assert worksheet.cell(
        row=row,
        column=EXPENSE_COLUMNS["description"],
    ).value == "Taxi to airport"

    assert worksheet.cell(
        row=row,
        column=EXPENSE_COLUMNS["amount"],
    ).value == 500

    workbook.close()


def test_write_employee_information(reimbursement_file):
    employee_info = {
        "employee_name": "Test User",
        "employee_id": "EMP001",
        "department": "Engineering",
    }

    result = write_excel_module.write_excel(
        filename="test.xlsx",
        employee_info=employee_info,
    )

    assert result["status"] == "success"
    assert set(result["updated_fields"]) == set(employee_info.keys())

    workbook = load_workbook(reimbursement_file)
    worksheet = workbook[EMPLOYEE_SHEET]

    for field, value in employee_info.items():
        cell = EMPLOYEE_FIELDS[field]["cell"]
        assert worksheet[cell].value == value

    workbook.close()


def test_write_employee_information_and_declaration(
    reimbursement_file,
):
    result = write_excel_module.write_excel(
        filename="test.xlsx",
        employee_info={
            "employee_name": "Test User",
        },
        declaration="Yes",
    )

    assert result["status"] == "success"
    assert result["declaration"] == "Yes"

    workbook = load_workbook(reimbursement_file)

    assert (
        workbook[SUMMARY_SHEET][DECLARATION_CELL].value
        == "Yes"
    )

    workbook.close()


def test_reject_invalid_expense_category(reimbursement_file):
    result = write_excel_module.write_excel(
        filename="test.xlsx",
        expense_id="EXP002",
        expense_date="2026-11-25",
        category="Shopping",
        description="Shopping expense",
        amount=500,
    )

    assert result["status"] == "error"
    assert "Invalid category" in result["message"]


def test_reject_negative_expense_amount(reimbursement_file):
    result = write_excel_module.write_excel(
        filename="test.xlsx",
        expense_id="EXP003",
        expense_date="2026-11-25",
        category="Transport",
        description="Taxi",
        amount=-500,
    )

    assert result["status"] == "error"


def test_reject_duplicate_expense_id(reimbursement_file):
    first_result = write_excel_module.write_excel(
        filename="test.xlsx",
        expense_id="EXP004",
        expense_date="2026-11-25",
        category="Food",
        description="Lunch",
        amount=200,
    )

    assert first_result["status"] == "success"

    second_result = write_excel_module.write_excel(
        filename="test.xlsx",
        expense_id="EXP004",
        expense_date="2026-11-25",
        category="Food",
        description="Dinner",
        amount=300,
    )

    assert second_result["status"] == "error"
    assert "already exists" in second_result["message"]


def test_reject_formula_injection_in_description(
    reimbursement_file,
):
    result = write_excel_module.write_excel(
        filename="test.xlsx",
        expense_id="EXP005",
        expense_date="2026-11-25",
        category="Transport",
        description="=SUM(A1:A5)",
        amount=100,
    )

    assert result["status"] == "error"


def test_reject_invalid_employee_field(reimbursement_file):
    result = write_excel_module.write_excel(
        filename="test.xlsx",
        employee_info={
            "unsupported_field": "value",
        },
    )

    assert result["status"] == "error"
    assert "Unsupported employee fields" in result["message"]


def test_reject_fractional_travel_days(reimbursement_file):
    result = write_excel_module.write_excel(
        filename="test.xlsx",
        employee_info={
            "travel_days": 2.5,
        },
    )

    assert result["status"] == "error"
    assert "positive whole number" in result["message"]


def test_reject_invalid_declaration(reimbursement_file):
    result = write_excel_module.write_excel(
        filename="test.xlsx",
        declaration="Maybe",
    )

    assert result["status"] == "error"


def test_write_to_missing_workbook(tmp_path, monkeypatch):
    monkeypatch.setattr(
        write_excel_module,
        "OUTPUT_DIR",
        tmp_path,
    )

    result = write_excel_module.write_excel(
        filename="missing.xlsx",
        expense_id="EXP006",
        expense_date="2026-11-25",
        category="Food",
        description="Lunch",
        amount=100,
    )

    assert result["status"] == "error"
    assert "not found" in result["message"]
