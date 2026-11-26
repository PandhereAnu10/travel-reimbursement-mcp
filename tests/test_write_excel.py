
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


def set_test_travel_period(workbook_path):
    workbook = load_workbook(workbook_path)
    employee = workbook[EMPLOYEE_SHEET]
    employee[EMPLOYEE_FIELDS["travel_start_date"]["cell"]] = "2026-11-20"
    employee[EMPLOYEE_FIELDS["travel_end_date"]["cell"]] = "2026-11-30"
    workbook.save(workbook_path)
    workbook.close()


def test_write_expense_record(reimbursement_file):
    set_test_travel_period(reimbursement_file)
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
        "designation": "Engineer",
        "email": "test@example.com",
        "travel_purpose": "Business Meeting",
        "source_location": "Mumbai",
        "destination_location": "Bengaluru",
        "travel_start_date": "2026-11-20",
        "travel_end_date": "2026-11-22",
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
        actual = worksheet[cell].value
        if field in {"travel_start_date", "travel_end_date"}:
            assert actual.date().isoformat() == value
        else:
            assert actual == value

    travel_days_cell = EMPLOYEE_FIELDS["travel_days"]["cell"]
    assert worksheet[travel_days_cell].data_type == "f"

    workbook.close()


def test_write_employee_information_and_declaration(
    reimbursement_file,
):
    result = write_excel_module.write_excel(
        filename="test.xlsx",
        employee_info={
            "employee_name": "Test User",
            "employee_id": "EMP001",
            "department": "Engineering",
            "designation": "Engineer",
            "email": "test@example.com",
            "travel_purpose": "Business Meeting",
            "source_location": "Mumbai",
            "destination_location": "Bengaluru",
            "travel_start_date": "2026-11-20",
            "travel_end_date": "2026-11-22",
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


def test_reject_incomplete_employee_information_without_saving(reimbursement_file):
    result = write_excel_module.write_excel(
        filename="test.xlsx",
        employee_info={"employee_name": "Test User", "employee_id": "EMP001"},
    )
    assert result["status"] == "error"
    assert "Missing required fields" in result["message"]

    workbook = load_workbook(reimbursement_file)
    worksheet = workbook[EMPLOYEE_SHEET]
    assert worksheet[EMPLOYEE_FIELDS["employee_name"]["cell"]].value is None
    workbook.close()

def test_allow_partial_employee_information_only_as_explicit_draft(reimbursement_file):
    result = write_excel_module.write_excel(
        filename="test.xlsx",
        employee_info={"employee_name": "Test User", "employee_id": "EMP001"},
        save_as_draft=True,
    )
    assert result["status"] == "success"
    assert result["record_status"] == "Draft - incomplete"
    assert set(result["missing_fields"]) == set(__import__("templates.reimbursement", fromlist=["REQUIRED_EMPLOYEE_FIELDS"]).REQUIRED_EMPLOYEE_FIELDS) - {"employee_name", "employee_id"}


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
    set_test_travel_period(reimbursement_file)
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
    assert "calculated automatically" in result["message"]


def test_reject_invalid_declaration(reimbursement_file):
    result = write_excel_module.write_excel(
        filename="test.xlsx",
        declaration="Maybe",
    )

    assert result["status"] == "error"


def test_write_to_missing_workbook(output_dir):

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


def test_reject_end_date_before_start_date(reimbursement_file):
    result = write_excel_module.write_excel(
        filename="test.xlsx",
        employee_info={
            "employee_name": "Test User",
            "employee_id": "EMP001",
            "department": "Engineering",
            "designation": "Engineer",
            "email": "test@example.com",
            "travel_purpose": "Business Meeting",
            "source_location": "Mumbai",
            "destination_location": "Bengaluru",
            "travel_start_date": "2026-11-25",
            "travel_end_date": "2026-11-20",
        },
    )

    assert result["status"] == "error"
    assert "Travel end date" in result["message"]


def test_reject_invalid_employee_email(reimbursement_file):
    info = {
        "employee_name": "Test User", "employee_id": "EMP001",
        "department": "Engineering", "designation": "Engineer",
        "email": "not-an-email", "travel_purpose": "Business Meeting",
        "source_location": "Mumbai", "destination_location": "Bengaluru",
        "travel_start_date": "2026-11-20", "travel_end_date": "2026-11-22",
    }
    result = write_excel_module.write_excel(filename="test.xlsx", employee_info=info)
    assert result["status"] == "error"
    assert "valid email" in result["message"]


def test_expense_date_must_be_within_existing_travel_period(reimbursement_file):
    info = {
        "employee_name": "Test User", "employee_id": "EMP001",
        "department": "Engineering", "designation": "Engineer",
        "email": "test@example.com", "travel_purpose": "Business Meeting",
        "source_location": "Mumbai", "destination_location": "Bengaluru",
        "travel_start_date": "2026-11-20", "travel_end_date": "2026-11-22",
    }
    assert write_excel_module.write_excel(filename="test.xlsx", employee_info=info)["status"] == "success"
    result = write_excel_module.write_excel(
        filename="test.xlsx", expense_id="EXP-RANGE", expense_date="2026-11-25",
        category="Transport", description="Taxi", amount=500,
    )
    assert result["status"] == "error"
    assert "between the travel start date" in result["message"]

    workbook = load_workbook(reimbursement_file)
    worksheet = workbook[EXPENSE_SHEET]
    assert all(
        worksheet.cell(row, EXPENSE_COLUMNS["expense_id"]).value is None
        for row in range(EXPENSE_FIRST_ROW, EXPENSE_FIRST_ROW + 3)
    )
    workbook.close()


def test_partial_employee_write_merges_existing_values(reimbursement_file):
    info = {
        "employee_name": "Test User", "employee_id": "EMP001",
        "department": "Engineering", "designation": "Engineer",
        "email": "test@example.com", "travel_purpose": "Business Meeting",
        "source_location": "Mumbai", "destination_location": "Bengaluru",
        "travel_start_date": "2026-11-20", "travel_end_date": "2026-11-22",
    }
    assert write_excel_module.write_excel(filename="test.xlsx", employee_info=info)["status"] == "success"
    result = write_excel_module.write_excel(
        filename="test.xlsx", employee_info={"department": "Research"}
    )
    assert result["status"] == "success"
    assert result["record_status"] == "Complete"


def test_new_expense_is_rejected_when_travel_period_is_missing(reimbursement_file):
    result = write_excel_module.write_excel(
        filename="test.xlsx",
        expense_id="EXP-NO-PERIOD",
        expense_date="2026-11-25",
        category="Food",
        description="Lunch",
        amount=250,
    )
    assert result["status"] == "error"
    assert "Travel start and end dates must be recorded" in result["message"]

    workbook = load_workbook(reimbursement_file)
    worksheet = workbook[EXPENSE_SHEET]
    assert worksheet.cell(
        EXPENSE_FIRST_ROW, EXPENSE_COLUMNS["expense_id"]
    ).value is None
    workbook.close()
