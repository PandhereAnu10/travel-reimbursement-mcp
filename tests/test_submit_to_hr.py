
from datetime import date

import pytest
from openpyxl import load_workbook

from tools import create_excel as create_excel_module
from tools import submit_to_hr as submit_module

from templates.reimbursement import (
    EMPLOYEE_FIELDS,
    REQUIRED_EMPLOYEE_FIELDS,
    EMPLOYEE_SHEET,
    EXPENSE_SHEET,
    EXPENSE_FIRST_ROW,
    EXPENSE_COLUMNS,
    SUMMARY_SHEET,
    DECLARATION_CELL,
    SUBMISSION_STATUS_CELL,
    SUBMISSION_DATE_CELL,
    CLAIM_REFERENCE_ID_CELL,
    SUBMITTED_STATUS,
)


def test_submit_valid_claim(valid_claim):
    result = submit_module.submit_to_hr(valid_claim.name)

    assert result["status"] == "success"
    assert result["filename"] == valid_claim.name
    assert result["submission_status"] == SUBMITTED_STATUS
    assert result["total_expenses"] == 1
    assert result["claim_reference"].startswith(
        f"TRV-{date.today().year}-"
    )

    workbook = load_workbook(valid_claim)
    summary_sheet = workbook[SUMMARY_SHEET]

    assert summary_sheet[SUBMISSION_STATUS_CELL].value == SUBMITTED_STATUS
    saved_submission_date = summary_sheet[SUBMISSION_DATE_CELL].value

    if hasattr(saved_submission_date, "date"):
        saved_submission_date = saved_submission_date.date()

    assert saved_submission_date == date.today()
    assert (
        summary_sheet[CLAIM_REFERENCE_ID_CELL].value
        == result["claim_reference"]
    )

    workbook.close()


def test_prevent_duplicate_submission(valid_claim):
    first_result = submit_module.submit_to_hr(valid_claim.name)
    second_result = submit_module.submit_to_hr(valid_claim.name)

    assert first_result["status"] == "success"
    assert second_result["status"] == "already_submitted"


def test_reject_missing_employee_information(valid_claim):
    workbook = load_workbook(valid_claim)
    employee_sheet = workbook[EMPLOYEE_SHEET]

    employee_sheet[
        EMPLOYEE_FIELDS["employee_name"]["cell"]
    ] = None

    workbook.save(valid_claim)
    workbook.close()

    result = submit_module.submit_to_hr(valid_claim.name)

    assert result["status"] == "incomplete"
    assert EMPLOYEE_FIELDS["employee_name"]["label"] in result["missing_fields"]


def test_reject_claim_without_expenses(valid_claim):
    workbook = load_workbook(valid_claim)
    expense_sheet = workbook[EXPENSE_SHEET]

    for field in (
    "expense_id",
    "date",
    "category",
    "description",
    "amount",
    "receipt_reference",
):
        expense_sheet.cell(
            EXPENSE_FIRST_ROW,
            EXPENSE_COLUMNS[field],
        ).value = None

    workbook.save(valid_claim)
    workbook.close()

    result = submit_module.submit_to_hr(valid_claim.name)

    assert result["status"] == "incomplete"
    assert "At least one expense record" in result["message"]


def test_reject_incomplete_expense(valid_claim):
    workbook = load_workbook(valid_claim)
    expense_sheet = workbook[EXPENSE_SHEET]

    expense_sheet.cell(
        EXPENSE_FIRST_ROW,
        EXPENSE_COLUMNS["description"],
    ).value = None

    workbook.save(valid_claim)
    workbook.close()

    result = submit_module.submit_to_hr(valid_claim.name)

    assert result["status"] == "incomplete"
    assert len(result["incomplete_expenses"]) == 1
    assert "Description" in result["incomplete_expenses"][0]["missing_fields"]


def test_reject_invalid_expense_category(valid_claim):
    workbook = load_workbook(valid_claim)
    expense_sheet = workbook[EXPENSE_SHEET]

    expense_sheet.cell(
        EXPENSE_FIRST_ROW,
        EXPENSE_COLUMNS["category"],
    ).value = "Invalid Category"

    workbook.save(valid_claim)
    workbook.close()

    result = submit_module.submit_to_hr(valid_claim.name)

    assert result["status"] == "error"
    assert "Invalid expense category" in result["message"]


def test_reject_expense_date_outside_travel_period(valid_claim):
    workbook = load_workbook(valid_claim)
    expense_sheet = workbook[EXPENSE_SHEET]

    expense_sheet.cell(
        EXPENSE_FIRST_ROW,
        EXPENSE_COLUMNS["date"],
    ).value = date(2026, 11, 25)

    workbook.save(valid_claim)
    workbook.close()

    result = submit_module.submit_to_hr(valid_claim.name)

    assert result["status"] == "error"
    assert "must fall between" in result["message"]


def test_reject_duplicate_expense_ids(valid_claim):
    workbook = load_workbook(valid_claim)
    expense_sheet = workbook[EXPENSE_SHEET]

    second_row = EXPENSE_FIRST_ROW + 1

    expense_values = {
        "expense_id": "exp-001",
        "date": date(2026, 11, 21),
        "category": "Food",
        "description": "Lunch",
        "amount": 200,
        "receipt_reference": None,
    }

    for field, value in expense_values.items():
        expense_sheet.cell(
            second_row,
            EXPENSE_COLUMNS[field],
        ).value = value

    workbook.save(valid_claim)
    workbook.close()

    result = submit_module.submit_to_hr(valid_claim.name)

    assert result["status"] == "error"
    assert "Duplicate expense ID" in result["message"]


def test_reject_claim_without_accepted_declaration(valid_claim):
    workbook = load_workbook(valid_claim)
    workbook[SUMMARY_SHEET][DECLARATION_CELL] = "No"

    workbook.save(valid_claim)
    workbook.close()

    result = submit_module.submit_to_hr(valid_claim.name)

    assert result["status"] == "incomplete"
    assert result["required_declaration"] == "Yes"


def test_reject_end_date_before_start_date(valid_claim):
    workbook = load_workbook(valid_claim)
    employee_sheet = workbook[EMPLOYEE_SHEET]

    employee_sheet[
        EMPLOYEE_FIELDS["travel_start_date"]["cell"]
    ] = date(2026, 11, 23)

    employee_sheet[
        EMPLOYEE_FIELDS["travel_end_date"]["cell"]
    ] = date(2026, 11, 20)

    workbook.save(valid_claim)
    workbook.close()

    result = submit_module.submit_to_hr(valid_claim.name)

    assert result["status"] == "error"
    assert "end date cannot be earlier" in result["message"]


@pytest.mark.parametrize(
    "filename",
    [
        "",
        "../test_claim.xlsx",
        "folder/test_claim.xlsx",
        "folder\\test_claim.xlsx",
        "test_claim.csv",
        None,
    ],
)
def test_reject_invalid_filename(filename):
    result = submit_module.submit_to_hr(filename)

    assert result["status"] == "error"
    assert "valid Excel filename" in result["message"]


def test_reject_missing_workbook(valid_claim):
    result = submit_module.submit_to_hr("missing_claim.xlsx")

    assert result["status"] == "error"
    assert "does not exist" in result["message"]


def test_reject_corrupted_workbook(valid_claim):
    valid_claim.write_bytes(b"not a valid Excel workbook")

    result = submit_module.submit_to_hr(valid_claim.name)

    assert result["status"] == "error"
    assert "invalid or corrupted" in result["message"]


def test_reject_workbook_with_missing_sheet(valid_claim):
    workbook = load_workbook(valid_claim)
    workbook.remove(workbook[EXPENSE_SHEET])

    workbook.save(valid_claim)
    workbook.close()

    result = submit_module.submit_to_hr(valid_claim.name)

    assert result["status"] == "error"
    assert "missing required worksheets" in result["message"]


def test_generate_next_claim_reference(valid_claim, tmp_path):
    previous_filename = "previous_claim.xlsx"
    create_excel_module.create_excel(previous_filename)

    previous_path = tmp_path / previous_filename
    workbook = load_workbook(previous_path)

    workbook[SUMMARY_SHEET][CLAIM_REFERENCE_ID_CELL] = (
        f"TRV-{date.today().year}-0007"
    )

    workbook.save(previous_path)
    workbook.close()

    result = submit_module.generate_claim_reference()

    assert result == f"TRV-{date.today().year}-0008"


def test_submit_multiple_expenses(valid_claim):
    workbook = load_workbook(valid_claim)
    expense_sheet = workbook[EXPENSE_SHEET]

    second_row = EXPENSE_FIRST_ROW + 1

    expense_values = {
        "expense_id": "EXP-002",
        "date": date(2026, 11, 22),
        "category": "Food",
        "description": "Lunch",
        "amount": 250,
        "receipt_reference": None,
    }

    for field, value in expense_values.items():
        expense_sheet.cell(
            second_row,
            EXPENSE_COLUMNS[field],
        ).value = value

    workbook.save(valid_claim)
    workbook.close()

    result = submit_module.submit_to_hr(valid_claim.name)

    assert result["status"] == "success"
    assert result["total_expenses"] == 2


def test_reject_negative_expense_amount(valid_claim):
    workbook = load_workbook(valid_claim)
    expense_sheet = workbook[EXPENSE_SHEET]

    expense_sheet.cell(
        EXPENSE_FIRST_ROW,
        EXPENSE_COLUMNS["amount"],
    ).value = -100

    workbook.save(valid_claim)
    workbook.close()

    result = submit_module.submit_to_hr(valid_claim.name)

    assert result["status"] == "error"
