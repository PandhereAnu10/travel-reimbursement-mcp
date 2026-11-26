
import pytest

from tools import create_excel
from tools import read_excel
from tools import list_excel_files
from tools import write_excel
from tools import update_excel
from tools import calculate_expense
from tools import submit_to_hr

from datetime import date
from openpyxl import load_workbook
from templates.reimbursement import (
    EMPLOYEE_FIELDS,
    REQUIRED_EMPLOYEE_FIELDS,
    EMPLOYEE_SHEET,
    EXPENSE_SHEET,
    EXPENSE_FIRST_ROW,
    EXPENSE_COLUMNS,
    SUMMARY_SHEET,
    DECLARATION_CELL,
)

@pytest.fixture
def output_dir(tmp_path, monkeypatch):
    # using temp o/p directory for tests

    modules = [
        create_excel,
        read_excel,
        list_excel_files,
        write_excel,
        update_excel,
        calculate_expense,
        submit_to_hr,
    ]

    for module in modules:
        monkeypatch.setattr(module, "OUTPUT_DIR", tmp_path)

    return tmp_path

@pytest.fixture
def reimbursement_file(output_dir):

    from tools.create_excel import create_excel

    create_excel("test.xlsx")
    return output_dir / "test.xlsx"


@pytest.fixture
def valid_claim(output_dir):

    from tools.create_excel import create_excel

    filename = "test_claim.xlsx"
    create_excel(filename)

    file_path = output_dir / filename
    workbook = load_workbook(file_path)

    employee_sheet = workbook[EMPLOYEE_SHEET]

    for field in REQUIRED_EMPLOYEE_FIELDS:
        cell = EMPLOYEE_FIELDS[field]["cell"]

        if field == "travel_start_date":
            employee_sheet[cell] = date(2026, 11, 20)
        elif field == "travel_end_date":
            employee_sheet[cell] = date(2026, 11, 22)
        elif field == "email":
            employee_sheet[cell] = "test@example.com"
        else:
            employee_sheet[cell] = f"Test {field}"

    expense_sheet = workbook[EXPENSE_SHEET]

    expense_values = {
        "expense_id": "EXP-001",
        "date": date(2026, 11, 21),
        "category": "Transport",
        "description": "Taxi fare",
        "amount": 500,
        "receipt_reference": "receipt-001",
    }

    for field, value in expense_values.items():
        expense_sheet.cell(
            EXPENSE_FIRST_ROW,
            EXPENSE_COLUMNS[field],
        ).value = value

    workbook[SUMMARY_SHEET][DECLARATION_CELL] = "Yes"

    workbook.save(file_path)
    workbook.close()
    return file_path

@pytest.fixture
def reimbursement_file_with_expense(reimbursement_file):
    file_path = reimbursement_file

    workbook = load_workbook(file_path)
    employee_sheet = workbook[EMPLOYEE_SHEET]
    employee_sheet[EMPLOYEE_FIELDS["travel_start_date"]["cell"]] = date(2026, 11, 20)
    employee_sheet[EMPLOYEE_FIELDS["travel_end_date"]["cell"]] = date(2026, 11, 22)

    expense_sheet = workbook[EXPENSE_SHEET]

    expense_values = {
        "expense_id": "EXP001",
        "date": date(2026, 11, 20),
        "category": "Transport",
        "description": "Taxi fare",
        "amount": 500,
        "receipt_reference": "REC001",
    }

    for field, value in expense_values.items():
        expense_sheet.cell(
            EXPENSE_FIRST_ROW,
            EXPENSE_COLUMNS[field],
        ).value = value

    workbook.save(file_path)
    workbook.close()

    return file_path
