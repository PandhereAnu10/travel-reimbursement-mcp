
from pathlib import Path

from openpyxl import Workbook

from templates.reimbursement import (
    EMPLOYEE_SHEET,
    EXPENSE_SHEET,
    CALCULATOR_SHEET,
    SUMMARY_SHEET,
    EMPLOYEE_TITLE_CELL,
    EXPENSE_TITLE_CELL,
    CALCULATOR_TITLE_CELL,
    SUMMARY_TITLE_CELL,
)


# Project's output directory
OUTPUT_DIR = Path(__file__).resolve().parents[1] / "data" / "generated"


def validate_filename(filename: str) -> str:

    if not filename or not filename.strip():
        raise ValueError("Filename cannot be empty.")

    filename = filename.strip()

    # Prevent directory traversal and paths outside OUTPUT_DIR.
    if "/" in filename or "\\" in filename:
        raise ValueError("Only a filename is allowed, not a directory path.")

    if filename in {".", ".."}:
        raise ValueError("Invalid filename.")

    # Add .xlsx if the user did not provide an extension.
    if not Path(filename).suffix:
        filename = f"{filename}.xlsx"

    if Path(filename).suffix.lower() != ".xlsx":
        raise ValueError("Only .xlsx files are supported.")

    return filename


def create_excel(filename: str) -> dict:
    # Step 1: Validate the filename.
    filename = validate_filename(filename)

    # Step 2: Create the output directory if it does not exist.
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # Step 3: Build the complete output path.
    file_path = OUTPUT_DIR / filename

    # Step 4: Prevent accidental overwriting.
    if file_path.exists():
        raise FileExistsError(
            f"A workbook named '{filename}' already exists."
        )

    # Step 5: Create a new workbook.
    workbook = Workbook()

    # Remove the default worksheet created by OpenPyXL.
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    # Step 6: Create the four worksheets.
    employee_ws = workbook.create_sheet(EMPLOYEE_SHEET)
    expense_ws = workbook.create_sheet(EXPENSE_SHEET)
    calculator_ws = workbook.create_sheet(CALCULATOR_SHEET)
    summary_ws = workbook.create_sheet(SUMMARY_SHEET)

    # Step 7: Add worksheet titles.
    employee_ws[EMPLOYEE_TITLE_CELL] = EMPLOYEE_SHEET
    expense_ws[EXPENSE_TITLE_CELL] = EXPENSE_SHEET
    calculator_ws[CALCULATOR_TITLE_CELL] = CALCULATOR_SHEET
    summary_ws[SUMMARY_TITLE_CELL] = SUMMARY_SHEET

    # Step 8: Save the workbook.
    workbook.save(file_path)

    # Step 9: Return information about the created workbook.
    return {
        "status": "success",
        "message": "Workbook created successfully.",
        "file_name": filename,
        "file_path": str(file_path),
        "worksheets": workbook.sheetnames,
    }
