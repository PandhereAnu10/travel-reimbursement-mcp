
import re
from datetime import date
from pathlib import Path
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from tools.create_excel import OUTPUT_DIR

from templates.reimbursement import (
    EMPLOYEE_SHEET,
    EMPLOYEE_FIELDS,
    REQUIRED_EMPLOYEE_FIELDS,
    EXPENSE_SHEET,
    EXPENSE_FIRST_ROW,
    EXPENSE_LAST_ROW,
    EXPENSE_COLUMNS,
    EXPENSE_CATEGORIES,
    SUMMARY_SHEET,
    SUBMISSION_STATUS_CELL,
    SUBMISSION_DATE_CELL,
    CLAIM_REFERENCE_ID_CELL,
    DECLARATION_CELL,
    SUBMITTED_STATUS,
    DATE_FORMAT,
)

from tools.common_func import (
    save_workbook_safely,
    validate_expense_date,
    validate_expense_amount,
    validate_text,
    validate_expense_headers,
)


def generate_claim_reference():
    # generate yearly claim reference no.

    current_year = date.today().year
    highest_number = 0

    for file_path in OUTPUT_DIR.glob("*.xlsx"):
        workbook = None

        try:
            workbook = load_workbook(
                file_path,
                read_only=True,
                data_only=True,
            )

            if SUMMARY_SHEET not in workbook.sheetnames:
                continue

            reference = workbook[SUMMARY_SHEET][
                CLAIM_REFERENCE_ID_CELL
            ].value

            if not reference:
                continue

            match = re.fullmatch(
                rf"TRV-{current_year}-(\d+)",
                str(reference).strip(),
            )

            if match:
                highest_number = max(
                    highest_number,
                    int(match.group(1)),
                )

        except (BadZipFile, InvalidFileException, OSError):
            continue

        finally:
            if workbook:
                workbook.close()

    return f"TRV-{current_year}-{highest_number + 1:04d}"


def submit_to_hr(filename: str) -> dict:
    # Validate and submit reimbursement claim to HR

    workbook = None

    try:
        if (
            not isinstance(filename, str)
            or not filename.strip()
            or Path(filename).name != filename
            or "/" in filename
            or "\\" in filename
            or Path(filename).suffix.lower() != ".xlsx"
        ):
            return {
                "status": "error",
                "message": "Provide a valid Excel filename ending in .xlsx.",
            }

        file_path = (OUTPUT_DIR / filename).resolve()

        if file_path.parent != OUTPUT_DIR.resolve():
            return {
                "status": "error",
                "message": "Access outside the generated workbook directory is not allowed.",
            }

        if not file_path.is_file():
            return {
                "status": "error",
                "message": f"Workbook '{filename}' does not exist.",
            }

        # Load the workbook.
        workbook = load_workbook(file_path)

        required_sheets = [
            EMPLOYEE_SHEET,
            EXPENSE_SHEET,
            SUMMARY_SHEET,
        ]

        missing_sheets = [
            name for name in required_sheets
            if name not in workbook.sheetnames
        ]

        if missing_sheets:
            return {
                "status": "error",
                "message": "Workbook is missing required worksheets.",
                "missing_sheets": missing_sheets,
            }

        employee_sheet = workbook[EMPLOYEE_SHEET]
        expense_sheet = workbook[EXPENSE_SHEET]
        summary_sheet = workbook[SUMMARY_SHEET]

        # Prevent duplicate submissions.
        existing_status = summary_sheet[
            SUBMISSION_STATUS_CELL
        ].value

        if existing_status == SUBMITTED_STATUS:
            return {
                "status": "already_submitted",
                "message": "This claim has already been submitted to HR.",
                "submission_status": existing_status,
                "submission_date": summary_sheet[
                    SUBMISSION_DATE_CELL
                ].value,
                "claim_reference": summary_sheet[
                    CLAIM_REFERENCE_ID_CELL
                ].value,
            }

        # Validate employee information.
        missing_fields = []

        for field in REQUIRED_EMPLOYEE_FIELDS:
            cell = EMPLOYEE_FIELDS[field]["cell"]
            label = EMPLOYEE_FIELDS[field]["label"]
            value = employee_sheet[cell].value

            if value is None or (
                isinstance(value, str) and not value.strip()
            ):
                missing_fields.append(label)

        if missing_fields:
            return {
                "status": "incomplete",
                "message": "Required employee and travel information is missing.",
                "missing_fields": missing_fields,
            }

        # Validate travel dates.
        start_date = validate_expense_date(
            employee_sheet[
                EMPLOYEE_FIELDS["travel_start_date"]["cell"]
            ].value
        )

        end_date = validate_expense_date(
            employee_sheet[
                EMPLOYEE_FIELDS["travel_end_date"]["cell"]
            ].value
        )

        if end_date < start_date:
            return {
                "status": "error",
                "message": "Travel end date cannot be earlier than travel start date.",
            }

        # Validate the expense worksheet structure.
        validate_expense_headers(expense_sheet)

        # Ensure at least one complete expense record exists.
        expense_count = 0
        incomplete_expenses = []
        seen_expense_ids = set()

        for row in range(EXPENSE_FIRST_ROW, EXPENSE_LAST_ROW + 1):
            expense_id = expense_sheet.cell(
                row, EXPENSE_COLUMNS["expense_id"]
            ).value

            expense_date = expense_sheet.cell(
                row, EXPENSE_COLUMNS["date"]
            ).value

            category = expense_sheet.cell(
                row, EXPENSE_COLUMNS["category"]
            ).value

            description = expense_sheet.cell(
                row, EXPENSE_COLUMNS["description"]
            ).value

            amount = expense_sheet.cell(
                row, EXPENSE_COLUMNS["amount"]
            ).value

            receipt = expense_sheet.cell(
                row, EXPENSE_COLUMNS["receipt_reference"]
            ).value

            values = [
                expense_id, expense_date, category,
                description, amount, receipt,
            ]

            if all(value is None or value == "" for value in values):
                continue

            required_values = {
                "Expense ID": expense_id,
                "Date": expense_date,
                "Category": category,
                "Description": description,
                "Amount": amount,
            }

            missing = [
                name for name, value in required_values.items()
                if value is None or (
                    isinstance(value, str) and not value.strip()
                )
            ]

            if missing:
                incomplete_expenses.append({
                    "row": row,
                    "missing_fields": missing,
                })
                continue

            expense_id = validate_text(expense_id, "Expense ID")
            expense_date = validate_expense_date(expense_date)
            category = validate_text(category, "Category")
            validate_text(description, "Description")
            validate_expense_amount(amount)

            if receipt not in (None, ""):
                validate_text(receipt, "Receipt reference")

            # Check for duplicate expense IDs.
            normalized_expense_id = expense_id.casefold()

            if normalized_expense_id in seen_expense_ids:
                return {
                    "status": "error",
                    "message": f"Duplicate expense ID '{expense_id}' found in row {row}.",
                }

            seen_expense_ids.add(normalized_expense_id)

            # Ensure expense date falls within the travel period.
            if not start_date <= expense_date <= end_date:
                return {
                    "status": "error",
                    "message": (
                        f"Expense date in row {row} must fall between "
                        "the travel start date and travel end date."
                    ),
                }

            if category not in EXPENSE_CATEGORIES:
                return {
                    "status": "error",
                    "message": (
                        f"Invalid expense category '{category}' "
                        f"in row {row}."
                    ),
                }

            expense_count += 1

        if incomplete_expenses:
            return {
                "status": "incomplete",
                "message": "Some expense records are incomplete.",
                "incomplete_expenses": incomplete_expenses,
            }

        if expense_count == 0:
            return {
                "status": "incomplete",
                "message": "At least one expense record is required before submission.",
            }

        # Validate employee declaration.
        declaration = summary_sheet[DECLARATION_CELL].value

        if not isinstance(declaration, str) or declaration.strip().lower() != "yes":
            return {
                "status": "incomplete",
                "message": (
                    "Employee declaration must be accepted before "
                    "the claim can be submitted."
                ),
                "required_declaration": "Yes",
            }

        # Generate submission details.
        submission_date = date.today()
        claim_reference = generate_claim_reference()

        summary_sheet[SUBMISSION_STATUS_CELL] = SUBMITTED_STATUS
        summary_sheet[SUBMISSION_DATE_CELL] = submission_date
        summary_sheet[SUBMISSION_DATE_CELL].number_format = DATE_FORMAT
        summary_sheet[CLAIM_REFERENCE_ID_CELL] = claim_reference

        # Save safely.
        save_workbook_safely(workbook, file_path)

        return {
            "status": "success",
            "message": "Travel reimbursement claim successfully submitted to HR.",
            "filename": filename,
            "submission_status": SUBMITTED_STATUS,
            "submission_date": submission_date.isoformat(),
            "claim_reference": claim_reference,
            "total_expenses": expense_count,
        }

    except (BadZipFile, InvalidFileException):
        return {
            "status": "error",
            "message": "The Excel workbook is invalid or corrupted.",
        }

    except (ValueError, TypeError) as error:
        return {
            "status": "error",
            "message": str(error),
        }

    except PermissionError:
        return {
            "status": "error",
            "message": "Workbook could not be saved. Check whether it is open in Excel.",
        }

    except OSError:
        return {
            "status": "error",
            "message": "A file operation failed while submitting the claim.",
        }

    except Exception:
        return {
            "status": "error",
            "message": "An unexpected error occurred while submitting the claim.",
        }

    finally:
        if workbook:
            workbook.close()
