import os
import tempfile
import math
import time

from datetime import date, datetime
from pathlib import Path
from decimal import Decimal, InvalidOperation
from templates.reimbursement import (
    EXPENSE_HEADER_ROW,
    EXPENSE_HEADERS,
    EMPLOYEE_SHEET,
    EMPLOYEE_FIELDS,
    SUMMARY_SHEET,
    SUBMISSION_STATUS_CELL,
    SUBMITTED_STATUS,
)

def save_workbook_safely(workbook, file_path):

    temporary_path = None

    try:
        with tempfile.NamedTemporaryFile(
            suffix=".xlsx",
            prefix=".reimbursement_",
            dir=file_path.parent,
            delete=False,
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)

        workbook.save(temporary_path)

        for attempt in range(3):
            try:
                os.replace(temporary_path, file_path)
                break
            except PermissionError as error:
                if attempt == 2:
                    raise PermissionError(
                        "Windows could not replace the workbook. Close the Excel file "
                        "and any other application using it, check folder permissions, "
                        "and try again."
                    ) from error
                time.sleep(0.15 * (attempt + 1))

    finally:
        if temporary_path and temporary_path.exists():
            temporary_path.unlink()


def validate_expense_date(value):
    """Validate and convert an expense date."""

    if isinstance(value, datetime):
        parsed_date = value.date()

    elif isinstance(value, date):
        parsed_date = value

    elif isinstance(value, str):
        import re
        normalized = value.strip()
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", normalized):
            raise ValueError("Invalid expense date. Use YYYY-MM-DD.")
        try:
            parsed_date = date.fromisoformat(normalized)
        except ValueError:
            raise ValueError("Invalid expense date. Use YYYY-MM-DD.")

    else:
        raise ValueError(
            "Expense date must be provided as YYYY-MM-DD."
        )

    if not date(2000, 1, 1) <= parsed_date <= date(2100, 12, 31):
        raise ValueError(
            "Expense date must be between 2000 and 2100."
        )

    return parsed_date


def validate_expense_amount(value):
    # validate non-negative finite numeric amount

    if isinstance(value, bool):
        raise ValueError("Amount must be numeric.")

    try:
        amount = Decimal(str(value))
    except (InvalidOperation, ValueError, TypeError):
        raise ValueError("Amount must be a valid number.")

    if not amount.is_finite() or amount < 0:
        raise ValueError(
            "Amount must be finite and non-negative."
        )

    if not math.isfinite(float(amount)):
        raise ValueError("Amount is outside the supported range.")

    return float(amount)


def validate_text(value, field_name):
    # Validate required text fields and reject formula-like input

    if not isinstance(value, str) or not value.strip():
        raise ValueError(
            f"{field_name} cannot be empty."
        )

    value = value.strip()

    if value.lstrip().startswith(("=", "+", "-", "@")):
        raise ValueError(
            f"{field_name} cannot start with a formula or "
            "spreadsheet command character."
        )

    return value


def validate_expense_headers(worksheet):
    # Ensure the worksheet matches configured expense layout
    actual_headers = [
        worksheet.cell(
            row=EXPENSE_HEADER_ROW,
            column=column,
        ).value
        for column in range(1, len(EXPENSE_HEADERS) + 1)
    ]

    if actual_headers != EXPENSE_HEADERS:
        raise ValueError(
            "Expense worksheet headers do not match the configured template."
        )


def ensure_claim_editable(workbook):
    if SUMMARY_SHEET in workbook.sheetnames:
        status = workbook[SUMMARY_SHEET][SUBMISSION_STATUS_CELL].value
        if status == SUBMITTED_STATUS:
            raise ValueError("This claim has already been submitted to HR and cannot be modified.")


def get_travel_period(workbook, required=False):
    if EMPLOYEE_SHEET not in workbook.sheetnames:
        if required:
            raise ValueError(
                "The employee/travel worksheet is missing; travel start and end dates "
                "must be recorded before adding or changing an expense date."
            )
        return None, None
    sheet = workbook[EMPLOYEE_SHEET]
    start = sheet[EMPLOYEE_FIELDS["travel_start_date"]["cell"]].value
    end = sheet[EMPLOYEE_FIELDS["travel_end_date"]["cell"]].value
    if start in (None, "") or end in (None, ""):
        if required:
            raise ValueError(
                "Travel start and end dates must be recorded before adding "
                "or changing an expense date."
            )
        return None, None
    start = validate_expense_date(start)
    end = validate_expense_date(end)
    if end < start:
        raise ValueError("Travel end date cannot be earlier than travel start date.")
    return start, end


def validate_date_within_travel_period(expense_date, workbook, require_period=False):
    start, end = get_travel_period(workbook, required=require_period)
    if start is not None and not start <= expense_date <= end:
        raise ValueError("Expense date must fall between the travel start date and travel end date.")


def validate_email(value):
    import re
    if not isinstance(value, str) or not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", value.strip()):
        raise ValueError("Email must be a valid email address.")
    return value.strip()
