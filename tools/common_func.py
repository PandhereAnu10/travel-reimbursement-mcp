import os
import tempfile
import math

from datetime import date, datetime
from pathlib import Path
from decimal import Decimal, InvalidOperation
from templates.reimbursement import (
    EXPENSE_HEADER_ROW,
    EXPENSE_HEADERS
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

        os.replace(temporary_path, file_path)

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
        try:
            parsed_date = date.fromisoformat(value.strip())
        except ValueError:
            raise ValueError(
                "Invalid expense date. Use YYYY-MM-DD."
            )

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

    if value.startswith(("=", "+", "-", "@")):
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
