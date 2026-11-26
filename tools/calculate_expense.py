
from decimal import Decimal, InvalidOperation
import math
from pathlib import Path
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from tools.create_excel import OUTPUT_DIR, validate_filename, resolve_workbook_path
from tools.common_func import save_workbook_safely, ensure_claim_editable
from templates.reimbursement import CALCULATOR_SHEET


def validate_number(value, field_name):
    if value is None or isinstance(value, bool):
        raise ValueError(f"{field_name} is required and must be numeric.")

    try:
        number = Decimal(str(value))
    except (InvalidOperation, ValueError):
        raise ValueError(f"{field_name} must be numeric.")

    if not number.is_finite():
        raise ValueError(f"{field_name} must be a finite number.")

    if number < 0:
        raise ValueError(f"{field_name} cannot be negative.")

    try:
        if not math.isfinite(float(number)):
            raise ValueError(f"{field_name} is outside the supported numeric range.")
    except (OverflowError, ValueError):
        raise ValueError(f"{field_name} is outside the supported numeric range.")

    return number


def calculate_expense(
    filename: str,
    calculation_type: str,
    distance_km: float = None,
    rate_per_km: float = None,
    travel_days: float = None,
    daily_allowance: float = None,
    foreign_amount: float = None,
    exchange_rate: float = None,
    currency_code: str = "",
    number1: float = None,
    number2: float = None,
    operation: str = None,
) -> dict:

    workbook = None

    try:
        filename = validate_filename(filename)
        file_path = resolve_workbook_path(filename, OUTPUT_DIR)

        if not file_path.is_file():
            raise FileNotFoundError(
                f"Workbook '{filename}' was not found."
            )

        if not isinstance(calculation_type, str):
            raise ValueError("Calculation type must be text.")
        calculation_type = calculation_type.strip().lower()

        if calculation_type not in {
            "mileage",
            "per_diem",
            "currency_conversion",
            "quick_calculation",
        }:
            raise ValueError(
                "Invalid calculation type. Choose mileage, per_diem, "
                "currency_conversion, or quick_calculation."
            )

        workbook = load_workbook(file_path)

        if CALCULATOR_SHEET not in workbook.sheetnames:
            raise ValueError(
                f"Worksheet '{CALCULATOR_SHEET}' does not exist."
            )

        ensure_claim_editable(workbook)
        sheet = workbook[CALCULATOR_SHEET]

        if calculation_type == "mileage":
            distance = validate_number(distance_km, "Distance")
            rate = validate_number(rate_per_km, "Rate per kilometre")
            result = distance * rate

            sheet["B4"] = float(distance)
            sheet["B5"] = float(rate)
            # C6 is a workbook formula; leave it intact.

            details = {
                "distance_km": float(distance),
                "rate_per_km": float(rate),
            }

        elif calculation_type == "per_diem":
            days = validate_number(travel_days, "Travel days")
            allowance = validate_number(
                daily_allowance, "Daily allowance"
            )

            if days <= 0 or days != days.to_integral_value():
                raise ValueError("Travel days must be a positive whole number.")

            result = days * allowance

            sheet["B9"] = float(days)
            sheet["B10"] = float(allowance)
            # C11 is a workbook formula; leave it intact.

            details = {
                "travel_days": float(days),
                "daily_allowance": float(allowance),
            }

        elif calculation_type == "currency_conversion":
            amount = validate_number(
                foreign_amount, "Foreign amount"
            )
            rate = validate_number(exchange_rate, "Exchange rate")

            if rate <= 0:
                raise ValueError("Exchange rate must be greater than zero.")

            if not isinstance(currency_code, str) or not currency_code.strip():
                raise ValueError("Currency code is required.")

            currency_code = currency_code.strip().upper()
            import re
            if not re.fullmatch(r"[A-Z]{3}", currency_code):
                raise ValueError("Currency code must be a three-letter ISO-style code, such as USD or INR.")

            result = amount * rate

            sheet["B14"] = float(amount)
            sheet["B15"] = currency_code
            sheet["B16"] = float(rate)
            # C17 is a workbook formula; leave it intact.

            details = {
                "foreign_amount": float(amount),
                "exchange_rate": float(rate),
                "currency_code": currency_code,
            }

        else:
            first = validate_number(number1, "First number")
            second = validate_number(number2, "Second number")

            if operation not in {"+", "-", "*", "/"}:
                raise ValueError(
                    "Invalid operation. Allowed operations: +, -, *, /."
                )

            if operation == "+":
                result = first + second
            elif operation == "-":
                result = first - second
            elif operation == "*":
                result = first * second
            else:
                if second == 0:
                    raise ValueError("Division by zero is not allowed.")
                result = first / second

            sheet["B20"] = float(first)
            sheet["B21"] = operation
            sheet["B22"] = float(second)
            # C23 is a workbook formula; leave it intact.

            details = {
                "number1": float(first),
                "number2": float(second),
                "operation": operation,
            }

        if not result.is_finite() or not math.isfinite(float(result)):
            raise ValueError("Calculation result is outside the supported numeric range.")

        save_workbook_safely(workbook, file_path)

        return {
            "status": "success",
            "filename": filename,
            "calculation_type": calculation_type,
            "details": details,
            "result": float(result),
            "message": "Expense calculation completed successfully.",
        }

    except (ValueError, FileNotFoundError) as error:
        return {
            "status": "error",
            "message": str(error),
        }

    except OSError as error:
        return {
            "status": "error",
            "message": f"Workbook could not be accessed or saved: {error}",
        }

    except (BadZipFile, InvalidFileException) as error:
        return {
            "status": "error",
            "message": f"Invalid or corrupted Excel workbook: {error}",
        }

    finally:
        if workbook is not None:
            workbook.close()
