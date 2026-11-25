from openpyxl import load_workbook
from decimal import Decimal, InvalidOperation
from tools.create_excel import OUTPUT_DIR, validate_filename

from zipfile import BadZipFile
from openpyxl.utils.exceptions import InvalidFileException

from templates.reimbursement import (
    EMPLOYEE_SHEET,
    EMPLOYEE_FIELDS,
    DECLARATION_CELL,
    DECLARATION_OPTIONS,
    EXPENSE_SHEET,
    EXPENSE_FIRST_ROW,
    EXPENSE_LAST_ROW,
    EXPENSE_COLUMNS,
    EXPENSE_CATEGORIES,
    DATE_FORMAT,
    CURRENCY_FORMAT,
    SUMMARY_SHEET,
)

from tools.common_func import (
    save_workbook_safely,
    validate_expense_date,
    validate_expense_amount,
    validate_text,
    validate_expense_headers,
)

def write_excel(
    filename: str,
    expense_id: str = "",
    expense_date: str = "",
    category: str = "",
    description: str = "",
    amount: float = None,
    receipt_reference: str = "",
    employee_info: dict = None,
    declaration: str = None,
) -> dict:
    # Add a new expense record

    workbook = None

    try:
        # 1. Validate filename.
        filename = validate_filename(filename)
        file_path = OUTPUT_DIR / filename

        if not file_path.is_file():
            raise FileNotFoundError(
                f"Workbook '{filename}' was not found."
            )

        # 2. Open workbook.
        workbook = load_workbook(file_path)
        
        # 3. Handle employee and travel information.
        if employee_info is not None or declaration is not None:

            if EMPLOYEE_SHEET not in workbook.sheetnames:
                raise ValueError(
                    f"Worksheet '{EMPLOYEE_SHEET}' does not exist."
                )

            employee_sheet = workbook[EMPLOYEE_SHEET]

            if employee_info is not None:
                if not isinstance(employee_info, dict) or not employee_info:
                    raise ValueError(
                        "Employee information must be a non-empty dictionary."
                    )

                invalid_fields = (
                    set(employee_info.keys()) - set(EMPLOYEE_FIELDS.keys())
                )

                if invalid_fields:
                    raise ValueError(
                        f"Unsupported employee fields: {sorted(invalid_fields)}"
                    )

                validated_info = {}

                for field, value in employee_info.items():
                    if value is None:
                        raise ValueError(
                            f"Value for '{field}' cannot be null."
                        )

                    if field in ("travel_start_date", "travel_end_date"):
                        validated_info[field] = validate_expense_date(value)

                    elif field == "travel_days":
                        if isinstance(value, bool):
                            raise ValueError(
                                "Travel days must be a positive whole number."
                            )

                        try:
                            days = Decimal(str(value))
                        except (InvalidOperation, ValueError, TypeError):
                            raise ValueError(
                                "Travel days must be a positive whole number."
                            )

                        if (
                            not days.is_finite()
                            or days <= 0
                            or days != days.to_integral_value()
                        ):
                            raise ValueError(
                                "Travel days must be a positive whole number."
                            )

                        validated_info[field] = int(days)

                    else:
                        validated_info[field] = validate_text(
                            value, EMPLOYEE_FIELDS[field]["label"]
                        )

                for field, value in validated_info.items():
                    cell = EMPLOYEE_FIELDS[field]["cell"]
                    employee_sheet[cell] = value

                    if field in ("travel_start_date", "travel_end_date"):
                        employee_sheet[cell].number_format = DATE_FORMAT

            if declaration is not None:
                if declaration not in DECLARATION_OPTIONS:
                    raise ValueError(
                        f"Declaration must be one of {DECLARATION_OPTIONS}."
                    )

                if SUMMARY_SHEET not in workbook.sheetnames:
                    raise ValueError(
                        f"Worksheet '{SUMMARY_SHEET}' does not exist."
                    )

                workbook[SUMMARY_SHEET][DECLARATION_CELL] = declaration

            save_workbook_safely(workbook, file_path)

            return {
                "status": "success",
                "filename": filename,
                "message": "Employee and travel information saved successfully.",
                "updated_fields": list(employee_info.keys())
                    if employee_info else [],
                "declaration": declaration,
            }

        
        # 4. Validate input fields.
        expense_id = validate_text(
            expense_id,
            "Expense ID",
        )

        description = validate_text(
            description,
            "Description",
        )

        if not isinstance(category, str):
            raise ValueError("Category must be text.")

        category = category.strip()

        if category not in EXPENSE_CATEGORIES:
            raise ValueError(
                f"Invalid category. Allowed categories: "
                f"{EXPENSE_CATEGORIES}"
            )

        parsed_date = validate_expense_date(expense_date)
        parsed_amount = validate_expense_amount(amount)

        if receipt_reference is None:
            receipt_reference = ""

        if not isinstance(receipt_reference, str):
            raise ValueError(
                "Receipt reference must be text."
            )

        receipt_reference = receipt_reference.strip()

        if receipt_reference.startswith(("=", "+", "-", "@")):
            raise ValueError(
                "Receipt reference cannot start with a "
                "spreadsheet command character."
            )

        if EXPENSE_SHEET not in workbook.sheetnames:
            raise ValueError(
                f"Worksheet '{EXPENSE_SHEET}' does not exist."
            )

        worksheet = workbook[EXPENSE_SHEET]

        # 5. Verify worksheet structure.
        validate_expense_headers(worksheet)

        id_column = EXPENSE_COLUMNS["expense_id"]

        # 6. Check for duplicate Expense IDs.
        for row in range(
            EXPENSE_FIRST_ROW,
            EXPENSE_LAST_ROW + 1,
        ):
            existing_id = worksheet.cell(
                row=row,
                column=id_column,
            ).value

            if (
                existing_id is not None
                and str(existing_id).strip().casefold()
                == expense_id.casefold()
            ):
                raise ValueError(
                    f"Expense ID '{expense_id}' already exists."
                )

        # 7. Find the next completely empty expense row.
        target_row = None

        editable_columns = [
            EXPENSE_COLUMNS[field]
            for field in (
                "expense_id",
                "date",
                "category",
                "description",
                "amount",
                "receipt_reference",
            )
        ]

        for row in range(
            EXPENSE_FIRST_ROW,
            EXPENSE_LAST_ROW + 1,
        ):
            if all(
                worksheet.cell(row=row, column=column).value is None
                for column in editable_columns
            ):
                target_row = row
                break

        if target_row is None:
            raise ValueError(
                "Expense worksheet is full. "
                "No empty expense rows remain."
            )

        # 8. Write only the six editable columns.
        # The existing Date Check formula in column G is untouched.
        values = {
            "expense_id": expense_id,
            "date": parsed_date,
            "category": category,
            "description": description,
            "amount": parsed_amount,
            "receipt_reference": receipt_reference or None,
        }

        for field, value in values.items():
            column = EXPENSE_COLUMNS[field]
            worksheet.cell(
                row=target_row,
                column=column,
            ).value = value

        # 9. Apply configured display formats.
        worksheet.cell(
            row=target_row,
            column=EXPENSE_COLUMNS["date"],
        ).number_format = DATE_FORMAT

        worksheet.cell(
            row=target_row,
            column=EXPENSE_COLUMNS["amount"],
        ).number_format = CURRENCY_FORMAT

        # 10. Save safely.
        save_workbook_safely(workbook, file_path)

        return {
            "status": "success",
            "filename": filename,
            "expense_id": expense_id,
            "worksheet": EXPENSE_SHEET,
            "row": target_row,
            "message": "Expense record added successfully.",
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
