
from openpyxl import load_workbook

from tools.create_excel import OUTPUT_DIR, validate_filename
from zipfile import BadZipFile
from openpyxl.utils.exceptions import InvalidFileException
from decimal import Decimal, InvalidOperation

from templates.reimbursement import (
    EMPLOYEE_SHEET,
    EMPLOYEE_FIELDS,
    DECLARATION_CELL,
    DECLARATION_OPTIONS,
    EXPENSE_SHEET,
    EXPENSE_HEADER_ROW,
    EXPENSE_FIRST_ROW,
    EXPENSE_LAST_ROW,
    EXPENSE_COLUMNS,
    EXPENSE_HEADERS,
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
)


EDITABLE_FIELDS = {
    "date",
    "category",
    "description",
    "amount",
    "receipt_reference",
}


def update_excel(
    filename: str,
    expense_id: str = "",
    updates: dict = None,
    employee_updates: dict = None,
    declaration: str = None,
) -> dict:
    # Update selected fields of an existing expense record

    workbook = None

    try:
        # 1. Validate filename.
        filename = validate_filename(filename)
        file_path = OUTPUT_DIR / filename

        if not file_path.is_file():
            raise FileNotFoundError(
                f"Workbook '{filename}' was not found."
            )

        # 2. Validate update payload.
        if updates is not None:
            if not isinstance(updates, dict) or not updates:
                raise ValueError(
                    "At least one field must be provided for updating."
                )

            expense_id = validate_text(
                expense_id,
                "Expense ID",
            )

            invalid_fields = set(updates.keys()) - EDITABLE_FIELDS

            if invalid_fields:
                raise ValueError(
                    f"Unsupported update fields: {sorted(invalid_fields)}. "
                    f"Allowed fields: {sorted(EDITABLE_FIELDS)}"
                )

        elif expense_id:
            raise ValueError(
                "Expense updates must be provided when an Expense ID is given."
            )

        # 3. Validate all updated values before opening the workbook.
        validated_updates = {}

        if updates is not None:
            for field, value in updates.items():

                if value is None:
                    raise ValueError(
                        f"Value for '{field}' cannot be null."
                    )

                if field == "date":
                    validated_updates[field] = validate_expense_date(value)

                elif field == "category":
                    if not isinstance(value, str):
                        raise ValueError("Category must be text.")

                    value = value.strip()

                    if value not in EXPENSE_CATEGORIES:
                        raise ValueError(
                            f"Invalid category. Allowed categories: "
                            f"{EXPENSE_CATEGORIES}"
                        )

                    validated_updates[field] = value

                elif field == "description":
                    validated_updates[field] = validate_text(
                        value,
                        "Description",
                    )

                elif field == "amount":
                    validated_updates[field] = validate_expense_amount(value)

                elif field == "receipt_reference":
                    if not isinstance(value, str):
                        raise ValueError("Receipt reference must be text.")

                    if not value.strip():
                        validated_updates[field] = None
                    else:
                        validated_updates[field] = validate_text(
                            value,
                            "Receipt reference",
                        )

        # Validate employee and travel information updates.
        validated_employee_updates = {}

        if employee_updates is not None:
            if not isinstance(employee_updates, dict) or not employee_updates:
                raise ValueError(
                    "Employee updates must be a non-empty dictionary."
                )

            invalid_fields = (
                set(employee_updates.keys()) - set(EMPLOYEE_FIELDS.keys())
            )

            if invalid_fields:
                raise ValueError(
                    f"Unsupported employee fields: {sorted(invalid_fields)}"
                )

            for field, value in employee_updates.items():
                if value is None:
                    raise ValueError(
                        f"Value for '{field}' cannot be null."
                    )

                if field in ("travel_start_date", "travel_end_date"):
                    value = validate_expense_date(value)

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

                    value = int(days)

                else:
                    value = validate_text(
                        value, EMPLOYEE_FIELDS[field]["label"]
                    )

                validated_employee_updates[field] = value

        if declaration is not None:
            if declaration not in DECLARATION_OPTIONS:
                raise ValueError(
                    f"Declaration must be one of {DECLARATION_OPTIONS}."
                )

        if (
            updates is None
            and employee_updates is None
            and declaration is None
        ):
            raise ValueError(
                "Provide expense updates, employee updates, or a declaration."
            )

        # 4. Open workbook.
        workbook = load_workbook(file_path)

        # Handle employee and travel information updates.
        if employee_updates is not None:
            if EMPLOYEE_SHEET not in workbook.sheetnames:
                raise ValueError(
                    f"Worksheet '{EMPLOYEE_SHEET}' does not exist."
                )

            employee_sheet = workbook[EMPLOYEE_SHEET]

            start_cell = EMPLOYEE_FIELDS["travel_start_date"]["cell"]
            end_cell = EMPLOYEE_FIELDS["travel_end_date"]["cell"]

            start_date = validated_employee_updates.get(
                "travel_start_date",
                employee_sheet[start_cell].value,
            )
            end_date = validated_employee_updates.get(
                "travel_end_date",
                employee_sheet[end_cell].value,
            )

            if start_date and end_date:
                start_date = validate_expense_date(start_date)
                end_date = validate_expense_date(end_date)

                if end_date < start_date:
                    raise ValueError(
                        "Travel end date cannot be earlier than travel start date."
                    )

            for field, value in validated_employee_updates.items():
                cell = EMPLOYEE_FIELDS[field]["cell"]
                employee_sheet[cell] = value

                if field in ("travel_start_date", "travel_end_date"):
                    employee_sheet[cell].number_format = DATE_FORMAT

        if declaration is not None:
            if SUMMARY_SHEET not in workbook.sheetnames:
                raise ValueError(
                    f"Worksheet '{SUMMARY_SHEET}' does not exist."
                )

            workbook[SUMMARY_SHEET][DECLARATION_CELL] = declaration

        # Handle expense record updates.
        target_row = None

        if updates is not None:
            if EXPENSE_SHEET not in workbook.sheetnames:
                raise ValueError(
                    f"Worksheet '{EXPENSE_SHEET}' does not exist."
                )

            worksheet = workbook[EXPENSE_SHEET]

            # 5. Verify worksheet headers.
            actual_headers = [
                worksheet.cell(
                    row=EXPENSE_HEADER_ROW,
                    column=column,
                ).value
                for column in range(1, len(EXPENSE_HEADERS) + 1)
            ]

            if actual_headers != EXPENSE_HEADERS:
                raise ValueError(
                    "Expense worksheet headers do not match "
                    "the configured template."
                )

            # 6. Locate the expense by its unique ID.
            id_column = EXPENSE_COLUMNS["expense_id"]
            matching_rows = []

            for row in range(
                EXPENSE_FIRST_ROW,
                EXPENSE_LAST_ROW + 1,
            ):
                current_id = worksheet.cell(
                    row=row,
                    column=id_column,
                ).value

                if (
                    current_id is not None
                    and str(current_id).strip().casefold()
                    == expense_id.casefold()
                ):
                    matching_rows.append(row)

            if not matching_rows:
                raise ValueError(
                    f"Expense ID '{expense_id}' was not found."
                )

            if len(matching_rows) > 1:
                raise ValueError(
                    f"Multiple records with Expense ID '{expense_id}' "
                    "were found. Update cancelled."
                )

            target_row = matching_rows[0]

            # 7. Check that target cells do not contain formulas.
            for field in validated_updates:
                column = EXPENSE_COLUMNS[field]
                cell = worksheet.cell(
                    row=target_row,
                    column=column,
                )

                if cell.data_type == "f":
                    raise ValueError(
                        f"Cannot update '{field}' because the target "
                        "cell contains a formula."
                    )

            # 8. Apply the validated updates.
            for field, value in validated_updates.items():
                column = EXPENSE_COLUMNS[field]

                worksheet.cell(
                    row=target_row,
                    column=column,
                ).value = value

                if field == "date":
                    worksheet.cell(
                        row=target_row,
                        column=column,
                    ).number_format = DATE_FORMAT

                elif field == "amount":
                    worksheet.cell(
                        row=target_row,
                        column=column,
                    ).number_format = CURRENCY_FORMAT

        # 9. Save safely.
        save_workbook_safely(workbook, file_path)

        if updates is not None and employee_updates is None and declaration is None:
            return {
                "status": "success",
                "filename": filename,
                "expense_id": expense_id,
                "worksheet": EXPENSE_SHEET,
                "row": target_row,
                "updated_fields": list(validated_updates.keys()),
                "message": "Expense record updated successfully.",
            }

        return {
            "status": "success",
            "filename": filename,
            "message": "Employee and travel information updated successfully.",
            "updated_fields": list(validated_employee_updates.keys())
                if employee_updates is not None else [],
            "declaration": declaration,
            "expense_id": expense_id if updates is not None else None,
            "expense_updated_fields": list(validated_updates.keys())
                if updates is not None else [],
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
