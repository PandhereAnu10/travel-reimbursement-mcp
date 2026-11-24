from openpyxl import load_workbook

from tools.create_excel import OUTPUT_DIR, validate_filename

from zipfile import BadZipFile
from openpyxl.utils.exceptions import InvalidFileException

from templates.reimbursement import (
    EXPENSE_SHEET,
    EXPENSE_FIRST_ROW,
    EXPENSE_LAST_ROW,
    EXPENSE_COLUMNS,
    EXPENSE_CATEGORIES,
    DATE_FORMAT,
    CURRENCY_FORMAT,
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
    expense_id: str,
    expense_date: str,
    category: str,
    description: str,
    amount: float,
    receipt_reference: str = "",
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

        # 2. Validate input fields.
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

        # 3. Open workbook.
        workbook = load_workbook(file_path)

        if EXPENSE_SHEET not in workbook.sheetnames:
            raise ValueError(
                f"Worksheet '{EXPENSE_SHEET}' does not exist."
            )

        worksheet = workbook[EXPENSE_SHEET]

        # 4. Verify worksheet structure.
        validate_expense_headers(worksheet)

        id_column = EXPENSE_COLUMNS["expense_id"]

        # 5. Check for duplicate Expense IDs.
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

        # 6. Find the next completely empty expense row.
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

        # 7. Write only the six editable columns.
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

        # 8. Apply configured display formats.
        worksheet.cell(
            row=target_row,
            column=EXPENSE_COLUMNS["date"],
        ).number_format = DATE_FORMAT

        worksheet.cell(
            row=target_row,
            column=EXPENSE_COLUMNS["amount"],
        ).number_format = CURRENCY_FORMAT

        # 9. Save safely.
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
