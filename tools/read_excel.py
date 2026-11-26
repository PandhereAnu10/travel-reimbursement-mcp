
from datetime import date, datetime, time
from pathlib import Path

from openpyxl import load_workbook

from tools.create_excel import OUTPUT_DIR, validate_filename, resolve_workbook_path


def make_json_compatible(value):
    #convert excel cells into json values

    if isinstance(value, (datetime, date, time)):
        return value.isoformat()

    return value


def read_excel(filename: str, sheet_name: str = None, include_formulas: bool = False) -> dict:
    # Read data from existing workbook

    # Step 1: Validate the filename.
    filename = validate_filename(filename)

    # Step 2: Check whether the workbook exists.
    file_path = resolve_workbook_path(filename, OUTPUT_DIR)

    if not file_path.exists():
        raise FileNotFoundError(
            f"Workbook '{filename}' was not found."
        )

    # Step 3: Open the workbook.
    workbook = load_workbook(
        file_path,
        data_only=False
    )

    try:
        # Step 4: Validate the requested worksheet.
        if sheet_name is not None and sheet_name not in workbook.sheetnames:
            raise ValueError(
                f"Worksheet '{sheet_name}' does not exist. "
                f"Available worksheets: {workbook.sheetnames}"
            )

        # Step 5: Decide which worksheets to read.
        selected_sheets = (
            [sheet_name]
            if sheet_name
            else workbook.sheetnames
        )

        result = {}
        contains_formulas = False

        if not isinstance(include_formulas, bool):
            raise ValueError("include_formulas must be true or false.")

        # Step 6: Read each selected worksheet.
        for current_sheet in selected_sheets:
            worksheet = workbook[current_sheet]

            rows = []

            for row in worksheet.iter_rows():
                converted_row = []
                for cell in row:
                    if cell.data_type == "f" and not include_formulas:
                        converted_row.append(None)
                    else:
                        converted_row.append(make_json_compatible(cell.value))

                # Skip rows that contain no literal values (or formulas when
                # formula inspection was explicitly requested). This prevents
                # hundreds of formula-only rows from bloating normal reads.
                if any(value is not None for value in converted_row):
                    rows.append(converted_row)

            # Include the worksheet even if it has no actual data.
            result[current_sheet] = rows
            if any(cell.data_type == "f" for row in worksheet.iter_rows() for cell in row):
                contains_formulas = True

        # Step 7: Return structured data.
        response = {
            "status": "success",
            "file_name": filename,
            "sheets": result,
        }
        if contains_formulas and include_formulas:
            response["formula_note"] = (
                "Formula cells are returned as formula expressions because include_formulas=true. "
                "openpyxl does not calculate formulas, and cached results may be absent; "
                "do not present a formula as a verified calculated value."
            )
        return response

    finally:
        # Always close the workbook, including when an error occurs.
        workbook.close()
