
from datetime import date, datetime, time
from pathlib import Path

from openpyxl import load_workbook

from tools.create_excel import OUTPUT_DIR, validate_filename


def make_json_compatible(value):
    #convert excel cells into json values

    if isinstance(value, (datetime, date, time)):
        return value.isoformat()

    return value


def read_excel(filename: str, sheet_name: str = None) -> dict:
    # Read data from existing workbook

    # Step 1: Validate the filename.
    filename = validate_filename(filename)

    # Step 2: Check whether the workbook exists.
    file_path = OUTPUT_DIR / filename

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

        # Step 6: Read each selected worksheet.
        for current_sheet in selected_sheets:
            worksheet = workbook[current_sheet]

            rows = []

            for row in worksheet.iter_rows(values_only=True):
                converted_row = [
                    make_json_compatible(value)
                    for value in row
                ]

                has_actual_data = any(
                    value is not None
                    and not (
                        isinstance(value, str)
                        and value.startswith("=")
                    )
                    for value in converted_row
                )

                if has_actual_data:
                    rows.append(converted_row)
                    result[current_sheet] = rows

        # Step 7: Return structured data.
        return {
            "status": "success",
            "file_name": filename,
            "sheets": result,
        }

    finally:
        # Always close the workbook, including when an error occurs.
        workbook.close()
