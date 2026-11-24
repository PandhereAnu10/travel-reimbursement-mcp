from openpyxl.styles import (
    Font,
    PatternFill,
    Border,
    Side,
    Alignment,
)

from templates.reimbursement import (
    EMPLOYEE_SHEET,
    EXPENSE_SHEET,
    CALCULATOR_SHEET,
    SUMMARY_SHEET,
    EMPLOYEE_TITLE_CELL,
    EXPENSE_TITLE_CELL,
    CALCULATOR_TITLE_CELL,
    SUMMARY_TITLE_CELL,
    EMPLOYEE_FIELDS,
    EMPLOYEE_LABEL_COLUMN,
    EXPENSE_HEADER_ROW,
    EXPENSE_FIRST_ROW,
    EXPENSE_LAST_ROW,
    EXPENSE_TOTAL_CELL,
    EXPENSE_HEADERS,
    EXPENSE_COLUMNS,
    MILEAGE_DISTANCE_CELL,
    MILEAGE_RATE_CELL,
    MILEAGE_RESULT_CELL,
    PER_DIEM_DAYS_CELL,
    PER_DIEM_ALLOWANCE_CELL,
    PER_DIEM_RESULT_CELL,
    CURRENCY_AMOUNT_CELL,
    CURRENCY_CODE_CELL,
    CURRENCY_RATE_CELL,
    CURRENCY_RESULT_CELL,
    QUICK_CALC_FIRST_VALUE_CELL,
    QUICK_CALC_OPERATION_CELL,
    QUICK_CALC_SECOND_VALUE_CELL,
    QUICK_CALC_RESULT_CELL,
    SUMMARY_CATEGORY_START_ROW,
    SUMMARY_CATEGORY_END_ROW,
    SUMMARY_CATEGORY_AMOUNT_COLUMN,
    TOTAL_CLAIMED_CELL,
    ADVANCE_RECEIVED_CELL,
    NET_PAYABLE_CELL,
    DECLARATION_CELL,
    EMPLOYEE_SIGNATURE_CELL,
    SUBMISSION_STATUS_CELL,
    SUBMISSION_DATE_CELL,
    CLAIM_REFERENCE_ID_CELL,
    HR_NOTES_CELL,
    DATE_FORMAT,
    CURRENCY_FORMAT,
    INPUT_FILL_COLOR,
    HEADER_FILL_COLOR,
    OUTPUT_FILL_COLOR,
    BORDER_COLOR,
)


# Shared styles used across all four worksheets.
HEADER_FILL = PatternFill(
    fill_type="solid",
    fgColor=HEADER_FILL_COLOR,
)

INPUT_FILL = PatternFill(
    fill_type="solid",
    fgColor=INPUT_FILL_COLOR,
)

OUTPUT_FILL = PatternFill(
    fill_type="solid",
    fgColor=OUTPUT_FILL_COLOR,
)

THIN_BORDER = Border(
    left=Side(style="thin", color=BORDER_COLOR),
    right=Side(style="thin", color=BORDER_COLOR),
    top=Side(style="thin", color=BORDER_COLOR),
    bottom=Side(style="thin", color=BORDER_COLOR),
)


# Applies the common title style to a worksheet.
def style_title(worksheet, cell_reference):
    cell = worksheet[cell_reference]
    cell.font = Font(size=16, bold=True, color="FFFFFFFF")
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(vertical="center")
    worksheet.row_dimensions[cell.row].height = 30


# Formats a row containing column headings.
def style_header_row(worksheet, row, start_column, end_column):
    for column in range(start_column, end_column + 1):
        cell = worksheet.cell(row=row, column=column)
        cell.font = Font(bold=True, color="FFFFFFFF")
        cell.fill = HEADER_FILL
        cell.border = THIN_BORDER
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )


# Distinguishes cells where users are expected to enter information.
def style_input_cell(cell):
    cell.fill = INPUT_FILL
    cell.border = THIN_BORDER
    cell.alignment = Alignment(vertical="center")


# Distinguishes calculated or automatically populated cells.
def style_output_cell(cell):
    cell.fill = OUTPUT_FILL
    cell.border = THIN_BORDER
    cell.font = Font(bold=True)


# Sets column widths to keep worksheet content readable.
def set_column_widths(worksheet, widths):
    for column, width in widths.items():
        worksheet.column_dimensions[column].width = width


# Formats employee details and highlights the calculated travel-days cell.
def format_employee_sheet(worksheet):
    style_title(worksheet, EMPLOYEE_TITLE_CELL)
    worksheet.merge_cells("A1:B1")

    set_column_widths(worksheet, {"A": 30, "B": 38})

    for field in EMPLOYEE_FIELDS.values():
        label_cell = worksheet.cell(
            row=worksheet[field["cell"]].row,
            column=1,
        )
        value_cell = worksheet[field["cell"]]

        label_cell.value = field["label"]
        label_cell.font = Font(bold=True)
        label_cell.border = THIN_BORDER

        if field["cell"] == EMPLOYEE_FIELDS["travel_days"]["cell"]:
            style_output_cell(value_cell)
        else:
            style_input_cell(value_cell)

        # Displays travel dates in the configured date format.
        if field["cell"] in (
            EMPLOYEE_FIELDS["travel_start_date"]["cell"],
            EMPLOYEE_FIELDS["travel_end_date"]["cell"],
        ):
            value_cell.number_format = DATE_FORMAT

    worksheet.freeze_panes = "A3"


# Formats the expense table, including its input rows and calculated date checks.
def format_expense_sheet(worksheet):
    style_title(worksheet, EXPENSE_TITLE_CELL)
    worksheet.merge_cells("A1:G1")

    worksheet["D2"] = "Total Expenses"
    worksheet["D2"].font = Font(bold=True)
    worksheet[EXPENSE_TOTAL_CELL].number_format = CURRENCY_FORMAT
    style_output_cell(worksheet[EXPENSE_TOTAL_CELL])

    # Populate the expense table headings from the shared configuration.
    for column, header in enumerate(EXPENSE_HEADERS, start=1):
        worksheet.cell(
            row=EXPENSE_HEADER_ROW,
            column=column,
            value=header,
        )

    style_header_row(
        worksheet,
        EXPENSE_HEADER_ROW,
        1,
        len(EXPENSE_HEADERS),
    )

    set_column_widths(
        worksheet,
        {
            "A": 16,
            "B": 18,
            "C": 23,
            "D": 35,
            "E": 20,
            "F": 25,
            "G": 25,
        },
    )

    # Apply formatting only to the 200 configured expense-entry rows.
    for row in range(EXPENSE_FIRST_ROW, EXPENSE_LAST_ROW + 1):
        for column in range(1, len(EXPENSE_HEADERS) + 1):
            cell = worksheet.cell(row=row, column=column)

            if column == EXPENSE_COLUMNS["date_check"]:
                style_output_cell(cell)
            else:
                style_input_cell(cell)

        worksheet.cell(
            row=row,
            column=EXPENSE_COLUMNS["date"],
        ).number_format = DATE_FORMAT

        worksheet.cell(
            row=row,
            column=EXPENSE_COLUMNS["amount"],
        ).number_format = CURRENCY_FORMAT

    worksheet.freeze_panes = "A5"
    worksheet.auto_filter.ref = (
        f"A{EXPENSE_HEADER_ROW}:G{EXPENSE_LAST_ROW}"
    )


# Formats the calculator inputs and highlights formula result cells.
def format_calculator_sheet(worksheet):
    style_title(worksheet, CALCULATOR_TITLE_CELL)
    worksheet.merge_cells("A1:C1")

    set_column_widths(
        worksheet,
        {
            "A": 32,
            "B": 22,
            "C": 24,
        },
    )

    labels = {
        "A3": "Mileage Calculator",
        "A4": "Distance (km)",
        "A5": "Rate per km",
        "B6": "Mileage Amount",
        "A8": "Per Diem Calculator",
        "A9": "Number of Days",
        "A10": "Daily Allowance",
        "B11": "Per Diem Amount",
        "A13": "Currency Conversion",
        "A14": "Foreign Amount",
        "A15": "Currency Code",
        "A16": "Exchange Rate (INR)",
        "B17": "Converted Amount",
        "A19": "Quick Calculator",
        "A20": "First Value",
        "A21": "Operation",
        "A22": "Second Value",
        "B23": "Result",
    }

    for cell_reference, label in labels.items():
        worksheet[cell_reference] = label

    # Give each calculator section a distinct heading.
    for row in (3, 8, 13, 19):
        cell = worksheet[f"A{row}"]
        cell.font = Font(bold=True, color="FFFFFFFF")
        cell.fill = HEADER_FILL

    input_cells = (
        MILEAGE_DISTANCE_CELL,
        MILEAGE_RATE_CELL,
        PER_DIEM_DAYS_CELL,
        PER_DIEM_ALLOWANCE_CELL,
        CURRENCY_AMOUNT_CELL,
        CURRENCY_CODE_CELL,
        CURRENCY_RATE_CELL,
        QUICK_CALC_FIRST_VALUE_CELL,
        QUICK_CALC_OPERATION_CELL,
        QUICK_CALC_SECOND_VALUE_CELL,
    )

    for cell_reference in input_cells:
        style_input_cell(worksheet[cell_reference])

    output_cells = (
        MILEAGE_RESULT_CELL,
        PER_DIEM_RESULT_CELL,
        CURRENCY_RESULT_CELL,
        QUICK_CALC_RESULT_CELL,
    )

    for cell_reference in output_cells:
        style_output_cell(worksheet[cell_reference])

    # Apply currency formatting to monetary inputs and outputs.
    for cell_reference in (
        MILEAGE_RATE_CELL,
        MILEAGE_RESULT_CELL,
        PER_DIEM_ALLOWANCE_CELL,
        PER_DIEM_RESULT_CELL,
        CURRENCY_AMOUNT_CELL,
        CURRENCY_RESULT_CELL,
    ):
        worksheet[cell_reference].number_format = CURRENCY_FORMAT

    worksheet.freeze_panes = "A3"


# Formats the category summary, claim totals and HR submission fields.
def format_summary_sheet(worksheet):
    style_title(worksheet, SUMMARY_TITLE_CELL)
    worksheet.merge_cells("A1:B1")

    set_column_widths(
        worksheet,
        {
            "A": 45,
            "B": 30,
        },
    )

    worksheet["A3"] = "Expense Category"
    worksheet.cell(
        row=SUMMARY_CATEGORY_START_ROW - 1,
        column=SUMMARY_CATEGORY_AMOUNT_COLUMN,
        value="Amount (INR)",
    )

    style_header_row(
        worksheet,
        SUMMARY_CATEGORY_START_ROW - 1,
        1,
        2,
    )

    labels = {
        "A9": "Total Amount Claimed",
        "A10": "Advance Received",
        "A11": "Net Amount Payable",
        "A13": "Employee Declaration Accepted?",
        "A14": "Employee Signature",
        "A15": "Submission Status",
        "A16": "Submission Date",
        "A17": "Claim Reference ID",
        "A18": "Notes for HR",
    }

    for cell_reference, label in labels.items():
        worksheet[cell_reference] = label
        worksheet[cell_reference].font = Font(bold=True)
        worksheet[cell_reference].border = THIN_BORDER

    # Format the category-wise calculated amounts.
    for row in range(
        SUMMARY_CATEGORY_START_ROW,
        SUMMARY_CATEGORY_END_ROW + 1,
    ):
        worksheet.cell(row=row, column=1).border = THIN_BORDER
        worksheet.cell(
            row=row,
            column=SUMMARY_CATEGORY_AMOUNT_COLUMN,
        ).number_format = CURRENCY_FORMAT
        style_output_cell(
            worksheet.cell(
                row=row,
                column=SUMMARY_CATEGORY_AMOUNT_COLUMN,
            )
        )

    # Mark fields that require user input.
    for cell_reference in (
        ADVANCE_RECEIVED_CELL,
        DECLARATION_CELL,
        SUBMISSION_STATUS_CELL,
        HR_NOTES_CELL,
    ):
        style_input_cell(worksheet[cell_reference])

    # Mark fields populated by formulas or the submission process.
    for cell_reference in (
        TOTAL_CLAIMED_CELL,
        NET_PAYABLE_CELL,
        EMPLOYEE_SIGNATURE_CELL,
        SUBMISSION_DATE_CELL,
        CLAIM_REFERENCE_ID_CELL,
    ):
        style_output_cell(worksheet[cell_reference])

    for cell_reference in (
        TOTAL_CLAIMED_CELL,
        ADVANCE_RECEIVED_CELL,
        NET_PAYABLE_CELL,
    ):
        worksheet[cell_reference].number_format = CURRENCY_FORMAT

    worksheet[SUBMISSION_DATE_CELL].number_format = DATE_FORMAT
    worksheet.freeze_panes = "A4"


# Entry point called by create_excel.py to format all four worksheets.
def format_workbook(workbook):
    format_employee_sheet(workbook[EMPLOYEE_SHEET])
    format_expense_sheet(workbook[EXPENSE_SHEET])
    format_calculator_sheet(workbook[CALCULATOR_SHEET])
    format_summary_sheet(workbook[SUMMARY_SHEET])