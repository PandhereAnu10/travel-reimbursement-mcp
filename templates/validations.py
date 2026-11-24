from openpyxl.worksheet.datavalidation import DataValidation

from templates.reimbursement import (
    EMPLOYEE_SHEET,
    EXPENSE_SHEET,
    CALCULATOR_SHEET,
    SUMMARY_SHEET,
    EMPLOYEE_FIELDS,
    EXPENSE_FIRST_ROW,
    EXPENSE_LAST_ROW,
    EXPENSE_COLUMNS,
    EXPENSE_CATEGORIES,
    QUICK_CALC_OPERATIONS,
    DECLARATION_OPTIONS,
    SUBMISSION_STATUS_OPTIONS,
    MILEAGE_DISTANCE_CELL,
    MILEAGE_RATE_CELL,
    PER_DIEM_ALLOWANCE_CELL,
    CURRENCY_AMOUNT_CELL,
    CURRENCY_RATE_CELL,
    QUICK_CALC_FIRST_VALUE_CELL,
    QUICK_CALC_SECOND_VALUE_CELL,
    QUICK_CALC_OPERATION_CELL,
    PER_DIEM_DAYS_CELL,
    ADVANCE_RECEIVED_CELL,
    DECLARATION_CELL,
    SUBMISSION_STATUS_CELL,
)


# Creates a dropdown list for cells with a predefined set of allowed values.
def add_dropdown(worksheet, cell_range, options, prompt):
    formula = '"' + ",".join(options) + '"'

    validation = DataValidation(
        type="list",
        formula1=formula,
        allow_blank=True,
    )

    validation.errorTitle = "Invalid Selection"
    validation.error = "Please select a valid option."
    validation.promptTitle = "Select a Value"
    validation.prompt = prompt
    validation.showErrorMessage = True
    validation.showInputMessage = True

    worksheet.add_data_validation(validation)
    validation.add(cell_range)


# Restricts numeric input to non-negative decimal values.
def add_decimal_validation(worksheet, cell_range):
    validation = DataValidation(
        type="decimal",
        operator="greaterThanOrEqual",
        formula1="0",
        allow_blank=True,
    )

    validation.errorTitle = "Invalid Number"
    validation.error = "Enter a non-negative number."
    validation.showErrorMessage = True

    worksheet.add_data_validation(validation)
    validation.add(cell_range)


# Restricts date input to the configured supported date range.
def add_date_validation(worksheet, cell_range):
    validation = DataValidation(
        type="date",
        operator="between",
        formula1="DATE(2000,1,1)",
        formula2="DATE(2100,12,31)",
        allow_blank=True,
    )

    validation.errorTitle = "Invalid Date"
    validation.error = "Enter a valid date between 2000 and 2100."
    validation.showErrorMessage = True

    worksheet.add_data_validation(validation)
    validation.add(cell_range)


# Validates employee travel dates and ensures the end date is not earlier.
def add_employee_validations(workbook):
    worksheet = workbook[EMPLOYEE_SHEET]

    start_date = EMPLOYEE_FIELDS["travel_start_date"]["cell"]
    end_date = EMPLOYEE_FIELDS["travel_end_date"]["cell"]

    add_date_validation(worksheet, start_date)
    add_date_validation(worksheet, end_date)

    validation = DataValidation(
        type="custom",
        formula1=f'=OR({end_date}="",{start_date}="",{end_date}>={start_date})',
        allow_blank=True,
    )

    validation.errorTitle = "Invalid Travel Dates"
    validation.error = "Travel end date cannot be earlier than travel start date."
    validation.showErrorMessage = True

    worksheet.add_data_validation(validation)
    validation.add(end_date)


# Adds category, date and amount restrictions to the expense-entry rows.
def add_expense_validations(workbook):
    worksheet = workbook[EXPENSE_SHEET]

    category_column = EXPENSE_COLUMNS["category"]
    date_column = EXPENSE_COLUMNS["date"]
    amount_column = EXPENSE_COLUMNS["amount"]

    category_letter = worksheet.cell(
        row=EXPENSE_FIRST_ROW,
        column=category_column,
    ).column_letter

    date_letter = worksheet.cell(
        row=EXPENSE_FIRST_ROW,
        column=date_column,
    ).column_letter

    amount_letter = worksheet.cell(
        row=EXPENSE_FIRST_ROW,
        column=amount_column,
    ).column_letter

    # Expense categories must be selected from the configured list.
    add_dropdown(
        worksheet,
        f"{category_letter}{EXPENSE_FIRST_ROW}:"
        f"{category_letter}{EXPENSE_LAST_ROW}",
        EXPENSE_CATEGORIES,
        "Select an expense category.",
    )

    add_date_validation(
        worksheet,
        f"{date_letter}{EXPENSE_FIRST_ROW}:"
        f"{date_letter}{EXPENSE_LAST_ROW}",
    )

    add_decimal_validation(
        worksheet,
        f"{amount_letter}{EXPENSE_FIRST_ROW}:"
        f"{amount_letter}{EXPENSE_LAST_ROW}",
    )


# Validates numeric inputs and the arithmetic operation in the calculator.
def add_calculator_validations(workbook):
    worksheet = workbook[CALCULATOR_SHEET]

    numeric_cells = (
        MILEAGE_DISTANCE_CELL,
        MILEAGE_RATE_CELL,
        PER_DIEM_ALLOWANCE_CELL,
        CURRENCY_AMOUNT_CELL,
        CURRENCY_RATE_CELL,
        QUICK_CALC_FIRST_VALUE_CELL,
        QUICK_CALC_SECOND_VALUE_CELL,
    )

    for cell in numeric_cells:
        add_decimal_validation(worksheet, cell)

    # Per diem days must be a non-negative whole number.
    days_validation = DataValidation(
        type="whole",
        operator="greaterThanOrEqual",
        formula1="0",
        allow_blank=True,
    )

    days_validation.errorTitle = "Invalid Days"
    days_validation.error = "Enter a non-negative whole number."
    days_validation.showErrorMessage = True

    worksheet.add_data_validation(days_validation)
    days_validation.add(PER_DIEM_DAYS_CELL)

    add_dropdown(
        worksheet,
        QUICK_CALC_OPERATION_CELL,
        QUICK_CALC_OPERATIONS,
        "Choose an arithmetic operation.",
    )


# Adds declaration, submission-status and advance-amount restrictions.
def add_summary_validations(workbook):
    worksheet = workbook[SUMMARY_SHEET]

    add_dropdown(
        worksheet,
        DECLARATION_CELL,
        DECLARATION_OPTIONS,
        "Select Yes or No to confirm the declaration.",
    )

    add_dropdown(
        worksheet,
        SUBMISSION_STATUS_CELL,
        SUBMISSION_STATUS_OPTIONS,
        "Select the claim submission status.",
    )

    add_decimal_validation(
        worksheet,
        ADVANCE_RECEIVED_CELL,
    )


# Entry point called by create_excel.py to apply all worksheet validations.
def add_workbook_validations(workbook):
    add_employee_validations(workbook)
    add_expense_validations(workbook)
    add_calculator_validations(workbook)
    add_summary_validations(workbook)