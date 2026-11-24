from openpyxl.workbook.properties import CalcProperties

from templates.reimbursement import (
    EMPLOYEE_SHEET,
    EXPENSE_SHEET,
    CALCULATOR_SHEET,
    SUMMARY_SHEET,
    EMPLOYEE_FIELDS,
    EXPENSE_FIRST_ROW,
    EXPENSE_LAST_ROW,
    EXPENSE_TOTAL_CELL,
    EXPENSE_COLUMNS,
    EXPENSE_CATEGORIES,
    SUMMARY_CATEGORY_START_ROW,
    SUMMARY_CATEGORY_END_ROW,
    MILEAGE_DISTANCE_CELL,
    MILEAGE_RATE_CELL,
    MILEAGE_RESULT_CELL,
    PER_DIEM_DAYS_CELL,
    PER_DIEM_ALLOWANCE_CELL,
    PER_DIEM_RESULT_CELL,
    CURRENCY_AMOUNT_CELL,
    CURRENCY_RATE_CELL,
    CURRENCY_RESULT_CELL,
    QUICK_CALC_FIRST_VALUE_CELL,
    QUICK_CALC_OPERATION_CELL,
    QUICK_CALC_SECOND_VALUE_CELL,
    QUICK_CALC_RESULT_CELL,
    TOTAL_CLAIMED_CELL,
    ADVANCE_RECEIVED_CELL,
    NET_PAYABLE_CELL,
    EMPLOYEE_SIGNATURE_CELL,
    WITHIN_TRAVEL_DATES_FLAG,
    OUTSIDE_TRAVEL_DATES_FLAG,
    MISSING_TRAVEL_DATES_FLAG,
)


# Calculates the total number of travel days, including both start and end dates.
def add_employee_formulas(workbook):
    worksheet = workbook[EMPLOYEE_SHEET]

    start_date = EMPLOYEE_FIELDS["travel_start_date"]["cell"]
    end_date = EMPLOYEE_FIELDS["travel_end_date"]["cell"]
    travel_days = EMPLOYEE_FIELDS["travel_days"]["cell"]

    worksheet[travel_days] = (
        f'=IF(OR({start_date}="",{end_date}=""),"",'
        f'IF({end_date}<{start_date},"Invalid dates",'
        f'{end_date}-{start_date}+1))'
    )


# Checks whether each expense date falls within the employee's travel period.
def add_expense_formulas(workbook):
    worksheet = workbook[EXPENSE_SHEET]

    start_date = EMPLOYEE_FIELDS["travel_start_date"]["cell"]
    end_date = EMPLOYEE_FIELDS["travel_end_date"]["cell"]

    start_ref = f"'{EMPLOYEE_SHEET}'!${start_date[0]}${start_date[1:]}"
    end_ref = f"'{EMPLOYEE_SHEET}'!${end_date[0]}${end_date[1:]}"

    date_column = EXPENSE_COLUMNS["date"]
    check_column = EXPENSE_COLUMNS["date_check"]
    amount_column = EXPENSE_COLUMNS["amount"]

    for row in range(EXPENSE_FIRST_ROW, EXPENSE_LAST_ROW + 1):
        expense_date = worksheet.cell(
            row=row,
            column=date_column,
        ).coordinate

        check_cell = worksheet.cell(
            row=row,
            column=check_column,
        )

        check_cell.value = (
            f'=IF({expense_date}="","",'
            f'IF(OR({start_ref}="",{end_ref}=""),'
            f'"{MISSING_TRAVEL_DATES_FLAG}",'
            f'IF(AND({expense_date}>={start_ref},'
            f'{expense_date}<={end_ref}),'
            f'"{WITHIN_TRAVEL_DATES_FLAG}",'
            f'"{OUTSIDE_TRAVEL_DATES_FLAG}")))'
        )

    # Calculates the total of all expense amounts in the configured rows.
    amount_letter = worksheet.cell(
        row=EXPENSE_FIRST_ROW,
        column=amount_column,
    ).column_letter

    worksheet[EXPENSE_TOTAL_CELL] = (
        f'=SUM({amount_letter}{EXPENSE_FIRST_ROW}:'
        f'{amount_letter}{EXPENSE_LAST_ROW})'
    )


# Adds formulas for mileage, per diem, currency conversion and arithmetic.
def add_calculator_formulas(workbook):
    worksheet = workbook[CALCULATOR_SHEET]

    # Mileage amount = distance travelled × rate per kilometre.
    worksheet[MILEAGE_RESULT_CELL] = (
        f'=IF(OR({MILEAGE_DISTANCE_CELL}="",'
        f'{MILEAGE_RATE_CELL}=""),"",'
        f'{MILEAGE_DISTANCE_CELL}*{MILEAGE_RATE_CELL})'
    )

    # Automatically uses the employee's calculated travel days.
    travel_days = EMPLOYEE_FIELDS["travel_days"]["cell"]

    worksheet[PER_DIEM_DAYS_CELL] = (
        f'=IF(\'{EMPLOYEE_SHEET}\'!{travel_days}="","",'
        f'\'{EMPLOYEE_SHEET}\'!{travel_days})'
    )

    # Per diem amount = number of days × daily allowance.
    worksheet[PER_DIEM_RESULT_CELL] = (
        f'=IF(OR({PER_DIEM_DAYS_CELL}="",'
        f'{PER_DIEM_ALLOWANCE_CELL}=""),"",'
        f'{PER_DIEM_DAYS_CELL}*{PER_DIEM_ALLOWANCE_CELL})'
    )

    # Converts the foreign currency amount into INR.
    worksheet[CURRENCY_RESULT_CELL] = (
        f'=IF(OR({CURRENCY_AMOUNT_CELL}="",'
        f'{CURRENCY_RATE_CELL}=""),"",'
        f'{CURRENCY_AMOUNT_CELL}*{CURRENCY_RATE_CELL})'
    )

    # Performs the selected arithmetic operation on two input values.
    first = QUICK_CALC_FIRST_VALUE_CELL
    operation = QUICK_CALC_OPERATION_CELL
    second = QUICK_CALC_SECOND_VALUE_CELL
    result = QUICK_CALC_RESULT_CELL

    worksheet[result] = (
        f'=IF(OR({first}="",{operation}="",{second}=""),"",'
        f'IF({operation}="+",{first}+{second},'
        f'IF({operation}="-",{first}-{second},'
        f'IF({operation}="*",{first}*{second},'
        f'IF({operation}="/",'
        f'IF({second}=0,"Division by zero",{first}/{second}),'
        f'"Invalid operation")))))'
    )


# Calculates category-wise totals and overall claim amounts.
def add_summary_formulas(workbook):
    worksheet = workbook[SUMMARY_SHEET]
    expense_sheet = workbook[EXPENSE_SHEET]

    category_column = EXPENSE_COLUMNS["category"]
    amount_column = EXPENSE_COLUMNS["amount"]

    category_letter = expense_sheet.cell(
        row=EXPENSE_FIRST_ROW,
        column=category_column,
    ).column_letter

    amount_letter = expense_sheet.cell(
        row=EXPENSE_FIRST_ROW,
        column=amount_column,
    ).column_letter

    # SUMIF groups expense amounts by their selected category.
    for row, category in zip(
        range(SUMMARY_CATEGORY_START_ROW, SUMMARY_CATEGORY_END_ROW + 1),
        EXPENSE_CATEGORIES,
    ):
        worksheet.cell(row=row, column=1).value = category

        worksheet.cell(row=row, column=2).value = (
            f'=SUMIF(\'{EXPENSE_SHEET}\'!'
            f'${category_letter}${EXPENSE_FIRST_ROW}:'
            f'${category_letter}${EXPENSE_LAST_ROW},'
            f'A{row},'
            f'\'{EXPENSE_SHEET}\'!'
            f'${amount_letter}${EXPENSE_FIRST_ROW}:'
            f'${amount_letter}${EXPENSE_LAST_ROW})'
        )

    # Links the total claimed amount to the expense worksheet.
    worksheet[TOTAL_CLAIMED_CELL] = (
        f"='{EXPENSE_SHEET}'!{EXPENSE_TOTAL_CELL}"
    )

    # Subtracts any advance received to calculate the net payable amount.
    worksheet[NET_PAYABLE_CELL] = (
        f'=IF({TOTAL_CLAIMED_CELL}="","",'
        f'{TOTAL_CLAIMED_CELL}-N({ADVANCE_RECEIVED_CELL}))'
    )

    # Links the employee name to the signature field.
    employee_name = EMPLOYEE_FIELDS["employee_name"]["cell"]

    worksheet[EMPLOYEE_SIGNATURE_CELL] = (
        f'=IF(\'{EMPLOYEE_SHEET}\'!{employee_name}="","",'
        f'\'{EMPLOYEE_SHEET}\'!{employee_name})'
    )


# Runs all formula functions and configures Excel to recalculate on opening.
def add_workbook_formulas(workbook):
    add_employee_formulas(workbook)
    add_expense_formulas(workbook)
    add_calculator_formulas(workbook)
    add_summary_formulas(workbook)

    workbook.calculation = CalcProperties(
        calcMode="auto",
        fullCalcOnLoad=True,
        forceFullCalc=True,
    )