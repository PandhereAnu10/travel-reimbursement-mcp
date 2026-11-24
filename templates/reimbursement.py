
# Workbook sheet names

EMPLOYEE_SHEET = "Employee & Travel Information"
EXPENSE_SHEET = "Expense Details"
CALCULATOR_SHEET = "Expense Calculator"
SUMMARY_SHEET = "Claim Summary & HR"

# 1. Employee & Travel Info

EMPLOYEE_TITLE_CELL = "A1"

EMPLOYEE_START_ROW = 3
EMPLOYEE_END_ROW = 13

EMPLOYEE_LABEL_COLUMN = 1
EMPLOYEE_VALUE_COLUMN = 2

EMPLOYEE_FIELDS = {
    "employee_name": {
        "label": "Employee Name",
        "cell": "B3"
    },
    "employee_id": {
        "label": "Employee ID",
        "cell": "B4"
    },
    "department": {
        "label": "Department",
        "cell": "B5"
    },
    "designation": {
        "label": "Designation",
        "cell": "B6"
    },
    "email": {
        "label": "Email",
        "cell": "B7"
    },
    "travel_purpose": {
        "label": "Travel Purpose",
        "cell": "B8"
    },
    "source_location": {
        "label": "Source Location",
        "cell": "B9"
    },
    "destination_location": {
        "label": "Destination Location",
        "cell": "B10"
    },
    "travel_start_date": {
        "label": "Travel Start Date",
        "cell": "B11"
    },
    "travel_end_date": {
        "label": "Travel End Date",
        "cell": "B12"
    },
    "travel_days": {
        "label": "Number of Travel Days",
        "cell": "B13"
    }
}



# 2. Expense Details

EXPENSE_TITLE_CELL = "A1"
EXPENSE_TOTAL_CELL = "E2"

EXPENSE_HEADER_ROW = 4
EXPENSE_FIRST_ROW = 5
EXPENSE_LAST_ROW = 204

EXPENSE_CATEGORIES = [
    "Transport",
    "Food",
    "Accommodation",
    "Local Conveyance",
    "Miscellaneous"
]

EXPENSE_COLUMNS = {
    "expense_id": 1,
    "date": 2,
    "category": 3,
    "description": 4,
    "amount": 5,
    "receipt_reference": 6,
    "date_check": 7
}

EXPENSE_HEADERS = [
    "Expense ID",
    "Date",
    "Category",
    "Description",
    "Amount (INR)",
    "Receipt Reference",
    "Date Check"
]



# 3. Expense Calculator

CALCULATOR_TITLE_CELL = "A1"

# Mileage Calculator
MILEAGE_DISTANCE_CELL = "B4"
MILEAGE_RATE_CELL = "B5"
MILEAGE_RESULT_CELL = "C6"

# Per Diem Calculator
PER_DIEM_DAYS_CELL = "B9"
PER_DIEM_ALLOWANCE_CELL = "B10"
PER_DIEM_RESULT_CELL = "C11"

# Currency Conversion
CURRENCY_AMOUNT_CELL = "B14"
CURRENCY_CODE_CELL = "B15"
CURRENCY_RATE_CELL = "B16"
CURRENCY_RESULT_CELL = "C17"

# Quick Calculator
QUICK_CALC_FIRST_VALUE_CELL = "B20"
QUICK_CALC_OPERATION_CELL = "B21"
QUICK_CALC_SECOND_VALUE_CELL = "B22"
QUICK_CALC_RESULT_CELL = "C23"

QUICK_CALC_OPERATIONS = [
    "+",
    "-",
    "*",
    "/"
]



# 4. Claim Summary & HR Submission

SUMMARY_TITLE_CELL = "A1"

# Category-wise expense summary
SUMMARY_CATEGORY_START_ROW = 4
SUMMARY_CATEGORY_END_ROW = 8

SUMMARY_CATEGORY_AMOUNT_COLUMN = 2

# Overall claim calculations
TOTAL_CLAIMED_CELL = "B9"
ADVANCE_RECEIVED_CELL = "B10"
NET_PAYABLE_CELL = "B11"

# Employee declaration and submission
DECLARATION_CELL = "B13"
EMPLOYEE_SIGNATURE_CELL = "B14"
SUBMISSION_STATUS_CELL = "B15"
SUBMISSION_DATE_CELL = "B16"
CLAIM_REFERENCE_ID_CELL = "B17"
HR_NOTES_CELL = "B18"

DECLARATION_OPTIONS = [
    "Yes",
    "No"
]

SUBMISSION_STATUS_OPTIONS = [
    "Draft",
    "Submitted to HR"
]



# Workbook default and formatting

# Claim submission status
DEFAULT_STATUS = "Draft"
SUBMITTED_STATUS = "Submitted to HR"

# Date and currency display formats
DATE_FORMAT = "DD-MMM-YYYY"
CURRENCY_FORMAT = '₹#,##,##0.00'

# Workbook styling colours (ARGB without the #)
INPUT_FILL_COLOR = "FFFFF2CC"   # Light yellow
HEADER_FILL_COLOR = "FF1F4E78"  # Dark blue
OUTPUT_FILL_COLOR = "FFE2F0D9"
BORDER_COLOR = "FFD9E1F2"

# Expense date validation
OUTSIDE_TRAVEL_DATES_FLAG = "Outside travel dates"
WITHIN_TRAVEL_DATES_FLAG = "Within travel dates"
MISSING_TRAVEL_DATES_FLAG = "Travel dates not provided"


# REquired Employee Fields for submitting to hr

REQUIRED_EMPLOYEE_FIELDS = (
    "employee_name",
    "employee_id",
    "department",
    "designation",
    "email",
    "travel_purpose",
    "source_location",
    "destination_location",
    "travel_start_date",
    "travel_end_date",
)
