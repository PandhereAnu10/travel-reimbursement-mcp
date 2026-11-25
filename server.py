
from mcp.server import MCPServer

from tools.create_excel import create_excel as create_excel_workbook
from tools.read_excel import read_excel as read_excel_workbook
from tools.list_excel_files import list_excel_files
from tools.write_excel import write_excel
from tools.update_excel import update_excel
from tools.submit_to_hr import submit_to_hr
from tools.calculate_expense import calculate_expense


# Initialize the MCP server
mcp = MCPServer("travel-reimbursement")


# Health-check tool
@mcp.tool()
def ping() -> str:
    # checking server is running
    return "pong"


# Excel workbook creation tool
@mcp.tool(
    name="create_excel",
    description=(
        "Create a new travel reimbursement Excel workbook "
        "with the four required worksheets. "
        "Accepts an Excel filename."
    ),
)
def create_excel_tool(filename: str) -> dict:
    return create_excel_workbook(filename)


# Excel workbook reading tool
@mcp.tool(
    name="read_excel",
    description=(
        "Read data from an existing travel reimbursement "
        "Excel workbook. Optionally read a specific worksheet."
    ),
)
def read_excel_tool(
    filename: str,
    sheet_name: str = None,
) -> dict:
    """Read workbook data using the Excel reading module."""
    return read_excel_workbook(filename, sheet_name)


# Excel workbook listing tool
@mcp.tool(
    name="list_excel_files",
    description=(
        "List all available Excel reimbursement workbooks "
        "with their filenames, sizes, modification dates, "
        "and worksheet names. Reports unreadable workbooks separately."
    ),
)
def list_excel_files_tool() -> dict:
    return list_excel_files()


# Excel workbook writing tool
@mcp.tool(
    name="write_excel",
    description=(
        "Add a new expense record to an existing reimbursement workbook. "
        "Requires filename, unique expense ID, expense date, category, "
        "description, and amount. Receipt reference is optional. "
        "Does not overwrite existing expense records."
    ),
)
def write_excel_tool(
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
    return write_excel(
        filename=filename,
        expense_id=expense_id,
        expense_date=expense_date,
        category=category,
        description=description,
        amount=amount,
        receipt_reference=receipt_reference,
        employee_info=employee_info,
        declaration=declaration,
    )

# Excel workbook update tool
@mcp.tool(
    name="update_excel",
    description=(
        "Update selected fields of an existing expense record "
        "using its unique Expense ID. Supported fields are date, "
        "category, description, amount, and receipt_reference. "
        "Only explicitly supplied fields are changed."
    ),
)
def update_excel_tool(
    filename: str,
    expense_id: str = "",
    updates: dict = None,
    employee_updates: dict = None,
    declaration: str = None,
) -> dict:
    return update_excel(
        filename=filename,
        expense_id=expense_id,
        updates=updates,
        employee_updates=employee_updates,
        declaration=declaration,
    )


# Submit reimbursement claim to HR
@mcp.tool(
    name="submit_to_hr",
    description=(
        "Validate and submit a travel reimbursement claim to HR. "
        "Checks required employee and travel information, at least one "
        "expense record, and employee declaration before submission. "
        "On success, updates the claim status, submission date, "
        "and unique claim reference in the workbook."
    ),
)
def submit_to_hr_tool(filename: str) -> dict:
    return submit_to_hr(filename=filename)

@mcp.tool(
    name="calculate_expense",
    description=(
        "Calculate travel expenses using mileage, per diem, "
        "currency conversion, or basic arithmetic operations. "
        "Saves calculation inputs and results in the workbook's "
        "Expense Calculator worksheet. Supports addition, "
        "subtraction, multiplication, and division."
    ),
)
def calculate_expense_tool(
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
    return calculate_expense(
        filename=filename,
        calculation_type=calculation_type,
        distance_km=distance_km,
        rate_per_km=rate_per_km,
        travel_days=travel_days,
        daily_allowance=daily_allowance,
        foreign_amount=foreign_amount,
        exchange_rate=exchange_rate,
        currency_code=currency_code,
        number1=number1,
        number2=number2,
        operation=operation,
    )


if __name__ == "__main__":
    # Runining using stdio transport by default
    mcp.run()
