
from mcp.server import MCPServer

from tools.create_excel import create_excel as create_excel_workbook
from tools.read_excel import read_excel as read_excel_workbook
from tools.list_excel_files import list_excel_files
from tools.write_excel import write_excel
from tools.update_excel import update_excel


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
    expense_id: str,
    expense_date: str,
    category: str,
    description: str,
    amount: float,
    receipt_reference: str = "",
) -> dict:
    return write_excel(
        filename=filename,
        expense_id=expense_id,
        expense_date=expense_date,
        category=category,
        description=description,
        amount=amount,
        receipt_reference=receipt_reference,
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
    expense_id: str,
    updates: dict,
) -> dict:
    """Update an existing expense record."""
    return update_excel(
        filename=filename,
        expense_id=expense_id,
        updates=updates,
    )


if __name__ == "__main__":
    # Runining using stdio transport by default
    mcp.run()
