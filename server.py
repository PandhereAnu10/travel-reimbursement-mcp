
import os
import argparse
from pathlib import Path

from dotenv import load_dotenv
from mcp.server import MCPServer

from security import BearerTokenMiddleware
from mcp.server.transport_security import TransportSecuritySettings

from tools.create_excel import create_excel as create_excel_workbook
from tools.read_excel import read_excel as read_excel_workbook
from tools.list_excel_files import list_excel_files
from tools.write_excel import write_excel
from tools.update_excel import update_excel
from tools.submit_to_hr import submit_to_hr
from tools.calculate_expense import calculate_expense


# Load project configuration when the server is started in its own terminal.
PROJECT_DIR = Path(__file__).resolve().parent
load_dotenv(PROJECT_DIR / ".env")

# Initialize the MCP server
mcp = MCPServer("travel-reimbursement")


# Health-check tool
@mcp.tool(
    description="Check whether the travel reimbursement MCP server is responding."
)
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
        "Read data from an existing travel reimbursement Excel workbook. "
        "If the user names a worksheet, ALWAYS pass its exact name as sheet_name; "
        "otherwise all worksheets are returned. Formula expressions are omitted by default "
        "to keep results concise; set include_formulas=true only when the user asks to inspect formulas."
    ),
)
def read_excel_tool(
    filename: str,
    sheet_name: str = None,
    include_formulas: bool = False,
) -> dict:
    """Read workbook data using the Excel reading module."""
    return read_excel_workbook(filename, sheet_name, include_formulas)


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
        "Write to an existing reimbursement workbook. Use employee_info to add "
        "employee/travel fields, or declaration to record Yes/No; these may be "
        "combined. Otherwise add one NEW expense using filename, unique expense_id, "
        "expense_date (YYYY-MM-DD), category, description, and amount; "
        "receipt_reference is optional. Do not combine employee/declaration data "
        "with expense creation. Does not overwrite existing expenses. "
        "A complete travel start and end date must already be recorded before an expense can be added. "
        "travel_days is calculated from the travel dates and cannot be set directly. "
        "For a normal employee/travel record, provide ALL required fields: employee_name, "
        "employee_id, department, designation, email, travel_purpose, source_location, "
        "destination_location, travel_start_date, and travel_end_date. The tool rejects "
        "incomplete records unless the user explicitly asks to save a draft and save_as_draft=true is passed."
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
    save_as_draft: bool = False,
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
        save_as_draft=save_as_draft,
    )

# Excel workbook update tool
@mcp.tool(
    name="update_excel",
    description=(
        "Update an existing reimbursement workbook. For an expense, provide "
        "expense_id and updates containing only date, category, description, "
        "amount, and/or receipt_reference. An explicit expense-date change requires "
        "a complete travel start/end period in the workbook. For employee/travel information, use "
        "employee_updates; declaration can be updated separately or alongside "
        "those fields. Use advance_received to set the advance amount on the Claim Summary & HR sheet. "
        "Never pass travel_days: it is calculated from travel dates. "
        "Only explicitly supplied fields are changed."
    ),
)
def update_excel_tool(
    filename: str,
    expense_id: str = "",
    updates: dict = None,
    employee_updates: dict = None,
    declaration: str = None,
    advance_received: float = None,
) -> dict:
    return update_excel(
        filename=filename,
        expense_id=expense_id,
        updates=updates,
        employee_updates=employee_updates,
        declaration=declaration,
        advance_received=advance_received,
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
        "Saves calculation inputs in the workbook's Expense Calculator worksheet "
        "and returns the calculated result. Workbook result cells remain formulas "
        "so Excel can recalculate them. Supports addition, subtraction, "
        "multiplication, and division."
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
    parser = argparse.ArgumentParser(description="Travel reimbursement MCP server")
    parser.add_argument(
        "--transport",
        choices=("stdio", "streamable-http"),
        default=os.getenv("MCP_TRANSPORT", "streamable-http"),
        help="MCP transport to expose (default: streamable-http)",
    )
    args = parser.parse_args()

    if args.transport == "stdio":
        # Local stdio mode is intended for MCP clients that launch this process.
        mcp.run(transport="stdio")
    else:
        # Require a configured secret so the HTTP endpoint never starts unprotected.
        auth_token = os.getenv("MCP_AUTH_TOKEN")
        if not auth_token or len(auth_token) < 32:
            raise RuntimeError(
                "MCP_AUTH_TOKEN is missing or too short. Generate one with "
                "python -c \"import secrets; print(secrets.token_urlsafe(32))\" and add it to .env."
            )

        app = mcp.streamable_http_app(
            streamable_http_path="/mcp",
            transport_security=TransportSecuritySettings(
                allowed_hosts=[
                    "127.0.0.1:*",
                    "localhost:*",
                    "ghostlier-adrianna-unsuccinctly.ngrok-free.dev",
                    "ghostlier-adrianna-unsuccinctly.ngrok-free.dev:*",
                ],
                allowed_origins=[
                    "https://ghostlier-adrianna-unsuccinctly.ngrok-free.dev",
                ],
            ),
        )
        app = BearerTokenMiddleware(app, auth_token, protected_path="/mcp")

        host = os.getenv("MCP_HOST", "127.0.0.1")
        port = int(os.getenv("MCP_PORT", "8000"))

        import uvicorn
        uvicorn.run(app, host=host, port=port)
