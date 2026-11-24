
from mcp.server import MCPServer

from tools.create_excel import create_excel as create_excel_workbook


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


if __name__ == "__main__":
    # Runining using stdio transport by default
    mcp.run()
