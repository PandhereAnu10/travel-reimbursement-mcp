from mcp.server import MCPServer

mcp = MCPServer('travel-reimbursement')


@mcp.tool()
def ping() -> str:
    """Check that the server is running."""
    return 'pong'


if __name__ == '__main__':
    mcp.run()  # stdio transport by default