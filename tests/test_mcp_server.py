import ast
import importlib
import sys
import types
from pathlib import Path


def load_server_with_mcp_stub(monkeypatch):
    registered = {}

    class FakeMCPServer:
        def __init__(self, name):
            self.name = name

        def tool(self, name=None, description=None):
            def decorator(function):
                registered[name or function.__name__] = {
                    "function": function,
                    "description": description or function.__doc__ or "",
                }
                return function
            return decorator

        def run(self, *args, **kwargs):
            return None

    mcp_package = types.ModuleType("mcp")
    mcp_package.__path__ = []
    mcp_server_module = types.ModuleType("mcp.server")
    mcp_server_module.MCPServer = FakeMCPServer
    monkeypatch.setitem(sys.modules, "mcp", mcp_package)
    monkeypatch.setitem(sys.modules, "mcp.server", mcp_server_module)
    sys.modules.pop("server", None)
    module = importlib.import_module("server")
    return module, registered


def test_all_expected_tools_are_registered(monkeypatch):
    _, registered = load_server_with_mcp_stub(monkeypatch)

    assert set(registered) == {
        "ping",
        "create_excel",
        "read_excel",
        "list_excel_files",
        "write_excel",
        "update_excel",
        "submit_to_hr",
        "calculate_expense",
    }
    assert all(item["description"] for item in registered.values())


def test_write_and_update_descriptions_explain_supported_modes(monkeypatch):
    _, registered = load_server_with_mcp_stub(monkeypatch)

    assert "employee_info" in registered["write_excel"]["description"]
    assert "employee_updates" in registered["update_excel"]["description"]
    assert "travel_days" in registered["write_excel"]["description"]


def test_server_uses_authenticated_streamable_http_transport():
    source = Path(__file__).resolve().parents[1].joinpath("server.py").read_text(encoding="utf-8")

    assert "MCP_AUTH_TOKEN" in source
    assert 'mcp.streamable_http_app(streamable_http_path="/mcp")' in source
    assert "BearerTokenMiddleware" in source
    assert "uvicorn.run(app, host=host, port=port)" in source


def test_server_supports_stdio_and_http_transports():
    source = Path(__file__).resolve().parents[1].joinpath("server.py").read_text(encoding="utf-8")
    assert 'choices=("stdio", "streamable-http")' in source
    assert 'mcp.run(transport="stdio")' in source
    assert 'mcp.streamable_http_app(streamable_http_path="/mcp")' in source
