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



def test_server_uses_streamable_http_transport(monkeypatch):
    module, _ = load_server_with_mcp_stub(monkeypatch)
    source = Path(module.__file__).read_text(encoding="utf-8")
    tree = ast.parse(source)

    main_guard = next(
        node for node in tree.body
        if isinstance(node, ast.If)
        and isinstance(node.test, ast.Compare)
        and isinstance(node.test.left, ast.Name)
        and node.test.left.id == "__name__"
    )

    run_call = next(
        node.value
        for node in main_guard.body
        if isinstance(node, ast.Expr)
        and isinstance(node.value, ast.Call)
        and isinstance(node.value.func, ast.Attribute)
        and node.value.func.attr == "run"
    )

    assert run_call.func.attr == "run"

    keywords = {
        keyword.arg: keyword.value
        for keyword in run_call.keywords
    }

    assert ast.literal_eval(keywords["transport"]) == "streamable-http"
    assert ast.literal_eval(keywords["streamable_http_path"]) == "/mcp"
