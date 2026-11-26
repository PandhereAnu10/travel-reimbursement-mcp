import importlib
import sys
import types
from types import SimpleNamespace

import pytest


@pytest.fixture

def client_module(monkeypatch):
    mcp_stub = types.ModuleType("mcp")
    mcp_stub.Client = object
    mcp_stub.StdioServerParameters = lambda **kwargs: kwargs
    litellm_stub = types.ModuleType("litellm")
    litellm_stub.acompletion = None
    monkeypatch.setitem(sys.modules, "mcp", mcp_stub)
    monkeypatch.setitem(sys.modules, "litellm", litellm_stub)
    sys.modules.pop("client", None)
    module = importlib.import_module("client")
    yield module
    sys.modules.pop("client", None)


def make_tool(name, properties):
    return SimpleNamespace(
        name=name,
        input_schema={"properties": {key: {} for key in properties}, "required": []},
    )


def test_expense_creation_requires_user_supplied_expense_id(client_module):
    tool = make_tool(
        "write_excel",
        {
            "filename", "expense_id", "expense_date", "category",
            "description", "amount", "receipt_reference", "employee_info",
            "declaration",
        },
    )
    arguments = {
        "filename": "claim.xlsx",
        "expense_id": "EXP001",
        "expense_date": "2026-11-25",
        "category": "Transport",
        "description": "Airport taxi",
        "amount": 950,
    }
    valid, message = client_module.validate_tool_arguments(
        tool,
        arguments,
        "Please add expense EXP001 dated 2026-11-25 for airport taxi amount 950 to claim.xlsx",
    )
    assert valid, message

    arguments["expense_id"] = "EXP999"
    valid, message = client_module.validate_tool_arguments(
        tool,
        arguments,
        "Please add expense EXP001 dated 2026-11-25 for airport taxi amount 950 to claim.xlsx",
    )
    assert not valid
    assert "never invent" in message


def test_update_expense_id_none_does_not_raise(client_module):
    tool = make_tool("update_excel", {"filename", "expense_id", "updates", "employee_updates", "declaration"})
    valid, message = client_module.validate_tool_arguments(
        tool,
        {"filename": "claim.xlsx", "expense_id": None, "updates": {"amount": 10}},
        "Update expense amount",
    )
    assert not valid
    assert "expense_id is required" in message


def test_employee_values_must_be_grounded(client_module):
    tool = make_tool("write_excel", {"filename", "employee_info", "declaration"})
    valid, message = client_module.validate_tool_arguments(
        tool,
        {"filename": "claim.xlsx", "employee_info": {
            "employee_name": "John Doe", "employee_id": "E001",
            "department": "Engineering", "designation": "Engineer",
            "email": "john@example.com", "travel_purpose": "Meeting",
            "source_location": "Mumbai", "destination_location": "Pune",
            "travel_start_date": "2026-11-20", "travel_end_date": "2026-11-22",
        }},
        "Please record my employee information.",
    )
    assert not valid
    assert "do not invent it" in message


def test_non_finite_amount_is_rejected(client_module):
    tool = make_tool(
        "write_excel",
        {"filename", "expense_id", "expense_date", "category", "description", "amount", "receipt_reference", "employee_info", "declaration"},
    )
    valid, message = client_module.validate_tool_arguments(
        tool,
        {"filename": "claim.xlsx", "expense_id": "EXP001", "expense_date": "2026-11-25", "category": "Food", "description": "Lunch", "amount": float("nan")},
        "Add expense EXP001 dated 2026-11-25 Food Lunch amount nan claim.xlsx",
    )
    assert not valid
    assert "finite, non-negative" in message


def test_unmatched_think_tag_does_not_erase_response(client_module):
    assert client_module.clean_assistant_response("Answer: <think>unfinished") == "Answer: <think>unfinished"


def test_missing_result_content_is_handled(client_module):
    assert client_module.get_tool_result_text(SimpleNamespace(content=None)) == ""


def test_detects_clarification_after_a_successful_lookup(client_module):
    assert client_module.is_clarification_request(
        "I found two workbooks. Which workbook should I use?"
    )


def test_does_not_keep_completed_task_open_for_generic_offer(client_module):
    assert not client_module.is_clarification_request(
        "The expense was added successfully. Would you like anything else?"
    )


def test_employee_write_requires_all_mandatory_fields(client_module):
    tool = make_tool("write_excel", {"filename", "employee_info", "declaration", "save_as_draft"})
    valid, message = client_module.validate_tool_arguments(
        tool,
        {"filename": "claim.xlsx", "employee_info": {"employee_name": "Anushka"}},
        "Add employee information for Anushka.",
    )
    assert not valid
    assert "travel_purpose" in message
    assert "travel_start_date" in message


def test_explicit_draft_allows_partial_employee_fields(client_module):
    tool = make_tool("write_excel", {"filename", "employee_info", "declaration", "save_as_draft"})
    valid, message = client_module.validate_tool_arguments(
        tool,
        {"filename": "claim.xlsx", "employee_info": {"employee_name": "Anushka"}, "save_as_draft": True},
        "Save a draft employee record for Anushka.",
    )
    assert valid, message


def test_verified_employee_write_summary_uses_only_confirmed_fields(client_module):
    summary = client_module.format_verified_employee_write({
        "status": "success",
        "filename": "Manual_Test_Claim.xlsx",
        "updated_fields": ["employee_name", "employee_id", "email"],
        "record_status": "Complete",
        "missing_fields": [],
    })
    assert "Employee name" in summary
    assert "Employee ID" in summary
    assert "Email" in summary
    assert "Travel purpose" not in summary
    assert "Record status: complete." in summary


def test_verified_employee_draft_summary_reports_missing_fields(client_module):
    summary = client_module.format_verified_employee_write({
        "status": "success",
        "filename": "Manual_Test_Claim.xlsx",
        "updated_fields": ["employee_name"],
        "record_status": "Draft - incomplete",
        "missing_fields": ["travel_purpose", "travel_start_date"],
    })
    assert "incomplete draft" in summary
    assert "Travel purpose" in summary
    assert "Travel start date" in summary


def test_currency_conversion_accepts_equivalent_decimal_formatting(client_module):
    tool = make_tool(
        "calculate_expense",
        {"filename", "calculation_type", "foreign_amount", "exchange_rate", "currency_code"},
    )
    valid, message = client_module.validate_tool_arguments(
        tool,
        {
            "filename": "claim.xlsx",
            "calculation_type": "currency_conversion",
            "foreign_amount": 100,
            "exchange_rate": 83.5,
            "currency_code": "USD",
        },
        "Convert USD 100 into INR using an exchange rate of 83.50 INR per USD "
        "and save the calculation in claim.xlsx.",
    )
    assert valid, message


def test_clarification_detector_catches_update_confirmation_question(client_module):
    assert client_module.is_clarification_request(
        "Would you like me to update that existing record with the new details?"
    )


def test_clarification_detector_catches_yes_no_action_question(client_module):
    assert client_module.is_clarification_request(
        "Do you want me to replace the existing expense?"
    )


def test_advance_amount_is_forwarded_only_when_user_provided(client_module):
    tool = make_tool(
        "update_excel",
        {"filename", "expense_id", "updates", "employee_updates", "declaration", "advance_received"},
    )
    valid, message = client_module.validate_tool_arguments(
        tool,
        {"filename": "claim.xlsx", "advance_received": 1000},
        "I received an advance of INR 1,000 for claim.xlsx.",
    )
    assert valid, message

    valid, message = client_module.validate_tool_arguments(
        tool,
        {"filename": "claim.xlsx", "advance_received": 1000},
        "Update the advance amount in claim.xlsx.",
    )
    assert not valid
    assert "not present in the user's messages" in message
