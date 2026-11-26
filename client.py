
import asyncio
from contextlib import asynccontextmanager
from pathlib import Path
import json
import os
import re
import sys
import math
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

from templates.reimbursement import REQUIRED_EMPLOYEE_FIELDS

from mcp import Client
from dotenv import load_dotenv
from litellm import acompletion

# project dir
PROJECT_DIR = Path(__file__).resolve().parent
load_dotenv(PROJECT_DIR / ".env")

LLM_BASE_URL = os.getenv("LLM_BASE_URL")
LLM_API_KEY = os.getenv("LLM_API_KEY")
LLM_MODEL = os.getenv("LLM_MODEL")

# MCP Streamable HTTP server endpoint. Start server.py in a separate terminal.
MCP_SERVER_URL = os.getenv("MCP_SERVER_URL", "http://127.0.0.1:8000/mcp")
MCP_AUTH_TOKEN = os.getenv("MCP_AUTH_TOKEN")
MCP_TRANSPORT = os.getenv("MCP_TRANSPORT", "streamable-http").strip().lower()


@asynccontextmanager
async def authenticated_mcp_client():
    if MCP_TRANSPORT == "stdio":
        from mcp.client.stdio import stdio_client, StdioServerParameters

        server_parameters = StdioServerParameters(
            command=sys.executable,
            args=[str(PROJECT_DIR / "server.py"), "--transport", "stdio"],
            env=os.environ.copy(),
        )
        async with stdio_client(server_parameters) as transport:
            async with Client(transport) as mcp_client:
                yield mcp_client
        return

    if MCP_TRANSPORT != "streamable-http":
        raise ValueError("MCP_TRANSPORT must be 'stdio' or 'streamable-http'.")

    import httpx
    from mcp.client.streamable_http import streamable_http_client

    if not MCP_AUTH_TOKEN or len(MCP_AUTH_TOKEN) < 32:
        raise ValueError(
            "MCP_AUTH_TOKEN is missing or too short. Add the same generated token "
            "(at least 32 characters) to the server and client .env configuration."
        )

    async with httpx.AsyncClient(
        headers={
            "Authorization": f"Bearer {MCP_AUTH_TOKEN}",
            "ngrok-skip-browser-warning": "true",
        },
        timeout=httpx.Timeout(30.0, read=300.0),
    ) as http_client:
        transport = streamable_http_client(
            MCP_SERVER_URL,
            http_client=http_client,
        )
        async with Client(transport) as mcp_client:
            yield mcp_client


def convert_mcp_tools_to_llm_tools(mcp_tools):
    llm_tools = []

    for tool in mcp_tools:
        llm_tools.append({
            "type": "function",
            "function": {
                "name": tool.name,
                "description": tool.description or "",
                "parameters": tool.input_schema,
            },
        })

    return llm_tools


def get_tool_result_text(result):
    structured_content = getattr(
        result,
        "structuredContent",
        None
    )

    if structured_content is None:
        structured_content = getattr(
            result,
            "structured_content",
            None
        )

    if structured_content is not None:
        return structured_content

    content_parts = []

    for item in (getattr(result, "content", None) or []):
        if hasattr(item, "text"):
            content_parts.append(item.text)
        else:
            content_parts.append(str(item))

    text_result = "\n".join(content_parts)

    try:
        return json.loads(text_result)
    except json.JSONDecodeError:
        return text_result


def clean_assistant_response(content):
    if not content:
        return ""

    # Remove only complete thinking blocks. An unmatched opening tag should
    # not cause the entire user-facing answer to be discarded.
    content = re.sub(
        r"<think>.*?</think>",
        "",
        content,
        flags=re.DOTALL | re.IGNORECASE,
    )

    return content.strip()


def is_clarification_request(content):
    if not content:
        return False

    text = content.casefold()
    clarification_phrases = (
        "which workbook",
        "which file",
        "which expense",
        "which expense id",
        "what is the",
        "what's the",
        "what date",
        "what amount",
        "what is your",
        "please provide",
        "could you provide",
        "can you provide",
        "could you confirm",
        "please confirm",
        "please tell me",
        "could you tell me",
        "missing information",
        "missing field",
        "need the",
        "provide the",
        "would you like me to",
        "do you want me to",
        "would you prefer me to",
        "should i",
        "shall i",
    )
    return "?" in text and any(phrase in text for phrase in clarification_phrases)



def format_verified_employee_write(result):

    if not isinstance(result, dict) or result.get("status") != "success":
        return None

    labels = {
        "employee_name": "Employee name",
        "employee_id": "Employee ID",
        "department": "Department",
        "designation": "Designation",
        "email": "Email",
        "travel_purpose": "Travel purpose",
        "source_location": "Source location",
        "destination_location": "Destination location",
        "travel_start_date": "Travel start date",
        "travel_end_date": "Travel end date",
    }
    updated_fields = result.get("updated_fields") or []
    lines = [
        f"Employee/travel information was saved to {result.get('filename', 'the workbook')}."
    ]
    if updated_fields:
        lines.append("Confirmed saved fields:")
        lines.extend(f"- {labels.get(field, field)}" for field in updated_fields)
    else:
        lines.append("No employee fields were reported as updated.")

    record_status = result.get("record_status")
    if record_status == "Draft - incomplete":
        missing = result.get("missing_fields") or []
        lines.append("Record status: incomplete draft.")
        if missing:
            lines.append("Still missing: " + ", ".join(labels.get(field, field) for field in missing) + ".")
    elif record_status == "Complete":
        lines.append("Record status: complete.")

    if result.get("declaration") is not None:
        lines.append(f"Declaration recorded: {result['declaration']}.")
    return "\n".join(lines)

def validate_tool_arguments(tool, arguments, conversation_evidence=""):


    if not isinstance(arguments, dict):
        return False, "Tool arguments must be a JSON object."

    schema = tool.input_schema or {}
    properties = schema.get("properties", {})
    required = schema.get("required", [])

    # Check arguments required by the MCP schema.
    missing = [
        field for field in required
        if arguments.get(field) is None
        or (
            isinstance(arguments.get(field), str)
            and not arguments[field].strip()
        )
    ]

    if missing:
        return False, f"Missing required arguments: {', '.join(missing)}"

    # Reject arguments that are not part of the tool schema.
    unknown = set(arguments.keys()) - set(properties.keys())

    if unknown:
        return False, f"Unexpected arguments: {', '.join(sorted(unknown))}"

    def is_number(value):
        if not isinstance(value, (int, float)) or isinstance(value, bool):
            return False
        try:
            return math.isfinite(float(value)) and value >= 0
        except (OverflowError, TypeError, ValueError):
            return False

    def non_empty_dict(value):
        return isinstance(value, dict) and len(value) > 0

    def value_was_supplied(value):
        if not isinstance(value, (str, int, float)) or isinstance(value, bool):
            return False
        if isinstance(value, (int, float)):
            try:
                expected_number = Decimal(str(value))
            except (InvalidOperation, ValueError):
                return False
            if not expected_number.is_finite():
                return False

            # Treat numerically equivalent forms (83.5 and 83.50) as the
            # same user-supplied value, without allowing 83.5 to match 835.
            number_tokens = re.findall(
                r"(?<![a-z0-9])(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?(?![a-z0-9])",
                conversation_evidence.casefold(),
            )
            for token in number_tokens:
                try:
                    if Decimal(token.replace(",", "")) == expected_number:
                        return True
                except InvalidOperation:
                    continue
            return False

        normalized_value = re.sub(r"[^a-z0-9]+", " ", str(value).casefold()).strip()
        normalized_evidence = re.sub(
            r"[^a-z0-9]+", " ", conversation_evidence.casefold()
        )
        if not normalized_value:
            return False
        # Require token boundaries so IDs and words are not partial matches.
        return re.search(
            rf"(?<![a-z0-9]){re.escape(normalized_value)}(?![a-z0-9])",
            normalized_evidence,
        ) is not None

    def validate_employee_values(employee_data):
        if "travel_days" in employee_data:
            return False, (
                "travel_days is calculated from travel_start_date and "
                "travel_end_date; do not provide it as an editable field."
            )
        for field, value in employee_data.items():
            if not isinstance(value, str) or not value.strip():
                return False, f"{field} must be supplied as non-empty text."
            if not value_was_supplied(value):
                return False, (
                    f"The value for '{field}' is not present in the user's "
                    "messages or verified tool results. Ask the user to provide "
                    "or confirm it; do not invent it."
                )
            if field in {"travel_start_date", "travel_end_date"}:
                try:
                    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", value.strip()):
                        raise ValueError
                    date.fromisoformat(value.strip())
                except ValueError:
                    return False, f"{field} must use YYYY-MM-DD format."
            if field == "email" and not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", value.strip()):
                return False, "email must be a valid email address."
        return True, "Employee values are grounded in the conversation."

    employee_fields = {
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
        "travel_days",
    }

    expense_categories = {
        "Transport",
        "Food",
        "Accommodation",
        "Local Conveyance",
        "Miscellaneous",
    }

    declaration_options = {"Yes", "No"}

    # WRITE EXCEL
    if tool.name == "write_excel":
        employee_info = arguments.get("employee_info")
        declaration = arguments.get("declaration")
        save_as_draft = arguments.get("save_as_draft", False)

        # Employee information / declaration operation
        if employee_info is not None or declaration is not None:

            expense_fields = {
                "expense_id",
                "expense_date",
                "category",
                "description",
                "amount",
                "receipt_reference",
            }

            if any(arguments.get(field) is not None for field in expense_fields):
                return False, (
                    "Do not combine employee/declaration updates "
                    "with expense creation. Make separate tool calls."
                )

            if not isinstance(save_as_draft, bool):
                return False, "save_as_draft must be true or false."
            if save_as_draft and not value_was_supplied("draft"):
                return False, (
                    "Partial saving is allowed only when the user explicitly asks to save a draft. "
                    "Ask the user to confirm a draft save instead of setting save_as_draft=true."
                )

            if employee_info is not None:
                if not non_empty_dict(employee_info):
                    return False, "employee_info must be a non-empty object."

                missing_fields = [
                    field for field in REQUIRED_EMPLOYEE_FIELDS
                    if field not in employee_info
                    or employee_info[field] is None
                    or (isinstance(employee_info[field], str) and not employee_info[field].strip())
                ]
                if missing_fields and not save_as_draft:
                    return False, (
                        "Employee/travel record is incomplete. Missing required fields: "
                        + ", ".join(missing_fields)
                        + ". Recover values already supplied in the conversation; if any are genuinely missing, ask the user. "
                        "Only pass save_as_draft=true if the user explicitly requested a draft."
                    )

                valid_values, grounding_message = validate_employee_values(employee_info)
                if not valid_values:
                    return False, grounding_message

                invalid_fields = set(employee_info.keys()) - employee_fields

                if invalid_fields:
                    return False, (
                        f"Unsupported employee fields: "
                        f"{', '.join(sorted(invalid_fields))}"
                    )

            if declaration is not None and declaration not in declaration_options:
                return False, "declaration must be either 'Yes' or 'No'."

            return True, "Valid employee/declaration arguments."

        # Expense creation operation
        expense_required = [
            "filename",
            "expense_id",
            "expense_date",
            "category",
            "description",
            "amount",
        ]

        missing = [
            field for field in expense_required
            if arguments.get(field) is None
            or (
                isinstance(arguments.get(field), str)
                and not arguments[field].strip()
            )
        ]

        if missing:
            return False, (
                f"Expense creation requires: {', '.join(missing)}"
            )
            
         # Normalize expense dates before provenance validation.
        # The LLM may convert YYYY-MM-DD into an ISO datetime.
        expense_date = arguments.get("expense_date")

        if not isinstance(expense_date, str):
            return False, "expense_date must be a valid date in YYYY-MM-DD format."

        expense_date = expense_date.strip()

        try:
            if re.fullmatch(r"\d{4}-\d{2}-\d{2}", expense_date):
                expense_date = date.fromisoformat(expense_date).isoformat()
            else:
                expense_date = datetime.fromisoformat(
                    expense_date.replace("Z", "+00:00")
                ).date().isoformat()

        except ValueError:
            return False, (
                "Invalid expense date. Use a valid date in YYYY-MM-DD format."
            )

        # Pass the normalized date to the MCP tool.
        arguments["expense_date"] = expense_date


        if arguments["category"] not in expense_categories:
            return False, (
                f"Invalid category. Allowed categories: "
                f"{', '.join(sorted(expense_categories))}"
            )

        if not is_number(arguments["amount"]):
            return False, "amount must be a finite, non-negative number."
        if not value_was_supplied(arguments["amount"]):
            return False, "The expense amount is not present in the user's messages. Ask the user to confirm it."

        if not value_was_supplied(arguments["expense_id"]):
            return False, (
                "The Expense ID was not provided by the user or returned by a "
                "verified tool. Ask the user for the exact Expense ID; never invent one."
            )
        for field in ("filename", "expense_date", "description"):
            if not value_was_supplied(arguments[field]):
                return False, f"The value for '{field}' is not present in the user's messages or verified tool results. Ask the user to confirm it instead of guessing."

    # UPDATE EXCEL
    elif tool.name == "update_excel":
        updates = arguments.get("updates")
        employee_updates = arguments.get("employee_updates")
        declaration = arguments.get("declaration")
        advance_received = arguments.get("advance_received")

        if (updates is None and employee_updates is None and declaration is None
                and advance_received is None):
            return False, (
                "Provide at least one of updates, employee_updates, "
                "or declaration."
            )

        if updates is not None:
            if not non_empty_dict(updates):
                return False, "updates must be a non-empty object."

            expense_id = arguments.get("expense_id")
            if not isinstance(expense_id, str) or not expense_id.strip():
                return False, "expense_id is required when updating an expense."
            if not value_was_supplied(expense_id):
                return False, (
                    "The Expense ID was not provided by the user or returned by a "
                    "verified tool. Ask the user for the exact Expense ID; never invent one."
                )

            allowed_updates = {
                "date",
                "category",
                "description",
                "amount",
                "receipt_reference",
            }

            invalid_fields = set(updates.keys()) - allowed_updates

            if invalid_fields:
                return False, (
                    f"Unsupported expense update fields: "
                    f"{', '.join(sorted(invalid_fields))}"
                )

            if "amount" in updates and not is_number(updates["amount"]):
                return False, "Updated amount must be a finite, non-negative number."
            for field, value in updates.items():
                if field == "receipt_reference" and value in (None, ""):
                    continue
                if not value_was_supplied(value):
                    return False, f"The updated value for '{field}' is not present in the user's messages or verified tool results. Ask the user to confirm it."

            if (
                "category" in updates
                and updates["category"] not in expense_categories
            ):
                return False, "Invalid expense category."

        if employee_updates is not None:
            if not non_empty_dict(employee_updates):
                return False, "employee_updates must be a non-empty object."

            valid_values, grounding_message = validate_employee_values(employee_updates)
            if not valid_values:
                return False, grounding_message

            invalid_fields = set(employee_updates.keys()) - employee_fields

            if invalid_fields:
                return False, (
                    f"Unsupported employee update fields: "
                    f"{', '.join(sorted(invalid_fields))}"
                )

        if declaration is not None and declaration not in declaration_options:
            return False, "declaration must be either 'Yes' or 'No'."

        if advance_received is not None:
            if not is_number(advance_received):
                return False, "advance_received must be a finite, non-negative number."
            if not value_was_supplied(advance_received):
                return False, "The advance amount is not present in the user's messages. Ask the user to confirm it."

    # CALCULATE EXPENSE
    elif tool.name == "calculate_expense":
        calculation_type = arguments.get("calculation_type")

        allowed_types = {
            "mileage",
            "per_diem",
            "currency_conversion",
            "quick_calculation",
        }

        if calculation_type not in allowed_types:
            return False, (
                f"Invalid calculation_type. Allowed values: "
                f"{', '.join(sorted(allowed_types))}"
            )

        required_by_type = {
            "mileage": ["distance_km", "rate_per_km"],
            "per_diem": ["travel_days", "daily_allowance"],
            "currency_conversion": [
                "foreign_amount",
                "exchange_rate",
                "currency_code",
            ],
            "quick_calculation": ["number1", "number2", "operation"],
        }

        missing = [
            field for field in required_by_type[calculation_type]
            if arguments.get(field) is None
            or (
                isinstance(arguments.get(field), str)
                and not arguments[field].strip()
            )
        ]

        if missing:
            return False, (
                f"{calculation_type} requires: {', '.join(missing)}"
            )

        numeric_fields = {
            "mileage": ["distance_km", "rate_per_km"],
            "per_diem": ["travel_days", "daily_allowance"],
            "currency_conversion": ["foreign_amount", "exchange_rate"],
            "quick_calculation": ["number1", "number2"],
        }

        for field in numeric_fields[calculation_type]:
            if not is_number(arguments[field]):
                return False, f"{field} must be a finite, non-negative number."
            if not value_was_supplied(arguments[field]):
                return False, f"{field} was not supplied by the user or a verified tool. Ask the user instead of guessing."

        if calculation_type == "per_diem":
            days = arguments["travel_days"]
            if days <= 0 or int(days) != days:
                return False, "travel_days must be a positive whole number."

        if calculation_type == "currency_conversion":
            if not isinstance(arguments["currency_code"], str) or not arguments["currency_code"].strip():
                return False, "currency_code must be provided."
            if not value_was_supplied(arguments["currency_code"]):
                return False, "currency_code was not supplied by the user. Ask the user instead of guessing."
            if arguments["exchange_rate"] <= 0:
                return False, "exchange_rate must be greater than zero."

        if calculation_type == "quick_calculation":
            operation = arguments["operation"]

            if operation not in {"+", "-", "*", "/"}:
                return False, "operation must be one of +, -, *, /."

            if operation == "/" and arguments["number2"] == 0:
                return False, "Division by zero is not allowed."

    return True, "Arguments are valid."

async def main():

    if not all([LLM_BASE_URL, LLM_API_KEY, LLM_MODEL]):
        raise ValueError(
            "Missing required environment variables: set LLM_BASE_URL, "
            "LLM_API_KEY, and LLM_MODEL in the .env file."
        )

    print("Connecting to MCP server...")

    async with authenticated_mcp_client() as mcp_client:

        print("Connected successfully!")

        # 1. Discover available MCP tools
        tools_response = await mcp_client.list_tools()

        # 2. Convert MCP tools into LiteLLM-compatible tools
        llm_tools = convert_mcp_tools_to_llm_tools(
            tools_response.tools
        )

        print("\nTools available to the LLM:")

        for tool in llm_tools:
            print("-", tool["function"]["name"])

        # Keep the actual MCP tool definitions for validation.
        available_tools = {
            tool.name: tool
            for tool in tools_response.tools
        }

        # 3. Get the user's natural-language request
        user_prompt = input("\nWhat would you like to do? ")

        messages = [
            {
                "role": "system",
                "content": (
                    "You are an intelligent travel reimbursement assistant connected to an MCP server. "
                    "Understand the user's intent, select the correct MCP tool, provide accurate arguments, "
                    "and report only results confirmed by the tools.\n\n"

                    "TOOL SELECTION RULES:\n"
                    "1. Create a new workbook -> create_excel.\n"
                    "2. Read, inspect, retrieve, or summarize an existing workbook -> read_excel.\n"
                    "3. List available workbooks -> list_excel_files.\n"
                    "4. Add a NEW expense -> write_excel.\n"
                    "5. Modify an EXISTING expense -> update_excel.\n"
                    "6. Add employee or travel information -> write_excel with employee_info.\n"
                    "7. Modify employee or travel information -> update_excel with employee_updates.\n"
                    "8. Record or change a declaration -> write_excel or update_excel, respectively.\n"
                    "9. Update an advance received amount -> update_excel with advance_received.\n"
                    "10. Calculate mileage, per diem, currency conversion, or arithmetic -> calculate_expense.\n"
                    "10. Submit a claim to HR -> submit_to_hr.\n"
                    "11. Never use write_excel to modify an existing expense.\n"
                    "12. Never use update_excel to create a new expense.\n"
                    "13. Never create a workbook when the user specifies an existing workbook.\n"
                    "14. Never interpret a filename alone as a request to create a workbook.\n"
                    "15. If the user asks to read a workbook without specifying a worksheet, retrieve all worksheets.\n"
                    "16. If the user names a worksheet, always pass that exact sheet_name.\n"
                    "17. Do not request formula expressions unless the user explicitly asks to inspect formulas.\n\n"

                    "WORKBOOK SELECTION RULES:\n"
                    "1. If the user provides a filename, use that exact filename.\n"
                    "2. Do NOT call list_excel_files when the user has already provided a filename.\n"
                    "3. If the user needs to work with an existing workbook but gives no filename, "
                    "call list_excel_files first.\n"
                    "4. If only one relevant workbook is available, use it.\n"
                    "5. If multiple workbooks are available and the intended one is unclear, ask the user.\n"
                    "6. Do not ask the user to repeat information that is already available in the workbook.\n"
                    "7. Do not call read_excel before a calculation if the filename and all calculation "
                    "inputs are already provided.\n"
                    "8. For updates, inspect the workbook when necessary to identify the target record.\n"
                    "9. If the user requests a new workbook but gives no filename, ask for one.\n\n"

                    "EXPENSE CREATION RULES:\n"
                    "1. Use write_excel to add exactly ONE new expense per call.\n"
                    "2. Required arguments are filename, expense_id, expense_date, category, description, and amount.\n"
                    "3. receipt_reference is optional.\n"
                    "4. Pass expense_id as a single string, never as a dictionary or multiple IDs.\n"
                    "5. Supported categories are Transport, Food, Accommodation, Local Conveyance, and Miscellaneous.\n"
                    "6. Expense dates must use YYYY-MM-DD format.\n"
                    "7. Never invent an expense ID, date, description, category, or amount.\n"
                    "8. If a required value is missing, ask the user for it. Do not call write_excel until "
                    "all required values are available.\n"
                    "9. Do not assume the current date is the expense date.\n"
                    "10. A complete travel start date and end date must already be recorded before adding an expense. "
                    "If either date is missing, ask the user for the travel dates and record them before creating the expense.\n"
                    "11. Do not use generic descriptions such as 'Travel expense' unless the user explicitly "
                    "provides or confirms that description.\n"
                    "12. If multiple expenses are requested, make one separate write_excel call per expense.\n"
                    "13. Wait for each tool result before reporting success.\n\n"

                    "EMPLOYEE AND TRAVEL INFORMATION RULES:\n"
                    "1. Use write_excel with employee_info to add employee or travel information.\n"
                    "2. Supported editable fields are employee_name, employee_id, department, designation, email, "
                    "travel_purpose, source_location, destination_location, travel_start_date, and travel_end_date.\n"
                    "3. Use the exact supported field names.\n"
                    "4. Never invent employee details, dates, locations, or other values.\n"
                    "5. Ask for missing essential information.\n"
                    "6. Dates must use YYYY-MM-DD format.\n"
                    "7. travel_days is calculated automatically from travel_start_date and travel_end_date; never write it.\n"
                    "8. Employee information and expense creation are separate operations. "
                    "Use separate calls when both are requested.\n"
                    "9. Use update_excel with employee_updates to modify existing employee information.\n"
                    "10. Use only supported employee field names in employee_updates.\n"
                    "11. The declaration argument accepts only 'Yes' or 'No'. Never assume the user's declaration.\n"
                    "12. Employee information and declaration may be recorded together when requested.\n"
                    "13. If the user provides an advance received amount, pass it as advance_received in update_excel; never omit it.\n"
                    "13. A normal employee/travel record must include every required field: employee_name, employee_id, department, designation, email, travel_purpose, source_location, destination_location, travel_start_date, and travel_end_date.\n"
                    "14. If any required field is missing, ask for it and do not save a partial record.\n"
                    "15. Partial saving is allowed only when the user explicitly asks to save a draft; then pass save_as_draft=true.\n"
                    "16. Never claim fields were saved unless they appear in the tool result's updated_fields.\n\n"

                    "EXPENSE UPDATE RULES:\n"
                    "1. Use update_excel only for an existing expense.\n"
                    "2. Identify the expense using its Expense ID.\n"
                    "3. If the Expense ID is not provided, inspect the workbook to identify the intended expense.\n"
                    "4. If the intended expense remains unclear, ask the user.\n"
                    "5. The updates argument may contain only these supported fields: date, category, "
                    "description, amount, and receipt_reference.\n"
                    "6. Change only the fields explicitly requested by the user.\n"
                    "7. Never modify Expense ID or the calculated Date Check field.\n"
                    "8. Do not overwrite unrelated information.\n"
                    "9. Include all requested changes for the same expense in one update_excel call.\n\n"

                    "CALCULATION RULES:\n"
                    "1. Use calculate_expense for supported reimbursement calculations.\n"
                    "2. calculation_type must be EXACTLY one of: mileage, per_diem, "
                    "currency_conversion, quick_calculation.\n"
                    "3. Never use shortened or alternative values such as 'conversion' or 'per diem' "
                    "as the calculation_type. Use the exact values above.\n"
                    "4. Mileage requires filename, distance_km, and rate_per_km.\n"
                    "5. Per diem requires filename, travel_days, and daily_allowance.\n"
                    "6. Currency conversion requires filename, foreign_amount, exchange_rate, and currency_code.\n"
                    "7. Quick calculation requires filename, number1, number2, and operation.\n"
                    "8. Supported quick-calculation operations are +, -, *, and /.\n"
                    "9. Map the user's wording directly to the correct arguments.\n"
                    "10. Example: 'Calculate per diem for 3 days at 1500 per day in Test_Claim.xlsx' "
                    "means calculation_type='per_diem', travel_days=3, daily_allowance=1500, "
                    "filename='Test_Claim.xlsx'.\n"
                    "11. Example: 'Convert 100 USD using an exchange rate of 83 in Test_Claim.xlsx' "
                    "means calculation_type='currency_conversion', foreign_amount=100, "
                    "exchange_rate=83, currency_code='USD', filename='Test_Claim.xlsx'.\n"
                    "12. Never omit a value that the user has explicitly provided.\n"
                    "13. Never invent a mileage rate, daily allowance, exchange rate, currency code, "
                    "or calculation operand.\n"
                    "14. If a required input is genuinely missing, ask the user before calling the tool.\n"
                    "15. Never divide by zero.\n"
                    "16. Do not claim a calculation was saved unless the tool confirms success.\n\n"

                    "CLAIM SUBMISSION RULES:\n"
                    "1. Call submit_to_hr only when the user explicitly asks to submit or send the claim to HR.\n"
                    "2. Use the existing workbook filename.\n"
                    "3. Never create a new workbook just to submit a claim.\n"
                    "4. Never submit automatically after adding or updating an expense.\n"
                    "5. If submission fails because information is missing, explain which fields need completion.\n"
                    "6. Never invent missing information or claim submission success without tool confirmation.\n"
                    "7. If the claim is already submitted, report its existing status and reference.\n\n"

                    "MCP EXECUTION RULES:\n"
                    "1. Use the available MCP tools to perform the requested operations.\n"
                    "2. Select the tool that directly matches the user's intent.\n"
                    "3. Pass the exact argument names and values required by that tool.\n"
                    "4. Do not send unexpected or unnecessary arguments.\n"
                    "5. Do not call unrelated tools or perform operations the user did not request.\n"
                    "6. Before calling a tool, check that all required arguments are available.\n"
                    "7. If an essential argument is missing, ask the user instead of guessing.\n"
                    "8. After a tool call, inspect its actual result.\n"
                    "9. If the tool returns status='error', report the error. Do not claim success.\n"
                    "10. Use actual workbook data when answering workbook-related questions.\n"
                    "11. Never fabricate workbook contents, calculation results, tool outputs, or confirmations.\n"
                    "12. Distinguish blank cells from actual values and calculated results.\n"
                    "13. When read_excel returns a formula expression or formula_note, do not invent its calculated value; explain that the workbook formula may need Excel recalculation.\n"
                    "14. Never claim an operation succeeded unless the MCP tool confirms it.\n"
                    "15. Do not repeat an operation that has already succeeded.\n"
                    "16. If a tool call fails, explain the actual issue instead of inventing an explanation.\n"
                    "17. Use only facts stated by the user or returned by verified tools. Never invent IDs, names, dates, locations, rates, or amounts.\n"
                    "18. When validation or execution identifies missing/invalid information, ask the user for that specific information and wait for the reply before retrying.\n"
                    "19. For employee writes, report only fields listed in updated_fields and report record_status/missing_fields accurately.\n"
                    "20. When a tool succeeds, provide the confirmed result and finish the task; do not ask unnecessary follow-up questions.\n"
                    "21. If the user declines, says no, stop, or cancels a proposed action, do not call a tool for that action; acknowledge the cancellation.\n"
                ),
            },
            {
                "role": "user",
                "content": user_prompt,
            },
        ]

        # 4. Ask the LLM to respond and select a tool. The loop continues
        # only when clarification is needed or the current task needs more
        # tool steps. A completed task exits without an unnecessary prompt.
        awaiting_user_reply = False
        verified_employee_write = None
        max_rounds = 12

        for _ in range(max_rounds):
            try:
                response = await acompletion(
                    model=f"openai/{LLM_MODEL}",
                    api_base=LLM_BASE_URL,
                    api_key=LLM_API_KEY,
                    messages=messages,
                    tools=llm_tools,
                    tool_choice="auto",
                    temperature=0,
                    max_tokens=2048,
                )
            except Exception as error:
                print(f"\nLLM request failed: {error}")
                print("No further action was taken. Check the model endpoint/key and retry.")
                break
            
            
            choices = getattr(response, "choices", None) or []
            if not choices:
                print("\nThe model returned no response choices. No MCP tool was called.")
                break
            assistant_message = choices[0].message

            if not assistant_message.tool_calls:
                final_response = clean_assistant_response(
                    assistant_message.content
                )
                verified_response = format_verified_employee_write(verified_employee_write)
                if verified_response:
                    final_response = verified_response
                print("\nAssistant:")
                print(final_response or "The model returned an empty response.")

                # Ask for another user message only when the model is asking
                # a clarification question or a previous tool call failed.
                clarification_needed = (
                    awaiting_user_reply
                    or is_clarification_request(final_response)
                )
                if clarification_needed:
                    user_reply = input("\nYou (provide the missing/corrected information, or press Enter to stop): ").strip()
                    if not user_reply:
                        break
                    messages.append({"role": "user", "content": user_reply})
                    awaiting_user_reply = False
                    continue

                break

            messages.append(assistant_message.model_dump(exclude_none=True))

            for tool_call in assistant_message.tool_calls:
                tool_name = tool_call.function.name
                try:
                    tool_arguments = json.loads(tool_call.function.arguments)
                except (json.JSONDecodeError, TypeError):
                    tool_arguments = None

                # If the latest user request names exactly one known worksheet,
                # preserve that explicit scope even if the model omits sheet_name.
                if tool_name == "read_excel" and isinstance(tool_arguments, dict):
                    latest_user_text = next(
                        (str(message.get("content", "")) for message in reversed(messages)
                         if message.get("role") == "user"),
                        "",
                    ).casefold()
                    known_sheets = (
                        "Employee & Travel Information",
                        "Expense Details",
                        "Expense Calculator",
                        "Claim Summary & HR",
                    )
                    mentioned_sheets = [
                        name for name in known_sheets if name.casefold() in latest_user_text
                    ]
                    if len(mentioned_sheets) == 1 and not tool_arguments.get("sheet_name"):
                        tool_arguments["sheet_name"] = mentioned_sheets[0]

                print(f"\nLLM selected tool: {tool_name}")
                print("Arguments:", tool_arguments)
                if not (
                    tool_name == "write_excel"
                    and isinstance(tool_arguments, dict)
                    and tool_arguments.get("employee_info") is not None
                ):
                    verified_employee_write = None

                if awaiting_user_reply:
                    tool_output = {
                        "status": "error",
                        "message": (
                            "A previous tool call failed. Ask the user for the "
                            "missing or corrected information before making another tool call."
                        ),
                    }
                    tool_failed = True
                elif tool_name not in available_tools:
                    tool_output = {"status": "error", "message": "Requested tool is not available."}
                    tool_failed = True
                elif tool_arguments is None:
                    tool_output = {"status": "error", "message": "Invalid JSON tool arguments."}
                    tool_failed = True
                else:
                    conversation_evidence = "\n".join(
                        str(message.get("content", ""))
                        for message in messages
                        if message.get("role") in {"user", "tool"}
                    )
                    is_valid, validation_message = validate_tool_arguments(
                        available_tools[tool_name],
                        tool_arguments,
                        conversation_evidence,
                    )

                    if not is_valid:
                        tool_output = {"status": "error", "message": validation_message}
                        tool_failed = True
                    else:
                        try:
                            result = await mcp_client.call_tool(tool_name, tool_arguments)
                            result_data = get_tool_result_text(result)
                            tool_failed = bool(getattr(result, "is_error", False))
                            if isinstance(result_data, dict) and result_data.get("status") == "error":
                                tool_failed = True
                            if (
                                tool_name == "write_excel"
                                and isinstance(tool_arguments, dict)
                                and tool_arguments.get("employee_info") is not None
                                and isinstance(result_data, dict)
                                and result_data.get("status") == "success"
                            ):
                                verified_employee_write = result_data
                            tool_output = {
                                "is_error": tool_failed,
                                "result": result_data,
                            }
                        except Exception as error:
                            tool_output = {"status": "error", "message": str(error)}
                            tool_failed = True

                if tool_failed:
                    awaiting_user_reply = True
                    verified_employee_write = None

                print("Tool result:", json.dumps(tool_output, ensure_ascii=False, default=str))
                messages.append({
                    "role": "tool",
                    "tool_call_id": tool_call.id,
                    "content": json.dumps(tool_output, ensure_ascii=False, separators=(",", ":"), default=str),
                })
        else:
            print("\nMaximum interaction rounds reached. The task may be incomplete; review the last tool result before retrying.")



if __name__ == "__main__":
    asyncio.run(main())
