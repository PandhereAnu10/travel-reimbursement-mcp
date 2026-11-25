
import asyncio
import sys
from pathlib import Path
import json
import os

from mcp import Client, StdioServerParameters
from dotenv import load_dotenv
from litellm import acompletion

# project dir
PROJECT_DIR = Path(__file__).resolve().parent
load_dotenv(PROJECT_DIR / ".env")

LLM_BASE_URL = os.getenv("LLM_BASE_URL")
LLM_API_KEY = os.getenv("LLM_API_KEY")
LLM_MODEL = os.getenv("LLM_MODEL")

# for missing env var
if not all([LLM_BASE_URL, LLM_API_KEY, LLM_MODEL]):
    raise ValueError(
        "Missing required env variables."
    )

# MCP server config
SERVER_PARAMS = StdioServerParameters(
    command=sys.executable,
    args=[str(PROJECT_DIR / "server.py")],
)


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

    for item in result.content:
        if hasattr(item, "text"):
            content_parts.append(item.text)
        else:
            content_parts.append(str(item))

    text_result = "\n".join(content_parts)

    try:
        return json.loads(text_result)
    except json.JSONDecodeError:
        return text_result


async def main():

    print("Connecting to MCP server...")

    async with Client(SERVER_PARAMS) as mcp_client:

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

        # 3. Get the user's natural-language request
        user_prompt = input("\nWhat would you like to do? ")

        messages = [
            {
                "role": "system",
                "content": (
                    "You are an intelligent travel reimbursement assistant. "
                    "Understand the user's actual intent and select the most appropriate "
                    "available MCP tool to perform the requested operation.\n\n"

                    "TOOL SELECTION RULES:\n"
                    "1. If the user asks to create, generate, or make a new workbook, "
                    "call create_excel.\n"
                    "2. If the user asks to read, retrieve, inspect, summarize, explain, "
                    "or provide information from an existing workbook, call read_excel.\n"
                    "3. If the user asks to list, show, or find available Excel workbooks, "
                    "call list_excel_files.\n"
                    "4. If the user asks to add, insert, record, or append a NEW expense, "
                    "call write_excel.\n"
                    "5. If the user asks to modify, correct, or change an EXISTING expense, "
                    "call update_excel.\n"
                    "6. Never use write_excel to modify an existing expense.\n"
                    "7. Never use update_excel to create a new expense.\n"
                    "8. Reading an existing workbook must NEVER trigger create_excel.\n"
                    "9. Creating a new workbook must NEVER trigger read_excel unless "
                    "the user explicitly requests both operations.\n"
                    "10. Do not interpret a workbook filename alone as a request "
                    "to create a workbook.\n"
                    "11. If the user asks to read a workbook without specifying a "
                    "worksheet, retrieve all available worksheets.\n"
                    "12. If the user specifies a worksheet, focus on that worksheet.\n"
                    "13. If the filename or another essential argument is missing, "
                    "ask the user rather than guessing.\n\n"

                    "EXPENSE CREATION RULES:\n"
                    "1. A new expense requires filename, expense_id, expense_date, "
                    "category, description, and amount.\n"
                    "2. Receipt reference is optional.\n"
                    "3. Use a unique Expense ID for every expense.\n"
                    "4. Use only the configured expense categories: Transport, Food, "
                    "Accommodation, Local Conveyance, and Miscellaneous.\n"
                    "5. Expense dates must use YYYY-MM-DD format.\n"
                    "6. Never invent expense information or IDs.\n"
                    "7. Do not overwrite existing expense records.\n"
                    "8. The write_excel tool adds EXACTLY ONE expense per call.\n"
                    "9. Never pass multiple expenses inside the expense_id argument.\n"
                    "10. Never pass a dictionary or nested object of expense records "
                    "as the expense_id argument.\n"
                    "11. Always pass expense_id as a single string.\n"
                    "12. For each write_excel call, pass these separate arguments: "
                    "filename, expense_id, expense_date, category, description, "
                    "amount, and receipt_reference.\n"
                    "13. If the user requests multiple new expenses, call write_excel "
                    "separately for each expense, using one individual tool call "
                    "per expense.\n"
                    "14. Never combine multiple expense records into one tool call.\n"
                    "15. Wait for each tool result before reporting that expense "
                    "as successfully added.\n"
                    "16. If any expense insertion fails, report its error and do not "
                    "claim that it was added.\n"
                    "17. Do not create a new workbook when the user has specified "
                    "an existing workbook for expense insertion.\n\n"

                    "EXPENSE UPDATE RULES:\n"
                    "1. An existing expense must be identified using its Expense ID.\n"
                    "2. The updates argument must contain only the fields the user "
                    "explicitly wants to change.\n"
                    "3. Supported fields are date, category, description, amount, "
                    "and receipt_reference.\n"
                    "4. Never modify Expense ID or the calculated Date Check field.\n"
                    "5. Never overwrite unrelated expense information.\n"
                    "6. If the Expense ID is missing, ask the user for it.\n"
                    "7. If the user requests an update but the target expense "
                    "cannot be identified, do not guess which record to modify.\n"
                    "8. Use one update_excel call for the specified existing expense "
                    "and include all requested changes in its updates dictionary.\n\n"

                    "CLAIM SUBMISSION RULES:\n"
                    "1. If the user explicitly asks to submit, finalize, "
                    "or send a reimbursement claim to HR, call submit_to_hr.\n"
                    "2. The submit_to_hr tool requires the existing workbook filename.\n"
                    "3. Never create a new workbook just to submit a claim.\n"
                    "4. Never submit a claim automatically after adding or updating "
                    "an expense.\n"
                    "5. The tool validates required employee and travel information, "
                    "at least one expense record, and employee declaration.\n"
                    "6. If required information is missing, explain exactly which "
                    "fields need to be completed. Do not invent values.\n"
                    "7. The employee declaration must be accepted before submission.\n"
                    "8. A successful submission returns the submission status, "
                    "submission date, and unique claim reference.\n"
                    "9. If the claim is already submitted, report its existing "
                    "status and reference. Do not create a duplicate submission.\n"
                    "10. Never claim that a claim was submitted unless the "
                    "submit_to_hr tool confirms success.\n\n"

                    "EXECUTION RULES:\n"
                    "1. Use available MCP tools to perform requested operations.\n"
                    "2. Pass the correct arguments required by the selected tool.\n"
                    "3. Do not call unrelated tools or perform additional operations "
                    "that the user did not request.\n"
                    "4. Ask for missing information only when essential.\n"
                    "5. After receiving a tool result, inspect its contents and "
                    "provide the actual result clearly.\n"
                    "6. If a tool returns status 'error', explain the error. "
                    "Do not claim that the operation succeeded.\n"
                    "7. For workbook-reading requests, use retrieved workbook data. "
                    "Never invent missing values.\n"
                    "8. Distinguish between blank cells, formula strings, and actual "
                    "calculated values.\n"
                    "9. Do not interpret blank template fields as confirmed information.\n"
                    "10. Avoid unnecessary explanations of Excel formulas when "
                    "the user asks for actual workbook information.\n"
                    "11. Never claim an operation succeeded unless the MCP tool "
                    "confirms successful execution.\n"
                    "12. Never output raw tool-call markup as a substitute for "
                    "executing an MCP tool.\n"
                    "13. If a tool call cannot be executed, clearly report the issue "
                    "instead of pretending it succeeded.\n"
                    "14. Do not fabricate tool results or execution confirmations.\n"
                ),
            },
            {
                "role": "user",
                "content": user_prompt,
            },
        ]

        # 4. Ask the LLM to respond and select a tool
        for _ in range(6):

            response = await acompletion(
                model=f"openai/{LLM_MODEL}",
                api_base=LLM_BASE_URL,
                api_key=LLM_API_KEY,
                messages=messages,
                tools=llm_tools,
                tool_choice="auto",
                max_tokens=2048,
            )

            assistant_message = response.choices[0].message
            
            #print("Assistant:", assistant_message.model_dump_json(indent=2))

            # If the LLM doesn't request a tool, show its final answer.
            if not assistant_message.tool_calls:

                print("\nAssistant:")
                print(assistant_message.content)
                break

            # Add the LLM's tool-call message to the conversation.
            messages.append(
                assistant_message.model_dump(exclude_none=True)
            )

            # 5. Execute each requested MCP tool
            for tool_call in assistant_message.tool_calls:

                tool_name = tool_call.function.name
                tool_arguments = json.loads(
                    tool_call.function.arguments
                )

                print(f"\nLLM selected tool: {tool_name}")
                print("Arguments:", tool_arguments)

                # Only execute tools discovered from this MCP server.
                available_names = {
                    tool.name for tool in tools_response.tools
                }

                if tool_name not in available_names:
                    tool_output = {
                        "status": "error",
                        "message": "Requested tool is not available.",
                    }

                else:
                    try:
                        result = await mcp_client.call_tool(
                            tool_name,
                            tool_arguments
                        )

                        tool_output = {
                            "is_error": result.is_error,
                            "result": get_tool_result_text(result),
                        }

                        #print("\nMCP Tool Result:")
                        #print(json.dumps(tool_output, indent=2))

                    except Exception as error:
                        tool_output = {
                            "status": "error",
                            "message": str(error),
                        }

                # 6. Return the MCP result to the LLM
                messages.append({
                    "role": "tool",
                    "tool_call_id": tool_call.id,
                    "content": json.dumps(
                        tool_output,
                        ensure_ascii=False,
                        separators=(",", ":")
                    ),
                })

        else:
            print("\nMaximum tool-call rounds reached.")


if __name__ == "__main__":
    asyncio.run(main())

