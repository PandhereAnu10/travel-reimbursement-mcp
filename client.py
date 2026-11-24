
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
                    "3. Reading an existing workbook must NEVER trigger create_excel.\n"
                    "4. Creating a new workbook must NEVER trigger read_excel unless "
                    "the user explicitly requests both operations.\n"
                    "5. Do not interpret a workbook filename as a request to create "
                    "a workbook. Determine the operation from the user's actual request.\n"
                    "6. When the user asks to read a workbook without specifying a "
                    "particular worksheet, retrieve and inspect all available worksheets.\n"
                    "7. When the user specifies a particular worksheet, retrieve and "
                    "focus on that worksheet.\n"
                    "8. If the user asks to read an existing workbook but does not "
                    "provide its filename, ask for the filename rather than creating "
                    "a new workbook.\n\n"

                    "EXECUTION RULES:\n"
                    "1. Use the available MCP tools to perform the requested operation.\n"
                    "2. Pass the correct arguments required by the selected tool.\n"
                    "3. Do not call unrelated tools or perform additional operations "
                    "that the user did not request.\n"
                    "4. Ask for missing information only when it is essential "
                    "to complete the requested operation.\n"
                    "5. After receiving a tool result, inspect its contents and "
                    "provide the user with the actual result in a clear format.\n"
                    "6. For workbook-reading requests, answer using the retrieved "
                    "workbook data. Do not invent or assume missing values.\n"
                    "7. Never claim that an operation succeeded unless the MCP tool "
                    "confirms successful execution.\n"
                ),
            },

            {
                "role": "user",
                "content": user_prompt,
            },
        ]

        # 4. Ask the LLM to respond and select a tool
        for _ in range(5):

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

