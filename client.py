
import asyncio
import sys
from pathlib import Path
import json
import os

from mcp import Client, StdioServerParameters
from dotenv import load_dotenv
from litellm import acompletion

LLM_BASE_URL = os.getenv("LLM_BASE_URL")
LLM_API_KEY = os.getenv("LLM_API_KEY")
LLM_MODEL = os.getenv("LLM_MODEL")

# project dir
PROJECT_DIR = Path(__file__).resolve().parent
load_dotenv(PROJECT_DIR / ".env")

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
    # Extract MCP tool res as text form LLM

    content_parts = []

    for item in result.content:
        if hasattr(item, "text"):
            content_parts.append(item.text)
        else:
            content_parts.append(str(item))

    return "\n".join(content_parts)


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
                    "You are a travel reimbursement assistant. "
                    "Use the available tools to complete the user's "
                    "request. Do not claim a task succeeded unless "
                    "the tool confirms it. Ask for missing information "
                    "when necessary."
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
            )

            assistant_message = response
            
            print(assistant_message.model_dump_json(indent=2))

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
                            tool_arguments,
                        )

                        tool_output = {
                            "is_error": result.isError,
                            "result": get_tool_result_text(result),
                        }

                    except Exception as error:
                        tool_output = {
                            "status": "error",
                            "message": str(error),
                        }

                # 6. Return the MCP result to the LLM
                messages.append({
                    "role": "tool",
                    "tool_call_id": tool_call.id,
                    "content": json.dumps(tool_output),
                })

        else:
            print("\nMaximum tool-call rounds reached.")


if __name__ == "__main__":
    asyncio.run(main())

