# Travel Reimbursement MCP Server

An MCP-based Travel Reimbursement system that enables employees to create, manage, calculate, and submit travel expense claims through Excel workbooks using natural-language instructions and an AI-powered client.

## 1. Project Overview

This project implements a custom Model Context Protocol (MCP) server in Python. It exposes tools that allow an AI assistant to interact with Excel workbooks for travel reimbursement management.

The system follows this workflow:

**User → LLM → MCP Client → MCP Server → Excel Workbook**

Each reimbursement workbook contains four worksheets:

1. **Employee & Travel Information** – Stores employee details, department, designation, travel purpose, destination, and travel dates.
2. **Expense Details** – Maintains expense records, categories, dates, amounts, receipt references, and date validation.
3. **Expense Calculator** – Supports mileage, per-diem, currency-conversion, and quick arithmetic calculations.
4. **Claim Summary & HR** – Displays expense summaries, net payable amount, employee declaration, and HR submission details.

## 2. Architecture

```text
User
  |
  v
LiteLLM (LLM Provider)
  |
  v
client.py (MCP Client)
  |
  | MCP tool calls
  v
server.py (MCP Server)
  |
  +---- security.py (Authentication / Security)
  |
  +---- tools/
  |       |
  |       +---- create_excel
  |       +---- read_excel
  |       +---- write_excel
  |       +---- update_excel
  |       +---- list_excel_files
  |       +---- calculate_expense
  |       +---- submit_to_hr
  |
  v
templates/ + OpenPyXL
  |
  v
Excel Workbook
  |
  +---- Employee & Travel Information
  +---- Expense Details
  +---- Expense Calculator
  +---- Claim Summary & HR
```

The server exposes individual tools for workbook operations. The LLM client interprets user requests, selects the appropriate tools, handles missing information through follow-up questions, and reports the result.

## 3. Technology Stack

| Technology | Purpose |
|---|---|
| Python 3.11+ | Core implementation |
| Official Python MCP SDK | MCP server and client communication |
| OpenPyXL | Excel workbook creation and manipulation |
| LiteLLM | LLM API integration |
| Pydantic | Input validation and tool argument schemas |
| Pytest | Automated testing |
| MCP Inspector | Manual MCP tool inspection and testing |
| ngrok | Public access to the Streamable HTTP server |

## 4. Project Structure

```text
travel-reimbursement-mcp/
│
├── templates/
│   ├── formatting.py
│   ├── formulas.py
│   ├── reimbursement.py
│   └── validations.py
│
├── tools/
│   ├── __init__.py
│   ├── common_func.py
│   ├── create_excel.py
│   ├── list_excel_files.py
│   ├── read_excel.py
│   ├── calculate_expense.py
│   ├── submit_to_hr.py
│   ├── update_excel.py
│   └── write_excel.py
│
├── tests/
│   ├── test_calculate_expense.py
│   ├── test_client_validation.py
│   ├── test_common_func.py
│   ├── test_create_excel.py
│   ├── test_list_excel_files.py
│   ├── test_mcp_server.py
│   ├── test_read_excel.py
│   ├── test_submit_to_hr.py
│   ├── test_update_excel.py
│   └── test_write_excel.py
│
├── data/
│   └── generated/
│       └── Demo_Travel_Reimbursement.xlsx
│
├── security.py
├── server.py
├── client.py
├── requirements.txt
├── .env.example
├── .gitignore
└── README.md
```

The implementation separates workbook templates, MCP tools, security, and automated tests into individual modules to keep the project organized and maintainable.

## 5. Available MCP Tools

| Tool | Description |
|---|---|
| `create_excel` | Creates a reimbursement workbook with the required worksheets. |
| `read_excel` | Reads workbook data and optionally a specified worksheet. |
| `write_excel` | Adds employee information and expense records. |
| `update_excel` | Updates existing employee or expense information. |
| `list_excel_files` | Lists available reimbursement workbooks. |
| `calculate_expense` | Performs mileage, per-diem, currency-conversion, and quick calculations. |
| `submit_to_hr` | Validates a claim and records its HR submission details. |

## 6. Implementation Highlights

- [x] Workbook creation with four structured worksheets.
- [x] Excel formatting, formulas, and input validations.
- [x] Employee and travel information management.
- [x] Expense creation and updating.
- [x] Expense totals and reimbursement calculations.
- [x] Mileage, per-diem, currency-conversion, and arithmetic calculations.
- [x] Claim validation and HR submission workflow.
- [x] LiteLLM-powered natural-language client.
- [x] Follow-up questions for missing or invalid information.
- [x] Automated tests.
- [x] Local MCP server execution.
- [x] Streamable HTTP access through ngrok.

Travel days are calculated using the travel start and end dates. The calculator writes input values while preserving formula-driven result cells.

## 7. Installation and Configuration

### Prerequisites

- Python 3.11 or newer
- Git
- A LiteLLM-compatible API endpoint and credentials
- ngrok (for public HTTP access)

### Clone the repository

```bash
git clone https://github.com/PandhereAnu10/travel-reimbursement-mcp.git
cd travel-reimbursement-mcp
```

### Create a virtual environment

```bash
python -m venv .venv
```

Activate the virtual environment:

```powershell
.\.venv\Scripts\activate
```

### Install dependencies

```bash
pip install -r requirements.txt
```

### Configure environment variables

Copy `.env.example` to `.env` and configure the values supplied by your model provider.

The environment configuration includes:

```env
LLM_BASE_URL=
LLM_API_KEY=
LLM_MODEL=
MCP_SERVER_URL=http://127.0.0.1:8000/mcp
MCP_HOST=127.0.0.1
MCP_PORT=8000
MCP_AUTH_TOKEN=
```

Generate a secure authentication token using:

```bash
python -c "import secrets; print(secrets.token_urlsafe(32))"
```

Copy the generated token into the `MCP_AUTH_TOKEN` field in your `.env` file.

## 8. Running the Application

### Run the LLM client

The client starts the local MCP server as a child process and handles natural-language interactions.

```bash
python client.py
```

### Run the MCP server using Streamable HTTP

Start the server:

```bash
python server.py
```

The local MCP endpoint is:

```text
http://127.0.0.1:8000/mcp
```

Generated workbooks are stored under:

```text
data/generated/
```

## 9. Connect via ngrok

The MCP server can be exposed publicly using ngrok, allowing an external MCP client or reviewer to connect remotely.

### Step 1: Start the MCP server

```bash
python server.py
```

### Step 2: Start ngrok in another terminal

```bash
ngrok http 8000
```

### Step 3: Public MCP endpoint

The current public MCP endpoint is:

```text
https://ghostlier-adrianna-unsuccinctly.ngrok-free.dev/mcp
```

This is the endpoint to use when connecting through MCP Inspector or another compatible remote MCP client.

### Step 4: Configure authentication

The MCP endpoint uses bearer-token authentication configured through `MCP_AUTH_TOKEN` in `.env`.

Include the following HTTP header when connecting:

```text
Authorization: Bearer <MCP_AUTH_TOKEN>
```

For ngrok browser-warning handling, include:

```text
ngrok-skip-browser-warning: true
```

The authentication token must be shared privately with the reviewer. Never commit the actual token to GitHub or include it in this README.

### Step 5: Verify using MCP Inspector

1. Open MCP Inspector.
2. Select Streamable HTTP as the transport.
3. Enter the public MCP endpoint shown above.
4. Add the required Authorization header.
5. Add the ngrok browser-warning header if needed.
6. Connect and inspect the available MCP tools.
7. Execute individual tools to verify their responses.


## 10. Testing

Run the automated test suite:

```bash
python -m pytest -q
```

**Latest test result:** 183 tests passed.

MCP tools can also be tested independently using MCP Inspector.

The testing covers:

- Workbook creation and reading.
- Employee information and expense writing.
- Updating existing expenses.
- Expense calculations.
- Invalid calculator inputs.
- Missing workbook and record handling.
- Invalid expense categories.
- Incomplete and complete claim submissions.
- Natural-language client interactions and input validation.

## 11. Example User Prompts

The following prompts demonstrate the intended interactions:

**Create a workbook**

> Create a reimbursement workbook for employee EMP001, Rahul Sharma, travelling from Mumbai to Pune.

**Add expenses**

> Add ₹850 for train travel, ₹350 for lunch and ₹2,000 for accommodation.

**Calculate mileage**

> I drove 120 km locally at ₹12 per km. Calculate it and add it as a local conveyance expense.

**Read expenses**

> Show me all expenses recorded for EMP001 and the total claim amount.

**Update an expense**

> Change the accommodation expense from ₹2,000 to ₹2,500.

**Submit to HR**

> I received an advance of ₹1,000. Accept the declaration and submit the claim to HR.

**List workbooks**

> Show me all reimbursement Excel files created so far.

**One-shot reimbursement request**

> Make a travel reimbursement for me. I flew from Delhi to Mumbai on 10 Nov for ₹6,200 and back on 14 Nov for ₹5,800. Hotel from 10 to 14 Nov at ₹3,500 per night.

The client should ask follow-up questions when required details are missing instead of inventing values.

## 12. Known Limitations

- The system uses Excel workbooks as its primary data store rather than a database.
- Formula-driven results may require Excel or another compatible spreadsheet application to recalculate when the workbook is opened.
- Employee travel days are derived from travel dates and are not directly editable.
- The LLM client and direct MCP tool calls use their respective supported workflows for collecting and updating claim information.
- The public ngrok endpoint is temporary and may change when the tunnel restarts.

## 13. Sample Output

A sample reimbursement workbook is included in `data/generated/`:

```text
Demo_Travel_Reimbursement.xlsx
```

It demonstrates the four-sheet workbook structure, employee and travel details, expense records, calculator inputs, and HR submission information.

## 14. Demo Recording

A screen recording demonstrates the main reimbursement workflow, MCP tool execution through MCP Inspector, calculator operations, and the generated Excel workbook.

**Demo video:** <Add your demo video link here>

---

## Project Links

- **GitHub Repository:** https://github.com/PandhereAnu10/travel-reimbursement-mcp
- **Public MCP Endpoint:** https://ghostlier-adrianna-unsuccinctly.ngrok-free.dev/mcp
- **Demo Recording:** <Add your demo video link here>
