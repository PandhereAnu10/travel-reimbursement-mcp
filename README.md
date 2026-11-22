
# Travel Reimbursement MCP Server

An MCP-based Travel Reimbursement system that helps employees manage travel expense claims through an Excel workbook and AI-powered tool interactions.

## Project Overview

The project uses the Model Context Protocol (MCP) to expose tools for creating, reading, writing, updating, calculating, and submitting travel reimbursement claims.

The Excel workbook consists of four worksheets:

1. **Employee & Travel Information** – Stores employee details and travel information.
2. **Expense Details** – Maintains individual expense records, categories, dates, amounts, and receipt references.
3. **Expense Calculator** – Provides mileage, per diem, currency conversion, and basic arithmetic calculations.
4. **Claim Summary & HR** – Displays category-wise expense totals, net payable amount, employee declaration, and HR submission details.

## Technology Stack

- Python
- MCP SDK
- OpenPyXL
- LiteLLM
- Pydantic
- Pytest

## Project Structure

```text
travel-reimbursement-mcp/
│
├── templates/
│   ├── __init__.py
│   └── reimbursement.py
│
├── tools/
│   └── __init__.py
│
├── tests/
│
├── data/
│   └── generated/
│
├── server.py
├── client.py
├── requirements.txt
├── .env.example
├── .gitignore
└── README.md
```

## Current Implementation Status

- [x] Initial project structure
- [x] Dependency configuration
- [x] MCP server initialization and connectivity check
- [x] Workbook configuration for all four worksheets
- [ ] Excel workbook creation
- [ ] Excel read, write, and update operations
- [ ] Expense calculation tools
- [ ] HR submission functionality
- [ ] MCP client integration with LiteLLM
- [ ] Automated testing

## Installation

Clone the repository:

```bash
git clone <repository-url>
cd travel-reimbursement-mcp
```

Create and activate a virtual environment:

```bash
python -m venv .venv
```

Install dependencies:

```bash
pip install -r requirements.txt
```

