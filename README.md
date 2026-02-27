# boq-word-automation

This script generates a Word business case from Excel even when fields are **not in one table**.

## Supported Excel layouts

The script reads from `data/master_projects.xlsx` and supports:

1. **Tabular layout** (headers in one row):
   - `Contract_Number`, `Work_Order_Number`, `Project_Name`, `Date`, `Start_Date`, `End_Date`, `Total_Value`
2. **Key-value layout** across one or multiple sheets, for example:
   - `Contract Number` in one cell and its value in the next cell
   - `Project Name` on another sheet with value next to it

The script automatically searches sheets and common alias labels (e.g., `Contract Number`, `Work Order`, `Start Date`, `Contract Value`).

## How to place placeholders in Word (very simple)

Open `templates/business_case_template.docx` and type placeholders exactly where values should appear:

- `{{Contract_Number}}`
- `{{Work_Order_Number}}`
- `{{Project_Name}}`
- `{{Date}}`
- `{{Total_Value}}`
- `{{CASHFLOW_TABLE}}` (this is where the monthly table will be inserted)

### Example paragraph in Word

`This business case is for project {{Project_Name}} under contract {{Contract_Number}}.`

### Example for cashflow location

Add a heading like `Cashflow` and on the next line type only:

`{{CASHFLOW_TABLE}}`

When script runs, that marker is replaced by a generated table.

## Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run

```bash
python generate_business_cases.py
```

## Output

Generated file:

- `output/output.docx`
