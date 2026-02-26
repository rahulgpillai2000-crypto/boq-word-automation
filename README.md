# boq-word-automation

This project generates a Word business case from an Excel master file using Python.

## What the script does

`generate_business_cases.py` performs the following flow:

1. Reads Excel data from `data/master_projects.xlsx`.
2. Extracts required fields:
   - `Contract_Number`
   - `Work_Order_Number`
   - `Project_Name`
   - `Date`
   - `Start_Date`
   - `End_Date`
   - `Total_Value`
3. Calculates project duration in days.
4. Converts duration to months using `ceil(days / 30)`.
5. Divides total value equally across the calculated months.
6. Opens Word template `templates/business_case_template.docx`.
7. Replaces placeholders:
   - `{{Contract_Number}}`
   - `{{Work_Order_Number}}`
   - `{{Project_Name}}`
   - `{{Date}}`
   - `{{Total_Value}}`
8. Finds `{{CASHFLOW_TABLE}}` and replaces it with a dynamic 2-row table:
   - Row 1: `Month 1`, `Month 2`, ...
   - Row 2: equal monthly contractor fee values
9. Saves output to `output/output.docx`.

## Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## How to run

```bash
python generate_business_cases.py
```

## Template placeholder example

Use these placeholders inside `templates/business_case_template.docx`:

- Contract Number: `{{Contract_Number}}`
- Work Order Number: `{{Work_Order_Number}}`
- Project Name: `{{Project_Name}}`
- Date: `{{Date}}`
- Total Value: `{{Total_Value}}`
- Cashflow insertion point: `{{CASHFLOW_TABLE}}`
