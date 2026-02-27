# boq-word-automation

Excel-to-Word automation that keeps project business case documents fully aligned with a master Excel source.

## What this does

- Reads project/contract records from Excel (`pandas`).
- Replaces placeholders inside a Word `.docx` template (`python-docx`).
- Generates one output document per Excel row (or a single selected row).
- Overwrites generated outputs when re-run, so updates are repeatable and consistent.

## 1) Excel master data format

Create an Excel sheet with one row per project and clear column headers.

Recommended headers:

- `Project_Name`
- `Contract_Number`
- `Contract_Value`
- `Date`
- `Client_Name`
- `Engineer_Consultant`
- `Work_Order_Details`
- `Financial_Total`
- `Variation_Total`

You can add more columns at any time; each column is automatically available as a Word placeholder key.

## 2) Word business case template format

Use placeholders wrapped in double curly braces anywhere in the document body, tables, headers, or footers.

Example placeholders:

- `{{Project_Name}}`
- `{{Contract_Number}}`
- `{{Contract_Value}}`
- `{{Date}}`
- `{{Client_Name}}`
- `{{Engineer_Consultant}}`
- `{{Work_Order_Details}}`
- `{{Financial_Total}}`
- `{{Variation_Total}}`

> Keep the template file unchanged and re-run the script whenever Excel data changes.

## 3) Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 4) Run automation (one-click)

Generate documents for **all projects**:

```bash
python generate_business_cases.py \
  --excel data/master_projects.xlsx \
  --sheet 0 \
  --template templates/business_case_template.docx \
  --output-dir output
```

Generate document for **one project only**:

```bash
python generate_business_cases.py \
  --excel data/master_projects.xlsx \
  --template templates/business_case_template.docx \
  --output-dir output \
  --project-id-column Contract_Number \
  --project-id CN-1002
```

## 5) How consistency is enforced

- Excel is the single source of truth.
- Word files are regenerated directly from Excel values.
- No manual typing is needed inside generated Word files.
- Running the script repeatedly updates outputs with latest Excel values.

## 6) Notes

- Output filenames are based on `Contract_Number` by default (configurable via `--filename-column`).
- If a placeholder exists in Word but not in Excel headers, the script clears it and prints a warning.
- Date cells are exported in `YYYY-MM-DD` format.
