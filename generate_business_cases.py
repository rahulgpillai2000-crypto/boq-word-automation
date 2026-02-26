#!/usr/bin/env python3
"""Generate a business case Word document from a master Excel data source.

Required behavior:
- Read project data from data/master_projects.xlsx
- Use templates/business_case_template.docx as template
- Replace key placeholders
- Replace {{CASHFLOW_TABLE}} with a generated 2-row monthly cashflow table
- Save final file to output/output.docx
"""

from __future__ import annotations

import math
from datetime import datetime
from pathlib import Path
from typing import Dict

import pandas as pd
from docx import Document

EXCEL_PATH = Path("data/master_projects.xlsx")
TEMPLATE_PATH = Path("templates/business_case_template.docx")
OUTPUT_PATH = Path("output/output.docx")

REQUIRED_COLUMNS = [
    "Contract_Number",
    "Work_Order_Number",
    "Project_Name",
    "Date",
    "Start_Date",
    "End_Date",
    "Total_Value",
]


class BusinessCaseError(RuntimeError):
    """Raised for business case generation errors."""


def _to_datetime(value: object, column_name: str) -> datetime:
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        raise BusinessCaseError(f"Invalid date in column '{column_name}': {value!r}")
    return parsed.to_pydatetime()


def _format_value(value: object) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    return str(value)


def load_first_project_row(excel_path: Path) -> pd.Series:
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    data = pd.read_excel(excel_path)
    if data.empty:
        raise BusinessCaseError("Excel file contains no project rows.")

    data.columns = [str(col).strip() for col in data.columns]
    missing = [col for col in REQUIRED_COLUMNS if col not in data.columns]
    if missing:
        raise BusinessCaseError(f"Missing required Excel columns: {', '.join(missing)}")

    return data.iloc[0]


def compute_cashflow(row: pd.Series) -> Dict[str, float | int]:
    start_date = _to_datetime(row["Start_Date"], "Start_Date")
    end_date = _to_datetime(row["End_Date"], "End_Date")
    if end_date < start_date:
        raise BusinessCaseError("End_Date cannot be earlier than Start_Date.")

    duration_days = (end_date - start_date).days + 1
    duration_months = max(1, math.ceil(duration_days / 30))

    total_value = float(row["Total_Value"])
    monthly_fee = total_value / duration_months

    return {
        "duration_days": duration_days,
        "duration_months": duration_months,
        "total_value": total_value,
        "monthly_fee": monthly_fee,
    }


def replace_simple_placeholders(document: Document, replacements: Dict[str, str]) -> None:
    for paragraph in document.paragraphs:
        for key, value in replacements.items():
            placeholder = "{{" + key + "}}"
            if placeholder in paragraph.text:
                paragraph.text = paragraph.text.replace(placeholder, value)

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for key, value in replacements.items():
                        placeholder = "{{" + key + "}}"
                        if placeholder in paragraph.text:
                            paragraph.text = paragraph.text.replace(placeholder, value)


def insert_cashflow_table(document: Document, duration_months: int, monthly_fee: float) -> None:
    marker = "{{CASHFLOW_TABLE}}"

    for paragraph in document.paragraphs:
        if marker in paragraph.text:
            paragraph.text = paragraph.text.replace(marker, "")
            table = document.add_table(rows=2, cols=duration_months)
            table.style = "Table Grid"

            for month in range(duration_months):
                table.cell(0, month).text = f"Month {month + 1}"
                table.cell(1, month).text = f"{monthly_fee:,.2f}"

            paragraph._p.addnext(table._tbl)
            return

    raise BusinessCaseError("Placeholder {{CASHFLOW_TABLE}} not found in template.")


def main() -> None:
    row = load_first_project_row(EXCEL_PATH)
    cashflow = compute_cashflow(row)

    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Word template not found: {TEMPLATE_PATH}")

    document = Document(TEMPLATE_PATH)

    replacements = {
        "Contract_Number": _format_value(row["Contract_Number"]),
        "Work_Order_Number": _format_value(row["Work_Order_Number"]),
        "Project_Name": _format_value(row["Project_Name"]),
        "Date": _format_value(pd.to_datetime(row["Date"], errors="coerce")),
        "Total_Value": f"{float(row['Total_Value']):,.2f}",
    }

    replace_simple_placeholders(document, replacements)
    insert_cashflow_table(document, int(cashflow["duration_months"]), float(cashflow["monthly_fee"]))

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    document.save(OUTPUT_PATH)

    print(f"Generated Word document: {OUTPUT_PATH}")
    print(
        "Computed values -> "
        f"Duration (days): {cashflow['duration_days']}, "
        f"Duration (months): {cashflow['duration_months']}, "
        f"Monthly contractor fee: {cashflow['monthly_fee']:,.2f}"
    )


if __name__ == "__main__":
    main()
