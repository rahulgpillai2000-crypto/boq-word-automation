#!/usr/bin/env python3
"""Generate Microsoft Word business case documents from an Excel master data source.

Usage examples:
  python generate_business_cases.py --excel data/projects.xlsx --template templates/business_case_template.docx --output-dir output
  python generate_business_cases.py --excel data/projects.xlsx --template templates/business_case_template.docx --output-dir output --project-id-column Contract_Number --project-id CN-1002
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Dict, Iterable, List, Tuple

import pandas as pd
from docx import Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import Table
from docx.text.paragraph import Paragraph

PLACEHOLDER_PATTERN = re.compile(r"\{\{\s*([A-Za-z0-9_]+)\s*\}\}")


class TemplateError(RuntimeError):
    """Raised when template rendering cannot be completed."""


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Generate or update Word business-case documents using Excel as the single "
            "source of truth."
        )
    )
    parser.add_argument("--excel", required=True, help="Path to Excel master data file.")
    parser.add_argument("--sheet", default=0, help="Excel sheet name or index (default: first sheet).")
    parser.add_argument(
        "--template",
        required=True,
        help="Path to the .docx business case template containing placeholders like {{Project_Name}}.",
    )
    parser.add_argument("--output-dir", required=True, help="Directory where generated .docx files are written.")
    parser.add_argument(
        "--filename-column",
        default="Contract_Number",
        help="Column to use for output filename stem (default: Contract_Number).",
    )
    parser.add_argument(
        "--project-id-column",
        default=None,
        help="Optional column used with --project-id to generate only one selected project.",
    )
    parser.add_argument(
        "--project-id",
        default=None,
        help="Optional project identifier value to generate a single document.",
    )
    return parser.parse_args()


def sanitize_filename(name: str) -> str:
    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", name.strip())
    return safe.strip("._") or "document"


def format_cell_value(value: object) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    return str(value)


def build_mapping(row: pd.Series) -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    for column, value in row.items():
        mapping[str(column)] = format_cell_value(value)
    return mapping


def replace_in_runs(paragraph: Paragraph, mapping: Dict[str, str]) -> Tuple[int, List[str]]:
    if not paragraph.runs:
        return 0, []

    original_text = "".join(run.text for run in paragraph.runs)
    placeholders_found = PLACEHOLDER_PATTERN.findall(original_text)
    if not placeholders_found:
        return 0, []

    replaced_text = original_text
    missing_keys: List[str] = []
    for key in placeholders_found:
        placeholder = "{{" + key + "}}"
        if key in mapping:
            replaced_text = re.sub(r"\{\{\s*" + re.escape(key) + r"\s*\}\}", mapping[key], replaced_text)
        else:
            missing_keys.append(key)
            replaced_text = re.sub(r"\{\{\s*" + re.escape(key) + r"\s*\}\}", "", replaced_text)

    # Keep formatting by writing full text to first run and clearing others.
    paragraph.runs[0].text = replaced_text
    for run in paragraph.runs[1:]:
        run.text = ""

    return len(placeholders_found), sorted(set(missing_keys))


def iter_all_paragraphs(document: Document) -> Iterable[Paragraph]:
    for paragraph in document.paragraphs:
        yield paragraph

    for table in document.tables:
        yield from iter_table_paragraphs(table)

    for section in document.sections:
        for paragraph in section.header.paragraphs:
            yield paragraph
        for paragraph in section.footer.paragraphs:
            yield paragraph
        for table in section.header.tables:
            yield from iter_table_paragraphs(table)
        for table in section.footer.tables:
            yield from iter_table_paragraphs(table)


def iter_table_paragraphs(table: Table) -> Iterable[Paragraph]:
    for row in table.rows:
        for cell in row.cells:
            for child in cell._tc.iterchildren():
                if isinstance(child, CT_P):
                    yield Paragraph(child, cell)
                elif isinstance(child, CT_Tbl):
                    nested_table = Table(child, cell)
                    yield from iter_table_paragraphs(nested_table)


def render_document(template_path: Path, output_path: Path, mapping: Dict[str, str]) -> Tuple[int, List[str]]:
    document = Document(template_path)
    replacement_count = 0
    missing_keys: List[str] = []

    for paragraph in iter_all_paragraphs(document):
        count, missing = replace_in_runs(paragraph, mapping)
        replacement_count += count
        missing_keys.extend(missing)

    document.save(output_path)
    return replacement_count, sorted(set(missing_keys))


def load_excel_data(excel_path: Path, sheet: str | int) -> pd.DataFrame:
    data_frame = pd.read_excel(excel_path, sheet_name=sheet)
    if data_frame.empty:
        raise TemplateError("The Excel file does not contain any data rows.")

    data_frame.columns = [str(column).strip() for column in data_frame.columns]
    return data_frame


def filter_rows(data_frame: pd.DataFrame, project_id_column: str | None, project_id: str | None) -> pd.DataFrame:
    if project_id_column and project_id is None:
        raise TemplateError("--project-id-column was provided, but --project-id is missing.")
    if project_id and project_id_column is None:
        raise TemplateError("--project-id was provided, but --project-id-column is missing.")

    if not project_id_column:
        return data_frame

    if project_id_column not in data_frame.columns:
        raise TemplateError(
            f"Project ID column '{project_id_column}' was not found in Excel headers: {list(data_frame.columns)}"
        )

    filtered = data_frame[data_frame[project_id_column].astype(str) == str(project_id)]
    if filtered.empty:
        raise TemplateError(f"No row found where {project_id_column} == '{project_id}'.")
    return filtered


def main() -> None:
    args = parse_args()

    excel_path = Path(args.excel).expanduser().resolve()
    template_path = Path(args.template).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve()

    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")
    if not template_path.exists():
        raise FileNotFoundError(f"Word template not found: {template_path}")

    output_dir.mkdir(parents=True, exist_ok=True)

    data_frame = load_excel_data(excel_path, args.sheet)
    data_frame = filter_rows(data_frame, args.project_id_column, args.project_id)

    if args.filename_column not in data_frame.columns:
        raise TemplateError(
            f"Filename column '{args.filename_column}' was not found in Excel headers: {list(data_frame.columns)}"
        )

    for _, row in data_frame.iterrows():
        mapping = build_mapping(row)
        filename_value = mapping.get(args.filename_column, "") or "business_case"
        file_stem = sanitize_filename(filename_value)
        output_path = output_dir / f"{file_stem}.docx"

        replacements, missing_keys = render_document(template_path, output_path, mapping)
        print(f"Generated: {output_path} | placeholders processed: {replacements}")
        if missing_keys:
            print(f"  Warning: placeholders missing in Excel data -> {', '.join(missing_keys)}")


if __name__ == "__main__":
    main()
