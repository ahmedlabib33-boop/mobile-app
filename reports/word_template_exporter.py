from __future__ import annotations

from copy import deepcopy
import re
from pathlib import Path
from typing import Any

import pandas as pd


def sanitize_filename(value: str) -> str:
    cleaned = re.sub(r"[^\w\-\.]+", "_", str(value or "").strip())
    cleaned = re.sub(r"_+", "_", cleaned).strip("._")
    return cleaned or "output"


def normalize_header(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().lower())


def replace_text_preserve_format(doc, replacements: dict) -> None:
    replacements = {str(k): str(v) for k, v in replacements.items() if str(k)}
    if not replacements:
        return

    def replace_in_paragraph(paragraph) -> None:
        full_text = "".join(run.text for run in paragraph.runs)
        if not full_text:
            return
        updated = full_text
        for source, target in replacements.items():
            updated = updated.replace(source, target)
        if updated == full_text:
            return
        if paragraph.runs:
            paragraph.runs[0].text = updated
            for run in paragraph.runs[1:]:
                run.text = ""
        else:
            paragraph.add_run(updated)

    def replace_in_table(table) -> None:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_in_paragraph(paragraph)
                for nested in cell.tables:
                    replace_in_table(nested)

    for paragraph in doc.paragraphs:
        replace_in_paragraph(paragraph)
    for table in doc.tables:
        replace_in_table(table)
    for section in doc.sections:
        for paragraph in section.header.paragraphs:
            replace_in_paragraph(paragraph)
        for table in section.header.tables:
            replace_in_table(table)
        for paragraph in section.footer.paragraphs:
            replace_in_paragraph(paragraph)
        for table in section.footer.tables:
            replace_in_table(table)


def _find_header_row_index(table, required_headers: list[str]) -> int | None:
    required = [normalize_header(header) for header in required_headers]
    for row_index, row in enumerate(table.rows[: min(5, len(table.rows))]):
        row_headers = [normalize_header(cell.text) for cell in row.cells]
        if all(any(req == hdr or req in hdr for hdr in row_headers) for req in required):
            return row_index
    return None


def find_table_by_headers(doc, required_headers: list[str]):
    for table in doc.tables:
        if _find_header_row_index(table, required_headers) is not None:
            return table
    return None


def clone_table_row_format(table, source_row_index: int):
    new_tr = deepcopy(table.rows[source_row_index]._tr)
    table._tbl.append(new_tr)
    return table.rows[len(table.rows) - 1]


def _clear_cell_preserve_format(cell) -> None:
    for paragraph in cell.paragraphs:
        if paragraph.runs:
            paragraph.runs[0].text = ""
            for run in paragraph.runs[1:]:
                run.text = ""
        else:
            paragraph.add_run("")


def _write_cell_preserve_format(cell, value: Any) -> None:
    text = "" if value is None else str(value)
    if cell.paragraphs and cell.paragraphs[0].runs:
        cell.paragraphs[0].runs[0].text = text
        for run in cell.paragraphs[0].runs[1:]:
            run.text = ""
    elif cell.paragraphs:
        cell.paragraphs[0].add_run(text)
    else:
        paragraph = cell.add_paragraph()
        paragraph.add_run(text)


def write_cell_preserve_format(cell, value: Any) -> None:
    _clear_cell_preserve_format(cell)
    _write_cell_preserve_format(cell, value)


def update_table_preserve_format(
    table,
    data,
    column_mapping: dict[str, str] | None = None,
    required_headers: list[str] | None = None,
    allow_remove_extra: bool = False,
) -> None:
    if data is None:
        return
    df = data.copy() if isinstance(data, pd.DataFrame) else pd.DataFrame(data)
    if df.empty:
        return
    header_row_index = _find_header_row_index(table, required_headers or list(df.columns))
    if header_row_index is None:
        return

    header_cells = [cell.text.strip() for cell in table.rows[header_row_index].cells]
    column_mapping = column_mapping or {}
    resolved_columns: list[str | None] = []
    for header in header_cells:
        direct = df.columns[df.columns.map(lambda c: normalize_header(c)) == normalize_header(header)]
        if len(direct):
            resolved_columns.append(direct[0])
            continue
        mapped = column_mapping.get(header) or column_mapping.get(normalize_header(header))
        resolved_columns.append(mapped if mapped in df.columns else None)

    existing_data_rows = len(table.rows) - (header_row_index + 1)
    source_row_index = header_row_index + 1 if existing_data_rows > 0 else header_row_index
    while len(table.rows) < header_row_index + 1 + len(df):
        clone_table_row_format(table, source_row_index)

    for offset, (_, df_row) in enumerate(df.iterrows(), start=1):
        row = table.rows[header_row_index + offset]
        for cell_index, cell in enumerate(row.cells):
            _clear_cell_preserve_format(cell)
            if cell_index < len(resolved_columns) and resolved_columns[cell_index]:
                _write_cell_preserve_format(cell, df_row.get(resolved_columns[cell_index], ""))

    if allow_remove_extra:
        target_len = header_row_index + 1 + len(df)
        while len(table.rows) > target_len:
            table._tbl.remove(table.rows[-1]._tr)


def update_keyed_table_preserve_format(
    table,
    row_key_column_index: int,
    value_updates: dict[str, list[Any]],
) -> None:
    normalized_map = {normalize_header(key): values for key, values in value_updates.items()}
    for row in table.rows[1:]:
        if row_key_column_index >= len(row.cells):
            continue
        row_key = normalize_header(row.cells[row_key_column_index].text)
        values = normalized_map.get(row_key)
        if not values:
            continue
        for idx, value in enumerate(values, start=row_key_column_index + 1):
            if idx < len(row.cells):
                _write_cell_preserve_format(row.cells[idx], value)
