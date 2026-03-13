"""Orchestrate PDF → XLSX conversion."""

import re
import unicodedata
from pathlib import Path

from .config import COLUMNS_TO_EXCLUDE, HAS_HEADER_ROW, OUTPUT_COLUMN_ORDER


def _normalize_header(name: str) -> str:
    """Normalize header for comparison: strip, collapse spaces, uppercase, remove accents (e.g. CORREÇÃO → CORRECAO)."""
    s = re.sub(r"\s+", " ", str(name).strip()).upper()
    nfd = unicodedata.normalize("NFD", s)
    return "".join(c for c in nfd if unicodedata.category(c) != "Mn")
from .excel_writer import write_xlsx, write_spreadsheet
from .pdf_reader import extract_tables, extract_header_and_tables


def _find_header_index(header: list, after_name: str, exact: bool = False) -> int | None:
    """Return 0-based index of column with name after_name. If exact, match by string; else normalized."""
    for i, h in enumerate(header):
        if exact:
            if str(h).strip() == after_name.strip():
                return i
        else:
            if _normalize_header(h) == _normalize_header(after_name):
                return i
    return None


def _add_blank_columns(rows: list[list], inserts: list[tuple[str | list[str], str]]) -> list[list]:
    """
    Insert blank columns after specified header names.
    inserts: list of (after_this_header, new_column_name). after_this_header can be a single name
    or a list of alternatives (tried in order; first match wins). Matching is normalized (no accents).
    """
    if not rows:
        return rows
    for after_name_or_list, new_name in inserts:
        header = rows[0]
        candidates = after_name_or_list if isinstance(after_name_or_list, list) else [after_name_or_list]
        idx = None
        for c in candidates:
            idx = _find_header_index(header, c, exact=False)
            if idx is not None:
                break
        if idx is None:
            continue
        insert_at = idx + 1
        new_header = list(header[:insert_at]) + [new_name] + list(header[insert_at:])
        new_rows = [new_header]
        for row in rows[1:]:
            new_row = list(row[:insert_at]) + [""] + list(row[insert_at:])
            new_rows.append(new_row)
        rows = new_rows
    return rows


def convert_pdf_to_xlsx(
    pdf_path: str | Path,
    xlsx_path: str | Path | None = None,
    *,
    column_order: list[str] | list[int] | None = None,
    has_header: bool | None = None,
    first_table_only: bool = True,
) -> Path:
    """
    Read table(s) from PDF and write to XLSX with the specified column order.

    Args:
        pdf_path: Input PDF file path.
        xlsx_path: Output XLSX path. If None, same name as PDF with .xlsx.
        column_order: Override config column order if provided.
        has_header: Override config HAS_HEADER_ROW if provided.
        first_table_only: If True, only the first table is written; else all tables concatenated.

    Returns:
        Path to the created XLSX file.
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    tables = extract_tables(pdf_path)
    if not tables:
        raise ValueError(f"No tables found in PDF: {pdf_path}")

    if first_table_only:
        rows = tables[0]
    else:
        # Concatenate: use first table's header, then all data rows from all tables
        header = tables[0][0] if tables[0] else []
        data_rows = tables[0][1:] if len(tables[0]) > 1 else []
        for t in tables[1:]:
            if t and (not HAS_HEADER_ROW or t[0] != header):
                data_rows.extend(t)
            elif t and HAS_HEADER_ROW:
                data_rows.extend(t[1:])
        rows = [header] + data_rows if header else data_rows

    out_path = Path(xlsx_path) if xlsx_path else pdf_path.with_suffix(".xlsx")
    order = column_order if column_order is not None else OUTPUT_COLUMN_ORDER
    header_flag = has_header if has_header is not None else HAS_HEADER_ROW

    return write_xlsx(
        rows,
        out_path,
        column_order=order if order else None,
        has_header=header_flag,
    )


def convert_pdf_to_spreadsheet(
    pdf_path: str | Path,
    output_dir: str | Path | None = None,
    *,
    column_order: list[str] | list[int] | None = None,
    has_header: bool | None = None,
    first_table_only: bool = False,
) -> Path:
    """
    Read PDF (header + table from all pages), build one combined dataset,
    then write to XLSX once. No per-page writes; no overwriting.

    Output file is named exported_<timestamp>.xlsx in output_dir (default: Desktop).
    """
    from datetime import datetime

    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    # 1) Extract everything into two lists (reader only appends)
    all_header_rows, all_tables = extract_header_and_tables(pdf_path)

    # 2) Build one combined list of table rows from all tables (append only)
    combined_table_rows: list[list[str]] = []
    if all_tables:
        first_header = all_tables[0][0] if (HAS_HEADER_ROW and all_tables[0]) else []
        if first_table_only:
            combined_table_rows = list(all_tables[0])
        else:
            combined_table_rows.append(first_header)
            for t in all_tables:
                data_start = 1 if (t and HAS_HEADER_ROW and t[0] == first_header) else 0
                for row in t[data_start:]:
                    combined_table_rows.append(row)

    order = column_order if column_order is not None else OUTPUT_COLUMN_ORDER
    header_flag = has_header if has_header is not None else HAS_HEADER_ROW

    # 3) Drop excluded columns (TÍTULO, ATRASO)
    if combined_table_rows and COLUMNS_TO_EXCLUDE and header_flag:
        header_row = combined_table_rows[0]
        exclude_set = {c.strip() for c in COLUMNS_TO_EXCLUDE}
        keep_indices = [i for i, h in enumerate(header_row) if str(h).strip() not in exclude_set]
        combined_table_rows = [[(row[i] if i < len(row) else "") for i in keep_indices] for row in combined_table_rows]

    # 4) Add blank columns VCM (after CORREÇÃO MONETÁRIA / Correção or VALOR), HR (after HONORÁRIOS), HA (after HR)
    combined_table_rows = _add_blank_columns(
        combined_table_rows,
        [
            (["CORREÇÃO MONETÁRIA", "Correção", "VALOR"], "VCM"),
            ("HONORÁRIOS", "HR"),
            ("HR", "HA"),
        ],
    )

    # 5) Output path and options
    if output_dir is None:
        output_dir = Path.home() / "Desktop"
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    out_path = output_dir / f"exported_{timestamp}.xlsx"

    # 6) Write once, outside any page loop
    return write_spreadsheet(
        all_header_rows,
        combined_table_rows,
        out_path,
        column_order=order if order else None,
        has_header=header_flag,
    )
