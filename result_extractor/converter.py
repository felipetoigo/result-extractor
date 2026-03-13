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


def _cell_to_float(cell) -> float:
    """Parse cell to float (Brazilian: comma=decimal, dot=thousands). Return 0.0 if not parseable."""
    if cell is None or cell == "":
        return 0.0
    if isinstance(cell, (int, float)):
        return float(cell)
    s = str(cell).strip()
    s = re.sub(r"R\$\s*", "", s, flags=re.I)
    s = "".join(c for c in s if c in "0123456789,.-").strip()
    if not s:
        return 0.0
    if "," in s:
        parts = s.split(",", 1)
        int_part = parts[0].replace(".", "").strip()
        dec_str = "".join(c for c in parts[1] if c.isdigit())[:2].ljust(2, "0") or "00"
        try:
            sign = -1 if int_part.startswith("-") else 1
            return sign * (abs(int(int_part)) + int(dec_str) / 100.0)
        except ValueError:
            return 0.0
    try:
        return float(s.replace(".", ""))
    except ValueError:
        return 0.0


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


def _fill_vcm_column(rows: list[list]) -> list[list]:
    """
    Fill VCM column with VALOR + CORREÇÃO MONETÁRIA for each data row.
    Uses normalized header matching (CORREÇÃO MONETÁRIA or short form Correção).
    """
    if not rows or len(rows) < 2:
        return rows
    header = rows[0]
    idx_valor = _find_header_index(header, "VALOR", exact=False)
    idx_correcao = _find_header_index(header, "CORREÇÃO MONETÁRIA", exact=False)
    if idx_correcao is None:
        idx_correcao = _find_header_index(header, "Correção", exact=False)
    idx_vcm = next((i for i, h in enumerate(header) if str(h).strip().upper() == "VCM"), None)
    if idx_valor is None or idx_correcao is None or idx_vcm is None:
        return rows
    result = [list(header)]
    for row in rows[1:]:
        new_row = list(row)
        if idx_vcm < len(new_row):
            valor = _cell_to_float(new_row[idx_valor] if idx_valor < len(row) else 0)
            correcao = _cell_to_float(new_row[idx_correcao] if idx_correcao < len(row) else 0)
            new_row[idx_vcm] = round(valor + correcao, 2)
        result.append(new_row)
    return result


def _fill_hr_ha_columns(rows: list[list]) -> list[list]:
    """
    Fill HR and HA columns with 20% of (VCM + JUROS + MULTA) for each data row.
    Same value in both columns. Example: (236.89 + 227.81 + 4.74) * 0.2 = 93.89.
    """
    if not rows or len(rows) < 2:
        return rows
    header = rows[0]
    idx_vcm = next((i for i, h in enumerate(header) if str(h).strip().upper() == "VCM"), None)
    idx_juros = _find_header_index(header, "JUROS", exact=False)
    idx_multa = _find_header_index(header, "MULTA", exact=False)
    idx_hr = next((i for i, h in enumerate(header) if str(h).strip().upper() == "HR"), None)
    idx_ha = next((i for i, h in enumerate(header) if str(h).strip().upper() == "HA"), None)
    if idx_vcm is None or idx_juros is None or idx_multa is None or idx_hr is None or idx_ha is None:
        return rows
    result = [list(header)]
    for row in rows[1:]:
        new_row = list(row)
        vcm = _cell_to_float(new_row[idx_vcm] if idx_vcm < len(row) else 0)
        juros = _cell_to_float(new_row[idx_juros] if idx_juros < len(row) else 0)
        multa = _cell_to_float(new_row[idx_multa] if idx_multa < len(row) else 0)
        value = round((vcm + juros + multa) * 0.2, 2)
        if idx_hr < len(new_row):
            new_row[idx_hr] = value
        if idx_ha < len(new_row):
            new_row[idx_ha] = value
        result.append(new_row)
    return result


def _drop_columns(rows: list[list], column_names: list[str]) -> list[list]:
    """
    Remove columns by header name (normalized match). Keeps all other columns in order.
    """
    if not rows:
        return rows
    header = rows[0]
    drop_set = {_normalize_header(name) for name in column_names}
    keep_indices = [
        i for i, h in enumerate(header)
        if _normalize_header(str(h).strip()) not in drop_set
    ]
    return [
        [(row[i] if i < len(row) else "") for i in keep_indices]
        for row in rows
    ]


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

    # 5) Fill VCM column: VCM = VALOR + CORREÇÃO MONETÁRIA for each data row
    combined_table_rows = _fill_vcm_column(combined_table_rows)

    # 6) Fill HR and HA columns: 20% of (VCM + JUROS + MULTA) for each data row
    combined_table_rows = _fill_hr_ha_columns(combined_table_rows)

    # 7) Drop HONORÁRIOS column (after all calculations)
    combined_table_rows = _drop_columns(combined_table_rows, ["HONORÁRIOS"])

    # 8) Output path and options
    if output_dir is None:
        output_dir = Path.home() / "Desktop"
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    out_path = output_dir / f"exported_{timestamp}.xlsx"

    # 9) Write once, outside any page loop
    return write_spreadsheet(
        all_header_rows,
        combined_table_rows,
        out_path,
        column_order=order if order else None,
        has_header=header_flag,
    )
