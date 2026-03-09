"""Orchestrate PDF → XLSX conversion."""

from pathlib import Path

from .config import HAS_HEADER_ROW, OUTPUT_COLUMN_ORDER
from .excel_writer import write_xlsx, write_spreadsheet
from .pdf_reader import extract_tables, extract_header_and_tables


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

    # 3) Output path and options
    if output_dir is None:
        output_dir = Path.home() / "Desktop"
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    out_path = output_dir / f"exported_{timestamp}.xlsx"
    order = column_order if column_order is not None else OUTPUT_COLUMN_ORDER
    header_flag = has_header if has_header is not None else HAS_HEADER_ROW

    # 4) Write once, outside any page loop
    return write_spreadsheet(
        all_header_rows,
        combined_table_rows,
        out_path,
        column_order=order if order else None,
        has_header=header_flag,
    )
