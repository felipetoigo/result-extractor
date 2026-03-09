"""Write table data to XLSX with configurable column order."""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from .config import OUTPUT_COLUMN_ORDER, OUTPUT_SHEET_NAME


def reorder_columns(
    rows: list[list[str]],
    column_order: list[str] | list[int],
    has_header: bool = True,
) -> list[list[str]]:
    """
    Reorder columns so they match the desired order.

    If column_order contains strings, the first row is treated as header and
    column order is by header name. If column_order contains integers, they
    are 0-based column indices.

    Args:
        rows: Table rows (list of list of cell values).
        column_order: Desired column order (header names or indices).
        has_header: If True, first row is header and used for name-based order.

    Returns:
        New list of rows with columns reordered.
    """
    if not rows:
        return []

    if not column_order:
        return rows

    header = rows[0] if has_header else None
    data_rows = rows[1:] if has_header else rows

    # Build index mapping: output position -> source column index
    if column_order and isinstance(column_order[0], int):
        # Indices: assume column_order is list of 0-based indices
        index_map = list(column_order)
    else:
        # Header names: find index for each name in the header
        if not header:
            return rows
        name_to_index = {str(h).strip(): i for i, h in enumerate(header)}
        index_map = []
        for name in column_order:
            idx = name_to_index.get(str(name).strip())
            if idx is not None:
                index_map.append(idx)
            # If name not in header, we could append empty column; for now skip

    def reorder_row(row: list[str]) -> list[str]:
        return [row[i] if i < len(row) else "" for i in index_map]

    out_header = reorder_row(header) if header else []
    out_data = [reorder_row(r) for r in data_rows]
    return [out_header] + out_data if out_header else out_data


def write_xlsx(
    rows: list[list[str]],
    output_path: str | Path,
    sheet_name: str | None = None,
    column_order: list[str] | list[int] | None = None,
    has_header: bool = True,
) -> Path:
    """
    Write rows to an XLSX file with optional column reordering.

    Args:
        rows: Table data (list of rows, each row is list of cells).
        output_path: Path for the output .xlsx file.
        sheet_name: Sheet name; uses config default if None.
        column_order: Desired column order; uses config if None.
        has_header: Whether first row is header (for column order by name).

    Returns:
        Resolved output path.
    """
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    order = column_order if column_order is not None else OUTPUT_COLUMN_ORDER
    name = sheet_name if sheet_name is not None else OUTPUT_SHEET_NAME

    ordered = reorder_columns(rows, order, has_header=has_header) if order else rows

    wb = Workbook()
    ws: Worksheet = wb.active
    ws.title = name[:31]  # Excel sheet name length limit

    for row in ordered:
        ws.append(row)

    wb.save(path)
    return path.resolve()


def write_spreadsheet(
    header_rows: list[tuple[int, str, str]] | list[tuple[str, str]],
    table_rows: list[list[str]],
    output_path: str | Path,
    sheet_name: str | None = None,
    column_order: list[str] | list[int] | None = None,
    has_header: bool = True,
) -> Path:
    """
    Write a single sheet with (1) header section, (2) blank row, (3) table.

    Args:
        header_rows: Personal/header data. Either [(page_num, field, value), ...]
            (with page numbers from all pages) or [(field, value), ...] for a single block.
        table_rows: Table data (list of rows); first row is header if has_header.
        output_path: Path for the output .xlsx file.
        sheet_name: Sheet name; uses config default if None.
        column_order: Optional column order for the table section.
        has_header: Whether first row of table_rows is a header row.

    Returns:
        Resolved output path.
    """
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    name = sheet_name if sheet_name is not None else OUTPUT_SHEET_NAME
    order = column_order if column_order is not None else OUTPUT_COLUMN_ORDER

    table_part = (
        reorder_columns(table_rows, order, has_header=has_header)
        if order and table_rows
        else table_rows
    )

    wb = Workbook()
    ws: Worksheet = wb.active
    ws.title = name[:31]

    # Header section: with page numbers (Page | Field | Value) or plain (Field | Value)
    if header_rows and isinstance(header_rows[0][0], int):
        ws.append(["Page", "Field", "Value"])
        for row in header_rows:
            ws.append([row[0], row[1], row[2]])
    else:
        ws.append(["Field", "Value"])
        for row in header_rows:
            ws.append([row[0], row[1]])
    ws.append([])  # blank row

    # Table
    for row in table_part:
        ws.append(row)

    wb.save(path)
    return path.resolve()
