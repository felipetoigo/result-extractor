"""Extract table data and header (personal data) from PDF files."""

import re
from pathlib import Path

import pdfplumber


def _find_pagina_markers(page) -> list[tuple[float, float, int]]:
    """
    Find 'Página X/Y' on the page and return sorted (top, bottom, page_num) for each.
    Used to split one physical page into N logical pages.
    """
    words = page.extract_words() or []
    if not words:
        return []
    # Build (top, bottom, page_num) for each "Página X/Y"
    markers: list[tuple[float, float, int]] = []
    # Match "Página" and then "X/Y" (e.g. "1/4") in following words
    for i, w in enumerate(words):
        text = (w.get("text") or "").strip()
        if "Página" not in text and "página" not in text:
            continue
        # Get bbox of this word (and optionally next if "1/4" is separate)
        x0, top, x1, bottom = w["x0"], w["top"], w["x1"], w["bottom"]
        page_num = None
        # Check if number is in same word: "Página 1/4" or "1/4"
        match = re.search(r"(\d+)\s*/\s*(\d+)", text)
        if not match and i + 1 < len(words):
            next_text = (words[i + 1].get("text") or "").strip()
            match = re.search(r"(\d+)\s*/\s*(\d+)", next_text)
            if match:
                bottom = max(bottom, words[i + 1]["bottom"])
                x1 = max(x1, words[i + 1]["x1"])
        if match:
            page_num = int(match.group(1))
        if page_num is not None:
            markers.append((top, bottom, page_num))
    # Sort by vertical position (top), then by page_num
    markers.sort(key=lambda m: (m[0], m[2]))
    return markers


def _logical_pages_from_physical(page) -> list[tuple[object, int]]:
    """
    Return list of (cropped_page, logical_page_num) for this physical page.
    If 'Página X/Y' markers are found, split into N logical pages; else return [(page, 1)].
    """
    markers = _find_pagina_markers(page)
    if len(markers) <= 1:
        return [(page, 1)]
    # Sort by top to get order; use bottom of each marker as end of that logical page
    markers_sorted = sorted(markers, key=lambda m: m[0])
    results: list[tuple[object, int]] = []
    for idx, (top, bottom, page_num) in enumerate(markers_sorted):
        top_y = markers_sorted[idx - 1][1] if idx > 0 else 0
        bottom_y = bottom
        if bottom_y <= top_y:
            continue
        cropped = page.within_bbox((0, top_y, page.width, bottom_y))
        results.append((cropped, page_num))
    return results if results else [(page, 1)]


def _parse_header_text(text: str | None) -> list[tuple[str, str]]:
    """Parse header block into key-value pairs. Lines with ':' or tab split into Field, Value."""
    if not text or not text.strip():
        return []
    pairs: list[tuple[str, str]] = []
    for line in text.strip().splitlines():
        line = line.strip()
        if not line:
            continue
        if ":" in line:
            key, _, value = line.partition(":")
            pairs.append((key.strip(), value.strip()))
        elif "\t" in line:
            key, _, value = line.partition("\t")
            pairs.append((key.strip(), value.strip()))
        else:
            pairs.append((line, ""))
    return pairs


def _extract_from_one_logical_page(cropped_page, logical_page_num: int) -> tuple[list[tuple[int, str, str]], list[list[list[str]]]]:
    """Extract header pairs and tables from one logical page (a physical page or a cropped region)."""
    header_rows: list[tuple[int, str, str]] = []
    tables: list[list[list[str]]] = []
    found_tables = cropped_page.find_tables()
    if found_tables:
        first_table = found_tables[0]
        top_y = first_table.bbox[1]
        if top_y > 0:
            header_region = cropped_page.within_bbox((0, 0, cropped_page.width, top_y))
            header_text = header_region.extract_text()
            pairs = _parse_header_text(header_text)
            for field, value in pairs:
                header_rows.append((logical_page_num, field, value))
        for table in cropped_page.extract_tables():
            normalized = [
                [str(cell).strip() if cell is not None else "" for cell in row]
                for row in table
            ]
            if normalized:
                tables.append(normalized)
    else:
        header_text = cropped_page.extract_text()
        pairs = _parse_header_text(header_text or "")
        for field, value in pairs:
            header_rows.append((logical_page_num, field, value))
        for table in cropped_page.extract_tables():
            normalized = [
                [str(cell).strip() if cell is not None else "" for cell in row]
                for row in table
            ]
            if normalized:
                tables.append(normalized)
    return header_rows, tables


def extract_header_and_tables(
    pdf_path: str | Path,
) -> tuple[list[tuple[int, str, str]], list[list[list[str]]]]:
    """
    Extract header section and tables from every page. All data is appended into
    two single lists (no overwriting). Supports multi-page PDFs and single long
    pages with "Página 1/4" etc.

    Returns:
        (all_header_rows, all_tables) — append-only; write to XLSX once outside.
    """
    path = Path(pdf_path)
    if not path.exists():
        raise FileNotFoundError(f"PDF not found: {path}")

    # Single accumulation lists — only append, never replace
    all_header_rows: list[tuple[int, str, str]] = []
    all_tables: list[list[list[str]]] = []

    with pdfplumber.open(path) as pdf:
        num_physical = len(pdf.pages)
        if num_physical > 1:
            # Multiple physical pages: process each page as page 1, 2, 3, 4 (no splitting)
            for page_num, physical_page in enumerate(pdf.pages, start=1):
                h, t = _extract_from_one_logical_page(physical_page, page_num)
                all_header_rows.extend(h)
                all_tables.extend(t)
        else:
            # Single physical page: try Página split, else process once
            physical_page = pdf.pages[0]
            logical_list = _logical_pages_from_physical(physical_page)
            for cropped_page, logical_page_num in logical_list:
                h, t = _extract_from_one_logical_page(cropped_page, logical_page_num)
                all_header_rows.extend(h)
                all_tables.extend(t)
    return all_header_rows, all_tables


def extract_tables(pdf_path: str | Path) -> list[list[list[str]]]:
    """
    Extract all tables from a PDF as a list of tables; each table is a list of rows.

    Each row is a list of cell strings. The first row is typically the header
    if HAS_HEADER_ROW is True in config.

    Args:
        pdf_path: Path to the PDF file.

    Returns:
        List of tables; each table is list of rows; each row is list of cell strings.
    """
    path = Path(pdf_path)
    if not path.exists():
        raise FileNotFoundError(f"PDF not found: {path}")

    tables: list[list[list[str]]] = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            page_tables = page.extract_tables()
            if page_tables:
                for table in page_tables:
                    # Normalize: ensure every cell is str, empty cells as ""
                    normalized = [
                        [str(cell).strip() if cell is not None else "" for cell in row]
                        for row in table
                    ]
                    if normalized:
                        tables.append(normalized)
    return tables


def extract_first_table(pdf_path: str | Path) -> list[list[str]]:
    """
    Extract the first table only (all rows as list of lists).

    Convenience when the PDF has a single table.

    Args:
        pdf_path: Path to the PDF file.

    Returns:
        First table as list of rows.
    """
    tables = extract_tables(pdf_path)
    if not tables:
        raise ValueError(f"No tables found in PDF: {pdf_path}")
    return tables[0]
