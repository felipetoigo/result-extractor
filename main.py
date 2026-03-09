#!/usr/bin/env python3
"""
Result Extractor CLI: convert PDF tables to XLSX with configurable column order.

Usage:
    python main.py input.pdf
    python main.py input.pdf -o output.xlsx
"""

import argparse
import sys
from pathlib import Path

# Allow running as script from repo root
sys.path.insert(0, str(Path(__file__).resolve().parent))

from result_extractor.converter import convert_pdf_to_xlsx


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extract table(s) from a PDF and save as XLSX with configurable column order."
    )
    parser.add_argument(
        "pdf",
        type=Path,
        help="Path to the input PDF file",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Output XLSX path (default: same name as PDF with .xlsx)",
    )
    parser.add_argument(
        "--all-tables",
        action="store_true",
        help="Include all tables from the PDF (default: first table only)",
    )
    args = parser.parse_args()

    if not args.pdf.exists():
        print(f"Error: file not found: {args.pdf}", file=sys.stderr)
        return 1

    try:
        out = convert_pdf_to_xlsx(
            args.pdf,
            xlsx_path=args.output,
            first_table_only=not args.all_tables,
        )
        print(f"Created: {out}")
        return 0
    except FileNotFoundError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
