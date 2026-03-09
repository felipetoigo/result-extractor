"""
Column order and extraction settings.

Once you have the PDF, update OUTPUT_COLUMN_ORDER to match the desired
Excel column order. Use the exact header names as they appear in the PDF
(or as normalized), or use column indices if headers vary.
"""

# Desired column order in the output XLSX (left to right).
# Update this list when you have the PDF: use the PDF's column headers (strings)
# or 0-based indices (integers). Example: ["Name", "Date", "Result"] or [0, 2, 1]
OUTPUT_COLUMN_ORDER: list[str] | list[int] = [
    # Placeholder: add your column names or indices after you have the PDF.
    # Example: "Column A", "Column B", "Column C"
]

# If the PDF has no header row, set to False and we'll use column indices.
HAS_HEADER_ROW: bool = True

# Sheet name in the output workbook.
OUTPUT_SHEET_NAME: str = "Results"
