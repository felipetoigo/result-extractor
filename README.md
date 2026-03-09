# Result Extractor

Python application that reads a PDF with **two sections**—a header (personal data such as residencial, CNPJ, etc.) and a table—and converts **all of this data** into a single spreadsheet. The output file is named **exported_&lt;timestamp&gt;.xlsx** (e.g. `exported_2025-03-09_14-30-00.xlsx`) and is saved on your Desktop.

## Setup

1. Create a virtual environment (recommended):

   ```bash
   python3 -m venv .venv
   source .venv/bin/activate   # Windows: .venv\Scripts\activate
   ```

2. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

## Configuration (after you have the PDF)

Edit `result_extractor/config.py`:

- **`OUTPUT_COLUMN_ORDER`**  
  Set the desired column order in the output Excel file. Use either:
  - **Header names**: exact strings as they appear in the PDF table (e.g. `["Name", "Date", "Result"]`), or  
  - **0-based indices**: e.g. `[0, 2, 1]` to put column 0 first, then 2, then 1.

- **`HAS_HEADER_ROW`**  
  Set to `True` if the first row of the table is a header; set to `False` if there is no header (then use indices in `OUTPUT_COLUMN_ORDER`).

- **`OUTPUT_SHEET_NAME`**  
  Name of the sheet in the generated workbook (optional).

## Usage

### GUI (recommended)

```bash
python gui.py
```

A window opens with an **Import and Convert** button. Click it to choose a PDF (only PDF files are shown in the file picker). The file dialog opens in the **pdf-to-convert** folder if it exists—place your PDF there for convenience. The spreadsheet (header data + table) is saved on your **Desktop** as **exported_&lt;timestamp&gt;.xlsx** (e.g. `exported_2025-03-09_14-30-00.xlsx`). Works on macOS and Windows.

### Command line

```bash
# Output file will be input.pdf → input.xlsx (same directory)
python main.py path/to/input.pdf

# Specify output path
python main.py path/to/input.pdf -o path/to/output.xlsx

# Use all tables in the PDF (default: first table only)
python main.py input.pdf --all-tables
```

## Building executables (Windows and macOS)

You can build standalone executables so others can run the tool without installing Python.

1. Install PyInstaller:

   ```bash
   pip install pyinstaller
   ```

2. Build on **macOS** (creates a Mac executable; use `gui.py` for the windowed app):

   ```bash
   pyinstaller --onefile --name result-extractor-mac gui.py
   ```

   The executable will be in `dist/result-extractor-mac`.

3. Build on **Windows** (creates a Windows executable; use `gui.py` for the windowed app):

   ```cmd
   pyinstaller --onefile --name result-extractor-win gui.py
   ```

   The executable will be in `dist\result-extractor-win.exe`.

**Note:** Build the Windows executable on a Windows machine and the Mac executable on a Mac (or use CI). PyInstaller produces an executable for the OS it runs on.

## Project layout

```
result-extractor/
├── main.py                 # CLI entry point (table-only conversion)
├── gui.py                  # GUI: Import and Convert → header + table → exported_<timestamp>.xlsx on Desktop
├── pdf-to-convert/         # Put PDFs here; file dialog opens in this folder
├── requirements.txt
├── README.md
├── result_extractor/
│   ├── __init__.py
│   ├── config.py           # Column order and options (edit when you have the PDF)
│   ├── pdf_reader.py       # PDF header + table extraction (pdfplumber)
│   ├── excel_writer.py     # XLSX: header section (Field, Value) + table
│   └── converter.py        # convert_pdf_to_spreadsheet() for full conversion
└── .gitignore
```

## When you have the PDF

1. Run once with the current (empty) column order to see the extracted table:  
   `python main.py yourfile.pdf -o preview.xlsx`  
   Open `preview.xlsx` to see the headers and column order.

2. Update `OUTPUT_COLUMN_ORDER` in `result_extractor/config.py` with the desired order (header names or indices).

3. Run again to generate the final XLSX with columns in the right order.
