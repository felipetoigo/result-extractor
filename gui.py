#!/usr/bin/env python3
"""
Result Extractor GUI: select a PDF and convert to XLSX on the Desktop.
"""

import sys
from pathlib import Path

# Allow running as script from repo root
sys.path.insert(0, str(Path(__file__).resolve().parent))

import tkinter as tk
from tkinter import filedialog, messagebox

from result_extractor.converter import convert_pdf_to_spreadsheet


def get_desktop_path() -> Path:
    """Return the user's Desktop folder (Mac and Windows)."""
    return Path.home() / "Desktop"


def get_pdf_to_convert_dir() -> Path:
    """Return the pdf-to-convert folder next to the app."""
    return Path(__file__).resolve().parent / "pdf-to-convert"


def import_and_convert() -> None:
    pdf_dir = get_pdf_to_convert_dir()
    initial_dir = str(pdf_dir) if pdf_dir.is_dir() else None
    pdf_path = filedialog.askopenfilename(
        title="Select PDF file",
        initialdir=initial_dir,
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
    )
    if not pdf_path:
        return

    desktop = get_desktop_path()
    desktop.mkdir(parents=True, exist_ok=True)

    try:
        out_path = convert_pdf_to_spreadsheet(pdf_path, output_dir=desktop)
        messagebox.showinfo(
            "Done",
            f"File saved to:\n{out_path}",
        )
    except FileNotFoundError as e:
        messagebox.showerror("Error", str(e))
    except ValueError as e:
        messagebox.showerror("Error", str(e))
    except Exception as e:
        messagebox.showerror("Error", str(e))


def imobiliarias_import_and_convert() -> None:
    """IMOBILIÁRIAS flow: import PDF and convert (rules to be defined)."""
    # TODO: define and implement IMOBILIÁRIAS-specific rules
    pass


def main() -> None:
    root = tk.Tk()
    root.title("Result Extractor")
    root.geometry("480x260")
    root.resizable(True, True)

    frame = tk.Frame(root, padx=48, pady=48)
    frame.pack()

    btn_condominios = tk.Button(
        frame,
        text="CONDOMÍNIOS - Importar e Converter",
        command=import_and_convert,
        font=("", 12),
        width=38,
        padx=20,
        pady=10,
        cursor="hand2",
    )
    btn_condominios.pack(pady=(0, 8))

    btn_imobiliarias = tk.Button(
        frame,
        text="IMOBILIÁRIAS - Importar e converter",
        command=imobiliarias_import_and_convert,
        font=("", 12),
        width=38,
        padx=20,
        pady=10,
        cursor="hand2",
    )
    btn_imobiliarias.pack()

    root.mainloop()


if __name__ == "__main__":
    main()
