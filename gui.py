#!/usr/bin/env python3
"""
Result Extractor GUI: select a PDF and convert to XLSX on the Desktop.
Modern UI with CustomTkinter (Result Cobranças visual identity).
"""

import sys
from pathlib import Path

# Allow running as script from repo root
sys.path.insert(0, str(Path(__file__).resolve().parent))

import customtkinter as ctk
from tkinter import filedialog, messagebox

from result_extractor.converter import (
    convert_pdf_to_spreadsheet,
    convert_pdf_to_spreadsheet_imobiliarias,
)

# --- Visual identity (Result Cobranças: dark teal/petrol + orange) ---
COLOR_TEAL = "#0a4d4e"       # Primary: dark teal / petrol blue (headers)
COLOR_ORANGE = "#e8762e"     # Accent: primary action buttons
COLOR_ORANGE_HOVER = "#f08a45"
COLOR_BG = "#f5f5f5"         # Light background
COLOR_CARD = "#ffffff"       # Card/panel background
FONT_FAMILY = "Segoe UI"     # Modern sans-serif (Windows native)
FONT_SIZE_TITLE = 18
FONT_SIZE_BUTTON = 14
PADDING_X = 56
PADDING_Y = 40
BUTTON_WIDTH = 380
BUTTON_HEIGHT = 48
CORNER_RADIUS = 12


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
    """IMOBILIÁRIAS flow: select PDF, exclude TÍTULO/ATRASO, same extraction pattern as condomínios."""
    pdf_dir = get_pdf_to_convert_dir()
    initial_dir = str(pdf_dir) if pdf_dir.is_dir() else None
    pdf_path = filedialog.askopenfilename(
        title="Select PDF file (Imobiliárias)",
        initialdir=initial_dir,
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
    )
    if not pdf_path:
        return

    desktop = get_desktop_path()
    desktop.mkdir(parents=True, exist_ok=True)

    try:
        out_path = convert_pdf_to_spreadsheet_imobiliarias(pdf_path, output_dir=desktop)
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


def main() -> None:
    # High DPI: CustomTkinter enables DPI awareness on Windows by default
    ctk.set_appearance_mode("Light")
    ctk.set_default_color_theme("blue")  # base theme; we override button colors

    root = ctk.CTk()
    root.title("Result Extractor")
    root.geometry("520x380")
    root.resizable(True, True)
    root.configure(fg_color=COLOR_BG)

    # Minimalist layout: centered content with plenty of whitespace
    # CTkFrame does not support padding; we use pack padx/pady for inset
    main_container = ctk.CTkFrame(
        root,
        fg_color=COLOR_CARD,
        corner_radius=16,
        border_width=0,
    )
    main_container.place(relx=0.5, rely=0.5, anchor="center")

    # Header (dark teal)
    header = ctk.CTkLabel(
        main_container,
        text="Result Extractor",
        font=ctk.CTkFont(family=FONT_FAMILY, size=FONT_SIZE_TITLE, weight="bold"),
        text_color="white",
        fg_color=COLOR_TEAL,
        corner_radius=10,
        height=44,
    )
    header.pack(fill="x", padx=PADDING_X, pady=(PADDING_Y, 28))

    # Buttons: large, centered, orange primary style
    btn_condominios = ctk.CTkButton(
        main_container,
        text="CONDOMÍNIOS – Importar e Converter",
        command=import_and_convert,
        font=ctk.CTkFont(family=FONT_FAMILY, size=FONT_SIZE_BUTTON),
        width=BUTTON_WIDTH,
        height=BUTTON_HEIGHT,
        corner_radius=CORNER_RADIUS,
        fg_color=COLOR_ORANGE,
        hover_color=COLOR_ORANGE_HOVER,
        text_color="white",
        cursor="hand2",
    )
    btn_condominios.pack(padx=PADDING_X, pady=(0, 16))

    btn_imobiliarias = ctk.CTkButton(
        main_container,
        text="IMOBILIÁRIAS – Importar e converter",
        command=imobiliarias_import_and_convert,
        font=ctk.CTkFont(family=FONT_FAMILY, size=FONT_SIZE_BUTTON),
        width=BUTTON_WIDTH,
        height=BUTTON_HEIGHT,
        corner_radius=CORNER_RADIUS,
        fg_color=COLOR_ORANGE,
        hover_color=COLOR_ORANGE_HOVER,
        text_color="white",
        cursor="hand2",
    )
    btn_imobiliarias.pack(padx=PADDING_X, pady=(0, PADDING_Y))

    root.mainloop()


if __name__ == "__main__":
    main()
