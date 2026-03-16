# Building the Windows .exe

Build the executable **on a Windows machine** (PyInstaller produces an exe for the OS it runs on).

## 1. Prerequisites

- **Python 3.10+** installed and on PATH (or use the same Python you use for development).
- **Windows**: Build must be done on Windows to produce a Windows `.exe`.

## 2. Install dependencies

From the project root (same folder as `gui.py` and `result_extractor.spec`):

```bash
pip install -r requirements.txt
```

This installs PDF/Excel libraries, **customtkinter** (GUI), and PyInstaller. If `pyinstaller` is not on your PATH after install, use step 3b below.

## 3. Build the .exe

From the project root:

```bash
pyinstaller result_extractor.spec
```

If the `pyinstaller` command is not found (Scripts folder not on PATH), run:

```bash
python -m PyInstaller result_extractor.spec
```

## 4. Output

- The `.exe` is created in **`dist/Result Extractor.exe`**.
- You can copy that single file to any Windows PC and run it (no Python needed).
- The first run may be a bit slow while it unpacks; after that it runs normally.

## 5. Notes

- **GUI**: The app uses **CustomTkinter**. The spec includes CustomTkinter’s data files (themes, fonts) so the window and buttons look correct. If the window appears blank or styles are wrong, ensure you have the latest `result_extractor.spec` (with `collect_data_files('customtkinter')` in `datas`).
- **Antivirus**: Some antivirus software may flag new PyInstaller exes. If that happens, add an exception or build on the target machine.
- **One-file**: The spec is set up for a single executable (no separate folder of DLLs).
- **Console**: The app is built with `console=False`, so no command-prompt window appears when you run it.

## 6. Testing without building

To run the app from source (e.g. on Mac or Windows) without building an exe:

```bash
pip install -r requirements.txt
python gui.py
```
