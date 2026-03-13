# Building the Windows .exe

Build the executable **on a Windows machine** (PyInstaller produces an exe for the OS it runs on).

## 1. Install dependencies

```bash
pip install -r requirements.txt
```

## 2. Build the .exe

From the project root (same folder as `gui.py` and `result_extractor.spec`):

```bash
pyinstaller result_extractor.spec
```

## 3. Output

- The `.exe` is created in **`dist/Result Extractor.exe`**.
- You can copy that single file to any Windows PC and run it (no Python needed).
- The first run may be a bit slow while it unpacks; after that it runs normally.

## Notes

- **Antivirus**: Some antivirus software may flag new PyInstaller exes. If that happens, add an exception or build on the target machine.
- **One-file**: The spec is set up for a single executable (no separate folder of DLLs).
- **Console**: The app is built with `console=False`, so no command-prompt window appears when you run it.
