# AGENTS.md

## Project Overview

Quadstate Renewal Processor is a Windows desktop app for processing renewal Excel exports.

The app uses:
- `CustomTkinter` for the GUI
- `pandas` / `openpyxl` for unprotected Excel files
- `msoffcrypto-tool` for password-protected Excel files
- `xlsxwriter` for formatted Excel output
- `python-dotenv` for password/environment configuration
- `pywin32` retained for Excel COM helper support

## Entry Point

Keep `QuadstateRenewals.py` as the small launcher entry point:

```python
from quadstate_renewals.app import main
```

Do not move main startup logic back into this file. This file should remain PyInstaller-friendly.

## Package Layout

- `quadstate_renewals/app.py`: CustomTkinter UI, file dialogs, status colors, progress bar, button callbacks
- `quadstate_renewals/processor.py`: Coordinates reading, validation, transformation, export, status/progress callbacks
- `quadstate_renewals/excel_reader.py`: Reads Excel files; uses `msoffcrypto` for protected files
- `quadstate_renewals/excel_writer.py`: Writes formatted output workbook
- `quadstate_renewals/constants.py`: Required columns, output columns, dropdown values, row colors
- `quadstate_renewals/config.py`: `.env` loading, password lookup, default path resolution
- `quadstate_renewals/logging_config.py`: Console/file logging setup
- `quadstate_renewals/com_utils.py`: Retained Excel COM helper; not the active read path

## Important Behavior

Required input columns live in `quadstate_renewals/constants.py`.

If the carrier/source Excel format changes, update:
- `REQUIRED_COLUMNS`
- `OUTPUT_COLUMNS`
- any transformation logic in `processor.py`
- any formatting assumptions in `excel_writer.py`

The UI shows format-change warnings in yellow when required columns are missing. The short UI message is intentionally concise; full missing-column details are logged to `app.log`.

The record tracker distinguishes:
- data records: `len(df)`
- visible Excel rows: `len(df) + 1`, because row 1 is the header

## GUI Notes

Do not call `python QuadstateRenewals.py` in automated checks unless you intend to open the GUI. It enters the CustomTkinter mainloop.

For non-GUI verification, prefer import/syntax checks.

Status coloring is handled in `QuadstateRenewalsApp.set_status()`:
- warnings and source-format errors are yellow
- normal processing/success messages return to the default text color

Keep the Process button visible when changing layout. If adding vertical content, increase the window height or compact existing rows.

## Verification Commands

Use these lightweight checks after edits:

```powershell
python -c "import ast, pathlib; paths=[pathlib.Path('QuadstateRenewals.py'), *pathlib.Path('quadstate_renewals').glob('*.py')]; [ast.parse(p.read_text()) for p in paths]; print('syntax ok:', len(paths), 'files')"
```

```powershell
python -c "import QuadstateRenewals; from quadstate_renewals import app, config, excel_reader, excel_writer, processor; print('imports ok')"
```

Avoid relying only on `python -m py_compile` if local `__pycache__` permissions are problematic.

## Packaging Notes

Keep imports normal and explicit. Avoid dynamic imports unless there is a strong reason.

PyInstaller should target `QuadstateRenewals.py`.

CustomTkinter may require explicit data-file handling in PyInstaller packaging. Be careful when changing CustomTkinter themes, fonts, or assets.

Inno Setup should package the PyInstaller output folder rather than individual source files.

## README Encoding

`README.md` is UTF-16LE. Preserve its encoding when editing.

In PowerShell, read/write it with:

```powershell
Get-Content -Path .\README.md -Encoding Unicode
Set-Content -Path .\README.md -Encoding Unicode
```

Avoid tools that assume UTF-8 unless intentionally converting the file.

## Editing Guidance

Keep changes scoped:
- UI changes belong in `app.py`
- processing workflow changes belong in `processor.py`
- Excel input/decryption changes belong in `excel_reader.py`
- output workbook formatting belongs in `excel_writer.py`
- column/dropdown/color constants belong in `constants.py`

Do not mix GUI layout changes with Excel processing changes unless the task truly requires both.
