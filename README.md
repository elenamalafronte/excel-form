# Excel Form App

A desktop app built with CustomTkinter to manage "Heat Number" records in Excel.

The app provides:
- an Insert tab for adding new rows
- a Search tab for searching, reviewing, and updating existing rows
- a customizable field configuration UI

## Security

- The Customize Fields action is password-protected.
- local credentials are in `secrets.py` 
- `secrets.py` should stay in `.gitignore`.

## Features

- Insert form for row creation
  - Auto-generates `File Number`
  - ItemCode-based Description autofill
  - File picker support for `FileLink`
  - Success message includes the physical worksheet row where the record was saved
- Search table
  - Search by any configured column
  - Column visibility controls
  - Horizontal and vertical scrolling
  - Open workbook and open file links from table
- Retroactive PDF upload from Search
  - Select a row and upload/replace `FileLink`
  - Saves to workbook and refreshes table
- Customizable fields panel
  - Add/remove fields
  - Drag-and-drop reorder fields
  - Undo remove
  - Password prompt required before opening the panel
  - Persist field config into `config.py`
  - Sync workbook schema to updated field set
- Save feedback UX
  - Save buttons show `Saving...` with active visual state while processing

## Project Structure

- `main.py` - app bootstrap and tab mounting
- `insert_tab.py` - insert form and customize-fields UI
- `search_tab.py` - search table and row actions
- `excel.py` - workbook read/write logic and sync helpers
- `config.py` - column schema, validation, and formula template
- `ui_style.py` - shared UI constants

## Requirements

- Python 3.10+ (tested on Python 3.13)
- Packages:
  - `customtkinter`
  - `openpyxl`

Install dependencies:

```bash
pip install customtkinter openpyxl
```

## Running the App

From the project folder:

```bash
python main.py
```

## Workbook Notes

The app expects an Excel workbook with:
- source sheet: `CREXPD01`
- form/output sheet: `Heat Number`

The app uses `EXCEL_FILE` in `config.py` to decide which file to read/write.

Data placement behavior:
- App-managed records are written starting from row 3 in the form/output sheet.
- Search reads from the app-managed table region that starts at row 3.
- This avoids mixing app rows with deep template/history ranges in large client workbooks.

## Common Troubleshooting

- "Cannot save workbook" errors:
  - Close the workbook in Excel and try again.
- Description not appearing immediately:
  - Use Refresh in Search tab, or reopen the workbook if external recalculation is needed.
- Rows appear in Search but not where expected in Excel:
  - App rows are written starting at row 3 by design.
- Field customization issues:
  - Ensure `ItemCode` exists if `Description` is enabled, otherwise formula-based autofill cannot work.
