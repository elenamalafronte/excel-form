import zipfile
import re
from pathlib import Path

from openpyxl import Workbook, load_workbook

from config import (
    COLUMNS,
    EXCEL_FILE,
    FILE_NUMBER_PATTERN,
    SEARCH_BY,
)


def _open_workbook():
    path = Path(EXCEL_FILE)
    if path.exists():
        return load_workbook(path, keep_vba=path.suffix.lower() == ".xlsm")
    return Workbook()


def _get_layout_sheets(wb):
    ws_source = None
    ws_form = None

    if "CREXPD01" in wb.sheetnames:
        ws_source = wb["CREXPD01"]
    elif len(wb.worksheets) > 0:
        ws_source = wb.worksheets[0]

    if "Heat Number" in wb.sheetnames:
        ws_form = wb["Heat Number"]
    elif len(wb.worksheets) > 1:
        ws_form = wb.worksheets[1]

    if ws_source is None:
        ws_source = wb.active
        ws_source.title = "CREXPD01"
    if ws_form is None:
        ws_form = wb.create_sheet(title="Heat Number")

    return ws_source, ws_form


def _get_form_sheet_for_read(wb):
    if "Heat Number" in wb.sheetnames:
        return wb["Heat Number"]
    if len(wb.worksheets) < 2:
        return None
    return wb.worksheets[1]


def _ensure_workbook_and_sheets():
    wb = _open_workbook()
    ws_source, ws_form = _get_layout_sheets(wb)

    headers = [c["name"] for c in COLUMNS]
    if ws_form.max_row == 1 and ws_form.cell(row=1, column=1).value is None:
        ws_form.append(headers)
    else:
        for idx, header in enumerate(headers, start=1):
            if ws_form.cell(row=1, column=idx).value != header:
                ws_form.cell(row=1, column=idx, value=header)

    return wb, ws_source, ws_form


# Column O (1-based = 15, 0-based = 14) is the ItemCode column in CREXPD01,
# as confirmed by the formula: MATCH(B{row}, CREXPD01!$O$1:$O$15000, 0)
_CREXPD01_ITEMCODE_COL_FALLBACK = 14  # 0-based index


def _build_description_index(wb):
    """Return {ItemCode.upper(): description} from the CREXPD01 sheet.

    Locates ItemCode and Detailed Description columns by header name
    (case- and whitespace-insensitive), falling back to column O for ItemCode
    when the header scan finds no match.
    """
    ws_source = None
    if "CREXPD01" in wb.sheetnames:
        ws_source = wb["CREXPD01"]
    elif wb.worksheets:
        ws_source = wb.worksheets[0]

    if ws_source is None:
        return {}

    item_code_col = None
    description_col = None

    header_row = next(ws_source.iter_rows(min_row=1, max_row=1), None)
    if header_row:
        for col_idx, cell in enumerate(header_row):
            val = str(cell.value or "").strip().lower().replace(" ", "").replace("_", "")
            if "itemcode" in val:
                item_code_col = col_idx
            if "detaileddescription" in val or val == "description":
                description_col = col_idx

    if item_code_col is None:
        item_code_col = _CREXPD01_ITEMCODE_COL_FALLBACK

    if description_col is None:
        return {}

    index = {}
    for row in ws_source.iter_rows(min_row=2):
        try:
            ic_val = row[item_code_col].value
            desc_val = row[description_col].value
        except IndexError:
            continue
        if ic_val:
            index[str(ic_val).strip().upper()] = desc_val or ""
    return index


def load_sheet():
    """Return worksheet rows as list[dict], excluding the header row.

    Description cells whose cached value is None (Excel formula not yet
    recalculated after an openpyxl save) are back-filled from a Python-side
    lookup against CREXPD01.  The desc_index is built lazily — only on the
    first row that actually needs it — so routine searches never pay the cost
    of scanning all 15 000 source rows.
    """
    path = Path(EXCEL_FILE)
    if not path.exists():
        return []

    wb = load_workbook(path, read_only=True, data_only=True)
    ws = _get_form_sheet_for_read(wb)
    if ws is None:
        wb.close()
        return []

    headers = [c["name"] for c in COLUMNS]
    file_link_col_index = next(
        (idx for idx, col in enumerate(COLUMNS) if col["name"] == "FileLink"), None
    )

    desc_index = None  # built lazily only if a row has a missing description

    rows = []
    consecutive_empty = 0
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row_dict = {}
        for i in range(len(headers)):
            cell = row[i] if i < len(row) else None
            if cell is None:
                row_dict[headers[i]] = None
                continue
            hyperlink = getattr(cell, "hyperlink", None)
            if file_link_col_index is not None and i == file_link_col_index and hyperlink:
                row_dict[headers[i]] = getattr(hyperlink, "target", None) or cell.value
            else:
                row_dict[headers[i]] = cell.value

        if any(v not in (None, "") for v in row_dict.values()):
            consecutive_empty = 0
            if not row_dict.get("Description") and row_dict.get("ItemCode"):
                if desc_index is None:
                    desc_index = _build_description_index(wb)
                ic = str(row_dict["ItemCode"]).strip().upper()
                row_dict["Description"] = desc_index.get(ic, "")
            rows.append(row_dict)
        else:
            consecutive_empty += 1
            if consecutive_empty > 10:
                break

    wb.close()
    return rows


def append_row(data: dict):
    """Append one row to the Heat Number sheet and save.

    Strategy to avoid the inflated-max_row bug:
      Phase 1 — read-only streaming pass to build the description index and
                 find the true last data row (Excel's stored max_row metadata
                 is inflated by formatting and cannot be trusted for ws.append).
      Phase 2 — write-mode: write directly to last_data_row + 1 with no
                 backward scan, then save.
    """
    file_path = Path(EXCEL_FILE)
    if not file_path.exists():
        raise FileNotFoundError(f"Workbook not found: {EXCEL_FILE}")

    # Phase 1 – fast read-only pass
    desc_index = {}
    last_data_row = 1
    try:
        wb_ro = load_workbook(file_path, read_only=True, data_only=True)
        desc_index = _build_description_index(wb_ro)
        ws_ro = _get_form_sheet_for_read(wb_ro)
        if ws_ro is not None:
            for row in ws_ro.iter_rows(min_row=2):
                if any(cell.value is not None for cell in row):
                    last_data_row = row[0].row
        wb_ro.close()
    except Exception:
        pass

    # Phase 2 – write-mode pass (no scanning, write at exact row)
    wb, _ws_source, ws = _ensure_workbook_and_sheets()

    row_values = [data.get(col["name"], "") for col in COLUMNS]

    item_code = str(data.get("ItemCode", "") or "").strip().upper()
    desc_col_idx = next(
        (i for i, c in enumerate(COLUMNS) if c["name"] == "Description"), None
    )
    if desc_col_idx is not None:
        row_values[desc_col_idx] = desc_index.get(item_code, "")

    new_row_idx = last_data_row + 1
    for col_idx, value in enumerate(row_values, start=1):
        ws.cell(row=new_row_idx, column=col_idx, value=value)

    for col_idx, col in enumerate(COLUMNS, start=1):
        if col["name"] != "FileLink":
            continue
        link_value = data.get("FileLink", "")
        if link_value:
            cell = ws.cell(row=new_row_idx, column=col_idx)
            cell.value = link_value
            cell.hyperlink = str(link_value)
            cell.style = "Hyperlink"

    try:
        wb.save(str(file_path))
    except PermissionError as exc:
        raise PermissionError(
            f"Cannot save workbook. Close '{EXCEL_FILE}' in Excel and try again."
        ) from exc
    finally:
        try:
            wb.close()
        except Exception:
            pass


def search_rows(search_value, search_column="ItemCode"):
    rows = load_sheet()
    if not search_value:
        return rows

    search_value = str(search_value).strip().lower()
    column = search_column if search_column in SEARCH_BY else "ItemCode"

    results = []
    for row in rows:
        candidate = str(row.get(column, "")).strip().lower()
        if search_value in candidate:
            results.append(row)
    return results
