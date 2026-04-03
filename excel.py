import zipfile
import re
from pathlib import Path

from openpyxl import Workbook, load_workbook

from config import (
    COLUMNS,
    EXCEL_FILE,
    FILE_NUMBER_PATTERN,
    SEARCH_BY,
    build_description_formula,
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


def _col_idx_to_letter(col_idx):
    """Convert 0-based column index to Excel column letter (A, B, ..., AA, ...)."""
    letters = []
    n = col_idx + 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters.append(chr(65 + rem))
    return "".join(reversed(letters))


def _find_sheet_zip_path(file_path, sheet_name):
    """Return the ZIP-internal path to a named worksheet's XML file."""
    with zipfile.ZipFile(file_path, "r") as z:
        wb_xml = z.read("xl/workbook.xml").decode("utf-8")
        rels_xml = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")

    id_match = re.search(
        rf'<sheet\b[^>]*\bname="{re.escape(sheet_name)}"[^>]*\br:id="([^"]+)"', wb_xml
    ) or re.search(
        rf'<sheet\b[^>]*\br:id="([^"]+)"[^>]*\bname="{re.escape(sheet_name)}"', wb_xml
    )
    if not id_match:
        raise ValueError(f"Sheet '{sheet_name}' not found in workbook.xml")
    r_id = id_match.group(1)

    target_match = re.search(
        rf'<Relationship\b[^>]*\bId="{re.escape(r_id)}"[^>]*\bTarget="([^"]+)"', rels_xml
    ) or re.search(
        rf'<Relationship\b[^>]*\bTarget="([^"]+)"[^>]*\bId="{re.escape(r_id)}"', rels_xml
    )
    if not target_match:
        raise ValueError(f"Could not resolve relationship {r_id}")
    target = target_match.group(1)
    return f"xl/{target}" if not target.startswith("/") else target


def _build_row_xml(row_idx, row_values):
    """Return the XML string for one worksheet row, using inline strings."""
    cells = []
    for col_idx, value in enumerate(row_values):
        if value is None or value == "":
            continue
        col_letter = _col_idx_to_letter(col_idx)
        cell_ref = f"{col_letter}{row_idx}"
        val = str(value)
        if val.startswith("="):
            formula = val[1:].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            cells.append(f'<c r="{cell_ref}"><f>{formula}</f></c>')
        else:
            try:
                num = float(val)
                stored = int(num) if num == int(num) and "." not in val else num
                cells.append(f'<c r="{cell_ref}"><v>{stored}</v></c>')
            except ValueError:
                safe = val.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                cells.append(f'<c r="{cell_ref}" t="inlineStr"><is><t>{safe}</t></is></c>')
    return f'<row r="{row_idx}">{"".join(cells)}</row>'


def _zip_append_row(file_path, row_values, new_row_idx):
    """Append one row by rewriting only the Heat Number worksheet XML in the ZIP.

    CREXPD01's bytes are copied verbatim — it is never parsed or serialized.
    This eliminates the save latency that came from openpyxl re-serializing
    15 000+ rows of source data on every insert.
    """
    sheet_zip_path = _find_sheet_zip_path(file_path, "Heat Number")
    new_row_xml = _build_row_xml(new_row_idx, row_values)

    with zipfile.ZipFile(file_path, "r") as z:
        sheet_str = z.read(sheet_zip_path).decode("utf-8")

    if "</sheetData>" not in sheet_str:
        raise ValueError(f"</sheetData> not found in {sheet_zip_path}")

    sheet_str = sheet_str.replace("</sheetData>", f"{new_row_xml}</sheetData>", 1)
    sheet_str = re.sub(
        r'(<dimension ref="[A-Z]+\d+:)([A-Z]+)(\d+)(")',
        lambda m: f"{m.group(1)}{m.group(2)}{new_row_idx}{m.group(4)}",
        sheet_str,
    )

    tmp_path = file_path.with_suffix(".xlsx.tmp")
    try:
        with zipfile.ZipFile(file_path, "r") as z_in, \
             zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as z_out:
            for item in z_in.infolist():
                data = sheet_str.encode("utf-8") if item.filename == sheet_zip_path \
                    else z_in.read(item.filename)
                z_out.writestr(item, data)
        tmp_path.replace(file_path)
    except Exception:
        tmp_path.unlink(missing_ok=True)
        raise


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

    # Phase 2 – ZIP surgery: rewrite only the Heat Number sheet XML
    new_row_idx = last_data_row + 1

    row_values = [data.get(col["name"], "") for col in COLUMNS]

    desc_col_idx = next(
        (i for i, c in enumerate(COLUMNS) if c["name"] == "Description"), None
    )
    item_code_col_idx = next(
        (i for i, c in enumerate(COLUMNS) if c["name"] == "ItemCode"), None
    )

    if desc_col_idx is not None and item_code_col_idx is not None:
        col_letter = chr(ord("A") + item_code_col_idx)
        row_values[desc_col_idx] = build_description_formula(
            f"{col_letter}{new_row_idx}", "CREXPD01"
        )

    try:
        _zip_append_row(file_path, row_values, new_row_idx)
    except PermissionError as exc:
        raise PermissionError(
            f"Cannot save workbook. Close '{EXCEL_FILE}' in Excel and try again."
        ) from exc


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
