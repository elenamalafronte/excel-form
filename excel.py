import os
import subprocess
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


def _build_row_xml(row_idx, row_values, cached_values=None):
    """Return the XML string for one worksheet row, using inline strings.

    cached_values: optional {col_idx: value} — for formula cells, the pre-computed
    result is written as <v> so openpyxl data_only=True can read it immediately
    without waiting for Excel to recalculate.
    """
    cells = []
    cached_values = cached_values or {}
    for col_idx, value in enumerate(row_values):
        if value is None or value == "":
            continue
        col_letter = _col_idx_to_letter(col_idx)
        cell_ref = f"{col_letter}{row_idx}"
        val = str(value)
        if val.startswith("="):
            formula = val[1:].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            cached = cached_values.get(col_idx)
            if cached is not None and cached != "":
                safe_v = str(cached).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                cells.append(f'<c r="{cell_ref}" t="str"><f>{formula}</f><v>{safe_v}</v></c>')
            else:
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


def _zip_append_row(file_path, row_values, new_row_idx, cached_values=None):
    """Append one row by rewriting only the Heat Number worksheet XML in the ZIP.

    CREXPD01's bytes are copied verbatim — it is never parsed or serialized.
    This eliminates the save latency that came from openpyxl re-serializing
    15 000+ rows of source data on every insert.
    """
    sheet_zip_path = _find_sheet_zip_path(file_path, "Heat Number")
    new_row_xml = _build_row_xml(new_row_idx, row_values, cached_values)

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


def _col_str_to_idx(col_str: str) -> int:
    """Convert a column letter string like 'A' or 'AV' to a 0-based index."""
    n = 0
    for ch in col_str:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def _build_description_index(file_path: Path) -> dict:
    """Return {ItemCode.upper(): description} by parsing CREXPD01 XML directly.

    Reads sharedStrings.xml and the CREXPD01 worksheet XML straight from the
    ZIP archive, bypassing openpyxl's read-only streaming layer entirely.
    This avoids the known issue where openpyxl read_only=True returns None
    for shared-string cells when the string table hasn't been fully loaded.
    """
    try:
        crexpd01_zip_path = _find_sheet_zip_path(file_path, "CREXPD01")
    except (ValueError, KeyError):
        return {}

    try:
        with zipfile.ZipFile(file_path, "r") as z:
            # Build the shared-strings lookup (most string cells use t="s")
            shared_strings: list = []
            if "xl/sharedStrings.xml" in z.namelist():
                ss_xml = z.read("xl/sharedStrings.xml").decode("utf-8")
                for si in re.findall(r"<si>(.*?)</si>", ss_xml, re.DOTALL):
                    parts = re.findall(r"<t(?:[^>]*)?>([^<]*)</t>", si)
                    raw = "".join(parts)
                    # Unescape XML entities so item codes / descriptions look right
                    raw = (raw.replace("&amp;", "&").replace("&lt;", "<")
                               .replace("&gt;", ">").replace("&quot;", '"')
                               .replace("&apos;", "'"))
                    shared_strings.append(raw)

            sheet_xml = z.read(crexpd01_zip_path).decode("utf-8")
    except Exception:
        return {}

    def _cell_value(t: str, inner: str) -> str:
        if t == "s":
            m = re.search(r"<v>(\d+)</v>", inner)
            if m:
                idx = int(m.group(1))
                return shared_strings[idx] if idx < len(shared_strings) else ""
        elif t == "inlineStr":
            m = re.search(r"<t[^>]*>([^<]*)</t>", inner)
            return m.group(1) if m else ""
        else:
            m = re.search(r"<v>([^<]*)</v>", inner)
            return m.group(1) if m else ""
        return ""

    item_code_col: int = _CREXPD01_ITEMCODE_COL_FALLBACK
    description_col: "int | None" = None
    index: dict = {}

    for row_match in re.finditer(
        r'<row\b[^>]*\br="(\d+)"[^>]*>(.*?)</row>', sheet_xml, re.DOTALL
    ):
        row_num = int(row_match.group(1))
        row_xml = row_match.group(2)

        cells: dict = {}
        for cell_match in re.finditer(r"<c\b([^>]*)>(.*?)</c>", row_xml, re.DOTALL):
            attrs = cell_match.group(1)
            inner = cell_match.group(2)
            r_m = re.search(r'\br="([A-Z]+)', attrs)
            if not r_m:
                continue
            col_idx = _col_str_to_idx(r_m.group(1))
            t_m = re.search(r'\bt="([^"]*)"', attrs)
            cells[col_idx] = _cell_value(t_m.group(1) if t_m else "", inner)

        if row_num == 1:
            for col_idx, val in cells.items():
                normalized = val.strip().lower().replace(" ", "").replace("_", "")
                if "itemcode" in normalized:
                    item_code_col = col_idx
                if "detaileddescription" in normalized:
                    description_col = col_idx
            if description_col is None:
                return {}
        elif description_col is not None:
            ic = cells.get(item_code_col, "").strip()
            desc = cells.get(description_col, "").strip()
            if ic:
                index[ic.upper()] = desc

    return index


# Module-level description index cache — rebuilt only when the file changes.
# Avoids re-scanning 15 000 CREXPD01 rows on every search click.
_desc_index_cache: dict = {}
_desc_index_mtime: float | None = None


def _get_desc_index(path: Path) -> dict:
    """Return a cached {ItemCode.upper(): description} dict for the given file.

    The cache is invalidated whenever the file's mtime changes, so it stays
    fresh after every insert without rescanning on every search.
    """
    global _desc_index_cache, _desc_index_mtime
    try:
        mtime = path.stat().st_mtime
    except OSError:
        return {}
    if mtime != _desc_index_mtime:
        _desc_index_cache = _build_description_index(path)
        _desc_index_mtime = mtime
    return _desc_index_cache


def load_sheet():
    """Return worksheet rows as list[dict], excluding the header row.

    Description is resolved from a Python-side lookup against CREXPD01 so
    that the Search tab always shows the right value, independent of whether
    Excel has recalculated and cached the formula result.  The index is built
    once and cached by file mtime, so repeated searches are instant.
    """
    path = Path(EXCEL_FILE)
    if not path.exists():
        return []

    # Build (or reuse) the description index BEFORE opening the Heat Number
    # sheet.  Both sheets live in the same ZIP, and opening one sheet's XML
    # stream mid-iteration of another can cause the read-only ZipFile handle
    # to return empty data for the second sheet.
    desc_index = _get_desc_index(path)

    wb = load_workbook(path, read_only=True, data_only=True)
    ws = _get_form_sheet_for_read(wb)
    if ws is None:
        wb.close()
        return []

    headers = [c["name"] for c in COLUMNS]
    file_link_col_index = next(
        (idx for idx, col in enumerate(COLUMNS) if col["name"] == "FileLink"), None
    )

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
            if row_dict.get("ItemCode") and desc_index:
                ic = str(row_dict["ItemCode"]).strip().upper()
                looked_up = desc_index.get(ic, "")
                if looked_up:
                    row_dict["Description"] = looked_up
            rows.append(row_dict)
        else:
            consecutive_empty += 1
            if consecutive_empty > 10:
                break

    wb.close()
    return rows


def _excel_recalc_and_save(file_path: Path) -> None:
    """Open the workbook in Excel via PowerShell COM, recalculate, and save.

    This caches all formula results as <v> elements in the XML so that
    openpyxl data_only=True can read them immediately on the next load.
    Runs silently in the background — errors are suppressed so a missing
    Excel installation doesn't break the insert flow.
    """
    ps = (
        "$xl = New-Object -ComObject Excel.Application;"
        "$xl.Visible = $false; $xl.DisplayAlerts = $false;"
        "$wb = $xl.Workbooks.Open($env:_EXCEL_PATH);"
        "$xl.CalculateFull(); $wb.Save(); $wb.Close($false); $xl.Quit()"
    )
    try:
        subprocess.run(
            ["powershell", "-NonInteractive", "-NoProfile", "-Command", ps],
            env={**os.environ, "_EXCEL_PATH": str(file_path.absolute())},
            timeout=30,
            capture_output=True,
        )
    except Exception:
        pass  # non-fatal: description will update next time Excel saves


def append_row(data: dict):
    """Append one row to the Heat Number sheet via direct ZIP surgery.

    Phase 1 — read-only streaming pass to find the true last data row.
    Phase 2 — rewrite only the Heat Number worksheet XML inside the ZIP,
               leaving CREXPD01 untouched (never parsed, never serialized).
    """
    file_path = Path(EXCEL_FILE)
    if not file_path.exists():
        raise FileNotFoundError(f"Workbook not found: {EXCEL_FILE}")

    # Phase 1 – find true last data row (Heat Number sheet only).
    # desc_index is fetched from the module-level mtime cache; no extra I/O if
    # the file hasn't changed since the last search.
    last_data_row = 1
    try:
        wb_ro = load_workbook(file_path, read_only=True, data_only=True)
        ws_ro = _get_form_sheet_for_read(wb_ro)
        if ws_ro is not None:
            for row in ws_ro.iter_rows(min_row=2):
                if any(cell.value is not None for cell in row):
                    last_data_row = row[0].row
        wb_ro.close()
    except Exception:
        pass

    desc_index = _get_desc_index(file_path)

    # Phase 2 – ZIP surgery: rewrite only the Heat Number sheet XML
    new_row_idx = last_data_row + 1

    row_values = [data.get(col["name"], "") for col in COLUMNS]

    desc_col_idx = next(
        (i for i, c in enumerate(COLUMNS) if c["name"] == "Description"), None
    )
    item_code_col_idx = next(
        (i for i, c in enumerate(COLUMNS) if c["name"] == "ItemCode"), None
    )

    cached_values = {}
    if desc_col_idx is not None and item_code_col_idx is not None:
        col_letter = chr(ord("A") + item_code_col_idx)
        row_values[desc_col_idx] = build_description_formula(
            f"{col_letter}{new_row_idx}", "CREXPD01"
        )
        item_code = data.get("ItemCode", "")
        if item_code and desc_index:
            description = desc_index.get(str(item_code).strip().upper(), "")
            if description:
                cached_values[desc_col_idx] = description

    try:
        _zip_append_row(file_path, row_values, new_row_idx, cached_values)
    except PermissionError as exc:
        raise PermissionError(
            f"Cannot save workbook. Close '{EXCEL_FILE}' in Excel and try again."
        ) from exc

    _excel_recalc_and_save(file_path)


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
