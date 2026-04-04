import os
import subprocess
import zipfile
import re
from pathlib import Path

from openpyxl import Workbook, load_workbook

from config import (
    COLUMNS,
    build_description_formula,
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

    # Normalize relationship target to actual ZIP member format.
    # Some workbooks store absolute-like targets such as /xl/worksheets/sheet2.xml
    # while ZIP members are stored without leading slash.
    normalized = target.replace("\\", "/").lstrip("/")
    if not normalized.startswith("xl/"):
        normalized = f"xl/{normalized}"
    return normalized


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

    if "</sheetData>" in sheet_str:
        sheet_str = sheet_str.replace("</sheetData>", f"{new_row_xml}</sheetData>", 1)
    elif re.search(r"<sheetData\s*/>", sheet_str):
        # Empty sheets can use a self-closing sheetData tag.
        sheet_str = re.sub(
            r"<sheetData\s*/>",
            f"<sheetData>{new_row_xml}</sheetData>",
            sheet_str,
            count=1,
        )
    else:
        raise ValueError(f"sheetData section not found in {sheet_zip_path}")
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

    # If XML parsing yielded nothing (or became incompatible with workbook
    # serialization), fall back to openpyxl-based extraction.
    if not index:
        return _build_description_index_fallback(file_path)

    return index


def _build_description_index_fallback(file_path: Path) -> dict:
    """Fallback description index builder using openpyxl row values."""
    try:
        wb = load_workbook(file_path, read_only=True, data_only=True)
        if "CREXPD01" in wb.sheetnames:
            ws = wb["CREXPD01"]
        elif wb.worksheets:
            ws = wb.worksheets[0]
        else:
            wb.close()
            return {}

        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if not header_row:
            wb.close()
            return {}

        normalized_headers = [
            str(h or "").strip().lower().replace(" ", "").replace("_", "")
            for h in header_row
        ]

        item_idx = -1
        desc_idx = -1
        for idx, name in enumerate(normalized_headers):
            if item_idx < 0 and "itemcode" in name:
                item_idx = idx
            if desc_idx < 0 and "detaileddescription" in name:
                desc_idx = idx

        if item_idx < 0 or desc_idx < 0:
            wb.close()
            return {}

        index = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row:
                continue

            item_code = ""
            description = ""

            if item_idx < len(row) and row[item_idx] is not None:
                item_code = str(row[item_idx]).strip().upper()
            if desc_idx < len(row) and row[desc_idx] is not None:
                description = str(row[desc_idx]).strip()

            if item_code:
                index[item_code] = description

        wb.close()
        return index
    except Exception:
        return {}


# Module-level description index cache — rebuilt only when the file changes.
# Avoids re-scanning 15 000 CREXPD01 rows on every search click.
_desc_index_cache: dict = {}
_desc_index_mtime: float | None = None


def _invalidate_desc_index_cache() -> None:
    global _desc_index_cache, _desc_index_mtime
    _desc_index_cache = {}
    _desc_index_mtime = None


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


def get_description_for_itemcode(item_code: str) -> str:
    """Return description for an ItemCode from CREXPD01 index, if available."""
    path = Path(EXCEL_FILE)
    if not path.exists() or not item_code:
        return ""

    desc_index = _get_desc_index(path)
    if not desc_index:
        return ""

    return desc_index.get(str(item_code).strip().upper(), "") or ""


def load_sheet(_allow_recalc=True):
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

    # If a row has an ItemCode that exists in CREXPD01 but Description is still
    # blank, force one Excel recalc/save pass and reload once.
    if _allow_recalc and desc_index:
        should_recalc = False
        for row in rows:
            item_code = str(row.get("ItemCode", "")).strip().upper()
            description = str(row.get("Description", "") or "").strip()
            if item_code and not description and item_code in desc_index:
                should_recalc = True
                break

        if should_recalc:
            _excel_recalc_and_save(path)
            return load_sheet(_allow_recalc=False)

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


def recalc_workbook() -> None:
    """Force Excel recalculation/save for the configured workbook."""
    file_path = Path(EXCEL_FILE)
    if not file_path.exists():
        return
    _excel_recalc_and_save(file_path)
    _invalidate_desc_index_cache()


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
    last_data_row = 1
    desc_index = {}
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

    # Build description index once for this append. This allows writing a
    # cached formula result so Search can display Description immediately.
    desc_index = _get_desc_index(file_path)

    # Phase 2 – ZIP surgery: rewrite only the Heat Number sheet XML
    new_row_idx = last_data_row + 1

    # Build row values against the actual worksheet header order so formula
    # references remain correct even after drag-and-drop field reordering.
    effective_columns = [col["name"] for col in COLUMNS]
    try:
        wb_headers = load_workbook(file_path, read_only=True, data_only=True)
        ws_headers = _get_form_sheet_for_read(wb_headers)
        if ws_headers is not None:
            header_row = next(ws_headers.iter_rows(min_row=1, max_row=1, values_only=True), None)
            if header_row:
                header_names = [str(v).strip() for v in header_row if v not in (None, "")]
                if header_names:
                    effective_columns = header_names
        wb_headers.close()
    except Exception:
        pass

    row_values = [data.get(name, "") for name in effective_columns]

    desc_col_idx = next(
        (i for i, name in enumerate(effective_columns) if name == "Description"), None
    )
    item_code_col_idx = next(
        (i for i, name in enumerate(effective_columns) if name == "ItemCode"), None
    )

    cached_values = {}
    if desc_col_idx is not None and item_code_col_idx is not None:
        item_code_col_letter = _col_idx_to_letter(item_code_col_idx)
        row_values[desc_col_idx] = build_description_formula(
            f"{item_code_col_letter}{new_row_idx}", "CREXPD01"
        )

        item_code = str(data.get("ItemCode", "")).strip().upper()
        if item_code and desc_index:
            description = desc_index.get(item_code, "")
            if description:
                cached_values[desc_col_idx] = description

    try:
        _zip_append_row(file_path, row_values, new_row_idx, cached_values)
    except PermissionError as exc:
        raise PermissionError(
            f"Cannot save workbook. Close '{EXCEL_FILE}' in Excel and try again."
        ) from exc

    _excel_recalc_and_save(file_path)


def sync_form_sheet_columns(old_columns, new_columns):
    """Sync Heat Number sheet schema to match new columns.

    Existing data is preserved by matching values via old column names.
    Columns removed from config are removed from the sheet.
    Newly added columns are created with blank values for existing rows.
    """
    file_path = Path(EXCEL_FILE)

    wb = _open_workbook()
    _, ws_form = _get_layout_sheets(wb)

    old_names = [c.get("name") for c in old_columns if c.get("name")]
    new_names = [c.get("name") for c in new_columns if c.get("name")]

    if not new_names:
        raise ValueError("Cannot sync workbook: no target columns defined")

    # Prefer real sheet headers when present, otherwise fall back to old config names.
    sheet_headers = []
    if ws_form.max_row >= 1:
        for col_idx in range(1, ws_form.max_column + 1):
            val = ws_form.cell(row=1, column=col_idx).value
            sheet_headers.append(str(val).strip() if val is not None else "")

    effective_old_names = [h for h in sheet_headers if h] or old_names

    # Capture existing rows as dictionaries keyed by old headers.
    existing_row_dicts = []
    if ws_form.max_row >= 2 and effective_old_names:
        for row_idx in range(2, ws_form.max_row + 1):
            row_dict = {}
            has_data = False
            for col_idx, header in enumerate(effective_old_names, start=1):
                if not header:
                    continue
                value = ws_form.cell(row=row_idx, column=col_idx).value
                row_dict[header] = value
                if value not in (None, ""):
                    has_data = True
            if has_data:
                existing_row_dicts.append(row_dict)

    # Rebuild the sheet with the new schema.
    ws_form.delete_rows(1, ws_form.max_row)
    ws_form.append(new_names)

    desc_col_idx_1based = None
    item_code_col_idx_1based = None
    if "Description" in new_names:
        desc_col_idx_1based = new_names.index("Description") + 1
    if "ItemCode" in new_names:
        item_code_col_idx_1based = new_names.index("ItemCode") + 1

    for row_dict in existing_row_dicts:
        ws_form.append([row_dict.get(name, "") for name in new_names])

        # Ensure existing rows keep a correct Description formula after column
        # reordering. This prevents stale references to old ItemCode columns.
        if desc_col_idx_1based is not None and item_code_col_idx_1based is not None:
            row_num = ws_form.max_row
            item_code_cell_ref = f"{_col_idx_to_letter(item_code_col_idx_1based - 1)}{row_num}"
            ws_form.cell(
                row=row_num,
                column=desc_col_idx_1based,
                value=build_description_formula(item_code_cell_ref, "CREXPD01"),
            )

    try:
        wb.save(file_path)
    finally:
        wb.close()

    # Schema sync can change formulas/references; reset in-memory indexes and
    # trigger one recalc pass so Insert/Search resolve descriptions immediately.
    _invalidate_desc_index_cache()
    _excel_recalc_and_save(file_path)
    _invalidate_desc_index_cache()


def update_file_link(file_number: str, file_link: str) -> bool:
    """Update FileLink for a row in Heat Number sheet identified by File Number.

    Returns True when a matching row was updated, else False.
    """
    if not file_number:
        return False

    file_path = Path(EXCEL_FILE)
    if not file_path.exists():
        raise FileNotFoundError(f"Workbook not found: {EXCEL_FILE}")

    wb = _open_workbook()
    try:
        _, ws_form = _get_layout_sheets(wb)

        headers = []
        for col_idx in range(1, ws_form.max_column + 1):
            value = ws_form.cell(row=1, column=col_idx).value
            headers.append(str(value).strip() if value is not None else "")

        try:
            file_number_idx = headers.index("File Number") + 1
            file_link_idx = headers.index("FileLink") + 1
        except ValueError:
            return False

        target = str(file_number).strip().upper()
        updated = False

        for row_idx in range(2, ws_form.max_row + 1):
            current_value = ws_form.cell(row=row_idx, column=file_number_idx).value
            if str(current_value or "").strip().upper() != target:
                continue

            link_cell = ws_form.cell(row=row_idx, column=file_link_idx)
            link_cell.value = file_link
            if file_link:
                link_cell.hyperlink = file_link
                link_cell.style = "Hyperlink"
            else:
                link_cell.hyperlink = None

            updated = True
            break

        if updated:
            wb.save(file_path)
            _invalidate_desc_index_cache()
        return updated
    finally:
        wb.close()


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
