import re
import ast
import pprint
import sys
import os
from pathlib import Path


def _runtime_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _user_config_dir() -> Path:
    """Return user-writable config directory."""
    if sys.platform == "win32":
        base = Path(os.getenv("LOCALAPPDATA", os.path.expanduser("~\\AppData\\Local")))
    else:
        base = Path.home() / ".config"
    config_dir = base / "ExcelForm"
    config_dir.mkdir(parents=True, exist_ok=True)
    return config_dir


def _user_config_file() -> Path:
    """Return path to user config file."""
    return _user_config_dir() / "config.py"


EXCEL_FILE = ""
USER_HAS_SELECTED_WORKBOOK = False
SOURCE_SHEET_NAME = 'CREXPD01'
FORM_SHEET_NAME = 'Heat Number'


def get_excel_file_path() -> Path | None:
    """Return the configured workbook path, or None when no workbook is loaded."""
    value = str(EXCEL_FILE).strip()
    if not value:
        return None
    return Path(value).expanduser()

# Workbook layout:
# 1) first sheet = raw data, read-only
# 2) second sheet = form output, read/write
# 3) third sheet = generated final sheet, untouched by this app

COLUMNS = [{'name': 'File Number',
  'type': 'text',
  'required': True,
  'unique': True,
  'validate': 'is_valid_fileNumber'},
 {'name': 'ItemCode', 'type': 'text', 'required': True, 'unique': False, 'validate': None},
 {'name': 'Description', 'type': 'general', 'required': False, 'validate': None},
 {'name': 'Qty_EA', 'type': 'text', 'required': False},
 {'name': 'Qty_mt', 'type': 'text', 'required': False},
 {'name': 'HeatNumber', 'type': 'text', 'required': False},
 {'name': 'ManufacturerTestReport(MTR)No', 'type': 'text', 'required': False},
 {'name': 'Manufacturer/Supplier', 'type': 'text', 'required': False},
 {'name': 'PackinglistNo', 'type': 'text', 'required': False},
 {'name': 'ShippingNotice', 'type': 'text', 'required': False},
 {'name': 'PO/MR', 'type': 'text', 'required': False},
 {'name': 'QualityControlManufactDossier(QCMD)',
  'type': 'text',
  'required': False,
  'unique': False,
  'validate': None},
 {'name': 'Rev', 'type': 'general', 'required': False, 'unique': False, 'validate': None},
 {'name': 'PAGENr', 'type': 'text', 'required': False, 'unique': False, 'validate': None},
 {'name': 'FileLink', 'type': 'filelink', 'required': False, 'validate': None}]


def save_columns_config(new_columns, config_file_path=None):
    """Persist updated COLUMNS into user config directory.

    Saves to user-writable AppData location instead of bundled config.py.
    """
    target_path = Path(config_file_path) if config_file_path else _user_config_file()
    
    # Read existing user config if it exists, otherwise start fresh
    if target_path.exists():
        source = target_path.read_text(encoding="utf-8")
        module = ast.parse(source)
    else:
        # Create new user config with just COLUMNS
        source = ""

    module = ast.parse(source) if source else ast.parse("COLUMNS = []")
    columns_assign = None
    for node in module.body:
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name) and target.id == "COLUMNS":
                    columns_assign = node
                    break
        if columns_assign is not None:
            break

    lines = source.splitlines(keepends=True) if source else []
    if columns_assign is not None:
        start_idx = columns_assign.lineno - 1
        end_idx = columns_assign.end_lineno
        updated_source = "".join(lines[:start_idx]) + f"COLUMNS = {pprint.pformat(new_columns, width=100, sort_dicts=False)}\n" + "".join(lines[end_idx:])
    else:
        # No existing COLUMNS assignment, append it
        updated_source = f"COLUMNS = {pprint.pformat(new_columns, width=100, sort_dicts=False)}\n"

    target_path.write_text(updated_source, encoding="utf-8")


def _replace_constant_assignment(source, constant_name, new_value):
    """Replace or add a constant assignment in source code."""
    module = ast.parse(source) if source else ast.parse("")
    target_node = None

    for node in module.body:
        if isinstance(node, ast.Assign) and len(node.targets) == 1:
            target = node.targets[0]
            if isinstance(target, ast.Name) and target.id == constant_name:
                target_node = node
                break

    if target_node is None:
        # Constant not found, append it
        return source + f"{constant_name} = {new_value!r}\n"

    lines = source.splitlines(keepends=True)
    start_idx = target_node.lineno - 1
    end_idx = target_node.end_lineno

    replacement = f"{constant_name} = {new_value!r}\n"
    return "".join(lines[:start_idx]) + replacement + "".join(lines[end_idx:])


def save_workbook_settings(excel_file=None, source_sheet_name=None, form_sheet_name=None, config_file_path=None):
    """Persist workbook path and sheet names into user config directory.

    Saves to user-writable AppData location instead of bundled config.py.
    Updates the module globals too so the running app can use the new values
    immediately without a restart.
    """
    target_path = Path(config_file_path) if config_file_path else _user_config_file()
    
    # Read existing user config if it exists, otherwise start with empty
    if target_path.exists():
        source = target_path.read_text(encoding="utf-8")
    else:
        source = ""

    updates = {}
    if excel_file is not None:
        updates["EXCEL_FILE"] = str(excel_file)
        updates["USER_HAS_SELECTED_WORKBOOK"] = True
    if source_sheet_name is not None:
        updates["SOURCE_SHEET_NAME"] = str(source_sheet_name)
    if form_sheet_name is not None:
        updates["FORM_SHEET_NAME"] = str(form_sheet_name)

    updated_source = source
    for constant_name, value in updates.items():
        updated_source = _replace_constant_assignment(updated_source, constant_name, value)
        globals()[constant_name] = value

    target_path.write_text(updated_source, encoding="utf-8")


def load_user_config():
    """Load user config overrides from AppData and apply to globals."""
    user_config_path = _user_config_file()
    if not user_config_path.exists():
        return

    try:
        user_config_source = user_config_path.read_text(encoding="utf-8")
        user_module = ast.parse(user_config_source)

        user_has_selected_workbook = False
        for node in user_module.body:
            if isinstance(node, ast.Assign):
                for target in node.targets:
                    if isinstance(target, ast.Name) and target.id == "USER_HAS_SELECTED_WORKBOOK":
                        if isinstance(node.value, ast.Constant):
                            user_has_selected_workbook = bool(node.value.value)
                        break
        
        # Extract and apply overrides
        for node in user_module.body:
            if isinstance(node, ast.Assign):
                for target in node.targets:
                    if isinstance(target, ast.Name):
                        if target.id == "EXCEL_FILE" and not user_has_selected_workbook:
                            continue
                        # Evaluate the value safely (only for simple literals)
                        if isinstance(node.value, ast.Constant):
                            globals()[target.id] = node.value.value
                        elif isinstance(node.value, ast.List):
                            # For COLUMNS, we need to eval carefully
                            try:
                                value = ast.literal_eval(node.value)
                                globals()[target.id] = value
                            except (ValueError, TypeError):
                                pass
    except Exception as e:
        # If user config is malformed, just skip and use defaults
        print(f"Warning: could not load user config from {user_config_path}: {e}", file=sys.stderr)


# Load user config overrides if they exist
load_user_config()

SEARCH_BY = [col["name"] for col in COLUMNS]  # should be able to search by any column

DESCRIPTION_LOOKUP_FORMULA_TEMPLATE = (
    '=IFERROR(INDEX({source_sheet}!$A$1:$AV$15000,MATCH({item_code_cell},{source_sheet}!$O$1:$O$15000,0),'
    'MATCH("Detailed Description",{source_sheet}!$A$1:$AV$1,0)),"")'
)

FILE_NUMBER_PATTERN = re.compile(r"^(\d{2})-(\d{2})([A-Z])$")



def _to_file_number_index(file_number):
    m = FILE_NUMBER_PATTERN.match(file_number)
    if not m:
        return None
    nn = int(m.group(1))
    sub = int(m.group(2))
    letter = ord(m.group(3)) - ord("A")
    return ((nn - 1) * 26 * 99) + (letter * 99) + (sub - 1)


def _from_file_number_index(index):
    if index < 0 or index >= (99 * 26 * 99):
        return None

    nn = (index // (26 * 99)) + 1
    rem = index % (26 * 99)
    letter_index = rem // 99
    sub = (rem % 99) + 1
    letter = chr(ord("A") + letter_index)
    return f"{nn:02d}-{sub:02d}{letter}"


def _extract_existing_file_numbers(all_rows):
    existing_values = []
    for row in all_rows or []:
        if isinstance(row, dict):
            raw = row.get("File Number") or row.get("FileNumber")
        elif isinstance(row, (list, tuple)) and row:
            raw = row[0]
        else:
            raw = None

        if raw is None:
            continue

        value = str(raw).strip().upper()
        if FILE_NUMBER_PATTERN.match(value):
            existing_values.append(value)

    return existing_values


def get_next_fileNumber(all_rows):
    existing_values = _extract_existing_file_numbers(all_rows)
    if not existing_values:
        return "01-01A"

    max_index = max(_to_file_number_index(v) for v in existing_values)
    next_value = _from_file_number_index(max_index + 1)
    if next_value is None:
        raise ValueError("File Number sequence exceeded 99-99Z")
    return next_value


def get_next_fileNumber_from_value(value):
    index = _to_file_number_index(value)
    if index is None:
        raise ValueError("Could not advance File Number")

    next_value = _from_file_number_index(index + 1)
    if next_value is None:
        raise ValueError("File Number sequence exceeded 99-99Z")
    return next_value


def build_description_formula(item_code_cell, source_sheet):
    """Return the Excel formula used to auto-populate Description.

    The formula looks up the detailed description in the source sheet using the
    current row's ItemCode cell reference.
    """
    return DESCRIPTION_LOOKUP_FORMULA_TEMPLATE.format(
        source_sheet=source_sheet,
        item_code_cell=item_code_cell,
    )

def is_valid_fileNumber(value, all_rows):
    # Expected order: 01-01A ... 01-99A, 01-01B ... 01-99Z, 02-01A ... 99-99Z.

    if value is None:
        return "File Number is required"

    value = str(value).strip().upper()
    if not value:
        return "File Number is required"

    match = FILE_NUMBER_PATTERN.match(value)
    if not match:
        return "File Number must match format NN-NNL (example: 01-01A)"

    part_1 = int(match.group(1))
    part_2 = int(match.group(2))
    if not (1 <= part_1 <= 99 and 1 <= part_2 <= 99):
        return "File Number numeric parts must be between 01 and 99"

    existing_values = _extract_existing_file_numbers(all_rows)

    if value in existing_values:
        return "File Number must be unique"

    expected_next = get_next_fileNumber(all_rows)
    if value != expected_next:
        return f"File Number must be the next value in sequence: {expected_next}"

    return None

