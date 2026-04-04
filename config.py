import re
import ast
import pprint
from pathlib import Path

EXCEL_FILE = "Heat number summary.xlsx"

# Workbook layout:
# 1) first sheet = raw data, read-only
# 2) second sheet = form output, read/write
# 3) third sheet = generated final sheet, untouched by this app

COLUMNS = [{'name': 'File Number',
  'type': 'text',
  'required': True,
  'unique': True,
  'validate': 'is_valid_fileNumber'},
 {'name': 'PO/MR', 'type': 'text', 'required': True, 'unique': False, 'validate': None},
 {'name': 'Manufacturer/Supplier',
  'type': 'text',
  'required': True,
  'unique': False,
  'validate': None},
 {'name': 'ItemCode', 'type': 'text', 'required': True, 'unique': False, 'validate': None},
 {'name': 'Description', 'type': 'general', 'required': False, 'validate': None},
 {'name': 'ManufacturerTestReport(MTR)No',
  'type': 'text',
  'required': True,
  'unique': False,
  'validate': None},
 {'name': 'HeatNumber', 'type': 'text', 'required': True, 'unique': False, 'validate': None},
 {'name': 'PackiglistNo', 'type': 'text', 'required': True, 'unique': False, 'validate': None},
 {'name': 'ShippingNotice', 'type': 'text', 'required': True, 'unique': False, 'validate': None},
 {'name': 'QualityControlManufactDossier(QCMD)',
  'type': 'text',
  'required': True,
  'unique': False,
  'validate': None},
 {'name': 'Rev', 'type': 'general', 'required': True, 'unique': False, 'validate': None},
 {'name': 'PAGENr', 'type': 'text', 'required': True, 'unique': False, 'validate': None},
 {'name': 'FileLink', 'type': 'filelink', 'required': False, 'validate': None},
 {'name': 'Qt_Test', 'type': 'text', 'required': False}]


def save_columns_config(new_columns, config_file_path=None):
    """Persist updated COLUMNS into config.py.

    This rewrites only the COLUMNS assignment block and keeps the rest of
    config.py untouched.
    """
    target_path = Path(config_file_path) if config_file_path else Path(__file__)
    source = target_path.read_text(encoding="utf-8")

    module = ast.parse(source)
    columns_assign = None
    for node in module.body:
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name) and target.id == "COLUMNS":
                    columns_assign = node
                    break
        if columns_assign is not None:
            break

    if columns_assign is None:
        raise ValueError("Could not find COLUMNS assignment in config.py")

    lines = source.splitlines(keepends=True)
    start_idx = columns_assign.lineno - 1
    end_idx = columns_assign.end_lineno

    formatted_columns = pprint.pformat(new_columns, width=100, sort_dicts=False)
    replacement = f"COLUMNS = {formatted_columns}\n"

    updated_source = "".join(lines[:start_idx]) + replacement + "".join(lines[end_idx:])
    target_path.write_text(updated_source, encoding="utf-8")

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

