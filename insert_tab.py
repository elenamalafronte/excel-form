import time
import threading
from tkinter import filedialog, messagebox
from pathlib import Path

from customtkinter import CTkButton, CTkEntry, CTkFont, CTkFrame, CTkLabel, CTkTextbox, CTkScrollableFrame

from config import COLUMNS, EXCEL_FILE, get_next_fileNumber, get_next_fileNumber_from_value
from excel import append_row, load_sheet
from ui_style import (
    BODY_FONT_SIZE,
    BUTTON_CORNER_RADIUS,
    BUTTON_HEIGHT,
    CARD_CORNER_RADIUS,
    CONTROL_CORNER_RADIUS,
    ENTRY_HEIGHT,
    ENTRY_WIDTH,
    FORM_HEIGHT,
    FORM_WIDTH,
    LABEL_COLUMN_MIN_WIDTH,
    LABEL_FONT_SIZE,
    ROW_PADX,
    ROW_PADY,
    SECTION_TITLE_SIZE,
    TEXTBOX_HEIGHT,
)

# TODO: make font in table in search tab bigger (now too small)
# TODO: add dragger/something where you yourself can customise the width of the fields
# TODO: add button where you click and it shows a panel where you can customise the number of fiels + their properties(required, text/number/general format etc) 
#   then save that config chnage into the COLUMNS variable in config.py

def _open_file_picker(entry_widget):
    file_path = filedialog.askopenfilename(title="Select file for FileLink")
    if file_path:
        entry_widget.delete(0, "end")
        entry_widget.insert(0, file_path)


def _show_error(field_name, message):
    messagebox.showerror("Validation Error", f"{field_name}: {message}")


def _show_success(file_number):
    messagebox.showinfo("Saved", f"Row saved with File Number: {file_number}")


def _show_timed_success(file_number, elapsed_seconds, elapsed_before_save=None, elapsed_after_save=None):
    msg = f"Row saved with File Number: {file_number}"
    # optionally include save vs refresh breakdown for debugging performance issues
    # if elapsed_before_save is not None and elapsed_after_save is not None:
    #     save_duration = elapsed_after_save - elapsed_before_save
    #     refresh_duration = elapsed_seconds - elapsed_after_save
    #     msg += f"\n(Save: {save_duration:.3f}s, Refresh: {refresh_duration:.3f}s)"
    messagebox.showinfo("Saved", msg)


def _show_timed_error(message, elapsed_seconds):
    _show_error("Save", f"{message}\nElapsed: {elapsed_seconds:.3f}s")


_sheet_rows_cache = None
_sheet_rows_cache_mtime = None


def _get_cached_sheet_rows():
    """Return workbook rows for autofill, reloading only when the file changes.

    TODO: if you later want even faster autofill, keep this cache warm after
    startup instead of waiting for the first ItemCode lookup.
    """
    global _sheet_rows_cache, _sheet_rows_cache_mtime

    workbook_path = Path(EXCEL_FILE)
    if not workbook_path.exists():
        _sheet_rows_cache = []
        _sheet_rows_cache_mtime = None
        return _sheet_rows_cache

    try:
        current_mtime = workbook_path.stat().st_mtime
    except OSError:
        return load_sheet()

    if current_mtime != _sheet_rows_cache_mtime:
        _sheet_rows_cache = load_sheet()
        _sheet_rows_cache_mtime = current_mtime

    return _sheet_rows_cache


def _lookup_description_for_itemcode(item_code, rows):
    """Return the description for an ItemCode.

    This is the one place where the autofill rule lives.

    If your workbook ever changes shape, update the matching logic here and
    leave the rest of the UI code alone.
    """
    item_code = str(item_code).strip().upper()

    for row in rows or []:
        if str(row.get("ItemCode", "")).strip().upper() == item_code:
            # TODO: if your source description column has a different name,
            # change "Description" here only.
            return row.get("Description", "") or ""

    # Return an empty string when nothing matches so the field clears cleanly.
    return ""


def _update_description_field(description_widget, item_code_widget):
    """Update the Description field from the current ItemCode value.

    This is intentionally small: get the current ItemCode, look up the
    description, and write it into the disabled field.
    """
    item_code = item_code_widget.get().strip()

    if not item_code:
        description_widget.configure(state="normal")
        description_widget.delete("1.0", "end")
        description_widget.configure(state="disabled")
        return

    rows = _get_cached_sheet_rows()
    description = _lookup_description_for_itemcode(item_code, rows)

    description_widget.configure(state="normal")
    description_widget.delete("1.0", "end")
    description_widget.insert("1.0", description)
    description_widget.configure(state="disabled")


def _bind_itemcode_autofill(item_code_widget, description_widget):
    """Wire ItemCode edits to description autofill.
    """
    def _on_itemcode_event(event=None):
        _update_description_field(description_widget, item_code_widget)

    item_code_widget.bind("<FocusOut>", _on_itemcode_event)

    # make autofill to run on every keystroke.
    item_code_widget.bind("<KeyRelease>", _on_itemcode_event)

    return _on_itemcode_event


def build_insert_tab(tab):
    label_font = CTkFont(size=LABEL_FONT_SIZE)
    body_font = CTkFont(size=BODY_FONT_SIZE)
    title_font = CTkFont(size=SECTION_TITLE_SIZE, weight="bold")

    outer_container = CTkFrame(tab, fg_color="transparent")
    outer_container.pack(fill="both", expand=True, padx=12, pady=12)

    # Centered, scrollable card to keep the insert form tidy and readable.
    container = CTkScrollableFrame(
        outer_container,
        width=FORM_WIDTH,
        height=FORM_HEIGHT,
        corner_radius=CARD_CORNER_RADIUS,
    )
    container.pack(pady=8)
    container.grid_columnconfigure(0, minsize=LABEL_COLUMN_MIN_WIDTH)

    next_file_number_state = {"value": None}

    try:
        next_file_number_state["value"] = get_next_fileNumber(load_sheet())
    except Exception:
        next_file_number_state["value"] = "01-01A"

    fields = {}
    item_code_widget = None
    description_widget = None
    for row_idx, col in enumerate(COLUMNS):
        filelink_header = None
        if col.get("type") == "filelink":
            filelink_header = CTkFrame(container, fg_color="transparent")
            filelink_header.grid(
                row=row_idx,
                column=0,
                sticky="ew",
                padx=ROW_PADX,
                pady=ROW_PADY,
            )
            filelink_header.grid_columnconfigure(0, weight=1)
            CTkLabel(filelink_header, text=col["name"], font=label_font).grid(
                row=0,
                column=0,
                sticky="w",
            )
        else:
            CTkLabel(container, text=col["name"], font=label_font).grid(
                row=row_idx,
                column=0,
                sticky="w",
                padx=ROW_PADX,
                pady=ROW_PADY,
            )

        if col["name"] == "File Number":
            widget = CTkEntry(
                container,
                width=ENTRY_WIDTH,
                height=ENTRY_HEIGHT,
                corner_radius=CONTROL_CORNER_RADIUS,
                font=body_font,
            )
            widget.insert(0, "Auto-generated on Save")
            widget.configure(state="disabled")
            widget.grid(row=row_idx, column=1, sticky="ew", padx=ROW_PADX, pady=ROW_PADY)
        elif col["name"] == "Description":
            widget = CTkTextbox(
                container,
                width=ENTRY_WIDTH,
                height=TEXTBOX_HEIGHT + 10,  # add a bit of extra height so its visually clear this is a multi-line field
                corner_radius=CONTROL_CORNER_RADIUS,
                font=body_font,
                border_width=2,
                border_color="#979DA2",
            )
            widget.configure(state="disabled")
            widget.grid(row=row_idx, column=1, sticky="ew", padx=ROW_PADX, pady=ROW_PADY)
            description_widget = widget
        elif col["name"] == "ItemCode":
            widget = CTkEntry(
                container,
                width=ENTRY_WIDTH,
                height=ENTRY_HEIGHT,
                corner_radius=CONTROL_CORNER_RADIUS,
                font=body_font,
            )
            widget.grid(row=row_idx, column=1, sticky="ew", padx=ROW_PADX, pady=ROW_PADY)
            fields["ItemCode"] = widget
            item_code_widget = widget
        elif col.get("type") == "filelink":
            widget = CTkEntry(
                container,
                width=ENTRY_WIDTH,
                height=ENTRY_HEIGHT,
                corner_radius=CONTROL_CORNER_RADIUS,
                font=body_font,
            )
            CTkButton(
                filelink_header,
                text="Browse",
                width=100,
                height=BUTTON_HEIGHT,
                corner_radius=BUTTON_CORNER_RADIUS,
                font=body_font,
                command=lambda w=widget: _open_file_picker(w),
            ).grid(row=0, column=1, sticky="e")
            widget.grid(row=row_idx, column=1, sticky="ew", padx=ROW_PADX, pady=ROW_PADY)
        else:
            widget = CTkEntry(
                container,
                width=ENTRY_WIDTH,
                height=ENTRY_HEIGHT,
                corner_radius=CONTROL_CORNER_RADIUS,
                font=body_font,
            )
            widget.grid(row=row_idx, column=1, sticky="ew", padx=ROW_PADX, pady=ROW_PADY)

        fields[col["name"]] = widget

    if item_code_widget is not None and description_widget is not None:
        _bind_itemcode_autofill(item_code_widget, description_widget)

    # keep a reference so the thread callback can re-enable it
    save_button = CTkButton(
        container,
        text="Save Row",
        height=BUTTON_HEIGHT,
        corner_radius=BUTTON_CORNER_RADIUS,
        font=body_font,
    )

    def on_submit():
        started_at = time.perf_counter()
        file_number = next_file_number_state["value"]
        if not file_number:
            _show_error("File Number", "Could not determine the next File Number")
            return

        data = {}
        for col in COLUMNS:
            name = col["name"]
            if name == "File Number":
                data[name] = file_number
                continue
            if name == "Description":
                # The insert tab owns this value now.
                # If the ItemCode lookup has not run yet, this may still be blank.
                data[name] = fields[name].get("1.0", "end").strip()
                continue

            value = fields[name].get().strip()
            if col.get("required") and not value:
                _show_error(name, "This field is required")
                return
            data[name] = value

        # disable button so user can't double-submit while saving
        save_button.configure(state="disabled", text="Saving…")

        def do_save():
            try:
                elapsed_before_save = time.perf_counter() - started_at
                append_row(data)
                elapsed_after_save = time.perf_counter() - started_at
                tab.after(0, _on_save_success, file_number, started_at, elapsed_before_save, elapsed_after_save)
            except Exception as exc:
                tab.after(0, _on_save_error, str(exc), started_at)

        threading.Thread(target=do_save, daemon=True).start()

    def _on_save_success(file_number, started_at, elapsed_before_save, elapsed_after_save):
        elapsed_seconds = time.perf_counter() - started_at
        save_button.configure(state="normal", text="Save Row")
        _show_timed_success(file_number, elapsed_seconds, elapsed_before_save, elapsed_after_save)
        try:
            next_file_number_state["value"] = get_next_fileNumber_from_value(file_number)
        except Exception:
            next_file_number_state["value"] = None

        for col in COLUMNS:
            # Skip auto-filled fields. They are managed separately and should not
            # be cleared here unless you explicitly want to reset them after save.
            if col["name"] in ("File Number", "Description"):
                continue
            fields[col["name"]].delete(0, "end")

        refresh_search = getattr(tab, "refresh_search", None)
        if callable(refresh_search):
            refresh_search()

    def _on_save_error(message, started_at):
        elapsed_seconds = time.perf_counter() - started_at
        save_button.configure(state="normal", text="Save Row")
        _show_timed_error(message, elapsed_seconds)

    save_button.configure(command=on_submit)
    save_button.grid(
        row=len(COLUMNS) + 1,
        column=0,
        columnspan=3,
        sticky="ew",
        padx=ROW_PADX,
        pady=(12, 16),
    )
