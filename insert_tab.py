import time
import threading
from tkinter import BooleanVar, filedialog, messagebox
from pathlib import Path

from customtkinter import (
    CTkButton,
    CTkCheckBox,
    CTkComboBox,
    CTkEntry,
    CTkFont,
    CTkFrame,
    CTkLabel,
    CTkScrollableFrame,
    CTkTextbox,
    CTkToplevel,
)

from config import (
    COLUMNS,
    EXCEL_FILE,
    SEARCH_BY,
    get_next_fileNumber,
    get_next_fileNumber_from_value,
    save_columns_config,
)
from excel import append_row, get_description_for_itemcode, load_sheet, sync_form_sheet_columns
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


def _invalidate_sheet_rows_cache():
    global _sheet_rows_cache, _sheet_rows_cache_mtime
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
            # Prefer a non-empty description when duplicates exist.
            description = (row.get("Description", "") or "").strip()
            if description:
                return description

    # Fallback to source sheet lookup so new ItemCodes autofill before save.
    return get_description_for_itemcode(item_code)


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
    field_types = ["text", "number", "general", "filelink"]
    existing_by_name = {c.get("name"): c for c in COLUMNS}

    def _rebuild_insert_tab():
        for child in tab.winfo_children():
            child.destroy()
        build_insert_tab(tab)

    def _open_fields_customizer():
        panel = CTkToplevel(tab)
        panel.title("Customize Fields")
        panel.geometry("720x560")
        panel.transient(tab.winfo_toplevel())
        panel.grab_set()

        CTkLabel(panel, text="Customize Insert Fields", font=CTkFont(size=18, weight="bold")).pack(
            anchor="w", padx=12, pady=(12, 8)
        )

        rows_frame = CTkScrollableFrame(panel)
        rows_frame.pack(fill="both", expand=True, padx=12, pady=(0, 10))
        rows_frame.grid_columnconfigure(0, weight=0)
        rows_frame.grid_columnconfigure(1, weight=3)
        rows_frame.grid_columnconfigure(2, weight=2)
        rows_frame.grid_columnconfigure(3, weight=0)
        required_lane_width = 76
        checkbox_size = 22
        checkbox_pad_x = max((required_lane_width - checkbox_size) // 2, 0)

        CTkLabel(rows_frame, text="Field Name", font=CTkFont(size=13, weight="bold")).grid(
            row=0, column=1, sticky="w", padx=8, pady=(8, 6)
        )
        CTkLabel(rows_frame, text="Type", font=CTkFont(size=13, weight="bold")).grid(
            row=0, column=2, sticky="w", padx=8, pady=(8, 6)
        )
        CTkLabel(
            rows_frame,
            text="Required",
            font=CTkFont(size=13, weight="bold"),
            width=required_lane_width,
            anchor="center",
        ).grid(row=0, column=3, sticky="w", padx=(0, 0), pady=(8, 6))

        row_models = []
        removed_rows_stack = []
        undo_button = None
        drag_state = {"model": None}

        def _update_undo_button_state():
            if undo_button is None:
                return
            undo_button.configure(state="normal" if removed_rows_stack else "disabled")

        def _reflow_rows():
            for idx, row in enumerate(row_models, start=1):
                row["drag_handle"].grid(row=idx, column=0, sticky="w", padx=(8, 4), pady=4)
                row["name_entry"].grid(row=idx, column=1, sticky="ew", padx=8, pady=4)
                row["type_combo"].grid(row=idx, column=2, sticky="ew", padx=8, pady=4)
                row["action_cell"].grid(row=idx, column=3, sticky="w", padx=(6, 6), pady=4)

        def _move_row(model, target_index):
            if model not in row_models:
                return
            current_index = row_models.index(model)
            target_index = max(0, min(target_index, len(row_models) - 1))
            if current_index == target_index:
                return
            row_models.pop(current_index)
            row_models.insert(target_index, model)
            _reflow_rows()

        def _drag_target_index():
            pointer_y = rows_frame.winfo_pointery()
            for idx, row in enumerate(row_models):
                center_y = row["name_entry"].winfo_rooty() + (row["name_entry"].winfo_height() // 2)
                if pointer_y < center_y:
                    return idx
            return len(row_models) - 1

        def _on_drag_start(event, model):
            drag_state["model"] = model
            handle = model.get("drag_handle")
            if handle is not None:
                handle.configure(fg_color="#2C8C67")

        def _on_drag_motion(event):
            model = drag_state.get("model")
            if model is None:
                return
            _move_row(model, _drag_target_index())

        def _on_drag_end(event):
            model = drag_state.get("model")
            if model is not None:
                handle = model.get("drag_handle")
                if handle is not None:
                    handle.configure(fg_color="transparent")
            drag_state["model"] = None

        def _remove_row(model):
            if len(row_models) <= 1:
                messagebox.showwarning("Customize Fields", "At least one field is required.")
                return

            row_index = row_models.index(model)
            removed_field = {
                "name": model["name_entry"].get().strip(),
                "type": model["type_combo"].get().strip().lower() or "text",
                "required": bool(model["required_var"].get()),
                "original_name": model.get("original_name"),
            }

            for widget in (
                model["name_entry"],
                model["type_combo"],
                model["drag_handle"],
                model["action_cell"],
            ):
                widget.destroy()

            row_models.remove(model)
            removed_rows_stack.append({"index": row_index, "field": removed_field})
            _reflow_rows()
            _update_undo_button_state()

        def _undo_remove_row():
            if not removed_rows_stack:
                return

            removed_entry = removed_rows_stack.pop()
            _add_row(
                initial=removed_entry["field"],
                insert_at=min(max(0, removed_entry["index"]), len(row_models)),
            )
            _update_undo_button_state()

        def _add_row(initial=None, insert_at=None):
            initial = initial or {}
            required_var = BooleanVar(value=bool(initial.get("required", False)))

            name_entry = CTkEntry(rows_frame)
            name_entry.insert(0, str(initial.get("name", "")))

            type_combo = CTkComboBox(rows_frame, values=field_types)
            initial_type = str(initial.get("type", "text")).strip().lower()
            type_combo.set(initial_type if initial_type in field_types else "text")

            action_cell = CTkFrame(rows_frame, fg_color="transparent")

            required_check = CTkCheckBox(
                action_cell,
                text="",
                variable=required_var,
                width=26,
                checkbox_width=checkbox_size,
                checkbox_height=checkbox_size,
            )
            required_check.pack(side="left", padx=(checkbox_pad_x, checkbox_pad_x))

            drag_handle = CTkLabel(
                rows_frame,
                text="::",
                width=26,
                corner_radius=6,
                fg_color="transparent",
                cursor="fleur",
            )

            model = {
                "original_name": initial.get("original_name", initial.get("name")),
                "name_entry": name_entry,
                "type_combo": type_combo,
                "required_var": required_var,
                "required_check": required_check,
                "action_cell": action_cell,
                "drag_handle": drag_handle,
                "remove_btn": None,
            }

            drag_handle.bind("<ButtonPress-1>", lambda event, m=model: _on_drag_start(event, m))
            drag_handle.bind("<B1-Motion>", _on_drag_motion)
            drag_handle.bind("<ButtonRelease-1>", _on_drag_end)

            remove_btn = CTkButton(
                action_cell,
                text="Remove",
                width=86,
                command=lambda m=model: _remove_row(m),
            )
            remove_btn.pack(side="left", padx=(12, 0))
            model["remove_btn"] = remove_btn

            if insert_at is None or insert_at < 0 or insert_at > len(row_models):
                row_models.append(model)
            else:
                row_models.insert(insert_at, model)
            _reflow_rows()

        def _save_customized_fields():
            old_columns_snapshot = [dict(c) for c in COLUMNS]
            new_columns = []
            seen_names = set()

            for row in row_models:
                name = row["name_entry"].get().strip()
                col_type = row["type_combo"].get().strip().lower()
                required = bool(row["required_var"].get())

                if not name:
                    messagebox.showerror("Customize Fields", "Field name cannot be empty.")
                    return
                if name in seen_names:
                    messagebox.showerror("Customize Fields", f"Duplicate field name: {name}")
                    return
                if col_type not in field_types:
                    messagebox.showerror("Customize Fields", f"Invalid type for {name}: {col_type}")
                    return

                seen_names.add(name)

                original_name = row.get("original_name")
                existing = existing_by_name.get(original_name) or existing_by_name.get(name) or {}
                col_def = {
                    "name": name,
                    "type": col_type,
                    "required": required,
                }

                if "unique" in existing:
                    col_def["unique"] = existing["unique"]
                if "validate" in existing:
                    col_def["validate"] = existing["validate"]

                if name == "File Number":
                    col_def["required"] = True
                    col_def["unique"] = True
                    col_def["validate"] = "is_valid_fileNumber"

                new_columns.append(col_def)

            if not new_columns:
                messagebox.showerror("Customize Fields", "At least one field is required.")
                return

            has_item_code = any(col.get("name") == "ItemCode" for col in new_columns)
            has_description = any(col.get("name") == "Description" for col in new_columns)
            if has_description and not has_item_code:
                messagebox.showerror(
                    "Customize Fields",
                    "Description requires ItemCode so the auto-formula can keep working.",
                )
                return

            try:
                save_columns_config(new_columns)
            except Exception as exc:
                messagebox.showerror("Customize Fields", f"Could not save config.py:\n{exc}")
                return

            try:
                sync_form_sheet_columns(old_columns_snapshot, new_columns)
            except Exception as exc:
                # Roll back config file so app config and workbook schema do not diverge.
                try:
                    save_columns_config(old_columns_snapshot)
                except Exception:
                    pass
                messagebox.showerror(
                    "Customize Fields",
                    "Config was not applied because workbook schema sync failed.\n"
                    f"Reason: {exc}",
                )
                return

            COLUMNS.clear()
            COLUMNS.extend(new_columns)

            _invalidate_sheet_rows_cache()

            SEARCH_BY.clear()
            SEARCH_BY.extend([col["name"] for col in new_columns])

            rebuild_search = getattr(tab, "rebuild_search", None)
            if callable(rebuild_search):
                rebuild_search()

            panel.destroy()
            messagebox.showinfo(
                "Customize Fields",
                "Field configuration saved and workbook columns updated.",
            )
            _rebuild_insert_tab()

        for col in COLUMNS:
            _add_row(col)

        buttons = CTkFrame(panel, fg_color="transparent")
        buttons.pack(fill="x", padx=12, pady=(0, 12))

        CTkButton(buttons, text="Add Field", width=100, command=lambda: _add_row()).pack(side="left")
        undo_button = CTkButton(buttons, text="Undo Remove", width=120, command=_undo_remove_row)
        undo_button.pack(side="left", padx=(8, 0))
        _update_undo_button_state()
        CTkButton(buttons, text="Cancel", width=100, command=panel.destroy).pack(side="right")
        CTkButton(buttons, text="Save", width=100, command=_save_customized_fields).pack(
            side="right", padx=(0, 8)
        )

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

    header_row = CTkFrame(container, fg_color="transparent")
    header_row.grid(row=0, column=0, columnspan=3, sticky="ew", padx=ROW_PADX, pady=(10, 6))
    header_row.grid_columnconfigure(0, weight=1)

    CTkLabel(header_row, text="Insert Material Record", font=title_font).grid(
        row=0,
        column=0,
        sticky="w",
    )
    CTkButton(
        header_row,
        text="Customize Fields",
        height=BUTTON_HEIGHT,
        corner_radius=BUTTON_CORNER_RADIUS,
        font=body_font,
        command=_open_fields_customizer,
    ).grid(row=0, column=1, sticky="e")

    next_file_number_state = {"value": None}

    try:
        next_file_number_state["value"] = get_next_fileNumber(load_sheet())
    except Exception:
        next_file_number_state["value"] = "01-01A"

    fields = {}
    item_code_widget = None
    description_widget = None
    for row_idx, col in enumerate(COLUMNS, start=1):
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
        row=len(COLUMNS) + 2,
        column=0,
        columnspan=3,
        sticky="ew",
        padx=ROW_PADX,
        pady=(12, 16),
    )
