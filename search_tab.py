import os
import threading
import time
from tkinter import BooleanVar, Menu, filedialog, messagebox
from tkinter import font as tkfont

import config as cfg

from customtkinter import (
    CTkButton,
    CTkCheckBox,
    CTkComboBox,
    CTkEntry,
    CTkFont,
    CTkFrame,
    CTkLabel,
    CTkScrollableFrame,
    CTkToplevel,
)
from tksheet import Sheet

from config import COLUMNS, SEARCH_BY
from excel import recalc_workbook, search_rows, update_file_link
from ui_style import (
    BODY_FONT_SIZE,
    BUTTON_CORNER_RADIUS,
    BUTTON_HEIGHT,
    CARD_CORNER_RADIUS,
    CONTROL_CORNER_RADIUS,
    ENTRY_HEIGHT,
    LABEL_FONT_SIZE,
    ROW_PADX,
    ROW_PADY,
    SECTION_TITLE_SIZE,
    TABLE_FONT_SIZE,
    TABLE_HEADING_FONT_SIZE
)


_BUTTON_IDLE_FG = ("#3B8ED0", "#1F6AA5")
_BUTTON_IDLE_HOVER = ("#36719F", "#144870")
_BUTTON_BUSY_FG = "#1F4D82"
_ROW_SELECTED_BG = "#DCEBFF"
_ROW_SELECTED_FG = "#0D1B2A"
_CELL_SELECTED_BG = "#6E9FEA"
_CELL_SELECTED_FG = "#FFFFFF"
_MAX_CUSTOM_HIGHLIGHT_ROWS = 1200
_MAX_WRAP_RECALC_ROWS = 800


def _set_button_loading_state(button, is_loading, idle_text, busy_text, refresh_widget=None):
    if button is None:
        return

    if is_loading:
        button.configure(
            state="normal",
            text=busy_text,
            fg_color=_BUTTON_BUSY_FG,
            hover_color=_BUTTON_BUSY_FG,
        )
        if refresh_widget is not None:
            refresh_widget.update_idletasks()
    else:
        button.configure(
            state="normal",
            text=idle_text,
            fg_color=_BUTTON_IDLE_FG,
            hover_color=_BUTTON_IDLE_HOVER,
        )


def build_search_tab(tab):
    label_font = CTkFont(size=LABEL_FONT_SIZE)
    body_font = CTkFont(size=BODY_FONT_SIZE)
    title_font = CTkFont(size=SECTION_TITLE_SIZE, weight="bold")

    outer_container = CTkFrame(tab, fg_color="transparent")
    outer_container.pack(fill="both", expand=True, padx=12, pady=12)

    container = CTkFrame(outer_container, corner_radius=CARD_CORNER_RADIUS)
    container.pack(fill="both", expand=True)

    toolbar = CTkFrame(container, fg_color="transparent")
    toolbar.grid(row=0, column=0, columnspan=5, sticky="e", padx=ROW_PADX, pady=(8, 0))

    CTkLabel(container, text="Search Value", font=label_font).grid(
        row=1, column=0, sticky="w", padx=ROW_PADX, pady=ROW_PADY
    )
    search_entry = CTkEntry(
        container,
        height=ENTRY_HEIGHT,
        corner_radius=CONTROL_CORNER_RADIUS,
        font=body_font,
    )
    search_entry.grid(row=1, column=1, sticky="ew", padx=ROW_PADX, pady=ROW_PADY)

    CTkLabel(container, text="Search By", font=label_font).grid(
        row=1, column=2, sticky="w", padx=ROW_PADX, pady=ROW_PADY
    )
    search_by = CTkComboBox(
        container,
        values=SEARCH_BY,
        height=ENTRY_HEIGHT,
        corner_radius=CONTROL_CORNER_RADIUS,
        font=body_font,
    )
    search_by.set("ItemCode" if "ItemCode" in SEARCH_BY else SEARCH_BY[0])
    search_by.grid(row=1, column=3, sticky="ew", padx=ROW_PADX, pady=ROW_PADY)

    columns = [c["name"] for c in COLUMNS]
    default_column_width = 138
    max_auto_row_height = 96
    results_sheet = Sheet(
        container,
        headers=columns,
        data=[],
        show_row_index=False,
        show_top_left=False,
        show_x_scrollbar=True,
        show_y_scrollbar=True,
        table_wrap="w",
        header_wrap="w",
        align="center",
        header_align="center",
    )
    results_sheet.grid(row=2, column=0, columnspan=6, sticky="nsew", padx=ROW_PADX, pady=8)
    # Keep native wrapping, but compute row heights ourselves for stable/minimum sizing.
    results_sheet.set_options(auto_resize_rows=False, table_wrap="w", header_wrap="w")
    results_sheet.enable_bindings("all")
    results_sheet.font(("Segoe UI", TABLE_FONT_SIZE, "normal"))
    results_sheet.header_font(("Segoe UI", TABLE_HEADING_FONT_SIZE, "bold"))
    results_sheet.set_all_column_widths(default_column_width, redraw=False)
    results_sheet.redraw()

    hidden_columns = set()
    column_visibility_vars = {}
    current_rows = []
    search_request_id = {"value": 0}
    search_in_progress = {"value": False}
    search_started_at = {"value": None}
    last_fitted_width = -1
    auto_fit_columns = True
    row_height_recalc_job = None
    max_header_lines = 3
    sort_state = {"column": None, "ascending": True}

    def fit_columns_to_available_width(redraw=False):
        visible_indexes = [idx for idx, col_name in enumerate(columns) if col_name not in hidden_columns]
        if not visible_indexes:
            visible_indexes = list(range(len(columns)))

        visible_count = len(visible_indexes)
        if visible_count <= 0:
            return False

        available_width = max(
            results_sheet.winfo_width(),
            container.winfo_width() - (ROW_PADX * 2),
        )
        if available_width <= 20:
            return False

        # Use an exact width budget so visible columns fill the table width.
        usable_width = max(visible_count * 60, int(available_width))
        description_weight = 2.4
        weights = [description_weight if columns[idx] == "Description" else 1.0 for idx in visible_indexes]

        total_weight = sum(weights) if weights else 1.0
        assigned_widths = [int(usable_width * (w / total_weight)) for w in weights]

        # Guarantee the budget is fully used after integer rounding.
        remainder = usable_width - sum(assigned_widths)
        if assigned_widths:
            assigned_widths[-1] += remainder

        # Keep columns usable on narrow windows.
        min_width = 60
        deficit = 0
        for i, w in enumerate(assigned_widths):
            if w < min_width:
                deficit += (min_width - w)
                assigned_widths[i] = min_width

        if deficit > 0 and assigned_widths:
            flex_indexes = [i for i, idx in enumerate(visible_indexes) if columns[idx] != "Description"]
            if not flex_indexes:
                flex_indexes = list(range(len(assigned_widths) - 1))
            for i in reversed(flex_indexes):
                if deficit <= 0:
                    break
                reducible = max(0, assigned_widths[i] - min_width)
                take = min(reducible, deficit)
                assigned_widths[i] -= take
                deficit -= take
            if deficit > 0:
                assigned_widths[-1] = max(min_width, assigned_widths[-1] - deficit)

        for col_idx, width in zip(visible_indexes, assigned_widths):
            results_sheet.column_width(col_idx, width, redraw=False)

        if redraw:
            results_sheet.redraw()
        return True

    def recompute_row_heights(redraw=False):
        # Measuring wrapped text for thousands of rows can freeze the UI.
        # Keep default row heights for large result sets.
        if len(current_rows) > _MAX_WRAP_RECALC_ROWS:
            if redraw:
                results_sheet.redraw()
            return

        # Compute the minimum needed height per row from wrapped visible text.
        # This avoids oversized rows from generic auto-resize behavior.
        visible_indexes = [idx for idx, col_name in enumerate(columns) if col_name not in hidden_columns]
        if not visible_indexes:
            visible_indexes = list(range(len(columns)))

        widths = results_sheet.get_column_widths()
        cell_font = tkfont.Font(font=results_sheet.font())
        line_height = max(cell_font.metrics("linespace"), 12)
        min_row_height = max(24, line_height + 6)

        def wrapped_line_count(text, max_width):
            text = str(text or "")
            if not text:
                return 1
            if max_width <= 8:
                return max(1, text.count("\n") + 1)

            lines = 0
            for paragraph in text.split("\n"):
                if paragraph == "":
                    lines += 1
                    continue

                current = ""
                for token in paragraph.split(" "):
                    candidate = token if not current else f"{current} {token}"
                    if cell_font.measure(candidate) <= max_width:
                        current = candidate
                        continue

                    if current:
                        lines += 1
                        current = ""

                    # Hard-break long tokens that exceed the cell width.
                    piece = ""
                    for ch in token:
                        candidate_piece = f"{piece}{ch}"
                        if cell_font.measure(candidate_piece) <= max_width:
                            piece = candidate_piece
                        else:
                            lines += 1
                            piece = ch
                    current = piece

                if current:
                    lines += 1

            return max(lines, 1)

        for row_idx, row_data in enumerate(current_rows):
            needed_lines = 1
            for col_idx in visible_indexes:
                col_name = columns[col_idx]
                text = row_data.get(col_name, "")
                col_width = int(widths[col_idx]) if col_idx < len(widths) else default_column_width
                usable_width = max(10, col_width - 14)
                needed_lines = max(needed_lines, wrapped_line_count(text, usable_width))

            target_height = min(max_auto_row_height, max(min_row_height, (needed_lines * line_height) + 6))
            results_sheet.row_height(row_idx, target_height, redraw=False)

        if redraw:
            results_sheet.redraw()

    def recompute_header_height():
        # Keep header compact: 1 line when possible, otherwise cap at 2 lines.
        header_font = tkfont.Font(font=results_sheet.header_font())
        widths = results_sheet.get_column_widths()

        def wrapped_line_count(text, max_width):
            text = str(text or "")
            if max_width <= 8 or not text:
                return 1

            lines = 0
            for paragraph in text.split("\n"):
                if paragraph == "":
                    lines += 1
                    continue

                current = ""
                for token in paragraph.split(" "):
                    candidate = token if not current else f"{current} {token}"
                    if header_font.measure(candidate) <= max_width:
                        current = candidate
                        continue

                    if current:
                        lines += 1
                        current = ""

                    # Hard-break very long tokens that exceed the cell width.
                    piece = ""
                    for ch in token:
                        candidate_piece = f"{piece}{ch}"
                        if header_font.measure(candidate_piece) <= max_width:
                            piece = candidate_piece
                        else:
                            lines += 1
                            piece = ch
                    current = piece

                if current:
                    lines += 1

            return max(lines, 1)

        needed_lines = 1
        for idx, col_name in enumerate(columns):
            if col_name in hidden_columns:
                continue
            col_width = int(widths[idx]) if idx < len(widths) else default_column_width
            usable_width = max(10, col_width - 14)
            needed_lines = max(needed_lines, wrapped_line_count(col_name, usable_width))

        results_sheet.set_header_height_lines(min(max(needed_lines, 1), 2), redraw=False)

    def schedule_row_height_recalc(delay_ms=80):
        nonlocal row_height_recalc_job
        if len(current_rows) > _MAX_WRAP_RECALC_ROWS:
            return

        if row_height_recalc_job is not None:
            tab.after_cancel(row_height_recalc_job)

        def _run():
            nonlocal row_height_recalc_job
            row_height_recalc_job = None
            recompute_header_height()
            recompute_row_heights(redraw=True)

        row_height_recalc_job = tab.after(delay_ms, _run)

    def apply_display_columns():
        if not hidden_columns:
            results_sheet.display_columns(
                columns="all",
                all_columns_displayed=True,
                reset_col_positions=False,
                redraw=True,
                deselect_all=False,
            )
            if auto_fit_columns:
                fit_columns_to_available_width(redraw=True)
            schedule_row_height_recalc()
            return

        visible_indexes = [idx for idx, col_name in enumerate(columns) if col_name not in hidden_columns]
        results_sheet.display_columns(
            columns=visible_indexes,
            all_columns_displayed=False,
            reset_col_positions=False,
            redraw=True,
            deselect_all=False,
        )
        if auto_fit_columns:
            fit_columns_to_available_width(redraw=True)
        schedule_row_height_recalc()

    def get_selected_row_index():
        selected_rows = results_sheet.get_selected_rows(get_cells_as_rows=True)
        if not selected_rows:
            current = results_sheet.get_currently_selected()
            if current and len(current) >= 1 and isinstance(current[0], int):
                selected_rows = [current[0]]
            else:
                return None

        normalized_rows = []
        for selected in selected_rows:
            if isinstance(selected, tuple) and selected:
                normalized_rows.append(selected[0])
            elif isinstance(selected, int):
                normalized_rows.append(selected)

        if not normalized_rows:
            return None

        row_index = min(normalized_rows)
        if row_index < 0 or row_index >= len(current_rows):
            return None
        return row_index

    def hide_column(selected):
        if selected not in columns:
            return

        hidden_columns.add(selected)
        apply_display_columns()

        var = column_visibility_vars.get(selected)
        if var is not None and var.get():
            var.set(False)

    def show_column(selected):
        if selected not in columns:
            return

        hidden_columns.discard(selected)
        apply_display_columns()

        var = column_visibility_vars.get(selected)
        if var is not None and not var.get():
            var.set(True)

    def show_all_columns():
        hidden_columns.clear()
        apply_display_columns()

        for col_name in columns:
            var = column_visibility_vars.get(col_name)
            if var is not None:
                var.set(True)

    def hide_all_columns():
        hidden_columns.clear()
        for col_name in columns:
            hidden_columns.add(col_name)

        apply_display_columns()

        for col_name in columns:
            var = column_visibility_vars.get(col_name)
            if var is not None:
                var.set(False)

    def _on_column_toggle(column_name):
        var = column_visibility_vars.get(column_name)
        if var is None:
            return

        if var.get():
            show_column(column_name)
        else:
            hide_column(column_name)

    def open_columns_panel():
        panel = CTkToplevel(tab)
        panel.title("Choose Columns")
        panel.geometry("360x520")
        panel.transient(tab.winfo_toplevel())
        panel.grab_set()

        CTkLabel(panel, text="Column Visibility", font=CTkFont(size=16, weight="bold")).pack(
            anchor="w", padx=12, pady=(12, 8)
        )

        list_frame = CTkScrollableFrame(panel, corner_radius=10)
        list_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        column_visibility_vars.clear()
        for col_name in columns:
            is_visible = col_name not in hidden_columns
            var = BooleanVar(value=is_visible)
            column_visibility_vars[col_name] = var
            CTkCheckBox(
                list_frame,
                text=col_name,
                variable=var,
                command=lambda c=col_name: _on_column_toggle(c),
            ).pack(anchor="w", padx=8, pady=6)

        actions = CTkFrame(panel, fg_color="transparent")
        actions.pack(fill="x", padx=12, pady=(0, 12))
        actions.grid_columnconfigure(0, weight=1)
        actions.grid_columnconfigure(1, weight=1)
        actions.grid_columnconfigure(2, weight=1)

        CTkButton(
            actions,
            text="Hide All",
            command=hide_all_columns,
            height=BUTTON_HEIGHT,
            corner_radius=BUTTON_CORNER_RADIUS,
            font=body_font,
        ).grid(row=0, column=0, sticky="ew", padx=(0, 6))

        CTkButton(
            actions,
            text="Show All",
            command=show_all_columns,
            height=BUTTON_HEIGHT,
            corner_radius=BUTTON_CORNER_RADIUS,
            font=body_font,
        ).grid(row=0, column=1, sticky="ew", padx=6)

        CTkButton(
            actions,
            text="Close",
            command=panel.destroy,
            height=BUTTON_HEIGHT,
            corner_radius=BUTTON_CORNER_RADIUS,
            font=body_font,
        ).grid(row=0, column=2, sticky="ew", padx=(6, 0))

    header_menu = Menu(results_sheet, tearoff=0)

    def on_header_right_click(event):
        if results_sheet.identify_region(event) != "header":
            return

        displayed_col_index = results_sheet.identify_column(event)
        if displayed_col_index is None:
            return

        try:
            col_index = results_sheet.displayed_column_to_data(displayed_col_index)
        except Exception:
            return

        if col_index < 0 or col_index >= len(columns):
            return

        col_name = columns[col_index]

        header_menu.delete(0, "end")
        header_menu.add_command(label=f"Hide '{col_name}'", command=lambda c=col_name: hide_column(c))
        header_menu.add_separator()
        header_menu.add_command(label="Column Visibility", command=open_columns_panel)
        header_menu.add_command(label="Show All Columns", command=show_all_columns)
        header_menu.tk_popup(event.x_root, event.y_root)

    refresh_button = None
    search_button = None
    search_status_label = None

    def _set_search_status(text, color="#767676"):
        if search_status_label is None:
            return
        search_status_label.configure(text=text, text_color=color)

    file_number_col_idx = columns.index("File Number") if "File Number" in columns else -1
    file_link_col_idx = columns.index("FileLink") if "FileLink" in columns else -1

    selected_info_label = None
    upload_pdf_button = None
    selected_actions = CTkFrame(container, fg_color="transparent")
    selected_actions.grid(row=4, column=1, columnspan=3, sticky="w", padx=ROW_PADX, pady=10)
    selected_actions.grid_remove()

    def _selected_row_values():
        row_index = get_selected_row_index()
        if row_index is None:
            return None

        row = current_rows[row_index]
        return [("" if row.get(col_name) is None else row.get(col_name)) for col_name in columns]

    def _set_selected_actions_visible(visible):
        if visible:
            selected_actions.grid()
        else:
            selected_actions.grid_remove()

    def _highlight_selected_row_and_cell():
        # Full dehighlight/highlight on every click is expensive on very large
        # tables and can lock the UI. Fall back to native tksheet selection.
        if len(current_rows) > _MAX_CUSTOM_HIGHLIGHT_ROWS:
            return

        row_index = get_selected_row_index()
        selected = results_sheet.get_currently_selected()
        col_index = None
        if selected and len(selected) >= 2 and isinstance(selected[1], int):
            col_index = selected[1]

        try:
            results_sheet.dehighlight_rows("all", redraw=False)
            results_sheet.dehighlight_cells(all_=True, redraw=False)

            if row_index is not None:
                results_sheet.highlight_rows(
                    row_index,
                    bg=_ROW_SELECTED_BG,
                    fg=_ROW_SELECTED_FG,
                    redraw=False,
                )
                if isinstance(col_index, int) and col_index >= 0:
                    results_sheet.highlight_cells(
                        row=row_index,
                        column=col_index,
                        bg=_CELL_SELECTED_BG,
                        fg=_CELL_SELECTED_FG,
                        redraw=False,
                    )

            results_sheet.redraw()
        except Exception:
            # Keep selection behavior usable even if highlight styling fails.
            pass

    def _update_selected_actions_ui():
        if selected_info_label is None or upload_pdf_button is None:
            return

        _highlight_selected_row_and_cell()

        values = _selected_row_values()
        if not values:
            _set_selected_actions_visible(False)
            return

        file_number = ""
        file_link = ""
        if 0 <= file_number_col_idx < len(values):
            file_number = str(values[file_number_col_idx] or "").strip()
        if 0 <= file_link_col_idx < len(values):
            file_link = str(values[file_link_col_idx] or "").strip()

        if not file_number:
            _set_selected_actions_visible(False)
            return

        selected_info_label.configure(text=f"Selected: {file_number}")
        upload_pdf_button.configure(text="Upload PDF" if not file_link else "Replace PDF")
        _set_selected_actions_visible(True)

    def _sort_key(row, column_name):
        value = row.get(column_name)
        if value is None:
            return (2, "")

        if isinstance(value, (int, float)):
            return (0, float(value))

        text = str(value).strip()
        if text == "":
            return (2, "")

        try:
            return (0, float(text))
        except ValueError:
            return (1, text.lower())

    def _apply_sort_state():
        nonlocal current_rows
        col_name = sort_state.get("column")
        if not col_name:
            return

        ascending = bool(sort_state.get("ascending", True))
        current_rows = sorted(current_rows, key=lambda row: _sort_key(row, col_name), reverse=not ascending)

    def _render_rows():
        data = []
        for row in current_rows:
            data.append([("" if row.get(c["name"]) is None else row.get(c["name"])) for c in COLUMNS])

        large_result_set = len(current_rows) > _MAX_WRAP_RECALC_ROWS
        results_sheet.set_options(table_wrap="n" if large_result_set else "w")

        results_sheet.headers(columns, redraw=False)
        recompute_header_height()
        results_sheet.set_sheet_data(
            data,
            reset_col_positions=False,
            reset_row_positions=True,
            redraw=False,
            keep_formatting=False,
        )
        # Defensive fallback: some tksheet states can show headers but skip body rows.
        if data and results_sheet.total_rows() == 0:
            results_sheet.set_sheet_data(
                data,
                reset_col_positions=True,
                reset_row_positions=True,
                redraw=False,
                keep_formatting=False,
            )
        if auto_fit_columns:
            fit_columns_to_available_width(redraw=False)
        apply_display_columns()
        if not large_result_set:
            recompute_row_heights(redraw=False)
        results_sheet.deselect("all", redraw=False)
        results_sheet.redraw()
        _update_selected_actions_ui()

    def _set_search_loading(is_loading):
        search_in_progress["value"] = is_loading

        if is_loading:
            _set_search_status("Loading rows...", "#4E7DA8")

        try:
            search_entry.configure(state="disabled" if is_loading else "normal")
            search_by.configure(state="disabled" if is_loading else "normal")
        except Exception:
            pass

        if refresh_button is not None:
            refresh_button.configure(
                state="disabled" if is_loading else "normal",
                text="Refreshing..." if is_loading else "Refresh",
            )

        if search_button is not None:
            _set_button_loading_state(
                search_button,
                is_loading,
                idle_text="Search",
                busy_text="Loading...",
                refresh_widget=container if is_loading else None,
            )

    def sort_by_column(col_name):
        if col_name not in columns:
            return

        if sort_state.get("column") == col_name:
            sort_state["ascending"] = not bool(sort_state.get("ascending", True))
        else:
            sort_state["column"] = col_name
            sort_state["ascending"] = True

        _apply_sort_state()
        _render_rows()

    def on_search():
        nonlocal current_rows
        if search_in_progress["value"]:
            return

        search_value = search_entry.get().strip()
        search_column = search_by.get()
        search_request_id["value"] += 1
        request_id = search_request_id["value"]
        search_started_at["value"] = time.perf_counter()
        _set_search_loading(True)

        def _do_search():
            error_msg = None
            rows = []
            try:
                rows = search_rows(search_value, search_column)
            except Exception as exc:
                error_msg = str(exc)
            finally:
                tab.after(0, _finish_search, request_id, rows, error_msg)

        def _finish_search(request_id_from_thread, rows, error_msg):
            nonlocal current_rows
            # Ignore stale responses from older requests.
            if request_id_from_thread != search_request_id["value"]:
                return

            _set_search_loading(False)

            if error_msg:
                _set_search_status("Load failed", "#B23A48")
                messagebox.showerror("Search", f"Could not load rows:\n{error_msg}")
                return

            current_rows = list(rows)
            _apply_sort_state()
            _render_rows()

            elapsed = 0.0
            started = search_started_at.get("value")
            if started is not None:
                elapsed = max(0.0, time.perf_counter() - started)

            if search_value:
                if current_rows:
                    _set_search_status(f"Found {len(current_rows)} rows in {elapsed:.1f}s", "#2E8B57")
                else:
                    _set_search_status(f"No matches ({elapsed:.1f}s)", "#B27A2F")
            else:
                _set_search_status(f"Loaded {len(current_rows)} rows in {elapsed:.1f}s", "#2E8B57")

        threading.Thread(target=_do_search, daemon=True).start()

    def refresh_with_feedback():
        if refresh_button is None:
            on_search()
            return

        on_search()

    def refresh_after_recalc():
        on_search()

    def upload_pdf_for_selected_row():
        row_index = get_selected_row_index()
        values = _selected_row_values()
        if row_index is None or not values or file_number_col_idx < 0:
            messagebox.showwarning("Upload PDF", "Select a row first.")
            return

        file_number = str(values[file_number_col_idx] if file_number_col_idx < len(values) else "").strip()
        if not file_number:
            messagebox.showwarning("Upload PDF", "Selected row has no File Number.")
            return

        file_path = filedialog.askopenfilename(
            title="Select PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if not file_path:
            return

        original_text = upload_pdf_button.cget("text") if upload_pdf_button is not None else "Upload PDF"
        if upload_pdf_button is not None:
            _set_button_loading_state(
                upload_pdf_button,
                True,
                idle_text=original_text,
                busy_text="Uploading....",
                refresh_widget=container,
            )

        def do_upload():
            """Background thread worker for PDF upload."""
            error_msg = None
            updated = False
            try:
                updated = update_file_link(file_number, file_path)
            except Exception as exc:
                error_msg = f"Could not save PDF link:\n{exc}"
            finally:
                # Re-enable button on main thread
                container.after(0, _finish_upload, error_msg, updated)

        def _finish_upload(error_msg, updated):
            """Finish upload and update UI on main thread."""
            if upload_pdf_button is not None:
                _set_button_loading_state(
                    upload_pdf_button,
                    False,
                    idle_text=original_text,
                    busy_text="Uploading....",
                )

            if error_msg:
                messagebox.showerror("Upload PDF", error_msg)
                return

            if not updated:
                messagebox.showwarning("Upload PDF", "Could not find the selected row in workbook.")
                return

            refresh_after_recalc()
            for idx, row in enumerate(current_rows):
                value = "" if row.get("File Number") is None else str(row.get("File Number")).strip()
                if value == file_number:
                    results_sheet.select_row(idx, redraw=True, run_binding_func=False)
                    break
            _update_selected_actions_ui()

        # Start background upload thread
        upload_thread = threading.Thread(target=do_upload, daemon=True)
        upload_thread.start()

    tab.refresh_search = on_search
    tab.refresh_search_with_recalc = refresh_after_recalc
    tab.auto_refresh_search = lambda: on_search()

    def open_workbook():
        try:
            os.startfile(cfg.EXCEL_FILE)
        except Exception as exc:
            messagebox.showerror("Open Workbook", f"Could not open workbook:\n{exc}")

    def handle_double_click(event):
        region = results_sheet.identify_region(event)

        if region == "header":
            displayed_col_index = results_sheet.identify_column(event)
            if displayed_col_index is None:
                return

            # Ignore double-click near header borders to avoid conflicting with resize actions.
            try:
                col_positions = results_sheet.get_column_widths(canvas_positions=True)
                if 0 <= displayed_col_index + 1 < len(col_positions):
                    right_edge_x = col_positions[displayed_col_index + 1]
                    if abs(int(event.x) - int(right_edge_x)) <= 6:
                        return
            except Exception:
                pass

            try:
                data_col_index = results_sheet.displayed_column_to_data(displayed_col_index)
            except Exception:
                return

            if 0 <= data_col_index < len(columns):
                sort_by_column(columns[data_col_index])
            return

        if region != "table":
            return

        row_index = results_sheet.identify_row(event, exclude_index=True)
        displayed_col_index = results_sheet.identify_column(event, exclude_header=True)
        if row_index is None or displayed_col_index is None:
            return

        try:
            col_index = results_sheet.displayed_column_to_data(displayed_col_index)
        except Exception:
            return

        if col_index < 0 or col_index >= len(COLUMNS):
            return

        if COLUMNS[col_index]["name"] != "FileLink":
            return

        if row_index < 0 or row_index >= len(current_rows):
            return

        file_link = current_rows[row_index].get("FileLink", "")
        if not file_link:
            return

        try:
            os.startfile(file_link)  # Windows-specific
        except Exception as exc:
            messagebox.showerror("Open File", f"Could not open file link:\n{exc}")

    def handle_left_release(event):
        _update_selected_actions_ui()

    results_sheet.bind("<Double-1>", handle_double_click)
    results_sheet.bind("<Button-3>", on_header_right_click)
    results_sheet.bind("<ButtonRelease-1>", handle_left_release)
    results_sheet.bind("<KeyRelease>", lambda event: _update_selected_actions_ui())

    def on_user_column_resize(event=None):
        nonlocal auto_fit_columns
        auto_fit_columns = False
        schedule_row_height_recalc()

    results_sheet.extra_bindings("column_width_resize", func=on_user_column_resize)
    results_sheet.extra_bindings("double_click_column_resize", func=on_user_column_resize)

    def on_sheet_configure(event):
        nonlocal last_fitted_width
        if event.width <= 20:
            return
        if abs(event.width - last_fitted_width) < 3:
            return

        # Always recompute wrapped heights when viewport width changes.
        # Auto-fit columns only while user hasn't manually resized columns.
        if auto_fit_columns:
            fit_columns_to_available_width(redraw=True)
        recompute_header_height()
        if len(current_rows) <= _MAX_WRAP_RECALC_ROWS:
            recompute_row_heights(redraw=True)
            schedule_row_height_recalc(140)
        else:
            results_sheet.redraw()
        last_fitted_width = event.width

    results_sheet.bind("<Configure>", on_sheet_configure, add="+")

    # Auto-load rows when Search tab is created, but do it after first paint
    # so the tab opens immediately and loading remains non-blocking.
    tab.after(120, on_search)

    column_visibility_text = "Column Visibility"
    column_visibility_button_width = body_font.measure(column_visibility_text) + 26

    CTkLabel(
        toolbar,
        text="Tip: double-click a column header to sort ascending/descending.",
        font=CTkFont(size=max(BODY_FONT_SIZE - 1, 10)),
        text_color="#767676",
    ).pack(side="left", padx=(0, 10))

    CTkButton(
        toolbar,
        text=column_visibility_text,
        command=open_columns_panel,
        height=BUTTON_HEIGHT,
        width=max(column_visibility_button_width, column_visibility_button_width + 10),
        corner_radius=BUTTON_CORNER_RADIUS,
        font=body_font,
    ).pack(side="right")

    top_actions = CTkFrame(container, fg_color="transparent")
    top_actions.grid(row=1, column=4, sticky="e", padx=ROW_PADX, pady=ROW_PADY)

    search_button = CTkButton(
        top_actions,
        text="Search",
        command=on_search,
        height=BUTTON_HEIGHT,
        corner_radius=BUTTON_CORNER_RADIUS,
        font=body_font,
        width=110,
    )
    search_button.pack(side="left")
    refresh_button = CTkButton(
        container,
        text="Refresh",
        command=refresh_with_feedback,
        height=BUTTON_HEIGHT,
        corner_radius=BUTTON_CORNER_RADIUS,
        font=body_font,
    )
    refresh_button.grid(
        row=4, column=0, sticky="w", padx=ROW_PADX, pady=10
    )
    search_status_label = CTkLabel(
        container,
        text="Ready",
        font=CTkFont(size=max(BODY_FONT_SIZE - 1, 10), weight="bold"),
        text_color="#767676",
        anchor="w",
    )
    search_status_label.grid(row=5, column=0, sticky="w", padx=ROW_PADX, pady=(0, 10))

    CTkButton(
        container,
        text="Open Workbook",
        command=open_workbook,
        height=BUTTON_HEIGHT,
        corner_radius=BUTTON_CORNER_RADIUS,
        font=body_font,
        fg_color="#2E8B57",
        hover_color="#236B43",
    ).grid(
        row=4, column=4, sticky="ew", padx=ROW_PADX, pady=10
    )

    selected_info_label = CTkLabel(selected_actions, text="", font=label_font)
    selected_info_label.pack(side="left", padx=(0, 10))
    upload_pdf_button = CTkButton(
        selected_actions,
        text="Upload PDF",
        command=upload_pdf_for_selected_row,
        height=BUTTON_HEIGHT,
        corner_radius=BUTTON_CORNER_RADIUS,
        font=body_font,
    )
    upload_pdf_button.pack(side="left")

    container.grid_columnconfigure(1, weight=1, minsize=220)
    container.grid_columnconfigure(3, weight=1, minsize=170)
    container.grid_rowconfigure(2, weight=1)
