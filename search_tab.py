import os
from tkinter import BooleanVar, Menu, filedialog, messagebox, ttk

from customtkinter import (
    CTkButton,
    CTkCheckBox,
    CTkComboBox,
    CTkEntry,
    CTkFont,
    CTkFrame,
    CTkLabel,
    CTkScrollbar,
    CTkScrollableFrame,
    CTkToplevel,
)

from config import COLUMNS, EXCEL_FILE, SEARCH_BY
from excel import search_rows, update_file_link
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
    TABLE_HEADING_FONT_SIZE,
    TABLE_ROW_HEIGHT,
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

    table_style = ttk.Style()
    table_style.configure(
        "App.Treeview",
        font=("Segoe UI", TABLE_FONT_SIZE),
        rowheight=TABLE_ROW_HEIGHT,
    )
    table_style.configure(
        "App.Treeview.Heading",
        font=("Segoe UI", TABLE_HEADING_FONT_SIZE, "bold"),
    )

    columns = [c["name"] for c in COLUMNS]
    default_column_width = 138
    results_tree = ttk.Treeview(container, columns=columns, show="headings", style="App.Treeview")
    for col_name in columns:
        results_tree.heading(col_name, text=col_name, anchor="center")
        results_tree.column(col_name, width=default_column_width, stretch=True, anchor="center")

    results_tree.grid(row=2, column=0, columnspan=5, sticky="nsew", padx=ROW_PADX, pady=8)

    y_scroll = CTkScrollbar(container, orientation="vertical", command=results_tree.yview)
    y_scroll.grid(row=2, column=5, sticky="ns", pady=8)
    results_tree.configure(yscrollcommand=y_scroll.set)

    x_scroll = CTkScrollbar(container, orientation="horizontal", command=results_tree.xview)
    x_scroll.grid(row=3, column=0, columnspan=5, sticky="ew", padx=ROW_PADX, pady=(0, 8))
    results_tree.configure(xscrollcommand=x_scroll.set)

    hidden_columns = set()
    column_visibility_vars = {}

    def hide_column(selected):
        if selected not in columns:
            return

        hidden_columns.add(selected)
        results_tree.column(selected, width=0, minwidth=0, stretch=False)

        var = column_visibility_vars.get(selected)
        if var is not None and var.get():
            var.set(False)

    def show_column(selected):
        if selected not in columns:
            return

        hidden_columns.discard(selected)
        results_tree.column(selected, width=default_column_width, minwidth=20, stretch=True)

        var = column_visibility_vars.get(selected)
        if var is not None and not var.get():
            var.set(True)

    def show_all_columns():
        hidden_columns.clear()
        for col_name in columns:
            results_tree.column(col_name, width=default_column_width, minwidth=20, stretch=True)

        for col_name in columns:
            var = column_visibility_vars.get(col_name)
            if var is not None:
                var.set(True)

    def hide_all_columns():
        hidden_columns.clear()
        for col_name in columns:
            hidden_columns.add(col_name)
            results_tree.column(col_name, width=0, minwidth=0, stretch=False)

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

    header_menu = Menu(results_tree, tearoff=0)

    def on_header_right_click(event):
        if results_tree.identify_region(event.x, event.y) != "heading":
            return

        col_id = results_tree.identify_column(event.x)
        if not col_id or col_id == "#0":
            return

        col_index = int(col_id.replace("#", "")) - 1
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

    file_number_col_idx = columns.index("File Number") if "File Number" in columns else -1
    file_link_col_idx = columns.index("FileLink") if "FileLink" in columns else -1

    selected_info_label = None
    upload_pdf_button = None
    selected_actions = CTkFrame(container, fg_color="transparent")
    selected_actions.grid(row=4, column=1, columnspan=3, sticky="w", padx=ROW_PADX, pady=10)
    selected_actions.grid_remove()

    def _selected_row_values():
        selected = results_tree.selection()
        if not selected:
            return None
        values = results_tree.item(selected[0], "values")
        return list(values) if values else None

    def _set_selected_actions_visible(visible):
        if visible:
            selected_actions.grid()
        else:
            selected_actions.grid_remove()

    def _update_selected_actions_ui():
        if selected_info_label is None or upload_pdf_button is None:
            return

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

    def on_search():
        rows = search_rows(search_entry.get().strip(), search_by.get())
        results_tree.delete(*results_tree.get_children())
        for row in rows:
            values = [("" if row.get(c["name"]) is None else row.get(c["name"])) for c in COLUMNS]
            results_tree.insert("", "end", values=values)

    def refresh_with_feedback():
        if refresh_button is None:
            on_search()
            return

        original_text = refresh_button.cget("text")
        refresh_button.configure(state="disabled", text="Refreshing...")
        container.update_idletasks()
        try:
            on_search()
        finally:
            refresh_button.configure(state="normal", text=original_text)

    def upload_pdf_for_selected_row():
        values = _selected_row_values()
        if not values or file_number_col_idx < 0:
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
            upload_pdf_button.configure(state="disabled", text="Uploading PDF...")
            container.update_idletasks()

        try:
            updated = update_file_link(file_number, file_path)
            if not updated:
                messagebox.showwarning("Upload PDF", "Could not find the selected row in workbook.")
                return

            on_search()
            # Restore selection after refresh when possible.
            for item_id in results_tree.get_children():
                row_values = results_tree.item(item_id, "values")
                if file_number_col_idx < len(row_values) and str(row_values[file_number_col_idx] or "").strip() == file_number:
                    results_tree.selection_set(item_id)
                    results_tree.focus(item_id)
                    break
            _update_selected_actions_ui()
        except Exception as exc:
            messagebox.showerror("Upload PDF", f"Could not save PDF link:\n{exc}")
        finally:
            if upload_pdf_button is not None:
                upload_pdf_button.configure(state="normal", text=original_text)

    tab.refresh_search = on_search
    tab.auto_refresh_search = lambda: on_search()

    def open_workbook():
        try:
            os.startfile(EXCEL_FILE)
        except Exception as exc:
            messagebox.showerror("Open Workbook", f"Could not open workbook:\n{exc}")

    def handle_cell_click(event):
        item_id = results_tree.identify_row(event.y)
        col_id = results_tree.identify_column(event.x)
        if not item_id or not col_id:
            return

        col_index = int(col_id.replace("#", "")) - 1
        if col_index < 0 or col_index >= len(COLUMNS):
            return

        if COLUMNS[col_index]["name"] != "FileLink":
            return

        values = results_tree.item(item_id, "values")
        file_link = values[col_index] if col_index < len(values) else ""
        if not file_link:
            return

        try:
            os.startfile(file_link)  # Windows-specific
        except Exception as exc:
            messagebox.showerror("Open File", f"Could not open file link:\n{exc}")

    results_tree.bind("<Double-1>", handle_cell_click)
    results_tree.bind("<Button-3>", on_header_right_click)
    results_tree.bind("<<TreeviewSelect>>", lambda event: _update_selected_actions_ui())

    on_search()

    column_visibility_text = "Column Visibility"
    column_visibility_button_width = body_font.measure(column_visibility_text) + 26

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

    CTkButton(
        top_actions,
        text="Search",
        command=on_search,
        height=BUTTON_HEIGHT,
        corner_radius=BUTTON_CORNER_RADIUS,
        font=body_font,
        width=110,
    ).pack(side="left")
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
    CTkButton(
        container,
        text="Open Workbook",
        command=open_workbook,
        height=BUTTON_HEIGHT,
        corner_radius=BUTTON_CORNER_RADIUS,
        font=body_font,
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
