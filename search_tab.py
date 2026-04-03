import os
from tkinter import messagebox, ttk

from customtkinter import CTkButton, CTkComboBox, CTkEntry, CTkFont, CTkFrame, CTkLabel

from config import COLUMNS, EXCEL_FILE, SEARCH_BY
from excel import search_rows
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
    results_tree = ttk.Treeview(container, columns=columns, show="headings", style="App.Treeview")
    centered_columns = {"File Number", "Qty_EA", "Qty_mt", "Rev", "PAGENr"}
    for col_name in columns:
        results_tree.heading(col_name, text=col_name)
        anchor = "center" if col_name in centered_columns else "w"
        results_tree.column(col_name, width=138, stretch=True, anchor=anchor)

    results_tree.grid(row=2, column=0, columnspan=5, sticky="nsew", padx=ROW_PADX, pady=8)

    y_scroll = ttk.Scrollbar(container, orient="vertical", command=results_tree.yview)
    y_scroll.grid(row=2, column=5, sticky="ns", pady=8)
    results_tree.configure(yscrollcommand=y_scroll.set)

    refresh_button = None

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

    on_search()

    CTkButton(
        container,
        text="Search",
        command=on_search,
        height=BUTTON_HEIGHT,
        corner_radius=BUTTON_CORNER_RADIUS,
        font=body_font,
    ).grid(
        row=1, column=4, sticky="ew", padx=ROW_PADX, pady=ROW_PADY
    )
    refresh_button = CTkButton(
        container,
        text="Refresh",
        command=refresh_with_feedback,
        height=BUTTON_HEIGHT,
        corner_radius=BUTTON_CORNER_RADIUS,
        font=body_font,
    )
    refresh_button.grid(
        row=3, column=0, sticky="w", padx=ROW_PADX, pady=10
    )
    CTkButton(
        container,
        text="Open Workbook",
        command=open_workbook,
        height=BUTTON_HEIGHT,
        corner_radius=BUTTON_CORNER_RADIUS,
        font=body_font,
    ).grid(
        row=3, column=4, sticky="ew", padx=ROW_PADX, pady=10
    )

    container.grid_columnconfigure(1, weight=1)
    container.grid_columnconfigure(3, weight=1)
    container.grid_rowconfigure(2, weight=1)
