import os
from tkinter import messagebox, ttk

from customtkinter import CTkButton, CTkComboBox, CTkEntry, CTkFrame, CTkLabel

from config import COLUMNS, EXCEL_FILE, SEARCH_BY
from excel import search_rows


def build_search_tab(tab):
    container = CTkFrame(tab)
    container.pack(fill="both", expand=True, padx=12, pady=12)

    CTkLabel(container, text="Search Value").grid(row=0, column=0, sticky="w", padx=8, pady=6)
    search_entry = CTkEntry(container)
    search_entry.grid(row=0, column=1, sticky="ew", padx=8, pady=6)

    CTkLabel(container, text="Search By").grid(row=0, column=2, sticky="w", padx=8, pady=6)
    search_by = CTkComboBox(container, values=SEARCH_BY)
    search_by.set("ItemCode" if "ItemCode" in SEARCH_BY else SEARCH_BY[0])
    search_by.grid(row=0, column=3, sticky="ew", padx=8, pady=6)

    columns = [c["name"] for c in COLUMNS]
    results_tree = ttk.Treeview(container, columns=columns, show="headings")
    for col_name in columns:
        results_tree.heading(col_name, text=col_name)
        results_tree.column(col_name, width=130, stretch=True)

    results_tree.grid(row=1, column=0, columnspan=5, sticky="nsew", padx=8, pady=8)

    y_scroll = ttk.Scrollbar(container, orient="vertical", command=results_tree.yview)
    y_scroll.grid(row=1, column=5, sticky="ns", pady=8)
    results_tree.configure(yscrollcommand=y_scroll.set)

    def on_search():
        rows = search_rows(search_entry.get().strip(), search_by.get())
        results_tree.delete(*results_tree.get_children())
        for row in rows:
            values = [row.get(c["name"], "") for c in COLUMNS]
            results_tree.insert("", "end", values=values)

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

    CTkButton(container, text="Search", command=on_search).grid(
        row=0, column=4, sticky="ew", padx=8, pady=6
    )
    CTkButton(container, text="Refresh", command=on_search).grid(
        row=2, column=0, sticky="w", padx=8, pady=6
    )
    CTkButton(container, text="Open Workbook", command=open_workbook).grid(
        row=2, column=4, sticky="ew", padx=8, pady=6
    )

    container.grid_columnconfigure(1, weight=1)
    container.grid_columnconfigure(3, weight=1)
    container.grid_rowconfigure(1, weight=1)
