import time
import threading
from tkinter import filedialog, messagebox

from customtkinter import CTkButton, CTkEntry, CTkFrame, CTkLabel

from config import COLUMNS, get_next_fileNumber, get_next_fileNumber_from_value
from excel import append_row, load_sheet


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
    msg = f"Row saved with File Number: {file_number}\nElapsed: {elapsed_seconds:.3f}s"
    if elapsed_before_save is not None and elapsed_after_save is not None:
        save_duration = elapsed_after_save - elapsed_before_save
        refresh_duration = elapsed_seconds - elapsed_after_save
        msg += f"\n(Save: {save_duration:.3f}s, Refresh: {refresh_duration:.3f}s)"
    messagebox.showinfo("Saved", msg)


def _show_timed_error(message, elapsed_seconds):
    _show_error("Save", f"{message}\nElapsed: {elapsed_seconds:.3f}s")


def build_insert_tab(tab):
    container = CTkFrame(tab)
    container.pack(fill="both", expand=True, padx=12, pady=12)

    next_file_number_state = {"value": None}

    try:
        next_file_number_state["value"] = get_next_fileNumber(load_sheet())
    except Exception:
        next_file_number_state["value"] = "01-01A"

    fields = {}
    for row_idx, col in enumerate(COLUMNS):
        CTkLabel(container, text=col["name"]).grid(row=row_idx, column=0, sticky="w", padx=8, pady=6)

        if col["name"] == "File Number":
            widget = CTkEntry(container)
            widget.insert(0, "Auto-generated on Save")
            widget.configure(state="disabled")
            widget.grid(row=row_idx, column=1, sticky="ew", padx=8, pady=6)
        elif col["name"] == "Description":
            widget = CTkEntry(container)
            widget.insert(0, "Auto-filled from ItemCode")
            widget.configure(state="disabled")
            widget.grid(row=row_idx, column=1, sticky="ew", padx=8, pady=6)
        elif col.get("type") == "filelink":
            widget = CTkEntry(container)
            widget.grid(row=row_idx, column=1, sticky="ew", padx=8, pady=6)
            CTkButton(
                container,
                text="Browse",
                width=90,
                command=lambda w=widget: _open_file_picker(w),
            ).grid(row=row_idx, column=2, sticky="w", padx=8, pady=6)
        else:
            widget = CTkEntry(container)
            widget.grid(row=row_idx, column=1, sticky="ew", padx=8, pady=6)

        fields[col["name"]] = widget

    container.grid_columnconfigure(1, weight=1)

    # keep a reference so the thread callback can re-enable it
    save_button = CTkButton(container, text="Save Row")

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
                # auto-filled by formula in excel.py — don't read the disabled widget
                data[name] = ""
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
            # skip disabled/auto fields — calling .delete() on them raises TclError
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
        padx=8,
        pady=14,
    )
