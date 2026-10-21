"""Windows desktop interface. Run with python app.py."""

from __future__ import annotations

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from checker import Cancelled, Options, Source, check_files, export_report, read_table, sheet_names
from mapping_dialog import MappingDialog

FILE_TYPES = [("Excel and CSV", "*.xlsx *.xlsm *.csv"), ("Excel workbook", "*.xlsx *.xlsm"), ("CSV", "*.csv")]
UPLOADED_ROWS = "Validate uploaded rows (mapped fields)"
FULL_COPY = "Compare complete database copy"


class CheckerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CHCODE Database Checker")
        self.geometry("1180x880")
        self.minsize(940, 760)
        self.configure(bg="#f1f5f9")
        self.events = queue.Queue()
        self.busy = False
        self.running_check = False
        self.stop = threading.Event()
        self.database = None
        self.files = {}
        self.results = []
        self.result_database = None
        self.result_options = None
        self.serial = 0
        self.ignore_case = tk.BooleanVar(value=True)
        self.trim_spaces = tk.BooleanVar(value=True)
        self.filter_value = tk.StringVar(value="All rows")
        self.mode_value = tk.StringVar(value=UPLOADED_ROWS)
        self.coverage = tk.StringVar(value="Select a file to review the fields being checked.")
        self.status = tk.StringVar(value="Start by choosing your database. Files stay on this computer.")
        self.db_label = tk.StringVar(value="No database selected")
        self.metrics = tk.StringVar(value="0 files  •  0 passed  •  0 matched rows  •  0 changed  •  0 not found")
        self.action_widgets = []
        self._build()
        self.after(100, self._poll)
        self.protocol("WM_DELETE_WINDOW", self._close)

    def _button(self, parent, text, command, **kwargs):
        widget = ttk.Button(parent, text=text, command=command, **kwargs)
        self.action_widgets.append(widget)
        return widget

    def _build(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background="#f1f5f9")
        style.configure("TLabel", background="#f1f5f9", foreground="#0f172a", font=("Segoe UI", 10))
        style.configure("Title.TLabel", font=("Segoe UI", 24, "bold"))
        style.configure("Sub.TLabel", foreground="#475569")
        style.configure("TButton", font=("Segoe UI", 10), padding=(12, 7))
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=29)
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), padding=6)
        style.configure("TCheckbutton", background="#f1f5f9", font=("Segoe UI", 10))
        main = ttk.Frame(self, padding=24)
        main.pack(fill="both", expand=True)
        ttk.Label(main, text="CHCODE Database Checker", style="Title.TLabel").pack(anchor="w")
        ttk.Label(main, text="Validate the rows in multiple files against your database, using CHCODE.", style="Sub.TLabel").pack(anchor="w", pady=(3, 18))

        database_row = ttk.Frame(main)
        database_row.pack(fill="x")
        self._button(database_row, "1  Choose database", self._choose_database).pack(side="left")
        ttk.Label(database_row, textvariable=self.db_label, wraplength=650).pack(side="left", padx=14)
        db_sheet_row = ttk.Frame(main)
        db_sheet_row.pack(fill="x", pady=(8, 16))
        ttk.Label(db_sheet_row, text="Database worksheet").pack(side="left")
        self.db_sheet = ttk.Combobox(db_sheet_row, state="disabled", width=28)
        self.db_sheet.pack(side="left", padx=10)
        self.db_sheet.bind("<<ComboboxSelected>>", self._change_database_sheet)
        ttk.Label(db_sheet_row, text="Required header: CHCODE", style="Sub.TLabel").pack(side="left", padx=8)

        toolbar = ttk.Frame(main)
        toolbar.pack(fill="x", pady=(0, 8))
        self._button(toolbar, "2  Add files", self._add_files).pack(side="left")
        self._button(toolbar, "Remove selected", self._remove).pack(side="left", padx=8)
        self._button(toolbar, "Clear files", self._clear).pack(side="left")
        ttk.Label(toolbar, text="Excel (.xlsx / .xlsm) and CSV", style="Sub.TLabel").pack(side="right")

        self.file_tree = self._tree(main, [
            ("name", "File", 240), ("sheet", "Worksheet", 130), ("column", "CHCODE header", 120),
            ("rows", "Rows", 65), ("matched", "Matched", 75), ("different", "Changed", 75),
            ("extra", "Extra", 65), ("missing", "Missing", 75), ("blank", "Blank", 65), ("status", "Comparison", 200),
        ], height=6)
        self.file_tree.bind("<<TreeviewSelect>>", self._select_file)
        selection_row = ttk.Frame(main)
        selection_row.pack(fill="x", pady=8)
        ttk.Label(selection_row, text="Selected file worksheet").pack(side="left")
        self.file_sheet = ttk.Combobox(selection_row, state="disabled", width=28)
        self.file_sheet.pack(side="left", padx=10)
        self.file_sheet.bind("<<ComboboxSelected>>", self._change_file_sheet)
        self.mapping_button = self._button(selection_row, "Review field mappings", self._edit_mappings)
        self.mapping_button.pack(side="left")

        mode_row = ttk.Frame(main)
        mode_row.pack(fill="x", pady=(0, 6))
        ttk.Label(mode_row, text="Check mode").pack(side="left")
        self.mode_combo = ttk.Combobox(mode_row, textvariable=self.mode_value, values=[UPLOADED_ROWS, FULL_COPY], width=42, state="readonly")
        self.mode_combo.pack(side="left", padx=10)
        self.mode_combo.bind("<<ComboboxSelected>>", lambda event: self._invalidate())
        ttk.Label(main, textvariable=self.coverage, wraplength=1090, style="Sub.TLabel").pack(anchor="w", pady=(0, 4))

        options = ttk.Frame(main)
        options.pack(fill="x", pady=(4, 12))
        for label, variable in [("Ignore letter case", self.ignore_case), ("Trim outer spaces", self.trim_spaces)]:
            widget = ttk.Checkbutton(options, text=label, variable=variable, command=self._invalidate)
            widget.pack(side="left", padx=(0, 18))
            self.action_widgets.append(widget)
        ttk.Label(options, text="Options apply to every cell. Row order does not matter.", style="Sub.TLabel").pack(side="left")

        actions = ttk.Frame(main)
        actions.pack(fill="x")
        self.run_button = self._button(actions, "3  Check all files", self._run)
        self.run_button.pack(side="left")
        self.export_button = self._button(actions, "Export Excel report", self._export)
        self.export_button.pack(side="left", padx=8)
        self.cancel_button = ttk.Button(actions, text="Cancel", command=self.stop.set, state="disabled")
        self.cancel_button.pack(side="left")
        self.progress = ttk.Progressbar(actions, mode="determinate")
        self.progress.pack(side="right", fill="x", expand=True, padx=(20, 0))
        ttk.Label(main, textvariable=self.metrics, font=("Segoe UI", 11, "bold"), wraplength=1000).pack(anchor="w", pady=(16, 8))

        preview_bar = ttk.Frame(main)
        preview_bar.pack(fill="x", pady=(0, 5))
        ttk.Label(preview_bar, text="Selected file results", font=("Segoe UI", 10, "bold")).pack(side="left")
        self.filter = ttk.Combobox(preview_bar, textvariable=self.filter_value,
                                  values=["All rows", "MATCHED", "DIFFERENT", "NOT FOUND", "EXTRA ROW", "BLANK KEY", "MISSING FROM FILE"], width=21, state="readonly")
        self.filter.pack(side="right")
        self.filter.bind("<<ComboboxSelected>>", lambda event: self._preview())
        self.preview_tree = self._tree(main, [("row", "Source row", 90), ("code", "CHCODE", 250),
                                              ("result", "Result", 180), ("db", "Database row", 100),
                                              ("changes", "Differences (file value / database value)", 450)], height=6, expand=True)
        ttk.Label(main, text="Preview: up to 500 rows. Export includes every checked row, cell differences, and field coverage.", style="Sub.TLabel").pack(anchor="w", pady=(7, 3))
        ttk.Label(main, textvariable=self.status, wraplength=1030, style="Sub.TLabel").pack(anchor="w")
        self._refresh_controls()

    def _tree(self, parent, columns, height, expand=False):
        frame = ttk.Frame(parent)
        frame.pack(fill="both" if expand else "x", expand=expand)
        tree = ttk.Treeview(frame, columns=[item[0] for item in columns], show="headings", height=height)
        for key, label, width in columns:
            tree.heading(key, text=label)
            tree.column(key, width=width, minwidth=55)
        vertical = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        horizontal = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        return tree

    def _refresh_controls(self):
        for widget in self.action_widgets:
            widget.configure(state="disabled" if self.busy else "normal")
        self.run_button.configure(state="normal" if not self.busy and self.database and self.database['column'] and self.files else "disabled")
        self.export_button.configure(state="normal" if not self.busy and self.results else "disabled")
        self.cancel_button.configure(state="normal" if self.busy and self.running_check else "disabled")
        self.db_sheet.configure(state="readonly" if not self.busy and self.database else "disabled")
        self.file_sheet.configure(state="readonly" if not self.busy and len(self.file_tree.selection()) == 1 else "disabled")
        self.mode_combo.configure(state="disabled" if self.busy else "readonly")
        selected = self.file_tree.selection()
        can_map = (not self.busy and self.mode_value.get() == UPLOADED_ROWS and self.database
                   and self.database['column'] and len(selected) == 1
                   and self.files[selected[0]].get('column'))
        self.mapping_button.configure(state="normal" if can_map else "disabled")

    def _background(self, job, done, message, check=False):
        if self.busy:
            return
        self.busy = True
        self.running_check = check
        self.stop.clear()
        self.status.set(message)
        self._refresh_controls()

        def work():
            try:
                value = job()
                self.events.put(("done", done, value))
            except Cancelled:
                self.events.put(("cancelled",))
            except Exception as error:
                self.events.put(("error", str(error)))
        threading.Thread(target=work, daemon=True).start()

    def _poll(self):
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == "progress":
                    index, total, result = event[1:]
                    self.progress.configure(value=index, maximum=total)
                    self.status.set(f"Checked {index} of {total}: {result.source.path.name}")
                    continue
                self.busy = False
                if event[0] == "done":
                    event[1](event[2])
                elif event[0] == "cancelled":
                    self.status.set("Check cancelled. Run again to generate a complete report.")
                else:
                    self.status.set(event[1])
                    messagebox.showerror("Unable to complete action", event[1], parent=self)
                self._refresh_controls()
        except queue.Empty:
            pass
        self.after(100, self._poll)

    @staticmethod
    def _inspect(path, sheet=None, sheets=None):
        names = sheets or sheet_names(path)
        table = read_table(path, sheet or names[0])
        column = next((name for name in table.headers if name.casefold() == "chcode"), "")
        return {"path": Path(path), "sheet": table.sheet, "sheets": names,
                "column": column, "headers": table.headers, "header_row": table.header_row,
                "rows": len(table.rows), "error": "" if column else "Missing CHCODE header"}

    def _set_database(self, data):
        self.database = data
        for item in self.files.values():
            item.pop('mappings', None)
        self.db_label.set(str(data['path']))
        self.db_sheet.configure(values=data['sheets'])
        self.db_sheet.set(data['sheet'])
        self._invalidate()
        self.status.set(data['error'] or f"Database ready: CHCODE header found on row {data['header_row']}; {data['rows']:,} data rows. Add files to check.")
        if data['error']:
            messagebox.showwarning("CHCODE header required", "The selected database worksheet must have a CHCODE header. Headers below the first row are detected automatically. Choose another worksheet or correct the file.", parent=self)

    def _choose_database(self):
        path = filedialog.askopenfilename(title="Choose the CHCODE database", filetypes=FILE_TYPES, parent=self)
        if path:
            self._background(lambda: self._inspect(Path(path)), self._set_database, "Reading database…")

    def _change_database_sheet(self, event=None):
        data = self.database
        sheet = self.db_sheet.get()
        self._background(lambda: self._inspect(data['path'], sheet, data['sheets']), self._set_database, "Reading database worksheet…")

    def _add_files(self):
        paths = filedialog.askopenfilenames(title="Choose files to check (select multiple)", filetypes=FILE_TYPES, parent=self)
        existing = {str(data['path']).casefold() for data in self.files.values()}
        paths = [Path(path) for path in paths if str(Path(path)).casefold() not in existing]
        if not paths:
            return

        def inspect():
            items = []
            for path in paths:
                try:
                    items.append(self._inspect(path))
                except Exception as error:
                    items.append({"path": path, "sheet": "", "sheets": [], "column": "", "headers": [], "rows": 0, "error": str(error)})
            return items

        def done(items):
            for item in items:
                self.serial += 1
                key = str(self.serial)
                self.files[key] = item
                self.file_tree.insert("", "end", iid=key, values=self._file_values(item))
            self._invalidate()
            self.status.set(f"Added {len(items)} file(s). Files with missing CHCODE headers will be reported as errors.")
        self._background(inspect, done, "Reading selected files…")

    @staticmethod
    def _file_values(item):
        return (item['path'].name, item['sheet'], item['column'] or "Missing", item['rows'], "—", "—", "—", "—", "—", item['error'] or "Ready")

    def _change_file_sheet(self, event=None):
        selected = self.file_tree.selection()
        if len(selected) != 1:
            return
        key = selected[0]
        data = self.files[key]
        sheet = self.file_sheet.get()

        def done(item):
            self.files[key] = item
            self._invalidate()
            self._select_file()
            self.status.set(item['error'] or "Worksheet updated. Ready to check.")
        self._background(lambda: self._inspect(data['path'], sheet, data['sheets']), done, "Reading selected worksheet…")

    def _select_file(self, event=None):
        selected = self.file_tree.selection()
        if len(selected) == 1:
            item = self.files[selected[0]]
            self.file_sheet.configure(values=item['sheets'])
            self.file_sheet.set(item['sheet'])
            if item['error']:
                self.status.set(f"{item['path'].name}: {item['error']}")
        else:
            self.file_sheet.set("")
        self._preview()
        self._update_coverage()
        self._refresh_controls()

    def _edit_mappings(self):
        selected = self.file_tree.selection()
        if self.busy or len(selected) != 1 or not self.database:
            return
        item = self.files[selected[0]]

        def save(mappings, apply_similar):
            for key, other in self.files.items():
                if key == selected[0] or (apply_similar and other['headers'] == item['headers']):
                    other['mappings'] = mappings
            self._invalidate()
            self.status.set("Field mappings saved. Run the check again to update results.")
        MappingDialog(self, item, self.database, save)

    def _update_coverage(self):
        from checker import suggest_mappings
        if self.mode_value.get() == FULL_COPY:
            self.coverage.set("Checks all columns and all database rows, including duplicate counts.")
            return
        selected = self.file_tree.selection()
        if len(selected) != 1 or not self.database:
            self.coverage.set("Uploaded-row mode checks mapped fields only. Select a file to review field coverage.")
            return
        item = self.files[selected[0]]
        mappings = item.get('mappings')
        if mappings is None:
            mappings = suggest_mappings(item['headers'], self.database['headers'], item['column'], self.database['column'])
        checked = {mapping.file_column for mapping in mappings}
        ignored = [header for header in item['headers'] if header not in checked]
        self.coverage.set(f"{len(mappings)} fields checked. Not checked: {', '.join(ignored) or 'none'}.")

    def _remove(self):
        for key in self.file_tree.selection():
            self.file_tree.delete(key)
            del self.files[key]
        self._invalidate()
        self._select_file()

    def _clear(self):
        self.file_tree.delete(*self.file_tree.get_children())
        self.files.clear()
        self._invalidate()
        self._select_file()

    def _invalidate(self):
        self.results = []
        self.metrics.set(f"{len(self.files)} files  •  0 passed  •  0 matched rows  •  0 changed  •  0 not found")
        self.preview_tree.delete(*self.preview_tree.get_children())
        for key, item in self.files.items():
            self.file_tree.item(key, values=self._file_values(item))
        self.progress.configure(value=0)
        self._update_coverage()
        self._refresh_controls()

    @staticmethod
    def _source(item):
        return Source(item['path'], item['sheet'], item['column'] or "CHCODE", item.get('mappings'))

    def _run(self):
        self._invalidate()
        database = self._source(self.database)
        sources = [self._source(item) for item in self.files.values()]
        options = Options(self.ignore_case.get(), self.trim_spaces.get(), self.mode_value.get() == UPLOADED_ROWS)

        def progress(index, total, result):
            self.events.put(("progress", index, total, result))

        def done(results):
            self.results = results
            self.result_database = database
            self.result_options = options
            for key, result in zip(self.files, results):
                self.file_tree.item(key, values=(result.source.path.name, result.source.sheet, result.source.column,
                    len(result.records), result.count("MATCHED"), result.count("DIFFERENT"),
                    result.count("NOT FOUND") + result.count("EXTRA ROW"), len(result.missing_rows),
                    result.count("BLANK KEY"), result.error or result.verdict))
            same = sum(result.verdict in {"SAME ROWS", "MAPPED ROWS MATCH"} for result in results)
            counts = [sum(result.count(status) for result in results) for status in ["MATCHED", "DIFFERENT", "NOT FOUND", "EXTRA ROW", "MISSING FROM FILE"]]
            self.metrics.set(f"{len(results)} files  •  {same} passed  •  {counts[0]:,} matched rows  •  {counts[1]:,} changed  •  {counts[2] + counts[3]:,} extra / not found  •  {counts[4]:,} missing")
            errors = sum(bool(result.error) for result in results)
            self.status.set(f"Check complete. {errors} file error(s). Export the report for full details.")
            if not self.file_tree.selection() and self.files:
                self.file_tree.selection_set(next(iter(self.files)))
            self._preview()
        self._background(lambda: check_files(database, sources, options, progress, self.stop.is_set), done, "Comparing complete rows against the database…", check=True)

    def _preview(self):
        self.preview_tree.delete(*self.preview_tree.get_children())
        selected = self.file_tree.selection()
        if len(selected) != 1 or not self.results:
            return
        position = list(self.files).index(selected[0])
        result = self.results[position]
        if result.error:
            self.status.set(result.error)
            return
        if result.missing_columns or result.extra_columns:
            self.status.set(f"Column differences — missing: {', '.join(result.missing_columns) or 'none'}; extra: {', '.join(result.extra_columns) or 'none'}.")
        column = result.headers.index(result.source.column)
        changes = {}
        for number, database_row, name, actual, expected in result.differences:
            changes.setdefault(number, []).append(f"{name}: {actual!r} / {expected!r}")
        count = 0
        for number, status, matches, values in result.records:
            if self.filter_value.get() != "All rows" and status != self.filter_value.get():
                continue
            code = values[column]
            self.preview_tree.insert("", "end", values=(number, "" if code is None else code, status, matches, "; ".join(changes.get(number, []))))
            count += 1
            if count >= 500:
                break
        if count < 500 and self.filter_value.get() in {"All rows", "MISSING FROM FILE"}:
            database_column = result.database_headers.index(self.result_database.column)
            for number, values in result.missing_rows[:500 - count]:
                self.preview_tree.insert("", "end", values=("", values[database_column], "MISSING FROM FILE", number, "Database row is absent from this file"))

    def _export(self):
        path = filedialog.asksaveasfilename(title="Save complete Excel report", defaultextension=".xlsx",
                    initialfile="CHCODE_check_report.xlsx", filetypes=[("Excel workbook", "*.xlsx")], parent=self)
        if not path:
            return
        destination = Path(path).resolve()
        inputs = [self.result_database.path, *(result.source.path for result in self.results)]
        if any(destination == item.resolve() for item in inputs):
            messagebox.showerror("Choose a different name", "The report cannot overwrite your database or a file being checked.", parent=self)
            return
        self._background(lambda: export_report(destination, self.result_database, self.result_options, self.results),
                         lambda _: self.status.set(f"Report saved: {destination}"), "Saving complete Excel report…")

    def _close(self):
        if self.busy:
            messagebox.showinfo("Action in progress", "Please wait for the current action to finish. You can cancel a running check.", parent=self)
            return
        self.destroy()


if __name__ == "__main__":
    CheckerApp().mainloop()
