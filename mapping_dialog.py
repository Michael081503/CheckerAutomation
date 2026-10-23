"""Review which uploaded columns are checked against database columns."""

import tkinter as tk
from tkinter import ttk

from checker import ColumnMapping, suggest_mappings

SKIP = "(Do not check)"


class MappingDialog(tk.Toplevel):
    def __init__(self, parent, item, database, save):
        super().__init__(parent)
        self.title("Review field mappings")
        self.geometry("900x650")
        self.minsize(760, 480)
        self.transient(parent)
        self.grab_set()
        self.item = item
        self.database = database
        self.save_callback = save
        self.fields = []
        self.apply_similar = tk.BooleanVar(value=True)
        self.coverage = tk.StringVar()
        main = ttk.Frame(self, padding=18)
        main.pack(fill="both", expand=True)
        ttk.Label(main, text=item['path'].name, font=("Segoe UI", 12, "bold"), wraplength=830).pack(anchor="w")
        ttk.Label(main, text="Only mapped fields are validated. Unchecked fields are listed in the report.", wraplength=830).pack(anchor="w", pady=(5, 12))
        area = ttk.Frame(main)
        area.pack(fill="both", expand=True)
        canvas = tk.Canvas(area, highlightthickness=0, bg="#f1f5f9")
        scroll = ttk.Scrollbar(area, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        fields = ttk.Frame(canvas)
        window = canvas.create_window((0, 0), window=fields, anchor="nw")
        fields.bind("<Configure>", lambda event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(window, width=event.width))
        for column, label in enumerate(["Uploaded file column", "Database column", "Compare as"]):
            ttk.Label(fields, text=label, font=("Segoe UI", 10, "bold")).grid(row=0, column=column, sticky="w", padx=5, pady=8)
        fields.columnconfigure(1, weight=1)
        for row, name in enumerate(item['headers'], 1):
            ttk.Label(fields, text=name).grid(row=row, column=0, sticky="w", padx=5, pady=5)
            destination = ttk.Combobox(fields, values=[SKIP, *database['headers']], state="readonly", width=40)
            destination.grid(row=row, column=1, sticky="ew", padx=5, pady=5)
            comparison = ttk.Combobox(fields, values=["Text", "Number", "Product"], state="readonly", width=10)
            comparison.grid(row=row, column=2, padx=5, pady=5)
            destination.bind("<<ComboboxSelected>>", self._coverage)
            self.fields.append((name, destination, comparison))
        ttk.Label(main, textvariable=self.coverage, font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(10, 4))
        ttk.Label(main, text="Number ignores commas. Product recognizes Credit Card / CARDS and Personal Loan / LOAN.", wraplength=830).pack(anchor="w")
        ttk.Checkbutton(main, text="Apply to other uploaded files with the same headers and lookup column", variable=self.apply_similar).pack(anchor="w", pady=8)
        footer = ttk.Frame(main)
        footer.pack(fill="x", pady=(8, 0))
        ttk.Button(footer, text="Use suggested mappings", command=self._suggest).pack(side="left")
        ttk.Button(footer, text="Save mappings", command=self._save).pack(side="right")
        ttk.Button(footer, text="Cancel", command=self.destroy).pack(side="right", padx=8)
        initial = item.get('mappings')
        self._populate(initial if initial is not None else self._suggested())

    def _suggested(self):
        return suggest_mappings(self.item['headers'], self.database['headers'], self.item['column'], self.database['column'])

    def _suggest(self):
        self._populate(self._suggested())

    def _populate(self, mappings):
        by_column = {item.file_column: item for item in mappings}
        for name, destination, comparison in self.fields:
            item = by_column.get(name)
            destination.set(item.database_column if item else SKIP)
            comparison.set(item.comparison if item else "Text")
            if name == self.item['column']:
                destination.set(self.database['column'])
                comparison.set("Text")
                destination.configure(state="disabled")
                comparison.configure(state="disabled")
        self._coverage()

    def _coverage(self, event=None):
        checked = sum(destination.get() != SKIP for _, destination, _ in self.fields)
        self.coverage.set(f"{checked} fields checked; {len(self.fields) - checked} fields not checked.")

    def _save(self):
        mappings = tuple(ColumnMapping(name, destination.get(), comparison.get())
                         for name, destination, comparison in self.fields if destination.get() != SKIP)
        self.save_callback(mappings, self.apply_similar.get())
        self.destroy()
