"""
Appraisal Analysis Tool — v0.4.1 (timestamped)
Fixes & Features:
- Q presets row moved directly below Subject/Adjustments section
- Header Mapper saves/loads plain strings (no widget objects)
- Header Mapper popup enlarged; Save button always visible
- Add/Remove Custom Fields fully working and stays within popup
- Entry fields: light yellow bg; focus highlight to light green
- Adjustment entries auto-format to $xxx,xxx.00 on focus-out
- Clear/Reset in red; Help editor & viewer retained
"""

import os
import csv
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, simpledialog
from datetime import datetime

# ---------- Version with timestamp ----------
__version__ = f"v0.4.1 — {datetime.now().strftime('%Y-%m-%d %H:%M')}"

# ---------- Persistence file names ----------
PRESETS_FILE = "presets.json"
HEADERS_MAP_FILE = "headers_map.json"
CUSTOM_FIELDS_FILE = "custom_fields.json"
HELP_CONTENT_FILE = "help_content.json"

# ---------- UI colors / styles ----------
PRESET_COLORS = {
    "Q1": "#4A90E2",  # Blue
    "Q2": "#50C878",  # Green
    "Q3": "#FFA500",  # Orange
    "Q4": "#9370DB",  # Purple
    "Q5": "#FF6347",  # Red,
}
ACTIVE_BG = "#3A5F8F"          # active preset bg
ACTIVE_FG = "black"            # active preset text color
ENTRY_BG = "#FFFACD"           # light yellow default entry bg
ENTRY_FOCUS_BG = "#E0FFE0"     # light green on focus
BAR_BG = "#2F6DB3"

# ---------- Logical fields (left column) ----------
BASE_SUBJECT_FIELDS = [
    "GLA (sf)",
    "Basement (sf)",
    "Garage Bays",
    "Lot/Acres",
    "Basement (Beds)",
    "Bathrooms",
    "Basement (Family)",
    "Basement (Other)",
    "Total Fireplaces",
]

# ---------- Adjustment fields (right column) ----------
BASE_ADJ_FIELDS = [
    "GLA $/sf",
    "Basement $/sf",
    "Garage $/bay",
    "Bedrooms $ each",
    "Bathrooms $ each",
    "Family Rooms $ each",
    "Other Rooms $ each",
    "Fireplaces $ each",
]

# ---------- Default help content ----------
DEFAULT_HELP = {
    "general": (
        "This tool collects Subject data and Adjustment amounts, lets you save multiple preset sets (Q1–Q5), "
        "and maps your CSV headers to the fields used in analysis.\n\n"
        "Workflow:\n"
        "1) Enter Subject info (Case, Address).\n"
        "2) Fill or load adjustments (Q1–Q5). Click a Q button to load; Save Preset to store.\n"
        "3) Select Market CSV and/or Lot CSV.\n"
        "4) Configure → Set File Headers… to map logical fields to your CSV columns. "
        "Dropdowns support type-ahead and show 2–3 sample values.\n"
        "5) Run Analysis (future stage) for regression, outliers, and PDF.\n\n"
        "Local files: presets.json, headers_map.json, custom_fields.json, help_content.json"
    ),
    "subject": "Enter the subject’s address at the top. In Subject/Adjustments, left is data (counts/sizes) and right is $ per unit.",
    "files": "Choose your CSVs, then Set File Headers… to map each field. Use the sample preview to confirm format.",
}

# ---------- Utility: JSON load/save ----------
def load_json(path, default):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default

def save_json(path, data):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception as e:
        messagebox.showwarning("Save Error", f"Could not save to {path}\n{e}")

# ---------- Utility: currency formatting ----------
def to_currency_string(text: str) -> str:
    """
    Convert a string like '82', '82000', '44.5' to '$82.00', '$82,000.00', '$44.50'.
    If not parseable, return original text unchanged.
    """
    t = (text or "").strip().replace("$", "").replace(",", "")
    if not t:
        return ""
    try:
        num = float(t)
        return f"${num:,.2f}"
    except ValueError:
        return text


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Appraisal Analysis Tool — {__version__}")
        self.geometry("1200x900")
        self.configure(bg="#F5F7FA")

        # ----- State -----
        self.presets = load_json(PRESETS_FILE, {"Q1": {}, "Q2": {}, "Q3": {}, "Q4": {}, "Q5": {}})
        self.headers_map = load_json(HEADERS_MAP_FILE, {})        # logical_field -> header string
        self.custom_fields = load_json(CUSTOM_FIELDS_FILE, [])    # list of {"name":..., "header":...}
        self.help_content = load_json(HELP_CONTENT_FILE, DEFAULT_HELP.copy())

        self.active_preset = None
        self.subject_entries = {}
        self.adjustment_entries = {}

        self.market_csv_path = None
        self.lot_csv_path = None
        self.cached_headers = []
        self.headers_samples = {}

        self._mapper_widgets = {}     # transient: logical_field -> Combobox widget (not persisted)
        self._custom_widgets = []     # transient: list of (name_entry, header_combobox)

        # ----- Build UI -----
        self._build_menus()
        self._build_header_bar()
        self._build_subject_info_row()
        self._build_subject_adjustments_grid()
        self._build_presets_row()     # << moved here: appears right under Subject/Adjustments
        self._build_files_section()
        self._build_actions_row()

    # ---------- Styled Entry ----------
    def styled_entry(self, parent, width=14):
        e = tk.Entry(parent, width=width, bg=ENTRY_BG, relief="sunken", bd=2)
        e.bind("<FocusIn>", lambda ev: e.config(bg=ENTRY_FOCUS_BG))
        e.bind("<FocusOut>", lambda ev: e.config(bg=ENTRY_BG))
        return e

    # ---------- Header Bar ----------
    def _build_header_bar(self):
        bar = tk.Frame(self, bg=BAR_BG)
        bar.pack(fill="x")
        tk.Label(
            bar, text=f"Appraisal Analysis Tool — {__version__}",
            bg=BAR_BG, fg="white", font=("Segoe UI", 14, "bold"),
            padx=12, pady=10
        ).pack(side="left")

    # ---------- Subject Info (Case/Address) ----------
    def _build_subject_info_row(self):
        box = tk.LabelFrame(self, text="Subject Information", padx=12, pady=12, bg="#F5F7FA")
        box.pack(fill="x", padx=12, pady=10)

        tk.Label(box, text="Case #:", bg="#F5F7FA").grid(row=0, column=0, sticky="w", pady=4)
        self.case_entry = self.styled_entry(box, width=24)
        self.case_entry.grid(row=0, column=1, padx=5, pady=4)

        tk.Label(box, text="Address:", bg="#F5F7FA").grid(row=0, column=2, sticky="w", pady=4)
        self.addr_entry = self.styled_entry(box, width=36)
        self.addr_entry.grid(row=0, column=3, padx=5, pady=4)

        tk.Label(box, text="City:", bg="#F5F7FA").grid(row=1, column=0, sticky="w", pady=4)
        self.city_entry = self.styled_entry(box, width=20)
        self.city_entry.grid(row=1, column=1, padx=5, pady=4)

        tk.Label(box, text="State:", bg="#F5F7FA").grid(row=1, column=2, sticky="w", pady=4)
        self.state_entry = self.styled_entry(box, width=10)
        self.state_entry.grid(row=1, column=3, sticky="w", pady=4)

        tk.Label(box, text="ZIP:", bg="#F5F7FA").grid(row=1, column=4, sticky="w", pady=4)
        self.zip_entry = self.styled_entry(box, width=10)
        self.zip_entry.grid(row=1, column=5, sticky="w", pady=4)

    # ---------- Subject + Adjustments Grid ----------
    def _build_subject_adjustments_grid(self):
        frame = tk.LabelFrame(self, text="Subject Data & Adjustment Amounts", padx=12, pady=12, bg="#F5F7FA")
        frame.pack(fill="x", padx=12, pady=10)

        tk.Label(frame, text="Subject Data", bg="#D0E7FF", font=("Segoe UI", 10, "bold")).grid(
            row=0, column=0, sticky="we", padx=5, pady=5
        )
        tk.Label(frame, text="Adjustment Amounts", bg="#D6F5D6", font=("Segoe UI", 10, "bold")).grid(
            row=0, column=1, sticky="we", padx=5, pady=5
        )

        all_subject_fields = BASE_SUBJECT_FIELDS + [c.get("name", "") for c in self.custom_fields if c.get("name")]
        all_adj_fields = BASE_ADJ_FIELDS + [
            c.get("adj_label", f'{c.get("name", "Custom")} $ each') for c in self.custom_fields
        ]

        max_rows = max(len(all_subject_fields), len(all_adj_fields))
        for i in range(max_rows):
            if i < len(all_subject_fields):
                s_field = all_subject_fields[i]
                tk.Label(frame, text=s_field + ":", bg="#F5F7FA").grid(row=i+1, column=0, sticky="w", pady=3)
                s_entry = self.styled_entry(frame, width=14)
                s_entry.grid(row=i+1, column=0, sticky="e", padx=(130, 5), pady=3)
                self.subject_entries[s_field] = s_entry

            if i < len(all_adj_fields):
                a_field = all_adj_fields[i]
                tk.Label(frame, text=a_field + ":", bg="#F5F7FA").grid(row=i+1, column=1, sticky="w", pady=3)
                a_entry = self.styled_entry(frame, width=14)

                def on_focus_out(event, e=a_entry):
                    e.insert(0, to_currency_string(e.get()))
                    val = e.get().replace("$$", "$")
                    e.delete(0, tk.END)
                    e.insert(0, val)

                a_entry.bind("<FocusOut>", on_focus_out)
                a_entry.grid(row=i+1, column=1, sticky="e", padx=(200, 5), pady=3)
                self.adjustment_entries[a_field] = a_entry

    # ---------- Preset Buttons (now directly under Subject/Adjustments) ----------
    def _build_presets_row(self):
        row = tk.Frame(self, bg="#F5F7FA")
        row.pack(fill="x", pady=(0, 10))

        self.preset_buttons = {}
        colors = {"Q1": "#FFCCCC", "Q2": "#FFE5B4", "Q3": "#FFFFCC", "Q4": "#CCFFCC", "Q5": "#CCE5FF"}
        for i, q in enumerate(["Q1", "Q2", "Q3", "Q4", "Q5"]):
            btn = tk.Button(
                row, text=q,
                bg=colors[q], fg="black",
                relief="raised", width=6,
                command=lambda q=q: self._apply_preset(q)
            )
            btn.grid(row=0, column=i, padx=6)
            self.preset_buttons[q] = btn

        tk.Button(row, text="Save Preset", command=self._save_preset).grid(row=0, column=6, padx=10)
        tk.Button(row, text="Clear Adjustments", command=self._clear_adjustments).grid(row=0, column=7, padx=10)

    def _apply_preset(self, name):
        data = self.presets.get(name, {})
        for field, entry in self.adjustment_entries.items():
            entry.delete(0, tk.END)
            if field in data:
                entry.insert(0, data[field])

        for q, btn in self.preset_buttons.items():
            btn.config(relief="raised", bd=1)
        self.preset_buttons[name].config(relief="sunken", fg="black", bd=3)
        self.active_preset = name

    def _save_preset(self):
        if not self.active_preset:
            messagebox.showinfo("Save Preset", "Select a Q button first.")
            return
        data = {}
        for field, entry in self.adjustment_entries.items():
            data[field] = entry.get()
        self.presets[self.active_preset] = data
        save_json(PRESETS_FILE, self.presets)
        messagebox.showinfo("Preset Saved", f"{self.active_preset} saved.")

    def _clear_adjustments(self):
        for entry in self.adjustment_entries.values():
            entry.delete(0, tk.END)
        if self.active_preset:
            self.preset_buttons[self.active_preset].config(relief="raised", bd=1)
            self.active_preset = None

    # ---------- Files Section ----------
    def _build_files_section(self):
        box = tk.LabelFrame(self, text="File Selection", padx=12, pady=12, bg="#F5F7FA")
        box.pack(fill="x", padx=12, pady=10)

        tk.Label(box, text="Market Export CSV:", bg="#F5F7FA").grid(row=0, column=0, sticky="w", pady=4)
        self.market_label = tk.Label(box, text="(not selected)", bg="#F5F7FA")
        self.market_label.grid(row=0, column=1, sticky="w", pady=4)
        tk.Button(box, text="Browse", command=lambda: self._pick_file("market")).grid(row=0, column=2, pady=4)

        tk.Label(box, text="Lot Sales CSV:", bg="#F5F7FA").grid(row=1, column=0, sticky="w", pady=4)
        self.lot_label = tk.Label(box, text="(not selected)", bg="#F5F7FA")
        self.lot_label.grid(row=1, column=1, sticky="w", pady=4)
        tk.Button(box, text="Browse", command=lambda: self._pick_file("lot")).grid(row=1, column=2, pady=4)

        tk.Button(box, text="Set File Headers…", command=self._open_header_mapper).grid(row=2, column=0, pady=10, sticky="w")

    # ---------- File Picking / Header Cache ----------
    def _pick_file(self, kind):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")])
        if not path:
            return
        if kind == "market":
            self.market_csv_path = path
            self.market_label.config(text=os.path.basename(path))
            self._cache_headers(path)
        elif kind == "lot":
            self.lot_csv_path = path
            self.lot_label.config(text=os.path.basename(path))
            self._cache_headers(path)

    def _cache_headers(self, path):
        try:
            with open(path, newline='', encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                self.cached_headers = reader.fieldnames or []
                samples = {h: [] for h in self.cached_headers}
                for i, row in enumerate(reader):
                    if i >= 3:
                        break
                    for h, val in row.items():
                        samples[h].append("" if val is None else str(val))
                self.headers_samples = samples
        except Exception as e:
            messagebox.showerror("Header Cache Error", str(e))

    # ---------- Header Mapper ----------
    def _open_header_mapper(self):
        if not self.cached_headers:
            messagebox.showinfo("No headers", "Please select at least one CSV so I can read its headers.")
            return

        self._mapper_widgets = {}   # reset transient widgets map
        self._custom_widgets = []   # reset custom field widgets

        win = tk.Toplevel(self)
        win.title("Set File Headers")
        win.geometry("820x600")  # larger popup so buttons are visible
        win.transient(self)

        container = tk.Frame(win)
        container.pack(fill="both", expand=True, padx=10, pady=10)

        canvas = tk.Canvas(container, highlightthickness=0)
        vscroll = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas)
        inner_id = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=vscroll.set)

        def _on_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        inner.bind("<Configure>", _on_configure)

        canvas.pack(side="left", fill="both", expand=True)
        vscroll.pack(side="right", fill="y")

        row = 0
        tk.Label(inner, text="Map base fields to your CSV headers (type-ahead supported).").grid(
            row=row, column=0, columnspan=3, sticky="w", pady=(0,8)
        )
        row += 1

        # Base fields (Subject + Adjustment fields names for mapping)
        for field in BASE_SUBJECT_FIELDS + BASE_ADJ_FIELDS:
            tk.Label(inner, text=field).grid(row=row, column=0, sticky="w", padx=4, pady=3)
            combo = ttk.Combobox(inner, values=self.cached_headers, width=38)
            prev_value = self.headers_map.get(field, "")
            if prev_value:
                combo.set(prev_value)
            combo.grid(row=row, column=1, padx=4, pady=3, sticky="w")

            preview = tk.Label(inner, text="", fg="gray")
            preview.grid(row=row, column=2, sticky="w")

            combo.bind(
                "<KeyRelease>",
                lambda ev, c=combo, p=preview: self._typeahead_and_preview(c, p)
            )
            combo.bind(
                "<<ComboboxSelected>>",
                lambda ev, c=combo, p=preview: self._update_preview_from_combo(c, p)
            )
            # initialize preview if preset exists
            self._update_preview_from_combo(combo, preview)

            self._mapper_widgets[field] = combo
            row += 1

        # ---- Custom Fields section ----
        tk.Label(inner, text="Custom Fields", font=("Segoe UI", 10, "bold")).grid(
            row=row, column=0, sticky="w", pady=(12, 4)
        )
        row += 1

        for cf in self.custom_fields:
            name = cf.get("name", "")
            header = cf.get("header", "")
            n_entry = tk.Entry(inner, width=20, bg="#E6FFED")  # light green so it's obvious
            n_entry.insert(0, name)
            n_entry.grid(row=row, column=0, padx=4, pady=2, sticky="w")
            c_combo = ttk.Combobox(inner, values=self.cached_headers, width=38)
            c_combo.set(header)
            c_combo.grid(row=row, column=1, padx=4, pady=2, sticky="w")
            self._custom_widgets.append((n_entry, c_combo))
            row += 1

        def add_custom_field():
            n_entry = tk.Entry(inner, width=20, bg="#E6FFED")
            n_entry.grid()
            c_combo = ttk.Combobox(inner, values=self.cached_headers, width=38)
            c_combo.grid(row=n_entry.grid_info()['row'], column=1, padx=4, pady=2, sticky="w")
            self._custom_widgets.append((n_entry, c_combo))

        def remove_custom_field():
            if not self._custom_widgets:
                return
            n_entry, c_combo = self._custom_widgets.pop()
            n_entry.destroy()
            c_combo.destroy()

        tk.Button(inner, text="Add Custom Field", bg="#28A745", fg="white",
                  command=add_custom_field).grid(row=row, column=0, pady=10, sticky="w")
        tk.Button(inner, text="Remove Custom Field", bg="#D73A49", fg="white",
                  command=remove_custom_field).grid(row=row, column=1, pady=10, sticky="w")
        row += 1

        # ---- Save row (pins at end of scroll content) ----
        def save_all():
            # collect headers from widgets into plain strings
            final_map = {}
            for field, combo in self._mapper_widgets.items():
                if isinstance(combo, ttk.Combobox):
                    final_map[field] = combo.get().strip()
            save_json(HEADERS_MAP_FILE, final_map)
            self.headers_map = final_map

            # collect custom fields into plain list of dicts
            new_custom_fields = []
            for n_entry, c_combo in self._custom_widgets:
                name = n_entry.get().strip()
                header = c_combo.get().strip()
                if name:
                    new_custom_fields.append({"name": name, "header": header})
            save_json(CUSTOM_FIELDS_FILE, new_custom_fields)
            self.custom_fields = new_custom_fields

            messagebox.showinfo("Saved", "Headers and custom fields saved.")
            win.destroy()

        tk.Button(inner, text="Save All", bg="#0366D6", fg="white", command=save_all).grid(
            row=row, column=0, pady=12, sticky="w"
        )

    def _typeahead_and_preview(self, combo: ttk.Combobox, preview_label: tk.Label):
        typed = combo.get().strip()
        if not self.cached_headers:
            preview_label.config(text="")
            return
        if typed:
            filtered = [h for h in self.cached_headers if typed.lower() in h.lower()]
            combo["values"] = filtered or self.cached_headers
        else:
            combo["values"] = self.cached_headers
        self._update_preview_from_combo(combo, preview_label)

    def _update_preview_from_combo(self, combo: ttk.Combobox, preview_label: tk.Label):
        h = combo.get().strip()
        vals = self.headers_samples.get(h, [])[:3]
        if vals:
            preview_label.config(text=" | ".join(vals))
        else:
            preview_label.config(text="")

    # ---------- Help System ----------
    def _open_help_editor(self):
        win = tk.Toplevel(self)
        win.title("Edit Help Content")
        win.geometry("700x520")
        win.transient(self)

        frame = tk.Frame(win)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        self._help_widgets = {}
        row = 0
        for section in ["general", "subject", "files"]:
            tk.Label(frame, text=section, font=("Segoe UI", 10, "bold")).grid(
                row=row, column=0, sticky="w", pady=(8, 2)
            )
            txt = tk.Text(frame, width=84, height=6, wrap="word")
            txt.insert("1.0", self.help_content.get(section, ""))
            txt.grid(row=row+1, column=0, padx=4, pady=4)
            self._help_widgets[section] = txt
            row += 2

        def save_help():
            for sec, widget in self._help_widgets.items():
                self.help_content[sec] = widget.get("1.0", "end-1c")
            save_json(HELP_CONTENT_FILE, self.help_content)
            messagebox.showinfo("Help Saved", "Help content saved.")
            win.destroy()

        tk.Button(frame, text="Save Help", bg="#0366D6", fg="white", command=save_help).grid(
            row=row, column=0, pady=12, sticky="w"
        )

    def _show_help_popup(self):
        win = tk.Toplevel(self)
        win.title("Help")
        win.geometry("700x500")
        win.transient(self)

        txt = tk.Text(win, wrap="word")
        txt.pack(fill="both", expand=True, padx=10, pady=10)
        txt.tag_configure("section", font=("Segoe UI", 10, "bold"))

        for section in ["general", "subject", "files"]:
            txt.insert("end", f"{section}\n", "section")
            txt.insert("end", self.help_content.get(section, "") + "\n\n")

    # ---------- Clear / Reset & Analysis Placeholder ----------
    def _clear_all(self):
        self.case_entry.delete(0, tk.END)
        self.addr_entry.delete(0, tk.END)
        self.city_entry.delete(0, tk.END)
        self.state_entry.delete(0, tk.END)
        self.zip_entry.delete(0, tk.END)

        for e in self.subject_entries.values():
            e.delete(0, tk.END)
        for e in self.adjustment_entries.values():
            e.delete(0, tk.END)

        self.market_csv_path = None
        self.lot_csv_path = None
        self.market_label.config(text="(not selected)")
        self.lot_label.config(text="(not selected)")

        if hasattr(self, "preset_buttons"):
            for btn in self.preset_buttons.values():
                btn.config(relief="raised", bd=1)
        self.active_preset = None

    def _run_analysis_placeholder(self):
        msg = []
        msg.append(f"Case #: {self.case_entry.get()}")
        msg.append(f"Address: {self.addr_entry.get()}, {self.city_entry.get()} {self.state_entry.get()} {self.zip_entry.get()}")
        msg.append(f"Market file: {self.market_csv_path or '(none)'}")
        msg.append(f"Lot file: {self.lot_csv_path or '(none)'}")
        adjustments = {f: e.get() for f, e in self.adjustment_entries.items()}
        msg.append(f"Adjustments: {adjustments}")
        messagebox.showinfo("Run Analysis (Demo)", "\n".join(msg))

    # ---------- Menus ----------
    def _build_menus(self):
        menubar = tk.Menu(self)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Set File Headers…", command=self._open_header_mapper)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.destroy)
        menubar.add_cascade(label="File", menu=file_menu)

        config_menu = tk.Menu(menubar, tearoff=0)
        config_menu.add_command(label="Edit Help Content…", command=self._open_help_editor)
        menubar.add_cascade(label="Configure", menu=config_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="View Help", command=self._show_help_popup)
        help_menu.add_command(label="About", command=self._show_about)
        menubar.add_cascade(label="Help", menu=help_menu)

        self.config(menu=menubar)

    def _show_about(self):
        messagebox.showinfo(
            "About",
            f"Appraisal Analysis Tool\nVersion: {__version__}\n\n"
            "Subject data, adjustments, header mapping, presets, custom fields, and help editor."
        )

    # ---------- Bottom Action Row ----------
    def _build_actions_row(self):
        row = tk.Frame(self, bg="#F5F7FA")
        row.pack(fill="x", pady=20)
        tk.Button(row, text="Clear / Reset All", bg="#D73A49", fg="white", command=self._clear_all).pack(side="right", padx=10)
        tk.Button(row, text="Run Analysis", bg="#2DA44E", fg="white", command=self._run_analysis_placeholder).pack(side="right", padx=10)


if __name__ == "__main__":
    app = App()
    app.mainloop()
