
import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

APP_TITLE = "Appraisal Analysis Tool — Stage 2.4 (Presets + Help)"

CONFIG_FILE = "presets.json"

# Preset colors (fixed)
PRESET_COLORS = {
    "Q1": "#4A90E2",  # Blue
    "Q2": "#50C878",  # Green
    "Q3": "#FFA500",  # Orange
    "Q4": "#9370DB",  # Purple
    "Q5": "#FF6347"   # Red
}

# Load and save presets
def load_presets():
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"Q1": {}, "Q2": {}, "Q3": {}, "Q4": {}, "Q5": {}}

def save_presets(presets):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(presets, f, indent=2)
    except Exception:
        pass

presets = load_presets()
current_preset = None
preset_buttons = {}

def highlight_preset(preset_name):
    for name, button in preset_buttons.items():
        if name == preset_name:
            button.config(relief="sunken", font=("Segoe UI", 9, "bold"))
        else:
            button.config(relief="raised", font=("Segoe UI", 9, "normal"))

def apply_preset(preset_name, entries):
    global current_preset
    current_preset = preset_name
    highlight_preset(preset_name)
    data = presets.get(preset_name, {})
    for field, entry in entries.items():
        entry.delete(0, tk.END)
        if field in data:
            entry.insert(0, str(data[field]))

def save_preset(preset_name, entries):
    if not preset_name:
        messagebox.showwarning("No preset selected", "Select Q1–Q5 first before saving.")
        return
    data = {}
    for field, entry in entries.items():
        val = entry.get().strip()
        if val:
            data[field] = val
    presets[preset_name] = data
    save_presets(presets)
    messagebox.showinfo("Preset Saved", f"Preset {preset_name} updated and saved.")

def clear_adjustments(entries):
    for entry in entries.values():
        entry.delete(0, tk.END)

def clear_all(subject_entries, adj_entries, case_entry, addr_entries):
    for entry in subject_entries.values():
        entry.delete(0, tk.END)
    for entry in adj_entries.values():
        entry.delete(0, tk.END)
    case_entry.delete(0, tk.END)
    for entry in addr_entries:
        entry.delete(0, tk.END)

def run_analysis():
    messagebox.showinfo("Run Analysis", "Stage 3 will perform regression and create the PDF.")

def pick_file(label):
    path = filedialog.askopenfilename(title="Select CSV file",
                                      filetypes=[("CSV Files","*.csv")])
    if path:
        label.config(text=f"Selected: {os.path.basename(path)}")

def show_help():
    help_text = """
Appraisal Analysis Tool — Instructions

1. Enter Subject Information at the top (Case #, Address, City, State, ZIP).
2. In the 'Subject Data' column, fill in details for GLA, Basement, Garage Bays, Lot Size, and rooms.
3. In the 'Adjustment Amounts' column, enter the dollar adjustments (per sf, per bay, per room).
4. Presets:
   - Q1–Q5 buttons load saved sets of adjustments.
   - After editing adjustments, click 'Save Preset' to update the currently selected one.
   - Presets are saved to presets.json and will persist when reopening the app.
   - 'Clear Adjustments' resets only the right column. 'Clear/Reset All' clears everything.
5. File Selection:
   - Browse for Market Export CSV and Lot Sales CSV files.
6. Run Analysis:
   - Click 'Run Analysis' to process the data, perform regression, and generate a PDF report.
7. Outliers:
   - The analysis removes sales more than 2 standard deviations from regression unless dataset < 12.
8. Notes:
   - All adjustments and methods are user-provided, based on appraisal best practices.
   - The tool ensures results are transparent, factual, and replicable.

"""
    help_window = tk.Toplevel()
    help_window.title("Help & Instructions")
    st = scrolledtext.ScrolledText(help_window, wrap="word", width=80, height=30)
    st.insert("1.0", help_text)
    st.configure(state="disabled")
    st.pack(padx=10, pady=10)

def main():
    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("1050x800")
    root.configure(bg="#F5F7FA")

    # Menu bar with Help
    menubar = tk.Menu(root)
    helpmenu = tk.Menu(menubar, tearoff=0)
    helpmenu.add_command(label="Instructions", command=show_help)
    menubar.add_cascade(label="Help", menu=helpmenu)
    root.config(menu=menubar)

    # Header
    header = tk.Frame(root, bg="#2F6DB3")
    header.pack(fill="x")
    tk.Label(header, text=APP_TITLE, bg="#2F6DB3", fg="white",
             font=("Segoe UI", 14, "bold"), padx=12, pady=10).pack(side="left")

    # Subject Information
    subj_info = tk.LabelFrame(root, text="Subject Information", padx=12, pady=12, bg="#F5F7FA")
    subj_info.pack(fill="x", padx=12, pady=10)
    tk.Label(subj_info, text="Case #:", bg="#F5F7FA").grid(row=0, column=0, sticky="w", pady=4)
    case_entry = tk.Entry(subj_info, width=20)
    case_entry.grid(row=0, column=1, padx=5, pady=4)

    tk.Label(subj_info, text="Address:", bg="#F5F7FA").grid(row=0, column=2, sticky="w", pady=4)
    addr_entry = tk.Entry(subj_info, width=40)
    addr_entry.grid(row=0, column=3, padx=5, pady=4)

    tk.Label(subj_info, text="City:", bg="#F5F7FA").grid(row=1, column=0, sticky="w", pady=4)
    city_entry = tk.Entry(subj_info, width=20)
    city_entry.grid(row=1, column=1, padx=5, pady=4)

    tk.Label(subj_info, text="State:", bg="#F5F7FA").grid(row=1, column=2, sticky="w", pady=4)
    state_entry = tk.Entry(subj_info, width=10)
    state_entry.grid(row=1, column=3, sticky="w", pady=4)

    tk.Label(subj_info, text="ZIP:", bg="#F5F7FA").grid(row=1, column=4, sticky="w", pady=4)
    zip_entry = tk.Entry(subj_info, width=10)
    zip_entry.grid(row=1, column=5, sticky="w", pady=4)

    # Subject Data + Adjustments side by side
    subj_adj_frame = tk.LabelFrame(root, text="Subject Data & Adjustment Amounts", padx=12, pady=12, bg="#F5F7FA")
    subj_adj_frame.pack(fill="x", padx=12, pady=10)

    subject_fields = ["GLA (sf)", "Basement (sf)", "Garage Bays", "Lot Size (sf)",
                      "Bedrooms", "Family Rooms", "Bathrooms", "Other Rooms"]
    adjustment_fields = ["GLA $/sf", "Basement $/sf", "Garage $/bay",
                         "Bedrooms $ each", "Family Rooms $ each",
                         "Bathrooms $ each", "Other Rooms $ each"]

    subject_entries = {}
    adj_entries = {}

    tk.Label(subj_adj_frame, text="Subject Data", bg="#D0E7FF",
             font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="we", padx=5, pady=5)
    tk.Label(subj_adj_frame, text="Adjustment Amounts", bg="#D6F5D6",
             font=("Segoe UI", 10, "bold")).grid(row=0, column=1, sticky="we", padx=5, pady=5)

    for i, (s_field, a_field) in enumerate(zip(subject_fields, adjustment_fields), start=1):
        tk.Label(subj_adj_frame, text=s_field+":", bg="#F5F7FA").grid(row=i, column=0, sticky="w", pady=3)
        e1 = tk.Entry(subj_adj_frame, width=12)
        e1.grid(row=i, column=0, sticky="e", padx=(120,5), pady=3)
        subject_entries[s_field] = e1

        tk.Label(subj_adj_frame, text=a_field+":", bg="#F5F7FA").grid(row=i, column=1, sticky="w", pady=3, padx=(10,0))
        e2 = tk.Entry(subj_adj_frame, width=12)
        e2.grid(row=i, column=1, sticky="e", padx=(180,5), pady=3)
        adj_entries[a_field] = e2

    # Preset buttons
    preset_frame = tk.Frame(subj_adj_frame, bg="#F5F7FA")
    preset_frame.grid(row=len(subject_fields)+1, column=0, columnspan=2, pady=10)

    for p in ["Q1", "Q2", "Q3", "Q4", "Q5"]:
        btn = tk.Button(preset_frame, text=p, width=6,
                        bg=PRESET_COLORS[p], fg="white",
                        command=lambda pn=p: apply_preset(pn, adj_entries))
        btn.pack(side="left", padx=5)
        preset_buttons[p] = btn

    tk.Button(preset_frame, text="Save Preset",
              command=lambda: save_preset(current_preset, adj_entries)).pack(side="left", padx=10)
    tk.Button(preset_frame, text="Clear Adjustments",
              command=lambda: clear_adjustments(adj_entries)).pack(side="left", padx=10)

    # File selection
    files_frame = tk.LabelFrame(root, text="File Selection", padx=12, pady=12, bg="#F5F7FA")
    files_frame.pack(fill="x", padx=12, pady=10)

    tk.Label(files_frame, text="Market Export CSV:", bg="#F5F7FA").grid(row=0, column=0, sticky="w", pady=4)
    market_label = tk.Label(files_frame, text="(not selected)", bg="#F5F7FA")
    market_label.grid(row=0, column=1, sticky="w", pady=4)
    tk.Button(files_frame, text="Browse", command=lambda: pick_file(market_label)).grid(row=0, column=2, pady=4)

    tk.Label(files_frame, text="Lot Sales CSV:", bg="#F5F7FA").grid(row=1, column=0, sticky="w", pady=4)
    lot_label = tk.Label(files_frame, text="(not selected)", bg="#F5F7FA")
    lot_label.grid(row=1, column=1, sticky="w", pady=4)
    tk.Button(files_frame, text="Browse", command=lambda: pick_file(lot_label)).grid(row=1, column=2, pady=4)

    # Action buttons
    btn_frame = tk.Frame(root, bg="#F5F7FA")
    btn_frame.pack(fill="x", pady=20)
    tk.Button(btn_frame, text="Run Analysis", bg="#2DA44E", fg="white",
              command=run_analysis).pack(side="right", padx=10)
    tk.Button(btn_frame, text="Clear / Reset All",
              command=lambda: clear_all(subject_entries, adj_entries, case_entry,
                                        [addr_entry, city_entry, state_entry, zip_entry])).pack(side="right", padx=10)

    root.mainloop()

if __name__ == "__main__":
    main()
