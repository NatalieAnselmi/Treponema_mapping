import pandas as pd
import numpy as np
import os
import tkinter as tk
from tkinter import filedialog, messagebox

PSEUDOCOUNT = 0.1
SHEETS = ["Healthy", "PD", "Stable PD", "Fluctuating PD", "Progressing PD"]

# ---------- Core analysis ----------
def analyze_file(file_path):
    dfs = {s: pd.read_excel(file_path, sheet_name=s) for s in SHEETS}

    # --- Top abundant ---
    top50_abundant, top100_abundant = {}, {}
    for s, df in dfs.items():
        abundance = df.iloc[:, 2:].sum(axis=0).sort_values(ascending=False)
        top50_abundant[s] = abundance.head(50).index.tolist()
        top100_abundant[s] = abundance.head(100).index.tolist()

    # --- Top variable across all sheets ---
    combined = pd.concat([df.iloc[:, 2:] for df in dfs.values()], axis=0)
    variability = combined.var(axis=0).sort_values(ascending=False)
    top50_var = variability.head(50).index.tolist()
    top100_var = variability.head(100).index.tolist()

    # --- Log fold-change vs Healthy ---
    means = {s: df.iloc[:, 2:].mean(axis=0) for s, df in dfs.items()}
    comps = ["PD", "Stable PD", "Fluctuating PD", "Progressing PD"]
    top50_vh, top100_vh = {}, {}
    for c in comps:
        log2fc = np.log2((means[c] + PSEUDOCOUNT) / (means["Healthy"] + PSEUDOCOUNT))
        top50_vh[f"{c}_Higher"] = log2fc.nlargest(50).index
        top50_vh[f"{c}_Lower"] = log2fc.nsmallest(50).index
        top100_vh[f"{c}_Higher"] = log2fc.nlargest(100).index
        top100_vh[f"{c}_Lower"] = log2fc.nsmallest(100).index

    # --- Save all to one file ---
    out_path = file_path.replace(".xlsx", "_Tops.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        # Top50s
        pd.DataFrame({
            "Healthy": pd.Series(top50_abundant["Healthy"]),
            "PD": pd.Series(top50_abundant["PD"]),
            "Stable PD": pd.Series(top50_abundant["Stable PD"]),
            "Fluctuating PD": pd.Series(top50_abundant["Fluctuating PD"]),
            "Progressing PD": pd.Series(top50_abundant["Progressing PD"]),
            "Most Variable": pd.Series(top50_var),
        }).to_excel(writer, sheet_name="Top50s", index=False)

        # Top100s
        pd.DataFrame({
            "Healthy": pd.Series(top100_abundant["Healthy"]),
            "PD": pd.Series(top100_abundant["PD"]),
            "Stable PD": pd.Series(top100_abundant["Stable PD"]),
            "Fluctuating PD": pd.Series(top100_abundant["Fluctuating PD"]),
            "Progressing PD": pd.Series(top100_abundant["Progressing PD"]),
            "Most Variable": pd.Series(top100_var),
        }).to_excel(writer, sheet_name="Top100s", index=False)

        # Top50vHealth
        pd.DataFrame({k: pd.Series(v) for k, v in top50_vh.items()}).to_excel(
            writer, sheet_name="Top50vHealth", index=False
        )
        # Top100vHealth
        pd.DataFrame({k: pd.Series(v) for k, v in top100_vh.items()}).to_excel(
            writer, sheet_name="Top100vHealth", index=False
        )

    print(f"Saved: {out_path}")

# ---------- GUI ----------
root = tk.Tk()
root.title("Gene Analysis")
root.geometry("500x500")

file_buttons = []
selected_paths = []

file_input_frame = tk.Frame(root)
file_input_frame.pack(pady=10)

run_button = tk.Button(root, text="Run Analysis", fg="blue",
                       command=lambda: [analyze_file(p) for p in selected_paths if p])

def check_all_selected():
    if all(selected_paths):
        run_button.pack(pady=20)
    else:
        run_button.forget()

def browse_file(idx):
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if path:
        selected_paths[idx] = path
        file_buttons[idx].config(text=os.path.basename(path))
        check_all_selected()

def load_file_inputs(count):
    for w in file_input_frame.winfo_children():
        w.destroy()
    file_buttons.clear()
    selected_paths.clear()
    for i in range(count):
        frame = tk.Frame(file_input_frame)
        frame.pack(pady=5)
        tk.Label(frame, text=f"File {i+1}:").pack(side="left", padx=5)
        btn = tk.Button(frame, text="Select", command=lambda idx=i: browse_file(idx))
        btn.pack(side="left", padx=5)
        file_buttons.append(btn)
        selected_paths.append(None)
    check_all_selected()

def get_file_count():
    try:
        count = int(file_count_entry.get())
        if count < 1:
            raise ValueError
        load_file_inputs(count)
    except ValueError:
        messagebox.showerror("Invalid input", "Please enter an integer ≥ 1.")

# Layout
tk.Label(root, text="How many files to analyze?").pack(pady=10)
file_count_entry = tk.Entry(root, justify="center")
file_count_entry.pack()
tk.Button(root, text="Confirm", command=get_file_count).pack(pady=5)

# Pack the frame for file selectors AFTER Confirm
file_input_frame = tk.Frame(root)
file_input_frame.pack(pady=10)

root.mainloop()
