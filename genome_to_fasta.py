from Bio import SeqIO
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# --- Core parsing function ---
def parse_gbff(file_path, use_product):
    # Counters for repeated names
    product_counts = {}

    with open(file_path) as handle:
        first_record = next(SeqIO.parse(handle, "genbank"))
    definition = first_record.description
    outfile_base = re.sub(r",.*", "", definition).strip()
    fasta_file = os.path.join(os.path.dirname(file_path), outfile_base + ".fasta")

    with open(fasta_file, "w") as fasta_out:
        for record in SeqIO.parse(file_path, "genbank"):
            for feature in record.features:
                if feature.type == "CDS":
                    locus = feature.qualifiers.get("locus_tag", ["unknown_locus"])[0]
                    product = feature.qualifiers.get("product", ["unknown_product"])[0]
                    seq = feature.qualifiers.get("translation", [""])[0]
                    if not seq:
                        continue

                    if use_product:
                        # Remove commas
                        clean_name = product.replace(",", "")
                        # Count duplicates (case-sensitive)
                        key = clean_name
                        if key in product_counts:
                            product_counts[key] += 1
                            clean_name = f"{clean_name}_{product_counts[key]}"
                        else:
                            product_counts[key] = 1
                        header = re.sub(r"[ \t/-]+", "_", clean_name)  # Replace spaces, /, and - with _
                    else:
                        header = locus

                    fasta_out.write(f">{header}\n{seq}\n")

    print(f"FASTA file written: {fasta_file}")
    return fasta_file

# --- GUI setup ---
root = tk.Tk()
root.title("GBFF to FASTA Parser")
root.geometry("550x600")

selected_paths = []
file_buttons = []
use_product_var = tk.BooleanVar(value=False)

def check_all_selected():
    if all(selected_paths):
        run_button.pack(pady=20)
    else:
        run_button.forget()

def browse_file(idx):
    path = filedialog.askopenfilename(
        title=f"Select GBFF file {idx+1}",
        filetypes=[("GenBank files", "*.gbff *.gbk"), ("All files", "*.*")]
    )
    if path:
        selected_paths[idx] = path
        file_buttons[idx].config(text=os.path.basename(path))
        check_all_selected()

def load_file_inputs(count):
    for widget in file_frame.winfo_children():
        widget.destroy()
    file_buttons.clear()
    selected_paths.clear()
    for i in range(count):
        frame = tk.Frame(file_frame)
        frame.pack(pady=5)
        tk.Label(frame, text=f"File {i+1}:").pack(side="left", padx=5)
        btn = tk.Button(frame, text="Browse", command=lambda idx=i: browse_file(idx))
        btn.pack(side="left", padx=5)
        file_buttons.append(btn)
        selected_paths.append(None)
    check_all_selected()

def confirm_count():
    try:
        count = int(file_count_entry.get())
        if count < 1:
            raise ValueError
        load_file_inputs(count)
    except ValueError:
        messagebox.showerror("Invalid input", "Please enter an integer ≥ 1.")

def run_parser():
    print("Parsing selected files...")
    for path in selected_paths:
        out = parse_gbff(path, use_product_var.get())
        print(f"Parsed: {path}\nSaved FASTA: {out}")
    messagebox.showinfo("Done", "Parsing complete! Check your console for output paths.")

# --- Widgets ---
tk.Label(root, text="How many GBFF files to parse?").pack(pady=10)
file_count_entry = tk.Entry(root, justify="center")
file_count_entry.pack()
tk.Button(root, text="Confirm", command=confirm_count).pack(pady=5)

# Radio buttons for naming choice
tk.Label(root, text="Name sequences using:").pack(pady=10)
tk.Radiobutton(root, text="Gene loci", variable=use_product_var, value=False).pack()
tk.Radiobutton(root, text="Protein products", variable=use_product_var, value=True).pack()

file_frame = tk.Frame(root)
file_frame.pack(pady=10)

run_button = tk.Button(root, text="Run", bg="lightgreen", command=run_parser)

root.mainloop()
