from Bio import SeqIO
import re
import csv
import tkinter as tk
from tkinter import filedialog

# --- INPUT ---
root = tk.Tk()
root.withdraw()
gbff_file = filedialog.askopenfilename(
    title="Select gbff file",
)

# --- STEP 1: Get DEFINITION line (first record's description) ---
with open(gbff_file) as handle:
    first_record = next(SeqIO.parse(handle, "genbank"))
definition = first_record.description
# Example: "Treponema denticola ATCC 35405, complete genome"
outfile_base = re.sub(r",.*", "", definition).strip()
fasta_file = outfile_base + ".fasta"

# --- STEP 2: Extract locus_tag and protein sequences ---
with open(fasta_file, "w") as fasta_out: #, open(csv_file, "w", newline="") as csv_out:
    
    for record in SeqIO.parse(gbff_file, "genbank"):
        for feature in record.features:
            if feature.type == "CDS":
                locus = feature.qualifiers.get("locus_tag", [""])[0]
                gene = feature.qualifiers.get("gene", [""])[0]
                product = feature.qualifiers.get("product", [""])[0]
                protein_id = feature.qualifiers.get("protein_id", [""])[0]
                seq = feature.qualifiers.get("translation", [""])[0]

                # Write FASTA (skip if no sequence)
                if locus and seq:
                    fasta_out.write(f">{locus}\n{seq}\n")

print(f"FASTA file written: {fasta_file}")
