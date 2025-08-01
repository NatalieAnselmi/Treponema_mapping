import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from tkinter import Tk, filedialog


def bin_gene_counts(values):
    bins = [(0, 5), (5, 10), (10, 20), (20, 30), (30, 40), (40, 50),
            (50, 60), (60, 70), (70, 80), (80, 90), (90, 100)]
    labels = ["(0, 5]", "(5, 10]", "(10, 20]", "(20, 30]", "(30, 40]",
              "(40, 50]", "(50, 60]", "(60, 70]", "(70, 80]", "(80, 90]", "(90, 100]"]
    bins.append((100, float("inf")))
    labels.append(">100")

    bin_counts = {label: 0 for label in labels}

    for val in values:
        for (low, high), label in zip(bins, labels):
            if low < val <= high:
                bin_counts[label] += 1
                break

    return bin_counts

def select_excel_file():
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select Excel File", filetypes=[("Excel files", "*.xlsx")]
    )
    return file_path

def extract_gene_summary_data(wb):
    gene_data = {}  # {gene: {health_status: [count, avg, sd]}}
    count_labels = {}
    health_statuses = []

    for ws in wb.worksheets:
        if ws.title in ("Summary", "Full_Data"):
            continue
        status = ws.title
        health_statuses.append(status)

        # Extract labels and values
        col_start = 3
        num_cols = ws.max_column

        gene_names = [ws.cell(row=1, column=col).value for col in range(col_start, num_cols + 1)]

        row_count = ws.max_row
        count_label = ws.cell(row=row_count - 2, column=2).value
        count_labels[status] = count_label

        # Read last 3 rows (count, avg, sd)
        counts_row = row_count - 2
        avg_row = row_count - 1
        sd_row = row_count

        for i, gene in enumerate(gene_names):
            if gene not in gene_data:
                gene_data[gene] = {}

            col = col_start + i
            count = ws.cell(row=counts_row, column=col).value
            avg = ws.cell(row=avg_row, column=col).value
            sd = ws.cell(row=sd_row, column=col).value

            # Sanitize and round
            count = count if isinstance(count, (int, float)) else 0
            avg = round(avg, 3) if isinstance(avg, (int, float)) and avg != 0 else 0
            sd = round(sd, 3) if isinstance(sd, (int, float)) and sd != 0 else 0

            gene_data[gene][status] = [count, avg, sd]

    return gene_data, count_labels, health_statuses


def write_summary_table(ws, gene_data, count_labels, health_statuses, start_col=1):
    ws.cell(row=1, column=start_col).value = "Gene"
    ws.cell(row=1, column=start_col).font = Font(bold=True)

    # Header rows
    col = start_col + 1
    for status in health_statuses:
        ws.cell(row=1, column=col).value = status
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 2)
        ws.cell(row=1, column=col).font = Font(bold=True)

        ws.cell(row=2, column=col).value = count_labels[status]
        ws.cell(row=2, column=col + 1).value = "Average"
        ws.cell(row=2, column=col + 2).value = "SD"

        for i in range(col, col + 3):
            ws.cell(row=2, column=i).font = Font(bold=True)

        col += 3

    # Write gene rows
    row = 3
    for gene in sorted(gene_data.keys()):
        ws.cell(row=row, column=start_col).value = gene
        ws.cell(row=row, column=start_col).font = Font(bold=True)

        col = start_col + 1
        for status in health_statuses:
            values = gene_data[gene].get(status, [0, 0, 0])
            for val in values:
                ws.cell(row=row, column=col).value = val
                col += 1
        row += 1

    return row  # Return the row after last summary data


def write_distribution_table(ws, gene_data, health_statuses, start_col):
    start_row = 1
    ws.cell(row=start_row, column=start_col).value = "# genes with reads"
    ws.cell(row=start_row, column=start_col).font = Font(bold=True)

    for i, status in enumerate(health_statuses):
        ws.cell(row=start_row, column=start_col + 1 + i).value = status
        ws.cell(row=start_row, column=start_col + 1 + i).font = Font(bold=True)

    status_gene_counts = {status: [] for status in health_statuses}
    for gene in gene_data:
        for status in health_statuses:
            count = gene_data[gene].get(status, [0])[0]
            if isinstance(count, (int, float)):
                status_gene_counts[status].append(count)

    bin_labels = list(bin_gene_counts([]).keys())
    for i, label in enumerate(bin_labels, start=start_row + 1):
        ws.cell(row=i, column=start_col).value = label
        ws.cell(row=i, column=start_col).font = Font(bold=True)

        for j, status in enumerate(health_statuses):
            bins = bin_gene_counts(status_gene_counts[status])
            ws.cell(row=i, column=start_col + 1 + j).value = bins[label]


def autosize_columns(ws):
    for col in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(col)].auto_size = True


def main():
    file_path = select_excel_file()
    if not file_path:
        print("No file selected.")
        return

    wb = openpyxl.load_workbook(file_path, data_only=True)
    summary_ws = wb.create_sheet("Summary")

    gene_data, count_labels, health_statuses = extract_gene_summary_data(wb)

    # Write summary table first (left side)
    last_summary_row = write_summary_table(summary_ws, gene_data, count_labels, health_statuses, start_col=1)

    # Write distribution table to the right of summary
    summary_col_width = 1 + 3 * len(health_statuses)
    write_distribution_table(summary_ws, gene_data, health_statuses, start_col=summary_col_width + 2)

    autosize_columns(summary_ws)

    wb.save(file_path)
    print(f"Summary and distribution table added to: {file_path}")


if __name__ == "__main__":
    main()
