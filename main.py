import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
import copy

def safe_number(value):
    """Convert value to a number if possible, otherwise return 0."""
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0

def copy_formatting(src_cell, dst_cell):
    """Copy the formatting from one cell to another."""
    if src_cell.has_style:
        dst_cell.font = copy.copy(src_cell.font)
        dst_cell.border = copy.copy(src_cell.border)
        dst_cell.fill = copy.copy(src_cell.fill)
        dst_cell.number_format = copy.copy(src_cell.number_format)
        dst_cell.protection = copy.copy(src_cell.protection)
        dst_cell.alignment = copy.copy(src_cell.alignment)

def process_excel_file(file_path):
    workbook = openpyxl.load_workbook(file_path, keep_vba=True)
    sheet = workbook.active

    # Assuming headers are at row 3, and you want to insert a row just below the last existing data row.
    last_row = sheet.max_row + 1
    sheet.insert_rows(idx=last_row)

    # Apply a simple formula to a specific cell in the new row, for example, to sum values of column A and B in the new row
    # Adjust the column letters and row index according to your needs
    sheet[f'A{last_row}'].value = f'=SUM(A{last_row-1}:B{last_row-1})'

    # No need to copy formatting for the new row in this specific request,
    # but you can add it here if needed, using the copy_formatting function defined above.

    # Save the modified file with a new name to avoid overwriting the original.
    new_file_path = file_path.replace('.xlsm', '_modified.xlsm')
    workbook.save(new_file_path)
    messagebox.showinfo("Success", f"File has been processed and saved as {new_file_path}")

def select_excel_file():
    tk.Tk().withdraw()  # We use withdraw to hide the empty tk window
    filename = filedialog.askopenfilename(filetypes=[("Excel Macro-Enabled Workbook", "*.xlsm")])
    if filename:
        print(f"Selected file: {filename}")
        process_excel_file(filename)

def create_gui():
    window = tk.Tk()
    window.title("Excel Macro-Enabled File Modifier")
    window.geometry("300x300")

    open_file_btn = tk.Button(window, text="Open Excel File", bg="green", command=select_excel_file)
    open_file_btn.pack(expand=True)

    window.mainloop()

create_gui()
