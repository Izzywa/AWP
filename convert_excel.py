import tkinter as tk
from tkinter import filedialog
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
root = tk.Tk()
filename = filedialog.askopenfilename(
    initialdir=BASE_DIR,
    title='Select an Excel File to READ',
    filetypes=[("Excel Files", "*.xlsx")]
)

output_filename = filedialog.askopenfilename(
    initialdir=BASE_DIR,
    title='Select OUTPUT excel file, press CANCEL if none',
    filetypes=[("Excel Files", "*.xlsx")]
)
root.destroy()
print(filename)
print(output_filename == '')
