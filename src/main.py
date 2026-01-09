import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from data_cleaning import clean_data
from analysis import generate_summary
from report_generator import format_excel
import os
import sys


def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def run_automation():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx")]
    )

    if not file_path:
        return

    try:
        df = clean_data(file_path)
        summary, top = generate_summary(df)

        base_dir = get_base_dir()
        output_dir = os.path.join(base_dir, "output_files")
        os.makedirs(output_dir, exist_ok=True)

        output_path = os.path.join(output_dir, "final_report.xlsx")

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Clean Data", index=False)
            summary.to_excel(writer, sheet_name="Summary", index=False)
            top.to_excel(writer, sheet_name="Top 5 Transactions", index=False)

        format_excel(output_path)

        messagebox.showinfo("Success", f"Report generated successfully!\n\nSaved in:\n{output_path}")

    except Exception as e:
        messagebox.showerror("Error", str(e))


app = tk.Tk()
app.title("Excel Report Automation")
app.geometry("420x220")
app.resizable(False, False)

label = tk.Label(app, text="Select an Excel file to generate report", font=("Arial", 12))
label.pack(pady=25)

btn = tk.Button(
    app,
    text="Choose Excel File & Generate Report",
    command=run_automation,
    height=2,
    width=32
)
btn.pack()

app.mainloop()
