import pandas as pd
from data_cleaning import clean_data
from analysis import generate_summary
from report_generator import format_excel

input_file = "input_files/raw_data.xlsx"
output_file = "output_files/final_report.xlsx"

df = clean_data(input_file)

summary, top = generate_summary(df)

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Clean Data", index=False)
    summary.to_excel(writer, sheet_name="Summary", index=False)
    top.to_excel(writer, sheet_name="Top 5 Transactions", index=False)

format_excel(output_file)

print("Report generated successfully.")
