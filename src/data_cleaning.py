import pandas as pd

CATEGORY_KEYS = ["category", "department", "type"]
AMOUNT_KEYS = ["amount", "total", "cost", "value"]
PERSON_KEYS = ["person", "employee", "name"]

def find_column(columns, keys):
    for col in columns:
        for key in keys:
            if key in col.lower():
                return col
    return None

def clean_data(file_path):
    df = pd.read_excel(file_path)

    df.columns = df.columns.str.strip().str.lower()

    category_col = find_column(df.columns, CATEGORY_KEYS)
    amount_col = find_column(df.columns, AMOUNT_KEYS)
    person_col = find_column(df.columns, PERSON_KEYS)

    if not category_col or not amount_col:
        raise Exception("Required columns not found in Excel file.")

    df["category"] = df[category_col].astype(str).str.strip()
    df["amount"] = pd.to_numeric(df[amount_col], errors="coerce")

    if person_col:
        df["person"] = df[person_col]

    df = df.dropna(subset=["amount"])
    df = df.reset_index(drop=True)

    return df
