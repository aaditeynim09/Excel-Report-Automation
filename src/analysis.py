def generate_summary(df):
    category_summary = df.groupby("category")["amount"].sum().reset_index()

    top_expenses = df.sort_values("amount", ascending=False).head(5)

    return category_summary, top_expenses
