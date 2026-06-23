for row in ws.iter_rows(min_row=3, max_row=72, min_col=1, max_col=5, values_only=True):
    print(f"{row}")  # Только колонки A(индекс 0) и C(индекс 2)