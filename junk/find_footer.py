import openpyxl
import os

filename = 'Шаблон.xlsx'
if os.path.exists(filename):
    try:
        wb = openpyxl.load_workbook(filename, data_only=True)
        ws = wb.active
        # Проверяем строки с 50 по 70
        for r in range(50, 75):
            row_vals = [ws.cell(r, c).value for c in range(1, 7)]
            if any(row_vals):
                print(f"Row {r}: {row_vals}")
    except Exception as e:
        print(f"Error: {e}")
else:
    print(f"File {filename} not found.")
