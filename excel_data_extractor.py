from openpyxl import load_workbook
import os

excel_files = os.listdir("data")
wb = load_workbook(filename=f"data/{excel_files[0]}")

print(wb.active)