import openpyxl

wb = openpyxl.load_workbook("data/Functional DMB Review Sheets-16th Aug.xlsx")
ws = wb["7.Customer Service"]

print("Current rows 48-75:")
for r in range(48, 75):
    row_vals = [ws.cell(r, c).value for c in range(1, 10)]
    if any(v is not None for v in row_vals):
        print(f"Row {r}: {row_vals[:7]}")
