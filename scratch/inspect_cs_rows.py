import openpyxl

wb = openpyxl.load_workbook("data/Functional DMB Review Sheets-16th Aug.xlsx")
ws = wb["7.Customer Service"]

print("Current rows 48 to 72:")
for r in range(48, 72):
    row_vals = [ws.cell(r, c).value for c in range(1, 8)]
    print(f"Row {r:2d}: {row_vals}")
