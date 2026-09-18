import openpyxl
from pathlib import Path

files_to_update = [
    Path("data/Functional DMB Review Sheets-16th Aug.xlsx"),
    Path(r"c:\Users\320320898\OneDrive - Philips\Functional DMB Review Sheets.xlsx"),
]

for p in files_to_update:
    if not p.exists():
        print(f"Skipping non-existent: {p}")
        continue
    try:
        wb = openpyxl.load_workbook(p)
        if "7.Customer Service" in wb.sheetnames:
            ws = wb["7.Customer Service"]
            
            # Check if CP% is already in rows 48-80
            has_cp = False
            for r in range(48, min(90, ws.max_row + 1)):
                if str(ws.cell(r, 1).value or "").strip().lower() in {"cp%", "cp %", "cp"}:
                    has_cp = True
                    break
            
            if not has_cp:
                # Find row with 'Paretos for RCA' or 'Action tracker'
                target_r = 71
                for r in range(48, min(100, ws.max_row + 1)):
                    v = str(ws.cell(r, 1).value or "").strip().lower()
                    if "pareto" in v or "action tracker" in v:
                        target_r = r
                        break
                
                ws.insert_rows(target_r, 3)
                
                ws.cell(target_r, 1, "Red KPI")
                ws.cell(target_r, 2, "Cause 1 ")
                ws.cell(target_r, 3, "Cause 2")
                ws.cell(target_r, 4, "Cause 3")
                ws.cell(target_r, 5, "Cause 4")
                ws.cell(target_r, 6, "Cause 5")

                ws.cell(target_r + 1, 1, "CP%")
                ws.cell(target_r + 1, 2, "META Contract Renewal GAP")
                ws.cell(target_r + 1, 3, "GRC CP drop")
                ws.cell(target_r + 1, 4, "EOL / EOS Impact")
                ws.cell(target_r + 1, 5, "Pending delayed renewal (Global)")

                ws.cell(target_r + 2, 1, "% impact on KPI gap")
                ws.cell(target_r + 2, 2, 0.40)
                ws.cell(target_r + 2, 3, 0.30)
                ws.cell(target_r + 2, 4, 0.20)
                ws.cell(target_r + 2, 5, 0.10)
                
                wb.save(p)
                print(f"Successfully added CP% to {p}")
            else:
                print(f"CP% already exists in {p}")
        wb.close()
    except Exception as e:
        print(f"Error updating {p}: {e}")
