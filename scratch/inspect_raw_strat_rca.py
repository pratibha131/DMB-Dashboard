import sys
from pathlib import Path
sys.path.insert(0, str(Path('.')))
import pandas as pd
from data_loader import load_strategic_rca_actions, find_strategic_workbook

# Let's inspect what causes and actions exist for all 22 KPIs in AOP Critical
excel_path = find_strategic_workbook()
df = pd.read_excel(excel_path, sheet_name="AOP Critical", header=None)

for r in range(len(df)):
    strat = df.iloc[r, 0] if pd.notna(df.iloc[r, 0]) else ""
    kpi = df.iloc[r, 1] if pd.notna(df.iloc[r, 1]) else ""
    cause = df.iloc[r, 23] if df.shape[1] > 23 and pd.notna(df.iloc[r, 23]) else ""
    action = df.iloc[r, 24] if df.shape[1] > 24 and pd.notna(df.iloc[r, 24]) else ""
    if cause or action:
        print(f"Row {r:2d} | KPI: {str(kpi):30} | Cause: {str(cause)[:40]} | Action: {str(action)[:40]}")
