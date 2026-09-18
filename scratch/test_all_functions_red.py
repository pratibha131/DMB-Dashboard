import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
from app import get_active_dmb_data, load_function_rca_details, function_key

active_dmb = get_active_dmb_data()
rca_data = load_function_rca_details()

print("Functions in DMB:", active_dmb['function'].dropna().unique())
print("Functions in RCA causes:", rca_data['causes']['function'].dropna().unique())
print("Functions in RCA actions:", rca_data['actions']['function'].dropna().unique())

for fn in active_dmb['function'].dropna().unique():
    for month in [pd.Timestamp('2026-07-01'), pd.Timestamp('2026-08-01')]:
        red_kpis = active_dmb[(active_dmb['function'] == fn) & (active_dmb['month'] == month) & (active_dmb['status'] == 'Not Met')]['kpi_name'].tolist()
        print(f"[{fn}] {month.strftime('%b %Y')} Red KPIs ({len(red_kpis)}): {red_kpis}")
