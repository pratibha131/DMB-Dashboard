import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
from app import get_active_dmb_data, load_function_rca_details, create_continuous_red_detail

# Let's inspect Customer Service July 2026
# In July 2026, CS has 4 Red KPIs: CP %, CS Profitability, CS Sales, First Vist Fix (Parts)
# Let's see what happens when only Red KPIs are rendered
active_dmb = get_active_dmb_data()
rca_data = load_function_rca_details()

print("Testing all functions for July and August 2026:")
for fn in active_dmb['function'].dropna().unique():
    for month_str in ['2026-07-01', '2026-08-01']:
        month = pd.Timestamp(month_str)
        red_kpis = active_dmb[(active_dmb['function'] == fn) & (active_dmb['month'] == month) & (active_dmb['status'] == 'Not Met')]['kpi_name'].tolist()
        res = create_continuous_red_detail(fn, month_str)
        print(f"Function: {fn:20} | Month: {month.strftime('%b %Y')} | Red KPIs count: {len(red_kpis)}")
