import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
from app import create_continuous_red_detail, get_active_dmb_data

active_dmb = get_active_dmb_data()
months = ['2026-01-01', '2026-02-01', '2026-03-01', '2026-04-01', '2026-05-01', '2026-06-01', '2026-07-01', '2026-08-01']

for fn in sorted(active_dmb['function'].dropna().unique()):
    for m in months:
        dt = pd.Timestamp(m)
        red_kpis = active_dmb[(active_dmb['function'] == fn) & (active_dmb['month'] == dt) & (active_dmb['status'] == 'Not Met')]['kpi_name'].tolist()
        res = create_continuous_red_detail(fn, m)
        is_empty = getattr(res, 'className', '') == 'modal-empty-state'
        print(f"[{fn:20}] {dt.strftime('%b %Y')} -> {len(red_kpis)} Red KPIs | EmptyState: {is_empty}")

print("\nALL FUNCTIONS AND MONTHS TESTED WITH ZERO ERRORS!")
