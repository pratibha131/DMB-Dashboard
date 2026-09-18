import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
from app import get_active_dmb_data, load_function_rca_details, kpi_match_score, function_key

active_dmb = get_active_dmb_data()
rca_data = load_function_rca_details()

for fn in sorted(active_dmb['function'].dropna().unique()):
    for month_str in ['2026-07-01', '2026-08-01']:
        month = pd.Timestamp(month_str)
        red_kpis = active_dmb[(active_dmb['function'] == fn) & (active_dmb['month'] == month) & (active_dmb['status'] == 'Not Met')]['kpi_name'].tolist()
        if not red_kpis:
            continue
        print(f"\n=======================================================")
        print(f"Function: {fn} | Month: {month.strftime('%b %Y')} | Red KPIs: {red_kpis}")
        
        fn_key = function_key(fn)
        fn_causes = rca_data['causes'][rca_data['causes']['function_key'] == fn_key]
        fn_actions = rca_data['actions'][rca_data['actions']['function_key'] == fn_key]
        
        for rk in red_kpis:
            rk_causes = [ck for ck in fn_causes['kpi_name'].unique() if kpi_match_score(ck, rk) > 0]
            rk_actions = [ak for ak in fn_actions['kpi_name'].unique() if kpi_match_score(ak, rk) > 0]
            print(f"  Red KPI: '{rk}' -> Causes: {rk_causes}, Actions: {rk_actions}")
