import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
from app import get_active_dmb_data, load_function_rca_details, kpi_match_score, function_key

active_dmb = get_active_dmb_data()
rca_data = load_function_rca_details()

cs_causes = rca_data['causes'][rca_data['causes']['function_key'] == 'customer service']
cs_actions = rca_data['actions'][rca_data['actions']['function_key'] == 'customer service']

red_kpis = ['CP %', 'CS Profitability', 'CS Sales', 'First Vist Fix (Parts)']

print("Causes unique KPIs:", cs_causes['kpi_name'].unique().tolist())
print("Actions unique KPIs:", cs_actions['kpi_name'].unique().tolist())

for rk in red_kpis:
    print(f"\nMatching for Red KPI: {rk}")
    for ck in cs_causes['kpi_name'].unique():
        score = kpi_match_score(ck, rk)
        if score > 0:
            print(f"  Cause KPI '{ck}' score: {score}")
    for ak in cs_actions['kpi_name'].unique():
        score = kpi_match_score(ak, rk)
        if score > 0:
            print(f"  Action KPI '{ak}' score: {score}")
