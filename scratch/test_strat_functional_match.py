import sys
from pathlib import Path
sys.path.insert(0, str(Path('.')))
import pandas as pd
from app import load_function_rca_details, kpi_match_score

rca = load_function_rca_details()
causes = rca['causes']
actions = rca['actions']

for kpi in ['NPI Schedule Adherence', '% of CAPA investgation', 'Inventory MAT sales', 'Contract Penetration']:
    print(f"\nChecking functional matches for: {kpi}")
    for c_kpi in causes['kpi_name'].unique():
        sc = kpi_match_score(c_kpi, kpi)
        if sc > 0:
            print(f"  Matched cause KPI: '{c_kpi}' (score {sc})")
    for a_kpi in actions['kpi_name'].unique():
        sc = kpi_match_score(a_kpi, kpi)
        if sc > 0:
            print(f"  Matched action KPI: '{a_kpi}' (score {sc})")
