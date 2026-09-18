import sys
from pathlib import Path
sys.path.insert(0, str(Path('.')))
import pandas as pd
from data_loader import load_strategic_rca_actions, find_strategic_workbook

file_path = find_strategic_workbook()
print("Strategic workbook:", file_path)

strat_df = load_strategic_rca_actions()
print("Columns:", strat_df.columns.tolist())
print("\nUnique KPIs and their Causes / Actions:")
for kpi, grp in strat_df.groupby('kpi_name'):
    causes = grp['cause'].replace('', pd.NA).dropna().unique()
    actions = grp['action'].replace('', pd.NA).dropna().unique()
    print(f"KPI: {kpi:35} | Cause: {causes} | Action: {actions}")
