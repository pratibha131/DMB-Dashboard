import sys
from pathlib import Path
sys.path.insert(0, str(Path('.')))
import pandas as pd
from data_loader import get_data_store, load_strategic_rca_actions

ds = get_data_store()
ds.reload()
rca_actions = ds.get_strat_rca()

print("Checking May 2026 Red KPIs in Strategic RCA:")
may_red = rca_actions[rca_actions['reporting_month'].eq(pd.Timestamp('2026-05-01')) & rca_actions['is_red']]
print(may_red[['strategic_imperative', 'kpi_name', 'cause', 'action']])
