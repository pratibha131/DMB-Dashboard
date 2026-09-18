import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
import data_loader

strat_df = data_loader.load_strategic_rca_actions()
print(f"Total strategic records: {len(strat_df)}")
print("Columns:", strat_df.columns.tolist())

# Group by month and see red KPIs
for month, grp in strat_df.groupby("reporting_month"):
    print(f"\nMonth: {month}")
    print(f"Total KPIs in file: {len(grp)}")
    reds = grp[grp["is_red"]]
    print(f"Red KPIs ({len(reds)}): {reds['kpi_name'].tolist()}")
    greens = grp[grp["actual"].notna() & (~grp["is_red"])]
    print(f"Green KPIs ({len(greens)}): {greens['kpi_name'].tolist()}")
