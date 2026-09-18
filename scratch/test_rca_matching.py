import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import app
import pandas as pd
import re

def slug(s):
    return "rca-card-" + re.sub(r"[^a-zA-Z0-9_-]", "_", str(s)).lower()

m = pd.Timestamp("2026-07-01")
rca = app.get_function_rca_data()
causes_df = rca["causes"]
actions_df = rca["actions"]
review_kpis = rca["review_kpis"]
mpr_data = app.get_active_mpr_data()
strat_rca = app.get_active_strat_rca()

red_mpr = mpr_data[mpr_data["month"].eq(m) & (mpr_data["status"] == "Not Met")]["kpi_name"].unique().tolist()
red_strat = strat_rca[strat_rca["reporting_month"].eq(m) & strat_rca["is_red"]]["kpi_name"].unique().tolist() if not strat_rca.empty and "is_red" in strat_rca.columns else []

all_kpis = list(dict.fromkeys(red_mpr + red_strat))
print("Processing KPIs for July 2026:", len(all_kpis), all_kpis[:10])
for kpi in all_kpis:
    k_key = app.kpi_key(kpi)
    c_match = causes_df[causes_df["kpi_name"].apply(app.kpi_key).eq(k_key)]
    a_match = actions_df[actions_df["kpi_name"].apply(app.kpi_key).eq(k_key)]
    print(f"  KPI: '{kpi}' (slug: {slug(kpi)}) -> Causes: {len(c_match)}, Actions: {len(a_match)}")
