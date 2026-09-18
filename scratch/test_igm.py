import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import app

rca = app.get_function_rca_data()
causes_df = rca["causes"]
actions_df = rca["actions"]

igm_causes = causes_df[causes_df["kpi_name"].str.contains("IGM", case=False, na=False)]
igm_actions = actions_df[actions_df["kpi_name"].str.contains("IGM", case=False, na=False)]

print("IGM Causes:\n", igm_causes[["kpi_name", "cause_rank", "cause", "impact_percent"]])
print("\nIGM Actions:\n", igm_actions[["kpi_name", "corrective_action", "root_cause", "owner", "status"]])
