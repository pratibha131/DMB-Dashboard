import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import pandas as pd
from app import get_function_rca_data, create_continuous_red_detail

rca_data = get_function_rca_data()
actions = rca_data["actions"]
print("Actions columns:", actions.columns.tolist())
print("Actions function_keys:", actions["function_key"].unique())
print("Actions source_sheets:", actions["source_sheet"].unique())
cs_actions = actions[actions["source_sheet"] == "7.Customer Service"]
print("Customer Service actions:")
print(cs_actions[["kpi_name", "root_cause", "corrective_action", "owner", "status"]])

detail = create_continuous_red_detail("Customer Service", "2026-07-01")
print("\nRendered CS Detail for July 2026 successfully!")
