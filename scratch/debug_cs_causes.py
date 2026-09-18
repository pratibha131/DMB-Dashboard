import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import pandas as pd
from app import get_function_rca_data, function_key, get_active_dmb_data

rca_data = get_function_rca_data()
causes = rca_data["causes"]
actions = rca_data["actions"]

print("=== ALL CAUSES FOR CUSTOMER SERVICE ===")
cs_causes = causes[causes["source_sheet"].str.contains("Customer Service", case=False, na=False)]
print(cs_causes[["source_sheet", "kpi_name", "kpi_key", "cause_rank", "cause", "impact_percent"]])

print("\n=== ALL ACTIONS FOR CUSTOMER SERVICE ===")
cs_actions = actions[actions["source_sheet"].str.contains("Customer Service", case=False, na=False)]
print(cs_actions[["source_sheet", "kpi_name", "kpi_key", "root_cause", "corrective_action", "owner", "status"]])

print("\n=== DMB RED KPIS IN JULY 2026 FOR CUSTOMER SERVICE ===")
dmb = get_active_dmb_data()
cs_dmb = dmb[(dmb["function"] == "Customer Service") & (dmb["month"] == "2026-07-01")]
print(cs_dmb[["kpi_name", "status", "is_continuous_red"]])
