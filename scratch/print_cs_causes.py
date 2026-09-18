import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import pandas as pd
from app import get_function_rca_data

rca_data = get_function_rca_data()
causes = rca_data["causes"]
cs_causes = causes[causes["source_sheet"].str.contains("Customer Service", case=False, na=False)]
print(cs_causes[["kpi_name", "kpi_key", "cause_rank", "cause", "impact_percent"]])
