import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
from app import get_active_dmb_data, load_function_rca_details, function_key

rca_data = load_function_rca_details()
causes = rca_data['causes']
actions = rca_data['actions']

cs_causes = causes[causes['function_key'].eq('customer service')]
cs_actions = actions[actions['function_key'].eq('customer service')]

print("Customer Service Causes KPIs in Excel:")
print(cs_causes[['kpi_name', 'cause', 'impact_percent']])

print("\nCustomer Service Actions KPIs in Excel:")
print(cs_actions[['kpi_name', 'root_cause', 'corrective_action', 'owner', 'status']])
