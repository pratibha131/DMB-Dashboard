import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import pandas as pd
from app import create_continuous_red_detail, get_active_dmb_data

dmb_df = get_active_dmb_data()
print("Columns:", dmb_df.columns.tolist())
month_col = "reporting_month" if "reporting_month" in dmb_df.columns else "month"
latest_month = dmb_df[month_col].max()
print("Latest month:", latest_month)

content = create_continuous_red_detail("Customer Service", latest_month)
print("Modal content created successfully for Customer Service!")
print(f"Content type: {type(content)}")

# Test all functions
functions = dmb_df["function"].dropna().unique()
for fn in functions:
    try:
        modal_c = create_continuous_red_detail(fn, latest_month)
        print(f"Function '{fn}': Modal content rendered successfully.")
    except Exception as e:
        print(f"ERROR for '{fn}': {e}")
