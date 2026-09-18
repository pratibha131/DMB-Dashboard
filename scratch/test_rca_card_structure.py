import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import pandas as pd
from app import create_continuous_red_detail, get_active_dmb_data, create_kpi_rca_card
from data_loader import load_rca_data

data = load_rca_data()
causes_df = pd.DataFrame(data) if isinstance(data, list) else pd.DataFrame()
print(f"Total RCA records: {len(causes_df)}")
if "kpi_name" in actions_df.columns:
    print("Action kpis:", actions_df["kpi_name"].unique())

card = create_kpi_rca_card("CP %", pd.DataFrame(), actions_df[actions_df["kpi_name"] == "CP"] if "kpi_name" in actions_df.columns else pd.DataFrame(), function_badge="Customer Service")

# Drill into table
table = card.children[2].children
thead = table.children[0]
tbody = table.children[1]

print("\nThead row 1 (Group headers):")
for th in thead.children[0].children:
    print(f"  - text: '{th.children}', colSpan: {th.colSpan}, className: '{th.className}'")

print("\nThead row 2 (Subhead columns):")
for th in thead.children[1].children:
    print(f"  - text: '{th.children}', className: '{th.className}'")

print(f"\nTbody rows count: {len(tbody.children)}")
for idx, tr in enumerate(tbody.children):
    print(f"Row {idx + 1}: {len(tr.children)} cells")
    for c_idx, td in enumerate(tr.children):
        # Extract text
        txt = td.children
        if hasattr(txt, 'children'):
            txt = txt.children
        print(f"   Cell {c_idx + 1}: colSpan={getattr(td, 'colSpan', None)}, rowSpan={getattr(td, 'rowSpan', None)}, class='{getattr(td, 'className', '')}', content='{txt}'")

# Test CP %
cp_causes = causes_df[causes_df["kpi"].str.strip().str.lower() == "cp %"] if not causes_df.empty else pd.DataFrame()
cp_actions = actions_df[actions_df["kpi"].str.strip().str.lower() == "cp %"] if not actions_df.empty else pd.DataFrame()

print(f"\nCP % -> Causes: {len(cp_causes)}, Actions: {len(cp_actions)}")
if not cp_actions.empty:
    print(cp_actions[["kpi", "root_cause", "corrective_action", "owner", "status"]])

card = create_kpi_rca_card("CP %", cp_causes, cp_actions, function_badge="Customer Service")
print("\nCard created successfully for CP %!")

# Test all red KPIs from Customer Service
cs_actions = actions_df[actions_df["function"] == "Customer Service"] if "function" in actions_df.columns else pd.DataFrame()
print(f"\nCustomer service actions count: {len(cs_actions)}")
if not cs_actions.empty:
    unique_kpis = cs_actions["kpi"].unique()
    print("Customer service KPIs in actions:", unique_kpis)
    for kpi in unique_kpis:
        c_kpi = causes_df[(causes_df["kpi"] == kpi) & (causes_df["function"] == "Customer Service")] if not causes_df.empty else pd.DataFrame()
        a_kpi = cs_actions[cs_actions["kpi"] == kpi]
        kpi_card = create_kpi_rca_card(kpi, c_kpi, a_kpi, function_badge="Customer Service")
        print(f"Card for '{kpi}': {len(kpi_card.children)} main children components")

print("\nALL RCA CARD TESTS PASSED!")
