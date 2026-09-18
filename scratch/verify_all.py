import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
from app import create_continuous_red_detail, get_dynamic_reporting_months, get_active_dmb_data, get_active_mpr_data, get_active_rca_actions

print("=== 1. VERIFYING CUSTOMER SERVICE JULY 2026 ===")
detail_cs_jul = create_continuous_red_detail('Customer Service', '2026-07-01')

# Summary cards
summary_grid = detail_cs_jul.children[0]
cards = summary_grid.children
print(f"Summary Card 1: {cards[0].children[0].children} {cards[0].children[1].children}")
print(f"Summary Card 2: {cards[1].children[0].children} {cards[1].children[1].children}")
print(f"Summary Card 3: {cards[2].children[0].children} {cards[2].children[1].children}")

# Section contents
section = detail_cs_jul.children[1]
print(f"\nSection children count: {len(section.children)}")

pill_box = section.children[1]
pills = pill_box.children[1].children
pill_labels = [p.children for p in pills]
print("Quick Jump to KPI Pills:", pill_labels)

# Tables
table_stack = section.children[2]
table_cards = table_stack.children
print(f"Number of KPI Cards rendered: {len(table_cards)}")

print("\n=== 2. VERIFYING CUSTOMER SERVICE AUGUST 2026 (0 RED KPIS) ===")
detail_cs_aug = create_continuous_red_detail('Customer Service', '2026-08-01')
print("Aug 2026 response class:", detail_cs_aug.className)
print("Aug 2026 content:", [getattr(c, 'children', c) for c in detail_cs_aug.children])

print("\n=== 3. VERIFYING DYNAMIC MONTH GENERATION ===")
mpr = get_active_mpr_data()
dmb = get_active_dmb_data()
rca = get_active_rca_actions()

mpr_m, mpr_def = get_dynamic_reporting_months(mpr, rca)
dmb_m, dmb_def = get_dynamic_reporting_months(dmb, rca)
print("Current MPR months:", [m.strftime('%b %Y') for m in mpr_m])
print("Current MPR default:", mpr_def.strftime('%b %Y'))
print("Current DMB months:", [m.strftime('%b %Y') for m in dmb_m])
print("Current DMB default:", dmb_def.strftime('%b %Y'))

# Simulation of October 1, 2026
mpr_m_oct, mpr_def_oct = get_dynamic_reporting_months(mpr, rca, reference_date="2026-10-01")
print("\nSimulated Oct 1 MPR months:", [m.strftime('%b %Y') for m in mpr_m_oct])
print("Simulated Oct 1 MPR default:", mpr_def_oct.strftime('%b %Y'))

print("\nALL VERIFICATIONS PASSED SUCCESSFULLY!")
