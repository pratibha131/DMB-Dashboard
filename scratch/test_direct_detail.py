import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import pandas as pd
from app import create_continuous_red_detail, create_function_card, get_active_dmb_data

print("=== 1. Testing create_continuous_red_detail for Customer Service ===")
detail = create_continuous_red_detail("Customer Service", "2026-07-01")

summary = detail.children[0]
causes_card = summary.children[2]
top_causes_num = causes_card.children[0].children
print(f"Top causes displayed: {top_causes_num}")

section = detail.children[1]
cards = section.children[2].children
print(f"Total recovery cards generated: {len(cards)}")
for c_idx, card in enumerate(cards):
    title_group = card.children[0].children[0]
    badge_group = title_group.children[0]
    badge_text = badge_group.children[1].children
    title_text = title_group.children[1].children
    
    tbody = card.children[1].children.children[1]
    print(f"Card {c_idx+1}: [{badge_text}] Title: '{title_text}', Rows: {len(tbody.children)}")
    if title_text == "CP%":
        for r_idx, tr in enumerate(tbody.children):
            # Print cause, impact, action description
            cause = tr.children[0].children
            impact = tr.children[1].children
            impact_str = impact.children if hasattr(impact, 'children') else str(impact)
            print(f"   CP% Row {r_idx+1}: Cause='{cause}', Impact='{impact_str}'")

print("\n=== 2. Testing create_function_card (Red KPI vs 0 Red KPI) ===")
dmb = get_active_dmb_data()
dmb_jul = dmb[dmb["month"].eq(pd.Timestamp("2026-07-01"))]

# Test NAR (has 1 not_met)
nar_data = dmb_jul[dmb_jul["function"].eq("NAR")]
nar_card = create_function_card("NAR", dmb_jul)
print("NAR Card structure:")
# Check that there is no function-stat-grid
class_names = [child.className for child in nar_card.children if hasattr(child, 'className')]
print(f"  Child classes in NAR Card: {class_names}")
assert "function-stat-grid" not in class_names, "function-stat-grid should not be present!"
btn = nar_card.children[3].children[1]
btn_num = btn.children[0].children
btn_lbl = btn.children[1].children
btn_sub = btn.children[2].children
print(f"  NAR Button: Count={btn_num}, Label='{btn_lbl}', Subtext='{btn_sub}'")

# Test a function with 0 Red KPIs if any, or mock
zero_data = dmb_jul.copy()
zero_data.loc[zero_data["function"].eq("NAR"), "status"] = "Met"
zero_card = create_function_card("NAR", zero_data)
btn0 = zero_card.children[3].children[1]
btn0_num = btn0.children[0].children
btn0_lbl = btn0.children[1].children
btn0_sub = btn0.children[2].children
print(f"  Zero Card Button: Count={btn0_num}, Label='{btn0_lbl}', Subtext='{btn0_sub}'")
print("=== All checks passed! ===")
