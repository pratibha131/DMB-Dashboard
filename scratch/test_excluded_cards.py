import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import pandas as pd
from app import update_dmb_function_cards, get_active_dmb_data, create_continuous_red_detail

print("=== 1. Testing update_dmb_function_cards for July 2026 ===")
cards_jul = update_dmb_function_cards("2026-07-01", None)
function_names_jul = [card.children[0].children[0].children for card in cards_jul]
print(f"Rendered functions in July 2026 ({len(function_names_jul)}): {function_names_jul}")

assert "NAR" not in function_names_jul, "NAR must be excluded!"
assert "Europe" not in function_names_jul, "Europe must be excluded!"
assert "Growth" not in function_names_jul, "Growth must be excluded!"

print("\n=== 2. Testing update_dmb_function_cards for August 2026 ===")
cards_aug = update_dmb_function_cards("2026-08-01", None)
function_names_aug = [card.children[0].children[0].children for card in cards_aug]
print(f"Rendered functions in August 2026 ({len(function_names_aug)}): {function_names_aug}")

assert "NAR" not in function_names_aug, "NAR must be excluded!"
assert "Europe" not in function_names_aug, "Europe must be excluded!"
assert "Growth" not in function_names_aug, "Growth must be excluded!"

print("\n=== 3. Testing outer card button count for Customer Service (July 2026) ===")
cs_card = next(c for c in cards_jul if c.children[0].children[0].children == "Customer Service")
btn = cs_card.children[3].children[1]
btn_num = btn.children[0].children
btn_lbl = btn.children[1].children
btn_sub = btn.children[2].children
print(f"Customer Service Outer Card Button: Count={btn_num}, Label='{btn_lbl}', Subtext='{btn_sub}'")

print("\n=== 4. Testing Modal Summary Cards for Customer Service (July 2026) ===")
detail = create_continuous_red_detail("Customer Service", "2026-07-01")
summary = detail.children[0]
card1_num = summary.children[0].children[0].children
card1_lbl = summary.children[0].children[1].children
card2_num = summary.children[1].children[0].children
card2_lbl = summary.children[1].children[1].children
print(f"Modal Summary Box 1: {card1_num} | {card1_lbl}")
print(f"Modal Summary Box 2: {card2_num} | {card2_lbl}")

print("\n=== All assertions passed successfully! ===")
