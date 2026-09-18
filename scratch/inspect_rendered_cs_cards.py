import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import pandas as pd
from app import create_continuous_red_detail

detail = create_continuous_red_detail("Customer Service", "2026-07-01")

# Find the section and cards (index 0 is header, index 1 is pill_box, index 2 is cause-table-stack)
section_children = detail.children[1].children
stack = section_children[-1].children
print(f"Total cards rendered: {len(stack)}")

for card in stack:
    # Title
    title_box = card.children[0].children[0]
    title_text = title_box.children
    print(f"\n======================================")
    print(f"Card Title: {title_text}")
    table = card.children[2].children
    tbody = table.children[1]
    print(f"Rows count: {len(tbody.children)}")
    for r_idx, tr in enumerate(tbody.children):
        cell_contents = []
        for td in tr.children:
            txt = td.children
            if hasattr(txt, 'children'):
                txt = txt.children
            cell_contents.append(str(txt))
        print(f"  Row {r_idx + 1}: {' | '.join(cell_contents)}")
