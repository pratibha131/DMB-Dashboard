import sys
from pathlib import Path
sys.path.insert(0, str(Path('.')))
import pandas as pd
from app import create_continuous_red_detail

detail = create_continuous_red_detail('Customer Service', '2026-07-01')
section = detail.children[1]
table_stack = section.children[2]
cp_card = table_stack.children[0]

table_wrap = cp_card.children[1]
table = table_wrap.children if not isinstance(table_wrap.children, list) else table_wrap.children[0]
tbody = table.children[1]
print(f"CP% table rows count: {len(tbody.children)}")
for idx, row in enumerate(tbody.children):
    print(f"\n--- Row {idx} ({len(row.children)} cells) ---")
    for c_idx, cell in enumerate(row.children):
        span = getattr(cell, 'rowSpan', 1) or 1
        content = getattr(cell.children, 'children', cell.children) if hasattr(cell.children, 'children') else cell.children
        print(f"  Cell {c_idx} (rowSpan={span}): {content}")
