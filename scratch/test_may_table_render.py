import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import pandas as pd
from data_loader import get_data_store, load_strategic_rca_actions
from app import create_rca_table, get_active_rca_actions

ds = get_data_store()
ds.reload()

table_may = create_rca_table('2026-05-01')
table_obj = table_may.children
tbody = table_obj.children[1]

print(f"May 2026 Table Rows: {len(tbody.children)}")
for idx, row in enumerate(tbody.children):
    cells = [c.children.children if hasattr(c.children, 'children') else c.children for c in row.children]
    print(f"\nRow {idx}:")
    print(f"  Imperative: {cells[0]}")
    print(f"  KPI Name:   {cells[1]}")
    print(f"  Cause:      {str(cells[2])[:60]}")
    print(f"  Action:     {str(cells[3])[:60]}")
