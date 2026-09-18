import sys
from pathlib import Path
sys.path.insert(0, str(Path('.')))
import pandas as pd
from dash import html

def calculate_row_spans(total_rows, count):
    if count <= 0:
        return []
    base = total_rows // count
    extra = total_rows % count
    spans = []
    curr = 0
    for i in range(count):
        span = base + (1 if i < extra else 0)
        spans.append((curr, span))
        curr += span
    return spans

print("4 rows, 2 items:", calculate_row_spans(4, 2))
print("4 rows, 3 items:", calculate_row_spans(4, 3))
print("5 rows, 2 items:", calculate_row_spans(5, 2))
print("2 rows, 4 items:", calculate_row_spans(max(2, 4), 2))
