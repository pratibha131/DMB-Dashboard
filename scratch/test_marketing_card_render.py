import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import pandas as pd
from data_loader import get_data_store
from app import create_function_card

store = get_data_store()
store.reload(force=True)
mpr, dmb, rca = store.get_data()

current = dmb[dmb["month"] == pd.Timestamp("2026-09-01")].copy()
card = create_function_card("Marketing", current)
print("Marketing card rendered successfully.")
# Let's inspect bar segments inside card
# card is html.Div with children: [header, subtitle, stacked_bar, bottom]
stacked_bar = card.children[2]
print("Stacked bar segments:")
for seg in stacked_bar.children:
    print("  Segment:", seg.children, "class:", getattr(seg, 'className', ''), "title:", getattr(seg, 'title', ''))
