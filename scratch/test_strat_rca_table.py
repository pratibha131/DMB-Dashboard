import sys
from pathlib import Path
sys.path.insert(0, str(Path('.')))
import pandas as pd
from data_loader import load_strategic_rca_actions
from app import create_rca_table, get_active_rca_actions, get_active_mpr_data

strat_actions = load_strategic_rca_actions()
print("Total strat_actions rows:", len(strat_actions))
print("Rows with cause:", len(strat_actions[strat_actions['cause'].ne('')]))
print("Rows with action:", len(strat_actions[strat_actions['action'].ne('')]))

# Check May 2026
table_may = create_rca_table('2026-05-01')
print("May 2026 table type:", type(table_may))
