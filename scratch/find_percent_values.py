import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import pandas as pd
from data_loader import get_data_store

store = get_data_store()
store.reload(force=True)

for name, df in [('MPR', store._mpr_raw), ('DMB', store._dmb_raw)]:
    entity_col = 'strategic_imperative' if name == 'MPR' else 'function'
    print(f'=== Testing {name} cleaning ===')
    for (ent, kpi), grp in df.groupby([entity_col, 'kpi_name']):
        for col in ['Actual', 'Target']:
            for idx, val in grp[col].items():
                if isinstance(val, str) and '%' in val:
                    m_str = grp.loc[idx, 'month'].strftime('%Y-%m')
                    print(f"{ent} | {kpi} | {m_str} {col} = {val!r}")
