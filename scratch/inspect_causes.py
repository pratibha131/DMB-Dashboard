import pandas as pd
from pathlib import Path

def inspect_all():
    files = list(Path('data').glob('*Strategic*.xlsx'))
    fpath = files[0]
    xl = pd.ExcelFile(fpath)
    for s in xl.sheet_names:
        df = xl.parse(s, header=None)
        print(f"=== Sheet: {s} ({df.shape}) ===")
        for r in range(len(df)):
            cause = df.iloc[r, 23] if df.shape[1] > 23 else None
            act = df.iloc[r, 24] if df.shape[1] > 24 else None
            if pd.notna(cause) or pd.notna(act):
                print(f"  Row {r}: kpi={df.iloc[r, 1]} | cause={str(cause)[:40]} | act={str(act)[:40]}")

if __name__ == '__main__':
    inspect_all()
