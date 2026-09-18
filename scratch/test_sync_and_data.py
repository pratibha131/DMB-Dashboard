import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import email_sync
import sharepoint_sync
import data_loader
import app

print("--- 1. Testing Workbook Standardization ---")
test_names = [
    "Masterfile_DMB_Dashboard.xlsx",
    "MPR Masterfile 2026.xlsx",
    "Functional DMB Review Sheets-.xlsx",
    "Functional Review.xlsx",
    "Strategic Execution Dashboard-11th_Sept.xlsx",
    "Strategic Execution Dashboard.xlsx",
    "AOP Critical Red KPIs.xlsx",
]
for name in test_names:
    std = email_sync.standardize_workbook_filename(name)
    print(f"  '{name}' -> '{std}'")

print("\n--- 2. Testing Data Store Reload ---")
store = data_loader.get_data_store()
reloaded = store.reload(force=True)
mpr, dmb, strat = store.get_data()
print(f"  Data Store Loaded: MPR={mpr.shape}, DMB={dmb.shape}, Strategic RCA={strat.shape}")

print("\n--- 3. Testing Sync Manager Status ---")
mgr = sharepoint_sync.get_sync_manager()
status = mgr.get_status()
print(f"  Sync Status: {status['status']}, Poll Interval={status['poll_interval']}s, Mailbox={status['mailbox_user']}")

print("\n--- 4. Testing Live Sync Callback (1-Hour Refresh) ---")
res = app.handle_live_sync_trigger(0)
print(f"  Live Sync Callback: status={res.get('status')}, updated={res.get('updated')}")

print("\n--- 5. Testing GitHub Repository Detection ---")
repo = email_sync.get_configured_github_repo()
print(f"  Target GitHub Repo: {repo}")

print("\n>>> ALL TESTS COMPLETED SUCCESSFULLY! <<<")
