"""
DMB-Dashboard Poller Integration
================================
Provides background polling for mailbox (dmbdashboard@gmail.com),
SharePoint web URLs, and OneDrive synced workbooks with auto-commit to GitHub.
"""

import sharepoint_sync
import data_loader

def start_dmb_mailbox_poller(reload_callback=None):
    """
    Starts a background daemon thread that polls mailbox / SharePoint every 60 seconds.
    Whenever a new Excel file arrives:
      1. Downloads into data/ (Masterfile / Functional Review / Strategic Execution)
      2. Commits & pushes to GitHub repository
      3. Calls reload_callback to hot-reload in-memory data store.
    """
    callback = reload_callback or (lambda: data_loader.reload_all_data(force=True))
    sharepoint_sync.start_live_sync(callback=callback)
