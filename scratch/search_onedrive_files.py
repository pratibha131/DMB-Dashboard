import os
from pathlib import Path

# search in OneDrive
onedrive_dir = Path(r"c:\Users\320320898\OneDrive - Philips")
print("Searching in:", onedrive_dir)

try:
    for p in onedrive_dir.rglob("*.xlsx"):
        if "Functional" in p.name or "Review" in p.name or "DMB" in p.name:
            print("Found:", p, "size:", p.stat().st_size)
except Exception as e:
    print("Search error:", e)
