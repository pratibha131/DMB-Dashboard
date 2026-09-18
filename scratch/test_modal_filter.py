import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
from app import get_active_dmb_data, load_function_rca_details, create_continuous_red_detail

detail = create_continuous_red_detail('Customer Service', '2026-07-01')
# Let's inspect the children of detail
print("Top level type:", type(detail))
print("Children count:", len(detail.children))
# Print sections
section = detail.children[1]
print("Section children count:", len(section.children))
for i, c in enumerate(section.children):
    print(f"Child {i}:", type(c), getattr(c, 'className', ''))
