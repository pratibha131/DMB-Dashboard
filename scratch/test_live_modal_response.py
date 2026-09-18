import requests
import json

# Send request to Dash server to trigger modal for Customer Service
url = "http://127.0.0.1:10000/_dash-update-component"

payload = {
    "output": "..continuous-red-modal.className...continuous-red-modal-title.children...continuous-red-modal-month.children...continuous-red-modal-body.children..",
    "inputs": [
        {"id": {"type": "continuous-red-card", "function": "Customer Service"}, "property": "n_clicks", "value": 1}
    ],
    "changedPropIds": ['{"function":"Customer Service","type":"continuous-red-card"}.n_clicks'],
    "state": [
        {"id": "dmb-month-filter", "property": "value", "value": "2026-07-01"}
    ]
}

res = requests.post(url, json=payload, headers={"Content-Type": "application/json"})
print("Response status:", res.status_code)
text = res.text
print("Length of response:", len(text))
print("Contains 'META Contract Renewal GAP':", "META Contract Renewal GAP" in text)
print("Contains 'GRC CP drop':", "GRC CP drop" in text)
print("Contains '28':", "28" in text)
print("Contains '40%':", "40%" in text)
