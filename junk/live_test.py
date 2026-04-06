import requests
import json

data = {
    "items": [
        {"name": "Туя западная Aureospicata", "parameters": "180-200 rb", "quantity": 1, "price": 16875, "planting": 5906.25, "total": 22781.25},
    ],
    "material_total": 16875,
    "labor_total": 5906.25,
    "tax_rate": 0,
    "tax_amount": 0,
    "grand_total": 22781.25
}

try:
    r = requests.post("http://localhost:5000/generate-xlsx", json=data, timeout=30)
    print(f"XLSX status: {r.status_code}")
    print(f"Content-Type: {r.headers.get('Content-Type')}")
    print(f"Size: {len(r.content)} bytes")
    if r.status_code != 200:
        print(f"Error body: {r.text[:2000]}")
    else:
        with open("test_download.xlsx", "wb") as f:
            f.write(r.content)
        print("Saved to test_download.xlsx - SUCCESS")
except Exception as e:
    print(f"Exception: {e}")
