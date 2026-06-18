import requests

key = "WSRJauAh1VrZwQofU6iBwAsgMfEMcpBhdMAdWpKqwJf"
base = "https://api.docuseal.com"
headers = {"X-Auth-Token": key, "Accept": "application/json"}

r = requests.get(f"{base}/templates/4356184", headers=headers)
t = r.json()

for f in t.get("fields", []):
    print(f"\nField: {f['name']}")
    for area in f.get("areas", []):
        print(f"  page={area['page']} x={area['x']} y={area['y']} w={area['w']} h={area['h']}")