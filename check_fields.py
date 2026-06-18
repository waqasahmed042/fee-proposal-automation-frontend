import requests

key = "WSRJauAh1VrZwQofU6iBwAsgMfEMcpBhdMAdWpKqwJf"
base = "https://api.docuseal.com"
headers = {"X-Auth-Token": key, "Accept": "application/json"}

r = requests.get(f"{base}/templates", headers=headers)
data = r.json()
templates = data.get("data", data) if isinstance(data, dict) else data

if not templates:
    print("No templates found")
else:
    t = templates[0]
    print("ID:", t["id"])
    print("Name:", t["name"])
    print("Field count:", len(t.get("fields", [])))
    print("Fields:", [f["name"] for f in t.get("fields", [])])