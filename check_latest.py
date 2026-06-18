import requests

key = "WSRJauAh1VrZwQofU6iBwAsgMfEMcpBhdMAdWpKqwJf"
base = "https://api.docuseal.com"
headers = {"X-Auth-Token": key, "Accept": "application/json"}

r = requests.get(f"{base}/templates", headers=headers)
data = r.json()
templates = data.get("data", data) if isinstance(data, dict) else data

# Sort by ID descending to get newest first
templates.sort(key=lambda x: x["id"], reverse=True)

print(f"Total templates: {len(templates)}")
for t in templates[:5]:  # show 5 most recent
    fields = [f["name"] for f in t.get("fields", [])]
    print(f"\nID: {t['id']} | Name: {t['name']}")
    print(f"Fields ({len(fields)}): {fields}")