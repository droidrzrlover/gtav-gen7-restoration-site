#!/usr/bin/env python3
"""Copy approved PNG screenshots from SharePoint into the public site repository."""
import json
import os
import pathlib
import urllib.error
import urllib.parse
import urllib.request

TENANT = os.environ["AZURE_TENANT_ID"]
CLIENT = os.environ["AZURE_CLIENT_ID"]
SECRET = os.environ["AZURE_CLIENT_SECRET"]
DRIVE = "b!R0bc46Q_MkaB_7y7Wj67eC-pWXyWfV9Mvo3eEJsAaYjeqPlaMT-zRbTjbOm63vLA"
FOLDER = "Media/showcase"
DEST = pathlib.Path("showcase")
ALLOWED = {
    "gen7-full-map.png",
    "gen7-gta-online.png",
    "gen7-hud-minimap.png",
    "gen7-pause-menu.png",
    "gen7-weapon-wheel.png",
}
MAX_SIZE = 15 * 1024 * 1024

def request(url, *, data=None, headers=None):
    req = urllib.request.Request(url, data=data, headers=headers or {}, method="POST" if data is not None else "GET")
    with urllib.request.urlopen(req, timeout=90) as response:
        return response.read()

payload = urllib.parse.urlencode({
    "client_id": CLIENT,
    "client_secret": SECRET,
    "scope": "https://graph.microsoft.com/.default",
    "grant_type": "client_credentials",
}).encode()
auth = json.loads(request(
    f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token",
    data=payload,
    headers={"Content-Type": "application/x-www-form-urlencoded"},
))
token = auth["access_token"]
headers = {"Authorization": f"Bearer {token}"}
folder_path = urllib.parse.quote(FOLDER, safe="/")
drive = urllib.parse.quote(DRIVE, safe="")
url = f"https://graph.microsoft.com/v1.0/drives/{drive}/root:/{folder_path}:/children?$select=id,name,size,file&$top=200"
items = []
while url:
    page = json.loads(request(url, headers=headers))
    items.extend(page.get("value", []))
    url = page.get("@odata.nextLink")

DEST.mkdir(parents=True, exist_ok=True)
synced = []
for item in items:
    name = item.get("name", "")
    if name not in ALLOWED or not item.get("file"):
        continue
    if int(item.get("size", MAX_SIZE + 1)) > MAX_SIZE:
        raise ValueError(f"Screenshot exceeds {MAX_SIZE} bytes: {name}")
    item_id = urllib.parse.quote(item["id"], safe="")
    # /content redirects to a short-lived download URL; urllib follows it.
    data = request(f"https://graph.microsoft.com/v1.0/drives/{drive}/items/{item_id}/content", headers=headers)
    if len(data) > MAX_SIZE or not data.startswith(bytes.fromhex("89504e470d0a1a0a")):
        raise ValueError(f"Invalid PNG or oversized file: {name}")
    (DEST / name).write_bytes(data)
    synced.append(name)

missing = ALLOWED.difference(synced)
if missing:
    raise RuntimeError("Expected screenshots missing from SharePoint: " + ", ".join(sorted(missing)))
print("Synchronized " + str(len(synced)) + " PNG screenshots: " + ", ".join(sorted(synced)))
