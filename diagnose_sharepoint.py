"""Step-by-step SharePoint/Graph diagnostic. Run this directly:

    python diagnose_sharepoint.py

It signs you in the same way the app does, then walks through each Graph
call one at a time, printing the URL and result so we can see exactly
which step 404s.
"""

import sys
import urllib.parse

import requests
from msal import PublicClientApplication

import config

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPES = ["Sites.Read.All", "Files.Read.All"]


def step(label, url):
    print(f"\n--- {label} ---")
    print(f"GET {url}")
    resp = requests.get(url, headers=headers, timeout=30)
    print(f"Status: {resp.status_code}")
    if not resp.ok:
        print(f"Body: {resp.text[:2000]}")
        sys.exit(1)
    data = resp.json()
    return data


print(f"CLIENT_ID = {config.CLIENT_ID}")
print(f"TENANT_ID = {config.TENANT_ID}")
print(f"SHAREPOINT_HOSTNAME = {config.SHAREPOINT_HOSTNAME}")
print(f"SHAREPOINT_SITE_PATH = {config.SHAREPOINT_SITE_PATH}")
print(f"SHAREPOINT_DEFAULT_PATH = {config.SHAREPOINT_DEFAULT_PATH}")

app = PublicClientApplication(
    config.CLIENT_ID,
    authority=f"https://login.microsoftonline.com/{config.TENANT_ID}",
)
result = app.acquire_token_interactive(scopes=SCOPES)
if not result or "access_token" not in result:
    print("Sign-in failed:", (result or {}).get("error_description"))
    sys.exit(1)
headers = {"Authorization": f"Bearer {result['access_token']}"}
print("\nSigned in OK.")
account = (result.get("id_token_claims") or {})
print(f"Signed in as: {account.get('preferred_username') or account.get('upn') or account.get('email')}")
print(f"  name: {account.get('name')}")
print(f"  oid (object id): {account.get('oid')}")
print(f"  tid (tenant id): {account.get('tid')}")

site_url = f"{GRAPH_BASE}/sites/{config.SHAREPOINT_HOSTNAME}:{config.SHAREPOINT_SITE_PATH}"
site = step("Resolve site", site_url)
site_id = site["id"]
print(f"site id = {site_id}")
print(f"site displayName = {site.get('displayName')}")
print(f"site webUrl = {site.get('webUrl')}")

drives_url = f"{GRAPH_BASE}/sites/{site_id}/drives"
drives = step("List ALL drives (document libraries) on this site", drives_url)
print("\nDrives on this site:")
for d in drives.get("value", []):
    print(f"  - name={d.get('name')!r} id={d['id']} webUrl={d.get('webUrl')}")

default_drive = step("Resolve DEFAULT drive (what the app actually uses)", f"{GRAPH_BASE}/sites/{site_id}/drive")
drive_id = default_drive["id"]
print(f"\nDefault drive name = {default_drive.get('name')!r}")
print(f"Default drive webUrl = {default_drive.get('webUrl')}")

root = step("List default drive ROOT children", f"{GRAPH_BASE}/drives/{drive_id}/root/children")
print("\nRoot-level items in default drive:")
for item in root.get("value", []):
    kind = "folder" if "folder" in item else "file"
    print(f"  - [{kind}] {item['name']}")

encoded = urllib.parse.quote(config.SHAREPOINT_DEFAULT_PATH)
target_url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{encoded}:/children"
print(f"\n--- List target path children ({config.SHAREPOINT_DEFAULT_PATH!r}) ---")
print(f"GET {target_url}")
resp = requests.get(target_url, headers=headers, timeout=30)
print(f"Status: {resp.status_code}")
if not resp.ok:
    print(f"Body: {resp.text[:2000]}")
else:
    target = resp.json()
    print("\nItems in target folder:")
    for item in target.get("value", []):
        kind = "folder" if "folder" in item else "file"
        print(f"  - [{kind}] {item['name']}")

shared = step("List items shared directly with me", f"{GRAPH_BASE}/me/drive/sharedWithMe")
print("\nItems shared directly with you:")
for item in shared.get("value", []):
    remote = item.get("remoteItem", {})
    parent = remote.get("parentReference", {})
    kind = "folder" if "folder" in remote else "file"
    print(f"  - [{kind}] name={item.get('name')!r} driveId={parent.get('driveId')} itemId={remote.get('id')}")
    if "folder" in remote and parent.get("driveId") and remote.get("id"):
        children_url = f"{GRAPH_BASE}/drives/{parent['driveId']}/items/{remote['id']}/children"
        print(f"    GET {children_url}")
        r2 = requests.get(children_url, headers=headers, timeout=30)
        print(f"    Status: {r2.status_code}")
        if r2.ok:
            for child in r2.json().get("value", []):
                ckind = "folder" if "folder" in child else "file"
                print(f"      - [{ckind}] {child['name']}")
        else:
            print(f"    Body: {r2.text[:500]}")

print("\nDone.")
