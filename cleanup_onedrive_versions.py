#!/usr/bin/env python3
"""
cleanup_onedrive_versions.py

Walks a OneDrive for Business / SharePoint document library (via Microsoft
Graph) and deletes every non-current version of every file, freeing up the
storage that accumulates from OneDrive's version history feature.

Auth: app-only (client credentials) using the same Azure AD app registration
your upload Action already uses. Requires the Files.ReadWrite.All
application permission with admin consent (Sites.ReadWrite.All also works
if your drive is backed by a SharePoint site).

Required environment variables (set as GitHub Actions secrets):
    AZURE_TENANT_ID
    AZURE_CLIENT_ID
    AZURE_CLIENT_SECRET
    GRAPH_DRIVE_ID        - the id of the drive containing your files
Optional:
    GRAPH_FOLDER_PATH     - limit the walk to this folder, e.g. "Data/CSVs"
                            (path relative to the drive root). If unset,
                            the whole drive is scanned.
    DRY_RUN               - "true" to log what would be deleted without
                            deleting anything. Defaults to "false".
    KEEP_VERSIONS         - how many of the most recent versions to keep
                            per file. Defaults to "1" (current version only).
"""

import os
import sys
import time
import requests
import msal

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def get_access_token():
    tenant_id = os.environ["AZURE_TENANT_ID"]
    client_id = os.environ["AZURE_CLIENT_ID"]
    client_secret = os.environ["AZURE_CLIENT_SECRET"]

    app = msal.ConfidentialClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential=client_secret,
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"Failed to acquire token: {result.get('error_description')}")
    return result["access_token"]


def graph_get(session, url):
    resp = session.get(url)
    if resp.status_code == 429:
        retry_after = int(resp.headers.get("Retry-After", "5"))
        time.sleep(retry_after)
        return graph_get(session, url)
    resp.raise_for_status()
    return resp.json()


def iter_files(session, drive_id, folder_path=None):
    """Yield every file (not folder) item under the given drive/folder, recursively."""
    if folder_path:
        start_url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{folder_path}:/children"
    else:
        start_url = f"{GRAPH_BASE}/drives/{drive_id}/root/children"

    stack = [start_url]
    while stack:
        url = stack.pop()
        while url:
            data = graph_get(session, url)
            for item in data.get("value", []):
                if "folder" in item:
                    stack.append(f"{GRAPH_BASE}/drives/{drive_id}/items/{item['id']}/children")
                elif "file" in item:
                    yield item
            url = data.get("@odata.nextLink")


def cleanup_versions(session, drive_id, item, keep_versions, dry_run):
    item_id = item["id"]
    name = item.get("name", item_id)
    versions_url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/versions"
    data = graph_get(session, versions_url)
    versions = data.get("value", [])

    # Graph returns versions newest-first; keep the first `keep_versions`.
    if len(versions) <= keep_versions:
        return 0

    to_delete = versions[keep_versions:]
    deleted = 0
    for v in to_delete:
        v_id = v["id"]
        del_url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/versions/{v_id}"
        if dry_run:
            print(f"  [DRY RUN] would delete version {v_id} of '{name}'")
            deleted += 1
            continue
        resp = session.delete(del_url)
        if resp.status_code == 429:
            retry_after = int(resp.headers.get("Retry-After", "5"))
            time.sleep(retry_after)
            resp = session.delete(del_url)
        if resp.status_code in (204, 200):
            deleted += 1
        else:
            print(f"  WARNING: failed to delete version {v_id} of '{name}': "
                  f"{resp.status_code} {resp.text}")
    return deleted


def main():
    drive_id = os.environ["GRAPH_DRIVE_ID"]
    folder_path = os.environ.get("GRAPH_FOLDER_PATH") or None
    dry_run = os.environ.get("DRY_RUN", "false").strip().lower() == "true"
    keep_versions = int(os.environ.get("KEEP_VERSIONS", "1"))

    token = get_access_token()
    session = requests.Session()
    session.headers.update({"Authorization": f"Bearer {token}"})

    total_files = 0
    total_deleted = 0

    print(f"Scanning drive {drive_id}"
          f"{f' / {folder_path}' if folder_path else ''} "
          f"(dry_run={dry_run}, keep_versions={keep_versions})")

    for item in iter_files(session, drive_id, folder_path):
        total_files += 1
        deleted = cleanup_versions(session, drive_id, item, keep_versions, dry_run)
        if deleted:
            print(f"'{item.get('name')}': deleted {deleted} old version(s)")
        total_deleted += deleted

    print(f"\nDone. Scanned {total_files} file(s), deleted {total_deleted} old version(s).")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
