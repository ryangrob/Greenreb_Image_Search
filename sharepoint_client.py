"""Microsoft Graph client for browsing/downloading files from a SharePoint
document library, using interactive user sign-in (no client secret)."""

import json
import os
import urllib.parse

import requests
from msal import PublicClientApplication, SerializableTokenCache

import config

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPES = ["Sites.Read.All", "Files.ReadWrite.All"]
INDEX_FILENAME = ".imagesearch_sp_index.json"
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg"}
TOKEN_CACHE_PATH = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "ImageSearch", "msal_cache.bin"
)


class DeltaExpired(Exception):
    """Raised when a saved delta token is no longer valid (Graph returns 410)."""


class SharePointClient:
    def __init__(self):
        # Persisted to disk so a signed-in user isn't prompted to sign in
        # again every single time the app launches - without this, MSAL's
        # default in-memory cache is thrown away when the process exits.
        self._cache = SerializableTokenCache()
        if os.path.exists(TOKEN_CACHE_PATH):
            try:
                with open(TOKEN_CACHE_PATH, "r", encoding="utf-8") as f:
                    self._cache.deserialize(f.read())
            except OSError:
                pass
        self._app = PublicClientApplication(
            config.CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{config.TENANT_ID}",
            token_cache=self._cache,
        )
        self._token = None
        self._site_id = None
        self._drive_id = None

    def _save_cache(self):
        if self._cache.has_state_changed:
            os.makedirs(os.path.dirname(TOKEN_CACHE_PATH), exist_ok=True)
            with open(TOKEN_CACHE_PATH, "w", encoding="utf-8") as f:
                f.write(self._cache.serialize())

    def sign_in(self):
        accounts = self._app.get_accounts()
        result = None
        if accounts:
            result = self._app.acquire_token_silent(SCOPES, account=accounts[0])
        if not result:
            result = self._app.acquire_token_interactive(scopes=SCOPES)
        if not result or "access_token" not in result:
            error = (result or {}).get("error_description", "Sign-in failed.")
            raise RuntimeError(error)
        self._token = result["access_token"]
        self._save_cache()

    def _headers(self):
        # Access tokens expire after ~60-90 minutes; a full indexing run (or
        # simply coming back to the app later) can easily outlast that, so
        # silently refresh via MSAL's cached refresh token before every call
        # instead of relying on the token captured once at sign_in().
        accounts = self._app.get_accounts()
        if accounts:
            result = self._app.acquire_token_silent(SCOPES, account=accounts[0])
            if result and "access_token" in result:
                self._token = result["access_token"]
                self._save_cache()
        return {"Authorization": f"Bearer {self._token}"}

    def _get(self, url):
        resp = requests.get(url, headers=self._headers(), timeout=30)
        resp.raise_for_status()
        return resp.json()

    def _resolve_site_and_drive(self):
        if self._site_id:
            return
        site_url = f"{GRAPH_BASE}/sites/{config.SHAREPOINT_HOSTNAME}:{config.SHAREPOINT_SITE_PATH}"
        site = self._get(site_url)
        self._site_id = site["id"]
        drive = self._get(f"{GRAPH_BASE}/sites/{self._site_id}/drive")
        self._drive_id = drive["id"]

    def get_search_root_item(self):
        """Resolves the driveItem (id, eTag, ...) for SHAREPOINT_SEARCH_ROOT_PATH itself."""
        return self.get_item_by_path(config.SHAREPOINT_SEARCH_ROOT_PATH, by_id=False)

    @staticmethod
    def is_wanted_image(item):
        """Extension-based check (not mimeType) so formats like .dng never
        get picked up regardless of how Graph classifies them."""
        name = item.get("name", "")
        ext = name[name.rfind(".") :].lower() if "." in name else ""
        return ext in IMAGE_EXTENSIONS

    def get_delta_items(self, root_item_id, delta_link=None):
        """Pages through Graph's delta query for a folder subtree.

        With delta_link=None, returns every current item (first-run/full
        bootstrap cost). With a previously-saved delta_link, returns only
        what's been added/changed/deleted since then. Returns
        (items, new_delta_link) - the caller should persist new_delta_link
        for next time. Raises DeltaExpired if a saved delta_link is no
        longer valid (Graph returns 410 Gone) - retry with delta_link=None.
        """
        self._resolve_site_and_drive()
        url = delta_link or f"{GRAPH_BASE}/drives/{self._drive_id}/items/{root_item_id}/delta"
        items = []
        new_delta_link = None
        try:
            while url:
                data = self._get(url)
                items.extend(data.get("value", []))
                url = data.get("@odata.nextLink")
                if "@odata.deltaLink" in data:
                    new_delta_link = data["@odata.deltaLink"]
        except requests.exceptions.HTTPError as exc:
            if delta_link and exc.response is not None and exc.response.status_code == 410:
                raise DeltaExpired from exc
            raise
        return items, new_delta_link

    def list_children(self, path_or_item_id, by_id):
        """Lists a folder's children, either by drive-relative path or by driveItem id."""
        self._resolve_site_and_drive()
        if by_id:
            url = f"{GRAPH_BASE}/drives/{self._drive_id}/items/{path_or_item_id}/children"
        else:
            encoded = urllib.parse.quote(path_or_item_id)
            url = f"{GRAPH_BASE}/drives/{self._drive_id}/root:/{encoded}:/children"
        items = []
        while url:
            data = self._get(url)
            items.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
        return items

    def download_file(self, item, dest_path):
        download_url = item.get("@microsoft.graph.downloadUrl")
        if not download_url:
            detail = self._get(f"{GRAPH_BASE}/drives/{self._drive_id}/items/{item['id']}")
            download_url = detail["@microsoft.graph.downloadUrl"]
        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
        resp = requests.get(download_url, stream=True, timeout=60)
        resp.raise_for_status()
        with open(dest_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=1 << 16):
                f.write(chunk)

    def get_item_by_path(self, path_or_item_id, by_id):
        """Resolves a folder's own driveItem (id, eTag, ...) - not its children."""
        self._resolve_site_and_drive()
        if by_id:
            url = f"{GRAPH_BASE}/drives/{self._drive_id}/items/{path_or_item_id}"
        else:
            encoded = urllib.parse.quote(path_or_item_id)
            url = f"{GRAPH_BASE}/drives/{self._drive_id}/root:/{encoded}"
        return self._get(url)

    def get_thumbnail_url(self, item, size="medium"):
        self._resolve_site_and_drive()
        url = f"{GRAPH_BASE}/drives/{self._drive_id}/items/{item['id']}/thumbnails"
        data = self._get(url)
        sets = data.get("value", [])
        if not sets:
            return None
        return sets[0].get(size, {}).get("url")

    def download_thumbnail(self, item, dest_path, size="medium"):
        thumb_url = self.get_thumbnail_url(item, size=size)
        if not thumb_url:
            return False
        resp = requests.get(thumb_url, stream=True, timeout=30)
        resp.raise_for_status()
        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
        with open(dest_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=1 << 16):
                f.write(chunk)
        return True

    def download_index_file(self, folder_item_id):
        """Returns the parsed shared index dict for a folder, or None if not present."""
        self._resolve_site_and_drive()
        children = self.list_children(folder_item_id, by_id=True)
        entry = next((c for c in children if c.get("name") == INDEX_FILENAME), None)
        if entry is None:
            return None
        download_url = entry.get("@microsoft.graph.downloadUrl")
        if not download_url:
            detail = self._get(f"{GRAPH_BASE}/drives/{self._drive_id}/items/{entry['id']}")
            download_url = detail.get("@microsoft.graph.downloadUrl")
        resp = requests.get(download_url, timeout=60)
        resp.raise_for_status()
        return resp.json()

    def upload_index_file(self, folder_item_id, data):
        """Uploads (creates/overwrites) the shared index file into a folder.

        Uses the simple-upload PUT, which caps at 4MB - fine for base64'd
        thumbnails at this scale, but would need the resumable upload-session
        API if a folder's index ever grows past that.
        """
        self._resolve_site_and_drive()
        url = f"{GRAPH_BASE}/drives/{self._drive_id}/items/{folder_item_id}:/{INDEX_FILENAME}:/content"
        payload = json.dumps(data).encode("utf-8")
        resp = requests.put(
            url,
            headers={**self._headers(), "Content-Type": "application/json"},
            data=payload,
            timeout=60,
        )
        resp.raise_for_status()
        return resp.json()
