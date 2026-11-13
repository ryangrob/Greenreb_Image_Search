"""Microsoft Graph client for browsing/downloading files from a SharePoint
document library, using interactive user sign-in (no client secret)."""

import json
import os
import sys
import threading
import urllib.parse

import requests
from msal import PublicClientApplication, SerializableTokenCache

import config

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPES = ["Sites.Read.All", "Files.ReadWrite.All"]
# Asked for separately, only when someone sends a support message, so the
# normal sign-in never needs this permission.
MAIL_SCOPES = ["Mail.Send"]
INDEX_FILENAME = ".imagesearch_sp_index.json"
# Each machine writes its own feedback file rather than all sharing one, so
# concurrent users can never overwrite each other's contributions.
FEEDBACK_PREFIX = ".imagesearch_feedback_"
FAVOURITES_PREFIX = ".imagesearch_favourites_"
# One file per shared collection: who owns it, who it is shared with, its photos.
SHARED_COLLECTION_PREFIX = ".imagesearch_collection_"
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg"}
# Payloads at/above this go through a resumable upload session rather than a
# simple PUT. Graph's chunked uploads require each chunk (except the last) to
# be a multiple of 320 KiB.
SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024
UPLOAD_CHUNK_SIZE = 16 * 320 * 1024  # 5 MiB
TOKEN_CACHE_PATH = os.path.join(
    (os.path.join(os.path.expanduser("~"), "Library", "Application Support", "Topshot")
     if sys.platform == "darwin"
     else os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "ImageSearch")),
    "msal_cache.bin"
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
        # Thumbnail downloads happen concurrently from a thread pool (see
        # tkinter_app._apply_folder_changes) - MSAL's token acquisition/cache
        # save isn't documented as thread-safe, so serialize it.
        self._token_lock = threading.Lock()

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

    def sign_in_silent(self):
        """Signs in from the saved session only - never opens a browser.
        For checks that run on their own in the background."""
        accounts = self._app.get_accounts()
        if not accounts:
            return False
        result = self._app.acquire_token_silent(SCOPES, account=accounts[0])
        if not result or "access_token" not in result:
            return False
        self._token = result["access_token"]
        self._save_cache()
        return True

    def send_mail(self, to, subject, body):
        """Sends an email from the signed-in user's own mailbox (it appears
        in their Sent Items). `to` is one address or a list of them. The
        first time, Microsoft may ask them to allow the app to send mail on
        their behalf."""
        recipients = [to] if isinstance(to, str) else [a for a in to if a]
        accounts = self._app.get_accounts()
        result = None
        if accounts:
            result = self._app.acquire_token_silent(MAIL_SCOPES, account=accounts[0])
        if not result:
            result = self._app.acquire_token_interactive(scopes=MAIL_SCOPES)
        if not result or "access_token" not in result:
            raise RuntimeError((result or {}).get("error_description", "Not allowed to send mail."))
        self._save_cache()
        resp = requests.post(
            f"{GRAPH_BASE}/me/sendMail",
            headers={"Authorization": f"Bearer {result['access_token']}",
                     "Content-Type": "application/json"},
            json={
                "message": {
                    "subject": subject,
                    "body": {"contentType": "Text", "content": body},
                    "toRecipients": [{"emailAddress": {"address": a}} for a in recipients],
                },
                "saveToSentItems": True,
            },
            timeout=30,
        )
        resp.raise_for_status()

    def account_email(self):
        """The signed-in user's address (their Microsoft sign-in name)."""
        accounts = self._app.get_accounts()
        return (accounts[0].get("username") or "").strip().lower() if accounts else ""

    def _headers(self):
        # Access tokens expire after ~60-90 minutes; a full indexing run (or
        # simply coming back to the app later) can easily outlast that, so
        # silently refresh via MSAL's cached refresh token before every call
        # instead of relying on the token captured once at sign_in().
        with self._token_lock:
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
        get picked up regardless of how Graph classifies them. Also skips
        macOS AppleDouble sidecar files (e.g. "._DSC1234.JPG") - tiny hidden
        metadata files macOS creates alongside real photos when copying to
        non-Mac storage, which otherwise pass the extension check and get
        wastefully "indexed" as if they were real images."""
        name = item.get("name", "")
        if name.startswith("._"):
            return False
        ext = name[name.rfind(".") :].lower() if "." in name else ""
        return ext in IMAGE_EXTENSIONS

    def get_delta_items(self, root_item_id, delta_link=None, status_callback=None):
        """Pages through Graph's delta query for a folder subtree.

        With delta_link=None, returns every current item (first-run/full
        bootstrap cost). With a previously-saved delta_link, returns only
        what's been added/changed/deleted since then. Returns
        (items, new_delta_link) - the caller should persist new_delta_link
        for next time. Raises DeltaExpired if a saved delta_link is no
        longer valid (Graph returns 410 Gone) - retry with delta_link=None.

        Calls status_callback(str) after each page so a full-tree listing
        (which can take a couple minutes on a large library) shows visible
        progress instead of looking stuck on "Checking for changes...".
        """
        self._resolve_site_and_drive()
        url = delta_link or f"{GRAPH_BASE}/drives/{self._drive_id}/items/{root_item_id}/delta"
        items = []
        new_delta_link = None
        try:
            while url:
                data = self._get(url)
                items.extend(data.get("value", []))
                if status_callback:
                    status_callback(f"Checking SharePoint for changes... ({len(items)} items found so far)")
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
            # The drive id is only known after this lookup. Every other
            # method that builds a /drives/... URL made it first; this one
            # did not, so downloading an original worked only if a sync had
            # already run in the same session - otherwise the request went
            # to /drives/None/... and failed with 400 Bad Request.
            self._resolve_site_and_drive()
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

    def download_json_file(self, folder_item_id, filename):
        """Returns a parsed JSON file from a folder, or None if not present.

        Fetches the specific filename directly via Graph's colon-path
        addressing instead of listing every child in the folder just to
        find one entry - one API call instead of a full (possibly paged)
        folder listing, which adds up across hundreds of folders.
        """
        self._resolve_site_and_drive()
        url = f"{GRAPH_BASE}/drives/{self._drive_id}/items/{folder_item_id}:/{filename}"
        try:
            entry = self._get(url)
        except requests.exceptions.HTTPError as exc:
            if exc.response is not None and exc.response.status_code == 404:
                return None
            raise
        download_url = entry.get("@microsoft.graph.downloadUrl")
        if not download_url:
            detail = self._get(f"{GRAPH_BASE}/drives/{self._drive_id}/items/{entry['id']}")
            download_url = detail.get("@microsoft.graph.downloadUrl")
        resp = requests.get(download_url, timeout=60)
        resp.raise_for_status()
        return resp.json()

    def delete_file(self, folder_item_id, filename):
        """Deletes a file from a folder (to the site's recycle bin). A file
        that is already gone counts as deleted."""
        self._resolve_site_and_drive()
        url = f"{GRAPH_BASE}/drives/{self._drive_id}/items/{folder_item_id}:/{filename}"
        resp = requests.delete(url, headers=self._headers(), timeout=30)
        if resp.status_code not in (204, 404):
            resp.raise_for_status()

    # ----- files on another site, by path (the releases folder) -----

    def _drive_for_site(self, site_path):
        cache = self.__dict__.setdefault("_site_drives", {})
        if site_path not in cache:
            site = self._get(f"{GRAPH_BASE}/sites/{config.SHAREPOINT_HOSTNAME}:{site_path}")
            cache[site_path] = self._get(f"{GRAPH_BASE}/sites/{site['id']}/drive")["id"]
        return cache[site_path]

    def download_bytes_by_path(self, site_path, file_path):
        """A file's contents by its path on a site, or None if it isn't there."""
        drive = self._drive_for_site(site_path)
        url = f"{GRAPH_BASE}/drives/{drive}/root:/{urllib.parse.quote(file_path)}"
        try:
            entry = self._get(url)
        except requests.exceptions.HTTPError as exc:
            if exc.response is not None and exc.response.status_code == 404:
                return None
            raise
        download_url = entry.get("@microsoft.graph.downloadUrl")
        if not download_url:
            detail = self._get(f"{GRAPH_BASE}/drives/{drive}/items/{entry['id']}")
            download_url = detail["@microsoft.graph.downloadUrl"]
        resp = requests.get(download_url, timeout=120)
        resp.raise_for_status()
        return resp.content

    def ensure_folder(self, site_path, folder_path):
        """Creates a folder path on a site, one level at a time, if missing."""
        drive = self._drive_for_site(site_path)
        built = ""
        for part in [p for p in folder_path.split("/") if p]:
            parent = built
            built = f"{built}/{part}" if built else part
            try:
                self._get(f"{GRAPH_BASE}/drives/{drive}/root:/{urllib.parse.quote(built)}")
                continue
            except requests.exceptions.HTTPError as exc:
                if exc.response is None or exc.response.status_code != 404:
                    raise
            parent_url = (
                f"{GRAPH_BASE}/drives/{drive}/root:/{urllib.parse.quote(parent)}:/children"
                if parent else f"{GRAPH_BASE}/drives/{drive}/root/children"
            )
            resp = requests.post(
                parent_url,
                headers={**self._headers(), "Content-Type": "application/json"},
                json={"name": part, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"},
                timeout=30,
            )
            if resp.status_code != 409:  # created by someone else meanwhile is fine
                resp.raise_for_status()

    def list_names_by_path(self, site_path, folder_path, prefix=""):
        """Names of the .json files in a folder on a site that start with
        `prefix` - an empty list if the folder doesn't exist yet."""
        drive = self._drive_for_site(site_path)
        url = f"{GRAPH_BASE}/drives/{drive}/root:/{urllib.parse.quote(folder_path)}:/children"
        names = []
        while url:
            try:
                data = self._get(url)
            except requests.exceptions.HTTPError as exc:
                if exc.response is not None and exc.response.status_code == 404:
                    return []
                raise
            for child in data.get("value", []):
                name = child.get("name", "")
                if name.startswith(prefix) and name.endswith(".json"):
                    names.append(name)
            url = data.get("@odata.nextLink")
        return names

    def delete_by_path(self, site_path, file_path):
        """Deletes a file on a site (to its recycle bin); already gone is fine."""
        drive = self._drive_for_site(site_path)
        url = f"{GRAPH_BASE}/drives/{drive}/root:/{urllib.parse.quote(file_path)}"
        resp = requests.delete(url, headers=self._headers(), timeout=30)
        if resp.status_code not in (204, 404):
            resp.raise_for_status()

    def upload_bytes_by_path(self, site_path, folder_path, filename, payload):
        """Uploads (creates/overwrites) a file of any size into a folder on a site."""
        drive = self._drive_for_site(site_path)
        target = urllib.parse.quote(f"{folder_path}/{filename}")
        resp = requests.post(
            f"{GRAPH_BASE}/drives/{drive}/root:/{target}:/createUploadSession",
            headers={**self._headers(), "Content-Type": "application/json"},
            json={"item": {"@microsoft.graph.conflictBehavior": "replace"}},
            timeout=60,
        )
        resp.raise_for_status()
        upload_url = resp.json()["uploadUrl"]
        total = len(payload)
        start = 0
        while start < total:
            end = min(start + UPLOAD_CHUNK_SIZE, total) - 1
            chunk_resp = requests.put(
                upload_url,
                headers={"Content-Range": f"bytes {start}-{end}/{total}"},
                data=payload[start:end + 1],
                timeout=120,
            )
            chunk_resp.raise_for_status()
            start = end + 1

    def download_index_file(self, folder_item_id):
        """Returns the parsed shared index dict for a folder, or None."""
        return self.download_json_file(folder_item_id, INDEX_FILENAME)

    def list_shared_files(self, folder_item_id, prefix):
        """Names of every per-machine shared file with a given prefix.

        Each machine writes only its own file, so merging is a matter of
        reading them all - no machine can overwrite another's contributions.
        """
        children = self.list_children(folder_item_id, by_id=True)
        return [
            c["name"]
            for c in children
            if c.get("name", "").startswith(prefix) and c.get("name", "").endswith(".json")
        ]

    def list_feedback_files(self, folder_item_id):
        return self.list_shared_files(folder_item_id, FEEDBACK_PREFIX)

    def get_index_file_size(self, folder_item_id):
        """Returns the byte size of a folder's shared index, or None if it
        isn't there. Metadata-only (no content download), so it's a cheap
        way to tell whether a folder's index already reached SharePoint.
        """
        self._resolve_site_and_drive()
        url = f"{GRAPH_BASE}/drives/{self._drive_id}/items/{folder_item_id}:/{INDEX_FILENAME}"
        try:
            return self._get(url).get("size")
        except requests.exceptions.HTTPError as exc:
            if exc.response is not None and exc.response.status_code == 404:
                return None
            raise

    def upload_index_file(self, folder_item_id, data):
        """Uploads (creates/overwrites) the shared index file into a folder."""
        return self.upload_json_file(folder_item_id, INDEX_FILENAME, data)

    def upload_json_file(self, folder_item_id, filename, data):
        """Uploads (creates/overwrites) a JSON file into a folder.

        Small payloads go via the simple-upload PUT; anything at/over
        SIMPLE_UPLOAD_LIMIT uses Graph's resumable upload-session API
        instead. A folder with a few hundred photos easily exceeds the
        simple-upload ceiling once base64'd thumbnails are included, and
        when that upload fails the folder's index never reaches SharePoint
        at all - so every other user re-embeds that whole folder from
        scratch, defeating the point of the shared index.
        """
        self._resolve_site_and_drive()
        payload = json.dumps(data).encode("utf-8")
        if len(payload) >= SIMPLE_UPLOAD_LIMIT:
            return self._upload_file_chunked(folder_item_id, filename, payload)

        url = f"{GRAPH_BASE}/drives/{self._drive_id}/items/{folder_item_id}:/{filename}:/content"
        resp = requests.put(
            url,
            headers={**self._headers(), "Content-Type": "application/json"},
            data=payload,
            timeout=60,
        )
        resp.raise_for_status()
        return resp.json()

    def _upload_file_chunked(self, folder_item_id, filename, payload):
        """Uploads a large file via Graph's resumable upload session."""
        self._resolve_site_and_drive()  # cheap once resolved; see download_file
        session_url = (
            f"{GRAPH_BASE}/drives/{self._drive_id}/items/"
            f"{folder_item_id}:/{filename}:/createUploadSession"
        )
        resp = requests.post(
            session_url,
            headers={**self._headers(), "Content-Type": "application/json"},
            json={"item": {"@microsoft.graph.conflictBehavior": "replace"}},
            timeout=60,
        )
        resp.raise_for_status()
        upload_url = resp.json()["uploadUrl"]

        total = len(payload)
        start = 0
        result = {}
        while start < total:
            end = min(start + UPLOAD_CHUNK_SIZE, total) - 1
            chunk = payload[start : end + 1]
            # The upload URL is already pre-authenticated - deliberately no
            # Authorization header here (Graph rejects the request if sent).
            chunk_resp = requests.put(
                upload_url,
                headers={"Content-Range": f"bytes {start}-{end}/{total}"},
                data=chunk,
                timeout=120,
            )
            chunk_resp.raise_for_status()
            if chunk_resp.status_code in (200, 201) and chunk_resp.content:
                result = chunk_resp.json()
            start = end + 1
        return result
