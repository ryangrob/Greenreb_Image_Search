"""Image Search desktop app.

Point it at a folder of photos, index them once, then search by typing a
description ("a dog on a beach") or by picking a reference image.
Indexing uses a local CLIP model, so search works fully offline after the
model has been downloaded once.
"""

import base64
import concurrent.futures
import datetime
import glob
import json
import os
import queue
import re
import subprocess
import sys
import threading
import time
import traceback
import unicodedata
import tkinter as tk
from tkinter import filedialog, messagebox

# When running as a bundled .exe (PyInstaller --onedir), redirect the CLIP
# model's cache to the copy bundled alongside the exe (see build.bat /
# download_model.py) so end users never wait on the ~350-600MB first-run
# download. Must happen before anything imports huggingface_hub, which reads
# HF_HOME once at import time - hence this sits above every other import.
if getattr(sys, "frozen", False):
    os.environ["HF_HOME"] = os.path.join(sys._MEIPASS, "model_cache")

import customtkinter as ctk
import numpy as np
from PIL import Image

from search_engine import ImageSearchEngine, MODEL_TAG
from sharepoint_client import DeltaExpired, SharePointClient

CRASH_LOG_PATH = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "ImageSearch", "crash.log"
)
REUSE_LOG_PATH = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "ImageSearch", "reuse_debug.log"
)


def _log_crash(context, exc_info):
    """Appends a timestamped traceback to crash.log - so an unexpected failure
    (e.g. during a long SharePoint indexing run) leaves evidence behind
    instead of the app just silently disappearing."""
    try:
        os.makedirs(os.path.dirname(CRASH_LOG_PATH), exist_ok=True)
        with open(CRASH_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"\n--- {datetime.datetime.now().isoformat()} [{context}] ---\n")
            f.write("".join(traceback.format_exception(*exc_info)))
    except OSError:
        pass


def _log_reuse(line):
    """Temporary diagnostic log for tracking down why the SharePoint
    per-folder cache/shared-index reuse check does or doesn't hit."""
    try:
        os.makedirs(os.path.dirname(REUSE_LOG_PATH), exist_ok=True)
        with open(REUSE_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"{datetime.datetime.now().isoformat()} {line}\n")
    except OSError:
        pass


THUMB_SIZE = (128, 128)
GRID_COLUMNS = 5
DOWNLOAD_WORKERS = 8
EMBED_BATCH_SIZE = 16

# Added to an image's similarity score when the query matches the name of the
# SharePoint folder it lives in. CLIP cannot know event names ("Weihnachtsfeier
# 2024"), so without this those photos are unreachable by name. Sized against
# measured score spread (~0.15-0.30) to promote strongly without overwhelming a
# genuinely better visual match.
FOLDER_NAME_BONUS = 0.08
# Added when people previously opened an image for this same query. Capped so a
# popular result is promoted rather than permanently pinned to the top.
FEEDBACK_BONUS = 0.04
FEEDBACK_BONUS_CAP = 0.10
DEFAULT_TOP_K = 50

BG_COLOR = "#0d1b2a"
CARD_COLOR = "#0a141f"
ACCENT_COLOR = "#0b4c8c"
ACCENT_ACTIVE = "#004fff"
FIELD_BG = "#0a141f"
TEXT_COLOR = "#eef3fa"
MUTED_TEXT_COLOR = "#8fa3b8"

CARD_RADIUS = 14
BUTTON_RADIUS = 17

LOGO_HEIGHT = 32


def _resource_path(*parts):
    # PyInstaller (--onedir) unpacks bundled data next to the exe under
    # sys._MEIPASS; running from source, it's just this file's folder.
    base = sys._MEIPASS if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, *parts)


def _load_logo_image(height=LOGO_HEIGHT):
    try:
        img = Image.open(_resource_path("assets", "logo.png")).convert("RGBA")
        ratio = height / img.height
        width = max(1, round(img.width * ratio))
        return ctk.CTkImage(light_image=img, dark_image=img, size=(width, height))
    except Exception:
        return None


def normalize_text(text):
    """Lowercases and strips accents so folder-name matching is tolerant of
    German spelling variants - "Weihnachtsfeier" vs "weihnachtsfeier",
    "Grunreb" vs "Grünreb", "Strasse" vs "Straße"."""
    text = (text or "").lower().replace("ß", "ss")
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def folder_name_match(query, folder_name):
    """How strongly a query matches a folder name, from 0.0 to 1.0.

    Substring matching (not just whole words) because German compounds the
    words a searcher is likely to type: "feier" should match
    "Weihnachtsfeier", and "bier" should match "Bierfest".
    """
    q = normalize_text(query)
    f = normalize_text(folder_name)
    if not q or not f:
        return 0.0
    terms = [t for t in q.split() if len(t) > 2]
    if not terms:
        return 0.0
    hits = sum(1 for t in terms if t in f)
    return hits / len(terms)


def open_in_system_viewer(path):
    if sys.platform.startswith("win"):
        os.startfile(path)  # noqa: S606 (Windows-only convenience)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


def _make_button(parent, text, command, width=140):
    return ctk.CTkButton(
        parent,
        text=text,
        command=command,
        width=width,
        height=34,
        corner_radius=BUTTON_RADIUS,
        fg_color=ACCENT_COLOR,
        hover_color=ACCENT_ACTIVE,
        text_color=TEXT_COLOR,
        font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
    )


class ImageSearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Greenreb_Image_Search")
        self.root.geometry("980x720")
        self.root.minsize(760, 520)

        self.engine = ImageSearchEngine()
        self.sp_client = None
        self._sp_mode_active = False
        self._sp_total_images = 0
        self._sp_done_images = 0
        self.event_queue = queue.Queue()
        self.worker_thread = None
        self.cancel_requested = False
        self.thumbnail_refs = []  # keep CTkImage refs alive
        self._last_query = ""
        self._feedback = None

        self._build_widgets()
        self.root.after(100, self._poll_queue)

    # ---------- UI construction ----------

    def _build_widgets(self):
        header = ctk.CTkFrame(self.root, fg_color="transparent")
        header.pack(fill=tk.X, padx=12, pady=(12, 0))
        self._logo_image = _load_logo_image()
        if self._logo_image:
            ctk.CTkLabel(header, image=self._logo_image, text="").pack(side=tk.LEFT)

        top = ctk.CTkFrame(self.root, fg_color=CARD_COLOR, corner_radius=CARD_RADIUS)
        top.pack(fill=tk.X, padx=12, pady=12)

        self.folder_var = tk.StringVar(value="No folder selected")
        _make_button(top, "Choose Folder...", self._choose_folder, width=150).pack(
            side=tk.LEFT, padx=(14, 8), pady=14
        )
        ctk.CTkLabel(
            top, textvariable=self.folder_var, width=340, anchor="w",
            text_color=MUTED_TEXT_COLOR,
        ).pack(side=tk.LEFT, padx=8, pady=14)
        self.index_button = _make_button(top, "Index Folder", self._start_indexing, width=130)
        self.index_button.pack(side=tk.LEFT, padx=4, pady=14)
        self.index_button.configure(state=tk.DISABLED)
        self.sharepoint_button = _make_button(
            top, "Search Marketing Photos", self._browse_sharepoint, width=190
        )
        self.sharepoint_button.pack(side=tk.LEFT, padx=(4, 14), pady=14)

        search_frame = ctk.CTkFrame(self.root, fg_color=CARD_COLOR, corner_radius=CARD_RADIUS)
        search_frame.pack(fill=tk.X, padx=12, pady=(0, 12))

        self.query_var = tk.StringVar()
        query_entry = ctk.CTkEntry(
            search_frame, textvariable=self.query_var, fg_color=FIELD_BG,
            text_color=TEXT_COLOR, border_width=0, corner_radius=10,
        )
        query_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(14, 8), pady=14)
        query_entry.bind("<Return>", lambda _e: self._start_text_search())

        self.search_button = _make_button(
            search_frame, "Search", self._start_text_search, width=100
        )
        self.search_button.pack(side=tk.LEFT, padx=(0, 14), pady=14)
        self.search_button.configure(state=tk.DISABLED)

        status_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        status_frame.pack(fill=tk.X, padx=12, pady=(0, 8))
        self.status_var = tk.StringVar(value="Choose a folder to begin.")
        ctk.CTkLabel(
            status_frame, textvariable=self.status_var, text_color=TEXT_COLOR, anchor="w",
        ).pack(side=tk.LEFT)
        self.progress = ctk.CTkProgressBar(
            status_frame, width=200, fg_color=FIELD_BG, progress_color=ACCENT_COLOR,
        )
        self.progress.set(0)
        self.progress.pack(side=tk.RIGHT)

        # Scrollable results grid - CTkScrollableFrame already handles its
        # own internal canvas, scrollbar, and mousewheel binding.
        self.results_frame = ctk.CTkScrollableFrame(self.root, fg_color=BG_COLOR)
        self.results_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 12))

    # ---------- folder selection & indexing ----------

    def _choose_folder(self):
        folder = filedialog.askdirectory(title="Choose a folder of images")
        if not folder:
            return
        self._sp_mode_active = False
        self.folder_var.set(folder)
        self.index_button.configure(state=tk.NORMAL)

        count = self.engine.load_cached_index(folder)
        self._update_search_buttons_state()
        if count:
            self.status_var.set(f"Loaded {count} cached images. Click 'Index Folder' to refresh.")
        else:
            self.status_var.set("No cached index found. Click 'Index Folder' to build one.")

    def _start_indexing(self):
        folder = self.folder_var.get()
        if not os.path.isdir(folder):
            return
        self._set_busy(True)
        self.progress.configure(mode="indeterminate")
        self.progress.start()

        def work():
            def on_progress(done, total):
                self.event_queue.put(("progress", done, total))

            def on_status(msg):
                self.event_queue.put(("status", msg))

            try:
                count = self.engine.build_index(
                    folder, progress_callback=on_progress, status_callback=on_status
                )
                self.event_queue.put(("index_done", count))
            except Exception as exc:  # surfaces to the GUI instead of crashing the thread
                self.event_queue.put(("error", str(exc)))

        self._run_in_background(work)

    def _start_text_search(self):
        query = self.query_var.get().strip()
        if not query:
            return
        if self.engine.embeddings is None or len(self.engine.paths) == 0:
            messagebox.showinfo("Image Search", "Index the folder first before searching.")
            return
        self._last_query = query
        self._set_busy(True)
        self.progress.configure(mode="indeterminate")
        self.progress.start()

        def work():
            def on_status(msg):
                self.event_queue.put(("status", msg))

            try:
                bonus = self._compute_bonus(query)
                results = self.engine.search_text(
                    query, DEFAULT_TOP_K, status_callback=on_status, bonus=bonus
                )
                self.event_queue.put(("search_done", results))
            except Exception as exc:
                self.event_queue.put(("error", str(exc)))

        self._run_in_background(work)

    def _run_in_background(self, fn):
        def guarded():
            try:
                fn()
            except Exception:
                # Every current work() already has its own try/except that
                # routes failures to the "error" event - this is a backstop
                # so a future/unexpected failure still leaves a traceback in
                # crash.log instead of the thread just dying silently.
                _log_crash("background thread", sys.exc_info())
                raise

        self.worker_thread = threading.Thread(target=guarded, daemon=True)
        self.worker_thread.start()

    def _set_busy(self, busy):
        # "Index Folder" re-runs local-folder indexing, which doesn't apply
        # in SharePoint mode (that index is already built via the shared
        # thumbnail sync) - keep it disabled while SharePoint mode is active.
        self.index_button.configure(
            state=tk.NORMAL if (not busy and not self._sp_mode_active) else tk.DISABLED
        )
        self.sharepoint_button.configure(state=tk.DISABLED if busy else tk.NORMAL)
        if busy:
            self.search_button.configure(state=tk.DISABLED)
        else:
            self._update_search_buttons_state()

    def _update_search_buttons_state(self):
        has_index = self.engine.embeddings is not None and len(self.engine.paths) > 0
        self.search_button.configure(state=tk.NORMAL if has_index else tk.DISABLED)

    # ---------- search-time ranking signals ----------

    def _feedback_path(self):
        local_folder, _t, _f, _i = self._sp_cache_dirs()
        return os.path.join(local_folder, "click_feedback.json")

    def _load_feedback(self):
        """{normalized query: {item_id: times opened}}"""
        if self._feedback is None:
            self._feedback = {}
            try:
                with open(self._feedback_path(), "r", encoding="utf-8") as f:
                    self._feedback = json.load(f)
            except (json.JSONDecodeError, OSError):
                pass
        return self._feedback

    def _record_click(self, item_id):
        """Remembers that an image was opened for the last query run.

        Opening a result is the clearest signal available that it was the
        right answer, and it was previously discarded. Recording it lets
        frequently-searched topics improve with use, without any labelling.
        """
        query = normalize_text(self._last_query)
        if not query or not item_id:
            return
        feedback = self._load_feedback()
        per_query = feedback.setdefault(query, {})
        per_query[item_id] = per_query.get(item_id, 0) + 1
        try:
            with open(self._feedback_path(), "w", encoding="utf-8") as f:
                json.dump(feedback, f, ensure_ascii=False)
        except OSError:
            pass

    def _compute_bonus(self, query):
        """Per-image score adjustments the image embedding can't provide:
        the name of the folder an image sits in, and prior opens for this
        same query. Returns None when neither applies."""
        if not self.engine.meta:
            return None
        feedback_for_query = self._load_feedback().get(normalize_text(query), {})
        folder_cache = {}
        bonus = np.zeros(len(self.engine.meta), dtype=np.float32)
        any_bonus = False

        for i, meta in enumerate(self.engine.meta):
            if not meta:
                continue
            folder = meta.get("folder") or ""
            if folder:
                if folder not in folder_cache:
                    folder_cache[folder] = folder_name_match(query, folder)
                score = folder_cache[folder]
                if score:
                    bonus[i] += FOLDER_NAME_BONUS * score
                    any_bonus = True
            opens = feedback_for_query.get(meta.get("item_id"))
            if opens:
                bonus[i] += min(FEEDBACK_BONUS * opens, FEEDBACK_BONUS_CAP)
                any_bonus = True

        return bonus if any_bonus else None

    # ---------- SharePoint browsing ----------

    def _browse_sharepoint(self):
        # A browser window only actually appears the very first time (or if
        # the saved sign-in has expired) - after that, sign-in is silent.
        self.status_var.set("Signing in to Microsoft...")
        self._set_busy(True)

        def work():
            try:
                if self.sp_client is None:
                    self.sp_client = SharePointClient()
                self.sp_client.sign_in()
                self.event_queue.put(("sp_signed_in",))
            except Exception as exc:
                self.event_queue.put(("error", f"Sign-in failed: {exc}"))

        self._run_in_background(work)

    def _sp_cache_dirs(self):
        local_folder = os.path.join(
            os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
            "ImageSearch",
            "sharepoint_cache",
            "FotosVideos",
        )
        return (
            local_folder,
            os.path.join(local_folder, "thumbs"),
            os.path.join(local_folder, "full"),
            os.path.join(local_folder, "folder_indexes"),
        )

    def _download_thumbnail_with_retry(self, item, thumb_path):
        """One quick retry on a transient failure (e.g. a brief network
        hiccup) before giving up - cuts down how often the pending-retry/
        next-run mechanism has to be relied on for otherwise-fine
        connections. Does NOT retry a clean "no thumbnail available" (a
        return of False) since that isn't a transient error."""
        try:
            return self.sp_client.download_thumbnail(item, thumb_path)
        except Exception:
            time.sleep(1.5)
            return self.sp_client.download_thumbnail(item, thumb_path)

    def _upload_folder_index(self, folder_id, data, local_index_path, uploaded_ok):
        """Uploads one folder's shared index, recording success/failure.

        `uploaded_ok` is the set of folder ids known to have reached
        SharePoint - persisted across runs so a folder whose upload failed
        can be retried later even though delta will never report it as
        changed again.
        """
        try:
            self.sp_client.upload_index_file(folder_id, data)
            uploaded_ok.add(folder_id)
            return True
        except Exception as exc:
            uploaded_ok.discard(folder_id)
            # Logged, not just flashed in the status bar: a failed upload here
            # means this folder's work never reaches SharePoint, so every
            # other user silently re-embeds the whole folder. That's worth
            # leaving durable evidence of rather than letting it scroll past.
            try:
                size_mb = os.path.getsize(local_index_path) / (1024 * 1024)
            except OSError:
                size_mb = -1
            _log_reuse(
                f"UPLOAD FAILED folder={folder_id} size={size_mb:.2f}MB "
                f"items={len(data.get('items', {}))} error={exc!r}"
            )
            self.event_queue.put(("status", f"Warning: couldn't upload shared index: {exc}"))
            return False

    def _backfill_folder_names(self, folder_ids, folder_names):
        """Looks up names for folders we don't have one for yet.

        Delta only reports folders that changed, so on an incremental run
        almost no folder names arrive - including for a library indexed
        before folder names were captured at all. This fills the gap once;
        afterwards the names are cached and this does nothing.
        """
        missing = [fid for fid in folder_ids if fid not in folder_names]
        if not missing:
            return 0

        done = 0
        with concurrent.futures.ThreadPoolExecutor(max_workers=DOWNLOAD_WORKERS) as pool:
            futures = {
                pool.submit(self.sp_client.get_item_by_path, fid, True): fid for fid in missing
            }
            for future in concurrent.futures.as_completed(futures):
                fid = futures[future]
                done += 1
                if done % 25 == 0 or done == len(missing):
                    self.event_queue.put(
                        ("status", f"Reading folder names... {done}/{len(missing)}")
                    )
                try:
                    name = (future.result() or {}).get("name")
                    if name:
                        folder_names[fid] = name
                except Exception:
                    continue  # a folder we can't read just misses out on name matching
        _log_reuse(f"FOLDER NAMES: looked_up={len(missing)} known_total={len(folder_names)}")
        return len(missing)

    def _repair_unuploaded_indexes(self, indexes_dir, uploaded_ok):
        """Re-uploads any locally-built folder index that never made it to
        SharePoint.

        Without this, a folder whose upload failed once stays broken
        permanently: delta only reports genuinely changed folders, so an
        unchanged-but-never-uploaded folder is never revisited, and every
        other user keeps re-embedding it from scratch.
        """
        candidates = []
        for path in glob.glob(os.path.join(indexes_dir, "*.json")):
            folder_id = os.path.splitext(os.path.basename(path))[0]
            if folder_id not in uploaded_ok:
                candidates.append((folder_id, path))
        if not candidates:
            return 0

        # Confirm against SharePoint before re-uploading. On the very first
        # run after this check was introduced nothing is recorded as
        # uploaded yet, so without this every folder would be re-uploaded
        # (hundreds of MB) even though most are already fine. A
        # metadata-only size lookup is far cheaper than an upload, so only
        # genuinely missing or mismatched indexes get sent.
        pending = []
        for i, (folder_id, path) in enumerate(candidates, 1):
            self.event_queue.put(
                ("status", f"Verifying shared index {i}/{len(candidates)}...")
            )
            try:
                remote_size = self.sp_client.get_index_file_size(folder_id)
                if remote_size is not None and remote_size == os.path.getsize(path):
                    uploaded_ok.add(folder_id)
                    continue
            except Exception:
                pass  # can't confirm - fall through and re-upload to be safe
            pending.append((folder_id, path))

        already_ok = len(candidates) - len(pending)
        repaired = 0
        for i, (folder_id, path) in enumerate(pending, 1):
            self.event_queue.put(
                ("status", f"Sharing index for folder {i}/{len(pending)} with the team...")
            )
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except (json.JSONDecodeError, OSError):
                continue
            if self._upload_folder_index(folder_id, data, path, uploaded_ok):
                repaired += 1
        _log_reuse(
            f"REPAIR PASS: checked={len(candidates)} already_on_sharepoint={already_ok} "
            f"needed_upload={len(pending)} succeeded={repaired}"
        )
        return repaired

    def _apply_folder_changes(
        self, folder_id, changed_items, deleted_ids, thumbs_dir, indexes_dir, failed_items,
        uploaded_ok,
    ):
        """Applies a delta-reported set of adds/updates/deletes to one
        SharePoint folder's shared, thumbnail-based index.

        Reuses a shared `.imagesearch_sp_index.json` uploaded into the
        SharePoint folder itself when possible (keyed by driveItem id +
        eTag), so only genuinely new/changed images ever need their
        thumbnail downloaded and embedded - by anyone, on any machine.

        Any item whose download/embed fails (e.g. lost connectivity partway
        through a run) is recorded into `failed_items` (item_id -> raw item
        dict) so the caller can persist it for a forced retry next run -
        delta sync alone would never re-surface it, since nothing about an
        unchanged-but-previously-failed item looks "changed" to SharePoint.
        Returns the folder's updated {item_id: entry} dict.
        """
        local_index_path = os.path.join(indexes_dir, f"{folder_id}.json")
        local_items = {}
        if os.path.exists(local_index_path):
            try:
                with open(local_index_path, "r", encoding="utf-8") as f:
                    local_items = json.load(f).get("items", {})
            except (json.JSONDecodeError, OSError):
                local_items = {}

        had_local_cache = len(local_items) > 0
        remote_items = {}
        remote_index_error = None
        if changed_items:
            try:
                remote_index = self.sp_client.download_index_file(folder_id)
                remote_items = (remote_index or {}).get("items", {})
            except Exception as exc:
                remote_index_error = str(exc)
                self.event_queue.put(("status", f"Couldn't read shared index: {exc}"))

        for item_id in deleted_ids:
            local_items.pop(item_id, None)

        reused_count = 0
        to_download = []  # (item, thumb_path) pairs needing the full pipeline

        for item in changed_items:
            item_id = item["id"]
            etag = item.get("eTag")
            thumb_path = os.path.join(thumbs_dir, f"{item_id}.jpg")
            existing = local_items.get(item_id) or remote_items.get(item_id)

            if existing and existing.get("etag") == etag:
                reused_count += 1
                local_items[item_id] = existing
                if not os.path.exists(thumb_path) and existing.get("thumbnail_b64"):
                    with open(thumb_path, "wb") as f:
                        f.write(base64.b64decode(existing["thumbnail_b64"]))
                self.event_queue.put(("sp_item_done",))
            else:
                to_download.append((item, thumb_path))

        # Downloading a thumbnail is pure network I/O - overlapping several
        # at once hides Graph's per-request latency behind the CPU-bound
        # CLIP embed of whichever item downloaded first, instead of paying
        # network latency + inference time back-to-back per image. The
        # embed itself stays sequential in this thread: PyTorch's CPU ops
        # already use multiple cores per call, so running several embeds
        # concurrently would mostly just contend with itself.
        reembedded_count = 0
        if to_download:
            downloaded = []  # (item, thumb_path) that downloaded successfully
            with concurrent.futures.ThreadPoolExecutor(max_workers=DOWNLOAD_WORKERS) as pool:
                future_to_item = {
                    pool.submit(self._download_thumbnail_with_retry, item, thumb_path): (item, thumb_path)
                    for item, thumb_path in to_download
                }
                for future in concurrent.futures.as_completed(future_to_item):
                    item, thumb_path = future_to_item[future]
                    try:
                        if not future.result():
                            raise RuntimeError("no thumbnail available")
                        downloaded.append((item, thumb_path))
                    except Exception as exc:
                        self.event_queue.put(("status", f"Skipped {item['name']}: {exc}"))
                        failed_items[item["id"]] = item
                        self.event_queue.put(("sp_item_done",))

            # Batch the actual CLIP embedding - a single forward pass over
            # several images at once is meaningfully faster per-image on
            # CPU than embedding one image at a time.
            for batch_start in range(0, len(downloaded), EMBED_BATCH_SIZE):
                batch = downloaded[batch_start : batch_start + EMBED_BATCH_SIZE]
                self.event_queue.put(("status", f"Indexing {len(batch)} image(s)..."))
                embeddings = self.engine.embed_image_files([tp for _, tp in batch])
                for item, thumb_path in batch:
                    item_id = item["id"]
                    etag = item.get("eTag")
                    emb = embeddings.get(thumb_path)
                    try:
                        if emb is None:
                            raise RuntimeError("could not read downloaded image")
                        with open(thumb_path, "rb") as f:
                            thumb_b64 = base64.b64encode(f.read()).decode("ascii")
                        local_items[item_id] = {
                            "name": item["name"],
                            "etag": etag,
                            "embedding": emb.tolist(),
                            "thumbnail_b64": thumb_b64,
                        }
                        reembedded_count += 1
                        failed_items.pop(item_id, None)
                    except Exception as exc:
                        self.event_queue.put(("status", f"Skipped {item['name']}: {exc}"))
                        failed_items[item_id] = item
                    self.event_queue.put(("sp_item_done",))

        if changed_items:
            _log_reuse(
                f"folder={folder_id} had_local_cache={had_local_cache} "
                f"had_remote_index={bool(remote_items)} remote_error={remote_index_error!r} "
                f"total={len(changed_items)} reused={reused_count} reembedded={reembedded_count}"
            )

        data = {"model": MODEL_TAG, "items": local_items}
        with open(local_index_path, "w", encoding="utf-8") as f:
            json.dump(data, f)
        self._upload_folder_index(folder_id, data, local_index_path, uploaded_ok)

        return local_items

    def _index_sharepoint_tree(self):
        """Delta-syncs every folder under SHAREPOINT_SEARCH_ROOT_PATH and
        makes the combined result searchable.

        Only folders Graph reports as actually changed since the last run
        get touched (first run: everything, since there's no saved delta
        token yet - same cost as a full scan, but still reuses any
        already-indexed folder's cached entries instead of redoing them).
        """
        local_folder, thumbs_dir, _full_dir, indexes_dir = self._sp_cache_dirs()
        os.makedirs(thumbs_dir, exist_ok=True)
        os.makedirs(indexes_dir, exist_ok=True)
        delta_link_path = os.path.join(local_folder, "delta_link.txt")
        folder_map_path = os.path.join(local_folder, "item_folder_map.json")
        pending_retries_path = os.path.join(local_folder, "pending_retries.json")
        uploaded_ok_path = os.path.join(local_folder, "uploaded_ok.json")
        folder_names_path = os.path.join(local_folder, "folder_names.json")

        self._sp_mode_active = True
        self._set_busy(True)
        self.progress.configure(mode="indeterminate")
        self.progress.start()

        def work():
            try:
                delta_link = None
                if os.path.exists(delta_link_path):
                    with open(delta_link_path, "r", encoding="utf-8") as f:
                        delta_link = f.read().strip() or None

                item_folder_map = {}
                if os.path.exists(folder_map_path):
                    try:
                        with open(folder_map_path, "r", encoding="utf-8") as f:
                            item_folder_map = json.load(f)
                    except (json.JSONDecodeError, OSError):
                        item_folder_map = {}

                # Items that failed to download/embed on a previous run (e.g.
                # connectivity dropped mid-run) - delta sync alone would never
                # surface these again since nothing about them looks
                # "changed" to SharePoint, so they're force-retried here.
                pending_retries = {}
                if os.path.exists(pending_retries_path):
                    try:
                        with open(pending_retries_path, "r", encoding="utf-8") as f:
                            pending_retries = json.load(f)
                    except (json.JSONDecodeError, OSError):
                        pending_retries = {}

                # Folder ids whose shared index is confirmed to have reached
                # SharePoint. Anything built locally but missing from here
                # gets re-uploaded by the repair pass below.
                uploaded_ok = set()
                if os.path.exists(uploaded_ok_path):
                    try:
                        with open(uploaded_ok_path, "r", encoding="utf-8") as f:
                            uploaded_ok = set(json.load(f))
                    except (json.JSONDecodeError, OSError, TypeError):
                        uploaded_ok = set()

                self.event_queue.put(("status", "Checking SharePoint for changes..."))
                delta_status = lambda msg: self.event_queue.put(("status", msg))
                root = self.sp_client.get_search_root_item()
                try:
                    raw_items, new_delta_link = self.sp_client.get_delta_items(
                        root["id"], delta_link, status_callback=delta_status
                    )
                except DeltaExpired:
                    self.event_queue.put(("status", "Delta expired - doing a full resync..."))
                    raw_items, new_delta_link = self.sp_client.get_delta_items(
                        root["id"], None, status_callback=delta_status
                    )

                folder_names = {}
                if os.path.exists(folder_names_path):
                    try:
                        with open(folder_names_path, "r", encoding="utf-8") as f:
                            folder_names = json.load(f)
                    except (json.JSONDecodeError, OSError):
                        folder_names = {}

                changed_by_folder = {}
                deleted_ids = []
                for item in raw_items:
                    item_id = item["id"]
                    if item.get("deleted"):
                        deleted_ids.append(item_id)
                        continue
                    if "folder" in item:
                        # Folder names are what make event-based searches
                        # ("Weihnachtsfeier") possible at all - CLIP has no
                        # way to know them from pixels. They arrive free in
                        # the delta listing, so capture rather than discard.
                        if item.get("name"):
                            folder_names[item_id] = item["name"]
                        continue
                    if not self.sp_client.is_wanted_image(item):
                        continue
                    parent_id = item.get("parentReference", {}).get("id")
                    if not parent_id:
                        continue
                    changed_by_folder.setdefault(parent_id, []).append(item)
                    item_folder_map[item_id] = parent_id

                deleted_by_folder = {}
                for item_id in deleted_ids:
                    folder_id = item_folder_map.pop(item_id, None)
                    if folder_id:
                        deleted_by_folder.setdefault(folder_id, []).append(item_id)
                    pending_retries.pop(item_id, None)

                already_queued = {item["id"] for items in changed_by_folder.values() for item in items}
                for item_id, item in pending_retries.items():
                    if item_id in already_queued:
                        continue
                    parent_id = item.get("parentReference", {}).get("id")
                    if not parent_id:
                        continue
                    changed_by_folder.setdefault(parent_id, []).append(item)
                    item_folder_map[item_id] = parent_id

                affected_folders = set(changed_by_folder) | set(deleted_by_folder)
                total_changed = sum(len(v) for v in changed_by_folder.values())
                self.event_queue.put(("sp_scan_done", total_changed))
                self.engine.load_model(lambda msg: self.event_queue.put(("status", msg)))

                failed_items = {}
                for folder_id in affected_folders:
                    self._apply_folder_changes(
                        folder_id,
                        changed_by_folder.get(folder_id, []),
                        deleted_by_folder.get(folder_id, []),
                        thumbs_dir,
                        indexes_dir,
                        failed_items,
                        uploaded_ok,
                    )

                with open(pending_retries_path, "w", encoding="utf-8") as f:
                    json.dump(failed_items, f)

                repaired = self._repair_unuploaded_indexes(indexes_dir, uploaded_ok)
                with open(uploaded_ok_path, "w", encoding="utf-8") as f:
                    json.dump(sorted(uploaded_ok), f)

                # Only name+embedding are needed to build the searchable index below -
                # discard each folder's thumbnail_b64 data (the bulk of its JSON) right
                # after parsing instead of retaining all ~33k images' worth of it in
                # memory at once alongside the CLIP model.
                all_folder_ids = set(item_folder_map.values())
                self._backfill_folder_names(all_folder_ids, folder_names)
                with open(folder_names_path, "w", encoding="utf-8") as f:
                    json.dump(folder_names, f, ensure_ascii=False)

                all_entries = {}
                for folder_id in all_folder_ids:
                    path = os.path.join(indexes_dir, f"{folder_id}.json")
                    if os.path.exists(path):
                        try:
                            with open(path, "r", encoding="utf-8") as f:
                                folder_items = json.load(f).get("items", {})
                            for item_id, entry in folder_items.items():
                                all_entries[item_id] = {
                                    "name": entry["name"],
                                    "embedding": entry["embedding"],
                                    "folder": folder_names.get(folder_id, ""),
                                }
                        except (json.JSONDecodeError, OSError, KeyError):
                            pass

                sp_items = [
                    {
                        "path": os.path.join(thumbs_dir, f"{item_id}.jpg"),
                        "embedding": np.array(entry["embedding"], dtype=np.float32),
                        "meta": {
                            "item_id": item_id,
                            "name": entry["name"],
                            "folder": entry.get("folder", ""),
                        },
                    }
                    for item_id, entry in all_entries.items()
                ]
                count = self.engine.load_sp_items(sp_items)

                _log_reuse(
                    f"RUN COMPLETE: delta_link_was={'set' if delta_link else 'None (full listing)'} "
                    f"new_delta_link={'received' if new_delta_link else 'MISSING'} "
                    f"raw_items={len(raw_items)} affected_folders={len(affected_folders)} "
                    f"item_folder_map_size={len(item_folder_map)} final_searchable_count={count} "
                    f"failed_pending_retry={len(failed_items)} shared_indexes_repaired={repaired} "
                    f"shared_indexes_confirmed={len(uploaded_ok)}"
                )
                if new_delta_link:
                    with open(delta_link_path, "w", encoding="utf-8") as f:
                        f.write(new_delta_link)
                with open(folder_map_path, "w", encoding="utf-8") as f:
                    json.dump(item_folder_map, f)

                self.event_queue.put(("sp_index_done", count, len(affected_folders), len(failed_items)))
            except Exception as exc:
                _log_crash("SharePoint indexing", sys.exc_info())
                self.event_queue.put(("error", str(exc)))

        self._run_in_background(work)

    def _open_sp_result(self, item_id, name):
        self._record_click(item_id)
        _local_folder, _thumbs_dir, full_dir, _indexes_dir = self._sp_cache_dirs()
        dest = os.path.join(full_dir, f"{item_id}_{name}")
        if os.path.exists(dest):
            open_in_system_viewer(dest)
            return

        self._set_busy(True)
        self.status_var.set(f"Downloading {name}...")

        def work():
            try:
                item = {"id": item_id}
                self.sp_client.download_file(item, dest)
                self.event_queue.put(("sp_open_ready", dest))
            except Exception as exc:
                self.event_queue.put(("error", f"Failed to download {name}: {exc}"))

        self._run_in_background(work)

    # ---------- event loop / queue polling ----------

    def _poll_queue(self):
        try:
            while True:
                event = self.event_queue.get_nowait()
                kind = event[0]
                if kind == "progress":
                    _, done, total = event
                    self.progress.stop()
                    self.progress.configure(mode="determinate")
                    self.progress.set(done / max(total, 1))
                    self.status_var.set(f"Indexing... {done}/{total}")
                elif kind == "status":
                    self.status_var.set(event[1])
                elif kind == "index_done":
                    self.progress.stop()
                    self.progress.set(0)
                    self._set_busy(False)
                    self.status_var.set(f"Indexed {event[1]} images. Ready to search.")
                elif kind == "search_done":
                    self.progress.stop()
                    self.progress.set(0)
                    self._set_busy(False)
                    self._render_results(event[1])
                    self.status_var.set(f"Found {len(event[1])} results.")
                elif kind == "sp_signed_in":
                    self._set_busy(False)
                    self._index_sharepoint_tree()
                elif kind == "sp_scan_done":
                    self._sp_total_images = event[1]
                    self._sp_done_images = 0
                    self.progress.stop()
                    self.progress.configure(mode="determinate")
                    self.progress.set(0)
                elif kind == "sp_item_done":
                    self._sp_done_images += 1
                    self.progress.set(self._sp_done_images / max(self._sp_total_images, 1))
                    self.status_var.set(
                        f"Indexing from SharePoint... {self._sp_done_images}/{self._sp_total_images}"
                    )
                elif kind == "sp_index_done":
                    _, count, changed_folder_count, failed_count = event
                    self.progress.stop()
                    self.progress.set(0)
                    self.folder_var.set("SharePoint: Fotos & Videos (all subfolders)")
                    self._set_busy(False)
                    retry_note = (
                        f" {failed_count} image(s) failed (e.g. lost connection) and will "
                        "automatically retry next run."
                        if failed_count
                        else ""
                    )
                    if changed_folder_count:
                        self.status_var.set(
                            f"Indexed {count} image(s) total ({changed_folder_count} folder(s) "
                            f"had changes). Ready to search.{retry_note}"
                        )
                    else:
                        self.status_var.set(
                            f"Up to date - {count} image(s) already indexed. "
                            f"Ready to search.{retry_note}"
                        )
                elif kind == "sp_open_ready":
                    self._set_busy(False)
                    self.status_var.set("Ready to search.")
                    open_in_system_viewer(event[1])
                elif kind == "error":
                    self.progress.stop()
                    self.progress.set(0)
                    self._set_busy(False)
                    self.status_var.set("Error.")
                    messagebox.showerror("Image Search", event[1])
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    # ---------- results rendering ----------

    def _render_results(self, results):
        for child in self.results_frame.winfo_children():
            child.destroy()
        self.thumbnail_refs.clear()

        for idx, (path, score, meta) in enumerate(results):
            row, col = divmod(idx, GRID_COLUMNS)
            cell = ctk.CTkFrame(self.results_frame, fg_color="transparent")
            cell.grid(row=row, column=col, sticky="n", padx=4, pady=4)

            try:
                img = Image.open(path)
                img.thumbnail(THUMB_SIZE)
                photo = ctk.CTkImage(light_image=img, dark_image=img, size=img.size)
            except Exception:
                continue
            self.thumbnail_refs.append(photo)

            label = ctk.CTkLabel(cell, image=photo, text="", cursor="hand2")
            label.pack()
            if meta is None:
                label.bind("<Double-Button-1>", lambda _e, p=path: open_in_system_viewer(p))
            else:
                label.bind(
                    "<Double-Button-1>",
                    lambda _e, m=meta: self._open_sp_result(m["item_id"], m["name"]),
                )

            # The SharePoint filename is almost always a camera default
            # (DSC05540.jpg), so the folder is the only human-meaningful
            # label a result can carry - show it when we know it.
            folder = (meta or {}).get("folder") if meta else None
            caption = folder if folder else os.path.basename(path)
            ctk.CTkLabel(
                cell, text=caption, justify=tk.CENTER, wraplength=THUMB_SIZE[0],
                text_color=MUTED_TEXT_COLOR, font=ctk.CTkFont(size=11),
            ).pack()


def _thread_excepthook(args):
    _log_crash("unhandled thread exception", (args.exc_type, args.exc_value, args.exc_traceback))


def main():
    threading.excepthook = _thread_excepthook
    ctk.set_appearance_mode("dark")
    root = ctk.CTk()
    try:
        root.configure(fg_color=BG_COLOR)
    except Exception:
        pass
    try:
        root.iconbitmap(_resource_path("assets", "app_icon.ico"))
    except Exception:
        pass
    ImageSearchApp(root)
    try:
        root.mainloop()
    except Exception:
        _log_crash("mainloop", sys.exc_info())
        raise


if __name__ == "__main__":
    main()
