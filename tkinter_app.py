"""Image Search desktop app.

Point it at a folder of photos, index them once, then search by typing a
description ("a dog on a beach") or by picking a reference image.
Indexing uses a local CLIP model, so search works fully offline after the
model has been downloaded once.
"""

import base64
import concurrent.futures
import datetime
import json
import os
import queue
import subprocess
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# When running as a bundled .exe (PyInstaller --onedir), redirect the CLIP
# model's cache to the copy bundled alongside the exe (see build.bat /
# download_model.py) so end users never wait on the ~350-600MB first-run
# download. Must happen before anything imports huggingface_hub, which reads
# HF_HOME once at import time - hence this sits above every other import.
if getattr(sys, "frozen", False):
    os.environ["HF_HOME"] = os.path.join(sys._MEIPASS, "model_cache")

import numpy as np
from PIL import Image, ImageTk

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
DEFAULT_TOP_K = 50

BG_COLOR = "#0d1b2a"
CARD_COLOR = "#0a141f"
ACCENT_COLOR = "#0b4c8c"
ACCENT_ACTIVE = "#004fff"
ACCENT_DISABLED = "#33404c"
FIELD_BG = "#0a141f"
TEXT_COLOR = "#eef3fa"
MUTED_TEXT_COLOR = "#8fa3b8"
DISABLED_TEXT_COLOR = "#5c6b7a"

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
        img = img.resize((max(1, round(img.width * ratio)), height), Image.LANCZOS)
        return ImageTk.PhotoImage(img)
    except Exception:
        return None


class RoundedButton(tk.Canvas):
    """A pill-shaped button - ttk has no border-radius support, so this
    draws its own rounded rect on a Canvas instead."""

    def __init__(self, parent, text, command, width=140, height=34, radius=17, font=None,
                 canvas_bg=CARD_COLOR):
        super().__init__(
            parent, width=width, height=height, background=canvas_bg,
            highlightthickness=0, cursor="hand2",
        )
        self._command = command
        self._width = width
        self._height = height
        self._radius = radius
        self._text = text
        self._font = font or ("Segoe UI", 10, "bold")
        self._enabled = True
        self._draw(ACCENT_COLOR, TEXT_COLOR)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", self._on_click)

    @staticmethod
    def _round_rect_points(x1, y1, x2, y2, r):
        return [
            x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
            x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
            x1, y2, x1, y2 - r, x1, y1 + r, x1, y1,
        ]

    def _draw(self, fill, textfill):
        self.delete("all")
        points = self._round_rect_points(1, 1, self._width - 1, self._height - 1, self._radius)
        self.create_polygon(points, smooth=True, fill=fill, outline=fill)
        self.create_text(
            self._width / 2, self._height / 2, text=self._text, fill=textfill, font=self._font
        )

    def _on_enter(self, _e):
        if self._enabled:
            self._draw(ACCENT_ACTIVE, TEXT_COLOR)

    def _on_leave(self, _e):
        if self._enabled:
            self._draw(ACCENT_COLOR, TEXT_COLOR)

    def _on_click(self, _e):
        if self._enabled and self._command:
            self._command()

    def config_state(self, enabled):
        self._enabled = enabled
        if enabled:
            self.configure(cursor="hand2")
            self._draw(ACCENT_COLOR, TEXT_COLOR)
        else:
            self.configure(cursor="arrow")
            self._draw(ACCENT_DISABLED, DISABLED_TEXT_COLOR)


def open_in_system_viewer(path):
    if sys.platform.startswith("win"):
        os.startfile(path)  # noqa: S606 (Windows-only convenience)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


class ImageSearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Image Search")
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
        self.thumbnail_refs = []  # keep PhotoImage refs alive

        self._build_widgets()
        self.root.after(100, self._poll_queue)

    # ---------- UI construction ----------

    def _build_widgets(self):
        header = ttk.Frame(self.root, padding=(8, 8, 8, 0))
        header.pack(fill=tk.X)
        self._logo_photo = _load_logo_image()
        if self._logo_photo:
            ttk.Label(header, image=self._logo_photo).pack(side=tk.LEFT)

        top = ttk.Frame(self.root, style="Card.TFrame", padding=8)
        top.pack(fill=tk.X, padx=8, pady=8)

        self.folder_var = tk.StringVar(value="No folder selected")
        RoundedButton(top, "Choose Folder...", self._choose_folder, width=150).pack(
            side=tk.LEFT
        )
        ttk.Label(top, textvariable=self.folder_var, width=50, style="Card.TLabel").pack(
            side=tk.LEFT, padx=8
        )
        self.index_button = RoundedButton(top, "Index Folder", self._start_indexing, width=130)
        self.index_button.pack(side=tk.LEFT, padx=4)
        self.index_button.config_state(False)
        self.sharepoint_button = RoundedButton(
            top, "Search Marketing Photos", self._browse_sharepoint, width=190
        )
        self.sharepoint_button.pack(side=tk.LEFT, padx=4)

        search_frame = ttk.Frame(self.root, style="Card.TFrame", padding=8)
        search_frame.pack(fill=tk.X, padx=8, pady=(0, 8))

        self.query_var = tk.StringVar()
        query_entry = ttk.Entry(search_frame, textvariable=self.query_var)
        query_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        query_entry.bind("<Return>", lambda _e: self._start_text_search())

        self.search_button = RoundedButton(
            search_frame, "Search", self._start_text_search, width=100
        )
        self.search_button.pack(side=tk.LEFT, padx=(8, 0))
        self.search_button.config_state(False)

        status_frame = ttk.Frame(self.root, padding=(8, 0, 8, 4))
        status_frame.pack(fill=tk.X)
        self.status_var = tk.StringVar(value="Choose a folder to begin.")
        ttk.Label(status_frame, textvariable=self.status_var).pack(side=tk.LEFT)
        self.progress = ttk.Progressbar(status_frame, mode="determinate", length=200)
        self.progress.pack(side=tk.RIGHT)

        # Scrollable results grid
        results_container = ttk.Frame(self.root)
        results_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        canvas = tk.Canvas(results_container, highlightthickness=0, background=BG_COLOR)
        scrollbar = ttk.Scrollbar(results_container, orient=tk.VERTICAL, command=canvas.yview)
        self.results_frame = ttk.Frame(canvas)
        self.results_frame.bind(
            "<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.results_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)

    # ---------- folder selection & indexing ----------

    def _choose_folder(self):
        folder = filedialog.askdirectory(title="Choose a folder of images")
        if not folder:
            return
        self._sp_mode_active = False
        self.folder_var.set(folder)
        self.index_button.config_state(True)

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
        self.progress.config(mode="indeterminate")
        self.progress.start(10)

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
        self._set_busy(True)
        self.progress.config(mode="indeterminate")
        self.progress.start(10)

        def work():
            def on_status(msg):
                self.event_queue.put(("status", msg))

            try:
                results = self.engine.search_text(query, DEFAULT_TOP_K, status_callback=on_status)
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
        self.index_button.config_state(not busy and not self._sp_mode_active)
        self.sharepoint_button.config_state(not busy)
        if busy:
            self.search_button.config_state(False)
        else:
            self._update_search_buttons_state()

    def _update_search_buttons_state(self):
        has_index = self.engine.embeddings is not None and len(self.engine.paths) > 0
        self.search_button.config_state(has_index)

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

    def _apply_folder_changes(self, folder_id, changed_items, deleted_ids, thumbs_dir, indexes_dir):
        """Applies a delta-reported set of adds/updates/deletes to one
        SharePoint folder's shared, thumbnail-based index.

        Reuses a shared `.imagesearch_sp_index.json` uploaded into the
        SharePoint folder itself when possible (keyed by driveItem id +
        eTag), so only genuinely new/changed images ever need their
        thumbnail downloaded and embedded - by anyone, on any machine.
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
            with concurrent.futures.ThreadPoolExecutor(max_workers=DOWNLOAD_WORKERS) as pool:
                future_to_item = {
                    pool.submit(self.sp_client.download_thumbnail, item, thumb_path): (item, thumb_path)
                    for item, thumb_path in to_download
                }
                for future in concurrent.futures.as_completed(future_to_item):
                    item, thumb_path = future_to_item[future]
                    item_id = item["id"]
                    etag = item.get("eTag")
                    try:
                        if not future.result():
                            raise RuntimeError("no thumbnail available")
                        self.event_queue.put(("status", f"Indexing {item['name']}..."))
                        emb = self.engine.embed_image_file(thumb_path)
                        with open(thumb_path, "rb") as f:
                            thumb_b64 = base64.b64encode(f.read()).decode("ascii")
                        local_items[item_id] = {
                            "name": item["name"],
                            "etag": etag,
                            "embedding": emb.tolist(),
                            "thumbnail_b64": thumb_b64,
                        }
                        reembedded_count += 1
                    except Exception as exc:
                        self.event_queue.put(("status", f"Skipped {item['name']}: {exc}"))
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
        try:
            self.sp_client.upload_index_file(folder_id, data)
        except Exception as exc:
            self.event_queue.put(("status", f"Warning: couldn't upload shared index: {exc}"))

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

        self._sp_mode_active = True
        self._set_busy(True)
        self.progress.config(mode="indeterminate")
        self.progress.start(10)

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

                self.event_queue.put(("status", "Checking SharePoint for changes..."))
                root = self.sp_client.get_search_root_item()
                try:
                    raw_items, new_delta_link = self.sp_client.get_delta_items(
                        root["id"], delta_link
                    )
                except DeltaExpired:
                    self.event_queue.put(("status", "Delta expired - doing a full resync..."))
                    raw_items, new_delta_link = self.sp_client.get_delta_items(root["id"], None)

                changed_by_folder = {}
                deleted_ids = []
                for item in raw_items:
                    item_id = item["id"]
                    if item.get("deleted"):
                        deleted_ids.append(item_id)
                        continue
                    if "folder" in item:
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

                affected_folders = set(changed_by_folder) | set(deleted_by_folder)
                total_changed = sum(len(v) for v in changed_by_folder.values())
                self.event_queue.put(("sp_scan_done", total_changed))
                self.engine.load_model(lambda msg: self.event_queue.put(("status", msg)))

                for folder_id in affected_folders:
                    self._apply_folder_changes(
                        folder_id,
                        changed_by_folder.get(folder_id, []),
                        deleted_by_folder.get(folder_id, []),
                        thumbs_dir,
                        indexes_dir,
                    )

                # Only name+embedding are needed to build the searchable index below -
                # discard each folder's thumbnail_b64 data (the bulk of its JSON) right
                # after parsing instead of retaining all ~33k images' worth of it in
                # memory at once alongside the CLIP model.
                all_entries = {}
                for folder_id in set(item_folder_map.values()):
                    path = os.path.join(indexes_dir, f"{folder_id}.json")
                    if os.path.exists(path):
                        try:
                            with open(path, "r", encoding="utf-8") as f:
                                folder_items = json.load(f).get("items", {})
                            for item_id, entry in folder_items.items():
                                all_entries[item_id] = {
                                    "name": entry["name"],
                                    "embedding": entry["embedding"],
                                }
                        except (json.JSONDecodeError, OSError, KeyError):
                            pass

                sp_items = [
                    {
                        "path": os.path.join(thumbs_dir, f"{item_id}.jpg"),
                        "embedding": np.array(entry["embedding"], dtype=np.float32),
                        "meta": {"item_id": item_id, "name": entry["name"]},
                    }
                    for item_id, entry in all_entries.items()
                ]
                count = self.engine.load_sp_items(sp_items)

                _log_reuse(
                    f"RUN COMPLETE: delta_link_was={'set' if delta_link else 'None (full listing)'} "
                    f"new_delta_link={'received' if new_delta_link else 'MISSING'} "
                    f"raw_items={len(raw_items)} affected_folders={len(affected_folders)} "
                    f"item_folder_map_size={len(item_folder_map)} final_searchable_count={count}"
                )
                if new_delta_link:
                    with open(delta_link_path, "w", encoding="utf-8") as f:
                        f.write(new_delta_link)
                with open(folder_map_path, "w", encoding="utf-8") as f:
                    json.dump(item_folder_map, f)

                self.event_queue.put(("sp_index_done", count, len(affected_folders)))
            except Exception as exc:
                _log_crash("SharePoint indexing", sys.exc_info())
                self.event_queue.put(("error", str(exc)))

        self._run_in_background(work)

    def _open_sp_result(self, item_id, name):
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
                    self.progress.config(mode="determinate", maximum=max(total, 1), value=done)
                    self.status_var.set(f"Indexing... {done}/{total}")
                elif kind == "status":
                    self.status_var.set(event[1])
                elif kind == "index_done":
                    self.progress.stop()
                    self.progress.config(value=0)
                    self._set_busy(False)
                    self.status_var.set(f"Indexed {event[1]} images. Ready to search.")
                elif kind == "search_done":
                    self.progress.stop()
                    self.progress.config(value=0)
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
                    self.progress.config(mode="determinate", maximum=max(event[1], 1), value=0)
                elif kind == "sp_item_done":
                    self._sp_done_images += 1
                    self.progress.config(value=self._sp_done_images)
                    self.status_var.set(
                        f"Indexing from SharePoint... {self._sp_done_images}/{self._sp_total_images}"
                    )
                elif kind == "sp_index_done":
                    _, count, changed_folder_count = event
                    self.progress.stop()
                    self.progress.config(value=0)
                    self.folder_var.set("SharePoint: Fotos & Videos (all subfolders)")
                    self._set_busy(False)
                    if changed_folder_count:
                        self.status_var.set(
                            f"Indexed {count} image(s) total ({changed_folder_count} folder(s) "
                            "had changes). Ready to search."
                        )
                    else:
                        self.status_var.set(
                            f"Up to date - {count} image(s) already indexed. Ready to search."
                        )
                elif kind == "sp_open_ready":
                    self._set_busy(False)
                    self.status_var.set("Ready to search.")
                    open_in_system_viewer(event[1])
                elif kind == "error":
                    self.progress.stop()
                    self.progress.config(value=0)
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
            cell = ttk.Frame(self.results_frame, padding=4)
            cell.grid(row=row, column=col, sticky="n")

            try:
                img = Image.open(path)
                img.thumbnail(THUMB_SIZE)
                photo = ImageTk.PhotoImage(img)
            except Exception:
                continue
            self.thumbnail_refs.append(photo)

            label = ttk.Label(cell, image=photo, cursor="hand2")
            label.pack()
            if meta is None:
                label.bind("<Double-Button-1>", lambda _e, p=path: open_in_system_viewer(p))
            else:
                label.bind(
                    "<Double-Button-1>",
                    lambda _e, m=meta: self._open_sp_result(m["item_id"], m["name"]),
                )

            caption = f"{os.path.basename(path)}\n{score * 100:.1f}%"
            ttk.Label(cell, text=caption, justify=tk.CENTER, wraplength=THUMB_SIZE[0]).pack()


def _apply_theme(root):
    root.configure(background=BG_COLOR)
    style = ttk.Style()
    style.theme_use("clam")

    style.configure("TFrame", background=BG_COLOR)
    style.configure("TLabel", background=BG_COLOR, foreground=TEXT_COLOR)
    style.configure("Card.TFrame", background=CARD_COLOR)
    style.configure("Card.TLabel", background=CARD_COLOR, foreground=MUTED_TEXT_COLOR)
    style.configure(
        "TEntry",
        fieldbackground=FIELD_BG,
        foreground=TEXT_COLOR,
        insertcolor=TEXT_COLOR,
        borderwidth=0,
    )
    style.configure(
        "Horizontal.TProgressbar",
        background=ACCENT_COLOR,
        troughcolor=FIELD_BG,
        borderwidth=0,
    )
    style.configure(
        "Vertical.TScrollbar",
        background=ACCENT_COLOR,
        troughcolor=BG_COLOR,
        arrowcolor=TEXT_COLOR,
        borderwidth=0,
    )
    style.map("Vertical.TScrollbar", background=[("active", ACCENT_ACTIVE)])


def _thread_excepthook(args):
    _log_crash("unhandled thread exception", (args.exc_type, args.exc_value, args.exc_traceback))


def main():
    threading.excepthook = _thread_excepthook
    root = tk.Tk()
    try:
        _apply_theme(root)
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
