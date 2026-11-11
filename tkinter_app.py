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
import hashlib
import json
import os
import platform
import queue
import re
import shutil
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
elif "HF_HOME" not in os.environ:
    # Running from source, the same pre-downloaded copy sits next to this
    # file. Without this the weights are looked for in the default
    # HuggingFace cache, which on a machine that has never downloaded them
    # means either a 600MB wait or - where the network blocks it - the
    # "Failed to download weights" crash recorded in crash.log.
    _SOURCE_MODEL_CACHE = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "model_cache"
    )
    if os.path.isdir(_SOURCE_MODEL_CACHE):
        os.environ["HF_HOME"] = _SOURCE_MODEL_CACHE

import customtkinter as ctk
import numpy as np
from PIL import Image

from search_engine import ImageSearchEngine, MODEL_TAG
from sharepoint_client import (
    FAVOURITES_PREFIX,
    FEEDBACK_PREFIX,
    DeltaExpired,
    SharePointClient,
)

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


# Matches the long edge of the thumbnails SharePoint gives us (Graph
# "medium" = 176px): displaying smaller than the source threw away detail
# we already had on disk, and displaying larger would only upscale blur.
THUMB_SIZE = (176, 176)
# Decoded thumbnails kept for reuse across re-layouts; a few hundred
# 176px images is a handful of MB, far cheaper than re-decoding them.
THUMB_CACHE_MAX = 800
GRID_COLUMNS = 5  # fallback used only before the results area has a real size
GRID_CELL_PADDING = 24  # horizontal space a grid cell needs beyond the thumbnail itself
RESIZE_DEBOUNCE_MS = 150
# CTkScrollableFrame moves 1px per scroll "unit" on Windows and ~20 units per
# wheel notch; multiplying the increment scales the distance per notch
# without touching the wheel event handling itself.
SCROLL_SPEED_MULTIPLIER = 3
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

# Collections are curated highlight reels, not catch-alls: only images
# scoring in the top slice for a collection are admitted, and each is capped
# so browsing stays quick. Photos containing guests are ordered ahead of
# product and object shots within every collection.
COLLECTION_PERCENTILE = 92
COLLECTION_MAX = 150
# How much more "graphic with text on it" than "plain photograph" an image
# has to look before it's kept out of collections. Slightly above zero so
# only clear cases are excluded - a photo that merely contains a sign in
# the background shouldn't be treated as finished artwork.
TEXT_OVERLAY_MARGIN = -0.012
HISTORY_MAX = 100
# How many photos the opening view shows. Enough to fill a maximised
# window several rows deep without making startup wait on decoding.
ALL_PHOTOS_PREVIEW = 120
# A collection with one or two photos in it reads as a bug rather than
# a category, so small ones are kept off the sidebar.
MIN_COLLECTION_SIZE = 10
# How many results keep strict relevance order before the variety rules are
# allowed to reorder anything. The engine ranks the whole library correctly;
# interleaving by folder and grouping by orientation then moved genuinely
# better matches down the grid to make room for variety. Keeping the opening
# rows in pure relevance order means a query with a qualifier in it
# ("nighttime", "no people") is answered by the top of the grid, and the
# variety rules still shape the long tail underneath.
RELEVANCE_HEAD = 24
# The collections in the sidebar. They start empty and are filled by hand -
# dragging a photo onto one, or right-clicking it - rather than being
# guessed from the image. An automatic grouping was close but never right,
# and a shortlist someone curated is worth more than one nobody trusts.
MANUAL_COLLECTIONS = ["Golf", "Food", "Drinks", "Family", "Kids"]
# How far the pointer has to move with the button held before it counts as a
# drag rather than a click, so double-clicking a photo still opens it.
DRAG_THRESHOLD_PX = 8
# Largest the in-app preview will draw an image, in px.
PREVIEW_MAX = (1120, 740)
# Results built per pass before handing control back to the event loop.
RENDER_CHUNK = 24

# The two directions used both to score images ("does this photo have guests
# in it") and to read a query ("is this asking for a scene or for a thing").
PEOPLE_PROMPTS = [
    "people enjoying themselves",
    "a group of happy people",
    "guests having fun",
]
EMPTY_PROMPTS = [
    "an empty room with no people",
    "an empty venue interior",
    "a close up of an object",
]
OBJECT_PROMPTS = [
    "a close up photograph of food on a plate",
    "a drink in a glass on a bar",
    "a product shot of a single object",
    "a detail photograph of equipment",
]
DEFAULT_TOP_K = 150

BG_COLOR = "#0d1b2a"
CARD_COLOR = "#0a141f"
ACCENT_COLOR = "#0b4c8c"
ACCENT_ACTIVE = "#004fff"
FIELD_BG = "#0a141f"
TEXT_COLOR = "#eef3fa"
MUTED_TEXT_COLOR = "#8fa3b8"

CARD_RADIUS = 14
BUTTON_RADIUS = 17

LOGO_HEIGHT = 40


def _build_stamp():
    """A short "which version is this" marker for the header.

    Running from source and running the built .exe look identical on
    screen, and the two can be weeks apart - so the window says when the
    code it is running was last changed, and nobody has to guess.
    """
    try:
        target = sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__)
        stamp = datetime.datetime.fromtimestamp(os.path.getmtime(target))
        return stamp.strftime("%d %b %H:%M")
    except Exception:
        return ""


def _resource_path(*parts):
    # PyInstaller (--onedir) unpacks bundled data next to the exe under
    # sys._MEIPASS; running from source, it's just this file's folder.
    base = sys._MEIPASS if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, *parts)


def _load_logo_image(height=LOGO_HEIGHT):
    try:
        img = Image.open(_resource_path("assets", "logo.png")).convert("RGBA")
        # The file carries a wide transparent margin - the artwork is only
        # 1048x764 inside a 1440x1282 canvas - so scaling the whole canvas
        # left the mark itself around 24px tall and looking soft. Cropping
        # to the artwork first lets it use the full height asked for.
        box = img.getbbox()
        if box:
            img = img.crop(box)
        ratio = height / img.height
        width = max(1, round(img.width * ratio))
        # Resampled here rather than left to the widget. The source mark is
        # 1440px tall, and a plain scale to 40px aliases the fine strokes -
        # LANCZOS keeps the edges clean at this size.
        img = img.resize((width, height), Image.LANCZOS)
        return ctk.CTkImage(light_image=img, dark_image=img, size=(width, height))
    except Exception:
        return None


def query_key(text):
    """Stable, non-readable key for a search query.

    Click feedback is shared with the team through SharePoint, and the raw
    query text would otherwise sit there as a readable log of what everyone
    searched for. Hashing keeps matching exact (the same query always yields
    the same key on every machine) while leaving nothing legible in the file.

    Obfuscation, not encryption: this app is open source, so short common
    words could be recovered by hashing a dictionary and comparing. It stops
    the file being readable, not a determined attacker.
    """
    normalized = normalize_text(text)
    if not normalized:
        return ""
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()[:16]


def machine_key():
    """Opaque per-machine id used to name this machine's feedback file, so
    the filenames don't advertise who searched for what either."""
    raw = f"{platform.node()}|{os.environ.get('USERNAME', '')}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:12]


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
        self.root.title("Greenreb Image Search")
        self.root.geometry("1180x760")
        self.root.minsize(1000, 560)
        # Opens filling the screen: more of the grid is visible at once, and
        # the 1180x760 default looked small on a desk of large monitors.
        # The geometry above stays as the size it restores to.
        try:
            self.root.state("zoomed")
        except Exception:
            pass

        self.engine = ImageSearchEngine()
        self.sp_client = None
        self._sp_mode_active = False
        self._sp_total_images = 0
        self._sp_done_images = 0
        self.event_queue = queue.Queue()
        self.worker_thread = None
        self.cancel_requested = False
        self.thumbnail_refs = []  # keep CTkImage refs alive
        self._thumb_cache = {}  # path -> CTkImage, see _thumb_for
        self._last_query = ""
        self._feedback = None
        self._favourites = None
        self._settings = None
        self._history = None
        self._landscape_cache = {}
        self._people_vec = None
        self._empty_vec = None
        self._graphic_vec = None
        self._object_vec = None
        self._plain_vec = None
        self._current_columns = GRID_COLUMNS
        self._resize_after_id = None
        self._render_after_id = None
        self._drag = None
        self._preview = None
        self._view = "all"
        # Mirrors the state _set_busy puts the buttons in, so paths that
        # do not go through a button (the Enter key) can honour it too.
        self._busy = False
        self._collections = {}
        self.all_card = None
        self.fav_card = None
        self.collection_cards = {}

        self._build_widgets()
        self.root.after(100, self._poll_queue)
        self._startup_load()

    # ---------- UI construction ----------

    def _build_widgets(self):
        header = ctk.CTkFrame(self.root, fg_color="transparent")
        header.pack(fill=tk.X, padx=12, pady=(12, 0))
        self._logo_image = _load_logo_image()
        if self._logo_image:
            ctk.CTkLabel(header, image=self._logo_image, text="").pack(side=tk.LEFT)
        # The header was a small logo adrift in an otherwise empty band; naming
        # the app there makes the window read as a product rather than a tool.
        ctk.CTkLabel(
            header, text="Image Search",
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
            text_color=TEXT_COLOR,
        ).pack(side=tk.LEFT, padx=(12, 0))
        stamp = _build_stamp()
        subtitle = "Marketing photo library"
        if stamp:
            subtitle = f"{subtitle}   \u00b7   version {stamp}"
        ctk.CTkLabel(
            header, text=subtitle,
            font=ctk.CTkFont(family="Segoe UI", size=13),
            text_color=MUTED_TEXT_COLOR,
        ).pack(side=tk.LEFT, padx=(12, 0), pady=(4, 0))

        top = ctk.CTkFrame(self.root, fg_color=CARD_COLOR, corner_radius=CARD_RADIUS)
        top.pack(fill=tk.X, padx=12, pady=12)

        self.folder_var = tk.StringVar(value="No folder selected")
        _make_button(top, "Choose Folder...", self._choose_folder, width=150).pack(
            side=tk.LEFT, padx=(14, 8), pady=14
        )
        # Sized so the buttons to its right still fit at the minimum window
        # width rather than being clipped off the edge.
        ctk.CTkLabel(
            top, textvariable=self.folder_var, width=150, anchor="w",
            text_color=MUTED_TEXT_COLOR,
        ).pack(side=tk.LEFT, padx=8, pady=14, fill=tk.X, expand=True)
        self.index_button = _make_button(top, "Index Folder", self._start_indexing, width=130)
        self.index_button.pack(side=tk.LEFT, padx=4, pady=14)
        self.index_button.configure(state=tk.DISABLED)
        self.sharepoint_button = _make_button(
            top, "Check for new photos", self._browse_sharepoint, width=190
        )
        self.sharepoint_button.pack(side=tk.LEFT, padx=4, pady=14)
        _make_button(top, "Settings", self._open_settings, width=100).pack(
            side=tk.LEFT, padx=(4, 14), pady=14
        )

        # Sidebar (collection cards) beside the search area and results, so
        # switching collection keeps the same search box rather than moving it.
        body = ctk.CTkFrame(self.root, fg_color="transparent")
        body.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 12))

        sidebar = ctk.CTkFrame(body, fg_color="transparent", width=176)
        sidebar.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 12))
        sidebar.pack_propagate(False)

        self.sidebar = sidebar
        self.all_card = self._make_collection_card(sidebar, "All photos", lambda: self._set_view("all"))
        self.fav_card = self._make_collection_card(
            sidebar, "Favourites", lambda: self._set_view("favourites")
        )
        self.history_card = self._make_collection_card(
            sidebar, "Search history", lambda: self._set_view("history")
        )
        # Separates the three ways of looking at the whole library from the
        # automatic collections below.
        ctk.CTkFrame(sidebar, fg_color=ACCENT_ACTIVE, height=3, corner_radius=2).pack(
            fill=tk.X, pady=(6, 10)
        )
        self.collection_cards = {}

        main = ctk.CTkFrame(body, fg_color="transparent")
        main.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        search_frame = ctk.CTkFrame(main, fg_color=CARD_COLOR, corner_radius=CARD_RADIUS)
        search_frame.pack(fill=tk.X, pady=(0, 12))

        self.query_var = tk.StringVar()
        self.query_entry = ctk.CTkEntry(
            search_frame, textvariable=self.query_var, fg_color=FIELD_BG,
            text_color=TEXT_COLOR, border_width=0, corner_radius=10,
            placeholder_text="Describe an image in English...",
        )
        self.query_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(14, 8), pady=14)
        self.query_entry.bind("<Return>", lambda _e: self._start_text_search())

        self.image_search_button = _make_button(
            search_frame, "Find similar...", self._start_image_search, width=130
        )
        self.image_search_button.pack(side=tk.LEFT, padx=(0, 8), pady=14)
        self.image_search_button.configure(state=tk.DISABLED)

        self.search_button = _make_button(
            search_frame, "Search", self._start_text_search, width=100
        )
        self.search_button.pack(side=tk.LEFT, padx=(0, 14), pady=14)
        self.search_button.configure(state=tk.DISABLED)

        self.suggestion_bar = ctk.CTkFrame(main, fg_color="transparent")
        self.suggestion_bar.pack(fill=tk.X, pady=(0, 8))

        status_frame = ctk.CTkFrame(main, fg_color="transparent")
        status_frame.pack(fill=tk.X, pady=(0, 8))
        self.status_var = tk.StringVar(
            value="Loading your photo library..."
        )
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
        self.results_frame = ctk.CTkScrollableFrame(main, fg_color=BG_COLOR)
        self.results_frame.pack(fill=tk.BOTH, expand=True)
        # CustomTkinter moves 1 pixel per scroll "unit" on Windows, which
        # feels sluggish over a tall grid of 150 results - scaling the
        # increment (rather than rebinding the wheel handler, which would
        # stack a second handler on top of CTk's own) makes each notch move
        # further without changing how scrolling behaves in any other way.
        self.results_frame._parent_canvas.configure(
            yscrollincrement=SCROLL_SPEED_MULTIPLIER
        )
        self._last_results = []
        # add="+" is essential here: CTkScrollableFrame already binds its own
        # <Configure> handler on this same widget to keep the canvas's
        # scrollregion in sync with the grid's actual height. A plain bind()
        # call replaces that handler instead of stacking on top of it, which
        # freezes the scrollregion at whatever it was on first render and
        # makes it impossible to scroll past that point - only the first
        # screenful of results was ever reachable.
        self.results_frame.bind(
            "<Configure>", lambda _e: self._on_results_area_resized(), add="+"
        )

        self._update_favourites_button()
        self._set_view("all")

    def _make_collection_card(self, parent, text, command):
        card = ctk.CTkButton(
            parent,
            text=text,
            command=command,
            height=48,
            corner_radius=CARD_RADIUS,
            fg_color=CARD_COLOR,
            hover_color=ACCENT_COLOR,
            text_color=TEXT_COLOR,
            anchor="w",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
        )
        card.pack(fill=tk.X, pady=(0, 8))
        return card

    def _update_favourites_button(self):
        if getattr(self, "fav_card", None) is None:
            return
        self.fav_card.configure(text=f"  Favourites ({len(self._load_favourites())})")

    def _update_history_button(self):
        if getattr(self, "history_card", None) is None:
            return
        self.history_card.configure(text=f"  Search history ({len(self._load_history())})")

    def _collections_cache_path(self):
        local_folder, _t, _f, _i = self._sp_cache_dirs()
        return os.path.join(local_folder, "collections.json")

    def _save_collections(self, collections):
        """Stores collections by image id rather than position, so they
        survive the index being rebuilt in a different order."""
        try:
            by_id = {
                name: [self.engine.meta[i].get("item_id") for i in idxs if self.engine.meta[i]]
                for name, idxs in collections.items()
            }
            with open(self._collections_cache_path(), "w", encoding="utf-8") as f:
                json.dump(by_id, f, ensure_ascii=False)
        except (OSError, IndexError) as exc:
            _log_reuse(f"COLLECTIONS SAVE FAILED: {exc!r}")

    def _load_collections(self):
        """Restores cached collections without needing the CLIP model.

        Grouping the library requires embedding the collection descriptions,
        which means loading the model - far too slow to do before the app is
        usable. The grouping only changes when the library does, so it is
        computed during a sync and simply read back here.
        """
        by_id = self._read_json(self._collections_cache_path())
        if not by_id:
            return {name: [] for name in MANUAL_COLLECTIONS}
        position = {
            (meta or {}).get("item_id"): i for i, meta in enumerate(self.engine.meta or [])
        }
        result = {name: [] for name in MANUAL_COLLECTIONS}
        for name, item_ids in by_id.items():
            if not isinstance(item_ids, list):
                continue
            # Empty collections are kept: they are the drop targets, so they
            # have to stay on screen before anything has been put in them.
            result[name] = [position[i] for i in item_ids if i in position]
        return result

    def _refresh_collections(self, from_cache=False):
        """Groups the loaded index into collections, off the UI thread."""
        def work():
            try:
                # Always read what was curated by hand. The automatic
                # grouping (_build_categories) is no longer applied - it
                # needed the CLIP model and produced sets nobody chose.
                collections = self._load_collections()
                self.event_queue.put(("collections_ready", collections))
            except Exception:
                _log_crash("building collections", sys.exc_info())

        self._run_in_background(work)

    def _rebuild_collection_cards(self):
        """Recreates the auto-collection cards after an index finishes."""
        for card in self.collection_cards.values():
            card.destroy()
        self.collection_cards = {}
        for name, idxs in (self._collections or {}).items():
            self.collection_cards[name] = self._make_collection_card(
                self.sidebar, f"  {name} ({len(idxs)})", lambda n=name: self._set_view(n)
            )

    def _set_view(self, view):
        """Switches which collection the search box searches."""
        self._view = view
        self.all_card.configure(
            text="  All photos", fg_color=ACCENT_COLOR if view == "all" else CARD_COLOR
        )
        self.fav_card.configure(fg_color=ACCENT_COLOR if view == "favourites" else CARD_COLOR)
        self.history_card.configure(fg_color=ACCENT_COLOR if view == "history" else CARD_COLOR)
        self._update_favourites_button()
        self._update_history_button()
        for name, card in self.collection_cards.items():
            card.configure(fg_color=ACCENT_COLOR if view == name else CARD_COLOR)

        if view == "history":
            self.query_entry.configure(placeholder_text="Describe an image in English...")
            self._show_history()
        elif view == "favourites":
            self.query_entry.configure(placeholder_text="Search your favourites...")
            self._show_favourites()
        elif view in (self._collections or {}):
            self.query_entry.configure(placeholder_text=f"Search within {view}...")
            self._show_collection(view)
        else:
            self.query_entry.configure(placeholder_text="Describe an image in English...")
            self._show_all_photos()

    def _show_all_photos(self):
        """Opening view: a cross-section of the library rather than a blank grid.

        Showing nothing until the first search makes an app holding tens of
        thousands of photos look empty, and gives no sense of what is in
        there. Stepping through the whole library (rather than taking the
        first N, which would all come from whichever folder indexed first)
        and then interleaving by folder gives a varied sample in one screen.
        """
        total = len(self.engine.paths)
        if not total:
            self._render_results([])
            self.status_var.set(
                "No photos loaded yet. Click 'Check for new photos' to load the "
                "shared drive, or 'Choose Folder...' for photos on this computer."
            )
            return
        step = max(1, total // ALL_PHOTOS_PREVIEW)
        idxs = list(range(0, total, step))[:ALL_PHOTOS_PREVIEW]
        results = [(self.engine.paths[i], 1.0, self.engine.meta[i]) for i in idxs]
        self._render_results(self._interleave_by_folder(results))
        self.status_var.set(
            f"{total:,} photos ready. Type what you are looking for, "
            "or pick a collection on the left."
        )

    def _show_collection(self, name):
        idxs = (self._collections or {}).get(name, [])
        if not idxs:
            self._render_results([])
            self.status_var.set(
                f"{name} is empty - drag photos onto it, or right-click a photo "
                f"and choose 'Add to {name}'."
            )
            return
        results = [(self.engine.paths[i], 1.0, self.engine.meta[i]) for i in idxs[:DEFAULT_TOP_K]]
        self._render_results(results)
        extra = f" (showing first {DEFAULT_TOP_K})" if len(idxs) > DEFAULT_TOP_K else ""
        self.status_var.set(f"{len(idxs)} photo(s) in {name}{extra}. Search to narrow it down.")

    def _show_favourites(self):
        """Lists every favourite, newest first, with no search term needed."""
        favs = self._load_favourites()
        if not favs:
            self._render_results([])
            self.status_var.set("No favourites yet - click the star on any result to save it.")
            return
        by_key = {}
        for i, meta in enumerate(self.engine.meta or []):
            by_key[self._fav_key(self.engine.paths[i], meta)] = i

        results = []
        for key, entry in reversed(list(favs.items())):
            idx = by_key.get(key)
            if idx is not None:
                results.append((self.engine.paths[idx], 1.0, self.engine.meta[idx]))
            elif entry.get("path") and os.path.exists(entry["path"]):
                # Favourited before the current index was loaded - still show it.
                results.append((entry["path"], 1.0, entry))
        self._render_results(results)
        self.status_var.set(f"{len(results)} favourite(s).")

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

    def _active_subset(self):
        """Indices the current collection restricts search to, or None."""
        if self._view == "favourites":
            return self._favourite_indices()
        if self._view in (self._collections or {}):
            return self._collections[self._view]
        return None

    def _exclude_graphic(self, subset):
        """Narrows a search subset to exclude finished graphics (posters,
        adverts, anything with text burnt into the image) - see
        _graphic_mask. Collections already exclude these at build time, so
        this mainly matters for "All photos" (subset is None, so it isn't
        narrowed to anything else yet), but is applied uniformly so a
        graphic can never surface from a text search regardless of view.
        """
        mask = self._graphic_mask()
        if len(mask) == 0:
            return subset
        non_graphic = set(np.nonzero(~mask)[0].tolist())
        return non_graphic if subset is None else (set(subset) & non_graphic)

    def _interleave_by_folder(self, results):
        """Round-robins results across SharePoint folders instead of
        leaving them in pure relevance order.

        A query can legitimately have its best 40+ matches all sitting in
        one folder (e.g. one heavily-photographed shoot) - the near-
        duplicate filter in search_engine.py doesn't touch this, and
        shouldn't, since these are genuinely different, non-duplicate
        photos. Without this, that one folder fills the entire top of the
        results grid before any other folder gets a look in.

        Purely reorders - nothing is dropped or re-scored, and each
        folder's own internal relevance order is preserved. Items with no
        known folder (local-folder mode, or a name that wasn't captured)
        never compete with each other for a turn; each is its own group.
        """
        if not results:
            return results
        groups = {}
        order = []
        for i, (path, score, meta) in enumerate(results):
            folder = (meta or {}).get("folder") if meta else None
            key = folder if folder else f"__solo_{i}"
            if key not in groups:
                groups[key] = []
                order.append(key)
            groups[key].append((path, score, meta))

        interleaved = []
        while len(interleaved) < len(results):
            for key in order:
                queue = groups[key]
                if queue:
                    interleaved.append(queue.pop(0))
        return interleaved

    def _start_text_search(self):
        query = self.query_var.get().strip()
        if not query:
            return
        # The Search button is disabled while indexing, but <Return> in the
        # entry reaches this directly - without this guard a search could run
        # against an index the sync is in the middle of replacing.
        if self._busy:
            self.status_var.set(
                "Still loading photos - one moment, then try that search again."
            )
            return
        if self.engine.embeddings is None or len(self.engine.paths) == 0:
            messagebox.showinfo("Image Search", "Index the folder first before searching.")
            return
        subset = self._active_subset()
        if self._view == "favourites" and not subset:
            self.status_var.set("No favourites yet - click the star on any result to save it.")
            return

        self._last_query = query
        self._record_search(query)
        self._set_busy(True)
        self.progress.configure(mode="indeterminate")
        self.progress.start()

        def work():
            def on_status(msg):
                self.event_queue.put(("status", msg))

            try:
                self.engine.load_model(on_status)
                effective_subset = self._exclude_graphic(subset)
                bonus = self._compute_bonus(query)
                results = self.engine.search_text(
                    query, DEFAULT_TOP_K, status_callback=on_status, bonus=bonus,
                    subset=effective_subset,
                )
                if self._query_wants_objects(query):
                    self._last_search_mode = "objects"
                    # Relevance-preserving: splits on "has guests in it" and
                    # keeps each half in the engine's order, so this can be
                    # applied to the whole list without scrambling it.
                    results = self._order_for_object_query(results)
                    head, tail = results[:RELEVANCE_HEAD], results[RELEVANCE_HEAD:]
                    results = head + self._interleave_by_folder(tail)
                else:
                    self._last_search_mode = "scene"
                    head, tail = results[:RELEVANCE_HEAD], results[RELEVANCE_HEAD:]
                    tail = self._interleave_by_folder(tail)
                    tail = self._order_by_orientation_and_people(tail)
                    results = head + tail
                self.event_queue.put(("search_done", results))
            except Exception as exc:
                self.event_queue.put(("error", str(exc)))

        self._run_in_background(work)

    def _start_image_search(self):
        """Searches with a photo instead of words.

        Some things are far easier to show than to describe - the look of a
        campaign, a particular lighting, a composition. The reference image
        is embedded with the same model as the library, so it ranks exactly
        as a typed query would.
        """
        if self._busy:
            self.status_var.set("Still loading photos - one moment, then try again.")
            return
        if self.engine.embeddings is None or len(self.engine.paths) == 0:
            messagebox.showinfo("Image Search", "Load the photo library first.")
            return
        path = filedialog.askopenfilename(
            title="Pick a photo to find similar ones",
            filetypes=[
                ("Images", "*.jpg *.jpeg *.png *.bmp *.gif *.webp *.tif *.tiff"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return

        subset = self._exclude_graphic(self._active_subset())
        self._last_query = ""
        self._set_busy(True)
        self.progress.configure(mode="indeterminate")
        self.progress.start()
        self.status_var.set(f"Finding photos like {os.path.basename(path)}...")

        def work():
            def on_status(msg):
                self.event_queue.put(("status", msg))

            try:
                results = self.engine.search_image(
                    path, DEFAULT_TOP_K, status_callback=on_status, subset=subset
                )
                # Same shape as a text search: the best matches stay in the
                # engine's order, variety is applied only below them.
                head, tail = results[:RELEVANCE_HEAD], results[RELEVANCE_HEAD:]
                self._last_search_mode = "image"
                self.event_queue.put(("search_done", head + self._interleave_by_folder(tail)))
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
        self._busy = busy
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
        state = tk.NORMAL if has_index else tk.DISABLED
        self.search_button.configure(state=state)
        if getattr(self, "image_search_button", None) is not None:
            self.image_search_button.configure(state=state)

    # ---------- search-time ranking signals ----------

    # ---------- search history ----------

    def _history_path(self):
        local_folder, _t, _f, _i = self._sp_cache_dirs()
        return os.path.join(local_folder, "search_history.json")

    def _load_history(self):
        """Most recent first: [{"term": ..., "at": iso timestamp}, ...]"""
        if self._history is None:
            data = self._read_json(self._history_path())
            entries = data.get("searches") if isinstance(data, dict) else None
            self._history = entries if isinstance(entries, list) else []
        return self._history

    def _record_search(self, term):
        term = (term or "").strip()
        if not term:
            return
        history = self._load_history()
        # Repeating a search updates its time rather than adding a duplicate,
        # so the list stays a record of what was looked for, not how often
        # the same thing was retyped.
        normalized = normalize_text(term)
        history[:] = [h for h in history if normalize_text(h.get("term", "")) != normalized]
        history.insert(0, {"term": term, "at": datetime.datetime.now().isoformat(timespec="minutes")})
        del history[HISTORY_MAX:]
        try:
            with open(self._history_path(), "w", encoding="utf-8") as f:
                json.dump({"searches": history}, f, ensure_ascii=False)
        except OSError:
            pass
        self._refresh_suggestions()
        self._update_history_button()

    def _top_terms(self, count=5):
        """The terms searched most often, for the shortcut buttons."""
        counts = {}
        for entry in self._load_history():
            term = (entry.get("term") or "").strip()
            if term:
                counts[term] = counts.get(term, 0) + entry.get("uses", 1)
        return [t for t, _ in sorted(counts.items(), key=lambda kv: -kv[1])][:count]

    def _refresh_suggestions(self):
        """Rebuilds the shortcut buttons under the search box."""
        if getattr(self, "suggestion_bar", None) is None:
            return
        for child in self.suggestion_bar.winfo_children():
            child.destroy()
        terms = [h.get("term") for h in self._load_history()[:5] if h.get("term")]
        if not terms:
            return
        ctk.CTkLabel(
            self.suggestion_bar, text="Recent:", text_color=MUTED_TEXT_COLOR,
            font=ctk.CTkFont(size=11),
        ).pack(side=tk.LEFT, padx=(2, 6))
        for term in terms:
            ctk.CTkButton(
                self.suggestion_bar, text=term, height=26, corner_radius=13,
                fg_color=CARD_COLOR, hover_color=ACCENT_COLOR, text_color=TEXT_COLOR,
                font=ctk.CTkFont(size=11), width=0,
                command=lambda t=term: self._search_term(t),
            ).pack(side=tk.LEFT, padx=3)

    def _search_term(self, term):
        """Runs a search for a saved term, switching back to the full library
        so a stored search isn't silently narrowed by whatever collection
        happened to be open."""
        if self._view not in ("all", "favourites") and self._view != "history":
            pass
        if self._view == "history":
            self._set_view("all")
        self.query_var.set(term)
        self._start_text_search()

    def _show_history(self):
        history = self._load_history()
        for child in self.results_frame.winfo_children():
            child.destroy()
        self.thumbnail_refs.clear()
        if not history:
            self.status_var.set("No searches yet - your search history will appear here.")
            return
        self.status_var.set(f"{len(history)} previous search(es). Click one to run it again.")
        for entry in history:
            term = entry.get("term", "")
            when = (entry.get("at") or "").replace("T", " ")
            row = ctk.CTkFrame(self.results_frame, fg_color=CARD_COLOR, corner_radius=10)
            row.pack(fill=tk.X, padx=4, pady=3)
            ctk.CTkButton(
                row, text=term, anchor="w", height=34, corner_radius=8,
                fg_color="transparent", hover_color=ACCENT_COLOR, text_color=TEXT_COLOR,
                font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
                command=lambda t=term: self._search_term(t),
            ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 4), pady=4)
            ctk.CTkLabel(
                row, text=when, text_color=MUTED_TEXT_COLOR, font=ctk.CTkFont(size=11),
            ).pack(side=tk.RIGHT, padx=12)

    # ---------- fast local index ----------

    def _fast_cache_paths(self):
        local_folder, _t, _f, _i = self._sp_cache_dirs()
        return (
            os.path.join(local_folder, "index_embeddings.npy"),
            os.path.join(local_folder, "index_meta.json"),
        )

    def _save_fast_cache(self, sp_items):
        """Writes the searchable index in a form that loads quickly.

        The per-folder JSON files are the shared, authoritative copy, but
        parsing them all costs seconds: most of their bytes are base64
        thumbnails that get discarded, and every embedding is stored as
        text. Keeping a compact binary copy means opening the app doesn't
        have to pay that.
        """
        emb_path, meta_path = self._fast_cache_paths()
        try:
            if sp_items:
                np.save(emb_path, np.stack([i["embedding"] for i in sp_items]))
            meta = {
                "paths": [i["path"] for i in sp_items],
                "meta": [i["meta"] for i in sp_items],
            }
            with open(meta_path, "w", encoding="utf-8") as f:
                json.dump(meta, f, ensure_ascii=False)
        except (OSError, ValueError) as exc:
            _log_reuse(f"FAST CACHE SAVE FAILED: {exc!r}")

    def _load_fast_cache(self):
        """Loads the compact index if present. Returns image count, or 0."""
        emb_path, meta_path = self._fast_cache_paths()
        if not (os.path.exists(emb_path) and os.path.exists(meta_path)):
            return 0
        try:
            embeddings = np.load(emb_path)
            with open(meta_path, "r", encoding="utf-8") as f:
                meta = json.load(f)
            paths, metas = meta.get("paths") or [], meta.get("meta") or []
            if len(paths) != len(embeddings) or len(metas) != len(embeddings):
                return 0
            self.engine.folder = None
            self.engine.paths = paths
            self.engine.meta = metas
            self.engine.embeddings = embeddings
            return len(paths)
        except Exception as exc:
            _log_reuse(f"FAST CACHE LOAD FAILED: {exc!r}")
            return 0

    def _startup_load(self):
        """Makes the app usable immediately, without signing in.

        Everything needed to search is already on disk from the last run,
        so waiting for a SharePoint sync before the first search is pure
        dead time. The sync still happens - when the user asks for it - and
        replaces this with fresh results.
        """
        def work():
            try:
                count = self._load_fast_cache()
                if count:
                    self.event_queue.put(("startup_ready", count))
                # Warm the CLIP model in the background. It's needed to
                # encode a search query, and loading it on the first search
                # would put several seconds squarely in the user's way -
                # here it happens while they're still reading the window.
                self.engine.load_model()
            except Exception:
                _log_crash("startup load", sys.exc_info())

        self._run_in_background(work)

    # ---------- settings ----------

    def _settings_path(self):
        local_folder, _t, _f, _i = self._sp_cache_dirs()
        return os.path.join(os.path.dirname(local_folder), "settings.json")

    def _load_settings(self):
        if self._settings is None:
            self._settings = self._read_json(self._settings_path())
        return self._settings

    def _save_settings(self):
        try:
            os.makedirs(os.path.dirname(self._settings_path()), exist_ok=True)
            with open(self._settings_path(), "w", encoding="utf-8") as f:
                json.dump(self._load_settings(), f, ensure_ascii=False)
        except OSError:
            pass

    def _download_folder(self):
        folder = self._load_settings().get("download_folder") or ""
        return folder if folder and os.path.isdir(folder) else ""

    def _open_settings(self):
        """Small modal for the one setting worth exposing: where opened
        images should be saved."""
        win = ctk.CTkToplevel(self.root)
        win.title("Settings")
        win.geometry("540x230")
        win.configure(fg_color=BG_COLOR)
        win.transient(self.root)
        win.grab_set()

        ctk.CTkLabel(
            win, text="Save opened images to", text_color=TEXT_COLOR,
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
        ).pack(anchor="w", padx=20, pady=(20, 2))
        ctk.CTkLabel(
            win,
            text="Double-clicking a result saves a copy here, then opens it.\n"
                 "Leave empty to just open images without keeping a copy.",
            text_color=MUTED_TEXT_COLOR, justify="left",
        ).pack(anchor="w", padx=20, pady=(0, 10))

        row = ctk.CTkFrame(win, fg_color=CARD_COLOR, corner_radius=CARD_RADIUS)
        row.pack(fill=tk.X, padx=20)
        folder_var = tk.StringVar(value=self._download_folder() or "Not set")
        ctk.CTkLabel(
            row, textvariable=folder_var, text_color=MUTED_TEXT_COLOR, anchor="w",
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=14, pady=12)

        def choose():
            chosen = filedialog.askdirectory(title="Choose a download folder", parent=win)
            if chosen:
                self._load_settings()["download_folder"] = chosen
                self._save_settings()
                folder_var.set(chosen)

        def clear():
            self._load_settings().pop("download_folder", None)
            self._save_settings()
            folder_var.set("Not set")

        _make_button(row, "Choose...", choose, width=110).pack(side=tk.LEFT, padx=(0, 8), pady=12)
        _make_button(row, "Clear", clear, width=80).pack(side=tk.LEFT, padx=(0, 14), pady=12)
        _make_button(win, "Done", win.destroy, width=100).pack(pady=18)

    # ---------- collections ----------

    def _load_category_defs(self):
        """Collection definitions, from the user's copy if they've made one.

        A user-editable override means categories can be retuned (or new
        ones added) without rebuilding and redistributing the app.
        """
        local_folder, _t, _f, _i = self._sp_cache_dirs()
        override = os.path.join(os.path.dirname(local_folder), "categories.json")
        for path in (override, _resource_path("assets", "categories.json")):
            data = self._read_json(path)
            cats = data.get("categories")
            if isinstance(cats, list) and cats:
                return cats
        return []

    def _people_score(self):
        """Per-image "has guests in it" score for the whole loaded index,
        cached since the concept vectors never change between calls."""
        if self._people_vec is None:
            self._people_vec = self.engine.embed_concept(PEOPLE_PROMPTS)
            self._empty_vec = self.engine.embed_concept(EMPTY_PROMPTS)
        return self.engine.score_against(self._people_vec) - self.engine.score_against(self._empty_vec)

    def _graphic_mask(self):
        """Per-image "is this finished artwork with text on it" mask for
        the whole loaded index. The concept vectors are cached since they
        never change between calls, but the mask itself is recomputed each
        time (same pattern as _people_score) since it depends on whichever
        index is currently loaded.

        Used to keep posters, adverts, and anything with wording burnt into
        the image out of collections *and* out of plain search results -
        they're not raw source material a marketer would want to reuse.
        """
        if self._graphic_vec is None:
            self._graphic_vec = self.engine.embed_concept([
                "a poster with large text written on it",
                "an advertisement with a headline and words",
                "a graphic design with a slogan overlaid",
                "a sign with writing on it",
            ])
            self._plain_vec = self.engine.embed_concept([
                "a candid photograph of people with no text",
                "a plain photograph with no writing",
                "a natural photo without any words",
            ])
        return (
            self.engine.score_against(self._graphic_vec)
            - self.engine.score_against(self._plain_vec)
        ) > TEXT_OVERLAY_MARGIN

    def _query_wants_objects(self, query):
        """True when the query asks for a thing rather than for a scene.

        This is a lifestyle photo library, so nearly every image has guests
        somewhere in it. Ordering people-first is right for "birthday party"
        and wrong for "burger", where the wanted photo is the dish, not the
        table it is sitting on. Comparing the query itself against the two
        directions decides which rule applies - rather than a hand-written
        list of food words, which would never cover a whole menu in two
        languages.
        """
        if self._people_vec is None:
            self._people_vec = self.engine.embed_concept(PEOPLE_PROMPTS)
            self._empty_vec = self.engine.embed_concept(EMPTY_PROMPTS)
        if self._object_vec is None:
            self._object_vec = self.engine.embed_concept(OBJECT_PROMPTS)
        try:
            q = self.engine.embed_concept([query])
        except Exception:
            return False
        return float(q @ self._object_vec) > float(q @ self._people_vec)

    def _order_for_object_query(self, results):
        """Puts photos that are actually OF the thing ahead of photos that
        merely contain it.

        A stable partition on "has guests in it", so relevance order is kept
        inside each half and the people shots stay directly underneath
        rather than being thrown away - the opposite of the orientation
        sort below, which is what a query for a scene wants.
        """
        if not results:
            return results
        path_to_index = {p: i for i, p in enumerate(self.engine.paths)}
        scores = self._people_score()

        def has_people(path):
            idx = path_to_index.get(path)
            return bool(idx is not None and idx < len(scores) and scores[idx] > 0)

        objects = [r for r in results if not has_people(r[0])]
        people = [r for r in results if has_people(r[0])]
        return objects + people

    def _order_by_orientation_and_people(self, results):
        """Re-orders a set of search results: landscape-with-guests,
        landscape, portrait-with-guests, portrait.

        Applies the same rule collections already use, but to plain search
        results too - a text search only ranks by relevance, so without
        this a query like "golf gameplay nighttime" mixes orientations and
        empty-venue shots in with photos of guests, in whatever order
        similarity happened to produce.

        Only touches the returned results (at most a couple hundred
        images), not the whole library, and is a stable sort - within each
        of the four groups, the original relevance order is preserved.
        """
        if not results:
            return results
        path_to_index = {p: i for i, p in enumerate(self.engine.paths)}
        idxs = [path_to_index.get(path) for path, _score, _meta in results]

        has_people = {}
        valid = [i for i in idxs if i is not None]
        if valid:
            scores = self._people_score()
            for i in valid:
                has_people[i] = bool(scores[i] > 0)

        def key(item, idx):
            if idx is None:
                return (False, False)
            return (self._is_landscape(idx), has_people.get(idx, False))

        order = sorted(range(len(results)), key=lambda n: key(results[n], idxs[n]), reverse=True)
        return [results[n] for n in order]

    def _build_categories(self):
        """Assigns each indexed image to at most one collection.

        Two signals, deliberately weighted differently:
        - the SharePoint folder name, treated as decisive, because a photo
          in a folder called "Familie" is a family photo whatever the
          pixels suggest;
        - otherwise visual similarity to the collection's description.

        Collections marked require_people exclude photos with no guests in
        them, so an empty venue stays searchable without turning up in a
        highlight collection.
        """
        defs = self._load_category_defs()
        if not defs or not self.engine.paths:
            return {}

        has_people = self._people_score() > 0

        # Finished graphics - posters, adverts, anything with wording burnt
        # into the image - are not raw source material, so they're kept out
        # of collections entirely, and out of plain search too (see
        # _start_text_search).
        is_graphic = self._graphic_mask()

        scores = []
        for cat in defs:
            prompts = cat.get("prompts") or [cat.get("name", "")]
            scores.append(self.engine.score_against(self.engine.embed_concept(prompts)))
        S = np.stack(scores, axis=1)

        # Deliberately strict: a collection is a curated highlight reel, not
        # everything that vaguely matches. Only the strongest matches are
        # admitted, so a browsable collection stays worth browsing.
        best = np.argmax(S, axis=1)
        best_score = S[np.arange(len(S)), best]
        threshold = (
            float(np.percentile(best_score, COLLECTION_PERCENTILE)) if len(best_score) else 0.0
        )

        collections = {cat["name"]: [] for cat in defs}
        for i, meta in enumerate(self.engine.meta):
            if is_graphic[i]:
                continue  # finished artwork, not a photo to reuse
            folder = (meta or {}).get("folder") or ""
            placed = False
            if folder:
                normalized = normalize_text(folder)
                for ci, cat in enumerate(defs):
                    if any(normalize_text(k) in normalized for k in cat.get("folder_keywords") or []):
                        collections[cat["name"]].append(i)
                        placed = True
                        break
            if placed:
                continue
            ci = int(best[i])
            if best_score[i] <= threshold:
                continue
            if defs[ci].get("require_people", True) and not has_people[i]:
                continue
            collections[defs[ci]["name"]].append(i)

        # Ordering, strongest first: landscape with guests, landscape,
        # portrait with guests, portrait. Landscape leads because that's
        # what marketing material is usually laid out for; a tall phone
        # photo rarely fits a banner however good it is.
        result = {}
        for ci, cat in enumerate(defs):
            idxs = collections[cat["name"]]
            if not idxs:
                continue
            idxs.sort(
                key=lambda i: (self._is_landscape(i), bool(has_people[i]), float(S[i, ci])),
                reverse=True,
            )
            result[cat["name"]] = idxs[:COLLECTION_MAX]
        return result

    def _is_landscape(self, index):
        """Whether an image is wider than it is tall.

        Read from the cached thumbnail's header only - the pixel data is
        never decoded - and memoised, since the same image is compared many
        times while sorting. Unknown sizes count as portrait so a missing
        thumbnail can't jump the queue.
        """
        cached = self._landscape_cache.get(index)
        if cached is not None:
            return cached
        try:
            with Image.open(self.engine.paths[index]) as img:
                width, height = img.size
            landscape = width > height
        except Exception:
            landscape = False
        self._landscape_cache[index] = landscape
        return landscape

    # ---------- favourites ----------

    @staticmethod
    def _fav_key(path, meta):
        """Stable id for a result. SharePoint items keep their driveItem id
        so favourites survive re-indexing and cache clears; local-folder
        images fall back to their path."""
        if meta and meta.get("item_id"):
            return meta["item_id"]
        return path

    def _favourites_path(self):
        local_folder, _t, _f, _i = self._sp_cache_dirs()
        return os.path.join(local_folder, "favourites.json")

    def _load_favourites(self):
        """This machine's saved images, plus any shared by the team."""
        if self._favourites is None:
            merged = dict(self._read_json(self._team_favourites_path()).get("items") or {})
            # Local saves may not have been shared yet, so they go on top.
            merged.update(self._read_json(self._favourites_path()).get("items") or {})
            self._favourites = merged
        return self._favourites

    def _own_favourites(self):
        return self._read_json(self._favourites_path()).get("items") or {}

    def _is_favourite(self, key):
        """Whether *you* saved this image.

        Deliberately not the merged team view: the star is a toggle for your
        own list, so it has to reflect what clicking it will change. The
        Favourites collection still shows everything you and the team saved.
        """
        return key in self._own_favourites()

    def _toggle_favourite(self, key, path, meta):
        """Adds or removes one image from this machine's own favourites.

        Only this machine's file is written; entries shared by teammates are
        left alone, so un-starring can't delete someone else's shortlist.
        """
        own = self._own_favourites()
        if key in own:
            del own[key]
            added = False
        else:
            # Store enough to render a favourite even when it isn't part of
            # the currently loaded index (e.g. before a SharePoint sync).
            own[key] = {
                "path": path,
                "name": (meta or {}).get("name", os.path.basename(path)),
                "folder": (meta or {}).get("folder", ""),
                "item_id": (meta or {}).get("item_id", ""),
            }
            added = True
        try:
            os.makedirs(os.path.dirname(self._favourites_path()), exist_ok=True)
            with open(self._favourites_path(), "w", encoding="utf-8") as f:
                json.dump({"items": own}, f, ensure_ascii=False)
        except OSError:
            pass
        self._favourites = None  # rebuild the merged view
        self._update_favourites_button()
        return added

    def _team_favourites_path(self):
        local_folder, _t, _f, _i = self._sp_cache_dirs()
        return os.path.join(local_folder, "favourites_team.json")

    def _sync_favourites(self, root_id):
        """Publishes this machine's favourites and merges in the team's.

        Same one-file-per-machine approach as the click feedback: nobody
        writes to anyone else's file, so two people saving at once cannot
        clobber each other and the merge is a plain union.
        """
        own = self._read_json(self._favourites_path()).get("items") or {}
        own_name = f"{FAVOURITES_PREFIX}{machine_key()}.json"
        try:
            # Uploaded even when empty: skipping the upload left a removed
            # favourite sitting in this machine's shared file forever, so
            # un-starring never reached anyone else.
            self.sp_client.upload_json_file(root_id, own_name, {"items": own})
            merged = {}
            for name in self.sp_client.list_shared_files(root_id, FAVOURITES_PREFIX):
                # This machine's own file is skipped: the local copy is
                # authoritative for it. Merging it back in meant a favourite
                # you had just un-starred returned through the team list and
                # could never be removed from the Favourites view.
                if os.path.basename(name) == own_name:
                    continue
                remote = self.sp_client.download_json_file(root_id, name)
                for key, entry in ((remote or {}).get("items") or {}).items():
                    if isinstance(entry, dict):
                        merged.setdefault(key, entry)
            with open(self._team_favourites_path(), "w", encoding="utf-8") as f:
                json.dump({"items": merged}, f, ensure_ascii=False)
            self._favourites = None  # reload including the team's
            _log_reuse(f"FAVOURITES SYNC: shared={len(own)} team_total={len(merged)}")
        except Exception as exc:
            _log_reuse(f"FAVOURITES SYNC FAILED: {exc!r}")

    def _favourite_indices(self):
        """Positions in the loaded index that are favourited."""
        favs = self._load_favourites()
        if not favs or not self.engine.meta:
            return []
        return [
            i
            for i, meta in enumerate(self.engine.meta)
            if self._fav_key(self.engine.paths[i], meta) in favs
        ]

    def _feedback_path(self):
        """This machine's own contributions - the file that gets uploaded."""
        local_folder, _t, _f, _i = self._sp_cache_dirs()
        return os.path.join(local_folder, "click_feedback.json")

    def _team_feedback_path(self):
        """Everyone's contributions merged together - used for ranking."""
        local_folder, _t, _f, _i = self._sp_cache_dirs()
        return os.path.join(local_folder, "click_feedback_team.json")

    def _read_json(self, path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except (json.JSONDecodeError, OSError):
            return {}

    @staticmethod
    def _merge_feedback(into, other):
        """Adds one feedback map into another. Counts are additive, so
        merging is order-independent and can't lose anyone's data."""
        for key, items in (other or {}).items():
            if not isinstance(items, dict):
                continue
            target = into.setdefault(key, {})
            for item_id, count in items.items():
                try:
                    target[item_id] = target.get(item_id, 0) + int(count)
                except (TypeError, ValueError):
                    continue
        return into

    def _load_feedback(self):
        """{hashed query: {item_id: times opened}}, this machine + team."""
        if self._feedback is None:
            merged = self._read_json(self._team_feedback_path())
            # Local clicks may not be uploaded/merged yet, so overlay them.
            self._merge_feedback(merged, self._read_json(self._feedback_path()))
            self._feedback = merged
        return self._feedback

    def _record_click(self, item_id):
        """Remembers that an image was opened for the last query run.

        Opening a result is the clearest signal available that it was the
        right answer, and it was previously discarded. Recording it lets
        frequently-searched topics improve with use, without any labelling.
        """
        key = query_key(self._last_query)
        if not key or not item_id:
            return
        own = self._read_json(self._feedback_path())
        own.setdefault(key, {})[item_id] = own.get(key, {}).get(item_id, 0) + 1
        try:
            with open(self._feedback_path(), "w", encoding="utf-8") as f:
                json.dump(own, f)
        except OSError:
            pass
        # Keep the in-memory view current so the next search reflects it.
        if self._feedback is not None:
            self._feedback.setdefault(key, {})[item_id] = (
                self._feedback.get(key, {}).get(item_id, 0) + 1
            )

    def _sync_feedback(self, root_id):
        """Publishes this machine's feedback and merges in everyone else's.

        Each machine owns exactly one file, so no upload can clobber another
        user's data and the merge is a simple sum.
        """
        own = self._read_json(self._feedback_path())
        try:
            if own:
                self.sp_client.upload_json_file(
                    root_id, f"{FEEDBACK_PREFIX}{machine_key()}.json", own
                )
            merged = {}
            for name in self.sp_client.list_feedback_files(root_id):
                remote = self.sp_client.download_json_file(root_id, name)
                if remote:
                    self._merge_feedback(merged, remote)
            with open(self._team_feedback_path(), "w", encoding="utf-8") as f:
                json.dump(merged, f)
            self._feedback = None  # force reload with the merged data
            _log_reuse(f"FEEDBACK SYNC: queries={len(merged)}")
        except Exception as exc:
            _log_reuse(f"FEEDBACK SYNC FAILED: {exc!r}")

    def _compute_bonus(self, query):
        """Per-image score adjustments the image embedding can't provide:
        the name of the folder an image sits in, and prior opens for this
        same query. Returns None when neither applies."""
        if not self.engine.meta:
            return None
        feedback_for_query = self._load_feedback().get(query_key(query), {})
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
                self._save_fast_cache(sp_items)

                self.event_queue.put(("status", "Syncing team favourites and feedback..."))
                self._sync_feedback(root["id"])
                self._sync_favourites(root["id"])

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

    def _deliver_to_download_folder(self, source, name):
        """Copies an opened image into the user's chosen folder, if set.

        Returns the delivered path, or the original when no folder is
        configured. Never overwrites: a second copy of the same filename
        gets a numeric suffix rather than silently replacing what's there.
        """
        folder = self._download_folder()
        if not folder:
            return source
        target = os.path.join(folder, name)
        if os.path.exists(target):
            stem, ext = os.path.splitext(name)
            n = 2
            while os.path.exists(os.path.join(folder, f"{stem} ({n}){ext}")):
                n += 1
            target = os.path.join(folder, f"{stem} ({n}){ext}")
        try:
            shutil.copy2(source, target)
            return target
        except OSError:
            return source

    def _open_sp_result(self, item_id, name):
        self._record_click(item_id)
        _local_folder, _thumbs_dir, full_dir, _indexes_dir = self._sp_cache_dirs()
        dest = os.path.join(full_dir, f"{item_id}_{name}")
        if os.path.exists(dest):
            self.event_queue.put(("sp_open_ready", dest, name))
            return

        self._set_busy(True)
        self.status_var.set(f"Downloading {name}...")

        def work():
            try:
                if self.sp_client is None:
                    # Instant startup means a result can be opened before
                    # "Check for new photos" has ever run this session,
                    # so the SharePoint connection may not exist yet.
                    self.event_queue.put(("status", "Signing in to Microsoft..."))
                    self.sp_client = SharePointClient()
                    self.sp_client.sign_in()
                    self.event_queue.put(("status", f"Downloading {name}..."))
                item = {"id": item_id}
                self.sp_client.download_file(item, dest)
                self.event_queue.put(("sp_open_ready", dest, name))
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
                    mode = getattr(self, "_last_search_mode", "")
                    note = ""
                    if mode == "objects":
                        note = "  \u00b7  Close-ups first, people shots below."
                    elif mode == "image":
                        note = "  \u00b7  Most visually similar first."
                    self.status_var.set(f"Found {len(event[1])} results.{note}")
                elif kind == "clipboard_status":
                    message, target = event[2], event[1]
                    preview = self._preview
                    shown = False
                    if target is not None and preview is not None and preview["token"] is target:
                        try:
                            preview["caption"].configure(text=message)
                            shown = True
                        except Exception:
                            shown = False
                    if not shown:
                        self.status_var.set(message)
                elif kind == "preview_ready":
                    preview = self._preview
                    if preview is not None and preview["token"] is event[1]:
                        try:
                            img = self._fit(
                                Image.open(event[2]).convert("RGB"), PREVIEW_MAX
                            )
                            photo = ctk.CTkImage(
                                light_image=img, dark_image=img, size=img.size
                            )
                            preview["label"].configure(image=photo)
                            preview["label"]._preview_ref = photo
                            text = preview["caption"].cget("text")
                            preview["caption"].configure(
                                text=text.replace("   (sharpening...)", "")
                            )
                        except Exception:
                            pass
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
                    self._refresh_collections()
                elif kind == "startup_ready":
                    self._sp_mode_active = True
                    self.folder_var.set("SharePoint: Fotos & Videos (all subfolders)")
                    self._set_busy(False)
                    # Fills the opening grid now the cached index is loaded,
                    # instead of leaving an empty window until the first search.
                    self._show_all_photos()
                    # Past searches are on disk, so the shortcut chips can be
                    # there from the start rather than only after a search.
                    self._refresh_suggestions()
                    self._refresh_collections(from_cache=True)
                elif kind == "collections_ready":
                    self._collections = event[1]
                    self._rebuild_collection_cards()
                elif kind == "sp_open_ready":
                    self._set_busy(False)
                    delivered = self._deliver_to_download_folder(event[1], event[2])
                    if self._download_folder():
                        self.status_var.set(f"Saved to {os.path.dirname(delivered)}")
                    else:
                        self.status_var.set("Ready to search.")
                    open_in_system_viewer(delivered)
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

    def _desired_columns(self):
        """How many result columns fit the results area at its current
        width, so the grid fills the window (and expands to fill it when
        maximized) instead of sitting at a fixed column count."""
        width = self.results_frame.winfo_width()
        if width <= 1:
            return GRID_COLUMNS
        cell_width = THUMB_SIZE[0] + GRID_CELL_PADDING
        return max(1, (width - GRID_CELL_PADDING) // cell_width)

    def _on_results_area_resized(self):
        """Re-lays out the current results at the new column count, once
        resizing has paused rather than on every intermediate size during a
        window drag."""
        if self._resize_after_id is not None:
            try:
                self.root.after_cancel(self._resize_after_id)
            except Exception:
                pass
        self._resize_after_id = self.root.after(RESIZE_DEBOUNCE_MS, self._apply_resize)

    def _apply_resize(self):
        self._resize_after_id = None
        columns = self._desired_columns()
        if columns != self._current_columns:
            self._current_columns = columns
            if self._last_results:
                self._render_results(self._last_results)

    def _thumb_for(self, path):
        """Returns a ready CTkImage for `path`, decoding it only once.

        Changing the window width re-lays out the whole grid, and re-reading
        150 JPEGs off disk each time is a visible stall - so each decoded
        thumbnail is kept and handed back on later renders.
        """
        cached = self._thumb_cache.get(path)
        if cached is not None:
            return cached
        img = Image.open(path)
        img.thumbnail(THUMB_SIZE)
        photo = ctk.CTkImage(light_image=img, dark_image=img, size=img.size)
        # Bounded so a long session over a 31k-image library cannot grow
        # without limit; dropping the whole cache is fine, it just refills.
        if len(self._thumb_cache) >= THUMB_CACHE_MAX:
            self._thumb_cache.clear()
        self._thumb_cache[path] = photo
        return photo

    # ---------- preview, clipboard ----------

    def _preview_dir(self):
        local_folder, _t, _f, _i = self._sp_cache_dirs()
        return os.path.join(local_folder, "preview")

    def _best_local_copy(self, path, meta):
        """The sharpest version of a photo already on this machine.

        Prefers a downloaded original, then the 800px preview, then the
        thumbnail the grid is drawn from - so copying or previewing never
        waits on the network when it does not have to.
        """
        item_id = (meta or {}).get("item_id")
        name = (meta or {}).get("name") or ""
        if item_id:
            _lf, _t, full_dir, _i = self._sp_cache_dirs()
            full = os.path.join(full_dir, f"{item_id}_{name}")
            if os.path.exists(full):
                return full
            large = os.path.join(self._preview_dir(), f"{item_id}.jpg")
            if os.path.exists(large):
                return large
        return path

    def _fit(self, img, box):
        """Scales an image to fill `box` as closely as its shape allows.

        PIL's thumbnail() only ever shrinks, so a 176px cached thumbnail
        stayed 176px inside a 1120px preview window - the reason the first
        version of this opened postage-stamp sized.
        """
        scale = min(box[0] / img.width, box[1] / img.height)
        size = (max(1, round(img.width * scale)), max(1, round(img.height * scale)))
        return img.resize(size, Image.LANCZOS)

    def _show_preview(self, path, meta):
        """A large view inside the app.

        Opens immediately from whatever is already on disk, then replaces
        the image with an 800px copy fetched in the background. Waiting on
        the full original (often 8MB) before showing anything made looking
        at a photo feel like a download rather than a glance.
        """
        name = (meta or {}).get("name") or os.path.basename(path)
        win = ctk.CTkToplevel(self.root)
        win.title(name)
        win.configure(fg_color=BG_COLOR)
        win.geometry("1180x840")
        try:
            win.after(200, lambda: win.iconbitmap(_resource_path("assets", "app_icon.ico")))
        except Exception:
            pass

        token = object()
        holder = ctk.CTkFrame(win, fg_color=BG_COLOR)
        holder.pack(fill=tk.BOTH, expand=True, padx=12, pady=(12, 0))
        image_label = ctk.CTkLabel(holder, text="", fg_color=BG_COLOR)
        image_label.pack(fill=tk.BOTH, expand=True)

        bar = ctk.CTkFrame(win, fg_color="transparent")
        bar.pack(fill=tk.X, padx=12, pady=12)
        caption = ctk.CTkLabel(
            bar, text=(meta or {}).get("folder") or name,
            text_color=MUTED_TEXT_COLOR, anchor="w",
        )
        caption.pack(side=tk.LEFT)
        _make_button(bar, "Close", win.destroy, width=90).pack(side=tk.RIGHT, padx=(8, 0))
        if meta and meta.get("item_id"):
            _make_button(
                bar, "Open full size",
                lambda: self._open_sp_result(meta["item_id"], meta["name"]), width=140,
            ).pack(side=tk.RIGHT, padx=(8, 0))
        _make_button(
            bar, "Copy", lambda: self._copy_to_clipboard(path, meta, token=token), width=90,
        ).pack(side=tk.RIGHT, padx=(8, 0))

        self._preview = {"token": token, "label": image_label, "win": win, "caption": caption}

        def draw(source_path):
            try:
                img = self._fit(Image.open(source_path).convert("RGB"), PREVIEW_MAX)
            except Exception:
                return
            photo = ctk.CTkImage(light_image=img, dark_image=img, size=img.size)
            image_label.configure(image=photo)
            image_label._preview_ref = photo  # keep it alive with the widget

        draw(self._best_local_copy(path, meta))

        item_id = (meta or {}).get("item_id")
        if item_id:
            if not os.path.exists(os.path.join(self._preview_dir(), f"{item_id}.jpg")):
                caption.configure(text=f"{caption.cget('text')}   (sharpening...)")
                self._fetch_preview(item_id, token)
            # Opening a preview is a deliberate act, so it is a good moment
            # to pull the original in the background - Copy and Open full
            # size are then instant instead of starting a download.
            self._prefetch_full(item_id, (meta or {}).get("name") or name)

    def _prefetch_full(self, item_id, name):
        def work():
            try:
                self._ensure_full_image(item_id, name)
            except Exception as exc:
                _log_reuse(f"FULL PREFETCH FAILED: {exc!r}")

        self._run_in_background(work)

    def _fetch_preview(self, item_id, token):
        """Pulls the 800px version of one photo in the background."""
        def work():
            try:
                if self.sp_client is None:
                    self.sp_client = SharePointClient()
                    self.sp_client.sign_in()
                dest = os.path.join(self._preview_dir(), f"{item_id}.jpg")
                ok = self.sp_client.download_thumbnail({"id": item_id}, dest, size="large")
                if ok:
                    self.event_queue.put(("preview_ready", token, dest))
            except Exception as exc:
                _log_reuse(f"PREVIEW FETCH FAILED: {exc!r}")

        self._run_in_background(work)

    def _ensure_full_image(self, item_id, name):
        """Local path to the original photo, downloading it if needed.

        Called from a worker thread - it can block on an 8MB download.
        """
        _lf, _t, full_dir, _i = self._sp_cache_dirs()
        dest = os.path.join(full_dir, f"{item_id}_{name}")
        if os.path.exists(dest) and os.path.getsize(dest) > 0:
            return dest
        if self.sp_client is None:
            self.sp_client = SharePointClient()
            self.sp_client.sign_in()
        self.sp_client.download_file({"id": item_id}, dest)
        return dest

    def _put_image_on_clipboard(self, source):
        """Writes one image to the Windows clipboard as a device independent
        bitmap - the format Office pastes as a picture rather than as a
        file reference. Returns its pixel size."""
        import ctypes
        from io import BytesIO

        img = Image.open(source).convert("RGB")
        buf = BytesIO()
        img.save(buf, "BMP")
        data = buf.getvalue()[14:]  # a DIB is a BMP minus its file header
        buf.close()

        kernel32, user32 = ctypes.windll.kernel32, ctypes.windll.user32
        # Handles are pointer-sized; without these the values are truncated
        # on 64-bit Python and the clipboard write corrupts.
        kernel32.GlobalAlloc.restype = ctypes.c_void_p
        kernel32.GlobalAlloc.argtypes = [ctypes.c_uint, ctypes.c_size_t]
        kernel32.GlobalLock.restype = ctypes.c_void_p
        kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
        kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
        user32.SetClipboardData.restype = ctypes.c_void_p
        user32.SetClipboardData.argtypes = [ctypes.c_uint, ctypes.c_void_p]

        handle = kernel32.GlobalAlloc(0x0002, len(data))  # GMEM_MOVEABLE
        pointer = kernel32.GlobalLock(handle)
        ctypes.memmove(pointer, data, len(data))
        kernel32.GlobalUnlock(handle)
        if not user32.OpenClipboard(0):
            raise OSError("clipboard busy")
        try:
            user32.EmptyClipboard()
            user32.SetClipboardData(8, handle)  # CF_DIB
        finally:
            user32.CloseClipboard()
        return img.width, img.height

    def _copy_to_clipboard(self, path, meta, token=None):
        """Copies the photo at full resolution.

        Always the original, never the thumbnail the grid is drawn from -
        a 176px copy pasted into a deck is useless. If the original is not
        on this machine yet it is fetched first, so this can take a moment
        the first time and is instant afterwards.
        """
        item_id = (meta or {}).get("item_id")
        name = (meta or {}).get("name") or os.path.basename(path)

        def report(msg):
            self.event_queue.put(("clipboard_status", token, msg))

        def work():
            try:
                source = path
                if item_id:
                    cached = os.path.join(self._sp_cache_dirs()[2], f"{item_id}_{name}")
                    if not os.path.exists(cached):
                        report("Fetching the full-resolution photo...")
                    source = self._ensure_full_image(item_id, name)
                width, height = self._put_image_on_clipboard(source)
                report(f"Copied at full resolution - {width} x {height}.")
            except Exception as exc:
                _log_reuse(f"CLIPBOARD COPY FAILED: {exc!r}")
                report("Could not copy this photo to the clipboard.")

        self._run_in_background(work)

    # ---------- manual collections ----------

    def _collection_ids(self):
        """{name: [item_id, ...]} exactly as stored on disk.

        Item ids rather than positions, so a collection survives the index
        being rebuilt in a different order after a sync.
        """
        data = self._read_json(self._collections_cache_path()) or {}
        out = {name: [] for name in MANUAL_COLLECTIONS}
        for name, ids in data.items():
            if isinstance(ids, list):
                out[name] = [i for i in ids if i]
        return out

    def _write_collection_ids(self, by_id):
        try:
            os.makedirs(os.path.dirname(self._collections_cache_path()), exist_ok=True)
            with open(self._collections_cache_path(), "w", encoding="utf-8") as f:
                json.dump(by_id, f, ensure_ascii=False)
            return True
        except OSError as exc:
            _log_reuse(f"COLLECTIONS SAVE FAILED: {exc!r}")
            self.status_var.set("Could not save that collection - see crash.log.")
            return False

    def _collection_names(self):
        return list((self._collections or {}).keys()) or list(MANUAL_COLLECTIONS)

    def _add_to_collection(self, name, path, meta):
        item_id = (meta or {}).get("item_id")
        if not item_id:
            self.status_var.set("Only photos from the shared drive can be collected.")
            return
        by_id = self._collection_ids()
        ids = by_id.setdefault(name, [])
        if item_id in ids:
            self.status_var.set(f"Already in {name}.")
            return
        ids.append(item_id)
        if not self._write_collection_ids(by_id):
            return
        self._collections = self._load_collections()
        self._rebuild_collection_cards()
        self.status_var.set(f"Added to {name} - {len(ids)} photo(s).")
        if self._view == name:
            self._show_collection(name)

    def _remove_from_collection(self, name, meta):
        item_id = (meta or {}).get("item_id")
        by_id = self._collection_ids()
        ids = by_id.get(name) or []
        if not item_id or item_id not in ids:
            return
        ids.remove(item_id)
        if not self._write_collection_ids(by_id):
            return
        self._collections = self._load_collections()
        self._rebuild_collection_cards()
        self.status_var.set(f"Removed from {name} - {len(ids)} photo(s) left.")
        if self._view == name:
            self._show_collection(name)

    def _show_result_menu(self, event, path, meta):
        """Right-click menu on a result: the keyboard-and-mouse route into a
        collection, for anyone who would rather not drag."""
        menu = tk.Menu(self.root, tearoff=0)
        for name in self._collection_names():
            menu.add_command(
                label=f"Add to {name}",
                command=lambda n=name: self._add_to_collection(n, path, meta),
            )
        if self._view in self._collection_names():
            menu.add_separator()
            menu.add_command(
                label=f"Remove from {self._view}",
                command=lambda: self._remove_from_collection(self._view, meta),
            )
        menu.add_separator()
        menu.add_command(label="Preview", command=lambda: self._show_preview(path, meta))
        menu.add_command(label="Copy image", command=lambda: self._copy_to_clipboard(path, meta))
        if meta and meta.get("item_id"):
            menu.add_command(
                label="Open full size",
                command=lambda: self._open_sp_result(meta["item_id"], meta["name"]),
            )
        else:
            menu.add_command(label="Open full size", command=lambda: open_in_system_viewer(path))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    # ---------- drag a photo onto a collection ----------

    def _drag_start(self, event, path, meta):
        # Only records where the press happened. The drag does not begin
        # until the pointer actually moves, so a double-click still opens
        # the photo instead of being swallowed here.
        self._drag = {
            "path": path, "meta": meta,
            "x": event.x_root, "y": event.y_root,
            "active": False, "win": None, "hover": None,
        }

    def _card_under_pointer(self, x_root, y_root):
        for name, card in (self.collection_cards or {}).items():
            try:
                cx, cy = card.winfo_rootx(), card.winfo_rooty()
                if (cx <= x_root <= cx + card.winfo_width()
                        and cy <= y_root <= cy + card.winfo_height()):
                    return name
            except Exception:
                continue
        return None

    def _set_drag_hover(self, name):
        drag = self._drag
        if drag is None or drag.get("hover") == name:
            return
        previous = drag.get("hover")
        if previous and previous in (self.collection_cards or {}):
            self.collection_cards[previous].configure(
                fg_color=ACCENT_COLOR if self._view == previous else CARD_COLOR
            )
        if name and name in (self.collection_cards or {}):
            self.collection_cards[name].configure(fg_color=ACCENT_ACTIVE)
        drag["hover"] = name

    def _drag_motion(self, event):
        drag = self._drag
        if drag is None:
            return
        if not drag["active"]:
            moved = abs(event.x_root - drag["x"]) + abs(event.y_root - drag["y"])
            if moved < DRAG_THRESHOLD_PX:
                return
            drag["active"] = True
            drag["win"] = self._make_drag_preview(drag["path"])
        win = drag.get("win")
        if win is not None:
            try:
                win.geometry(f"+{event.x_root + 12}+{event.y_root + 12}")
            except Exception:
                pass
        self._set_drag_hover(self._card_under_pointer(event.x_root, event.y_root))

    def _make_drag_preview(self, path):
        """A small translucent window that follows the pointer, so it is
        obvious something is being carried."""
        try:
            win = tk.Toplevel(self.root)
            win.overrideredirect(True)
            try:
                win.attributes("-alpha", 0.85)
                win.attributes("-topmost", True)
            except Exception:
                pass
            try:
                ctk.CTkLabel(win, image=self._thumb_for(path), text="").pack()
            except Exception:
                tk.Label(win, text="Moving photo...", padx=10, pady=6).pack()
            return win
        except Exception:
            return None

    def _drag_release(self, event):
        drag = self._drag
        self._drag = None
        if drag is None:
            return
        win = drag.get("win")
        if win is not None:
            try:
                win.destroy()
            except Exception:
                pass
        hover = drag.get("hover")
        if hover and hover in (self.collection_cards or {}):
            self.collection_cards[hover].configure(
                fg_color=ACCENT_COLOR if self._view == hover else CARD_COLOR
            )
        if not drag["active"]:
            return  # a plain click, not a drag
        target = self._card_under_pointer(event.x_root, event.y_root)
        if target:
            self._add_to_collection(target, drag["path"], drag["meta"])

    def _render_results(self, results):
        self._last_results = results
        # A fill from the previous render may still be in flight.
        if self._render_after_id is not None:
            try:
                self.root.after_cancel(self._render_after_id)
            except Exception:
                pass
            self._render_after_id = None
        for child in self.results_frame.winfo_children():
            child.destroy()
        self.thumbnail_refs.clear()
        self._render_chunk(results, 0, max(1, self._current_columns))

    def _render_chunk(self, results, start, columns):
        """Builds one screenful of results, then yields to the event loop
        before the next.

        A full grid is 150 cells, each of which is several Tk widgets plus a
        JPEG decode - building them in a single pass froze the window for
        seconds after every search, with nothing on screen until the last
        one was done. Rendering in chunks puts the top of the grid up
        immediately and fills the rest in behind it, so a search feels
        instant even though the same total work still happens.
        """
        self._render_after_id = None
        end = min(start + RENDER_CHUNK, len(results))

        for idx in range(start, end):
            path, _score, meta = results[idx]
            row, col = divmod(idx, columns)
            cell = ctk.CTkFrame(self.results_frame, fg_color="transparent")
            cell.grid(row=row, column=col, sticky="n", padx=4, pady=4)

            try:
                photo = self._thumb_for(path)
            except Exception:
                continue
            self.thumbnail_refs.append(photo)

            holder = ctk.CTkFrame(cell, fg_color="transparent")
            holder.pack()
            label = ctk.CTkLabel(holder, image=photo, text="", cursor="hand2")
            label.pack()
            # Double-click previews in-app. Opening the original meant an
            # 8MB download before anything appeared; the full file is still
            # one click away from inside the preview.
            label.bind(
                "<Double-Button-1>",
                lambda _e, p=path, m=meta: self._show_preview(p, m),
            )
            label.bind("<Button-3>", lambda e, p=path, m=meta: self._show_result_menu(e, p, m))
            label.bind("<ButtonPress-1>", lambda e, p=path, m=meta: self._drag_start(e, p, m))
            label.bind("<B1-Motion>", self._drag_motion)
            label.bind("<ButtonRelease-1>", self._drag_release)

            key = self._fav_key(path, meta)
            star = ctk.CTkLabel(
                holder,
                text="\u2605" if self._is_favourite(key) else "\u2606",
                text_color="#ffd54a" if self._is_favourite(key) else "#ffffff",
                font=ctk.CTkFont(size=20),
                fg_color=CARD_COLOR,
                corner_radius=8,
                width=26,
                height=24,
                cursor="hand2",
            )
            star.place(relx=1.0, rely=0.0, anchor="ne", x=-3, y=3)
            star.bind(
                "<Button-1>",
                lambda _e, k=key, p=path, m=meta, w=star: self._on_star_clicked(k, p, m, w),
            )

            # The SharePoint filename is almost always a camera default
            # (DSC05540.jpg), so the folder is the only human-meaningful
            # label a result can carry - show it when we know it.
            folder = (meta or {}).get("folder") if meta else None
            caption = folder if folder else os.path.basename(path)
            ctk.CTkLabel(
                cell, text=caption, justify=tk.CENTER, wraplength=THUMB_SIZE[0],
                text_color=MUTED_TEXT_COLOR, font=ctk.CTkFont(size=12),
            ).pack()

        if end < len(results):
            self._render_after_id = self.root.after(
                1, lambda: self._render_chunk(results, end, columns)
            )

    def _on_star_clicked(self, key, path, meta, widget):
        added = self._toggle_favourite(key, path, meta)
        widget.configure(
            text="★" if added else "☆",
            text_color="#ffd54a" if added else "#ffffff",
        )
        # Un-starring while viewing favourites should drop it from the list.
        if self._view == "favourites" and not added:
            self._show_favourites()


def _thread_excepthook(args):
    _log_crash("unhandled thread exception", (args.exc_type, args.exc_value, args.exc_traceback))


def main():
    threading.excepthook = _thread_excepthook
    # Windows groups a window under the executable that started it, so a
    # script run through pythonw shows the Python icon in the taskbar and in
    # Alt-Tab. Claiming an explicit identity makes it use this app's own icon.
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "Greenreb.ImageSearch"
        )
    except Exception:
        pass
    ctk.set_appearance_mode("dark")
    # CustomTkinter re-scales every widget when the window moves to a monitor
    # with a different display scaling, which on a multi-monitor desk means
    # the whole grid is rebuilt mid-drag and the window judders. The app is
    # laid out in plain pixels and does not need that, so it is turned off.
    try:
        ctk.deactivate_automatic_dpi_awareness()
    except Exception:
        pass
    root = ctk.CTk()
    try:
        root.configure(fg_color=BG_COLOR)
    except Exception:
        pass
    try:
        # default=... applies the icon to every window the app opens (the
        # preview and settings windows too), not just the main one.
        root.iconbitmap(default=_resource_path("assets", "app_icon.ico"))
    except Exception:
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
