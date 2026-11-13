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
import io
import json
import math
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
import urllib.parse
import uuid
import zipfile
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

_IMPORT_T0 = time.perf_counter()
import customtkinter as ctk
_T_CTK = time.perf_counter()
import numpy as np
_T_NUMPY = time.perf_counter()
from PIL import Image, ImageDraw, ImageFilter, ImageFont, ImageOps, ImageTk
_T_PIL = time.perf_counter()

from search_engine import ImageSearchEngine, MODEL_TAG, to_english
from app_version import (
    APP_AUTHOR, APP_MODULES, APP_NAME, APP_VERSION, RUNTIME_VERSION, WHATS_NEW, version_tuple,
)
_T_ENGINE = time.perf_counter()
_SHAREPOINT = None


def _sharepoint():
    """The SharePoint module, imported the first time it is needed.

    It pulls in msal, requests and cryptography - most of a second of import
    time that used to sit in front of the first window, although starting up
    only reads the photo library cache from disk and never talks to
    SharePoint.
    """
    global _SHAREPOINT
    if _SHAREPOINT is None:
        import sharepoint_client
        _SHAREPOINT = sharepoint_client
    return _SHAREPOINT

IS_MAC = sys.platform == "darwin"


def _data_dir():
    """Settings, caches, logs and downloaded updates: %LOCALAPPDATA%\\ImageSearch
    on Windows, ~/Library/Application Support/Topshot on a Mac."""
    if IS_MAC:
        return os.path.join(os.path.expanduser("~"), "Library", "Application Support", "Topshot")
    return os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "ImageSearch")


DATA_DIR = _data_dir()

CRASH_LOG_PATH = os.path.join(
    DATA_DIR, "crash.log"
)
BOOT_LOG_PATH = os.path.join(
    DATA_DIR, "startup.log"
)
REUSE_LOG_PATH = os.path.join(
    DATA_DIR, "reuse_debug.log"
)


_BOOT_T0 = _IMPORT_T0
_BOOT_LOGGED = False


def _process_age():
    """Seconds since Windows created this process. Includes the time before
    Python was running at all, which a timer started in this file can't see."""
    try:
        import ctypes
        from ctypes import wintypes

        kernel32 = ctypes.windll.kernel32
        kernel32.GetCurrentProcess.restype = wintypes.HANDLE
        kernel32.GetProcessTimes.argtypes = (
            [wintypes.HANDLE] + [ctypes.POINTER(wintypes.FILETIME)] * 4
        )
        times = [wintypes.FILETIME() for _ in range(4)]
        if not kernel32.GetProcessTimes(
            kernel32.GetCurrentProcess(), *[ctypes.byref(t) for t in times]
        ):
            return None
        now = wintypes.FILETIME()
        kernel32.GetSystemTimeAsFileTime(ctypes.byref(now))

        def ticks(t):
            return (t.dwHighDateTime << 32) | t.dwLowDateTime

        return (ticks(now) - ticks(times[0])) / 1e7
    except Exception:
        return None


def _boot_mark(label):
    """Appends one line to startup.log: how long after launch `label` happened.

    So "it starts slowly" can be answered with numbers - which phase took
    the time - rather than by feel.
    """
    global _BOOT_LOGGED
    try:
        os.makedirs(os.path.dirname(BOOT_LOG_PATH), exist_ok=True)
        if os.path.exists(BOOT_LOG_PATH) and os.path.getsize(BOOT_LOG_PATH) > 256_000:
            os.remove(BOOT_LOG_PATH)
        age = _process_age()
        since_launch = f"   ({age:5.2f}s since launch)" if age is not None else ""
        with open(BOOT_LOG_PATH, "a", encoding="utf-8") as f:
            if not _BOOT_LOGGED:
                stamp = datetime.datetime.now().isoformat(timespec="seconds")
                f.write(f"--- {stamp}\n")
                _BOOT_LOGGED = True
            f.write(f"+{time.perf_counter() - _BOOT_T0:5.2f}s{since_launch}  {label}\n")
    except OSError:
        pass


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
# The collections a new install starts with. They are suggestions, not
# fixtures: people can add their own and remove any of these.
DEFAULT_COLLECTIONS = ["Golf", "Food", "Drinks", "Family & Kids", "Groups"]
COLLECTION_NAME_MAX = 32
# Collection names double as view keys (see _set_view), so these would
# collide with the built-in views.
RESERVED_COLLECTION_NAMES = {"all", "all photos", "favourites", "history", "search history"}
# Collections that existed under an earlier name. Their photos are folded
# into the new collection when read, so renaming never loses anything.
LEGACY_COLLECTIONS = {"Family": "Family & Kids", "Kids": "Family & Kids"}
# How far the pointer has to move with the button held before it counts as a
# drag rather than a click, so double-clicking a photo still opens it.
DRAG_THRESHOLD_PX = 8
# Shared collections are re-read from SharePoint this often while the app
# is open (and on start, and when Sync is clicked).
SHARED_SYNC_MS = 10 * 60 * 1000
# How often an open app looks for a newer version (also on every start).
UPDATE_CHECK_MS = 4 * 60 * 60 * 1000
# What an update package may contain. Anything else means the package is
# not one of ours, and it is refused.
UPDATE_FILE_TYPES = {".py", ".ttf", ".css", ".txt", ".json", ".png", ".ico"}
EMAIL_PATTERN = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
# Drag and drop motion. The carried photo eases towards where it should be
# every frame (an exponential approach: reads as ease-out, and retargets
# smoothly when the pointer changes its mind) instead of jumping.
DRAG_SCALE = 0.62        # carried size, as a share of the tile
DRAG_SCALE_OVER = 0.5    # a little smaller over a collection: "it goes in here"
DRAG_ALPHA = 0.92
DRAG_FOLLOW_TAU_MS = 26  # how closely it trails the pointer
DRAG_SETTLE_TAU_MS = 60  # how quickly it shrinks / grows
DRAG_HOVER_MS = 140      # collection card lighting up or dimming
DROP_FLY_MS = 280        # flying into the collection
DROP_BACK_MS = 240       # gliding home when dropped elsewhere
CARD_PULSE_MS = 560      # the collection's confirmation glow
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
# For the People / No people filter: is anyone in the photo at all - a
# single golfer counts - rather than "guests having fun", which is what
# ranking uses.
PERSON_PROMPTS = [
    "a photo of a person",
    "a photo of people",
    "a man or a woman in the picture",
    "a golfer swinging a club",
    "a portrait of someone smiling",
]
NOBODY_PROMPTS = [
    "a photo with nobody in it",
    "a close up of food on a plate",
    "a drink on a table",
    "an empty room",
    "an empty golf range",
    "the outside of a building",
]
OBJECT_PROMPTS = [
    "a close up photograph of food on a plate",
    "a drink in a glass on a bar",
    "a product shot of a single object",
    "a detail photograph of equipment",
]
DEFAULT_TOP_K = 150

BG_COLOR = "#090b0a"
CARD_COLOR = "#141917"
ACCENT_COLOR = "#1d2621"
ACCENT_ACTIVE = "#1b4a37"
FIELD_BG = "#141917"
TEXT_COLOR = "#fafafa"
MUTED_TEXT_COLOR = "#9ca8a2"

CARD_RADIUS = 14

# Topshot's look: the layout of the Topgolf Sales Cockpit (near-black page,
# a slightly lighter full-height sidebar, cards with hairline borders, white
# text) tinted with the greens of the app icon. One green does the
# emphasis: the Search button, the selected item, focus and progress.
# Inter (bundled in assets/fonts, loaded for this process only - nothing is
# installed on the machine). Falls back to Segoe UI if the files are missing.
UI_FONT = "Segoe UI"
UI_FONT_MEDIUM = "Segoe UI Semibold"
UI_FONT_SEMIBOLD = "Segoe UI Semibold"
if sys.platform == "darwin":
    UI_FONT = UI_FONT_MEDIUM = UI_FONT_SEMIBOLD = "Helvetica Neue"
_INTER_FILES = ("Inter-Regular.ttf", "Inter-Medium.ttf", "Inter-SemiBold.ttf", "Inter-Bold.ttf")


def _register_bundled_fonts():
    """Makes the bundled Inter files usable by this process (FR_PRIVATE):
    nothing is installed, and they vanish when the app closes."""
    if IS_MAC:
        return _register_fonts_mac()
    try:
        import ctypes

        add = ctypes.windll.gdi32.AddFontResourceExW
    except Exception:
        return 0
    added = 0
    for name in _INTER_FILES:
        path = _resource_path("assets", "fonts", name)
        if os.path.exists(path):
            try:
                added += add(path, 0x10, 0)
            except Exception:
                pass
    return added


def _register_fonts_mac():
    """The Mac version of the above: Core Text, for this process only."""
    try:
        import ctypes

        cf = ctypes.cdll.LoadLibrary(
            "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation")
        ct = ctypes.cdll.LoadLibrary("/System/Library/Frameworks/CoreText.framework/CoreText")
        cf.CFURLCreateFromFileSystemRepresentation.restype = ctypes.c_void_p
        cf.CFURLCreateFromFileSystemRepresentation.argtypes = [
            ctypes.c_void_p, ctypes.c_char_p, ctypes.c_long, ctypes.c_bool]
        cf.CFRelease.argtypes = [ctypes.c_void_p]
        ct.CTFontManagerRegisterFontsForURL.restype = ctypes.c_bool
        ct.CTFontManagerRegisterFontsForURL.argtypes = [
            ctypes.c_void_p, ctypes.c_uint32, ctypes.c_void_p]
    except Exception:
        return 0
    added = 0
    for name in _INTER_FILES:
        path = _resource_path("assets", "fonts", name)
        if not os.path.exists(path):
            continue
        raw = os.fsencode(path)
        url = cf.CFURLCreateFromFileSystemRepresentation(None, raw, len(raw), False)
        if not url:
            continue
        try:
            added += bool(ct.CTFontManagerRegisterFontsForURL(url, 1, None))  # process scope
        except Exception:
            pass
        finally:
            cf.CFRelease(url)
    return added


def _choose_ui_fonts(root):
    """Uses Inter where Windows can see it, Segoe UI otherwise."""
    global UI_FONT, UI_FONT_MEDIUM, UI_FONT_SEMIBOLD
    try:
        import tkinter.font as tkfont

        families = set(tkfont.families(root))
    except Exception:
        return
    if "Inter" in families:
        UI_FONT = "Inter"
        UI_FONT_MEDIUM = "Inter Medium" if "Inter Medium" in families else "Inter"
        UI_FONT_SEMIBOLD = "Inter SemiBold" if "Inter SemiBold" in families else "Inter"


def _font(size, weight="regular"):
    """The app's type scale in one place: Inter Regular / Medium / SemiBold,
    falling back to Segoe UI and Segoe UI Semibold."""
    family = {
        "regular": UI_FONT, "medium": UI_FONT_MEDIUM,
        "semibold": UI_FONT_SEMIBOLD, "bold": UI_FONT,
    }[weight]
    bold = weight == "bold" or (weight != "regular" and family == UI_FONT)
    return ctk.CTkFont(family=family, size=size, weight="bold" if bold else "normal")
SURFACE = "#141917"
SURFACE_HOVER = "#1c2420"
BORDER = "#24302a"
TEXT_SOFT = "#e5e5e5"
FAINT_TEXT = "#6b7a72"
SIDEBAR_BG = "#0f1311"
NAV_BG = SIDEBAR_BG
PRIMARY_TEXT = "#ffffff"
FOCUS_BORDER = "#10a573"
PULSE_COLOUR = "#1f6b4c"
SOURCE_DOT = "#34d399"
NAV_HOVER = "#17201b"
NAV_SELECTED = "#123628"
SYSTEM_TEXT = "#84928b"  # System items sit back from the rest of the sidebar
MENU_BORDER = "#2a3830"
DANGER = "#f87171"
PRIMARY = "#10a573"
PRIMARY_HOVER = "#14c088"

# Segoe MDL2 Assets ships with Windows 10 and 11; the glyphs are drawn into
# images so they sit beside Segoe UI text at any size. Without the font the
# buttons simply have no icon.
_ICON_GLYPHS = {
    "photos": "\uE91B", "star": "\uE734", "history": "\uE81C", "folder": "\uE8B7",
    "search": "\uE721", "refresh": "\uE72C", "settings": "\uE713", "add": "\uE710",
    "similar": "\uE8B9", "open": "\uE838",
    "shared": "\uE716", "update": "\uE896", "sync": "\uE72C", "info": "\uE946",
    "support": "\uE897", "local": "\uEDA2",
    "back": "\uE72B", "next": "\uE72A", "done": "\uE73E",
    "preview": "\uE7B3", "copy": "\uE8C8", "external": "\uE8A7", "newfolder": "\uE8F4",
    "move": "\uE8DE", "remove": "\uE738", "trash": "\uE74D", "edit": "\uE70F",
    "whatsnew": "\uE734",
}
_icon_cache = {}
# The same outline icon set as the Sales Cockpit (Lucide), from the bundled
# icon font; its codepoints are read from the stylesheet that ships with it.
_LUCIDE_NAMES = {
    "photos": "images", "star": "star", "history": "history", "folder": "folder",
    "search": "search", "refresh": "refresh-cw", "settings": "settings", "add": "plus",
    "similar": "scan-search", "open": "folder-open", "index": "database",
    "shared": "users", "update": "download", "sync": "refresh-cw", "info": "info",
    "support": "life-buoy", "local": "hard-drive",
    "back": "arrow-left", "next": "arrow-right", "done": "check",
    "preview": "eye", "copy": "copy", "external": "external-link", "newfolder": "folder-plus",
    "move": "folder-input", "remove": "folder-minus", "trash": "trash-2", "edit": "pencil",
    "whatsnew": "sparkles", "more": "ellipsis", "save": "download", "close": "x",
    "selected": "check", "landscape": "rectangle-horizontal", "portrait": "rectangle-vertical",
    "people": "users", "nopeople": "user-x", "filter": "filter", "tour": "mouse-pointer-click",
}
_lucide_codes = None


def _lucide_glyph(name):
    global _lucide_codes
    if _lucide_codes is None:
        _lucide_codes = {}
        try:
            with open(_resource_path("assets", "fonts", "lucide.css"), encoding="utf-8") as f:
                css = f.read()
            for icon, code in re.findall(
                r'\.icon-([a-z0-9-]+):+before\s*\{\s*content:\s*"\\([0-9a-fA-F]+)"', css
            ):
                _lucide_codes[icon] = chr(int(code, 16))
        except Exception:
            pass
    return _lucide_codes.get(_LUCIDE_NAMES.get(name, name))


_icon_pil_cache = {}


def _icon_tk(name, size=16, colour=None):
    """The same icon as a plain Tk image, or None."""
    if _icon(name, size, colour) is None:
        return None
    small = _icon_pil_cache.get((name, size, colour))
    return ImageTk.PhotoImage(small) if small is not None else None


def _icon(name, size=16, colour=None):
    key = (name, size, colour)
    if key in _icon_cache:
        return _icon_cache[key]
    image = None
    try:
        k = 4
        glyph = _lucide_glyph(name)
        lucide_path = _resource_path("assets", "fonts", "lucide.ttf")
        if glyph and os.path.exists(lucide_path):
            font = ImageFont.truetype(lucide_path, size * k)
        else:
            glyph = _ICON_GLYPHS[name]
            font = ImageFont.truetype(
                os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "segmdl2.ttf"), size * k
            )
        canvas = Image.new("RGBA", (size * k, size * k), (0, 0, 0, 0))
        draw = ImageDraw.Draw(canvas)
        left, top, right, bottom = draw.textbbox((0, 0), glyph, font=font)
        draw.text(
            ((size * k - (right - left)) / 2 - left, (size * k - (bottom - top)) / 2 - top),
            glyph, font=font, fill=_hex_rgb(colour or TEXT_SOFT) + (255,),
        )
        small = canvas.resize((size, size), Image.LANCZOS)
        _icon_pil_cache[key] = small
        image = ctk.CTkImage(light_image=small, dark_image=small, size=(size, size))
    except Exception:
        image = None
    _icon_cache[key] = image
    return image
BUTTON_RADIUS = 17

LOGO_HEIGHT = 36


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
    # A downloaded update runs from its own folder and brings its own
    # assets, so this file's folder comes first in every case.
    here = os.path.join(os.path.dirname(os.path.abspath(__file__)), *parts)
    if os.path.exists(here) or not getattr(sys, "frozen", False):
        return here
    return os.path.join(sys._MEIPASS, *parts)


def _load_mark(height):
    """The app's mark, `height` px tall: the Topshot icon when it is there,
    otherwise the shield cut from the old logo."""
    mark = _resource_path("assets", "topshot_mark.png")
    if os.path.exists(mark):
        with Image.open(mark) as img:
            return img.convert("RGBA").resize((height, height), Image.LANCZOS)
    return _load_shield(height)


def _load_shield(height):
    """The Topgolf shield on its own, as an RGBA image `height` px tall.

    logo.png is the full lockup - the shield above the TOPGOLF wordmark -
    inside a wide transparent margin. Drawn at header size the wordmark came
    out about 7px tall: unreadable, and it made the whole mark look low
    quality. The shield is the part that survives small sizes, so it is cut
    out at the empty band separating it from the wordmark.
    """
    img = Image.open(_resource_path("assets", "logo.png")).convert("RGBA")
    box = img.getbbox()
    if box:
        img = img.crop(box)
    filled = (np.asarray(img)[:, :, 3] > 8).any(axis=1)
    widest, start = None, None
    for row, has_ink in enumerate(filled):
        if not has_ink and start is None:
            start = row
        elif has_ink and start is not None:
            if widest is None or row - start > widest[1] - widest[0]:
                widest = (start, row)
            start = None
    # Only treat it as the shield/wordmark gap if it sits in the lower part;
    # anything else means the artwork has changed shape, so keep it whole.
    if widest is not None and widest[0] > img.height * 0.4:
        img = img.crop((0, 0, img.width, widest[0]))
        box = img.getbbox()
        if box:
            img = img.crop(box)
    width = max(1, round(img.width * height / img.height))
    return img.resize((width, height), Image.LANCZOS)


def _load_logo_image(height=LOGO_HEIGHT):
    try:
        img = _load_mark(height)
        return ctk.CTkImage(light_image=img, dark_image=img, size=img.size)
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


def _open_external(target):
    """Opens a link, a mailto: or a file with whatever the system uses for it.
    Raises if nothing could be opened, so callers can offer another way."""
    if sys.platform.startswith("win"):
        os.startfile(target)  # noqa: S606
    elif sys.platform == "darwin":
        subprocess.run(["open", target], check=True, timeout=20)
    else:
        subprocess.run(["xdg-open", target], check=True, timeout=20)


def _apply_icon(win):
    """The app icon on a window. A Mac shows the app bundle's icon instead."""
    if sys.platform == "darwin":
        return
    try:
        win.iconbitmap(_resource_path("assets", "app_icon.ico"))
    except Exception:
        pass


def _bind_context_menu(widget, handler):
    """Right-click. On a Mac a two-finger click or the right button arrives as
    Button-2 under Tk 8.6 (Button-3 from Tk 8.7 on), so both are bound."""
    widget.bind("<Button-3>", handler)
    if sys.platform == "darwin":
        widget.bind("<Button-2>", handler)


def _system_description():
    if sys.platform == "darwin":
        return f"macOS {platform.mac_ver()[0]} ({platform.machine()})"
    return f"Windows {platform.release()} ({platform.version()})"


_selection_badge_cache = []


def _selection_badge(d=22):
    """A green disc with a white tick, drawn once."""
    if _selection_badge_cache:
        return _selection_badge_cache[0]
    k = 4
    big = Image.new("RGBA", (d * k, d * k), (0, 0, 0, 0))
    draw = ImageDraw.Draw(big)
    draw.ellipse((0, 0, d * k - 1, d * k - 1), fill=_hex_rgb(PRIMARY) + (255,),
                 outline=(255, 255, 255, 235), width=2 * k)
    draw.line([(d * k * 0.28, d * k * 0.52), (d * k * 0.44, d * k * 0.68), (d * k * 0.73, d * k * 0.36)],
              fill=(255, 255, 255, 255), width=int(2.4 * k), joint="curve")
    badge = big.resize((d, d), Image.LANCZOS)
    _selection_badge_cache.append(badge)
    return badge


def _make_ghost_button(parent, text, command, width=0, icon=None):
    """A quiet, borderless button for occasional actions."""
    return ctk.CTkButton(
        parent, text=text, command=command, width=width, height=32, corner_radius=8,
        fg_color="transparent", hover_color=SURFACE_HOVER, text_color=TEXT_SOFT,
        text_color_disabled=FAINT_TEXT, image=_icon(icon, 15) if icon else None,
        compound="left", font=_font(13, "medium"),
    )


def _make_secondary_button(parent, text, command, width=120, icon=None):
    """An outlined button: clearly clickable, but quieter than the primary."""
    return ctk.CTkButton(
        parent, text=text, command=command, width=width, height=38, corner_radius=10,
        fg_color="transparent", hover_color=SURFACE_HOVER, border_width=1,
        border_color=BORDER, text_color=TEXT_COLOR, text_color_disabled=FAINT_TEXT,
        image=_icon(icon, 15) if icon else None, compound="left",
        font=_font(13, "semibold"),
    )


def _make_button(parent, text, command, width=140):
    return ctk.CTkButton(
        parent,
        text=text,
        command=command,
        width=width,
        height=34,
        corner_radius=8,
        fg_color="#1d2621",
        hover_color="#2a3830",
        text_color_disabled=FAINT_TEXT,
        text_color=TEXT_COLOR,
        font=_font(13, "medium"),
    )


def _make_primary_button(parent, text, command, width=96):
    """The one emphasised action: white on near-black, as in the Cockpit."""
    return ctk.CTkButton(
        parent, text=text, command=command, width=width, height=36, corner_radius=8,
        fg_color=PRIMARY, hover_color=PRIMARY_HOVER, text_color=PRIMARY_TEXT,
        text_color_disabled="#a7d8c4", font=_font(13, "semibold"),
    )


STAR_BADGE_SIZE = (26, 24)
STAR_BADGE_INSET = 6  # clear of the tile's rounded corner
# The results grid is a justified layout (as in Google Photos): every
# photo in a row shares one height and keeps its own shape, and each full
# row is sized to end exactly at the right edge. No captions - the folder
# name is shown in the preview, where it is actually readable.
ROW_HEIGHT = 140  # target height; full rows shrink a little to fit the width
TILE_GAP = 6
TILE_RADIUS = 8
GRID_SIDE_PAD = 0  # tiles line up with the search field's left edge
ROWS_PER_CHUNK = 3  # rows built per event-loop turn, so the grid fills in smoothly
CAPTION_FONT = ("Roboto", -12)  # what CustomTkinter used for captions on Windows


def _hex_rgb(colour):
    return tuple(int(colour[i:i + 2], 16) for i in (1, 3, 5))


def _draw_star_badge(filled):
    """The favourite star on a small rounded badge, drawn at 8x and scaled
    down so its edges are smooth. Drawn rather than typed, so it needs no
    font and looks the same on every machine."""
    w, h = STAR_BADGE_SIZE
    k = 8
    big = Image.new("RGBA", (w * k, h * k), (0, 0, 0, 0))
    draw = ImageDraw.Draw(big)
    draw.rounded_rectangle(
        (0, 0, w * k - 1, h * k - 1), radius=8 * k, fill=_hex_rgb(CARD_COLOR) + (230,)
    )
    cx, cy = w * k / 2, h * k / 2 + 0.4 * k
    outer = 8.0 * k
    inner = outer * 0.42
    points = []
    for i in range(10):
        angle = -math.pi / 2 + i * math.pi / 5
        radius = outer if i % 2 == 0 else inner
        points.append((cx + radius * math.cos(angle), cy + radius * math.sin(angle)))
    if filled:
        draw.polygon(points, fill=(255, 213, 74, 255))
    else:
        try:
            draw.polygon(points, outline=(255, 255, 255, 255), width=round(1.4 * k))
        except TypeError:  # older Pillow: no outline width on polygons
            draw.line(points + [points[0]], fill=(255, 255, 255, 255), width=round(1.4 * k),
                      joint="curve")
    return big.resize((w, h), Image.LANCZOS)


class ImageSearchApp:
    def __init__(self, root, boot=None):
        self.root = root
        self.root.title(APP_NAME)
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
        self._thumb_cache = {}  # path -> decoded PIL thumbnail, see _thumb_image
        self._badges = None
        self._shared = None            # shared collections, see _shared_state
        self._shared_syncing = False
        self._shared_lock = threading.Lock()  # one SharePoint write at a time
        self._sp_root_id = None
        self._update_checking = False
        self._update_ready = None      # version downloaded and waiting for a restart
        self._tile_corners = None
        self._render_favs = set()
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
        self._sp_lock = threading.Lock()
        # Set by BootSequence while the app starts behind its splash.
        self._boot = boot
        self._render_done_callback = None
        self._warming = False
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
        # Laid out like the Sales Cockpit: a full-height sidebar for where you
        # are (library, collections) and the occasional system actions at
        # its foot, and a main column with the page title, the search field
        # and the photos.
        shell = ctk.CTkFrame(self.root, fg_color=SIDEBAR_BG, corner_radius=0, width=240)
        shell.pack(side=tk.LEFT, fill=tk.Y)
        shell.pack_propagate(False)
        ctk.CTkFrame(self.root, fg_color=BORDER, width=1, corner_radius=0).pack(
            side=tk.LEFT, fill=tk.Y
        )
        nav = ctk.CTkFrame(shell, fg_color="transparent")
        nav.pack(fill=tk.BOTH, expand=True, padx=10, pady=(16, 12))

        brand = ctk.CTkFrame(nav, fg_color="transparent")
        brand.pack(fill=tk.X, padx=6, pady=(0, 20))
        self._logo_image = _load_logo_image()
        if self._logo_image:
            ctk.CTkLabel(brand, image=self._logo_image, text="").pack(side=tk.LEFT)
        names = ctk.CTkFrame(brand, fg_color="transparent")
        names.pack(side=tk.LEFT, padx=(10, 0))
        ctk.CTkLabel(
            names, text=APP_NAME, font=_font(14, "semibold"), text_color=TEXT_COLOR,
            height=18, anchor="w",
        ).pack(anchor="w")
        ctk.CTkLabel(
            names, text="Marketing photo library", font=_font(11), text_color=MUTED_TEXT_COLOR,
            height=16, anchor="w",
        ).pack(anchor="w")

        def section(parent, text, pady=(0, 4)):
            row = ctk.CTkFrame(parent, fg_color="transparent")
            row.pack(fill=tk.X, pady=pady)
            ctk.CTkLabel(
                row, text=text, text_color=FAINT_TEXT, font=_font(11, "medium"), height=22,
            ).pack(side=tk.LEFT, padx=(10, 0))
            return row

        # Footer first (packed from the bottom up), so the lists above can
        # grow without pushing it off the window.
        footer = ctk.CTkFrame(nav, fg_color="transparent")
        footer.pack(side=tk.BOTTOM, fill=tk.X)
        # The update card: always says which version this is, and when a newer
        # one is ready, what it brings and a button to switch to it.
        card = self.update_card = ctk.CTkFrame(
            footer, fg_color=SURFACE, corner_radius=10, border_width=1, border_color=BORDER,
        )
        card.pack(fill=tk.X, pady=(0, 10))
        top = self._update_top = ctk.CTkFrame(card, fg_color="transparent")
        top.pack(fill=tk.X, padx=12, pady=(9, 0))
        self.update_title = ctk.CTkLabel(top, text="", font=_font(12, "semibold"),
                                         text_color=TEXT_COLOR, anchor="w", height=18)
        self.update_title.pack(side=tk.LEFT)
        self.update_pill = ctk.CTkLabel(top, text="", font=_font(10, "medium"),
                                        text_color=FAINT_TEXT, anchor="e", height=18)
        self.update_pill.pack(side=tk.RIGHT)
        # On a short screen the card is one line: this small button stands in
        # for the What's new link below.
        self._whats_new_icon = ctk.CTkButton(
            top, text="", width=22, height=20, corner_radius=6, fg_color="transparent",
            hover_color=SURFACE_HOVER, image=_icon("whatsnew", 13, MUTED_TEXT_COLOR),
            command=self._show_whats_new,
        )
        self.update_detail = ctk.CTkLabel(card, text="", font=_font(11), text_color=MUTED_TEXT_COLOR,
                                          anchor="w", justify="left", wraplength=196)
        self.update_detail.pack(fill=tk.X, padx=12, pady=(1, 0))
        actions = self._update_actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.pack(fill=tk.X, padx=8, pady=(4, 7))
        self.update_button = _make_primary_button(actions, "Restart to update", self._restart_app,
                                                  width=0)
        self.update_button.configure(height=28, font=_font(12, "semibold"))
        self.whats_new_link = ctk.CTkButton(
            actions, text="What's new", width=0, height=26, corner_radius=6,
            fg_color="transparent", hover_color=SURFACE_HOVER, text_color=MUTED_TEXT_COLOR,
            font=_font(11, "medium"), image=_icon("whatsnew", 13, MUTED_TEXT_COLOR),
            compound="left", command=self._show_whats_new,
        )
        self.whats_new_link.pack(side=tk.LEFT)
        self._update_info = {"version": APP_VERSION, "changes": list(WHATS_NEW)}
        self._set_update_card("current")
        self._footer_divider = ctk.CTkFrame(footer, fg_color=BORDER, height=1, corner_radius=0)
        self._footer_divider.pack(fill=tk.X, pady=(0, 10))
        source = self._source_row = ctk.CTkFrame(footer, fg_color="transparent")
        source.pack(fill=tk.X, padx=10)
        ctk.CTkLabel(
            source, text="\u25cf", text_color=SOURCE_DOT, font=_font(10), width=12, height=18,
        ).pack(side=tk.LEFT, anchor="n")
        self.folder_var = tk.StringVar(value="No folder selected")
        ctk.CTkLabel(
            source, textvariable=self.folder_var, text_color=MUTED_TEXT_COLOR, font=_font(12),
            justify="left", anchor="w", wraplength=180,
        ).pack(side=tk.LEFT, padx=(6, 0), fill=tk.X)
        stamp = _build_stamp()
        self._stamp_label = None
        if stamp and not getattr(sys, "frozen", False):
            # Running from source: when the code last changed.
            self._stamp_label = ctk.CTkLabel(
                footer, text=f"Source changed {stamp}", text_color=FAINT_TEXT, font=_font(11),
                anchor="w", height=18,
            )
            self._stamp_label.pack(fill=tk.X, padx=(28, 10), pady=(2, 0))

        system = self.system_section = ctk.CTkFrame(nav, fg_color="transparent")
        system.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 10))
        # Everything one click away, but small: a full row for the thing done
        # most often, the rest as compact pairs - about half the height of a
        # plain list, so the collections above keep their room.
        self._system_header = section(system, "SYSTEM", pady=(4, 0))
        self.sharepoint_button = self._make_system_item(
            system, "  Check for new photos", self._browse_sharepoint, icon="refresh"
        )
        pairs = ctk.CTkFrame(system, fg_color="transparent")
        pairs.pack(fill=tk.X)
        pairs.grid_columnconfigure((0, 1), weight=1, uniform="system")
        entries = [
            ("Updates", "update", lambda: self._check_for_updates(interactive=True)),
            ("Local folder", "local", None),
            ("Settings", "settings", self._open_settings),
            ("Support", "support", self._show_support_form),
            ("Tour", "tour", self._start_tour),
            ("About", "info", self._show_about),
        ]
        for n, (text, icon, command) in enumerate(entries):
            item = self._make_system_item(pairs, f" {text}", command, icon=icon, compact=True)
            item.grid(row=n // 2, column=n % 2, sticky="ew")
            if command is None:
                item.configure(command=lambda w=item: self._show_local_folder_menu(w))
        # Not shown: only its enabled state is used, by the Local folder menu.
        self.index_button = ctk.CTkButton(system, text="", command=self._start_indexing)
        self.index_button.configure(state=tk.DISABLED)

        section(nav, "LIBRARY")
        self.all_card = self._make_collection_card(
            nav, "  All photos", lambda: self._set_view("all"), icon="photos"
        )
        self.fav_card = self._make_collection_card(
            nav, "  Favourites", lambda: self._set_view("favourites"), icon="star"
        )
        self.history_card = self._make_collection_card(
            nav, "  Search history", lambda: self._set_view("history"), icon="history"
        )
        heading = self.collections_heading = section(nav, "COLLECTIONS", pady=(20, 4))
        add_button = ctk.CTkButton(
            heading, text="" if _icon("add", 14) else "+", image=_icon("add", 14),
            width=26, height=24, corner_radius=6,
            fg_color="transparent", hover_color=NAV_HOVER, text_color=TEXT_SOFT,
            font=_font(15, "semibold"),
        )
        add_button.configure(command=lambda: self._show_add_collection_menu(add_button))
        self.collection_add_button = add_button
        add_button.pack(side=tk.RIGHT, padx=(0, 4))
        # Fetches shared collections now - after someone has been added to
        # one, or to pick up photos others just put in.
        self.shared_sync_button = ctk.CTkButton(
            heading, text="" if _icon("sync", 13) else "\u21bb", image=_icon("sync", 13),
            width=26, height=24, corner_radius=6,
            fg_color="transparent", hover_color=NAV_HOVER, text_color=TEXT_SOFT,
            font=_font(13), command=lambda: self._sync_shared(interactive=True, announce=True),
        )
        self.shared_sync_button.pack(side=tk.RIGHT)
        # Collection items are rebuilt into this frame (see
        # _rebuild_collection_cards), between the heading and the footer.
        # Scrolls on its own when there are more collections than fit, so
        # the System section below never covers the last ones.
        self.sidebar = ctk.CTkScrollableFrame(
            nav, fg_color=SIDEBAR_BG, corner_radius=0,
            scrollbar_button_color="#1d2621", scrollbar_button_hover_color="#33443b",
        )
        self.sidebar.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self._collections_hint = None
        self.collection_cards = {}

        main = ctk.CTkFrame(self.root, fg_color="transparent")
        main.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(32, 24), pady=(24, 0))

        # Page title and a one-line status beneath it, like the Cockpit's
        # "Guten Tag, Ryan" header.
        head = ctk.CTkFrame(main, fg_color="transparent")
        head.pack(fill=tk.X, pady=(0, 18))
        titles = ctk.CTkFrame(head, fg_color="transparent")
        titles.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.title_var = tk.StringVar(value="All photos")
        ctk.CTkLabel(
            titles, textvariable=self.title_var, font=_font(22, "semibold"),
            text_color=TEXT_COLOR, anchor="w", height=30,
        ).pack(anchor="w")
        self.status_var = tk.StringVar(value="Loading your photo library...")
        ctk.CTkLabel(
            titles, textvariable=self.status_var, font=_font(13), text_color=MUTED_TEXT_COLOR,
            anchor="w", height=20,
        ).pack(anchor="w", pady=(2, 0))
        # A hairline that only shows while something is running: its track is
        # the background colour, so at rest there is nothing to see.
        self.progress = ctk.CTkProgressBar(
            head, width=140, height=3, corner_radius=2,
            fg_color=BG_COLOR, progress_color=PRIMARY_HOVER,
        )
        # CustomTkinter still draws a round dot at 0%, so the bar takes the
        # background colour whenever it is empty and idle.
        _set, _start = self.progress.set, self.progress.start

        def progress_set(value, *args, **kwargs):
            try:
                idle = value <= 0 and self.progress.cget("mode") == "determinate"
                self.progress.configure(progress_color=BG_COLOR if idle else PRIMARY_HOVER)
            except Exception:
                pass
            return _set(value, *args, **kwargs)

        def progress_start(*args, **kwargs):
            try:
                self.progress.configure(progress_color=PRIMARY_HOVER)
            except Exception:
                pass
            return _start(*args, **kwargs)

        self.progress.set, self.progress.start = progress_set, progress_start
        self.progress.set(0)
        self.progress.pack(side=tk.RIGHT, anchor="s", pady=(0, 8))

        search_frame = self.search_frame = ctk.CTkFrame(
            main, fg_color=SURFACE, corner_radius=12, border_width=1, border_color=BORDER,
        )
        search_frame.pack(fill=tk.X, pady=(0, 12))
        search_icon = _icon("search", 16, MUTED_TEXT_COLOR)
        if search_icon is not None:
            ctk.CTkLabel(search_frame, image=search_icon, text="").pack(side=tk.LEFT, padx=(16, 0))

        self.query_var = tk.StringVar()
        self.query_entry = ctk.CTkEntry(
            search_frame, textvariable=self.query_var, fg_color=SURFACE,
            text_color=TEXT_COLOR, border_width=0, corner_radius=8, height=38,
            font=_font(14),
            placeholder_text="Describe a photo - in English or German…",
            placeholder_text_color=FAINT_TEXT,
        )
        self.query_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 8), pady=8)
        self.query_entry.bind("<Return>", lambda _e: self._start_text_search())
        self.query_entry.bind(
            "<FocusIn>", lambda _e: search_frame.configure(border_color=FOCUS_BORDER)
        )
        self.query_entry.bind("<FocusOut>", lambda _e: search_frame.configure(border_color=BORDER))

        self.image_search_button = _make_secondary_button(
            search_frame, "Find similar…", self._start_image_search, width=136, icon="similar"
        )
        self.image_search_button.configure(height=36, corner_radius=8)
        self.image_search_button.pack(side=tk.LEFT, padx=(0, 8), pady=8)
        self.image_search_button.configure(state=tk.DISABLED)

        self.search_button = _make_primary_button(search_frame, "Search", self._start_text_search)
        self.search_button.pack(side=tk.LEFT, padx=(0, 8), pady=8)
        self.search_button.configure(state=tk.DISABLED)

        # height=1: an empty CTkFrame is 200 px tall by default, which left a
        # gap under the search box on a computer without any history yet.
        self.suggestion_bar = ctk.CTkFrame(main, fg_color="transparent", height=1)
        self.suggestion_bar.pack(fill=tk.X, pady=(0, 14))

        # Scrollable results grid - CTkScrollableFrame already handles its
        # own internal canvas, scrollbar, and mousewheel binding.
        self.results_frame = ctk.CTkScrollableFrame(
            main, fg_color=BG_COLOR, corner_radius=0, scrollbar_button_color=BORDER,
            scrollbar_button_hover_color="#3a4c42",
        )
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
        self._bind_selection_keys()
        self._refresh_suggestions()  # the filters are there from the start
        self.root.bind("<Configure>", self._on_root_configure, add="+")
        self.root.after(300, lambda: self._apply_sidebar_density(force=True))
        self._set_view("all")

    # ---------- short screens ----------
    #
    # On a laptop screen the sidebar tightens up by itself so the
    # collections keep their room: the update card becomes one line, the
    # SYSTEM heading and the source lines at the bottom step aside. A tall
    # screen shows everything.

    SHORT_SCREEN = 860  # window height, in px, below which the sidebar tightens

    def _on_root_configure(self, event):
        if event.widget is not self.root:
            return
        if getattr(self, "_density_after", None) is not None:
            try:
                self.root.after_cancel(self._density_after)
            except Exception:
                pass
        self._density_after = self.root.after(120, self._apply_sidebar_density)

    def _sidebar_is_compact(self):
        return bool(getattr(self, "_sidebar_compact", False))

    def _apply_sidebar_density(self, force=False):
        self._density_after = None
        try:
            height = self.root.winfo_height()
        except tk.TclError:
            return
        if height <= 1:
            return
        compact = height < self.SHORT_SCREEN
        if not force and compact == getattr(self, "_sidebar_compact", None):
            return
        self._sidebar_compact = compact
        footer_rows = [
            (self._footer_divider, {"fill": tk.X, "pady": (0, 10)}, self.update_card),
            (self._source_row, {"fill": tk.X, "padx": 10}, self._footer_divider),
            (self._stamp_label, {"fill": tk.X, "padx": (28, 10), "pady": (2, 0)}, self._source_row),
        ]
        try:
            if compact:
                for widget, _opts, _after in footer_rows:
                    if widget is not None:
                        widget.pack_forget()
                self._system_header.pack_forget()
                self.update_card.pack_configure(pady=(0, 8))
            else:
                for widget, opts, after in footer_rows:
                    if widget is not None:
                        widget.pack(after=after, **opts)
                self._system_header.pack(fill=tk.X, pady=(4, 0), before=self.sharepoint_button)
                self.update_card.pack_configure(pady=(0, 10))
        except tk.TclError:
            pass
        self._apply_update_card_density()

    def _apply_update_card_density(self):
        """One line on a short screen, unless an update needs its button."""
        if not hasattr(self, "_update_actions"):
            return
        quiet = getattr(self, "_update_state", "current") in ("current", "checking", "uptodate")
        one_line = self._sidebar_is_compact() and quiet
        try:
            if one_line:
                self._update_actions.pack_forget()
                self._update_top.pack_configure(pady=(8, 8))
                self._whats_new_icon.pack(side=tk.RIGHT, padx=(0, 4), before=self.update_pill)
            else:
                self._whats_new_icon.pack_forget()
                self._update_top.pack_configure(pady=(9, 0))
                if not self._update_actions.winfo_manager():
                    self._update_actions.pack(fill=tk.X, padx=8, pady=(4, 7))
        except tk.TclError:
            pass

    def _make_system_item(self, parent, text, command, icon, compact=False):
        """A System entry: smaller and quieter than the library and
        collections above it - used now and then, so it steps back and
        leaves room for more collections. Compact ones sit two to a row."""
        item = ctk.CTkButton(
            parent, text=text, command=command, height=24 if compact else 26, corner_radius=7,
            fg_color=NAV_BG, hover_color=NAV_HOVER, text_color=SYSTEM_TEXT,
            text_color_disabled=FAINT_TEXT, anchor="w", width=0,
            image=_icon(icon, 12 if compact else 13, SYSTEM_TEXT),
            compound="left", font=_font(11, "medium"),
        )
        if not compact:
            item.pack(fill=tk.X, pady=0)
        return item

    def _make_collection_card(self, parent, text, command, icon="folder"):
        """A flat sidebar item: no panel of its own, a soft highlight on
        hover and a filled one when selected."""
        card = ctk.CTkButton(
            parent,
            text=text,
            command=command,
            height=34,
            corner_radius=8,
            fg_color=NAV_BG,
            hover_color=NAV_HOVER,
            text_color=TEXT_SOFT,
            text_color_disabled=FAINT_TEXT,
            anchor="w",
            image=_icon(icon, 16),
            compound="left",
            font=_font(14, "medium"),
        )
        card.pack(fill=tk.X, pady=(0, 1))
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
        """The hand-curated collections, as positions in the loaded index.

        Stored by photo id (see _collection_ids), so they survive the index
        being rebuilt in a different order. Empty collections are kept: they
        are the drop targets, so they have to stay on screen before anything
        has been put in them.
        """
        position = {
            (meta or {}).get("item_id"): i for i, meta in enumerate(self.engine.meta or [])
        }
        return {
            name: [position[i] for i in ids if i in position]
            for name, ids in self._collection_ids().items()
        }


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
        """Recreates the collection cards in the sidebar."""
        for card in self.collection_cards.values():
            card.destroy()
        self.collection_cards = {}
        if getattr(self, "_collections_hint", None) is not None:
            self._collections_hint.destroy()
            self._collections_hint = None
        shared_names = set(self._shared_by_name())
        for name, idxs in (self._collections or {}).items():
            # Long names are shortened on the card only; the sidebar is narrow.
            shown = name if len(name) <= 17 else name[:16].rstrip() + "…"
            card = self._make_collection_card(
                self.sidebar, f"  {shown} ({len(idxs)})", lambda n=name: self._set_view(n),
                icon="shared" if name in shared_names else "folder",
            )
            _bind_context_menu(card, lambda e, n=name: self._show_collection_menu(e, n))
            card.configure(
                fg_color=self._card_rest_colour(name),
                text_color=TEXT_COLOR if self._view == name else TEXT_SOFT,
            )
            self.collection_cards[name] = card
        if not self.collection_cards:
            self._collections_hint = ctk.CTkLabel(
                self.sidebar, text="No collections.\nClick + to make one.",
                text_color=MUTED_TEXT_COLOR, font=_font(12), justify="left",
            )
            self._collections_hint.pack(anchor="w", padx=6, pady=(2, 0))

    def _set_view(self, view):
        """Switches which collection the search box searches."""
        self._view = view
        if getattr(self, "title_var", None) is not None:
            self.title_var.set(
                {"all": "All photos", "favourites": "Favourites", "history": "Search history"}
                .get(view, view)
            )
        def look(selected):
            return {
                "fg_color": NAV_SELECTED if selected else NAV_BG,
                "text_color": TEXT_COLOR if selected else TEXT_SOFT,
            }

        self.all_card.configure(**look(view == "all"))
        self.fav_card.configure(**look(view == "favourites"))
        self.history_card.configure(**look(view == "history"))
        self._update_favourites_button()
        self._update_history_button()
        for name, card in self.collection_cards.items():
            card.configure(**look(view == name))

        if view == "history":
            self.query_entry.configure(placeholder_text="Describe a photo - in English or German…")
            self._show_history()
        elif view == "favourites":
            self.query_entry.configure(placeholder_text="Search your favourites...")
            self._show_favourites()
        elif view in (self._collections or {}):
            self.query_entry.configure(placeholder_text=f"Search within {view}...")
            self._show_collection(view)
        else:
            self.query_entry.configure(placeholder_text="Describe a photo - in English or German…")
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
            self._render_placeholder_grid()
            self.status_var.set(
                "No photos yet - click Check for new photos (bottom left) to load the library "
                "from SharePoint. The first time takes a few minutes."
            )
            return
        idxs = (self._filtered_sample_indices() if self._filters_active()
                else self._startup_sample_indices())
        results = [(self.engine.paths[i], 1.0, self.engine.meta[i]) for i in idxs]
        self._render_results(self._interleave_by_folder(results))
        self._sample_view = True
        if getattr(self, "_filter_count_var", None) is not None:
            self._filter_count_var.set("")
        self.status_var.set(
            f"{total:,} photos ready. Type what you are looking for, "
            "or pick a collection on the left."
        )

    def _show_collection(self, name, keep_scroll=False):
        idxs = (self._collections or {}).get(name, [])
        if not idxs:
            self._render_results([])
            self.status_var.set(
                f"{name} is empty - drag photos onto it, or right-click a photo "
                f"and choose 'Add to {name}'."
            )
            return
        results = [(self.engine.paths[i], 1.0, self.engine.meta[i]) for i in idxs[:DEFAULT_TOP_K]]
        self._render_results(results, keep_scroll=keep_scroll)
        extra = f" (showing first {DEFAULT_TOP_K})" if len(idxs) > DEFAULT_TOP_K else ""
        self.status_var.set(f"{len(idxs)} photo(s) in {name}{extra}. Search to narrow it down.")

    def _show_favourites(self, keep_scroll=False):
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
        self._render_results(results, keep_scroll=keep_scroll)
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
            messagebox.showinfo(APP_NAME, "Index the folder first before searching.")
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
                if self._query_wants_objects(to_english(query)):
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
            messagebox.showinfo(APP_NAME, "Load the photo library first.")
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

        active = self._active_subset()
        self._last_query = ""
        self._set_busy(True)
        self.progress.configure(mode="indeterminate")
        self.progress.start()
        self.status_var.set(f"Finding photos like {os.path.basename(path)}...")

        def work():
            def on_status(msg):
                self.event_queue.put(("status", msg))

            try:
                # Off the UI thread: the graphic filter needs the model, and
                # loading it here (not in the click handler) keeps the window
                # responsive if Find similar is used before it has warmed up.
                subset = self._exclude_graphic(active)
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
        """Rebuilds the row under the search box: recent searches on the
        left, the filters on the right."""
        if getattr(self, "suggestion_bar", None) is None:
            return
        for child in self.suggestion_bar.winfo_children():
            child.destroy()
        self._build_filter_chips(self.suggestion_bar)
        terms = [h.get("term") for h in self._load_history()[:4] if h.get("term")]
        if not terms:
            return
        ctk.CTkLabel(
            self.suggestion_bar, text="Recent", text_color=FAINT_TEXT,
            font=_font(12),
        ).pack(side=tk.LEFT, padx=(4, 8))
        for term in terms:
            ctk.CTkButton(
                self.suggestion_bar, text=term, height=28, corner_radius=14,
                fg_color="transparent", hover_color=SURFACE_HOVER, text_color=TEXT_SOFT,
                border_width=1, border_color=BORDER,
                font=_font(12), width=0,
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
        self._scroll_results_to_top()
        self.thumbnail_refs.clear()
        if not history:
            self.status_var.set("No searches yet - your search history will appear here.")
            return
        self.status_var.set(f"{len(history)} previous search(es). Click one to run it again.")
        top = ctk.CTkFrame(self.results_frame, fg_color="transparent")
        top.pack(fill=tk.X, padx=4, pady=(2, 8))
        ctk.CTkButton(
            top, text="Clear search history", width=170, height=30, corner_radius=8,
            # A muted red on hover: the one destructive action on this screen
            # should look different from the searches beside it.
            fg_color=CARD_COLOR, hover_color="#7a2330", text_color=TEXT_COLOR,
            font=_font(12, "semibold"),
            command=self._clear_history,
        ).pack(side=tk.RIGHT)
        for entry in history:
            term = entry.get("term", "")
            when = (entry.get("at") or "").replace("T", " ")
            row = ctk.CTkFrame(self.results_frame, fg_color=CARD_COLOR, corner_radius=10)
            row.pack(fill=tk.X, padx=4, pady=3)
            ctk.CTkButton(
                row, text=term, anchor="w", height=34, corner_radius=8,
                fg_color="transparent", hover_color=ACCENT_COLOR, text_color=TEXT_COLOR,
                font=_font(13, "semibold"),
                command=lambda t=term: self._search_term(t),
            ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 4), pady=4)
            ctk.CTkLabel(
                row, text=when, text_color=MUTED_TEXT_COLOR, font=_font(11),
            ).pack(side=tk.RIGHT, padx=12)

    def _clear_history(self):
        """Empties this machine's search history, after asking.

        Only the list of past searches goes. The click feedback that helps
        rank results for everyone is a separate file and is left alone.
        """
        count = len(self._load_history())
        if not count:
            return
        if not messagebox.askyesno(
            "Clear search history",
            f"Clear all {count} searches from your history?\n\n"
            "This also clears the Recent shortcuts under the search box. "
            "It can't be undone.",
            icon="warning",
        ):
            return
        try:
            with open(self._history_path(), "w", encoding="utf-8") as f:
                json.dump({"searches": []}, f)
        except OSError as exc:
            _log_reuse(f"HISTORY CLEAR FAILED: {exc!r}")
            self.status_var.set("Could not clear the search history.")
            return
        self._history = []
        self._refresh_suggestions()
        self._update_history_button()
        self._show_history()
        self.status_var.set("Search history cleared.")

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
            count = 0
            try:
                count = 0 if FIRST_RUN_PREVIEW else self._load_fast_cache()
            except Exception:
                _log_crash("startup load", sys.exc_info())
            if count:
                self.event_queue.put(("startup_loaded", count))
                # Decode the opening screen's thumbnails here, off the main
                # thread: Pillow lets go of the interpreter while decoding,
                # so the splash keeps animating and the window only has to
                # build widgets once it is ready.
                try:
                    for i in self._startup_sample_indices():
                        self._thumb_image(self.engine.paths[i])
                except Exception:
                    _log_crash("decoding startup thumbnails", sys.exc_info())
            # The search model is no longer loaded here - see _warm_model.
            self.event_queue.put(("startup_ready", count) if count else ("startup_empty",))

        self._run_in_background(work)

    def _warm_model(self):
        """Loads the search model in the background.

        Started once the window is on screen rather than during startup:
        importing torch holds the interpreter for seconds, which would
        freeze the intro animation if the two overlapped. A search started
        before this finishes simply waits for it.
        """
        if self._warming or self.engine.model is not None:
            return
        self._warming = True

        def work():
            try:
                self.engine.load_model()
                _boot_mark("search model ready")
            except Exception:
                _log_crash("loading the search model", sys.exc_info())

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
            font=_font(14, "semibold"),
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
            # Team picks this person took out of their own view stay out.
            for key in self._hidden_favourites():
                merged.pop(key, None)
            # Local saves may not have been shared yet, so they go on top.
            merged.update(self._read_json(self._favourites_path()).get("items") or {})
            self._favourites = merged
        return self._favourites

    def _hidden_favourites_path(self):
        local_folder, _t, _f, _i = self._sp_cache_dirs()
        return os.path.join(local_folder, "favourites_hidden.json")

    def _hidden_favourites(self):
        """Team favourites this person removed from their own view. Only
        hidden here - the teammate's shortlist is never touched."""
        return set(self._read_json(self._hidden_favourites_path()).get("items") or [])

    def _set_hidden_favourite(self, key, hidden):
        keys = self._hidden_favourites()
        (keys.add if hidden else keys.discard)(key)
        try:
            os.makedirs(os.path.dirname(self._hidden_favourites_path()), exist_ok=True)
            with open(self._hidden_favourites_path(), "w", encoding="utf-8") as f:
                json.dump({"items": sorted(keys)}, f)
        except OSError:
            pass

    def _unfavourite(self, key, path, meta):
        """Takes a photo out of this person's Favourites: their own star is
        removed, a teammate's pick is hidden for them only."""
        if key in self._own_favourites():
            self._toggle_favourite(key, path, meta)
        else:
            self._set_hidden_favourite(key, True)
        self._favourites = None
        self._update_favourites_button()

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
            if key in self._hidden_favourites():
                self._set_hidden_favourite(key, False)  # starred again: back in view
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
        own_name = f"{_sharepoint().FAVOURITES_PREFIX}{machine_key()}.json"
        try:
            # Uploaded even when empty: skipping the upload left a removed
            # favourite sitting in this machine's shared file forever, so
            # un-starring never reached anyone else.
            self.sp_client.upload_json_file(root_id, own_name, {"items": own})
            merged = {}
            for name in self.sp_client.list_shared_files(root_id, _sharepoint().FAVOURITES_PREFIX):
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
                    root_id, f"{_sharepoint().FEEDBACK_PREFIX}{machine_key()}.json", own
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
                # Signs in on every sync, as it always has - not just when no
                # connection exists - so a long-running session refreshes its
                # Microsoft sign-in here. Under the same lock as everything
                # else that signs in, so a preview opening meanwhile can't
                # start a second sign-in alongside it.
                with self._sp_lock:
                    if self.sp_client is None:
                        self.sp_client = _sharepoint().SharePointClient()
                    self.sp_client.sign_in()
                self.event_queue.put(("sp_signed_in",))
            except Exception as exc:
                self.event_queue.put(("error", f"Sign-in failed: {exc}"))

        self._run_in_background(work)

    def _sp_cache_dirs(self):
        local_folder = os.path.join(
            DATA_DIR,
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
                except _sharepoint().DeltaExpired:
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
                    self._ensure_sp_client()
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
                    self.progress.configure(mode="determinate")
                    self.progress.set(0)
                    self._set_busy(False)
                    self.status_var.set(f"Indexed {event[1]} images. Ready to search.")
                elif kind == "search_done":
                    self.progress.stop()
                    self.progress.configure(mode="determinate")
                    self.progress.set(0)
                    self._set_busy(False)
                    self._render_results(event[1])
                    mode = getattr(self, "_last_search_mode", "")
                    note = ""
                    if mode == "objects":
                        note = "  \u00b7  Close-ups first, people shots below."
                    elif mode == "image":
                        note = "  \u00b7  Most visually similar first."
                    searched_as = getattr(self.engine, "last_translation", None)
                    if searched_as and mode != "image":
                        note = f"  \u00b7  Searched as “{searched_as}”." + note
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
                elif kind == "preview_image":
                    self._apply_preview(event[1], event[2], event[3])
                elif kind == "preview_done":
                    self._end_sharpening(event[1])
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
                    self.progress.configure(mode="determinate")
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
                    if self._boot is not None:
                        # Real width before the first layout, so the grid is
                        # drawn once at the right column count, not redrawn.
                        self.root.update_idletasks()
                        self._render_done_callback = self._boot.render_done
                    # Fills the opening grid now the cached index is loaded,
                    # instead of leaving an empty window until the first search.
                    self._show_all_photos()
                    # Past searches are on disk, so the shortcut chips can be
                    # there from the start rather than only after a search.
                    self._refresh_suggestions()
                    self._refresh_collections(from_cache=True)
                    if self._boot is None:
                        self._warm_model()
                elif kind == "startup_loaded":
                    if self._boot is not None:
                        self._boot.library_loaded(event[1])
                elif kind == "startup_empty":
                    self._refresh_collections(from_cache=True)
                    if self._boot is not None:
                        self._boot.library_empty()
                    else:
                        self._warm_model()
                elif kind == "collections_ready":
                    self._collections = event[1]
                    self._rebuild_collection_cards()
                    if self._boot is not None:
                        self._boot.collections_done()
                elif kind == "sp_open_ready":
                    self._set_busy(False)
                    delivered = self._deliver_to_download_folder(event[1], event[2])
                    if self._download_folder():
                        self.status_var.set(f"Saved to {os.path.dirname(delivered)}")
                    else:
                        self.status_var.set("Ready to search.")
                    open_in_system_viewer(delivered)
                elif kind == "shared_records":
                    self._apply_shared_records(event[1], event[2], event[3])
                elif kind == "shared_upsert":
                    self._apply_shared_upsert(event[1], event[2], event[3])
                elif kind == "shared_removed":
                    self._apply_shared_removed(event[1], event[2])
                elif kind == "support_fallback":
                    self._support_fallback(event[1], event[2], event[3])
                elif kind == "shared_invite":
                    self._offer_invite(event[1], event[2])
                elif kind == "invite_fallback":
                    self._invite_fallback(event[1], event[2], event[3])
                elif kind in ("shared_status", "update_status"):
                    self.status_var.set(event[1])
                elif kind == "update_ready":
                    self._show_update_notice(event[1], restart=True, changes=event[2])
                elif kind == "update_full":
                    self._show_update_notice(event[1], restart=False, note=event[2], changes=event[3])
                elif kind == "update_uptodate":
                    if not self._update_ready:
                        self._set_update_card("uptodate")
                    if event[1]:
                        self.status_var.set(f"You're up to date - version {APP_VERSION}.")
                elif kind == "update_card":
                    self._set_update_card(event[1])
                elif kind == "error":
                    self.progress.stop()
                    self.progress.configure(mode="determinate")
                    self.progress.set(0)
                    self._set_busy(False)
                    self.status_var.set("Error.")
                    messagebox.showerror(APP_NAME, event[1])
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    # ---------- results rendering ----------

    def _desired_columns(self):
        """How many result columns fit the results area at its current
        width, so the grid fills the window (and expands to fill it when
        maximized) instead of sitting at a fixed column count."""
        # The scroll canvas is the visible area; its size is settled by
        # geometry management, so it is right even before the inner frame's
        # own resize event has been processed.
        try:
            width = self.results_frame._parent_canvas.winfo_width()
        except Exception:
            width = self.results_frame.winfo_width()
        if width <= 1:
            # Not laid out yet (the window is built hidden during the intro);
            # the resize that follows the reveal re-lays it out at the real width.
            return GRID_COLUMNS * (THUMB_SIZE[0] + GRID_CELL_PADDING)
        # The justified grid depends on the exact width, so the "column
        # count" is now simply the width: any real change re-lays it out.
        return width

    def _on_results_area_resized(self):
        self.root.after_idle(self._place_selection_bar)
        self._on_results_area_resized_grid()

    def _on_results_area_resized_grid(self):
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
            elif getattr(self, "_placeholder_grid", False):
                self._render_placeholder_grid()

    def _thumb_image(self, path):
        """The decoded thumbnail as a PIL image, cached.

        Never touches Tk, so it is safe from a worker thread - which is how
        the opening screen's thumbnails get decoded while the splash plays.
        """
        img = self._thumb_cache.get(path)
        if img is None:
            with Image.open(path) as source:
                img = source.convert("RGB")
            img.thumbnail(THUMB_SIZE)
            # Bounded so a long session over a 33k-image library cannot grow
            # without limit; dropping the whole cache is fine, it refills.
            if len(self._thumb_cache) >= THUMB_CACHE_MAX:
                self._thumb_cache.clear()
            self._thumb_cache[path] = img
        return img

    def _star_badges(self):
        if self._badges is None:
            self._badges = {True: _draw_star_badge(True), False: _draw_star_badge(False)}
        return self._badges

    def _tile_mask(self, size):
        """An alpha mask with rounded corners for one tile. The corner is
        drawn once at 4x and reused, so a mask costs four pastes."""
        r = TILE_RADIUS
        if self._tile_corners is None:
            s = 4
            big = Image.new("L", (r * s, r * s), 0)
            ImageDraw.Draw(big).pieslice((0, 0, 2 * r * s - 1, 2 * r * s - 1), 180, 270, fill=255)
            c = big.resize((r, r), Image.LANCZOS)
            self._tile_corners = (
                c, c.transpose(Image.FLIP_LEFT_RIGHT),
                c.transpose(Image.FLIP_TOP_BOTTOM), c.transpose(Image.ROTATE_180),
            )
        tl, tr, bl, br = self._tile_corners
        w, h = size
        mask = Image.new("L", size, 255)
        mask.paste(tl, (0, 0))
        mask.paste(tr, (w - r, 0))
        mask.paste(bl, (0, h - r))
        mask.paste(br, (w - r, h - r))
        return mask

    def _tile_photo(self, path, size, favourite, hover=False, selected=False):
        return ImageTk.PhotoImage(self._tile_pil(path, size, favourite, hover, selected))

    def _tile_pil(self, path, size, favourite, hover=False, selected=False):
        """One grid tile, ready for Tk: the photo filled to `size` (centre
        crop of at most a few pixels, never stretched), rounded corners, and
        the star - always when it is a favourite, otherwise only on hover,
        so a screen of photos is not covered in empty stars.
        """
        img = ImageOps.fit(self._thumb_image(path), size, Image.LANCZOS)
        tile = Image.new("RGBA", size, _hex_rgb(BG_COLOR) + (255,))
        tile.paste(img, (0, 0), self._tile_mask(size))
        if favourite or hover:
            badge = self._star_badges()[bool(favourite)]
            tile.alpha_composite(
                badge, (size[0] - badge.width - STAR_BADGE_INSET, STAR_BADGE_INSET)
            )
        if selected:
            tile = self._selected_look(tile, size)
        return tile

    def _tile_hover(self, label, on):
        if on:
            self._hover_label = label
        elif getattr(self, "_hover_label", None) is label:
            self._hover_label = None
        tile = getattr(label, "tile", None)
        if tile is None or getattr(label, "dragging", False) or getattr(label, "selected", False):
            return
        path, size, key = tile
        if key in self._render_favs:
            return  # the gold star already shows
        try:
            if on:
                if getattr(label, "hover_image", None) is None:
                    label.hover_image = self._tile_photo(path, size, False, hover=True)
                label.configure(image=label.hover_image)
            else:
                label.configure(image=label.image)
        except Exception:
            pass

    @staticmethod
    def _in_star(event):
        """Whether a click on a thumbnail landed on its star (with a little
        slack, since the badge is small)."""
        try:
            width = event.widget.winfo_width()
        except tk.TclError:
            return False
        badge_w, badge_h = STAR_BADGE_SIZE
        return (event.x >= width - badge_w - STAR_BADGE_INSET - 4
                and event.y <= STAR_BADGE_INSET + badge_h + 4)

    def _on_thumb_release(self, event, key, path, meta):
        if getattr(self, "_suppress_release", False):
            self._suppress_release = False  # this click was a Ctrl/Shift selection
            return
        dragged = self._drag is not None and self._drag.get("active")
        self._drag_release(event)
        if not dragged and getattr(self, "_sel", None):
            # While photos are picked, a plain click adds or drops one.
            pos = getattr(event.widget, "pos", None)
            if pos is not None:
                self._select_pos(pos, pos not in self._sel)
                self._sel_anchor = pos
                self._update_selection_bar()
            return
        if not dragged and self._in_star(event):
            self._on_star_clicked(key, path, meta, event.widget)

    def _on_thumb_double(self, event, path, meta):
        if getattr(self, "_sel", None):
            return  # picking photos, not opening one
        if self._in_star(event):
            return  # the clicks were for the star, not for opening the photo
        self._show_preview(path, meta)

    def _startup_sample_indices(self):
        """The photos the opening view shows: a step through the whole
        library rather than the first N, which would all be from whichever
        folder indexed first."""
        total = len(self.engine.paths)
        step = max(1, total // ALL_PHOTOS_PREVIEW)
        return list(range(0, total, step))[:ALL_PHOTOS_PREVIEW]

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

    def _preview_box(self):
        """The largest the preview draws a photo: most of the screen, so it
        is worth opening, while leaving the desktop visible around it."""
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        return max(640, int(sw * 0.82)), max(420, int(sh * 0.72))

    def _show_preview(self, path, meta):
        """A large view inside the app, at full resolution.

        Opens at once from what is already on disk and sharpens in steps:
        the grid thumbnail, then SharePoint's 800px preview, then the
        original itself (downloaded in the background, which also makes
        Copy and Open original instant afterwards). Each step only ever
        replaces a lower-quality one.
        """
        name = (meta or {}).get("name") or os.path.basename(path)
        item_id = (meta or {}).get("item_id")
        box = self._preview_box()
        win = ctk.CTkToplevel(self.root)
        win.title(name)
        win.configure(fg_color=BG_COLOR)
        width, height = box[0] + 24, box[1] + 84
        x = max(0, (win.winfo_screenwidth() - width) // 2)
        y = max(0, (win.winfo_screenheight() - height) // 2 - 20)
        win.geometry(f"{width}x{height}+{x}+{y}")
        try:
            win.after(200, lambda: _apply_icon(win))
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
        if item_id:
            open_original = lambda: self._open_sp_result(item_id, meta["name"])
        else:
            open_original = lambda: open_in_system_viewer(path)
        _make_button(bar, "Open original", open_original, width=140).pack(side=tk.RIGHT, padx=(8, 0))
        _make_button(
            bar, "Copy", lambda: self._copy_to_clipboard(path, meta, token=token), width=90,
        ).pack(side=tk.RIGHT, padx=(8, 0))

        self._preview = {
            "token": token, "label": image_label, "win": win, "caption": caption,
            "quality": -1, "box": box,
        }

        large = os.path.join(self._preview_dir(), f"{item_id}.jpg") if item_id else None
        try:
            if large and os.path.exists(large):
                self._apply_preview(token, Image.open(large).convert("RGB"), 1)
            else:
                self._apply_preview(token, self._thumb_image(path), 0)
        except Exception:
            pass

        caption.configure(text=f"{caption.cget('text')}   (sharpening...)")
        if item_id:
            if not (large and os.path.exists(large)):
                self._fetch_preview(item_id, token, box)
            self._load_original_preview(item_id, name, token, box)
        else:
            self._load_original_preview(None, path, token, box)

    def _apply_preview(self, token, img, quality):
        """Shows `img` in the open preview - unless it was closed, replaced
        by another preview, or already showing something sharper."""
        preview = self._preview
        if preview is None or preview["token"] is not token or quality <= preview["quality"]:
            return
        try:
            fitted = self._fit(img, preview["box"]) if img.size != preview["box"] else img
            photo = ctk.CTkImage(light_image=fitted, dark_image=fitted, size=fitted.size)
            preview["label"].configure(image=photo)
            preview["label"]._preview_ref = photo  # keep it alive with the widget
            preview["quality"] = quality
            if quality >= 2:
                self._end_sharpening(token)
        except Exception:
            pass

    def _end_sharpening(self, token):
        preview = self._preview
        if preview is None or preview["token"] is not token:
            return
        try:
            text = preview["caption"].cget("text")
            preview["caption"].configure(text=text.replace("   (sharpening...)", ""))
        except Exception:
            pass

    @staticmethod
    def _open_upright(source):
        """Opens a photo the right way up. Phones and many cameras save the
        pixels as shot and record the rotation in an EXIF tag, so a portrait
        photo opened naively comes out sideways."""
        img = Image.open(source)
        return ImageOps.exif_transpose(img) or img

    def _decode_to_fit(self, source, box):
        """Decodes a photo straight to roughly the size it will be shown at.
        For JPEGs, draft() has the decoder skip detail that would be thrown
        away anyway - a 24-megapixel original decodes many times faster."""
        with Image.open(source) as probe:
            try:
                probe.draft("RGB", box)
            except Exception:
                pass
            img = ImageOps.exif_transpose(probe) or probe
            img = img.convert("RGB")
        return self._fit(img, box)

    def _load_original_preview(self, item_id, name_or_path, token, box):
        """Fetches the original (or reads it, for a photo on this computer)
        and hands it to the preview at the size it will be shown."""
        def work():
            try:
                if item_id:
                    source = self._ensure_full_image(item_id, name_or_path)
                else:
                    source = name_or_path
                img = self._decode_to_fit(source, box)
                self.event_queue.put(("preview_image", token, img, 2))
            except Exception as exc:
                _log_reuse(f"ORIGINAL PREVIEW FAILED: {exc!r}")
                self.event_queue.put(("preview_done", token))

        self._run_in_background(work)

    def _ensure_sp_client(self):
        """The SharePoint connection, signing in if needed. Locked: a preview
        fetches two versions of a photo at once, and two threads must not
        both start a sign-in."""
        with self._sp_lock:
            if self.sp_client is None:
                client = _sharepoint().SharePointClient()
                client.sign_in()
                self.sp_client = client
        return self.sp_client

    def _fetch_preview(self, item_id, token, box):
        """Pulls SharePoint's 800px version of one photo - the quick middle
        step while the original downloads."""
        def work():
            try:
                client = self._ensure_sp_client()
                dest = os.path.join(self._preview_dir(), f"{item_id}.jpg")
                if client.download_thumbnail({"id": item_id}, dest, size="large"):
                    with Image.open(dest) as im:
                        img = im.convert("RGB")
                    self.event_queue.put(("preview_image", token, self._fit(img, box), 1))
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
        self._ensure_sp_client().download_file({"id": item_id}, dest)
        return dest

    def _put_image_on_clipboard(self, source):
        """Writes one image to the Windows clipboard as a device independent
        bitmap - the format Office pastes as a picture rather than as a
        file reference. Returns its pixel size."""
        import ctypes
        from io import BytesIO

        img = self._open_upright(source).convert("RGB")
        if IS_MAC:
            return self._put_image_on_mac_clipboard(img)
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

    def _put_image_on_mac_clipboard(self, img):
        """The Mac clipboard takes the picture itself, so Keynote, PowerPoint
        and Mail paste it as an image."""
        import tempfile

        tmp = os.path.join(tempfile.gettempdir(), "topshot-clipboard.jpg")
        img.save(tmp, "JPEG", quality=95)
        script = f'set the clipboard to (read (POSIX file "{tmp}") as JPEG picture)'
        subprocess.run(["osascript", "-e", script], check=True, timeout=30,
                       capture_output=True)
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

    def _local_collection_ids(self):
        """{name: [item_id, ...]} in sidebar order, exactly as curated.

        Item ids rather than positions, so a collection survives the index
        being rebuilt in a different order after a sync. Version 2 of the
        file records which collections exist, so one that was removed stays
        removed. Older files only listed photos per collection: they are read
        into the new form - renamed collections folded together, any default
        the old file lacked added - and written back on the next change.
        """
        data = self._read_json(self._collections_cache_path())
        if data.get("version") == 2:
            out = {}
            for entry in data.get("collections") or []:
                name = (entry or {}).get("name") if isinstance(entry, dict) else None
                if isinstance(name, str) and name and name not in out:
                    out[name] = [i for i in (entry.get("items") or []) if i]
            return out
        out = {name: [] for name in DEFAULT_COLLECTIONS}
        for name, ids in data.items():
            if not isinstance(ids, list):
                continue
            bucket = out.setdefault(LEGACY_COLLECTIONS.get(name, name), [])
            for item_id in ids:
                if item_id and item_id not in bucket:
                    bucket.append(item_id)
        return out

    def _collection_ids(self):
        """Every collection in the sidebar, {name: [item_id, ...]}: this
        computer's own, then the shared ones this user is in."""
        out = self._local_collection_ids()
        for name, record in self._shared_by_name().items():
            out[name] = [i for i in (record.get("items") or []) if i]
        return out

    def _write_collection_ids(self, by_id):
        """Saves collections. Local ones go to this computer's file; a
        change to a shared one is saved here at once and sent to SharePoint
        in the background, as the photos added and removed - so two people
        filling the same collection at once do not undo each other."""
        shared = self._shared_by_name()
        for name, record in shared.items():
            if name not in by_id:
                continue
            new = [i for i in by_id[name] if i]
            old = [i for i in (record.get("items") or []) if i]
            if new != old:
                added = [i for i in new if i not in old]
                removed = [i for i in old if i not in new]
                record["items"] = new
                self._save_shared_state()
                self._push_shared(dict(record), add=added, remove=removed)
        by_id = {name: ids for name, ids in by_id.items() if name not in shared}
        payload = {
            "version": 2,
            "collections": [{"name": name, "items": ids} for name, ids in by_id.items()],
        }
        try:
            os.makedirs(os.path.dirname(self._collections_cache_path()), exist_ok=True)
            with open(self._collections_cache_path(), "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False)
            return True
        except OSError as exc:
            _log_reuse(f"COLLECTIONS SAVE FAILED: {exc!r}")
            self.status_var.set("Could not save that collection - see crash.log.")
            return False

    def _collection_names(self):
        return list((self._collections or self._collection_ids()).keys())

    def _reload_collections(self):
        self._collections = self._load_collections()
        self._rebuild_collection_cards()

    def _ask_collection_name(self, title, prompt, current=None):
        """Asks for a collection name. Returns None if cancelled or unusable,
        after saying why when it is the latter."""
        dialog = ctk.CTkInputDialog(title=title, text=prompt)
        # CustomTkinter puts its own icon on new windows after ~200ms.
        dialog.after(250, lambda: _apply_icon(dialog))
        raw = dialog.get_input()
        if raw is None:
            return None
        name = " ".join(raw.split())
        if not name:
            return None
        if len(name) > COLLECTION_NAME_MAX:
            messagebox.showinfo(
                APP_NAME, f"Please keep collection names to {COLLECTION_NAME_MAX} characters."
            )
            return None
        if name.lower() in RESERVED_COLLECTION_NAMES:
            messagebox.showinfo(
                APP_NAME, f"“{name}” is already used by the app - please choose another name."
            )
            return None
        taken = {n.lower() for n in self._collection_ids() if n != current}
        if current is None:
            # Shared collections show under their own names; a new local one
            # may not borrow one of those either.
            taken |= {(r.get("name") or "").lower() for r in self._shared_state()["collections"]}
        if name.lower() in taken:
            messagebox.showinfo(APP_NAME, f"There is already a collection called “{name}”.")
            return None
        return name

    def _new_collection(self):
        name = self._ask_collection_name("New collection", "Name for the new collection:")
        if name is None:
            return
        by_id = self._collection_ids()
        by_id[name] = []
        if self._write_collection_ids(by_id):
            self._reload_collections()
            self.status_var.set(
                f"Created {name} - drag photos onto it, or right-click a photo and choose 'Add to {name}'."
            )

    def _rename_collection(self, name):
        if self._is_shared(name):
            self._manage_shared(name)
            return
        new = self._ask_collection_name(
            "Rename collection", f"New name for “{name}”:", current=name
        )
        if new is None or new == name:
            return
        by_id = self._collection_ids()
        if name not in by_id:
            return
        by_id = {(new if key == name else key): ids for key, ids in by_id.items()}  # keeps its place
        if self._write_collection_ids(by_id):
            viewing = self._view == name
            self._reload_collections()
            if viewing:
                self._set_view(new)
            self.status_var.set(f"Renamed {name} to {new}.")

    def _remove_collection(self, name):
        if self._is_shared(name):
            self._leave_or_delete_shared(name)
            return
        by_id = self._collection_ids()
        if name not in by_id:
            return
        count = len(by_id[name])
        detail = (
            f"The {count} photo(s) in it stay in the library - only the collection goes."
            if count else "It is empty."
        )
        if not messagebox.askyesno(
            "Remove collection", f"Remove “{name}”?\n\n{detail}", icon="warning"
        ):
            return
        del by_id[name]
        if self._write_collection_ids(by_id):
            viewing = self._view == name
            self._reload_collections()
            if viewing:
                self._set_view("all")
            self.status_var.set(f"Removed {name}.")

    def _show_collection_menu(self, event, name):
        if self._is_shared(name):
            self._show_shared_menu(event, name)
            return
        self._popup(event.x_root, event.y_root, [
            ("item", "Rename…", "edit", lambda: self._rename_collection(name)),
            ("sep",),
            ("item", "Remove…", "trash", lambda: self._remove_collection(name), {"danger": True}),
        ])

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

    # ---------- shared collections ----------
    #
    # A shared collection is one small file on SharePoint, next to the
    # favourites: its name, who owns it, the email addresses it is shared
    # with, and its photos. Everyone listed sees it in their sidebar and can
    # add and remove photos; only the owner changes the name and the people.
    # Note that this decides what the app shows - anyone who can open the
    # photo library on SharePoint could also open these files.

    def _shared_cache_path(self):
        return os.path.join(self._sp_cache_dirs()[0], "shared_collections.json")

    def _shared_state(self):
        """{"me": email, "collections": [record, ...]} - the last synced
        copy, kept on disk so shared collections show at once on start."""
        if self._shared is None:
            data = self._read_json(self._shared_cache_path())
            records = [
                r for r in (data.get("collections") or [])
                if isinstance(r, dict) and isinstance(r.get("id"), str)
            ]
            self._shared = {"me": (data.get("me") or "").lower(), "collections": records}
        return self._shared

    def _save_shared_state(self):
        try:
            os.makedirs(os.path.dirname(self._shared_cache_path()), exist_ok=True)
            with open(self._shared_cache_path(), "w", encoding="utf-8") as f:
                json.dump(self._shared_state(), f, ensure_ascii=False)
        except OSError as exc:
            _log_reuse(f"SHARED COLLECTIONS SAVE FAILED: {exc!r}")

    def _shared_by_name(self):
        """{sidebar name: record}. A shared collection whose name is already
        used by one of this computer's own collections is shown as
        "Name (shared)" so the two never get mixed up."""
        local = {n.lower() for n in self._local_collection_ids()}
        out = {}
        for record in sorted(
            self._shared_state()["collections"], key=lambda r: (r.get("name") or "").lower()
        ):
            base = (record.get("name") or "Shared").strip() or "Shared"
            shown, n = base, 1
            while shown.lower() in local or shown in out:
                n += 1
                shown = f"{base} (shared)" if n == 2 else f"{base} (shared {n - 1})"
            out[shown] = record
        return out

    def _is_shared(self, name):
        return name in self._shared_by_name()

    def _shared_name_for(self, collection_id):
        for name, record in self._shared_by_name().items():
            if record.get("id") == collection_id:
                return name
        return None

    @staticmethod
    def _visible_to(record, me):
        if not me:
            return False
        members = [str(m).lower() for m in (record.get("members") or [])]
        return me == str(record.get("owner") or "").lower() or me in members

    def _sp_client(self, interactive):
        """The SharePoint connection. Background checks pass
        interactive=False: they use the saved sign-in or give up quietly,
        and never open a sign-in window on their own."""
        if interactive:
            return self._ensure_sp_client()
        with self._sp_lock:
            if self.sp_client is None:
                client = _sharepoint().SharePointClient()
                if not client.sign_in_silent():
                    return None
                self.sp_client = client
        return self.sp_client

    @staticmethod
    def _shared_location():
        import config

        return (
            getattr(config, "SHARED_SITE_PATH", config.SHAREPOINT_SITE_PATH),
            getattr(config, "SHARED_FOLDER_PATH", config.SHAREPOINT_SEARCH_ROOT_PATH),
        )

    def _shared_read(self, client, filename):
        site, folder = self._shared_location()
        raw = client.download_bytes_by_path(site, f"{folder}/{filename}")
        return json.loads(raw.decode("utf-8")) if raw else None

    def _shared_write(self, client, filename, record):
        site, folder = self._shared_location()
        client.ensure_folder(site, folder)
        client.upload_bytes_by_path(
            site, folder, filename, json.dumps(record, ensure_ascii=False).encode("utf-8")
        )

    def _sp_root(self, client):
        if self._sp_root_id is None:
            self._sp_root_id = client.get_search_root_item()["id"]
        return self._sp_root_id

    @staticmethod
    def _now_iso():
        return datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds")

    def _start_background_checks(self):
        self._sync_shared()
        self.root.after(3000, self._check_for_updates)
        self.root.after(SHARED_SYNC_MS, self._periodic_shared_sync)
        self.root.after(UPDATE_CHECK_MS, self._periodic_update_check)

    def _periodic_shared_sync(self):
        self._sync_shared()
        self.root.after(SHARED_SYNC_MS, self._periodic_shared_sync)

    def _periodic_update_check(self):
        self._check_for_updates()
        self.root.after(UPDATE_CHECK_MS, self._periodic_update_check)

    def _sync_shared(self, interactive=False, announce=False):
        """Reads every shared collection on SharePoint and keeps the ones
        this user owns or is listed in."""
        if self._shared_syncing:
            return
        self._shared_syncing = True
        if announce:
            self.status_var.set("Syncing shared collections…")

        def work():
            try:
                client = self._sp_client(interactive)
                if client is None:
                    return
                me = client.account_email()
                site, folder = self._shared_location()
                prefix = _sharepoint().SHARED_COLLECTION_PREFIX
                records = []
                for filename in client.list_names_by_path(site, folder, prefix):
                    record = self._shared_read(client, filename)
                    if isinstance(record, dict) and isinstance(record.get("id"), str) \
                            and self._visible_to(record, me):
                        records.append(record)
                self.event_queue.put(("shared_records", me, records, announce))
            except Exception as exc:
                _log_reuse(f"SHARED COLLECTIONS SYNC FAILED: {exc!r}")
                if announce:
                    self.event_queue.put((
                        "shared_status",
                        "Couldn't reach SharePoint to sync shared collections - try again in a moment.",
                    ))
            finally:
                self._shared_syncing = False

        threading.Thread(target=work, daemon=True).start()

    def _viewing_shared(self):
        record = self._shared_by_name().get(self._view)
        return (record.get("id"), list(record.get("items") or [])) if record else (None, None)

    def _after_shared_change(self, viewing_id, old_items):
        """Redraws the sidebar, and the open collection if it changed."""
        self._collections = self._load_collections()
        self._rebuild_collection_cards()
        if not viewing_id:
            return
        name = self._shared_name_for(viewing_id)
        if name is None:
            self._set_view("all")
            self.status_var.set("That shared collection was removed, or you are no longer in it.")
        elif name != self._view:
            self._set_view(name)
        else:
            record = self._shared_by_name()[name]
            if list(record.get("items") or []) != list(old_items or []):
                self._show_collection(name, keep_scroll=True)

    def _apply_shared_records(self, me, records, announce):
        viewing_id, old_items = self._viewing_shared()
        self._shared = {"me": (me or "").lower(), "collections": records}
        self._save_shared_state()
        self._after_shared_change(viewing_id, old_items)
        if announce:
            n = len(records)
            self.status_var.set(
                f"Shared collections are up to date - {n} shared with you." if n
                else "No shared collections yet - use + to create one."
            )

    def _apply_shared_upsert(self, record, me, message):
        viewing_id, old_items = self._viewing_shared()
        state = self._shared_state()
        if me:
            state["me"] = me.lower()
        state["collections"] = [r for r in state["collections"] if r.get("id") != record["id"]]
        if self._visible_to(record, state["me"]):
            state["collections"].append(record)
        self._save_shared_state()
        self._after_shared_change(viewing_id, old_items)
        if message:
            self.status_var.set(message)

    def _apply_shared_removed(self, collection_id, message):
        viewing_id, old_items = self._viewing_shared()
        state = self._shared_state()
        state["collections"] = [r for r in state["collections"] if r.get("id") != collection_id]
        self._save_shared_state()
        self._after_shared_change(viewing_id, old_items)
        if message:
            self.status_var.set(message)

    def _push_shared(self, snapshot, add=(), remove=(), name=None, members=None,
                     leave=False, delete=False, message=None):
        """Applies one change to a shared collection's file on SharePoint.

        Reads the current file first and applies only this change to it, so
        photos someone else added a moment ago are kept."""
        collection_id = snapshot["id"]

        def work():
            with self._shared_lock:
                try:
                    client = self._sp_client(True)
                    me = client.account_email()
                    filename = f"{_sharepoint().SHARED_COLLECTION_PREFIX}{collection_id}.json"
                    if delete:
                        site, folder = self._shared_location()
                        client.delete_by_path(site, f"{folder}/{filename}")
                        self.event_queue.put(("shared_removed", collection_id, message))
                        return
                    remote = self._shared_read(client, filename)
                    if not (isinstance(remote, dict) and remote.get("id") == collection_id):
                        self.event_queue.put((
                            "shared_removed", collection_id,
                            "That shared collection no longer exists on SharePoint.",
                        ))
                        return
                    items = [i for i in (remote.get("items") or []) if i and i not in remove]
                    for item_id in add:
                        if item_id not in items:
                            items.append(item_id)
                    remote["items"] = items
                    if name:
                        remote["name"] = name
                    if members is not None:
                        remote["members"] = members
                    if leave:
                        remote["members"] = [
                            m for m in (remote.get("members") or []) if str(m).lower() != me
                        ]
                    remote["updated"] = self._now_iso()
                    remote["updated_by"] = me
                    self._shared_write(client, filename, remote)
                    if leave:
                        self.event_queue.put(("shared_removed", collection_id, message))
                    else:
                        self.event_queue.put(("shared_upsert", remote, me, message))
                except Exception as exc:
                    _log_reuse(f"SHARED COLLECTION SAVE FAILED: {exc!r}")
                    self.event_queue.put((
                        "shared_status",
                        "Couldn't save that change to SharePoint, so others won't see it yet. "
                        "Check your connection and try again.",
                    ))

        threading.Thread(target=work, daemon=True).start()

    def _show_add_collection_menu(self, button):
        self._popup(button.winfo_rootx(), button.winfo_rooty() + button.winfo_height() + 4, [
            ("item", "New collection…", "newfolder", self._new_collection),
            ("item", "New shared collection…", "shared", self._new_shared_collection),
        ])

    def _show_shared_menu(self, event, name):
        record = self._shared_by_name().get(name) or {}
        me = self._shared_state()["me"]
        owner = str(record.get("owner") or "")
        is_owner = bool(me) and owner.lower() == me
        people = len(record.get("members") or [])
        items = [("header", f"Shared with {people} {'person' if people == 1 else 'people'}"
                  + ("" if is_owner else f" \u00b7 by {owner or 'someone else'}"))]
        if is_owner:
            items.append(("item", "Name and people…", "edit", lambda: self._manage_shared(name)))
        items += [
            ("item", "Sync now", "sync", lambda: self._sync_shared(interactive=True, announce=True)),
            ("sep",),
            ("item", "Delete shared collection…" if is_owner else "Leave this collection…",
             "trash", lambda: self._leave_or_delete_shared(name), {"danger": True}),
        ]
        self._popup(event.x_root, event.y_root, items)

    @staticmethod
    def _parse_emails(text):
        found, bad = [], []
        for piece in re.split(r"[\s,;]+", text or ""):
            piece = piece.strip().strip("<>").lower()
            if not piece:
                continue
            if EMAIL_PATTERN.match(piece):
                if piece not in found:
                    found.append(piece)
            else:
                bad.append(piece)
        return found, bad

    def _shared_dialog(self, title, action, name="", members=(), current=None):
        """Asks for a name and the people to share with.
        Returns (name, [emails]) or None if cancelled."""
        win = ctk.CTkToplevel(self.root)
        win.title(title)
        win.configure(fg_color=BG_COLOR)
        win.resizable(False, False)
        win.transient(self.root)
        win.after(250, lambda: _apply_icon(win))
        result = {}

        body = ctk.CTkFrame(win, fg_color="transparent")
        body.pack(fill=tk.BOTH, expand=True, padx=24, pady=(20, 18))
        ctk.CTkLabel(body, text=title, font=_font(17, "semibold"), text_color=TEXT_COLOR,
                     anchor="w").pack(fill=tk.X)
        ctk.CTkLabel(
            body, font=_font(12), text_color=MUTED_TEXT_COLOR, anchor="w", justify="left",
            wraplength=400,
            text=(f"Everyone you list sees this collection in their {APP_NAME} and can add "
                  "or remove photos. They need access to the photo library on SharePoint."),
        ).pack(fill=tk.X, pady=(4, 16))

        ctk.CTkLabel(body, text="Name", font=_font(12, "medium"), text_color=TEXT_SOFT,
                     anchor="w").pack(fill=tk.X)
        name_entry = ctk.CTkEntry(
            body, width=400, height=36, fg_color=SURFACE, border_color=BORDER, border_width=1,
            corner_radius=8, text_color=TEXT_COLOR, font=_font(13),
        )
        name_entry.insert(0, name)
        name_entry.pack(fill=tk.X, pady=(4, 14))

        ctk.CTkLabel(body, text="People - work email addresses, one per line",
                     font=_font(12, "medium"), text_color=TEXT_SOFT, anchor="w").pack(fill=tk.X)
        people = ctk.CTkTextbox(
            body, width=400, height=130, fg_color=SURFACE, border_color=BORDER, border_width=1,
            corner_radius=8, text_color=TEXT_COLOR, font=_font(13),
        )
        if members:
            people.insert("1.0", "\n".join(members))
        people.pack(fill=tk.X, pady=(4, 6))
        error = ctk.CTkLabel(body, text="", font=_font(12), text_color="#f87171", anchor="w",
                             justify="left", wraplength=400)
        error.pack(fill=tk.X)

        def submit():
            chosen = " ".join(name_entry.get().split())
            if not chosen:
                error.configure(text="Please give the collection a name.")
                return
            if len(chosen) > COLLECTION_NAME_MAX:
                error.configure(text=f"Please keep the name to {COLLECTION_NAME_MAX} characters.")
                return
            if chosen.lower() in RESERVED_COLLECTION_NAMES:
                error.configure(text=f"“{chosen}” is used by the app - please choose another name.")
                return
            taken = {n.lower() for n in self._local_collection_ids()}
            taken |= {
                (r.get("name") or "").lower() for r in self._shared_state()["collections"]
                if (r.get("name") or "") != current
            }
            if chosen.lower() in taken:
                error.configure(text=f"There is already a collection called “{chosen}”.")
                return
            emails, bad = self._parse_emails(people.get("1.0", "end"))
            if bad:
                error.configure(text="These don't look like email addresses: " + ", ".join(bad[:4]))
                return
            result["value"] = (chosen, emails)
            win.destroy()

        buttons = ctk.CTkFrame(body, fg_color="transparent")
        buttons.pack(fill=tk.X, pady=(12, 0))
        _make_primary_button(buttons, action, submit, width=110).pack(side=tk.RIGHT)
        _make_secondary_button(buttons, "Cancel", win.destroy, width=100).pack(
            side=tk.RIGHT, padx=(0, 8)
        )
        win.bind("<Escape>", lambda _e: win.destroy())
        try:
            win.grab_set()
        except tk.TclError:
            pass
        name_entry.focus_set()
        self.root.wait_window(win)
        return result.get("value")

    def _new_shared_collection(self):
        value = self._shared_dialog("New shared collection", "Create")
        if not value:
            return
        name, members = value
        self.status_var.set(f"Creating {name} on SharePoint…")

        def work():
            try:
                client = self._sp_client(True)
                me = client.account_email()
                now = self._now_iso()
                record = {
                    "version": 1,
                    "id": uuid.uuid4().hex[:16],
                    "name": name,
                    "owner": me,
                    "members": [m for m in members if m != me],
                    "items": [],
                    "created": now,
                    "updated": now,
                    "updated_by": me,
                }
                filename = f"{_sharepoint().SHARED_COLLECTION_PREFIX}{record['id']}.json"
                self._shared_write(client, filename, record)
                people = len(record["members"])
                if record["members"]:
                    self.event_queue.put(("shared_invite", name, list(record["members"])))
                self.event_queue.put((
                    "shared_upsert", record, me,
                    f"Created {name}, shared with {people} {'person' if people == 1 else 'people'}. "
                    "They see it after their next sync - or straight away if they click Sync.",
                ))
            except Exception as exc:
                _log_reuse(f"SHARED COLLECTION CREATE FAILED: {exc!r}")
                self.event_queue.put((
                    "shared_status",
                    "Couldn't create the shared collection on SharePoint - check your connection and try again.",
                ))

        threading.Thread(target=work, daemon=True).start()

    def _offer_invite(self, name, emails):
        """Sends the people just added a short invitation, straight from the
        user's own mailbox (it shows in their Sent Items) - one click, no
        email app. If the organisation doesn't let the app send mail, the
        same invitation opens ready-written in the email app instead."""
        if not emails:
            return
        emails = list(emails)
        who = emails[0] if len(emails) == 1 else f"{len(emails)} people"
        if not messagebox.askyesno(
            "Invite people",
            f"Send {who} an email that you've shared “{name}” with them?\n\n"
            "It goes out from your own mailbox.",
        ):
            return
        subject, body = self._invite_text(name)
        self.status_var.set(f"Sending the invitation to {who}…")

        def work():
            try:
                client = self._sp_client(True)
                if client is None:
                    raise RuntimeError("not signed in")
                client.send_mail(emails, subject, body)
                _log_reuse(f"INVITE SENT: {name!r} to {len(emails)} recipient(s)")
                self.event_queue.put(("shared_status", f"Invitation sent to {who}."))
            except Exception as exc:
                _log_reuse(f"INVITE MAIL FAILED: {exc!r}")
                self.event_queue.put(("invite_fallback", emails, subject, body))

        threading.Thread(target=work, daemon=True).start()

    def _invite_text(self, name):
        import config

        download = getattr(config, "DOWNLOAD_URL", "")
        # A link into the private GitHub repository opens for nobody else,
        # so it is only offered once DOWNLOAD_URL points somewhere shared.
        if "github.com" in download:
            download = ""
        subject = f"I've shared the photo collection “{name}” with you"
        body = (
            "Hi,\n\n"
            f"I've shared the photo collection “{name}” with you in {APP_NAME}, our app for "
            "searching the marketing photo library on SharePoint.\n\n"
            "To see it: open the app and click the sync button (↻) next to Collections - or just "
            f"restart it. “{name}” then appears under Collections, and you can add or remove "
            "photos by dragging them onto it.\n\n"
            + (f"Don't have {APP_NAME} yet? Download it here:\n{download}\n\n" if download
               else "Don't have the app yet? Just reply and I'll send it to you.\n\n")
            + "Best"
        )
        return subject, body

    def _invite_fallback(self, emails, subject, body):
        """Sending from the app wasn't possible: the invitation opens in the
        email app, ready to send."""
        self.status_var.set("Couldn't send from the app - the invitation opens in your email app instead.")
        url = "mailto:" + ",".join(emails) + "?" + urllib.parse.urlencode(
            {"subject": subject, "body": body}, quote_via=urllib.parse.quote
        )
        try:
            _open_external(url)
        except Exception as exc:
            _log_reuse(f"INVITE EMAIL FAILED: {exc!r}")
            messagebox.showinfo(APP_NAME, "Couldn't send the invitation - please let them know yourself.")

    def _manage_shared(self, shown_name):
        record = self._shared_by_name().get(shown_name)
        if not record:
            return
        me = self._shared_state()["me"]
        if str(record.get("owner") or "").lower() != me:
            messagebox.showinfo(
                APP_NAME,
                f"Only {record.get('owner') or 'the owner'} can change the name and people of this collection.",
            )
            return
        value = self._shared_dialog(
            "Shared collection", "Save", name=record.get("name") or "",
            members=record.get("members") or [], current=record.get("name"),
        )
        if not value:
            return
        name, members = value
        members = [m for m in members if m != me]
        added = [m for m in members if m not in [str(x).lower() for x in (record.get("members") or [])]]
        if added:
            self.root.after(300, lambda: self._offer_invite(name, added))
        self.status_var.set(f"Saving {name}…")
        self._push_shared(dict(record), name=name, members=members,
                          message=f"Saved {name} - shared with {len(members)} people.")

    def _leave_or_delete_shared(self, shown_name):
        record = self._shared_by_name().get(shown_name)
        if not record:
            return
        me = self._shared_state()["me"]
        if str(record.get("owner") or "").lower() == me:
            if not messagebox.askyesno(
                "Delete shared collection",
                f"Delete “{shown_name}” for everyone it is shared with?\n\n"
                "The photos stay in the library - only the collection goes.",
                icon="warning",
            ):
                return
            self._push_shared(dict(record), delete=True, message=f"Deleted {shown_name}.")
        else:
            if not messagebox.askyesno(
                "Leave collection",
                f"Leave “{shown_name}”? It disappears from your sidebar; the others keep it.",
            ):
                return
            self._push_shared(dict(record), leave=True, message=f"You left {shown_name}.")

    # ---------- local folder, support ----------

    def _show_local_folder_menu(self, anchor):
        try:
            enabled = str(self.index_button.cget("state")) == "normal"
        except Exception:
            enabled = False
        self._popup(anchor.winfo_rootx() + 12, anchor.winfo_rooty(), [
            ("item", "Choose folder…", "open", self._choose_folder),
            ("item", "Index folder", "index", self._start_indexing, {"enabled": enabled}),
        ])

    def _app_details(self):
        lines = [
            f"{APP_NAME} {APP_VERSION} (runtime {RUNTIME_VERSION}, "
            f"{'installed app' if getattr(sys, 'frozen', False) else 'from source'})",
            _system_description(),
            f"Signed in as: {self._shared_state().get('me') or 'unknown'}",
            f"Photos in library: {len(getattr(self.engine, 'paths', []) or []):,}",
        ]
        try:
            with open(CRASH_LOG_PATH, "r", encoding="utf-8") as f:
                tail = f.read()[-3000:]
            if tail.strip():
                lines += ["", "Last error log entries:", tail]
        except Exception:
            pass
        return "\n".join(lines)

    def _show_support_form(self):
        """A short form that emails the app's owner from the user's own
        Microsoft account - no email app needed."""
        import config

        to = getattr(config, "SUPPORT_EMAIL", "")
        win = ctk.CTkToplevel(self.root)
        win.title("Support")
        win.configure(fg_color=BG_COLOR)
        win.resizable(False, False)
        win.transient(self.root)
        win.after(250, lambda: _apply_icon(win))
        body = ctk.CTkFrame(win, fg_color="transparent")
        body.pack(fill=tk.BOTH, expand=True, padx=24, pady=(20, 18))
        ctk.CTkLabel(body, text="How can we help?", font=_font(17, "semibold"),
                     text_color=TEXT_COLOR, anchor="w").pack(fill=tk.X)
        ctk.CTkLabel(body, text=f"Your message goes straight to {APP_AUTHOR}, sent from your "
                     "Greenreb email so a reply comes back to you.", font=_font(12),
                     text_color=MUTED_TEXT_COLOR, anchor="w", justify="left",
                     wraplength=420).pack(fill=tk.X, pady=(4, 14))

        topic = tk.StringVar(value="Problem")
        ctk.CTkSegmentedButton(
            body, values=["Problem", "Idea", "Question"], variable=topic, height=32,
            font=_font(13, "medium"), fg_color=SURFACE, selected_color="#2a3830",
            selected_hover_color="#33443b", unselected_color=SURFACE,
            unselected_hover_color=SURFACE_HOVER, text_color=TEXT_COLOR,
        ).pack(anchor="w")
        ctk.CTkLabel(body, text="What's on your mind?", font=_font(12, "medium"),
                     text_color=TEXT_SOFT, anchor="w").pack(fill=tk.X, pady=(14, 0))
        message = ctk.CTkTextbox(
            body, width=420, height=150, fg_color=SURFACE, border_color=BORDER, border_width=1,
            corner_radius=8, text_color=TEXT_COLOR, font=_font(13), wrap="word",
        )
        message.pack(fill=tk.X, pady=(4, 8))
        details = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            body, text="Include app details (version, system, recent error log)",
            variable=details, font=_font(12), text_color=TEXT_SOFT, fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER, checkmark_color="#ffffff", border_color="#3f4a44",
            checkbox_width=18, checkbox_height=18,
        ).pack(anchor="w")
        error = ctk.CTkLabel(body, text="", font=_font(12), text_color="#f87171", anchor="w")
        error.pack(fill=tk.X, pady=(6, 0))

        def send():
            text = message.get("1.0", "end").strip()
            if not text:
                error.configure(text="Please write a few words first.")
                return
            first = text.splitlines()[0][:60]
            subject = f"[{APP_NAME} {topic.get()}] {first}"
            content = text
            if details.get():
                content += "\n\n---\n" + self._app_details()
            win.destroy()
            self.status_var.set("Sending your message…")

            def work():
                try:
                    client = self._sp_client(True)
                    client.send_mail(to, subject, content)
                    self.event_queue.put(("shared_status", f"Thanks - your message was sent to {APP_AUTHOR}."))
                except Exception as exc:
                    _log_reuse(f"SUPPORT MAIL FAILED: {exc!r}")
                    self.event_queue.put(("support_fallback", to, subject, content))

            threading.Thread(target=work, daemon=True).start()

        buttons = ctk.CTkFrame(body, fg_color="transparent")
        buttons.pack(fill=tk.X, pady=(10, 0))
        _make_primary_button(buttons, "Send", send, width=100).pack(side=tk.RIGHT)
        _make_secondary_button(buttons, "Cancel", win.destroy, width=100).pack(side=tk.RIGHT, padx=(0, 8))
        ctk.CTkButton(
            buttons, text="Take the tour", width=0, height=32, corner_radius=8,
            fg_color="transparent", hover_color=SURFACE_HOVER, text_color=MUTED_TEXT_COLOR,
            font=_font(12, "medium"), command=lambda: (win.destroy(), self._start_tour()),
        ).pack(side=tk.LEFT)
        win.bind("<Escape>", lambda _e: win.destroy())
        try:
            win.grab_set()
        except tk.TclError:
            pass
        message.focus_set()

    def _support_fallback(self, to, subject, content):
        """If sending from the app isn't allowed (some organisations need IT
        to approve it once), the same message opens in Outlook instead."""
        self.status_var.set("Couldn't send from the app - opening it in your email app instead.")
        url = f"mailto:{to}?" + urllib.parse.urlencode(
            {"subject": subject, "body": content[:1500]}, quote_via=urllib.parse.quote
        )
        try:
            _open_external(url)
        except Exception:
            messagebox.showinfo("Support", f"Please email {to} - your message couldn't be sent.")

    # ---------- tour ----------

    def _tour_path(self):
        return os.path.join(
            DATA_DIR, "tour.json"
        )

    def _tour_seen(self):
        if FIRST_RUN_PREVIEW:
            return False
        return bool(self._read_json(self._tour_path()).get("seen"))

    def _start_tour(self):
        if getattr(self, "_tour", None) is not None:
            return
        # A new computer has no photos until the library is loaded once, so
        # the tour ends on the button that does it - and then does it.
        empty = not len(self.engine.paths)
        steps = [
            (lambda: [self.search_frame, self.suggestion_bar], "How to browse through the library",
             ("Once your photos are loaded - that's the last step of this tour - type what you "
              "need in plain English, like “burger on a tray” or “friends cheering at night”, and "
              "press Enter. Below: your recent searches, and filters for landscape, portrait and "
              "people." if empty else
              "Type what you need in plain English - “burger on a tray”, “friends cheering at "
              "night” - and press Enter. Got a photo that's close? Use Find similar. Below: your "
              "recent searches, and filters for landscape, portrait and people.")),
            (lambda: [self.collections_heading,
                      getattr(self.sidebar, "_parent_frame", self.sidebar)], "Collections",
             "Your own folders for the photos you keep using. Drag any photo onto one to file "
             "it, or right-click a photo and pick Add to collection."),
            (lambda: [self.collection_add_button, self.shared_sync_button], "Shared collections",
             "Click + and choose New shared collection, then add your colleagues' email "
             "addresses - Topshot sends them the invite. Everyone sees the same photos and "
             "can add their own; ↻ fetches what others added. Right-click a shared collection "
             "to change who's in it."),
            (lambda: getattr(self.results_frame, "_parent_frame", self.results_frame), "Your photos",
             ("This is where your photos will be - the grey tiles are placeholders until the "
              "library is loaded in the last step. Then: double-click for a big preview, "
              "right-click to copy in full resolution, hover to star a favourite, and "
              + ("⌘" if IS_MAC else "Ctrl") + "-click or Shift-click to pick several."
              if empty else
              "Double-click for a big preview, right-click to copy it in full resolution, and "
              "hover to star a favourite. " + ("⌘" if IS_MAC else "Ctrl") + "-click or Shift-click "
              "to pick several - then add, copy or save them in one go.")),
            (lambda: [self.fav_card, self.history_card], "Favourites and history",
             "Every photo you starred and every search you ran - one click to get back to them."),
            (lambda: self.sharepoint_button,
             "Load your photos" if empty else "New photos, updates and help",
             ("Last step: click ✓ and Topshot loads the photo library from SharePoint. A "
              "Microsoft sign-in page opens in your browser - sign in with your Greenreb account, "
              "then come back here. The first load takes a few minutes and shows its progress at "
              "the top; after that, Topshot opens in seconds. Updates, settings and help live "
              "down here."
              if empty else
              "Check for new photos brings in the latest shots. The app keeps itself up to date - "
              "settings, support and this tour live down here. Enjoy!")),
        ]
        self._tour = Tour(self, steps, on_done=self._end_tour)

    def _end_tour(self, completed=True):
        self._tour = None
        if not completed:
            return  # it broke off on its own: shown again next start
        try:
            os.makedirs(os.path.dirname(self._tour_path()), exist_ok=True)
            with open(self._tour_path(), "w", encoding="utf-8") as f:
                json.dump({"seen": True, "version": APP_VERSION}, f)
        except OSError:
            pass
        # Still no photos (a first start): load the library now, whether the
        # tour was finished or skipped - nothing works until this has run once.
        if FIRST_RUN_PREVIEW:
            self.status_var.set(
                "First-run preview finished - on a new computer the library starts loading now. "
                "Start Topshot normally to get your photos back.")
            return
        if not len(self.engine.paths) and not self._busy:
            self.root.after(350, self._browse_sharepoint)

    # ---------- about ----------

    @staticmethod
    def _code_stats():
        """How many Python files and lines the running app is made of -
        counted from the source when it is there (running from source, or
        from a downloaded update), otherwise the figures for this release."""
        folder = os.path.dirname(os.path.abspath(__file__))
        files = lines = 0
        for name in os.listdir(folder):
            if name.endswith(".py"):
                try:
                    with open(os.path.join(folder, name), "r", encoding="utf-8") as f:
                        lines += sum(1 for _ in f)
                    files += 1
                except OSError:
                    pass
        if files:
            return files, lines
        return 9, 6700  # the .exe holds compiled code only

    def _show_about(self):
        win = ctk.CTkToplevel(self.root)
        win.title(f"About {APP_NAME}")
        win.configure(fg_color=BG_COLOR)
        win.resizable(False, False)
        win.transient(self.root)
        win.after(250, lambda: _apply_icon(win))
        body = ctk.CTkFrame(win, fg_color="transparent")
        body.pack(fill=tk.BOTH, expand=True, padx=28, pady=(24, 20))

        top = ctk.CTkFrame(body, fg_color="transparent")
        top.pack(fill=tk.X)
        if getattr(self, "_logo_image", None):
            ctk.CTkLabel(top, image=self._logo_image, text="").pack(side=tk.LEFT)
        names = ctk.CTkFrame(top, fg_color="transparent")
        names.pack(side=tk.LEFT, padx=(12, 0))
        ctk.CTkLabel(names, text=APP_NAME, font=_font(20, "semibold"), text_color=TEXT_COLOR,
                     anchor="w", height=26).pack(anchor="w")
        ctk.CTkLabel(names, text=f"Version {APP_VERSION}  ·  Marketing photo library",
                     font=_font(12), text_color=MUTED_TEXT_COLOR, anchor="w").pack(anchor="w")

        files, lines = self._code_stats()

        def block(heading, text, pady=(18, 0)):
            ctk.CTkLabel(body, text=heading.upper(), font=_font(11, "medium"),
                         text_color=FAINT_TEXT, anchor="w").pack(fill=tk.X, pady=pady)
            ctk.CTkLabel(body, text=text, font=_font(13), text_color=TEXT_SOFT, anchor="w",
                         justify="left", wraplength=420).pack(fill=tk.X, pady=(4, 0))

        block("Created by", f"{APP_AUTHOR}")
        block(
            "How it was built",
            f"Designed and developed by {APP_AUTHOR} with Claude, Anthropic's AI assistant, "
            f"as a coding partner: {files} Python files with about {lines:,} lines of code. "
            "It reads the marketing photo library on Microsoft SharePoint through Microsoft "
            "Graph, signing in with your Greenreb Microsoft account.",
        )
        block(
            "Built on",
            "OpenAI CLIP (open_clip) to understand what is in each photo · PyTorch · "
            "CustomTkinter · Pillow · Microsoft MSAL · the Inter typeface (SIL Open "
            "Font License) · Lucide icons (ISC License).",
        )
        ctk.CTkLabel(body, text=f"© {datetime.date.today().year} {APP_AUTHOR}",
                     font=_font(11), text_color=FAINT_TEXT, anchor="w").pack(fill=tk.X, pady=(18, 0))
        row = ctk.CTkFrame(body, fg_color="transparent")
        row.pack(fill=tk.X, pady=(14, 0))
        _make_secondary_button(row, "Close", win.destroy, width=100).pack(side=tk.RIGHT)
        _make_secondary_button(
            row, "Take the tour", lambda: (win.destroy(), self._start_tour()), width=130
        ).pack(side=tk.RIGHT, padx=(0, 8))
        win.bind("<Escape>", lambda _e: win.destroy())
        try:
            win.grab_set()
        except tk.TclError:
            pass

    # ---------- updates ----------
    #
    # The app's own code (about 4 MB) is updated separately from the .exe's
    # Python, libraries and model (about 600 MB). make_update.py publishes a
    # package and a latest.json to the releases folder on SharePoint; the
    # app downloads the package, checks it against the SHA-256 in
    # latest.json, unpacks it, and launcher.py runs it from the next start.

    def _updates_dir(self):
        return os.path.join(
            DATA_DIR, "app"
        )

    def _staged_update(self):
        try:
            with open(os.path.join(self._updates_dir(), "current.json"), "r", encoding="utf-8") as f:
                return str(json.load(f).get("version") or "")
        except Exception:
            return ""

    def _check_for_updates(self, interactive=False):
        if self._update_checking:
            return
        if self._update_ready:
            if interactive:
                self.status_var.set(f"Version {self._update_ready} is ready - restart to update.")
            return
        self._update_checking = True
        if interactive:
            self.status_var.set("Checking for updates…")
            self._set_update_card("checking")

        def work():
            try:
                import config

                client = self._sp_client(interactive)
                if client is None:
                    return
                raw = client.download_bytes_by_path(
                    config.UPDATE_SITE_PATH, f"{config.UPDATE_FOLDER_PATH}/latest.json"
                )
                if raw is None:
                    self.event_queue.put(("update_uptodate", interactive))
                    return
                manifest = json.loads(raw.decode("utf-8"))
                version = str(manifest["version"])
                changes = [str(c) for c in (manifest.get("changes") or [])][:8]
                if version_tuple(version) <= version_tuple(APP_VERSION):
                    self.event_queue.put(("update_uptodate", interactive))
                    return
                if int(manifest.get("runtime", -1)) != RUNTIME_VERSION:
                    self.event_queue.put(("update_full", version, manifest.get("notes") or "", changes))
                    return
                if self._staged_update() != version:
                    package = client.download_bytes_by_path(
                        config.UPDATE_SITE_PATH, f"{config.UPDATE_FOLDER_PATH}/{manifest['package']}"
                    )
                    if package is None:
                        raise ValueError("the update package is missing")
                    if hashlib.sha256(package).hexdigest() != str(manifest["sha256"]).lower():
                        raise ValueError("the update package does not match its checksum")
                    self._stage_update(version, package)
                self.event_queue.put(("update_ready", version, changes))
            except Exception as exc:
                _log_reuse(f"UPDATE CHECK FAILED: {exc!r}")
                if interactive:
                    self.event_queue.put(("update_status", "Couldn't check for updates - try again later."))
                    self.event_queue.put(("update_card", "current"))
            finally:
                self._update_checking = False

        threading.Thread(target=work, daemon=True).start()

    def _stage_update(self, version, package):
        """Unpacks an update next to the others and points the launcher at
        it. Refuses anything that isn't plainly one of our packages."""
        root = self._updates_dir()
        partial = os.path.join(root, f"{version}.partial")
        shutil.rmtree(partial, ignore_errors=True)
        os.makedirs(partial)
        try:
            self._unpack_update(package, partial)
        except Exception:
            shutil.rmtree(partial, ignore_errors=True)
            raise
        final = os.path.join(root, version)
        shutil.rmtree(final, ignore_errors=True)
        os.replace(partial, final)
        with open(os.path.join(root, "current.json"), "w", encoding="utf-8") as f:
            json.dump({"version": version, "runtime": RUNTIME_VERSION}, f)
        # Older downloaded versions are no longer needed.
        for entry in os.listdir(root):
            path = os.path.join(root, entry)
            if os.path.isdir(path) and entry not in (version, APP_VERSION):
                shutil.rmtree(path, ignore_errors=True)

    @staticmethod
    def _unpack_update(package, partial):
        with zipfile.ZipFile(io.BytesIO(package)) as z:
            for info in z.infolist():
                name = info.filename.replace("\\", "/")
                if name.endswith("/"):
                    continue
                parts = name.split("/")
                if name.startswith("/") or ".." in parts or ":" in name:
                    raise ValueError(f"unsafe path in update: {name}")
                if os.path.splitext(name)[1].lower() not in UPDATE_FILE_TYPES:
                    raise ValueError(f"unexpected file in update: {name}")
                dest = os.path.join(partial, *parts)
                os.makedirs(os.path.dirname(dest), exist_ok=True)
                with z.open(info) as source, open(dest, "wb") as out:
                    shutil.copyfileobj(source, out)
        for module in APP_MODULES:
            if not os.path.isfile(os.path.join(partial, f"{module}.py")):
                raise ValueError(f"update is missing {module}.py")

    def _set_update_card(self, state, version=None, changes=None, note=""):
        """current / checking / uptodate / ready / download."""
        version = version or APP_VERSION
        # The one-line card (short screens) packs its button row away; it is
        # put back first, so the button and the text below can find their place.
        actions = getattr(self, "_update_actions", None)
        if actions is not None and not actions.winfo_manager():
            actions.pack(fill=tk.X, padx=8, pady=(4, 7))
        self.update_button.pack_forget()
        pill, pill_colour, border = "", FAINT_TEXT, BORDER
        title, detail = f"{APP_NAME} {APP_VERSION}", ""
        if state == "checking":
            pill = "Checking…"
        elif state == "uptodate":
            pill, pill_colour = "Up to date", SOURCE_DOT
        elif state == "ready":
            title, border = f"Update ready · {version}", PRIMARY
            detail = " · ".join((changes or [])[:2]) or "A newer version is ready."
            self.update_button.configure(text="Restart to update", command=self._restart_app)
            self.update_button.pack(side=tk.LEFT, padx=(4, 6), before=self.whats_new_link)
        elif state == "download":
            title, border = f"Version {version} available", PRIMARY
            detail = note or "This one needs a new download."
            self.update_button.configure(text="Download", command=self._open_download_page)
            self.update_button.pack(side=tk.LEFT, padx=(4, 6), before=self.whats_new_link)
        self._update_state = state
        if changes is not None:
            self._update_info = {"version": version, "changes": list(changes)}
        self.update_title.configure(text=title)
        self.update_pill.configure(text=pill, text_color=pill_colour)
        self.update_detail.configure(text=detail)
        if detail:
            self.update_detail.pack(fill=tk.X, padx=12, pady=(1, 0), before=self.update_button.master)
        else:
            self.update_detail.pack_forget()
        previous = self.update_card.cget("border_color")
        if state in ("ready", "download") and previous != border:
            # Draws the eye once: the border warms to green.
            _animate(self.root, 500, lambda p: self.update_card.configure(
                border_color=_mix(BORDER, border, p)))
        else:
            self.update_card.configure(border_color=border)
        self._apply_update_card_density()

    def _show_update_notice(self, version, restart, note="", changes=None):
        if restart:
            self._update_ready = version
            self._set_update_card("ready", version, changes)
        else:
            self._set_update_card("download", version, changes, note)

    def _open_download_page(self):
        import config

        url = getattr(config, "DOWNLOAD_URL", "")
        if url:
            try:
                _open_external(url)
            except Exception:
                pass

    def _show_whats_new(self, fresh=False):
        """The version and what changed, on a card over the dimmed window.

        Shown by itself once after an update (fresh=True), and from the
        What's new link any time - then it describes an update that is
        waiting, if there is one. Rare, so it can take a moment: the card
        rises and fades in (ease-out, 260 ms), the lines follow one after
        another, and it leaves faster than it came (150 ms).
        """
        if getattr(self, "_whats_new_win", None) is not None:
            return
        info = (self._update_info if (self._update_info and not fresh)
                else {"version": APP_VERSION, "changes": WHATS_NEW})
        upcoming = info["version"] != APP_VERSION
        reduced = self._motion_reduced()
        root = self.root
        root.update_idletasks()
        rx, ry = root.winfo_rootx(), root.winfo_rooty()
        rw, rh = root.winfo_width(), root.winfo_height()

        scrim = tk.Toplevel(root)
        scrim.overrideredirect(True)
        scrim.configure(bg="#000000")
        try:
            scrim.attributes("-alpha", 0.0)
            scrim.transient(root)
        except tk.TclError:
            pass
        scrim.geometry(f"{rw}x{rh}+{rx}+{ry}")

        width = 440
        win = ctk.CTkToplevel(root)
        win.withdraw()
        win.overrideredirect(True)
        win.configure(fg_color=SURFACE)
        try:
            win.attributes("-alpha", 0.0)
            win.transient(root)
        except tk.TclError:
            pass
        self._whats_new_win = win
        body = ctk.CTkFrame(win, fg_color="transparent", width=width)
        body.pack(fill=tk.BOTH, expand=True, padx=28, pady=(24, 22))

        head = ctk.CTkFrame(body, fg_color="transparent")
        head.pack(fill=tk.X)
        try:
            mark = _load_mark(44)
            ctk.CTkLabel(head, text="", image=ctk.CTkImage(mark, mark, size=(44, 44))).pack(
                side=tk.LEFT, padx=(0, 14))
        except Exception:
            pass
        words = ctk.CTkFrame(head, fg_color="transparent")
        words.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ctk.CTkLabel(words, text="WHAT'S NEW", font=_font(11, "semibold"),
                     text_color=SPLASH_ACCENT, anchor="w").pack(fill=tk.X)
        ctk.CTkLabel(words, text=f"{APP_NAME} {info['version']}", font=_font(22, "semibold"),
                     text_color=TEXT_COLOR, anchor="w").pack(fill=tk.X)
        ctk.CTkLabel(words, text=("Ready to install - restart to switch." if upcoming
                                  else "Updated just now." if fresh else "The version you are using."),
                     font=_font(12), text_color=MUTED_TEXT_COLOR, anchor="w").pack(fill=tk.X)

        ctk.CTkFrame(body, height=1, fg_color=BORDER).pack(fill=tk.X, pady=(18, 12))
        lines = []
        for text in info["changes"] or ["Improvements and fixes."]:
            row = ctk.CTkFrame(body, fg_color="transparent")
            row.pack(fill=tk.X, pady=4)
            ctk.CTkLabel(row, text="", image=_icon("done", 14, SOURCE_DOT), width=18).pack(
                side=tk.LEFT, anchor="n", pady=(2, 0))
            label = ctk.CTkLabel(row, text=text, font=_font(13),
                                 text_color=TEXT_SOFT if reduced else SURFACE,
                                 anchor="w", justify="left", wraplength=width - 100)
            label.pack(side=tk.LEFT, padx=(10, 0), fill=tk.X)
            lines.append(label)

        buttons = ctk.CTkFrame(body, fg_color="transparent")
        buttons.pack(fill=tk.X, pady=(20, 0))
        closing = {"done": False}

        def close(then=None):
            if closing["done"]:
                return
            closing["done"] = True

            def step(p):
                try:
                    win.attributes("-alpha", 1.0 - p)
                    scrim.attributes("-alpha", 0.4 * (1.0 - p))
                    win.geometry(f"+{x}+{y + round(6 * p)}")
                except tk.TclError:
                    pass

            def gone():
                for w in (win, scrim):
                    try:
                        w.destroy()
                    except tk.TclError:
                        pass
                self._whats_new_win = None
                if then:
                    then()

            _animate(root, 0 if reduced else 150, step, done=gone, ease=EASE_OUT)

        if upcoming and self._update_ready:
            _make_primary_button(buttons, "Restart to update",
                                 lambda: close(self._restart_app), width=150).pack(side=tk.RIGHT)
            _make_secondary_button(buttons, "Later", close, width=90).pack(
                side=tk.RIGHT, padx=(0, 8))
        else:
            _make_primary_button(buttons, "Got it", close, width=110).pack(side=tk.RIGHT)

        win.update_idletasks()
        h = win.winfo_reqheight()
        x = rx + max(0, (rw - width) // 2)
        y = ry + max(0, (rh - h) // 2 - 12)
        rise = 0 if reduced else 14
        win.geometry(f"{width}x{h}+{x}+{y + rise}")
        win.deiconify()
        win.update_idletasks()
        _round_corners(win, border=MENU_BORDER)
        try:
            scrim.lift()  # over the window, under the card
        except tk.TclError:
            pass
        win.lift()
        win.focus_force()
        win.bind("<Escape>", lambda _e: close())
        win.bind("<Return>", lambda _e: close())
        scrim.bind("<Button-1>", lambda _e: close())

        def enter(p):
            try:
                win.attributes("-alpha", p)
                scrim.attributes("-alpha", 0.4 * p)
                win.geometry(f"+{x}+{y + round(rise * (1 - p))}")
            except tk.TclError:
                pass

        _animate(root, 120 if reduced else 260, enter, ease=EASE_OUT)
        if not reduced:
            for i, label in enumerate(lines):
                def fade(label=label):
                    _animate(root, 220, lambda p: label.winfo_exists() and label.configure(
                        text_color=_mix(SURFACE, TEXT_SOFT, p)), ease=EASE_OUT)
                root.after(140 + i * 55, fade)

    def _restart_app(self):
        """Starts the app again through the launcher, which picks up the
        downloaded version, then closes this window."""
        if getattr(sys, "frozen", False):
            command = [sys.executable]
        else:
            first = os.path.abspath(sys.argv[0]) if sys.argv and sys.argv[0] else ""
            if os.path.basename(first).lower() != "launcher.py":
                first = os.path.join(os.path.dirname(os.path.abspath(__file__)), "launcher.py")
            command = [sys.executable, first]
        command.append("--after-update")
        try:
            os.remove(_handover_path())
        except OSError:
            pass
        try:
            subprocess.Popen(command, close_fds=True)
        except OSError as exc:
            messagebox.showerror(APP_NAME, f"Couldn't restart: {exc}")
            return
        # The same intro as at start-up, filling the screen, while the new
        # version loads behind it. The new version opens its own identical
        # screen on top and says so; then this window goes.
        version = self._update_ready or "the new version"
        try:
            _fine_timer(True)
            overlay = BootSplash(self.root, full=True)
            overlay.show(on_shown=lambda: overlay.set_progress(0.35))
            overlay.set_status(f"Updating to version {version}…")
        except Exception:
            overlay = None
        started = time.perf_counter()

        def wait_for_handover():
            elapsed = time.perf_counter() - started
            if os.path.exists(_handover_path()) or elapsed > 60:
                if overlay is not None:
                    overlay.set_progress(1.0)
                self.root.after(300, self.root.destroy)
                return
            if overlay is not None and elapsed > 1.5:
                # Creeps on while waiting, so it never looks stuck.
                overlay.set_progress(min(0.9, 0.35 + elapsed / 20))
            self.root.after(150, wait_for_handover)

        self.root.after(400, wait_for_handover)

    # ---------- More menu ----------

    def _show_more_menu(self, anchor):
        def later(action):
            # A menu item opening another menu: after this click has finished,
            # or the click itself would count as "outside" and close it.
            return lambda: self.root.after(80, action)

        self._popup(anchor.winfo_rootx() + 12, anchor.winfo_rooty(), [
            ("item", "Check for updates", "update", lambda: self._check_for_updates(interactive=True)),
            ("item", "Local folder…", "local", later(lambda: self._show_local_folder_menu(anchor))),
            ("item", "Settings", "settings", self._open_settings),
            ("sep",),
            ("item", "Take the tour", "tour", later(self._start_tour)),
            ("item", "Support", "support", self._show_support_form),
            ("item", f"About {APP_NAME}", "info", self._show_about),
        ])

    # ---------- filters ----------
    #
    # Two small pairs on the right of the Recent row: landscape / portrait,
    # and with / without people. They narrow whatever is on screen - a
    # search, a collection, favourites - and All photos samples the whole
    # library through them. Click an active one again to switch it off.

    def _filters_state(self):
        f = self.__dict__.get("_filters")
        if f is None:
            f = self._filters = {"shape": None, "people": None}
        return f

    def _filters_active(self):
        f = self._filters_state()
        return bool(f["shape"] or f["people"])

    def _build_filter_chips(self, bar):
        f = self._filters_state()
        box = ctk.CTkFrame(bar, fg_color="transparent")
        box.pack(side=tk.RIGHT, padx=(8, 0))
        if not hasattr(self, "_filter_count_var"):
            self._filter_count_var = tk.StringVar(value="")
        ctk.CTkLabel(box, textvariable=self._filter_count_var, text_color=FAINT_TEXT,
                     font=_font(12)).pack(side=tk.LEFT, padx=(0, 6))

        def chip(text, icon, group, value, gap):
            on = f[group] == value
            ink = TEXT_COLOR if on else MUTED_TEXT_COLOR
            ctk.CTkButton(
                box, text=text, image=_icon(icon, 13, ink), compound="left", width=0,
                height=28, corner_radius=14, font=_font(12),
                fg_color=NAV_SELECTED if on else "transparent",
                hover_color=NAV_SELECTED if on else SURFACE_HOVER, text_color=ink,
                border_width=1, border_color=PRIMARY if on else BORDER,
                command=lambda: self._toggle_filter(group, value),
            ).pack(side=tk.LEFT, padx=(gap, 3))

        chip(" Landscape", "landscape", "shape", "landscape", 3)
        chip(" Portrait", "portrait", "shape", "portrait", 0)
        chip(" People", "people", "people", "with", 10)
        chip(" No people", "nopeople", "people", "without", 0)

    def _toggle_filter(self, group, value):
        f = self._filters_state()
        turning_on = f[group] != value
        if (group == "people" and turning_on and self.engine.model is None
                and self.__dict__.get("_person_vec") is None and not self._load_concept_cache()):
            self._warm_model()
            self.status_var.set(
                "The search engine is still starting - try the people filter again in a few seconds.")
            return
        f[group] = value if turning_on else None
        self._refresh_suggestions()
        if getattr(self, "_sample_view", False):
            self._show_all_photos()
        else:
            self._render_results(self._last_results)
        names = {"landscape": "landscape", "portrait": "portrait",
                 "with": "with people", "without": "without people"}
        on = [names[v] for v in (f["shape"], f["people"]) if v]
        self.status_var.set(f"Showing {' and '.join(on)} photos only - click a filter again to switch it off."
                            if on else "Filters off - showing everything again.")

    def _path_index(self):
        paths = self.engine.paths
        cache = self.__dict__.get("_path_index_cache")
        if cache is None or cache[0] is not paths or cache[1] != len(paths):
            cache = self._path_index_cache = (paths, len(paths), {p: i for i, p in enumerate(paths)})
        return cache[2]

    def _people_flags(self):
        """Per photo: is anyone in it - seen by either the "a person" check or
        the "guests having fun" one, so a lone golfer and a table of friends
        both count. None while that can't be told yet."""
        try:
            if self.__dict__.get("_person_vec") is None or self._people_vec is None:
                if not self._load_concept_cache():
                    if self.engine.model is None:
                        return None
                    embed = self.engine.embed_concept
                    self._person_vec, self._nobody_vec = embed(PERSON_PROMPTS), embed(NOBODY_PROMPTS)
                    self._people_vec, self._empty_vec = embed(PEOPLE_PROMPTS), embed(EMPTY_PROMPTS)
                    self._save_concept_cache()
            score = self.engine.score_against
            return (((score(self._person_vec) - score(self._nobody_vec)) > 0)
                    | ((score(self._people_vec) - score(self._empty_vec)) > 0))
        except Exception:
            return None

    def _concept_cache_path(self):
        return os.path.join(DATA_DIR, "concepts.npz")

    @staticmethod
    def _concept_prompts_key():
        return repr((PERSON_PROMPTS, NOBODY_PROMPTS, PEOPLE_PROMPTS, EMPTY_PROMPTS))

    def _load_concept_cache(self):
        """The "is anyone in it" directions, kept from an earlier run so the
        people filter works the moment the app opens, before the search
        engine has loaded."""
        try:
            data = np.load(self._concept_cache_path())
            if str(data["tag"]) != MODEL_TAG or str(data["prompts"]) != self._concept_prompts_key():
                return False
            self._person_vec, self._nobody_vec = data["person"], data["nobody"]
            self._people_vec, self._empty_vec = data["people"], data["empty"]
            return True
        except Exception:
            return False

    def _save_concept_cache(self):
        try:
            os.makedirs(DATA_DIR, exist_ok=True)
            with open(self._concept_cache_path(), "wb") as f:
                np.savez(f, tag=MODEL_TAG, prompts=self._concept_prompts_key(),
                         person=self._person_vec, nobody=self._nobody_vec,
                         people=self._people_vec, empty=self._empty_vec)
        except Exception:
            pass

    def _passes_filters(self, idx, people):
        f = self._filters_state()
        if idx is None:
            return False
        if f["shape"]:
            if self._is_landscape(idx) != (f["shape"] == "landscape"):
                return False
        if f["people"] and people is not None and idx < len(people):
            if bool(people[idx]) != (f["people"] == "with"):
                return False
        return True

    def _apply_filters(self, results):
        if not results or not self._filters_active():
            return results
        index = self._path_index()
        people = self._people_flags() if self._filters_state()["people"] else None
        return [r for r in results if self._passes_filters(index.get(r[0]), people)]

    def _filtered_sample_indices(self):
        """A cross-section of the whole library through the filters, checking
        only as many photos as it takes to fill the opening screen."""
        total = len(self.engine.paths)
        if not total:
            return []
        people = self._people_flags() if self._filters_state()["people"] else None
        step = max(1, total // (ALL_PHOTOS_PREVIEW * 6))
        picked = []
        for offset in range(step):
            for i in range(offset, total, step):
                if self._passes_filters(i, people):
                    picked.append(i)
                    if len(picked) >= ALL_PHOTOS_PREVIEW:
                        return picked
        return picked

    def _show_filter_count(self, total, shown):
        var = getattr(self, "_filter_count_var", None)
        if var is None:
            return
        if self._filters_active() and not getattr(self, "_sample_view", False) and total:
            var.set(f"{shown} of {total}")
        else:
            var.set("")
        if self._filters_active() and total and not shown:
            self.status_var.set("No photos here match these filters - switch one off to see more.")

    # ---------- picking several photos ----------

    def _bind_selection_keys(self):
        self._sel = {}
        self._sel_anchor = None
        self._tile_by_pos = {}
        self._shown_results = []

        def escape(_e):
            if getattr(self, "_menu", None) is None and self._sel:
                self._clear_selection()

        def select_all(_e):
            try:
                if self.root.focus_get() is self.query_entry or str(self.root.focus_get()).startswith(
                        str(self.query_entry)):
                    return None  # Ctrl+A in the search box selects its text
            except Exception:
                pass
            for pos in range(len(self._shown_results)):
                self._select_pos(pos, True)
            self._update_selection_bar()
            return "break"

        self.root.bind("<Escape>", escape, add="+")
        self.root.bind("<Command-a>" if IS_MAC else "<Control-a>", select_all, add="+")
        self._bind_shortcuts()

    def _typing(self):
        """Whether the keyboard is in a text field, where keys keep their
        usual meaning (Ctrl+C copies text, / types a slash)."""
        try:
            return isinstance(self.root.focus_get(), (tk.Entry, tk.Text))
        except Exception:
            return False

    def _bind_shortcuts(self):
        """Ctrl+F (Cmd+F) or / jumps to the search box, Ctrl+C copies the
        picked photos or the one under the pointer, Esc empties the search
        box. Instant, without animation - they are used all day."""
        mod = "Command" if IS_MAC else "Control"

        def find(_e=None):
            try:
                self.query_entry.focus_set()
                inner = getattr(self.query_entry, "_entry", self.query_entry)
                inner.select_range(0, "end")
                inner.icursor("end")
            except Exception:
                pass
            return "break"

        def slash(_e):
            return None if self._typing() else find()

        def copy(_e):
            if self._typing():
                return None
            if getattr(self, "_sel", None):
                self._copy_selection()
                return "break"
            label = getattr(self, "_hover_label", None)
            try:
                if label is not None and label.winfo_exists():
                    path, _size, _key = label.tile
                    self._copy_to_clipboard(path, getattr(label, "meta", None))
                    return "break"
            except Exception:
                pass
            return None

        def clear_search(_e):
            if self.query_var.get():
                self.query_var.set("")
                return "break"
            return None

        for key in ("f", "F"):
            self.root.bind(f"<{mod}-{key}>", find, add="+")
        self.root.bind("<slash>", slash, add="+")
        for key in ("c", "C"):
            self.root.bind(f"<{mod}-{key}>", copy, add="+")
        self.query_entry.bind("<Escape>", clear_search, add="+")

    def _select_click(self, pos, mode):
        self._suppress_release = True
        if mode == "range" and self._sel_anchor is not None:
            lo, hi = sorted((self._sel_anchor, pos))
            for p in range(lo, hi + 1):
                self._select_pos(p, True)
        else:
            self._select_pos(pos, pos not in self._sel)
            self._sel_anchor = pos
        self._update_selection_bar()
        return "break"

    def _select_pos(self, pos, on):
        shown = getattr(self, "_shown_results", None) or []
        if not 0 <= pos < len(shown):
            return
        if on:
            path, _score, meta = shown[pos]
            self._sel[pos] = (path, meta)
        else:
            self._sel.pop(pos, None)
        label = self._tile_by_pos.get(pos)
        if label is not None:
            label.selected = on
            self._paint_tile(label)

    def _paint_tile(self, label):
        try:
            if getattr(label, "dragging", False):
                return
            if getattr(label, "selected", False):
                if getattr(label, "sel_image", None) is None:
                    path, size, key = label.tile
                    label.sel_image = self._tile_photo(path, size, key in self._render_favs,
                                                       selected=True)
                label.configure(image=label.sel_image)
            else:
                label.configure(image=label.image)
        except Exception:
            pass

    def _selection_items(self):
        return [self._sel[p] for p in sorted(self._sel)]

    def _clear_selection(self, animate=True):
        sel = getattr(self, "_sel", None)
        if sel is None:
            return
        for pos in list(sel):
            label = self._tile_by_pos.get(pos)
            if label is not None:
                label.selected = False
                self._paint_tile(label)
        sel.clear()
        self._sel_anchor = None
        self._update_selection_bar(animate=animate)

    def _drag_group(self, label):
        pos = getattr(label, "pos", None)
        sel = getattr(self, "_sel", None) or {}
        if pos in sel and len(sel) > 1:
            return self._selection_items()
        return None

    def _selected_look(self, tile, size):
        """Picked: a soft green veil, a green ring and a tick, so a screen of
        picks reads at a glance."""
        w, h = size
        k = 3
        veil = Image.new("RGBA", size, _hex_rgb(PRIMARY) + (0,))
        veil.putalpha(self._tile_mask(size).point(lambda a: a * 46 // 255))
        tile = Image.alpha_composite(tile, veil)
        ring = Image.new("RGBA", (w * k, h * k), (0, 0, 0, 0))
        ImageDraw.Draw(ring).rounded_rectangle(
            (k, k, w * k - k - 1, h * k - k - 1), radius=TILE_RADIUS * k,
            outline=_hex_rgb(PRIMARY_HOVER) + (255,), width=3 * k)
        tile.alpha_composite(ring.resize(size, Image.LANCZOS))
        tile.alpha_composite(_selection_badge(), (8, 8))
        return tile

    def _with_count_badge(self, img, count):
        """The carried photo, with how many are coming along."""
        base = img.convert("RGBA").copy()
        w, h = base.size
        d = max(26, w // 6)
        k = 4
        badge = Image.new("RGBA", (d * k, d * k), (0, 0, 0, 0))
        draw = ImageDraw.Draw(badge)
        draw.ellipse((0, 0, d * k - 1, d * k - 1), fill=_hex_rgb(PRIMARY) + (255,))
        try:
            font = ImageFont.truetype(_resource_path("assets", "fonts", "Inter-SemiBold.ttf"),
                                      int(d * k * 0.48))
        except Exception:
            font = ImageFont.load_default()
        text = str(count)
        l, t, r, b = draw.textbbox((0, 0), text, font=font)
        draw.text(((d * k - (r - l)) / 2 - l, (d * k - (b - t)) / 2 - t), text, font=font,
                  fill=(255, 255, 255, 255))
        base.alpha_composite(badge.resize((d, d), Image.LANCZOS), (w - d - 6, 6))
        return base

    # the bar that appears while photos are picked

    def _selection_bar_parts(self):
        bar = self.__dict__.get("_sel_bar")
        if bar is not None:
            return bar
        holder = tk.Frame(self.root, bg=BG_COLOR, bd=0, highlightthickness=0)
        card = ctk.CTkFrame(holder, fg_color=SURFACE, corner_radius=14, border_width=1,
                            border_color=MENU_BORDER, bg_color=BG_COLOR)
        card.pack(fill=tk.BOTH, expand=True)
        row = ctk.CTkFrame(card, fg_color="transparent")
        row.pack(padx=10, pady=8)
        count = ctk.CTkLabel(row, text="", font=_font(13, "semibold"), text_color=TEXT_COLOR)
        count.pack(side=tk.LEFT, padx=(8, 12))

        def action(text, icon, command):
            button = ctk.CTkButton(
                row, text=text, image=_icon(icon, 14, TEXT_SOFT), compound="left", width=0,
                height=32, corner_radius=8, fg_color="transparent", hover_color=SURFACE_HOVER,
                text_color=TEXT_SOFT, font=_font(12, "medium"), command=command)
            button.pack(side=tk.LEFT, padx=2)
            return button

        collect = action(" Add to…", "folder", None)
        collect.configure(command=lambda: self._selection_collect_menu(collect))
        action(" Copy", "copy", self._copy_selection)
        action(" Save…", "save", self._save_selection)
        fav = action(" Favourite", "star", self._favourite_selection)
        ctk.CTkFrame(row, width=1, height=20, fg_color=BORDER).pack(side=tk.LEFT, padx=6)
        ctk.CTkButton(
            row, text="", image=_icon("close", 14, MUTED_TEXT_COLOR), width=32, height=32,
            corner_radius=8, fg_color="transparent", hover_color=SURFACE_HOVER,
            command=self._clear_selection,
        ).pack(side=tk.LEFT)
        bar = self._sel_bar = {"holder": holder, "count": count, "collect": collect, "fav": fav,
                               "shown": False, "gen": 0, "offset": 0}
        return bar

    def _selection_bar_spot(self, bar):
        holder = bar["holder"]
        holder.update_idletasks()
        w = holder.winfo_reqwidth()
        h = holder.winfo_reqheight()
        area = getattr(self.results_frame, "_parent_frame", self.results_frame)
        ax = area.winfo_rootx() - self.root.winfo_rootx()
        ay = area.winfo_rooty() - self.root.winfo_rooty()
        x = ax + max(0, (area.winfo_width() - w) // 2)
        y = min(self.root.winfo_height(), ay + area.winfo_height()) - h - 20
        return x, y, w, h

    def _place_selection_bar(self):
        bar = self.__dict__.get("_sel_bar")
        if not bar or not bar["shown"]:
            return
        try:
            x, y, w, h = self._selection_bar_spot(bar)
            bar["holder"].place(x=x, y=y + bar["offset"], width=w, height=h)
            bar["holder"].lift()
        except tk.TclError:
            pass

    def _update_selection_bar(self, animate=True):
        n = len(getattr(self, "_sel", None) or {})
        bar = self.__dict__.get("_sel_bar")
        if n == 0:
            if bar and bar["shown"]:
                self._slide_selection_bar(False, animate)
            return
        bar = self._selection_bar_parts()
        bar["count"].configure(text=f"{n} selected")
        in_collection = self._view in self._collection_names()
        bar["collect"].configure(text=" Move to…" if in_collection else " Add to…")
        bar["fav"].configure(text=" Remove from favourites" if self._view == "favourites"
                             else " Favourite")
        if not bar["shown"]:
            self._slide_selection_bar(True, animate)
        else:
            self._place_selection_bar()

    def _slide_selection_bar(self, show, animate=True):
        """Rises from the bottom edge and leaves the same way: in with a
        quick ease-out, out a little faster."""
        bar = self._selection_bar_parts()
        bar["gen"] += 1
        gen = bar["gen"]
        try:
            _x, _y, _w, h = self._selection_bar_spot(bar)
        except tk.TclError:
            return
        travel = h + 26
        start = bar["offset"] if bar["shown"] else travel
        end = 0 if show else travel
        bar["shown"] = True

        def step(p):
            if bar["gen"] != gen:
                return
            bar["offset"] = round(start + (end - start) * p)
            self._place_selection_bar()

        def done():
            if bar["gen"] != gen:
                return
            if not show:
                bar["shown"] = False
                bar["offset"] = travel
                try:
                    bar["holder"].place_forget()
                except tk.TclError:
                    pass

        quick = not animate or self._motion_reduced()
        _animate(self.root, 0 if quick else (220 if show else 150), step, done=done, ease=EASE_OUT)

    # what the bar does

    def _selection_collect_menu(self, anchor):
        names = self._collection_names()
        in_collection = self._view in names
        shared = set(self._shared_by_name())
        targets = [n for n in names if n != self._view]
        n = len(self._sel)
        items = []
        if targets:
            items.append(("header", f"Move {n} to" if in_collection else f"Add {n} to"))
            for name in targets:
                icon = "shared" if name in shared else "folder"
                items.append(("item", name, icon, lambda t=name, v=self._view: (
                    self._bulk_collect(t, move_from=v if in_collection else None),
                    self._pulse_card(t))))
        else:
            items.append(("item", "New collection…", "newfolder", self._new_collection))
        if in_collection:
            items += [("sep",), ("item", f"Remove {n} from {self._view}", "remove",
                                 lambda v=self._view: self._bulk_remove(v), {"danger": True})]
        self._popup(anchor.winfo_rootx(), anchor.winfo_rooty(), items)

    def _bulk_collect(self, target, move_from=None, items=None):
        items = items if items is not None else self._selection_items()
        ids = [i for i in ((m or {}).get("item_id") for _p, m in items) if i]
        if not ids:
            self.status_var.set("Only photos from the shared drive can be collected.")
            return
        by_id = self._collection_ids()
        target_ids = by_id.setdefault(target, [])
        added = 0
        for item_id in ids:
            if item_id not in target_ids:
                target_ids.append(item_id)
                added += 1
        moving = set(ids)
        if move_from and move_from in by_id:
            by_id[move_from] = [i for i in by_id[move_from] if i not in moving]
        if not self._write_collection_ids(by_id):
            return
        self._collections = self._load_collections()
        self._rebuild_collection_cards()
        already = len(ids) - added
        note = f" ({already} already there)" if already else ""
        if move_from and self._view == move_from:
            remaining = [r for r in self._last_results
                         if (r[2] or {}).get("item_id") not in moving]
            self._render_results(remaining, keep_scroll=True)
            self.status_var.set(f"Moved {len(ids)} photo(s) to {target}{note}.")
        else:
            self._clear_selection()
            self.status_var.set(f"Added {added} photo(s) to {target}{note}.")
            if self._view == target:
                self._show_collection(target)

    def _bulk_remove(self, name):
        items = self._selection_items()
        gone = {i for i in ((m or {}).get("item_id") for _p, m in items) if i}
        by_id = self._collection_ids()
        if not gone or name not in by_id:
            return
        by_id[name] = [i for i in by_id[name] if i not in gone]
        if not self._write_collection_ids(by_id):
            return
        self._collections = self._load_collections()
        self._rebuild_collection_cards()
        self.status_var.set(f"Removed {len(gone)} photo(s) from {name}.")
        if self._view == name:
            self._show_collection(name, keep_scroll=True)

    def _fetch_originals(self, items, folder, progress=None):
        """Full-resolution originals of `items`, copied into `folder` under
        their own names. Returns the copies' paths."""
        os.makedirs(folder, exist_ok=True)
        taken = set()
        out = []
        for n, (path, meta) in enumerate(items, start=1):
            if progress:
                progress(n)
            item_id = (meta or {}).get("item_id")
            name = (meta or {}).get("name") or os.path.basename(path)
            source = self._ensure_full_image(item_id, name) if item_id else path
            stem, ext = os.path.splitext(name)
            candidate, k = name, 2
            while candidate.lower() in taken or os.path.exists(os.path.join(folder, candidate)):
                candidate = f"{stem} ({k}){ext}"
                k += 1
            taken.add(candidate.lower())
            target = os.path.join(folder, candidate)
            shutil.copy2(source, target)
            out.append(target)
        return out

    def _copy_selection(self):
        items = self._selection_items()
        if not items:
            return
        if len(items) == 1:
            path, meta = items[0]
            self._copy_to_clipboard(path, meta)
            return
        total = len(items)
        self.status_var.set(f"Getting {total} photos ready to paste…")

        def work():
            import tempfile

            folder = os.path.join(tempfile.gettempdir(), "Topshot copies")
            try:
                for old in os.listdir(folder):
                    try:
                        os.remove(os.path.join(folder, old))
                    except OSError:
                        pass
            except OSError:
                pass
            try:
                files = self._fetch_originals(
                    items, folder,
                    lambda n: self.event_queue.put(("shared_status", f"Fetching photo {n} of {total}…")))
                self._put_files_on_clipboard(files)
                self.event_queue.put(("shared_status",
                                      f"Copied {len(files)} photos at full resolution - paste them "
                                      "into PowerPoint, Teams or an email."))
            except Exception as exc:
                _log_reuse(f"MULTI COPY FAILED: {exc!r}")
                self.event_queue.put(("shared_status", "Couldn't copy those photos - try again in a moment."))

        threading.Thread(target=work, daemon=True).start()

    def _put_files_on_clipboard(self, paths):
        """Puts files (not pixels) on the clipboard - how several photos
        travel at once, the same as copying them in Explorer or Finder."""
        if IS_MAC:
            files = ", ".join('POSIX file "{}"'.format(p.replace("\\", "\\\\").replace('"', '\\"'))
                              for p in paths)
            subprocess.run(["osascript", "-e", f"set the clipboard to {{{files}}}"],
                           check=True, timeout=30, capture_output=True)
            return
        import ctypes
        import struct

        names = ("\0".join(os.path.abspath(p) for p in paths) + "\0\0").encode("utf-16-le")
        payload = struct.pack("<IiiII", 20, 0, 0, 0, 1) + names  # DROPFILES, wide names
        kernel32, user32 = ctypes.windll.kernel32, ctypes.windll.user32
        kernel32.GlobalAlloc.restype = ctypes.c_void_p
        kernel32.GlobalAlloc.argtypes = [ctypes.c_uint, ctypes.c_size_t]
        kernel32.GlobalLock.restype = ctypes.c_void_p
        kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
        kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
        user32.SetClipboardData.restype = ctypes.c_void_p
        user32.SetClipboardData.argtypes = [ctypes.c_uint, ctypes.c_void_p]
        user32.RegisterClipboardFormatW.restype = ctypes.c_uint
        user32.RegisterClipboardFormatW.argtypes = [ctypes.c_wchar_p]

        def block(data):
            handle = kernel32.GlobalAlloc(0x0042, len(data))  # GMEM_MOVEABLE | GMEM_ZEROINIT
            pointer = kernel32.GlobalLock(handle)
            ctypes.memmove(pointer, data, len(data))
            kernel32.GlobalUnlock(handle)
            return handle

        files, effect = block(payload), block(struct.pack("<I", 1))  # DROPEFFECT_COPY
        for _attempt in range(10):
            if user32.OpenClipboard(0):
                break
            time.sleep(0.05)
        else:
            raise OSError("clipboard busy")
        try:
            user32.EmptyClipboard()
            user32.SetClipboardData(15, files)  # CF_HDROP
            user32.SetClipboardData(user32.RegisterClipboardFormatW("Preferred DropEffect"), effect)
        finally:
            user32.CloseClipboard()

    def _save_selection(self):
        items = self._selection_items()
        if not items:
            return
        folder = filedialog.askdirectory(title=f"Save {len(items)} photo(s) to…", parent=self.root)
        if not folder:
            return
        total = len(items)
        self.status_var.set(f"Saving {total} photo(s)…")

        def work():
            try:
                files = self._fetch_originals(
                    items, folder,
                    lambda n: self.event_queue.put(("shared_status", f"Saving photo {n} of {total}…")))
                self.event_queue.put(("shared_status", f"Saved {len(files)} photo(s) to {folder}."))
                try:
                    _open_external(folder)
                except Exception:
                    pass
            except Exception as exc:
                _log_reuse(f"SAVE SELECTION FAILED: {exc!r}")
                self.event_queue.put(("shared_status", "Couldn't save all of them - check the folder and try again."))

        threading.Thread(target=work, daemon=True).start()

    def _favourite_selection(self):
        items = self._selection_items()
        if not items:
            return
        if self._view == "favourites":
            for path, meta in items:
                self._unfavourite(self._fav_key(path, meta), path, meta)
            self._show_favourites(keep_scroll=True)
            self.status_var.set(f"Removed {len(items)} from favourites.")
            return
        own = self._own_favourites()
        picks = [(self._fav_key(p, m), p, m) for p, m in items]
        all_on = all(key in own for key, _p, _m in picks)
        for key, path, meta in picks:
            if (key in own) == all_on:
                self._toggle_favourite(key, path, meta)
        if self._view == "favourites":
            self._show_favourites()
        else:
            self._render_results(self._last_results, keep_scroll=True)
        self.status_var.set(
            f"Removed {len(picks)} from favourites." if all_on
            else f"Starred {len(picks)} photo(s) - find them under Favourites.")

    # ---------- after start: the tour for someone new, What's new after an update ----------

    def _whats_new_path(self):
        return os.path.join(DATA_DIR, "whats_new.json")

    def _mark_whats_new_seen(self):
        try:
            os.makedirs(DATA_DIR, exist_ok=True)
            with open(self._whats_new_path(), "w", encoding="utf-8") as f:
                json.dump({"seen": APP_VERSION}, f)
        except OSError:
            pass

    def _first_screen_extras(self):
        if not self._tour_seen():
            # Someone new gets the tour, not a list of changes.
            self._mark_whats_new_seen()
            self._start_tour_when_ready()
            return
        seen = str(self._read_json(self._whats_new_path()).get("seen") or "0")
        if seen != APP_VERSION and version_tuple(APP_VERSION) > version_tuple(seen):
            self._mark_whats_new_seen()
            self._show_whats_new(fresh=True)

    def _start_tour_when_ready(self, waited=0):
        """Starts the first-run tour once the window is really on screen -
        not minimised, not still hidden behind the intro - however long that
        takes."""
        try:
            ready = bool(self.root.winfo_viewable()) and self.root.state() != "iconic"
        except tk.TclError:
            return
        if ready and getattr(self, "_whats_new_win", None) is None:
            self._start_tour()
        else:
            self.root.after(700, lambda: self._start_tour_when_ready(waited + 700))

    # ---------- a card over the dimmed window ----------

    def _card_shell(self, width=440):
        """The What's new look for small panels: a card that rises in over
        the dimmed window (ease-out, 240 ms) and leaves faster (150 ms).
        Returns dict(win, body, show, close, refit)."""
        root = self.root
        reduced = self._motion_reduced()
        root.update_idletasks()
        rx, ry = root.winfo_rootx(), root.winfo_rooty()
        rw, rh = root.winfo_width(), root.winfo_height()
        scrim = tk.Toplevel(root)
        scrim.overrideredirect(True)
        scrim.configure(bg="#000000")
        try:
            scrim.attributes("-alpha", 0.0)
            scrim.transient(root)
        except tk.TclError:
            pass
        scrim.geometry(f"{rw}x{rh}+{rx}+{ry}")
        win = ctk.CTkToplevel(root)
        win.withdraw()
        win.overrideredirect(True)
        win.configure(fg_color=SURFACE)
        try:
            win.attributes("-alpha", 0.0)
            win.transient(root)
        except tk.TclError:
            pass
        body = ctk.CTkFrame(win, fg_color="transparent", width=width)
        body.pack(fill=tk.BOTH, expand=True, padx=26, pady=(22, 20))
        spot = {"x": 0, "y": 0, "closing": False}

        def place(rise=0):
            win.update_idletasks()
            h = win.winfo_reqheight()
            spot["x"] = rx + max(0, (rw - width) // 2)
            spot["y"] = ry + max(0, (rh - h) // 2 - 12)
            win.geometry(f"{width}x{h}+{spot['x']}+{spot['y'] + rise}")

        def close(then=None):
            if spot["closing"]:
                return
            spot["closing"] = True

            def step(p):
                try:
                    win.attributes("-alpha", 1.0 - p)
                    scrim.attributes("-alpha", 0.4 * (1.0 - p))
                    win.geometry(f"+{spot['x']}+{spot['y'] + round(6 * p)}")
                except tk.TclError:
                    pass

            def gone():
                for w in (win, scrim):
                    try:
                        w.destroy()
                    except tk.TclError:
                        pass
                if then:
                    then()

            _animate(root, 0 if reduced else 150, step, done=gone, ease=EASE_OUT)

        def show():
            rise = 0 if reduced else 12
            place(rise)
            win.deiconify()
            win.update_idletasks()
            _round_corners(win, border=MENU_BORDER)
            try:
                scrim.lift()
            except tk.TclError:
                pass
            win.lift()
            win.focus_force()
            win.bind("<Escape>", lambda _e: close())
            scrim.bind("<Button-1>", lambda _e: close())

            def enter(p):
                try:
                    win.attributes("-alpha", p)
                    scrim.attributes("-alpha", 0.4 * p)
                    win.geometry(f"+{spot['x']}+{spot['y'] + round(rise * (1 - p))}")
                except tk.TclError:
                    pass

            _animate(root, 120 if reduced else 240, enter, ease=EASE_OUT)

        def refit():
            if not spot["closing"]:
                try:
                    place()
                except tk.TclError:
                    pass

        return {"win": win, "body": body, "show": show, "close": close, "refit": refit}

    # ---------- photo details ----------

    def _show_details(self, path, meta):
        """Right-click > Details: what the file is and where it lives. Shown
        on request only, not under every photo."""
        meta = meta or {}
        card = self._card_shell(440)
        body = card["body"]
        name = meta.get("name") or os.path.basename(path)
        head = ctk.CTkFrame(body, fg_color="transparent")
        head.pack(fill=tk.X)
        try:
            thumb = self._thumb_image(path).copy()
            thumb.thumbnail((96, 96))
            pic = ctk.CTkImage(thumb, thumb, size=thumb.size)
            ctk.CTkLabel(head, text="", image=pic).pack(side=tk.LEFT, padx=(0, 16), anchor="n")
        except Exception:
            pass
        words = ctk.CTkFrame(head, fg_color="transparent")
        words.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ctk.CTkLabel(words, text="DETAILS", font=_font(11, "semibold"), text_color=SPLASH_ACCENT,
                     anchor="w").pack(fill=tk.X)
        ctk.CTkLabel(words, text=name, font=_font(16, "semibold"), text_color=TEXT_COLOR,
                     anchor="w", justify="left", wraplength=280).pack(fill=tk.X)
        ctk.CTkLabel(words, text=meta.get("folder") or "This computer", font=_font(12),
                     text_color=MUTED_TEXT_COLOR, anchor="w", justify="left",
                     wraplength=280).pack(fill=tk.X)
        ctk.CTkFrame(body, height=1, fg_color=BORDER).pack(fill=tk.X, pady=(16, 10))
        grid = ctk.CTkFrame(body, fg_color="transparent")
        grid.pack(fill=tk.X)
        grid.grid_columnconfigure(1, weight=1)
        fields = ("Dimensions", "File size", "Taken", "Camera", "Added", "Added by", "Folder")
        values = {}
        for row, label in enumerate(fields):
            ctk.CTkLabel(grid, text=label, font=_font(12), text_color=FAINT_TEXT, anchor="w",
                         width=90).grid(row=row, column=0, sticky="nw", pady=2)
            value = ctk.CTkLabel(grid, text="…", font=_font(12), text_color=TEXT_SOFT,
                                 anchor="w", justify="left", wraplength=280)
            value.grid(row=row, column=1, sticky="w", pady=2)
            values[label] = value
        buttons = ctk.CTkFrame(body, fg_color="transparent")
        buttons.pack(fill=tk.X, pady=(18, 0))
        link = {"url": ""}
        open_sp = _make_secondary_button(buttons, "Open in SharePoint",
                                         lambda: link["url"] and _open_external(link["url"]),
                                         width=160)
        _make_primary_button(buttons, "Close", lambda: card["close"](), width=90).pack(side=tk.RIGHT)
        card["show"]()

        result = {}

        def work():
            info = {}
            try:
                item_id = meta.get("item_id")
                if item_id:
                    client = self._sp_client(False)
                    if client is not None:
                        info = client.get_item_by_path(item_id, by_id=True) or {}
                result["info"] = info
            except Exception as exc:
                _log_reuse(f"DETAILS FAILED: {exc!r}")
                result["info"] = {}

        threading.Thread(target=work, daemon=True).start()

        def when(text):
            try:
                stamp = datetime.datetime.fromisoformat(str(text).replace("Z", "+00:00"))
                return stamp.astimezone().strftime("%d %b %Y, %H:%M")
            except Exception:
                return ""

        def fill():
            if "info" not in result:
                card["win"].after(120, fill)
                return
            info = result["info"]
            image, photo = info.get("image") or {}, info.get("photo") or {}
            size = info.get("size")
            width, height = image.get("width"), image.get("height")
            if not (width and height):
                try:
                    with Image.open(path) as im:  # this computer's copy, if that's all there is
                        if not meta.get("item_id"):
                            width, height = im.size
                except Exception:
                    pass
            if size is None and not meta.get("item_id"):
                try:
                    size = os.path.getsize(path)
                except OSError:
                    size = None
            folder_path = ((info.get("parentReference") or {}).get("path") or "").split("root:", 1)[-1]
            shown = {
                "Dimensions": f"{width} × {height} px" if width and height else "-",
                "File size": f"{size / 1048576:.1f} MB" if size else "-",
                "Taken": when(photo.get("takenDateTime")) or "-",
                "Camera": " ".join(x for x in (photo.get("cameraMake"), photo.get("cameraModel")) if x) or "-",
                "Added": when(info.get("createdDateTime")) or "-",
                "Added by": ((info.get("createdBy") or {}).get("user") or {}).get("displayName") or "-",
                "Folder": folder_path.strip("/") or meta.get("folder") or "-",
            }
            for label, widget in values.items():
                try:
                    widget.configure(text=shown[label])
                except tk.TclError:
                    return
            link["url"] = info.get("webUrl") or ""
            if link["url"]:
                open_sp.pack(side=tk.LEFT)
            card["refit"]()

        card["win"].after(120, fill)

    def _popup(self, x_root, y_root, items):
        """Opens the app's own menu at a screen point, closing any open one."""
        if getattr(self, "_menu", None) is not None:
            self._menu.close(animate=False)
        self._menu = PopMenu(self, x_root, y_root, items)

    def _show_result_menu(self, event, path, meta):
        """Right-click menu on a result: the keyboard-and-mouse route into a
        collection, for anyone who would rather not drag."""
        names = self._collection_names()
        in_collection = self._view in names
        shared = set(self._shared_by_name())
        if meta and meta.get("item_id"):
            open_original = lambda: self._open_sp_result(meta["item_id"], meta["name"])
        else:
            open_original = lambda: open_in_system_viewer(path)
        items = [
            ("item", "Preview", "preview", lambda: self._show_preview(path, meta)),
            ("item", "Copy image", "copy", lambda: self._copy_to_clipboard(path, meta)),
            ("item", "Open original", "external", open_original),
            ("item", "Details…", "info", lambda: self._show_details(path, meta)),
        ]
        targets = [n for n in names if n != self._view]
        if targets:
            items += [("sep",), ("header", "Move to" if in_collection else "Add to collection")]
            for name in targets:
                icon = "shared" if name in shared else "folder"
                if in_collection:
                    action = lambda n=name, v=self._view: (
                        self._move_to_collection(v, n, path, meta), self._pulse_card(n))
                else:
                    action = lambda n=name: (self._add_to_collection(n, path, meta), self._pulse_card(n))
                items.append(("item", name, icon, action))
        if in_collection:
            items += [("sep",), ("item", f"Remove from {self._view}", "remove",
                                 lambda: self._remove_from_collection(self._view, meta), {"danger": True})]
        self._popup(event.x_root, event.y_root, items)

    # ---------- drag a photo onto a collection ----------

    # ---------- drag and drop into collections ----------
    #
    # The photo lifts off the grid into a small card that follows the
    # pointer, and its place in the grid dims. The collection under the
    # pointer lights up and the card shrinks a little more to say "it goes
    # in here". Released there, it flies into the collection, which glows
    # to confirm; when the photo came from another collection it is moved,
    # and its place in the grid fades out. Released anywhere else, it
    # glides back to where it came from.

    def _drag_start(self, event, path, meta):
        # Only records where the press happened. The drag does not begin
        # until the pointer actually moves, so a double-click still opens
        # the photo instead of being swallowed here.
        try:
            w, h = event.widget.winfo_width(), event.widget.winfo_height()
        except tk.TclError:
            w, h = 1, 1
        self._drag = {
            "path": path, "meta": meta, "label": event.widget,
            "x": event.x_root, "y": event.y_root,
            "px": event.x_root, "py": event.y_root,
            "fx": min(1.0, max(0.0, event.x / max(1, w))),
            "fy": min(1.0, max(0.0, event.y / max(1, h))),
            "active": False, "ghost": None, "hover": None, "tick": None,
            "group": self._drag_group(event.widget),
        }

    def _motion_reduced(self):
        if getattr(self, "_reduced_motion", None) is None:
            self._reduced_motion = _prefers_reduced_motion()
        return self._reduced_motion

    def _card_under_pointer(self, x_root, y_root):
        # Cards scrolled out of view in the sidebar still have screen
        # positions - over the System section - so only the visible part of
        # the list counts.
        try:
            area = self.sidebar._parent_canvas
            top = area.winfo_rooty()
            if not (top <= y_root <= top + area.winfo_height()):
                return None
        except Exception:
            pass
        for name, card in (self.collection_cards or {}).items():
            try:
                cx, cy = card.winfo_rootx(), card.winfo_rooty()
                if (cx <= x_root <= cx + card.winfo_width()
                        and cy <= y_root <= cy + card.winfo_height()):
                    return name
            except Exception:
                continue
        return None

    def _drop_target(self, x_root, y_root):
        """The collection a drop would land in - never the one being viewed,
        since dropping a photo where it already is does nothing."""
        name = self._card_under_pointer(x_root, y_root)
        return None if name == self._view else name

    def _card_rest_colour(self, name):
        return NAV_SELECTED if self._view == name else NAV_BG

    def _fade_card(self, name, to_colour, ms, start=None, ease=None):
        """Eases a collection card's colour. A newer fade on the same card
        takes over from wherever this one has got to."""
        card = (self.collection_cards or {}).get(name)
        if card is None:
            return
        if start is None:
            try:
                start = card.cget("fg_color")
            except Exception:
                start = None
            if not (isinstance(start, str) and start.startswith("#") and len(start) == 7):
                start = self._card_rest_colour(name)
        gens = self.__dict__.setdefault("_card_fade_gen", {})
        gen = gens.get(name, 0) + 1
        gens[name] = gen

        def step(p):
            if gens.get(name) != gen:
                return
            try:
                card.configure(fg_color=_mix(start, to_colour, p))
            except Exception:
                pass

        _animate(self.root, ms, step, ease=ease or EASE_OUT)

    def _pulse_card(self, name):
        """A collection glows and settles back, confirming what landed in it."""
        self._fade_card(name, self._card_rest_colour(name), CARD_PULSE_MS, start=PULSE_COLOUR)

    def _set_drag_hover(self, name):
        drag = self._drag
        if drag is None or drag.get("hover") == name:
            return
        previous = drag.get("hover")
        if previous:
            self._fade_card(previous, self._card_rest_colour(previous), DRAG_HOVER_MS)
        if name:
            self._fade_card(name, ACCENT_ACTIVE, DRAG_HOVER_MS)
        drag["hover"] = name

    def _drag_motion(self, event):
        drag = self._drag
        if drag is None:
            return
        drag["px"], drag["py"] = event.x_root, event.y_root
        if not drag["active"]:
            moved = abs(event.x_root - drag["x"]) + abs(event.y_root - drag["y"])
            if moved < DRAG_THRESHOLD_PX:
                return
            drag["active"] = True
            self._dim_source(drag)
            drag["ghost"] = self._make_ghost(drag)
            if drag["ghost"] is not None:
                _fine_timer(True)
                drag["t"] = time.perf_counter()
                self._drag_tick(drag)
        self._set_drag_hover(self._drop_target(event.x_root, event.y_root))

    def _dim_source(self, drag):
        """The photo's place in the grid dims while it is being carried, and
        can fade out the rest of the way if it is moved away."""
        label = drag["label"]
        try:
            path, size, key = label.tile
            tile = self._tile_pil(path, size, key in self._render_favs)
            bg = Image.new("RGBA", size, _hex_rgb(BG_COLOR) + (255,))
            drag["dim"] = [
                ImageTk.PhotoImage(Image.blend(tile, bg, a)) for a in (0.62, 0.74, 0.85, 0.94, 1.0)
            ]
            label.dragging = True
            label.configure(image=drag["dim"][0])
        except Exception:
            drag["dim"] = None

    def _restore_source(self, drag):
        label = drag["label"]
        try:
            label.dragging = False
        except Exception:
            pass
        self._paint_tile(label)

    def _make_ghost(self, drag):
        """The carried photo: a borderless window that starts exactly over
        the tile, so lifting it reads as one object leaving the grid."""
        label = drag["label"]
        try:
            tile_w, tile_h = label.winfo_width(), label.winfo_height()
            x0, y0 = label.winfo_rootx(), label.winfo_rooty()
            win = tk.Toplevel(self.root)
            win.overrideredirect(True)
            win.configure(bg=BG_COLOR)
            try:
                win.attributes("-topmost", True)
            except Exception:
                pass
            inner = tk.Label(win, bd=0, highlightthickness=0, bg=BG_COLOR)
            inner.pack()
            g = {
                "win": win, "label": inner,
                "base": self._with_count_badge(self._thumb_image(drag["path"]), len(drag["group"]))
                if drag.get("group") else self._thumb_image(drag["path"]),
                "w": max(8, tile_w), "h": max(8, tile_h), "size": None,
                "cx": x0 + tile_w / 2, "cy": y0 + tile_h / 2, "scale": 1.0, "alpha": 1.0,
            }
            self._place_ghost(g)
            win.update_idletasks()
            _round_corners(win, border=SPLASH_ACCENT)
            return g
        except Exception:
            return None

    def _place_ghost(self, g):
        w = max(8, round(g["w"] * g["scale"]))
        h = max(8, round(g["h"] * g["scale"]))
        if g["size"] != (w, h):
            photo = ImageTk.PhotoImage(ImageOps.fit(g["base"], (w, h), Image.BILINEAR))
            g["label"].configure(image=photo)
            g["label"].image = photo
            g["size"] = (w, h)
        g["win"].geometry(f"{w}x{h}+{round(g['cx'] - w / 2)}+{round(g['cy'] - h / 2)}")
        try:
            g["win"].attributes("-alpha", max(0.0, min(1.0, g["alpha"])))
        except Exception:
            pass

    def _drag_tick(self, drag):
        """One frame of the carried photo easing after the pointer."""
        drag["tick"] = None
        g = drag.get("ghost")
        if g is None or drag.get("released"):
            return
        now = time.perf_counter()
        dt = max(1.0, (now - drag["t"]) * 1000.0)
        drag["t"] = now
        scale = DRAG_SCALE_OVER if drag.get("hover") else DRAG_SCALE
        # The spot that was grabbed stays under the pointer as the card shrinks.
        tx = drag["px"] + (0.5 - drag["fx"]) * g["w"] * scale
        ty = drag["py"] + (0.5 - drag["fy"]) * g["h"] * scale
        if self._motion_reduced():
            kp = ks = 1.0
        else:
            kp = 1.0 - math.exp(-dt / DRAG_FOLLOW_TAU_MS)
            ks = 1.0 - math.exp(-dt / DRAG_SETTLE_TAU_MS)
        g["cx"] += (tx - g["cx"]) * kp
        g["cy"] += (ty - g["cy"]) * kp
        g["scale"] += (scale - g["scale"]) * ks
        g["alpha"] += (DRAG_ALPHA - g["alpha"]) * ks
        try:
            self._place_ghost(g)
        except tk.TclError:
            return
        drag["tick"] = self.root.after(FRAME_MS, lambda: self._drag_tick(drag))

    def _drag_release(self, event):
        drag = self._drag
        self._drag = None
        if drag is None:
            return
        drag["released"] = True
        if drag.get("tick") is not None:
            try:
                self.root.after_cancel(drag["tick"])
            except Exception:
                pass
        if not drag["active"]:
            return  # a plain click, not a drag
        target = self._drop_target(event.x_root, event.y_root)
        hover = drag.get("hover")
        if hover and hover != target:
            self._fade_card(hover, self._card_rest_colour(hover), DRAG_HOVER_MS)
        if target:
            self._fly_into(drag, target)
        else:
            self._fly_back(drag)

    def _end_ghost(self, drag):
        g = drag.get("ghost")
        drag["ghost"] = None
        if g is not None:
            try:
                g["win"].destroy()
            except Exception:
                pass
            _fine_timer(False)

    def _fly_into(self, drag, target):
        g = drag.get("ghost")
        moving = self._view in self._collection_names() and self._view != target
        source = self._view

        def land():
            self._end_ghost(drag)
            if drag.get("group"):
                self._restore_source(drag)
                self._bulk_collect(target, move_from=source if moving else None,
                                   items=drag["group"])
            elif moving:
                self._move_to_collection(source, target, drag["path"], drag["meta"])
            else:
                self._restore_source(drag)
                self._add_to_collection(target, drag["path"], drag["meta"])
            self._pulse_card(target)

        card = (self.collection_cards or {}).get(target)
        if g is None or card is None:
            land()
            return
        try:
            ex = card.winfo_rootx() + card.winfo_width() / 2
            ey = card.winfo_rooty() + card.winfo_height() / 2
        except tk.TclError:
            land()
            return
        x0, y0, s0, a0 = g["cx"], g["cy"], g["scale"], g["alpha"]
        dim = drag.get("dim") if moving else None

        def step(p):
            g["cx"] = x0 + (ex - x0) * p
            g["cy"] = y0 + (ey - y0) * p
            g["scale"] = s0 + (0.14 - s0) * p
            g["alpha"] = a0 * (1.0 - p ** 3)  # stays visible most of the way in
            self._place_ghost(g)
            if dim:  # the photo's old place in the grid fades out alongside
                try:
                    drag["label"].configure(image=dim[min(len(dim) - 1, int(p * len(dim)))])
                except Exception:
                    pass

        ms = 0 if self._motion_reduced() else DROP_FLY_MS
        _animate(self.root, ms, step, land, ease=EASE_IN_OUT)

    def _fly_back(self, drag):
        g = drag.get("ghost")

        def home():
            self._end_ghost(drag)
            self._restore_source(drag)

        if g is None:
            home()
            return
        label = drag["label"]
        try:
            hx = label.winfo_rootx() + label.winfo_width() / 2
            hy = label.winfo_rooty() + label.winfo_height() / 2
        except tk.TclError:
            home()
            return
        x0, y0, s0, a0 = g["cx"], g["cy"], g["scale"], g["alpha"]

        def step(p):
            g["cx"] = x0 + (hx - x0) * p
            g["cy"] = y0 + (hy - y0) * p
            g["scale"] = s0 + (1.0 - s0) * p
            g["alpha"] = a0 + (1.0 - a0) * p
            self._place_ghost(g)

        ms = 0 if self._motion_reduced() else DROP_BACK_MS
        _animate(self.root, ms, step, home, ease=EASE_OUT)

    def _move_to_collection(self, source, target, path, meta):
        """Moves a photo out of one collection and into another."""
        item_id = (meta or {}).get("item_id")
        if not item_id:
            self.status_var.set("Only photos from the shared drive can be collected.")
            return
        by_id = self._collection_ids()
        source_ids = by_id.get(source) or []
        target_ids = by_id.setdefault(target, [])
        already = item_id in target_ids
        if item_id in source_ids:
            source_ids.remove(item_id)
        if not already:
            target_ids.append(item_id)
        if not self._write_collection_ids(by_id):
            return
        self._collections = self._load_collections()
        self._rebuild_collection_cards()
        if self._view == source:
            remaining = [r for r in self._last_results
                         if ((r[2] or {}).get("item_id") if r[2] else None) != item_id]
            self._render_results(remaining, keep_scroll=True)
        note = f" (it was already in {target})" if already else ""
        self.status_var.set(
            f"Moved to {target}{note}. {len(source_ids)} photo(s) left in {source}."
        )

    def _scroll_results_to_top(self):
        """A new set of results starts at the top. Without this, switching
        from a long view (scrolled down) to a short one kept the old scroll
        position - below the end of the new content, so it looked empty."""
        try:
            self.results_frame._parent_canvas.yview_moveto(0)
        except Exception:
            pass

    # ---------- before the library is loaded ----------

    # Widths of the placeholder tiles, as aspect ratios: the mix of landscape
    # and portrait a real screen of photos has, so the grid looks like itself.
    PLACEHOLDER_ASPECTS = (1.5, 0.75, 1.33, 1.0, 1.6, 0.8, 1.25, 1.5, 0.7, 1.4, 1.0, 1.78,
                           0.75, 1.33, 1.5, 0.8, 1.2, 1.6, 1.0, 1.4, 0.75, 1.5, 1.33, 0.9)

    def _render_placeholder_grid(self):
        """Calm empty tiles in the shape of the real grid, shown until the
        library has been loaded once: the window never looks broken, and the
        tour has something to point at. Rows further down fade into the page."""
        for child in self.results_frame.winfo_children():
            child.destroy()
        self._placeholder_grid = True
        self._scroll_results_to_top()
        avail = max(320, self._desired_columns() - 2 * GRID_SIDE_PAD - 4)
        aspects = self.PLACEHOLDER_ASPECTS
        n = 0
        for depth in range(4):
            row_aspects, total = [], 0.0
            while not row_aspects or total * ROW_HEIGHT + TILE_GAP * (len(row_aspects) - 1) < avail:
                a = aspects[n % len(aspects)]
                n += 1
                row_aspects.append(a)
                total += a
            gaps = TILE_GAP * (len(row_aspects) - 1)
            height = max(40, round((avail - gaps) / total))
            widths = [max(24, round(a * height)) for a in row_aspects]
            widths[-1] = max(24, widths[-1] + (avail - gaps) - sum(widths))
            row = tk.Frame(self.results_frame, bg=BG_COLOR, bd=0, highlightthickness=0)
            row.pack(anchor="w", padx=(GRID_SIDE_PAD, 0), pady=(0, TILE_GAP))
            for i, w in enumerate(widths):
                try:
                    photo = ImageTk.PhotoImage(self._placeholder_tile((w, height), depth))
                except Exception:
                    continue
                label = tk.Label(row, image=photo, bd=0, highlightthickness=0, bg=BG_COLOR)
                label.image = photo
                label.pack(side=tk.LEFT, padx=(0 if i == 0 else TILE_GAP, 0))

    def _placeholder_tile(self, size, depth=0):
        fade = min(0.75, 0.2 * depth)
        top = Image.new("RGBA", size, _hex_rgb(_mix(SURFACE_HOVER, BG_COLOR, fade)) + (255,))
        bottom = Image.new("RGBA", size, _hex_rgb(_mix(SURFACE, BG_COLOR, fade)) + (255,))
        body = Image.composite(bottom, top, Image.linear_gradient("L").resize(size))
        tile = Image.new("RGBA", size, _hex_rgb(BG_COLOR) + (255,))
        tile.paste(body, (0, 0), self._tile_mask(size))
        ink = _mix(FAINT_TEXT, BG_COLOR, min(0.85, 0.45 + fade))
        icon_size = max(16, min(30, size[1] // 5))
        if _icon("photos", icon_size, ink) is not None:
            glyph = _icon_pil_cache.get(("photos", icon_size, ink))
            if glyph is not None:
                tile.alpha_composite(glyph, ((size[0] - icon_size) // 2, (size[1] - icon_size) // 2))
        return tile

    def _render_results(self, results, keep_scroll=False):
        self._placeholder_grid = False
        if results is not self._last_results:
            # A re-draw of the same set (the window was resized) is still
            # whatever it was - the All photos sample stays the sample.
            self._sample_view = False
        self._last_results = results
        self._clear_selection(animate=False)
        shown = self._apply_filters(results)
        self._shown_results = shown
        self._tile_by_pos = {}
        self._show_filter_count(len(results), len(shown))
        self._restore_scroll = None
        if keep_scroll:
            try:
                self._restore_scroll = self.results_frame._parent_canvas.yview()[0]
            except Exception:
                pass
        # A fill from the previous render may still be in flight.
        if self._render_after_id is not None:
            try:
                self.root.after_cancel(self._render_after_id)
            except Exception:
                pass
            self._render_after_id = None
        for child in self.results_frame.winfo_children():
            child.destroy()
        if self._restore_scroll is None:
            self._scroll_results_to_top()
        self.thumbnail_refs.clear()
        # Read once per render, not once per photo: it is a file on disk.
        # In Favourites every photo is one - all of them carry the gold star.
        self._render_favs = set(self._load_favourites() if self._view == "favourites"
                                else self._own_favourites())
        # Measured now rather than trusting the last resize event, so the
        # first layout after startup already has the right column count.
        self._current_columns = self._desired_columns()
        avail = max(320, self._current_columns - 2 * GRID_SIDE_PAD - 4)
        self._render_chunk(shown, 0, avail)

    def _render_chunk(self, results, start, avail):
        """Builds a few rows of results, then yields to the event loop
        before the next, so the first rows appear at once.

        Rows are justified: photos are added at ROW_HEIGHT until the row
        reaches the available width, then the row is scaled to end exactly
        at the right edge. Each photo is one plain Tk label (CustomTkinter
        widgets cost ~33ms each); the star and the rounded corners are
        drawn into the image itself.
        """
        self._render_after_id = None
        n = len(results)
        idx = start
        rows = 0
        while idx < n and rows < ROWS_PER_CHUNK:
            items = []
            total_aspect = 0.0
            full = False
            while idx < n:
                path, _score, meta = results[idx]
                idx += 1
                try:
                    img = self._thumb_image(path)
                except Exception:
                    continue
                aspect = min(3.0, max(0.4, img.width / max(1, img.height)))
                items.append((path, meta, aspect, idx - 1))
                total_aspect += aspect
                if total_aspect * ROW_HEIGHT + TILE_GAP * (len(items) - 1) >= avail:
                    full = True
                    break
            if not items:
                break
            gaps = TILE_GAP * (len(items) - 1)
            if full:
                height = max(40, round((avail - gaps) / total_aspect))
                widths = [max(24, round(a * height)) for _p, _m, a, _i in items]
                widths[-1] += (avail - gaps) - sum(widths)  # land exactly on the edge
                widths[-1] = max(24, widths[-1])
            else:
                height = ROW_HEIGHT
                widths = [max(24, round(a * height)) for _p, _m, a, _i in items]
            self._build_row(items, widths, height)
            rows += 1

        if getattr(self, "_restore_scroll", None) is not None:
            try:
                self.root.update_idletasks()
                self.results_frame._parent_canvas.yview_moveto(self._restore_scroll)
            except Exception:
                pass
            if idx >= n:
                self._restore_scroll = None

        if idx < n:
            self._render_after_id = self.root.after(
                1, lambda: self._render_chunk(results, idx, avail)
            )
        elif self._render_done_callback is not None:
            callback, self._render_done_callback = self._render_done_callback, None
            callback()

    def _build_row(self, items, widths, height):
        row = tk.Frame(self.results_frame, bg=BG_COLOR, bd=0, highlightthickness=0)
        row.pack(anchor="w", padx=(GRID_SIDE_PAD, 0), pady=(0, TILE_GAP))
        for i, ((path, meta, _a, pos), w) in enumerate(zip(items, widths)):
            key = self._fav_key(path, meta)
            size = (w, height)
            try:
                photo = self._tile_photo(path, size, key in self._render_favs)
            except Exception:
                continue
            label = tk.Label(
                row, image=photo, bd=0, highlightthickness=0, padx=0, pady=0,
                bg=BG_COLOR, cursor="hand2",
            )
            label.image = photo  # a PhotoImage only lives while something refers to it
            label.hover_image = None
            label.tile = (path, size, key)
            label.meta = meta
            label.pack(side=tk.LEFT, padx=(0 if i == 0 else TILE_GAP, 0))
            # Double-click previews in-app; the original is one click away
            # from inside the preview. A click on the star toggles it.
            label.bind("<Double-Button-1>", lambda e, p=path, m=meta: self._on_thumb_double(e, p, m))
            _bind_context_menu(label, lambda e, p=path, m=meta: self._show_result_menu(e, p, m))
            label.bind("<ButtonPress-1>", lambda e, p=path, m=meta: self._drag_start(e, p, m))
            label.bind("<B1-Motion>", self._drag_motion)
            label.bind(
                "<ButtonRelease-1>",
                lambda e, k=key, p=path, m=meta: self._on_thumb_release(e, k, p, m),
            )
            label.bind("<Enter>", lambda e: self._tile_hover(e.widget, True))
            label.bind("<Leave>", lambda e: self._tile_hover(e.widget, False))
            # Picking several: Ctrl-click (Cmd on a Mac) adds or drops one,
            # Shift-click takes everything from the last pick to this one.
            label.pos = pos
            label.selected = False
            label.sel_image = None
            self._tile_by_pos[pos] = label
            label.bind("<Shift-Button-1>", lambda e, p=pos: self._select_click(p, "range"))
            label.bind("<Command-Button-1>" if IS_MAC else "<Control-Button-1>",
                       lambda e, p=pos: self._select_click(p, "toggle"))

    def _fade_tile_out(self, widget, done):
        """A photo leaving the grid: it sinks into the page (ease-out, 180 ms)."""
        try:
            path, size, _key = widget.tile
            tile = self._tile_pil(path, size, False)
            page = Image.new("RGBA", size, _hex_rgb(BG_COLOR) + (255,))
            frames = [ImageTk.PhotoImage(Image.blend(tile, page, a)) for a in (0.3, 0.55, 0.8, 1.0)]
        except Exception:
            done()
            return
        widget.fade_frames = frames
        widget.dragging = True  # no hover star while it goes

        def step(p):
            try:
                widget.configure(image=frames[min(len(frames) - 1, int(p * len(frames)))])
            except tk.TclError:
                pass

        _animate(self.root, 0 if self._motion_reduced() else 180, step, done=done, ease=EASE_OUT)

    def _on_star_clicked(self, key, path, meta, widget):
        if self._view == "favourites":
            # Everything here is a favourite: the star takes it out, and the
            # photo fades away before the list closes the gap.
            self._unfavourite(key, path, meta)
            self._render_favs.discard(key)
            self.status_var.set("Removed from favourites.")
            self._fade_tile_out(widget, lambda: self._show_favourites(keep_scroll=True))
            return
        added = self._toggle_favourite(key, path, meta)
        (self._render_favs.add if added else self._render_favs.discard)(key)
        try:
            _p, size, _k = widget.tile
            widget.image = self._tile_photo(path, size, added)
            # The pointer is still on the photo, so an un-starred photo
            # keeps showing the outline star until the pointer leaves.
            widget.hover_image = None if added else self._tile_photo(path, size, False, hover=True)
            widget.configure(image=widget.image if added else widget.hover_image)
        except Exception:
            pass
        # Un-starring while viewing favourites should drop it from the list.
        if self._view == "favourites" and not added:
            self._show_favourites()

# ---------- startup: splash and reveal ----------

SPLASH_BG = "#090b0a"  # the window's own background, so the handover is seamless
SPLASH_TRACK = "#1d2621"
SPLASH_ACCENT = "#34d399"  # a soft white, matching the neutral palette
SPLASH_FAINT = "#6b7a72"
SPLASH_BORDER = "#24302a"
FRAME_MS = 8


def _cubic_bezier(x1, y1, x2, y2):
    """An easing function equivalent to CSS cubic-bezier(x1, y1, x2, y2)."""
    cx = 3 * x1
    bx = 3 * (x2 - x1) - cx
    ax = 1 - cx - bx
    cy = 3 * y1
    by = 3 * (y2 - y1) - cy
    ay = 1 - cy - by

    def x_at(t):
        return ((ax * t + bx) * t + cx) * t

    def y_at(t):
        return ((ay * t + by) * t + cy) * t

    def ease(x):
        if x <= 0:
            return 0.0
        if x >= 1:
            return 1.0
        t = x
        for _ in range(8):  # Newton's method usually lands in a few steps
            err = x_at(t) - x
            if abs(err) < 1e-6:
                return y_at(t)
            slope = (3 * ax * t + 2 * bx) * t + cx
            if abs(slope) < 1e-6:
                break
            t -= err / slope
            if not 0.0 <= t <= 1.0:
                break
        lo, hi, t = 0.0, 1.0, x  # bisection where Newton cannot cope
        for _ in range(40):
            if x_at(t) < x:
                lo = t
            else:
                hi = t
            t = (lo + hi) / 2
        return y_at(t)

    return ease


# A strong ease-out: most of the movement happens at once and then settles,
# so an entrance reads as immediate. The stock "ease-out" is too timid.
EASE_OUT = _cubic_bezier(0.23, 1, 0.32, 1)
# For things travelling across the screen (the progress bar): accelerates
# away and settles, instead of lurching off at full speed.
EASE_IN_OUT = _cubic_bezier(0.65, 0, 0.35, 1)


def _linear(x):
    return x


def _mix(start, end, amount):
    """The colour `amount` of the way from `start` to `end` (both #rrggbb).
    Tk can't fade a canvas item's opacity, but on a solid background
    fading its colour into the background looks the same."""
    a, b = _hex_rgb(start), _hex_rgb(end)
    return "#%02x%02x%02x" % tuple(round(x + (y - x) * amount) for x, y in zip(a, b))


def _fade_frames(image, levels):
    """`image` at `levels` opacities, from invisible to solid, so it can be
    faded by swapping frames rather than re-rendering on every tick."""
    alpha = image.getchannel("A")
    frames = []
    for i in range(levels):
        frame = image.copy()
        frame.putalpha(alpha.point(lambda v, s=i / (levels - 1): round(v * s)))
        frames.append(ImageTk.PhotoImage(frame))
    return frames


def _glow_frames(levels=12):
    """A soft light for the tip of the progress bar, at a range of
    strengths, so it can breathe without re-rendering each frame."""
    w, h, k = 44, 18, 4
    blob = Image.new("L", (w * k, h * k), 0)
    ImageDraw.Draw(blob).ellipse((w * k * 0.22, h * k * 0.30, w * k * 0.78, h * k * 0.70), fill=255)
    blob = blob.filter(ImageFilter.GaussianBlur(5 * k)).resize((w, h), Image.LANCZOS)
    tint = Image.new("RGB", (w, h), _hex_rgb(SPLASH_ACCENT))
    frames = []
    for i in range(levels):
        frame = tint.copy()
        frame.putalpha(blob.point(lambda v, s=0.9 * i / (levels - 1): round(v * s)))
        frames.append(ImageTk.PhotoImage(frame))
    return frames


def _animate(scheduler, duration_ms, step, done=None, ease=EASE_OUT):
    """Calls step(progress) from 0 to 1 over duration_ms, then done().

    Driven by the clock, not by counting frames: if the main thread is busy
    and a frame comes late, the animation jumps to where it should be rather
    than stretching out. done() always runs, even if a window vanished
    mid-way, so a sequence built on these can never stall.
    """
    start = time.perf_counter()

    def frame():
        try:
            elapsed = (time.perf_counter() - start) * 1000.0
            if duration_ms > 0 and elapsed < duration_ms:
                step(ease(elapsed / duration_ms))
                scheduler.after(FRAME_MS, frame)
                return
            step(1.0)
        except tk.TclError:
            pass
        if done is not None:
            done()

    frame()


def _prefers_reduced_motion():
    """Windows' "Animation effects" switch (Settings > Accessibility > Visual
    effects). When it is off, movement is dropped and fades are shortened -
    fades stay, because they help the eye follow what changed."""
    try:
        import ctypes

        enabled = ctypes.c_int(1)
        ok = ctypes.windll.user32.SystemParametersInfoW(
            0x1042, 0, ctypes.byref(enabled), 0  # SPI_GETCLIENTAREAANIMATION
        )
        return bool(ok) and not enabled.value
    except Exception:
        return False


def _start_freeze_watchdog(root):
    """Writes what every thread was doing to freeze.log whenever the window
    stops responding for a few seconds (and again if it stays stuck), plus a
    line when it recovers. A hang on someone else's computer can't be
    watched; this records where it happened."""
    import faulthandler

    path = os.path.join(DATA_DIR, "freeze.log")
    beat = [time.monotonic()]

    def pulse():
        beat[0] = time.monotonic()
        try:
            root.after(250, pulse)
        except tk.TclError:
            pass

    def note(text, dump=False):
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "a", encoding="utf-8") as f:
                f.write(f"--- {datetime.datetime.now().isoformat(timespec='seconds')} "
                        f"{APP_VERSION}: {text}\n")
                f.flush()
                if dump:
                    faulthandler.dump_traceback(file=f, all_threads=True)
                    f.write("\n")
        except Exception:
            pass

    def watch():
        dumps = 0
        while True:
            time.sleep(0.5)
            stalled = time.monotonic() - beat[0]
            if stalled >= 3.0 and dumps == 0 or stalled >= 12.0 and dumps == 1:
                dumps += 1
                note(f"window not responding for {stalled:.1f}s", dump=True)
            elif stalled < 1.0 and dumps:
                note(f"responding again (dumped {dumps}x)")
                dumps = 0

    pulse()
    threading.Thread(target=watch, name="freeze-watchdog", daemon=True).start()


def _fine_timer(on):
    """Windows runs Tk's timers at about 15ms by default - 64 frames a second
    at best, visibly steppy on a fast display. A 1ms resolution is asked for
    only while the intro plays, then handed back."""
    try:
        import ctypes

        winmm = ctypes.windll.winmm
        (winmm.timeBeginPeriod if on else winmm.timeEndPeriod)(1)
    except Exception:
        pass


def _round_corners(win, border=SPLASH_BORDER):
    """Asks Windows 11 to draw a borderless window with rounded corners and a
    hairline border, like its own flyouts. A no-op on Windows 10."""
    try:
        import ctypes

        hwnd = ctypes.windll.user32.GetParent(win.winfo_id())
        dwm = ctypes.windll.dwmapi
        corner = ctypes.c_int(2)  # DWMWCP_ROUND
        dwm.DwmSetWindowAttribute(hwnd, 33, ctypes.byref(corner), ctypes.sizeof(corner))
        r, g, b = (int(border[i:i + 2], 16) for i in (1, 3, 5))
        colour = ctypes.c_uint(r | (g << 8) | (b << 16))  # COLORREF is 0x00BBGGRR
        dwm.DwmSetWindowAttribute(hwnd, 34, ctypes.byref(colour), ctypes.sizeof(colour))
    except Exception:
        pass


# Started with this flag by "Restart to update": the intro then fills the
# screen and appears without its entrance, continuing the screen the old
# version is showing, so the update reads as one window reloading.
AFTER_UPDATE = "--after-update" in sys.argv
# py launcher.py --first-run-preview: the app as a brand-new computer sees it -
# no photos, the tour - without touching the library on this one.
FIRST_RUN_PREVIEW = "--first-run-preview" in sys.argv


def _set_window_icons(root):
    """Hands Windows the 256 px and 32 px icons directly. Tk only gives it the
    32 px one, which the taskbar then stretches - why the icon looked soft
    next to other apps."""
    try:
        import ctypes

        user32 = ctypes.windll.user32
        user32.LoadImageW.restype = ctypes.c_void_p
        user32.LoadImageW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_uint,
                                      ctypes.c_int, ctypes.c_int, ctypes.c_uint]
        user32.GetParent.restype = ctypes.c_void_p
        user32.SendMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_void_p,
                                        ctypes.c_void_p]
        path = _resource_path("assets", "app_icon.ico")
        hwnd = user32.GetParent(root.winfo_id())
        big = user32.LoadImageW(None, path, 1, 256, 256, 0x10)    # IMAGE_ICON, LR_LOADFROMFILE
        small = user32.LoadImageW(None, path, 1, 32, 32, 0x10)
        if big:
            user32.SendMessageW(hwnd, 0x80, 1, big)               # WM_SETICON, ICON_BIG
        if small:
            user32.SendMessageW(hwnd, 0x80, 0, small)             # ICON_SMALL
        root._icon_handles = (big, small)
    except Exception:
        pass


def _handover_path():
    return os.path.join(
        DATA_DIR, "app", "handover"
    )


def _work_area(win):
    """The primary screen minus the taskbar - where a maximised window sits."""
    try:
        import ctypes
        import ctypes.wintypes

        rect = ctypes.wintypes.RECT()
        if ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0):
            return rect.left, rect.top, rect.right, rect.bottom
    except Exception:
        pass
    return 0, 0, win.winfo_screenwidth(), win.winfo_screenheight()


class PopMenu:
    """The app's own right-click menu, in its colours and type.

    Drawn inside the main window (not as a separate pop-up window), so it
    never lands on the wrong monitor or behind anything. It opens from the
    point that was clicked - down and to the right, or up/left near an edge -
    unfolding in one short ease-out, and folds back the same way, faster.
    A click anywhere else, or Esc, closes it.
    """

    WIDTH = 248
    OPEN_MS, CLOSE_MS = 170, 110

    def __init__(self, app, x_root, y_root, items):
        self.app, self.root = app, app.root
        self.closed = False
        # CustomTkinter panels can't be sized by place(), so a plain Tk frame
        # is placed and sized (it clips the panel while the menu unfolds) and
        # the styled panel sits inside it at its natural size.
        self.holder = tk.Frame(self.root, bg=BG_COLOR, bd=0, highlightthickness=0)
        f = self.frame = ctk.CTkFrame(self.holder, fg_color=SURFACE, corner_radius=10,
                                      border_width=1, border_color=MENU_BORDER,
                                      bg_color=BG_COLOR)
        inner = ctk.CTkFrame(f, fg_color="transparent")
        inner.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        for item in items:
            kind = item[0]
            if kind == "sep":
                ctk.CTkFrame(inner, fg_color=BORDER, height=1, corner_radius=0).pack(
                    fill=tk.X, padx=8, pady=4)
            elif kind == "header":
                ctk.CTkLabel(inner, text=item[1].upper(), font=_font(10, "medium"),
                             text_color=FAINT_TEXT, anchor="w", height=22).pack(fill=tk.X, padx=10)
            else:
                label, icon, command = item[1], item[2], item[3]
                options = item[4] if len(item) > 4 else {}
                colour = DANGER if options.get("danger") else TEXT_SOFT
                enabled = options.get("enabled", True)
                shown = label if len(label) <= 30 else label[:29].rstrip() + "…"
                ctk.CTkButton(
                    inner, text=f"  {shown}", image=_icon(icon, 15, colour if enabled else FAINT_TEXT),
                    compound="left", anchor="w", height=30, corner_radius=6,
                    fg_color="transparent", hover_color=NAV_SELECTED, text_color=colour,
                    text_color_disabled=FAINT_TEXT, font=_font(13),
                    state=tk.NORMAL if enabled else tk.DISABLED,
                    command=lambda c=command: self._choose(c),
                ).pack(fill=tk.X)
        f.place(x=0, y=0, relwidth=1)
        self.root.update_idletasks()
        w, h = self.WIDTH, f.winfo_reqheight()
        rx, ry = x_root - self.root.winfo_rootx(), y_root - self.root.winfo_rooty()
        W, H = self.root.winfo_width(), self.root.winfo_height()
        x = rx if rx + w + 8 <= W else rx - w
        self.upward = ry + h + 8 > H
        y = ry - h if self.upward else ry
        self.x, self.y = max(8, min(x, W - w - 8)), max(8, min(y, H - h - 8))
        self.w, self.h = w, h
        self._show(0.0)
        self.holder.lift()
        app.__dict__.setdefault("_menu_clicks_bound", False)
        if not app._menu_clicks_bound:
            # One binding for every menu: a click elsewhere closes the open one.
            self.root.bind("<Button-1>", lambda e: app._menu and app._menu._outside(e), add="+")
            self.root.bind("<Escape>", lambda _e: app._menu and app._menu.close(), add="+")
            app._menu_clicks_bound = True
        reduced = _prefers_reduced_motion()
        _animate(self.root, 0 if reduced else self.OPEN_MS, self._show, ease=EASE_OUT)

    def _show(self, p):
        """p=0: folded towards the click point; p=1: fully open."""
        if self.closed and p > 0.99:
            return
        visible = round(self.h * (0.78 + 0.22 * p))
        drift = 5 * (1 - p)
        y = self.y + (self.h - visible + drift if self.upward else -drift)
        try:
            self.holder.place(x=self.x, y=round(y), width=self.w, height=visible)
            # Opening upwards the panel is anchored at the bottom, so it
            # unfolds from the point that was clicked.
            self.frame.place(x=0, y=(visible - self.h) if self.upward else 0, relwidth=1)
        except tk.TclError:
            pass

    def _outside(self, event):
        if self.closed:
            return
        if str(event.widget).startswith(str(self.holder)):
            return
        self.close()

    def _choose(self, command):
        self.close(animate=False)
        self.root.after(10, command)

    def close(self, animate=True):
        if self.closed:
            return
        self.closed = True
        if self.app._menu is self:
            self.app._menu = None

        def gone():
            try:
                self.holder.destroy()
            except tk.TclError:
                pass

        if not animate or _prefers_reduced_motion():
            gone()
            return

        def fold(p):
            visible = round(self.h * (1 - 0.18 * p))
            y = self.y + (self.h - visible if self.upward else 0)
            try:
                self.holder.place(x=self.x, y=round(y), width=self.w, height=visible)
                self.frame.place(x=0, y=(visible - self.h) if self.upward else 0, relwidth=1)
            except tk.TclError:
                pass

        _animate(self.root, self.CLOSE_MS, fold, done=gone, ease=EASE_OUT)


class _OutlineBars:
    """macOS stand-in for the tour's dimmed layer, drawn inside the app's own
    window. Tk on a Mac can't make part of a window see-through, and small
    separate windows stay on the normal desktop when Topshot is in full
    screen or on another desktop - so the highlight is four thin green bars
    placed inside the main window. They fade in and out by colour; nothing
    is dimmed and nothing covers the card."""

    WIDTH = 3
    GAP = 4  # space between the part being explained and the frame

    def __init__(self, root, colour):
        self.root = root
        self.colour = colour
        self.level = 0.0
        self.rect = None
        self.bars = [tk.Frame(root, bg=BG_COLOR, bd=0, highlightthickness=0) for _ in range(4)]

    def _spots(self):
        x0, y0, x1, y1 = (round(v) for v in self.rect)
        g, t = self.GAP, self.WIDTH
        x0, y0, x1, y1 = x0 - g, y0 - g, x1 + g, y1 + g
        w, h = max(1, x1 - x0), max(1, y1 - y0)
        return ((x0, y0, w, t), (x0, y1 - t, w, t), (x0, y0, t, h), (x1 - t, y0, t, h))

    def _show(self):
        if self.rect is None or self.level <= 0.02:
            for bar in self.bars:
                try:
                    bar.place_forget()
                except tk.TclError:
                    pass
            return
        for bar, (x, y, w, h) in zip(self.bars, self._spots()):
            try:
                bar.place(x=x, y=y, width=w, height=h)
                bar.lift()
            except tk.TclError:
                pass

    def place(self, rect):
        self.rect = rect
        self._show()

    def alpha(self, value):
        self.level = max(0.0, min(1.0, value))
        colour = _mix(BG_COLOR, self.colour, self.level)
        for bar in self.bars:
            try:
                bar.configure(bg=colour)
            except tk.TclError:
                pass
        self._show()

    def lift(self):
        for bar in self.bars:
            try:
                bar.lift()
            except tk.TclError:
                pass

    def destroy(self):
        for bar in self.bars:
            try:
                bar.destroy()
            except tk.TclError:
                pass


class Tour:
    """A five-step guided tour. The window dims, the part being explained
    stays lit with a ring around it, and a card beside it says what it is.
    Moving on, the lit area glides and resizes to the next part in one eased
    motion. Clicking anywhere in the app outside the card ends the tour.

    The dimmed layer is a translucent window over the app with holes in it
    (a colour Windows treats as see-through): one over the part being
    explained, one over the card. The card itself is part of the main
    window, so it is always drawn, on the right monitor, at the right size.
    """

    KEY = "#ff00fe"      # drawn in this colour = see-through
    DIM = 0.62           # how dark the rest of the window goes
    PAD = 8              # space between a control and its ring
    MOVE_MS = 420        # the lit area gliding to the next part
    CARD_MS = 180

    def __init__(self, app, steps, on_done):
        self.app, self.root, self.steps, self.on_done = app, app.root, steps, on_done
        self.reduced = _prefers_reduced_motion()
        self.index = 0
        self.rect = None
        self.closed = False
        self.bubble = None
        self.bubble_box = None
        self._moving = False
        self._still = 0
        self.outline = _OutlineBars(self.root, "#34d399") if IS_MAC else None
        self.overlay = tk.Toplevel(self.root)
        if self.outline is not None:
            self.overlay.withdraw()  # kept only so the code below has a window
        self.overlay.overrideredirect(True)
        self.overlay.transient(self.root)
        self.overlay.configure(bg="#000000")
        try:
            self.overlay.attributes("-transparentcolor", self.KEY)
            self.overlay.attributes("-alpha", 0.0)
        except tk.TclError:
            pass
        self.canvas = tk.Canvas(self.overlay, bg="#000000", highlightthickness=0, bd=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.hole = self.canvas.create_polygon(0, 0, 0, 0, fill=self.KEY, outline="", smooth=True)
        self.card_hole = self.canvas.create_polygon(0, 0, 0, 0, fill=self.KEY, outline="",
                                                    smooth=True)
        self.ring = self.canvas.create_polygon(0, 0, 0, 0, fill="", outline="#34d399", width=2,
                                               smooth=True)
        self._place_overlay()
        # Never unbound (tkinter's unbind would drop other handlers too);
        # they do nothing once the tour is closed.
        # A click beside the card no longer ends the tour (one click to bring
        # the window forward used to close it): only Skip, Esc and Done do.
        self.canvas.bind("<Button-1>", lambda _e: "break")
        self.root.bind("<Button-1>", self._on_root_click, add="+")
        for key, action in (("<Escape>", self.close), ("<Right>", self.next),
                            ("<Left>", self.back)):
            self.root.bind(key, lambda _e, a=action: None if self.closed else a(), add="+")
        self._last_geometry = self._window_geometry()
        self._follow_window()
        _fine_timer(True)
        _animate(self.root, 0 if self.reduced else 240,
                 lambda p: self._overlay_alpha(self.DIM * p))
        self._go(0, first=True)

    # --- geometry ---

    def _window_geometry(self):
        r = self.root
        return (r.winfo_rootx(), r.winfo_rooty(), r.winfo_width(), r.winfo_height())

    def _place_overlay(self):
        if self.outline is not None:
            return
        x, y, w, h = self._window_geometry()
        self.overlay.geometry(f"{w}x{h}+{x}+{y}")
        self.canvas.configure(width=w, height=h)
        self.overlay.lift()

    def _overlay_alpha(self, value):
        if self.outline is not None:
            self.outline.alpha(value / self.DIM)
            return
        try:
            self.overlay.attributes("-alpha", value)
        except tk.TclError:
            pass

    def _target_rect(self, step):
        """The area to light, in window coordinates: one widget or the span of several."""
        widgets = self.steps[step][0]()
        if not isinstance(widgets, (list, tuple)):
            widgets = [widgets]
        self.root.update_idletasks()
        ox, oy = self.root.winfo_rootx(), self.root.winfo_rooty()
        boxes = [(w.winfo_rootx(), w.winfo_rooty(),
                  w.winfo_rootx() + w.winfo_width(), w.winfo_rooty() + w.winfo_height())
                 for w in widgets]
        x0 = min(b[0] for b in boxes) - ox - self.PAD
        y0 = min(b[1] for b in boxes) - oy - self.PAD
        x1 = max(b[2] for b in boxes) - ox + self.PAD
        y1 = max(b[3] for b in boxes) - oy + self.PAD
        # Never past the window's edges: the photo grid runs on below the
        # window, so its ring stops just above the bottom edge.
        w, h = self.root.winfo_width(), self.root.winfo_height()
        x1 = min(x1, w - self.PAD)
        y1 = min(y1, h - self.PAD)
        return [x0, y0, x1, y1]

    @staticmethod
    def _rounded(x0, y0, x1, y1, r=12):
        r = max(2, min(r, (x1 - x0) / 2, (y1 - y0) / 2))
        return [x0 + r, y0, x1 - r, y0, x1, y0, x1, y0 + r, x1, y1 - r, x1, y1,
                x1 - r, y1, x0 + r, y1, x0, y1, x0, y1 - r, x0, y0 + r, x0, y0]

    def _draw(self, rect):
        self.rect = rect
        if self.outline is not None:
            self.outline.place(rect)
            return
        self.canvas.coords(self.hole, *self._rounded(*rect))
        x0, y0, x1, y1 = rect
        self.canvas.coords(self.ring, *self._rounded(x0 - 2, y0 - 2, x1 + 2, y1 + 2, 14))

    def _follow_window(self):
        """While the window is being moved or resized (or dragged to another
        monitor) the dimming steps aside instead of chasing it; once it has
        been still for a moment, everything is re-measured and fades back."""
        if self.closed:
            return
        try:
            geometry = self._window_geometry()
        except tk.TclError:
            return
        if geometry != self._last_geometry:
            if not self._moving:
                self._moving = True
                self._overlay_alpha(0.0)
                _log_reuse(f"TOUR window moving: {self._last_geometry} -> {geometry}")
            self._still = 0
        elif self._moving:
            self._still += 1
            if self._still >= 4:  # about 160 ms without movement
                self._moving = False
                try:
                    self._place_overlay()
                    self._draw(self._target_rect(self.index))
                    self._place_bubble()
                except Exception:
                    pass
                _animate(self.root, 0 if self.reduced else 180,
                         lambda p: self._overlay_alpha(self.DIM * p))
        self._last_geometry = geometry
        if not self._moving:
            self._keep_on_top()
        self.root.after(40, self._follow_window)

    def _keep_on_top(self):
        """Clicking the app (or its taskbar button) raises the main window
        above the dimmed layer, which hid the highlight. While the app is the
        active window the layer is put back on top; when another app is in
        front it is left alone, so it never covers other programs."""
        if self.outline is not None:
            try:
                if self.root.focus_displayof() is not None:
                    self.outline.lift()
            except (tk.TclError, KeyError):
                pass
            return
        try:
            import ctypes

            user32 = ctypes.windll.user32
            user32.GetParent.restype = ctypes.c_void_p
            user32.GetForegroundWindow.restype = ctypes.c_void_p
            mine = {user32.GetParent(self.root.winfo_id()),
                    user32.GetParent(self.overlay.winfo_id())}
            if user32.GetForegroundWindow() in mine:
                self.overlay.lift()
        except Exception:
            pass

    def _on_root_click(self, event):
        if self.closed:
            return
        return  # see __init__: clicks elsewhere leave the tour open

    # --- steps ---

    def _settled(self):
        """Ignores a second step request that arrives while the light is still
        travelling - a held arrow key would otherwise skip whole steps."""
        now = time.perf_counter()
        if now - getattr(self, "_last_step_at", 0.0) < 0.3:
            return False
        self._last_step_at = now
        return True

    def next(self):
        if not self._settled():
            return
        if self.index + 1 >= len(self.steps):
            self.close()
        else:
            self._go(self.index + 1)

    def back(self):
        if self.index > 0 and self._settled():
            self._go(self.index - 1)

    def _go(self, index, first=False):
        if self.closed:
            return
        try:
            target = self._target_rect(index)
        except Exception:
            _log_crash("tour step", sys.exc_info())
            self.close(completed=False)
            return
        self.index = index
        start = self.rect or target
        self._remove_bubble()
        token = self._move_token = object()

        def step(p):
            # Position and size travel together on one curve, so the light
            # reads as one shape gliding and reshaping, not jumping.
            if self._move_token is token and not self.closed:
                self._draw([a + (b - a) * p for a, b in zip(start, target)])

        def arrived():
            if self._move_token is token and not self.closed:
                try:
                    self._build_bubble()
                except Exception:
                    _log_crash("tour card", sys.exc_info())

        _animate(self.root, 0 if (self.reduced or first) else self.MOVE_MS, step,
                 done=arrived, ease=EASE_IN_OUT)

    # --- the card ---

    def _remove_bubble(self):
        bubble, self.bubble, self.bubble_box = self.bubble, None, None
        try:
            self.canvas.coords(self.card_hole, 0, 0, 0, 0)
        except tk.TclError:
            pass
        if bubble is not None:
            try:
                bubble.destroy()
            except tk.TclError:
                pass

    def _build_bubble(self):
        _getter, title, text = self.steps[self.index]
        total = len(self.steps)
        last = self.index + 1 == total
        # A plain Tk frame is placed and sized; the styled card fills it.
        # (CustomTkinter panels refuse a size from place(), which is what
        # kept the card from ever appearing.)
        holder = self.bubble = tk.Frame(self.root, bg=BG_COLOR, bd=0, highlightthickness=0)
        card = ctk.CTkFrame(holder, fg_color=SURFACE, corner_radius=12, border_width=1,
                            border_color=MENU_BORDER, bg_color=BG_COLOR)
        card.pack(fill=tk.BOTH, expand=True)
        body = ctk.CTkFrame(card, fg_color="transparent")
        body.pack(fill=tk.BOTH, expand=True, padx=18, pady=(15, 13))
        ctk.CTkLabel(body, text=title, font=_font(15, "semibold"), text_color=TEXT_COLOR,
                     anchor="w").pack(fill=tk.X)
        ctk.CTkLabel(body, text=text, font=_font(13), text_color=TEXT_SOFT, anchor="w",
                     justify="left", wraplength=300).pack(fill=tk.X, pady=(4, 14))
        row = ctk.CTkFrame(body, fg_color="transparent")
        row.pack(fill=tk.X)

        def arrow(icon, fallback, command, primary=False, enabled=True):
            ink = PRIMARY_TEXT if primary else (TEXT_COLOR if enabled else FAINT_TEXT)
            image = _icon(icon, 16, ink)
            return ctk.CTkButton(
                row, text="" if image else fallback, image=image, width=36, height=32,
                corner_radius=8, command=command, state=tk.NORMAL if enabled else tk.DISABLED,
                fg_color=PRIMARY if primary else "transparent",
                hover_color=PRIMARY_HOVER if primary else SURFACE_HOVER,
                border_width=0 if primary else 1, border_color=MENU_BORDER,
                text_color=ink, text_color_disabled=FAINT_TEXT, font=_font(14, "semibold"),
            )

        arrow("done" if last else "next", "✓" if last else "→", self.next, primary=True).pack(
            side=tk.RIGHT)
        ctk.CTkLabel(row, text=f"{self.index + 1}/{total}", font=_font(12, "medium"),
                     text_color=MUTED_TEXT_COLOR, width=38).pack(side=tk.RIGHT, padx=4)
        arrow("back", "←", self.back, enabled=self.index > 0).pack(side=tk.RIGHT)
        if not last:
            ctk.CTkButton(row, text="Skip", width=0, height=32, fg_color="transparent",
                          hover_color=SURFACE_HOVER, text_color=MUTED_TEXT_COLOR,
                          font=_font(12, "medium"), command=self.close).pack(side=tk.LEFT)
        self.root.update_idletasks()
        self.bubble_box = (card.winfo_reqwidth(), card.winfo_reqheight())
        rise = 0 if self.reduced else 8

        def appear(p):
            if self.bubble is holder:
                self._place_bubble(rise=round(rise * (1 - p)))

        _animate(self.root, 0 if self.reduced else self.CARD_MS, appear, ease=EASE_OUT)
        try:
            state = (f"overlay alpha={self.overlay.attributes('-alpha')} "
                     f"geometry={self.overlay.winfo_geometry()} mapped={self.overlay.winfo_ismapped()} "
                     f"moving={self._moving} window={self._window_geometry()}")
        except Exception as exc:
            state = repr(exc)
        _log_reuse(f"TOUR step {self.index + 1}: card {self.bubble_box}; {state}")

    def _place_bubble(self, rise=0):
        card = self.bubble
        if card is None or self.rect is None or self.bubble_box is None:
            return
        bw, bh = self.bubble_box
        x0, y0, x1, y1 = [round(v) for v in self.rect]
        w, h = self.root.winfo_width(), self.root.winfo_height()
        gap = 16
        # Beside the lit part if there is room, otherwise below or above it.
        if x1 + gap + bw <= w:
            x, y = x1 + gap, y0
        elif y1 + gap + bh <= h:
            x, y = x0, y1 + gap
        elif y0 - gap - bh >= 0:
            x, y = x0, y0 - gap - bh
        else:
            x, y = max(0, x0 - gap - bw), y0
        x = max(8, min(x, w - bw - 8))
        y = max(8, min(y, h - bh - 8)) + rise
        try:
            # Width is fixed; the height follows the content, so the last
            # row (the arrows) can never be cut off.
            card.place(x=x, y=y, width=bw)
            card.lift()
            card.update_idletasks()
            real = card.winfo_height()
            if real > 1 and real != bh:
                self.bubble_box = (bw, real)
                bh = real
            self.canvas.coords(self.card_hole, *self._rounded(x, y, x + bw, y + bh, 12))
        except (tk.TclError, ValueError):
            _log_crash("tour card placement", sys.exc_info())

    def close(self, completed=True):
        if self.closed:
            return
        self.closed = True
        self._remove_bubble()

        def finish():
            try:
                self.overlay.destroy()
            except tk.TclError:
                pass
            if self.outline is not None:
                self.outline.destroy()
            _fine_timer(False)
            self.on_done(completed)

        _animate(self.root, 0 if self.reduced else 200,
                 lambda p: self._overlay_alpha(self.DIM * (1 - p)), done=finish)


class BootSplash:
    """The card on screen from the moment the app starts until its window is
    ready.

    Starting used to be a blank pause, then a small window that jumped to
    full screen and filled in piece by piece. This makes it one composed
    motion: the card arrives, its contents follow in a short cascade, the
    bar reports real progress, and the card dissolves into a window that is
    already finished.
    """

    WIDTH, HEIGHT = 460, 300
    CARD_RISE = 8  # px the card travels as it arrives
    ITEM_RISE = 8  # px each line of content travels as it follows
    CARD_IN_MS = 220
    ITEM_IN_MS = 420
    STATUS_OUT_MS, STATUS_IN_MS = 130, 210
    BAR_MS = 420
    GLOW_PERIOD_S = 1.8

    CARD_W, CARD_H = 460, 300

    def __init__(self, root, full=False, static=False):
        """full: fill the screen like the maximised window instead of a
        small card. static: appear complete at once, with no entrance -
        used when continuing a screen another window is already showing."""
        self.root = root
        self.reduced = _prefers_reduced_motion()
        self.full = full
        self.static = static
        self._closing = False
        win = self.win = tk.Toplevel(root)
        win.withdraw()
        win.overrideredirect(True)
        win.configure(bg=SPLASH_BG)
        try:
            win.attributes("-alpha", 0.0)
            win.attributes("-topmost", True)
        except tk.TclError:
            pass
        screen_w, screen_h = win.winfo_screenwidth(), win.winfo_screenheight()
        if full:
            left, top, right, bottom = _work_area(win)
            self.WIDTH, self.HEIGHT = right - left, bottom - top
            self.x, self.y = left, top
            # The card's contents, centred on the screen (a little high:
            # the optical centre of a screen sits above the geometric one).
            ox = (self.WIDTH - self.CARD_W) // 2
            oy = (self.HEIGHT - self.CARD_H) // 2 - self.HEIGHT // 40
        else:
            self.x = (screen_w - self.WIDTH) // 2
            # A little above centre: the optical centre of a screen sits higher
            # than the geometric one, and a dead-centred card looks low.
            self.y = (screen_h - self.HEIGHT) // 2 - screen_h // 40
            ox = oy = 0
        self._place(0 if (self.reduced or full) else self.CARD_RISE)

        w, h = self.WIDTH, self.HEIGHT
        cx = ox + self.CARD_W // 2
        c = self.canvas = tk.Canvas(
            win, width=w, height=h, bg=SPLASH_BG, highlightthickness=0, bd=0
        )
        c.pack()
        # Everything starts in the background colour and is brought in by
        # the entrance cascade; each entry is (start delay ms, apply(0..1)).
        self._entrance = []

        try:
            self._shield_frames = _fade_frames(_load_mark(84), 14)
            shield = c.create_image(cx, oy + 86, image=self._shield_frames[0])
            self._entrance.append(
                (40, lambda q: self._fade_image(shield, self._shield_frames, cx, oy + 86, q))
            )
        except Exception:
            self._shield_frames = []

        title = c.create_text(cx, oy + 160, text=APP_NAME, fill=SPLASH_BG,
                              font=(UI_FONT_SEMIBOLD, -26))
        self._entrance.append((90, lambda q: self._fade_text(title, TEXT_COLOR, cx, oy + 160, q)))
        subtitle = c.create_text(cx, oy + 190, text="Marketing photo library", fill=SPLASH_BG,
                                 font=(UI_FONT, -14))
        self._entrance.append(
            (130, lambda q: self._fade_text(subtitle, MUTED_TEXT_COLOR, cx, oy + 190, q))
        )

        self._bar_w = 200
        self._bar_x = ox + (self.CARD_W - self._bar_w) // 2
        self._bar_y = oy + 228
        track = c.create_rectangle(self._bar_x, self._bar_y, self._bar_x + self._bar_w,
                                   self._bar_y + 3, fill=SPLASH_BG, width=0)
        self._bar = c.create_rectangle(self._bar_x, self._bar_y, self._bar_x,
                                       self._bar_y + 3, fill=SPLASH_BG, width=0)
        try:
            self._glow_frames = _glow_frames()
            self._glow = c.create_image(self._bar_x, self._bar_y + 1.5,
                                        image=self._glow_frames[0])
        except Exception:
            self._glow_frames, self._glow = [], None
        self._glow_in = 0.0

        def bar_in(q):
            c.itemconfigure(track, fill=_mix(SPLASH_BG, SPLASH_TRACK, q))
            c.itemconfigure(self._bar, fill=_mix(SPLASH_BG, SPLASH_ACCENT, q))
            self._glow_in = q

        self._entrance.append((170, bar_in))

        self._status_xy = (cx, oy + 252)
        self._status = c.create_text(*self._status_xy, text="Starting…", fill=SPLASH_BG,
                                     font=(UI_FONT, -12))
        self._entrance.append(
            (200, lambda q: self._fade_text(self._status, MUTED_TEXT_COLOR, *self._status_xy, q))
        )
        stamp = _build_stamp()
        if stamp:
            version = c.create_text(ox + self.CARD_W - 16, oy + self.CARD_H - 12,
                                    text=f"version {APP_VERSION}", anchor="se",
                                    fill=SPLASH_BG, font=(UI_FONT, -11))
            self._entrance.append(
                (240, lambda q: c.itemconfigure(version, fill=_mix(SPLASH_BG, SPLASH_FAINT, q)))
            )

        self._drawn = 0.0
        self._target = 0.12
        self._bar_token = None
        self._status_token = None
        self._draw_bar(0.12)

    # --- drawing primitives ---

    def _place(self, offset):
        self.win.geometry(f"{self.WIDTH}x{self.HEIGHT}+{self.x}+{self.y + offset}")

    def _alpha(self, value):
        self.win.attributes("-alpha", max(0.0, min(1.0, value)))

    def _rise(self):
        return 0 if self.reduced else self.ITEM_RISE

    def _fade_text(self, item, colour, x, y, q):
        self.canvas.itemconfigure(item, fill=_mix(SPLASH_BG, colour, q))
        self.canvas.coords(item, x, y + self._rise() * (1 - q))

    def _fade_image(self, item, frames, x, y, q):
        self.canvas.itemconfigure(item, image=frames[round(q * (len(frames) - 1))])
        self.canvas.coords(item, x, y + self._rise() * (1 - q))

    def _draw_bar(self, fraction):
        self._drawn = fraction
        right = self._bar_x + round(self._bar_w * fraction)
        self.canvas.coords(self._bar, self._bar_x, self._bar_y, right, self._bar_y + 3)
        if self._glow is not None:
            self.canvas.coords(self._glow, right, self._bar_y + 1.5)

    # --- motion ---

    def show(self, on_shown):
        self.win.deiconify()
        self.win.update_idletasks()
        if not self.full:
            _round_corners(self.win)
        self._breathe()
        if self.static:
            for _delay, apply in self._entrance:
                apply(1.0)
            self._alpha(1.0)
            self.root.after(0, on_shown)
            return
        if self.reduced:
            for _delay, apply in self._entrance:
                apply(1.0)
            _animate(self.root, 120, self._alpha, done=on_shown)
            return
        last = max((delay for delay, _ in self._entrance), default=0)
        total = max(self.CARD_IN_MS, last + self.ITEM_IN_MS)

        def frame(p):
            t = p * total
            card = EASE_OUT(min(1.0, t / self.CARD_IN_MS))
            self._alpha(card)
            if not self.full:
                self._place(round(self.CARD_RISE * (1 - card)))
            for delay, apply in self._entrance:
                apply(EASE_OUT(max(0.0, min(1.0, (t - delay) / self.ITEM_IN_MS))))

        # One clock drives the whole cascade, each part easing on its own
        # schedule, so the pieces can never drift out of step.
        _animate(self.root, total, frame, done=on_shown, ease=_linear)

    def _breathe(self):
        """Keeps the light at the bar's tip gently pulsing while the app
        loads - constant, slow motion that says "still working", not frozen."""
        if self._glow is None or not self._glow_frames:
            return
        start = time.perf_counter()
        frames = self._glow_frames

        def tick():
            if self._closing:
                return
            try:
                if self.reduced:
                    level = self._glow_in * 0.7
                else:
                    phase = (time.perf_counter() - start) / self.GLOW_PERIOD_S
                    breath = 0.5 - 0.5 * math.cos(2 * math.pi * phase)  # smooth 0..1..0
                    level = self._glow_in * (0.5 + 0.5 * breath)
                self.canvas.itemconfigure(self._glow, image=frames[round(level * (len(frames) - 1))])
                self.root.after(16, tick)
            except tk.TclError:
                pass

        tick()

    def set_status(self, text):
        """Crossfades the status line: the old text lifts away, the new one
        rises into place. An instant swap reads as a flicker."""
        x, y = self._status_xy
        if self.reduced:
            self.canvas.itemconfigure(self._status, text=text)
            return
        token = self._status_token = object()
        out_ms, in_ms = self.STATUS_OUT_MS, self.STATUS_IN_MS
        swapped = [False]

        def step(p):
            if self._status_token is not token:
                return  # a newer status has taken over
            t = p * (out_ms + in_ms)
            if t < out_ms:
                q = EASE_OUT(t / out_ms)
                self.canvas.itemconfigure(self._status, fill=_mix(MUTED_TEXT_COLOR, SPLASH_BG, q))
                self.canvas.coords(self._status, x, y - 4 * q)
                return
            if not swapped[0]:
                self.canvas.itemconfigure(self._status, text=text)
                swapped[0] = True
            q = EASE_OUT(min(1.0, (t - out_ms) / in_ms))
            self.canvas.itemconfigure(self._status, fill=_mix(SPLASH_BG, MUTED_TEXT_COLOR, q))
            self.canvas.coords(self._status, x, y + 4 * (1 - q))

        _animate(self.root, out_ms + in_ms, step, ease=_linear)

    def set_progress(self, target):
        """Moves the bar to `target` (0-1). Eased in and out, since it is
        travelling across the card; and only ever forwards - a bar that
        slides back reads as something going wrong."""
        target = max(self._target, min(1.0, target))
        self._target = target
        start = self._drawn
        token = self._bar_token = object()

        def step(p):
            if self._bar_token is token:  # a newer target has taken over
                self._draw_bar(start + (target - start) * p)

        _animate(self.root, 0 if self.reduced else self.BAR_MS, step, ease=EASE_IN_OUT)

    def hide(self, done=None):
        """Fades out faster than it came in - leaving should never hold
        anything up - and keeps drifting the way it arrived."""
        drift = 0 if (self.reduced or self.full) else 6

        def step(p):
            self._alpha(1 - p)
            self._place(-round(drift * p))

        def finished():
            self.destroy()
            if done is not None:
                done()

        _animate(self.root, 120 if self.reduced else 200, step, done=finished)

    def destroy(self):
        self._closing = True
        try:
            self.win.destroy()
        except tk.TclError:
            pass


class BootSequence:
    """Runs startup: splash in, build the window out of sight, reveal it.

    The window is built and laid out while fully transparent, so by the time
    it fades in it is already maximised, its sidebar is complete and its
    first screen of photos is drawn. Nothing on it moves after it appears.
    """

    MIN_VISIBLE_MS = 350  # never flash the card in and straight out again
    TIMEOUT_MS = 12000  # reveal regardless, rather than hide a problem

    def __init__(self, root):
        self.root = root
        self.app = None
        self.splash = None
        self._shown_at = None
        self._render_done = False
        self._collections_done = False
        self._revealing = False

    def start(self):
        _fine_timer(True)
        try:
            self.splash = BootSplash(self.root, full=True, static=AFTER_UPDATE)
            if AFTER_UPDATE:
                self.splash.set_status(f"Starting version {APP_VERSION}…")
            _boot_mark("splash built")
            self.splash.show(on_shown=self._on_splash_shown)
        except Exception:
            _log_crash("showing the splash", sys.exc_info())
            if self.splash is not None:
                self.splash.destroy()
            self.splash = None
            self.root.after(0, self._on_splash_shown)
        self.root.after(self.TIMEOUT_MS, self._reveal)

    def _on_splash_shown(self):
        _boot_mark("splash on screen")
        if AFTER_UPDATE:
            # Tells the old version its screen is now covered and it can go.
            try:
                with open(_handover_path(), "w", encoding="utf-8") as f:
                    f.write(str(os.getpid()))
            except OSError:
                pass
        self._shown_at = time.perf_counter()
        # Built while the card is still: building the window holds the main
        # thread for a few hundred ms, and a text change started now would
        # freeze half-way through.
        self.root.after(16, self._build)

    def _build(self):
        try:
            self.app = ImageSearchApp(self.root, boot=self)
        except Exception:
            _log_crash("building the window", sys.exc_info())
            self._fail()
            return
        _boot_mark("window built")
        self._stage("Loading your photo library…", 0.4)

    # --- called by the app as startup progresses ---

    def library_loaded(self, count):
        _boot_mark(f"photo library loaded ({count:,} photos)")
        self._stage(f"Preparing {count:,} photos…", 0.75)

    def library_empty(self):
        _boot_mark("no photo library cached yet")
        self._render_done = True
        self._maybe_reveal()

    def render_done(self):
        _boot_mark("first screen of photos drawn")
        self._render_done = True
        self._maybe_reveal()

    def collections_done(self):
        self._collections_done = True
        self._maybe_reveal()

    def _maybe_reveal(self):
        if self._render_done and self._collections_done:
            self._reveal()

    # --- the handover ---

    def _stage(self, text, progress=None):
        if self.splash is None:
            return
        if text is not None:
            self.splash.set_status(text)
        if progress is not None:
            self.splash.set_progress(progress)

    def _reveal(self):
        if self._revealing or self.app is None:
            return  # still building; the build's own progress will get here
        self._revealing = True
        shown_ms = 0 if self._shown_at is None else (time.perf_counter() - self._shown_at) * 1000
        wait = max(0, self.MIN_VISIBLE_MS - shown_ms)
        self._stage("Ready", 1.0)
        # Long enough for the bar to visibly arrive before the card leaves.
        self.root.after(int(wait) + 460, self._crossfade)

    def _crossfade(self):
        root = self.root
        reduced = self.splash.reduced if self.splash is not None else _prefers_reduced_motion()
        try:
            if root.state() == "withdrawn":
                root.deiconify()
        except tk.TclError:
            pass

        def step(p):
            root.attributes("-alpha", p)

        # The window comes up as the card goes down: one crossfade, with the
        # card's exit shorter than the window's entrance.
        _animate(root, 150 if reduced else 320, step, done=self._finished)
        if self.splash is not None:
            self.splash.hide()
            self.splash = None

    def _finished(self):
        _fine_timer(False)
        try:
            root = self.root
            root.attributes("-alpha", 1.0)
            root.lift()
            root.focus_force()
        except Exception:
            pass
        _boot_mark("window revealed")
        if self.app is not None:
            # After the reveal, not during it: importing torch holds the
            # interpreter for seconds and would freeze the crossfade.
            self.root.after(700, self.app._warm_model)
            # Shared collections and updates are looked up quietly a few
            # seconds later, from the saved sign-in only - no browser pops up.
            self.root.after(5000, self.app._start_background_checks)
            self.root.after(50, lambda: _set_window_icons(self.root))
            self.root.after(1200, self.app._first_screen_extras)
            if AFTER_UPDATE:
                self.root.after(400, lambda: self.app.status_var.set(
                    f"Updated to version {APP_VERSION}. Everything you had is still here."
                ))

    def _fail(self):
        _fine_timer(False)
        if self.splash is not None:
            self.splash.destroy()
            self.splash = None
        messagebox.showerror(
            APP_NAME,
            f"{APP_NAME} could not start. The details were written to "
            f"crash.log in {DATA_DIR}.",
        )
        self.root.destroy()


def _thread_excepthook(args):
    _log_crash("unhandled thread exception", (args.exc_type, args.exc_value, args.exc_traceback))


def main():
    threading.excepthook = _thread_excepthook
    _boot_mark("python and libraries loaded")
    _boot_mark(
        f"    of which: customtkinter {_T_CTK - _IMPORT_T0:.2f}s, numpy {_T_NUMPY - _T_CTK:.2f}s, "
        f"Pillow {_T_PIL - _T_NUMPY:.2f}s, search engine {_T_ENGINE - _T_PIL:.2f}s"
    )
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
    # One typeface throughout. CustomTkinter defaults to Roboto, which mixed
    # with the Segoe UI used for headings made the window look assembled.
    try:
        font_theme = ctk.ThemeManager.theme["CTkFont"]
        if "family" in font_theme:
            font_theme["family"] = UI_FONT
        for platform_font in font_theme.values():
            if isinstance(platform_font, dict) and "family" in platform_font:
                platform_font["family"] = UI_FONT
    except Exception:
        pass
    # CustomTkinter re-scales every widget when the window moves to a monitor
    # with a different display scaling, which on a multi-monitor desk means
    # the whole grid is rebuilt mid-drag and the window judders. The app is
    # laid out in plain pixels and does not need that, so it is turned off.
    try:
        ctk.deactivate_automatic_dpi_awareness()
    except Exception:
        pass
    _register_bundled_fonts()
    root = ctk.CTk()
    # Errors inside button clicks, timers and animations are otherwise
    # printed to a console nobody sees; this keeps a record of them.
    root.report_callback_exception = lambda kind, value, tb: _log_crash(
        "UI callback", (kind, value, tb)
    )
    _start_freeze_watchdog(root)
    _choose_ui_fonts(root)
    try:
        font_theme = ctk.ThemeManager.theme["CTkFont"]
        if "family" in font_theme:
            font_theme["family"] = UI_FONT
        for platform_font in font_theme.values():
            if isinstance(platform_font, dict) and "family" in platform_font:
                platform_font["family"] = UI_FONT
    except Exception:
        pass
    # Hidden and fully transparent until BootSequence reveals it: the window
    # is built and laid out out of sight, so it never appears half-made.
    try:
        root.attributes("-alpha", 0.0)
    except Exception:
        pass
    root.withdraw()
    _boot_mark("window system started")
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
            _apply_icon(root)
        except Exception:
            pass
    BootSequence(root).start()
    root.after(0, lambda: _boot_mark("event loop running"))
    try:
        root.mainloop()
    except Exception:
        _log_crash("mainloop", sys.exc_info())
        raise


if __name__ == "__main__":
    main()
