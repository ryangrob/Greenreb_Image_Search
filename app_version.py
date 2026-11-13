"""Which version of the app's own code this is, and which runtime it needs.

The .exe has two parts. The runtime - Python, the libraries and the AI
model, about 600 MB - only changes with a full download. The app's own code
(this file, the screens, the search logic, fonts and icons) is about 4 MB
and can be updated on its own: the app downloads a newer copy from the
releases folder on SharePoint and launcher.py runs that copy instead of the
one inside the .exe.
"""

APP_VERSION = "3.8.0"

# What people see the app called (window title, sidebar, intro, About).
# Folders, settings and file names keep "ImageSearch" so a rename never
# loses anyone's cache, favourites or collections.
APP_NAME = "Topshot"
APP_AUTHOR = "Ryan Grob"

# Shown under "What's new" in the app, and sent with an update so people see
# what they are about to get. Keep each line short; the first two are the
# summary on the update card.
WHATS_NEW = [
    "Search in German too - „Burger mit Pommes“ works like “burger with fries”",
    "Drag photos straight into PowerPoint, Outlook, Teams or a folder - as the original",
    "Export with crop: 16:9, 9:16, Instagram or free, at full quality",
    "A bigger preview: zoom, previous and next, compare two photos",
    "Near-identical shots fold into a stack - click the number to open it",
    "Settings: light or dark, text size, English or German",
    "A sharper People filter, Ctrl+F for search, Details on right-click",
]

# Bump only when a release changes what the .exe itself contains (Python,
# a library, the model). Code-only updates cannot deliver those, so an
# update built for a different runtime asks for a full download instead.
RUNTIME_VERSION = 1

# The modules an update replaces, in load order: each may import the ones
# before it, so they must already be the updated versions when it loads.
APP_MODULES = ["app_version", "config", "sharepoint_client", "search_engine", "tkinter_app"]


def version_tuple(text):
    parts = []
    for piece in str(text).split("."):
        digits = "".join(ch for ch in piece if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts)
