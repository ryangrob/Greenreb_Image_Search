"""Starts Image Search - the newest downloaded update if there is one,
otherwise the code built into the .exe.

This file is the .exe's entry point and is the one part an update cannot
replace, so it is kept small and careful: if an update fails to load, it is
set aside and the built-in version starts instead.
"""

import importlib.util
import json
import os
import sys
import traceback

os.environ["IMAGESEARCH_LAUNCHER"] = "1"

if not getattr(sys, "frozen", False) and "HF_HOME" not in os.environ:
    # Running from source: the pre-downloaded model sits next to this file.
    # (An update's own folder has no model in it, so the app could not find
    # it relative to its own location.)
    _cache = os.path.join(os.path.dirname(os.path.abspath(__file__)), "model_cache")
    if os.path.isdir(_cache):
        os.environ["HF_HOME"] = _cache

import app_version as _bundled  # noqa: E402  the version inside the .exe

UPDATES_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "ImageSearch", "app"
)
STATE_FILE = os.path.join(UPDATES_DIR, "current.json")


def _log(message):
    try:
        os.makedirs(UPDATES_DIR, exist_ok=True)
        with open(os.path.join(UPDATES_DIR, "launcher.log"), "a", encoding="utf-8") as f:
            f.write(message.rstrip() + "\n")
    except OSError:
        pass


def _pick_update():
    """The folder of a downloaded update that is newer than the built-in
    code and made for this runtime, or None."""
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            state = json.load(f)
        version = str(state["version"])
        runtime = int(state.get("runtime", -1))
    except Exception:
        return None
    if runtime != _bundled.RUNTIME_VERSION:
        return None
    if _bundled.version_tuple(version) <= _bundled.version_tuple(_bundled.APP_VERSION):
        return None
    folder = os.path.join(UPDATES_DIR, version)
    if not all(os.path.isfile(os.path.join(folder, f"{m}.py")) for m in _bundled.APP_MODULES):
        return None
    return folder


def _load_update(folder):
    """Imports the update's modules in order, so each one's imports find
    the updated versions of the others."""
    for name in _bundled.APP_MODULES:
        path = os.path.join(folder, f"{name}.py")
        spec = importlib.util.spec_from_file_location(name, path)
        module = importlib.util.module_from_spec(spec)
        sys.modules[name] = module
        spec.loader.exec_module(module)
    return sys.modules["tkinter_app"]


def _set_aside(reason):
    _log(f"Update set aside: {reason}")
    try:
        os.replace(STATE_FILE, os.path.join(UPDATES_DIR, "current.bad.json"))
    except OSError:
        pass


def main():
    app = None
    folder = _pick_update()
    if folder is not None:
        try:
            app = _load_update(folder)
        except Exception:
            _set_aside(traceback.format_exc())
            for name in _bundled.APP_MODULES:
                sys.modules.pop(name, None)
            sys.modules["app_version"] = _bundled
            app = None
    if app is None:
        import tkinter_app as app  # the version built into the .exe
    app.main()


if __name__ == "__main__":
    main()
