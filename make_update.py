"""Packages the app's code as an update and (optionally) publishes it.

    py make_update.py            builds the package into dist_update\\
    py make_update.py --upload   builds it and uploads it to the releases
                                 folder on SharePoint (you sign in)

Everyone's app checks that folder, downloads the package, checks its
SHA-256 against latest.json and offers "Restart to update". Bump
APP_VERSION in app_version.py first - an update is only taken when its
version is higher than the one running.

If a release changes the .exe itself (a library, Python, the model), bump
RUNTIME_VERSION too and publish a full build: apps on the old runtime will
then say a new download is needed instead of loading code they can't run.
"""

import argparse
import datetime
import hashlib
import io
import json
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from app_version import APP_MODULES, APP_VERSION, RUNTIME_VERSION, WHATS_NEW  # noqa: E402

ASSET_EXTENSIONS = {".ttf", ".css", ".txt", ".json", ".png", ".ico"}
SKIP_ASSETS = {"app_icon_master.png"}  # the 1024px source of the icon; not used at runtime


def build_package():
    buffer = io.BytesIO()
    files = []
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
        for name in APP_MODULES:
            path = os.path.join(HERE, f"{name}.py")
            z.write(path, f"{name}.py")
            files.append(f"{name}.py")
        assets = os.path.join(HERE, "assets")
        for folder, _dirs, names in os.walk(assets):
            for name in sorted(names):
                ext = os.path.splitext(name)[1].lower()
                if ext not in ASSET_EXTENSIONS or ".bak" in name or name in SKIP_ASSETS:
                    continue
                full = os.path.join(folder, name)
                rel = os.path.relpath(full, HERE).replace(os.sep, "/")
                z.write(full, rel)
                files.append(rel)
    return buffer.getvalue(), files


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument("--upload", action="store_true", help="upload to the SharePoint releases folder")
    parser.add_argument("--notes", default="", help="one line shown to users (optional)")
    args = parser.parse_args()

    package, files = build_package()
    package_name = f"ImageSearch-update-{APP_VERSION}.zip"
    manifest = {
        "version": APP_VERSION,
        "runtime": RUNTIME_VERSION,
        "package": package_name,
        "sha256": hashlib.sha256(package).hexdigest(),
        "size": len(package),
        "published": datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds"),
        "notes": args.notes,
        "changes": list(WHATS_NEW),
    }
    out = os.path.join(HERE, "dist_update")
    os.makedirs(out, exist_ok=True)
    with open(os.path.join(out, package_name), "wb") as f:
        f.write(package)
    with open(os.path.join(out, "latest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2)
    print(f"Built {package_name}: {len(files)} files, {len(package) / 1e6:.1f} MB")
    print(f"  in {out}")

    if not args.upload:
        print("Not uploaded. Run again with --upload to publish it.")
        return

    import config
    from sharepoint_client import SharePointClient

    client = SharePointClient()
    client.sign_in()
    client.ensure_folder(config.UPDATE_SITE_PATH, config.UPDATE_FOLDER_PATH)
    # Package first, manifest last: an app must never see a manifest whose
    # package isn't there yet.
    client.upload_bytes_by_path(config.UPDATE_SITE_PATH, config.UPDATE_FOLDER_PATH, package_name, package)
    client.upload_bytes_by_path(
        config.UPDATE_SITE_PATH, config.UPDATE_FOLDER_PATH, "latest.json",
        json.dumps(manifest, indent=2).encode("utf-8"),
    )
    print(f"Published version {APP_VERSION} to {config.UPDATE_SITE_PATH} / {config.UPDATE_FOLDER_PATH}")
    print("Everyone gets it the next time their app checks (at start, then every few hours).")


if __name__ == "__main__":
    main()
