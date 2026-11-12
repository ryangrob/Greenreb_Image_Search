#!/bin/bash
# Builds Topshot.app and Topshot-<version>-mac.dmg on a Mac (Apple Silicon).
# Runs on GitHub (see .github/workflows/build-mac.yml) or on any Mac:
#   python3 -m pip install -r requirements.txt && ./build_mac.sh
# Use the python.org Python: it brings the Tk that the window is drawn with.
set -euo pipefail
cd "$(dirname "$0")"
PY="${PYTHON:-python3}"
VERSION="$("$PY" -c 'import app_version; print(app_version.APP_VERSION)')"
echo "Building Topshot $VERSION for macOS ($(uname -m))"

# The CLIP model goes inside the app, so nobody waits for a download.
if [ ! -d model_cache ]; then
    "$PY" download_model.py
fi

rm -rf build dist
"$PY" -m PyInstaller --noconfirm --onedir --windowed --name Topshot \
    --collect-all open_clip \
    --collect-all torch \
    --collect-all torchvision \
    --collect-all msal \
    --collect-all customtkinter \
    --hidden-import unicodedata \
    --hidden-import tkinter_app \
    --hidden-import app_version \
    --hidden-import sharepoint_client \
    --hidden-import config \
    --add-data "model_cache:model_cache" \
    --add-data "assets:assets" \
    --icon "assets/app_icon.icns" \
    --osx-bundle-identifier "com.greenreb.topshot" \
    launcher.py

APP="dist/Topshot.app"
PLIST="$APP/Contents/Info.plist"
/usr/libexec/PlistBuddy -c "Set :CFBundleShortVersionString $VERSION" "$PLIST" || \
    /usr/libexec/PlistBuddy -c "Add :CFBundleShortVersionString string $VERSION" "$PLIST"
/usr/libexec/PlistBuddy -c "Set :CFBundleVersion $VERSION" "$PLIST" || \
    /usr/libexec/PlistBuddy -c "Add :CFBundleVersion string $VERSION" "$PLIST"
/usr/libexec/PlistBuddy -c "Add :NSHighResolutionCapable bool true" "$PLIST" 2>/dev/null || true
/usr/libexec/PlistBuddy -c "Add :LSMinimumSystemVersion string 12.0" "$PLIST" 2>/dev/null || true
/usr/libexec/PlistBuddy -c "Add :NSHumanReadableCopyright string 'Ryan Grob, Greenreb'" "$PLIST" 2>/dev/null || true

# Editing Info.plist breaks PyInstaller's signature; sign again. Without an
# Apple Developer ID this is an ad-hoc signature: enough for Apple Silicon
# to run the app, but macOS asks once before opening it (see READ ME FIRST).
codesign --force --deep --sign "${SIGN_IDENTITY:--}" "$APP"
codesign --verify --deep --strict "$APP" && echo "Signature OK"

# The disk image: the app, a shortcut to Applications, and how to open it.
STAGE="dist/dmg"
rm -rf "$STAGE" && mkdir -p "$STAGE"
cp -R "$APP" "$STAGE/"
ln -s /Applications "$STAGE/Applications"
cat > "$STAGE/READ ME FIRST.txt" <<TXT
Topshot $VERSION for Mac

1. Drag Topshot onto the Applications folder.
2. Open Topshot from Applications (or Launchpad / Spotlight).
3. The first time, macOS says it can't check the app, because it is not
   signed with an Apple certificate yet. Click Done, then open
   System Settings > Privacy & Security, scroll down and click
   "Open Anyway" next to Topshot. Confirm with your password.
   This is needed once only.
4. Sign in with your Greenreb Microsoft account.

If macOS says the app "is damaged": open Terminal and run
   xattr -cr /Applications/Topshot.app
then open it again.

Logs, if something goes wrong:
   ~/Library/Application Support/Topshot/crash.log
Please send that file to Ryan Grob.
TXT
DMG="dist/Topshot-$VERSION-mac.dmg"
hdiutil create -volname "Topshot $VERSION" -srcfolder "$STAGE" -ov -format UDZO "$DMG"
echo "Done: $DMG ($(du -h "$DMG" | cut -f1))"
