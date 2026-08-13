# Image Search

A desktop app for searching a folder of photos by typing a description
("golden retriever on a beach") or by picking a reference image. Uses a
local CLIP model, so once the model is downloaded, searching works fully
offline. Can also pull images directly from a SharePoint document library.

## Files

- `tkinter_app.py` — the GUI.
- `search_engine.py` — CLIP indexing/search logic used by the GUI.
- `sharepoint_client.py` — signs in to Microsoft 365 and talks to Microsoft
  Graph to recursively walk a SharePoint folder tree, fetch thumbnails, and
  read/write each subfolder's shared index file.
- `config.py` — Azure AD app registration IDs and the SharePoint site/folder
  this app points at (already filled in).
- `requirements.txt` — Python dependencies.
- `download_model.py` — one-time helper that pre-downloads the CLIP weights
  into `model_cache/` so `build.bat` can bundle them into the `.exe`.
- `build.bat` — builds a standalone Windows app with PyInstaller.

## Run it from source

```
pip install -r requirements.txt
python tkinter_app.py
```

**Local folder:**
1. Click **Choose Folder...** and pick a folder of images.
2. Click **Index Folder**. The first time this runs it downloads the CLIP
   model (~350MB, one-time, needs internet) then embeds every image in the
   folder. An index cache file (`.imagesearch_cache.json`) is written into
   that folder so re-indexing later only processes new/changed files.

**SharePoint (Marketing Photos):**
1. Click **Search Marketing Photos**. The first time on a machine, a browser
   window opens to sign in with your Microsoft 365 account. Sign-in is saved
   to disk (`%LOCALAPPDATA%\ImageSearch\msal_cache.bin`) and reused silently
   on every future launch — no repeat browser prompts unless that saved
   sign-in actually expires.
2. The app automatically indexes every subfolder under `Fotos & Videos`
   (`SHAREPOINT_SEARCH_ROOT_PATH` in `config.py`) — there's no folder picker,
   and no separate "Index Folder" click needed. Only `.png`/`.jpg`/`.jpeg`
   files are considered (RAW formats like `.dng` are skipped entirely).
   It works off small thumbnails (not the full-resolution originals), and
   shares the resulting index by uploading a small `.imagesearch_sp_index.json`
   file back into each subfolder. The first person to index a given folder
   pays the one-time cost of downloading+embedding its thumbnails; everyone
   after that (any machine) just downloads that already-built index instead
   of redoing the work. Only when you double-click a search result to open
   it does the real full-resolution file get downloaded.

   Every click checks SharePoint for what's changed using Microsoft Graph's
   delta query, so after the first (necessarily slow, full) pass, later runs
   are fast — only genuinely new/changed/deleted images get touched, not a
   full rescan of the whole tree. This state is cached locally under
   `%LOCALAPPDATA%\ImageSearch\sharepoint_cache\FotosVideos\` (`delta_link.txt`,
   `item_folder_map.json`, plus the per-folder index/thumbnail cache) — delete
   that folder to force a full resync from scratch if it's ever needed.

**Searching (either source):**
- Type a description and click **Search**, or click **Search by Image...**
  and pick a photo to find similar ones.
- Double-click any result thumbnail to open it in your default viewer.

## Build a standalone .exe

```
pip install -r requirements.txt
build.bat
```

This produces `dist\Greenreb_Image_Search\Greenreb_Image_Search.exe`. Zip up
the whole `dist\Greenreb_Image_Search` folder (not just the .exe — it needs the files
PyInstaller collects alongside it) and send that to anyone; they can run
it on Windows without installing Python.

**The CLIP model is bundled into the build** (`build.bat` runs
`download_model.py` automatically the first time, if `model_cache/` doesn't
already exist, then packages it into the `.exe` via `--add-data`). This adds
~580MB to the build/zip size, but means an end user's very first click works
immediately — no waiting on a model download before they can start
indexing/searching. The only thing still requiring internet on first use is
signing in to Microsoft 365 and the SharePoint indexing itself.

## Troubleshooting

- **Build is huge / slow to start:** this is expected — it bundles torch.
  `--onedir` (used by `build.bat`) starts faster than a one-file build.
- **Antivirus flags the .exe:** common false positive with PyInstaller
  builds; codesigning it removes the warning but requires a certificate.
- **"Search Marketing Photos" fails to sign in / permission errors:** the
  Azure AD app (`config.py`) needs admin consent granted for
  `Sites.Read.All`, `Files.ReadWrite.All`, and `offline_access` in whichever
  tenant a given user signs in with (`Files.ReadWrite.All` is required, not
  just `Files.Read.All`, because the app uploads a shared thumbnail index
  back into each SharePoint folder it indexes). If a user's account belongs to a
  different tenant than the one the app was registered in (e.g. an
  external/guest account), sign-in may need the app registration's
  "Supported account types" changed to multitenant — see whoever set up
  `config.py` for details.
- **SharePoint search finds nothing / wrong folder:** the folder tree that
  gets indexed is set by `SHAREPOINT_SEARCH_ROOT_PATH` in `config.py`.
  Update it if the target folder path in SharePoint changes.
- **App closes/crashes unexpectedly during a long SharePoint indexing run:**
  check `%LOCALAPPDATA%\ImageSearch\crash.log` — any caught exception
  (including ones that only ever showed as a generic error dialog) is
  appended there with a full traceback, which is the fastest way to find the
  actual cause. If that file has no new entry at all after a crash, the
  failure happened below Python (e.g. the OS killing the process for using
  too much memory) rather than something Python itself could catch.
