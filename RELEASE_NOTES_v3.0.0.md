# Greenreb Image Search v3.0.0

Search the marketing photo library on SharePoint by describing what you
need, or by showing it a picture. 32,000 photos, no tagging required.

## Install

1. Download **Greenreb_Image_Search_v3.0.0.zip** below (about 600 MB).
2. **Right-click the zip → Properties → tick "Unblock" → OK.** Windows
   blocks downloaded programs, and skipping this causes the app to fail
   with no explanation.
3. Extract to a short path such as `C:\ImageSearch` — not Downloads.
   Windows has a 260-character path limit and a long path silently drops
   files during extraction.
4. Run `Greenreb_Image_Search.exe`.
5. Sign in with your Greenreb Microsoft account when prompted.

The window appears within a few seconds. The first search after starting
can take a few seconds longer while the search model loads. Windows
SmartScreen will warn that the publisher is unknown: choose **More info →
Run anyway**. The app is not code-signed yet.

## What's new

**Better search results.** The top results now keep their true relevance
order. Previously the list was reordered twice after ranking, pushing
genuinely better matches down the grid — a specific search gave you a few
good results followed by loose ones.

**Searches understand what you're asking for.** A search for a thing
("burger", "cocktail") leads with photos of that thing. A search for a
scene ("birthday party") leads with photos of guests. Previously every
search assumed you wanted people in shot.

**Find similar.** Pick any photo on your computer and find the ones in the
library that look like it. No words needed.

**Collections you curate.** Start with Golf, Food, Drinks, Family & Kids
and Groups, or make your own with the **+** in the sidebar. Fill them by
dragging a photo onto one, or right-click a photo. Right-click a collection
to rename or remove it - removing never deletes photos. Automatic categories were close but never right; a shortlist someone
chose is worth more.

**Preview and copy.** Double-click any photo for a large view inside the
app. **Copy** puts the full-resolution original on your clipboard, ready
to paste into PowerPoint or an email.

**Starts properly.** A short intro card while the app loads, then the
window appears complete - no blank pause, no window jumping to full screen,
nothing filling in piece by piece.

**Faster.** Search results and collections draw many times faster, and the
window no longer stutters when moved between monitors.

**Fixed:** un-starring a favourite now actually removes it.

## Known limitations

- Not code-signed, so Windows and some antivirus software will warn.
- No installer and no automatic updates — a new version means a new zip.
- Search is English-only.
- You need access to the SharePoint site to sign in.
