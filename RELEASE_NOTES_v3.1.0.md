# Greenreb Image Search v3.1.0

Search the marketing photo library on SharePoint by describing what you
need, or by showing it a picture. 33,000 photos, no tagging required.

## Install

1. Download **Greenreb_Image_Search_v3.1.0.zip** below (about 600 MB).
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

**A fresh, modern look, matching the Sales Cockpit.** Neutral near-black
greys instead of navy and blue, a full-height sidebar (Library,
Collections, and System actions at the foot), a page title that follows
what you're viewing, one rounded search field, the Inter typeface and
Lucide outline icons. The colour on screen now comes from the photos.

**A cleaner photo grid.** Photos sit in tidy rows that fill the full width
of the window, each keeping its own shape, with rounded corners and even
spacing - no more ragged gaps around portrait photos. The folder names
under every photo are gone; the folder is shown in the preview instead.
The star appears when you point at a photo, and favourites keep their
gold star.

**Move photos between collections.** Drag a photo from one collection
onto another and it moves: it leaves the first and appears in the second.
Dragging from All photos, a search or Favourites still adds a copy.
Right-click a photo inside a collection for **Move to ...**.

**Drag and drop that feels physical.** The photo lifts off the grid and
follows the pointer, the collection under it lights up, and on release it
flies into the collection, which glows to confirm. Dropped anywhere else,
it glides back to where it came from. Respects Windows' "Animation
effects" setting.

**Fixed:** switching from a long list you had scrolled down to a shorter
collection no longer shows an empty page; every new view now starts at the
top.

Everything from v3.0.0 is included: better ranking, Find similar,
collections, full-resolution preview and Copy, the startup intro, and
Clear search history.

## Known limitations

- Not code-signed, so Windows and some antivirus software will warn.
- No installer and no automatic updates — a new version means a new zip.
- Search is English-only.
- You need access to the SharePoint site to sign in.
