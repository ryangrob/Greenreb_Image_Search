# Topshot v3.2.0 (formerly Image Search)

Search the marketing photo library on SharePoint by describing what you
need, or by showing it a picture. 33,000 photos, no tagging required.

## Install

1. Download **Topshot_v3.2.0.zip** below (about 600 MB).
2. **Right-click the zip → Properties → tick "Unblock" → OK.** Windows
   blocks downloaded programs, and skipping this causes the app to fail
   with no explanation.
3. Extract to a short path such as `C:\ImageSearch` — not Downloads.
   Windows has a 260-character path limit and a long path silently drops
   files during extraction.
4. Run `Topshot.exe`.
5. Sign in with your Greenreb Microsoft account when prompted.

The window appears within a few seconds. The first search after starting
can take a few seconds longer while the search model loads. Windows
SmartScreen will warn that the publisher is unknown: choose **More info →
Run anyway**. The app is not code-signed yet.

## What's new

**A new name and icon: Topshot.** The app has its own identity now - a golf
ball that is also a camera aperture - with the interface in matching greens.
The icon stays sharp at every taskbar size.

**The app now updates itself.** It checks for a newer version when it
starts and every few hours, downloads it in the background (a few MB),
checks that it arrived intact, and shows **Restart to update** at the
bottom of the sidebar. **System → Check for updates** looks straight away.
This is the last version you need to download by hand - unless the
engine underneath changes, in which case the app tells you.

**Shared collections.** Click **+** next to Collections and choose **New
shared collection**: give it a name and the work email addresses of the
people to share it with. Everyone listed sees it in their sidebar (with a
people icon) and can drag photos in and out; photos added by others
appear automatically, or straight away with the sync button next to +.
The owner can rename it, change the people or delete it (right-click);
everyone else can leave it.

**A guided tour.** The first time you open the app, a short tour lights up
each part of the window and explains it in a sentence. Take it again any
time from About or Support.

**Support from inside the app.** System → Support: write what's wrong or
what you'd like, and it's sent straight to the app's owner from your own
Greenreb email.

**A smoother start.** The intro now fills the screen, and an update
reloads the window with the same animation.

**About.** System → About shows who made the app, how it was built and the
open-source parts it uses.

Everything from v3.1.0 is included.

## Known limitations

- Not code-signed, so Windows and some antivirus software will warn.
- No installer yet. Updates to the app arrive by themselves; a change to the
  engine (Python, libraries, AI model) still needs a new zip.
- Search is English-only.
- You need access to the SharePoint site to sign in.
