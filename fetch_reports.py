"""Downloads Topshot's error reports from SharePoint.

Run in C:\ImageSearchBuild-v2.4.1:   py fetch_reports.py          (new reports)
                                      py fetch_reports.py --all    (summarise all)

Every Topshot uploads the new entries of its crash.log and freeze.log to
the reports folder (config.REPORTS_SITE_PATH / REPORTS_FOLDER_PATH). This
signs in with your saved Microsoft sign-in, saves the reports that are not
in error_reports\ yet, and prints a summary, newest first. The full text is
in the .json files and in error_reports\SUMMARY.txt - hand those to Claude.
Nothing on SharePoint is changed.
"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

PREFIX = "report_"


def last_line(text):
    """The line that names the error: the last one that isn't indented."""
    lines = [l for l in (text or "").splitlines() if l.strip()]
    for line in reversed(lines):
        if not line.startswith((" ", "\t")):
            return line.strip()
    return lines[-1].strip() if lines else ""


def summarise(report, name):
    out = [f"{report.get('created', '?')}  machine {report.get('machine', '?')}  "
           f"Topshot {report.get('app_version', '?')} ({report.get('build', '?')})  "
           f"{report.get('system', '?')}",
           f"    user: {report.get('user') or 'unknown'}   photos: {report.get('photos', '?')}   "
           f"file: {name}"]
    for log, part in (report.get("logs") or {}).items():
        skipped = part.get("skipped_bytes") or 0
        out.append(f"    {log}" + (f"  (older {skipped:,} bytes left out)" if skipped else ""))
        for entry in part.get("entries") or []:
            times = f"  x{entry['count']}, last {entry.get('last', '')}" if entry.get("count", 1) > 1 else ""
            line = last_line(entry.get("text"))
            out.append(f"      {entry.get('at', '')}  {entry.get('what', '')}"
                       + (f"  ->  {line}" if line else "") + times)
    return "\n".join(out)


def main():
    import config
    from sharepoint_client import SharePointClient

    show_all = "--all" in sys.argv
    site = getattr(config, "REPORTS_SITE_PATH", None) or config.UPDATE_SITE_PATH
    folder = getattr(config, "REPORTS_FOLDER_PATH", None) or "General/Topshot Reports"
    out_dir = os.path.join(HERE, "error_reports")
    os.makedirs(out_dir, exist_ok=True)

    client = SharePointClient()
    client.sign_in()
    names = client.list_names_by_path(site, folder, PREFIX)
    new = [n for n in names if not os.path.exists(os.path.join(out_dir, n))]
    print(f"{len(names)} report(s) in {site} / {folder}; {len(new)} new.")
    for n, name in enumerate(new, start=1):
        raw = client.download_bytes_by_path(site, f"{folder}/{name}")
        if raw is None:
            continue
        with open(os.path.join(out_dir, name), "wb") as f:
            f.write(raw)
        print(f"  downloaded {n}/{len(new)}: {name}")

    chosen = sorted(os.listdir(out_dir)) if show_all else new
    reports = []
    for name in chosen:
        if not (name.startswith(PREFIX) and name.endswith(".json")):
            continue
        try:
            with open(os.path.join(out_dir, name), "r", encoding="utf-8") as f:
                reports.append((json.load(f), name))
        except (OSError, ValueError):
            print(f"  could not read {name}")
    reports.sort(key=lambda r: str(r[0].get("created", "")), reverse=True)
    if not reports:
        print("Nothing new." if not show_all else "No reports yet.")
        return
    text = "\n\n".join(summarise(report, name) for report, name in reports)
    print()
    print(text)
    with open(os.path.join(out_dir, "SUMMARY.txt"), "w", encoding="utf-8") as f:
        f.write(text + "\n")
    print(f"\nSaved in {out_dir} (SUMMARY.txt and one .json per report).")


if __name__ == "__main__":
    main()
