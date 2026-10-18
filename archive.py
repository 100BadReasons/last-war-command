"""Archive completed plan steps at the start of an analysis run.

Usage: python3 archive.py YYYY-MM-DD
Reads the live checkmarks from the site's Google Sheet sync (sync_url in data/site.json).
Every step that is ticked (or already status "done") moves out of the plan and into
site.json "archive" under the given date. A step with "subs" moves only when every
sub-step is ticked; ticked sub-steps of an unfinished step are saved as status "done". Prints the archived step IDs so their rows
can be removed from the sheet's Site Checklist tab.
"""
import json
import os
import sys
import urllib.request

ROOT = os.path.dirname(os.path.abspath(__file__))
PATH = os.path.join(ROOT, "data", "site.json")


def main():
    date = sys.argv[1]
    site = json.load(open(PATH))
    with urllib.request.urlopen(site["sync_url"]) as r:
        checks = json.load(r).get("checks", {})
    moved = []
    for phase in site["plan"]:
        keep = []
        for step in phase["steps"]:
            subs = step.get("subs") or []
            for sub in subs:
                # keep ticks on partly finished steps; the sheet rows are cleared after archiving
                if checks.get(sub["id"], sub.get("status") == "done"):
                    sub["status"] = "done"
                else:
                    sub.pop("status", None)
            finished = all(s.get("status") == "done" for s in subs) if subs else checks.get(step["id"], step["status"] == "done")
            if finished:
                moved.append(dict(step, status="done", phase=phase["phase"], phase_name=phase["name"]))
            else:
                keep.append(step)
        phase["steps"] = keep
    if moved:
        archive = site.setdefault("archive", [])
        entry = next((a for a in archive if a["date"] == date), None)
        if entry is None:
            entry = {"date": date, "steps": []}
            archive.insert(0, entry)
        entry["steps"].extend(moved)
    json.dump(site, open(PATH, "w"), indent=1, ensure_ascii=False)
    for s in moved:
        print(s["id"])


if __name__ == "__main__":
    main()
