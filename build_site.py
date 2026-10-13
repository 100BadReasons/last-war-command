"""Build the static recap site into docs/ for GitHub Pages.

Usage: python3 build_site.py
Reads data/site.json (goal, plan, recaps) and every data/YYYY-MM-DD.json snapshot.
"""
import datetime
import glob
import json
import os
import re

ROOT = os.path.dirname(os.path.abspath(__file__))


def load():
    site = json.load(open(os.path.join(ROOT, "data", "site.json")))
    snaps = []
    for p in sorted(glob.glob(os.path.join(ROOT, "data", "*.json"))):
        if re.match(r"\d{4}-\d{2}-\d{2}\.json$", os.path.basename(p)):
            snaps.append(json.load(open(p)))
    return site, snaps


def power(v):
    """'29.56M' -> 29560000; ints pass through."""
    if isinstance(v, (int, float)):
        return v
    m = re.match(r"([\d.]+)\s*M", str(v))
    return int(float(m.group(1)) * 1_000_000) if m else None


def payload(site, snaps):
    latest = snaps[-1]
    hero_power = {h[0]: h[6] for h in latest["heroes"]}
    squads = []
    for sq in latest["squads"]:
        p = power(sq[1])
        if p:
            squads.append({"n": sq[0], "power": p, "heroes": [h for h in sq[2:7] if h], "mix": sq[7]})
    squad1 = [{"name": n, "power": hero_power.get(n)} for n in squads[0]["heroes"]]
    history = [{
        "date": s["snapshot_date"],
        "total": s["account"]["total_power"],
        "squad1": power(s["squads"][0][1]),
    } for s in snaps]
    return {
        "site": site,
        "updated": latest["snapshot_date"],
        "account": latest["account"],
        "squads": squads,
        "squad1": squad1,
        "history": history,
        "built": datetime.datetime.now().strftime("%b %-d, %-I:%M %p"),
    }


def main():
    site, snaps = load()
    data = json.dumps(payload(site, snaps), separators=(",", ":")).replace("</", "<\\/")
    tpl = open(os.path.join(ROOT, "site_template.html")).read()
    os.makedirs(os.path.join(ROOT, "docs"), exist_ok=True)
    out = os.path.join(ROOT, "docs", "index.html")
    open(out, "w").write(tpl.replace("/*__DATA__*/null", data))
    open(os.path.join(ROOT, "docs", ".nojekyll"), "w").close()
    print("wrote", out)


if __name__ == "__main__":
    main()
