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
    tanks = []
    for name in squads[0]["heroes"]:
        rows = []
        for s in snaps:
            hero = next((h for h in s["heroes"] if h[0] == name), None)
            ew = next((e[1] for e in s["exclusive_weapons"] if e[0] == name), None)
            skills = next((k[1:4] for k in s["skills"] if k[0] == name), None)
            rows.append({
                "date": s["snapshot_date"],
                "power": hero[6] if hero else None,
                "ew": ew if isinstance(ew, int) else None,
                "honor": dict(s.get("wall_of_honor") or []).get(name),
                "skills": skills,
            })
        hero = next(h for h in latest["heroes"] if h[0] == name)
        tanks.append({"name": name, "role": hero[3], "attack": hero[7], "hp": hero[8], "defense": hero[9], "history": rows})
    keys = ["hero_power", "unit_power", "building_power", "drone_power", "tech_power", "overlord_power"]
    power_split = [{"date": s["snapshot_date"], **{k: s["account"].get(k) for k in keys}} for s in snaps]
    hero_split = [{"date": s["snapshot_date"], **s["account"]["hero_power_detail"]}
                  for s in snaps if isinstance(s["account"].get("hero_power_detail"), dict)]
    return {
        "site": site,
        "tanks": tanks,
        "power_split": power_split,
        "hero_split": hero_split,
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
