"""Build the Last War tracker workbook from a snapshot JSON.

Usage: python3 build_sheet.py data/2026-10-13.json [more snapshots...]
The last snapshot drives every tab; all snapshots feed the History tab.
"""
import json
import sys
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HEAD_FILL = PatternFill("solid", fgColor="1F3A5F")
HEAD_FONT = Font(bold=True, color="FFFFFF")
FLAG_FILL = PatternFill("solid", fgColor="FFF2CC")
TANK_FILL = PatternFill("solid", fgColor="E2EFDA")

STRATEGY = [
    # (priority, area, action, free path, paid path, why)
    (1, "Troops", "Fill the 16,249 empty troop slots with Lv10 units",
     "Train Lv10 non-stop; save training speedups for Arms Race / VS 'unit training' days for double points",
     "Training speedup packs on VS unit day",
     "Unit power is 38.9M; capacity is 43,029 but only 26,780 units exist. Each Lv10 = 1,647 power -> up to ~26M power waiting."),
    (2, "Troops", "Promote the 7,290 Lv8 units to Lv10",
     "Use Promote in the Drill Ground (cheaper than training new)",
     "Same, plus speedups",
     "Lv8 -> Lv10 is +584 power each (~4.3M total) and much stronger in fights."),
    (3, "Skills", "Spend your 49,813 skill medals on Squad 1",
     "Order: Kimberly skill 1 (29/36) -> Murphy (22/34) -> Williams / Stetmann / Marshall (19-20/30)",
     "Skill medal packs only after free medals are spent",
     "You have ~4 upgrades' worth of medals sitting unused. Free power today."),
    (4, "Research", "Push Tank Mastery (11%)",
     "Tank Synergy HP/ATK/DEF/Damage nodes first, then Tank HP/ATK II; keep both research queues busy",
     "Research speedups in Alliance Duel research day",
     "Tank-only boosts every hero in Squad 1. Tech Center Lv30 requirement already met."),
    (5, "Gear", "Raise Williams' cannon + data chip and Murphy's data chip from 30 to 40",
     "Gear ore from events, Arms Race chests, alliance store",
     "Gear ore packs",
     "Only Squad 1 gear still under 40. Then work on gear stars for Williams/Stetmann/Marshall (0-2 stars)."),
    (6, "Exclusive Weapon", "Murphy EW 12 -> 20",
     "EW shards from season events/store; Lv20 unlocks the big squad-wide boost",
     "Season pass / EW shard packs (best value spend for a light spender)",
     "Kimberly is already 20. Murphy is your front-line tank; Lv20 is the cost/benefit sweet spot. Don't put shards into Tesla."),
    (7, "Hero EXP", "Use the 434.9M hero EXP on squad 2 & 3 heroes",
     "Monica 149 -> 150 (185M), then Maxwell 72 and Schuyler",
     "None needed",
     "All Squad 1 heroes are 150. Spare EXP raises squad 2 power for free."),
    (8, "Squads", "Make Squad 2 all-Aircraft",
     "Swap Monica (Tank) -> Morrison or Lucius (UR Aircraft, Lv150). Consider replacing Maxwell (Lv72) with the other",
     "None needed",
     "5 same-type heroes gives the biggest type bonus; Squad 2 has 4 Aircraft + 1 Tank. Check squad power after swapping."),
    (9, "Hero Stars", "Promote heroes whose Tier button is ready",
     "Swift (needs 5* to unlock EW), Carlie, Fiona, Schuyler, Maxwell, Farhad, Braz, Kane all show green Promote",
     "UR universal shard packs -> Swift first",
     "Free stars = free power. Prioritise heroes in a squad."),
    (10, "Wall of Honor", "Promote Honor levels on Squad 1 tanks",
     "Murphy 23, Kimberly 25, Williams 11, Marshall 10, Stetmann 0 - honor medals into tanks",
     "-",
     "Each level adds Tank HP/ATK/DEF % for every tank hero."),
    (11, "Gear", "Williams is your weakest front-liner: 27% damage resistance vs Murphy 50% / Kimberly 44%",
     "Promote Williams' armor + radar stars (1 star each) and raise cannon/chip 30 -> 40",
     "Gear ore / blueprint packs aimed at Williams",
     "Front row soaks the hits. Low resist on Williams is the main hole in Squad 1 (Tank Details tab)."),
    (12, "Season", "Season bonus is 0% on all five tanks",
     "Check the Season tab for unclaimed season-tech / season boost; join season events",
     "Season pass",
     "Season line shows +0% for Attack/HP/Defense on every tank hero."),
    (13, "Squad 4", "Unlock Squad 4 (research 4%)",
     "Finish Squad 4 research in the Tech tree",
     "Research speedups",
     "A 4th squad adds another march and room for more hero power."),
]


# Goal: Squad 1 from 29.56M to 40M without upgrading HQ (set 2026-10-13).
SQUAD1_PLAN = [
    ("1", "Stetmann skill 1 (Orb Lightning) 19 -> 20", "9,200 medals", "To do"),
    ("1", "Williams skill 1 (Blast Shell) 20 -> 21", "10,400 medals", "To do"),
    ("1", "Marshall skill 1 (Triad Harmony) 20 -> 21", "10,400 medals", "To do"),
    ("1", "Murphy skill 1 (Cannon Fire II) 22 -> 23", "12,800 medals", "To do"),
    ("1", "Save leftover ~7,000 medals toward next level (Stetmann 20 -> 21 is 10,400)", "-", "To do"),
    ("1", "Murphy data chip 30 -> 40", "Gear ore", "To do"),
    ("1", "Williams cannon 30 -> 40, then data chip 30 -> 40", "Gear ore", "To do"),
    ("1", "Check Season screen - season bonus shows 0% on all tanks", "-", "To do"),
    ("1", "Wall of Honor: honor medals into Stetmann (0), Marshall (10), Williams (11)", "Honor medals", "To do"),
    ("2", "Gear stars: Marshall -> Williams -> Stetmann", "Blueprints", "Later"),
    ("2", "Murphy Exclusive Weapon 12 -> 20", "EW shards", "Later"),
    ("2", "Tank Mastery research - keep queues busy", "Research", "Later"),
    ("3", "Kimberly Exclusive Weapon 20 -> 30; Marshall EW if available", "EW shards", "Later"),
    ("3", "Skills toward max; drone levels", "Medals / drone parts", "Later"),
]


def header(ws, cols, row=1):
    for i, c in enumerate(cols, 1):
        cell = ws.cell(row=row, column=i, value=c)
        cell.fill, cell.font = HEAD_FILL, HEAD_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.freeze_panes = ws.cell(row=row + 1, column=1)


def widths(ws, w):
    for i, x in enumerate(w, 1):
        ws.column_dimensions[get_column_letter(i)].width = x


def table(ws, cols, rows, w=None, flag_col=None):
    header(ws, cols)
    for r in rows:
        ws.append(list(r))
        if flag_col is not None and isinstance(r[flag_col], str) and "NOT CAPTURED" in r[flag_col]:
            for c in ws[ws.max_row]:
                c.fill = FLAG_FILL
    if w:
        widths(ws, w)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")


def pretty(cols):
    return [c.replace("_", " ").title() for c in cols]


def build(paths, out):
    snaps = [json.load(open(p)) for p in paths]
    s = snaps[-1]
    a = s["account"]
    wb = Workbook()

    ws = wb.active
    ws.title = "Dashboard"
    ws["A1"] = f"Last War snapshot - {s['snapshot_date']}"
    ws["A1"].font = Font(bold=True, size=14)
    rows = [
        ("HQ level", a["hq_level"]), ("VIP", a["vip"]), ("Total power", a["total_power"]),
        ("Hero power", a["hero_power"]), ("Unit (troop) power", a["unit_power"]),
        ("Building power", a["building_power"]), ("Decor power", a["decor_power"]),
        ("Drone power", a["drone_power"]), ("Tech power", a["tech_power"]),
        ("Overlord power", a["overlord_power"]), ("Drone level", a["drone_level"]),
        ("Unused hero EXP", a["hero_exp_available"]), ("Unused skill medals", a["skill_medals_available"]),
        ("Troops", f"{s['troop_summary']['unit_count']:,} / {s['troop_summary']['capacity']:,}"),
        ("Empty troop slots", s["troop_summary"]["empty_capacity"]),
    ]
    for i, (k, v) in enumerate(rows, 3):
        ws.cell(row=i, column=1, value=k).font = Font(bold=True)
        c = ws.cell(row=i, column=2, value=v)
        if isinstance(v, int) and v > 9999:
            c.number_format = "#,##0"
    ws.cell(row=len(rows) + 4, column=1, value="Squads").font = Font(bold=True)
    for j, sq in enumerate(s["squads"]):
        ws.cell(row=len(rows) + 5 + j, column=1, value=f"Squad {sq[0]}")
        ws.cell(row=len(rows) + 5 + j, column=2, value=f"{sq[1]}  {sq[7] or ''}")
    widths(ws, [26, 34])

    ws = wb.create_sheet("Strategy")
    table(ws, ["#", "Area", "Action", "Free path", "Paid path (light spend)", "Why"], STRATEGY,
          [5, 16, 38, 48, 34, 55])

    ws = wb.create_sheet("Squad 1 Plan")
    table(ws, ["Phase", "Step", "Cost", "Status"], SQUAD1_PLAN, [12, 70, 22, 12])

    ws = wb.create_sheet("Heroes")
    table(ws, pretty(s["hero_columns"]), s["heroes"], [12, 8, 10, 10, 7, 12, 12, 9, 8, 9, 8, 7])
    for row in ws.iter_rows(min_row=2):
        if row[2].value == "Tank":
            row[0].fill = TANK_FILL
        row[6].number_format = "#,##0"

    if "tank_details" in s:
        ws = wb.create_sheet("Tank Details")
        table(ws, pretty(s["tank_details_columns"]), s["tank_details"], [11] + [12] * 15)
        for row in ws.iter_rows(min_row=2):
            for c in row[1:4]:
                c.number_format = "#,##0"

    ws = wb.create_sheet("Skills")
    table(ws, ["Name", "Skill 1", "Skill 2", "Skill 3", "Skill 4", "Skill 1 Name", "Skill 1 Max"],
          s["skills"], [12, 8, 8, 8, 8, 26, 11], flag_col=5)

    ws = wb.create_sheet("Gear")
    table(ws, pretty(s["gear_columns"]), s["gear"], [16, 14, 14, 14, 14, 34])

    ws = wb.create_sheet("Exclusive Weapons")
    table(ws, pretty(s["exclusive_weapons_columns"]), s["exclusive_weapons"],
          [12, 14, 11, 12, 13, 11, 40], flag_col=1)

    ws = wb.create_sheet("Squads")
    table(ws, pretty(s["squads_columns"]), s["squads"], [7, 9, 11, 11, 11, 11, 11, 20])

    ws = wb.create_sheet("Troops")
    table(ws, pretty(s["troops_columns"]), s["troops"], [8, 14, 14, 10, 12, 13, 8, 8])
    r = ws.max_row + 2
    for k, v in s["troop_summary"].items():
        ws.cell(row=r, column=1, value=k.replace("_", " ").title()).font = Font(bold=True)
        ws.cell(row=r, column=3, value=v)
        r += 1

    ws = wb.create_sheet("Research")
    table(ws, ["Tree", "Progress"], [(k.replace("_", " ").title(), v) for k, v in s["research"].items()], [24, 40])

    ws = wb.create_sheet("Missing Data")
    table(ws, ["Area", "What's missing", "Impact"], s["missing"], [18, 80, 10])
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.fill = FLAG_FILL

    ws = wb.create_sheet("History")
    keys = ["total_power", "hero_power", "unit_power", "building_power", "drone_power", "tech_power", "decor_power"]
    table(ws, ["Date"] + [k.replace("_", " ").title() for k in keys] + ["Troops"],
          [[x["snapshot_date"]] + [x["account"][k] for k in keys] + [x["troop_summary"]["unit_count"]] for x in snaps],
          [12] + [15] * len(keys) + [10])
    for row in ws.iter_rows(min_row=2):
        for c in row[1:]:
            c.number_format = "#,##0"

    wb.save(out)
    print("wrote", out)


if __name__ == "__main__":
    paths = sys.argv[1:] or ["data/2026-10-13.json"]
    build(sorted(paths), "Last War Tracker.xlsx")
