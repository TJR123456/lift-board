#!/usr/bin/env python3
"""Builds index.html from the lift sheets in ./data.

Every logged cell is a formula like 80*(8+6)+70*(9), i.e. weight*(reps+reps...).
This script reads those formulas back into individual sets, decides for each
lift whether you're ready to move up, and writes a self-contained page.

Run:  python3 build.py
"""
import datetime as dt
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

HERE = Path(__file__).resolve().parent
HISTORY = HERE / "data" / "Lifts (history).xlsx"      # old bro-split sheet, frozen
CURRENT = HERE / "data" / "New Lift Sheet.xlsx"        # refreshed daily
OUT = HERE / "index.html"

# ---- Rules you can tune -----------------------------------------------------
# Move up when you'd still average FLOOR reps a set at the next weight.
# Predicted average after a jump = avg - (30 + avg) * jump / weight (Epley).
FLOOR = 8
FLOOR_OVERRIDES = {              # high-rep lifts
    "cable forearm curl": 10, "wrist curl": 10, "face pulls": 10,
    "shrugs": 10, "cable lat raise": 10,
}
REP_CAP = 15                    # averaging this many at a weight = move up regardless
CLOSE_WITHIN = 1.0             # "close" when the prediction is this many reps under the floor
HOLD_SESSIONS = 6               # sessions in a row without a new best est. max = holding (normal on a cut)
FATIGUE_LIFTS = 4              # this many lifts dropping in one week = fatigue warning
FATIGUE_PCT = 5                # a "drop" = est. max down this % vs the last time on that day
TOO_HEAVY_AVG = 6              # average under this after a jump = drop back
# The average only counts sets at the day's heaviest weight, so ramp-up sets
# are left out. Lifts with a back-off marked like "S4 70%" in the header are
# judged on the max sets before it (S1-S3), even if one of those dropped weight.
# Start fresh comparisons after a form change (lift -> ISO date).
FORM_RESET = {"hammer curl": "2026-10-19", "wrist curl": "2026-11-06",
              "cable forearm curl": "2026-11-06"}
# Muscles each lift trains: (primary, secondary). Drawn as the body map.
# Based on EMG and anatomy references; primary = prime movers, secondary =
# synergists and stabilizers that do real work. Muscle ids match template.html.
MUSCLES = {
    # push
    "machine shoulder press": (["front_delts"], ["side_delts", "triceps", "upper_chest", "traps"]),
    "inc db bench": (["upper_chest"], ["front_delts", "triceps", "chest"]),
    "db bench": (["chest"], ["upper_chest", "front_delts", "triceps"]),
    "machine chest press": (["chest"], ["upper_chest", "front_delts", "triceps"]),
    "decline bench": (["chest"], ["triceps", "front_delts"]),
    "dips": (["chest", "triceps"], ["front_delts"]),
    "inc db fly": (["upper_chest"], ["chest", "front_delts"]),
    "db flys": (["chest"], ["upper_chest", "front_delts"]),
    "pec dec": (["chest"], ["upper_chest", "front_delts"]),
    "ez-bar skull crusher": (["triceps"], []),
    "straight-bar tricep pushdown": (["triceps"], []),
    "overhead rope extension": (["triceps"], []),
    "tricep extension": (["triceps"], []),
    "cable lat raise": (["side_delts"], ["front_delts", "traps"]),
    "lateral raise": (["side_delts"], ["front_delts", "traps"]),
    "lean away lat raise": (["side_delts"], ["front_delts", "traps"]),
    # pull
    "neutral-grip lat pulldown": (["lats"], ["upper_back", "biceps", "brachialis", "rear_delts"]),
    "s-a lat pulldown": (["lats"], ["upper_back", "biceps", "brachialis"]),
    "chest supported t-bar row": (["upper_back", "lats"], ["rear_delts", "biceps", "brachialis"]),
    "s-a machine row": (["lats", "upper_back"], ["rear_delts", "biceps", "brachialis"]),
    "seated cable row": (["upper_back", "lats"], ["rear_delts", "biceps", "lower_back"]),
    "cable lat pullover": (["lats"], ["triceps", "chest"]),
    "face pulls": (["rear_delts"], ["upper_back", "traps", "side_delts"]),
    "rear delt fly": (["rear_delts"], ["upper_back"]),
    "dumbbell er": (["rear_delts"], ["upper_back"]),          # rotator cuff sits under the rear delt
    "shrugs": (["traps"], ["upper_back", "forearm_flex"]),
    # arms
    "ez-bar cable curl": (["biceps"], ["brachialis", "forearm_flex"]),
    "single arm cable curls": (["biceps"], ["brachialis", "forearm_flex"]),
    "s-a machine preacher curl": (["biceps"], ["brachialis"]),
    "machine preacher curl": (["biceps"], ["brachialis"]),
    "db spider curl": (["biceps"], ["brachialis"]),
    "hammer curl": (["brachialis", "forearm_ext"], ["biceps"]),
    "reverse curl": (["forearm_ext", "brachialis"], ["biceps"]),
    "cable forearm curl": (["forearm_flex"], []),
    "wrist curl": (["forearm_flex"], []),
    # legs
    "hack squat": (["quads"], ["glutes", "adductors"]),
    "bulgarian split": (["quads", "glutes"], ["adductors", "hamstrings"]),
    "db rdl": (["hamstrings", "glutes"], ["lower_back", "adductors", "forearm_flex"]),
    "hip thrust": (["glutes"], ["hamstrings", "quads", "adductors"]),
    "leg extension": (["quads"], []),
    "hamstring curl": (["hamstrings"], ["calves"]),
    # core (retired)
    "crunch machine": (["abs"], ["obliques"]),
    "cable chops": (["obliques"], ["abs"]),
    "cable side rotation": (["obliques"], ["abs"]),
    "side cable rotation": (["obliques"], ["abs"]),
    "russian twists": (["obliques"], ["abs"]),
}
# Week of Nov 2 was logged on the new tabs under the old routine's headers.
# Each of those sessions is moved to the day it belongs to on the new split,
# and counts there for calls and volume. None = keep it off every day page
# (retired lifts). Keys are (date, lift); days match the start of a tab name.
REASSIGN = {
    ("2026-11-02", "ez-bar cable curl"): "Tue", ("2026-11-02", "hammer curl"): "Tue",
    ("2026-11-02", "s-a machine preacher curl"): "Tue",
    ("2026-11-02", "straight-bar tricep pushdown"): "Mon", ("2026-11-02", "ez-bar skull crusher"): "Mon",
    ("2026-11-02", "overhead rope extension"): "Mon",
    # Wed Legs didn't change, so 11/3 stays as a normal Wednesday session (a real baseline).
    ("2026-11-04", "inc db bench"): "Thurs", ("2026-11-04", "db bench"): "Thurs",
    ("2026-11-04", "machine chest press"): "Thurs", ("2026-11-04", "inc db fly"): "Thurs",
    ("2026-11-04", "straight-bar tricep pushdown"): "Thurs", ("2026-11-04", "machine shoulder press"): "Mon",
    ("2026-11-05", "neutral-grip lat pulldown"): "Fri", ("2026-11-05", "chest supported t-bar row"): "Tue",
    ("2026-11-05", "cable lat pullover"): "Fri", ("2026-11-05", "db spider curl"): "Fri",
    ("2026-11-05", "machine preacher curl"): None, ("2026-11-05", "seated cable row"): None,
    ("2026-11-06", "machine shoulder press"): "Mon", ("2026-11-06", "cable lat raise"): "Mon",
    ("2026-11-06", "rear delt fly"): "Fri", ("2026-11-06", "face pulls"): "Tue",
    ("2026-11-06", "shrugs"): "Fri", ("2026-11-06", "reverse curl"): "Tue",
    ("2026-11-06", "cable forearm curl"): "Tue",
}
# Machines whose empty sled/arm weighs something. You log the plates you add;
# the math (est. max, jump size, move-up prediction, pounds moved) adds this on.
BASE_WEIGHT = {"hack squat": 105, "hip thrust": 50}   # both confirmed at the gym 11/11/2026
# Cell fills: cyan = death set, either blue = different machine (skip it).
DEATH_FILL = "FF00FFFF"
SKEWED_FILLS = {"FF3C78D8", "FF4A86E8"}

# Same movement logged under different names.
ALIASES = {
    "curl machine": "machine preacher curl",
    "straight bar cable curl": "ez-bar cable curl",
    "seated hammer curl": "hammer curl",
    "tricep pushdown": "straight-bar tricep pushdown",
    "skull crushers": "ez-bar skull crusher",
    "oh extensions": "overhead rope extension",
    "over-head tricep extension w rope": "overhead rope extension",
    "wrist curls": "wrist curl",
    "reverse grip bb curl": "reverse curl", "rev bb curl": "reverse curl",
    "rev grip ez-bar curl": "reverse curl",
    "spider curl": "db spider curl",
    "lat pulldown": "neutral-grip lat pulldown",
    "chest supported row": "chest supported t-bar row",
    "row": "seated cable row", "cable row": "seated cable row",
    "lat pullover": "cable lat pullover", "cable pullover": "cable lat pullover",
    "chest press": "machine chest press",
    "shoulder press": "machine shoulder press",
    "seated machine shoulder press": "machine shoulder press",
    "leg ext": "leg extension", "hammy curl": "hamstring curl",
    "hip thrusts": "hip thrust",
    "lean shrug": "shrugs", "neutral grip cable shrugs": "shrugs",
    "tricep dip": "dips",
    "cable lateral raise": "cable lat raise",
    "s-a machine preacher curls": "s-a machine preacher curl",
    "ez bar cable curl": "ez-bar cable curl", "cable forearm curls": "cable forearm curl",
    "hamstring curls": "hamstring curl",
}
# Tab where Trent logs each death set's drops week to week (not a workout day).
TRACKER_TAB = re.compile(r"death set tracker", re.I)
# -----------------------------------------------------------------------------

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def read_workbook(path):
    """Yield (tab, row) where row maps column letter -> (formula, value, fill)."""
    with zipfile.ZipFile(path) as z:
        def xml(name):
            return ET.fromstring(z.read(name))
        strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            strings = ["".join(si.itertext()) for si in xml("xl/sharedStrings.xml").findall("m:si", NS)]
        styles = xml("xl/styles.xml")
        fills = []
        for f in styles.find("m:fills", NS):
            fg = f.find("m:patternFill/m:fgColor", NS)
            fills.append(fg.get("rgb") if fg is not None else None)
        xf_fill = [int(x.get("fillId", 0)) for x in styles.find("m:cellXfs", NS)]
        rels = {r.get("Id"): r.get("Target") for r in xml("xl/_rels/workbook.xml.rels")}
        for sheet in xml("xl/workbook.xml").find("m:sheets", NS):
            target = rels[sheet.get(REL)].lstrip("/")
            target = target if target.startswith("xl/") else "xl/" + target
            rows = {}
            for c in xml(target).iter("{%s}c" % NS["m"]):
                f, v = c.find("m:f", NS), c.find("m:v", NS)
                if f is None and v is None:
                    continue
                val = v.text if v is not None else ""
                if c.get("t") == "s":
                    val = strings[int(val)]
                s = int(c.get("s", 0))
                fill = fills[xf_fill[s]] if s < len(xf_fill) else None
                col, num = re.match(r"([A-Z]+)(\d+)", c.get("r")).groups()
                rows.setdefault(int(num), {})[col] = (f.text if f is not None else None, val, fill)
            for num in sorted(rows):
                yield sheet.get("name"), rows[num]


def lift_key(header):
    name = re.sub(r"\s*\d+\s*xf.*$", "", header, flags=re.I).strip().lower()
    return ALIASES.get(name, name)


def parse_sets(formula):
    sets = []
    for w, reps in re.findall(r"([\d.]+)\*\(([\d+\s]+)\)", formula.replace(" ", "")):
        sets += [(float(w), int(r)) for r in reps.split("+") if r]
    return sets


def is_header(row):
    b = row.get("B")
    return bool(b and b[0] is None and not re.fullmatch(r"[\d.]+", b[1] or ""))


def load(path, source):
    """Sessions from one workbook. A tab can hold several header rows (each new
    routine starts a new one); only rows under a tab's last header are that
    day's current routine."""
    sessions = []
    header, block = {}, {}
    for tab, row in read_workbook(path):
        if TRACKER_TAB.search(tab):
            continue
        if is_header(row):
            header = {col: cell[1] for col, cell in row.items() if col != "A"}
            block[tab] = block.get(tab, 0) + 1
            continue
        try:
            date = dt.date(1899, 12, 30) + dt.timedelta(days=int(float(row["A"][1])))
        except (KeyError, ValueError, TypeError):
            continue
        for col, (formula, _, fill) in row.items():
            if not formula or col not in header:
                continue
            sets = parse_sets(formula)
            if not sets:
                continue
            title = header[col]
            sessions.append(dict(
                lift=lift_key(title), title=re.sub(r"\s*\d+\s*xf.*$", "", title, flags=re.I).strip(),
                scheme=(re.search(r"\d+\s*xf.*$", title, re.I) or [""])[0],
                backoff=[int(n) - 1 for n in re.findall(r"S(\d+)\s*\d+\s*%", title, re.I)],
                tab=tab, block=block[tab], date=date.isoformat(), sets=sets, formula=formula, source=source,
                death=fill == DEATH_FILL or bool(re.search(r"death|\bS\d+\s*d\b", title, re.I)),
                skewed=fill in SKEWED_FILLS))
    for s in sessions:
        s["current_block"] = s["block"] == block[s["tab"]]
        s["day"] = s["tab"] if s["current_block"] else "Old routine"
        s["last_block"] = block[s["tab"]]
    return sessions


def reassign(sessions):
    """Move sessions listed in REASSIGN onto their day in the new split."""
    last_block = {s["tab"]: s["last_block"] for s in sessions if s["source"] == "current"}
    for s in sessions:
        if s["source"] != "current" or (s["date"], s["lift"]) not in REASSIGN:
            continue
        target = REASSIGN[(s["date"], s["lift"])]
        if target is None:
            s["off_day"] = True
            continue
        tab = next(t for t in last_block if t.split()[0] == target)
        s.update(tab=tab, day=tab, block=last_block[tab], current_block=True, moved=True)


def e1rm(w, r):
    return w * (1 + r / 30)


def best_e1rm(lift, sets):
    """Best est. max of a session, in the weight you load (base weight taken back off)."""
    b = BASE_WEIGHT.get(lift, 0)
    return max(e1rm(w + b, r) for w, r in sets) - b


def increment(lift, w):
    """Smallest sensible jump: about 2.5-5% where add-ons allow it."""
    if "db" in lift.split():
        return 2.5 if w < 30 else 5
    return 5 if w + BASE_WEIGHT.get(lift, 0) >= 100 else 2.5   # smallest add-on is 2.5 lbs


def predicted_avg(avg, w, inc):
    return avg - (30 + avg) * inc / w


def needed_avg(floor, w, inc):
    f = inc / w
    return (floor + 30 * f) / (1 - f)


def since_best(lift, usable):
    """Sessions logged after the one with the best est. max."""
    scores = [best_e1rm(lift, s["sets"]) for s in usable]
    return len(scores) - 1 - scores.index(max(scores))


def avg_reps(s):
    if s["backoff"]:
        # "S4 70%" lifts: judge only the max sets before the back-off. A max set
        # that dropped weight still counts, as the reps it equals at the top weight.
        working = s["sets"][:min(s["backoff"])] or s["sets"]
        top = max(w for w, _ in working)
        reps = [r if w == top else max(0.0, 30 * (e1rm(w, r) / top - 1)) for w, r in working]
        return sum(reps) / len(reps)
    top = max(w for w, _ in s["sets"])
    reps = [r for w, r in s["sets"] if w == top]
    return sum(reps) / len(reps)


def assess(lift, history, day=None):
    """Decide what to do next time for one lift on one workout day.

    The move-up call looks only at that day's sessions, so a lift done on two
    days is judged separately on each. Stalls and "too heavy" use the lift's
    whole history.
    """
    reset = FORM_RESET.get(lift)
    usable = [s for s in history if not s["skewed"] and (not reset or s["date"] >= reset)]
    floor = FLOOR_OVERRIDES.get(lift, FLOOR)
    if not usable:
        return dict(status="new", floor=floor)
    for s in usable:
        top = max(w for w, _ in s["sets"])
        s["top"] = top
        s["top_reps"] = max(r for w, r in s["sets"] if w == top)
        s["avg"] = avg_reps(s)
    on_day = [s for s in usable if s["tab"] == day and s["current_block"]]
    judged = on_day or usable
    last = judged[-1]
    weight = last["top"]
    at_weight = []
    for s in reversed(judged):
        if s["top"] != weight:
            break
        at_weight.insert(0, s)
    prev_weight = next((s["top"] for s in reversed(judged[:-len(at_weight)]) if s["top"] != weight), None)
    avg = round(last["avg"], 1)
    inc = increment(lift, weight)
    moved = weight + BASE_WEIGHT.get(lift, 0)      # what you actually move
    pred = predicted_avg(last["avg"], moved, inc)
    out = dict(floor=floor, weight=weight, reps=last["top_reps"], avg=avg,
               predicted=round(pred, 1), target=round(min(REP_CAP, needed_avg(floor, moved, inc)), 1),
               capped=needed_avg(floor, moved, inc) > REP_CAP,
               jump_pct=round(100 * inc / moved, 1),
               sessions_at_weight=len(at_weight), prev_weight=prev_weight, judged_date=last["date"],
               other_day=None if on_day or day is None else last["day"],
               next_weight=weight + inc)
    # Too heavy looks at every session at this weight since you moved up to it,
    # on any day, old split included: a weight that's too heavy is too heavy anywhere.
    jumped = next((i for i in range(len(usable) - 1, -1, -1) if usable[i]["top"] < weight), None)
    since_jump = [s for s in usable[jumped + 1:] if s["top"] == weight] if jumped is not None else []
    too_heavy = (usable[jumped]["top"] if len(since_jump) >= 2
                 and max(s["avg"] for s in since_jump) < TOO_HEAVY_AVG else None)
    if pred >= floor or last["avg"] >= REP_CAP:
        out["status"] = "move"
    elif too_heavy:
        out.update(status="drop", prev_weight=too_heavy)
    elif since_best(lift, usable) >= HOLD_SESSIONS:
        best = usable[len(usable) - 1 - since_best(lift, usable)]
        out.update(status="holding", since_best=since_best(lift, usable), best_weight=best["top"],
                   best_reps=best["top_reps"], best_date=best["date"])
    elif pred >= floor - CLOSE_WITHIN or last["avg"] >= REP_CAP - CLOSE_WITHIN:
        out["status"] = "close"
    else:
        out["status"] = "build"
    return out


def weekly_drops(sessions):
    """Per week (Monday date), lifts whose est. max fell vs. the last session on the same day."""
    prev, weeks = {}, {}
    for s in sessions:
        reset = FORM_RESET.get(s["lift"])
        if s["skewed"] or (reset and s["date"] < reset):
            continue
        b = BASE_WEIGHT.get(s["lift"], 0)
        score = best_e1rm(s["lift"], s["sets"]) + b      # compare full loads
        k = (s["lift"], s["tab"], s["block"])   # same day of the same routine
        if k in prev and score <= prev[k] * (1 - FATIGUE_PCT / 100):
            d = dt.date.fromisoformat(s["date"])
            week = (d - dt.timedelta(days=d.weekday())).isoformat()
            weeks.setdefault(week, []).append(dict(
                name=s["title"], day=s["day"], from_=round(prev[k] - b), to=round(score - b),
                pct=round(100 * (1 - score / prev[k]))))
        prev[k] = score
    return weeks


def weekly_volume(sessions, current_tabs):
    """Sets and pounds moved per week (Mon-Sun), by muscle and by workout tab.
    Every logged set counts, different-machine ones too: the work was done."""
    muscles, workouts, by_day = {}, {}, {}
    for s in sessions:
        d = dt.date.fromisoformat(s["date"])
        week = (d - dt.timedelta(days=d.weekday())).isoformat()
        b = BASE_WEIGHT.get(s["lift"], 0)
        n, lbs = len(s["sets"]), sum((w + b) * r for w, r in s["sets"])
        primary, secondary = MUSCLES.get(s["lift"], ([], []))
        for m in primary + secondary:     # [direct sets, assist sets, direct lbs, assist lbs]
            x = muscles.setdefault(week, {}).setdefault(m, [0, 0, 0, 0])
            x[0 if m in primary else 1] += n
            x[2 if m in primary else 3] += round(lbs)
        if s["tab"] in current_tabs and not s.get("off_day"):     # per lift, kept apart by day so Mon and Thu never mix
            y = by_day.setdefault(s["tab"], {}).setdefault(s["lift"], dict(title=s["title"], weeks={}))
            y["title"] = s["title"]
            z = y["weeks"].setdefault(week, [0, 0])
            z[0] += n; z[1] += round(lbs)
        x = workouts.setdefault(week, {}).setdefault(s["tab"], dict(sets=0, lbs=0, dates=set()))
        x["sets"] += n; x["lbs"] += lbs; x["dates"].add(s["date"])
    for week in workouts.values():
        for x in week.values():
            x["lbs"] = round(x["lbs"]); x["sessions"] = len(x.pop("dates"))
    tabs = list(current_tabs) + sorted({t for w in workouts.values() for t in w} - set(current_tabs))
    return dict(weeks=sorted(workouts, reverse=True), muscles=muscles, workouts=workouts, tabs=tabs, by_day=by_day)


def load_death(path, day_tabs):
    """Death set drops from the tracker tab. Header: "Week" then "Day - Lift";
    each row is one week, labelled like "1 (11/9-11/13)"; each cell is the drop
    chain as weight*(reps)+weight*(reps)..."""
    columns, weeks, header = [], {}, None
    today = dt.date.today()
    for tab, row in read_workbook(path):
        if not TRACKER_TAB.search(tab):
            continue
        if row.get("A", (None, ""))[1].strip().lower() == "week":
            header = {}
            for col in sorted(row, key=lambda c: (len(c), c)):
                text = row[col][1]
                if col == "A" or not text or "-" not in text:
                    continue
                day, lift = (x.strip() for x in text.split("-", 1))
                tab_name = next((t for t in day_tabs if t.split()[0].lower() == day.lower()), None)
                header[col] = dict(id=col, day=tab_name, key=lift_key(lift), name=lift)
                columns.append(header[col])
            continue
        m = re.search(r"(\d+)/(\d+)", row.get("A", (None, ""))[1] or "")
        if not header or not m:
            continue
        start = dt.date(today.year, int(m[1]), int(m[2]))
        if start > today + dt.timedelta(days=60):
            start = start.replace(year=today.year - 1)
        week = (start - dt.timedelta(days=start.weekday())).isoformat()
        for col, (formula, _, _) in row.items():
            drops = parse_sets(formula) if formula and col in header else None
            if drops:
                weeks.setdefault(week, {})[col] = dict(drops=drops, lbs=round(sum(w * r for w, r in drops)),
                                                       reps=sum(r for _, r in drops))
    return dict(columns=columns, weeks=weeks)


def load_times(path, sessions):
    """Workout length from each day tab's "Time" column (a duration cell, stored as
    a fraction of a day), with that date's sets and pounds for pace."""
    out, header = [], {}
    for tab, row in read_workbook(path):
        if TRACKER_TAB.search(tab):
            continue
        if is_header(row):
            header = {c: (v[1] or "").strip().lower() for c, v in row.items()}
            continue
        col = next((c for c, h in header.items() if h == "time"), None)
        try:
            date = dt.date(1899, 12, 30) + dt.timedelta(days=int(float(row["A"][1])))
            frac = float(row[col][1])
        except (KeyError, ValueError, TypeError):
            continue
        if frac <= 0:
            continue
        mine = [x for x in sessions if x["source"] == "current" and x["date"] == date.isoformat()]
        out.append(dict(date=date.isoformat(), day=tab, minutes=round(frac * 1440, 1),
                        sets=sum(len(x["sets"]) for x in mine),
                        lbs=round(sum((w + BASE_WEIGHT.get(x["lift"], 0)) * r for x in mine for w, r in x["sets"]))))
    return sorted(out, key=lambda t: t["date"])


def build():
    sessions = load(HISTORY, "history") + load(CURRENT, "current")
    reassign(sessions)
    sessions.sort(key=lambda s: s["date"])
    by_lift = {}
    for s in sessions:
        by_lift.setdefault(s["lift"], []).append(s)

    # The current routine is the last header row on each tab of the newest sheet.
    last_header = {}
    for tab, row in read_workbook(CURRENT):
        if is_header(row) and not TRACKER_TAB.search(tab):
            last_header[tab] = row
    current_order, tab_of, tabs, header_text, col_pos = [], {}, list(last_header), {}, {}
    for tab, row in last_header.items():
        for col in sorted(row, key=lambda c: (len(c), c)):
            text = row[col][1]
            if col == "A" or not text or text.strip().lower() in ("time", "date"):
                continue
            k = lift_key(text)
            header_text[(k, tab)] = text
            col_pos[(k, tab)] = len(col_pos)
            if tab not in tab_of.setdefault(k, []):
                tab_of[k].append(tab)
            if k not in current_order:
                current_order.append(k)

    def split_header(text):
        return (re.sub(r"\s*\d+\s*xf.*$", "", text, flags=re.I).strip(),
                (re.search(r"\d+\s*xf.*$", text, re.I) or [""])[0])

    lifts = []
    for k in list(by_lift) + [k for k in tab_of if k not in by_lift]:   # new lifts have no history yet
        hist = by_lift.get(k, [])
        history = [dict(date=s["date"], day=s["day"], sets=s["sets"], death=s["death"], skewed=s["skewed"],
                         moved=s.get("moved", False),
                         avg=round(avg_reps(s), 1), e1rm=round(best_e1rm(k, s["sets"]), 1))
                    for s in hist]
        for day in tab_of.get(k) or [None]:
            a = assess(k, hist, day)
            name, scheme = split_header(header_text[(k, day)]) if day else (hist[-1]["title"], hist[-1]["scheme"])
            lifts.append(dict(
                id=f"{k}|{day or ''}", key=k, name=name, scheme=scheme, day=day,
                current=day is not None, order=col_pos.get((k, day), 999),
                form_reset=FORM_RESET.get(k), base=BASE_WEIGHT.get(k, 0), primary=MUSCLES.get(k, ([], []))[0],
                secondary=MUSCLES.get(k, ([], []))[1], **a, sessions=history))
    lifts.sort(key=lambda l: (not l["current"], l["order"], l["key"]))
    unmapped = sorted({l["key"] for l in lifts if l["key"] not in MUSCLES})
    if unmapped:
        print("No muscle map for:", ", ".join(unmapped), "(add them to MUSCLES)")
    weeks = weekly_drops(sessions)
    last = dt.date.fromisoformat(sessions[-1]["date"])
    this_week = (last - dt.timedelta(days=last.weekday())).isoformat()
    drops = [{**d, "from": d.pop("from_")} for d in weeks.get(this_week, [])]
    data = dict(fatigue=dict(week=this_week, drops=drops) if len(drops) >= FATIGUE_LIFTS else None,built=dt.datetime.now().strftime("%a %b %-d, %-I:%M %p"),
                first=sessions[0]["date"], last=sessions[-1]["date"],
                logged=len({(s["date"], s["tab"]) for s in sessions}), days=tabs,
                rules=dict(floor=FLOOR, overrides=FLOOR_OVERRIDES, cap=REP_CAP, close=CLOSE_WITHIN,
                           hold=HOLD_SESSIONS, too_heavy=TOO_HEAVY_AVG,
                           fatigue_lifts=FATIGUE_LIFTS, fatigue_pct=FATIGUE_PCT),
                lifts=lifts, volume=weekly_volume(sessions, tabs), death=load_death(CURRENT, tabs),
                times=load_times(CURRENT, sessions))
    page = (HERE / "template.html").read_text()
    OUT.write_text(page.replace("/*__DATA__*/null", json.dumps(data)))
    ready = [f'{l["name"]} ({l["day"]})' for l in lifts if l["current"] and l["status"] == "move"]
    print(f"Wrote {OUT} — {len(lifts)} lifts, ready to move up: {', '.join(ready) or 'none'}")


if __name__ == "__main__":
    build()
