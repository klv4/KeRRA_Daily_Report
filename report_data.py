"""
report_data.py
==============
Defaults, chainage maths and the save / carry-forward logic for the KeRRA daily report app.
No Streamlit code in here, so it can be tested on its own.
"""

import datetime
import json
import math
import re

import pandas as pd

# ------------------------------------------------------------------ choices
WEATHER_OPTIONS = ["Sunny", "Partly cloudy", "Cloudy", "Overcast", "Light rain", "Rain", "Heavy rain", "Windy", "Foggy"]
GROUND_OPTIONS = ["Good", "Bad", "Wet", "Dry"]                      # fixed list
STATUS_OPTIONS = ["Working", "Not working"]                         # fixed list
SIDE_OPTIONS = ["RHS", "LHS"]                                       # fixed list
RAIN_LOCATIONS = ["Wahambla site (:0+854) km", "Majiwa sec. school(9+000)km"]
ACTIVITY_OPTIONS = ["Embankment fill", "Removal of existing culvert", "Site clearance", "Excavation of culverts",
                    "Earthworks; cut to fill", "Concrete casting", "Bed casting & BRC", "Rock fill",
                    "Casting of access culvert", "TSG", "BSG", "Base cement stabilization", "Cut to spoil"]
UNIT_OPTIONS = ["pcs", "m3", "kg", "m", "trucks"]
TIME_OPTIONS = [f"{h % 12 or 12}:00 {'am' if h < 12 else 'pm'}" for h in range(24)]
CHAINAGE_STEP = 20                                                  # metres between chainages in the dropdown

# the lists that grow when a new value is typed (saved in lists.json next to the app)
GROWING_LISTS = {"activity": ACTIVITY_OPTIONS, "unit": UNIT_OPTIONS, "time": TIME_OPTIONS,
                 "weather": WEATHER_OPTIONS, "rainloc": RAIN_LOCATIONS, "extra_chainage": []}

# ------------------------------------------------------------------ table columns
ACT_COLS = ["Operation", "Chainage from", "Chainage to", "Side", "Previous (m)", "Today (m)", "Remark / status"]
EQ_COLS = ["Equipment / machinery", "Qty", "Reg. no.", "From", "To", "Idle hours", "Status"]
LAB_COLS = ["Category", "Number", "From", "To", "Idle hours"]
MAT_COLS = ["Material", "Quantity", "Unit"]
VEH_COLS = ["Vehicle reg. no.", "Assigned", "Insurance status", "Daily condition", "Remarks"]
STAFF_COLS = ["Team", "Activity"]
RAIN_COLS = ["Location", "Rainfall (mm)"]

TABLE_COLS = {"activities": ACT_COLS, "equipment": EQ_COLS, "labour": LAB_COLS, "materials": MAT_COLS,
              "vehicles": VEH_COLS, "staff": STAFF_COLS, "rainfall": RAIN_COLS}
NUMERIC_COLS = {"Previous (m)", "Today (m)", "Qty", "Idle hours", "Number", "Rainfall (mm)", "Quantity"}

START, END = "8:00 am", "5:00 pm"
NO_DELAYS = "There were no delays."


def fmt_chainage(m):
    """2260 -> '2+260'."""
    m = int(round(m))
    return f"{m // 1000}+{m % 1000:03d}"


def chainage_options(length_m, extra=()):
    """0+000, 0+020, ... up to the road length (the end is added if it is not a multiple of the step),
    plus any chainages typed by hand, in order."""
    vals = set()
    if length_m and length_m > 0:
        n = int(length_m) // CHAINAGE_STEP
        vals.update(i * CHAINAGE_STEP for i in range(n + 1))
        vals.add(int(length_m))
    out = {fmt_chainage(v): v for v in vals}
    for e in extra:
        m = chainage_m(e)
        if m is not None and e not in out:
            out[e] = m
    return [k for k, _ in sorted(out.items(), key=lambda kv: (kv[1], kv[0]))]


def caption_text(op, frm, to, side):
    """'Subbase stabilization 2+260 to 2+660 RHS' - the operation merged with its chainages."""
    ch = f"{frm} to {to}" if (frm and to) else (frm or to or "")
    return " ".join(x for x in (op, ch, side) if x).strip()


def empty_df(name, rows=1):
    cols = TABLE_COLS[name]
    data = {c: [None if c in NUMERIC_COLS else ""] * rows for c in cols}
    df = pd.DataFrame(data, columns=cols)
    for c in cols:
        df[c] = df[c].astype("float64") if c in NUMERIC_COLS else df[c].astype("object")
    return df


def _df(name, rows):
    cols = TABLE_COLS[name]
    df = pd.DataFrame(rows, columns=cols)
    for c in cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").astype("float64") if c in NUMERIC_COLS else df[c].fillna("").astype("object")
    return df


def default_equipment():
    rows = []
    for name, reg in [("Double cabin", "KBS 072 K"), ("Double cabin", "KBV 351 R"), ("Double cabin", "KDZ 972 A"),
                      ("Single cabin", "KBQ 912 R"), ("Excavator", ""), ("Grader", ""), ("Grader", ""),
                      ("Tipper", "KDZ 090 F"), ("Tipper", "KDZ 225 F"), ("Tipper", "KDJ 228 F"),
                      ("Tipper", "KDZ 230 F"), ("Tipper", "KDZ 236 F"), ("Tipper", "KDJ 236 H"),
                      ("Tipper", "KBS 093 L"), ("Wheel shovel", ""), ("Flat roller", ""), ("Sheep foot", ""),
                      ("Water bowser", ""), ("Back hoe", ""), ("Back hoe", "")]:
        rows.append([name, 1, reg, START, END, 0, "Working"])
    return _df("equipment", rows)


def default_labour():
    rows = [("Site engineer", 1), ("Site manager", 1), ("Surveyor", 1), ("Foreman", 3), ("Operator", 8),
            ("Driver", 13), ("Survey team", 4), ("Material team", 4), ("General labourers", 46), ("Intern", 3),
            ("Security personnel", 1)]
    return _df("labour", [[c, n, START, END, 0] for c, n in rows])


def default_vehicles():
    return _df("vehicles", [["KDV 984 M", "Field", "Updated", "Good", "Working"], ["", "", "", "", ""], ["", "", "", "", ""]])


def default_staff():
    return _df("staff", [["Survey", ""], ["Inspectorate", ""], ["Materials", ""]])


def default_state():
    """Everything the form holds, as plain values (dates as date objects, tables as DataFrames)."""
    return {
        "project": "", "contract_no": "", "contractor": "",
        "date": datetime.date.today(), "road_length": None,
        "station_from": None, "station_to": None,
        "morning": "Sunny", "afternoon": "Sunny", "ground": "Good", "tmax": None, "tmin": None,
        "rainfall": _df("rainfall", [[loc, 0.0] for loc in RAIN_LOCATIONS]),
        "activities": empty_df("activities", 1),
        "equipment": default_equipment(),
        "visitors": "Nil",
        "labour": default_labour(),
        "materials": empty_df("materials", 0),
        "issues": NO_DELAYS,
        "vehicles": default_vehicles(),
        "staff": default_staff(),
    }


# ------------------------------------------------------------------ number / chainage helpers
def is_blank(v):
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    try:
        if pd.isna(v):
            return True
    except (TypeError, ValueError):
        pass
    return str(v).strip() == ""


def fmt_num(v):
    """17 -> '17', 0.5 -> '0.5', blank -> ''."""
    if is_blank(v):
        return ""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    return str(int(round(f))) if abs(f - round(f)) < 1e-9 else f"{f:.2f}".rstrip("0").rstrip(".")


def to_float(v):
    if is_blank(v):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


_CH = re.compile(r"^\s*(\d+)\s*\+\s*(\d{1,3}(?:\.\d+)?)")


def chainage_m(text):
    """'2+260' -> 2260.0 (metres). Trailing words such as '2+370 RHS' are ignored. Not a chainage -> None."""
    m = _CH.match(str(text or ""))
    return int(m.group(1)) * 1000 + float(m.group(2)) if m else None


def activity_rows(df):
    """Rows of the activities table with the TOTAL worked out.
    If 'Today (m)' is left blank but both chainages are valid, today's length is taken from them."""
    out = []
    for _, r in df.iterrows():
        op = "" if is_blank(r["Operation"]) else str(r["Operation"]).strip()
        frm = "" if is_blank(r["Chainage from"]) else str(r["Chainage from"]).strip()
        to = "" if is_blank(r["Chainage to"]) else str(r["Chainage to"]).strip()
        remark = "" if is_blank(r["Remark / status"]) else str(r["Remark / status"]).strip()
        side = "" if is_blank(r["Side"]) else str(r["Side"]).strip()
        if not (op or frm or to or remark or not is_blank(r["Previous (m)"]) or not is_blank(r["Today (m)"])):
            continue                                                  # completely empty row
        prev, today = to_float(r["Previous (m)"]), to_float(r["Today (m)"])
        auto = False
        if today is None:
            a, b = chainage_m(frm), chainage_m(to)
            if a is not None and b is not None and b != a:
                today, auto = abs(b - a), True
        total = None
        if prev is not None or today is not None:
            total = (prev or 0) + (today or 0)
        out.append({"op": op, "from": frm, "to": to, "side": side, "prev": prev, "today": today, "total": total,
                    "remark": remark, "auto": auto,
                    "from_txt": f"{frm} {side}".strip() if frm else "",      # chainage + RHS / LHS in one cell
                    "to_txt": f"{to} {side}".strip() if to else "",
                    "caption": caption_text(op, frm, to, side)})
    return out


def clean_rows(df):
    """List of lists of display strings, completely empty rows dropped."""
    rows = []
    for _, r in df.iterrows():
        cells = ["" if is_blank(v) else (fmt_num(v) if isinstance(v, (int, float)) else str(v).strip()) for v in r]
        if any(cells):
            rows.append(cells)
    return rows


# ------------------------------------------------------------------ save / open / carry forward
def to_json(state):
    d = {}
    for k, v in state.items():
        if isinstance(v, pd.DataFrame):
            d[k] = json.loads(v.to_json(orient="records"))
        elif isinstance(v, datetime.date):
            d[k] = v.isoformat()
        else:
            d[k] = v
    d["_app"], d["_v"] = "kerra-daily-report", 2
    return json.dumps(d, indent=1)


def from_json(text):
    d = json.loads(text)
    if d.get("_app") != "kerra-daily-report":
        raise ValueError("This is not a saved KeRRA daily report file.")
    st = default_state()
    for k, base in st.items():
        if k not in d:
            continue
        if isinstance(base, pd.DataFrame):
            cols = TABLE_COLS[k]
            st[k] = _df(k, [[row.get(c) for c in cols] for row in d[k]]) if d[k] else empty_df(k, 0)
            if k == "rainfall" and not d[k]:
                st[k] = _df(k, [[loc, 0.0] for loc in RAIN_LOCATIONS])
        elif k == "date":
            st[k] = datetime.date.fromisoformat(d[k])
        else:
            st[k] = d[k]
    return st


def carry_forward(st):
    """Tomorrow's report: same project, plant, labour, vehicles and teams; previous length = old total;
    weather, rainfall, today's work, deliveries, delays and notes are cleared."""
    n = default_state()
    for k in ("project", "contract_no", "contractor", "road_length", "equipment", "labour", "vehicles"):
        n[k] = st[k]
    n["date"] = st["date"] + datetime.timedelta(days=1)
    rows = []
    for a in activity_rows(st["activities"]):
        if a["op"]:
            rows.append([a["op"], None, None, None, a["total"], None, ""])
    n["activities"] = _df("activities", rows) if rows else empty_df("activities", 1)
    n["staff"] = _df("staff", [[t, ""] for t in st["staff"]["Team"]]) if len(st["staff"]) else default_staff()
    return n


def totals(st):
    """Figures for the on-screen check."""
    eq = [r for r in clean_rows(st["equipment"])]
    qty = sum(to_float(r[1]) or 0 for r in eq)
    lab = sum(to_float(r[1]) or 0 for r in clean_rows(st["labour"]))
    return {"equipment_units": int(qty), "workers": int(lab),
            "idle_equipment": sum(1 for r in eq if (to_float(r[5]) or 0) > 0)}


# ------------------------------------------------------------------ dropdown lists that grow
def load_lists(path):
    """The growing dropdown lists: the built-in values plus anything saved earlier in lists.json."""
    lists = {k: list(v) for k, v in GROWING_LISTS.items()}
    try:
        with open(path, encoding="utf-8") as f:
            saved = json.load(f)
        for k in lists:
            for v in saved.get(k, []):
                if v not in lists[k]:
                    lists[k].append(v)
    except (OSError, ValueError):
        pass
    return lists


def save_lists(path, lists):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(lists, f, indent=1, ensure_ascii=False)
        return True
    except OSError:
        return False
