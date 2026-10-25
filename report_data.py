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
GROUND_OPTIONS = ["Good", "Fair", "Damp", "Wet", "Muddy", "Waterlogged"]
STATUS_OPTIONS = ["Working", "Idle", "Standby", "Breakdown", "Under repair", "Left site"]

# ------------------------------------------------------------------ table columns
ACT_COLS = ["Operation", "Chainage from", "Chainage to", "Previous (m)", "Today (m)", "Remark / status"]
EQ_COLS = ["Equipment / machinery", "Qty", "Reg. no.", "From", "To", "Idle hours", "Status"]
LAB_COLS = ["Category", "Number", "From", "To", "Idle hours"]
MAT_COLS = ["Material", "Total quantity delivered"]
VEH_COLS = ["Vehicle reg. no.", "Assigned", "Insurance status", "Daily condition", "Remarks"]
STAFF_COLS = ["Team", "Activity"]
RAIN_COLS = ["Location", "Rainfall (mm)"]

TABLE_COLS = {"activities": ACT_COLS, "equipment": EQ_COLS, "labour": LAB_COLS, "materials": MAT_COLS,
              "vehicles": VEH_COLS, "staff": STAFF_COLS, "rainfall": RAIN_COLS}
NUMERIC_COLS = {"Previous (m)", "Today (m)", "Qty", "Idle hours", "Number", "Rainfall (mm)"}

START, END = "8:00 am", "5:00 pm"
NO_DELAYS = "There were no delays."


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
    return _df("vehicles", [["", "", "", "", ""]] * 3)


def default_staff():
    return _df("staff", [["Survey", ""], ["Inspectorate", ""], ["Materials", ""]])


def default_state():
    """Everything the form holds, as plain values (dates as date objects, tables as DataFrames)."""
    return {
        "project": "", "contract_no": "", "contractor": "",
        "date": datetime.date.today(),
        "station_from": "", "station_to": "",
        "morning": "Sunny", "afternoon": "Sunny", "ground": "Good", "tmax": None, "tmin": None,
        "rainfall": empty_df("rainfall", 0),
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
        out.append({"op": op, "from": frm, "to": to, "prev": prev, "today": today, "total": total,
                    "remark": remark, "auto": auto})
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
    d["_app"], d["_v"] = "kerra-daily-report", 1
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
        elif k == "date":
            st[k] = datetime.date.fromisoformat(d[k])
        else:
            st[k] = d[k]
    return st


def carry_forward(st):
    """Tomorrow's report: same project, plant, labour, vehicles and teams; previous length = old total;
    weather, rainfall, today's work, deliveries, delays and notes are cleared."""
    n = default_state()
    for k in ("project", "contract_no", "contractor", "equipment", "labour", "vehicles"):
        n[k] = st[k]
    n["date"] = st["date"] + datetime.timedelta(days=1)
    rows = []
    for a in activity_rows(st["activities"]):
        if a["op"]:
            rows.append([a["op"], "", "", a["total"], None, ""])
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
