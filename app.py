"""
KeRRA — Daily Progress Report / Site Diary
-------------------------------------------
Tabs:  Home  |  Daily report

Run locally:
    pip install -r requirements.txt
    streamlit run app.py

Keep these files in the SAME folder:
    app.py  version.py  report_data.py  report_docx.py  logo.png
"""

import base64
import datetime
import os

import pandas as pd
import streamlit as st

st.set_page_config(page_title="KeRRA — Daily Report", page_icon="🛣️", layout="wide")

import report_data as rd              # noqa: E402
import report_docx as dx              # noqa: E402
from version import VERSION, RELEASE_DATE   # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
ss = st.session_state

SCALARS = ["project", "contract_no", "contractor", "date", "road_length", "station_from", "station_to", "morning",
           "afternoon", "ground", "tmax", "tmin", "visitors", "issues"]
EDITOR_TABLES = ["vehicles", "staff"]                 # plain-text tables (no dropdowns)
# Tables built from one row of widgets per line, so every dropdown can take a new typed value.
# Column kinds: text | int | num | dd:<growing list> | fixed:<options>
ROW_TABLES = {
    "activities": ["dd:activity", "dd:chainage", "dd:chainage", "fixed:side", "num", "num", "text"],
    "equipment": ["text", "int", "text", "dd:time", "dd:time", "num", "fixed:status"],
    "labour": ["text", "int", "dd:time", "dd:time", "num"],
    "materials": ["text", "num", "dd:unit"],
    "rainfall": ["dd:rainloc", "num"],
}
FIXED = {"side": rd.SIDE_OPTIONS, "status": rd.STATUS_OPTIONS}
LISTS_FILE = os.path.join(HERE, "lists.json")
if "lists" not in ss:
    ss["lists"] = rd.load_lists(LISTS_FILE)


def _key(table, rid, ci):
    return f"r|{table}|{rid}|{ci}"


def _num_or_none(v, integer=False):
    f = rd.to_float(v)
    return None if f is None else (int(round(f)) if integer else float(f))


def learn(listname, value):
    """Add a value typed into a dropdown to its list (and remember it in lists.json). True if it was new."""
    if value in (None, "") or value in ss["lists"][listname]:
        return False
    ss["lists"][listname].append(value)
    rd.save_lists(LISTS_FILE, ss["lists"])
    return True


def chainage_list():
    return rd.chainage_options(ss.get("road_length"), ss["lists"]["extra_chainage"])


def _store_row(table, rid, values):
    for ci, (kind, v) in enumerate(zip(ROW_TABLES[table], values)):
        k = _key(table, rid, ci)
        if kind in ("int", "num"):
            ss[k] = _num_or_none(v, kind == "int")
        elif kind == "text":
            ss[k] = "" if rd.is_blank(v) else str(v)
        else:                                          # dropdowns: blank -> nothing selected
            v = None if rd.is_blank(v) else str(v)
            ss[k] = v
            if v and kind.startswith("dd:"):
                if kind == "dd:chainage":
                    learn("extra_chainage", v) if v not in rd.chainage_options(ss.get("road_length")) else None
                else:
                    learn(kind[3:], v)


def _new_id():
    ss["nid"] = ss.get("nid", 0) + 1
    return ss["nid"]


NEW_ROW = {
    "activities": [None, None, None, None, None, None, ""],
    "equipment": ["", 1, "", rd.START, rd.END, 0.0, "Working"],
    "labour": ["", 1, rd.START, rd.END, 0.0],
    "materials": ["", None, None],
}


def add_row(table):
    rid = _new_id()
    ss["rows_" + table].append(rid)
    _store_row(table, rid, NEW_ROW[table])


def del_row(table, rid):
    ss["rows_" + table].remove(rid)
    for ci in range(len(ROW_TABLES[table])):
        ss.pop(_key(table, rid, ci), None)


def table_df(table):
    """The rows on screen as a DataFrame (the form the report builder and the save file use)."""
    cols = rd.TABLE_COLS[table]
    rows = []
    for rid in ss["rows_" + table]:
        rows.append([ss.get(_key(table, rid, ci)) for ci in range(len(cols))])
    return rd._df(table, rows) if rows else rd.empty_df(table, 0)


# ============================================================ state handling
def apply_state(state):
    """Put a whole report (from defaults, a saved file or a carry-forward) into the form."""
    for k in SCALARS:
        v = state[k]
        ss[k] = None if (k in ("station_from", "station_to") and rd.is_blank(v)) else v
        if k in ("road_length", "tmax", "tmin"):
            ss[k] = _num_or_none(v)
    for key in [k for k in ss.keys() if isinstance(k, str) and (k.startswith("r|") or k.startswith("ph|"))]:
        del ss[key]
    for table in ROW_TABLES:
        ss["rows_" + table] = []
        df = state[table]
        for _, r in df.iterrows():
            rid = _new_id()
            ss["rows_" + table].append(rid)
            _store_row(table, rid, list(r))
    for k in ("morning", "afternoon"):
        learn("weather", ss[k])
    for k in EDITOR_TABLES:
        ss["base_" + k] = state[k]
        ss.pop("out_" + k, None)
    ss["ver"] = ss.get("ver", 0) + 1          # new editor keys, so the plain tables reload
    ss.pop("docx_bytes", None)


def current_state():
    s = {k: ss[k] for k in SCALARS}
    for k in ("station_from", "station_to"):
        s[k] = s[k] or ""
    for k in ROW_TABLES:
        s[k] = table_df(k)
    for k in EDITOR_TABLES:
        s[k] = ss.get("out_" + k, ss["base_" + k])
    return s


if "ver" not in ss:
    apply_state(rd.default_state())


def _open_saved():
    f = ss.get("up_json")
    if f is None:
        ss["_msg"] = ("warning", "Choose a saved report file first.")
        return
    try:
        state = rd.from_json(f.getvalue().decode("utf-8"))
    except Exception as e:                                        # noqa: BLE001
        ss["_msg"] = ("error", f"That file could not be opened: {e}")
        return
    if ss.get("open_mode", "").startswith("Start the next"):
        state = rd.carry_forward(state)
        ss["_msg"] = ("success", f"Next day's report started for {state['date'].strftime('%d/%m/%Y')}: plant, labour and "
                                 "vehicles kept; previous lengths updated; today's work, weather and photos cleared.")
    else:
        ss["_msg"] = ("success", f"Report of {state['date'].strftime('%d/%m/%Y')} reopened.")
    apply_state(state)


def _blank():
    apply_state(rd.default_state())
    ss["_msg"] = ("success", "Blank report started.")


# ============================================================ styling
def _logo_b64():
    try:
        with open(os.path.join(HERE, "logo.png"), "rb") as f:
            return base64.b64encode(f.read()).decode("ascii")
    except Exception:                                             # noqa: BLE001
        return ""


CSS = """
<style>
.hero{background:linear-gradient(120deg,#0b4d2c 0%,#0f7a43 55%,#17a05a 100%);border-radius:18px;padding:30px 32px;
      display:flex;align-items:center;justify-content:space-between;gap:24px;margin:4px 0 22px 0;
      box-shadow:0 8px 24px rgba(11,77,44,.25);}
.hero .eyebrow{letter-spacing:.18em;font-size:.72rem;font-weight:700;color:#b8ebcf;margin-bottom:6px;}
.hero h1{color:#fff;font-size:1.85rem;line-height:1.2;margin:0 0 10px 0;padding:0;font-weight:700;}
.hero p{color:#dcf5e6;margin:0;font-size:.98rem;line-height:1.5;}
.hero .logo{background:#fff;border-radius:14px;padding:12px 14px;flex:0 0 auto;box-shadow:0 4px 14px rgba(0,0,0,.18);}
.hero .logo img{height:84px;display:block;}
.card{background:#fff;border:1px solid #d3e8dc;border-radius:16px;padding:20px 22px 14px 22px;height:100%;
      box-shadow:0 2px 10px rgba(15,122,67,.08);}
.card .ico{width:44px;height:44px;border-radius:12px;background:#e6f5ec;display:flex;align-items:center;
           justify-content:center;font-size:1.4rem;margin-bottom:10px;}
.card h3{color:#0b4d2c;margin:0 0 6px 0;padding:0;font-size:1.15rem;}
.card p{color:#4a5b6d;font-size:.9rem;margin:0 0 8px 0;}
.card ul{color:#33465a;font-size:.86rem;margin:0 0 6px 18px;padding:0;line-height:1.55;}
.steps{display:flex;gap:12px;margin:22px 0 8px 0;flex-wrap:wrap;}
.step{flex:1 1 150px;background:#f1f9f4;border:1px solid #d3e8dc;border-radius:12px;padding:12px 14px;}
.step b{color:#0f7a43;display:block;font-size:.78rem;letter-spacing:.08em;margin-bottom:3px;}
.step span{color:#33465a;font-size:.86rem;}
.foot{margin-top:26px;padding-top:14px;border-top:1px solid #d3e8dc;color:#6b7c8e;font-size:.8rem;text-align:center;}
@media (max-width:640px){.hero{flex-direction:column;align-items:flex-start}}
</style>
"""


def html(block):
    """Render an HTML block with the indentation flattened (so Markdown never treats it as code)."""
    st.markdown("".join(line.strip() for line in block.splitlines()), unsafe_allow_html=True)


st.markdown(CSS, unsafe_allow_html=True)

LABELS = ["Home", "Daily report"]
CAN_SWITCH = True
try:
    tabs = st.tabs(LABELS, default="Home", key="main_tabs", on_change="rerun")
except TypeError:
    tabs = st.tabs(LABELS)
    CAN_SWITCH = False


def _goto(label):
    ss["main_tabs"] = label


def _open_button(col, label, tab_label, key):
    if CAN_SWITCH:
        col.button(label, key=key, on_click=_goto, args=(tab_label,), type="primary", width="stretch")
    else:
        col.caption(f"Open the “{tab_label}” tab above.")


# ============================================================ HOME
with tabs[0]:
    logo = _logo_b64()
    logo_html = (f'<div class="logo"><img src="data:image/png;base64,{logo}" alt="KeRRA"/></div>' if logo else "")
    html(f"""
    <div class="hero">
      <div>
        <div class="eyebrow">KENYA RURAL ROADS AUTHORITY</div>
        <h1>From site notes to a finished daily report.</h1>
        <p>Fill in the day's weather, work, plant, labour and photos — get the Word site diary in the KeRRA format.</p>
      </div>
      {logo_html}
    </div>
    """)

    c1, c2 = st.columns(2, gap="medium")
    with c1:
        html("""
        <div class="card"><div class="ico">📝</div>
        <h3>Daily report</h3>
        <p>The ten sections of the Daily Progress Report / Site Diary, in order.</p>
        <ul><li>Weather, rainfall and ground conditions</li>
            <li>Activities with chainages — today's length and totals worked out for you</li>
            <li>Equipment, labour force and materials delivered</li>
            <li>Delays, supervision vehicles, staff activities and progress photos</li></ul></div>
        """)
        _open_button(c1, "Open Daily report  →", "Daily report", "go_report")
    with c2:
        html("""
        <div class="card"><div class="ico">⏭️</div>
        <h3>Tomorrow in one step</h3>
        <p>Plant and labour rarely change day to day, so don't type them again.</p>
        <ul><li>Save today's report as a small file</li>
            <li>Tomorrow, open it and choose “Start the next day's report”</li>
            <li>Previous lengths become yesterday's totals; weather, photos and today's work are cleared</li>
            <li>Or reopen it unchanged to correct a report</li></ul></div>
        """)
        _open_button(c2, "Start from a saved report  →", "Daily report", "go_report2")

    html("""
    <div class="steps">
      <div class="step"><b>1 · ENTER</b><span>Project, weather, work done, plant and labour.</span></div>
      <div class="step"><b>2 · CHECK</b><span>Totals and warnings update as you type.</span></div>
      <div class="step"><b>3 · DOWNLOAD</b><span>Get the report as a Word (.docx) file, ready to print or sign.</span></div>
      <div class="step"><b>4 · SAVE</b><span>Keep the saved file to reopen or carry forward tomorrow.</span></div>
    </div>
    """)
    html(f"""
    <div class="foot">Kenya Rural Roads Authority &nbsp;·&nbsp; Daily Progress Report / Site Diary<br/>Version {VERSION} &nbsp;·&nbsp; {RELEASE_DATE}</div>
    """)

# ============================================================ widgets for the report tab
def dd(container, key, listname, label, fixed=None, placeholder="Select…"):
    """A dropdown. fixed=[...] -> only those values. Otherwise the list grows: a new typed value is added to it."""
    if fixed is not None:
        opts = list(fixed)
        if ss.get(key) not in opts:
            ss[key] = None
        return container.selectbox(label, opts, index=None, key=key, label_visibility="collapsed", placeholder=placeholder)
    opts = chainage_list() if listname == "chainage" else ss["lists"][listname]
    cur = ss.get(key)
    if cur and cur not in opts:                                   # a new value was typed (or came from a file)
        if listname == "chainage":
            added = rd.chainage_m(cur) is not None and learn("extra_chainage", cur)
        else:
            added = learn(listname, cur)
        if added:
            st.rerun()                                            # redraw every dropdown with the longer list
        opts = list(opts) + [cur]
    return container.selectbox(label, opts, index=None, key=key, accept_new_options=True,
                               label_visibility="collapsed", placeholder=placeholder)


def cell(container, table, rid, ci, kind):
    key, label = _key(table, rid, ci), rd.TABLE_COLS[table][ci]
    if kind == "text":
        container.text_input(label, key=key, label_visibility="collapsed")
    elif kind == "int":
        container.number_input(label, key=key, value=None, min_value=0, step=1, label_visibility="collapsed")
    elif kind == "num":
        container.number_input(label, key=key, value=None, min_value=0.0, step=0.5, format="%g",
                               label_visibility="collapsed")
    elif kind.startswith("fixed:"):
        dd(container, key, None, label, fixed=FIXED[kind[6:]], placeholder="—")
    else:
        dd(container, key, kind[3:], label)


def row_table(table, widths, headers, editable=True, add_label="＋ Add row"):
    w = list(widths) + ([0.45] if editable else [])
    hc = st.columns(w)
    for c, h in zip(hc, headers):
        c.markdown(f"<small><b>{h}</b></small>", unsafe_allow_html=True)
    for rid in list(ss["rows_" + table]):
        cols = st.columns(w)
        for ci, kind in enumerate(ROW_TABLES[table]):
            cell(cols[ci], table, rid, ci, kind)
        if editable:
            cols[-1].button("✕", key=f"del|{table}|{rid}", on_click=del_row, args=(table, rid), help="Delete this row")
    if editable:
        st.button(add_label, key=f"add|{table}", on_click=add_row, args=(table,))


def editor(name, **cfg):
    """A plain table (no dropdowns). Edits are kept in out_<name>."""
    out = st.data_editor(ss["base_" + name], key=f"ed_{name}_{ss['ver']}", num_rows="dynamic",
                         hide_index=True, width="stretch", **cfg)
    ss["out_" + name] = out
    return out


def activity_entries():
    """[(row id, caption)] for every activity row that has an operation or chainage."""
    out = []
    for rid in ss["rows_activities"]:
        g = lambda ci: ss.get(_key("activities", rid, ci)) or ""            # noqa: E731
        cap = rd.caption_text(g(0), g(1), g(2), g(3))
        if cap:
            out.append((rid, cap))
    return out


# ============================================================ DAILY REPORT
with tabs[1]:
    st.header("Daily progress report / site diary")
    st.caption("Fill in the sections below, check the totals, then download the Word report. "
               "Dropdowns that are not fixed accept a new typed value and remember it.")

    # ---------------- start from a saved report
    with st.expander("Start from a saved report (reopen, or carry forward to the next day)"):
        st.file_uploader("Saved report file (.json)", type=["json"], key="up_json")
        st.radio("What do you want to do with it?",
                 ["Reopen it to edit", "Start the next day's report (carry forward)"], key="open_mode")
        b1, b2 = st.columns(2)
        b1.button("Open", on_click=_open_saved, type="primary", width="stretch")
        b2.button("Start a blank report", on_click=_blank, width="stretch")
    if ss.get("_msg"):
        kind, text = ss.pop("_msg")
        getattr(st, kind)(text)

    # ---------------- project
    st.subheader("Project")
    st.text_input("Project name", key="project", placeholder="e.g. Upgrading to bitumen standard of … Road")
    c1, c2, c3, c4 = st.columns([1, 1.4, 1, 1])
    c1.text_input("Contract no.", key="contract_no", placeholder="e.g. RWC 682")
    c2.text_input("Contractor", key="contractor")
    c3.date_input("Date", key="date", format="DD/MM/YYYY")
    c4.number_input("Road length (m)", key="road_length", value=None, min_value=0.0, step=20.0, format="%g",
                    help="The chainage dropdowns are built from this: 0+000, 0+020, 0+040 … up to the road length.")
    if ss.get("road_length"):
        n = len(rd.chainage_options(ss["road_length"]))
        st.caption(f"{ss['date'].strftime('%A, %d %B %Y')}  ·  chainage list: {n} values, "
                   f"0+000 to {rd.fmt_chainage(ss['road_length'])}, every {rd.CHAINAGE_STEP} m")
    else:
        st.caption(f"{ss['date'].strftime('%A, %d %B %Y')}  ·  enter the road length to get the chainage dropdowns "
                   "(you can still type a chainage such as 2+260).")

    # ---------------- I weather
    st.subheader("I. Weather conditions")
    c1, c2 = st.columns(2)
    c1.caption("Station — from (chainage)")
    c2.caption("Station — to (chainage)")
    dd(c1, "station_from", "chainage", "Station from", placeholder="Select a chainage")
    dd(c2, "station_to", "chainage", "Station to", placeholder="Select a chainage")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.caption("Morning")
    dd(c1, "morning", "weather", "Morning")
    c2.caption("Afternoon")
    dd(c2, "afternoon", "weather", "Afternoon")
    c3.caption("Ground conditions")
    dd(c3, "ground", None, "Ground conditions", fixed=rd.GROUND_OPTIONS)
    c4.number_input("Max temperature (°C)", key="tmax", value=None, step=1.0, format="%g")
    c5.number_input("Min temperature (°C)", key="tmin", value=None, step=1.0, format="%g")
    st.markdown("**Rainfall (mm)** — the two rain-gauge locations")
    row_table("rainfall", [3, 1.2], ["Location", "Rainfall (mm)"], editable=False)

    # ---------------- II activities
    st.subheader("II. Activities")
    st.caption("Leave “Today (m)” empty and the length is worked out from the two chainages "
               "(2+260 to 2+660 = 400 m). Total = previous + today. RHS / LHS is joined to the chainage in the report.")
    row_table("activities", [2.6, 1.15, 1.15, 0.9, 1, 1, 1.6],
              ["Operation", "Chainage from", "Chainage to", "Side", "Prev. (m)", "Today (m)", "Remark / status"],
              add_label="＋ Add activity")
    acts = rd.activity_rows(table_df("activities"))
    if acts:
        st.caption("As it will appear in the report:")
        st.dataframe(pd.DataFrame([{"Operation": a["op"], "From": a["from_txt"], "To": a["to_txt"],
                                    "Prev.": rd.fmt_num(a["prev"]),
                                    "Today": rd.fmt_num(a["today"]) + (" (from chainages)" if a["auto"] else ""),
                                    "Total": rd.fmt_num(a["total"]), "Remark": a["remark"]} for a in acts]),
                     hide_index=True, width="stretch")

    # ---------------- III equipment
    st.subheader("III. Contractor's equipment / machinery on site")
    row_table("equipment", [2, 0.7, 1.4, 1.1, 1.1, 0.9, 1.3],
              ["Equipment / machinery", "Qty", "Reg. no.", "From", "To", "Idle hours", "Status"],
              add_label="＋ Add machine")

    # ---------------- IV visitors
    st.subheader("IV. Visitors")
    st.text_area("Visitors (one per line, or “Nil”)", key="visitors", height=80)

    # ---------------- V labour
    st.subheader("V. Contractor's labour force")
    row_table("labour", [2.4, 0.9, 1.2, 1.2, 1], ["Category", "Number", "From", "To", "Idle hours"],
              add_label="＋ Add category")

    # ---------------- VI materials
    st.subheader("VI. Materials delivered")
    row_table("materials", [3, 1.2, 1.2], ["Material", "Quantity", "Unit"], add_label="＋ Add material")

    # ---------------- VII delays
    st.subheader("VII. Accidents, delays & problems")
    st.text_area("One item per line", key="issues", height=100)

    # ---------------- VIII vehicles, IX staff
    st.subheader("VIII. Resident Engineer's supervision vehicles")
    editor("vehicles")
    st.subheader("IX. Resident Engineer's staff activities")
    editor("staff", column_config={"Activity": st.column_config.TextColumn(width="large")})

    # ---------------- X photos, one group per activity
    st.subheader("X. Progress photos")
    st.caption("Add the photos of each activity under its own heading. The caption is made automatically from the "
               "operation and chainages. Photos are cut to the same size and placed two per row.")
    photos = []
    entries = activity_entries()
    if not entries:
        st.info("Enter the activities in section II first — a photo group appears for each one.")
    for rid, cap in entries:
        files = st.file_uploader(f"Photos for: {cap}", type=["jpg", "jpeg", "png"], accept_multiple_files=True,
                                 key=f"ph|{rid}")
        if files:
            tc = st.columns(8)
            for i, f in enumerate(files):
                tc[i % 8].image(f, width=110)
                photos.append((f.getvalue(), cap))
    with st.expander("Other photos (with your own caption)"):
        ocap = st.text_input("Caption", key="ph|other_cap", placeholder="e.g. Site meeting")
        ofiles = st.file_uploader("Photos", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="ph|other")
        for f in ofiles or []:
            if ocap.strip():
                photos.append((f.getvalue(), ocap.strip()))
        if ofiles and not ocap.strip():
            st.warning("Type a caption, otherwise these photos are left out of the report.")

    # ---------------- check
    st.header("Check")
    state = current_state()
    tot = rd.totals(state)
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Activities", len(acts))
    m2.metric("Machines on site", tot["equipment_units"])
    m3.metric("Workers on site", tot["workers"])
    m4.metric("Photos", len(photos))

    warnings = []
    for label, key in (("Project name", "project"), ("Contract no.", "contract_no"), ("Contractor", "contractor")):
        if not str(state[key]).strip():
            warnings.append(f"{label} is empty.")
    if not state["road_length"]:
        warnings.append("Road length is empty, so the chainage dropdowns have no values.")
    if not acts:
        warnings.append("No activities entered for the day.")
    have_photo = {cap for _, cap in photos}
    for a in acts:
        for part, txt in (("from", a["from"]), ("to", a["to"])):
            if txt and rd.chainage_m(txt) is None:
                warnings.append(f"Activity “{a['op'] or '(no name)'}”: the “{part}” chainage “{txt}” is not in the form 2+260.")
        if a["today"] is None and a["op"] and not a["remark"]:
            warnings.append(f"Activity “{a['op']}” has no length for today and no remark.")
        if a["caption"] and a["caption"] not in have_photo:
            warnings.append(f"No photo for “{a['caption']}”.")
    if state["tmax"] is not None and state["tmin"] is not None and state["tmin"] > state["tmax"]:
        warnings.append("Minimum temperature is higher than the maximum.")
    if tot["idle_equipment"]:
        warnings.append(f"{tot['idle_equipment']} machine(s) have idle hours — explain in Accidents, delays & problems if needed.")
    if warnings:
        for w in warnings:
            st.warning(w, icon="⚠️")
    else:
        st.success("Everything needed is filled in.")

    # ---------------- output
    st.header("Download")
    fname = dx.report_filename(state)
    if st.button("Generate Word report", type="primary"):
        try:
            ss["docx_bytes"] = dx.build_report(state, photos)
            ss["docx_name"] = fname
        except Exception as e:                                    # noqa: BLE001
            st.error(f"The report could not be created: {e}")
    if ss.get("docx_bytes"):
        st.download_button("Download Word report", ss["docx_bytes"], file_name=ss.get("docx_name", fname),
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        st.caption("Regenerate after any change — the file above is the version made when you last pressed the button.")

    st.subheader("Save this report")
    st.caption("Download a small file that holds everything except the photos. Tomorrow, open it under "
               "“Start from a saved report” at the top of this tab.")
    st.download_button("Download saved report (.json)", rd.to_json(state),
                       file_name=f"Daily report - {state['date'].strftime('%d-%m-%Y')} (saved).json", mime="application/json")
