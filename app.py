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

st.set_page_config(page_title="KeRRA — Daily Report", page_icon="🛣️", layout="centered")

import report_data as rd              # noqa: E402
import report_docx as dx              # noqa: E402
from version import VERSION, RELEASE_DATE   # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
ss = st.session_state

SCALARS = ["project", "contract_no", "contractor", "date", "station_from", "station_to", "morning", "afternoon",
           "ground", "tmax", "tmin", "visitors", "issues"]
TABLES = ["rainfall", "activities", "equipment", "labour", "materials", "vehicles", "staff"]


# ============================================================ state handling
def apply_state(state):
    """Put a whole report (from defaults, a saved file or a carry-forward) into the form."""
    for k in SCALARS:
        ss[k] = state[k]
    for k in TABLES:
        ss["base_" + k] = state[k]
        ss.pop("out_" + k, None)
    ss["ver"] = ss.get("ver", 0) + 1          # new editor keys, so the tables reload
    ss.pop("docx_bytes", None)


def current_state():
    s = {k: ss[k] for k in SCALARS}
    for k in TABLES:
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

# ============================================================ DAILY REPORT
with tabs[1]:
    st.header("Daily progress report / site diary")
    st.caption("Fill in the sections below, check the totals, then download the Word report.")

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
    c1, c2, c3 = st.columns(3)
    c1.text_input("Contract no.", key="contract_no", placeholder="e.g. RWC 682")
    c2.text_input("Contractor", key="contractor")
    c3.date_input("Date", key="date", format="DD/MM/YYYY")
    st.caption(ss["date"].strftime("%A, %d %B %Y"))

    def editor(name, label_help=None, **cfg):
        """A table the user can add / delete rows in. Edits are kept in out_<name>."""
        out = st.data_editor(ss["base_" + name], key=f"ed_{name}_{ss['ver']}", num_rows="dynamic",
                             hide_index=True, width="stretch", **cfg)
        ss["out_" + name] = out
        return out

    # ---------------- I weather
    st.subheader("I. Weather conditions")
    c1, c2 = st.columns(2)
    c1.text_input("Station — from (chainage)", key="station_from", placeholder="2+260")
    c2.text_input("Station — to (chainage)", key="station_to", placeholder="2+660")
    c1, c2, c3 = st.columns(3)
    for col, key, label in ((c1, "morning", "Morning"), (c2, "afternoon", "Afternoon"), (c3, "ground", "Ground conditions")):
        base = rd.GROUND_OPTIONS if key == "ground" else rd.WEATHER_OPTIONS
        opts = base if ss[key] in base else base + [ss[key]]
        col.selectbox(label, opts, key=key)
    c1, c2 = st.columns(2)
    c1.number_input("Max temperature (°C)", key="tmax", value=None, step=1.0, format="%g")
    c2.number_input("Min temperature (°C)", key="tmin", value=None, step=1.0, format="%g")
    st.markdown("**Rainfall** — one row per rain gauge / location (leave empty for “Nil”)")
    editor("rainfall", column_config={"Rainfall (mm)": st.column_config.NumberColumn(min_value=0.0, format="%g")})

    # ---------------- II activities
    st.subheader("II. Activities")
    st.caption("Leave “Today (m)” empty and the length is worked out from the two chainages "
               "(2+260 to 2+660 = 400 m). Total = previous + today.")
    acts_df = editor("activities", column_config={
        "Operation": st.column_config.TextColumn(width="large"),
        "Chainage from": st.column_config.TextColumn("From", help="e.g. 2+260"),
        "Chainage to": st.column_config.TextColumn("To", help="e.g. 2+660"),
        "Previous (m)": st.column_config.NumberColumn("Prev. (m)", min_value=0.0, format="%g"),
        "Today (m)": st.column_config.NumberColumn("Today (m)", min_value=0.0, format="%g"),
        "Remark / status": st.column_config.TextColumn("Remark"),
    })
    acts = rd.activity_rows(acts_df)
    if acts:
        st.caption("As it will appear in the report:")
        st.dataframe(pd.DataFrame([{"Operation": a["op"], "From": a["from"], "To": a["to"],
                                    "Prev.": rd.fmt_num(a["prev"]),
                                    "Today": rd.fmt_num(a["today"]) + (" (from chainages)" if a["auto"] else ""),
                                    "Total": rd.fmt_num(a["total"])} for a in acts]),
                     hide_index=True, width="stretch")

    # ---------------- III equipment
    st.subheader("III. Contractor's equipment / machinery on site")
    st.caption("One row per machine. Times are typed as text, e.g. 8:00 am.")
    eq_df = editor("equipment", column_config={
        "Equipment / machinery": st.column_config.TextColumn("Equipment", width="medium"),
        "Qty": st.column_config.NumberColumn(min_value=0, step=1, format="%d", width="small"),
        "Reg. no.": st.column_config.TextColumn(width="small"),
        "From": st.column_config.TextColumn(width="small"),
        "To": st.column_config.TextColumn(width="small"),
        "Idle hours": st.column_config.NumberColumn(min_value=0.0, format="%g", width="small"),
        "Status": st.column_config.SelectboxColumn(options=rd.STATUS_OPTIONS, width="small"),
    })

    # ---------------- IV visitors
    st.subheader("IV. Visitors")
    st.text_area("Visitors (one per line, or “Nil”)", key="visitors", height=80)

    # ---------------- V labour
    st.subheader("V. Contractor's labour force")
    lab_df = editor("labour", column_config={
        "Number": st.column_config.NumberColumn(min_value=0, step=1, format="%d"),
        "Idle hours": st.column_config.NumberColumn(min_value=0.0, format="%g"),
    })

    # ---------------- VI materials
    st.subheader("VI. Materials delivered")
    st.caption("Include the unit in the quantity, e.g. “30 m³” or “200 bags”.")
    editor("materials")

    # ---------------- VII delays
    st.subheader("VII. Accidents, delays & problems")
    st.text_area("One item per line", key="issues", height=100)

    # ---------------- VIII vehicles, IX staff
    st.subheader("VIII. Resident Engineer's supervision vehicles")
    editor("vehicles")
    st.subheader("IX. Resident Engineer's staff activities")
    editor("staff", column_config={"Activity": st.column_config.TextColumn(width="large")})

    # ---------------- X photos
    st.subheader("X. Progress photos")
    files = st.file_uploader("Add photos (JPG or PNG)", type=["jpg", "jpeg", "png"], accept_multiple_files=True,
                             key="photos")
    photos = []
    for i, f in enumerate(files or []):
        pc1, pc2 = st.columns([1, 2])
        pc1.image(f, width="stretch")
        cap = pc2.text_input(f"Caption for photo {i + 1}", key=f"cap_{f.file_id}", placeholder="e.g. Subbase stabilisation 2+260 to 2+660")
        photos.append((f.getvalue(), cap))

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
    if not acts:
        warnings.append("No activities entered for the day.")
    for a in acts:
        for part, txt in (("from", a["from"]), ("to", a["to"])):
            if txt and rd.chainage_m(txt) is None:
                warnings.append(f"Activity “{a['op'] or '(no name)'}”: the “{part}” chainage “{txt}” is not in the form 2+260.")
        if a["today"] is None and a["op"] and not a["remark"]:
            warnings.append(f"Activity “{a['op']}” has no length for today and no remark.")
    if state["tmax"] is not None and state["tmin"] is not None and state["tmin"] > state["tmax"]:
        warnings.append("Minimum temperature is higher than the maximum.")
    if not photos:
        warnings.append("No progress photo attached.")
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
