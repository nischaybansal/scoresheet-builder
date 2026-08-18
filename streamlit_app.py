"""
Scoresheet Builder - Streamlit app.

This file is UI only. Every decision about what ends up in the workbook lives in
generator.py, which is imported unchanged - so the interface can be restyled or replaced
without touching the Excel engine.

Run locally:  streamlit run streamlit_app.py
"""
import io
import re

import streamlit as st

import generator as g

st.set_page_config(page_title="Scoresheet Builder", page_icon="clipboard",
                   layout="centered", initial_sidebar_state="expanded")

# --------------------------------------------------------------------------- styling
st.markdown(
    """
    <style>
      .block-container{max-width:880px;padding-top:2.2rem}

      /* masthead */
      .sb-mast{background:linear-gradient(180deg,#0F2137,#16304C);border-bottom:3px solid #0E7C86;
        border-radius:12px;padding:26px 28px;margin-bottom:20px;color:#EAF1F4}
      .sb-brand{font-weight:700;letter-spacing:.14em;font-size:13px;color:#fff}
      .sb-mast h1{color:#fff;font-size:32px;margin:10px 0 8px;font-weight:700}
      .sb-mast p{color:#C4D4DC;margin:0;font-size:15px}

      /* step headings */
      .sb-step{display:flex;align-items:center;gap:10px;font-size:19px;font-weight:700;
        color:#12233A;margin:26px 0 2px}
      .sb-num{display:inline-flex;align-items:center;justify-content:center;width:26px;height:26px;
        border-radius:50%;background:#0E7C86;color:#fff;font-size:13px;font-weight:700;flex:none}

      /* tool group headers */
      .sb-group{display:flex;justify-content:space-between;align-items:baseline;
        border-bottom:2px solid #E3EAEE;padding-bottom:6px;margin-bottom:10px}
      .sb-online{color:#0E7C86;font-weight:700;font-size:13px;letter-spacing:.05em;
        text-transform:uppercase}
      .sb-offline{color:#B4741F;font-weight:700;font-size:13px;letter-spacing:.05em;
        text-transform:uppercase}
      .sb-count{font-size:12px;color:#6B7C88;font-weight:600}

      /* framework preview table */
      .sb-tbl{width:100%;border-collapse:collapse;font-size:13px}
      .sb-tbl th{background:#16304C;color:#fff;text-align:left;padding:6px 10px;font-weight:600}
      .sb-tbl td{border-bottom:1px solid #E3EAEE;padding:6px 10px;color:#12233A}
      .sb-tbl td.num{text-align:center;width:90px;color:#6B7C88}

      /* sidebar summary */
      .sb-sum-h{font-weight:700;font-size:13px;letter-spacing:.08em;text-transform:uppercase;
        color:#0E7C86;margin:14px 0 4px}
      .sb-line{display:flex;justify-content:space-between;font-size:13px;padding:3px 0;
        border-bottom:1px solid #E7EDF0;color:#12233A}
      .sb-line span:last-child{font-weight:600;text-align:right}
      .sb-none{font-size:13px;color:#8A98A3;font-style:italic;padding:3px 0}
      .sb-pill{display:inline-block;background:#EEF3F5;border:1px solid #D6E2E7;border-radius:11px;
        padding:1px 9px;font-size:12px;margin:2px 3px 2px 0;color:#12233A}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="sb-mast">
      <div class="sb-brand">SHL &middot; DEVELOPMENT CENTRES</div>
      <h1>Scoresheet Builder</h1>
      <p>Add a competency framework, choose the tools, download the assessor workbook &mdash;
      the Detailed Integration tab is wired for you.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# --------------------------------------------------------------------------- helpers
ONLINE_TOOLS = [
    ("OPQ",    "Personality — 32 dimensions"),
    ("MQ",     "Motivation — 18 dimensions"),
    ("SJT",    "Situational judgement"),
    ("Verify", "Ability — inductive, deductive, numerical"),
]
OFFLINE_TOOLS = [
    ("Case Study",         "Rubric anchors"),
    ("Group Discussion",   "Rubric anchors"),
    ("Inbox Simulation",   "Rubric anchors"),
    ("Written Analysis",   "Rubric anchors"),
    ("Business Role Play", "Rubric anchors"),
    ("Coaching Role Play", "Rubric anchors"),
    ("BEI",                "Interview questions"),
]


def _safe_name(name, fallback="DC_Scoresheet"):
    name = (name or "").strip()
    if not name:
        return fallback
    name = re.sub(r"[^A-Za-z0-9 _-]", "", name).strip().replace(" ", "_")
    return name or fallback


def _step(num, title, caption=None):
    st.markdown(f'<div class="sb-step"><span class="sb-num">{num}</span>{title}</div>',
                unsafe_allow_html=True)
    if caption:
        st.caption(caption)


def _n(count, word, plural=None):
    """'1 tool' / '3 tools' - counts read badly without this."""
    return f"{count} {word if count == 1 else (plural or word + 's')}"


def _esc(v):
    return (str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


# --------------------------------------------------------------------------- 1. framework
_step("1", "Framework",
      "Competency and Behaviour are required. Theme and Level are optional - and you can "
      "rename those two columns to whatever the client calls them.")

_tbuf = io.BytesIO()
g.build_template().save(_tbuf)
st.download_button(
    "Download framework template",
    data=_tbuf.getvalue(),
    file_name="Competency_Framework_Template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

mode = st.radio("Add the framework by", ["Upload file", "Paste columns"],
                horizontal=True, label_visibility="collapsed")
uploaded, pasted = None, ""
if mode == "Upload file":
    uploaded = st.file_uploader("Upload the filled-in .xlsx", type=["xlsx", "xlsm"])
else:
    pasted = st.text_area(
        "Paste the framework",
        height=170,
        placeholder="Drives Results, Sets clear goals and holds the team accountable\n"
                    "Drives Results, Removes obstacles to keep work moving\n"
                    "Builds Collaboration, Works across teams to solve shared problems\n\n"
                    "One behaviour per line. Leave the competency blank to repeat the one above.",
    )

# Parse live so the tool step can offer per-tool competency mapping.
framework, fw_warnings, fw_error, fw_meta = None, [], None, None
try:
    if uploaded is not None:
        uploaded.seek(0)
        framework, fw_warnings, fw_meta = g.parse_framework(uploaded)
    elif pasted.strip():
        framework, fw_warnings, fw_meta = g.parse_framework_text(pasted)
except ValueError as e:
    fw_error = str(e)
except Exception:
    fw_error = ("That file couldn't be read. Make sure it's the .xlsx template with "
                "Competency and Behaviour columns.")

comps_list = []
if framework:
    for c, _ in framework:
        if c not in comps_list:
            comps_list.append(c)

if fw_error:
    st.error(fw_error)
elif framework:
    meta = fw_meta or {}
    themes, levels = meta.get("themes") or {}, meta.get("levels") or {}
    extras = []
    if themes:
        extras.append(meta.get("theme_label") or "Theme")
    if levels:
        extras.append(meta.get("level_label") or "Level")
    st.success(f"{_n(len(comps_list), 'competency', 'competencies')} · "
               f"{_n(len(framework), 'behaviour')}"
               + (f" · also read: {', '.join(extras)}" if extras else ""))

    # Warnings the parser raised - rows it had to skip. Surfaced here, not just at build.
    for w in fw_warnings:
        st.warning(w)

    with st.expander("Check the framework", expanded=False):
        head = (f"<th>{_esc(meta.get('theme_label') or 'Theme')}</th>" if themes else "") \
               + "<th>Competency</th><th class='num'>Behaviours</th>"
        body = []
        for c in comps_list:
            n = sum(1 for cc, _ in framework if cc == c)
            cells = (f"<td>{_esc(themes.get(c, ''))}</td>" if themes else "") \
                    + f"<td>{_esc(c)}</td><td class='num'>{n}</td>"
            body.append(f"<tr>{cells}</tr>")
        st.markdown(f"<table class='sb-tbl'><tr>{head}</tr>{''.join(body)}</table>",
                    unsafe_allow_html=True)
        if levels:
            st.caption(f"'{meta.get('level_label') or 'Level'}' is read per behaviour. Where a "
                       "competency carries one value it is merged into a single cell; where it "
                       "varies behaviour to behaviour each row keeps its own.")

# A new framework invalidates any per-tool competency choice held in session state.
# Cleared BEFORE those widgets render, so nothing is written to a live widget's value.
_sig = tuple(comps_list)
if st.session_state.get("_fw_sig") != _sig:
    for k in [k for k in list(st.session_state) if k.startswith("comps_")]:
        del st.session_state[k]
    st.session_state["_fw_sig"] = _sig

# --------------------------------------------------------------------------- 2. tools
_step("2", "Tools",
      "Only ticked tools become tabs. Under each one, drop the competencies that tool does "
      "not measure - those show as a black block in the Detailed Integration.")
if not framework:
    st.info("Add the framework above to choose competencies per tool.")


def comp_picker(tool, enabled):
    """Per-tool competency mapping, folded away so four ticked tools don't produce four
    long lists. The label carries the count, so it reads without opening."""
    if not (enabled and comps_list):
        return None
    key = f"comps_{tool}"
    chosen = st.session_state.get(key, comps_list)
    n = len(chosen)
    flag = "" if n else "  -  none selected"
    with st.expander(f"Competencies: {n} of {len(comps_list)}{flag}", expanded=False):
        return st.multiselect("Competencies this tool measures", comps_list,
                              default=comps_list, key=key, label_visibility="collapsed")


def group_header(css, label, picked, total):
    st.markdown(
        f'<div class="sb-group"><span class="{css}">{label}</span>'
        f'<span class="sb-count">{picked} of {total}</span></div>',
        unsafe_allow_html=True)


picked_online = sum(1 for t, _ in ONLINE_TOOLS if st.session_state.get(f"tool_{t}"))
picked_offline = sum(1 for t, _ in OFFLINE_TOOLS if st.session_state.get(f"tool_{t}"))

col_on, col_off = st.columns(2)
chosen, tool_comps = {}, {}

with col_on:
    group_header("sb-online", "Online &middot; psychometrics", picked_online, len(ONLINE_TOOLS))
    for tool, blurb in ONLINE_TOOLS:
        on = st.checkbox(f"**{tool}** - {blurb}", key=f"tool_{tool}")
        chosen[tool] = on
        if tool == "SJT":
            sjt_variant = (st.selectbox("SJT type", ["Managerial", "Executive"],
                                        key="sjt_variant",
                                        help="Shapes the SJT sheet and the Onlines Data columns")
                           if on else "Managerial")
        if tool == "MQ":
            mq_mode = (st.selectbox(
                "MQ rubric style", ["5-point", "low-high"], key="mq_mode",
                format_func=lambda v: ("Rating 1-5 rubric" if v == "5-point"
                                       else "Low / High rubric (stens 1-3 / 8-10)"),
                help="Low/High writes two rubric columns; only stens 1-3 or 8-10 produce "
                     "a statement") if on else "5-point")
            if on and mq_mode == "low-high":
                st.caption("Standalone profile: no competency mapping, and it stays out of "
                           "Detailed Integration and Integration.")
                tool_comps["MQ"] = None
                continue
        picked = comp_picker(tool, on)
        if picked is not None:
            tool_comps[tool] = picked

with col_off:
    group_header("sb-offline", "Offline &middot; consultant writes rubrics",
                 picked_offline, len(OFFLINE_TOOLS))
    for tool, blurb in OFFLINE_TOOLS:
        on = st.checkbox(f"**{tool}** - {blurb}", key=f"tool_{tool}")
        chosen[tool] = on
        picked = comp_picker(tool, on)
        if picked is not None:
            tool_comps[tool] = picked

selected = [t for t in g.CANONICAL if chosen.get(t)]
tool_comps = {t: c for t, c in tool_comps.items() if c is not None}
any_online = any(chosen.get(t) for t, _ in ONLINE_TOOLS)

# --------------------------------------------------------------------------- 3. options
_step("3", "Workbook options")
if any_online:
    _mode = st.radio(
        "How do the online scores get in?",
        ["Consultant pastes the score extract", "Assessors type the stens"],
        help="Pasting adds the Consultant Paste and Onlines Data sheets and fills the stens "
             "automatically. Typing removes both sheets; stens go in on each tool sheet.")
    online_paste = _mode.startswith("Consultant")
    st.caption("Pick the participant in Detailed Integration and every sten fills itself."
               if online_paste else
               "No Consultant Paste or Onlines Data sheet - the sten goes in on the tool "
               "sheet next to each mapped trait.")
else:
    online_paste = True
    st.caption("Tick an online tool to choose how the stens arrive.")

show_summary = st.checkbox(
    "Add the competency summary table to Detailed Integration", value=False,
    help="Individual / Cohort / Ideal / Previous DC")

# --------------------------------------------------------------------------- 4. generate
_step("4", "Generate")

blockers = []
if framework is None:
    blockers.append("Add a framework in step 1.")
if not selected:
    blockers.append("Tick at least one assessment tool in step 2.")
empty_tools = [t for t, c in tool_comps.items() if not c]
if empty_tools:
    blockers.append("No competencies selected for " + ", ".join(empty_tools)
                    + " - either pick some or untick the tool.")

dc_name = st.text_input("File name (optional)", placeholder="e.g. Navitasys_DC_June")

for b in blockers:
    st.warning(b)

cfg_sig = repr((tuple(framework or ()), tuple(selected),
                sjt_variant, mq_mode, online_paste, show_summary,
                {k: tuple(v) for k, v in sorted(tool_comps.items())}))

if st.button("Build scoresheet", type="primary", disabled=bool(blockers)):
    st.session_state.pop("result", None)
    try:
        wb = g.build_workbook(framework, selected, sjt_variant=sjt_variant,
                              tool_comps=tool_comps, mq_mode=mq_mode,
                              online_paste=online_paste, meta=fw_meta,
                              show_summary=show_summary)
        buf = io.BytesIO()
        wb.save(buf)
        st.session_state.result = buf.getvalue()
        st.session_state.result_name = _safe_name(dc_name) + ".xlsx"
        st.session_state.result_sig = cfg_sig
        st.session_state.info = (f"{_n(len(comps_list), 'competency', 'competencies')} · "
                                 f"{_n(len(framework), 'behaviour')} · "
                                 f"{_n(len(selected), 'tool')}")
        st.session_state.warnings = fw_warnings
    except ValueError as e:
        st.error(str(e))
    except Exception:
        st.error("The workbook couldn't be built. Check the framework file and try again.")

if st.session_state.get("result"):
    if st.session_state.get("result_sig") != cfg_sig:
        st.info("Settings have changed since this was built - build again to include them.")
    else:
        st.success("Ready · " + st.session_state.info)
    for w in st.session_state.get("warnings", []):
        st.warning(w)
    st.download_button(
        "Download scoresheet",
        data=st.session_state.result,
        file_name=st.session_state.result_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# --------------------------------------------------------------------------- live summary
# Rendered last so it reflects every choice made above; appears in the sidebar, where it
# stays visible while the form is scrolled.
with st.sidebar:
    st.markdown('<div class="sb-sum-h">This build</div>', unsafe_allow_html=True)
    if framework:
        st.markdown(
            f'<div class="sb-line"><span>Competencies</span><span>{len(comps_list)}</span></div>'
            f'<div class="sb-line"><span>Behaviours</span><span>{len(framework)}</span></div>',
            unsafe_allow_html=True)
    else:
        st.markdown('<div class="sb-none">No framework yet</div>', unsafe_allow_html=True)

    st.markdown('<div class="sb-sum-h">Tools</div>', unsafe_allow_html=True)
    if selected:
        st.markdown("".join(f'<span class="sb-pill">{_esc(t)}</span>' for t in selected),
                    unsafe_allow_html=True)
        if framework:
            st.caption(f"{sum(1 for t in selected if g.TOOLS[t]['family'] == 'online')} online, "
                       f"{sum(1 for t in selected if g.TOOLS[t]['family'] == 'offline')} offline")
    else:
        st.markdown('<div class="sb-none">No tools ticked</div>', unsafe_allow_html=True)

    st.markdown('<div class="sb-sum-h">Settings</div>', unsafe_allow_html=True)
    rows = []
    if chosen.get("SJT"):
        rows.append(("SJT", sjt_variant))
    if chosen.get("MQ"):
        rows.append(("MQ rubric", "Low / High" if mq_mode == "low-high" else "Rating 1-5"))
    if any_online:
        rows.append(("Stens", "Pasted" if online_paste else "Typed"))
    rows.append(("Summary table", "On" if show_summary else "Off"))
    st.markdown("".join(f'<div class="sb-line"><span>{_esc(k)}</span><span>{_esc(v)}</span></div>'
                        for k, v in rows), unsafe_allow_html=True)

    if blockers:
        st.markdown('<div class="sb-sum-h">Before building</div>', unsafe_allow_html=True)
        for b in blockers:
            st.caption(b)
