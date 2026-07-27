"""
Scoresheet Builder — Streamlit app.
Reuses generator.py (the Excel engine) unchanged.
Run locally:  streamlit run streamlit_app.py
"""
import io
import re
import streamlit as st

import generator as g

st.set_page_config(page_title="Scoresheet Builder", page_icon="📋", layout="centered")

# --- light brand styling ---
st.markdown(
    """
    <style>
      .block-container{max-width:820px;padding-top:2.2rem}
      .sb-mast{background:linear-gradient(180deg,#0F2137,#16304C);border-bottom:3px solid #0E7C86;
        border-radius:12px;padding:26px 28px;margin-bottom:22px;color:#EAF1F4}
      .sb-brand{font-weight:700;letter-spacing:.14em;font-size:13px;color:#fff}
      .sb-mast h1{color:#fff;font-size:32px;margin:10px 0 8px;font-weight:700}
      .sb-mast p{color:#C4D4DC;margin:0;font-size:15px}
      .sb-step{font-size:19px;font-weight:700;color:#12233A;margin:6px 0 2px}
      .sb-online{color:#0E7C86;font-weight:700;font-size:13px;letter-spacing:.05em;text-transform:uppercase}
      .sb-offline{color:#B4741F;font-weight:700;font-size:13px;letter-spacing:.05em;text-transform:uppercase}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="sb-mast">
      <div class="sb-brand">SHL · DEVELOPMENT CENTRES</div>
      <h1>Scoresheet Builder</h1>
      <p>Add a competency framework, choose the tools, download the assessor workbook — the
      Detailed Integration tab is wired for you.</p>
    </div>
    """,
    unsafe_allow_html=True,
)


def _safe_name(name, fallback="DC_Scoresheet"):
    name = (name or "").strip()
    if not name:
        return fallback
    name = re.sub(r"[^A-Za-z0-9 _-]", "", name).strip().replace(" ", "_")
    return name or fallback


# ---------------- STEP 1: framework ----------------
st.markdown('<div class="sb-step">1 · Framework</div>', unsafe_allow_html=True)
st.caption("Two columns — Competency and Behaviour. Start from the template so the format is right.")

# build the template bytes for the download button
_tbuf = io.BytesIO()
g.build_template().save(_tbuf)
st.download_button(
    "⬇ Download framework template",
    data=_tbuf.getvalue(),
    file_name="Competency_Framework_Template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

mode = st.radio("Add the framework by", ["Upload file", "Paste columns"], horizontal=True, label_visibility="collapsed")
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

# ---------------- STEP 2: tools ----------------
st.markdown('<div class="sb-step">2 · Tools</div>', unsafe_allow_html=True)
st.caption("Only the tools you tick become tabs, and only those feed the Detailed Integration.")

col_on, col_off = st.columns(2)
with col_on:
    st.markdown('<div class="sb-online">Online · paste-stens</div>', unsafe_allow_html=True)
    opq = st.checkbox("OPQ — Personality")
    mq = st.checkbox("MQ — Motivation")
    sjt = st.checkbox("SJT — Situational judgement")
    if sjt:
        sjt_variant = st.selectbox("SJT type", ["Managerial", "Executive"],
                                   help="Shapes the SJT sheet and the Onlines Data columns")
    else:
        sjt_variant = "Managerial"
    verify = st.checkbox("Verify — Ability (Inductive / Deductive / Numerical)")
with col_off:
    st.markdown('<div class="sb-offline">Offline · consultant writes rubrics</div>', unsafe_allow_html=True)
    cs = st.checkbox("Case Study — Rubric anchors")
    gd = st.checkbox("Group Discussion — Rubric anchors")
    inbox = st.checkbox("Inbox Simulation — Rubric anchors")
    wa = st.checkbox("Written Analysis — Rubric anchors")
    rp = st.checkbox("Role Play — Rubric anchors")
    bei = st.checkbox("BEI — Interview questions")

selected = [t for t, on in [("OPQ", opq), ("MQ", mq), ("SJT", sjt), ("Verify", verify),
                            ("Case Study", cs), ("Group Discussion", gd),
                            ("Inbox Simulation", inbox), ("Written Analysis", wa),
                            ("Role Play", rp), ("BEI", bei)] if on]

# ---------------- STEP 3: generate ----------------
st.markdown('<div class="sb-step">3 · Generate</div>', unsafe_allow_html=True)
dc_name = st.text_input("File name (optional)", placeholder="e.g. Navitasys_DC_June")

if st.button("Build scoresheet", type="primary"):
    st.session_state.pop("result", None)
    if not selected:
        st.error("Pick at least one assessment tool.")
    elif not uploaded and not pasted.strip():
        st.error("Add a framework — upload the template or paste two columns.")
    else:
        try:
            if uploaded is not None:
                framework, warnings = g.parse_framework(uploaded)
            else:
                framework, warnings = g.parse_framework_text(pasted)
            wb = g.build_workbook(framework, selected, sjt_variant=sjt_variant)
            buf = io.BytesIO()
            wb.save(buf)
            n_comp = len({c for c, _ in framework})
            st.session_state.result = buf.getvalue()
            st.session_state.result_name = _safe_name(dc_name) + ".xlsx"
            st.session_state.info = f"{len(framework)} behaviours · {n_comp} competencies · {len(selected)} tools"
            st.session_state.warnings = warnings
        except ValueError as e:
            st.error(str(e))
        except Exception:
            st.error("That file couldn't be read. Make sure it's the .xlsx template with "
                     "Competency and Behaviour columns.")

if st.session_state.get("result"):
    st.success("Ready — " + st.session_state.info)
    for w in st.session_state.get("warnings", []):
        st.warning(w)
    st.download_button(
        "⬇ Download scoresheet",
        data=st.session_state.result,
        file_name=st.session_state.result_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
