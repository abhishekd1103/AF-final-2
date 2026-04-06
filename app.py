"""
Arc Flash Study Report Generator — Streamlit App
Improved version: maintenance mode per-scenario, flexible comparison, report options.
streamlit run app.py
"""
import streamlit as st
import pandas as pd
import os, sys, tempfile, copy
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from modules.config import AF_COLS, is_wide_table
from modules.af_processor import (
    read_excel, auto_map, filter_junk, apply_rules,
    summary_stats, auto_conclusion, detect_voltages,
)
from services.comparison_service import COMPARISON_PARAMS
from services.report_service import generate_report

st.set_page_config(page_title="AF Report Generator", page_icon="⚡", layout="wide")

# ── Session-state defaults ───────────────────────────────────────────────────
_DEFAULTS = dict(
    template_path=None,
    scenarios=[],
    col_map={},
    project={},
    conclusion={},
    selected_cols=[k for k, v in AF_COLS.items() if v["req"]],
    # Comparison (replaces the old maint_mode / maint_df block)
    comparison_mode=False,
    scenario_a="",
    scenario_b="",
    comparison_params=["total_energy", "afb"],
    # Report options
    show_mitigation_section=False,
    annexures=[],
    consultant_logo=None,
    client_logo=None,
)
for k, v in _DEFAULTS.items():
    st.session_state.setdefault(k, copy.deepcopy(v))

# ── Sidebar ───────────────────────────────────────────────────────────────────
st.sidebar.title("⚡ AF Report Platform")
step = st.sidebar.radio("Step", [
    "1. Upload & Scenarios",
    "2. Preview & Edit",
    "3. Settings & Conclusion",
    "4. Generate",
])
st.sidebar.markdown("---")
st.sidebar.metric("Template", "Yes" if st.session_state.template_path else "No")
st.sidebar.metric("Scenarios", str(len(st.session_state.scenarios)))

# Count how many scenarios are maintenance mode
maint_count = sum(1 for s in st.session_state.scenarios if s.get("is_maintenance", False))
normal_count = len(st.session_state.scenarios) - maint_count
mode_label = "Sequential"
if maint_count > 0 and normal_count > 0:
    mode_label = "Maint + Normal"
elif maint_count > 0:
    mode_label = "Maintenance Only"
st.sidebar.metric("Mode", mode_label)
if st.session_state.comparison_mode:
    st.sidebar.success("🔀 Comparison ON")


# ════════════════════════════════════════════════════════════════════════════
# STEP 1 — Upload Template, Logos & Scenarios
# ════════════════════════════════════════════════════════════════════════════
if step == "1. Upload & Scenarios":
    st.header("Step 1 — Upload Template, Logos & Scenarios")

    c1, c2, c3 = st.columns(3)
    with c1:
        tpl = st.file_uploader("Report Template (.docx)", type=["docx"], key="u_tpl")
        if tpl:
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
            tmp.write(tpl.read()); tmp.close()
            st.session_state.template_path = tmp.name
            st.success("Template loaded")
    with c2:
        cl = st.file_uploader("Consultant Logo", type=["png", "jpg", "jpeg"], key="u_cl")
        if cl:
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            tmp.write(cl.read()); tmp.close()
            st.session_state.consultant_logo = tmp.name
            st.success("Consultant logo loaded")
    with c3:
        cll = st.file_uploader("Client Logo", type=["png", "jpg", "jpeg"], key="u_cll")
        if cll:
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            tmp.write(cll.read()); tmp.close()
            st.session_state.client_logo = tmp.name
            st.success("Client logo loaded")

    st.markdown("---")

    # ── Add New Scenario ────────────────────────────────────────────────────
    st.subheader("Add Scenario")
    with st.expander("➕ New Scenario", expanded=not st.session_state.scenarios):
        sc1, sc2 = st.columns(2)
        with sc1:
            sname = st.text_input("Scenario Name", key="ns_n",
                                  placeholder="e.g. Normal Operation")
            sdesc = st.text_input("Description", key="ns_d",
                                  placeholder="e.g. All breakers at normal settings")
            # ── Maintenance Mode Checkbox (Part 1.1) ────────────────────────
            st.markdown("**Scenario Configuration**")
            is_maint = st.checkbox(
                "☑ Maintenance Mode (Mitigation Mode)",
                key="ns_maint",
                help=(
                    "Mark this scenario as a Maintenance / Mitigation scenario. "
                    "It will appear in Section 6 and can be selected as Scenario B "
                    "in the comparison table."
                ),
            )
            if is_maint:
                st.info("ℹ️ This scenario will be treated as the **mitigated** case.")
        with sc2:
            sfile = st.file_uploader("Result Excel", type=["xlsx", "xls"], key="ns_f")

        if st.button("Add Scenario", type="primary") and sname and sfile:
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            tmp.write(sfile.read()); tmp.close()
            try:
                raw = read_excel(tmp.name)
                cm = auto_map(list(raw.columns))
                if not st.session_state.col_map:
                    st.session_state.col_map = cm
                proc = apply_rules(filter_junk(raw, cm), cm)
                st.session_state.scenarios.append({
                    "name": sname,
                    "desc": sdesc,
                    "filepath": tmp.name,
                    "df_raw": raw,
                    "df": proc,
                    "exclude_set": set(),
                    "is_maintenance": is_maint,   # ← NEW: per-scenario flag
                })
                st.success("✅ {}: {} buses loaded{}".format(
                    sname, len(proc),
                    " [Maintenance Mode]" if is_maint else ""))
                st.rerun()
            except Exception as e:
                st.error(str(e))

    # ── Existing Scenarios ───────────────────────────────────────────────────
    for i, s in enumerate(st.session_state.scenarios):
        cm = st.session_state.col_map
        badge = "🔧 " if s.get("is_maintenance", False) else ""
        with st.expander("{}{} — {} buses".format(badge, s["name"], len(s["df"]))):
            stats = summary_stats(s["df"], cm)
            mc = st.columns(5)
            mc[0].metric("Buses", stats["total"])
            mc[1].metric("Max Energy", stats.get("max_energy", "N/A"))
            mc[2].metric("Max Level", stats.get("max_level", "N/A"))
            mc[3].metric("FCT Unknown", stats.get("fct_unknown", 0))
            mc[4].metric("Mode", "Maintenance" if s.get("is_maintenance") else "Normal")

            vl = detect_voltages(s["df"], cm)
            if vl:
                st.info("Voltages: {}".format(vl))
            if stats.get("maintenance_recommended") and not s.get("is_maintenance"):
                st.warning("⚠️ Level F/G detected — consider adding a Maintenance Mode scenario!")

            # Allow toggling maintenance mode on existing scenarios
            new_maint = st.checkbox(
                "Maintenance Mode (Mitigation Mode)",
                value=s.get("is_maintenance", False),
                key="maint_toggle_{}".format(i),
            )
            if new_maint != s.get("is_maintenance", False):
                s["is_maintenance"] = new_maint
                st.rerun()

            if st.button("🗑 Remove", key="rm_{}".format(i)):
                st.session_state.scenarios.pop(i)
                st.rerun()

    st.markdown("---")

    # ── Comparison Mode (Part 1.2 & 1.3) ─────────────────────────────────────
    st.subheader("🔀 Scenario Comparison")
    st.session_state.comparison_mode = st.toggle(
        "Enable Comparison Mode",
        value=st.session_state.comparison_mode,
        help="Compare two scenarios side-by-side in the report comparison table.",
    )

    if st.session_state.comparison_mode:
        if len(st.session_state.scenarios) < 2:
            st.warning("Add at least 2 scenarios to enable comparison.")
        else:
            sc_names = [s["name"] for s in st.session_state.scenarios]
            col_a, col_b = st.columns(2)
            with col_a:
                st.session_state.scenario_a = st.selectbox(
                    "Scenario A (Base)",
                    sc_names,
                    index=sc_names.index(st.session_state.scenario_a)
                          if st.session_state.scenario_a in sc_names else 0,
                    key="sel_a",
                )
            with col_b:
                # Default B to the first maintenance-mode scenario, if any
                default_b = next(
                    (s["name"] for s in st.session_state.scenarios
                     if s.get("is_maintenance") and s["name"] != st.session_state.scenario_a),
                    sc_names[-1]
                )
                st.session_state.scenario_b = st.selectbox(
                    "Scenario B (Compare / Mitigated)",
                    [n for n in sc_names if n != st.session_state.scenario_a],
                    index=0,
                    key="sel_b",
                )

            st.markdown("**Comparison Parameters**")
            st.caption("Architecture supports adding new parameters in `comparison_service.py` without code changes here.")
            all_params = list(COMPARISON_PARAMS.keys())
            st.session_state.comparison_params = st.multiselect(
                "Parameters to compare",
                all_params,
                default=st.session_state.comparison_params or ["total_energy", "afb"],
                format_func=lambda k: COMPARISON_PARAMS[k][0],
                key="cmp_params",
            )
            if st.session_state.scenario_a == st.session_state.scenario_b:
                st.error("Scenario A and B must be different.")


# ════════════════════════════════════════════════════════════════════════════
# STEP 2 — Preview & Edit
# ════════════════════════════════════════════════════════════════════════════
elif step == "2. Preview & Edit":
    st.header("Step 2 — Column Selection & Preview")
    if not st.session_state.scenarios:
        st.warning("Add scenarios first.")
        st.stop()

    ao = [(k, v["label"]) for k, v in AF_COLS.items()]
    st.session_state.selected_cols = st.multiselect(
        "Report columns", [k for k, _ in ao],
        default=st.session_state.selected_cols,
        format_func=lambda k: AF_COLS[k]["label"])

    n_cols = len(st.session_state.selected_cols)
    if is_wide_table(n_cols):
        st.info("📄 {} columns — Section 6 will use **A3 Landscape**".format(n_cols))
    else:
        st.info("📄 {} columns — Section 6 will use **A3 Portrait** (body remains A4 Portrait)".format(n_cols))

    cm = st.session_state.col_map
    tabs = st.tabs([
        ("🔧 " if s.get("is_maintenance") else "") + s["name"]
        for s in st.session_state.scenarios
    ])
    for idx, s in enumerate(st.session_state.scenarios):
        with tabs[idx]:
            df = s["df"]
            disp = [cm.get(k) for k in st.session_state.selected_cols
                    if cm.get(k) and cm.get(k) in df.columns]
            if not disp:
                disp = list(df.columns[:8])
            ds = df[disp].copy()
            ds.insert(0, "#", range(1, len(ds) + 1))
            excl = s.get("exclude_set", set())
            ds.insert(1, "Exclude", [i in excl for i in range(len(ds))])
            ed = st.data_editor(
                ds,
                column_config={"Exclude": st.column_config.CheckboxColumn("Exclude")},
                disabled=[c for c in ds.columns if c != "Exclude"],
                use_container_width=True, height=500,
                key="de_{}".format(idx),
            )
            if ed is not None:
                s["exclude_set"] = {i for i, r in ed.iterrows() if r.get("Exclude", False)}
                ne = len(s["exclude_set"])
                if ne:
                    st.info("{} rows excluded".format(ne))


# ════════════════════════════════════════════════════════════════════════════
# STEP 3 — Settings & Conclusion
# ════════════════════════════════════════════════════════════════════════════
elif step == "3. Settings & Conclusion":
    st.header("Step 3 — Project Info, Conclusion & Report Options")

    pi = st.session_state.project
    c1, c2 = st.columns(2)
    with c1:
        pi["project_name"]     = st.text_input("Project Name", pi.get("project_name", ""))
        pi["client_name"]      = st.text_input("Client", pi.get("client_name", ""))
        pi["project_location"] = st.text_input("Location", pi.get("project_location", ""))
        pi["document_no"]      = st.text_input("Doc No.", pi.get("document_no", ""))
        rd = st.date_input("Report Date", value=date.today())
        pi["report_date"] = rd.strftime("%d-%b-%Y")
        pi["rev_date"]    = pi["report_date"]
    with c2:
        pi["software_name"]    = st.text_input("Software", pi.get("software_name", "ETAP"))
        pi["software_version"] = st.text_input("Version", pi.get("software_version", ""))
        pi["study_standard"]   = st.text_input("Standard", pi.get("study_standard", "IEEE 1584-2018 / NFPA 70E"))
        pi["system_frequency"] = st.selectbox("Frequency", ["50 Hz", "60 Hz"])
        pi["produced_by"]      = st.text_input("Produced By", pi.get("produced_by", ""))
        pi["checked_by"]       = st.text_input("Checked By", pi.get("checked_by", ""))

    with st.expander("Revision"):
        pi["rev_no"]      = st.text_input("Rev", pi.get("rev_no", "0"))
        pi["revision"]    = "Rev. " + pi["rev_no"]
        pi["rev_remark"]  = st.text_input("Remark", pi.get("rev_remark", "Initial Issue"))
        pi["rev_status"]  = st.selectbox("Status", ["Issued", "Draft", "For Review", "Approved"])
    pi["report_type"] = "Arc Flash Hazard Analysis"

    st.markdown("---")

    # ── Report Options (Part 1.4) ──────────────────────────────────────────────
    st.subheader("📋 Report Options")
    st.session_state.show_mitigation_section = st.checkbox(
        "Include Mitigation / Maintenance Section in Report",
        value=st.session_state.show_mitigation_section,
        help=(
            "When unchecked, the recommendation text defaults to the generic PPE advisory. "
            "When checked, mitigation-specific recommendations are generated."
        ),
    )
    if not st.session_state.show_mitigation_section:
        st.info(
            "📌 Recommendation will auto-populate as: "
            "_\"Use Recommended PPE while doing any work with energized equipment.\"_"
        )
    else:
        st.info("📌 Mitigation-specific recommendation text will be generated.")

    st.markdown("---")

    # ── Conclusion ─────────────────────────────────────────────────────────────
    st.subheader("Conclusion")
    cn = st.session_state.conclusion
    cm = st.session_state.col_map

    has_maint = any(s.get("is_maintenance") for s in st.session_state.scenarios)

    if st.session_state.scenarios:
        normal_scenarios = [s for s in st.session_state.scenarios
                            if not s.get("is_maintenance", False)]
        if not normal_scenarios:
            normal_scenarios = st.session_state.scenarios

        ac = auto_conclusion(
            [summary_stats(s["df"], cm) for s in normal_scenarios],
            normal_scenarios,
            has_maint,
        )
        cn["conclusion_text"] = st.text_area(
            "Conclusion",
            cn.get("conclusion_text", "") or ac.get("conclusion_text", ""),
            height=80,
        )
        cn["observation_text"] = st.text_area(
            "Observation",
            cn.get("observation_text", "") or ac.get("observation_text", ""),
            height=60,
        )
        # Recommendation — the service layer will apply generic text if mitigation
        # section is not selected, but let the user pre-fill or override here.
        default_rec = (
            ac.get("recommendation_text", "")
            if st.session_state.show_mitigation_section
            else "Use Recommended PPE while doing any work with energized equipment."
        )
        cn["recommendation_text"] = st.text_area(
            "Recommendation (auto-populated based on Report Options above)",
            cn.get("recommendation_text", "") or default_rec,
            height=60,
        )

        if st.session_state.show_mitigation_section:
            cn["maintenance_section"] = st.text_area(
                "Maintenance Section",
                cn.get("maintenance_section", "") or ac.get("maintenance_section", ""),
                height=60,
            )
            cn["mitigation_text"] = st.text_area(
                "Mitigation",
                cn.get("mitigation_text", "") or ac.get("mitigation_text", ""),
                height=60,
            )
        else:
            cn["maintenance_section"] = ""
            cn["mitigation_text"] = ""

        cn["user_remarks"] = st.text_area("Remarks", cn.get("user_remarks", ""), height=60)

    st.markdown("---")
    st.subheader("Annexures")
    na = st.number_input("Count", 0, 20, len(st.session_state.annexures))
    while len(st.session_state.annexures) < na:
        st.session_state.annexures.append({
            "letter": chr(65 + len(st.session_state.annexures)),
            "title": "", "content": "",
        })
    st.session_state.annexures = st.session_state.annexures[:na]
    for i, a in enumerate(st.session_state.annexures):
        with st.expander("Annexure {}".format(a["letter"])):
            a["title"]   = st.text_input("Title",   a["title"],   key="at_{}".format(i))
            a["content"] = st.text_area("Content", a["content"], key="ac_{}".format(i), height=60)


# ════════════════════════════════════════════════════════════════════════════
# STEP 4 — Generate
# ════════════════════════════════════════════════════════════════════════════
elif step == "4. Generate":
    st.header("Step 4 — Generate Report")

    issues = []
    if not st.session_state.template_path:
        issues.append("Upload a report template (Step 1)")
    if not st.session_state.scenarios:
        issues.append("Add at least one scenario (Step 1)")
    if (st.session_state.comparison_mode
            and st.session_state.scenario_a == st.session_state.scenario_b):
        issues.append("Comparison Mode: Scenario A and B must be different (Step 1)")
    for issue in issues:
        st.error(issue)
    if issues:
        st.stop()

    pi = st.session_state.project

    # Summary panel
    col1, col2 = st.columns(2)
    with col1:
        st.info("**{}** | Doc: {} | Rev: {} | {}".format(
            pi.get("project_name", "—"), pi.get("document_no", "—"),
            pi.get("rev_no", "0"), pi.get("report_date", "—")))
    with col2:
        n_cols = len(st.session_state.selected_cols)
        sc_layout = "A3 Landscape" if is_wide_table(n_cols) else "A3 Portrait"
        st.info("Columns: {} | Sec.6 layout: {} | Comparison: {}".format(
            n_cols, sc_layout,
            "ON ({} vs {})".format(
                st.session_state.scenario_a, st.session_state.scenario_b)
            if st.session_state.comparison_mode else "OFF"))

    # Scenario summary
    with st.expander("📋 Scenarios in this report"):
        for s in st.session_state.scenarios:
            badge = "🔧 **[Maintenance]** " if s.get("is_maintenance") else "▶ "
            st.markdown("{}{} — {} buses".format(badge, s["name"], len(s["df"])))
        maint_sc = [s for s in st.session_state.scenarios if s.get("is_maintenance")]
        if maint_sc:
            st.caption(
                "Maintenance scenarios appear in Section 6 and can act as "
                "Scenario B in the comparison table."
            )

    st.markdown("---")

    if st.button("⚡ Generate Report", type="primary", use_container_width=True):
        with st.spinner("Generating report…"):
            try:
                od = tempfile.mkdtemp()
                op, fn = generate_report(
                    template_path=st.session_state.template_path,
                    col_map=st.session_state.col_map,
                    project=pi,
                    scenarios=st.session_state.scenarios,
                    selected_cols=st.session_state.selected_cols,
                    conclusion_overrides=st.session_state.conclusion,
                    output_dir=od,
                    # Comparison
                    comparison_mode=st.session_state.comparison_mode,
                    scenario_a_name=st.session_state.scenario_a,
                    scenario_b_name=st.session_state.scenario_b,
                    comparison_param_keys=st.session_state.comparison_params,
                    # Report options
                    show_mitigation_section=st.session_state.show_mitigation_section,
                    # Logos
                    consultant_logo=st.session_state.consultant_logo,
                    client_logo=st.session_state.client_logo,
                    # Annexures
                    annexures=st.session_state.annexures,
                )
                st.success("✅ Report generated: **{}**".format(fn))
                with open(op, "rb") as f:
                    st.download_button(
                        "⬇ Download {}".format(fn),
                        f.read(),
                        file_name=fn,
                        mime=(
                            "application/vnd.openxmlformats-officedocument"
                            ".wordprocessingml.document"
                        ),
                        use_container_width=True,
                    )
            except Exception as e:
                st.error(str(e))
                import traceback
                st.code(traceback.format_exc())
