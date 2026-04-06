"""report_generator.py — Orchestrator. FastAPI-ready.

Changes vs. previous version:
  - Accepts pre-built comparison headers/rows from comparison_service (Flexible A↔B).
  - Accepts comparison_mode, comparison_heading, comparison_description directly.
  - Maintenance mode is now per-scenario (is_maintenance flag) handled in report_service.
  - show_mitigation_section respected via conclusion_overrides.
  - Fix 9: Maintenance Mode appended as a distinct scenario in Section 6 (preserved).
  - Fix 10: No dynamic Protective Device table (preserved).
"""
import pandas as pd
import os
from .config import AF_COLS, is_wide_table, sanitize_filename
from .af_processor import detect_voltages, summary_stats, auto_conclusion, build_comparison
from .template_engine import AFTemplateEngine


class AFReportGenerator:
    def __init__(self, template_path, col_map):
        self.tmpl = template_path
        self.cm = col_map

    def generate(
        self,
        project,
        scenarios,
        selected_cols,
        conclusion_overrides,
        output_dir,
        # Legacy maintenance mode (kept for backward compat from old app.py path)
        maintenance_mode=False,
        mitigation_df=None,
        comparison_scenarios=None,   # legacy: list of scenario names
        comparison_cols=None,        # legacy: list of col keys
        # New flexible comparison (populated by report_service)
        comparison_mode=False,
        comparison_headers=None,
        comparison_rows=None,
        comparison_heading="",
        comparison_description="",
        # Report options
        annexures=None,
        consultant_logo=None,
        client_logo=None,
    ):
        engine = AFTemplateEngine(self.tmpl)
        engine.set_wide_mode(is_wide_table(len(selected_cols)))
        engine.set_logos(consultant_logo, client_logo)

        # ── Build effective scenario list (Fix 9) ──
        effective_scenarios = list(scenarios)
        if maintenance_mode and mitigation_df is not None and len(mitigation_df) > 0:
            effective_scenarios.append({
                "name": "Maintenance Mode",
                "desc": (
                    "Mitigated results with maintenance-mode protective "
                    "device settings applied (reduced clearing times)."
                ),
                "df": mitigation_df,
                "exclude_set": set(),
            })

        # Auto-detect voltages across all scenarios
        all_dfs = [s["df"] for s in effective_scenarios]
        voltages = (
            detect_voltages(pd.concat(all_dfs, ignore_index=True), self.cm)
            if all_dfs else ""
        )

        # Conclusion from original (normal) scenarios
        all_stats = [summary_stats(s["df"], self.cm) for s in scenarios]
        auto_conc = auto_conclusion(all_stats, scenarios, maintenance_mode)
        conc = {}
        conc.update(auto_conc)
        for k, v in conclusion_overrides.items():
            if v and not k.startswith("_"):   # skip internal _user_* keys
                conc[k] = v

        # ── Merge fields ──
        fields = {}
        for k, v in project.items():
            fields["{{" + k + "}}"] = v
        fields["{{voltage_levels}}"] = voltages or project.get("voltage_levels", "")
        fields["{{operating_scenarios}}"] = "; ".join(
            "{} - {}".format(s["name"], s.get("desc", "")) for s in effective_scenarios
        )
        fields["{{system_frequency}}"] = project.get("system_frequency", "50 Hz")
        fields["{{conclusion_text}}"] = conc.get("conclusion_text", "")
        fields["{{maintenance_section}}"] = conc.get("maintenance_section", "")
        fields["{{observation_text}}"] = conc.get("observation_text", "")
        fields["{{recommendation_text}}"] = conc.get("recommendation_text", "")
        fields["{{mitigation_text}}"] = conc.get("mitigation_text", "")
        fields["{{user_remarks}}"] = conc.get("user_remarks", "")

        # ── Comparison section heading/description ──
        # Priority: new flexible comparison → legacy maintenance → blank
        if comparison_mode and comparison_heading:
            fields["{{comparison_heading}}"] = comparison_heading
            fields["{{comparison_description}}"] = comparison_description
        elif maintenance_mode and comparison_scenarios:
            fields["{{comparison_heading}}"] = (
                "5.1 Maintenance Mode - Incident Energy Comparison"
            )
            fields["{{comparison_description}}"] = (
                "The following table compares Normal vs Mitigated parameters "
                "for the selected scenario(s)."
            )
        else:
            fields["{{comparison_heading}}"] = ""
            fields["{{comparison_description}}"] = ""

        engine.set_fields(fields)

        # ── Scenario tables (Section 6) ──
        col_headers = [(k, AF_COLS[k]["short"]) for k in selected_cols if k in AF_COLS]
        for s in effective_scenarios:
            df = s["df"]
            excl = s.get("exclude_set", set())
            if excl:
                df = df.drop(index=list(excl), errors="ignore").reset_index(drop=True)
            rows = self._df_to_rows(df, selected_cols)
            engine.add_scenario(s["name"], s.get("desc", ""), col_headers, rows)

        # ── Comparison table ──
        if comparison_mode and comparison_headers and comparison_rows:
            # New flexible comparison — pre-built by comparison_service
            engine.set_comparison(comparison_headers, comparison_rows)

        elif (maintenance_mode and mitigation_df is not None
              and comparison_scenarios and comparison_cols):
            # Legacy maintenance comparison path
            for sname in comparison_scenarios:
                s = next((s for s in scenarios if s["name"] == sname), None)
                if s:
                    comp_rows = build_comparison(
                        s["df"], mitigation_df, self.cm, comparison_cols
                    )
                    if comp_rows:
                        comp_hdr = [("bus_id", "Bus ID")]
                        for ck in comparison_cols:
                            lbl = AF_COLS.get(ck, {}).get("short", ck)
                            comp_hdr.append(("normal_" + ck, "Normal " + lbl))
                            comp_hdr.append(("mitigated_" + ck, "Mitigated " + lbl))
                            if AF_COLS.get(ck, {}).get("dtype") == "num":
                                comp_hdr.append(("reduction_" + ck, "Reduction %"))
                        engine.set_comparison(comp_hdr, comp_rows)
                    break

        # ── Annexures ──
        if annexures:
            for a in annexures:
                engine.add_annexure(a["letter"], a["title"], a.get("content", ""))

        # Filename
        pname = sanitize_filename(project.get("project_name", "AF_Report"))
        rev = project.get("rev_no", "0")
        rdate = project.get("report_date", "").replace("/", "-")
        fname = (
            "{}_Rev{}_{}.docx".format(pname, rev, rdate)
            if rdate else "{}_Rev{}.docx".format(pname, rev)
        )
        output_path = os.path.join(output_dir, fname)

        engine.generate(output_path)
        return output_path, fname

    def _df_to_rows(self, df, col_keys):
        """
        Convert a DataFrame to list[dict] keyed by logical column keys.
        Always includes `energy_level` so row-level color lookup works.
        """
        rows = []
        keys = list(col_keys)
        if "energy_level" not in keys:
            keys = keys + ["energy_level"]
        for idx, (_, row) in enumerate(df.iterrows()):
            rd = {"s_no": str(idx + 1)}
            for k in keys:
                c = self.cm.get(k)
                rd[k] = (
                    str(row[c])
                    if c and c in df.columns and pd.notna(row[c])
                    else ""
                )
            rows.append(rd)
        return rows


def generate_af_report(**kw):
    gen = AFReportGenerator(kw.pop("template_path"), kw.pop("col_map"))
    return gen.generate(**kw)
