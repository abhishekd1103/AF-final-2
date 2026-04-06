"""
report_service.py — FastAPI-ready report generation service.

Thin orchestration wrapper around AFReportGenerator.
This module can be imported by FastAPI routes without any Streamlit dependency.

Future usage:
    POST /generate-report
    Body: ReportRequest (JSON)
    Response: binary .docx or {"filename": ..., "download_url": ...}

Current usage:
    Called directly from app.py (Streamlit) via generate_report().
"""
import os
import sys
import tempfile

# Ensure modules are importable when called from different working dirs
_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_HERE)
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from modules.report_generator import AFReportGenerator
from services.recommendation_service import get_recommendation_text, get_mitigation_text
from services.comparison_service import compare_scenarios, COMPARISON_PARAMS


def generate_report(
    template_path: str,
    col_map: dict,
    project: dict,
    scenarios: list,
    selected_cols: list,
    conclusion_overrides: dict,
    output_dir: str,
    # Comparison
    comparison_mode: bool = False,
    scenario_a_name: str = "",
    scenario_b_name: str = "",
    comparison_param_keys: list = None,
    # Mitigation / report options
    show_mitigation_section: bool = False,
    # Logos
    consultant_logo: str = None,
    client_logo: str = None,
    # Annexures
    annexures: list = None,
) -> tuple:
    """
    Generate an Arc Flash Report and return (output_path, filename).

    This function is the single entry point for report generation.
    All business logic lives here; the UI (Streamlit) and the API (FastAPI)
    call this function with the same signature.
    """
    if comparison_param_keys is None:
        comparison_param_keys = ["total_energy", "afb"]
    if annexures is None:
        annexures = []

    # ── Resolve recommendation text before passing to generator ──
    auto_rec = conclusion_overrides.get("recommendation_text", "")
    auto_mit = conclusion_overrides.get("mitigation_text", "")

    final_overrides = dict(conclusion_overrides)
    final_overrides["recommendation_text"] = get_recommendation_text(
        show_mitigation_section=show_mitigation_section,
        auto_text=auto_rec,
        user_override=conclusion_overrides.get("_user_recommendation", ""),
    )
    final_overrides["mitigation_text"] = get_mitigation_text(
        show_mitigation_section=show_mitigation_section,
        auto_text=auto_mit,
        user_override=conclusion_overrides.get("_user_mitigation", ""),
    )
    if not show_mitigation_section:
        final_overrides["maintenance_section"] = ""

    # ── Identify maintenance-mode scenario (marked per-scenario) ──
    # A scenario with is_maintenance=True acts as the "mitigated" side.
    maintenance_scenario = next(
        (s for s in scenarios if s.get("is_maintenance", False)), None
    )
    normal_scenarios = [s for s in scenarios if not s.get("is_maintenance", False)]
    maintenance_mode = maintenance_scenario is not None
    mitigation_df = maintenance_scenario["df"] if maintenance_scenario else None

    # ── Build comparison if requested ──
    comp_headers = []
    comp_rows = []
    if comparison_mode and scenario_a_name and scenario_b_name:
        sc_a = next((s for s in scenarios if s["name"] == scenario_a_name), None)
        sc_b = next((s for s in scenarios if s["name"] == scenario_b_name), None)
        if sc_a and sc_b:
            result = compare_scenarios(
                df_a=sc_a["df"],
                df_b=sc_b["df"],
                col_map=col_map,
                param_keys=comparison_param_keys,
                scenario_a_name=scenario_a_name,
                scenario_b_name=scenario_b_name,
            )
            comp_headers, comp_rows = result.to_table()

    # ── Delegate to existing generator (unchanged calc engine) ──
    gen = AFReportGenerator(template_path, col_map)
    return gen.generate(
        project=project,
        scenarios=normal_scenarios if maintenance_scenario else scenarios,
        selected_cols=selected_cols,
        conclusion_overrides=final_overrides,
        output_dir=output_dir,
        maintenance_mode=maintenance_mode,
        mitigation_df=mitigation_df,
        # Pass pre-built comparison instead of letting generator rebuild it
        comparison_mode=comparison_mode,
        comparison_headers=comp_headers,
        comparison_rows=comp_rows,
        comparison_heading=(
            f"5.1 Scenario Comparison: {scenario_a_name} vs {scenario_b_name}"
            if comparison_mode else ""
        ),
        comparison_description=(
            "The following table compares key parameters between the two selected scenarios."
            if comparison_mode else ""
        ),
        annexures=annexures,
        consultant_logo=consultant_logo,
        client_logo=client_logo,
    )
