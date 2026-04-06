"""
comparison_service.py — Flexible Scenario Comparison Engine.

Architecture: Designed for 2-scenario comparison today, extensible to N scenarios.
Future FastAPI exposure: POST /compare-scenarios

Design decisions:
  - Any two scenarios (A ↔ B) can be compared — not just maintenance vs normal.
  - If a bus exists in Scenario X but NOT in Scenario Y, display "OFF".
  - Supported parameters are defined in COMPARISON_PARAMS (add new ones here).
  - A ComparisonResult dataclass lets the report layer stay decoupled from pandas.
"""
from __future__ import annotations

import pandas as pd
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

# ─── Supported comparison parameters ─────────────────────────────────────────
# To add a new parameter in the future: add one entry here. No other code change.
# Key → (label, dtype)  dtype: "num" | "text"
COMPARISON_PARAMS: Dict[str, Tuple[str, str]] = {
    "total_energy": ("Incident Energy (cal/cm²)", "num"),
    "afb":          ("Arc Flash Boundary (m)",     "num"),
    "energy_level": ("Energy Level",               "text"),
    "final_fct":    ("Clearing Time (s)",           "num"),
}

ABSENT_MARKER = "OFF"   # shown when a bus is missing from one scenario


# ─── Data model ──────────────────────────────────────────────────────────────
@dataclass
class ComparisonRow:
    s_no: str
    bus_id: str
    values_a: Dict[str, str] = field(default_factory=dict)  # {param_key: value}
    values_b: Dict[str, str] = field(default_factory=dict)
    reductions: Dict[str, str] = field(default_factory=dict)  # numeric params only


@dataclass
class ComparisonResult:
    scenario_a_name: str
    scenario_b_name: str
    param_keys: List[str]
    rows: List[ComparisonRow]

    # ── Convenience: convert to the (headers, rows) format the template engine expects ──
    def to_table(self) -> Tuple[List[Tuple[str, str]], List[Dict[str, str]]]:
        """Return (col_headers, row_dicts) compatible with AFTemplateEngine.set_comparison()."""
        headers: List[Tuple[str, str]] = [("bus_id", "Bus ID")]
        for pk in self.param_keys:
            label, dtype = COMPARISON_PARAMS.get(pk, (pk, "text"))
            short = label.split("(")[0].strip()
            headers.append((f"a_{pk}", f"{self.scenario_a_name}\n{short}"))
            headers.append((f"b_{pk}", f"{self.scenario_b_name}\n{short}"))
            if dtype == "num":
                headers.append((f"reduction_{pk}", "Reduction %"))

        rows: List[Dict[str, str]] = []
        for cr in self.rows:
            rd: Dict[str, str] = {"s_no": cr.s_no, "bus_id": cr.bus_id}
            for pk in self.param_keys:
                rd[f"a_{pk}"] = cr.values_a.get(pk, ABSENT_MARKER)
                rd[f"b_{pk}"] = cr.values_b.get(pk, ABSENT_MARKER)
                _, dtype = COMPARISON_PARAMS.get(pk, (pk, "text"))
                if dtype == "num":
                    rd[f"reduction_{pk}"] = cr.reductions.get(pk, "—")
            rows.append(rd)
        return headers, rows


# ─── Core comparison function ─────────────────────────────────────────────────
def compare_scenarios(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    col_map: Dict[str, str],
    param_keys: List[str],
    scenario_a_name: str = "Scenario A",
    scenario_b_name: str = "Scenario B",
) -> ComparisonResult:
    """
    Compare two scenario DataFrames on the given parameter keys.

    Rules:
    - Master bus list = union of both scenarios, iterated in df_a order then extra from df_b.
    - If bus in A but not B → b values = ABSENT_MARKER ("OFF").
    - If bus in B but not A → a values = ABSENT_MARKER ("OFF").
    - Reduction (%) computed only for numeric params and only when both values are present.

    Returns a ComparisonResult. Call .to_table() to get the template-engine format.
    """
    bus_col = col_map.get("bus_id", "")

    # Build lookup dicts: bus_id → row Series
    def _build_lookup(df: pd.DataFrame) -> Dict[str, pd.Series]:
        if not bus_col or bus_col not in df.columns:
            return {}
        return {str(r[bus_col]).strip(): r for _, r in df.iterrows()}

    lookup_a = _build_lookup(df_a)
    lookup_b = _build_lookup(df_b)

    # Union of bus IDs — preserve A order, then extra from B
    seen: set = set()
    bus_order: List[str] = []
    for bid in lookup_a:
        if bid not in seen:
            bus_order.append(bid)
            seen.add(bid)
    for bid in lookup_b:
        if bid not in seen:
            bus_order.append(bid)
            seen.add(bid)

    # Validate requested params against COMPARISON_PARAMS
    valid_keys = [k for k in param_keys if k in COMPARISON_PARAMS]

    rows: List[ComparisonRow] = []
    for idx, bus_id in enumerate(bus_order):
        row_a = lookup_a.get(bus_id)
        row_b = lookup_b.get(bus_id)

        va: Dict[str, str] = {}
        vb: Dict[str, str] = {}
        reductions: Dict[str, str] = {}

        for pk in valid_keys:
            col = col_map.get(pk)
            _, dtype = COMPARISON_PARAMS[pk]

            # Value from Scenario A
            if row_a is not None and col and col in df_a.columns:
                raw = row_a.get(col)
                va[pk] = str(raw) if pd.notna(raw) and raw != "" else "N/A"
            else:
                va[pk] = ABSENT_MARKER

            # Value from Scenario B
            if row_b is not None and col and col in df_b.columns:
                raw = row_b.get(col)
                vb[pk] = str(raw) if pd.notna(raw) and raw != "" else "N/A"
            else:
                vb[pk] = ABSENT_MARKER

            # Reduction % (numeric only, both present)
            if dtype == "num" and va[pk] not in (ABSENT_MARKER, "N/A") and vb[pk] not in (ABSENT_MARKER, "N/A"):
                try:
                    fa, fb = float(va[pk]), float(vb[pk])
                    if fa != 0:
                        reductions[pk] = str(round((fa - fb) / fa * 100, 1))
                    else:
                        reductions[pk] = "0"
                except (ValueError, TypeError):
                    reductions[pk] = "N/A"
            else:
                reductions[pk] = "—"

        rows.append(ComparisonRow(
            s_no=str(idx + 1),
            bus_id=bus_id,
            values_a=va,
            values_b=vb,
            reductions=reductions,
        ))

    return ComparisonResult(
        scenario_a_name=scenario_a_name,
        scenario_b_name=scenario_b_name,
        param_keys=valid_keys,
        rows=rows,
    )
