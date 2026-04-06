"""af_processor.py — Excel to DataFrame. Auto column mapping in backend."""
import pandas as pd
import re
from .config import ROUND_KEYS, JUNK_PATTERNS, AF_COLS, AUTO_PATTERNS, LEVEL_ORDER, level_to_ppe

def _detect_header(fp, sheet=0, scan=10):
    raw = pd.read_excel(fp, sheet_name=sheet, header=None, nrows=scan)
    best, bn = 0, 0
    for i in range(len(raw)):
        n = sum(1 for v in raw.iloc[i] if isinstance(v, str) and len(v.strip()) > 2)
        if n > bn:
            bn, best = n, i
    return best

def read_excel(fp, sheet=0):
    hr = _detect_header(fp, sheet)
    df = pd.read_excel(fp, sheet_name=sheet, header=hr)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def auto_map(columns):
    """Auto-detect column mapping. Runs in backend only."""
    det, used = {}, set()
    for key, pats in AUTO_PATTERNS.items():
        for col in columns:
            if col in used:
                continue
            for p in pats:
                if re.search(p, col.lower().strip()):
                    det[key] = col
                    used.add(col)
                    break
            if key in det:
                break
    return det

def filter_junk(df, cm):
    bus = cm.get("bus_id", df.columns[0])
    out = df[df[bus].apply(lambda v: isinstance(v, str) and len(v.strip()) > 0)].copy()
    pat = '|'.join(JUNK_PATTERNS)
    for c in [bus] + list(out.columns[:5]):
        if c in out.columns:
            out = out[~out[c].apply(lambda v: bool(re.search(pat, str(v).strip(), re.I)) if pd.notna(v) else False)]
    kv = cm.get("kv")
    if kv and kv in out.columns:
        out = out[out[kv].apply(lambda v: _num(v) or pd.isna(v) or str(v).strip() == "")]
    return out.reset_index(drop=True)

def apply_rules(df, cm):
    df = df.copy()
    for k in ROUND_KEYS:
        c = cm.get(k)
        if c and c in df.columns:
            df[c] = df[c].apply(lambda v: round(float(v), 2) if _num(v) else v)
    el = cm.get("energy_level")
    if el and el in df.columns:
        df[el] = df[el].apply(lambda v: "FCT Not Determined" if pd.isna(v) or str(v).strip() == "" else str(v).strip())
    gc = cm.get("glove_class")
    if gc and gc in df.columns:
        df[gc] = df[gc].apply(lambda v: "00" if _num(v) and float(v) == 0 else str(int(float(v))) if _num(v) else str(v).strip() if pd.notna(v) else "")
    sp = cm.get("source_pd")
    if sp and sp in df.columns:
        df[sp] = df[sp].apply(lambda v: "N/A" if pd.isna(v) or str(v).strip() == "" else str(v).strip())
    for k in ["total_energy", "afb", "final_fct"]:
        c = cm.get(k)
        if c and c in df.columns:
            df[c] = df[c].apply(lambda v: "N/A" if pd.isna(v) or str(v).strip() == "" else v)
    return df

def process_file(fp, sheet=0):
    raw = read_excel(fp, sheet)
    cm = auto_map(list(raw.columns))
    clean = filter_junk(raw, cm)
    proc = apply_rules(clean, cm)
    return raw, proc, cm

def detect_voltages(df, cm):
    kv = cm.get("kv")
    if not kv or kv not in df.columns:
        return ""
    vals = pd.to_numeric(df[kv], errors="coerce").dropna().unique()
    vals = sorted(set(round(v, 3) for v in vals))
    parts = []
    for v in vals:
        if v < 1:
            parts.append("{:.0f}V ({} kV)".format(v * 1000, v))
        else:
            parts.append("{} kV".format(v))
    return ", ".join(parts)

def summary_stats(df, cm):
    s = {"total": len(df)}
    en = cm.get("total_energy")
    if en and en in df.columns:
        vals = pd.to_numeric(df[en], errors="coerce").dropna()
        if len(vals):
            s["max_energy"] = round(vals.max(), 2)
            bus = cm.get("bus_id")
            if bus and bus in df.columns:
                s["max_bus"] = str(df.loc[vals.idxmax(), bus])
    for k, n in [("afb", "max_afb"), ("total_ia", "max_ia"), ("total_ibf", "max_ibf")]:
        c = cm.get(k)
        if c and c in df.columns:
            v = pd.to_numeric(df[c], errors="coerce").dropna()
            s[n] = round(v.max(), 2) if len(v) else "N/A"
    el = cm.get("energy_level")
    if el and el in df.columns:
        s["levels"] = df[el].value_counts().to_dict()
        s["fct_unknown"] = int((df[el] == "FCT Not Determined").sum())
        known = [l for l in df[el].unique() if l in LEVEL_ORDER]
        if known:
            s["max_level"] = max(known, key=lambda x: LEVEL_ORDER.get(x, 0))
        s["maintenance_recommended"] = any(l in ["Level F", "Level G", "F", "G"] for l in df[el].unique())
    else:
        s["maintenance_recommended"] = False
    return s

def auto_conclusion(all_stats, scenarios, maint_mode):
    max_e = max((s.get("max_energy", 0) or 0 for s in all_stats), default=0)
    max_bus = next((s.get("max_bus", "") for s in all_stats if s.get("max_energy") == max_e), "")
    max_level = ""
    for st in all_stats:
        ml = st.get("max_level", "")
        if ml and LEVEL_ORDER.get(ml, 0) > LEVEL_ORDER.get(max_level, 0):
            max_level = ml
    ie_level, ppe_cat = level_to_ppe(max_e)
    sc_names = ", ".join(s["name"] for s in scenarios)
    c = {"max_energy": max_e, "max_bus": max_bus, "max_level": max_level}
    if LEVEL_ORDER.get(max_level, 0) <= 5:
        c["conclusion_text"] = "The arc flash study for {} indicates a maximum incident energy of {} cal/cm2 (Level {}) at bus {}. All energy levels are within Level E or below. Personnel must use appropriate PPE as per Section 4.2.".format(sc_names, max_e, ie_level, max_bus)
    else:
        c["conclusion_text"] = "The arc flash study for {} indicates a maximum incident energy of {} cal/cm2 (Level {}) at bus {}. Some buses exceed Level E. Mitigation measures are recommended.".format(sc_names, max_e, ie_level, max_bus)
    c["observation_text"] = "For buses with Incident Energy > Level E, the highest energy is {} cal/cm2 at {}.".format(max_e, max_bus)
    c["recommendation_text"] = "If protective devices are set with FCT <= 0.08s, incident energy can be reduced to Level D or below."
    if maint_mode:
        c["mitigation_text"] = "Mitigation: By providing maintenance settings in the instantaneous zone of protective devices with operating time < 0.08s, incident energy can be significantly reduced."
        c["maintenance_section"] = "Maintenance mode settings are recommended for protective devices at high-energy bus locations."
    else:
        c["mitigation_text"] = ""
        c["maintenance_section"] = ""
    return c

def build_comparison(normal_df, mitigated_df, cm, selected_cols):
    bus = cm.get("bus_id", "ID")
    norm_map = {}
    if bus in normal_df.columns:
        norm_map = {str(r[bus]).strip(): r for _, r in normal_df.iterrows()}
    rows = []
    for idx, (_, mr) in enumerate(mitigated_df.iterrows()):
        bid = str(mr.get(bus, "")).strip()
        nr = norm_map.get(bid, {})
        rd = {"s_no": str(idx + 1), "bus_id": bid}
        for key in selected_cols:
            col = cm.get(key)
            if not col:
                continue
            n_val = nr.get(col) if isinstance(nr, pd.Series) else None
            m_val = mr.get(col)
            rd["normal_" + key] = str(n_val) if pd.notna(n_val) and n_val is not None else "N/A"
            rd["mitigated_" + key] = str(m_val) if pd.notna(m_val) else "N/A"
            if AF_COLS.get(key, {}).get("dtype") == "num":
                try:
                    nf, mf = float(n_val), float(m_val)
                    rd["reduction_" + key] = str(round((nf - mf) / nf * 100, 1)) if nf > 0 else "0"
                except (ValueError, TypeError):
                    rd["reduction_" + key] = "N/A"
        rows.append(rd)
    return rows

def _num(v):
    if pd.isna(v):
        return False
    try:
        float(v)
        return True
    except (ValueError, TypeError):
        return False
