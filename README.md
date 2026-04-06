# Arc Flash Study Report Generator

## Quick Start

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Workflow

1. **Upload Template & Logos** — Upload the `.docx` merge-field template and optional company logos
2. **Add Scenarios** — Upload Excel result files (columns auto-mapped, no user config needed)
3. **Preview & Edit** — Select report columns, exclude rows via checkbox
4. **Project Info & Conclusion** — Fill project details (date picker included), conclusion auto-derived
5. **Generate** — Download as `ProjectName_RevX_DD-MMM-YYYY.docx`

## Page Size Logic

| Condition | Page Size |
|---|---|
| User selects ≤ 6 columns | Template default (A4 Portrait) |
| User selects > 6 columns | Table section switches to A3 Landscape, then back to template size |

## Result Table Heading Customization

The result table column headers are fully controlled by **user column selection** in Step 2.

### How to customize table headings:

1. Open `modules/config.py`
2. Find the `AF_COLS` dictionary
3. Each column has a `"short"` key — this is the table header text

**Example — change "Energy" to "Inc. Energy (cal/cm²)":**

```python
# In AF_COLS:
"total_energy": {
    "label": "Incident Energy (cal/cm2)",  # UI label
    "short": "Inc. Energy (cal/cm2)",       # <-- TABLE HEADER TEXT
    ...
},
```

### To add a completely new column:

```python
"new_column": {
    "label": "My Custom Column",     # Shown in UI
    "short": "Custom",               # Table header
    "req": False,                    # Required?
    "dtype": "num",                  # "num" or "text"
    "round": True,                   # Round to 2 decimals?
    "compare": False,                # Available for comparison?
},
```

Then add the auto-detect pattern in `AUTO_PATTERNS`:

```python
"new_column": [r"my.*custom", r"^custom$"],
```

## Level Color Coding

Result tables automatically color-code rows by Energy Level:

| Level | Color | Hex |
|---|---|---|
| A, B | Green | `#E2EFDA` |
| C | Blue | `#D6E4F0` |
| D | Yellow | `#FFF3CD` |
| E, F, G | Red | `#FFCDD2` |

To modify colors, edit `LEVEL_COLORS` in `modules/config.py`.

## File Structure

```
af_deploy/
├── app.py                   # Streamlit UI
├── AF_Template.docx         # Merge-field template
├── requirements.txt
├── README.md
└── modules/
    ├── config.py            # Colors, columns, thresholds
    ├── af_processor.py      # Excel reader, auto-mapping, rules
    ├── template_engine.py   # Template-driven report builder
    └── report_generator.py  # Orchestrator (FastAPI-ready)
```

## FastAPI Migration

`report_generator.py` exposes `generate_af_report(**kwargs)` — call directly from a FastAPI endpoint.
