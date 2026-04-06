"""AF Report Platform"""
from .config import AF_COLS, LANDSCAPE_COL_THRESHOLD, is_wide_table, LEVEL_COLORS
from .af_processor import (
    read_excel, auto_map, filter_junk, apply_rules,
    summary_stats, auto_conclusion, build_comparison, detect_voltages,
    process_file
)
from .template_engine import AFTemplateEngine
from .report_generator import AFReportGenerator, generate_af_report
