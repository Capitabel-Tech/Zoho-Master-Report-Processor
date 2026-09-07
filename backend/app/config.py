from pathlib import Path

QUARTER_LABELS = {
    "Q1": "Q1 (Apr / May / Jun)",
    "Q2": "Q2 (Jul / Aug / Sep)",
    "Q3": "Q3 (Oct / Nov / Dec)",
    "Q4": "Q4 (Jan / Feb / Mar)",
}
QUARTERS = list(QUARTER_LABELS.keys())

QUARTER_MONTHS = {
    "Q1": (4, 5, 6),
    "Q2": (7, 8, 9),
    "Q3": (10, 11, 12),
    "Q4": (1, 2, 3),
}

# Pipeline isn't quarter-scoped the way Deals is - it only ever tracks
# whichever quarter's deals are currently in progress, and comes in exactly
# two variants of that same data: the full pipeline, and a subset filtered to
# probability < 40%. There's no month range behind these (unlike QUARTER_MONTHS),
# so the quarter-mismatch check in excel_engine simply doesn't apply here.
PIPELINE_PERIODS = {
    "FULL": "Q2",
    "LT40": "Q2 - <40%",
}

# What "period" options are valid for each report type, and what label each
# one substitutes into the template's {QUARTER} placeholder.
PERIOD_OPTIONS = {
    "deals": QUARTER_LABELS,
    "pipeline": PIPELINE_PERIODS,
}

REPORT_TYPES = ["deals", "pipeline"]

BASE_DIR = Path(__file__).resolve().parent.parent
TEMPLATES_DIR = BASE_DIR / "data" / "templates"


def template_path(report_type: str) -> Path:
    return TEMPLATES_DIR / f"{report_type}.xlsx"
