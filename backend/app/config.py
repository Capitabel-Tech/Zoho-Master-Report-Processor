from datetime import date
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

REPORT_TYPES = ["deals", "pipeline", "leads"]

LEAD_STATUS_FILTER = "new"

# Stage values (normalized lowercase) that count as "closed" for a Deals sheet.
# Every other stage in a real master export (Cold, the various Resolved-* dead-
# lead outcomes, ...) is excluded from Deals entirely.
DEALS_STAGES = {"closed won", "resolved-completed"}

# Pipeline is split into two sheets by STAGE (not probability) - the two sets
# are mutually exclusive, and every other stage (Cold, Resolved-*, Technical /
# Valuation, ...) is excluded from both.
PIPELINE_HIGH_STAGES = {"disbursement", "sanction", "login"}
PIPELINE_LOW_STAGES = {"document collection", "identify bank", "customer visit / telecommunication"}

QUARTER_ORDER = ["Q1", "Q2", "Q3", "Q4"]

BASE_DIR = Path(__file__).resolve().parent.parent
TEMPLATES_DIR = BASE_DIR / "data" / "templates"


def template_path(report_type: str) -> Path:
    return TEMPLATES_DIR / f"{report_type}.xlsx"


def fiscal_year_start_year(today: date) -> int:
    """The calendar year Q1 (April) starts in, for the fiscal year containing `today`.

    Jan-Mar belongs to the fiscal year that started the *previous* April.
    """
    return today.year if today.month >= 4 else today.year - 1


def quarter_date_range(quarter: str, fy_start_year: int) -> tuple[date, date]:
    """(start, end) calendar dates for `quarter` within the fiscal year that
    starts in April of `fy_start_year`. Q4 spans into fy_start_year + 1."""
    ranges = {
        "Q1": (date(fy_start_year, 4, 1), date(fy_start_year, 6, 30)),
        "Q2": (date(fy_start_year, 7, 1), date(fy_start_year, 9, 30)),
        "Q3": (date(fy_start_year, 10, 1), date(fy_start_year, 12, 31)),
        "Q4": (date(fy_start_year + 1, 1, 1), date(fy_start_year + 1, 3, 31)),
    }
    return ranges[quarter]


def fiscal_year_range(fy_start_year: int) -> tuple[date, date]:
    """(start, end) calendar dates spanning the WHOLE fiscal year that starts
    in April of `fy_start_year` - used for Pipeline, which (unlike Deals) isn't
    scoped to a single quarter's months."""
    return (date(fy_start_year, 4, 1), date(fy_start_year + 1, 3, 31))


def fiscal_year_label(fy_start_year: int) -> str:
    """e.g. 2026 -> "2026 - 2027" - for titles like "New Leads FY - 2026 - 2027"."""
    return f"{fy_start_year} - {fy_start_year + 1}"


def current_and_complete_quarters(today: date) -> tuple[str, list[str]]:
    """Returns (current_quarter, [quarters of this fiscal year already fully
    elapsed, in order starting from Q1])."""
    current = next(q for q in QUARTER_ORDER if today.month in QUARTER_MONTHS[q])
    idx = QUARTER_ORDER.index(current)
    return current, QUARTER_ORDER[:idx]
