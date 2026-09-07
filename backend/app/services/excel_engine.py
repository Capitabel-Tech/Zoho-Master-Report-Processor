"""
Template-driven formula engine.

A quarter's "master template" is an .xlsx with:
  - an optional title row
  - a header row (must contain a cell that reads "Deal Name")
  - exactly one example data row directly below the header

Each template column is classified by looking at its example-row cell:
  - starts with "="  -> FORMULA column: reused for every output row, with the
    row number in its cell references shifted to match.
  - anything else     -> PASSTHROUGH column: for each output row, the value is
    looked up from the uploaded raw file by matching header text (not position).

This lets the client add a new column just by editing the template in Excel -
no code changes needed.
"""
import copy
import io
import re
import textwrap
from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime
from typing import Optional

from openpyxl.styles import Font, PatternFill
from openpyxl import Workbook, load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


class EngineError(Exception):
    """Raised for problems the UI should show directly to the user."""


def normalize_header(value) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().lower()


def _find_header_row(ws: Worksheet, anchor_header: str = "deal name") -> int:
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 10)):
        for cell in row:
            if normalize_header(cell.value) == anchor_header:
                return cell.row
    raise EngineError(
        f'Could not find a header row containing "{anchor_header.title()}" in the first 10 rows.'
    )


CELL_REF_RE = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)")

# Catches the #1 mistake when hand-editing a template: writing a calculation
# like "E3+F3" without the leading "=" that tells Excel it's a formula.
# Without "=" it's just text, and the column silently comes out blank.
LOOKS_LIKE_MISSING_EQUALS_RE = re.compile(
    r"^\(?\s*\$?[A-Za-z]{1,3}\$?\d+\s*[-+*/]\s*\$?[A-Za-z]{1,3}\$?\d+"
)


MAX_WRAPPED_LINES = 6

# Excel sometimes drops a column's explicit width when a file is re-saved
# (observed after hand-editing in Excel). This is the fallback used BOTH when
# estimating how many lines a cell's text will wrap to AND when actually
# setting that column's width in the output - the same file that generated
# the estimate. Any mismatch between those two is what causes clipped text.
DEFAULT_COL_WIDTH = 24


def wrapped_line_count(text: str, col_width: float) -> int:
    if not text:
        return 1
    wrapped = textwrap.wrap(str(text), width=max(int(col_width or DEFAULT_COL_WIDTH), 1)) or [""]
    return min(len(wrapped), MAX_WRAPPED_LINES)


def shift_formula_row(formula: str, from_row: int, to_row: int) -> str:
    """Shift bare (non-$-row) references to `from_row` onto `to_row`.

    Absolute row references (e.g. $E$3) are left untouched, since those are
    presumed intentional fixed references.
    """

    def repl(m: re.Match) -> str:
        col_dollar, col, row_dollar, row_digits = m.groups()
        if row_dollar == "$" or int(row_digits) != from_row:
            return m.group(0)
        return f"{col_dollar}{col}{row_dollar}{to_row}"

    return CELL_REF_RE.sub(repl, formula)


@dataclass
class TemplateColumn:
    index: int  # 1-based column index
    header: str
    header_norm: str
    is_formula: bool
    formula: Optional[str]
    number_format: str


@dataclass
class ParsedTemplate:
    title_row: Optional[int]
    header_row: int
    example_row: int
    columns: list[TemplateColumn]
    column_widths: dict[str, float]
    wb_path_for_style: str  # kept for re-opening to copy styles when building output


def read_template(path, anchor_header: str = "deal name") -> ParsedTemplate:
    wb = load_workbook(path, data_only=False)
    ws = wb.active

    header_row = _find_header_row(ws, anchor_header)
    title_row = header_row - 1 if header_row > 1 else None
    if title_row is not None and all(
        c.value is None for c in ws[title_row]
    ):
        title_row = None
    example_row = header_row + 1

    columns: list[TemplateColumn] = []
    max_col = ws.max_column
    for col_idx in range(1, max_col + 1):
        header_cell = ws.cell(header_row, col_idx)
        if header_cell.value is None or str(header_cell.value).strip() == "":
            continue
        example_cell = ws.cell(example_row, col_idx)
        value = example_cell.value
        is_formula = isinstance(value, str) and value.startswith("=")

        if (
            isinstance(value, str)
            and not is_formula
            and LOOKS_LIKE_MISSING_EQUALS_RE.match(value.strip())
        ):
            raise EngineError(
                f'Column "{header_cell.value}" has "{value}" in row {example_row}, which '
                f'looks like a formula missing its "=" at the start. Did you mean "={value}"? '
                f"Without the \"=\", Excel treats it as plain text and the column comes out blank."
            )

        columns.append(
            TemplateColumn(
                index=col_idx,
                header=str(header_cell.value),
                header_norm=normalize_header(header_cell.value),
                is_formula=is_formula,
                formula=value if is_formula else None,
                number_format=example_cell.number_format,
            )
        )

    if not columns:
        raise EngineError("Template has no columns after the header row.")

    widths = {}
    for col_idx in range(1, max_col + 1):
        letter = get_column_letter(col_idx)
        dim = ws.column_dimensions.get(letter)
        if dim and dim.width:
            widths[letter] = dim.width

    return ParsedTemplate(
        title_row=title_row,
        header_row=header_row,
        example_row=example_row,
        columns=columns,
        column_widths=widths,
        wb_path_for_style=str(path),
    )


MONTH_NAMES = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec",
}


def check_quarter_match(
    raw_rows: list[dict[str, object]],
    expected_months: tuple[int, ...],
    quarter_months_map: dict[str, tuple[int, ...]],
) -> Optional[str]:
    """Returns a warning message if most dated rows fall outside the selected
    quarter's months, or None if there's nothing to flag (including when the
    raw file has no recognizable date column at all - e.g. Pipeline deals,
    which don't have a closing date yet).
    """
    date_header = next((k for k in raw_rows[0] if "date" in k), None)
    if not date_header:
        return None

    months = [
        v.month
        for row in raw_rows
        if isinstance((v := row.get(date_header)), (date, datetime))
    ]
    if not months:
        return None

    mismatched = sum(1 for m in months if m not in expected_months)
    if mismatched <= len(months) / 2:
        return None

    actual_month, _ = Counter(months).most_common(1)[0]
    likely_quarter = next(
        (q for q, months_ in quarter_months_map.items() if actual_month in months_),
        None,
    )
    expected_names = "/".join(MONTH_NAMES[m] for m in expected_months)
    if likely_quarter:
        return (
            f"Most of this file's dates are from {MONTH_NAMES[actual_month]}, which falls in "
            f"{likely_quarter}, not the selected quarter ({expected_names}). Did you mean to "
            f"pick {likely_quarter} instead?"
        )
    return (
        f"Most of this file's dates ({MONTH_NAMES[actual_month]}) don't fall in the selected "
        f"quarter's months ({expected_names}). Double-check you picked the right quarter."
    )


def read_raw_upload(
    file_bytes: bytes, anchor_header: str = "deal name"
) -> list[dict[str, object]]:
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active

    header_row = _find_header_row(ws, anchor_header)
    headers: dict[int, str] = {}
    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(header_row, col_idx)
        if cell.value is not None and str(cell.value).strip() != "":
            headers[col_idx] = normalize_header(cell.value)

    rows: list[dict[str, object]] = []
    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row):
        if all(c.value is None for c in row):
            continue
        # stop at the first fully-blank row (raw exports may have trailing pivot data)
        row_data = {}
        for col_idx, header_norm in headers.items():
            row_data[header_norm] = row[col_idx - 1].value
        # skip rows that look like a pivot / summary block (e.g. "Grand Total")
        if row_data.get(anchor_header) is None:
            continue
        rows.append(row_data)

    if not rows:
        raise EngineError("No data rows found in the uploaded file.")
    return rows


def write_report_sheet(
    out_ws: Worksheet,
    template: ParsedTemplate,
    raw_rows: list[dict[str, object]],
    quarter_label: Optional[str] = None,
) -> None:
    """Writes one template-driven report into `out_ws` (an existing, already-titled
    worksheet). `raw_rows` may be empty - that just produces a title/header-only sheet,
    e.g. a quarter with no matching deals yet.
    """
    src_wb = load_workbook(template.wb_path_for_style, data_only=False)
    src_ws = src_wb.active

    # A raw export's own pre-existing "Total"/"Grand Total" row (Zoho's Deals
    # export includes one) is noise here - the Total row this function adds at
    # the end is always freshly computed from the actual rows in *this*
    # output, so any raw one is dropped rather than treated as a real deal.
    label_col = template.columns[0]
    raw_rows = [
        r
        for r in raw_rows
        if str(r.get(label_col.header_norm) or "").strip().lower()
        not in ("total", "grand total")
    ]

    # Every passthrough column (Status, Lead Source, Gross Revenue, ...) is
    # copied best-effort - if the raw file doesn't have it at all, every row
    # just gets None for it, same as if the raw file had the column but left
    # it blank for that row. Nothing here blocks the whole report: a formula
    # that depends on a missing/blank column simply comes out blank for that
    # row too (below), rather than the file being rejected outright.
    col_by_index = {c.index: c for c in template.columns}

    # Per formula column, which passthrough columns' raw values it needs.
    formula_dependencies: dict[int, set[str]] = {}
    for col in template.columns:
        if not col.is_formula:
            continue
        deps: set[str] = set()
        for m in CELL_REF_RE.finditer(col.formula):
            _, letters, _, _ = m.groups()
            try:
                dep_idx = column_index_from_string(letters.upper())
            except ValueError:
                continue
            dep_col = col_by_index.get(dep_idx)
            if dep_col and not dep_col.is_formula:
                deps.add(dep_col.header_norm)
        formula_dependencies[col.index] = deps

    def copy_style(src_cell, dst_cell):
        dst_cell.font = copy.copy(src_cell.font)
        dst_cell.fill = copy.copy(src_cell.fill)
        dst_cell.border = copy.copy(src_cell.border)
        dst_cell.alignment = copy.copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format

    def copy_row_height(src_row_idx: int, dst_row_idx: int):
        src_dim = src_ws.row_dimensions.get(src_row_idx)
        if src_dim and src_dim.height:
            out_ws.row_dimensions[dst_row_idx].height = src_dim.height

    # title row
    if template.title_row is not None:
        for col in template.columns:
            src_cell = src_ws.cell(template.title_row, col.index)
            dst_cell = out_ws.cell(template.title_row, col.index)
            value = src_cell.value
            if quarter_label and isinstance(value, str) and "{QUARTER}" in value:
                value = value.replace("{QUARTER}", quarter_label)
            dst_cell.value = value
            copy_style(src_cell, dst_cell)
        copy_row_height(template.title_row, template.title_row)
        if src_ws.merged_cells:
            for merged_range in src_ws.merged_cells.ranges:
                if merged_range.min_row == template.title_row == merged_range.max_row:
                    out_ws.merge_cells(str(merged_range))

    # header row - height is computed from the actual header text (not copied
    # verbatim from the template), so a newly added long column name is never
    # visually clipped even though the template's own fixed height wouldn't fit it.
    HEADER_LINE_HEIGHT = 15
    header_max_lines = 1
    for col in template.columns:
        src_cell = src_ws.cell(template.header_row, col.index)
        dst_cell = out_ws.cell(template.header_row, col.index)
        dst_cell.value = col.header
        copy_style(src_cell, dst_cell)
        col_width = template.column_widths.get(get_column_letter(col.index))
        header_max_lines = max(header_max_lines, wrapped_line_count(col.header, col_width))
    template_header_height = src_ws.row_dimensions.get(template.header_row)
    min_header_height = (
        template_header_height.height
        if template_header_height and template_header_height.height
        else 24
    )
    out_ws.row_dimensions[template.header_row].height = max(
        min_header_height, header_max_lines * HEADER_LINE_HEIGHT + 14
    )

    # column widths - every column gets an explicit width (falling back to the
    # same DEFAULT_COL_WIDTH used to estimate wrap line counts above), so a
    # column that lost its width metadata never ends up narrower in the
    # actual output than what the height calculation assumed it would be.
    for col in template.columns:
        letter = get_column_letter(col.index)
        out_ws.column_dimensions[letter].width = template.column_widths.get(
            letter, DEFAULT_COL_WIDTH
        )

    # data rows - extra padding beyond the wrapped text height gives each row
    # visible breathing room, since Excel has no true "gap between rows" like
    # a webpage does.
    first_data_row = template.header_row + 1
    example_row_height = src_ws.row_dimensions.get(template.example_row)
    min_row_height = example_row_height.height if example_row_height and example_row_height.height else 26
    LINE_HEIGHT = 15
    ROW_PADDING = 14

    STRIPE_FILL = PatternFill(start_color="F2F6FB", end_color="F2F6FB", fill_type="solid")
    TOTAL_FILL = PatternFill(start_color="D9E2EC", end_color="D9E2EC", fill_type="solid")

    def bold_copy(font):
        return Font(name=font.name, size=font.size, bold=True, color=font.color)

    passthrough_values: dict[int, list[object]] = {c.index: [] for c in template.columns}

    for i, raw_row in enumerate(raw_rows):
        out_row = first_data_row + i
        max_lines = 1
        for col in template.columns:
            example_cell = src_ws.cell(template.example_row, col.index)
            dst_cell = out_ws.cell(out_row, col.index)
            if col.is_formula:
                deps = formula_dependencies.get(col.index, set())
                inputs_present = all(raw_row.get(dep) is not None for dep in deps)
                value = (
                    shift_formula_row(col.formula, template.example_row, out_row)
                    if inputs_present
                    else None
                )
            else:
                value = raw_row.get(col.header_norm)
                passthrough_values[col.index].append(value)
            dst_cell.value = value
            copy_style(example_cell, dst_cell)

            if isinstance(value, str) and not value.startswith("="):
                col_width = template.column_widths.get(get_column_letter(col.index))
                max_lines = max(max_lines, wrapped_line_count(value, col_width))

        if i % 2 == 1:
            for col in template.columns:
                out_ws.cell(out_row, col.index).fill = STRIPE_FILL

        out_ws.row_dimensions[out_row].height = max(min_row_height, max_lines * LINE_HEIGHT + ROW_PADDING)

    # Total row - auto-generated fresh from whatever's actually in this sheet,
    # rather than relying on a raw export's own (possibly missing, possibly
    # differently-shaped) total. A column is "summable" if it's a passthrough
    # column whose every non-blank value is numeric, or a formula column built
    # only from +/- (a sum of differences equals the difference of sums, so
    # summing the column is valid). A formula column involving division (a
    # ratio/percentage) instead gets that same formula reapplied at the total
    # row, referencing the just-computed totals in that row - never summed
    # directly, since summing percentages isn't meaningful. If nothing in the
    # sheet is summable, no Total row is added at all (e.g. an all-text sheet
    # like New Leads).
    if raw_rows:
        last_data_row = first_data_row + len(raw_rows) - 1
        total_row = last_data_row + 1

        def is_numeric_column(col_idx: int) -> bool:
            values = passthrough_values.get(col_idx, [])
            present = [v for v in values if v is not None]
            return bool(present) and all(isinstance(v, (int, float)) for v in present)

        ratio_cols = []
        summable_cols = []
        for col in template.columns[1:]:
            if col.is_formula:
                if "/" in col.formula:
                    ratio_cols.append(col)
                else:
                    summable_cols.append(col)
            # A passthrough percentage (e.g. "Probability (%)") isn't a plain
            # amount - summing percentages across rows isn't meaningful, so
            # it's left blank in the Total row rather than summed like Requested/
            # Sanctioned/Disbursed amounts are. Checked via the header text too,
            # since a percentage stored as a whole number (50, not 0.5) may not
            # use Excel's own "%" number format at all.
            elif (
                "%" not in col.number_format
                and "%" not in col.header
                and is_numeric_column(col.index)
            ):
                summable_cols.append(col)

        if summable_cols or ratio_cols:
            totaled_indices = {col.index for col in summable_cols}

            def ratio_inputs_totaled(formula: str) -> bool:
                for m in CELL_REF_RE.finditer(formula):
                    _, letters, _, _ = m.groups()
                    try:
                        ref_idx = column_index_from_string(letters.upper())
                    except ValueError:
                        continue
                    if ref_idx != label_col.index and ref_idx not in totaled_indices:
                        return False
                return True

            label_cell = out_ws.cell(total_row, label_col.index, "Total")
            copy_style(src_ws.cell(template.example_row, label_col.index), label_cell)

            for col in summable_cols:
                letter = get_column_letter(col.index)
                cell = out_ws.cell(total_row, col.index, f"=SUM({letter}{first_data_row}:{letter}{last_data_row})")
                copy_style(src_ws.cell(template.example_row, col.index), cell)

            # A ratio's total only makes sense if every amount it divides by
            # was itself actually totaled - otherwise (e.g. Login Amount has
            # no total because no deal in this sheet has reached Login yet)
            # it would silently show a misleading 0% rather than being blank.
            for col in ratio_cols:
                if not ratio_inputs_totaled(col.formula):
                    continue
                cell = out_ws.cell(
                    total_row, col.index, shift_formula_row(col.formula, template.example_row, total_row)
                )
                copy_style(src_ws.cell(template.example_row, col.index), cell)

            for col in template.columns:
                cell = out_ws.cell(total_row, col.index)
                cell.font = bold_copy(cell.font)
                cell.fill = TOTAL_FILL

            out_ws.row_dimensions[total_row].height = min_row_height


def build_output(
    template: ParsedTemplate,
    raw_rows: list[dict[str, object]],
    quarter_label: Optional[str] = None,
) -> bytes:
    out_wb = Workbook()
    out_ws = out_wb.active
    src_wb = load_workbook(template.wb_path_for_style, data_only=False)
    out_ws.title = src_wb.active.title or "Report"

    write_report_sheet(out_ws, template, raw_rows, quarter_label)

    # openpyxl writes formula text but no cached result; force Excel (or any
    # compliant reader) to fully recalculate the moment the file is opened,
    # instead of relying on a cached value that was never written.
    out_wb.calculation.fullCalcOnLoad = True

    buffer = io.BytesIO()
    out_wb.save(buffer)
    buffer.seek(0)
    return buffer.read()


def process_quarter(
    template_path,
    raw_file_bytes: bytes,
    quarter_label: Optional[str] = None,
    expected_months: Optional[tuple[int, ...]] = None,
    quarter_months_map: Optional[dict[str, tuple[int, ...]]] = None,
) -> bytes:
    template = read_template(template_path)
    raw_rows = read_raw_upload(raw_file_bytes)

    if expected_months and quarter_months_map:
        mismatch = check_quarter_match(raw_rows, expected_months, quarter_months_map)
        if mismatch:
            raise EngineError(mismatch)

    return build_output(template, raw_rows, quarter_label)


DATE_STRING_FORMATS = ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d")


def parse_date_value(value: object) -> Optional[date]:
    """Coerces a raw cell value to a `date`, handling real Excel dates AND
    plain text dates (some Zoho exports - e.g. the Leads module's "Created
    Time" - store dates as strings like "2024-08-09 21:30:59" rather than
    native Excel datetimes)."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        for fmt in DATE_STRING_FORMATS:
            try:
                return datetime.strptime(value.strip(), fmt).date()
            except ValueError:
                continue
    return None


def filter_rows(
    raw_rows: list[dict[str, object]],
    field: str,
    allowed_values: set[str],
    date_header: Optional[str] = None,
    date_range: Optional[tuple[date, date]] = None,
) -> list[dict[str, object]]:
    """Keep rows whose normalized `field` value is in `allowed_values`,
    optionally also requiring `date_header`'s value to fall within
    `date_range` (inclusive)."""
    out = []
    for row in raw_rows:
        if normalize_header(row.get(field)) not in allowed_values:
            continue
        if date_header and date_range:
            value_date = parse_date_value(row.get(date_header))
            if value_date is None or not (date_range[0] <= value_date <= date_range[1]):
                continue
        out.append(row)
    return out


def filter_rows_by_stage(
    raw_rows: list[dict[str, object]],
    stages: set[str],
    date_header: Optional[str] = None,
    date_range: Optional[tuple[date, date]] = None,
) -> list[dict[str, object]]:
    """Keep rows whose normalized "stage" value is in `stages`, optionally also
    requiring `date_header`'s value to fall within `date_range` (inclusive)."""
    return filter_rows(raw_rows, "stage", stages, date_header, date_range)


def build_master_workbook(
    deals_template: ParsedTemplate,
    pipeline_template: ParsedTemplate,
    master_raw_rows: list[dict[str, object]],
    current_quarter: str,
    complete_quarters: list[str],
    quarter_date_ranges: dict[str, tuple[date, date]],
    fiscal_year_range_: tuple[date, date],
    deals_stages: set[str],
    pipeline_high_stages: set[str],
    pipeline_low_stages: set[str],
    leads_template: Optional[ParsedTemplate] = None,
    leads_raw_rows: Optional[list[dict[str, object]]] = None,
    lead_status_filter: Optional[str] = None,
    fiscal_year_label_: Optional[str] = None,
) -> bytes:
    """Builds one workbook with a sheet per quarter of the current fiscal year
    already fully elapsed (Deals), a "Till Date" Deals sheet for the quarter in
    progress, two Pipeline sheets covering the WHOLE fiscal year (not just the
    current quarter) split by stage rather than probability, and (if leads
    data is supplied) a New Leads sheet also scoped to the whole fiscal year -
    all derived from unfiltered raw exports, replacing the manual per-sheet
    Zoho filtering.
    """
    out_wb = Workbook()
    first_sheet = True

    def next_ws(title: str) -> Worksheet:
        nonlocal first_sheet
        if first_sheet:
            ws = out_wb.active
            ws.title = title
            first_sheet = False
        else:
            ws = out_wb.create_sheet(title)
        return ws

    for q in complete_quarters:
        rows = filter_rows_by_stage(
            master_raw_rows, deals_stages, "closing date", quarter_date_ranges[q]
        )
        write_report_sheet(next_ws(f"DEALS For {q}"), deals_template, rows, q)

    till_date_rows = filter_rows_by_stage(
        master_raw_rows, deals_stages, "closing date", quarter_date_ranges[current_quarter]
    )
    write_report_sheet(
        next_ws(f"DEALS For {current_quarter} - Till Date"),
        deals_template,
        till_date_rows,
        f"{current_quarter} - Till Date",
    )

    high = filter_rows_by_stage(
        master_raw_rows, pipeline_high_stages, "closing date", fiscal_year_range_
    )
    low = filter_rows_by_stage(
        master_raw_rows, pipeline_low_stages, "closing date", fiscal_year_range_
    )

    write_report_sheet(
        next_ws(f"Pipeline Deals for {current_quarter}"),
        pipeline_template,
        high,
        current_quarter,
    )
    write_report_sheet(
        next_ws(f"Pipeline Deals for {current_quarter} - <40%"),
        pipeline_template,
        low,
        f"{current_quarter} - <40%",
    )

    if leads_template is not None and leads_raw_rows is not None:
        lead_rows = filter_rows(
            leads_raw_rows,
            "lead status",
            {lead_status_filter},
            "created time",
            fiscal_year_range_,
        )
        write_report_sheet(
            next_ws(f"New Leads FY - {fiscal_year_label_}"),
            leads_template,
            lead_rows,
            fiscal_year_label_,
        )

    out_wb.calculation.fullCalcOnLoad = True
    buffer = io.BytesIO()
    out_wb.save(buffer)
    buffer.seek(0)
    return buffer.read()
