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
        f'Could not find a header row containing "Deal Name" in the first 10 rows.'
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


def read_template(path) -> ParsedTemplate:
    wb = load_workbook(path, data_only=False)
    ws = wb.active

    header_row = _find_header_row(ws)
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


def read_raw_upload(file_bytes: bytes) -> list[dict[str, object]]:
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active

    header_row = _find_header_row(ws)
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
        deal_name = row_data.get("deal name")
        if deal_name is None:
            continue
        rows.append(row_data)

    if not rows:
        raise EngineError("No data rows found in the uploaded file.")
    return rows


def build_output(
    template: ParsedTemplate,
    raw_rows: list[dict[str, object]],
    quarter_label: Optional[str] = None,
) -> bytes:
    src_wb = load_workbook(template.wb_path_for_style, data_only=False)
    src_ws = src_wb.active

    out_wb = Workbook()
    out_ws = out_wb.active
    out_ws.title = src_ws.title or "Report"

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

    deal_name_col = next(
        (c.index for c in template.columns if c.header_norm == "deal name"), None
    )
    STRIPE_FILL = PatternFill(start_color="F2F6FB", end_color="F2F6FB", fill_type="solid")
    TOTAL_FILL = PatternFill(start_color="D9E2EC", end_color="D9E2EC", fill_type="solid")

    for i, raw_row in enumerate(raw_rows):
        out_row = first_data_row + i
        max_lines = 1
        row_values = {}
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
            dst_cell.value = value
            row_values[col.index] = value
            copy_style(example_cell, dst_cell)

            if isinstance(value, str) and not value.startswith("="):
                col_width = template.column_widths.get(get_column_letter(col.index))
                max_lines = max(max_lines, wrapped_line_count(value, col_width))

        is_total_row = deal_name_col is not None and str(
            row_values.get(deal_name_col) or ""
        ).strip().lower() in ("total", "grand total")

        for col in template.columns:
            dst_cell = out_ws.cell(out_row, col.index)
            if is_total_row:
                bold_font = copy.copy(dst_cell.font)
                bold_font = Font(
                    name=bold_font.name,
                    size=bold_font.size,
                    bold=True,
                    color=bold_font.color,
                )
                dst_cell.font = bold_font
                dst_cell.fill = TOTAL_FILL
            elif i % 2 == 1:
                dst_cell.fill = STRIPE_FILL

        out_ws.row_dimensions[out_row].height = max(min_row_height, max_lines * LINE_HEIGHT + ROW_PADDING)

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
