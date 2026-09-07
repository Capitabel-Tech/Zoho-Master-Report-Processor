from datetime import date

from fastapi import APIRouter, HTTPException, UploadFile
from fastapi.responses import Response

from app.config import (
    DEALS_STAGES,
    LEAD_STATUS_FILTER,
    PERIOD_OPTIONS,
    PIPELINE_HIGH_STAGES,
    PIPELINE_LOW_STAGES,
    QUARTER_MONTHS,
    REPORT_TYPES,
    current_and_complete_quarters,
    fiscal_year_label,
    fiscal_year_range,
    fiscal_year_start_year,
    quarter_date_range,
)
from app.services import template_store
from app.services.excel_engine import (
    EngineError,
    build_master_workbook,
    process_quarter,
    read_raw_upload,
    read_template,
)

router = APIRouter(prefix="/api/process", tags=["process"])


@router.post("/master")
async def process_master_report(file: UploadFile, leads_file: UploadFile):
    for f in (file, leads_file):
        if not f.filename.lower().endswith((".xlsx", ".xlsm")):
            raise HTTPException(status_code=400, detail="Please upload .xlsx files.")

    try:
        deals_path = template_store.get_template_path("deals")
        pipeline_path = template_store.get_template_path("pipeline")
        leads_path = template_store.get_template_path("leads")
    except template_store.TemplateNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))

    content = await file.read()
    leads_content = await leads_file.read()

    today = date.today()
    current_quarter, complete_quarters = current_and_complete_quarters(today)
    fy_start = fiscal_year_start_year(today)
    quarter_ranges = {q: quarter_date_range(q, fy_start) for q in QUARTER_MONTHS}
    fy_range = fiscal_year_range(fy_start)
    fy_label = fiscal_year_label(fy_start)

    try:
        deals_template = read_template(deals_path)
        pipeline_template = read_template(pipeline_path)
        leads_template = read_template(leads_path, anchor_header="lead name")
        raw_rows = read_raw_upload(content)
        leads_rows = read_raw_upload(leads_content, anchor_header="lead name")
        output_bytes = build_master_workbook(
            deals_template,
            pipeline_template,
            raw_rows,
            current_quarter,
            complete_quarters,
            quarter_ranges,
            fy_range,
            DEALS_STAGES,
            PIPELINE_HIGH_STAGES,
            PIPELINE_LOW_STAGES,
            leads_template=leads_template,
            leads_raw_rows=leads_rows,
            lead_status_filter=LEAD_STATUS_FILTER,
            fiscal_year_label_=fy_label,
        )
    except EngineError as e:
        raise HTTPException(status_code=400, detail=str(e))

    out_name = f"MIS {today.isoformat()}.xlsx"
    return Response(
        content=output_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{out_name}"'},
    )


@router.post("/{report_type}/{period}")
async def process_raw_upload(report_type: str, period: str, file: UploadFile):
    report_type = report_type.lower()
    period = period.upper()
    if report_type not in REPORT_TYPES:
        raise HTTPException(status_code=404, detail=f"Unknown report type: {report_type}")
    period_labels = PERIOD_OPTIONS.get(report_type)
    if period_labels is None:
        raise HTTPException(
            status_code=404, detail=f"No single-period flow for report type: {report_type}"
        )
    if period not in period_labels:
        raise HTTPException(status_code=404, detail=f"Unknown period: {period}")
    if not file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="Please upload an .xlsx file.")

    try:
        template_path = template_store.get_template_path(report_type)
    except template_store.TemplateNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))

    content = await file.read()

    # The quarter-mismatch date check only makes sense for Deals, which is
    # scoped to an actual calendar quarter. Pipeline's "periods" are just the
    # full pipeline vs. a <40% probability subset, not different date ranges.
    quarter_kwargs = {}
    if report_type == "deals":
        quarter_kwargs = {
            "expected_months": QUARTER_MONTHS[period],
            "quarter_months_map": QUARTER_MONTHS,
        }

    try:
        output_bytes = process_quarter(
            template_path,
            content,
            period_labels[period],
            **quarter_kwargs,
        )
    except EngineError as e:
        raise HTTPException(status_code=400, detail=str(e))

    out_name = f"{report_type}_{period}_processed.xlsx"
    return Response(
        content=output_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{out_name}"'},
    )
