from fastapi import APIRouter, HTTPException, UploadFile
from fastapi.responses import Response

from app.config import PERIOD_OPTIONS, QUARTER_MONTHS, REPORT_TYPES
from app.services import template_store
from app.services.excel_engine import EngineError, process_quarter

router = APIRouter(prefix="/api/process", tags=["process"])


@router.post("/{report_type}/{period}")
async def process_raw_upload(report_type: str, period: str, file: UploadFile):
    report_type = report_type.lower()
    period = period.upper()
    if report_type not in REPORT_TYPES:
        raise HTTPException(status_code=404, detail=f"Unknown report type: {report_type}")
    period_labels = PERIOD_OPTIONS[report_type]
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
