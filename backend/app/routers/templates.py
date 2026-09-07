import io

from fastapi import APIRouter, HTTPException, UploadFile
from fastapi.responses import FileResponse

from app.config import REPORT_TYPES
from app.services import template_store
from app.services.excel_engine import EngineError, read_template

router = APIRouter(prefix="/api/templates", tags=["templates"])


def _validate_report_type(report_type: str) -> str:
    report_type = report_type.lower()
    if report_type not in REPORT_TYPES:
        raise HTTPException(status_code=404, detail=f"Unknown report type: {report_type}")
    return report_type


@router.get("/{report_type}")
def download_template(report_type: str):
    report_type = _validate_report_type(report_type)
    try:
        path = template_store.get_template_path(report_type)
    except template_store.TemplateNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))
    return FileResponse(
        path,
        filename=f"{report_type}_template.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@router.post("/{report_type}")
async def upload_template(report_type: str, file: UploadFile):
    report_type = _validate_report_type(report_type)
    if not file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="Please upload an .xlsx file.")

    content = await file.read()

    # validate it parses as a well-formed template before accepting it
    try:
        read_template(io.BytesIO(content))
    except EngineError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception:
        raise HTTPException(status_code=400, detail="Could not read this file as a valid template.")

    template_store.save_template(report_type, content)
    return {"status": "ok", "report_type": report_type}
