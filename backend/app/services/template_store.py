from pathlib import Path

from app.config import REPORT_TYPES, template_path


class TemplateNotFoundError(Exception):
    pass


def get_template_path(report_type: str) -> Path:
    if report_type not in REPORT_TYPES:
        raise ValueError(f"Unknown report type: {report_type}")
    path = template_path(report_type)
    if not path.exists():
        raise TemplateNotFoundError(f"No template has been uploaded yet for {report_type}.")
    return path


def save_template(report_type: str, file_bytes: bytes) -> Path:
    if report_type not in REPORT_TYPES:
        raise ValueError(f"Unknown report type: {report_type}")
    path = template_path(report_type)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(file_bytes)
    return path
