from __future__ import annotations

from datetime import date, datetime, timedelta
from typing import Any

# Excel's day-0 epoch used by openpyxl for Windows workbooks
_EXCEL_EPOCH = datetime(1899, 12, 30)

_DATE_FORMATS = ("%Y-%m-%d", "%d.%m.%Y", "%Y/%m/%d", "%d/%m/%Y")


def looks_like_date(value: str) -> bool:
    """Return True if the string matches a supported numeric date format."""
    if not isinstance(value, str):
        return False
    value = value.strip()
    for fmt in _DATE_FORMATS:
        try:
            datetime.strptime(value, fmt)
            return True
        except ValueError:
            continue
    return False


def is_date_like(value: Any) -> bool:
    """Return True if the cell is a real date, Excel serial, or numeric date string."""
    if isinstance(value, (datetime, date)):
        return True
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return _from_excel_serial(value) is not None
    if isinstance(value, str) and looks_like_date(value):
        return True
    return False


def format_header_date(value: Any) -> str:
    parsed = parse_date(value)
    if parsed is not None:
        return parsed.strftime("%d.%m.%Y")
    return str(value)


def _from_excel_serial(value: int | float) -> date | None:
    """Convert an Excel serial day number to a date, if it looks plausible."""
    try:
        serial = float(value)
    except (TypeError, ValueError):
        return None
    # Seminar dates live in a narrow modern range; reject junk / indexes / scores
    if serial < 30000 or serial > 60000:  # ~1982–2064
        return None
    try:
        return (_EXCEL_EPOCH + timedelta(days=serial)).date()
    except OverflowError:
        return None


def parse_date(value: Any) -> date | None:
    """Convert Excel datetime/date, serial number, or DD.MM.YYYY-style string to date."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return _from_excel_serial(value)
    if not isinstance(value, str):
        return None

    text = value.strip()
    if not text:
        return None

    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None
