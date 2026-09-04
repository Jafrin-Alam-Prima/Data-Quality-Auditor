"""
Date parsing for the data-cleaning (preprocessing) feature.

This module never guesses. A value is only converted when it can be
recognised with real confidence; anything else is left exactly as it was in
the workbook and reported separately, so nothing is silently corrupted for a
data migration.
"""

from __future__ import annotations

import datetime as _dt
import re

from openpyxl.utils.datetime import from_excel

from .utils import cell_text

# The workbook's own Instructions sheet asks for this exact shape:
#   "YYYY-MM-DD 00:00:00.000"
TARGET_NUMBER_FORMAT = "yyyy-mm-dd hh:mm:ss.000"

_SLASH_DATE_RE = re.compile(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})$")
_ISO_DATE_RE = re.compile(
    r"^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$"
)

# A number this large or small in a date column is not a plausible Excel
# serial date -- treat it as unparseable rather than risk turning an
# unrelated number into a fabricated date. The lower bound matters more
# than it looks: a small leftover number like "11" or "12" (a real,
# genuinely corrupted value found in production data) is a valid Excel
# serial date for early January 1900, and would otherwise be silently
# "successfully" parsed into a nonsense date instead of being reported as
# unparseable. 1970-01-01 (serial 25569) is a conservative floor -- no
# genuine asset purchase or data-collection date is plausibly older than
# that -- and 2100-12-31 (serial 73415) is a generous ceiling.
_MIN_SERIAL, _MAX_SERIAL = 25569, 73415


def parse_date(value, day_first: bool) -> _dt.datetime | None:
    """Return a `datetime` for `value`, or None if it cannot be parsed with
    confidence.

    `day_first` resolves the only genuine ambiguity: a slash/dash date such
    as "01/02/2020" is DD/MM/YYYY when True, MM/DD/YYYY when False. An
    ISO-style date ("2020-02-01") is unambiguous and is always read as
    YYYY-MM-DD regardless of `day_first`.
    """
    if value is None:
        return None

    if isinstance(value, _dt.datetime):
        return value
    if isinstance(value, _dt.date):
        return _dt.datetime(value.year, value.month, value.day)

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if _MIN_SERIAL <= value <= _MAX_SERIAL:
            try:
                return from_excel(value)
            except (ValueError, OverflowError):
                return None
        return None

    text = cell_text(value)
    if not text:
        return None

    iso = _ISO_DATE_RE.match(text)
    if iso:
        year, month, day, hour, minute, second = iso.groups()
        return _safe_datetime(int(year), int(month), int(day),
                              int(hour or 0), int(minute or 0), int(second or 0))

    slash = _SLASH_DATE_RE.match(text)
    if slash:
        first, second, year = (int(g) for g in slash.groups())
        day, month = (first, second) if day_first else (second, first)
        return _safe_datetime(year, month, day)

    return None


def _safe_datetime(year, month, day, hour=0, minute=0, second=0) -> _dt.datetime | None:
    try:
        return _dt.datetime(year, month, day, hour, minute, second)
    except ValueError:
        return None       # e.g. day 31 in a 30-day month -- not a real date
