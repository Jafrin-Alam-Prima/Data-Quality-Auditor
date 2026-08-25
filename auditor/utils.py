"""Small value helpers shared by the whole audit."""

from __future__ import annotations

import datetime as _dt
import re
import unicodedata

from .config import ERROR_VALUES, PLACEHOLDERS

_WS_RE = re.compile(r"\s+")
_NON_ALNUM_RE = re.compile(r"[^a-z0-9]+")


def cell_text(value) -> str:
    """Turn any cell value into a clean, trimmed string.

    Floats that are whole numbers become "4540045" rather than "4540045.0",
    which matters because IDs and codes are often stored as numbers.
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float):
        if value != value:            # NaN
            return ""
        if value.is_integer():
            return str(int(value))
        return repr(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, (_dt.datetime, _dt.date, _dt.time)):
        return str(value)
    text = str(value)
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\xa0", " ").replace("​", "")
    return _WS_RE.sub(" ", text).strip()


def norm_key(value) -> str:
    """Comparison key: trimmed, single-spaced, lower-case."""
    return cell_text(value).casefold()


def reduce_name(value) -> str:
    """Reduce a sheet or column name to letters and digits only."""
    return _NON_ALNUM_RE.sub("", cell_text(value).casefold())


def is_blank(value) -> bool:
    return cell_text(value) == ""


def is_placeholder(value) -> bool:
    """True when a value was typed instead of real data (N/A, TBA, ...)."""
    key = norm_key(value)
    if not key:
        return False
    if key in PLACEHOLDERS or key in ERROR_VALUES:
        return True
    # "n/a - not purchased", "unknown supplier", ...
    if key.startswith(("n/a ", "n/a-", "n/a–", "na ", "tba ", "tbd ")):
        return True
    return False


def is_error_value(value) -> bool:
    return norm_key(value) in ERROR_VALUES


def to_number(value):
    """Parse a numeric amount, or return None when the value is not a number.

    Accepts real numbers and numeric strings such as "1,234.50", "$1 234.50"
    and "(500)" for negatives.  Rejects descriptions such as
    "OLD MATERIAL/UNKNOWN".
    """
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and value != value:
            return None
        return float(value)
    text = cell_text(value)
    if not text:
        return None

    negative = text.startswith("(") and text.endswith(")")
    if negative:
        text = text[1:-1]

    # Strip currency symbols, currency codes and thousands separators.
    text = re.sub(r"(?i)\b(usd|eur|gbp|mwk|zmw|zwl|kes|ugx|tzs|zar)\b", "", text)
    text = re.sub(r"[^\d,.\-+eE]", "", text)
    if not text or not re.search(r"\d", text):
        return None

    # 1.234,56 (European) vs 1,234.56 (English)
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        parts = text.split(",")
        if len(parts) > 1 and all(len(p) == 3 for p in parts[1:]):
            text = text.replace(",", "")       # 1,234,567
        else:
            text = text.replace(",", ".")      # 1234,56
    try:
        number = float(text)
    except ValueError:
        return None
    return -number if negative else number


def is_zero(value) -> bool:
    number = to_number(value)
    return number is not None and abs(number) < 1e-12


def quote_list(values, limit=3) -> str:
    """Render sample values as “A”, “B” and “C” for the feedback text."""
    items = [f"“{cell_text(v)}”" for v in list(values)[:limit] if cell_text(v)]
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    return ", ".join(items[:-1]) + " and " + items[-1]


def guess_country(filename: str, fallback_countries=()) -> str:
    """Guess the country / CO name from the file name.

    "Data Upload Template V3_Zambia RM.xlsx"        -> "Zambia RM"
    "Data Upload Template v3_Malawi_Zimbabwe.xlsx"  -> "Malawi & Zimbabwe"
    """
    from .config import FILENAME_NOISE_WORDS

    stem = re.sub(r"\.(xlsx|xlsm|xls)$", "", cell_text(filename), flags=re.I)
    segments = [s.strip() for s in re.split(r"[_]+", stem) if s.strip()]

    kept = []
    for segment in segments:
        words = [w for w in re.split(r"[\s\-]+", segment) if w]
        useful = [
            w for w in words
            if w.casefold() not in FILENAME_NOISE_WORDS
            and not re.fullmatch(r"(?i)v?\d+(\.\d+)*", w)
            and not re.fullmatch(r"\d{2,8}", w)
        ]
        if useful:
            kept.append(" ".join(useful))

    if kept:
        return " & ".join(kept)
    names = sorted({cell_text(c).title() for c in fallback_countries if cell_text(c)})
    return " & ".join(names) if names else ""
