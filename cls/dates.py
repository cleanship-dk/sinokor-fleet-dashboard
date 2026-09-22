"""Date coercion for a workbook that mixes real datetimes, Excel serials and free text.

The Sinokor workbook stores dates four different ways in the same column:
  * real ``datetime`` values (most cells)
  * Excel serial numbers (if a sheet was re-saved oddly)
  * day-first text such as ``13/7/2026`` (Fleet Schedule)
  * month-first text such as ``1/24/2026`` (Fleet Contact Info)
  * dotted text such as ``2022.10.31``

``parse_date_column`` resolves the ambiguous ``a/b/yyyy`` case per column by taking a majority
vote from the unambiguous values in that same column, so a future dump that flips convention
still parses correctly.
"""
from __future__ import annotations

import datetime as _dt
import re
from typing import Any, Iterable

import pandas as pd

_EXCEL_EPOCH = _dt.datetime(1899, 12, 30)

# 2026-07-13 / 2026.07.13 / 2026年7月13日-ish
_ISO_RE = re.compile(r"^(\d{4})\s*[-./년]\s*(\d{1,2})\s*[-./월]\s*(\d{1,2})\s*일?$")
# 13/7/2026 or 1/24/2026 -- ambiguous without context
_SLASH_RE = re.compile(r"^(\d{1,2})\s*[-./]\s*(\d{1,2})\s*[-./]\s*(\d{2,4})$")
# "2028년 3월(차기입거)"  -> year/month only, no day
_KO_YM_RE = re.compile(r"^(\d{4})\s*년\s*(\d{1,2})\s*월")
_YM_RE = re.compile(r"^(\d{4})[-./](\d{1,2})$")

_NOT_A_DATE = {
    "", "x", "n/a", "na", "-", "--", "tbc", "tba", "none", "nil",
    "not planned", "unknown", "확인불가", "미정",
}


def _clean(text: Any) -> str:
    return re.sub(r"\s+", " ", str(text)).strip()


def coerce_date(value: Any, dayfirst: bool | None = None) -> _dt.date | None:
    """Best-effort single-value coercion. Returns ``None`` for anything that is not a date.

    ``dayfirst`` only matters for the ambiguous ``a/b/yyyy`` form where both numbers are <= 12.
    ``None`` means "undecided" and such values are rejected rather than guessed; use
    :func:`parse_date_column` to have the convention inferred from sibling values.
    """
    if value is None:
        return None
    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value
    if isinstance(value, pd.Timestamp):
        return None if pd.isna(value) else value.date()
    if isinstance(value, float) and pd.isna(value):
        return None

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        # Excel serial. Guard the range so IMO numbers / LOA metres are never mistaken for dates.
        if 20000 <= float(value) <= 80000:
            return (_EXCEL_EPOCH + _dt.timedelta(days=float(value))).date()
        return None

    text = _clean(value)
    if text.lower() in _NOT_A_DATE:
        return None

    m = _ISO_RE.match(text)
    if m:
        return _safe(int(m.group(1)), int(m.group(2)), int(m.group(3)))

    m = _KO_YM_RE.match(text) or _YM_RE.match(text)
    if m:
        # Month-only precision: anchor to the 1st so it still sorts and filters sensibly.
        return _safe(int(m.group(1)), int(m.group(2)), 1)

    m = _SLASH_RE.match(text)
    if m:
        a, b, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000
        if a > 12 and b <= 12:
            return _safe(y, b, a)          # unambiguously day-first
        if b > 12 and a <= 12:
            return _safe(y, a, b)          # unambiguously month-first
        if a <= 12 and b <= 12:
            if dayfirst is True:
                return _safe(y, b, a)
            if dayfirst is False:
                return _safe(y, a, b)
            return None                    # genuinely ambiguous -- caller decides
        return None

    # Last resort: let pandas try, but never let it invent a date from a bare number.
    if any(ch.isalpha() for ch in text):
        try:
            ts = pd.to_datetime(text, dayfirst=bool(dayfirst), errors="raise")
            return None if pd.isna(ts) else ts.date()
        except Exception:
            return None
    return None


def _safe(year: int, month: int, day: int) -> _dt.date | None:
    try:
        return _dt.date(year, month, day)
    except ValueError:
        return None


def infer_dayfirst(values: Iterable[Any], report: bool = False):
    """Majority vote over the unambiguous ``a/b/yyyy`` values in a column.

    With ``report=True`` returns ``(dayfirst, guessed)``, where ``guessed`` is True when nothing
    in the column settled the question and the default had to be assumed. Callers use that to
    warn, because a silently wrong convention shifts every date by up to eleven months.
    """
    day_first = month_first = 0
    for value in values:
        if isinstance(value, (_dt.date, _dt.datetime)) or value is None:
            continue
        m = _SLASH_RE.match(_clean(value))
        if not m:
            continue
        a, b = int(m.group(1)), int(m.group(2))
        if a > 12 >= b:
            day_first += 1
        elif b > 12 >= a:
            month_first += 1
    if day_first or month_first:
        decided, guessed = day_first >= month_first, False
    else:
        # The workbook is Korean/European in origin, so day-first is the safer default.
        decided, guessed = True, any(_SLASH_RE.match(_clean(v)) for v in values
                                     if not isinstance(v, (_dt.date, _dt.datetime)) and v is not None)
    return (decided, guessed) if report else decided


def parse_date_column(values: Iterable[Any]) -> list[_dt.date | None]:
    values = list(values)
    dayfirst = infer_dayfirst(values)
    return [coerce_date(v, dayfirst=dayfirst) for v in values]


def is_datelike_text(value: Any) -> bool:
    """True when a cell holds a date written as text (used by the data-quality report)."""
    if value is None or isinstance(value, (_dt.date, _dt.datetime)):
        return False
    text = _clean(value)
    if not text or text.lower() in _NOT_A_DATE:
        return False
    return bool(_ISO_RE.match(text) or _SLASH_RE.match(text) or _KO_YM_RE.match(text) or _YM_RE.match(text))


def months_between(earlier: _dt.date | None, later: _dt.date | None) -> float | None:
    """Fractional months between two dates (30.44-day month), signed."""
    if earlier is None or later is None:
        return None
    return (later - earlier).days / 30.4375


def years_between(earlier: _dt.date | None, later: _dt.date | None) -> float | None:
    if earlier is None or later is None:
        return None
    return (later - earlier).days / 365.25


def fmt(value: _dt.date | None, style: str = "%d %b %Y") -> str:
    return "" if value is None else value.strftime(style)
