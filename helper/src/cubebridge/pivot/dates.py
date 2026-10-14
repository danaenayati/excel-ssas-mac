"""Excel's Date Filters as date windows, worked out from today's date.

Shared by the MDX builder and the built-in sample cube, so that both agree on what
"This Week" or "Last Quarter" means. As in Excel, weeks run Sunday to Saturday and
"Year to Date" ends today. Dates are the Mac's local calendar dates.
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass

from .model import DateFilter

_DAY = datetime.timedelta(days=1)


@dataclass(frozen=True)
class DateWindow:
    """Dates d with start <= d < end; either bound may be open. Or, for "All Dates in
    the Period", every date in one quarter or month of any year."""

    start: datetime.date | None = None
    end: datetime.date | None = None
    quarter: int | None = None
    month: int | None = None

    def contains(self, value: datetime.date) -> bool:
        if self.quarter is not None:
            return (value.month - 1) // 3 + 1 == self.quarter
        if self.month is not None:
            return value.month == self.month
        return (self.start is None or value >= self.start) and (self.end is None or value < self.end)


def _month_start(year: int, month: int) -> datetime.date:
    year += (month - 1) // 12
    month = (month - 1) % 12 + 1
    return datetime.date(year, month, 1)


def window(date_filter: DateFilter, today: datetime.date) -> DateWindow:
    op = date_filter.operator
    first, last = date_filter.value, date_filter.value2
    if op == "equals":
        return DateWindow(first, first + _DAY)
    if op == "before":
        return DateWindow(end=first)
    if op == "after":
        return DateWindow(start=first + _DAY)
    if op == "between":
        return DateWindow(first, last + _DAY)
    if op.startswith("quarter"):
        return DateWindow(quarter=int(op[len("quarter"):]))
    if op.startswith("month"):
        return DateWindow(month=int(op[len("month"):]))

    shift = {"next": 1, "this": 0, "last": -1}
    if op in ("tomorrow", "today", "yesterday"):
        day = today + _DAY * {"tomorrow": 1, "today": 0, "yesterday": -1}[op]
        return DateWindow(day, day + _DAY)
    if op == "yearToDate":
        return DateWindow(datetime.date(today.year, 1, 1), today + _DAY)
    for prefix, step in shift.items():
        if not op.startswith(prefix):
            continue
        period = op[len(prefix):]
        if period == "Week":
            # Sunday is day 0 of Excel's week.
            sunday = today - _DAY * ((today.weekday() + 1) % 7) + _DAY * (7 * step)
            return DateWindow(sunday, sunday + _DAY * 7)
        if period == "Month":
            start = _month_start(today.year, today.month + step)
            return DateWindow(start, _month_start(start.year, start.month + 1))
        if period == "Quarter":
            first_month = (today.month - 1) // 3 * 3 + 1
            start = _month_start(today.year, first_month + 3 * step)
            return DateWindow(start, _month_start(start.year, start.month + 3))
        if period == "Year":
            return DateWindow(datetime.date(today.year + step, 1, 1), datetime.date(today.year + step + 1, 1, 1))
    raise ValueError(f"unknown date filter {op!r}")
