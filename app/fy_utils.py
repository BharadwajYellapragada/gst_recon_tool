"""
Indian financial year helpers (April to March).
"""
import datetime
import calendar
import pandas as pd

MONTH_ABBR = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]


def parse_ddmmyyyy(s):
    return datetime.datetime.strptime(s, "%d/%m/%Y").date()


def fy_label_for_date(d):
    """Indian FY: Apr Y to Mar Y+1 is labelled 'Y-(Y+1 last two digits)', e.g. 2022-23."""
    if d.month >= 4:
        start_year = d.year
    else:
        start_year = d.year - 1
    return f"{start_year}-{str(start_year + 1)[-2:]}"


def fy_bounds(fy_label):
    """Returns (start_date, end_date) for e.g. '2022-23' -> (2022-04-01, 2023-03-31)."""
    start_year = int(fy_label.split("-")[0])
    start = datetime.date(start_year, 4, 1)
    end = datetime.date(start_year + 1, 3, 31)
    return start, end


def add_months(d, n):
    month = d.month - 1 + n
    year = d.year + month // 12
    month = month % 12 + 1
    day = min(d.day, calendar.monthrange(year, month)[1])
    return datetime.date(year, month, day)


def period_str(d):
    """e.g. 2022-04-15 -> 'APR-22'"""
    return f"{d.strftime('%b').upper()}-{str(d.year)[-2:]}"


def fy_period_window(fy_label, extra_months=6):
    """
    List of 'MMM-YY' period strings covering the financial year plus `extra_months`
    beyond its end (to catch suppliers filing GSTR-1 late).
    """
    start, end = fy_bounds(fy_label)
    window_end = add_months(end, extra_months)
    periods = []
    cur = datetime.date(start.year, start.month, 1)
    end_marker = datetime.date(window_end.year, window_end.month, 1)
    while cur <= end_marker:
        periods.append(period_str(cur))
        cur = add_months(cur, 1)
    return periods


def fy_cutoff_date(fy_label, extra_months=6):
    """The date after which a supplier filing is considered outside the allowed window."""
    _, end = fy_bounds(fy_label)
    return add_months(end, extra_months)


def days_until_cutoff(fy_label, extra_months=6, as_of=None):
    as_of = as_of or datetime.date.today()
    cutoff = fy_cutoff_date(fy_label, extra_months)
    return (cutoff - as_of).days


def latest_ddmmyyyy(series):
    """Correct max() over a column of 'DD/MM/YYYY' strings — plain string max sorts wrong
    (e.g. '31/12/2022' > '31/03/2023' lexically). Returns '' for an empty/all-blank series."""
    if series is None or len(series) == 0:
        return ""
    parsed = pd.to_datetime(series, format="%d/%m/%Y", errors="coerce")
    parsed = parsed.dropna()
    if len(parsed) == 0:
        return ""
    return parsed.max().strftime("%d/%m/%Y")
