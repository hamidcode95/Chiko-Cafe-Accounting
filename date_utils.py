# -*- coding: utf-8 -*-
"""
Small helper around jdatetime so the rest of the app only deals with
Jalali (Shamsi) date strings in the format YYYY-MM-DD (e.g. 1405-06-03).
This format sorts correctly as plain text, which matters because we use
it directly as the Excel sheet name and rely on lexical ordering for
date-range reports.
"""

import jdatetime
import datetime

FA_MONTHS = [
    "فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور",
    "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند",
]

FA_WEEKDAYS = ["شنبه", "یکشنبه", "دوشنبه", "سه‌شنبه", "چهارشنبه", "پنجشنبه", "جمعه"]


def today_str():
    """Today's date as YYYY-MM-DD in the Jalali calendar."""
    return jdatetime.date.today().strftime("%Y-%m-%d")


def today_display():
    """Today's date, nicely formatted for humans, e.g. 'سه‌شنبه ۳ شهریور ۱۴۰۵'."""
    return to_display(today_str())


def to_display(jalali_str):
    """Convert 'YYYY-MM-DD' Jalali string to a human friendly Persian string."""
    try:
        y, m, d = (int(part) for part in jalali_str.split("-"))
        jd = jdatetime.date(y, m, d)
        weekday = FA_WEEKDAYS[jd.weekday()]
        return f"{weekday} {d} {FA_MONTHS[m - 1]} {y}"
    except (ValueError, IndexError):
        return jalali_str


def is_valid(jalali_str):
    try:
        y, m, d = (int(part) for part in jalali_str.split("-"))
        jdatetime.date(y, m, d)
        return True
    except (ValueError, IndexError):
        return False


def recent_dates(n=60):
    """Last n Jalali date strings ending today, most recent first.
    Handy for letting the user pick 'which day was that sale on?'."""
    today = jdatetime.date.today()
    return [(today - datetime.timedelta(days=i)).strftime("%Y-%m-%d") for i in range(n)]
