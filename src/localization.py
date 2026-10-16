"""Indonesian date/time display using the computer's current local timezone."""

from __future__ import annotations

import datetime as dt

INDONESIAN_WEEKDAYS = (
    "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu",
)
INDONESIAN_MONTHS = (
    "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
    "Agustus", "September", "Oktober", "November", "Desember",
)


def format_local_datetime(value: str | dt.datetime) -> str:
    """Format a stored or current timestamp in local time, in Bahasa Indonesia."""
    moment = dt.datetime.fromisoformat(value) if isinstance(value, str) else value
    moment = moment.astimezone()
    timezone = moment.tzname() or "waktu lokal"
    return (
        f"{INDONESIAN_WEEKDAYS[moment.weekday()]}, {moment.day} "
        f"{INDONESIAN_MONTHS[moment.month - 1]} {moment.year}"
        f" · {moment:%H:%M:%S} {timezone}"
    )
