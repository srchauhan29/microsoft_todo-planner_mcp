"""Date helpers. Users type local dates; Graph wants specific shapes per workload."""

from __future__ import annotations

from datetime import datetime, time, timezone
from zoneinfo import ZoneInfo


def _parse_local(value: str, tz: ZoneInfo) -> datetime:
    value = value.strip()
    try:
        if len(value) == 10:  # YYYY-MM-DD
            return datetime.combine(datetime.strptime(value, "%Y-%m-%d").date(), time(0, 0), tzinfo=tz)
        dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as exc:
        raise ValueError(f"Invalid date '{value}'. Use YYYY-MM-DD or YYYY-MM-DDTHH:MM.") from exc
    return dt if dt.tzinfo else dt.replace(tzinfo=tz)


def todo_datetime(value: str, tz: ZoneInfo) -> dict[str, str]:
    """To Do dateTimeTimeZone, expressed in the user's zone so the due *date* stays correct."""
    local = _parse_local(value, tz).astimezone(tz)
    return {"dateTime": local.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": tz.key}


def planner_datetime(value: str, tz: ZoneInfo) -> str:
    """Planner wants an ISO 8601 UTC timestamp. Midnight local keeps the date correct in the UI."""
    return _parse_local(value, tz).astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
