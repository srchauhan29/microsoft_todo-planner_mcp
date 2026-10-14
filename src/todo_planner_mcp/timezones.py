"""Time zone detection.

'auto' resolves, in order:
  1. The signed-in user's Microsoft 365 time zone (Outlook mailbox settings) - the one
     To Do and Outlook use to display dates. Needs MailboxSettings.Read.
  2. This computer's Windows time zone.
  3. UTC.
Microsoft 365 reports Windows names (e.g. "AUS Eastern Standard Time"); they are
mapped to IANA names (e.g. "Australia/Sydney") using the CLDR table shipped with tzlocal.
"""

from __future__ import annotations

import logging
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import tzlocal
from tzlocal.windows_tz import win_tz

log = logging.getLogger(__name__)

# Legacy IANA aliases that some sources still return.
_CANONICAL = {"Asia/Calcutta": "Asia/Kolkata", "Asia/Katmandu": "Asia/Kathmandu", "Asia/Saigon": "Asia/Ho_Chi_Minh"}


def to_iana(name: str | None) -> str | None:
    """Map a Windows or IANA time zone name to a canonical IANA name, or None if unknown."""
    if not name:
        return None
    name = name.strip()
    candidate = win_tz.get(name, name)
    candidate = _CANONICAL.get(candidate, candidate)
    try:
        ZoneInfo(candidate)
    except (ZoneInfoNotFoundError, ValueError):
        log.warning("Unrecognised time zone name: %s", name)
        return None
    return candidate


def local_zone() -> str | None:
    try:
        return to_iana(tzlocal.get_localzone_name())
    except Exception as exc:  # tzlocal raises various errors on unusual systems
        log.warning("Could not read the computer's time zone: %s", exc)
        return None
