"""Pure diary merge / conversion logic (no COM, no file IO, stdlib only).

This module is the Python port of the pure helpers in Export-Diary.ps1
(Split-DiaryList / Split-DiaryNames / ConvertTo-DiaryEvent / Merge-DiaryEvents /
ConvertTo-Utc8601). It has ZERO Windows / pywin32 dependencies so it imports and
unit-tests cleanly on any platform (Linux, macOS, Windows), Python 3.8+.

It reads and writes the exact same diary.json event shape as the PowerShell
version, so a diary.json produced by either implementation is a valid input to
the other:
  * lastModified / cancelledAt / lastRunUtc : UTC "yyyy-MM-ddTHH:mm:ssZ"
  * start / end / windowFrom / windowTo     : local offset "yyyy-MM-ddTHH:mm:ss+HH:MM"
  * id scheme                               : "<GlobalAppointmentID>|<startUtcIso>"
  * field ORDER of an event dict            : identical to the PS [ordered] object
"""

from datetime import datetime, timezone


# ---------------------------------------------------------------------------
#  Formatting helpers
# ---------------------------------------------------------------------------

def iso_utc(dt):
    """UTC ISO-8601 with a trailing Z, e.g. 2026-07-03T12:00:00Z.

    Mirrors ConvertTo-Utc8601: converts the (timezone-aware) value to UTC and
    formats without fractional seconds. Naive datetimes are assumed to already
    be UTC (the COM layer always passes aware datetimes).
    """
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _iso_local_offset(dt):
    """Local-offset ISO-8601, e.g. 2026-07-06T14:00:00+01:00.

    Mirrors PowerShell's .ToString("yyyy-MM-ddTHH:mm:sszzz") -- the offset is
    rendered as +HH:MM (with a colon), which Python's %z does NOT do, so the
    offset is formatted by hand from the datetime's own utcoffset().
    """
    base = dt.strftime("%Y-%m-%dT%H:%M:%S")
    off = dt.utcoffset()
    if off is None:
        # Should not happen on the real path (COM datetimes are aware); treat
        # as UTC to stay well-defined rather than crashing.
        return base + "+00:00"
    total = int(off.total_seconds())
    sign = "+" if total >= 0 else "-"
    total = abs(total)
    hours = total // 3600
    minutes = (total % 3600) // 60
    return "%s%s%02d:%02d" % (base, sign, hours, minutes)


# ---------------------------------------------------------------------------
#  Splitters
# ---------------------------------------------------------------------------

def split_list(value, delimiter):
    """Split a delimited string into a trimmed, empty-dropped list.

    Mirrors Split-DiaryList: None / whitespace-only -> []; parts are trimmed and
    empties dropped.
    """
    if value is None:
        return []
    if not str(value).strip():
        return []
    out = []
    for part in str(value).split(delimiter):
        trimmed = part.strip()
        if trimmed:
            out.append(trimmed)
    return out


def split_names(value):
    """Split Outlook's semicolon-joined display-name string. Mirrors Split-DiaryNames."""
    return split_list(value, ";")


# ---------------------------------------------------------------------------
#  Event builder
# ---------------------------------------------------------------------------

def build_event(subject, start, end, all_day, organizer, required, optional,
                location, categories, global_id, last_modified, meeting_status,
                is_recurring, description=None):
    """Build one canonical diary event dict from scalar values.

    Mirrors ConvertTo-DiaryEvent. PURE: callers read COM properties and pass
    them in. start / end / last_modified must be timezone-aware datetimes.

    The returned dict preserves the exact field order of the PowerShell
    [ordered] object (Python dicts keep insertion order). MeetingStatus 5
    (olMeetingCanceled) or 7 (olMeetingReceivedAndCanceled) -> status
    "cancelled". A 'description' key is present only when description is not
    None, and is capped at 2000 characters.
    """
    start_utc_iso = iso_utc(start)
    event_id = "%s|%s" % (global_id, start_utc_iso)

    status = "active"
    if meeting_status == 5 or meeting_status == 7:
        status = "cancelled"

    event = {
        "id": event_id,
        "subject": subject,
        "startDate": start.strftime("%Y-%m-%d"),
        "startTime": start.strftime("%H:%M"),
        "endDate": end.strftime("%Y-%m-%d"),
        "endTime": end.strftime("%H:%M"),
        "start": _iso_local_offset(start),
        "end": _iso_local_offset(end),
        "organizer": organizer,
        "requiredAttendees": split_names(required),
        "optionalAttendees": split_names(optional),
        "location": location,
        "categories": split_list(categories, ","),
        "isRecurring": bool(is_recurring),
        "isAllDay": bool(all_day),
        "status": status,
        "lastModified": iso_utc(last_modified),
        "cancelledAt": None,
    }

    # Only attach a description when body extraction was requested.
    if description is not None:
        body = description
        if len(body) > 2000:
            body = body[:2000]
        event["description"] = body

    return event


# ---------------------------------------------------------------------------
#  Merge engine
# ---------------------------------------------------------------------------

def _parse_offset_to_utc(value):
    """Parse a diary start/end ISO string to a UTC datetime, or None on failure.

    Mirrors [datetimeoffset]::Parse(..., InvariantCulture).UtcDateTime used in
    Merge-DiaryEvents pass-2. Accepts the local-offset form the events use
    (e.g. 2026-07-10T09:00:00+00:00) and also tolerates a trailing Z.
    """
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    if text.endswith("Z") or text.endswith("z"):
        text = text[:-1] + "+00:00"
    try:
        parsed = datetime.fromisoformat(text)
    except ValueError:
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed.astimezone(timezone.utc)


def merge_events(existing, pulled, window_from, window_to, now_utc):
    """Pure incremental merge. Mirrors Merge-DiaryEvents.

    Args:
        existing: events already in diary.json (list of dict).
        pulled: events freshly pulled for the window (list of dict).
        window_from, window_to: window bounds (aware datetimes), compared in UTC.
        now_utc: timestamp used when stamping cancelledAt (aware datetime).

    Returns:
        (events_list, counts) where counts is a dict with keys
        added / updated / cancelled / unchanged.
    """
    if existing is None:
        existing = []
    if pulled is None:
        pulled = []

    now_iso = iso_utc(now_utc)
    from_utc = window_from.astimezone(timezone.utc) if window_from.tzinfo else window_from.replace(tzinfo=timezone.utc)
    to_utc = window_to.astimezone(timezone.utc) if window_to.tzinfo else window_to.replace(tzinfo=timezone.utc)

    existing_by_id = {}
    for e in existing:
        existing_by_id[e["id"]] = e
    pulled_by_id = {}
    for p in pulled:
        pulled_by_id[p["id"]] = p

    merged = []
    added = 0
    updated = 0
    cancelled = 0
    unchanged = 0

    # ---- pass 1: everything present in the fresh pull ----
    for p in pulled:
        if p["id"] not in existing_by_id:
            # brand new occurrence (may itself already be cancelled)
            merged.append(p)
            added += 1
            continue

        ex = existing_by_id[p["id"]]

        if p.get("status") == "cancelled":
            if ex.get("status") == "active":
                # active -> cancelled this run (host cancelled the meeting)
                p["cancelledAt"] = now_iso
                merged.append(p)
                cancelled += 1
            else:
                # was already cancelled -- carry forward, DO NOT re-stamp cancelledAt
                p["cancelledAt"] = ex.get("cancelledAt")
                merged.append(p)
                unchanged += 1
            continue

        # pulled event is active
        if ex.get("status") == "active" and ex.get("lastModified") == p.get("lastModified"):
            # unchanged -- preserve existing identity exactly
            merged.append(ex)
            unchanged += 1
        else:
            # content changed (or reactivated) -> update; active means cancelledAt null
            p["cancelledAt"] = None
            merged.append(p)
            updated += 1

    # ---- pass 2: existing events NOT in the fresh pull ----
    for e in existing:
        if e["id"] in pulled_by_id:
            continue

        # "In window" must mirror the pull's Restrict filter exactly
        # ([Start] >= from AND [End] <= to). An event whose start is inside but
        # whose end extends past windowTo is NOT in the pull, so its absence
        # proves nothing -- treating it as in-window would falsely cancel it.
        in_window = False
        start_utc = _parse_offset_to_utc(e.get("start"))
        if start_utc is not None:
            end_utc = _parse_offset_to_utc(e.get("end"))
            if end_utc is None:
                end_utc = start_utc
            if start_utc >= from_utc and end_utc <= to_utc:
                in_window = True

        if e.get("status") == "active" and in_window:
            # was active, sits inside the refreshed window, but vanished from
            # the pull -> it was deleted/cancelled in Outlook -> flip to cancelled.
            e["status"] = "cancelled"
            e["cancelledAt"] = now_iso
            merged.append(e)
            cancelled += 1
        else:
            # outside window, or already cancelled -> carry forward untouched.
            merged.append(e)
            unchanged += 1

    counts = {
        "added": added,
        "updated": updated,
        "cancelled": cancelled,
        "unchanged": unchanged,
    }
    return merged, counts
