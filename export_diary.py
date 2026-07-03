#!/usr/bin/env python3
"""Extract a delegate mailbox's Outlook calendar to JSON via desktop Outlook COM.

Python port of Export-Diary.ps1. The corporate target machine runs PowerShell in
Constrained Language Mode (which blocks COM); CLM does not constrain Python, so
this port drives the same automation through pywin32.

The pure merge/diff/conversion logic lives in diary_merge.py (no COM, importable
anywhere). All COM access lives inside get_diary_events_from_outlook, and pywin32
is imported lazily INSIDE that function so this module imports fine on Linux for
testing (`python3 -c "import export_diary"`).

This module reads and writes the exact same diary.json format as the PowerShell
version, so files round-trip between the two implementations.

Target: Python 3.8+ (no match statements, no runtime X | Y annotations).
"""

import argparse
import json
import os
import sys
import time
import unittest
from datetime import datetime, timedelta

import diary_merge


# ============================================================================
#  Window / IO helpers (no COM)
# ============================================================================

def _parse_date_arg(text):
    """Parse a --from / --to value. Try ISO first, then common fallbacks.

    Mirrors [datetime]::Parse in Get-DiaryWindow but with an explicit, portable
    set of accepted formats.
    """
    try:
        return datetime.fromisoformat(text)
    except ValueError:
        pass
    for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S",
                "%m/%d/%Y", "%m/%d/%Y %H:%M", "%d/%m/%Y", "%d/%m/%Y %H:%M"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    raise ValueError("Could not parse date '%s'." % text)


def get_diary_window(from_str, to_str, days_ahead):
    """Resolve [from, to] from explicit args or the rolling days_ahead default.

    Mirrors Get-DiaryWindow. Returns (window_from, window_to) as local-time
    aware datetimes.
    """
    local_tz = datetime.now().astimezone().tzinfo
    if from_str and to_str:
        wf = _parse_date_arg(from_str)
        wt = _parse_date_arg(to_str)
        if wf.tzinfo is None:
            wf = wf.replace(tzinfo=local_tz)
        if wt.tzinfo is None:
            wt = wt.replace(tzinfo=local_tz)
        return wf, wt
    if from_str or to_str:
        raise ValueError("Specify BOTH --from and --to (an explicit range), or neither (use --days-ahead).")
    start = datetime.now().astimezone().replace(hour=0, minute=0, second=0, microsecond=0)
    return start, start + timedelta(days=days_ahead)


def read_existing_diary(path):
    """Load events from an existing diary.json. Mirrors Read-ExistingDiary.

    Missing / empty / corrupt file -> warn (except for plain missing) + empty list.
    """
    if not os.path.exists(path):
        return []
    try:
        with open(path, "r", encoding="utf-8") as fh:
            raw = fh.read()
        if not raw.strip():
            return []
        parsed = json.loads(raw)
        events = parsed.get("events")
        if events is not None:
            return list(events)
        return []
    except Exception as ex:  # noqa: BLE001 - match PS "treat as empty" behaviour
        sys.stderr.write("WARNING: Could not parse existing '%s' (%s). Treating as empty.\n" % (path, ex))
        return []


def write_diary_file(path, mailbox, window_from, window_to, events):
    """Write meta + events atomically as UTF-8 (no BOM). Mirrors Write-DiaryFile."""
    meta = {
        "mailbox": mailbox,
        "lastRunUtc": diary_merge.iso_utc(datetime.now().astimezone()),
        "windowFrom": diary_merge._iso_local_offset(window_from),
        "windowTo": diary_merge._iso_local_offset(window_to),
    }
    doc = {
        "meta": meta,
        "events": list(events),
    }

    directory = os.path.dirname(path)
    if directory and not os.path.isdir(directory):
        os.makedirs(directory, exist_ok=True)

    tmp = path + ".tmp"
    # UTF-8 without BOM; ensure_ascii=False keeps non-ASCII text readable.
    with open(tmp, "w", encoding="utf-8", newline="\n") as fh:
        json.dump(doc, fh, indent=2, ensure_ascii=False)
    os.replace(tmp, path)


# ============================================================================
#  Locale-safe Restrict date formatting (Windows-only code path)
# ============================================================================

def get_restrict_datetime(dt):
    """Format a datetime for an Outlook Restrict filter using the machine's
    Windows locale short-date + time format.

    THE CLASSIC RESTRICT GOTCHA: Outlook parses Restrict date strings using the
    current Windows user locale, NOT a fixed format. Hard-coding
    strftime('%m/%d/%Y') would break on any non-US locale (e.g. dd/MM/yyyy in the
    UK, yyyy.MM.dd elsewhere) and silently return wrong or zero results. So we ask
    Windows itself to format the date/time via GetDateFormatW / GetTimeFormatW
    with LOCALE_USER_DEFAULT, exactly matching how Outlook will re-parse it.

    This uses ctypes and only ever runs on Windows.
    """
    import ctypes  # local import: only reached on the Windows run path
    from ctypes import wintypes

    LOCALE_USER_DEFAULT = 0x0400
    DATE_SHORTDATE = 0x00000001
    TIME_NOSECONDS = 0x00000002

    class SYSTEMTIME(ctypes.Structure):
        _fields_ = [
            ("wYear", wintypes.WORD),
            ("wMonth", wintypes.WORD),
            ("wDayOfWeek", wintypes.WORD),
            ("wDay", wintypes.WORD),
            ("wHour", wintypes.WORD),
            ("wMinute", wintypes.WORD),
            ("wSecond", wintypes.WORD),
            ("wMilliseconds", wintypes.WORD),
        ]

    st = SYSTEMTIME()
    st.wYear = dt.year
    st.wMonth = dt.month
    st.wDayOfWeek = 0  # ignored by the formatting APIs
    st.wDay = dt.day
    st.wHour = dt.hour
    st.wMinute = dt.minute
    st.wSecond = dt.second
    st.wMilliseconds = 0

    # use_last_error=True is required for ctypes.get_last_error() to capture
    # the real Windows error code on the failure paths below.
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

    date_buf = ctypes.create_unicode_buffer(128)
    n = kernel32.GetDateFormatW(LOCALE_USER_DEFAULT, DATE_SHORTDATE,
                                ctypes.byref(st), None, date_buf, len(date_buf))
    if n == 0:
        raise ctypes.WinError(ctypes.get_last_error())

    time_buf = ctypes.create_unicode_buffer(128)
    n = kernel32.GetTimeFormatW(LOCALE_USER_DEFAULT, TIME_NOSECONDS,
                                ctypes.byref(st), None, time_buf, len(time_buf))
    if n == 0:
        raise ctypes.WinError(ctypes.get_last_error())

    return date_buf.value + " " + time_buf.value


# ============================================================================
#  COM layer -- only ever reached from run_export (real run path)
# ============================================================================

def _ensure_aware(dt):
    """COM datetimes from pywin32 are timezone-aware (pywintypes uses local tz).
    Normalize defensively: if naive, attach the local timezone."""
    if dt.tzinfo is None:
        return dt.astimezone()
    return dt


def get_diary_events_from_outlook(mailbox, window_from, window_to, include_body):
    """Attach to running Outlook, resolve the shared calendar, expand recurrences
    inside [window_from, window_to], and return a list of plain diary event dicts.

    Mirrors Get-DiaryEventsFromOutlook. pywin32 is imported here (lazily) so this
    module imports cleanly on non-Windows platforms for testing.
    """
    # win32com.client.Dispatch("Outlook.Application") attaches to the already
    # running single-instance desktop Outlook; we NEVER call .Quit() on it.
    import win32com.client  # noqa: import here so the module loads on Linux

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
    except Exception as ex:  # noqa: BLE001
        raise RuntimeError(
            "Could not attach to Outlook. Ensure the desktop Outlook client is "
            "running. Underlying error: %s" % ex)

    ns = outlook.GetNamespace("MAPI")

    recipient = ns.CreateRecipient(mailbox)
    recipient.Resolve()
    if not recipient.Resolved:
        raise RuntimeError(
            "Could not resolve mailbox '%s'. Check the address/name and that it "
            "exists in the address book." % mailbox)

    try:
        # 9 = olFolderCalendar
        calendar = ns.GetSharedDefaultFolder(recipient, 9)
    except Exception as ex:  # noqa: BLE001
        raise RuntimeError(
            "Access denied or unavailable for the shared calendar of '%s'. "
            "Confirm delegate/reviewer permission has been granted and the "
            "mailbox is reachable. Underlying error: %s" % (mailbox, ex))

    items = calendar.Items

    # ---- canonical fast recurrence pattern -- ORDER IS MANDATORY ----
    # Sort THEN IncludeRecurrences THEN Restrict. Any other order breaks
    # recurrence expansion / windowing.
    items.Sort("[Start]")
    items.IncludeRecurrences = True
    # Restrict date strings MUST use the machine's Windows locale short date/time
    # (see get_restrict_datetime); a hard-coded format breaks on non-US locales.
    filter_str = ("[Start] >= '" + get_restrict_datetime(window_from) +
                  "' AND [End] <= '" + get_restrict_datetime(window_to) + "'")
    restricted = items.Restrict(filter_str)

    results = []

    # Never use .Count / indexing on an IncludeRecurrences collection -- iterate
    # GetFirst()/GetNext() and hard-break once Start passes the window end.
    item = restricted.GetFirst()
    while item is not None:
        # Read Start first so we can hard-break once we pass the window end.
        s_start = _ensure_aware(item.Start)
        if s_start > window_to:
            del item
            break

        # Each property access is an RPC -- read ONLY what we need, once each.
        s_subject = item.Subject
        s_end = _ensure_aware(item.End)
        s_all_day = item.AllDayEvent
        s_org = item.Organizer
        s_req = item.RequiredAttendees
        s_opt = item.OptionalAttendees
        s_loc = item.Location
        s_cat = item.Categories
        s_gid = item.GlobalAppointmentID
        s_lm = _ensure_aware(item.LastModificationTime)
        s_ms = item.MeetingStatus
        s_rec = item.IsRecurring
        s_body = None
        if include_body:
            s_body = item.Body

        evt = diary_merge.build_event(
            subject=s_subject, start=s_start, end=s_end, all_day=s_all_day,
            organizer=s_org, required=s_req, optional=s_opt, location=s_loc,
            categories=s_cat, global_id=s_gid, last_modified=s_lm,
            meeting_status=s_ms, is_recurring=s_rec, description=s_body)

        results.append(evt)

        # pywin32 refcounts COM objects automatically -- no manual release
        # needed; just drop the per-iteration reference and keep no strays.
        item = restricted.GetNext()

    return results


# ============================================================================
#  Orchestration (real run path)
# ============================================================================

def run_export(mailbox, days_ahead, from_str, to_str, out_file, include_body,
               full_resync):
    """The real run path -- the only function that reaches COM. Mirrors
    Invoke-DiaryExport."""
    start_time = time.monotonic()

    # Resolve to an absolute path once: relative paths resolve against the
    # process CWD, which can differ under Task Scheduler.
    out_file = os.path.abspath(out_file)

    window_from, window_to = get_diary_window(from_str, to_str, days_ahead)

    print("Diary window: %s -> %s  (mailbox: %s)" % (
        window_from.strftime("%Y-%m-%d %H:%M"),
        window_to.strftime("%Y-%m-%d %H:%M"),
        mailbox))

    existing = []
    if full_resync:
        print("Full resync: existing diary ignored.")
    else:
        existing = read_existing_diary(out_file)

    pulled = get_diary_events_from_outlook(mailbox, window_from, window_to, include_body)

    now_utc = datetime.now().astimezone()
    merged, counts = diary_merge.merge_events(
        existing, pulled, window_from, window_to, now_utc)

    # Secondary key on id keeps output deterministic for tied start times
    # (avoids spurious diffs between runs and between implementations).
    ordered = sorted(merged, key=lambda e: (e.get("start") or "", e.get("id") or ""))

    write_diary_file(out_file, mailbox, window_from, window_to, ordered)

    elapsed = round(time.monotonic() - start_time, 2)

    print("")
    print("Diary written to %s" % out_file)
    print("  added:     %d" % counts["added"])
    print("  updated:   %d" % counts["updated"])
    print("  cancelled: %d" % counts["cancelled"])
    print("  unchanged: %d" % counts["unchanged"])
    print("  total:     %d" % len(ordered))
    print("  elapsed:   %ss" % elapsed)


# ============================================================================
#  CLI
# ============================================================================

def _build_parser():
    parser = argparse.ArgumentParser(
        prog="export_diary.py",
        description="Extract a delegate mailbox's Outlook calendar to JSON via "
                    "desktop Outlook COM, with incremental add/update/cancel merge.")
    parser.add_argument("--mailbox", default="",
                        help="Delegate mailbox to read (required), "
                             "e.g. delegate.name@yourorg.com.")
    parser.add_argument("--days-ahead", type=int, default=60,
                        help="Rolling window length; window = today 00:00 -> today+N (default 60).")
    parser.add_argument("--from", dest="from_str", default="",
                        help="Explicit range start (supply with --to, or neither).")
    parser.add_argument("--to", dest="to_str", default="",
                        help="Explicit range end (supply with --from, or neither).")
    parser.add_argument("--out-file", default="./diary.json",
                        help="Output JSON path (written atomically; dir auto-created).")
    parser.add_argument("--include-body", action="store_true",
                        help="Add a 'description' field (plain text, capped at 2000 chars).")
    parser.add_argument("--full-resync", action="store_true",
                        help="Ignore the existing file; rebuild the window from scratch.")
    parser.add_argument("--self-test", action="store_true",
                        help="Run the merge-logic tests and exit (no COM touched).")
    return parser


def main(argv=None):
    parser = _build_parser()
    args = parser.parse_args(argv)

    if args.self_test:
        # Run tests/test_merge.py via unittest; exit with its result. No COM.
        loader = unittest.TestLoader()
        repo_root = os.path.dirname(os.path.abspath(__file__))
        suite = loader.discover(start_dir=os.path.join(repo_root, "tests"),
                                pattern="test_merge.py")
        runner = unittest.TextTestRunner(verbosity=2)
        result = runner.run(suite)
        return 0 if result.wasSuccessful() else 1

    # Fail fast rather than argparse 'required': a hard requirement keeps the
    # error message clear for the scheduled non-interactive case.
    if not args.mailbox.strip():
        sys.stderr.write(
            "ERROR: --mailbox is required: the delegate mailbox whose calendar "
            "to extract, e.g. --mailbox 'first.last@yourorg.com'.\n")
        return 2

    try:
        # Validate the window pairing up front so a usage error exits with 2.
        get_diary_window(args.from_str, args.to_str, args.days_ahead)
    except ValueError as ex:
        sys.stderr.write("ERROR: %s\n" % ex)
        return 2

    try:
        run_export(
            mailbox=args.mailbox,
            days_ahead=args.days_ahead,
            from_str=args.from_str,
            to_str=args.to_str,
            out_file=args.out_file,
            include_body=args.include_body,
            full_resync=args.full_resync)
    except Exception as ex:  # noqa: BLE001 - surface a friendly message + exit 1
        sys.stderr.write("ERROR: %s\n" % ex)
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
