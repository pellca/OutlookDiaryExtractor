"""stdlib unittest port of tests/Merge.Tests.ps1 for the pure diary logic.

Runs on any OS (no COM, no pywin32):
    python -m unittest tests.test_merge          # from the repo root
    python -m unittest tests.test_merge -v
    python export_diary.py --self-test           # on the target machine

Ports EVERY test group / assertion from Merge.Tests.ps1, using the same
deterministic UTC window/clock values so behaviour matches the PowerShell suite.
"""

import os
import sys
import unittest
from datetime import datetime, timezone

# Make diary_merge importable when run from anywhere.
_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

import diary_merge  # noqa: E402


def new_test_event(id, start, end="", status="active",
                   last_modified="2026-07-01T00:00:00Z", cancelled_at=None,
                   subject="Test"):
    """Mirror of New-TestEvent: a plain diary-event dict."""
    return {
        "id": id,
        "subject": subject,
        "start": start,
        "end": end,
        "status": status,
        "lastModified": last_modified,
        "cancelledAt": cancelled_at,
    }


def _utc(text):
    """Parse an ISO string with trailing Z into a UTC-aware datetime."""
    return datetime.fromisoformat(text.replace("Z", "+00:00")).astimezone(timezone.utc)


# Deterministic UTC window / clock (independent of the runner's timezone).
WFROM = _utc("2026-07-01T00:00:00Z")
WTO = _utc("2026-08-30T00:00:00Z")
NOW_UTC = _utc("2026-07-03T12:00:00Z")
NOW_ISO = "2026-07-03T12:00:00Z"

IN_WINDOW_START = "2026-07-10T09:00:00+00:00"
OUT_WINDOW_START = "2026-05-01T09:00:00+00:00"


def _merge(existing, pulled):
    return diary_merge.merge_events(existing, pulled, WFROM, WTO, NOW_UTC)


class MergeTests(unittest.TestCase):

    def test_add_new_event(self):
        existing = []
        pulled = [new_test_event("A", IN_WINDOW_START)]
        events, counts = _merge(existing, pulled)
        self.assertEqual(1, counts["added"], "new event counted as added")
        self.assertEqual(0, counts["updated"], "no updates")
        self.assertEqual(1, len(events), "one event in output")
        self.assertEqual("active", events[0]["status"], "new event active")

    def test_unchanged_event_identity_preserved(self):
        existing = [new_test_event("A", IN_WINDOW_START,
                                   last_modified="2026-07-01T00:00:00Z", subject="ORIGINAL")]
        # Same id + same lastModified but different subject in the pull -> keep EXISTING.
        pulled = [new_test_event("A", IN_WINDOW_START,
                                 last_modified="2026-07-01T00:00:00Z", subject="CHANGED")]
        events, counts = _merge(existing, pulled)
        self.assertEqual(1, counts["unchanged"], "unchanged counted")
        self.assertEqual(0, counts["updated"], "not updated")
        self.assertEqual("ORIGINAL", events[0]["subject"], "existing identity preserved")
        # identity preserved: the exact existing object is kept.
        self.assertIs(existing[0], events[0], "same object instance kept")

    def test_modified_event_updated(self):
        existing = [new_test_event("A", IN_WINDOW_START,
                                   last_modified="2026-07-01T00:00:00Z", subject="ORIGINAL")]
        pulled = [new_test_event("A", IN_WINDOW_START,
                                 last_modified="2026-07-02T09:30:00Z", subject="CHANGED")]
        events, counts = _merge(existing, pulled)
        self.assertEqual(1, counts["updated"], "changed lastModified counted as updated")
        self.assertEqual("CHANGED", events[0]["subject"], "updated to pulled content")
        self.assertIsNone(events[0]["cancelledAt"], "active update keeps cancelledAt null")

    def test_missing_from_pull_in_window_cancelled_and_stamped(self):
        existing = [new_test_event("A", IN_WINDOW_START, status="active")]
        pulled = []
        events, counts = _merge(existing, pulled)
        self.assertEqual(1, counts["cancelled"], "in-window absence counted as cancelled")
        self.assertEqual("cancelled", events[0]["status"], "status flipped to cancelled")
        self.assertEqual(NOW_ISO, events[0]["cancelledAt"], "cancelledAt stamped with now")

    def test_legacy_event_missing_cancelledat_does_not_crash(self):
        # Older-schema / hand-edited events may lack a cancelledAt key. The merge
        # must add it and stamp without raising.
        legacy = {
            "id": "A",
            "status": "active",
            "start": IN_WINDOW_START,
            "lastModified": "2026-07-01T00:00:00Z",
        }
        existing = [legacy]
        pulled = []
        threw = False
        try:
            events, counts = _merge(existing, pulled)
        except Exception:  # noqa: BLE001
            threw = True
        self.assertFalse(threw, "merge does not throw when existing event lacks cancelledAt")
        self.assertEqual(1, counts["cancelled"], "legacy in-window absence counted as cancelled")
        self.assertEqual("cancelled", events[0]["status"], "legacy event flipped to cancelled")
        self.assertEqual(NOW_ISO, events[0]["cancelledAt"], "cancelledAt added and stamped")

    def test_already_cancelled_not_restamped(self):
        existing = [new_test_event("A", IN_WINDOW_START, status="cancelled",
                                   cancelled_at="2026-06-15T08:00:00Z")]
        pulled = []
        events, counts = _merge(existing, pulled)
        self.assertEqual(0, counts["cancelled"], "already-cancelled not re-counted")
        self.assertEqual("2026-06-15T08:00:00Z", events[0]["cancelledAt"],
                         "original cancelledAt preserved (not restamped)")

    def test_event_outside_window_untouched(self):
        existing = [new_test_event("A", OUT_WINDOW_START, status="active")]
        pulled = []
        events, counts = _merge(existing, pulled)
        self.assertEqual(0, counts["cancelled"], "out-of-window event not cancelled")
        self.assertEqual("active", events[0]["status"], "out-of-window event stays active")
        self.assertIsNone(events[0]["cancelledAt"], "out-of-window cancelledAt still null")

    def test_event_straddling_window_end_not_falsely_cancelled(self):
        # Start inside the window, end past windowTo -> NOT in the pull; its
        # absence proves nothing, so it must be carried forward active.
        existing = [new_test_event("A", "2026-08-29T23:00:00+00:00",
                                   end="2026-08-30T02:00:00+00:00", status="active")]
        pulled = []
        events, counts = _merge(existing, pulled)
        self.assertEqual(0, counts["cancelled"], "straddling event not counted cancelled")
        self.assertEqual("active", events[0]["status"], "straddling event stays active")
        self.assertIsNone(events[0]["cancelledAt"], "straddling event cancelledAt still null")

    def test_meetingstatus_cancelled_in_pull_cancelled_and_stamped(self):
        existing = [new_test_event("A", IN_WINDOW_START, status="active",
                                   last_modified="2026-07-01T00:00:00Z")]
        pulled = [new_test_event("A", IN_WINDOW_START, status="cancelled",
                                 last_modified="2026-07-02T00:00:00Z")]
        events, counts = _merge(existing, pulled)
        self.assertEqual(1, counts["cancelled"], "pulled-cancelled active event counted as cancelled")
        self.assertEqual("cancelled", events[0]["status"], "status cancelled from pull")
        self.assertEqual(NOW_ISO, events[0]["cancelledAt"], "cancelledAt stamped on flip")

    def test_new_pulled_already_cancelled_added_as_cancelled(self):
        existing = []
        pulled = [new_test_event("Z", IN_WINDOW_START, status="cancelled")]
        events, counts = _merge(existing, pulled)
        self.assertEqual(1, counts["added"], "new cancelled event added")
        self.assertEqual("cancelled", events[0]["status"], "new event carries cancelled status")
        self.assertIsNone(events[0]["cancelledAt"], "new cancelled event cancelledAt None")

    def test_full_resync_empty_existing_all_added(self):
        existing = []
        pulled = [new_test_event("A", IN_WINDOW_START),
                  new_test_event("B", IN_WINDOW_START)]
        events, counts = _merge(existing, pulled)
        self.assertEqual(2, counts["added"], "full resync adds everything")
        self.assertEqual(0, counts["unchanged"], "full resync has no carry-over")


class BuildEventTests(unittest.TestCase):

    def test_all_day_event_field_shape(self):
        d1 = _utc("2026-07-10T00:00:00Z")
        d2 = _utc("2026-07-11T00:00:00Z")
        lm = _utc("2026-07-01T00:00:00Z")
        ev = diary_merge.build_event(
            subject="Holiday", start=d1, end=d2, all_day=True, organizer="Someone",
            required="", optional="", location="", categories="", global_id="GID123",
            last_modified=lm, meeting_status=1, is_recurring=False)
        self.assertTrue(ev["isAllDay"], "isAllDay is true")
        self.assertFalse(ev["isRecurring"], "isRecurring is false")
        self.assertEqual("active", ev["status"], "non-cancelled MeetingStatus -> active")
        self.assertTrue(ev["id"].startswith("GID123|"), "id starts with GlobalAppointmentID")
        self.assertTrue(ev["id"].endswith("Z"), "id ends with startUtc Z")
        self.assertEqual(0, len(ev["requiredAttendees"]), "empty attendees -> empty array")
        self.assertEqual(0, len(ev["categories"]), "empty categories -> empty array")
        self.assertIsInstance(ev["requiredAttendees"], list, "requiredAttendees is a list")
        self.assertNotIn("description", ev, "no description key without body")

    def test_field_order_matches_ps(self):
        d1 = _utc("2026-07-10T00:00:00Z")
        ev = diary_merge.build_event(
            subject="x", start=d1, end=d1, all_day=False, organizer="", required="",
            optional="", location="", categories="", global_id="G", last_modified=d1,
            meeting_status=1, is_recurring=False)
        expected = ["id", "subject", "startDate", "startTime", "endDate", "endTime",
                    "start", "end", "organizer", "requiredAttendees",
                    "optionalAttendees", "location", "categories", "isRecurring",
                    "isAllDay", "status", "lastModified", "cancelledAt"]
        self.assertEqual(expected, list(ev.keys()), "field order matches the PS ordered object")

    def test_description_present_and_capped(self):
        d1 = _utc("2026-07-10T00:00:00Z")
        long_body = "x" * 2500
        ev = diary_merge.build_event(
            subject="x", start=d1, end=d1, all_day=False, organizer="", required="",
            optional="", location="", categories="", global_id="G", last_modified=d1,
            meeting_status=1, is_recurring=False, description=long_body)
        self.assertIn("description", ev, "description key present when body given")
        self.assertEqual(2000, len(ev["description"]), "description capped at 2000 chars")

    def test_meetingstatus_5_and_7_cancelled(self):
        d1 = _utc("2026-07-10T00:00:00Z")
        d2 = _utc("2026-07-11T00:00:00Z")
        lm = _utc("2026-07-01T00:00:00Z")
        ev5 = diary_merge.build_event(
            subject="x", start=d1, end=d2, all_day=False, organizer="", required="",
            optional="", location="", categories="", global_id="G5", last_modified=lm,
            meeting_status=5, is_recurring=False)
        ev7 = diary_merge.build_event(
            subject="x", start=d1, end=d2, all_day=False, organizer="", required="",
            optional="", location="", categories="", global_id="G7", last_modified=lm,
            meeting_status=7, is_recurring=False)
        self.assertEqual("cancelled", ev5["status"], "MeetingStatus 5 -> cancelled")
        self.assertEqual("cancelled", ev7["status"], "MeetingStatus 7 -> cancelled")


class SplitTests(unittest.TestCase):

    def test_split_names_edge_cases(self):
        self.assertEqual(0, len(diary_merge.split_names("")), "empty string -> 0 names")
        self.assertEqual(0, len(diary_merge.split_names(None)), "None -> 0 names")
        self.assertEqual(2, len(diary_merge.split_names("Alice Smith; Bob Jones")), "two names")
        self.assertEqual(2, len(diary_merge.split_names("Alice; ; Bob;")),
                         "empties and trailing semicolons dropped")
        solo = diary_merge.split_names("  Solo Person  ")
        self.assertEqual(1, len(solo), "single padded name -> 1")
        self.assertEqual("Solo Person", solo[0], "name is trimmed")

    def test_split_list_categories(self):
        cats = diary_merge.split_list("Red, Green ,, ", ",")
        self.assertEqual(2, len(cats), "categories split, trimmed, empties dropped")
        self.assertEqual(["Red", "Green"], cats, "trimmed values")

    def test_split_list_none_and_whitespace(self):
        self.assertEqual([], diary_merge.split_list(None, ","), "None -> []")
        self.assertEqual([], diary_merge.split_list("   ", ","), "whitespace -> []")


class IsoUtcTests(unittest.TestCase):

    def test_iso_utc_format(self):
        dt = _utc("2026-07-03T12:00:00Z")
        self.assertEqual("2026-07-03T12:00:00Z", diary_merge.iso_utc(dt), "UTC ISO with Z")

    def test_iso_utc_converts_offset(self):
        from datetime import timedelta
        tz = timezone(timedelta(hours=1))
        dt = datetime(2026, 7, 3, 13, 0, 0, tzinfo=tz)  # 12:00 UTC
        self.assertEqual("2026-07-03T12:00:00Z", diary_merge.iso_utc(dt),
                         "offset datetime converted to UTC")


if __name__ == "__main__":
    unittest.main()
