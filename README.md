# Outlook Diary Extractor

Extracts a delegate mailbox's Outlook calendar to a JSON file for a downstream
AI "Diary Assistant", using **desktop Outlook COM automation** (Graph is blocked,
EWS is being retired). Recurring meetings are expanded to one entry per
occurrence, and re-runs perform an **incremental** add / update / cancel merge so
frequent scheduled refreshes stay cheap and never re-emit everything.

**Primary tool: `export_diary.py` (Python + pywin32).** The corporate target
machine runs PowerShell in **Constrained Language Mode**, which blocks COM. CLM
does not constrain Python, so the Python version is the one to use there.

The original `Export-Diary.ps1` (Windows PowerShell 5.1) is kept as an
**alternative for machines whose PowerShell is in `FullLanguage` mode** — see
[PowerShell alternative](#powershell-alternative-fulllanguage-only) below.

> Both implementations produce and consume the **identical** `diary.json`
> (same field names, same ISO formats, same `id` scheme, same `meta` block, same
> merge semantics), so a file written by one is a valid input to the other and
> you can switch between them freely.

---

## 10-second pre-flight (locked-down machine)

Confirm Python can drive Outlook COM. Run this one-liner (Outlook should already
be running and signed in):

```bat
python -c "import win32com.client as w; print(w.Dispatch('Outlook.Application').Session.CurrentUser.Name)"
```

If it prints your name, COM works. If it fails with an import error, install
pywin32 (below). If it fails to attach, start the Outlook desktop client first.

### Installing pywin32

```bat
pip install pywin32
```

On a corporate build without direct PyPI access, point pip at the internal
mirror, e.g.:

```bat
pip install --index-url https://<your-internal-mirror>/simple pywin32
```

(pywin32 is the only third-party dependency; the merge logic and tests are pure
stdlib.)

---

## Requirements

- Windows with the **desktop Outlook client running** and signed in.
- **Delegate / reviewer access** to the target mailbox's calendar.
- **Python 3.8+** and **pywin32**.
- The tool **attaches** to the running Outlook instance and **never calls
  `.Quit()`** — it will not close the user's Outlook.

---

## Usage

Default rolling window (today 00:00 -> today + 60 days), writes `./diary.json`:

```bat
python export_diary.py --mailbox "delegate.name@yourorg.com"
```

Different horizon:

```bat
python export_diary.py --mailbox "delegate.name@yourorg.com" --days-ahead 14
```

Explicit one-off range (both required together):

```bat
python export_diary.py --mailbox "delegate.name@yourorg.com" --from "2026-07-01" --to "2026-07-31"
```

Custom output file and include (truncated) plain-text bodies:

```bat
python export_diary.py --mailbox "delegate.name@yourorg.com" --out-file "C:\diary\diary.json" --include-body
```

Full rebuild (ignore the existing file, re-pull the whole window):

```bat
python export_diary.py --mailbox "delegate.name@yourorg.com" --full-resync
```

Run the built-in pure-logic tests and exit (no Outlook/COM touched):

```bat
python export_diary.py --self-test
```

### Parameters

| Parameter        | Default          | Meaning                                                        |
|------------------|------------------|----------------------------------------------------------------|
| `--mailbox`      | (required)       | Delegate mailbox to read.                                       |
| `--days-ahead`   | `60`             | Rolling window length; window = today 00:00 -> today+N.        |
| `--from` `--to`  | (none)           | Explicit range override (supply **both** or neither).          |
| `--out-file`     | `./diary.json`   | Output JSON path (written atomically; directory auto-created).  |
| `--include-body` | off              | Adds a `description` field (plain text, capped at 2000 chars). |
| `--full-resync`  | off              | Ignore existing file; rebuild the window from scratch.         |
| `--self-test`    | off              | Run merge-logic tests, exit 0 (pass) / 1 (fail).               |

Exit codes: `2` for usage errors (missing `--mailbox`, `--from`/`--to` mismatch),
`1` for runtime failures (Outlook not attachable, mailbox not resolved, shared
calendar denied), `0` on success.

---

## Scheduling frequent refresh (Task Scheduler)

Create a task that runs while the user is logged on (Outlook must be running):

- **Program/script:** `python.exe` (or the full path, e.g.
  `C:\Python312\python.exe`)
- **Add arguments:**

  ```
  "C:\Tools\OutlookDiaryExtractor\export_diary.py" --mailbox "delegate.name@yourorg.com" --days-ahead 60 --out-file "C:\Tools\OutlookDiaryExtractor\diary.json"
  ```

- **Trigger:** e.g. every 15 minutes.
- Run **only when user is logged on** (COM needs the interactive Outlook session).

Or register from a prompt:

```bat
schtasks /Create /TN "OutlookDiary" /SC MINUTE /MO 15 /TR "python.exe \"C:\Tools\OutlookDiaryExtractor\export_diary.py\" --mailbox \"delegate.name@yourorg.com\" --days-ahead 60" /RL LIMITED /F
```

Each run does a small, bounded COM query and an incremental merge, so it stays
fast even at high frequency.

---

## Output JSON

*(Identical for the Python and PowerShell implementations.)*

File shape:

```json
{
  "meta": {
    "mailbox":    "delegate.name@yourorg.com",
    "lastRunUtc": "2026-07-03T11:00:00Z",
    "windowFrom": "2026-07-03T00:00:00+01:00",
    "windowTo":   "2026-09-01T00:00:00+01:00"
  },
  "events": [ ... ]
}
```

Each event:

| Field               | Type            | Notes                                                                 |
|---------------------|-----------------|-----------------------------------------------------------------------|
| `id`                | string          | Merge key: `<GlobalAppointmentID>|<occurrenceStartUtc ISO>`. Stable per occurrence. |
| `subject`           | string          | Meeting subject.                                                      |
| `startDate`         | string          | `yyyy-MM-dd` (local).                                                 |
| `startTime`         | string          | `HH:mm` (local, 24h).                                                 |
| `endDate`           | string          | `yyyy-MM-dd` (local).                                                 |
| `endTime`           | string          | `HH:mm` (local, 24h).                                                 |
| `start`             | string          | ISO 8601 with local offset, e.g. `2026-07-06T14:00:00+01:00`.        |
| `end`               | string          | ISO 8601 with local offset.                                          |
| `organizer`         | string          | Organizer display name.                                              |
| `requiredAttendees` | array of string | Display names (semicolon list split, trimmed, empties dropped).     |
| `optionalAttendees` | array of string | Display names.                                                      |
| `location`          | string          | Location text.                                                       |
| `categories`        | array of string | Outlook categories (comma list split, trimmed).                      |
| `isRecurring`       | bool            | True if the appointment is part of a recurring series.               |
| `isAllDay`          | bool            | True for all-day events.                                             |
| `status`            | string          | `active` or `cancelled`.                                              |
| `lastModified`      | string          | Outlook `LastModificationTime`, UTC ISO (`...Z`).                     |
| `cancelledAt`       | string or null  | UTC ISO timestamp set when `status` first flips to `cancelled`; else `null`. |
| `description`       | string          | Present only with `--include-body`; plain text, max 2000 chars.      |

ISO formats are exact and shared by both implementations: `lastModified`,
`cancelledAt` and `lastRunUtc` are UTC `yyyy-MM-ddTHH:mm:ssZ`; `start`, `end`,
`windowFrom` and `windowTo` carry a local offset `yyyy-MM-ddTHH:mm:ss+HH:MM`.

### Incremental merge rules

*(Identical for the Python and PowerShell implementations.)*

- **New `id`** -> added.
- **Existing `id`, same `lastModified`** -> unchanged (existing entry kept as-is).
- **Existing `id`, changed `lastModified`** -> updated in place.
- **`MeetingStatus` 5 / 7** (cancelled / received-and-cancelled) -> `status: cancelled`.
- **Existing active event inside the refreshed window but absent from the pull**
  -> treated as deleted -> `status: cancelled`, `cancelledAt` stamped. ("Inside
  the window" mirrors the pull filter exactly: start >= from AND end <= to, so
  an event straddling the window's far edge is never falsely cancelled.)
- **Already-cancelled events** are kept and are **not** re-stamped.
- **Events outside the refreshed window** are left untouched (history accumulates).
- **Rescheduled meetings appear as a pair**: the occurrence-start is part of the
  merge `id`, so moving a meeting produces one `cancelled` entry (the old slot)
  plus one new `active` entry (the new slot). Downstream consumers should treat
  a cancel+add sharing the same `GlobalAppointmentID` prefix as a reschedule.
- `--full-resync` ignores the existing file and rebuilds the window.

The file is written atomically (`diary.json.tmp` then `os.replace`), so a
downstream reader never sees a torn file. Encoding is UTF-8 without BOM.

---

## Testing

The pure merge/diff and conversion logic (`diary_merge.py`) has **no COM
dependency** and is tested directly with stdlib `unittest` on any OS:

```bat
python -m unittest tests.test_merge       :: any OS, no pywin32 needed
python -m unittest tests.test_merge -v
python export_diary.py --self-test        :: same tests on the target box
```

`diary_merge.py` imports cleanly on Linux/macOS/Windows (stdlib only), and
`export_diary.py` imports pywin32 **lazily** (only when it actually reaches
Outlook), so both modules import — and the tests run — anywhere.

---

## Troubleshooting

**PowerShell `LanguageMode` is `ConstrainedLanguage`.**
This is exactly why the Python tool is the primary one. CLM blocks
`New-Object -ComObject`, so `Export-Diary.ps1` cannot run; Python is not subject
to CLM. **Use `export_diary.py`** (this repo's primary tool) — no policy change
required.

**`import win32com.client` fails (`ModuleNotFoundError`).**
pywin32 is not installed for the Python you invoked. Run `pip install pywin32`
(see [Installing pywin32](#installing-pywin32)); on locked-down builds use the
internal mirror.

**"Could not resolve mailbox ..."** (exit 1)
`CreateRecipient(...).Resolve()` failed. Check the address/name is exactly right
and resolves in Outlook's address book (try typing it into a new message's To
field). A UPN/email or a full display name usually works.

**"Access denied or unavailable for the shared calendar ..."** (exit 1)
`GetSharedDefaultFolder` failed. You need delegate or at least **Reviewer**
permission on the mailbox's calendar, and the mailbox owner's account must be
reachable. Ask the owner to share their Calendar, then retry.

**"Could not attach to Outlook ..."** (exit 1)
The desktop Outlook client is not running. Start Outlook, wait for it to finish
loading, and re-run.

**Wrong / no events, or a `Restrict` error about dates.**
The `Restrict` filter uses the machine's Windows locale short date/time (the
tool asks Windows itself to format the bounds via `GetDateFormatW` /
`GetTimeFormatW`, so non-US regional formats are handled correctly). If results
still look off, verify the window bounds printed at the top of the run.

---

## PowerShell alternative (FullLanguage only)

`Export-Diary.ps1` is the original single-file PowerShell 5.1 implementation. It
only works where `$ExecutionContext.SessionState.LanguageMode` is `FullLanguage`
(Constrained Language Mode blocks its COM calls — use the Python tool instead
there). It produces the **identical** `diary.json`.

Pre-flight:

```powershell
# Must be FullLanguage (Constrained Language Mode blocks COM):
$ExecutionContext.SessionState.LanguageMode
# Outlook COM must be creatable:
$o = New-Object -ComObject Outlook.Application; $o.Session.CurrentUser.Name
```

Usage (invoke with an execution-policy bypass rather than changing machine policy):

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Export-Diary.ps1 -Mailbox "delegate.name@yourorg.com"
.\Export-Diary.ps1 -Mailbox "delegate.name@yourorg.com" -DaysAhead 14
.\Export-Diary.ps1 -Mailbox "delegate.name@yourorg.com" -From "2026-07-01" -To "2026-07-31"
.\Export-Diary.ps1 -Mailbox "delegate.name@yourorg.com" -OutFile "C:\diary\diary.json" -IncludeBody
.\Export-Diary.ps1 -Mailbox "delegate.name@yourorg.com" -FullResync
.\Export-Diary.ps1 -SelfTest
```

Its pure-logic tests:

```powershell
pwsh -File tests/Merge.Tests.ps1        # any OS
powershell -File tests\Merge.Tests.ps1  # Windows PowerShell 5.1
.\Export-Diary.ps1 -SelfTest            # on the target box
```
