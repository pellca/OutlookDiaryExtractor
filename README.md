# Outlook Diary Extractor

Extracts a delegate mailbox's Outlook calendar to a JSON file for a downstream
AI "Diary Assistant", using **desktop Outlook COM automation** (Graph is blocked,
EWS is being retired). Recurring meetings are expanded to one entry per
occurrence, and re-runs perform an **incremental** add / update / cancel merge so
frequent scheduled refreshes stay cheap and never re-emit everything.

Single tool: `Export-Diary.ps1`. Windows PowerShell 5.1 compatible (also runs the
pure-logic tests under `pwsh` 7 on any OS).

---

## 10-second pre-flight (locked-down machine)

Run these two one-liners in the same Windows PowerShell session you'll use:

```powershell
# 1. Language mode MUST be FullLanguage (Constrained Language Mode blocks COM):
$ExecutionContext.SessionState.LanguageMode

# 2. Outlook COM must be creatable (Outlook should already be running):
$o = New-Object -ComObject Outlook.Application; $o.Session.CurrentUser.Name
```

If (1) prints anything other than `FullLanguage`, COM automation is disabled by
policy (see Troubleshooting). If (2) errors, start the Outlook desktop client
first.

`ExecutionPolicy` note: invoke the script with an explicit bypass rather than
changing machine policy:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Export-Diary.ps1 -Mailbox "delegate.name@yourorg.com"
```

---

## Requirements

- Windows with the **desktop Outlook client running** and signed in.
- **Delegate / reviewer access** to the target mailbox's calendar.
- Windows PowerShell 5.1 (built in) or PowerShell 7. Session in `FullLanguage`.
- No installation, no external modules.

The script **attaches** to the running Outlook instance and **never calls
`.Quit()`** — it will not close the user's Outlook.

---

## Usage

Default rolling window (today 00:00 -> today + 60 days), writes `.\diary.json`:

```powershell
.\Export-Diary.ps1 -Mailbox "delegate.name@yourorg.com"
```

Different horizon:

```powershell
.\Export-Diary.ps1 -Mailbox "delegate.name@yourorg.com" -DaysAhead 14
```

Explicit one-off range (both required together):

```powershell
.\Export-Diary.ps1 -Mailbox "delegate.name@yourorg.com" -From "2026-07-01" -To "2026-07-31"
```

Custom output file and include (truncated) plain-text bodies:

```powershell
.\Export-Diary.ps1 -Mailbox "delegate.name@yourorg.com" -OutFile "C:\diary\diary.json" -IncludeBody
```

Full rebuild (ignore the existing file, re-pull the whole window):

```powershell
.\Export-Diary.ps1 -Mailbox "delegate.name@yourorg.com" -FullResync
```

Run the built-in pure-logic tests and exit (no Outlook/COM touched):

```powershell
.\Export-Diary.ps1 -SelfTest
```

### Parameters

| Parameter      | Default             | Meaning                                                        |
|----------------|---------------------|----------------------------------------------------------------|
| `-Mailbox`     | (required)          | Delegate mailbox to read.                                       |
| `-DaysAhead`   | `60`                | Rolling window length; window = today 00:00 -> today+N.        |
| `-From` `-To`  | (none)              | Explicit range override (supply **both** or neither).          |
| `-OutFile`     | `.\diary.json`      | Output JSON path (written atomically; directory auto-created).  |
| `-IncludeBody` | off                 | Adds a `description` field (plain text, capped at 2000 chars). |
| `-FullResync`  | off                 | Ignore existing file; rebuild the window from scratch.         |
| `-SelfTest`    | off                 | Run merge-logic tests, exit with the failure count.            |

---

## Scheduling frequent refresh (Task Scheduler)

Create a task that runs while the user is logged on (Outlook must be running):

- **Program/script:** `powershell.exe`
- **Add arguments:**

  ```
  -NoProfile -ExecutionPolicy Bypass -File "C:\Tools\OutlookDiaryExtractor\Export-Diary.ps1" -Mailbox "delegate.name@yourorg.com" -DaysAhead 60 -OutFile "C:\Tools\OutlookDiaryExtractor\diary.json"
  ```

- **Trigger:** e.g. every 15 minutes.
- Run **only when user is logged on** (COM needs the interactive Outlook session).

Or register from a prompt:

```powershell
$taskArgs = '-NoProfile -ExecutionPolicy Bypass -File "C:\Tools\OutlookDiaryExtractor\Export-Diary.ps1" -Mailbox "delegate.name@yourorg.com" -DaysAhead 60'
schtasks /Create /TN "OutlookDiary" /SC MINUTE /MO 15 /TR "powershell.exe $taskArgs" /RL LIMITED /F
```

Each run does a small, bounded COM query and an incremental merge, so it stays
fast even at high frequency.

---

## Output JSON

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
| `description`       | string          | Present only with `-IncludeBody`; plain text, max 2000 chars.        |

### Incremental merge rules

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
- `-FullResync` ignores the existing file and rebuilds the window.

The file is written atomically (`diary.json.tmp` then `Move-Item -Force`), so a
downstream reader never sees a torn file. Encoding is UTF-8 without BOM.

---

## Testing

The pure merge/diff and conversion logic has **no COM dependency** and is tested
directly:

```powershell
pwsh -File tests/Merge.Tests.ps1        # any OS
powershell -File tests\Merge.Tests.ps1  # Windows PowerShell 5.1
.\Export-Diary.ps1 -SelfTest            # runs the same tests on the target box
```

The tests dot-source `Export-Diary.ps1` in **library mode** (they set
`$DiaryLibraryMode = $true` first) so only the functions load — the COM/export
path never runs. Exit code equals the number of failed assertions.

---

## Troubleshooting

**`LanguageMode` is `ConstrainedLanguage`.**
COM automation is blocked by policy (AppLocker / WDAC). `New-Object -ComObject`
will fail. There is no script-side workaround; this must be relaxed for the
account, or the tool run under an allowed context. Confirm with the pre-flight
one-liner.

**"Could not resolve mailbox ..."**
`CreateRecipient(...).Resolve()` failed. Check the address/name is exactly right
and resolves in Outlook's address book (try typing it into a new message's To
field). A UPN/email or a full display name usually works.

**"Access denied or unavailable for the shared calendar ..."**
`GetSharedDefaultFolder` failed. You need delegate or at least **Reviewer**
permission on the mailbox's calendar, and the mailbox owner's account must be
reachable. Ask the owner to share their Calendar, then retry.

**"Could not attach to Outlook ..."**
The desktop Outlook client is not running (or a policy blocked COM). Start
Outlook, wait for it to finish loading, and re-run.

**Wrong / no events, or a `Restrict` error about dates.**
The `Restrict` filter uses the machine's current culture short date/time. If the
regional format is unusual, verify the window bounds printed at the top of the
run look correct.
