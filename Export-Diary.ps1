<#
.SYNOPSIS
    Extract a delegate mailbox's Outlook calendar to JSON via desktop Outlook COM,
    with incremental (add / update / cancel) refresh.

.DESCRIPTION
    Attaches to the already-running desktop Outlook instance, resolves a shared
    (delegate) calendar, expands recurring appointments inside a bounded window,
    and merges the result into a JSON file for a downstream AI "Diary Assistant".

    Windows PowerShell 5.1 compatible (no PS7-only syntax). The pure merge/diff
    and conversion logic contains no COM calls and can be dot-sourced and unit
    tested anywhere (see tests/Merge.Tests.ps1). All COM access lives inside
    Get-DiaryEventsFromOutlook / Invoke-DiaryExport, which are only called on the
    real run path -- so dot-sourcing this file never touches COM.

.NOTES
    Library-mode guard: the main run body is wrapped in `if (-not $DiaryLibraryMode)`.
    Tests set `$DiaryLibraryMode = $true` in their own scope BEFORE dot-sourcing this
    file; because dot-sourcing executes in the caller's scope, the guard sees $true and
    the main body is skipped (functions are still defined). Running the script normally
    leaves $DiaryLibraryMode undefined ($null), so the guard runs the export.
#>
param(
    [string]$Mailbox = "",
    [int]$DaysAhead = 60,
    [string]$From = "",
    [string]$To = "",
    [string]$OutFile = ".\diary.json",
    [switch]$IncludeBody,
    [switch]$FullResync,
    [switch]$SelfTest
)

# ============================================================================
#  PURE HELPERS (no COM) -- safe to dot-source / unit test
# ============================================================================

function ConvertTo-Utc8601 {
    # UTC ISO-8601 with trailing Z, e.g. 2026-07-03T12:00:00Z
    param([datetime]$Value)
    return $Value.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss'Z'")
}

function Split-DiaryList {
    # Split a delimited string into a trimmed, empty-dropped array.
    param(
        [string]$Value,
        [string]$Delimiter
    )
    if ([string]::IsNullOrWhiteSpace($Value)) { return @() }
    $parts = $Value -split [regex]::Escape($Delimiter)
    $out = New-Object System.Collections.ArrayList
    foreach ($p in $parts) {
        $t = $p.Trim()
        if ($t.Length -gt 0) { [void]$out.Add($t) }
    }
    return $out.ToArray()
}

function Split-DiaryNames {
    # Outlook RequiredAttendees / OptionalAttendees are semicolon-joined display names.
    param([string]$Value)
    return (Split-DiaryList -Value $Value -Delimiter ';')
}

function ConvertTo-DiaryEvent {
    # Build one canonical diary event object from scalar values.
    # PURE: no COM here -- callers read COM properties and pass them in.
    param(
        [string]$Subject,
        [datetime]$Start,
        [datetime]$End,
        [bool]$AllDay,
        [string]$Organizer,
        [string]$Required,
        [string]$Optional,
        [string]$Location,
        [string]$Categories,
        [string]$GlobalId,
        [datetime]$LastModified,
        [int]$MeetingStatus,
        [bool]$IsRecurring,
        [string]$Description = $null
    )

    $startUtcIso = ConvertTo-Utc8601 -Value $Start
    $id = "$GlobalId|$startUtcIso"

    # olMeetingCanceled = 5, olMeetingReceivedAndCanceled = 7
    $status = "active"
    if ($MeetingStatus -eq 5 -or $MeetingStatus -eq 7) { $status = "cancelled" }

    $obj = [pscustomobject][ordered]@{
        id                = $id
        subject           = $Subject
        startDate         = $Start.ToString("yyyy-MM-dd")
        startTime         = $Start.ToString("HH:mm")
        endDate           = $End.ToString("yyyy-MM-dd")
        endTime           = $End.ToString("HH:mm")
        start             = $Start.ToString("yyyy-MM-ddTHH:mm:sszzz")
        end               = $End.ToString("yyyy-MM-ddTHH:mm:sszzz")
        organizer         = $Organizer
        requiredAttendees = @(Split-DiaryNames -Value $Required)
        optionalAttendees = @(Split-DiaryNames -Value $Optional)
        location          = $Location
        categories        = @(Split-DiaryList -Value $Categories -Delimiter ',')
        isRecurring       = [bool]$IsRecurring
        isAllDay          = [bool]$AllDay
        status            = $status
        lastModified      = (ConvertTo-Utc8601 -Value $LastModified)
        cancelledAt       = $null
    }

    # Only attach a description when body extraction was requested.
    if ($PSBoundParameters.ContainsKey('Description') -and $null -ne $Description) {
        $body = $Description
        if ($body.Length -gt 2000) { $body = $body.Substring(0, 2000) }
        Add-Member -InputObject $obj -MemberType NoteProperty -Name 'description' -Value $body
    }

    return $obj
}

function Set-DiaryValue {
    # Set a property on a PSCustomObject, adding it (NoteProperty) if absent.
    # Existing diary events loaded from older-schema or hand-edited JSON may lack
    # fields (e.g. cancelledAt); ConvertFrom-Json objects do NOT auto-create a
    # property on assignment -- they throw -- so add it when missing.
    param($Object, [string]$Name, $Value)
    if ($null -ne $Object.PSObject.Properties[$Name]) {
        $Object.$Name = $Value
    } else {
        Add-Member -InputObject $Object -MemberType NoteProperty -Name $Name -Value $Value -Force
    }
}

function Merge-DiaryEvents {
    # PURE incremental merge. No COM, no file IO.
    #   $ExistingEvents : events already in diary.json (array of PSCustomObject)
    #   $PulledEvents   : events freshly pulled for the window (array of PSCustomObject)
    #   $WindowFrom/$WindowTo : window bounds (datetime); compared in UTC
    #   $NowUtc         : timestamp used when stamping cancelledAt (datetime)
    # Returns object with .Events (array) + .Added/.Updated/.Cancelled/.Unchanged counts.
    param(
        [object[]]$ExistingEvents,
        [object[]]$PulledEvents,
        [datetime]$WindowFrom,
        [datetime]$WindowTo,
        [datetime]$NowUtc
    )

    if ($null -eq $ExistingEvents) { $ExistingEvents = @() }
    if ($null -eq $PulledEvents)   { $PulledEvents   = @() }

    $nowIso  = ConvertTo-Utc8601 -Value $NowUtc
    $fromUtc = $WindowFrom.ToUniversalTime()
    $toUtc   = $WindowTo.ToUniversalTime()

    $existingById = @{}
    foreach ($e in $ExistingEvents) { $existingById[$e.id] = $e }
    $pulledById = @{}
    foreach ($p in $PulledEvents) { $pulledById[$p.id] = $p }

    $merged    = New-Object System.Collections.ArrayList
    $added     = 0
    $updated   = 0
    $cancelled = 0
    $unchanged = 0

    # ---- pass 1: everything present in the fresh pull ----
    foreach ($p in $PulledEvents) {
        if (-not $existingById.ContainsKey($p.id)) {
            # brand new occurrence (may itself already be cancelled)
            [void]$merged.Add($p)
            $added++
            continue
        }

        $ex = $existingById[$p.id]

        if ($p.status -eq "cancelled") {
            if ($ex.status -eq "active") {
                # active -> cancelled this run (host cancelled the meeting)
                $p.cancelledAt = $nowIso
                [void]$merged.Add($p)
                $cancelled++
            } else {
                # was already cancelled -- carry forward, DO NOT re-stamp cancelledAt
                $p.cancelledAt = $ex.cancelledAt
                [void]$merged.Add($p)
                $unchanged++
            }
            continue
        }

        # pulled event is active
        if ($ex.status -eq "active" -and $ex.lastModified -eq $p.lastModified) {
            # unchanged -- preserve existing identity exactly
            [void]$merged.Add($ex)
            $unchanged++
        } else {
            # content changed (or reactivated) -> update; active means cancelledAt null
            $p.cancelledAt = $null
            [void]$merged.Add($p)
            $updated++
        }
    }

    # ---- pass 2: existing events NOT in the fresh pull ----
    foreach ($e in $ExistingEvents) {
        if ($pulledById.ContainsKey($e.id)) { continue }

        # "In window" must mirror the pull's Restrict filter exactly
        # ([Start] >= from AND [End] <= to). An event whose start is inside but
        # whose end extends past windowTo is NOT in the pull, so its absence
        # proves nothing -- treating it as in-window would falsely cancel it.
        $inWindow = $false
        if ($null -ne $e.start -and $e.start -ne "") {
            try {
                $startUtc = [datetimeoffset]::Parse($e.start, [System.Globalization.CultureInfo]::InvariantCulture).UtcDateTime
                $endUtc = $startUtc
                if ($null -ne $e.end -and $e.end -ne "") {
                    $endUtc = [datetimeoffset]::Parse($e.end, [System.Globalization.CultureInfo]::InvariantCulture).UtcDateTime
                }
                if ($startUtc -ge $fromUtc -and $endUtc -le $toUtc) { $inWindow = $true }
            } catch {
                $inWindow = $false
            }
        }

        if ($e.status -eq "active" -and $inWindow) {
            # was active, sits inside the refreshed window, but vanished from the
            # pull -> it was deleted/cancelled in Outlook -> flip to cancelled.
            Set-DiaryValue -Object $e -Name 'status' -Value 'cancelled'
            Set-DiaryValue -Object $e -Name 'cancelledAt' -Value $nowIso
            [void]$merged.Add($e)
            $cancelled++
        } else {
            # outside window, or already cancelled -> carry forward untouched.
            [void]$merged.Add($e)
            $unchanged++
        }
    }

    return [pscustomobject]@{
        Events    = $merged.ToArray()
        Added     = $added
        Updated   = $updated
        Cancelled = $cancelled
        Unchanged = $unchanged
    }
}

# ============================================================================
#  COM LAYER -- only ever reached from Invoke-DiaryExport (real run path)
# ============================================================================

function Release-ComObject {
    param($ComObject)
    if ($null -ne $ComObject) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) } catch { }
    }
}

function Get-DiaryEventsFromOutlook {
    # Attach to running Outlook, resolve the shared calendar, expand recurrences
    # inside [WindowFrom, WindowTo], and return an array of plain diary events.
    param(
        [string]$Mailbox,
        [datetime]$WindowFrom,
        [datetime]$WindowTo,
        [switch]$IncludeBody
    )

    $outlook    = $null
    $ns         = $null
    $recipient  = $null
    $calendar   = $null
    $items      = $null
    $restricted = $null

    try {
        try {
            $outlook = New-Object -ComObject Outlook.Application
        } catch {
            throw "Could not attach to Outlook. Ensure the desktop Outlook client is running and this session is not in Constrained Language Mode. Underlying error: $($_.Exception.Message)"
        }

        $ns = $outlook.GetNamespace("MAPI")

        $recipient = $ns.CreateRecipient($Mailbox)
        [void]$recipient.Resolve()
        if (-not $recipient.Resolved) {
            throw "Could not resolve mailbox '$Mailbox'. Check the address/name and that it exists in the address book."
        }

        try {
            # 9 = olFolderCalendar
            $calendar = $ns.GetSharedDefaultFolder($recipient, 9)
        } catch {
            throw "Access denied or unavailable for the shared calendar of '$Mailbox'. Confirm delegate/reviewer permission has been granted and the mailbox is reachable. Underlying error: $($_.Exception.Message)"
        }

        $items = $calendar.Items

        # ---- canonical fast recurrence pattern -- ORDER IS MANDATORY ----
        $items.Sort("[Start]")
        $items.IncludeRecurrences = $true
        # Restrict date strings must use current-culture short date/time (.ToString("g")).
        $filter = "[Start] >= '" + $WindowFrom.ToString("g") + "' AND [End] <= '" + $WindowTo.ToString("g") + "'"
        $restricted = $items.Restrict($filter)

        $results = New-Object System.Collections.ArrayList

        # Never use .Count / indexing on an IncludeRecurrences collection.
        $item = $restricted.GetFirst()
        while ($null -ne $item) {

            # Read Start first so we can hard-break once we pass the window end.
            $sStart = $item.Start
            if ($sStart -gt $WindowTo) {
                $tail = $item
                $item = $null
                Release-ComObject $tail
                break
            }

            # Each property access is an RPC -- read ONLY what we need, once each.
            $sSubject = $item.Subject
            $sEnd     = $item.End
            $sAllDay  = $item.AllDayEvent
            $sOrg     = $item.Organizer
            $sReq     = $item.RequiredAttendees
            $sOpt     = $item.OptionalAttendees
            $sLoc     = $item.Location
            $sCat     = $item.Categories
            $sGid     = $item.GlobalAppointmentID
            $sLm      = $item.LastModificationTime
            $sMs      = $item.MeetingStatus
            $sRec     = $item.IsRecurring
            $sBody    = $null
            if ($IncludeBody) { $sBody = $item.Body }

            if ($IncludeBody) {
                $evt = ConvertTo-DiaryEvent -Subject $sSubject -Start $sStart -End $sEnd -AllDay $sAllDay `
                    -Organizer $sOrg -Required $sReq -Optional $sOpt -Location $sLoc -Categories $sCat `
                    -GlobalId $sGid -LastModified $sLm -MeetingStatus $sMs -IsRecurring $sRec -Description $sBody
            } else {
                $evt = ConvertTo-DiaryEvent -Subject $sSubject -Start $sStart -End $sEnd -AllDay $sAllDay `
                    -Organizer $sOrg -Required $sReq -Optional $sOpt -Location $sLoc -Categories $sCat `
                    -GlobalId $sGid -LastModified $sLm -MeetingStatus $sMs -IsRecurring $sRec
            }

            [void]$results.Add($evt)

            $next = $restricted.GetNext()
            Release-ComObject $item
            $item = $next
        }

        return $results.ToArray()
    }
    finally {
        Release-ComObject $restricted
        Release-ComObject $items
        Release-ComObject $calendar
        Release-ComObject $recipient
        Release-ComObject $ns
        # NEVER call $outlook.Quit() -- we attached to the user's running instance.
        Release-ComObject $outlook
    }
}

# ============================================================================
#  ORCHESTRATION (real run path)
# ============================================================================

function Get-DiaryWindow {
    # Resolve [From,To] from explicit args or the rolling DaysAhead default.
    param(
        [string]$From,
        [string]$To,
        [int]$DaysAhead
    )
    if ($From -ne "" -and $To -ne "") {
        return [pscustomobject]@{
            From = [datetime]::Parse($From)
            To   = [datetime]::Parse($To)
        }
    }
    if ($From -ne "" -or $To -ne "") {
        throw "Specify BOTH -From and -To (an explicit range), or neither (use -DaysAhead)."
    }
    $start = (Get-Date).Date
    return [pscustomobject]@{
        From = $start
        To   = $start.AddDays($DaysAhead)
    }
}

function Read-ExistingDiary {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return @() }
        $parsed = $raw | ConvertFrom-Json
        if ($null -ne $parsed.events) { return @($parsed.events) }
        return @()
    } catch {
        Write-Warning "Could not parse existing '$Path' ($($_.Exception.Message)). Treating as empty."
        return @()
    }
}

function Write-DiaryFile {
    param(
        [string]$Path,
        [string]$Mailbox,
        [datetime]$WindowFrom,
        [datetime]$WindowTo,
        [object[]]$Events
    )
    $meta = [ordered]@{
        mailbox    = $Mailbox
        lastRunUtc = (ConvertTo-Utc8601 -Value (Get-Date))
        windowFrom = $WindowFrom.ToString("yyyy-MM-ddTHH:mm:sszzz")
        windowTo   = $WindowTo.ToString("yyyy-MM-ddTHH:mm:sszzz")
    }
    $doc = [ordered]@{
        meta   = $meta
        events = @($Events)
    }
    $json = $doc | ConvertTo-Json -Depth 6

    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        [void](New-Item -ItemType Directory -Path $dir -Force)
    }

    $tmp = "$Path.tmp"
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($tmp, $json, $utf8NoBom)
    Move-Item -LiteralPath $tmp -Destination $Path -Force
}

function Invoke-DiaryExport {
    # The real run path. This is the ONLY function that reaches COM.
    param(
        [string]$Mailbox,
        [int]$DaysAhead,
        [string]$From,
        [string]$To,
        [string]$OutFile,
        [switch]$IncludeBody,
        [switch]$FullResync
    )

    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    # Resolve to an absolute path once: [IO.File]::WriteAllText resolves relative
    # paths against the process CWD, which can differ from $PWD (e.g. Task Scheduler).
    $OutFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutFile)

    $window = Get-DiaryWindow -From $From -To $To -DaysAhead $DaysAhead
    $windowFrom = $window.From
    $windowTo   = $window.To

    Write-Host "Diary window: $($windowFrom.ToString('yyyy-MM-dd HH:mm')) -> $($windowTo.ToString('yyyy-MM-dd HH:mm'))  (mailbox: $Mailbox)"

    $existing = @()
    if ($FullResync) {
        Write-Host "Full resync: existing diary ignored."
    } else {
        $existing = Read-ExistingDiary -Path $OutFile
    }

    $pulled = Get-DiaryEventsFromOutlook -Mailbox $Mailbox -WindowFrom $windowFrom -WindowTo $windowTo -IncludeBody:$IncludeBody

    $merge = Merge-DiaryEvents -ExistingEvents $existing -PulledEvents $pulled `
        -WindowFrom $windowFrom -WindowTo $windowTo -NowUtc (Get-Date).ToUniversalTime()

    $ordered = @($merge.Events | Sort-Object -Property start)

    Write-DiaryFile -Path $OutFile -Mailbox $Mailbox -WindowFrom $windowFrom -WindowTo $windowTo -Events $ordered

    $sw.Stop()
    $elapsed = [math]::Round($sw.Elapsed.TotalSeconds, 2)

    Write-Host ""
    Write-Host "Diary written to $OutFile"
    Write-Host ("  added:     {0}" -f $merge.Added)
    Write-Host ("  updated:   {0}" -f $merge.Updated)
    Write-Host ("  cancelled: {0}" -f $merge.Cancelled)
    Write-Host ("  unchanged: {0}" -f $merge.Unchanged)
    Write-Host ("  total:     {0}" -f $ordered.Count)
    Write-Host ("  elapsed:   {0}s" -f $elapsed)
}

# ============================================================================
#  MAIN -- guarded so dot-sourcing (tests) never runs COM or the export
# ============================================================================

if (-not $DiaryLibraryMode) {

    if ($SelfTest) {
        $testPath = Join-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath 'tests') -ChildPath 'Merge.Tests.ps1'
        if (-not (Test-Path -LiteralPath $testPath)) {
            Write-Error "Self-test file not found at '$testPath'."
            exit 3
        }
        & $testPath
        exit $LASTEXITCODE
    }

    # Fail fast rather than [Parameter(Mandatory)]: a mandatory prompt would
    # hang forever under a non-interactive scheduled task.
    if ([string]::IsNullOrWhiteSpace($Mailbox)) {
        Write-Error "-Mailbox is required: the delegate mailbox whose calendar to extract, e.g. -Mailbox 'first.last@yourorg.com'."
        exit 2
    }

    try {
        Invoke-DiaryExport -Mailbox $Mailbox -DaysAhead $DaysAhead -From $From -To $To `
            -OutFile $OutFile -IncludeBody:$IncludeBody -FullResync:$FullResync
    } catch {
        Write-Error $_.Exception.Message
        exit 1
    }
}
