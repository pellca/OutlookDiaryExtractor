<#
    Plain-PowerShell unit tests for the pure logic in Export-Diary.ps1.
    No Pester required. Runs on Windows PowerShell 5.1 and pwsh 7 (Linux/Windows).

    Usage:
        pwsh -File tests/Merge.Tests.ps1
        powershell -File tests\Merge.Tests.ps1
    or, on the target machine:
        .\Export-Diary.ps1 -SelfTest

    Exit code = number of failed assertions (0 = all green).

    Library-mode guard: we set $DiaryLibraryMode = $true BEFORE dot-sourcing
    Export-Diary.ps1. Because dot-sourcing executes in THIS scope, the guarded
    main body of Export-Diary.ps1 (`if (-not $DiaryLibraryMode) { ... }`) sees
    $true and is skipped -- so only the functions load, never the COM/export path.
#>

$DiaryLibraryMode = $true
. (Join-Path -Path (Split-Path -Parent $PSScriptRoot) -ChildPath 'Export-Diary.ps1')

# ---------------------------------------------------------------------------
#  tiny assert harness
# ---------------------------------------------------------------------------
$script:Total = 0
$script:Fail  = 0

function Assert-True {
    param([bool]$Condition, [string]$Message)
    $script:Total++
    if ($Condition) {
        Write-Host "  PASS  $Message"
    } else {
        Write-Host "  FAIL  $Message"
        $script:Fail++
    }
}

function Assert-Equal {
    param($Expected, $Actual, [string]$Message)
    Assert-True (($Expected) -eq ($Actual)) ("{0} (expected '{1}', got '{2}')" -f $Message, $Expected, $Actual)
}

function New-TestEvent {
    param(
        [string]$Id,
        [string]$Start,
        [string]$End = "",
        [string]$Status = "active",
        [string]$LastModified = "2026-07-01T00:00:00Z",
        $CancelledAt = $null,
        [string]$Subject = "Test"
    )
    return [pscustomobject]@{
        id           = $Id
        subject      = $Subject
        start        = $Start
        end          = $End
        status       = $Status
        lastModified = $LastModified
        cancelledAt  = $CancelledAt
    }
}

# Deterministic UTC window / clock (independent of the runner's timezone).
$WFrom  = [datetime]::Parse("2026-07-01T00:00:00Z").ToUniversalTime()
$WTo    = [datetime]::Parse("2026-08-30T00:00:00Z").ToUniversalTime()
$NowUtc = [datetime]::Parse("2026-07-03T12:00:00Z").ToUniversalTime()
$NowIso = "2026-07-03T12:00:00Z"

$InWindowStart  = "2026-07-10T09:00:00+00:00"
$OutWindowStart = "2026-05-01T09:00:00+00:00"

Write-Host "Running diary merge-logic tests..."
Write-Host ""

# ---------------------------------------------------------------------------
Write-Host "[ add new event ]"
$existing = @()
$pulled   = @( (New-TestEvent -Id "A" -Start $InWindowStart) )
$r = Merge-DiaryEvents -ExistingEvents $existing -PulledEvents $pulled -WindowFrom $WFrom -WindowTo $WTo -NowUtc $NowUtc
Assert-Equal 1 $r.Added     "new event counted as added"
Assert-Equal 0 $r.Updated   "no updates"
Assert-Equal 1 @($r.Events).Count "one event in output"
Assert-Equal "active" @($r.Events)[0].status "new event active"

# ---------------------------------------------------------------------------
Write-Host "[ unchanged event skipped / identity preserved ]"
$existing = @( (New-TestEvent -Id "A" -Start $InWindowStart -LastModified "2026-07-01T00:00:00Z" -Subject "ORIGINAL") )
# Same id + same lastModified but different subject in the pull -> must keep EXISTING.
$pulled   = @( (New-TestEvent -Id "A" -Start $InWindowStart -LastModified "2026-07-01T00:00:00Z" -Subject "CHANGED") )
$r = Merge-DiaryEvents -ExistingEvents $existing -PulledEvents $pulled -WindowFrom $WFrom -WindowTo $WTo -NowUtc $NowUtc
Assert-Equal 1 $r.Unchanged "unchanged counted"
Assert-Equal 0 $r.Updated   "not updated"
Assert-Equal "ORIGINAL" @($r.Events)[0].subject "existing identity preserved (kept ORIGINAL)"

# ---------------------------------------------------------------------------
Write-Host "[ modified event updated ]"
$existing = @( (New-TestEvent -Id "A" -Start $InWindowStart -LastModified "2026-07-01T00:00:00Z" -Subject "ORIGINAL") )
$pulled   = @( (New-TestEvent -Id "A" -Start $InWindowStart -LastModified "2026-07-02T09:30:00Z" -Subject "CHANGED") )
$r = Merge-DiaryEvents -ExistingEvents $existing -PulledEvents $pulled -WindowFrom $WFrom -WindowTo $WTo -NowUtc $NowUtc
Assert-Equal 1 $r.Updated   "changed lastModified counted as updated"
Assert-Equal "CHANGED" @($r.Events)[0].subject "updated to pulled content"
Assert-Equal $null @($r.Events)[0].cancelledAt "active update keeps cancelledAt null"

# ---------------------------------------------------------------------------
Write-Host "[ missing from pull inside window -> cancelled + stamped ]"
$existing = @( (New-TestEvent -Id "A" -Start $InWindowStart -Status "active") )
$pulled   = @()
$r = Merge-DiaryEvents -ExistingEvents $existing -PulledEvents $pulled -WindowFrom $WFrom -WindowTo $WTo -NowUtc $NowUtc
Assert-Equal 1 $r.Cancelled "in-window absence counted as cancelled"
Assert-Equal "cancelled" @($r.Events)[0].status "status flipped to cancelled"
Assert-Equal $NowIso @($r.Events)[0].cancelledAt "cancelledAt stamped with now"

# ---------------------------------------------------------------------------
Write-Host "[ older-schema existing event lacking cancelledAt -> cancels without throwing ]"
# Events from an older schema or a hand-edited file may have no cancelledAt
# property. ConvertFrom-Json objects throw on assignment to a missing property,
# so pass-2 must add it. Build via ConvertFrom-Json to reproduce faithfully.
$legacyJson = '{"events":[{"id":"A","status":"active","start":"' + $InWindowStart + '","lastModified":"2026-07-01T00:00:00Z"}]}'
$existing = @(($legacyJson | ConvertFrom-Json).events)
$pulled   = @()
$threw = $false
try {
    $r = Merge-DiaryEvents -ExistingEvents $existing -PulledEvents $pulled -WindowFrom $WFrom -WindowTo $WTo -NowUtc $NowUtc
} catch { $threw = $true }
Assert-True (-not $threw) "merge does not throw when existing event lacks cancelledAt"
Assert-Equal 1 $r.Cancelled "legacy in-window absence counted as cancelled"
Assert-Equal "cancelled" @($r.Events)[0].status "legacy event flipped to cancelled"
Assert-Equal $NowIso @($r.Events)[0].cancelledAt "cancelledAt added and stamped on legacy event"

# ---------------------------------------------------------------------------
Write-Host "[ already-cancelled event not re-stamped ]"
$existing = @( (New-TestEvent -Id "A" -Start $InWindowStart -Status "cancelled" -CancelledAt "2026-06-15T08:00:00Z") )
$pulled   = @()
$r = Merge-DiaryEvents -ExistingEvents $existing -PulledEvents $pulled -WindowFrom $WFrom -WindowTo $WTo -NowUtc $NowUtc
Assert-Equal 0 $r.Cancelled "already-cancelled not re-counted"
Assert-Equal "2026-06-15T08:00:00Z" @($r.Events)[0].cancelledAt "original cancelledAt preserved (not restamped)"

# ---------------------------------------------------------------------------
Write-Host "[ event outside window untouched ]"
$existing = @( (New-TestEvent -Id "A" -Start $OutWindowStart -Status "active") )
$pulled   = @()
$r = Merge-DiaryEvents -ExistingEvents $existing -PulledEvents $pulled -WindowFrom $WFrom -WindowTo $WTo -NowUtc $NowUtc
Assert-Equal 0 $r.Cancelled "out-of-window event not cancelled"
Assert-Equal "active" @($r.Events)[0].status "out-of-window event stays active"
Assert-Equal $null @($r.Events)[0].cancelledAt "out-of-window cancelledAt still null"

# ---------------------------------------------------------------------------
Write-Host "[ event straddling window end not falsely cancelled ]"
# The pull's Restrict filter requires [End] <= windowTo, so an event whose start
# is inside the window but whose end extends past it is NOT in the pull. Its
# absence proves nothing -- it must be carried forward active, not cancelled.
$existing = @( (New-TestEvent -Id "A" -Start "2026-08-29T23:00:00+00:00" -End "2026-08-30T02:00:00+00:00" -Status "active") )
$pulled   = @()
$r = Merge-DiaryEvents -ExistingEvents $existing -PulledEvents $pulled -WindowFrom $WFrom -WindowTo $WTo -NowUtc $NowUtc
Assert-Equal 0 $r.Cancelled "straddling event not counted cancelled"
Assert-Equal "active" @($r.Events)[0].status "straddling event stays active"
Assert-Equal $null @($r.Events)[0].cancelledAt "straddling event cancelledAt still null"

# ---------------------------------------------------------------------------
Write-Host "[ MeetingStatus-cancelled event in pull -> cancelled + stamped ]"
$existing = @( (New-TestEvent -Id "A" -Start $InWindowStart -Status "active" -LastModified "2026-07-01T00:00:00Z") )
$pulled   = @( (New-TestEvent -Id "A" -Start $InWindowStart -Status "cancelled" -LastModified "2026-07-02T00:00:00Z") )
$r = Merge-DiaryEvents -ExistingEvents $existing -PulledEvents $pulled -WindowFrom $WFrom -WindowTo $WTo -NowUtc $NowUtc
Assert-Equal 1 $r.Cancelled "pulled-cancelled active event counted as cancelled"
Assert-Equal "cancelled" @($r.Events)[0].status "status cancelled from pull"
Assert-Equal $NowIso @($r.Events)[0].cancelledAt "cancelledAt stamped on flip"

# ---------------------------------------------------------------------------
Write-Host "[ new pulled event already cancelled -> added as cancelled ]"
$existing = @()
$pulled   = @( (New-TestEvent -Id "Z" -Start $InWindowStart -Status "cancelled") )
$r = Merge-DiaryEvents -ExistingEvents $existing -PulledEvents $pulled -WindowFrom $WFrom -WindowTo $WTo -NowUtc $NowUtc
Assert-Equal 1 $r.Added "new cancelled event added"
Assert-Equal "cancelled" @($r.Events)[0].status "new event carries cancelled status"

# ---------------------------------------------------------------------------
Write-Host "[ full-resync semantics: empty existing -> all added ]"
# -FullResync ignores the existing file; the pure function models that as empty existing.
$existing = @()
$pulled   = @(
    (New-TestEvent -Id "A" -Start $InWindowStart),
    (New-TestEvent -Id "B" -Start $InWindowStart)
)
$r = Merge-DiaryEvents -ExistingEvents $existing -PulledEvents $pulled -WindowFrom $WFrom -WindowTo $WTo -NowUtc $NowUtc
Assert-Equal 2 $r.Added     "full resync adds everything"
Assert-Equal 0 $r.Unchanged "full resync has no carry-over"

# ---------------------------------------------------------------------------
Write-Host "[ ConvertTo-DiaryEvent: all-day event field shape ]"
$d1 = [datetime]::Parse("2026-07-10T00:00:00Z")
$d2 = [datetime]::Parse("2026-07-11T00:00:00Z")
$lm = [datetime]::Parse("2026-07-01T00:00:00Z")
$ev = ConvertTo-DiaryEvent -Subject "Holiday" -Start $d1 -End $d2 -AllDay $true `
    -Organizer "Someone" -Required "" -Optional "" -Location "" -Categories "" `
    -GlobalId "GID123" -LastModified $lm -MeetingStatus 1 -IsRecurring $false
Assert-True ($ev.isAllDay -eq $true)  "isAllDay is true"
Assert-True ($ev.isRecurring -eq $false) "isRecurring is false"
Assert-Equal "active" $ev.status "non-cancelled MeetingStatus -> active"
Assert-True ($ev.id -like "GID123|*Z") "id = GlobalAppointmentID|<startUtc Z>"
Assert-Equal 0 @($ev.requiredAttendees).Count "empty attendees -> empty array"
Assert-Equal 0 @($ev.categories).Count "empty categories -> empty array"
Assert-True (($ev.requiredAttendees -is [array]) -or (@($ev.requiredAttendees).Count -eq 0)) "requiredAttendees is array shaped"

# ---------------------------------------------------------------------------
Write-Host "[ ConvertTo-DiaryEvent: MeetingStatus 5 and 7 -> cancelled ]"
$ev5 = ConvertTo-DiaryEvent -Subject "x" -Start $d1 -End $d2 -AllDay $false -Organizer "" -Required "" -Optional "" -Location "" -Categories "" -GlobalId "G5" -LastModified $lm -MeetingStatus 5 -IsRecurring $false
$ev7 = ConvertTo-DiaryEvent -Subject "x" -Start $d1 -End $d2 -AllDay $false -Organizer "" -Required "" -Optional "" -Location "" -Categories "" -GlobalId "G7" -LastModified $lm -MeetingStatus 7 -IsRecurring $false
Assert-Equal "cancelled" $ev5.status "MeetingStatus 5 -> cancelled"
Assert-Equal "cancelled" $ev7.status "MeetingStatus 7 -> cancelled"

# ---------------------------------------------------------------------------
Write-Host "[ attendee string splitting edge cases ]"
Assert-Equal 0 @(Split-DiaryNames -Value "").Count            "empty string -> 0 names"
Assert-Equal 0 @(Split-DiaryNames -Value $null).Count         "null -> 0 names"
Assert-Equal 2 @(Split-DiaryNames -Value "Alice Smith; Bob Jones").Count "two names"
Assert-Equal 2 @(Split-DiaryNames -Value "Alice; ; Bob;").Count "empties and trailing semicolons dropped"
$solo = @(Split-DiaryNames -Value "  Solo Person  ")
Assert-Equal 1 $solo.Count "single padded name -> 1"
Assert-Equal "Solo Person" $solo[0] "name is trimmed"
Assert-Equal 2 @(Split-DiaryList -Value "Red, Green ,, " -Delimiter ",").Count "categories split, trimmed, empties dropped"

# ---------------------------------------------------------------------------
Write-Host ""
if ($script:Fail -eq 0) {
    Write-Host ("ALL PASS  ({0} assertions)" -f $script:Total)
} else {
    Write-Host ("FAILURES: {0} of {1} assertions failed" -f $script:Fail, $script:Total)
}
exit $script:Fail
