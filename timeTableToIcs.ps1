<#
.SYNOPSIS
    Gets and converts a WebUntis timetable to an ICS calendar file.

.DESCRIPTION
    This script retrieves timetable data from the WebUntis API and converts it into a subscribable ICS calendar file format.
    It allows specifying a date range for the timetable data.

.PARAMETER baseUrl
    The base URL of the WebUntis API. This is required to connect to the WebUntis server.

.PARAMETER elementType
    Specifies the type of element to retrieve. Default is 1, which typically corresponds to a timetable.

.PARAMETER elementId
    The unique identifier for the class timetable. This ID is used to fetch the timetable data.

.PARAMETER dates
    An array of dates (either as strings or DateTime objects) for which to retrieve timetable data.
    Defaults to the previous, current, and next three weeks. The maximum range is defined by the WebUntis administrator.

.PARAMETER OutputFilePath
    The file path where the ICS file will be saved. Must have a .ics extension. Default is "calendar.ics".

.PARAMETER dontCreateMultiDayEvents
    If specified, the script will skip generating "summary" multi-day events.

.PARAMETER dontSplitOnGapDays
    If specified, multi-day events will not be split even if there are gap days between periods.

.PARAMETER overrideSummaries
    A hashtable to override course summaries. The key is the original (short) course name, and the value is the new course name.

.PARAMETER appendToPreviousICSat
    The path to an existing ICS file to which new timetable data should be appended.
    The file must have a .ics extension and be a valid ICS file.

.PARAMETER connectMaxGapMinutes
    Specifies the maximum gap in minutes between consecutive lessons to connect. Set to -1 to disable connection. Default is 15 minutes.

.PARAMETER dontCreateBreakEntriesForGaps
    If specified, gaps in consecutive events will not result in a (low-Prio) "Break" entry.

.PARAMETER splitByCourse
    If specified, timetable data will be split into separate ICS files for each course.

.PARAMETER splitByOverrides
    If specified, timetable data will be split into separate ICS files for each course defined in overrideSummaries and the remaining miscellaneous classes.

.PARAMETER dontSplitHigherPrio
    If specified, higher-priority lessons will not be moved into the separate ICS file "PRIO".

.PARAMETER groupByPrio
    If specified, higher-priority lessons will also be grouped into separate ICS files "PRIO($val)". Implies -dontSplitHigherPrio.
    File names for priorities can be overridden by overrideSummaries with the key "PRIO$val".
    (Example: -overrideSummaries @{PRIO8 = "MySpecialName"})

.PARAMETER dontRemoveHighPrioFromNormal
    If specified, higher-priority lessons will not be removed from the normal timetable(s).
    By default, higher-priority lessons are removed from the normal timetable(s) and placed in a separate ICS file "PRIO($val)".

.PARAMETER outAllFormats
    If specified, outputs the timetable data in all available formats. Implies -splitByCourse $true and -dontSplitHigherPrio $false.

.PARAMETER culture
    Specifies the culture info used for DST/TZ adjustments and formatting date/time values.
    Defaults to the system culture as returned by Get-Culture.

.PARAMETER TimeZoneID
    Specifies the time zone ID for the calendar entries in IANA format. Default is 'Europe/Berlin'.
    Other options include 'Europe/Paris', 'Europe/Rome', 'Europe/Madrid', 'Europe/Brussels', 'Europe/Warsaw', and 'UTC'.

.PARAMETER cookie
    The session cookie value for the WebUntis API. Required for authentication.

.PARAMETER tenantId
    The tenant ID for the WebUntis session. Required for authentication.

.NOTES
    Author: Chaos_02
    Date: 2025-05-15
    Version: 1.9.4
    This script is designed to work with the WebUntis API to generate ICS calendar files from timetable data.
#>

param (
    [ValidateNotNullOrEmpty()]
    [Alias('URL')]
    [string]$baseUrl,

    [int]$elementType = 1,

    [Alias('TimeTableID')]
    [int]$elementId,

    [Parameter(Mandatory = $false)]
    [Alias('DateRange')]
    [ValidateScript({
            if ($_.GetType().Name -eq 'String') {
                if (-not [datetime]::TryParse($_, [ref] $null)) {
                    throw 'Invalid date format. Please provide a valid date string parse-able by `[datetime]::TryParse()`.'
                }
            } elseif ($_.GetType().Name -ne 'DateTime') {
                throw 'Invalid date format. Provide a date string or DateTime[] object.'
            }
            $true
        })]
    [System.Object[]]$dates = (@(-7, 0, 7, 14, 21, 28) | ForEach-Object { return (Get-Date).Date.AddHours(12).AddDays($_) }),

    [switch]$dontCreateMultiDayEvents,

    [ValidateScript({ if (($_ -and -not $dontCreateMultiDayEvents) -eq $false) {throw "Can't use together with -dontCreateMultiDayEvents"} else {$true} })]
    [switch]$dontSplitOnGapDays,

    [Parameter(
        Position = 0,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true,
        HelpMessage = 'Specifies the output file path for the generated ICS file. Must have a .ics extension.'
    )]
    [Alias('PSPath')]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({
        if ($_.Extension.ToLower() -eq '.ics') {
            $true
        }
        else {
            throw "Output file must have a .ics extension."
        }
    })]
    [System.IO.FileInfo]$OutputFilePath = [System.IO.FileInfo]"calendar.ics",

    [ValidateNotNullOrEmpty()]
    [ValidateScript({
        if (-not ($_ -is [hashtable])) {
            throw "The parameter must be a hashtable."
        }
        foreach ($key in $_.Keys) {
            if (-not ($key -is [string]) -or [string]::IsNullOrWhiteSpace($key)) {
                throw "All keys in overrideSummaries must be non-empty strings. Invalid key: '$key'."
            }
            $value = $_[$key]
            if (-not ($value -is [string]) -or [string]::IsNullOrWhiteSpace($value)) {
                throw "All values in overrideSummaries must be non-empty strings. Invalid value for key '$key': '$value'."
            }
        }
        $true
    })]
    [Parameter(
        ValueFromPipelineByPropertyName = $true,
        HelpMessage = 'A hashtable to override the summaries of the courses. The key is the original course (short) name, the value is the new course name.'
    )]
    [hashtable]$overrideSummaries,

    [Parameter(
        ValueFromPipelineByPropertyName = $true,
        HelpMessage = 'Path to an existing ICS file to which the new timetable data should be appended. Must be a valid ICS file.'
    )]
    [ValidateNotNull()]
    [ValidateScript({
        if (-not $_.Exists) {
            throw "The file '$($_.FullName)' does not exist."
        }
        if ($_.Extension.ToLower() -ne ".ics") {
            throw "The file '$($_.FullName)' does not have a .ics extension."
        }
        $content = Get-Content -Path $_.FullName -Raw
        if ($content -notmatch '^BEGIN:VCALENDAR' -or $content -notmatch 'END:VCALENDAR\s*$') {
            throw "The file '$($_.FullName)' does not appear to be a valid ICS file."
        }
        $true
    })]
    [System.IO.FileInfo]$appendToPreviousICSat,

    [Parameter(HelpMessage = 'Specifies the maximum gap in minutes between consecutive lessons to connect. Set to -1 to disable connection. Default is 15 minutes.')]
    [ValidateScript({ if ($_ -ge -1) { $true } else { throw 'The value must be a non-negative integer.' } })]
    [int]$connectMaxGapMinutes = 15,

    [Parameter(HelpMessage = 'If set, gaps in consecutive events will result in a less prioritized "Break" entry.')]
    [switch]$dontCreateBreakEntriesForGaps,

    [Parameter(
        ParameterSetName = 'OutputControl',
        HelpMessage = 'Split the timetable data into separate ICS files for each course.'
    )]
    [switch]$splitByCourse,

    [Parameter(
        ParameterSetName = 'OutputControl',
        HelpMessage = 'Split only by courses defined in overrideSummaries and misc. classes.'
    )]
    [ValidateScript({if ($_ -and -not $overrideSummaries) {throw 'The parameter -splitByOverrides requires the parameter overrideSummaries to be set.'} else {$true}})]
    [switch]$splitByOverrides,

    [Parameter(
        ParameterSetName = 'OutputControl',
        HelpMessage = 'If set, higher-priority lessons will not be moved at least into the separate ICS file "PRIO".'
    )]
    [switch]$dontSplitHigherPrio,

    [Parameter(
        ParameterSetName = 'OutputControl',
        HelpMessage = 'If set, higher-priority lessons will not be removed from the normal timetable(s).'
    )]
    [switch]$dontRemoveHighPrioFromNormal,

    [Parameter(
        ParameterSetName = 'OutputControl',
        HelpMessage = 'If set, higher-priority lessons will also be grouped into separate ICS file(s) "PRIO($val)".'
    )]
    [ValidateScript({if ($_ -and $dontSplitHigherPrio) {throw 'The parameter -groupByPrio requires the parameter -dontSplitHigherPrio to be NOT set.'} else {$true}})]
    [switch]$groupByPrio,

    [Parameter(
        ParameterSetName = 'OutputControl',
        HelpMessage = 'If set, outputs the timetable data in all available formats.'
    )]
    [switch]$outAllFormats,

    [Parameter(HelpMessage = 'Specifies the culture info used for DST/TZ adjustments and formatting date/time values. Default is the system culture.')]
    [ArgumentCompleter({
        param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
        # Retrieve all available cultures (you can choose to restrict to a subset if desired)
        $allCultures = [System.Globalization.CultureInfo]::GetCultures([System.Globalization.CultureTypes]::AllCultures)

        if ($wordToComplete -eq '') {
            # If no input is provided, return all culture names - seems to be ignored because it cycles through files
            return $allCultures | ForEach-Object { $_.Name }
        }
        $allCultures |
            Where-Object { $_.Name -like "*$wordToComplete*" } |
            ForEach-Object {
                # Create a CompletionResult for each matching culture.
                # The first argument is the text inserted upon selection,
                # the second is the text displayed,
                # the third is the result type,
                # and the fourth is a tooltip (here we use the DisplayName).
                [System.Management.Automation.CompletionResult]::new(
                    $_.Name,
                    $_.Name,
                    'ParameterValue',
                    $_.DisplayName
                )
            }
    })]
    [System.Globalization.CultureInfo]$culture = (Get-Culture),

    [Parameter(HelpMessage = 'Specifies the time zone ID for the calendar entries.')]
    [ValidateNotNullOrEmpty()]
    [ValidateSet( 
        'Europe/Berlin',
        'Europe/Paris',
        'Europe/Rome',
        'Europe/Madrid',
        'Europe/Brussels',
        'Europe/Warsaw',
        'UTC'
    )]
    [string]$TimeZoneID = 'Europe/Berlin',

    [ValidateNotNullOrEmpty()]
    [Parameter(HelpMessage = 'Specifies the cookie value for the WebUntis session.')]
    [string]$cookie,

    [ValidateNotNullOrEmpty()]
    [Parameter(HelpMessage = 'Specifies the tenant ID for the WebUntis session.')]
    [string]$tenantId
)

if ($groupByPrio) {
    $dontSplitHigherPrio = $false
}

if ($outAllFormats) {
    $splitByCourse = $true
    $dontSplitHigherPrio = $false
}
if ($appendToPreviousICSat) {
    if (-not (Test-Path $appendToPreviousICSat)) {
        $appendToPreviousICSat = $null
    }
}

# Convert any string inputs to DateTime objects
$dates = $dates | ForEach-Object {
    if ($_ -is [string]) { [datetime]::Parse($_) } else { $_ }
}

if ($culture -isnot [System.Globalization.CultureInfo]) {
    $culture = [System.Globalization.CultureInfo]::GetCultureInfo($culture)
}
Write-Host "::notice::Using $($culture.NativeName) as formatting culture."

function Get-SingleElement {
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Object[]]$collection
    )

    process {
        if ($collection.Length -eq 0) {
            throw [System.InvalidOperationException]::new("No elements match the predicate. Call stack: $((Get-PSCallStack | Out-String).Trim())")
        } elseif ($collection.Length -gt 1) {
            throw [System.InvalidOperationException]::new("More than one element matches the predicate. Call stack: $((Get-PSCallStack | Out-String).Trim())")
        }

        return $collection[0]
    }
}

# Function to calculate the week start date (Monday) for a given date
function Get-WeekStartDate($date) {
    $offset = ($date.DayOfWeek.value__ + 6) % 7
    return $date.Date.AddDays(-$offset)
}

$headers = @{
    'authority'                 = "$baseUrl"
    'accept'                    = 'application/json'
    'accept-encoding'           = 'gzip, deflate, br, zstd'
    'accept-language'           = 'de-DE,de;q=0.9,en-US;q=0.8,en;q=0.7'
    'cache-control'             = 'max-age=0'
    'dnt'                       = '1'
    'pragma'                    = 'no-cache'
    'priority'                  = 'u=0, i'
    'sec-ch-ua'                 = "`"Google Chrome`";v=`"131`", `"Chromium`";v=`"131`", `"Not_A Brand`";v=`"24`""
    'sec-ch-ua-mobile'          = '?0'
    'sec-ch-ua-platform'        = "`"Windows`""
    'sec-fetch-dest'            = 'document'
    'sec-fetch-mode'            = 'navigate'
    'sec-fetch-site'            = 'none'
    'sec-fetch-user'            = '?1'
    'upgrade-insecure-requests' = '1'
}

$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$session.UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'
$session.Cookies.Add((New-Object System.Net.Cookie('schoolname', "`"$cookie`"", '/', "$baseUrl")))
$session.Cookies.Add((New-Object System.Net.Cookie('Tenant-Id', "`"$tenantId`"", '/', "$baseUrl")))
$session.Cookies.Add((New-Object System.Net.Cookie('schoolname', "`"$cookie==`"", '/', "$baseUrl")))
$session.Cookies.Add((New-Object System.Net.Cookie('Tenant-Id', "`"$tenantId`"", '/', "$baseUrl")))
$session.Cookies.Add((New-Object System.Net.Cookie('schoolname', "`"$cookie==`"", '/', "$baseUrl")))
$session.Cookies.Add((New-Object System.Net.Cookie('Tenant-Id', "`"$tenantId`"", '/', "$baseUrl")))
$session.Cookies.Add((New-Object System.Net.Cookie('traceId', '9de4710537aa097594b039ee3a591cfc22a6dd99', '/', "$baseUrl")))
$session.Cookies.Add((New-Object System.Net.Cookie('JSESSIONID', 'B9ED9B2D36BE7D25A7A9EF21E8144D3F', '/', "$baseUrl")))

$periods = [System.Collections.Generic.List[PeriodEntry]]::new()
$courses = [System.Collections.Generic.List[Course]]::new()
$rooms = [System.Collections.Generic.List[Room]]::new()
$legende = [System.Collections.Generic.List[PeriodTableEntry]]::new();


$lastImportTimeStamp = $null
foreach ($date in $dates) {

    $week = Get-WeekStartDate $date
    Write-Host "Getting Data for week of $([System.String]::Format($culture, "{0:d}", $week.Date)) - $([System.String]::Format($culture, "{0:d}", $week.Date.addDays(7)))"

    $url = "https://$baseUrl/WebUntis/api/public/timetable/weekly/data?elementType=$elementType&elementId=$elementId&date=$($date.ToString('yyyy-MM-dd'))&formatId=14"

    $response = Invoke-WebRequest -UseBasicParsing -Uri $url -Method Get -WebSession $session -Headers $headers
    $object = $response | ConvertFrom-Json -ErrorAction Stop

    if ($null -ne $object.data.error) {
        $continue = $false
        switch ($object.data.error.data.messageKey) {
            'ERR_TTVIEW_NOTALLOWED_ONDATE' {
                Write-Warning "::warning::Query for $([System.String]::Format($culture, "{0:ddd}, {0:d}", $date)) not allowed. (Query time frame limited by WebUntis API, maximum defined by admin)"
                $continue = $true # PS5 does not support "continue 2", "continue" counts for the switch statement then
                break;
            }
            default {
                Write-Warning "::warning::$($object.data.error.data.messageKey) for value: $($object.data.error.data.messageArgs[0])"
                $continue = $true
                break;
            }
        }
        if ($continue) { continue }
    }


    try {
        $object.data.result.data.elements | ForEach-Object {
            if ($legende.FindAll({ param($e) $e.id -eq $_.id -and $e.type -eq $_.type }).Count -eq 0) {
                # prevent duplicates
                $legende.Add([PeriodTableEntry]::new($_)) 
            } 
        }
        $legende | Where-Object { $_.type -eq 4 } | ForEach-Object {
            if ($rooms.FindAll({ param($e) $e.id -eq $_.id }).Count -eq 0) {
                # prevent duplicates
                $rooms.Add([Room]::new($_)) 
            } 
        }
        $legende | Where-Object { $_.type -eq 3 } | ForEach-Object {
            if ($courses.FindAll({ param($e) $e.id -eq $_.id }).Count -eq 0) {
                # prevent duplicates
                if ($overrideSummaries) {
                    if ($overrideSummaries.Contains($_.name)) {
                        $_.longName = $overrideSummaries[$_.name]
                    }
                }
                $courses.Add([Course]::new($_))
            } 
        }

        $class = ($legende.Where({ $_.id -eq $elementId }) | Get-SingleElement)
        $class = [PSCustomObject]@{
            name          = $class.name
            longName      = $class.longName
            displayname   = $class.displayname
            alternatename = $class.alternatename
            backColor     = $class.backColor
        }


        $object.data.result.data.elementPeriods.$elementId | ForEach-Object {
            try {
                $element = $_
                $periods.Add([PeriodEntry]::new($_, $rooms, $courses)) 
            } catch {
                #[FormatException] {
                Write-Error ($element | Format-List | Out-String)
                throw
            }
        }

        $lastImportTimeStamp = [System.DateTimeOffset]::FromUnixTimeMilliseconds($object.data.result.lastImportTimeStamp).DateTime
    } catch [FormatException] {
        Write-Error 'Invalid Response regarding datetime format:'
        throw
        exit 1
    }
}
$lastImportTimeStamp = [datetime]::SpecifyKind($lastImportTimeStamp, [DateTimeKind]::Local)
if (-not (Get-Date).IsDaylightSavingTime()) { # website seems to display time +1h from the unix time stamp it serves. (shouldn't it be -1 for DST?)
    $lastImportTimeStamp = $lastImportTimeStamp.AddHours(1);
} else {
    $lastImportTimeStamp = $lastImportTimeStamp.AddHours(-1);
}

$periods.sort({ param($a, $b) $a.startTime.CompareTo($b.startTime) })

if ($periods.Count -eq 0 -or $null -eq $periods) {
    Write-Host "::warning::No Periods in the specified time frame"
    exit 0
}

if ($connectMaxGapMinutes -ne -1 -and $null -ne $connectMaxGapMinutes) {
    $breakEntries = [System.Collections.Generic.List[PeriodEntry]]::new()
    for ($i = 1; $i -lt $periods.Count; $i++) {
        $period = $periods[$i]
        $nextperiod = $periods[$i+1]
        if (($nextperiod.startTime - $period.endTime).TotalMinutes -lt 0) {
            continue # skip overlapping periods
        }
        $i2 = $i - 1
        while ($null -eq $periods[$i2]) {
            $i2--
        }
        $prevperiod = $periods[$i2]

        if (($period.startTime - $prevperiod.endTime).TotalMinutes -le $connectMaxGapMinutes -and `
            $period.course.course.id -eq $prevperiod.course.course.id -and `
            $period.room.room.id -eq $prevperiod.room.room.id -and `
            $period.cellState -eq $prevperiod.cellState -and `
            $period.substText -eq $prevperiod.substText
        ) {
            if (-not $dontCreateBreakEntriesForGaps -and ($period.startTime - $prevperiod.endTime).TotalMinutes -gt 0) {
                $breakJson = [PSCustomObject]@{
                    id         = Get-Random -Minimum 100000 -Maximum $([int]::MaxValue)
                    date       = $period.startTime.Date.ToString('yyyyMMdd')
                    startTime  = $period.endTime.ToString('hhmm')
                    endTime    = $nextperiod.startTime.ToString('hhmm')
                    elements   = @(@{
                        type = 3
                        id = 0
                    })
                    substText  = "$(($period.startTime - $prevperiod.endTime).TotalMinutes)m break in $($period.course.course.name)" # Entry Description
                    lessonCode = 'BREAK'
                    cellstate  = 'ADDITIONAL'
                    priority   = 1
                }
                $breakEntry = [PeriodEntry]::new(
                    $breakJson,
                    $rooms, 
                    [Course]::new([PeriodTableEntry]::new(@{
                        type = 3
                        id = 0
                        longName = "$(($period.startTime - $prevperiod.endTime).TotalMinutes)m break"
                    }))
                )
                Write-Host "Adding break entry: $($breakEntry.ToString())"
                if ([string]::IsNullOrWhiteSpace($breakEntry.course.course.longName)) {
                    Write-Host "::warning::Break entry has no course name set? $($breakEntry.ToString()), $($breakEntry.course.ToString()), $($breakEntry.course.course.ToString())" 
                    # this should not happen but when it happens it seems to produce many break entries like this:
                    # BEGIN:VEVENT
                    # UID:1367233
                    # DTSTART;TZID=Europe/Berlin:20250411T110000
                    # DTEND;TZID=Europe/Berlin:20250411T114500
                    # LOCATION:116 Klz
                    # SUMMARY:Wi, Wirtschafts- und Sozialkunde
                    # DESCRIPTION:
                    # STATUS:CONFIRMED
                    # CATEGORIES:LESSON
                    # PRIORITY:5
                    # TRANSP:OPAQUE
                    # END:VEVENT
                    # BEGIN:VEVENT ---> this entry is present like 20 times with differend UIDs
                    # UID:72828535
                    # DTSTART;TZID=Europe/Berlin:20250516T100000
                    # DTEND;TZID=Europe/Berlin:20250516T100000
                    # LOCATION:
                    # SUMMARY:
                    # DESCRIPTION:0m break in LBT4
                    # STATUS:TENTATIVE
                    # CATEGORIES:BREAK
                    # PRIORITY:9
                    # TRANSP:TRANSPARENT
                    # END:VEVENT
                    # BEGIN:VEVENT
                    # UID:1367242
                    # DTSTART;TZID=Europe/Berlin:20250516T110000
                    # DTEND;TZID=Europe/Berlin:20250516T114500
                    # LOCATION:116 Klz
                    # SUMMARY:Wi, Wirtschafts- und Sozialkunde
                    # DESCRIPTION:
                    # STATUS:CONFIRMED
                    # CATEGORIES:LESSON
                    # PRIORITY:5
                    # TRANSP:OPAQUE
                    # END:VEVENT
                }
                $breakEntries.add($breakEntry)
            }
       
            $prevperiod.endTime = $period.endTime
            $periods[$i] = $null
        }
    }
    $periods.RemoveAll({ param($a) $a -eq $null })
    $periods.AddRange($breakEntries)
    $periods.Sort({ param($a,$b) $a.startTime.CompareTo($b.startTime) -or $a.endTime.CompareTo($b.endTime) })
}

if (-not $dontCreateMultiDayEvents) {

    # Always create a dummy Summary event so a file exists (prevents issues with outlook)
    if ($periods.Count -eq 0) {
        $summaryJson = [PSCustomObject]@{
            id         = 0
            date       = ([datetime]"2020-01-01T00:00:00Z").Date.ToString('yyyyMMdd')
            startTime  = ([datetime]"2020-01-01T00:00:00Z").ToString('hhmm')
            endTime    = ([datetime]"2020-01-01T00:00:00Z").AddMinutes(1)
            course = @{
                course = @{
                    longName = "DUMMY"
                }
            }
            substText  = "Dummy because No events in timeframe $([System.String]::Format($culture, "{0:d}", $week.Date)) - $([System.String]::Format($culture, "{0:d}", $week.Date.addDays(7)))"
            lessonCode = 'SUMMARY'
            cellstate  = 'CANCEL'
        }
        $newSummary = [PeriodEntry]::new($summaryJson, $rooms, $courses)
    }

    # Add WeekStartDate property to each period
    $periods | ForEach-Object {
        $_ | Add-Member -NotePropertyName WeekStartDate -NotePropertyValue (Get-WeekStartDate $_.startTime) -Force
    }

    # Group periods by WeekStartDate
    $periodsGroupedByWeek = $periods.Where({ $_.isCancelled -ne $true }) | Group-Object -Property WeekStartDate

    # Initialize an array to hold the new multi-day elements
    $multiDayEvents = [System.Collections.Generic.List[PeriodEntry]]::new()

    # Process each week group
    foreach ($group in $periodsGroupedByWeek) {
        $sortedPeriods = $group.Group # [System.Collections.Generic.List[PeriodEntry]]::new($($group.Group | Sort-Object startTime))


        $first = $sortedPeriods[0]
        $sortedPeriods.remove($first) | Out-Null

        $dayGroups = [System.Collections.Generic.List[System.Collections.Generic.List[PeriodEntry]]]::new()
        $previousDate = $first.startTime.Date

        foreach ($period in $sortedPeriods) {
            $currentDate = $period.startTime.Date
            $daysDifference = ($currentDate - $previousDate).Days
            if (((-not $dontSplitOnGapDays) -and $daysDifference -gt 1) -or $dayGroups.Count -eq 0) {
                # Gap detected, add new group
                $dayGroups.Add([System.Collections.Generic.List[PeriodEntry]]::new())
            }

            # (No gap), add to current group
            $dayGroups[$dayGroups.Count - 1].Add($period)
            $previousDate = $currentDate
        }

        $i = 0
        foreach ($dayGroup in $dayGroups) {
            $i++
            $firstPeriod = $dayGroup[0]
            $lastPeriod = $dayGroup[$dayGroup.Count - 1]

            $calendar = $culture.Calendar
            $weekOfYear = $calendar.GetWeekOfYear($firstPeriod.startTime, $culture.DateTimeFormat.CalendarWeekRule, $culture.DateTimeFormat.FirstDayOfWeek)

            if ($dayGroups.Length -gt 1) {
                $weekOfYear = "$weekOfYear ($i/${dayGroups.Length})"
            }

            do {
                $id = [System.Math]::Abs([System.BitConverter]::ToInt32([System.Guid]::NewGuid().ToByteArray(), 0))
            } while ($periods.Where({ $_.id -eq $id }).Count -ne 0 -or $multiDayEvents.Where({ $_.id -eq $id }).Count -ne 0)

            # Create a new JSON object with necessary properties
            $summaryJson = [PSCustomObject]@{
                id         = $i
                date       = $firstPeriod.startTime.Date.ToString('yyyyMMdd')
                startTime  = $firstPeriod.startTime.ToString('hhmm')
                endTime    = $lastPeriod.endTime
                elements   = @(@{
                    type = 3
                    id = 0
                })
                substText  = "Calendar Week $weekOfYear; For setting longer notifications after some weeks of absence"
                lessonCode = 'SUMMARY'
                cellstate  = 'ADDITIONAL'
            }
            $newSummary = [PeriodEntry]::new(
                $summaryJson,
                $rooms, 
                [Course]::new([PeriodTableEntry]::new(@{
                    type = 3
                    id = 0
                    longName = "Last Gen: $([System.String]::Format($culture, "{0:ddd}, {0:d} {0:t}", (Get-Date))), Source updated: $([System.String]::Format($culture, "{0:ddd}, {0:d} {0:t}", $lastImportTimeStamp))"
                }))
            )
            $multiDayEvents.Add($newSummary)
        }
    }

    $finalList = [System.Collections.Generic.List[PeriodEntry]]::new()
    $finalList.AddRange($multiDayEvents)
    $finalList.AddRange($periods)
    $periods = $finalList
}

$existingPeriods = [System.Collections.Generic.List[PeriodEntry]]::new()

if ($appendToPreviousICSat) {
    Write-Information "Appending to previous ICS file $appendToPreviousICSat"
    $content = Get-Content $appendToPreviousICSat -Raw
    $veventPattern = '(?s)BEGIN:VEVENT.*?END:VEVENT'
    $existingEntries = [regex]::Matches($content, $veventPattern) | ForEach-Object { $_.Value }
    
    foreach ($entry in $existingEntries) {
        $previousIcsEvent = [IcsEvent]::new($entry)
        if ($previousIcsEvent.Category -ne 'SUMMARY') {
            $previousPeriod = [PeriodEntry]::new($previousIcsEvent, $rooms, $courses)
            if ($periods.where({ $_.ID -eq $previousPeriod.ID }).Count -lt 1) {
                $existingPeriods.Add($previousPeriod)
            } else {
                Write-Verbose "Skipping existing entry $($previousPeriod.ID) ($($previousPeriod.StartTime) - $($previousPeriod.EndTime))"
            }
        } else { Write-Verbose "Skipping SUMMARY entry $($previousIcsEvent.StartTime) - $($previousIcsEvent.EndTime)" }
    }
    if ($periods.count -ne 0 -and $null -ne $periods) {
        $existingPeriods.AddRange($periods)
    }
    $periods = $existingPeriods
}

$highPrioPeriods = $periods.Where({ $_.Priority -gt 5 })
if (-not $dontRemoveHighPrioFromNormal) {
    $periods.RemoveAll({ param($a) $highPrioPeriods.Contains($a) })
}

$tmpPeriods = $periods

if ($splitByCourse -and -not $splitByOverrides) {
    $periods = $periods | Group-Object -Property { if (-not [string]::IsNullOrEmpty($_.course.course.name)) {
            $_.course.course.name 
        } else { 
            $_.lessonCode
        } 
    }
    if ($outAllFormats) {
        $periods += ($tmpPeriods | Group-Object -Property { 'All' })
    }
} elseif ($splitByOverrides) {
    $periods = $periods | Group-Object -Property { if (-not [string]::IsNullOrEmpty($_.course.course.name)) { 
            Write-Verbose "Checking for override: $($_.course.course.name) $($overrideSummaries.Keys -contains $_.course.course.name)"
            if ($overrideSummaries.Keys -contains $_.course.course.name) {
                ($overrideSummaries[$_.course.course.name] -split ',')[0]
            } else {
                'Misc'
            }
        } else {
            $_.lessonCode
        } }
    if ($outAllFormats) {
        $periods += ($tmpPeriods | Group-Object -Property { 'All' })
    }
} else {
    $periods = $periods | Group-Object -Property { 'All' }
}

if (-not $dontSplitHigherPrio -or $outAllFormats) {

    $prioJson = [PSCustomObject]@{
        id         = 0
        date       = ([datetime]"2020-01-01T00:00:00Z").Date.ToString('yyyyMMdd')
        startTime  = ([datetime]"2020-01-01T00:00:00Z").ToString('hhmm')
        endTime    = ([datetime]"2020-01-01T00:00:00Z").AddMinutes(1)
        course = @{
            course = @{
                longName = "DUMMY"
            }
        }
        substText  = "Dummy because No events in timeframe $([System.String]::Format($culture, "{0:d}", $week.Date)) - $([System.String]::Format($culture, "{0:d}", $week.Date.addDays(7)))"
        lessonCode = 'PRIO'
        cellstate  = 'CANCEL'
        priority   = 5
    }
    $newPrio = [PeriodEntry]::new($prioJson, $rooms, $courses)

    $highPrioPeriods.Insert(0, $newPrio)

    $highPrioGroups = $highPrioPeriods | Group-Object -Property { "PRIO" }
    $highPrioGroups = @($highPrioGroups) # so I can add the new groups

    $highPrioPeriods.RemoveAt(0) # remove the dummy entry

    if ($groupByPrio -or $outAllFormats) {
        # clone into PRIO group
        $highPrioGroups += ($highPrioPeriods | Group-Object -Property {
            $prioStr = "PRIO$($_.Priority)"
            
            Write-Verbose "Checking for override: $prioStr $($overrideSummaries.Keys -contains $prioStr)"
            if ($overrideSummaries.Keys -contains $prioStr) {
                ($overrideSummaries[$prioStr] -split ',')[0]
            } else {
                $prioStr
            }
        })
    }

    foreach ($prioGroup in $highPrioGroups) {
        $periods += $prioGroup
    }
}

foreach ($group in $periods) {
    
    $calendarEntries = [System.Collections.Generic.List[IcsEvent]]::new()

    # Iterate over each period and create calendar entries
    foreach ($period in $group.Group) {
        $calendarEntries.Add([IcsEvent]::new($period, $TimeZoneID))
    }

    # Get all properties except StartTime and EndTime
    $properties = $calendarEntries | Get-Member -MemberType Properties | Where-Object {
        $_.Name -ne 'StartTime' -and $_.Name -ne 'EndTime' -and $_.Name -ne 'Description' -and $_.Name -ne 'UID' -and $_.Name -ne 'preExist'
    } | Select-Object -ExpandProperty Name

    if ($splitByCourse) {
        if ($group -ne $periods[0]) {
            Write-Host "`n`n`n`n`n`n`n" # make cmdline output more readable
        }
        Write-Host "ICS content for $($group.Name):`n============================================================="
    }
    # Use Select-Object to reorder properties and reformat for better cmdline output
    $calendarEntries | Select-Object (@(
            @{ Name = 'pre'; Expression = { if ($_.preExist) { '[X]' } else { '[ ]' } } },
            @{ Name = 'StartTimeF'; Expression = { 
                $datetime = $_.StartTime
                if ($datetime -match ';.*:(\d{8}T\d{6})') { # workaraound because IcsEvent doesn't know if it's Summary (see .ToIcsEntry())
                    $datetime = $matches[1]
                }
                [DateTime]::ParseExact($datetime, 'yyyyMMddTHHmmss', $null).ToString("g", $culture)
            } },
            @{ Name = 'EndTimeF'; Expression = { 
                $datetime = $_.EndTime
                if ($datetime -match ';.*:(\d{8}T\d{6})') {
                    $datetime = $matches[1]
                }
                [DateTime]::ParseExact($datetime, 'yyyyMMddTHHmmss', $null).ToString("g", $culture)
            } }
        ) + $properties + @{ 
            Name       = 'DescriptionF'; 
            Expression = {
                $_.Description -replace '`n', ';; '
            } 
        }
    ) | Format-Table -Wrap -AutoSize | Out-String -Width 4096


    $IcsEntries = [System.Collections.Generic.List[string]]::new()
    foreach ($icsEvent in $calendarEntries) {
        $IcsEntries += $icsEvent.ToIcsEntry()
    }
   

    # Create the .ics file content
    $icsContent = @"
BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Chaos_02//WebUntisToIcs//EN
CALSCALE:GREGORIAN
METHOD:PUBLISH
X-WR-TIMEZONE:$TimeZoneID
X-MS-OLK-WKHRSTART:073000
X-MS-OLK-WKHREND:130000
REFRESH-INTERVAL;VALUE=DURATION:PT3H
X-PUBLISHED-TTL:PT6H
X-WR-CALNAME:$(if (-not $splitByCourse) {$class.displayname} else {$class.displayname + " - $($group.Name)"})
BEGIN:VTIMEZONE
TZID:$TimeZoneID
LAST-MODIFIED:$((Get-Date).ToUniversalTime().ToString('yyyyMMddTHHmmssZ'))
X-LIC-LOCATION:$TimeZoneID
BEGIN:DAYLIGHT
DTSTART:19810329T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU
TZOFFSETFROM:+0100
TZOFFSETTO:+0200
TZNAME:$($culture.DateTimeFormat.DaylightName)
END:DAYLIGHT
BEGIN:STANDARD
DTSTART:19961027T030000
RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU
TZOFFSETFROM:+0200
TZOFFSETTO:+0100
TZNAME:$($culture.DateTimeFormat.StandardName)
END:STANDARD
END:VTIMEZONE
$(($IcsEntries -join "`n"))
END:VCALENDAR
"@

    try {
        if ($OutputFilePath) {
            # Write the .ics content to a file
            if ($splitByCourse) {
                if ($group.Name -ne 'All') {
                    $OutputPath = $OutputFilePath.FullName.Insert($OutputFilePath.FullName.LastIndexOf('.'), "_$($group.Name -replace '[^a-zA-Z0-9]', '_')")
                } else {
                    $OutputPath = $OutputFilePath
                }
            } else {
                $OutputPath = $OutputFilePath
            }
            Set-Content -Path $OutputPath -Value $icsContent
            Write-Output "ICS file created at $((Get-Item -Path $OutputPath).FullName)"
        } else {
            # for writing the .ics content to a variable
            Write-Output $icsContent
            if (-not $splitByCourse) {
                return $icsContent
            } else {
                # TODO: return after all period groups
            }
        }
    } catch {
        Write-Error "An error occurred while creating the ICS file: $_"
        throw
    }

}

####### Class definitions #######

class IcsEvent {
    [string]$UID
    [string]$StartTime
    [string]$EndTime
    [string]$Location
    [string]$Summary
    [string]$Description
    [string]$Status
    [string]$Category
    [int]$Priority
    [bool]$Transparent
    [bool]$preExist = $false

    IcsEvent([PeriodEntry]$period, [string]$TimeZoneID) {
        $this.preExist = $period.preExist
        $this.UID = $period.id
        $adjustedStartTime = if ((Get-Date).IsDaylightSavingTime()) { $period.startTime.AddHours(-1) } else { $period.startTime }
        $adjustedEndTime = if ((Get-Date).IsDaylightSavingTime()) { $period.endTime.AddHours(-1) } else { $period.endTime }
        if ($period.lessonCode -ne 'SUMMARY') {
            $this.startTime = ";TZID=$($TimeZoneID):" + $adjustedStartTime.ToString('yyyyMMddTHHmmss')
            $this.endTime = ";TZID=$($TimeZoneID):" + $adjustedEndTime.ToString('yyyyMMddTHHmmss')
        } else {
            $this.startTime = ';VALUE=DATE:' + $adjustedStartTime.ToString('yyyyMMdd')
            $this.endTime = ';VALUE=DATE:' + $adjustedEndTime.AddDays(1).ToString('yyyyMMdd')
        }
        $this.location = $period.room.room.longName
        $this.summary = $period.course.course.longName
        $this.description = $period.substText
        if ($null -ne $period.rescheduleInfo) {
            $this.description += "`nReschedule:`n" + $period.rescheduleInfo.ToString()
        }
        $this.status = switch ($period.cellState) {
            'STANDARD' { 'CONFIRMED' }
            'ADDITIONAL' { 'TENTATIVE' }
            'CANCEL' { 'CANCELLED' }
            'SHIFT' { 'CONFIRMED' }
            'ROOMSUBSTITUTION' { 'CONFIRMED' }
            'SUBSTITUTION' { 'CONFIRMED' }
            'BREAK' { 'TENTATIVE' }
            'SUMMARY' { 'CONFIRMED' }
            'CONFIRMED' { $_ }
            'TENTATIVE' { $_ }
            'CANCELLED' { $_ }
            default { 'CONFIRMED' }
        }
        $this.category = switch ($period.lessonCode) {
            'UNTIS_ADDITIONAL' { 'Additional' }
            default { $_ }
        }
        $this.Priority = 10 - $period.priority
        $this.Transparent = ($this.Status -ne 'CONFIRMED')
    }

    IcsEvent([string]$icsText) {
        if ($icsText -notmatch 'BEGIN:VEVENT' -or ($icsText -match 'BEGIN:VEVENT' -and ($icsText -match 'BEGIN:VEVENT.*BEGIN:VEVENT'))) {
            throw 'Invalid Syntax in ICS entry. Only one VEVENT element is allowed.'
        }
        if ($icsText -notmatch 'END:VEVENT' -or ($icsText -match 'END:VEVENT' -and ($icsText -match 'END:VEVENT.*END:VEVENT'))) {
            throw 'Invalid Syntax in ICS entry. Only one VEVENT element is allowed.'
        }
        $this.preExist = $true
        if ($icsText -match 'UID:(.*)') { $this.UID = $matches[1].Trim() } else { throw 'UID not found in ICS entry.' }
        if ($icsText -match 'DTSTART;(?:TZID=.*|VALUE=DATE):(.*)') { $this.StartTime = $matches[1].Trim() } else { throw 'StartTime not found in ICS entry.' }
        if ($icsText -match 'DTEND;(?:TZID=.*|VALUE=DATE):(.*)') { $this.EndTime = $matches[1].Trim() } else { throw 'EndTime not found in ICS entry.' }
        if ($icsText -match 'LOCATION:(.*)') { $this.Location = $matches[1].Trim() } else { throw 'Location not found in ICS entry.' }
        if ($icsText -match 'SUMMARY:(.*)') { $this.Summary = $matches[1].Trim() } else { throw 'Summary not found in ICS entry.' }
        if ($icsText -match 'DESCRIPTION:(.*)') { $this.Description = $matches[1].Trim() } else { throw 'Description not found in ICS entry.' }
        if ($icsText -match 'STATUS:(.*)') { $this.Status = $matches[1].Trim() } else { throw 'Status not found in ICS entry.' }
        if ($icsText -match 'CATEGORIES:(.*)') { $this.Category = $matches[1].Trim() } else { throw 'Category not found in ICS entry.' }
        if ($icsText -match 'PRIORITY:(.*)') {$this.Priority = $matches[1].Trim() } else { Write-Warning 'Priority not found in ICS entry'; $this.Priority = 5 }
        if ($icsText -match 'TRANSP:(.*)') {$this.Transparent = ($matches[1].Trim() -eq 'TRANSPARENT') } else { Write-Warning 'Transparency not found in ICS entry'; $this.Transparent = $false }
    }

    [string] ToIcsEntry() {
        return @"
BEGIN:VEVENT
UID:$($this.UID)
DTSTART$($this.StartTime)
DTEND$($this.EndTime)
LOCATION:$($this.Location)
SUMMARY:$($this.Summary)
DESCRIPTION:$($this.Description)
STATUS:$($this.Status)
CATEGORIES:$($this.Category)
PRIORITY:$($this.Priority)
TRANSP:$(switch ($this.Transparent) { $true {'TRANSPARENT'} $false {'OPAQUE'}})
END:VEVENT
"@
    }
    
}

class rescheduleInfo {
    [datetime]$startTime
    [datetime]$endTime
    [bool]$isSource

    [datetime] date() {
        return $this.startTime.Date
    }

    rescheduleInfo([PSCustomObject]$jsonObject) {
        $this.startTime = [datetime]::ParseExact($jsonObject.date.ToString(), 'yyyyMMdd', $null).Add([timespan]::ParseExact($jsonObject.startTime.ToString().PadLeft(4, '0'), 'hhmm', $null))
        $this.endTime = $this.date().Add([timespan]::ParseExact($jsonObject.endTime.ToString().PadLeft(4, '0'), 'hhmm', $null))
        $this.isSource = $jsonObject.isSource
    }

    [string] ToString() {
        return "Start Time: $($this.startTime), End Time: $($this.endTime), Is Source: $($this.isSource)"
    }
}

class PeriodEntry {
    [int]$id
    [int]$lessonId
    [int]$lessonNumber
    [string]$lessonCode
    [string]$lessonText
    [string]$periodText
    [bool]$hasPeriodText
    [string]$periodInfo
    [array]$periodAttachments
    [string]$substText
    [datetime]$startTime
    [datetime]$endTime
    [RoomEntry]$room
    [CourseEntry]$course
    [string]$studentGroup
    [int]$code
    [string]$cellState
    [int]$priority
    [bool]$isStandard
    [bool]$isCancelled
    [bool]$isEvent
    [rescheduleInfo]$rescheduleInfo
    [int]$roomCapacity
    [int]$studentCount
    [bool]$preExist = $false

    [datetime] date() {
        return $this.startTime.Date
    }

    PeriodEntry([PSCustomObject]$jsonObject, [System.Collections.Generic.List[Room]]$rooms, [System.Collections.Generic.List[Course]]$courses) {
        $this.id = $jsonObject.id
        $this.lessonId = $jsonObject.lessonId
        $this.lessonNumber = $jsonObject.lessonNumber
        $this.lessonCode = $jsonObject.lessonCode
        $this.lessonText = $jsonObject.lessonText
        $this.periodText = $jsonObject.periodText
        $this.hasPeriodText = $jsonObject.hasPeriodText
        $this.periodInfo = $jsonObject.periodInfo
        $this.periodAttachments = $jsonObject.periodAttachments
        $this.substText = $jsonObject.substText
        $this.startTime = [datetime]::ParseExact($jsonObject.date.ToString(), 'yyyyMMdd', $null).Add([timespan]::ParseExact($jsonObject.startTime.ToString().PadLeft(4, '0'), 'hhmm', $null))
        if ($jsonObject.endTime -is [DateTime]) {
            $this.endTime = $jsonObject.endTime
        } else {
            $this.endTime = $this.date().Add([timespan]::ParseExact($jsonObject.endTime.ToString().PadLeft(4, '0'), 'hhmm', $null))
        }
        $this.startTime = [datetime]::SpecifyKind($this.startTime, [System.DateTimeKind]::Local)
        $this.endTime = [datetime]::SpecifyKind($this.endTime, [System.DateTimeKind]::Local)
        $this.room = [RoomEntry]::new(($jsonObject.elements | Where-Object { $_.type -eq 4 } | Get-SingleElement), $rooms)
        $this.course = [CourseEntry]::new(($jsonObject.elements | Where-Object { $_.type -eq 3 } | Get-SingleElement), $courses)
        $this.studentGroup = $jsonObject.studentGroup
        $this.code = $jsonObject.code
        $this.cellState = $jsonObject.cellState
        $this.priority = switch ($jsonObject.priority) {$null {5} default {$_}}
        $this.isCancelled = $jsonObject.is.cancelled
        $this.isStandard = $jsonObject.is.standard
        $this.isEvent = $jsonObject.is.event
        if ($null -ne $jsonObject.rescheduleInfo) {
            $this.rescheduleInfo = [rescheduleInfo]::new($jsonObject.rescheduleInfo)
        } else {
            $this.rescheduleInfo = $null
        }
        $this.roomCapacity = $jsonObject.roomCapacity
        $this.studentCount = $jsonObject.studentCount
    }

    PeriodEntry([IcsEvent]$icsEvent, [System.Collections.Generic.List[Room]]$rooms, [System.Collections.Generic.List[Course]]$courses) {
        $this.preExist = $true
        $this.id = $icsEvent.UID
        $this.course = [CourseEntry]::new(($courses.Where({ $_.longName -eq $icsEvent.Summary }) | Get-SingleElement))
        $this.room = [RoomEntry]::new(($rooms.Where({ $_.longName -eq $icsEvent.Location }) | Get-SingleElement))
        $this.lessonCode = $icsEvent.Category
        $this.substText = $icsEvent.Description
        $this.cellState = $icsEvent.Status
        $this.priority = 10 - $icsEvent.Priority

        try {
            $this.startTime = [datetime]::ParseExact($icsEvent.StartTime, 'yyyyMMddTHHmmss', $null)
            $this.endTime = [datetime]::ParseExact($icsEvent.EndTime, 'yyyyMMddTHHmmss', $null)
        } catch {
            $this.startTime = [datetime]::ParseExact($icsEvent.StartTime, 'yyyyMMdd', $null)
            $this.endTime = [datetime]::ParseExact($icsEvent.EndTime, 'yyyyMMdd', $null)
        }
        
        #[datetime]::ParseExact($icsEvent.StartTime, "yyyyMMdd", $null)
        #[datetime]::ParseExact($icsEvent.StartTime, @("yyyyMMddTHHmmss", "yyyyMMdd

    }

    [string] ToString() {
        return "Start Time: $($this.startTime), End Time: $($this.endTime), Cell State: $($this.cellState), ID: $($this.id), Lesson ID: $($this.lessonId), Lesson Number: $($this.lessonNumber), Lesson Code: $($this.lessonCode), Lesson Text: $($this.lessonText), Period Text: $($this.periodText), Has Period Text: $($this.hasPeriodText), Period Info: $($this.periodInfo), Period Attachments: $($this.periodAttachments), Subst Text: $($this.substText), Elements: $($this.elements), Student Group: $($this.studentGroup), Code: $($this.code), Priority: $($this.priority), Is Standard: $($this.isStandard), Is Event: $($this.isEvent), Room Capacity: $($this.roomCapacity), Student Count: $($this.studentCount)"
    }
}

class RoomEntry {
    [Room]$room
    [int]$orgId
    [bool]$missing
    [string]$state

    RoomEntry([PSCustomObject]$jsonObject, [System.Collections.Generic.List[Room]]$rooms) {
        Write-Debug "RoomEntry: $($jsonObject)"
        $this.room = $rooms | Where-Object { $_.id -eq $jsonObject.id } | Get-SingleElement
        $this.orgId = $jsonObject.orgId
        $this.missing = $jsonObject.missing
        $this.state = $jsonObject.state
    }

    RoomEntry([Room]$room) {
        $this.room = $room
        $this.orgId = $null
        $this.missing = $null
        $this.state = $null
    }
}

class Room {
    [int]$id
    [string]$name
    [string]$longName
    [string]$displayname
    [string]$alternatename
    [int]$roomCapacity

    Room([PeriodTableEntry]$legende) {
        if ($legende.type -ne 4) {
            throw [System.ArgumentException]::new('The provided object is not a room.')
        }
        $this.id = $legende.id
        $this.name = $legende.name
        $this.longName = $legende.longName
        $this.displayname = $legende.displayname
        $this.alternatename = $legende.alternatename
        $this.roomCapacity = $legende.roomCapacity
    }
}


class CourseEntry {
    [Course]$course
    [int]$orgId
    [bool]$missing
    [string]$state

    CourseEntry([PSCustomObject]$jsonObject, [System.Collections.Generic.List[Course]]$courses) {
        $this.course = $courses | Where-Object { $_.id -eq $jsonObject.id } | Get-SingleElement
        $this.orgId = $jsonObject.orgId
        $this.missing = $jsonObject.missing
        $this.state = $jsonObject.state
    }

    CourseEntry([Course]$course) {
        $this.course = $course
        $this.orgId = $null
        $this.missing = $null
        $this.state = $null
    }

    [string] ToString() {
        return "Course: $($this.course.name), ID: $($this.course.id), Org ID: $($this.orgId), Missing: $($this.missing), State: $($this.state)"
    }
}

class Course {
    [int]$id
    [string]$name
    [string]$longName
    [string]$displayname
    [string]$alternatename
    [int]$courseCapacity

    Course([PeriodTableEntry]$legende) {
        if ($legende.type -ne 3) {
            throw [System.ArgumentException]::new('The provided object is not a course.')
        }
        $this.id = $legende.id
        $this.name = $legende.name
        $this.longName = $legende.longName
        $this.displayname = $legende.displayname
        $this.alternatename = $legende.alternatename
        $this.courseCapacity = $legende.courseCapacity
    }

    [string] ToString() {
        return "Course: $($this.name), ID: $($this.id), Long Name: $($this.longName), Display Name: $($this.displayname), Alternate Name: $($this.alternatename), Course Capacity: $($this.courseCapacity)"
    }
}


class PeriodTableEntry {
    [int]$type
    [int]$id
    [string]$name
    [string]$longName
    [string]$displayname
    [string]$alternatename
    [string]$backColor
    [bool]$canViewTimetable
    [int]$roomCapacity

    PeriodTableEntry([PSCustomObject]$jsonObject) {
        $this.type = $jsonObject.type
        $this.id = $jsonObject.id
        $this.name = $jsonObject.name
        $this.longName = $jsonObject.longName
        $this.displayname = $jsonObject.displayname
        $this.alternatename = $jsonObject.alternatename
        $this.backColor = $jsonObject.backColor
        $this.canViewTimetable = $jsonObject.canViewTimetable
        $this.roomCapacity = $jsonObject.roomCapacity
    }

    [string] ToString() {
        return "Type: $($this.type), ID: $($this.id), Name: $($this.name), Long Name: $($this.longName), Display Name: $($this.displayname), Alternate Name: $($this.alternatename), Back Color: $($this.backColor), Can View Timetable: $($this.canViewTimetable), Room Capacity: $($this.roomCapacity)"
    }
}
