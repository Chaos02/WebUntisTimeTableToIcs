<#
.SYNOPSIS
    Gets and converts a WebUntis timetable to an ICS calendar file.

.DESCRIPTION
    This script retrieves timetable data from the WebUntis API and converts it into a subscribable ICS calendar file format.
    It allows specifying a date range for the timetable data.

.PARAMETER baseUrl
    The base URL of the WebUntis API.

.PARAMETER elementType
    The type of element (default is 1; typically returns a timetable).

.PARAMETER elementId
    The class timetable ID.

.PARAMETER dates
    An array of dates (either as strings or DateTime objects) for which to retrieve timetable data.
    The default is the current week and the next three weeks.
    Maximum range is defined by WebUntis admin afaik.

.PARAMETER OutputFilePath
    The file path where the ICS file will be saved. Default is "calendar.ics".

.PARAMETER dontCreateMultiDayEvents
    If set, generating "summary" multi-day events will be skipped.

.PARAMETER dontSplitOnGapDays
    If set, multi-day events will not be split even if gap days exist between periods.

.PARAMETER overrideSummaries
    A hashtable to override course summaries. The key is the original (short) course name, and the value is the new course name.

.PARAMETER appendToPreviousICSat
    The path to an existing ICS file to which new timetable data should be appended.
    (The file must have a .ics extension and be a valid ICS file.)

.PARAMETER connectLessons
    If set, consecutive lessons will be connected/merged if they occur within a specified maximum gap.

.PARAMETER connectMaxGapMinutes
    Maximum gap in minutes between consecutive lessons to connect.
    Set to -1 to disable connection. Default is 15 minutes.

.PARAMETER splitByCourse
    If set, timetable data will be split into separate ICS files for each course.

.PARAMETER splitByOverrides
    If set, timetable data will be split into separate ICS files for each course defined in overrideSummaries and the remaining misc. classes.

.PARAMETER outAllFormats
    If set, outputs the timetable data in all available formats (implies -splitByCourse).

.PARAMETER culture
    Culture info used for DST/TZ adjustments and formatting date/time values.
    Default is the system culture as returned by Get-Culture.

.PARAMETER cookie
    The cookie value for the WebUntis session.

.PARAMETER tenantId
    The tenant ID for the WebUntis session.

.EXAMPLE
    .\timeTableToIcs.ps1 -baseUrl "your.webuntis.url" -elementType 1 -elementId 12345 -dates "2023-01-01", "2023-01-08" -OutputFilePath "mycalendar.ics" -cookie "your_cookie" -tenantId "your_tenant_id"

.NOTES
    Author: Chaos_02
    Date: 2024-09-28
    Version: 1.5
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
    [System.Object[]]$dates = @( (@(-7, 0, 7, 14, 21, 28) | ForEach-Object { 
        $date = (Get-Date).Date
        $date.AddHours(12)
        $date.AddDays($_) }) ),
    [switch]$dontCreateMultiDayEvents,
    [ValidateScript({ if (($_ -and -not $dontCreateMultiDayEvents) -eq $false) {throw "Can't use together with -dontCreateMultiDayEvents"} else {$true} })]
    [switch]$dontSplitOnGapDays,
    [Parameter(
        Position = 0,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true,
        HelpMessage = 'Output file path.'
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
        HelpMessage = 'Path to an existing ICS file to which the new timetable data should be appended.'
    )]
    [ValidateNotNull()]
    [ValidateScript({
        # Ensure the file exists.
        if (-not $_.Exists) {
            throw "The file '$($_.FullName)' does not exist."
        }
        # Check that the file has a .ics extension (case-insensitive).
        if ($_.Extension.ToLower() -ne ".ics") {
            throw "The file '$($_.FullName)' does not have a .ics extension."
        }
        # Read the file content and verify it looks like a valid ICS file.
        $content = Get-Content -Path $_.FullName -Raw
        if ($content -notmatch '^BEGIN:VCALENDAR' -or $content -notmatch 'END:VCALENDAR\s*$') {
            throw "The file '$($_.FullName)' does not appear to be a valid ICS file."
        }
        $true
    })]
    [System.IO.FileInfo]$appendToPreviousICSat,
    [Parameter(HelpMessage = 'Maximum gap in minutes to connect lessons set to -1 to disable connection')]
    [ValidateScript({ if ($_ -ge -1) { $true } else { throw 'The value must be a non-negative integer.' } })]
    [int]$connectMaxGapMinutes = 15,
    [Parameter(
        ParameterSetName = 'OutputControl',
        HelpMessage = 'Split the timetable data into separate ICS files for each course.'
        #Mandatory = {$splitByOverrides}
    )]
    [switch]$splitByCourse,
    [Parameter(
        ParameterSetName = 'OutputControl',
        HelpMessage = 'Split only by courses defined in overrideSummaries and misc. classes.'
    )]
    [ValidateScript({if ($_ -and -not $overrideSummaries) {throw 'The parameter -splitByOverrides requires the parameter overrideSummaries to be set.'} else {$true}})]
    [switch]$splitByOverrides,
    [Parameter(ParameterSetName = 'OutputControl')]
    [switch]$outAllFormats,
    [Parameter(HelpMessage = 'Controls DST/TZ info (and how DateTimes are formatted)')]
    [ArgumentCompleter({
        param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
        # Retrieve all available cultures (you can choose to restrict to a subset if desired)
        $allCultures = [System.Globalization.CultureInfo]::GetCultures([System.Globalization.CultureTypes]::AllCultures)
        # Filter cultures by matching the user’s typed text (case-insensitive)
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
    [ValidateNotNullOrEmpty()]
    [string]$cookie,
    [ValidateNotNullOrEmpty()]
    [string]$tenantId
)

if ($outAllFormats) {
    $splitByCourse = $true
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

    Write-Verbose "Getting Data for week of $([System.String]::Format($culture, "{0:ddd}, {0:d}", $date)))"

    $url = "https://$baseUrl/WebUntis/api/public/timetable/weekly/data?elementType=$elementType&elementId=$elementId&date=$($date.ToString('yyyy-MM-dd'))&formatId=14"

    $response = Invoke-WebRequest -UseBasicParsing -Uri $url -Method Get -WebSession $session -Headers $headers
    $object = $response | ConvertFrom-Json -ErrorAction Stop

    if ($null -ne $object.data.error) {
        switch ($object.data.error.data.messageKey) {
            'ERR_TTVIEW_NOTALLOWED_ONDATE' {
                Write-Warning "::warning::Query for $([System.String]::Format($culture, "{0:ddd}, {0:d}", $date)) not allowed. (Query time frame limited by WebUntis API, maximum defined by admin)"
                continue
            }
            default {
                Write-Warning "::warning::$($object.data.error.data.messageKey) for value: $($object.data.error.data.messageArgs[0])"
                continue
            }
        }
    }

    $class = [PeriodTableEntry]

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

$periods = $periods | Sort-Object -Property startTime

if ($periods.Length -eq 0 -or $null -eq $periods) {
    Write-Host "::warning::No Periods in the specified time frame"
    exit 0
}

if ($connectMaxGapMinutes -ne -1 -and $null -ne $connectMaxGapMinutes) {
    for ($i = 0; $i -lt $periods.Count; $i++) {
        if ($null -eq $periods[$i]) {
            continue
        }
        $period = $periods[$i]
        $lperiod = $periods[$i+1]
        if (($lperiod.startTime - $period.endTime).TotalMinutes -lt 0) {
            continue # skip overlapping periods
        }
        if (($lperiod.startTime - $period.endTime).TotalMinutes -le $connectMaxGapMinutes -and `
            $lperiod.course.course.id -eq $period.course.course.id -and `
            $lperiod.room.room.id -eq $period.room.room.id -and `
            $lperiod.cellState -eq $period.cellState -and `
            $lperiod.substText -eq $period.substText
        ) {
            $periods[$i].endTime = $lperiod.EndTime
            $periods[$i+1] = $null
            $i++ # skip null entry
            continue
        }
        if ($i -eq 0) {
            continue
        }
        $i2 = $i - 1
        while ($null -eq $periods[$i2]) {
            $i2--
        }
        if (($periods[$i2].endTime - $period.startTime).TotalMinutes -le $connectMaxGapMinutes -and `
            $period.course.course.id -eq $periods[$i2].course.course.id -and `
            $period.room.room.id -eq $periods[$i2].room.room.id -and `
            $period.cellState -eq $periods[$i2].cellState -and `
            $period.substText -eq $periods[$i2].substText
        ) {
            $periods[$i2].endTime = $periods[$i].endTime
            $periods[$i] = $null
        }
    }
    $periods = $periods | Where-Object { $_ -ne $null }
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
            substText  = "Dummy because No events in timeframe $($dates[0]) - $($dates[$dates.Count - 1])" # next is not guaranteed
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

    $periods = ($multiDayEvents + $periods)
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
        $periods = ($existingPeriods + $periods)
    } else {
        $periods = $existingPeriods
    }
}

if ($splitByCourse -and -not $splitByOverrides) {
    $tmpPeriods = $periods
    $periods = $periods | Group-Object -Property { if (-not [string]::IsNullOrEmpty($_.course.course.name)) {
            $_.course.course.name 
        } else { 
            $_.lessonCode
        } }
    if ($outAllFormats) {
        $periods += ($tmpPeriods | Group-Object -Property { 'All' })
    }
} elseif ($splitByOverrides) {
    $tmpPeriods = $periods
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

foreach ($group in $periods) {
    
    $calendarEntries = [System.Collections.Generic.List[IcsEvent]]::new()

    # Iterate over each period and create calendar entries
    foreach ($period in $group.Group) {
        $calendarEntries.Add([IcsEvent]::new($period))
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
X-WR-TIMEZONE:$($culture.DateTimeFormat.TimeZoneInfo.Id)
X-MS-OLK-WKHRSTART:073000
X-MS-OLK-WKHREND:130000
REFRESH-INTERVAL;VALUE=DURATION:PT3H
X-PUBLISHED-TTL:PT6H
X-WR-CALNAME:$(if (-not $splitByCourse) {$class.displayname} else {$class.displayname + " - $($group.Name)"})
BEGIN:VTIMEZONE
TZID:$($culture.DateTimeFormat.TimeZoneInfo.Id)
LAST-MODIFIED:$((Get-Date).ToUniversalTime().ToString('yyyyMMddTHHmmssZ'))
X-LIC-LOCATION:$($culture.DateTimeFormat.TimeZoneInfo.Id)
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

    IcsEvent([PeriodEntry]$period) {
        $this.preExist = $period.preExist
        $this.UID = $period.id
        $adjustedStartTime = if ((Get-Date).IsDaylightSavingTime()) { $period.startTime.AddHours(-1) } else { $period.startTime }
        $adjustedEndTime = if ((Get-Date).IsDaylightSavingTime()) { $period.endTime.AddHours(-1) } else { $period.endTime }
        if ($period.lessonCode -ne 'SUMMARY') {
            $this.startTime = ';TZID=Europe/Berlin:' + $adjustedStartTime.ToString('yyyyMMddTHHmmss')
            $this.endTime = ';TZID=Europe/Berlin:' + $adjustedEndTime.ToString('yyyyMMddTHHmmss')
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
        $problemObject = $jsonObject 
        $tmp = $rooms | Where-Object { $_.id -eq $jsonObject.id } | Get-SingleElement
        $this.room = $tmp
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
