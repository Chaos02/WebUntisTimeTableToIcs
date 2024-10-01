<#
.SYNOPSIS
    Converts a WebUntis timetable to an ICS calendar file.

.DESCRIPTION
    This script retrieves timetable data from the WebUntis API and converts it into an ICS calendar file format. 
    It allows specifying a date range for the timetable data.

.PARAMETER baseUrl
    The base URL of the WebUntis API.

.PARAMETER elementType
    The type of element to filter by (default is 1, this should return a timetable).

.PARAMETER elementId
    The classes timetable ID.

.PARAMETER dates
    An array of dates (either as strings or DateTime objects) for which to retrieve timetable data. 
    The default is the current week and the next three weeks.

.PARAMETER OutputFilePath
    The file path where the ICS file will be saved. The default is "calendar.ics".

.PARAMETER cookie
    The cookie value for the WebUntis session.

.PARAMETER tenantId
    The tenant ID for the WebUntis session.

.EXAMPLE
    .\timeTableToIcs.ps1 -baseUrl "your.webuntis.url" -elementType 1 -elementId 12345 -dates "2023-01-01", "2023-01-08" -OutputFilePath "mycalendar.ics" -cookie "your_cookie" -tenantId "your_tenant_id"

.NOTES
    Author: Markus Noack
    Date: 2024-09-28
    Version: 1.0
#>

param (
    [ValidateNotNullOrEmpty()]
    [Alias("URL")]
    [string]$baseUrl,
    [int]$elementType = 1,
    [Alias("TimeTableID")]
    [int]$elementId,
    [Parameter(Mandatory = $false)]
    [Alias("Date")]
    [ValidateScript({
            if ($_.GetType().Name -eq 'String') {
                if (-not [datetime]::TryParse($_, [ref] $null)) {
                    throw "Invalid date format. Please provide a valid date string."
                }
            }
            elseif ($_.GetType().Name -ne 'DateTime') {
                throw "Invalid date format. Provide a date string or DateTime object."
            }
            $true
        })]
    [System.Object[]]$dates = @( (@(-7, 0, 7, 14) | ForEach-Object { (Get-Date).AddDays($_) }) ),
    [ValidateNotNullOrEmpty()]
    [string]$OutputFilePath = "calendar.ics",
    [ValidateNotNullOrEmpty()]
    [string]$cookie,
    [ValidateNotNullOrEmpty()]
    [string]$tenantId
)

$datesCount = 0
# Convert any string inputs to DateTime objects
$dates = $dates | ForEach-Object {
    $datescount++
    if ($_ -is [string]) { [datetime]::Parse($_) } else { $_ }
}
if ($datesCount -gt 4) {
    throw "The maximum number of weeks is 4. (Limit by WebUntis API)"
}

function Get-SingleElement {
    param (
        [Parameter(ValueFromPipeline = $true)]#, ValueFromPipelineByPropertyName = $true)]
        [System.Object[]]$collection
    )

    process {
        $elements = @()
        foreach ($item in $collection) { # Unify for generic access
            $elements += $item
        }

        if ($elements.Length -eq 0) {
            throw [System.InvalidOperationException]::new("No elements match the predicate. Call stack: $((Get-PSCallStack | Out-String).Trim())")
        }
        elseif ($elements.Length -gt 1) {
            throw [System.InvalidOperationException]::new("More than one element matches the predicate. Call stack: $((Get-PSCallStack | Out-String).Trim())")
        }

        return $elements[0]
    }
}

$headers = @{
    "authority"                 = "$baseUrl"
    "accept"                    = "application/json"
    "accept-encoding"           = "gzip, deflate, br, zstd"
    "accept-language"           = "de-DE,de;q=0.9,en-US;q=0.8,en;q=0.7"
    "cache-control"             = "max-age=0"
    "dnt"                       = "1"
    "pragma"                    = "no-cache"
    "priority"                  = "u=0, i"
    "sec-ch-ua"                 = "`"Google Chrome`";v=`"131`", `"Chromium`";v=`"131`", `"Not_A Brand`";v=`"24`""
    "sec-ch-ua-mobile"          = "?0"
    "sec-ch-ua-platform"        = "`"Windows`""
    "sec-fetch-dest"            = "document"
    "sec-fetch-mode"            = "navigate"
    "sec-fetch-site"            = "none"
    "sec-fetch-user"            = "?1"
    "upgrade-insecure-requests" = "1"
}

$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$session.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
$session.Cookies.Add((New-Object System.Net.Cookie("schoolname", "`"$cookie`"", "/", "$baseUrl")))
$session.Cookies.Add((New-Object System.Net.Cookie("Tenant-Id", "`"$tenantId`"", "/", "$baseUrl")))
$session.Cookies.Add((New-Object System.Net.Cookie("schoolname", "`"$cookie==`"", "/", "$baseUrl")))
$session.Cookies.Add((New-Object System.Net.Cookie("Tenant-Id", "`"$tenantId`"", "/", "$baseUrl")))
$session.Cookies.Add((New-Object System.Net.Cookie("schoolname", "`"$cookie==`"", "/", "$baseUrl")))
$session.Cookies.Add((New-Object System.Net.Cookie("Tenant-Id", "`"$tenantId`"", "/", "$baseUrl")))
$session.Cookies.Add((New-Object System.Net.Cookie("traceId", "9de4710537aa097594b039ee3a591cfc22a6dd99", "/", "$baseUrl")))
$session.Cookies.Add((New-Object System.Net.Cookie("JSESSIONID", "B9ED9B2D36BE7D25A7A9EF21E8144D3F", "/", "$baseUrl")))

$periods = [System.Collections.Generic.List[PeriodEntry]]::new()
$courses = [System.Collections.Generic.List[Course]]::new()
$rooms = [System.Collections.Generic.List[Room]]::new()
$legende = [System.Collections.Generic.List[PeriodTableEntry]]::new();

foreach ($date in $dates) {

    Write-Host "Getting Data for week of $($date.ToString("yyyy-mm-dd"))"

    $url = "https://$baseUrl/WebUntis/api/public/timetable/weekly/data?elementType=$elementType&elementId=$elementId&date=$($date.ToString("yyyy-MM-dd"))&formatId=14"

    $response = Invoke-WebRequest -UseBasicParsing -Uri $url -Method Get -WebSession $session -Headers $headers
    $object = $response | ConvertFrom-Json -ErrorAction Stop

    $class = [PeriodTableEntry]

    try {
        $object.data.result.data.elements | ForEach-Object {
            if ($legende.FindAll({ param($e) $e.id -eq $_.id -and $e.type -eq $_.type }).Count -eq 0) { # prevent duplicates
                $legende.Add([PeriodTableEntry]::new($_)) 
            } 
        }
        $legende | Where-Object { $_.type -eq 4 } | ForEach-Object {
            if ($rooms.FindAll({ param($e) $e.id -eq $_.id}).Count -eq 0) { # prevent duplicates
                $rooms.Add([Room]::new($_)) 
            } 
        }
        $legende | Where-Object { $_.type -eq 3 } | ForEach-Object {
            if ($courses.FindAll({ param($e) $e.id -eq $_.id}).Count -eq 0) { # prevent duplicates
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
            }
            catch { #[FormatException] {
                Write-Error ($element | Format-List | Out-String)
                throw
            }
        }
    }
    catch [FormatException] {
        Write-Error "Invalid Response regarding datetime format:"
        throw
        exit 1
    }

}

$periods = $periods | Sort-Object -Property startTime

$calendarEntries = [System.Collections.Generic.List[IcsEvent]]::new()

# Iterate over each period and create calendar entries
foreach ($period in $periods) {
    $calendarEntries.Add([IcsEvent]::new($period))
}

# Get all properties except StartTime and EndTime
$properties = $calendarEntries | Get-Member -MemberType Properties | Where-Object {
    $_.Name -ne 'StartTime' -and $_.Name -ne 'EndTime' -and $_.Name -ne 'Description'
} | Select-Object -ExpandProperty Name

# Use Select-Object to reorder properties and add calculated properties
$calendarEntries | Select-Object (@(
        @{ Name = 'StartTimeF'; Expression = { [DateTime]::ParseExact($_.StartTime, "yyyyMMddTHHmmss", $null).ToString("dd.MM.yy HH:mm") } },
        @{ Name = 'EndTimeF'; Expression = { [DateTime]::ParseExact($_.EndTime, "yyyyMMddTHHmmss", $null).ToString("dd.MM.yy HH:mm") } }
    ) + $properties + @{ 
        Name       = 'DescriptionF'; 
        Expression = { 
            $parts = $_.Description -split ';'
            $firstPart = $parts[0].Trim()
            $secondPart = if ($parts.Count -gt 1) { $parts[1].Trim() } else { "" }
            
            if ($secondPart -ne "") {
                $sourceIndex = $secondPart.IndexOf("source:")
                if ($sourceIndex -ne -1) {
                    $source = $secondPart.Substring($sourceIndex + 7).Trim()
                    "$firstPart; source: $source"
                }
                else {
                    $firstPart
                }
            }
            else {
                $firstPart
            }
        } 
    }
) | Format-Table -AutoSize


$IcsEntries = [System.Collections.Generic.List[string]]::new()
foreach ($icsEvent in $calendarEntries) {
    $IcsEntries += $icsEvent.ToIcsEntry()
}
   

# Create the .ics file content
$icsContent = @"
BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Chaos_02//WebUntisToIcs//EN
X-WR-CALNAME:$($class.displayname)
$(($IcsEntries -join "`n"))
END:VCALENDAR
"@

try {
    if ($OutputFilePath) {
        # Write the .ics content to a file
        Set-Content -Path $OutputFilePath -Value $icsContent
        Write-Output "ICS file created at $((Get-Item -Path $OutputFilePath).FullName)"
    }
    else {
        # Write the .ics content to a variable
        $icsVariable = $icsContent
        Write-Output $icsVariable
        return $icsVariable
    }
}
catch {
    Write-Error "An error occurred while creating the ICS file: $_"
    throw
}

####### Class definitions #######

class IcsEvent {
    [string]$StartTime
    [string]$EndTime
    [string]$Location
    [string]$Summary
    [string]$Description
    [string]$Status
    [string]$Category

    IcsEvent([PeriodEntry]$period) {
        $this.startTime = $period.startTime.ToString("yyyyMMddTHHmmss")
        $this.endTime = $period.endTime.ToString("yyyyMMddTHHmmss")
        $this.location = $period.room.room.longName
        $this.summary = $period.course.course.longName
        $this.description = $period.substText
        if ($null -ne $period.rescheduleInfo) {
            $this.description += "`nReschedule:`n" + $period.rescheduleInfo.ToString()
        }
        $this.status = switch ($period.cellState) {
            "STANDARD" { "CONFIRMED" }
            "ADDITIONAL" { "TENTATIVE" }
            "CANCEL" { "CANCELLED" }
            default { "CONFIRMED" }
        }
        $this.category = switch ($period.lessonCode) {
            "UNTIS_ADDITIONAL" { "Additional" }
            default { $_ }
        }
    }

    [string] ToIcsEntry() {
        return @"
BEGIN:VEVENT
DTSTART:$($this.StartTime)
DTEND:$($this.EndTime)
LOCATION:$($this.Location)
SUMMARY:$($this.Summary)
DESCRIPTION:$($this.Description)
STATUS:$($this.Status)
CATEGORIES:$($this.Category)
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
        $this.startTime = [datetime]::ParseExact($jsonObject.date.ToString(), "yyyyMMdd", $null).Add([timespan]::ParseExact($jsonObject.startTime.ToString().PadLeft(4, '0'), "hhmm", $null))
        $this.endTime = $this.date().Add([timespan]::ParseExact($jsonObject.endTime.ToString().PadLeft(4, '0'), "hhmm", $null))
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
    [bool]$isEvent
    [rescheduleInfo]$rescheduleInfo
    [int]$roomCapacity
    [int]$studentCount

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
        $this.startTime = [datetime]::ParseExact($jsonObject.date.ToString(), "yyyyMMdd", $null).Add([timespan]::ParseExact($jsonObject.startTime.ToString().PadLeft(4, '0'), "hhmm", $null))
        $this.endTime = $this.date().Add([timespan]::ParseExact($jsonObject.endTime.ToString().PadLeft(4, '0'), "hhmm", $null))
        $this.room = [RoomEntry]::new(($jsonObject.elements | Where-Object { $_.type -eq 4 } | Get-SingleElement), $rooms)
        $this.course = [CourseEntry]::new(($jsonObject.elements | Where-Object { $_.type -eq 3 } | Get-SingleElement), $courses)
        $this.studentGroup = $jsonObject.studentGroup
        $this.code = $jsonObject.code
        $this.cellState = $jsonObject.cellState
        $this.priority = $jsonObject.priority
        $this.isStandard = $jsonObject.is.standard
        $this.isEvent = $jsonObject.is.event
        if ($null -ne $jsonObject.rescheduleInfo) {
            $this.rescheduleInfo = [rescheduleInfo]::new($jsonObject.rescheduleInfo)
        }
        else {
            $this.rescheduleInfo = $null
        }
        $this.roomCapacity = $jsonObject.roomCapacity
        $this.studentCount = $jsonObject.studentCount
    }

    [string] ToString() {
        return "ID: $($this.id), Lesson ID: $($this.lessonId), Lesson Number: $($this.lessonNumber), Lesson Code: $($this.lessonCode), Lesson Text: $($this.lessonText), Period Text: $($this.periodText), Has Period Text: $($this.hasPeriodText), Period Info: $($this.periodInfo), Period Attachments: $($this.periodAttachments), Subst Text: $($this.substText), Date: $($this.date), Start Time: $($this.startTime), End Time: $($this.endTime), Elements: $($this.elements), Student Group: $($this.studentGroup), Code: $($this.code), Cell State: $($this.cellState), Priority: $($this.priority), Is Standard: $($this.isStandard), Is Event: $($this.isEvent), Room Capacity: $($this.roomCapacity), Student Count: $($this.studentCount)"
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
            throw [System.ArgumentException]::new("The provided object is not a room.")
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
            throw [System.ArgumentException]::new("The provided object is not a course.")
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
