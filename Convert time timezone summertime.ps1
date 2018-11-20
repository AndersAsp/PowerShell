$ScheduledStartDateString = "2012-01-01"
$ScheduledEndDateString = "2012-01-01"
$DowntimeStartDateString = "2012-01-01"
$DowntimeEndDateString = ""



$LocalZone = [System.TimeZoneInfo]::Local
$Hours = [system.Math]::Abs($LocalZone.BaseUtcOffset.Hours)
if ($LocalZone.IsdaylightSavingTime([system.DateTime]::Now)) { $Hours += 1 }

if ($ScheduledStartDateString -ne "" -and $ScheduledStartDateString -ne $NULL) {

    [datetime]$ScheduledStartDate = $ScheduledStartDateString
    $ssd = $ScheduledStartDate.AddHours($Hours)
} else {
    $ssd = "Not specified"
}

if ($ScheduledEndDateString -ne "" -and $ScheduledEndDateString -ne $NULL) {
    
    [datetime]$ScheduledEndDate = $ScheduledEndDateString
    $sed = $ScheduledEndDate.AddHours($Hours)

} else {
    $sed = "Not specified"
}

if ($DowntimeStartDateString -ne "" -and $DowntimeStartDateString -ne $NULL) {
    
    [datetime]$DowntimeStartDate = $DowntimeStartDateString
    $dsd = $DowntimeStartDate.AddHours($Hours)

} else {
    $dsd = "Not specified"
}

if ($DowntimeEndDateString -ne "" -and $DowntimeEndDateString -ne $NULL) {

    [datetime]$DowntimeEndDate = $DowntimeEndDateString
    $ded = $DowntimeEndDate.AddHours($Hours)

} else {
    $ded = "Not specified"
}



$ssd
$sed
$dsd
$ded


