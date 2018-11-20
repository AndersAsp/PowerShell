$CLOSEDAYS = 7

Function Get-GMTDate () {
$LocalZone = [System.TimeZoneInfo]::Local
$Hours = [system.Math]::Abs($LocalZone.BaseUtcOffset.Hours)
$Mins = [System.Math]::Abs($LocalZone.BaseUtcOffset.Minutes)
if ($LocalZone.IsdaylightSavingTime([system.DateTime]::Now)) { $Hours -= 1 }
$TimeDiff = New-Object TimeSpan 0,$Hours,$Mins,0,0
(Get-Date).Subtract($TimeDiff)
}

Import-Module -Name SMLets

$GMTDate = Get-GMTDate
$ResolvedStatus = Get-SCSMEnumeration IncidentStatusEnum.Resolved$

$IncidentClass = Get-SCSMClass -Name System.WorkItem.Incident$

$ResolvedIncidents = Get-SCSMObject -class $IncidentClass | where{ $_.Status -eq $ResolvedStatus -and ($_.ResolvedDate).AddDays($CLOSEDAYS) -lt $GMTDate }
$ResolvedIncidents

if ($ResolvedIncidents -ne $null) {
$ResolvedIncidents |Set-SCSMIncident -Status Closed -comment "Auto-closed after $CLOSEDAYS days of inactivity"
}