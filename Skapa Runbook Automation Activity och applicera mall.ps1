Import-Module Smlets 
$AAClass = Get-SCSMClass Microsoft.SystemCenter.Orchestrator.RunbookAutomationActivity$

$Params = @{
    ID="RB{0}"
    Title="Runbook Automated Activity " + [datetime]::now
    } 
    
$o = New-SCSMObject -Class $AAClass -PropertyHashtable $Params -pass 

$Projection = Get-SCSMTypeProjection Microsoft.SystemCenter.Orchestrator.RunbookAutomationActivity.Projection

$title = $o.Id
$AAObject = get-scsmobjectprojection $projection -filter "Id -eq '$title'" 
$AAObject

$template = Get-SCSMObjectTemplate -displayname "Norma SW deployment initiation"
$template

$AAObject.__base.ApplyTemplate($template)
$AAObject.__base.Commit()