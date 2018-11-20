Import-Module smlets
$class = Get-SCSMClass System.Service$
$RelClass = Get-SCSMRelationshipClass System.WorkItemAboutConfigItem$

$InitService = Get-SCSMObject -Class $class -Filter "DisplayName -eq 'Service1'"
$TargetService = Get-SCSMObject -Class $class -Filter "DisplayName -eq 'Service2'"

$WIs = Get-SCSMRelationshipObject -TargetRelationship $RelClass -TargetObject $InitService

Foreach ($WI in $WIs) {
    
    $Obj = $WI.SourceObject
    
    $SourceObject = Get-SCSMObject -Id $Obj.Id
    New-SCSMRelationshipObject -Relationship $RelClass -Source $SourceObject -Target $TargetService -Bulk
}