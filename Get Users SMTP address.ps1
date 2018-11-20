$SRId = "SR3667"
$SRId = "SR5504"

Import-Module SMLets

$SRClass = Get-SCSMClass System.WorkItem.ServiceRequest$
$AffectedUserRel = Get-SCSMRelationshipClass System.WorkItemAffectedUser$

try{
    $SR = Get-SCSMObject -Class $SRClass -Filter "Id -eq $SRId"
    $AffectedUser = Get-SCSMRelatedObject -SMObject $SR -Relationship $AffectedUserRel
}
catch{
    $EmailAddress = "it-operations@sca.com"
    break
}

If($AffectedUser){
    try{
        $endPoint = Get-SCSMRelatedObject -SMObject $AffectedUser -Relationship $userPerf|?{$_.DisplayName -like '*SMTP'}
        $EmailAddress = $endPoint.TargetAddress
    }
    catch{
        $EmailAddress = "it-operations@sca.com"
        break
    }
}