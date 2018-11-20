Import-Module SMLets

$WorkItemId = "SR3317"

$SRClass = Get-SCSMClass System.WorkItem.ServiceRequest$
$IRClass = Get-SCSMClass System.WorkItem.Incident$

$AffectedUserRel = Get-SCSMRelationshipClass System.WorkItemAffectedUser$
$AffectedCIRel = Get-SCSMRelationshipClass System.WorkItemAboutConfigItem$
$AssignedToRel = Get-SCSMRelationshipClass System.WorkItemAssignedToUser$
$ContainsActivityRel = Get-SCSMRelationshipClass System.WorkItemContainsActivity$

Function Convert-IRtoSR($WorkItemId){
    
    # Get Incident 
    $IR = Get-SCSMObject -Class $IRClass -Filter "Id -eq $WorkItemId"

    # Ensure that we got the Incident 
    If($IR.count -ne 1){
        Throw "Unable to retrieve incident with ID: $WorkItemId"
    }

    # Get all the needed properties
    $Title = $IR.Title
    $Description = "[Ärendet har konverterats från $($IR.Id)]`n" + $IR.Description
    $AltContact = $IR.ContactMethod
    $Notification = $IR.Notification
    $OnSite = $IR.OnSite

    # -- Get Enumerations
    # ------------------------------------------------------------------------------------

    # Urgency
    $IRUrgency = $IR.Urgency.DisplayName

    switch ($IRUrgency)
    {
	    "Low" {$SRUrgency = Get-SCSMEnumeration -Name ServiceRequestUrgencyEnum.Low}
        "Låg" {$SRUrgency = Get-SCSMEnumeration -Name ServiceRequestUrgencyEnum.Low}
	    "Medium" {$SRUrgency = Get-SCSMEnumeration -Name ServiceRequestUrgencyEnum.Medium}
	    "High" {$SRUrgency = Get-SCSMEnumeration -Name ServiceRequestUrgencyEnum.High}
	    "Hög" {$SRUrgency = Get-SCSMEnumeration -Name ServiceRequestUrgencyEnum.High}
	    default {$SRUrgency = Get-SCSMEnumeration -Name ServiceRequestUrgencyEnum.Low}
    }
    
    # Classification / Area
    $IRClassification = $IR.Classification.DisplayName
    $SRArea = Get-SCSMEnumeration | ?{$_.DisplayName -eq $IRClassification -and $_.Parent.Id -eq "dac418a9-17c2-f0b5-8e53-3ddec62092ab"}
    If($SRArea.count -ne 1){
        # If several hits, or no hits, default to Software
        $SRArea = Get-SCSMEnumeration Enum.e8192569ca4c4297b933455c4846ce66
    }

    # Support Group
    $IRSupportGroup = $IR.TierQueue.DisplayName
    $SRSupportGroup = Get-SCSMEnumeration | ?{$_.DisplayName -eq $IRSupportGroup -and $_.Parent.Id -eq "23c243f6-9365-d46f-dff2-03826e24d228"}
    If($SRSupportGroup.count -ne 1){
        # If several hits, or no hits, default to Dispatcher
        $SRSupportGroup = Get-SCSMEnumeration Enum.9bd76cabed234bc1a58fa30ef027f609
    }

    # Source 
    $IRSource = $IR.Source.DisplayName

    switch ($IRSource)
    {
	    "Console" {$SRSource = Get-SCSMEnumeration -Name Enum.dc07afd99ba64c57b72761fe5d574d02}
        "Konsol" {$SRSource = Get-SCSMEnumeration -Name Enum.dc07afd99ba64c57b72761fe5d574d02}
	    "E-Mail" {$SRSource = Get-SCSMEnumeration -Name ServiceRequestSourceEnum.Email}
        "E-post" {$SRSource = Get-SCSMEnumeration -Name ServiceRequestSourceEnum.Email}
	    "Portal" {$SRSource = Get-SCSMEnumeration -Name ServiceRequestSourceEnum.Portal}
	    "Other" {$SRSource = Get-SCSMEnumeration -Name ServiceRequestSourceEnum.Other}
	    default {$SRSource = Get-SCSMEnumeration -Name Enum.dc07afd99ba64c57b72761fe5d574d02}
    }

    # Priority

    $IRPrio = $IR.Priority

    switch ($IRPrio){
        "1" {$SRPrio = Get-SCSMEnumeration -Name ServiceRequestPriorityEnum.High}
        "2" {$SRPrio = Get-SCSMEnumeration -Name ServiceRequestPriorityEnum.Medium}
        "3" {$SRPrio = Get-SCSMEnumeration -Name ServiceRequestPriorityEnum.Low}
        default {$SRPrio = Get-SCSMEnumeration -Name ServiceRequestPriorityEnum.Low}
    }

    # ------------------------------------------------------------------------------------

    # Get Affected User
    $AffectedUser = Get-SCSMRelatedObject -SMObject $IR -Relationship $AffectedUserRel
    
    # Get Affected CIs
    $AffectedCIs = Get-SCSMRelatedObject -SMObject $IR -Relationship $AffectedCIRel

    # Get Assigned User 
    $AssignedUser = Get-SCSMRelatedObject -SMObject $IR -Relationship $AssignedToRel

    # -------------------------------------------------
    # -- All information gathered, lets create the SR
    # -------------------------------------------------

    $SRPropertyHash = @{"Id" = "SR{0}";
                        "Title" = $Title;
                        "Description" = $Description;
                        "ContactMethod" = $AltContact;
                        "Notification" = $Notification;
                        "Onsite" = $OnSite;
                        "LTVArea" = $SRArea;
                        "Source" = $SRSource;
                        "Urgency" = $SRUrgency;
                        "SupportGroup" = $SRSupportGroup;
                        "Priority" = $SRPrio;
                        }

    $NewSR = New-SCSMObject -Class $SRClass -PropertyHashtable $SRPropertyHash -PassThru

    If(!$NewSR){
        throw "Unable to create Service Request"
    }

    # Set Affected User
    If($AffectedUser){
        New-SCSMRelationshipObject -Relationship $AffectedUserRel -Source $NewSR -Target $AffectedUser -Bulk
    }

    # Set Assigned User
    If($AssignedUser){
        New-SCSMRelationshipObject -Relationship $AssignedToRel -Source $NewSR -Target $AssignedUser -Bulk
    }

    # Set Affected CIs
    If($AffectedCIs){
        Foreach($CI in $AffectedCIs){
            New-SCSMRelationshipObject -Relationship $AffectedCIRel -Source $NewSR -Target $CI -Bulk
        }
    }

    # Resolve the Incident that has been converted
    $IRResolutionCategory = Get-SCSMEnumeration IncidentResolutionCategoryEnum.Cancelled$
    $IRResolutionDescription = "Ärendet har konverterats till ett nytt ärende med ID: $($NewSR.Id)"
    $IRStatus = Get-SCSMEnumeration IncidentStatusEnum.Resolved$
    $IRResolvedDate = Get-Date
    $IRUpdatedDescription = "[Ärendet har konverterats till en Service Request med ID: $($NewSR.Id)]`n" + $IR.Description

    $IRPropertyHash = @{"Status" = $IRStatus;
                        "ResolutionCategory" = $IRResolutionCategory;
                        "ResolutionDescription" = $IRResolutionDescription;
                        "ResolvedDate" = $IRResolvedDate;
                        "Description" = $IRUpdatedDescription;
                        }

    Set-SCSMObject -SMObject $IR -PropertyHashtable $IRPropertyHash

    Return $NewSR

}

Function Convert-SRtoIR($WorkItemId){

    # Get Service Request 
    $SR = Get-SCSMObject -Class $SRClass -Filter "Id -eq $WorkItemId"

    # Ensure that we got the Service Request 
    If($SR.count -ne 1){
        Throw "Unable to retrieve Service Request with ID: $WorkItemId"
    }

    # Get all the needed properties
    $Title = $SR.Title
    $Description = "[Ärendet har konverterats från $($SR.Id)]`n" + $SR.Description
    $AltContact = $SR.ContactMethod
    $Notification = $SR.Notification
    $OnSite = $SR.OnSite

    # -- Get Enumerations
    # ------------------------------------------------------------------------------------

    # Urgency
    $SRUrgency = $SR.Urgency.DisplayName

    switch ($SRUrgency)
    {
	    "Low" {$IRUrgency = Get-SCSMEnumeration -Name System.WorkItem.TroubleTicket.UrgencyEnum.Low}
        "Låg" {$IRUrgency = Get-SCSMEnumeration -Name System.WorkItem.TroubleTicket.UrgencyEnum.Low}
	    "Medium" {$IRUrgency = Get-SCSMEnumeration -Name System.WorkItem.TroubleTicket.UrgencyEnum.Medium}
	    "High" {$IRUrgency = Get-SCSMEnumeration -Name System.WorkItem.TroubleTicket.UrgencyEnum.High}
	    "Hög" {$IRUrgency = Get-SCSMEnumeration -Name System.WorkItem.TroubleTicket.UrgencyEnum.High}
	    default {$IRUrgency = Get-SCSMEnumeration -Name System.WorkItem.TroubleTicket.UrgencyEnum.Low}
    }
    
    # Classification / Area
    $SRLTVArea = $SR.LTVArea.DisplayName
    $IRClassification = Get-SCSMEnumeration | ?{$_.DisplayName -eq $SRLTVArea -and $_.Parent.Id -eq "dac418a9-17c2-f0b5-8e53-3ddec62092ab"}
    If($IRClassification.count -ne 1){
        # If several hits, or no hits, default to Software
        $IRClassification = Get-SCSMEnumeration Enum.492c27e6138449ff9681277bc20d73cb
    }

    # Support Group
    $SRSupportGroup = $SR.SupportGroup.DisplayName
    $IRSupportGroup = Get-SCSMEnumeration | ?{$_.DisplayName -eq $SRSupportGroup -and $_.Parent.Id -eq "23c243f6-9365-d46f-dff2-03826e24d228"}
    If($IRSupportGroup.count -ne 1){
        # If several hits, or no hits, default to Dispatcher
        $IRSupportGroup = Get-SCSMEnumeration Enum.859ae6bf99254d9699122fcaec2cffc5
    }

    # Source 
    $SRSource = $SR.Source.DisplayName

    switch ($SRSource)
    {
	    "Console" {$IRSource = Get-SCSMEnumeration -Name IncidentSourceEnum.Console}
        "Konsol" {$IRSource = Get-SCSMEnumeration -Name IncidentSourceEnum.Console}
	    "E-Mail" {$IRSource = Get-SCSMEnumeration -Name IncidentSourceEnum.Email}
        "E-post" {$IRSource = Get-SCSMEnumeration -Name IncidentSourceEnum.Email}
	    "Portal" {$IRSource = Get-SCSMEnumeration -Name IncidentSourceEnum.Portal}
	    "Other" {$IRSource = Get-SCSMEnumeration -Name IncidentSourceEnum.Console}
	    default {$IRSource = Get-SCSMEnumeration -Name IncidentSourceEnum.Console}
    }

    # Priority

    $SRPrio = $SR.Priority

    switch ($SRPrio){
        "Low" {$IRImpact = Get-SCSMEnumeration -Name System.WorkItem.TroubleTicket.ImpactEnum.Low}
        "Låg" {$IRImpact = Get-SCSMEnumeration -Name System.WorkItem.TroubleTicket.ImpactEnum.Low}
        "Medium" {$IRImpact = Get-SCSMEnumeration -Name System.WorkItem.TroubleTicket.ImpactEnum.Medium}
        "High" {$IRImpact = Get-SCSMEnumeration -Name System.WorkItem.TroubleTicket.ImpactEnum.High}
        "Hög" {$IRImpact = Get-SCSMEnumeration -Name System.WorkItem.TroubleTicket.ImpactEnum.High}
        default {$IRImpact = Get-SCSMEnumeration -Name System.WorkItem.TroubleTicket.ImpactEnum.Low}
    }

    # ------------------------------------------------------------------------------------

    # Get Affected User
    $AffectedUser = Get-SCSMRelatedObject -SMObject $SR -Relationship $AffectedUserRel
    
    # Get Affected CIs
    $AffectedCIs = Get-SCSMRelatedObject -SMObject $SR -Relationship $AffectedCIRel

    # Get Assigned User 
    $AssignedUser = Get-SCSMRelatedObject -SMObject $SR -Relationship $AssignedToRel

    # -------------------------------------------------
    # -- All information gathered, lets create the SR
    # -------------------------------------------------

    $IRPropertyHash = @{"Id" = "IR{0}";
                        "Title" = $Title;
                        "Description" = $Description;
                        "ContactMethod" = $AltContact;
                        "Notification" = $Notification;
                        "Onsite" = $OnSite;
                        "Classification" = $IRClassification;
                        "Source" = $IRSource;
                        "Urgency" = $IRUrgency;
                        "TierQueue" = $IRSupportGroup;
                        "Impact" = $IRImpact;
                        }

    $NewIR = New-SCSMObject -Class $IRClass -PropertyHashtable $IRPropertyHash -PassThru

    If(!$NewIR){
        throw "Unable to create Incident"
    }

    # Set Affected User
    If($AffectedUser){
        New-SCSMRelationshipObject -Relationship $AffectedUserRel -Source $NewIR -Target $AffectedUser -Bulk
    }

    # Set Assigned User
    If($AssignedUser){
        New-SCSMRelationshipObject -Relationship $AssignedToRel -Source $NewIR -Target $AssignedUser -Bulk
    }

    # Set Affected CIs
    If($AffectedCIs){
        New-SCSMRelationshipObject -Relationship $AffectedCIRel -Source $NewIR -Target $AffectedCIs -Bulk
    }


    # Complete the Service Request that has been converted
    $SRImplementationResults = Get-SCSMEnumeration ServiceRequestImplementationResultsEnum.Canceled$
    $SRNotes = "Ärendet har konverterats till ett nytt ärende med ID: $($NewIR.Id)"
    $SRStatus = Get-SCSMEnumeration ServiceRequestStatusEnum.Completed$
    $SRCompletedDate = Get-Date
    $SRUpdatedDescription = "[Ärendet har konverterats till en Incident med ID: $($NewIR.Id)]`n" + $SR.Description

    $SRPropertyHash = @{"Status" = $SRStatus;
                        "ImplementationResults" = $SRImplementationResults;
                        "Notes" = $SRNotes;
                        "CompletedDate" = $SRCompletedDate;
                        "Description" = $SRUpdatedDescription;
                        }

    $Activities = Get-SCSMRelatedObject -SMObject $SR -Relationship $ContainsActivityRel

    If($Activities){
        Foreach($Act in $Activities){
            Set-SCSMObject -SMObject $Act -Property Skip -Value $TRUE
        }
    }
    
    Set-SCSMObject -SMObject $SR -PropertyHashtable $SRPropertyHash

    Return $NewIR

}


If($WorkItemId.Substring(0,2) -eq "IR"){
    Convert-IRtoSR $WorkItemId
} elseif($WorkItemId.Substring(0,2) -eq "SR") {
    Convert-SRtoIR $WorkItemId
} else {
    Throw "No IR/SR prefix found!"
}