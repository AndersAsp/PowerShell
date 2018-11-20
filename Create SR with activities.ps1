Import-Module SMLets

$SRClass = Get-SCSMClass System.WorkItem.ServiceRequest$
$MAClass = Get-SCSMClass System.WorkItem.Activity.ManualActivity$

$UrgencyLowEnum = Get-SCSMEnumeration ServiceRequestUrgencyEnum.Low$
$PriorityLowEnum = Get-SCSMEnumeration ServiceRequestPriorityEnum.Low$
$SourceOtherEnum = Get-SCSMEnumeration ServiceRequestSourceEnum.Other$
$StatusNewEnum = Get-SCSMEnumeration ServiceRequestStatusEnum.New$

$SRPropertyHash = @{"Id" = "SR{0}";
                    "Title" = "Monthly restore tests";
                    "Description" = "This is a recurring Service Request to ensure that we do restore test of our backups";
                    "Urgency" = $UrgencyLowEnum;
                    "Priority" = $PriorityLowEnum;
                    "Source" = $SourceOtherEnum;
                    "Status" = $StatusNewEnum;
                    }

$MAPropertyHash = @{"Id" = "MA{0}";
                    "SequenceID" = "0";
                    "Title" = "Perform restore test";
                    "Description" = "Perform a restore test of a random backup that hasn't been tested for a while";
                    }

$MAPropertyHash2 = @{"Id" = "MA{0}";
                    "SequenceID" = "1";
                    "Title" = "Create and deliver report to manager";
                    "Description" = "Create a report of the results of the restore tests and deliver it to the system owner and your manager";
                    }

# Create the object projection with properties
$Projection = @{__CLASS = "System.WorkItem.ServiceRequest";
                __OBJECT = $SRPropertyHash;
                Activity = @{__CLASS = $MAClass.Name;
                             __OBJECT = $MAPropertyHash;
                             },
                           @{__CLASS = $MAClass.Name;
                             __OBJECT = $MAPropertyHash2;
                             }
                }


New-SCSMObjectProjection -Type "System.WorkItem.ServiceRequestAndActivityViewProjection" -Projection $Projection