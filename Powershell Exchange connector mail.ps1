# -------------------------------------------------------------------
# -- Functions
# -------------------------------------------------------------------
 
# This function adds a comment to the SR Action Log
Function Add-SRComment {
    param (
        [parameter(Mandatory=$True,Position=0)][Alias('SRObject')]$pSRObject,
        [parameter(Mandatory=$True,Position=1)][Alias('UserComment')]$pUserComment,
        [parameter(Mandatory=$True,Position=2)][Alias('EnteredBy')]$pEnteredBy
    )
 
    # Make sure that the SR Object it passed to the function
    If ($pSRObject) {
 
        # Generate a new GUID for the comment
        $NewGUID = ([guid]::NewGuid()).ToString()
 
        # Create the object projection with properties
        $Projection = @{__CLASS = "System.WorkItem.ServiceRequest";
                        __SEED = $pSRObject;
                        EndUserCommentLog = @{__CLASS = "System.WorkItem.TroubleTicket.UserCommentLog";
                                            __OBJECT = @{Id = $NewGUID;
                                                        DisplayName = $NewGUID;
                                                        Comment = $pUserComment;
                                                        EnteredBy = $pEnteredBy;
                                                        EnteredDate = (Get-Date).ToUniversalTime();
                                            }
                        }
        }
 
        # Create the actual comment
        New-SCSMObjectProjection -Type "System.WorkItem.ServiceRequestProjection" -Projection $Projection
    }
 
}
 
# This function adds an attachment to the SR
Function Add-SRAttachment {
    param (
        [parameter(Mandatory=$True,Position=0)][Alias('SRObject')]$pSRObject,
        [parameter(Mandatory=$True,Position=1)][Alias('UserComment')]$pFullFilePath
    )
 
    # Make sure that the SR Object it passed to the function
    If ($pSRObject) {
 
        # Get filename
        $file = Get-Item $pFullFilePath
 
        # Open file
        $fMode = [System.IO.FileMode]::Open
        $fRead = New-Object System.IO.FileStream $pFullFilePath, $fMode
 
        # Generate a new GUID for the comment
        $NewGUID = ([guid]::NewGuid()).ToString()
 
        # Create the object projection with properties
        $Projection = @{__CLASS = "System.WorkItem.ServiceRequest";
                        __SEED = $pSRObject;
                        FileAttachment = @{__CLASS = "System.FileAttachment";
                                            __OBJECT = @{Id = $NewGUID;
                                                        DisplayName = $file.Name;
                                                        Description = $file.Name;
                                                        Extension = $file.Extension;
                                                        Size = $file.Length;
                                                        AddedDate = (Get-Date).ToUniversalTime();
                                                        Content = $fRead;
                                            }
                        }
        }
 
        # Create the actual attachment
        New-SCSMObjectProjection -Type "ServiceRequestProjection.FileAttachment" -Projection $Projection
        $fRead.Close()
 
        Remove-Item $pFullFilePath
    }
 
}
 
# This function will return the user object of the specified Email Address. If a matching user isn't found, the function will create an Internal user within the SCSM CMDB
Function Get-UserByEmail {
    param (
        [parameter(Mandatory=$True,Position=0)][Alias('EmailAddress')]$pEmailAddress,
        [parameter(Mandatory=$False,Position=1)][Alias('Name')]$pName
    )
 
    # Get all the classes and relationships
    $UserPreferenceClass = Get-SCSMClass System.UserPreference$
    $UserPrefRel = Get-SCSMRelationshipClass System.UserHasPreference$
    $ADUser = Get-SCSMClass Microsoft.Ad.User$
 
    # Check if the user exist
    $SMTPObj = Get-SCSMObject -Class $UserPreferenceClass -Filter "DisplayName -like '*SMTP'" | ?{$_.TargetAddress -eq $pEmailAddress}
 
    If ($SMTPObj) {
        # A matching user exist, return the object
        $RelObj = Get-SCSMRelationshipObject -TargetRelationship $UserPrefRel -TargetObject $SMTPObj
        Return $AffectedUser = Get-scsmobject -Id ($RelObj[0].SourceObject).Get_Id()
    } else {
        # A matching user does NOT exist. Do some processing to get the needed properties for creating the user object
        If (!$pName -or $pName -eq '') {
            $Name = $pEmailAddress.Substring(0,$pEmailAddress.IndexOf("@"))
            $UserName = $Name.Replace(",","")
            $UserName = $UserName.Replace(" ","")
        } else {
            $Name = $pName
            $UserName = $Name.Replace(",","")
            $UserName = $UserName.Replace(" ","")
        }
 
        # Try Username to make sure we have a unique username
        $Loop = $TRUE
        $i = 1
 
        While ($Loop -eq $TRUE) {
            $tempUser = Get-SCSMObject -Class (Get-SCSMClass System.Domain.User$) -Filter "UserName -eq $UserName"
 
            If ($tempUser) {
                $UserName = $UserName + $i
                $i = $i +1
            } elseif ($i -gt 15) {
                Break
            } else {
                $Loop = $False
            }
        }
 
        # Create the Property Hash for the new user object
        $PropertyHash = @{"DisplayName" = $Name;
                            "Domain" = "SMINTERNAL";
                            "UserName" = $UserName;
        }
 
        # Create the actual user object
        $AffectedUser = New-SCSMObject -Class (Get-SCSMClass System.Domain.User$) -PropertyHashtable $PropertyHash -PassThru
 
        # Add the SMTP notification address to the created user object
 
        If ($AffectedUser) {
            $NewGUID = ([guid]::NewGuid()).ToString()
 
            $DisplayName = $pEmailAddress + "_SMTP"
 
            $Projection = @{__CLASS = "System.Domain.User";
                            __SEED = $AffectedUser;
                            Notification = @{__CLASS = "System.Notification.Endpoint";
                                             __OBJECT = @{Id = $NewGUID;
                                             DisplayName = $DisplayName;
                                             ChannelName = "SMTP";
                                             TargetAddress = $pEmailAddress;
                                             Description = $pEmailAddress;
                                             }
                            }
            }
 
            New-SCSMObjectProjection -Type "System.User.Preferences.Projection" -Projection $Projection
 
        }
 
        # Return the created user object
        Return $AffectedUser
 
 
    }
 
}
# -------------------------------------------------------------------
# -------------------------------------------------------------------
# -------------------------------------------------------------------
 
# -- Start of the actual script
# -------------------------------------------------------------------
 
# -- C O N F I G U R E   T H E S E
# -------------------------------------------------------------------
 
# Enter the email addresses to monitor here. Separate the addresses with ";"
$ParseEmailAddresses = 'lonejour.agneta.waktnas@sr.se;lonejour.angela.karlsson@sr.se;lonejour.anneli.g.johansson@sr.se;lonejour.annika.johansson@sr.se;lonejour.ann-marie.arleteg@sr.se;lonejour.christina.bakke@sr.se;lonejour.daniel.qvarnstrom@sr.se;lonejour.eivor.furhoff@sr.se;lonejour.jenny.traner@sr.se;lonejour.majlis.nilsson@sr.se;lonejour.margaretha.larsson@sr.se;lonejour.suzanne.kreuger@sr.se;lonejour.asa.brannholm@sr.se;lonejour.asa.persson@sr.se;logon@sr.se;personalwebben@sr.se;lonejour@sr.se'
 
# Enter path where script can temp store exported emails before they are imported to SCSM
$FilePath = 'C:\SCSM\Scheduled_PowerShell_Scripts\ParseEmail\'
 
# ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
# -------------------------------------------------------------------
 
# Split all the entered email addresses to handle them one by one
$ParseEmailAddresses = $ParseEmailAddresses -split ";"
 
# Import SMLets
Import-Module SMLets
 
# Load Exhange Web Services
$EWS = [Reflection.Assembly]::LoadFile("C:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll")
 
if(!$EWS) { Throw "Unable to load EWS!" }
 
$ExchSvc = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013)
 
# If you want to connect to Exchange using another user, uncomment and configure the line below
#$ExchSvc.Credentials = New-Object Net.NetworkCredential('svcscorchmgmtac01', 'JzuZqhwg0kaMljtteLre', 'ref.sr.se')
$psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.ItemSchema]::MimeContent)
 
# Get the Service Request class
$SRClass = Get-SCSMClass System.WorkItem.ServiceRequest$
 
# Loop through all monitored email addresses
Foreach($ParseEmailAddress in $ParseEmailAddresses) {
 
    #write-host $ParseEmailAddress
 
    # User Autodiscover to find the Exchange service and then connect to it
    $ExchSvc.AutodiscoverUrl($ParseEmailAddress)
    $ExchSvc.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $ParseEmailAddress)
 
    # Get the mappings to Inbox and Deleted Items
    $inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ExchSvc,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
    $deletedItems = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ExchSvc,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems)
 
    # Get all unread emails from the Inbox
    $iv = new-object Microsoft.Exchange.WebServices.Data.ItemView(50)
    $inboxfilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead,$false)
    $msgs = $ExchSvc.FindItems($inbox.Id, $inboxfilter, $iv)
 
    # If any emails is found, process them
    If ($msgs.TotalCount -gt 0) {
 
        # Get the property set of the email
        $psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
        $psPropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text
        $LoadProperties = $ExchSvc.LoadPropertiesForItems($msgs,$psPropertySet)
 
        Foreach ($mail in $msgs.Items) {
 
            # Get nescessary information from the email
            $FromEmail = $mail.From.Address
            $FromName = $mail.From.Name
            $Subject = $mail.Subject
            $Body = $mail.body.text
 
            # Check if User Exist
            $AffectedUser = Get-UserByEmail -pEmailAddress $FromEmail -pName $FromName
 
            # Check if new or update
            # -------------------------------------------------------------------
 
            $IdStart = $Subject.IndexOf("[")+1
            $IdEnd = $Subject.IndexOf("]")
 
            If ($IdEnd -gt 0 -and $IdStart -ge 0) {
                $WI_ID = $Subject.Substring($IdStart,$IdEnd-$IdStart)
 
                If ($WI_ID.Substring(0,2) -eq "SR") {
        
                    $SR = Get-SCSMObject -Class $SRClass -Filter "Id -eq $WI_ID"
 
                    If ($SR) {
                        $Action = "Update"
                    } else {
                        $Action = "Error"
                    }
                } else {
                    $Action = "Error"
                }
    
            } else {
                $Action = "Create"
            }
            # -------------------------------------------------------------------
 
            If ($Action -eq "Create") { 
                
                # Get enumerations
                $StatusEnum = Get-SCSMEnumeration ServiceRequestStatusEnum.Submitted$
                $PriorityEnum = Get-SCSMEnumeration ServiceRequestPriorityEnum.Low$
                $UrgencyEnum = Get-SCSMEnumeration ServiceRequestUrgencyEnum.Low$
                $SourceEnum = Get-SCSMEnumeration  ServiceRequestSourceEnum.Email$
                $AreaEnum = Get-SCSMEnumeration ServiceRequestAreaEnum.Other$
                $LonejourenSupportGroupEnum = Get-SCSMEnumeration Enum.cc1eea17d3144e0093df1a34da6b9294
                $PersonalwebbenSupportGroupEnum = Get-SCSMEnumeration Enum.009af7b1aed0497c81d8723558701c19
                $LogonSupportGroupEnum = Get-SCSMEnumeration Enum.f6b7d4dd66d84385992debd68a4dc1c4
                $Automail = Get-SCSMEnumeration Enum.cc1eea17d3144e0093df1a34da6b9294
 
                # Create the Property Hash table
                $PropertyHash = @{
                    "Id" = "SR{0}";
                    "Title" = $Subject;
                    "Description" = $Body;
                    "Status" = $StatusEnum;
                    "Priority" = $PriorityEnum;
                    "Urgency" = $UrgencyEnum;
                    "Source" = $SourceEnum;
                    "Area" = $AreaEnum;
                    "AutoMail" = $Automail;
                }
                               
                If($ParseEmailAddress -eq 'logon@sr.se') {
                    $PropertyHash.Add("SupportGroup", $LogonSupportGroupEnum)                    
                } elseif($ParseEmailAddress -eq 'personalwebben@sr.se') {
                    $PropertyHash.Add("SupportGroup", $PersonalwebbenSupportGroupEnum)
                } else {
                    $PropertyHash.Add("SupportGroup", $LonejourenSupportGroupEnum)
                }
 
                # Create SR
                $NewSR = New-SCSMObject -Class $SRClass -PropertyHashtable $PropertyHash -PassThru
 
                # Create Relationships
                $NewAffectedUserRelationship = New-SCSMRelationshipObject -Relationship (Get-SCSMRelationshipClass System.WorkItemAffectedUser$) -Source $NewSR -Target $AffectedUser -Bulk -PassThru
                $NewCreatedByRelationship = New-SCSMRelationshipObject -Relationship (Get-SCSMRelationshipClass System.WorkItemCreatedByUser$) -Source $NewSR -Target $AffectedUser -Bulk -PassThru
 
                If($ParseEmailAddress -ne 'logon@sr.se' -and $ParseEmailAddress -ne 'personalwebben@sr.se' -and $ParseEmailAddress -ne 'lonejour@sr.se') {
                
                    # Get Assignee
                    $AssgineeMail = $ParseEmailAddress.Substring($ParseEmailAddress.IndexOf(".")+1,$ParseEmailAddress.Length-$ParseEmailAddress.IndexOf(".")-1)
                    $Assignee = Get-UserByEmail -pEmailAddress $AssgineeMail
                
                    $NewAssignedUserRelationship = New-SCSMRelationshipObject -Relationship (Get-SCSMRelationshipClass System.WorkItemAssignedToUser$) -Source $NewSR -Target $Assignee -Bulk -PassThru
                }
 
                # Save original mail as file
                $mail.load($psPropset)  
                $fileName = $FilePath+$NewSR.Id+"_original_email.eml"  
                $File = new-object System.IO.FileStream($fileName, [System.IO.FileMode]::Create)   
                $File.Write($mail.MimeContent.Content, 0,$mail.MimeContent.Content.Length)  
                $File.Close() 
                
                # Attach file to SR
                Add-SRAttachment -pSRObject $NewSR -pFullFilePath $fileName
 
            } elseif ($Action -eq "Update") {
                # Update SR with End User comment
                # NOTE: Seems like there is a bug in the SCSM console which make these comments unreadable. The comments are visible through the SDK and the Cireson Portal
                Add-SRComment -SRObject $SR -UserComment $Body -EnteredBy $FromName
            } else {
                # ERROR
            }
 
            # Move the processed E-mail to Deleted Items and mark it as read
            $move = $mail.Move($deletedItems.Id)
            $isRead = $mail.IsRead = $true
 
        }
    }
}
