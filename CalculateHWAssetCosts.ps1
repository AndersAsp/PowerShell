Import-Module Smlets
$HWAssetClass = Get-SCSMClass Cireson.AssetManagement.HardwareAsset$
$Assets = Get-SCSMObject -Class $HWAssetClass -Filter "DisplayName -like 'A183*'"
$Now = Get-Date


Foreach ($Asset in $Assets) {
    $HWCost = $NULL
    $SvcCost = $NULL
    $TotalCost = $NULL
    $PropertyHash = @{}

    if ($Asset.PriskodHardware -ne $NULL -and $Asset.Cost -ne $NULL) {
        $EnumValue = Get-SCSMEnumeration $Asset.PriskodHardware
        
        switch ($EnumValue.DisplayName) { 
            "HV1" {
                    $Months = 36
                    
                    If ($Asset.ReceivedDate.AddMonths($Months) -lt $Now -and !$Asset.AvskrivningsDatum) {
                        $PropertyHash.Add("AvskrivningsDatum",$Asset.ReceivedDate.AddMonths($Months))
                        $HWCost = 0
                    } elseif (!$Asset.Cost) {
                        $HWCost = ($Asset.Cost / $Months)
                    }
            }
            "HV2" {
                    $Months = 24
                    $HWCost = ($Asset.Cost / $Months)
                    If ($Asset.ReceivedDate.AddMonths($Months) -lt $Now -and !$Asset.AvskrivningsDatum) {
                        $PropertyHash.Add("AvskrivningsDatum",$Asset.ReceivedDate.AddMonths($Months))
                        $HWCost = 0
                    } elseif (!$Asset.Cost) {
                        $HWCost = ($Asset.Cost / $Months)
                    }
            }
            "HVD" {
                    $Months = 1
                    $HWCost = ($Asset.Cost / $Months)
                    If ($Asset.ReceivedDate.AddMonths($Months) -lt $Now -and !$Asset.AvskrivningsDatum) {
                        $PropertyHash.Add("AvskrivningsDatum",$Asset.ReceivedDate.AddMonths($Months))
                        $HWCost = 0
                    } elseif (!$Asset.Cost) {
                        $HWCost = ($Asset.Cost / $Months)
                    }
            }
        }
        
        If ($HWCost) {$HWCost = [Math]::Round($HWCost, 2)}

    }

    if ($Asset.PriskodService -ne $NULL) {
        $EnumValue = Get-SCSMEnumeration $Asset.PriskodService

        switch ($EnumValue.DisplayName) { 
            "ADM1" {$SvcCost = 234} 
            "UTB1" {$SvcCost = 93} 
            "UTB2" {$SvcCost = 54} 
            "SER1" {$SvcCost = 31} 
            "SER2" {$SvcCost = 40} 
            "SER3" {$SvcCost = 9} 
            "SER4" {$SvcCost = 18} 
            "INFO1" {$SvcCost = 104}  
            "UTB2F" {$SvcCost = 74} 
            "ADM1F" {$SvcCost = 253} 
            "UTB1F" {$SvcCost = 112} 
            "HYR1" {$SvcCost = 100} 
            "OMA1" {$SvcCost = 18}  
            "SER0" {$SvcCost = 0}  
        }

        If ($Asset.KostnadService -eq $SvcCost) {$SvcCost = $NULL}
    }

    $NewDisplayName = $Asset.Name + " (" + $Asset.AssetTagHBG + ")"

    if ($Asset.DisplayName -ne $NewDisplayName) {
        $PropertyHash.Add("DisplayName",$NewDisplayName)
    }


    if($HWCost -or $HWCost -eq 0) {$PropertyHash.Add("KostnadHW",$HWCost)}
    if($SvcCost -or $SvcCost -eq 0) {$PropertyHash.Add("KostnadService",$SvcCost)}
    if($HWCost -and $SvcCost) {
        [decimal]$TotalCost = [decimal]$HWCost+[decimal]$SvcCost
        $PropertyHash.Add("KostnadTotal",$TotalCost)
    } elseif($HWCost -and !$SvcCost) {
        [decimal]$TotalCost = $HWCost
        $PropertyHash.Add("KostnadTotal",$TotalCost)
    } elseif(!$HWCost -and $SvcCost) {
        [decimal]$TotalCost = $SvcCost
        $PropertyHash.Add("KostnadTotal",$TotalCost)
    }


    #Write-host "-----------------"

    if($PropertyHash.Count -gt 0) {
        #$Asset.DisplayName
        #$Asset.Id
        #$PropertyHash
        Set-SCSMObject -SMObject $Asset -PropertyHashtable $PropertyHash

    }

}