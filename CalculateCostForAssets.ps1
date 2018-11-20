Import-Module Smlets
$HWAssetClass = Get-SCSMClass Cireson.AssetManagement.HardwareAsset$
$Assets = Get-SCSMObject -Class $HWAssetClass

Foreach ($Asset in $Assets) {
    $HWCost = $NULL
    $SvcCost = $NULL
    $TotalCost = $NULL
    $PropertyHash = @{}

    if ($Asset.PriskodHardware -ne $NULL -and $Asset.Cost -ne $NULL) {
        $EnumValue = Get-SCSMEnumeration $Asset.PriskodHardware
        
        switch ($EnumValue.DisplayName) { 
            "HV1" {$HWCost = ($Asset.Cost / 36)}
            "HV2" {$HWCost = ($Asset.Cost / 24)} 
            "HV3" {$HWCost = ($Asset.Cost / 1)} 
        }
        $HWCost = [Math]::Round($HWCost, 2)
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
        }
    }

    if($HWCost) {$PropertyHash.Add("KostnadHW",$HWCost)}
    if($SvcCost) {$PropertyHash.Add("KostnadService",$SvcCost)}
    if($HWCost -and $SvcCost) {
        [decimal]$TotalCost = [decimal]$HWCost+[decimal]$SvcCost
        $PropertyHash.Add("TotalKostnad",$TotalCost)
    }

    if($PropertyHash.Count -gt 0) {
        $Asset.DisplayName
        $PropertyHash
        Set-SCSMObject -SMObject $Asset -PropertyHashtable $PropertyHash

    }

}