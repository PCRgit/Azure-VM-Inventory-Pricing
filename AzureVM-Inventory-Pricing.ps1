<#
.SYNOPSIS
    Azure VM Inventory with Pricing Report - Multi-Subscription
.DESCRIPTION
    Scans all subscriptions in the tenant, collects VM settings (networking,
    SKU, OS, tags), all attached disks (OS + data), performs live price
    lookups via Azure Retail Prices API, and exports to a formatted Excel file.
.NOTES
    Requirements:
      - Az PowerShell Module     : Install-Module Az
      - ImportExcel Module       : Install-Module ImportExcel
      - Connect-AzAccount before running

# Basic run - outputs to Desktop
.\AzureVM-Inventory-Pricing.ps1

# Custom output path and exclude a subscription
.\AzureVM-Inventory-Pricingps1 -OutputPath "C:\Reports\VMReport.xlsx" -ExcludeSubscriptions @("sub-id-to-skip") -CurrencyCode "USD"    
#>

#Requires -Modules Az.Accounts, Az.Compute, Az.Network, ImportExcel

[CmdletBinding()]
param(
    [string]$OutputPath = "$env:USERPROFILE\Desktop\AzureVM_Inventory_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx",
    [string]$CurrencyCode = "USD",
    [string[]]$ExcludeSubscriptions = @()
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

# ----------------------------
# Pricing cache / helpers
# ----------------------------

$script:PriceCache = @{}

function Invoke-AzRetailPriceQuery {
    param(
        [Parameter(Mandatory = $true)][string]$Uri
    )

    $maxRetries = 4
    $attempt = 0

    do {
        try {
            return Invoke-RestMethod -Uri $Uri -Method Get -ErrorAction Stop
        }
        catch {
            $attempt++
            $msg = $_.Exception.Message

            if ($msg -match 'Too many requests' -and $attempt -lt $maxRetries) {
                $sleepSeconds = 5 * $attempt
                Write-Warning "Retail Pricing API rate limit hit. Retrying in $sleepSeconds second(s)..."
                Start-Sleep -Seconds $sleepSeconds
            }
            else {
                throw
            }
        }
    } while ($attempt -lt $maxRetries)
}

function Get-AzRetailPrice {
    param(
        [Parameter(Mandatory = $true)][string]$Filter,
        [Parameter(Mandatory = $true)][string]$CacheKey
    )

    if ($script:PriceCache.ContainsKey($CacheKey)) {
        return $script:PriceCache[$CacheKey]
    }

    try {
        $encodedFilter = [System.Uri]::EscapeDataString($Filter)
        $uri = "https://prices.azure.com/api/retail/prices?api-version=2023-01-01-preview&currencyCode=$CurrencyCode&`$filter=$encodedFilter"
        $resp = Invoke-AzRetailPriceQuery -Uri $uri

        $items = @()
        if ($resp -and $resp.Items) {
            $items = @($resp.Items)
        }

        $item = $items |
            Where-Object {
                $_.type -eq 'Consumption' -and
                $_.meterName -notmatch 'Spot|Low Priority'
            } |
            Sort-Object effectiveStartDate -Descending |
            Select-Object -First 1

        $price = $null
        if ($item) {
            $price = $item.retailPrice
        }

        $script:PriceCache[$CacheKey] = $price
        return $price
    }
    catch {
        Write-Warning "Pricing API error for key '$CacheKey': $($_.Exception.Message)"
        $script:PriceCache[$CacheKey] = $null
        return $null
    }
}

function Get-VMHourlyPrice {
    param(
        [Parameter(Mandatory = $true)][string]$SkuName,
        [Parameter(Mandatory = $true)][string]$Region
    )

    $cacheKey = "VM|$SkuName|$Region|$CurrencyCode"
    $filter = "serviceName eq 'Virtual Machines' and armSkuName eq '$SkuName' and armRegionName eq '$Region' and priceType eq 'Consumption'"
    Get-AzRetailPrice -Filter $filter -CacheKey $cacheKey
}

function Get-DiskTierLabel {
    param(
        [Parameter(Mandatory = $true)][string]$DiskSkuName,
        [Parameter(Mandatory = $true)][int]$DiskSizeGB
    )

    $family = switch -Regex ($DiskSkuName) {
        'Premium'     { 'P'; break }
        'StandardSSD' { 'E'; break }
        'Standard_LRS' { 'S'; break }
        'Ultra'       { 'U'; break }
        default       { 'S' }
    }

    if ($family -eq 'U') {
        return 'Ultra'
    }

    $tierNumber = switch ($DiskSizeGB) {
        { $_ -le 4 }     { 1; break }
        { $_ -le 8 }     { 2; break }
        { $_ -le 16 }    { 3; break }
        { $_ -le 32 }    { 4; break }
        { $_ -le 64 }    { 6; break }
        { $_ -le 128 }   { 10; break }
        { $_ -le 256 }   { 15; break }
        { $_ -le 512 }   { 20; break }
        { $_ -le 1024 }  { 30; break }
        { $_ -le 2048 }  { 40; break }
        { $_ -le 4096 }  { 50; break }
        { $_ -le 8192 }  { 60; break }
        { $_ -le 16384 } { 70; break }
        default          { 80 }
    }

    return "$family$tierNumber"
}

function Get-DiskProductName {
    param(
        [Parameter(Mandatory = $true)][string]$DiskSkuName
    )

    switch -Regex ($DiskSkuName) {
        'Premium'     { 'Premium SSD Managed Disks'; break }
        'StandardSSD' { 'Standard SSD Managed Disks'; break }
        'Standard_LRS' { 'Standard HDD Managed Disks'; break }
        'Ultra'       { 'Ultra Disks'; break }
        default       { 'Standard HDD Managed Disks' }
    }
}

function Get-DiskMonthlyPrice {
    param(
        [Parameter(Mandatory = $true)][string]$DiskSkuName,
        [Parameter(Mandatory = $true)][string]$Region,
        [Parameter(Mandatory = $true)][int]$DiskSizeGB
    )

    $productName = Get-DiskProductName -DiskSkuName $DiskSkuName
    $tierLabel = Get-DiskTierLabel -DiskSkuName $DiskSkuName -DiskSizeGB $DiskSizeGB

    if ($tierLabel -eq 'Ultra') {
        return $null
    }

    $skuName = "$tierLabel LRS"
    if ($DiskSkuName -match 'ZRS') {
        $skuName = "$tierLabel ZRS"
    }

    $cacheKey = "DISK|$productName|$skuName|$Region|$CurrencyCode"
    $filter = "serviceFamily eq 'Storage' and armRegionName eq '$Region' and skuName eq '$skuName' and productName eq '$productName' and priceType eq 'Consumption'"

    Get-AzRetailPrice -Filter $filter -CacheKey $cacheKey
}

function Get-VMPowerState {
    param(
        [Parameter(Mandatory = $true)]$VmObject
    )

    $powerState = $null

    if ($VmObject.PSObject.Properties.Name -contains 'Statuses' -and $VmObject.Statuses) {
        $statusList = @($VmObject.Statuses)
        $powerState = ($statusList | Where-Object { $_.Code -match 'PowerState' } | Select-Object -First 1).DisplayStatus
    }

    if (-not $powerState) {
        try {
            $vmWithStatus = Get-AzVM -ResourceGroupName $VmObject.ResourceGroupName -Name $VmObject.Name -Status -ErrorAction Stop
            if ($vmWithStatus.PSObject.Properties.Name -contains 'Statuses' -and $vmWithStatus.Statuses) {
                $statusList = @($vmWithStatus.Statuses)
                $powerState = ($statusList | Where-Object { $_.Code -match 'PowerState' } | Select-Object -First 1).DisplayStatus
            }
        }
        catch {
            $powerState = $null
        }
    }

    if (-not $powerState) {
        $powerState = 'Unknown'
    }

    return $powerState
}

function Get-SafeProperty {
    param(
        $Object,
        [string]$PropertyName
    )

    if ($null -ne $Object -and $Object.PSObject.Properties.Name -contains $PropertyName) {
        return $Object.$PropertyName
    }

    return $null
}

# ----------------------------
# Collect inventory
# ----------------------------

$allVMs = New-Object System.Collections.Generic.List[object]
$allDisks = New-Object System.Collections.Generic.List[object]

$subscriptions = Get-AzSubscription | Where-Object {
    $_.State -eq 'Enabled' -and $_.Id -notin $ExcludeSubscriptions
}

Write-Host ""
Write-Host "[INFO] Found $($subscriptions.Count) enabled subscription(s)." -ForegroundColor Cyan
Write-Host ""

$subIndex = 0

foreach ($sub in $subscriptions) {
    $subIndex++
    Write-Host "[$subIndex/$($subscriptions.Count)] Processing subscription: $($sub.Name) ($($sub.Id))" -ForegroundColor Yellow

    try {
        Set-AzContext -SubscriptionId $sub.Id -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Warning "Could not switch to subscription $($sub.Id): $($_.Exception.Message)"
        continue
    }

    $vms = @(Get-AzVM -ErrorAction SilentlyContinue)

    if (-not $vms -or $vms.Count -eq 0) {
        Write-Host "  No VMs found." -ForegroundColor DarkGray
        continue
    }

    Write-Host "  Found $($vms.Count) VM(s). Collecting details..." -ForegroundColor Green

    foreach ($vm in $vms) {
        $powerState = Get-VMPowerState -VmObject $vm

        $nicNames = New-Object System.Collections.Generic.List[string]
        $privateIPs = New-Object System.Collections.Generic.List[string]
        $publicIPs = New-Object System.Collections.Generic.List[string]
        $vnetNames = New-Object System.Collections.Generic.List[string]
        $subnetNames = New-Object System.Collections.Generic.List[string]
        $nsgNames = New-Object System.Collections.Generic.List[string]
        $acceleratedNetworking = $false

        foreach ($nicRef in @($vm.NetworkProfile.NetworkInterfaces)) {
            $nicIdParts = $nicRef.Id -split '/'
            $nicName = $nicIdParts[-1]
            $nicRg = $nicIdParts[4]

            try {
                $nic = Get-AzNetworkInterface -Name $nicName -ResourceGroupName $nicRg -ErrorAction Stop
                $nicNames.Add($nic.Name)

                if ($nic.EnableAcceleratedNetworking) {
                    $acceleratedNetworking = $true
                }

                foreach ($ipConfig in @($nic.IpConfigurations)) {
                    if ($ipConfig.PrivateIpAddress) {
                        $privateIPs.Add($ipConfig.PrivateIpAddress)
                    }

                    if ($ipConfig.Subnet -and $ipConfig.Subnet.Id) {
                        $subnetIdParts = $ipConfig.Subnet.Id -split '/'
                        if ($subnetIdParts.Count -ge 11) {
                            $vnetNames.Add($subnetIdParts[8])
                            $subnetNames.Add($subnetIdParts[10])
                        }
                    }

                    if ($ipConfig.PublicIpAddress -and $ipConfig.PublicIpAddress.Id) {
                        $pipIdParts = $ipConfig.PublicIpAddress.Id -split '/'
                        $pipName = $pipIdParts[-1]
                        $pipRg = $pipIdParts[4]

                        try {
                            $pip = Get-AzPublicIpAddress -Name $pipName -ResourceGroupName $pipRg -ErrorAction Stop
                            if ($pip.IpAddress) {
                                $publicIPs.Add($pip.IpAddress)
                            }
                        }
                        catch {
                            $publicIPs.Add('N/A')
                        }
                    }
                }

                if ($nic.NetworkSecurityGroup -and $nic.NetworkSecurityGroup.Id) {
                    $nsgNames.Add(($nic.NetworkSecurityGroup.Id -split '/')[-1])
                }
            }
            catch {
                Write-Warning "Could not retrieve NIC '$nicName': $($_.Exception.Message)"
            }
        }

        $region = $vm.Location
        $vmSku = $vm.HardwareProfile.VmSize
        $vmHourlyPrice = Get-VMHourlyPrice -SkuName $vmSku -Region $region
        $vmMonthlyPrice = $null

        if ($null -ne $vmHourlyPrice) {
            $vmMonthlyPrice = [math]::Round(([double]$vmHourlyPrice * 730), 4)
        }

        $tagString = ''
        if ($vm.Tags) {
            $tagString = ($vm.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
        }

        $osDisk = $vm.StorageProfile.OsDisk
        $osDiskResource = $null
        $osDiskSizeGB = $null
        $osDiskSku = $null
        $osDiskPrice = $null
        $osDiskManagedId = $null
        $osDiskState = $null
        $osDiskEncType = $null

        if ($osDisk.ManagedDisk -and $osDisk.ManagedDisk.Id) {
            $osDiskManagedId = $osDisk.ManagedDisk.Id
            $osDiskRg = ($osDisk.ManagedDisk.Id -split '/')[4]
            $osDiskName = ($osDisk.ManagedDisk.Id -split '/')[-1]

            try {
                $osDiskResource = Get-AzDisk -ResourceGroupName $osDiskRg -DiskName $osDiskName -ErrorAction Stop
            }
            catch {
                $osDiskResource = $null
            }
        }

        if ($osDiskResource) {
            $osDiskSizeGB = $osDiskResource.DiskSizeGB
            $osDiskSku = $osDiskResource.Sku.Name
            $osDiskState = Get-SafeProperty -Object $osDiskResource -PropertyName 'DiskState'

            if ($osDiskResource.Encryption) {
                $osDiskEncType = Get-SafeProperty -Object $osDiskResource.Encryption -PropertyName 'Type'
            }
        }
        else {
            $osDiskSizeGB = Get-SafeProperty -Object $osDisk -PropertyName 'DiskSizeGB'
            $osDiskSku = 'Unknown'
        }

        if ($osDiskSizeGB -and $osDiskSku -and $osDiskSku -ne 'Unknown') {
            $osDiskPrice = Get-DiskMonthlyPrice -DiskSkuName $osDiskSku -Region $region -DiskSizeGB ([int]$osDiskSizeGB)
        }

        $allDisks.Add([PSCustomObject]@{
            SubscriptionName      = $sub.Name
            SubscriptionId        = $sub.Id
            ResourceGroup         = $vm.ResourceGroupName
            VMName                = $vm.Name
            DiskName              = $osDisk.Name
            DiskType              = 'OS Disk'
            DiskSKU               = $osDiskSku
            DiskSizeGB            = $osDiskSizeGB
            Caching               = $osDisk.Caching
            StorageAccountType    = $osDiskSku
            ManagedDiskId         = $osDiskManagedId
            LUN                   = $null
            DiskState             = $osDiskState
            EncryptionType        = $osDiskEncType
            Region                = $region
            EstMonthlyPrice_USD   = $osDiskPrice
        })

        foreach ($dataDisk in @($vm.StorageProfile.DataDisks)) {
            $ddResource = $null
            $ddSizeGB = $null
            $ddSku = $null
            $ddPrice = $null
            $ddManagedId = $null
            $ddState = $null
            $ddEncType = $null

            if ($dataDisk.ManagedDisk -and $dataDisk.ManagedDisk.Id) {
                $ddManagedId = $dataDisk.ManagedDisk.Id
                $ddRg = ($dataDisk.ManagedDisk.Id -split '/')[4]
                $ddName = ($dataDisk.ManagedDisk.Id -split '/')[-1]

                try {
                    $ddResource = Get-AzDisk -ResourceGroupName $ddRg -DiskName $ddName -ErrorAction Stop
                }
                catch {
                    $ddResource = $null
                }
            }

            if ($ddResource) {
                $ddSizeGB = $ddResource.DiskSizeGB
                $ddSku = $ddResource.Sku.Name
                $ddState = Get-SafeProperty -Object $ddResource -PropertyName 'DiskState'

                if ($ddResource.Encryption) {
                    $ddEncType = Get-SafeProperty -Object $ddResource.Encryption -PropertyName 'Type'
                }
            }
            else {
                $ddSizeGB = Get-SafeProperty -Object $dataDisk -PropertyName 'DiskSizeGB'
                $ddSku = 'Unknown'
            }

            if ($ddSizeGB -and $ddSku -and $ddSku -ne 'Unknown') {
                $ddPrice = Get-DiskMonthlyPrice -DiskSkuName $ddSku -Region $region -DiskSizeGB ([int]$ddSizeGB)
            }

            $allDisks.Add([PSCustomObject]@{
                SubscriptionName      = $sub.Name
                SubscriptionId        = $sub.Id
                ResourceGroup         = $vm.ResourceGroupName
                VMName                = $vm.Name
                DiskName              = $dataDisk.Name
                DiskType              = 'Data Disk'
                DiskSKU               = $ddSku
                DiskSizeGB            = $ddSizeGB
                Caching               = $dataDisk.Caching
                StorageAccountType    = $ddSku
                ManagedDiskId         = $ddManagedId
                LUN                   = $dataDisk.Lun
                DiskState             = $ddState
                EncryptionType        = $ddEncType
                Region                = $region
                EstMonthlyPrice_USD   = $ddPrice
            })
        }

        $vmDiskRows = @($allDisks | Where-Object { $_.VMName -eq $vm.Name -and $_.SubscriptionId -eq $sub.Id })
        $diskMeasure = $vmDiskRows | Measure-Object -Property EstMonthlyPrice_USD -Sum
        $diskSum = 0

        if ($diskMeasure -and $null -ne $diskMeasure.Sum) {
            $diskSum = [double]$diskMeasure.Sum
        }

        $imageRef = $vm.StorageProfile.ImageReference
        $osProfile = $vm.OSProfile
        $diagProfile = $vm.DiagnosticsProfile

        $availabilitySet = $null
        if ($vm.AvailabilitySetReference -and $vm.AvailabilitySetReference.Id) {
            $availabilitySet = ($vm.AvailabilitySetReference.Id -split '/')[-1]
        }

        $vmssName = $null
        if ($vm.VirtualMachineScaleSet -and $vm.VirtualMachineScaleSet.Id) {
            $vmssName = ($vm.VirtualMachineScaleSet.Id -split '/')[-1]
        }

        $ppgName = $null
        if ($vm.ProximityPlacementGroup -and $vm.ProximityPlacementGroup.Id) {
            $ppgName = ($vm.ProximityPlacementGroup.Id -split '/')[-1]
        }

        $bootDiagEnabled = $false
        if ($diagProfile -and $diagProfile.BootDiagnostics) {
            if ($null -ne $diagProfile.BootDiagnostics.Enabled) {
                $bootDiagEnabled = [bool]$diagProfile.BootDiagnostics.Enabled
            }
        }

        $allVMs.Add([PSCustomObject]@{
            SubscriptionName         = $sub.Name
            SubscriptionId           = $sub.Id
            ResourceGroup            = $vm.ResourceGroupName
            VMName                   = $vm.Name
            PowerState               = $powerState
            Location                 = $region
            VMSize                   = $vmSku
            OSType                   = $osDisk.OsType
            OSPublisher              = $(if ($imageRef) { $imageRef.Publisher } else { $null })
            OSOffer                  = $(if ($imageRef) { $imageRef.Offer } else { $null })
            OSSKU                    = $(if ($imageRef) { $imageRef.Sku } else { $null })
            OSVersion                = $(if ($imageRef) { $imageRef.Version } else { $null })
            ComputerName             = $(if ($osProfile) { $osProfile.ComputerName } else { $null })
            AdminUsername            = $(if ($osProfile) { $osProfile.AdminUsername } else { $null })
            AvailabilitySet          = $availabilitySet
            VirtualMachineScaleSet   = $vmssName
            AvailabilityZone         = ($vm.Zones -join ',')
            ProximityPlacementGroup  = $ppgName
            LicenseType              = $vm.LicenseType
            BootDiagnosticsEnabled   = $bootDiagEnabled
            AcceleratedNetworking    = $acceleratedNetworking
            NICNames                 = ($nicNames -join ', ')
            PrivateIPAddresses       = ($privateIPs -join ', ')
            PublicIPAddresses        = ($publicIPs -join ', ')
            VNetNames                = (($vnetNames | Select-Object -Unique) -join ', ')
            SubnetNames              = (($subnetNames | Select-Object -Unique) -join ', ')
            NSGNames                 = (($nsgNames | Select-Object -Unique) -join ', ')
            OSDiskName               = $osDisk.Name
            OSDiskSizeGB             = $osDiskSizeGB
            OSDiskSKU                = $osDiskSku
            DataDiskCount            = @($vm.StorageProfile.DataDisks).Count
            Tags                     = $tagString
            VMHourlyPrice_USD        = $vmHourlyPrice
            VMMonthlyPrice_USD       = $vmMonthlyPrice
            EstTotalDiskCost_USD     = [math]::Round($diskSum, 4)
        })

        Write-Host ("    ✓ {0} [{1}] | VM: ${2}/hr" -f $vm.Name, $vmSku, $(if ($null -ne $vmHourlyPrice) { $vmHourlyPrice } else { 'N/A' })) -ForegroundColor DarkGreen
    }
}

# ----------------------------
# Build cost summary
# ----------------------------

$costSummary = foreach ($group in ($allVMs | Group-Object SubscriptionName)) {
    $subName = $group.Name
    $grpVMs = @($group.Group)
    $subDisks = @($allDisks | Where-Object { $_.SubscriptionName -eq $subName })

    $vmCostMeasure = $grpVMs | Measure-Object -Property VMMonthlyPrice_USD -Sum
    $diskCostMeasure = $subDisks | Measure-Object -Property EstMonthlyPrice_USD -Sum
    $dataDiskCountMeasure = $grpVMs | Measure-Object -Property DataDiskCount -Sum

    [PSCustomObject]@{
        SubscriptionName           = $subName
        TotalVMs                   = $grpVMs.Count
        RunningVMs                 = @($grpVMs | Where-Object { $_.PowerState -match 'running' }).Count
        DeallocatedVMs             = @($grpVMs | Where-Object { $_.PowerState -match 'deallocated' }).Count
        TotalDataDisks             = $(if ($null -ne $dataDiskCountMeasure.Sum) { [int]$dataDiskCountMeasure.Sum } else { 0 })
        TotalVMCost_Monthly_USD    = [math]::Round($(if ($null -ne $vmCostMeasure.Sum) { [double]$vmCostMeasure.Sum } else { 0 }), 2)
        TotalDiskCost_Monthly_USD  = [math]::Round($(if ($null -ne $diskCostMeasure.Sum) { [double]$diskCostMeasure.Sum } else { 0 }), 2)
    }
}

# ----------------------------
# Export to Excel
# ----------------------------

Write-Host ""
Write-Host "[INFO] Exporting to Excel: $OutputPath" -ForegroundColor Cyan

$excelPkg = $allVMs | Export-Excel `
    -Path $OutputPath `
    -WorksheetName 'VM Inventory' `
    -AutoSize `
    -AutoFilter `
    -FreezeTopRow `
    -BoldTopRow `
    -TableName 'VMInventory' `
    -TableStyle 'Medium9' `
    -PassThru

$ws1 = $excelPkg.Workbook.Worksheets['VM Inventory']

if ($allVMs.Count -gt 0) {
    $headers = @($allVMs[0].PSObject.Properties.Name)
    $pwrIndex = [array]::IndexOf($headers, 'PowerState')

    if ($pwrIndex -ge 0) {
        $excelCol = $pwrIndex + 1
        $letters = ""
        $dividend = $excelCol
        while ($dividend -gt 0) {
            $modulo = ($dividend - 1) % 26
            $letters = [char](65 + $modulo) + $letters
            $dividend = [math]::Floor(($dividend - $modulo) / 26)
        }

        $lastRow = $allVMs.Count + 1
        $range = "$letters" + "2:" + "$letters" + "$lastRow"

        Add-ConditionalFormatting -Worksheet $ws1 -Address $range -RuleType ContainsText -ConditionValue 'running' -BackgroundColor ([System.Drawing.Color]::LightGreen)
        Add-ConditionalFormatting -Worksheet $ws1 -Address $range -RuleType ContainsText -ConditionValue 'deallocated' -BackgroundColor ([System.Drawing.Color]::LightCoral)
        Add-ConditionalFormatting -Worksheet $ws1 -Address $range -RuleType ContainsText -ConditionValue 'stopped' -BackgroundColor ([System.Drawing.Color]::LightYellow)
    }
}

$excelPkg = $allDisks | Export-Excel `
    -ExcelPackage $excelPkg `
    -WorksheetName 'Disk Inventory' `
    -AutoSize `
    -AutoFilter `
    -FreezeTopRow `
    -BoldTopRow `
    -TableName 'DiskInventory' `
    -TableStyle 'Medium6' `
    -PassThru

$excelPkg = $costSummary | Export-Excel `
    -ExcelPackage $excelPkg `
    -WorksheetName 'Cost Summary' `
    -AutoSize `
    -FreezeTopRow `
    -BoldTopRow `
    -TableName 'CostSummary' `
    -TableStyle 'Medium2' `
    -PassThru

Close-ExcelPackage $excelPkg

Write-Host ""
Write-Host "[DONE] Report saved to: $OutputPath" -ForegroundColor Green
Write-Host "  - VM Inventory   : $($allVMs.Count) VM(s)" -ForegroundColor White
Write-Host "  - Disk Inventory : $($allDisks.Count) disk(s)" -ForegroundColor White
Write-Host "  - Cost Summary   : $(@($costSummary).Count) subscription row(s)" -ForegroundColor White