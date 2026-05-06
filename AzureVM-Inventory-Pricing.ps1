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
    [string]$OutputPath   = "$env:USERPROFILE\Desktop\AzureVM_Inventory_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx",
    [string]$CurrencyCode = "USD",
    [string[]]$ExcludeSubscriptions = @()   # Subscription IDs to skip
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

#region ── Helper: Azure Retail Price Lookup ──────────────────────────────────

# Cache to avoid duplicate API calls for the same SKU+Region
$script:PriceCache = @{}

function Get-AzRetailPrice {
    param(
        [string]$Filter,
        [string]$CacheKey
    )
    if ($script:PriceCache.ContainsKey($CacheKey)) {
        return $script:PriceCache[$CacheKey]
    }
    try {
        $uri    = "https://prices.azure.com/api/retail/prices?api-version=2023-01-01-preview"
        $params = @{ '$filter' = $Filter; currencyCode = $CurrencyCode }
        $resp   = Invoke-RestMethod -Uri $uri -Method Get -Body $params -ErrorAction Stop
        $item   = $resp.Items |
                  Where-Object { $_.type -eq 'Consumption' -and $_.meterName -notmatch 'Spot|Low Priority' } |
                  Sort-Object effectiveStartDate -Descending |
                  Select-Object -First 1
        $price  = if ($item) { $item.retailPrice } else { $null }
        $script:PriceCache[$CacheKey] = $price
        return $price
    } catch {
        Write-Warning "Pricing API error for key '$CacheKey': $_"
        $script:PriceCache[$CacheKey] = $null
        return $null
    }
}

function Get-VMHourlyPrice {
    param([string]$SkuName, [string]$Region)
    $cacheKey = "VM_${SkuName}_${Region}"
    $filter   = "serviceName eq 'Virtual Machines' and armSkuName eq '$SkuName' " +
                "and armRegionName eq '$Region' and priceType eq 'Consumption'"
    return Get-AzRetailPrice -Filter $filter -CacheKey $cacheKey
}

function Get-DiskMonthlyPrice {
    param([string]$DiskSkuName, [string]$Region, [int]$DiskSizeGB)

    # Map disk SKU tier + size to Azure pricing product name
    $tier = switch -Regex ($DiskSkuName) {
        'Premium_LRS'     { 'Premium SSD' }
        'StandardSSD_LRS' { 'Standard SSD' }
        'Standard_LRS'    { 'Standard HDD' }
        'UltraSSD_LRS'    { 'Ultra Disk'   }
        'Premium_ZRS'     { 'Premium SSD'  }
        'StandardSSD_ZRS' { 'Standard SSD' }
        default           { 'Standard HDD' }
    }

    # Determine disk size label (Azure pricing tiers)
    $sizeLabel = switch ($DiskSizeGB) {
        { $_ -le 4   } { 'P1'; break }
        { $_ -le 8   } { 'P2'; break }
        { $_ -le 16  } { 'P3'; break }
        { $_ -le 32  } { 'P4'; break }
        { $_ -le 64  } { 'P6'; break }
        { $_ -le 128 } { 'P10'; break }
        { $_ -le 256 } { 'P15'; break }
        { $_ -le 512 } { 'P20'; break }
        { $_ -le 1024 }{ 'P30'; break }
        { $_ -le 2048 }{ 'P40'; break }
        { $_ -le 4096 }{ 'P50'; break }
        { $_ -le 8192 }{ 'P60'; break }
        { $_ -le 16384}{ 'P70'; break }
        default         { 'P80' }
    }

    # For Standard HDD/SSD, use E/S labels
    if ($tier -eq 'Standard SSD') { $sizeLabel = $sizeLabel -replace '^P', 'E' }
    if ($tier -eq 'Standard HDD') { $sizeLabel = $sizeLabel -replace '^P', 'S' }

    $cacheKey = "DISK_${tier}_${sizeLabel}_${Region}"
    $filter   = "serviceFamily eq 'Storage' and armRegionName eq '$Region' " +
                "and skuName eq '$sizeLabel' and productName eq 'Premium SSD Managed Disks'"

    # Build correct productName per tier
    $productName = switch ($tier) {
        'Premium SSD'  { 'Premium SSD Managed Disks'  }
        'Standard SSD' { 'Standard SSD Managed Disks' }
        'Standard HDD' { 'Standard HDD Managed Disks' }
        'Ultra Disk'   { 'Ultra Disks'                }
        default        { 'Premium SSD Managed Disks'  }
    }

    $filter = "serviceFamily eq 'Storage' and armRegionName eq '$Region' " +
              "and skuName eq '$sizeLabel LRS' and productName eq '$productName'"

    return Get-AzRetailPrice -Filter $filter -CacheKey $cacheKey
}

#endregion

#region ── Data Collection ────────────────────────────────────────────────────

$allVMs   = [System.Collections.Generic.List[PSCustomObject]]::new()
$allDisks = [System.Collections.Generic.List[PSCustomObject]]::new()

# Get all accessible subscriptions
$subscriptions = Get-AzSubscription | Where-Object {
    $_.State -eq 'Enabled' -and $_.Id -notin $ExcludeSubscriptions
}

Write-Host "`n[INFO] Found $($subscriptions.Count) enabled subscription(s).`n" -ForegroundColor Cyan

$subCount = 0
foreach ($sub in $subscriptions) {
    $subCount++
    Write-Host "[$subCount/$($subscriptions.Count)] Processing subscription: $($sub.Name) ($($sub.Id))" -ForegroundColor Yellow

    try {
        Set-AzContext -SubscriptionId $sub.Id -ErrorAction Stop | Out-Null
    } catch {
        Write-Warning "  Could not switch to subscription $($sub.Id): $_"
        continue
    }

    # Get all VMs in subscription (expand instanceView for power state)
    $vms = Get-AzVM -Status -ErrorAction SilentlyContinue
    if (-not $vms) {
        Write-Host "  No VMs found." -ForegroundColor DarkGray
        continue
    }

    Write-Host "  Found $($vms.Count) VM(s). Collecting details..." -ForegroundColor Green

    foreach ($vm in $vms) {

        # ── Power State ───────────────────────────────────────────────────────
        $powerState = ($vm.Statuses |
                       Where-Object { $_.Code -match 'PowerState' } |
                       Select-Object -First 1).DisplayStatus

        # ── Networking ────────────────────────────────────────────────────────
        $nicDetails       = [System.Collections.Generic.List[string]]::new()
        $privateIPs       = [System.Collections.Generic.List[string]]::new()
        $publicIPs        = [System.Collections.Generic.List[string]]::new()
        $vnetNames        = [System.Collections.Generic.List[string]]::new()
        $subnetNames      = [System.Collections.Generic.List[string]]::new()
        $nsgNames         = [System.Collections.Generic.List[string]]::new()
        $acceleratedNetw  = $false

        foreach ($nicRef in $vm.NetworkProfile.NetworkInterfaces) {
            $nicName = ($nicRef.Id -split '/')[-1]
            $nicRg   = ($nicRef.Id -split '/')[4]
            try {
                $nic = Get-AzNetworkInterface -Name $nicName -ResourceGroupName $nicRg -ErrorAction Stop
                $nicDetails.Add($nicName)
                $acceleratedNetw = $nic.EnableAcceleratedNetworking

                foreach ($ipConfig in $nic.IpConfigurations) {
                    $privateIPs.Add($ipConfig.PrivateIpAddress)
                    $subnetNames.Add(($ipConfig.Subnet.Id -split '/')[-1])
                    $vnetNames.Add(($ipConfig.Subnet.Id -split '/')[-3])

                    if ($ipConfig.PublicIpAddress) {
                        $pipName = ($ipConfig.PublicIpAddress.Id -split '/')[-1]
                        $pipRg   = ($ipConfig.PublicIpAddress.Id -split '/')[4]
                        try {
                            $pip = Get-AzPublicIpAddress -Name $pipName -ResourceGroupName $pipRg -ErrorAction Stop
                            $publicIPs.Add($pip.IpAddress)
                        } catch { $publicIPs.Add("N/A") }
                    }
                }
                if ($nic.NetworkSecurityGroup) {
                    $nsgNames.Add(($nic.NetworkSecurityGroup.Id -split '/')[-1])
                }
            } catch {
                Write-Warning "    Could not retrieve NIC '$nicName': $_"
            }
        }

        # ── VM Pricing ────────────────────────────────────────────────────────
        $region          = $vm.Location
        $vmSku           = $vm.HardwareProfile.VmSize
        $vmHourlyPrice   = Get-VMHourlyPrice -SkuName $vmSku -Region $region
        $vmMonthlyPrice  = if ($vmHourlyPrice) { [math]::Round($vmHourlyPrice * 730, 4) } else { $null }

        # ── Tags ──────────────────────────────────────────────────────────────
        $tagString = if ($vm.Tags) {
            ($vm.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
        } else { "" }

        # ── OS Disk ───────────────────────────────────────────────────────────
        $osDisk = $vm.StorageProfile.OsDisk
        $osDiskResource = $null
        if ($osDisk.ManagedDisk) {
            $osDiskRg   = ($osDisk.ManagedDisk.Id -split '/')[4]
            $osDiskName = ($osDisk.ManagedDisk.Id -split '/')[-1]
            try {
                $osDiskResource = Get-AzDisk -ResourceGroupName $osDiskRg -DiskName $osDiskName -ErrorAction Stop
            } catch {}
        }

        $osDiskSizeGB  = if ($osDiskResource) { $osDiskResource.DiskSizeGB } else { $osDisk.DiskSizeGB }
        $osDiskSku     = if ($osDiskResource) { $osDiskResource.Sku.Name } else { "Unknown" }
        $osDiskPrice   = if ($osDiskSizeGB -and $osDiskSku) {
                            Get-DiskMonthlyPrice -DiskSkuName $osDiskSku -Region $region -DiskSizeGB $osDiskSizeGB
                         } else { $null }

        # ── Add OS Disk row ───────────────────────────────────────────────────
        $allDisks.Add([PSCustomObject]@{
            SubscriptionName     = $sub.Name
            SubscriptionId       = $sub.Id
            ResourceGroup        = $vm.ResourceGroupName
            VMName               = $vm.Name
            DiskName             = $osDisk.Name
            DiskType             = "OS Disk"
            DiskSKU              = $osDiskSku
            DiskSizeGB           = $osDiskSizeGB
            Caching              = $osDisk.Caching
            StorageAccountType   = $osDiskSku
            ManagedDiskId        = $osDisk.ManagedDisk?.Id
            DiskState            = $osDiskResource?.DiskState
            EncryptionType       = $osDiskResource?.Encryption?.Type
            Region               = $region
            'EstMonthlyPrice_USD' = $osDiskPrice
        })

        # ── Data Disks ────────────────────────────────────────────────────────
        foreach ($dataDisk in $vm.StorageProfile.DataDisks) {
            $ddResource = $null
            if ($dataDisk.ManagedDisk) {
                $ddRg   = ($dataDisk.ManagedDisk.Id -split '/')[4]
                $ddName = ($dataDisk.ManagedDisk.Id -split '/')[-1]
                try {
                    $ddResource = Get-AzDisk -ResourceGroupName $ddRg -DiskName $ddName -ErrorAction Stop
                } catch {}
            }

            $ddSizeGB = if ($ddResource) { $ddResource.DiskSizeGB } else { $dataDisk.DiskSizeGB }
            $ddSku    = if ($ddResource) { $ddResource.Sku.Name } else { "Unknown" }
            $ddPrice  = if ($ddSizeGB -and $ddSku) {
                            Get-DiskMonthlyPrice -DiskSkuName $ddSku -Region $region -DiskSizeGB $ddSizeGB
                        } else { $null }

            $allDisks.Add([PSCustomObject]@{
                SubscriptionName     = $sub.Name
                SubscriptionId       = $sub.Id
                ResourceGroup        = $vm.ResourceGroupName
                VMName               = $vm.Name
                DiskName             = $dataDisk.Name
                DiskType             = "Data Disk"
                DiskSKU              = $ddSku
                DiskSizeGB           = $ddSizeGB
                Caching              = $dataDisk.Caching
                StorageAccountType   = $ddSku
                ManagedDiskId        = $dataDisk.ManagedDisk?.Id
                LUN                  = $dataDisk.Lun
                DiskState            = $ddResource?.DiskState
                EncryptionType       = $ddResource?.Encryption?.Type
                Region               = $region
                'EstMonthlyPrice_USD' = $ddPrice
            })
        }

        # ── VM Row ────────────────────────────────────────────────────────────
        $allVMs.Add([PSCustomObject]@{
            SubscriptionName         = $sub.Name
            SubscriptionId           = $sub.Id
            ResourceGroup            = $vm.ResourceGroupName
            VMName                   = $vm.Name
            PowerState               = $powerState
            Location                 = $region
            VMSize                   = $vmSku
            OSType                   = $vm.StorageProfile.OsDisk.OsType
            OSPublisher              = $vm.StorageProfile.ImageReference?.Publisher
            OSOffer                  = $vm.StorageProfile.ImageReference?.Offer
            OSSKU                    = $vm.StorageProfile.ImageReference?.Sku
            OSVersion                = $vm.StorageProfile.ImageReference?.Version
            ComputerName             = $vm.OSProfile?.ComputerName
            AdminUsername            = $vm.OSProfile?.AdminUsername
            AvailabilitySet          = ($vm.AvailabilitySetReference?.Id -split '/')[-1]
            VirtualMachineScaleSet   = ($vm.VirtualMachineScaleSet?.Id -split '/')[-1]
            AvailabilityZone         = ($vm.Zones -join ',')
            ProximityPlacementGroup  = ($vm.ProximityPlacementGroup?.Id -split '/')[-1]
            LicenseType              = $vm.LicenseType
            BootDiagnosticsEnabled   = $vm.DiagnosticsProfile?.BootDiagnostics?.Enabled
            AcceleratedNetworking    = $acceleratedNetw
            NICNames                 = ($nicDetails -join ', ')
            PrivateIPAddresses       = ($privateIPs -join ', ')
            PublicIPAddresses        = ($publicIPs -join ', ')
            VNetNames                = ($vnetNames -join ', ')
            SubnetNames              = ($subnetNames -join ', ')
            NSGNames                 = ($nsgNames -join ', ')
            OSDiskName               = $osDisk.Name
            OSDiskSizeGB             = $osDiskSizeGB
            OSDiskSKU                = $osDiskSku
            DataDiskCount            = $vm.StorageProfile.DataDisks.Count
            Tags                     = $tagString
            'VMHourlyPrice_USD'      = $vmHourlyPrice
            'VMMonthlyPrice_USD'     = $vmMonthlyPrice
            'EstTotalDiskCost_USD'   = ($allDisks | Where-Object { $_.VMName -eq $vm.Name -and $_.SubscriptionId -eq $sub.Id } |
                                         Measure-Object -Property EstMonthlyPrice_USD -Sum).Sum
        })

        Write-Host "    ✓ $($vm.Name) [$vmSku] | VM: `$$vmHourlyPrice/hr" -ForegroundColor DarkGreen
    }
}

#endregion

#region ── Excel Export ───────────────────────────────────────────────────────

Write-Host "`n[INFO] Exporting to Excel: $OutputPath" -ForegroundColor Cyan

# ── Sheet 1: VM Inventory ─────────────────────────────────────────────────────
$vmExcelParams = @{
    Path          = $OutputPath
    WorksheetName = "VM Inventory"
    AutoSize      = $true
    AutoFilter    = $true
    FreezeTopRow  = $true
    BoldTopRow    = $true
    TableName     = "VMInventory"
    TableStyle    = "Medium9"
    PassThru      = $true
}
$excelPkg = $allVMs | Export-Excel @vmExcelParams

# Conditional formatting: color power state column
$ws1   = $excelPkg.Workbook.Worksheets["VM Inventory"]
$pwrCol = ($allVMs | Get-Member -MemberType NoteProperty).Name.IndexOf('PowerState') + 1
if ($pwrCol -gt 0 -and $allVMs.Count -gt 0) {
    $lastRow = $allVMs.Count + 1
    Add-ConditionalFormatting -Worksheet $ws1 -Range "$([char](64+$pwrCol))2:$([char](64+$pwrCol))$lastRow" `
        -RuleType ContainsText -ConditionValue "running" -BackgroundColor ([System.Drawing.Color]::LightGreen)
    Add-ConditionalFormatting -Worksheet $ws1 -Range "$([char](64+$pwrCol))2:$([char](64+$pwrCol))$lastRow" `
        -RuleType ContainsText -ConditionValue "deallocated" -BackgroundColor ([System.Drawing.Color]::LightCoral)
    Add-ConditionalFormatting -Worksheet $ws1 -Range "$([char](64+$pwrCol))2:$([char](64+$pwrCol))$lastRow" `
        -RuleType ContainsText -ConditionValue "stopped" -BackgroundColor ([System.Drawing.Color]::LightYellow)
}

# ── Sheet 2: Disk Inventory ────────────────────────────────────────────────────
$diskExcelParams = @{
    ExcelPackage  = $excelPkg
    WorksheetName = "Disk Inventory"
    AutoSize      = $true
    AutoFilter    = $true
    FreezeTopRow  = $true
    BoldTopRow    = $true
    TableName     = "DiskInventory"
    TableStyle    = "Medium6"
    PassThru      = $true
}
$excelPkg = $allDisks | Export-Excel @diskExcelParams

# ── Sheet 3: Cost Summary by Subscription ─────────────────────────────────────
$costSummary = $allVMs |
    Group-Object SubscriptionName |
    ForEach-Object {
        $grpVMs   = $_.Group
        $subDisks = $allDisks | Where-Object { $_.SubscriptionName -eq $_.Name }
        [PSCustomObject]@{
            SubscriptionName     = $_.Name
            TotalVMs             = $grpVMs.Count
            RunningVMs           = ($grpVMs | Where-Object { $_.PowerState -match 'running' }).Count
            DeallocatedVMs       = ($grpVMs | Where-Object { $_.PowerState -match 'deallocated' }).Count
            TotalDataDisks       = ($grpVMs | Measure-Object DataDiskCount -Sum).Sum
            'TotalVMCost_Monthly_USD'   = [math]::Round(($grpVMs | Measure-Object VMMonthlyPrice_USD -Sum).Sum, 2)
            'TotalDiskCost_Monthly_USD' = [math]::Round(($allDisks |
                Where-Object { $_.SubscriptionName -eq $_.Name } |
                Measure-Object EstMonthlyPrice_USD -Sum).Sum, 2)
        }
    }

$summaryExcelParams = @{
    ExcelPackage  = $excelPkg
    WorksheetName = "Cost Summary"
    AutoSize      = $true
    FreezeTopRow  = $true
    BoldTopRow    = $true
    TableName     = "CostSummary"
    TableStyle    = "Medium2"
    PassThru      = $true
}
$excelPkg = $costSummary | Export-Excel @summaryExcelParams

# Save and close
Close-ExcelPackage $excelPkg

Write-Host "`n[DONE] Report saved to: $OutputPath" -ForegroundColor Green
Write-Host "  - VM Inventory  : $($allVMs.Count) VMs"   -ForegroundColor White
Write-Host "  - Disk Inventory: $($allDisks.Count) Disks" -ForegroundColor White

#endregion