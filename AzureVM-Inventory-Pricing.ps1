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
    [string[]]$ExcludeSubscriptions = @(),
    [string]$TenantId = ""   # Optional: pass your tenant ID to avoid cross-tenant errors
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

$script:PriceCache = @{}

# ─────────────────────────────────────────────
# Helper: Build full pricing URL manually
# ─────────────────────────────────────────────
function Build-PriceUrl {
    param(
        [Parameter(Mandatory = $true)][string]$Filter
    )
    $encodedFilter = [System.Uri]::EscapeDataString($Filter)
    return "https://prices.azure.com/api/retail/prices?api-version=2023-01-01-preview&currencyCode=$CurrencyCode&`$filter=$encodedFilter"
}

# ─────────────────────────────────────────────
# Helper: REST call with retry on rate limit
# ─────────────────────────────────────────────
function Invoke-RetailPriceRequest {
    param(
        [Parameter(Mandatory = $true)][string]$Uri
    )

    $maxRetries = 4
    $attempt    = 0

    do {
        try {
            return Invoke-RestMethod -Uri $Uri -Method Get -UseBasicParsing -ErrorAction Stop
        }
        catch {
            $attempt++
            $msg = $_.Exception.Message
            if ($msg -match 'Too many requests|429' -and $attempt -lt $maxRetries) {
                $wait = 5 * $attempt
                Write-Warning "  Rate limited by Pricing API. Retrying in $wait second(s) (attempt $attempt/$maxRetries)..."
                Start-Sleep -Seconds $wait
            }
            else {
                throw
            }
        }
    } while ($attempt -lt $maxRetries)
}

# ─────────────────────────────────────────────
# Helper: Generic retail price lookup w/ cache
# ─────────────────────────────────────────────
function Get-AzRetailPrice {
    param(
        [Parameter(Mandatory = $true)][string]$Filter,
        [Parameter(Mandatory = $true)][string]$CacheKey
    )

    if ($script:PriceCache.ContainsKey($CacheKey)) {
        return $script:PriceCache[$CacheKey]
    }

    try {
        $uri       = Build-PriceUrl -Filter $Filter
        $allItems  = New-Object System.Collections.Generic.List[object]

        do {
            $resp = Invoke-RetailPriceRequest -Uri $uri
            if ($resp -and $resp.Items) {
                foreach ($i in @($resp.Items)) { $allItems.Add($i) }
            }
            $uri = $null
            if ($resp -and $resp.NextPageLink -and $resp.NextPageLink -ne '') {
                $uri = $resp.NextPageLink
            }
        } while ($uri)

        $item = @($allItems) |
            Where-Object {
                $_.type -eq 'Consumption' -and
                $_.meterName -notmatch 'Spot|Low Priority'
            } |
            Sort-Object effectiveStartDate -Descending |
            Select-Object -First 1

        $price = $null
        if ($item) { $price = $item.retailPrice }

        $script:PriceCache[$CacheKey] = $price
        return $price
    }
    catch {
        Write-Warning "  Pricing API error [$CacheKey]: $($_.Exception.Message)"
        $script:PriceCache[$CacheKey] = $null
        return $null
    }
}

# ─────────────────────────────────────────────
# VM pricing
# ─────────────────────────────────────────────
function Get-VMHourlyPrice {
    param(
        [Parameter(Mandatory = $true)][string]$SkuName,
        [Parameter(Mandatory = $true)][string]$Region
    )
    $cacheKey = "VM|$SkuName|$Region|$CurrencyCode"
    $filter   = "serviceName eq 'Virtual Machines' and armSkuName eq '$SkuName' and armRegionName eq '$Region' and priceType eq 'Consumption'"
    return Get-AzRetailPrice -Filter $filter -CacheKey $cacheKey
}

# ─────────────────────────────────────────────
# Disk tier label mapping
# ─────────────────────────────────────────────
function Get-DiskTierLabel {
    param(
        [Parameter(Mandatory = $true)][string]$DiskSkuName,
        [Parameter(Mandatory = $true)][int]$DiskSizeGB
    )

    $family = 'S'
    if ($DiskSkuName -match '^Premium')      { $family = 'P' }
    elseif ($DiskSkuName -match '^StandardSSD') { $family = 'E' }
    elseif ($DiskSkuName -match '^Standard_LRS') { $family = 'S' }
    elseif ($DiskSkuName -match '^Ultra')    { return 'Ultra' }

    $num = switch ($DiskSizeGB) {
        { $_ -le 4 }      { 1; break }
        { $_ -le 8 }      { 2; break }
        { $_ -le 16 }     { 3; break }
        { $_ -le 32 }     { 4; break }
        { $_ -le 64 }     { 6; break }
        { $_ -le 128 }    { 10; break }
        { $_ -le 256 }    { 15; break }
        { $_ -le 512 }    { 20; break }
        { $_ -le 1024 }   { 30; break }
        { $_ -le 2048 }   { 40; break }
        { $_ -le 4096 }   { 50; break }
        { $_ -le 8192 }   { 60; break }
        { $_ -le 16384 }  { 70; break }
        default           { 80 }
    }

    return "$family$num"
}

function Get-DiskProductName {
    param([Parameter(Mandatory = $true)][string]$DiskSkuName)

    if ($DiskSkuName -match '^Premium')         { return 'Premium SSD Managed Disks' }
    elseif ($DiskSkuName -match '^StandardSSD') { return 'Standard SSD Managed Disks' }
    elseif ($DiskSkuName -match '^Ultra')       { return 'Ultra Disks' }
    else                                        { return 'Standard HDD Managed Disks' }
}

function Get-DiskMonthlyPrice {
    param(
        [Parameter(Mandatory = $true)][string]$DiskSkuName,
        [Parameter(Mandatory = $true)][string]$Region,
        [Parameter(Mandatory = $true)][int]$DiskSizeGB
    )

    $tierLabel   = Get-DiskTierLabel -DiskSkuName $DiskSkuName -DiskSizeGB $DiskSizeGB
    if ($tierLabel -eq 'Ultra') { return $null }

    $productName = Get-DiskProductName -DiskSkuName $DiskSkuName
    $redundancy  = if ($DiskSkuName -match 'ZRS') { 'ZRS' } else { 'LRS' }
    $skuName     = "$tierLabel $redundancy"

    $cacheKey = "DISK|$productName|$skuName|$Region|$CurrencyCode"
    $filter   = "serviceFamily eq 'Storage' and armRegionName eq '$Region' and skuName eq '$skuName' and productName eq '$productName' and priceType eq 'Consumption'"

    return Get-AzRetailPrice -Filter $filter -CacheKey $cacheKey
}

# ─────────────────────────────────────────────
# Power state with safe fallback
# ─────────────────────────────────────────────
function Get-VMPowerState {
    param([Parameter(Mandatory = $true)]$VmObject)

    # Some Az versions expose .PowerState directly
    if ($VmObject.PSObject.Properties.Name -contains 'PowerState' -and $VmObject.PowerState) {
        return $VmObject.PowerState
    }

    # Try inline Statuses
    if ($VmObject.PSObject.Properties.Name -contains 'Statuses' -and $VmObject.Statuses) {
        $ps = (@($VmObject.Statuses) | Where-Object { $_.Code -match 'PowerState' } | Select-Object -First 1).DisplayStatus
        if ($ps) { return $ps }
    }

    # Explicit -Status call fallback
    try {
        $vmStatus = Get-AzVM -ResourceGroupName $VmObject.ResourceGroupName -Name $VmObject.Name -Status -ErrorAction Stop
        if ($vmStatus.PSObject.Properties.Name -contains 'PowerState' -and $vmStatus.PowerState) {
            return $vmStatus.PowerState
        }
        if ($vmStatus.PSObject.Properties.Name -contains 'Statuses' -and $vmStatus.Statuses) {
            $ps = (@($vmStatus.Statuses) | Where-Object { $_.Code -match 'PowerState' } | Select-Object -First 1).DisplayStatus
            if ($ps) { return $ps }
        }
    }
    catch {}

    return 'Unknown'
}

# ─────────────────────────────────────────────
# Safe property getter
# ─────────────────────────────────────────────
function Get-SafeProperty {
    param($Object, [string]$PropertyName)
    if ($null -ne $Object -and $Object.PSObject.Properties.Name -contains $PropertyName) {
        return $Object.$PropertyName
    }
    return $null
}

# ─────────────────────────────────────────────
# Excel column letter helper (handles AA, AB...)
# ─────────────────────────────────────────────
function Get-ExcelColumnName {
    param([Parameter(Mandatory = $true)][int]$ColumnNumber)
    $letters  = ""
    $dividend = $ColumnNumber
    while ($dividend -gt 0) {
        $mod      = ($dividend - 1) % 26
        $letters  = [char](65 + $mod) + $letters
        $dividend = [math]::Floor(($dividend - $mod) / 26)
    }
    return $letters
}

# ─────────────────────────────────────────────
# Resolve tenant ID from current context
# ─────────────────────────────────────────────
if (-not $TenantId) {
    $currentContext = Get-AzContext -ErrorAction SilentlyContinue
    if ($currentContext -and $currentContext.Tenant) {
        $TenantId = $currentContext.Tenant.Id
        Write-Host "[INFO] Using tenant ID from current context: $TenantId" -ForegroundColor Cyan
    }
}

if (-not $TenantId) {
    Write-Error "Could not determine TenantId. Please run: Connect-AzAccount -TenantId '<your-tenant-id>' or pass -TenantId parameter."
    exit 1
}

# ─────────────────────────────────────────────
# Get subscriptions — filter to current tenant only
# ─────────────────────────────────────────────
$allSubs = Get-AzSubscription -TenantId $TenantId -ErrorAction SilentlyContinue

if (-not $allSubs) {
    # Fallback: get all and filter by tenant
    $allSubs = Get-AzSubscription -ErrorAction SilentlyContinue |
               Where-Object { $_.TenantId -eq $TenantId }
}

$subscriptions = @($allSubs | Where-Object {
    $_.State -eq 'Enabled' -and $_.Id -notin $ExcludeSubscriptions
})

Write-Host ""
Write-Host "[INFO] Found $($subscriptions.Count) enabled subscription(s) in tenant $TenantId." -ForegroundColor Cyan
Write-Host ""

$allVMs   = New-Object System.Collections.Generic.List[object]
$allDisks = New-Object System.Collections.Generic.List[object]
$subIndex = 0

foreach ($sub in $subscriptions) {
    $subIndex++
    Write-Host "[$subIndex/$($subscriptions.Count)] Processing: $($sub.Name) ($($sub.Id))" -ForegroundColor Yellow

    try {
        Set-AzContext -SubscriptionId $sub.Id -TenantId $TenantId -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Warning "  Skipping subscription — cannot switch context: $($_.Exception.Message)"
        continue
    }

    $vms = @(Get-AzVM -ErrorAction SilentlyContinue)

    if (-not $vms -or $vms.Count -eq 0) {
        Write-Host "  No VMs found." -ForegroundColor DarkGray
        continue
    }

    Write-Host "  Found $($vms.Count) VM(s)..." -ForegroundColor Green

    foreach ($vm in $vms) {

        $powerState = Get-VMPowerState -VmObject $vm

        $nicNames    = New-Object System.Collections.Generic.List[string]
        $privateIPs  = New-Object System.Collections.Generic.List[string]
        $publicIPs   = New-Object System.Collections.Generic.List[string]
        $vnetNames   = New-Object System.Collections.Generic.List[string]
        $subnetNames = New-Object System.Collections.Generic.List[string]
        $nsgNames    = New-Object System.Collections.Generic.List[string]
        $accelNet    = $false

        foreach ($nicRef in @($vm.NetworkProfile.NetworkInterfaces)) {
            if (-not $nicRef.Id) { continue }

            $nicParts = $nicRef.Id -split '/'
            $nicName  = $nicParts[-1]
            $nicRg    = $nicParts[4]

            try {
                $nic = Get-AzNetworkInterface -Name $nicName -ResourceGroupName $nicRg -ErrorAction Stop
                $nicNames.Add($nic.Name)
                if ($nic.EnableAcceleratedNetworking) { $accelNet = $true }

                foreach ($ipCfg in @($nic.IpConfigurations)) {
                    if ($ipCfg.PrivateIpAddress) { $privateIPs.Add($ipCfg.PrivateIpAddress) }

                    if ($ipCfg.Subnet -and $ipCfg.Subnet.Id) {
                        $snetParts = $ipCfg.Subnet.Id -split '/'
                        if ($snetParts.Count -ge 11) {
                            $vnetNames.Add($snetParts[8])
                            $subnetNames.Add($snetParts[10])
                        }
                    }

                    if ($ipCfg.PublicIpAddress -and $ipCfg.PublicIpAddress.Id) {
                        $pipParts = $ipCfg.PublicIpAddress.Id -split '/'
                        try {
                            $pip = Get-AzPublicIpAddress -Name $pipParts[-1] -ResourceGroupName $pipParts[4] -ErrorAction Stop
                            $publicIPs.Add($(if ($pip.IpAddress) { $pip.IpAddress } else { 'Not Assigned' }))
                        }
                        catch { $publicIPs.Add('N/A') }
                    }
                }

                if ($nic.NetworkSecurityGroup -and $nic.NetworkSecurityGroup.Id) {
                    $nsgNames.Add(($nic.NetworkSecurityGroup.Id -split '/')[-1])
                }
            }
            catch {
                Write-Warning "  NIC '$nicName': $($_.Exception.Message)"
            }
        }

        $region          = $vm.Location
        $vmSku           = $vm.HardwareProfile.VmSize
        $vmHourlyPrice   = Get-VMHourlyPrice -SkuName $vmSku -Region $region
        $vmMonthlyPrice  = $null
        if ($null -ne $vmHourlyPrice) {
            $vmMonthlyPrice = [math]::Round([double]$vmHourlyPrice * 730, 4)
        }

        $tagString = ''
        if ($vm.Tags) {
            $tagString = ($vm.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
        }

        # ── OS Disk ──────────────────────────────────────────────
        $osDisk        = $vm.StorageProfile.OsDisk
        $osDiskRes     = $null
        $osDiskSizeGB  = $null
        $osDiskSku     = 'Unknown'
        $osDiskPrice   = $null
        $osDiskMgdId   = $null
        $osDiskState   = $null
        $osDiskEncType = $null

        if ($osDisk.ManagedDisk -and $osDisk.ManagedDisk.Id) {
            $osDiskMgdId  = $osDisk.ManagedDisk.Id
            $osDiskRg     = ($osDisk.ManagedDisk.Id -split '/')[4]
            $osDiskResName = ($osDisk.ManagedDisk.Id -split '/')[-1]
            try {
                $osDiskRes = Get-AzDisk -ResourceGroupName $osDiskRg -DiskName $osDiskResName -ErrorAction Stop
            }
            catch { $osDiskRes = $null }
        }

        if ($osDiskRes) {
            $osDiskSizeGB  = $osDiskRes.DiskSizeGB
            $osDiskSku     = $osDiskRes.Sku.Name
            $osDiskState   = Get-SafeProperty -Object $osDiskRes -PropertyName 'DiskState'
            $osDiskEncType = Get-SafeProperty -Object $osDiskRes.Encryption -PropertyName 'Type'
        }
        else {
            $osDiskSizeGB = Get-SafeProperty -Object $osDisk -PropertyName 'DiskSizeGB'
        }

        if ($osDiskSizeGB -and $osDiskSku -ne 'Unknown') {
            $osDiskPrice = Get-DiskMonthlyPrice -DiskSkuName $osDiskSku -Region $region -DiskSizeGB ([int]$osDiskSizeGB)
        }

        $allDisks.Add([PSCustomObject]@{
            SubscriptionName    = $sub.Name
            SubscriptionId      = $sub.Id
            ResourceGroup       = $vm.ResourceGroupName
            VMName              = $vm.Name
            DiskName            = $osDisk.Name
            DiskType            = 'OS Disk'
            DiskSKU             = $osDiskSku
            DiskSizeGB          = $osDiskSizeGB
            Caching             = $osDisk.Caching
            StorageAccountType  = $osDiskSku
            ManagedDiskId       = $osDiskMgdId
            LUN                 = $null
            DiskState           = $osDiskState
            EncryptionType      = $osDiskEncType
            Region              = $region
            EstMonthlyPrice_USD = $osDiskPrice
        })

        # ── Data Disks ────────────────────────────────────────────
        foreach ($dd in @($vm.StorageProfile.DataDisks)) {
            $ddRes     = $null
            $ddSizeGB  = $null
            $ddSku     = 'Unknown'
            $ddPrice   = $null
            $ddMgdId   = $null
            $ddState   = $null
            $ddEncType = $null

            if ($dd.ManagedDisk -and $dd.ManagedDisk.Id) {
                $ddMgdId = $dd.ManagedDisk.Id
                $ddRg    = ($dd.ManagedDisk.Id -split '/')[4]
                $ddN     = ($dd.ManagedDisk.Id -split '/')[-1]
                try {
                    $ddRes = Get-AzDisk -ResourceGroupName $ddRg -DiskName $ddN -ErrorAction Stop
                }
                catch { $ddRes = $null }
            }

            if ($ddRes) {
                $ddSizeGB  = $ddRes.DiskSizeGB
                $ddSku     = $ddRes.Sku.Name
                $ddState   = Get-SafeProperty -Object $ddRes -PropertyName 'DiskState'
                $ddEncType = Get-SafeProperty -Object $ddRes.Encryption -PropertyName 'Type'
            }
            else {
                $ddSizeGB = Get-SafeProperty -Object $dd -PropertyName 'DiskSizeGB'
            }

            if ($ddSizeGB -and $ddSku -ne 'Unknown') {
                $ddPrice = Get-DiskMonthlyPrice -DiskSkuName $ddSku -Region $region -DiskSizeGB ([int]$ddSizeGB)
            }

            $allDisks.Add([PSCustomObject]@{
                SubscriptionName    = $sub.Name
                SubscriptionId      = $sub.Id
                ResourceGroup       = $vm.ResourceGroupName
                VMName              = $vm.Name
                DiskName            = $dd.Name
                DiskType            = 'Data Disk'
                DiskSKU             = $ddSku
                DiskSizeGB          = $ddSizeGB
                Caching             = $dd.Caching
                StorageAccountType  = $ddSku
                ManagedDiskId       = $ddMgdId
                LUN                 = $dd.Lun
                DiskState           = $ddState
                EncryptionType      = $ddEncType
                Region              = $region
                EstMonthlyPrice_USD = $ddPrice
            })
        }

        # ── VM row ─────────────────────────────────────────────────
        $vmDiskRows = @($allDisks | Where-Object { $_.VMName -eq $vm.Name -and $_.SubscriptionId -eq $sub.Id })
        $diskMeasure = $vmDiskRows | Measure-Object -Property EstMonthlyPrice_USD -Sum
        $diskSum = if ($diskMeasure -and $null -ne $diskMeasure.Sum) { [double]$diskMeasure.Sum } else { 0 }

        $imgRef      = $vm.StorageProfile.ImageReference
        $osProf      = $vm.OSProfile
        $diagProf    = $vm.DiagnosticsProfile
        $bootDiag    = $false

        if ($diagProf -and $diagProf.BootDiagnostics -and $null -ne $diagProf.BootDiagnostics.Enabled) {
            $bootDiag = [bool]$diagProf.BootDiagnostics.Enabled
        }

        $availSet = $null
        if ($vm.AvailabilitySetReference -and $vm.AvailabilitySetReference.Id) {
            $availSet = ($vm.AvailabilitySetReference.Id -split '/')[-1]
        }
        $vmssRef = $null
        if ($vm.VirtualMachineScaleSet -and $vm.VirtualMachineScaleSet.Id) {
            $vmssRef = ($vm.VirtualMachineScaleSet.Id -split '/')[-1]
        }
        $ppgRef = $null
        if ($vm.ProximityPlacementGroup -and $vm.ProximityPlacementGroup.Id) {
            $ppgRef = ($vm.ProximityPlacementGroup.Id -split '/')[-1]
        }

        $allVMs.Add([PSCustomObject]@{
            SubscriptionName        = $sub.Name
            SubscriptionId          = $sub.Id
            ResourceGroup           = $vm.ResourceGroupName
            VMName                  = $vm.Name
            PowerState              = $powerState
            Location                = $region
            VMSize                  = $vmSku
            OSType                  = $osDisk.OsType
            OSPublisher             = $(if ($imgRef) { $imgRef.Publisher } else { $null })
            OSOffer                 = $(if ($imgRef) { $imgRef.Offer } else { $null })
            OSSKU                   = $(if ($imgRef) { $imgRef.Sku } else { $null })
            OSVersion               = $(if ($imgRef) { $imgRef.Version } else { $null })
            ComputerName            = $(if ($osProf) { $osProf.ComputerName } else { $null })
            AdminUsername           = $(if ($osProf) { $osProf.AdminUsername } else { $null })
            AvailabilitySet         = $availSet
            VirtualMachineScaleSet  = $vmssRef
            AvailabilityZone        = ($vm.Zones -join ',')
            ProximityPlacementGroup = $ppgRef
            LicenseType             = $vm.LicenseType
            BootDiagnosticsEnabled  = $bootDiag
            AcceleratedNetworking   = $accelNet
            NICNames                = ($nicNames -join ', ')
            PrivateIPAddresses      = ($privateIPs -join ', ')
            PublicIPAddresses       = ($publicIPs -join ', ')
            VNetNames               = (($vnetNames | Select-Object -Unique) -join ', ')
            SubnetNames             = (($subnetNames | Select-Object -Unique) -join ', ')
            NSGNames                = (($nsgNames | Select-Object -Unique) -join ', ')
            OSDiskName              = $osDisk.Name
            OSDiskSizeGB            = $osDiskSizeGB
            OSDiskSKU               = $osDiskSku
            DataDiskCount           = @($vm.StorageProfile.DataDisks).Count
            Tags                    = $tagString
            VMHourlyPrice_USD       = $vmHourlyPrice
            VMMonthlyPrice_USD      = $vmMonthlyPrice
            EstTotalDiskCost_USD    = [math]::Round($diskSum, 4)
        })

        $priceDisplay = if ($null -ne $vmHourlyPrice) { $vmHourlyPrice } else { 'N/A' }
        Write-Host "    ✓ $($vm.Name) [$vmSku] | VM: `$$priceDisplay/hr" -ForegroundColor DarkGreen
    }
}

# ─────────────────────────────────────────────
# Cost Summary per subscription
# ─────────────────────────────────────────────
$costSummary = foreach ($grp in ($allVMs | Group-Object SubscriptionName)) {
    $sName     = $grp.Name
    $grpVMs    = @($grp.Group)
    $grpDisks  = @($allDisks | Where-Object { $_.SubscriptionName -eq $sName })

    $vmCost    = $grpVMs  | Measure-Object -Property VMMonthlyPrice_USD   -Sum
    $dskCost   = $grpDisks | Measure-Object -Property EstMonthlyPrice_USD -Sum
    $ddCount   = $grpVMs  | Measure-Object -Property DataDiskCount         -Sum

    [PSCustomObject]@{
        SubscriptionName          = $sName
        TotalVMs                  = $grpVMs.Count
        RunningVMs                = @($grpVMs | Where-Object { $_.PowerState -match 'running' }).Count
        DeallocatedVMs            = @($grpVMs | Where-Object { $_.PowerState -match 'deallocated' }).Count
        TotalDataDisks            = $(if ($null -ne $ddCount.Sum)  { [int]$ddCount.Sum }    else { 0 })
        TotalVMCost_Monthly_USD   = [math]::Round($(if ($null -ne $vmCost.Sum)  { [double]$vmCost.Sum }  else { 0 }), 2)
        TotalDiskCost_Monthly_USD = [math]::Round($(if ($null -ne $dskCost.Sum) { [double]$dskCost.Sum } else { 0 }), 2)
    }
}

# ─────────────────────────────────────────────
# Excel Export
# ─────────────────────────────────────────────
Write-Host ""
Write-Host "[INFO] Exporting to Excel: $OutputPath" -ForegroundColor Cyan

$excelPkg = $allVMs | Export-Excel `
    -Path          $OutputPath `
    -WorksheetName 'VM Inventory' `
    -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
    -TableName     'VMInventory' `
    -TableStyle    'Medium9' `
    -PassThru

$ws1 = $excelPkg.Workbook.Worksheets['VM Inventory']

if ($allVMs.Count -gt 0) {
    $headers  = @($allVMs[0].PSObject.Properties.Name)
    $pwrIdx   = [array]::IndexOf($headers, 'PowerState')

    if ($pwrIdx -ge 0) {
        $col     = Get-ExcelColumnName -ColumnNumber ($pwrIdx + 1)
        $lastRow = $allVMs.Count + 1
        $range   = "${col}2:${col}${lastRow}"

        Add-ConditionalFormatting -Worksheet $ws1 -Address $range -RuleType ContainsText -ConditionValue 'running'     -BackgroundColor ([System.Drawing.Color]::LightGreen)
        Add-ConditionalFormatting -Worksheet $ws1 -Address $range -RuleType ContainsText -ConditionValue 'deallocated' -BackgroundColor ([System.Drawing.Color]::LightCoral)
        Add-ConditionalFormatting -Worksheet $ws1 -Address $range -RuleType ContainsText -ConditionValue 'stopped'     -BackgroundColor ([System.Drawing.Color]::LightYellow)
    }
}

$excelPkg = $allDisks | Export-Excel `
    -ExcelPackage  $excelPkg `
    -WorksheetName 'Disk Inventory' `
    -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
    -TableName     'DiskInventory' `
    -TableStyle    'Medium6' `
    -PassThru

$excelPkg = $costSummary | Export-Excel `
    -ExcelPackage  $excelPkg `
    -WorksheetName 'Cost Summary' `
    -AutoSize -FreezeTopRow -BoldTopRow `
    -TableName     'CostSummary' `
    -TableStyle    'Medium2' `
    -PassThru

Close-ExcelPackage $excelPkg

Write-Host ""
Write-Host "[DONE] Report saved to: $OutputPath" -ForegroundColor Green
Write-Host "  - VM Inventory   : $($allVMs.Count) VM(s)"        -ForegroundColor White
Write-Host "  - Disk Inventory : $($allDisks.Count) disk(s)"    -ForegroundColor White
Write-Host "  - Cost Summary   : $(@($costSummary).Count) row(s)" -ForegroundColor White