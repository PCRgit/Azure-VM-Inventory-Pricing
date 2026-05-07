<#
.SYNOPSIS
    Azure VM Inventory + Pricing Export to Excel

.DESCRIPTION
    Scans all enabled subscriptions in your tenant, collects full VM settings
    (compute, networking, OS, tags, availability), all attached managed disks,
    performs live retail price lookups for VM SKUs and each disk, and exports
    everything to a richly formatted multi-sheet Excel workbook.

.NOTES
    Required modules:
      Install-Module Az -Scope CurrentUser -Force
      Install-Module ImportExcel -Scope CurrentUser -Force

    Authenticate first:
      Connect-AzAccount -TenantId "<your-tenant-id>"

.EXAMPLE
    .\AzureVM-Inventory-Pricing.ps1

.EXAMPLE
    .\AzureVM-Inventory-Pricing.ps1 -OutputPath "C:\Reports\VMReport.xlsx" -ExcludeSubscriptions @("sub-id-to-skip") -CurrencyCode "USD"
#>

#Requires -Modules Az.Accounts, Az.Compute, Az.Network, ImportExcel

[CmdletBinding()]
param(
    [string]   $OutputPath           = "$env:USERPROFILE\Desktop\AzureVM_Inventory_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx",
    [string]   $CurrencyCode         = "USD",
    [string[]] $ExcludeSubscriptions = @(),
    [string]   $TenantId             = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

$script:PriceCache = @{}

#region ── Pricing helpers ─────────────────────────────────────────────────────

function Build-PriceUrl {
    param([string]$Filter)

    $encodedFilter = [System.Uri]::EscapeDataString($Filter)
    return "https://prices.azure.com/api/retail/prices?api-version=2023-01-01-preview&currencyCode=$CurrencyCode&meterRegion='primary'&`$filter=$encodedFilter"
}

function Invoke-RetailPriceRequest {
    param([string]$Uri)

    $maxRetries = 8
    $attempt    = 0

    do {
        try {
            return Invoke-RestMethod -Uri $Uri -Method Get -UseBasicParsing -ErrorAction Stop
        }
        catch {
            $attempt++

            $isRateLimited = $_.Exception.Message -match '429|Too many requests'
            if (-not $isRateLimited -or $attempt -ge $maxRetries) {
                throw
            }

            $retryAfter = $null
            try {
                if ($_.Exception.Response -and $_.Exception.Response.Headers['Retry-After']) {
                    $retryAfter = [int]$_.Exception.Response.Headers['Retry-After']
                }
            }
            catch {
                $retryAfter = $null
            }

            if (-not $retryAfter) {
                # Exponential backoff + small jitter
                $retryAfter = [math]::Min(60, ([math]::Pow(2, $attempt) + (Get-Random -Minimum 0 -Maximum 3)))
            }

            Write-Warning "  Rate limited. Retrying in $retryAfter sec (attempt $attempt of $maxRetries)..."
            Start-Sleep -Seconds $retryAfter
        }
    } while ($attempt -lt $maxRetries)
}

function Get-AzRetailPriceItems {
    param(
        [Parameter(Mandatory)][string]$Filter,
        [Parameter(Mandatory)][string]$CacheKey
    )

    if ($script:PriceCache.ContainsKey($CacheKey)) {
        return $script:PriceCache[$CacheKey]
    }

    try {
        $uri      = Build-PriceUrl -Filter $Filter
        $allItems = [System.Collections.Generic.List[object]]::new()

        do {
            $resp = Invoke-RetailPriceRequest -Uri $uri
            if ($resp -and $resp.Items) {
                foreach ($item in $resp.Items) {
                    [void]$allItems.Add($item)
                }
            }
            $uri = if ($resp -and $resp.NextPageLink -and $resp.NextPageLink -ne '') { $resp.NextPageLink } else { $null }
        } while ($uri)

        $items = @($allItems)
        $script:PriceCache[$CacheKey] = $items
        return $items
    }
    catch {
        Write-Warning "  Pricing API error [$CacheKey]: $($_.Exception.Message)"
        $script:PriceCache[$CacheKey] = @()
        return @()
    }
}

function Get-LatestConsumptionItem {
    param(
        [object[]]$Items,
        [scriptblock]$Predicate
    )

    $filtered = @(
        $Items | Where-Object {
            $primaryOk = $true
            if ($_.PSObject.Properties.Name -contains 'isPrimaryMeterRegion') {
                $primaryOk = ($_.isPrimaryMeterRegion -eq $true)
            }

            $_.type -eq 'Consumption' -and
            $primaryOk -and
            (& $Predicate $_)
        }
    )

    if (-not $filtered -or $filtered.Count -eq 0) {
        return $null
    }

    return $filtered |
        Sort-Object effectiveStartDate -Descending |
        Select-Object -First 1
}

function Convert-UnitPriceToMonthlyEstimate {
    param(
        [double]$Price,
        [string]$UnitOfMeasure
    )

    if ([string]::IsNullOrWhiteSpace($UnitOfMeasure)) {
        return [math]::Round($Price, 6)
    }

    $u = $UnitOfMeasure.ToLowerInvariant()

    # Hourly meters -> monthly estimate using 730 hours
    if ($u -match 'hour') {
        return [math]::Round(($Price * 730), 6)
    }

    # Month / GiB-month / IOPS-month / MBps-month meters are already monthly
    return [math]::Round($Price, 6)
}

function Select-PreferredVmMeter {
    param(
        [object[]]$Items,
        [string]$OsType,
        [string]$LicenseType
    )

    if (-not $Items -or $Items.Count -eq 0) {
        return $null
    }

    $preferBaseRate = $false
    if ($OsType -eq 'Windows' -and $LicenseType -in @('Windows_Server', 'Windows_Client')) {
        $preferBaseRate = $true
    }

    $commonValid = {
        param($i)
        $i.type -eq 'Consumption' -and
        $i.unitOfMeasure -match 'Hour' -and
        $i.meterName -notmatch 'Spot|Low Priority' -and
        $i.skuName   -notmatch 'Spot|Low Priority' -and
        $i.productName -notmatch 'Dedicated Host'
    }

    if ($preferBaseRate) {
        $match = @(
            $Items | Where-Object {
                (& $commonValid $_) -and
                $_.productName -notmatch 'Windows'
            }
        ) | Sort-Object effectiveStartDate -Descending | Select-Object -First 1

        if ($match) {
            return @{
                Item         = $match
                BillingModel = "Licensed Base Compute ($LicenseType)"
            }
        }
    }

    if ($OsType -eq 'Windows') {
        $match = @(
            $Items | Where-Object {
                (& $commonValid $_) -and
                $_.productName -match 'Windows'
            }
        ) | Sort-Object effectiveStartDate -Descending | Select-Object -First 1

        if ($match) {
            return @{
                Item         = $match
                BillingModel = 'Windows PAYG'
            }
        }
    }
    else {
        $match = @(
            $Items | Where-Object {
                (& $commonValid $_) -and
                $_.productName -notmatch 'Windows'
            }
        ) | Sort-Object effectiveStartDate -Descending | Select-Object -First 1

        if ($match) {
            return @{
                Item         = $match
                BillingModel = 'Linux / Non-Windows'
            }
        }
    }

    # Final generic fallback: cheapest valid hourly compute meter
    $fallback = @(
        $Items | Where-Object { & $commonValid $_ }
    ) | Sort-Object retailPrice, effectiveStartDate -Descending | Select-Object -First 1

    if ($fallback) {
        return @{
            Item         = $fallback
            BillingModel = 'Generic fallback compute'
        }
    }

    return $null
}

function Get-VMRetailPriceInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SkuName,
        [Parameter(Mandatory)][string]$Region,
        [string]$OsType,
        [string]$LicenseType
    )

    # Pass 1: strict VM query
    $key1    = "VM|STRICT|$SkuName|$Region|$OsType|$LicenseType|$CurrencyCode"
    $filter1 = "serviceName eq 'Virtual Machines' and armSkuName eq '$SkuName' and armRegionName eq '$Region' and priceType eq 'Consumption'"

    $items1 = Get-AzRetailPriceItems -Filter $filter1 -CacheKey $key1
    $selection = Select-PreferredVmMeter -Items $items1 -OsType $OsType -LicenseType $LicenseType

    # Pass 2: broader compute-family fallback
    if (-not $selection) {
        $key2    = "VM|BROAD|$SkuName|$Region|$OsType|$LicenseType|$CurrencyCode"
        $filter2 = "serviceFamily eq 'Compute' and armSkuName eq '$SkuName' and armRegionName eq '$Region' and priceType eq 'Consumption'"

        $items2 = Get-AzRetailPriceItems -Filter $filter2 -CacheKey $key2
        $selection = Select-PreferredVmMeter -Items $items2 -OsType $OsType -LicenseType $LicenseType
    }

    if (-not $selection) {
        return [PSCustomObject]@{
            HourlyPrice      = $null
            MonthlyPrice     = $null
            BillingModel     = 'No valid hourly meter'
            ProductName      = $null
            MeterName        = $null
            SkuNameReturned  = $null
            UnitOfMeasure    = $null
            EffectiveDate    = $null
        }
    }

    $item = $selection.Item
    $hourlyPrice = [math]::Round([double]$item.retailPrice, 6)

    return [PSCustomObject]@{
        HourlyPrice      = $hourlyPrice
        MonthlyPrice     = [math]::Round(($hourlyPrice * 730), 2)
        BillingModel     = $selection.BillingModel
        ProductName      = $item.productName
        MeterName        = $item.meterName
        SkuNameReturned  = $item.skuName
        UnitOfMeasure    = $item.unitOfMeasure
        EffectiveDate    = $item.effectiveStartDate
    }
}

function Get-DiskTierLabel {
    param(
        [string]$DiskSkuName,
        [int]$DiskSizeGB,
        [string]$ExplicitTier = $null
    )

    # If Azure reports explicit disk tier (Pxx/Eyy/Szz), prefer it
    if ($ExplicitTier -and $ExplicitTier -match '^[PES]\d+$') {
        return $ExplicitTier
    }

    if ($DiskSkuName -match '^Ultra')     { return 'Ultra' }
    if ($DiskSkuName -match '^PremiumV2') { return 'PremiumV2' }

    if ($DiskSkuName -match '^Premium') {
        $prefix = 'P'
    }
    elseif ($DiskSkuName -match '^StandardSSD') {
        $prefix = 'E'
    }
    else {
        $prefix = 'S'
    }

    $n = switch ($DiskSizeGB) {
        { $_ -le 4     } { 1;  break }
        { $_ -le 8     } { 2;  break }
        { $_ -le 16    } { 3;  break }
        { $_ -le 32    } { 4;  break }
        { $_ -le 64    } { 6;  break }
        { $_ -le 128   } { 10; break }
        { $_ -le 256   } { 15; break }
        { $_ -le 512   } { 20; break }
        { $_ -le 1024  } { 30; break }
        { $_ -le 2048  } { 40; break }
        { $_ -le 4096  } { 50; break }
        { $_ -le 8192  } { 60; break }
        { $_ -le 16384 } { 70; break }
        default          { 80 }
    }

    return "$prefix$n"
}

function Get-PremiumV2DiskMonthlyPrice {
    param(
        [Parameter(Mandatory)][string]$Region,
        [Parameter(Mandatory)][int]$DiskSizeGB,
        [int]$ProvisionedIOPS = 0,
        [int]$ProvisionedMBps = 0
    )

    $productName = 'Azure Premium SSD v2'
    $baseKey     = "DISK|PremiumV2|$Region|$CurrencyCode"
    $filter      = "serviceName eq 'Storage' and armRegionName eq '$Region' and productName eq '$productName' and priceType eq 'Consumption'"

    $items = Get-AzRetailPriceItems -Filter $filter -CacheKey $baseKey
    if (-not $items -or $items.Count -eq 0) {
        return $null
    }

    $capacityItem = Get-LatestConsumptionItem -Items $items -Predicate {
        param($i)
        $i.meterName -eq 'Premium LRS provisioned capacity'
    }

    $iopsItem = Get-LatestConsumptionItem -Items $items -Predicate {
        param($i)
        $i.meterName -eq 'Premium LRS provisioned IOPS'
    }

    $throughputItem = Get-LatestConsumptionItem -Items $items -Predicate {
        param($i)
        $i.meterName -eq 'Premium LRS provisioned throughput (MB/s)'
    }

    if (-not $capacityItem) {
        return $null
    }

    $capacityMonthlyPerGiB = Convert-UnitPriceToMonthlyEstimate -Price ([double]$capacityItem.retailPrice) -UnitOfMeasure $capacityItem.unitOfMeasure
    $capacityCost          = $DiskSizeGB * $capacityMonthlyPerGiB

    $billableIops = [math]::Max(0, ($ProvisionedIOPS - 3000))
    $billableMbps = [math]::Max(0, ($ProvisionedMBps - 125))

    $iopsCost = 0
    if ($billableIops -gt 0 -and $iopsItem) {
        $iopsMonthly = Convert-UnitPriceToMonthlyEstimate -Price ([double]$iopsItem.retailPrice) -UnitOfMeasure $iopsItem.unitOfMeasure
        $iopsCost    = $billableIops * $iopsMonthly
    }

    $throughputCost = 0
    if ($billableMbps -gt 0 -and $throughputItem) {
        $mbpsMonthly    = Convert-UnitPriceToMonthlyEstimate -Price ([double]$throughputItem.retailPrice) -UnitOfMeasure $throughputItem.unitOfMeasure
        $throughputCost = $billableMbps * $mbpsMonthly
    }

    return [math]::Round(($capacityCost + $iopsCost + $throughputCost), 2)
}

function Get-DiskMonthlyPrice {
    param(
        [Parameter(Mandatory)][string]$DiskSkuName,
        [Parameter(Mandatory)][string]$Region,
        [Parameter(Mandatory)][int]$DiskSizeGB,
        [string]$DiskTier = $null,
        [int]$ProvisionedIOPS = 0,
        [int]$ProvisionedMBps = 0
    )

    if ($DiskSkuName -match '^PremiumV2') {
        return Get-PremiumV2DiskMonthlyPrice `
            -Region $Region `
            -DiskSizeGB $DiskSizeGB `
            -ProvisionedIOPS $ProvisionedIOPS `
            -ProvisionedMBps $ProvisionedMBps
    }

    if ($DiskSkuName -match '^Ultra') {
        return $null
    }

    $tier       = Get-DiskTierLabel -DiskSkuName $DiskSkuName -DiskSizeGB $DiskSizeGB -ExplicitTier $DiskTier
    $redundancy = if ($DiskSkuName -match 'ZRS') { 'ZRS' } else { 'LRS' }

    $productName = switch -Regex ($DiskSkuName) {
        '^Premium'      { 'Premium SSD Managed Disks'; break }
        '^StandardSSD'  { 'Standard SSD Managed Disks'; break }
        default         { 'Standard HDD Managed Disks' }
    }

    $meterName = "$tier $redundancy Disk"
    $key       = "DISK|$productName|$meterName|$Region|$CurrencyCode"
    $filter    = "serviceName eq 'Storage' and armRegionName eq '$Region' and productName eq '$productName' and meterName eq '$meterName' and priceType eq 'Consumption'"

    $items = Get-AzRetailPriceItems -Filter $filter -CacheKey $key
    if (-not $items -or $items.Count -eq 0) {
        return $null
    }

    $item = Get-LatestConsumptionItem -Items $items -Predicate {
        param($i)
        $i.meterName -eq $meterName
    }

    if (-not $item) {
        return $null
    }

    return Convert-UnitPriceToMonthlyEstimate -Price ([double]$item.retailPrice) -UnitOfMeasure $item.unitOfMeasure
}

#endregion

#region ── VM helpers ──────────────────────────────────────────────────────────

function Get-VMPowerState {
    param($VmObject)

    if ($VmObject.PSObject.Properties.Name -contains 'PowerState' -and $VmObject.PowerState) {
        return $VmObject.PowerState
    }

    if ($VmObject.PSObject.Properties.Name -contains 'Statuses' -and $VmObject.Statuses) {
        $ps = (@($VmObject.Statuses) | Where-Object { $_.Code -match 'PowerState' } | Select-Object -First 1).DisplayStatus
        if ($ps) { return $ps }
    }

    try {
        $s = Get-AzVM -ResourceGroupName $VmObject.ResourceGroupName -Name $VmObject.Name -Status -ErrorAction Stop
        if ($s.PSObject.Properties.Name -contains 'PowerState' -and $s.PowerState) {
            return $s.PowerState
        }

        if ($s.Statuses) {
            $ps = (@($s.Statuses) | Where-Object { $_.Code -match 'PowerState' } | Select-Object -First 1).DisplayStatus
            if ($ps) { return $ps }
        }
    }
    catch {}

    return 'Unknown'
}

function Get-SafeProp {
    param($Object, [string]$Prop)

    if ($null -ne $Object -and $Object.PSObject.Properties.Name -contains $Prop) {
        return $Object.$Prop
    }

    return $null
}

function Get-IntProp {
    param(
        $Object,
        [string]$Prop,
        [int]$Default = 0
    )

    $value = Get-SafeProp -Object $Object -Prop $Prop
    if ($null -eq $value -or $value -eq '') {
        return $Default
    }

    try {
        return [int]$value
    }
    catch {
        return $Default
    }
}

function Get-ExcelColumn {
    param([int]$Col)

    $s = ""
    $d = [int]$Col

    while ($d -gt 0) {
        $m = [int](($d - 1) % 26)
        $s = ([char](65 + $m)).ToString() + $s
        $d = [int][math]::Floor(($d - 1) / 26)
    }

    return $s
}

#endregion

#region ── Resolve tenant + subscriptions ──────────────────────────────────────

if (-not $TenantId) {
    $ctx = Get-AzContext -ErrorAction SilentlyContinue
    if ($ctx -and $ctx.Tenant) {
        $TenantId = $ctx.Tenant.Id
    }
}

if (-not $TenantId) {
    Write-Error "Cannot determine TenantId. Run: Connect-AzAccount -TenantId '<id>'"
    exit 1
}

Write-Host "`n[INFO] Tenant: $TenantId  |  Currency: $CurrencyCode" -ForegroundColor Cyan

$subscriptions = @(
    Get-AzSubscription -TenantId $TenantId -ErrorAction SilentlyContinue |
    Where-Object { $_.State -eq 'Enabled' -and $_.Id -notin $ExcludeSubscriptions }
)

Write-Host "[INFO] Found $($subscriptions.Count) enabled subscription(s).`n" -ForegroundColor Cyan

#endregion

#region ── Inventory collection ────────────────────────────────────────────────

$allVMs   = [System.Collections.Generic.List[object]]::new()
$allDisks = [System.Collections.Generic.List[object]]::new()
$subIndex = 0

foreach ($sub in $subscriptions) {
    $subIndex++
    Write-Host "[$subIndex/$($subscriptions.Count)] $($sub.Name) ($($sub.Id))" -ForegroundColor Yellow

    try {
        Set-AzContext -SubscriptionId $sub.Id -TenantId $TenantId -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Warning "  Skipping — cannot set context: $($_.Exception.Message)"
        continue
    }

    $vms = @(Get-AzVM -ErrorAction SilentlyContinue)
    if (-not $vms -or $vms.Count -eq 0) {
        Write-Host "  No VMs found." -ForegroundColor DarkGray
        continue
    }

    Write-Host "  Found $($vms.Count) VM(s)." -ForegroundColor Green

    foreach ($vm in $vms) {
        $powerState = Get-VMPowerState -VmObject $vm

        $nicNames    = [System.Collections.Generic.List[string]]::new()
        $privateIPs  = [System.Collections.Generic.List[string]]::new()
        $publicIPs   = [System.Collections.Generic.List[string]]::new()
        $vnetNames   = [System.Collections.Generic.List[string]]::new()
        $subnetNames = [System.Collections.Generic.List[string]]::new()
        $nsgNames    = [System.Collections.Generic.List[string]]::new()
        $accelNet    = $false

        foreach ($nicRef in @($vm.NetworkProfile.NetworkInterfaces)) {
            if (-not $nicRef.Id) { continue }

            $np   = $nicRef.Id -split '/'
            $nicN = $np[-1]
            $nicG = $np[4]

            try {
                $nic = Get-AzNetworkInterface -Name $nicN -ResourceGroupName $nicG -ErrorAction Stop
                [void]$nicNames.Add($nic.Name)

                if ($nic.EnableAcceleratedNetworking) {
                    $accelNet = $true
                }

                foreach ($ip in @($nic.IpConfigurations)) {
                    if ($ip.PrivateIpAddress) {
                        [void]$privateIPs.Add($ip.PrivateIpAddress)
                    }

                    if ($ip.Subnet -and $ip.Subnet.Id) {
                        $sp = $ip.Subnet.Id -split '/'
                        if ($sp.Count -ge 11) {
                            [void]$vnetNames.Add($sp[8])
                            [void]$subnetNames.Add($sp[10])
                        }
                    }

                    if ($ip.PublicIpAddress -and $ip.PublicIpAddress.Id) {
                        $pp = $ip.PublicIpAddress.Id -split '/'
                        try {
                            $pip = Get-AzPublicIpAddress -Name $pp[-1] -ResourceGroupName $pp[4] -ErrorAction Stop
                            [void]$publicIPs.Add($(if ($pip.IpAddress) { $pip.IpAddress } else { 'Dynamic/Unassigned' }))
                        }
                        catch {
                            [void]$publicIPs.Add('N/A')
                        }
                    }
                }

                if ($nic.NetworkSecurityGroup -and $nic.NetworkSecurityGroup.Id) {
                    [void]$nsgNames.Add(($nic.NetworkSecurityGroup.Id -split '/')[-1])
                }
            }
            catch {
                Write-Warning "  NIC '$nicN': $($_.Exception.Message)"
            }
        }

        $region = $vm.Location
        $vmSku  = $vm.HardwareProfile.VmSize
        $osDisk = $vm.StorageProfile.OsDisk

        $vmPriceInfo = Get-VMRetailPriceInfo `
            -SkuName $vmSku `
            -Region $region `
            -OsType $osDisk.OsType `
            -LicenseType $vm.LicenseType

        $vmHourlyPrice  = $vmPriceInfo.HourlyPrice
        $vmMonthlyPrice = $vmPriceInfo.MonthlyPrice

        $tagString = ''
        if ($vm.Tags) {
            $tagString = ($vm.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
        }

        # ── OS Disk ───────────────────────────────────────────────
        $osDiskRes     = $null
        $osDiskSizeGB  = $null
        $osDiskSku     = 'Unknown'
        $osDiskState   = $null
        $osDiskEncType = $null
        $osDiskMgdId   = $null
        $osDiskPrice   = $null
        $osDiskTier    = $null
        $osDiskIOPS    = 0
        $osDiskMBps    = 0

        if ($osDisk.ManagedDisk -and $osDisk.ManagedDisk.Id) {
            $osDiskMgdId = $osDisk.ManagedDisk.Id
            $dp = $osDisk.ManagedDisk.Id -split '/'
            try {
                $osDiskRes = Get-AzDisk -ResourceGroupName $dp[4] -DiskName $dp[-1] -ErrorAction Stop
            }
            catch {
                $osDiskRes = $null
            }
        }

        if ($osDiskRes) {
            $osDiskSizeGB  = $osDiskRes.DiskSizeGB
            $osDiskSku     = $osDiskRes.Sku.Name
            $osDiskState   = Get-SafeProp -Object $osDiskRes -Prop 'DiskState'
            $osDiskEncType = Get-SafeProp -Object $osDiskRes.Encryption -Prop 'Type'
            $osDiskTier    = Get-SafeProp -Object $osDiskRes -Prop 'Tier'
            $osDiskIOPS    = Get-IntProp  -Object $osDiskRes -Prop 'DiskIOPSReadWrite' -Default 0
            $osDiskMBps    = Get-IntProp  -Object $osDiskRes -Prop 'DiskMBpsReadWrite' -Default 0
        }
        else {
            $osDiskSizeGB = Get-SafeProp -Object $osDisk -Prop 'DiskSizeGB'
        }

        if ($osDiskSizeGB -and $osDiskSku -ne 'Unknown') {
            $osDiskPrice = Get-DiskMonthlyPrice `
                -DiskSkuName $osDiskSku `
                -Region $region `
                -DiskSizeGB ([int]$osDiskSizeGB) `
                -DiskTier $osDiskTier `
                -ProvisionedIOPS $osDiskIOPS `
                -ProvisionedMBps $osDiskMBps
        }

        [void]$allDisks.Add([PSCustomObject]@{
            SubscriptionName    = $sub.Name
            SubscriptionId      = $sub.Id
            ResourceGroup       = $vm.ResourceGroupName
            VMName              = $vm.Name
            DiskName            = $osDisk.Name
            DiskType            = 'OS Disk'
            DiskSKU             = $osDiskSku
            DiskTier            = $osDiskTier
            DiskSizeGB          = $osDiskSizeGB
            ProvisionedIOPS     = $osDiskIOPS
            ProvisionedMBps     = $osDiskMBps
            Caching             = $osDisk.Caching
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
            $ddState   = $null
            $ddEncType = $null
            $ddMgdId   = $null
            $ddPrice   = $null
            $ddTier    = $null
            $ddIOPS    = 0
            $ddMBps    = 0

            if ($dd.ManagedDisk -and $dd.ManagedDisk.Id) {
                $ddMgdId = $dd.ManagedDisk.Id
                $dp = $dd.ManagedDisk.Id -split '/'
                try {
                    $ddRes = Get-AzDisk -ResourceGroupName $dp[4] -DiskName $dp[-1] -ErrorAction Stop
                }
                catch {
                    $ddRes = $null
                }
            }

            if ($ddRes) {
                $ddSizeGB  = $ddRes.DiskSizeGB
                $ddSku     = $ddRes.Sku.Name
                $ddState   = Get-SafeProp -Object $ddRes -Prop 'DiskState'
                $ddEncType = Get-SafeProp -Object $ddRes.Encryption -Prop 'Type'
                $ddTier    = Get-SafeProp -Object $ddRes -Prop 'Tier'
                $ddIOPS    = Get-IntProp  -Object $ddRes -Prop 'DiskIOPSReadWrite' -Default 0
                $ddMBps    = Get-IntProp  -Object $ddRes -Prop 'DiskMBpsReadWrite' -Default 0
            }
            else {
                $ddSizeGB = Get-SafeProp -Object $dd -Prop 'DiskSizeGB'
            }

            if ($ddSizeGB -and $ddSku -ne 'Unknown') {
                $ddPrice = Get-DiskMonthlyPrice `
                    -DiskSkuName $ddSku `
                    -Region $region `
                    -DiskSizeGB ([int]$ddSizeGB) `
                    -DiskTier $ddTier `
                    -ProvisionedIOPS $ddIOPS `
                    -ProvisionedMBps $ddMBps
            }

            [void]$allDisks.Add([PSCustomObject]@{
                SubscriptionName    = $sub.Name
                SubscriptionId      = $sub.Id
                ResourceGroup       = $vm.ResourceGroupName
                VMName              = $vm.Name
                DiskName            = $dd.Name
                DiskType            = 'Data Disk'
                DiskSKU             = $ddSku
                DiskTier            = $ddTier
                DiskSizeGB          = $ddSizeGB
                ProvisionedIOPS     = $ddIOPS
                ProvisionedMBps     = $ddMBps
                Caching             = $dd.Caching
                ManagedDiskId       = $ddMgdId
                LUN                 = $dd.Lun
                DiskState           = $ddState
                EncryptionType      = $ddEncType
                Region              = $region
                EstMonthlyPrice_USD = $ddPrice
            })
        }

        # ── Compute VM row totals ──────────────────────────────────
        $vmDiskRows = @($allDisks | Where-Object { $_.VMName -eq $vm.Name -and $_.SubscriptionId -eq $sub.Id })
        $diskSum = 0
        $m = $vmDiskRows | Measure-Object -Property EstMonthlyPrice_USD -Sum
        if ($m -and $null -ne $m.Sum) {
            $diskSum = [double]$m.Sum
        }

        $imgRef   = $vm.StorageProfile.ImageReference
        $osProf   = $vm.OSProfile
        $diagProf = $vm.DiagnosticsProfile
        $bootDiag = $false

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

        [void]$allVMs.Add([PSCustomObject]@{
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
            VMPriceBillingModel     = $vmPriceInfo.BillingModel
            VMPriceProductName      = $vmPriceInfo.ProductName
            VMPriceMeterName        = $vmPriceInfo.MeterName
            VMPriceSkuNameReturned  = $vmPriceInfo.SkuNameReturned
            VMPriceUnitOfMeasure    = $vmPriceInfo.UnitOfMeasure
            VMPriceEffectiveDate    = $vmPriceInfo.EffectiveDate

            EstTotalDiskCost_USD    = [math]::Round($diskSum, 2)
            EstTotalMonthlyCost_USD = [math]::Round($(if ($null -ne $vmMonthlyPrice) { $vmMonthlyPrice + $diskSum } else { $diskSum }), 2)
        })

        $pd = if ($null -ne $vmHourlyPrice) { $vmHourlyPrice } else { 'N/A' }
        Write-Host "    ✓ $($vm.Name) [$vmSku] | VM: `$$pd/hr | $($vmPriceInfo.BillingModel)" -ForegroundColor DarkGreen
        Write-Host "      ↳ PriceSource: Product='$($vmPriceInfo.ProductName)' | Meter='$($vmPriceInfo.MeterName)' | ReturnedSku='$($vmPriceInfo.SkuNameReturned)'" -ForegroundColor DarkGray
    }
}

#endregion

#region ── Cost Summary ────────────────────────────────────────────────────────

$costSummary = foreach ($grp in ($allVMs | Group-Object SubscriptionName)) {
    $sn      = $grp.Name
    $gvms    = @($grp.Group)
    $gdisks  = @($allDisks | Where-Object { $_.SubscriptionName -eq $sn })

    $mVM    = $gvms   | Measure-Object -Property VMMonthlyPrice_USD      -Sum
    $mDisk  = $gdisks | Measure-Object -Property EstMonthlyPrice_USD     -Sum
    $mTotal = $gvms   | Measure-Object -Property EstTotalMonthlyCost_USD -Sum
    $mDD    = $gvms   | Measure-Object -Property DataDiskCount           -Sum

    $vmCost    = if ($null -ne $mVM.Sum)    { [math]::Round([double]$mVM.Sum,    2) } else { 0 }
    $diskCost  = if ($null -ne $mDisk.Sum)  { [math]::Round([double]$mDisk.Sum,  2) } else { 0 }
    $totalCost = if ($null -ne $mTotal.Sum) { [math]::Round([double]$mTotal.Sum, 2) } else { 0 }
    $ddCount   = if ($null -ne $mDD.Sum)    { [int]$mDD.Sum } else { 0 }

    [PSCustomObject]@{
        SubscriptionName          = $sn
        TotalVMs                  = $gvms.Count
        RunningVMs                = @($gvms | Where-Object { $_.PowerState -match 'running' }).Count
        DeallocatedVMs            = @($gvms | Where-Object { $_.PowerState -match 'deallocated' }).Count
        StoppedVMs                = @($gvms | Where-Object { $_.PowerState -match '^VM stopped$' }).Count
        TotalDataDisks            = $ddCount
        VMComputeCost_Monthly_USD = $vmCost
        DiskCost_Monthly_USD      = $diskCost
        TotalCost_Monthly_USD     = $totalCost
    }
}

#endregion

#region ── Excel Export ────────────────────────────────────────────────────────

Write-Host "`n[INFO] Exporting to Excel: $OutputPath" -ForegroundColor Cyan

# ── Sheet 1: VM Inventory ─────────────────────────────────────────────────────
$excelPkg = $allVMs | Export-Excel `
    -Path $OutputPath `
    -WorksheetName 'VM Inventory' `
    -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
    -TableName 'VMInventory' `
    -TableStyle 'Medium9' `
    -PassThru

$ws = $excelPkg.Workbook.Worksheets['VM Inventory']

if ($allVMs.Count -gt 0) {
    $headers = @($allVMs[0].PSObject.Properties.Name)
    $lastRow = $allVMs.Count + 1

    $pwrIdx = [array]::IndexOf($headers, 'PowerState')
    if ($pwrIdx -ge 0) {
        $col   = Get-ExcelColumn -Col ($pwrIdx + 1)
        $range = "${col}2:${col}${lastRow}"
        Add-ConditionalFormatting -Worksheet $ws -Address $range -RuleType ContainsText -ConditionValue 'running'     -BackgroundColor ([System.Drawing.Color]::LightGreen)
        Add-ConditionalFormatting -Worksheet $ws -Address $range -RuleType ContainsText -ConditionValue 'deallocated' -BackgroundColor ([System.Drawing.Color]::LightCoral)
        Add-ConditionalFormatting -Worksheet $ws -Address $range -RuleType ContainsText -ConditionValue 'stopped'     -BackgroundColor ([System.Drawing.Color]::LightYellow)
        Add-ConditionalFormatting -Worksheet $ws -Address $range -RuleType ContainsText -ConditionValue 'Unknown'     -BackgroundColor ([System.Drawing.Color]::LightGray)
    }

    foreach ($colName in @('VMHourlyPrice_USD','VMMonthlyPrice_USD','EstTotalDiskCost_USD','EstTotalMonthlyCost_USD')) {
        $idx = [array]::IndexOf($headers, $colName)
        if ($idx -ge 0) {
            $ws.Column($idx + 1).Style.Numberformat.Format = '$#,##0.0000'
        }
    }

    $totalIdx = [array]::IndexOf($headers, 'EstTotalMonthlyCost_USD')
    if ($totalIdx -ge 0) {
        $col   = Get-ExcelColumn -Col ($totalIdx + 1)
        $range = "${col}2:${col}${lastRow}"
        Add-ConditionalFormatting -Worksheet $ws -Address $range -RuleType GreaterThan -ConditionValue 1000 -BackgroundColor ([System.Drawing.Color]::FromArgb(255, 200, 100))
    }
}

# ── Sheet 2: Disk Inventory ───────────────────────────────────────────────────
$excelPkg = $allDisks | Export-Excel `
    -ExcelPackage $excelPkg `
    -WorksheetName 'Disk Inventory' `
    -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
    -TableName 'DiskInventory' `
    -TableStyle 'Medium6' `
    -PassThru

$wsDisk = $excelPkg.Workbook.Worksheets['Disk Inventory']

if ($allDisks.Count -gt 0) {
    $dHeaders = @($allDisks[0].PSObject.Properties.Name)
    $dLastRow = $allDisks.Count + 1

    $dtIdx = [array]::IndexOf($dHeaders, 'DiskType')
    if ($dtIdx -ge 0) {
        $col   = Get-ExcelColumn -Col ($dtIdx + 1)
        $range = "${col}2:${col}${dLastRow}"
        Add-ConditionalFormatting -Worksheet $wsDisk -Address $range -RuleType ContainsText -ConditionValue 'OS Disk'   -BackgroundColor ([System.Drawing.Color]::LightSteelBlue)
        Add-ConditionalFormatting -Worksheet $wsDisk -Address $range -RuleType ContainsText -ConditionValue 'Data Disk' -BackgroundColor ([System.Drawing.Color]::LightCyan)
    }

    $dpIdx = [array]::IndexOf($dHeaders, 'EstMonthlyPrice_USD')
    if ($dpIdx -ge 0) {
        $wsDisk.Column($dpIdx + 1).Style.Numberformat.Format = '$#,##0.00'
    }

    $dsIdx = [array]::IndexOf($dHeaders, 'DiskState')
    if ($dsIdx -ge 0) {
        $col   = Get-ExcelColumn -Col ($dsIdx + 1)
        $range = "${col}2:${col}${dLastRow}"
        Add-ConditionalFormatting -Worksheet $wsDisk -Address $range -RuleType ContainsText -ConditionValue 'Unattached' -BackgroundColor ([System.Drawing.Color]::LightCoral)
    }
}

# ── Sheet 3: Cost Summary ─────────────────────────────────────────────────────
$excelPkg = $costSummary | Export-Excel `
    -ExcelPackage $excelPkg `
    -WorksheetName 'Cost Summary' `
    -AutoSize -FreezeTopRow -BoldTopRow `
    -TableName 'CostSummary' `
    -TableStyle 'Medium2' `
    -PassThru

$wsSum = $excelPkg.Workbook.Worksheets['Cost Summary']

if (@($costSummary).Count -gt 0) {
    $sHeaders = @(@($costSummary)[0].PSObject.Properties.Name)
    $sLastRow = @($costSummary).Count + 1

    foreach ($colName in @('VMComputeCost_Monthly_USD','DiskCost_Monthly_USD','TotalCost_Monthly_USD')) {
        $idx = [array]::IndexOf($sHeaders, $colName)
        if ($idx -ge 0) {
            $wsSum.Column($idx + 1).Style.Numberformat.Format = '$#,##0.00'
        }
    }

    $tcIdx = [array]::IndexOf($sHeaders, 'TotalCost_Monthly_USD')
    if ($tcIdx -ge 0) {
        $col   = Get-ExcelColumn -Col ($tcIdx + 1)
        $range = "${col}2:${col}${sLastRow}"
        Add-ConditionalFormatting -Worksheet $wsSum -Address $range -DataBarColor ([System.Drawing.Color]::SteelBlue)
    }
}

Close-ExcelPackage $excelPkg

Write-Host "`n[DONE] Report saved to: $OutputPath" -ForegroundColor Green
Write-Host "  VM Inventory   : $($allVMs.Count) VM(s)"          -ForegroundColor White
Write-Host "  Disk Inventory : $($allDisks.Count) disk(s)"      -ForegroundColor White
Write-Host "  Cost Summary   : $(@($costSummary).Count) sub(s)" -ForegroundColor White

#endregion