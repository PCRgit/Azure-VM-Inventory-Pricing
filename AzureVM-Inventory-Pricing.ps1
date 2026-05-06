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
      Install-Module Az          -Scope CurrentUser -Force
      Install-Module ImportExcel -Scope CurrentUser -Force

    Authenticate first:
      Connect-AzAccount -TenantId "<your-tenant-id>"
#>

#Requires -Modules Az.Accounts, Az.Compute, Az.Network, ImportExcel

[CmdletBinding()]
param(
    [string]   $OutputPath            = "$env:USERPROFILE\Desktop\AzureVM_Inventory_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx",
    [string]   $CurrencyCode          = "USD",
    [string[]] $ExcludeSubscriptions  = @(),
    [string]   $TenantId              = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

# Load System.Web for proper URL encoding
Add-Type -AssemblyName System.Web

$script:PriceCache = @{}

#region ── Pricing helpers ─────────────────────────────────────────────────────

function Build-PriceUrl {
    param([string]$Filter)
    # Encode ONLY the filter value — keep $filter literal in the URL key
    $encodedFilter = [System.Web.HttpUtility]::UrlEncode($Filter)
    return "https://prices.azure.com/api/retail/prices?api-version=2023-01-01-preview&currencyCode=$CurrencyCode&`$filter=$encodedFilter"
}

function Invoke-RetailPriceRequest {
    param([string]$Uri)

    $maxRetries = 4
    $attempt    = 0

    do {
        try {
            return Invoke-RestMethod -Uri $Uri -Method Get -UseBasicParsing -ErrorAction Stop
        } catch {
            $attempt++
            if ($_.Exception.Message -match '429|Too many requests' -and $attempt -lt $maxRetries) {
                $wait = 5 * $attempt
                Write-Warning "  Rate limited. Retrying in $wait sec (attempt $attempt of $maxRetries)..."
                Start-Sleep -Seconds $wait
            } else { throw }
        }
    } while ($attempt -lt $maxRetries)
}

function Get-AzRetailPrice {
    param([string]$Filter, [string]$CacheKey)

    if ($script:PriceCache.ContainsKey($CacheKey)) {
        return $script:PriceCache[$CacheKey]
    }

    try {
        $uri      = Build-PriceUrl -Filter $Filter
        $allItems = [System.Collections.Generic.List[object]]::new()

        do {
            $resp = Invoke-RetailPriceRequest -Uri $uri
            if ($resp -and $resp.Items) {
                $resp.Items | ForEach-Object { $allItems.Add($_) }
            }
            $uri = if ($resp -and $resp.NextPageLink -and $resp.NextPageLink -ne '') { $resp.NextPageLink } else { $null }
        } while ($uri)

        $item = @($allItems) |
            Where-Object { $_.type -eq 'Consumption' -and $_.meterName -notmatch 'Spot|Low Priority' } |
            Sort-Object effectiveStartDate -Descending |
            Select-Object -First 1

        $price = if ($item) { [math]::Round([double]$item.retailPrice, 6) } else { $null }
        $script:PriceCache[$CacheKey] = $price
        return $price
    } catch {
        Write-Warning "  Pricing API error [$CacheKey]: $($_.Exception.Message)"
        $script:PriceCache[$CacheKey] = $null
        return $null
    }
}

function Get-VMHourlyPrice {
    param([string]$SkuName, [string]$Region)
    $key    = "VM|$SkuName|$Region|$CurrencyCode"
    $filter = "serviceName eq 'Virtual Machines' and armSkuName eq '$SkuName' and armRegionName eq '$Region' and priceType eq 'Consumption'"
    return Get-AzRetailPrice -Filter $filter -CacheKey $key
}

function Get-DiskTierLabel {
    param([string]$DiskSkuName, [int]$DiskSizeGB)

    if ($DiskSkuName -match '^Ultra')        { return 'Ultra' }
    if ($DiskSkuName -match '^Premium')       { $f = 'P' }
    elseif ($DiskSkuName -match '^StandardSSD') { $f = 'E' }
    else                                         { $f = 'S' }

    $n = switch ($DiskSizeGB) {
        { $_ -le 4    } { 1;  break }
        { $_ -le 8    } { 2;  break }
        { $_ -le 16   } { 3;  break }
        { $_ -le 32   } { 4;  break }
        { $_ -le 64   } { 6;  break }
        { $_ -le 128  } { 10; break }
        { $_ -le 256  } { 15; break }
        { $_ -le 512  } { 20; break }
        { $_ -le 1024 } { 30; break }
        { $_ -le 2048 } { 40; break }
        { $_ -le 4096 } { 50; break }
        { $_ -le 8192 } { 60; break }
        { $_ -le 16384} { 70; break }
        default          { 80 }
    }
    return "$f$n"
}

function Get-DiskMonthlyPrice {
    param([string]$DiskSkuName, [string]$Region, [int]$DiskSizeGB)

    $tier = Get-DiskTierLabel -DiskSkuName $DiskSkuName -DiskSizeGB $DiskSizeGB
    if ($tier -eq 'Ultra') { return $null }

    $productName = switch -Regex ($DiskSkuName) {
        '^Premium'      { 'Premium SSD Managed Disks';  break }
        '^StandardSSD'  { 'Standard SSD Managed Disks'; break }
        '^Ultra'        { 'Ultra Disks';                break }
        default         { 'Standard HDD Managed Disks' }
    }

    $redundancy = if ($DiskSkuName -match 'ZRS') { 'ZRS' } else { 'LRS' }
    $skuName    = "$tier $redundancy"
    $key        = "DISK|$productName|$skuName|$Region|$CurrencyCode"
    $filter     = "serviceFamily eq 'Storage' and armRegionName eq '$Region' and skuName eq '$skuName' and productName eq '$productName' and priceType eq 'Consumption'"

    return Get-AzRetailPrice -Filter $filter -CacheKey $key
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
        if ($s.PSObject.Properties.Name -contains 'PowerState' -and $s.PowerState) { return $s.PowerState }
        if ($s.Statuses) {
            $ps = (@($s.Statuses) | Where-Object { $_.Code -match 'PowerState' } | Select-Object -First 1).DisplayStatus
            if ($ps) { return $ps }
        }
    } catch {}
    return 'Unknown'
}

function Get-SafeProp {
    param($Object, [string]$Prop)
    if ($null -ne $Object -and $Object.PSObject.Properties.Name -contains $Prop) { return $Object.$Prop }
    return $null
}

function Get-ExcelColumn {
    param([int]$Col)
    $s = ""; $d = $Col
    while ($d -gt 0) {
        $m = ($d - 1) % 26
        $s = [char](65 + $m) + $s
        $d = [math]::Floor(($d - $m) / 26)
    }
    return $s
}

#endregion

#region ── Resolve tenant + subscriptions ──────────────────────────────────────

if (-not $TenantId) {
    $ctx = Get-AzContext -ErrorAction SilentlyContinue
    if ($ctx -and $ctx.Tenant) { $TenantId = $ctx.Tenant.Id }
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
    } catch {
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
                $nicNames.Add($nic.Name)
                if ($nic.EnableAcceleratedNetworking) { $accelNet = $true }

                foreach ($ip in @($nic.IpConfigurations)) {
                    if ($ip.PrivateIpAddress) { $privateIPs.Add($ip.PrivateIpAddress) }

                    if ($ip.Subnet -and $ip.Subnet.Id) {
                        $sp = $ip.Subnet.Id -split '/'
                        if ($sp.Count -ge 11) { $vnetNames.Add($sp[8]); $subnetNames.Add($sp[10]) }
                    }

                    if ($ip.PublicIpAddress -and $ip.PublicIpAddress.Id) {
                        $pp = $ip.PublicIpAddress.Id -split '/'
                        try {
                            $pip = Get-AzPublicIpAddress -Name $pp[-1] -ResourceGroupName $pp[4] -ErrorAction Stop
                            $publicIPs.Add($(if ($pip.IpAddress) { $pip.IpAddress } else { 'Dynamic/Unassigned' }))
                        } catch { $publicIPs.Add('N/A') }
                    }
                }

                if ($nic.NetworkSecurityGroup -and $nic.NetworkSecurityGroup.Id) {
                    $nsgNames.Add(($nic.NetworkSecurityGroup.Id -split '/')[-1])
                }
            } catch {
                Write-Warning "  NIC '$nicN': $($_.Exception.Message)"
            }
        }

        $region         = $vm.Location
        $vmSku          = $vm.HardwareProfile.VmSize
        $vmHourlyPrice  = Get-VMHourlyPrice -SkuName $vmSku -Region $region
        $vmMonthlyPrice = $null
        if ($null -ne $vmHourlyPrice) {
            $vmMonthlyPrice = [math]::Round([double]$vmHourlyPrice * 730, 2)
        }

        $tagString = ''
        if ($vm.Tags) {
            $tagString = ($vm.Tags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
        }

        # ── OS Disk ───────────────────────────────────────────────
        $osDisk        = $vm.StorageProfile.OsDisk
        $osDiskRes     = $null
        $osDiskSizeGB  = $null
        $osDiskSku     = 'Unknown'
        $osDiskState   = $null
        $osDiskEncType = $null
        $osDiskMgdId   = $null
        $osDiskPrice   = $null

        if ($osDisk.ManagedDisk -and $osDisk.ManagedDisk.Id) {
            $osDiskMgdId = $osDisk.ManagedDisk.Id
            $dp = $osDisk.ManagedDisk.Id -split '/'
            try {
                $osDiskRes = Get-AzDisk -ResourceGroupName $dp[4] -DiskName $dp[-1] -ErrorAction Stop
            } catch { $osDiskRes = $null }
        }

        if ($osDiskRes) {
            $osDiskSizeGB  = $osDiskRes.DiskSizeGB
            $osDiskSku     = $osDiskRes.Sku.Name
            $osDiskState   = Get-SafeProp -Object $osDiskRes -Prop 'DiskState'
            $osDiskEncType = Get-SafeProp -Object $osDiskRes.Encryption -Prop 'Type'
        } else {
            $osDiskSizeGB = Get-SafeProp -Object $osDisk -Prop 'DiskSizeGB'
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

            if ($dd.ManagedDisk -and $dd.ManagedDisk.Id) {
                $ddMgdId = $dd.ManagedDisk.Id
                $dp = $dd.ManagedDisk.Id -split '/'
                try {
                    $ddRes = Get-AzDisk -ResourceGroupName $dp[4] -DiskName $dp[-1] -ErrorAction Stop
                } catch { $ddRes = $null }
            }

            if ($ddRes) {
                $ddSizeGB  = $ddRes.DiskSizeGB
                $ddSku     = $ddRes.Sku.Name
                $ddState   = Get-SafeProp -Object $ddRes -Prop 'DiskState'
                $ddEncType = Get-SafeProp -Object $ddRes.Encryption -Prop 'Type'
            } else {
                $ddSizeGB = Get-SafeProp -Object $dd -Prop 'DiskSizeGB'
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
        if ($m -and $null -ne $m.Sum) { $diskSum = [double]$m.Sum }

        $imgRef   = $vm.StorageProfile.ImageReference
        $osProf   = $vm.OSProfile
        $diagProf = $vm.DiagnosticsProfile
        $bootDiag = $false
        if ($diagProf -and $diagProf.BootDiagnostics -and $null -ne $diagProf.BootDiagnostics.Enabled) {
            $bootDiag = [bool]$diagProf.BootDiagnostics.Enabled
        }

        $availSet = $null; if ($vm.AvailabilitySetReference -and $vm.AvailabilitySetReference.Id) { $availSet = ($vm.AvailabilitySetReference.Id -split '/')[-1] }
        $vmssRef  = $null; if ($vm.VirtualMachineScaleSet -and $vm.VirtualMachineScaleSet.Id)   { $vmssRef  = ($vm.VirtualMachineScaleSet.Id -split '/')[-1]  }
        $ppgRef   = $null; if ($vm.ProximityPlacementGroup -and $vm.ProximityPlacementGroup.Id) { $ppgRef   = ($vm.ProximityPlacementGroup.Id -split '/')[-1]  }

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
            OSOffer                 = $(if ($imgRef) { $imgRef.Offer }     else { $null })
            OSSKU                   = $(if ($imgRef) { $imgRef.Sku }       else { $null })
            OSVersion               = $(if ($imgRef) { $imgRef.Version }   else { $null })
            ComputerName            = $(if ($osProf) { $osProf.ComputerName }  else { $null })
            AdminUsername           = $(if ($osProf) { $osProf.AdminUsername } else { $null })
            AvailabilitySet         = $availSet
            VirtualMachineScaleSet  = $vmssRef
            AvailabilityZone        = ($vm.Zones -join ',')
            ProximityPlacementGroup = $ppgRef
            LicenseType             = $vm.LicenseType
            BootDiagnosticsEnabled  = $bootDiag
            AcceleratedNetworking   = $accelNet
            NICNames                = ($nicNames    -join ', ')
            PrivateIPAddresses      = ($privateIPs  -join ', ')
            PublicIPAddresses       = ($publicIPs   -join ', ')
            VNetNames               = (($vnetNames  | Select-Object -Unique) -join ', ')
            SubnetNames             = (($subnetNames | Select-Object -Unique) -join ', ')
            NSGNames                = (($nsgNames   | Select-Object -Unique) -join ', ')
            OSDiskName              = $osDisk.Name
            OSDiskSizeGB            = $osDiskSizeGB
            OSDiskSKU               = $osDiskSku
            DataDiskCount           = @($vm.StorageProfile.DataDisks).Count
            Tags                    = $tagString
            VMHourlyPrice_USD       = $vmHourlyPrice
            VMMonthlyPrice_USD      = $vmMonthlyPrice
            EstTotalDiskCost_USD    = [math]::Round($diskSum, 2)
            EstTotalMonthlyCost_USD = [math]::Round($(if ($null -ne $vmMonthlyPrice) { $vmMonthlyPrice + $diskSum } else { $diskSum }), 2)
        })

        $pd = if ($null -ne $vmHourlyPrice) { $vmHourlyPrice } else { 'N/A' }
        Write-Host "    ✓ $($vm.Name) [$vmSku] | VM: `$$pd/hr" -ForegroundColor DarkGreen
    }
}

#endregion

#region ── Cost Summary ────────────────────────────────────────────────────────

$costSummary = foreach ($grp in ($allVMs | Group-Object SubscriptionName)) {
    $sn      = $grp.Name
    $gvms    = @($grp.Group)
    $gdisks  = @($allDisks | Where-Object { $_.SubscriptionName -eq $sn })

    $mVM    = $gvms   | Measure-Object -Property VMMonthlyPrice_USD   -Sum
    $mDisk  = $gdisks | Measure-Object -Property EstMonthlyPrice_USD  -Sum
    $mTotal = $gvms   | Measure-Object -Property EstTotalMonthlyCost_USD -Sum
    $mDD    = $gvms   | Measure-Object -Property DataDiskCount         -Sum

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

#region ── Excel Export (Improved) ────────────────────────────────────────────

Write-Host "`n[INFO] Exporting to Excel: $OutputPath" -ForegroundColor Cyan

# ── Sheet 1: VM Inventory ─────────────────────────────────────────────────────
$excelPkg = $allVMs | Export-Excel `
    -Path          $OutputPath `
    -WorksheetName 'VM Inventory' `
    -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
    -TableName     'VMInventory' `
    -TableStyle    'Medium9' `
    -PassThru

$ws = $excelPkg.Workbook.Worksheets['VM Inventory']

if ($allVMs.Count -gt 0) {
    $headers = @($allVMs[0].PSObject.Properties.Name)
    $lastRow = $allVMs.Count + 1

    # Power State — color coding
    $pwrIdx = [array]::IndexOf($headers, 'PowerState')
    if ($pwrIdx -ge 0) {
        $col   = Get-ExcelColumn -Col ($pwrIdx + 1)
        $range = "${col}2:${col}${lastRow}"
        Add-ConditionalFormatting -Worksheet $ws -Address $range -RuleType ContainsText -ConditionValue 'running'     -BackgroundColor ([System.Drawing.Color]::LightGreen)
        Add-ConditionalFormatting -Worksheet $ws -Address $range -RuleType ContainsText -ConditionValue 'deallocated' -BackgroundColor ([System.Drawing.Color]::LightCoral)
        Add-ConditionalFormatting -Worksheet $ws -Address $range -RuleType ContainsText -ConditionValue 'stopped'     -BackgroundColor ([System.Drawing.Color]::LightYellow)
        Add-ConditionalFormatting -Worksheet $ws -Address $range -RuleType ContainsText -ConditionValue 'Unknown'     -BackgroundColor ([System.Drawing.Color]::LightGray)
    }

    # Format currency columns with $ number format
    foreach ($col in @('VMHourlyPrice_USD','VMMonthlyPrice_USD','EstTotalDiskCost_USD','EstTotalMonthlyCost_USD')) {
        $idx = [array]::IndexOf($headers, $col)
        if ($idx -ge 0) {
            $colLetter = Get-ExcelColumn -Col ($idx + 1)
            $ws.Column($idx + 1).Style.Numberformat.Format = '$#,##0.0000'
        }
    }

    # Highlight high-cost VMs (monthly total > $1000) in orange
    $totalIdx = [array]::IndexOf($headers, 'EstTotalMonthlyCost_USD')
    if ($totalIdx -ge 0) {
        $col   = Get-ExcelColumn -Col ($totalIdx + 1)
        $range = "${col}2:${col}${lastRow}"
        Add-ConditionalFormatting -Worksheet $ws -Address $range -RuleType GreaterThan -ConditionValue 1000 -BackgroundColor ([System.Drawing.Color]::FromArgb(255, 200, 100))
    }
}

# ── Sheet 2: Disk Inventory ───────────────────────────────────────────────────
$excelPkg = $allDisks | Export-Excel `
    -ExcelPackage  $excelPkg `
    -WorksheetName 'Disk Inventory' `
    -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
    -TableName     'DiskInventory' `
    -TableStyle    'Medium6' `
    -PassThru

$wsDisk = $excelPkg.Workbook.Worksheets['Disk Inventory']

if ($allDisks.Count -gt 0) {
    $dHeaders = @($allDisks[0].PSObject.Properties.Name)
    $dLastRow = $allDisks.Count + 1

    # Color disk type column
    $dtIdx = [array]::IndexOf($dHeaders, 'DiskType')
    if ($dtIdx -ge 0) {
        $col   = Get-ExcelColumn -Col ($dtIdx + 1)
        $range = "${col}2:${col}${dLastRow}"
        Add-ConditionalFormatting -Worksheet $wsDisk -Address $range -RuleType ContainsText -ConditionValue 'OS Disk'   -BackgroundColor ([System.Drawing.Color]::LightSteelBlue)
        Add-ConditionalFormatting -Worksheet $wsDisk -Address $range -RuleType ContainsText -ConditionValue 'Data Disk' -BackgroundColor ([System.Drawing.Color]::LightCyan)
    }

    # Format disk price column
    $dpIdx = [array]::IndexOf($dHeaders, 'EstMonthlyPrice_USD')
    if ($dpIdx -ge 0) {
        $wsDisk.Column($dpIdx + 1).Style.Numberformat.Format = '$#,##0.00'
    }

    # Color DiskState column
    $dsIdx = [array]::IndexOf($dHeaders, 'DiskState')
    if ($dsIdx -ge 0) {
        $col   = Get-ExcelColumn -Col ($dsIdx + 1)
        $range = "${col}2:${col}${dLastRow}"
        Add-ConditionalFormatting -Worksheet $wsDisk -Address $range -RuleType ContainsText -ConditionValue 'Unattached' -BackgroundColor ([System.Drawing.Color]::LightCoral)
    }
}

# ── Sheet 3: Cost Summary ─────────────────────────────────────────────────────
$excelPkg = $costSummary | Export-Excel `
    -ExcelPackage  $excelPkg `
    -WorksheetName 'Cost Summary' `
    -AutoSize -FreezeTopRow -BoldTopRow `
    -TableName     'CostSummary' `
    -TableStyle    'Medium2' `
    -PassThru

$wsSum = $excelPkg.Workbook.Worksheets['Cost Summary']

if (@($costSummary).Count -gt 0) {
    $sHeaders = @(@($costSummary)[0].PSObject.Properties.Name)
    $sLastRow = @($costSummary).Count + 1

    foreach ($col in @('VMComputeCost_Monthly_USD','DiskCost_Monthly_USD','TotalCost_Monthly_USD')) {
        $idx = [array]::IndexOf($sHeaders, $col)
        if ($idx -ge 0) {
            $wsSum.Column($idx + 1).Style.Numberformat.Format = '$#,##0.00'
        }
    }

    # Highlight total cost column with data bars
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