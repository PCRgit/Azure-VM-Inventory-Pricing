<#
.SYNOPSIS
    Azure VM Right-Sizing & Cost Savings Report

.DESCRIPTION
    For every VM across all subscriptions, pulls 30 days of native Azure Monitor
    metrics (CPU, Available Memory, Disk IOPS, Network) with NO dependency on
    Log Analytics. Combines with Azure Advisor right-sizing recommendations and
    live retail pricing to calculate potential monthly savings.
    Exports a multi-sheet Excel workbook with colour-coded recommendations.

.NOTES
    Required modules:
      Install-Module Az          -Scope CurrentUser -Force
      Install-Module ImportExcel -Scope CurrentUser -Force
    Connect-AzAccount before running.

    Thresholds (all tunable via parameters):
      - Idle      : Avg CPU <= 5%  AND Avg Mem Available >= 90% of total
      - Underused : Avg CPU <= 20% AND P95 CPU <= 40%
      - Overused  : Avg CPU >= 85% OR  P95 CPU >= 95%
      - OK        : everything else
#>

#Requires -Modules Az.Accounts, Az.Compute, Az.Monitor, ImportExcel

[CmdletBinding()]
param(
    [string]   $OutputPath            = "$env:USERPROFILE\Desktop\AzureVM_RightSizing_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx",
    [string]   $CurrencyCode          = 'USD',
    [string[]] $ExcludeSubscriptions  = @(),
    [int]      $LookbackDays          = 30,
    [double]   $IdleCpuThreshold      = 5.0,
    [double]   $UnderusedCpuThreshold = 20.0,
    [double]   $UnderusedP95Threshold = 40.0,
    [double]   $OverusedCpuThreshold  = 85.0,
    [double]   $OverusedP95Threshold  = 95.0
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName System.Web

$script:PriceCache = @{}

#region ── Helpers ─────────────────────────────────────────────────────────────

function Get-ExcelColumn {
    param([int]$Col)
    $s = ''; $d = [int]$Col
    while ($d -gt 0) {
        $m = [int](($d - 1) % 26)
        $s = [string][char](65 + $m) + $s
        $d = [int][math]::Floor(($d - $m) / 26)
    }
    return $s
}

function Get-AzRetailPrice {
    param([string]$Filter, [string]$CacheKey)
    if ($script:PriceCache.ContainsKey($CacheKey)) { return $script:PriceCache[$CacheKey] }
    try {
        $enc  = [System.Web.HttpUtility]::UrlEncode($Filter)
        $uri  = "https://prices.azure.com/api/retail/prices?api-version=2023-01-01-preview&currencyCode=$CurrencyCode&`$filter=$enc"
        $resp = Invoke-RestMethod -Uri $uri -Method Get -UseBasicParsing -ErrorAction Stop
        $item = @($resp.Items) |
            Where-Object { $_.type -eq 'Consumption' -and $_.meterName -notmatch 'Spot|Low Priority' } |
            Sort-Object effectiveStartDate -Descending | Select-Object -First 1
        $price = if ($item) { [math]::Round([double]$item.retailPrice, 6) } else { $null }
        $script:PriceCache[$CacheKey] = $price
        return $price
    } catch {
        $script:PriceCache[$CacheKey] = $null
        return $null
    }
}

function Get-VMPrice {
    param([string]$Sku, [string]$Region)
    $key    = "VM|$Sku|$Region|$CurrencyCode"
    $filter = "serviceName eq 'Virtual Machines' and armSkuName eq '$Sku' and armRegionName eq '$Region' and priceType eq 'Consumption'"
    return Get-AzRetailPrice -Filter $filter -CacheKey $key
}

# Get a single Azure Monitor metric stat over lookback period
function Get-VMMetricStat {
    param(
        [string]$ResourceId,
        [string]$MetricName,
        [string]$Aggregation,   # Average, Maximum, Minimum, Total, Count
        [datetime]$StartTime,
        [datetime]$EndTime
    )
    try {
        $result = Get-AzMetric `
            -ResourceId     $ResourceId `
            -MetricName     $MetricName `
            -StartTime      $StartTime `
            -EndTime        $EndTime `
            -TimeGrain      '01:00:00' `
            -AggregationType $Aggregation `
            -ErrorAction Stop `
            -WarningAction SilentlyContinue

        $values = @($result.Data | Where-Object { $null -ne $_.$Aggregation } | ForEach-Object { $_.$Aggregation })
        if ($values.Count -eq 0) { return $null }
        return [math]::Round(($values | Measure-Object -Average).Average, 2)
    } catch {
        return $null
    }
}

# Get P95 by pulling hourly Average and taking 95th percentile of that array
function Get-VMMetricP95 {
    param(
        [string]$ResourceId,
        [string]$MetricName,
        [datetime]$StartTime,
        [datetime]$EndTime
    )
    try {
        $result = Get-AzMetric `
            -ResourceId      $ResourceId `
            -MetricName      $MetricName `
            -StartTime       $StartTime `
            -EndTime         $EndTime `
            -TimeGrain       '01:00:00' `
            -AggregationType 'Average' `
            -ErrorAction Stop `
            -WarningAction SilentlyContinue

        $values = @($result.Data | Where-Object { $null -ne $_.Average } | ForEach-Object { $_.Average }) | Sort-Object
        if ($values.Count -eq 0) { return $null }
        $idx = [int][math]::Ceiling($values.Count * 0.95) - 1
        if ($idx -lt 0) { $idx = 0 }
        return [math]::Round($values[$idx], 2)
    } catch {
        return $null
    }
}

# VM SKU family → next smaller SKU (basic right-size mapping)
function Get-SuggestedSku {
    param([string]$CurrentSku, [string]$Region)
    # Derive series, vCPU count and propose half-step down
    # Pattern: Standard_D4s_v3 → Standard_D2s_v3
    if ($CurrentSku -match '^(Standard_[A-Za-z]+)(\d+)(.*)$') {
        $prefix  = $Matches[1]
        $cpuSize = [int]$Matches[2]
        $suffix  = $Matches[3]
        $newSize = [int][math]::Floor($cpuSize / 2)
        if ($newSize -ge 1) { return "${prefix}${newSize}${suffix}" }
    }
    return $null
}

function Get-Recommendation {
    param(
        [double]$AvgCpu,
        [double]$P95Cpu,
        [double]$AvgMemPct
    )
    if ($null -eq $AvgCpu) { return 'No Data' }
    if ($AvgCpu -le $IdleCpuThreshold) { return 'Idle — Consider Deallocate/Delete' }
    if ($AvgCpu -le $UnderusedCpuThreshold -and ($null -eq $P95Cpu -or $P95Cpu -le $UnderusedP95Threshold)) { return 'Underused — Downsize' }
    if ($AvgCpu -ge $OverusedCpuThreshold -or ($null -ne $P95Cpu -and $P95Cpu -ge $OverusedP95Threshold)) { return 'Overused — Upsize' }
    return 'Right-Sized'
}

#endregion

#region ── Advisor Recommendations ────────────────────────────────────────────

function Get-AdvisorRightSizeRecs {
    param([string]$SubscriptionId)
    $recs = @()
    try {
        $uri  = "https://management.azure.com/subscriptions/$SubscriptionId/providers/Microsoft.Advisor/recommendations?api-version=2023-01-01&`$filter=Category eq 'Cost'"
        $token = (Get-AzAccessToken -ResourceUrl 'https://management.azure.com' -ErrorAction Stop).Token
        $headers = @{ Authorization = "Bearer $token"; 'Content-Type' = 'application/json' }
        do {
            $resp = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -UseBasicParsing -ErrorAction Stop
            $recs += @($resp.value | Where-Object {
                $_.properties.impactedField -eq 'Microsoft.Compute/virtualMachines' -and
                $_.properties.recommendationTypeId -match 'e10b1381|4d2e7b53|48eda464|09c2ec8e'
            })
            $uri = if ($resp.nextLink) { $resp.nextLink } else { $null }
        } while ($uri)
    } catch {
        Write-Warning "  Could not retrieve Advisor recs for $SubscriptionId : $($_.Exception.Message)"
    }
    return $recs
}

#endregion

#region ── Main Collection ─────────────────────────────────────────────────────

$endTime   = (Get-Date).ToUniversalTime()
$startTime = $endTime.AddDays(-$LookbackDays)

Write-Host "`n[INFO] Lookback: $LookbackDays days ($($startTime.ToString('yyyy-MM-dd')) → $($endTime.ToString('yyyy-MM-dd')))" -ForegroundColor Cyan

$ctx   = Get-AzContext -ErrorAction SilentlyContinue
$tenId = if ($ctx -and $ctx.Tenant) { $ctx.Tenant.Id } else { '' }

$subscriptions = @(
    Get-AzSubscription -TenantId $tenId -ErrorAction SilentlyContinue |
    Where-Object { $_.State -eq 'Enabled' -and $_.Id -notin $ExcludeSubscriptions }
)

Write-Host "[INFO] Scanning $($subscriptions.Count) subscription(s).`n" -ForegroundColor Cyan

$report    = [System.Collections.Generic.List[object]]::new()
$subIdx    = 0

foreach ($sub in $subscriptions) {
    $subIdx++
    Write-Host "[$subIdx/$($subscriptions.Count)] $($sub.Name)" -ForegroundColor Yellow
    Set-AzContext -SubscriptionId $sub.Id -TenantId $tenId -ErrorAction SilentlyContinue | Out-Null

    # Pull Advisor recommendations once per subscription
    $advisorRecs = Get-AdvisorRightSizeRecs -SubscriptionId $sub.Id
    $advisorMap  = @{}
    foreach ($rec in $advisorRecs) {
        $vmName = $rec.properties.impactedValue
        $advisorMap[$vmName] = $rec
    }

    $vms = @(Get-AzVM -ErrorAction SilentlyContinue)
    if (-not $vms -or $vms.Count -eq 0) { Write-Host '  No VMs.' -ForegroundColor DarkGray; continue }

    Write-Host "  $($vms.Count) VM(s) found. Collecting metrics..." -ForegroundColor Green

    foreach ($vm in $vms) {
        $rid    = $vm.Id
        $region = $vm.Location
        $sku    = $vm.HardwareProfile.VmSize

        Write-Host "    → $($vm.Name) [$sku]" -ForegroundColor DarkCyan

        # ── Metrics ──────────────────────────────────────────────
        $avgCpu   = Get-VMMetricStat -ResourceId $rid -MetricName 'Percentage CPU'       -Aggregation 'Average' -StartTime $startTime -EndTime $endTime
        $maxCpu   = Get-VMMetricStat -ResourceId $rid -MetricName 'Percentage CPU'       -Aggregation 'Maximum' -StartTime $startTime -EndTime $endTime
        $p95Cpu   = Get-VMMetricP95  -ResourceId $rid -MetricName 'Percentage CPU'       -StartTime $startTime -EndTime $endTime
        $avgMemMB = Get-VMMetricStat -ResourceId $rid -MetricName 'Available Memory Bytes' -Aggregation 'Average' -StartTime $startTime -EndTime $endTime
        $minMemMB = Get-VMMetricStat -ResourceId $rid -MetricName 'Available Memory Bytes' -Aggregation 'Minimum' -StartTime $startTime -EndTime $endTime
        $avgNetIn = Get-VMMetricStat -ResourceId $rid -MetricName 'Network In Total'     -Aggregation 'Total'   -StartTime $startTime -EndTime $endTime
        $avgNetOut= Get-VMMetricStat -ResourceId $rid -MetricName 'Network Out Total'    -Aggregation 'Total'   -StartTime $startTime -EndTime $endTime
        $avgDiskR = Get-VMMetricStat -ResourceId $rid -MetricName 'Disk Read Operations/Sec'  -Aggregation 'Average' -StartTime $startTime -EndTime $endTime
        $avgDiskW = Get-VMMetricStat -ResourceId $rid -MetricName 'Disk Write Operations/Sec' -Aggregation 'Average' -StartTime $startTime -EndTime $endTime

        # Convert bytes → GB
        $avgMemGB = if ($null -ne $avgMemMB) { [math]::Round($avgMemMB / 1GB, 2) } else { $null }
        $minMemGB = if ($null -ne $minMemMB) { [math]::Round($minMemMB / 1GB, 2) } else { $null }

        # Network in/out MB average per day
        $netInMB  = if ($null -ne $avgNetIn)  { [math]::Round($avgNetIn  / 1MB, 1) } else { $null }
        $netOutMB = if ($null -ne $avgNetOut) { [math]::Round($avgNetOut / 1MB, 1) } else { $null }

        # ── Classification ────────────────────────────────────────
        $memPct = $null  # Available memory % — need total RAM; estimated from SKU below
        $recommendation = Get-Recommendation -AvgCpu $avgCpu -P95Cpu $p95Cpu -AvgMemPct $memPct

        # ── Sizing suggestion ─────────────────────────────────────
        $suggestedSku    = $null
        $suggestedPrice  = $null
        $currentPrice    = Get-VMPrice -Sku $sku -Region $region
        $monthlyCurrent  = if ($null -ne $currentPrice) { [math]::Round($currentPrice * 730, 2) } else { $null }
        $monthlySuggested= $null
        $estimatedSaving = $null

        if ($recommendation -match 'Underused|Idle') {
            # Check Advisor first
            $advisorRec = $advisorMap[$vm.Name]
            if ($advisorRec) {
                $extData = $advisorRec.properties.extendedProperties
                if ($extData -and $extData.recommendedResourceSku) {
                    $suggestedSku = $extData.recommendedResourceSku
                } elseif ($extData -and $extData.targetResourceId) {
                    $suggestedSku = $extData.targetResourceId
                }
            }
            if (-not $suggestedSku) {
                $suggestedSku = Get-SuggestedSku -CurrentSku $sku -Region $region
            }

            if ($suggestedSku) {
                $suggestedPrice   = Get-VMPrice -Sku $suggestedSku -Region $region
                $monthlySuggested = if ($null -ne $suggestedPrice) { [math]::Round($suggestedPrice * 730, 2) } else { $null }
                if ($null -ne $monthlyCurrent -and $null -ne $monthlySuggested) {
                    $estimatedSaving = [math]::Round($monthlyCurrent - $monthlySuggested, 2)
                    if ($estimatedSaving -lt 0) { $estimatedSaving = 0 }
                }
            }
        }

        # Advisor saving amount
        $advisorSaving = $null
        $advisorAction = $null
        $adv = $advisorMap[$vm.Name]
        if ($adv) {
            $ext = $adv.properties.extendedProperties
            if ($ext -and $ext.savingsAmount) { $advisorSaving = [double]$ext.savingsAmount }
            $advisorAction = $adv.properties.shortDescription.solution
        }

        $report.Add([PSCustomObject]@{
            SubscriptionName    = $sub.Name
            SubscriptionId      = $sub.Id
            ResourceGroup       = $vm.ResourceGroupName
            VMName              = $vm.Name
            Location            = $region
            CurrentSKU          = $sku
            OSType              = $vm.StorageProfile.OsDisk.OsType
            LookbackDays        = $LookbackDays
            # CPU
            AvgCPU_Pct          = $avgCpu
            MaxCPU_Pct          = $maxCpu
            P95CPU_Pct          = $p95Cpu
            # Memory
            AvgAvailMem_GB      = $avgMemGB
            MinAvailMem_GB      = $minMemGB
            # Disk
            AvgDiskReadIOPS     = $avgDiskR
            AvgDiskWriteIOPS    = $avgDiskW
            # Network
            TotalNetIn_MB       = $netInMB
            TotalNetOut_MB      = $netOutMB
            # Recommendation
            Recommendation      = $recommendation
            SuggestedSKU        = $suggestedSku
            CurrentPrice_Hr     = $currentPrice
            CurrentCost_Month   = $monthlyCurrent
            SuggestedPrice_Hr   = $suggestedPrice
            SuggestedCost_Month = $monthlySuggested
            EstSaving_Month     = $estimatedSaving
            EstSaving_Annual    = if ($null -ne $estimatedSaving) { [math]::Round($estimatedSaving * 12, 2) } else { $null }
            AdvisorRecommendation = $advisorAction
            AdvisorSaving_Month = $advisorSaving
        })
    }
}

#endregion

#region ── Summary aggregates ──────────────────────────────────────────────────

$summaryBySub = foreach ($g in ($report | Group-Object SubscriptionName)) {
    $rows = @($g.Group)
    [PSCustomObject]@{
        SubscriptionName   = $g.Name
        TotalVMs           = $rows.Count
        IdleVMs            = @($rows | Where-Object { $_.Recommendation -match 'Idle'      }).Count
        UnderusedVMs       = @($rows | Where-Object { $_.Recommendation -match 'Underused' }).Count
        RightSizedVMs      = @($rows | Where-Object { $_.Recommendation -eq  'Right-Sized' }).Count
        OverusedVMs        = @($rows | Where-Object { $_.Recommendation -match 'Overused'  }).Count
        NoDataVMs          = @($rows | Where-Object { $_.Recommendation -eq  'No Data'     }).Count
        TotalCurrentCost   = [math]::Round(($rows | Measure-Object CurrentCost_Month -Sum).Sum, 2)
        PotentialSaving_Mo = [math]::Round(($rows | Measure-Object EstSaving_Month   -Sum).Sum, 2)
        PotentialSaving_Yr = [math]::Round(($rows | Measure-Object EstSaving_Annual  -Sum).Sum, 2)
    }
}

$topSavers = $report |
    Where-Object { $null -ne $_.EstSaving_Month -and $_.EstSaving_Month -gt 0 } |
    Sort-Object EstSaving_Month -Descending |
    Select-Object -First 20

#endregion

#region ── Excel Export ────────────────────────────────────────────────────────

Write-Host "`n[INFO] Building Excel report: $OutputPath" -ForegroundColor Cyan

$reportArr  = @($report)
$summaryArr = @($summaryBySub)
$topArr     = @($topSavers)

# Totals row (for summary)
$grandTotal = [PSCustomObject]@{
    SubscriptionName   = '** TOTAL **'
    TotalVMs           = ($summaryArr | Measure-Object TotalVMs           -Sum).Sum
    IdleVMs            = ($summaryArr | Measure-Object IdleVMs            -Sum).Sum
    UnderusedVMs       = ($summaryArr | Measure-Object UnderusedVMs       -Sum).Sum
    RightSizedVMs      = ($summaryArr | Measure-Object RightSizedVMs      -Sum).Sum
    OverusedVMs        = ($summaryArr | Measure-Object OverusedVMs        -Sum).Sum
    NoDataVMs          = ($summaryArr | Measure-Object NoDataVMs          -Sum).Sum
    TotalCurrentCost   = [math]::Round(($summaryArr | Measure-Object TotalCurrentCost   -Sum).Sum, 2)
    PotentialSaving_Mo = [math]::Round(($summaryArr | Measure-Object PotentialSaving_Mo -Sum).Sum, 2)
    PotentialSaving_Yr = [math]::Round(($summaryArr | Measure-Object PotentialSaving_Yr -Sum).Sum, 2)
}
$summaryWithTotal = @($summaryArr) + @($grandTotal)

# ── Sheet 1: Executive Summary ───────────────────────────────────────────────
$pkg = $summaryWithTotal | Export-Excel `
    -Path $OutputPath `
    -WorksheetName 'Executive Summary' `
    -AutoSize -FreezeTopRow -BoldTopRow `
    -TableName 'ExecSummary' -TableStyle 'Medium9' `
    -PassThru

$wsSumm = $pkg.Workbook.Worksheets['Executive Summary']
$esh    = @($summaryWithTotal[0].PSObject.Properties.Name)

foreach ($col in @('TotalCurrentCost','PotentialSaving_Mo','PotentialSaving_Yr')) {
    $i = [array]::IndexOf($esh, $col)
    if ($i -ge 0) { $wsSumm.Column([int]($i+1)).Style.Numberformat.Format = '$#,##0.00' }
}

# Bold totals row
$totRow = $summaryWithTotal.Count + 1
$wsSumm.Row([int]$totRow).Style.Font.Bold = $true
$wsSumm.Row([int]$totRow).Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$wsSumm.Row([int]$totRow).Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(221,235,247))

# ── Sheet 2: Full VM Recommendations ────────────────────────────────────────
$pkg = $reportArr | Export-Excel `
    -ExcelPackage $pkg `
    -WorksheetName 'VM Recommendations' `
    -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
    -TableName 'VMRecs' -TableStyle 'Medium6' `
    -PassThru

$wsRec   = $pkg.Workbook.Worksheets['VM Recommendations']
$recH    = @($reportArr[0].PSObject.Properties.Name)
$recLast = $reportArr.Count + 1

foreach ($col in @('CurrentPrice_Hr','CurrentCost_Month','SuggestedPrice_Hr','SuggestedCost_Month','EstSaving_Month','EstSaving_Annual','AdvisorSaving_Month')) {
    $i = [array]::IndexOf($recH, $col)
    if ($i -ge 0) { $wsRec.Column([int]($i+1)).Style.Numberformat.Format = '$#,##0.00' }
}

foreach ($col in @('AvgCPU_Pct','MaxCPU_Pct','P95CPU_Pct')) {
    $i = [array]::IndexOf($recH, $col)
    if ($i -ge 0) { $wsRec.Column([int]($i+1)).Style.Numberformat.Format = '0.00%' }
}

# Colour-code Recommendation column
$recIdx = [array]::IndexOf($recH, 'Recommendation')
if ($recIdx -ge 0) {
    $col = Get-ExcelColumn ([int]($recIdx + 1))
    $range = "${col}2:${col}${recLast}"
    Add-ConditionalFormatting -Worksheet $wsRec -Address $range -RuleType ContainsText -ConditionValue 'Idle'        -BackgroundColor ([System.Drawing.Color]::LightCoral)
    Add-ConditionalFormatting -Worksheet $wsRec -Address $range -RuleType ContainsText -ConditionValue 'Underused'   -BackgroundColor ([System.Drawing.Color]::LightYellow)
    Add-ConditionalFormatting -Worksheet $wsRec -Address $range -RuleType ContainsText -ConditionValue 'Right-Sized' -BackgroundColor ([System.Drawing.Color]::LightGreen)
    Add-ConditionalFormatting -Worksheet $wsRec -Address $range -RuleType ContainsText -ConditionValue 'Overused'    -BackgroundColor ([System.Drawing.Color]::FromArgb(255,153,51))
    Add-ConditionalFormatting -Worksheet $wsRec -Address $range -RuleType ContainsText -ConditionValue 'No Data'     -BackgroundColor ([System.Drawing.Color]::LightGray)
}

# Data bars on savings
$savIdx = [array]::IndexOf($recH, 'EstSaving_Month')
if ($savIdx -ge 0) {
    $col = Get-ExcelColumn ([int]($savIdx + 1))
    $range = "${col}2:${col}${recLast}"
    Add-ConditionalFormatting -Worksheet $wsRec -Address $range -DataBarColor ([System.Drawing.Color]::Green)
}

# ── Sheet 3: Top Savings Opportunities ───────────────────────────────────────
if ($topArr.Count -gt 0) {
    $pkg = $topArr | Export-Excel `
        -ExcelPackage $pkg `
        -WorksheetName 'Top Savings' `
        -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
        -TableName 'TopSavings' -TableStyle 'Medium3' `
        -PassThru

    $wsTop  = $pkg.Workbook.Worksheets['Top Savings']
    $topH   = @($topArr[0].PSObject.Properties.Name)
    $topLast = $topArr.Count + 1

    foreach ($col in @('CurrentPrice_Hr','CurrentCost_Month','SuggestedPrice_Hr','SuggestedCost_Month','EstSaving_Month','EstSaving_Annual')) {
        $i = [array]::IndexOf($topH, $col)
        if ($i -ge 0) { $wsTop.Column([int]($i+1)).Style.Numberformat.Format = '$#,##0.00' }
    }

    $savIdx2 = [array]::IndexOf($topH, 'EstSaving_Month')
    if ($savIdx2 -ge 0) {
        $col = Get-ExcelColumn ([int]($savIdx2 + 1))
        $range = "${col}2:${col}${topLast}"
        Add-ConditionalFormatting -Worksheet $wsTop -Address $range -DataBarColor ([System.Drawing.Color]::DarkGreen)
    }
}

# ── Sheet 4: Idle VMs ────────────────────────────────────────────────────────
$idleVMs = @($reportArr | Where-Object { $_.Recommendation -match 'Idle' })
if ($idleVMs.Count -gt 0) {
    $pkg = $idleVMs | Export-Excel `
        -ExcelPackage $pkg `
        -WorksheetName 'Idle VMs' `
        -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
        -TableName 'IdleVMs' -TableStyle 'Medium10' `
        -PassThru
} else {
    $ws = Add-Worksheet -ExcelPackage $pkg -WorksheetName 'Idle VMs'
    $ws.Cells['A1'].Value = 'No idle VMs found based on current thresholds.'
}

# ── Sheet 5: Overused VMs ─────────────────────────────────────────────────────
$overVMs = @($reportArr | Where-Object { $_.Recommendation -match 'Overused' })
if ($overVMs.Count -gt 0) {
    $pkg = $overVMs | Export-Excel `
        -ExcelPackage $pkg `
        -WorksheetName 'Overused VMs' `
        -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
        -TableName 'OverusedVMs' -TableStyle 'Medium2' `
        -PassThru
} else {
    $ws = Add-Worksheet -ExcelPackage $pkg -WorksheetName 'Overused VMs'
    $ws.Cells['A1'].Value = 'No overused VMs found based on current thresholds.'
}

Close-ExcelPackage $pkg

Write-Host "[DONE] Report: $OutputPath" -ForegroundColor Green

$totalSaving = [math]::Round(($reportArr | Measure-Object EstSaving_Month -Sum).Sum, 2)
$totalAnnual = [math]::Round($totalSaving * 12, 2)
Write-Host ""
Write-Host "  VMs analysed        : $($reportArr.Count)"                           -ForegroundColor White
Write-Host "  Idle VMs            : $(@($reportArr | Where-Object { $_.Recommendation -match 'Idle' }).Count)"      -ForegroundColor Red
Write-Host "  Underused VMs       : $(@($reportArr | Where-Object { $_.Recommendation -match 'Underused' }).Count)" -ForegroundColor Yellow
Write-Host "  Right-Sized VMs     : $(@($reportArr | Where-Object { $_.Recommendation -eq 'Right-Sized' }).Count)"  -ForegroundColor Green
Write-Host "  Overused VMs        : $(@($reportArr | Where-Object { $_.Recommendation -match 'Overused' }).Count)"  -ForegroundColor DarkYellow
Write-Host "  Est. Monthly Saving : `$$totalSaving"                                 -ForegroundColor Cyan
Write-Host "  Est. Annual Saving  : `$$totalAnnual"                                 -ForegroundColor Cyan
Write-Host ""

#endregion
