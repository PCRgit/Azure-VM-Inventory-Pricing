
<#
.SYNOPSIS
    Azure VM Cost Excel Dashboard Generator

.DESCRIPTION
    Takes the VM inventory data collected by AzureVM-Inventory-Pricing.ps1 and
    builds a rich Excel dashboard with:
      - Sheet 1: Executive Dashboard  (KPI tiles, cost breakdown table, SKU summary)
      - Sheet 2: VM Inventory         (full VM details with conditional formatting)
      - Sheet 3: Disk Inventory       (all disks with orphan detection)
      - Sheet 4: Cost by Subscription (bar-style cost summary)
      - Sheet 5: Cost by Region       (region cost breakdown)
      - Sheet 6: Cost by VM Size      (top SKUs by spend)
      - Sheet 7: Deallocated VMs      (waste/savings opportunity list)

    This script can be run standalone by passing pre-collected $allVMs and $allDisks
    arrays, OR dot-sourced after AzureVM-Inventory-Pricing.ps1.

.NOTES
    Requires: ImportExcel module
      Install-Module ImportExcel -Scope CurrentUser -Force
#>

param(
    [Parameter(Mandatory = $true)]
    [System.Collections.Generic.List[object]]$AllVMs,

    [Parameter(Mandatory = $true)]
    [System.Collections.Generic.List[object]]$AllDisks,

    [string]$OutputPath = "$env:USERPROFILE\Desktop\Azure_Cost_Dashboard_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx",
    [string]$CurrencyCode = "USD"
)

#Requires -Modules ImportExcel

Add-Type -AssemblyName System.Drawing

# ─────────────────────────────────────────────────────────────────────────────
# Helper: Get Excel column letter from number
# ─────────────────────────────────────────────────────────────────────────────
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

# ─────────────────────────────────────────────────────────────────────────────
# Helper: Write a styled KPI tile into a worksheet cell range
# ─────────────────────────────────────────────────────────────────────────────
function Write-KpiTile {
    param(
        $Worksheet,
        [int]$Row,
        [int]$Col,
        [string]$Label,
        [string]$Value,
        [System.Drawing.Color]$BgColor,
        [System.Drawing.Color]$TextColor = [System.Drawing.Color]::White,
        [int]$MergeWidth  = 2,
        [int]$MergeHeight = 3
    )

    # Write label row
    $labelCell = $Worksheet.Cells[$Row, $Col]
    $labelCell.Value = $Label
    $labelCell.Style.Font.Bold = $true
    $labelCell.Style.Font.Size = 10
    $labelCell.Style.Font.Color.SetColor($TextColor)
    $labelCell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $labelCell.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    $labelCell.Style.Fill.PatternType    = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $labelCell.Style.Fill.BackgroundColor.SetColor($BgColor)

    # Write value row
    $valueCell = $Worksheet.Cells[$Row + 1, $Col]
    $valueCell.Value = $Value
    $valueCell.Style.Font.Bold = $true
    $valueCell.Style.Font.Size = 18
    $valueCell.Style.Font.Color.SetColor($TextColor)
    $valueCell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $valueCell.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    $valueCell.Style.Fill.PatternType    = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $valueCell.Style.Fill.BackgroundColor.SetColor($BgColor)

    # Spacer row
    $spacerCell = $Worksheet.Cells[$Row + 2, $Col]
    $spacerCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $spacerCell.Style.Fill.BackgroundColor.SetColor($BgColor)

    # Merge if spanning multiple columns
    if ($MergeWidth -gt 1) {
        $endCol = $Col + $MergeWidth - 1
        for ($r = $Row; $r -le $Row + $MergeHeight - 1; $r++) {
            $Worksheet.Cells[$r, $Col, $r, $endCol].Merge = $true
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Helper: Style a header row
# ─────────────────────────────────────────────────────────────────────────────
function Set-HeaderStyle {
    param($Worksheet, [int]$Row, [int]$ColCount, [System.Drawing.Color]$BgColor)
    $range = $Worksheet.Cells[$Row, 1, $Row, $ColCount]
    $range.Style.Font.Bold = $true
    $range.Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $range.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $range.Style.Fill.BackgroundColor.SetColor($BgColor)
    $range.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $range.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
    $range.Style.Border.Bottom.Color.SetColor([System.Drawing.Color]::White)
}

# ─────────────────────────────────────────────────────────────────────────────
# Compute aggregate data
# ─────────────────────────────────────────────────────────────────────────────
$vms   = @($AllVMs)
$disks = @($AllDisks)

$totalVMs           = $vms.Count
$runningVMs         = @($vms | Where-Object { $_.PowerState -match 'running' }).Count
$deallocatedVMs     = @($vms | Where-Object { $_.PowerState -match 'deallocated' }).Count
$stoppedVMs         = @($vms | Where-Object { $_.PowerState -match '^VM stopped$' }).Count
$unknownVMs         = @($vms | Where-Object { $_.PowerState -eq 'Unknown' }).Count
$totalDisks         = $disks.Count
$unattachedDisks    = @($disks | Where-Object { $_.DiskState -eq 'Unattached' }).Count

$vmCostMeasure      = $vms   | Measure-Object -Property VMMonthlyPrice_USD   -Sum
$diskCostMeasure    = $disks | Measure-Object -Property EstMonthlyPrice_USD  -Sum
$totalCostMeasure   = $vms   | Measure-Object -Property EstTotalMonthlyCost_USD -Sum

$totalVMCost     = if ($vmCostMeasure.Sum)    { [math]::Round([double]$vmCostMeasure.Sum,   2) } else { 0 }
$totalDiskCost   = if ($diskCostMeasure.Sum)  { [math]::Round([double]$diskCostMeasure.Sum, 2) } else { 0 }
$totalCombined   = if ($totalCostMeasure.Sum) { [math]::Round([double]$totalCostMeasure.Sum,2) } else { 0 }

# Deallocated VM waste (they still incur disk cost)
$deallocatedDiskCost = [math]::Round(
    ([double](($disks | Where-Object { $vmName = $_.VMName; @($vms | Where-Object { $_.VMName -eq $vmName -and $_.PowerState -match 'deallocated' }).Count -gt 0 } |
        Measure-Object -Property EstMonthlyPrice_USD -Sum).Sum ?? 0)), 2)

# ─── Cost by Subscription ────────────────────────────────────────────────────
$costBySub = $vms | Group-Object SubscriptionName | ForEach-Object {
    $g     = @($_.Group)
    $sDisks = @($disks | Where-Object { $_.SubscriptionName -eq $_.Name })
    $vmC   = $g | Measure-Object -Property VMMonthlyPrice_USD -Sum
    $dkC   = $sDisks | Measure-Object -Property EstMonthlyPrice_USD -Sum
    [PSCustomObject]@{
        Subscription      = $_.Name
        VMCount           = $g.Count
        RunningVMs        = @($g | Where-Object { $_.PowerState -match 'running' }).Count
        DeallocatedVMs    = @($g | Where-Object { $_.PowerState -match 'deallocated' }).Count
        VMCompute_USD     = [math]::Round($(if ($vmC.Sum) { [double]$vmC.Sum } else { 0 }), 2)
        DiskCost_USD      = [math]::Round($(if ($dkC.Sum) { [double]$dkC.Sum } else { 0 }), 2)
        Total_Monthly_USD = [math]::Round($(if ($vmC.Sum) { [double]$vmC.Sum } else { 0 }) + $(if ($dkC.Sum) { [double]$dkC.Sum } else { 0 }), 2)
    }
} | Sort-Object Total_Monthly_USD -Descending

# ─── Cost by Region ──────────────────────────────────────────────────────────
$costByRegion = $vms | Group-Object Location | ForEach-Object {
    $g  = @($_.Group)
    $vmC = $g | Measure-Object -Property VMMonthlyPrice_USD -Sum
    $dkC = @($disks | Where-Object { $_.Region -eq $_.Name }) | Measure-Object -Property EstMonthlyPrice_USD -Sum
    [PSCustomObject]@{
        Region            = $_.Name
        VMCount           = $g.Count
        VMCompute_USD     = [math]::Round($(if ($vmC.Sum) { [double]$vmC.Sum } else { 0 }), 2)
        DiskCost_USD      = [math]::Round($(if ($dkC.Sum) { [double]$dkC.Sum } else { 0 }), 2)
        Total_Monthly_USD = [math]::Round($(if ($vmC.Sum) { [double]$vmC.Sum } else { 0 }) + $(if ($dkC.Sum) { [double]$dkC.Sum } else { 0 }), 2)
    }
} | Sort-Object Total_Monthly_USD -Descending

# ─── Cost by VM Size ─────────────────────────────────────────────────────────
$costBySize = $vms | Group-Object VMSize | ForEach-Object {
    $g   = @($_.Group)
    $vmC = $g | Measure-Object -Property VMMonthlyPrice_USD -Sum
    $hr  = $g[0].VMHourlyPrice_USD
    [PSCustomObject]@{
        VMSize              = $_.Name
        VMCount             = $g.Count
        UnitHourlyPrice_USD = $hr
        Total_Monthly_USD   = [math]::Round($(if ($vmC.Sum) { [double]$vmC.Sum } else { 0 }), 2)
        RunningCount        = @($g | Where-Object { $_.PowerState -match 'running' }).Count
    }
} | Sort-Object Total_Monthly_USD -Descending

# ─── Deallocated VMs (savings opportunity) ───────────────────────────────────
$deallocList = @($vms | Where-Object { $_.PowerState -match 'deallocated' } |
    Select-Object SubscriptionName, ResourceGroup, VMName, VMSize, Location,
                  OSDiskSKU, OSDiskSizeGB, DataDiskCount,
                  VMMonthlyPrice_USD, EstTotalDiskCost_USD, EstTotalMonthlyCost_USD,
                  Tags |
    Sort-Object EstTotalMonthlyCost_USD -Descending)

# ─────────────────────────────────────────────────────────────────────────────
# Build Excel package
# ─────────────────────────────────────────────────────────────────────────────
Write-Host "`n[INFO] Building Excel dashboard: $OutputPath" -ForegroundColor Cyan

# Use EPPlus directly via ImportExcel's underlying package
$excel = [OfficeOpenXml.ExcelPackage]::new([System.IO.FileInfo]$OutputPath)

# Color palette
$colorDarkBlue    = [System.Drawing.Color]::FromArgb(31,  73, 125)
$colorMidBlue     = [System.Drawing.Color]::FromArgb(68, 114, 196)
$colorLightBlue   = [System.Drawing.Color]::FromArgb(189, 215, 238)
$colorGreen       = [System.Drawing.Color]::FromArgb(70,  130, 80)
$colorOrange      = [System.Drawing.Color]::FromArgb(197, 90,  17)
$colorRed         = [System.Drawing.Color]::FromArgb(192, 0,   0)
$colorGray        = [System.Drawing.Color]::FromArgb(89,  89,  89)
$colorYellow      = [System.Drawing.Color]::FromArgb(255, 230, 153)
$colorLightGray   = [System.Drawing.Color]::FromArgb(242, 242, 242)
$colorWhite       = [System.Drawing.Color]::White

# ═════════════════════════════════════════════════════════════════════════════
# SHEET 1: Executive Dashboard
# ═════════════════════════════════════════════════════════════════════════════
$wsDash = $excel.Workbook.Worksheets.Add("Dashboard")
$wsDash.View.ShowGridLines = $false
$wsDash.TabColor = $colorDarkBlue

# ─── Title Banner ─────────────────────────────────────────────────────────────
$titleRange = $wsDash.Cells["A1:P2"]
$titleRange.Merge = $true
$titleRange.Value = "AZURE VM COST DASHBOARD"
$titleRange.Style.Font.Size = 22
$titleRange.Style.Font.Bold = $true
$titleRange.Style.Font.Color.SetColor($colorWhite)
$titleRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$titleRange.Style.Fill.BackgroundColor.SetColor($colorDarkBlue)
$titleRange.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
$titleRange.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

$subTitleRange = $wsDash.Cells["A3:P3"]
$subTitleRange.Merge = $true
$subTitleRange.Value = "Generated: $(Get-Date -Format 'dddd, MMMM dd yyyy  HH:mm')  |  Currency: $CurrencyCode  |  Retail (PAYG) pricing"
$subTitleRange.Style.Font.Size = 10
$subTitleRange.Style.Font.Italic = $true
$subTitleRange.Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(166,166,166))
$subTitleRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$subTitleRange.Style.Fill.BackgroundColor.SetColor($colorDarkBlue)
$subTitleRange.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center

# ─── KPI Tiles Row 1 (row 5-7) ────────────────────────────────────────────────
# Tile layout: col 1-2, 4-5, 7-8, 10-11, 13-14, 16-17
$kpiData = @(
    @{ Label="TOTAL VMs";          Value=$totalVMs.ToString();                    Color=$colorMidBlue  }
    @{ Label="RUNNING";            Value=$runningVMs.ToString();                  Color=$colorGreen    }
    @{ Label="DEALLOCATED";        Value=$deallocatedVMs.ToString();              Color=$colorOrange   }
    @{ Label="TOTAL DISKS";        Value=$totalDisks.ToString();                  Color=$colorGray     }
    @{ Label="UNATTACHED DISKS";   Value=$unattachedDisks.ToString();             Color=$colorRed      }
    @{ Label="SUBSCRIPTIONS";      Value=($vms | Select-Object -Unique SubscriptionName).Count.ToString(); Color=$colorDarkBlue }
)

$kpiCols = @(1, 4, 7, 10, 13, 16)
for ($k = 0; $k -lt $kpiData.Count; $k++) {
    $kd  = $kpiData[$k]
    $col = $kpiCols[$k]
    Write-KpiTile -Worksheet $wsDash -Row 5 -Col $col `
        -Label $kd.Label -Value $kd.Value -BgColor $kd.Color
}

# ─── KPI Tiles Row 2 — Cost KPIs (row 9-11) ────────────────────────────────
$costKpis = @(
    @{ Label="VM COMPUTE / MONTH";    Value="`$$("{0:N2}" -f $totalVMCost)";    Color=$colorMidBlue  }
    @{ Label="DISK COST / MONTH";     Value="`$$("{0:N2}" -f $totalDiskCost)";  Color=$colorGray     }
    @{ Label="TOTAL COST / MONTH";    Value="`$$("{0:N2}" -f $totalCombined)";  Color=$colorDarkBlue }
    @{ Label="DEALLOCATED DISK WASTE";Value="`$$("{0:N2}" -f $deallocatedDiskCost)"; Color=$colorOrange }
    @{ Label="ANNUAL ESTIMATE";       Value="`$$("{0:N0}" -f ($totalCombined * 12))"; Color=$colorGreen }
)

$costKpiCols = @(1, 4, 7, 10, 13)
for ($k = 0; $k -lt $costKpis.Count; $k++) {
    $kd  = $costKpis[$k]
    $col = $costKpiCols[$k]
    Write-KpiTile -Worksheet $wsDash -Row 9 -Col $col `
        -Label $kd.Label -Value $kd.Value -BgColor $kd.Color
}

# ─── Section: Cost by Subscription ──────────────────────────────────────────
$secRow = 14

# Section header
$hRange = $wsDash.Cells[$secRow, 1, $secRow, 7]
$hRange.Merge = $true
$hRange.Value = "COST BY SUBSCRIPTION"
$hRange.Style.Font.Bold = $true
$hRange.Style.Font.Size = 11
$hRange.Style.Font.Color.SetColor($colorWhite)
$hRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$hRange.Style.Fill.BackgroundColor.SetColor($colorDarkBlue)
$hRange.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
$hRange.Style.Indent = 1

# Table headers
$subHeaders = @("Subscription","VMs","Running","Deallocated","VM Compute ($CurrencyCode)","Disk Cost ($CurrencyCode)","Total/Month ($CurrencyCode)")
for ($c = 0; $c -lt $subHeaders.Count; $c++) {
    $cell = $wsDash.Cells[$secRow + 1, $c + 1]
    $cell.Value = $subHeaders[$c]
    $cell.Style.Font.Bold = $true
    $cell.Style.Font.Color.SetColor($colorWhite)
    $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $cell.Style.Fill.BackgroundColor.SetColor($colorMidBlue)
    $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $cell.Style.Border.BorderAround([OfficeOpenXml.Style.ExcelBorderStyle]::Thin, $colorWhite)
}

$dataRow = $secRow + 2
$rowToggle = $false
foreach ($row in $costBySub) {
    $rowBg = if ($rowToggle) { $colorLightGray } else { $colorWhite }
    $rowToggle = -not $rowToggle

    $vals = @($row.Subscription, $row.VMCount, $row.RunningVMs, $row.DeallocatedVMs,
              $row.VMCompute_USD, $row.DiskCost_USD, $row.Total_Monthly_USD)

    for ($c = 0; $c -lt $vals.Count; $c++) {
        $cell = $wsDash.Cells[$dataRow, $c + 1]
        $cell.Value = $vals[$c]
        $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $cell.Style.Fill.BackgroundColor.SetColor($rowBg)
        $cell.Style.Border.BorderAround([OfficeOpenXml.Style.ExcelBorderStyle]::Hair)

        if ($c -ge 4) {
            $cell.Style.Numberformat.Format = '$#,##0.00'
            $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Right
        } else {
            $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
        }
    }

    # Highlight total column with data bar style (manual orange for high cost)
    $totalCell = $wsDash.Cells[$dataRow, 7]
    if ($row.Total_Monthly_USD -gt 1000) {
        $totalCell.Style.Font.Bold = $true
        $totalCell.Style.Font.Color.SetColor($colorRed)
    }

    $dataRow++
}

# Grand Total row
$totalLabelCell = $wsDash.Cells[$dataRow, 1]
$totalLabelCell.Value = "GRAND TOTAL"
$totalLabelCell.Style.Font.Bold = $true
$totalLabelCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$totalLabelCell.Style.Fill.BackgroundColor.SetColor($colorLightBlue)

$grandTotals = @("", $totalVMs, $runningVMs, $deallocatedVMs, $totalVMCost, $totalDiskCost, $totalCombined)
for ($c = 1; $c -lt $grandTotals.Count; $c++) {
    $cell = $wsDash.Cells[$dataRow, $c + 1]
    $cell.Value = $grandTotals[$c]
    $cell.Style.Font.Bold = $true
    $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $cell.Style.Fill.BackgroundColor.SetColor($colorLightBlue)
    if ($c -ge 4) { $cell.Style.Numberformat.Format = '$#,##0.00' }
    $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
}

# ─── Section: Top 10 Most Expensive VMs ─────────────────────────────────────
$secRow2 = $dataRow + 2

$h2Range = $wsDash.Cells[$secRow2, 1, $secRow2, 7]
$h2Range.Merge = $true
$h2Range.Value = "TOP 10 MOST EXPENSIVE VMs"
$h2Range.Style.Font.Bold = $true
$h2Range.Style.Font.Size = 11
$h2Range.Style.Font.Color.SetColor($colorWhite)
$h2Range.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$h2Range.Style.Fill.BackgroundColor.SetColor($colorRed)
$h2Range.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
$h2Range.Style.Indent = 1

$top10Headers = @("VM Name","Subscription","Location","VM Size","VM/Month","Disk/Month","Total/Month")
for ($c = 0; $c -lt $top10Headers.Count; $c++) {
    $cell = $wsDash.Cells[$secRow2 + 1, $c + 1]
    $cell.Value = $top10Headers[$c]
    $cell.Style.Font.Bold = $true
    $cell.Style.Font.Color.SetColor($colorWhite)
    $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $cell.Style.Fill.BackgroundColor.SetColor($colorOrange)
    $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
}

$top10VMs = @($vms | Sort-Object EstTotalMonthlyCost_USD -Descending | Select-Object -First 10)
$t10Row = $secRow2 + 2
$toggle = $false
foreach ($v in $top10VMs) {
    $bg = if ($toggle) { $colorLightGray } else { $colorWhite }
    $toggle = -not $toggle
    $t10Vals = @($v.VMName, $v.SubscriptionName, $v.Location, $v.VMSize,
                 $v.VMMonthlyPrice_USD, $v.EstTotalDiskCost_USD, $v.EstTotalMonthlyCost_USD)
    for ($c = 0; $c -lt $t10Vals.Count; $c++) {
        $cell = $wsDash.Cells[$t10Row, $c + 1]
        $cell.Value = $t10Vals[$c]
        $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $cell.Style.Fill.BackgroundColor.SetColor($bg)
        $cell.Style.Border.BorderAround([OfficeOpenXml.Style.ExcelBorderStyle]::Hair)
        if ($c -ge 4) { $cell.Style.Numberformat.Format = '$#,##0.00' }
    }
    $t10Row++
}

# ─── Section: Top VM Sizes by Spend ─────────────────────────────────────────
$secRow3 = $t10Row + 2
$colOffset = 9  # place this section in columns 9-16

$h3Range = $wsDash.Cells[$secRow3, $colOffset, $secRow3, $colOffset + 6]
$h3Range.Merge = $true
$h3Range.Value = "TOP VM SIZES BY MONTHLY SPEND"
$h3Range.Style.Font.Bold = $true
$h3Range.Style.Font.Size = 11
$h3Range.Style.Font.Color.SetColor($colorWhite)
$h3Range.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$h3Range.Style.Fill.BackgroundColor.SetColor($colorGray)
$h3Range.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
$h3Range.Style.Indent = 1

$sizeHeaders = @("VM Size","Count","Running","Unit Price/hr","Total/Month","% of Compute")
for ($c = 0; $c -lt $sizeHeaders.Count; $c++) {
    $cell = $wsDash.Cells[$secRow3 + 1, $colOffset + $c]
    $cell.Value = $sizeHeaders[$c]
    $cell.Style.Font.Bold = $true
    $cell.Style.Font.Color.SetColor($colorWhite)
    $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $cell.Style.Fill.BackgroundColor.SetColor($colorGray)
    $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
}

$sizeRow = $secRow3 + 2
$toggle2 = $false
$topSizes = @($costBySize | Select-Object -First 15)
foreach ($sz in $topSizes) {
    $bg = if ($toggle2) { $colorLightGray } else { $colorWhite }
    $toggle2 = -not $toggle2
    $pct = if ($totalVMCost -gt 0) { [math]::Round($sz.Total_Monthly_USD / $totalVMCost * 100, 1) } else { 0 }

    $szVals = @($sz.VMSize, $sz.VMCount, $sz.RunningCount, $sz.UnitHourlyPrice_USD, $sz.Total_Monthly_USD, "$pct%")
    for ($c = 0; $c -lt $szVals.Count; $c++) {
        $cell = $wsDash.Cells[$sizeRow, $colOffset + $c]
        $cell.Value = $szVals[$c]
        $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $cell.Style.Fill.BackgroundColor.SetColor($bg)
        $cell.Style.Border.BorderAround([OfficeOpenXml.Style.ExcelBorderStyle]::Hair)
        if ($c -eq 3 -or $c -eq 4) { $cell.Style.Numberformat.Format = '$#,##0.0000' }
        $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    }
    $sizeRow++
}

# ─── Auto column width on dashboard ──────────────────────────────────────────
$wsDash.Cells[$wsDash.Dimension.Address].AutoFitColumns()
# Ensure row heights are visible for KPI tiles
5..12 | ForEach-Object { $wsDash.Row($_).Height = 24 }

# ═════════════════════════════════════════════════════════════════════════════
# SHEET 2: VM Inventory (full detail)
# ═════════════════════════════════════════════════════════════════════════════
$excelPkg = $AllVMs | Export-Excel `
    -ExcelPackage  $excel `
    -WorksheetName 'VM Inventory' `
    -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
    -TableName     'VMInventory' `
    -TableStyle    'Medium9' `
    -PassThru

$wsVM = $excelPkg.Workbook.Worksheets['VM Inventory']
if ($vms.Count -gt 0) {
    $headers = @($vms[0].PSObject.Properties.Name)
    $lastRow = $vms.Count + 1

    # Power State conditional formatting
    $pwrIdx = [array]::IndexOf($headers, 'PowerState')
    if ($pwrIdx -ge 0) {
        $col   = Get-ExcelColumn -Col ($pwrIdx + 1)
        $range = "${col}2:${col}${lastRow}"
        Add-ConditionalFormatting -Worksheet $wsVM -Address $range -RuleType ContainsText -ConditionValue 'running'     -BackgroundColor ([System.Drawing.Color]::LightGreen)
        Add-ConditionalFormatting -Worksheet $wsVM -Address $range -RuleType ContainsText -ConditionValue 'deallocated' -BackgroundColor ([System.Drawing.Color]::LightCoral)
        Add-ConditionalFormatting -Worksheet $wsVM -Address $range -RuleType ContainsText -ConditionValue 'stopped'     -BackgroundColor ([System.Drawing.Color]::LightYellow)
        Add-ConditionalFormatting -Worksheet $wsVM -Address $range -RuleType ContainsText -ConditionValue 'Unknown'     -BackgroundColor ([System.Drawing.Color]::LightGray)
    }

    # Currency formatting for price columns
    foreach ($pCol in @('VMHourlyPrice_USD','VMMonthlyPrice_USD','EstTotalDiskCost_USD','EstTotalMonthlyCost_USD')) {
        $idx = [array]::IndexOf($headers, $pCol)
        if ($idx -ge 0) { $wsVM.Column($idx + 1).Style.Numberformat.Format = '$#,##0.0000' }
    }

    # High cost highlight (total > $1000/month)
    $tcIdx = [array]::IndexOf($headers, 'EstTotalMonthlyCost_USD')
    if ($tcIdx -ge 0) {
        $col   = Get-ExcelColumn -Col ($tcIdx + 1)
        $range = "${col}2:${col}${lastRow}"
        Add-ConditionalFormatting -Worksheet $wsVM -Address $range -RuleType GreaterThan -ConditionValue 1000 -BackgroundColor ([System.Drawing.Color]::FromArgb(255, 200, 100))
    }
}

# ═════════════════════════════════════════════════════════════════════════════
# SHEET 3: Disk Inventory
# ═════════════════════════════════════════════════════════════════════════════
$excelPkg = $AllDisks | Export-Excel `
    -ExcelPackage  $excelPkg `
    -WorksheetName 'Disk Inventory' `
    -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
    -TableName     'DiskInventory' `
    -TableStyle    'Medium6' `
    -PassThru

$wsDisk = $excelPkg.Workbook.Worksheets['Disk Inventory']
if ($disks.Count -gt 0) {
    $dHeaders = @($disks[0].PSObject.Properties.Name)
    $dLastRow = $disks.Count + 1

    $dtIdx = [array]::IndexOf($dHeaders, 'DiskType')
    if ($dtIdx -ge 0) {
        $col   = Get-ExcelColumn -Col ($dtIdx + 1)
        $range = "${col}2:${col}${dLastRow}"
        Add-ConditionalFormatting -Worksheet $wsDisk -Address $range -RuleType ContainsText -ConditionValue 'OS Disk'   -BackgroundColor ([System.Drawing.Color]::LightSteelBlue)
        Add-ConditionalFormatting -Worksheet $wsDisk -Address $range -RuleType ContainsText -ConditionValue 'Data Disk' -BackgroundColor ([System.Drawing.Color]::LightCyan)
    }

    $dsIdx = [array]::IndexOf($dHeaders, 'DiskState')
    if ($dsIdx -ge 0) {
        $col   = Get-ExcelColumn -Col ($dsIdx + 1)
        $range = "${col}2:${col}${dLastRow}"
        Add-ConditionalFormatting -Worksheet $wsDisk -Address $range -RuleType ContainsText -ConditionValue 'Unattached' -BackgroundColor ([System.Drawing.Color]::LightCoral)
    }

    $dpIdx = [array]::IndexOf($dHeaders, 'EstMonthlyPrice_USD')
    if ($dpIdx -ge 0) { $wsDisk.Column($dpIdx + 1).Style.Numberformat.Format = '$#,##0.00' }
}

# ═════════════════════════════════════════════════════════════════════════════
# SHEET 4: Cost by Subscription
# ═════════════════════════════════════════════════════════════════════════════
$excelPkg = $costBySub | Export-Excel `
    -ExcelPackage  $excelPkg `
    -WorksheetName 'Cost by Subscription' `
    -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
    -TableName     'CostBySub' `
    -TableStyle    'Medium2' `
    -PassThru

$wsCBS = $excelPkg.Workbook.Worksheets['Cost by Subscription']
foreach ($col in @('VMCompute_USD','DiskCost_USD','Total_Monthly_USD')) {
    $idx = [array]::IndexOf(@($costBySub[0].PSObject.Properties.Name), $col)
    if ($idx -ge 0) { $wsCBS.Column($idx + 1).Style.Numberformat.Format = '$#,##0.00' }
}
if (@($costBySub).Count -gt 0) {
    $tcIdx2 = [array]::IndexOf(@($costBySub[0].PSObject.Properties.Name), 'Total_Monthly_USD')
    if ($tcIdx2 -ge 0) {
        $col   = Get-ExcelColumn -Col ($tcIdx2 + 1)
        $range = "${col}2:${col}$(@($costBySub).Count + 1)"
        Add-ConditionalFormatting -Worksheet $wsCBS -Address $range -DataBarColor $colorMidBlue
    }
}

# ═════════════════════════════════════════════════════════════════════════════
# SHEET 5: Cost by Region
# ═════════════════════════════════════════════════════════════════════════════
$excelPkg = $costByRegion | Export-Excel `
    -ExcelPackage  $excelPkg `
    -WorksheetName 'Cost by Region' `
    -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
    -TableName     'CostByRegion' `
    -TableStyle    'Medium4' `
    -PassThru

$wsCBR = $excelPkg.Workbook.Worksheets['Cost by Region']
foreach ($col in @('VMCompute_USD','DiskCost_USD','Total_Monthly_USD')) {
    $idx = [array]::IndexOf(@($costByRegion[0].PSObject.Properties.Name), $col)
    if ($idx -ge 0) { $wsCBR.Column($idx + 1).Style.Numberformat.Format = '$#,##0.00' }
}
if (@($costByRegion).Count -gt 0) {
    $tcIdxR = [array]::IndexOf(@($costByRegion[0].PSObject.Properties.Name), 'Total_Monthly_USD')
    if ($tcIdxR -ge 0) {
        $col   = Get-ExcelColumn -Col ($tcIdxR + 1)
        $range = "${col}2:${col}$(@($costByRegion).Count + 1)"
        Add-ConditionalFormatting -Worksheet $wsCBR -Address $range -DataBarColor $colorGreen
    }
}

# ═════════════════════════════════════════════════════════════════════════════
# SHEET 6: Cost by VM Size
# ═════════════════════════════════════════════════════════════════════════════
$excelPkg = $costBySize | Export-Excel `
    -ExcelPackage  $excelPkg `
    -WorksheetName 'Cost by VM Size' `
    -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
    -TableName     'CostBySize' `
    -TableStyle    'Medium5' `
    -PassThru

$wsCBSZ = $excelPkg.Workbook.Worksheets['Cost by VM Size']
foreach ($col in @('UnitHourlyPrice_USD','Total_Monthly_USD')) {
    $idx = [array]::IndexOf(@($costBySize[0].PSObject.Properties.Name), $col)
    if ($idx -ge 0) { $wsCBSZ.Column($idx + 1).Style.Numberformat.Format = '$#,##0.0000' }
}
if (@($costBySize).Count -gt 0) {
    $tcIdxSZ = [array]::IndexOf(@($costBySize[0].PSObject.Properties.Name), 'Total_Monthly_USD')
    if ($tcIdxSZ -ge 0) {
        $col   = Get-ExcelColumn -Col ($tcIdxSZ + 1)
        $range = "${col}2:${col}$(@($costBySize).Count + 1)"
        Add-ConditionalFormatting -Worksheet $wsCBSZ -Address $range -DataBarColor $colorOrange
    }
}

# ═════════════════════════════════════════════════════════════════════════════
# SHEET 7: Deallocated VMs (Savings Opportunity)
# ═════════════════════════════════════════════════════════════════════════════
if ($deallocList.Count -gt 0) {
    $excelPkg = $deallocList | Export-Excel `
        -ExcelPackage  $excelPkg `
        -WorksheetName 'Savings Opportunities' `
        -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
        -TableName     'DeallocVMs' `
        -TableStyle    'Medium3' `
        -PassThru

    $wsSav = $excelPkg.Workbook.Worksheets['Savings Opportunities']
    $savHeaders = @($deallocList[0].PSObject.Properties.Name)

    foreach ($col in @('VMMonthlyPrice_USD','EstTotalDiskCost_USD','EstTotalMonthlyCost_USD')) {
        $idx = [array]::IndexOf($savHeaders, $col)
        if ($idx -ge 0) { $wsSav.Column($idx + 1).Style.Numberformat.Format = '$#,##0.00' }
    }

    # Banner note above table
    $noteRow = $deallocList.Count + 3
    $noteRange = $wsSav.Cells[$noteRow, 1, $noteRow, 5]
    $noteRange.Merge = $true
    $noteRange.Value = "⚠ These VMs are deallocated and not billing for compute — but their disks ARE still incurring charges. Consider deleting unused VMs and their disks to eliminate waste."
    $noteRange.Style.Font.Italic = $true
    $noteRange.Style.Font.Color.SetColor($colorOrange)
    $noteRange.Style.WrapText = $true
    $wsSav.Row($noteRow).Height = 40
}

# ─────────────────────────────────────────────────────────────────────────────
# Reorder sheets — Dashboard first
# ─────────────────────────────────────────────────────────────────────────────
$excelPkg.Workbook.Worksheets.MoveToStart("Dashboard")

# Set tab colors
$tabColors = @{
    "Dashboard"            = $colorDarkBlue
    "VM Inventory"         = $colorMidBlue
    "Disk Inventory"       = $colorGray
    "Cost by Subscription" = $colorGreen
    "Cost by Region"       = [System.Drawing.Color]::FromArgb(112, 48, 160)
    "Cost by VM Size"      = $colorOrange
    "Savings Opportunities"= $colorRed
}

foreach ($kvp in $tabColors.GetEnumerator()) {
    $ws = $excelPkg.Workbook.Worksheets[$kvp.Key]
    if ($ws) { $ws.TabColor = $kvp.Value }
}

# ─────────────────────────────────────────────────────────────────────────────
# Save
# ─────────────────────────────────────────────────────────────────────────────
Close-ExcelPackage $excelPkg -SaveAs $OutputPath 2>$null
if (-not (Test-Path $OutputPath)) {
    $excelPkg.SaveAs([System.IO.FileInfo]$OutputPath)
}

Write-Host "`n[DONE] Dashboard saved: $OutputPath" -ForegroundColor Green
Write-Host "  Sheets: Dashboard | VM Inventory | Disk Inventory | Cost by Subscription | Cost by Region | Cost by VM Size | Savings Opportunities" -ForegroundColor White
