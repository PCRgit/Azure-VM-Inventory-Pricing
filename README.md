# Azure VM Inventory & Pricing Report

A PowerShell script that scans **all subscriptions** in your Azure tenant, collects comprehensive VM and disk configurations, performs live price lookups via the Azure Retail Prices API, and exports everything into a formatted multi-sheet Excel workbook.

---

## 📋 Table of Contents

- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [Parameters](#parameters)
- [Output](#output)
- [Excel Workbook Structure](#excel-workbook-structure)
- [Pricing Logic](#pricing-logic)
- [Permissions Required](#permissions-required)
- [Known Limitations](#known-limitations)
- [Troubleshooting](#troubleshooting)

---

## ✨ Features

- 🔍 **Multi-subscription scan** — iterates every enabled subscription in the tenant automatically
- 💻 **Full VM settings** — compute, OS, availability, license type, boot diagnostics, tags
- 🌐 **Networking details** — NICs, private/public IPs, VNet, subnet, NSG, accelerated networking
- 💾 **All disks** — OS disk and every attached data disk with SKU, size, caching, LUN, encryption
- 💰 **Live pricing** — VM hourly/monthly cost and per-disk monthly cost via Azure Retail Prices API
- 📊 **Formatted Excel output** — 3-sheet workbook with tables, auto-filters, frozen headers, and conditional formatting (no Excel install required)
- ⚡ **Price caching** — deduplicates API calls; same SKU+region is only queried once per run
- 🎨 **Power state color coding** — Running (green), Deallocated (red), Stopped (yellow)

---

## ✅ Prerequisites

| Requirement | Version | Install Command |
|---|---|---|
| PowerShell | 5.1+ or 7.x | [Download](https://aka.ms/powershell) |
| Az PowerShell Module | Latest | `Install-Module Az -Scope CurrentUser` |
| ImportExcel Module | Latest | `Install-Module ImportExcel -Scope CurrentUser` |
| Azure Authentication | — | `Connect-AzAccount` |

> **Note:** The Azure Retail Prices API (`prices.azure.com`) is publicly accessible and does **not** require authentication.

---

## 🚀 Installation

1. **Clone or download** this repository:
   ```powershell
   git clone https://github.com/your-org/azure-vm-inventory.git
   cd azure-vm-inventory
   ```

2. **Install required modules** (one-time):
   ```powershell
   Install-Module Az -Scope CurrentUser -Force
   Install-Module ImportExcel -Scope CurrentUser -Force
   ```

3. **Authenticate to Azure**:
   ```powershell
   Connect-AzAccount

   # For service principal / automation:
   Connect-AzAccount -ServicePrincipal -Tenant $TenantId `
       -Credential (Get-Credential) -ApplicationId $AppId
   ```

---

## 💻 Usage

### Basic — Output to Desktop
```powershell
.\AzureVM-Inventory.ps1
```

### Custom Output Path
```powershell
.\AzureVM-Inventory.ps1 -OutputPath "C:\Reports\VMReport.xlsx"
```

### Exclude Specific Subscriptions
```powershell
.\AzureVM-Inventory.ps1 -ExcludeSubscriptions @("sub-id-1", "sub-id-2")
```

### Full Example
```powershell
.\AzureVM-Inventory.ps1 `
    -OutputPath   "C:\Reports\AzureVMs_$(Get-Date -Format 'yyyyMMdd').xlsx" `
    -CurrencyCode "USD" `
    -ExcludeSubscriptions @("xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx")
```

---

## ⚙️ Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-OutputPath` | String | `Desktop\AzureVM_Inventory_<timestamp>.xlsx` | Full path for the output Excel file |
| `-CurrencyCode` | String | `USD` | Currency for pricing (e.g. `USD`, `EUR`, `GBP`) |
| `-ExcludeSubscriptions` | String[] | `@()` | Array of Subscription IDs to skip |

---

## 📁 Output

The script generates a single `.xlsx` file with **3 worksheets**:

### Sheet 1 — VM Inventory
Complete VM configuration per row. Key columns:

| Column Group | Columns |
|---|---|
| Identity | SubscriptionName, SubscriptionId, ResourceGroup, VMName, ComputerName, AdminUsername |
| Compute | VMSize, OSType, OSPublisher, OSOffer, OSSKU, OSVersion |
| Power | PowerState *(color-coded)* |
| Networking | NICNames, PrivateIPAddresses, PublicIPAddresses, VNetNames, SubnetNames, NSGNames, AcceleratedNetworking |
| Availability | AvailabilitySet, AvailabilityZone, VirtualMachineScaleSet, ProximityPlacementGroup |
| OS Disk | OSDiskName, OSDiskSizeGB, OSDiskSKU |
| Misc | LicenseType, BootDiagnosticsEnabled, DataDiskCount, Tags |
| **Pricing** | **VMHourlyPrice_USD, VMMonthlyPrice_USD, EstTotalDiskCost_USD** |

### Sheet 2 — Disk Inventory
One row per disk (OS + all data disks). Key columns:

| Column | Description |
|---|---|
| DiskType | `OS Disk` or `Data Disk` |
| DiskSKU | `Premium_LRS`, `StandardSSD_LRS`, `Standard_LRS`, `UltraSSD_LRS`, etc. |
| DiskSizeGB | Provisioned size in GB |
| LUN | Logical Unit Number (data disks only) |
| Caching | `ReadWrite`, `ReadOnly`, `None` |
| DiskState | `Attached`, `Unattached`, etc. |
| EncryptionType | `EncryptionAtRestWithPlatformKey`, `CustomerManaged`, etc. |
| **EstMonthlyPrice_USD** | **Live price lookup from Azure Retail API** |

### Sheet 3 — Cost Summary
Per-subscription rollup:

| Column | Description |
|---|---|
| TotalVMs | Total VM count |
| RunningVMs | VMs currently running |
| DeallocatedVMs | Deallocated VMs (not billed for compute) |
| TotalDataDisks | Sum of all data disks |
| TotalVMCost_Monthly_USD | Estimated monthly VM compute cost |
| TotalDiskCost_Monthly_USD | Estimated monthly disk cost |

---

## 💲 Pricing Logic

### VM Pricing
- Queries `prices.azure.com` using `armSkuName` (e.g. `Standard_D4s_v5`) and `armRegionName`
- Filters for `Consumption` (pay-as-you-go) type, excludes **Spot** and **Low Priority**
- Monthly estimate = **Hourly price × 730 hours**

### Disk Pricing
- Maps disk SKU + size (GB) to Azure pricing tier label:
  - **Premium SSD**: P1–P80
  - **Standard SSD**: E1–E80
  - **Standard HDD**: S4–S80
  - **Ultra Disk**: queried separately
- Queries correct `productName` per storage tier
- Returns the current monthly **retail price** per disk

### Price Cache
All pricing results are stored in `$script:PriceCache` (hashtable keyed by `SKU_Region`).
Identical SKU+Region combinations only trigger **one API call** per script run, significantly reducing execution time on large tenants.

---

## 🔐 Permissions Required

The identity running this script needs the following Azure RBAC roles:

| Scope | Role | Purpose |
|---|---|---|
| Management Group / Root | `Reader` | Enumerate subscriptions |
| Each Subscription | `Reader` | List VMs, disks, NICs, Public IPs |

> 💡 **Tip:** Assign `Reader` at the **Management Group** level to automatically cover all child subscriptions.

To grant via PowerShell:
```powershell
New-AzRoleAssignment `
    -ObjectId         "<service-principal-or-user-object-id>" `
    -RoleDefinitionName "Reader" `
    -Scope            "/providers/Microsoft.Management/managementGroups/<mg-id>"
```

---

## ⚠️ Known Limitations

- **Pricing is retail (PAYG)** — does not reflect EA, CSP, Reserved Instance, or Spot discounts
- **Ultra Disk pricing** may not always return results depending on region availability in the API
- **Very large tenants** (500+ VMs) may take 15–30+ minutes due to per-NIC and per-disk API calls
- Public IP lookup requires access to the NIC's resource group
- **Unmanaged disks** (VHD blobs in Storage Accounts) are not priced — only Managed Disks are supported

---

## 🔧 Troubleshooting

| Issue | Resolution |
|---|---|
| `No VMs found` for a subscription | Verify `Reader` role is assigned on that subscription |
| `Pricing API error` warnings | Transient; the script continues and marks price as `null` |
| NIC details missing | Ensure the running identity has access to the NIC resource group |
| Excel file locked / in use | Close the file before re-running the script |
| `ImportExcel` not found | Run `Install-Module ImportExcel -Scope CurrentUser -Force` |
| Script runs slowly | Normal for large tenants — NIC/disk/pricing API calls are sequential per VM |

---

## 📄 License

MIT License. See [LICENSE](LICENSE) for details.

---

## 🙌 Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

---

*Generated for enterprise Azure inventory and FinOps cost visibility.*
