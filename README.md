# SCCM Power BI Template Fix for Report Server

PowerShell scripts to fix SCCM/ConfigMgr Power BI templates (`.pbit`) for compatibility with **Power BI Report Server**.

## Problem

Microsoft's SCCM Power BI templates use **DirectQuery** mode with **composite model** compatibility. Power BI Report Server does not support composite models, resulting in:

> *"This file uses a composite model which combines DirectQuery sources and/or imported data. These models aren't currently supported in Power BI Report Server."*

Additionally, some templates have data model constraints (`isNullable=false`, incorrect relationship cardinalities) that only fail in Import mode because DirectQuery doesn't validate these at load time.

## What These Scripts Fix

| Fix | Description |
|-----|-------------|
| DirectQuery → Import | Converts all partition modes and `IsDirectQuery` flags |
| `defaultMode` | Adds `"defaultMode": "import"` to the model root |
| `isNullable` constraints | Relaxes `isNullable=false` on data columns that can have blanks |
| Relationship cardinality | Fixes incorrect `fromCardinality: "one"` definitions |
| ZIP path preservation | Modifies entries in-place to preserve forward-slash paths |

## Prerequisites

- **Power BI Desktop optimized for Report Server** (for opening fixed templates)
- **PowerShell 5.1+** (included with Windows)
- **SCCM/ConfigMgr Power BI templates** (`.pbit` files) — download from Microsoft:
  - **[Microsoft Endpoint Configuration Manager Sample Power BI Reports](https://www.microsoft.com/en-us/download/details.aspx?id=101452)**
  - Or from your SCCM installation: `<SCCM Install Dir>\tools\ServerReportingService\Templates\`

## Quick Start

### Fix a Single Template (DirectQuery → Import)

```powershell
.\Fix-PbitTemplate.ps1 -PbitPath ".\Client Status.pbit"
```

### Fix All Templates in a Folder

```powershell
.\Fix-PbitTemplate.ps1 -PbitPath ".\*.pbit"
```

### Fix Content Status Template (Additional Model Fixes)

The Content Status template requires extra fixes beyond DirectQuery conversion:

```powershell
.\Fix-ContentStatusPbit.ps1 -PbitPath ".\Content Status.pbit"
```

### Pipeline Support

```powershell
Get-ChildItem ".\*.pbit" | .\Fix-PbitTemplate.ps1
```

### Skip Backups

```powershell
.\Fix-PbitTemplate.ps1 -PbitPath ".\*.pbit" -NoBackup
```

## Scripts

### `Fix-PbitTemplate.ps1`

General-purpose fix for any SCCM `.pbit` template. Handles:
- DirectQuery → Import partition mode conversion
- `IsDirectQuery` flag byte-level replacement in DataMashup
- `defaultMode: import` addition
- Automatic `.backup` file creation

### `Fix-ContentStatusPbit.ps1`

Targeted fix for the **Content Status** template which has additional issues:
- Removes redundant `PackageID` column from `v_PackageStatusDistPointsSumm`
- Sets all `isNullable=false` columns to `true` (55 columns across all tables)
- Removes incorrect `fromCardinality: "one"` from `v_ClientDownloadHistoryDP_BG` relationship
- Includes all DirectQuery → Import fixes

## After Fixing

1. Open the `.pbit` file in **Power BI Desktop optimized for Report Server**
2. Enter your SCCM database server and catalog name when prompted
3. Data loads into memory (Import mode)
4. Verify visuals populate with data
5. Save as `.pbix` → publish to Report Server

## Technical Reference

### `.pbit` File Structure

A `.pbit` file is a ZIP archive:

```
├── DataModelSchema        # JSON (UTF-16 LE, no BOM) – tabular model definition
├── DataMashup             # 8-byte header + OPC ZIP – Power Query M expressions
├── Report/Layout          # JSON – report pages and visual definitions
├── [Content_Types].xml    # OPC content type definitions
├── Version, Metadata, Settings, SecurityBindings, DiagramLayout
└── Report/CustomVisuals/, Report/StaticResources/
```

### Key Modifications

| File | Property | Original | Fixed |
|------|----------|----------|-------|
| `DataModelSchema` | Partition `"mode"` | `"directQuery"` | `"import"` |
| `DataModelSchema` | Model `"defaultMode"` | *(missing)* | `"import"` |
| `DataModelSchema` | Column `"isNullable"` | `false` | `true` |
| `DataMashup` | `IsDirectQuery Value` | `"l1"` (true) | `"l0"` (false) |

### ⚠️ ZIP Path Warning

Never extract and re-zip a `.pbit` file using `Compress-Archive` or `ZipFile.CreateFromDirectory` on Windows. These tools create **backslash** paths (`Report\Layout`) but Power BI requires **forward slashes** (`Report/Layout`). Always modify entries **in-place** using `ZipArchiveMode.Update`.

## Tested Environment

- **Power BI Report Server** — January 2026
- **Power BI Desktop (optimized for Report Server)** — 2026

## Tested Templates

| Template | Fix Script | Status |
|----------|-----------|--------|
| Client Status | `Fix-PbitTemplate.ps1` | ✅ |
| Content Status | `Fix-ContentStatusPbit.ps1` | ✅ |
| Microsoft Edge Management | `Fix-PbitTemplate.ps1` | ✅ |
| Software Update Compliance Status | `Fix-PbitTemplate.ps1` | ✅ |
| Software Update Deployment Status | `Fix-PbitTemplate.ps1` | ✅ |

## License

MIT License – see [LICENSE](LICENSE)

> **Note:** These scripts are community tools that modify Microsoft's copyrighted `.pbit` template files. The templates themselves are not included in this repository. Obtain them from your SCCM/ConfigMgr installation.
