<#
.SYNOPSIS
    Fixes the SCCM Content Status Power BI template (.pbit) for Report Server.

.DESCRIPTION
    The Content Status template has additional issues beyond DirectQuery mode:
    
    1. Column constraints: 55 columns across all tables have isNullable=false,
       but the actual database data contains blanks/NULLs. DirectQuery doesn't 
       validate this at load time, but Import mode does.
    
    2. Redundant PackageID column: v_PackageStatusDistPointsSumm has a PackageID
       column from pkgsum.* that duplicates ContentPackageID (which the relationship
       uses). The source data can have blank PackageIDs.
    
    3. Incorrect relationship cardinality: v_ClientDownloadHistoryDP_BG has 
       fromCardinality="one" on PackageID, but multiple download records can 
       share the same PackageID.
    
    This script uses JSON parsing to comprehensively fix all issues.

.PARAMETER PbitPath
    Path to the Content Status .pbit file.

.PARAMETER NoBackup
    Skip creating a backup file.

.EXAMPLE
    .\Fix-ContentStatusPbit.ps1 -PbitPath ".\Content Status.pbit"

.EXAMPLE
    .\Fix-ContentStatusPbit.ps1 -PbitPath ".\Content Status.pbit" -NoBackup
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$PbitPath,

    [switch]$NoBackup
)

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-ZipEntry {
    param($ZipArchive, [string]$EntryName)
    $entry = $ZipArchive.GetEntry($EntryName)
    if (-not $entry) { return $null }
    $stream = $entry.Open()
    $ms = New-Object System.IO.MemoryStream
    $stream.CopyTo($ms)
    $result = $ms.ToArray()
    $ms.Dispose()
    $stream.Dispose()
    return $result
}

function Write-ZipEntry {
    param($ZipArchive, [string]$EntryName, [byte[]]$Data)
    $entry = $ZipArchive.GetEntry($EntryName)
    if (-not $entry) { return $false }
    $stream = $entry.Open()
    $stream.SetLength(0)
    $stream.Write($Data, 0, $Data.Length)
    $stream.Dispose()
    return $true
}

$fileName = Split-Path $PbitPath -Leaf

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Fixing: $fileName" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# --- Backup ---
if (-not $NoBackup) {
    $backupPath = "$PbitPath.backup"
    if (-not (Test-Path $backupPath)) {
        Copy-Item $PbitPath $backupPath
        Write-Host "[OK] Backup created: $backupPath" -ForegroundColor Green
    } else {
        Write-Host "[--] Backup already exists, skipping" -ForegroundColor Yellow
    }
}

# --- Read ---
Write-Host "`n[..] Reading PBIT..." -ForegroundColor Gray
$zip = [System.IO.Compression.ZipFile]::OpenRead($PbitPath)
$schemaBytes = Read-ZipEntry $zip "DataModelSchema"
$mashupBytes = Read-ZipEntry $zip "DataMashup"
$zip.Dispose()

if (-not $schemaBytes) {
    Write-Error "DataModelSchema not found in $fileName"
    exit 1
}

$encoding = New-Object System.Text.UnicodeEncoding($false, $false)
$text = $encoding.GetString($schemaBytes)
$json = $text | ConvertFrom-Json

# ============================================
# Fix 1: Set ALL non-RowNumber columns to nullable
# ============================================
Write-Host "[..] Fix 1: Setting all data columns to nullable..." -ForegroundColor Gray
$nullableFixCount = 0

foreach ($table in $json.model.tables) {
    foreach ($col in $table.columns) {
        if ($col.type -ne "rowNumber" -and $col.PSObject.Properties["isNullable"]) {
            if ($col.isNullable -eq $false) {
                $col.isNullable = $true
                $nullableFixCount++
            }
        }
    }
}
Write-Host "  [OK] Changed $nullableFixCount columns from isNullable=false to true" -ForegroundColor Green

# ============================================
# Fix 2: Remove PackageID column from v_PackageStatusDistPointsSumm
# ============================================
Write-Host "[..] Fix 2: Removing redundant PackageID column..." -ForegroundColor Gray

$psdp = $json.model.tables | Where-Object { $_.name -eq "v_PackageStatusDistPointsSumm" }
if ($psdp) {
    $beforeCount = $psdp.columns.Count
    $psdp.columns = @($psdp.columns | Where-Object { $_.name -ne "PackageID" })
    $afterCount = $psdp.columns.Count

    if ($afterCount -lt $beforeCount) {
        Write-Host "  [OK] Removed PackageID column ($beforeCount -> $afterCount columns)" -ForegroundColor Green
    } else {
        Write-Host "  [--] PackageID column not found" -ForegroundColor Yellow
    }
} else {
    Write-Host "  [--] Table v_PackageStatusDistPointsSumm not found" -ForegroundColor Yellow
}

# ============================================
# Fix 3: Remove fromCardinality "one" from v_ClientDownloadHistoryDP_BG
# ============================================
Write-Host "[..] Fix 3: Fixing relationship cardinality..." -ForegroundColor Gray

$fixedRel = $false
foreach ($rel in $json.model.relationships) {
    if ($rel.fromTable -eq "v_ClientDownloadHistoryDP_BG" -and $rel.fromColumn -eq "PackageID") {
        if ($rel.PSObject.Properties["fromCardinality"]) {
            $rel.PSObject.Properties.Remove("fromCardinality")
            $fixedRel = $true
        }
    }
}
if ($fixedRel) {
    Write-Host "  [OK] Removed fromCardinality 'one'" -ForegroundColor Green
} else {
    Write-Host "  [--] No cardinality to fix" -ForegroundColor Yellow
}

# ============================================
# Fix 4: DirectQuery -> Import
# ============================================
Write-Host "[..] Fix 4: DirectQuery -> Import conversion..." -ForegroundColor Gray

$dqCount = 0
foreach ($table in $json.model.tables) {
    foreach ($p in $table.partitions) {
        if ($p.mode -eq "directQuery") {
            $p.mode = "import"
            $dqCount++
        }
    }
}
Write-Host "  [OK] Converted $dqCount partitions to Import" -ForegroundColor Green

# Add defaultMode
if (-not $json.model.PSObject.Properties["defaultMode"]) {
    $json.model | Add-Member -MemberType NoteProperty -Name "defaultMode" -Value "import" -Force
    Write-Host "  [OK] Added defaultMode: import" -ForegroundColor Green
} else {
    $json.model.defaultMode = "import"
    Write-Host "  [--] defaultMode set to import" -ForegroundColor Yellow
}

# ============================================
# Serialize schema
# ============================================
Write-Host "[..] Serializing modified schema..." -ForegroundColor Gray
$newText = $json | ConvertTo-Json -Depth 20
$modifiedSchemaBytes = $encoding.GetBytes($newText)

# ============================================
# Fix 5: DataMashup IsDirectQuery flags
# ============================================
Write-Host "[..] Fix 5: DataMashup IsDirectQuery flags..." -ForegroundColor Gray

$mashupFixCount = 0
$modifiedMashupBytes = $null
if ($mashupBytes) {
    $searchBytes = [System.Text.Encoding]::UTF8.GetBytes('IsDirectQuery" Value="l1"')
    $replaceBytes = [System.Text.Encoding]::UTF8.GetBytes('IsDirectQuery" Value="l0"')
    $modifiedMashupBytes = [byte[]]$mashupBytes.Clone()
    
    for ($i = 0; $i -le $modifiedMashupBytes.Length - $searchBytes.Length; $i++) {
        $match = $true
        for ($j = 0; $j -lt $searchBytes.Length; $j++) {
            if ($modifiedMashupBytes[$i + $j] -ne $searchBytes[$j]) {
                $match = $false
                break
            }
        }
        if ($match) {
            for ($j = 0; $j -lt $replaceBytes.Length; $j++) {
                $modifiedMashupBytes[$i + $j] = $replaceBytes[$j]
            }
            $mashupFixCount++
        }
    }
    Write-Host "  [OK] Fixed $mashupFixCount IsDirectQuery flags" -ForegroundColor Green
}

# ============================================
# Write all changes
# ============================================
Write-Host "[..] Writing to PBIT (preserving ZIP structure)..." -ForegroundColor Gray
$zip = [System.IO.Compression.ZipFile]::Open($PbitPath, [System.IO.Compression.ZipArchiveMode]::Update)
Write-ZipEntry $zip "DataModelSchema" $modifiedSchemaBytes | Out-Null
if ($modifiedMashupBytes) {
    Write-ZipEntry $zip "DataMashup" $modifiedMashupBytes | Out-Null
}
$zip.Dispose()
Write-Host "[OK] PBIT saved successfully" -ForegroundColor Green

# ============================================
# Summary
# ============================================
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  Summary for $fileName" -ForegroundColor White
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "    Nullable columns fixed:          $nullableFixCount"
Write-Host "    DirectQuery partitions converted: $dqCount"
Write-Host "    IsDirectQuery flags fixed:        $mashupFixCount"
Write-Host ""
Write-Host "Next steps:" -ForegroundColor White
Write-Host "  1. Open the .pbit in Power BI Desktop for Report Server"
Write-Host "  2. Enter your SCCM database connection details when prompted"
Write-Host "  3. Wait for data to load"
Write-Host "  4. Save as .pbix and publish to Report Server"
