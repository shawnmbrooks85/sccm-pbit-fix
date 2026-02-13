<#
.SYNOPSIS
    Fixes SCCM/ConfigMgr Power BI templates (.pbit) for Power BI Report Server compatibility.

.DESCRIPTION
    Power BI Report Server does not support composite models (DirectQuery + Import).
    This script converts DirectQuery templates to Import mode by modifying:
    1. DataModelSchema - partition modes from "directQuery" to "import"
    2. DataModelSchema - adds "defaultMode": "import" to the model
    3. DataMashup - changes IsDirectQuery flags from true (l1) to false (l0)
    
    The original ZIP structure (forward-slash paths) is preserved by modifying
    entries in-place rather than extracting and re-zipping.

.PARAMETER PbitPath
    Path to the .pbit file to fix. Can be a single file or wildcard pattern.

.PARAMETER BackupSuffix
    Suffix for backup files. Default: ".backup"

.PARAMETER NoBackup
    Skip creating backup files.

.EXAMPLE
    .\Fix-PbitTemplate.ps1 -PbitPath ".\Client Status.pbit"
    
.EXAMPLE
    .\Fix-PbitTemplate.ps1 -PbitPath ".\*.pbit"
    
.EXAMPLE
    Get-ChildItem ".\*.pbit" | .\Fix-PbitTemplate.ps1
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
    [Alias("FullName")]
    [string[]]$PbitPath,

    [string]$BackupSuffix = ".backup",

    [switch]$NoBackup
)

begin {
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    function Fix-DataModelSchema {
        param([byte[]]$OriginalBytes)
        
        # Read as UTF-16 LE (no BOM) - standard PBIT encoding
        $encoding = New-Object System.Text.UnicodeEncoding($false, $false)
        $text = $encoding.GetString($OriginalBytes)
        
        # 1. Replace all directQuery partition modes with import
        $before = [regex]::Matches($text, '"mode":\s*"directQuery"').Count
        $text = $text -replace '"mode":\s*"directQuery"', '"mode": "import"'
        
        # 2. Add defaultMode if not present
        if ($text -notmatch '"defaultMode"') {
            $text = $text.Replace('"culture": "en-US",', '"defaultMode": "import",' + "`r`n" + '  "culture": "en-US",')
        }
        
        $after = [regex]::Matches($text, '"mode":\s*"import"').Count
        $modifiedBytes = $encoding.GetBytes($text)
        
        return @{
            Bytes = $modifiedBytes
            DirectQueryConverted = $before
            ImportModeTotal = $after
        }
    }

    function Fix-DataMashup {
        param([byte[]]$OriginalBytes)
        
        # Byte-level replacement: IsDirectQuery" Value="l1" -> IsDirectQuery" Value="l0"
        # l1 = true (DirectQuery), l0 = false (Import)
        # Same byte length so no offset corruption
        $searchBytes = [System.Text.Encoding]::UTF8.GetBytes('IsDirectQuery" Value="l1"')
        $replaceBytes = [System.Text.Encoding]::UTF8.GetBytes('IsDirectQuery" Value="l0"')
        
        $modifiedBytes = [byte[]]$OriginalBytes.Clone()
        $replacements = 0
        
        for ($i = 0; $i -le $modifiedBytes.Length - $searchBytes.Length; $i++) {
            $match = $true
            for ($j = 0; $j -lt $searchBytes.Length; $j++) {
                if ($modifiedBytes[$i + $j] -ne $searchBytes[$j]) {
                    $match = $false
                    break
                }
            }
            if ($match) {
                for ($j = 0; $j -lt $replaceBytes.Length; $j++) {
                    $modifiedBytes[$i + $j] = $replaceBytes[$j]
                }
                $replacements++
            }
        }
        
        return @{
            Bytes = $modifiedBytes
            FlagsFixed = $replacements
        }
    }

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
}

process {
    foreach ($path in $PbitPath) {
        # Resolve wildcards
        $files = Resolve-Path $path -ErrorAction SilentlyContinue
        if (-not $files) {
            Write-Warning "File not found: $path"
            continue
        }

        foreach ($file in $files) {
            $filePath = $file.Path
            $fileName = Split-Path $filePath -Leaf
            
            Write-Host "`n========================================" -ForegroundColor Cyan
            Write-Host "Processing: $fileName" -ForegroundColor Cyan
            Write-Host "========================================" -ForegroundColor Cyan
            
            # --- Step 1: Create backup ---
            if (-not $NoBackup) {
                $backupPath = "$filePath$BackupSuffix"
                if (-not (Test-Path $backupPath)) {
                    Copy-Item $filePath $backupPath
                    Write-Host "[OK] Backup created: $backupPath" -ForegroundColor Green
                } else {
                    Write-Host "[--] Backup already exists, skipping" -ForegroundColor Yellow
                }
            }

            # --- Step 2: Read original entries ---
            Write-Host "`n[..] Reading original PBIT..." -ForegroundColor Gray
            $zip = [System.IO.Compression.ZipFile]::OpenRead($filePath)
            
            $schemaBytes = Read-ZipEntry $zip "DataModelSchema"
            $mashupBytes = Read-ZipEntry $zip "DataMashup"
            $zip.Dispose()

            if (-not $schemaBytes) {
                Write-Warning "DataModelSchema not found in $fileName - skipping"
                continue
            }
            if (-not $mashupBytes) {
                Write-Warning "DataMashup not found in $fileName - skipping"
                continue
            }

            # --- Step 3: Fix DataModelSchema ---
            Write-Host "[..] Fixing DataModelSchema..." -ForegroundColor Gray
            $schemaResult = Fix-DataModelSchema $schemaBytes
            Write-Host "[OK] Converted $($schemaResult.DirectQueryConverted) DirectQuery partitions to Import" -ForegroundColor Green

            if ($schemaResult.DirectQueryConverted -eq 0) {
                Write-Host "[--] No DirectQuery partitions found - template may already be Import mode" -ForegroundColor Yellow
            }

            # --- Step 4: Fix DataMashup ---
            Write-Host "[..] Fixing DataMashup IsDirectQuery flags..." -ForegroundColor Gray
            $mashupResult = Fix-DataMashup $mashupBytes
            Write-Host "[OK] Fixed $($mashupResult.FlagsFixed) IsDirectQuery flags" -ForegroundColor Green

            # --- Step 5: Write back (in-place, preserving ZIP structure) ---
            Write-Host "[..] Writing fixes to PBIT (preserving ZIP structure)..." -ForegroundColor Gray
            $zip = [System.IO.Compression.ZipFile]::Open($filePath, [System.IO.Compression.ZipArchiveMode]::Update)
            
            Write-ZipEntry $zip "DataModelSchema" $schemaResult.Bytes | Out-Null
            Write-ZipEntry $zip "DataMashup" $mashupResult.Bytes | Out-Null
            
            $zip.Dispose()
            Write-Host "[OK] PBIT saved successfully" -ForegroundColor Green

            # --- Summary ---
            Write-Host "`n  Summary for $fileName`:" -ForegroundColor White
            Write-Host "    DirectQuery -> Import partitions: $($schemaResult.DirectQueryConverted)" 
            Write-Host "    IsDirectQuery flags fixed:        $($mashupResult.FlagsFixed)"
            Write-Host "    Total Import mode partitions:     $($schemaResult.ImportModeTotal)"
        }
    }
}

end {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "All templates processed." -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "`nNext steps:" -ForegroundColor White
    Write-Host "  1. Open each .pbit in Power BI Desktop for Report Server"
    Write-Host "  2. Enter your SCCM database connection details when prompted"
    Write-Host "  3. Wait for data to load"
    Write-Host "  4. Save as .pbix and publish to Report Server"
}
