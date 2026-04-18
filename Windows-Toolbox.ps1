<#
.SYNOPSIS
    Windows Management & Automation Toolbox

.DESCRIPTION
    A professional, menu-driven PowerShell utility that integrates:
      - Software installation and updates via Winget
      - WiFi profile backup and restore
      - Device driver export and import (Native & OSD methods)
      - Windows activation via BIOS OEM key

.NOTES
    Author  : Yoennis Olmo
    Version : v1.3
    Date    : 2026-04-01
    Requires: Administrator privileges, PowerShell 5.1+

.EXAMPLE
    # PowerShell 5.1
    powershell.exe -ExecutionPolicy Bypass -File Windows-Toolbox.ps1

    # PowerShell 7+
    pwsh.exe -ExecutionPolicy Bypass -File Windows-Toolbox.ps1
#>

#Requires -RunAsAdministrator

[CmdletBinding()]
param()

# ===========================================================================
# CONFIGURATION
# ===========================================================================

$script:LogFile    = "$env:TEMP\Windows-Toolbox_$(Get-Date -Format 'yyyyMMdd').log"
$script:Utf8NoBom  = [System.Text.UTF8Encoding]::new($false)   # UTF-8 without BOM — consistent on PS5.1 and PS7


# ===========================================================================
# LOGGING
# ===========================================================================

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet("INFO", "WARN", "ERROR")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$timestamp] [$Level] $Message"
    try {
        # AppendAllText with explicit UTF-8 no-BOM is consistent on PS5.1 and PS7.
        # (Add-Content writes ANSI on PS5.1 but UTF-8 on PS7 when no -Encoding is given.)
        [System.IO.File]::AppendAllText($script:LogFile, "$entry`n", $script:Utf8NoBom)
    } catch {
        # Silently continue if log write fails  -  don't crash the UI
    }
}

# ===========================================================================
# FOLDER BROWSER
# ===========================================================================

function Get-FolderFromBrowser {
    [CmdletBinding()]
    param(
        [string]$Description = "Select a folder",
        [string]$InitialPath  = [Environment]::GetFolderPath('Desktop')
    )
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        $dialog                  = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description      = $Description
        $dialog.ShowNewFolderButton = $true
        $dialog.SelectedPath     = $InitialPath
        $result = $dialog.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            return $dialog.SelectedPath
        }
    } catch {
        Write-Host "  WARNING: Folder browser unavailable  -  $_" -ForegroundColor Yellow
    }
    return $null
}

# ===========================================================================
# MENU
# ===========================================================================

# ---------------------------------------------------------------------------
# Menu item definitions  -  each entry drives both the display and dispatch.
# ---------------------------------------------------------------------------
$script:MenuItems = @(
    [PSCustomObject]@{ Number = 1;  Label = 'Export Installed Software List (Winget)';         Section = 'Software';   Action = { Export-WingetList } }
    [PSCustomObject]@{ Number = 2;  Label = 'Import & Install from Software List (Winget)';    Section = 'Software';   Action = { Import-WingetList } }
    [PSCustomObject]@{ Number = 3;  Label = 'Backup WiFi Profiles';                            Section = 'WiFi';       Action = { Export-WifiProfiles } }
    [PSCustomObject]@{ Number = 4;  Label = 'Restore WiFi Profiles';                           Section = 'WiFi';       Action = { Import-WifiProfiles } }
    [PSCustomObject]@{ Number = 5;  Label = 'Export Drivers  (Native  -  Export-WindowsDriver)'; Section = 'Drivers';    Action = { Export-DriversNative } }
    [PSCustomObject]@{ Number = 6;  Label = 'Import Drivers  (Native  -  pnputil)';              Section = 'Drivers';    Action = { Import-DriversNative } }
    [PSCustomObject]@{ Number = 7;  Label = 'Download Driver Pack  (OSD  -  Save-MyDriverPack)'; Section = 'Drivers';    Action = { Export-DriversOSD } }
    [PSCustomObject]@{ Number = 8;  Label = 'Install Drivers  (OSD Pack  -  Smart HWID Match)';  Section = 'Drivers';    Action = { Install-DriversFromOSDPack } }
    [PSCustomObject]@{ Number = 9;  Label = 'Install Single Driver  (Browse to INF file)';     Section = 'Drivers';    Action = { Install-SingleDriver } }
    [PSCustomObject]@{ Number = 10; Label = 'Activate Windows (BIOS OEM Key)';                 Section = 'Activation'; Action = { Invoke-WindowsActivation } }
    [PSCustomObject]@{ Number = 11; Label = 'Export BIOS OEM Key to File';                     Section = 'Activation'; Action = { Export-BiosOemKey } }
    [PSCustomObject]@{ Number = 12; Label = 'Activate from Key File';                          Section = 'Activation'; Action = { Invoke-WindowsActivationFromFile } }
    [PSCustomObject]@{ Number = 13; Label = 'Check Activation Status';                         Section = 'Activation'; Action = { Get-ActivationStatus } }
    [PSCustomObject]@{ Number = 0;  Label = 'Exit';                                            Section = 'Exit';       Action = $null }
)

# ---------------------------------------------------------------------------
# Show-Menu  -  draw the menu; $SelectedIndex is highlighted
# ---------------------------------------------------------------------------
function Show-Menu {
    [CmdletBinding()]
    param([int]$SelectedIndex = 0)

    Clear-Host

    $width = [Math]::Min([Console]::WindowWidth - 4, 72)
    $sep   = '-' * $width

    # -- Title --------------------------------------------------------------
    Write-Host ''
    $titleText = 'WINDOWS MANAGEMENT TOOLBOX  v1.3'
    $padLeft   = [int][Math]::Floor(($width - $titleText.Length) / 2)
    Write-Host (' ' * ($padLeft + 2)) -NoNewline
    Write-Host $titleText -ForegroundColor Cyan
    Write-Host "  $sep" -ForegroundColor DarkGray
    Write-Host ''

    # -- Menu items ----------------------------------------------------------
    $lastSection = ''
    $idx = 0
    foreach ($item in $script:MenuItems) {
        $isSelected = ($idx -eq $SelectedIndex)

        # Section header
        if ($item.Section -ne $lastSection) {
            if ($lastSection -ne '' -and $item.Section -ne 'Exit') {
                Write-Host ''
            }
            if ($item.Section -ne 'Exit') {
                Write-Host "  $($item.Section)" -ForegroundColor Yellow
            }
            $lastSection = $item.Section
        }

        $numPad = $item.Number.ToString().PadLeft(2)
        $label  = "[$numPad]  $($item.Label)"

        if ($isSelected) {
            Write-Host '  > ' -NoNewline -ForegroundColor Cyan
            Write-Host $label -ForegroundColor White -BackgroundColor DarkBlue
        } elseif ($item.Section -eq 'Exit') {
            Write-Host "    $label" -ForegroundColor Red
        } else {
            Write-Host "    $label" -ForegroundColor Gray
        }

        $idx++
    }

    # -- Footer -------------------------------------------------------------
    Write-Host ''
    Write-Host "  $sep" -ForegroundColor DarkGray
    Write-Host '  UpDn Navigate   Enter: Select   0-9: Quick jump   Esc: Exit' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host "  Log: $script:LogFile" -ForegroundColor DarkGray
}

# ===========================================================================
# SOFTWARE MANAGEMENT (WINGET)
# ===========================================================================

function Invoke-WingetInstall {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$PackageIds
    )

    Write-Host ""
    Write-Host "  Refreshing package sources..." -ForegroundColor DarkGray
    $null = winget source update 2>&1
    Write-Host "  Starting software installation..." -ForegroundColor Cyan
    Write-Host "  $($PackageIds.Count) package(s) queued." -ForegroundColor White
    Write-Host ""
    Write-Log -Message "Winget install started  -  $($PackageIds.Count) packages"

    $successCount = 0
    $skipCount    = 0
    $errorCount   = 0

    foreach ($id in $PackageIds) {
        Write-Host "  Installing: $id" -NoNewline -ForegroundColor Cyan
        try {
            $null = winget install --id $id --silent --accept-package-agreements --accept-source-agreements 2>&1
            $code = $LASTEXITCODE

            if ($code -eq 0) {
                Write-Host "  [OK]" -ForegroundColor Green
                Write-Log -Message "Installed: $id"
                $successCount++
            } elseif ($code -eq -1978335189 -or $code -eq -1978335163) {
                Write-Host "  [SKIP] Already up to date." -ForegroundColor DarkGray
                Write-Log -Message "Already installed/up to date (skipped): $id"
                $skipCount++
            } elseif ($code -eq -1978335212) {
                Write-Host "  [WARN] Package not found  -  ID may have changed or source unavailable." -ForegroundColor Yellow
                Write-Log -Message "Package not found: $id (exit $code)" -Level "WARN"
                $errorCount++
            } elseif ($code -eq -1978335159) {
                Write-Host "  [WARN] Reboot required  -  restart Windows and run this again." -ForegroundColor Yellow
                Write-Log -Message "Install blocked (reboot required): $id (exit $code)" -Level "WARN"
                $errorCount++
            } elseif ($code -eq -1978335164) {
                Write-Host "  [SKIP] Not applicable for this system (internal/system package)." -ForegroundColor DarkGray
                Write-Log -Message "Package not applicable (skipped): $id (exit $code)"
                $skipCount++
            } else {
                Write-Host "  [WARN] Exit code: $code" -ForegroundColor Yellow
                Write-Log -Message "Winget exit $code for $id" -Level "WARN"
                $errorCount++
            }
        } catch {
            Write-Host "  [ERROR] $_" -ForegroundColor Red
            Write-Log -Message "Failed to install $id  -  $_" -Level "ERROR"
            $errorCount++
        }
    }

    Write-Host "  ----------------------------------------" -ForegroundColor DarkGray
    Write-Host "  Results: " -NoNewline -ForegroundColor White
    Write-Host "$successCount installed  " -NoNewline -ForegroundColor Green
    Write-Host "$skipCount skipped  " -NoNewline -ForegroundColor Yellow
    Write-Host "$errorCount warnings" -ForegroundColor Red
    Write-Log -Message "Winget install complete  -  Installed:$successCount Skipped:$skipCount Warnings:$errorCount"
}

function Export-WingetList {
    [CmdletBinding()]
    param()

    Write-Host ""
    Write-Host "  Select the destination folder in the Explorer window..." -ForegroundColor Cyan
    Write-Host ""

    $root = Get-FolderFromBrowser -Description "Select destination for software list export"
    if (-not $root) {
        Write-Host "  Cancelled  -  no folder selected." -ForegroundColor Yellow
        Write-Log -Message "Winget list export cancelled  -  no folder selected" -Level "WARN"
        return
    }

    $outFile = Join-Path $root "winget-software-list.txt"
    Write-Host "  Reading installed software via winget list..." -ForegroundColor Cyan
    Write-Log -Message "Winget list export started  -  destination: $outFile"

    try {
        # Capture winget list output (suppress progress/error noise)
        $raw = winget list --accept-source-agreements 2>&1 | Where-Object { $_ -is [string] }

        # Find the separator line (row of dashes) to locate column positions
        $sepIdx = -1
        for ($i = 0; $i -lt $raw.Count; $i++) {
            if ($raw[$i] -match '^[\s\-]+$' -and ($raw[$i] -replace '\s','').Length -gt 10) {
                $sepIdx = $i
                break
            }
        }
        if ($sepIdx -lt 1) { throw "Could not parse winget list output  -  unexpected format." }

        $headerLine = $raw[$sepIdx - 1]
        $nameCol    = $headerLine.IndexOf('Name')
        $idCol      = $headerLine.IndexOf('Id')
        $verCol     = $headerLine.IndexOf('Version')

        if ($nameCol -lt 0 -or $idCol -lt 0 -or $verCol -lt 0) {
            throw "Could not find expected columns (Name, Id, Version) in winget list header."
        }

        # Parse each data row using column positions
        $entries = for ($i = $sepIdx + 1; $i -lt $raw.Count; $i++) {
            $line = $raw[$i]
            if ($line.Length -le $verCol) { continue }
            $name    = $line.Substring($nameCol, $idCol - $nameCol).Trim()
            $id      = $line.Substring($idCol,   $verCol - $idCol).Trim()
            $version = $line.Substring($verCol).Trim() -replace '\s.*$'  # stop at next column

            # Include only real user-installable applications:
            #   - ID must start with a letter (excludes numeric version strings like 10.0.5)
            #   - ID must contain a dot separating two letter-based segments (Publisher.App)
            #   - ID must NOT start with MSIX\ (internal Windows workload packages)
            #   - ID must NOT be a pure version number (digits and dots only)
            if ($id -match '^[A-Za-z][A-Za-z0-9\-]+\.[A-Za-z][A-Za-z0-9\.\-]+$' -and
                $id -notmatch '^MSIX\\') {
                [PSCustomObject]@{ Name = $name; Id = $id; Version = $version }
            }
        }

        if (-not $entries -or $entries.Count -eq 0) {
            throw "No packages found in winget list output."
        }

        # Build output file
        $colName = 40; $colId = 45; $colVer = 12
        $lines = [System.Collections.Generic.List[string]]::new()
        $lines.Add("# Installed Software  -  exported $(Get-Date -Format 'yyyy-MM-dd HH:mm') by Windows-Toolbox")
        $lines.Add("# To reinstall: use option [2] Import & Install from Software List")
        $lines.Add("#")
        $lines.Add("# $('Name'.PadRight($colName))$('ID'.PadRight($colId))Version")
        $lines.Add("# $('-' * $colName)$('-' * $colId)$('-' * $colVer)")
        foreach ($e in $entries) {
            $lines.Add("  $($e.Name.PadRight($colName))$($e.Id.PadRight($colId))$($e.Version)")
        }
        # WriteAllLines with explicit no-BOM UTF-8 produces identical files on PS5.1 and PS7.
        # (Out-File -Encoding UTF8 adds a BOM on PS5.1 but not on PS7.)
        [System.IO.File]::WriteAllLines($outFile, [string[]]$lines, $script:Utf8NoBom)

        Write-Host "  $($entries.Count) package(s) found." -ForegroundColor Green
        Write-Host ""

        # Preview table
        Write-Host ("  " + "Name".PadRight($colName) + "ID".PadRight($colId) + "Version") -ForegroundColor DarkGray
        Write-Host ("  " + ('-' * ($colName + $colId + $colVer))) -ForegroundColor DarkGray
        foreach ($e in $entries) {
            $nameTrunc = if ($e.Name.Length -gt $colName - 2) { $e.Name.Substring(0, $colName - 3) + '...' } else { $e.Name }
            Write-Host ("  " + $nameTrunc.PadRight($colName) + $e.Id.PadRight($colId) + $e.Version) -ForegroundColor Gray
        }

        Write-Host ""
        Write-Host "  Saved to: $outFile" -ForegroundColor Green
        Write-Log -Message "Winget list exported  -  $($entries.Count) packages to $outFile"
    } catch {
        Write-Host "  ERROR: Export failed  -  $_" -ForegroundColor Red
        Write-Log -Message "Winget list export failed  -  $_" -Level "ERROR"
    }
}

function Import-WingetList {
    [CmdletBinding()]
    param()

    Write-Host ""
    Write-Host "  Select the 'winget-software-list.txt' file in the Explorer window..." -ForegroundColor Cyan
    Write-Host ""

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Title            = "Select software list file"
        $dialog.Filter           = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        $dialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
        $result = $dialog.ShowDialog()
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
            Write-Host "  Cancelled  -  no file selected." -ForegroundColor Yellow
            Write-Log -Message "Winget list import cancelled  -  no file selected" -Level "WARN"
            return
        }
        $importFile = $dialog.FileName
    } catch {
        $importFile = Read-Host "  Enter full path to the software list file"
    }

    if (-not (Test-Path $importFile)) {
        Write-Host "  ERROR: File not found: $importFile" -ForegroundColor Red
        Write-Log -Message "Winget list import failed: file not found  -  $importFile" -Level "ERROR"
        return
    }

    Write-Host "  File: $importFile" -ForegroundColor White
    Write-Log -Message "Winget list import started  -  source: $importFile"

    try {
        # Each data line: "  <Name padded>  <ID padded>  <Version>"
        # Comment lines start with #; extract the ID (second whitespace-separated token,
        # but columns may be fixed-width  -  safest: split by 2+ spaces and take token [1])
        $ids = Get-Content $importFile |
            Where-Object { $_.Trim() -ne '' -and -not $_.TrimStart().StartsWith('#') } |
            ForEach-Object {
                $parts = ($_.Trim() -split '\s{2,}')
                if ($parts.Count -ge 2) { $parts[1].Trim() }
            } |
            Where-Object { $_ -match '^[A-Za-z][A-Za-z0-9\-]+\.[A-Za-z][A-Za-z0-9\.\-]+$' -and $_ -notmatch '^MSIX\\' }

        if (-not $ids -or $ids.Count -eq 0) {
            Write-Host "  ERROR: No valid package IDs found in the file." -ForegroundColor Red
            Write-Log -Message "Winget list import: no valid IDs parsed from $importFile" -Level "ERROR"
            return
        }

        Write-Host "  $($ids.Count) package(s) to install." -ForegroundColor Cyan
        Write-Host ""
        Invoke-WingetInstall -PackageIds $ids
    } catch {
        Write-Host "  ERROR: Import failed  -  $_" -ForegroundColor Red
        Write-Log -Message "Winget list import failed  -  $_" -Level "ERROR"
    }
}

# ===========================================================================
# WIFI PROFILE MANAGEMENT
# ===========================================================================

function Export-WifiProfiles {
    [CmdletBinding()]
    param()

    Write-Host ""
    Write-Host "  Select the destination folder in the Explorer window..." -ForegroundColor Cyan
    Write-Host "  A 'WiFi-Profiles' subfolder will be created there." -ForegroundColor DarkGray
    Write-Host ""

    $root = Get-FolderFromBrowser -Description "Select destination for WiFi profile backup"
    if (-not $root) {
        Write-Host "  Cancelled  -  no folder selected." -ForegroundColor Yellow
        Write-Log -Message "WiFi export cancelled  -  no folder selected" -Level "WARN"
        return
    }

    $backupPath = Join-Path $root "WiFi-Profiles"
    Write-Host "  Destination: $backupPath" -ForegroundColor White
    Write-Log -Message "WiFi export started  -  destination: $backupPath"

    try {
        if (-not (Test-Path $backupPath)) {
            New-Item -ItemType Directory -Path $backupPath -Force | Out-Null
            Write-Host "  Created folder: $backupPath" -ForegroundColor DarkGray
        }

        # Pass folder= as a single quoted argument so paths with spaces work correctly
        $output = & netsh wlan export profile "folder=$backupPath" key=clear 2>&1

        # Count saved profiles from netsh success lines; surface only errors
        $savedCount = ($output | Where-Object { $_ -match 'is saved in file' }).Count
        $errorLines = $output | Where-Object { $_ -match '(?i)error|failed|not found' }

        if ($savedCount -gt 0) {
            Write-Host "  $savedCount profile(s) exported." -ForegroundColor Green
            Write-Host "  Saved to: $backupPath" -ForegroundColor Green
            Write-Log -Message "WiFi profiles exported  -  $savedCount profile(s) to $backupPath"
        } else {
            Write-Host "  WARNING: No profiles were exported." -ForegroundColor Yellow
            if ($errorLines) { $errorLines | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow } }
            Write-Log -Message "WiFi export: no profiles saved  -  check WLAN service" -Level "WARN"
        }
    } catch {
        Write-Host "  ERROR: WiFi export failed  -  $_" -ForegroundColor Red
        Write-Log -Message "WiFi export failed  -  $_" -Level "ERROR"
    }
}

function Import-WifiProfiles {
    [CmdletBinding()]
    param()

    Write-Host ""
    Write-Host "  Select the 'WiFi-Profiles' backup folder in the Explorer window..." -ForegroundColor Cyan
    Write-Host ""

    $backupPath = Get-FolderFromBrowser -Description "Select the WiFi-Profiles backup folder to restore from"
    if (-not $backupPath) {
        Write-Host "  Cancelled  -  no folder selected." -ForegroundColor Yellow
        Write-Log -Message "WiFi import cancelled  -  no folder selected" -Level "WARN"
        return
    }

    Write-Host "  Restoring from: $backupPath" -ForegroundColor White
    Write-Host ""
    Write-Log -Message "WiFi import started  -  source: $backupPath"

    try {
        if (-not (Test-Path $backupPath)) {
            Write-Host "  ERROR: Folder not found: $backupPath" -ForegroundColor Red
            Write-Log -Message "WiFi import failed: path not found  -  $backupPath" -Level "ERROR"
            return
        }

        $profiles = Get-ChildItem -Path $backupPath -Filter "*.xml" -ErrorAction Stop
        if ($profiles.Count -eq 0) {
            Write-Host "  WARNING: No XML profile files found in $backupPath" -ForegroundColor Yellow
            Write-Log -Message "WiFi import: no XML files found in $backupPath" -Level "WARN"
            return
        }

        $okCount   = 0
        $skipCount = 0
        $warnCount = 0

        foreach ($wifiProfile in $profiles) {
            Write-Host "  Importing: $($wifiProfile.Name)" -ForegroundColor Cyan
            try {
                $result = & netsh wlan add profile "filename=$($wifiProfile.FullName)" user=all 2>&1
                $resultStr = ($result | Out-String).Trim()

                if ($resultStr -match "added on interface|updated on interface") {
                    Write-Host "    [OK]" -ForegroundColor Green
                    Write-Log -Message "WiFi profile imported: $($wifiProfile.Name)"
                    $okCount++
                } elseif ($resultStr -match "already exists") {
                    Write-Host "    [SKIP] Profile already exists." -ForegroundColor Yellow
                    Write-Log -Message "WiFi profile already exists (skipped): $($wifiProfile.Name)" -Level "WARN"
                    $skipCount++
                } else {
                    Write-Host "    [WARN] $resultStr" -ForegroundColor Yellow
                    Write-Log -Message "WiFi import warning for $($wifiProfile.Name): $resultStr" -Level "WARN"
                    $warnCount++
                }
            } catch {
                Write-Host "    [ERROR] $_" -ForegroundColor Red
                Write-Log -Message "WiFi import error for $($wifiProfile.Name)  -  $_" -Level "ERROR"
                $warnCount++
            }
        }

        Write-Host ""
        Write-Host "  ----------------------------------------" -ForegroundColor DarkGray
        Write-Host "  Results: " -NoNewline -ForegroundColor White
        Write-Host "$okCount imported  " -NoNewline -ForegroundColor Green
        Write-Host "$skipCount skipped  " -NoNewline -ForegroundColor Yellow
        Write-Host "$warnCount warnings" -ForegroundColor Red
        Write-Log -Message "WiFi import complete  -  OK:$okCount Skipped:$skipCount Warnings:$warnCount"
    } catch {
        Write-Host "  ERROR: WiFi restore failed  -  $_" -ForegroundColor Red
        Write-Log -Message "WiFi restore failed  -  $_" -Level "ERROR"
    }
}

# ===========================================================================
# 5 & 6  -  DRIVER EXPORT / IMPORT (NATIVE)
# ===========================================================================

function Export-DriversNative {
    [CmdletBinding()]
    param()

    Write-Host ""
    Write-Host "  Select the destination folder in the Explorer window..." -ForegroundColor Cyan
    Write-Host "  A 'Windows-Drivers' subfolder will be created there." -ForegroundColor DarkGray
    Write-Host ""

    $root = Get-FolderFromBrowser -Description "Select destination for driver export"
    if (-not $root) {
        Write-Host "  Cancelled  -  no folder selected." -ForegroundColor Yellow
        Write-Log -Message "Driver export (native) cancelled  -  no folder selected" -Level "WARN"
        return
    }

    $destination = Join-Path $root "Windows-Drivers"
    Write-Host "  Destination: $destination" -ForegroundColor White
    Write-Log -Message "Driver export (native) started  -  destination: $destination"

    try {
        if (-not (Test-Path $destination)) {
            New-Item -ItemType Directory -Path $destination -Force | Out-Null
            Write-Host "  Created folder: $destination" -ForegroundColor DarkGray
        }

        Write-Host "  Exporting drivers... (this may take a moment)" -ForegroundColor Cyan
        $exported = Export-WindowsDriver -Online -Destination $destination -ErrorAction Stop
        $driverCount = ($exported | Measure-Object).Count
        Write-Host ""
        Write-Host "  Drivers backed up : $driverCount" -ForegroundColor Green
        Write-Host "  Driver export complete." -ForegroundColor Green
        Write-Host "  Saved to: $destination" -ForegroundColor Green
        Write-Log -Message "Drivers exported (native) to $destination  -  $driverCount drivers"
    } catch {
        Write-Host "  ERROR: Driver export failed  -  $_" -ForegroundColor Red
        Write-Log -Message "Driver export (native) failed  -  $_" -Level "ERROR"
    }
}

function Import-DriversNative {
    [CmdletBinding()]
    param()

    Write-Host ""
    Write-Host "  Select the 'Windows-Drivers' backup folder in the Explorer window..." -ForegroundColor Cyan
    Write-Host ""

    $sourcePath = Get-FolderFromBrowser -Description "Select the Windows-Drivers backup folder to restore from"
    if (-not $sourcePath) {
        Write-Host "  Cancelled  -  no folder selected." -ForegroundColor Yellow
        Write-Log -Message "Driver import (native) cancelled  -  no folder selected" -Level "WARN"
        return
    }

    Write-Host "  Source: $sourcePath" -ForegroundColor White
    Write-Log -Message "Driver import (native) started  -  source: $sourcePath"

    try {
        if (-not (Test-Path $sourcePath)) {
            Write-Host "  ERROR: Folder not found: $sourcePath" -ForegroundColor Red
            Write-Log -Message "Driver import failed: path not found  -  $sourcePath" -Level "ERROR"
            return
        }

        Write-Host "  Installing drivers from: $sourcePath" -ForegroundColor Cyan
        Write-Host "  Scanning for INF files..." -ForegroundColor DarkGray
        $infFiles = Get-ChildItem -Path $sourcePath -Filter '*.inf' -Recurse -ErrorAction SilentlyContinue
        if ($infFiles.Count -eq 0) {
            Write-Host "  WARNING: No INF files found in $sourcePath" -ForegroundColor Yellow
            Write-Log -Message "Driver import: no INF files found in $sourcePath" -Level "WARN"
            return
        }
        Write-Host "  Found $($infFiles.Count) INF file(s). Installing via pnputil..." -ForegroundColor Cyan
        Write-Host ""
        $okCount     = 0
        $rebootCount = 0
        $warnCount   = 0
        $total       = $infFiles.Count
        $i           = 0

        foreach ($inf in $infFiles) {
            $i++
            $rel = $inf.FullName.Replace($sourcePath, '').TrimStart('\/')
            Write-Progress -Activity "Installing drivers" `
                           -Status "$i of $total : $($inf.Name)" `
                           -PercentComplete (($i / $total) * 100)

            $null = pnputil /add-driver "`"$($inf.FullName)`"" /install 2>&1
            switch ($LASTEXITCODE) {
                0    { $okCount++;     Write-Log -Message "Driver installed: $rel" }
                3010 { $rebootCount++; Write-Log -Message "Driver installed (reboot required): $rel" }
                default {
                    $warnCount++
                    Write-Log -Message "Driver skipped (exit $LASTEXITCODE): $rel" -Level "WARN"
                }
            }
        }

        Write-Progress -Activity "Installing drivers" -Completed
        Write-Host "  Results:" -ForegroundColor White
        Write-Host "    Installed  : $okCount" -ForegroundColor Green
        if ($rebootCount -gt 0) {
            Write-Host "    Need reboot: $rebootCount" -ForegroundColor Yellow
        }
        Write-Host "    Skipped    : $warnCount" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  Driver import complete." -ForegroundColor Green
        Write-Log -Message "Drivers imported (native) from $sourcePath  -  OK:$okCount Reboot:$rebootCount Skipped:$warnCount"
        if ($okCount -gt 0 -or $rebootCount -gt 0) {
            Write-Host "  A system restart may be required for all drivers to take effect." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "  ERROR: Driver import failed  -  $_" -ForegroundColor Red
        Write-Log -Message "Driver import (native) failed  -  $_" -Level "ERROR"
    }
}

# ===========================================================================
# 7  -  DRIVER EXPORT (OSD MODULE)
# ===========================================================================

function Export-DriversOSD {
    [CmdletBinding()]
    param()

    Write-Host ""
    Write-Host "  Checking for OSD PowerShell module..." -ForegroundColor Cyan
    Write-Log -Message "Driver export (OSD) started"

    try {
        if (-not (Get-Module -ListAvailable -Name OSD)) {
            Write-Host ""
            Write-Host "  The OSD module is not installed." -ForegroundColor Yellow
            $confirm = Read-Host "  Install OSD module from PSGallery now? [Y/N]"
            if ($confirm -notmatch '^[Yy]') {
                Write-Host "  Cancelled." -ForegroundColor Yellow
                Write-Log -Message "OSD module install cancelled by user" -Level "WARN"
                return
            }

            Write-Host "  Installing OSD module..." -ForegroundColor Cyan
            Install-Module OSD -Force -Scope CurrentUser -ErrorAction Stop
            Write-Host "  OSD module installed." -ForegroundColor Green
            Write-Log -Message "OSD module installed from PSGallery"
        }

        Import-Module OSD -ErrorAction Stop

        # Show device product info
        Write-Host ""
        Write-Host "  Detecting device product ID..." -ForegroundColor Cyan
        $product = Get-MyComputerProduct
        Write-Host "  Computer Product ID: $product" -ForegroundColor White

        # Browse for destination
        Write-Host ""
        Write-Host "  Select the destination folder in the Explorer window..." -ForegroundColor Cyan
        Write-Host "  A 'Windows-DriverPack' subfolder will be created there." -ForegroundColor DarkGray
        Write-Host ""
        $root = Get-FolderFromBrowser -Description "Select destination for OSD driver pack download"
        if (-not $root) {
            Write-Host "  Cancelled  -  no folder selected." -ForegroundColor Yellow
            Write-Log -Message "OSD driver pack cancelled  -  no folder selected" -Level "WARN"
            return
        }

        $packPath = Join-Path $root "Windows-DriverPack"
        if (-not (Test-Path $packPath)) {
            New-Item -ItemType Directory -Path $packPath -Force | Out-Null
        }

        Write-Host "  Downloading driver pack to: $packPath" -ForegroundColor Cyan
        Write-Host "  Please wait  -  downloading (this may take several minutes)..." -ForegroundColor DarkGray

        $job = Start-Job -ScriptBlock {
            param($p) Import-Module OSD -ErrorAction SilentlyContinue; Save-MyDriverPack -DownloadPath $p
        } -ArgumentList $packPath

        $spin = '|','/','-','\'
        $i = 0
        while ($job.State -eq 'Running') {
            Write-Host "`r  [$($spin[$i % 4])] Downloading..." -NoNewline -ForegroundColor Cyan
            $i++
            Start-Sleep -Milliseconds 200
        }
        Receive-Job $job | Out-Null
        Remove-Job $job

        Write-Host "`r  Download complete.              " -ForegroundColor Green
        Write-Host ""
        Write-Host "  OSD driver pack saved successfully." -ForegroundColor Green
        Write-Host "  Saved to: $packPath" -ForegroundColor Green
        Write-Log -Message "OSD driver pack saved  -  Product: $product  -  Path: $packPath"
    } catch {
        Write-Host "  ERROR: OSD driver export failed  -  $_" -ForegroundColor Red
        Write-Log -Message "OSD driver export failed  -  $_" -Level "ERROR"
    }
}

# ===========================================================================
# 8  -  SMART DRIVER INSTALL FROM OSD PACK (HWID MATCHING)
# ===========================================================================

function Get-SystemHardwareIds {
    # Collects every Hardware ID and Compatible ID present on this device
    [CmdletBinding()]
    param()

    $hwids = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    $devices = Get-PnpDevice -ErrorAction SilentlyContinue
    $total   = ($devices | Measure-Object).Count
    $i       = 0

    foreach ($device in $devices) {
        $i++
        Write-Progress -Activity "Reading device hardware IDs" `
                       -Status "$i of $total  -  $($device.FriendlyName)" `
                       -PercentComplete (($i / $total) * 100)
        try {
            # Hardware IDs  -  most specific (VEN/DEV/SUBSYS/REV)
            $hProp = Get-PnpDeviceProperty -InstanceId $device.InstanceId `
                         -KeyName 'DEVPKEY_Device_HardwareIds' -ErrorAction SilentlyContinue
            if ($hProp -and $hProp.Data) {
                foreach ($id in $hProp.Data) {
                    if ($id -and $id.Trim()) { [void]$hwids.Add($id.Trim().ToUpper()) }
                }
            }
            # Compatible IDs  -  broader fallback matches
            $cProp = Get-PnpDeviceProperty -InstanceId $device.InstanceId `
                         -KeyName 'DEVPKEY_Device_CompatibleIds' -ErrorAction SilentlyContinue
            if ($cProp -and $cProp.Data) {
                foreach ($id in $cProp.Data) {
                    if ($id -and $id.Trim()) { [void]$hwids.Add($id.Trim().ToUpper()) }
                }
            }
        } catch { }
    }

    Write-Progress -Activity "Reading device hardware IDs" -Completed
    return $hwids
}

function Get-InfHardwareIds {
    # Parses an INF file and returns all hardware IDs it declares support for
    [CmdletBinding()]
    param([string]$InfPath)

    $hwids = [System.Collections.Generic.List[string]]::new()

    # Bus prefixes that identify hardware ID strings inside INF files
    $busPrefixes = @(
        'PCI\\', 'USB\\', 'USBSTOR\\', 'ACPI\\', 'HDAUDIO\\', 'HID\\',
        'ROOT\\', 'DISPLAY\\', 'SCSI\\', 'BTH\\', 'BTHENUM\\', 'SWD\\',
        'MF\\', 'SD\\', '1394\\', 'ISAPNP\\', 'PCMCIA\\', 'STORAGE\\',
        'MEDIA\\', 'NET\\', 'SMBUS\\', 'WUDF\\', 'FTDIBUS\\'
    )

    try {
        # INF files may be ANSI or UTF-16; try both
        $lines = $null
        try   { $lines = Get-Content -Path $InfPath -Encoding Default -ErrorAction Stop }
        catch { $lines = Get-Content -Path $InfPath -Encoding Unicode -ErrorAction Stop }

        foreach ($line in $lines) {
            $trimmed = $line.Trim()
            # Skip comments and blank lines
            if ($trimmed -eq '' -or $trimmed -match '^\s*;') { continue }

            # Split by comma  -  HWIDs appear as comma-separated tokens on install lines
            foreach ($part in ($trimmed -split ',')) {
                # Strip inline comments
                $token = (($part -split ';')[0]).Trim().ToUpper()
                if ($token -eq '') { continue }

                foreach ($prefix in $busPrefixes) {
                    if ($token -match "^$prefix") {
                        $hwids.Add($token)
                        break
                    }
                }
            }
        }
    } catch { }

    return $hwids
}

function Install-DriversFromOSDPack {
    [CmdletBinding()]
    param()

    Write-Host ""
    Write-Host "  === Smart Driver Install from OSD Pack ===" -ForegroundColor Cyan
    Write-Host "  Only INF drivers whose Hardware ID matches this device will be installed." -ForegroundColor DarkGray
    Write-Host "  EXE files are listed for manual review  -  not auto-installed." -ForegroundColor DarkGray
    Write-Host ""

    $stagingRoot = 'C:\Windows\Temp\SWSetup'

    # --- Check if staging folder already has extracted content ---
    $stagingInfs = Get-ChildItem -Path $stagingRoot -Filter '*.inf' -Recurse -ErrorAction SilentlyContinue
    if ($stagingInfs.Count -gt 0) {
        Write-Host "  Previously extracted content found:" -ForegroundColor Cyan
        Write-Host "  $stagingRoot  ($($stagingInfs.Count) INF files)" -ForegroundColor White
        Write-Host ""
        $useStaging = Read-Host "  Use existing extracted content? [Y/N]"
        if ($useStaging -match '^[Yy]') {
            Write-Host "  Using existing extracted content." -ForegroundColor Green
            Write-Host ""
            $packFolder = $stagingRoot
            Write-Log -Message "Smart driver install using existing staging content: $stagingRoot"
        } else {
            $packFolder = $null   # fall through to browse
        }
    } else {
        $packFolder = $null
    }

    # --- Browse for pack folder (only if not using staging) ---
    if (-not $packFolder) {
        Write-Host "  Select the driver pack folder in the Explorer window..." -ForegroundColor Cyan
        Write-Host ""

        $packFolder = Get-FolderFromBrowser -Description "Select the OSD driver pack folder"
        if (-not $packFolder) {
            Write-Host "  Cancelled  -  no folder selected." -ForegroundColor Yellow
            Write-Log -Message "Smart driver install cancelled  -  no folder selected" -Level "WARN"
            return
        }

        if (-not (Test-Path $packFolder)) {
            Write-Host "  ERROR: Folder not found: $packFolder" -ForegroundColor Red
            Write-Log -Message "Smart driver install failed: folder not found  -  $packFolder" -Level "ERROR"
            return
        }

        Write-Host ""
        Write-Log -Message "Smart driver install started  -  folder: $packFolder"

        # --- Extract any EXE packs to staging ---
        $exeInPack = Get-ChildItem -Path $packFolder -Filter '*.exe' -Recurse -ErrorAction SilentlyContinue
        if ($exeInPack.Count -gt 0) {
            Write-Host "  Found $($exeInPack.Count) EXE installer(s). Extracting to staging folder..." -ForegroundColor Cyan
            foreach ($exe in $exeInPack) {
                $exeStaging = Join-Path $stagingRoot ($exe.BaseName)
                if (Test-Path $exeStaging) {
                    Write-Host "  Already extracted: $($exe.Name)" -ForegroundColor DarkGray
                    Write-Log -Message "EXE already extracted (staging exists): $($exe.Name)"
                } else {
                    New-Item -ItemType Directory -Path $exeStaging -Force | Out-Null
                    Write-Host "  Extracting: $($exe.Name)..." -ForegroundColor Cyan
                    Write-Log -Message "Extracting EXE: $($exe.FullName) -> $exeStaging"
                    try {
                        $p = Start-Process -FilePath $exe.FullName -ArgumentList "/s /e /f `"$exeStaging`"" -Wait -PassThru -ErrorAction Stop
                        if ($p.ExitCode -ne 0) {
                            Start-Process -FilePath $exe.FullName -ArgumentList "-s -e -f `"$exeStaging`"" -Wait -ErrorAction SilentlyContinue
                        }
                    } catch {
                        Write-Host "  [WARN] Extraction failed for $($exe.Name): $_" -ForegroundColor Yellow
                        Write-Log -Message "EXE extraction failed: $($exe.Name)  -  $_" -Level "WARN"
                    }
                }
            }
            $packFolder = $stagingRoot
            Write-Host ""
        }
    }

    # --- Step 1: Collect system Hardware IDs ---
    Write-Host "  [1/4] Reading hardware IDs from this device..." -ForegroundColor Cyan
    $systemHwids = Get-SystemHardwareIds
    Write-Host "        $($systemHwids.Count) hardware IDs detected on this device." -ForegroundColor White
    Write-Log -Message "System hardware IDs collected: $($systemHwids.Count)"

    # --- Step 2: Scan pack folder ---
    Write-Host ""
    Write-Host "  [2/4] Scanning driver pack folder for INF and EXE files..." -ForegroundColor Cyan
    $infFiles = Get-ChildItem -Path $packFolder -Filter '*.inf' -Recurse -ErrorAction SilentlyContinue
    $exeFiles = Get-ChildItem -Path $packFolder -Filter '*.exe' -Recurse -ErrorAction SilentlyContinue
    Write-Host "        $($infFiles.Count) INF files found." -ForegroundColor White
    Write-Host "        $($exeFiles.Count) EXE file(s) found (listed at end for manual review)." -ForegroundColor White

    if ($infFiles.Count -eq 0) {
        Write-Host "  No INF files found in: $packFolder" -ForegroundColor Yellow
        Write-Log -Message "Smart driver install: no INF files found in $packFolder" -Level "WARN"
        return
    }

    # --- Step 3: Match INF files to this device's hardware ---
    Write-Host ""
    Write-Host "  [3/4] Matching INF files to this device's hardware IDs..." -ForegroundColor Cyan

    $matched   = [System.Collections.Generic.List[hashtable]]::new()
    $unmatched = [System.Collections.Generic.List[string]]::new()
    $i = 0

    foreach ($inf in $infFiles) {
        $i++
        Write-Progress -Activity "Matching INF files to hardware" `
                       -Status "$i of $($infFiles.Count): $($inf.Name)" `
                       -PercentComplete (($i / $infFiles.Count) * 100)

        $infHwids    = Get-InfHardwareIds -InfPath $inf.FullName
        $matchedHwid = $null

        foreach ($infHwid in $infHwids) {
            foreach ($sysHwid in $systemHwids) {
                # System HWID starts with the INF HWID  -  correct Windows driver matching behavior
                # e.g. system has PCI\VEN_10EC&DEV_8168&SUBSYS_84321043&REV_15
                #      INF lists  PCI\VEN_10EC&DEV_8168  (less specific, still valid)
                if ($sysHwid -like "$infHwid*") {
                    $matchedHwid = $infHwid
                    break
                }
            }
            if ($matchedHwid) { break }
        }

        $relPath = $inf.FullName.Replace($packFolder, '').TrimStart('\/')
        if ($matchedHwid) {
            $matched.Add(@{ Path = $inf.FullName; Name = $inf.Name; MatchedHwid = $matchedHwid; RelPath = $relPath })
        } else {
            $unmatched.Add($relPath)
        }
    }

    Write-Progress -Activity "Matching INF files to hardware" -Completed

    # --- Display results ---
    Write-Host ""
    Write-Host "  ---- Matched drivers (your hardware) ----" -ForegroundColor Green
    if ($matched.Count -eq 0) {
        Write-Host "  None found." -ForegroundColor Yellow
    } else {
        Write-Host "  $($matched.Count) driver(s) matched this device." -ForegroundColor Green
    }

    Write-Host ""
    Write-Host "  Summary: $($matched.Count) matched  /  $($unmatched.Count) not matched  /  $($exeFiles.Count) EXE(s)" -ForegroundColor White
    Write-Log -Message "Match results  -  Matched:$($matched.Count) Skipped:$($unmatched.Count) EXEs:$($exeFiles.Count)"

    if ($matched.Count -eq 0) {
        Write-Host "  No matching drivers to install." -ForegroundColor Yellow
        Write-Log -Message "Smart driver install: no matching INFs found" -Level "WARN"
        return
    }

    # --- Step 4: Install matched drivers ---
    Write-Host ""
    $confirm = Read-Host "  Install $($matched.Count) matched driver(s) now? [Y/N]"
    if ($confirm -notmatch '^[Yy]') {
        Write-Host "  Cancelled." -ForegroundColor Yellow
        Write-Log -Message "Smart driver install cancelled by user"
        return
    }

    Write-Host ""
    Write-Host "  [4/4] Installing $($matched.Count) matched driver(s) via pnputil..." -ForegroundColor Cyan
    Write-Host ""
    Write-Log -Message "Smart driver install: installing $($matched.Count) drivers"

    $okCount     = 0
    $rebootCount = 0
    $skipCount   = 0
    $total       = $matched.Count
    $i           = 0

    foreach ($m in $matched) {
        $i++
        Write-Progress -Activity "Installing matched drivers" `
                       -Status "$i of $total : $([System.IO.Path]::GetFileName($m.Path))" `
                       -PercentComplete (($i / $total) * 100)
        try {
            $null = pnputil /add-driver "`"$($m.Path)`"" /install 2>&1
            switch ($LASTEXITCODE) {
                0    { $okCount++;     Write-Log -Message "Driver installed: $($m.RelPath)  -  HWID: $($m.MatchedHwid)" }
                3010 { $rebootCount++; Write-Log -Message "Driver installed (reboot required): $($m.RelPath)" }
                123  { $skipCount++;   Write-Log -Message "Driver already in store (skipped): $($m.RelPath)" }
                default {
                    $skipCount++
                    Write-Log -Message "Driver skipped (exit $LASTEXITCODE): $($m.RelPath)" -Level "WARN"
                }
            }
        } catch {
            $skipCount++
            Write-Log -Message "Driver install error: $($m.RelPath)  -  $_" -Level "ERROR"
        }
    }

    Write-Progress -Activity "Installing matched drivers" -Completed
    Write-Host "  Results:" -ForegroundColor White
    Write-Host "    Installed  : $okCount" -ForegroundColor Green
    if ($rebootCount -gt 0) {
        Write-Host "    Need reboot: $rebootCount" -ForegroundColor Yellow
    }
    Write-Host "    Skipped    : $skipCount" -ForegroundColor DarkGray
    Write-Host ""
    Write-Log -Message "Smart driver install complete  -  OK:$okCount Reboot:$rebootCount Skipped:$skipCount"

    if ($okCount -gt 0 -or $rebootCount -gt 0) {
        Write-Host "  A system restart may be required for all drivers to take effect." -ForegroundColor Yellow
    }
}

# ===========================================================================
# 9 & 10  -  WINDOWS ACTIVATION
# ===========================================================================

function Export-BiosOemKey {
    [CmdletBinding()]
    param()

    Write-Host ""
    Write-Host "  Retrieving BIOS OEM product key..." -ForegroundColor Cyan
    Write-Log -Message "BIOS OEM key export started"

    try {
        $biosKey = (Get-CimInstance -Query 'SELECT * FROM SoftwareLicensingService').OA3xOriginalProductKey

        if (-not $biosKey -or $biosKey.Trim() -eq '') {
            Write-Host "  ERROR: No OEM product key found in BIOS/UEFI firmware." -ForegroundColor Red
            Write-Host "  This device may not have an embedded OEM key." -ForegroundColor Yellow
            Write-Log -Message "BIOS key export: no key found" -Level "ERROR"
            return
        }

        Write-Host "  Key found: $biosKey" -ForegroundColor Green
        Write-Host ""
        Write-Host "  Select the destination folder in the Explorer window..." -ForegroundColor Cyan
        Write-Host ""

        $root = Get-FolderFromBrowser -Description "Select destination folder for the key file"
        if (-not $root) {
            Write-Host "  Cancelled  -  no folder selected." -ForegroundColor Yellow
            Write-Log -Message "BIOS key export cancelled  -  no folder selected" -Level "WARN"
            return
        }

        $outFile = Join-Path $root "Windows-ActivationKey.txt"
        $lines   = @(
            "# Windows Activation Key  -  exported $(Get-Date -Format 'yyyy-MM-dd HH:mm') by Windows-Toolbox"
            "# Device : $env:COMPUTERNAME"
            "# Source : BIOS/UEFI OEM firmware (OA3xOriginalProductKey)"
            "#"
            "# To activate: use option [11] Activate from Key File"
            "#"
            $biosKey
        )
        [System.IO.File]::WriteAllLines($outFile, [string[]]$lines, $script:Utf8NoBom)

        Write-Host "  Saved to: $outFile" -ForegroundColor Green
        Write-Log -Message "BIOS OEM key exported to $outFile"
    } catch {
        Write-Host "  ERROR: Key export failed  -  $_" -ForegroundColor Red
        Write-Log -Message "BIOS key export failed  -  $_" -Level "ERROR"
    }
}

function Invoke-WindowsActivationFromFile {
    [CmdletBinding()]
    param()

    Write-Host ""
    Write-Host "  Select the key file in the Explorer window..." -ForegroundColor Cyan
    Write-Host ""

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Title            = "Select Windows activation key file"
        $dialog.Filter           = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        $dialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
        $result = $dialog.ShowDialog()
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
            Write-Host "  Cancelled  -  no file selected." -ForegroundColor Yellow
            Write-Log -Message "Activation from file cancelled  -  no file selected" -Level "WARN"
            return
        }
        $keyFile = $dialog.FileName
    } catch {
        $keyFile = Read-Host "  Enter full path to the key file"
    }

    if (-not (Test-Path $keyFile)) {
        Write-Host "  ERROR: File not found: $keyFile" -ForegroundColor Red
        Write-Log -Message "Activation from file failed: file not found  -  $keyFile" -Level "ERROR"
        return
    }

    Write-Host "  File: $keyFile" -ForegroundColor White
    Write-Log -Message "Activation from file started  -  source: $keyFile"

    try {
        # Extract first line matching the XXXXX-XXXXX-XXXXX-XXXXX-XXXXX key format
        $productKey = Get-Content $keyFile |
            Where-Object { $_ -match '^[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}$' } |
            Select-Object -First 1

        if (-not $productKey) {
            Write-Host "  ERROR: No valid product key found in the file." -ForegroundColor Red
            Write-Host "  Expected format: XXXXX-XXXXX-XXXXX-XXXXX-XXXXX" -ForegroundColor Yellow
            Write-Log -Message "Activation from file: no valid key found in $keyFile" -Level "ERROR"
            return
        }

        Write-Host "  Key found. Installing..." -ForegroundColor Cyan
        Write-Log -Message "Activating Windows with key from file: $keyFile"

        slmgr /ipk $productKey
        Start-Sleep -Seconds 3

        Write-Host "  Activating Windows online..." -ForegroundColor Cyan
        slmgr /ato

        Write-Host ""
        Write-Host "  Activation command sent." -ForegroundColor Green
        Write-Host "  Run option [13] to verify the result." -ForegroundColor White
        Write-Log -Message "Activation attempted with key from file"
    } catch {
        Write-Host "  ERROR: Activation failed  -  $_" -ForegroundColor Red
        Write-Log -Message "Activation from file failed  -  $_" -Level "ERROR"
    }
}

function Install-SingleDriver {
    [CmdletBinding()]
    param()

    Write-Host ""
    Write-Host "  Select the INF driver file in the Explorer window..." -ForegroundColor Cyan
    Write-Host ""

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Title            = "Select driver INF file"
        $dialog.Filter           = "Driver files (*.inf)|*.inf|All files (*.*)|*.*"
        $dialog.InitialDirectory = 'C:\Windows\Temp\SWSetup'
        if (-not (Test-Path $dialog.InitialDirectory)) {
            $dialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
        }
        $result = $dialog.ShowDialog()
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
            Write-Host "  Cancelled  -  no file selected." -ForegroundColor Yellow
            Write-Log -Message "Single driver install cancelled  -  no file selected" -Level "WARN"
            return
        }
        $infPath = $dialog.FileName
    } catch {
        $infPath = Read-Host "  Enter full path to the INF file"
    }

    if (-not (Test-Path $infPath)) {
        Write-Host "  ERROR: File not found: $infPath" -ForegroundColor Red
        Write-Log -Message "Single driver install failed: file not found  -  $infPath" -Level "ERROR"
        return
    }

    Write-Host "  Installing via pnputil..." -ForegroundColor Cyan
    Write-Log -Message "Single driver install started  -  $infPath"

    try {
        $null = pnputil /add-driver "`"$infPath`"" /install 2>&1
        switch ($LASTEXITCODE) {
            0 {
                Write-Host "  Driver installed successfully." -ForegroundColor Green
                Write-Log -Message "Single driver installed: $infPath"
            }
            3010 {
                Write-Host "  Driver installed. A restart is required to complete installation." -ForegroundColor Yellow
                Write-Log -Message "Single driver installed (reboot required): $infPath"
            }
            123 {
                Write-Host "  Driver is already present in the driver store  -  no change made." -ForegroundColor DarkGray
                Write-Log -Message "Single driver already in store: $infPath"
            }
            default {
                Write-Host "  WARNING: pnputil exited with code $LASTEXITCODE." -ForegroundColor Yellow
                Write-Log -Message "Single driver install exit $LASTEXITCODE : $infPath" -Level "WARN"
            }
        }
    } catch {
        Write-Host "  ERROR: Driver install failed  -  $_" -ForegroundColor Red
        Write-Log -Message "Single driver install error  -  $_" -Level "ERROR"
    }
}

function Get-ActivationStatus {
    [CmdletBinding()]
    param()

    Write-Host ""
    Write-Host "  Checking Windows activation status..." -ForegroundColor Cyan
    Write-Log -Message "Activation status check"

    try {
        $products = Get-CimInstance -ClassName SoftwareLicensingProduct |
            Where-Object { $_.name -match 'windows' -and $_.PartialProductKey }

        if (-not $products) {
            Write-Host "  Could not retrieve licensing information." -ForegroundColor Yellow
            Write-Log -Message "Activation check: no SoftwareLicensingProduct found" -Level "WARN"
            return
        }

        foreach ($p in $products) {
            $statusText = switch ($p.LicenseStatus) {
                1 { "Activated" }
                2 { "Out-of-Box Grace Period" }
                3 { "Out-of-Tolerance Grace Period" }
                4 { "Non-Genuine Grace Period" }
                5 { "Not Activated" }
                6 { "Extended Grace Period" }
                default { "Unknown (Code: $($p.LicenseStatus))" }
            }
            $color = if ($p.LicenseStatus -eq 1) { "Green" } else { "Yellow" }
            Write-Host ""
            Write-Host "  Product   : $($p.Name)" -ForegroundColor White
            Write-Host "  Status    : $statusText" -ForegroundColor $color
            Write-Host "  Partial Key: $($p.PartialProductKey)" -ForegroundColor White
        }

        Write-Host ""
        Write-Log -Message "Activation status retrieved"
    } catch {
        Write-Host "  ERROR: Could not check activation status  -  $_" -ForegroundColor Red
        Write-Log -Message "Activation status check failed  -  $_" -Level "ERROR"
    }
}

function Invoke-WindowsActivation {
    [CmdletBinding()]
    param()

    Write-Host ""
    Write-Host "  Checking current activation status..." -ForegroundColor Cyan
    Write-Log -Message "Windows activation started"

    try {
        $licenseStatus = Get-CimInstance -ClassName SoftwareLicensingProduct |
            Where-Object { $_.name -match 'windows' -and $_.PartialProductKey } |
            Select-Object -ExpandProperty LicenseStatus -First 1

        if ($licenseStatus -eq 1) {
            Write-Host "  Windows is already activated. No action needed." -ForegroundColor Green
            Write-Log -Message "Activation skipped  -  already activated"
            return
        }

        Write-Host "  Windows is not activated. Retrieving BIOS OEM key..." -ForegroundColor Yellow
        Write-Log -Message "Activation: Windows not activated, retrieving BIOS key"

        $biosKey = (Get-CimInstance -Query 'SELECT * FROM SoftwareLicensingService').OA3xOriginalProductKey

        if (-not $biosKey -or $biosKey.Trim() -eq '') {
            Write-Host "  ERROR: No OEM product key found in BIOS/UEFI firmware." -ForegroundColor Red
            Write-Host "  This device may not have an embedded OEM key." -ForegroundColor Yellow
            Write-Log -Message "Activation failed: no BIOS OEM key found" -Level "ERROR"
            return
        }

        Write-Host "  BIOS OEM key found. Installing key..." -ForegroundColor Cyan
        slmgr /ipk $biosKey
        Start-Sleep -Seconds 3

        Write-Host "  Activating Windows online..." -ForegroundColor Cyan
        slmgr /ato

        Write-Host ""
        Write-Host "  Activation command sent." -ForegroundColor Green
        Write-Host "  Run option [13] to verify the result." -ForegroundColor White
        Write-Log -Message "Activation attempted with BIOS OEM key"
    } catch {
        Write-Host "  ERROR: Activation failed  -  $_" -ForegroundColor Red
        Write-Log -Message "Activation failed  -  $_" -Level "ERROR"
    }
}

# ===========================================================================
# MAIN  -  INTERACTIVE MENU LOOP
# ===========================================================================

function Main {
    Write-Log -Message "=== Windows-Toolbox session started ==="

    $selectedIndex = 0
    $totalItems    = $script:MenuItems.Count
    $running       = $true

    while ($running) {

        Show-Menu -SelectedIndex $selectedIndex

        # Read one raw keypress  -  no Enter required
        $key = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        $vk  = [int]$key.VirtualKeyCode

        if ($vk -eq 38) {
            # Up arrow / Numpad8 with NumLock off
            if ($selectedIndex -gt 0) { $selectedIndex-- }
            else                      { $selectedIndex = $totalItems - 1 }
        }
        elseif ($vk -eq 40) {
            # Down arrow / Numpad2 with NumLock off
            if ($selectedIndex -lt ($totalItems - 1)) { $selectedIndex++ }
            else                                       { $selectedIndex = 0 }
        }
        elseif ($vk -eq 36) {
            # Home key  -  jump to first item
            $selectedIndex = 0
        }
        elseif ($vk -eq 35) {
            # End key  -  jump to last item
            $selectedIndex = $totalItems - 1
        }
        elseif ($vk -eq 13) {
            # Enter  -  invoke selected item
            $item = $script:MenuItems[$selectedIndex]
            if ($null -eq $item.Action) {
                $running = $false
            } else {
                Clear-Host
                Write-Host ''
                Write-Host "  $($item.Label)" -ForegroundColor Cyan
                Write-Host "  $('-' * ([Math]::Min([Console]::WindowWidth - 4, 72)))" -ForegroundColor DarkGray
                Write-Host ''
                & $item.Action
                Write-Log -Message "Executed: [$($item.Number)] $($item.Label)"
                Write-Host ''
                Read-Host '  Press Enter to return to the menu'
            }
        }
        elseif ($vk -eq 27) {
            # Escape  -  exit
            $running = $false
        }
        elseif (($vk -ge 48 -and $vk -le 57) -or ($vk -ge 96 -and $vk -le 105)) {
            # Direct number shortcut  -  top-row digits (VK 48-57) or numpad with NumLock on (VK 96-105)
            $numPressed = if ($vk -ge 96) { $vk - 96 } else { $vk - 48 }
            $found      = $null
            foreach ($mi in $script:MenuItems) {
                if ($mi.Number -eq $numPressed) { $found = $mi ; break }
            }
            if ($null -ne $found) {
                if ($null -eq $found.Action) {
                    $running = $false
                } else {
                    Clear-Host
                    Write-Host ''
                    Write-Host "  $($found.Label)" -ForegroundColor Cyan
                    Write-Host "  $('-' * ([Math]::Min([Console]::WindowWidth - 4, 72)))" -ForegroundColor DarkGray
                    Write-Host ''
                    & $found.Action
                    Write-Log -Message "Executed: [$($found.Number)] $($found.Label)"
                    Write-Host ''
                    Read-Host '  Press Enter to return to the menu'
                    # Re-sync selectedIndex to the item that was run
                    $j = 0
                    foreach ($mi in $script:MenuItems) {
                        if ($mi.Number -eq $found.Number) { $selectedIndex = $j ; break }
                        $j++
                    }
                }
            }
        }
    }

    Write-Host '  Goodbye.' -ForegroundColor Cyan
    Write-Log -Message "=== Windows-Toolbox session ended ==="
}

# Entry point
Main
