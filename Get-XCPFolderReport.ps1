<#
.SYNOPSIS
    Generates a folder report from NetApp XCP scan showing first-level subdirectories with ownership and size information.

.DESCRIPTION
    This script uses NetApp XCP to scan an SMB share and generates a CSV report containing:
    - Owner of each first-level subdirectory
    - Full path to the directory
    - Total size of all files within and beneath that directory
    - Age of the newest file within each directory

    The script parses XCP's output format and aggregates file sizes for each top-level folder.

    Optimizations:
    - Stream processing: Lines processed as they arrive, minimizing memory usage
    - Pre-compiled regex: Pattern matching compiled once for faster multi-million line scans
    - Hashtable aggregation: O(1) lookups instead of repeated Where-Object filtering

.PARAMETER NetLocation
    The SMB path to scan in UNC format (e.g., \\server\share or \\192.168.1.100\data).
    This parameter is mandatory.

.PARAMETER Parallel
    Number of parallel XCP processes to use for scanning. Default is 8.
    Valid range: 1-61.

.PARAMETER LogToFile
    Enable transcript logging to a file in the script's directory.
    The log file will be named: XCPFolderReport_<timestamp>.log

.PARAMETER NoAutoOpen
    Suppress automatic opening of CSV (and transcript) files after completion.

.PARAMETER Debug
    Use the built-in -Debug common parameter to see verbose variable output during execution.

.INPUTS
    None. This script does not accept pipeline input.

.OUTPUTS
    - Console: Displays the folder report as a formatted table
    - CSV File: Creates <ShareName>_FolderReport_<timestamp>.csv in the current directory
    - Log File: (Optional) Creates transcript when -LogToFile is specified

.EXAMPLE
    .\Get-XCPFolderReport.ps1 -NetLocation "\\fileserver\projects"
    Scans the projects share with 8 parallel processes (default), generates a CSV report, and opens it.

.EXAMPLE
    .\Get-XCPFolderReport.ps1 -NetLocation "\\192.168.1.50\home" -Parallel 16 -LogToFile
    Scans with 16 parallel processes and transcript logging enabled.

.EXAMPLE
    .\Get-XCPFolderReport.ps1 -NetLocation "\\fileserver\projects\users" -Parallel 4 -Debug -NoAutoOpen
    Scans the users subfolder with 4 parallel processes, debug output, and no auto-open.

.NOTES
    Author: Converted from Bash script
    Requires: NetApp XCP installed and configured for SMB scanning
    Version: 2.1
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "SMB path to scan (e.g., \\server\share)")]
    [Alias("Path", "Location")]
    [string]$NetLocation,

    [Parameter(Mandatory = $false, HelpMessage = "Number of parallel XCP processes (default: 8)")]
    [ValidateRange(1, 61)]
    [int]$Parallel = 8,

    [Parameter(Mandatory = $false, HelpMessage = "Enable transcript logging to file")]
    [switch]$LogToFile,

    [Parameter(Mandatory = $false, HelpMessage = "Suppress auto-opening of output files")]
    [switch]$NoAutoOpen
)

#region Script Setup and Transcript Logging

# Get script directory for log file placement
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile = Join-Path $ScriptDir "XCPFolderReport_$Timestamp.log"

# Start transcript if logging enabled
if ($LogToFile) {
    $DebugPreference = "Continue"  # Ensure debug output goes to transcript without prompting
    Start-Transcript -Path $LogFile
    Write-Host "Transcript logging enabled: $LogFile" -ForegroundColor Cyan
}

Write-Debug "ScriptDir = $ScriptDir"
Write-Debug "Timestamp = $Timestamp"
Write-Debug "LogFile = $LogFile"
Write-Debug "NetLocation (input) = $NetLocation"

#endregion

#region Pre-compiled Regex Patterns

# Pre-compile regex patterns for performance on large scans
# These are compiled to IL code once, avoiding per-line pattern parsing overhead
$SizeParseRegex = [regex]::new('^([\d.]+)\s*(B|KiB|MiB|GiB|TiB|KB|MB|GB|TB)?$', 'Compiled, IgnoreCase')
$PlainNumberRegex = [regex]::new('^\d+$', 'Compiled')

# XCP line format: type owner size age path
# Example: f S-1-22-1-1000  65.1KiB   59d4h pi-hole\etc-pihole\pihole.toml
# Example: f S-1-22-1-1000      729   59d4h pi-hole\etc-pihole\tls_ca.crt (no unit = bytes)
# Example: f S-1-22-1-992   40.2MiB    +0s  cluster\gitlab\... (+ prefix for very recent files)
# Fields are separated by varying whitespace, so we match by known patterns:
# - Type: single char (d or f)
# - Owner: non-whitespace (SID, DOMAIN\user, or username)
# - Size: number + optional unit (e.g., 65.1KiB, 1.5GiB, 729)
# - Age: optional + prefix, then number + time unit (e.g., 59d4h, 7y0d, 1h30m, +0s)
# - Path: everything remaining
$XCPLineRegex = [regex]::new(
    '^(?<type>[df])\s+(?<owner>\S+)\s+(?<size>[\d.]+[KMGTP]?i?B?)\s+(?<age>\+?\d+\w+)\s+(?<path>.+)$',
    'Compiled, IgnoreCase'
)

Write-Debug "Pre-compiled regex patterns initialized"

#endregion

#region Helper Functions

# Function to convert human-readable size to bytes
# Uses pre-compiled regex for performance
function Convert-SizeToBytes {
    param([string]$SizeString)

    $match = $SizeParseRegex.Match($SizeString)
    if ($match.Success) {
        $Number = [double]$match.Groups[1].Value
        $Unit = $match.Groups[2].Value.ToUpper()

        switch ($Unit) {
            'B'   { return [long]$Number }
            'KIB' { return [long]($Number * 1024) }
            'KB'  { return [long]($Number * 1024) }
            'MIB' { return [long]($Number * 1024 * 1024) }
            'MB'  { return [long]($Number * 1024 * 1024) }
            'GIB' { return [long]($Number * 1024 * 1024 * 1024) }
            'GB'  { return [long]($Number * 1024 * 1024 * 1024) }
            'TIB' { return [long]($Number * 1024 * 1024 * 1024 * 1024) }
            'TB'  { return [long]($Number * 1024 * 1024 * 1024 * 1024) }
            default { return [long]$Number }
        }
    }

    # If no unit or unrecognized, try to parse as plain number
    if ($PlainNumberRegex.IsMatch($SizeString)) {
        return [long]$SizeString
    }

    return 0
}

# Function to convert bytes back to human-readable format
function Convert-BytesToSize {
    param([long]$Bytes)

    if ($Bytes -ge 1TB) {
        return "{0:N2} TB" -f ($Bytes / 1TB)
    }
    elseif ($Bytes -ge 1GB) {
        return "{0:N2} GB" -f ($Bytes / 1GB)
    }
    elseif ($Bytes -ge 1MB) {
        return "{0:N2} MB" -f ($Bytes / 1MB)
    }
    elseif ($Bytes -ge 1KB) {
        return "{0:N2} KB" -f ($Bytes / 1KB)
    }
    else {
        return "$Bytes B"
    }
}

# Function to convert XCP age string to total seconds for comparison
# Examples: 59d4h, 7y0d, 1h30m, +0s, 5m20s
function Convert-AgeToSeconds {
    param([string]$AgeString)

    # Remove leading + if present
    $AgeString = $AgeString.TrimStart('+')

    $totalSeconds = [long]0

    # Match patterns like 7y, 59d, 4h, 30m, 20s
    $yearMatch = [regex]::Match($AgeString, '(\d+)y')
    $dayMatch = [regex]::Match($AgeString, '(\d+)d')
    $hourMatch = [regex]::Match($AgeString, '(\d+)h')
    $minMatch = [regex]::Match($AgeString, '(\d+)m')
    $secMatch = [regex]::Match($AgeString, '(\d+)s')

    if ($yearMatch.Success) {
        $totalSeconds += [long]$yearMatch.Groups[1].Value * 365 * 24 * 60 * 60
    }
    if ($dayMatch.Success) {
        $totalSeconds += [long]$dayMatch.Groups[1].Value * 24 * 60 * 60
    }
    if ($hourMatch.Success) {
        $totalSeconds += [long]$hourMatch.Groups[1].Value * 60 * 60
    }
    if ($minMatch.Success) {
        $totalSeconds += [long]$minMatch.Groups[1].Value * 60
    }
    if ($secMatch.Success) {
        $totalSeconds += [long]$secMatch.Groups[1].Value
    }

    return $totalSeconds
}

# Function to convert seconds back to human-readable age format
function Convert-SecondsToAge {
    param([long]$Seconds)

    if ($Seconds -ge (365 * 24 * 60 * 60)) {
        $years = [math]::Floor($Seconds / (365 * 24 * 60 * 60))
        $days = [math]::Floor(($Seconds % (365 * 24 * 60 * 60)) / (24 * 60 * 60))
        return "{0}y{1}d" -f $years, $days
    }
    elseif ($Seconds -ge (24 * 60 * 60)) {
        $days = [math]::Floor($Seconds / (24 * 60 * 60))
        $hours = [math]::Floor(($Seconds % (24 * 60 * 60)) / (60 * 60))
        return "{0}d{1}h" -f $days, $hours
    }
    elseif ($Seconds -ge (60 * 60)) {
        $hours = [math]::Floor($Seconds / (60 * 60))
        $mins = [math]::Floor(($Seconds % (60 * 60)) / 60)
        return "{0}h{1}m" -f $hours, $mins
    }
    elseif ($Seconds -ge 60) {
        $mins = [math]::Floor($Seconds / 60)
        $secs = $Seconds % 60
        return "{0}m{1}s" -f $mins, $secs
    }
    else {
        return "{0}s" -f $Seconds
    }
}

# Function to extract the first-level directory from a relative XCP path
# XCP returns paths relative to scan location, e.g., "pi-hole\etc-pihole\file.txt"
# where "pi-hole" is the scanned folder and "etc-pihole" is a first-level subdir
# We want to return "pi-hole\etc-pihole" for aggregation
function Get-FirstLevelDir {
    param(
        [string]$RelativePath
    )

    # Find the first and second backslash positions
    $firstSlash = $RelativePath.IndexOf('\')

    # If no backslash, this is a file/dir in the root of the scanned folder
    # We don't aggregate these (or it's the root folder itself)
    if ($firstSlash -eq -1) {
        return $null
    }

    $secondSlash = $RelativePath.IndexOf('\', $firstSlash + 1)

    # If no second backslash, this path IS a first-level item (e.g., "pi-hole\etc-pihole")
    if ($secondSlash -eq -1) {
        return $RelativePath
    }

    # Otherwise, extract up to the second backslash (e.g., "pi-hole\etc-pihole" from "pi-hole\etc-pihole\file.txt")
    return $RelativePath.Substring(0, $secondSlash)
}

#endregion

#region Input Validation

# Remove trailing backslash if present for consistency
$NetLocation = $NetLocation.TrimEnd('\')
Write-Debug "NetLocation (trimmed) = $NetLocation"

# Validate UNC path format (must start with \\)
if ($NetLocation -notmatch '^\\\\[^\\]+\\[^\\]+') {
    Write-Error "Error: Path should be in UNC format (\\server\share)"
    Write-Error "Example: \\fileserver\data or \\192.168.1.100\home"
    Write-Debug "VALIDATION FAILED: Invalid UNC path format"
    if ($LogToFile) { Stop-Transcript }
    exit 1
}

Write-Debug "UNC path validation passed"

# Extract server and share components for later use
# UNC format: \\server\share\optional\subpath
$PathParts = $NetLocation -replace '^\\\\', '' -split '\\'
$Server = $PathParts[0]
$ShareName = $PathParts[1]
# The base share path (just \\server\share)
$BaseSharePath = "\\$Server\$ShareName"
# Any subpath beneath the share
$SubPath = if ($PathParts.Count -gt 2) { ($PathParts[2..($PathParts.Count-1)] -join '\') } else { "" }

Write-Debug "Server = $Server"
Write-Debug "ShareName = $ShareName"
Write-Debug "BaseSharePath = $BaseSharePath"
Write-Debug "SubPath = $SubPath"

#endregion

#region XCP Detection

# Common XCP installation paths - check in order of likelihood
$XCPPaths = @(
    "C:\Program Files\NetApp\XCP\xcp.exe",
    "C:\NetApp\XCP\xcp.exe",
    "C:\Program Files (x86)\NetApp\XCP\xcp.exe",
    "$env:LOCALAPPDATA\NetApp\XCP\xcp.exe"
)

Write-Debug "Searching for XCP executable..."

$XCP = $null
foreach ($Path in $XCPPaths) {
    Write-Debug "Checking path: $Path"
    if (Test-Path $Path) {
        $XCP = $Path
        Write-Debug "XCP found at: $Path"
        break
    }
}

# Also check if xcp is in PATH
if (-not $XCP) {
    $XCPInPath = Get-Command "xcp.exe" -ErrorAction SilentlyContinue
    if ($XCPInPath) {
        $XCP = $XCPInPath.Source
        Write-Debug "XCP found in PATH: $XCP"
    }
}

if (-not $XCP) {
    Write-Error "Error: XCP executable not found"
    Write-Error "Searched locations:"
    $XCPPaths | ForEach-Object { Write-Error "  - $_" }
    Write-Error "Ensure NetApp XCP is installed or add it to your PATH"
    Write-Debug "FATAL: XCP executable not found"
    if ($LogToFile) { Stop-Transcript }
    exit 1
}

Write-Debug "XCP = $XCP"

#endregion

#region XCP Scan and Stream Processing

Write-Host "Scanning $NetLocation with XCP (this may take a moment)..." -ForegroundColor Yellow
Write-Debug "Starting XCP scan with stream processing..."

# Build XCP command
# XCP SMB scan format: xcp scan -l -ownership -parallel # \\server\share
# Output format: type  owner  size  age  path
$XCPArgs = @("scan", "-l", "-ownership", "-parallel", $Parallel, $NetLocation)
Write-Debug "XCPArgs = $($XCPArgs -join ' ')"

# Initialize hashtable for aggregating folder stats
# Key: first-level directory path
# Value: PSCustomObject with Owner and TotalBytes
# Using hashtable gives O(1) lookups vs O(n) for repeated Where-Object
$FolderStats = @{}

$LineCount = 0
$ProcessedCount = 0
$ErrorLines = @()

try {
    # Stream processing: process each line as it arrives from XCP
    # This reduces memory footprint from O(n) to O(1) for raw output
    # and allows aggregation to begin immediately
    & $XCP $XCPArgs 2>&1 | ForEach-Object {
        $Line = $_
        $LineCount++

        # Skip empty lines
        if ([string]::IsNullOrWhiteSpace($Line)) { return }

        # Check for error output (ErrorRecord objects from 2>&1)
        if ($Line -is [System.Management.Automation.ErrorRecord]) {
            $ErrorLines += $Line.ToString()
            return
        }

        # Parse line using regex that matches XCP output format
        # Format: type owner size age path
        $Match = $XCPLineRegex.Match($Line)

        if (-not $Match.Success) {
            Write-Debug "Skipping line $LineCount (not a data line): $Line"
            return
        }

        $ItemType = $Match.Groups['type'].Value
        $Owner = $Match.Groups['owner'].Value
        $SizeStr = $Match.Groups['size'].Value
        $AgeStr = $Match.Groups['age'].Value
        $ItemPath = $Match.Groups['path'].Value

        # Determine which first-level directory this item belongs to
        $FirstLevelDir = Get-FirstLevelDir -RelativePath $ItemPath

        if ($FirstLevelDir) {
            # Convert size to bytes
            $SizeBytes = Convert-SizeToBytes -SizeString $SizeStr

            # Convert age to seconds for comparison (smaller = newer)
            $AgeSeconds = Convert-AgeToSeconds -AgeString $AgeStr

            # Aggregate into hashtable
            if ($FolderStats.ContainsKey($FirstLevelDir)) {
                # Add to existing total
                $FolderStats[$FirstLevelDir].TotalBytes += $SizeBytes

                # Track newest file (smallest age)
                if ($AgeSeconds -lt $FolderStats[$FirstLevelDir].NewestAgeSeconds) {
                    $FolderStats[$FirstLevelDir].NewestAgeSeconds = $AgeSeconds
                    $FolderStats[$FirstLevelDir].NewestAgeStr = $AgeStr
                }
            }
            else {
                # First time seeing this directory - initialize entry
                # For owner, use the owner of the directory itself (type 'd' with exact path match)
                # or the first file's owner as fallback
                $FolderStats[$FirstLevelDir] = [PSCustomObject]@{
                    Owner            = $Owner
                    TotalBytes       = $SizeBytes
                    IsDir            = ($ItemType -eq 'd' -and $ItemPath -eq $FirstLevelDir)
                    NewestAgeSeconds = $AgeSeconds
                    NewestAgeStr     = $AgeStr
                }
                Write-Debug "New first-level dir: $FirstLevelDir (Owner: $Owner)"
            }

            # If this IS the directory entry itself, update owner to be accurate
            if ($ItemType -eq 'd' -and $ItemPath -eq $FirstLevelDir) {
                $FolderStats[$FirstLevelDir].Owner = $Owner
                $FolderStats[$FirstLevelDir].IsDir = $true
            }

            $ProcessedCount++
        }

        # Progress indicator every 10000 lines
        if ($LineCount % 10000 -eq 0) {
            Write-Host "  Processed $LineCount lines..." -ForegroundColor DarkGray
        }
    }

    $ExitCode = $LASTEXITCODE
    Write-Debug "XCP ExitCode = $ExitCode"
    Write-Debug "Total lines read = $LineCount"
    Write-Debug "Items processed = $ProcessedCount"
    Write-Debug "First-level directories found = $($FolderStats.Count)"

    if ($ExitCode -ne 0 -and $FolderStats.Count -eq 0) {
        Write-Error "Error: XCP scan failed with exit code $ExitCode"
        if ($ErrorLines.Count -gt 0) {
            Write-Error "XCP Errors:"
            $ErrorLines | ForEach-Object { Write-Error "  $_" }
        }
        Write-Debug "FATAL: XCP scan failed with no results"
        if ($LogToFile) { Stop-Transcript }
        exit 1
    }
}
catch {
    Write-Error "Error executing XCP: $_"
    Write-Debug "FATAL: Exception during XCP execution: $_"
    if ($LogToFile) { Stop-Transcript }
    exit 1
}

Write-Host "Processing complete. Found $($FolderStats.Count) first-level directories." -ForegroundColor Yellow

#endregion

#region Build Results

Write-Debug "Building results from aggregated data..."

# Get the parent of the scanned path for building full UNC paths
# XCP paths start with the last component of NetLocation, so we need the parent
# Handle edge case where NetLocation is the root of a share (no parent)
$NetLocationParent = Split-Path -Parent $NetLocation
if ([string]::IsNullOrEmpty($NetLocationParent)) {
    # NetLocation is at share root, so the XCP relative paths ARE the full paths under the share
    # Just prepend the NetLocation itself
    $NetLocationParent = $NetLocation
}

Write-Debug "NetLocationParent = $NetLocationParent"

# Convert hashtable to sorted results array
$Results = $FolderStats.GetEnumerator() | Sort-Object -Property Name | ForEach-Object {
    $TotalSize = Convert-BytesToSize -Bytes $_.Value.TotalBytes

    # Build full UNC path: parent + relative first-level dir
    # e.g., "\\server\share" + "pi-hole\etc-pihole" = "\\server\share\pi-hole\etc-pihole"
    $FullPath = Join-Path $NetLocationParent $_.Name

    [PSCustomObject]@{
        Owner           = $_.Value.Owner
        Path            = $FullPath
        TotalSize       = $TotalSize
        NewestFileAge   = $_.Value.NewestAgeStr
        SizeBytes       = $_.Value.TotalBytes
        NewestAgeSeconds = $_.Value.NewestAgeSeconds
    }
}

# Handle case where Results might be a single object (not array)
if ($Results -and $Results -isnot [array]) {
    $Results = @($Results)
}

Write-Debug "Total Results = $($Results.Count)"

#endregion

#region Output Results

# Generate output filenames
$CSVFileName = "${ShareName}_FolderReport_$Timestamp.csv"
$CSVPath = Join-Path (Get-Location) $CSVFileName

Write-Debug "CSVPath = $CSVPath"

# Export to CSV (select only the columns we want, in order)
$Results | Select-Object Owner, Path, TotalSize, NewestFileAge | Export-Csv -Path $CSVPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host ("=" * 80) -ForegroundColor Green
Write-Host "XCP Folder Report - $NetLocation" -ForegroundColor Green
Write-Host ("=" * 80) -ForegroundColor Green
Write-Host ""

# Display to console as formatted table
if ($Results.Count -gt 0) {
    $Results | Select-Object Owner, Path, TotalSize, NewestFileAge | Format-Table -AutoSize
}
else {
    Write-Host "No first-level subdirectories found." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "CSV Report saved to: $CSVPath" -ForegroundColor Cyan

if ($LogToFile) {
    Write-Host "Transcript log saved to: $LogFile" -ForegroundColor Cyan
    Stop-Transcript
}

Write-Host "Done." -ForegroundColor Green

#endregion

#region Auto-Open Output Files

if (-not $NoAutoOpen) {
    Write-Debug "Opening output files with default applications..."

    # Open CSV in default application (typically Excel)
    try {
        Invoke-Item -Path $CSVPath
        Write-Debug "Opened CSV: $CSVPath"
    }
    catch {
        Write-Warning "Could not auto-open CSV file: $_"
    }

    # Open transcript log if it was created
    if ($LogToFile -and (Test-Path $LogFile)) {
        try {
            Invoke-Item -Path $LogFile
            Write-Debug "Opened transcript: $LogFile"
        }
        catch {
            Write-Warning "Could not auto-open transcript file: $_"
        }
    }
}

#endregion