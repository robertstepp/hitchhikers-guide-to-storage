<#
.SYNOPSIS
    Generates a folder report from NetApp XCP scan showing first-level subdirectories with ownership and size information.

.DESCRIPTION
    This script uses NetApp XCP to scan an SMB share and generates a CSV report containing:
    - Owner of each first-level subdirectory
    - Full path to the directory
    - Total size of all files within and beneath that directory

    The script parses XCP's output format and aggregates file sizes for each top-level folder.

    Optimizations:
    - Stream processing: Lines processed as they arrive, minimizing memory usage
    - Pre-compiled regex: Pattern matching compiled once for faster multi-million line scans
    - Hashtable aggregation: O(1) lookups instead of repeated Where-Object filtering

.PARAMETER NetLocation
    The SMB path to scan in UNC format (e.g., \\server\share or \\192.168.1.100\data).
    This parameter is mandatory.

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
    Scans the projects share, generates a CSV report, and opens it in the default application.

.EXAMPLE
    .\Get-XCPFolderReport.ps1 -NetLocation "\\192.168.1.50\home" -LogToFile
    Scans with transcript logging enabled, opens both CSV and log file after completion.

.EXAMPLE
    .\Get-XCPFolderReport.ps1 -NetLocation "\\fileserver\projects\users" -Debug -NoAutoOpen
    Scans the users subfolder with debug output, does not auto-open files.

.NOTES
    Author: Converted from Bash script
    Requires: NetApp XCP installed and configured for SMB scanning
    Version: 2.0
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "SMB path to scan (e.g., \\server\share)")]
    [Alias("Path", "Location")]
    [string]$NetLocation,

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
$SkipLineRegex = [regex]::new('^\s*(scanned|XCP|---|\d+\s+scanned)', 'Compiled, IgnoreCase')
$SplitRegex = [regex]::new('\s{2,}|\t', 'Compiled')
$SizeParseRegex = [regex]::new('^([\d.]+)\s*(B|KiB|MiB|GiB|TiB|KB|MB|GB|TB)$', 'Compiled, IgnoreCase')
$PlainNumberRegex = [regex]::new('^\d+$', 'Compiled')

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

# Function to extract the first-level directory from a full path
# Given \\server\share\users\jsmith\docs, with base \\server\share\users,
# returns \\server\share\users\jsmith
function Get-FirstLevelDir {
    param(
        [string]$FullPath,
        [int]$TargetDepth
    )

    # Count backslashes to find depth
    $slashCount = 0
    $cutoffIndex = -1

    for ($i = 0; $i -lt $FullPath.Length; $i++) {
        if ($FullPath[$i] -eq '\') {
            $slashCount++
            if ($slashCount -eq $TargetDepth) {
                $cutoffIndex = $i
            }
            elseif ($slashCount -gt $TargetDepth) {
                # Path is deeper than first level, truncate to first-level dir
                return $FullPath.Substring(0, $cutoffIndex)
            }
        }
    }

    # If we reach here with exact target depth, return the full path (it IS a first-level dir)
    if ($slashCount -eq $TargetDepth) {
        return $FullPath
    }

    # Path is shallower than target depth (shouldn't happen for valid items)
    return $null
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

# Calculate base depth by counting backslashes in the scan path
# \\server\share = 3 backslashes, so first-level dirs have 4 backslashes
$BaseDepth = ($NetLocation.ToCharArray() | Where-Object { $_ -eq '\' }).Count
$TargetDepth = $BaseDepth + 1

Write-Debug "BaseDepth = $BaseDepth"
Write-Debug "TargetDepth = $TargetDepth"

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
# XCP SMB scan format: xcp scan -l -ownership \\server\share
# Output format: type  owner  size  age  path
$XCPArgs = @("scan", "-l", "-ownership", $NetLocation)
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

        # Skip summary/header lines using pre-compiled regex
        if ($SkipLineRegex.IsMatch($Line)) {
            Write-Debug "Skipping line $LineCount (header/summary)"
            return
        }

        # Check for error output (ErrorRecord objects from 2>&1)
        if ($Line -is [System.Management.Automation.ErrorRecord]) {
            $ErrorLines += $Line.ToString()
            return
        }

        # Split line using pre-compiled regex
        # Format: Type  Owner  Size  Age  Path
        $Parts = $SplitRegex.Split($Line) | Where-Object { $_ -ne '' }

        if ($Parts.Count -ge 5) {
            $ItemType = $Parts[0]
            $Owner = $Parts[1]
            $SizeStr = $Parts[2]
            # $Age = $Parts[3]  # Not used in current report
            # Path may contain spaces, so join remaining parts
            $ItemPath = ($Parts[4..($Parts.Count-1)] -join ' ')

            # Determine which first-level directory this item belongs to
            $FirstLevelDir = Get-FirstLevelDir -FullPath $ItemPath -TargetDepth $TargetDepth

            if ($FirstLevelDir) {
                # Convert size to bytes
                $SizeBytes = Convert-SizeToBytes -SizeString $SizeStr

                # Aggregate into hashtable
                if ($FolderStats.ContainsKey($FirstLevelDir)) {
                    # Add to existing total
                    $FolderStats[$FirstLevelDir].TotalBytes += $SizeBytes
                }
                else {
                    # First time seeing this directory - initialize entry
                    # For owner, use the owner of the directory itself (type 'd' with exact path match)
                    # or the first file's owner as fallback
                    $FolderStats[$FirstLevelDir] = [PSCustomObject]@{
                        Owner      = $Owner
                        TotalBytes = $SizeBytes
                        IsDir      = ($ItemType -eq 'd' -and $ItemPath -eq $FirstLevelDir)
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

# Convert hashtable to sorted results array
$Results = $FolderStats.GetEnumerator() | Sort-Object -Property Name | ForEach-Object {
    $TotalSize = Convert-BytesToSize -Bytes $_.Value.TotalBytes

    [PSCustomObject]@{
        Owner     = $_.Value.Owner
        Path      = $_.Name
        SizeBytes = $_.Value.TotalBytes
        TotalSize = $TotalSize
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
$Results | Select-Object Owner, Path, TotalSize | Export-Csv -Path $CSVPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host ("=" * 80) -ForegroundColor Green
Write-Host "XCP Folder Report - $NetLocation" -ForegroundColor Green
Write-Host ("=" * 80) -ForegroundColor Green
Write-Host ""

# Display to console as formatted table
if ($Results.Count -gt 0) {
    $Results | Select-Object Owner, Path, TotalSize | Format-Table -AutoSize
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