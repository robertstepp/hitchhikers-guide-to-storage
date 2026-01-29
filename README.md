# Hitchhiker's Guide to Storage 🚀

> *"Space is big. You just won't believe how vastly, hugely, mind-bogglingly big it is."*
> — Douglas Adams

...and so is your file share, apparently. This tool helps you figure out where it all went.

## What Is This?

A PowerShell script that scans SMB/CIFS shares using [NetApp XCP](https://xcp.netapp.com/) and generates a report showing:

- **Owner** of each first-level subdirectory
- **Path** to the directory (full UNC path)
- **Total Size** of everything beneath it
- **Newest File Age** — how recently the directory was modified

Perfect for answering questions like "Which user home folder is eating 500GB?" or "What folders haven't been touched in years?"

## Requirements

- Windows PowerShell 5.1+ or PowerShell Core 7+
- [NetApp XCP](https://xcp.netapp.com/) installed and licensed for SMB scanning
- Read access to the target share

## Usage

### Basic Scan

```powershell
.\Get-XCPFolderReport.ps1 -NetLocation "\\fileserver\share"
```

This scans the share, displays results in the console, saves a CSV report, and opens it in Excel.

### Scan a Subfolder

Want to see all user home directories? Point it at the users folder:

```powershell
.\Get-XCPFolderReport.ps1 -NetLocation "\\fileserver\home\users"
```

Output:
```
Owner              Path                                    TotalSize    NewestFileAge
-----              ----                                    ---------    -------------
DOMAIN\arthur.dent \\fileserver\home\users\arthur.dent    42.00 GB     2d5h
DOMAIN\ford.prefect \\fileserver\home\users\ford.prefect  1.21 GB      7y0d
DOMAIN\zaphod      \\fileserver\home\users\zaphod         999.99 GB    +0s
```

### Adjust Parallel Processes

Speed up large scans by increasing parallelism (default: 8, max: 61):

```powershell
.\Get-XCPFolderReport.ps1 -NetLocation "\\fileserver\share" -Parallel 16
```

### With Transcript Logging

```powershell
.\Get-XCPFolderReport.ps1 -NetLocation "\\fileserver\share" -LogToFile
```

Creates a full transcript log in the script directory for auditing or troubleshooting.

### Debug Mode

```powershell
.\Get-XCPFolderReport.ps1 -NetLocation "\\fileserver\share" -Debug
```

Shows all variable assignments and processing steps in real-time.

### Suppress Auto-Open (for Scheduled Tasks)

```powershell
.\Get-XCPFolderReport.ps1 -NetLocation "\\fileserver\share" -NoAutoOpen
```

Generates the report without opening files afterward—useful for automation.

## Parameters

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `-NetLocation` | Yes | — | UNC path to scan (e.g., `\\server\share`) |
| `-Parallel` | No | 8 | Number of parallel XCP processes (1-61) |
| `-LogToFile` | No | — | Enable transcript logging |
| `-NoAutoOpen` | No | — | Don't auto-open output files |
| `-Debug` | No | — | Show debug output |

## Output

### Console
A formatted table displaying Owner, Path, TotalSize, and NewestFileAge for each first-level subdirectory.

### CSV File
`<ShareName>_FolderReport_<timestamp>.csv` in the current directory, containing:
- `Owner` — Domain\User or SID who owns the directory
- `Path` — Full UNC path
- `TotalSize` — Human-readable size (e.g., "42.00 GB")
- `NewestFileAge` — Age of the most recently modified file (e.g., "7y0d", "2d5h", "+0s")

### Transcript Log (Optional)
`XCPFolderReport_<timestamp>.log` in the script directory when `-LogToFile` is specified.

## Performance Notes

This script is optimized for large-scale scans:

- **Stream Processing** — Lines are processed as they arrive from XCP, keeping memory usage constant regardless of share size
- **Pre-compiled Regex** — Pattern matching is compiled once at startup, not on every line
- **Hashtable Aggregation** — O(1) lookups instead of repeated filtering
- **Parallel Scanning** — Configurable parallelism for faster scans

Tested on shares with millions of files. Progress indicator displays every 10,000 lines.

## Why "Hitchhiker's Guide"?

Because when your manager asks why the file server is full and you have 47 TB of unstructured data to analyze, the worst thing you can do is panic.

Just run the script. Have a cup of tea. The answer might not be 42, but at least you'll know which folder to look at first.

*So long, and thanks for all the fish.* 🐬