# SharePoint Data Resolver v2.2a

## Author: Jeffrey Bunnaman  
IT Specialist - Alpha Corporate Technologies

---

## Change Log:
- **20240906** - Added file hash verification for removing duplicate files.
- **20240907** - Additional logic added for invalid item names for SharePoint.
- **20250207** - Added additional logic to improve user experience.

---

## Overview

The **SharePoint Data Resolver** is a PowerShell script designed to help migrate data from on-premise file systems to Microsoft 365 SharePoint Online. This script resolves common issues encountered during migration by preparing the data for smooth transfer. 

### Key Features:
- Removes hidden file attributes
- Renames invalid item names for SharePoint compatibility
- Removes empty files and folders
- Removes duplicate files and archives
- Removes unsupported file types (e.g., `.exe`, `.dll`, etc.)
- Removes QuickBooks Desktop files
- Shortens file paths that exceed length limits

---

## Usage

1. **Source Path**: The folder where your data is currently stored (on-premise file system).
2. **Log File**: The path where the log file will be generated for tracking script execution results.

```powershell
$sourcePath = "C:\ALPHA\SharePoint Test"
$logFile = "C:\ALPHA\Logs\SharePoint_testlog.txt"
```

**Important Notes**:
- After running this script, use the **SharePoint Migration Tool**'s scan feature to ensure data compatibility.
- If errors persist, consult the log file and corrective action might be necessary.

---

## Script Execution Flow:

1. **Starting the Script**: 
   - Initializes the script and provides feedback during execution.
   - Clears or creates the log file based on existing status.
  
2. **File Processing**:
   - The script performs operations like removing hidden file attributes, resolving invalid item names, and checking for unsupported file types.

3. **Log Creation**:
   - Detailed logs are written for each operation (success or failure), making troubleshooting easier.

4. **File Operations**:
   - Files that don’t meet SharePoint compatibility (e.g., invalid names or unsupported types) are either renamed or removed.

---

## Functions

The script includes several functions to resolve specific issues commonly found during file migrations:

- **`Resolve-InvalidItemNames`**: Resolves invalid SharePoint item names by replacing illegal characters with underscores.
- **`Remove-EmptyItems`**: Removes files and folders with a size of 0 bytes.
- **`Remove-DuplicateArchives`**: Removes duplicate `.zip` archive files if the original unzipped content exists.
- **`Remove-DuplicateFiles`**: Detects and removes exact duplicate files using SHA256 file hash verification.
- **`Check-UnsupportedFileTypes`**: Removes files with extensions that are unsupported by SharePoint (e.g., `.exe`, `.dll`).
- **`Remove-QuickBooksFiles`**: Removes files related to QuickBooks Desktop (`.qbw`, `.qbb`, etc.).
- **`Check-PathLengthLimits`**: Shortens file paths that exceed the Windows path length limit (260 characters).

---

## Example Execution

```powershell
# Set the source directory and log file location
$sourcePath = "C:\ALPHA\SharePoint Test"
$logFile = "C:\ALPHA\Logs\SharePoint_testlog.txt"

# Execute the main script
Write-Host "Starting SharePoint Data Resolver..."
Start-Sleep -Seconds 2
# The script will begin processing as per the defined functions
```

---

## Requirements

- PowerShell v5.1 or later
- Administrative privileges may be required for file modifications.
- Ensure that the **SharePoint Migration Tool** is used after running the script for the final validation and migration steps.

---

## Troubleshooting

- **Log File**: Always check the log file (`$logFile`) for detailed errors or warnings during script execution.
- **Permissions**: Make sure the script has the necessary permissions to modify files and folders.
- **File Conflicts**: In case the script encounters any file conflicts (e.g., duplicate file names), it will log the details for further resolution.

---

## License

This script is developed by Jeffrey Bunnaman, IT Specialist at Alpha Corporate Technologies. It is provided **as-is** with no warranties.

---

Feel free to contribute or improve upon this script by opening issues or pull requests on the repository.