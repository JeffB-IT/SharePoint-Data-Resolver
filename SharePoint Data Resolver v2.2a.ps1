# SharePoint Data Resolver v2.2a
# Author: Jeffrey Bunnaman, IT Specialist - Alpha Corporate Technologies
# 
# CHANGE LOG:
# 20240906 - Added file hash verification for removing duplicate files
# 20240907 - Additional logic added for invalid item names for SharePoint
# 20250207 - Added additional logic to improve user experience.

# This PowerShell script will prepare data for migration from on-premise to Microsoft 365 SharePoint Online.
# Please note: Utilize SharePoint Migration scan function to scan source data once script execution is complete; if errors persist, consult scan log CSV downloaded from SharePoint Admin Center and take corrective action if necessary.


# -----------------------FUNCTIONS-----------------------

# Function to resolve invalid item names
function Resolve-InvalidItemNames {
    param (
        [string]$sourcePath,
        [string]$logFile
    )
    Write-Host "Resolving invalid item names..." -ForegroundColor Green
    Start-sleep -seconds 2
    # Logic to rename invalid items
    Get-ChildItem -Path $sourcePath -Recurse | ForEach-Object {
        if ($_ -match '[*:"<>?|/\\]') {
            $newName = $_.Name -replace '[*:"<>?|/\\]', '_'
            try {
                Rename-Item -Path $_.FullName -NewName $newName -ErrorAction Stop
                Write-Host "Renamed item: $($_.FullName) to $newName"
            } catch {
                Add-Content -Path $logFile -Value "Failed to rename item: $($_.FullName) - Error: $_"
                Write-Host "Failed to rename item: $($_.FullName)"
            }
        }
    }
}

# Function to remove empty items
function Remove-EmptyItems {
    param (
        [string]$sourcePath,
        [string]$logFile
    )
    Write-Host "Removing empty items..." -ForegroundColor Green
    Start-sleep -seconds 2
    Get-ChildItem -Path $sourcePath -Recurse -File | Where-Object { $_.Length -eq 0 } | ForEach-Object {
        try {
            Remove-Item -Path $_.FullName -ErrorAction Stop
            Write-Host "Removed empty item: $($_.FullName)"
        } catch {
            Add-Content -Path $logFile -Value "Failed to remove empty item: $($_.FullName) - Error: $_"
            Write-Host "Failed to remove empty item: $($_.FullName)"
        }
    }
}

# Function to remove duplicate archives
function Remove-DuplicateArchives {
    param (
        [string]$sourcePath,
        [string]$logFile
    )
    Write-Host "Removing duplicate archives..." -ForegroundColor Green
    Start-sleep -seconds 2
    Get-ChildItem -Path $sourcePath -Recurse -Filter "*.zip" | ForEach-Object {
        $originalFile = $_.FullName -replace '\.zip$', ''
        if (Test-Path -Path $originalFile) {
            try {
                Remove-Item -Path $_.FullName -ErrorAction Stop
                Add-Content -Path $logFile -Value "Removed duplicate archive: $($_.FullName)"
                Write-Host "Removed duplicate archive: $($_.FullName)"
            } catch {
                Add-Content -Path $logFile -Value "Failed to remove duplicate archive: $($_.FullName) - Error: $_"
                Write-Host "Failed to remove duplicate archive: $($_.FullName)"
            }
        }
    }
}

# Function to compute file hash
function Get-FileHash {
    param (
        [string]$filePath
    )
    Write-Host "Verifying file hash..." -ForegroundColor Green
    Start-sleep -seconds 2
    try {
        $hashAlgorithm = [System.Security.Cryptography.HashAlgorithm]::Create("SHA256")
        $fileStream = [System.IO.File]::OpenRead($filePath)
        $hashBytes = $hashAlgorithm.ComputeHash($fileStream)
        $fileStream.Close()
        $hashAlgorithm.Dispose()
        return [BitConverter]::ToString($hashBytes) -replace '-'
    } catch {
        Write-Host "Failed to compute hash for file: $filePath - Error: $_"
        return $null
    }
}

# Function to remove exact duplicate files
function Remove-DuplicateFiles {
    param (
        [string]$sourcePath,
        [string]$logFile
    )
    Write-Host "Removing exact duplicate files & verifying via Hash..." -ForegroundColor Green
    Start-sleep -seconds 2
    $fileHashes = @{}
    Get-ChildItem -Path $sourcePath -Recurse -File | ForEach-Object {
        $fileHash = Get-FileHash -filePath $_.FullName
        if ($fileHash) {
            if ($fileHashes.ContainsKey($fileHash)) {
                try {
                    Remove-Item -Path $_.FullName -ErrorAction Stop
                    Add-Content -Path $logFile -Value "Removed duplicate file: $($_.FullName)"
                    Write-Host "Removed duplicate file: $($_.FullName)"
                } catch {
                    Add-Content -Path $logFile -Value "Failed to remove duplicate file: $($_.FullName) - Error: $_"
                    Write-Host "Failed to remove duplicate file: $($_.FullName)"
                }
            } else {
                $fileHashes[$fileHash] = $_.FullName
            }
        }
    }
}

# Function to check for unsupported file types and remove them
function Check-UnsupportedFileTypes {
    param (
        [string]$sourcePath,
        [string]$logFile,
        [string[]]$unsupportedExtensions = @(".exe", ".dll", ".bat", ".ini", "~$") # Incompatible SharePoint Online extensions
    )
    Write-Host "Checking for unsupported file types..." -ForegroundColor Green
    Start-sleep -seconds 2
    Get-ChildItem -Path $sourcePath -Recurse -File | ForEach-Object {
        if ($unsupportedExtensions -contains $_.Extension -or $_.Name -match '^\~\$') {
            try {
                Remove-Item -Path $_.FullName -ErrorAction Stop
                Add-Content -Path $logFile -Value "Removed unsupported file type: $($_.FullName)"
                Write-Host "Removed unsupported file type: $($_.FullName)"
            } catch {
                Add-Content -Path $logFile -Value "Failed to remove unsupported file type: $($_.FullName) - Error: $_"
                Write-Host "Failed to remove unsupported file type: $($_.FullName)"
            }
        }
    }
}

# Function to remove QuickBooks Desktop files
function Remove-QuickBooksFiles {
    param (
        [string]$sourcePath,
        [string]$logFile,
        [string[]]$quickBooksExtensions = @(".qbw", ".qbb", ".qba", ".qbx", ".qby")
    )
    Write-Host "Removing QuickBooks Desktop files..." -ForegroundColor Green
    Start-sleep -seconds 2
    Get-ChildItem -Path $sourcePath -Recurse -File | ForEach-Object {
        if ($quickBooksExtensions -contains $_.Extension) {
            try {
                Remove-Item -Path $_.FullName -ErrorAction Stop
                Add-Content -Path $logFile -Value "Removed QuickBooks file: $($_.FullName)"
                Write-Host "Removed QuickBooks file: $($_.FullName)"
            } catch {
                Add-Content -Path $logFile -Value "Failed to remove QuickBooks file: $($_.FullName) - Error: $_"
                Write-Host "Failed to remove QuickBooks file: $($_.FullName)"
            }
        }
    }
}

# Function to check for path length limits and shorten paths
function Check-PathLengthLimits {
    param (
        [string]$sourcePath,
        [string]$logFile,
        [int]$maxPathLength = 260
    )
    Write-Host "Checking for path length limits..." -ForegroundColor Green
    Start-sleep -seconds 2
    Get-ChildItem -Path $sourcePath -Recurse | ForEach-Object {
        if ($_.FullName.Length -gt $maxPathLength) {
            $shortenedPath = $_.FullName.Substring(0, $maxPathLength - 10) + "~" + $_.FullName.Substring($_.FullName.Length - 10)
            try {
                Rename-Item -Path $_.FullName -NewName $shortenedPath -ErrorAction Stop
                Add-Content -Path $logFile -Value "Shortened path: $($_.FullName) to $shortenedPath"
                Write-Host "Shortened path: $($_.FullName) to $shortenedPath"
            } catch {
                Add-Content -Path $logFile -Value "Failed to shorten path: $($_.FullName) - Error: $_"
                Write-Host "Failed to shorten path: $($_.FullName)"
            }
        }
    }
}

# -----------------------MAIN SCRIPT EXECUTION-----------------------

# Define variables
# sourcePath = "Source NTFS address of data.
# $logFile = "Destination NTFS address to store script logs."
$sourcePath = "C:\ALPHA\SharePoint Test"
$logFile = "C:\ALPHA\Logs\SharePoint_testlog.txt"

Write-Host "Starting SharePoint Data Resolver, please wait..." -ForegroundColor Green
Start-sleep -seconds 2
# Create or clear the log file
if (Test-Path $logFile) {
    Clear-Content -Path $logFile
    Write-Host "Cleared existing log file: $logFile"
} else {
    New-Item -Path $logFile -ItemType File
    Write-Host "Created new log file: $logFile"
}

# Remove hidden attribute from all files

Write-Host "Checking for hidden attribute on data and resolving..." -ForegroundColor Green
Start-sleep -seconds 2
Get-ChildItem -Path $sourcePath -Recurse -Force | ForEach-Object {
    if ($_.Attributes -match "Hidden") {
        try {
            $newAttributes = $_.Attributes -band (-bnot [System.IO.FileAttributes]::Hidden)
            $_.Attributes = $newAttributes
            Write-Host "Removed hidden attribute from: $($_.FullName)"
        } catch {
            Add-Content -Path $logFile -Value "Failed to remove hidden attribute from: $($_.FullName) - Error: $_"
            Write-Host "Failed to remove hidden attribute from: $($_.FullName)"
        }
    }
}

# Resolve invalid item names
Resolve-InvalidItemNames -sourcePath $sourcePath -logFile $logFile

# Remove empty items
Remove-EmptyItems -sourcePath $sourcePath -logFile $logFile

# Remove duplicate archives
Remove-DuplicateArchives -sourcePath $sourcePath -logFile $logFile

# Remove exact duplicate files
Remove-DuplicateFiles -sourcePath $sourcePath -logFile $logFile

# Check for unsupported file types and remove them
Check-UnsupportedFileTypes -sourcePath $sourcePath -logFile $logFile

# Remove QuickBooks Desktop files
Remove-QuickBooksFiles -sourcePath $sourcePath -logFile $logFile

# Check and shorten paths that exceed length limits
Check-PathLengthLimits -sourcePath $sourcePath -logFile $logFile

Start-sleep -seconds 2
Write-Host "SharePoint Data Resolver completed successfully - refer to log file for report." -ForegroundColor Yellow
