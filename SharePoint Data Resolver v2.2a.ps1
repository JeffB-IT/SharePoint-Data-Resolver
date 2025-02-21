# SharePoint Data Resolver v2.2a
# Author: Jeffrey Bunnaman, IT Specialist - Alpha Corporate Technologies
# 
# CHANGE LOG:
# 20240906 - Added file hash verification for removing duplicate files
# 20240907 - Additional logic added for invalid item names for SharePoint
# 20250207 - Added additional logic to improve user experience.
# 20250221 - Made QOL changes to the remove duplicate files logic.

# This PowerShell script will prepare data for migration from on-premise to Microsoft 365 SharePoint Online.
# Please note: Utilize SharePoint Migration scan function to scan source data once script execution is complete; if errors persist, consult scan log CSV downloaded from SharePoint Admin Center and take corrective action if necessary.


# -----------------------FUNCTIONS-----------------------

# Function to remove duplicate files
function Remove-DuplicateFiles {
    param (
        [string]$sourcePath,
        [string]$logFile
    )
    Write-Host "Removing duplicate files, checking path lengths, and verifying hash..." -ForegroundColor Green
    Start-sleep -seconds 2
    $fileHashes = @{ }
    
    # Ensure the source path is valid
    if (-not (Test-Path -Path $sourcePath)) {
        Write-Host "The source path $sourcePath is invalid."
        return
    }

    # Iterate through files in the source directory
    $files = Get-ChildItem -Path $sourcePath -Recurse -File
    foreach ($file in $files) {
        try {
            # Check if the file path length exceeds the maximum allowable path length
            if ($file.FullName.Length -gt 260) {
                Write-Host "Skipping file: $($file.FullName) (Path is too long)"
                Add-Content -Path $logFile -Value "Skipped file: $($file.FullName) (Path is too long)"
                continue
            }

              # Compute the file hash using the updated custom hash function
              $hashValue = Get-FileHashCustom($file.FullName)
            
              if ($hashValue) {
                  if ($fileHashes.ContainsKey($hashValue)) {
                      try {
                          # Remove the duplicate file
                          Remove-Item -Path $file.FullName -ErrorAction Stop
                          Add-Content -Path $logFile -Value "Removed duplicate file: $($file.FullName)"
                          Write-Host "Removed duplicate file: $($file.FullName)"
                      } catch {
                          Add-Content -Path $logFile -Value "Failed to remove duplicate file: $($file.FullName) - Error: $_"
                          Write-Host "Failed to remove duplicate file: $($file.FullName)"
                      }
                  } else {
                      # Add the hash to the list of already seen hashes
                      $fileHashes[$hashValue] = $file.FullName
                  }
              } else {
                  Write-Host "Skipping file: $($file.FullName) (unable to compute hash)"
                  Add-Content -Path $logFile -Value "Skipped file: $($file.FullName) (unable to compute hash)"
              }
          } catch {
              Write-Host "Error processing file: $($file.FullName) - Error: $_"
              Add-Content -Path $logFile -Value "Error processing file: $($file.FullName) - Error: $_"
          }

            # Replace special characters with underscores
            $newName = $file.Name -replace '[^a-zA-Z0-9\s\.\-]', '_'
            
            # If the file name changed, rename it
            if ($file.Name -ne $newName) {
                try {
                    Rename-Item -Path $file.FullName -NewName $newName -ErrorAction Stop
                    Write-Host "Renamed file: $($file.FullName) to $newName"
                    Add-Content -Path $logFile -Value "Renamed file: $($file.FullName) to $newName"
                } catch {
                    Add-Content -Path $logFile -Value "Failed to rename file: $($file.FullName) - Error: $_"
                    Write-Host "Failed to rename file: $($file.FullName)"
                    continue
                }
            }

          
    }
}


# Custom function to calculate file hash using .NET HashAlgorithm
function Get-FileHashCustom {
    param (
        [string]$filePath
    )

    try {
        # If the file path exceeds 260 characters, use the UNC path format \\?\
        if ($filePath.Length -gt 260) {
            $filePath = "\\?\$filePath"
        }

        # Open file stream with the UNC path if required
        $fileStream = [System.IO.File]::OpenRead($filePath)
        
        # Create SHA256 instance
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        
        # Compute the hash
        $hashBytes = $sha256.ComputeHash($fileStream)
        
        # Convert hash bytes to hexadecimal string
        $hashString = [BitConverter]::ToString($hashBytes) -replace '-'
        
        # Close the file stream
        $fileStream.Close()
        
        return $hashString
    } 
    
    catch {
        Write-Host "Error computing hash for file: $filePath"
        return $null
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
    Write-Host "Peforming additional checks for path length limits..." -ForegroundColor Green
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


# Remove exact duplicate files
Remove-DuplicateFiles -sourcePath $sourcePath -logFile $logFile

# Check and shorten paths that exceed length limits
Check-PathLengthLimits -sourcePath $sourcePath -logFile $logFile

# Remove empty items
Remove-EmptyItems -sourcePath $sourcePath -logFile $logFile

# Check for unsupported file types and remove them
Check-UnsupportedFileTypes -sourcePath $sourcePath -logFile $logFile

# Remove QuickBooks Desktop files
Remove-QuickBooksFiles -sourcePath $sourcePath -logFile $logFile

Start-sleep -seconds 2
Write-Host "SharePoint Data Resolver completed successfully - refer to log file for report." -ForegroundColor Yellow
