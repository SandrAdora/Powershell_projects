#######################################################
# This program deletes duplicate files in your system #
# Author: Sandra E.                                   #
# Creation date: 08.11.2025                           #
# Version: 1.5 (subfolder-only search)                #
#######################################################

# Get current user info
try {
    $CurrentUser = $env:USERNAME
    $CurrentProfile = $env:USERPROFILE
    Write-Host "User: $CurrentUser" -ForegroundColor DarkCyan
    Write-Host "Profile: $CurrentProfile" -ForegroundColor DarkCyan
} catch {
    Write-Host "Could not read user credentials." -ForegroundColor Red
    Start-Sleep -Seconds 5
    exit
}

Write-Host "========================================================================" -ForegroundColor Magenta
Write-Host " Searching for duplicates in subfolders only..." -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Magenta

# Prompt for root directory
$rootDirectory = Read-Host "Enter root directory to search from (e.g. C:\Users\)"

# Validate root path
if (!(Test-Path $rootDirectory)) {
    Write-Host "Directory does not exist: $rootDirectory" -ForegroundColor Red
    exit
}

# Setup logging
$LogFile = "$env:TEMP\duplicate_deletion_log.txt"
if (!(Test-Path $LogFile)) {
    New-Item -Path $LogFile -ItemType File -Force | Out-Null
    Write-Host "Log file created at $LogFile" -ForegroundColor Green
}

# Define file types to check
$Extensions = @(".pdf", ".docx", ".jpg", ".png", ".txt", ".pptx", ".html", ".drawio")
$dublicateCounter = 0

# Get subfolders only
$subFolders = Get-ChildItem -Path $rootDirectory -Directory

foreach ($folder in $subFolders) {
    foreach ($extension in $Extensions) {
        $files = Get-ChildItem -Path $folder.FullName -Recurse -File | Where-Object { $_.Extension -eq $extension }

        foreach ($file in $files) {
            $fileName = $file.Name
            $sourcePath = $file.FullName

            # Check for duplicates in other subfolders
            $duplicateFound = $false
            foreach ($compareFolder in $subFolders) {
                if ($compareFolder.FullName -ne $folder.FullName) {
                    $comparePath = Join-Path -Path $compareFolder.FullName -ChildPath $fileName
                    if (Test-Path $comparePath) {
                        try {
                            Remove-Item -Path $sourcePath -Force
                            $dublicateCounter++
                            Write-Host "Deleted duplicate: $fileName from $($folder.Name)" -ForegroundColor Green
                            Add-Content -Path $LogFile -Value "Deleted: $sourcePath"
                            $duplicateFound = $true
                            break
                        } catch {
                            Write-Host "Error deleting $fileName" -ForegroundColor Red
                            Add-Content -Path $LogFile -Value "Error deleting $sourcePath"
                        }
                    }
                }
            }

            if (-not $duplicateFound) {
                Write-Host "No duplicate found for: $fileName in other subfolders" -ForegroundColor Yellow
            }
        }
    }
}

Write-Host "========================================================================" -ForegroundColor Magenta
Write-Host " Process completed. $dublicateCounter duplicates deleted." -ForegroundColor Cyan
Write-Host " Log saved to: $LogFile" -ForegroundColor Green
Write-Host "========================================================================" -ForegroundColor Magenta

Start-Sleep -Seconds 4
