#######################################################
# This program deletes duplicate files in your system #
# Author: Sandra E.                                   #
# Creation date: 08.11.2025                           #
# Version: 1.1                                        #
#######################################################

param (
    [switch]$DryRun,
    [string]$LogFile = "$env:TEMP\duplicate_deletion_log.txt"
)

Write-Host "========================================================================" -ForegroundColor Magenta
Write-Host " Searching for duplicates started....." -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Magenta

$sourceDir = Read-Host "Enter root directory to search from (e.g. C:\)"

$file_duplicates = Get-ChildItem -Path $sourceDir -Recurse -File |
    Group-Object -Property Length |
    Where-Object { $_.Count -gt 1 } |
    Select-Object -ExpandProperty Group |
    Get-FileHash |
    Group-Object -Property Hash |
    Where-Object { $_.Count -gt 1 } |
    ForEach-Object {
        $_.Group | Select-Object Path, Hash
    }

$file_duplicates_grouped = $file_duplicates | Group-Object -Property Hash

Write-Host "Search completed..." -ForegroundColor Green
Write-Host "========================================================================" -ForegroundColor Magenta

foreach ($group in $file_duplicates_grouped) {
    $files = $group.Group
    $filesToDelete = $files | Select-Object -Skip 1

    foreach ($file in $filesToDelete) {
        $message = "Duplicate found: $($file.Path)"
        Write-Host $message -ForegroundColor Yellow
        Add-Content -Path $LogFile -Value $message

        if ($DryRun) {
            Write-Host "DryRun mode: Skipping deletion of $($file.Path)" -ForegroundColor Cyan
        } else {
            try {
                Remove-Item -Path $file.Path -Force
                Write-Host "Deleted: $($file.Path)" -ForegroundColor Red
                Add-Content -Path $LogFile -Value "Deleted: $($file.Path)"
            } catch {
                Write-Host "Error deleting $($file.Path): $_" -ForegroundColor DarkRed
                Add-Content -Path $LogFile -Value "Error deleting $($file.Path): $_"
            }
        }
    }
}

Write-Host "========================================================================" -ForegroundColor Magenta
Write-Host " Process completed. Log saved to: $LogFile" -ForegroundColor Green
Write-Host "========================================================================" -ForegroundColor Magenta
