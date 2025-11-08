#######################################################
# This program deletes duplicate files in your system #
# Author: Sandra E.                                   #
# Creation date: 08.11.2025                           #
# Version: 1.0                                        #
#######################################################

Write-Host "========================================================================" -ForegroundColor Magenta
Write-Host " Searching for duplicates started....." -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Magenta

$sourceDir = Read-Host "Enter root directory to search from (e.g. C:\)"

$file_duplicates = Get-ChildItem -Path $sourceDir -Recurse |
    Group-Object -Property Length |
    Where-Object { $_.Count -gt 1 } |
    Select-Object -ExpandProperty Group |
    Get-FileHash |
    Group-Object -Property Hash |
    Where-Object { $_.Count -gt 1 } |
    ForEach-Object {
        $_.Group | Select-Object Path, Hash
    }

Write-Host "Search completed..." -ForegroundColor Green
foreach ($file in $file_duplicates) {
    Write-Host "Found duplicate: $($file.Path)" -ForegroundColor Magenta
}
