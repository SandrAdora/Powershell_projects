
# Pfade für den Downloads-Ordner und Zielordner
$sourcePath = "C:\Users\$env:USERNAME\Downloads"

# Diestination folder 
$destinationFolders = @{
    "Documents"   = "C:\Users\$env:USERNAME\Documents"
    "Pictures"    = "C:\Users\$env:USERNAME\Pictures"
    "Music"       = "C:\Users\$env:USERNAME\Music"
    "Videos"      = "C:\Users\$env:USERNAME\Videos"
    "PDFs"        = "C:\Users\$env:USERNAME\Documents\PDFs"
    "Zips"        = "C:\Users\$env:USERNAME\Dokuments\Zips"
    "Others"      = "C:\Users\$env:USERNAME\others"

	
	
    
}

# Überprüfen, ob der Quellordner existiert
if (-not (Test-Path $sourcePath)) {
    Write-Host "Der Quellordner existiert nicht: $sourcePath" -ForegroundColor Red
    exit
}
# Dateien die sich im Downloadordner befinden anzeigen
Write-Host "Folgende Dateien befinden sich in deinem Downloadsordner:" -ForegroundColor Magenta
Get-ChildItem -Path $sourcePath -File | ForEach-Object{
	Write-Host "$_.Name" -ForegroundColor Cyan
}

# Sortieren der Dateien im Downloads-Ordner
# Costume location for specific names
Write-Host "Sortiervorgang wurde gestartet..." -ForegroundColor Magenta
Get-ChildItem -Path $sourcePath -File | ForEach-Object {
    $fileExtension = $_.Extension.ToLower()
    $destinationPath = $null
    # Anderer Speicherort für bestimmte datei
   if ($fileExtension -eq ".pdf" -and $_.Name -like "ehg*" -or $_.Name -like "DIC*") {
        $destinationPath = $destinationFolders["Others"]
	}
   else{	

    # Zielordner basierend auf Dateitypen festlegen
    switch ($fileExtension) {
        ".jpg" { $destinationPath = $destinationFolders["Pictures"] }
        ".jpeg" {$destinationPath = $destinationFolders["Pictures"]}
        ".png" { $destinationPath = $destinationFolders["Pictures"] }
        ".mp3" { $destinationPath = $destinationFolders["Music"] }
        ".mp4" { $destinationPath = $destinationFolders["Videos"] }
        ".pdf" { $destinationPath = $destinationFolders["PDFs"] }
        ".docx" { $destinationPath = $destinationFolders["Documents"] }
        ".xlsx" { $destinationPath = $destinationFolders["Excel"] }
        ".csv" {$destinationPath = $destinationFolders["Excel"]}
        ".zip" {$destinationPath = $destinationFolders["Zips"]}

		
		

    }
}

    # Datei verschieben, falls ein Zielordner definiert ist
    if ($destinationPath) {
        if (-not (Test-Path $destinationPath)) {
            New-Item -ItemType Directory -Path $destinationPath | Out-Null
        }
        Move-Item -Path $_.FullName -Destination $destinationPath -Force
        Write-Host "Datei $($_.Name) wurde nach $destinationPath verschoben." -ForegroundColor Cyan
    } else {
        Write-Host "Keine Regel fuer Dateityp $fileExtension gefunden. Datei $($_.Name) wird uebersprungen." -ForegroundColor Yellow
    }
}

Write-Host "Sortiervorgang abgeschlossen!" -ForegroundColor Green
