
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
    "Bachelor"    = "D:\BachelorThesisExperiment\Bilder"
    "temp"        = "C:\Users\$env:USERNAME\Temp"
    "kaggle"      = "D:\KaggleProjects"
    "sonstiges"   = "D:\Sonstiges"
    "Excel"       = "D:\Excel"
    "eHealth"     = "D:\Studium\E_Health\Vorlesungen"
    "datascience_v" = "D:\Studium\Applied Datascience with python\Vorlesungen"
    "datascience_u" = "D:\Studium\Applied Datascience with python\Uebungen"
    "datascience_p" = "D:\Studium\Applied Datascience with python\Project\Documents"
    "batFiles"     = "D:\WindowsBatDateien"
    "commerzbank" = "D:\Commerzbank"
    "bio_k"	  = "D:\Studium\Biometrie\Kausuren"
	"obs" = "D:\OBS"
	"jsonFiles" = "D:\JSONFILES"
	"powerpoints" = "D:\Powerpoints"
	"database" = "D:\Database"
	
	
    
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
        $destinationPath = $destinationFolders["eHealth"]
	}elseif($fileExtension -eq ".pdf" -and  $_.Name -like "Applied*" ){
		$destinationPath = $destinationFolders["datascience_v"]
	}elseif($fileExtension -eq ".pdf" -and $_.Name -like "*parkonto*" -or $_.Name -like "*edingungen_*parlkonto*"){
        $destinationPath = $destinationFolders["commerzbank"]
    }
    elseif($fileExtension -eq ".zip" -and $_.Name -like "Lecture*" ){
		$destinationPath = $destinationFolders["datascience_u"]
	}elseif($fileExtension -eq ".zip" -and $_.Name -like "Module*"){
		$destinationPath = $destinationFolders["datascience_v"]
	}elseif($fileExtension -eq ".pdf" -and $_.Name -like "BIO*"){
		$destinationPath = $destinationFolders["bio_k"]
}elseif($fileExtension -eq ".pdf" -and $_.Name -like "Scientific*" -or $_.Name -like "Data*"){
	$destinationPath = $destinationFolders["datascience_p"]
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
        ".xlsx" { $destinationPath = $destinationFolders["Documents"] }
        ".csv" {$destinationPath = $destinationFolders["Excel"]}
        ".zip" {$destinationPath = $destinationFolders["Zips"]}
        ".svg" {$destinationPath = $destinationFolders["Bachelor"]}
        ".exe"  {$destinationPath = $destinationFolders["temp"]}
	".Browser.for.SQLite-v3.13.1-win64.msi" {$destinationPath = $destinationFolders["temp"]}
        ".ipynb" {$destinationPath =  $destinationFolders["Kaggle"]}
        ".npy" {$destinationPath = $destinationFolders["sonstiges"]}
        ".html" {$destinationPath = $destinationFolders["sonstiges"]}
        ".ris" {$destinationPath = $destinationFolders["sonstiges"]}
		".msi" {$destinationPath = $destinationFolders["sonstiges"]}
        ".bat" {$destinationPath = $destinationFolders["batFiles"]}
		".json" {$destinationPath = $destinationFolders["jsonFiles"]}
		".pptx" {$destinationPath = $destinationFolders["powerpoints"]}
		".mkv" {$destinationPath  = $destinationFolders["obs"]}
		".db"  {$destinationPath = $destinationFolders["database"]}
		
		

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
