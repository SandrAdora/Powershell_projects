###############################################
# Dieses Programm erstellt in einer Excel     #
# den Stundenplan jedes Mitarbeiters          #
# und fügt neue Daten hinzu.                  #
#                                             #
# Version: 1.4                                #
# Autor:   Sandra Edigin                      #
# Datum:   23.03.2025                         #
###############################################

Write-Host "##################################################################################################" -ForegroundColor DarkYellow
Write-Host "# Willkommen beim Arbeitszeiterfassungs-Skript des Bavaria Gym Neutraubling.                     #" -ForegroundColor DarkYellow
Write-Host "# Dieses Tool hilft dir dabei, Deine Arbeitszeiten schnell und einfach zu erfassen.              #" -ForegroundColor DarkYellow
Write-Host "#                                                                                                #" -ForegroundColor DarkYellow
Write-Host "# Version: 1.3                                                                                   #" -ForegroundColor DarkYellow
Write-Host "# Autor: Sandra Edigin                                                                           #" -ForegroundColor DarkYellow
Write-Host "# Datum: 23.03.2025                                                                              #" -ForegroundColor DarkYellow
Write-Host "##################################################################################################" -ForegroundColor DarkYellow

# Heutige Datum 
$CurrentDate = Get-Date
Write-Host "$CurrentDate" -ForegroundColor Blue

Write-Host "`n!!! WICHTIGER HINWEIS !!!`n" -ForegroundColor Yellow
Write-Host "`n-----------------------------------------------------------------------------------------------------------------------`n" -ForegroundColor Magenta
Write-Host "`nVielen Dank, dass du dieses Tool verwendest!`n" -ForegroundColor Cyan
Write-Host "`nBitte speichere deine Excel-Datei direkt nach der ersten Erstellung ab [strg + s]." -ForegroundColor Yellow
Write-Host "Wenn du Aenderung in deinem Arbeitsordner oder an der Datei vornimmst, denke immer daran, die Datei zu speichern und [ja] zu klicken, um Aenderungen zu uebernehmen.`n" -ForegroundColor Yellow
Write-Host "Du wirst Schritt fuer Schritt durch die Eingabe deiner Arbeitszeiten gefuehrt." -ForegroundColor Cyan
Write-Host "`nLass uns starten!`n" -ForegroundColor Green
Write-Host "`n------------------------------------------------------------------------------------------------------------------------`n" -ForegroundColor Magenta

# Monatsliste
$month = (Get-Date).ToString("MMMM")
# Name des Arbeitsblatts 
$sheetName = "$(Get-Date -Format 'MMMM yyyy')"
# Pfad zum Logo Bavarian Gym
$logopath = "C:\Users\$env:USERNAME\Pictures\bav_gym.jpg"


# Name des Arbeiters ermitteln
$workername = Read-Host "Bitte gib deinen Nachnamen ein"

# überprüfen ob Logo vorhanden ist 
if(-not (Test-Path $logopath)){
    Write-Host "Bild konnte nicht gefunden werden" -ForegroundColor Yellow
    exit
}

# Funktionen 
function ConvertTo-Date {
    param (
        [string]$inputDate
    )
    try {
        return Get-Date -Date $inputDate
    } catch {
        Write-Host "Ungültiges Datumformat. Bitte erneut eingeben." -ForegroundColor Yellow
        return $null
    }
}

function Open-ExcelWorkbook {
    param (
        [string]$filePath,
        [ref]$excelApp
    )
    if (Test-Path $filePath) {
        $excelApp.Value = New-Object -ComObject Excel.Application
        $excelApp.Value.Visible = $false
        return $excelApp.Value.Workbooks.Open($filePath)
    } else {
        Write-Host "Die Datei wurde nicht gefunden." -ForegroundColor Red
        return $null
    }
}

function Update-ExcelCells {
    param (
        $worksheet,
        [int]$column,
        [datetime]$dateToFind,
        $updateAction
    )
    $lastRow = $worksheet.Cells($worksheet.Rows.Count, $column).End(-4162).Row
    for ($row = 4; $row -le $lastRow; $row++) {
        $cell = $worksheet.Cells($row, $column)
        $cellValue = $cell.Value2

        if ($null -ne $cellValue -and [DateTime]::FromOADate($cellValue) -eq $dateToFind) {
            $updateAction.Invoke($worksheet, $row)
        }
    }
}
function Update-ExcelFile-For-Writing {
    param (
        [string]$filePath
    )
    Set-ItemProperty -Path $filePath -Name IsReadOnly -Value $false
    Write-Host "Datei erfolgreich vorbereitet." -ForegroundColor Cyan
}


# Ende der Funktionen

do{
    try{
    
    $typeOfJob = Read-Host "Gib deine Beschaeftigungsart ein (Minijob, Teilzeit, Vollzeit)"
    if($typeOfJob -notin  @("minijob", "teilzeit", "vollzeit")){
        throw "Eingabe ist ungueltig."
    }
    $richtigeEingabe = $true
    }catch{
        Write-Host "Fehlermeldung: Bitte Minijob, Teilzeit oder Vollzeit angeben." -ForegroundColor Red
        $richtigeEingabe = $false
    }
}while(-not $richtigeEingabe)

$nameKleinbuchstaben = $workername.ToLower()
$jobKeinbuchstaben = $typeOfJob.ToLower()
$ordnerPath = "D:\Bavaria_Gym_Arbeitsordner\$jobKeinbuchstaben\$nameKleinbuchstaben"
# Speicherort festlegen
$excelPath = "$ordnerPath\Arbeitszeiterfassung.xlsx"

# Nutzer soll seine Arbeitszeiten ansehen ohne Zeiten eintragen zu müssen
Write-Host "`nHinweis: Bitte nur [2], [3] und [4] auswaehlen wenn fur dich bereits ein Arbeitsordner erstellt wurde!!!!`n" -ForegroundColor Yellow

    $vorhaben = Read-Host "Folgende Moeglichkeiten: 
    1. Arbeitszeiten eintragen [ENTER] (Vor allem dann auswaehlen, wenn du noch keinen Ordner hast)
    2. Arbeitszeiten ansehen [2]
    3. Arbeitsstunden fuer diesen Monat auszurechnen [3]
    4. Arbeitsstunden bearbeiten [4]
    5. Skript beenden [exit] eingeben
    Bitte moeglichkeit waehlen"
    switch ($vorhaben) {
        "2" {  
            if (-not (Test-Path $excelPath)) {
                Write-Host "Die Datei existiert in diesem Pfad nicht: $excelPath" -ForegroundColor Red
                exit
            }
            Start-Process $excelPath
                exit  
            }
        "4"{
            Write-Host "Prozess Arbeitsstunden bearbeiten wurde gestartet" -ForegroundColor Cyan

            # Excel-Anwendung öffnen
            $Excel = $null
            if (Test-Path $excelPath) {
                Write-Host "Arbeitserfassung wird geöffnet" -ForegroundColor Cyan
                $workbook = Open-ExcelWorkbook -filePath $excelPath -excelApp ([ref]$Excel)
                if ($workbook) {
                    Write-Host "Hinweis: Nur das aktuelle Monat kann mithilfe des Skripts bearbeitet werden" -ForegroundColor Magenta
                    $arbeitsblatt = $workbook.Sheets.Item($sheetName)
            
                    # Eingabe des zu korrigierenden Datums
                    $zu_korrigieren = Read-Host "Bitte Datum eingeben, das zu korrigieren gilt [z. B. 25.04.2025]"
                    $zu_korrigieren_date = Get-Date -Format "$zu_korrigieren"
            
            
                    # Letzte Zeile in der Spalte bestimmen
                    $last_row = $arbeitsblatt.Cells($arbeitsblatt.Rows.Count, 4).End(-4162).Row
            
                    # Zellen durchlaufen und Datum korrigieren
                    $korrekturDurchgefuehrt = $false
                    for ($row = 4; $row -le $last_row; $row++) {
                        $datum = $arbeitsblatt.Cells.Item($row, 4).Value2
                    
                        Write-Host "Eintrag fuer: $($datum) gefunden" -ForegroundColor Cyan
                        # Vergleich mit dem zu korrigierenden Datum
                        if ($null -ne $datum -and $datum -eq $zu_korrigieren_date) {
                            Write-Host "Datum gefunden in Zeile $row. Bitte neue Arbeitszeit eingeben." -ForegroundColor Cyan
                            $startTime = Read-Host "Bitte neuer Arbeitszeitsbeginn eingeben (HH:mm)"
                            $startDateTime = [datetime]::ParseExact($startTime, "HH:mm", $null)
            
                            $endTime = Read-Host "Bitte neues Arbeitsschluss eingeben (HH:mm)"
                            $endDateTime = if ($endTime -eq "24:00") { 
                                [datetime]::ParseExact("00:00", "HH:mm", $null).AddDays(1) 
                            } else { 
                                [datetime]::ParseExact($endTime, "HH:mm", $null) 
                            }
            
                            # Berechnung der Arbeitsstunden
                            $arbeitsstunden = ($endDateTime - $startDateTime).TotalHours
            
                            # Arbeitszeit in entsprechende Spalten eintragen
                            $arbeitsblatt.Cells.Item($row, 5).Value = "$startTime"
                            $arbeitsblatt.Cells.Item($row, 6).Value = "$endTime"
                            $arbeitsblatt.Cells.Item($row, 7).Value = "$arbeitsstunden"
                            Write-Host "Arbeitszeit für Zeile $row erfolgreich aktualisiert." -ForegroundColor Green
            
                            $korrekturDurchgefuehrt = $true
                            break # Schleife abbrechen, da nur eine Stelle bearbeitet werden soll
                        }
                    }
            
                    # Feedback für den Nutzer
                    if (-not $korrekturDurchgefuehrt) {
                        Write-Host "Das eingegebene Datum wurde nicht gefunden." -ForegroundColor Yellow
                    }
            
                    # Änderungen speichern und Datei schließen
                    $workbook.Save()
                    Write-Host "Skript wird jetzt beendet. Auf Wiedersehen." -ForegroundColor Green
                    Start-Process $excelPath
                    } else {
                    Write-Host "Die Arbeitsmappe konnte nicht geöffnet werden." -ForegroundColor Red
                    }
                                    # Excel schließen und freigeben
    
                Start-Process $excelPath
                exit
            }           

            elseif($bearbeiten -eq "2"){
                Write-Host "Prozess Stunden zu korigieren wurde gestartet" -ForegroundColor Cyan
            }

            else{
                    Write-Host "Workbook konnte nicht geoffnet werden. Entweder Sie ist beschaedigt oder sie existiert nicht." -ForegroundColor Red
                    Write-Host "Das Skript wird beendet. Aufwieder sehen" -ForegroundColor Yellow
                    exit
            }
        }
        "3" {
            # Excel-Anwendung öffnen
            Get-Process -Name Excel -ErrorAction SilentlyContinue | Stop-Process -Force
            # Excel-Objekt erstellen
            $Excel = New-Object -ComObject Excel.Application
            $Excel.Visible = $true
    
            # Excel-Datei öffnen
            $workBook = $Excel.Workbooks.Open($excelPath)
    
            # Maximale Stunden für jeden Mitarbeiter abhängig von seiner Arbeitsart
            $JobMaxHrs = 0
            if($workBook){
                Write-Host "`nProcess der Arbeitsstunden Berechnung wurde gestartet...`n" -ForegroundColor Cyan
                $workbookSheet = $workBook.Sheets.Item("$sheetName")
            }else{
                Write-Host "Datei konnte nicht geoeffnet werden" -ForegroundColor Red 
                exit
            }
            $JobMaxHrs = $workbookSheet.Range("B2").Value2
            Write-Host "Die Anzahl an maximale Stunden fuer $jobKeinbuchstaben Mitarbeiter betraegt: $JobMaxHrs" -ForegroundColor Blue
            $values = 0
            # Letzte ausgefüllte Zeile in der Spalte 7 ermitteln
            $last_row = $workbookSheet.Cells($workbookSheet.Rows.Count, 7).End(-4162).Row # -4162 entspricht xlUp
    
            # Werte aus Spalte 7 summieren
            $values = 0
            $result = 0
            for ($row = 4; $row -le $last_row; $row++) {
                $cellValues = $workbookSheet.Cells.Item($row, 7).Value2
                if ($null -ne $cellValues) { # Überprüfung, dass die Zelle nicht leer ist
                    $values += $cellValues
                }
            }
            # Ergebnis ausgeben
            Write-Host "`nDie Berechnung ist abgeschlossen`n" -ForegroundColor Green
            if($values -gt $JobMaxHrs){
                $result = $values - $JobMaxHrs
                $w = ""
                if($result -eq "1"){
                    $w = "Stunde"
                }else{
                    $w = "Stunden"
                }
                Write-Host "Du hast $($result) Plus $($w)" -ForegroundColor Blue
            }elseif($values -eq $JobMaxHrs){
                $result = 0
                Write-Host "Du hast die Maximale Stunden fuer diesen Monat erreicht $($values)" -ForegroundColor Yellow
            }else{
                $result = $JobMaxHrs -$values
                Write-Host "Noch $($result) um die Stunden fuer diesen Monat zu erreichen." -ForegroundColor Cyan
            }
            $word = ""
            if($result -eq "1"){
                $word = "Stunde"
            } elseif($result -eq "0" -or $result -gt "1"){
                $word = "Stunden"
            }
            Write-Host "In diesem Monat hat du insgesamt $($values) $($word) von $($JobMaxHrs) gearbeitet." -ForegroundColor Cyan
            $speichern = Read-Host "Soll die geleisteten Stunden in deinem Arbeitsordner uebernommen werden? [ja] oder [nein]"
            if($speichern -eq "ja" -or $speichern -eq "Ja"){
                # Speicher den Wert rechts neben ber Maxmalen Anzahl an Stunden pro monat
                Write-Host "`nDie Anzahl gearbeitet Stunden werden in deinem Arbeitsblatt eingefuegt`n" -ForegroundColor Cyan
                $workbookSheet.Cells.Item(3,2) = $values
                Write-Host "Die Aenderungen wurden eingetragen" -ForegroundColor Green
                Write-Host "Bitte Aenderung an deiner Datei mit [strg + s] abspeichern." -ForegroundColor Yellow
                Write-Host "Das Skript wird beendet. Bitte die [enter taste] druecken" -ForegroundColor Green 
    
                Write-Host "`nDas Bavaria Gym wuenscht dir einen erfolgreichen Tag!!!`n" -ForegroundColor DarkYellow
                
            }else{
                Write-Host "Die Berechnung wird nicht uebernommen" -ForegroundColor Cyan
                exit
            }
            # automatische Speicherung
            $workBook.SaveAs($excelPath)
            exit
        }
        "exit"{
            #Beendet Das Skript 
            Write-Host "Das Skript wurde auf Wunsch von $workername beendet" -ForegroundColor Yellow
            exit
        }
    
    }
# Beschäftigungsart
$maxMonatsStunden = 0

if ($typeOfJob -eq "Minijob" -or $typeOfJob -eq "minijob") {
    $maxMonatsStunden = 43
} elseif ($typeOfJob -eq "Teilzeit" -or $typeOfJob -eq "teilzeit") {
    $maxMonatsStunden = 80
} elseif ($typeOfJob -eq "Vollzeit" -or $typeOfJob -eq "vollzeit") {
    $maxMonatsStunden = 160
}

do {
    try {
        # Arbeitszeiten eingeben und prüfen
        $startTime = Read-Host "Bitte Arbeitsbeginn eingeben (HH:mm)"
        $startDateTime = [datetime]::ParseExact($startTime, "HH:mm", $null)
        $endTime = Read-Host "Bitte Arbeitsschluss eingeben (HH:mm)"
            
        # Sonderfall: 24:00 in 00:00 des nächsten Tages umwandeln
        if ($endTime -eq "24:00") {
            $endDateTime = [datetime]::ParseExact("00:00", "HH:mm", $null).AddDays(1)
        } else {
            $endDateTime = [datetime]::ParseExact($endTime, "HH:mm", $null)
        }
        if ($endDateTime -le $startDateTime) {
            throw "Die Endzeit muss spaeter als die Startzeit sein!"
        }
        $validInput = $true
    } catch {
        Write-Host "Fehler bei der Zeiteingabe: $_. Bitte im Format HH:mm erneut versuchen." -ForegroundColor Red
        $validInput = $false
    }
} while (-not $validInput)

# Arbeitsstunden berechnen
$workingHours = ($endDateTime - $startDateTime).TotalHours

Write-Host "----------------------- $month $year ---------------------" -ForegroundColor DarkCyan
Write-Host " Hallo $workername                                        "                                          
Write-Host " Die $workingHours Stunden werden dir eingetragen.        " -ForegroundColor Green
Write-Host "---------------------------------------------------------" -ForegroundColor DarkCyan

if (-not (Test-Path $ordnerPath)) {
    # Erstelle Ordner, falls nicht vorhanden
    New-Item -ItemType Directory -Path $ordnerPath | Out-Null
    Write-Host "Dein Arbeitsordner wurde erstellt......." -ForegroundColor DarkGreen
}
try {
    # Excel-Objekt erstellen
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $row = 4  # Daten fangen in Zeile 4 an, da Zeile 1 Überschriften enthält
    if (Test-Path $excelPath) {

        Write-Host "Datei existert und kann geoeffnet werden" -ForegroundColor Yellow
        # Existierende Datei öffnen
        $workBook = $Excel.Workbooks.Open($excelPath)

        # Arbeitsblattnamen prüfen und sicherstellen, dass es existiert
        $worksheet = $null
        foreach ($sheet in $workBook.Sheets) {
            if ($sheet.Name -eq $sheetName) {
                $worksheet = $sheet
                break
            }
        }
        # Neues Arbeitsblatt erstellen, falls es nicht existiert
        if (-not $worksheet) {
            Write-Host "Das Arbeitsblatt '$sheetName' existiert nicht." -ForegroundColor Yellow
            
            # Überschriften hinzufügen
            $worksheet = $workBook.Sheets.Add()
            $worksheet.Name = $sheetName
            $worksheet.Cells.Item(1,1).Value = "$month"
            $worksheet.Cells.Item(1,2).Value = "$workername"
            $worksheet.Cells.Item(1,3).Value = " "
            $worksheet.Cells.Item(1,4).Value = "Datum"
            $worksheet.Cells.Item(1,5).Value = "Beginn"
            $worksheet.Cells.Item(1,6).Value = "Ende"
            $worksheet.Cells.Item(1,7).Value = "Stunden"
            $worksheet.Cells.Item(1,8).Value = " "
            $worksheet.Cells.item(1,9).Value ="Urlaubstage Insgesamt"
            $worksheet.Cells.item(1,10).Value ="Urlaubstage genommen"

            Write-Host "Ein neues Arbeitsblatt fuer $month wurde erstellt." -ForegroundColor Green

        } else {
            Write-Host "Das Arbeitsblatt '$sheetName' existiert bereits." -ForegroundColor Green
        }   
        # Finde die erste freie Zeile
        $row = $worksheet.UsedRange.Rows.Count + 1
    }
    else {
        # Neue Arbeitsmappe erstellen
        $workBook = $Excel.Workbooks.Add()
        $worksheet = $workBook.Sheets.Item(1)
    
        $worksheet.Name = $sheetName

        # Überschriften hinzufügen
        $worksheet.Cells.Item(1,1).Value = "$month"
        $worksheet.Cells.Item(1,2).Value = "$workername"
        $worksheet.Cells.Item(1,3).Value = " "
        $worksheet.Cells.Item(1,4).Value = "Datum"
        $worksheet.Cells.Item(1,5).Value = "Beginn"
        $worksheet.Cells.Item(1,6).Value = "Ende"
        $worksheet.Cells.Item(1,7).Value = "Stunden"
        $worksheet.Cells.Item(1,8).Value = " "
        $worksheet.Cells.item(1,9).Value ="Urlaubstage Insgesamt"
        $worksheet.Cells.item(1,10).Value ="Urlaubstage genommen"
    }

    # Daten in die Tabelle einfügen
    $worksheet.Cells.Item(2,1).Value ="Max Stunden"
    $worksheet.Cells.Item(2,2).Value = "$maxMonatsStunden"
    $worksheet.Cells.Item(3, 1) = "Stunden gearbeitet"
    $worksheet.Cells.Item(3,2) = " "
    #$worksheet.Cells.Item($row,1).Value = "$typeOfJob"
    $worksheet.Cells.Item($row, 4).Value = Get-Date -Format "dd.MM.yyyy"
    $worksheet.Cells.Item($row, 5).Value = "$startTime"
    $worksheet.Cells.Item($row, 6).Value = "$endTime"
    $worksheet.Cells.Item($row, 7).Value = "$workingHours"

    # Farbliche Kennzeichnung der Überschriften
    $headerRange = $worksheet.Range("A1:J1")  # Bereich der Überschriften (Spalten A bis I, Zeile 1)
    $headerRange.Interior.Color = 49407 # Farbe Orange
    $headerRange.Font.Size = 12               # Schriftgröße auf 14 setzen

    # Weitere Headers anpassen
    $headerMinijob = $worksheet.Range("A2:A3")
    $headerMinijob.Interior.Color  = 49407 #Orange
    $headerMinijob.Font.Size = 11
    
    $bildExistiert = $false
    foreach($pic in $worksheet.Pictures()){
        if($pic.Filename -eq $logopath){
            $bildExistiert = $true 
            break
        }
    }
    if(-not($bildExistiert)){
        # Letzte verfügbare Zeile in Spalte A
        $lastRow = $worksheet.Cells($worksheet.Rows.Count, 1).End(-4162).Row + 3 # Letzte Zeile + 1

        # Zelle unterhalb der letzten befüllten Zeile in Spalte A
        $cell = $worksheet.Cells.Item($lastRow, 1)

        # Position für das Logo festlegen (in Pixeln)
        $logoLeft = $cell.Left + 5 # Gleiche Spalte wie Zelle A
        $logoTop = $cell.Top + 1 # Eine Zeile unterhalb

        # Bild einfügen
        $picture = $worksheet.Shapes.AddPicture(
            $logoPath,          # Pfad zum Logo
            $false,             # Verknüpfung mit Datei (False bedeutet, es wird eingebettet)
            $true,              # Bildgröße relativ zum Original beibehalten
            $logoLeft,          # Position von links (entsprechend der Zelle)
            $logoTop,           # Position von oben (entsprechend der Zelle)
            60,                 # Breite des Bildes (anpassbar)
            60                  # Höhe des Bildes (anpassbar)
        )
    }
    # Spalten automatisch anpassen
    $worksheet.UsedRange.Columns.Autofit()
   
    # Daten speichern
    $workBook.SaveAs($excelPath)
    Write-Host "Deine Stunden wurden in deinem Ordner unter $excelPath gespeichert." -ForegroundColor Green
} finally {

    # Excel-Prozess schließen
    $Excel.Quit()
}
 # Schreibschutz entfernen
 Set-ItemProperty -Path $excelPath -Name IsReadOnly -Value $false

# Datei öffnen
Start-Process $excelPath

