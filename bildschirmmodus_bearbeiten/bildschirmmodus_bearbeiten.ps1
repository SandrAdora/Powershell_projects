################################################
# Dieses Programm bearbeitet das Bildschirm    #
# lässt sie an wenn dieser von User eingestellt#
# wird                                         #
# Version: 1.0                                 #
# Autor:   Sandra Edigin                       #
# Datum:   12.06.2025                          #
################################################

Write-Host "You have started the Screen settings Script" -ForegroundColor Green

Write-Host "--------------------------------------------------------" -ForegroundColor Magenta
# Get the current display timeout settings for AC (plugged in) and DC (battery)
$acTimeout = powercfg /query SCHEME_CURRENT SUB_VIDEO VIDEOIDLE | 
Select-String -Pattern "Power Setting Index" | ForEach-Object { ($_ -split ':')[1].Trim() }

$dcTimeout = powercfg /query SCHEME_CURRENT SUB_VIDEO VIDEOIDLE | 
Select-String -Pattern "Power Setting Index" | ForEach-Object { ($_ -split ':')[1].Trim() }

Write-Host " Your Current Screen Modus: " -ForegroundColor Yellow
Write-Host "Display timeout when plugged in (AC): $($acTimeout) seconds"
Write-Host "Display timeout on battery (DC): $($dcTimeout) seconds"

Write-Host "--------------------------------------------------------" -ForegroundColor Magenta 
    Write-Host "You can know change the standby modus of your screen" -ForegroundColor Blue

    $choice = Read-Host "Do you want to do this following
    1. For ( Screen On all the time): [on]
    2. For (Back to Default): [off]"

    switch ($choice) {
        "on" { 
            Write-Host "Screen will be on when using network" -ForegroundColor Blue
            powercfg -change -monitor-timeout-ac 0
            Write-Host "Screen will be one when using Batterie" -ForegroundColor Blue
            powercfg -change -monitor-timeout-dc 0

        }
        "off" {
            Write-Host "Changing all Screen Settings back to default modus" -ForegroundColor Blue
            # Bildschirm bei Netzbetrieb nach 10 Minuten ausschalten
            powercfg -change -monitor-timeout-ac 10

            # Bildschirm bei Akkubetrieb nach 5 Minuten ausschalten
            powercfg -change -monitor-timeout-dc 5

        }
        Default {
            Write-Host "No Changes were made" -ForegroundColor Red
        }
    }
    
Write-Host "----------------Script ran sucessfully---------------------------------------" -ForegroundColor Magenta