###########################################
# Project: Folder File Creator & Renamer  #
# This program creates or renames a       #
# folder/file in a preferred directory    #
#                                         #
# Author:        Sandra E.                #
# Creation date: 08.11.2025               #
# Version:        1.1                     #
###########################################

Write-Host "===========================================================================================" -ForegroundColor Yellow
Write-Host ".......Script Folder File Creator started....." -ForegroundColor Magenta
Write-Host ".......Script Folder File Renaming started....." -ForegroundColor Magenta
Write-Host "===========================================================================================" -ForegroundColor Yellow

# Step 1: Get user preferred root directory
do {
    $sourceDir = Read-Host "Enter preferred root directory (e.g. C:\ )"
    if (Test-Path -Path $sourceDir) {
        break
    } else {
        Write-Host "Attention---> ($sourceDir) does not exist. Try again" -ForegroundColor Yellow
    }
} while ($true)

# Step 2: Ask user what they want to do
do {
    $whatUserWantsToDo = Read-Host "Do you want to CREATE [create] a new folder/file or RENAME [rename] an existing folder/file"
    if ($whatUserWantsToDo -notin ("create", "rename")) {
        Write-Host "Invalid Input. Try again" -ForegroundColor Red
    } else {
        break
    }
} while ($true)

# Step 3: Handle CREATE
if ($whatUserWantsToDo -eq "create") {
    $sourcePath = Read-Host "Enter complete path (including folder name) relative to $sourceDir"
    $pathToFolder = Join-Path $sourceDir $sourcePath

    if (!(Test-Path -Path $pathToFolder)) {
        Write-Host "Path or folder does not exist...creating one....." -ForegroundColor Cyan
        New-Item -ItemType Directory -Path $pathToFolder | Out-Null
        Write-Host "($pathToFolder) was created" -ForegroundColor Green
    }

    do {
        $createFile = Read-Host "Do you want to create a file? yes [y] or no [n]"
        if ($createFile -notin @("y", "n")) {
            Write-Host "Invalid Input. Try again" -ForegroundColor Yellow
        } elseif ($createFile -eq "y") {
            $fileName = Read-Host "Enter filename with extension (e.g. example.txt)"
            $filePath = Join-Path $pathToFolder $fileName

            if (!(Test-Path -Path $filePath -PathType Leaf)) {
                Write-Host "Creating file...." -ForegroundColor Cyan
                New-Item -Path $pathToFolder -Name $fileName -ItemType "File" | Out-Null
                Write-Host "Creation complete, File ($fileName) was created in folder $pathToFolder" -ForegroundColor Green
            } else {
                Write-Host "File already exists at $filePath" -ForegroundColor Yellow
            }
            break
        } else {
            Write-Host "Ending Script. Thank You" -ForegroundColor Green
            break
        }
    } while ($true)

# Step 4: Handle RENAME
} elseif ($whatUserWantsToDo -eq "rename") {
    do {
        $folderOrFile = Read-Host "Do you want to rename a folder [fd] or file [fl]"
        if ($folderOrFile -notin ("fd", "fl")) {
            Write-Host "Invalid choice. Try again" -ForegroundColor Red
        } else {
            break
        }
    } while ($true)

    # Rename Folder
    if ($folderOrFile -eq "fd") {
        Write-Host "Renaming Folder..." -ForegroundColor Cyan
        do {
            $DirectoryToSearchFrom = Read-Host "Enter path to search from or entire folder path"
            
            if (!(Test-Path -Path $DirectoryToSearchFrom)) {
                Write-Host "Directory does not exist." -ForegroundColor Red
            } else {
                break
            }
        } while ($true)

        $fullFolderPathUserChoice = Read-Host "Did you enter the entire folder path? yes [y] no [n]"
        if ($fullFolderPathUserChoice -eq "n") {
            $folderName = Read-Host "Enter the name of the folder to rename"
            $folders = Get-ChildItem -Path $DirectoryToSearchFrom -Recurse -Directory | Where-Object { $_.Name -eq $folderName }

            if ($folders) {
                Write-Host "Folder '$folderName' found." -ForegroundColor Green
                $folderNewName = Read-Host "Enter the new folder name"
                foreach ($folder in $folders) {
                    Rename-Item -Path $folder.FullName -NewName $folderNewName
                    Write-Host "Renamed folder $($folder.Name) to $folderNewName" -ForegroundColor Green
                }
            } else {
                Write-Host "Folder not found" -ForegroundColor Yellow
            }
        }else {
            Write-Host "Entire Folder path was given searching and renaming folder...." -ForegroundColor Cyan
            $newFolderName = Read-Host "Enter new folder name" 
            Rename-Item -Path $DirectoryToSearchFrom -NewName $newFolderName
            Write-Host "Renamed folder $($DirectoryToSearchFrom) to  $($newFolderName)"

        }

    # Rename File
    } elseif ($folderOrFile -eq "fl") {
        Write-Host "Renaming File..." -ForegroundColor Cyan
        do {
            $DirectoryToSearchFrom = Read-Host "Enter directory to search from (or entire file path) File will be searched from $($sourceDir) Directory"           
            if (!(Test-Path $DirectoryToSearchFrom)) {
                Write-Host "Directory does not exist. Try again" -ForegroundColor Red
            } else {
                break
            }
        } while ($true)

        $fullPathInputChoice = Read-Host "Did you enter the entire file path? yes [y] no [n]"
        if ($fullPathInputChoice -eq "n") {
            $oldFileName = Read-Host "Enter old file name with extension (e.g. example.txt)"
            $newFileName = Read-Host "Enter new file name with extension (e.g. newFileName.txt)"

            $files = Get-ChildItem -Path $DirectoryToSearchFrom -Recurse | Where-Object { $_.Name -eq $oldFileName }
            foreach ($file in $files) {
                Write-Host "Renaming file $($file.Name) to $newFileName" -ForegroundColor Cyan
                Rename-Item -Path $file.FullName -NewName $newFileName
                Write-Host "Renamed $($file.FullName) to $($newFileName)" -ForegroundColor Green
            }
        }else{
            Write-Host "You entered the complete file path..." -ForegroundColor Cyan 
            $newFileName = Read-Host "Enter new file name"
            Rename-Item -Path $DirectoryToSearchFrom -NewName $newFileName
            Write-Host "Successfully renamed $($DirectoryToSearchFrom) to $($newFileName)" -ForegroundColor Green
        }
    }
} else {
    Write-Host "No changes made." -ForegroundColor Magenta
}

Write-Host "===========================================================================================" -ForegroundColor Yellow
Write-Host ".......Script Folder File Creator ended....." -ForegroundColor Magenta
Write-Host ".......Script Folder File Renaming ended....." -ForegroundColor Magenta
Write-Host "===========================================================================================" -ForegroundColor Yellow
