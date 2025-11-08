###########################################
# Project: Folder File Creator            #
# This programm creates a folder or file  #
#  in a prefered directory location       #
#                                         #
#  Author:        Sandra E.               #
#  Creation date: 08.11.2025              #
# Version:        1.0                     #
###########################################


Write-Host "===========================================================================================" -ForegroundColor Yellow
Write-Host ".......Script Folder File Creator started....." -ForegroundColor Magenta
Write-Host ".......Script Folder File Renaming started....." -ForegroundColor Magenta

Write-Host "===========================================================================================" -ForegroundColor Yellow

# Function section 
# --- Renaming folder
function Rename-Folder{
    param (
        [string]$Path,
        [string]$OldName,
        [string]$NewName,
        [switch]$Rename
    )
    # search for matching folders
    $items = Get-ChildItem -Path $Path -Recurse -Directory | Where-Object {$_.Name -match $OldName}
    foreach($item in $items )
    {
        Write-Host "Found: $($item.FullName)"
        if($Rename)
        {
            $newName = $item.Name -replace $OldName, $NewName
            Rename-Item -Path $item.Name -NewName $newName
            Write-Host "Renamed to: $newName"
        }

    }

}
function Rename-File {
    param (
        [string]$Path,
        [string]$OldFileName,
        [string]$NewFileName
    )
    # search for matching file name and rename them 
        $files = Get-ChildItem -Path $Path -File -Recurse | Where-Object{ $_.Name -match $OldFileName}
        foreach($file in $files){
        $newName = $file.Name -replace $OldFileName, $NewFileName
        Rename-Item -Path $file.Name -NewName $newName
        Write-Host "Renamed $OldFileNam to: $neName"
    }
}
# --------------------------- End of function section
do{
# Get user prefered Dateination

$sourceDir = Read-Host "Enter Prefered root directory (e.g. C:\ )"

# Search if path exists
if(Test-Path -Path $sourceDir)
{
    break # Leave the loop if directory already existis
}
else{
    Write-Host "Attention---> ($sourceDir) does not exist.Try again" -ForegroundColor Yellow
}
}while($true)

# Get what the user wants to do rename or create
do{
$whatUserWantsToDo = Read-Host "Do you want to CREATE [create] a new folder/file or do you want to RENAME  [rename] exsisting folder or file"
if($whatUserWantsToDo -notin ("create", "rename"))
{
    Write-Host "Invalid Input. Try again"
}
else{
    break
}
}while($true)

if($whatUserWantsToDo -eq "create")
{

$sourcePath = Read-Host "Enter <path_to_folder\foldername>"

# check if path exist create if not
$pathToFolder = "$sourceDir\$sourcePath"
if(!(Test-Path -Path $pathToFolder))
{
    # create new folder 
    Write-Host "Path or Foldername does not exist...creating one....." -ForegroundColor Cyan
    New-Item -ItemType Directory -Path $pathToFolder | Out-Null
    Write-Host "($pathToFolder) was created" -ForegroundColor Green
}
do{
    $createFile = Read-Host "Do you want to create a file? yes [y] no [n] "
    if($createFile -notin @("y", "n")){
        Write-Host "Invalid Input. Try again" -ForegroundColor Yellow
    }elseif ($createFile -eq "y") {
        # Get File name
        $fileName = Read-Host "Enter filename (eg. example.txt)"
        # Search if file exist and create if not 
        if(!(Test-Path -Path $fileName -PathType Leaf))
        {
            Write-Host "Creating file...." -ForegroundColor Cyan
            New-Item -Path $pathToFolder -Name $fileName -ItemType "File"
            Write-Host "Creation complete, File ($fileName) was created in folder $pathToFolder" -ForegroundColor Green
            Write-Host "Ending Script. Thank You" -ForegroundColor Green
            break
        }
    }else{
        Write-Host "Ending Script. Thank You" -ForegroundColor Green
        break      
    }
}while($true)

}elseif($whatUserWantsToDo -eq "rename")
{
    $folderOrFile = Read-Host "Do you want to rename a folder [fd] or file [fl]"
    switch ($folderOrFile) {
        "fd" {
            $getFolderName = Read-Host "Enter existing folder in $sourceDir you want to rename"
            $getNewName = Read-Host "Enter newfolder name" 

            # Search for folder name and rename it 
            Rename-Folder -Path $sourceDir -OldName $getFolderName -NewName $getNewName -Rename
            Write-Host "Changed $getFolderName to $getNewName " -ForegroundColor Green
            break
            
        }
        
        "fl" {
            $fileName = Read-Host "Enter file name"
            $newFileName = Read-Host "Enter new file name"
            Rename-File -Path $sourceDir -OldFileName $fileName -NewFileName $newFileName
            Write-Host "Changed $fileName to $newFileName " -ForegroundColor Green
            break

        }
        Default {
            Write-Host "Standard Default no changes "
            break
        }
    }
    # Search for a specific folder or file name in the directory 


}else{
    Write-Host "No changes" -ForegroundColor Magenta


}
Write-Host "===========================================================================================" -ForegroundColor Yellow
Write-Host ".......Script Folder File Creator ended....." -ForegroundColor Magenta
Write-Host ".......Script Folder File Renaming ended....." -ForegroundColor Magenta

Write-Host "===========================================================================================" -ForegroundColor Yellow


