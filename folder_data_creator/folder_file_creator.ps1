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


$sourcePath = Read-Host "Enter compete path (including the folder name) you would want to save the folder in"

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
    $createFile = Read-Host "Do you want to create a file? yes [y] or no [n] "
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
            $items = Get-ChildItem -Path $sourceDir -Recurse | Where-Object {$_.Name -eq $getFolderName}
            foreach($item in $items)
            {
                Rename-Item -Path $item.Name -NewName  $getNewName
            }
            Write-Host "Changed $getFolderName to $getNewName " -ForegroundColor Green
            break
            
        }
        
        "fl" {
            $fileName = Read-Host "Enter file name"
            $newFileName = Read-Host "Enter new file name"
            $items = Get-ChildItem -Path $sourceDir -Recurse | Where-Object {$_.Name -eq $fileName}
            foreach($item in $items)
            {
                Rename-Item -Path $item.Name -NewName  $newFileName
            }
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


