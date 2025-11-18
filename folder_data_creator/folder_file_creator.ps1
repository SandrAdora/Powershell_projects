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

# check what the user wants to do
if($whatUserWantsToDo -eq "create")
{

# enter Source path
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
    # Verify if user want to create a file in the created folder 
    $createFile = Read-Host "Do you want to create a file? yes [y] or no [n] "
    if($createFile -notin @("y", "n")){
        Write-Host "Invalid Input. Try again" -ForegroundColor Yellow
    }elseif ($createFile -eq "y") {
        # Get File name
        $fileName = Read-Host "Enter filename with extensions (eg. example.txt)"
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
        # End Script if user declines to create a file in created folder 
        Write-Host "Ending Script. Thank You" -ForegroundColor Green
        break      
    }
}while($true)

}elseif($whatUserWantsToDo -eq "rename")
{
    # Check if the user wants to rename a folder or a file
    do{
    $folderOrFile = Read-Host "Do you want to rename a folder [fd] or file [fl]"
    if($folderOrFile -notin ("fd", "fl"))
    {
        Write-Host "Invalid choice. Try again" -ForegroundColor Red
    }
    if($folderOrFile -eq "fd" -or $folderOrFile -eq "fl")
    {
        # exist loop if either user chose rename folder or file
        break 
    }

    }while($true)
    
    # Search for a specific folder or file name in the directory 
    if($folderOrFile -eq "fd")
    {
        Write-Host "Renaming Folder..." -ForegroundColor Cyan 
        $folderName = Read-Host "Enter the name of the folder "
        # search folder in system
        $
    }
}else{
    Write-Host "No changes" -ForegroundColor Magenta


}
Write-Host "===========================================================================================" -ForegroundColor Yellow
Write-Host ".......Script Folder File Creator ended....." -ForegroundColor Magenta
Write-Host ".......Script Folder File Renaming ended....." -ForegroundColor Magenta

Write-Host "===========================================================================================" -ForegroundColor Yellow


