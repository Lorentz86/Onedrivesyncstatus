#
#
# Onedrive sync status scipt. 
# Can NOT run as administrator
#
# Download latest "OneDriveLib.dll" from https://github.com/rodneyviana/ODSyncService
# Make a directory or shared folder for powershell to import the module. 
# 
#
# Tested on Windows 10
# Tested with powershell 7.2
# Should work with powershell 5.1
#
# Location of file, change if you want to manually decide where the file os going to be
$Fileloc = "C:\ODtool\OneDriveLib.dll"

# Auto-Download file. Comment the following lines if doing it manual. 
$ODfolder = "C:\ODtool"
if ((Test-Path $ODfile) -eq $true) {Write-Host "File already exist"}
Else { New-Item -ItemType Directory -Path $ODfolder 
$ODuri = "https://github.com/rodneyviana/ODSyncService/releases/download/1.0.0/OneDriveLib.dll"
Invoke-WebRequest -Uri $ODuri -OutFile $ODfile
	}


#Unlocks the file. 
Unblock-File -LiteralPath $Fileloc
Import-Module $Fileloc

# Test location of onedrive. 
# If you have a diffirent location for onedrive uncomment the $ODpath and put the fullpath of onedrive.exe there and comment out the other 2 lines. 
# $ODpath = "D:\filepath\onedrive.exe"
$path = @("%localappdata%\Microsoft\OneDrive\onedrive.exe","C:\Program Files\Microsoft OneDrive\onedrive.exe","C:\Program Files (x86)\Microsoft OneDrive\onedrive.exe")
Foreach ($p in $path) { if ((Test-Path $p) -eq $True) {$ODpath = $p} Else {Write-Host "No Match"}}

# Get onedrive status
$Onedrivestatus = Get-ODStatus

Foreach ($ODstate in $Onedrivestatus) { 
	if ($ODstate.StatusString -like "*error*") {
		$ODstatus = (Get-ODStatus -ByPath $ODstate.LocalPath)
		if ($ODstatus -like "*unknown*") {

# Gives a yes/No option to reset onedrive
			$caption = "Restart Onedrive"    
    $message = "Are you Sure You Want To Proceed:"
    [int]$defaultChoice = 0
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Do the job."
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Do not do the job."
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
    $choiceRTN = $host.ui.PromptForChoice($caption,$message, $options,$defaultChoice)

if ( $choiceRTN -ne 1 )
{
   Start-Process -NoNewWindow -FilePath $ODpath -ArgumentList /reset
		Start-Process $ODpath
}
else
{
   "Host does not require an action"
}
			
	}
		}
	
	Else {Write-Host "Onedrive is functioning properly"}
}
