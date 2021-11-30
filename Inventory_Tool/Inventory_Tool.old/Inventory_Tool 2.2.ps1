#Variables
$FolderPath = 'C:\Windows 7 Inventarisatie'
$ApplicatiePath = 'C:\Windows 7 Inventarisatie\newtextfile.txt'
$DriversPath = 'C:\Windows 7 Inventarisatie\windrivers.txt'
$COMportsPath = 'C:\Windows 7 Inventarisatie\COMports.txt'

#Creates a folder if it doesn't exist yet
If(!(test-path $FolderPath))
{
    #Creates a new folder and leaves a message after
    New-Item -ItemType Directory -Force -Path $FolderPath
    Write-Host "
    New Folder has been made at:" $FolderPath
}
else
{
    Write-Host "Folder already exists at:" $FolderPath
}

#Creates new text file
New-Item -Path $ApplicatiePath -ItemType File
#Finds all installed applications on pc
Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object DisplayName |
Select-Object -Property DisplayName, DisplayVersion, Publisher, InstallDate | Sort-Object -Property DisplayName |
#Saves the application information to the newly created text file
Out-File $ApplicatiePath
#Renames text file after the computer's name
Rename-Item -Path $ApplicatiePath -NewName "$env:computername Applicaties.txt"

#Creates new text file
New-Item -Path $DriversPath -ItemType File
#Finds all drivers on device
Get-WmiObject Win32_PnPSignedDriver | Where-Object DeviceName | 
Select-Object DeviceName, Manufacturer, DriverVersion | Sort-Object -Property DeviceName, Manufacturer | Sort-Object -Property DeviceName |
#Saves the application information to the newly created text file
Out-File $DriversPath
#Renames text file after the computer's name
Rename-Item -Path $DriversPath -NewName "$env:computername Drivers.txt"

#Check for COMports
$ports = (Get-WmiObject -query "SELECT * FROM Win32_PnPEntity" | Where-Object {$_.Name -Match "COM\d+"}).name
if ($null -ne $ports)
{
    #Creates a new .txt file and saves the COMport data to it
    $ports | Sort-Object -Property Size_MB -Descending | New-Item $COMportsPath |
    Out-File $COMportsPath; Rename-Item -Path $COMportsPath -NewName "$env:computername COMPORT.txt"
}
Else
{
    #If no COMport's have been found it leaves a message
    Write-Host "No COM Ports found"
}