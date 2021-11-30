
#First run the Set-ExecutionPolicy.ps1 script

#Variables
$Folder_data = $PSScriptRoot + '\Inventory_Tool_Data'
$Folder_pcname = "$Folder_data\$env:computername"
$Folder_timestamp = "$Folder_pcname\$Timestamp"

$Timestamp = $(get-date -f yyyy-MM-dd_HH-mm-ss)
$FileName = "$Folder_timestamp\$Timestamp"
$Comports = (Get-WmiObject -query "SELECT * FROM Win32_PnPEntity" | Where-Object {$_.Name -Match "COM\d+"}).name

$AppPath = "$FileName Applicaties.txt"
$DriverPath = "$FileName Drivers.txt"
$USBPath = "$FileName USB Poorten.txt"
$COMportPath = "$FileName COMports.txt"

#New folders 'Directory\Inventory_Tool_Data\pcname\pcname timestamp\x.txt'
New-Item -ItemType Directory -Force -Path $Folder_data
New-Item -ItemType Directory -Force -Path $Folder_pcname
New-Item -ItemType Directory -Force -Path $Folder_timestamp

#Applicaties
#If you want .csv files instead, just # the Out-File and remove # from Export-Csv and change .txt to .csv
New-Item -Path $AppPath -ItemType File
Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |
    Where-Object DisplayName |
    Select-Object -Property DisplayName, Publisher, DisplayVersion, InstallDate |
    Sort-Object -Property DisplayName, Publisher |
Out-File $AppPath
#Export-Csv $AppPath

#Drivers
New-Item -Path $DriverPath -ItemType File
Get-WmiObject Win32_PnPSignedDriver |
    Where-Object DeviceName |
    Select-Object DeviceName, Manufacturer, DriverVersion |
    Sort-Object -Property DeviceName, Manufacturer |
Out-File $DriverPath

#USB Poorten
New-Item -Path $USBPath -ItemType File
Get-PnpDevice -class USB |
     Sort-Object -Property Status, Class, FriendlyName |
Out-File $USBPath

#COMports
if ($null -ne $Comports) {
    New-Item -Path $COMportPath -ItemType File
    $Comports |
        Out-File $COMportPath
    }

#Open folder location
if (-not(Get-Process explorer | Where-Object{$_.MainWindowTitle -eq ($Folder_data | Split-Path -Leaf)})) {
    explorer.exe $Folder_data
    }