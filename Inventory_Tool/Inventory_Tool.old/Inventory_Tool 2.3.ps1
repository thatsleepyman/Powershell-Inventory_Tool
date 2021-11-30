    #Variables
$FolderPath = 'C:\Windows 7 Inventarisatie'
$ApplicatiePath = "$FolderPath\$env:computername Applicaties.txt"
$DriversPath = "$FolderPath\$env:computername Drivers.txt"
$USBConnectPath = "$FolderPath\$env:computername USBConnections.txt"
$COMportsPath = "$FolderPath\$env:computername COMports.txt"

    #New folder 'C:\Windows 7 Inventarisatie'
If(!(test-path $FolderPath))
{
    New-Item -ItemType Directory -Force -Path $FolderPath
}
else
{
    Write-Host "Catch Error Folder"
}

    #Applicaties
If(!(test-path $ApplicatiePath))
{             
    New-Item -Path $ApplicatiePath -ItemType File
    Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object DisplayName | Select-Object -Property DisplayName, Publisher, DisplayVersion, InstallDate | Sort-Object -Property DisplayName, Publisher | Out-File $ApplicatiePath
}
else
{
    Write-Host "Catch Error Applicaties"
}

    #Drivers
If(!(test-path $DriversPath))
{    
    New-Item -Path $DriversPath -ItemType File
    Get-WmiObject Win32_PnPSignedDriver | Where-Object DeviceName | Select-Object DeviceName, Manufacturer, DriverVersion | Sort-Object -Property DeviceName, Manufacturer | Out-File $DriversPath
}
else
{
    Write-Host "Catch Error Drivers"
}

    #USBs
If(!(test-path $USBConnectPath))
{    
    New-Item -Path $USBConnectPath -ItemType File
    Get-PnpDevice | Sort-Object -Property Status, Class, FriendlyName | Out-File $USBConnectPath
}
else
{
    Write-Host "Catch Error USBs"
} 

    #COMports
$ports = (Get-WmiObject -query "SELECT * FROM Win32_PnPEntity" | Where-Object {$_.Name -Match "COM\d+"}).name
if ($null -ne $ports)
{
    $ports | Sort-Object -Property Size_MB -Descending | New-Item $COMportsPath | Out-File $COMportsPath
}
Else
{
    Write-Host "Catch Error COMports"
}

    #Open folder location
explorer.exe $FolderPath