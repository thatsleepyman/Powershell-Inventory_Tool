    #Variables
$FolderPath = 'C:\Inventory_Tool'
$Timestamp = $(get-date -f yyyy-MM-dd_HH-mm-ss)
$comports = (Get-WmiObject -query "SELECT * FROM Win32_PnPEntity" | Where-Object {$_.Name -Match "COM\d+"}).name

$AppPath = "$FolderPath\$env:computername $Timestamp Apps.txt"
$DriverPath = "$FolderPath\$env:computername $Timestamp Drivers.txt"
$USBPath = "$FolderPath\$env:computername $Timestamp USB.txt"
$COMportPath = "$FolderPath\$env:computername $Timestamp COMports.txt"

    #New folder 'C:\Inventory Tool'
New-Item -ItemType Directory -Force -Path $FolderPath
Write-Host "$FolderPath has been created"

    #Applicaties            
New-Item -Path $AppPath -ItemType File
Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |
    Where-Object DisplayName |
    Select-Object -Property DisplayName, Publisher, DisplayVersion, InstallDate |
    Sort-Object -Property DisplayName, Publisher |
    Out-File $AppPath
Write-Host "$env:computername - $Timestamp - Apps.txt"

    #Drivers   
New-Item -Path $DriverPath -ItemType File
    Get-WmiObject Win32_PnPSignedDriver |
    Where-Object DeviceName |
    Select-Object DeviceName, Manufacturer, DriverVersion |
    Sort-Object -Property DeviceName, Manufacturer |
    Out-File $DriverPath
Write-Host "$env:computername - $Timestamp - Drivers.txt has been created"

    #USBs   
New-Item -Path $USBPath -ItemType File
     Get-PnpDevice -class USB |
     Sort-Object -Property Status, Class, FriendlyName |
     Out-File $USBPath
Write-Host "$env:computername - $Timestamp - USB.txt has been created"

    #COMports
if ($null -ne $ports)
    {
    $comports |
    New-Item $COMportPath |
        Out-File $COMportPath
    Write-Host "$env:computername - $Timestamp - COMports.txt has been created"
    }
else
    {
    Write-Host 'No COMports found'
    }

    #Open folder location
explorer.exe $FolderPath