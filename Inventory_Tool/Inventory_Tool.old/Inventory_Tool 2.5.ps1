#First run the Set-ExecutionPolicy.ps1 script
    
    #Variables
$FolderPath = $PSScriptRoot + "\Inventory_Tool_Data"
$Timestamp = $(get-date -f yyyy-MM-dd_HH-mm-ss)
$Comports = (Get-WmiObject -query "SELECT * FROM Win32_PnPEntity" | Where-Object {$_.Name -Match "COM\d+"}).name

    $AppPath = "$FolderPath\$env:computername $Timestamp Apps.txt"
    $DriverPath = "$FolderPath\$env:computername $Timestamp Drivers.txt"
    $USBPath = "$FolderPath\$env:computername $Timestamp USB.txt"
    $COMportPath = "$FolderPath\$env:computername $Timestamp COMports.txt"

    #New folder 'Directory\Inventory_Tool_Data'
New-Item -ItemType Directory -Force -Path $FolderPath

    #Applicaties
#If you want .csv just # the Out-File and remove # from Export-Csv and change .txt to .csv
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

    #USBs
New-Item -Path $USBPath -ItemType File
     Get-PnpDevice -class USB |
     Sort-Object -Property Status, Class, FriendlyName |
     Out-File $USBPath

    #COMports
if ($null -ne $Comports)
    {
    $Comports |
    New-Item $COMportPath |
        Out-File $COMportPath
    Write-Host "$env:computername $Timestamp COMports.txt has been created"
    }
else
    {
    Write-Host 'No COMports found'
    }

    #Open folder location
if (-not(Get-Process explorer | Where-Object{$_.MainWindowTitle -eq ($FolderPath | Split-Path -Leaf)}))
    {
    explorer.exe $FolderPath
    }