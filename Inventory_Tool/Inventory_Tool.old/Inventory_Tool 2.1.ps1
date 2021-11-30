#Creates a folder if it doesn't exist yet
$FolderPath = '\\gemeentenet.local\dfs\EQU\SSC\34-Agile Teams\007. Team Werkplek\Medewerkers\Stefan\test\Windows 7 Inventarisatie'
If(!(test-path $FolderPath))
{
    New-Item -ItemType Directory -Force -Path $FolderPath
    Write-Host "
    New Folder has been made at:" $FolderPath
}
else
{
    Write-Host "Folder already exists at:" $FolderPath   
}

$FolderPath2 = '\\gemeentenet.local\dfs\EQU\SSC\34-Agile Teams\007. Team Werkplek\Medewerkers\Stefan\test\Windows 7 Inventarisatie\Desktop'
New-Item -ItemType Directory -Force -Path $FolderPath2

#Creates new text file
New-Item -Path '\\gemeentenet.local\dfs\EQU\SSC\34-Agile Teams\007. Team Werkplek\Medewerkers\Stefan\test\Windows 7 Inventarisatie\Desktop\newtextfile.txt' -ItemType File

#Finds all installed applications on pc
Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object DisplayName |
Select-Object -Property DisplayName, DisplayVersion, Publisher, InstallDate | Sort-Object -Property DisplayName | 

#Saves the application information to the newly created text file
Out-File '\\gemeentenet.local\dfs\EQU\SSC\34-Agile Teams\007. Team Werkplek\Medewerkers\Stefan\test\Windows 7 Inventarisatie\Desktop\newtextfile.txt'
#Renames text file after the computer's name
Rename-Item -Path '\\gemeentenet.local\dfs\EQU\SSC\34-Agile Teams\007. Team Werkplek\Medewerkers\Stefan\test\Windows 7 Inventarisatie\Desktop\newtextfile.txt'-NewName $env:computername+.txt

#Bit of text display
Write-Host "
Device Name"
Write-Host "-----------"
#Shows computer's name
$env:computername | Select-Object

#Shows all applications on pc in Powershell
Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object DisplayName |
Select-Object -Property DisplayName, DisplayVersion, Publisher,InstallDate, Version | Sort-Object -Property DisplayName | Format-Table -Autosize

#Bit of a text display :3
Write-Host "COM Ports"
Write-Host "---------"
#Check for COM Ports
$ports = (Get-WmiObject -query "SELECT * FROM Win32_PnPEntity" | Where {$_.Name -Match "COM\d+"}).name
if ($null -ne $ports)
{
    #Creates a new .txt file and saves the COMport data to it
    $ports | Sort-Object -Property Size_MB -Descending | New-Item '\\gemeentenet.local\dfs\EQU\SSC\34-Agile Teams\007. Team Werkplek\Medewerkers\Stefan\test\Windows 7 InventarisatieTest Folder\Desktop\COMports.txt' |
    Out-File '\\gemeentenet.local\dfs\EQU\SSC\34-Agile Teams\007. Team Werkplek\Medewerkers\Stefan\test\Windows 7 Inventarisatie\Desktop\COMports.txt'; Rename-Item -Path '\\gemeentenet.local\dfs\EQU\SSC\34-Agile Teams\007. Team Werkplek\Medewerkers\Stefan\test\Windows 7 Inventarisatie\Desktop\COMports.txt' -NewName $env:computername+COMPORT.txt
} 
Else
{
    #If no COMport's have been found
    Write-Host "No COM Ports found
    "
}

Rename-Item -Path '\\gemeentenet.local\dfs\EQU\SSC\34-Agile Teams\007. Team Werkplek\Medewerkers\Stefan\test\Windows 7 Inventarisatie\Desktop'-NewName $env:computername