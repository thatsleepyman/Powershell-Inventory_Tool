$rootFolderPath = 'C:\Program Files (x86)', 'C:\Program Files'
$excludeDirectories = ("bin", "obj", "temp");

function Exclude-Directories
{
    process
    {
        $allowThrough = $true
        foreach ($directoryToExclude in $excludeDirectories)
        {
            $directoryText = "*\" + $directoryToExclude
            $childText = "*\" + $directoryToExclude + "\*"
            if (($_.FullName -Like $directoryText -And $_.PsIsContainer) `
                -Or $_.FullName -Like $childText)
            {
                $allowThrough = $false
                break
            }
        }
        if ($allowThrough)
        {
            return $_
        }
    }
}

Clear-Host

Get-ChildItem $rootFolderPath -include *.exe -Recurse `
    | Exclude-Directories