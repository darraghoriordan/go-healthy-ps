# script parses BOM for code, version and description
# recurses through -bomPath and all sub directories
# writes to csv file

param (
    [Parameter(Mandatory=$False)]
    [string]$bomPath
    )

$ErrorActionPreference = "Stop"

# SETUP
$fileSuffix = '*.doc'
$matchRegex = '^([A-Z]{3}[0-9]{4})[ ]*[v|V]([0-9]+)(.*)\.doc$'
$results = @()
$outputCsvFile = 'c:\temp\bom_versions.csv'
$defaultBomPAth = 'C:\Users\'


# for testing
if ([string]::IsNullOrEmpty($bomPath)) {
    $bomPath = $defaultBomPAth
}


$filenames = Get-ChildItem -Path $bomPath -Recurse -Include *.doc | Select -exp Name

foreach ($filename in $filenames){

 Write-Output "Testing file ($filename)"
    $match = [regex]::Match($filename, $matchRegex)
    $details = $null
    if ($match.Success) {
        $details = @{            
                Name             = $match.Captures.Groups[1]              
                Version     = $match.Captures.Groups[2]                  
                Description      = $match.Captures.Groups[3]  
        }                           
        $results += New-Object PSObject -Property $details  
    }
    if ($details -ne $null ) {
        Write-Output 'MATCH!'
        Write-Output $details}
    else {
        Write-Warning 'file did not match ($filename)'
    }
}

Write-Output "Writing results to: $outputCsvFile"
$results | Select-Object Name, Version, Description | Export-Csv $outputCsvFile -NoTypeInformation

