# script parses BOM for code, version and description
# recurses through -bomPath and all sub directories
# writes to csv file

param (
    [Parameter(Mandatory=$False)]
    [string]$bomPath
    )

$ErrorActionPreference = "Stop"

# SETUP
# don't know enough regex to make this a single with an OR
$regExs = @()
$regExs += new-object System.Text.RegularExpressions.Regex('^(?<productCode>[A-Z]{3}[0-9]{4}[GMP]*)[ ]*[v|V](?<versionNumber>[0-9]+)(?<productDescription>.*)\.(doc|pdf|DOC|PDF)$')
$regExs += new-object System.Text.RegularExpressions.Regex('^(?<productCode>[A-Z]{3}[0-9]{4}[GMP]*)[ ]*(?<productDescription>.*)[v|V](?<versionNumber>[0-9]+)\.(doc|pdf|DOC|PDF)$')

$results = @()
$outputCsvFile = 'c:\temp\bom_versions.csv'
$defaultBomPAth = 'C:\Users\darragh\Downloads\'

# for testing
if ([string]::IsNullOrEmpty($bomPath)) {
    $bomPath = $defaultBomPAth
}

$filenames = Get-ChildItem -Path $bomPath -Recurse | Select -exp Name

foreach ($filename in $filenames){
    $details = $null
    Write-Output "Testing file ($filename)"

    foreach ($regex in $regExs){
        $match = $regex.Match($filename)
        
        if ($match.Success) {
            Write-Output 'MATCH!'
            $details = @{            
                Name = $match.Groups["productCode"]              
                Version = $match.Groups['versionNumber']                  
                Description = $match.Groups['productDescription']
                RawFileName = $filename
            }                           
            Write-Output $details
            break
        }
    }
    if ($details -eq $null){
        $details = @{            
                Name = ''              
                Version = ''              
                Description = ''
                RawFileName = $filename
            }                   
    }
    $results += New-Object PSObject -Property $details
}

Write-Output "Writing results to: $outputCsvFile"
$results | Select-Object Name, Version, Description, RawFileName | Export-Csv $outputCsvFile -NoTypeInformation

