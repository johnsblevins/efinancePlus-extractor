param(
    [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
    [string]$csvFolder,

    [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
    [string]$outputFolder
)         

Get-ChildItem -path $csvFolder -Filter *.csv | Select-Object -ExpandProperty FullName | Import-Csv | Export-Csv $outputFolder\combinedcsvs.csv -NoTypeInformation -Append
