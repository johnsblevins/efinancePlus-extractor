param(
    [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
    [string]$pdfDirectory,

    [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
    [string]$outputFolder
)         

get-childitem -Path $pdfDirectory -Filter "*.pdf" | foreach{
.\VendorPaymentHistory2CSV.ps1 -pdfPath $_.FullName -outputFolder $outputFolder
}