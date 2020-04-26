param(
    [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
    [string]$pdfPath,

    [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
    [string]$outputFolder,

    [Parameter()]
    [string]$pdfExtractorLibPath="libraries"
)         

# Set Variables
$pdfExtractorLibraries = Get-ChildItem $pdfExtractorLibPath
$transactions = @() 

# Load PDF Extractor Libraries
foreach($library in $pdfExtractorLibraries)
{
    [Reflection.Assembly]::LoadFile($library.FullName)    >$null 2>&1
}

# Load File
$reader =  New-Object iText.Kernel.Pdf.PdfReader ((get-item $pdfPath).FullName)
$PDFdocument = New-Object iText.Kernel.Pdf.PdfDocument($reader)
$numPages = $PDFDocument.GetNumberOfPages()
write-host "PDF File ""$pdfPath"" loaded with $numPages Pages"

# For Each Page
for($page=1; $page -le $numPages; $page++)
{
    $i = [math]::floor($page / $numPages * 100)
    Write-Progress -Activity "Processing Pages" -Status "$i% Complete:" -PercentComplete $i

    # Load Page
    $currentPage = $PDFdocument.GetPage($page)     
    $pageText = [iText.Kernel.Pdf.Canvas.Parser.PdfTextExtractor]::GetTextFromPage($currentPage)
    $pageTextArray = $pageText.Split([Environment]::NewLine) 

    # Get Selection Criteria
    if($pageText -match "(?m)^SELECTION CRITERIA: .*$")
    { 
        $selectionCrieria = $Matches[0].Replace("SELECTION CRITERIA: ","")
    }

    for($line=6; $line -le 46; $line++)
    {
        $lineText = $pageTextArray[$line]

        # Vendor Line A - 33556 4IMPRINT 171712 N 69131 1111420401009999-329-2200 ITEM #107814 0.00 1432.00
        if ($lineText -match "^.*\d{16}-\d{3}-\d{4}.*$")
        {
            
            $line_a = $pageTextArray[$line]
            $line_b = $pageTextArray[$line+1]
            if ($line_b -contains "ferpa")
            {
                $line_b = $pageTextArray[$line+2]
            }
            #$page
            #$line
            #$line_a
            #$line_b

            $vendorId = $line_a.Substring(0,8)
            $vendorName = $line_a.Substring(9,22)
            $purchaseOrder = $line_a.Substring(32,12)
            $1099 = $line_a.Substring(45,1)
            $checkNumber = $line_a.Substring(47, 9)
            $accountingUnitId = $line_a.Substring(57,25)
            
            
            if($line_a.Length -gt 106)
            {
                $description= $line_a.Substring(83,22)
                $salesTax = $line_a.Substring(107,11)
                $amount = $line_a.Substring(119,$line_a.Length-119)
            }
            else {
                $salesTax = ""
                $amount = ""
                $description= $line_a.Substring(83,$line_a.Length-83)
            }
            $invoice = $line_b.Substring(12,8)
            $pf = $line_b.Substring(45,1)
            $date = $line_b.Substring(48,8)
            $control = $line_b.Substring(83,$line_b.Length-83)

            $transaction = New-Object System.Object
            $transaction | Add-Member -MemberType NoteProperty -Name "SelectionCritera" -Value $selectionCrieria
            $transaction | Add-Member -MemberType NoteProperty -Name "vendorId" -Value $vendorId
            $transaction | Add-Member -MemberType NoteProperty -Name "vendorName" -Value $vendorName
            $transaction | Add-Member -MemberType NoteProperty -Name "purchaseOrder" -Value $purchaseOrder
            $transaction | Add-Member -MemberType NoteProperty -Name "1099" -Value $1099
            $transaction | Add-Member -MemberType NoteProperty -Name "checkNumber" -Value $checkNumber
            $transaction | Add-Member -MemberType NoteProperty -Name "accountingUnitId" -Value $accountingUnitId
            $transaction | Add-Member -MemberType NoteProperty -Name "description" -Value $description
            $transaction | Add-Member -MemberType NoteProperty -Name "salesTax" -Value $salesTax
            $transaction | Add-Member -MemberType NoteProperty -Name "amount" -Value $amount
            $transaction | Add-Member -MemberType NoteProperty -Name "invoice" -Value $invoice
            $transaction | Add-Member -MemberType NoteProperty -Name "pf" -Value $pf
            $transaction | Add-Member -MemberType NoteProperty -Name "date" -Value $date
            $transaction | Add-Member -MemberType NoteProperty -Name "control" -Value $control
           

            $transactions += $transaction
        
            $vendorId = ""
            $vendorName = ""
            $purchaseOrder = ""
            $1099 = ""
            $checkNumber = ""
            $accountingUnitId = ""
            $description= ""
            $salesTax = ""
            $amount = ""
            $invoice = ""
            $pf = ""
            $date = ""
            $control = ""

        }
    }
}
if( $currentPDFDocument -and -not $currentPDFDocument.IsClosed())
{
    $currentPDFDocument.Close()
}

$outputFile = "$outputFolder\" + $pdfPath.split("\")[-1].Split(".")[0] + ".csv"
$transactions | Export-Csv -Path $outputFile -NoTypeInformation