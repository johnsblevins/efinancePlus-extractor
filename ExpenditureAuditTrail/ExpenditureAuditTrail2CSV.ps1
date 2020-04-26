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

    # Get Accounting Periods
    if($pageText -match "(?m)^ACCOUNTING PERIODS:  .*$")
    { 
        $accountingPeriods = $Matches[0].Replace("ACCOUNTING PERIODS:  ","")
    } 

    for($line=12; $line -le 46; $line++)
    {
        $lineText = $pageTextArray[$line]

        # Accounting Unit Line
        if ($lineText -match "^\d{1}-\d{1}-\d{3}-\d{3}-\d{4}-\d{4}-\d{16}.*$")
        {
            $accountingUnit = $matches[0]
            $accountingUnitID = $accountingUnit.Substring(0,38)
            $accountingUnitDesc = $accountingUnit.Substring(41)
        }
        # Account Line
        if ($lineText -match "^(?!.*cont'd)\d{3}-\d{4}.*$")
        {
            #while($lineText.contains("  ")){$lineText = $lineText.replace("  ", " ")}
            $accountID = $lineText.Split(" ")[0]
            
            $accountDesc = $lineText.Substring(9,45)
            while($accountDesc.contains("  ")){$accountDesc = $accountDesc.replace("  ", " ")}
            $accountDesc = $accountDesc.TrimEnd(" ")

            $endString = $lineText.substring(46).TrimStart(" ")
            while($endString.contains("  ")){$endString = $endString.replace("  ", " ")}
            $accountStartingBudget = $endString.split(" ")[0]
            $accountBeginExpenditureBalance = $endString.split(" ")[1]
            $accountBeginEncumberanceBalance = $endString.split(" ")[2]
        }

        # Transaction Line
        if ($lineText -match "^   \d{2}/\d{2}/\d{2}.*$")
        {
            $beginString = $lineText.Substring(3,40)
            while($beginString.contains("  ")){$beginString = $beginString.replace("  ", " ")}
            if ($beginString -match "^\d{2}/\d{2}/\d{2}") { $date = $matches[0] }
            if ($beginString -match " \d{2}-\d{1,2} ") { $tc = $matches[0].Replace(" ","") }
            if ($beginString -match " \d{6}-\d{2} ") { $po = $matches[0].Replace(" ","") }
            if ($beginString -match " \d{1,6} ") { $reference = $matches[0].Replace(" ","") }

            $midString = $lineText.Substring(43,25).trimstart(" ").trimend(" ")
            if($midString.Replace(",","") -match "^-?\d*\.{0,1}\d+$")
            {
                $budgetChange = $midString
                $vendorId = ""
                $VendorDesc = ""
            }
            else {
                $budgetChange = ""
                $VendorId = $midString.split(" ")[0].trimend(" ")
                $VendorDesc = $midstring.Substring($midstring.IndexOf(" ")+1).trimend(" ")
            }
            try{
                        $expenditure = $linetext.substring(68,18).replace(" ","")
            }catch{$expenditure = ""}
            try{
                $encumberance = $linetext.substring(86,18).replace(" ","")
            }catch{$encumberance = ""}
            #$lineText
            try{
            $endString = $lineText.Substring(105).trimend(" ")
            $description = $endString            
            }
            catch{
                $description = ""
            }
            $transaction = New-Object System.Object
            $transaction | Add-Member -MemberType NoteProperty -Name "SelectionCritera" -Value $selectionCrieria
            $transaction | Add-Member -MemberType NoteProperty -Name "AccountingPeriods" -Value $accountingPeriods
            $transaction | Add-Member -MemberType NoteProperty -Name "AccountingUnitId" -Value $accountingUnitId
            $transaction | Add-Member -MemberType NoteProperty -Name "accountingUnitDesc" -Value $AccountingUnitDesc
            $transaction | Add-Member -MemberType NoteProperty -Name "AccountID" -Value $accountID
            $transaction | Add-Member -MemberType NoteProperty -Name "AccountDesc" -Value $accountDesc
            $transaction | Add-Member -MemberType NoteProperty -Name "accountStartingBudget" -Value $accountStartingBudget
            $transaction | Add-Member -MemberType NoteProperty -Name "AccountBeginExpenditureBalance" -Value $accountBeginExpenditureBalance
            $transaction | Add-Member -MemberType NoteProperty -Name "AccountBeginEncumberanceBalance" -Value $accountBeginEncumberanceBalance
            $transaction | Add-Member -MemberType NoteProperty -Name "date" -Value $date
            $transaction | Add-Member -MemberType NoteProperty -Name "tc" -Value $tc
            $transaction | Add-Member -MemberType NoteProperty -Name "po" -Value $po
            $transaction | Add-Member -MemberType NoteProperty -Name "reference" -Value $reference
            $transaction | Add-Member -MemberType NoteProperty -Name "VendorId" -Value $VendorId
            $transaction | Add-Member -MemberType NoteProperty -Name "VendorDesc" -Value $VendorDesc
            $transaction | Add-Member -MemberType NoteProperty -Name "BudgetChange" -Value $budgetChange
            $transaction | Add-Member -MemberType NoteProperty -Name "expenditure" -Value $expenditure
            $transaction | Add-Member -MemberType NoteProperty -Name "encumberance" -Value $encumberance
            $transaction | Add-Member -MemberType NoteProperty -Name "description" -Value $description

            $transactions += $transaction
        
            $tc = ""
            $po = ""
            $reference = ""
            $VendorId = ""
            $VendorDesc = ""
            $budgetChange = ""
            $expenditure = ""
            $encumberance = ""
            $description = ""
        }
    }      
}
if( $currentPDFDocument -and -not $currentPDFDocument.IsClosed())
{
    $currentPDFDocument.Close()
}

$outputFile = "$outputFolder\" + $pdfPath.split("\")[-1].Split(".")[0] + ".csv"
$transactions | Export-Csv -Path $outputFile -NoTypeInformation