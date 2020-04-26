
# Expenditure Audit Trail Report

## Sample scripts to process individual report file:

```
.\ExpenditureAuditTrail2CSV.ps1 -pdfPath "Samples\FY17_ExpenditureAuditTrail\General Fund
Expenditure Audit Trail - FY 2017.pdf" -outputFolder ".\Samples\FY17_ExpenditureAuditTrail"
```

```
.\ExpenditureAuditTrail2CSV.ps1 -pdfPath "Samples\FY18_ExpenditureAuditTrail\General Fund
Expenditure Audit Trail - FY 2018.pdf" -outputFolder ".\Samples\FY18_ExpenditureAuditTrail"
```

# Vendor Payment History

## Sample script to process individual report file:

.\VendorPaymentHistory2CSV.ps1 -pdfPath "Samples\VendorPaymentHistory\DRDR 23223-23611 (Sep 2016).PDF" -outputFolder ".\Samples\VendorPaymentHistory"

## Sample script to process a directory of report files:
.\wrapper_VendorPaymentHistory2CSV.ps1 -pdfDirectory "Samples\VendorPaymentHistory" -outputFolder "Samples\VendorPaymentHistory"

## Sample script to combine multiple CSV files into a single combined file:
.\combine_csvs.ps1 -csvFolder "Samples\VendorPaymentHistory" -outputFolder "Samples\VendorPaymentHistory"