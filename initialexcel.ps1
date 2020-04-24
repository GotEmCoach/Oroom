$ExcelObject=New-Object -ComObject excel.application
$ExcelObject.visible=$true

$ExcelFiles=Get-ChildItem -Path C:\Users\xxxx\Documents\Excel

foreach($ExcelFile in $ExcelFiles){
    $ExcelFile.FullName
}

$Workbook=$ExcelObject.Workbooks.add()
$Worksheet=$Workbook.Sheets.Item("Sheet1")

foreach($ExcelFile in $ExcelFiles){
    $Everyexcel=$ExcelObject.Workbooks.Open($ExcelFile.FullName)
    $Everysheet=$Everyexcel.sheets.item(1)
    $Everysheet.Copy($Worksheet)
    $Everyexcel.Close()
}

$Workbook.SaveAs("C:\Users\xxxx\Documents\excel\merge.xlsx")
$ExcelObject.Quit()