# EDITABLE
$csvFilePath = "C:\LENOVO_WARRANTIES.csv"
$excelFilePath = "C:\LENOVO_WARRANTIES.xlsx"
$nameOfSerialNumbersColumn = "Serial Numbers"
$excelWriteEndDateRow = 2
$tokenId = "TOKEN_ID"

# Creates a header for authentication for Lenovo API.
$Headers = @{}
$Headers.Add("ClientID", "$($tokenId)")
$Headers.Add("accept", "application/json")
$Headers.Add("content-type", "application/json")

$serialNum = Import-Csv $csvFilePath | Select-Object -ExpandProperty $nameOfSerialNumbersColumn

# Setup excel file for adding end dates.
$excel =  New-Object -ComObject excel.application
$workbook = $excel.Workbooks.Open($excelFilePath)
$excel.DisplayAlerts = $false
$Data = $workbook.Worksheets.Item(1)
$Data.Name = 'Sheet1'

# Checks the serial number through Lenovo's API, and returns the end date.
function LenovoWarranty
{
    Param ([string]$serial)

    $response = Invoke-RestMethod http://supportapi.lenovo.com/v2.5/warranty?Serial=$serial -ContentType "application/JSON" -Headers $Headers
    [string]$e = $response.Warranty.End
    $endDate = $e.Substring(0, $e.IndexOf("T"))
    return $endDate
}

# Loops through the list of serial numbers, and writes them to the excel file "Warranty End Dates" column if no errors happen.
for ($i = 0; $i -lt $serialNum.Count; $i++)
{
    try {
        $warranty = LenovoWarranty $serialNum[$i]
        $Data.Cells.Item($i + 2, $excelWriteEndDateRow) = $warranty
    } catch {
        $Data.Cells.Item($i + 2, $excelWriteEndDateRow) = "ERROR! SN Not Found."
        Write-Host "SN: $($serialNum[$i]) not found. Error found at row $($i + 2)."
    }
}

# Format, save and quit the excel document.
$usedRange = $Data.UsedRange                                                                        
$usedRange.EntireColumn.AutoFit() | Out-Null
$workbook.SaveAs($excelFilePath)
$excel.Quit()