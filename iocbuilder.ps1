Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -Path ".\Microsoft.Office.Interop.Excel.dll"

# Set dimensions for PowerShell window

$Width = 170
$Height = 60
[Console]::SetWindowSize($Width, $Height)

# Banner start
Write-Host ""
Write-Host -ForegroundColor White '8888888 .d88888b.   .d8888b.       888888b.            d8b 888      888                  '
Write-Host -ForegroundColor White '  888  d88P" "Y88b d88P  Y88b      888  "88b           Y8P 888      888                  '
Write-Host -ForegroundColor White '  888  888     888 888    888      888  .88P               888      888                  '
Write-Host -ForegroundColor White '  888  888     888 888             8888888K.  888  888 888 888  .d88888  .d88b.  888d888 '
Write-Host -ForegroundColor White '  888  888     888 888             888  "Y88b 888  888 888 888 d88" 888 d8P  Y8b 888P"   '
Write-Host -ForegroundColor White '  888  888     888 888    888      888    888 888  888 888 888 888  888 88888888 888     '
Write-Host -ForegroundColor White '  888  Y88b. .d88P Y88b  d88P      888   d88P Y88b 888 888 888 Y88b 888 Y8b.     888     '
Write-Host -ForegroundColor White '8888888 "Y88888P"   "Y8888P"       8888888P"   "Y88888 888 888  "Y88888  "Y8888  888     '
#Banner end

Write-Host ""
Write-Host "Please select your file for processing.."

# Prompt user to select file
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
$openFileDialog.Title = "Select a file"
$dialogResult = $openFileDialog.ShowDialog()

if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "No file selected"
    return
}

$excelFilePath = $openFileDialog.FileName

# Load Excel file
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Open($excelFilePath)
$worksheet = $workbook.ActiveSheet

# Find the next available column and insert the URL and IP header
$lastColumn = $worksheet.UsedRange.Columns.Count
$urlColumnIndex = $lastColumn + 1
$urlColumn = $worksheet.Cells.Item(1, $urlColumnIndex)
$urlColumn.Value2 = "URLs"
$ipColumnIndex = $lastColumn + 2
$ipColumn = $worksheet.Cells.Item(1, $ipColumnIndex)
$ipColumn.Value2 = "IP Addresses"

# Get any values that look like a URL or IP address and put them in the new columns
$range = $worksheet.UsedRange
$urlRegex = "(?:https?://|www\.)[\w\-]+(?:\.[\w\-]+)+[/#?]?.*?(?=[\s]|$|\(|\))"
$ipRegex = "\b(?:\d{1,3}\.){3}\d{1,3}\b"
foreach ($cell in $range.Cells) {
    $value = $cell.Value2
    if ($value -match $urlRegex) {
        # Remove parentheses from URL
        $url = $Matches[0] -replace '[()]'
        $urlCell = $worksheet.Cells.Item($cell.Row, $urlColumnIndex)
        $urlCell.Value2 = $url
    }
    if ($value -match $ipRegex) {
        $ipCell = $worksheet.Cells.Item($cell.Row, $ipColumnIndex)
        $ipCell.Value2 = $value
    }
}

# Remove any empty cells under the URL and IP columns
$lastRow = $worksheet.UsedRange.Rows.Count
$urlRange = $worksheet.Range($worksheet.Cells.Item(2, $urlColumnIndex), $worksheet.Cells.Item($lastRow, $urlColumnIndex))
$urlRange.SpecialCells([Microsoft.Office.Interop.Excel.XlCellType]::xlCellTypeBlanks).Delete()
$ipRange = $worksheet.Range($worksheet.Cells.Item(2, $ipColumnIndex), $worksheet.Cells.Item($lastRow, $ipColumnIndex))
$ipRange.SpecialCells([Microsoft.Office.Interop.Excel.XlCellType]::xlCellTypeBlanks).Delete()

Write-Host "Processing complete.."

# Display dialog box to prompt user for file name and path to save
$saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveFileDialog.Title = "Save File"
$saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv"
$saveFileDialog.ShowDialog() | Out-Null

# Get the file name and path from the dialog box
$filePath = $saveFileDialog.FileName
$fileExtension = [System.IO.Path]::GetExtension($filePath)

# Save as either Excel or CSV
if ($fileExtension -eq ".xlsx") {
    $worksheet.SaveAs($filePath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook)
}
elseif ($fileExtension -eq ".csv") {
    $worksheet.SaveAs($filePath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlCSV)
}
else {
    Write-Host "Invalid file type selected"
    return
}

# Clean up
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "File saved as $csvFilePath"

# Kill excel processes that are not visible
Get-Process Excel | Where-Object {$_.MainWindowTitle -eq ''} | Stop-Process