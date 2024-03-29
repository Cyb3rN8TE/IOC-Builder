Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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

# Show an open file dialog to let the user select a CSV file
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Filter = "CSV Files (*.csv)|*.csv"
$openFileDialog.Title = "Select CSV file"
$openFileDialog.ShowDialog() | Out-Null

if ([string]::IsNullOrEmpty($openFileDialog.FileName)) {
    Write-Host "No file selected. Exiting script."
    Exit
}

# Ask the user if they want to defang or refang
$choice = Read-Host "Do you want to defang or refang the data? Enter 'd' for defang or 'r' for refang"

# Check if the user entered a valid choice
if ($choice -ne 'd' -and $choice -ne 'r') {
    Write-Host "Invalid choice. Exiting script."
    Exit
}

# Load the CSV file
$data = Import-Csv $openFileDialog.FileName

# Process all rows, including the header row
for ($i = 0; $i -lt $data.Count; $i++) {
    $row = $data[$i]
    foreach ($column in $row.PSObject.Properties) {
        if ($column.Value -is [string]) {
            if ($choice -eq 'd') {
                # Defang by replacing '.' with '[.]'
                $column.Value = $column.Value -replace '\.', '[.]'
            } else {
                # Refang by finding '[.]' and replacing with '.'
                $column.Value = $column.Value -replace '\[\.\]', '.'
            }
        }
    }
}

# Show a save file dialog to let the user save the modified CSV as a new file
$saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
$saveFileDialog.Title = "Save modified CSV file"
$saveFileDialog.ShowDialog() | Out-Null

if ([string]::IsNullOrEmpty($saveFileDialog.FileName)) {
    Write-Host "No file selected. Exiting script."
    Exit
}

$data | Export-Csv $saveFileDialog.FileName -NoTypeInformation

Write-Host "CSV file has been modified and saved as $($saveFileDialog.FileName)"
