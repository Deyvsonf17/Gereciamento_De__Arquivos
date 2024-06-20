param (
    [string]$filePath,
    [string]$sheet,
    [string]$cell
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$workbook = $excel.Workbooks.Open($filePath)

# Ensure the workbook is used
$null = $workbook

$excel.Run("GoToSheetAndCell", $sheet, $cell)


