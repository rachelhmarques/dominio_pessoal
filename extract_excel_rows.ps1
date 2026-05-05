param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath
)

$ErrorActionPreference = "Stop"
$excel = $null
$workbook = $null

try {
    $resolvedPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($resolvedPath)

    foreach ($worksheet in $workbook.Worksheets) {
        $usedRange = $worksheet.UsedRange
        $rowCount = $usedRange.Rows.Count
        $colCount = $usedRange.Columns.Count

        for ($row = 1; $row -le $rowCount; $row++) {
            $cells = New-Object System.Collections.Generic.List[string]

            for ($col = 1; $col -le $colCount; $col++) {
                $text = [string]$worksheet.Cells.Item($row, $col).Text
                if (-not [string]::IsNullOrWhiteSpace($text)) {
                    [void]$cells.Add($text.Trim())
                }
            }

            if ($cells.Count -gt 0) {
                [PSCustomObject]@{
                    sheet = [string]$worksheet.Name
                    row = $row
                    cells = $cells
                } | ConvertTo-Json -Compress -Depth 4
            }
        }
    }
}
finally {
    if ($workbook) {
        $workbook.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }

    if ($excel) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
