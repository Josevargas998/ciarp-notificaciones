function Inspect-Headers($path) {
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($path)
        
        foreach ($sheet in $workbook.Sheets) {
            Write-Output "Sheet: $($sheet.Name)"
            # Find the header row (typically row 1 or 2)
            $headerRow = 1
            $cols = @()
            for ($c = 1; $c -le 25; $c++) {
                $val = $sheet.Cells.Item(1, $c).Text
                if ($val) { $cols += "$($c): $val" }
            }
            if ($cols.Count -eq 0) {
                # try row 2
                $headerRow = 2
                for ($c = 1; $c -le 25; $c++) {
                    $val = $sheet.Cells.Item(2, $c).Text
                    if ($val) { $cols += "$($c): $val" }
                }
            }
            Write-Output "  Header Row $($headerRow): $($cols -join ' | ')"
        }
        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    } catch {
        Write-Output "Error: $_"
    }
}

Inspect-Headers "C:\Users\Admin\Downloads\docentes\ciarp 2.xlsx"
