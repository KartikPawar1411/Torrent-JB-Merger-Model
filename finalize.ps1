# Opens the built workbook in desktop Excel, creates the three data tables,
# recalculates everything and saves. Requires Excel for Windows.
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$path = Join-Path $root "output\Torrent_JB_Merger_Model.xlsx"
$tables = Get-Content (Join-Path $root "output\tables.txt")

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open($path)
    $ws = $wb.Worksheets.Item("Sensitivity")
    foreach ($line in $tables) {
        $p = $line.Split("|")
        [void]$ws.Range($p[1]).Table($ws.Range($p[2]), $ws.Range($p[3]))
    }
    $excel.CalculateFull()
    $wb.Worksheets.Item("Summary").Activate()
    $wb.Save()
    $wb.Close($false)
    Write-Output "finalized $path"
}
finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
