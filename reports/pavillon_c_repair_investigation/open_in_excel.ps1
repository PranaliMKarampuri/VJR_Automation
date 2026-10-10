$ErrorActionPreference = 'Stop'
$diag = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot 'diagnostic_copy.xlsx')).Path
$resaved = Join-Path $PSScriptRoot 'excel_resaved_diagnostic.xlsx'
$log = Join-Path $PSScriptRoot 'excel_open_session.txt'
$start = Get-Date
$before = (Get-FileHash -LiteralPath $diag -Algorithm SHA256).Hash
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $true
    # Opening the investigation copy normally lets Excel display its own recovery prompt.
    $book = $excel.Workbooks.Open($diag)
    "Opened: $($book.Name)" | Set-Content -LiteralPath $log -Encoding UTF8
    "Workbook count: $($excel.Workbooks.Count)" | Add-Content -LiteralPath $log -Encoding UTF8
    "Excel version: $($excel.Version)" | Add-Content -LiteralPath $log -Encoding UTF8
    # Save only to a new diagnostic path, never back to the input copy.
    $book.SaveAs($resaved, 51)
    "Saved recovered/reserialized diagnostic as: $resaved" | Add-Content -LiteralPath $log -Encoding UTF8
    $book.Close($false)
    $book = $null
    $after = (Get-FileHash -LiteralPath $diag -Algorithm SHA256).Hash
    "Diagnostic input hash before/after: $before / $after" | Add-Content -LiteralPath $log -Encoding UTF8
    $candidates = Get-ChildItem -LiteralPath $env:TEMP -File -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -ge $start.AddMinutes(-1) -and ($_.Name -match '^(error|recoveryLog).*\.xml$' -or $_.Name -match 'repair.*\.xml$') }
    foreach ($item in $candidates) {
        "Recovery candidate: $($item.FullName)" | Add-Content -LiteralPath $log -Encoding UTF8
        Get-Content -LiteralPath $item.FullName -Raw | Add-Content -LiteralPath $log -Encoding UTF8
    }
} catch {
    "Excel automation error: $($_ | Out-String)" | Add-Content -LiteralPath $log -Encoding UTF8
    throw
} finally {
    if ($book -ne $null) { try { $book.Close($false) } catch {} }
    if ($excel -ne $null) { try { $excel.Quit() } catch {} }
    if ($book -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
    if ($excel -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
Get-Content -LiteralPath $log
