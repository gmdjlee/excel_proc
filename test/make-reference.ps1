# Generates test/reference.xlsx from the existing VBA workbook.
# Run ONCE. Do not regenerate - this file is the answer key for the SPA's
# fidelity test, and a moving answer key makes the test meaningless.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$xlsm = (Get-ChildItem $root -Filter '*.xlsm').FullName
$csv  = Join-Path $root 'random_data.csv'
$out  = Join-Path $PSScriptRoot 'reference.xlsx'
Remove-Item $out -Force -ErrorAction SilentlyContinue

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1
$wb = $xl.Workbooks.Open($xlsm)
$ok = $xl.Run('BuildFromCSV', $csv, '8/28 111G', 'XX', $out)
"BuildFromCSV = $ok"
$wb.Close($false)
$xl.Quit()
