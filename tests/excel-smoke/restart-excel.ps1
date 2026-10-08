# Restarts Excel so a rebuilt add-in can load: Excel keys its add-in list by file name and keeps a loaded .xll locked
# for its lifetime, so a new build at any path is never loaded until Excel restarts. Refuses if any workbook has
# unsaved changes (nothing is ever saved here). Windows PowerShell 5.1.
#   powershell -ExecutionPolicy Bypass -File tests/excel-smoke/restart-excel.ps1 [-Stop | -Start]
# -Stop only closes Excel and -Start only starts it: build between the two, since Excel loads the installed
# add-in at startup and locks it before a build could replace it. Without a switch it does both.
param([switch]$Stop, [switch]$Start)
$ErrorActionPreference = 'Stop'
if (-not $Stop -and -not $Start) { $Stop = $true; $Start = $true }
$xl = $null
try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { }
if ($Stop -and $xl) {
  $unsaved = @($xl.Workbooks | Where-Object { -not $_.Saved } | ForEach-Object { $_.Name })
  if ($unsaved.Count -gt 0) { throw "ABORT: unsaved changes in: $($unsaved -join ', '). Excel was not closed." }
  "Closing Excel (open, all saved: $(@($xl.Workbooks | ForEach-Object { $_.Name }) -join ', '))"
  $pid0 = 0; try { $pid0 = (Get-Process EXCEL | Select-Object -First 1).Id } catch { }
  $xl.DisplayAlerts = $false
  $xl.Quit()
  [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
  $xl = $null
  for ($i = 0; $i -lt 30 -and (Get-Process EXCEL -ErrorAction SilentlyContinue); $i++) { Start-Sleep -Milliseconds 500 }
  if (Get-Process EXCEL -ErrorAction SilentlyContinue) { "Excel did not exit; killing it (all workbooks were saved)"; Stop-Process -Name EXCEL -Force; Start-Sleep -Seconds 2 }
}
if (-not $Start) { "Excel is closed."; return }
Start-Process excel.exe
for ($i = 0; $i -lt 60; $i++) {
  Start-Sleep -Milliseconds 500
  try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application'); if ($xl.Ready) { break } } catch { }
}
if (-not $xl) { throw "Excel did not come back within 30 s." }
"Excel restarted: version $($xl.Version) build $($xl.Build) pid $((Get-Process EXCEL | Select-Object -First 1).Id); workbooks: $(@($xl.Workbooks | ForEach-Object { $_.Name }) -join ', ')"
