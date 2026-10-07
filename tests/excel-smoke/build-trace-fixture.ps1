# Builds the Trace In fixture (docs/PLAN.md Phase 4 exit criteria) through COM in a running Excel:
#   %TEMP%\emt-trace-fixture\TraceFixture.xlsx   the workbook to trace (left closed)
#   %TEMP%\emt-trace-fixture\TraceExternal.xlsx  a second workbook it links to (left closed, so Trace In must open it)
# Run with Windows PowerShell 5.1 (needs Marshal.GetActiveObject), with Excel open:
#   powershell -ExecutionPolicy Bypass -File tests/excel-smoke/build-trace-fixture.ps1
# trace-smoke.ps1 calls it with its own Excel object (-Excel). Only the two fixture files are touched: they are
# closed without saving if open, then rebuilt. Excel's iteration setting is turned on while the circular reference
# is written (otherwise Excel warns about it) and put back afterwards.
#
# TraceFixture.xlsx, sheet Calc (every formula the smoke test traces):
#   A1  =B2+B3+B4+B5+B6+B7+B8+B10+B11+B13     the main audited cell: 10 references, all on Calc
#   B1  (empty)                                the undo check types here
#   B2  =Inputs!B2*(1+Growth)                  cross-sheet reference; workbook-level name
#   B3  =B2*Rate                               same-sheet formula cell; name scoped to Calc (Calc!Rate)
#   B4  =SUM(Data!A1:A60)                      a range over 50 cells (paged)
#   B5  =Hidden!A1+TaxRate                     a hidden sheet; a name holding a constant (no cells)
#   B6  =SUM(Sales[Amount])+Sales[[#Totals],[Qty]]  qualified table references (Sales is Data!D1:G7; its
#                                          Amount column, G2:G6, is =[@Qty]*[@Price]: unqualified)
#   B7  =INDIRECT("Inputs!B3")+OFFSET(Inputs!B2,1,0)+INDEX(Data!A1:A60,5)  computed references
#   B8  =B9+1, B9 =B8*0.5                      a circular reference
#   B10 =Data!L2+Data!K1                       a merged cell (Data!K1:L2)
#   B11 =[TraceExternal.xlsx]Rates!B3*2        an external reference (the file is closed, so the path is written)
#   B13 =Inputs!B6                             a hidden row (Inputs row 6)
param(
    [object]$Excel,
    [string]$OutDir = (Join-Path $env:TEMP 'emt-trace-fixture')
)
$ErrorActionPreference = 'Stop'

$xl = if ($Excel) { $Excel } else { [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') }
$mainName = 'TraceFixture.xlsx'
$extName = 'TraceExternal.xlsx'
$mainPath = Join-Path $OutDir $mainName
$extPath = Join-Path $OutDir $extName
$xlOpenXmlWorkbook = 51
$xlSrcRange = 1
$xlYes = 1
$xlSheetHidden = 0

New-Item -ItemType Directory -Force $OutDir | Out-Null
foreach ($name in @($mainName, $extName)) {
    foreach ($open in @($xl.Workbooks)) {
        if ($open.Name -eq $name) { $open.Close($false); "Closed open copy of $name without saving" }
    }
}
foreach ($path in @($mainPath, $extPath)) { if (Test-Path $path) { Remove-Item $path -Force } }

$alerts = $xl.DisplayAlerts
$xl.DisplayAlerts = $false
$ext = $null
$wb = $null
$iteration = $null
try {
    # --- The external workbook: Rates!B3 is a formula, so tracing it goes one level further. ---
    $ext = $xl.Workbooks.Add()
    $rates = $ext.Worksheets.Item(1)
    $rates.Name = 'Rates'
    $rates.Range('A2').Value2 = 'Base rate'
    $rates.Range('B2').Value2 = 0.07
    $rates.Range('A3').Value2 = 'Doubled'
    $rates.Range('B3').Formula = '=B2*2'
    $ext.SaveAs($extPath, $xlOpenXmlWorkbook)
    $extName = $ext.Name

    # --- The main workbook. ---
    $wb = $xl.Workbooks.Add()
    $iteration = $xl.Iteration
    $xl.Iteration = $true   # the circular reference below would otherwise raise Excel's warning
    while ($wb.Worksheets.Count -lt 4) { $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) | Out-Null }
    $calc = $wb.Worksheets.Item(1); $calc.Name = 'Calc'
    $inputs = $wb.Worksheets.Item(2); $inputs.Name = 'Inputs'
    $data = $wb.Worksheets.Item(3); $data.Name = 'Data'
    $hidden = $wb.Worksheets.Item(4); $hidden.Name = 'Hidden'

    # Inputs: values, a hidden row, and the names.
    $inputs.Range('A2').Value2 = 'Revenue';  $inputs.Range('B2').Value2 = 100
    $inputs.Range('A3').Value2 = 'Margin';   $inputs.Range('B3').Value2 = 0.25
    $inputs.Range('A4').Value2 = 'Growth';   $inputs.Range('B4').Value2 = 0.05
    $inputs.Range('A6').Value2 = 'Hidden row'; $inputs.Range('B6').Value2 = 5
    $inputs.Rows.Item(6).Hidden = $true
    $wb.Names.Add('Growth', '=Inputs!$B$4') | Out-Null                 # workbook-level
    $calc.Names.Add('Rate', '=Inputs!$B$3') | Out-Null                 # scoped to Calc
    $wb.Names.Add('TaxRate', '=0.21') | Out-Null                       # a constant: no cells

    # Data: 60 numbers, a table with an unqualified structured reference and a totals row, a merged cell.
    for ($i = 1; $i -le 60; $i++) { $data.Cells.Item($i, 1).Value2 = $i }
    $data.Range('D1').Value2 = 'Region'; $data.Range('E1').Value2 = 'Qty'; $data.Range('F1').Value2 = 'Price'; $data.Range('G1').Value2 = 'Amount'
    $rows = @(@('North', 3, 10), @('South', 5, 12), @('East', 2, 9), @('West', 7, 11), @('Central', 4, 8))
    for ($r = 0; $r -lt $rows.Count; $r++) {
        $data.Cells.Item($r + 2, 4).Value2 = $rows[$r][0]
        $data.Cells.Item($r + 2, 5).Value2 = $rows[$r][1]
        $data.Cells.Item($r + 2, 6).Value2 = $rows[$r][2]
    }
    $table = $data.ListObjects.Add($xlSrcRange, $data.Range('D1:G6'), [Type]::Missing, $xlYes)
    $table.Name = 'Sales'
    $table.ListColumns.Item('Amount').DataBodyRange.Formula = '=[@Qty]*[@Price]'   # unqualified structured references
    $table.ShowTotals = $true                        # the totals row is row 7
    $data.Range('K1').Value2 = 99
    $data.Range('K1:L2').Merge()

    # Hidden sheet.
    $hidden.Range('A1').Value2 = 42
    $hidden.Visible = $xlSheetHidden

    # Calc: the traced formulas. B11 is written while TraceExternal.xlsx is open, so Excel resolves the link.
    $calc.Range('B2').Formula = '=Inputs!B2*(1+Growth)'
    $calc.Range('B3').Formula = '=B2*Rate'
    $calc.Range('B4').Formula = '=SUM(Data!A1:A60)'
    $calc.Range('B5').Formula = '=Hidden!A1+TaxRate'
    $calc.Range('B6').Formula = '=SUM(Sales[Amount])+Sales[[#Totals],[Qty]]'
    $calc.Range('B7').Formula = '=INDIRECT("Inputs!B3")+OFFSET(Inputs!B2,1,0)+INDEX(Data!A1:A60,5)'
    $calc.Range('B8').Formula = '=B9+1'
    $calc.Range('B9').Formula = '=B8*0.5'
    $calc.Range('B10').Formula = '=Data!L2+Data!K1'
    $calc.Range('B11').Formula = "=[$extName]Rates!B3*2"
    $calc.Range('B13').Formula = '=Inputs!B6'
    $calc.Range('A1').Formula = '=B2+B3+B4+B5+B6+B7+B8+B10+B11+B13'
    $calc.Activate()
    $calc.Range('A1').Select() | Out-Null

    $ext.Close($false)   # saved above; now closed, so B11 holds the full path
    $ext = $null
    $wb.SaveAs($mainPath, $xlOpenXmlWorkbook)
    "B11 formula with the external workbook closed: $($calc.Range('B11').Formula)"
    "A1 formula: $($calc.Range('A1').Formula)"
    $wb.Close($false)
    $wb = $null
    "Fixture written: $mainPath"
    "External:        $extPath"
}
finally {
    if ($ext) { try { $ext.Close($false) } catch {} }
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($null -ne $iteration) {
        # Iteration can only be set while a workbook is open.
        $tmp = $xl.Workbooks.Add()
        try { $xl.Iteration = $iteration } finally { $tmp.Close($false) }
    }
    $xl.DisplayAlerts = $alerts
}
