# Creates sample.xlsx with Excel.
# Invoked by ../make-fixtures.ps1; each Office app runs in its own PowerShell process.
param([Parameter(Mandatory)][string]$Out, [string]$Png)
$ErrorActionPreference = 'Stop'

function Release($obj) {
  if ($obj) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Add()
  while ($wb.Worksheets.Count -lt 3) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }

  $ws = $wb.Worksheets.Item(1)
  $ws.Name = '売上'
  $ws.Range('A1').Value2 = '2026年 売上レポート'
  $ws.Range('A1:E1').Merge()
  $ws.Range('A1').Font.Size = 16
  $ws.Range('A1').Font.Bold = $true
  $ws.Range('A1').HorizontalAlignment = -4108 # xlCenter

  $headers = @('商品', '日付', '数量', '単価', '金額')
  for ($i = 0; $i -lt $headers.Count; $i++) { $ws.Cells.Item(3, $i + 1).Value2 = $headers[$i] }
  $header = $ws.Range('A3:E3')
  $header.Font.Bold = $true
  $header.Font.Color = 0xFFFFFF
  $header.Interior.Color = 0x7F3F1F # BGR
  $header.HorizontalAlignment = -4108

  $rows = @(
    @('りんご', '2026-01-15', 120, 98.5),
    @('みかん', '2026-02-03', 80, 45),
    @('Banana', '2026-03-21', 1500, 30.25),
    @('ぶどう（長い商品名のテキスト）', '2026-04-30', 12, 1200)
  )
  for ($r = 0; $r -lt $rows.Count; $r++) {
    $row = 4 + $r
    $ws.Range("A$row").Value2 = $rows[$r][0]
    # PowerShell caches COM property setters by argument type, so numbers go through Formula strings.
    $ws.Range("B$row").Formula = [string]([datetime]$rows[$r][1]).ToOADate()
    $ws.Range("C$row").Formula = [string]$rows[$r][2]
    $ws.Range("D$row").Formula = [string]$rows[$r][3]
    $ws.Range("E$row").Formula = "=C$row*D$row"
  }
  $ws.Range('B4:B7').NumberFormat = 'yyyy/m/d'
  $ws.Range('C4:C7').NumberFormat = '#,##0'
  $ws.Range('D4:E7').NumberFormat = '#,##0.00'
  $ws.Range('A8').Value2 = '合計'
  $ws.Range('A8').Font.Bold = $true
  $ws.Range('E8').Formula = '=SUM(E4:E7)'
  $ws.Range('E8').NumberFormat = '"¥"#,##0'
  $ws.Range('E8').Font.Bold = $true
  $ws.Range('A3:E8').Borders.LineStyle = 1
  $ws.Range('A10').Value2 = '達成率'
  $ws.Range('B10').Formula = '0.875'
  $ws.Range('B10').NumberFormat = '0.0%'
  $ws.Range('A11').Value2 = '有効'
  $ws.Range('B11').Formula = 'TRUE'
  $ws.Range('A12').Value2 = 'エラー'
  $ws.Range('B12').Formula = '=1/0'
  $ws.Range('A13').Value2 = 'This is a long text that spills over into the empty cells to its right'
  $ws.Range('A14').Value2 = "折り返し`nテキスト"
  $ws.Range('A14').WrapText = $true
  $ws.Range('C14').Value2 = 'Red italic'
  $ws.Range('C14').Font.Color = 0x0000FF
  $ws.Range('C14').Font.Italic = $true
  $ws.Columns.Item(1).ColumnWidth = 18
  $ws.Columns.Item(2).ColumnWidth = 12
  $ws.Rows.Item(14).RowHeight = 36
  $ws.Activate()
  $ws.Range('A4').Select()
  $excel.ActiveWindow.FreezePanes = $true

  $chart = $ws.Shapes.AddChart2(201, 51, 420, 30, 320, 200)
  [void]$chart.Chart.SetSourceData($ws.Range('A3:A7,C3:C7'))

  $data = $wb.Worksheets.Item(2)
  $data.Name = 'Data'
  # 5,000 rows x 6 columns of formulas (30,000 cells with cached values).
  $data.Range('A1:A5000').Formula = '="Item "&ROW()'
  $data.Range('B1:F5000').Formula = '=ROW()*(COLUMN()-1)+0.5'

  $hidden = $wb.Worksheets.Item(3)
  $hidden.Name = 'Hidden'
  $hidden.Range('A1').Value2 = 'secret'
  $hidden.Visible = 0

  $wb.SaveAs([string](Join-Path $Out 'sample.xlsx'), 51) # xlOpenXMLWorkbook
  $wb.Close($false)
} finally {
  $excel.Quit()
  Release $excel
}
