# Creates sample.pptx with PowerPoint.
# Invoked by ../make-fixtures.ps1; each Office app runs in its own PowerShell process.
param([Parameter(Mandatory)][string]$Out, [string]$Png)
$ErrorActionPreference = 'Stop'

function Release($obj) {
  if ($obj) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) }
}

$ppt = New-Object -ComObject PowerPoint.Application
try {
  $pres = $ppt.Presentations.Add(0) # msoFalse: no window
  # Layout indices in the default template: 1 Title, 2 Title and Content, 6 Title Only, 7 Blank.
  $s1 = $pres.Slides.Add(1, 1)
  $s1.Shapes.Item(1).TextFrame.TextRange.Text = 'Office Workbench'
  $s1.Shapes.Item(2).TextFrame.TextRange.Text = 'PowerPoint 表示テスト'
  $s1.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'タイトルスライドのノートです。'

  $s2 = $pres.Slides.Add(2, 2)
  $s2.Shapes.Item(1).TextFrame.TextRange.Text = '箇条書き'
  $body = $s2.Shapes.Item(2).TextFrame.TextRange
  $body.Text = "第一レベルの項目`r第二レベルの項目`r第一レベル（太字）`rLast item"
  $body.Paragraphs(2).IndentLevel = 2
  $body.Paragraphs(3).Font.Bold = $true

  $s3 = $pres.Slides.Add(3, 6)
  $s3.Shapes.Item(1).TextFrame.TextRange.Text = '図形'
  $rect = $s3.Shapes.AddShape(1, 40, 140, 160, 90)        # msoShapeRectangle
  $rect.TextFrame.TextRange.Text = '四角形'
  $oval = $s3.Shapes.AddShape(9, 230, 140, 140, 90)       # msoShapeOval
  $oval.Fill.ForeColor.RGB = 0x3399FF
  $round = $s3.Shapes.AddShape(5, 400, 140, 160, 90)      # msoShapeRoundedRectangle
  $round.Fill.TwoColorGradient(1, 1)
  $round.Fill.ForeColor.RGB = 0xFF9933
  $round.Fill.BackColor.RGB = 0x33CCFF
  $round.TextFrame.TextRange.Text = 'グラデーション'
  $arrow = $s3.Shapes.AddShape(33, 600, 150, 120, 70)     # msoShapeRightArrow
  $star = $s3.Shapes.AddShape(92, 760, 140, 100, 100)     # msoShape5pointStar
  $star.Rotation = 20
  $line = $s3.Shapes.AddLine(40, 300, 400, 380)
  $line.Line.EndArrowheadStyle = 2                        # msoArrowheadTriangle
  $line.Line.Weight = 3
  $box = $s3.Shapes.AddTextbox(1, 450, 300, 400, 80)
  $box.TextFrame.TextRange.Text = 'テキストボックス: 中央揃え'
  $box.TextFrame.TextRange.ParagraphFormat.Alignment = 2  # ppAlignCenter
  $box.TextFrame.TextRange.Font.Size = 24
  $box.TextFrame.TextRange.Font.Color.RGB = 0x0000C0
  [void]$s3.Shapes.AddPicture([string]$Png, 0, -1, 40, 400, 180, 120)

  $s4 = $pres.Slides.Add(4, 6)
  $s4.Shapes.Item(1).TextFrame.TextRange.Text = '表とグラフ'
  $tbl = $s4.Shapes.AddTable(3, 3, 40, 140, 420, 150).Table
  $vals = @(@('地域', '売上', '前年比'), @('東日本', '1,200', '105%'), @('西日本', '980', '97%'))
  for ($r = 0; $r -lt 3; $r++) { for ($c = 0; $c -lt 3; $c++) { $tbl.Cell($r + 1, $c + 1).Shape.TextFrame.TextRange.Text = $vals[$r][$c] } }
  # Charts need a presentation window, which this windowless session lacks, so the chart is
  # best-effort; the unit tests cover the "unsupported content" placeholder separately.
  try { [void]$s4.Shapes.AddChart2(201, 51, 500, 140, 400, 300) } catch { Write-Host '  (chart skipped)' }

  $s5 = $pres.Slides.Add(5, 2)
  $s5.Shapes.Item(1).TextFrame.TextRange.Text = '非表示スライド'
  $s5.Shapes.Item(2).TextFrame.TextRange.Text = 'このスライドは非表示です。'
  $s5.SlideShowTransition.Hidden = -1

  $pres.SaveAs([string](Join-Path $Out 'sample.pptx'), 24) # ppSaveAsOpenXMLPresentation
  $pres.Close()
} finally {
  $ppt.Quit()
  Release $ppt
}
