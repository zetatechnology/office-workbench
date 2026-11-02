# Creates sample.docx with Word.
# Invoked by ../make-fixtures.ps1; each Office app runs in its own PowerShell process.
param([Parameter(Mandatory)][string]$Out, [string]$Png)
$ErrorActionPreference = 'Stop'

function Release($obj) {
  if ($obj) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) }
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  $doc = $word.Documents.Add()
  $sel = $word.Selection
  $sel.Style = $doc.Styles.Item(-2) # Heading 1
  $sel.TypeText('Office Workbench サンプル文書')
  $sel.TypeParagraph()
  $sel.Style = $doc.Styles.Item(-1) # Normal
  $sel.TypeText('これは Word の表示テスト用の文書です。')
  $sel.Font.Bold = $true
  $sel.TypeText('太字')
  $sel.Font.Bold = $false
  $sel.TypeText('、')
  $sel.Font.Italic = $true
  $sel.TypeText('斜体')
  $sel.Font.Italic = $false
  $sel.TypeText('、')
  $sel.Font.Color = 255
  $sel.TypeText('赤い文字')
  $sel.Font.Color = -16777216
  $sel.TypeText(' and English text.')
  $sel.TypeParagraph()
  $sel.Style = $doc.Styles.Item(-3) # Heading 2
  $sel.TypeText('箇条書き')
  $sel.TypeParagraph()
  $sel.Style = $doc.Styles.Item(-49) # List Bullet
  foreach ($item in @('一つ目の項目', '二つ目の項目', '三つ目の項目')) { $sel.TypeText($item); $sel.TypeParagraph() }
  $sel.Style = $doc.Styles.Item(-1)
  $sel.TypeText('表:')
  $sel.TypeParagraph()
  $table = $doc.Tables.Add($sel.Range, 3, 3)
  $table.Borders.Enable = $true
  $cells = @(@('項目', '数量', '備考'), @('A', '10', 'テスト'), @('B', '20', ''))
  for ($r = 0; $r -lt 3; $r++) { for ($c = 0; $c -lt 3; $c++) { $table.Cell($r + 1, $c + 1).Range.Text = $cells[$r][$c] } }
  $table.Rows.Item(1).Range.Font.Bold = $true
  $sel.EndKey(6) | Out-Null # wdStory
  $sel.TypeParagraph()
  [void]$sel.InlineShapes.AddPicture([string]$Png)
  $sel.TypeParagraph()
  $sel.InsertBreak(7) # wdPageBreak
  $sel.Style = $doc.Styles.Item(-2)
  $sel.TypeText('2 ページ目')
  $sel.TypeParagraph()
  $sel.Style = $doc.Styles.Item(-1)
  $sel.TypeText('改ページ後のテキストです。')
  $doc.Sections.Item(1).Headers.Item(1).Range.Text = 'ヘッダー'
  $doc.Sections.Item(1).Footers.Item(1).PageNumbers.Add() | Out-Null
  $doc.SaveAs2([string](Join-Path $Out 'sample.docx'), 16) # wdFormatDocumentDefault
  $doc.Close($false)
} finally {
  $word.Quit()
  Release $word
}
