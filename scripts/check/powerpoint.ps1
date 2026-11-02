# Opens one presentation read-only with PowerPoint and prints its text; fails if PowerPoint
# cannot open it as is. Invoked by ../check-office.ps1.
param([Parameter(Mandatory)][string]$File)
$ErrorActionPreference = 'Stop'

# PowerPoint has a single instance; only quit it when this script started it.
$running = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
$ppt = New-Object -ComObject PowerPoint.Application
try {
  $pres = $ppt.Presentations.Open([string]$File, -1, 0, 0) # ReadOnly, not Untitled, no window
  Write-Host "  slides: $($pres.Slides.Count)"
  foreach ($slide in $pres.Slides) {
    $texts = @()
    foreach ($shape in $slide.Shapes) {
      if ($shape.HasTextFrame -and $shape.TextFrame.HasText) { $texts += $shape.TextFrame.TextRange.Text.Replace([string][char]13, ' / ') }
      if ($shape.HasTable) {
        $table = $shape.Table
        for ($r = 1; $r -le $table.Rows.Count; $r++) {
          $cells = @()
          for ($c = 1; $c -le $table.Columns.Count; $c++) { $cells += $table.Cell($r, $c).Shape.TextFrame.TextRange.Text.Replace([string][char]13, ' / ') }
          $texts += '[' + ($cells -join ' | ') + ']'
        }
      }
    }
    Write-Host "  $($slide.SlideIndex): $($texts -join ' ; ')"
  }
  $pres.Close()
} finally {
  if (-not $running) { $ppt.Quit() }
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)
}
