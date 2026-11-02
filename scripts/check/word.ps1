# Opens one document read-only with Word and prints its text; fails if Word cannot open it as is.
# Invoked by ../check-office.ps1.
param([Parameter(Mandatory)][string]$File)
$ErrorActionPreference = 'Stop'

# Word may already be running for the user; only quit the instance this script started.
$running = [bool](Get-Process WINWORD -ErrorAction SilentlyContinue)
$word = New-Object -ComObject Word.Application
$word.DisplayAlerts = 0
try {
  # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, ..., OpenAndRepair is not requested)
  $doc = $word.Documents.Open([string]$File, $false, $true, $false)
  Write-Host ("  paragraphs: {0}, tables: {1}" -f $doc.Paragraphs.Count, $doc.Tables.Count)
  foreach ($line in $doc.Content.Text.Split([char]13)) { if ($line.Trim()) { Write-Host "  | $line" } }
  $doc.Close($false)
} finally {
  if (-not $running) { $word.Quit() }
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
