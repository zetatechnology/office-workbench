# Opens edited Word / PowerPoint files with the locally installed Office (Windows only) to check
# that they open without repair, and prints their text. Make the files first with
#   node scripts/make-edited-samples.mjs
# or pass -Dir with files saved from the editor. Each file is opened in its own PowerShell process
# with a timeout (a hidden repair prompt would otherwise hang).
param([string]$Dir = (Join-Path ([System.IO.Path]::GetTempPath()) 'office-workbench-edited'), [int]$TimeoutSeconds = 120)
$ErrorActionPreference = 'Stop'
$files = Get-ChildItem $Dir -File | Where-Object { $_.Extension -in '.docx', '.docm', '.pptx', '.pptm' }
if (-not $files) { throw "No .docx / .pptx files in $Dir." }
# Office writes repair logs (error*.xml) to %TEMP%; new ones mean something was repaired.
$logsBefore = @(Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue | ForEach-Object FullName)
$failed = 0

foreach ($file in $files) {
  $app = if ($file.Extension -like '.doc*') { 'word' } else { 'powerpoint' }
  Write-Host "$($file.Name) ($app)"
  $script = Join-Path $PSScriptRoot "check\$app.ps1"
  $proc = Start-Process powershell -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$script`"", '-File', "`"$($file.FullName)`"") -NoNewWindow -PassThru
  [void]$proc.Handle # Caches the handle so ExitCode is available after exit (PowerShell 5.1 quirk).
  if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {
    $proc.Kill()
    Write-Host "  FAILED: timed out after $TimeoutSeconds s (Office may be showing a repair prompt)." -ForegroundColor Red
    $failed++
  } elseif ($proc.ExitCode -ne 0) {
    Write-Host "  FAILED: exit code $($proc.ExitCode)." -ForegroundColor Red
    $failed++
  }
}

$newLogs = @(Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue | Where-Object { $logsBefore -notcontains $_.FullName })
foreach ($log in $newLogs) {
  Write-Host "Repair log: $($log.FullName)" -ForegroundColor Yellow
  Get-Content $log.FullName | Write-Host
  $failed++
}
if ($failed) { throw "$failed problem(s) found." }
Write-Host 'All files opened without repair.' -ForegroundColor Green
