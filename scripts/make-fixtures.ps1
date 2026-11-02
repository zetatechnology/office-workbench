# Generates test fixtures with the locally installed Microsoft Office (Windows only), so the
# files have exactly the structure real Office writes. The generated files are committed;
# rerun this script only when the fixtures need to change.
param([string[]]$Only = @('excel', 'word', 'powerpoint'), [int]$TimeoutSeconds = 180)
$ErrorActionPreference = 'Stop'
$out = (Resolve-Path (Join-Path $PSScriptRoot '..\test-fixtures')).Path
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) "ow-fixtures-$PID"
New-Item -ItemType Directory -Force $tmp | Out-Null

# A small PNG used by the Word and PowerPoint fixtures.
Add-Type -AssemblyName System.Drawing
$png = Join-Path $tmp 'picture.png'
$bmp = New-Object System.Drawing.Bitmap 240, 160
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.Clear([System.Drawing.Color]::FromArgb(255, 230, 240, 255))
$g.FillEllipse([System.Drawing.Brushes]::OrangeRed, 20, 20, 120, 120)
$g.FillRectangle([System.Drawing.Brushes]::SteelBlue, 120, 60, 100, 80)
$g.Dispose()
$bmp.Save($png, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()

try {
  foreach ($app in $Only) {
    Write-Host "Generating with $app..."
    $script = Join-Path $PSScriptRoot "fixtures\$app.ps1"
    $proc = Start-Process powershell -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$script`"", '-Out', "`"$out`"", '-Png', "`"$png`"") -NoNewWindow -PassThru
    [void]$proc.Handle # Caches the handle so ExitCode is available after exit (PowerShell 5.1 quirk).
    if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {
      $proc.Kill()
      throw "$app timed out after $TimeoutSeconds s (Office may be showing a hidden prompt)."
    }
    if ($proc.ExitCode -ne 0) { throw "$app failed with exit code $($proc.ExitCode)." }
  }
} finally {
  Remove-Item -Recurse -Force $tmp
}
Get-ChildItem $out -Filter 'sample.*' | Select-Object Name, Length
