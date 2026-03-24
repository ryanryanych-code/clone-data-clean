param(
  [string]$SourcePath = "",
  [string]$OutputPath = ""
)

$ErrorActionPreference = "Stop"

if (-not $OutputPath) {
  $projectRoot = Split-Path -Parent $PSScriptRoot
  $OutputPath = Join-Path $projectRoot "autoload-data.js"
}

if (-not $SourcePath) {
  $desktopPath = [Environment]::GetFolderPath("Desktop")
  $candidate = Get-ChildItem -LiteralPath $desktopPath -File -Filter "M12-M2*.csv" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
  if ($null -eq $candidate) {
    throw "No M12-M2*.csv file found on Desktop."
  }
  $SourcePath = $candidate.FullName
}

if (-not (Test-Path -LiteralPath $SourcePath)) {
  throw "Source file not found: $SourcePath"
}

$csvText = Get-Content -LiteralPath $SourcePath -Raw -Encoding UTF8
$base64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($csvText))
$sourceName = [System.IO.Path]::GetFileName($SourcePath)
$updatedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

$safeSourcePath = $SourcePath.Replace("\", "\\").Replace("`"", "\`"")
$safeSourceName = $sourceName.Replace("`"", "\`"")
$safeUpdatedAt = $updatedAt.Replace("`"", "\`"")

$content = @"
window.AUTOLOAD_SOURCE_PATH = "$safeSourcePath";
window.AUTOLOAD_SOURCE_NAME = "$safeSourceName";
window.AUTOLOAD_UPDATED_AT = "$safeUpdatedAt";
window.AUTOLOAD_CSV_BASE64 = "$base64";
"@

Set-Content -LiteralPath $OutputPath -Value $content -Encoding UTF8
Write-Output "Autoload data file updated: $OutputPath"
