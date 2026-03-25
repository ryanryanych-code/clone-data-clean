param(
  [string]$SourcePath = "",
  [string]$OutputPath = "",
  [switch]$AutoPush,
  [string]$RepoPath = "",
  [string]$RemoteName = "origin",
  [string]$BranchName = "",
  [string]$GitExePath = ""
)

$ErrorActionPreference = "Stop"

if (-not $OutputPath) {
  $projectRoot = Split-Path -Parent $PSScriptRoot
  $OutputPath = Join-Path $projectRoot "autoload-data.js"
}
if (-not $RepoPath) {
  $RepoPath = Split-Path -Parent $PSScriptRoot
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

if ($AutoPush) {
  $resolvedGitExe = $GitExePath
  if (-not $resolvedGitExe) {
    $gitCommand = Get-Command git -ErrorAction SilentlyContinue
    if ($gitCommand) {
      $resolvedGitExe = $gitCommand.Source
    }
  }
  if (-not $resolvedGitExe) {
    $defaultGitExe = "C:\Program Files\Git\cmd\git.exe"
    if (Test-Path -LiteralPath $defaultGitExe) {
      $resolvedGitExe = $defaultGitExe
    }
  }
  if (-not $resolvedGitExe) {
    throw "Git executable not found. Set -GitExePath explicitly."
  }

  if (-not (Test-Path -LiteralPath $RepoPath)) {
    throw "RepoPath not found: $RepoPath"
  }

  & $resolvedGitExe -C $RepoPath rev-parse --is-inside-work-tree *> $null
  if ($LASTEXITCODE -ne 0) {
    throw "RepoPath is not a git repository: $RepoPath"
  }

  $relativeOutput = Resolve-Path -LiteralPath $OutputPath | ForEach-Object { $_.Path.Replace((Resolve-Path -LiteralPath $RepoPath).Path + "\", "") }
  & $resolvedGitExe -C $RepoPath add -- $relativeOutput
  if ($LASTEXITCODE -ne 0) {
    throw "git add failed for $relativeOutput"
  }

  & $resolvedGitExe -C $RepoPath diff --cached --quiet -- $relativeOutput
  if ($LASTEXITCODE -eq 0) {
    Write-Output "No changes to commit for $relativeOutput"
    exit 0
  }
  if ($LASTEXITCODE -ne 1) {
    throw "git diff --cached failed."
  }

  if (-not $BranchName) {
    $detectedBranch = (& $resolvedGitExe -C $RepoPath branch --show-current).Trim()
    if (-not $detectedBranch) {
      throw "Unable to detect current branch. Set -BranchName."
    }
    $BranchName = $detectedBranch
  }

  $commitMessage = "chore: auto update autoload data ($updatedAt)"
  & $resolvedGitExe -C $RepoPath commit -m $commitMessage
  if ($LASTEXITCODE -ne 0) {
    throw "git commit failed."
  }

  & $resolvedGitExe -C $RepoPath push $RemoteName $BranchName
  if ($LASTEXITCODE -ne 0) {
    throw "git push failed."
  }

  Write-Output "Auto push completed: $RemoteName/$BranchName"
}
