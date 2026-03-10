Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$versionFile = Join-Path $repoRoot "VERSION.txt"

if (-not (Test-Path $versionFile)) {
    throw "VERSION.txt nicht gefunden: $versionFile"
}

$current = (Get-Content $versionFile -Raw).Trim()
if ($current -notmatch '^(?<major>\d+)\.(?<minor>\d+)$') {
    throw "Ungueltiges Versionsformat in VERSION.txt: $current"
}

$major = [int]$Matches.major
$minor = [int]$Matches.minor + 1
$nextVersion = "$major.$minor"

Set-Content -Path $versionFile -Value $nextVersion -Encoding utf8NoBOM

Push-Location $repoRoot
try {
    & .\.venv\Scripts\python.exe -m PyInstaller TreeSolutionHelper.spec
    $srcExe = Join-Path $repoRoot "dist\TreeSolutionHelper.exe"
    if (-not (Test-Path $srcExe)) {
        throw "Gebuildete EXE nicht gefunden: $srcExe"
    }
    $dstExe = Join-Path $repoRoot "dist\TreeSolutionHelper (V$nextVersion).exe"
    Copy-Item -Path $srcExe -Destination $dstExe -Force
    Remove-Item -Path $srcExe -Force
    Write-Host "Version aktualisiert: $current -> $nextVersion"
    Write-Host "Erstellt: $dstExe"
}
catch {
    Set-Content -Path $versionFile -Value $current -Encoding utf8NoBOM
    throw
}
finally {
    Pop-Location
}
