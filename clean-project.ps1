Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = $PSScriptRoot
$fileTargets = @(
    "python-3.13.14-amd64.exe",
    "tesseract-ocr-w64-setup-5.4.0.20240606.exe",
    "gs10071w64.exe"
)

foreach ($name in $fileTargets) {
    $path = Join-Path $root $name
    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
        Write-Host "Removed: $name"
    }
}

$directoryTargets = @(
    "test-output",
    "__pycache__"
)

foreach ($name in $directoryTargets) {
    $path = Join-Path $root $name
    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Recurse -Force
        Write-Host "Removed: $name"
    }
}

Get-ChildItem -Path $root -Recurse -Directory -Filter "__pycache__" -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase) } |
    ForEach-Object {
        Remove-Item -LiteralPath $_.FullName -Recurse -Force
        Write-Host ("Removed cache: " + $_.FullName)
    }

Write-Host "Cleanup complete."

