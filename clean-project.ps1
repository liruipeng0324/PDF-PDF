param(
    [switch]$RemoveInstallers
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = $PSScriptRoot

$directoryTargets = @(
    "test-output"
)

foreach ($name in $directoryTargets) {
    $path = Join-Path $root $name
    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Recurse -Force
        Write-Host "Removed: $name"
    }
}

Remove-Item -LiteralPath (Join-Path $root "web-server.log") -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath (Join-Path $root "web-server-error.log") -Force -ErrorAction SilentlyContinue

Get-ChildItem -Path $root -Recurse -Directory -Filter "__pycache__" -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase) } |
    ForEach-Object {
        Remove-Item -LiteralPath $_.FullName -Recurse -Force
        Write-Host ("Removed cache: " + $_.FullName)
    }

if ($RemoveInstallers) {
    $fileTargets = @(
        "python-3.13.14-amd64.exe",
        "tesseract-ocr-w64-setup-5.4.0.20240606.exe",
        "gs10071w64.exe"
    )

    foreach ($name in $fileTargets) {
        $path = Join-Path $root $name
        if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
            Write-Host "Removed installer: $name"
        }
    }
}
else {
    Write-Host "Installers and tool components were kept."
}

Write-Host "Safe cleanup complete."
