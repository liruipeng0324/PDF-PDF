param(
    [string]$OutputPath = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not $OutputPath) {
    $OutputPath = Join-Path $PSScriptRoot ("PDF-PDF-portable-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".zip")
}

$staging = Join-Path ([IO.Path]::GetTempPath()) ("PDF-PDF-portable-" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $staging -Force | Out-Null

$excludeNames = @(".git", ".agents", "test-output", "__pycache__")
$excludeExtensions = @(".zip")

Get-ChildItem -LiteralPath $PSScriptRoot -Force | ForEach-Object {
    if ($excludeNames -contains $_.Name) {
        return
    }
    if (-not $_.PSIsContainer -and $excludeExtensions -contains $_.Extension.ToLowerInvariant()) {
        return
    }

    Copy-Item -LiteralPath $_.FullName -Destination $staging -Recurse -Force
}

if (Test-Path -LiteralPath $OutputPath) {
    Remove-Item -LiteralPath $OutputPath -Force
}

Compress-Archive -Path (Join-Path $staging "*") -DestinationPath $OutputPath -Force
Remove-Item -LiteralPath $staging -Recurse -Force

Write-Host "Portable package created:"
Write-Host $OutputPath
