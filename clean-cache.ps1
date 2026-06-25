Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = $PSScriptRoot
$removed = 0

Get-ChildItem -Path $root -Recurse -Directory -Filter "__pycache__" -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase) } |
    ForEach-Object {
        Remove-Item -LiteralPath $_.FullName -Recurse -Force
        $script:removed += 1
        Write-Host ("Removed cache: " + $_.FullName)
    }

Write-Host "Cache cleanup complete. Removed folders: $removed"

