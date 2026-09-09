Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$installer = Get-Content -Raw -LiteralPath (Join-Path $root "install-components.ps1")
$readme = Get-Content -Raw -LiteralPath (Join-Path $root "README.md")
$runtime = Get-Content -Raw -LiteralPath (Join-Path $root "make-dual-pdf.ps1")
$guidePath = Get-ChildItem -LiteralPath $root -Filter "*.md" |
    Where-Object { $_.Name -ne "README.md" } |
    Select-Object -First 1
$guide = Get-Content -Raw -LiteralPath $guidePath.FullName

if ($installer -notmatch 'function Test-PythonModule') {
    throw "missing Python module check"
}
if ($installer -notmatch 'function Install-PythonPackage') {
    throw "missing per-package install helper"
}
if ($installer -notmatch '(?is)Ghostscript.{0,120}optional') {
    throw "missing optional Ghostscript behavior"
}
if ($installer -match 'install-python-packages\.ps1') {
    throw "installer still depends on duplicate package script"
}
if ($installer -notmatch 'pypdfium2') {
    throw "installer does not check pypdfium2"
}

foreach ($document in @($readme, $guide)) {
    if ($document -match 'install-python\.ps1|install-python-packages\.ps1') {
        throw "documentation still references duplicate Python installer"
    }
}
if ($runtime -match 'install-python\.ps1|install-python-packages\.ps1') {
    throw "runtime still references duplicate Python installer"
}

foreach ($duplicate in @("install-python.ps1", "install-python-packages.ps1")) {
    if (Test-Path -LiteralPath (Join-Path $root $duplicate)) {
        throw "duplicate installer still exists: $duplicate"
    }
}

Write-Host "PASS installer structure checks"
