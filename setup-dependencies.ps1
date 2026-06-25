Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Add-ToolDirectory {
    param([string]$Path)

    if (Test-Path -LiteralPath $Path) {
        $env:Path = "$Path;$env:Path"
    }
}

function Find-Python {
    $projectPython = Join-Path $PSScriptRoot "tools\Python313\python.exe"
    if (Test-Path -LiteralPath $projectPython) {
        return $projectPython
    }

    $localPython = "C:\Users\Administrator\AppData\Local\Python\bin\python.exe"
    if (Test-Path -LiteralPath $localPython) {
        return $localPython
    }

    $python = Get-Command "python" -ErrorAction SilentlyContinue
    if ($python -and $python.Source -notlike "*\Microsoft\WindowsApps\*") {
        return $python.Source
    }

    return ""
}

Add-ToolDirectory (Join-Path $PSScriptRoot "tools\Python313")
Add-ToolDirectory (Join-Path $PSScriptRoot "tools\Python313\Scripts")
Add-ToolDirectory (Join-Path $PSScriptRoot "tools\poppler-26.02.0-0\poppler-26.02.0\Library\bin")
Add-ToolDirectory (Join-Path $PSScriptRoot "tools\qpdf-12.3.2\qpdf-12.3.2-msvc64\bin")
Add-ToolDirectory (Join-Path $PSScriptRoot "tools\tesseract")
Add-ToolDirectory (Join-Path $PSScriptRoot "tools\tesseract-nsis")
Add-ToolDirectory (Join-Path $PSScriptRoot "tools\ghostscript-10.07.1\bin")

$packages = Join-Path $PSScriptRoot "tools\python-packages"
if (Test-Path -LiteralPath $packages) {
    $env:PYTHONPATH = "$packages;$env:PYTHONPATH"
}

$tessdata = Join-Path $PSScriptRoot "tools\tessdata"
if (Test-Path -LiteralPath $tessdata) {
    $env:TESSDATA_PREFIX = $tessdata
}

$python = Find-Python
Write-Host "PDF-PDF dependency check"
Write-Host "Project: $PSScriptRoot"
Write-Host ""

if ($python) {
    Write-Host "[found] Python: $python"
}
else {
    Write-Host "[missing] Python"
    Write-Host "          Install Python, or place a portable Python at tools\Python313\python.exe."
}

& powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "check-tools.ps1")

Write-Host ""
Write-Host "If all required items show [found], double-click start-web.bat."
Write-Host "If Microsoft Word is missing, use the already-exported PDF mode."
