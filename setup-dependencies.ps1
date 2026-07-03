Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Add-ToolDirectory {
    param([string]$Path)

    if (Test-Path -LiteralPath $Path) {
        $env:Path = "$Path;$env:Path"
    }
}

function Add-MatchingToolDirectories {
    param(
        [string]$Root,
        [string]$Pattern
    )

    if (-not (Test-Path -LiteralPath $Root)) {
        return
    }

    Get-ChildItem -LiteralPath $Root -Recurse -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -like $Pattern } |
        ForEach-Object { Add-ToolDirectory $_.FullName }
}

function Find-Python {
    $projectPython312 = Join-Path $PSScriptRoot "tools\Python312\python.exe"
    if (Test-Path -LiteralPath $projectPython312) {
        return $projectPython312
    }

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
Add-ToolDirectory (Join-Path $PSScriptRoot "tools\Python312")
Add-ToolDirectory (Join-Path $PSScriptRoot "tools\Python312\Scripts")
Add-ToolDirectory (Join-Path $PSScriptRoot "tools\poppler-26.02.0-0\poppler-26.02.0\Library\bin")
Add-ToolDirectory (Join-Path $PSScriptRoot "tools\qpdf-12.3.2\qpdf-12.3.2-msvc64\bin")
Add-ToolDirectory (Join-Path $PSScriptRoot "tools\ghostscript-10.07.1\bin")
Add-MatchingToolDirectories -Root (Join-Path $PSScriptRoot "tools") -Pattern "*\poppler-*\Library\bin"
Add-MatchingToolDirectories -Root (Join-Path $PSScriptRoot "tools") -Pattern "*\qpdf-*\bin"
Add-MatchingToolDirectories -Root (Join-Path $PSScriptRoot "tools") -Pattern "*\gs*\bin"
Add-MatchingToolDirectories -Root (Join-Path $PSScriptRoot "tools") -Pattern "*\ghostscript*\bin"

$packages = Join-Path $PSScriptRoot "tools\python-packages"
if (Test-Path -LiteralPath $packages) {
    $env:PYTHONPATH = "$packages;$env:PYTHONPATH"
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
