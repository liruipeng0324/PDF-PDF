param(
    [string]$Python = "python"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$pythonCommand = Get-Command $Python -ErrorAction SilentlyContinue
if (-not $pythonCommand) {
    throw "Python was not found. Pass the full path, for example: .\install-python-packages.ps1 -Python C:\Path\To\python.exe"
}

if ($pythonCommand.Source -like "*\Microsoft\WindowsApps\*") {
    throw "Only the Windows Python placeholder was found. Pass the real python.exe path with -Python."
}

& $pythonCommand.Source --version
& $pythonCommand.Source -m pip install --upgrade pip
& $pythonCommand.Source -m pip install -r (Join-Path $PSScriptRoot "requirements.txt")
& $pythonCommand.Source -m pip install img2pdf

Write-Host "Python packages installed."
