Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$python = Join-Path $PSScriptRoot "tools\Python313\python.exe"
if (-not (Test-Path -LiteralPath $python)) {
    $localPython = "C:\Users\Administrator\AppData\Local\Python\bin\python.exe"
    if (Test-Path -LiteralPath $localPython) {
        $python = $localPython
    }
    else {
        $pythonCommand = Get-Command python -ErrorAction SilentlyContinue
        if (-not $pythonCommand -or $pythonCommand.Source -like "*\Microsoft\WindowsApps\*") {
            throw "Python was not found. Run setup-dependencies.ps1 first."
        }
        $python = $pythonCommand.Source
    }
}

$url = "http://127.0.0.1:8765/"
$serverScript = Join-Path $PSScriptRoot "web-server.py"

$existing = Get-NetTCPConnection -LocalPort 8765 -ErrorAction SilentlyContinue
if (-not $existing) {
    Start-Process -FilePath $python -ArgumentList @($serverScript, "--host", "127.0.0.1", "--port", "8765") -WindowStyle Hidden
    Start-Sleep -Seconds 2
}

Start-Process $url
Write-Host $url
