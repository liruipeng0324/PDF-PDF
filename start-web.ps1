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
$logPath = Join-Path $PSScriptRoot "web-server.log"
$errorLogPath = Join-Path $PSScriptRoot "web-server-error.log"

$existing = Get-NetTCPConnection -LocalPort 8765 -ErrorAction SilentlyContinue
if (-not $existing) {
    if (Test-Path -LiteralPath $logPath) {
        Remove-Item -LiteralPath $logPath -Force
    }
    if (Test-Path -LiteralPath $errorLogPath) {
        Remove-Item -LiteralPath $errorLogPath -Force
    }
    Start-Process -FilePath $python -ArgumentList @($serverScript, "--host", "127.0.0.1", "--port", "8765") -RedirectStandardOutput $logPath -RedirectStandardError $errorLogPath -WindowStyle Hidden

    $ready = $false
    foreach ($i in 1..20) {
        try {
            Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 1 | Out-Null
            $ready = $true
            break
        }
        catch {
            Start-Sleep -Seconds 1
        }
    }

    if (-not $ready) {
        Write-Host "Web service did not start. Log: $logPath"
        if (Test-Path -LiteralPath $logPath) {
            Get-Content -LiteralPath $logPath
        }
        if (Test-Path -LiteralPath $errorLogPath) {
            Get-Content -LiteralPath $errorLogPath
        }
        exit 1
    }
}

Start-Process $url
Write-Host $url
