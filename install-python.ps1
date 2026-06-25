Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$version = "3.13.14"
$installerName = "python-$version-amd64.exe"
$installerUrl = "https://www.python.org/ftp/python/$version/$installerName"
$downloadPath = Join-Path $PSScriptRoot $installerName

function Test-RealPython {
    $command = Get-Command python -ErrorAction SilentlyContinue
    if (-not $command) {
        return $false
    }

    return $command.Source -notlike "*\Microsoft\WindowsApps\python.exe"
}

if (Test-RealPython) {
    python --version
}
else {
    Write-Host "Downloading Python $version..."
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $installerUrl -OutFile $downloadPath -UseBasicParsing

    Write-Host "Installing Python $version for current user..."
    Start-Process -FilePath $downloadPath -ArgumentList @(
        "/quiet",
        "InstallAllUsers=0",
        "PrependPath=1",
        "Include_pip=1",
        "Include_launcher=1"
    ) -Wait
}

Write-Host "Installing Python package: pikepdf"
python -m pip install --upgrade pip
python -m pip install -r (Join-Path $PSScriptRoot "requirements.txt")

Write-Host "Python setup complete."
python --version

