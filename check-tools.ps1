Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Add-ToolDirectory {
    param([string]$Path)

    if (Test-Path -LiteralPath $Path) {
        $env:Path = "$Path;$env:Path"
    }
}

Add-ToolDirectory (Join-Path $PSScriptRoot "tools\poppler-26.02.0-0\poppler-26.02.0\Library\bin")
Add-ToolDirectory (Join-Path $PSScriptRoot "tools\qpdf-12.3.2\qpdf-12.3.2-msvc64\bin")
Add-ToolDirectory (Join-Path $PSScriptRoot "tools\ghostscript-10.07.1\bin")
Add-ToolDirectory (Join-Path $PSScriptRoot "tools\Python312")
Add-ToolDirectory (Join-Path $PSScriptRoot "tools\Python312\Scripts")
Add-ToolDirectory "D:\GPL\gs10.07.1\bin"
$env:Path = [Environment]::GetEnvironmentVariable("Path", "User") + ";" +
    [Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
    $env:Path
$projectPython = "C:\Users\Administrator\AppData\Local\Python\bin\python.exe"
$projectPythonPackages = Join-Path $PSScriptRoot "tools\python-packages"
if (Test-Path -LiteralPath $projectPythonPackages) {
    if ($env:PYTHONPATH) {
        $env:PYTHONPATH = "$projectPythonPackages;$env:PYTHONPATH"
    }
    else {
        $env:PYTHONPATH = $projectPythonPackages
    }
}
$env:HOME = $PSScriptRoot
$env:USERPROFILE = $PSScriptRoot

$projectPython312 = Join-Path $PSScriptRoot "tools\Python312\python.exe"
$useProjectPython312 = Test-Path -LiteralPath $projectPython312
$pythonCommandForChecks = if ($useProjectPython312) { $projectPython312 } elseif (Test-Path -LiteralPath $projectPython) { $projectPython } else { "python" }
if ($useProjectPython312) {
    $env:PYTHONPATH = ""
}

$checks = @(
    @{ Name = "Python"; Command = $pythonCommandForChecks; Args = @("--version") },
    @{ Name = "pip"; Command = $pythonCommandForChecks; Args = @("-m", "pip", "--version") },
    @{ Name = "Poppler pdftoppm"; Command = "pdftoppm"; Args = @("-v") },
    @{ Name = "qpdf"; Command = "qpdf"; Args = @("--version") },
    @{ Name = "Ghostscript"; Command = "gswin64c"; Args = @("--version") },
    @{ Name = "img2pdf"; Command = $pythonCommandForChecks; Args = @("-m", "img2pdf", "--version") },
    @{ Name = "pikepdf"; Command = $pythonCommandForChecks; Args = @("-c", "import pikepdf; print(pikepdf.__version__)") },
    @{ Name = "pywin32"; Command = $pythonCommandForChecks; Args = @("-c", "import win32com.client; print('available')") }
)

foreach ($check in $checks) {
    $cmd = Get-Command $check.Command -ErrorAction SilentlyContinue
    if (-not $cmd) {
        Write-Host ("[missing] " + $check.Name)
        continue
    }

    if (($check.Command -eq "python" -or $check.Command -eq "py") -and $cmd.Source -like "*\Microsoft\WindowsApps\*") {
        Write-Host ("[missing] " + $check.Name + " (Windows placeholder found, real install not visible)")
        continue
    }

    $oldPreference = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    $output = & $cmd.Source @($check.Args) 2>&1
    $exitCode = $LASTEXITCODE
    $ErrorActionPreference = $oldPreference
    if ($exitCode -ne 0) {
        Write-Host ("[missing] " + $check.Name + " -> " + $cmd.Source)
        $output | Select-Object -First 1 | ForEach-Object { Write-Host ("          " + $_) }
        continue
    }

    Write-Host ("[found]   " + $check.Name + " -> " + $cmd.Source)
    $output | Select-Object -First 1 | ForEach-Object { Write-Host ("          " + $_) }
}
