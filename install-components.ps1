param(
    [switch]$Force,
    [switch]$SkipDownloads,
    [switch]$SkipGhostscript
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$toolsDir = Join-Path $PSScriptRoot "tools"
$downloadDir = Join-Path $PSScriptRoot ".cache\downloads"
$pythonDir = Join-Path $toolsDir "Python312"

$pythonVersion = "3.12.10"
$pythonInstallerUrl = "https://www.python.org/ftp/python/$pythonVersion/python-$pythonVersion-amd64.exe"

function Write-Step {
    param([string]$Message)
    Write-Host ""
    Write-Host ("== " + $Message)
}

function Ensure-Directory {
    param([string]$Path)
    [System.IO.Directory]::CreateDirectory($Path) | Out-Null
}

function Download-File {
    param(
        [string]$Url,
        [string]$TargetPath
    )

    if ((Test-Path -LiteralPath $TargetPath) -and -not $Force) {
        Write-Host ("Using cached download: " + $TargetPath)
        return
    }

    Write-Host ("Downloading: " + $Url)
    Invoke-WebRequest -Uri $Url -OutFile $TargetPath -UseBasicParsing
}

function Get-GitHubReleaseAssetUrl {
    param(
        [string]$Repository,
        [string]$AssetLike,
        [string]$Tag = ""
    )

    $apiUrl = if ($Tag) {
        "https://api.github.com/repos/$Repository/releases/tags/$Tag"
    }
    else {
        "https://api.github.com/repos/$Repository/releases/latest"
    }

    Write-Host ("GitHub release lookup: " + $apiUrl)
    $release = Invoke-RestMethod -Uri $apiUrl -Headers @{ "User-Agent" = "PDF-PDF-installer" }
    $asset = @($release.assets | Where-Object { $_.name -like $AssetLike } | Select-Object -First 1)
    if (-not $asset) {
        throw "Could not find GitHub release asset '$AssetLike' in $Repository."
    }
    return $asset.browser_download_url
}

function Expand-ToolZip {
    param(
        [string]$ZipPath,
        [string]$Destination
    )

    Ensure-Directory $Destination
    Write-Host ("Extracting: " + $ZipPath)
    Expand-Archive -LiteralPath $ZipPath -DestinationPath $Destination -Force
}

function Install-PortablePython {
    $pythonExe = Join-Path $pythonDir "python.exe"
    if ((Test-Path -LiteralPath $pythonExe) -and -not $Force) {
        Write-Host ("Python already exists: " + $pythonExe)
        return $pythonExe
    }

    if ($SkipDownloads) {
        throw "Python is missing and -SkipDownloads was used."
    }

    Ensure-Directory $downloadDir
    Ensure-Directory $toolsDir
    $installer = Join-Path $downloadDir ("python-$pythonVersion-amd64.exe")
    Download-File -Url $pythonInstallerUrl -TargetPath $installer

    Write-Host ("Installing Python to: " + $pythonDir)
    $arguments = @(
        "/quiet",
        "InstallAllUsers=0",
        "TargetDir=$pythonDir",
        "Include_pip=1",
        "PrependPath=0",
        "Include_test=0",
        "Include_doc=0",
        "Shortcuts=0"
    )
    Start-Process -FilePath $installer -ArgumentList $arguments -Wait

    if (-not (Test-Path -LiteralPath $pythonExe)) {
        throw "Python installation failed: $pythonExe was not created."
    }

    return $pythonExe
}

function Install-PythonPackages {
    param([string]$PythonExe)

    Write-Step "Install Python packages"
    & $PythonExe --version
    & $PythonExe -m pip install --upgrade pip
    & $PythonExe -m pip install -r (Join-Path $PSScriptRoot "requirements.txt")
    & $PythonExe -m pip install img2pdf
}

function Install-Qpdf {
    if ((Get-ChildItem -LiteralPath $toolsDir -Recurse -File -Filter "qpdf.exe" -ErrorAction SilentlyContinue | Select-Object -First 1) -and -not $Force) {
        Write-Host "qpdf already exists."
        return
    }
    if ($SkipDownloads) {
        Write-Host "qpdf missing, but -SkipDownloads was used."
        return
    }

    Ensure-Directory $downloadDir
    $url = Get-GitHubReleaseAssetUrl -Repository "qpdf/qpdf" -AssetLike "qpdf-*-msvc64.zip"
    $zip = Join-Path $downloadDir ([IO.Path]::GetFileName(([Uri]$url).AbsolutePath))
    Download-File -Url $url -TargetPath $zip
    Expand-ToolZip -ZipPath $zip -Destination $toolsDir
}

function Install-Poppler {
    if ((Get-ChildItem -LiteralPath $toolsDir -Recurse -File -Filter "pdftoppm.exe" -ErrorAction SilentlyContinue | Select-Object -First 1) -and -not $Force) {
        Write-Host "Poppler already exists."
        return
    }
    if ($SkipDownloads) {
        Write-Host "Poppler missing, but -SkipDownloads was used."
        return
    }

    Ensure-Directory $downloadDir
    $url = Get-GitHubReleaseAssetUrl -Repository "oschwartz10612/poppler-windows" -AssetLike "Release-*.zip"
    $zip = Join-Path $downloadDir ([IO.Path]::GetFileName(([Uri]$url).AbsolutePath))
    Download-File -Url $url -TargetPath $zip
    Expand-ToolZip -ZipPath $zip -Destination $toolsDir
}

function Install-Ghostscript {
    if ($SkipGhostscript) {
        Write-Host "Ghostscript installation skipped."
        return
    }
    if ((Get-ChildItem -LiteralPath $toolsDir -Recurse -File -Filter "gswin64c.exe" -ErrorAction SilentlyContinue | Select-Object -First 1) -and -not $Force) {
        Write-Host "Ghostscript already exists."
        return
    }
    if ($SkipDownloads) {
        Write-Host "Ghostscript missing, but -SkipDownloads was used."
        return
    }

    Ensure-Directory $downloadDir
    $url = Get-GitHubReleaseAssetUrl -Repository "ArtifexSoftware/ghostpdl-downloads" -AssetLike "gs*w64.exe"
    $installer = Join-Path $downloadDir ([IO.Path]::GetFileName(([Uri]$url).AbsolutePath))
    $installDir = Join-Path $toolsDir "ghostscript"
    Download-File -Url $url -TargetPath $installer

    Write-Host ("Installing Ghostscript to: " + $installDir)
    Ensure-Directory $installDir
    Start-Process -FilePath $installer -ArgumentList @("/S", "/D=$installDir") -Wait
}

function Test-DesktopApplications {
    Write-Step "Check desktop applications"

    try {
        $word = New-Object -ComObject Word.Application
        $version = $word.Version
        $word.Quit()
        Write-Host ("[found] Microsoft Word COM version " + $version)
    }
    catch {
        Write-Host "[missing] Microsoft Word COM. Word input mode will not work until Word is installed and activated."
    }

    $acrobatCandidates = @(
        "D:\Acrobat\Acrobat\Acrobat.exe",
        "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
        "C:\Program Files\Adobe\Acrobat\Acrobat\Acrobat.exe",
        "C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
    )
    $acrobat = @($acrobatCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1)
    if ($acrobat) {
        Write-Host ("[found] Adobe Acrobat executable: " + $acrobat)
    }
    else {
        Write-Host "[missing] Adobe Acrobat Pro. Install Acrobat Pro before using Acrobat OCR mode."
    }
}

Ensure-Directory $toolsDir
Ensure-Directory $downloadDir

Write-Host "PDF-PDF one-click component installer"
Write-Host ("Project: " + $PSScriptRoot)

Write-Step "Install Python"
$python = Install-PortablePython

if (-not $SkipDownloads) {
    Write-Step "Install PDF command line tools"
    Install-Poppler
    Install-Qpdf
    Install-Ghostscript
}

Install-PythonPackages -PythonExe $python
Test-DesktopApplications

Write-Step "Final dependency check"
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "setup-dependencies.ps1")

Write-Host ""
Write-Host "Done. If Acrobat Pro and Word are available, run start-web.bat."
