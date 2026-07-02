param(
    [string]$AcrobatExe = "D:\Acrobat\Acrobat\Acrobat.exe",
    [int]$TimeoutSeconds = 20
)

$ErrorActionPreference = "Continue"

function Write-Step {
    param([string]$Message)
    Write-Host ""
    Write-Host "== $Message =="
}

function Remove-UserComOverride {
    $clsids = @(
        "{85DE1C45-2C66-101B-B02E-04021C009402}",
        "{FF76CB60-2E68-101B-B02E-04021C009402}"
    )

    foreach ($clsid in $clsids) {
        $key = "HKCU:\Software\Classes\CLSID\$clsid"
        if (Test-Path -LiteralPath $key) {
            Remove-Item -LiteralPath $key -Recurse -Force
            Write-Host "Removed user COM override: $key"
        }
        else {
            Write-Host "No user COM override: $key"
        }
    }
}

function Stop-AcrobatProcesses {
    Get-Process Acrobat, AcroCEF, AcroRd32 -ErrorAction SilentlyContinue | ForEach-Object {
        Write-Host "Stopping $($_.ProcessName) PID $($_.Id)"
        Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
    }
}

function Show-ComRegistration {
    $items = @(
        "Registry::HKEY_CLASSES_ROOT\AcroExch.App\CLSID",
        "Registry::HKEY_CLASSES_ROOT\AcroExch.PDDoc\CLSID",
        "Registry::HKEY_CLASSES_ROOT\CLSID\{85DE1C45-2C66-101B-B02E-04021C009402}\LocalServer32",
        "Registry::HKEY_CLASSES_ROOT\CLSID\{FF76CB60-2E68-101B-B02E-04021C009402}\LocalServer32"
    )

    foreach ($item in $items) {
        Write-Host $item
        if (Test-Path -LiteralPath $item) {
            Get-ItemProperty -LiteralPath $item | Format-List
        }
        else {
            Write-Host "MISSING"
        }
    }
}

function Set-AcrobatAutomationPreferences {
    $settings = @(
        @{ Path = "HKCU:\Software\Adobe\Adobe Acrobat\DC\Privileged"; Name = "bProtectedMode"; Value = 0 },
        @{ Path = "HKCU:\Software\Adobe\Adobe Acrobat\DC\TrustManager"; Name = "bEnhancedSecurityStandalone"; Value = 0 },
        @{ Path = "HKCU:\Software\Adobe\Adobe Acrobat\DC\TrustManager"; Name = "bEnhancedSecurityInBrowser"; Value = 0 },
        @{ Path = "HKCU:\Software\Adobe\Adobe Acrobat\DC\FeatureLockDown"; Name = "bProtectedMode"; Value = 0 },
        @{ Path = "HKCU:\Software\Adobe\Adobe Acrobat\DC\FeatureLockDown"; Name = "bEnhancedSecurityStandalone"; Value = 0 },
        @{ Path = "HKCU:\Software\Adobe\Adobe Acrobat\DC\FeatureLockDown"; Name = "bToggleAdobeDocumentServices"; Value = 1 },
        @{ Path = "HKCU:\Software\Adobe\Adobe Acrobat\DC\FeatureLockDown"; Name = "bToggleAdobeSign"; Value = 1 },
        @{ Path = "HKCU:\Software\Adobe\Adobe Acrobat\DC\FeatureLockDown"; Name = "bUpdater"; Value = 0 },
        @{ Path = "HKCU:\Software\Adobe\Adobe Acrobat\DC\FeatureLockDown"; Name = "bDisableJavaScript"; Value = 0 },
        @{ Path = "HKCU:\Software\Adobe\Adobe Acrobat\DC\FeatureLockDown\cServices"; Name = "bToggleAdobeDocumentServices"; Value = 1 },
        @{ Path = "HKCU:\Software\Adobe\Adobe Acrobat\DC\FeatureLockDown\cServices"; Name = "bToggleAdobeSign"; Value = 1 },
        @{ Path = "HKCU:\Software\Adobe\Adobe Acrobat\DC\FeatureLockDown\cServices"; Name = "bUpdater"; Value = 0 },
        @{ Path = "HKCU:\Software\Adobe\Adobe Acrobat\DC\JSPrefs"; Name = "bEnableJS"; Value = 1 },
        @{ Path = "HKCU:\Software\Adobe\Adobe Acrobat\DC\JSPrefs"; Name = "bEnableMenuItems"; Value = 1 }
    )

    foreach ($setting in $settings) {
        New-Item -Path $setting.Path -Force | Out-Null
        New-ItemProperty -Path $setting.Path -Name $setting.Name -Value $setting.Value -PropertyType DWord -Force | Out-Null
        Write-Host "Set $($setting.Path)\$($setting.Name) = $($setting.Value)"
    }
}

function Invoke-AcrobatRegistration {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Acrobat.exe not found: $Path"
    }

    Write-Host "Running Acrobat registration command..."
    $process = Start-Process -FilePath $Path -ArgumentList "/regserver" -PassThru -WindowStyle Hidden
    if (-not $process.WaitForExit(15000)) {
        Write-Host "Registration command is still running; stopping it."
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
    }
}

function Test-ComObject {
    param(
        [string]$Name,
        [string]$Code,
        [int]$Timeout
    )

    Write-Host "Testing $Name..."
    $tempScript = Join-Path $env:TEMP ("acrobat-com-test-" + [guid]::NewGuid().ToString("N") + ".ps1")
    $tempOut = Join-Path $env:TEMP ("acrobat-com-test-" + [guid]::NewGuid().ToString("N") + ".log")
    $tempErr = Join-Path $env:TEMP ("acrobat-com-test-" + [guid]::NewGuid().ToString("N") + ".err")

    try {
        Set-Content -LiteralPath $tempScript -Value $Code -Encoding UTF8
        $powerShellExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $powerShellExe
        $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$tempScript`""
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $process = [System.Diagnostics.Process]::Start($psi)

        if ($process.WaitForExit($Timeout * 1000)) {
            Write-Host ($process.StandardOutput.ReadToEnd())
            Write-Host ($process.StandardError.ReadToEnd())
            return ($process.ExitCode -eq 0)
        }

        Write-Host "$Name TIMEOUT after $Timeout seconds"
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
        Write-Host ($process.StandardOutput.ReadToEnd())
        Write-Host ($process.StandardError.ReadToEnd())
        return $false
    }
    finally {
        Remove-Item -LiteralPath $tempScript -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $tempOut -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $tempErr -Force -ErrorAction SilentlyContinue
    }
}

Write-Step "Close Acrobat"
Stop-AcrobatProcesses
Start-Sleep -Seconds 2

Write-Step "Remove current-user COM overrides"
Remove-UserComOverride

Write-Step "Set Acrobat automation preferences"
Set-AcrobatAutomationPreferences

Write-Step "Show registration before repair"
Show-ComRegistration

Write-Step "Run Acrobat /regserver"
Invoke-AcrobatRegistration -Path $AcrobatExe
Start-Sleep -Seconds 2

Write-Step "Close Acrobat after registration"
Stop-AcrobatProcesses
Start-Sleep -Seconds 2

Write-Step "Show registration after repair"
Show-ComRegistration

Write-Step "Self-check COM"
$appOk = Test-ComObject -Name "AcroExch.App" -Timeout $TimeoutSeconds -Code @'
Write-Output "before AcroExch.App"
$app = New-Object -ComObject AcroExch.App
Write-Output "after AcroExch.App"
$app.Hide()
Write-Output "after Hide"
$app.Exit()
'@

$pddocOk = Test-ComObject -Name "AcroExch.PDDoc" -Timeout $TimeoutSeconds -Code @'
Write-Output "before AcroExch.PDDoc"
$pddoc = New-Object -ComObject AcroExch.PDDoc
Write-Output "after AcroExch.PDDoc"
$pddoc.Close()
'@

Write-Step "Result"
if ($appOk -and $pddocOk) {
    Write-Host "ACROBAT_COM_OK"
    exit 0
}

Write-Host "ACROBAT_COM_FAILED"
Write-Host "If this still fails, run Acrobat once as Administrator, finish login/activation/update prompts, then run this script again."
exit 1
