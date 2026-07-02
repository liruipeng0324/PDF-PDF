param(
    [int]$TimeoutSeconds = 20
)

$ErrorActionPreference = "Continue"

function Stop-AcrobatProcesses {
    Get-Process Acrobat, AcroCEF, AcroRd32 -ErrorAction SilentlyContinue | ForEach-Object {
        Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
    }
}

function Test-ComProcess {
    param(
        [string]$Name,
        [string]$Code,
        [int]$Timeout
    )

    $tempScript = Join-Path $env:TEMP ("acrobat-com-check-" + [guid]::NewGuid().ToString("N") + ".ps1")
    $tempOut = Join-Path $env:TEMP ("acrobat-com-check-" + [guid]::NewGuid().ToString("N") + ".out")
    $tempErr = Join-Path $env:TEMP ("acrobat-com-check-" + [guid]::NewGuid().ToString("N") + ".err")

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

        if (-not $process.WaitForExit($Timeout * 1000)) {
            Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
            Write-Host "$Name TIMEOUT"
            Write-Host ($process.StandardOutput.ReadToEnd())
            Write-Host ($process.StandardError.ReadToEnd())
            return $false
        }

        Write-Host ($process.StandardOutput.ReadToEnd())
        Write-Host ($process.StandardError.ReadToEnd())
        if ($process.ExitCode -ne 0) {
            Write-Host "$Name FAILED with exit code $($process.ExitCode)"
            return $false
        }

        Write-Host "$Name OK"
        return $true
    }
    finally {
        Remove-Item -LiteralPath $tempScript -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $tempOut -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $tempErr -Force -ErrorAction SilentlyContinue
    }
}

Stop-AcrobatProcesses
Start-Sleep -Seconds 1

$appOk = Test-ComProcess -Name "AcroExch.App" -Timeout $TimeoutSeconds -Code @'
$app = $null
try {
    Write-Output "before AcroExch.App"
    $app = New-Object -ComObject AcroExch.App
    Write-Output "after AcroExch.App"
    $app.Hide()
    Write-Output "after Hide"
    exit 0
}
finally {
    if ($app -ne $null) {
        try { $app.Exit() } catch {}
    }
}
'@

$pddocOk = Test-ComProcess -Name "AcroExch.PDDoc" -Timeout $TimeoutSeconds -Code @'
$pddoc = $null
try {
    Write-Output "before AcroExch.PDDoc"
    $pddoc = New-Object -ComObject AcroExch.PDDoc
    Write-Output "after AcroExch.PDDoc"
    exit 0
}
finally {
    if ($pddoc -ne $null) {
        try { $pddoc.Close() } catch {}
    }
}
'@

$python = Join-Path $PSScriptRoot "tools\Python312\python.exe"
if (-not (Test-Path -LiteralPath $python)) {
    $python = "python"
}

$jsOk = Test-ComProcess -Name "PDDoc.GetJSObject low-level eval" -Timeout $TimeoutSeconds -Code @"
`$ErrorActionPreference = "Stop"
& "$python" -u -c "import pathlib, pikepdf, pythoncom, win32com.client; p=pathlib.Path(r'$PSScriptRoot')/'tmp-com-check.pdf'; pdf=pikepdf.Pdf.new(); pdf.add_blank_page(page_size=(595,842)); pdf.save(p); pdf.close(); d=win32com.client.Dispatch('AcroExch.PDDoc'); print('before Open', flush=True); print('Open', d.Open(str(p)), flush=True); js=d.GetJSObject(); dispid=js._oleobj_.GetIDsOfNames('eval'); result=js._oleobj_.Invoke(dispid,0,pythoncom.DISPATCH_METHOD,True,'1+1'); print('eval', result, flush=True); d.Close(); p.unlink(missing_ok=True)"
"@

if ($appOk -and $pddocOk -and $jsOk) {
    Stop-AcrobatProcesses
    Write-Host "ACROBAT_COM_OK"
    exit 0
}

Stop-AcrobatProcesses
Write-Host "ACROBAT_COM_FAILED"
exit 1
