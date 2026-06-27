param(
    [ValidateNotNullOrEmpty()]
    [string]$InputDir = "",

    [ValidateNotNullOrEmpty()]
    [string]$QueueCsv = "",

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputDir,

    [ValidateSet("pdf", "word", "auto")]
    [string]$InputType = "pdf",

    [ValidateNotNullOrEmpty()]
    [string]$Language = "chi_sim+eng",

    [ValidateRange(72, 600)]
    [int]$Dpi = 300,

    [switch]$Deskew,
    [switch]$RotatePages,
    [switch]$Optimize,
    [switch]$KeepWork,
    [switch]$DetailedPngProgress,
    [switch]$Recurse,
    [switch]$SkipExisting,

    [ValidateSet("pdfium", "poppler")]
    [string]$RenderEngine = "pdfium"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-Directory {
    param(
        [string]$Path,
        [string]$Name
    )

    $resolved = Resolve-Path -LiteralPath $Path -ErrorAction SilentlyContinue
    if (-not $resolved) {
        throw "$Name directory was not found: $Path"
    }

    $item = Get-Item -LiteralPath $resolved.Path
    if (-not $item.PSIsContainer) {
        throw "$Name is not a directory: $Path"
    }

    return $item.FullName
}

function Get-QueueType {
    param(
        [string]$Path,
        [string]$PreferredType
    )

    if ($PreferredType -ne "auto") {
        return $PreferredType
    }

    $extension = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($extension -eq ".pdf") {
        return "pdf"
    }

    if ($extension -eq ".doc" -or $extension -eq ".docx") {
        return "word"
    }

    throw "Unsupported input extension: $Path"
}

function New-OutputPath {
    param(
        [string]$InputPath,
        [string]$TargetDir
    )

    $name = [IO.Path]::GetFileNameWithoutExtension($InputPath)
    return Join-Path $TargetDir ($name + "-dual-layer.pdf")
}

function New-QueueItem {
    param(
        [string]$InputPath,
        [string]$OutputPath,
        [string]$Type
    )

    [pscustomobject]@{
        InputPath = $InputPath
        OutputPath = $OutputPath
        Type = $Type
    }
}

function Join-ProcessArguments {
    param([string[]]$Arguments)

    return ($Arguments | ForEach-Object {
        if ($_ -match '[\s"]') {
            '"' + ($_ -replace '"', '\"') + '"'
        }
        else {
            $_
        }
    }) -join " "
}

function Read-QueueCsv {
    param(
        [string]$Path,
        [string]$TargetDir
    )

    $resolved = Resolve-Path -LiteralPath $Path -ErrorAction SilentlyContinue
    if (-not $resolved) {
        throw "Queue CSV was not found: $Path"
    }

    $rows = Import-Csv -LiteralPath $resolved.Path
    foreach ($row in $rows) {
        if (-not $row.InputPath) {
            throw "Queue CSV row is missing InputPath."
        }

        $inputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($row.InputPath)
        $type = if ($row.Type) { $row.Type } else { "auto" }
        $type = Get-QueueType -Path $inputPath -PreferredType $type
        $outputPath = if ($row.OutputPath) {
            $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($row.OutputPath)
        }
        else {
            New-OutputPath -InputPath $inputPath -TargetDir $TargetDir
        }

        New-QueueItem -InputPath $inputPath -OutputPath $outputPath -Type $type
    }
}

function Read-QueueDirectory {
    param(
        [string]$Path,
        [string]$TargetDir,
        [string]$PreferredType,
        [switch]$Recursive
    )

    $sourceDir = Resolve-Directory -Path $Path -Name "Input"
    $patterns = if ($PreferredType -eq "pdf") {
        @("*.pdf")
    }
    elseif ($PreferredType -eq "word") {
        @("*.doc", "*.docx")
    }
    else {
        @("*.pdf", "*.doc", "*.docx")
    }

    foreach ($pattern in $patterns) {
        Get-ChildItem -LiteralPath $sourceDir -Filter $pattern -File -Recurse:$Recursive |
            Sort-Object FullName |
            ForEach-Object {
                $type = Get-QueueType -Path $_.FullName -PreferredType $PreferredType
                $outputPath = New-OutputPath -InputPath $_.FullName -TargetDir $TargetDir
                New-QueueItem -InputPath $_.FullName -OutputPath $outputPath -Type $type
            }
    }
}

function Invoke-QueueChild {
    param(
        [string[]]$Arguments,
        [string]$LogPath
    )

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "powershell"
    $psi.Arguments = Join-ProcessArguments $Arguments
    $psi.WorkingDirectory = $PSScriptRoot
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $psi.StandardErrorEncoding = [System.Text.Encoding]::UTF8

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi

    $logEncoding = New-Object System.Text.UTF8Encoding($false)
    $writer = [System.IO.StreamWriter]::new($LogPath, $false, $logEncoding)

    try {
        [void]$process.Start()

        while (-not $process.HasExited) {
            while (-not $process.StandardOutput.EndOfStream) {
                $line = $process.StandardOutput.ReadLine()
                $writer.WriteLine($line)
                $writer.Flush()
                Write-Host $line
            }

            while (-not $process.StandardError.EndOfStream) {
                $line = $process.StandardError.ReadLine()
                $writer.WriteLine($line)
                $writer.Flush()
                Write-Host $line
            }

            Start-Sleep -Milliseconds 200
        }

        while (-not $process.StandardOutput.EndOfStream) {
            $line = $process.StandardOutput.ReadLine()
            $writer.WriteLine($line)
            Write-Host $line
        }

        while (-not $process.StandardError.EndOfStream) {
            $line = $process.StandardError.ReadLine()
            $writer.WriteLine($line)
            Write-Host $line
        }

        $writer.Flush()
        return $process.ExitCode
    }
    finally {
        $writer.Dispose()
        if (-not $process.HasExited) {
            $process.Kill()
        }
        $process.Dispose()
    }
}

if (-not $InputDir -and -not $QueueCsv) {
    throw "Provide -InputDir or -QueueCsv."
}

if ($InputDir -and $QueueCsv) {
    throw "Provide either -InputDir or -QueueCsv, not both."
}

$outputRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDir)
New-Item -ItemType Directory -Path $outputRoot -Force | Out-Null

$logDir = Join-Path $outputRoot "_queue-logs"
New-Item -ItemType Directory -Path $logDir -Force | Out-Null
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$summaryPath = Join-Path $logDir ("summary-" + $timestamp + ".csv")
$controlDir = Join-Path $logDir "control"
$pauseFile = Join-Path $controlDir "pause.flag"
$skipFile = Join-Path $controlDir "skip-next.flag"
New-Item -ItemType Directory -Path $controlDir -Force | Out-Null
Remove-Item -LiteralPath $pauseFile, $skipFile -Force -ErrorAction SilentlyContinue

$queue = @(if ($QueueCsv) {
    @(Read-QueueCsv -Path $QueueCsv -TargetDir $outputRoot)
}
else {
    @(Read-QueueDirectory -Path $InputDir -TargetDir $outputRoot -PreferredType $InputType -Recursive:$Recurse)
})

if ($queue.Count -eq 0) {
    throw "Queue is empty."
}

Write-Host ("Queue items: " + $queue.Count)

$results = New-Object System.Collections.Generic.List[object]
$index = 0

foreach ($item in $queue) {
    $index += 1
    $logPath = Join-Path $logDir ("item-{0:000}-{1}.log" -f $index, ([IO.Path]::GetFileNameWithoutExtension($item.InputPath)))
    $status = "failed"
    $message = ""

    try {
        while (Test-Path -LiteralPath $pauseFile) {
            Write-Host ("[{0}/{1}] PAUSED" -f $index, $queue.Count)
            Start-Sleep -Seconds 2
        }

        if (Test-Path -LiteralPath $skipFile) {
            Remove-Item -LiteralPath $skipFile -Force -ErrorAction SilentlyContinue
            $status = "skipped"
            $message = "Skipped by user."
            Write-Host ("[{0}/{1}] SKIP {2}" -f $index, $queue.Count, $item.InputPath)
        }
        elseif ($SkipExisting -and (Test-Path -LiteralPath $item.OutputPath)) {
            $status = "skipped"
            $message = "Output already exists."
            Write-Host ("[{0}/{1}] SKIP {2}" -f $index, $queue.Count, $item.InputPath)
        }
        else {
            Write-Host ("[{0}/{1}] START {2}" -f $index, $queue.Count, $item.InputPath)

            $arguments = @(
                "-ExecutionPolicy", "Bypass",
                "-File", (Join-Path $PSScriptRoot "make-dual-pdf.ps1"),
                "-OutputPath", $item.OutputPath,
                "-Language", $Language,
                "-Dpi", $Dpi,
                "-RenderEngine", $RenderEngine
            )

            if ($item.Type -eq "word") {
                $arguments += @("-WordPath", $item.InputPath)
            }
            else {
                $arguments += @("-TaggedPdfPath", $item.InputPath)
            }

            if ($Deskew) {
                $arguments += "-Deskew"
            }

            if ($RotatePages) {
                $arguments += "-RotatePages"
            }

            if ($Optimize) {
                $arguments += "-Optimize"
            }

            if ($KeepWork) {
                $arguments += "-KeepWork"
            }

            if ($DetailedPngProgress) {
                $arguments += "-DetailedPngProgress"
            }

            $childExitCode = Invoke-QueueChild -Arguments $arguments -LogPath $logPath

            if ($childExitCode -ne 0) {
                throw "make-dual-pdf.ps1 failed with exit code $childExitCode. See $logPath"
            }

            $status = "success"
            $message = "Created."
            Write-Host ("[{0}/{1}] OK {2}" -f $index, $queue.Count, $item.OutputPath)
        }
    }
    catch {
        $message = $_.Exception.Message
        if (-not (Test-Path -LiteralPath $logPath)) {
            $message | Set-Content -LiteralPath $logPath -Encoding UTF8
        }
        Write-Host ("[{0}/{1}] FAIL {2}" -f $index, $queue.Count, $item.InputPath)
        Write-Host ("          " + $message)
    }

    $results.Add([pscustomobject]@{
        Index = $index
        Status = $status
        Type = $item.Type
        InputPath = $item.InputPath
        OutputPath = $item.OutputPath
        LogPath = $logPath
        Message = $message
    })
}

$results | Export-Csv -LiteralPath $summaryPath -NoTypeInformation -Encoding UTF8

$success = @($results | Where-Object { $_.Status -eq "success" }).Count
$failed = @($results | Where-Object { $_.Status -eq "failed" }).Count
$skipped = @($results | Where-Object { $_.Status -eq "skipped" }).Count

Write-Host "Queue complete."
Write-Host "Success: $success"
Write-Host "Skipped: $skipped"
Write-Host "Failed: $failed"
Write-Host "Summary: $summaryPath"

if ($failed -gt 0) {
    exit 1
}
