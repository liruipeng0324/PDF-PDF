param(
    [ValidateSet("manual-ocr", "full", "all")]
    [string]$Profile = "manual-ocr",
    [switch]$Offline
)
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$manifestPath = Join-Path $PSScriptRoot "parts-manifest.json"
$baseUrl = "https://github.com/liruipeng0324/PDF-PDF/releases/download/v2026.10.13-offline/"
if (-not (Test-Path -LiteralPath $manifestPath)) {
    if ($Offline) { throw "parts-manifest.json is required for offline merge." }
    Write-Host "Downloading package manifest..."
    Invoke-WebRequest -UseBasicParsing -Uri ($baseUrl + "parts-manifest.json") -OutFile $manifestPath -TimeoutSec 180
}
$manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ($manifest.build -ne "20261013-offline-gs" -or $manifest.baseUrl -ne $baseUrl) { throw "Unexpected package manifest." }
$partsDir = Join-Path $PSScriptRoot "parts"
$outputDir = Join-Path $PSScriptRoot "packages"
[IO.Directory]::CreateDirectory($partsDir) | Out-Null
[IO.Directory]::CreateDirectory($outputDir) | Out-Null
$logPath = Join-Path $PSScriptRoot ("Get-Packages-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".log")
Start-Transcript -LiteralPath $logPath | Out-Null
try {
    foreach ($package in $manifest.packages) {
        if ($Profile -ne "all" -and $package.profile -ne $Profile) { continue }
        if ($package.file -notmatch '^PDF-PDF-(full|manual-ocr)\.zip$') { throw "Invalid output filename." }
        $output = Join-Path $outputDir $package.file
        if (Test-Path -LiteralPath $output) {
            if ((Get-FileHash -LiteralPath $output -Algorithm SHA256).Hash -ieq $package.sha256) {
                Write-Host ("Already verified: " + $output)
                continue
            }
            throw ("Existing ZIP has another checksum; move it aside before retrying: " + $output)
        }
        $index = 0
        foreach ($part in $package.parts) {
            $index++
            if ($part.file -notmatch '^PDF-PDF-(full|manual-ocr)\.zip\.part[0-9]{3}$') { throw "Invalid part filename." }
            $local = Join-Path $partsDir $part.file
            $valid = (Test-Path -LiteralPath $local) -and
                (Get-Item -LiteralPath $local).Length -eq $part.size -and
                (Get-FileHash -LiteralPath $local -Algorithm SHA256).Hash -ieq $part.sha256
            if (-not $valid) {
                if ($Offline) { throw ("Missing or invalid offline part: " + $part.file) }
                Write-Progress -Activity ("Downloading " + $package.file) -Status ("Part " + $index + "/" + $package.parts.Count) -PercentComplete ([int](100 * ($index - 1) / $package.parts.Count))
                Write-Host ("Download " + $part.file + " [" + $index + "/" + $package.parts.Count + "]")
                $temporary = $local + ".download"
                Invoke-WebRequest -UseBasicParsing -Uri ($baseUrl + $part.file) -OutFile $temporary -TimeoutSec 180
                if ((Get-Item -LiteralPath $temporary).Length -ne $part.size -or
                    (Get-FileHash -LiteralPath $temporary -Algorithm SHA256).Hash -ine $part.sha256) {
                    throw ("Downloaded part checksum failed: " + $part.file)
                }
                Move-Item -LiteralPath $temporary -Destination $local -Force
            }
        }
        Write-Host ("Merging " + $package.file)
        $partial = $output + ".partial"
        $destination = [IO.File]::Open($partial, [IO.FileMode]::Create, [IO.FileAccess]::Write)
        try {
            foreach ($part in $package.parts) {
                $source = [IO.File]::OpenRead((Join-Path $partsDir $part.file))
                try { $source.CopyTo($destination) } finally { $source.Dispose() }
            }
        } finally { $destination.Dispose() }
        if ((Get-Item -LiteralPath $partial).Length -ne $package.size -or
            (Get-FileHash -LiteralPath $partial -Algorithm SHA256).Hash -ine $package.sha256) {
            throw ("Merged ZIP checksum failed: " + $package.file)
        }
        Move-Item -LiteralPath $partial -Destination $output
        Write-Progress -Activity ("Downloading " + $package.file) -Completed
        Write-Host ("Verified ZIP ready: " + $output)
    }
}
finally { Stop-Transcript | Out-Null }
