param(
    [ValidateNotNullOrEmpty()]
    [string]$WordPath = "",

    [ValidateNotNullOrEmpty()]
    [string]$TaggedPdfPath = "",

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath,

    [ValidateNotNullOrEmpty()]
    [string]$Language = "chi_sim+eng",

    [ValidateRange(72, 600)]
    [int]$Dpi = 300,

    [string]$WorkDir = "",

    [switch]$Deskew,
    [switch]$RotatePages,
    [switch]$Optimize,
    [switch]$KeepWork,
    [switch]$DetailedPngProgress,

    [ValidateSet("pdfium", "poppler")]
    [string]$RenderEngine = "pdfium",

    [ValidateSet("jpeg")]
    [string]$ImageFormat = "jpeg",

    [ValidateRange(1, 100)]
    [int]$JpegQuality = 50,

    [ValidateRange(1, 16)]
    [int]$RenderWorkers = 4
)

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
Add-ToolDirectory (Join-Path $PSScriptRoot "tools\tesseract")
Add-ToolDirectory (Join-Path $PSScriptRoot "tools\tesseract-nsis")
Add-ToolDirectory (Join-Path $PSScriptRoot "tools\ghostscript-10.07.1\bin")
Add-ToolDirectory "D:\OCR"
Add-ToolDirectory "D:\GPL\gs10.07.1\bin"
$env:Path = [Environment]::GetEnvironmentVariable("Path", "User") + ";" +
    [Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
    $env:Path
$projectPythonPackages = Join-Path $PSScriptRoot "tools\python-packages"
if (Test-Path -LiteralPath $projectPythonPackages) {
    if ($env:PYTHONPATH) {
        $env:PYTHONPATH = "$projectPythonPackages;$env:PYTHONPATH"
    }
    else {
        $env:PYTHONPATH = $projectPythonPackages
    }
}
$projectTessdata = Join-Path $PSScriptRoot "tools\tessdata"
if (Test-Path -LiteralPath $projectTessdata) {
    $env:TESSDATA_PREFIX = $projectTessdata
}

function Test-Command {
    param([string]$Name)
    return [bool](Get-Command $Name -ErrorAction SilentlyContinue)
}

function Require-Command {
    param(
        [string]$Name,
        [string]$InstallHint
    )

    if (-not (Test-Command $Name)) {
        throw "$Name was not found. $InstallHint"
    }
}

function Require-PythonModule {
    param(
        [string]$Module,
        [string]$InstallHint
    )

    $python = Get-PythonCommand
    if (-not $python) {
        throw "Python was not found. $InstallHint"
    }

    & $python -m $Module --version *> $null
    if ($LASTEXITCODE -ne 0) {
        throw "$Module was not found. $InstallHint"
    }
}

function Require-PythonImport {
    param(
        [string]$Module,
        [string]$InstallHint
    )

    $python = Get-PythonCommand
    if (-not $python) {
        throw "Python was not found. $InstallHint"
    }

    & $python -c "import $Module; print('ok')" *> $null
    if ($LASTEXITCODE -ne 0) {
        throw "$Module was not found. $InstallHint"
    }
}

function Require-TesseractLanguages {
    param([string]$LanguageSpec)

    $availableOutput = & tesseract --list-langs 2>&1
    $available = @($availableOutput | Where-Object { $_ -and $_ -notmatch "^List of available languages" })
    $requested = @($LanguageSpec -split "\+" | Where-Object { $_ })
    $missing = @($requested | Where-Object { $available -notcontains $_ })

    if ($missing.Count -gt 0) {
        throw (
            "Tesseract language data is missing: " + ($missing -join ", ") + "`n" +
            "Current tessdata: " + $env:TESSDATA_PREFIX + "`n" +
            "Put the missing .traineddata files in tools\\tessdata, or choose an installed OCR language."
        )
    }
}

function Resolve-InputFile {
    param(
        [string]$Path,
        [string]$Kind = "input"
    )

    $resolved = Resolve-Path -LiteralPath $Path -ErrorAction SilentlyContinue
    if (-not $resolved) {
        throw "$Kind file was not found: $Path"
    }

    $item = Get-Item -LiteralPath $resolved.Path
    if ($item.PSIsContainer) {
        throw "$Kind path is a folder. Please choose a file: $Path"
    }

    return $item.FullName
}

function New-CleanDirectory {
    param([string]$Path)

    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Recurse -Force
    }

    New-Item -ItemType Directory -Path $Path | Out-Null
}

function Export-WordTaggedPdf {
    param(
        [string]$SourceWord,
        [string]$TargetPdf
    )

    Write-Host "[1/5] Exporting tagged PDF from Word..."

    $word = $null
    $document = $null

    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false

        $document = $word.Documents.Open($SourceWord, $false, $true)

        $wdExportFormatPDF = 17
        $wdExportOptimizeForPrint = 0
        $wdExportAllDocument = 0
        $wdExportDocumentContent = 0
        $wdExportCreateHeadingBookmarks = 1

        $document.ExportAsFixedFormat(
            $TargetPdf,
            $wdExportFormatPDF,
            $false,
            $wdExportOptimizeForPrint,
            $wdExportAllDocument,
            1,
            1,
            $wdExportDocumentContent,
            $true,
            $true,
            $wdExportCreateHeadingBookmarks,
            $true,
            $true,
            $false
        )
    }
    catch {
        throw (
            "Word PDF export failed: " + $_.Exception.Message + "`n" +
            "If this is a background or restricted session, open Word on the desktop, export the document as a bookmarked/tagged PDF, " +
            "then rerun this tool with -TaggedPdfPath."
        )
    }
    finally {
        if ($document -ne $null) {
            $document.Close($false)
            [Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
        }

        if ($word -ne $null) {
            $word.Quit()
            [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        }
    }
}

function Convert-PdfToPngPages {
    param(
        [string]$SourcePdf,
        [string]$PagesDir,
        [int]$Resolution,
        [string]$Engine = "pdfium",
        [string]$Format = "jpeg",
        [int]$Quality = 75,
        [int]$Workers = 4
    )

    Write-Host "[2/5] Rendering PDF pages to images..."
    New-CleanDirectory -Path $PagesDir

    $python = Get-PythonCommand
    if (-not $python) {
        throw "Python was not found. Cannot count PDF pages before rendering."
    }

    $pageCountText = & $python -c "import pikepdf, sys; pdf=pikepdf.open(sys.argv[1]); print(len(pdf.pages))" $SourcePdf
    if ($LASTEXITCODE -ne 0) {
        throw "Could not count PDF pages before rendering."
    }

    $pageCount = [int]($pageCountText | Select-Object -First 1)
    Write-Host ("IMAGE_TOTAL_PAGES " + $pageCount)

    if ($Engine -eq "pdfium") {
        Require-PythonImport "pypdfium2" "Install pypdfium2 with install-python-packages.ps1."

        $Workers = [Math]::Max(1, [Math]::Min($Workers, $pageCount))
        Write-Host ("IMAGE_ENGINE pdfium")
        Write-Host ("IMAGE_FORMAT " + $Format)
        Write-Host ("IMAGE_FAST_MODE pdfium parallel rendering with " + $Workers + " workers...")

        $script = Join-Path $PSScriptRoot "tools\render-pdfium-pages.py"
        $arguments = @(
            $script,
            "--input", $SourcePdf,
            "--output-dir", $PagesDir,
            "--dpi", $Resolution,
            "--workers", $Workers,
            "--format", $Format,
            "--jpeg-quality", $Quality
        )

        & $python @arguments 2>&1 | ForEach-Object { Write-Host $_ }
        if ($LASTEXITCODE -ne 0) {
            throw "PDFium rendering failed with exit code: $LASTEXITCODE"
        }
    }
    elseif ($DetailedPngProgress) {
        for ($page = 1; $page -le $pageCount; $page += 1) {
            Write-Host ("IMAGE_PAGE " + $page + "/" + $pageCount)
            $prefix = Join-Path $PagesDir ("page-" + ("{0:D6}" -f $page))
            & pdftoppm -f $page -l $page -r $Resolution -jpeg -jpegopt ("quality=" + $Quality) $SourcePdf $prefix

            if ($LASTEXITCODE -ne 0) {
                throw "pdftoppm failed on page $page with exit code: $LASTEXITCODE"
            }
        }
    }
    else {
        $Workers = [Math]::Max(1, [Math]::Min($Workers, $pageCount))
        Write-Host ("IMAGE_FAST_MODE parallel rendering with " + $Workers + " workers...")
        $prefix = Join-Path $PagesDir "page"

        if ($Workers -eq 1) {
            & pdftoppm -r $Resolution -jpeg -jpegopt ("quality=" + $Quality) $SourcePdf $prefix

            if ($LASTEXITCODE -ne 0) {
                throw "pdftoppm failed with exit code: $LASTEXITCODE"
            }
        }
        else {
            $processes = New-Object System.Collections.Generic.List[object]
            $chunkSize = [Math]::Ceiling($pageCount / $Workers)

            for ($worker = 0; $worker -lt $Workers; $worker += 1) {
                $startPage = [int]($worker * $chunkSize + 1)
                $endPage = [int]([Math]::Min(($worker + 1) * $chunkSize, $pageCount))
                if ($startPage -gt $pageCount) {
                    continue
                }

                Write-Host ("IMAGE_RANGE " + $startPage + "-" + $endPage + "/" + $pageCount)

                $arguments = @("-f", $startPage, "-l", $endPage, "-r", $Resolution, "-jpeg", "-jpegopt", ("quality=" + $Quality), $SourcePdf, $prefix)
                $process = Start-Process -FilePath "pdftoppm" -ArgumentList $arguments -PassThru -WindowStyle Hidden
                [void]$processes.Add($process)
            }

            $lastCount = -1
            while (@($processes | Where-Object { -not $_.HasExited }).Count -gt 0) {
                $currentCount = @(Get-ChildItem -LiteralPath $PagesDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension.ToLowerInvariant() -in @(".jpg", ".jpeg") }).Count
                if ($currentCount -ne $lastCount) {
                    Write-Host ("IMAGE_PAGE " + [Math]::Min($currentCount, $pageCount) + "/" + $pageCount)
                    $lastCount = $currentCount
                }
                Start-Sleep -Seconds 2
            }

            $currentCount = @(Get-ChildItem -LiteralPath $PagesDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension.ToLowerInvariant() -in @(".jpg", ".jpeg") }).Count
            Write-Host ("IMAGE_PAGE " + [Math]::Min($currentCount, $pageCount) + "/" + $pageCount)

            foreach ($process in $processes) {
                $process.WaitForExit()
                if ($process.ExitCode -ne 0) {
                    throw "pdftoppm failed with exit code: $($process.ExitCode)"
                }
                $process.Dispose()
            }
        }
    }

    $pages = @(Get-ChildItem -LiteralPath $PagesDir -File |
        Where-Object { $_.Extension.ToLowerInvariant() -in @(".jpg", ".jpeg") } |
        Sort-Object {
            if ($_.BaseName -match "(\d+)$") { [int]$Matches[1] } else { 0 }
        })

    if ($pages.Count -eq 0) {
        throw "No JPG pages were created."
    }

    return @($pages.FullName)
}

function Merge-PngPagesToPdf {
    param(
        [string[]]$ImageFiles,
        [string]$SourcePdf,
        [string]$TargetPdf,
        [string]$WorkDirectory
    )

    Write-Host "[3/5] Merging rendered pages into image-only PDF..."

    $python = Get-PythonCommand
    if (-not $python) {
        throw "Python was not found. Install Python and img2pdf, then rerun the tool."
    }

    $listPath = Join-Path $WorkDirectory "image-pages.txt"
    [System.IO.File]::WriteAllLines($listPath, $ImageFiles, [System.Text.UTF8Encoding]::new($false))

    $script = Join-Path $PSScriptRoot "tools\merge-png-pdf.py"
    & $python $script --list $listPath --source-pdf $SourcePdf --output $TargetPdf

    if ($LASTEXITCODE -ne 0) {
        throw "Image-to-PDF merge failed with exit code: $LASTEXITCODE"
    }
}

function Invoke-Ocr {
    param(
        [string]$SourcePdf,
        [string]$TargetPdf,
        [string]$OcrLanguage
    )

    Write-Host "[4/5] Running OCR and creating hidden text layer..."

    $outputType = if ($Optimize) { "pdfa" } else { "pdf" }
    $arguments = @(
        "--language", $OcrLanguage,
        "--output-type", $outputType,
        "--force-ocr"
    )

    if ($Deskew) {
        $arguments += "--deskew"
    }

    if ($RotatePages) {
        $arguments += "--rotate-pages"
    }

    if ($Optimize) {
        $arguments += @(
            "--optimize", "1",
            "--pdfa-image-compression", "jpeg",
            "--ghostscript-jpeg-quality", "50",
            "--fast-web-view", "0"
        )
    }

    $arguments += @($SourcePdf, $TargetPdf)

    $python = Get-PythonCommand
    if ($python) {
        & $python -m ocrmypdf @arguments
    }
    else {
        & ocrmypdf @arguments
    }

    if ($LASTEXITCODE -ne 0) {
        throw "OCRmyPDF failed with exit code: $LASTEXITCODE"
    }
}

function Get-PythonCommand {
    $projectPython = Join-Path $PSScriptRoot "tools\Python313\python.exe"
    if (Test-Path -LiteralPath $projectPython) {
        return $projectPython
    }

    $localPython = "C:\Users\Administrator\AppData\Local\Python\bin\python.exe"
    if (Test-Path -LiteralPath $localPython) {
        return $localPython
    }

    $pyCommand = Get-Command "py" -ErrorAction SilentlyContinue
    if ($pyCommand -and $pyCommand.Source -notlike "*\Microsoft\WindowsApps\*") {
        return $pyCommand.Source
    }

    $pythonCommand = Get-Command "python" -ErrorAction SilentlyContinue
    if ($pythonCommand -and $pythonCommand.Source -notlike "*\Microsoft\WindowsApps\*") {
        return $pythonCommand.Source
    }

    return ""
}

function Restore-Bookmarks {
    param(
        [string]$TaggedPdf,
        [string]$OcrPdf,
        [string]$FinalPdf
    )

    Write-Host "[5/5] Restoring bookmarks to final PDF..."

    $python = Get-PythonCommand
    if (-not $python) {
        throw "Python was not found. Install Python and pikepdf, then rerun the tool."
    }

    $script = Join-Path $PSScriptRoot "tools\restore-bookmarks.py"
    & $python $script --source $TaggedPdf --target $OcrPdf --output $FinalPdf

    if ($LASTEXITCODE -ne 0) {
        throw "Bookmark restore failed with exit code: $LASTEXITCODE"
    }
}

if (-not $WordPath -and -not $TaggedPdfPath) {
    throw "Please provide -WordPath for the full workflow, or -TaggedPdfPath to start from an already exported PDF."
}

if ($WordPath -and $TaggedPdfPath) {
    throw "Please provide either -WordPath or -TaggedPdfPath, not both."
}

$sourceWord = ""
if ($WordPath) {
    $sourceWord = Resolve-InputFile -Path $WordPath -Kind "Word"
}

$sourceTaggedPdf = ""
if ($TaggedPdfPath) {
    $sourceTaggedPdf = Resolve-InputFile -Path $TaggedPdfPath -Kind "Tagged PDF"
}

$finalPdf = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
$finalFolder = Split-Path -Parent $finalPdf

if ($finalFolder -and -not (Test-Path -LiteralPath $finalFolder)) {
    New-Item -ItemType Directory -Path $finalFolder | Out-Null
}

if ($RenderEngine -eq "poppler") {
    Require-Command "pdftoppm" "Install Poppler for Windows and add its bin folder to PATH."
}
Require-Command "tesseract" "Install Tesseract OCR and the required language packs."
Require-TesseractLanguages $Language
Require-PythonModule "ocrmypdf" "Install OCRmyPDF with install-python-packages.ps1."
Require-PythonModule "img2pdf" "Install img2pdf with install-python-packages.ps1."

$baseName = [IO.Path]::GetFileNameWithoutExtension($sourceWord)
if ($sourceTaggedPdf) {
    $baseName = [IO.Path]::GetFileNameWithoutExtension($sourceTaggedPdf)
}
$runName = "dual-pdf-" + $baseName + "-" + [guid]::NewGuid().ToString("N")

if (-not $WorkDir) {
    $WorkRoot = [IO.Path]::GetTempPath()
}
else {
    $WorkRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($WorkDir)
    New-Item -ItemType Directory -Path $WorkRoot -Force | Out-Null
}

$WorkDir = Join-Path $WorkRoot $runName
New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null

$taggedPdf = Join-Path $WorkDir "01-word-tagged.pdf"
$pagesDir = Join-Path $WorkDir "02-pages"
$imagePdf = Join-Path $WorkDir "03-image-only.pdf"
$ocrPdf = Join-Path $WorkDir "04-ocr.pdf"
$renderImageFormat = $ImageFormat

try {
    if ($sourceWord) {
        Export-WordTaggedPdf -SourceWord $sourceWord -TargetPdf $taggedPdf
    }
    else {
        Write-Host "[1/5] Using existing tagged/bookmarked PDF..."
        Copy-Item -LiteralPath $sourceTaggedPdf -Destination $taggedPdf -Force
    }

    $imagePages = @(Convert-PdfToPngPages -SourcePdf $taggedPdf -PagesDir $pagesDir -Resolution $Dpi -Engine $RenderEngine -Format $renderImageFormat -Quality $JpegQuality -Workers $RenderWorkers)
    Write-Host ("OCR_TOTAL_PAGES " + $imagePages.Count)
    Merge-PngPagesToPdf -ImageFiles $imagePages -SourcePdf $taggedPdf -TargetPdf $imagePdf -WorkDirectory $WorkDir
    Invoke-Ocr -SourcePdf $imagePdf -TargetPdf $ocrPdf -OcrLanguage $Language
    Restore-Bookmarks -TaggedPdf $taggedPdf -OcrPdf $ocrPdf -FinalPdf $finalPdf

    Write-Host "Done: $finalPdf"
}
finally {
    if (-not $KeepWork -and (Test-Path -LiteralPath $WorkDir)) {
        Remove-Item -LiteralPath $WorkDir -Recurse -Force
    }
    elseif ($KeepWork) {
        Write-Host "Work files kept at: $WorkDir"
    }
}
