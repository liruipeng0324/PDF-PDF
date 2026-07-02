param(
    [ValidateNotNullOrEmpty()]
    [string]$WordPath = "",

    [ValidateNotNullOrEmpty()]
    [string]$TaggedPdfPath = "",

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath,

    [ValidateRange(72, 600)]
    [int]$Dpi = 300,

    [string]$WorkDir = "",

    [switch]$Optimize,
    [switch]$KeepWork,
    [switch]$DetailedPngProgress,
    [switch]$AcrobatDoubleLayer,

    [ValidateSet("CHS", "ENG", "CHT")]
    [string]$AcrobatOcrLanguage = "CHS",

    [ValidateSet("pdfium", "poppler")]
    [string]$RenderEngine = "pdfium",

    [ValidateSet("png", "jpeg")]
    [string]$ImageFormat = "png",

    [ValidateRange(1, 100)]
    [int]$JpegQuality = 95,

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
Add-ToolDirectory (Join-Path $PSScriptRoot "tools\ghostscript-10.07.1\bin")
Add-ToolDirectory (Join-Path $PSScriptRoot "tools\Python312")
Add-ToolDirectory (Join-Path $PSScriptRoot "tools\Python312\Scripts")
Add-ToolDirectory "D:\GPL\gs10.07.1\bin"
$env:Path = [Environment]::GetEnvironmentVariable("Path", "User") + ";" +
    [Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
    $env:Path
$projectPythonPackages = Join-Path $PSScriptRoot "tools\python-packages"
$projectPython312 = Join-Path $PSScriptRoot "tools\Python312\python.exe"
if ((Test-Path -LiteralPath $projectPythonPackages) -and -not (Test-Path -LiteralPath $projectPython312)) {
    if ($env:PYTHONPATH) {
        $env:PYTHONPATH = "$projectPythonPackages;$env:PYTHONPATH"
    }
    else {
        $env:PYTHONPATH = $projectPythonPackages
    }
}
elseif (Test-Path -LiteralPath $projectPython312) {
    $env:PYTHONPATH = ""
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
    $openWordPath = $SourceWord

    try {
        $targetFolder = Split-Path -Parent $TargetPdf
        if ($targetFolder) {
            $extension = [IO.Path]::GetExtension($SourceWord)
            if (-not $extension) {
                $extension = ".docx"
            }
            $openWordPath = Join-Path $targetFolder ("00-source" + $extension)
            Copy-Item -LiteralPath $SourceWord -Destination $openWordPath -Force
            Write-Host ("WORD_OPEN_COPY " + $openWordPath)
        }

        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0

        try {
            $document = $word.Documents.Open($openWordPath, $false, $true)
        }
        catch {
            $document = $word.Documents.OpenNoRepairDialog($openWordPath, $false, $true)
        }
        if ($null -eq $document) {
            throw "Word opened but did not return a document. Open Microsoft Word once on the desktop, close any activation/protected-view prompts, then retry."
        }

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
            "Open Microsoft Word once on the desktop and close any activation, update, file recovery, or protected-view prompts. " +
            "If Word still cannot export automatically, export the document as a bookmarked/tagged PDF in Word, then choose the PDF input mode."
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
            if ($Format -eq "jpeg") {
                & pdftoppm -f $page -l $page -r $Resolution -jpeg -jpegopt ("quality=" + $Quality) $SourcePdf $prefix
            }
            else {
                & pdftoppm -f $page -l $page -r $Resolution -png $SourcePdf $prefix
            }

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
            if ($Format -eq "jpeg") {
                & pdftoppm -r $Resolution -jpeg -jpegopt ("quality=" + $Quality) $SourcePdf $prefix
            }
            else {
                & pdftoppm -r $Resolution -png $SourcePdf $prefix
            }

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

                $arguments = @("-f", $startPage, "-l", $endPage, "-r", $Resolution)
                if ($Format -eq "jpeg") {
                    $arguments += @("-jpeg", "-jpegopt", ("quality=" + $Quality))
                }
                else {
                    $arguments += "-png"
                }
                $arguments += @($SourcePdf, $prefix)
                $process = Start-Process -FilePath "pdftoppm" -ArgumentList $arguments -PassThru -WindowStyle Hidden
                [void]$processes.Add($process)
            }

            $lastCount = -1
            while (@($processes | Where-Object { -not $_.HasExited }).Count -gt 0) {
                $currentCount = @(Get-ChildItem -LiteralPath $PagesDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension.ToLowerInvariant() -in @(".png", ".jpg", ".jpeg") }).Count
                if ($currentCount -ne $lastCount) {
                    Write-Host ("IMAGE_PAGE " + [Math]::Min($currentCount, $pageCount) + "/" + $pageCount)
                    $lastCount = $currentCount
                }
                Start-Sleep -Seconds 2
            }

            $currentCount = @(Get-ChildItem -LiteralPath $PagesDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension.ToLowerInvariant() -in @(".png", ".jpg", ".jpeg") }).Count
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
        Where-Object { $_.Extension.ToLowerInvariant() -in @(".png", ".jpg", ".jpeg") } |
        Sort-Object {
            if ($_.BaseName -match "(\d+)$") { [int]$Matches[1] } else { 0 }
        })

    if ($pages.Count -eq 0) {
        throw "No rendered image pages were created."
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

function Get-PythonCommand {
    $projectPython312 = Join-Path $PSScriptRoot "tools\Python312\python.exe"
    if (Test-Path -LiteralPath $projectPython312) {
        return $projectPython312
    }

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

function Compress-ScannedPdf {
    param(
        [string]$SourcePdf,
        [string]$TargetPdf
    )

    if (-not $Optimize) {
        Copy-Item -LiteralPath $SourcePdf -Destination $TargetPdf -Force
        return
    }

    $gsCommand = Get-Command "gswin64c" -ErrorAction SilentlyContinue
    if (-not $gsCommand) {
        Write-Host "SCAN_OPTIMIZE Ghostscript not found; skipping scanned image compression"
        Copy-Item -LiteralPath $SourcePdf -Destination $TargetPdf -Force
        return
    }

    Write-Host "SCAN_OPTIMIZE Ghostscript 300DPI scanned PDF compression at JPEG quality 85"
    & $gsCommand.Source `
        "-dSAFER" `
        "-dBATCH" `
        "-dNOPAUSE" `
        "-sDEVICE=pdfwrite" `
        "-dCompatibilityLevel=1.6" `
        "-dDetectDuplicateImages=true" `
        "-dCompressFonts=true" `
        "-dSubsetFonts=true" `
        "-dAutoRotatePages=/None" `
        "-dColorImageDownsampleType=/Bicubic" `
        "-dColorImageResolution=300" `
        "-dColorImageDownsampleThreshold=1.0" `
        "-dGrayImageDownsampleType=/Bicubic" `
        "-dGrayImageResolution=300" `
        "-dGrayImageDownsampleThreshold=1.0" `
        "-dMonoImageDownsampleType=/Subsample" `
        "-dMonoImageResolution=300" `
        "-dMonoImageDownsampleThreshold=1.0" `
        "-dAutoFilterColorImages=false" `
        "-dColorImageFilter=/DCTEncode" `
        "-dAutoFilterGrayImages=false" `
        "-dGrayImageFilter=/DCTEncode" `
        "-dJPEGQ=85" `
        "-sOutputFile=$TargetPdf" `
        $SourcePdf

    if ($LASTEXITCODE -ne 0) {
        throw "Ghostscript scanned PDF compression failed with exit code: $LASTEXITCODE"
    }

    $sourceSize = (Get-Item -LiteralPath $SourcePdf).Length
    $targetSize = (Get-Item -LiteralPath $TargetPdf).Length
    if ($targetSize -ge $sourceSize) {
        Write-Host ("SCAN_OPTIMIZE Ghostscript output was not smaller; keeping image PDF ({0} -> {1} bytes)" -f $sourceSize, $targetSize)
        Copy-Item -LiteralPath $SourcePdf -Destination $TargetPdf -Force
    }
    else {
        Write-Host ("SCAN_OPTIMIZE Ghostscript reduced size ({0} -> {1} bytes)" -f $sourceSize, $targetSize)
    }
}

function Convert-WithAcrobatDoubleLayer {
    param(
        [string]$SourcePdf,
        [string]$TargetPdf,
        [string]$Language = "CHS"
    )

    Write-Host "[4/5] Creating Acrobat searchable-image PDF..."

    $python = Get-PythonCommand
    if (-not $python) {
        throw "Python was not found. Install Python and pywin32, then rerun the tool."
    }

    $modulePath = Join-Path $PSScriptRoot "acrobat_double_pdf.py"
    if (-not (Test-Path -LiteralPath $modulePath)) {
        throw "Acrobat double-layer module was not found: $modulePath"
    }

    & "$python" "$modulePath" --input "$SourcePdf" --output "$TargetPdf" --lang "$Language"
    if ($LASTEXITCODE -ne 0) {
        throw "Acrobat double-layer conversion failed with exit code: $LASTEXITCODE"
    }
}

function Restore-Bookmarks {
    param(
        [string]$TaggedPdf,
        [string]$TargetPdf,
        [string]$FinalPdf
    )

    Write-Host "[5/5] Restoring bookmarks to final PDF..."

    $python = Get-PythonCommand
    if (-not $python) {
        throw "Python was not found. Install Python and pikepdf, then rerun the tool."
    }

    $script = Join-Path $PSScriptRoot "tools\restore-bookmarks.py"
    & $python $script --source $TaggedPdf --target $TargetPdf --output $FinalPdf

    if ($LASTEXITCODE -ne 0) {
        throw "Bookmark restore failed with exit code: $LASTEXITCODE"
    }
}

function Optimize-FinalPdf {
    param(
        [string]$PdfPath
    )

    if (-not $Optimize) {
        return
    }

    $qpdfCommand = Get-Command "qpdf" -ErrorAction SilentlyContinue
    if (-not $qpdfCommand) {
        Write-Host "SCAN_OPTIMIZE qpdf not found; skipping final structure compression"
        return
    }

    $tempPdf = [IO.Path]::Combine(
        [IO.Path]::GetDirectoryName($PdfPath),
        ([IO.Path]::GetFileNameWithoutExtension($PdfPath) + ".qpdf.tmp.pdf")
    )

    try {
        Write-Host "SCAN_OPTIMIZE qpdf object and stream compression"
        & $qpdfCommand.Source `
            "--object-streams=generate" `
            "--stream-data=compress" `
            "--recompress-flate" `
            "--linearize" `
            $PdfPath `
            $tempPdf

        if ($LASTEXITCODE -ne 0) {
            throw "qpdf failed with exit code: $LASTEXITCODE"
        }

        $sourceSize = (Get-Item -LiteralPath $PdfPath).Length
        $targetSize = (Get-Item -LiteralPath $tempPdf).Length
        if ($targetSize -lt $sourceSize) {
            Write-Host ("SCAN_OPTIMIZE qpdf reduced size ({0} -> {1} bytes)" -f $sourceSize, $targetSize)
            Move-Item -LiteralPath $tempPdf -Destination $PdfPath -Force
        }
        else {
            Write-Host ("SCAN_OPTIMIZE qpdf output was not smaller; keeping restored PDF ({0} -> {1} bytes)" -f $sourceSize, $targetSize)
        }
    }
    finally {
        if (Test-Path -LiteralPath $tempPdf) {
            Remove-Item -LiteralPath $tempPdf -Force
        }
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
$compressedImagePdf = Join-Path $WorkDir "04-image-compressed.pdf"
$doubleLayerPdf = Join-Path $WorkDir "04-acrobat-double-layer.pdf"
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
    Merge-PngPagesToPdf -ImageFiles $imagePages -SourcePdf $taggedPdf -TargetPdf $imagePdf -WorkDirectory $WorkDir
    if ($AcrobatDoubleLayer) {
        Convert-WithAcrobatDoubleLayer -SourcePdf $imagePdf -TargetPdf $doubleLayerPdf -Language $AcrobatOcrLanguage
        Restore-Bookmarks -TaggedPdf $taggedPdf -TargetPdf $doubleLayerPdf -FinalPdf $finalPdf
    }
    else {
        Compress-ScannedPdf -SourcePdf $imagePdf -TargetPdf $compressedImagePdf
        Restore-Bookmarks -TaggedPdf $taggedPdf -TargetPdf $compressedImagePdf -FinalPdf $finalPdf
    }
    Optimize-FinalPdf -PdfPath $finalPdf

    Write-Host "Done: $finalPdf"
}
catch {
    if ($AcrobatDoubleLayer) {
        Write-Host "Work files kept after Acrobat failure at: $WorkDir"
        $KeepWork = $true
    }
    throw
}
finally {
    if (-not $KeepWork -and (Test-Path -LiteralPath $WorkDir)) {
        Remove-Item -LiteralPath $WorkDir -Recurse -Force
    }
    elseif ($KeepWork) {
        Write-Host "Work files kept at: $WorkDir"
    }
}
