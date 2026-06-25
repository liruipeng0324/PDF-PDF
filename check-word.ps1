Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

try {
    $word = New-Object -ComObject Word.Application
    $version = $word.Version
    $word.Quit()
    Write-Host "[found]   Microsoft Word COM -> version $version"
    Write-Host "          Full -WordPath workflow can be used from this session."
}
catch {
    Write-Host "[missing] Microsoft Word COM is not available in this session"
    Write-Host ("          " + $_.Exception.Message)
    Write-Host "          Use Word's desktop Export/Save As PDF first, then run make-dual-pdf.ps1 with -TaggedPdfPath."
}

