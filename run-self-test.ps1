Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$python = "C:\Users\Administrator\AppData\Local\Python\bin\python.exe"
if (-not (Test-Path -LiteralPath $python)) {
    throw "Python was not found at $python"
}

$env:PYTHONPATH = Join-Path $PSScriptRoot "tools\python-packages"
$testDir = Join-Path $PSScriptRoot "test-output"
New-Item -ItemType Directory -Path $testDir -Force | Out-Null

$sourcePdf = Join-Path $testDir "source-bookmarked.pdf"
$finalPdf = Join-Path $testDir "final-ocr-bookmarked.pdf"

@'
from pathlib import Path
from fpdf import FPDF
import pikepdf
from pikepdf import OutlineItem

out = Path("test-output/source-bookmarked.pdf")
pdf = FPDF(unit="pt", format="A4")
for title, body in [
    ("Chapter One", "This is a test page for OCR."),
    ("Chapter Two", "Second page with searchable text after OCR."),
]:
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 24)
    pdf.cell(0, 40, title, new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 14)
    pdf.multi_cell(0, 22, body)
pdf.output(str(out))

with pikepdf.open(out, allow_overwriting_input=True) as doc:
    with doc.open_outline() as outline:
        outline.root.append(OutlineItem("Chapter One", doc.pages[0].obj))
        outline.root.append(OutlineItem("Chapter Two", doc.pages[1].obj))
    doc.save(out)
'@ | & $python

& powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "make-dual-pdf.ps1") `
    -TaggedPdfPath $sourcePdf `
    -OutputPath $finalPdf `
    -Language "eng" `
    -Dpi 150

@'
from pathlib import Path
import pikepdf
from pdfminer.high_level import extract_text

pdf_path = Path("test-output/final-ocr-bookmarked.pdf")
with pikepdf.open(pdf_path) as pdf:
    assert len(pdf.pages) == 2, "Expected 2 pages"
    outlines = pdf.Root.get("/Outlines")
    assert outlines is not None, "Expected bookmarks"
    titles = []
    node = outlines.get("/First")
    while node is not None:
        titles.append(str(node.get("/Title")))
        node = node.get("/Next")
    assert titles == ["Chapter One", "Chapter Two"], titles

text = extract_text(str(pdf_path))
assert "Chapter One" in text, "OCR text missing Chapter One"
assert "Chapter Two" in text, "OCR text missing Chapter Two"
print("Self-test passed:", pdf_path.resolve())
'@ | & $python

