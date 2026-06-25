# Word to Dual-Layer PDF Tool

This project builds a searchable image-based PDF from a Word document or from an already exported bookmarked PDF.

## 中文说明

这是一个本地网页工具，用来制作“双层 PDF”：外观是逐页图片，底层有 OCR 文字层，并把原 PDF 书签加回最终文件。

在一台新电脑上使用时：

1. 如果要直接处理 Word 文件，需要先安装 Microsoft Word。
2. 如果没有 Word，可以先手动把 Word 导出为带书签的 PDF，再在网页里选择“已导出的 PDF”模式。
3. 本仓库默认只保存源码和启动脚本，不把 `tools` 目录、安装包、测试输出提交到 Git。
4. 需要准备 Python、Tesseract OCR、Ghostscript、Poppler、qpdf，以及 `requirements.txt` 里的 Python 依赖。
5. 配好依赖后双击 `start-web.bat`，浏览器会打开 `http://127.0.0.1:8765/`。

Workflow:

1. Export Word to a tagged/bookmarked PDF.
2. Render PDF pages to PNG.
3. Merge PNG pages into a new image-only PDF.
4. Run OCR to add a hidden text layer.
5. Copy bookmarks from the source PDF back to the final OCR PDF.

Important: Word structural tags are lost when the PDF is rasterized to PNG. The final PDF keeps page appearance, OCR text, and restored bookmarks.

## Files

- `make-dual-pdf.ps1`: main workflow.
- `run-queue.ps1`: batch queue workflow.
- `queue-template.csv`: CSV queue template.
- `check-tools.ps1`: dependency checker.
- `check-word.ps1`: checks whether Word automation is available in the current session.
- `run-self-test.ps1`: creates a small test PDF and verifies OCR text plus bookmarks.
- `tools/restore-bookmarks.py`: bookmark restore helper.
- `index.html`, `styles.css`, `app.js`: local command builder UI.
- `web-server.py`, `start-web.ps1`: local web UI with file/folder picker and start button.
- `clean-project.ps1`: removes downloaded installers, test output, and Python cache files.
- `clean-cache.ps1`: removes only Python cache files.

## Check Tools

```powershell
powershell -ExecutionPolicy Bypass -File .\check-tools.ps1
```

All required tools should show `[found]`.

Check whether automatic Word export is available:

```powershell
powershell -ExecutionPolicy Bypass -File .\check-word.ps1
```

## Self Test

```powershell
powershell -ExecutionPolicy Bypass -File .\run-self-test.ps1
```

Expected result: `Self-test passed`.

## Web UI

Start the local web UI:

```powershell
powershell -ExecutionPolicy Bypass -File .\start-web.ps1
```

The browser opens `http://127.0.0.1:8765/`.

If you do not want to use PowerShell manually, double-click:

```text
start-web.bat
```

The web UI can:

- Browse local drives and folders.
- Select a source Word/PDF/CSV or input folder.
- Select an output folder. In single-file mode, the final PDF name is generated automatically.
- Start single-file or queue jobs.
- Show live job status, progress, and logs.

## Cleanup

After all tools are installed, remove large installer files and test artifacts:

```powershell
powershell -ExecutionPolicy Bypass -File .\clean-project.ps1
```

Only remove cache files:

```powershell
powershell -ExecutionPolicy Bypass -File .\clean-cache.ps1
```

## Use From Word

```powershell
powershell -ExecutionPolicy Bypass -File .\make-dual-pdf.ps1 `
  -WordPath "source.docx" `
  -OutputPath "final-dual-layer.pdf" `
  -Language "chi_sim+eng" `
  -Dpi 300 `
  -Deskew `
  -RotatePages
```

This requires Microsoft Word COM automation to work in the current desktop session.

## Use From Existing Bookmarked PDF

```powershell
powershell -ExecutionPolicy Bypass -File .\make-dual-pdf.ps1 `
  -TaggedPdfPath "source-bookmarked.pdf" `
  -OutputPath "final-dual-layer.pdf" `
  -Language "chi_sim+eng" `
  -Dpi 300 `
  -Deskew `
  -RotatePages
```

Use this mode when Word export has already been done manually.

## Queue From Folder

Process all PDFs in a folder:

```powershell
powershell -ExecutionPolicy Bypass -File .\run-queue.ps1 `
  -InputDir "E:\Docs\Input" `
  -OutputDir "E:\Docs\Output" `
  -InputType pdf `
  -Language "chi_sim+eng" `
  -Dpi 300 `
  -Deskew `
  -RotatePages `
  -SkipExisting
```

Process all Word files in a folder:

```powershell
powershell -ExecutionPolicy Bypass -File .\run-queue.ps1 `
  -InputDir "E:\Docs\Input" `
  -OutputDir "E:\Docs\Output" `
  -InputType word `
  -Language "chi_sim+eng" `
  -Dpi 300
```

Use `-Recurse` to include subfolders.

## Queue From CSV

CSV columns:

- `InputPath`: source Word or bookmarked PDF.
- `OutputPath`: optional final PDF path.
- `Type`: `pdf`, `word`, or `auto`.

Example:

```csv
InputPath,OutputPath,Type
E:\Docs\a.pdf,E:\Docs\Output\a-dual-layer.pdf,pdf
E:\Docs\b.docx,E:\Docs\Output\b-dual-layer.pdf,word
```

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\run-queue.ps1 `
  -QueueCsv "E:\Docs\queue.csv" `
  -OutputDir "E:\Docs\Output" `
  -Language "chi_sim+eng" `
  -Dpi 300 `
  -SkipExisting
```

Each run writes logs to `_queue-logs` in the output folder. Failed items do not stop the rest of the queue.
