# PDF-PDF 双层 PDF 制作工具

这是一个本地网页工具，用来制作“双层 PDF”：页面外观来自原文档渲染出的图片，底层再加入 OCR 文字层，最后把原 PDF 的书签加回最终 PDF。

## 工作流程

1. Word 导出为带标签/书签的 PDF，或直接使用已经导出的带书签 PDF。
2. PDF 页面转为 PNG 图片。
3. PNG 图片合并为新的图片型 PDF。
4. OCR 扫描并加入隐藏文字层。
5. 把原 PDF 书签恢复到最终 OCR PDF。

注意：PDF 转图片后，原来的结构标签会丢失；最终 PDF 保留页面外观、OCR 可搜索文字和书签。

## 直接使用

双击：

```text
start-web.bat
```

浏览器会打开：

```text
http://127.0.0.1:8765/
```

网页支持：

- 单个 Word / PDF 文件制作。
- 文件夹队列制作。
- CSV 队列制作。
- 本地文件/文件夹选择。
- 实时进度、PNG 页数、OCR 页数。
- 依赖检测。
- 取消任务。
- 查看最近输出。
- 打开输出位置和日志目录。
- 清理缓存。

## 新电脑使用

推荐两种方式：

1. 源码方式：下载仓库后，安装或准备 Python、Tesseract OCR、Ghostscript、Poppler、qpdf，再运行 `setup-dependencies.ps1` 检查。
2. 便携方式：在已经配置好的电脑上运行 `build-portable-package.ps1`，把生成的 zip 复制到新电脑，解压后双击 `start-web.bat`。

如果要直接处理 Word 文件，新电脑必须安装 Microsoft Word。没有 Word 时，可以先手动把 Word 导出为带书签 PDF，再在网页里选择“已导出的带书签 PDF”模式。

## 依赖检查

```powershell
powershell -ExecutionPolicy Bypass -File .\setup-dependencies.ps1
```

或者：

```powershell
powershell -ExecutionPolicy Bypass -File .\check-tools.ps1
```

检查 Word 自动导出：

```powershell
powershell -ExecutionPolicy Bypass -File .\check-word.ps1
```

## 自测

```powershell
powershell -ExecutionPolicy Bypass -File .\run-self-test.ps1
```

预期结果：

```text
Self-test passed
```

## 制作便携发布包

```powershell
powershell -ExecutionPolicy Bypass -File .\build-portable-package.ps1
```

这个包会包含当前目录下的源码、启动脚本和 `tools` 便携组件，但不会包含 `.git`、测试输出和已有 zip。

## 命令行用法

从 Word 处理：

```powershell
powershell -ExecutionPolicy Bypass -File .\make-dual-pdf.ps1 `
  -WordPath "source.docx" `
  -OutputPath "final-dual-layer.pdf" `
  -Language "chi_sim+eng" `
  -Dpi 300 `
  -Deskew `
  -RotatePages
```

从已导出的带书签 PDF 处理：

```powershell
powershell -ExecutionPolicy Bypass -File .\make-dual-pdf.ps1 `
  -TaggedPdfPath "source-bookmarked.pdf" `
  -OutputPath "final-dual-layer.pdf" `
  -Language "chi_sim+eng" `
  -Dpi 300 `
  -Deskew `
  -RotatePages
```

文件夹队列：

```powershell
powershell -ExecutionPolicy Bypass -File .\run-queue.ps1 `
  -InputDir "E:\Docs\Input" `
  -OutputDir "E:\Docs\Output" `
  -InputType pdf `
  -Language "chi_sim+eng" `
  -Dpi 300 `
  -SkipExisting
```

CSV 队列列名：

```csv
InputPath,OutputPath,Type
E:\Docs\a.pdf,E:\Docs\Output\a-dual-layer.pdf,pdf
E:\Docs\b.docx,E:\Docs\Output\b-dual-layer.pdf,word
```

每次队列运行会在输出目录的 `_queue-logs` 里生成日志和汇总 CSV。

## 清理

只清理缓存：

```powershell
powershell -ExecutionPolicy Bypass -File .\clean-cache.ps1
```

清理安装包、测试输出和缓存：

```powershell
powershell -ExecutionPolicy Bypass -File .\clean-project.ps1
```
