# PDF-PDF 双层 PDF 制作工具

## 最新 EXE 交付版 v1.0.0

**[下载完整交付包（约 176 MiB）](https://github.com/liruipeng0324/PDF-PDF/releases/download/v1.0.0/PDF-Tools-EXE-v1.0.0.zip)** · [版本说明与校验文件](https://github.com/liruipeng0324/PDF-PDF/releases/tag/v1.0.0) · [PDF 使用说明书](https://github.com/liruipeng0324/PDF-PDF/releases/download/v1.0.0/PDF-Tools-v1.0.0-Manual.pdf)

完整解压到专用目录，双击 **PDF工具.exe**；无需安装 Python。支持 PDF、DOC、DOCX、跨目录追加文件及“生成在原目录”，包内附一键卸载。

DOC/DOCX 输入需要已安装并激活的 Microsoft Word。**Acrobat OCR 尚未完成验收，请保持“Acrobat 双层PDF”不勾选；未执行 OCR 的输出没有可搜索文字层。**

以下为仓库早期源码版说明。EXE 交付版请以本页上方的 Release 及随包说明书为准；GitHub 自动生成的 Source code 压缩包不含 EXE 运行组件。

---

这是一个 Windows 本地 Web 工具，用来把 Word 文档或已导出的带书签 PDF 制作成可搜索、带书签的双层 PDF。

当前推荐流程使用 Adobe Acrobat Pro 的“扫描和 OCR”功能生成文字层，再把原始目录书签恢复到最终 PDF。

## 快速开始

第一次从 GitHub 下载后，先双击：

```text
一键安装组件.bat
```

安装完成后，双击启动：

```text
start-web.bat
```

浏览器打开后：

1. 选择输入文件，默认是 Word 文档。
2. 选择输出文件夹。
3. 保持 300 DPI 和 Acrobat OCR 默认设置。
4. 点击“开始制作”。

## 主要功能

- Word 自动导出为带书签 PDF
- 已导出的带书签 PDF 直接制作
- 多文件队列处理
- 300 DPI 图片化中间页
- Adobe Acrobat Pro 自动 OCR
- OCR 期间识别中文扫描状态，避免重复点击扫描
- 扫描完成后自动保存、验证文字层、恢复书签
- 输出日志、队列汇总、缓存清理

## 必要组件

必须准备：

- Microsoft Word：处理 Word 输入时需要
- Adobe Acrobat Pro：用于“扫描和 OCR”，Reader 不支持
- Python 3.12：一键安装脚本会安装到项目目录
- Poppler / PDFium / img2pdf / pikepdf / qpdf / pywinauto：一键安装脚本会检查并只补装缺失项

重复运行 `一键安装组件.bat` 不会重复下载或安装已满足的组件；只有显式运行 `install-components.ps1 -Force` 才会强制升级或重装。

## 详细说明

完整安装、使用、队列和常见问题请看：

[使用说明.md](./使用说明.md)
