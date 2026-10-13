# PDF-PDF 2026-10-13 离线安装包

适用于 Windows 10/11 x64。

大文件直传在当前网络下被 GitHub 返回 408 超时，因此本次使用分片发布。
请从 [本次 Release](https://github.com/liruipeng0324/PDF-PDF/releases/tag/v2026.10.13-offline)
下载 Get-Packages.bat、Get-Packages.ps1、parts-manifest.json，放在同一目录，
双击 Get-Packages.bat，选择手动版、完整版或两者。脚本会下载缺失分片并校验、合并到 packages 目录。
失败后可再次运行，已校验的分片会跳过。合并后得到：
- PDF-PDF-full.zip：完整版本，293,921,875 字节（280.3 MiB）。
- PDF-PDF-manual-ocr.zip：手动 OCR 版本，293,578,806 字节（280.0 MiB）。

GitHub 自动生成的 Source code ZIP/TAR 仅包含仓库源码，不包含离线运行组件，请勿将它当作安装包。
如需离线合并，将所需的 .partNNN 附件放入脚本同目录的 parts 文件夹，
执行 powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Get-Packages.ps1 -Profile all -Offline。
合并前逐片验证 SHA-256，合并后验证原 ZIP 的大小和 SHA-256。

## 修复内容

- 内置 Ghostscript 10.07.1、初始化/字体/ICC 资源及 Visual C++ 运行库。
- 安装、网页检查与任务执行使用统一的组件查找规则。
- Ghostscript 检查不访问 GitHub；已满足的组件跳过，重复运行不重装。
- 缺失或执行失败明确返回非零退出码；最终检查错误码逐层传递。
- 修复带空格路径的 Web 启动；启动日志按时间保留。
- 手动版明确提示“待手动 OCR”，保持手动识别文字的使用方式。
- 删除交付包内无用的打包脚本、缓存、测试日志与开发运行库文件。

## 使用

完整解压到新目录，双击“一键安装组件.bat”，检查通过后双击 start-web.bat。
安装日志保存在 install-logs/install-latest.log；启动日志在 run-logs。

Word 与 Acrobat Pro 不随包分发。Word 输入需要 Microsoft Word。
完整版自动 OCR 需要激活的 Acrobat Pro、可交互桌面及匹配的 ui_coords.json。
手动版网页生成带书签的图片 PDF，后续手动 OCR 并用 finish-manual-ocr.ps1 验证及恢复书签。

## 验收范围

已验证禁止下载的一键检查、重复检查、中文/空格/括号目录启动、网页 Ghostscript 发现、
PDFium 渲染、Ghostscript 压缩、qpdf 优化及书签恢复。缺失组件与子检查失败会返回非零退出码。
这不是全新主机/虚拟机验收；没有运行真实 Word 导出或 Acrobat GUI 自动 OCR。

本次提交记录交付包发布说明、下载/合并工具、校验值和验收结果，未同步包内修复到项目源码。
此前工作区已有的未提交源码改动保持原状。

## SHA-256

```text
33218d08c6916afcd7e1d732d2756a8f5564ffe198e1cf971cfb946c19d18dd5  PDF-PDF-full.zip
604b0b8859a99f4ef231845d26118f73d943cd7653933e83d04557c85cfea89a  PDF-PDF-manual-ocr.zip
```

验收记录随 Release 附件提供，并在仓库同目录保存为 verification.json。
Ghostscript 许可证及对应版本源码归档随两个安装包提供。
