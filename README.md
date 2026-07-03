# PDF-PDF 双层 PDF 制作工具

这是一个本地 Web 工具，用于把 Word 文档或已导出的带书签 PDF 制作成可检索的双层 PDF。

当前主流程使用 Adobe Acrobat Pro 的“扫描和 OCR”桌面功能生成文字层，然后把原 PDF 书签恢复到最终文件。

详细安装和使用步骤请看：

[使用说明.md](./使用说明.md)

## 快速启动

第一次从 GitHub 下载项目后，先双击：

```text
一键安装组件.bat
```

安装完成后再双击：

```text
start-web.bat
```

## 日常启动

双击：

```text
start-web.bat
```

打开网页后，选择输入文件或多文件队列，再选择输出目录，点击“开始制作”。

## 主要功能

- 单个 Word / PDF 制作
- 多文件队列制作
- CSV 队列制作
- Windows 原生文件选择
- 300 DPI 图片渲染
- Acrobat Pro 自动 OCR
- 扫描完成后等待保存并验证文字层
- 恢复原 PDF 书签
- 输出日志和队列汇总
- 缓存清理和依赖检测

## 注意

直接处理 Word 文件需要安装 Microsoft Word。

生成双层 PDF 需要安装 Adobe Acrobat Pro，并保证 Acrobat 能正常打开“扫描和 OCR”。Reader 不支持本工具需要的 OCR 流程。
