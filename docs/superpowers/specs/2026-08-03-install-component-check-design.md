# 一键安装组件检查设计

## 目标

让 `一键安装组件.bat` 对项目组件执行可重复的预检：已满足的组件只报告并跳过，缺失的组件才下载或安装；Ghostscript 缺失时不阻断主流程。同时将项目维护为单一安装入口。

## 范围

- 保留 `install-components.ps1` 作为唯一安装入口。
- 检查项目内置 Python 3.12、Python 包、Poppler、qpdf、Ghostscript，以及 Word/Acrobat 的可用性。
- Python 包按实际导入结果逐项补装，不因每次运行都执行完整 `pip install`。
- Ghostscript 是可选组件，缺失时跳过并提示。
- 删除重复的 `install-python.ps1` 和 `install-python-packages.ps1`，同步更新文档。

## 行为设计

安装器输出每项组件的状态：`[found]`、`[missing]`、`[installing]`、`[skipped]`。Python 与外部工具通过可执行文件检查，Python 包通过目标 Python 解释器执行 import 检查。普通运行只对缺失项执行操作；`-Force` 是显式重新安装/升级入口。

主流程顺序为：建立目录 -> 检查/安装 Python -> 检查/补装 Python 包 -> 检查/下载 Poppler 与 qpdf -> 检查/可选安装 Ghostscript -> 检查桌面应用 -> 执行最终依赖检查。

## 错误处理与验证

缺失且允许下载的必需组件安装失败时立即报错；`-SkipDownloads` 下不下载并明确报告缺失。Ghostscript 缺失只报告可选能力不可用。安装后重新调用依赖检查，确保运行时可发现组件。

## 验收标准

1. 重复运行一键安装脚本不会重新下载已存在的 Python、Poppler、qpdf 或 Ghostscript。
2. 已能导入的 Python 包不会触发 pip 安装；缺失包只补装缺失项。
3. `-Force` 能显式触发 Python 包升级和组件重新安装。
4. 缺少 Ghostscript 不会使安装失败。
5. 文档不再指导用户使用已删除的两个重复安装脚本。
