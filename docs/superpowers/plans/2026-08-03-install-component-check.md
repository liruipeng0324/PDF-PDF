# 一键安装组件检查 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将一键安装脚本改为可重复执行、只安装缺失组件的唯一安装入口。

**Architecture:** 在现有 `install-components.ps1` 中增加统一的 Python 包检查和外部工具查找逻辑，保留现有下载缓存机制；安装流程只对检查失败的组件执行动作，Ghostscript 作为可选组件处理。删除两个重复的独立 Python 安装脚本，并同步调整文档。

**Tech Stack:** PowerShell 5+/Windows, Python/pip, Git, Markdown.

## Global Constraints

- 保留项目内置 Python 3.12 路径 `tools\Python312`。
- 必需 Python 包来自 `requirements.txt`：`pikepdf`、`img2pdf`、`pypdfium2`、`pywin32`、`pywinauto`。
- Poppler、qpdf 是必需外部工具；Ghostscript 是可选外部工具。
- 普通运行不重复下载或安装；仅 `-Force` 才允许强制重装/升级。
- 保留用户现有未提交修改，不重置或覆盖无关文件。

---

### Task 1: 建立安装器行为测试脚本

**Files:**
- Create: `tests/test-install-components.ps1`

**Interfaces:**
- Tests the script text and isolated helper behavior through a temporary fake tool tree; no network or real installation is used.

- [ ] **Step 1: Write the failing tests**

```powershell
$script = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot '..\install-components.ps1')
if ($script -notmatch 'Test-PythonModule') { throw 'missing Python module check' }
if ($script -notmatch 'Install-PythonPackage') { throw 'missing per-package install helper' }
if ($script -notmatch 'Ghostscript.*optional|optional.*Ghostscript') { throw 'missing optional Ghostscript behavior' }
if ($script -match 'install-python-packages\.ps1') { throw 'installer still depends on duplicate package script' }
Write-Host 'PASS installer structure checks'
```

- [ ] **Step 2: Run test to verify it fails**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\test-install-components.ps1`

Expected: FAIL because the current installer has no per-package helper and no explicit optional Ghostscript status.

- [ ] **Step 3: Keep the test as a regression check**

The test must remain network-free and assert only the installer contract, so it can run in CI or on a machine without the optional desktop applications.

### Task 2: Implement component preflight and selective installation

**Files:**
- Modify: `install-components.ps1`

**Interfaces:**
- Add `Find-ToolExecutable([string[]]$Names)` returning the first callable executable path.
- Add `Test-PythonModule([string]$PythonExe, [string]$Module)` returning a Boolean import result.
- Add `Install-PythonPackage([string]$PythonExe, [string]$Requirement, [string]$Module)` that skips on success and invokes pip only when missing or `-Force` is set.

- [ ] **Step 1: Write the failing test**

Extend `tests/test-install-components.ps1` with checks that the script contains `Test-PythonModule`, checks `pypdfium2`, and reports `Ghostscript` as optional.

- [ ] **Step 2: Run the test and confirm the expected failure**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\test-install-components.ps1`

Expected: FAIL on the missing helper/behavior assertions.

- [ ] **Step 3: Implement the minimal selective behavior**

Use the target portable Python for every package check. Check each required module before pip, install only missing requirements, and preserve the existing download cache. Replace unconditional package installation with the per-package loop. Make Ghostscript print `[skipped]` when absent and continue.

- [ ] **Step 4: Run the test to verify it passes**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\test-install-components.ps1`

Expected: `PASS installer structure checks` and exit code 0.

### Task 3: Remove duplicate entry points

**Files:**
- Delete: `install-python.ps1`
- Delete: `install-python-packages.ps1`
- Modify: `README.md`
- Modify: `使用说明.md`

**Interfaces:**
- Documentation directs all installation and repair actions to `一键安装组件.bat` or `install-components.ps1`.

- [ ] **Step 1: Add documentation assertions**

Extend the test to fail if either deleted script is referenced by the two Markdown documents.

- [ ] **Step 2: Run the test and confirm it fails before cleanup**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\test-install-components.ps1`

Expected: FAIL while references or files remain.

- [ ] **Step 3: Delete the duplicate scripts and update instructions**

Explain that rerunning the one-click installer performs a component check and installs only missing items; document `-Force` and the optional Ghostscript behavior. Remove manual instructions that invoke the deleted scripts.

- [ ] **Step 4: Run the regression test**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\test-install-components.ps1`

Expected: exit code 0.

### Task 4: Verify PowerShell syntax and repository requirements

**Files:**
- Verify: `install-components.ps1`, `tests/test-install-components.ps1`, `README.md`, `使用说明.md`

- [ ] **Step 1: Run PowerShell parser checks**

Run: `powershell.exe -NoProfile -Command "$files=@('.\install-components.ps1','.\tests\test-install-components.ps1'); foreach($f in $files){[System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path $f),[ref]$null,[ref]$null) | Out-Null; Write-Host ('parsed ' + $f)}"`

Expected: both files report `parsed` and no parser errors.

- [ ] **Step 2: Run the full installer contract test**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\test-install-components.ps1`

Expected: exit code 0.

- [ ] **Step 3: Review the diff and status**

Run: `git diff --check; git status --short`

Expected: no whitespace errors; only scoped installer, test, documentation, and design/plan changes appear alongside pre-existing user files.
