from __future__ import annotations

import argparse
import csv
import json
import os
import re
import shutil
import string
import subprocess
import tempfile
import threading
import time
import uuid
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse


ROOT = Path(__file__).resolve().parent
POWERSHELL = os.path.join(
    os.environ.get("SystemRoot", r"C:\Windows"),
    "System32",
    "WindowsPowerShell",
    "v1.0",
    "powershell.exe",
)
if not os.path.exists(POWERSHELL):
    POWERSHELL = "powershell"

JOBS: dict[str, dict] = {}

STAGE_RE = re.compile(r"\[(\d+)/(\d+)\]\s*(.*)")
QUEUE_START_RE = re.compile(r"\[(\d+)/(\d+)\]\s+START\s+(.*)")
QUEUE_END_RE = re.compile(r"\[(\d+)/(\d+)\]\s+(OK|FAIL|SKIP)\s+(.*)")
PNG_TOTAL_RE = re.compile(r"(?:PNG|IMAGE)_TOTAL_PAGES\s+(\d+)")
PNG_PAGE_RE = re.compile(r"(?:PNG|IMAGE)_PAGE\s+(\d+)\s*/\s*(\d+)")
QUEUE_LOG_FALLBACK_RE = re.compile(r"QUEUE_LOG_DIR_FALLBACK\s+(.+)")


def json_response(handler: SimpleHTTPRequestHandler, status: int, payload: dict) -> None:
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(data)))
    handler.end_headers()
    handler.wfile.write(data)


def read_json(handler: SimpleHTTPRequestHandler) -> dict:
    length = int(handler.headers.get("Content-Length", "0"))
    raw = handler.rfile.read(length)
    if not raw:
        return {}
    return json.loads(raw.decode("utf-8"))


def add_tool_paths(env: dict[str, str]) -> dict[str, str]:
    paths = [
        ROOT / "tools" / "Python312",
        ROOT / "tools" / "Python312" / "Scripts",
        ROOT / "tools" / "Python313",
        ROOT / "tools" / "Python313" / "Scripts",
        ROOT / "tools" / "poppler-26.02.0-0" / "poppler-26.02.0" / "Library" / "bin",
        ROOT / "tools" / "qpdf-12.3.2" / "qpdf-12.3.2-msvc64" / "bin",
        ROOT / "tools" / "ghostscript-10.07.1" / "bin",
        Path(r"D:\GPL\gs10.07.1\bin"),
    ]
    existing = [str(path) for path in paths if path.exists()]
    env["PATH"] = ";".join(existing + [env.get("PATH", "")])
    packages = ROOT / "tools" / "python-packages"
    if packages.exists() and not (ROOT / "tools" / "Python312" / "python.exe").exists():
        env["PYTHONPATH"] = str(packages) + (";" + env["PYTHONPATH"] if env.get("PYTHONPATH") else "")
    return env


def run_short(command: list[str], timeout: int = 12) -> tuple[bool, str]:
    env = add_tool_paths(os.environ.copy())
    executable = command[0]
    if not Path(executable).exists():
        resolved = shutil.which(executable, path=env.get("PATH", ""))
        if resolved:
            command = [resolved, *command[1:]]
    try:
        result = subprocess.run(
            command,
            cwd=str(ROOT),
            env=env,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
        )
        output = (result.stdout or "").strip()
        return result.returncode == 0, output
    except Exception as exc:  # noqa: BLE001
        return False, str(exc)


def find_python() -> str:
    candidates = [
        ROOT / "tools" / "Python312" / "python.exe",
        ROOT / "tools" / "Python313" / "python.exe",
        Path(r"C:\Users\Administrator\AppData\Local\Python\bin\python.exe"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return str(candidate)
    for name in ("py", "python"):
        path = shutil.which(name)
        if path and r"\Microsoft\WindowsApps\\" not in path:
            return path
    return ""


def drive_roots() -> list[dict]:
    roots = []
    for letter in string.ascii_uppercase:
        path = f"{letter}:\\"
        if os.path.exists(path):
            roots.append({"name": path, "path": path, "type": "dir"})
    return roots


def list_directory(path: str) -> dict:
    if not path:
        return {"path": "", "parent": "", "items": drive_roots()}

    current = Path(path)
    if not current.exists():
        raise FileNotFoundError(path)
    if not current.is_dir():
        raise NotADirectoryError(path)

    items = []
    try:
        children = sorted(current.iterdir(), key=lambda p: (not p.is_dir(), p.name.lower()))
    except PermissionError:
        children = []

    for child in children:
        try:
            is_dir = child.is_dir()
        except OSError:
            continue
        if is_dir:
            item_type = "dir"
        else:
            ext = child.suffix.lower()
            item_type = "file" if ext in {".pdf", ".doc", ".docx", ".csv"} else "other"
        if item_type != "other":
            items.append({"name": child.name, "path": str(child), "type": item_type})

    parent = str(current.parent) if current.parent != current else ""
    return {"path": str(current), "parent": parent, "items": items}


def explain_error(log: str) -> str:
    patterns = [
        ("Word PDF export failed", "Word 已安装，但自动导出 PDF 失败。请先打开 Word 关闭激活、恢复、更新或受保护视图提示；如果仍失败，可先用 Word 手动另存为 PDF，再选择 PDF 输入模式。"),
        ("No JPG pages were created", "PDF 转图片没有生成页面。请确认源 PDF 能正常打开。"),
        ("No PNG pages were created", "PDF 转图片没有生成页面。请确认源 PDF 能正常打开。"),
        ("No rendered image pages were created", "PDF 转图片没有生成页面。请确认源 PDF 能正常打开。"),
        ("pdftoppm failed", "PDF 转图片失败。请检查 PDF 是否损坏/加密。"),
        ("Access is denied", "路径没有写入权限，或目标 PDF 正被其他程序打开。"),
        ("Permission denied", "路径没有写入权限，或目标 PDF 正被其他程序打开。"),
        ("Queue is empty", "队列为空。请确认已经选择 Word/PDF 文件，或 CSV 内容正确。"),
    ]
    for marker, message in patterns:
        if marker.lower() in log.lower():
            return message
    return ""


def quote_command(command: list[str]) -> str:
    parts = []
    for item in command:
        if any(ch.isspace() for ch in item) or '"' in item:
            parts.append('"' + item.replace('"', '\\"') + '"')
        else:
            parts.append(item)
    return " ".join(parts)


def append_log_file(job: dict, line: str) -> None:
    log_path = (job.get("outputs") or {}).get("summaryHint")
    if not log_path:
        return
    try:
        path = Path(log_path)
        if path.suffix.lower() != ".log":
            path.mkdir(parents=True, exist_ok=True)
            path = path / "web-job.log"
        else:
            path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("a", encoding="utf-8") as handle:
            handle.write(line + "\n")
    except Exception:
        return


def run_job(job_id: str, command: list[str], cwd: Path) -> None:
    job = JOBS[job_id]
    job["status"] = "running"
    job["progressText"] = "正在运行"
    job["startedAt"] = time.strftime("%Y-%m-%d %H:%M:%S")
    job["lastOutputAt"] = time.time()
    lines = [
        "WEB_JOB_STARTED " + job["startedAt"],
        "WEB_COMMAND " + quote_command(command),
    ]
    job["log"] = "\n".join(lines)
    for initial_line in lines:
        append_log_file(job, initial_line)
    try:
        process = subprocess.Popen(
            command,
            cwd=str(cwd),
            env=add_tool_paths(os.environ.copy()),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        job["pid"] = process.pid
        job["process"] = process
        monitor = threading.Thread(target=monitor_job_stall, args=(job_id,), daemon=True)
        monitor.start()
        assert process.stdout is not None
        for line in process.stdout:
            clean_line = line.rstrip()
            job["lastOutputAt"] = time.time()
            lines.append(clean_line)
            append_log_file(job, clean_line)
            update_progress(job, clean_line)
            text = "\n".join(lines[-700:])
            job["log"] = text
            job["friendlyError"] = explain_error(text)
            if job.get("cancelRequested"):
                break
        if job.get("cancelRequested") and process.poll() is None:
            terminate_process_tree(process.pid)
        code = process.wait()
        job["exitCode"] = code
        if job.get("cancelRequested"):
            job["status"] = "cancelled"
            job["progressText"] = "已取消"
        else:
            job["status"] = "success" if code == 0 else "failed"
            if code == 0:
                job["progress"] = 100
                job["progressText"] = "完成"
    except Exception as exc:  # noqa: BLE001
        job["status"] = "failed"
        job["exitCode"] = -1
        job["log"] = (job.get("log", "") + "\n" + str(exc)).strip()
        job["friendlyError"] = explain_error(job["log"]) or str(exc)
    finally:
        job.pop("process", None)
        job["finishedAt"] = time.strftime("%Y-%m-%d %H:%M:%S")
        append_log_file(job, "WEB_JOB_FINISHED " + job["finishedAt"])


def terminate_process_tree(pid: int) -> None:
    if not pid:
        return
    try:
        subprocess.run(
            ["taskkill", "/PID", str(pid), "/T", "/F"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=15,
        )
    except Exception:
        try:
            subprocess.run(
                [POWERSHELL, "-NoProfile", "-Command", f"Stop-Process -Id {int(pid)} -Force -ErrorAction SilentlyContinue"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=10,
            )
        except Exception:
            return


def monitor_job_stall(job_id: str) -> None:
    while True:
        time.sleep(15)
        job = JOBS.get(job_id)
        if not job or job.get("status") not in {"running", "queued"}:
            return
        process = job.get("process")
        if not process or process.poll() is not None:
            return

        idle_seconds = time.time() - float(job.get("lastOutputAt") or time.time())
        if idle_seconds < 180:
            continue

        idle_minutes = int(idle_seconds // 60)
        job["progressText"] = f"任务仍在运行，已 {idle_minutes} 分钟没有新日志"


def update_progress(job: dict, line: str) -> None:
    fallback_match = QUEUE_LOG_FALLBACK_RE.search(line)
    if fallback_match:
        fallback_dir = fallback_match.group(1).strip()
        outputs = job.setdefault("outputs", {})
        outputs["summaryHint"] = fallback_dir
        outputs["controlDir"] = str(Path(fallback_dir) / "control")
        return

    queue_start_match = QUEUE_START_RE.search(line)
    if queue_start_match:
        current = int(queue_start_match.group(1))
        total = max(int(queue_start_match.group(2)), 1)
        job["queueCurrent"] = current
        job["queueTotal"] = total
        job["progress"] = max(1, int(((current - 1) / total) * 100))
        job["progressText"] = f"队列 {current}/{total}: 开始处理"
        return

    queue_end_match = QUEUE_END_RE.search(line)
    if queue_end_match:
        current = int(queue_end_match.group(1))
        total = max(int(queue_end_match.group(2)), 1)
        status = queue_end_match.group(3)
        job["queueCurrent"] = current
        job["queueTotal"] = total
        job["progress"] = min(99, int((current / total) * 100))
        status_text = {"OK": "完成", "FAIL": "失败", "SKIP": "跳过"}.get(status, status)
        job["progressText"] = f"队列 {current}/{total}: {status_text}"
        return

    png_total_match = PNG_TOTAL_RE.search(line)
    if png_total_match:
        job["pngTotalPages"] = int(png_total_match.group(1))
        return

    png_page_match = PNG_PAGE_RE.search(line)
    if png_page_match:
        current_page = int(png_page_match.group(1))
        total_pages = int(png_page_match.group(2))
        job["pngCurrentPage"] = min(current_page, total_pages)
        job["pngTotalPages"] = total_pages
        update_page_progress(job, 0.20, 0.20, job["pngCurrentPage"], total_pages)
        job["progressText"] = f"图片正在转换第 {job['pngCurrentPage']} / {total_pages} 页"
        return

    match = STAGE_RE.search(line)
    if not match:
        if "Queue complete." in line:
            job["progress"] = 100
            job["progressText"] = "队列完成"
        return

    current = int(match.group(1))
    total = max(int(match.group(2)), 1)
    detail = match.group(3).strip()

    if total == 5:
        job["currentStage"] = current
        progress = min(99, max(1, int(((current - 1) / total) * 100)))
        job["progressText"] = f"步骤 {current}/{total}: {detail}" if detail else f"步骤 {current}/{total}"
        if current == 2 and job.get("pngTotalPages"):
            total_pages = job["pngTotalPages"]
            current_page = job.get("pngCurrentPage", 0)
            job["progressText"] = (
                f"图片正在转换第 {current_page} / {total_pages} 页"
                if current_page
                else f"图片正在转换，共 {total_pages} 页"
            )
    else:
        progress = min(99, max(1, int((current / total) * 100)))
        job["progressText"] = f"队列 {current}/{total}: {detail}" if detail else f"队列 {current}/{total}"

    job["progress"] = progress


def update_page_progress(job: dict, stage_start: float, stage_span: float, current_page: int, total_pages: int) -> None:
    queue_current = job.get("queueCurrent")
    queue_total = job.get("queueTotal")
    page_ratio = current_page / max(total_pages, 1)
    item_ratio = min(0.99, stage_start + stage_span * page_ratio)
    if queue_current and queue_total:
        overall = ((queue_current - 1) + item_ratio) / max(queue_total, 1)
        job["progress"] = min(99, max(1, int(overall * 100)))
    else:
        job["progress"] = min(99, max(1, int(item_ratio * 100)))


def switch_arg(enabled: bool, name: str) -> list[str]:
    return [name] if enabled else []


def build_command(payload: dict) -> list[str]:
    mode = payload.get("mode", "single")
    dpi = str(payload.get("dpi", 300))
    render_engine = payload.get("renderEngine", "pdfium")
    optimize = bool(payload.get("optimize"))
    detailed_png = bool(payload.get("detailedPngProgress"))
    acrobat_double_layer = bool(payload.get("acrobatDoubleLayer"))
    acrobat_ocr_language = payload.get("acrobatOcrLanguage", "CHS")

    base = [POWERSHELL, "-NoProfile", "-STA", "-ExecutionPolicy", "Bypass"]

    if mode == "single":
        input_type = payload.get("inputType", "pdf")
        source_path = payload.get("sourcePath", "")
        output_path = normalize_single_output_path(source_path, payload.get("outputPath", ""))
        if not source_path or not output_path:
            raise ValueError("源文件和输出 PDF 都不能为空。")
        script = ROOT / "make-dual-pdf.ps1"
        command = base + ["-File", str(script)]
        command += ["-WordPath" if input_type == "word" else "-TaggedPdfPath", source_path]
        command += ["-OutputPath", output_path]
    else:
        source_path = payload.get("sourcePath", "")
        source_paths = payload.get("sourcePaths") or []
        output_path = payload.get("outputPath", "")
        if not (source_path or source_paths) or not output_path:
            raise ValueError("输入位置和输出目录都不能为空。")
        script = ROOT / "run-queue.ps1"
        command = base + ["-File", str(script)]
        if mode == "queue-dir" and source_paths:
            source_path = create_queue_csv(source_paths, payload.get("inputType", "pdf"), output_path)
            command += ["-QueueCsv", source_path]
        else:
            command += ["-InputDir" if mode == "queue-dir" else "-QueueCsv", source_path]
        command += ["-OutputDir", output_path]
        command += ["-InputType", payload.get("inputType", "pdf")]

    command += ["-Dpi", dpi]
    command += ["-RenderEngine", render_engine]
    command += switch_arg(optimize, "-Optimize")
    if acrobat_double_layer:
        command += ["-AcrobatDoubleLayer", "-AcrobatOcrLanguage", acrobat_ocr_language]
    command += switch_arg(bool(payload.get("keepWork")), "-KeepWork")
    command += switch_arg(detailed_png, "-DetailedPngProgress")
    if mode == "queue-dir":
        command += switch_arg(bool(payload.get("recurse")), "-Recurse")
    if mode != "single":
        command += switch_arg(bool(payload.get("skipExisting")), "-SkipExisting")
    return command


def create_queue_csv(source_paths: list[str], input_type: str, output_dir: str) -> str:
    queue_dir = ROOT / ".cache" / "web-queues"
    queue_dir.mkdir(parents=True, exist_ok=True)
    queue_path = queue_dir / ("queue-" + time.strftime("%Y%m%d-%H%M%S") + "-" + uuid.uuid4().hex[:8] + ".csv")
    normalized_type = input_type if input_type in {"word", "pdf", "auto"} else "auto"
    with queue_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["InputPath", "OutputPath", "Type"])
        writer.writeheader()
        for source_path in source_paths:
            if not source_path:
                continue
            writer.writerow({
                "InputPath": source_path,
                "OutputPath": "",
                "Type": normalized_type,
            })
    return str(queue_path)


def normalize_single_output_path(source_path: str, output_path: str) -> str:
    if not output_path:
        return ""
    output = Path(output_path)
    if output.suffix.lower() == ".pdf":
        return str(output)
    source_stem = Path(source_path).stem or "output"
    return str(output / f"{source_stem}-dual-layer.pdf")


def expected_outputs(payload: dict) -> dict:
    mode = payload.get("mode", "single")
    output_path = payload.get("outputPath", "")
    if mode == "single":
        output_path = normalize_single_output_path(payload.get("sourcePath", ""), output_path)
        output = Path(output_path) if output_path else Path("")
        return {
            "outputPath": output_path,
            "outputDir": str(output.parent) if output_path else "",
            "summaryHint": str(output.with_suffix(".log")) if output_path else "",
        }
    return {
        "outputDir": output_path,
        "summaryHint": str(Path(output_path) / "_queue-logs") if output_path else "",
        "controlDir": str(Path(output_path) / "_queue-logs" / "control") if output_path else "",
    }


def clean_cache() -> dict:
    removed = []
    for path in ROOT.rglob("__pycache__"):
        if not path.is_dir():
            continue
        for child in sorted(path.rglob("*"), reverse=True):
            if child.is_file():
                child.unlink(missing_ok=True)
            elif child.is_dir():
                child.rmdir()
        path.rmdir()
        removed.append(str(path))
    return {"removed": removed, "count": len(removed)}


def recent_outputs() -> dict:
    pdfs = sorted(ROOT.rglob("*dual-layer*.pdf"), key=lambda p: p.stat().st_mtime, reverse=True)[:20]
    summaries = sorted(ROOT.rglob("summary-*.csv"), key=lambda p: p.stat().st_mtime, reverse=True)[:10]
    logs = sorted(ROOT.rglob("item-*.log"), key=lambda p: p.stat().st_mtime, reverse=True)[:10]

    def item(path: Path) -> dict:
        stat = path.stat()
        return {
            "path": str(path),
            "size": stat.st_size,
            "modified": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(stat.st_mtime)),
        }

    return {"pdfs": [item(path) for path in pdfs], "summaries": [item(path) for path in summaries], "logs": [item(path) for path in logs]}


def dependency_checks() -> dict:
    python = find_python()
    checks = []

    def add(name: str, ok: bool, detail: str, fix: str = "") -> None:
        checks.append({"name": name, "ok": ok, "detail": detail.splitlines()[0] if detail else "", "fix": fix})

    add("Python", bool(python), python or "未找到 Python", "安装 Python，或放入 tools\\Python313。")
    if python:
        ok, output = run_short([python, "--version"])
        add("Python 版本", ok, output, "确认 Python 可以正常启动。")
        for module in ("img2pdf", "pikepdf", "pypdfium2", "win32com.client", "pywinauto"):
            command = [python, "-m", module, "--version"]
            if module == "pikepdf":
                command = [python, "-c", "import pikepdf; print(pikepdf.__version__)"]
            elif module == "pypdfium2":
                command = [python, "-c", "import pypdfium2; print('available')"]
            elif module == "win32com.client":
                command = [python, "-c", "import win32com.client; print('available')"]
            elif module == "pywinauto":
                command = [python, "-c", "import pywinauto; print(pywinauto.__version__)"]
            ok, output = run_short(command)
            label = "pywin32" if module == "win32com.client" else module
            add(f"Python 模块 {label}", ok, output, "运行 install-python-packages.ps1 安装依赖。")

    for name, command, args, fix in [
        ("Poppler pdftoppm", "pdftoppm", ["-v"], "安装 Poppler，或放入 tools\\poppler-26.02.0-0。"),
        ("qpdf", "qpdf", ["--version"], "安装 qpdf，或放入 tools\\qpdf-12.3.2。"),
        ("Ghostscript", "gswin64c", ["--version"], "安装 Ghostscript，或放入 tools\\ghostscript-10.07.1。"),
    ]:
        ok, output = run_short([command, *args])
        add(name, ok, output, fix)

    ok, output = run_short([POWERSHELL, "-ExecutionPolicy", "Bypass", "-File", str(ROOT / "check-word.ps1")], timeout=20)
    add("Microsoft Word 自动导出", ok and ("available" in output.lower() or "[found]" in output.lower()), output, "安装 Microsoft Word；没有 Word 时请选择“已导出的 PDF”模式。")

    ok, output = run_short([POWERSHELL, "-ExecutionPolicy", "Bypass", "-File", str(ROOT / "check-acrobat-com.ps1")], timeout=50)
    add("Acrobat COM 自动化", ok and "ACROBAT_COM_OK" in output, output, "运行 repair-acrobat-com.ps1 修复 Acrobat COM，或在桌面完成 Acrobat 授权/更新提示。")

    return {"checks": checks, "allOk": all(item["ok"] for item in checks)}


def cancel_job(job_id: str) -> dict:
    job = JOBS.get(job_id)
    if not job:
        raise ValueError("任务不存在。")
    job["cancelRequested"] = True
    process = job.get("process")
    if process and process.poll() is None:
        terminate_process_tree(process.pid)
    return {"status": job.get("status", "unknown")}


def control_job(job_id: str, action: str) -> dict:
    job = JOBS.get(job_id)
    if not job:
        raise ValueError("任务不存在。")
    control_dir = (job.get("outputs") or {}).get("controlDir")
    if not control_dir:
        raise ValueError("当前任务不是队列任务，不能暂停或跳过。")
    path = Path(control_dir)
    path.mkdir(parents=True, exist_ok=True)
    pause_file = path / "pause.flag"
    skip_file = path / "skip-next.flag"
    if action == "pause":
        pause_file.write_text("pause", encoding="utf-8")
        job["queueControl"] = "paused"
    elif action == "resume":
        pause_file.unlink(missing_ok=True)
        job["queueControl"] = "running"
    elif action == "skip":
        skip_file.write_text("skip", encoding="utf-8")
        job["queueControl"] = "skip-next"
    else:
        raise ValueError("未知队列控制命令。")
    return {"queueControl": job.get("queueControl", "")}


def open_path(path: str) -> None:
    if not path:
        raise ValueError("路径不能为空。")
    target = Path(path)
    if target.is_file():
        subprocess.Popen(["explorer.exe", "/select,", str(target)])
    else:
        subprocess.Popen(["explorer.exe", str(target)])


def native_pick(payload: dict) -> dict:
    kind = payload.get("kind", "file")
    original_kind = kind
    target = payload.get("target", "")
    title = payload.get("title", "选择路径")
    file_filter = payload.get("filter", "所有文件 (*.*)|*.*")
    initial = payload.get("initial") or payload.get("initialPath", "")
    title_text = str(title)
    if target == "output" or (kind in {"save-file", "file"} and ("输出" in title_text or "output" in title_text.lower())):
        kind = "dir"
    if target != "output" and original_kind != "save-file" and kind == "dir" and any(token in file_filter.lower() for token in (".doc", ".docx", ".pdf")):
        kind = "files"
    script = r'''
param(
    [string]$Kind,
    [string]$Title,
    [string]$Filter,
    [string]$Initial,
    [string]$ResultPath
)

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$owner = New-Object System.Windows.Forms.Form
$owner.TopMost = $true
$owner.ShowInTaskbar = $false
$owner.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$owner.Width = 1
$owner.Height = 1
$owner.Opacity = 0
$owner.Show()
$owner.Activate()

if ($Kind -eq "dir") {
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Title
    $dialog.ShowNewFolderButton = $true
    if ($Initial -and (Test-Path -LiteralPath $Initial)) {
        if ((Get-Item -LiteralPath $Initial).PSIsContainer) {
            $dialog.SelectedPath = $Initial
        }
        else {
            $dialog.SelectedPath = Split-Path -Parent $Initial
        }
    }
    if ($dialog.ShowDialog($owner) -eq [System.Windows.Forms.DialogResult]::OK) {
        Set-Content -LiteralPath $ResultPath -Value $dialog.SelectedPath -Encoding UTF8
    }
    $owner.Close()
    exit 0
}

if ($Kind -eq "save-file") {
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Title = $Title
    $dialog.Filter = $Filter
    $dialog.OverwritePrompt = $true
    $dialog.AddExtension = $true
    $dialog.DefaultExt = "pdf"
    if ($Initial) {
        if (Test-Path -LiteralPath $Initial) {
            if ((Get-Item -LiteralPath $Initial).PSIsContainer) {
                $dialog.InitialDirectory = $Initial
            }
            else {
                $dialog.InitialDirectory = Split-Path -Parent $Initial
                $dialog.FileName = Split-Path -Leaf $Initial
            }
        }
        else {
            $dialog.InitialDirectory = Split-Path -Parent $Initial
            $dialog.FileName = Split-Path -Leaf $Initial
        }
    }
    if ($dialog.ShowDialog($owner) -eq [System.Windows.Forms.DialogResult]::OK) {
        Set-Content -LiteralPath $ResultPath -Value $dialog.FileName -Encoding UTF8
    }
    $owner.Close()
    exit 0
}

$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Title = $Title
$dialog.Filter = $Filter
$dialog.CheckFileExists = $true
$dialog.Multiselect = ($Kind -eq "files")
if ($Initial -and (Test-Path -LiteralPath $Initial)) {
    if ((Get-Item -LiteralPath $Initial).PSIsContainer) {
        $dialog.InitialDirectory = $Initial
    }
    else {
        $dialog.InitialDirectory = Split-Path -Parent $Initial
    }
}
if ($dialog.ShowDialog($owner) -eq [System.Windows.Forms.DialogResult]::OK) {
    if ($Kind -eq "files") {
        Set-Content -LiteralPath $ResultPath -Value $dialog.FileNames -Encoding UTF8
    }
    else {
        Set-Content -LiteralPath $ResultPath -Value $dialog.FileName -Encoding UTF8
    }
}
$owner.Close()
'''
    temp_path = ""
    result_path = ""
    try:
        with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False, encoding="utf-8") as handle:
            handle.write(script)
            temp_path = handle.name
        result_handle = tempfile.NamedTemporaryFile("w", suffix=".txt", delete=False, encoding="utf-8")
        result_path = result_handle.name
        result_handle.close()
        Path(result_path).write_text("", encoding="utf-8")

        startupinfo = None
        creationflags = 0
        if os.name == "nt":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            creationflags = subprocess.CREATE_NO_WINDOW

        process = subprocess.Popen(
            [POWERSHELL, "-NoProfile", "-WindowStyle", "Hidden", "-STA", "-ExecutionPolicy", "Bypass", "-File", temp_path, kind, title, file_filter, initial, result_path],
            cwd=str(ROOT),
            startupinfo=startupinfo,
            creationflags=creationflags,
        )
        code = process.wait(timeout=600)
        if code != 0:
            raise RuntimeError("系统文件资源管理器选择窗口打开失败。")
        selected = Path(result_path).read_text(encoding="utf-8-sig").strip().splitlines()
        return {"selected": selected[-1] if selected else "", "selectedPaths": selected}
    finally:
        if temp_path:
            Path(temp_path).unlink(missing_ok=True)
        if result_path:
            Path(result_path).unlink(missing_ok=True)


def native_pick_tk(kind: str, title: str, file_filter: str, initial: str) -> dict:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception:
        return {"selected": ""}

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    root.update()

    initial_path = Path(initial) if initial else None
    initial_dir = ""
    if initial_path and initial_path.exists():
        initial_dir = str(initial_path if initial_path.is_dir() else initial_path.parent)

    try:
        if kind == "dir":
            selected = filedialog.askdirectory(title=title, initialdir=initial_dir or None, mustexist=False)
        else:
            filetypes = [("所有文件", "*.*")]
            if "PDF" in file_filter:
                filetypes = [("PDF 文件", "*.pdf"), ("所有文件", "*.*")]
            elif "Word" in file_filter:
                filetypes = [("Word 文档", "*.doc *.docx"), ("所有文件", "*.*")]
            elif "CSV" in file_filter:
                filetypes = [("CSV 队列", "*.csv"), ("所有文件", "*.*")]
            selected = filedialog.askopenfilename(title=title, initialdir=initial_dir or None, filetypes=filetypes)
        return {"selected": selected or ""}
    finally:
        root.destroy()


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def end_headers(self) -> None:
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        super().end_headers()

    def log_message(self, format: str, *args) -> None:
        return

    def do_GET(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        if parsed.path == "/api/list":
            query = parse_qs(parsed.query)
            path = query.get("path", [""])[0]
            try:
                json_response(self, 200, {"ok": True, **list_directory(path)})
            except Exception as exc:  # noqa: BLE001
                json_response(self, 400, {"ok": False, "error": str(exc)})
            return

        if parsed.path == "/api/job":
            query = parse_qs(parsed.query)
            job_id = query.get("id", [""])[0]
            job = JOBS.get(job_id)
            if not job:
                json_response(self, 404, {"ok": False, "error": "任务不存在。"})
                return
            visible = {key: value for key, value in job.items() if key != "process"}
            json_response(self, 200, {"ok": True, "job": visible})
            return

        if parsed.path == "/api/recent":
            try:
                json_response(self, 200, {"ok": True, **recent_outputs()})
            except Exception as exc:  # noqa: BLE001
                json_response(self, 400, {"ok": False, "error": str(exc)})
            return

        if parsed.path == "/api/tools":
            json_response(self, 200, {"ok": True, **dependency_checks()})
            return

        if parsed.path == "/":
            self.path = "/index.html"
        super().do_GET()

    def do_POST(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        if parsed.path == "/api/clean-cache":
            try:
                result = clean_cache()
                json_response(self, 200, {"ok": True, **result})
            except Exception as exc:  # noqa: BLE001
                json_response(self, 400, {"ok": False, "error": str(exc)})
            return

        if parsed.path == "/api/cancel":
            try:
                payload = read_json(self)
                json_response(self, 200, {"ok": True, **cancel_job(payload.get("jobId", ""))})
            except Exception as exc:  # noqa: BLE001
                json_response(self, 400, {"ok": False, "error": str(exc)})
            return

        if parsed.path == "/api/control":
            try:
                payload = read_json(self)
                json_response(self, 200, {"ok": True, **control_job(payload.get("jobId", ""), payload.get("action", ""))})
            except Exception as exc:  # noqa: BLE001
                json_response(self, 400, {"ok": False, "error": str(exc)})
            return

        if parsed.path == "/api/open-path":
            try:
                payload = read_json(self)
                open_path(payload.get("path", ""))
                json_response(self, 200, {"ok": True})
            except Exception as exc:  # noqa: BLE001
                json_response(self, 400, {"ok": False, "error": str(exc)})
            return

        if parsed.path == "/api/pick":
            try:
                payload = read_json(self)
                json_response(self, 200, {"ok": True, **native_pick(payload)})
            except Exception as exc:  # noqa: BLE001
                json_response(self, 400, {"ok": False, "error": str(exc)})
            return

        if parsed.path != "/api/start":
            json_response(self, 404, {"ok": False, "error": "接口不存在。"})
            return

        try:
            payload = read_json(self)
            command = build_command(payload)
            job_id = uuid.uuid4().hex
            JOBS[job_id] = {
                "id": job_id,
                "status": "queued",
                "command": command,
                "outputs": expected_outputs(payload),
                "log": "",
                "friendlyError": "",
                "exitCode": None,
                "progress": 0,
                "progressText": "等待开始",
                "startedAt": "",
                "finishedAt": "",
                "cancelRequested": False,
            }
            thread = threading.Thread(target=run_job, args=(job_id, command, ROOT), daemon=True)
            thread.start()
            json_response(self, 200, {"ok": True, "jobId": job_id})
        except Exception as exc:  # noqa: BLE001
            json_response(self, 400, {"ok": False, "error": str(exc)})


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    args = parser.parse_args()

    server = ThreadingHTTPServer((args.host, args.port), Handler)
    print(f"Dual-layer PDF web UI: http://{args.host}:{args.port}/")
    server.serve_forever()


if __name__ == "__main__":
    main()
