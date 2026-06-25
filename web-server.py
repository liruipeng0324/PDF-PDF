from __future__ import annotations

import argparse
import json
import os
import re
import string
import subprocess
import threading
import time
import uuid
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse


ROOT = Path(__file__).resolve().parent
POWERSHELL = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "WindowsPowerShell", "v1.0", "powershell.exe")
if not os.path.exists(POWERSHELL):
    POWERSHELL = "powershell"
JOBS: dict[str, dict] = {}


STAGE_RE = re.compile(r"\[(\d+)/(\d+)\]\s*(.*)")
OCR_TOTAL_RE = re.compile(r"OCR_TOTAL_PAGES\s+(\d+)")
OCR_PAGE_RE = re.compile(r"(?:page|Page|OCR).*?(\d+)\s*/\s*(\d+)|(?:page|Page)\s+(\d+)")
PNG_TOTAL_RE = re.compile(r"PNG_TOTAL_PAGES\s+(\d+)")
PNG_PAGE_RE = re.compile(r"PNG_PAGE\s+(\d+)\s*/\s*(\d+)")


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


def run_job(job_id: str, command: list[str], cwd: Path) -> None:
    job = JOBS[job_id]
    job["status"] = "running"
    job["progressText"] = "正在运行"
    job["startedAt"] = time.strftime("%Y-%m-%d %H:%M:%S")
    try:
        process = subprocess.Popen(
            command,
            cwd=str(cwd),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        job["pid"] = process.pid
        lines = []
        assert process.stdout is not None
        for line in process.stdout:
            clean_line = line.rstrip()
            lines.append(clean_line)
            update_progress(job, clean_line)
            job["log"] = "\n".join(lines[-500:])
        code = process.wait()
        job["exitCode"] = code
        job["status"] = "success" if code == 0 else "failed"
        if code == 0:
            job["progress"] = 100
            job["progressText"] = "完成"
    except Exception as exc:  # noqa: BLE001 - return readable job error to local UI
        job["status"] = "failed"
        job["exitCode"] = -1
        job["log"] = (job.get("log", "") + "\n" + str(exc)).strip()
    finally:
        job["finishedAt"] = time.strftime("%Y-%m-%d %H:%M:%S")


def update_progress(job: dict, line: str) -> None:
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
        page_progress = int((job["pngCurrentPage"] / max(total_pages, 1)) * 20)
        job["progress"] = min(39, 20 + page_progress)
        job["progressText"] = f"PNG 正在转换第 {job['pngCurrentPage']} / {total_pages} 页"
        return

    total_match = OCR_TOTAL_RE.search(line)
    if total_match:
        job["ocrTotalPages"] = int(total_match.group(1))
        return

    match = STAGE_RE.search(line)
    if not match:
        if job.get("ocrTotalPages") and job.get("currentStage") == 4:
            page_match = OCR_PAGE_RE.search(line)
            if page_match:
                if page_match.group(1) and page_match.group(2):
                    current_page = int(page_match.group(1))
                    total_pages = int(page_match.group(2))
                else:
                    current_page = int(page_match.group(3))
                    total_pages = int(job["ocrTotalPages"])
                job["ocrCurrentPage"] = min(current_page, total_pages)
                job["ocrTotalPages"] = total_pages
                page_progress = int((job["ocrCurrentPage"] / max(total_pages, 1)) * 20)
                job["progress"] = min(79, 60 + page_progress)
                job["progressText"] = f"OCR 正在扫描第 {job['ocrCurrentPage']} / {total_pages} 页"
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
        if detail:
            job["progressText"] = f"步骤 {current}/{total}: {detail}"
        else:
            job["progressText"] = f"步骤 {current}/{total}"
        if current == 4 and job.get("ocrTotalPages"):
            total_pages = job["ocrTotalPages"]
            current_page = job.get("ocrCurrentPage", 0)
            if current_page:
                job["progressText"] = f"OCR 正在扫描第 {current_page} / {total_pages} 页"
            else:
                job["progressText"] = f"OCR 正在扫描，共 {total_pages} 页"
        if current == 2 and job.get("pngTotalPages"):
            total_pages = job["pngTotalPages"]
            current_page = job.get("pngCurrentPage", 0)
            if current_page:
                job["progressText"] = f"PNG 正在转换第 {current_page} / {total_pages} 页"
            else:
                job["progressText"] = f"PNG 正在转换，共 {total_pages} 页"
    else:
        progress = min(99, max(1, int((current / total) * 100)))
        job["progressText"] = f"队列 {current}/{total}: {detail}" if detail else f"队列 {current}/{total}"

    job["progress"] = progress


def switch_arg(enabled: bool, name: str) -> list[str]:
    return [name] if enabled else []


def build_command(payload: dict) -> list[str]:
    mode = payload.get("mode", "single")
    language = payload.get("language", "chi_sim+eng")
    dpi = str(payload.get("dpi", 300))

    base = [POWERSHELL, "-ExecutionPolicy", "Bypass"]

    if mode == "single":
        input_type = payload.get("inputType", "pdf")
        source_path = payload.get("sourcePath", "")
        output_path = payload.get("outputPath", "")
        if not source_path or not output_path:
            raise ValueError("源文件和输出 PDF 都不能为空。")
        script = ROOT / "make-dual-pdf.ps1"
        command = base + ["-File", str(script)]
        command += ["-WordPath" if input_type == "word" else "-TaggedPdfPath", source_path]
        command += ["-OutputPath", output_path]
    else:
        source_path = payload.get("sourcePath", "")
        output_path = payload.get("outputPath", "")
        if not source_path or not output_path:
            raise ValueError("输入位置和输出目录都不能为空。")
        script = ROOT / "run-queue.ps1"
        command = base + ["-File", str(script)]
        command += ["-InputDir" if mode == "queue-dir" else "-QueueCsv", source_path]
        command += ["-OutputDir", output_path]
        command += ["-InputType", payload.get("inputType", "pdf")]

    command += ["-Language", language, "-Dpi", dpi]
    command += switch_arg(bool(payload.get("deskew")), "-Deskew")
    command += switch_arg(bool(payload.get("rotatePages")), "-RotatePages")
    command += switch_arg(bool(payload.get("optimize")), "-Optimize")
    command += switch_arg(bool(payload.get("keepWork")), "-KeepWork")
    command += switch_arg(bool(payload.get("detailedPngProgress")), "-DetailedPngProgress")
    if mode == "queue-dir":
        command += switch_arg(bool(payload.get("recurse")), "-Recurse")
    if mode != "single":
        command += switch_arg(bool(payload.get("skipExisting")), "-SkipExisting")
    return command


def expected_outputs(payload: dict) -> dict:
    mode = payload.get("mode", "single")
    output_path = payload.get("outputPath", "")
    if mode == "single":
        return {"outputPath": output_path, "outputDir": str(Path(output_path).parent) if output_path else ""}
    return {"outputDir": output_path, "summaryHint": str(Path(output_path) / "_queue-logs") if output_path else ""}


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

    return {
        "pdfs": [item(path) for path in pdfs],
        "summaries": [item(path) for path in summaries],
        "logs": [item(path) for path in logs],
    }


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def log_message(self, format: str, *args) -> None:
        return

    def do_GET(self) -> None:  # noqa: N802 - stdlib hook
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
            json_response(self, 200, {"ok": True, "job": job})
            return

        if parsed.path == "/api/recent":
            try:
                json_response(self, 200, {"ok": True, **recent_outputs()})
            except Exception as exc:  # noqa: BLE001
                json_response(self, 400, {"ok": False, "error": str(exc)})
            return

        if parsed.path == "/":
            self.path = "/index.html"
        super().do_GET()

    def do_POST(self) -> None:  # noqa: N802 - stdlib hook
        parsed = urlparse(self.path)
        if parsed.path == "/api/clean-cache":
            try:
                result = clean_cache()
                json_response(self, 200, {"ok": True, **result})
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
                "exitCode": None,
                "progress": 0,
                "progressText": "等待开始",
                "startedAt": "",
                "finishedAt": "",
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
