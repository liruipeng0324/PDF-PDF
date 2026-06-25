const fields = {
  mode: document.querySelector("#mode"),
  inputType: document.querySelector("#inputType"),
  sourcePath: document.querySelector("#sourcePath"),
  outputPath: document.querySelector("#outputPath"),
  language: document.querySelector("#language"),
  dpi: document.querySelector("#dpi"),
  deskew: document.querySelector("#deskew"),
  rotatePages: document.querySelector("#rotatePages"),
  optimize: document.querySelector("#optimize"),
  keepWork: document.querySelector("#keepWork"),
  detailedPngProgress: document.querySelector("#detailedPngProgress"),
  recurse: document.querySelector("#recurse"),
  skipExisting: document.querySelector("#skipExisting"),
};

const sourceLabel = document.querySelector("#sourceLabel");
const outputLabel = document.querySelector("#outputLabel");
const pickSource = document.querySelector("#pickSource");
const pickOutput = document.querySelector("#pickOutput");
const startJob = document.querySelector("#startJob");
const cleanCache = document.querySelector("#cleanCache");
const showRecent = document.querySelector("#showRecent");
const jobState = document.querySelector("#jobState");
const logBox = document.querySelector("#logBox");
const progressText = document.querySelector("#progressText");
const progressPercent = document.querySelector("#progressPercent");
const progressBar = document.querySelector("#progressBar");
const dialog = document.querySelector("#fileDialog");
const dialogTitle = document.querySelector("#dialogTitle");
const dialogPath = document.querySelector("#dialogPath");
const fileList = document.querySelector("#fileList");
const closeDialog = document.querySelector("#closeDialog");
const goParent = document.querySelector("#goParent");
const selectCurrent = document.querySelector("#selectCurrent");

let pickerTarget = "source";
let pickerKind = "file";
let currentPath = "";
let pollTimer = null;

function modeInfo() {
  const mode = fields.mode.value;
  const isWord = fields.inputType.value === "word";
  return { mode, isWord };
}

function updateLabels() {
  const { mode, isWord } = modeInfo();
  sourceLabel.textContent = mode === "queue-dir"
    ? "输入文件夹"
    : mode === "queue-csv"
      ? "队列 CSV"
      : isWord
        ? "Word 文件"
        : "带书签 PDF";
  outputLabel.textContent = mode === "single" ? "输出 PDF" : "输出文件夹";
  fields.sourcePath.placeholder = "点击选择";
  fields.outputPath.placeholder = "点击选择";
  selectCurrent.hidden = pickerKind !== "dir";
}

async function apiGet(url) {
  const response = await fetch(url);
  const data = await response.json();
  if (!data.ok) throw new Error(data.error || "请求失败。");
  return data;
}

async function apiPost(url, payload) {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const data = await response.json();
  if (!data.ok) throw new Error(data.error || "请求失败。");
  return data;
}

function allowedFile(item) {
  const { mode, isWord } = modeInfo();
  const name = item.name.toLowerCase();
  if (item.type === "dir") return true;
  if (pickerKind === "dir") return false;
  if (mode === "queue-csv") return name.endsWith(".csv");
  if (isWord) return name.endsWith(".doc") || name.endsWith(".docx");
  return name.endsWith(".pdf");
}

function defaultOutputFile(folderPath) {
  const source = fields.sourcePath.value || "final.pdf";
  const normalized = source.replaceAll("\\", "/");
  const filename = normalized.split("/").pop() || "final.pdf";
  const stem = filename.replace(/\.[^.]+$/, "");
  const separator = folderPath.endsWith("\\") ? "" : "\\";
  return `${folderPath}${separator}${stem}-dual-layer.pdf`;
}

async function loadDirectory(path = "") {
  const data = await apiGet(`/api/list?path=${encodeURIComponent(path)}`);
  currentPath = data.path;
  dialogPath.textContent = currentPath || "此电脑";
  fileList.innerHTML = "";
  const items = data.items.filter(allowedFile);
  for (const item of items) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `file-item ${item.type}`;
    button.textContent = `${item.type === "dir" ? "[文件夹]" : "[文件]"} ${item.name}`;
    button.addEventListener("click", async () => {
      if (item.type === "dir") {
        if (pickerKind === "dir" && pickerTarget === "source" && fields.mode.value === "queue-dir") {
          fields.sourcePath.value = item.path;
          dialog.close();
        } else if (pickerKind === "dir" && pickerTarget === "output") {
          fields.outputPath.value = fields.mode.value === "single" ? defaultOutputFile(item.path) : item.path;
          dialog.close();
        } else {
          await loadDirectory(item.path);
        }
      } else {
        if (pickerTarget === "source") {
          fields.sourcePath.value = item.path;
        } else {
          fields.outputPath.value = item.path;
        }
        dialog.close();
      }
    });
    fileList.appendChild(button);
  }
  goParent.disabled = !currentPath;
  goParent.dataset.parent = data.parent || "";
}

async function openPicker(target) {
  pickerTarget = target;
  const { mode } = modeInfo();
  if (target === "source") {
    pickerKind = mode === "queue-dir" ? "dir" : "file";
  } else {
    pickerKind = "dir";
  }
  dialogTitle.textContent = target === "source" ? "选择输入" : "选择输出";
  selectCurrent.hidden = pickerKind !== "dir";
  dialog.showModal();
  await loadDirectory(currentPath);
}

function buildPayload() {
  const { mode, isWord } = modeInfo();
  return {
    mode,
    inputType: isWord ? "word" : "pdf",
    sourcePath: fields.sourcePath.value,
    outputPath: fields.outputPath.value,
    language: fields.language.value,
    dpi: Number(fields.dpi.value),
    deskew: fields.deskew.checked,
    rotatePages: fields.rotatePages.checked,
    optimize: fields.optimize.checked,
    keepWork: fields.keepWork.checked,
    detailedPngProgress: fields.detailedPngProgress.checked,
    recurse: fields.recurse.checked,
    skipExisting: fields.skipExisting.checked,
  };
}

function outputHint(payload) {
  if (payload.mode === "single") {
    return `预计输出 PDF：${payload.outputPath}`;
  }
  return `预计输出文件夹：${payload.outputPath}\n队列日志：${payload.outputPath}\\_queue-logs`;
}

async function pollJob(jobId) {
  const data = await apiGet(`/api/job?id=${encodeURIComponent(jobId)}`);
  const job = data.job;
  jobState.textContent = `状态：${job.status}`;
  renderProgress(job.progress || 0, job.progressText || "");
  const outputs = job.outputs || {};
  const outputLines = [];
  if (outputs.outputPath) outputLines.push(`输出 PDF：${outputs.outputPath}`);
  if (outputs.outputDir) outputLines.push(`输出文件夹：${outputs.outputDir}`);
  if (outputs.summaryHint) outputLines.push(`队列日志：${outputs.summaryHint}`);
  const helperLines = [];
  if (job.status === "running" && String(job.progressText || "").includes("OCR")) {
    helperLines.push("OCR 正在执行。页数多、DPI 高、中文识别都会比较慢，请等待。");
  }
  logBox.textContent = [outputLines.join("\n"), helperLines.join("\n"), job.log || ""].filter(Boolean).join("\n\n");
  logBox.scrollTop = logBox.scrollHeight;
  if (job.status === "success" || job.status === "failed") {
    clearInterval(pollTimer);
    pollTimer = null;
    startJob.disabled = false;
    cleanCache.disabled = false;
    showRecent.disabled = false;
  }
}

function renderProgress(value, text) {
  const progress = Math.max(0, Math.min(100, Number(value) || 0));
  progressBar.style.width = `${progress}%`;
  progressPercent.textContent = `${progress}%`;
  progressText.textContent = text || "等待开始";
}

async function startCurrentJob() {
  startJob.disabled = true;
  cleanCache.disabled = true;
  showRecent.disabled = true;
  logBox.textContent = "";
  jobState.textContent = "正在启动...";
  renderProgress(0, "正在启动");
  try {
    const payload = buildPayload();
    logBox.textContent = outputHint(payload);
    const data = await apiPost("/api/start", payload);
    jobState.textContent = "状态：queued";
    await pollJob(data.jobId);
    pollTimer = setInterval(() => pollJob(data.jobId).catch((error) => {
      logBox.textContent = error.message;
      startJob.disabled = false;
      clearInterval(pollTimer);
    }), 3000);
  } catch (error) {
    jobState.textContent = "启动失败。";
    logBox.textContent = error.message;
    renderProgress(0, "启动失败");
    startJob.disabled = false;
    cleanCache.disabled = false;
    showRecent.disabled = false;
  }
}

fields.mode.addEventListener("change", updateLabels);
fields.inputType.addEventListener("change", updateLabels);
pickSource.addEventListener("click", () => openPicker("source"));
pickOutput.addEventListener("click", () => openPicker("output"));
startJob.addEventListener("click", startCurrentJob);
cleanCache.addEventListener("click", async () => {
  cleanCache.disabled = true;
  jobState.textContent = "正在清理缓存...";
  try {
    const data = await apiPost("/api/clean-cache", {});
    jobState.textContent = `缓存清理完成：${data.count} 个文件夹`;
    logBox.textContent = data.removed.join("\n") || "没有发现缓存。";
  } catch (error) {
    jobState.textContent = "缓存清理失败。";
    logBox.textContent = error.message;
  } finally {
    cleanCache.disabled = false;
  }
});
showRecent.addEventListener("click", async () => {
  jobState.textContent = "最近输出";
  try {
    const data = await apiGet("/api/recent");
    const lines = [];
    lines.push("最近生成的 PDF：");
    lines.push(...(data.pdfs.length ? data.pdfs.map((item) => `${item.modified}  ${item.path}`) : ["没有找到。"]));
    lines.push("");
    lines.push("队列汇总：");
    lines.push(...(data.summaries.length ? data.summaries.map((item) => `${item.modified}  ${item.path}`) : ["没有找到。"]));
    lines.push("");
    lines.push("任务日志：");
    lines.push(...(data.logs.length ? data.logs.map((item) => `${item.modified}  ${item.path}`) : ["没有找到。"]));
    logBox.textContent = lines.join("\n");
  } catch (error) {
    logBox.textContent = error.message;
  }
});
closeDialog.addEventListener("click", () => dialog.close());
goParent.addEventListener("click", () => loadDirectory(goParent.dataset.parent || ""));
selectCurrent.addEventListener("click", () => {
  if (pickerTarget === "source") fields.sourcePath.value = currentPath;
  if (pickerTarget === "output") {
    fields.outputPath.value = fields.mode.value === "single" ? defaultOutputFile(currentPath) : currentPath;
  }
  dialog.close();
});

updateLabels();
renderProgress(0, "尚未开始");
