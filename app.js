const fields = {
  mode: document.querySelector("#mode"),
  inputType: document.querySelector("#inputType"),
  sourcePath: document.querySelector("#sourcePath"),
  outputPath: document.querySelector("#outputPath"),
  ocrMode: document.querySelector("#ocrMode"),
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
const cancelJob = document.querySelector("#cancelJob");
const pauseJob = document.querySelector("#pauseJob");
const resumeJob = document.querySelector("#resumeJob");
const skipJob = document.querySelector("#skipJob");
const cleanCache = document.querySelector("#cleanCache");
const showRecent = document.querySelector("#showRecent");
const refreshTools = document.querySelector("#refreshTools");
const toolChecks = document.querySelector("#toolChecks");
const jobState = document.querySelector("#jobState");
const logBox = document.querySelector("#logBox");
const progressText = document.querySelector("#progressText");
const progressPercent = document.querySelector("#progressPercent");
const progressBar = document.querySelector("#progressBar");
const resultActions = document.querySelector("#resultActions");
const openOutput = document.querySelector("#openOutput");
const openLogDir = document.querySelector("#openLogDir");
const friendlyError = document.querySelector("#friendlyError");

let pollTimer = null;
let currentJobId = "";
let lastOutputs = {};

function modeInfo() {
  return {
    mode: fields.mode.value,
    isWord: fields.inputType.value === "word",
  };
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
  outputLabel.textContent = "输出文件夹";
  fields.sourcePath.placeholder = "点击选择输入";
  fields.outputPath.placeholder = "点击选择输出目录";
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

function defaultOutputFile(folderPath) {
  const source = fields.sourcePath.value || "final.pdf";
  const normalized = source.replaceAll("\\", "/");
  const filename = normalized.split("/").pop() || "final.pdf";
  const stem = filename.replace(/\.[^.]+$/, "");
  const separator = folderPath.endsWith("\\") ? "" : "\\";
  return `${folderPath}${separator}${stem}-dual-layer.pdf`;
}

function pickerFilter(target) {
  const { mode, isWord } = modeInfo();
  if (target === "output") return "所有文件 (*.*)|*.*";
  if (mode === "queue-csv") return "CSV 队列 (*.csv)|*.csv|所有文件 (*.*)|*.*";
  if (isWord) return "Word 文档 (*.doc;*.docx)|*.doc;*.docx|所有文件 (*.*)|*.*";
  return "PDF 文件 (*.pdf)|*.pdf|所有文件 (*.*)|*.*";
}

async function openSystemPicker(target) {
  const { mode } = modeInfo();
  const kind = target === "source" && mode !== "queue-dir" ? "file" : "dir";
  const title = target === "source" ? "选择输入" : "选择输出目录";
  try {
    const data = await apiPost("/api/pick", {
      kind,
      title,
      filter: pickerFilter(target),
      initialPath: target === "source" ? fields.sourcePath.value : fields.outputPath.value,
    });
    if (!data.selected) return;
    if (target === "source") {
      fields.sourcePath.value = data.selected;
    } else {
      fields.outputPath.value = fields.mode.value === "single" ? defaultOutputFile(data.selected) : data.selected;
    }
  } catch (error) {
    jobState.textContent = "文件资源管理器没有打开。";
    logBox.textContent = error.message;
  }
}

function buildPayload() {
  const { mode, isWord } = modeInfo();
  return {
    mode,
    inputType: isWord ? "word" : "pdf",
    sourcePath: fields.sourcePath.value,
    outputPath: fields.outputPath.value,
    ocrMode: fields.ocrMode.value,
    renderEngine: "pdfium",
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

function setBusy(isBusy) {
  const isQueue = fields.mode.value !== "single";
  startJob.disabled = isBusy;
  cancelJob.disabled = !isBusy;
  pauseJob.disabled = !isBusy || !isQueue;
  resumeJob.disabled = !isBusy || !isQueue;
  skipJob.disabled = !isBusy || !isQueue;
  cleanCache.disabled = isBusy;
  showRecent.disabled = isBusy;
}

function renderProgress(value, text) {
  const progress = Math.max(0, Math.min(100, Number(value) || 0));
  progressBar.style.width = `${progress}%`;
  progressPercent.textContent = `${progress}%`;
  progressText.textContent = text || "等待开始";
}

function renderOutputs(outputs) {
  lastOutputs = outputs || {};
  const hasOutput = Boolean(lastOutputs.outputPath || lastOutputs.outputDir || lastOutputs.summaryHint);
  resultActions.hidden = !hasOutput;
  openOutput.disabled = !Boolean(lastOutputs.outputPath || lastOutputs.outputDir);
  openLogDir.disabled = !Boolean(lastOutputs.summaryHint);
}

async function pollJob(jobId) {
  const data = await apiGet(`/api/job?id=${encodeURIComponent(jobId)}`);
  const job = data.job;
  jobState.textContent = `状态：${job.status}`;
  renderProgress(job.progress || 0, job.progressText || "");
  renderOutputs(job.outputs || {});

  const outputs = job.outputs || {};
  const outputLines = [];
  if (outputs.outputPath) outputLines.push(`输出 PDF：${outputs.outputPath}`);
  if (outputs.outputDir) outputLines.push(`输出文件夹：${outputs.outputDir}`);
  if (outputs.summaryHint) outputLines.push(`队列日志：${outputs.summaryHint}`);

  const helperLines = [];
  if (job.status === "running" && String(job.progressText || "").includes("OCR")) {
    helperLines.push("OCR 正在执行。页数多、DPI 高、中文识别都会比较慢，请等待。");
  }

  friendlyError.hidden = !job.friendlyError;
  friendlyError.textContent = job.friendlyError || "";
  logBox.textContent = [outputLines.join("\n"), helperLines.join("\n"), job.log || ""].filter(Boolean).join("\n\n");
  logBox.scrollTop = logBox.scrollHeight;

  if (["success", "failed", "cancelled"].includes(job.status)) {
    clearInterval(pollTimer);
    pollTimer = null;
    currentJobId = "";
    setBusy(false);
  }
}

async function startCurrentJob() {
  setBusy(true);
  resultActions.hidden = true;
  friendlyError.hidden = true;
  friendlyError.textContent = "";
  logBox.textContent = "";
  jobState.textContent = "正在启动...";
  renderProgress(0, "正在启动");
  try {
    const payload = buildPayload();
    logBox.textContent = outputHint(payload);
    const data = await apiPost("/api/start", payload);
    currentJobId = data.jobId;
    jobState.textContent = "状态：queued";
    await pollJob(data.jobId);
    pollTimer = setInterval(() => pollJob(data.jobId).catch((error) => {
      logBox.textContent = error.message;
      setBusy(false);
      clearInterval(pollTimer);
    }), 3000);
  } catch (error) {
    jobState.textContent = "启动失败。";
    logBox.textContent = error.message;
    renderProgress(0, "启动失败");
    setBusy(false);
  }
}

async function cancelCurrentJob() {
  if (!currentJobId) return;
  cancelJob.disabled = true;
  jobState.textContent = "正在取消...";
  try {
    await apiPost("/api/cancel", { jobId: currentJobId });
  } catch (error) {
    logBox.textContent = `${logBox.textContent}\n\n取消失败：${error.message}`;
  }
}

async function controlCurrentJob(action) {
  if (!currentJobId) return;
  try {
    await apiPost("/api/control", { jobId: currentJobId, action });
    const label = action === "pause" ? "队列将在当前文件结束后暂停。" : action === "resume" ? "队列继续执行。" : "下一项将被跳过。";
    jobState.textContent = label;
  } catch (error) {
    logBox.textContent = `${logBox.textContent}\n\n队列控制失败：${error.message}`;
  }
}

async function loadToolChecks() {
  toolChecks.textContent = "正在检测组件...";
  try {
    const data = await apiGet("/api/tools");
    toolChecks.innerHTML = "";
    for (const item of data.checks) {
      const div = document.createElement("div");
      div.className = `check-item ${item.ok ? "ok" : "missing"}`;
      div.innerHTML = `<strong>${item.ok ? "可用" : "缺失"}</strong><span>${item.name}</span><small>${item.detail || item.fix}</small>`;
      toolChecks.appendChild(div);
    }
  } catch (error) {
    toolChecks.textContent = `检测失败：${error.message}`;
  }
}

async function openPath(path) {
  if (!path) return;
  await apiPost("/api/open-path", { path });
}

fields.mode.addEventListener("change", updateLabels);
fields.inputType.addEventListener("change", updateLabels);
fields.ocrMode.addEventListener("change", () => {
  if (fields.ocrMode.value === "fast") {
    fields.dpi.value = "200";
    fields.deskew.checked = false;
    fields.rotatePages.checked = false;
  }
  if (fields.ocrMode.value === "accurate") {
    fields.dpi.value = "300";
    fields.deskew.checked = true;
    fields.rotatePages.checked = true;
  }
});
pickSource.addEventListener("click", () => openSystemPicker("source"));
pickOutput.addEventListener("click", () => openSystemPicker("output"));
fields.sourcePath.addEventListener("click", () => openSystemPicker("source"));
fields.outputPath.addEventListener("click", () => openSystemPicker("output"));
startJob.addEventListener("click", startCurrentJob);
cancelJob.addEventListener("click", cancelCurrentJob);
pauseJob.addEventListener("click", () => controlCurrentJob("pause"));
resumeJob.addEventListener("click", () => controlCurrentJob("resume"));
skipJob.addEventListener("click", () => controlCurrentJob("skip"));
refreshTools.addEventListener("click", loadToolChecks);
openOutput.addEventListener("click", () => openPath(lastOutputs.outputPath || lastOutputs.outputDir));
openLogDir.addEventListener("click", () => openPath(lastOutputs.summaryHint));
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

updateLabels();
renderProgress(0, "尚未开始");
loadToolChecks();
