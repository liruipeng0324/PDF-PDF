const fields = {
  mode: document.querySelector("#mode"),
  inputType: document.querySelector("#inputType"),
  sourcePath: document.querySelector("#sourcePath"),
  outputPath: document.querySelector("#outputPath"),
  dpi: document.querySelector("#dpi"),
  optimize: document.querySelector("#optimize"),
  acrobatDoubleLayer: document.querySelector("#acrobatDoubleLayer"),
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

let currentJobId = "";
let pollTimer = null;
let selectedSourcePaths = [];
let lastOutputs = {};
const outputPathCacheKey = "dualPdfOutputPath";

const modeInfo = {
  single: {
    sourceLabel: "\u8f93\u5165\u6587\u4ef6",
    outputLabel: "\u8f93\u51fa\u6587\u4ef6\u5939",
    sourceButton: "\u9009\u62e9\u6587\u4ef6",
    outputButton: "\u9009\u62e9\u76ee\u5f55",
    sourcePlaceholder: "\u70b9\u51fb\u9009\u62e9 Word \u6216 PDF \u6587\u4ef6",
    outputPlaceholder: "\u70b9\u51fb\u9009\u62e9\u8f93\u51fa\u76ee\u5f55",
  },
  "queue-dir": {
    sourceLabel: "\u8f93\u5165\u6587\u4ef6",
    outputLabel: "\u8f93\u51fa\u6587\u4ef6\u5939",
    sourceButton: "\u9009\u62e9\u591a\u4e2a\u6587\u4ef6",
    outputButton: "\u9009\u62e9\u76ee\u5f55",
    sourcePlaceholder: "\u70b9\u51fb\u9009\u62e9\u591a\u4e2a Word/PDF \u6587\u4ef6",
    outputPlaceholder: "\u70b9\u51fb\u9009\u62e9\u8f93\u51fa\u76ee\u5f55",
  },
  "queue-csv": {
    sourceLabel: "\u961f\u5217\u8868\u683c",
    outputLabel: "\u8f93\u51fa\u6587\u4ef6\u5939",
    sourceButton: "\u9009\u62e9 CSV",
    outputButton: "\u9009\u62e9\u76ee\u5f55",
    sourcePlaceholder: "\u70b9\u51fb\u9009\u62e9\u961f\u5217 CSV",
    outputPlaceholder: "\u70b9\u51fb\u9009\u62e9\u8f93\u51fa\u76ee\u5f55",
  },
};

async function apiGet(url) {
  const response = await fetch(url);
  const data = await response.json();
  if (!response.ok || !data.ok) {
    throw new Error(data.error || "请求失败");
  }
  return data;
}

async function apiPost(url, payload) {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const data = await response.json();
  if (!response.ok || !data.ok) {
    throw new Error(data.error || "请求失败");
  }
  return data;
}

function isQueueMode() {
  return fields.mode.value !== "single";
}

function updateLabels() {
  const info = modeInfo[fields.mode.value] || modeInfo.single;
  sourceLabel.textContent = info.sourceLabel;
  outputLabel.textContent = info.outputLabel;
  pickSource.textContent = info.sourceButton;
  pickOutput.textContent = info.outputButton;
  fields.sourcePath.placeholder = info.sourcePlaceholder;
  fields.outputPath.placeholder = info.outputPlaceholder;
  fields.recurse.closest("label").hidden = fields.mode.value !== "queue-dir";
  fields.skipExisting.closest("label").hidden = !isQueueMode();
  clearSourceSelection(false);
}

function clearSourceSelection(clearInput = true) {
  selectedSourcePaths = [];
  if (clearInput) {
    fields.sourcePath.value = "";
  }
}

function sourceFileFilter() {
  if (fields.mode.value === "queue-csv") {
    return "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
  }
  if (fields.inputType.value === "word") {
    return "Word files (*.doc;*.docx)|*.doc;*.docx|All files (*.*)|*.*";
  }
  if (fields.inputType.value === "pdf") {
    return "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*";
  }
  return "Word/PDF files (*.doc;*.docx;*.pdf)|*.doc;*.docx;*.pdf|All files (*.*)|*.*";
}

function filenameStem(sourcePath) {
  if (!sourcePath) {
    return "";
  }
  const slash = Math.max(sourcePath.lastIndexOf("\\"), sourcePath.lastIndexOf("/"));
  const name = slash >= 0 ? sourcePath.slice(slash + 1) : sourcePath;
  return name.replace(/\.[^.]+$/, "") || "output";
}

function joinPath(dir, filename) {
  if (!dir) {
    return filename;
  }
  const separator = dir.endsWith("\\") || dir.endsWith("/") ? "" : "\\";
  return `${dir}${separator}${filename}`;
}

async function openSystemPicker(target) {
  const isSource = target === "source";
  const multiFiles = isSource && fields.mode.value === "queue-dir";
  const kind = isSource ? (multiFiles ? "files" : "file") : "dir";
  const title = isSource
    ? (multiFiles ? "\u9009\u62e9\u591a\u4e2a\u8f93\u5165\u6587\u4ef6\uff0c\u6309\u4f4f Ctrl \u6216 Shift \u591a\u9009" : "\u9009\u62e9\u8f93\u5165\u6587\u4ef6")
    : "\u9009\u62e9\u8f93\u51fa\u6587\u4ef6\u5939";
  const payload = {
    kind,
    target,
    title,
    initial: isSource ? fields.sourcePath.value : fields.outputPath.value,
    filter: isSource ? sourceFileFilter() : "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*",
  };

  const button = isSource ? pickSource : pickOutput;
  button.disabled = true;
  try {
    const data = await apiPost("/api/pick", payload);
    const paths = data.selectedPaths || (data.selected ? [data.selected] : []);
    if (!paths.length) {
      return;
    }

    if (isSource) {
      if (multiFiles) {
        selectedSourcePaths = paths;
        fields.sourcePath.value = `\u5df2\u9009\u62e9 ${paths.length} \u4e2a\u6587\u4ef6`;
      } else {
        selectedSourcePaths = [];
        fields.sourcePath.value = paths[0];
      }
    } else {
      fields.outputPath.value = paths[0];
      localStorage.setItem(outputPathCacheKey, paths[0]);
    }
  } catch (error) {
    logBox.textContent = error.message;
  } finally {
    button.disabled = false;
  }
}

function buildPayload() {
  const sourcePath = fields.sourcePath.value.trim();
  let outputPath = fields.outputPath.value.trim();
  if (fields.mode.value === "single" && sourcePath && outputPath) {
    outputPath = joinPath(outputPath, `${filenameStem(sourcePath)}-dual-layer.pdf`);
  }

  const payload = {
    mode: fields.mode.value,
    inputType: fields.inputType.value,
    sourcePath,
    outputPath,
    dpi: Number(fields.dpi.value || 300),
    optimize: fields.optimize.checked,
    acrobatDoubleLayer: fields.acrobatDoubleLayer.checked,
    acrobatOcrLanguage: "CHS",
    keepWork: fields.keepWork.checked,
    detailedPngProgress: fields.detailedPngProgress.checked,
    recurse: fields.recurse.checked,
    skipExisting: fields.skipExisting.checked,
  };

  if (fields.mode.value === "queue-dir") {
    payload.sourcePaths = selectedSourcePaths.slice();
    payload.sourcePath = selectedSourcePaths.length ? selectedSourcePaths[0] : "";
  }
  return payload;
}

function validatePayload(payload) {
  if (fields.mode.value === "queue-dir") {
    if (!payload.sourcePaths.length) {
      return "\u8bf7\u5148\u9009\u62e9\u4e00\u4e2a\u6216\u591a\u4e2a\u8f93\u5165\u6587\u4ef6\u3002";
    }
    if (!payload.outputPath) {
      return "\u8bf7\u9009\u62e9\u8f93\u51fa\u6587\u4ef6\u5939\u3002";
    }
    return "";
  }
  if (!payload.sourcePath) {
    return "\u8bf7\u9009\u62e9\u8f93\u5165\u6587\u4ef6\u3002";
  }
  if (!payload.outputPath) {
    return "\u8bf7\u9009\u62e9\u8f93\u51fa\u6587\u4ef6\u5939\u3002";
  }
  return "";
}

function setBusy(running) {
  startJob.disabled = running;
  cancelJob.disabled = !running;
  pauseJob.disabled = !running || fields.mode.value === "single";
  resumeJob.disabled = !running || fields.mode.value === "single";
  skipJob.disabled = !running || fields.mode.value === "single";
}

function renderProgress(value, text) {
  const progress = Math.max(0, Math.min(100, Number(value || 0)));
  progressBar.style.width = `${progress}%`;
  progressPercent.textContent = `${progress}%`;
  progressText.textContent = text || "正在运行";
}

function renderOutputs(outputs = {}) {
  lastOutputs = outputs;
  const hasOutput = Boolean(outputs.outputPath || outputs.outputDir || outputs.summaryHint);
  resultActions.hidden = !hasOutput;
  openOutput.disabled = !(outputs.outputPath || outputs.outputDir);
  openLogDir.disabled = !outputs.summaryHint;
}

function renderFriendlyError(message) {
  friendlyError.hidden = !message;
  friendlyError.textContent = message || "";
}

async function pollJob() {
  if (!currentJobId) {
    return;
  }
  try {
    const data = await apiGet(`/api/job?id=${encodeURIComponent(currentJobId)}`);
    const job = data.job;
    jobState.textContent = `状态：${job.status || "unknown"}`;
    renderProgress(job.progress, job.progressText);
    renderOutputs(job.outputs);
    renderFriendlyError(job.friendlyError);
    logBox.textContent = job.log || "";
    logBox.scrollTop = logBox.scrollHeight;

    if (["success", "failed", "cancelled"].includes(job.status)) {
      clearInterval(pollTimer);
      pollTimer = null;
      setBusy(false);
    }
  } catch (error) {
    clearInterval(pollTimer);
    pollTimer = null;
    setBusy(false);
    logBox.textContent = error.message;
  }
}

async function startCurrentJob() {
  const payload = buildPayload();
  const validation = validatePayload(payload);
  if (validation) {
    logBox.textContent = validation;
    return;
  }
  if (fields.outputPath.value.trim()) {
    localStorage.setItem(outputPathCacheKey, fields.outputPath.value.trim());
  }

  setBusy(true);
  renderFriendlyError("");
  renderProgress(0, "正在提交任务");
  jobState.textContent = "状态：提交中";
  logBox.textContent = "";

  try {
    const data = await apiPost("/api/start", payload);
    currentJobId = data.jobId;
    jobState.textContent = "状态：running";
    pollJob();
    pollTimer = setInterval(pollJob, 1000);
  } catch (error) {
    setBusy(false);
    jobState.textContent = "状态：failed";
    logBox.textContent = error.message;
  }
}

async function cancelCurrentJob() {
  if (!currentJobId) {
    return;
  }
  cancelJob.disabled = true;
  try {
    await apiPost("/api/cancel", { jobId: currentJobId });
    jobState.textContent = "状态：正在取消";
  } catch (error) {
    logBox.textContent = error.message;
  }
}

async function controlCurrentJob(action) {
  if (!currentJobId) {
    return;
  }
  try {
    await apiPost("/api/control", { jobId: currentJobId, action });
  } catch (error) {
    logBox.textContent = error.message;
  }
}

async function openPath(path) {
  if (!path) {
    return;
  }
  try {
    await apiPost("/api/open-path", { path });
  } catch (error) {
    logBox.textContent = error.message;
  }
}

async function loadToolChecks() {
  toolChecks.textContent = "正在检测组件...";
  try {
    const data = await apiGet("/api/tools");
    const checks = data.checks || [];
    toolChecks.innerHTML = checks.map((item) => {
      const ok = item.ok ? "可用" : "缺失";
      const className = item.ok ? "ok" : "missing";
      const detail = item.detail ? `<p>${item.detail}</p>` : "";
      return `<article class="${className}"><span>${ok}</span><strong>${item.name}</strong>${detail}</article>`;
    }).join("");
  } catch (error) {
    toolChecks.textContent = error.message;
  }
}

pickSource.addEventListener("click", () => openSystemPicker("source"));
pickOutput.addEventListener("click", () => openSystemPicker("output"));
startJob.addEventListener("click", startCurrentJob);
cancelJob.addEventListener("click", cancelCurrentJob);
pauseJob.addEventListener("click", () => controlCurrentJob("pause"));
resumeJob.addEventListener("click", () => controlCurrentJob("resume"));
skipJob.addEventListener("click", () => controlCurrentJob("skip"));
refreshTools.addEventListener("click", loadToolChecks);
openOutput.addEventListener("click", () => openPath(lastOutputs.outputPath || lastOutputs.outputDir));
openLogDir.addEventListener("click", () => openPath(lastOutputs.summaryHint));
fields.mode.addEventListener("change", updateLabels);
fields.inputType.addEventListener("change", () => clearSourceSelection(true));

cleanCache.addEventListener("click", async () => {
  cleanCache.disabled = true;
  jobState.textContent = "正在清理缓存...";
  try {
    const data = await apiPost("/api/clean-cache", {});
    jobState.textContent = `缓存清理完成：${data.count} 个文件夹`;
    logBox.textContent = (data.removed || []).join("\n") || "没有发现缓存。";
  } catch (error) {
    jobState.textContent = "缓存清理失败";
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

fields.inputType.value = "word";
fields.outputPath.value = localStorage.getItem(outputPathCacheKey) || "";
updateLabels();
fields.dpi.value = "300";
fields.optimize.checked = true;
renderProgress(0, "尚未开始");
loadToolChecks();
