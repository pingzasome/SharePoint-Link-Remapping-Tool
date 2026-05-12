const elements = {
  fileInput: document.getElementById("fileInput"),
  previewBtn: document.getElementById("previewBtn"),
  runBtn: document.getElementById("runBtn"),
  exportBtn: document.getElementById("exportBtn"),
  clearBtn: document.getElementById("clearBtn"),
  toggleSecretBtn: document.getElementById("toggleSecretBtn"),
  clientSecret: document.getElementById("clientSecret"),
  fileInfo: document.getElementById("fileInfo"),
  configCheck: document.getElementById("configCheck"),
  fileCheck: document.getElementById("fileCheck"),
  resultCheck: document.getElementById("resultCheck"),
  errorPanel: document.getElementById("errorPanel"),
  resultBody: document.getElementById("resultBody"),
  tableHint: document.getElementById("tableHint"),
  loadingState: document.getElementById("loadingState"),
  totalRows: document.getElementById("totalRows"),
  foundRows: document.getElementById("foundRows"),
  notFoundRows: document.getElementById("notFoundRows"),
  multipleRows: document.getElementById("multipleRows"),
  errorRows: document.getElementById("errorRows"),
  progressPanel: document.getElementById("progressPanel"),
  progressMessage: document.getElementById("progressMessage"),
  startedAtText: document.getElementById("startedAtText"),
  elapsedText: document.getElementById("elapsedText"),
  progressRows: document.getElementById("progressRows"),
  progressPercent: document.getElementById("progressPercent"),
  progressBar: document.getElementById("progressBar"),
};

let hasPreview = false;
let progressPollId = null;
let elapsedTimerId = null;
let currentJobStartedAt = null;

elements.previewBtn.addEventListener("click", previewFile);
elements.runBtn.addEventListener("click", runMatching);
elements.clearBtn.addEventListener("click", clearResults);
elements.toggleSecretBtn.addEventListener("click", toggleSecretVisibility);
["tenantId", "clientId", "sharepointHost", "sharepointSitePath"].forEach((id) => {
  document.getElementById(id).addEventListener("input", updateConfigCheck);
});
elements.exportBtn.addEventListener("click", (event) => {
  if (elements.exportBtn.getAttribute("aria-disabled") === "true") {
    event.preventDefault();
  }
});

disableExport();
setRunReady(false);
renderEmptyState("Upload and preview a ReplaceMagic export to begin.");
updateConfigCheck();

async function previewFile() {
  clearError();
  const file = elements.fileInput.files[0];
  if (!file) {
    showError("Choose a CSV or XLSX ReplaceMagic export file first.");
    return;
  }

  const formData = new FormData();
  formData.append("file", file);

  setBusy(true, "Previewing file");
  try {
    const payload = await requestJson("/api/upload", { method: "POST", body: formData });
    hasPreview = true;
    resetProgress();
    setRunReady(true);
    setFileInfo(`${payload.filename} - ${payload.totalRows} rows loaded`);
    setChecklistState(elements.fileCheck, "ReplaceMagic file", `${payload.totalRows} rows ready`, "ready");
    setChecklistState(elements.resultCheck, "Latest result", "Preview only", "idle");
    setSummary({ total: payload.totalRows, found: 0, notFound: 0, multipleMatch: 0, error: 0 });
    renderRows(
      payload.preview.map((row) => ({
        OldPath: row.OldPath,
        OldURL: row.OldURL,
        SearchFileName: row.SearchFileName,
        MatchedFileName: "",
        NewURL: "",
        Status: "PREVIEW",
        Remark: `Row ${row.rowNumber}`,
      })),
    );
    elements.tableHint.textContent = `Showing ${payload.preview.length} preview rows from ${payload.totalRows} uploaded rows.`;
    disableExport();
  } catch (error) {
    showError(error.message);
  } finally {
    setBusy(false);
  }
}

async function runMatching() {
  clearError();
  if (!hasPreview) {
    showError("Preview a ReplaceMagic file before running matching.");
    return;
  }

  const config = {
    tenantId: valueOf("tenantId"),
    clientId: valueOf("clientId"),
    clientSecret: valueOf("clientSecret"),
    sharepointHost: valueOf("sharepointHost"),
    sharepointSitePath: valueOf("sharepointSitePath"),
  };

  const configError = validateConfig(config);
  if (configError) {
    showError(configError);
    return;
  }

  setBusy(true, "Searching SharePoint");
  try {
    const payload = await requestJson("/api/run", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(config),
    });
    startProgress(payload);
  } catch (error) {
    showError(error.message);
    setBusy(false);
  }
}

async function clearResults() {
  clearError();
  setBusy(true, "Clearing results");
  try {
    await requestJson("/api/clear", { method: "POST" });
    hasPreview = false;
    elements.fileInput.value = "";
    setRunReady(false);
    setFileInfo("");
    renderEmptyState("Upload and preview a ReplaceMagic export to begin.");
    elements.tableHint.textContent = "Preview rows will appear here before matching.";
    setChecklistState(elements.fileCheck, "ReplaceMagic file", "Not previewed", "idle");
    setChecklistState(elements.resultCheck, "Latest result", "No result", "idle");
    setSummary({ total: 0, found: 0, notFound: 0, multipleMatch: 0, error: 0 });
    resetProgress();
    disableExport();
  } catch (error) {
    showError(error.message);
  } finally {
    setBusy(false);
  }
}

async function requestJson(url, options) {
  const response = await fetch(url, options);
  const contentType = response.headers.get("content-type") || "";
  const payload = contentType.includes("application/json") ? await response.json() : {};
  if (!response.ok) {
    throw new Error(payload.detail || `Request failed with HTTP ${response.status}`);
  }
  return payload;
}

function startProgress(job) {
  stopProgressTimers();
  disableExport();
  currentJobStartedAt = job.startedAt ? new Date(job.startedAt) : new Date();
  elements.progressPanel.classList.remove("hidden");
  elements.startedAtText.textContent = formatTime(currentJobStartedAt);
  updateProgressUi({
    status: job.status,
    total: job.total || 0,
    processed: 0,
    percent: 0,
    elapsedSeconds: 0,
    message: "Starting SharePoint matching",
    summary: { total: job.total || 0, found: 0, notFound: 0, multipleMatch: 0, error: 0 },
  });

  elapsedTimerId = window.setInterval(() => {
    if (!currentJobStartedAt) {
      return;
    }
    const elapsedSeconds = Math.max(0, Math.floor((Date.now() - currentJobStartedAt.getTime()) / 1000));
    elements.elapsedText.textContent = formatDuration(elapsedSeconds);
  }, 1000);

  progressPollId = window.setInterval(() => pollProgress(job.jobId), 1000);
  pollProgress(job.jobId);
}

async function pollProgress(jobId) {
  try {
    const progress = await requestJson(`/api/progress/${jobId}`, { method: "GET" });
    updateProgressUi(progress);

    if (Array.isArray(progress.results) && progress.results.length > 0) {
      renderRows(progress.results);
    }
    if (progress.summary) {
      setSummary(progress.summary);
    }

    if (progress.status === "COMPLETED") {
      stopProgressTimers();
      updateProgressUi(progress);
      renderRows(progress.results || []);
      setSummary(progress.summary || {});
      elements.tableHint.textContent = `Matching complete in ${formatDuration(progress.elapsedSeconds || 0)}. Export the CSV when ready.`;
      setChecklistState(elements.resultCheck, "Latest result", `Completed in ${formatDuration(progress.elapsedSeconds || 0)}`, "ready");
      setBusy(false);
      enableExport();
    } else if (progress.status === "ERROR") {
      stopProgressTimers();
      setBusy(false);
      setChecklistState(elements.resultCheck, "Latest result", "Failed", "error");
      showError(progress.error || "Matching failed.");
      elements.tableHint.textContent = "Matching failed. Review the error and try again.";
    }
  } catch (error) {
    stopProgressTimers();
    setBusy(false);
    showError(error.message);
  }
}

function updateProgressUi(progress) {
  const total = progress.total || 0;
  const processed = progress.processed || 0;
  const percent = Math.max(0, Math.min(100, progress.percent || 0));

  elements.progressMessage.textContent = progress.message || "Running";
  elements.progressRows.textContent = `${processed} / ${total}`;
  elements.progressPercent.textContent = `${percent}%`;
  elements.progressBar.style.width = `${percent}%`;
  elements.elapsedText.textContent = formatDuration(progress.elapsedSeconds || 0);
  if (progress.startedAt) {
    currentJobStartedAt = new Date(progress.startedAt);
    elements.startedAtText.textContent = formatTime(currentJobStartedAt);
  }
}

function resetProgress() {
  stopProgressTimers();
  currentJobStartedAt = null;
  elements.progressPanel.classList.add("hidden");
  elements.progressMessage.textContent = "Waiting to start";
  elements.startedAtText.textContent = "-";
  elements.elapsedText.textContent = "00:00";
  elements.progressRows.textContent = "0 / 0";
  elements.progressPercent.textContent = "0%";
  elements.progressBar.style.width = "0%";
}

function stopProgressTimers() {
  if (progressPollId) {
    window.clearInterval(progressPollId);
    progressPollId = null;
  }
  if (elapsedTimerId) {
    window.clearInterval(elapsedTimerId);
    elapsedTimerId = null;
  }
}

function renderRows(rows) {
  elements.resultBody.innerHTML = "";
  if (!rows || rows.length === 0) {
    renderEmptyState("No rows to display yet.");
    return;
  }

  const fragment = document.createDocumentFragment();

  rows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.className = "align-top hover:bg-slate-50";
    ["OldPath", "OldURL", "SearchFileName", "MatchedFileName", "NewURL", "Status", "Remark"].forEach((key) => {
      const td = document.createElement("td");
      td.className = cellClass(key);
      if (key === "Status") {
        td.appendChild(statusBadge(row[key] || ""));
      } else if (key === "NewURL" && isHttpUrl(row[key])) {
        const link = document.createElement("a");
        link.href = row[key];
        link.target = "_blank";
        link.rel = "noreferrer";
        link.className = "text-cyan-800 underline decoration-cyan-800/30 underline-offset-2 hover:text-cyan-950";
        link.textContent = row[key];
        td.appendChild(link);
      } else {
        td.textContent = row[key] || "";
      }
      tr.appendChild(td);
    });
    fragment.appendChild(tr);
  });

  elements.resultBody.appendChild(fragment);
}

function renderEmptyState(message) {
  elements.resultBody.innerHTML = "";
  const tr = document.createElement("tr");
  const td = document.createElement("td");
  td.colSpan = 7;
  td.className = "px-4 py-10 text-center text-sm text-slate-500";
  td.textContent = message;
  tr.appendChild(td);
  elements.resultBody.appendChild(tr);
}

function cellClass(key) {
  const base = "break-words px-4 py-3 text-slate-700";
  const classes = {
    Status: "px-4 py-3 text-slate-700 whitespace-nowrap",
    Remark: "break-words px-4 py-3 text-slate-700 text-sm",
    NewURL: "break-all px-4 py-3 text-slate-700",
  };
  return classes[key] || base;
}

function isHttpUrl(value) {
  return typeof value === "string" && /^https?:\/\//i.test(value);
}

function statusBadge(status) {
  const span = document.createElement("span");
  const colors = {
    FOUND: "bg-emerald-50 text-emerald-800 ring-emerald-200",
    NOT_FOUND: "bg-amber-50 text-amber-800 ring-amber-200",
    MULTIPLE_MATCH: "bg-indigo-50 text-indigo-800 ring-indigo-200",
    ERROR: "bg-red-50 text-red-800 ring-red-200",
    PREVIEW: "bg-slate-100 text-slate-700 ring-slate-200",
  };
  span.className = `inline-flex rounded px-2 py-1 text-xs font-semibold ring-1 ${colors[status] || colors.PREVIEW}`;
  span.textContent = status;
  return span;
}

function setSummary(summary) {
  elements.totalRows.textContent = summary.total || 0;
  elements.foundRows.textContent = summary.found || 0;
  elements.notFoundRows.textContent = summary.notFound || 0;
  elements.multipleRows.textContent = summary.multipleMatch || 0;
  elements.errorRows.textContent = summary.error || 0;
}

function setBusy(isBusy, label = "Searching SharePoint") {
  elements.previewBtn.disabled = isBusy;
  elements.runBtn.disabled = isBusy || !hasPreview;
  elements.clearBtn.disabled = isBusy;
  elements.loadingState.classList.toggle("hidden", !isBusy);
  elements.loadingState.classList.toggle("flex", isBusy);
  elements.loadingState.lastChild.textContent = label;
}

function setRunReady(isReady) {
  elements.runBtn.disabled = !isReady;
}

function setFileInfo(message) {
  if (!message) {
    elements.fileInfo.textContent = "";
    elements.fileInfo.classList.add("hidden");
    return;
  }
  elements.fileInfo.textContent = message;
  elements.fileInfo.classList.remove("hidden");
}

function validateConfig(config) {
  const missing = [];
  if (!config.tenantId) missing.push("Tenant ID");
  if (!config.clientId) missing.push("Client ID");
  if (!config.sharepointHost) missing.push("SharePoint Host");
  if (!config.sharepointSitePath) missing.push("SharePoint Site Path");
  if (missing.length > 0) {
    return `Missing required configuration: ${missing.join(", ")}`;
  }
  if (config.sharepointHost.startsWith("http")) {
    return "SharePoint Host should be only the host name, for example contoso.sharepoint.com.";
  }
  if (!config.sharepointSitePath.startsWith("/sites/")) {
    return "SharePoint Site Path should look like /sites/SiteName.";
  }
  return "";
}

function updateConfigCheck() {
  const config = {
    tenantId: valueOf("tenantId"),
    clientId: valueOf("clientId"),
    sharepointHost: valueOf("sharepointHost"),
    sharepointSitePath: valueOf("sharepointSitePath"),
  };
  if (validateConfig(config)) {
    setChecklistState(elements.configCheck, "SharePoint config", "Needs attention", "idle");
  } else {
    setChecklistState(elements.configCheck, "SharePoint config", "Looks ready", "ready");
  }
}

function setChecklistState(element, label, value, state) {
  const classes = {
    ready: "flex items-center justify-between gap-3 rounded-md bg-emerald-50 px-3 py-2 text-emerald-900",
    error: "flex items-center justify-between gap-3 rounded-md bg-red-50 px-3 py-2 text-red-900",
    idle: "flex items-center justify-between gap-3 rounded-md bg-slate-50 px-3 py-2 text-slate-600",
  };
  element.className = classes[state] || classes.idle;
  element.innerHTML = `<span>${label}</span><span class="font-semibold">${value}</span>`;
}

function toggleSecretVisibility() {
  const isPassword = elements.clientSecret.type === "password";
  elements.clientSecret.type = isPassword ? "text" : "password";
  elements.toggleSecretBtn.textContent = isPassword ? "Hide" : "Show";
}

function showError(message) {
  elements.errorPanel.textContent = message;
  elements.errorPanel.classList.remove("hidden");
}

function clearError() {
  elements.errorPanel.textContent = "";
  elements.errorPanel.classList.add("hidden");
}

function enableExport() {
  elements.exportBtn.className = "rounded-md bg-slate-900 px-4 py-2 text-center text-sm font-semibold text-white hover:bg-slate-700";
  elements.exportBtn.classList.remove("pointer-events-none");
  elements.exportBtn.setAttribute("aria-disabled", "false");
  elements.exportBtn.setAttribute("tabindex", "0");
}

function disableExport() {
  elements.exportBtn.className = "pointer-events-none rounded-md bg-slate-300 px-4 py-2 text-center text-sm font-semibold text-white";
  elements.exportBtn.setAttribute("aria-disabled", "true");
  elements.exportBtn.setAttribute("tabindex", "-1");
}

function valueOf(id) {
  return document.getElementById(id).value.trim();
}

function formatDuration(totalSeconds) {
  const seconds = Math.max(0, Math.floor(Number(totalSeconds) || 0));
  const hours = Math.floor(seconds / 3600);
  const minutes = Math.floor((seconds % 3600) / 60);
  const remainingSeconds = seconds % 60;
  if (hours > 0) {
    return `${pad(hours)}:${pad(minutes)}:${pad(remainingSeconds)}`;
  }
  return `${pad(minutes)}:${pad(remainingSeconds)}`;
}

function formatTime(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) {
    return "-";
  }
  return date.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", second: "2-digit" });
}

function pad(value) {
  return String(value).padStart(2, "0");
}
