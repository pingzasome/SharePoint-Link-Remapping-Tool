const elements = {
  fileInput: document.getElementById("fileInput"),
  previewBtn: document.getElementById("previewBtn"),
  runBtn: document.getElementById("runBtn"),
  exportBtn: document.getElementById("exportBtn"),
  clearBtn: document.getElementById("clearBtn"),
  errorPanel: document.getElementById("errorPanel"),
  resultBody: document.getElementById("resultBody"),
  tableHint: document.getElementById("tableHint"),
  loadingState: document.getElementById("loadingState"),
  totalRows: document.getElementById("totalRows"),
  foundRows: document.getElementById("foundRows"),
  notFoundRows: document.getElementById("notFoundRows"),
  multipleRows: document.getElementById("multipleRows"),
  errorRows: document.getElementById("errorRows"),
};

let hasPreview = false;

elements.previewBtn.addEventListener("click", previewFile);
elements.runBtn.addEventListener("click", runMatching);
elements.clearBtn.addEventListener("click", clearResults);
elements.exportBtn.addEventListener("click", (event) => {
  if (elements.exportBtn.getAttribute("aria-disabled") === "true") {
    event.preventDefault();
  }
});

disableExport();

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

  setBusy(true, "Searching SharePoint");
  try {
    const payload = await requestJson("/api/run", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(config),
    });
    renderRows(payload.results);
    setSummary(payload.summary);
    elements.tableHint.textContent = "Matching complete. Export the CSV when ready.";
    enableExport();
  } catch (error) {
    showError(error.message);
  } finally {
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
    elements.resultBody.innerHTML = "";
    elements.tableHint.textContent = "Preview rows will appear here before matching.";
    setSummary({ total: 0, found: 0, notFound: 0, multipleMatch: 0, error: 0 });
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

function renderRows(rows) {
  elements.resultBody.innerHTML = "";
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
  elements.runBtn.disabled = isBusy;
  elements.clearBtn.disabled = isBusy;
  elements.loadingState.classList.toggle("hidden", !isBusy);
  elements.loadingState.classList.toggle("flex", isBusy);
  elements.loadingState.lastChild.textContent = label;
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
