const state = {
  masterId: "",
  surveyId: "",
  masterName: "",
  surveyName: "",
  jobId: "",
  page: 1,
  totalResults: 0,
  pageSize: 50,
  summary: null,
  lookupSource: "current",
  lookupMasterId: "",
  lookupMasterName: "",
  lookupRows: [],
  lookupPage: 1,
  lookupPageSize: 200,
  lookupTotalRows: 0,
  files: [],
  filesPage: 1,
  filesPageSize: 8,
  surveyFiles: [],
  surveyFilesPage: 1,
  surveyFilesPageSize: 5,
  surveyFilesExpanded: false,
  roundFilesExpanded: false,
  justUploadedMasterId: "",
  expandedBatchRows: {},
  lookupOptions: {
    supplier_codes: [],
    suppliers: [],
    materials: [],
    categories: [],
    substances: [],
  },
  surveyOptions: {
    "原物料": { supplier_codes: [], suppliers: [], materials: [], categories: [], substances: [] },
    "包材": { supplier_codes: [], suppliers: [], materials: [], categories: [], substances: [] },
  },
  surveyOptionsMasterId: "",
  latestMaster: null,
  workingMaster: null,
  surveyBatches: { "原物料": [], "包材": [] },
  batchStatus: { "原物料": "pending", "包材": "pending" },
  roundProgress: {
    "原物料": { total: 0, completed: 0 },
    "包材": { total: 0, completed: 0 },
  },
  currentBatch: "",
  shownNoticeKeys: new Set(),
};

const labels = { added: "新增", modified: "修改", same: "相同", conflict: "衝突" };
const allowedActionsByType = {
  added: ["accept_new", "confirm_add", "ignore"],
  modified: ["accept_new", "keep_old", "ignore"],
  same: ["ignore"],
};

const $ = (id) => document.getElementById(id);

async function api(url, options = {}) {
  const res = await fetch(url, options);
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "系統發生錯誤");
  return data;
}

function text(v) {
  return (v ?? "").toString().replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;" }[c]));
}

function effectiveMaster() {
  if (state.workingMaster) return { ...state.workingMaster, source: "working" };
  if (state.masterId) return { id: state.masterId, original_name: state.masterName, source: "manual" };
  if (state.latestMaster) return { ...state.latestMaster, source: "latest" };
  return null;
}

function baselineLabel(master) {
  if (!master) return "尚未選定比對基準";
  if (master.source === "working") return `本輪暫存更動版 ${master.original_name}`;
  if (master.source === "latest") return `目前最新總表 ${master.original_name}`;
  return `本次指定總表 ${master.original_name}`;
}

function baselineHelp(master) {
  if (!master) return "請在「上傳」分頁指定本次比對總表，或將某份更動版定為最新總表。";
  if (master.source === "working") return "接下來若未重新選用其他總表，系統會使用這份暫存更動版繼續比對。所有調查表都處理完後，請按「完成本輪更新並定版」。";
  if (master.source === "latest") return "若未上傳並選用新的總表，系統會預設使用此最新總表進行比對。";
  return "這是本次臨時指定的比對基準，不會改變系統最新總表。";
}

function showInfo(el, file) {
  const sheets = file.info.sheets.map((s) => `${s.name}（${s.rows} 列 / ${s.columns} 欄${s.hidden ? "，隱藏" : ""}）`).join("<br>");
  el.innerHTML = `<strong>${file.name}</strong><div class="muted">${sheets}</div>`;
}

function surveyKindWarning(fileInfo, expectedSheet) {
  const name = (fileInfo.name || "").toLowerCase();
  const sheets = (fileInfo.info?.sheets || []).map((sheet) => {
    const previewText = (sheet.preview || []).flat().join(" ");
    return `${sheet.name} ${previewText}`;
  }).join(" ").toLowerCase();
  const textSource = `${name} ${sheets}`;
  const rawHints = ["原物料", "raw", "material", "原料"];
  const packageHints = ["包材", "包裝", "package", "packaging"];
  const rawScore = rawHints.filter((hint) => textSource.includes(hint.toLowerCase())).length;
  const packageScore = packageHints.filter((hint) => textSource.includes(hint.toLowerCase())).length;
  if (expectedSheet === "原物料" && packageScore > rawScore) return "這份檔案看起來可能是包材調查表";
  if (expectedSheet === "包材" && rawScore > packageScore) return "這份檔案看起來可能是原物料調查表";
  if (!rawScore && !packageScore) return "格式需人工確認";
  return "";
}

function cleanPastedSupplierCode(value) {
  return (value ?? "").toString().replace(/\r?\n/g, "").trim();
}

function supplierFieldLabel(sheetName) {
  return sheetName === "包材" ? "供應商名稱/代號" : "供應商代號";
}

function supplierDisplayValue(row) {
  return row?.supplier_code || row?.supplier || "";
}

async function clearWorkingVersion() {
  await api("/api/clear-working-version", { method: "POST" });
  state.workingMaster = null;
}

function resetRoundProgress() {
  state.roundProgress = {
    "原物料": { total: 0, completed: 0 },
    "包材": { total: 0, completed: 0 },
  };
}

function activateTab(targetId) {
  document.querySelectorAll(".tab").forEach((tab) => tab.classList.toggle("active", tab.dataset.tabTarget === targetId));
  document.querySelectorAll(".tabPage").forEach((page) => page.classList.toggle("active", page.id === targetId));
  if (targetId === "lookupPage") {
    loadLookupOptions().catch((err) => alert(err.message));
  }
  if (targetId === "surveyPage") {
    loadSurveyOptions().catch(() => {});
  }
}

function formForBatch(sheetName) {
  return $(sheetName === "包材" ? "packageSurveyForm" : "rawSurveyForm");
}

function clearSurveyForm(sheetName) {
  const form = formForBatch(sheetName);
  if (!form) return;
  form.file.value = "";
  form.supplier_code.value = "";
  form.material_name.value = "";
  if (form.subcategory) form.subcategory.value = "";
  $(sheetName === "包材" ? "packageSurveyInfo" : "rawSurveyInfo").innerHTML = "";
  updateSurveyFileLabel(form);
}

function updateSurveyFileLabel(form) {
  const label = document.querySelector(`[data-file-label="${form.id}"]`);
  if (!label) return;
  const count = form.file?.files?.length || 0;
  label.textContent = count ? `已選 ${count} 份` : "上傳 1 份 Excel";
}

function clearCompletedBatch(sheetName) {
  if (!sheetName) return;
  state.surveyBatches[sheetName] = [];
  state.batchStatus[sheetName] = "applied";
  clearSurveyForm(sheetName);
  renderBatchLists();
}

async function upload(form, kind, pickedFile = null) {
  const fd = new FormData(form);
  fd.set("kind", kind);
  if (pickedFile) fd.set("file", pickedFile);
  const data = await api("/api/upload", { method: "POST", body: fd });
  if (kind === "master") {
    state.justUploadedMasterId = data.id;
    showInfo($("masterInfo"), data);
  }
  await loadDashboard();
}

async function uploadSurveyBatch(form, sheetName) {
  const files = [...form.file.files];
  if (!files.length) return alert("請先選擇調查表檔案");
  if (files.length > 1) return alert("一次只能上傳 1 份調查表，避免同一組供應商/品項資料誤套到多份檔案。");
  const target = state.surveyBatches[sheetName];
  const defaults = {
    supplier_code: form.supplier_code.value.trim(),
    material_name: form.material_name.value.trim(),
    subcategory: form.subcategory ? form.subcategory.value.trim() : "",
  };
  if (!defaults.supplier_code) return alert(`請先填寫${supplierFieldLabel(sheetName)}`);
  if (!defaults.material_name) return alert(`請先填寫${sheetName === "包材" ? "品項/包材名稱" : "品項/原料名稱"}`);
  for (const file of files) {
    const fd = new FormData();
    fd.set("kind", "survey");
    fd.set("file", file);
    const data = await api("/api/upload", { method: "POST", body: fd });
    const warning = surveyKindWarning(data, sheetName);
    if (warning && warning !== "格式需人工確認") {
      const ok = confirm(`${data.name}：${warning}，是否仍加入${sheetName}批次？`);
      if (!ok) {
        await api(`/api/files/${data.id}`, { method: "DELETE" });
        continue;
      }
    }
    target.push({
      survey_file_id: data.id,
      original_name: data.name,
      type: sheetName,
      supplier_code: defaults.supplier_code,
      material_name: defaults.material_name,
      subcategory: defaults.subcategory,
      format_warning: warning,
    });
    state.batchStatus[sheetName] = "pending";
    state.roundProgress[sheetName].total += 1;
  }
  form.file.value = "";
  updateSurveyFileLabel(form);
  $(sheetName === "原物料" ? "rawSurveyInfo" : "packageSurveyInfo").innerHTML = `<strong>已加入目前確認的${sheetName}調查表</strong>`;
  renderBatchLists();
  renderSelectedFiles();
  renderBatchProgress();
  await loadDashboard();
}

async function uploadLookupMaster(form) {
  const fd = new FormData(form);
  fd.set("kind", "master");
  const data = await api("/api/upload", { method: "POST", body: fd });
  state.lookupSource = "upload";
  state.lookupMasterId = data.id;
  state.lookupMasterName = data.name;
  $("lookupUploadInfo").innerHTML = `<strong>${text(data.name)}</strong><div class="muted">已上傳並用於查詢</div>`;
  renderSelectedFiles();
  await loadDashboard();
  await loadLookupOptions();
}

async function selectFile(file, options = {}) {
  if (file.kind === "master" || file.kind === "updated_master") {
    if (!options.skipWorkingWarning && state.workingMaster && state.workingMaster.id !== file.id) {
      const ok = confirm(`目前已有本輪暫存更動版：${state.workingMaster.original_name}。若改選其他總表，後續比對將改用新選用的總表，原暫存更動版仍可下載，但不再作為本輪比對基準。確定改選？`);
      if (!ok) return;
      await clearWorkingVersion();
    }
    state.masterId = file.id;
    state.masterName = file.original_name;
    resetRoundProgress();
  }
  state.jobId = "";
  state.summary = null;
  state.currentBatch = "";
  $("compareSummary").innerHTML = "";
  $("resultsBody").innerHTML = "";
  $("pager").textContent = "";
  renderDiffCounts();
  renderDiffPanel();
  updatePagerControls();
  renderSelectedFiles();
  renderBatchProgress();
  loadLookupOptions().catch(() => {});
  loadSurveyOptions().catch(() => {});
}

function renderSelectedFiles() {
  const master = effectiveMaster();
  $("lookupMasterName").textContent = lookupMasterName() || "尚未選用";
  const basis = master ? baselineLabel(master) : "尚未選定比對基準";
  const recoveryNotice = state.workingMaster && !state.jobId ? `
    <div class="recoveryNotice">
      <strong>偵測到尚未定版的暫存更動版。</strong>
      <span>目前會以這份暫存更動版作為比對基準；請在下方本批套用完成提示中選擇下一步。</span>
    </div>
  ` : "";
  $("compareBasisNotice").innerHTML = `
    <strong>目前比對基準：</strong>${text(basis)}
    ${recoveryNotice}
  `;
  $("finalizeBtn").classList.toggle("hidden", !state.workingMaster || Boolean(state.jobId));
}

function renderBatchLists() {
  renderUploadBatchSummary("原物料", "rawBatchList");
  renderUploadBatchSummary("包材", "packageBatchList");
  renderRoundBatchDetails();
}

function renderUploadBatchSummary(sheetName, targetId) {
  const rows = state.surveyBatches[sheetName];
  $(targetId).innerHTML = rows.length
    ? `<span class="compactCount">本輪已加入 ${rows.length} 份</span>`
    : `<span class="muted">尚未加入${sheetName}調查表。</span>`;
}

function renderRoundBatchDetails() {
  if (!$("roundFilesList")) return;
  const raw = renderBatchList("原物料");
  const pkg = renderBatchList("包材");
  $("roundFilesList").innerHTML = `${raw}${pkg}`;
  $("roundFilesList").classList.toggle("hidden", !state.roundFilesExpanded);
  bindBatchListEvents("roundFilesList");
}

function renderBatchList(sheetName) {
  const rows = state.surveyBatches[sheetName];
  return `
    <div class="roundBatchGroup">
      <div class="roundBatchGroupTitle"><strong>${sheetName}</strong><span>${rows.length} 份</span></div>
      ${rows.length ? rows.map((row, index) => `
    <div class="batchItem compactBatchItem" data-batch-type="${sheetName}" data-batch-index="${index}">
      <div class="batchRow">
        <div class="batchFileName">
          <strong>${text(row.original_name)}</strong>
          ${row.format_warning && row.format_warning !== "格式需人工確認" ? `<span class="warningText">${text(row.format_warning)}</span>` : ""}
        </div>
        <span>${text(supplierDisplayValue(row) || "未填供應商")}</span>
        <span>${text(row.material_name || "未填品項")}</span>
        <span class="tag same">${sheetName}</span>
        <div class="batchRowActions">
          <button type="button" data-toggle-batch="${sheetName}" data-toggle-index="${index}">${state.expandedBatchRows[`${sheetName}:${index}`] ? "收合" : "編輯"}</button>
          <button type="button" data-remove-batch="${sheetName}" data-remove-index="${index}">移除</button>
        </div>
      </div>
      <div class="batchEdit ${state.expandedBatchRows[`${sheetName}:${index}`] ? "" : "hidden"}">
        <label class="suggestField">${supplierFieldLabel(sheetName)}<input data-batch-field="supplier_code" value="${text(row.supplier_code)}" autocomplete="off" required /><div class="suggestPanel batchSuggestPanel" data-batch-suggest="supplier_code"></div></label>
        <label class="suggestField">品項名稱<input data-batch-field="material_name" value="${text(row.material_name)}" autocomplete="off" placeholder="補填總表右側品項名稱" required /><div class="suggestPanel batchSuggestPanel" data-batch-suggest="material_name"></div></label>
        ${sheetName === "原物料" ? `<label class="suggestField">細項分類<input data-batch-field="subcategory" value="${text(row.subcategory)}" autocomplete="off" placeholder="例如 Solvent(溶劑)，可留空" /><div class="suggestPanel batchSuggestPanel" data-batch-suggest="subcategory"></div></label>` : ""}
      </div>
    </div>
      `).join("") : `<div class="muted">尚未加入${sheetName}調查表。</div>`}
    </div>
  `;
}

function bindBatchListEvents(targetId) {
  document.querySelectorAll(`#${targetId} [data-toggle-batch]`).forEach((button) => {
    button.addEventListener("click", () => {
      const key = `${button.dataset.toggleBatch}:${button.dataset.toggleIndex}`;
      state.expandedBatchRows[key] = !state.expandedBatchRows[key];
      renderBatchLists();
    });
  });
  document.querySelectorAll(`#${targetId} [data-batch-field]`).forEach((input) => {
    input.addEventListener("input", () => {
      const item = input.closest(".batchItem");
      const row = state.surveyBatches[item.dataset.batchType][Number(item.dataset.batchIndex)];
      row[input.dataset.batchField] = input.value;
      renderSelectedFiles();
    });
    input.addEventListener("focus", () => renderBatchEditSuggestions(input, true));
    input.addEventListener("input", () => renderBatchEditSuggestions(input, true));
    if (input.dataset.batchField === "supplier_code") {
      input.addEventListener("paste", () => setTimeout(() => {
        input.value = cleanPastedSupplierCode(input.value);
        input.dispatchEvent(new Event("input", { bubbles: true }));
      }, 0));
    }
  });
  document.querySelectorAll(`#${targetId} [data-remove-batch]`).forEach((button) => {
    button.addEventListener("click", () => {
      const removedType = button.dataset.removeBatch;
      state.surveyBatches[removedType].splice(Number(button.dataset.removeIndex), 1);
      state.roundProgress[removedType].total = Math.max(state.roundProgress[removedType].completed, state.roundProgress[removedType].total - 1);
      renderBatchLists();
      renderSelectedFiles();
      renderBatchProgress();
    });
  });
}

function batchEditOptions(sheetName, fieldName) {
  const options = state.surveyOptions[sheetName] || {};
  if (fieldName === "supplier_code") {
    return sheetName === "包材" ? [...(options.suppliers || []), ...(options.supplier_codes || [])] : (options.supplier_codes || []);
  }
  if (fieldName === "material_name") return options.materials || [];
  if (fieldName === "subcategory") return options.categories || [];
  return [];
}

function renderBatchEditSuggestions(input, open = false) {
  const item = input.closest(".batchItem");
  if (!item) return;
  const fieldName = input.dataset.batchField;
  const panel = input.closest(".suggestField")?.querySelector(".batchSuggestPanel");
  if (!panel) return;
  const suggestions = rankLookupSuggestions(batchEditOptions(item.dataset.batchType, fieldName), input.value);
  panel.innerHTML = suggestions.map((value) => `
    <button type="button" class="suggestOption" data-value="${text(value)}">${text(value)}</button>
  `).join("");
  document.querySelectorAll(".batchSuggestPanel").forEach((other) => {
    if (other !== panel) other.classList.remove("active");
  });
  closeSurveySuggestions();
  panel.classList.toggle("active", open && suggestions.length > 0);
}

function lookupMasterId() {
  return state.latestMaster?.id || "";
}

function lookupMasterName() {
  return state.latestMaster ? `目前最新總表 ${state.latestMaster.original_name}` : "";
}

function nextBatchType() {
  if (state.surveyBatches["原物料"].length && state.batchStatus["原物料"] !== "applied") return "原物料";
  if (state.surveyBatches["包材"].length && state.batchStatus["包材"] !== "applied") return "包材";
  return "";
}

function batchStatusText(sheetName) {
  const count = state.surveyBatches[sheetName].length;
  if (state.batchStatus[sheetName] === "applied") return "已套用";
  if (state.currentBatch === sheetName && state.batchStatus[sheetName] === "processing") return "比對中";
  if (state.currentBatch === sheetName && state.jobId) return "已比對，待確認套用";
  if (!count) return "未加入";
  if (sheetName === "包材" && state.surveyBatches["原物料"].length && state.batchStatus["原物料"] !== "applied") return "等待原物料完成";
  return "待比對";
}

function renderBatchSummary(sheetName) {
  const rows = state.surveyBatches[sheetName];
  const title = `目前上傳的${sheetName}調查表`;
  const status = batchStatusText(sheetName);
  if (!rows.length) {
    const stats = progressStats(sheetName);
    return `
      <div class="batchSummaryCard">
        <div class="batchSummaryHeader"><strong>${title}</strong><span class="tag">${status}</span></div>
        <div class="muted">${stats.completed ? `本輪已完成 ${stats.completed} 份${sheetName}調查表。` : `尚未加入${sheetName}調查表。`}</div>
      </div>
    `;
  }
  return `
    <div class="batchSummaryCard">
      <div class="batchSummaryHeader"><strong>${title}</strong><span class="tag same">${rows.length} 份｜${status}</span></div>
      <div class="batchSummaryList">
        ${rows.map((row) => `
          <div class="batchSummaryItem">
            <strong>${text(row.original_name)}</strong>
            <span>${supplierFieldLabel(sheetName)}：${text(supplierDisplayValue(row) || "未填")}</span>
            <span>品項名稱：${text(row.material_name || "未填")}</span>
            ${row.subcategory ? `<span>細項分類：${text(row.subcategory)}</span>` : ""}
            ${row.format_warning && row.format_warning !== "格式需人工確認" ? `<span class="warningText">${text(row.format_warning)}</span>` : ""}
          </div>
        `).join("")}
      </div>
    </div>
  `;
}

function progressStats(sheetName) {
  const progress = state.roundProgress[sheetName] || { total: 0, completed: 0 };
  const isActive = state.currentBatch === sheetName && ["processing", "confirming"].includes(state.batchStatus[sheetName]);
  const waiting = isActive || state.batchStatus[sheetName] === "applied" ? 0 : state.surveyBatches[sheetName].length;
  const total = Math.max(progress.total, progress.completed, progress.completed + waiting);
  const completed = Math.min(progress.completed, total);
  const percent = total ? Math.round((completed / total) * 100) : 0;
  return { total, completed, waiting, percent };
}

function renderProgressChip(sheetName) {
  const stats = progressStats(sheetName);
  const statusText = stats.total ? `${stats.completed}/${stats.total}${stats.waiting ? `，待 ${stats.waiting}` : ""}` : "未加入";
  const status = batchStatusText(sheetName);
  return `
    <div class="roundProgressChip">
      <div class="roundProgressMeta">
        <strong>${sheetName}</strong>
        <span>${statusText}｜${text(status)}</span>
      </div>
      <div class="progressTrack" aria-label="${sheetName}進度 ${stats.percent}%">
        <span style="width: ${stats.percent}%"></span>
      </div>
    </div>
  `;
}

function renderBatchMini(sheetName) {
  const rows = state.surveyBatches[sheetName];
  const stats = progressStats(sheetName);
  const names = rows.slice(0, 2).map((row) => row.original_name).join("、");
  const more = rows.length > 2 ? ` 等 ${rows.length} 份` : "";
  const doneText = stats.completed && !rows.length ? `已完成 ${stats.completed} 份` : "";
  return `
    <div class="batchMini">
      <strong>${sheetName}</strong>
      <span>${text(names ? names + more : doneText || "尚未加入")}</span>
    </div>
  `;
}

function renderVersionManagement() {
  const latest = state.latestMaster;
  $("latestVersionName").textContent = latest?.original_name || "尚未選定";
  $("latestVersionDownload").classList.toggle("hidden", !latest);
  if (latest) $("latestVersionDownload").href = `/download-file/${latest.id}`;
}

function renderBatchProgress() {
  if (!$("batchProgress")) return;
  const next = nextBatchType();
  const confirming = Boolean(state.jobId);
  const rawStats = progressStats("原物料");
  const packageStats = progressStats("包材");
  const total = rawStats.total + packageStats.total;
  const completed = rawStats.completed + packageStats.completed;
  const percent = total ? Math.round((completed / total) * 100) : 0;
  const totalStatus = total ? `${completed} / ${total} 份完成` : "尚未開始";
  const nextText = confirming ? "先確認並套用差異" : (next ? `按開始比對，系統將處理${next}` : "");
  const rawCount = state.surveyBatches["原物料"].length;
  const packageCount = state.surveyBatches["包材"].length;
  $("batchProgress").innerHTML = `
    <div class="roundProgressPanel compactRoundProgress">
      <div class="roundProgressHeader">
        <div>
          <span>本輪進度</span>
          <strong>${totalStatus}</strong>
        </div>
        <span class="tag same">${percent}%</span>
      </div>
      <div class="roundProgressChips">
        ${renderProgressChip("原物料")}
        ${renderProgressChip("包材")}
      </div>
      <div class="roundFileSummary">
        <span>本輪待比對：原物料 ${rawCount} 份｜包材 ${packageCount} 份</span>
        <button id="toggleRoundFilesBtn" type="button">${state.roundFilesExpanded ? "收合本輪檔案" : "查看本輪檔案"}</button>
      </div>
    </div>
    ${nextText ? `<div class="stateLine"><span>下一步</span><strong>${text(nextText)}</strong></div>` : ""}
  `;
  $("toggleRoundFilesBtn").addEventListener("click", () => {
    state.roundFilesExpanded = !state.roundFilesExpanded;
    renderBatchProgress();
    renderRoundBatchDetails();
  });
  $("compareBtn").textContent = confirming ? "差異確認中" : "開始比對";
  $("compareBtn").disabled = confirming || !next;
  $("cancelJobBtn").classList.toggle("hidden", !(confirming || state.workingMaster));
  $("finalizeBtn").classList.toggle("hidden", !state.workingMaster || confirming);
}

function updateLookupSheetUi() {
  const isPackage = $("lookupSheet").value === "包材";
  $("lookupCategoryLabel").textContent = isPackage ? "品項分類（包材免填）" : "品項分類";
  $("lookupCategory").placeholder = isPackage ? "包材免填" : "可輸入分類關鍵字";
  $("lookupCategory").disabled = isPackage;
  if (isPackage) {
    $("lookupCategory").value = "";
    closeLookupSuggestions("lookupCategory");
  }
}

function setLookupSource(source) {
  state.lookupSource = source;
  document.querySelectorAll("input[name='lookupSource']").forEach((input) => { input.checked = input.value === source; });
  if ($("lookupExistingPanel")) $("lookupExistingPanel").classList.toggle("active", source === "existing");
  if ($("lookupUploadPanel")) $("lookupUploadPanel").classList.toggle("active", source === "upload");
  if (source === "current") {
    state.lookupMasterId = "";
    state.lookupMasterName = "";
  }
  renderSelectedFiles();
}

const lookupFieldMap = {
  lookupSupplierCode: { panelId: "lookupSupplierCodeSuggest", optionKey: "supplier_codes" },
  lookupSupplier: { panelId: "lookupSupplierSuggest", optionKey: "suppliers" },
  lookupMaterial: { panelId: "lookupMaterialSuggest", optionKey: "materials" },
  lookupCategory: { panelId: "lookupCategorySuggest", optionKey: "categories" },
  lookupSubstance: { panelId: "lookupSubstanceSuggest", optionKey: "substances" },
};

const surveyFieldMap = {
  rawSurveySupplierCode: { panelId: "rawSurveySupplierSuggest", sheetName: "原物料", optionKey: "supplier_codes" },
  rawSurveyMaterial: { panelId: "rawSurveyMaterialSuggest", sheetName: "原物料", optionKey: "materials" },
  rawSurveySubcategory: { panelId: "rawSurveyCategorySuggest", sheetName: "原物料", optionKey: "categories" },
  packageSurveySupplier: { panelId: "packageSurveySupplierSuggest", sheetName: "包材", optionKey: "supplier_mix" },
  packageSurveyMaterial: { panelId: "packageSurveyMaterialSuggest", sheetName: "包材", optionKey: "materials" },
};

function normalizeLookupText(value) {
  return (value ?? "").toString()
    .replace(/\u3000/g, " ")
    .replace(/\s+/g, " ")
    .replace(/\s*([()（）])\s*/g, "$1")
    .trim()
    .toLowerCase();
}

function displayLookupText(value) {
  return (value ?? "").toString()
    .replace(/\u3000/g, " ")
    .replace(/\s+/g, " ")
    .replace(/\s*([()（）])\s*/g, "$1")
    .trim();
}

function uniqueLookupValues(values) {
  const seen = new Set();
  const result = [];
  (values || []).forEach((value) => {
    const displayValue = displayLookupText(value);
    const key = normalizeLookupText(displayValue);
    if (!key || seen.has(key)) return;
    seen.add(key);
    result.push(displayValue);
  });
  return result;
}

function rankLookupSuggestions(values, query) {
  const normalizedQuery = normalizeLookupText(query);
  const uniqueValues = uniqueLookupValues(values);
  const limit = 500;
  if (!normalizedQuery) return uniqueValues.slice(0, limit);

  return uniqueValues
    .map((value, index) => {
      const normalized = normalizeLookupText(value);
      if (normalized.startsWith(normalizedQuery)) return { value, rank: 0, index };
      if (normalized.includes(normalizedQuery)) return { value, rank: 1, index };
      return null;
    })
    .filter(Boolean)
    .sort((a, b) => a.rank - b.rank || a.index - b.index)
    .slice(0, limit)
    .map((item) => item.value);
}

function closeSuggestionPanels(fieldMap, exceptId = "") {
  Object.values(fieldMap).forEach((field) => {
    if (field.panelId !== exceptId) $(field.panelId).classList.remove("active");
  });
}

function renderSuggestions(inputId, fieldMap, optionsProvider, open = false) {
  const field = fieldMap[inputId];
  if (!field) return;
  const panel = $(field.panelId);
  const options = optionsProvider(field);
  const suggestions = rankLookupSuggestions(options, $(inputId).value);
  panel.innerHTML = suggestions.map((value) => `
    <button type="button" class="suggestOption" data-value="${text(value)}">${text(value)}</button>
  `).join("");
  panel.classList.toggle("active", open && suggestions.length > 0);
  closeSuggestionPanels(fieldMap, field.panelId);
}

function closeLookupSuggestions(exceptId = "") {
  closeSuggestionPanels(lookupFieldMap, exceptId);
}

function closeSurveySuggestions(exceptId = "") {
  closeSuggestionPanels(surveyFieldMap, exceptId);
}

function renderLookupSuggestions(inputId, open = false) {
  renderSuggestions(inputId, lookupFieldMap, (field) => state.lookupOptions[field.optionKey], open);
}

function surveyOptionsForField(field) {
  const options = state.surveyOptions[field.sheetName] || {};
  if (field.optionKey === "supplier_mix") return [...(options.suppliers || []), ...(options.supplier_codes || [])];
  return options[field.optionKey] || [];
}

function renderSurveySuggestions(inputId, open = false) {
  renderSuggestions(inputId, surveyFieldMap, surveyOptionsForField, open);
}

function renderAllLookupSuggestions() {
  Object.keys(lookupFieldMap).forEach((inputId) => renderLookupSuggestions(inputId, false));
}

function renderAllSurveySuggestions() {
  Object.keys(surveyFieldMap).forEach((inputId) => renderSurveySuggestions(inputId, false));
}

function emptyFilterOptions() {
  return { supplier_codes: [], suppliers: [], materials: [], categories: [], substances: [] };
}

async function loadFilterOptions(masterId, masterSheet) {
  return api("/api/master-filter-options", {
    method: "POST",
    body: JSON.stringify({ master_file_id: masterId, master_sheet: masterSheet }),
    headers: { "Content-Type": "application/json" },
  });
}

async function loadLookupOptions() {
  const masterId = lookupMasterId();
  updateLookupSheetUi();
  if (!masterId) {
    state.lookupOptions = emptyFilterOptions();
    renderAllLookupSuggestions();
    $("lookupSummary").innerHTML = `<span class="muted">請先選擇查詢總表來源，或上傳新的查詢總表。</span>`;
    return;
  }
  $("lookupSummary").innerHTML = `<strong>正在讀取總表篩選選項...</strong>`;
  const data = await loadFilterOptions(masterId, $("lookupSheet").value);
  state.lookupOptions = {
    supplier_codes: data.supplier_codes || [],
    suppliers: data.suppliers || [],
    materials: data.materials || [],
    categories: data.categories || [],
    substances: data.substances || [],
  };
  renderAllLookupSuggestions();
  if (!state.lookupRows.length) {
    $("lookupSummary").innerHTML = `
      <strong>篩選選項已載入</strong>
      ${text(data.file_name)}｜${text(data.sheet)}｜
      供應商代號 ${data.supplier_codes.length} 筆，供應商 ${data.suppliers.length} 筆，品項 ${data.materials.length} 筆，分類 ${data.categories.length} 筆，化學物質 ${data.substances.length} 筆
    `;
  }
}

async function loadSurveyOptions() {
  let master = effectiveMaster();
  if (!master?.id) {
    try {
      const data = await api("/api/dashboard");
      state.latestMaster = data.version_state?.latest_master || null;
      state.workingMaster = data.version_state?.working_master || null;
      master = effectiveMaster();
    } catch (err) {
      master = null;
    }
  }
  if (!master?.id) {
    state.surveyOptions = { "原物料": emptyFilterOptions(), "包材": emptyFilterOptions() };
    state.surveyOptionsMasterId = "";
    renderAllSurveySuggestions();
    return;
  }
  if (state.surveyOptionsMasterId === master.id) return;
  const [rawData, packageData] = await Promise.all([
    loadFilterOptions(master.id, "原物料"),
    loadFilterOptions(master.id, "包材"),
  ]);
  state.surveyOptions = {
    "原物料": {
      supplier_codes: rawData.supplier_codes || [],
      suppliers: rawData.suppliers || [],
      materials: rawData.materials || [],
      categories: rawData.categories || [],
      substances: rawData.substances || [],
    },
    "包材": {
      supplier_codes: packageData.supplier_codes || [],
      suppliers: packageData.suppliers || [],
      materials: packageData.materials || [],
      categories: packageData.categories || [],
      substances: packageData.substances || [],
    },
  };
  state.surveyOptionsMasterId = master.id;
}

async function showSurveySuggestions(inputId) {
  try {
    await loadSurveyOptions();
  } catch (err) {
    return;
  }
  renderSurveySuggestions(inputId, true);
}

function surveyPayload() {
  const master = effectiveMaster();
  const sheetName = nextBatchType();
  const surveyItems = state.surveyBatches[sheetName].map((row) => ({ ...row }));
  return {
    master_file_id: master?.id || "",
    survey_file_id: state.surveyId,
    survey_items: surveyItems,
    master_sheet: sheetName,
    supplier_code: "",
    material_name: "",
    survey_version: "",
  };
}

async function compare() {
  const sheetName = nextBatchType();
  if (!sheetName) throw new Error("目前沒有待比對批次");
  const rows = state.surveyBatches[sheetName];
  if (!effectiveMaster()?.id || !rows.length) throw new Error(`請先選定比對基準總表並加入${sheetName}調查表批次`);
  const missing = rows.find((row) => !row.supplier_code.trim());
  if (missing) throw new Error(`${missing.original_name} 未填${supplierFieldLabel(sheetName)}`);
  const missingMaterial = rows.find((row) => !row.material_name.trim());
  if (missingMaterial) throw new Error(`${missingMaterial.original_name} 未填品項/原料名稱`);
  const body = surveyPayload();
  state.currentBatch = sheetName;
  state.batchStatus[sheetName] = "processing";
  state.roundProgress[sheetName].total = Math.max(state.roundProgress[sheetName].total, state.roundProgress[sheetName].completed + rows.length);
  renderBatchProgress();
  $("compareBtn").disabled = true;
  $("compareBtn").textContent = "比對中...";
  $("compareSummary").innerHTML = `<strong>比對中...</strong><div class="muted">正在解析${text(sheetName)}調查表並比對總表，請稍候。</div>`;
  let data;
  try {
    data = await api("/api/compare", { method: "POST", body: JSON.stringify(body), headers: { "Content-Type": "application/json" } });
  } catch (err) {
    state.currentBatch = "";
    state.batchStatus[sheetName] = "pending";
    renderBatchProgress();
    throw err;
  }
  state.jobId = data.job_id;
  state.batchStatus[sheetName] = "confirming";
  state.roundProgress[sheetName].completed = Math.max(
    state.roundProgress[sheetName].completed,
    state.roundProgress[sheetName].total
  );
  state.page = 1;
  const s = data.summary;
  state.summary = s;
  renderDiffCounts();
  renderDiffPanel();
  renderBatchProgress();
  $("compareSummary").innerHTML = `
    <strong>比對完成</strong>
    舊版 ${s.old_count} 筆，新版 ${s.new_count} 筆；
    <span class="tag added">新增 ${s.added}</span>
    <span class="tag modified">修改 ${s.modified}</span>
    <span class="tag same">相同 ${s.same}</span>
    <span class="tag conflict">衝突 ${s.conflict}</span>
    <div class="muted">比對基準：${text(baselineLabel(effectiveMaster()))}</div>
    <div class="muted">批次類型：${text(sheetName)}；判斷規則：來源、化學物質名稱、CAS 主號與 Level；調查表未列出的總表細項不視為刪除。</div>
    <div class="muted">橫向供應品項：${s.batch ? `批次會依每份檔案的${supplierFieldLabel(sheetName)}與品項名稱更新或新增右側欄位` : (s.horizontal_new_column ? "套用時會新增右側欄位" : "已存在或資料不足，套用時不新增欄位")}</div>
    ${s.warnings?.length ? `<div class="muted">警告：${s.warnings.map(text).join("；")}</div>` : ""}
  `;
  await loadResults();
  await loadDashboard();
  renderBatchProgress();
}

function renderDiffCounts() {
  const s = state.summary || {};
  $("diffCounts").innerHTML = `
    <span class="tag added">新增 ${Number(s.added || 0)}</span>
    <span class="tag modified">修改 ${Number(s.modified || 0)}</span>
    <span class="tag same">相同 ${Number(s.same || 0)}</span>
    <span class="tag conflict">衝突 ${Number(s.conflict || 0)}</span>
  `;
}

async function loadResults() {
  if (!state.jobId) return;
  const type = $("typeFilter").value;
  const pageSize = Number($("pageSize").value || 50);
  state.pageSize = pageSize;
  const data = await api(`/api/jobs/${state.jobId}/results?type=${type}&page=${state.page}&page_size=${pageSize}`);
  state.totalResults = data.total;
  $("resultsBody").innerHTML = data.rows.map((r) => {
    const source = Object.keys(r.new).length ? r.new : r.old;
    const changes = r.changes.map((c) => `${text(c.field)}: ${text(c.old)} → ${text(c.new)}`).join("<br>");
    const level = r.new["等級"] || r.old["等級"] || "";
    return `<tr>
      <td><input type="checkbox" class="rowCheck" value="${r.id}" data-diff-type="${r.diff_type}" ${r.diff_type === "conflict" ? "disabled" : ""}></td>
      <td><span class="tag ${r.diff_type}">${labels[r.diff_type] || r.diff_type}</span></td>
      <td>${text(r.action)}</td>
      <td>${text(source["調查表類型"])}</td>
      <td>${text(source["供應商代號"] || source["供應商"])}</td>
      <td>${text(source["物料名稱"])}</td>
      <td>${text(source["化學物質"])}</td>
      <td>${text(source.CAS)}</td>
      <td>${text(level)}</td>
      <td>${changes || "-"}</td>
      <td>${text(source["來源檔案"])}</td>
    </tr>`;
  }).join("");
  state.page = data.page;
  $("selectAll").checked = false;
  updatePagerControls();
  renderDiffPanel();
}

async function autoConfirm() {
  if (!state.jobId) return alert("尚未建立比對工作");
  if (!confirm("將一鍵確認所有新增與修改：新增資料會加入總表，修改資料會採用調查表內容覆蓋總表。確定執行？")) return;
  const data = await api(`/api/jobs/${state.jobId}/auto-confirm`, { method: "POST" });
  alert(`已確認新增 ${data.added} 筆、修改 ${data.modified} 筆`);
  await loadResults();
}

async function confirmAction(action) {
  const selected = [...document.querySelectorAll(".rowCheck:checked")];
  const ids = selected.map((i) => i.value);
  if (!ids.length) return alert("請先勾選差異資料");
  const invalidTypes = [...new Set(selected.map((i) => i.dataset.diffType).filter((type) => !(allowedActionsByType[type] || []).includes(action)))];
  if (invalidTypes.length) {
    const names = invalidTypes.map((type) => labels[type] || type).join("、");
    return alert(`這個確認方式不能套用在「${names}」資料，請改用符合差異類型的按鈕。`);
  }
  await api(`/api/jobs/${state.jobId}/confirm`, {
    method: "POST",
    body: JSON.stringify({ result_ids: ids, action }),
    headers: { "Content-Type": "application/json" },
  });
  await loadResults();
}

function updatePagerControls() {
  const totalPages = Math.max(1, Math.ceil((state.totalResults || 0) / (state.pageSize || 50)));
  const currentPage = Math.min(state.page || 1, totalPages);
  $("pager").textContent = state.jobId ? `第 ${currentPage} / ${totalPages} 頁，每頁 ${state.pageSize} 筆，共 ${state.totalResults} 筆` : "";
  $("prevPageBtn").disabled = !state.jobId || currentPage <= 1;
  $("nextPageBtn").disabled = !state.jobId || currentPage >= totalPages;
}

function changeResultsPage(delta) {
  const totalPages = Math.max(1, Math.ceil((state.totalResults || 0) / (state.pageSize || 50)));
  const nextPage = Math.min(Math.max(1, (state.page || 1) + delta), totalPages);
  if (nextPage === state.page) return;
  state.page = nextPage;
  loadResults().catch((err) => alert(err.message));
}

async function applyJob() {
  if (!state.jobId) return alert("尚未建立比對工作");
  if (!confirm("即將套用目前這批已確認差異，產生本輪暫存更動版與 Search Report。確定套用本批修改？")) return;
  const applyingBatch = state.currentBatch;
  if (applyingBatch) {
    state.batchStatus[applyingBatch] = "confirming";
    renderBatchProgress();
  }
  const data = await api(`/api/jobs/${state.jobId}/apply`, { method: "POST" });
  const completedBatch = state.currentBatch || applyingBatch;
  if (completedBatch) {
    const progress = state.roundProgress[completedBatch] || { total: 0, completed: 0 };
    const waitingCount = state.surveyBatches[completedBatch].length;
    state.roundProgress[completedBatch].total = Math.max(progress.total || 0, progress.completed || 0, waitingCount);
    state.roundProgress[completedBatch].completed = state.roundProgress[completedBatch].total;
  }
  if (completedBatch) clearCompletedBatch(completedBatch);
  state.jobId = "";
  state.currentBatch = "";
  state.summary = null;
  state.totalResults = 0;
  state.page = 1;
  $("resultsBody").innerHTML = "";
  $("selectAll").checked = false;
  renderDiffCounts();
  updatePagerControls();
  const next = nextBatchType();
  const nextMessage = next
    ? `${completedBatch}批次已套用。接下來會使用暫存更動版繼續比對${next}。`
    : "目前已加入的檔案都已套用。請完成本輪更新並定版。";
  const nextButton = next
    ? `<button class="primary inlineNextBtn" type="button" data-next-step="compare">繼續比對</button>`
    : `<button class="primary inlineNextBtn" type="button" data-next-step="finalize">完成本輪更新並定版</button>`;
  $("compareSummary").innerHTML += `<div class="notice"><strong>已套用本批修改並產生暫存更動版：</strong>${text(data.updated_file)}。${text(nextMessage)}<br>
    <a href="/download/${encodeURIComponent(data.updated_file)}">${text(data.updated_file)}</a>、
    <a href="/download/${encodeURIComponent(data.report_file)}">${text(data.report_file)}</a>
    <div class="inlineNextAction">${nextButton}</div>
  </div>`;
  await loadDashboard();
  renderSelectedFiles();
  renderBatchProgress();
}

async function cancelJob() {
  const hasJob = Boolean(state.jobId);
  const hasWorking = Boolean(state.workingMaster);
  if (!hasJob && !hasWorking) return alert("目前沒有可取消的比對工作");
  const message = hasWorking
    ? "確定取消本次比對？尚未定版的暫存更動版會一併刪除，本輪已加入的調查表也會清空。"
    : "確定取消本次比對？系統會刪除這次比對的差異暫存結果，尚未套用任何修改，調查表會回到待比對清單。";
  if (!confirm(message)) return;
  if (hasJob) {
    await api(`/api/jobs/${state.jobId}/cancel`, { method: "POST" });
  }
  const batch = state.currentBatch;
  if (batch) {
    const progress = state.roundProgress[batch] || { total: 0, completed: 0 };
    const batchCount = state.surveyBatches[batch].length;
    state.roundProgress[batch].completed = Math.max(0, (progress.completed || 0) - batchCount);
    state.batchStatus[batch] = "pending";
  }
  if (hasWorking) {
    await clearWorkingVersion();
    state.surveyBatches = { "原物料": [], "包材": [] };
    state.batchStatus = { "原物料": "pending", "包材": "pending" };
    resetRoundProgress();
    clearSurveyForm("原物料");
    clearSurveyForm("包材");
  }
  state.jobId = "";
  state.currentBatch = "";
  state.summary = null;
  state.totalResults = 0;
  state.page = 1;
  $("compareSummary").innerHTML = `<div class="notice"><strong>本次比對已取消。</strong>${hasWorking ? "未定版暫存更動版已清除。" : "尚未套用任何修改，調查表已回到待比對清單。"}</div>`;
  $("resultsBody").innerHTML = "";
  $("selectAll").checked = false;
  renderDiffCounts();
  renderDiffPanel();
  updatePagerControls();
  renderBatchLists();
  renderBatchProgress();
  await loadDashboard();
}

async function finalizeWorkingVersion() {
  if (!confirm("確定完成本輪更新並定版？目前暫存更動版將成為最新總表。之後若未重新選用總表，系統會預設使用此版本進行比對。")) return;
  const data = await api("/api/finalize-working-version", { method: "POST" });
  alert(data.message);
  state.masterId = "";
  state.masterName = "";
  state.jobId = "";
  state.currentBatch = "";
  state.summary = null;
  state.totalResults = 0;
  state.page = 1;
  state.workingMaster = null;
  state.surveyBatches = { "原物料": [], "包材": [] };
  state.batchStatus = { "原物料": "pending", "包材": "pending" };
  resetRoundProgress();
  $("resultsBody").innerHTML = "";
  $("selectAll").checked = false;
  $("compareSummary").innerHTML = `<div class="notice"><strong>已完成本輪更新並定版。</strong>目前頁面已回到下一輪開始前狀態。</div>`;
  renderDiffCounts();
  renderDiffPanel();
  updatePagerControls();
  await loadDashboard();
  renderBatchLists();
  renderBatchProgress();
}

function renderDiffPanel() {
  const hasJob = Boolean(state.jobId);
  $("diffPanel").classList.toggle("collapsed", !hasJob);
  $("typeFilter").disabled = !hasJob;
  $("pageSize").disabled = !hasJob;
  $("autoConfirmBtn").disabled = !hasJob;
  $("applyBtn").disabled = !hasJob;
  document.querySelectorAll(".actions button[data-action]").forEach((button) => { button.disabled = !hasJob; });
  if (!hasJob) {
    $("resultsBody").innerHTML = "";
    $("pager").textContent = "";
  }
}

async function setLatest(fileId) {
  if (!confirm("此動作會把這份總表定為系統預設最新版。之後若未手動指定其他總表，系統會用它進行比對。確定嗎？")) return;
  const data = await api(`/api/files/${fileId}/set-latest`, { method: "POST" });
  alert(`${data.message}。接下來若未指定其他總表，系統會預設使用此版本進行比對。`);
  state.masterId = "";
  state.masterName = "";
  await loadDashboard();
}

async function deleteFile(fileId, name) {
  if (state.latestMaster?.id === fileId) return alert("目前最新總表不能直接刪除；請先將其他總表定為最新版。");
  if (state.workingMaster?.id === fileId) return alert("本輪暫存更動版請用「取消本次比對」清除。");
  if (isSurveyInBatch(fileId)) return alert("已加入本輪的調查表不能直接刪除；請先從批次清單移除或取消本次比對。");
  if (!confirm(`確定刪除檔案「${name}」？若此檔案已被未套用的比對工作引用，該工作會標示為檔案已刪，不能再套用。`)) return;
  await api(`/api/files/${fileId}`, { method: "DELETE" });
  if (state.masterId === fileId) {
    state.masterId = "";
    state.masterName = "";
  }
  if (state.surveyId === fileId) {
    state.surveyId = "";
    state.surveyName = "";
  }
  renderSelectedFiles();
  await loadDashboard();
}

function lookupPayload() {
  return {
    master_file_id: lookupMasterId(),
    master_sheet: $("lookupSheet").value,
    supplier_code: $("lookupSupplierCode").value,
    supplier: $("lookupSupplier").value,
    material_name: $("lookupMaterial").value,
    category: $("lookupCategory").value,
    substance: $("lookupSubstance").value,
    include_empty: false,
    page: state.lookupPage,
    page_size: state.lookupPageSize,
  };
}

async function lookupMaster(page = 1) {
  if (!lookupMasterId()) throw new Error("請先選擇查詢總表來源，或上傳新的查詢總表");
  state.lookupPage = page;
  $("lookupSummary").innerHTML = "<strong>查詢中...</strong>";
  const data = await api("/api/master-search", {
    method: "POST",
    body: JSON.stringify(lookupPayload()),
    headers: { "Content-Type": "application/json" },
  });
  $("lookupSummary").innerHTML = `
    <strong>查詢完成</strong>
    ${text(data.file_name)}｜${text(data.sheet)}｜
    細項 ${data.total_rows ?? data.rows.length} 筆
  `;
  state.lookupRows = data.rows || [];
  state.lookupTotalRows = data.total_rows ?? state.lookupRows.length;
  state.lookupPage = data.page || page;
  state.lookupPageSize = data.page_size || state.lookupPageSize;
  renderLookupRows();
}

function renderLookupRows() {
  const totalRows = state.lookupTotalRows || state.lookupRows.length;
  const totalPages = Math.max(1, Math.ceil(totalRows / state.lookupPageSize));
  state.lookupPage = Math.min(Math.max(1, state.lookupPage), totalPages);
  const start = totalRows ? (state.lookupPage - 1) * state.lookupPageSize + 1 : 0;
  const pageRows = state.lookupRows;
  const end = totalRows ? Math.min(totalRows, start + pageRows.length - 1) : 0;

  $("lookupCount").textContent = totalRows ? `顯示 ${start}-${end} 筆` : "顯示 0 筆";
  $("lookupPager").textContent = totalRows ? `第 ${state.lookupPage} / ${totalPages} 頁，每頁 ${state.lookupPageSize} 筆，共 ${state.lookupTotalRows} 筆` : "";
  $("lookupPrevPageBtn").disabled = state.lookupPage <= 1 || totalRows === 0;
  $("lookupNextPageBtn").disabled = state.lookupPage >= totalPages || totalRows === 0;
  $("lookupRows").innerHTML = pageRows.map((r) => `
    <tr>
      <td>${text(r.sheet)}</td>
      <td>${text(r.supplier_code)}</td>
      <td>${text(r.supplier)}</td>
      <td>${text(r.material_name)}</td>
      <td>${text(r.category || "-")}</td>
      <td>${text(r.substance)}</td>
      <td>${text(r.cas)}</td>
      <td>${text(r.level)}</td>
      <td>${text(r.content || "-")}</td>
    </tr>
  `).join("");
}

function changeLookupPage(delta) {
  const totalPages = Math.max(1, Math.ceil((state.lookupTotalRows || 0) / state.lookupPageSize));
  const nextPage = Math.min(Math.max(1, state.lookupPage + delta), totalPages);
  if (nextPage === state.lookupPage) return;
  lookupMaster(nextPage).catch((err) => alert(err.message));
}

function clearLookup() {
  ["lookupSupplierCode", "lookupSupplier", "lookupMaterial", "lookupCategory", "lookupSubstance"].forEach((id) => { $(id).value = ""; });
  renderAllLookupSuggestions();
  closeLookupSuggestions();
  state.lookupRows = [];
  state.lookupTotalRows = 0;
  state.lookupPage = 1;
  $("lookupSummary").innerHTML = "";
  $("lookupRows").innerHTML = "";
  $("lookupCount").textContent = "";
  $("lookupPager").textContent = "";
  $("lookupPrevPageBtn").disabled = true;
  $("lookupNextPageBtn").disabled = true;
}

function fileStatus(f) {
  if (state.latestMaster?.id === f.id) return "目前最新總表";
  return "";
}

function fileSortRank(f) {
  if (state.latestMaster?.id === f.id) return 0;
  if (state.justUploadedMasterId === f.id) return 1;
  return 3;
}

function isSurveyInBatch(fileId) {
  return Object.values(state.surveyBatches).some((rows) => rows.some((row) => row.survey_file_id === fileId));
}

function masterFileTags(f) {
  const tags = [];
  if (state.latestMaster?.id === f.id) tags.push(`<span class="tag same">目前最新總表</span>`);
  if (state.justUploadedMasterId === f.id) tags.push(`<span class="tag fresh">剛剛上傳</span>`);
  return tags.join("");
}

function renderFilesList() {
  const total = state.files.length;
  const totalPages = Math.max(1, Math.ceil(total / state.filesPageSize));
  state.filesPage = Math.min(Math.max(1, state.filesPage), totalPages);
  const start = (state.filesPage - 1) * state.filesPageSize;
  const pageFiles = state.files.slice(start, start + state.filesPageSize);
  const end = total ? start + pageFiles.length : 0;
  $("filesPagerTop").textContent = total ? `顯示 ${start + 1}-${end} / ${total}` : "尚無檔案";
  $("filesPager").textContent = total ? `第 ${state.filesPage} / ${totalPages} 頁，每頁 ${state.filesPageSize} 筆` : "";
  $("filesPrevPageBtn").disabled = state.filesPage <= 1 || total === 0;
  $("filesNextPageBtn").disabled = state.filesPage >= totalPages || total === 0;
  $("filesList").innerHTML = pageFiles.map((f) => {
    const isLatest = state.latestMaster?.id === f.id;
    const isFresh = state.justUploadedMasterId === f.id;
    return `<div class="fileRow ${isLatest ? "selectedItem" : ""} ${isFresh ? "freshItem" : ""}">
      <div class="fileNameBlock">
        <strong>${text(f.original_name)}</strong>
        <span>${text(f.kind)}｜${text(f.uploaded_at)}</span>
      </div>
      <div class="fileTags">${masterFileTags(f)}</div>
      <a href="/download-file/${f.id}">下載</a>
      <div class="fileActions compact">
        <button type="button" data-set-latest="${f.id}">定最新版</button>
        <button type="button" data-delete-file="${f.id}" data-file-name="${text(f.original_name)}">刪除</button>
      </div>
    </div>`;
  }).join("");
  document.querySelectorAll("#filesList [data-set-latest]").forEach((button) => {
    button.addEventListener("click", () => setLatest(button.dataset.setLatest).catch((err) => alert(err.message)));
  });
  document.querySelectorAll("#filesList [data-delete-file]").forEach((button) => {
    button.addEventListener("click", () => deleteFile(button.dataset.deleteFile, button.dataset.fileName).catch((err) => alert(err.message)));
  });
}

function changeFilesPage(delta) {
  const totalPages = Math.max(1, Math.ceil((state.files.length || 0) / state.filesPageSize));
  const nextPage = Math.min(Math.max(1, state.filesPage + delta), totalPages);
  if (nextPage === state.filesPage) return;
  state.filesPage = nextPage;
  renderFilesList();
}

function renderSurveyFilesList() {
  const total = state.surveyFiles.length;
  const totalPages = Math.max(1, Math.ceil(total / state.surveyFilesPageSize));
  state.surveyFilesPage = Math.min(Math.max(1, state.surveyFilesPage), totalPages);
  const start = (state.surveyFilesPage - 1) * state.surveyFilesPageSize;
  const pageFiles = state.surveyFiles.slice(start, start + state.surveyFilesPageSize);
  $("surveyFilesSummary").innerHTML = `
    <strong>已上傳調查表：${total} 份</strong>
    <span>${state.surveyFilesExpanded ? "清單已展開" : "需要核對檔案時再展開"}</span>
  `;
  $("toggleSurveyFilesBtn").textContent = state.surveyFilesExpanded ? "收合清單" : "查看清單";
  $("surveyFilesList").classList.toggle("hidden", !state.surveyFilesExpanded);
  $("surveyFilesPagerBar").classList.toggle("hidden", !state.surveyFilesExpanded);
  $("surveyFilesPager").textContent = total ? `第 ${state.surveyFilesPage} / ${totalPages} 頁，每頁 ${state.surveyFilesPageSize} 筆` : "";
  $("surveyFilesPrevPageBtn").disabled = state.surveyFilesPage <= 1 || total === 0;
  $("surveyFilesNextPageBtn").disabled = state.surveyFilesPage >= totalPages || total === 0;
  $("surveyFilesList").innerHTML = `
    <div class="tableLikeHeader">
      <span>檔名</span>
      <span>上傳時間</span>
      <span>狀態</span>
      <span>動作</span>
    </div>
    ${pageFiles.map((f) => {
    const inBatch = isSurveyInBatch(f.id);
    return `<div class="tableLikeRow ${inBatch ? "selectedItem" : ""}">
      <strong>${text(f.original_name)}</strong>
      <span>${text(f.uploaded_at)}</span>
      <span class="tag ${inBatch ? "same" : ""}">${inBatch ? "已加入本輪" : "未加入本輪"}</span>
      <div class="fileActions compact">
        <a href="/download-file/${f.id}">下載</a>
        <button type="button" data-delete-file="${f.id}" data-file-name="${text(f.original_name)}" ${inBatch ? "disabled" : ""}>刪除</button>
      </div>
    </div>`;
  }).join("")}`;
  document.querySelectorAll("#surveyFilesList [data-delete-file]").forEach((button) => {
    button.addEventListener("click", () => deleteFile(button.dataset.deleteFile, button.dataset.fileName).catch((err) => alert(err.message)));
  });
}

function changeSurveyFilesPage(delta) {
  const totalPages = Math.max(1, Math.ceil((state.surveyFiles.length || 0) / state.surveyFilesPageSize));
  const nextPage = Math.min(Math.max(1, state.surveyFilesPage + delta), totalPages);
  if (nextPage === state.surveyFilesPage) return;
  state.surveyFilesPage = nextPage;
  renderSurveyFilesList();
}

async function loadDashboard() {
  const data = await api("/api/dashboard");
  (data.system_notices || []).forEach((notice) => {
    const key = `${notice.type || "notice"}:${notice.file_id || ""}:${notice.message || ""}`;
    if (!state.shownNoticeKeys.has(key)) {
      state.shownNoticeKeys.add(key);
      alert(notice.message || "系統狀態已自動修復，請確認目前最新總表。");
    }
  });
  state.latestMaster = data.version_state?.latest_master || null;
  state.workingMaster = data.version_state?.working_master || null;
  const files = [...data.files].sort((a, b) => {
    const rank = fileSortRank(a) - fileSortRank(b);
    if (rank) return rank;
    return (b.uploaded_at || "").localeCompare(a.uploaded_at || "");
  });
  const masterFiles = files.filter((f) => (f.kind === "master" || f.kind === "updated_master") && f.id !== state.workingMaster?.id);
  const surveyFiles = files.filter((f) => f.kind === "survey");
  state.files = masterFiles;
  state.surveyFiles = surveyFiles;
  $("versionText").textContent = data.app_version || "v0.7.1";
  if (!state.masterId && !state.latestMaster && !state.workingMaster && masterFiles.length === 1) {
    state.masterId = masterFiles[0].id;
    state.masterName = masterFiles[0].original_name;
  }
  renderSelectedFiles();
  renderBatchProgress();
  renderVersionManagement();
  renderDiffPanel();
  renderFilesList();
  renderSurveyFilesList();
  loadSurveyOptions().catch(() => {});
  $("versionsList").innerHTML = data.versions.map((v) => {
    const latestTag = state.latestMaster?.id === v.file_id ? `<span class="tag same">目前最新總表</span>` : "";
    return `<div class="listItem">
      <div class="listTitle"><strong>${text(v.version_no)}</strong>${latestTag}</div>
      <div class="muted">${v.created_at}｜${text(v.note || "")}</div>
    </div>`;
  }).join("");
}

$("masterForm").addEventListener("submit", async (e) => { e.preventDefault(); try { await upload(e.target, "master"); } catch (err) { alert(err.message); } });
$("rawSurveyForm").addEventListener("submit", async (e) => { e.preventDefault(); try { await uploadSurveyBatch(e.target, "原物料"); } catch (err) { alert(err.message); } });
$("packageSurveyForm").addEventListener("submit", async (e) => { e.preventDefault(); try { await uploadSurveyBatch(e.target, "包材"); } catch (err) { alert(err.message); } });
if ($("lookupMasterForm")) $("lookupMasterForm").addEventListener("submit", async (e) => { e.preventDefault(); try { await uploadLookupMaster(e.target); } catch (err) { alert(err.message); } });
document.querySelectorAll("#rawSurveyForm input[name='supplier_code'], #packageSurveyForm input[name='supplier_code']").forEach((input) => {
  input.addEventListener("paste", () => setTimeout(() => { input.value = cleanPastedSupplierCode(input.value); }, 0));
});
$("compareBtn").addEventListener("click", async () => { try { await compare(); } catch (err) { alert(err.message); } });
$("compareSummary").addEventListener("click", async (event) => {
  const button = event.target.closest("[data-next-step]");
  if (!button) return;
  try {
    if (button.dataset.nextStep === "compare") await compare();
    if (button.dataset.nextStep === "survey") activateTab("surveyPage");
    if (button.dataset.nextStep === "finalize") await finalizeWorkingVersion();
  } catch (err) {
    alert(err.message);
  }
});
$("autoConfirmBtn").addEventListener("click", () => autoConfirm().catch((err) => alert(err.message)));
$("finalizeBtn").addEventListener("click", () => finalizeWorkingVersion().catch((err) => alert(err.message)));
$("cancelJobBtn").addEventListener("click", () => cancelJob().catch((err) => alert(err.message)));
$("typeFilter").addEventListener("change", () => { state.page = 1; loadResults(); });
$("pageSize").addEventListener("change", () => { state.page = 1; loadResults(); });
$("selectAll").addEventListener("change", (e) => document.querySelectorAll(".rowCheck:not(:disabled)").forEach((c) => c.checked = e.target.checked));
$("prevPageBtn").addEventListener("click", () => changeResultsPage(-1));
$("nextPageBtn").addEventListener("click", () => changeResultsPage(1));
$("filesPrevPageBtn").addEventListener("click", () => changeFilesPage(-1));
$("filesNextPageBtn").addEventListener("click", () => changeFilesPage(1));
$("surveyFilesPrevPageBtn").addEventListener("click", () => changeSurveyFilesPage(-1));
$("surveyFilesNextPageBtn").addEventListener("click", () => changeSurveyFilesPage(1));
$("toggleSurveyFilesBtn").addEventListener("click", () => {
  state.surveyFilesExpanded = !state.surveyFilesExpanded;
  renderSurveyFilesList();
});
document.querySelectorAll("#rawSurveyForm input[type='file'], #packageSurveyForm input[type='file']").forEach((input) => {
  input.addEventListener("change", () => updateSurveyFileLabel(input.form));
});
document.querySelectorAll(".actions button[data-action]").forEach((b) => b.addEventListener("click", () => confirmAction(b.dataset.action).catch((err) => alert(err.message))));
$("applyBtn").addEventListener("click", () => applyJob().catch((err) => alert(err.message)));
$("lookupBtn").addEventListener("click", () => lookupMaster().catch((err) => alert(err.message)));
$("lookupClearBtn").addEventListener("click", clearLookup);
$("lookupPrevPageBtn").addEventListener("click", () => changeLookupPage(-1));
$("lookupNextPageBtn").addEventListener("click", () => changeLookupPage(1));
$("lookupSheet").addEventListener("change", () => loadLookupOptions().catch((err) => alert(err.message)));
Object.keys(lookupFieldMap).forEach((inputId) => {
  $(inputId).addEventListener("input", () => renderLookupSuggestions(inputId, true));
  $(inputId).addEventListener("focus", () => renderLookupSuggestions(inputId, true));
  $(lookupFieldMap[inputId].panelId).addEventListener("click", (event) => {
    const option = event.target.closest(".suggestOption");
    if (!option) return;
    $(inputId).value = option.dataset.value || "";
    renderLookupSuggestions(inputId, false);
  });
});
Object.keys(surveyFieldMap).forEach((inputId) => {
  $(inputId).addEventListener("input", () => showSurveySuggestions(inputId));
  $(inputId).addEventListener("focus", () => showSurveySuggestions(inputId));
  $(surveyFieldMap[inputId].panelId).addEventListener("click", (event) => {
    const option = event.target.closest(".suggestOption");
    if (!option) return;
    $(inputId).value = option.dataset.value || "";
    renderSurveySuggestions(inputId, false);
  });
});
document.addEventListener("click", (event) => {
  if (!event.target.closest(".lookupGrid")) closeLookupSuggestions();
  if (!event.target.closest("#surveyPage .suggestField")) {
    closeSurveySuggestions();
    document.querySelectorAll(".batchSuggestPanel").forEach((panel) => panel.classList.remove("active"));
  }
});
document.addEventListener("click", (event) => {
  const option = event.target.closest(".batchSuggestPanel .suggestOption");
  if (!option) return;
  const field = option.closest(".suggestField");
  const input = field?.querySelector("[data-batch-field]");
  if (!input) return;
  input.value = option.dataset.value || "";
  input.dispatchEvent(new Event("input", { bubbles: true }));
  option.closest(".batchSuggestPanel").classList.remove("active");
});
if ($("lookupExistingMaster")) $("lookupExistingMaster").addEventListener("change", () => {
  const selected = $("lookupExistingMaster").selectedOptions[0];
  state.lookupSource = "existing";
  state.lookupMasterId = $("lookupExistingMaster").value;
  state.lookupMasterName = selected ? selected.textContent.split("｜")[0] : "";
  setLookupSource("existing");
  loadLookupOptions().catch((err) => alert(err.message));
});
document.querySelectorAll("input[name='lookupSource']").forEach((input) => {
  input.addEventListener("change", () => {
    setLookupSource(input.value);
    loadLookupOptions().catch((err) => alert(err.message));
  });
});
document.querySelectorAll(".tab").forEach((button) => {
  button.addEventListener("click", () => activateTab(button.dataset.tabTarget));
});

const masterDropZone = $("masterDropZone");
const masterFileInput = $("masterForm").querySelector("input[type='file']");
["dragenter", "dragover"].forEach((eventName) => {
  masterDropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    masterDropZone.classList.add("dragOver");
  });
});
["dragleave", "drop"].forEach((eventName) => {
  masterDropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    masterDropZone.classList.remove("dragOver");
  });
});
masterDropZone.addEventListener("drop", async (event) => {
  const file = [...event.dataTransfer.files].find((item) => /\.(xlsx|xls|xlsm)$/i.test(item.name));
  if (!file) return alert("請拖曳 Excel 檔案（.xlsx、.xls、.xlsm）");
  try {
    await upload($("masterForm"), "master", file);
    masterFileInput.value = "";
  } catch (err) {
    alert(err.message);
  }
});

loadDashboard();
renderBatchLists();
