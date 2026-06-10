"use strict";

/* ---------- State ---------- */
let rawRows = [];      // full sheet as array-of-arrays (each = a row of cells)
let headers = [];      // resolved header labels for current header row
let dataRows = [];     // [{ __i, fields: [{label, value}] }] built from current config
let headerRowIdx = 0;  // index into rawRows used as header
let currentSessionId = null;        // id of the session being viewed/edited
let hiddenCols = new Set();         // column indices hidden from every form
let selectedForms = new Set();      // dataRow.__i values the user picked
let showSelectedOnly = false;       // limit view to selectedForms

const STORE_KEY = "rowsToForms.sessions.v1";

/* ---------- DOM ---------- */
const $ = (id) => document.getElementById(id);
const dropzone   = $("dropzone");
const fileInput  = $("fileInput");
const config     = $("config");
const headerSel  = $("headerRow");
const titleSel   = $("titleCol");
const searchInp  = $("search");
const formsEl    = $("forms");
const metaEl     = $("meta");
const emptyEl    = $("empty");
const toastEl    = $("toast");
const sessionsEl     = $("sessions");
const sessionsList   = $("sessionsList");
const clearSessions  = $("clearSessions");
const fieldsBtn      = $("fieldsBtn");
const fieldsPop      = $("fieldsPop");
const fieldsList     = $("fieldsList");
const fieldsSearch   = $("fieldsSearch");
const selectedOnlyBtn = $("selectedOnlyBtn");

/* ---------- Icons ---------- */
const ICON_COPY = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="11" height="11" rx="2"/><path d="M5 15V5a2 2 0 0 1 2-2h10"/></svg>';
const ICON_CHECK = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg>';
const ICON_CHEVRON = '<svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m6 9 6 6 6-6"/></svg>';
const ICON_FILE = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M14 3v4a1 1 0 0 0 1 1h4"/><path d="M17 21H7a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h7l5 5v11a2 2 0 0 1-2 2Z"/></svg>';
const ICON_TRASH = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6h18M8 6V4a1 1 0 0 1 1-1h6a1 1 0 0 1 1 1v2m2 0v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6"/></svg>';

/* ---------- File handling ---------- */
dropzone.addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", (e) => {
  if (e.target.files[0]) handleFile(e.target.files[0]);
});
["dragenter", "dragover"].forEach((ev) =>
  dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.add("drag"); })
);
["dragleave", "drop"].forEach((ev) =>
  dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.remove("drag"); })
);
dropzone.addEventListener("drop", (e) => {
  const f = e.dataTransfer.files[0];
  if (f) handleFile(f);
});

function handleFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.data ?? e.target.result, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      // header:1 => array of arrays; defval keeps empty cells as ""
      rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", blankrows: false, raw: false });
      if (!rawRows.length) { toast("That file looks empty."); return; }
      headerRowIdx = 0;
      hiddenCols = new Set();
      selectedForms = new Set();
      showSelectedOnly = false;
      currentSessionId = "s_" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
      buildHeaderRowOptions();
      onConfigChange();
      saveCurrentSession(file.name);
      showWorkspace();
    } catch (err) {
      console.error(err);
      toast("Couldn't read that file.");
    }
  };
  reader.readAsArrayBuffer(file);
}

/* ---------- Config: header row + title column ---------- */
function buildHeaderRowOptions() {
  // Offer the first N rows as candidate header rows (cap at 20 for sanity).
  const max = Math.min(rawRows.length, 20);
  headerSel.innerHTML = "";
  for (let i = 0; i < max; i++) {
    const preview = rawRows[i].filter((c) => c !== "").slice(0, 4).join(", ");
    const opt = document.createElement("option");
    opt.value = i;
    opt.textContent = `Row ${i + 1}${preview ? " — " + truncate(preview, 40) : ""}`;
    headerSel.appendChild(opt);
  }
  headerSel.value = String(headerRowIdx);
}

function buildTitleColOptions() {
  const prev = titleSel.value;
  titleSel.innerHTML = "";
  const noneOpt = document.createElement("option");
  noneOpt.value = "-1";
  noneOpt.textContent = "(Row number only)";
  titleSel.appendChild(noneOpt);
  headers.forEach((h, i) => {
    const opt = document.createElement("option");
    opt.value = i;
    opt.textContent = truncate(h, 40);
    titleSel.appendChild(opt);
  });
  // keep previous choice if still valid, else default to first column
  if (prev !== "" && Number(prev) < headers.length) titleSel.value = prev;
  else titleSel.value = headers.length ? "0" : "-1";
}

/* ---------- Field visibility popover ---------- */
function buildFieldsPopover() {
  fieldsList.innerHTML = "";
  headers.forEach((h, i) => {
    const item = document.createElement("label");
    item.className = "pop-item";
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = !hiddenCols.has(i);
    cb.addEventListener("change", () => {
      if (cb.checked) hiddenCols.delete(i);
      else hiddenCols.add(i);
      updateFieldsBtn();
      render();
      persistConfig();
    });
    const span = document.createElement("span");
    span.textContent = h;
    span.title = h;
    item.appendChild(cb);
    item.appendChild(span);
    item.dataset.label = h.toLowerCase();
    fieldsList.appendChild(item);
  });
  filterFields();
  updateFieldsBtn();
}

function filterFields() {
  const q = fieldsSearch.value.trim().toLowerCase();
  fieldsList.querySelectorAll(".pop-item").forEach((item) => {
    item.classList.toggle("hidden", !!q && !item.dataset.label.includes(q));
  });
}

fieldsSearch.addEventListener("input", filterFields);

function updateFieldsBtn() {
  const shown = headers.length - hiddenCols.size;
  fieldsBtn.textContent =
    (shown === headers.length ? "All fields" : `${shown} of ${headers.length} fields`) + " ▾";
}

function setAllFields(visible) {
  hiddenCols = visible ? new Set() : new Set(headers.map((_, i) => i));
  buildFieldsPopover();
  render();
  persistConfig();
}

fieldsBtn.addEventListener("click", (e) => {
  e.stopPropagation();
  fieldsPop.classList.toggle("hidden");
});
fieldsPop.addEventListener("click", (e) => e.stopPropagation());
document.addEventListener("click", () => fieldsPop.classList.add("hidden"));
$("fieldsAll").addEventListener("click", () => setAllFields(true));
$("fieldsNone").addEventListener("click", () => setAllFields(false));

/* ---------- Form selection ---------- */
function updateSelectedBtn() {
  selectedOnlyBtn.classList.toggle("active", showSelectedOnly);
  selectedOnlyBtn.textContent = showSelectedOnly
    ? `Showing ${selectedForms.size} selected`
    : (selectedForms.size ? `Show ${selectedForms.size} selected only` : "Showing all forms");
}

selectedOnlyBtn.addEventListener("click", () => {
  showSelectedOnly = !showSelectedOnly;
  updateSelectedBtn();
  render();
  persistConfig();
});

headerSel.addEventListener("change", () => {
  headerRowIdx = Number(headerSel.value);
  selectedForms.clear();   // row indices shift when the header row changes
  showSelectedOnly = false;
  onConfigChange();
  persistConfig();
});
titleSel.addEventListener("change", () => { render(); persistConfig(); });
searchInp.addEventListener("input", render);
$("expandAll").addEventListener("click", () => toggleAll(false));
$("collapseAll").addEventListener("click", () => toggleAll(true));
$("reset").addEventListener("click", resetApp);

function onConfigChange() {
  const headerCells = rawRows[headerRowIdx] || [];
  const colCount = rawRows.reduce((m, r) => Math.max(m, r.length), 0);
  headers = [];
  for (let c = 0; c < colCount; c++) {
    const label = (headerCells[c] ?? "").toString().trim();
    headers.push(label || `Column ${c + 1}`);
  }
  // data = everything after the header row
  dataRows = [];
  for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
    const row = rawRows[r];
    if (!row || row.every((c) => c === "")) continue; // skip fully empty
    const fields = headers.map((label, c) => ({
      label,
      value: (row[c] ?? "").toString(),
    }));
    dataRows.push({ __i: dataRows.length, fields });
  }
  buildTitleColOptions();
  buildFieldsPopover();
  updateSelectedBtn();
  render();
}

/* ---------- Render ---------- */
function render() {
  const titleColIdx = Number(titleSel.value);
  const q = searchInp.value.trim().toLowerCase();
  formsEl.innerHTML = "";

  const visible = dataRows.filter((row) => {
    if (showSelectedOnly && !selectedForms.has(row.__i)) return false;
    if (!q) return true;
    return row.fields.some((f) =>
      f.value.toLowerCase().includes(q) || f.label.toLowerCase().includes(q));
  });

  const shownFields = headers.length - hiddenCols.size;
  metaEl.textContent = `${dataRows.length} form${dataRows.length === 1 ? "" : "s"}` +
    ((q || showSelectedOnly) ? ` · ${visible.length} shown` : "") +
    ` · ${shownFields} field${shownFields === 1 ? "" : "s"} each`;

  emptyEl.classList.toggle("hidden", visible.length !== 0);

  visible.forEach((row) => {
    formsEl.appendChild(renderSection(row, titleColIdx));
  });
}

function renderSection(row, titleColIdx) {
  const section = document.createElement("div");
  section.className = "form-section";
  if (selectedForms.has(row.__i)) section.classList.add("selected");

  /* head */
  const head = document.createElement("div");
  head.className = "fs-head";

  const title =
    titleColIdx >= 0 && row.fields[titleColIdx] && row.fields[titleColIdx].value
      ? row.fields[titleColIdx].value
      : `Form ${row.__i + 1}`;

  // selection checkbox
  const sel = document.createElement("input");
  sel.type = "checkbox";
  sel.className = "fs-select";
  sel.title = "Select this form";
  sel.checked = selectedForms.has(row.__i);
  sel.addEventListener("click", (e) => e.stopPropagation());
  sel.addEventListener("change", () => {
    if (sel.checked) selectedForms.add(row.__i);
    else selectedForms.delete(row.__i);
    section.classList.toggle("selected", sel.checked);
    updateSelectedBtn();
    persistConfig();
    if (showSelectedOnly) render();
  });

  head.innerHTML =
    ICON_CHEVRON +
    `<span class="fs-index">#${row.__i + 1}</span>` +
    `<h3 class="fs-title"></h3>`;
  head.insertBefore(sel, head.firstChild);
  head.querySelector(".fs-title").textContent = title;

  const copyJsonBtn = document.createElement("button");
  copyJsonBtn.className = "btn accent fs-copy-json";
  copyJsonBtn.innerHTML = ICON_COPY + "<span>Copy JSON</span>";
  copyJsonBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    const obj = {};
    row.fields.forEach((f, i) => { if (!hiddenCols.has(i)) obj[f.label] = f.value; });
    copyText(JSON.stringify(obj, null, 2), "Form copied as JSON");
  });
  head.appendChild(copyJsonBtn);

  head.addEventListener("click", () => section.classList.toggle("collapsed"));

  /* body */
  const body = document.createElement("div");
  body.className = "fs-body";
  row.fields.forEach((f, i) => {
    if (hiddenCols.has(i)) return;
    body.appendChild(renderField(f));
  });

  section.appendChild(head);
  section.appendChild(body);
  return section;
}

function renderField(f) {
  const wrap = document.createElement("div");
  wrap.className = "field";

  const label = document.createElement("div");
  label.className = "field-label";
  label.textContent = f.label;
  label.title = f.label;

  const valRow = document.createElement("div");
  valRow.className = "field-value-row";

  const ta = document.createElement("textarea");
  ta.className = "field-value";
  ta.readOnly = true;
  ta.rows = 1;
  ta.value = f.value;
  autosize(ta);

  const btn = document.createElement("button");
  btn.className = "copy-btn";
  btn.title = "Copy value";
  btn.innerHTML = ICON_COPY;
  btn.addEventListener("click", () => {
    copyText(f.value, "Copied");
    flash(btn);
  });

  valRow.appendChild(ta);
  valRow.appendChild(btn);
  wrap.appendChild(label);
  wrap.appendChild(valRow);
  return wrap;
}

/* ---------- Session persistence (localStorage) ---------- */
function loadStore() {
  try { return JSON.parse(localStorage.getItem(STORE_KEY)) || {}; }
  catch (e) { return {}; }
}
function writeStore(obj) {
  try {
    localStorage.setItem(STORE_KEY, JSON.stringify(obj));
    return true;
  } catch (e) {
    toast("Couldn't save — storage full or unavailable.");
    return false;
  }
}

function saveCurrentSession(name) {
  if (!currentSessionId) return;
  const store = loadStore();
  store[currentSessionId] = {
    id: currentSessionId,
    name: name || (store[currentSessionId] && store[currentSessionId].name) || "Untitled",
    savedAt: Date.now(),
    headerRowIdx,
    titleCol: titleSel.value,
    hiddenCols: [...hiddenCols],
    selectedForms: [...selectedForms],
    showSelectedOnly,
    rows: rawRows,
  };
  writeStore(store);
}

// Update only the config of the active session (no-op if not yet saved).
function persistConfig() {
  if (!currentSessionId) return;
  const store = loadStore();
  if (!store[currentSessionId]) return;
  store[currentSessionId].headerRowIdx = headerRowIdx;
  store[currentSessionId].titleCol = titleSel.value;
  store[currentSessionId].hiddenCols = [...hiddenCols];
  store[currentSessionId].selectedForms = [...selectedForms];
  store[currentSessionId].showSelectedOnly = showSelectedOnly;
  store[currentSessionId].savedAt = Date.now();
  writeStore(store);
}

function loadSession(id) {
  const store = loadStore();
  const s = store[id];
  if (!s) return;
  rawRows = s.rows || [];
  headerRowIdx = s.headerRowIdx || 0;
  currentSessionId = id;
  hiddenCols = new Set(s.hiddenCols || []);
  selectedForms = new Set(s.selectedForms || []);
  showSelectedOnly = !!s.showSelectedOnly;
  buildHeaderRowOptions();
  onConfigChange();
  if (s.titleCol != null) {
    titleSel.value = s.titleCol;
    render();
  }
  showWorkspace();
}

function deleteSession(id) {
  const store = loadStore();
  delete store[id];
  writeStore(store);
  renderSessions();
}

function renderSessions() {
  const store = loadStore();
  const items = Object.values(store).sort((a, b) => b.savedAt - a.savedAt);
  sessionsList.innerHTML = "";
  sessionsEl.classList.toggle("hidden", items.length === 0);

  items.forEach((s) => {
    const rowCount = Math.max(0, (s.rows ? s.rows.length : 0) - 1 - (s.headerRowIdx || 0));
    const item = document.createElement("div");
    item.className = "session-item";
    item.innerHTML =
      `<span class="session-icon">${ICON_FILE}</span>` +
      `<div class="session-meta">` +
        `<div class="session-name"></div>` +
        `<div class="session-sub">${rowCount} form${rowCount === 1 ? "" : "s"} · saved ${timeAgo(s.savedAt)}</div>` +
      `</div>`;
    item.querySelector(".session-name").textContent = s.name;
    item.addEventListener("click", () => loadSession(s.id));

    const del = document.createElement("button");
    del.className = "session-del";
    del.title = "Delete session";
    del.innerHTML = ICON_TRASH;
    del.addEventListener("click", (e) => {
      e.stopPropagation();
      deleteSession(s.id);
    });
    item.appendChild(del);
    sessionsList.appendChild(item);
  });
}

clearSessions.addEventListener("click", () => {
  if (!confirm("Delete all saved sessions?")) return;
  writeStore({});
  renderSessions();
});

function showWorkspace() {
  dropzone.classList.add("hidden");
  sessionsEl.classList.add("hidden");
  config.classList.remove("hidden");
  metaEl.classList.remove("hidden");
}
function showLanding() {
  dropzone.classList.remove("hidden");
  config.classList.add("hidden");
  metaEl.classList.add("hidden");
  emptyEl.classList.add("hidden");
  formsEl.innerHTML = "";
  renderSessions();
}

// Show saved sessions on first load.
renderSessions();

/* ---------- Helpers ---------- */
function toggleAll(collapsed) {
  document.querySelectorAll(".form-section").forEach((s) =>
    s.classList.toggle("collapsed", collapsed));
}

function resetApp() {
  rawRows = []; headers = []; dataRows = [];
  currentSessionId = null;
  hiddenCols = new Set();
  selectedForms = new Set();
  showSelectedOnly = false;
  fileInput.value = "";
  searchInp.value = "";
  showLanding();
}

function autosize(ta) {
  requestAnimationFrame(() => {
    ta.style.height = "auto";
    ta.style.height = Math.min(ta.scrollHeight + 2, 160) + "px";
  });
}

function flash(btn) {
  btn.classList.add("copied");
  btn.innerHTML = ICON_CHECK;
  setTimeout(() => { btn.classList.remove("copied"); btn.innerHTML = ICON_COPY; }, 1100);
}

function copyText(text, msg) {
  const done = () => toast(msg || "Copied");
  if (navigator.clipboard && window.isSecureContext) {
    navigator.clipboard.writeText(text).then(done).catch(() => fallbackCopy(text, done));
  } else {
    fallbackCopy(text, done);
  }
}
function fallbackCopy(text, done) {
  const t = document.createElement("textarea");
  t.value = text;
  t.style.position = "fixed";
  t.style.opacity = "0";
  document.body.appendChild(t);
  t.select();
  try { document.execCommand("copy"); done(); } catch (e) { toast("Copy failed"); }
  document.body.removeChild(t);
}

let toastTimer;
function toast(msg) {
  toastEl.textContent = msg;
  toastEl.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toastEl.classList.remove("show"), 1600);
}

function truncate(s, n) {
  s = String(s);
  return s.length > n ? s.slice(0, n - 1) + "…" : s;
}

function timeAgo(ts) {
  const diff = Date.now() - ts;
  const m = Math.floor(diff / 60000);
  if (m < 1) return "just now";
  if (m < 60) return m + "m ago";
  const h = Math.floor(m / 60);
  if (h < 24) return h + "h ago";
  const d = Math.floor(h / 24);
  if (d < 7) return d + "d ago";
  return new Date(ts).toLocaleDateString();
}
