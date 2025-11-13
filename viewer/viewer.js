// viewer.js (ES module) — UI orchestration only. Relies on global `window.deps` helpers.

/* --------- Config copied from backup --------- */
const files = [
  { path: "../models/3_Statement_Model.xlsx", label: "3-Statement Model" },
  { path: "../models/DCF_SN.xlsx",            label: "DCF Valuations" },
  { path: "../models/Comps_Precs.xlsx",       label: "Comparables and Precedent Valuations" },
  { path: "../models/Valuation_Overview.xlsx",label: "Valuation Overview" }
];

const SCENARIO_OPTS = ["Bear","Central","Bull"];
const SCENARIOS = [
  { sheet: "DCF",  targets: ["J22","J34"] },
  { sheet: "SOTP", targets: ["B28","H28","N28"] },
  { sheet: "rNPV", targets: ["A1"] }
];

const LINKED_FILES = ["../models/3_Statement_Model.xlsx","../models/DCF_SN.xlsx"];
const MAX_R = 1500, MAX_C = 80;

/* --------- DOM --------- */
const fileSel     = document.getElementById("file");
const sheetSel    = document.getElementById("sheet");
const scenarioSel = document.getElementById("scenario");
const showForm    = document.getElementById("showFormula");
const filterInp   = document.getElementById("filter");
const copyBtn     = document.getElementById("copyLink");
const dlLink      = document.getElementById("downloadLink");
const out         = document.getElementById("out");
const statusEl    = document.getElementById("status");
const cellInfo    = document.getElementById("cellInfo");

/* --------- State --------- */
let currentPath = null;
let currentWB   = null;
let currentWS   = null;
let hf          = null;
let styleCtx    = null;
let sheetOffsets = {};     // { [sheetName]: { r0, c0 } }
let date1904BySheet = {};  // { [sheetName]: boolean }
let currentHTMLTable = null;

/* --------- Init UI --------- */
(function bootstrap() {
  // Files dropdown
  files.forEach(f => {
    const opt = document.createElement("option");
    opt.value = f.path; opt.textContent = f.label;
    fileSel.appendChild(opt);
  });
  fileSel.value = files[0].path;

  // Scenario dropdown
  scenarioSel.innerHTML = SCENARIO_OPTS.map(v => `<option value="${v}">${v}</option>`).join("");

  // Wire controls
  fileSel.addEventListener("change", () => loadWorkbook(fileSel.value));
  sheetSel.addEventListener("change", () => {
    if (!currentWB) return;
    const name = sheetSel.value;
    currentWS = currentWB.Sheets[name];
    renderActiveSheet();
    syncLink();
  });
  showForm.addEventListener("change", () => renderActiveSheet());
  filterInp.addEventListener("input", applyFilter);
  copyBtn.addEventListener("click", copyDeepLink);

  // Scenarios
  scenarioSel.addEventListener("change", onScenarioChange);

  // Deep-link boot
  initFromQueryAndLoad();
})();

/* --------- Orchestration --------- */
async function loadWorkbook(path) {
  try {
    currentPath = path;
    status("Loading workbook…");
    out.innerHTML = "";
    sheetSel.innerHTML = "";
    styleCtx = null;

    const { wb, buf } = await deps.loadWorkbook(path);
    currentWB = wb;

    styleCtx = await deps.extractStyles(buf);

    // Populate sheets
    currentWB.SheetNames.forEach(s => {
      const opt = document.createElement("option");
      opt.value = s; opt.textContent = s;
      sheetSel.appendChild(opt);
    });
    const initial = currentWB.SheetNames.includes("Summary") ? "Summary" : currentWB.SheetNames[0];
    sheetSel.value = initial;
    currentWS = currentWB.Sheets[initial];

    // Build recalc model (HF) across current + linked files
    const hfBundle = await deps.buildHF({
      primaryPath: currentPath,
      linkedFiles: LINKED_FILES
    });
    hf = hfBundle.hf;
    sheetOffsets = hfBundle.sheetOffsets;
    date1904BySheet = hfBundle.date1904BySheet;

    // Init scenario from workbook / URL
    initScenarioFromHF();

    // Render
    renderActiveSheet();
    status(`${path} • ${currentWB.SheetNames.length} sheet(s)`);
    syncLink();
  } catch (e) {
    out.innerHTML = `<div class="error">${e.message}. Check the file path (case-sensitive), ensure the workbook is under 100 MB, and not stored via Git LFS.</div>`;
    status("Error loading workbook");
  }
}

function renderActiveSheet() {
  if (!currentWS) return;
  const res = deps.renderSheet({
    container: out,
    ws: currentWS,
    sheetName: sheetSel.value,
    hf,
    styleCtx,
    showFormulae: !!showForm.checked,
    sheetOffsets,
    date1904BySheet,
    MAX_R,
    MAX_C
  });
  currentHTMLTable = res.table;

  // Cell inspector
  if (currentHTMLTable) {
    currentHTMLTable.addEventListener("click", onCellClick);
  }
  applyFilter();
}

function onCellClick(ev){
  const td = ev.target.closest("td");
  if(!td || !td.dataset.address) return;
  const addr = td.dataset.address;
  const cell = currentWS[addr] || {};
  const fmt  = XLSX.utils.format_cell(cell);
  const info = {
    Address: addr,
    Value: fmt || "",
    Raw: (cell.v!==undefined ? String(cell.v) : ""),
    Formula: (cell.f ? "=" + cell.f : ""),
    Type: cell.t || "",
    Format: cell.z || "",
    Hyperlink: (cell.l && (cell.l.Target || cell.l.target)) ? (cell.l.Target || cell.l.target) : ""
  };
  cellInfo.innerHTML = "";
  Object.entries(info).forEach(([k,v])=>{
    const b = document.createElement("b"); b.textContent = k;
    const s = document.createElement("span"); s.textContent = v || "–";
    cellInfo.appendChild(b); cellInfo.appendChild(s);
  });
}

/* --------- Scenario helpers --------- */
function onScenarioChange(){
  if (!hf) return;
  try {
    setScenarioInHF(scenarioSel.value);
    if (hf.recompute) hf.recompute();
    renderActiveSheet();
    status("Scenario updated");
  } catch (e) {
    status("Scenario update failed: " + e.message);
  }
}
function setScenarioInHF(value){
  SCENARIOS.forEach(({sheet, targets}) => {
    const id = hf.getSheetId(sheet);
    const { r0, c0 } = sheetOffsets[sheet] || { r0: 0, c0: 0 };
    targets.forEach(addr => {
      const { r, c } = XLSX.utils.decode_cell(addr);
      hf.setCellContents({ sheet: id, row: r - r0, col: c - c0 }, [[value]]);
    });
  });
}
function readScenarioFromHF(){
  for (const {sheet, targets} of SCENARIOS){
    const id = hf.getSheetId(sheet);
    const { r0, c0 } = sheetOffsets[sheet] || { r0: 0, c0: 0 };
    for (const addr of targets){
      const { r, c } = XLSX.utils.decode_cell(addr);
      const v = hf.getCellValue({ sheet: id, row: r - r0, col: c - c0 });
      if (v != null) return String(v);
    }
  }
  return null;
}
function initScenarioFromHF(){
  try {
    if (!hf) return;
    const curVal = readScenarioFromHF();
    if (curVal && SCENARIO_OPTS.includes(curVal)) scenarioSel.value = curVal;

    const params = new URLSearchParams(location.search);
    const qSc = params.get("sc");
    if (qSc && SCENARIO_OPTS.includes(qSc)) {
      scenarioSel.value = qSc;
      setScenarioInHF(qSc);
    }
  } catch(_) {}
}

/* --------- Utilities --------- */
function applyFilter(){
  if(!currentHTMLTable) return;
  const q = (filterInp.value||"").toLowerCase();
  const rows = currentHTMLTable.querySelectorAll("tbody tr");
  rows.forEach(row=>{
    if(!q){ row.classList.remove("hidden"); return; }
    const text = row.textContent.toLowerCase();
    row.classList.toggle("hidden", !text.includes(q));
  });
}
function copyDeepLink(){
  const short = (currentPath||"").split("/").pop();
  const url = new URL(location.href);
  url.searchParams.set("file", short);
  url.searchParams.set("sheet", sheetSel.value);
  url.searchParams.set("f", showForm.checked ? "1":"0");
  url.searchParams.set("sc", scenarioSel.value);
  navigator.clipboard.writeText(url.toString()).then(() => {
    status("Link copied");
  }).catch(() => status("Could not copy link"));
}
function syncLink(){
  dlLink.href = currentPath;
  dlLink.textContent = "Download " + (files.find(f=>f.path===currentPath)?.label || "workbook");
}
function status(t){ statusEl.textContent = t; }

function initFromQueryAndLoad(){
  const params = new URLSearchParams(location.search);
  const qFile  = params.get("file");
  const qSheet = params.get("sheet");
  const qF     = params.get("f");
  const qSc    = params.get("sc");
  if(qF === "1") showForm.checked = true;
  if (qSc && SCENARIO_OPTS.includes(qSc)) scenarioSel.value = qSc;

  if(qFile){
    const match = files.find(f => f.path.endsWith("/"+qFile) || f.label === qFile);
    if(match) fileSel.value = match.path;
  }
  loadWorkbook(fileSel.value).then(()=>{
    if(qSheet && currentWB.SheetNames.includes(qSheet)){
      sheetSel.value = qSheet;
      currentWS = currentWB.Sheets[qSheet];
      renderActiveSheet();
    }
  });
}
