// viewer.js (ES module) — UI orchestration + chart extraction/rendering.
// Relies on global third-party libs loaded by index.html: XLSX, HyperFormula, fflate.
// Loads Chart.js + plugins on-demand.

/* --------- Config copied from backup (unchanged) --------- */
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

/* NEW: raw workbook bytes (for chart scan) + chart state */
let wbBuf = null;
let chartsBySheet = {};    // { [sheetName]: ChartDef[] }
let chartInstances = [];   // live Chart.js instances for cleanup

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
    renderChartsForActiveSheet();            // NEW
    syncLink();
  });
  showForm.addEventListener("change", () => { renderActiveSheet(); renderChartsForActiveSheet(); }); // NEW
  filterInp.addEventListener("input", applyFilter);
  copyBtn.addEventListener("click", copyDeepLink);

  // Scenarios
  scenarioSel.addEventListener("change", onScenarioChange);

  // Ensure a Charts panel exists (no HTML change required)
  ensureChartsPanel();

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
    chartsBySheet = {};
    destroyAllCharts();
    setChartsStatus("Scanning for charts…");

    const { wb, buf } = await deps.loadWorkbook(path);
    currentWB = wb;
    wbBuf = buf;

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

    // Extract charts from the raw XLSX
    chartsBySheet = await extractChartsFromXLSX(wbBuf);
    debugLogCharts(chartsBySheet);

    // Init scenario from workbook / URL
    initScenarioFromHF();

    // Render sheet + charts
    renderActiveSheet();
    renderChartsForActiveSheet();

    status(`${path} • ${currentWB.SheetNames.length} sheet(s)`);
    syncLink();
  } catch (e) {
    out.innerHTML = `<div class="error">${e.message}. Check the file path (case-sensitive), ensure the workbook is under 100 MB, and not stored via Git LFS.</div>`;
    setChartsStatus("No charts");
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
    renderChartsForActiveSheet();        // NEW
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
      renderChartsForActiveSheet();  // NEW
    }
  });
}

/* ===================== CHARTS (NEW) ===================== */

/* Panel plumbing */
function ensureChartsPanel(){
  let panel = document.getElementById("chartsPanel");
  if (panel) return;
  const aside = document.querySelector("aside");
  if (!aside) return;
  panel = document.createElement("div");
  panel.className = "panel";
  panel.id = "chartsPanel";
  const h2 = document.createElement("h2");
  h2.textContent = "Charts";
  const meta = document.createElement("div");
  meta.className = "meta";
  meta.textContent = "Charts parsed from the workbook.";
  const out = document.createElement("div");
  out.id = "chartsOut";
  panel.appendChild(h2); panel.appendChild(meta); panel.appendChild(out);
  aside.appendChild(panel);

  // quick inline sizing so we don't need CSS edits
  const css = document.createElement("style");
  css.textContent = "#chartsOut canvas{width:100%;height:260px;display:block;margin:10px 0}";
  document.head.appendChild(css);
}
function chartsOutEl(){ return document.getElementById("chartsOut"); }
function setChartsStatus(text){
  const co = chartsOutEl();
  if (!co) return;
  co.innerHTML = `<div class="meta">${text}</div>`;
}

/* Dynamic loader for Chart.js + time adapter + financial plugin */
let chartLibPromise = null;
function loadChartLibsOnce(){
  if (window.Chart && window.Chart.registry) return Promise.resolve();
  if (chartLibPromise) return chartLibPromise;

  function loadScript(src){
    return new Promise((res, rej)=>{
      const s = document.createElement("script"); s.src = src; s.onload = res; s.onerror = () => rej(new Error("Failed to load " + src));
      document.head.appendChild(s);
    });
  }
  chartLibPromise = (async ()=>{
    await loadScript("https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js");
    await loadScript("https://cdn.jsdelivr.net/npm/chartjs-adapter-date-fns@3.0.0/dist/chartjs-adapter-date-fns.bundle.min.js");

    let pluginModule = null;
    try{
      pluginModule = await import("https://cdn.jsdelivr.net/npm/chartjs-chart-financial@3.3.0/dist/chartjs-chart-financial.esm.js");
    }catch(e){
      console.warn("[charts] Failed to import financial plugin via ESM:", e);
    }

    if (window.Chart && window.Chart.register) {
      const registrables = [];

      if (pluginModule && typeof pluginModule === "object") {
        Object.values(pluginModule).forEach(obj => {
          if (obj && (obj.id || obj.prototype?.id)) registrables.push(obj);
        });
      }

      // Fallback: if module failed but global auto-registered, nothing to do.
      if (!registrables.length && window.Chart.registry?.getController("ohlc")) {
        console.info("[charts] Financial plugin already registered (ohlc controller present).");
        return;
      }

      if (registrables.length) {
        const seen = new Set();
        const unique = registrables.filter(obj => {
          const key = obj.id || obj.prototype?.id || obj.name || String(obj);
          if (seen.has(key)) return false;
          seen.add(key);
          return true;
        });
        try{
          window.Chart.register(...unique);
          console.info("[charts] Registered financial plugin controllers:", unique.map(o => o.id || o.name || "unknown"));
        }catch(e){
          console.warn("[charts] Failed to register financial plugin controllers:", e);
        }
      } else {
        console.warn("[charts] Financial plugin not available; stock charts will fall back.");
      }
    }
  })();
  return chartLibPromise;
}

/* Helpers shared by extractor and renderer */
const tdDecoder = new TextDecoder("utf-8");
const byLocal = (root, name) => Array.from(root.getElementsByTagName("*")).filter(n => n.localName === name);
const CHART_TYPE_ORDER = [
  "lineChart","line3DChart","barChart","bar3DChart","columnChart","column3DChart",
  "areaChart","area3DChart","scatterChart","bubbleChart","pieChart","pie3DChart",
  "doughnutChart","radarChart","histogramChart","stockChart","waterfallChart",
  "funnelChart","boxWhiskerChart","sunburstChart","treemapChart","surfaceChart",
  "surface3DChart","wireframeSurfaceChart","wireframeSurface3DChart","paretoChart",
  "comboChart","ofPieChart"
];
const CHART_TYPE_FALLBACK = {
  lineChart: "line",
  line3DChart: "line",
  barChart: "bar",
  bar3DChart: "bar",
  columnChart: "bar",
  column3DChart: "bar",
  areaChart: "area",
  area3DChart: "area",
  scatterChart: "scatter",
  bubbleChart: "bubble",
  pieChart: "pie",
  pie3DChart: "pie",
  doughnutChart: "doughnut",
  radarChart: "radar",
  histogramChart: "histogram",
  paretoChart: "bar",
  stockChart: "stock",
  waterfallChart: "waterfall",
  funnelChart: "funnel",
  boxWhiskerChart: "boxWhisker",
  sunburstChart: "sunburst",
  treemapChart: "treemap",
  surfaceChart: "surface",
  surface3DChart: "surface",
  wireframeSurfaceChart: "surface",
  wireframeSurface3DChart: "surface",
  comboChart: "combo",
  ofPieChart: "pie"
};

function normalisePath(base, target){
  if (!target) return null;
  const clean = target.replace(/\\/g, "/");
  if (/^\//.test(clean)) {
    return clean.replace(/^\//, "").toLowerCase();
  }

  const baseDirParts = (base ? base.split("/").slice(0, -1) : []);
  const rawParts = clean.split("/");
  const stack = [];

  function pushParts(parts){
    parts.forEach(part => {
      if (!part || part === ".") return;
      if (part === "..") {
        if (stack.length) stack.pop();
      } else {
        stack.push(part);
      }
    });
  }

  if (/^xl\//i.test(clean)) {
    pushParts(clean.split("/"));
  } else {
    pushParts(baseDirParts);
    pushParts(rawParts);
  }

  if (stack.length && stack[0].toLowerCase() !== "xl") {
    stack.unshift("xl");
  }
  return stack.join("/").toLowerCase();
}

function excelSerialToDate(serial, use1904){
  if (serial == null || isNaN(serial)) return null;
  let s = Number(serial);
  if (use1904) s += 1462;
  const epoch = Date.UTC(1899, 11, 30); // Excel 1900 epoch with leap bug
  return new Date(epoch + s * 86400000);
}

/* Extract charts from XLSX ArrayBuffer */
const FFLATE_URL = "https://cdn.jsdelivr.net/npm/fflate@0.8.1/umd/fflate.min.js";
let fflateReady = null;
async function ensureFflateReady(){
  if (globalThis.fflate && typeof globalThis.fflate.unzipSync === "function") return globalThis.fflate;
  if (typeof globalThis.ensureFflate === "function") {
    try {
      const lib = await globalThis.ensureFflate();
      if (lib && typeof lib.unzipSync === "function") return lib;
    } catch(_) {}
  }
  if (fflateReady) return fflateReady;
  fflateReady = new Promise((resolve, reject)=>{
    const script = document.createElement("script");
    script.src = FFLATE_URL;
    script.async = true;
    script.onload = () => {
      if (globalThis.fflate && typeof globalThis.fflate.unzipSync === "function") resolve(globalThis.fflate);
      else reject(new Error("fflate failed to initialise"));
    };
    script.onerror = () => reject(new Error("Unable to load fflate library"));
    document.head.appendChild(script);
  }).catch(err => {
    fflateReady = null;
    throw err;
  });
  return fflateReady;
}

async function extractChartsFromXLSX(arrayBuffer){
  try{
    const fflateLib = await ensureFflateReady();
    const zipRaw = fflateLib.unzipSync(new Uint8Array(arrayBuffer));
    // case-insensitive view over zip entries
    const zip = {};
    Object.keys(zipRaw).forEach(k => zip[k.toLowerCase()] = zipRaw[k]);

    const wbXml = zip["xl/workbook.xml"]; const relsXml = zip["xl/_rels/workbook.xml.rels"];
    if(!wbXml || !relsXml){ setChartsStatus("No charts"); return {}; }

    const parse = (u8) => (new DOMParser()).parseFromString(tdDecoder.decode(u8), "application/xml");
    const wdoc = parse(wbXml);
    const rdoc = parse(relsXml);

    const sheets = Array.from(wdoc.getElementsByTagName("sheet")).map(s => ({
      name: s.getAttribute("name"),
      rid:  s.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships","id") || s.getAttribute("r:id")
    }));
    const wbRels = {};
    Array.from(rdoc.getElementsByTagName("Relationship")).forEach(r=>{
      wbRels[r.getAttribute("Id")] = r.getAttribute("Target");
    });

    const meta = sheets.map(s => {
      const tgt = wbRels[s.rid] || "";
      const path = ("xl/" + tgt.replace(/^\//,"")).toLowerCase();
      const kind = /chartsheets\//i.test(tgt) ? "chartsheet" : "worksheet";
      return { name: s.name, path, kind };
    });

    function collectPoints(node){
      if (!node) return null;
      const pts = [];
      const lvlNodes = byLocal(node, "lvl");
      if (lvlNodes.length){
        const buckets = [];
        lvlNodes.forEach((lvl, depth)=>{
          const members = byLocal(lvl, "pt");
          members.forEach(member=>{
            const idx = parseInt(member.getAttribute("idx") ?? `${buckets.length}`, 10);
            const valAttr = member.getAttribute("val");
            const vNode = byLocal(member, "v")[0];
            const text = valAttr != null ? valAttr : (vNode ? vNode.textContent : member.textContent);
            if (!buckets[idx]) buckets[idx] = [];
            buckets[idx][depth] = text;
          });
        });
        return buckets.map(path => (path || []).filter(Boolean).join(" / "));
      }

      const pointNodes = byLocal(node, "pt");
      pointNodes.forEach(pt=>{
        const idxAttr = pt.getAttribute("idx");
        const idx = idxAttr != null ? parseInt(idxAttr, 10) : pts.length;
        const valAttr = pt.getAttribute("val");
        const vNode = byLocal(pt, "v")[0];
        const text = valAttr != null ? valAttr : (vNode ? vNode.textContent : pt.textContent);
        pts[idx] = text;
      });
      if (pts.length) return pts;

      const nary = byLocal(node, "ptValue");
      if (nary.length){
        nary.forEach((pt, i)=>{
          const idxAttr = pt.getAttribute("idx");
          const idx = idxAttr != null ? parseInt(idxAttr, 10) : i;
          pts[idx] = pt.textContent;
        });
      }
      return pts.length ? pts : null;
    }

    function firstDescendantWithLocal(node, localNames){
      if (!node) return null;
      const wanted = Array.isArray(localNames) ? localNames : [localNames];
      const all = node.getElementsByTagName("*");
      for (let i=0;i<all.length;i++){
        const el = all[i];
        if (wanted.includes(el.localName)) return el;
      }
      return null;
    }

    function pickRefOrData(seriesNode, candidates){
      let refF = null, data = null;
      for (const candidate of candidates){
        const containers = byLocal(seriesNode, candidate);
        for (const container of containers){
          const refNode = firstDescendantWithLocal(container, ["numRef","strRef","numData","strData","multiLvlStrRef"]);
          if (refNode && !refF){
            const fNode = firstDescendantWithLocal(refNode, "f");
            if (fNode && fNode.textContent) refF = fNode.textContent.trim();
          }
          if (!data){
            data = collectPoints(refNode) || collectPoints(container);
          }
          if (refF || data) break;
        }
        if (refF || data) break;
      }
      return { refF, data };
    }

    function readSeriesName(txNode){
      if (!txNode) return { nameF:null, nameV:null };
      const ref = firstDescendantWithLocal(txNode, ["strRef","strData"]);
      if (ref){
        const fNode = firstDescendantWithLocal(ref, "f");
        if (fNode && fNode.textContent) return { nameF: fNode.textContent.trim(), nameV: null };
        const data = collectPoints(ref);
        if (data && data.length) return { nameF: null, nameV: String(data[0] ?? "") };
      }
      const vNode = firstDescendantWithLocal(txNode, ["v","t","r"]);
      if (vNode && vNode.textContent) return { nameF:null, nameV: vNode.textContent.trim() };
      const text = txNode.textContent?.trim();
      return text ? { nameF:null, nameV:text } : { nameF:null, nameV:null };
    }

    function extractTitle(doc){
      const titleNode = byLocal(doc, "title")[0];
      if (!titleNode) return { titleF:null, titleText:null };
      const ref = firstDescendantWithLocal(titleNode, ["strRef","strData"]);
      if (ref){
        const fNode = firstDescendantWithLocal(ref, "f");
        if (fNode && fNode.textContent) return { titleF: fNode.textContent.trim(), titleText: null };
        const data = collectPoints(ref);
        if (data && data.length) return { titleF:null, titleText:String(data[0] ?? "") };
      }
      const vNode = firstDescendantWithLocal(titleNode, ["v","t"]);
      if (vNode && vNode.textContent) return { titleF:null, titleText: vNode.textContent.trim() };
      const text = titleNode.textContent?.trim();
      return text ? { titleF:null, titleText:text } : { titleF:null, titleText:null };
    }

    function parseClassicChart(doc){
      const typeNode = CHART_TYPE_ORDER.map(t => byLocal(doc, t)[0]).find(Boolean);
      if (!typeNode) return null;
      const mappedType = CHART_TYPE_FALLBACK[typeNode.localName] || "line";
      const { titleF, titleText } = extractTitle(doc);
      const serNodes = byLocal(typeNode, "ser").concat(byLocal(typeNode, "series"));
      const series = serNodes.map(sN => {
        const { nameF, nameV } = readSeriesName(byLocal(sN, "tx")[0]);
        const cat = pickRefOrData(sN, ["cat","category","xVal","categories"]);
        const y   = pickRefOrData(sN, ["val","yVal","values"]);
        const z   = pickRefOrData(sN, ["bubbleSize","zVal","size","sizes"]);
        return {
          nameF,
          nameV,
          catRef: cat.refF || null,
          catData: cat.data || null,
          xRef: cat.refF || null,
          xData: cat.data || null,
          yRef: y.refF || null,
          yData: y.data || null,
          zRef: z.refF || null,
          zData: z.data || null
        };
      });
      return { type: mappedType, titleF, titleText, series };
    }

    function parseChartPart(u8){
      const doc = parse(u8);
      const classic = parseClassicChart(doc);
      if (classic) return classic;
      return null;
    }

    const chartsBySheet = {};
    // Walk sheets
    for (const m of meta){
      if (m.kind === "worksheet"){
        // locate drawings via sheet rels
        const relPath = `xl/worksheets/_rels/${m.path.split("/").pop()}.rels`;
        const relEntry = zip[relPath];
        if (relEntry){
          const relDoc = parse(relEntry);
          const drawRels = Array.from(relDoc.getElementsByTagName("Relationship")).filter(r => /drawing/i.test(r.getAttribute("Type")||""));
          for (const dr of drawRels){
            const target = dr.getAttribute("Target");
            const drawingPath = normalisePath(m.path, target);
            if (!drawingPath) {
              console.warn(`[charts] Failed to normalize drawing path from ${m.path} + ${target}`);
              continue;
            }
            const drawingXml = zip[drawingPath];
            if (!drawingXml) {
              console.warn(`[charts] Drawing not found: ${drawingPath}`);
              continue;
            }

            // map drawing rId -> chart/chartEx
            const dRelsPath = `xl/drawings/_rels/${drawingPath.split("/").pop()}.rels`;
            const dRelsXml = zip[dRelsPath];
            if (!dRelsXml) {
              console.warn(`[charts] Drawing rels not found: ${dRelsPath}`);
              continue;
            }
            const dRels = {};
            Array.from(parse(dRelsXml).getElementsByTagName("Relationship")).forEach(r=>{
              const relTarget = r.getAttribute("Target");
              const relPath = normalisePath(drawingPath, relTarget);
              if (relPath) {
                dRels[r.getAttribute("Id")] = relPath;
              } else {
                console.warn(`[charts] Failed to normalize chart path from ${drawingPath} + ${relTarget}`);
              }
            });

            const dDoc = parse(drawingXml);
            const chartElems = byLocal(dDoc, "chart"); // catches c:chart and cx:chart
            for (const cEl of chartElems){
              const rid = cEl.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships","id") || cEl.getAttribute("r:id");
              const chartPath = dRels[rid];
              if (!chartPath) {
                console.warn(`[charts] No chart path for rId ${rid} in drawing`);
                continue;
              }
              if (!zip[chartPath]) {
                console.warn(`[charts] Chart file not found: ${chartPath} (tried rId ${rid})`);
                continue;
              }
              const def = parseChartPart(zip[chartPath]);
              if (!def) {
                console.warn(`[charts] Failed to parse chart: ${chartPath}`);
                continue;
              }
              if (!chartsBySheet[m.name]) chartsBySheet[m.name] = [];
              chartsBySheet[m.name].push(def);
            }
          }
        }
      } else if (m.kind === "chartsheet"){
        // chartsheet -> chart directly via rels
        const csRelsPath = `xl/chartsheets/_rels/${m.path.split("/").pop()}.rels`;
        const csXml = zip[csRelsPath];
        if (csXml){
          const csDoc = parse(csXml);
          const chartRel = Array.from(csDoc.getElementsByTagName("Relationship"))
            .find(r => /relationships\/chart/i.test(r.getAttribute("Type")||""));
          if (chartRel){
            const chartPath = normalisePath(m.path, chartRel.getAttribute("Target"));
            const chartXml = zip[chartPath];
            if (chartXml){
              const def = parseChartPart(chartXml);
              if (def){
                if (!chartsBySheet[m.name]) chartsBySheet[m.name] = [];
                chartsBySheet[m.name].push(def);
              }
            }
          }
        }
      }
    }
    return chartsBySheet;
  }catch(e){
    console.warn("Chart extraction failed:", e);
    setChartsStatus("Charts unavailable");
    return {};
  }
}

function debugLogCharts(map){
  try{
    const names = Object.keys(map);
    console.info("[charts] sheets:", names.length ? names.join(", ") : "(none)");
    names.forEach(n => {
      console.info(`[charts] ${n}: ${map[n].length} chart(s)`);
      map[n].forEach((ch, i) => {
        console.info(`  [${i}] type: ${ch.type}, series: ${ch.series?.length || 0}, title: ${ch.titleText || ch.titleF || "(none)"}`);
      });
    });
  }catch(e){
    console.warn("[charts] debug failed:", e);
  }
}

function destroyAllCharts(){
  chartInstances.forEach(ch => { try{ ch.destroy(); }catch{} });
  chartInstances = [];
}

function rangeToVector(vals2d){
  if(!Array.isArray(vals2d) || !vals2d.length) return [];
  if (vals2d.length === 1) return vals2d[0].slice();
  if (vals2d[0].length === 1) return vals2d.map(r => r[0]);
  return vals2d.flat();
}
function parseSheetAndRange(a1){
  if(!a1) return null;
  // 'Sheet Name'!$A$1:$B$9
  const m = a1.match(/^(?:'([^']+)'|([^'!]+))!(\$?[A-Z]+\$?[0-9]+(?::\$?[A-Z]+\$?[0-9]+)?)$/);
  if(!m) return null;
  const sheet = (m[1] || m[2]);
  const range = m[3].replace(/\$/g,'');
  return { sheet, range };
}
function getRangeValues(sheetName, rangeA1){
  if(!hf) return [];
  const id = hf.getSheetId(sheetName);
  if (id == null) return [];
  const { r0, c0 } = sheetOffsets[sheetName] || { r0:0, c0:0 };
  const r = XLSX.utils.decode_range(rangeA1);
  const out = [];
  for(let R=r.s.r; R<=r.e.r; ++R){
    const row = [];
    for(let C=r.s.c; C<=r.e.c; ++C){
      let v = hf.getCellValue({ sheet:id, row: R - r0, col: C - c0 });
      if (v && typeof v === "object" && "type" in v) v = null; // HF error -> null
      row.push(v);
    }
    out.push(row);
  }
  return out;
}
function evalTitle(def){
  if (def.titleF){
    const pr = parseSheetAndRange(def.titleF);
    if (pr){
      const vec = rangeToVector(getRangeValues(pr.sheet, pr.range));
      if (vec.length) return String(vec[0] ?? "");
    }
  }
  return def.titleText || "";
}
function evalSeriesName(s){
  if (s.nameF){
    const pr = parseSheetAndRange(s.nameF);
    if (pr){
      const vec = rangeToVector(getRangeValues(pr.sheet, pr.range));
      if (vec.length) return String(vec[0] ?? "");
    }
  }
  if (s.nameV) return s.nameV;
  return "";
}

/* Build histogram bins if Excel didn't */
function buildHistogram(values){
  const nums = values.map(Number).filter(v => isFinite(v));
  if (!nums.length) return { labels:[], counts:[] };
  const min = Math.min(...nums), max = Math.max(...nums);
  const k = Math.max(1, Math.ceil(Math.log2(nums.length) + 1));
  const width = (max - min) / k || 1;
  const counts = new Array(k).fill(0);
  nums.forEach(v => {
    let idx = Math.floor((v - min) / width);
    if (idx >= k) idx = k - 1;
    if (idx < 0) idx = 0;
    counts[idx]++;
  });
  const labels = counts.map((_,i)=>{
    const a = min + i*width, b = a + width;
    return `${a.toFixed(2)}–${b.toFixed(2)}`;
  });
  return { labels, counts };
}

/* Render charts for currently selected sheet */
async function renderChartsForActiveSheet(){
  const sheet = sheetSel.value;
  const co = chartsOutEl();
  if (!co) return;

  destroyAllCharts();
  co.innerHTML = "";

  const defs = chartsBySheet[sheet] || [];
  if (!defs.length){
    setChartsStatus("No charts found on this sheet.");
    return;
  }

  await loadChartLibsOnce();

  defs.slice(0, 12).forEach((def, idx) => {
    const title = evalTitle(def);
    const cnv = document.createElement("canvas");
    cnv.setAttribute("aria-label", title || ("Chart "+(idx+1)));
    co.appendChild(cnv);

    // Helpers to resolve labels/x/y/z either from range refs (HF) or cached data
    function resolveLabels(firstSeries){
      if (firstSeries?.catRef){
        const pr = parseSheetAndRange(firstSeries.catRef);
        if (pr) return rangeToVector(getRangeValues(pr.sheet, pr.range)).map(v => String(v ?? ""));
      }
      if (Array.isArray(firstSeries?.catData)) return firstSeries.catData.map(v => String(v ?? ""));
      return [];
    }
    function resolveY(s){
      if (s.yRef){
        const pr = parseSheetAndRange(s.yRef);
        if (pr) return rangeToVector(getRangeValues(pr.sheet, pr.range)).map(Number);
      }
      if (Array.isArray(s.yData)) return s.yData.map(Number);
      return [];
    }
    function resolveX(s){
      if (s.xRef){
        const pr = parseSheetAndRange(s.xRef);
        if (pr) return rangeToVector(getRangeValues(pr.sheet, pr.range));
      }
      if (Array.isArray(s.xData)) return s.xData.slice();
      return [];
    }
    function resolveZ(s){
      if (s.zRef){
        const pr = parseSheetAndRange(s.zRef);
        if (pr) return rangeToVector(getRangeValues(pr.sheet, pr.range)).map(Number);
      }
      if (Array.isArray(s.zData)) return s.zData.map(Number);
      return [];
    }

    if (def.type === "pie" || def.type === "doughnut"){
      const s0 = def.series[0] || {};
      const labels = resolveLabels(s0);
      const data   = resolveY(s0);
      const cfg = {
        type: def.type,
        data: { labels, datasets: [{ label: evalSeriesName(s0) || title || "Series", data }] },
        options: { responsive: true, plugins: { legend: { position:'top' }, title: { display: !!title, text: title } } }
      };
      chartInstances.push(new Chart(cnv.getContext("2d"), cfg));

    } else if (def.type === "scatter"){
      const datasets = def.series.map((s, i) => {
        const xs = resolveX(s).map(Number);
        const ys = resolveY(s);
        const n = Math.min(xs.length, ys.length);
        const data = Array.from({length:n}, (_,k)=>({x: xs[k], y: ys[k]}));
        return { label: evalSeriesName(s) || ("Series " + (i+1)), data, showLine:false };
      });
      const cfg = {
        type: "scatter",
        data: { datasets },
        options: { responsive:true, scales:{ x:{ type:"linear" }, y:{ type:"linear" } }, plugins:{ title:{ display: !!title, text: title } } }
      };
      chartInstances.push(new Chart(cnv.getContext("2d"), cfg));

    } else if (def.type === "bubble"){
      const datasets = def.series.map((s, i) => {
        const xs = resolveX(s).map(Number);
        const ys = resolveY(s);
        const zs = resolveZ(s);
        const n = Math.min(xs.length, ys.length, zs.length || Infinity);
        const zMin = Math.min(...zs.filter(isFinite), 0), zMax = Math.max(...zs.filter(isFinite), 1);
        const size = (z)=> (!isFinite(z) || zMax===zMin) ? 6 : (4 + 12 * (z - zMin) / (zMax - zMin));
        const data = Array.from({length:n}, (_,k)=>({x: xs[k], y: ys[k], r: size(zs[k])}));
        return { label: evalSeriesName(s) || ("Series " + (i+1)), data };
      });
      const cfg = { type: "bubble", data: { datasets }, options: { responsive:true, plugins:{ title:{ display: !!title, text: title } } } };
      chartInstances.push(new Chart(cnv.getContext("2d"), cfg));

    } else if (def.type === "radar"){
      const s0 = def.series[0] || {};
      const labels = resolveLabels(s0);
      const datasets = def.series.map((s, i) => ({
        label: evalSeriesName(s) || ("Series " + (i+1)),
        data: resolveY(s)
      }));
      const cfg = { type: "radar", data: { labels, datasets }, options: { responsive:true, plugins:{ title:{ display: !!title, text: title } } } };
      chartInstances.push(new Chart(cnv.getContext("2d"), cfg));

    } else if (def.type === "histogram"){
      const s0 = def.series[0] || {};
      const x = resolveX(s0), y = resolveY(s0);
      let labels = [], counts = [];
      if (x.length && y.length && x.length === y.length){
        labels = x.map(v=>String(v ?? ""));
        counts = y;
      } else {
        const values = y.length ? y : x.map(Number);
        const h = buildHistogram(values);
        labels = h.labels; counts = h.counts;
      }
      const cfg = { type: "bar", data: { labels, datasets:[{ label: evalSeriesName(s0) || title || "Histogram", data: counts }] }, options: { responsive:true, plugins:{ title:{ display: !!title, text: title } } } };
      chartInstances.push(new Chart(cnv.getContext("2d"), cfg));

    } else if (def.type === "stock"){
      // Expect datasets for O/H/L/C; accept either named series or positional order
      const s = def.series;
      if (!s || !s.length) {
        console.warn("[charts] Stock chart has no series");
        setChartsStatus("Stock chart: no data");
        return;
      }
      
      const nameMap = {};
      s.forEach((ser, i)=>{
        const nm = (evalSeriesName(ser) || "").toLowerCase();
        if (/open/.test(nm)) nameMap.open = i;
        else if (/high/.test(nm)) nameMap.high = i;
        else if (/low/.test(nm)) nameMap.low = i;
        else if (/close|last/.test(nm)) nameMap.close = i;
      });
      const idxOpen = nameMap.open ?? 0;
      const idxHigh = nameMap.high ?? 1;
      const idxLow  = nameMap.low  ?? 2;
      const idxClose= nameMap.close?? 3;

      const refSeries = s[idxOpen] || s[0] || {};
      let labels = [];
      if (refSeries.catRef){
        const pr = parseSheetAndRange(refSeries.catRef);
        if (pr) labels = rangeToVector(getRangeValues(pr.sheet, pr.range));
      } else if (Array.isArray(refSeries.catData)){
        labels = refSeries.catData.slice();
      }

      const opens = (s[idxOpen]) ? resolveY(s[idxOpen]) : [];
      const highs = (s[idxHigh]) ? resolveY(s[idxHigh]) : [];
      const lows  = (s[idxLow ]) ? resolveY(s[idxLow ]) : [];
      const closes= (s[idxClose])? resolveY(s[idxClose]) : [];
      const n = Math.min(labels.length, opens.length, highs.length, lows.length, closes.length);

      if (n === 0) {
        console.warn("[charts] Stock chart: no valid data points", {labels: labels.length, opens: opens.length, highs: highs.length, lows: lows.length, closes: closes.length});
        setChartsStatus("Stock chart: insufficient data");
        return;
      }

      const use1904 = !!date1904BySheet[sheet];
      const data = Array.from({length:n}, (_,i)=>{
        const L = labels[i];
        let x = L;
        if (typeof L === "number") {
          const d = excelSerialToDate(L, use1904);
          if (d) x = d.getTime();
        } else if (typeof L === "string") {
          const d = new Date(L);
          if (!isNaN(d.getTime())) x = d.getTime();
        }
        return { x, o: +opens[i] || 0, h: +highs[i] || 0, l: +lows[i] || 0, c: +closes[i] || 0 };
      });

      let cfg = {
        type: "ohlc",
        data: { datasets: [{ label: title || "OHLC", data }] },
        options: {
          responsive: true,
          parsing: false,
          scales: { x: { type: "time", time: { unit: "day" } } },
          plugins: { title: { display: !!title, text: title } }
        }
      };
      try{
        if (!window.Chart || !window.Chart.registry || !window.Chart.registry.getController("ohlc")) {
          throw new Error("OHLC controller not registered");
        }
        chartInstances.push(new Chart(cnv.getContext("2d"), cfg));
      }catch(e){
        console.warn("[charts] OHLC chart failed, falling back:", e);
        // Degrade gracefully if plugin unavailable
        cfg = {
          type: "bar",
          data: { labels: labels.map(v=>String(v ?? "")), datasets:[{ label: "Close", data: closes }] },
          options: { responsive:true, plugins:{ title:{ display:true, text: (title? title+" (Close only)" : "Close") } } }
        };
        chartInstances.push(new Chart(cnv.getContext("2d"), cfg));
      }

    } else {
      // line / area / bar (default)
      const s0 = def.series[0] || {};
      const labels = resolveLabels(s0);
      const datasets = def.series.map((s, i) => ({
        label: evalSeriesName(s) || ("Series " + (i+1)),
        data: resolveY(s)
      }));
      const chartTypeMap = {
        area: "line",
        bar: "bar",
        waterfall: "bar",
        funnel: "bar",
        boxWhisker: "bar",
        treemap: "bar",
        sunburst: "bar",
        surface: "line",
        combo: "bar",
        line: "line"
      };
      const chartType = chartTypeMap[def.type] || "line";
      const options = { responsive:true, plugins:{ title:{ display: !!title, text: title } } };
      if (chartType === "line" && def.type === "area") {
        datasets.forEach(d => d.fill = true);
      }
      const cfg = { type: chartType, data: { labels, datasets }, options };
      chartInstances.push(new Chart(cnv.getContext("2d"), cfg));
    }
  });
}

/* ===================== END CHARTS ===================== */
