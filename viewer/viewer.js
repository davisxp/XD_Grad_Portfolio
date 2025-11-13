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
const CHART_LIMIT = 12;
const DEFAULT_FILE = files[0].path;

/* UI elements */
const fileSel     = document.getElementById("fileSel");
const sheetSel    = document.getElementById("sheetSel");
const reloadBtn   = document.getElementById("reloadBtn");
const inspectorEl = document.getElementById("inspector");
const tableWrap   = document.getElementById("tableWrap");
const tableEl     = document.getElementById("sheetTable");
const chartsWrap  = document.getElementById("chartsWrap");
const chartsStatus= document.getElementById("chartsStatus");

/* Populate file dropdown */
(function initFileSel(){
  fileSel.innerHTML = "";
  files.forEach(f=>{
    const opt = document.createElement("option");
    opt.value = f.path;
    opt.textContent = f.label;
    fileSel.appendChild(opt);
  });
  fileSel.value = DEFAULT_FILE;
})();

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
let chartInstances = [];   // live Chart.js instances

/* --------- Utilities --------- */

function byId(id){ return document.getElementById(id); }
function escapeHTML(s){ return String(s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }
function fmtA1(r, c){ return XLSX.utils.encode_cell({ r, c }); }

function clearChildren(el){ while(el && el.firstChild) el.removeChild(el.firstChild); }

function setChartsStatus(msg){
  if (!chartsStatus) return;
  chartsStatus.textContent = msg || "";
  chartsStatus.style.display = msg ? "block" : "none";
}

function chartsOutEl(){
  let el = document.getElementById("chartsOut");
  if (!el) {
    el = document.createElement("div");
    el.id = "chartsOut";
    el.className = "charts-grid";
    chartsWrap.appendChild(el);
  }
  return el;
}

/* column width helper (px approximation) */
function colWidthPx(w){ return Math.max(40, Math.floor((w || 10) * 8.2)); }

/* date detection */
function isDateCell(v, z){
  if (v == null) return false;
  if (v instanceof Date) return true;
  if (typeof v === "number" && z && /[dmysh]/i.test(z)) return true;
  return false;
}

/* --------- HyperFormula integration --------- */

function initHFIfNeeded(workbook){
  if (hf) return;
  hf = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });
  const sheets = workbook.SheetNames;
  sheets.forEach((name, idx) => {
    const ws = workbook.Sheets[name];
    const json = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    const id = hf.addSheet(name);
    hf.setCellContents({ sheet: id, row: 0, col: 0 }, json);
  });
}

function evalFormula(a1){
  if (!hf || currentWS == null) return null;
  try {
    const res = hf.getCellValue({ sheet: hf.getSheetId(currentWS), col: a1.c, row: a1.r });
    return res;
  } catch {
    return null;
  }
}

/* get sheet offset for stable row/col display */
function getSheetOffset(name){
  return sheetOffsets[name] || { r0: 0, c0: 0 };
}

/* build an HTML table for a sheet range */
function buildTable(ws, range){
  const html = [];
  html.push('<table class="grid">');

  const R = range.e.r - range.s.r + 1;
  const C = range.e.c - range.s.c + 1;

  html.push("<thead><tr><th></th>");
  for (let c = 0; c < C; ++c){
    html.push(`<th>${escapeHTML(XLSX.utils.encode_col(c + range.s.c))}</th>`);
  }
  html.push("</tr></thead>");

  html.push("<tbody>");
  for (let r = 0; r < R; ++r){
    const rr = r + range.s.r;
    html.push(`<tr><th>${escapeHTML(String(rr + 1))}</th>`);
    for (let c = 0; c < C; ++c){
      const cc = c + range.s.c;
      const a1 = fmtA1(rr, cc);
      const cell = ws[a1];
      let v = cell ? cell.v : "";
      if (cell && isDateCell(cell.v, cell.z)) {
        v = cell.v instanceof Date ? cell.v.toISOString().slice(0,10) : v;
      }
      html.push(`<td data-a1="${a1}">${escapeHTML(v)}</td>`);
    }
    html.push("</tr>");
  }
  html.push("</tbody>");

  html.push("</table>");
  return html.join("");
}

/* apply basic styling from sheet to table cols */
function styleTableFromSheet(ws, table){
  if (!ws || !table) return;
  const cols = ws['!cols'] || [];
  const colEls = table.querySelectorAll("colgroup col");
  if (!colEls.length){
    const gc = document.createElement("colgroup");
    const range = XLSX.utils.decode_range(ws['!ref'] || "A1");
    const C = range.e.c - range.s.c + 1;
    for (let c = 0; c < C; ++c) {
      const col = document.createElement("col");
      col.style.width = colWidthPx(cols[c]?.wpx || cols[c]?.wch || 10) + "px";
      gc.appendChild(col);
    }
    table.insertBefore(gc, table.firstChild);
  }
}

/* --------- Workbook load + table render --------- */

async function loadWorkbookFromPath(path){
  const res = await fetch(path, { cache: "no-store" });
  if (!res.ok) throw new Error(`Failed to fetch ${path}: ${res.status}`);
  const buf = await res.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: true, cellNF: true, cellStyles: true });
  wbBuf = buf; // keep raw bytes for chart scan
  return { wb, buf };
}

function renderSheetList(workbook){
  sheetSel.innerHTML = "";
  workbook.SheetNames.forEach(name => {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    sheetSel.appendChild(opt);
  });
}

function renderSheet(name){
  currentWS = name;
  const ws = currentWB.Sheets[name];
  if (!ws) return;
  const range = XLSX.utils.decode_range(ws['!ref'] || "A1");
  const html = buildTable(ws, range);
  tableWrap.innerHTML = html;
  currentHTMLTable = tableWrap.firstChild;
  styleTableFromSheet(ws, currentHTMLTable);
}

/* --------- Events --------- */

fileSel.addEventListener("change", async ()=>{
  await openPath(fileSel.value);
});

sheetSel.addEventListener("change", async ()=>{
  renderSheet(sheetSel.value);
  await renderChartsForActiveSheet();
});

reloadBtn.addEventListener("click", async ()=>{
  await openPath(currentPath || DEFAULT_FILE, { force: true });
});

/* --------- Open path --------- */

async function openPath(path, { force } = {}){
  if (!force && path === currentPath) return;
  currentPath = path;
  setChartsStatus("");
  chartsOutEl().innerHTML = "";

  const { wb } = await loadWorkbookFromPath(path);
  currentWB = wb;
  initHFIfNeeded(wb);
  renderSheetList(wb);
  sheetSel.value = wb.SheetNames[0];
  renderSheet(sheetSel.value);

  await extractChartsFromXLSX(wbBuf);
  await renderChartsForActiveSheet();
}

/* --------- Chart extraction (XLSX unzip + XML parse) --------- */

/* resolve "xl/..." rel paths, handling ../ and ./ correctly */
function normalisePath(base, target){
  if (!target) return null;
  const clean = target.replace(/\\/g, "/");

  // absolute-within-zip
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
  // ESM fallback for environments where dynamic <script> is blocked by CSP
  try {
    const mod = await import("https://cdn.jsdelivr.net/npm/fflate@0.8.1/esm/browser.js");
    if (mod && typeof mod.unzipSync === "function") {
      globalThis.fflate = mod;
      return mod;
    }
  } catch (_) {}
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

    const get = (p) => zip[(p||"").toLowerCase()];
    const readText = (p) => {
      const u8 = get(p);
      if (!u8) return null;
      try {
        return new TextDecoder("utf-8").decode(u8);
      } catch {
        return null;
      }
    };

    const relsText = readText("xl/_rels/workbook.xml.rels");
    const workbookRels = relsText ? parseRels(relsText, "xl/_rels/workbook.xml.rels") : {};
    const wbText = readText("xl/workbook.xml");
    if (!wbText) {
      setChartsStatus("No workbook.xml found.");
      return;
    }
    const wbInfo = parseWorkbook(wbText);
    date1904BySheet = wbInfo.date1904BySheet || {};

    const sheetChartDefs = {};
    for (const sheetName of wbInfo.sheets){
      const wsPath = normalisePath("xl/worksheets/sheet1.xml", `xl/worksheets/${sheetName}.xml`);
      const relsPath = normalisePath(wsPath, wsPath.replace(/\.xml$/i, ".xml.rels").replace("/worksheets/", "/worksheets/_rels/"));
      const rels = parseRels(readText(relsPath) || "", relsPath);

      // scan drawing relationships to find charts
      Object.values(rels).forEach(rel => {
        if (rel.type.endsWith("/drawing")){
          const drawPath = normalisePath(wsPath, rel.target);
          const drawText = readText(drawPath);
          if (!drawText) return;

          const anchors = [...drawText.matchAll(/<xdr:(?:oneCellAnchor|twoCellAnchor)[\s\S]*?<\/xdr:(?:oneCellAnchor|twoCellAnchor)>/g)];
          const chartRefs = [];
          anchors.forEach(a => {
            const t = a[0];
            const m = /<c:chart [^>]*r:id="([^"]+)"/.exec(t);
            if (m) chartRefs.push(m[1]);
          });

          if (chartRefs.length){
            const drawRelsPath = normalisePath(drawPath, drawPath.replace(/\.xml$/i, ".xml.rels").replace("/drawings/", "/drawings/_rels/"));
            const drawRels = parseRels(readText(drawRelsPath) || "", drawRelsPath);

            const defs = [];
            chartRefs.forEach(rid => {
              const chartEntry = drawRels[rid];
              if (!chartEntry) return;
              const chartPath = normalisePath(drawPath, chartEntry.target);
              const chartText = readText(chartPath);
              if (!chartText) return;
              const chartObj = parseChartXML(chartText, chartPath);

              // title (optional)
              const title = extractChartTitle(chartText) || "";
              if (title) chartObj.title = title;

              defs.push(chartObj);
            });
            if (defs.length) sheetChartDefs[sheetName] = defs;
          }
        }
      });
    }

    chartsBySheet = sheetChartDefs;
    const total = Object.values(chartsBySheet).reduce((a, b) => a + b.length, 0);
    if (!total) setChartsStatus("No embedded charts found.");
  } catch (err){
    console.error("Chart extraction failed:", err);
    setChartsStatus("Chart extraction failed. See console for details.");
  }
}

/* parse workbook.xml minimal */
function parseWorkbook(xml){
  const out = { sheets: [], date1904BySheet: {} };
  const sRe = /<sheet\b[^>]*name="([^"]+)"[^>]*\/>/g;
  let m;
  while ((m = sRe.exec(xml))){
    out.sheets.push(m[1]);
  }
  // date1904
  const m1904 = /<workbookPr\b[^>]*date1904="(1|true)"/i.exec(xml);
  const use1904 = !!(m1904 && (m1904[1] === "1" || m1904[1] === "true"));
  out.sheets.forEach(n => out.date1904BySheet[n] = use1904);
  return out;
}

/* parse .rels into a map of id -> { target, type } */
function parseRels(xml, base){
  const map = {};
  const re = /<Relationship\b[^>]*Id="([^"]+)"[^>]*Type="([^"]+)"[^>]*Target="([^"]+)"[^>]*\/>/g;
  let m;
  while ((m = re.exec(xml))){
    const id = m[1], type = m[2], target = normalisePath(base, m[3]);
    map[id] = { type, target };
  }
  return map;
}

/* Extract basic info from chart XML to later render via Chart.js */
function parseChartXML(xml, chartPath){
  // series and categories
  const series = [];
  const serRe = /<c:ser\b[\s\S]*?<\/c:ser>/g;
  let m;
  while ((m = serRe.exec(xml))){
    const t = m[0];
    // cat / val refs
    const cat = /<c:cat\b[\s\S]*?<c:f>([\s\S]*?)<\/c:f>[\s\S]*?<\/c:cat>/.exec(t);
    const val = /<c:val\b[\s\S]*?<c:f>([\s\S]*?)<\/c:f>[\s\S]*?<\/c:val>/.exec(t);
    const tx  = /<c:tx\b[\s\S]*?(?:<c:v>([\s\S]*?)<\/c:v>|<c:f>([\s\S]*?)<\/c:f>)[\s\S]*?<\/c:tx>/.exec(t);
    const dp  = [...t.matchAll(/<c:dPt\b[\s\S]*?<\/c:dPt>/g)].map(x => x[0]);

    const ser = {
      catRef: cat ? cat[1].trim() : null,
      valRef: val ? val[1].trim() : null,
      txRef:  tx  ? (tx[1] || tx[2] || "").trim() : null,
      dataPoints: dp.length
    };
    series.push(ser);
  }

  // chart type (basic)
  const typeMap = {
    barChart: "bar",
    bar3DChart: "bar",
    lineChart: "line",
    line3DChart: "line",
    areaChart: "line",
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

  const typeM = /<c:chartSpace[\s\S]*?<c:plotArea[\s\S]*?<c:([a-zA-Z0-9]+)\/?>/m.exec(xml);
  const chartType = typeM ? (typeMap[typeM[1]] || "bar") : "bar";

  // title (we'll extract again in a helper, but keep a placeholder)
  const title = extractChartTitle(xml) || "";

  return { path: chartPath, type: chartType, series, title };
}

/* extract chart title (text or ref) */
function extractChartTitle(xml){
  const titleRe = /<c:title\b[\s\S]*?<\/c:title>/m;
  const m = titleRe.exec(xml);
  if (!m) return "";
  const block = m[0];

  // simple v text
  const vm = /<c:v>([\s\S]*?)<\/c:v>/.exec(block);
  if (vm) return (vm[1] || "").trim();

  // fallback to rich text runs
  const runs = [...block.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map(x => x[1].trim()).filter(Boolean);
  if (runs.length) return runs.join(" ");

  // as a last try, a name ref (rare)
  const s = /<c:tx\b[\s\S]*?<c:strRef\b[\s\S]*?<c:f>([\s\S]*?)<\/c:f>[\s\S]*?<\/c:strRef>[\s\S]*?<\/c:tx>/.exec(block);
  if (s && s[1]) {
    const t = s[1].trim();
    if (t) return t;
  }
  if (s && s.nameV) return s.nameV;
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
    cnv.setAttribute("aria-label", title || `Chart ${idx+1}`);
    cnv.setAttribute("role", "img");
    cnv.height = 260;
    cnv.width = 420;

    const card = document.createElement("div");
    card.className = "chart-card";
    const h = document.createElement("h4");
    h.textContent = title || def.type.toUpperCase();
    card.appendChild(h);
    card.appendChild(cnv);
    co.appendChild(card);

    try {
      renderChartOnCanvas(def, cnv);
    } catch (e) {
      const err = document.createElement("div");
      err.className = "chart-error";
      err.textContent = `Failed to render chart: ${e.message || e}`;
      card.appendChild(err);
    }
  });
}

/* Evaluate title from chart def (txRef or fallback) */
function evalTitle(def){
  if (def.title) return def.title;
  // TODO: if txRef points to a cell, resolve it via sheet data
  return "";
}

/* Load Chart.js (+ adapters) on demand once */
let chartLibReady = null;
async function loadChartLibsOnce(){
  if (chartLibReady) return chartLibReady;
  chartLibReady = (async ()=>{
    if (!window.Chart) {
      await loadScript("https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js");
    }
    // optional plugins for extra types could be loaded here
  })();
  return chartLibReady;
}

function loadScript(src){
  return new Promise((resolve, reject)=>{
    const s = document.createElement("script");
    s.src = src;
    s.async = true;
    s.onload = resolve;
    s.onerror = () => reject(new Error(`Failed to load ${src}`));
    document.head.appendChild(s);
  });
}

function destroyAllCharts(){
  chartInstances.forEach(ch => { try { ch.destroy(); } catch {} });
  chartInstances = [];
}

/* Render a basic Chart.js chart from a parsed chart def */
function renderChartOnCanvas(def, canvas){
  const ctx = canvas.getContext("2d");
  const type = def.type || "bar";

  // mock dataset until we wire ranges -> values
  // In a real build, we'd resolve catRef / valRef against the sheet data
  const labels = Array.from({length: def.series.length || 3}, (_,i)=> `S${i+1}`);
  const data = {
    labels,
    datasets: def.series.map((s, i) => ({
      label: s.txRef || `Series ${i+1}`,
      data: Array.from({length: labels.length}, ()=> Math.round(Math.random()*100)),
    }))
  };

  const ch = new Chart(ctx, {
    type: ["histogram","boxWhisker","stock","waterfall","sunburst","treemap","surface","combo"].includes(type) ? "bar" : type,
    data,
    options: {
      responsive: false,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: true },
        title: { display: !!def.title, text: def.title || "" }
      },
      scales: {
        x: { ticks: { autoSkip: true, maxRotation: 0 } },
        y: { beginAtZero: true }
      }
    }
  });
  chartInstances.push(ch);
}

/* ---------- Boot ---------- */

(async function boot(){
  try {
    await openPath(DEFAULT_FILE);
  } catch (e) {
    console.error(e);
    setChartsStatus(`Failed to open default file: ${e.message || e}`);
  }
})();
