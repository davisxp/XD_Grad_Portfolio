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

/* --------- Helpers --------- */
function status(msg){ if (statusEl) statusEl.textContent = msg || ""; }
function htmlEscape(s){ return String(s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c])); }

function parseCellAddress(a1){
  try { return XLSX.utils.decode_cell(a1); } catch { return null; }
}

/* use window.deps prepared by index.html */
const deps = window.deps || {};
if (!deps || !deps.loadWorkbook) {
  console.warn("[viewer] deps not available; ensure index.html script block ran before viewer.js");
}

/* --------- Init UI --------- */
(function initUI(){
  // files dropdown
  fileSel.innerHTML = "";
  files.forEach(f=>{
    const o = document.createElement("option");
    o.value = f.path; o.textContent = f.label;
    fileSel.appendChild(o);
  });

  // scenario dropdown
  scenarioSel.innerHTML = "";
  SCENARIO_OPTS.forEach(s=>{
    const o = document.createElement("option");
    o.value = s; o.textContent = s;
    scenarioSel.appendChild(o);
  });

  fileSel.value = files[0].path;
})();

/* ----- Rendering the grid ----- */
function renderActiveSheet(){
  const wsname = sheetSel.value;
  const ws = currentWB.Sheets[wsname];
  if (!ws) return;

  const { table } = deps.renderSheet({ ws, wsname, wb: currentWB, styleCtx, showFormula: showForm.checked, filterText: filterInp.value || "" });
  currentHTMLTable = table;

  // Wire inspector
  table.addEventListener("click", (e)=>{
    const td = e.target.closest("td[data-a1]");
    if (!td) return;
    const addr = td.getAttribute("data-a1");
    showCellInfo(wsname, addr);
  }, { passive:true });
}

function showCellInfo(sheet, a1){
  if (!hf || !cellInfo) return;
  const id = hf.getSheetId(sheet);
  if (id == null) return;

  const { r, c } = XLSX.utils.decode_cell(a1);
  const { r0, c0 } = sheetOffsets[sheet] || { r0:0, c0:0 };
  const v = hf.getCellValue({ sheet:id, row:r - r0, col:c - c0 });
  const f = hf.getCellFormula({ sheet:id, row:r - r0, col:c - c0 });

  cellInfo.innerHTML = `
    <div class="kv"><div>Cell</div><div>${htmlEscape(sheet)}!${htmlEscape(a1)}</div></div>
    <div class="kv"><div>Value</div><div>${htmlEscape(String(v))}</div></div>
    <div class="kv"><div>Formula</div><div><code>${htmlEscape(f || "")}</code></div></div>
  `;
}

/* ----- URL sync ----- */
function syncLink(){
  const url = new URL(location.href);
  url.searchParams.set("file", fileSel.value);
  url.searchParams.set("sheet", sheetSel.value);
  url.searchParams.set("sc", scenarioSel.value);
  history.replaceState(null, "", url.toString());
}

/* ----- Events ----- */
fileSel.addEventListener("change", async ()=>{
  await openPath(fileSel.value);
});
sheetSel.addEventListener("change", ()=>{
  renderActiveSheet();
  renderChartsForActiveSheet();
  syncLink();
});
scenarioSel.addEventListener("change", ()=>{
  applyScenarioToHF(scenarioSel.value);
  renderActiveSheet();
  renderChartsForActiveSheet();
});
showForm.addEventListener("change", ()=>{
  renderActiveSheet();
});
filterInp.addEventListener("input", ()=>{
  renderActiveSheet();
});
copyBtn.addEventListener("click", ()=>{
  navigator.clipboard.writeText(location.href).catch(()=>{});
});
dlLink.addEventListener("click", (e)=>{
  dlLink.href = fileSel.value;
});

/* ----- HF + Scenario wiring ----- */
async function openPath(path){
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
    out.innerHTML = `<div class="error">${e.message}. Check the file path (case-sensitive on most servers).</div>`;
    status("");
  }
}

function applyScenarioToHF(scenarioName){
  // write scenario into target cells
  SCENARIOS.forEach(({sheet, targets})=>{
    const id = hf.getSheetId(sheet);
    const { r0, c0 } = sheetOffsets[sheet] || { r0:0, c0:0 };
    targets.forEach(addr => {
      const { r, c } = XLSX.utils.decode_cell(addr);
      hf.setCellContents({ sheet: id, row: r - r0, col: c - c0 }, [[scenarioName]]);
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
      applyScenarioToHF(qSc);
    }
  } catch {}
}

/* ===================== CHARTS ===================== */

let chartsBySheet = {};
let wbBuf = null;
let chartInstances = [];

function ensureChartsHost(){
  const wrap = document.getElementById("sheetWrap");
  if (!document.getElementById("chartsOut")){
    const div = document.createElement("div");
    div.id = "chartsOut";
    div.className = "charts-grid";
    wrap.appendChild(div);
  }
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
      const s = document.createElement("script"); s.src = src; s.async = true;
      s.onload = res; s.onerror = () => rej(new Error("Failed to load " + src));
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
      // optional
    }
    if (pluginModule && pluginModule.CandlestickController) {
      Chart.register(pluginModule.CandlestickController, pluginModule.OHLCController, pluginModule.CandlestickElement, pluginModule.OHLCElement);
    }
  })();
  return chartLibPromise;
}

const chartTypeMap = {
  barChart: "bar",
  bar3DChart: "bar",
  lineChart: "line",
  line3DChart: "line",
  areaChart: "area",
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

    const wbXml = zip["xl/workbook.xml"]; const relsXml = zip["xl/_rels/workbook.xml.rels"];
    if(!wbXml || !relsXml){ setChartsStatus("No charts"); return {}; }

    const tdDecoder = new TextDecoder("utf-8");
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

    function byLocal(doc, local){
      return Array.from(doc.getElementsByTagNameNS("*", local)).concat(Array.from(doc.getElementsByTagName(local)));
    }
    function firstDescendantWithLocal(node, locals){
      for (const l of locals){
        const el = node.querySelector(`*:is(${l}, *|${l})`);
        if (el) return el;
      }
      return null;
    }
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
            buckets[idx] = buckets[idx] || new Array(depth+1).fill("");
            const txt = member.textContent || "";
            buckets[idx][depth] = txt;
          });
        });
        return buckets.map(arr => arr.join(" / "));
      } else {
        byLocal(node, "pt").forEach(pt=>{
          const idx = parseInt(pt.getAttribute("idx") ?? `${pts.length}`, 10);
          pts[idx] = pt.textContent || "";
        });
        return pts.filter(v=>v!==undefined);
      }
    }
    function readXY(refNode){
      if (!refNode) return { refF:null, data:null };
      const fNode = firstDescendantWithLocal(refNode, ["f"]);
      if (fNode && fNode.textContent) return { refF: fNode.textContent.trim(), data: null };
      const data = collectPoints(refNode);
      return { refF:null, data };
    }
    function readSeriesName(txNode){
      if (!txNode) return { nameF:null, nameV:null };
      const ref = firstDescendantWithLocal(txNode, ["strRef","strData"]);
      if (ref){
        const fNode = firstDescendantWithLocal(ref, ["f"]);
        if (fNode && fNode.textContent) return { nameF: fNode.textContent.trim(), nameV: null };
        const data = collectPoints(ref);
        if (data && data.length) return { nameF:null, nameV:String(data[0] ?? "") };
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
        const fNode = firstDescendantWithLocal(ref, ["f"]);
        if (fNode && fNode.textContent) return { titleF: fNode.textContent.trim(), titleText: null };
        const data = collectPoints(ref);
        if (data && data.length) return { titleF:null, titleText:String(data[0] ?? "") };
      }
      const vNode = firstDescendantWithLocal(titleNode, ["v","t"]);
      if (vNode && vNode.textContent) return { titleF:null, titleText: vNode.textContent.trim() };
      return { titleF:null, titleText:null };
    }

    const chartsBySheetLocal = {};

    for (const { name: sheetName, path: sheetPath, kind } of meta){
      if (kind !== "worksheet" && kind !== "chartsheet") continue;

      // read drawing rels from the sheet rels
      const relsPath = sheetPath.replace(/(^xl\/)(worksheets|chartsheets)\//, "$1$2/_rels/").replace(/\.xml$/i, ".xml.rels");
      const relsXml2 = zip[relsPath];
      const rels = {};
      if (relsXml2){
        const rdoc2 = parse(relsXml2);
        Array.from(rdoc2.getElementsByTagName("Relationship")).forEach(r=>{
          rels[r.getAttribute("Id")] = r.getAttribute("Target");
        });
      }

      const drawingTargets = Object.values(rels).filter(t => /drawing/i.test(t));
      const defs = [];

      for (const drawT of drawingTargets){
        const drawPath = normalisePath(sheetPath, drawT);
        const drawingXml = zip[drawPath];
        if (!drawingXml) continue;
        const ddoc = parse(drawingXml);

        // map drawing rId -> chart path
        const dRelsPath = `xl/drawings/_rels/${drawPath.split("/").pop()}.rels`;
        const dRelsXml = zip[dRelsPath];
        const dRels = {};
        if (dRelsXml){
          const ddoc2 = parse(dRelsXml);
          Array.from(ddoc2.getElementsByTagName("Relationship")).forEach(r=>{
            const relTarget = r.getAttribute("Target");
            const relPath = normalisePath(drawPath, relTarget);
            if (relPath) dRels[r.getAttribute("Id")] = relPath;
          });
        }

        const chartElems = ddoc.getElementsByTagNameNS("*", "chart"); // catches c:chart and cx:chart
        for (const cEl of chartElems){
          const rid = cEl.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships","id") || cEl.getAttribute("r:id");
          const chartPath = dRels[rid];
          if (!chartPath) continue;
          const chartXml = zip[chartPath];
          if (!chartXml) continue;
          const cdoc = parse(chartXml);

          let chartType = "bar";
          const plotArea = cdoc.getElementsByTagNameNS("*","plotArea")[0];
          if (plotArea){
            const firstChartElem = Array.from(plotArea.children).find(n => /chart$/i.test(n.localName));
            if (firstChartElem) chartType = chartTypeMap[firstChartElem.localName] || "bar";
          }
          const title = extractTitle(cdoc);

          function firstDescendantWithLocal2(node, locals){
            for (const l of locals){
              const el = node.querySelector(`*:is(${l}, *|${l})`);
              if (el) return el;
            }
            return null;
          }
          function readXY2(refNode){
            if (!refNode) return { refF:null, data:null };
            const fNode = firstDescendantWithLocal2(refNode, ["f"]);
            if (fNode && fNode.textContent) return { refF: fNode.textContent.trim(), data: null };
            const data = (function collect(node){
              if (!node) return null;
              const pts = [];
              const lvlNodes = Array.from(node.querySelectorAll("*|lvl, lvl"));
              if (lvlNodes.length){
                const buckets = [];
                lvlNodes.forEach((lvl, depth)=>{
                  const members = Array.from(lvl.querySelectorAll("*|pt, pt"));
                  members.forEach(member=>{
                    const idx = parseInt(member.getAttribute("idx") ?? `${buckets.length}`, 10);
                    buckets[idx] = buckets[idx] || new Array(depth+1).fill("");
                    const txt = member.textContent || "";
                    buckets[idx][depth] = txt;
                  });
                });
                return buckets.map(arr => arr.join(" / "));
              } else {
                Array.from(node.querySelectorAll("*|pt, pt")).forEach(pt=>{
                  const idx = parseInt(pt.getAttribute("idx") ?? `${pts.length}`, 10);
                  pts[idx] = pt.textContent || "";
                });
                return pts.filter(v=>v!==undefined);
              }
            })(refNode);
            return { refF:null, data };
          }
          function readSeriesName2(txNode){
            if (!txNode) return { nameF:null, nameV:null };
            const ref = firstDescendantWithLocal2(txNode, ["strRef","strData"]);
            if (ref){
              const fNode = firstDescendantWithLocal2(ref, ["f"]);
              if (fNode && fNode.textContent) return { nameF: fNode.textContent.trim(), nameV: null };
              const data = (function collect(node){
                if (!node) return null;
                const pts = [];
                Array.from(node.querySelectorAll("*|pt, pt")).forEach(pt=>{
                  const idx = parseInt(pt.getAttribute("idx") ?? `${pts.length}`, 10);
                  pts[idx] = pt.textContent || "";
                });
                return pts.filter(v=>v!==undefined);
              })(ref);
              if (data && data.length) return { nameF:null, nameV:String(data[0] ?? "") };
            }
            const vNode = firstDescendantWithLocal2(txNode, ["v","t","r"]);
            if (vNode && vNode.textContent) return { nameF:null, nameV: vNode.textContent.trim() };
            const text = txNode.textContent?.trim();
            return text ? { nameF:null, nameV:text } : { nameF:null, nameV:null };
          }

          const serNodes = Array.from(cdoc.querySelectorAll("c\\:ser, cx\\:ser, ser"));
          const series = serNodes.map(ser=>{
            const tx = ser.querySelector("*|tx, tx");
            const xRef = ser.querySelector("*|cat, *|xVal, *|xValues, cat, xVal, xValues");
            const yRef = ser.querySelector("*|val, *|yVal, *|yValues, val, yVal, yValues");
            const zRef = ser.querySelector("*|zVal, *|zValues, zVal, zValues");
            const lbls = ser.querySelector("*|dLbls, *|dLabels, dLbls, dLabels");
            const lRef = lbls ? (lbls.querySelector("*|numRef, *|strRef, *|numLit, *|strLit, numRef, strRef, numLit, strLit")) : null;

            const { nameF, nameV } = readSeriesName2(tx);
            const { refF: xF, data: xData } = xRef ? readXY2(xRef) : { refF:null, data:null };
            const { refF: yF, data: yData } = yRef ? readXY2(yRef) : { refF:null, data:null };
            const { refF: zF, data: zData } = zRef ? readXY2(zRef) : { refF:null, data:null };
            const labels = lRef ? readXY2(lRef) : { refF:null, data:null };

            return {
              txRef: nameF, txText: nameV,
              xRef: xF, xData,
              yRef: yF, yData,
              zRef: zF, zData,
              labelsRef: labels.refF, labelsData: labels.data
            };
          });

          defs.push({
            type: chartType,
            title: title.titleText || "",
            titleRef: title.titleF || null,
            series
          });
        }
      }

      if (defs.length) chartsBySheetLocal[sheetName] = defs;
    }

    return chartsBySheetLocal;
  } catch (err){
    console.error("Chart extraction failed:", err);
    setChartsStatus("Chart extraction failed. See console for details.");
    return {};
  }
}

/* Render charts for selected sheet */
function renderChartsForActiveSheet(){
  ensureChartsHost();

  const sheet = sheetSel.value;
  const defs = chartsBySheet[sheet] || [];
  const host = chartsOutEl();
  if (!host) return;

  host.innerHTML = "";
  if (!defs.length){
    setChartsStatus("No charts found on this sheet.");
    return;
  }

  loadChartLibsOnce().then(()=>{
    defs.slice(0, 12).forEach((def, idx) => {
      const title = evalTitle(def);
      const cnv = document.createElement("canvas");
      cnv.setAttribute("aria-label", title || `Chart ${idx+1}`);
      cnv.setAttribute("role", "img");
      cnv.height = 260;

      const card = document.createElement("div");
      card.className = "chart-card";
      const h = document.createElement("h4");
      h.textContent = title || def.type.toUpperCase();
      card.appendChild(h);
      card.appendChild(cnv);
      host.appendChild(card);

      try {
        renderChartOnCanvas(def, cnv);
      } catch (e) {
        const err = document.createElement("div");
        err.className = "chart-error";
        err.textContent = `Failed to render chart: ${e.message || e}`;
        card.appendChild(err);
      }
    });
  });
}

function evalTitle(def){
  if (def.title) return def.title;
  if (def.titleRef){
    const pr = parseSheetAndRange(def.titleRef);
    if (pr){
      const vals = getRangeValues(pr.sheet, pr.range);
      const flat = rangeToVector(vals);
      const first = flat.find(v => v != null);
      if (first != null) return String(first);
    }
  }
  return "";
}

function destroyAllCharts(){
  chartInstances.forEach(ch => { try { ch.destroy(); } catch {} });
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

function evalSeriesName(ser){
  return ser.txText || (function(){
    if (!ser.txRef) return "";
    const pr = parseSheetAndRange(ser.txRef);
    if (!pr) return "";
    const vals = rangeToVector(getRangeValues(pr.sheet, pr.range));
    return String((vals.find(v=>v!=null) ?? ""));
  })();
}

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

function renderChartOnCanvas(def, cnv){
  const title = evalTitle(def);

  function resolveLabels(s){
    if (s.labelsRef){
      const pr = parseSheetAndRange(s.labelsRef);
      if (pr) return rangeToVector(getRangeValues(pr.sheet, pr.range)).map(v=>String(v ?? ""));
    }
    if (Array.isArray(s.labelsData)) return s.labelsData.map(v=>String(v ?? ""));
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

  } else if (def.type === "scatter" || def.type === "bubble") {
    const ds = def.series.map((s, i)=>{
      const x = resolveX(s), y = resolveY(s);
      const z = resolveZ(s);
      const data = x.map((xx, idx)=> ({ x: Number(xx), y: Number(y[idx] ?? 0), r: Number(z[idx] ?? 3) }));
      return { label: evalSeriesName(s) || ("Series " + (i+1)), data, showLine: false };
    });
    const cfg = { type: "scatter", data: { datasets: ds }, options: { responsive: true, plugins: { title: { display: !!title, text: title } } } };
    chartInstances.push(new Chart(cnv.getContext("2d"), cfg));

  } else if (def.type === "histogram"){
    const s0 = def.series[0] || {};
    const x = resolveX(s0);
    const y = resolveY(s0);
    let labels = [];
    let counts = [];
    if (y.length){
      labels = x.map(v=>String(v ?? ""));
      counts = y;
    } else {
      const values = y.length ? y : x.map(Number);
      const h = buildHistogram(values);
      labels = h.labels; counts = h.counts;
    }
    const cfg = { type: "bar", data: { labels, datasets:[{ label: title || "Histogram", data: counts }] },
      options:{ responsive:true, plugins:{ title:{ display: !!title, text: title } } } };
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

    const sOpen = s[idxOpen] || s[0] || {};
    const sHigh = s[idxHigh] || s[1] || {};
    const sLow  = s[idxLow]  || s[2] || {};
    const sClose= s[idxClose]|| s[3] || {};

    const labels = resolveLabels(sClose);
    const open = resolveY(sOpen), high = resolveY(sHigh), low = resolveY(sLow), close = resolveY(sClose);

    const data = labels.map((lab, i)=>({ t: lab, o: open[i] ?? null, h: high[i] ?? null, l: low[i] ?? null, c: close[i] ?? null }));
    const cfg = { type: "bar", data: { labels, datasets:[{ label: title || "OHLC", data: data.map(d=>d.c) }] }, options:{ plugins:{ title:{ display: !!title, text: title } } } };
    chartInstances.push(new Chart(cnv.getContext("2d"), cfg));

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
}

/* ===================== END CHARTS ===================== */

/* ---------- Boot ---------- */
(async function boot(){
  try {
    await openPath(fileSel.value);
  } catch (e) {
    console.error(e);
    status(`Failed to open default file: ${e.message || e}`);
  }
})();

/* ===== Patch v2: Remove CDN dependency by using native DecompressionStream to read XLSX =====
   This overrides extractChartsFromXLSX() to prefer a CSP-safe, no-network path.
   If DecompressionStream is unavailable, it will fall back to the previous fflate-based path.
----------------------------------------------------------------------------------------------*/

async function unzipWithNative(arrayBuffer){
  if (typeof DecompressionStream === "undefined") throw new Error("DecompressionStream not available");
  const bytes = new Uint8Array(arrayBuffer);
  const len = bytes.length;

  // Find End of Central Directory (EOCD) signature 0x06054b50
  function readU32LE(i){ return (bytes[i] | (bytes[i+1]<<8) | (bytes[i+2]<<16) | (bytes[i+3]<<24)) >>> 0; }
  function readU16LE(i){ return (bytes[i] | (bytes[i+1]<<8)) >>> 0; }

  const eocdSig = 0x06054b50 >>> 0;
  let eocd = -1;
  const maxBack = Math.min(len, 22 + 65536 + 1024); // EOCD + max comment + buffer
  for (let i = len - 22; i >= len - maxBack; i--){
    if (i < 0) break;
    if (readU32LE(i) === eocdSig){ eocd = i; break; }
  }
  if (eocd < 0) throw new Error("ZIP EOCD not found");

  const cdSize   = readU32LE(eocd + 12);
  const cdOffset = readU32LE(eocd + 16);
  const cdEnd = cdOffset + cdSize;

  // Iterate Central Directory entries
  const cdSig = 0x02014b50 >>> 0;
  const lfSig = 0x04034b50 >>> 0;

  const decoder = new TextDecoder("utf-8");
  const entries = [];
  let p = cdOffset;
  while (p + 46 <= cdEnd && readU32LE(p) === cdSig){
    const compression = readU16LE(p + 10);
    const compSize = readU32LE(p + 20);
    const uncompSize = readU32LE(p + 24);
    const nameLen = readU16LE(p + 28);
    const extraLen = readU16LE(p + 30);
    const commentLen = readU16LE(p + 32);
    const localHeaderOffset = readU32LE(p + 42);

    const nameBytes = bytes.subarray(p + 46, p + 46 + nameLen);
    const name = decoder.decode(nameBytes);

    entries.push({ name, compression, compSize, uncompSize, localHeaderOffset });
    p += 46 + nameLen + extraLen + commentLen;
  }

  async function inflateRaw(u8){
    const ds = new DecompressionStream('deflate-raw');
    const resp = new Response(new Blob([u8]).stream().pipeThrough(ds));
    const ab = await resp.arrayBuffer();
    return new Uint8Array(ab);
  }

  const out = {};
  for (const ent of entries){
    if (readU32LE(ent.localHeaderOffset) !== lfSig) continue;
    const nlen = readU16LE(ent.localHeaderOffset + 26);
    const xlen = readU16LE(ent.localHeaderOffset + 28);
    const dataStart = ent.localHeaderOffset + 30 + nlen + xlen;
    const comp = bytes.subarray(dataStart, dataStart + ent.compSize);
    let u8;
    if (ent.compression === 0){ // stored
      u8 = comp.slice();
    } else if (ent.compression === 8){ // deflate
      u8 = await inflateRaw(comp);
    } else {
      // Unsupported method; skip quietly
      continue;
    }
    out[ent.name.toLowerCase()] = u8;
  }
  return out;
}

// Keep a handle to the old implementation as fallback
const __extractChartsFromXLSX_fflate = extractChartsFromXLSX;

/** Override: prefer native unzip, fallback to fflate */
async function extractChartsFromXLSX(arrayBuffer){
  // native path
  if (typeof DecompressionStream !== "undefined"){
    try{
      const zip = await unzipWithNative(arrayBuffer);
      const tdDecoder = new TextDecoder("utf-8");

      const get = (p) => zip[(p||"").toLowerCase()];
      const readXML = (p) => {
        const u8 = get(p);
        return u8 ? tdDecoder.decode(u8) : null;
      };

      const wbXml = readXML("xl/workbook.xml");
      const relsXml = readXML("xl/_rels/workbook.xml.rels");
      if (!wbXml || !relsXml){
        setChartsStatus("No workbook parts found.");
        return {};
      }

      const parse = (xml) => (new DOMParser()).parseFromString(xml, "application/xml");
      const wdoc = parse(wbXml);
      const rdoc = parse(relsXml);

      // sheet -> target map from workbook rels
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

      function byLocal(doc, local){
        return Array.from(doc.getElementsByTagNameNS("*", local)).concat(Array.from(doc.getElementsByTagName(local)));
      }
      function firstDescendantWithLocal(node, locals){
        for (const l of locals){
          const el = node.querySelector(`*:is(${l}, *|${l})`);
          if (el) return el;
        }
        return null;
      }
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
              buckets[idx] = buckets[idx] || new Array(depth+1).fill("");
              const txt = member.textContent || "";
              buckets[idx][depth] = txt;
            });
          });
          return buckets.map(arr => arr.join(" / "));
        } else {
          byLocal(node, "pt").forEach(pt=>{
            const idx = parseInt(pt.getAttribute("idx") ?? `${pts.length}`, 10);
            pts[idx] = pt.textContent || "";
          });
          return pts.filter(v=>v!==undefined);
        }
      }

      function readXY(refNode){
        if (!refNode) return { refF:null, data:null };
        const fNode = firstDescendantWithLocal(refNode, ["f"]);
        if (fNode && fNode.textContent) return { refF: fNode.textContent.trim(), data: null };
        const data = collectPoints(refNode);
        return { refF:null, data };
      }
      function readSeriesName(txNode){
        if (!txNode) return { nameF:null, nameV:null };
        const ref = firstDescendantWithLocal(txNode, ["strRef","strData"]);
        if (ref){
          const fNode = firstDescendantWithLocal(ref, ["f"]);
          if (fNode && fNode.textContent) return { nameF: fNode.textContent.trim(), nameV: null };
          const data = collectPoints(ref);
          if (data && data.length) return { nameF:null, nameV:String(data[0] ?? "") };
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
          const fNode = firstDescendantWithLocal(ref, ["f"]);
          if (fNode && fNode.textContent) return { titleF: fNode.textContent.trim(), titleText: null };
          const data = collectPoints(ref);
          if (data && data.length) return { titleF:null, titleText:String(data[0] ?? "") };
        }
        const vNode = firstDescendantWithLocal(titleNode, ["v","t"]);
        if (vNode && vNode.textContent) return { titleF:null, titleText: vNode.textContent.trim() };
        return { titleF:null, titleText:null };
      }

      const chartTypeMap = {
        barChart: "bar", bar3DChart: "bar", lineChart: "line", line3DChart: "line",
        areaChart: "area", scatterChart: "scatter", bubbleChart: "bubble",
        pieChart: "pie", pie3DChart: "pie", doughnutChart: "doughnut",
        radarChart: "radar", histogramChart: "histogram", paretoChart: "bar",
        stockChart: "stock", waterfallChart: "waterfall", funnelChart: "funnel",
        boxWhiskerChart: "boxWhisker", sunburstChart: "sunburst", treemapChart: "treemap",
        surfaceChart: "surface", surface3DChart: "surface", wireframeSurfaceChart: "surface",
        wireframeSurface3DChart: "surface", comboChart: "combo", ofPieChart: "pie"
      };

      const chartsBySheetLocal = {};

      for (const { name: sheetName, path: sheetPath, kind } of meta){
        if (kind !== "worksheet" && kind !== "chartsheet") continue;

        // read drawing rels from the sheet rels
        const relsPath = sheetPath.replace(/(^xl\/)(worksheets|chartsheets)\//, "$1$2/_rels/").replace(/\.xml$/i, ".xml.rels");
        const relsXml2 = readXML(relsPath);
        const rels = {};
        if (relsXml2){
          const rdoc2 = parse(relsXml2);
          Array.from(rdoc2.getElementsByTagName("Relationship")).forEach(r=>{
            rels[r.getAttribute("Id")] = r.getAttribute("Target");
          });
        }

        const drawingTarget = rels["rId1"]; // simple case; otherwise scan for Type contains '/drawing'
        const drawingTargets = Object.values(rels).filter(t => /drawing/i.test(t));
        const drawList = drawingTargets.length ? drawingTargets : (drawingTarget ? [drawingTarget] : []);
        const defs = [];

        for (const drawT of drawList){
          const drawPath = normalisePath(sheetPath, drawT);
          const drawingXml = readXML(drawPath);
          if (!drawingXml) continue;
          const ddoc = parse(drawingXml);

          // map drawing rId -> chart path
          const dRelsPath = drawPath.replace(/(^xl\/)drawings\//, "$1drawings/_rels/") + ".rels";
          const dRelsXml = readXML(dRelsPath);
          const dRels = {};
          if (dRelsXml){
            const ddoc2 = parse(dRelsXml);
            Array.from(ddoc2.getElementsByTagName("Relationship")).forEach(r=>{
              const relTarget = r.getAttribute("Target");
              const relPath = normalisePath(drawPath, relTarget);
              if (relPath) dRels[r.getAttribute("Id")] = relPath;
            });
          }

          const chartElems = ddoc.getElementsByTagNameNS("*", "chart");
          for (const cEl of chartElems){
            const rid = cEl.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships","id") || cEl.getAttribute("r:id");
            const chartPath = dRels[rid];
            if (!chartPath) continue;
            const chartXml = readXML(chartPath);
            if (!chartXml) continue;
            const cdoc = parse(chartXml);

            let chartType = "bar";
            const plotArea = cdoc.getElementsByTagNameNS("*","plotArea")[0];
            if (plotArea){
              const firstChartElem = Array.from(plotArea.children).find(n => /chart$/i.test(n.localName));
              if (firstChartElem) chartType = chartTypeMap[firstChartElem.localName] || "bar";
            }
            const title = extractTitle(cdoc);

            function readSerNodes(selector){
              return Array.from(cdoc.querySelectorAll(selector));
            }
            function readSeriesFromNodes(nodes){
              const arr = [];
              for (const ser of nodes){
                const tx    = firstDescendantWithLocal(ser, ["tx"]);
                const xRef  = firstDescendantWithLocal(ser, ["cat","xVal","xValues"]);
                const yRef  = firstDescendantWithLocal(ser, ["val","yVal","yValues"]);
                const zRef  = firstDescendantWithLocal(ser, ["zVal","zValues"]);
                const lbls  = firstDescendantWithLocal(ser, ["dLbls","dLabels"]);
                const lRef  = lbls ? firstDescendantWithLocal(lbls, ["numRef","strRef","numLit","strLit"]) : null;

                const { nameF, nameV } = readSeriesName(tx);
                const { refF: xF, data: xData } = xRef ? readXY(xRef) : { refF:null, data:null };
                const { refF: yF, data: yData } = yRef ? readXY(yRef) : { refF:null, data:null };
                const { refF: zF, data: zData } = zRef ? readXY(zRef) : { refF:null, data:null };
                const labels = lRef ? readXY(lRef) : { refF:null, data:null };

                arr.push({
                  txRef: nameF, txText: nameV,
                  xRef: xF, xData,
                  yRef: yF, yData,
                  zRef: zF, zData,
                  labelsRef: labels.refF, labelsData: labels.data
                });
              }
              return arr;
            }

            // Gather series across known chart node types
            const serNodes = Array.from(cdoc.querySelectorAll("c\\:ser, cx\\:ser, ser"));
            const series = readSeriesFromNodes(serNodes);

            defs.push({
              type: chartType,
              title: title.titleText || "",
              titleRef: title.titleF || null,
              series
            });
          }
        }

        if (defs.length) chartsBySheetLocal[sheetName] = defs;
      }

      return chartsBySheetLocal;
    } catch (err){
      console.warn("[charts] Native unzip failed; falling back to fflate:", err);
      // fall through to fflate path
    }
  }

  // fallback to original fflate-based implementation
  return await __extractChartsFromXLSX_fflate(arrayBuffer);
}
