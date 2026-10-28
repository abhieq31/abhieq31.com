/* AdigiPQR worker - the heavy work (reading the export, writing .xlsx, QA's workbook) runs here, off the page's
   main thread, so the page never freezes. Same code as the page (core.js, compare.js); nothing is sent anywhere:
   the worker is part of this tab and has no network access of its own. Each step is reported to the page. */
/* Also loaded by the page itself: there it only registers the same jobs (AdigiJobs) for browsers without workers. */
(function () {
"use strict";
var inWorker = typeof WorkerGlobalScope !== "undefined" && self instanceof WorkerGlobalScope;
if (inWorker) self.AdigiProgress = function (key, detail) { self.postMessage({ step: key, detail: detail || "" }); };

function step(key, detail) { if (self.AdigiProgress) self.AdigiProgress(key, detail); }
function ab(b) { return b instanceof ArrayBuffer ? b : b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength); }   // transferable

async function sha256(buf) {
  try {
    var d = await crypto.subtle.digest("SHA-256", buf);
    return Array.prototype.map.call(new Uint8Array(d), function (x) { return (x < 16 ? "0" : "") + x.toString(16); }).join("");
  } catch (e) { return null; }
}

async function readTable(buf, name) {
  if (/\.xlsx$/i.test(name)) {
    var wb = new ExcelJS.Workbook(); await wb.xlsx.load(buf);
    var ws = wb.worksheets[0], head = [], out = [];
    ws.eachRow({ includeEmpty: false }, function (row, i) {
      var vals = []; for (var c = 1; c <= ws.columnCount; c++) { var v = row.getCell(c).value;
        if (v && typeof v === "object") v = "result" in v ? v.result : v.text !== undefined ? v.text : v instanceof Date ? v.toISOString() : v;
        vals.push(v === null || v === undefined ? "" : String(v)); }
      if (i === 1) head = vals; else { var o = {}; head.forEach(function (h, j) { o[h] = vals[j]; }); out.push(o); }
    });
    return { rows: out, cols: head };
  }
  var res = Papa.parse(new TextDecoder("utf-8").decode(buf), { header: true, skipEmptyLines: true, dynamicTyping: false });
  return { rows: res.data, cols: res.meta.fields };
}

/* The export stays here: the page gets back only what it draws (no 10,000-row copies between the two). */
var DATA = null, SEL = null;
function s(v) { return v === null || v === undefined ? "" : String(v).trim(); }

/* QA asks for all batches of a material in a period, or for a list of batch numbers (material optional).
   Order: first manufacturing start, then batch number (= pivot.in_batch_order). Accepted scan records join their batches. */
function select(d) {
  var byBatch = d.mode === "batch", want = byBatch ? d.want : [], wantSet = {};
  want.forEach(function (b) { wantSet[b] = 1; });
  var mat = d.mat, from = byBatch ? "" : d.from, to = byBatch ? "" : d.to, lo = from ? from + " 00:00" : null, hi = to ? to + " 23:59" : null;
  var undated = {}, rows = DATA.rows.filter(function (r) {
    if (byBatch && !wantSet[r.BATCH_NO]) return false;
    if (mat && s(r.MATERIAL_CODE) !== mat) return false;
    if (!lo && !hi) return true;
    if (!r._start) { undated[r.BATCH_NO] = 1; return false; }
    return (!lo || r._start >= lo) && (!hi || r._start <= hi);
  });
  var first = {};
  rows.forEach(function (r) { if (r._start && (!first[r.BATCH_NO] || r._start < first[r.BATCH_NO])) first[r.BATCH_NO] = r._start; });
  var order = Object.keys(rows.reduce(function (a, r) { a[r.BATCH_NO] = 1; return a; }, {}))
    .sort(function (a, b) { var x = first[a] || "9999", y = first[b] || "9999"; return x < y ? -1 : x > y ? 1 : a < b ? -1 : a > b ? 1 : 0; });
  var pos = {}; order.forEach(function (b, i) { pos[b] = i; });
  rows = rows.slice().sort(function (a, b) { return pos[a.BATCH_NO] - pos[b.BATCH_NO]; });
  var scanRows = (d.scans || []).filter(function (row) {
    if (byBatch) return !!wantSet[row.BATCH_NO];
    var st = AdigiCore.startKey(row.FIRST_READING_ON);
    return pos[row.BATCH_NO] !== undefined || ((!mat || row.MATERIAL_CODE === mat) && (!lo || (st && st >= lo)) && (!hi || (st && st <= hi)));
  });
  var all = rows.concat(scanRows.map(function (r) { var o = Object.assign({}, r); o._scan = true; return o; }));
  all.sourceColumns = DATA.cols;
  return { rows: all, meta: { mat: mat, from: from, to: to, undated: Object.keys(undated), first: first, nScan: scanRows.length,
    byBatch: byBatch, want: want, missing: want.filter(function (b) { return pos[b] === undefined; }) } };
}

var OPS = {
  parse: async function (d) {
    step("hash", (d.buf.byteLength / 1048576).toFixed(1) + " MB");
    var sha = await sha256(d.buf);
    step("parse");
    var t = await readTable(d.buf, d.name);
    step("norm", t.rows.length.toLocaleString("en-IN") + " rows");
    var rows = AdigiCore.normalise(t.rows, t.cols), mats = {};
    rows.forEach(function (r) { var m = s(r.MATERIAL_CODE); if (!(m in mats)) mats[m] = s(r.PRODUCT_NAME); r._start = AdigiCore.startKey(r.MFG_START); });
    DATA = { rows: rows, cols: rows.sourceColumns }; SEL = null;
    return { n: rows.length, sha: sha, mats: mats };
  },
  // the overview: selection + QA sheets + statistics, and per batch what the batch table shows
  select: async function (d) {
    step("model");
    var sel = select(d); SEL = sel;
    var out = { meta: sel.meta, n: sel.rows.length };
    if (!sel.rows.length) return out;
    var model = AdigiCore.build(sel.rows, d.pivot), params = AdigiCore.parameters(model, d.cppList);
    var hdr = {}, oos = {}, oosRows = 0;
    sel.rows.forEach(function (r) {
      if (!r._scan && !hdr[r.BATCH_NO]) hdr[r.BATCH_NO] = { MATERIAL_CODE: r.MATERIAL_CODE, BATCH_SIZE: r.BATCH_SIZE, BATCH_SIZE_UOM: r.BATCH_SIZE_UOM,
        MBR_ID: r.MBR_ID, MBR_VERSION: r.MBR_VERSION, MANUFACTURING_ORDER: r.MANUFACTURING_ORDER, MO_STATUS: r.MO_STATUS };
      if (/^OUT/i.test(s(r.LIMIT_CHECK))) { oos[r.BATCH_NO] = (oos[r.BATCH_NO] || 0) + 1; oosRows++; }
    });
    out.model = model; out.params = params; out.hdr = hdr; out.oos = oos; out.oosRows = oosRows;
    return out;
  },
  workbook: async function (d) {
    var buf = ab((await AdigiCore.workbook(SEL.rows, d.opts)).buffer);
    return { buffer: buf, transfer: [buf] };
  },
  compare: async function (d) { return AdigiCompare.run(d.buf, DATA.rows, ""); },
  fill: async function (d) {
    var res = await AdigiCompare.fillWorkbook(d.buf, SEL.rows, d.opts), buf = ab(res.buffer);
    return { buffer: buf, report: res.report, transfer: [buf] };
  },
  qaReport: async function (d) {
    var buf = ab(await AdigiCompare.workbook(d.res, d.info));
    return { buffer: buf, transfer: [buf] };
  }
};

if (!inWorker) self.AdigiJobs = OPS;
else self.onmessage = async function (e) {
  var d = e.data;
  if (d.op === "init") {
    try { if (d.scripts.length) importScripts.apply(null, d.scripts); self.postMessage({ ready: true }); }   // bundle: already inside
    catch (err) { self.postMessage({ ready: false, error: String(err && err.message || err) }); }
    return;
  }
  try {
    var out = await OPS[d.op](d), transfer = out.transfer || [];
    delete out.transfer;
    self.postMessage({ id: d.id, result: out }, transfer);
  } catch (err) {
    self.postMessage({ id: d.id, error: String(err && err.message || err) });
  }
};
})();
