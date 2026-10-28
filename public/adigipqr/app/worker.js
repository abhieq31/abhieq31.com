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
function rowsIn(d) { var r = d.rows; r.sourceColumns = d.cols; return r; }

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

var OPS = {
  parse: async function (d) {
    step("hash", (d.buf.byteLength / 1048576).toFixed(1) + " MB");
    var sha = await sha256(d.buf);
    step("parse");
    var t = await readTable(d.buf, d.name);
    step("norm", t.rows.length.toLocaleString("en-IN") + " rows");
    var rows = AdigiCore.normalise(t.rows, t.cols);
    return { rows: rows, cols: rows.sourceColumns, sha: sha };
  },
  workbook: async function (d) {
    var buf = ab((await AdigiCore.workbook(rowsIn(d), d.opts)).buffer);
    return { buffer: buf, transfer: [buf] };
  },
  compare: async function (d) { return AdigiCompare.run(d.buf, rowsIn(d), ""); },
  fill: async function (d) {
    var res = await AdigiCompare.fillWorkbook(d.buf, rowsIn(d), d.opts), buf = ab(res.buffer);
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
    try { importScripts.apply(null, d.scripts); self.postMessage({ ready: true }); }
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
