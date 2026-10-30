/* AdigiPQR worker - the heavy work (reading the export, writing .xlsx, QA's workbook) runs here, off the page's
   main thread, so the page never freezes. Same code as the page (core.js); nothing is sent anywhere:
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

/* .xlsx export -> rows, read straight from the sheet XML (JSZip + a scanner): the export is one flat table, so the
   full ExcelJS object model (seconds per 10,000 rows, minutes on a slow PC) is not needed. Same rules as Excel:
   shared strings, inline strings, numbers in shortest form, date-formatted cells as DD-MM-YYYY HH:MM (= the SQL). */
function unxml(t) {
  return t.indexOf("&") < 0 && t.indexOf("_x") < 0 ? t : t
    .replace(/_x([0-9A-Fa-f]{4})_/g, function (m, h) { return String.fromCharCode(parseInt(h, 16)); })
    .replace(/&(lt|gt|amp|quot|apos|#\d+|#x[0-9A-Fa-f]+);/g, function (m, e) {
      return e === "lt" ? "<" : e === "gt" ? ">" : e === "amp" ? "&" : e === "quot" ? '"' : e === "apos" ? "'" :
        String.fromCodePoint(e.charAt(1) === "x" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
    });
}
function texts(xml) {                            // all <t> runs of one string item (rich text), phonetic runs left out
  var out = "", re = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>|<t(?:\s[^>]*)?\/>/g, m;
  xml = xml.replace(/<rPh\b[\s\S]*?<\/rPh>/g, "");
  while ((m = re.exec(xml))) out += m[1] || "";
  return unxml(out);
}
function colIndex(ref) { var n = 0; for (var i = 0; i < ref.length; i++) { var c = ref.charCodeAt(i); if (c < 65) break; n = n * 26 + c - 64; } return n - 1; }
var DATE_IDS = { 14: 1, 15: 1, 16: 1, 17: 1, 18: 1, 19: 1, 20: 1, 21: 1, 22: 1, 45: 1, 46: 1, 47: 1 };
function dateStyles(xml) {                       // style index -> 'd' (date), 'dt' (date + time), 't' (time) for date formats
  var fmts = {}, m, re = /<numFmt\b[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"/g;
  while ((m = re.exec(xml))) fmts[m[1]] = unxml(m[2]);
  var xfs = (xml.match(/<cellXfs\b[\s\S]*?<\/cellXfs>/) || [""])[0], out = [], i = 0, x = /<xf\b([^>]*)/g;
  while ((m = x.exec(xfs))) {
    var id = (m[1].match(/numFmtId="(\d+)"/) || [0, "0"])[1], code = (fmts[id] || "").replace(/"[^"]*"|\[[^\]]*\]|\\./g, "").toLowerCase();
    var isDate = DATE_IDS[id] || /[dy]/.test(code) || (/m/.test(code) && /[dyh]/.test(code)), isTime = /h|s/.test(code) || id === "22" || id === "20" || id === "21";
    out[i++] = isDate ? (isTime ? "dt" : "d") : isTime && /h/.test(code) ? "t" : "";
  }
  return out;
}
function serialDate(v, kind) {                   // Excel serial (1900 system) -> 'DD-MM-YYYY HH:MM'
  var ms = Math.round((v - 25569) * 86400000), d = new Date(ms), p = function (n) { return (n < 10 ? "0" : "") + n; };
  var dd = p(d.getUTCDate()) + "-" + p(d.getUTCMonth() + 1) + "-" + d.getUTCFullYear(), tt = p(d.getUTCHours()) + ":" + p(d.getUTCMinutes());
  return kind === "d" ? dd : kind === "t" ? tt : dd + " " + tt;
}
async function readXlsx(buf) {
  // the export may sit on any sheet and below a title: the header is the first row (of the first 30) with BATCH_NO and
  // STEP_CODE_PATH (an APQR workbook's trace sheet has BATCH_NO too: it is not an export)
  var zip = await JSZip.loadAsync(buf), file = function (n) { var f = zip.file(n); return f ? f.async("string") : Promise.resolve(""); };
  var wbx = await file("xl/workbook.xml"), rels = await file("xl/_rels/workbook.xml.rels"), sheets = [], m;
  var sre = /<sheet\b[^>]*?name="([^"]*)"[^>]*?r:id="([^"]+)"/g;
  while ((m = sre.exec(wbx))) {
    var rel = (rels.match(new RegExp('<Relationship\\b[^>]*Id="' + m[2] + '"[^>]*>')) || [""])[0], tg = (rel.match(/Target="([^"]+)"/) || [])[1];
    if (tg) sheets.push({ name: unxml(m[1]), target: tg.charAt(0) === "/" ? tg.slice(1) : "xl/" + tg.replace(/^\.\//, "") });
  }
  if (!sheets.length) sheets.push({ name: "Sheet1", target: "xl/worksheets/sheet1.xml" });
  var sst = [], sx = await file("xl/sharedStrings.xml"), re = /<si>([\s\S]*?)<\/si>|<si\/>/g;
  while ((m = re.exec(sx))) sst.push(m[1] ? texts(m[1]) : "");
  var styles = dateStyles(await file("xl/styles.xml"));
  for (var si = 0; si < sheets.length; si++) {
    var xml = await file(sheets[si].target);
    if (!xml) continue;
    var rows = sheetRows(xml, sst, styles), h = -1;
    for (var k = 0; k < Math.min(rows.length, 30) && h < 0; k++)
      if (["BATCH_NO", "STEP_CODE_PATH"].every(function (c) { return (rows[k] || []).some(function (x) { return x !== undefined && String(x).trim().toUpperCase() === c; }); })) h = k;
    if (h < 0) continue;
    var head = rows[h].map(function (x) { return x === undefined ? "" : String(x).trim().toUpperCase(); }), out = [];
    for (var i = h + 1; i < rows.length; i++) {
      var o = {}, rw = rows[i] || [], any = false;
      for (var j = 0; j < head.length; j++) if (head[j]) { o[head[j]] = rw[j] === undefined ? "" : rw[j]; if (o[head[j]] !== "") any = true; }
      if (any) out.push(o);
    }
    return { rows: out, cols: head.filter(Boolean) };
  }
  var names = sheets.map(function (x) { return x.name; });
  if (names.some(function (n) { return /^(BMR DETAILS?|APQR_Source_Trace|APQR_Rules|Compression|Coating)$/i.test(n.trim()); }))
    throw new Error("This is an APQR workbook (QA format: " + names.slice(0, 4).join(", ") + " …), not the PAS-X export. Drop the export of final/APQR_FINAL_QUERY.sql here; QA's own workbook goes under “Optional: QA's workbook” below.");
  throw new Error("No sheet of this file has the BATCH_NO and STEP_CODE_PATH columns. Drop the export of final/APQR_FINAL_QUERY.sql (column headers on).");
}

function sheetRows(xml, sst, styles) {
  var rows = [], cur = null, cell = /<row\b[^>]*?(\/)?>|<\/row>|<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, next = 0, m;
  while ((m = cell.exec(xml))) {
    var tag = m[0];
    if (tag.charCodeAt(1) === 114) { cur = []; rows.push(cur); next = 0; continue; }       // <row
    if (tag.charCodeAt(1) === 47) { cur = null; continue; }                                 // </row>
    if (!cur) continue;
    var at = m[2] || "", body = m[3] || "", r = at.match(/\br="([A-Z]+)/), ci = r ? colIndex(r[1]) : next;
    next = ci + 1;
    if (!body) continue;
    var t = (at.match(/\bt="(\w+)"/) || [])[1], v, vm;
    if (t === "inlineStr") v = texts(body);
    else {
      vm = body.match(/<v>([\s\S]*?)<\/v>/); if (!vm) continue;
      v = vm[1];
      if (t === "s") v = sst[+v];
      else if (t === "str" || t === "e") v = unxml(v);
      else if (t !== "b") {
        var st = styles[+((at.match(/\bs="(\d+)"/) || [0, 0])[1])], n = +v;
        v = st && isFinite(n) ? serialDate(n, st) : isFinite(n) ? String(n) : v;
      }
    }
    cur[ci] = v;
  }
  return rows;
}

async function readTable(buf, name) {
  if (/\.xlsx$/i.test(name)) return readXlsx(buf);
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
