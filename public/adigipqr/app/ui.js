/* AdigiPQR web app - page logic. Everything stays in this browser tab. */
(function () {
  "use strict";
  var C = window.AdigiCore, $ = function (id) { return document.getElementById(id); };
  var S = { all: null, file: null, sha: null, rows: null, model: null, params: null, sheet: 0, pf: "", cf: "",
    scans: [], scanFile: null, scanThumbs: {} };
  $("ver").textContent = "v" + C.VERSION;

  function esc(t) { return String(t === null || t === undefined ? "" : t).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function f(v, d) { if (v === null || v === undefined || !isFinite(v)) return "–"; var r = C.pyRound(v, d === undefined ? 3 : d); return String(r); }
  function msg(el, t, kind) { var m = $(el); m.textContent = t || ""; m.className = "msg" + (kind ? " " + kind : ""); }
  function store(k, v) { try { if (v === undefined) return localStorage.getItem(k); localStorage.setItem(k, v); } catch (e) { return null; } }

  /* ------------------------------------------------------------------ tabs */
  function show(tab) {
    document.querySelectorAll("nav [role=tab]").forEach(function (b) { b.setAttribute("aria-selected", b.dataset.tab === tab ? "true" : "false"); });
    ["apqr", "scan", "sql", "guide"].forEach(function (t) { $("tab-" + t).hidden = t !== tab; });
    if (location.hash.slice(1) !== tab) history.replaceState(null, "", "#" + tab);
  }
  document.querySelectorAll("nav [role=tab]").forEach(function (b) { b.onclick = function () { show(b.dataset.tab); }; });
  // The offline zip exists only on the public site; the offline copy and the plant portal hide the link.
  if (!/(^|\.)abhieq31\.com$/.test(location.hostname)) { var z = $("offlineZip"); if (z) z.parentNode.hidden = true; }
  if (window.ADIGI_SQL) $("sqlTabBtn").hidden = false;
  var start = location.hash.slice(1);
  show(["apqr", "scan", "sql", "guide"].indexOf(start) >= 0 && (start !== "sql" || window.ADIGI_SQL) ? start : "apqr");

  /* ------------------------------------------------------------------ file intake */
  function parseStart(t) {
    var m = String(t || "").trim().match(/^(\d{2})[-/](\d{2})[-/](\d{4})(?:\D+(\d{2}):(\d{2}))?/);
    if (m) return m[3] + "-" + m[2] + "-" + m[1] + " " + (m[4] || "00") + ":" + (m[5] || "00");
    m = String(t || "").trim().match(/^(\d{4})-(\d{2})-(\d{2})(?:\D+(\d{2}):(\d{2}))?/);
    return m ? m[1] + "-" + m[2] + "-" + m[3] + " " + (m[4] || "00") + ":" + (m[5] || "00") : null;
  }
  async function sha256(buf) {
    if (!(window.crypto && crypto.subtle)) return null;
    try { var d = await crypto.subtle.digest("SHA-256", buf); return Array.prototype.map.call(new Uint8Array(d), function (x) { return (x < 16 ? "0" : "") + x.toString(16); }).join(""); }
    catch (e) { return null; }
  }
  async function readTable(file) {
    var buf = await file.arrayBuffer();
    S.sha = await sha256(buf);
    if (/\.xlsx$/i.test(file.name)) {
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
    var text = new TextDecoder("utf-8").decode(buf);
    var res = Papa.parse(text, { header: true, skipEmptyLines: true, dynamicTyping: false });
    return { rows: res.data, cols: res.meta.fields };
  }
  async function load(file) {
    msg("msg", "Reading " + file.name + "…");
    try {
      var t = await readTable(file), rows = C.normalise(t.rows, t.cols);
      if (!rows.length) throw new Error("The file has no rows with a BATCH_NO.");
      S.all = rows; S.file = file.name;
      var mats = {};
      rows.forEach(function (r) { var m = C.s(r.MATERIAL_CODE); if (!mats[m]) mats[m] = C.s(r.PRODUCT_NAME); r._start = parseStart(r.MFG_START); });
      var sel = $("fMat"); sel.innerHTML = "";
      var keys = Object.keys(mats).sort();
      if (keys.length > 1) sel.add(new Option("All materials (" + keys.length + ")", ""));
      keys.forEach(function (m) { sel.add(new Option((m || "(none)") + (mats[m] ? " – " + mats[m] : ""), m)); });
      $("fFrom").value = ""; $("fTo").value = "";
      if (S.mode === "batch" && !sel.querySelector('option[value=""]')) { sel.insertBefore(new Option("Any material", ""), sel.firstChild); sel.value = ""; }
      $("filechip").textContent = file.name + " · " + rows.length + " rows" + (S.sha ? " · SHA-256 " + S.sha.slice(0, 12) + "…" : "");
      msg("msg", ""); $("loaded").hidden = false;
      document.body.classList.add("has-data"); loadCpp(); recompute();
    } catch (e) { msg("msg", "Not loaded: " + e.message, "err"); }
  }
  $("file").addEventListener("change", function (e) { if (e.target.files[0]) load(e.target.files[0]); e.target.value = ""; });
  // invented example files shipped with the app (web/example): a download from this site, nothing is uploaded
  async function example(name, type) {
    var r = await fetch("example/" + name);
    if (!r.ok) throw new Error("HTTP " + r.status);
    return new File([await r.blob()], name, { type: type });
  }
  $("exLoad").onclick = async function () {
    try { await load(await example("example_pasx_export.csv", "text/csv")); S.example = true; }
    catch (e) { msg("msg", "The example file could not be opened here (" + e.message + "). It is in the folder example/ next to this page.", "err"); }
  };
  $("exTpl").onclick = async function () {
    try { await loadTpl(await example("example_qa_workbook_template.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")); }
    catch (e) { msg("tmsg", "The example file could not be opened here (" + e.message + ").", "err"); }
  };
  $("exQa").onclick = async function () {
    try { await loadQa(await example("example_qa_manual.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")); }
    catch (e) { msg("qmsg", "The example file could not be opened here (" + e.message + ").", "err"); }
  };
  dropZone($("drop"), load);
  function dropZone(el, fn) {
    ["dragenter", "dragover"].forEach(function (ev) { el.addEventListener(ev, function (e) { e.preventDefault(); el.classList.add("over"); }); });
    ["dragleave", "drop"].forEach(function (ev) { el.addEventListener(ev, function (e) { e.preventDefault(); el.classList.remove("over"); }); });
    el.addEventListener("drop", function (e) { if (e.dataTransfer.files[0]) fn(e.dataTransfer.files[0]); });
  }

  /* ------------------------------------------------------------------ selection */
  function scanRowsFor(batchSet, mat, lo, hi) {
    var out = [];
    S.scans.filter(function (r) { return r.status === "accepted"; }).forEach(function (r) {
      var row = r.row, st = parseStart(row.FIRST_READING_ON);
      var mine = batchSet[row.BATCH_NO] || ((!mat || row.MATERIAL_CODE === mat) && (!lo || (st && st >= lo)) && (!hi || (st && st <= hi)));
      if (mine) out.push(row);
    });
    return out;
  }
  // QA asks in two ways: all batches of a material in a period, or a list of batch numbers (material optional)
  function batchList() { return uniqList($("fBatches").value.split(/[\s,;]+/).filter(Boolean)); }
  function uniqList(a) { var o = []; a.forEach(function (x) { if (o.indexOf(x) < 0) o.push(x); }); return o; }
  function selected() {
    var byBatch = S.mode === "batch", want = byBatch ? batchList() : [], wantSet = {};
    want.forEach(function (b) { wantSet[b] = 1; });
    var mat = $("fMat").value, from = byBatch ? "" : $("fFrom").value, to = byBatch ? "" : $("fTo").value, lo = from ? from + " 00:00" : null, hi = to ? to + " 23:59" : null;
    var undated = {}, rows = (S.all || []).filter(function (r) {
      if (byBatch && !wantSet[r.BATCH_NO]) return false;
      if (mat && C.s(r.MATERIAL_CODE) !== mat) return false;
      if (!lo && !hi) return true;
      if (!r._start) { undated[r.BATCH_NO] = 1; return false; }
      return (!lo || r._start >= lo) && (!hi || r._start <= hi);
    });
    var first = {};
    rows.forEach(function (r) { if (r._start && (!first[r.BATCH_NO] || r._start < first[r.BATCH_NO])) first[r.BATCH_NO] = r._start; });
    // manufacturing order: batch start, then batch number (same as the portal)
    var order = Object.keys(rows.reduce(function (a, r) { a[r.BATCH_NO] = 1; return a; }, {}))
      .sort(function (a, b) { var x = first[a] || "9999", y = first[b] || "9999"; return x < y ? -1 : x > y ? 1 : a < b ? -1 : a > b ? 1 : 0; });
    var pos = {}; order.forEach(function (b, i) { pos[b] = i; });
    rows = rows.slice().sort(function (a, b) { return pos[a.BATCH_NO] - pos[b.BATCH_NO]; });
    var bs = {}; order.forEach(function (b) { bs[b] = 1; });
    var scanRows = byBatch ? S.scans.filter(function (r) { return r.status === "accepted" && wantSet[r.row.BATCH_NO]; }).map(function (r) { return r.row; })
      : scanRowsFor(bs, mat, lo, hi);
    var missing = want.filter(function (b) { return !bs[b]; });
    var all = rows.concat(scanRows.map(function (r) { var o = Object.assign({}, r); o._scan = true; return o; }));
    all.sourceColumns = S.all.sourceColumns;
    return { rows: all, mat: mat, from: from, to: to, undated: Object.keys(undated), first: first, nScan: scanRows.length,
      byBatch: byBatch, want: want, missing: missing };
  }
  function opt() { return { merge: $("oMerge").checked, minutes: $("oMin").checked, inspection: $("oInsp").checked }; }
  // QA's CPP list for the product: typed once per material, remembered in this browser (product knowledge, no batch data)
  function cppKey() { return "apqr_cpp_" + ($("fMat").value || "all"); }
  function cppText() { return $("oCpp").value; }
  function loadCpp() { $("oCpp").value = store(cppKey()) || ""; }

  function recompute() {
    if (!S.all) return;
    var sel = selected();
    S.sel = sel; S.rows = sel.rows;
    if (!sel.rows.length) {
      S.model = null; $("kpis").innerHTML = "";
      msg("msg", sel.byBatch ? (sel.want.length ? "None of the requested batch numbers are in this export" + (sel.mat ? " under material " + sel.mat : "") + ": " + sel.want.slice(0, 10).join(", ") + (sel.want.length > 10 ? " …" : "") :
        "Enter the batch numbers QA asked for, then Apply.") : "No batches of this material in this period.", "err");
      ["secBatches", "secParams", "secCharts", "secSheets"].forEach(function (id) { $(id).hidden = true; });
      $("dl").disabled = true; $("tdl").disabled = true; return;
    }
    ["secBatches", "secParams", "secCharts", "secSheets"].forEach(function (id) { $(id).hidden = false; });
    $("dl").disabled = false; $("tdl").disabled = !S.tpl; msg("msg", "");
    S.model = C.build(sel.rows, opt()); S.params = C.parameters(S.model, cppText());
    var starts = S.model.batches.map(function (b) { return sel.first[b]; }).filter(Boolean).sort();
    $("period").textContent = (sel.byBatch ? "Requested " + sel.want.length + " batch number(s): " + (sel.want.length - sel.missing.length) + " found" +
        (sel.missing.length ? ", not in this export" + (sel.mat ? " under material " + sel.mat : "") + ": " + sel.missing.join(", ") : "") + ". " : "") +
      (starts.length ? "Manufactured " + starts[0].slice(0, 10) + " → " + starts[starts.length - 1].slice(0, 10) + ". " : "") +
      (sel.undated.length ? sel.undated.length + " batch(es) without a start date were left out by the date filter. " : "") +
      (sel.nScan ? sel.nScan + " record(s) from BMR scans included." : "");
    renderKpis(); renderBatches(); renderParams();
    if ($("secCharts").open) renderCharts(); else $("charts").innerHTML = "";
    if ($("secSheets").open) renderSheet();
  }
  ["apply"].forEach(function (id) { $(id).onclick = recompute; });
  S.mode = "period";
  function setMode(m) {
    S.mode = m;
    $("reqMode").querySelectorAll("button").forEach(function (b) { b.setAttribute("aria-pressed", b.dataset.m === m ? "true" : "false"); });
    document.querySelectorAll(".m-period").forEach(function (el) { el.hidden = m !== "period"; });
    document.querySelectorAll(".m-batch").forEach(function (el) { el.hidden = m !== "batch"; });
    var any = $("fMat").querySelector('option[value=""]');
    if (m === "batch" && !any && $("fMat").options.length) { $("fMat").insertBefore(new Option("Any material", ""), $("fMat").firstChild); $("fMat").value = ""; }
    if (m === "batch") $("fBatches").focus();
    recompute();
  }
  $("reqMode").querySelectorAll("button").forEach(function (b) { b.onclick = function () { setMode(b.dataset.m); }; });
  $("fBatches").addEventListener("change", recompute);
  $("fMat").addEventListener("change", loadCpp);
  ["fMat", "oMerge", "oMin", "oInsp"].forEach(function (id) { $(id).addEventListener("change", recompute); });
  $("oCpp").addEventListener("change", function () { store(cppKey(), this.value); recompute(); });
  $("secCharts").addEventListener("toggle", function () { if (this.open) renderCharts(); });
  $("secSheets").addEventListener("toggle", function () { if (this.open) renderSheet(); });
  $("who").value = store("apqr_by") || "";
  $("who").addEventListener("change", function () { store("apqr_by", this.value); $("sby").value = this.value; });

  /* ------------------------------------------------------------------ overview */
  function renderKpis() {
    var p = S.params, oosRows = S.model.sheets[S.model.sheets.length - 1].rows.length;
    var weak = p.filter(function (x) { return x.cap.cpk !== null && x.cap.n >= 6 && x.cap.cpk < 1.33; }).length;
    var sig = p.filter(function (x) { return x.nelson.length; }).length;
    var cpps = p.filter(function (x) { return x.crit === C.CPP || x.crit === C.CPP_TYPICAL; });
    var sure = cpps.filter(function (x) { return x.crit === C.CPP; }).length, cqa = p.filter(function (x) { return x.crit === C.CQA; }).length;
    var tiles = [[S.model.batches.length, "batches"], [p.length, "parameters trended"],
      [cpps.length, "CPPs identified · " + sure + " confirmed"], [cqa, "in-process CQAs (IPQC)"],
      [oosRows, "values outside MBR limit", oosRows ? "warn" : ""], [weak, "parameters with Cpk < 1.33", weak ? "amber" : ""],
      [sig, "parameters with trend signals", sig ? "amber" : ""], [S.model.sheets.filter(function (s) { return s.kind === "grid"; }).length, "QA stage sheets"]];
    $("kpis").innerHTML = tiles.map(function (t) { return '<div class="kpi ' + (t[2] || "") + '"><b>' + t[0] + "</b><span>" + esc(t[1]) + "</span></div>"; }).join("");
  }
  function renderBatches() {
    var hdr = {}; S.rows.forEach(function (r) { if (!r._scan && !hdr[r.BATCH_NO]) hdr[r.BATCH_NO] = r; });
    var oos = {}; S.rows.forEach(function (r) { if (/^OUT/i.test(C.s(r.LIMIT_CHECK))) oos[r.BATCH_NO] = (oos[r.BATCH_NO] || 0) + 1; });
    var h = '<tr><th>#</th><th>Batch</th><th>Material</th><th>Started</th><th>Size</th><th>MBR</th><th>Order</th><th>Status</th><th class="n">Out of limit</th></tr>';
    $("tBatches").innerHTML = h + S.model.batches.map(function (b, i) {
      var r = hdr[b] || {};
      return '<tr class="' + (oos[b] ? "bad" : "") + '"><td>' + (i + 1) + '</td><td class="mono">' + esc(b) + '</td><td class="mono">' + esc(r.MATERIAL_CODE) +
        '</td><td class="mono">' + esc((S.sel.first[b] || "").slice(0, 16)) + "</td><td>" + esc((C.s(r.BATCH_SIZE) + " " + C.s(r.BATCH_SIZE_UOM)).trim()) +
        '</td><td class="mono">' + esc(C.s(r.MBR_ID) + (C.s(r.MBR_VERSION) ? "-" + C.s(r.MBR_VERSION) : "")) + '</td><td class="mono">' + esc(r.MANUFACTURING_ORDER) +
        "</td><td>" + esc(r.MO_STATUS) + '</td><td class="n">' + (oos[b] || "") + "</td></tr>";
    }).join("");
    $("secBatches").querySelector("summary span").textContent = "Batches (" + S.model.batches.length + ")";
  }
  function pFilter(p, mode) {
    if (mode === "crit") return !!p.crit;
    if (mode === "out") return p.out.length > 0;
    if (mode === "cpk") return p.cap.cpk !== null && p.cap.n >= 6 && p.cap.cpk < 1.33;
    if (mode === "sig" || mode === "flag") return p.nelson.length > 0 || p.out.length > 0 || (mode === "flag" && p.cap.cpk !== null && p.cap.n >= 6 && p.cap.cpk < 1.33);
    return true;
  }
  function rulePills(p) {                       // one pill per rule, with the number of runs
    var c = {}; p.nelson.forEach(function (x) { c[x.rule] = (c[x.rule] || 0) + 1; });
    return Object.keys(c).map(function (r) {
      return '<span class="pill amb" title="' + esc(C.NELSON[r]) + '">R' + r + (c[r] > 1 ? " ×" + c[r] : "") + "</span>";
    }).join(" ");
  }
  function critPill(p) {
    if (!p.crit) return "";
    var t = p.crit === C.CPP ? ["ok", "CPP"] : p.crit === C.CPP_TYPICAL ? ["typ", "CPP?"] : ["cqa", "CQA"];
    return '<span class="pill ' + t[0] + '" title="' + esc(p.crit + " – " + p.crit_basis) + '">' + t[1] + "</span>";
  }
  function verdictPill(p) {
    var n = C.note(p), cls = n === "capable" ? "ok" : n === "not capable" ? "bad" : n === "marginal" ? "amb" : "";
    return '<span class="pill ' + cls + '">' + esc(n) + "</span>";
  }
  function renderParams() {
    var list = S.params.filter(function (p) { return pFilter(p, S.pf); });
    var h = '<tr><th>Sheet</th><th>Stage</th><th>Parameter</th><th>Criticality</th><th class="n">n</th><th class="n">LSL</th><th class="n">USL</th><th class="n">Mean</th><th class="n">SD</th>' +
      '<th class="n">Cpk</th><th class="n">Ppk</th><th>Capability</th><th>Outside limit</th><th>Signals</th></tr>';
    $("tParams").innerHTML = h + list.map(function (p) {
      return '<tr class="' + (p.out.length ? "bad" : p.nelson.length ? "amb" : "") + '"><td>' + esc(p.sheet) + "</td><td>" + esc(p.stage) + "<br><span class=note>" + esc(p.section) +
        "</span></td><td>" + esc(p.param) + (p.uom ? ' <span class="note">[' + esc(p.uom) + "]</span>" : "") + "</td><td>" + critPill(p) +
        (p.crit ? '<br><span class="note">' + esc(p.crit_basis) + "</span>" : "") + '</td><td class="n">' + p.cap.n +
        '</td><td class="n">' + f(p.lo) + '</td><td class="n">' + f(p.hi) + '</td><td class="n">' + f(p.cap.mean) + '</td><td class="n">' + f(p.cap.sd) +
        '</td><td class="n">' + f(p.cap.cpk, 2) + '</td><td class="n">' + f(p.cap.ppk, 2) + "</td><td>" + verdictPill(p) + '</td><td class="mono">' + esc(p.out.join(", ")) +
        "</td><td>" + rulePills(p) + "</td></tr>";
    }).join("") + (list.length ? "" : '<tr><td colspan="14" class="note">No parameter matches this filter.</td></tr>');
    $("secParams").querySelector("summary span").textContent = "Parameters, capability and trend signals (" + S.params.length + ")";
  }
  chips("pFilter", function (v) { S.pf = v; renderParams(); });
  chips("cFilter", function (v) { S.cf = v; renderCharts(); });
  function chips(id, fn) {
    $(id).querySelectorAll("button").forEach(function (b) {
      b.onclick = function () { $(id).querySelectorAll("button").forEach(function (x) { x.setAttribute("aria-pressed", x === b ? "true" : "false"); }); fn(b.dataset.f); };
    });
  }

  /* ------------------------------------------------------------------ charts (SVG, individuals chart with limits) */
  function chartSvg(p) {
    var W = 420, H = 190, L = 44, R = 10, T = 10, B = 34, x = p.values, n = x.length;
    var ch = p.chart || {}, lines = [p.lo, p.hi, ch.ucl, ch.lcl].filter(function (v) { return v !== null && v !== undefined && isFinite(v); });
    var lo = Math.min.apply(null, x.concat(lines)), hi = Math.max.apply(null, x.concat(lines)), span = (hi - lo) || Math.abs(hi) * 0.1 || 1;
    lo -= span * 0.08; hi += span * 0.08;
    var X = function (i) { return L + (n === 1 ? (W - L - R) / 2 : i * (W - L - R) / (n - 1)); }, Y = function (v) { return T + (hi - v) * (H - T - B) / (hi - lo); };
    var flagged = {}; p.nelson.forEach(function (fl) { fl.points.forEach(function (i) { flagged[i] = 1; }); });
    var out = {}; p.out.forEach(function (b) { out[b] = 1; });
    var g = '<svg viewBox="0 0 ' + W + " " + H + '" role="img" aria-label="' + esc(p.param) + ' trend">';
    [0, 0.5, 1].forEach(function (t) { var v = lo + t * (hi - lo); g += '<text x="' + (L - 6) + '" y="' + (Y(v) + 4) + '" font-size="10" text-anchor="end" fill="currentColor" opacity=".6">' + f(v, 2) + "</text>"; });
    function hl(v, color, dash, label, left) {
      if (v === null || v === undefined || !isFinite(v)) return "";
      return '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + Y(v) + '" y2="' + Y(v) + '" style="stroke:' + color + '" stroke-width="1.2"' + (dash ? ' stroke-dasharray="' + dash + '"' : "") +
        '/><text x="' + (left ? L + 3 : W - R) + '" y="' + (Y(v) - 3) + '" font-size="9.5" text-anchor="' + (left ? "start" : "end") + '" style="fill:' + color + '">' + label + "</text>";
    }
    g += hl(p.hi, "var(--limit)", "", "USL " + f(p.hi)) + hl(p.lo, "var(--limit)", "", "LSL " + f(p.lo));
    if (ch.ucl) g += hl(ch.ucl, "var(--ctl)", "4 3", "UCL", true) + hl(ch.lcl, "var(--ctl)", "4 3", "LCL", true);
    if (ch.cl !== undefined) g += hl(ch.cl, "var(--ctl)", "1 3", "");
    g += '<polyline fill="none" style="stroke:var(--series)" stroke-width="1.6" points="' + x.map(function (v, i) { return X(i) + "," + Y(v); }).join(" ") + '"/>';
    x.forEach(function (v, i) {
      var b = p.batches[i], col = out[b] ? "var(--limit)" : flagged[i] ? "var(--flag)" : "var(--series)", r = out[b] || flagged[i] ? 4.2 : 2.8;
      g += '<circle cx="' + X(i) + '" cy="' + Y(v) + '" r="' + r + '" style="fill:' + col + '"><title>' + esc(b) + ": " + f(v, 4) + "</title></circle>";
    });
    var step = Math.max(1, Math.ceil(n / 6));
    for (var i = 0; i < n; i += step) g += '<text x="' + X(i) + '" y="' + (H - 14) + '" font-size="9" text-anchor="middle" fill="currentColor" opacity=".6">' + esc(p.batches[i].slice(-5)) + "</text>";
    return g + "</svg>";
  }
  function renderCharts() {
    if (!S.params) return;
    var list = S.params.filter(function (p) { return p.values.length >= 2 && pFilter(p, S.cf); });
    $("charts").innerHTML = list.map(function (p) {
      var c = p.cap;
      return '<div class="chart"><h3>' + critPill(p) + " " + esc(p.param) + (p.uom ? " [" + esc(p.uom) + "]" : "") + '</h3><div class="note">' + esc(p.sheet + " · " + p.stage) + "</div>" + chartSvg(p) +
        "<p>n " + c.n + " · mean " + f(c.mean) + " · Cpk " + f(c.cpk, 2) + " · Ppk " + f(c.ppk, 2) + " · " + verdictPill(p) +
        (p.out.length ? ' · <span class="pill bad">' + p.out.length + " outside limit</span>" : "") +
        " " + rulePills(p) + "</p></div>";
    }).join("") || '<p class="note">No charts for this filter.</p>';
    $("secCharts").querySelector("summary span").textContent = "Trend charts (" + S.params.filter(function (p) { return p.values.length >= 2; }).length + ")";
  }

  /* ------------------------------------------------------------------ sheet preview */
  function renderSheet() {
    var sheets = S.model.sheets;
    if (S.sheet >= sheets.length) S.sheet = 0;
    $("sheetTabs").innerHTML = sheets.map(function (s, i) { return '<button aria-selected="' + (i === S.sheet) + '" data-i="' + i + '">' + esc(s.name) + "</button>"; }).join("");
    $("sheetTabs").querySelectorAll("button").forEach(function (b) { b.onclick = function () { S.sheet = +b.dataset.i; renderSheet(); }; });
    var s = sheets[S.sheet], h = "";
    if (s.kind === "simple") {
      h = "<tr>" + s.cols.map(function (c) { return "<th>" + esc(c) + "</th>"; }).join("") + "</tr>" +
        s.rows.slice(0, 400).map(function (r) { return "<tr>" + r.map(function (v) { return "<td>" + esc(v) + "</td>"; }).join("") + "</tr>"; }).join("");
    } else {
      for (var lv = 0; lv < 5; lv++) {
        h += "<tr>" + (lv === 0 ? '<th rowspan="5">Batch</th>' : "");
        var i = 0;
        while (i < s.cols.length) {
          var j = i + 1;
          while (j < s.cols.length && lv < 4 && s.cols[j].r.slice(0, lv + 1).join("\u0001") === s.cols[i].r.slice(0, lv + 1).join("\u0001") && s.cols[j].text === s.cols[i].text) j++;
          h += '<th colspan="' + (j - i) + '">' + esc(s.cols[i].r[lv]) + "</th>"; i = j;
        }
        h += "</tr>";
      }
      h += s.batches.map(function (b) {
        return '<tr><td class="mono">' + esc(b) + "</td>" + s.cols.map(function (c) {
          var v = c.values[b];
          return c.text ? "<td>" + esc(v) + "</td>" : '<td class="n' + (c.oos[b] && v !== undefined ? " oos" : "") + '">' + (v === undefined || v === null ? "" : f(v, 4)) + "</td>";
        }).join("") + "</tr>";
      }).join("");
    }
    $("tSheet").innerHTML = h;
  }

  /* ------------------------------------------------------------------ check against QA's manual workbook */
  async function loadQa(file) {
    if (!S.all) { msg("qmsg", "Load the PAS-X export first.", "err"); return; }
    msg("qmsg", "Comparing…"); $("qres").hidden = true;
    try {
      var res = await window.AdigiCompare.run(await file.arrayBuffer(), S.all, "");
      S.qa = res; S.qaFile = file.name; S.qf = "";
      $("qchip").textContent = file.name + " · " + res.qaValues + " values in " + res.qaBatches.length + " batch(es)";
      if (!res.results.length) {
        msg("qmsg", res.qaBatches.length ? "None of QA's batches are in this export (" + res.missing.slice(0, 5).join(", ") + (res.missing.length > 5 ? " …" : "") + "). Export the same batches from PAS-X." :
          "No batch rows found. Expected sheets: " + window.AdigiCompare.SHEETS.join(", ") + ", with 10-digit batch numbers in one of the first four columns.", "err");
        return;
      }
      msg("qmsg", res.missing.length ? res.missing.length + " of QA's batches are not in this export and were skipped: " + res.missing.slice(0, 8).join(", ") + (res.missing.length > 8 ? " …" : "") : "", res.missing.length ? "err" : "");
      renderQa(); $("qres").hidden = false;
    } catch (e) { msg("qmsg", "Could not read the workbook: " + e.message, "err"); }
  }
  function renderQa() {
    var r = S.qa.results, count = function (f) { return r.filter(f).length; };
    var typed = r.filter(function (x) { return x.STATUS !== "QA_BLANK"; });
    var match = count(function (x) { return /^MATCH/.test(x.STATUS); }), mis = count(function (x) { return x.STATUS === "MISMATCH"; });
    var comparable = typed.filter(function (x) { return /^MATCH|MISMATCH|AMBIGUOUS/.test(x.STATUS); }).length;
    var tiles = [[typed.length, "values typed by QA"], [match, "match PAS-X" + (comparable ? " (" + Math.round(100 * match / comparable) + " % of comparable)" : "")],
      [mis, "mismatch", mis ? "warn" : ""], [count(function (x) { return x.STATUS === "AMBIGUOUS"; }), "ambiguous", "amber"],
      [count(function (x) { return x.STATUS === "NOT_IN_MES"; }), "not recorded in PAS-X"], [count(function (x) { return x.STATUS === "NO_MAPPING_RULE"; }), "columns without a mapping rule"]];
    $("qkpis").innerHTML = tiles.map(function (t) { return '<div class="kpi ' + (t[2] && t[0] ? t[2] : "") + '"><b>' + t[0] + "</b><span>" + esc(t[1]) + "</span></div>"; }).join("");
    var sts = [], sheets = [];
    r.forEach(function (x) { if (sts.indexOf(x.STATUS) < 0) sts.push(x.STATUS); if (sheets.indexOf(x.QA_SHEET) < 0) sheets.push(x.QA_SHEET); });
    sts.sort();
    $("qsum").innerHTML = "<tr><th>QA sheet</th>" + sts.map(function (st) { return '<th class="n">' + esc(st) + "</th>"; }).join("") + "</tr>" +
      sheets.map(function (sh) {
        return "<tr><td>" + esc(sh) + "</td>" + sts.map(function (st) { var n = count(function (x) { return x.QA_SHEET === sh && x.STATUS === st; }); return '<td class="n">' + (n || "") + "</td>"; }).join("") + "</tr>";
      }).join("");
    renderQaDetail();
  }
  function renderQaDetail() {
    var list = S.qa.results.filter(function (x) { return S.qf ? x.STATUS === S.qf : x.STATUS !== "QA_BLANK"; }), more = list.length > 500;
    $("qdet").innerHTML = '<tr><th>Batch</th><th>QA sheet › column</th><th>Stat</th><th class="n">QA value</th><th class="n">PAS-X value</th><th>Status</th><th>PAS-X stage › step › parameter</th></tr>' +
      list.slice(0, 500).map(function (x) {
        var cls = x.STATUS === "MISMATCH" ? "bad" : x.STATUS === "AMBIGUOUS" ? "amb" : "";
        var qv = x.QA_VALUE instanceof Date ? x.QA_VALUE.toISOString().slice(11, 19) : x.QA_VALUE;
        return '<tr class="' + cls + '"><td class="mono">' + esc(x.BATCH_NO) + "</td><td>" + esc(x.QA_SHEET + " › " + x.QA_COLUMN) + "</td><td>" + esc(x.STAT) +
          '</td><td class="n">' + esc(qv === null || qv === undefined ? "" : qv) + '</td><td class="n">' + esc(typeof x.MES_VALUE === "number" ? f(x.MES_VALUE, 4) : x.MES_VALUE || "") +
          "</td><td>" + esc(x.STATUS) + "</td><td>" + esc([x.MES_STAGE, x.MES_STEP, x.MES_PARAMETER].filter(Boolean).join(" › ")) + "</td></tr>";
      }).join("") + (more ? '<tr><td colspan="7" class="note">First 500 of ' + list.length + " shown – the download has all.</td></tr>" : "") +
      (list.length ? "" : '<tr><td colspan="7" class="note">Nothing in this group.</td></tr>');
  }
  $("qfile").addEventListener("change", function (e) { if (e.target.files[0]) loadQa(e.target.files[0]); e.target.value = ""; });
  dropZone($("qdrop"), loadQa);
  chips("qFilter", function (v) { S.qf = v; renderQaDetail(); });
  $("qdl").onclick = async function () {
    if (!S.qa) return;
    var buf = await window.AdigiCompare.workbook(S.qa, [["QA workbook", S.qaFile], ["PAS-X export", S.file + (S.sha ? " · SHA-256 " + S.sha : "")],
      ["Generated", new Date().toISOString()], ["Tool", "AdigiPQR web app v" + C.VERSION + " (processed in the browser, nothing uploaded)"]]);
    save(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), "APQR_check_vs_QA_" + stamp() + ".xlsx");
  };

  /* ------------------------------------------------------------------ QA's official workbook, filled */
  async function loadTpl(file) {
    S.tpl = { name: file.name, buffer: await file.arrayBuffer() };
    $("tchip").textContent = file.name + " · " + Math.round(file.size / 1024) + " KB";
    $("tdl").disabled = !S.rows || !S.rows.length; $("tres").hidden = true; msg("tmsg", "");
  }
  $("tfile").addEventListener("change", function (e) { if (e.target.files[0]) loadTpl(e.target.files[0]); e.target.value = ""; });
  dropZone($("tdrop"), loadTpl);
  $("tdl").onclick = async function () {
    if (!S.tpl || !S.rows || !S.rows.length) return;
    var btn = this; btn.disabled = true; msg("tmsg", "Filling QA's workbook… (a workbook with thousands of batches takes about 15 seconds)");
    try {
      var sel = S.sel, req = sel.byBatch ? "batch numbers: " + sel.want.join(", ") : "material " + (sel.mat || "all") + ", manufactured " + (sel.from || "any") + " to " + (sel.to || "any");
      var info = [["QA workbook", S.tpl.name], ["PAS-X export", S.file + (S.sha ? " · SHA-256 " + S.sha : "")], ["QA request", req],
        ["Generated", new Date().toISOString()], ["Tool", "AdigiPQR web app v" + C.VERSION + " (processed in the browser, nothing uploaded)"],
        ["Status", "Prepared from the PAS-X MES execution record for QA review. Verify against the batch record before use in the APQR."]];
      var res = await window.AdigiCompare.fillWorkbook(S.tpl.buffer.slice(0), S.rows, { leaveYellow: $("tYellow").checked, info: info });
      save(new Blob([res.buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
        S.tpl.name.replace(/\.xlsx$/i, "") + "_filled_" + stamp() + ".xlsx");
      var cnt = {}; res.report.forEach(function (r) { cnt[r.STATUS] = (cnt[r.STATUS] || 0) + 1; });
      var T = window.AdigiCompare.STATUS_TEXT, notes = res.report.filter(function (r) { return !r.CELL; });
      $("tsum").innerHTML = "<tr><th>Result</th><th class=\"n\">Cells</th><th>Meaning</th></tr>" + Object.keys(T).filter(function (k) { return cnt[k]; }).map(function (k) {
        return '<tr class="' + (k === "AMBIGUOUS" ? "amb" : "") + '"><td>' + esc(k) + '</td><td class="n">' + cnt[k] + "</td><td>" + esc(T[k]) + "</td></tr>";
      }).join("") + notes.map(function (r) { return '<tr><td>' + esc(r.SHEET || "–") + "</td><td></td><td>" + esc(r.STATUS + ": " + r.SOURCE) + "</td></tr>"; }).join("");
      $("tres").hidden = false;
      msg("tmsg", (cnt.FILLED || 0) + " values written into QA's workbook. Details are in its last sheet, 'AdigiPQR fill report'.", "ok");
    } catch (e) { msg("tmsg", "Could not fill the workbook: " + e.message, "err"); }
    btn.disabled = false;
  };

  /* ------------------------------------------------------------------ download */
  function stamp() { var d = new Date(), p = function (n) { return (n < 10 ? "0" : "") + n; }; return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()); }
  function save(blob, name) {
    var url = URL.createObjectURL(blob), a = document.createElement("a");
    a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 15000);
  }
  async function buildAndSave(rows, nameBase, sourceDesc) {
    var o = opt(), oos = rows.filter(function (r) { return /^OUT/i.test(C.s(r.LIMIT_CHECK)); }).length;
    var accepted = S.scans.filter(function (r) { return r.status === "accepted"; });
    var batches = {}; rows.forEach(function (r) { batches[r.BATCH_NO] = 1; });
    var scans = accepted.filter(function (r) { return batches[r.row.BATCH_NO]; });
    var info = [["Tool", "AdigiPQR web app v" + C.VERSION + " (processed in the browser, nothing uploaded)"], ["Generated", new Date().toISOString()],
      ["Requested by", $("who").value || "-"]].concat(sourceDesc).concat([
      ["Batches", Object.keys(batches).length], ["Rows", rows.length], ["Out-of-limit rows", oos],
      ["Scanned records", scans.length ? scans.length + " accepted (sheet 'Scanned records')" : "none"],
      ["Options", "merge IPC=" + o.merge + "; minutes=" + o.minutes + "; inspection=" + o.inspection],
      ["Status", "Prepared from the PAS-X MES execution record for QA review. Verify against the batch record before use in the APQR."]]);
    var list = C.parseCppList(cppText());
    if (list.length) info.splice(info.length - 1, 0, ["QA CPP list", list.map(function (e) { return e[0]; }).join("; ")]);
    var res = await C.workbook(rows, { pivot: o, info: info, scans: scans, cppList: cppText() });
    save(new Blob([res.buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), nameBase + "_" + stamp() + ".xlsx");
  }
  $("dl").onclick = async function () {
    var btn = this; btn.disabled = true; msg("msg", "Building the workbook…");
    try {
      var sel = S.sel;
      var desc = [["Source file", S.file], ["Source SHA-256", S.sha || "not available in this browser context"], ["Material", sel.mat || "all in the file"]];
      if (sel.byBatch) desc = desc.concat([["QA request", "batch numbers (" + sel.want.length + ")"], ["Batch numbers requested", sel.want.join(", ")],
        ["Requested but not in the export", sel.missing.join(", ") || "none"]]);
      else desc = desc.concat([["QA request", "material + period"], ["Manufactured from", sel.from || "any"], ["Manufactured to", sel.to || "any"]]);
      await buildAndSave(S.rows, "APQR_" + (sel.mat || (sel.byBatch ? "batches" : "all")) + (sel.byBatch ? "_" + sel.want.length + "batches" :
        (sel.from || sel.to ? "_" + (sel.from || "start") + "_" + (sel.to || "end") : "")), desc);
      msg("msg", "Workbook saved.", "ok");
    } catch (e) { msg("msg", "Could not build the workbook: " + e.message, "err"); }
    btn.disabled = false;
  };

  /* ------------------------------------------------------------------ scans */
  var sfile = null;
  function setScanFile(fl) { sfile = fl; $("schip").textContent = fl.name + " · " + Math.round(fl.size / 1024) + " KB"; $("sread").disabled = false; }
  $("sfile").addEventListener("change", function (e) { if (e.target.files[0]) setScanFile(e.target.files[0]); e.target.value = ""; });
  dropZone($("sdrop"), setScanFile);
  $("sby").value = store("apqr_by") || "";
  function parsePages(t) {
    var set = new Set();
    String(t || "").split(/[,\s]+/).filter(Boolean).forEach(function (part) {
      var m = part.match(/^(\d+)(?:-(\d+))?$/);
      if (!m) throw new Error("Page list not understood: " + part + " (use e.g. 347-357, 360)");
      var a = +m[1], b = +(m[2] || m[1]);
      for (var i = Math.min(a, b); i <= Math.max(a, b); i++) set.add(i);
    });
    return set.size ? set : null;
  }
  $("sread").onclick = async function () {
    var btn = this; btn.disabled = true;
    try {
      var pages = parsePages($("spages").value), t0 = performance.now();
      var res = await AdigiOCR.readDocument(sfile, pages, function (p) {
        msg("smsg", p.stage || "");
        if (p.pages) $("sbar").style.width = Math.round(100 * p.page / p.pages) + "%";
      });
      $("sbar").style.width = "100%";
      var sha = await sha256(await sfile.arrayBuffer());
      res.records.forEach(function (r) { r.status = "pending"; r.file = sfile.name; r.fileSha = sha; S.scans.push(r); });
      Object.keys(res.thumbs).forEach(function (n) { S.scanThumbs[sfile.name + "#" + n] = res.thumbs[n]; });
      var found = res.log.filter(function (x) { return x[1]; }).length;
      msg("smsg", "Read " + res.log.length + " page(s) in " + Math.round((performance.now() - t0) / 1000) + " s: " + found + " printout page(s), " +
        res.records.length + " record(s), " + res.records.filter(function (r) { return r.ok; }).length + " passed every check.", "ok");
      renderScans();
    } catch (e) { msg("smsg", e.message, "err"); }
    btn.disabled = false;
  };
  function renderScans() {
    $("sresult").hidden = !S.scans.length;
    var keys = ["n", "avg", "min", "max", "sd", "lsl", "usl"];
    var h = "<tr><th>Batch</th><th>Serial</th><th>Interval · side</th><th>Balance</th><th>n</th><th>Average</th><th>Min</th><th>Max</th><th>SD</th><th>Limit</th><th>Result</th><th>Checks</th><th>Status</th></tr>";
    $("srecs").innerHTML = h + S.scans.map(function (r, i) {
      var v = r.values, failed = r.checks.filter(function (c) { return !c.ok; });
      var row = '<tr class="' + (r.ok ? "" : "amb") + '"><td class="mono">' + esc(v.batch) + "</td><td>" + esc(v.serial || "?") + "</td><td>" + esc(v.interval + " · " + v.side) +
        '</td><td class="mono">' + esc(v.balance) + "</td><td>" + f(v.n, 0) + "</td><td>" + f(v.avg, 2) + "</td><td>" + f(v.min, 1) + "</td><td>" + f(v.max, 1) +
        "</td><td>" + f(v.sd, 1) + '</td><td class="mono">' + f(v.lsl, 1) + "–" + f(v.usl, 1) + "</td><td>" + esc(v.result) + "</td><td>" +
        (r.ok ? '<span class="pill ok">✓ ' + r.checks.length + "/" + r.checks.length + "</span>" : '<span class="pill bad">✗ ' + failed.length + " failed</span>") +
        "</td><td>" + esc(r.status) + (r.acceptedBy ? '<br><span class="note">' + esc(r.acceptedBy) + "</span>" : "") + "</td></tr>";
      var detail = '<tr><td colspan="13"><details' + (!r.ok && r.status === "pending" ? " open" : "") + '><summary class="note">Pages ' + r.pages.join(", ") + " · checks</summary>" +
        '<div class="grid2" style="margin-top:8px"><div>' + r.pages.map(function (n) { var t = S.scanThumbs[r.file + "#" + n]; return t ? '<img class="thumb" alt="Page ' + n + '" src="' + t + '">' : ""; }).join("") +
        "</div><div><table>" + r.checks.map(function (c) { return "<tr><td>" + (c.ok ? "✓" : "✗") + "</td><td>" + esc(c.name) + '<br><span class="note">' + esc(c.detail) + "</span></td></tr>"; }).join("") + "</table>" +
        (r.status === "pending" ? '<div class="row" style="margin-top:8px">' + keys.map(function (k) { return '<label style="width:80px">' + k + '<input data-k="' + k + '" data-i="' + i + '" value="' + (v[k] === null ? "" : esc(v[k])) + '"></label>'; }).join("") +
          '<label style="flex:1 1 220px">Note' + (r.ok ? "" : " (required: what you checked on the page)") + '<input data-note="' + i + '"></label>' +
          '<button class="b primary" data-acc="' + i + '">Accept</button><button class="b" data-rej="' + i + '">Reject</button></div>' : (r.note ? '<p class="note">Note: ' + esc(r.note) + "</p>" : "")) +
        "</div></div></details></td></tr>";
      return row + detail;
    }).join("");
    $("srecs").querySelectorAll("[data-acc]").forEach(function (b) { b.onclick = function () { decide(+b.dataset.acc, true); }; });
    $("srecs").querySelectorAll("[data-rej]").forEach(function (b) { b.onclick = function () { decide(+b.dataset.rej, false); }; });
  }
  function decide(i, accept) {
    var r = S.scans[i], by = $("sby").value.trim();
    if (!by) { msg("sdecide", "Enter your name first: it is recorded as the verifier.", "err"); $("sby").focus(); return false; }
    store("apqr_by", by);
    var note = ($("srecs").querySelector('[data-note="' + i + '"]') || {}).value || "", changed = {};
    if (accept) {
      var v = Object.assign({}, r.values);
      $("srecs").querySelectorAll('[data-i="' + i + '"]').forEach(function (inp) {
        var k = inp.dataset.k, nv = C.num(inp.value);
        if (nv !== null && nv !== v[k]) { changed[k] = [v[k], nv]; v[k] = nv; }
      });
      if (!r.ok && !note.trim()) { msg("sdecide", "This record did not pass every check: compare it with the page and write what you checked in the note.", "err"); return false; }
      r.note = (note + (Object.keys(changed).length ? " | corrected: " + JSON.stringify(changed) : "")).replace(/^ \| /, "");
      r.values = v;
      r.row = AdigiOCR.contractRow(v, r.pages, r.ok && !Object.keys(changed).length, r.file, by, new Date().toISOString(), Object.keys(changed).length > 0);
    } else r.note = note;
    r.status = accept ? "accepted" : "rejected"; r.acceptedBy = by; r.acceptedAt = new Date().toISOString();
    msg("sdecide", accept ? "Accepted. It is added to the APQR Excel of its batch." : "Rejected; it will not be used.", "ok");
    renderScans(); if (S.all) recompute();
    return true;
  }
  $("sacceptAll").onclick = function () {
    if (!$("sby").value.trim()) { msg("sdecide", "Enter your name first: it is recorded as the verifier.", "err"); $("sby").focus(); return; }
    var n = 0;
    S.scans.forEach(function (r, i) { if (r.status === "pending" && r.ok && decide(i, true)) n++; });
    if (!n) { msg("sdecide", "No pending record passed every check: review the remaining ones one by one.", ""); return; }
    msg("sdecide", n + " record(s) accepted." + (S.all ? " They are included in the APQR sheet download." : " Load the PAS-X export on the APQR sheet tab, or download the records as CSV."), n ? "ok" : "");
  };
  $("scsv").onclick = function () {
    var acc = S.scans.filter(function (r) { return r.status === "accepted"; });
    if (!acc.length) { msg("sdecide", "No accepted records yet.", "err"); return; }
    var cols = Object.keys(acc[0].row), lines = [cols.join(",")].concat(acc.map(function (r) {
      return cols.map(function (c) { var t = String(r.row[c] === undefined ? "" : r.row[c]); return /[",\n]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t; }).join(",");
    }));
    save(new Blob([lines.join("\n") + "\n"], { type: "text/csv" }), "BMR_scan_records_" + stamp() + ".csv");
  };

  /* ------------------------------------------------------------------ SQL builder (offline package only) */
  if (window.ADIGI_SQL) {
    var lit = function (v) { return "'" + String(v).replace(/'/g, "''") + "'"; };
    var dmy = function (iso) { var m = (iso || "").match(/^(\d{4})-(\d{2})-(\d{2})$/); return m ? m[3] + "-" + m[2] + "-" + m[1] : null; };
    var buildSql = function () {
      var mats = $("qMat").value.split(/[,\s]+/).filter(Boolean), bats = $("qBatches").value.split(/[,\s]+/).filter(Boolean);
      var bad = mats.filter(function (c) { return !/^[A-Za-z0-9_-]{1,40}$/.test(c); }).concat(bats.filter(function (c) { return !/^[A-Za-z0-9_/-]{1,40}$/.test(c); }));
      if (bad.length) throw new Error("Not a material or batch number: " + bad.join(", "));
      var from = dmy($("qFrom").value), to = dmy($("qTo").value);
      if (!bats.length && (!mats.length || !from || !to)) throw new Error("Enter material code(s) and both dates, or a batch list.");
      var sql = window.ADIGI_SQL.text, subs = [
        [/CAST\((?:NULL|'[^']*') AS VARCHAR2\(4000\)\)(\s+AS material_list)/, (mats.length ? "CAST(" + lit(mats.join(",")) + " AS VARCHAR2(4000))" : "CAST(NULL AS VARCHAR2(4000))") + "$1"],
        [/TO_DATE\('[^']*', 'DD-MM-YYYY'\)(\s+AS from_date)/, "TO_DATE('" + (from || "01-01-2000") + "', 'DD-MM-YYYY')$1"],
        [/TO_DATE\('[^']*', 'DD-MM-YYYY'\)(\s+AS to_date)/, "TO_DATE('" + (to || "31-12-2099") + "', 'DD-MM-YYYY')$1"],
        [/CAST\((?:NULL|'[^']*') AS VARCHAR2\(4000\)\)(\s+AS batch_list)/, (bats.length ? "CAST(" + lit(bats.join(",")) + " AS VARCHAR2(4000))" : "CAST(NULL AS VARCHAR2(4000))") + "$1"],
        [/'[YN]'(\s+AS apqr_only)/, "'" + $("qOnly").value + "'$1"]];
      subs.forEach(function (x) { if (!x[0].test(sql)) throw new Error("The SQL template changed; cannot fill " + x[0]); sql = sql.replace(x[0], x[1]); });
      return sql;
    };
    var refresh = function () { try { $("qText").value = buildSql(); msg("qMsg", ""); } catch (e) { $("qText").value = ""; msg("qMsg", e.message, "err"); } };
    ["qMat", "qFrom", "qTo", "qBatches", "qOnly"].forEach(function (id) { $(id).addEventListener("input", refresh); $(id).addEventListener("change", refresh); });
    $("qCopy").onclick = function () { refresh(); if (!$("qText").value) return; $("qText").select(); try { navigator.clipboard.writeText($("qText").value); } catch (e) { document.execCommand("copy"); } msg("qMsg", "Copied.", "ok"); };
    $("qSave").onclick = function () { refresh(); if (!$("qText").value) return; save(new Blob([$("qText").value], { type: "text/plain" }), "APQR_" + ($("qMat").value.split(/[,\s]+/)[0] || "batches") + ".sql"); };
    refresh();
  }

  window.AdigiUI = { load: load, state: S, recompute: recompute };
})();
