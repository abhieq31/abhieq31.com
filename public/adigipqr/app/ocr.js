/* AdigiPQR - read scanned BMR printouts in the browser (no upload, no server).
 *
 * Supported printout: INDIVIDUAL WEIGHT VARIATION RECORD (balance software, 1-2 pages per check).
 * Method: render the PDF page (pdf.js), find the ruled table grid geometrically, cut every cell and read each
 * cell on its own with Tesseract (single-line mode; numeric cells limited to digits). Labels are matched
 * to the cell on their right; the weight table maps row x column to the serial number.
 * Every record is then checked against itself exactly like adigipqr/docread.py (Python) does.
 * Libraries are loaded from ./vendor only (works offline and under a strict Content-Security-Policy).
 */
(function (global) {
  "use strict";
  var DPI = 200;
  var state = { base: "", pdfjs: null, txt: null, num: null };

  function assetBase() {
    // works at abhieq31.com/adigipqr (no trailing slash), /adigipqr/, a local server and file://
    var b = document.querySelector("base");
    return b ? b.href : location.href.replace(/[^/]*([?#].*)?$/, "");
  }

  async function init(progress) {
    if (state.txt) return;
    if (location.protocol === "file:") {
      throw new Error("Reading scans needs the page to be served (browsers block workers on file://). " +
        "Use the online page, or start_web_app.bat for a local copy.");
    }
    state.base = assetBase();
    progress && progress("Loading the PDF reader…");
    state.pdfjs = await import(state.base + "vendor/pdfjs/pdf.min.mjs");
    state.pdfjs.GlobalWorkerOptions.workerSrc = state.base + "vendor/pdfjs/pdf.worker.min.mjs";
    progress && progress("Loading the text recogniser (first time ~10 MB, then cached by the browser)…");
    var opts = { workerPath: state.base + "vendor/tesseract/worker.min.js", corePath: state.base + "vendor/tesseract/core",
      langPath: state.base + "vendor/tesseract/lang", gzip: true };
    state.txt = await Tesseract.createWorker("eng", 1, opts);
    await state.txt.setParameters({ tessedit_pageseg_mode: "7", preserve_interword_spaces: "1" });
    state.num = await Tesseract.createWorker("eng", 1, opts);
    await state.num.setParameters({ tessedit_pageseg_mode: "7", tessedit_char_whitelist: "0123456789.-%/:" });
    state.strip = await Tesseract.createWorker("eng", 1, opts);
    await state.strip.setParameters({ tessedit_pageseg_mode: "11" });
  }

  async function openPdf(bytes) {
    return state.pdfjs.getDocument({ data: bytes }).promise;
  }

  async function renderPage(pdf, n, dpi) {
    var page = await pdf.getPage(n);
    var vp = page.getViewport({ scale: (dpi || DPI) / 72 });
    var c = document.createElement("canvas");
    c.width = Math.round(vp.width); c.height = Math.round(vp.height);
    var ctx = c.getContext("2d", { willReadFrequently: true });
    ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, c.width, c.height);
    await page.render({ canvasContext: ctx, viewport: vp }).promise;
    return c;
  }

  function imageToCanvas(img) {
    var c = document.createElement("canvas"); c.width = img.naturalWidth; c.height = img.naturalHeight;
    var ctx = c.getContext("2d", { willReadFrequently: true }); ctx.drawImage(img, 0, 0); return c;
  }

  /* ---------------------------------------------------------------- geometry */
  function binarize(canvas) {
    var W = canvas.width, H = canvas.height, d = canvas.getContext("2d", { willReadFrequently: true }).getImageData(0, 0, W, H).data;
    var ink = new Uint8Array(W * H);
    for (var i = 0; i < W * H; i++) ink[i] = (0.299 * d[4 * i] + 0.587 * d[4 * i + 1] + 0.114 * d[4 * i + 2]) < 150 ? 1 : 0;
    return { W: W, H: H, ink: ink };
  }

  function cluster(vals, gap) {
    var out = [], cur = [];
    vals.forEach(function (v) {
      if (cur.length && v - cur[cur.length - 1] > gap) { out.push(cur); cur = []; }
      cur.push(v);
    });
    if (cur.length) out.push(cur);
    return out.map(function (g) { return { a: g[0], b: g[g.length - 1], c: (g[0] + g[g.length - 1]) / 2 }; });
  }

  /* Table rows = bands between long horizontal rules; cells = gaps between vertical rules inside a band. */
  function grid(bin) {
    var W = bin.W, H = bin.H, ink = bin.ink, minRun = Math.round(W * 0.18), ys = [];
    for (var y = 0; y < H; y++) {
      var run = 0, best = 0;
      for (var x = 0; x < W; x++) { if (ink[y * W + x]) { run++; if (run > best) best = run; } else run = 0; }
      if (best >= minRun) ys.push(y);
    }
    var lines = cluster(ys, 3), rows = [];
    for (var i = 0; i + 1 < lines.length; i++) {
      var y0 = lines[i].b + 1, y1 = lines[i + 1].a - 1, h = y1 - y0;
      if (h < 18 || h > 140) continue;
      var t0 = y0 + Math.round(h * 0.15), t1 = y1 - Math.round(h * 0.15), need = (t1 - t0) * 0.85, xs = [];
      for (var x2 = 0; x2 < W; x2++) {
        var cnt = 0;
        for (var y2 = t0; y2 < t1; y2++) cnt += ink[y2 * W + x2];
        if (cnt >= need) xs.push(x2);
      }
      var seps = cluster(xs, 3), cells = [];
      for (var k = 0; k + 1 < seps.length; k++) {
        var x0 = seps[k].b + 1, x1 = seps[k + 1].a - 1;
        if (x1 - x0 > 14) cells.push({ x0: x0, x1: x1, y0: y0, y1: y1 });
      }
      // the scanner's page edge adds empty cells at the row ends: drop those (blank cells inside a row stay)
      while (cells.length && !hasInk(bin, cells[0])) cells.shift();
      while (cells.length && !hasInk(bin, cells[cells.length - 1])) cells.pop();
      if (cells.length >= 2) rows.push({ y0: y0, y1: y1, cells: cells });
    }
    // a rule lost in the scan leaves the last cell open: close it at the right edge of the neighbouring rows
    rows.forEach(function (r, i) {
      var near = rows.filter(function (o, j) { return j !== i && Math.abs(o.y0 - r.y0) < 260; })
        .map(function (o) { return o.cells[o.cells.length - 1].x1; }).sort(function (a, b) { return a - b; });
      if (!near.length) return;
      var edge = near[Math.floor(near.length / 2)], last = r.cells[r.cells.length - 1];
      if (edge - last.x1 > 60) {
        var extra = { x0: last.x1 + 5, x1: edge, y0: r.y0, y1: r.y1 };
        if (hasInk(bin, extra)) r.cells.push(extra);
      }
    });
    return rows;
  }

  function hasInk(bin, c) {
    var pad = 4, n = 0;
    for (var y = c.y0 + pad; y < c.y1 - pad; y++) for (var x = c.x0 + pad; x < c.x1 - pad; x++) n += bin.ink[y * bin.W + x];
    return n > 12;                                         // more than a speck of dust
  }

  /* Cut a cell, whiten leftover rule fragments at its edges, pad, return a canvas for Tesseract. */
  function cellCanvas(src, bin, c, scale) {
    var pad = 3, x0 = c.x0 + pad, y0 = c.y0 + pad, w = c.x1 - c.x0 - 2 * pad, h = c.y1 - c.y0 - 2 * pad;
    if (w < 6 || h < 6) return null;
    var any = false;
    for (var y = y0; y < y0 + h && !any; y++) for (var x = x0; x < x0 + w; x++) if (bin.ink[y * bin.W + x]) { any = true; break; }
    if (!any) return null;
    // 2x upscale and pure black/white: Tesseract reads best at ~30 px character height, scans give ~18 px
    var s = scale || 2, m = 16, out = document.createElement("canvas");
    out.width = w * s + 2 * m; out.height = h * s + 2 * m;
    var ctx = out.getContext("2d", { willReadFrequently: true });
    ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, out.width, out.height);
    ctx.imageSmoothingEnabled = true; ctx.imageSmoothingQuality = "high";
    ctx.drawImage(src, x0, y0, w, h, m, m, w * s, h * s);
    var img = ctx.getImageData(0, 0, out.width, out.height), d = img.data;
    for (var i = 0; i < d.length; i += 4) {
      var v = (0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2]) < 160 ? 0 : 255;
      d[i] = d[i + 1] = d[i + 2] = v; d[i + 3] = 255;
    }
    ctx.putImageData(img, 0, 0);
    return out;
  }

  async function read(worker, canvas) {
    if (!canvas) return { text: "", conf: null };
    var r = await worker.recognize(canvas);
    return { text: (r.data.text || "").replace(/\s+/g, " ").trim(), conf: r.data.confidence };
  }

  /* ---------------------------------------------------------------- labels */
  function norm(t) { return String(t || "").toLowerCase().replace(/[^a-z0-9%]/g, ""); }
  function ratio(a, b) {                                   // similarity like difflib, via LCS length
    if (!a || !b) return 0;
    var m = a.length, n = b.length, prev = new Array(n + 1).fill(0);
    for (var i = 1; i <= m; i++) {
      var cur = [0];
      for (var j = 1; j <= n; j++) cur[j] = a[i - 1] === b[j - 1] ? prev[j - 1] + 1 : Math.max(prev[j], cur[j - 1]);
      prev = cur;
    }
    return 2 * prev[n] / (m + n);
  }
  function matchLabel(text, label, cutoff) {
    var a = norm(text), b = norm(label);
    if (!a || !b) return 0;
    if (a === b) return 1;
    if (a.length >= 0.6 * b.length && b.slice(-a.length) === a) return 0.95;   // punch hole clipped the start
    var r = ratio(a, b);
    return r >= (cutoff || 0.8) ? r : 0;
  }

  var HEADER = { product_name: "Product Name", product_code: "Product Code", batch_no: "Batch No", serial: "Serial No.",
    balance: "Balance ID No.", start: "Start Date & Time", n_samples: "No. of Samples", std_weight: "Standard Weight of Tab. (mg)",
    stage: "Stage", batch_size: "Batch Size", interval: "Interval", side: "Side", end: "End Date & Time", limit: "Limit",
    result: "Test Result" };
  var SUMMARY = { avg: "Average Weight of Tab. (mg)", avg_pct: "Average %", above1: "No. of Tablet Above Limit 1",
    below1: "No. of Tablet Below Limit 1", above2: "No. of Tablet Above Limit 2", below2: "No. of Tablet Below Limit 2",
    min: "Minimum Weight (mg)", max: "Maximum Weight (mg)", min_pct: "Minimum %", max_pct: "Maximum %",
    sd: "Standard Deviation", range: "R( Max. Wt. - Min. Wt.)" };
  // cells without a unit are read digits-only; cells printed with "mg" / "%" are read as text (a digits-only
  // reader turns "mg" into digits) and the number is taken from the text
  var NUMERIC = { serial: 1, n_samples: 1, sd: 1, above1: 1, below1: 1, above2: 1, below2: 1, batch_no: 1, product_code: 1 };
  var ALL = Object.assign({}, HEADER, SUMMARY);

  /* "Maximum Weight (mg) 24.1 mg" -> "24.1 mg" when the text starts with the label */
  function splitLabel(text, label) {
    var words = String(text || "").split(" "), target = norm(label);
    for (var i = words.length - 1; i >= 1; i--) {
      var head = norm(words.slice(0, i).join(" ")), tail = words.slice(i).join(" ");
      if (head.length >= 6 && /\d/.test(tail) && (head === target || ratio(head, target) >= 0.9)) return tail;
    }
    return null;
  }

  function classifyText(t) {
    var n = norm(t);
    if (n.indexOf("weightvariation") >= 0 || (n.indexOf("individual") >= 0 && n.indexOf("variation") >= 0)) return "weight_variation";
    return null;
  }

  async function classify(canvas) {
    var h = Math.round(canvas.height * 0.12), strip = document.createElement("canvas");
    strip.width = Math.round(canvas.width / 2); strip.height = Math.round(h / 2);
    strip.getContext("2d").drawImage(canvas, 0, 0, canvas.width, h, 0, 0, strip.width, strip.height);
    var r = await state.strip.recognize(strip);
    return classifyText(r.data.text);
  }

  function fixDecimal(text, decimals) {
    // printed "23.5" read as "235": the format is fixed, so put the point back; the checks verify the result
    var t = text.replace(/[^0-9.]/g, "");
    if (!t) return null;
    if (t.indexOf(".") < 0 && decimals && t.length > decimals) t = t.slice(0, -decimals) + "." + t.slice(-decimals);
    var v = parseFloat(t);
    return isFinite(v) ? v : null;
  }

  function overlap(a, b) { return Math.max(0, Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0)); }

  async function parseWeightVariation(canvas, progress) {
    var bin = binarize(canvas), rows = grid(bin), fields = {}, weights = {}, head = null, rowNo = 0;
    var done = 0, total = rows.reduce(function (s, r) { return s + r.cells.length; }, 0);
    for (var ri = 0; ri < rows.length; ri++) {
      var cells = rows[ri].cells;
      if (head && ri > 0 && rows[ri].y0 - rows[ri - 1].y1 > 25) head = null;    // white space ends the table
      if (head) {
        // weight table row: map each header column to the cell that overlaps it most; stray cells are ignored
        var mapped = head.map(function (h) {
          var best = null, ov = 0, hw = h.x1 - h.x0;
          cells.forEach(function (c) {
            var o = overlap(c, h);
            if (o > ov && (c.x1 - c.x0) < hw * 1.5) { ov = o; best = c; }   // a table cell is about one column wide
          });
          return ov > hw * 0.6 ? best : null;
        });
        if (mapped.filter(Boolean).length >= head.length * 0.8) {
          rowNo++;
          for (var c = 0; c < head.length; c += 2) {
            var wc = mapped[c + 1];
            if (!wc) continue;
            var wv = await read(state.num, cellCanvas(canvas, bin, wc));
            done += 2;
            var val = fixDecimal(wv.text, 1);
            if (val !== null) weights[(rowNo - 1) * (head.length / 2) + c / 2 + 1] = { v: val, conf: wv.conf };
          }
          progress && progress(done, total);
          continue;
        }
        head = null;                                          // the table has ended
      }
      var texts = [];
      for (var k = 0; k < cells.length; k++) {
        texts.push(await read(state.txt, cellCanvas(canvas, bin, cells[k])));
        done++;
      }
      if (state.debugRows) state.debugRows.push([ri, cells.map(function (c) { return c.x0 + "-" + c.x1; }).join(" "), texts.map(function (x) { return x.text; }).join(" | ")]);
      // weight table header: alternating "Sr No." / "Weight in (mg)"
      var srs = texts.filter(function (x) { return /^sr/i.test(norm(x.text)); }).length;
      var wts = texts.filter(function (x) { return /weight|mg/i.test(x.text); }).length;
      if (srs >= 3 && wts >= 3) {
        var start = texts.findIndex(function (x) { return /^sr/i.test(norm(x.text)); });
        head = cells.slice(start);
        if (head.length % 2) head = head.slice(0, -1);
        rowNo = 0;
        continue;
      }
      // label | value: a cell that matches a known label takes the next cell as its value. When no rule
      // separates them, the label cell also holds the value ("Maximum Weight (mg) 24.1 mg").
      for (var j = 0; j < cells.length; j++) {
        var best = null, bestScore = 0, rest = null;
        Object.keys(ALL).forEach(function (key) {
          if (fields[key]) return;
          var sc = matchLabel(texts[j].text, ALL[key]), tail = null;
          if (!sc) {
            var cut = splitLabel(texts[j].text, ALL[key]);
            if (cut) { sc = 0.9; tail = cut; }
          }
          if (sc > bestScore) { best = key; bestScore = sc; rest = tail; }
        });
        if (!best) continue;
        if (rest !== null) { fields[best] = { text: rest, conf: texts[j].conf }; continue; }
        if (j + 1 >= cells.length) continue;
        var value = texts[j + 1];
        if (NUMERIC[best]) {
          var n = await read(state.num, cellCanvas(canvas, bin, cells[j + 1]));
          // two readings of the same cell: prefer the digits-only one unless the text reading holds more digits
          if (n.text && n.text.replace(/\D/g, "").length >= value.text.replace(/\D/g, "").length) value = n;
          if (best === "serial" || best === "n_samples") {
            // short numbers: "11" is two thin strokes; a 4x reading separates them - keep the more confident one
            var n4 = await read(state.num, cellCanvas(canvas, bin, cells[j + 1], 4));
            if (n4.text && (n4.conf || 0) > (value.conf || 0)) value = n4;
          }
        }
        fields[best] = value;
        j++;
      }
      progress && progress(done, total);
    }
    return { type: "weight_variation", fields: fields, weights: weights };
  }

  /* ---------------------------------------------------------------- records and checks (= docread.py) */
  function num(t) { var m = String(t || "").replace(",", ".").match(/-?\d+(?:\.\d+)?/); return m ? parseFloat(m[0]) : null; }
  function digits(t) { return String(t || "").replace(/\D/g, ""); }

  function mergePages(parsed) {
    var recs = {}, order = [];
    parsed.forEach(function (p) {
      var f = p.fields, g = function (k) { return f[k] ? f[k].text : ""; };
      var key = [digits(g("batch_no")), digits(g("start")), digits(g("end")), norm(g("side"))].join("|");
      if (!recs[key]) { recs[key] = { type: p.type, fields: {}, weights: {}, pages: [] }; order.push(key); }
      var r = recs[key];
      Object.keys(f).forEach(function (k) { if (!r.fields[k] || !r.fields[k].text) r.fields[k] = f[k]; });
      Object.keys(p.weights).forEach(function (k) { r.weights[k] = p.weights[k]; });
      r.pages.push(p.page_no);
    });
    return order.map(function (k) { return recs[k]; });
  }

  function mean(a) { return a.reduce(function (s, x) { return s + x; }, 0) / a.length; }
  function sd(a) { var m = mean(a); return Math.sqrt(a.reduce(function (s, x) { return s + (x - m) * (x - m); }, 0) / (a.length - 1)); }

  function evaluate(r) {
    var f = r.fields, g = function (k) { return f[k] ? f[k].text : ""; };
    // limits print with one decimal ("21.2 - 25.9 mg"); a lost point is put back and verified by the limit check
    var lim = (g("limit").replace(/,/g, ".").match(/\d+(?:\.\d+)?/g) || []).map(function (x) {
      return x.indexOf(".") < 0 && x.length >= 3 ? fixDecimal(x, 1) : parseFloat(x);
    });
    var lsl = lim.length >= 2 ? lim[0] : null, usl = lim.length >= 2 ? lim[1] : null;
    var v = {
      batch: digits(g("batch_no")), product_code: digits(g("product_code")), serial: digits(g("serial")),
      interval: g("interval").replace(/\s+/g, "").toUpperCase(), side: g("side").replace(/[^A-Za-z]/g, "").toUpperCase(),
      balance: g("balance").replace(/\s+/g, "").replace(/([A-Z])\.(?=[A-Z])/g, "$1"), start: g("start"), end: g("end"), stage: g("stage"), product: g("product_name"),
      result: g("result"), n: num(g("n_samples")), lsl: lsl, usl: usl, std: num(g("std_weight")),
      avg: fixDecimal(g("avg").split(/m/i)[0], 2), min: fixDecimal(g("min").split(/m/i)[0], 1),
      max: fixDecimal(g("max").split(/m/i)[0], 1), sd: fixDecimal(g("sd"), 1), range: fixDecimal(g("range").split(/m/i)[0], 1),
      min_pct: fixDecimal(g("min_pct"), 2), max_pct: fixDecimal(g("max_pct"), 2), avg_pct: fixDecimal(g("avg_pct"), 2)
    };
    var keys = Object.keys(r.weights).map(Number).sort(function (a, b) { return a - b; });
    var w = keys.map(function (k) { return r.weights[k].v; });
    v.n_read = w.length;
    var checks = [];
    function chk(name, ok, detail) { checks.push({ name: name, ok: !!ok, detail: detail }); }
    chk("Batch, serial and limit read", v.batch && v.serial && lsl !== null,
      "batch " + (v.batch || "?") + ", serial " + (v.serial || "?") + ", limit " + lsl + "-" + usl);
    chk("Summary block read", [v.avg, v.min, v.max, v.sd].every(function (x) { return x !== null; }),
      "avg " + v.avg + ", min " + v.min + ", max " + v.max + ", SD " + v.sd);
    var exp = v.n ? Math.round(v.n) : null, missing = [];
    if (exp !== null && !w.length) chk("All individual weights read", false, "0 of " + exp + " read (weight table not found)");
    if (w.length) {
      if (exp) for (var i = 1; i <= exp; i++) if (!(i in r.weights)) missing.push(i);
      chk("All individual weights read", exp !== null && w.length === exp && !missing.length,
        w.length + " of " + exp + " read" + (missing.length ? ", missing Sr " + missing.slice(0, 12).join(", ") : ""));
      if (v.min !== null) {
        chk("Min = smallest weight", Math.abs(Math.min.apply(null, w) - v.min) < 1e-6, Math.min.apply(null, w) + " vs printed " + v.min);
        chk("Max = largest weight", Math.abs(Math.max.apply(null, w) - v.max) < 1e-6, Math.max.apply(null, w) + " vs printed " + v.max);
      }
      if (v.avg !== null) {
        var m = mean(w);
        chk("Average of weights = printed average", Math.abs(m - v.avg) <= 0.0051 + 0.05 / Math.max(1, w.length), m.toFixed(3) + " vs " + v.avg);
      }
      if (v.sd !== null && w.length > 1) chk("SD of weights = printed SD", Math.abs(sd(w) - v.sd) <= 0.051, sd(w).toFixed(3) + " vs " + v.sd);
    }
    if (v.avg !== null && v.min !== null && v.max !== null && v.avg) {
      if (v.range !== null) chk("Range = max - min", Math.abs(v.max - v.min - v.range) <= 0.051, (v.max - v.min).toFixed(1) + " vs " + v.range);
      if (v.min_pct !== null && v.max_pct !== null) {
        var a = (v.avg - v.min) / v.avg * 100, b = (v.max - v.avg) / v.avg * 100;
        chk("Min % / Max % consistent", Math.abs(a - v.min_pct) <= 0.06 && Math.abs(b - v.max_pct) <= 0.06,
          a.toFixed(2) + "/" + b.toFixed(2) + " vs " + v.min_pct + "/" + v.max_pct);
      }
    }
    if (lsl !== null && v.min !== null) {
      var inside = v.min >= lsl && v.max <= usl;
      chk("Printed result agrees with limits", inside === /pass/i.test(v.result), v.min + "-" + v.max + " in " + lsl + "-" + usl + ": printed '" + v.result + "'");
    }
    if (lsl !== null && usl !== null && (v.avg !== null || v.std !== null)) {
      var ok = [v.avg, v.std].filter(function (x) { return x !== null; }).some(function (centre) {
        return [5, 7.5, 10].some(function (p) {
          return Math.abs(lsl - Math.round(centre * (1 - p / 100) * 10) / 10) <= 0.051 && Math.abs(usl - Math.round(centre * (1 + p / 100) * 10) / 10) <= 0.051;
        });
      });
      chk("Limit = average or standard weight ± 5 / 7.5 / 10 %", ok, lsl + "-" + usl + " for average " + v.avg + ", standard " + v.std);
    }
    r.values = v; r.checks = checks; r.ok = checks.every(function (c) { return c.ok; });
    return r;
  }

  function startKey(r) {
    var m = String(r.values.start || "").match(/(\d{2})\/(\d{2})\/(\d{4})\D*(\d{2}):(\d{2})/);
    return m ? m[3] + m[2] + m[1] + m[4] + m[5] : "9999";
  }

  function checkSequence(recs) {
    var by = {};
    recs.forEach(function (r) { (by[r.values.batch] = by[r.values.batch] || []).push(r); });
    Object.keys(by).forEach(function (b) {
      var rs = by[b].sort(function (x, y) { return startKey(x) < startKey(y) ? -1 : startKey(x) > startKey(y) ? 1 : 0; });
      var s = rs.map(function (r) { return /^\d+$/.test(r.values.serial) ? parseInt(r.values.serial, 10) : null; });
      rs.forEach(function (r, i) {
        var prev = null, next = null;
        for (var j = i - 1; j >= 0; j--) if (s[j] !== null) { prev = s[j]; break; }
        for (var k = i + 1; k < s.length; k++) if (s[k] !== null) { next = s[k]; break; }
        var ok = s[i] !== null && (prev === null || s[i] > prev) && (next === null || s[i] < next);
        r.checks.push({ name: "Serial no. rises with start time", ok: ok, detail: "serial " + s[i] + " between " + prev + " and " + next });
        r.ok = r.checks.every(function (c) { return c.ok; });
      });
    });
    return recs;
  }

  /* One APQR row per check, in the common long format (same as docread.contract_rows). */
  function contractRow(v, pages, autoOk, sourceFile, by, at, corrected) {
    var stage = (v.stage || "Compression").trim(), lower = stage.toLowerCase();
    var sheet = ["compression", "coating", "capsule", "granulation", "packing"].indexOf(lower) >= 0
      ? stage.charAt(0).toUpperCase() + stage.slice(1).toLowerCase() : "Compression";
    var step = ("Weight variation " + v.interval.charAt(0) + v.interval.slice(1).toLowerCase() + " " + v.side + " #" + v.serial).trim();
    var f = function (x) { return x === null || x === undefined ? "" : String(x); };
    var out = v.lsl !== null && v.min !== null && (v.min < v.lsl || v.max > v.usl);
    return { BATCH_NO: v.batch, MATERIAL_CODE: v.product_code, PRODUCT_NAME: v.product, APQR_SHEET: sheet, STAGE: stage, STEP: step,
      STEP_PATH: stage + " > " + step, STEP_CODE_PATH: "SCAN > WV > " + v.interval + " > " + v.side + " > " + v.serial,
      APQR_PARAMETER: "Uniformity of Weight (balance printout)", PARAMETER_CLASS: "IPQC (multi-value)", UOM: "mg",
      SPEC_MIN: f(v.lsl), SPEC_MAX: f(v.usl), TARGET_SET_VALUE: f(v.std), READING_MIN: f(v.avg), READING_MAX: f(v.avg),
      READING_AVG: f(v.avg), READING_COUNT: "1", INDIVIDUAL_MIN: f(v.min), INDIVIDUAL_MAX: f(v.max), INDIVIDUAL_AVG: f(v.avg),
      INDIVIDUAL_COUNT: String(Math.round(v.n || v.n_read || 0)), LIMIT_CHECK: out ? "OUT OF LIMIT - CHECK" : "WITHIN LIMIT",
      EQUIPMENT_IDS: v.balance, FIRST_READING_ON: v.start, LAST_READING_ON: v.end, SOURCE_SYSTEM: "BMR scan - balance printout",
      SOURCE_FILE: sourceFile, SOURCE_PAGE: pages.join(","), CONFIDENCE: corrected ? "corrected by person"
        : autoOk ? "all arithmetic checks passed" : "checked against page by person", VERIFIED_BY: by, VERIFIED_ON: at };
  }

  /* Read a PDF or image file. progress({stage, page, pages, found}) */
  async function readDocument(file, pageFilter, progress) {
    await init(function (msg) { progress && progress({ stage: msg }); });
    var bytes = new Uint8Array(await file.arrayBuffer()), parsed = [], log = [], thumbs = {};
    var pages = [], getCanvas;
    if (/\.pdf$/i.test(file.name) || file.type === "application/pdf") {
      var pdf = await openPdf(bytes);
      for (var i = 1; i <= pdf.numPages; i++) if (!pageFilter || pageFilter.has(i)) pages.push(i);
      getCanvas = function (n) { return renderPage(pdf, n); };
    } else {
      var url = URL.createObjectURL(file), img = new Image();
      await new Promise(function (res, rej) { img.onload = res; img.onerror = rej; img.src = url; });
      pages = [1];
      getCanvas = async function () { return imageToCanvas(img); };
    }
    for (var p = 0; p < pages.length; p++) {
      var n = pages[p], canvas = await getCanvas(n), typ = await classify(canvas);
      log.push([n, typ]);
      progress && progress({ stage: "Page " + n + (typ ? ": weight variation record, reading cells…" : ": not a known printout"),
        page: p + 1, pages: pages.length, found: parsed.length });
      if (typ) {
        var res = await parseWeightVariation(canvas, function (done, total) {
          progress && progress({ stage: "Page " + n + ": " + done + "/" + total + " cells", page: p + 1, pages: pages.length, found: parsed.length });
        });
        res.page_no = n;
        parsed.push(res);
        thumbs[n] = canvas.toDataURL("image/jpeg", 0.7);
      }
    }
    var recs = checkSequence(mergePages(parsed).map(evaluate));
    return { records: recs, log: log, thumbs: thumbs };
  }

  global.AdigiOCR = { readDocument: readDocument, evaluate: evaluate, mergePages: mergePages, checkSequence: checkSequence,
    contractRow: contractRow, grid: grid, binarize: binarize, matchLabel: matchLabel, fixDecimal: fixDecimal, _state: state };
})(window);
