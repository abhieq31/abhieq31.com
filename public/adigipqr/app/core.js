/* AdigiPQR core - the APQR workbook built in the browser. Port of adigipqr/pivot.py + extras.py + stats.py:
 * same sheets, same numbers (tests/test_parity.py compares every cell with the Python output).
 * Nothing is sent anywhere: the export is read, converted and saved on this computer.
 */
(function (global) {
  "use strict";
  var VERSION = "2.2.0";
  var IPC_RE = /in.?process|ipc|check|record|parameter|monitor/i;
  var PHASE_RE = /^(PRE.?LUB\w*|LUB\w*|LOT\w*|PART\w*|PREHEAT\w*|PRE\w*IPC|SPRAY\w*|DRY\w*|COOL\w*)$/i;
  var SEC_RE = /^(s|sec|secs|second|seconds)$/i;
  var SHEETS = ["Granulation", "Compression", "Coating", "Capsule", "Packing", "Inspection"];
  var REQUIRED = ["BATCH_NO", "APQR_SHEET", "STAGE", "APQR_PARAMETER", "PARAMETER_CLASS", "READING_MIN"];
  var FILL_COLS = ["STEP", "STEP_PATH", "STEP_CODE_PATH", "UOM", "EQUIPMENT_IDS", "PRODUCTION_UNIT_ID", "LIMIT_CHECK",
    "SPEC_MIN", "SPEC_MAX", "TARGET_SET_VALUE", "READING_MAX", "READING_AVG", "READING_COUNT", "INDIVIDUAL_MIN",
    "INDIVIDUAL_MAX", "INDIVIDUAL_AVG", "INDIVIDUAL_COUNT", "SUPERSEDED_VALUES"];
  var NELSON = { 1: "One point beyond 3 sigma", 2: "Nine points in a row on one side of the centre line",
    3: "Six points in a row steadily increasing or decreasing" };
  var BPR_COLS = ["Sr. No.", "Manufacturing Batch No.", "Packing Order", "Packing Material", "Stage", "Packing Start",
    "Packing End", "Production Unit", "Equipment ID", "Pack Style", "Market", "Remarks"];

  /* ---------------------------------------------------------------- helpers (= pivot.py) */
  function s(v) { return v === null || v === undefined || (typeof v === "number" && isNaN(v)) ? "" : String(v).trim(); }
  function num(v) {
    var t = s(v);
    if (!t) return null;
    if (/^-?\d+,\d+$/.test(t)) t = t.replace(",", ".");
    t = t.replace(/,/g, "");
    if (!/^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(t)) return null;
    var n = Number(t);
    return isFinite(n) ? n : null;
  }
  function pyRound(n, d) {
    // Python round(): decided on the exact decimal expansion of the double, ties to the even digit
    if (!isFinite(n) || n === 0) return n;
    var neg = n < 0, digits = Math.abs(n).toPrecision(21);           // 21 significant digits of the exact value
    var m = digits.match(/^(\d+)(?:\.(\d+))?$/);
    if (!m) return Number(n.toFixed(d));                                 // exponent form: tiny/huge, no rounding issue
    var ip = m[1], fp = (m[2] || "").padEnd(d + 1, "0"), keep = fp.slice(0, d), rest = fp.slice(d);
    var up = rest[0] > "5" || (rest[0] === "5" && (/[1-9]/.test(rest.slice(1)) || ((ip + keep).slice(-1) % 2 === 1)));
    var v = Number(ip + (d ? "." + keep : ""));
    if (up) v = Number((v + Math.pow(10, -d)).toFixed(d));
    return neg ? -v : v;
  }
  function fmt(n) {                               // '%g' % round(n, 4)
    if (n === null || n === undefined) return "";
    var r = pyRound(n, 4);
    if (r === 0) return "0";
    var a = Math.abs(r);
    if (a >= 1e6 || a < 1e-4) {
      var e = r.toExponential(5).replace(/\.?0+e/, "e");
      return e.replace(/e([+-])(\d)$/, "e$10$2");
    }
    return String(parseFloat(r.toPrecision(6)));
  }
  function cleanStage(x) { return s(x).replace(/\s*\(\s*[\d.,]+\s*kg\s*\)\s*/ig, " ").replace(/\s+/g, " ").trim() || "Stage"; }
  function phaseOf(code) {
    var t = s(code).split(">").map(function (x) { return x.trim(); });
    for (var i = t.length - 1; i >= 0; i--) if (PHASE_RE.test(t[i])) return t[i].toUpperCase();
    return "";
  }
  function limitText(lo, hi, tg) {
    var t = lo !== null && hi !== null ? fmt(lo) + " to " + fmt(hi) : hi !== null ? "NMT " + fmt(hi)
      : lo !== null ? "NLT " + fmt(lo) : "To be recorded";
    return t + (tg !== null && tg !== undefined ? " (Target " + fmt(tg) + ")" : "");
  }
  function pySort(a) { return a.slice().sort(function (x, y) { return x < y ? -1 : x > y ? 1 : 0; }); }

  /* ---------------------------------------------------------------- input */
  function normalise(raw, columns) {
    var cols = (columns || (raw[0] ? Object.keys(raw[0]) : [])).map(function (c) { return String(c).trim().toUpperCase(); });
    var miss = REQUIRED.filter(function (c) { return cols.indexOf(c) < 0; });
    if (miss.length) throw new Error("Missing columns " + miss.join(", ") + ". Export the result of 01_APQR_Final_PASX.sql with column headers on.");
    var rows = raw.map(function (r) {
      var o = {};
      Object.keys(r).forEach(function (k) { o[String(k).trim().toUpperCase()] = r[k] === null || r[k] === undefined ? "" : r[k]; });
      return o;
    }).filter(function (r) { return s(r.BATCH_NO) !== ""; });
    rows.forEach(function (r) {
      r.BATCH_NO = s(r.BATCH_NO).replace(/\.0$/, "");
      FILL_COLS.forEach(function (c) { if (!(c in r)) r[c] = ""; });
    });
    rows.sourceColumns = cols;
    return rows;
  }

  /* ---------------------------------------------------------------- model (= pivot.make_sheet / build) */
  function Col(r, values, oos, text, meta) { return { r: r, values: values, oos: oos || {}, text: !!text, meta: meta || {} }; }

  function makeSheet(rows, batches, sel, keyFn, sectionFn, paramFn, kind, opt) {
    var sub = rows.filter(sel);
    if (!sub.length) return null;
    var groups = {}, order = [];
    sub.forEach(function (r) {
      var k = keyFn(r);
      if (!groups[k]) { groups[k] = { stage: cleanStage(r.STAGE), param: paramFn(r), uom: s(r.UOM), rows: [], sections: [] }; order.push(k); }
      var g = groups[k];
      g.rows.push(r);
      var sec = sectionFn(r);
      if (g.sections.indexOf(sec) < 0) g.sections.push(sec);
    });
    var cols = [];
    if (kind !== "reading") {
      var stages = [];
      order.forEach(function (k) { if (stages.indexOf(groups[k].stage) < 0) stages.push(groups[k].stage); });
      var eq = {};
      sub.forEach(function (r) {
        var ident = s(r.EQUIPMENT_IDS) || s(r.PRODUCTION_UNIT_ID);
        if (!ident) return;
        var key = r.BATCH_NO + "\u0001" + cleanStage(r.STAGE);
        eq[key] = eq[key] || {};
        ident.split(",").forEach(function (i) { eq[key][i.trim()] = 1; });
      });
      stages.forEach(function (st) {
        var vals = {};
        batches.forEach(function (b) { vals[b] = pySort(Object.keys(eq[b + "\u0001" + st] || {})).join(", "); });
        cols.push(Col(["Equipment ID", st, "", "", ""], vals, null, true));
      });
    }
    order.forEach(function (k) {
      var g = groups[k];
      var section = g.sections.length === 1 ? g.sections[0] : g.sections.slice(0, 3).join(" / ");
      var toMin = opt.minutes && SEC_RE.test(g.uom), f = toMin ? 1 / 60 : 1, uom = toMin ? "min" : g.uom;
      var per = {}, perOrder = [], los = [], his = [], tgs = [];
      g.rows.forEach(function (r) {
        var ind = kind !== "reading" && (num(r.INDIVIDUAL_COUNT) || 0) > 0;
        var mn = num(ind ? r.INDIVIDUAL_MIN : r.READING_MIN), mx = num(ind ? r.INDIVIDUAL_MAX : r.READING_MAX);
        var av = num(ind ? r.INDIVIDUAL_AVG : r.READING_AVG), n = num(ind ? r.INDIVIDUAL_COUNT : r.READING_COUNT) || 1;
        if (mn === null && mx === null) return;
        if (mn === null) mn = mx;
        if (mx === null) mx = mn;
        if (av === null) av = (mn + mx) / 2;
        var p = per[r.BATCH_NO];
        if (!p) { p = per[r.BATCH_NO] = { mn: Infinity, mx: -Infinity, sm: 0, n: 0, oos: false }; perOrder.push(r.BATCH_NO); }
        p.mn = Math.min(p.mn, mn * f); p.mx = Math.max(p.mx, mx * f);
        p.sm += av * f * n; p.n += n;
        if (/^OUT/.test(s(r.LIMIT_CHECK).toUpperCase())) p.oos = true;
        [[los, "SPEC_MIN"], [his, "SPEC_MAX"], [tgs, "TARGET_SET_VALUE"]].forEach(function (x) {
          var v = num(r[x[1]]);
          if (v !== null) x[0].push(v * f);
        });
      });
      if (!perOrder.length) return;
      var lo = los.length ? Math.min.apply(null, los) : null, hi = his.length ? Math.max.apply(null, his) : null,
        tg = tgs.length ? tgs[0] : null;
      var lim = limitText(lo, hi, tg), p3 = g.param + (uom ? " [" + uom + "]" : "");
      var oos = {}, mins = {}, maxs = {}, avgs = {};
      perOrder.forEach(function (b) {
        var p = per[b];
        oos[b] = p.oos; mins[b] = p.mn; maxs[b] = p.mx; avgs[b] = p.n ? p.sm / p.n : null;
      });
      var single = perOrder.every(function (b) { return Math.abs(per[b].mx - per[b].mn) < 1e-9; });
      var r0 = g.rows[0];
      var meta = { key: k, stage: g.stage, section: section, param: g.param, uom: uom, lo: lo, hi: hi, target: tg,
        cls: s(r0.PARAMETER_CLASS), critical: s(r0.PARAMETER_CRITICALITY), batch_min: mins, batch_max: maxs };
      if (single) cols.push(Col([g.stage, section, p3, lim, "Value"], mins, oos, false, meta));
      else {
        cols.push(Col([g.stage, section, p3, lim, "Min"], mins, oos, false, meta));
        cols.push(Col([g.stage, section, p3, lim, "Max"], maxs, oos, false, meta));
        cols.push(Col([g.stage, section, p3, lim, "Avg"], avgs, oos, false, meta));
      }
    });
    var used = batches.filter(function (b) {
      return cols.some(function (c) { return !c.text && c.values[b] !== undefined && c.values[b] !== null; });
    });
    return [cols, used];
  }

  function bprDetails(rows) {
    var out = [], seen = {};
    rows.forEach(function (r) {
      if (s(r.APQR_SHEET) !== "Packing") return;
      var key = s(r.BATCH_NO) + "\u0001" + (s(r.SFO_ORDER_NO) || s(r.STAGE));
      if (!(key in seen)) {
        seen[key] = out.length;
        out.push([out.length + 1, s(r.BATCH_NO), s(r.SFO_ORDER_NO), s(r.SFO_MATERIAL), cleanStage(s(r.STAGE)), s(r.STAGE_START),
          s(r.STAGE_END), s(r.PRODUCTION_UNIT_ID) || s(r.PRODUCTION_UNIT), {}, "", "", ""]);
      }
      s(r.EQUIPMENT_IDS).split(",").forEach(function (x) { if (x.trim()) out[seen[key]][8][x.trim()] = 1; });
    });
    out.forEach(function (r) { r[8] = pySort(Object.keys(r[8])).join(", "); });
    return out;
  }

  function build(rows, opt) {
    opt = opt || { merge: true, minutes: true, inspection: false };
    var batches = [], hdr = {};
    rows.forEach(function (r) { if (!(r.BATCH_NO in hdr)) { hdr[r.BATCH_NO] = r; batches.push(r.BATCH_NO); } });
    var g = function (b, c) { return s(hdr[b][c]); };
    var out = [];
    var bmrCols = ["Sr. No.", "Manufacturing Batch No.", "Product Name", "Batch Size", "Mfg. Date", "Exp. Date", "SFG code",
      "Market", "BMR No.", "Master BMR No.", "Shelf Life", "Manufacturing Order", "MO Status", "OOS/OOT/Deviation/Change control", "Remarks"];
    out.push({ name: "BMR detail", kind: "simple", cols: bmrCols, rows: batches.map(function (b, i) {
      var bmr = g(b, "MBR_ID") + (g(b, "MBR_VERSION") ? "-" + g(b, "MBR_VERSION") : "");
      return [i + 1, b, g(b, "PRODUCT_NAME"), (g(b, "BATCH_SIZE") + " " + g(b, "BATCH_SIZE_UOM")).trim(), g(b, "MFG_MONTH"),
        g(b, "EXP_MONTH"), g(b, "MATERIAL_CODE"), "", bmr, "", (g(b, "SHELF_LIFE") + " " + g(b, "SHELF_LIFE_UOM")).trim(),
        g(b, "MANUFACTURING_ORDER"), g(b, "MO_STATUS"), "", ""];
    }) });
    function grid(name, res) {
      if (!res) return;
      res[0].forEach(function (c) { c.meta.sheet = name; });
      out.push({ name: name, kind: "grid", cols: res[0], batches: res[1] });
    }
    SHEETS.filter(function (x) { return x !== "Inspection" || opt.inspection; }).forEach(function (sh) {
      grid(sh, makeSheet(rows, batches,
        function (r) { return s(r.APQR_SHEET) === sh && (s(r.PARAMETER_CLASS) === "Process Parameter" || s(r.PARAMETER_CLASS) === "IPQC (multi-value)"); },
        function (r) {
          var ipc = IPC_RE.test(s(r.STEP)) || s(r.PARAMETER_CLASS) === "IPQC (multi-value)";
          var base = sh + "|" + cleanStage(r.STAGE) + "|" + s(r.APQR_PARAMETER).toUpperCase();
          return opt.merge && ipc ? base + "|" + phaseOf(r.STEP_CODE_PATH) : base + "|" + s(r.STEP_CODE_PATH);
        },
        function (r) { return s(r.STEP) || s(r.STEP_PATH).split(" > ").slice(1).join(" > "); },
        function (r) { return s(r.APQR_PARAMETER); }, "auto", opt));
    });
    var yld = function (r) { return s(r.PARAMETER_CLASS) === "Yield & Rejection" && /actual yield|accountab|reject|%.*yield/i.test(s(r.APQR_PARAMETER)); };
    var isPack = function (r) { return s(r.APQR_SHEET) === "Packing"; };
    grid("Yield data", makeSheet(rows, batches, function (r) { return yld(r) && !isPack(r); },
      function (r) { return cleanStage(r.STAGE) + "|" + s(r.APQR_PARAMETER).toUpperCase(); },
      function (r) { return s(r.STEP) || "Yield"; }, function (r) { return s(r.APQR_PARAMETER); }, "reading", opt));
    var isT = function (r) { return /temp/i.test(s(r.APQR_PARAMETER)); };
    grid("Temp. & RH", makeSheet(rows, batches,
      function (r) { return s(r.PARAMETER_CLASS) === "Environment (Temp/RH)" && /temp|humid|\brh\b/i.test(s(r.APQR_PARAMETER)); },
      function (r) { return cleanStage(r.STAGE) + "|" + (isT(r) ? "T" : "RH"); },
      function () { return "Line clearance"; }, function (r) { return isT(r) ? "Temp" : "RH"; }, "reading", opt));
    grid("IP/FP-Analytical", makeSheet(rows, batches, function (r) { return s(r.PARAMETER_CLASS) === "CQA (analytical)"; },
      function (r) { return "Analytical|" + cleanStage(r.STAGE) + "|" + s(r.APQR_PARAMETER).toUpperCase(); },
      function (r) { return s(r.STEP) || "Result"; }, function (r) { return s(r.APQR_PARAMETER); }, "reading", opt));
    var bpr = bprDetails(rows);
    if (bpr.length) out.push({ name: "BPR Details", kind: "simple", cols: BPR_COLS, rows: bpr });
    grid("Packing Yield", makeSheet(rows, batches, function (r) { return yld(r) && isPack(r); },
      function (r) { return cleanStage(r.STAGE) + "|" + s(r.APQR_PARAMETER).toUpperCase(); },
      function (r) { return s(r.STEP) || "Yield"; }, function (r) { return s(r.APQR_PARAMETER); }, "auto", opt));
    var oosRows = rows.filter(function (r) { return /^OUT/.test(s(r.LIMIT_CHECK).toUpperCase()); }).map(function (r) {
      return ["BATCH_NO", "APQR_SHEET", "STAGE", "STEP_PATH", "APQR_PARAMETER", "UOM", "SPEC_MIN", "SPEC_MAX", "READING_MIN",
        "READING_MAX", "INDIVIDUAL_MIN", "INDIVIDUAL_MAX", "SUPERSEDED_VALUES"].map(function (c) { return r[c] === undefined ? "" : r[c]; });
    });
    out.push({ name: "Out of limit", kind: "simple", cols: ["Batch No.", "Sheet", "Stage", "Step", "Parameter", "UOM", "Spec min",
      "Spec max", "Reading min", "Reading max", "Individual min", "Individual max", "Superseded values"], rows: oosRows });
    return { sheets: out, batches: batches };
  }

  /* ---------------------------------------------------------------- statistics (= stats.py) */
  function fsum(x) {                              // Neumaier compensated sum = Python 3.12+ sum() of floats
    var s1 = 0, c = 0;
    for (var i = 0; i < x.length; i++) {
      var t = s1 + x[i];
      c += Math.abs(s1) >= Math.abs(x[i]) ? (s1 - t) + x[i] : (x[i] - t) + s1;
      s1 = t;
    }
    return s1 + c;
  }
  function mean(x) { return x.length ? fsum(x) / x.length : null; }
  function sd(x) {
    if (x.length < 2) return null;
    var m = mean(x), v = Math.sqrt(fsum(x.map(function (a) { return (a - m) * (a - m); })) / (x.length - 1));
    return v < 1e-12 * Math.max(1, Math.abs(m)) ? 0 : v;           // constant values: float noise is not variation
  }
  function imr(x) {
    if (!x.length) return null;
    var cl = mean(x), mr = [];
    for (var i = 1; i < x.length; i++) mr.push(Math.abs(x[i] - x[i - 1]));
    var mrbar = mr.length ? mean(mr) : null, sw = mrbar ? mrbar / 1.128 : null;
    return { cl: cl, sigma: sw, ucl: sw ? cl + 3 * sw : null, lcl: sw ? cl - 3 * sw : null, mrbar: mrbar };
  }
  function capability(x, lsl, usl, sw) {
    var n = x.length, m = mean(x), s1 = sd(x);
    var out = { n: n, mean: m, sd: s1, cp: null, cpk: null, pp: null, ppk: null };
    if (n < 2 || (lsl === null && usl === null)) return out;
    function idx(sig) {
      if (!sig) return [null, null];
      var c = lsl !== null && usl !== null ? (usl - lsl) / (6 * sig) : null;
      var parts = [usl !== null ? (usl - m) / (3 * sig) : null, lsl !== null ? (m - lsl) / (3 * sig) : null].filter(function (v) { return v !== null; });
      return [c, Math.min.apply(null, parts)];
    }
    var a = idx(sw), b = idx(s1);
    out.cp = a[0]; out.cpk = a[1]; out.pp = b[0]; out.ppk = b[1];
    return out;
  }
  function nelson(x, cl, sigma, rules) {
    var out = [], n = x.length;
    if (!sigma || !n) return out;
    var z = x.map(function (v) { return (v - cl) / sigma; }), side = z.map(function (v) { return v > 0 ? 1 : v < 0 ? -1 : 0; });
    function add(rule, pts) {
      var last = out[out.length - 1];
      if (last && last.rule === rule && pts[0] <= last.points[last.points.length - 1] + 1) {
        var set = {}; last.points.concat(pts).forEach(function (p) { set[p] = 1; });
        last.points = Object.keys(set).map(Number).sort(function (a, b) { return a - b; });
      } else out.push({ rule: rule, text: NELSON[rule], points: pts.slice() });
    }
    function range(a, b) { var r = []; for (var i = a; i < b; i++) r.push(i); return r; }
    (rules || [1, 2, 3]).slice().sort().forEach(function (rule) {
      var i;
      if (rule === 1) for (i = 0; i < n; i++) { if (Math.abs(z[i]) > 3) add(1, [i]); }
      if (rule === 2) for (i = 8; i < n; i++) {
        var w = side.slice(i - 8, i + 1);
        if (w.every(function (v) { return v === 1; }) || w.every(function (v) { return v === -1; })) add(2, range(i - 8, i + 1));
      }
      if (rule === 3) for (i = 5; i < n; i++) {
        var d = []; for (var j = 0; j < 5; j++) d.push(x[i - 5 + j + 1] - x[i - 5 + j]);
        if (d.every(function (v) { return v > 0; }) || d.every(function (v) { return v < 0; })) add(3, range(i - 5, i + 1));
      }
    });
    return out;
  }

  /* ---------------------------------------------------------------- criticality (= adigipqr/cpp.py) */
  var CPP = "CPP", CPP_TYPICAL = "CPP (typical - confirm)", CQA = "CQA (IPQC)";
  var CPP_RULES = [
    ["", "uniformity|weight variation|average (weight|mass)|av\\w*\\.? (weight|mass)|(mass|weight) of \\d+|group weight|net weight|fill weight", "CQA (IPQC)", "tablet/capsule weight"],
    ["", "hardness|crushing|breaking", "CQA (IPQC)", "hardness"],
    ["", "thickness|length|width|diameter", "CQA (IPQC)", "dimensions"],
    ["", "disintegrat|\\bdt\\b", "CQA (IPQC)", "disintegration"],
    ["", "friab", "CQA (IPQC)", "friability"],
    ["", "weight gain", "CQA (IPQC)", "coating weight gain"],
    ["", "\\blod\\b|loss on drying|moisture|water content", "CQA (IPQC)", "drying end point (LOD)"],
    ["", "assay|dissolution|content uniformity|related substance|impurit", "CQA (IPQC)", "analytical CQA"],
    ["", "bulk density|tapped density|particle size|sieve analysis|\\bpsd\\b", "CQA (IPQC)", "granule properties"],
    ["", "binder.*(add|rate|time|qty|quantity)|(water|solution|solvent).*(add|rate|qty|quantity)|addition (time|rate)", "CPP (typical - confirm)", "binder / solvent addition"],
    ["", "wet mix|kneading|massing", "CPP (typical - confirm)", "wet massing"],
    ["granul|rmg|mixer", "(impeller|chopper).*(amp|current|power|kw|load)|power consumption|end ?point", "CPP (typical - confirm)", "granulation end point"],
    ["granul|rmg|mixer", "impeller|chopper", "CPP (typical - confirm)", "impeller / chopper speed"],
    ["dry|fbd|fbe|fluid", "inlet.*temp", "CPP (typical - confirm)", "drying inlet air temperature"],
    ["dry|fbd|fbe|fluid", "(product|bed|outlet|exhaust).*temp", "CPP (typical - confirm)", "product temperature during drying"],
    ["dry|fbd|fbe|fluid", "drying time|^time", "CPP (typical - confirm)", "drying time"],
    ["roll|compact|chilson", "hydraulic|roll\\w*.*(pressure|force)", "CPP (typical - confirm)", "roller pressure"],
    ["roll|compact|chilson", "gap", "CPP (typical - confirm)", "roller gap"],
    ["roll|compact|chilson", "roll\\w*.*(speed|rpm)|roller unit", "CPP (typical - confirm)", "roller speed"],
    ["roll|compact|chilson", "feed|cvf|screw", "CPP (typical - confirm)", "feed screw speed"],
    ["blend|lubric|\\bbin\\b|mix", "time|revolution", "CPP (typical - confirm)", "blending time"],
    ["blend|lubric|\\bbin\\b|mix", "speed|rpm", "CPP (typical - confirm)", "blender speed"],
    ["compress|tablet", "pre.?compression", "CPP (typical - confirm)", "pre-compression force"],
    ["compress|tablet", "compression force|main.*force|\\bmcf\\b", "CPP (typical - confirm)", "main compression force"],
    ["compress|tablet", "turret|press speed|machine speed|tablets per", "CPP (typical - confirm)", "press speed"],
    ["compress|tablet", "feeder|paddle", "CPP (typical - confirm)", "feeder speed"],
    ["compress|tablet", "fill(ing)? depth|depth of fill", "CPP (typical - confirm)", "fill depth"],
    ["coat", "inlet.*air.*(flow|volume|cfm)|air ?flow", "CPP (typical - confirm)", "inlet air flow"],
    ["coat", "inlet.*temp", "CPP (typical - confirm)", "inlet air temperature"],
    ["coat", "(bed|product).*temp", "CPP (typical - confirm)", "product (bed) temperature"],
    ["coat", "exhaust.*temp", "CPP (typical - confirm)", "exhaust temperature (product temperature proxy)"],
    ["coat", "spray rate|peristaltic|pump", "CPP (typical - confirm)", "spray rate"],
    ["coat", "atomi[sz]|pattern|spray.*air.*pressure", "CPP (typical - confirm)", "atomisation / pattern air"],
    ["coat", "pan|drum", "CPP (typical - confirm)", "pan speed"],
    ["coat", "distance|gun to bed|bed to gun", "CPP (typical - confirm)", "gun-to-bed distance"],
    ["capsul|encapsul", "tamp|dosing|dosator|machine speed|capsules per", "CPP (typical - confirm)", "capsule filling settings"],
    ["pack|blister|strip|bottl|alu", "seal\\w*.*(temp|pressure|time)|dwell|forming.*temp|induction", "CPP (typical - confirm)", "sealing temperature / pressure"]
  ];
  var CPP_RX = CPP_RULES.map(function (r) { return [r[0] ? new RegExp(r[0], "i") : null, new RegExp(r[1], "i"), r[2], r[3]]; });
  var NON_CRITICAL = ["", "0", "false", "low", "minor", "n", "nan", "no", "non critical", "non-critical", "noncritical", "none", "normal", "not critical", "null", "uncritical"];
  var CRITICAL_WORD = /crit|cpp|cqa|high|major|key/i, NOT_WORD = /\b(non|not|un)/i;
  var CPP_NOTE = "Criticality: from QA's CPP list when given, else the PAS-X MBR criticality flag, else what is typical for the unit operation (oral solid dose risk assessment, ICH Q8). \"typical - confirm\" must be confirmed against the product's CPP list / process validation report.";
  function cnorm(t) { return String(t === null || t === undefined ? "" : t).toLowerCase().replace(/[^a-z0-9%.]+/g, " ").trim().split(/\s+/).filter(Boolean).join(" "); }
  function parseCppList(text) {
    var out = [];
    String(text || "").split(/\r\n|\r|\n/).forEach(function (line) {
      line = line.split("#")[0].trim();
      if (!line) return;
      var i = line.lastIndexOf(":");
      out.push([line, cnorm(i < 0 ? "" : line.slice(0, i)), cnorm(i < 0 ? line : line.slice(i + 1))]);
    });
    return out;
  }
  function cppBaseline(params) {
    var vals = params.map(function (p) { return cnorm(p.critical); });
    var coded = vals.filter(function (v) { return NON_CRITICAL.indexOf(v) < 0; });
    if (!coded.length) return [false, null];
    if (coded.length < vals.length) return [true, null];
    var n = {}, order = [], best = null;                  // = Counter.most_common: ties go to the first seen
    coded.forEach(function (v) { if (!n[v]) { n[v] = 0; order.push(v); } n[v]++; });
    order.forEach(function (v) { if (best === null || n[v] > n[best]) best = v; });
    return [true, best];
  }
  function mbrFlag(value, baseline) {
    var v = cnorm(value);
    if (NON_CRITICAL.indexOf(v) >= 0) return false;
    if (CRITICAL_WORD.test(v)) return !NOT_WORD.test(v);
    return v !== baseline;
  }
  function classify(params, cppList) {
    var entries = Array.isArray(cppList) ? cppList : parseCppList(cppList), base = cppBaseline(params);
    params.forEach(function (p) {
      var where = [p.sheet, p.stage, p.section].map(function (v) { return v === null || v === undefined ? "" : String(v); }).join(" ");
      var name = String(p.param || ""), cls = p.cls || "";
      var ipqc = cls === "IPQC (multi-value)" || cls === "CQA (analytical)", other = cls === "Environment (Temp/RH)" || cls === "Yield & Rejection";
      var crit = "", basis = "";
      for (var i = 0; i < entries.length && !crit; i++) {
        var e = entries[i];
        if (e[2] && cnorm(name).indexOf(e[2]) >= 0 && (!e[1] || cnorm(where).indexOf(e[1]) >= 0)) { crit = ipqc ? CQA : CPP; basis = 'QA CPP list: "' + e[0] + '"'; }
      }
      if (!crit && base[0] && mbrFlag(p.critical, base[1])) { crit = ipqc ? CQA : CPP; basis = "PAS-X MBR criticality = " + p.critical; }
      if (!crit && !other) {
        for (var j = 0; j < CPP_RX.length; j++) {
          var r = CPP_RX[j];
          if (r[1].test(name) && (!r[0] || r[0].test(where))) {
            crit = r[2]; basis = (r[2] === CQA ? "in-process quality attribute: " : "typical for the unit operation: ") + r[3]; break;
          }
        }
      }
      if (!crit && ipqc) { crit = CQA; basis = "in-process check in PAS-X"; }
      p.crit = crit; p.crit_basis = basis;
    });
    return params;
  }

  /* one entry per trended parameter (= extras.parameters) */
  function parameters(model, cppList) {
    var res = [], seen = {};
    model.sheets.forEach(function (item) {
      if (item.kind !== "grid") return;
      item.cols.forEach(function (c) {
        if (c.text || (c.r[4] !== "Value" && c.r[4] !== "Avg") || seen[c.meta.key]) return;
        seen[c.meta.key] = 1;
        var m = c.meta, order = model.batches.filter(function (b) { return c.values[b] !== undefined && c.values[b] !== null; });
        var x = order.map(function (b) { return c.values[b]; }), ch = x.length ? imr(x) : null;
        var cap = capability(x, m.lo, m.hi, ch ? ch.sigma : null);
        res.push({ sheet: item.name, stage: m.stage, section: m.section, param: m.param, uom: m.uom, lo: m.lo, hi: m.hi,
          batches: order, values: x, cap: cap, chart: ch, out: order.filter(function (b) { return c.oos[b]; }),
          nelson: ch ? nelson(x, ch.cl, ch.sigma, [1, 2, 3]) : [], label: c.r[4] === "Value" ? "Value" : "Batch average",
          batchMin: m.batch_min, batchMax: m.batch_max, cls: m.cls, critical: m.critical });
      });
    });
    return classify(res, cppList);
  }
  function note(p) {
    var n = p.cap.n, cpk = p.cap.cpk;
    if (p.lo === null && p.hi === null) return "no limit in MBR";
    if (n < 6) return "too few batches for Cpk";
    if (cpk === null) return "no variation";
    return cpk >= 1.33 ? "capable" : cpk >= 1.0 ? "marginal" : "not capable";
  }

  /* ---------------------------------------------------------------- Excel (= pivot.write + extras) */
  var HEAD_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDDEBF0" } };
  var OOS_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8CBAD" } };
  var RED = OOS_FILL, AMBER = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFCE4B6" } };
  var CRIT_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC6E8D5" } }, CRIT_TYP_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE6F4EC" } };
  var THIN = { style: "thin", color: { argb: "FFB7C4CC" } }, BORDER = { top: THIN, left: THIN, bottom: THIN, right: THIN };
  var CENTER = { wrapText: true, horizontal: "center", vertical: "middle" };

  function asCell(v) {                            // = pivot.as_cell
    var t = s(v);
    if (/^\d{8,}$/.test(t)) return t;
    var n = num(v);
    return n !== null && /^-?[\d.]+$/.test(t) ? n : (v === null || v === undefined ? "" : v);
  }
  function colLetter(n) { var s2 = ""; while (n > 0) { var m = (n - 1) % 26; s2 = String.fromCharCode(65 + m) + s2; n = Math.floor((n - 1) / 26); } return s2; }

  function headerRuns(cols, level) {
    function same(a, b) {
      if (level === 4) return false;
      if (level <= 1) { for (var p = 0; p <= level; p++) if (a.r[p] !== b.r[p]) return false; return true; }
      return !a.text && a.r[0] === b.r[0] && a.r[1] === b.r[1] && a.r[2] === b.r[2] && a.r[3] === b.r[3];
    }
    var runs = [], i = 0;
    while (i < cols.length) { var j = i + 1; while (j < cols.length && same(cols[i], cols[j])) j++; runs.push([i, j - i]); i = j; }
    return runs;
  }

  function writeQaSheets(wb, model) {
    model.sheets.forEach(function (item) {
      var ws = wb.addWorksheet(item.name.substring(0, 31));
      if (item.kind === "simple") {
        ws.addRow(item.cols);
        item.rows.forEach(function (r) { ws.addRow(r.map(asCell)); });
        ws.getRow(1).eachCell(function (c) { c.font = { bold: true }; c.fill = HEAD_FILL; c.border = BORDER; c.alignment = CENTER; });
        for (var i = 1; i <= item.cols.length; i++) ws.getColumn(i).width = 16;
        ws.views = [{ state: "frozen", xSplit: 2, ySplit: 1 }];
        return;
      }
      var cols = item.cols, first = 3;
      ws.getCell(1, 1).value = "Sr. No."; ws.getCell(1, 2).value = "Manufacturing Batch No.";
      ws.mergeCells(1, 1, 5, 1); ws.mergeCells(1, 2, 5, 2);
      for (var lv = 0; lv < 5; lv++) {
        headerRuns(cols, lv).forEach(function (run) {
          var c0 = first + run[0], txt = cols[run[0]].r[lv];
          if (txt) ws.getCell(lv + 1, c0).value = txt;
          if (run[1] > 1) ws.mergeCells(lv + 1, c0, lv + 1, c0 + run[1] - 1);
        });
      }
      item.batches.forEach(function (b, i) {
        var r = 6 + i;
        ws.getCell(r, 1).value = i + 1; ws.getCell(r, 1).border = BORDER;
        ws.getCell(r, 2).value = b; ws.getCell(r, 2).border = BORDER;
        cols.forEach(function (c, ci) {
          var v = c.values[b], cell = ws.getCell(r, first + ci);
          if (c.text) { if (v) cell.value = v; }
          else if (v !== undefined && v !== null && isFinite(v)) { cell.value = pyRound(v, 4); if (c.oos[b]) cell.fill = OOS_FILL; }
          cell.border = BORDER;
        });
      });
      for (var rr = 1; rr <= 5; rr++) for (var cc = 1; cc < first + cols.length; cc++) {
        var cell = ws.getCell(rr, cc);
        cell.fill = HEAD_FILL; cell.border = BORDER; cell.alignment = CENTER; cell.font = { bold: rr !== 4, size: rr === 4 ? 9 : 10 };
      }
      ws.getRow(3).height = 42; ws.getRow(4).height = 30;
      ws.getColumn(1).width = 7; ws.getColumn(2).width = 14;
      cols.forEach(function (c, k) { ws.getColumn(first + k).width = c.text ? 16 : 11; });
      ws.views = [{ state: "frozen", xSplit: 2, ySplit: 5 }];
    });
  }

  function r4(v) { return typeof v === "number" && !Number.isInteger(v) ? pyRound(v, 4) : v; }

  function writeSummary(wb, params) {
    var ws = wb.addWorksheet("Summary");
    var head = ["Sheet", "Stage", "Step", "Parameter", "Criticality", "Criticality basis", "UOM", "LSL", "USL", "Batches",
      "Min", "Max", "Mean", "SD", "Cpk", "Ppk", "Capability", "Batches outside limit", "Trend signals", "Plotted value"];
    ws.addRow(head);
    params.forEach(function (p) {
      var c = p.cap;
      var sig = p.nelson.map(function (f) {
        return "R" + f.rule + " " + NELSON[f.rule].toLowerCase() + ": " + f.points.map(function (i) { return p.batches[i]; }).join(", ");
      }).join("; ");
      var row = [p.sheet, p.stage, p.section, p.param, p.crit, p.crit_basis, p.uom, p.lo, p.hi, c.n,
        p.values.length ? Math.min.apply(null, p.values) : null, p.values.length ? Math.max.apply(null, p.values) : null,
        c.mean, c.sd, c.cpk, c.ppk, note(p), p.out.join(", "), sig, p.label].map(function (v) { return v === null ? undefined : r4(v); });
      var added = ws.addRow(row);
      if (p.crit) added.getCell(5).fill = p.crit === CPP ? CRIT_FILL : CRIT_TYP_FILL;
      if (p.out.length) added.getCell(18).fill = RED;
      if (c.cpk !== null && c.n >= 6) { if (c.cpk < 1.0) added.getCell(15).fill = RED; else if (c.cpk < 1.33) added.getCell(15).fill = AMBER; }
      if (sig) added.getCell(19).fill = AMBER;
    });
    ws.getRow(1).eachCell(function (c) { c.font = { bold: true }; c.fill = HEAD_FILL; c.border = BORDER; c.alignment = { wrapText: true, vertical: "middle" }; });
    [12, 22, 26, 28, 20, 34, 8, 9, 9, 8, 9, 9, 9, 9, 7, 7, 18, 22, 40, 13].forEach(function (w, i) { ws.getColumn(i + 1).width = w; });
    ws.views = [{ state: "frozen", xSplit: 4, ySplit: 1 }];
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: Math.max(1, params.length + 1), column: head.length } };
    var nr = params.length + 3;
    ws.getCell(nr, 1).value = "Cpk = min(USL-mean, mean-LSL) / 3 sigma, sigma = average moving range / 1.128 (batch to batch). " +
      "Ppk uses the overall standard deviation. Below 6 batches the indices are not interpreted. " +
      "Trend signals: R1 point beyond 3 sigma, R2 nine points on one side of the mean, R3 six points steadily rising or falling.";
    ws.getCell(nr, 1).font = { italic: true, size: 9 };
    ws.getCell(nr + 1, 1).value = CPP_NOTE;
    ws.getCell(nr + 1, 1).font = { italic: true, size: 9 };
  }

  /* Trend data sheet + chart specs; the charts themselves are injected into the xlsx afterwards (injectCharts) */
  function writeTrends(wb, params) {
    var plot = params.filter(function (p) { return p.values.length >= 3 && (p.lo !== null || p.hi !== null); }).slice(0, 150);
    if (!plot.length) return null;
    var charts = wb.addWorksheet("Trends"), dat = wb.addWorksheet("Trend data");
    charts.getCell("A1").value = "Batch values against the MBR limits, in manufacturing order. Data: sheet \"Trend data\". " +
      "Points outside a limit are listed in the Summary sheet.";
    charts.getCell("A1").font = { italic: true, size: 9 };
    var specs = [], col = 1;
    plot.forEach(function (p, k) {
      var title = p.param + (p.uom ? " [" + p.uom + "]" : "") + " - " + p.stage;
      dat.getCell(1, col).value = title; dat.getCell(1, col).font = { bold: true };
      ["Batch", p.label, "LSL", "USL"].forEach(function (h, j) { dat.getCell(2, col + j).value = h; dat.getCell(2, col + j).font = { bold: true }; });
      p.batches.forEach(function (b, i) {
        dat.getCell(3 + i, col).value = b;
        dat.getCell(3 + i, col + 1).value = pyRound(p.values[i], 4);
        if (p.lo !== null) dat.getCell(3 + i, col + 2).value = p.lo;
        if (p.hi !== null) dat.getCell(3 + i, col + 3).value = p.hi;
      });
      var pts = p.values.concat([p.lo, p.hi].filter(function (v) { return v !== null; }));
      var lo = Math.min.apply(null, pts), hi = Math.max.apply(null, pts), span = (hi - lo) || Math.abs(hi) * 0.1 || 1;
      specs.push({ title: title.substring(0, 120), uom: p.uom, col: col, last: 2 + p.values.length, n: p.values.length,
        ymin: pyRound(lo - 0.12 * span, 6), ymax: pyRound(hi + 0.12 * span, 6), anchorCol: k % 2 === 0 ? 0 : 9,
        anchorRow: 2 + Math.floor(k / 2) * 15 });
      col += 5;
    });
    return { sheet: charts, specs: specs };
  }

  function esc(t) { return String(t).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
  var A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
  function chartXml(sp) {                        // same structure as openpyxl's LineChart output
    function L(c) { return colLetter(c); }
    function ser(i, c, color, dash, marker) {
      return '<c:ser><c:idx val="' + i + '"/><c:order val="' + i + '"/><c:tx><c:strRef><c:f>\'Trend data\'!' + L(c) + '2</c:f></c:strRef></c:tx>' +
        '<c:spPr><a:ln><a:solidFill><a:srgbClr val="' + color + '"/></a:solidFill><a:prstDash val="' + dash + '"/></a:ln></c:spPr>' +
        (marker ? '<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr><a:solidFill><a:srgbClr val="' + color + '"/></a:solidFill></c:spPr></c:marker>'
          : '<c:marker><c:symbol val="none"/></c:marker>') +
        '<c:cat><c:numRef><c:f>\'Trend data\'!$' + L(sp.col) + '$3:$' + L(sp.col) + '$' + sp.last + '</c:f></c:numRef></c:cat>' +
        '<c:val><c:numRef><c:f>\'Trend data\'!$' + L(c) + '$3:$' + L(c) + '$' + sp.last + '</c:f></c:numRef></c:val><c:smooth val="0"/></c:ser>';
    }
    function rich(t) { return '<c:tx><c:rich><a:bodyPr/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:t>' + esc(t) + '</a:t></a:r></a:p></c:rich></c:tx>'; }
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' + A + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<c:chart><c:title>' + rich(sp.title) + '<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>' +
      ser(0, sp.col + 1, "0E5B63", "solid", true) + ser(1, sp.col + 2, "C0392B", "dash", false) + ser(2, sp.col + 3, "C0392B", "dash", false) +
      '<c:marker val="1"/><c:axId val="10"/><c:axId val="100"/></c:lineChart>' +
      '<c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>' +
      '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="100"/><c:crosses val="autoZero"/>' +
      '<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:tickLblSkip val="' + Math.max(1, Math.floor(sp.n / 12)) + '"/><c:noMultiLvlLbl val="0"/></c:catAx>' +
      '<c:valAx><c:axId val="100"/><c:scaling><c:orientation val="minMax"/><c:max val="' + sp.ymax + '"/><c:min val="' + sp.ymin + '"/></c:scaling>' +
      '<c:delete val="0"/><c:axPos val="l"/>' + (sp.uom ? '<c:title>' + rich(sp.uom) + '<c:overlay val="0"/></c:title>' : '') +
      '<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>' +
      '<c:crossAx val="10"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx></c:plotArea>' +
      '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>';
  }

  async function injectCharts(buffer, sheetIndex, specs) {
    if (!specs || !specs.length) return buffer;
    var zip = await JSZip.loadAsync(buffer);
    var sheetPath = "xl/worksheets/sheet" + sheetIndex + ".xml";
    var xml = await zip.file(sheetPath).async("string");
    var anchors = specs.map(function (sp, i) {
      return '<xdr:oneCellAnchor><xdr:from><xdr:col>' + sp.anchorCol + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + sp.anchorRow +
        '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="5580000" cy="2592000"/><xdr:graphicFrame macro="">' +
        '<xdr:nvGraphicFramePr><xdr:cNvPr id="' + (i + 2) + '" name="Chart ' + (i + 1) + '"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>' +
        '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
        '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId' + (i + 1) + '"/></a:graphicData></a:graphic>' +
        '</xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>';
    }).join("");
    zip.file("xl/drawings/drawing1.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
      A + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' + anchors + '</xdr:wsDr>');
    zip.file("xl/drawings/_rels/drawing1.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      specs.map(function (sp, i) { return '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart' + (i + 1) + '.xml"/>'; }).join("") + '</Relationships>');
    specs.forEach(function (sp, i) { zip.file("xl/charts/chart" + (i + 1) + ".xml", chartXml(sp)); });
    var relsPath = "xl/worksheets/_rels/sheet" + sheetIndex + ".xml.rels", rels = zip.file(relsPath) ? await zip.file(relsPath).async("string")
      : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    rels = rels.replace("</Relationships>", '<Relationship Id="rIdAdigiDrawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>');
    zip.file(relsPath, rels);
    if (!/xmlns:r=/.test(xml.slice(0, 600))) xml = xml.replace("<worksheet ", '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
    var tag = '<drawing r:id="rIdAdigiDrawing1"/>', pos = xml.search(/<(legacyDrawing|tableParts|extLst)\b|<\/worksheet>/);
    xml = xml.slice(0, pos) + tag + xml.slice(pos);
    zip.file(sheetPath, xml);
    var ct = await zip.file("[Content_Types].xml").async("string");
    ct = ct.replace("</Types>", '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' +
      specs.map(function (sp, i) { return '<Override PartName="/xl/charts/chart' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'; }).join("") + "</Types>");
    zip.file("[Content_Types].xml", ct);
    return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  }

  /* Complete workbook: QA sheets, Summary, Trends, Trend data, [Scanned records], MES long data, Report info */
  async function workbook(rows, opts) {
    opts = opts || {};
    var model = build(rows, opts.pivot), params = parameters(model, opts.cppList);
    var wb = new ExcelJS.Workbook();
    wb.creator = "AdigiPQR";
    writeQaSheets(wb, model);
    writeSummary(wb, params);
    var trends = writeTrends(wb, params);
    if (opts.scans && opts.scans.length) writeScans(wb, opts.scans);
    var long = wb.addWorksheet("MES long data"), keys = rows.sourceColumns || Object.keys(rows[0]);
    long.addRow(keys);
    rows.forEach(function (r) { if (!r._scan) long.addRow(keys.map(function (k) { return asCell(r[k]); })); });
    long.getRow(1).font = { bold: true };
    long.views = [{ state: "frozen", ySplit: 1 }];
    var info = wb.addWorksheet("Report info");
    (opts.info || []).forEach(function (r) { var row = info.addRow(r); row.getCell(1).font = { bold: true }; });
    info.getColumn(1).width = 24; info.getColumn(2).width = 100;
    var buf = await wb.xlsx.writeBuffer();
    if (trends) {
      var idx = wb.worksheets.indexOf(trends.sheet) + 1;
      buf = await injectCharts(buf, idx, trends.specs);
    }
    return { buffer: buf, model: model, params: params };
  }

  function writeScans(wb, scans) {
    var ws = wb.addWorksheet("Scanned records");
    ws.addRow(["Batch", "Serial", "Interval", "Side", "Balance", "Start", "End", "n", "Average (mg)", "Min (mg)", "Max (mg)", "SD",
      "LSL", "USL", "Printed result", "Checks", "Source pages", "Accepted by", "Accepted on", "Note"]);
    scans.forEach(function (sc) {
      var v = sc.values, failed = sc.checks.filter(function (c) { return !c.ok; }).map(function (c) { return c.name; });
      var row = ws.addRow([v.batch, v.serial, v.interval, v.side, v.balance, v.start, v.end, v.n, v.avg, v.min, v.max, v.sd, v.lsl, v.usl,
        v.result, failed.length ? "failed: " + failed.join("; ") : "all " + sc.checks.length + " passed", sc.pages.join(","),
        sc.acceptedBy || "", sc.acceptedAt || "", sc.note || ""]);
      if (failed.length) row.getCell(16).fill = AMBER;
    });
    ws.getRow(1).eachCell(function (c) { c.font = { bold: true }; c.fill = HEAD_FILL; c.border = BORDER; });
    ws.views = [{ state: "frozen", xSplit: 1, ySplit: 1 }];
  }

  global.AdigiCore = { VERSION: VERSION, normalise: normalise, build: build, parameters: parameters, workbook: workbook,
    stats: { imr: imr, capability: capability, nelson: nelson, mean: mean, sd: sd }, note: note, s: s, num: num, fmt: fmt,
    pyRound: pyRound, NELSON: NELSON, asCell: asCell, classify: classify, parseCppList: parseCppList,
    CPP_RULES: CPP_RULES, CPP: CPP, CPP_TYPICAL: CPP_TYPICAL, CQA: CQA };
})(typeof window !== "undefined" ? window : this);
