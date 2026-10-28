/* AdigiPQR - check the MES export against QA's manually compiled APQR workbook, cell by cell.
   Port of python/compare_apqr.py (same rules, same statuses; tests/test_compare.py keeps them equal).
   Everything stays in this browser tab. */
(function (global) {
  "use strict";
  var SHEETS = ["Granulation", "Compression", "Coating", "Temp. & RH", "Temp. & Humidity", "Yield data", "BMR detail",
    "Packing Yield", "BPR Details"];
  var BATCH_RE = /^\d{10}$/;
  // (qa_regex, mes_param_regex, stage_regex, step_regex, value_kind) = compare_apqr.RULES
  var RULES = [
    ["^(granulation|compression|coating|packing yield) > (equipment|granulation ?\\()", null, null, null, "E"],
    ["cvf feed", "cvf\\s*feed", "roll|compact", null, "R"],
    ["roller gap", "roll\\w*\\s*gap", "roll|compact", null, "R"],
    ["roller unit", "roller\\s*unit", "roll|compact", null, "R"],
    ["cvf stirrer", "cvf\\s*stir", "roll|compact", null, "R"],
    ["hydraulic", "hydraulic", "roll|compact", null, "R"],
    ["pre-?granulator", "pre\\s*-?\\s*granul", "roll|compact", null, "R"],
    ["fine-?granulator", "fine\\s*-?\\s*granul", "roll|compact", null, "R"],
    ["upper screen", "upper", "roll|compact", null, "R"],
    ["lower screen", "lower", "roll|compact", null, "R"],
    ["pre-?lub.*speed", "rpm|speed", "lub|blend", "prelub|pre-?lub", "R"],
    ["pre-?lub.*time", "mixing time|blending time|time", "lub|blend", "prelub|pre-?lub", "R"],
    ["(?<!pre-)lubrication.*speed", "rpm|speed", "lub|blend", "(?<!pre)lub(?!.*prelub)", "R"],
    ["(?<!pre-)lubrication.*time", "mixing time|blending time|time", "lub|blend", "(?<!pre)lub(?!.*prelub)", "R"],
    ["compression.*uniformity", "uniformity", "compress", null, "I"],
    ["compression.*thickness", "thickness", "compress", null, "I"],
    ["compression.*(hardness|crushing|breaking)", "hardness|breaking|crushing", "compress", null, "I"],
    ["compression.*disintegration", "disintegration", "compress", null, "I"],
    ["compression.*friability", "friab", "compress", null, "R"],
    ["compression.*length", "length", "compress", null, "I"],
    ["compression.*width", "width", "compress", null, "I"],
    ["compression.*average mass", "av\\w*\\.?\\s*mass|average weight|avg\\.? weight", "compress", null, "R"],
    ["mass of 20", "(mass|weight) of 20", "compress", null, "R"],
    ["turret", "turret", "compress", null, "R"],
    ["feeder", "feeder", "compress", null, "R"],
    ["pre- ?compression force", "pre\\s*comp", "compress", null, "R"],
    ["main compression force", "main\\s*comp", "compress", null, "R"],
    ["coating.*uniformity", "uniformity", "coat", null, "I"],
    ["coating.*disintegration", "disintegration", "coat", null, "I"],
    ["coating.*ave\\. weight", "average weight|av\\w*\\.?\\s*mass", "coat", null, "R"],
    ["group weight", "weight of 20|mass of 20", "coat", null, "R"],
    ["preheating.*bed", "bed|product", "coat", "preheat", "R"],
    ["preheating time", "preheat\\w*\\s*time|time", "coat", "preheat", "R"],
    ["spraying gun", "gun", "coat", null, "R"],
    ["distance between bed and gun", "distance|bed to gun|gun to", "coat", null, "R"],
    ["coating parameter.*inlet air temp", "inlet", "coat", "spray", "R"],
    ["coating parameter.*bed", "bed|product", "coat", "spray", "R"],
    ["coating parameter.*exhaust", "exhaust", "coat", "spray", "R"],
    ["pan rpm", "pan", "coat", null, "R"],
    ["peristaltic", "peristaltic|pp\\s*rpm", "coat", "spray", "R"],
    ["atomi", "atomi", "coat", "spray", "R"],
    ["spray rate", "spray\\s*rate", "coat", null, "R"],
    ["coating parameter.*inlet air flow", "inlet air flow|cfm", "coat", "spray", "R"],
    ["coating parameter.*% rh", "rh", "coat", "spray", "R"],
    ["drying time", "drying time", "coat", null, "R"],
    ["cooling time", "cooling time", "coat", null, "R"],
    ["weight gain", "weight gain", "coat", null, "R"],
    ["yield data.*lubrication", "actual yield|% yield", "lub|blend", null, "Y"],
    ["yield data.*compression >", "actual yield|% yield", "compress", null, "Y"],
    ["yield data.*coating >", "actual yield|% yield", "coat", null, "Y"],
    ["packing yield.*parameter > packing >", "yield", "pack", null, "Y"]
  ];
  var RX = {};
  function rx(p) { return RX[p] || (RX[p] = new RegExp(p)); }
  var ENV_STAGE = [["sifting", "sift"], ["blending", "blend|lub"], ["roll comp", "roll|compact"], ["mixing", "mix"],
    ["lubrication", "lub"], ["compression", "compress"], ["coating", "coat"], ["dispensing", "dispens"]];

  /* ---------------------------------------------------------------- 1. flatten QA's workbook */
  function cellValue(cell) {                       // openpyxl data_only semantics
    if (!cell) return null;
    if (cell.type === 1) return null;              // ExcelJS ValueType.Merge: a merged cell that is not the top-left one
    var v = cell.value;
    if (v === undefined || v === null) return null;
    if (v instanceof Date || typeof v !== "object") return v;
    if (v.richText) return v.richText.map(function (t) { return t.text; }).join("");
    if ("result" in v || "formula" in v || "sharedFormula" in v) return v.result === undefined ? null : (v.result && v.result.error ? v.result.error : v.result);
    if ("text" in v) return v.text;
    if ("error" in v) return v.error;
    return String(v);
  }
  function str(v) { return v instanceof Date ? v.toISOString() : String(v); }
  function words(v) { return str(v).split(/\s+/).filter(Boolean).join(" "); }

  function flattenQa(wb, batchPrefix) {
    var rows = [];
    wb.worksheets.forEach(function (ws) {
      var name = ws.name;
      if (SHEETS.indexOf(name) < 0) return;
      var merged = {};
      ((ws.model && ws.model.merges) || []).forEach(function (rng) {
        var p = rng.split(":"), a = ws.getCell(p[0]), b = ws.getCell(p[1] || p[0]), v = cellValue(a);
        for (var r = a.row; r <= b.row; r++) for (var c = a.col; c <= b.col; c++) merged[r + "," + c] = v;
      });
      var raw = function (r, c) { var row = ws.findRow(r); return row ? cellValue(row.findCell(c)) : null; };
      var val = function (r, c) { var k = r + "," + c; return k in merged ? merged[k] : raw(r, c); };
      var first = null, bcol = null;
      for (var r = 1; r < 40 && !first; r++) {
        for (var c = 1; c < 5; c++) {
          var v = raw(r, c);
          if (v !== null && BATCH_RE.test(str(v).trim())) { first = r; bcol = c; break; }
        }
      }
      if (!first) return;
      var maxCol = ws.columnCount, maxRow = ws.rowCount, headers = {};
      for (var hc = 1; hc <= maxCol; hc++) {
        var parts = [];
        for (var hr = 1; hr < first; hr++) {
          var hv = val(hr, hc);
          if (hv !== null) { var s = words(hv); if (!parts.length || parts[parts.length - 1] !== s) parts.push(s); }
        }
        headers[hc] = parts;
      }
      for (var rr = first; rr <= maxRow; rr++) {
        var b = raw(rr, bcol);
        if (b === null || !BATCH_RE.test(str(b).trim())) continue;
        b = str(b).trim();
        if (batchPrefix && b.indexOf(batchPrefix) !== 0) continue;
        var lastParts = null;
        for (var cc = bcol + 1; cc <= maxCol; cc++) {
          var ps = headers[cc];
          if (!ps || !ps.length) continue;
          var qv = raw(rr, cc), leaf = ps[ps.length - 1].toLowerCase().replace(/\.+$/, ""), stat = "VALUE";
          if (leaf === "min" || leaf === "minimum") stat = "MIN";
          else if (leaf === "max" || leaf === "maximum") stat = "MAX";
          else if (leaf === "ave" || leaf === "avg" || leaf === "average") {
            stat = "AVG";
            if (ps.length <= 2 && lastParts) ps = lastParts.slice(0, -1).concat(["Ave"]);   // bare "Ave" column
          }
          if (stat === "MIN" || stat === "MAX") lastParts = ps;
          rows.push({ BATCH_NO: b, QA_SHEET: name, QA_COLUMN: ps.join(" > "), STAT: stat, QA_VALUE: qv });
        }
      }
    });
    return rows;
  }

  /* ---------------------------------------------------------------- 2. compare */
  function toNum(v) {
    if (v === null || v === undefined) return null;
    if (v instanceof Date) {                      // a pure time cell (00:50) = minutes:seconds as QA types it
      var serial = (v.getTime() - Date.UTC(1899, 11, 30)) / 86400000;
      if (serial >= 0 && serial < 1) return v.getUTCHours() * 60 + v.getUTCMinutes() + v.getUTCSeconds() / 60;
      return null;
    }
    if (typeof v === "number") return isFinite(v) ? v : null;
    if (typeof v === "boolean") return null;        // float(str(True)) fails in Python too
    var t = String(v).replace(/,/g, "").trim();
    if (!/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(t)) return null;
    return Number(t);
  }
  function num(v) { var n = toNum(v); return n === null ? null : n; }
  function blank(v) { return v === null || v === undefined || (typeof v === "number" && isNaN(v)) || ["", "NA", "IA", "None", "nan"].indexOf(String(v).trim()) >= 0; }
  function pick(cands, kind, stat) {
    var col = { R: { MIN: "READING_MIN", MAX: "READING_MAX", AVG: "READING_AVG", VALUE: "READING_MAX" },
      I: { MIN: "INDIVIDUAL_MIN", MAX: "INDIVIDUAL_MAX", AVG: "INDIVIDUAL_AVG", VALUE: "INDIVIDUAL_AVG" },
      Y: { MIN: "READING_MIN", MAX: "READING_MAX", AVG: "READING_AVG", VALUE: "READING_MAX" } }[kind][stat];
    if (kind === "I" && cands.every(function (r) { return num(r[col]) === null; })) col = col.replace("INDIVIDUAL", "READING");
    return col;
  }
  function uniq(a) { var o = []; a.forEach(function (x) { if (o.indexOf(x) < 0) o.push(x); }); return o; }
  function r4(x) { return Math.round(x * 1e4) / 1e4; }
  function close(a, b) { return Math.abs(a - b) <= Math.max(0.011, Math.abs(b) * 0.005); }

  function prepareMes(rows) {
    return rows.map(function (r) {
      var g = function (k) { var v = r[k]; return v === null || v === undefined ? "" : String(v); };
      return { row: r, BATCH_NO: g("BATCH_NO").replace(/\.0$/, ""), STAGE: g("STAGE"), STEP_PATH: g("STEP_PATH"),
        APQR_PARAMETER: g("APQR_PARAMETER"), PARAMETER_CLASS: g("PARAMETER_CLASS"), EQUIPMENT_IDS: g("EQUIPMENT_IDS"),
        PRODUCTION_UNIT_ID: g("PRODUCTION_UNIT_ID"), _stage: g("STAGE").toLowerCase(),
        _step: (g("STEP_PATH") + " " + g("STEP_CODE_PATH")).toLowerCase(), _param: g("APQR_PARAMETER").toLowerCase() };
    });
  }

  function compare(qa, mesRows) {
    var mes = prepareMes(mesRows), byBatch = {};
    mes.forEach(function (m) { (byBatch[m.BATCH_NO] = byBatch[m.BATCH_NO] || []).push(m); });
    return qa.map(function (q) {
      var rec = { BATCH_NO: q.BATCH_NO, QA_SHEET: q.QA_SHEET, QA_COLUMN: q.QA_COLUMN, STAT: q.STAT, QA_VALUE: q.QA_VALUE,
        MES_VALUE: null, STATUS: null, MES_STAGE: null, MES_STEP: null, MES_PARAMETER: null, CANDIDATES: 0 };
      var qv = toNum(q.QA_VALUE);
      if (blank(q.QA_VALUE)) { rec.STATUS = "QA_BLANK"; return rec; }
      var b = byBatch[q.BATCH_NO] || [];
      if (!b.length) { rec.STATUS = "BATCH_NOT_IN_EXTRACT"; return rec; }
      var col = q.QA_COLUMN.toLowerCase(), full = (q.QA_SHEET + " > " + q.QA_COLUMN).toLowerCase(), c, rule;
      if (q.QA_SHEET.indexOf("Temp") === 0 || /^(temp|rh)\b/.test(col)) {
        var segs = col.split(">"), pkw = segs[segs.length - 1].indexOf("temp") >= 0 ? "temp" : "humid|rh";
        var skw = null;
        for (var i = 0; i < ENV_STAGE.length; i++) if (col.indexOf(ENV_STAGE[i][0]) >= 0) { skw = ENV_STAGE[i][1]; break; }
        if (skw === null && col.indexOf("pack") >= 0) skw = "pack";
        c = b.filter(function (m) { return m.PARAMETER_CLASS.indexOf("Environment (Temp") === 0 && rx(pkw).test(m._param); });
        if (skw) c = c.filter(function (m) { return rx(skw).test(m._stage); });
        rule = [null, "R"];
      } else {
        var found = null;
        for (var k = 0; k < RULES.length; k++) if (rx(RULES[k][0]).test(full)) { found = RULES[k]; break; }
        if (!found) { rec.STATUS = "NO_MAPPING_RULE"; return rec; }
        if (found[4] === "E") {
          var eb = b;
          var stages = [["dispens", "dispens"], ["lubrication", "lub|blend"], ["compression", "compress"], ["coating", "coat"],
            ["granulation", "roll|compact|granul"], ["packing", "pack|blister"]];
          for (var s = 0; s < stages.length; s++) {
            if (full.indexOf(stages[s][0]) >= 0 || (stages[s][0] === "packing" && q.QA_SHEET === "Packing Yield")) {   // sheet name counts too
              eb = b.filter(function (m) { return rx(stages[s][1]).test(m._stage); }); break;
            }
          }
          [["amlo", "amlo"], ["valsartan", "vals"]].forEach(function (t) {
            if (col.indexOf(t[0]) >= 0 && eb.some(function (m) { return rx(t[1]).test(m._stage); })) eb = eb.filter(function (m) { return rx(t[1]).test(m._stage); });
          });
          if (!eb.length) { rec.STATUS = "NOT_IN_MES"; return rec; }
          var eq = (uniq(eb.map(function (m) { return m.EQUIPMENT_IDS; })).join(" ") + " " + uniq(eb.map(function (m) { return m.PRODUCTION_UNIT_ID; })).join(" ")).trim();
          rec.MES_VALUE = eq.slice(0, 200);
          rec.STATUS = eq.toUpperCase().indexOf(str(q.QA_VALUE).trim().toUpperCase()) >= 0 ? "MATCH" : "MISMATCH";
          return rec;
        }
        c = b.filter(function (m) { return rx(found[1]).test(m._param); });
        if (found[2]) c = c.filter(function (m) { return rx(found[2]).test(m._stage) || rx(found[2]).test(m._step); });
        if (found[3]) c = c.filter(function (m) { return rx(found[3]).test(m._step); });
        if (found[4] === "Y") c = c.filter(function (m) { return m.PARAMETER_CLASS === "Yield & Rejection"; });
        rule = [null, found[4]];
      }
      // part A / part B, lot I / lot II hints: if QA names a part that MES rows do not carry, do not guess
      var unresolved = false;
      [["amlo", "amlo"], ["valsartan", "vals"], ["lot-i ", "lot.?i\\b"], ["lot-ii", "lot.?ii"]].forEach(function (t) {
        if (col.indexOf(t[0]) >= 0 && c.length) {
          var cc = c.filter(function (m) { return rx(t[1]).test(m._stage) || rx(t[1]).test(m._step); });
          if (cc.length) c = cc; else unresolved = true;
        }
      });
      rec.CANDIDATES = c.length;
      if (!c.length) { rec.STATUS = "NOT_IN_MES"; return rec; }
      var colname = pick(c.map(function (m) { return m.row; }), rule[1], q.STAT);
      var withVal = c.filter(function (m) { return num(m.row[colname]) !== null; });
      if (!withVal.length) { rec.STATUS = "NOT_IN_MES"; return rec; }
      var vals = withVal.map(function (m) { return num(m.row[colname]); });
      var mv = q.STAT === "MIN" ? Math.min.apply(null, vals) : q.STAT === "MAX" ? Math.max.apply(null, vals) : vals[0];
      var top = withVal[0];
      rec.MES_VALUE = mv; rec.MES_STAGE = top.STAGE; rec.MES_STEP = top.STEP_PATH; rec.MES_PARAMETER = top.APQR_PARAMETER;
      if (qv === null) rec.STATUS = "QA_NOT_NUMERIC";
      else if (close(mv, qv)) rec.STATUS = c.length === 1 || uniq(vals.map(r4)).length === 1 ? "MATCH" : "MATCH (several MES rows)";
      else if ([1 / 60, 60].some(function (f) { return close(mv * f, qv); })) rec.STATUS = "MATCH (unit sec/min)";
      else rec.STATUS = c.length === 1 && !unresolved ? "MISMATCH" : "AMBIGUOUS";
      return rec;
    });
  }

  /* QA workbook (ArrayBuffer) + MES rows -> { results, missing (QA batches not in the export), qaBatches } */
  async function run(qaBuffer, mesRows, batchPrefix) {
    var wb = new ExcelJS.Workbook();
    await wb.xlsx.load(qaBuffer);
    var qa = flattenQa(wb, batchPrefix || "");
    var inMes = {}; mesRows.forEach(function (r) { inMes[String(r.BATCH_NO).replace(/\.0$/, "")] = 1; });
    var qaBatches = uniq(qa.map(function (q) { return q.BATCH_NO; }));
    var missing = qaBatches.filter(function (b) { return !inMes[b]; }).sort();
    var results = compare(qa.filter(function (q) { return inMes[q.BATCH_NO]; }), mesRows);
    return { results: results, missing: missing, qaBatches: qaBatches, qaValues: qa.length };
  }

  /* Summary sheet (QA sheet x status counts) + Detail sheet, like compare_apqr.py */
  async function workbook(res, info) {
    var wb = new ExcelJS.Workbook(), statuses = uniq(res.results.map(function (r) { return r.STATUS; })).sort();
    var sheets = uniq(res.results.map(function (r) { return r.QA_SHEET; })).sort();
    var sm = wb.addWorksheet("Summary");
    sm.addRow(["QA_SHEET"].concat(statuses));
    sheets.forEach(function (sh) {
      sm.addRow([sh].concat(statuses.map(function (st) { return res.results.filter(function (r) { return r.QA_SHEET === sh && r.STATUS === st; }).length; })));
    });
    sm.addRow([]);
    (info || []).forEach(function (r) { sm.addRow(r); });
    sm.getRow(1).font = { bold: true }; sm.getColumn(1).width = 18;
    var dt = wb.addWorksheet("Detail"), keys = ["BATCH_NO", "QA_SHEET", "QA_COLUMN", "STAT", "QA_VALUE", "MES_VALUE", "STATUS",
      "MES_STAGE", "MES_STEP", "MES_PARAMETER", "CANDIDATES"];
    dt.addRow(keys);
    res.results.forEach(function (r) { dt.addRow(keys.map(function (k) { var v = r[k]; return v === null || v === undefined ? undefined : v; })); });
    dt.getRow(1).font = { bold: true }; dt.views = [{ state: "frozen", ySplit: 1 }];
    dt.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: keys.length } };
    [12, 14, 60, 7, 12, 14, 24, 18, 30, 26, 10].forEach(function (w, i) { dt.getColumn(i + 1).width = w; });
    return wb.xlsx.writeBuffer();
  }

  global.AdigiCompare = { flattenQa: flattenQa, compare: compare, run: run, workbook: workbook, RULES: RULES, SHEETS: SHEETS };
})(typeof window !== "undefined" ? window : this);
