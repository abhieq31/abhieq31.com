/* AdigiPQR - QA's own APQR workbook <-> the PAS-X export, in the browser: check every value QA typed, or fill
   QA's workbook (new batches into BMR detail / BPR Details, values into empty cells, never over what QA typed).
   Port of adigipqr/qamap.py (same rules, same statuses; tests/test_compare.py keeps them equal). Nothing leaves the tab. */
(function (global) {
  "use strict";
  // QA names sheets differently per product: the rules work on the kind of sheet (= qamap.SHEET_KINDS)
  var SHEET_KINDS = [
    ["analytical|\\bcoa\\b", "LIMS"],
    ["packing\\s*yield", "Packing Yield"],
    ["^\\s*bpr", "BPR Details"],
    ["bmr", "BMR detail"],
    ["yield", "Yield data"],
    ["temp|humid|\\brh\\b", "Temp. & Humidity"],
    ["granul", "Granulation"],
    ["compress", "Compression"],
    ["coat", "Coating"],
    ["capsul", "Capsule"],
    ["^\\s*pack", "Packing"]
  ];
  var SHEETS = ["BMR detail", "Granulation", "Compression", "Coating", "Capsule", "Temp. & Humidity", "Yield data", "BPR Details", "Packing Yield", "Packing"];
  var BATCH_RE = /^\d{10}$/, BATCH_LABEL = /batch\s*no/i, YELLOW = "FFFFFF00";
  // QA's own fields: nothing in PAS-X to compare or fill
  var QA_FIELD = new RegExp("(^|> )(remarks?|trend(ing)? by( name)?|market|master b[mp]r no\\.?|oos\\s*/\\s*oot\\s*/\\s*deviation\\s*/\\s*change\\s*control|pack style|status|packing batch no\\.?|bpr no\\.?|ref\\.?\\s*doc\\.?\\s*no\\.?|verified by|esta\\w*lished control limits?|control limits? esta\\w*lished|specification no\\.?)$");
  // batch header fields of the export: (QA column regex on 'sheet kind > column', field)
  var HEADER_RULES = [
    ["^(bmr detail|bpr details) > batch size", "BATCH_SIZE"],
    ["^(bmr detail|bpr details) > mfg\\.? ?date", "MFG_MONTH"],
    ["^(bmr detail|bpr details) > exp\\.? ?date", "EXP_MONTH"],
    ["^(bmr detail|bpr details) > (sfg( product)? code|product code)", "MATERIAL_CODE"],
    ["^bmr detail > (mini tablet|input|sfg|intermediate)\\s*batch", "INPUT_SFG_BATCHES"],
    ["^(bmr detail|bpr details) > product name", "PRODUCT_NAME"],
    ["^bmr detail > bmr no", "BMR_NO"],
    ["^(bmr detail|bpr details) > shelf life", "SHELF_LIFE"],
    ["^bpr details > fg code", "FG_CODE"],
    ["^bmr detail > (manufacturing order|process order|mo no)", "MANUFACTURING_ORDER"]
  ];
  // (qa_regex, mes_param_regex, stage_regex, step_regex, value_kind) = adigipqr/qamap.RULES
  var RULES = [
    ["^(granulation|compression|coating|capsule|packing yield) > (equipment|granulation ?\\()", null, null, null, "E"],
    ["^(granulation|compression|coating|capsule) > (dispensing|blending|lubrication|granulation|compression|coating|capsule filling)$", null, null, null, "E"],
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
    ["milled granules.*lod", "lod", "granul", "mill", "R"],
    ["^coating > coating > average weight >", "final average weight", "coat", null, "R"],
    ["coating.*uniformity", "uniformity", "coat", null, "I"],
    ["coating.*disintegration", "disintegration", "coat", null, "I"],
    ["coating.*ave\\. weight", "average weight|av\\w*\\.?\\s*mass", "coat", null, "R"],
    ["coating.*wt\\.? ?of 20", "(mass|weight|wt\\.?) of 20", "coat", null, "R"],
    ["group weight", "weight of 20|mass of 20", "coat", null, "R"],
    ["preheating.*bed", "bed|product", "coat", "preheat", "R"],
    ["preheating time", "preheat\\w*\\s*time|time", "coat", "preheat", "R"],
    ["spraying gun", "gun", "coat", null, "R"],
    ["distance between bed and gun", "distance|bed to gun|gun to", "coat", null, "R"],
    ["coating parameter.*inlet air temp", "inlet", "coat", "spray", "R"],
    ["coating parameter.*bed", "bed|product", "coat", "spray", "R"],
    ["coating parameter.*exhaust", "exhaust", "coat", "spray", "R"],
    ["pan rpm", "pan", "coat", null, "R"],
    ["peristaltic|\\bpp rpm", "peristaltic|pp\\s*rpm", "coat", "spray", "R"],
    ["atomi", "atomi", "coat", "spray", "R"],
    ["spray rate", "spray\\s*rate", "coat", null, "R"],
    ["coating parameter.*inlet air flow", "inlet air flow|cfm", "coat", "spray", "R"],
    ["coating parameter.*% rh", "rh", "coat", "spray", "R"],
    ["drying time", "drying time", "coat", null, "R"],
    ["cooling time", "cooling time", "coat", null, "R"],
    ["weight gain", "weight gain", "coat", null, "R"],
    ["yield data.*(lb|lubrication)\\s*-?\\s*rejection", "reject", "lub|blend", null, "Y"],
    ["yield data.*compression\\s*-?\\s*rejection", "reject", "compress", null, "Y"],
    ["yield data.*coating\\s*-?\\s*rejection", "reject", "coat", null, "Y"],
    ["yield data.*capsul\\w*\\s*-?\\s*rejection", "reject", "capsul|fill", null, "Y"],
    ["packing yield.*packing\\s*-?\\s*rejection", "reject", "pack", null, "Y"],
    ["yield data.*lubrication", "actual yield|% yield", "lub|blend", null, "Y"],
    ["yield data.*compression >", "actual yield|% yield", "compress", null, "Y"],
    ["yield data.*coating >", "actual yield|% yield", "coat", null, "Y"],
    ["yield data.*capsul", "actual yield|% yield|yield", "capsul|fill", null, "Y"],
    ["packing yield.*parameter > packing >", "yield", "pack", null, "Y"]
  ];
  var ENV_STAGE = [["dispens", "dispens"], ["sift", "sift"], ["blend|lubric", "blend|lub"], ["roll ?comp", "roll|compact"], ["gran", "granul|roll|compact"], ["mix", "mix"], ["capsul", "capsul|encapsul|fill"], ["\\bcomp\\b|compress", "compress"], ["coat", "coat"]];
  var PART_HINTS = [["amlo", "amlo"], ["valsartan", "vals"], ["lot-i ", "lot.?i\\b"], ["lot-ii", "lot.?ii"]];
  var EQUIP_STAGE = [["dispens", "dispens"], ["lubrication", "lub|blend"], ["blend", "blend|lub"], ["compression", "compress"], ["coating", "coat"], ["granulation", "roll|compact|granul"], ["capsul", "capsul|encapsul|fill"], ["packing", "pack|blister"]];
  var EQUIP_TOKEN = new RegExp("(?:[A-Z]{2}/\\d{2})?/[A-Z]{2,5}\\d{2,3}\\b", "g");
  var STATUS_TEXT = {"FILLED": "written from PAS-X", "KEPT": "QA had already typed a value: left as it was", "LEFT_FOR_QA": "QA's own field, LIMS sheet or yellow column", "NOT_IN_MES": "not recorded in this PAS-X export", "AMBIGUOUS": "several PAS-X readings could be meant: left for QA", "NO_MAPPING_RULE": "column not mapped yet: left for QA", "OTHER_EQUIPMENT": "column of another machine than the batch was made on: left empty", "SHEET_SKIPPED": "sheet not filled", "NOTHING_FILLED": "check these batches: no value was written for them"};
  var STOP = ["of", "the", "and", "in", "at", "to", "for", "by", "no", "min", "max", "ave", "avg", "nlt", "nmt", "between"];
  var RX = {};
  function rx(p) { return RX[p] || (RX[p] = new RegExp(p)); }
  function sheetKind(name) { var t = s(name).toLowerCase(); for (var i = 0; i < SHEET_KINDS.length; i++) if (rx(SHEET_KINDS[i][0]).test(t)) return SHEET_KINDS[i][1]; return null; }

  /* ---------------------------------------------------------------- values (= qamap.num / to_num / blank) */
  function num(v) {
    if (v === null || v === undefined || typeof v === "boolean") return null;
    if (typeof v === "number") return isFinite(v) ? v : null;
    if (v instanceof Date || typeof v === "object") return null;
    var t = String(v).replace(/,/g, "").trim();
    return /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(t) ? Number(t) : null;
  }
  function toNum(v) {                             // a pure time cell (00:50) = minutes:seconds as QA types it
    if (v instanceof Date) {
      var serial = (v.getTime() - Date.UTC(1899, 11, 30)) / 86400000;
      return serial >= 0 && serial < 1 ? v.getUTCHours() * 60 + v.getUTCMinutes() + v.getUTCSeconds() / 60 : null;
    }
    return num(v);
  }
  function blank(v) { return v === null || v === undefined || (typeof v === "number" && isNaN(v)) || ["", "NA", "IA", "None", "nan"].indexOf(String(v).trim()) >= 0; }
  function str(v) { return v instanceof Date ? v.toISOString() : String(v); }
  function s(v) { return v === null || v === undefined ? "" : str(v); }
  function words(v) { return str(v).split(/\s+/).filter(Boolean).join(" "); }
  function uniq(a) { var seen = new Set(), o = []; a.forEach(function (x) { if (!seen.has(x)) { seen.add(x); o.push(x); } }); return o; }
  function r4(x) { return Math.round(x * 1e4) / 1e4; }

  /* ---------------------------------------------------------------- cells */
  function cached(cell) {                          // what Excel last showed (openpyxl data_only)
    if (!cell || cell.type === 1) return null;     // ExcelJS ValueType.Merge: hidden cell of a merged range
    var v = cell.value;
    if (v === undefined || v === null) return null;
    if (v instanceof Date || typeof v !== "object") return v;
    if (v.richText) return v.richText.map(function (t) { return t.text; }).join("");
    if ("formula" in v || "sharedFormula" in v || "result" in v) return v.result === undefined ? null : (v.result && v.result.error ? v.result.error : v.result);
    if ("text" in v) return v.text;
    if ("error" in v) return v.error;
    return String(v);
  }
  function raw(cell) {                             // formulas as text (openpyxl without data_only)
    if (!cell || cell.type === 1) return null;
    var v = cell.value;
    if (v && typeof v === "object" && !(v instanceof Date) && ("formula" in v || "sharedFormula" in v)) return "=" + (cell.formula || v.formula);
    return cached(cell);
  }
  function at(ws, r, c, mode) { var row = ws.findRow(r); var cell = row ? row.findCell(c) : null; return mode === "raw" ? raw(cell) : cached(cell); }
  function isYellow(cell) {
    var f = cell && cell.fill;
    return !!(f && f.type === "pattern" && f.pattern === "solid" && f.fgColor && f.fgColor.argb && String(f.fgColor.argb).toUpperCase() === YELLOW);
  }
  function addr(r, c) { var t = ""; while (c > 0) { var m = (c - 1) % 26; t = String.fromCharCode(65 + m) + t; c = Math.floor((c - 1) / 26); } return t + r; }

  /* ---------------------------------------------------------------- layout (= qamap.layout) */
  function layout(ws, mode) {
    var merged = {};
    ((ws.model && ws.model.merges) || []).forEach(function (rng) {
      var p = rng.split(":"), a = ws.getCell(p[0]), b = ws.getCell(p[1] || p[0]), v = mode === "raw" ? raw(a) : cached(a);
      for (var r = a.row; r <= b.row; r++) for (var c = a.col; c <= b.col; c++) merged[r + "," + c] = v;
    });
    var cv = function (r, c) { return at(ws, r, c, mode); };
    var val = function (r, c) { var k = r + "," + c; return k in merged ? merged[k] : cv(r, c); };
    var first = null, bcol = null, r, c;
    for (r = 1; r < 40 && !first; r++) for (c = 1; c < 5; c++) { var v = cv(r, c); if (v !== null && BATCH_RE.test(str(v).trim())) { first = r; bcol = c; break; } }
    if (!first) {                                  // empty template: below the header block
      var label = null;
      for (r = 1; r < 16; r++) for (c = 1; c < 5; c++) { var lv = cv(r, c); if (lv !== null && BATCH_LABEL.test(str(lv)) && (!label || r > label[0])) label = [r, c]; }
      if (!label) return null;
      bcol = label[1];
      var end = label[0];
      ((ws.model && ws.model.merges) || []).forEach(function (rng) {
        var p = rng.split(":"), a = ws.getCell(p[0]), b = ws.getCell(p[1] || p[0]);
        if (a.row <= 3 && b.row <= 12) end = Math.max(end, b.row);
      });
      first = end + 1;
    }
    var bl = [];
    for (r = 1; r < first; r++) { var bv = val(r, bcol); if (bv !== null) bl.push(words(bv)); }
    var cols = [], lastParts = null, maxCol = ws.columnCount;
    for (c = bcol + 1; c <= maxCol; c++) {
      var parts = [];
      for (r = 1; r < first; r++) { var hv = val(r, c); if (hv !== null) { var t = words(hv); if (!parts.length || parts[parts.length - 1] !== t) parts.push(t); } }
      if (!parts.length) continue;
      var leaf = parts[parts.length - 1].toLowerCase().replace(/\.+$/, ""), stat = "VALUE";
      if (leaf === "min" || leaf === "minimum") stat = "MIN";
      else if (leaf === "max" || leaf === "maximum") stat = "MAX";
      else if (leaf === "ave" || leaf === "avg" || leaf === "average") {
        stat = "AVG";
        if (parts.length <= 2 && lastParts) parts = lastParts.slice(0, -1).concat(["Ave"]);   // bare "Ave" column
      }
      if (stat === "MIN" || stat === "MAX") lastParts = parts;
      var yellow = false;
      for (r = 1; r < first; r++) { var row = ws.findRow(r); if (row && isYellow(row.findCell(c))) yellow = true; }
      cols.push({ c: c, parts: parts, stat: stat, yellow: yellow });
    }
    var sr = bcol > 1 && /sr\.?\s*no/i.test(s(val(1, 1)));
    return { first: first, bcol: bcol, packing: bl.join(" ").toLowerCase().indexOf("packing") >= 0, cols: cols, sr: sr };
  }

  /* ---------------------------------------------------------------- 1. flatten QA's workbook */
  function flattenQa(wb, batchPrefix) {
    var rows = [];
    wb.worksheets.forEach(function (ws) {
      if (SHEETS.indexOf(sheetKind(ws.name)) < 0) return;
      var lay = layout(ws, "cached");
      if (!lay) return;
      for (var r = lay.first; r <= ws.rowCount; r++) {
        var b = at(ws, r, lay.bcol);
        if (b === null || !BATCH_RE.test(str(b).trim())) continue;
        b = str(b).trim();
        if (batchPrefix && b.indexOf(batchPrefix) !== 0) continue;
        lay.cols.forEach(function (col) {
          rows.push({ BATCH_NO: b, QA_SHEET: ws.name, QA_COLUMN: col.parts.join(" > "), STAT: col.stat, QA_VALUE: at(ws, r, col.c) });
        });
      }
    });
    return rows;
  }

  /* ---------------------------------------------------------------- 2. lookup (= qamap.lookup) */
  function prepare(rows) {
    var by = {};
    rows.forEach(function (r) {
      var g = function (k) { return s(r[k]); };
      var m = { row: r, BATCH_NO: g("BATCH_NO").replace(/\.0$/, ""), STAGE: g("STAGE"), STEP_PATH: g("STEP_PATH"),
        APQR_PARAMETER: g("APQR_PARAMETER"), PARAMETER_CLASS: g("PARAMETER_CLASS"), EQUIPMENT_IDS: g("EQUIPMENT_IDS"),
        PRODUCTION_UNIT_ID: g("PRODUCTION_UNIT_ID"), _stage: g("STAGE").toLowerCase(),
        _step: (g("STEP_PATH") + " " + g("STEP_CODE_PATH")).toLowerCase(), _param: g("APQR_PARAMETER").toLowerCase(),
        _sheet: g("APQR_SHEET").toLowerCase(), _uom: g("UOM").trim().toLowerCase() };
      // text derived once per row, not once per QA cell (by-name matching runs for thousands of cells)
      m._kindS = m._sheet ? memo("k" + m._sheet, function () { return sheetKind(m._sheet); }) : null;
      m._bare = memo("b" + m.APQR_PARAMETER, function () { return bare(m.APQR_PARAMETER); });
      m._bareTok = memo("t" + m._bare, function () { return tokens(m._bare); });
      m._fullTok = memo("f" + m.APQR_PARAMETER, function () { return tokens(m.APQR_PARAMETER).join(" "); });
      m._pathTok = memo("p" + m.STAGE + "\u0000" + m.STEP_PATH, function () { return uniq(tokens(m.STAGE + " " + m.STEP_PATH)); });
      (by[m.BATCH_NO] = by[m.BATCH_NO] || []).push(m);
    });
    return by;
  }
  var MEMO = {};
  function memo(k, fn) { return k in MEMO ? MEMO[k] : (MEMO[k] = fn()); }
  function headerValue(b, field) {
    var first = function (k) { for (var i = 0; i < b.length; i++) { var v = s(b[i].row[k]).trim(); if (v) return v; } return ""; };
    if (field === "BMR_NO") {                     // QA writes PAS-X-era BMR numbers as F1M00011/00000004
      var mbr = first("MBR_ID"), ver = first("MBR_VERSION");
      return mbr + (mbr && ver ? (/^\d+$/.test(ver) ? "/" + ("00000000" + ver).slice(-Math.max(8, ver.length)) : "-" + ver) : "");
    }
    if ((field === "SHELF_LIFE" || field === "EXP_MONTH") && num(first("SHELF_LIFE")) === 0) return "";   // no expiry (QA: NA)
    if (field === "SHELF_LIFE") { var n = first("SHELF_LIFE"); return n ? (n + " " + first("SHELF_LIFE_UOM")).trim() : ""; }
    if (field === "FG_CODE") return first("SFO_MATERIAL");
    if (field === "BATCH_SIZE") {
      var v = first("BATCH_SIZE"), x = num(v);
      if (["kg", "kgs", ""].indexOf(first("BATCH_SIZE_UOM").toLowerCase()) < 0) {   // e.g. 12000000 tablets: QA writes kg
        var kg = stageKg(b.map(function (m) { return m.STAGE; }));
        if (kg !== null) return kg;
      }
      return x !== null ? x : v;
    }
    return first(field);
  }
  function stageKg(stages) {                      // 'Granulation (297.520 Kg)' -> 297.52 (= qamap.stage_kg)
    for (var i = 0; i < stages.length; i++) {
      var k = s(stages[i]).match(/\(\s*([\d.,]+)\s*kgs?\s*\)/i);
      if (k) return num(k[1]);
    }
    return null;
  }
  function has(words, w, rev) {                   // w is one of words, or an abbreviation (>= 3 letters) of / by one (= qamap._has)
    return words.indexOf(w) >= 0 || words.some(function (x) { return rev ? x.indexOf(w) === 0 && w.length >= 3 : w.indexOf(x) === 0 && x.length >= 3; });
  }
  /* QA's Min / Max of a spraying parameter = PAS-X's own 'Min. Inlet Air Temp.' / 'Max. Inlet Air Temp.' rows (= qamap._minmax) */
  function minmax(c, b, column, stat) {
    if ((stat !== "MIN" && stat !== "MAX") || !c.length) return c;
    var want = stat.toLowerCase(), sect = tokens(bare(column.split(" > ")[0])), out = [];
    var bases = c.map(function (m) { return tokens(m._param); });
    var stages = {}; c.forEach(function (m) { stages[m._stage] = 1; });
    b.forEach(function (m) {
      var mm = m._param.match(/^(min|max)\.?\s+(.+)/);
      if (!mm || mm[1] !== want || !stages[m._stage] || !sect.length || !sect.every(function (w) { return m._step.indexOf(w) >= 0; })) return;
      var t = tokens(mm[2]);
      if (t.length && bases.some(function (base) { return t.length === base.length && t.every(function (x) { return has(base, x, true); }); })) out.push(m);
    });
    return out.length ? out : c;
  }
  function tokens(t) { return (s(t).toLowerCase().match(/[a-z0-9%]+/g) || []).filter(function (w) { return STOP.indexOf(w) < 0; }); }
  function bare(t) { return tokens(s(t).replace(/\[[^\]]*\]|\([^)]*\)/g, " ")).join(" "); }
  /* generic match for products the rules do not know (= qamap.by_name) */
  function byName(sheet, column, b) {
    var path = tokens(column), parts = column.split(" > ").map(bare), kindS = sheetKind(sheet);
    var fullParts = column.split(" > ").map(function (p) { return tokens(p).join(" "); }), lvs = new Map();
    var multi = parts.map(tokens).filter(function (t) { return t.length >= 2; });
    var c = b.filter(function (m) {
      if (m.PARAMETER_CLASS.indexOf("Environment (Temp") === 0 || (m._sheet && m._kindS !== kindS)) return false;
      var name = m._bare, ws = m._bareTok;
      // 3: same name with its brackets ('Uniformity of Mass (Net Content)'), 2: same name, 1: all its words appear
      // in QA's header, or all words of one part of QA's header appear in it ('% Flap' -> '% Flap Position');
      // QA abbreviates: 'Bed Temp.' = 'Bed Temperature'
      var lv = fullParts.indexOf(m._fullTok) >= 0 ? 3 : parts.indexOf(name) >= 0 ? 2 :
        name && ((ws.length >= 2 && ws.every(function (w) { return has(path, w); })) ||
          multi.some(function (p) { return p.every(function (w) { return has(ws, w, true); }); })) ? 1 : 0;
      if (name && lv) { lvs.set(m, lv); return true; }
      return false;
    });
    if (c.length) {
      var lvOf = function (m) { return lvs.get(m); };
      var bestLv = Math.max.apply(null, c.map(lvOf));
      c = c.filter(function (m) { return lvOf(m) === bestLv; });
      var nw = function (m) { return m._bareTok.length; }, longest = Math.max.apply(null, c.map(nw));
      c = c.filter(function (m) { return nw(m) === longest; });   // 'Milled Granules % LOD' beats '% LOD'
      var score = function (m) { return m._pathTok.filter(function (w) { return path.indexOf(w) >= 0; }).length; };
      var best = Math.max.apply(null, c.map(score));
      c = c.filter(function (m) { return score(m) === best; });
    }
    return [c, uniq(c.map(function (m) { return m._param; })).length > 1];
  }
  function ipqcKind(c, other) { return c.every(function (m) { return m.PARAMETER_CLASS === "IPQC (multi-value)"; }) ? "I" : other; }
  function pick(cands, kind, stat) {
    var col = { R: { MIN: "READING_MIN", MAX: "READING_MAX", AVG: "READING_AVG", VALUE: "READING_MAX" },
      I: { MIN: "INDIVIDUAL_MIN", MAX: "INDIVIDUAL_MAX", AVG: "INDIVIDUAL_AVG", VALUE: "INDIVIDUAL_AVG" },
      Y: { MIN: "READING_MIN", MAX: "READING_MAX", AVG: "READING_AVG", VALUE: "READING_MAX" } }[kind][stat];
    if (kind === "I" && cands.every(function (m) { return num(m.row[col]) === null; })) col = col.replace("INDIVIDUAL", "READING");
    return col;
  }
  function lookup(sheet, column, stat, b) {
    var res = { kind: null, status: null, value: null, vals: [], cands: 0, top: null, unresolved: false, byName: false, minutes: false };
    var kindS = sheetKind(sheet) || sheet, col = column.toLowerCase(), full = (kindS + " > " + column).toLowerCase(), c, kind, i;
    if (QA_FIELD.test(col)) { res.kind = "Q"; res.status = "QA_FIELD"; return res; }
    for (i = 0; i < HEADER_RULES.length; i++) {
      if (rx(HEADER_RULES[i][0]).test(full)) {
        var hv = headerValue(b, HEADER_RULES[i][1]);
        res.kind = "H"; res.value = hv !== "" ? hv : null; res.status = hv !== "" ? null : "NOT_IN_MES"; return res;
      }
    }
    var machines = column.toUpperCase().match(EQUIP_TOKEN) || [];
    if (machines.length) {
      var used = b.map(function (m) { return m.EQUIPMENT_IDS; }).join(" ").toUpperCase();
      if (!used.trim()) res.unresolved = true;     // no equipment in the export: cannot tell the machine
      else if (!machines.some(function (t) { return used.indexOf(t) >= 0; })) { res.status = "OTHER_EQUIPMENT"; return res; }
    }
    if (kindS === "Temp. & Humidity" || /^(temp|rh)\b/.test(col)) {
      var pkw = col.indexOf("temp") >= 0 && col.indexOf("humid") < 0 ? "temp" : "humid|rh", skw = null;
      for (i = 0; i < ENV_STAGE.length; i++) if (rx(ENV_STAGE[i][0]).test(col)) { skw = ENV_STAGE[i][1]; break; }
      if (skw === null && col.indexOf("pack") >= 0) skw = "pack";
      c = b.filter(function (m) { return m.PARAMETER_CLASS.indexOf("Environment (Temp") === 0 && rx(pkw).test(m._param); });
      if (skw) c = c.filter(function (m) { return rx(skw).test(m._stage); });
      // "Dispensing (Coating)": a stage named in brackets must match too, or the value is not in this export
      (col.match(/\(([^)]*)\)/g) || []).forEach(function (qq) {
        var q = qq.slice(1, -1), qr = null;
        for (var k = 0; k < ENV_STAGE.length; k++) if (rx(ENV_STAGE[k][0]).test(q) && ENV_STAGE[k][1] !== skw) { qr = ENV_STAGE[k][1]; break; }
        if (qr) c = c.filter(function (m) { return rx(qr).test(m._stage) || rx(qr).test(m._step); });
      });
      kind = "R";
    } else {
      var rule = null;
      for (i = 0; i < RULES.length; i++) if (rx(RULES[i][0]).test(full)) { rule = RULES[i]; break; }
      if (!rule) {
        var bn = byName(sheet, column, b);
        if (!bn[0].length) { res.status = "NO_MAPPING_RULE"; return res; }
        var cm = minmax(bn[0], b, column, stat);
        return values(res, cm, ipqcKind(bn[0], "R"), stat, cm === bn[0] ? bn[1] : uniq(cm.map(function (m) { return m._param; })).length > 1, column, true);
      }
      kind = rule[4];
      if (kind === "E") {
        var eb = b, er = null;                     // the column's own words first, then the sheet kind
        for (i = 0; i < EQUIP_STAGE.length && er === null; i++) if (col.indexOf(EQUIP_STAGE[i][0]) >= 0) er = EQUIP_STAGE[i][1];
        for (i = 0; i < EQUIP_STAGE.length && er === null; i++)
          if (full.indexOf(EQUIP_STAGE[i][0]) >= 0 || (EQUIP_STAGE[i][0] === "packing" && kindS === "Packing Yield")) er = EQUIP_STAGE[i][1];
        if (er !== null) eb = b.filter(function (m) { return rx(er).test(m._stage); });
        PART_HINTS.slice(0, 2).forEach(function (t) {
          if (col.indexOf(t[0]) >= 0 && eb.some(function (m) { return rx(t[1]).test(m._stage); })) eb = eb.filter(function (m) { return rx(t[1]).test(m._stage); });
        });
        res.kind = "E";
        if (!eb.length) { res.status = "NOT_IN_MES"; return res; }
        // equipment IDs only: the PAS-X production unit (COAT05A) is not the asset number QA writes (PD/09/COT14)
        var eq = uniq(eb.map(function (m) { return m.EQUIPMENT_IDS; })).join(" ").trim();
        res.value = eq.slice(0, 200) || null; res.status = eq ? null : "NOT_IN_MES"; return res;
      }
      c = b.filter(function (m) { return rx(rule[1]).test(m._param); });
      if (rule[2]) c = c.filter(function (m) { return rx(rule[2]).test(m._stage) || rx(rule[2]).test(m._step); });
      if (rule[3]) c = c.filter(function (m) { return rx(rule[3]).test(m._step); });
      if (kind === "Y") c = c.filter(function (m) { return m.PARAMETER_CLASS === "Yield & Rejection"; });
      if (!c.length && (kind === "R" || kind === "I")) {      // the rule's stage words do not fit this product: by name
        var bn2 = byName(sheet, column, b);
        if (bn2[0].length) {
          var cm2 = minmax(bn2[0], b, column, stat);
          return values(res, cm2, ipqcKind(bn2[0], kind), stat, cm2 === bn2[0] ? bn2[1] : uniq(cm2.map(function (m) { return m._param; })).length > 1, column, true);
        }
      }
    }
    // part A / part B, lot I / lot II: if QA names a part that the PAS-X rows do not carry, do not guess
    PART_HINTS.forEach(function (t) {
      if (col.indexOf(t[0]) >= 0 && c.length) {
        var cc = c.filter(function (m) { return rx(t[1]).test(m._stage) || rx(t[1]).test(m._step); });
        if (cc.length) c = cc; else res.unresolved = true;
      }
    });
    return values(res, kind === "R" ? minmax(c, b, column, stat) : c, kind, stat, false, column, false);
  }
  var SECONDS = ["s", "sec", "secs", "second", "seconds"], MINUTES = /\bmin(ute)?s?\b|\bmins?\./;
  function wantsMinutes(column, stat) {           // the last part is the Min/Max label, not a unit
    var parts = column.split(" > ");
    if (stat === "MIN" || stat === "MAX" || stat === "AVG") parts = parts.slice(0, -1);
    return MINUTES.test(parts.join(" ").toLowerCase());
  }
  function values(res, c, kind, stat, several, column, byNm) {
    res.kind = kind; res.cands = c.length; res.byName = byNm; res.unresolved = res.unresolved || several;
    if (!c.length) { res.status = "NOT_IN_MES"; return res; }
    var colname = pick(c, kind, stat), withVal = c.filter(function (m) { return num(m.row[colname]) !== null; });
    if (!withVal.length) { res.status = "NOT_IN_MES"; return res; }
    var vals = withVal.map(function (m) { return num(m.row[colname]); });
    if (withVal.every(function (m) { return SECONDS.indexOf(m._uom) >= 0; }) && wantsMinutes(column, stat)) {   // PAS-X seconds, QA minutes
      vals = vals.map(function (v) { return v / 60; }); res.minutes = true;
    }
    res.vals = vals; res.top = withVal[0];
    res.value = stat === "MIN" ? Math.min.apply(null, vals) : stat === "MAX" ? Math.max.apply(null, vals) : vals[0];
    return res;
  }
  function close(a, b) { return Math.abs(a - b) <= Math.max(0.011, Math.abs(b) * 0.005); }
  function sameText(a, b) {                       // 'F1M00011/00000004' = 'F1M00011/4', 'JAN.2026' = 'Jan 2026'
    var k = function (v) { return (s(v).toUpperCase().match(/[A-Z0-9]+/g) || []).map(function (ch) { return ch.replace(/^0+(?=\d)/, ""); }).join("/"); };
    return k(a) === k(b);
  }

  /* ---------------------------------------------------------------- 3. compare (= qamap.compare) */
  function compare(qa, mesRows) {
    var by = prepare(mesRows);
    return qa.map(function (q) {
      var rec = { BATCH_NO: q.BATCH_NO, QA_SHEET: q.QA_SHEET, QA_COLUMN: q.QA_COLUMN, STAT: q.STAT, QA_VALUE: q.QA_VALUE,
        MES_VALUE: null, STATUS: null, MES_STAGE: null, MES_STEP: null, MES_PARAMETER: null, CANDIDATES: 0 };
      if (blank(q.QA_VALUE)) { rec.STATUS = "QA_BLANK"; return rec; }
      var b = by[q.BATCH_NO];
      if (!b) { rec.STATUS = "BATCH_NOT_IN_EXTRACT"; return rec; }
      var lk = lookup(q.QA_SHEET, q.QA_COLUMN, q.STAT, b);
      rec.CANDIDATES = lk.cands;
      if (lk.status) { rec.STATUS = lk.status; return rec; }
      if (lk.kind === "E") {
        rec.MES_VALUE = lk.value;
        rec.STATUS = lk.value.toUpperCase().indexOf(s(q.QA_VALUE).trim().toUpperCase()) >= 0 ? "MATCH" : "MISMATCH";
        return rec;
      }
      if (lk.kind === "H") {
        rec.MES_VALUE = lk.value;
        var hq = toNum(q.QA_VALUE), hm = num(lk.value);
        rec.STATUS = (hq !== null && hm !== null && close(hm, hq)) || sameText(q.QA_VALUE, lk.value) ? "MATCH" : "MISMATCH";
        return rec;
      }
      var mv = lk.value, top = lk.top, qv = toNum(q.QA_VALUE);
      rec.MES_VALUE = mv; rec.MES_STAGE = top.STAGE; rec.MES_STEP = top.STEP_PATH; rec.MES_PARAMETER = top.APQR_PARAMETER;
      if (qv === null) rec.STATUS = "QA_NOT_NUMERIC";
      else if (close(mv, qv)) rec.STATUS = lk.cands === 1 || uniq(lk.vals.map(r4)).length === 1 ? "MATCH" : "MATCH (several MES rows)";
      else if ([1 / 60, 60].some(function (f) { return close(mv * f, qv); })) rec.STATUS = "MATCH (unit sec/min)";
      else rec.STATUS = lk.cands === 1 && !lk.unresolved ? "MISMATCH" : "AMBIGUOUS";
      return rec;
    });
  }

  /* ---------------------------------------------------------------- 4. fill QA's workbook (= qamap.fill) */
  var MASTERS = ["BMR detail", "BPR Details"];
  var REF = /^=?\(?\s*(?:'([^']+)'|([A-Za-z0-9_.]+))!\$?([A-Z]{1,3})\$?(\d+)\s*\)?$/;
  function formulaRef(v) {
    if (typeof v !== "string" || v.charAt(0) !== "=") return null;
    var m = v.trim().match(REF);
    return m ? [m[1] || m[2], m[3] + m[4]] : null;
  }
  function resolve(wb, ws, a, depth) {
    var v = raw(ws.getCell(a)), ref = formulaRef(v);
    if (ref && (depth || 0) < 3 && wb.getWorksheet(ref[0])) return resolve(wb, wb.getWorksheet(ref[0]), ref[1], (depth || 0) + 1);
    return ref ? null : v;
  }
  function fill(wb, mesRows, opts) {
    opts = opts || {};
    var by = prepare(mesRows), order = uniq(mesRows.map(function (r) { return s(r.BATCH_NO).replace(/\.0$/, ""); })), report = [];
    var note = function (sheet, status, text) { report.push({ SHEET: sheet, CELL: "", BATCH_NO: "", COLUMN: "", STAT: "", VALUE: null, STATUS: status, SOURCE: text }); };
    var masters = MASTERS.map(function (k) { return wb.worksheets.filter(function (w) { return sheetKind(w.name) === k; })[0]; }).filter(Boolean);
    var sheets = masters.concat(wb.worksheets.filter(function (w) { return masters.indexOf(w) < 0; }));
    sheets.forEach(function (ws) {
      var kindW = sheetKind(ws.name), mapped = SHEETS.indexOf(kindW) >= 0, lims = kindW === "LIMS";
      if (!(mapped || lims)) return;
      var lay = layout(ws, "raw");
      if (!lay) return;
      if (lay.packing) { note(ws.name, "SHEET_SKIPPED", "rows are packing batch numbers; the PAS-X export lists manufacturing batches"); return; }
      if (lims && !mapped) { note(ws.name, "LEFT_FOR_QA", "LIMS results: the batch rows come from the linked batch column"); return; }
      var bc = lay.bcol, rows = {}, last = lay.first - 1, srs = [], intBatch = false, linked = false, r, c;
      for (r = lay.first; r <= ws.rowCount; r++) {
        var rv = at(ws, r, bc, "raw"), isRef = formulaRef(rv) !== null;
        linked = linked || isRef;
        var v = resolve(wb, ws, addr(r, bc));
        if (v !== null && v !== undefined && BATCH_RE.test(str(v).trim())) {
          if (!(str(v).trim() in rows)) rows[str(v).trim()] = r;
          if (!isRef) last = Math.max(last, r);
          intBatch = intBatch || (typeof v === "number" && Number.isInteger(v));
          if (lay.sr && num(at(ws, r, 1, "raw")) !== null) srs.push(num(at(ws, r, 1, "raw")));
        }
      }
      var nSr = srs.length ? Math.floor(Math.max.apply(null, srs)) : 0;
      var width = Math.max.apply(null, [bc].concat(lay.cols.map(function (x) { return x.c; })));
      // batch numbers as numbers, as QA types them (text only if QA's own batch cells are text)
      var asInt = intBatch || !Object.keys(rows).some(function (k) { return formulaRef(at(ws, rows[k], bc, "raw")) === null; });
      order.forEach(function (bno) {
        if (!(bno in rows)) {
          if (linked) return;                      // the master sheet did not get this batch: nothing to fill
          last += 1; r = rows[bno] = last;
          if (r !== lay.first) {
            for (c = 1; c <= width; c++) {
              var src = ws.getCell(lay.first, c), dst = ws.getCell(r, c);
              if (dst.type !== 1 && src.style && Object.keys(src.style).length) dst.style = JSON.parse(JSON.stringify(src.style));
            }
          }
          ws.getCell(r, bc).value = asInt && /^\d+$/.test(bno) ? Number(bno) : bno;
          if (lay.sr) {
            if (num(raw(ws.getCell(r, 1))) === null) ws.getCell(r, 1).value = nSr + 1;
            nSr = Math.floor(num(raw(ws.getCell(r, 1))));
          }
        }
        r = rows[bno];
        var b = by[bno] || [];
        lay.cols.forEach(function (col) {
          var column = col.parts.join(" > "), cell = ws.getCell(r, col.c), cur = raw(cell);
          var rec = { SHEET: ws.name, CELL: addr(r, col.c), BATCH_NO: bno, COLUMN: column, STAT: col.stat, VALUE: null, STATUS: null, SOURCE: "" };
          report.push(rec);
          if (cur !== null && String(cur).trim() !== "") { rec.STATUS = "KEPT"; rec.VALUE = cur; return; }
          if (opts.leaveYellow && col.yellow) { rec.STATUS = "LEFT_FOR_QA"; rec.SOURCE = "yellow column (LIMS / Tantra)"; return; }
          var lk = lookup(ws.name, column, col.stat, b);
          if (lk.status) { rec.STATUS = lk.status === "QA_FIELD" ? "LEFT_FOR_QA" : lk.status; return; }
          var several = lk.kind !== "E" && lk.kind !== "H" && uniq(lk.vals.map(r4)).length > 1;
          if (lk.unresolved || (several && (col.stat === "VALUE" || col.stat === "AVG"))) { rec.STATUS = "AMBIGUOUS"; rec.SOURCE = lk.cands + " PAS-X rows could be meant"; return; }
          cell.value = lk.value; rec.VALUE = lk.value; rec.STATUS = "FILLED";
          if (lk.kind === "E") rec.SOURCE = "equipment IDs of the stage";
          else if (lk.kind === "H") rec.SOURCE = "batch header";
          else rec.SOURCE = [lk.top.STAGE, lk.top.STEP_PATH, lk.top.APQR_PARAMETER].filter(Boolean).join(" > ") +
            (several ? " (" + lk.cands + " PAS-X rows, " + col.stat.toLowerCase() + " of all)" : "") + (lk.byName ? " [matched by name]" : "") + (lk.minutes ? " [seconds → minutes]" : "");
        });
      });
    });
    var missing = order.filter(function (bno) { return !report.some(function (x) { return x.BATCH_NO === bno && x.STATUS === "FILLED"; }); });
    if (missing.length) note("", "NOTHING_FILLED", "batches without any value written: " + missing.join(", "));
    // linked cells (='BMR detail'!B6) keep their cached result: set it to what they now point to, so the file shows the
    // right batch numbers even before Excel recalculates (Protected View does not)
    wb.worksheets.forEach(function (ws) {
      ws.eachRow(function (row) {
        row.eachCell(function (cell) {
          var v = cell.value;
          if (!v || typeof v !== "object" || !("formula" in v) || !formulaRef("=" + v.formula)) return;
          var res = resolve(wb, ws, cell.address);
          cell.value = { formula: v.formula, result: res === null || res === undefined || res === "" ? 0 : res };
        });
      });
    });
    wb.calcProperties = wb.calcProperties || {};
    wb.calcProperties.fullCalcOnLoad = true;      // Excel recalculates the linked batch columns on opening
    return report;
  }
  var REPORT_COLS = ["SHEET", "CELL", "BATCH_NO", "COLUMN", "STAT", "VALUE", "STATUS", "SOURCE"];
  function addReport(wb, report, info) {
    var name = "AdigiPQR fill report", old = wb.getWorksheet(name);
    if (old) wb.removeWorksheet(old.id);
    var ws = wb.addWorksheet(name);
    (info || []).forEach(function (r) { var row = ws.addRow(r); row.getCell(1).font = { bold: true }; });
    ws.addRow([]);
    ws.addRow(["Status", "Cells", "Meaning"]).eachCell(function (c) { c.font = { bold: true }; });
    Object.keys(STATUS_TEXT).forEach(function (st) {
      var n = report.filter(function (r) { return r.STATUS === st; }).length;
      if (n) ws.addRow([st, n, STATUS_TEXT[st]]);
    });
    ws.addRow([]);
    var head = ws.addRow(["Sheet", "Cell", "Batch", "QA column", "Stat", "Value", "Status", "Source in PAS-X"]);
    head.eachCell(function (c) { c.font = { bold: true }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDDEBF0" } }; });
    report.forEach(function (r) {
      var row = ws.addRow(REPORT_COLS.map(function (k) { return r[k] === null || r[k] === undefined ? undefined : r[k]; }));
      if (r.STATUS === "AMBIGUOUS") row.getCell(7).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFCE4B6" } };
    });
    [18, 8, 13, 70, 7, 16, 16, 60].forEach(function (w, i) { ws.getColumn(i + 1).width = w; });
    ws.views = [{ state: "frozen", ySplit: head.number }];
    return ws;
  }

  /* ---------------------------------------------------------------- entry points for the page */
  function step(key, detail) { if (global.AdigiProgress) global.AdigiProgress(key, detail); }   // progress for the page
  async function load(buffer) {
    step("open", (buffer.byteLength / 1048576).toFixed(1) + " MB");
    var wb = new ExcelJS.Workbook(); await wb.xlsx.load(buffer); return wb;
  }
  /* QA workbook (ArrayBuffer) + export rows -> { results, missing (QA batches not in the export), qaBatches } */
  async function run(qaBuffer, mesRows, batchPrefix) {
    var wb = await load(qaBuffer);
    step("flat", wb.worksheets.length + " sheets");
    var qa = flattenQa(wb, batchPrefix || "");
    var inMes = {}; mesRows.forEach(function (r) { inMes[s(r.BATCH_NO).replace(/\.0$/, "")] = 1; });
    var qaBatches = uniq(qa.map(function (q) { return q.BATCH_NO; }));
    var missing = qaBatches.filter(function (b) { return !inMes[b]; }).sort();
    var mine = qa.filter(function (q) { return inMes[q.BATCH_NO]; });
    step("match", mine.length.toLocaleString("en-IN") + " QA cells · " + (qaBatches.length - missing.length) + " batches");
    var results = compare(mine, mesRows);
    return { results: results, missing: missing, qaBatches: qaBatches, qaValues: qa.length };
  }
  /* QA workbook (ArrayBuffer) + export rows of the requested batches -> { buffer (filled xlsx), report } */
  async function fillWorkbook(qaBuffer, mesRows, opts) {
    var wb = await load(qaBuffer);
    step("fill", uniq(mesRows.map(function (r) { return r.BATCH_NO; })).length + " batches");
    var report = fill(wb, mesRows, opts);
    step("report", report.filter(function (r) { return r.STATUS === "FILLED"; }).length.toLocaleString("en-IN") + " values written");
    addReport(wb, report, (opts && opts.info) || []);
    step("zip");
    return { buffer: await wb.xlsx.writeBuffer(), report: report };
  }
  /* comparison report: Summary sheet (QA sheet x status counts) + Detail sheet, like compare_apqr.py */
  async function workbook(res, info) {
    step("zip");
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

  global.AdigiCompare = { flattenQa: flattenQa, compare: compare, lookup: lookup, fill: fill, addReport: addReport, run: run,
    fillWorkbook: fillWorkbook, workbook: workbook, RULES: RULES, HEADER_RULES: HEADER_RULES, SHEETS: SHEETS, STATUS_TEXT: STATUS_TEXT };
})(typeof window !== "undefined" ? window : this);
