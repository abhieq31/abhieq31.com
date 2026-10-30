/* AdigiPQR - runs final/apqr_curate.py (the same file as the command line and the plant portal) in this browser tab.
   Python 3.12 = Pyodide, served from this site (vendor/pyodide); openpyxl from its wheel. Nothing is uploaded.
   The page sends files; this runs  python apqr_curate.py <export> [--template T] [--rules R] [--validate QA] [--batches ..]
   exactly as on a PC, and returns the workbook, the printed log and the APQR_Exceptions sheet. */
"use strict";
var base = self.location.href.replace(/app\/py-worker\.js.*$/, ""), py = null;

function step(k, d) { self.postMessage({ step: k, detail: d || "" }); }
async function bytes(url) {
  var r = await fetch(base + url, { cache: "no-cache" });          // revalidated: a changed curator is never served stale
  if (!r.ok) throw new Error(url + ": HTTP " + r.status);
  return new Uint8Array(await r.arrayBuffer());
}

async function start() {
  if (py) return py;
  step("py", "Python 3.12 (WebAssembly), first time about 13 MB");
  importScripts(base + "vendor/pyodide/pyodide.js");
  var p = await loadPyodide({ indexURL: base + "vendor/pyodide/" });
  step("libs", "openpyxl 3.1.5 + et_xmlfile 2.0.0");
  var site = p.runPython("import site; site.getsitepackages()[0]");
  for (var w of ["openpyxl-3.1.5-py2.py3-none-any.whl", "et_xmlfile-2.0.0-py3-none-any.whl"])
    p.unpackArchive(await bytes("vendor/pyodide/" + w), "zip", { extractDir: site });
  p.FS.mkdirTree("/work/lib");
  p.runPython(GLUE);
  py = p;
  return p;
}

// runs main() of apqr_curate.py with the same arguments as the command line; what it prints is the log
var GLUE = [
  "import contextlib, glob, io, os, shutil, sys, importlib",
  "sys.path.insert(0, '/work/lib')",
  "def adigi_run(argv):",
  "    for d in ('/work/in', '/work/qa', '/work/out'):",
  "        os.makedirs(d, exist_ok=True)",
  "    import apqr_curate as A",
  "    A = importlib.reload(A)",
  "    log, stop = io.StringIO(), ''",
  "    os.chdir('/work/out')",
  "    with contextlib.redirect_stdout(log), contextlib.redirect_stderr(log):",
  "        try:",
  "            A.main(argv)",
  "        except SystemExit as e:",
  "            stop = e.code if isinstance(e.code, str) else ('' if not e.code else 'stopped (exit %s)' % e.code)",
  "    out = sorted(glob.glob('/work/out/*.xlsx'))",
  "    exc = []",
  "    if out:",
  "        import openpyxl",
  "        wb = openpyxl.load_workbook(out[0], read_only=True)",
  "        if A.EXC_SHEET in wb.sheetnames:",
  "            exc = [[('' if v is None else str(v)) for v in r] for r in wb[A.EXC_SHEET].iter_rows(values_only=True)]",
  "        wb.close()",
  "    return dict(log=log.getvalue(), stop=stop, out=out[0] if out else '', exceptions=exc, version=A.VERSION)",
].join("\n");

async function run(d) {
  var p = await start();
  step("files", d.files.export.name);
  var code = await bytes("app/apqr_curate.py");
  p.FS.writeFile("/work/lib/apqr_curate.py", code);
  p.runPython("import shutil; [shutil.rmtree(d, ignore_errors=True) for d in ('/work/in', '/work/qa', '/work/out')]");
  ["/work/in", "/work/qa", "/work/out"].forEach(function (x) { p.FS.mkdirTree(x); });
  var argv = [], put = function (dir, f) { var path = dir + "/" + f.name.replace(/[\/\\]/g, "_"); p.FS.writeFile(path, new Uint8Array(f.buffer)); return path; };
  argv.push(put("/work/in", d.files.export));
  if (d.files.template) argv.push("--template", put("/work/in", d.files.template));
  if (d.files.rules) argv.push("--rules", put("/work/in", d.files.rules));
  if (d.files.qa) argv.push("--validate", put("/work/qa", d.files.qa));
  if (d.batches && d.batches.length) argv.push("--batches", d.batches.join(","));
  step("curate", "apqr_curate.py " + argv.map(function (a) { return a.replace(/^\/work\/\w+\//, ""); }).join(" "));
  var res = p.globals.get("adigi_run")(p.toPy(argv)).toJs({ dict_converter: Object.fromEntries });
  var out = { log: res.log, stop: res.stop, exceptions: res.exceptions, version: res.version, sha: await sha(code), argv: argv };
  if (res.out) {
    step("save");
    var b = p.FS.readFile(res.out);
    out.name = res.out.split("/").pop(); out.buffer = b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength);
  }
  return out;
}
async function sha(u8) {
  try { var d = await crypto.subtle.digest("SHA-256", u8); return Array.prototype.map.call(new Uint8Array(d), function (x) { return (x < 16 ? "0" : "") + x.toString(16); }).join(""); }
  catch (e) { return ""; }
}

self.onmessage = async function (e) {
  var d = e.data;
  try {
    var r = await run(d);
    self.postMessage({ id: d.id, result: r }, r.buffer ? [r.buffer] : []);
  } catch (err) { self.postMessage({ id: d.id, error: String(err && err.message || err) }); }
};
