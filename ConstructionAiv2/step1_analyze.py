"""
step1_analyze.py – STEP 1: scan an input folder, assign every file to a project, analyse it
=========================================================================================
For every file in the input folder (sub-folders and .zip archives included):
  * GAEB files (.X8x, GAEB .xml, .D8x)          -> read locally by gaeb_reader.py (exact, free)
  * PDF, images, Office, text, e-mail files     -> Azure Content Understanding (prebuilt analyzer)
  * anything else (.doc, .xls, .dwg, …)         -> listed as "unsupported" with a hint, not sent
Each analysed file gives one JSON (same format as the Azure result you already know) in
OUTPUT/json/, and one line in OUTPUT/file_index.csv (project, route, status, JSON path …).

PROJECT ASSIGNMENT – files are only ever combined inside the SAME project:
  1. --manifest projects.csv  (columns: file, project_id)        -> wins over everything
  2. the first sub-folder of the input folder: INPUT/HEM2301-2/... -> project "HEM2301-2"
  3. files directly in INPUT: project number found in the file (GAEB <LblPrj>, Azure fields
     TenderNumber / ProjectNumber); files with the same number end up in the same project
  4. otherwise the file is its OWN project ("single:<file name>") and is never merged
  Edit file_index.csv (column project_id) and re-run steps 2-3 if an assignment is wrong.

Already analysed files are not sent again (cache by file content hash) -> no double Azure cost.

Usage
    python step1_analyze.py INPUT_FOLDER --out output [--manifest projects.csv] [--workers 4] [--force]
Environment (.env):
    AZURE_CU_ENDPOINT=https://<resource>.services.ai.azure.com     (or ...cognitiveservices.azure.com)
    AZURE_CU_KEY=<key>                       (or AZURE_CU_TOKEN=<Entra ID bearer token>)
    AZURE_CU_ANALYZER=prebuilt-layout        (the analyzerId in your existing JSON also works)
    AZURE_CU_API_VERSION=2025-11-01
"""
import argparse
import csv
import hashlib
import json
import os
import re
import shutil
import time
import urllib.error
import urllib.request
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import features
import gaeb_reader

AZURE_TYPES = {  # extension -> (content type, size limit in bytes)
    ".pdf": ("application/pdf", 200e6), ".tif": ("image/tiff", 200e6), ".tiff": ("image/tiff", 200e6),
    ".jpg": ("image/jpeg", 200e6), ".jpeg": ("image/jpeg", 200e6), ".jpe": ("image/jpeg", 200e6),
    ".png": ("image/png", 200e6), ".bmp": ("image/bmp", 200e6), ".heif": ("image/heif", 200e6),
    ".heic": ("image/heic", 200e6),
    ".docx": ("application/vnd.openxmlformats-officedocument.wordprocessingml.document", 200e6),
    ".xlsx": ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 200e6),
    ".pptx": ("application/vnd.openxmlformats-officedocument.presentationml.presentation", 200e6),
    ".txt": ("text/plain", 1e6), ".md": ("text/markdown", 1e6), ".html": ("text/html", 1e6),
    ".htm": ("text/html", 1e6), ".rtf": ("application/rtf", 1e6), ".xml": ("application/xml", 1e6),
    ".eml": ("message/rfc822", 1e6), ".msg": ("application/vnd.ms-outlook", 1e6),
    ".csv": ("text/csv", 1e6),
}
HINTS = {".doc": "save as .docx or PDF", ".xls": "save as .xlsx", ".ppt": "save as .pptx",
         ".dwg": "CAD drawing – export the relevant sheets as PDF", ".dxf": "CAD drawing – export as PDF",
         ".gif": "convert to PNG", ".odt": "save as .docx or PDF", ".ods": "save as .xlsx"}
SKIP = re.compile(r"(^~\$|^\.|^Thumbs\.db$|^desktop\.ini$)", re.I)
INDEX_COLUMNS = ["project_id", "project_source", "file", "relative_path", "route", "status", "positions",
                 "json_path", "features_path", "sha256", "message"]


# ------------------------------------------------------------------ small helpers
def load_dotenv(path=".env"):
    for f in (Path(path), Path(__file__).with_name(".env")):
        if f.exists():
            for line in f.read_text(encoding="utf-8").splitlines():
                if "=" in line and not line.lstrip().startswith("#"):
                    k, v = line.split("=", 1)
                    os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))
            return


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def safe(name):
    return re.sub(r"[^\w.\-]+", "_", name)[:120]


def detected_project(result_json):
    """Project number inside an analysed file (GAEB <LblPrj>, Azure fields), or None."""
    for c in (result_json.get("result") or {}).get("contents", []):
        for key in ("ProjectNumber", "TenderNumber", "ProjectId", "LblPrj"):
            f = (c.get("fields") or {}).get(key)
            v = (f or {}).get("valueString") if isinstance(f, dict) else None
            if v and v.strip():
                return re.sub(r"\s+", " ", v.strip())
    return None


# ------------------------------------------------------------------ input scan
def scan(input_dir, out_dir):
    """All files (zip archives are extracted into OUTPUT/unzipped/…). Returns list of
    (path, relative_path_inside_input)."""
    input_dir = Path(input_dir).resolve()
    files = []
    for p in sorted(input_dir.rglob("*")):
        if not p.is_file() or SKIP.search(p.name):
            continue
        rel = p.relative_to(input_dir)
        if p.suffix.lower() == ".zip":
            target = Path(out_dir) / "unzipped" / rel.with_suffix("")
            if not target.exists():
                with zipfile.ZipFile(p) as z:
                    for m in z.infolist():                     # guard against "../" paths in archives
                        dest = (target / m.filename).resolve()
                        if not str(dest).startswith(str(target.resolve())):
                            continue
                        if m.is_dir():
                            dest.mkdir(parents=True, exist_ok=True)
                        else:
                            dest.parent.mkdir(parents=True, exist_ok=True)
                            with z.open(m) as src, open(dest, "wb") as dst:
                                shutil.copyfileobj(src, dst)
            for q in sorted(target.rglob("*")):
                if q.is_file() and not SKIP.search(q.name) and q.suffix.lower() != ".zip":
                    files.append((q, rel.with_suffix("") / q.relative_to(target)))
            continue
        files.append((p, rel))
    return files


def route_of(path):
    ext = path.suffix.lower()
    if gaeb_reader.is_gaeb(path):
        return "gaeb"
    if ext in AZURE_TYPES:
        return "azure"
    return "unsupported"


# ------------------------------------------------------------------ Azure Content Understanding
class AzureCU:
    def __init__(self):
        self.endpoint = os.environ.get("AZURE_CU_ENDPOINT", "").strip().rstrip("/")
        # a Foundry *project* endpoint (…/api/projects/<name>) was pasted: keep only the resource part
        self.endpoint = re.sub(r"/api/projects/.*$", "", self.endpoint)
        self.key = os.environ.get("AZURE_CU_KEY")
        self.token = os.environ.get("AZURE_CU_TOKEN")
        self.analyzer = os.environ.get("AZURE_CU_ANALYZER", "prebuilt-layout")
        self.version = os.environ.get("AZURE_CU_API_VERSION", "2025-11-01")

    def check(self):
        missing = [n for n, v in (("AZURE_CU_ENDPOINT", self.endpoint),
                                  ("AZURE_CU_KEY or AZURE_CU_TOKEN", self.key or self.token)) if not v]
        if missing:
            raise SystemExit(f"Missing setting(s) for Azure Content Understanding: {', '.join(missing)} "
                             f"(put them in .env – see the header of step1_analyze.py)")

    def _headers(self, extra=None):
        h = {"Ocp-Apim-Subscription-Key": self.key} if self.key else {"Authorization": f"Bearer {self.token}"}
        h.update(extra or {})
        return h

    def _request(self, url, data=None, headers=None, method="GET"):
        for attempt in range(6):
            req = urllib.request.Request(url, data=data, headers=self._headers(headers), method=method)
            try:
                with urllib.request.urlopen(req, timeout=300) as r:
                    return r.status, dict(r.headers), r.read()
            except urllib.error.HTTPError as e:
                body = e.read().decode("utf-8", "replace")[:800]
                if e.code in (429, 500, 502, 503, 504) and attempt < 5:
                    time.sleep(float(e.headers.get("Retry-After") or 2 ** attempt * 3))
                    continue
                raise RuntimeError(f"HTTP {e.code}: {body}")
            except urllib.error.URLError as e:
                if attempt < 5:
                    time.sleep(2 ** attempt * 3)
                    continue
                raise RuntimeError(f"network error: {e.reason}")

    def analyze(self, path, content_type, max_wait=1800):
        url = (f"{self.endpoint}/contentunderstanding/analyzers/{self.analyzer}:analyzeBinary"
               f"?api-version={self.version}")
        status, headers, body = self._request(url, data=Path(path).read_bytes(), method="POST",
                                              headers={"Content-Type": content_type})
        op = headers.get("Operation-Location") or headers.get("operation-location")
        if not op:
            raise RuntimeError(f"no Operation-Location in the answer (HTTP {status}): {body[:300]!r}")
        start = time.time()
        while time.time() - start < max_wait:
            _, h, body = self._request(op)
            res = json.loads(body)
            st = (res.get("status") or "").lower()
            if st == "succeeded":
                return res
            if st in ("failed", "canceled", "cancelled"):
                raise RuntimeError(f"analysis {st}: {json.dumps(res.get('error', res))[:600]}")
            time.sleep(float(h.get("Retry-After") or 3))
        raise RuntimeError(f"no result after {max_wait} s")


# ------------------------------------------------------------------ one file
REBUILD_FEATURES = False          # set by --rebuild-features


def write_features(res, file_name, origin, out_dir, name, row):
    """Readable features file (name / type / description / value / evidence) next to the raw JSON.
    An existing features file is kept (not computed again) unless --rebuild-features or --force is given."""
    fp = Path(out_dir) / "features" / origin / f"{name}.features.json"
    if fp.exists() and not REBUILD_FEATURES:
        row["features_path"] = str(fp)
        if not row.get("positions"):
            try:
                row["positions"] = json.loads(fp.read_text(encoding="utf-8"))["summary"]["positions"]
            except (ValueError, KeyError):
                pass
        return
    try:
        feats = features.extract(res, file_name, origin)
    except Exception as e:                                     # features are a view: never stop the run
        row["message"] = (row["message"] + " | " if row["message"] else "") + f"features not written: {e}"[:200]
        return
    fp.parent.mkdir(parents=True, exist_ok=True)
    fp.write_text(json.dumps(feats, ensure_ascii=False, indent=2), encoding="utf-8")
    row["features_path"] = str(fp)
    sm = feats["summary"]
    row["positions"] = row.get("positions") or sm["positions"]
    if sm.get("total_check", "").startswith("positions add up to") or sm.get("price_check_failed"):
        note = sm.get("total_check") if "check" in sm.get("total_check", "") else \
            f'{sm["price_check_failed"]} position(s) where quantity x unit price ≠ total price'
        row["message"] = (row["message"] + " | " if row["message"] else "") + note


def process(path, rel, route, out_dir, cu, force):
    row = {"file": path.name, "relative_path": str(rel), "route": route, "positions": "", "message": ""}
    digest = sha256(path)
    row["sha256"] = digest
    if route == "unsupported":
        row["status"] = "unsupported"
        row["message"] = HINTS.get(path.suffix.lower(), "file type not supported by Azure or the GAEB reader")
        return row, None
    origin = "local" if route == "gaeb" else "azure"          # Azure AI  vs  read locally (GAEB / XML)
    name = f"{digest[:12]}__{safe(path.name)}"
    json_path = Path(out_dir) / "json" / origin / f"{name}.json"
    row["json_path"] = str(json_path)
    old = Path(out_dir) / "json" / f"{name}.json"               # files of the earlier version: move, don't re-pay
    if old.exists() and not json_path.exists():
        json_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(old), str(json_path))
    if json_path.exists() and not force:                       # cached: no new Azure cost
        res = json.loads(json_path.read_text(encoding="utf-8"))
        row["status"], row["message"] = "ok", "cached"
        row["positions"] = res.get("_positions", "")
        write_features(res, path.name, origin, out_dir, name, row)
        return row, res
    try:
        if route == "gaeb":
            res, n = gaeb_reader.read_gaeb(path)
            res["_positions"] = n
            row["positions"] = n
            row["message"] = "; ".join(res["result"].get("warnings", []))
        else:
            ctype, limit = AZURE_TYPES[path.suffix.lower()]
            if path.stat().st_size > limit:
                row["status"] = "too_large"
                row["message"] = f"{path.stat().st_size / 1e6:.1f} MB > Azure limit {limit / 1e6:.0f} MB for this type"
                return row, None
            res = cu.analyze(path, ctype)
        json_path.parent.mkdir(parents=True, exist_ok=True)
        json_path.write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
        row["status"] = "ok"
        write_features(res, path.name, origin, out_dir, name, row)
        return row, res
    except Exception as e:                                     # one bad file never stops the run
        row["status"], row["message"] = "failed", str(e)[:500]
        return row, None


# ------------------------------------------------------------------ main
def _test_pdf():
    """A tiny one-page PDF with a bill-of-quantities line, built in memory (no file needed)."""
    text = "BT (01.01.1 Kugelhahn DN 50 liefern und montieren) Tj 0 -20 Td (5,000 St EP ........ GP ........) Tj ET"
    stream = f"BT /F1 12 Tf 72 720 Td {text[3:]}".encode()
    objs = [b"<< /Type /Catalog /Pages 2 0 R >>", b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R "
            b"/Resources << /Font << /F1 5 0 R >> >> >>",
            b"<< /Length %d >>\nstream\n" % len(stream) + stream + b"\nendstream",
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"]
    out, offs = bytearray(b"%PDF-1.4\n"), []
    for i, o in enumerate(objs, 1):
        offs.append(len(out))
        out += b"%d 0 obj\n" % i + o + b"\nendobj\n"
    xref = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1)
    out += b"".join(b"%010d 00000 n \n" % o for o in offs)
    out += b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objs) + 1, xref)
    return bytes(out)


def check_azure():
    """Send a tiny test PDF to Azure Content Understanding and show what comes back."""
    import tempfile
    cu = AzureCU()
    print(f"endpoint   : {cu.endpoint}\nanalyzer   : {cu.analyzer}\napi-version: {cu.version}")
    with tempfile.TemporaryDirectory() as d:
        pdf = Path(d) / "azure_test.pdf"
        pdf.write_bytes(_test_pdf())
        t = time.time()
        try:
            res = cu.analyze(pdf, "application/pdf")
        except Exception as e:
            raise SystemExit(f"\nAZURE TEST FAILED: {e}\n"
                             "  401 -> wrong key | 404 -> wrong endpoint, or analyzer/api-version/region not supported")
    md = "\n".join(c.get("markdown", "") for c in res.get("result", {}).get("contents", []))
    print(f"status     : {res.get('status')}  ({time.time() - t:.1f} s)")
    print("text read  :", " | ".join(l for l in md.splitlines() if l.strip())[:200] or "(empty)")
    print("\nAZURE OK" if "Kugelhahn" in md else "\nAzure answered, but the test text was not found – check the analyzer")


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input", nargs="?", help="folder with the tender files (one sub-folder per project recommended)")
    ap.add_argument("--check-azure", action="store_true", help="only test the Azure connection with a tiny PDF and exit")
    ap.add_argument("--out", default="output", help="output folder (default: output)")
    ap.add_argument("--manifest", help="CSV with columns file,project_id (file = name or relative path)")
    ap.add_argument("--workers", type=int, default=4, help="files analysed in parallel (default 4)")
    ap.add_argument("--force", action="store_true", help="analyse again even if a cached JSON exists")
    ap.add_argument("--rebuild-features", action="store_true",
                    help="compute the features files again (no Azure cost) – e.g. after a code update")
    a = ap.parse_args(argv)
    global REBUILD_FEATURES
    REBUILD_FEATURES = a.rebuild_features or a.force
    load_dotenv()
    if a.check_azure:
        return check_azure()
    if not a.input:
        ap.error("the input folder is required (or use --check-azure)")

    input_dir, out_dir = Path(a.input), Path(a.out)
    if not input_dir.is_dir():
        raise SystemExit(f"Input folder not found: {input_dir.resolve()}")
    out_dir.mkdir(parents=True, exist_ok=True)
    manifest = {}
    if a.manifest:
        with open(a.manifest, newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                manifest[r["file"].strip().replace("\\", "/")] = r["project_id"].strip()

    files = scan(input_dir, out_dir)
    routed = [(p, rel, route_of(p)) for p, rel in files]
    print(f"{len(files)} file(s): " + ", ".join(f"{sum(r == k for *_, r in routed)} {k}"
                                                for k in ("azure", "gaeb", "unsupported")))
    cu = AzureCU()
    if any(r == "azure" for *_, r in routed):
        cu.check()
        print(f"Azure Content Understanding: analyzer '{cu.analyzer}', api-version {cu.version}")

    rows, results = [], {}
    with ThreadPoolExecutor(max_workers=max(1, a.workers)) as pool:
        futs = {pool.submit(process, p, rel, route, out_dir, cu, a.force): (p, rel) for p, rel, route in routed}
        for i, fut in enumerate(as_completed(futs), 1):
            row, res = fut.result()
            rows.append(row)
            results[row["relative_path"]] = res
            print(f"  [{i}/{len(futs)}] {row['status']:<11} {row['relative_path']}"
                  + (f"  ({row['message'][:90]})" if row["message"] and row["message"] != "cached" else ""))

    # ---- project assignment
    for row in rows:
        rel = row["relative_path"].replace("\\", "/")
        parts = rel.split("/")
        found = detected_project(results.get(row["relative_path"]) or {})
        if rel in manifest or row["file"] in manifest:
            row["project_id"], row["project_source"] = manifest.get(rel) or manifest[row["file"]], "manifest"
        elif len(parts) > 1:
            row["project_id"], row["project_source"] = parts[0], "folder"
            if found and found != parts[0]:
                row["message"] = (row["message"] + " | " if row["message"] else "") + \
                    f"file says project '{found}' but is in folder '{parts[0]}' – check"
        elif found:
            row["project_id"], row["project_source"] = found, "detected in file"
        else:
            row["project_id"], row["project_source"] = f"single:{Path(row['file']).stem}", "own project (unknown)"

    rows.sort(key=lambda r: (r["project_id"], r["relative_path"]))
    index = out_dir / "file_index.csv"
    with open(index, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=INDEX_COLUMNS, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)

    print("\nProjects:")
    for pid in sorted({r["project_id"] for r in rows}):
        rs = [r for r in rows if r["project_id"] == pid]
        ok = sum(r["status"] == "ok" for r in rs)
        print(f"  {pid:<30} {len(rs)} file(s), {ok} analysed, source: {rs[0]['project_source']}")
    bad = [r for r in rows if r["status"] != "ok"]
    if bad:
        print(f"\n{len(bad)} file(s) not analysed (see status/message in {index}):")
        for r in bad:
            print(f"  {r['status']:<11} {r['relative_path']}: {r['message'][:120]}")
    print(f"\nIndex written to {index}. Check the project_id column, then run step 2 (llm_filter.py).")


if __name__ == "__main__":
    main()