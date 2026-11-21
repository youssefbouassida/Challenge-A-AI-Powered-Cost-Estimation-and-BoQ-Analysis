"""
extract_to_excel.py  –  FILE 2 of 2: deterministic number extraction + Excel
=============================================================================
Input : materials_filtered.json written by llm_filter.py
Output: Excel sheet  Project | Material | Quantity | Unit | Unit price | Total price | Description
        (+ a "Sources" sheet to trace every row back to the document)

No LLM here. Every number is parsed by Python from the VERBATIM snippets the LLM spotted:
  1. the snippet must really occur in the source text, otherwise it is discarded
  2. numbers are parsed with explicit locale rules (German: 1.234,56 ; English: 1,234.56)
  3. amounts written "per unit" (e.g. 8.300 kg per filter) are multiplied by the position
     quantity (e.g. 2 St filters) – the only calculation, done here in Python, not by the LLM
  4. if the LLM gave no quantity snippet, the position's quantity line ("175 St EP GP") is
     read from the source text as a deterministic fallback
  5. rows with the same material name, unit AND the same sizes/angles in the position title are
     merged (quantities summed). Same name but different sizes (e.g. 90° vs 45° branch) are kept
     apart and the position number is added -> material names are unique per project
  6. a missing value stays EMPTY – nothing is guessed; total = quantity × unit price only
     when both are known (written as an Excel formula)
Sheets: Materials | Sources (row -> position, snippets, notes) | Checks (things to review) |
        Not materials (LLM's reasons)

Usage
    python extract_to_excel.py materials_filtered.json -o materials.xlsx [--number-format de|en]
"""
import argparse
import json
import re
from collections import OrderedDict
from pathlib import Path

from features import position_kind

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

UNITS = {"st": "St", "stk": "St", "stück": "St", "pcs": "pcs", "pc": "pcs", "ea": "pcs", "m": "m",
         "lfm": "m", "m2": "m2", "m²": "m2", "qm": "m2", "m3": "m3", "m³": "m3", "cbm": "m3", "kg": "kg",
         "t": "t", "l": "l", "psch": "psch", "pauschal": "psch", "satz": "Satz", "set": "set",
         "gebinde": "drum", "h": "h", "std": "h"}
NUM = r"\d{1,3}(?:[.,' \u00a0]\d{3})+(?:[.,]\d+)?|\d+(?:[.,]\d+)?"
BOQ_QTY = re.compile(rf"(?<![\w,.])({NUM})\s*({'|'.join(sorted(map(re.escape, UNITS), key=len, reverse=True))})\b\.?[\s.]{{0,12}}(?:EP|GP)\b",
                     re.I)


# ------------------------------------------------------------------ number parsing
def parse_number(text, fmt="de"):
    """First number in text -> float, with explicit locale rules. None if no number."""
    if not text:
        return None
    m = re.search(NUM, text)
    if not m:
        return None
    s = m.group(0).replace(" ", "").replace("\u00a0", "").replace("'", "")
    dec, thou = (",", ".") if fmt == "de" else (".", ",")
    if dec in s and thou in s:                        # 1.234,56  /  1,234.56
        s = s.replace(thou, "").replace(dec, ".")
    elif thou in s:
        parts = s.split(thou)
        # "8.300" (de) = 8300 ; "1.5" (de) is not a valid thousands group -> treat as decimal
        s = s.replace(thou, "") if all(len(p) == 3 for p in parts[1:]) else s.replace(thou, ".")
    else:
        s = s.replace(dec, ".")
    try:
        return float(s)
    except ValueError:
        return None


def parse_unit(text):
    if not text:
        return None
    m = re.search(rf"(?:{NUM})\s*([A-Za-zäöüÄÖÜ]+[23²³]?)", text)
    return UNITS.get(m.group(1).lower(), m.group(1)) if m else None


def _in_source(snippet, source):
    norm = lambda s: re.sub(r"\s+", " ", s or "").strip().lower()
    return bool(snippet) and norm(snippet) in norm(source)


def _numbers(text):
    text = re.sub(r"\[[A-Z]+_\d+\]", "", text)                 # anonymisation tokens
    text = re.sub(r"^[\s#]*\d{1,2}(?:\.\d{1,4}){1,3}", "", text)  # the position number itself
    return tuple(sorted(set(re.findall(r"\d+(?:[.,/]\d+)?", text))))


def _signature(e):
    """Numbers in the position title (sizes, angles, DN, PN …). Two positions are only merged when
    these are identical. If the title has NO numbers (e.g. hidden by [PERSON_…] tokens), the numbers
    of the whole position text are used instead (fix: 1.5 kW and 5.5 kW pumps were merged)."""
    src = e.get("source_text", "")
    title = _numbers("\n".join(src.split("\n")[:3]))
    return ("title",) + title if title else ("text",) + _numbers(src)


# ------------------------------------------------------------------ row extraction
def extract_row(e, fmt):
    notes = []
    src = e.get("source_text", "")
    snips = {k: e.get(k) for k in ("quantity_text", "position_quantity_text", "unit_price_text", "total_price_text")}
    for k, val in snips.items():
        if val and not re.search(r"\d", val):
            snips[k] = None                            # no number at all, e.g. "EP." = empty price column
        elif val and not _in_source(val, src):         # grounding: snippet must exist in the source
            notes.append(f"LLM snippet '{val}' not found in source – ignored")
            snips[k] = None
    qty, unit = parse_number(snips["quantity_text"], fmt), parse_unit(snips["quantity_text"])
    if e.get("quantity_is_per_unit"):
        n = parse_number(snips["position_quantity_text"], fmt)
        if qty is not None and n is not None:
            notes.append(f"per-unit amount: {snips['quantity_text']} × {snips['position_quantity_text']}")
            qty = qty * n
        else:
            notes.append("per-unit amount but position quantity unknown – quantity left empty")
            qty = None
    if qty is None and not e.get("quantity_is_per_unit"):     # fallback: the position's quantity line
        hits = BOQ_QTY.findall(src)
        if hits:
            qty, unit = parse_number(hits[0][0], fmt), UNITS.get(hits[0][1].lower(), hits[0][1])
            notes.append(f"quantity from the position's quantity line '{hits[0][0]} {hits[0][1]}'")
    up, tp = parse_number(snips["unit_price_text"], fmt), parse_number(snips["total_price_text"], fmt)
    pos = e.get("position") or e.get("source_ref") or ""
    name = re.sub(r"\s+", " ", e["material_name"]).strip()
    if e.get("conflict"):                 # two documents give different quantities for this position
        name += f' [version {Path(str(e.get("file") or "?")).name}]'
        notes.append("CONFLICT: another document of this project gives a different quantity for this position")
    return {"project_id": e["project_id"], "material": name,
            "quantity": qty, "unit": unit, "unit_price": up, "total_price": tp,
            "description": e.get("description") or None, "position": pos, "signature": _signature(e),
            "source": f'{e.get("file")} / {pos}',
            "quotes": [snips["quantity_text"], snips["unit_price_text"], snips["total_price_text"]],
            "notes": notes}


AMOUNT = re.compile(rf"(?<![\w,.])({NUM})\s*(kg|t|m3|m³|l)\b", re.I)


def kind_checks(entries):
    """Materials from contingency / alternative positions: listed in 'Checks' (not a normal purchase)."""
    out, seen = [], set()
    for e in entries:
        k = position_kind(e.get("source_text"))
        key = (e["project_id"], e.get("position"))
        if k and key not in seen:
            seen.add(key)
            out.append({"project_id": e["project_id"], "position": e.get("position"),
                        "check": "special position", "detail": f'{k}: "{e.get("material_name", "")[:80]}"'})
    return out


def position_checks(entries, fmt):
    """Deterministic warnings per position (shown in sheet 'Checks'):
    a position whose amounts are given per unit (e.g. per filter) and that contains further amounts
    in kg / t / m3 / l that no extracted item uses -> a product may have been missed
    (fix: the gravel '1,8 m3' of the filter filling was dropped)."""
    out, by_pos = [], OrderedDict()
    for e in entries:
        by_pos.setdefault((e["project_id"], e.get("position")), []).append(e)
    for (pid, pos), es in by_pos.items():
        if not any(e.get("quantity_is_per_unit") for e in es):
            continue
        used = {parse_number(e.get("quantity_text"), fmt) for e in es}
        src = es[0].get("source_text", "")
        for m in AMOUNT.finditer(src):
            val = parse_number(m.group(1), fmt)
            if re.search(r"\bje\s*$", src[max(0, m.start() - 6):m.start()]):
                continue                               # "Big Bags von je 830 kg" = packaging size
            if val not in used:
                out.append({"project_id": pid, "position": pos, "check": "possible missing product",
                            "detail": f"amount '{m.group(0)}' in this position is not used by any row – "
                                      f"is it a separate product to buy?"})
    return out


def merge_rows(rows):
    """One row per material name per project. Rows are summed only when name, unit and the
    position-title sizes (signature) are identical; otherwise the name is made unique."""
    groups = OrderedDict()
    for r in rows:
        groups.setdefault((r["project_id"], r["material"].lower(), r["unit"], r["signature"]), []).append(r)
    # names that would collide (same name, different unit or size) get a distinguishing suffix
    by_name = {}
    for (pid, name, unit, sig) in groups:
        by_name.setdefault((pid, name), []).append((unit, sig))
    out = []
    for (pid, name, unit, sig), parts in groups.items():
        r = dict(parts[0])
        if len(by_name[(pid, name)]) > 1:
            units = {u for u, _ in by_name[(pid, name)]}
            suffix = f"[{unit or 'no unit'}]" if len(units) > 1 else ""
            same_unit = [s_ for u, s_ in by_name[(pid, name)] if u == unit]
            if len(same_unit) > 1:
                suffix += f" (pos. {', '.join(p['position'] for p in parts)})"
            r["material"] = f'{r["material"]} {suffix.strip()}'.strip()
        qs = [p["quantity"] for p in parts]
        r["quantity"] = sum(qs) if None not in qs else None           # one unknown -> total unknown
        ups = {p["unit_price"] for p in parts}
        r["unit_price"] = ups.pop() if len(ups) == 1 else None         # different prices -> no single price
        tps = [p["total_price"] for p in parts]
        r["total_price"] = sum(tps) if None not in tps else None
        descs = list(dict.fromkeys(p["description"] for p in parts if p["description"]))
        r["description"] = "; ".join(descs) or None
        r["parts"] = parts
        out.append(r)
    return out


# ------------------------------------------------------------------ excel
def write_excel(rows, path, rejected=(), not_reviewed=(), checks=(), projects=None, files=()):
    wb = Workbook()
    ws = wb.active
    projects = projects or {}
    if projects:                                    # overview first
        pv = ws
        pv.title = "Projects"
        pv.append(["Project", "Materials (rows)", "Material positions", "Rejected positions", "Not reviewed",
                   "Checks", "Duplicates removed", "Conflicts", "Documents used", "Documents without positions (not used)"])
        for c in pv[1]:
            c.font = Font(name="Arial", bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F3864")
        for pid in sorted(projects):
            info = projects[pid]
            pv.append([pid, sum(1 for r in rows if r["project_id"] == pid),
                       len({p["position"] for r in rows if r["project_id"] == pid for p in r["parts"]}),
                       sum(1 for r in rejected if r["project_id"] == pid),
                       sum(1 for r in not_reviewed if r["project_id"] == pid),
                       sum(1 for c in checks if c["project_id"] == pid),
                       info.get("duplicates_removed", 0), len(info.get("conflicts", [])),
                       "\n".join(f for f in info.get("files", []) if f not in info.get("skipped_documents", [])),
                       "\n".join(info.get("skipped_documents", []))])
        for col, w in zip("ABCDEFGHIJ", (22, 10, 10, 10, 10, 8, 11, 9, 50, 45)):
            pv.column_dimensions[col].width = w
        for row in pv.iter_rows(min_row=2):
            for c in row:
                c.alignment = Alignment(vertical="top", wrap_text=True)
        ws = wb.create_sheet("Materials")
    else:
        ws.title = "Materials"
    head = ["Project", "Material", "Quantity", "Unit", "Unit price (EUR)", "Total price (EUR)", "Description"]
    ws.append(head)
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3864")
    for i, r in enumerate(sorted(rows, key=lambda r: (r["project_id"], r["material"].lower())), 2):
        if r["total_price"] is not None:
            total = r["total_price"]                                      # written in the document
        elif r["quantity"] is not None and r["unit_price"] is not None:
            total = f"=C{i}*E{i}"                                         # computed, stays live
        else:
            total = None                                                  # missing stays missing
        ws.append([r["project_id"], r["material"], r["quantity"], r["unit"], r["unit_price"], total,
                   r["description"]])
        for c in ws[i]:
            c.font = Font(name="Arial", size=10)
            c.alignment = Alignment(vertical="top", wrap_text=c.column == 7)
        ws.cell(i, 3).number_format = "#,##0.###"
        ws.cell(i, 5).number_format = ws.cell(i, 6).number_format = '#,##0.00 "€"'
    for col, w in zip("ABCDEFG", (14, 46, 11, 7, 15, 16, 70)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions

    src = wb.create_sheet("Sources")
    src.append(["Project", "Material", "Source (file / position)", "Quantity snippet", "Unit price snippet",
                "Total price snippet", "Notes"])
    for c in src[1]:
        c.font = Font(name="Arial", bold=True)
    for r in rows:
        for p in r["parts"]:
            src.append([r["project_id"], r["material"], p["source"], *p["quotes"], " | ".join(p["notes"])])
    for col, w in zip("ABCDEFG", (14, 46, 40, 18, 16, 16, 60)):
        src.column_dimensions[col].width = w

    ck = wb.create_sheet("Checks")
    ck.append(["Project", "Position", "Check", "Detail"])
    for c in ck[1]:
        c.font = Font(name="Arial", bold=True)
    for pid, info in sorted(projects.items()):
        for cf in info.get("conflicts", []):
            ck.append([pid, cf["position"], "CONFLICT between documents", f'{cf["detail"]} ({cf["files"]})'])
            for c in ck[ck.max_row]:
                c.fill = PatternFill("solid", fgColor="FFF2CC")
        for x in info.get("same_lv", []):
            ck.append([pid, "", "same LV in two files", x + " – duplicates removed, GAEB version kept"])
    for c_ in checks:
        ck.append([c_["project_id"], c_["position"], c_["check"], c_["detail"]])
    for r in rows:
        if len(r["parts"]) > 1:
            ck.append([r["project_id"], ", ".join(p["position"] for p in r["parts"]), "positions summed",
                       f'"{r["material"]}" = sum of {len(r["parts"])} positions – check they are the same product'])
        for p in r["parts"]:
            for n in p["notes"]:
                ck.append([r["project_id"], p["position"], "note", n])
    if ck.max_row == 1:
        ck.append(["", "", "no checks", "nothing to review"])
    for col, w in zip("ABCD", (14, 22, 26, 100)):
        ck.column_dimensions[col].width = w

    nm = wb.create_sheet("Not materials")
    nm.append(["Project", "Position", "Status", "Reason (LLM)", "Text (start)"])
    for c in nm[1]:
        c.font = Font(name="Arial", bold=True)
    for r in not_reviewed:
        nm.append([r["project_id"], r["position"], "NOT REVIEWED – check manually", "", r.get("text", "")[:200]])
    for r in sorted(rejected, key=lambda r: (not r.get("check"), r["position"])):     # flagged ones first
        status = r["check"] if r.get("check") else "rejected"
        if r.get("check") and r.get("second_opinion"):
            status += " – re-checked by the LLM, still rejected: please glance at it"
        nm.append([r["project_id"], r["position"], status, r.get("reason", ""), r.get("text", "")[:200]])
        if r.get("check"):
            for c in nm[nm.max_row]:
                c.fill = PatternFill("solid", fgColor="FFF2CC")
    for col, w in zip("ABCDE", (14, 12, 45, 50, 80)):
        nm.column_dimensions[col].width = w

    if files:                                        # every input file and what happened to it
        fs = wb.create_sheet("Files")
        cols = ["project_id", "project_source", "relative_path", "route", "status", "positions", "message"]
        fs.append(["Project", "Project from", "File", "Route", "Status", "GAEB positions", "Message"])
        for c in fs[1]:
            c.font = Font(name="Arial", bold=True)
        for f in files:
            fs.append([f.get(k, "") for k in cols])
            if f.get("status") != "ok":
                for c in fs[fs.max_row]:
                    c.fill = PatternFill("solid", fgColor="FFF2CC")
        for col, w in zip("ABCDEFG", (22, 18, 50, 11, 12, 10, 70)):
            fs.column_dimensions[col].width = w
    if "Projects" in wb.sheetnames:                  # material list first (like before), overview at the end
        pv = wb["Projects"]
        wb.move_sheet(pv, offset=len(wb.sheetnames) - 1 - wb.sheetnames.index("Projects"))
        if "Files" in wb.sheetnames:
            wb.move_sheet(pv, offset=-1)
    wb.active = wb.sheetnames.index("Materials")
    for w in wb.worksheets:
        w.sheet_view.tabSelected = w.title == "Materials"
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def _F(name, type_, desc, value, evidence=None):
    return {"name": name, "type": type_, "description": desc, "value": value, "evidence": evidence}


def write_material_files(entries, rejected, not_reviewed, files, out_dir, fmt):
    """One readable materials file per document (same format as the features files):
    output/materials/azure/<doc>.materials.json  and  output/materials/local/<doc>.materials.json.
    Numbers are the same Python-checked values as in the Excel; rows are NOT yet summed per project."""
    by_file = {}
    for f in files:
        by_file[f.get("relative_path") or f.get("file")] = f
    docs = OrderedDict()
    for kind, xs in (("materials", entries), ("not_materials", rejected), ("not_reviewed", not_reviewed)):
        for x in xs:
            docs.setdefault(x.get("file") or "?", {"materials": [], "not_materials": [], "not_reviewed": []})[kind].append(x)
    written = []
    for rel, d in docs.items():
        info = by_file.get(rel, {})
        origin = "local" if info.get("route") == "gaeb" else "azure"
        stem = Path(info["features_path"]).name.replace(".features.json", "") if info.get("features_path") \
            else re.sub(r"[^\w.\-]+", "_", Path(rel).name)
        mats = []
        for e in d["materials"]:
            r = extract_row(e, fmt)
            q_txt, up_txt, tp_txt = r["quotes"]
            val = [_F("Project", "string", "Project the document belongs to", r["project_id"]),
                   _F("PositionNumber", "string", "Position (OZ) the material comes from", r["position"]),
                   _F("MaterialName", "string", "Name of the material to buy (LLM)", r["material"]),
                   _F("Description", "string", "Short description (LLM)", r["description"]),
                   _F("Quantity", "number", "Quantity to buy (extracted by Python, checked against the text)",
                      r["quantity"], q_txt),
                   _F("Unit", "string", "Unit of the quantity", r["unit"], q_txt),
                   _F("UnitPrice", "number", "Unit price in EUR (EP), empty if not in the document", r["unit_price"], up_txt),
                   _F("TotalPrice", "number", "Total price in EUR (GP), empty if not in the document", r["total_price"], tp_txt),
                   _F("PerUnitAmount", "string", "Amount given per unit (e.g. per filter) and multiplied by the "
                      "position quantity", e.get("quantity_is_per_unit") and
                      f'{e.get("quantity_text")} × {e.get("position_quantity_text")}' or None),
                   _F("PositionKind", "string", "Special kind of position; empty = normal position. A contingency "
                      "is only bought if needed, an alternative replaces another position",
                      position_kind(e.get("source_text"))),
                   _F("Conflict", "string", "Another document of the project gives a different quantity",
                      "yes – check" if e.get("conflict") else None),
                   _F("Notes", "array", "Checks done by Python", r["notes"] or None)]
            mats.append(_F(f'Material {r["position"]} – {r["material"]}', "object",
                           "Material to buy (decision by the LLM, numbers by Python)", val,
                           (e.get("source_text") or "")[:300]))
        rej = [_F(f'Position {x.get("position")}', "object", "Position that is NOT a material to buy",
                  [_F("PositionNumber", "string", "Position (OZ)", x.get("position")),
                   _F("Reason", "string", "Why the LLM rejected it", x.get("reason")),
                   _F("Check", "string", "Supply wording found – please glance at it", x.get("check") or None)],
                  (x.get("text") or "")[:300]) for x in d["not_materials"]]
        nrev = [_F(f'Position {x.get("position")}', "object", "Position the LLM could not review", None,
                   (x.get("text") or "")[:300]) for x in d["not_reviewed"]]
        out = {"document": rel, "project_id": info.get("project_id") or (d["materials"] or d["not_materials"] or
                                                                       d["not_reviewed"])[0].get("project_id"),
               "source": "Azure Content Understanding" if origin == "azure" else "local GAEB reader",
               "summary": {"positions_reviewed": len({e.get("position") for e in d["materials"]}) + len(d["not_materials"]),
                           "material_rows": len(mats), "not_materials": len(rej), "not_reviewed": len(nrev),
                           "with_quantity": sum(1 for m in mats if m["value"][4]["value"] is not None),
                           "with_unit_price": sum(1 for m in mats if m["value"][6]["value"] is not None),
                           "conflicts": sum(1 for e in d["materials"] if e.get("conflict"))},
               "materials": mats, "not_materials": rej, "not_reviewed": nrev}
        path = Path(out_dir) / "materials" / origin / f"{stem}.materials.json"
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
        written.append(path)
    return written


def safe_name(pid):
    return re.sub(r"[^\w.\-]+", "_", pid)[:80] or "project"


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input", help="materials_filtered.json from llm_filter.py")
    ap.add_argument("-o", "--output", help="combined workbook (default: <out-dir>/all_projects.xlsx)")
    ap.add_argument("--out-dir", help="folder for the workbooks (default: folder of the input JSON)")
    ap.add_argument("--no-per-project", action="store_true", help="only the combined workbook")
    ap.add_argument("--no-excel", action="store_true",
                    help="only the JSON files output/materials/azure|local/*.materials.json, no Excel")
    ap.add_argument("--number-format", choices=["de", "en"], default="de",
                    help="de: 1.234,56 (default, German tenders)  en: 1,234.56")
    a = ap.parse_args(argv)
    data = json.loads(Path(a.input).read_text(encoding="utf-8"))
    if isinstance(data, list):                      # older format: list of materials only
        data = {"materials": data, "not_materials": [], "not_reviewed": []}
    out_dir = Path(a.out_dir) if a.out_dir else Path(a.input).parent
    combined = Path(a.output) if a.output else out_dir / "all_projects.xlsx"
    entries = data["materials"]
    rows = merge_rows([extract_row(e, a.number_format) for e in entries])   # merge key contains the project
    checks = position_checks(entries, a.number_format) + kind_checks(entries)
    rejected, not_reviewed = data.get("not_materials", []), data.get("not_reviewed", [])
    projects, files = data.get("projects", {}), data.get("files", [])
    for pid in {r["project_id"] for r in rows} | {r["project_id"] for r in rejected}:
        projects.setdefault(pid, {"files": []})

    written = write_material_files(entries, rejected, not_reviewed, files, out_dir, a.number_format)
    n_az = sum(1 for w in written if w.parent.name == "azure")
    print(f"materials per document: {n_az} file(s) in {out_dir / 'materials' / 'azure'}, "
          f"{len(written) - n_az} in {out_dir / 'materials' / 'local'}")
    if a.no_excel:
        return
    write_excel(rows, combined, rejected, not_reviewed, checks, projects, files)
    print(f"combined workbook (all projects, rows never merged across projects): {combined}")
    if not a.no_per_project and len(projects) > 0:
        for pid in sorted(projects):
            sel = lambda xs: [x for x in xs if x["project_id"] == pid]
            path = out_dir / "excel" / f"{safe_name(pid)}.xlsx"
            write_excel(sel(rows), path, sel(rejected), sel(not_reviewed), sel(checks), {pid: projects[pid]},
                        [f for f in files if f.get("project_id") == pid])
            print(f"  {pid:<28} {len(sel(rows)):>4} materials -> {path}")
    miss = lambda k: sum(r[k] is None for r in rows)
    print(f"{len(entries)} entries -> {len(rows)} materials in {len({r['project_id'] for r in rows})} project(s)")
    print(f"missing: quantity {miss('quantity')}, unit price {miss('unit_price')}, total price {miss('total_price')}")
    summed = sum(1 for r in rows if len(r["parts"]) > 1)
    conflicts = sum(len(p.get("conflicts", [])) for p in projects.values())
    if checks or summed or conflicts:
        print(f"CHECK: {conflicts} conflict(s) between documents, {len(checks)} possible missing product(s), "
              f"{summed} row(s) summing several positions – see sheet 'Checks'")
    flagged = sum(1 for r in rejected if r.get("check"))
    if flagged:
        print(f"CHECK: {flagged} rejected position(s) contain supply wording – highlighted in sheet 'Not materials'")
    bad_files = [f for f in files if f.get("status") != "ok"]
    if bad_files:
        print(f"FILES: {len(bad_files)} file(s) not analysed – see sheet 'Files'")
    print(f"not materials: {len(rejected)} positions, not reviewed by the LLM: {len(not_reviewed)}")


if __name__ == "__main__":
    main()