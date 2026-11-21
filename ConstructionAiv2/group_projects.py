"""
group_projects.py – find out which files belong to which project (step 1b)
==========================================================================
For files that are NOT in a project sub-folder (a flat folder with everything mixed):

  A. the LLM reads the beginning of every file (title page, headers, GAEB project data) and returns
     project number, project name, client, site – each with a VERBATIM quote from the file.
     Python checks every quote really is in the file; an invented number is thrown away.
  B. Python groups files with the same project number (spaces/dashes/case ignored).
     A file whose NAME contains a known project number joins that project.
  C. files still without a project (no number found): the LLM compares them with the known projects
     (name, client, site) and assigns one only with a verbatim proof; otherwise the file stays alone.

Never merged on a guess: an unclear file becomes its own project "single:<file>", so its materials
are never added to another project. Result: the project_id column of output/file_index.csv is
updated (project_source says how) – check it, change it by hand if needed, then run step 2.

    python group_projects.py --out output            (run_all.py calls this automatically)
"""
import argparse
import csv
import json
import re
from collections import Counter, defaultdict
from pathlib import Path

import llm_filter

PROFILE_CHARS = 3500        # beginning of the file the LLM reads (title page, header, project data)
BATCH_FILES = 8

EXTRACT_SCHEMA = {
    "type": "object", "additionalProperties": False, "required": ["files"],
    "properties": {"files": {"type": "array", "items": {
        "type": "object", "additionalProperties": False,
        "required": ["file_id", "project_number", "project_name", "client", "site", "document_type", "evidence"],
        "properties": {
            "file_id": {"type": "string"},
            "project_number": {"type": ["string", "null"]},
            "project_name": {"type": ["string", "null"]},
            "client": {"type": ["string", "null"]},
            "site": {"type": ["string", "null"]},
            "document_type": {"type": ["string", "null"]},
            "evidence": {"type": ["string", "null"]}}}}}}

ASSIGN_SCHEMA = {
    "type": "object", "additionalProperties": False, "required": ["assignments"],
    "properties": {"assignments": {"type": "array", "items": {
        "type": "object", "additionalProperties": False,
        "required": ["file_id", "project", "reason", "evidence"],
        "properties": {"file_id": {"type": "string"}, "project": {"type": "string"},
                       "reason": {"type": "string"}, "evidence": {"type": ["string", "null"]}}}}}}

EXTRACT_SYSTEM = """You read the beginning of construction tender documents (German: Leistungsverzeichnis, \
Baubeschreibung, Vorbemerkungen, GAEB, plans, letters) and identify the PROJECT each file belongs to.
For every file return:
- project_number: the project / tender / award number exactly as written (Projektnummer, Projekt-Nr., \
Vergabenummer, Vergabe-Nr., Maßnahmennummer, Ausschreibungsnummer, LV-Nummer, Bauvorhaben-Nr., GAEB LblPrj). \
null if none. NOT a date, phone number, page number, postcode, position number (01.02.3) or norm (DIN 1988).
- project_name: name of the construction project (Bauvorhaben / Maßnahme / Projekt), null if none.
- client: the client / Auftraggeber / Bauherr, null if none.
- site: the location / Baustelle / Ort, null if none.
- document_type: short (e.g. "LV", "Baubeschreibung", "Vertrag", "Plan").
- evidence: ONE short verbatim quote (max 150 characters) from the file text that contains the project \
number, or else the project name. Copy it exactly – it is checked automatically. null if nothing.
Never guess and never take values from another file."""

ASSIGN_SYSTEM = """You assign construction documents that have NO project number to known projects.
For every file decide:
- the id of a known project, ONLY if the file clearly names the same construction project \
(same project name, or same building AND same site). Same client alone is NOT enough – a client has many projects.
- "new:<short label>" if the file belongs to a project that is not in the list (files that clearly belong \
together get the SAME label),
- "none" if unclear. When in doubt answer "none": a wrong assignment mixes the materials of two projects.
evidence: one short verbatim quote (max 150 characters) from the FILE that proves the match; null for "none"."""


def key_of(number):
    """Comparison key of a project number: case, spaces, dashes, dots ignored."""
    return re.sub(r"[^0-9A-Z]", "", (number or "").upper())


def _norm(text):
    return re.sub(r"\s+", " ", text or "").strip().lower()


def grounded(snippet, text):
    return bool(snippet) and _norm(snippet) in _norm(text)


KEY_LINE = re.compile(r"projekt|ma(ß|ss)nahme|vergabe|bauvorhaben|bauherr|auftraggeber|baustelle|"
                      r"ausschreibung|LV[- ]?n(r|ummer)|objekt|project", re.I)


def key_lines(text, limit=1800):
    """Lines anywhere in the document that name the project (+ the 3 lines after: tables put the
    value in the next row) – the project number is often not on the first page."""
    lines, out, seen = text.splitlines(), [], set()
    for i, l in enumerate(lines):
        if KEY_LINE.search(l):
            chunk = " | ".join(x.strip() for x in lines[i:i + 4] if x.strip())
            if chunk and chunk not in seen:
                seen.add(chunk)
                out.append(chunk[:300])
        if sum(map(len, out)) > limit:
            break
    return "\n".join(out)


def profile(row, out_dir):
    """Text the LLM sees for a file: file name + project data fields + beginning of the content."""
    text, fields = "", {}
    jp = row.get("json_path")
    if row.get("status") == "ok" and jp:
        p = Path(jp) if Path(jp).is_absolute() else out_dir / jp
        if not p.exists():
            p = Path(jp)
        if p.exists():
            data = json.loads(p.read_text(encoding="utf-8"))
            for c in (data.get("result") or {}).get("contents", []):
                text += (c.get("markdown") or "") + "\n"
                for k, v in (c.get("fields") or {}).items():
                    val = v.get("valueString") if isinstance(v, dict) else None
                    if val:
                        fields[k] = val
    head = "\n".join(f"{k}: {v}" for k, v in fields.items())
    more = key_lines(text[PROFILE_CHARS:])
    return (f"FILE NAME: {row['file']}\n{head}\n{text[:PROFILE_CHARS]}"
            + (f"\n[… lines naming the project elsewhere in the file:]\n{more}" if more else "")).strip()


def extract(rows, texts, llm, log):
    """Step A: project data per file (LLM), every value checked against the file."""
    info = {}
    todo = rows                       # files Azure could not read are judged by their name only
    for i in range(0, len(todo), BATCH_FILES):
        batch = todo[i:i + BATCH_FILES]
        ids = {f"F{i + j + 1}": r for j, r in enumerate(batch)}
        user = "\n\n".join(f"### FILE {fid}\n{texts[r['relative_path']]}" for fid, r in ids.items())
        log(f"  reading project data: files {i + 1}-{i + len(batch)} of {len(todo)}")
        ans = llm.ask(EXTRACT_SYSTEM, user, EXTRACT_SCHEMA, "project_data")
        for a in ans.get("files", []):
            r = ids.get(a.get("file_id"))
            if not r:
                continue
            t = texts[r["relative_path"]]
            num = a.get("project_number")
            # the number must really be in the file (or its name) – otherwise it is dropped
            if num and key_of(num) and key_of(num) not in key_of(t):
                num = None
            if num and len(key_of(num)) < 3:
                num = None
            ok_ev = grounded(a.get("evidence"), t)
            info[r["relative_path"]] = {
                "number": num, "name": a.get("project_name") if ok_ev or num else None,
                "client": a.get("client"), "site": a.get("site"),
                "doc_type": a.get("document_type"), "evidence": a.get("evidence") if ok_ev else None}
    return info


def group(rows, info, texts, llm, log):
    """Steps B and C. Returns {relative_path: (project_id, source, note)}."""
    result, by_key, display = {}, defaultdict(list), defaultdict(Counter)
    for r in rows:
        inf = info.get(r["relative_path"]) or {}
        if inf.get("number"):
            k = key_of(inf["number"])
            by_key[k].append(r)
            display[k][inf["number"].strip()] += 1
    pid_of = {k: display[k].most_common(1)[0][0] for k in by_key}
    for k, rs in by_key.items():
        for r in rs:
            ev = (info[r["relative_path"]].get("evidence") or "")[:120]
            result[r["relative_path"]] = (pid_of[k], "project no. in file", f'"{ev}"' if ev else "")

    # B2: file name contains a known project number
    rest = [r for r in rows if r["relative_path"] not in result]
    for r in list(rest):
        fk = key_of(Path(r["file"]).stem)
        hits = [k for k in by_key if len(k) >= 5 and k in fk]
        if len(hits) == 1:
            result[r["relative_path"]] = (pid_of[hits[0]], "project no. in file name", "")
            rest.remove(r)

    # C: files without a number -> LLM compares with the known projects (proof required)
    rest_text = [r for r in rest if r.get("status") == "ok"]
    known = {}
    for k, rs in by_key.items():
        inf = [info[r["relative_path"]] for r in rs]
        pick = lambda f: Counter(x[f] for x in inf if x.get(f)).most_common(1)[0][0] if any(x.get(f) for x in inf) else ""
        known[pid_of[k]] = {"name": pick("name"), "client": pick("client"), "site": pick("site"),
                            "files": [r["file"] for r in rs][:5]}
    new_labels = {}
    for i in range(0, len(rest_text), BATCH_FILES):
        batch = rest_text[i:i + BATCH_FILES]
        ids = {f"F{i + j + 1}": r for j, r in enumerate(batch)}
        projects = "\n".join(f"- id: {pid} | name: {v['name']} | client: {v['client']} | site: {v['site']} | "
                             f"files: {', '.join(v['files'])}" for pid, v in known.items()) or "(none yet)"
        if new_labels:
            projects += "\n" + "\n".join(f"- id: new:{lab} | (project without number) | example: {ex}"
                                         for lab, ex in new_labels.items())
        user = f"KNOWN PROJECTS:\n{projects}\n\n" + "\n\n".join(
            f"### FILE {fid}\n{texts[r['relative_path']][:2000]}" for fid, r in ids.items())
        log(f"  matching files without project number: {i + 1}-{i + len(batch)} of {len(rest_text)}")
        ans = llm.ask(ASSIGN_SYSTEM, user, ASSIGN_SCHEMA, "project_assignment")
        for a in ans.get("assignments", []):
            r = ids.get(a.get("file_id"))
            proj = (a.get("project") or "none").strip()
            if not r or proj.lower() == "none":
                continue
            if not grounded(a.get("evidence"), texts[r["relative_path"]]):
                continue                                         # no verifiable proof -> stays alone
            note = f'{a.get("reason", "")[:100]} – "{a["evidence"][:100]}"'
            if proj.startswith("new:"):
                label = re.sub(r"[^\w.\-]+", "_", proj[4:].strip())[:50] or "project"
                new_labels.setdefault(label, r["file"])
                result[r["relative_path"]] = (f"auto:{label}", "LLM: same project (no number)", note)
            elif proj in known:
                result[r["relative_path"]] = (proj, "LLM: matched by name/site", note)
    # a "new:" label used by one file only is the same as an own project
    count = Counter(v[0] for v in result.values())
    for rel, (pid, src, note) in list(result.items()):
        if pid.startswith("auto:") and count[pid] == 1:
            del result[rel]
    return result


def run(out_dir, llm=None, log=print):
    out_dir = Path(out_dir)
    index = out_dir / "file_index.csv"
    with open(index, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    fixed = {"manifest", "folder"}
    todo = [r for r in rows if r["project_source"] not in fixed]
    if not todo:
        log("all files are in project folders (or in the manifest) – nothing to group")
        return rows
    log(f"grouping {len(todo)} file(s) without project folder into projects …")
    if llm is None:
        llm_filter.check_config()
        llm = llm_filter.LLM(cache=str(out_dir / ".llm_cache"))
    texts = {r["relative_path"]: profile(r, out_dir) for r in todo}
    info = extract(todo, texts, llm, log)
    groups = group(todo, info, texts, llm, log)
    for r in todo:
        if r["relative_path"] in groups:
            pid, src, note = groups[r["relative_path"]]
            r["project_id"], r["project_source"] = pid, src
            if note:
                r["message"] = (r["message"] + " | " if r.get("message") else "") + note[:200]
        else:
            r["project_id"], r["project_source"] = f"single:{Path(r['file']).stem}", "own project (unclear)"
    rows.sort(key=lambda r: (r["project_id"], r["relative_path"]))
    with open(index, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)

    log("\nProjects found:")
    for pid in sorted({r["project_id"] for r in rows}):
        rs = [r for r in rows if r["project_id"] == pid]
        srcs = ", ".join(sorted({r["project_source"] for r in rs}))
        log(f"  {pid:<30} {len(rs)} file(s)  [{srcs}]")
        for r in rs:
            log(f"      - {r['relative_path']}")
    unclear = [r for r in rows if r["project_source"] == "own project (unclear)"]
    if unclear:
        log(f"\n{len(unclear)} file(s) could not be assigned safely and stay their own project. If you know the "
            f"project, write it in the project_id column of {index} (or use a manifest).")
    log(f"\nIndex updated: {index}")
    return rows


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default="output", help="output folder of step 1 (contains file_index.csv)")
    ap.add_argument("--provider")
    ap.add_argument("--model")
    a = ap.parse_args(argv)
    llm_filter.load_dotenv()
    llm = None
    if a.provider or a.model:
        llm_filter.check_config()
        llm = llm_filter.LLM(cache=str(Path(a.out) / ".llm_cache"), provider=a.provider, model=a.model)
    run(a.out, llm)


if __name__ == "__main__":
    main()
