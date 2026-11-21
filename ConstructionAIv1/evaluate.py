"""
evaluate.py  –  OPTIONAL test tool (not part of the 2-file pipeline)
====================================================================
Scores a materials.xlsx produced by the pipeline against a VERIFIED answer ("gold standard"),
so different models / prompts can be compared with numbers instead of by eye.

Built-in gold standard: project HEM2301-2 (HEM2301-2_Gesamt-LV.pdf), checked position by position:
127 material positions with their quantities, 32 positions that are NOT materials.

Usage
    python evaluate.py materials.xlsx                         # score against the built-in gold
    python evaluate.py materials.xlsx --gold my_gold.json     # your own verified documents
    python evaluate.py --export-gold my_gold.json             # write the built-in gold as a template

Gold file format (one entry per project id):
    {"PROJECT-ID": {"materials": {"05.01.36": [[175, "St"]], "03.05.2": [[16600, "kg"], [3.6, "m3"]]},
                    "not_materials": ["01.01.1", "01.02.1", ...]}}
"""
import argparse
import json
import re
from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook

GOLD = json.loads(r'''{"HEM2301-2":{"source":"HEM2301-2_Gesamt-LV.pdf, verified position by position","materials":{"01.01.7":[[200,"St"]],"01.03.10":[[100,"kg"]],"03.01.1":[[4,"St"]],"03.01.2":[[4,"St"]],"03.01.3":[[1,"St"]],"03.01.4":[[1,"St"]],"03.01.5":[[2,"St"]],"03.02.1":[[2,"St"]],"03.02.2":[[1,"St"]],"03.02.3":[[2,"St"]],"03.02.4":[[2,"St"]],"03.02.5":[[1,"St"]],"03.03.1":[[1,"St"]],"03.03.2":[[1,"St"]],"03.04.1":[[1,"psch"]],"03.05.1":[[820,"St"]],"03.05.2":[[16600,"kg"],[3.6,"m3"]],"04.01.1":[[6,"St"]],"04.02.1":[[6,"St"]],"04.02.2":[[5,"St"]],"05.01.1":[[20,"m"]],"05.01.2":[[5,"m"]],"05.01.3":[[25,"m"]],"05.01.4":[[5,"m"]],"05.01.5":[[100,"m"]],"05.01.6":[[100,"m"]],"05.01.7":[[2,"St"]],"05.01.8":[[1,"St"]],"05.01.9":[[25,"St"]],"05.01.10":[[10,"St"]],"05.01.11":[[60,"St"]],"05.01.12":[[35,"St"]],"05.01.13":[[2,"St"]],"05.01.14":[[5,"St"]],"05.01.15":[[6,"St"]],"05.01.16":[[1,"St"]],"05.01.17":[[1,"St"]],"05.01.18":[[10,"St"]],"05.01.19":[[12,"St"]],"05.01.20":[[1,"St"]],"05.01.21":[[2,"St"]],"05.01.22":[[2,"St"]],"05.01.23":[[2,"St"]],"05.01.24":[[2,"St"]],"05.01.25":[[1,"St"]],"05.01.26":[[5,"St"]],"05.01.27":[[12,"St"]],"05.01.28":[[5,"St"]],"05.01.29":[[1,"St"]],"05.01.30":[[3,"St"]],"05.01.31":[[6,"St"]],"05.01.32":[[32,"St"]],"05.01.33":[[5,"St"]],"05.01.34":[[60,"St"]],"05.01.35":[[15,"St"]],"05.01.36":[[175,"St"]],"05.01.37":[[130,"St"]],"05.01.38":[[1,"St"]],"05.01.39":[[1,"St"]],"05.01.40":[[1,"St"]],"05.01.41":[[1,"St"]],"05.01.42":[[15,"St"]],"05.01.43":[[1,"St"]],"05.01.44":[[1,"St"]],"05.01.45":[[1,"St"]],"05.01.46":[[1,"St"]],"05.01.47":[[2,"St"]],"05.01.48":[[2,"St"]],"05.01.49":[[10,"St"]],"05.01.50":[[10,"St"]],"05.01.51":[[3,"St"]],"05.01.52":[[700,"kg"]],"05.01.53":[[200,"kg"]],"05.01.54":[[30,"St"]],"05.01.55":[[10,"St"]],"05.01.56":[[2,"St"]],"05.01.57":[[10,"St"]],"05.01.58":[[10,"St"]],"05.01.59":[[2,"St"]],"05.01.60":[[10,"St"]],"05.01.61":[[5,"St"]],"05.01.62":[[5,"St"]],"05.01.63":[[2,"St"]],"05.02.1":[[100,"m"]],"05.02.2":[[5,"St"]],"05.02.3":[[10,"m"]],"05.02.4":[[30,"m"]],"05.02.5":[[5,"m"]],"05.02.6":[[30,"m"]],"05.02.7":[[5,"St"]],"05.02.8":[[5,"St"]],"05.02.9":[[5,"St"]],"05.02.10":[[2,"St"]],"05.02.11":[[10,"St"]],"05.03.1":[[2,"St"]],"05.03.2":[[5,"St"]],"05.03.3":[[2,"St"]],"05.03.4":[[1,"St"]],"06.01.1":[[4,"St"]],"06.01.2":[[2,"St"]],"06.01.3":[[3,"St"]],"06.01.4":[[1,"St"]],"06.01.7":[[8,"St"]],"06.01.8":[[8,"St"]],"06.02.1":[[5,"St"]],"06.02.4":[[2,"St"]],"06.03.1":[[10,"St"]],"06.03.2":[[5,"St"]],"06.03.3":[[5,"St"]],"06.03.4":[[1,"St"]],"06.03.5":[[10,"St"]],"06.03.6":[[2,"St"]],"06.03.7":[[20,"St"]],"06.03.8":[[2,"St"]],"06.03.9":[[2,"St"]],"06.03.10":[[1,"St"]],"06.03.11":[[1,"St"]],"06.03.12":[[10,"St"]],"06.04.1":[[2,"St"]],"06.04.2":[[4,"St"]],"06.04.3":[[1,"St"]],"06.04.4":[[1,"St"]],"06.04.5":[[1,"St"]],"06.05.1":[[3,"St"]],"07.01.1":[[1,"St"]],"07.01.2":[[1,"St"]],"07.01.3":[[2,"St"]]},"not_materials":["01.01.1","01.01.2","01.01.3","01.01.4","01.01.5","01.01.6","01.01.8","01.02.1","01.02.2","01.02.3","01.02.4","01.02.5","01.03.1","01.03.2","01.03.3","01.03.4","01.03.5","01.03.6","01.03.7","01.03.8","01.03.9","02.01.1","02.02.1","02.02.2","02.02.3","03.05.3","03.05.4","03.05.5","06.01.5","06.01.6","06.02.2","06.02.3"]}}''')


def read_result(path):
    """materials.xlsx -> {project: {position: [(row_name, row_qty, row_unit, n_positions_in_row)]}}"""
    wb = load_workbook(path, data_only=True)
    rows = {(r[0], r[1]): r for r in wb["Materials"].iter_rows(min_row=2, values_only=True) if r[1]}
    parts = defaultdict(list)                                     # (project, name) -> positions
    for r in wb["Sources"].iter_rows(min_row=2, values_only=True):
        if r[1]:
            pos = str(r[2]).split(" / ")[-1].replace("text:", "")
            if re.fullmatch(r"\d{1,2}(?:\.\d{1,4}){1,3}", pos):          # ignore rows without a position
                parts[(r[0], r[1])].append(pos)
            continue
        if r[1]:
            parts.setdefault((r[0], r[1]), [])
    out = defaultdict(lambda: defaultdict(list))
    out["__rows_without_position__"] = [k for k in rows if not parts.get(k)]
    for (pid, name), poss in parts.items():
        row = rows.get((pid, name))
        if row and poss:
            for p in poss:
                out[pid][p].append((name, row[2], row[3], poss))
    return out


def close(a, b):
    return a is not None and b is not None and abs(float(a) - float(b)) <= 1e-6 * max(1, abs(float(b)))


def score(result, gold):
    report, totals = [], defaultdict(int)
    for pid, g in gold.items():
        res = result.get(pid, {})
        if not res:
            report.append(f"[{pid}] project not found in the Excel (check the project id)")
            continue
        gm = g["materials"]
        for pos, expected in gm.items():
            totals["expected"] += 1
            got = res.get(pos, [])
            if not got:
                totals["missing"] += 1
                report.append(f"MISSING        {pid}:{pos}  expected {expected}")
                continue
            totals["found"] += 1
            ok = True
            for name, qty, unit, poss in got:
                if len(poss) > 1:                                     # row sums several positions
                    exp_sum = sum(gm.get(p, [[0, ""]])[0][0] for p in poss)
                    distinct = len({json.dumps(gm.get(p)) for p in poss}) > 1 or len(poss) != len(set(poss))
                    if not close(qty, exp_sum):
                        ok = False
                        report.append(f"WRONG SUM      {pid}:{pos}  row '{name}' = {qty}, positions {poss} sum to {exp_sum}")
                    elif pos == poss[0]:
                        totals["merged_rows"] += 1
                        report.append(f"MERGED (check) {pid}:{', '.join(poss)}  in one row '{name}' = {qty} {unit}")
                    continue
                if not any(close(qty, e[0]) for e in expected):
                    ok = False
                    report.append(f"WRONG QTY      {pid}:{pos}  row '{name}' = {qty} {unit}, expected {expected}")
            if len(expected) > 1 and len(got) < len(expected):
                ok = False
                report.append(f"INCOMPLETE     {pid}:{pos}  {len(got)} of {len(expected)} products found {expected}")
            totals["correct"] += ok
        for pos in g.get("not_materials", []):
            if pos in res:
                totals["false_positive"] += 1
                report.append(f"NOT A MATERIAL {pid}:{pos}  listed as '{res[pos][0][0]}'")
        extra = [p for p in res if p not in gm and p not in g.get("not_materials", [])]
        for p in extra:
            report.append(f"UNKNOWN        {pid}:{p}  not in the gold standard ('{res[p][0][0]}')")
    return totals, report


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("excel", nargs="?", help="materials.xlsx from extract_to_excel.py")
    ap.add_argument("--gold", help="gold standard JSON (default: built-in HEM2301-2)")
    ap.add_argument("--export-gold", help="write the built-in gold standard to this file and exit")
    a = ap.parse_args()
    if a.export_gold:
        Path(a.export_gold).write_text(json.dumps(GOLD, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"gold standard written to {a.export_gold}")
        return
    if not a.excel:
        ap.error("give the Excel file to score")
    gold = json.loads(Path(a.gold).read_text(encoding="utf-8")) if a.gold else GOLD
    result = read_result(a.excel)
    junk = result.pop("__rows_without_position__", [])
    t, report = score(result, gold)
    report += [f"NO POSITION    {pid}: row '{name}' cannot be traced to a position" for pid, name in junk]
    n = max(t["expected"], 1)
    print(f"Materials found        : {t['found']}/{t['expected']}  ({100 * t['found'] / n:.1f} %)")
    print(f"Found AND correct      : {t['correct']}/{t['expected']}  ({100 * t['correct'] / n:.1f} %)")
    print(f"Missing                : {t['missing']}")
    print(f"Wrongly included       : {t['false_positive']}")
    print(f"Rows summing positions : {t['merged_rows']}  (check they are really the same product)")
    print(f"Rows without position  : {len(junk)}  (noise: duplicates / general text)")
    print("-" * 70)
    print("\n".join(report) if report else "No errors – everything matches the gold standard.")


if __name__ == "__main__":
    main()