""""
run_all.py – the whole pipeline with one command
================================================
    python run_all.py INPUT_FOLDER --out output

Step 1  step1_analyze.py  : every file -> Azure Content Understanding (or GAEB reader) -> output/json/*.json
                            + output/file_index.csv
Step 1b group_projects.py : files NOT in a project folder: the LLM finds the project of each file
                            (project number / name / site, checked against the file) -> file_index.csv
Step 2  llm_filter.py     : per project, the LLM keeps only the materials -> output/materials_filtered.json
Step 3  extract_to_excel.py: numbers extracted in Python -> output/excel/<project>.xlsx (one per project)
                            + output/all_projects.xlsx (all projects, rows NEVER merged across projects)

Options
    --manifest projects.csv   assign files to projects explicitly (columns: file, project_id)
    --review                  stop after step 1 so you can check/edit output/file_index.csv first;
                              then continue with:  python run_all.py INPUT --out output --skip-analysis
    --skip-analysis           reuse output/file_index.csv (no Azure calls)
    --provider / --model      LLM choice for step 2 (otherwise from .env)
    --projects A B            only these projects in steps 2-3
Re-running is cheap: analysed files and LLM answers are cached.
"""
import argparse
import sys
from pathlib import Path

import extract_to_excel
import group_projects
import llm_filter
import step1_analyze


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input", help="folder with the tender files (ideally one sub-folder per project)")
    ap.add_argument("--out", default="output")
    ap.add_argument("--manifest")
    ap.add_argument("--workers", type=int, default=4)
    ap.add_argument("--review", action="store_true", help="stop after step 1")
    ap.add_argument("--no-grouping", action="store_true",
                    help="do not let the LLM group loose files (each loose file = own project)")
    ap.add_argument("--skip-analysis", action="store_true", help="reuse the existing file_index.csv")
    ap.add_argument("--provider")
    ap.add_argument("--model")
    ap.add_argument("--projects", nargs="*")
    ap.add_argument("--lang", default="English")
    ap.add_argument("--no-excel", action="store_true", help="only the materials JSON files, no Excel")
    a = ap.parse_args()
    out = Path(a.out)
    index = out / "file_index.csv"

    if not a.skip_analysis:
        print("=" * 70 + "\nSTEP 1 – analyse files\n" + "=" * 70)
        args = [a.input, "--out", str(out), "--workers", str(a.workers)]
        if a.manifest:
            args += ["--manifest", a.manifest]
        step1_analyze.main(args)
        if not a.no_grouping:
            print("=" * 70 + "\nSTEP 1b – find the project of each file (files not in a project folder)\n" + "=" * 70)
            llm = None
            if a.provider or a.model:
                llm_filter.check_config()
                llm = llm_filter.LLM(cache=str(out / ".llm_cache"), provider=a.provider, model=a.model)
            group_projects.run(out, llm)
        if a.review:
            print(f"\nStopped for review. Check the project_id column in {index}, then run:\n"
                  f"  python run_all.py {a.input} --out {a.out} --skip-analysis")
            return
    elif not index.exists():
        sys.exit(f"{index} not found – run without --skip-analysis first")

    print("=" * 70 + "\nSTEP 2 – LLM: keep only the materials (per project)\n" + "=" * 70)
    args = ["--index", str(index), "-o", str(out / "materials_filtered.json"), "--cache",
            str(out / ".llm_cache"), "--lang", a.lang]
    if a.provider:
        args += ["--provider", a.provider]
    if a.model:
        args += ["--model", a.model]
    if a.projects:
        args += ["--projects", *a.projects]
    llm_filter.main(args)

    print("=" * 70 + "\nSTEP 3 – numbers + Excel (one workbook per project + all_projects.xlsx)\n" + "=" * 70)
    extract_to_excel.main([str(out / "materials_filtered.json"), "--out-dir", str(out)]
                          + (["--no-excel"] if a.no_excel else []))
    print(f"\nDone. Open {out / 'all_projects.xlsx'} (sheet 'Projects' first, then 'Checks').")


if __name__ == "__main__":
    main()