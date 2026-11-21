# Materials from construction tenders → Excel

This tool reads construction tender documents (German LV / Leistungsverzeichnis, offers, GAEB files …)
and produces **the list of materials to buy, per project**:

| Project | Material | Quantity | Unit | Unit price | Total price | Description |
|---|---|---|---|---|---|---|
| 79-222 | Rubber skirting, approx. 8 cm high | 30 | m | | | Rubber wall-base skirting … |

Missing values stay empty – nothing is invented. Materials are only added together **inside the same project**.

---

## 1. How it works (the big picture)

```
 tenders\  (all your files, mixed)
     │
     ▼
 STEP 1  read every file ─────────────► output\json\      (raw results)
         • PDF, Word, Excel, scans → Azure AI          output\features\  (readable features)
         • GAEB X83 / D84 → read locally (exact)
     │
     ▼
 STEP 1b find the project of every file ─► output\file_index.csv
         (LLM reads project number / name, Python checks it)
     │
     ▼
 STEP 2  LLM decides: material to buy or not? ─► output\materials_filtered.json
     │
     ▼
 STEP 3  Python extracts the numbers and builds the results
         ├──► output\materials\        (materials per document, JSON)
         ├──► output\all_projects.xlsx (all projects in one workbook)
         └──► output\excel\<project>.xlsx (one workbook per project)
```

**Who does what**
* **Azure AI** reads the documents (text and tables). GAEB files don't need Azure.
* **The LLM** (your Azure OpenAI model, or Claude) only *decides*: which file belongs to which project, and
  which position is a material.
* **Python** does everything that must be exact: quantities, units, prices, sums. Every number is checked
  against the document text.

---

## 2. Setup (once)

### 2.1 Install

Open the PyCharm terminal in the project folder:

```powershell
cd "F:\BC Project\materials_pipeline"
python -m venv .venv          # only if you don't have one yet
.venv\Scripts\activate
pip install -r requirements.txt
```

### 2.2 The `.env` file (your keys)

Copy `.env.example` to `.env` and fill it in. Everything comes from the **Microsoft Foundry** page of your
resource (Home page: *Clé API*, *Point de terminaison de projet*, *Point de terminaison Azure OpenAI*).

```
# Azure AI – reads the documents
AZURE_CU_ENDPOINT=https://<your-resource>.services.ai.azure.com      ← "Point de terminaison de projet",
AZURE_CU_KEY=<Clé API>                                                  only the part before /api/projects
AZURE_CU_ANALYZER=prebuilt-layout
AZURE_CU_API_VERSION=2025-11-01

# LLM – decides what is a material
LLM_PROVIDER=azure_openai
AZURE_OPENAI_ENDPOINT=https://<your-resource>.openai.azure.com       ← "Point de terminaison Azure OpenAI"
AZURE_OPENAI_API_KEY=<Clé API>                                          (same key as above – that is normal)
AZURE_OPENAI_DEPLOYMENT=<your model deployment name>

# Later, to use Claude instead:
# LLM_PROVIDER=anthropic
# ANTHROPIC_API_KEY=<your Claude key>
# LLM_MODEL=claude-sonnet-5-5
```

No storage account and no extra Azure model are needed: `prebuilt-layout` works without them.
Never share the key or put it in a screenshot.

### 2.3 Test the connections

```powershell
python step1_analyze.py --check-azure      # expected: AZURE OK
python llm_filter.py --check-llm           # expected: C1 = material, C2 = not material
```

---

## 3. Put your documents in place

Create a folder `tenders` inside the project folder and copy **all** documents into it.
No sorting needed – the tool finds the projects itself.

```
F:\BC Project\materials_pipeline\
   tenders\
      LV_06_Kautschukboden.pdf
      HEM2301-2_LV.X83
      Angebot_AN24-0305.pdf
      ...                      (zip files are unpacked automatically)
```

| File types | Read by |
|---|---|
| PDF, scans (jpg, png, tif), Word (.docx), Excel (.xlsx), PowerPoint (.pptx), txt, html, email | Azure AI |
| GAEB `.X81–.X86`, `.D81–.D86`, GAEB `.xml` | read locally (exact, free) |
| `.doc`, `.xls`, `.dwg`, GAEB 2000 `.P8x` | not supported – listed in `file_index.csv` with a hint (e.g. "save as .docx") |

*Optional:* if you already know the projects, sub-folders work too (`tenders\79-222\…`). The folder name
is then the project, and the LLM is not asked.

---

## 4. Run

### Everything at once
```powershell
python run_all.py tenders --out output
```

### Step by step (recommended the first time)

| # | Command | What it does | Cost |
|---|---|---|---|
| 1 | `python step1_analyze.py tenders --out output` | reads all files, writes `json\` and `features\` | Azure, per page – only for new files |
| 1b | `python group_projects.py --out output` | finds the project of each file | small LLM cost |
| – | open `output\file_index.csv` | check the `project_id` column, correct by hand if needed | – |
| 2+3 | `python run_all.py tenders --out output --skip-analysis` | materials (LLM) + JSON + Excel | LLM |
| 3 only | `python extract_to_excel.py output\materials_filtered.json --out-dir output` | rebuilds JSON + Excel from the saved LLM result | free |

Useful options
* `--no-excel` – only the JSON files, no Excel
* `--rebuild-features` (step 1) – compute the features files again after a code update (free)
* `--force` (step 1) – send files to Azure again (costs money)
* `--model <name>` – use another LLM model

**Re-running is cheap:** a file is never sent to Azure twice, existing features files are kept, and LLM
answers are saved in `output\.llm_cache\`.

---

## 5. Where to find everything

```
output\
  json\azure\              raw Azure answers (big files – you don't need to open them)
  json\local\              raw GAEB results
  features\azure\          ★ readable features of each document read by Azure AI
  features\local\          ★ readable features of each GAEB / XML file
  file_index.csv           ★ every input file: project, how it was assigned, status, problems
  materials_filtered.json  LLM decisions for all projects (materials + rejected positions)
  materials\azure\         ★ materials to buy, per document (Azure documents)
  materials\local\         ★ materials to buy, per document (GAEB files)
  all_projects.xlsx        ★ THE RESULT – all projects in one workbook
  excel\<project>.xlsx     ★ one workbook per project
  .llm_cache\              saved LLM answers (ignore)
```

Open them from the terminal:

```powershell
start "output\all_projects.xlsx"
start "output\excel\79-222.xlsx"
start output\features
start output\materials
start "output\file_index.csv"
```

(PyCharm cannot show `.xlsx` files – open them with `start` or right-click → Open In → Associated Application.)

---

## 6. What the files look like

### 6.1 Features file (`features\…\*.features.json`)

Everything the document contains, each value with the same five keys:

| Key | Meaning |
|---|---|
| `name` | name of the feature, e.g. `OfferNumber`, `Position 1.1.20` |
| `type` | `string`, `number`, `date`, `object`, `table`, `array` |
| `description` | what it means |
| `value` | the value (for an object: its sub-features) |
| `evidence` | the exact text in the document it was read from |

```json
{"name": "OfferNumber", "type": "string", "description": "Offer number (Angebot-Nr.)",
 "value": "AN24-0305", "evidence": "Angebot-Nr.: | AN24-0305"}

{"name": "Position 1.1.20", "type": "object", "description": "Bill of quantities position",
 "value": [
   {"name": "PositionNumber", "value": "1.1.20"},
   {"name": "ShortText",      "value": "Baustelleneinrichtung vorhalten"},
   {"name": "Quantity",       "value": 6.25,   "evidence": "6,250 Mt  EP"},
   {"name": "Unit",           "value": "Mt"},
   {"name": "UnitPrice",      "value": 750.0},
   {"name": "TotalPrice",     "value": 4687.5},
   {"name": "PriceCheck",     "value": "ok"},
   {"name": "PositionKind",   "value": null}
 ]}
```

Sections of the file:
* `features` – document data: DocumentType, ProjectNumber, ProjectName, OfferNumber, TenderNumber, LVNumber,
  CustomerNumber, Date, Client, NetTotal, VATRate, GrossTotal
* `positions` – every position of the bill of quantities
* `tables` – other tables (cover sheet, summaries)
* `summary` – quick quality check:

```json
"summary": {"positions": 274, "with_quantity": 274, "with_unit_price": 274,
            "price_check_failed": 0, "total_check": "ok – positions add up to the document total"}
```
`with_quantity` lower than `positions` → some quantities were not found. `total_check` not "ok" → the
positions do not add up to the document total.

**Important:** a value always belongs to the object around it. `"Quantity": 1.0` alone says nothing – the
`Position …` object around it tells you which element it is.

### 6.2 Materials file (`materials\…\*.materials.json`)

Same format, only the materials to buy of one document:

| Feature | Meaning |
|---|---|
| Project, PositionNumber | where it comes from |
| MaterialName, Description | written by the LLM |
| Quantity, Unit, UnitPrice, TotalPrice | extracted by Python, with `evidence` |
| PerUnitAmount | amount per unit multiplied by the position quantity (e.g. "per filter × 2 filters") |
| PositionKind | **contingency** (only bought if needed) or **alternative** (replaces another position) – empty = normal |
| Conflict | another document of the project gives a different quantity |
| Notes | checks done by Python |

Plus `not_materials` (rejected positions with the LLM's reason) and `summary`.
Rows are **not yet summed** here – one entry per position.

### 6.3 The Excel workbook

| Sheet | Content |
|---|---|
| **Materials** | the list to buy: Project, Material, Quantity, Unit, Unit price, Total price, Description (same material in the same project = one row, quantities summed) |
| Sources | every row → file and position number, to check it in the PDF |
| **Checks** | things to look at: conflicts, contingency / alternative positions, possibly missing products, rows that sum several positions |
| Not materials | rejected positions with reason (yellow = contains supply wording, please glance at it) |
| Projects | one line per project: number of materials, documents used, duplicates, conflicts |
| Files | every input file and its status (yellow = not analysed, with the reason) |

Total price: taken from the document if it is there; otherwise calculated as quantity × unit price (Excel
formula); otherwise empty.

---

## 7. The rules the tool follows

* **Never mix projects.** A file that cannot be assigned safely becomes its own project `single:<file>`.
  Files are grouped by project number (checked in the text) or by a proven same project name.
* **Same LV twice** (e.g. as PDF and as X83): detected automatically, each position counts once, the GAEB
  version is kept.
* **Same position, different quantity** in two documents: both kept, never added, reported as **Conflict**.
* **Contingency / alternative positions** (Bedarfs-, Eventual-, Alternativ-, Wahlposition): kept but marked,
  and listed in *Checks* – they are not a normal purchase.
* **Numbers are never guessed.** Every quantity and price must appear in the document text; quantity ×
  unit price is compared with the total price.
* **Tables** are read row by row: prices glued to the text, quantity rows below the text and positions that
  continue on the next page are handled.
* General rules only – no rules written for one specific document.

---

## 8. Checking the quality

* In each features file: look at `summary` (see 6.1).
* Overview of all files at once:
  ```powershell
  python -c "import json,glob; [print(f.split('\\')[-1][:45].ljust(46), json.load(open(f,encoding='utf-8'))['summary']) for f in glob.glob('output/features/*/*.json')]"
  ```
* Compare a project with a verified result (available for HEM2301-2):
  ```powershell
  python evaluate.py output\excel\HEM2301-2.xlsx
  ```
  Results so far: HEM2301-2 → 127/127 materials found; offer AN24-0305 → 274 positions, sum = offer total
  to the cent.

---

## 9. Problems and solutions

| Message | What to do |
|---|---|
| `AZURE TEST FAILED: HTTP 401` | wrong key in `.env` |
| `HTTP 404` (Azure) | wrong endpoint – use the *Point de terminaison de projet* without `/api/projects/…` |
| `HTTP 404` (LLM) | wrong deployment name (`AZURE_OPENAI_DEPLOYMENT`) |
| `HTTP 400 … max_tokens` | nothing to do – the tool adapts automatically and retries |
| `Missing setting(s)` | a line is missing in `.env`, or `.env` is not in the project folder |
| `PermissionError` on `.xlsx` | close the workbook in Excel and run again |
| `cannot import name … from features` | one file is older than the others – replace all `.py` files with the latest version |
| `the following arguments are required: input` | old version of `step1_analyze.py`, or the folder name is missing (`tenders`) |
| PyCharm: `No loader for … .xlsx` | open it with `start "output\all_projects.xlsx"` |
| `start: file not found` | the files are in `output\` – write `start "output\all_projects.xlsx"` |
| `file says project 'X' but is in folder 'Y'` | check the folder or correct `project_id` in `file_index.csv` |

---

## 10. The files of the tool

| File | Role |
|---|---|
| `run_all.py` | runs all steps with one command |
| `step1_analyze.py` | step 1: sends files to Azure / reads GAEB, writes `json\` and `features\`, `--check-azure` |
| `gaeb_reader.py` | reads GAEB files locally |
| `features.py` | turns every result into the readable features format; reads tables |
| `group_projects.py` | step 1b: finds the project of every file |
| `llm_filter.py` | step 2: the LLM decides what is a material, `--check-llm` |
| `extract_to_excel.py` | step 3: numbers, materials JSON and Excel |
| `evaluate.py` | compares a result with a verified one |
| `requirements.txt`, `.env.example` | setup |
