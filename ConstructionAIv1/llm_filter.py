"""
llm_filter.py  –  FILE 1 of 2: data preparation + LLM filtering
=================================================================
Reads Azure Document Intelligence / Content Understanding result files (valid JSON), cuts the
text into one "candidate" per bill-of-quantities position (never cut in the middle) and lets an
LLM decide for EVERY position whether it is a material that must be bought.

The LLM does NOT produce numbers. For every material it returns material_name, description and
VERBATIM text snippets for quantity / unit price / total price (plus, for amounts written "per
filter / per unit", the position quantity). Numbers are extracted later, deterministically, by
extract_to_excel.py.

Output (materials_filtered.json):
    {"materials": [...], "not_materials": [...positions rejected, with the LLM's reason...],
     "not_reviewed": [...positions with a quantity that the LLM did not answer for...]}

Usage
    python llm_filter.py --check-llm                 # 1. test the connection (one tiny call)
    python llm_filter.py --dry-run                   # 2. test reading features.json, no LLM call
    python llm_filter.py                             # 3. real run on features.json
    python llm_filter.py results/*.json -o materials_filtered.json
    python llm_filter.py part1.json part2.json --project HEM2301-2 -o materials_filtered.json
Environment (Azure OpenAI, default) – as environment variables or in a .env file:
    AZURE_OPENAI_ENDPOINT=https://<resource>.openai.azure.com
    AZURE_OPENAI_API_KEY=<key>
    AZURE_OPENAI_DEPLOYMENT=<deployment name, e.g. gpt-4o>
    AZURE_OPENAI_API_VERSION=2024-10-21   (optional)
or  LLM_PROVIDER=anthropic with ANTHROPIC_API_KEY   (Claude, default model claude-sonnet-5-5)
or  LLM_PROVIDER=openai    with OPENAI_API_KEY      (default model gpt-4o)
Choose the model with LLM_MODEL=... in .env or per run:  --provider anthropic --model claude-opus-5-5
"""
import argparse
import hashlib
import json
import os
import re
import time
import urllib.error
import urllib.request
from collections import Counter
from pathlib import Path

# ============================================================== 1. loading
def _flatten_field(name, f, out):
    """Azure field object -> list of (path, text). Keeps arrays/objects readable."""
    if not isinstance(f, dict):
        return
    t = f.get("type")
    if t == "array":
        for i, x in enumerate(f.get("valueArray", [])):
            _flatten_field(f"{name}[{i}]", x, out)
    elif t == "object":
        for k, v in f.get("valueObject", {}).items():
            _flatten_field(f"{name}.{k}", v, out)
    else:
        for k in ("valueString", "valueNumber", "valueInteger", "valueDate", "valueTime",
                  "valueCurrency", "valueBoolean", "content"):
            if k in f:
                out.append((name, json.dumps(f[k], ensure_ascii=False) if not isinstance(f[k], str) else f[k]))
                break


def read_json_documents(path):
    """A file may contain ONE JSON document, a LIST of documents, or SEVERAL documents written
    one after another (concatenated / one per line). Returns a list of documents."""
    text = Path(path).read_text(encoding="utf-8-sig")          # -sig: ignore a Windows BOM
    dec, i, docs = json.JSONDecoder(), 0, []
    while True:
        while i < len(text) and text[i] in " \t\r\n,":
            i += 1
        if i >= len(text):
            break
        rest = text[i:]
        junk = re.sub(r"[\s\x00\x1a﻿​ }\]]", "", rest)
        if docs and not junk:
            # leftovers after the last document: invisible characters (NUL, Ctrl-Z, BOM,
            # zero-width / non-breaking space) or stray closing brackets -> ignore, but tell the user
            print(f"  warning: ignored {len(rest)} stray character(s) at the end of {Path(path).name}: "
                  f"{rest[:20]!r}")
            break
        try:
            obj, i = dec.raw_decode(text, i)
        except json.JSONDecodeError as e:
            has_content = any(isinstance(d, dict) and (d.get("result") or d).get("contents") for d in docs)
            if has_content:
                # e.g. one "}" too many earlier in the file: the document (result.contents) was read
                # completely, only a trailing part (often the "usage" block) is left -> ignore it
                line = text.count("\n", 0, i) + 1
                print(f"  warning: {Path(path).name}: the JSON closes early at line {line}; the remaining "
                      f"{len(text) - i:,} character(s) are not valid JSON and were ignored "
                      f"(they start with {text[i:i + 40]!r}). The document content was read completely.")
                break
            raise SystemExit(f"{Path(path).name}: invalid JSON after {len(docs)} document(s) – "
                             f"line {e.lineno}, column {e.colno}: {e.msg}")
        docs.extend(obj if isinstance(obj, list) else [obj])
    return docs


def load_results(path):
    """All Azure results in one file -> list of {file, markdown, fields, project_hint}."""
    docs = read_json_documents(path)
    out = []
    for n, doc in enumerate(docs, 1):
        r = load_result(doc)
        r["file"] = Path(path).name if len(docs) == 1 else f"{Path(path).name}#{n}"
        out.append(r)
    return out


def load_result(doc):
    """One Azure result:  result.contents[*].markdown  and  result.contents[*].fields."""
    if not isinstance(doc, dict):
        return {"file": "", "markdown": "", "fields": [], "project_hint": None}
    contents = (doc.get("result") or doc).get("contents", [])
    markdown = "\n".join(c.get("markdown", "") for c in contents)
    fields = []
    for c in contents:
        for name, f in (c.get("fields") or {}).items():
            _flatten_field(name, f, fields)
    get = dict(fields)
    project = get.get("TenderNumber") or get.get("ProjectNumber") or get.get("ProjectName")
    return {"file": "", "markdown": markdown, "fields": fields, "project_hint": project}


# ============================================================== 2. data preparation
QTY_LINE = re.compile(r"^[.\s]*\d[\d.,]*\s*[A-Za-zäöü³²€]{0,6}\.?(\s+(EP|GP)\.?)*\s*$")


def clean_markdown(md):
    """Generic cleaning, no document-specific rules:
    - drop images, HTML comments (page numbers / headers / footers), flatten tables
    - remove running headers/footers = lines that occur in the first/last lines of many pages.
      Lines that look like a quantity ("2 St", "175 St EP. GP") are never removed."""
    md = re.sub(r"!\[.*?\]\(.*?\)", "", md)
    pages = re.split(r"<!--\s*PageBreak\s*-->", md)

    def norm(page):
        page = re.sub(r"<!--.*?-->", "\n", page, flags=re.S)
        page = re.sub(r"<br\s*/?>", " ", page)
        page = re.sub(r"</t[dh]>", "\n", page)
        page = re.sub(r"<[^>]+>", "", page)
        return [l.strip() for l in page.split("\n") if l.strip()]
    pages = [norm(p) for p in pages]
    edge = Counter()
    for p in pages:
        edge.update(set(p[:30] + p[-8:]))
    limit = max(3, 0.3 * len(pages))
    repeated = {l for l, n in edge.items() if n >= limit and len(l) < 100 and not QTY_LINE.match(l)}
    out = []
    for p in pages:
        head, tail = p[:30], p[-8:]
        for i, l in enumerate(p):
            at_edge = i < 30 or i >= len(p) - 8
            if at_edge and l in repeated:
                continue
            out.append(l)
    return "\n".join(out).strip()


POS = re.compile(r"^(?:#+\s*)?(\d{1,2}(?:\.\d{1,4}){1,3})(?![\d.,])(?:\s|$)")
# the ordered quantity of a German bill-of-quantities position: "175 St  EP.... GP ...."
_Q = r"(?<![\w,.])\d{1,3}(?:\.\d{3})*(?:,\d+)?\s*[A-Za-zäöüÄÖÜ]{1,6}[23²³]?\.?"
QTY_IN_BLOCK = re.compile(_Q + r"[\s.]{0,15}(?:EP|GP)\b")      # "175 St  EP.... GP ...."
QTY_ALONE = re.compile(r"(?m)^\s*" + _Q + r"\s*$")              # fallback: "1 psch" alone on its line


def _find_quantity(body):
    """The position's ordered quantity: first amount followed by the EP/GP price columns; if the
    columns were lost (page break), the LAST amount standing alone on a line."""
    m = QTY_IN_BLOCK.search(body)
    if m:
        return m
    alone = list(QTY_ALONE.finditer(body))
    return alone[-1] if alone else None


def _position_regex(lines):
    """Detect the document's position numbering (e.g. 05.01.3): the most frequent deepest pattern,
    so headings like '03.01', grades like '1.4571' and dates are not taken as positions."""
    shapes = Counter()
    for l in lines:
        m = POS.match(l)
        if m and not re.fullmatch(r"1\.4\d{3}", m.group(1)):
            p = m.group(1).split(".")
            shapes[(len(p),) + tuple(len(x) for x in p[:-1])] += 1
    good = [s for s, n in shapes.items() if n >= 3]
    if not good:
        return None
    depth = max(s[0] for s in good)
    shape = max((s for s in good if s[0] == depth), key=lambda s: shapes[s])
    return re.compile(r"^(?:#+\s*)?(" + r"\.".join(rf"\d{{{w}}}" for w in shape[1:]) + r"\.\d{1,3})(?![\d.,])(?:\s|$)")


def split_blocks(text):
    """Returns a list of (ref, text, context, has_quantity).
    Documents WITH position numbers (bill of quantities):
      - one block per position, NEVER cut (fix: long positions were cut in two before)
      - the block ends at the position's quantity line; text after it (e.g. "Spezifikation Kugelhahn")
        describes the following positions and is passed to the next block as context only
      - text before the first position (general description of the works) is NOT sent
        (fix: general sentences became fake items)
    Documents WITHOUT position numbers: paragraphs of ~1500 characters."""
    lines = text.split("\n")
    rx = _position_regex(lines)
    if rx is None:
        out, cur, n = [], "", 0
        for para in re.split(r"\n\s*\n", text):
            if cur and len(cur) + len(para) > 1500:
                n += 1; out.append((f"para{n}", cur, "", bool(_find_quantity(cur)))); cur = ""
            cur += ("\n\n" if cur else "") + para
        if cur:
            out.append((f"para{n + 1}", cur, "", bool(_find_quantity(cur))))
        return out
    raw, cur, ref = [], None, None
    for l in lines:
        m = rx.match(l)
        if m:
            if cur is not None:
                raw.append((ref, cur))
            cur, ref = [l], m.group(1)
        elif cur is not None:
            cur.append(l)
    if cur is not None:
        raw.append((ref, cur))
    out, context = [], ""
    for ref, ls in raw:
        body = "\n".join(ls)
        q = _find_quantity(body)
        if q:
            cut = body.find("\n", q.end())
            cut = len(body) if cut < 0 else cut
            item, tail = body[:cut], body[cut:].strip()
        else:
            item, tail = body, ""
        out.append((ref, item, context, bool(q)))
        context = tail[-1500:]          # shared specification text for the next position
    return out


def build_candidates(project_id, docs):
    """Fix: Azure fields are used only for the project number; items come from the text, where
    description and quantity are together (fields gave each item a 2nd time WITHOUT quantity).
    Fields are used as items only when a document has no text at all."""
    cands = []
    for d in docs:
        if d["markdown"].strip():
            for ref, txt, ctx, has_q in split_blocks(clean_markdown(d["markdown"])):
                cands.append({"id": f"C{len(cands)}", "project_id": project_id, "file": d["file"],
                              "ref": ref, "text": txt, "context": ctx, "has_quantity": has_q})
        else:
            groups = {}
            for path, val in d["fields"]:
                top = re.match(r"[^.\[]+(?:\[\d+\])?", path).group(0)     # one candidate per field / array item
                groups.setdefault(top, []).append(f"{path}: {val}")
            for top, vals in groups.items():
                cands.append({"id": f"C{len(cands)}", "project_id": project_id, "file": d["file"],
                              "ref": f"field:{top}", "text": "\n".join(vals), "context": "",
                              "has_quantity": True})
    seen, uniq = set(), []                       # identical texts (overlapping parts) only once
    for c in cands:
        h = hashlib.sha1(re.sub(r"\s+", " ", c["text"]).encode()).hexdigest()
        if h not in seen:
            seen.add(h); uniq.append(c)
    return uniq


# ============================================================== 3. LLM
SYSTEM = """You review positions of construction tender documents (bills of quantities) for
PROCUREMENT: what must be bought. Each CANDIDATE is one position: its number, title, description
and its ordered quantity line. Give a decision for EVERY candidate.

GENERAL RULES
1. Decide ONLY from the POSITION text. The CONTEXT (shared text of neighbouring positions) helps you
   understand terms but is never a reason to accept or reject a position.
2. is_material = true when the position makes the contractor SUPPLY something NEW and physical
   (supply verbs such as "liefern", "Lieferung", "anzubieten", "herstellen", "einbauen", "neu", or
   their equivalents in other languages). This stays true when the same position also describes
   installation, assembly, connection, or the removal of old parts: the new part is still bought.
3. is_material = false when nothing new and physical is bought: services (cleaning, disinfection,
   commissioning, testing, training, planning, maintenance contracts, site setup, travel),
   documents of any kind (drawings, plans, diagrams, documentation – even if they are "delivered"
   or printed), labour or equipment hours, removal / disposal only, and re-installing EXISTING
   equipment.
4. Placeholder tokens like [PERSON_12] or [ORT_3] are anonymisation artefacts that replaced ordinary
   words, sometimes the whole title. Never decide because of them – read the description.

ITEMS OF A MATERIAL POSITION (normally exactly ONE)
5. One item per distinct product that has its OWN amount. If the position lists several different
   products, each with its own amount or layer, return one item for each.
6. Parts, accessories or sub-assemblies described as included in or belonging to the supplied
   product are NOT separate items.
7. material_name (in {lang}): the product actually supplied, using its standard technical name.
   For an add-on / surcharge to another item, name the add-on itself. Include every detail that
   distinguishes it from similar positions: type, size (DN, dimensions), pressure rating, angle,
   power, height, material grade, variant. Two different positions must never get the same name
   unless they are really the same product.
8. description: max 20 words, in {lang}.

NUMBERS – COPY, NEVER CALCULATE
9. quantity_text: copy the amount VERBATIM from the text. Normally the ordered quantity of the
   position. If the real amount to buy is stated PER UNIT of the position (per filter, per tank,
   per piece …), copy that per-unit amount into quantity_text, set quantity_is_per_unit = true and
   copy the position's ordered quantity into position_quantity_text. Otherwise
   quantity_is_per_unit = false and position_quantity_text = null.
10. unit_price_text / total_price_text: verbatim amounts that contain digits; null when the price
    columns are empty."""

ITEM = {"type": "object", "additionalProperties": False,
        "required": ["material_name", "description", "quantity_text", "quantity_is_per_unit",
                     "position_quantity_text", "unit_price_text", "total_price_text"],
        "properties": {"material_name": {"type": "string"}, "description": {"type": "string"},
                       "quantity_text": {"type": ["string", "null"]},
                       "quantity_is_per_unit": {"type": "boolean"},
                       "position_quantity_text": {"type": ["string", "null"]},
                       "unit_price_text": {"type": ["string", "null"]},
                       "total_price_text": {"type": ["string", "null"]}}}
SCHEMA = {"type": "object", "additionalProperties": False, "required": ["decisions"],
          "properties": {"decisions": {"type": "array", "items": {
              "type": "object", "additionalProperties": False,
              "required": ["candidate_id", "is_material", "reason", "items"],
              "properties": {"candidate_id": {"type": "string"}, "is_material": {"type": "boolean"},
                             "reason": {"type": "string"}, "items": {"type": "array", "items": ITEM}}}}}}


REQUIRED_ENV = {"azure_openai": ["AZURE_OPENAI_ENDPOINT", "AZURE_OPENAI_API_KEY", "AZURE_OPENAI_DEPLOYMENT"],
                "openai": ["OPENAI_API_KEY"],
                "anthropic": ["ANTHROPIC_API_KEY"]}
DEFAULT_MODEL = {"openai": "gpt-4o", "anthropic": "claude-sonnet-5-5"}


def check_config():
    provider = os.getenv("LLM_PROVIDER", "azure_openai")
    if provider not in REQUIRED_ENV:
        raise SystemExit(f"LLM_PROVIDER='{provider}' is not supported (use {', '.join(REQUIRED_ENV)})")
    missing = [v for v in REQUIRED_ENV[provider] if not os.getenv(v)]
    if missing:
        raise SystemExit(f"Missing environment variable(s) for {provider}: {', '.join(missing)}\n"
                         "Set them (or put them in a .env file next to this script) – see the header of this file.")
    return provider


def load_dotenv(path=".env"):
    """Optional: read KEY=VALUE lines from a .env file (does not override variables already set)."""
    for f in (Path(path), Path(__file__).with_name(".env")):
        if f.exists():
            for line in f.read_text(encoding="utf-8").splitlines():
                if "=" in line and not line.lstrip().startswith("#"):
                    k, v = line.split("=", 1)
                    os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))
            return


class LLM:
    """provider: azure_openai | openai | anthropic (env LLM_PROVIDER or --provider)
    model    : env LLM_MODEL or --model (Azure: the deployment name)"""

    def __init__(self, cache=".llm_cache", provider=None, model=None):
        self.provider = provider or os.getenv("LLM_PROVIDER", "azure_openai")
        if self.provider == "azure_openai":
            self.model = model or os.getenv("LLM_MODEL") or os.getenv("AZURE_OPENAI_DEPLOYMENT", "")
        else:
            self.model = model or os.getenv("LLM_MODEL") or DEFAULT_MODEL[self.provider]
        self.cache = Path(cache); self.cache.mkdir(exist_ok=True)
        self.param_fixes = {}       # e.g. {"max_tokens": "max_completion_tokens", "temperature": None}

    def _apply_fixes(self, body):
        for old, new in self.param_fixes.items():
            if old in body:
                val = body.pop(old)
                if new:
                    body[new] = val
        return body

    def _learn_fix(self, detail):
        """Newer models reject some parameters ('max_tokens', 'temperature' …). Learn the fix from the
        error message and retry. Returns True if a fix was learned."""
        try:
            err = json.loads(detail).get("error", {})
        except (ValueError, AttributeError):
            return False
        param, msg = err.get("param") or "", err.get("message", "")
        if err.get("code") not in ("unsupported_parameter", "unsupported_value") or not param:
            return False
        if param in self.param_fixes:
            return False
        m = re.search(r"[Uu]se '(\w+)' instead", msg)
        self.param_fixes[param] = m.group(1) if m else None     # rename, or drop the parameter
        print(f"  note: model does not support '{param}' -> "
              f"{'using ' + repr(m.group(1)) if m else 'leaving it out'} and retrying")
        return True

    def ask(self, system, user):
        # the model is part of the cache key: switching model never reuses another model's answers
        key = hashlib.sha256((self.provider + self.model + system + user).encode()).hexdigest()[:24]
        f = self.cache / f"{key}.json"
        if f.exists():                                 # re-runs are free and identical
            return json.loads(f.read_text(encoding="utf-8"))
        err = None
        for attempt in range(6):
            try:
                data = json.loads(self._call(system, user))
                f.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
                return data
            except urllib.error.HTTPError as e:
                detail = e.read().decode("utf-8", "replace")[:2000]
                err = RuntimeError(f"HTTP {e.code}: {detail[:600]}")
                if e.code == 400 and self.provider != "anthropic" and self._learn_fix(detail):
                    continue                                    # retry at once with the fixed request
                if e.code not in (408, 429, 500, 502, 503, 504, 529):   # 400/401/403/404 -> retrying won't help
                    break
                time.sleep(2 ** attempt * 5)
            except Exception as e:                      # network / bad JSON
                err = e; time.sleep(2 ** attempt * 2)
        raise RuntimeError(f"LLM call failed ({self.provider}, {self.model}): {err}")

    def _post(self, url, headers, body):
        req = urllib.request.Request(url, json.dumps(body).encode(), headers, method="POST")
        with urllib.request.urlopen(req, timeout=300) as r:
            return json.loads(r.read())

    def _call(self, system, user):
        if self.provider == "anthropic":
            # Claude: structured output via a forced tool whose input_schema is our JSON schema
            body = {"model": self.model, "max_tokens": 16000, "temperature": 0, "system": system,
                    "messages": [{"role": "user", "content": user}],
                    "tools": [{"name": "record_decisions", "description": "Record the decision for every candidate.",
                               "input_schema": SCHEMA}],
                    "tool_choice": {"type": "tool", "name": "record_decisions"}}
            res = self._post("https://api.anthropic.com/v1/messages",
                             {"x-api-key": os.environ["ANTHROPIC_API_KEY"], "anthropic-version": "2023-06-01",
                              "content-type": "application/json"}, body)
            if res.get("stop_reason") == "max_tokens":
                raise RuntimeError("answer truncated – lower --batch-chars")
            block = next(b for b in res["content"] if b["type"] == "tool_use")
            return json.dumps(block["input"])
        body = {"messages": [{"role": "system", "content": system}, {"role": "user", "content": user}],
                "temperature": 0, "max_tokens": 32000,
                "response_format": {"type": "json_schema",
                                    "json_schema": {"name": "decisions", "schema": SCHEMA, "strict": True}}}
        if self.provider == "azure_openai":
            url = (f"{os.environ['AZURE_OPENAI_ENDPOINT'].rstrip('/')}/openai/deployments/"
                   f"{self.model}/chat/completions"
                   f"?api-version={os.getenv('AZURE_OPENAI_API_VERSION', '2024-10-21')}")
            headers = {"api-key": os.environ["AZURE_OPENAI_API_KEY"], "Content-Type": "application/json"}
        else:
            url, body["model"] = "https://api.openai.com/v1/chat/completions", self.model
            headers = {"Authorization": f"Bearer {os.environ['OPENAI_API_KEY']}", "Content-Type": "application/json"}
        res = self._post(url, headers, self._apply_fixes(body))
        if res["choices"][0].get("finish_reason") == "length":
            raise RuntimeError("answer truncated – lower --batch-chars")
        return res["choices"][0]["message"]["content"]


SUPPLY_WORDS = re.compile(r"\bliefern\b|\bLieferung\b|anzubieten|\bneue[nrs]?\b|\bherstellen\b|\beinbauen\b",
                          re.I)
NO_PURCHASE = re.compile(r"versetzen|Wiedermontage|\bin Regie\b|Flie(ß|ss)bild|Dokumentation|Zeichnung", re.I)


def supply_wording(text):
    """Supply wording in a position text, or None. Used to double-check rejections."""
    m = SUPPLY_WORDS.search(text)
    return m.group(0) if m else None


def _prompt(batch):
    parts = []
    for c in batch:
        ctx = f"\n\nCONTEXT (shared text of neighbouring positions – NOT this position):\n{c['context']}" \
            if c["context"] else ""
        hint = f"\n\nNOTE: {c['hint']}" if c.get("hint") else ""
        parts.append(f"### CANDIDATE {c['id']}\nPOSITION:\n{c['text']}{ctx}{hint}")
    return "\n\n".join(parts)


def filter_materials(cands, llm, lang="English", batch_chars=12000, log=print):
    """Returns (materials, rejected, not_reviewed). Every candidate must get a decision; candidates
    the LLM skipped are asked again once, then reported (fix: items were silently missed)."""
    batches, cur, size = [], [], 0
    for c in cands:
        n = len(c["text"]) + len(c["context"])
        if cur and size + n > batch_chars:          # a long position simply gets its own batch
            batches.append(cur); cur, size = [], 0
        cur.append(c); size += n
    if cur:
        batches.append(cur)
    by_id, decided = {c["id"]: c for c in cands}, {}
    system = SYSTEM.replace("{lang}", lang)
    for i, b in enumerate(batches, 1):
        res = llm.ask(system, _prompt(b))["decisions"]
        missing = [c for c in b if c["id"] not in {d["candidate_id"] for d in res}]
        if missing:
            log(f"  batch {i}: LLM skipped {len(missing)} candidate(s) -> asking again")
            res += llm.ask(system, _prompt(missing))["decisions"]
        for d in res:
            if d["candidate_id"] in by_id and d["candidate_id"] not in decided:
                decided[d["candidate_id"]] = d
        log(f"  batch {i}/{len(batches)}: {len(b)} positions -> "
            f"{sum(1 for c in b if decided.get(c['id'], {}).get('is_material'))} materials")
    # second opinion: rejected positions whose text says something is supplied
    doubtful = []
    for c in cands:
        d = decided.get(c["id"])
        w = supply_wording(c["text"])
        if d and not d["is_material"] and w and not NO_PURCHASE.search(c["text"]):
            doubtful.append({**c, "hint": f'You rejected this position, but its text contains the supply wording '
                                          f'"{w}". Re-check: is a NEW product bought here (even if installation '
                                          f'or removal of old parts is also described)? Decide from POSITION only.'})
    if doubtful:
        log(f"  second opinion on {len(doubtful)} rejected position(s) with supply wording")
        res = llm.ask(system, _prompt(doubtful))["decisions"]
        for d in res:
            if d["candidate_id"] in by_id:
                d["second_opinion"] = True
                decided[d["candidate_id"]] = d
        changed = sum(1 for d in res if d.get("is_material"))
        log(f"  -> {changed} changed to material")

    materials, rejected, not_reviewed = [], [], []
    for c in cands:
        base = {"project_id": c["project_id"], "file": c["file"], "position": c["ref"]}
        d = decided.get(c["id"])
        if d is None:
            if c["has_quantity"]:
                not_reviewed.append({**base, "text": c["text"][:300]})
            continue
        if not d["is_material"] or not d["items"]:
            w = supply_wording(c["text"])
            check = (f'CHECK: text contains supply wording "{w}"'
                     if w and not NO_PURCHASE.search(c["text"]) else "")
            rejected.append({**base, "reason": d["reason"], "check": check,
                             "second_opinion": bool(d.get("second_opinion")), "text": c["text"][:200]})
            continue
        for it in d["items"]:
            materials.append({**base, **{k: (v.strip() if isinstance(v, str) else v) for k, v in it.items()},
                              "source_text": c["text"]})
    return materials, rejected, not_reviewed


# ============================================================== main
def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("inputs", nargs="*", default=["features.json"],
                    help="Azure result JSON files or folders (default: features.json)")
    ap.add_argument("-o", "--output", default="materials_filtered.json")
    ap.add_argument("--project", help="force one project id for ALL inputs (parts of one project)")
    ap.add_argument("--lang", default="English", help="language for material names/descriptions")
    ap.add_argument("--batch-chars", type=int, default=12000)
    ap.add_argument("--provider", choices=list(REQUIRED_ENV), help="overrides LLM_PROVIDER")
    ap.add_argument("--model", help="overrides LLM_MODEL (Azure: deployment name), e.g. claude-opus-5-5")
    ap.add_argument("--check-llm", action="store_true", help="only test the LLM connection and exit")
    ap.add_argument("--dry-run", action="store_true", help="prepare candidates without calling the LLM")
    a = ap.parse_args(argv)

    load_dotenv()
    if a.provider:
        os.environ["LLM_PROVIDER"] = a.provider
    if a.check_llm:
        provider = check_config()
        print(f"testing {provider} ...")
        llm = LLM(cache=".llm_cache_test", model=a.model)
        print(f"model: {llm.model}")
        res = llm.ask(
            SYSTEM.replace("{lang}", "English"),
            "### CANDIDATE C1\nPOSITION:\n06.03.5\nKugelhahn DN 50, PN 16 liefern und montieren\n10 St EP. GP"
            "\n\n### CANDIDATE C2\nPOSITION:\n01.02.1\nDesinfektion der Gesamtanlage\n1 psch GP")
        print("LLM OK, answer:", json.dumps(res, ensure_ascii=False, indent=1))
        print("Expected: C1 = material (ball valve, '10 St'), C2 = not material (service).")
        return

    files = []
    for i in a.inputs:
        p = Path(i)
        if not p.exists():
            raise SystemExit(f"Input not found: {p.resolve()}  (run the script from the folder that contains it, "
                             "or pass the full path)")
        files += sorted(x for x in p.iterdir() if x.suffix.lower() in (".json", ".txt")) if p.is_dir() else [p]
    # group files per project: --project > project number found in the file > file name
    projects = {}
    for f in files:
        docs = load_results(f)
        if len(docs) > 1:
            print(f"{f.name}: contains {len(docs)} JSON documents")
        for d in docs:
            # project: --project > project number found in the document > file name (+ #n per document)
            pid = a.project or d["project_hint"] or d["file"].replace(".json", "").replace(".txt", "")
            projects.setdefault(pid, []).append(d)
            print(f"  {d['file']}: project={pid}, {len(d['fields'])} field values, {len(d['markdown']):,} chars text")

    if a.dry_run:
        preview = []
        for pid, docs in projects.items():
            cands = build_candidates(pid, docs)
            print(f"project {pid}: {len(cands)} positions, {sum(c['has_quantity'] for c in cands)} with a "
                  f"quantity line (dry run, LLM not called)")
            preview += cands
        Path("candidates_preview.json").write_text(json.dumps(preview, ensure_ascii=False, indent=1), encoding="utf-8")
        print("candidates written to candidates_preview.json")
        return
    check_config()
    llm = LLM(model=a.model)
    print(f"LLM: {llm.provider} / {llm.model}")
    result = {"llm": f"{llm.provider}/{llm.model}", "materials": [], "not_materials": [], "not_reviewed": []}
    for pid, docs in projects.items():
        cands = build_candidates(pid, docs)
        print(f"project {pid}: {len(cands)} positions")
        m, r, n = filter_materials(cands, llm, a.lang, a.batch_chars)
        result["materials"] += m; result["not_materials"] += r; result["not_reviewed"] += n
    Path(a.output).write_text(json.dumps(result, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{len(result['materials'])} material entries, {len(result['not_materials'])} positions rejected, "
          f"{len(result['not_reviewed'])} not reviewed -> {a.output}")
    if result["not_reviewed"]:
        print("WARNING: some positions with a quantity got no answer from the LLM – see 'not_reviewed'.")


if __name__ == "__main__":
    main()