"""
features.py – readable features for every document
====================================================
Turns each analysed document (Azure Content Understanding JSON, or the local GAEB reader) into a
features file where every feature has:

    name         e.g. "OfferNumber", "Position 1.1.20"
    type         string | number | date | object | array | table
    description  what the feature means
    value        the value (objects: the sub-features, same format)
    evidence     the exact text in the document it was read from (so you can check it)

Document features : DocumentType, ProjectNumber, ProjectName, OfferNumber, TenderNumber, LVNumber,
                    CustomerNumber, Date, Client, NetTotal, VATRate, GrossTotal (+ any field Azure returned)
Positions         : one object per bill-of-quantities position: PositionNumber, Section, ShortText,
                    LongText, Quantity, Unit, UnitPrice, TotalPrice, PriceCheck
Other tables      : type "table", rows of cells

The table reader (render_tables) is also used by llm_filter.py, so the material filter sees table rows
like  "1.1.20 | 6,250 Mt Baustelleneinrichtung vorhalten | 750,00 | 4.687,50"  as a normal position
with quantity, unit price and total price.

    python features.py output\\json\\azure\\xxx.json        -> prints the features of one file
"""
import json
import re
import sys
from html import unescape
from pathlib import Path

POS_NO = re.compile(r"^(\d{1,2}(?:\.\d{1,4}){1,3})\.?$")
POS_START = re.compile(r"^(\d{1,2}(?:\.\d{1,4}){1,3})\.?\s+(.+)$")
MONEY = re.compile(r"^-?\d{1,3}(?:\.\d{3})*,\d{2}$|^-?\d+,\d{2}$")
QTY_UNIT = re.compile(r"^(\d{1,3}(?:\.\d{3})*(?:,\d+)?|\d+(?:,\d+)?)\s*([A-Za-zäöüÄÖÜ]{1,6}[23²³]?\.?)(?:\s+(.*))?$",
                      re.S)
QTY_CELL = re.compile(r"^\d{1,3}(?:\.\d{3})*(?:,\d+)?\s*[A-Za-zäöüÄÖÜ]{1,6}[23²³]?\.?$")
TOTAL_ROW = re.compile(r"^(Summe|Übertrag|Uebertrag|Zusammenstellung|Gesamt|Zwischensumme|Netto|Brutto|"
                       r"zzgl|zuzüglich|MwSt|Mehrwertsteuer)", re.I)
SECTION_ROW = re.compile(r"(Bereich|Abschnitt|Titel|Los)\s*\d", re.I)


# ------------------------------------------------------------------ special kinds of positions
KINDS = [("contingency – only ordered if needed (Bedarfs-/Eventualposition)",
          r"Bedarfsposition|Eventualposition|Bedarfspos\.|\bEventualpos|contingency item"),
         ("alternative – replaces a basic position, do not add both (Alternativ-/Wahlposition)",
          r"Alternativposition|Wahlposition|Alternativpos\.|Alternative zu (?:Pos|OZ)"),
         ("quantity to be given by the bidder", r"Menge vom Bieter|vom Bieter einzutragen")]


def position_kind(text):
    """'contingency …' / 'alternative …' / … if the position text says so, else None (= normal position)."""
    found = [k for k, pat in KINDS if re.search(pat, text or "", re.I)]
    return "; ".join(found) or None


# ------------------------------------------------------------------ numbers
def de_number(text):
    """'4.687,50' -> 4687.5 ; '6,250' -> 6.25 ; None if not a number."""
    if text is None:
        return None
    t = str(text).strip().replace(" ", "")
    if not re.fullmatch(r"-?[\d.]*\d(,\d+)?", t):
        return None
    try:
        return float(t.replace(".", "").replace(",", "."))
    except ValueError:
        return None


# ------------------------------------------------------------------ HTML tables of the Azure markdown
def _cells(row_html):
    out = []
    for m in re.finditer(r"<t([hd])([^>]*)>(.*?)</t[hd]>", row_html, re.S):
        txt = re.sub(r"<br\s*/?>", " ", m.group(3))
        txt = unescape(re.sub(r"<[^>]+>", "", txt))
        out.append(re.sub(r"\s+", " ", txt).strip())
    return out


def parse_table(table_html):
    return [_cells(r) for r in re.findall(r"<tr[^>]*>(.*?)</tr>", table_html, re.S)]


def _header_roles(cells):
    """Column roles from a header row (Pos-Nr. | Menge ME | EP / EUR | GP / EUR)."""
    roles = {}
    for i, c in enumerate(cells):
        lc = c.lower()
        if re.search(r"\bpos|\boz\b|ordnungszahl", lc):
            roles.setdefault("pos", i)
        if re.search(r"menge", lc):
            roles.setdefault("qty", i)
        if re.search(r"\bep\b|einheitspreis|e-preis", lc):
            roles.setdefault("ep", i)
        if re.search(r"\bgp\b|gesamtpreis|gesamtbetrag|g-preis", lc):
            roles.setdefault("gp", i)
    return roles if ("ep" in roles or "gp" in roles or "qty" in roles) else {}


TAIL_MONEY = re.compile(r"^(.*\S)\s+(-?\d{1,3}(?:\.\d{3})*,\d{2})$", re.S)


def row_position(cells, roles):
    """A table row that is a bill-of-quantities position -> dict, else None.
    If a price is missing because Azure glued it to the text ('… transportieren und 0,10'), the number
    at the end of the text is used as price ONLY when quantity x unit price = total price confirms it."""
    p = _row_position(cells, roles)
    if p and not (p["ep"] and p["gp"]):
        split = []
        for c in cells:
            m = TAIL_MONEY.match(c) if c and not MONEY.match(c) and not POS_NO.match(c) else None
            split += [m.group(1), m.group(2)] if m else [c]
        if split != cells:
            p2 = _row_position(split, {})
            q, ep, gp = de_number(p2 and p2["qty"]), de_number(p2 and p2["ep"]), de_number(p2 and p2["gp"])
            if q is not None and ep is not None and gp is not None and abs(q * ep - gp) <= max(0.02, 0.005 * gp):
                return p2
    return p


def _row_position(cells, roles):
    if not cells:
        return None
    first = next((c for c in cells if c), "")
    m = POS_NO.match(first)
    rest_start = cells.index(first) + 1 if first else 1
    rest = cells[rest_start:]
    if not m:
        m2 = POS_START.match(first)                    # "1.1.10 1,000 psch Baustelle …" in one cell
        if not m2 or MONEY.match(first):
            return None
        m, rest = m2, [m2.group(2)] + rest
    pos = m.group(1)
    money_idx = [i for i, c in enumerate(rest) if MONEY.match(c)]
    texts = [c for i, c in enumerate(rest) if c and i not in money_idx]
    text = " ".join(texts)
    qty = unit = None
    q = QTY_UNIT.match(text)
    if q:
        qty, unit, text = q.group(1), q.group(2).rstrip("."), (q.group(3) or "").strip()
    ep = gp = None
    vals = [rest[i] for i in money_idx]
    if len(vals) >= 2:
        ep, gp = vals[-2], vals[-1]
    elif len(vals) == 1:
        idx = rest_start + money_idx[0]
        if roles.get("gp") is not None and idx == roles["gp"] and roles.get("ep") != idx:
            gp = vals[0]
        elif roles.get("ep") is not None and idx == roles["ep"]:
            ep = vals[0]
        else:                                          # unknown column: treat it as the unit price
            ep = vals[0]
    return {"pos": pos, "text": text, "qty": qty, "unit": unit, "ep": ep, "gp": gp}


def render_tables(md):
    """Replace every HTML table by plain lines. Position rows become a normal position block:
         1.1.20
         Baustelleneinrichtung vorhalten
         6,250 Mt  EP 750,00  GP 4.687,50
    Rows without a number that continue the text (also on the next page, next to 'Übertrag') are added to
    the previous position; a quantity row below the text ('| | 10 h | EP. | GP |') becomes its quantity line."""
    tables = list(re.finditer(r"<table[^>]*>.*?</table>", md, flags=re.S))
    outs, roles_prev, last = [], {}, None                 # last = (table index, line index) of open position

    def add_text(text):
        ti, li = last
        lines = outs[ti][li].split("\n")
        if re.search(r"\bEP\b.*\bGP\b", lines[-1]):
            lines.insert(len(lines) - 1, text)
        else:
            lines.append(text)
        outs[ti][li] = "\n".join(lines)

    between_pos = re.compile(r"(?m)^\s*(?:#+\s*)?\d{1,2}(?:\.\d{1,4}){1,3}\.?(?:\s|$)")
    for ti, m in enumerate(tables):
        out = []
        outs.append(out)
        gap = md[tables[ti - 1].end():m.start()] if ti else md[:m.start()]
        if between_pos.search(re.sub(r"<!--.*?-->", "", gap, flags=re.S)):
            last = None              # positions/headings between the tables: the open position has ended
        for cells in parse_table(m.group(0)):
            if not any(cells):
                continue
            roles = _header_roles(cells)
            if roles and not any(POS_NO.match(c) or QTY_CELL.match(c) or MONEY.match(c) for c in cells) \
                    and ("pos" in roles or "qty" in roles) and all(len(c) <= 40 for c in cells):
                roles_prev = roles
                continue                                 # header row (repeated on every page)
            p = row_position(cells, roles_prev)
            if p:
                block = [p["pos"]]
                if p["text"]:
                    block.append(p["text"])
                qline = f'{p["qty"]} {p["unit"]}' if p["qty"] else ""
                if p["qty"] or p["ep"] or p["gp"]:
                    block.append(f'{qline}  EP {p["ep"] or "........"}  GP {p["gp"] or "........"}'.strip())
                out.append("\n" + "\n".join(block))
                last = (ti, len(out) - 1)
                continue
            joined = " | ".join(c for c in cells if c)
            carry = [c for c in cells if re.match(r"^(Übertrag|Uebertrag)", c, re.I)]
            if carry:                                    # page carry-over: keep only text next to it
                text = " ".join(c for c in cells if c and c not in carry and not MONEY.match(c))
                if last is not None and text:
                    add_text(text)
                continue
            qcell = next((c for c in cells if QTY_CELL.match(c)), None)
            only_qty = qcell and all(not c or c == qcell or MONEY.match(c) or re.fullmatch(r"(?:(?:EP|GP)\.?[\s./]*)+", c)
                                     for c in cells)        # "| | 10 h | EP. | GP |" – not "Förderhöhe: | 6,30 m"
            if last is not None and only_qty and not re.search(r"\bEP\b.*\bGP\b", outs[last[0]][last[1]].split("\n")[-1]):
                money = [c for c in cells if MONEY.match(c)]
                ep = money[-2] if len(money) >= 2 else (money[0] if len(money) == 1 else None)
                gp = money[-1] if len(money) >= 2 else None
                ti0, li0 = last
                outs[ti0][li0] += f'\n{qcell.rstrip(".")}  EP {ep or "........"}  GP {gp or "........"}'
                continue
            if TOTAL_ROW.match(joined) or SECTION_ROW.search(cells[0] or ""):
                out.append(joined)
                last = None                              # a sum or a new section closes the position
                continue
            if last is not None and not any(MONEY.match(c) for c in cells):
                add_text(joined)
                continue
            out.append(joined)
    parts, pos = [], 0
    for m, out in zip(tables, outs):
        parts.append(md[pos:m.start()])
        parts.append("\n" + "\n".join(out) + "\n")
        pos = m.end()
    parts.append(md[pos:])
    return "".join(parts)


# ------------------------------------------------------------------ document features
F = lambda name, type_, desc, value, evidence=None: {"name": name, "type": type_, "description": desc,
                                                     "value": value, "evidence": evidence}

DOC_KEYS = [  # name, type, description, label regex
    ("ProjectNumber", "string", "Project number (Projekt-Nr.)", r"Projekt[- ]?(?:Nr\.?|nummer)|^Projekt"),
    ("TenderNumber", "string", "Tender / award / measure number (Vergabe-Nr., Maßnahmennummer)",
     r"Vergabe[- ]?(?:Nr\.?|nummer)|Ma(?:ß|ss)nahmen[- ]?(?:Nr\.?|nummer)|Ausschreibungs[- ]?(?:Nr\.?|nummer)"),
    ("OfferNumber", "string", "Offer number (Angebot-Nr.)", r"Angebots?[- ]?(?:Nr\.?|nummer)"),
    ("LVNumber", "string", "Bill of quantities number (LV-Nr.)", r"^LV[- ]?(?:Nr\.?|nummer)?$|^LV[- ]Nr"),
    ("CustomerNumber", "string", "Customer number (Kunden-Nr.)", r"Kunden[- ]?(?:Nr\.?|nummer)"),
    ("ProjectName", "string", "Name of the construction project (Bauvorhaben / Baumaßnahme)",
     r"Bauvorhaben|Baumaßnahme|Baumassnahme|Bezeichnung der Bauleistung"),
    ("Client", "string", "Client (Auftraggeber / Bauherr / Vergabestelle)", r"Auftraggeber|Bauherr|Vergabestelle"),
    ("Date", "date", "Document date (Datum)", r"^Datum|Datum der Versendung"),
]
DOC_TYPES = [("Offer", r"\bAngebot\b(?!\s*(?:zur|s?frist|abgabe))"), ("Request for quotation", r"\bAnfrage\b"),
             ("Bill of quantities", r"Leistungsverzeichnis|\bLV\b"),
             ("Invitation to tender", r"Aufforderung zur Abgabe"), ("Construction description", r"Baubeschreibung"),
             ("Preliminary remarks", r"Vorbemerkung"), ("Contract", r"\bVertrag\b"), ("Drawing / plan", r"\bPlan\b")]


def _kv_pairs(md):
    """(label, value, evidence) from 2-column table rows, 'Label: value' lines and 'Label:' + next line."""
    pairs = []
    for t in re.findall(r"<table[^>]*>.*?</table>", md, re.S):
        rows = parse_table(t)
        for r in rows:
            r = [c for c in r if c]
            if len(r) == 2 and len(r[0]) < 60:
                pairs.append((r[0].rstrip(":").strip(), r[1], f"{r[0]} | {r[1]}"))
        if len(rows) >= 2 and len(rows[0]) == len(rows[1]) <= 4:      # header row + value row
            for h, v in zip(rows[0], rows[1]):
                if h and v and len(h) < 40:
                    pairs.append((h.rstrip(":").strip(), v, f"{h} | {v}"))
    text = re.sub(r"<table[^>]*>.*?</table>|<!--.*?-->", "\n", md, flags=re.S)
    lines = [l.strip().lstrip("#").strip() for l in text.splitlines()]
    lines = [l for l in lines if l]
    for i, l in enumerate(lines):
        m = re.match(r"^([A-Za-zÄÖÜäöüß][\w .\-/äöüß]{1,40}?):\s*(.*)$", l)
        if not m:
            continue
        label, val = m.group(1).strip(), m.group(2).strip()
        if not val and i + 1 < len(lines) and len(lines[i + 1]) < 120:
            val = lines[i + 1]
            l = f"{l} {val}"
        if val:
            pairs.append((label, val, l))
    return pairs


def document_features(md, fields_from_azure=None):
    feats, pairs = [], _kv_pairs(md)
    head = md[:3000]
    hits = [(m.start(), name, m) for name, pat in DOC_TYPES for m in [re.search(pat, head)] if m]
    if hits:                                          # the title that comes FIRST in the document decides
        pos, name, m = min(hits, key=lambda h: h[0])
        line = head[head.rfind("\n", 0, pos) + 1: head.find("\n", pos) if head.find("\n", pos) > 0 else None]
        line = re.sub(r"<!--\s*\w+:\s*|-->|^#+", "", line).strip()
        feats.append(F("DocumentType", "string", "Kind of document (from its title)", name, line[:150]))
    for name, type_, desc, label in DOC_KEYS:
        for lab, val, ev in pairs:
            m_in = re.search("(?:" + label + r")\.?\s*:?\s*(\S*\d[\w./\-]*)\s*$", lab, re.I)
            if name.endswith("Number") and m_in:          # "Vergabenummer HEM2301-2 | …": number is in the label
                val = m_in.group(1)
            if re.search(label, lab, re.I) and val and not re.fullmatch(r"[\s:.\-]*", val):
                v = re.split(r"[A-Za-zÄÖÜäöüß]*[-\s]?(?:Nr\.?|nummer|Datum|Seite)\s*:", val.strip())[0].strip(" |,;")
                if not v or (name.endswith("Number") and (not re.search(r"\d", v) or len(v) > 30 or v.count(" ") > 2)):
                    continue                          # a number must contain a digit ("Titel" is a header)
                if type_ == "date":
                    d = re.search(r"\d{1,2}\.\d{1,2}\.\d{2,4}", v)
                    if not d:
                        continue
                    v = d.group(0)
                feats.append(F(name, type_, desc, v, ev[:200]))
                break
    # totals (offer / LV sum, VAT, gross)
    flat = re.sub(r"<[^>]+>", " ", render_tables(md))
    money = r"(\d{1,3}(?:\.\d{3})*,\d{2})"
    for name, desc, pat in [("NetTotal", "Net total of the document (Summe netto / Angebotssumme / Summe LV)",
                             r"(?:Summe\s*\(?netto\)?|Angebots?summe(?:\s*netto)?|Summe\s+LV|Nettosumme|Gesamt\s*netto)"
                             r"[^\n\d]{0,40}" + money),
                            ("GrossTotal", "Gross total incl. VAT (Gesamtsumme / brutto)",
                             r"(?:Gesamtsumme|Bruttosumme|Summe\s*\(?brutto\)?|Gesamtbetrag\s+inkl[^\n\d]*)[^\n\d]{0,40}" + money)]:
        found = list(re.finditer(pat, flat, re.I))
        if found:
            m = found[-1]
            feats.append(F(name, "number", desc, de_number(m.group(1)), re.sub(r"\s+", " ", m.group(0))[:150]))
    m = re.search(r"(\d{1,2}(?:,\d{1,2})?)\s*%\s*(?:MwSt|Mehrwertsteuer|USt)|(?:MwSt|Mehrwertsteuer|USt)[^\n%]{0,40}?(\d{1,2}(?:,\d{1,2})?)\s*%",
                  flat, re.I)
    if m:
        feats.append(F("VATRate", "number", "VAT rate in %", de_number(m.group(1) or m.group(2)),
                       re.sub(r"\s+", " ", m.group(0))[:150]))
    have = {(f["name"], str(f["value"])) for f in feats}
    for k, f in (fields_from_azure or {}).items():          # fields an Azure analyzer returned (if any)
        v = _azure_value(f)
        if (k, str(v)) not in have:
            feats.append(F(k, f.get("type", "string"), "Field returned by the Azure analyzer (AI)", v, None))
    return feats


def _azure_value(f):
    if not isinstance(f, dict):
        return f
    t = f.get("type")
    if t == "object":
        return {k: _azure_value(v) for k, v in (f.get("valueObject") or {}).items()}
    if t == "array":
        return [_azure_value(v) for v in f.get("valueArray") or []]
    for k, v in f.items():
        if k.startswith("value"):
            return v
    return None


# ------------------------------------------------------------------ positions
def position_features(md):
    """Positions found the same way the material filter finds them (shared code -> same result)."""
    import llm_filter                                   # late import: llm_filter imports this module
    out = []
    section = None
    for ref, text, ctx, has_q in llm_filter.split_blocks(llm_filter.clean_markdown(md)):
        if ref.startswith("para"):
            continue
        lines = [l for l in text.split("\n") if l.strip()]
        body = "\n".join(lines[1:])
        q = llm_filter._find_quantity(text)
        qty = unit = ep = gp = None
        qline = None
        if q:
            qline = q.group(0)
            m = re.match(r"\s*([\d.,]+)\s*([A-Za-zäöüÄÖÜ]{1,6}[23²³]?)", qline)
            if m:
                qty, unit = de_number(m.group(1)), m.group(2)
            full = text[q.start():].split("\n")[0]
            mp = re.search(r"EP\s*([\d.]+,\d{2})", full)
            mg = re.search(r"GP\s*([\d.]+,\d{2})", full)
            ep = de_number(mp.group(1)) if mp else None
            gp = de_number(mg.group(1)) if mg else None
            body = body.replace(full.strip(), "").strip()
        text_lines = [l for l in body.split("\n") if l.strip()]
        short = text_lines[0] if text_lines else ""
        long_ = " ".join(text_lines[1:])
        check = None
        if qty is not None and ep is not None and gp is not None:
            check = "ok" if abs(qty * ep - gp) <= max(0.02, 0.005 * gp) else f"Quantity x UnitPrice = {qty * ep:.2f} ≠ {gp}"
        val = [F("PositionNumber", "string", "Position number (OZ)", ref),
               F("ShortText", "string", "Short text (first line of the description)", short),
               F("LongText", "string", "Rest of the description", long_ or None),
               F("Quantity", "number", "Ordered quantity", qty, qline),
               F("Unit", "string", "Unit of the quantity (St, m, m2, m3, psch …)", unit, qline),
               F("UnitPrice", "number", "Unit price in EUR (EP)", ep),
               F("TotalPrice", "number", "Total price in EUR (GP)", gp),
               F("PriceCheck", "string", "Quantity × UnitPrice compared with TotalPrice", check),
               F("PositionKind", "string", "Special kind of position (contingency, alternative …); empty = normal",
                 position_kind(text))]
        out.append(F(f"Position {ref}", "object", "Bill of quantities position", val, lines[0][:100]))
    return out


def other_tables(md, limit=40):
    """Tables that are not position lists (cover data, summaries …) as type 'table'."""
    out = []
    for i, t in enumerate(re.findall(r"<table[^>]*>.*?</table>", md, re.S), 1):
        rows = [r for r in parse_table(t) if any(r)]
        if any(row_position(r, {}) for r in rows):
            continue
        out.append(F(f"Table {i}", "table", f"Table {i} of the document ({len(rows)} rows)", rows[:limit]))
    return out


# ------------------------------------------------------------------ one document
def extract(result_json, file_name, origin):
    res = result_json.get("result") or result_json
    contents = res.get("contents", [])
    md = "\n".join(c.get("markdown", "") for c in contents)
    azure_fields = {}
    for c in contents:
        azure_fields.update(c.get("fields") or {})
    if origin == "local":                                 # GAEB fields are already in our own format
        gaeb = {k: v for k, v in azure_fields.items()}
        azure_fields = {}
    positions = position_features(md)
    doc = document_features(md, azure_fields)
    if origin == "local":
        names = {"ProjectNumber": "Project number (GAEB LblPrj)", "ProjectName": "Project name (GAEB NamePrj)",
                 "GaebPhase": "GAEB exchange phase (X83 = tender, X84 = offer …)"}
        doc = [F(k, "string", names.get(k, "GAEB field"), _azure_value(v), "GAEB file") for k, v in gaeb.items()] + \
              [f for f in doc if f["name"] not in gaeb]
    priced = [p for p in positions if p["value"][6]["value"] is not None]
    summary = {"positions": len(positions), "with_quantity": sum(p["value"][3]["value"] is not None for p in positions),
               "with_unit_price": sum(p["value"][5]["value"] is not None for p in positions),
               "with_total_price": len(priced),
               "sum_of_total_prices": round(sum(p["value"][6]["value"] for p in priced), 2) if priced else None,
               "price_check_failed": sum(1 for p in positions if (p["value"][7]["value"] or "ok") != "ok")}
    net = next((f["value"] for f in doc if f["name"] == "NetTotal"), None)
    if net is not None and summary["sum_of_total_prices"] is not None:
        diff = round(net - summary["sum_of_total_prices"], 2)
        summary["document_net_total"] = net
        summary["total_check"] = "ok – positions add up to the document total" if abs(diff) <= 0.05 else \
            f"positions add up to {summary['sum_of_total_prices']}, document says {net} (difference {diff}) – check"
    return {"document": file_name,
            "source": "Azure Content Understanding (" + res.get("analyzerId", "?") + ")" if origin == "azure"
            else "local GAEB reader (exact values from the XML / GAEB file)",
            "summary": summary,
            "features": doc,
            "positions": positions,
            "tables": other_tables(md) if origin == "azure" else []}


if __name__ == "__main__":
    for p in sys.argv[1:]:
        data = json.loads(Path(p).read_text(encoding="utf-8"))
        origin = "local" if (data.get("result") or {}).get("analyzerId") == "gaeb-reader" else "azure"
        print(json.dumps(extract(data, Path(p).name, origin), ensure_ascii=False, indent=1))