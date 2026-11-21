"""
gaeb_reader.py – read GAEB bill-of-quantities files WITHOUT Azure (exact, no OCR)
=================================================================================
Supported
  * GAEB DA XML   (.X81 … .X86, .X31 …, or .xml with a <GAEB> root) – complete support
  * GAEB 90       (.D81 … .D86, fixed 80-column text lines)          – best effort
                  record 21 = position (number, quantity, unit), 25 = short text, 26 = long text
Not supported (reported, not guessed): GAEB 2000 (.P8x)

Each file is turned into the SAME structure as an Azure Content Understanding result:
    {"status": "Succeeded", "result": {"analyzerId": "gaeb-reader", "contents": [
        {"kind": "gaeb", "markdown": "...one block per position...", "fields": {...}}]}}
so llm_filter.py treats it exactly like an analysed PDF. Every position is written as
    05.01.36
    <short text>
    <long text>
    175 St  EP 12,50  GP 2.187,50        (EP/GP only if the file contains prices)
"""
import re
import xml.etree.ElementTree as ET
from pathlib import Path

GAEB_XML_EXT = re.compile(r"\.x\d{2}$", re.I)
GAEB_90_EXT = re.compile(r"\.d\d{2}$", re.I)
GAEB_2000_EXT = re.compile(r"\.p\d{2}$", re.I)


def is_gaeb(path):
    p = str(path)
    if GAEB_XML_EXT.search(p) or GAEB_90_EXT.search(p) or GAEB_2000_EXT.search(p):
        return True
    if p.lower().endswith(".xml"):
        try:
            head = Path(path).read_bytes()[:2000].decode("utf-8", "ignore")
            return "<GAEB" in head
        except OSError:
            return False
    return False


def _num_de(x, decimals=3):
    """1.8 -> '1,8' ; 175.0 -> '175' (German format, no thousands separator)."""
    if x is None:
        return None
    s = f"{x:.{decimals}f}".rstrip("0").rstrip(".")
    return s.replace(".", ",")


def _money_de(x):
    if x is None:
        return None
    s = f"{x:,.2f}"                                   # 2,187.50
    return s.replace(",", "§").replace(".", ",").replace("§", ".")


def _result(markdown, fields, warnings):
    return {"status": "Succeeded",
            "result": {"analyzerId": "gaeb-reader", "apiVersion": "local", "warnings": warnings,
                       "contents": [{"kind": "gaeb", "markdown": markdown,
                                     "fields": {k: {"type": "string", "valueString": v}
                                                for k, v in fields.items() if v}}]}}


# ======================================================================= GAEB DA XML
def _local(tag):
    return tag.rsplit("}", 1)[-1]


def _child(el, name):
    for c in el:
        if _local(c.tag) == name:
            return c
    return None


def _find(el, path):
    """namespace-independent path like 'Description/CompleteText/DetailTxt'"""
    for part in path.split("/"):
        if el is None:
            return None
        el = _child(el, part)
    return el


def _text(el):
    """All text inside an element; each <p> paragraph becomes one line."""
    if el is None:
        return ""
    paras = [p for p in el.iter() if _local(p.tag) == "p"]
    chunks = ["".join(p.itertext()) for p in paras] if paras else ["".join(el.itertext())]
    return "\n".join(re.sub(r"[ \t]+", " ", c).strip() for c in chunks if c.strip())


def _float(el):
    try:
        return float(el.text.strip()) if el is not None and el.text and el.text.strip() else None
    except ValueError:
        return None


def read_gaeb_xml(path):
    warnings = []
    root = ET.parse(path).getroot()
    if _local(root.tag) != "GAEB":
        raise ValueError("not a GAEB DA XML file (root element is not <GAEB>)")
    prj = _child(root, "PrjInfo")
    fields = {"ProjectNumber": _text(_child(prj, "LblPrj")) if prj is not None else "",
              "ProjectName": _text(_child(prj, "NamePrj")) if prj is not None else ""}
    award = _child(root, "Award")
    dp = _text(_child(award, "DP")) if award is not None else ""
    fields["GaebPhase"] = f"X{dp}" if dp else ""
    blocks, n_items = [], 0

    def walk_body(body, prefix):
        nonlocal n_items
        for el in body:
            tag = _local(el.tag)
            if tag == "BoQCtgy":
                no = ".".join(p for p in prefix + [el.get("RNoPart", "")] if p)
                title = _text(_child(el, "LblTx"))
                blocks.append(f"## {no} {title}".strip())
                sub = _child(el, "BoQBody")
                if sub is not None:
                    walk_body(sub, prefix + [el.get("RNoPart", "")])
            elif tag == "Itemlist":
                for it in el:
                    if _local(it.tag) == "Item":
                        add_item(it, prefix)
                    elif _local(it.tag) == "Remark":                 # Hinweistext
                        t = _text(_find(it, "Description/CompleteText/DetailTxt"))
                        if t:
                            blocks.append(f"Hinweis: {t}")

    def add_item(it, prefix):
        nonlocal n_items
        if _child(it, "NotAppl") is not None:                       # "Position entfällt"
            return
        no = ".".join(p for p in prefix + [it.get("RNoPart", "")] if p)
        short = _text(_find(it, "Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt")) \
            or _text(_find(it, "Description/OutlineText/OutlTxt/TextOutlTxt"))
        long_ = _text(_find(it, "Description/CompleteText/DetailTxt"))
        qty, qu = _float(_child(it, "Qty")), _text(_child(it, "QU")) or "?"
        up, total = _float(_child(it, "UP")), _float(_child(it, "IT"))
        marks = []
        if _child(it, "Provis") is not None:
            marks.append("[Bedarfsposition / contingency item – only ordered if needed]")
        if _child(it, "ALNGroupNo") is not None and _child(it, "ALNSerNo") is not None:
            marks.append("[Alternativposition – alternative to a basic item, do not add to it]")
        if _child(it, "LumpSumItem") is not None:
            marks.append("[Pauschalposition / lump sum]")
        if _child(it, "QtyTBD") is not None:
            marks.append("[Menge vom Bieter – quantity to be given by the bidder]")
        qline = (f"{_num_de(qty)} {qu}" if qty is not None else f"(keine Menge) {qu}")
        qline += f"  EP {_money_de(up)}" if up is not None else "  EP ......"
        qline += f"  GP {_money_de(total)}" if total is not None else "  GP ......"
        blocks.append("\n".join(x for x in [no, short, long_, *marks, qline] if x))
        n_items += 1

    if award is None or _child(award, "BoQ") is None:
        raise ValueError("no <Award>/<BoQ> in the file")
    for boq in award:
        if _local(boq.tag) != "BoQ":
            continue
        name = _text(_find(boq, "BoQInfo/Name")) or _text(_find(boq, "BoQInfo/LblTx"))
        if name:
            blocks.append(f"# Leistungsverzeichnis {name}")
        body = _child(boq, "BoQBody")
        if body is not None:
            walk_body(body, [])
    if n_items == 0:
        warnings.append("GAEB XML contains no positions")
    return _result("\n\n".join(blocks), fields, warnings), n_items


# ======================================================================= GAEB 90
def read_gaeb90(path):
    """Best effort for GAEB 90 (.D8x). Lines are 80 characters, columns 1-2 = record type,
    columns 75-80 = line counter. Record 21: position number (cols 3-11) and quantity
    (11 digits, 3 implied decimals) followed by the unit; 25 = short text; 26 = long text."""
    raw = Path(path).read_bytes()
    for enc in ("cp437", "cp850", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    warnings = ["GAEB 90 is read best-effort: please spot-check quantities"]
    items, cur = [], None
    for line in text.splitlines():
        if len(line) < 2:
            continue
        rec, body = line[:2], line[2:74]
        if rec == "21":
            digits = re.match(r"[\d ]*", body[:9]).group(0)          # number groups only, no flag letters
            oz = ".".join(digits.split())
            m = re.search(r"(\d{11})\s*([^\d\s]\S{0,3})", body[9:])
            qty = int(m.group(1)) / 1000 if m else None
            unit = m.group(2) if m else "?"
            cur = {"no": oz, "qty": qty, "unit": unit, "short": [], "long": []}
            items.append(cur)
        elif rec == "25" and cur is not None:
            cur["short"].append(body.rstrip())
        elif rec == "26" and cur is not None:
            cur["long"].append(body.rstrip())
    blocks = []
    for it in items:
        q = f"{_num_de(it['qty'])} {it['unit']}" if it["qty"] is not None else f"(keine Menge) {it['unit']}"
        blocks.append("\n".join(x for x in [it["no"], " ".join(s.strip() for s in it["short"]),
                                             "\n".join(l.strip() for l in it["long"] if l.strip()),
                                             f"{q}  EP ......  GP ......"] if x))
    if not items:
        raise ValueError("no position records (type 21) found – not a GAEB 90 file?")
    return _result("\n\n".join(blocks), {}, warnings), len(items)


def read_gaeb(path):
    """Returns (azure_like_result, number_of_positions). Raises ValueError if unreadable."""
    p = str(path)
    if GAEB_2000_EXT.search(p):
        raise ValueError("GAEB 2000 (.P8x) is not supported – export the LV as GAEB DA XML (.X8x)")
    if GAEB_90_EXT.search(p):
        return read_gaeb90(path)
    return read_gaeb_xml(path)
