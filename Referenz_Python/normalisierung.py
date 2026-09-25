# Normalisierung der Diagnosetexte fuer den Abgleich abgeschlossen/aktiv (Referenzlogik, Python)
# Ausschnitt aus der Auswertung vom 20.9.2026; Abkuerzungen/Synonyme in seed3.py
import re, csv, collections, sys
from seed3 import SEED
S = ""
OUT = "/DATA/eigene Dateien/Programmierung/Dateilesen/"
WC = r"[\wäöüß]"
P0 = re.compile(r'^((Z\.n\.|V\.a\.|gesichert|ausgeschl\.)\s*)+', re.I)
def n0(t): return re.sub(r'\s+', ' ', P0.sub('', t)).strip().lower()
P1 = re.compile(r'^\s*(?:'
    r'[\[(]\s*(?:wohl\s+)?(?:Z\.?\s?n\.?|V\.?\s?a\.?|[ZVGA])\s*\??\s*[\])]'
    r'|[\[(][A-Z]\d\d[\dA-Z.\-+*!]*[\])]'
    r'|[A-Z]\d\d\.\d{1,2}[\dA-Z]?\s'
    r'|Z\.?\s?n\.?(?=\s)|Zn\.|Zustand\s+nach|V\.\s?a\.|Verdacht\s+auf|Verdacht|gesichert|ausgeschl\.?|Ausschluss|ausgeschlossen|DD\b\.?'
    r')\s*[:,\-]?\s*', re.I)
def strip_pre(t):
    prev = None
    while prev != t: prev = t; t = P1.sub('', t, count=1)
    return t
def suffix(t):
    t = re.sub(r'\s*\[.*$', '', t); t = re.sub(r'\s*#+\s*$', '', t)
    return re.sub(r'[\s,;:.\-]+$', '', t)
DT = r'(?:\d{1,2}\.\d{1,2}\.(?:\d{2,4})?|\d{1,2}/\d{2,4}|(?:19|20)\d\d)'
DTre = re.compile(DT); Dre = re.compile(r'\s*(?:(?:seit|bis|ab|vor|am|ED)\s+)?' + DT + r'\b\.?', re.I)
def clean(t): return re.sub(r'\s+', ' ', t).strip()
SIDE_RX = re.compile(r'(?<![\wäöüß])(links|li\.?|rechts|re\.?|beidseits|beidseitig\w*|bds\.?)(?![\wäöüß])', re.I)
def text_sides(t):
    out = set()
    for m in SIDE_RX.finditer(t):
        w = m.group(1).lower().rstrip(".")
        out |= {"links"} if w in ("links", "li") else {"rechts"} if w in ("rechts", "re") else {"links", "rechts"}
    return out
FLAG = {"1": {"rechts"}, "2": {"links"}, "3": {"links", "rechts"}}
# ---- Abkuerzungs-/Synonymliste
def esc(f): return re.escape(f).replace(r"\.", r"\.\s*").replace(r"\ ", r"\s+")
def form_body(f):
    f = f.strip().lower().replace("ß", "ss"); comp = f.startswith("-"); f = f.lstrip("-")
    body = esc(f[:-1]) + WC + "*" if f.endswith("*") else esc(f.rstrip(".")) + (r"\.?" if f.endswith(".") else "")
    return comp, body
def form_rx(f):
    comp, body = form_body(f)
    pre = r"(?<=[\wäöüß])" if comp else r"(?<![\wäöüß])"
    return re.compile(pre + "(?:" + body + r")(?![\wäöüß])", re.I)
ABK = []
for typ, canon, forms, note, kind in SEED:
    if kind == "regex": frx = [(forms, re.compile(forms, re.I))]; whole = frx[0][1]
    else:
        fl = [x for x in forms.split("|") if x.strip()]; frx = [(x, form_rx(x)) for x in fl]
        comp = [form_body(x)[1] for x in fl if x.strip().startswith("-")]; plain = [form_body(x)[1] for x in fl if not x.strip().startswith("-")]
        parts = ([r"(?<=[\wäöüß])(?:" + "|".join(sorted(comp, key=len, reverse=True)) + ")"] if comp else []) + ([r"(?<![\wäöüß])(?:" + "|".join(sorted(plain, key=len, reverse=True)) + ")"] if plain else [])
        whole = re.compile("(?:" + "|".join(parts) + r")(?![\wäöüß])", re.I)
    ABK.append((typ, canon, forms, note, kind, whole, frx))
form_hits = collections.Counter()
def norm_abk(t, count=False):
    t = t.lower().replace("ß", "ss").replace("-", " ")
    for i, (typ, canon, forms, note, kind, whole, frx) in enumerate(ABK):
        if whole.search(t):
            if count:
                for lab, rx in frx:
                    if rx.search(t): form_hits[(i, lab)] += 1
            t = whole.sub(canon, t)
    t = re.sub(r'\s+,', ',', t); t = re.sub(r'(,\s*)+', ', ', t)
    return clean(suffix(t))
def prep(raw, flag, count=False):
    s = clean(suffix(strip_pre(raw)))
    dates = frozenset(DTre.findall(s))
    t = Dre.sub('', s)
    sides = frozenset(text_sides(t) | FLAG.get(flag, set()))
    t2 = clean(SIDE_RX.sub(' ', t))
    b0 = clean(suffix(t2)).lower()
    b1 = norm_abk(t2, count)
    core1 = clean(suffix(re.split(r'\s*[,(]', b1, maxsplit=1)[0]))
    rest = frozenset(re.findall(WC + "+", b1[len(core1):]))
    return s, dates, b0, b1, core1, rest, sides

