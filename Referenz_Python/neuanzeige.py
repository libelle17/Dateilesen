# Referenzlogik der neuen Diagnosenanzeige (Python) - Vorlage/Pruefstand fuer DiagAnzeige.bas
import re, csv, collections, os, datetime, sys, subprocess
from normalisierung import *
OUTD = "/DATA/eigene Dateien/Programmierung/Dateilesen/"
S = ""   # Arbeitsdateien (beispiel_rows.tsv, beispiel_geb.tsv, reihe.tsv) im aktuellen Ordner
def klasse(x): return (x % 15) % 10
# ---------- Praefixe/Labels
LET = "A-Za-zÄÖÜäöüß"
LAB = [("Z", re.compile(r'^\s*(?:Z\.\s?n\.|Zn\.|Zustand\s+nach(?![A-Za-zÄÖÜäöüß]))\s*[:,\-]?\s*', re.I)),
       ("V", re.compile(r'^\s*(?:V\.\s?a\.|Verdacht\s+auf(?![A-Za-zÄÖÜäöüß])|Verdacht(?![A-Za-zÄÖÜäöüß]))\s*[:,\-]?\s*', re.I)),
       ("G", re.compile(r'^\s*gesichert(?![A-Za-zÄÖÜäöüß])\s*[:,\-]?\s*', re.I)),
       ("A", re.compile(r'^\s*(?:ausgeschlossen(?![A-Za-zÄÖÜäöüß])|ausgeschl\.|ausgeschl(?![A-Za-zÄÖÜäöüß])|Ausschluss(?![A-Za-zÄÖÜäöüß]))\s*[:,\-]?\s*', re.I)),
       ("Z", re.compile(r'^\s*[\[(]\s*(?:wohl\s+)?Z\.?\s?n\.?\s*\??\s*[\])]\s*', re.I))]
def parse_prefix(t):
    labs = []
    while True:
        for k, rx in LAB:
            m = rx.match(t)
            if m and m.end() > 0: labs.append(k); t = t[m.end():]; break
        else: return labs, t.strip()
# ---------- Entscheidungstabelle
RULES = []
for r in list(csv.reader(open(OUTD + "Diagnose_Entscheidungstabelle.csv", encoding="utf-8-sig"), delimiter=";"))[1:]:
    RULES.append({"pre": r[1], "txt": r[2], "age": int(r[3]) if r[3] else None, "d": r[4], "prio": int(r[5])})
def like(pat, s):
    if "%" not in pat: return s == pat.lower()
    return re.fullmatch(".*".join(re.escape(x) for x in pat.lower().split("%")), s, re.S) is not None
def decide(icd, tnorm, age):
    best = None
    for r in RULES:
        if r["pre"] and not icd.startswith(r["pre"]): continue
        if r["txt"] and not like(r["txt"], tnorm): continue
        if r["age"] is not None and (age is None or age <= r["age"]): continue
        key = (len(r["pre"]), r["prio"])
        if best is None or key > best[0]: best = (key, r["d"])
    return best[1] if best else "je"
# ---------- Paarregeln (aktive Z.n. gegen aktive gesichert)
PAIR = []
for r in list(csv.reader(open(OUTD + "Diagnose_Paarregeln.csv", encoding="utf-8-sig"), delimiter=";"))[1:]:
    PAIR.append({"pre": r[0], "ord": int(r[1]), "cond": r[2], "code": r[3]})
def text_year(t):
    ys = [int(y) for y in re.findall(r"(?<!\d)((?:19|20)\d\d)(?!\d)", t)]
    ys += [(2000 + int(y) if int(y) <= 30 else 1900 + int(y)) for y in re.findall(r"(?<!\d)(?:0?[1-9]|1[0-2])/(\d\d)(?!\d)", t)]
    return max(ys) if ys else 0
def year_of(r): return text_year(r["txt"]) or (int(r["date"][:4]) if r["date"][:4].isdigit() else 0)
def cond_true(c, z, g):
    if c == "": return True
    if c == "GLEICH": return match(z, g) or match(g, z)
    if c == "GJUENGER": return year_of(g) > year_of(z)
    if c in ("SEITE", "SEITE_OHNE_AUSSENINNEN"):
        if not (z["sides"] and g["sides"] and (z["sides"] & g["sides"])): return False
        if c == "SEITE": return True
        zt, gt = z["text"].lower(), g["text"].lower()
        ai = lambda t: ("außen" in t or "aussen" in t, "innen" in t)
        return not ((ai(zt)[0] and ai(gt)[1]) or (ai(gt)[0] and ai(zt)[1]))
    if c.startswith("GTEXT:"): return like(c[6:], g["tnorm"])
    if c.startswith("ZTEXT:"): return like(c[6:], z["tnorm"])
    return False
def pair_code(z, g):
    cand = [p for p in PAIR if z["icd"].startswith(p["pre"]) and g["icd"].startswith(p["pre"])]
    if not cand: return "b"
    ml = max(len(p["pre"]) for p in cand)
    for p in sorted((p for p in cand if len(p["pre"]) == ml), key=lambda p: p["ord"]):
        if cond_true(p["cond"], z, g): return p["code"]
    return "b"
# ---------- Zeilen aufbereiten
reihe = {}
for l in open(S + "reihe.tsv", encoding="utf-8"):
    p = l.rstrip("\n").split("\t"); reihe[p[0]] = (None if p[1] == "NULL" else int(p[1]), None if p[2] == "NULL" else int(p[2]))
geb = {}
for l in open(S + "beispiel_geb.tsv", encoding="utf-8"):
    p = l.rstrip("\n").split("\t")
    try: geb[p[0]] = datetime.date.fromisoformat(p[1][:10])
    except Exception: pass
def age_of(pid):
    g = geb.get(pid); 
    if not g: return None
    t = datetime.date(2026, 9, 20); return t.year - g.year - ((t.month, t.day) < (g.month, g.day))
SEITE = {"1": "R", "2": "L", "3": "B"}
def tokens(t): return frozenset(re.findall(WC + "+", t.lower()))
def sim_rows(raws):
    out = []
    for r in raws:
        sid, pid, icd, st, kl, d, flag, txt, erl = r[:9]; st = int(st); kl = klasse(int(kl))
        if st == 3 or (kl == 4 and st == 4): continue                       # neue MODiagnosen
        sich = {1: "V", 2: "G", 3: "Z", 4: "A"}.get(kl, " ")
        if st == 4:
            if sich == "V" and txt[:4] != "V.a.": txt = "V.a. " + txt
            sich = "Z"
        s, dates, b0, b1, core, rest, sides = prep(txt, flag)
        labs, stripped = parse_prefix(txt)
        out.append({"sid": sid, "icd": icd, "icdk": icd if icd == "-1" else icd.rstrip("-"), "st": st, "sich": sich, "txt": txt, "labs": labs, "text": stripped, "seite": SEITE.get(flag, " "), "attr": erl, "erl": tokens(erl),
                    "obD": 0 if st == 1 else 1, "date": d, "dates": dates, "b0": b0, "b1": b1, "core": core, "rest": rest, "sides": sides,
                    "tnorm": clean(suffix(Dre.sub("", s))).lower(), "w1": frozenset(re.findall(WC + "+", b1)),
                    "rb": (lambda a, b: a.lower() if a != b else None)(clean(render(stripped, icd, " ", "", "")), clean(stripped))})
    return out
def certainty(r):
    if r["sich"] == "A" or "A" in r["labs"]: return "A"
    return "V" if (r["sich"] == "V" or "V" in r["labs"]) else "G"
def own_past(r): return r["sich"] == "Z" and r["st"] != 4 or "Z" in r["labs"]
def match(cl, ac, loose=False):
    if not (cl["icdk"] == ac["icdk"] or cl["icd"] == "-1" or (loose and cl["icd"][:3] == ac["icd"][:3] and not cl["rb"] and not ac["rb"])): return False
    if not (all(dcov(d, ac) for d in cl["dates"]) and cl["sides"] <= ac["sides"] and cl["erl"] <= ac["erl"]): return False
    if cl["b0"] == ac["b0"] or cl["b1"] == ac["b1"] or (cl["rb"] and cl["rb"] == ac["rb"]): return True
    if bool(cl["core"]) and cl["core"] == ac["core"] and cl["rest"] <= (ac["rest"] | ac["erl"]): return True
    return bool(cl["w1"]) and cl["w1"] <= (ac["w1"] | ac["erl"])        # Wortteilmenge (gleiche ICD); Erlaeuterung der aktiven Zeile zaehlt mit
def dcov(d, ac):
    # Datum aus dem Text der abgeschlossenen Zeile ist gedeckt, wenn es im Text der aktiven vorkommt oder alle Zahlengruppen in deren Erlaeuterung stehen
    return d in ac["dates"] or all(g in ac["erl"] for g in re.findall(r"\d+", d))
def lab(cert, past, vz=False):
    if cert == "A": return "Ausschluss "
    if vz: return "V.a. Z.n. "
    return ("Z.n. V.a. " if past else "V.a. ") if cert == "V" else ("Z.n. " if past else "")
def key_of(r): return (r["icdk"], r["rb"] or r["b1"], tuple(sorted(r["sides"])), tuple(sorted(r["dates"])), tuple(sorted(r["erl"])))
def latest(rows): return max(rows, key=lambda r: (r["date"], len(r["text"])))
def visible(rows, age):
    active = [r for r in rows if r["st"] in (1, 2, 5)]; closed = [r for r in rows if r["st"] == 4]
    grp = collections.OrderedDict()
    for r in active: grp.setdefault((key_of(r), certainty(r), own_past(r)), []).append(r)
    act_u = [latest(g) for g in grp.values()]
    coll = collections.OrderedDict(); keep = []
    for r in act_u:
        if own_past(r) or certainty(r) == "A": keep.append(r); continue
        coll.setdefault((r["icd"][:3], certainty(r), r["txt"][:17].lower()), []).append(r)
    act_u = keep + [latest(g) for g in coll.values()]
    # aktive Z.n.: weniger spezifische durch spezifischere aktive Z.n. derselben Diagnose verdraengen
    zs = [r for r in act_u if own_past(r) and certainty(r) != "A"]
    def zsup(x):
        for y in zs:
            if y is x or certainty(y) != certainty(x) or not match(x, y): continue
            if not match(y, x) or (y["date"], len(y["text"]), id(y)) > (x["date"], len(x["text"]), id(x)): return True
        return False
    zdrop = {id(x) for x in zs if zsup(x)}
    # aktive gesicherte Zeilen: weniger spezifische durch spezifischere derselben ICD-3-Stelle verdraengen
    gs = [r for r in act_u if not own_past(r) and certainty(r) == "G"]
    def gsup(x):
        for y in gs:
            if y is x or not match(x, y, True): continue
            if not match(y, x, True) or (y["date"], len(y["text"]), id(y)) > (x["date"], len(x["text"]), id(x)): return True
        return False
    zdrop |= {id(x) for x in gs if gsup(x)}
    res = []   # (row, label, quelle)
    for r in act_u:
        if id(r) in zdrop: continue
        vz = r["labs"][:2] == ["V", "Z"]
        res.append((r, "V.a. Z.n. " if vz else lab(certainty(r), own_past(r)), "akt"))
    # abgeschlossene: Kopien zusammenfassen, Gegenstueck pruefen, Tabelle
    cg = collections.OrderedDict()
    for r in closed: cg.setdefault((key_of(r), certainty(r)), []).append(r)
    cl_u = [latest(g) for g in cg.values()]
    rest = [c for c in cl_u if not any(match(c, a) for a in active)]
    def superseded(x):
        for y in rest:
            if y is x or not match(x, y): continue
            if not match(y, x) or (y["date"], len(y["text"]), id(y)) > (x["date"], len(x["text"]), id(x)): return True
        return False
    rest = [c for c in rest if not superseded(c)]
    once = collections.OrderedDict()
    for c in rest:
        d = decide(c["icd"], c["tnorm"], age)
        if d == "n": continue
        if d == "je": res.append((c, lab(certainty(c), True), "abg"))
        else: once.setdefault((c["icd"][:3], certainty(c)), []).append((c, d))
    for k, g in once.items():
        c = latest([x for x, _ in g]); d = next(dd for x, dd in g if x is c)
        res.append((c, lab(certainty(c), d == "jz" or "Z" in c["labs"]), "abg"))
    # Paarregeln (diagpaare) ueber alle angezeigten Z.n.- und gesichert-Zeilen
    zn = [(r, l, q) for r, l, q in res if certainty(r) == "G" and l == "Z.n. "]
    ges = [(r, l, q) for r, l, q in res if certainty(r) == "G" and l == ""]
    hidden = set(); bZ = collections.defaultdict(set); bG = collections.defaultdict(set); aZ = set(); aG = set(); paired = set()
    for z, _, _ in zn:
        for g, _, _ in ges:
            if z["icd"][:3] != g["icd"][:3] or z["icd"] == "-1": continue
            paired.add(id(z)); paired.add(id(g)); code = pair_code(z, g); yz, yg = year_of(z), year_of(g)
            if code == "g": hidden.add(id(z))
            elif code == "z": hidden.add(id(g))
            elif code == "j": hidden.add(id(z) if yg >= yz else id(g))
            elif code == "b": bZ[z["icd"][:3]].add(id(z)); bG[g["icd"][:3]].add(id(g))
            elif code == "a": aZ.add(id(z)); aG.add(id(g))
            elif code.isdigit():
                if yg - yz >= int(code): aZ.add(id(z)); aG.add(id(g))
                else: hidden.add(id(z) if yg > yz else id(g))
    byid = {id(r): r for r, _, _ in zn + ges}
    for grp_ in (bZ, bG):
        for i3, ids in grp_.items():
            ids = [i for i in ids if i not in hidden and i not in aZ and i not in aG]
            if len(ids) > 1:
                keep_ = latest([byid[i] for i in ids])
                for i in ids:
                    if byid[i] is not keep_: hidden.add(i)
    single = collections.OrderedDict()
    for z, _, q in zn:
        if q != "akt" or id(z) in paired or id(z) in hidden: continue
        single.setdefault(z["icd"][:3], []).append(z)
    for i3, g in single.items():
        d = decide(g[0]["icd"], g[0]["tnorm"], age)
        if d != "je":
            k = latest(g)
            for z in g:
                if z is not k: hidden.add(id(z))
    # L89: aktive gesicherte Wagner-0-Zeile deckt Z.n.-Wagner-0-Zeilen ab, die keine andere Seite betreffen
    def wagner0(r): return r["icd"][:3] == "L89" and len(r["icd"]) >= 5 and r["icd"][4] in "01"
    g0 = [r for r, l, q in res if l == "" and wagner0(r) and id(r) not in hidden]
    if g0:
        covered = set().union(*(r["sides"] for r in g0))
        for r, l, q in res:
            if l == "Z.n. " and wagner0(r) and id(r) not in hidden and r["sides"] <= covered:
                hidden.add(id(r))
    return [(r, l) for r, l, q in res if id(r) not in hidden]
# ---------- Rendering (Umformungen wie in MachDiagnosen)
def render(text, icd, seite, attr, dmseit):
    t = text
    if icd[:2] == "E1":
        if "ankr" in t or "ekund" in t: pass
        else:
            t = {"0": "Diabetes mellitus Typ 1", "1": "Diabetes mellitus Typ 2", "2": "Diabetes mellitus in Verbindung mit Fehl- oder Mangelernährung", "3": "Diabetes mellitus (sekundär)", "4": "Diabetes mellitus"}.get(icd[2:3], t)
            if dmseit != "": t += " seit " + dmseit
    if icd[:5] == "O24.4": t = "Gestationsdiabetes"
    if icd[:5] == "N08.3" and "Glomeruläre Krankheiten" in t: t = "Diabetische Nephropathie"
    if "Niereninsuff." in t: t = t.replace("Niereninsuff.", "Niereninsuffizienz")
    if icd[:5] == "G99.0" and "Autonome Neuropathie bei endokrinen" in t: t = "Diabetische autonome Neuropathie"
    if icd[:5] in ("I79.9", "I79.2") and "Periphere Angiopathie bei anderenorts" in t: t = "Periphere Angiopathie"
    if icd[:5] == "M14.6" and "Neuropathische Arthropathie " in t: t = "Diabetische Osteoarthropathie"
    if icd[:5] == "M36.8" and "anderenorts" in t: t = "Diabetische Bindegewebserkrankung"
    if icd[:3] == "E66": t = "Übergewicht"
    if icd[:5] in ("K76.0", "K76.9", "K71.6", "K71.7", "K77.8") and "nbek" not in t and "nklar" not in t and "nklar" not in attr: t = "Hepatopathie"
    if icd[:3] == "L89":
        t = "Diabetisches Fußsyndrom "
        if icd[4:5] < "9": t += " im Stadium Wagner "
        t += {"0": "0", "1": "0", "2": "1", "3": "2", "4": "3"}.get(icd[4:5], "")
    if seite == "R" and not t.endswith(" re"): t += " re"
    elif seite == "L" and not t.endswith(" li"): t += " li"
    elif seite == "B" and "bds." not in t: t += " bds."
    if attr != "": t += " (" + attr + ")"
    return t.replace(", nicht näher bezeichnet", "")
def lines_for(pid, vis, dmseit):
    items = []
    for r, label in vis:
        icd = r["icd"]
        if icd == "" or re.match(r"Z25", icd): continue
        if re.match(r"(M20\.|M21\.|Q66\.|B35\.|K02\.)", icd) or icd in ("L84", "R26.8", "R29.6", "R52.2", "R68.8"): continue
        rf, gi2 = reihe.get(icd, (None, None))
        txt = render(r["text"], icd, r["seite"], r["attr"], dmseit)
        if label: txt = re.sub(r"Diab\.", "diab.", txt).replace("Diabeti", "diabeti").replace("Diabet.", "diabet.").replace("Auton", "auton")
        line = label + txt
        if r["obD"] == 0: line = "(q) " + line
        marker = "Z" if label.startswith("Z.n.") else ("V" if label.startswith("V.a.") else "")
        indent = " " if (rf == 1 and gi2 is not None and 50 < gi2 < 97) else ""
        items.append(((0 if rf is None else 1, rf or 0, 1 if gi2 is None else 0, -(gi2 or 0), icd), indent + line + "\t[" + icd + marker + "]"))
    items.sort(key=lambda x: x[0])
    out = []; seen = set()
    for k, l in items:
        kk = (k[4][:3], l.split("\t")[0])          # gleiche Anzeigezeile in derselben ICD-3-Stelle nur einmal
        if kk in seen: continue
        seen.add(kk); out.append(l)
    return out
if __name__ == "__main__":
    raws = collections.defaultdict(list)
    for l in open(S + "beispiel_rows.tsv", encoding="utf-8"):
        p = l.rstrip("\n").split("\t"); p += [""] * (9 - len(p)); raws[p[1]].append(p)
    dm = {}
    for l in subprocess.run(["mariadb", "quelle", "--batch", "--skip-column-names", "-e", "SELECT pat_id, COALESCE(`diabetes seit`,0) FROM anamnesebogen"], capture_output=True, text=True).stdout.split("\n"):
        if "\t" in l: a, b = l.split("\t"); dm[a] = b
    outd = OUTD + "Beispiel_Diagnosestrings_neu"; os.makedirs(outd, exist_ok=True)
    files = [f for f in os.listdir(OUTD + "Beispiel_Diagnosestrings") if f.endswith(".txt")]
    stat = collections.Counter(); tot_old = tot_new = 0
    for f in sorted(files):
        icd3, pid = f[:-4].split("_")
        rows = sim_rows(raws.get(pid, []))
        new = lines_for(pid, visible(rows, age_of(pid)), dm.get(pid, ""))
        old = [x for x in open(OUTD + "Beispiel_Diagnosestrings/" + f, encoding="utf-8").read().split("\n") if x]
        tot_old += len(old); tot_new += len(new)
        stat["weniger" if len(new) < len(old) else "gleich" if len(new) == len(old) else "mehr"] += 1
        open(os.path.join(outd, f), "w", encoding="utf-8", newline="\n").write("\n".join(new) + "\n")
    print("Dateien:", len(files), "| Zeilen alt:", tot_old, "-> neu:", tot_new, "|", dict(stat))
    r155 = sim_rows(raws["155"]); print("Pat 155: Zeilen simuliert (quelle-Form):", len(r155), "| davon abgeschlossen:", sum(1 for r in r155 if r["st"] == 4), "| Zeilen im neuen String:", len(lines_for("155", visible(r155, age_of("155")), dm.get("155", ""))))
