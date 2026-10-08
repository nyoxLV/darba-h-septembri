# -*- coding: utf-8 -*-
"""Visu darbinieku audits (tikai lasa): stundas, krāsas pret piezīmēm, brīvdienas, statusi, EDLUS.

Pārbauda rindās 14-131: stundas bez objekta krāsas; piezīmes 1. rindas objekts != šūnas krāsas
objekts; stundas / statusi brīvdienās; diena > 12 h; statuss + stundas vienā dienā;
"Stundu pārbaude" != 0; EDLUS spoguļrinda != pamatrinda. Neko nemaina.

python3 audit_all.py <fails.xlsm> [gads mēnesis] [--alias aliases.json]"""
import html
import json
import re
import sys
import unicodedata
import zipfile
from datetime import date

sys.path.insert(0, __import__("os").path.dirname(__file__))
from xl_engine import SST, Sheet, Styles, day_groups, n2col

F = sys.argv[1]
pos = [a for a in sys.argv[2:] if a.isdigit()]
Y, M = (int(pos[0]), int(pos[1])) if len(pos) >= 2 else (2026, 9)
z = zipfile.ZipFile(F)
rd = lambda p: z.read(p).decode("utf-8")
st = Styles(rd("xl/styles.xml"), rd("xl/theme/theme1.xml"))
sh = Sheet(rd("xl/worksheets/sheet1.xml"), SST(rd("xl/sharedStrings.xml")), st)
com = {}
for m in re.finditer(r'<comment ref="([A-Z]+\d+)"[^>]*>(.*?)</comment>', rd("xl/comments1.xml"), re.S):
    com[m.group(1)] = html.unescape(re.sub(r"<[^>]+>", "", m.group(2)))
g = day_groups(sh)
ndays = (date(Y, M % 12 + 1, 1) - date(Y, M, 1)).days if M < 12 else 31
WE = {d for d in range(1, ndays + 1) if date(Y, M, d).weekday() >= 5}
hdr = {}
dup = []
for c in range(1, 600):
    if str(sh.value(7, c) or "").strip() == "st":
        rgb = st.interior_rgb(sh.cells[(8, c)].s)
        if rgb in hdr:
            dup.append((rgb, hdr[rgb], sh.value(8, c)))
        hdr[rgb] = str(sh.value(8, c)).strip()
COLS = {str(sh.value(8, c)).strip(): c for c in range(1, 700) if sh.value(8, c) in ("Stundu pārbaude", "St.")}
vals = sh.recalc()


def norm(t):
    t = unicodedata.normalize("NFKD", str(t or ""))
    t = "".join(ch for ch in t if not unicodedata.combining(ch)).lower()
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9 ]+", " ", t)).strip()


STOP = {"un", "iela", "privatmaja", "projekti", "objekti", "latvijas", "valsts", "mezi", "kristapa", "darbinieks",
        "akumulatoru", "gatavosana", "palidzesana", "tamesana", "iepirkuma", "energo", "serviss", "apkope", "transeja",
        "pagrabs", "noliktava", "depo", "ka", "ir", "bet", "nav", "darba", "8h", "dzeki", "edlus"}
# objekta nosaukums -> atslēgvārdi piezīmēs (neobligāts JSON fails: --alias aliases.json)
ALIAS = json.load(open(sys.argv[sys.argv.index("--alias") + 1], encoding="utf-8")) if "--alias" in sys.argv else {}



def keys(name):
    if name in ALIAS:
        return ALIAS[name]
    return [w for w in norm(name).split() if len(w) > 2 and w not in STOP] or [norm(name)]


def obj_of(text):
    """Kuri objekti minēti piezīmes 1. rindā."""
    first = norm(text.split("\n")[0])
    return [n for n in hdr.values() if any(k in first for k in keys(n))]


issues = {k: [] for k in ("krāsa", "piezīme≠krāsa", "brīvdiena", "statuss brīvdienā", ">12h", "statuss+stundas", "pārbaude", "EDLUS")}
tot = {}
rows = {}
for r in range(14, 132):
    name = str(sh.value(r, 2) or "").strip()
    if not name:
        continue
    rows[name] = r
    tot[name] = vals.get((r, COLS["St."])) or 0
    chk = vals.get((r, COLS["Stundu pārbaude"])) or 0
    if abs(chk) > 1e-9:
        issues["pārbaude"].append(f"{name}: {chk:g}")
    for d in range(1, ndays + 1):
        nums, stats = [], []
        for c in g[d]:
            v = sh.value(r, c)
            if v in (None, ""):
                continue
            ref = f"{n2col(c)}{r}"
            if isinstance(v, (int, float)):
                nums.append(v)
                rgb = st.interior_rgb(sh.cells[(r, c)].s)
                if rgb not in hdr:
                    issues["krāsa"].append(f"{name} {d:02d}.{M:02d} {ref}={v:g} krāsa {rgb}")
                elif ref in com:
                    found = obj_of(com[ref])
                    if found and hdr[rgb] not in found:
                        issues["piezīme≠krāsa"].append(f"{name} {d:02d}.{M:02d} {ref}={v:g}: krāsa '{hdr[rgb]}', piezīme '{com[ref].splitlines()[0][:70]}'")
            else:
                stats.append(str(v).strip())
        if d in WE and nums:
            issues["brīvdiena"].append(f"{name} {d:02d}.{M:02d}: {sum(nums):g} h")
        if d in WE and stats:
            issues["statuss brīvdienā"].append(f"{name} {d:02d}.{M:02d}: {stats}")
        if sum(nums) > 12:
            issues[">12h"].append(f"{name} {d:02d}.{M:02d}: {sum(nums):g} h")
        if nums and stats:
            issues["statuss+stundas"].append(f"{name} {d:02d}.{M:02d}: {stats} + {sum(nums):g} h")

# EDLUS spoguļrindas pret pamatrindām
for name, r in rows.items():
    if "EDLUS" not in name:
        continue
    base = re.sub(r"\s*-?\s*EDLUS\s*$", "", name).strip()
    cand = [n for n in rows if n != name and norm(n) == norm(base)]
    if not cand:
        issues["EDLUS"].append(f"{name}: nav pamatrindas")
        continue
    rb = rows[cand[0]]
    for d in range(1, ndays + 1):
        def day(rr):
            n = sum(v for c in g[d] if isinstance((v := sh.value(rr, c)), (int, float)))
            s = sorted(str(v).strip() for c in g[d] if isinstance((v := sh.value(rr, c)), str) and v.strip())
            return n, s
        a, b = day(rb), day(r)
        if a[1] != b[1] or (b[0] and abs(a[0] - b[0]) > 1e-9):
            issues["EDLUS"].append(f"{cand[0]} {d:02d}.{M:02d}: pamatrinda {a[0]:g} h {a[1]} | EDLUS {b[0]:g} h {b[1]}")

print("dublētas galvenes krāsas:", dup)
for k, v in issues.items():
    print(f"\n### {k} ({len(v)})")
    for x in v:
        print("  ", x)
print("\n### St. pa personām (≠0)")
for n, t in tot.items():
    if t:
        print(f"   {t:7g}  {n}")
