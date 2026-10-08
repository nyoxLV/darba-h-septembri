# -*- coding: utf-8 -*-
"""Papildu kolonnas dienu grupām (Artūra norāde 08.10: "izveidot vairāk kolonnas").
py -3.12 insert_day_cols.py <ieeja.xlsm> <izeja.xlsm> 4 7 15 21 29   (dienas, kurām +1 kolonna)"""
import sys, zipfile
sys.path.insert(0, __import__("os").path.dirname(__file__))
from xl_engine import SST, Sheet, Styles, day_groups, insert_columns, same, n2col

SRC, DST, DAYS = sys.argv[1], sys.argv[2], [int(x) for x in sys.argv[3:]]
zin = zipfile.ZipFile(SRC)
parts = {i.filename: zin.read(i.filename).decode("utf-8") for i in zin.infolist()
         if i.filename.endswith((".xml", ".vml", ".rels"))}
sh0 = Sheet(parts["xl/worksheets/sheet1.xml"], SST(parts["xl/sharedStrings.xml"]),
            Styles(parts["xl/styles.xml"], parts["xl/theme/theme1.xml"]))
g = day_groups(sh0)
inserts = [(g[d + 1][0], 1) for d in DAYS]
print("ievieto pirms kolonnām:", [(d, n2col(p)) for d, (p, _) in zip(DAYS, inserts)])
new = insert_columns(parts, inserts)
with zipfile.ZipFile(DST, "w") as zout:
    for info in zin.infolist():
        if info.filename == "xl/calcChain.xml":
            continue  # Excel to izveido no jauna
        data = new[info.filename].encode("utf-8") if info.filename in new else zin.read(info.filename)
        zout.writestr(info, data, compress_type=zipfile.ZIP_DEFLATED)
# pārbaude: visas vecās vērtības == jaunās pārbīdītās vērtības
fc = lambda c: c + sum(k for p, k in inserts if c >= p)
sh1 = Sheet(new["xl/worksheets/sheet1.xml"], SST(new["xl/sharedStrings.xml"]), Styles(new["xl/styles.xml"], new["xl/theme/theme1.xml"]))
v0, v1 = sh0.recalc(), sh1.recalc()
bad_cache = [k for k, v in v1.items() if not same(sh1.cached[k], v)]
bad_map = [(n2col(c) + str(r), v0[(r, c)], v1.get((r, fc(c)))) for (r, c) in v0 if not same(v0[(r, c)], v1.get((r, fc(c))))]
const_bad = [(r, c) for (r, c), cell in sh0.cells.items() if cell.f is None and not same(cell.v, sh1.cells[(r, fc(c))].v)]
g1 = day_groups(sh1)
print("dienu platumi:", {d: len(g1[d]) for d in DAYS}, "| formulas:", len(v0), "->", len(v1))
print("kešs != pārrēķins:", len(bad_cache), "| vecās != jaunās vērtības:", len(bad_map), bad_map[:5], "| konstantes:", len(const_bad))
