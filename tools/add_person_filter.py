# -*- coding: utf-8 -*-
"""Personu filtrs (AutoFilter) darba stundu lapai: redzami tikai tie, kas mēnesī strādāja.

Excel: B8 (Uzvārds,vārds) bultiņa -> atķeksē / ieķeksē personas; "Notīrīt filtru" rāda visus.
Redzamas paliek rindas, kurās dienu režģī ir kaut kas (stundas vai statuss: Atv., SA, SB ...),
plus --keep vārdi (strādāja, bet dati vēl nav ievadīti). Pārējās rindas paslēptas (hidden),
nevis dzēstas: formulas, kopsummas un piezīmes nemainās. Rindas ar tukšu B (9.-13.) paliek redzamas.

python3 add_person_filter.py <ieeja.xlsm> <izeja.xlsm> [--keep "Vārds Uzvārds;..."] [--names-from <fails.xlsm>]
  --names-from: ņem redzamo personu sarakstu no cita faila filtra (piem. oktobrim - no septembra).
"""
from __future__ import annotations

import html
import re
import sys
import zipfile

sys.path.insert(0, __import__("os").path.dirname(__file__))
from xl_engine import SST, Sheet, Styles, day_groups  # noqa: E402

HEADER, FIRST, LAST, COL = 8, 9, 131, "B"


def arg(flag):
    return sys.argv[sys.argv.index(flag) + 1] if flag in sys.argv else None


def filter_names(path):
    x = zipfile.ZipFile(path).read("xl/worksheets/sheet1.xml").decode("utf-8")
    m = re.search(r"<autoFilter [^>]*>(.*?)</autoFilter>", x, re.S)
    assert m, f"{path}: nav filtra"
    return [html.unescape(v) for v in re.findall(r'<filter val="([^"]*)"/>', m.group(1))]


def main():
    src, dst = sys.argv[1], sys.argv[2]
    keep = [k.strip() for k in (arg("--keep") or "").split(";") if k.strip()]
    zin = zipfile.ZipFile(src)
    rd = lambda p: zin.read(p).decode("utf-8")
    sheet, wb = rd("xl/worksheets/sheet1.xml"), rd("xl/workbook.xml")
    assert "<autoFilter" not in sheet, "filtrs jau ir"
    sh = Sheet(sheet, SST(rd("xl/sharedStrings.xml")), Styles(rd("xl/styles.xml"), rd("xl/theme/theme1.xml")))
    day_cols = [c for cols in day_groups(sh).values() for c in cols]
    names = {r: str(sh.value(r, 2) or "") for r in range(FIRST, LAST + 1)}

    if arg("--names-from"):
        show = set(filter_names(arg("--names-from")))
    else:
        show = {n for r, n in names.items() if n.strip() and any(sh.value(r, c) not in (None, "") for c in day_cols)}
        missing = [k for k in keep if k not in {n.strip() for n in names.values()}]
        assert not missing, f"nav tādu rindu: {missing}"
        show |= {n for n in names.values() if n.strip() in keep}
    hidden = [r for r, n in names.items() if n.strip() and n not in show]

    # rindas: hidden="1"
    def row_repl(m):
        r = int(m.group(1))
        if not FIRST <= r <= LAST:
            return m.group(0)  # 1.-7. palīgrindas u.c. paliek, kā bija (tās ir paslēptas oriģinālā)
        attrs = re.sub(r'\s*hidden="[^"]*"', "", m.group(2))
        return f'<row r="{r}"{attrs}{" hidden=" + chr(34) + "1" + chr(34) if r in hidden else ""}>'
    sheet, n_rows = re.subn(r'<row r="(\d+)"((?:[^>/]|/(?!>))*)>', row_repl, sheet)
    assert all(re.search(rf'<row r="{r}"[^>]* hidden="1"', sheet) for r in hidden)
    before = {r for r, a in re.findall(r'<row r="(\d+)"([^>]*)>', rd("xl/worksheets/sheet1.xml")) if 'hidden="1"' in a}
    after = {r for r, a in re.findall(r'<row r="(\d+)"([^>]*)>', sheet) if 'hidden="1"' in a}
    assert {int(r) for r in before} - {int(r) for r in after} <= set(range(FIRST, LAST + 1)), "atslēptas rindas ārpus filtra"

    # autoFilter uzreiz pēc </sheetData> (pirms mergeCells), filtra vērtības = redzamie vārdi
    ref = f"{COL}{HEADER}:{COL}{LAST}"
    vals = "".join(f'<filter val="{html.escape(n, quote=True)}"/>'
                   for n in sorted(show, key=lambda n: next(r for r, x in names.items() if x == n)))
    af = f'<autoFilter ref="{ref}"><filterColumn colId="0"><filters blank="1">{vals}</filters></filterColumn></autoFilter>'
    sheet = sheet.replace("</sheetData>", "</sheetData>" + af, 1)

    # _FilterDatabase vārds lapai (localSheetId = lapas indekss <sheets> sarakstā)
    sheets = re.findall(r'<sheet name="([^"]*)"[^>]*r:id="([^"]+)"', wb)
    rels = rd("xl/_rels/workbook.xml.rels")
    idx = next(i for i, (_, rid) in enumerate(sheets)
               if re.search(rf'Id="{rid}"[^>]*Target="(?:/xl/)?worksheets/sheet1.xml"', rels)
               or re.search(rf'Target="(?:/xl/)?worksheets/sheet1.xml"[^>]*Id="{rid}"', rels))
    sname = sheets[idx][0].replace("'", "''")
    dn = (f'<definedName name="_xlnm._FilterDatabase" localSheetId="{idx}" hidden="1">'
          f"'{sname}'!${COL}${HEADER}:${COL}${LAST}</definedName>")
    assert f'name="_xlnm._FilterDatabase" localSheetId="{idx}"' not in wb
    wb = wb.replace("<definedNames>", "<definedNames>" + dn, 1) if "<definedNames>" in wb else \
        wb.replace("</sheets>", f"</sheets><definedNames>{dn}</definedNames>", 1)

    parts = {"xl/worksheets/sheet1.xml": sheet, "xl/workbook.xml": wb}
    with zipfile.ZipFile(dst, "w") as zout:
        for info in zin.infolist():
            data = parts[info.filename].encode("utf-8") if info.filename in parts else zin.read(info.filename)
            zout.writestr(info, data, compress_type=zipfile.ZIP_DEFLATED)
    print(f"filtrs {ref}: redzamas {len(show)} personu rindas, paslēptas {len(hidden)}")
    for n in sorted(show, key=lambda n: next(r for r, x in names.items() if x == n)):
        print("   +", n)


if __name__ == "__main__":
    main()
