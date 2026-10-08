# xl_engine — darba stundu tabeles pārrēķins bez Excel

`xl_engine.py` ļauj labot mēneša darba stundu `.xlsm` failu (lapa `Darbs - <Mēnesis>`)
tieši XML līmenī un pārrēķināt visas formulu vērtības Pythonā — bez Excel un bez openpyxl.

Kāpēc: openpyxl saglabā formulas bez rezultātiem, tāpēc failu pēc tam obligāti jāpārrēķina
Excelī (`optimize_vba.ps1`). Šis dzinējs:

* maina tikai norādītās šūnas; VBA (`Module1`), pārējās lapas, stili un piezīmes paliek
  baitu līmenī neaizskarti;
* imitē `SumByHeaderColor` (salīdzina šūnas fona krāsu ar objekta galvenes krāsu, ieskaitot
  tēmas krāsas ar tonējumu — Windows HLS algoritms, HLSMAX = 240), `SUM`, `SUMIF`
  (arī "izstiepto" divu rindu kritēriju diapazonu), aritmētiku;
* ieraksta pārrēķinātās vērtības `<v>`, tāpēc fails atveras ar pareizām summām arī tad,
  ja makro ir bloķēti.

## Pārbaude pirms lietošanas

```python
from xl_engine import Styles, SST, Sheet, same
sh = Sheet(sheet1_xml, SST(shared_strings_xml), Styles(styles_xml, theme_xml))
vals = sh.recalc()
bad = [k for k, v in vals.items() if not same(sh.cached[k], v)]
assert not bad   # dzinējs sakrīt ar Excel saglabātajām vērtībām
```

Septembrī 2026 pārbaudīts pret 12 602 Excel saglabātajām formulu vērtībām — 0 atšķirību.

## Rediģēšana

* `sh.set_cell(r, c, value, style)` — konstante (skaitlis / teksts / `None`);
* `styles.xf_with_fill(style, fill_id)` — tas pats stils ar citu fona krāsu (objekta krāsa
  vai `0` statusiem `Atv.` / `SA` / `SB`);
* `sh.dump(sh.recalc())` — jaunais `sheet1.xml` + saraksts ar mainītajām formulu vērtībām.

Mēneša aizpildes skripti ar darbinieku datiem šeit **netiek glabāti** (repozitorijs ir publisks).
