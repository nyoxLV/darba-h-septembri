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
* `sh.dump(sh.recalc())` — jaunais `sheet1.xml` + saraksts ar mainītajām formulu vērtībām
  (arī Excel kļūdas, piem. `#DIV/0!` tukšā rindā, tiek ierakstītas kā `t="e"`);
* `Notes(comments_xml, vml)` — piezīmju pievienošana / dzēšana (`add`, `drop`, `drop_many`);
* `day_groups(sh)` — dienu grupas pēc 8. rindas sapludinātajām šūnām (1..31).

Ar šiem rīkiem var arī sagatavot jauna mēneša failu no iepriekšējā: pārsaukt lapu, iztīrīt dienu
režģi un piezīmes, pārkrāsot brīvdienas pēc jaunā kalendāra.

## Kolonnu ievietošana

`insert_columns(parts, [(pos, k), ...])` ievieto `k` kolonnas pirms kolonnas `pos` lapā "Darbs":
pārbīda šūnas, formulas (koplietotās pārvērš parastās), sapludinājumus, kolonnu platumus,
nosacījuma formatējumu, piezīmes un VML, izmēru, atlasi, lapas pārtraukumus; 8. rindas dienas
sapludinājums tiek pagarināts; `calcChain.xml` tiek dzēsts (Excel to izveido no jauna).

`insert_day_cols.py <ieeja.xlsm> <izeja.xlsm> 4 7 15 16 16` — +1 kolonna norādītajām dienām (diena
atkārtota N reizes → +N kolonnas). Skripts pats pārbauda, ka visas formulu vērtības un konstantes pēc
pārbīdes ir nemainīgas.

## Personu filtrs

`add_person_filter.py <ieeja.xlsm> <izeja.xlsm> [--keep "Vārds Uzvārds;..."] [--names-from <fails.xlsm>]`
ieliek Excel automātisko filtru uz vārdu kolonnas (`B8:B131`). Redzamas paliek rindas, kurās mēneša
dienu režģī ir stundas vai statuss, plus `--keep` vārdi; pārējās rindas tiek paslēptas (ne dzēstas),
tāpēc formulas un kopsummas nemainās. Excelī: B8 bultiņa → ieķeksē / atķeksē personas.
Palīgrindas ārpus 9.–131. rindas paliek, kā bija. `--names-from` ņem to pašu personu sarakstu no
cita faila (piem. nākamajam mēnesim).

## Audits

`audit_all.py <fails.xlsm> [gads mēnesis] [--alias aliases.json]` — tikai lasa un izdrukā: stundas
bez objekta krāsas, piezīmes objekts ≠ šūnas krāsas objekts, stundas / statusi brīvdienās, diena
> 12 h, statuss + stundas vienā dienā, "Stundu pārbaude" ≠ 0, EDLUS spoguļrinda ≠ pamatrinda, un
katra cilvēka mēneša stundas. `aliases.json` = `{"Objekta nosaukums": ["atslēgvārds", ...]}`
(neobligāts; objektu atpazīšanai piezīmēs).

Mēneša aizpildes skripti ar darbinieku datiem šeit **netiek glabāti** (repozitorijs ir publisks).
