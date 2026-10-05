# DART-Ziele Antibiotika-Resistenz – Folienbericht

Erzeugt aus einer Excel-Tabelle die Folie „Wo stehen wir heute?“ (vier Diagramme: *E. coli* Cephalosporine 3. Gen., MRSA, *K. pneumoniae* Carbapeneme, VRE) als PDF.

## Webseite (GitHub Pages)

Die Seite liegt in `docs/`. Die Excel-Datei wird nur im Browser verarbeitet, es werden keine Daten übertragen.

Aktivieren: *Settings → Pages → Build and deployment → Source: Deploy from a branch*, Branch auswählen, Ordner `/docs`.

## Excel-Format

Erstes Tabellenblatt, erste Zeile ist die Kopfzeile:

| drug_bug | 2019 | 2020 | … | 2026 | target_perc | target_abs |
|---|---|---|---|---|---|---|
| 3GCECOLI | 12,02 | 10,74 | … | 12,5 | -0,12 | 10,58 |
| CRKP | … | | | | | |
| MRSA | … | | | | | |
| VRE | … | | | | | |

- Jahresspalten werden an der Kopfzeile (vierstellige Jahreszahl) erkannt; leere Jahre am Ende werden ignoriert.
- Der Zielwert 2030 wird aus der Spalte `target_abs` (bzw. „Target 2030 …“) gelesen, sonst aus der letzten Spalte.

## Python-Variante

```
pip install openpyxl matplotlib
python make_report.py data/input_updated.xlsx output/bericht_neu.pdf 05.10.2026
```

Liegt `data/logo.png` lokal vor, wird es oben rechts eingefügt (nicht im Repository enthalten).
