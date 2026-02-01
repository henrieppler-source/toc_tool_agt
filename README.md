# Inhaltsverzeichnis-Generator (PDF-Lesezeichen → PDF + Excel)

## Zweck
Dieses Tool erzeugt aus den **Lesezeichen (Outlines)** einer PDF-Datei:
- ein **Inhaltsverzeichnis als PDF** (Layout an Muster angelehnt)
- ein **Inhaltsverzeichnis als Excel** auf Basis des **Excel-Musters** (Format bleibt erhalten)
- zusätzlich eine **Master-Excel** (`MASTER_Inhaltsverzeichnis.xlsx`), die bei jedem Lauf **fortgeschrieben** wird.

## Eingabe
- Ein Ordner mit PDFs, die Lesezeichen enthalten.

## Ausgabe
Im Ausgabeordner entstehen pro PDF:
- `<PDFNAME>_Inhaltsverzeichnis.pdf`
- `<PDFNAME>_Inhaltsverzeichnis.xlsx`

Zusätzlich:
- `MASTER_Inhaltsverzeichnis.xlsx` (eine Tabelle mit allen Einträgen)

## Templates
Im Programmordner müssen liegen:
- `Muster-Inhaltsverzeichnis-EXCEL.xlsx`
- (optional) `Muster-Inhaltsverzeichnis-PDF.pdf` (nur Referenz)

## Nutzung (Python)
```bash
pip install pypdf reportlab openpyxl
python main.py
```

## EXE bauen (PyInstaller)
```bash
pip install pyinstaller
pyinstaller --onefile --noconsole --name "TOC_Generator" main.py
```

## Hinweise / Grenzen
- Wenn ein PDF **keine Lesezeichen** hat, wird es übersprungen.
- Master-Excel kann Duplikate verhindern (Checkbox).
- Hierarchie wird in Excel per Einrückung (Spaces) sichtbar gemacht.
