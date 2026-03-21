# Excel Import Verifier
![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![Status](https://img.shields.io/badge/status-POC-orange)
![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)
[![English](https://img.shields.io/badge/README-English-informational?style=flat-square)](README_en.md)
[![Deutsch](https://img.shields.io/badge/README-Deutsch-informational?style=flat-square)](README.md)

## Problem

Nach einem Import in HR-Systeme gibt es oft keine sauberen Logs die zeigen ob alle Daten korrekt übernommen wurden. Der manuelle Abgleich von Import- und Exportdatei ist zeitaufwendig und fehleranfällig — bei 500 Mitarbeitern kann das 40+ Minuten dauern.

## Lösung

Dieses Tool vergleicht Import- und Exportdatei automatisch und markiert Abweichungen direkt in der Importdatei. Ergebnis in 2 Minuten statt 40.

## Demo

![Excel Import Verifier Demo](demo.gif)

## Funktionen

- GUI zur Dateiauswahl — kein Code-Kontakt nötig
- Vergleicht zwei Excel-Dateien (Import vs. Export)
- Mitarbeiter-Matching anhand Personalnummer, Vorname, Nachname (mindestens 2 von 3)
- Markierungen direkt in der Importdatei:
  - **Rot** — kein Treffer oder Abweichung in einzelnen Feldern
  - **Orange** — mehrere mögliche Treffer gefunden
- Feldspezifische Vergleichslogik:
  - Namen (ignoriert Groß-/Kleinschreibung, Leerzeichen, Bindestriche)
  - E-Mail (case-insensitive)
  - Datumsfelder (normalisiert auf `TT.MM.JJJJ`)
  - HR-Felder wie Familienstand, Elterneigenschaft
- Speichert Ergebnis als neue Excel-Datei

## Nutzung

1. Abhängigkeiten installieren:
```bash
pip install openpyxl
```
2. Tool starten:
```bash
python imexportchk.py
```
3. Import-Datei, Export-Datei und Ausgabepfad auswählen
4. Vergleich starten

## Testdaten

Beispieldateien findest du im Ordner `example_data/`.

## Status

Entwickelt und getestet mit anonymisierten HR-Daten.
Nicht gegen alle HR-System-Exportformate getestet.

## Technologien

- Python 3.x
- openpyxl
- tkinter

## Lizenz

MIT License — siehe [LICENSE](LICENSE)
