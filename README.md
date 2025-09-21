# Excel Import-Export Vergleichstool (POC)
[English](README_en.md)

## Überblick
Dieses Repository enthält ein **Proof of Concept** für ein Python-Tool, das eine Importdatei mit dem zugehörigen Exportergebnis in Excel vergleicht.  
Ziel ist es, zu prüfen, ob ein Import erfolgreich war, und Abweichungen automatisch zu markieren – so entfällt die fehleranfällige manuelle Kontrolle.

## Funktionen
- Benutzerfreundliche **GUI mit Tkinter** zur Dateiauswahl  
- Vergleicht zwei Excel-Dateien (Import vs. Export)  
- Abgleich von Mitarbeitern anhand **Personalnummer, Vorname, Nachname** (mindestens 2 von 3 müssen übereinstimmen)  
- Markiert Abweichungen direkt in der Import-Datei:
  - **Rot**: kein Treffer oder Unterschiede in den Feldern  
  - **Orange**: mehrere mögliche Treffer gefunden  
- Feldspezifische Vergleichslogik:
  - Namen (case-insensitive, ignoriert Leerzeichen und Bindestriche)  
  - E-Mail (case-insensitive)  
  - Datumsfelder (normalisiert auf `TT.MM.JJJJ`)  
  - Bestimmte HR-Felder (z. B. Familienstand, Elterneigenschaft)  
- Speichert das Ergebnis als neue Excel-Datei mit markierten Abweichungen  

## Motivation
In einem Projekt war der Erfolg von Importen unsicher, da das System keine sauberen Logs erzeugte.  
Das manuelle Vergleichen von Import- und Exportdaten war zeitaufwendig und fehleranfällig.  
Dieses Tool automatisiert den Prozess und liefert ein klares, prüfbares Ergebnis.

## Nutzung
1. Tool starten (`imexportchk.py` oder via kompilierte `.exe`).  
2. Folgende Dateien auswählen:
   - Import-Datei  
   - Export-Datei  
   - Name der Ergebnisdatei  
3. Vergleich starten.  
4. Das Tool speichert eine neue Excel-Datei, in der Abweichungen direkt im Import markiert sind.

## Beispiel
- Eingabe:  
  - Import-Datei mit Mitarbeiterstammdaten  
  - Export-Datei nach Systemimport  
- Ausgabe:  
  - Kopie der Import-Datei mit hervorgehobenen Unterschieden  

## Status
- Proof of Concept (POC)  
- Nicht für den produktiven Einsatz gedacht  
- Getestet mit HR-/Payroll-ähnlichen Excel-Strukturen (bis ~500 Zeilen)  

## Technologien
- Python  
- OpenPyXL  
- Tkinter  

## Hinweis
Dies ist ein **POC-Projekt** und nicht offiziell freigegeben. Nutzung erfolgt auf eigene Verantwortung.

## Lizenz
Dieses Projekt ist unter der MIT-Lizenz veröffentlicht.  
Siehe die Datei [LICENSE](LICENSE) für Details.
