# Aktenverzeichnis — Kanzlei-Tool

Ein webbasiertes Tool für Kanzleien zur digitalen Aktenführung. Dokumente per Drag & Drop einwerfen, automatisch analysieren lassen und als strukturiertes Aktenverzeichnis mit interaktiver Tabelle verwalten.

## Was es kann

- **Drag & Drop** von PDF, Word, Excel und Bildern (Screenshots, Fotos) — einzeln oder mehrere gleichzeitig
- **KI-Analyse** erkennt automatisch: Datum, Dokumenttyp (Klageschrift, Vollmacht, Bescheid, …), Essentialia und beteiligte Personen
- **Drei Tabs**: Hauptakte, Chronologie (datumssortiert, per Knopfdruck generiert), Personenübersicht
- **Inline-Editing**: Alle Zellen direkt in der Tabelle bearbeitbar
- **Nachträgliches Hinzufügen**: Neue Dokumente jederzeit reinwerfen — die Tabelle wird ergänzt
- **Alte Tabelle importieren**: Bestehende XLSX-Aktenverzeichnisse einlesen und weiterführen
- **Excel-Export**: Formatierte XLSX mit 4 Blättern (Vorblatt, Aktenübersicht, Chronologie, Personenübersicht)
- **Wahlweise OpenAI (GPT-4o) oder Anthropic (Claude)** mit eigenem API-Key

## Funktionaler Ablauf

1. **Akte anlegen** — Aktenzeichen und Bezeichnung eingeben (z.B. „Az. 123/26 Müller ./. Schmidt")
2. **Dokumente einwerfen** — Per Drag & Drop oder Datei-Auswahl. Jedes Dokument wird automatisch analysiert und eingetragen
3. **Tabelle pflegen** — Alle Zellen inline editierbar. Kategorien per Dropdown. Blatt-Nummern, Essentialia, Anmerkungen ergänzen
4. **Chronologie generieren** — Per Knopfdruck aus den Hauptakte-Einträgen, datumssortiert
5. **Weitere Dokumente** — Jederzeit neue Dokumente per Drag & Drop hinzufügen
6. **Excel exportieren** — Formatierte XLSX mit allen 4 Tabellenblättern

## Installation

Keine Installation nötig — es ist eine statische Web-App.

### Lokal starten

```bash
git clone https://github.com/Klotzkette/aktenverzeichnis.git
cd aktenverzeichnis


# Beliebigen lokalen Server starten:
npx http-server ./app -p 8080 -c-1
# oder
cd app && python -m http.server 8080
```

Dann `http://localhost:8080` im Browser öffnen.

### API-Schlüssel einrichten

1. Zahnrad-Icon oben rechts klicken
2. KI-Anbieter wählen (OpenAI oder Anthropic)
3. API-Schlüssel eingeben und speichern

## Tabellenstruktur

### Hauptakte
| # | Blatt | Datum | Vorgang | Kategorie | Essentialia | Anm. Kanzlei | Anm. Mandant |

### Chronologie
| # | Datum | Blatt | Beteiligte | Vorgang | Anm. Kanzlei | Anm. Mandant |

### Personenübersicht
| # | Name | Adresse | Prozessrolle | Blatt | Anm. Kanzlei | Anm. Mandant |

## Erkannte Kategorien

Klageschrift, Klageerwiderung, Schriftsatz, Bescheid, Widerspruchsbescheid, Vollmacht, Gutachten, Rechnung, Vertrag, Beschluss, Urteil, Protokoll, Korrespondenz, Mahnung, Abrechnung, Anlage, Sonstige

## Technologie

- Vanilla HTML/CSS/JS — keine Build-Tools, keine Frameworks
- [SheetJS](https://sheetjs.com/) — XLSX Import/Export
- [pdf.js](https://mozilla.github.io/pdf.js/) — PDF-Textextraktion
- [mammoth.js](https://github.com/mwilliamson/mammoth.js) — Word-Dokumentverarbeitung
- OpenAI / Anthropic Vision API — Bildanalyse für Screenshots

## Design

Adaptiert vom PDF-Pseudonymizer: Soft blue-teal Palette, Swiss-style Minimalismus, serifenlose Typografie.

## Urheberrecht / Copyright

© 2026 Tom Braegelmann. Alle Rechte vorbehalten.

Diese Software und alle zugehörigen Dateien sind urheberrechtlich geschützt. Es wird keinerlei Lizenz — weder ausdrücklich noch stillschweigend — gewährt. Jede Nutzung, Vervielfältigung, Verbreitung oder Bearbeitung bedarf der ausdrücklichen schriftlichen Genehmigung des Urhebers.

© 2026 Tom Braegelmann. All rights reserved.


## Co-Autor

Dieses Projekt wurde gemeinsam mit [Claude](https://www.anthropic.com/claude) (Anthropic) als KI-Assistenten entwickelt.
This software and all associated files are protected by copyright. No license — express or implied — is granted. Any use, reproduction, distribution, or modification requires the explicit written permission of the copyright holder.
