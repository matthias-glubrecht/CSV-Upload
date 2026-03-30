# CSV Upload – SharePoint Framework Web Part

**Version 1.1.0**

Ein SPFx-Web-Part zum Importieren von CSV-Dateien in SharePoint-Listen.

## Funktionsumfang

- **Websitesammlung auswählen** — Type-Ahead-Suche über die SharePoint-Suchmaschine
- **Website auswählen** — Dropdown aller Unterwebsites der gewählten Websitesammlung
- **Liste auswählen** — Dropdown aller benutzerdefinierten Listen (keine Dokumentbibliotheken)
- **CSV-Datei hochladen** — Drag & Drop oder Dateiauswahl-Dialog
- **Automatische Zeichenkodierungs-Erkennung** — UTF-8 (mit/ohne BOM), UTF-16 LE und Windows-1252 (Excel-Standard) werden automatisch erkannt
- **Automatisches Feld-Mapping** — Abgleich der CSV-Spaltenüberschriften mit Anzeige- und internem Feldnamen
- **Schlüsselspalte** — Optional kann eine Spalte als Schlüssel markiert werden, um bestehende Einträge zu aktualisieren (Upsert)
- **Defaultwerte** — Für Pflichtfelder (die nicht Schlüssel sind) können Defaultwerte angegeben werden
- **Fortschrittsanzeige** — Zeigt den Importfortschritt mit Zählern für erstellt/aktualisiert/Fehler
- **Interaktiver Fehler-Dialog** — Bei nicht auflösbaren Werten (Benutzer, Nachschlagen, Auswahl, Taxonomie) wird ein Dialog angezeigt, in dem der Benutzer den Wert korrigieren, das Feld überspringen oder die gesamte Zeile überspringen kann
- **Benutzerfreundliche Fehlermeldungen** — SharePoint-REST-Fehler werden in verständliche Meldungen übersetzt

### Unterstützte Feldtypen

| Feldtyp | Behandlung |
|---|---|
| Text, Mehrzeiliger Text | Direktübernahme |
| Zahl, Währung | Komma → Punkt Konvertierung |
| Datum/Uhrzeit | Deutsche (`DD.MM.YYYY`), ISO- und US-Formate; `[Heute]`/`[Today]` als Defaultwert |
| Ja/Nein | Erkennt `true`, `1`, `ja`, `yes`, `wahr` |
| Auswahl / Mehrfachauswahl | Optional mit freier Eingabe; ungültige Werte werden mit Fehlermeldung zurückgewiesen |
| Nachschlagen / Nachschlagen (mehrere) | Automatische Auflösung des Anzeigewerts zur Item-ID; nicht auflösbare Werte werden gemeldet |
| Person / Person (mehrere) | Auflösung über `ensureUser` oder Anzeigename; nicht auflösbare Benutzer werden gemeldet |
| Hyperlink | Format: `URL, Beschreibung`; URLs ohne Protokoll erhalten automatisch `https://` |
| Verwaltete Metadaten | Auflösung über TaxonomyHiddenList mit Fallback auf den Term Store via `@pnp/sp-taxonomy` |
| Verwaltete Metadaten (mehrere) | Schreiben über das versteckte Note-Feld; mehrteilige Werte durch `;` getrennt |

### SharePoint-Export-Kompatibilität

Der Web Part erkennt und verarbeitet automatisch das SharePoint-Exportformat für Taxonomie-Felder:
- Einzel: `6;#Einsatz` → `Einsatz`
- Mehrfach: `12;#Lehre;#13;#Forschung` → `Lehre`, `Forschung`
- Sonderfall: `22;##IT-Verfahren` → `IT-Verfahren` (Extra-`#` wird entfernt)

## Technologie-Stack

| Aspekt | Detail |
|---|---|
| SPFx-Version | 1.4.1 |
| React | 15.6.2 (Klassenkomponenten) |
| UI-Bibliothek | office-ui-fabric-react v5 |
| Datenzugriff | @pnp/sp ^1.3.11, @pnp/sp-taxonomy ^1.3.11 |
| TypeScript-Ziel | ES5 |
| Build-Tool | Gulp 3.9 |
| Lokalisierung | Deutsch (de-de), Englisch (en-us) |

## Projektstruktur

```
src/webparts/uploadCsv/
├── UploadCsvWebPart.ts              # WebPart-Klasse (Einstiegspunkt)
├── components/
│   ├── UploadCsv.tsx                # Hauptkomponente (Orchestrierung)
│   ├── UploadCsv.module.scss        # Gemeinsame Styles (CSS Modules)
│   ├── SiteCollectionPicker/        # Websitesammlungs-Auswahl (ComboBox)
│   ├── WebPicker/                   # Website-Auswahl (Dropdown)
│   ├── ListPicker/                  # Listen-Auswahl (Dropdown)
│   ├── CsvDropZone/                 # Drag & Drop / Dateiauswahl
│   ├── MappingTable/                # Feldzuordnungstabelle
│   ├── ImportProgress/              # Fortschrittsanzeige
│   └── FieldErrorDialog/            # Interaktiver Fehler-Dialog für Feldwerte
├── service/
│   └── CsvUploadService.ts          # Datenzugriff via @pnp/sp + @pnp/sp-taxonomy
├── models/
│   └── IModels.ts                   # TypeScript-Interfaces
├── utils/
│   └── CsvParser.ts                 # CSV-Parser mit Delimiter- und Encoding-Erkennung
└── loc/
    ├── mystrings.d.ts               # String-Typdefinitionen
    ├── en-us.js                     # Englische Übersetzungen
    └── de-de.js                     # Deutsche Übersetzungen
```

## Entwicklung

### Voraussetzungen

- Node.js 8.x
- Gulp CLI (`npm install -g gulp-cli`)

### Einrichtung

```bash
npm install
```

### Lokaler Build

```bash
gulp bundle
```

### Lokales Testen

```bash
gulp serve
```

Dann die SharePoint Workbench öffnen:
`https://<tenant>.sharepoint.com/_layouts/15/workbench.aspx`

### Produktion

```bash
gulp bundle --ship
gulp package-solution --ship
```

Die fertige Lösung liegt unter `sharepoint/solution/csv-upload.sppkg`.

### Bereitstellung

1. Die `.sppkg`-Datei in den **App-Katalog** der SharePoint-Umgebung hochladen
2. Im App-Katalog auf **Bereitstellen** klicken
3. Die App **CSV Upload** zur gewünschten Site hinzufügen
4. Den Web Part **Upload CSV** auf einer Seite einfügen

## Versionierung

Die Versionsnummer wird synchron in zwei Dateien gepflegt:

- `package.json` → `a.b.c`
- `config/package-solution.json` → `a.b.c.0`

## Lizenz

Intern / Proprietär
