# Sheets – Excel & CSV Editor

Excel- und CSV-Dateien direkt in Visual Studio Code öffnen und bearbeiten. Sheets bringt einen Tabelleneditor mit Formeln, Formatierung, Filtern und Diagrammen mit. Microsoft Excel oder ein externer Dienst sind dafür nicht erforderlich.

![Sheets Tabelleneditor mit Tabellenformatierung und Diagramm](media/preview.png)

## Installation

Die Extension lässt sich als VSIX aus diesem Repository bauen:

```bash
npm ci
npm run package
```

Danach in VS Code **Extensions: Install from VSIX…** ausführen und `artifacts/sheets-vscode-0.2.3.vsix` auswählen. Alternativ:

```bash
code --install-extension artifacts/sheets-vscode-0.2.3.vsix
```

Voraussetzungen: VS Code 1.95 oder neuer; zum Entwickeln und Paketieren Node.js 22 oder neuer. Das Paket enthält alle Laufzeitabhängigkeiten.

## Erste Schritte

- Eine `.xlsx`-Datei öffnen: Sie erscheint direkt im Sheets-Editor.
- Eine `.csv`- oder `.tsv`-Datei im Explorer rechts anklicken und **Mit Sheets öffnen** wählen. Alternativ **Öffnen mit… → Sheets (CSV)** verwenden.
- Über die Befehlspalette **Sheets: Tabellendatei öffnen…** eine Datei auswählen oder mit **Sheets: Neue Arbeitsmappe** eine neue XLSX-Datei anlegen.
- Mit `Ctrl+S` / `Cmd+S` speichern. Rückgängig und Wiederholen sind mit VS Code verbunden.

Die Bezeichnungen folgen der Sprache von VS Code. Englisch und Deutsch werden unterstützt.

Falls XLSX weiterhin in einem anderen Editor öffnet: Nach der Installation **Developer: Reload Window** ausführen und die Datei schließen und erneut öffnen. Eine bereits konfigurierte Benutzer- oder Workspace-Zuordnung hat Vorrang vor dem Extension-Standard. In diesem Fall über **Öffnen mit… → Standardeditor für „*.xlsx“ konfigurieren… → Sheets** die Zuordnung ändern. In `settings.json` entspricht das:

```json
"workbench.editorAssociations": {
  "*.xlsx": "sheets.spreadsheet"
}
```

## Funktionen

- Zellen bearbeiten, kopieren, einfügen und automatisch ausfüllen; Zeilen, Spalten und Blätter verwalten.
- Zahlen-, Währungs-, Prozent-, Datums- und eigene Zellformate; Schrift, Farben, Rahmen und Ausrichtung. Zellformatierung über `Ctrl+1` / `Cmd+1`.
- Tabellen hervorheben, komplette Datensätze sortieren und über die Kopfzeilen nach Werten filtern. Wertefilter bieten Suche, Mehrfachauswahl und leere Werte.
- Formeln mit über 130 Funktionen, darunter `XLOOKUP`, `PMT`, `FILTER`, `SORT`, `UNIQUE` und `SEQUENCE`; Vorschläge und farbige Zellreferenzen während der Eingabe.
- Bedingte Formatierung mit Zellvergleichen, Formelregeln und Farbskalen.
- Säulen-, Balken-, Linien-, Flächen-, Kreis-, Ring- und Punktdiagramme erstellen und bearbeiten.
- Pivot-Auswertungen mit Gruppen, Aggregaten und Gesamtsummen erstellen und aktualisieren.
- Excel-Tabellen, Zellformate, Diagramme und Wertefilter in XLSX speichern. CSV-Trennzeichen bleiben beim Speichern erhalten.
- Helles und dunkles VS-Code-Theme sowie fixierte Zeilen und Spalten.

## KI-Agents

Im Copilot-Chat **Agent** wählen und unter **Configure Tools / Werkzeuge konfigurieren** die Sheets-Werkzeuge aktivieren. Für Agent Mode eine aktuelle VS-Code-Version verwenden (Extension-Werkzeuge werden dort seit VS Code 1.99 unterstützt).

Sheets registriert 41 Werkzeuge (`sheets_*`) für Clients, die die VS-Code-Language-Model-Tools unterstützen, beispielsweise Copilot im Agent-Modus. Sie können Bereiche lesen und schreiben, Formeln und Formate anwenden, Tabellen und Diagramme erstellen oder sortieren.

Änderungen an geöffneten Dateien erscheinen im Editor und lassen sich rückgängig machen. Die Einstellung `sheets.agentTools.headlessWrites` erlaubt zusätzlich Änderungen an ungeöffneten Dateien, die direkt gespeichert werden. Diese Funktion erfordert einen vertrauenswürdigen Arbeitsbereich. Für die manuelle Tabellenbearbeitung wird kein KI-Konto benötigt.

Andere Extensions können `sheets.executeTool` aufrufen:

```ts
await vscode.commands.executeCommand('sheets.executeTool', {
  name: 'read_range',
  file: '/path/to/workbook.xlsx',
  input: { range: 'A1:D10' },
});
```

Die Tabellenverarbeitung läuft lokal. Die Extension sendet selbst keine Tabellendaten an externe Dienste; ein von dir eingesetzter KI-Client kann Werkzeugergebnisse an seinen Anbieter übertragen.

## Unterstützte Formate und Grenzen

Unterstützt werden `.xlsx`, `.csv` und `.tsv`. Makros und das alte `.xls`-Format werden nicht unterstützt. CSV/TSV können nur Werte eines Blatts speichern, keine Formatierung oder Diagramme.

Nicht alle Excel-Funktionen werden im Editor dargestellt oder bearbeitet. Dazu gehören native Pivot-Caches, strukturierte Tabellenreferenzen, Datenüberprüfung, Bilder und Kommentare. Eingebaute Excel-Tabellenstile werden im Editor teilweise näherungsweise blau dargestellt; der ursprüngliche Stil bleibt in der Datei erhalten.

Beim Speichern bleibt das Originalpaket die Grundlage. Nicht implementierte Bestandteile werden mitgeführt. Ohne Änderung wird die Datei bytegenau zurückgegeben. Änderungen, die sich nicht sicher übertragen lassen – etwa Struktur- oder Sortieränderungen mit unbekannten Excel-Inhalten –, werden vor dem Überschreiben mit einer Fehlermeldung blockiert. Wiederherstellungssicherungen behalten auch solche ungespeicherten Änderungen. Details: [XLSX-Kompatibilität und Speicherschutz](docs/xlsx-compatibility.md).

## Entwicklung

```bash
npm ci
npm run build        # Typprüfung und beide Extension-Bundles
npm test             # Core-, UI-, Agent- und Extension-Tests
npm run package      # Build, VSIX und Prüfung des fertigen Pakets
npm run check:package
```

Das Repository in VS Code öffnen und **F5** drücken. Die Konfiguration „Sheets: Extension starten“ baut die Extension im Watch-Modus und öffnet einen Extension Development Host. Nach Änderungen am Host-Code dessen Fenster neu laden; nach Änderungen am Webview den Tabelleneditor erneut öffnen.

Die Web-Demo bleibt über `npm run dev:web` verfügbar. `npm run build:all` baut zusätzlich Web-Demo und MCP-Server.

## Lizenz

MIT. Siehe LICENSE.
