# Sheets

TypeScript-Framework, um Excel- (`.xlsx`) und CSV-Dateien im Browser und in VS Code wie in Excel zu bearbeiten: Daten ändern, Formeln eingeben, Diagramme erstellen. Eine Tool-API mit JSON-Schema macht dieselben Funktionen für AI-Agents nutzbar.

```
packages/
  core/    Modell, Formel-Engine, Neuberechnung, Undo/Redo, xlsx/CSV-I/O   (keine DOM-Abhängigkeit)
  ui/      SpreadsheetView: Canvas-Grid, Editor, Toolbar, Diagramme        (DOM)
  agent/   Agent-Tools (JSON-Schema), Adapter für Claude/OpenAI/MCP, postMessage-Bridge
  mcp/     MCP-Server (stdio) zum Bearbeiten von Dateien durch Agents
apps/
  web/     Web-App (Vite)
src/       VS-Code-Extension (Custom Editor + Language-Model-Tools)
```

Abhängigkeiten zur Laufzeit: nur `fflate` (ZIP) und `fast-xml-parser` im Core.

## Schnellstart

```bash
npm install
npm run dev:web      # Web-App auf http://localhost:5173
npm test             # Vitest
npm run typecheck    # tsc -b über alle Pakete
npm run build:all    # Pakete, Web-App, MCP-Server und VS-Code-Extension
```

Workspace-Pakete werden über die Export-Bedingung `source` direkt aus den TypeScript-Quellen eingebunden (Vite, esbuild, Vitest). Für die Veröffentlichung auf npm erzeugt `tsc -b` `dist/`.

## Architektur

**Alle Änderungen sind Operationen.** Jede Mutation ist eine JSON-serialisierbare `Operation` (`setCells`, `insertRows`, `setCharts`, …), die `Workbook.apply()` ausführt. Daraus ergeben sich:

- Undo/Redo: Jede Operation liefert ihre Umkehrung, Transaktionen bündeln sie zu einem Schritt.
- ein Change-Event pro Transaktion mit `origin` (`user`, `agent`, `undo`, …)
- eine Grundlage für Protokollierung, Kollaboration und Fernsteuerung

High-Level-Befehle in `commands.ts` (Eingabe, Formatieren, Kopieren/Einfügen, Ausfüllen, Sortieren, Diagramme, Suche) bauen auf diesen Operationen auf. UI und Agent verwenden dieselben Befehle.

**Formel-Engine:** Tokenizer → Pratt-Parser (Excel-Präzedenz) → Evaluator mit Array-Broadcasting. Über 200 Funktionen (Mathe, Statistik, Finanzmathematik, Logik, Lookup inkl. `XLOOKUP`, Text, Datum) sowie `LET`/`LAMBDA` mit `MAP`, `BYROW`, `BYCOL`, `REDUCE`, `SCAN` und `MAKEARRAY`. Strukturierte Tabellenverweise (`Tabelle1[Spalte]`, `[@Spalte]`, `Tabelle1[#Alle]`) werden gegen die Tabellen der Mappe aufgelöst und in XLSX in der von Excel erwarteten Form gespeichert. Ein Abhängigkeitsgraph sorgt dafür, dass nach einer Änderung die betroffenen Zellen neu berechnet werden. Mappen mit dynamischen Arrays werden vollständig neu berechnet, damit auch veränderte Überlaufbereiche und deren Abhängigkeiten berücksichtigt werden. Zirkelbezüge ergeben `#CIRC!`. Bei Formeln mit unbekannten Funktionen bleibt der in der Datei gespeicherte Wert von Excel erhalten.

**xlsx:** eigener Reader und Writer, kein SheetJS/ExcelJS. Unterstützt werden Werte, Formeln (inkl. Shared Formulas, `_xlfn.`-Präfixe), Styles (Schrift, Farben inkl. Theme/Tint, Füllung, Rahmen, Ausrichtung, Zahlenformate), Spaltenbreiten und Zeilenhöhen, verbundene Zellen, fixierte Bereiche, definierte Namen und **native Excel-Diagramme** (Säule, Balken, Linie, Fläche, Kreis, Ring, Punkt), lesend und schreibend. Mit openpyxl in beide Richtungen gegengeprüft.

**Finanzfunktionen:** `PMT`, `PV`, `FV`, `NPER`, `NPV`, `IPMT` und `PPMT` berechnen Raten, Barwerte, Endwerte, Laufzeiten und Zins-/Tilgungsanteile. Verfügbar in Formeln, der Autovervollständigung und über `list_functions` mit `category: 'financial'`. Zinssatz und Laufzeit müssen dieselbe Periodeneinheit verwenden; `type: 0` steht für Zahlungen am Periodenende, `type: 1` am Anfang. Auszahlungen werden negativ angegeben, Einnahmen positiv. Beispiel: `=PMT(6%/12,36,10000)` berechnet die Monatsrate eines Kredits.

`NPV` diskontiert die erste Zahlung um eine Periode; eine sofortige Anfangsinvestition wird separat addiert, z. B. `=NPV(10%,3000,4200,6800)-10000`. Nichtnumerische Zahlungswerte einschließlich Fehlerwerten werden übersprungen. Syntax und Referenzfälle richten sich nach der Microsoft-Dokumentation zu [PMT](https://support.microsoft.com/en-us/excel/functions/pmt-function), [IPMT](https://support.microsoft.com/en-us/excel/functions/ipmt-function) und [NPV](https://support.microsoft.com/en-us/excel/functions/npv-function).

## Verwendung

### Core (Browser oder Node)

```ts
import { Workbook, setValues, readXlsx, writeXlsx, insertChartFromRange, toRange } from '@sheets/core';

const wb = new Workbook();
setValues(wb, 0, 0, 0, [
  ['Monat', 'Umsatz'],
  ['Jan', 1200],
  ['Feb', 1500],
  ['Summe', '=SUM(B2:B3)'],
]);
wb.getValue(0, 3, 1);                  // 2700
insertChartFromRange(wb, 0, toRange('A1:B3'), 'column', { title: 'Umsatz' });
wb.undo();
const bytes = writeXlsx(wb);           // Uint8Array
```

Eigene Funktionen:

```ts
import { registerFunction } from '@sheets/core';
registerFunction({
  name: 'VAT', category: 'math', syntax: 'VAT(amount)', description: 'Adds 19% VAT.',
  minArgs: 1, maxArgs: 1,
  impl: ([a], ctx) => (typeof a === 'number' ? a * 1.19 : 0),
});
```

### UI

```ts
import { SpreadsheetView } from '@sheets/ui';
const view = new SpreadsheetView(document.getElementById('app')!, wb, {
  language: 'de',
  toolbarActions: [{ id: 'save', label: 'Speichern', onClick: () => save() }],
});
view.select('B2:B10');
view.events.on('selectionChange', (s) => console.log(s));
```

Bedienung wie in Excel:

- Tastatur: Pfeiltasten, Strg+Pfeil, Shift-Auswahl, F2, Enter/Tab, Strg+C/X/V (auch mit Excel und Google Sheets), Strg+Z/Y, Strg+B/I/U, Strg+D/R
- Ausfüllkästchen mit Erkennung von Reihen
- Zellreferenzen per Klick oder Pfeiltaste in Formeln einfügen, farbige Markierung der Referenzen, Autovervollständigung von Funktionen und Syntax-Hinweise
- Spaltenbreite per Ziehen ändern, Doppelklick passt sie automatisch an
- Kontextmenü, Blatt-Tabs, Summe/Mittelwert/Anzahl in der Statusleiste
- Diagramme verschieben, skalieren und per Doppelklick bearbeiten
- helles und dunkles Theme; in VS Code wird das Editor-Theme übernommen
- bedingte Formatierung über den Regeleditor (Zellvergleiche, Formelregeln, Farbskalen und Prioritäten)
- Pivot-Auswertungen aus der Auswahl erstellen, aktualisieren und entfernen
- „Als Tabelle formatieren…“ mit fünf Farbstilen, Kopfzeile, wechselnden Zeilenfarben und Sortier- und Wertefiltermenüs direkt in den Tabellenköpfen
- „Zellen formatieren…“ (Strg/Cmd+1, Kontextmenü oder „…“ im Zahlenformat-Menü) mit Formatvorschau, eigenen Zahlenformaten, Dezimalstellen, Schrift, Farben, Ausrichtung und optionaler Anpassung der Spaltenbreiten
- „Sortieren…“ mit mehreren Sortierspalten und ausdrücklicher Kopfzeilenoption; bei einzelnen Spalten wird der zusammengehörige Datenbereich verwendet

Tabellen-Highlighting ist unabhängig von den direkten Zellformaten: Zeilenfarben bleiben beim Sortieren an ihrer Position, während Werte, Formeln und eigene Zellformate gemeinsam verschoben werden. Direkte Zellformate haben Vorrang vor dem Tabellenstil, bedingte Regeln können sie überlagern. Tabellen lassen sich ohne Datenverlust wieder entfernen; alle Aktionen unterstützen Undo/Redo.

```ts
import { formatTable, TABLE_THEMES } from '@sheets/core';
formatTable(wb, 0, toRange('A1:C20'), {
  name: 'Sales', headerRow: true, bandedRows: true, theme: TABLE_THEMES.blue,
});
```

Tabellenbereiche, Kopfzeilen und Farben bleiben in JSON und als native XLSX-Tabellen einschließlich eigener Tabellenstile erhalten. Leere oder doppelte Überschriften werden beim Erstellen eindeutig benannt. Beim Import fremder eingebauter Excel-Tabellenstile dient Blau als Darstellungsstandard; eigene Tabellenstile mit Körper-, Kopf- und Zeilenstreifenformaten werden übernommen. Wertefilter mit Suche, Mehrfachauswahl und leeren Werten unterstützen Undo/Redo und bleiben als native XLSX-AutoFilter erhalten. Mehrere Spaltenfilter werden mit UND verknüpft. Erweiterte Excel-Filter (z. B. Datumsgruppen, Farb- und benutzerdefinierte Vergleichsfilter) und berechnete Tabellenspalten sind noch nicht implementiert.

### Dynamische Arrays

`FILTER`, `SORT`, `UNIQUE` und `SEQUENCE` geben Arrays aus. Auch bestehende Array-Funktionen wie `TRANSPOSE`, Array-Literale und Bereichsausdrücke können in benachbarte Zellen überlaufen. `A1#` referenziert den aktuellen Überlaufbereich einer Formel. Das Verhalten orientiert sich am [Excel-Überlaufoperator](https://support.microsoft.com/en-us/excel/spilled-range-operator).

```ts
setValues(wb, 0, 0, 3, [['=SORT(UNIQUE(FILTER(A2:B20,B2:B20>100)),2,-1)']]);
setValues(wb, 0, 0, 6, [['=SUM(D1#)']]);
```

Blockierende Zellinhalte, verbundene Zellen oder das Blattende ergeben `#SPILL!`. Einzelne Überlaufzellen lassen sich formatieren, aber nur über die Formelzelle inhaltlich ändern. Größenänderungen, Kopieren, Undo/Redo und blattübergreifende Abhängigkeiten werden berücksichtigt. Abgeleitete Zellen werden für JSON aus der Formel neu erzeugt; XLSX enthält Array-Formeln, zwischengespeicherte Ausgaben und dynamische Array-Metadaten. `Workbook.evaluate()` liefert weiterhin einen skalaren Wert (bei Arrays das erste Element).

### Bedingte Formatierung

```ts
import { setConditionalFormats } from '@sheets/core';

setConditionalFormats(wb, 0, [
  { id: 'high', range: 'B2:B20', type: 'cellIs', operator: 'greaterThan',
    formulas: ['1000'], style: { fill: '#C6EFCE', color: '#006100' } },
  { id: 'row', range: 'A2:C20', type: 'expression',
    formulas: ['$B2<0'], style: { fill: '#FFC7CE' }, stopIfTrue: true },
]);
const displayedStyle = wb.getCellStyle(0, 1, 1);
```

Frühere Regeln haben höhere Priorität. Formelreferenzen sind relativ zur linken oberen Ecke des Regelbereichs; absolute Referenzen bleiben fest. `colorScale` unterstützt zwei oder drei Hex-Farben, mit Minimum/Maximum und bei drei Farben der Mitte des Wertebereichs. Grundformate bleiben erhalten. Regeln werden mit Undo/Redo und JSON sowie als native [OOXML-Regeln und Differentialformate](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-conditional-formatting) in XLSX gespeichert. Für größere Lese-/Renderdurchläufe kann ein `ConditionalFormatEvaluator` wiederverwendet werden; nach Änderungen ist ein neuer Resolver nötig.

### Pivot-Auswertungen

```ts
import { createPivotTable, refreshPivotTable } from '@sheets/core';

wb.addSheet('Report');
createPivotTable(wb, {
  id: 'sales',
  source: { sheet: 'Sheet1', range: 'A1:C100' }, // erste Zeile: Überschriften
  target: { sheet: 'Report', cell: 'A1' },
  rows: ['Region'], columns: ['Product'],
  values: [{ field: 'Sales', aggregate: 'sum' }],
});
refreshPivotTable(wb, 'sales'); // nach Änderungen an den Quelldaten
```

Felder werden über Überschriften oder nullbasierte Spaltenpositionen innerhalb der Quelle angegeben. Mehrere Gruppen und Wertefelder, Kriterienfilter (`filters`), `sum`/`count`/`average`/`min`/`max` sowie Gesamtsummen werden unterstützt. Mittelwerte der Gesamtsummen werden aus den Quelldaten berechnet. `evaluatePivot()` erstellt eine Matrix ohne Mutation, `removePivotTable()` entfernt Definition und Ausgabe. Erstellen und Aktualisieren schützen fremde Zielzellen und sind jeweils ein Undo-Schritt.

Definitionen bleiben in JSON und in einem eigenen XML-Teil der XLSX-Datei erhalten; die Ausgabe besteht aus normalen Tabellenzellen. **Native Excel-Pivot-Tabellen und Pivot-Caches werden noch nicht importiert oder erzeugt.**

Blattumbenennungen sowie eingefügte/gelöschte Zeilen und Spalten passen die Definitionen an. Wird die Kopfzeile oder ein verwendetes Quellfeld gelöscht, entfällt die Pivot-Definition; die bisherige Ausgabe bleibt als normale Tabelle erhalten. Undo stellt die Definition wieder her.

### Agent-API

```ts
import { SpreadsheetAgent, toAnthropicTools, AGENT_SYSTEM_PROMPT } from '@sheets/agent';

const agent = new SpreadsheetAgent(wb, { host: view.agentHost }); // host optional (Auswahl/Navigation)
const tools = toAnthropicTools(agent.tools);                      // an die Messages-API übergeben

// tool_use-Block aus der Modellantwort ausführen:
const result = agent.execute('fill_formula', { range: 'D2:D100', formula: '=B2*C2' });
// → { ok: true, result: { filled: 'Sheet1!D2:D100', cells: 99, sample: [...] } }
```

45 Tools, unter anderem: `get_workbook_info`, `read_range` (JSON/Markdown), `write_range`, `fill_formula`, `format_range`, `sort_range`, `copy_range`, `find`, `evaluate_formula`, `inspect_cell` (Vorgänger/Nachfolger), `create_chart`/`update_chart`/`list_charts`, Zeilen/Spalten/Blätter, `freeze_panes`, `define_name`, `undo`/`redo`. Hinzu kommen `format_table`/`list_tables`/`remove_table`, `set_conditional_formats`/`list_conditional_formats` und `create_pivot_table`/`refresh_pivot_table`/`list_pivot_tables`/`remove_pivot_table`, `replace`, `set_hidden` sowie `set_data_validation`/`list_data_validations`.

Jeder Tool-Aufruf wird ein einzelner Undo-Schritt mit `origin: 'agent'`. Nutzer können also nachvollziehen und rückgängig machen, was ein Agent geändert hat. Fehler kommen als `{ ok: false, error }` mit einer Meldung, die ein Modell verstehen kann.

**Eingebettete UIs** (iframe/Webview) lassen sich mit `serveAgent()` und `connectAgent()` per JSON-RPC über `postMessage` fernsteuern. Die Web-App macht das automatisch, wenn sie in einem iframe läuft, und stellt zusätzlich `window.sheets.execute(name, input)` bereit (z. B. für Playwright oder Browser-Agents).

### MCP-Server

```bash
npm run build -w @sheets/mcp
node packages/mcp/dist/server.js --root ~/Dokumente/Tabellen [--readonly] [--no-autosave]
```

Konfiguration für Claude Code / Claude Desktop:

```json
{ "mcpServers": { "sheets": { "command": "node", "args": ["/pfad/zu/sheets/packages/mcp/dist/server.js", "--root", "/pfad/zu/daten"] } } }
```

Alle Tools erhalten einen Parameter `file`. Zusätzlich gibt es `create_workbook` und `save_workbook` (`save_as` konvertiert z. B. CSV → XLSX). Der Zugriff ist auf `--root` beschränkt.

### VS-Code-Extension

```bash
npm run build      # oder: npm run package → artifacts/*.vsix
code --extensionDevelopmentPath=$PWD
```

- `.xlsx` und `.xlsm` öffnen standardmäßig im Sheets-Editor, `.csv`/`.tsv` über „Öffnen mit…“ bzw. Kontextmenü → *Open with Sheets*.
- Speichern, Undo/Redo, Dirty-Status und Wiederherstellen nach Absturz laufen über VS Code. Beim Speichern einer CSV bleibt das erkannte Trennzeichen erhalten.
- **Language-Model-Tools** (`sheets_*`) stehen z. B. dem Copilot-Agent-Modus zur Verfügung. Ist die Datei geöffnet, wirkt die Änderung sofort im Editor und lässt sich rückgängig machen. Sonst wird die Datei geladen, geändert und gespeichert (abschaltbar über `sheets.agentTools.headlessWrites`).
- Andere Extensions nutzen `vscode.commands.executeCommand('sheets.executeTool', { name, input, file })` oder die API, die `activate()` zurückgibt.

## Bekannte Grenzen / nächste Schritte

- Native Excel-Pivot-Caches und Bilder fehlen noch. Notizen und Hyperlinks werden gelesen und angezeigt, aber im Editor nicht bearbeitet.
- Bedingte Formatierung unterstützt Zellvergleiche, Formelregeln, Min/Max-Farbskalen mit optionaler 50%-Mitte, Datenbalken und Symbolsätze. Excel-2010-Erweiterungen (z. B. negative Datenbalken) und andere Regeltypen werden noch nicht importiert oder dargestellt.
- Dynamische Arrays rechnen die Mappe vollständig neu; `SEQUENCE` und Pivot-Auswertungen sind auf eine Million Zellen begrenzt. Der implizite Schnittmengenoperator `@` außerhalb von Tabellenverweisen fehlt noch; eine Tabelle wächst beim Tippen in die Nachbarspalte nicht automatisch mit.
- Zeilen einfügen/löschen und Sortieren werden bei XLSX-Dateien mit nicht vollständig modellierten Inhalten weiterhin blockiert, da dafür alle abhängigen Teile im Original-XML verschoben werden müssten.
- Beim Löschen von Zeilen/Spalten oder Blättern speichert Undo eine Kopie der ganzen Mappe. Für sehr große Dateien sollte das durch gezielte Umkehr-Operationen ersetzt werden.
- Abhängigkeiten von Bereichen werden linear durchsucht; für sehr viele Bereichsformeln wäre ein Intervall-Index sinnvoll.
- Das Datumssystem 1904 wird beim Import nicht umgerechnet.
