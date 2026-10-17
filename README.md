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
  vscode/  VS-Code-Extension (Custom Editor + Language-Model-Tools)
```

Abhängigkeiten zur Laufzeit: nur `fflate` (ZIP) und `fast-xml-parser` im Core.

## Schnellstart

```bash
npm install
npm run dev          # Web-App auf http://localhost:5173
npm test             # Vitest
npm run typecheck    # tsc -b über alle Pakete
npm run build        # Pakete, Web-App, MCP-Server und VS-Code-Extension
```

Workspace-Pakete werden über die Export-Bedingung `source` direkt aus den TypeScript-Quellen eingebunden (Vite, esbuild, Vitest). Für die Veröffentlichung auf npm erzeugt `tsc -b` `dist/`.

## Architektur

**Alle Änderungen sind Operationen.** Jede Mutation ist eine JSON-serialisierbare `Operation` (`setCells`, `insertRows`, `setCharts`, …), die `Workbook.apply()` ausführt. Daraus ergeben sich:

- Undo/Redo: Jede Operation liefert ihre Umkehrung, Transaktionen bündeln sie zu einem Schritt.
- ein Change-Event pro Transaktion mit `origin` (`user`, `agent`, `undo`, …)
- eine Grundlage für Protokollierung, Kollaboration und Fernsteuerung

High-Level-Befehle in `commands.ts` (Eingabe, Formatieren, Kopieren/Einfügen, Ausfüllen, Sortieren, Diagramme, Suche) bauen auf diesen Operationen auf. UI und Agent verwenden dieselben Befehle.

**Formel-Engine:** Tokenizer → Pratt-Parser (Excel-Präzedenz) → Evaluator mit Array-Broadcasting. Rund 130 Funktionen (Mathe, Statistik, Logik, Lookup inkl. `XLOOKUP`, Text, Datum). Ein Abhängigkeitsgraph sorgt dafür, dass nach einer Änderung nur die betroffenen Zellen neu berechnet werden. Zirkelbezüge ergeben `#CIRC!`. Bei Formeln mit unbekannten Funktionen bleibt der in der Datei gespeicherte Wert von Excel erhalten.

**xlsx:** eigener Reader und Writer, kein SheetJS/ExcelJS. Unterstützt werden Werte, Formeln (inkl. Shared Formulas, `_xlfn.`-Präfixe), Styles (Schrift, Farben inkl. Theme/Tint, Füllung, Rahmen, Ausrichtung, Zahlenformate), Spaltenbreiten und Zeilenhöhen, verbundene Zellen, fixierte Bereiche, definierte Namen und **native Excel-Diagramme** (Säule, Balken, Linie, Fläche, Kreis, Ring, Punkt), lesend und schreibend. Mit openpyxl in beide Richtungen gegengeprüft.

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

### Agent-API

```ts
import { SpreadsheetAgent, toAnthropicTools, AGENT_SYSTEM_PROMPT } from '@sheets/agent';

const agent = new SpreadsheetAgent(wb, { host: view.agentHost }); // host optional (Auswahl/Navigation)
const tools = toAnthropicTools(agent.tools);                      // an die Messages-API übergeben

// tool_use-Block aus der Modellantwort ausführen:
const result = agent.execute('fill_formula', { range: 'D2:D100', formula: '=B2*C2' });
// → { ok: true, result: { filled: 'Sheet1!D2:D100', cells: 99, sample: [...] } }
```

32 Tools, unter anderem: `get_workbook_info`, `read_range` (JSON/Markdown), `write_range`, `fill_formula`, `format_range`, `sort_range`, `copy_range`, `find`, `evaluate_formula`, `inspect_cell` (Vorgänger/Nachfolger), `create_chart`/`update_chart`/`list_charts`, Zeilen/Spalten/Blätter, `freeze_panes`, `define_name`, `undo`/`redo`.

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
cd apps/vscode && npm run build      # oder: npm run package → .vsix
code --extensionDevelopmentPath=$PWD
```

- `.xlsx`/`.xlsm` öffnen standardmäßig im Sheets-Editor, `.csv`/`.tsv` über „Öffnen mit…“ bzw. Kontextmenü → *Open with Sheets*.
- Speichern, Undo/Redo, Dirty-Status und Wiederherstellen nach Absturz laufen über VS Code. Beim Speichern einer CSV bleibt das erkannte Trennzeichen erhalten.
- **Language-Model-Tools** (`sheets_*`) stehen z. B. dem Copilot-Agent-Modus zur Verfügung. Ist die Datei geöffnet, wirkt die Änderung sofort im Editor und lässt sich rückgängig machen. Sonst wird die Datei geladen, geändert und gespeichert (abschaltbar über `sheets.agentTools.headlessWrites`).
- Andere Extensions nutzen `vscode.commands.executeCommand('sheets.executeTool', { name, input, file })` oder die API, die `activate()` zurückgibt.

## Bekannte Grenzen / nächste Schritte

- Dynamische Arrays mit Überlauf (`FILTER`, `SORT`, `UNIQUE`), Pivot-Tabellen, bedingte Formatierung, Datenüberprüfung, Bilder und Kommentare fehlen noch.
- Beim Löschen von Zeilen/Spalten oder Blättern speichert Undo eine Kopie der ganzen Mappe. Für sehr große Dateien sollte das durch gezielte Umkehr-Operationen ersetzt werden.
- Abhängigkeiten von Bereichen werden linear durchsucht; für sehr viele Bereichsformeln wäre ein Intervall-Index sinnvoll.
- Das Datumssystem 1904 wird beim Import nicht umgerechnet.
