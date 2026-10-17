# Entwicklung

Die Repo-Wurzel ist die VS-Code-Extension. `package.json` ist gleichzeitig Extension-Manifest und npm-Workspace-Manifest.

```
src/              Extension-Host, Copilot-Adapter, Custom Editor, Webview-Einstieg
packages/core/    Tabellenmodell, Formeln, Operationen, XLSX/CSV
packages/ui/      Canvas-Editor und Dialoge
packages/agent/   Agent-Werkzeuge und Adapter
packages/mcp/     Optionaler MCP-Server
apps/web/         Optionale Browser-Demo
test/             Extension-Tests und echter VS-Code-Host-Test
media/            Icon-Quellen, PNG und Vorschau
scripts/          Build, Icon-Generierung, Paketierung und Paketprüfung
```

Die ausführliche [Framework-Dokumentation](docs/framework.md) beschreibt die internen APIs. Weitere Dokumente: [Copilot-Anbindung](docs/copilot.md) und [XLSX-Speicherschutz](docs/xlsx-compatibility.md).

## Lokal testen

1. Node.js 22+ installieren und `npm ci` ausführen.
2. Das Repo in VS Code öffnen und F5 drücken.
3. Im Extension Development Host eine XLSX-Datei öffnen oder „Sheets: Neue Arbeitsmappe“ aufrufen.
4. CSV/TSV über „Mit Sheets öffnen“ testen.

`npm run watch` baut Host und Webview ohne Minifizierung. Die F5-Task wartet, bis beide Bundles vorliegen. Host-Code wird durch „Developer: Reload Window“ neu geladen; für Webview-Code den Tabellen-Tab erneut öffnen. Produktionsbundles entstehen durch `npm run build`.

`npm test` prüft Core, UI, Agent-Werkzeuge und den realen Webview-Nachrichtenhandler mit einem ersetzten DOM-View. `npm run package` erzeugt die VSIX und prüft Inhalt, Manifest, Übersetzungen, Icon und Aktivierung des tatsächlich gepackten Host-Bundles mit einer nachgebildeten VS-Code-API. Diese Prüfungen ersetzen nicht den manuellen Test der VS-Code-Oberfläche.

## Aufräumen und Host-Test

`npm run clean` entfernt ausschließlich Build-Ausgaben und TypeScript-Caches der bekannten Projekte. Quellen und erzeugte VSIX-Artefakte bleiben erhalten. Der frühere Pfad `apps/vscode` wird entfernt, soweit dort nur noch erzeugte Dateien und leere Ordner liegen.

`npm run test:host` baut die Extension und prüft die echte VS-Code-Language-Model-Tool-API in einem separaten lokalen Profil. Details und Voraussetzungen stehen in [Copilot](docs/copilot.md). Der Test sendet keinen Chat an ein Sprachmodell.

## Icon

`media/icon.svg` ist die editierbare Quelle. `npm run icon` rendert daraus das 256×256-PNG für das Extension-Manifest. Beide Dateien einchecken. Das Manifest verwendet PNG entsprechend den [VS-Code-Paketierungsanforderungen](https://code.visualstudio.com/api/working-with-extensions/publishing-extension).

## Paket und Veröffentlichung

`npm run package` erstellt `artifacts/sheets-vscode-<version>.vsix`. Die Positivliste in `.vscodeignore` nimmt ausschließlich Manifest, Übersetzungen, Dokumentation, Lizenz, Icon, Vorschau und die beiden JavaScript-Bundles auf. TypeScript, Source Maps, node_modules, Testdateien und Beispieldateien bleiben draußen. Alle Laufzeitabhängigkeiten außer der VS-Code-API sind gebündelt.

Die Extension-ID bleibt `sheets.sheets-vscode`. Es ist kein Marketplace-Repository oder Publisher-Konto in diesem Repo verifiziert. Für eine Veröffentlichung einen eigenen/verifizierten Publisher und echte Repository-/Support-Links im Manifest hinterlegen und die Anmeldung für `vsce` einrichten. Die lokale Paketierung erlaubt das fehlende Repository ausdrücklich. Die CI erzeugt nur ein herunterladbares VSIX-Artefakt und veröffentlicht nichts.

Bei einem Versionswechsel `package.json`, `package-lock.json`, `CHANGELOG.md` und das Installationsbeispiel in `README.md` aktualisieren.
