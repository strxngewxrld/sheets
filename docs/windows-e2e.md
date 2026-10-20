# Windows-E2E-Testbericht

Geprüft am 20. Oktober 2026 unter Windows Server 2025 Datacenter Evaluation
(24H2, Build 26100, x64), VS Code 1.142.0, Node.js 26.8.1 und npm 11.19.0.
Extension-Version: 1.1.0, einschließlich der Änderungen in diesem Arbeitsstand.

## Ergebnis

- 843 automatisierte Tests in 41 Dateien bestanden.
- Typecheck und Produktionsbuild bestanden.
- VSIX erstellt und validiert: 14 Dateien, 15 Befehle, zwei Editoren, 45 Agent-Werkzeuge.
- VSIX erfolgreich in ein isoliertes Extension-Verzeichnis installiert.
- Mit dem aktualisierten installierten Paket zehn Host-Szenarien und 16 UI-Szenarien bestanden.
- Keine unbehandelten Webview-Ausnahmen im UI-Lauf; keine Meldungen
  „No custom document found“ mehr in den finalen UI-Host-Logs.

Die Host-Suite prüft Werkzeugregistrierung und Aufrufe über die echte VS-Code-LM-API,
native Pfade und Datei-URIs mit Leerzeichen, Umlauten, `#` und `%`, führende Nullen,
Windows-1252-CSV, Editorwahl, Live-Bearbeitung, Speichern/Wiederöffnen sowie Schutz
vor dem Überschreiben beschädigter Dateien.

Die UI-Suite prüft Unicode, Eingabe, Formeln, rechteckiges Einfügen,
Undo/Redo, Suchen/Ersetzen, Datenvalidierung, Kontextmenüs und Tastaturnavigation,
Fokuswiederherstellung, Listener-Abbau, schmale Dialoge, Tabellen, Filter,
Auswahlstatistik, Zoom, sieben Diagrammtypen, Blattverwaltung, Light/Dark Theme,
Speichern/Wiederöffnen und CSV-Trennzeichen/CRLF.
Unter Windows verwendet ein zusätzlicher Ablauf echte Chromium-Eingabeereignisse
für Texteingabe und Ctrl+Z/Y/F/S. Die übrigen UI-Schritte verwenden überwiegend
DOM-Ereignisse in der echten Webview.

## Gefundene und behobene Fehler

1. **Teststart unter Windows:** Geerbtes `ELECTRON_RUN_AS_NODE=1` ließ `Code.exe`
   als Node-Prozess starten. Die Starter bereinigen die Kindprozessumgebung.
2. **Editor-Ereignisse:** Dieselbe Änderungsquelle wurde für beide Editor-Typen
   registriert. Änderungen werden nun ausschließlich an die zuständige Registrierung
   weitergeleitet. Eine verspätete Dokumentfreigabe entfernt außerdem keine neu
   geöffnete Instanz derselben Datei mehr aus der Dokumentverwaltung.
3. **XLSX nach erneutem Öffnen:** Zelländerungen in selbst erzeugten Dateien konnten
   eine XML-Darstellung erzeugen, die bei späteren Tabellenänderungen den
   Speicherschutz auslöste. Nachgewiesen vollständig vom Serializer abgebildete
   Dateien werden jetzt auch bei Zelländerungen konsistent neu geschrieben.
   Die strikte Prüfung auf vollständig abgebildete Inhalte bleibt erhalten.
4. **UI-Testzuverlässigkeit:** Der synthetische Rechtsklick ist abbrechbar, damit
   VS Code kein zusätzliches natives Kontextmenü öffnet. Beim Wiederöffnen wartet
   die Suite auf das Entfernen der vorherigen Webview; Speicherprüfungen warten
   auf einen sauberen Dokumentzustand.
5. **Fokus nach „Replace all“:** Wird der angeklickte Ersetzen-Button mangels
   weiterer Treffer deaktiviert, geht der Fokus nun zurück ins Suchfeld.
   Escape schließt dadurch weiterhin zuverlässig das Suchfenster.
6. **Überlagerte Dialoge:** Suchen und Format-/Analyse-Dialoge schließen einander
   beim Öffnen. Der Formatdialog kann nicht mehr hinter dem Suchfenster liegen.
7. **XLSX beim Kaltstart:** VS Code konnte eine per Kommandozeile übergebene Datei
   vor der Editor-Registrierung als Binärdatei im Texteditor öffnen. Sheets übernimmt
   solche unveränderten XLSX-/XLSM-Tabs einmalig bei der Aktivierung und schließt
   den ursprünglichen Text-Tab. Gruppenposition, Vorschauzustand und Fokus werden
   berücksichtigt. Bei expliziten Editor-Zuordnungen auf Benutzer-/Workspace-Ebene
   greift diese konservative Wiederherstellung nicht ein.

Regressionstests decken die bereinigte Umgebung, Ereigniszuordnung,
Dokumentfreigabe und die XLSX-Sequenz Zelländerung → Speichern → Wiederöffnen →
Tabelle anlegen → Speichern ab.
Die UI-Suite prüft zusätzlich den Fokus nach dem letzten ersetzten Treffer,
Escape und den Dialogwechsel in beiden Richtungen.

## Manueller Windows-MCP-Durchlauf

Zusätzlich wurden Maus und Windows-Tastatur im sichtbaren, separaten VS-Code-Host
über Windows MCP verwendet. Neue Testdateien und Kopien liegen ausschließlich unter
`artifacts/windows-manual/`. Geprüft wurden:

- Unicode-Eingabe, Ctrl+Z und Ctrl+S.
- Copy/Paste einer Unicode-Zelle über die echte System-Zwischenablage.
- Ctrl+F/H, Suchen, „Replace all“ und Escape.
- Formatdialog, Fettformatierung und Wechsel zwischen Such- und Formatdialog.
- Kontextmenü per Rechtsklick, Zeile einfügen und rückgängig machen.
- Nativer Windows-Dialog „Speichern unter“ mit Umlaut im Dateinamen.
- Neue Arbeitsmappe über die Befehlspalette und den nativen Speicherdialog.
- „Open Spreadsheet“ mit nativem Dateidialog, CSV-Bearbeitung und Speichern.
- CSV-Export mit Semikolon über die Trennzeichenauswahl und den nativen Speicherdialog.
- Kaltstart des neu installierten finalen Pakets mit einer XLSX-Datei: Sheets
  öffnet sich automatisch, der ursprüngliche Text-Tab verschwindet, Tabelle und
  gespeicherte Unicode-Werte sind sichtbar.

Die beiden Fokus-/Dialogfehler wurden hierbei entdeckt und nach einem Neubau und
Reload erneut über Windows MCP geprüft. Danach wurden die gespeicherten Dateien
unabhängig eingelesen: Unicode-Zellen, Fettformatierung und native XLSX-Tabelle
blieben erhalten; CSV behielt Trennzeichen, CRLF und fehlenden letzten Zeilenumbruch.
Ergebnis: `artifacts/windows-manual/file-verification.json`.
Der später gefundene Kaltstartfehler wurde zusätzlich mit einem echten CLI-Start
in die Host-Suite aufgenommen und visuell mit dem installierten finalen VSIX geprüft.
MCP-Texteingabe mit Tabs/Zeilenumbrüchen wurde nicht als rechteckiger Clipboard-Paste
gewertet, da das Werkzeug diese Zeichen als einzelne Tastatureingaben sendet.

## Nachweise und Wiederholung

Lokale Ergebnisse liegen unter `artifacts/windows-e2e/`:

- `host/result.json` und `host/run.log`
- `ui/result.json`, Screenshots und erzeugte XLSX-/CSV-Dateien
- `package.log`

Das geprüfte Paket liegt unter `artifacts/sheets-vscode-1.1.0.vsix`.
Artefakte und Testprofile werden nicht eingecheckt.

```powershell
npm ci
npm test
npm run test:host
npm run test:ui
npm run package

# Optional: Tests gegen eine separat installierte VSIX-Version.
$env:SHEETS_EXTENSION_PATH = 'C:\Pfad\zum\Extension-Verzeichnis'
node scripts/test-host.mjs
node scripts/test-ui.mjs
```

## Grenzen

Die Änderungen wurden zusätzlich auf macOS-Verträglichkeit im Code geprüft.
Regressionstests decken Cmd+F, Cmd+Option+F, Cmd+Z/Cmd+Shift+Z und POSIX-Pfade
unter `/Users/...` ab; Undo/Redo bleiben beim VS-Code-Host. Der native macOS-Launcher
bleibt erhalten. Diese Tests liefen auf Windows und ersetzen keinen erneuten
macOS-E2E-Lauf des geänderten gemeinsamen Codes.

Dieser Nachweis gilt für die genannte Windows-Umgebung. Windows 10/11 wurden nicht
separat gestartet. Physische AltGr-/IME-Eingabe, UNC-Netzlaufwerke, reale Copilot-Chat-Anfragen und
Update-Downloads wurden nicht Ende zu Ende geprüft. Der Test benötigt keinen
Chat-Aufruf. Die VS-Code-Installation meldet unabhängig von Sheets einen
API-Proposal-Konflikt ihrer mitgelieferten Copilot-Chat-Erweiterung.
