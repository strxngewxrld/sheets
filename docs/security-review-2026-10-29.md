# Sicherheitsprüfung – 29. Oktober 2026

Geprüft wurde der aktuelle lokale Arbeitsstand einschließlich der bereits vorhandenen, nicht committeten Sicherheitskorrekturen. Dieser Bericht ergänzt `security-review.md`; dessen frühere Testergebnisse wurden nicht als neue Prüfergebnisse übernommen. Die unten historisch beschriebenen CSV-Befunde wurden anschließend auf Wunsch behoben; vorhandene Änderungen wurden beibehalten.

## Stand nach der Korrektur

- Alle CSV-/TSV-Schreibpfade begrenzen das vollständige Ausgaberechteck auf eine Million Positionen und die UTF-8-Ausgabe auf 16 MiB. Auch führende Leerzellen, erhaltene Originaldatensätze, Anführungszeichen, BOM und Zeilenenden werden berücksichtigt. Überschreitungen werden vor dem Schreiben abgewiesen. Der reproduzierte XLSX-Fall wird bereits vor dem Zugriff auf die erste Ausgabezelle abgelehnt; die Webview bleibt für weitere Werkzeugaufrufe nutzbar.
- `writeCsv` schützt Formeltext standardmäßig durch einen Tabulator innerhalb eines zitierten Felds. Das gilt auch für String-Ergebnisse von Formeln. Zahlen bleiben unverändert. Neue CSV-/TSV-Konvertierungen über `saveWorkbook` bzw. „Speichern unter“ verwenden ebenfalls diesen Textschutz. Der Exportdialog bietet Textschutz als erste Auswahl und ausdrücklich unveränderte Werte für Datenimporte. `escapeFormulas: false` erlaubt unveränderte Ausgabe über die Core-API.
- Das normale Speichern bestehender CSV-Dateien übernimmt weiterhin die ursprünglichen Werte und Datensätze. Excel-Textschutz verändert Werte und ist keine universelle Garantie für jedes Tabellenprogramm. Für unveränderte Textdaten zur Anzeige in Tabellenprogrammen ist XLSX vorzuziehen; CSV-Exporte können beim Weiterverarbeiten und erneuten Speichern ihren Schutz verlieren.
- Übergröße beim manuellen Export wird als Fehlermeldung angezeigt, bevor ein Zieldialog oder Schreibzugriff erfolgt. Kein zusätzlicher Worker wurde eingeführt; die Größenlimits beheben den nachgewiesenen Exportfall, nicht sämtliche potenziell teuren Formelberechnungen.

## Reproduzierte Befunde vor der Korrektur

### 1. Mittel: CSV-Export kann durch eine kleine XLSX blockiert werden

**Fundstelle:** `packages/core/src/io/csv.ts:293–334`, erreichbar über `src/webview/main.ts:208–211` sowie CSV-Speicherpfade.

`writeCsv` durchläuft jede Position von A1 bis zur letzten benutzten Zeile und Spalte, ohne vorher ein Zell- oder Ausgabebudget zu prüfen. Eine einzelne Zelle in XFD1048576 verursacht dadurch 17.179.869.184 Schleifendurchläufe. Die vorhandenen Limits für Kopieren und andere Bereichsoperationen greifen hier nicht.

**Reproduktion:** Aus einer leeren XLSX wurde mit fflate eine Datei mit genau einer Inline-Textzelle `XFD1048576 = x` erzeugt. Das resultierende Archiv war 3.304 Byte groß. `openWorkbook` importierte erfolgreich genau eine Zelle; der anschließende `writeCsv`-Aufruf überschritt im separaten Node-Prozess das Zwei-Sekunden-Limit und wurde mit SIGTERM beendet. Der Kindprozess war zusätzlich auf 256 MiB V8-Heap begrenzt. Das ist ein Core-Nachweis; ein neuer Desktop-Hängertest wurde nicht durchgeführt.

**Auswirkung:** Ein fremdes Dokument kann nach Auslösen des CSV-Exports oder einer CSV-Konvertierung den ausführenden Webview-Thread blockieren. Bei Headless-/MCP-Konvertierungen läuft die synchrone Verarbeitung im jeweiligen Hostprozess. Ein bloßes Öffnen dieser konkreten Testdatei löste die Blockade nicht aus.

**Empfehlung:** Vor jeder dichten Ausgabe das vollständige Ausgaberechteck inklusive führender Leerzellen begrenzen; zusätzlich ein Bytebudget verwenden. Große Exporte in einem abbrechbaren Worker ausführen. Ein Timeout im aufrufenden Code allein unterbricht die synchrone Schleife nicht.

### 2. Mittel, abhängig vom Zielprogramm: Text wird beim CSV-Export zur Formel

**Fundstelle:** `packages/core/src/io/csv.ts:324–330`, `src/webview/main.ts:208–211`.

Ein expliziter Textwert `{ v: '=1+1' }` wird von `writeCsv` als `=1+1\r\n` ausgegeben, auch ohne die Option `formulas`. Dies wurde im separaten Node-Prozess reproduziert. Der Writer maskiert CSV-Trennzeichen und Anführungszeichen, bietet aber keine Option zur Neutralisierung von Formelpräfixen.

**Auswirkung:** Bei Weitergabe des Exports an Excel oder andere Tabellenprogramme kann fremder Text als Formel interpretiert werden. Weitergehende Auswirkungen hängen von Formel, Zielprogramm, Sicherheitskonfiguration und Benutzerinteraktion ab. Eine Codeausführung oder Datenübertragung wurde nicht demonstriert. Sheets selbst behandelt Formeltext beim standardmäßigen CSV-Import bereits als Text.

**Empfehlung:** Einen ausdrücklich auswählbaren Exportmodus für die sichere Anzeige fremder Textdaten anbieten und vor Formelpräfixen warnen. Den verlustfreien CSV-Speicherpfad nicht stillschweigend verändern: Neutralisierung verändert Daten und ist je nach Zielprogramm unterschiedlich zuverlässig. Siehe [OWASP CSV Injection Testing](https://owasp.github.io/www-project-web-security-testing-guide/latest/4-Web_Application_Security_Testing/07-Injection/21-CSV_Injection) und [CWE-1236](https://cwe.mitre.org/data/definitions/1236.html).

## Weitere Sicherheitsgrenzen und Restrisiken

- **Ressourcen beim Import:** `packages/core/src/io/zip.ts:4–6` erlaubt 512 MiB je ZIP-Teil und 1 GiB insgesamt. Diese Grenzen sind vorhanden, aber hoch; entpackte Daten, XML-Bäume, Workbook und Base64-Kopien benötigen zusätzlich Speicher. Das ist ein verbleibendes Verfügbarkeitsrisiko, kein hier nachgewiesener Bypass des ZIP-Limits. Niedrigere konfigurierbare Budgets und Prozess-/Workerisolation prüfen.
- **Agentenzugriff:** Headless-Zugriffe sind auf Tabellen im Workspace begrenzt; Schreiben erfordert Trust und die entsprechende Einstellung. Bereits geöffnete Dokumente werden dagegen vor diesen Prüfungen verarbeitet (`src/extension.ts:194–213`), auch außerhalb des Workspace. Copilot fordert für Schreibwerkzeuge eine Bestätigung an; die öffentliche Extension-API besitzt diese eigene Bestätigungsoberfläche nicht. Das ist eine bewusst breitere Zugriffsgrenze, kein nachgewiesener Sandbox-Ausbruch. Sensible Tabelleninhalte können über Werkzeugergebnisse beim eingesetzten KI-Anbieter landen. Fremde Zellinhalte als mögliche Prompt-Injection behandeln. Die [VS-Code-Dokumentation](https://code.visualstudio.com/docs/editing/workspaces/workspace-trust) beschreibt die entsprechende Agenten-/Trust-Grenze.
- **Dateisystemrennen:** Realpath-/Symlinkprüfung, Vergleich der Originaldaten und abschließender Schreibzugriff sind getrennte Operationen (`src/extension.ts:206–235`, `packages/mcp/src/server.ts:82–91`). Ein konkurrierender Prozess mit Schreibrechten kann dazwischen eingreifen. Kein Race-Exploit wurde in dieser Prüfung ausgeführt. Remote-Sicherheit hängt zusätzlich von korrekten Symlinkangaben des Providers ab.
- **Update-Vertrauenskette:** Downloads sind größenbegrenzt und werden gegen den von GitHub gelieferten SHA-256 geprüft (`src/update-check.ts:77–110`); die Installation erfordert eine Benutzeraktion. Prüfsumme und Paket stammen jedoch aus derselben Release-Vertrauenskette. Eine Kompromittierung des Release-Kontos wird dadurch nicht abgefangen. Release-Zugänge, CI-Berechtigungen und gegebenenfalls unabhängig geprüfte Signaturen absichern.

## Aktuell ausgeführte Prüfungen

- `npm audit --json`: 0 bekannte Schwachstellen, einschließlich Entwicklungsabhängigkeiten.
- `npm test` nach der Korrektur: **925 Tests in 45 Dateien erfolgreich**. Neue Regressionstests prüfen den importierten XLSX-Angriffsfall, Zellgrenzen, UTF-8- und Quote-Budgets, erhaltene Originaldatensätze, Formeltext einschließlich formatierter Zahlen, unverändertes CSV-Speichern sowie Export und Fehlerbehandlung in der Webview.
- `npm run build:all`: Typprüfung und Builds von Extension, Webview und MCP-Server erfolgreich.
- `git diff --check`: erfolgreich.
- Zwei isolierte Core-Reproduktionen für CSV-Formeltext und übergroße CSV-Ausgabe; anschließend zusätzlicher Nachweis mit einer tatsächlich importierten 3.304-Byte-XLSX.
- Codeprüfung von Webview/CSP, Dateizugriffen, Copilot-Werkzeugen, optionalem MCP-Server, XML-/ZIP-Import und Update-Installation.

Kein neuer Paketvergleich, vollständiger Penetrationstest, Remote-SSH-Test oder Desktop-Test. Die Excel-Interpretation wurde nicht in einer echten Excel-Instanz nachgetestet. Die Aussagen betreffen den lokalen Quellstand und die neu gebauten Bundles, nicht automatisch eine bereits ausgelieferte VSIX. Ein sauberer Dependency-Scan und bestandene Tests schließen eigene Logikfehler nicht aus.
