# Security review — 2026-10-24

Geprüft wurde der lokale Arbeitsstand der VS-Code-Extension einschließlich Core, Webview, Agent-/Copilot-Schnittstellen, optionalem MCP-Server, Update-Installer und Build-/Publishing-Abhängigkeiten. Bereits vorhandene uncommittete Änderungen wurden beibehalten und mitgeprüft. Die folgende Tabelle unterscheidet neu ergänzte Korrekturen von bereits vorhandenen Schutzmaßnahmen.

## Ergänzte Korrekturen

| Befund | Bewertung | Korrektur und Nachweis |
| --- | --- | --- |
| Verwundbare transitive MSAL-Abhängigkeit über `@vscode/vsce → @azure/identity` | High laut npm/GitHub; Publishing-Toolchain, nicht im Extension-Bundle | `@azure/msal-node` über einen auf `@azure/identity` begrenzten Override auf `^7.1.1` gesetzt. Der anfängliche Scan meldete zwei High-Einträge für dieselbe Ursache (MSAL und abhängiges Identity-Paket); danach keine bekannten Schwachstellen. Paketierung mit der aktualisierten Toolchain erfolgreich. |
| Remote-Dateizugriff prüfte nur das letzte Pfadsegment auf Symlinks und ignorierte Providerfehler | Mittel | Alle Komponenten bis zur Workspace-Wurzel werden geprüft. Providerfehler außer `FileNotFound` sperren den Zugriff; neue Dateien bleiben möglich. Remote-Pfade werden unabhängig vom lokalen Betriebssystem hinsichtlich Groß-/Kleinschreibung verglichen. Mehrdeutige Backslash-Pfade werden abgewiesen. Tests prüfen, dass abgewiesene Zugriffe keine Dateiinhalte lesen oder schreiben. |
| Agenten konnten schreibgeschützte geöffnete Arbeitsmappen im Speicher ändern; Revert setzte `readOnly: false` | Mittel | Schreibwerkzeuge werden im Extension-Host und im Webview-Agenten blockiert. Revert übernimmt den Schreibschutz des Dateisystems. Regressionstests decken Lesen, Schreiben und wiederholte Initialisierung mit wechselndem Schreibschutz ab. |
| ZIP-Limits vertrauten bei unkomprimierten Einträgen auf `originalSize`, obwohl fflate nach komprimierter Größe kopiert; spätere Header wurden erst nach früheren Dekompressionen geprüft | Mittel | Inkonsistente Größen gespeicherter Einträge werden abgewiesen. Ein eigener Durchlauf validiert alle Einträge ohne Dekompression, erst danach wird entpackt. Regressionstests verwenden manipulierte Größen und einen beschädigten frühen Eintrag vor einem übergroßen späteren Eintrag. |
| Update-URLs wurden nur als rohe Zeichenketten mit einem Repository-Präfix verglichen | Niedrig; zusätzliche Absicherung bei manipulierten Release-Metadaten | Vergleich nach URL-Normalisierung mit Prüfung von HTTPS/Origin, Repository-Pfad und fehlenden Zugangsdaten. Dot-Segmente, kodierte Separatoren und Backslash-Traversal können die Repository-Grenze nicht umgehen. Regressionstests prüfen rohe und kodierte Traversal-Pfade. |

## Nachtest mit macOS-MCP: zusätzliche Befunde

Ein separater VS-Code-Extension-Development-Host mit temporärem Profil und eigens erzeugten Testdateien wurde über macOS-MCP und die echte Webview bedient. Die normalen VS-Code-Fenster und Nutzerdaten waren nicht Teil der Angriffssitzung.

| Befund | Reproduktion vor dem Fix | Korrektur und Nachtest |
| --- | --- | --- |
| XLSX mit `col min="Infinity" max="Infinity"` oder `9007199254740992` blockierte den Import | Beide kleinen Dateien überschritten im begrenzten Core-Prozess 1,8 Sekunden ohne Abschluss. Beim Öffnen in der echten Webview blieb der Editor leer und reagierte nach drei Sekunden nicht. | Numerische XML-Attribute müssen endlich sein. Spalten-, Zeilen- und Zellindizes müssen ganzzahlig und innerhalb der Excel-Grenzen liegen. Beide Dateien zeigen im neuen Desktop-Host sofort einen Ladefehler; der Editor sperrt das Speichern. Legale Grenzindizes bleiben lesbar. |
| Backtracking-RegExp in Suche und Formel-Wildcards blockierte die Ausführung | `(a+)+$` gegen 30 `a` und `!` überschritt 1,8 Sekunden; ein mehrfaches `*a` in `COUNTIF` überschritt 1,5 Sekunden. Bei größerer Eingabe wächst die Laufzeit stark. | Suche, Ersetzen und Formel-Wildcards verwenden nun `re2js`, einen Regex-Interpreter mit linearer Laufzeit. Muster sind auf 4096 Zeichen begrenzt; nicht unterstützte Regex-Funktionen wie Backreferences/Lookarounds werden als Fehler zurückgegeben. Im macOS-Host blieb die UI nach nativer Eingabe von `(a+)+$` bedienbar; der Webview-Status meldete „No matches“. |
| Große manuell oder durch Agenten ausgewählte Bereiche konnten bei Rahmen, Kopieren oder TSV-Ausgabe praktisch unbegrenzt Zellen anlegen | `A1:XFC1048575` ließ alle drei Operationen im begrenzten Prozess hängen. | Dichte Kopier-, Einfüge-, Rahmen-, TSV- und Fülloperationen prüfen den Bereich vor der Allokation und lehnen mehr als eine Million Zellen ab. Im echten UI-Test wird ein übergroßer Kopiervorgang direkt mit einer Fehlermeldung abgelehnt, ohne Webview-Ausnahme. |
| Ein nicht validierter Style-Farbwert injizierte HTML in die Zwischenablage | `format_range` mit `color: '#000"/><img src=x onerror=alert(1)><td style="color:red'` lieferte beim echten Kopier-Event `<img ... onerror=...>` in `text/html`. Das Zielprogramm hätte dieses HTML außerhalb der Webview-CSP verarbeitet. | HTML für die System-Zwischenablage übernimmt nur valide Hex-Farben. Der Desktop-Regressionsfall kopiert dieselbe Zelle und prüft, dass kein eingeschleustes Markup vorhanden ist. |

Quelle für die Eigenschaften und Syntax des eingesetzten Regex-Interpreters: [RE2JS-Projekt](https://github.com/le0pard/re2js). Die Bibliothek ist jetzt Laufzeitabhängigkeit von `@sheets/core`; ihre Lizenz steht in `THIRD_PARTY_NOTICES.md`.

MSAL-Quelle: [GHSA-xcx6-jx8p-x3h5 / CVE-2026-109461](https://github.com/advisories/GHSA-xcx6-jx8p-x3h5). Der Hersteller nennt Version 7.1.0 als korrigiert. Der [Hersteller-Changelog](https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-node/CHANGELOG.md) beschreibt die Änderung in Version 7; Azure Identity setzt im hier geprüften interaktiven Aufruf keine entfernte `responseMode`-Option. Der Override ist nötig, solange Identity selbst noch MSAL `^6.0.0` verlangt. Eine echte Anmeldung beim Marketplace/Azure wurde nicht ausgeführt.

## Bereits vorhandene Korrekturen mitgeprüft

- Headless-Dateizugriffe sind auf Tabellen im Workspace begrenzt; lokale und MCP-Pfade werden über reale Symlinkziele geprüft, auch bei noch nicht existierenden Dateien und hängenden Links. Workspace Trust und die Einstellung für Headless-Schreibzugriffe bleiben wirksam.
- Importierte Diagrammfarben werden validiert und SVG-Attribute maskiert. Die Webview verwendet eine restriktive CSP, einen kryptografischen Script-Nonce sowie gesperrte Formulare und Basis-URLs.
- XLSX-Archive haben Limits von 512 MiB je Teil, 1 GiB insgesamt und 20.000 Einträgen. XML-Entity-Verarbeitung ist deaktiviert; externe XML-Ressourcen werden nicht geladen. XLSM-Makros werden erhalten, nicht ausgeführt.
- Update-Installationen benötigen den von GitHub gelieferten SHA-256-Hash und begrenzen Downloads auf 100 MiB. Dateispeicherung prüft zwischenzeitliche externe Änderungen.
- CI-Actions sind auf Commit-SHAs fixiert. Zusätzlich wurde `npm audit --audit-level=high` als CI-Prüfung ergänzt.

## Ausgeführte Prüfungen

- `npm audit --json`: **0 bekannte Schwachstellen**, einschließlich Entwicklungsabhängigkeiten.
- `npm test`: **911 Tests in 44 Dateien erfolgreich** nach den zusätzlichen Regressionstests.
- `npm run build:all`: Typecheck, Extension-/Webview-Bundles und MCP-Bundle erfolgreich.
- `npm run package`: VSIX erfolgreich erstellt und geprüft; tatsächliches Paket aktiviert mit nachgebildeter VS-Code-API, 17 Commands, zwei Editoren und 45 Agentenwerkzeugen.
- Zusätzlicher Stdio-Test gegen den gebauten MCP-Server: Traversal, Symlinks in Elternverzeichnissen, hängende Symlinks, nicht unterstützte Dateitypen und schreibende Aufrufe mit `--readonly` abgewiesen; reguläres Lesen, Erstellen und Schreiben erfolgreich.
- Musterprüfung auf verbreitete Token-/Private-Key-Formate in Quellen, Tests, Skripten und CI: keine Treffer. Keine Prüfung der gesamten Git-Historie.
- `git diff --check`: erfolgreich.
- macOS-MCP: echte VS-Code-Testsitzung mit manipulierten XLSX-Dateien, nativer Such-Eingabe und Webview-Kopiervorgang. Reproduktionen vor dem Fix liegen als Testdateien und begrenzte Prozessprotokolle unter `artifacts/security-macos/`; anschließend öffneten beide fehlerhaften Dateien mit Ladefehler. `npm run test:ui` bestand mit **17 Desktop-Szenarien**, einschließlich der neuen Clipboard- und Großbereichsprüfungen.

## Grenzen

Dies ist ein lokaler Code- und Dependency-Review mit Regressionstests und einem grafischen macOS-Test, kein vollständiger Penetrationstest. Ein neuer Remote-SSH-E2E-Lauf wurde nicht durchgeführt. Remote-Symlink-Schutz setzt voraus, dass der Dateisystemprovider `FileType.SymbolicLink` korrekt meldet. Prüfung und anschließender Dateizugriff sind über die VS-Code-Provider-API nicht atomar; ein konkurrierender Prozess mit Schreibrechten im Workspace bleibt außerhalb dieser Garantie. ZIP-Größenlimits ersetzen keine vollständige CPU-/Speicherisolation beim Verarbeiten komplexer Tabellen. Der sichere Regex-Interpreter unterstützt nicht die gesamte JavaScript-RegExp-Syntax; insbesondere komplexe Lookarounds und Backreferences können nicht mehr als Suchmuster verwendet werden.

Artefakt: `artifacts/sheets-vscode-1.2.2.vsix`. Es wurde lokal erstellt; keine Veröffentlichung, Installation oder Versionsanhebung durchgeführt.
