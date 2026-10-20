# Entwicklung

Die Repo-Wurzel ist die VS-Code-Extension. `package.json` ist gleichzeitig Extension-Manifest und npm-Workspace-Manifest.

```
src/              Extension-Host, Copilot-Adapter, Custom Editor, Webview-Einstieg
packages/core/    Tabellenmodell, Formeln, Operationen, XLSX/CSV
packages/ui/      Canvas-Editor und Dialoge
packages/agent/   Agent-Werkzeuge und Adapter
packages/mcp/     Optionaler MCP-Server
test/             Extension-Tests und echter VS-Code-Host-Test
media/            Icon-Quellen, PNG und Vorschau
scripts/          Build, Icon-Generierung, Paketierung und Paketprüfung
tools/            Veröffentlichung des VSIX über GitHub Releases
```

Die ausführliche [Framework-Dokumentation](docs/framework.md) beschreibt die internen APIs. Weitere Dokumente: [Copilot-Anbindung](docs/copilot.md) und [XLSX-Speicherschutz](docs/xlsx-compatibility.md).

## Lokal testen

1. Node.js 22+ installieren und `npm ci` ausführen.
2. Das Repo in VS Code öffnen und F5 drücken.
3. Im Extension Development Host eine XLSX-Datei öffnen oder „Sheets: Neue Arbeitsmappe“ aufrufen.
4. CSV/TSV über „Mit Sheets öffnen“ testen.

`npm run watch` baut Host und Webview ohne Minifizierung. Die F5-Task wartet, bis beide Bundles vorliegen. Host-Code wird durch „Developer: Reload Window“ neu geladen; für Webview-Code den Tabellen-Tab erneut öffnen. Produktionsbundles entstehen durch `npm run build`.

`npm test` prüft Core, UI, Agent-Werkzeuge und den realen Webview-Nachrichtenhandler mit einem ersetzten DOM-View. `npm run package` erzeugt die VSIX und prüft Inhalt, Manifest, Übersetzungen, Icon und Aktivierung des tatsächlich gepackten Host-Bundles mit einer nachgebildeten VS-Code-API. Diese Prüfungen ersetzen nicht den manuellen Test der VS-Code-Oberfläche.

Die GitHub-Actions-Matrix führt Typecheck, Tests und VSIX-Paketierung unter Ubuntu, Windows und macOS aus. Die Artefakte heißen `sheets-vscode-<runner-os>` und enthalten jeweils dieselbe plattformunabhängige Extension. Die grafischen Host-Tests werden separat lokal ausgeführt.

## Aufräumen und Host-Test

`npm run clean` entfernt ausschließlich Build-Ausgaben und TypeScript-Caches der bekannten Projekte. Quellen und erzeugte VSIX-Artefakte bleiben erhalten. Der frühere Pfad `apps/vscode` wird entfernt, soweit dort nur noch erzeugte Dateien und leere Ordner liegen.

`npm run test:host` baut die Extension und prüft die echte VS-Code-Language-Model-Tool-API in einem separaten lokalen Profil. Details und Voraussetzungen stehen in [Copilot](docs/copilot.md). Der Test sendet keinen Chat an ein Sprachmodell.

`npm run test:ui` baut die Extension und prüft die gerenderte Custom-Editor-Webview in einem eigenen Extension Development Host über Chromium DevTools. Voraussetzung ist eine lokale VS-Code-Desktop-Installation; mit `VSCODE_EXECUTABLE` lässt sich deren Launcher vorgeben. Die Prüfung verwendet ausschließlich erzeugte XLSX-/CSV-Dateien und ein temporäres Profil. Sie testet unter anderem Eingabe, Formeln, Einfügen, Suchen/Ersetzen, Datenüberprüfung, Menütastatur und Listener-Lebensdauer, schmale Dialoge, Tabellenfilter, alle Diagrammtypen, Blattverwaltung, Zoom, Light/Dark Theme und Speichern/Wiederöffnen. Ergebnis, Testdateien und Screenshots bleiben im ausgegebenen temporären Verzeichnis erhalten. Die Verbindung bleibt lokal; zusätzliche Browserpakete werden nicht benötigt. Native macOS-Menüs und Systemdialoge sind weiterhin manuell zu prüfen.

### Windows-E2E-Nachweis

Der [Windows-Testbericht](docs/windows-e2e.md) beschreibt den geprüften Stand und die Grenzen der Abdeckung. Beide E2E-Starter entfernen geerbtes `ELECTRON_RUN_AS_NODE` und behalten ihre temporären Ergebnisdateien. Mit `SHEETS_EXTENSION_PATH` kann statt des Workspace-Builds das Verzeichnis einer separat installierten Extension geprüft werden; die Tests selbst kommen weiterhin aus diesem Repository.

Der [Remote-SSH-Testbericht](docs/remote-ssh-e2e.md) dokumentiert den Lauf mit entferntem Windows-Extension-Host. Mit `SHEETS_UI_ATTACH_DIR` und `SHEETS_UI_DEBUG_PORT` kann die UI-Suite eine bereits laufende, isolierte Testsitzung mit aktiver Test-Bridge verwenden.

## Icon

`media/icon.svg` ist die editierbare Quelle. `npm run icon` rendert daraus das 256×256-PNG für das Extension-Manifest. Beide Dateien einchecken. Das Manifest verwendet PNG entsprechend den [VS-Code-Paketierungsanforderungen](https://code.visualstudio.com/api/working-with-extensions/publishing-extension).

## Paket und Veröffentlichung

`npm run package` erstellt `artifacts/sheets-vscode-<version>.vsix`. Die Positivliste in `.vscodeignore` nimmt ausschließlich Manifest, Übersetzungen, Dokumentation, Lizenz, Icon, Vorschau und die beiden JavaScript-Bundles auf. TypeScript, Source Maps, node_modules, Testdateien und Beispieldateien bleiben draußen. Alle Laufzeitabhängigkeiten außer der VS-Code-API sind gebündelt.

Die Extension-ID lautet `Tom-Knueppel.sheets-vscode` (Publisher `Tom-Knueppel` aus `package.json`); der Host-Test (`test/host/index.cjs`) sucht die Extension unter dieser ID. Das Repository im Manifest verweist auf `strxngewxrld/sheets`. Für eine Marketplace-Veröffentlichung ist weiterhin ein eigener/verifizierter Publisher und die Anmeldung für `vsce` nötig. Die CI erzeugt nur ein herunterladbares VSIX-Artefakt und veröffentlicht nichts.

### GitHub Releases

`tools/publish-github.mjs` führt `npm test` und `npm run package` aus und lädt die geprüfte `artifacts/sheets-vscode-<version>.vsix` direkt als Asset eines Releases `v<version>` hoch. Es gibt keine npm-Hülle, keine npm-Veröffentlichung und keine `.npmrc`. Die Extension-ID `Tom-Knueppel.sheets-vscode` bleibt unverändert.

```bash
# Lokal prüfen: kein Token und keine GitHub CLI nötig, keine Versionsänderung
npm run publish:github -- --dry-run

# Den vorbereiteten und gepushten Stand veröffentlichen; Tag wird automatisch angelegt
npm run publish:github

# Optional nur einen Release-Entwurf mit VSIX anlegen
npm run publish:github -- --draft
```

Für einen echten Upload die [GitHub CLI](https://cli.github.com/) installieren und `gh auth login --hostname github.com` ausführen. Alternativ `GH_TOKEN` oder `GITHUB_TOKEN` in der Umgebung setzen. Der Token benötigt Schreibzugriff auf Repository-Inhalte (`Contents: write` bei einem Fine-grained Token; `contents: write` für den Actions-Token). `NODE_AUTH_TOKEN` und `packages: write` werden nicht mehr verwendet. Tokens werden nicht in Dateien geschrieben.

Das Ziel ist standardmäßig `strxngewxrld/sheets`; `--owner` oder `GITHUB_OWNER` können den Owner ändern. Das Skript verlangt ein sauberes Git-Arbeitsverzeichnis und einen bereits nach GitHub gepushten Release-Commit. Nach erfolgreichen Tests und Paketprüfung legt es den fehlenden Tag `v<package.json-Version>` auf GitHub exakt auf diesem Commit an. Ein vorhandener einfacher oder annotierter Tag muss auf denselben Commit zeigen; abweichende Tags werden nicht verschoben. Das gilt auch bei `--draft`: Der Tag wird angelegt, der Release bleibt ein Entwurf. Es erzeugt keine Commits und überschreibt keine vorhandenen Releases oder Assets. Der Tag bleibt bestehen, falls der anschließende Release-Upload fehlschlägt, und kann beim nächsten Versuch wiederverwendet werden. Bei Bedarf den neuen Tag lokal mit `git fetch origin --tags` holen. Ein gescheiterter Upload kann einen Release-Entwurf hinterlassen; diesen vor einer Wiederholung auf GitHub prüfen und gegebenenfalls entfernen.

Eine neue Version vorab vorbereiten:

```bash
npm version patch --no-git-tag-version --ignore-scripts --workspaces=false
# CHANGELOG.md und das Installationsbeispiel in README.md aktualisieren
# Änderungen committen und pushen; das Publishing legt den Versionstag an
npm run publish:github
```

Die frühere Option `--bump none` bleibt kompatibel; andere `--bump`-Werte werden mit einem Hinweis auf die Vorbereitung abgelehnt. Standardmäßig wird die bereits vorbereitete Version verwendet.

Der Dry-Run führt alle lokalen Tests und die VSIX-Prüfung aus und schreibt `artifacts/v<version>-release-plan.json` mit Ziel, Commit, Asset, SHA-256, der geplanten Tag-Erstellung und dem Release-CLI-Aufruf sowie `artifacts/v<version>-release-notes.md`. Er kontaktiert GitHub nicht und verändert weder Version noch Git-Historie. Remote-Berechtigungen, Commit-Verfügbarkeit und bestehende Tags werden erst bei einem echten Upload geprüft. Die Release Notes übernehmen den passenden Versionsabschnitt aus `CHANGELOG.md` und ergänzen den VSIX-Installationshinweis. In VS Code steht die Task **Sheets: Publishing-Probelauf** bereit.

Die Veröffentlichung nutzt [`gh release create`](https://cli.github.com/manual/gh_release_create) mit `--verify-tag` und dem VSIX-Pfad, nachdem der Job den Tag über die [Git-References-API](https://docs.github.com/en/rest/git/refs#create-a-reference) angelegt oder geprüft hat. Die GitHub CLI lädt Assets zunächst in einen Entwurf und veröffentlicht anschließend; `--draft` lässt den Entwurf bestehen. Die benötigten Berechtigungen beschreibt die [GitHub-Releases-API](https://docs.github.com/en/rest/releases/releases#create-a-release).

Bei einem Versionswechsel `package.json`, `package-lock.json`, `CHANGELOG.md` und das Installationsbeispiel in `README.md` aktualisieren.
