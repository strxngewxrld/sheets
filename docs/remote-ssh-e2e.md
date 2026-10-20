# Remote-SSH-E2E-Testbericht

Geprüft am 20. Oktober 2026 mit Sheets 1.1.0, VS Code 1.142.0 und
Microsoft Remote-SSH 0.128.0. Client und SSH-Host liefen auf Windows Server 2025
(x64, Build 26100). Die Verbindung ging über den Windows-OpenSSH-Server auf
`127.0.0.1:22`, mit einem isolierten VS-Code-Profil und Testordner.

## Ergebnis

Sieben Remote-Host-Prüfungen und alle 16 bestehenden UI-Szenarien bestanden.
In den geprüften Abläufen waren keine Änderungen am Extension-Laufzeitcode nötig.

Der Kontext wurde im entfernten Extension Host protokolliert:

- `vscode.env.remoteName === 'ssh-remote'`
- Sheets wurde aus `.vscode-server/extensions/tom-knueppel.sheets-vscode-1.1.0` geladen.
- `extensionKind === Workspace` für Sheets; auch der Prüf-Treiber lief entfernt.
- Das sichtbare VS-Code-Fenster zeigte `SSH: sheets-localtest`.

Damit wurde eine echte Remote-SSH-Verbindung mit entferntem Extension Host geprüft,
einschließlich SSH-Tunnel und lokal gerenderter Webview. Dass derselbe Rechner beide
Rollen übernahm, ersetzte diese Verbindung nicht durch einen lokalen Development Host.
VS-Code-Server und Extension wurden über SSH installiert; der Ablauf folgt der
[offiziellen Remote-SSH-Dokumentation](https://code.visualstudio.com/docs/remote/ssh).

## Geprüfte Abläufe

Die Remote-Host-Prüfung kontrollierte Werkzeugregistrierung, Aufrufe über die echte
VS-Code-LM-API, Headless-Bearbeitung, führende Nullen und Formeln, native und
workspace-relative Pfade sowie escapte Datei-URIs mit Umlauten, `#` und `%`.
Live-Bearbeitung und Speichern erhielten die neu berechneten Formelergebnisse.
Ein CSV-Test erhielt Windows-1252, Semikolon, Dezimalkomma und CRLF bytegenau.

Die UI-Suite lief an der bereits verbundenen Remote-SSH-Sitzung. Sie prüfte
Texteingabe, Windows-Tastenkürzel, Formeln, Einfügen, Undo/Redo, Suchen/Ersetzen,
Dialogwechsel und Fokus, Datenvalidierung, Kontextmenüs, schmale Dialoge,
Tabellen und Filter, Zoom, alle sieben Diagrammtypen, Blattverwaltung,
Light/Dark Theme, XLSX-Speichern/Wiederöffnen und CSV-Speichern.

Zusätzlich wurde über Windows MCP eine entfernte XLSX im Explorer geöffnet,
`RemoteSSH geprüft €` eingegeben und mit Ctrl+S gespeichert. Ctrl+F und Escape
funktionierten im Remote-Fenster. Das anschließende unabhängige Einlesen der
gespeicherten Datei bestätigte den Unicode-Wert und die vorhandene Tabelle.

## Nachweise und Testanbindung

Die lokalen Artefakte liegen unter `artifacts/remote-ssh/`:

- `workspace/remote-host-result.json`: Remote-Kontext und sieben Host-Prüfungen.
- `workspace/result.json`: alle 16 UI-Prüfungen; daneben Screenshots und Testdateien.
- `verification.json`: unabhängige Kontrolle von Ergebnissen und manuell gespeicherter XLSX.
- `cleanup.json`: Entfernung des temporären Zugangs und des Prüf-Treibers.

`scripts/test-ui.mjs` kann sich jetzt an einen bereits laufenden Testhost anbinden:

```powershell
$env:SHEETS_UI_ATTACH_DIR = 'C:\Pfad\zum\Testordner'
$env:SHEETS_UI_DEBUG_PORT = '9345'
node scripts/test-ui.mjs
```

Voraussetzung sind eine aktive Test-Bridge entsprechend `test/ui/host.cjs` im
entfernten Workspace-Host und ein nur für den Test geöffneter lokaler DevTools-Port.
Das Testverzeichnis muss hier für Client und Host zugänglich sein. Ohne die beiden
Variablen startet die Suite weiterhin ihren eigenen lokalen VS-Code-Testhost.

## Aufräumen und Grenzen

Der temporäre Administrator-Schlüssel war ausschließlich für localhost zugelassen.
Nach dem Test wurden sein Eintrag und der private Schlüssel entfernt sowie die
Prüf-Erweiterung deinstalliert. Sheets bleibt auf dem VS-Code-Server installiert;
der zuvor aktivierte SSH-Dienst bleibt aktiv.

Dieser Lauf prüfte Windows-Client zu Windows-Host über Loopback. Andere Host-Systeme,
ein macOS-Client, Internet-Latenz, Verbindungsabbrüche und Wiederverbindung wurden
damit nicht nachgewiesen. Die Ergebnisse des lokalen Windows-Tests stehen im
[Windows-E2E-Bericht](windows-e2e.md).
