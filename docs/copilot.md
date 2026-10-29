# Copilot-Agent-Werkzeuge

Sheets stellt seine Werkzeuge über `contributes.languageModelTools` und `vscode.lm.registerTool` bereit. Das ist die von VS Code vorgesehene Schnittstelle für Extension-Werkzeuge in Copilot Agent Mode. Ein separater MCP-Server ist dafür nicht erforderlich.

## Verwendung

1. Die aktuelle Sheets-VSIX installieren und ein Fenster mit aktueller VS-Code-Version verwenden. Extension-Werkzeuge in Agent Mode werden seit VS Code 1.99 unterstützt; der manuelle Tabelleneditor unterstützt weiterhin VS Code 1.95+.
2. Copilot Chat öffnen und **Agent** auswählen. Das gewählte Modell muss Werkzeugaufrufe unterstützen.
3. Über **Configure Tools / Werkzeuge konfigurieren** die Werkzeuge von **Sheets** aktivieren. Unternehmensrichtlinien können die Nutzung einschränken.
4. Eine XLSX-Datei in Sheets öffnen oder im Prompt einen eindeutigen Dateipfad angeben. Bei mehreren Dateien den Pfad ausdrücklich nennen.

Beispiel:

> Nutze #sheets_get_workbook_info und #sheets_read_range, um /pfad/Debitoren.xlsx zu prüfen. Lies zunächst A1:F20 und beschreibe die vorhandenen Spalten. Ändere noch keine Daten.

Für eine konkrete Migration:

> Erhalte Debitorennummern als Text mit führenden Nullen. Verwende beim Schreiben write_range mit parse:false und prüfe danach dieselben Zellen erneut mit read_range.

Ändernde Werkzeuge liefern eine Bestätigungsbeschreibung mit Ziel und Eingabedaten an VS Code. Die konkrete Freigabeoberfläche hängt von VS Code, euren Richtlinien und bereits erteilten Freigaben ab. Bei geöffneten Dateien erscheinen Änderungen im Editor und unterstützen Undo/Redo. Änderungen an ungeöffneten Dateien werden gespeichert und benötigen einen vertrauenswürdigen Arbeitsbereich sowie `sheets.agentTools.headlessWrites`.

Ungeöffnete Dateien erreichen die Werkzeuge nur, wenn es sich um `.xlsx`, `.xlsm`, `.csv` oder `.tsv` innerhalb eines Arbeitsbereichsordners handelt. Das gilt auch für lesende Werkzeuge. Symbolische Links, die aus dem Arbeitsbereich hinausführen, werden abgelehnt. So kann ein Agent, der etwa über präparierte Datei- oder Webinhalte (Prompt Injection) gesteuert wird, keine beliebigen Dateien des Benutzers lesen oder überschreiben. Eine Tabelle außerhalb des Arbeitsbereichs zuerst im Sheets-Editor öffnen; geöffnete Dateien bearbeiten die Werkzeuge live im Editor.

Andere Extensions können die Werkzeuge über den Befehl `sheets.executeTool` ohne eigene Bestätigung aufrufen. Es gelten dieselben Grenzen für ungeöffnete Dateien.

## Was geprüft wird

- `npm run check:package`: Manifest und Registrierung aller Werkzeuge, getrennte Bestätigungen für Lese-/Schreibwerkzeuge, Abbruch vor einer Mutation, Weiterleitung durch den Webview und Behandlung von Werkzeugfehlern; dafür wird die VS-Code-API nachgebildet.
- `npm run test:host`: echter VS-Code-Extension-Host mit separatem Profil. Prüft Sichtbarkeit in `vscode.lm.tools`, `vscode.lm.invokeTool` für gespeicherte und im Editor geöffnete XLSX-Dateien sowie die Zurückweisung ungültiger Eingaben. Es wird kein Sprachmodell aufgerufen und kein Copilot-Chat versendet.

Der Host-Test verwendet das lokal installierte VS Code. Auf macOS wird der Standardpfad verwendet; unter Linux wird `code` im PATH erwartet. Unter Windows wird `code.cmd` im PATH gesucht und die benachbarte `Code.exe` ohne Shell gestartet. Über `VSCODE_EXECUTABLE` lässt sich ein anderer CLI-Pfad oder unter Windows der direkte Pfad zur `Code.exe` angeben. Ein grafischer Desktop ist erforderlich.

Das prüft die technische Werkzeugstrecke, nicht die Auswahl eines bestimmten Werkzeugs durch ein bestimmtes Copilot-Modell. Für einen vollständigen Copilot-Chat-Test müssen die Werkzeuge in eurem angemeldeten Chat aktiviert sein.

## Fehlerbehebung

- **Werkzeuge fehlen:** aktuelle Version installieren, Agent Mode wählen und die Werkzeugauswahl kontrollieren. Copilot-Berechtigungen und Organisationsrichtlinien prüfen.
- **Keine Datei geöffnet:** einen vollständigen Dateipfad angeben oder die Datei zuerst im Sheets-Editor öffnen.
- **„Only spreadsheets inside the workspace …“:** die Datei liegt außerhalb des Arbeitsbereichs oder ist ein symbolischer Link dorthin. Die Datei im Sheets-Editor öffnen oder ihren Ordner zum Arbeitsbereich hinzufügen.
- **Schreibzugriff auf ungeöffnete Datei verweigert:** Workspace Trust und `sheets.agentTools.headlessWrites` kontrollieren.
- **Remote-Pfad mit symbolischem Link:** Remote-Dateisysteme können Linkziele nicht zuverlässig auflösen. Ungeöffnete Dateien werden deshalb auch dann abgewiesen, wenn ein übergeordnetes Verzeichnis ein Link ist oder der Provider bei der Prüfung einen Fehler meldet. Die Datei zuerst im Sheets-Editor öffnen. Schreibgeschützte Dateisysteme erlauben auch im geöffneten Editor keine Agentenänderungen.
- **Führende Nullen verschwinden:** beim Schreiben nicht-formelhafter Texte `parse:false` setzen. Der Standard ist Eingabeinterpretation; ein Wert wie `0004711` kann damit numerisch werden.

Quellen: [Language Model Tool API](https://code.visualstudio.com/api/extension-guides/ai/tools), [Extension-Werkzeuge in Agent Mode seit VS Code 1.99](https://code.visualstudio.com/updates/v1_99#_vs-code-extension-tools-in-agent-mode), [Werkzeuge in AI-Agents](https://code.visualstudio.com/docs/copilot/concepts/tools).
