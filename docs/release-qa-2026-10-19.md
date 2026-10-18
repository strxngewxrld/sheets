# Lokale Release-Prüfung vom 19. Oktober 2026

Geprüft: Sheets 1.0.2 im lokalen Extension Development Host, macOS 15.7.5 (Intel/x64), VS Code 1.142.0, Node.js 26.12.0. Ausgangspunkt war ein sauberes Arbeitsverzeichnis. Verwendet wurden erzeugte Testdateien und isolierte VS-Code-Profile.

## Gefundene und behobene UI-Probleme

- Das Kontextmenü war ungefähr 705 px hoch und ragte in einem 480 × 320 px großen Editorbereich unten heraus. Menüs bekommen jetzt eine begrenzte, scrollbar erreichbare Fläche.
- Menüs hatten keine Pfeiltasten-Navigation. ArrowUp/Down, Home/End, das Überspringen deaktivierter Einträge und zyklische Navigation funktionieren jetzt. Escape und Menüaktionen geben den Fokus zurück; Tab schließt das Menü.
- Beim Schließen über eine Menüaktion blieben globale Ereignislistener zurück. Jede Schließart entfernt jetzt Listener, ausstehende Timer und den Größenbeobachter. Wiederholtes Öffnen und Ausführen wird überprüft.
- Dialoge waren in einer 260 px breiten Editorgruppe links um 36 px abgeschnitten. Dialoge und Felder passen sich jetzt der verfügbaren Breite an; Aktionsbuttons bleiben scrollbar erreichbar.
- Menüs verwendeten feste Bildschirmpositionierung, obwohl ihre Koordinaten relativ zum Tabelleneditor berechnet wurden. Sie sind jetzt relativ zum Editor verankert und schließen sich bei Größenänderungen der Editorgruppe.

## Ausgeführte Prüfungen

| Prüfung | Ergebnis |
| --- | --- |
| `npm test` | 825 Tests in 39 Dateien bestanden |
| `npm run test:host` | Echte Tool-Registry und Editor-Lebenszyklus bestanden |
| `npm run test:ui` | 14 Szenarien in der gerenderten Custom-Editor-Webview bestanden |
| `npm run package` | Build, Typecheck, VSIX-Inhalt und Aktivierung des gepackten Bundles bestanden |
| macOS-MCP | Desktop und Development Host inspiziert; Cmd+F und Suche mit „Müller“ bedient und Treffer A2 in der Webview verifiziert |

Der Host-Test prüft die Registrierung aller 45 Werkzeuge, tatsächliche Werkzeugaufrufe, ungültige Eingaben, das automatische Öffnen von XLSX mit Sheets, Änderungen im Live-Editor, Speichern und Wiederöffnen einschließlich nativer Tabellen. Er prüft außerdem Windows-1252-CSV mit Dezimalkomma, Semikolon und CRLF sowie den Schutz beschädigter XLSX-Dateien vor Überschreiben.

Die 14 UI-Szenarien decken ab:

1. XLSX laden, Unicode und Formeln in der Formelleiste.
2. Suchen, Ersetzen, ungültige reguläre Ausdrücke und Rückgängig über den Host.
3. Zellbearbeitung, Formeln, rechteckiges Einfügen, Rückgängig und Wiederholen.
4. Ungültige Eingaben durch Datenüberprüfung zurückweisen, ohne die Zelle zu verändern.
5. Kontextmenü in kleinen Editorflächen, Scrollen bis zum letzten Eintrag und deaktivierte Menüeinträge.
6. Menütastatur, zyklische Navigation und Fokus nach Escape sowie Menüaktionen.
7. Wiederholte Menüaktionen ohne verbleibende globale Listener, versetzte Editorposition und Größenänderungen.
8. Format-, Tabellen-, Sortier-, bedingte Formatierungs-, Pivot- und Suchdialoge in schmalen Editorgruppen; fehlerhafte Bereichseingaben.
9. Tabellenformatierung, Formelwerte, Speichern und Wiederöffnen.
10. Wertefiltersuche und Filter anwenden, Auswahlstatistik und Zoom.
11. Diagramme einfügen, alle sieben Typen rendern, Diagrammeditor und Rückgängig.
12. Blätter hinzufügen und umbenennen, Rückgängig und Wiederholen.
13. Dark Theme und Dialogdarstellung; Light Theme wird zuvor geprüft.
14. CSV im Custom Editor bearbeiten und speichern, ohne Semikolon oder CRLF zu verändern.

Screenshots und das maschinenlesbare UI-Ergebnis liegen nach dem Lauf unter [artifacts/release-qa/final/](../artifacts/release-qa/final/). Bei einer Wiederholung gibt `npm run test:ui` den neuen temporären Artefaktpfad aus. Die UI-Automatisierung arbeitet über Chromium DevTools in der echten VS-Code-Webview; sie ersetzt keine vollständige Prüfung nativer Systemdialoge mit Maus und Tastatur.

## Aussagegrenzen

Die lokalen Prüfungen ergeben keine allgemeine Fehlerfreiheitsgarantie. Windows, Linux, VS Code 1.95 als Mindestversion, installierte Drittanbieter-Extensions, sehr große reale Arbeitsmappen, native Office-Gegenprüfungen, Screenreader und sämtliche nativen Datei-/Speicherdialoge wurden in dieser Sitzung nicht vollständig getestet. Die 45 Werkzeuge wurden auf Registrierung geprüft; nicht jede mögliche Eingabe jedes Werkzeugs wurde im echten Host ausgeführt. Deutschsprachige Dialogtexte wurden in dieser UI-Automatisierung nicht separat durchlaufen.

Es wurde keine Version erhöht und kein Release veröffentlicht. Das neu erzeugte lokale VSIX trägt weiterhin die vorhandene Version 1.0.2.

SHA-256 des geprüften VSIX: `abde6f4bc165c7abeedab589b90f6a4ee8d4dfd601a1ece652fafa744adb9668`.
