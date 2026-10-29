# Sheets-Handbuch

[English](manual.en.md) · [README](../README.de.md)

Dieses Handbuch beschreibt die Sheets-Extension für VS Code Desktop unter Windows, macOS und Linux. Die Befehlsnamen folgen der deutschen Oberfläche. Mit **Sheets: Handbuch öffnen** lässt sich diese Anleitung jederzeit aufrufen; **Sheets: Erste Schritte** öffnet die geführte Einführung. Bei englischer VS-Code-Sprache heißen die Befehle **Open User Manual** und **Getting Started**.

## Inhalt

- [Installation und Dateien öffnen](#installation-und-dateien-öffnen)
- [Die erste Arbeitsmappe](#die-erste-arbeitsmappe)
- [Navigation und Bearbeitung](#navigation-und-bearbeitung)
- [Formeln](#formeln)
- [Formatierung und Tabellen](#formatierung-und-tabellen)
- [Suchen und Ersetzen](#suchen-und-ersetzen)
- [Diagramme und Auswertungen](#diagramme-und-auswertungen)
- [Speichern und Exportieren](#speichern-und-exportieren)
- [KI-Unterstützung](#ki-unterstützung)
- [Einstellungen und Fehlerbehebung](#einstellungen-und-fehlerbehebung)

## Installation und Dateien öffnen

![Erste-Schritte-Walkthrough in VS Code mit englischer Oberfläche](images/walkthrough.png)

*Die Einführung mit fünf Schritten, aufgenommen mit Windows MCP. Die Texte folgen der VS-Code-Sprache.*

Die Release-VSIX über **Extensions: Install from VSIX…** installieren und VS Code bei Aufforderung neu laden. Erforderlich ist VS Code ab Version 1.95; Excel wird nicht benötigt. Eine Anleitung zum eigenen Bauen steht in der [README](../README.de.md#installation).

`.xlsx` und `.xlsm` im Explorer öffnen. Bei `.csv` und `.tsv` per Rechtsklick **Mit Sheets öffnen** oder über **Öffnen mit… → Sheets (CSV)** wählen. **Sheets: Tabellendatei öffnen…** bietet eine Dateiauswahl für alle unterstützten Formate. Alte `.xls`-Dateien müssen zuerst in einer anderen Tabellenkalkulation konvertiert werden.

Die Befehlspalette öffnet sich mit **Strg+Umschalt+P** unter Windows/Linux bzw. **Cmd+Umschalt+P** unter macOS. Der Walkthrough ist auch auf der VS-Code-Seite „Erste Schritte“ verfügbar. Manuelles Bearbeiten erfordert weder ein KI-Konto noch einen vertrauenswürdigen Arbeitsbereich.

## Die erste Arbeitsmappe

1. **Sheets: Neue Arbeitsmappe** ausführen und einen Dateinamen mit `.xlsx` wählen. Der Befehl erstellt und öffnet die Datei.
2. A1 auswählen und folgenden tabulatorgetrennten Text einfügen:

   ```text
   Product	Quantity	Unit price	Revenue
   Notebook	3	12.5	=B2*C2
   Pen	8	2	=B3*C3
   Folder	2	4.5	=B4*C4
   Total			=SUM(D2:D4)
   ```

3. D2:D4 enthalten **37,5**, **16** und **9**; D5 enthält **62,5**. Die angezeigten Dezimalzeichen hängen von der Formatierung ab.
4. A1:D1 markieren und über die Werkzeugleiste fett formatieren. C2:D5 markieren, mit **Strg+1** / **Cmd+1** die Zellformatierung öffnen und ein Zahlen- oder Währungsformat mit zwei Nachkommastellen wählen.
5. Mit **Strg+S** / **Cmd+S** speichern.

Alternativ enthält [getting-started.csv](examples/getting-started.csv) die drei Eingabespalten. Für die vollständige Übung XLSX verwenden, damit Formeln, Formate und Diagramme erhalten bleiben.

![Sheets-Editor mit einer Umsatz-Arbeitsmappe](images/editor.png)

*Echter Screenshot aus VS Code Desktop, aufgenommen mit Windows MCP. Die Screenshot-Daten zeigen die Oberfläche; für die Übung gelten die Werte oben.*

## Navigation und Bearbeitung

Eine Zelle anklicken und tippen, um den Wert zu ersetzen. Per Doppelklick den vorhandenen Inhalt bearbeiten oder die Formelleiste verwenden. Enter bestätigt, Escape bricht eine Eingabe ab. Einen Bereich durch Ziehen markieren. Über das Namensfeld oder **Sheets: Gehe zu Zelle…** zu einer Zelle, einem Bereich oder definierten Namen springen, beispielsweise `D5` oder `A1:D5`.

Rechteckige Daten mit den üblichen Tastenkürzeln kopieren und einfügen. **Nur Werte einfügen** übernimmt Ergebnisse ohne Formeln; **Inhalte einfügen…** bietet Formeln, Formate, das Überspringen leerer Zellen und Transponieren. Mit dem Ausfüllkästchen benachbarte Zellen füllen. Kontextmenüs der Zeilen- und Spaltenköpfe bieten Einfügen, Löschen und Ausblenden. Über die Blattregister und deren Kontextmenüs Blätter wechseln, umbenennen und verwalten. Der XLSX-Speicherschutz kann einzelne Änderungen einschränken.

Über **Fenster fixieren** in der Werkzeugleiste Zeilen und Spalten fixieren. Den Zoom über die Statusleiste oder **Sheets: Vergrößern**, **Verkleinern** und **Zoom zurücksetzen** ändern. Unterstützt werden 25–400 %.

| Aktion | Windows / Linux | macOS |
| --- | --- | --- |
| Speichern | Strg+S | Cmd+S |
| Rückgängig / Wiederholen | Strg+Z / Strg+Umschalt+Z | Cmd+Z / Cmd+Umschalt+Z |
| Kopieren / Ausschneiden / Einfügen | Strg+C / Strg+X / Strg+V | Cmd+C / Cmd+X / Cmd+V |
| Nur Werte einfügen | Strg+Umschalt+V | Cmd+Umschalt+V |
| Inhalte einfügen | Strg+Alt+V | Cmd+Alt+V |
| Zellen formatieren | Strg+1 | Cmd+1 |
| Suchen | Strg+F | Cmd+F |
| Ersetzen | Strg+H | Cmd+Option+F |
| Gehe zu | Strg+G | Cmd+G |

Die Kürzel gelten bei Fokus im Tabelleneditor. Falls das Betriebssystem ein Kürzel abfängt, sind die Sheets-Befehle für Navigation, Suche und Zoom auch in der Befehlspalette erreichbar.

## Formeln

Formeln beginnen mit `=`. Auch in der deutschen Oberfläche **englische Funktionsnamen, Kommas zwischen Argumenten und einen Dezimalpunkt in Formelliteralen** verwenden. Beispiel: `=IF(B2>2,SUM(D2:D4),0)`. Das Semikolon dient nicht als Argumenttrenner. Die Anzeige kann trotzdem lokalisierte Zahlenformate verwenden.

`B2` ist relativ; `$B$2` fixiert Zeile und Spalte. Beim Ausfüllen von `=B2*C2` nach unten wird daraus `=B3*C3`. Andere Blätter über `='Sales 2026'!D2` referenzieren. Bei der Eingabe helfen Vorschläge und farbig markierte Bezüge. Unterstützt werden unter anderem Suchfunktionen, dynamische Arrays und `LET`/`LAMBDA`, jedoch nicht jede Excel-Funktion.

Bei `#NAME?` Funktionsschreibweise und Unterstützung prüfen. Bei `#REF!` die referenzierten Zellen und Blätter kontrollieren; bei `#DIV/0!` den Nenner prüfen. Bei unbekannten Excel-Formeln kann der gespeicherte Ergebniswert der Quelldatei erhalten bleiben. Sheets berechnet keine beliebigen Excel-Funktionen und führt keine Makros aus.

## Formatierung und Tabellen

Den Zielbereich markieren und über Werkzeugleiste oder **Zellen formatieren…** Schrift, Farben, Ausrichtung, Rahmen und Zahlenformate ändern. Ein Format verändert die Darstellung, nicht den zugrunde liegenden Zahlenwert.

Für die drei Produktdatensätze der Übung **A1:D4** ohne Summenzeile markieren und **Als Tabelle formatieren…** wählen. Der Tabelle einen Namen geben und bestätigen, dass die erste Zeile Überschriften enthält. Im Filtermenü einer Überschrift Werte suchen und die anzuzeigenden Werte auswählen. Den Spaltenfilter löschen, um wieder alle Datensätze zu sehen. Filtern blendet Datensätze aus und löscht sie nicht.

Über das Tabellenmenü oder **Sortieren…** vollständige Datensätze nach einer oder mehreren Spalten sortieren. Zusammengehörige Spalten gemeinsam auswählen. Strukturierte Bezüge wie `Table1[Revenue]` und `[@Quantity]` verwenden den tatsächlichen Tabellen- und Spaltennamen.

## Suchen und Ersetzen

**Sheets: Suchen** oder **Sheets: Suchen und Ersetzen** öffnen. Suchtext eingeben und das aktuelle Blatt oder die ganze Arbeitsmappe sowie Werte oder Formeln als Suchziel wählen. Bei Bedarf Groß-/Kleinschreibung, vollständige Zelltreffer oder reguläre Ausdrücke aktivieren. Treffer vor dem Ersetzen prüfen oder alle Treffer im gewählten Suchbereich ersetzen. Escape schließt den Suchbereich.

![Suchbereich in Sheets](images/find.png)

*Suchbereich in VS Code Desktop, aufgenommen mit Windows MCP.*

## Diagramme und Auswertungen

Einen Bereich mit Beschriftungen und Zahlen markieren und **Diagramm einfügen** in der Werkzeugleiste wählen. Typ auswählen und Quellbereich sowie Titel in den Diagramm-Steuerelementen konfigurieren. Für die Übung zeigt A1:B4 Mengen pro Produkt. Diagramme lassen sich verschieben und skalieren. Zum Erhalten der Diagramme als XLSX speichern.

Über **Bedingte Formatierung** Vergleichs- und Formelregeln, Farbskalen, Datenbalken oder Symbolsätze auf einen Bereich anwenden. Formelregeln relativ zur linken oberen Zelle des Zielbereichs formulieren.

Mit **Datenüberprüfung** Dropdown-Listen und Beschränkungen für Zahlen, Datumswerte, Textlängen oder eigene Formeln festlegen. Zielbereich und Regel angeben sowie bei Bedarf Eingabe- und Fehlermeldungen ergänzen.

Für eine Zusammenfassung **Pivot-Auswertung** öffnen und Quellbereich, Zeilen-/Spaltengruppen, Wertfeld, Aggregation und Zielzelle wählen. Nach Änderungen an der Quelle **Pivot-Auswertungen aktualisieren** verwenden. Sheets-Auswertungen sind von nativen Excel-Pivot-Caches zu unterscheiden; Änderungen an deren Ausgabebereichen können zum Schutz der Originaldatei blockiert werden.

Notizen importierter Arbeitsmappen werden angezeigt, lassen sich aber nicht bearbeiten. Hyperlinks und Ergebnisse von `HYPERLINK` per Strg/Cmd+Klick öffnen.

## Speichern und Exportieren

**Strg+S** / **Cmd+S** speichert das offene Dokument; Rückgängig/Wiederholen ist in VS Code eingebunden. **Speichern unter** verwendet das Format der Zieldatei. Für eine separate Wertekopie **Sheets: Blatt als CSV exportieren…** ausführen, Komma, Semikolon oder Tabulator wählen und das Ziel festlegen. Der Export schreibt das aktive Blatt und ersetzt die Quellmappe nicht.

Beim Export zuerst den Verwendungszweck wählen: **Für Excel (Textschutz)** stellt Texten mit möglichen Formelpräfixen einen Tabulator innerhalb eines zitierten CSV-Felds voran. Das verändert die exportierten Textwerte; Zahlen bleiben Zahlen. **Unveränderte Werte** ist für Datenimporte gedacht und übernimmt Texte wörtlich, die Excel möglicherweise als Formel interpretiert. Auch neue CSV-/TSV-Konvertierungen über „Speichern unter“ verwenden den Textschutz. Das normale Speichern einer bestehenden CSV bleibt verlustfrei. Für die unveränderte Weitergabe von Texten an Tabellenprogramme eignet sich XLSX besser, da es Text- und Formelzellen unterscheidet; CSV-Schutz hängt vom Zielprogramm ab.

CSV-/TSV-Ausgaben sind auf eine Million Positionen und 16 MiB UTF-8-Ausgabe begrenzt. Führende Leerzeilen und Leerspalten zählen mit. Zu große Ausgaben werden vor dem Schreiben mit einer Fehlermeldung abgelehnt; die Mappe bleibt geöffnet und kann als XLSX gespeichert werden.

| Format | Gespeicherter Inhalt |
| --- | --- |
| XLSX | Mehrere Blätter, Formeln, unterstützte Formatierung, Tabellen und Diagramme. Unbekannte Originalteile bleiben erhalten, soweit sicher möglich. |
| XLSM | Makrobestandteile bleiben erhalten; Makros werden nie ausgeführt. |
| CSV / TSV | Nur Werte eines Datenblatts, ohne Arbeitsmappenformatierung, Diagramme und weitere Blätter. |

Beim CSV-Speichern bleiben erkanntes Trennzeichen, Zeilenenden, UTF-8-BOM und unterstützte Quellkodierung (UTF-8 oder Windows-1252) erhalten. Unveränderte Datensätze werden nach Möglichkeit wörtlich übernommen. Kennungen mit führenden Nullen bleiben Text. Lassen sich neue Zeichen nicht in Windows-1252 darstellen, wechselt Sheets mit einem Hinweis zu UTF-8 mit BOM.

Bei einer **XLSX-Schutzmeldung** über der Tabelle lässt sich eine Änderung nicht sicher speichern, ohne unbekannte Excel-Inhalte zu verlieren. Die betreffende Änderung rückgängig machen oder benötigte Werte separat als CSV exportieren. „Speichern unter“ umgeht den Schutz nicht. Unveränderte unbekannte Inhalte bleiben im Originalpaket. Wiederherstellungssicherungen bewahren offene Änderungen, heben Speichereinschränkungen aber nicht auf. Details: [XLSX-Kompatibilität und Speicherschutz](xlsx-compatibility.md).

Wurde die Quelldatei extern geändert, verweigert Sheets das Überschreiben. Offene Werte separat sichern, die aktuelle Datei erneut öffnen und die Änderungen abgleichen.

## KI-Unterstützung

Eine aktuelle VS-Code-Version mit einem kompatiblen KI-Client wie Copilot im Agent-Modus verwenden (Extension-Werkzeuge erfordern VS Code ab 1.99). Im Werkzeugmenü des Clients die `sheets_*`-Werkzeuge aktivieren. Beispiel: „Lies A1:D5 in der offenen Mappe und erkläre die Berechnung der Summe.“ Oder: „Formatiere C2:D5 mit zwei Nachkommastellen.“ Änderungen prüfen und die geöffnete Mappe selbst speichern.

Änderungen an offenen Dokumenten erscheinen direkt und lassen sich rückgängig machen. Bei aktivem `sheets.agentTools.headlessWrites` dürfen Agents ungeöffnete Dateien direkt ändern und speichern; das erfordert einen vertrauenswürdigen Arbeitsbereich. Die Einstellung deaktivieren, wenn Dateien vor Agent-Änderungen geöffnet werden sollen. Sheets verarbeitet Tabellen lokal; der KI-Client kann Werkzeugeingaben und Ergebnisse an seinen Dienst übertragen. Weitere Beispiele: [Copilot-Anleitung](copilot.md).

## Einstellungen und Fehlerbehebung

| Einstellung | Standard | Wirkung |
| --- | --- | --- |
| `sheets.checkForUpdates` | `true` | Prüft beim Start öffentliche GitHub Releases und bietet Updates an. |
| `sheets.agentTools.headlessWrites` | `true` | Agents dürfen ungeöffnete Dateien in vertrauenswürdigen Arbeitsbereichen speichern. |

**Eine Mappe öffnet sich als Binärtext oder in einer anderen Extension:** Fenster neu laden, Tab schließen und erneut öffnen. Über **Öffnen mit… → Sheets** den Editor wählen und bei Bedarf als Standard für `*.xlsx` festlegen. Vorhandene Benutzer-/Workspace-Zuordnungen haben Vorrang.

**CSV erscheint als Text:** Zunächst **Mit Sheets öffnen** oder **Sheets (CSV)** in der Editorauswahl wählen.

**Nach dem CSV-Speichern fehlt eine Formel oder ein Diagramm:** CSV speichert ausschließlich Werte. Für solche Inhalte XLSX verwenden.

**Ein Sheets-Befehl reagiert nicht:** Zuerst eine geöffnete Sheets-Tabelle fokussieren. Suche, Zoom und Export arbeiten mit dem aktiven Sheets-Dokument.

**Walkthrough oder Handbuch fehlen:** Die aktualisierte VSIX installieren, VS Code neu laden und **Sheets: Erste Schritte** oder **Sheets: Handbuch öffnen** ausführen.
