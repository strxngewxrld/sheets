# XLSX-Kompatibilität und Schutz vor Datenverlust

## Speicherprinzip

Beim Import behält Sheets die vollständige ursprüngliche XLSX-Datei samt einem normalisierten Vergleichsmodell. Diese Informationen bleiben auch in JSON-Snapshots, Undo-Zuständen und den VS-Code-Wiederherstellungssicherungen erhalten.

Ohne fachliche Änderung liefert der Writer die ursprünglichen Dateibytes zurück. Bei unterstützten Änderungen werden die betroffenen Zellwerte, Formeln, Größenattribute oder Style-Referenzen gezielt im ursprünglichen XML geändert. Alle anderen ZIP-Einträge werden mit ihrem ursprünglichen Inhalt übernommen. Vorhandene Relationships und Content Types bleiben erhalten; Style-Ressourcen werden angehängt, sodass ursprüngliche Indizes gültig bleiben.

Ein vollständiger Neuaufbau ist nur zulässig, wenn das Original nachweislich exakt dem Output unseres Writers für das Importmodell entspricht (abgesehen von Erstellungs-/Änderungszeitpunkten) und die Formeln verstanden werden. Andernfalls scheitern Struktur-, Sortier-, Tabellen-, Diagramm- und Regeländerungen mit `XLSX-Schutz`. Es gibt keinen Schalter zum Erzwingen eines verlustbehafteten XLSX-Exports.

## Aktueller Umfang

| Bestandteil | Verhalten beim Speichern |
| --- | --- |
| Bilder, Zeichnungen, Kommentare, VML | Originalteile und Verknüpfungen bleiben erhalten. Notizen (Kommentare) und Hyperlinks werden gelesen und angezeigt, aber nicht bearbeitet. |
| Native Pivot-Definitionen und Pivot-Caches | Originalteile bleiben erhalten. Wertänderungen in erkannten Pivot-Ausgabebereichen werden blockiert. |
| Datenüberprüfung | Wird gelesen und bearbeitet. Geänderte Regeln ersetzen gezielt nur das `<dataValidations>`-Element. Enthält das Blatt zusätzlich Excel-2010-Regeln in `extLst`, werden Änderungen blockiert. |
| Erweiterte bedingte Formate | Datenbalken und Symbolsätze werden gelesen und dargestellt. Unverändertes XML bleibt erhalten; Änderungen der Regeldefinitionen sind bei nicht vollständig modellierten Dateien blockiert. |
| Makros (`.xlsm`) | `vbaProject.bin` und alle übrigen Teile bleiben erhalten; Makros werden nie ausgeführt. Neue `.xlsm`-Dateien erhalten den makrofähigen Content-Type. |
| Native Tabellen und Excel-Tabellenstile | Originaldefinitionen bleiben erhalten. Änderungen von Kopfzeilen, berechneten Spalten, Summenzeilen oder Tabellen-Metadaten werden bei nicht sicher übertragbaren Dateien blockiert. |
| Unbekannte Formeln, Shared/Array-/Datentabellen-Metadaten | Unveränderte Formeln und ihre gespeicherten Ergebnisse bleiben erhalten. Änderungen komplexer Formelstrukturen und ihrer abhängigen Ausgabebereiche werden blockiert. |
| Berechnungskette | Bleibt erhalten. Änderungen am Formelgraphen mit vorhandener Kette werden blockiert. |
| Nicht dargestellte Zellformatierung | Beispielsweise Textrotation, Zeichensatz, Hochstellung, diagonale Rahmen und Schutzattribute bleiben beim Ändern anderer Formatanteile erhalten. |
| Rich Text und phonetische Annotationen | Bleiben erhalten. Textänderungen, die diese Inhalte flachschreiben würden, werden blockiert; Zellformatierung ist weiter möglich. |
| Spaltenbreiten und Zeilenhöhen | Größenattribute werden gezielt geändert. Andere Attribute wie Gliederung, Zell-Style und nicht explizit geänderte Ausblendungen bleiben erhalten. |
| Externe Links, Verbindungen, Custom XML, eingebettete Binärteile | Bleiben erhalten; werden nicht ausgeführt oder automatisch aktualisiert. |
| Versteckte Blätter, lokale Namen, Druckeinstellungen und Dokumenteigenschaften | Bleiben im Originalpaket erhalten, auch wenn das Modell sie nicht vollständig darstellt. |
| Digitale Signaturen | Unverändertes Speichern ist möglich. Änderungen werden blockiert, um Signaturen nicht stillschweigend ungültig zu machen. |
| Geschützte Blätter, Datumssystem 1904, unbekannte Zellmetadaten | Nicht sicher unterstützte Änderungen werden blockiert. Die Originaldatei bleibt bestehen. |

Die Neuberechnung im Editor unterstützt weiterhin nur die implementierten Formeln. Unveränderte Formel-Caches im Dateipaket werden beibehalten; geänderte Formeln erhalten keinen erfundenen Cache. Die Datei wird für eine vollständige Neuberechnung durch Excel markiert. Andere Programme, die ausschließlich gespeicherte Ergebnisse lesen und keine Berechnung durchführen, können bis dahin alte oder fehlende Formel-Ergebnisse sehen.

## Speichern, Wiederherstellen und Konflikte

Der Editor prüft nach jeder Änderung im Hintergrund, ob sie sich noch verlustfrei speichern lässt, und zeigt andernfalls den Grund über der Tabelle an. Eine Fehlermeldung des Writers tritt auf, bevor die Extension oder der MCP-Server die Zieldatei schreibt. Das gilt auch für Agent-Werkzeuge. „Speichern unter“ umgeht den XLSX-Schutz nicht.

VS Code legt Wiederherstellungssicherungen als separates JSON-Format ab. Dadurch bleiben ungespeicherte Änderungen samt Originalarchiv erhalten, selbst wenn sie sich derzeit nicht verlustfrei als XLSX speichern lassen. Nach dem Wiederherstellen gelten dieselben Schutzprüfungen. Alte Sicherungen ohne nachweisbare Dateiversion können nur in eine neue Datei gespeichert werden.

Vor dem Überschreiben wird außerdem geprüft, ob die Datei seit dem Einlesen beziehungsweise letzten Speichern extern verändert wurde. Erkannte Konflikte werden verweigert. Das ist keine Betriebssystem-Sperre gegen zeitgleiche Schreibzugriffe anderer Programme.

CSV/TSV sind weiterhin ausdrücklich einfache Werteformate: Sie können weder mehrere Blätter noch Excel-Metadaten tragen. Ein bewusster CSV-Export ist eine Formatkonvertierung; der XLSX-Erhaltungspfad betrifft XLSX-Dateien.

## Prüfungen

- Bytegenaues No-op-Speichern und Erhaltung aller Paketbestandteile bei Zell- und Formatänderungen.
- Serialisierung über JSON sowie VS-Code-Backup und Wiederherstellung.
- Kein Dateischreiben bei Schutzfehlern und erkannten externen Änderungen.
- Unbekannte Formeln, Matrixbereiche, Rich Text, numerische Originalpräzision, Datumsmodus 1904 und Signaturen.
- Append-only Styles mit unbekannten Schrift-, Rahmen-, Ausrichtungs- und Schutzattributen.
- Eine unabhängige openpyxl-Testdatei in `packages/core/test/fixtures/excel-features.xlsx` mit Bildern, Kommentaren, Verbindungen zwischen Teilen, Datenüberprüfung, Data Bars, Diagramm, lokalen Namen, Druckeinstellungen und ausgeblendeten Inhalten.
- Zusätzlich wurde die Microsoft-Datei „Financial Sample.xlsx“ geprüft: Die 26 ursprünglichen Teile bleiben erhalten; bei einer Zell-/Formatänderung werden ausschließlich Blatt, Style-Ergänzungen und Berechnungskennzeichen geändert. Der native Stil `TableStyleLight9` bleibt bestehen.

Das ist ein erhaltender Speicherpfad mit konservativen Schutzgrenzen, keine vollständige Implementierung aller Excel-Funktionen. Die beschriebenen Prüfungen ersetzen keine Zertifizierung aller Excel-Dateivarianten.

Technischer Hintergrund: [Aufbau eines SpreadsheetML-Dokuments](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/structure-of-a-spreadsheetml-document), [Berechnungskette](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-the-calculation-chain).
