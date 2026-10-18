# XLSX-Kompatibilität und Schutz vor Datenverlust

## Speicherprinzip

Beim Import behält Sheets die vollständige ursprüngliche XLSX-Datei samt einem normalisierten Vergleichsmodell. Diese Informationen bleiben auch in JSON-Snapshots, Undo-Zuständen und den VS-Code-Wiederherstellungssicherungen erhalten.

Ohne fachliche Änderung liefert der Writer die ursprünglichen Dateibytes zurück. Bei unterstützten Änderungen werden die betroffenen Zellwerte, Formeln, Größenattribute oder Style-Referenzen gezielt im ursprünglichen XML geändert. Alle anderen ZIP-Einträge werden mit ihrem ursprünglichen Inhalt übernommen. Vorhandene Relationships und Content Types bleiben erhalten; Style-Ressourcen werden angehängt, sodass ursprüngliche Indizes gültig bleiben.

Ein vollständiger Neuaufbau ist nur zulässig, wenn das Original nachweislich exakt dem Output unseres Writers für das Importmodell entspricht (abgesehen von Erstellungs-/Änderungszeitpunkten) und die Formeln verstanden werden. Andernfalls scheitern Struktur-, Tabellen-, Diagramm- und Regeländerungen mit `XLSX-Schutz`; für Sortieren, neue Blätter und fehlende Style-Tabellen gelten die gezielten Pfade unten. Es gibt keinen Schalter zum Erzwingen eines verlustbehafteten XLSX-Exports.

## Gezielt erlaubte Änderungen in beliebigen Dateien

Diese Änderungen werden direkt in das erhaltene Originalpaket eingefügt. Tests belegen jeweils, dass alle übrigen ZIP-Einträge byte-identisch bleiben (eigene Testpakete und die openpyxl-Datei `excel-features.xlsx`).

| Änderung | Was geschrieben wird | Bleibt blockiert |
| --- | --- | --- |
| **Sortieren** eines Zellbereichs | Nur die Zellen in `sheetData`, deren Inhalt sich ändert (wie eine Zelländerung), plus das Neuberechnungskennzeichen in `workbook.xml`. Jede Zelle erhält exakt den ursprünglichen Formatdatensatz (`s`-Index) der Zelle, deren Darstellung sie übernimmt, sodass auch nicht dargestellte Formatanteile (Drehung, Hochstellung, Schutz) mitwandern. Formeln, die nur auf die eigene Zeile verweisen (`=B2*2`), bleiben dadurch textgleich und unverändert. | Zellen mit Notizen/Kommentaren oder Hyperlinks, Bereiche in nativen Excel-Tabellen, mit Datenüberprüfung oder verbundenen Zellen; Formeln mit relativen Bezügen auf andere Zeilen oder unbekannten Funktionen; Zahlen, deren gespeicherter Text genauer ist als eine Gleitkommazahl (z. B. `9007199254740993`); gleich dargestellte Zellen mit unterschiedlichen Excel-Formatdatensätzen; zusätzlich alle Schutzregeln für Zelländerungen (Rich Text, Zellmetadaten, geteilte/Matrixformeln, Berechnungskette bei Formeländerungen, Pivot-Ausgaben, Datumssystem 1904, Blattschutz). Ausschneiden und Einfügen bleibt als Strukturänderung blockiert. |
| **Neues Blatt am Ende** | Ein neuer Teil `xl/worksheets/sheetN.xml` mit Zellen, Spaltenbreiten, Zeilenhöhen und Datenüberprüfung, ein zusätzlicher Eintrag in `<sheets>` (neue `sheetId`, neue Relationship-ID) sowie je ein angehängtes Element in `xl/_rels/workbook.xml.rels` und `[Content_Types].xml`. Bestehende Einträge, `definedNames` (inklusive `localSheetId`), Blattreihenfolge und alle anderen Teile bleiben unverändert. `docProps/app.xml` wird nicht angepasst (die Blattliste dort ist nur informativ). | Einfügen an anderer Position, Umbenennen, Löschen und Verschieben von Blättern; neue Blätter mit Diagrammen, Tabellen, bedingten Formaten, Notizen, verbundenen Zellen, fixierten Bereichen oder abweichender Standardgröße; Namen, die bereits ein (auch im Editor unsichtbares) Diagrammblatt trägt; Strict-Open-XML-Dateien. |
| **Formatierung ohne Style-Tabelle** | Fehlt `styles.xml` vollständig, wird eine minimale gültige Style-Tabelle neu angelegt (Relationship und Content Type werden angehängt) und Formate wie üblich angefügt. Unformatierte Zellen zeigen danach die Standardschrift Calibri 11. | Dateien, deren Style-Relationship auf einen fehlenden Teil zeigt, oder die bereits einen nicht verknüpften Teil `styles.xml` enthalten. |

Umbenennen und Löschen bleiben bewusst blockiert: Blattnamen stehen auch in Diagrammbezügen, definierten Namen, Pivot-Caches, Hyperlinks und Erweiterungen, und Löschen verschiebt `localSheetId`-Indizes und die Berechnungskette. Diese Stellen lassen sich nicht vollständig nachweisbar anpassen.

Die Gründe einer Schutzmeldung folgen der Sprache der Oberfläche (Deutsch oder Englisch, `setXlsxMessageLanguage`). Fehler sind `XlsxProtectionError` mit dem Code `XLSX_PROTECTION`; der Text beginnt mit `XLSX-Schutz:` bzw. `XLSX protection:`. MCP-Server und Agent-Werkzeuge melden standardmäßig Deutsch.

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

Die Neuberechnung im Editor unterstützt weiterhin nur die implementierten Formeln. Formel-Caches (das gespeicherte Ergebnis `<v>`) werden so behandelt:

- Ändert sich das Ergebnis einer unveränderten Formel durch eine Bearbeitung – auch auf einem anderen Blatt –, wird nur ihr `<v>` (und der Zelltyp `t`) auf das berechnete Ergebnis gesetzt; die Formel und alle übrigen Attribute bleiben bytegleich. Formeln, deren Ergebnis gleich bleibt, werden nicht angefasst.
- Neue und geänderte Formeln erhalten ihr berechnetes Ergebnis als Cache.
- Beides gilt nur, wenn die Engine alle Funktionen der Formel kennt. Formeln mit unbekannten Funktionen behalten ihren Cache bzw. erhalten keinen – ein Ergebnis wird nie erfunden. Ebenso unverändert bleiben Matrixformeln, Datentabellen, dynamische Matrizen, Zellen mit Metadaten, geschützte Blätter und Dateien mit Datumssystem 1904. Ein Cache-Abgleich führt nie zu einer Speichersperre.
- Die Datei wird zusätzlich für eine vollständige Neuberechnung durch Excel markiert.

Programme, die ausschließlich gespeicherte Ergebnisse lesen (Datenimporte, openpyxl mit `data_only`, pandas), sehen damit aktuelle Werte, außer bei den genannten Ausnahmen. Ergebnisse, die von Gebietsschema oder Programm abhängen (z. B. `TEXT` mit Dezimalzeichen), entsprechen der Berechnung im Editor, bis Excel neu rechnet.

Blattschutz verhindert Bearbeitungen des geschützten Blatts. Ändert eine Bearbeitung auf einem anderen Blatt nur die Ergebnisse von Formeln auf dem geschützten Blatt, wird gespeichert und das geschützte Blatt bleibt unverändert.

## Speichern, Wiederherstellen und Konflikte

Der Editor prüft nach jeder Änderung im Hintergrund, ob sie sich noch verlustfrei speichern lässt, und zeigt andernfalls den Grund über der Tabelle an. Eine Fehlermeldung des Writers tritt auf, bevor die Extension oder der MCP-Server die Zieldatei schreibt. Das gilt auch für Agent-Werkzeuge. „Speichern unter“ umgeht den XLSX-Schutz nicht.

VS Code legt Wiederherstellungssicherungen als separates JSON-Format ab. Dadurch bleiben ungespeicherte Änderungen samt Originalarchiv erhalten, selbst wenn sie sich derzeit nicht verlustfrei als XLSX speichern lassen. Nach dem Wiederherstellen gelten dieselben Schutzprüfungen. Alte Sicherungen ohne nachweisbare Dateiversion können nur in eine neue Datei gespeichert werden.

Vor dem Überschreiben wird außerdem geprüft, ob die Datei seit dem Einlesen beziehungsweise letzten Speichern extern verändert wurde. Erkannte Konflikte werden verweigert. Das ist keine Betriebssystem-Sperre gegen zeitgleiche Schreibzugriffe anderer Programme.

CSV/TSV sind weiterhin ausdrücklich einfache Werteformate: Sie können weder mehrere Blätter noch Excel-Metadaten tragen. Ein bewusster CSV-Export ist eine Formatkonvertierung; der XLSX-Erhaltungspfad betrifft XLSX-Dateien. Enthält ein CSV-Dokument beim Speichern Inhalte, die CSV nicht aufnehmen kann und die beim Öffnen noch nicht vorhanden waren (weitere Blätter, Diagramme, Formeln, Formatierungen, Tabellen, Regeln, Notizen), zeigt der Editor einmal pro Dokument einen nicht blockierenden Hinweis, dass nur die Werte des Datenblatts gespeichert werden. „Speichern unter“ schreibt das Format der Zieldatei.

### CSV ohne Änderung öffnen und speichern

Beim Lesen werden nur verlustfrei darstellbare Werte typisiert: Zahlen in der Schreibweise der Datei (Dezimalzeichen und korrekte Tausendergruppen des Gebietsschemas, höchstens 15 signifikante Stellen) und `TRUE`/`FALSE`. Führende Nullen, Exponenten, Vorzeichen `+`, `-0`, fremde Dezimalzeichen (`1.23` in einer Semikolon-Datei), Datumswerte, Uhrzeiten, Prozent- und Währungsangaben bleiben Text. Zeilenende (CRLF/LF), ein fehlender Schluss-Zeilenumbruch und eine UTF-8-BOM werden beim Speichern übernommen. Normalisiert werden weiterhin: unnötige Nullen und Tausendertrenner von Zahlen (`1.234,50` → `1234,5`), überflüssige Anführungszeichen, das Auffüllen kürzerer Zeilen auf die Breite der Tabelle und Anführungszeichen um Werte mit Leerzeichen am Rand.

Kodierung: Gültiges UTF-8 (mit oder ohne BOM) wird als UTF-8 gelesen, alles andere als Windows-1252 – die Kodierung von CSV-Exporten aus Excel unter deutschem oder westeuropäischem Windows. Beim Speichern wird die Kodierung der Datei beibehalten. Kommen Zeichen hinzu, die Windows-1252 nicht darstellen kann (z. B. chinesische Schriftzeichen), wird die Datei als UTF-8 mit BOM gespeichert statt Zeichen zu verlieren; der Editor weist einmal darauf hin. Andere Kodierungen (z. B. Mac Roman, UTF-16) werden nicht erkannt.

Format und Datenblatt hängen am Arbeitsmappenmodell (`csvSource`) und gelten damit für jeden Speicherweg: den Editor, die Agent-Werkzeuge für ungeöffnete Dateien und den MCP-Server. Trennzeichen, Dezimalzeichen, Zeilenende, Schluss-Zeilenumbruch, BOM und Kodierung bleiben erhalten. Gespeichert wird immer das Blatt mit den Daten der Datei, auch wenn ein anderes Blatt aktiv ist; es folgt einer Umbenennung.

### Formel-Injektion (CSV)

`writeCsv` schreibt Zellwerte unverändert, auch wenn sie mit `=`, `+`, `-` oder `@` beginnen; Formeln werden standardmäßig als Ergebnis geschrieben, nur mit `formulas: true` als `=…`. So bleibt eine geöffnete und gespeicherte Datei unverändert. Tabellenprogramme können solche Texte beim Öffnen als Formel ausführen („CSV Injection“). Wer CSV-Dateien für Dritte aus nicht vertrauenswürdigen Daten erzeugt, sollte das berücksichtigen. Ein möglicher, nicht eingebauter optionaler Export-Schalter wäre `writeCsv(wb, sheet, { escapeFormulas: true })` (bzw. ein eigener Befehl „Als CSV für Excel exportieren (Formeln entschärfen)“), der Textzellen mit diesen Anfangszeichen sowie Tab/CR ein `'` voranstellt, Zahlen wie `-5` aber unverändert lässt; er dürfte nie für das normale Speichern eines CSV-Dokuments gelten.

## Prüfungen

- Bytegenaues No-op-Speichern und Erhaltung aller Paketbestandteile bei Zell- und Formatänderungen.
- Serialisierung über JSON sowie VS-Code-Backup und Wiederherstellung.
- Kein Dateischreiben bei Schutzfehlern und erkannten externen Änderungen.
- Unbekannte Formeln, Matrixbereiche, Rich Text, numerische Originalpräzision, Datumsmodus 1904 und Signaturen.
- Append-only Styles mit unbekannten Schrift-, Rahmen-, Ausrichtungs- und Schutzattributen.
- Sortieren mit exakter Mitnahme der Formatdatensätze sowie alle oben genannten Sperren; neues Blatt am Ende (auch in `excel-features.xlsx` mit lokalen Namen und sehr verstecktem Blatt); neue Style-Tabelle für Dateien ohne `styles.xml`; leere und nur aus Leerraum bestehende XML-Elemente (`<cols/>`, `<sheetViews> </sheetViews>`, leere Style-Sammlungen).
- Eine unabhängige openpyxl-Testdatei in `packages/core/test/fixtures/excel-features.xlsx` mit Bildern, Kommentaren, Verbindungen zwischen Teilen, Datenüberprüfung, Data Bars, Diagramm, lokalen Namen, Druckeinstellungen und ausgeblendeten Inhalten.
- Zusätzlich wurde die Microsoft-Datei „Financial Sample.xlsx“ geprüft: Die 26 ursprünglichen Teile bleiben erhalten; bei einer Zell-/Formatänderung werden ausschließlich Blatt, Style-Ergänzungen und Berechnungskennzeichen geändert. Der native Stil `TableStyleLight9` bleibt bestehen.

Das ist ein erhaltender Speicherpfad mit konservativen Schutzgrenzen, keine vollständige Implementierung aller Excel-Funktionen. Die beschriebenen Prüfungen ersetzen keine Zertifizierung aller Excel-Dateivarianten.

Technischer Hintergrund: [Aufbau eines SpreadsheetML-Dokuments](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/structure-of-a-spreadsheetml-document), [Berechnungskette](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-the-calculation-chain).
