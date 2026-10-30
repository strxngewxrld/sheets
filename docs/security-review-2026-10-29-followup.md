# Ergänzende Sicherheitsprüfung vom 29.10.2026

Geprüft: Sheets 1.2.2, Commit `9ea9fc155726e08b0c886432fc83e1fddc1a62c2`.
Lokale Quellcodeprüfung und isolierte Reproduktionen unter Windows mit Node.js 26.8.1.
Dieser Bericht ergänzt `security-review-2026-10-29.md`; dessen CSV-Korrekturen und
historische Prüfergebnisse bleiben unverändert dokumentiert.

## Stand nach der Korrektur

Alle drei unten dokumentierten Befunde sind im lokalen Quellcode behoben:

- Datumsargumente werden auf den gültigen Excel-Bereich 0 bis 2.958.465 geprüft.
  Arbeitswochen werden rechnerisch gezählt; WORKDAY sucht das Ergebnis binär,
  statt jeden Tag einzeln zu durchlaufen. Feiertage und die fiktive Datumszahl 60
  bleiben berücksichtigt. Ungültige oder unerreichbare Ergebnisse ergeben `#NUM!`.
- Eine gemeinsame Prüfung begrenzt dichte Formelbereiche und Zwischenmatrizen vor
  der Allokation auf eine Million Positionen. Sie greift für Broadcasting, Wrapping,
  TEXTSPLIT-Padding, Kriterienmasken, Lookups, CHOOSEROWS/CHOOSECOLS und LAMBDA-Arrays.
  Ungültige Dimensionen und Überschreitungen ergeben `#CALC!`.
- OFFSET prüft zusätzlich ganzzahlige Dimensionen und Blattgrenzen; ungültige
  Referenzen ergeben `#REF!`. INDEX, ROWS und COLUMNS können Referenzen anhand ihrer
  Grenzen verarbeiten, ohne dafür das vollständige Blatt zu materialisieren.
- Die bisherigen XLSX-Reproduktionen öffnen kontrolliert mit Zellfehlern; der Rechner
  bleibt anschließend nutzbar. Es wurde kein zusätzlicher Worker eingeführt.
- Der MCP-Buildbefehl verwendet plattformübergreifend passende Anführungszeichen.

Bewusste Grenze: Formeln mit dichten Bereichen oder Zwischenmatrizen über einer
Million Positionen werden abgewiesen. Das gilt auch für Kriterienmasken oder
Lookup-Vektoren, die anhand der deklarierten vollständigen Referenz aufgebaut würden.
Das in der zweiten Ergänzung (unten) eingeführte Rechenbudget ersetzt keine
Workerisolation.

Aktuelle Validierung nach der Korrektur:

- `npm audit --json --cache .test-tools/npm-cache`: **0 bekannte Schwachstellen**,
  einschließlich Entwicklungsabhängigkeiten. Der Nutzer hat den Online-Abgleich
  ausdrücklich freigegeben; die Übertragung erfolgte danach erfolgreich.
- `npm test -- --reporter=dot`: **959 Tests in 46 Dateien bestanden**, darunter
  34 neue Regressionstests. Der bisherige Symlink-Rechtefehler entfiel beim Lauf
  außerhalb der Dateisystem-Sandbox.
- `npm run build:all`: Typprüfung sowie Extension-, Webview- und MCP-Build erfolgreich.
- `node scripts/security-review.mjs`: Alle sieben isolierten Proben bestanden;
  kein Timeout und kein Speichermangelfehler. Die sechs Angriffsproben liefern
  die erwarteten kontrollierten Fehlerwerte.

## Reproduzierter Ausgangszustand vor der Korrektur

Drei bestätigte Befunde betreffen die Verfügbarkeit. Eine fremde XLSX kann bereits beim
Öffnen den synchronen Formelrechner blockieren oder dessen Speicher erschöpfen.
Die beiden XLSX-Reproduktionen benötigen jeweils nur rund 3,3 KB Dateigröße.
Betroffen sind der Editor und bei headless Agent-Zugriffen auch der Extension-Host;
der eigenständige MCP-Server verwendet denselben Rechner.
Eine Ausführung fremden Codes oder ein Abfluss von Daten wurde bei dieser Prüfung
nicht nachgewiesen. Dies ist keine Garantie für deren Abwesenheit.

| ID | Priorität | Befund | Code |
| --- | --- | --- | --- |
| S-01 | Hoch | Endlosschleifen bei großen Datumsargumenten | `packages/core/src/formula/functions/date.ts:68`, `:76` |
| S-02 | Hoch | Unbegrenzte Materialisierung von Formelbereichen | `packages/core/src/calc.ts:398` |
| S-03 | Mittel | Array-Operationen umgehen die vorhandene Zellgrenze | `packages/core/src/formula/evaluator.ts:215`, `packages/core/src/formula/functions/array.ts:51`, `packages/core/src/formula/functions/text.ts:104` |

Die Prioritäten bewerten das Risiko für die Endanwender dieses Projekts; sie sind keine CVSS-Werte.

## S-01: Datumsfunktionen können dauerhaft blockieren

Reproduktionen in Zelle A1 einer ansonsten leeren XLSX:

```excel
=WORKDAY(45000,1e20)
=NETWORKDAYS(1e20,1e20)
```

Bei dieser Größenordnung kann IEEE-754-Arithmetik die Schleifenvariable durch
Addition oder Subtraktion von 1 nicht mehr verändern. `workday()` vermindert
`days` deshalb nie; `networkDays()` erhöht `d` nie. Es fehlen vorherige Grenzen
für gültige Datumswerte und Arbeitsdauer sowie eine Laufzeitbegrenzung.
Auch kleinere, aber extrem große Argumente können lange Schleifen verursachen.

Beide XLSX-Öffnungen mussten nach 2 Sekunden vom übergeordneten Prüfprozess beendet
werden. Die Endlosigkeit ergibt sich zusätzlich aus der unveränderten Schleifenvariable.
Die Kontrollprobe `WORKDAY(45000,1)` öffnete normal und ergab 45001.
`Workbook.load()` berechnet standardmäßig automatisch neu (`workbook.ts:106`).
Ein asynchroner Request-Timeout kann eine bereits laufende synchrone Schleife nicht unterbrechen.

Abhilfe: Eingaben vor Schleifen gegen endliche, fachlich gültige Datums- und
Dauergrenzen prüfen; Arbeitswochen rechnerisch überspringen statt jeden Tag
einzeln zu durchlaufen. Zusätzlich ein gemeinsames Rechenbudget beziehungsweise
einen abbrechbaren Worker für fremde Formeln einführen.

## S-02: Große Referenzen erschöpfen den Speicher

Reproduktion in A1 einer ansonsten leeren XLSX:

```excel
=SUM(B1:XFC1048575)
```

`rangeValues()` legt für jeden adressierten Wert einen Array-Eintrag an.
Nur Referenzen, deren Ende exakt auf der letzten Excel-Zeile oder -Spalte liegt,
werden auf den benutzten Bereich begrenzt. Die obige Referenz endet jeweils davor
und fordert über 17 Milliarden Einträge an, obwohl alle referenzierten Zellen leer sind.
Die Begrenzungen für Benutzerauswahlen und `SEQUENCE` greifen hier nicht.

Die rund 3,3 KB große XLSX führte im separaten Node-Prozess mit 160 MiB Heap
nach etwa 1,7 Sekunden zu einem fatalen Speichermangelfehler (Exitcode 134).
Das Limit begrenzte die Reproduktion; daraus folgt keine Aussage über die exakte
Absturzschwelle einer normalen VS-Code-Instanz.

Abhilfe: Vor jeder dichten Bereichsallokation die Dimensionen und deren Produkt
prüfen. Für Aggregate möglichst dünn besetzte Iteratoren einsetzen; die deklarierte
Bereichsform für Funktionen mit Positionssemantik erhalten. Bei Überschreitung
einen kontrollierten Zellfehler zurückgeben, bevor Speicher angefordert wird.

## S-03: Bestehende Array-Grenze wird umgangen

`MAX_ARRAY_CELLS = 1_000_000` wird nur in einzelnen Funktionen geprüft.
Die folgenden Ausdrücke lieferten in isolierten Reproduktionen größere Matrizen:

| Ausdruck | Tatsächlich erzeugte Zellen |
| --- | ---: |
| `=WRAPROWS({1},1000001)` | 1.000.001 |
| `=SEQUENCE(1100)+TRANSPOSE(SEQUENCE(1100))` | 1.210.000 |
| `=TEXTSPLIT(REPT(",",1100)&REPT(";",1100),",",";")` | 1.212.201 |

`wrap()` füllt ohne Obergrenze auf, `broadcast()` bildet das Produkt aus maximaler
Zeilen- und Spaltenzahl, und `TEXTSPLIT` füllt alle Zeilen auf die größte Breite auf.
Größere Argumente können so sehr viel Speicher beanspruchen. Eine spätere Prüfung
beim Ausgeben in das Tabellenblatt erfolgt erst nach der Allokation.
Die Reproduktionen belegen die Umgehung; ein Absturz mit größeren Array-Argumenten
wurde bewusst nicht zusätzlich provoziert.

Abhilfe: Gemeinsame Prüfung für endliche, ganzzahlige Dimensionen, Blattgrenzen und
Zellanzahl vor jeder Matrixerzeugung. Auch `WRAPCOLS`, Kombinationen von Matrizen
und das Auffüllen ungleich langer Textzeilen einbeziehen. Ein Gesamtbudget pro
Berechnung verhindert zusätzlich die wiederholte Erzeugung vieler erlaubter Arrays.

## Zweite Ergänzung vom 30.10.2026

Eine erneute Prüfung des korrigierten Stands fand fünf weitere Verfügbarkeitsbefunde.
Alle ließen sich mit einer 3 bis 5 KB großen XLSX allein durch das Öffnen auslösen.
Betroffen waren Editor, headless Agentenzugriffe im Extension-Host und der MCP-Server.
Codeausführung oder Datenabfluss wurden nicht gefunden. Alle Befunde sind behoben.

| ID | Priorität | Befund vor der Korrektur | Korrektur |
| --- | --- | --- | --- |
| S-04 | Hoch | `FACT`, `PERMUT`, `COMBIN`, `CUMIPMT`/`CUMPRINC` und `DB` liefen ohne Obergrenze über ihr Argument; `=FACT(1e12)` blockierte unbegrenzt. | `FACT` über 170 sowie Überläufe in `PERMUT`/`COMBIN` ergeben wie in Excel `#NUM!`. Periodenschleifen werden mit dem Rechenbudget verrechnet. |
| S-05 | Hoch | Text hatte keine Längengrenze. `=UPPER(REPT("a",4e8))`, `=FIXED(1,3e8)` und `=TEXT(1,REPT("0",3e8))` führten zu einem nicht abfangbaren Speichermangel (Exitcode 134). | Excel-Grenze von 32.767 Zeichen, vor der Erzeugung geprüft für `REPT`, `&`, `CONCAT`, `CONCATENATE`, `TEXTJOIN`, `SUBSTITUTE` und `ARRAYTOTEXT`, danach für jedes Funktionsergebnis. `FIXED` erlaubt wie Excel höchstens 127 Nachkommastellen. Zahlenformate mit mehr als 100 Nachkommastellen funktionieren jetzt statt einen Fehler auszulösen. |
| S-06 | Hoch | Diagrammkategorien wurden nicht auf den benutzten Bereich begrenzt (`packages/core/src/chart.ts`). Kategorien `Sheet1!B1:XFD1048576` erschöpften den Speicher beim Zeichnen und im Agentenwerkzeug `list_charts`. | Kategorien und Werte werden gemeinsam auf den benutzten Bereich und höchstens 32.000 Punkte je Referenz begrenzt. |
| S-07 | Mittel bis hoch | Ohne Gesamtbudget blockierten einzelne Formeln wie `=LET(f,LAMBDA(f,n,IF(n=0,0,f(f,n-1)+f(f,n-1))),f(f,40))` oder verschachteltes `MAKEARRAY` unbegrenzt. | Rechenbudget pro Formel und pro Neuberechnung, siehe unten. |
| S-08 | Mittel | Die Linkerkennung des Editors wertete das `HYPERLINK`-Argument bei jeder Zellauswahl ohne Budget aus (`packages/ui/src/view.ts`). | Auswertung über `CalcEngine.evaluateNode` mit demselben Budget. |

Zusätzlich wurde eine nie aufgerufene, unbegrenzte zweite `SEQUENCE`-Definition in
`math.ts` entfernt; `RANDARRAY`, `SEQUENCE` und Regressionen (`LINEST`, `TREND`,
`GROWTH`) werden mit dem Budget verrechnet.

### Rechenbudget

`packages/core/src/formula/limits.ts` zählt Arbeitseinheiten: ausgewertete Knoten,
LAMBDA-Aufrufe, erzeugte Matrix- und Bereichszellen sowie Textzeichen.

| Grenze | Pro Formel | Pro Neuberechnung |
| --- | ---: | ---: |
| Arbeitseinheiten (begrenzt den Speicher) | 10 Mio. | 100 Mio. |
| Laufzeit (begrenzt die CPU-Zeit) | 2 s | 5 s |

Eine Überschreitung ergibt `#CALC!`. Die Ausnahme ist kein `CellError`, damit
`IFERROR` oder andere Funktionen sie nicht in einen Wert verwandeln und weiterrechnen.
Formeln, die beim Rechnen als Abhängigkeit berechnet werden, erhalten ein eigenes
Formelbudget; die Neuberechnung teilt ein gemeinsames Budget. Die Zeitgrenze ist
nötig, weil LAMBDA-Aufrufe etwa zwanzigmal teurer sind als einfache Matrixoperationen.

Bewusste Grenzen:

- Legitime Formeln mit mehr als 2 s Rechenzeit ergeben `#CALC!`, auf langsamen
  Rechnern also früher. Beispiel: `MAKEARRAY` mit einer Million LAMBDA-Aufrufen
  benötigte auf dem Prüfrechner 2,6 s und wird abgewiesen.
- Ist das Budget einer Neuberechnung erschöpft, erhalten alle restlichen Formeln
  dieser Neuberechnung `#CALC!`. Welche das sind, hängt bei der Zeitgrenze von der
  Rechnergeschwindigkeit ab. Die nächste Bearbeitung berechnet nur die betroffenen
  Zellen neu.
- Eine fremde Datei kann den Editor weiterhin für bis zu etwa 5 s je Neuberechnung
  blockieren. Nur ein abbrechbarer Worker würde das vollständig verhindern.

### Validierung

- `npm test`: **990 Tests in 46 Dateien bestanden**, darunter 31 neue Regressionstests
  in `packages/core/test/resource-limits.test.ts`.
- `npm run build:all` und `git diff --check`: erfolgreich.
- `node scripts/security-review.mjs`: alle **18** isolierten Proben bestanden,
  darunter 11 neue. Exponentielle Rekursion und verschachteltes `MAKEARRAY` enden
  nach 2,3 bzw. 2,5 s, alle anderen nach unter 0,3 s.
- Eine XLSX mit 20 exponentiellen Formeln (3,4 KB) öffnet nach 5,2 s mit `#CALC!`
  in allen Zellen.
- Windows-MCP: isolierter Extension-Development-Host mit eigenem Profil. Die Dateien
  mit 10 exponentiellen Formeln, mit Textexplosion/`FACT`/`HYPERLINK` und mit
  ganzseitigem Diagramm öffneten mit Fehlerwerten bzw. gezeichnetem Diagramm.
  Eingaben danach wurden sofort berechnet; die Auswahl der `HYPERLINK`-Zelle
  blockierte nicht. Auch die Wiederherstellung aus der VS-Code-Sicherung blieb
  begrenzt. Eine Kontrolldatei rechnete normal.

## Bestehende Schutzmaßnahmen und ursprüngliche Prüfgrenzen

Geprüft wurden insbesondere Webview/CSP und HTML/SVG-Ausgaben, Linköffnung,
Workspace- und Symlinkgrenzen für Agenten, Workspace Trust, CSV-Formelschutz,
XLSX/ZIP/XML-Verarbeitung und Update-Download mit SHA-256-Prüfung.
Hier wurden über die genannten Befunde hinaus keine weiteren belastbaren
Schwachstellen bestätigt. Bestehende Sicherheits- und Fachtests wurden ausgeführt.
Windows-MCP war für die reproduzierten Core-Probleme nicht erforderlich;
eine interaktive UI-Sicherheitsprüfung oder Prüfung des ausgelieferten VSIX erfolgte nicht.

Die folgenden Laufresultate stammen aus der ursprünglichen Prüfung vor den
Korrekturen; die aktuellen Ergebnisse stehen oben.

- `npm run typecheck`: erfolgreich.
- `npm test -- --reporter=dot`: 924 bestanden, 1 fehlgeschlagen (925 insgesamt).
  Der Fehler betrifft das Anlegen eines Windows-Symlinks (`EPERM`) in
  `test/file-uri.test.ts:77`; der dynamische lokale Symlinktest bleibt damit offen.
- Die lokal fehlende, bereits deklarierte Abhängigkeit `re2js@2.9.0` wurde für die
  Tests nachgeladen. Paketmanifest und Lockfile wurden nicht geändert.
- Ein Online-Abgleich mit npm-Advisories wurde nicht abgeschlossen. Der normale
  Aufruf scheiterte am Netzwerk; die automatische Freigabeprüfung lehnte die
  Eskalation wegen Übertragung der Projekt-Paketnamen und -Versionen an npm ab.
  Deshalb wird kein Urteil über aktuelle bekannte Abhängigkeitslücken abgegeben.

## Reproduktion

Nach Installation der deklarierten Abhängigkeiten im Repository-Hauptordner:

```powershell
npm run typecheck
node scripts/security-review.mjs
```

Der Prüfer erstellt die XLSX-Daten ausschließlich im Arbeitsspeicher und verändert
keine Nutzertabellen. Jede Probe läuft in einem separaten Prozess mit 160 MiB
JavaScript-Heap und 5 Sekunden Zeitlimit. Er meldet JSON-Zeilen mit Ergebnis,
Timeout und Speicherfehler. Das Skript prüft die erwarteten kontrollierten Ergebnisse
und beendet sich bei Abweichungen, Timeouts oder Prozessfehlern mit Exitcode 1.
