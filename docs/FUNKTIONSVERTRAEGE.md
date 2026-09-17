# Funktionsvertrag des Ausgangsstands

Für jede VBS-Funktion und erkannte C#-Methode: Eingabe, Ausgabe, Prüfregel und Seiteneffekt. Die Quellen bleiben im unveränderten bereitgestellten ZIP; Zeilennummern beziehen sich auf dekodierten Text.

## KC_KitaFino_Drop.cs:115 — AddCol

- Eingang: Spaltenname, Breite
- Ausgang: Grid-Spalte ohne Sortierung
- Prüfregel: Breite/Name direkt übernommen
- Seiteneffekt: ändert UI

## KC_KitaFino_Drop.cs:117 — Busy

- Eingang: Boolean, Statustext
- Ausgang: Fortschrittsanzeige und gesperrte Buttons
- Prüfregel: während Verarbeitung keine konkurrierende Aktion
- Seiteneffekt: UI, Timer, DoEvents

## KC_KitaFino_Drop.cs:125 — OnDragEnter

- Eingang: Drag-Daten
- Ausgang: Copy/None
- Prüfregel: FileGroupDescriptor(W) oder FileDrop erforderlich
- Seiteneffekt: UI-Drag-Effekt

## KC_KitaFino_Drop.cs:130 — DraggedMailCount

- Eingang: Drag-Daten
- Ausgang: Anzahl, Fallback 1
- Prüfregel: Descriptor-Stream mindestens vier Bytes, positiver Int32
- Seiteneffekt: liest Speicherstream; Position wiederherstellen

## KC_KitaFino_Drop.cs:148 — AddCapture

- Eingang: Capture-Pfad
- Ausgang: Boolean und neue Grid-/Capture-Zeile
- Prüfregel: Datei vorhanden, erste Zeile OK|
- Seiteneffekt: liest Datei, ändert UI/temporäre Liste

## KC_KitaFino_Drop.cs:158 — OnDragDrop

- Eingang: Drag-Daten
- Ausgang: aufgenommene Capture-Liste
- Prüfregel: nach Vorschau gesperrt; Capture-Exitcode und OK| prüfen
- Seiteneffekt: startet Capture-VBS, schreibt/liest Tempdateien, UI

## KC_KitaFino_Drop.cs:194 — OnAction

- Eingang: kein Parameter
- Ausgang: Vorschau oder Verbuchung
- Prüfregel: previewReady entscheidet
- Seiteneffekt: ruft DoPreview/DoCommit

## KC_KitaFino_Drop.cs:195 — DoPreview

- Eingang: Capture-Liste und Workbook-Pfad
- Ausgang: farbige Vorschau, geprüfter Wochencache
- Prüfregel: Exitcode/Datei/TSV prüfen; mindestens sieben Spalten
- Seiteneffekt: externe Queue/TSV, VBS-Prozess, UI; keine Produktionswerte

## KC_KitaFino_Drop.cs:230 — DoCommit

- Eingang: Vorschau, Queue, Wochencache
- Ausgang: Writer-Verbuchung
- Prüfregel: ausdrückliches Ja, Writer-Exitcode; Altwriter prüft erneut
- Seiteneffekt: startet Writer, Produktionszellen/Protokoll, UI

## KC_KitaFino_Drop.cs:245 — OnGridCellContentClick

- Eingang: Grid-Event
- Ausgang: ggf. entfernte Zeile
- Prüfregel: gültige Zeile, DeleteMail-Spalte, keine fertige Vorschau
- Seiteneffekt: UI/Temp-Capture-Liste

## KC_KitaFino_Drop.cs:253 — DeleteSelectedRows

- Eingang: selektierte Grid-Zeilen
- Ausgang: entfernte Zeilen
- Prüfregel: vor Vorschau und ausdrückliche Bestätigung
- Seiteneffekt: UI und Capture-Tempdateien

## KC_KitaFino_Drop.cs:262 — DeleteRows

- Eingang: Zeilenindizes
- Ausgang: verkleinerte Grid-/Capture-Liste
- Prüfregel: Indizes sortieren und von hinten löschen
- Seiteneffekt: UI, Tempdateien

## KC_KitaFino_Drop.cs:275 — OnGridDoubleClick

- Eingang: Grid-Zeile
- Ausgang: Detaildialog
- Prüfregel: gültige Zeile, Vorschauwerte anzeigen
- Seiteneffekt: UI

## KC_KitaFino_Drop.cs:294 — C

- Eingang: Grid-Zeile, Spalte
- Ausgang: Zelltext oder leer
- Prüfregel: Null-/Indexprüfung im Quelltext
- Seiteneffekt: liest UI

## KC_KitaFino_Drop.cs:296 — AppendLive

- Eingang: Status, Text
- Ausgang: Protokollzeile
- Prüfregel: UI-Protokoll konsistent formatieren
- Seiteneffekt: UI

## KC_KitaFino_Drop.cs:308 — ReadProtocol

- Eingang: Writer-Protokollpfad
- Ausgang: neue Protokollzeilen
- Prüfregel: nur neue Zeilen seit gelesener Position
- Seiteneffekt: liest Datei, UI

## KC_KitaFino_Drop.cs:321 — RunWaitWithProtocol

- Eingang: Programm, Argumente
- Ausgang: Exitcode
- Prüfregel: während Prozesslauf Protokoll lesen, Prozessende abwarten
- Seiteneffekt: startet Fremdprozess, liest Protokoll, DoEvents

## KC_KitaFino_Drop.cs:336 — RunWait

- Eingang: Programm, Argumente
- Ausgang: Exitcode
- Prüfregel: Prozessende abwarten
- Seiteneffekt: startet Fremdprozess, DoEvents

## KC_KitaFino_Drop.cs:340 — ConvertSonderkost

- Eingang: Formtext
- Ausgang: Einzelkategorie oder leer
- Prüfregel: bekannte Kategorie-Synonyme; Mehrfachformen separat prüfen
- Seiteneffekt: keine

## KC_KitaFino_Drop.cs:359 — CustomerTarget

- Eingang: Einrichtung, Sonderkostdetail
- Ausgang: Kundenziel/PRÜFEN
- Prüfregel: explizite Cluster-/Lünern-/Strolche-/Hertinger-/Liedbach-Regeln
- Seiteneffekt: keine

## KC_KitaFino_Drop.cs:390 — CellColumnFromMainTarget

- Eingang: Vorschau-Zielkette
- Ausgang: Spaltenbuchstaben
- Prüfregel: Ziel/Ziele=Zelladresse
- Seiteneffekt: keine

## KC_KitaFino_Drop.cs:395 — BaseRowFromMainTarget

- Eingang: Vorschau-Zielkette
- Ausgang: Tagesbasiszeile/0
- Prüfregel: Zeile=n
- Seiteneffekt: keine

## KC_KitaFino_Drop.cs:399 — SonderkostRowOffset

- Eingang: Kategorie
- Ausgang: KW-Zeilenoffset/-1
- Prüfregel: elf Standardzeilen; Fisch nicht KW
- Seiteneffekt: keine

## KC_KitaFino_Drop.cs:407 — ConvertSonderkostMulti

- Eingang: Formbezeichnung
- Ausgang: eindeutige Kategorienliste
- Prüfregel: Kategorie einmal aufnehmen; keine Doppelzählung
- Seiteneffekt: keine

## KC_KitaFino_Drop.cs:424 — ManualCombinationCell

- Eingang: Kunde, Form
- Ausgang: freigegebenes Kombinationsfeld oder leer
- Prüfregel: nur Strolche bekannte drei Kombinationen, Kita2_4 bekannte Dreierkombination
- Seiteneffekt: keine

## KC_KitaFino_Drop.cs:445 — SonderkostCountFromLine

- Eingang: Detailzeile
- Ausgang: Menge [n]
- Prüfregel: Alt-C# Fallback 1; neuer Writer verlangt Kontrollsumme
- Seiteneffekt: keine

## KC_KitaFino_Drop.cs:450 — SonderkostRawFormFromLine

- Eingang: Detailzeile
- Ausgang: Form ohne Kontext/Menge
- Prüfregel: Doppelpunkt und letzter Mengenblock
- Seiteneffekt: keine

## KC_KitaFino_Drop.cs:461 — ShowSonderkostWritePreview

- Eingang: selektierte Vorschau und TSV
- Ausgang: aggregierte Schreibprüfung
- Prüfregel: Kunde/Kategorie/Kombination eindeutig, Quellessen=Planessen
- Seiteneffekt: öffnet UI, liest TSV; schreibt keine Excelwerte

## KC_KitaFino_Drop.cs:587 — ImportFromFile

- Eingang: Dateiauswahl/MSG
- Ausgang: Capture-Dateien und Grid-Zeilen
- Prüfregel: nach Vorschau gesperrt; Capture-Datei prüfen
- Seiteneffekt: Outlook-COM/Tempdateien, UI

## KC_KitaFino_DropImport.vbs:362 — IIfSafe

- Eingang: Boolean, zwei Varianten
- Ausgang: gewählte Variante
- Prüfregel: keine echte Lazy-Auswertung der Argumente
- Seiteneffekt: keine

## KC_KitaFino_DropImport.vbs:366 — ReadUtf8

- Eingang: Pfad
- Ausgang: Text
- Prüfregel: ADODB.Stream UTF-8
- Seiteneffekt: liest Datei; neue Laufzeit benötigt keine

## KC_KitaFino_DropImport.vbs:374 — WriteUtf8

- Eingang: Pfad, Text
- Ausgang: Datei
- Prüfregel: ADODB.Stream UTF-8
- Seiteneffekt: überschreibt Datei; im Port entfernt

## KC_KitaFino_DropImport.vbs:380 — ReadQueue

- Eingang: Queue-Pfad
- Ausgang: Array nichtleerer Capture-Pfade
- Prüfregel: Zeilen normalisieren
- Seiteneffekt: liest externe Queue; durch interne Blätter ersetzt

## KC_KitaFino_DropImport.vbs:396 — ReadCapture

- Eingang: Capture-Pfad
- Ausgang: Betreff/Body/ID/Sender-Array
- Prüfregel: Base64-Felder
- Seiteneffekt: liest Datei; durch Outlook-Objekt ersetzt

## KC_KitaFino_DropImport.vbs:411 — B64ToUtf8

- Eingang: Base64
- Ausgang: Text
- Prüfregel: MSXML Base64 und ADODB UTF-8
- Seiteneffekt: Speicher-COM, keine Persistenz

## KC_KitaFino_DropImport.vbs:420 — Rx1

- Eingang: Text, Regex
- Ausgang: erste Capture-Gruppe oder leer
- Prüfregel: case-insensitive, erster Treffer
- Seiteneffekt: keine

## KC_KitaFino_DropImport.vbs:427 — AddN

- Eingang: zwei Diagnosetexte
- Ausgang: verketteter Text
- Prüfregel: leerer erster Text ohne Trennzeichen, sonst Semikolon
- Seiteneffekt: keine

## KC_KitaFino_DropImport.vbs:431 — FlatText

- Eingang: Text
- Ausgang: normalisierte einzelne Textzeile
- Prüfregel: CR/LF/TAB und Mehrfachleerzeichen ersetzen
- Seiteneffekt: im Port ByVal; Eingabe bleibt unverändert

## KC_KitaFino_DropImport.vbs:441 — GermanDate

- Eingang: Betreff
- Ausgang: erstes ausgeschriebenes deutsches Datum oder leer
- Prüfregel: Monatsname und vierstelliges Jahr
- Seiteneffekt: keine

## KC_KitaFino_DropImport.vbs:452 — SecondGermanDate

- Eingang: Betreff
- Ausgang: zweites Datum oder leer
- Prüfregel: mindestens zwei Regex-Treffer
- Seiteneffekt: keine

## KC_KitaFino_DropImport.vbs:465 — MonthNo

- Eingang: Monatsname
- Ausgang: Monatsnummer
- Prüfregel: deutsche Monatsnamen inkl. Maerz
- Seiteneffekt: keine

## KC_KitaFino_DropImport.vbs:482 — DateText

- Eingang: Datum/Variant
- Ausgang: dd.mm.yyyy oder leer
- Prüfregel: IsDate
- Seiteneffekt: keine

## KC_KitaFino_DropImport.vbs:490 — DetectOrderType

- Eingang: Betreff
- Ausgang: TAG/WOCHEN/UNBEKANNT
- Prüfregel: ein bzw. zwei erkannte Daten
- Seiteneffekt: keine

## KC_KitaFino_DropImport.vbs:502 — DetectOrderRange

- Eingang: Betreff
- Ausgang: Datum oder Bereich
- Prüfregel: erkannte Daten
- Seiteneffekt: keine

## KC_KitaFino_DropImport.vbs:514 — NumberFromCell

- Eingang: Zellinhalt
- Ausgang: erste Ganzzahl oder 0
- Prüfregel: erste Ziffernfolge
- Seiteneffekt: keine

## KC_KitaFino_DropImport.vbs:521 — IsWeekCopySheet

- Eingang: Blattname
- Ausgang: Boolean
- Prüfregel: Kopie/Backup/Print-/Temp-Namen ausschließen
- Seiteneffekt: keine

## KC_KitaFino_DropImport.vbs:538 — IsRealWeekSheetName

- Eingang: Blattname
- Ausgang: Boolean
- Prüfregel: ^kw\s*\d*$
- Seiteneffekt: keine

## KC_KitaFino_DropImport.vbs:545 — WeekStartFromSheet

- Eingang: KW-Blatt
- Ausgang: Montag oder leer
- Prüfregel: alt A2 zuerst; Port eindeutig A8/Montag
- Seiteneffekt: liest Blatt

## KC_KitaFino_DropImport.vbs:558 — FindWeekSheetByDate

- Eingang: Workbook, Datum
- Ausgang: eindeutiges KW-Blatt und Diagnose
- Prüfregel: Originalblatt, ISO-KW, Fallback Datum; Port prüft Jahr
- Seiteneffekt: liest Blattstruktur

## KC_KitaFino_DropImport.vbs:615 — HeaderCol

- Eingang: Blatt, Kundenspalte
- Ausgang: Spaltennummer oder 0
- Prüfregel: exakter Header C:O, Liedbach-Alias
- Seiteneffekt: liest Zeile 3

## KC_KitaFino_DropImport.vbs:625 — DayBaseRow

- Eingang: KW-Blatt, Lieferdatum
- Ausgang: 4+13*Offset oder 0
- Prüfregel: Montag bis Freitag
- Seiteneffekt: liest A8 im Port

## KC_KitaFino_DropImport.vbs:634 — ColLetter

- Eingang: Spaltennummer
- Ausgang: Excel-Spaltenbuchstaben
- Prüfregel: Base26
- Seiteneffekt: keine

## KC_KitaFino_DropImport.vbs:643 — LueText

- Eingang: keine
- Ausgang: Lünern
- Prüfregel: ChrW(252)
- Seiteneffekt: keine

## KC_KitaFino_DropImport.vbs:647 — InstitutionById

- Eingang: KitaFino-ID
- Ausgang: Einrichtungsname oder leer
- Prüfregel: fünf explizit freigegebene IDs
- Seiteneffekt: keine

## KC_KitaFino_DropImport.vbs:658 — TargetLabel

- Eingang: KitaFino-ID
- Ausgang: Kundenziel
- Prüfregel: explizite ID-Zuordnung
- Seiteneffekt: keine

## KC_KitaFino_DropImport.vbs:669 — QtyAfterLabel

- Eingang: Mailtext, Labelregex
- Ausgang: Menge oder leer
- Prüfregel: normalisierter Text, Menge vor Blockgrenze
- Seiteneffekt: keine im Port

## KC_KitaFino_DropImport.vbs:678 — ClusterDetails

- Eingang: Mailtext
- Ausgang: vier Clusterwerte als Text
- Prüfregel: alle vier numerisch
- Seiteneffekt: keine

## KC_KitaFino_DropImport.vbs:691 — LuenernDetails

- Eingang: Mailtext
- Ausgang: Lünern1/2 als Text
- Prüfregel: beide numerisch, alternative Schreibweise
- Seiteneffekt: keine

## KC_KitaFino_DropImport.vbs:700 — KeyWasBooked

- Eingang: Protokoll, Mail-Key
- Ausgang: Boolean
- Prüfregel: Spalte F exakter Key, O=JA
- Seiteneffekt: liest Protokoll

## KC_KitaFino_DropImport.vbs:712 — GetWorkbookByPath

- Eingang: Pfad
- Ausgang: gebundenes Workbook oder Nothing
- Prüfregel: GetObject, COM-Fehler abfangen
- Seiteneffekt: COM-Bindung

## KC_KitaFino_DropImport.vbs:721 — MarkAutoCell

- Eingang: Zelle
- Ausgang: Rahmenmarkierung
- Prüfregel: AUTO-Kennzeichnung im Altwriter
- Seiteneffekt: ändert Zellrahmen; Port lässt bestehende Gestaltung erhalten

## KC_KitaFino_DropImport.vbs:730 — ColorOne

- Eingang: Zelle, Status
- Ausgang: Ampelfarbe
- Prüfregel: GRUEN/GELB/ROT
- Seiteneffekt: ändert Hintergrund

## KC_KitaFino_DropImport.vbs:738 — WeekDaySegment

- Eingang: Mailtext, Datum
- Ausgang: Tagesabschnitt oder leer
- Prüfregel: Zusammenfassung-Datumsmarker; Freitag vor Footer enden
- Seiteneffekt: keine

## KC_KitaFino_DropImport.vbs:763 — GermanMonthName

- Eingang: Monatsnummer
- Ausgang: Monatsname
- Prüfregel: Array; März im Altcode als Maerz
- Seiteneffekt: keine

## KC_KitaFino_DropImport.vbs:769 — BlockQtySum

- Eingang: Abschnitt, Start-/Grenzregex
- Ausgang: Ganzzahl oder leer
- Prüfregel: Endzahlen und Sonderkost-Unterzählungen; negativ/unbekannt leer
- Seiteneffekt: keine

## KC_KitaFino_DropImport.vbs:802 — LuenernWeekQtySum

- Eingang: Abschnitt, Gruppe 1/2
- Ausgang: Blockmenge
- Prüfregel: BlockQtySum mit Lünern-Grenzen
- Seiteneffekt: keine

## KC_KitaFino_DropImport.vbs:806 — GroupQtySum

- Eingang: Abschnitt, Gruppe 1–4
- Ausgang: Blockmenge
- Prüfregel: BlockQtySum mit Gruppengrenzen
- Seiteneffekt: keine

## KC_KitaFino_DropImport.vbs:810 — ClusterQtySum

- Eingang: Abschnitt, Cluster 1–4
- Ausgang: Ganzzahl als Text oder leer
- Prüfregel: Blockgrenzen, Endzahlen, Sonderkost-Unterzählung zweimal abziehen
- Seiteneffekt: keine

## KC_KitaFino_OutlookCapture.vbs:54 — One

- Eingang: Text
- Ausgang: einzeiliger Text
- Prüfregel: CR/LF ersetzen
- Seiteneffekt: keine

## KC_KitaFino_OutlookCapture.vbs:58 — B64

- Eingang: Text
- Ausgang: Base64
- Prüfregel: UTF-8
- Seiteneffekt: Speicher-COM

## KC_KitaFino_OutlookCapture.vbs:78 — WriteUtf8

- Eingang: Pfad, Text
- Ausgang: Datei
- Prüfregel: ADODB.Stream UTF-8
- Seiteneffekt: überschreibt Datei; im Port entfernt

## KC_KitaFino_OutlookMultiCapture.vbs:33 — Utf8BytesNoBom

- Eingang: Text
- Ausgang: UTF-8 Bytes
- Prüfregel: BOM entfernen
- Seiteneffekt: Speicher-COM

## KC_KitaFino_OutlookMultiCapture.vbs:44 — B64Utf8

- Eingang: Text
- Ausgang: Base64 ohne Zeilenumbrüche
- Prüfregel: UTF-8 ohne BOM
- Seiteneffekt: Speicher-COM

## KC_KitaFino_OutlookMultiCapture.vbs:53 — WriteUtf8NoBom

- Eingang: Pfad, Text
- Ausgang: Capture-Datei
- Prüfregel: UTF-8 ohne BOM
- Seiteneffekt: Datei schreiben; ersetzt

## KC_KitaFino_Preview.vbs:194 — WriteDiag

- Eingang: Pfad, Status, Meldung
- Ausgang: System-Diagnose-TSV
- Prüfregel: Status und Meldung in feste Spalten
- Seiteneffekt: überschreibt UTF8-Diagnosedatei

## KC_KitaFino_Preview.vbs:198 — ReadUtf8

- Eingang: Pfad
- Ausgang: Text
- Prüfregel: ADODB.Stream UTF-8
- Seiteneffekt: liest Datei; neue Laufzeit benötigt keine

## KC_KitaFino_Preview.vbs:206 — WriteUtf8

- Eingang: Pfad, Text
- Ausgang: Datei
- Prüfregel: ADODB.Stream UTF-8
- Seiteneffekt: überschreibt Datei; im Port entfernt

## KC_KitaFino_Preview.vbs:212 — ReadQueue

- Eingang: Queue-Pfad
- Ausgang: Array nichtleerer Capture-Pfade
- Prüfregel: Zeilen normalisieren
- Seiteneffekt: liest externe Queue; durch interne Blätter ersetzt

## KC_KitaFino_Preview.vbs:228 — ReadCapture

- Eingang: Capture-Pfad
- Ausgang: Betreff/Body/ID/Sender-Array
- Prüfregel: Base64-Felder
- Seiteneffekt: liest Datei; durch Outlook-Objekt ersetzt

## KC_KitaFino_Preview.vbs:243 — B64ToUtf8

- Eingang: Base64
- Ausgang: Text
- Prüfregel: MSXML Base64 und ADODB UTF-8
- Seiteneffekt: Speicher-COM, keine Persistenz

## KC_KitaFino_Preview.vbs:252 — Rx1

- Eingang: Text, Regex
- Ausgang: erste Capture-Gruppe oder leer
- Prüfregel: case-insensitive, erster Treffer
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:259 — AddN

- Eingang: zwei Diagnosetexte
- Ausgang: verketteter Text
- Prüfregel: leerer erster Text ohne Trennzeichen, sonst Semikolon
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:263 — FlatText

- Eingang: Text
- Ausgang: normalisierte einzelne Textzeile
- Prüfregel: CR/LF/TAB und Mehrfachleerzeichen ersetzen
- Seiteneffekt: im Port ByVal; Eingabe bleibt unverändert

## KC_KitaFino_Preview.vbs:273 — GermanDate

- Eingang: Betreff
- Ausgang: erstes ausgeschriebenes deutsches Datum oder leer
- Prüfregel: Monatsname und vierstelliges Jahr
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:284 — SecondGermanDate

- Eingang: Betreff
- Ausgang: zweites Datum oder leer
- Prüfregel: mindestens zwei Regex-Treffer
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:297 — MonthNo

- Eingang: Monatsname
- Ausgang: Monatsnummer
- Prüfregel: deutsche Monatsnamen inkl. Maerz
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:314 — DateText

- Eingang: Datum/Variant
- Ausgang: dd.mm.yyyy oder leer
- Prüfregel: IsDate
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:322 — DetectOrderType

- Eingang: Betreff
- Ausgang: TAG/WOCHEN/UNBEKANNT
- Prüfregel: ein bzw. zwei erkannte Daten
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:334 — DetectOrderRange

- Eingang: Betreff
- Ausgang: Datum oder Bereich
- Prüfregel: erkannte Daten
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:346 — NumberFromCell

- Eingang: Zellinhalt
- Ausgang: erste Ganzzahl oder 0
- Prüfregel: erste Ziffernfolge
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:353 — IsWeekCopySheet

- Eingang: Blattname
- Ausgang: Boolean
- Prüfregel: Kopie/Backup/Print-/Temp-Namen ausschließen
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:370 — IsRealWeekSheetName

- Eingang: Blattname
- Ausgang: Boolean
- Prüfregel: ^kw\s*\d*$
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:377 — WeekStartFromSheet

- Eingang: KW-Blatt
- Ausgang: Montag oder leer
- Prüfregel: alt A2 zuerst; Port eindeutig A8/Montag
- Seiteneffekt: liest Blatt

## KC_KitaFino_Preview.vbs:390 — FindWeekSheetByDate

- Eingang: Workbook, Datum
- Ausgang: eindeutiges KW-Blatt und Diagnose
- Prüfregel: Originalblatt, ISO-KW, Fallback Datum; Port prüft Jahr
- Seiteneffekt: liest Blattstruktur

## KC_KitaFino_Preview.vbs:447 — HeaderCol

- Eingang: Blatt, Kundenspalte
- Ausgang: Spaltennummer oder 0
- Prüfregel: exakter Header C:O, Liedbach-Alias
- Seiteneffekt: liest Zeile 3

## KC_KitaFino_Preview.vbs:457 — DayBaseRow

- Eingang: KW-Blatt, Lieferdatum
- Ausgang: 4+13*Offset oder 0
- Prüfregel: Montag bis Freitag
- Seiteneffekt: liest A8 im Port

## KC_KitaFino_Preview.vbs:466 — ColLetter

- Eingang: Spaltennummer
- Ausgang: Excel-Spaltenbuchstaben
- Prüfregel: Base26
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:475 — LueText

- Eingang: keine
- Ausgang: Lünern
- Prüfregel: ChrW(252)
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:479 — InstitutionById

- Eingang: KitaFino-ID
- Ausgang: Einrichtungsname oder leer
- Prüfregel: fünf explizit freigegebene IDs
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:490 — TargetLabel

- Eingang: KitaFino-ID
- Ausgang: Kundenziel
- Prüfregel: explizite ID-Zuordnung
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:501 — WeekDaySegment

- Eingang: Mailtext, Datum
- Ausgang: Tagesabschnitt oder leer
- Prüfregel: Zusammenfassung-Datumsmarker; Freitag vor Footer enden
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:526 — GermanMonthName

- Eingang: Monatsnummer
- Ausgang: Monatsname
- Prüfregel: Array; März im Altcode als Maerz
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:531 — ClusterQtySum

- Eingang: Abschnitt, Cluster 1–4
- Ausgang: Ganzzahl als Text oder leer
- Prüfregel: Blockgrenzen, Endzahlen, Sonderkost-Unterzählung zweimal abziehen
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:574 — GroupQtySum

- Eingang: Abschnitt, Gruppe 1–4
- Ausgang: Blockmenge
- Prüfregel: BlockQtySum mit Gruppengrenzen
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:578 — LuenernWeekQtySum

- Eingang: Abschnitt, Gruppe 1/2
- Ausgang: Blockmenge
- Prüfregel: BlockQtySum mit Lünern-Grenzen
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:582 — BlockQtySum

- Eingang: Abschnitt, Start-/Grenzregex
- Ausgang: Ganzzahl oder leer
- Prüfregel: Endzahlen und Sonderkost-Unterzählungen; negativ/unbekannt leer
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:614 — SummaryQtySum

- Eingang: Tagesabschnitt
- Ausgang: Summe der Menümengen oder leer
- Prüfregel: vor Davon/Gruppe enden; alt nur >0
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:630 — QtyAfterLabel

- Eingang: Mailtext, Labelregex
- Ausgang: Menge oder leer
- Prüfregel: normalisierter Text, Menge vor Blockgrenze
- Seiteneffekt: keine im Port

## KC_KitaFino_Preview.vbs:639 — ClusterDetails

- Eingang: Mailtext
- Ausgang: vier Clusterwerte als Text
- Prüfregel: alle vier numerisch
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:652 — LuenernDetails

- Eingang: Mailtext
- Ausgang: Lünern1/2 als Text
- Prüfregel: beide numerisch, alternative Schreibweise
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:661 — SonderkostPreview

- Eingang: Mailtext
- Ausgang: aggregierte Kontext:Form [Menge]-Liste
- Prüfregel: strikte Sonderkostsignale, numerischer Zeilenabschluss, Dictionary je Kontext/Form
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:714 — IsStrictSonderkostLine

- Eingang: Zeile
- Ausgang: Boolean
- Prüfregel: Allergie-/Intoleranz-/ohne-/vegetar-Signale; Port ergänzt exakt Laktosefrei
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:736 — SonderkostPretty

- Eingang: Sonderkostliste
- Ausgang: mehrzeilige Anzeige
- Prüfregel: ein Datensatz je Zeile
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:749 — SonderkostFormCount

- Eingang: Sonderkostliste
- Ausgang: Formenanzahl
- Prüfregel: aggregierte Einträge zählen
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:757 — SonderkostMealCount

- Eingang: Sonderkostliste
- Ausgang: Essenanzahl
- Prüfregel: Mengen am Ende [n] summieren
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:773 — IsSonderkostLine

- Eingang: Zeile
- Ausgang: Boolean
- Prüfregel: breiter alter Zutatenfilter; nicht als Schreibfreigabe verwenden
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:788 — RxTest

- Eingang: Text, Regex
- Ausgang: Boolean
- Prüfregel: case-insensitive
- Seiteneffekt: keine

## KC_KitaFino_Preview.vbs:795 — KeyWasBooked

- Eingang: Protokoll, Mail-Key
- Ausgang: Boolean
- Prüfregel: Spalte F exakter Key, O=JA
- Seiteneffekt: liest Protokoll

## KC_KitaFino_Preview.vbs:807 — GetWorkbookByPath

- Eingang: Pfad
- Ausgang: gebundenes Workbook oder Nothing
- Prüfregel: GetObject, COM-Fehler abfangen
- Seiteneffekt: COM-Bindung

## KC_KitaFino_Preview.vbs:816 — MarkAutoCell

- Eingang: Zelle
- Ausgang: Rahmenmarkierung
- Prüfregel: AUTO-Kennzeichnung im Altwriter
- Seiteneffekt: ändert Zellrahmen; Port lässt bestehende Gestaltung erhalten

## KC_KitaFino_Preview.vbs:825 — ColorOne

- Eingang: Zelle, Status
- Ausgang: Ampelfarbe
- Prüfregel: GRUEN/GELB/ROT
- Seiteneffekt: ändert Hintergrund
