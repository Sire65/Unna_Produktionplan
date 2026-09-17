# Bestandsanalyse – 17.09.2026

Referenz: MASTER_KW38_FUNKTIONIERT_16-09-2026.xlsm, SHA256 7c0cf8e3d596ce11e4f9e022eb998e46850b875c023b120470a7c71899f6dabf. 223 VBA-Komponenten extrahiert. Referenz unverändert, Makros bei Analyse deaktiviert.

## Datenfluss

OutlookCapture/MultiCapture: Outlook-Auswahl → Betreff, Klartext, Sender, Message-ID (Fallback EntryID) → Base64-Textdatei. Keine Bestellprüfung; COM-Lesefehler führen zum Abbruch.

V20-C#: Dateiauswahl/Drag-Drop → Capture-Dateien → Preview-VBS → TSV/checkedweek-Cache → Writer-VBS. UI, Prozessaufrufe und temporäre Dateien. Sonderkost-Konvertierung enthält Kunden-, Standard- und Kombinationsregeln und Kontrollsumme. **Die C#-Sonderkost-Schreibprüfung schreibt ausdrücklich noch nichts.**

Preview: Betreff-ID und Menge, deutsches Datum, Original-KW-Blatt, Tages-/Gruppenmengen und Wochensumme → GRUEN/GELB/ROT, Vorschautexte, geprüfte Wochenmatrix. Reines Lesen der Produktionswerte. Sonderkost nur strikte Allergie-/Sonderkostsignale; Speiseplan-Zutaten ausgeschlossen.

Writer: Capture + geprüfter Wochen-Cache → KW-Zellen und Import_Protokoll/Import_Kitafino. Wochenwerte mit Snapshot, Konfliktabfrage und Rücklesen; Gruppen 2–4 für 59426 aggregiert. Tagespfad ohne gleichwertigen Transaktionsschutz. Historische Dubletten: Protokoll Spalte F und O=JA.

VBA KC_KitaFino_Drop_Start startet EXE per Shell; KC_KitaFino_Start ist ein älterer separater MSG-Import. Beide Einstiegspunkte werden auf die interne Warteschlange umgeleitet. KC_UI_ImportKitafino öffnet ebenfalls die neue Queue; Protokollnavigation und Sonderkostdruck bleiben erhalten. Protokollfilter bleiben bestehen. Workbook_Open/BeforeClose bekommen nur den neuen Start/Stopp-Aufruf.

## Erhaltene Regeln

59432 → Cluster1–4; 59430 → Lünern/Lünern2; 59431 → Liedbach; 59500 → Strolche; 59426 Gruppe1 → Kita1, Gruppen2–4 → Kita2_4. Wochenfreigabe nur 59432/59430/59431/59426. Wochenzeitraum Montag bis Freitag, alle fünf Segmente und Gesamtsumme erforderlich. Blockmengen ziehen Sonderkost-Unterzählungen nicht zusätzlich zu den Essen heran.

Sonderkost-Kombination zählt genau ein Essen. V20 freigegebene Kombinationen: Strolche Laktose+Nüsse, Laktose+Fructose, Nüsse+Soja+Südfrüchte; Kita2_4 Hülsenfrüchte+Moslem+Nüsse. Andere Kombinationen gesperrt. Fisch gehört in die Matrix, nicht in eine erfundene KW-Zeile.

## Belegte Abweichungen / gezielte Korrekturen

Alte WeekStartFromSheet liest A2 zuerst. A2 enthält TODAY(), A8 den tatsächlichen Montag. Alter Tagespfad addiert einen Tag zum Betreffdatum. Neue Zuordnung verwendet A8, prüft Jahr und Woche und schreibt das tatsächliche Betreff-Lieferdatum. Kein Ändern vorhandener Blattdaten oder Formeln.

Die Referenz hat eine externe Verknüpfung auf eine historische Mittelwerte-Datei auf Y:. Sie bleibt erhalten; Autarkie betrifft den KitaFino-Import. Vor vollständiger Workbook-Autarkie ist diese fachliche Abhängigkeit separat zu bewerten.

Queue-Inhalt und Mail-ID bleiben dauerhaft in der XLSM; Protokollbereinigung löscht keine neuen Dublettenbelege. Outlook-Scan liest nur Mails, bewegt/löscht/markiert keine Nachrichten.

## Vom Auftraggeber freigegebene Erweiterungen

Cluster-Kunden dürfen die vorhandenen MASTER-Kombinationsfelder R/S/T/AD verwenden. Freigabe im Gespräch vom 17.09.2026. Zuordnung nach vollständigem bekannten Merkmalssatz; zusätzliche unbekannte Einschränkungen verhindern eine Verbuchung. Merkmale wie Südfrüchte und Äpfel verschwinden nicht beim Zuordnen.

Kombinationen/Fisch werden intern nach Lieferdatum getrennt gespeichert. Die bestehende Tagesmatrix ist weiterhin Tagesanzeige; fünf Wochenwerte werden niemals in einer Tageszelle summiert. Projektion in die Tagesmatrix mit eigenem gespeichertem Journal und Rücklesen. Automatische Projektion respektiert manuelle Abweichungen.

Weiter belegte Portkorrekturen: ByVal beim Normalisieren verhindert das versehentliche Entfernen von Mail-Zeilenumbrüchen; Cluster-Blocksumme berücksichtigt mehrere Menüs auch bei Tagesbestellungen; Laktosefrei wird als striktes Sonderkostsignal erkannt; Kontrollsumme endet an derselben Zeile und nimmt keine Footer-Zahlen auf. Bekannte Kombinationen zählen nicht mehrfach nach Einzelmerkmalen.

Die historische Mittelwerte-Verknüpfung bleibt auf ausdrückliche Weisung vom 17.09.2026 erhalten. Keine Änderung dieser Formeln oder Referenzwerte.
