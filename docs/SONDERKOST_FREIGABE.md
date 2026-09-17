# Manuelle Freigabe neuer Sonderkostformen

## Verhalten

Unbekannte Formen und nicht freigegebene Kombinationen bleiben zunächst nicht buchbar. Die interne Prüfung liefert zusätzlich eine Liste vollständiger Wortlaute mit Kunde, Lieferdatum und Menge. Nur der manuelle Buchungsweg darf für diese Liste einen Ja/Nein-Dialog öffnen. Andere Sperren, einschließlich falscher Woche, unbekanntem Kunden, unplausiblen Mengen oder älteren Bestellungen, gelten weiter.

Ja speichert die vollständige Form in der nächsten vollständig leeren Zeile AK29:AY200. Es werden keine Arbeitsblattzeilen eingefügt. Bestehende Matrixzeilen, Kundenüberschriften und die Zuordnungstabelle ab BA bleiben erhalten. Der Begriff wird als Text gespeichert; eine Kombination ist eine eigene Form und zählt einmal. Ohne freie Zeile oder bei Schutz/Verbindungs-/Formelkonflikten wird keine neue Definition angelegt.

Anschließend wird die ursprüngliche Mail erneut vollständig geprüft. Die normale Bestätigung zur Buchung bleibt erforderlich. Nein verhindert die betreffende Freigabe und Buchung. Bereits bestätigte Definitionen bleiben gespeichert, auch wenn eine weitere Form oder die anschließende Buchung abgelehnt wird.

## Speicherung

KC_NeueSonderkost verwaltet VeryHidden _KC_Forms und _KC_FormTxn. Der Schlüssel normalisiert nur Groß-/Kleinschreibung und Leerraum; es findet keine semantische Vereinfachung oder Entfernung einzelner Einschränkungen statt. Registry und sichtbarer Wortlaut müssen übereinstimmen. Die Registry speichert Matrixzeile, Zeitpunkt, Excel-Benutzer und Status. Eine Definition wird in der internen Protokollierung dokumentiert.

Vor Änderungen werden die vorherigen Formeln/Werte der Definition, der Summenbasis, der AE/AF-Summen, der KW-Tagesverknüpfungen und des Protokolleintrags im separaten Freigabejournal gespeichert und die Datei gespeichert. Eine abgebrochene Definition wird beim Initialisieren wiederhergestellt. Das Journal unterscheidet Formeln und als Text gespeicherte Formelzeichenfolgen.

Die ursprünglichen AE/AF-Ausdrücke werden als Basis erhalten. Zusätzliche Mengen aus den freigegebenen Zeilen werden je Kunden-Spalte ergänzt. Neue Definitionen bauen die Formel aus der gespeicherten Basis und allen aktiven Zeilen auf; es wird nicht bei jeder Definition erneut auf die bisherige Summe addiert. Nach der letzten Freigabe manuell veränderte Summen sperren weitere Freigaben. KW-Tagesformeln verwenden den passenden Kunden und A8 plus Tagesoffset.

## Prüf- und Buchungsweg

Bestehende V20-Zuordnungen bleiben vorrangig erhalten, solange kein exakt manuell freigegebener Wortlaut vorliegt. Neue Bezeichnungen ohne bekannte Sonderkost-Schlüsselwörter werden ausschließlich aus explizit beschrifteten und mengenmäßig begrenzten Detailblöcken übernommen. Die Detailblock-, Tages- und Bestellsummen müssen passen. Aus normalen Speiseplanzutaten werden keine neuen Formen abgeleitet.

Neu freigegebene Formen bleiben im Status WARNUNG und dürfen ausschließlich manuell verbucht werden. Der automatische Weg öffnet keine Dialoge und registriert keine Formen.

Die zusätzlichen Mengen werden je Lieferdatum in _KC_Dated geschrieben und für den Plantag projiziert. Das bestehende Buchungsjournal, Rücklesen, Rollback und rote Rahmen gelten ebenfalls. Ersatzbestellungen setzen nicht mehr bestellte freigegebene Formen ausdrücklich auf null. Bei vor Anlage einer Form bereits vollständig importierten Kunden-/Tagesbestellungen wird eine fehlende neue Form als null behandelt, anhand des existierenden datierten Kunden-Slots für Fisch. Manuelle abweichende Matrixwerte bleiben durch die Projektionskontrolle geschützt.

## Validierung

Lokale native Excel-Tests decken Ja/Nein, nächste freie Zeile, exakten Wortlaut, neue Kombinationen bei Einzel- und Cluster-Kunden, Summen ohne Mehrfachzählung, rote Rahmen, falsche Woche, Mengenfehler, beschädigte Zuordnung, fehlende Zustimmung, geschützte/volle Matrix, manuell veränderte Summen, Tageswechsel, Definition-Rollback und Neustart ab. Zusätzlich werden bestehende Sicherheitsfälle und die lokale 991er Wochenreferenz gegen den neuen Stand geprüft. Originalmails und Produktionsdateien werden nicht veröffentlicht.

Die Definitionen gehören zur jeweiligen Wochenmappe und werden mit deren Kopie übernommen. Es gibt keine zentrale externe Stammdaten- oder Queue-Datei.

