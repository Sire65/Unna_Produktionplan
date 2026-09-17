# Küchenversion – KitaFino in VBA

Entwicklung zu Issue #1 auf `codex/autarke-kuechenversion`. Ausgangspunkt ist ausschließlich die bereitgestellte funktionierende MASTER. Sie bleibt unverändert.

## Bedienung

1. `KC_Produktionsplan.xlsm` als Entwicklungskopie öffnen und Makros für diese Datei zulassen. Windows-Excel und klassisches Outlook mit eingerichtetem Postfach werden verwendet.
2. Der bisherige KitaFino-Importbutton und die KitaFino-Navigation öffnen jetzt die Warteschlange. Auf `KitaFino_Warteschlange` stehen Auto **AUS**, Outlook-Überwachung **EIN**, Postfach **Mensa**. Postfachname bei Bedarf in Stammdaten!Q33 ändern; Abfrageintervall in Q34 einstellen. Die erste Prüfung erfolgt nach etwa fünf Sekunden, danach alle 60 Sekunden; Excel muss dafür bereit sein.
3. `Jetzt prüfen` liest neue KitaFino-Bestellungen in die interne Queue. Alternativ `Outlook-Auswahl einlesen` oder `MSG-Dateien einlesen`.
4. Eine Queue-Zeile auswählen, `Mail und Zielwerte anzeigen` zur Kontrolle nutzen. `Markierte Bestellung verbuchen` fragt vor dem tatsächlichen Ersetzen vorhandener Werte nach.
5. Fehler und nicht eindeutige Sonderkost sind auch manuell gesperrt. Warnungen wegen vorhandener abweichender Werte oder einer neueren Bestellung verlangen manuelle Kontrolle. Auto lässt ausschließlich vollständig bereite Datensätze zu.
6. Kombinationsfelder und Fisch werden je Lieferdatum intern gespeichert. Zum gewählten Datum in `Produktionsplan!L4` erfolgt eine abgesicherte Übernahme in die bestehende Sonderkostmatrix beim Timerlauf oder mit `Sonderkost für Plantag`. Manuelle Abweichungen werden automatisch nicht ersetzt; der Button bietet eine ausdrückliche Bestätigung an.

Auto wird mit sichtbarem Button und Bestätigung aktiviert. Die Einstellung wird gespeichert; nach Aktivierung bleibt sie über einen Neustart erhalten. Die ausgelieferte Datei hat Auto AUS. Überwachung ist separat abschaltbar und wird beim Schließen beendet.

Keine EXE, kein Installer, keine externen VBS/CMD, kein Compiler und keine externen Queue-Dateien sind für den neuen Import nötig. Outlook-Mails werden nur gelesen. Die bestehende Mittelwerte-Verknüpfung auf Y: bleibt auf ausdrücklichen Wunsch erhalten; die gesamte Arbeitsmappe ist deshalb nicht frei von externen Datenverknüpfungen.

## Speicherung und Sicherheit

`_KC_Config`, `_KC_Q`, `_KC_Body`, `_KC_Log`, `_KC_Txn`, `_KC_Dated` sind VeryHidden. VeryHidden ist eine Bedienhilfe, keine Zugriffssicherung. Die XLSM enthält eingelesene Mailtexte und muss wie die Bestellungen geschützt und gesichert werden.

Die Queue speichert Outlook-EntryID und StoreID, stabile Message-ID soweit vorhanden sowie Originalbetreff/-text. Dubletten werden über Message-ID, EntryID/StoreID oder identischen Inhalt erkannt; historische verbuchte Mail-Keys bleiben berücksichtigt. Bei überlappenden Kunden/Liefertagen werden ältere/zeitgleiche Bestellungen gesperrt und neuere manuell geprüft.

Vor jeder Produktionsänderung werden alte und neue Zielwerte in einem Journal gespeichert und die XLSM erfolgreich gespeichert. Danach werden alle Werte geschrieben und rückgelesen. Ein Fehler setzt die Ziele zurück. Beim nächsten Öffnen wird eine offene Transaktion wiederhergestellt und die betroffene Bestellung gesperrt. Eine beschreibbare, zuverlässig speicherbare XLSM und ihre Datensicherung bleiben erforderlich.

Die datierte Matrix nutzt den Excel-Datumswert als Zeilennummer und `60*(Matrix-Zeile-1)+Matrix-Spalte` als Spaltennummer. Zeilen 2/3 enthalten den letzten importierten Projektionswert bzw. das Datum je Ziel. Nur belegte, verbuchte Tageswerte werden projiziert. Keine Tageswerte bedeutet keine automatische Löschung vorhandener manueller Matrixwerte.

## Freigegebene Kombinationen

V20: Strolche R/S/T; Kita2_4 AC. Zusätzlich vom Auftraggeber freigegeben: Cluster1–4 R (Laktose+Nüsse), S (Laktose+Fructose), T (Nüsse+Soja+Südfrüchte), AD (Gluten+Lactose+Fructose). Jedes Kombinationsergebnis zählt genau ein Essen. Unbekannte/zusätzliche Merkmale verlangen eine ausdrückliche manuelle Matrix-Freigabe; Einschränkungen dürfen bei der Zuordnung nicht verloren gehen.

## Entwicklung und Kontrolle

`src/vba` enthält zehn neue Standardmodule und die sechs gezielt ersetzten bestehenden Komponenten. Alle anderen vorhandenen VBA-Komponenten bleiben fachlich erhalten. `docs/BESTAND.md` und `docs/FUNKTIONSVERTRAEGE.md` dokumentieren Ausgangsstand und Portierung.

`tools/build.ps1 -SourcePath <unveränderte MASTER.xlsm> -OutputPath <neue Kopie.xlsm>` baut ausschließlich eine Kopie mittels installierten Windows-Excels. Zugriff auf das VBA-Projekt muss für den **Entwicklungsbau** bereits möglich sein; das Werkzeug ändert keine globalen Sicherheits-/Trust-Center-Einstellungen. Der Küchenbetrieb benötigt diesen Zugriff und dieses Werkzeug nicht.

Regressionstests, Outlook-Referenzen, Abnahmeberichte und die XLSM mit Produktionsdaten bleiben lokal. Dieses Repository enthält die VBA-Quellen, das Bauwerkzeug und technische Dokumentation. Die fachliche Kontrolle gegen die MASTER bleibt vor einer Ablösung im Küchenbetrieb erforderlich.

### Zielwochenfilter

Jede geöffnete Mappe liest nur Mails, deren Lieferdatum zum Original-KW-Blatt und dessen Montag in A8 passt. KW und Jahr müssen stimmen; Wochenbestellungen müssen Montag bis Freitag derselben Woche umfassen. Empfangsdatum und aktive Mappe bestimmen das Ziel nicht. Fremde Wochen werden vor Speicherung von Mailtext, IDs und Queue übersprungen. Eine andere geöffnete, aktualisierte Wochenmappe liest ihre eigenen Mails unabhängig aus Outlook. Geschlossene Zielmappen werden nicht geöffnet; die Mails bleiben in Outlook und werden innerhalb des eingestellten Suchzeitraums später eingelesen. Erneute Prüfung vor Verbuchung verhindert das Schreiben in eine inzwischen geänderte Zielwoche. Auto bleibt standardmäßig AUS.

Beim Erstellen der nächsten Woche aus einer Kopie müssen KW-Blattname und Montag in A8 zur neuen Woche passen. Ein geänderter Dateiname allein reicht nicht. Geerbte Mails anderer Wochen bleiben als Historie intern erhalten, werden in der aktiven Warteschlange ausgeblendet und von der Automatik übersprungen. Auch die Sonderkostprojektion ist auf Plantage der Zielwoche beschränkt.

Erfolgreich verbuchte Mengen und Sonderkostwerte erhalten wie im ursprünglichen Writer vier kräftige rote Außenkanten (RGB 255,0,0; xlThick), einschließlich ausdrücklich geschriebener Nullwerte. Interne Speicherzellen werden nicht markiert. Die ursprünglichen Rahmen werden vor dem Schreiben im internen Transaktionsjournal gesichert und bei Abbruch/Neustart wiederhergestellt. Dies gilt auch für die datierte Sonderkostprojektion.

Outlook-Abfrageintervall: Stammdaten!Q34, ganze Sekunden von 10 bis 3600; Standard 60. Änderungen gelten ab dem nächsten Timerlauf. Ungültige/pastete Werte behalten das letzte gültige Intervall bei. Einstellung und lokaler Zellname werden mit der Wochenmappe kopiert. Die erste Abfrage ist nach etwa fünf Sekunden vorgesehen. Bei beschäftigtem Excel kann Application.OnTime später ausgeführt werden. Überwachung startet beim Öffnen nur bei Überwachung EIN; Auto bleibt standardmäßig AUS.

Voraussetzungen: Windows mit Desktop-Excel und zugelassenen Makros, klassisches Outlook mit konfiguriertem Postfach. Standard-Postfachname ist Mensa; angepasst wird er in Stammdaten!Q33. Der Zugriff verwendet ausschließlich den Posteingang des eindeutig gefundenen Outlook-Stores. Fehlender oder mehrfach vorkommender Name sperrt den Zugriff, ohne ein anderes Postfach auszuwählen. Bestehende ActiveX-/VBA-Komponenten und die auf Wunsch erhaltene Mittelwerte-Verknüpfung auf Y: begrenzen die allgemeine Portabilität. Keine Zusage für jedes Gerät, Mac, Web-Excel oder neues Outlook.

Postfachname und Abfrageintervall werden zentral in Stammdaten!Q33 bzw. Q34 gespeichert. B5 in der Warteschlange zeigt nur das zuletzt verwendete Postfach an. Änderungen am Postfach setzen beim nächsten Scan die alte Suche zurück. Ein leerer/fehlerhafter Name sperrt jede Abfrage und fällt niemals auf ein anderes Postfach zurück. Migration von Q23 nach Q34 erhält vorhandene individuelle Intervalle und entfernt die alten Import-Einstellungsbeschriftungen.


### Neue Sonderkostformen manuell bestätigen

Bei der manuellen Verbuchung wird für eine eindeutig erkannte, bisher unbekannte Sonderkostform oder Kombination gefragt: „Unbekannte Sonderkostform gefunden. Soll ich sie in die Sonderkostform-Matrix übernehmen?“ Der Dialog zeigt den vollständigen Wortlaut, Kunde, Lieferdatum, Menge und vorgeschlagene freie Zeile; Nein ist vorausgewählt. Nein lässt die betreffende Form und Bestellung unverändert. Bereits einzeln bestätigte Formen bleiben als Stammdaten erhalten, wenn eine später angebotene Form abgelehnt oder die anschließende Buchung abgebrochen wird.

Ja ergänzt ausschließlich eine vollständig leere, unverbundene Zeile im rechten Matrixbereich AK29:AY200. Bestehende Zeilen und die Zuordnungstabelle ab BA werden nicht verschoben. Eine Kombination wird vollständig als eigene Form gespeichert und zählt einmal. Die exakte Zuordnung (nur Groß-/Kleinschreibung und Leerraum normalisiert) wird mit Zeitpunkt/Excel-Benutzer in VeryHidden _KC_Forms gespeichert. Die Kundensummen AE/AF erhalten die zusätzlichen Matrixmengen; die KW-Tagesformeln beziehen AF für den zugehörigen Kunden und den richtigen Tag ein. Vorhandene Summenbasis bleibt erhalten. Nach einer früheren Freigabe manuell veränderte Summen sperren weitere Freigaben.

Neue Bezeichnungen ohne bekannte Allergie-Schlüsselwörter werden ausschließlich innerhalb ausdrücklich beschrifteter, mengenmäßig begrenzter Sonderkost-Detailblöcke erkannt. Speiseplanzeilen werden nicht als neue Formen geraten. Gruppen- und Tageskontrollsummen müssen stimmen. Unbekannter Kunde, falsche Woche, Mengenfehler oder ungültige Zuordnung werden durch die Freigabe nicht aufgehoben. Nach Ja erfolgt eine vollständige erneute Prüfung und weiterhin die normale Buchungsbestätigung. Neue bestätigte Formen bleiben WARNUNG und werden ausschließlich manuell verbucht; die Automatik zeigt keine Freigabe-Dialoge und legt keine Formen an.

_KC_FormTxn speichert vor Änderungen die ursprünglichen Formeln/Werte der Definition und Summen. Ein Fehler oder Neustart stellt eine unvollständige Freigabe wieder her. Verbuchte neue Mengen nutzen denselben datierten Speicher, Rücklesetest, roten Rahmen und Buchungs-Rollback wie bestehende Sonderkost. Ersatzbestellungen schreiben ausdrückliche Nullen. Bei älteren vollständig importierten Kunden-/Tagesbestellungen, die vor Anlage einer neuen Form liegen, gilt diese Form für den betreffenden Tag als null; fremde Tagesmengen werden nicht übernommen.


### Transport und Mengenbewertung drucken

Transportschein druckt ausschließlich den Tourenplan im Blatt Transport (A1:K36, einschließlich Bemerkungsfeld). Mengenbewertung druckt ausschließlich den Bewertungszettel M5:AB32. Beide Dokumente werden als einzelne A4-Seite im Querformat ausgegeben. Die beiden bestehenden Button-Makronamen bleiben erhalten.

Die Auswahl Transport in Druck alles sowie der ältere Sammeldruck verwenden denselben Tourenplan-Druckweg. Es wird nicht die erste Seite einer zufällig zuletzt eingestellten Druckfläche verwendet. Der Druckbereich, Zoom, Seitenanpassung, Ausrichtung und Papierformat werden vor dem einzelnen Auftrag gesichert und danach auch bei Fehlern wiederhergestellt. Eine neu gebaute Entwicklungskopie erhält den Tourenplan als Standard-Druckbereich.

KC_DruckTransport.KC_TransportPrint unterstützt zusätzlich einen ausdrücklichen PDF-Dateipfad für die Prüfung derselben Druckfläche. Native Excel-PDFs, Druckpfadprüfungen und Abnahmeberichte bleiben lokal; sie werden nicht auf GitHub veröffentlicht.
