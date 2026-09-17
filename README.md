# Küchenversion – KitaFino in VBA

Entwicklung zu Issue #1 auf `codex/autarke-kuechenversion`. Ausgangspunkt ist ausschließlich die bereitgestellte funktionierende MASTER. Sie bleibt unverändert.

## Bedienung

1. `KC_Produktionsplan.xlsm` als Entwicklungskopie öffnen und Makros für diese Datei zulassen. Windows-Excel und klassisches Outlook mit eingerichtetem Postfach werden verwendet.
2. Der bisherige KitaFino-Importbutton und die KitaFino-Navigation öffnen jetzt die Warteschlange. Auf `KitaFino_Warteschlange` stehen Auto **AUS**, Outlook-Überwachung **EIN**, Postfach **Mensa**. Postfachname bei Bedarf in B5 ändern. Die erste Prüfung erfolgt nach etwa fünf Sekunden, danach alle 60 Sekunden; Excel muss dafür bereit sein.
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

V20: Strolche R/S/T; Kita2_4 AC. Zusätzlich vom Auftraggeber freigegeben: Cluster1–4 R (Laktose+Nüsse), S (Laktose+Fructose), T (Nüsse+Soja+Südfrüchte), AD (Gluten+Lactose+Fructose). Jedes Kombinationsergebnis zählt genau ein Essen. Unbekannte/zusätzliche Merkmale bleiben gesperrt. Nicht freigegebene Kombinationen werden vollständig gesperrt; Einschränkungen dürfen bei der Zuordnung nicht verloren gehen.

## Entwicklung und Kontrolle

`src/vba` enthält sechs neue Standardmodule und die vier gezielt ersetzten bestehenden Komponenten. Alle anderen vorhandenen VBA-Komponenten bleiben fachlich erhalten. `docs/BESTAND.md` und `docs/FUNKTIONSVERTRAEGE.md` dokumentieren Ausgangsstand und Portierung.

`tools/build.ps1 -SourcePath <unveränderte MASTER.xlsm> -OutputPath <neue Kopie.xlsm>` baut ausschließlich eine Kopie mittels installierten Windows-Excels. Zugriff auf das VBA-Projekt muss für den **Entwicklungsbau** bereits möglich sein; das Werkzeug ändert keine globalen Sicherheits-/Trust-Center-Einstellungen. Der Küchenbetrieb benötigt diesen Zugriff und dieses Werkzeug nicht.

Regressionstests, Outlook-Referenzen, Abnahmeberichte und die XLSM mit Produktionsdaten bleiben lokal. Dieses Repository enthält die VBA-Quellen, das Bauwerkzeug und technische Dokumentation. Die fachliche Kontrolle gegen die MASTER bleibt vor einer Ablösung im Küchenbetrieb erforderlich.

### Zielwochenfilter

Jede geöffnete Mappe liest nur Mails, deren Lieferdatum zum Original-KW-Blatt und dessen Montag in A8 passt. KW und Jahr müssen stimmen; Wochenbestellungen müssen Montag bis Freitag derselben Woche umfassen. Empfangsdatum und aktive Mappe bestimmen das Ziel nicht. Fremde Wochen werden vor Speicherung von Mailtext, IDs und Queue übersprungen. Eine andere geöffnete, aktualisierte Wochenmappe liest ihre eigenen Mails unabhängig aus Outlook. Geschlossene Zielmappen werden nicht geöffnet; die Mails bleiben in Outlook und werden innerhalb des eingestellten Suchzeitraums später eingelesen. Erneute Prüfung vor Verbuchung verhindert das Schreiben in eine inzwischen geänderte Zielwoche. Auto bleibt standardmäßig AUS.
