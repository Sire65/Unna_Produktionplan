# Getrennte Druckwege für Transport

Der bisherige Mengen-Button änderte Transport.PageSetup.PrintArea dauerhaft auf M5:AB32 und setzte die Seitenanpassung ebenfalls dauerhaft. Transportschein und Druck alles verwendeten Worksheet.PrintOut From:=1, To:=1 und druckten dadurch die erste Seite des zuletzt gespeicherten Bereichs.

Die bestehenden Makronamen DruckeTransportSeite1 und DruckeBereichMegenabfrage werden als bestätigte Aufrufe des gemeinsamen Druckwegs erhalten. DruckMakro_neu, Case TR, und der Transportabschnitt im älteren DruckMakro rufen denselben Tourenplan-Druck auf.

KC_DruckTransport legt pro Auftrag A1:K36 (Tourenplan) oder M5:AB32 (Mengenbewertung) fest, A4, Querformat, eine Seite breit/hoch, ein Exemplar. Die zuvor gelesenen sechs veränderten Seiteneinstellungen werden nach dem Druck oder Export wiederhergestellt. Auch ein Ausgabefehler führt in die Wiederherstellung; deren einzelne Schritte werden trotz eines Fehlers in einer Eigenschaft weiter versucht. Ein Fehler wird an den bestehenden Aufrufer gemeldet.

KC_TransportSetup wird beim Bau einer Entwicklungskopie aufgerufen und setzt den normalen Transport-Druckbereich auf den Tourenplan. Beim bloßen Öffnen werden individuelle Druckeinstellungen nicht erneut überschrieben. Alle übrigen Druckpositionen im Sammeldruck bleiben fachlich unverändert.

Die lokale Prüfung nutzt native Excel-PDF-Exporte desselben Druckwegs. In einer Wegwerfkopie wird ausschließlich der physische PrintOut-Aufruf durch PDF-Export ersetzt; die Sammelauswahl wird auf TR begrenzt. Dadurch lassen sich die realen Button-Makros und der Sammeldruck-Zweig prüfen, ohne Papierdruck auszulösen. Die Prüfung verifiziert Inhalte und Seitenzahlen, visuelle Darstellung, wechselnde Dokumente sowie unveränderte Einstellungen nach Erfolg und Fehler. Es wird keine Ausgabe an einen physischen Drucker als getestet behauptet.
