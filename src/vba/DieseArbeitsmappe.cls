Attribute VB_Name = "DieseArbeitsmappe"

Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"

Attribute VB_GlobalNameSpace = False

Attribute VB_Creatable = False

Attribute VB_PredeclaredId = True

Attribute VB_Exposed = True

Attribute VB_TemplateDerived = False

Attribute VB_Customizable = True

Option Explicit



Private Sub Workbook_Open()
    KC_OnOpen

    On Error GoTo FehlerHandler

    

    Dim Einstellung As String

    Dim ws As Worksheet, pt As PivotTable

    Dim cmb As MSForms.comboBox

    Dim startName As String





    ' ==========================================================

    ' 0) Startprotokollierung – MUSS GANZ AM ANFANG laufen

    ' ==========================================================

    On Error Resume Next

    Call Startseite_Anmelden

    On Error GoTo FehlerHandler







    ' ---------------------------------------------------------

    ' 1) Spruch des Tages

    ' ---------------------------------------------------------

    Einstellung = GetSetting("SpruecheApp", "Settings", "Enabled", "True")

    

    If Einstellung = "True" Then

        On Error Resume Next

        frmSpruch.show

        On Error GoTo FehlerHandler

    End If







'    ' ---------------------------------------------------------

'    ' 2) Undo-System initialisieren (ActiveX ToggleButton)

'    ' ---------------------------------------------------------

'    Call InitUndo(True)

'

'    With Sheets("Startseite").tglUndo

'        .value = True

'        .BackColor = RGB(0, 255, 0)

'    End With







    ' ---------------------------------------------------------

    ' 3) Outlook/Termin-ComboBox (ActiveX)

    ' ---------------------------------------------------------

'    Set cmb = ThisWorkbook.Sheets("Startseite").OLEObjects("ComboBoxTag").Object

'    Modul_Termine.InitComboBox cmb

'    FillComboBoxOnce





   

  ' ---------------------------------------------------------

' 4) Protokoll-Button – dynamisch

' ---------------------------------------------------------

    On Error Resume Next

    Application.Run "Mod_Shape_Click_Protokoll.UpdateButtonColorOnOpen"

    On Error GoTo FehlerHandler







 ' ---------------------------------------------------------

' 5) Gruppen-ToggleButtons – dynamisch ausführen

' ---------------------------------------------------------

    On Error Resume Next

    Application.Run "ModGruppenToggle.InitToggleButtons"

    On Error GoTo FehlerHandler







    ' ---------------------------------------------------------

    ' 6) ToggleButton2 (Spruch AN/AUS)

    ' ---------------------------------------------------------

    If WorksheetExists("Startseite") Then

        On Error Resume Next

        With Worksheets("Startseite").ToggleButton2

            .Value = (Einstellung = "True")

        End With

        Worksheets("Startseite").UpdateToggleButtonAppearance

        On Error GoTo FehlerHandler

    End If







    ' ---------------------------------------------------------

    ' 7) Dropdown-Aktualisierung (Stammdaten Q21)

    ' ---------------------------------------------------------

'    If WorksheetExists("Stammdaten") Then

'        Call AktualisiereDropdownHilfsspalte

'    End If







    ' ---------------------------------------------------------

    ' 8) PivotTables + Diagramme aktualisieren

    ' ---------------------------------------------------------

    For Each ws In ThisWorkbook.Worksheets

        For Each pt In ws.PivotTables

            On Error Resume Next

            pt.PivotCache.Refresh

            On Error GoTo FehlerHandler

        Next pt

    Next ws



    ' Pivot-Status nur an eine bereits geladene, sichtbare Spruch-Form senden.
    KC_UpdatePivotLabel

    ' ---------------------------------------------------------

    ' 9) Startblatt automatisch auswählen (Q21)

    ' ---------------------------------------------------------

    If WorksheetExists("Stammdaten") Then

        startName = Sheets("Stammdaten").Range("Q21").Value

    End If

    

    If startName <> "" Then

        If WorksheetExists(startName) Then

            Application.GoTo Sheets(startName).Cells(1, 1), True

        End If

    End If







    ' ---------------------------------------------------------

    ' 10) Uhr-Button beim Start sicher auf "AUS" (rot) setzen

    ' ---------------------------------------------------------

    On Error Resume Next

    SetUhrButtonStatus False

    On Error GoTo FehlerHandler



    ' ---------------------------------------------------------------------------------------------

    ' 11) Home button ein oder ausschalten, ist immer eingebeldnet um stratseite immer zu erreichen

    ' ---------------------------------------------------------------------------------------------

    'Private Sub Workbook_Open()

        Call UpdateHomeButtons

   

 ' ---------------------------------------------------------------------------------------------

    ' 12) Könnte zum Wetter modul gehören, ist nicht ganz klar

    ' ---------------------------------------------------------------------------------------------



    'Public Sub Auto_Open()

        'Call Workbook_Startup

    'End Sub





' ---------------------------------------------------------------------------------------------

    ' 13) Tooltip für X7 in Sheet "Stammdaten" zum manuellen Eingeben von RGB Werten.

    ' ---------------------------------------------------------------------------------------------



'                            Worksheets(SHEET_STAMM).Range("X7").AddComment _
'                                "Manuelle RGB-Eingabe" & vbCrLf & _
'                                "Format: R,G,B (0–255)"

'

'

'

With Worksheets(SHEET_STAMM).Range("X7").Validation

    .Delete

    .Add Type:=xlValidateInputOnly

    .InputTitle = "RGB-Eingabe"

    .InputMessage = "Manuelle RGB-Eingabe möglich!Format: R,G,B (0–255), z.B. 255,0,0 ist knallrot"







End With





Application.EnableEvents = True







    Exit Sub



FehlerHandler:

    MsgBox "Beim Öffnen der Mappe ist ein Fehler aufgetreten: " & Err.Description, vbExclamation, "Hinweis"

End Sub









' =====================================================================

' Eindeutige, konfliktfreie Funktion:

' Prüft, ob ein Arbeitsblatt existiert (ersetzt SheetExists)

' =====================================================================

Public Function WorksheetExists(name As String) As Boolean

    On Error Resume Next

    WorksheetExists = Not ThisWorkbook.Worksheets(name) Is Nothing

    On Error GoTo 0

End Function

















' =====================================================================

' Gehört zum Farbdesigner Modul. Das die Farbe aus Stammdaten in Startseite übernommen wird.

' =====================================================================



'                        Private Sub Workbook_SheetActivate(ByVal Sh As Object)

'                            On Error GoTo EH

'

'                            If Sh.name = "Startseite" Then

'                                mod_FarbDesigner.RefreshPreviewOnStammdaten

'                            End If

'

'                            Exit Sub

'EH:

'                            Debug.Print "Workbook_SheetActivate FEHLER " & Err.Number & ": " & Err.Description

'                        End Sub



Private Sub Workbook_SheetActivate(ByVal sh As Object)

    If sh.name = SHEET_START Then

        mod_FarbDesigner.ApplyDesignToStartseite

     

     End If

    

   'Gehört zum übernhemen der Farben für die KPI´S

   

    If sh.name = SHEET_START Then

    mod_FarbDesigner.ApplyBackgroundShapesToStartseite

End If

    

       

   

End Sub









'==========================================================

' Gehört zur Outlook ComboBox mit "heute, morgen" etc.

'============================================================





Private Sub FillComboBoxOnce()

    On Error Resume Next

    

    Dim ws As Worksheet

    Dim cmb As MSForms.comboBox

    

    Set ws = ThisWorkbook.Sheets("Startseite")

    Set cmb = ws.OLEObjects("ComboBoxTag").Object

    

    If cmb Is Nothing Then Exit Sub

    

    With cmb

        .Clear

        .AddItem "Heute"

        .AddItem "Morgen"

        .AddItem "Übermorgen"

        .AddItem "Woche"

        .AddItem "Monat"

        .Value = "Heute"

    End With

End Sub



'=====================================================================

'Gehört zum Protokolleintrag des angemeldeten Users, Abmeldung

'=====================================================================



Private Sub Workbook_BeforeClose(Cancel As Boolean)
    KC_Stop

    On Error Resume Next

    Call Mod_UserProtokoll.Startseite_Abmelden

End Sub








Public Sub KC_UpdatePivotLabel()
    Dim loadedForm As Object
    For Each loadedForm In VBA.UserForms
        If TypeName(loadedForm) = "frmSpruch" Then
            If loadedForm.Visible Then
                ' Controls-Zugriff bleibt spät gebunden: fehlendes Label blockiert das Öffnen nicht.
                On Error Resume Next
                loadedForm.Controls("lblPivotAbgleich").Caption = "Pivot-Abgleich erfolgreich durchgeführt – " & Format(Now, "HH:mm") & " Uhr"
                loadedForm.Repaint
                On Error GoTo 0
            End If
            Exit For
        End If
    Next loadedForm
End Sub
