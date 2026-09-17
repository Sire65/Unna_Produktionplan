Attribute VB_Name = "KC_Pruefung"
Option Explicit

Public Function KC_Evaluate(ByVal subject As String, ByVal body As String) As Object
    Dim e As Object, plan As Collection, id As String, qty As String, d1 As Variant, d2 As Variant
    Dim targetReason As String
    Dim ws As Worksheet, info As Variant, groups As Long, days As Long, day As Long, gx As Long
    Dim seg As String, q As Variant, values(3) As Long, daySum As Long, total As Long
    Dim labels As Variant, col As Long, base As Long, sk As String, declared As Long, n As Long
    Set e = CreateObject("Scripting.Dictionary"): Set plan = New Collection
    e.Add "plan", plan: e.Add "status", KC_READY: e.Add "note", ""
    e.Add "id", "": e.Add "customer", "": e.Add "date", "": e.Add "end", ""
    e.Add "type", "UNBEKANNT": e.Add "special", "": e.Add "forms", 0: e.Add "bookable", True
    Dim unknown As New Collection
    e.Add "unknown", unknown
    Set KC_Evaluate = e
    On Error GoTo Failed
    id = KC_V20.Rx1(subject, "\((\d{5})\)")
    qty = KC_V20.Rx1(subject, "kitafino\s*-\s*(\d+)\s+Bestellungen")
    d1 = KC_V20.GermanDate(subject): d2 = KC_V20.SecondGermanDate(subject)
    e("id") = id: e("customer") = KC_V20.InstitutionById(id): e("date") = d1
    e("type") = KC_V20.DetectOrderType(subject)
    If IsDate(d2) Then e("end") = d2 Else e("end") = d1
    If id = "" Or e("customer") = "" Then Block e, "Unbekannter Kunde oder KitaFino-ID fehlt"
    If qty = "" Then Block e, "Gesamtmenge fehlt"
    If Len(body) = 0 Then Block e, "Mailtext fehlt"
    If KC_V20.Rx1(body, "BESTELLUNGEN[^\r\n]*\((\d{5})\)") <> id Then Block e, "Kunden-ID im Mailtext fehlt oder widerspricht dem Betreff"
    If Not IsDate(d1) Then Block e, "Lieferdatum fehlt"
    If IsDate(d1) Then
        If Weekday(CDate(d1), vbMonday) > 5 Then Block e, "Lieferdatum liegt am Wochenende"
        If Not ValidSubjectDates(subject) Then Block e, "Kalenderdatum im Betreff ungültig"
    End If
    If Not KC_V20.KC_TargetWeekMatches(ThisWorkbook, subject, targetReason) Then Block e, "Zielwoche passt nicht: " & targetReason
    If e("status") = KC_ERROR Then Exit Function
    If CDbl(qty) > 100000 Or CDbl(qty) < 0 Then Block e, "Gesamtmenge unplausibel": Exit Function
    If Not KC_V20.KC_DetailBlocksValid(body) Then Block e, "Sonderkost-Detailblock widerspricht seiner Gruppen-Kontrollsumme oder ist unvollständig": Exit Function
    sk = KC_V20.SonderkostPreview(body)
    e("special") = KC_V20.SonderkostPretty(sk): e("forms") = KC_V20.SonderkostFormCount(sk)
    groups = 1: days = 1
    Select Case id
        Case "59432": groups = 4: labels = Array("Cluster1", "Cluster2", "Cluster3", "Cluster4")
        Case "59430": groups = 2: labels = Array(KC_V20.LueText(), KC_V20.LueText() & "2")
        Case "59426": groups = 4: labels = Array("Kita1", "Kita2_4")
        Case "59431": labels = Array("Liedbach")
        Case "59500": labels = Array("Strolche")
    End Select
    If IsDate(d2) Then
        days = 5
        If id = "59500" Then Warn e, "Wochenbestellung Strolche im bisherigen V20 nicht freigegeben", False
        If Weekday(CDate(d1), vbMonday) <> 1 Or DateDiff("d", d1, d2) <> 4 Then Block e, "Wochenzeitraum muss Montag bis Freitag sein": Exit Function
    End If
    info = "": Set ws = KC_V20.FindWeekSheetByDate(ThisWorkbook, d1, info)
    If ws Is Nothing Then Block e, "Kein eindeutiges Original-Wochenblatt: " & info: Exit Function
    For day = 0 To days - 1
        seg = KC_V20.WeekDaySegment(body, DateAdd("d", day, d1))
        If seg = "" Then Block e, "Tagessegment fehlt: " & KC_V20.DateText(DateAdd("d", day, d1)): Exit Function
        daySum = 0
        For gx = 1 To groups
            Select Case id
                Case "59432"
                    q = KC_V20.ClusterQtySum(seg, gx)
                Case "59430": q = KC_V20.LuenernWeekQtySum(seg, gx)
                Case "59426": q = KC_V20.GroupQtySum(seg, gx)
                Case Else
                    If days = 5 Then q = KC_V20.SummaryQtySum(seg) Else q = qty
            End Select
            If Not IsNumeric(q) Then Block e, "Gruppenmenge nicht eindeutig: Tag " & day + 1 & ", Gruppe " & gx: Exit Function
            If CDbl(q) < 0 Or CDbl(q) <> Fix(CDbl(q)) Then Block e, "Negative oder gebrochene Gruppenmenge": Exit Function
            values(gx - 1) = CLng(q): daySum = daySum + CLng(q)
        Next gx
        If id = "59426" Then values(1) = values(1) + values(2) + values(3)
        base = KC_V20.DayBaseRow(ws, DateAdd("d", day, d1))
        If base = 0 Then Block e, "Tagesblock fehlt": Exit Function
        For gx = 0 To UBound(labels)
            col = KC_V20.HeaderCol(ws, labels(gx))
            If col = 0 Then Block e, "Kundenspalte fehlt: " & labels(gx): Exit Function
            plan.Add Array(ws.name, ws.Cells(base, col).Address(False, False), values(gx), "Gesamt", DateAdd("d", day, d1), labels(gx))
        Next gx
        KC_SpecialPlan e, seg, id, DateAdd("d", day, d1), ws, base, labels
        q = KC_V20.SummaryQtySum(seg)
        If Not IsNumeric(q) Then
            Block e, "Tageszusammenfassung ohne eindeutige Menge": Exit Function
        ElseIf CLng(q) <> daySum Then
            Block e, "Tageszusammenfassung <> Gruppen-/Betreffsumme": Exit Function
        End If
        total = total + daySum
    Next day
    If total <> CLng(qty) Then Block e, "Summe der Tages-/Gruppenmengen " & total & " <> Betreffmenge " & qty
    If e("status") <> KC_ERROR Then
        KC_Conflicts e
        e("note") = KC_V20.AddN(e("note"), "Summengeprüft: " & total & " Essen; " & KC_V20.SonderkostMealCount(sk) & " Sonderkost-Essen / " & e("forms") & " Formen")
    End If
    Exit Function
Failed:
    Block e, "Auswertung fehlgeschlagen: " & Err.Description
End Function

Private Function ValidSubjectDates(ByVal subject As String) As Boolean
    Dim re As Object, m As Object, month As Long, dt As Date
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True: re.Global = True
    re.pattern = "(\d{1,2})\.\s*(Januar|Februar|März|Maerz|April|Mai|Juni|Juli|August|September|Oktober|November|Dezember)\s*(\d{4})"
    ValidSubjectDates = True
    For Each m In re.Execute(subject)
        month = KC_V20.MonthNo(m.SubMatches(1))
        dt = DateSerial(CLng(m.SubMatches(2)), month, CLng(m.SubMatches(0)))
        If Day(dt) <> CLng(m.SubMatches(0)) Or VBA.month(dt) <> month Then ValidSubjectDates = False
    Next m
End Function

Private Sub Block(e As Object, ByVal note As String)
    e("status") = KC_ERROR: e("bookable") = False
    e("note") = KC_V20.AddN(e("note"), note)
End Sub

Private Sub Warn(e As Object, ByVal note As String, ByVal bookable As Boolean)
    If e("status") <> KC_ERROR Then e("status") = KC_WARN
    If Not bookable Then e("bookable") = False
    If InStr(1, CStr(e("note")), note, vbTextCompare) = 0 Then e("note") = KC_V20.AddN(e("note"), note)
End Sub

Public Sub KC_Conflicts(e As Object)
    Dim item As Variant, cell As Range, plan As Collection, count As Long, examples As String, conflict As Boolean
    Set plan = e("plan")
    For Each item In plan
        Set cell = ThisWorkbook.Worksheets(item(0)).Range(item(1))
        If cell.HasFormula Then Block e, "Ziel enthält Formel: " & item(0) & "!" & item(1): Exit Sub
        If cell.MergeCells Then Block e, "Ziel ist verbunden: " & item(0) & "!" & item(1): Exit Sub
        If cell.Parent.ProtectContents And cell.Locked Then Block e, "Ziel ist geschützt: " & item(0) & "!" & item(1): Exit Sub
        If IsError(cell.Value) Then Block e, "Ziel enthält Fehler": Exit Sub
        conflict = False
        If Len(CStr(cell.Value)) > 0 Then
            If Not IsNumeric(cell.Value) Then
                conflict = True
            ElseIf CDbl(cell.Value) <> CDbl(item(2)) Then
                conflict = True
            End If
        End If
        If conflict Then
            count = count + 1
            If count <= 6 Then examples = examples & IIf(Len(examples) > 0, ", ", "") & item(0) & "!" & item(1)
        End If
    Next item
    If count > 0 Then Warn e, "Vorhandene Werte weichen in " & count & " Zielzellen ab: " & examples, True
End Sub

Private Function Categories(ByVal raw As String) As Collection
    Dim cats As New Collection, l As String
    l = LCase(raw)
    If (InStr(l, "kein schweinefleisch") > 0) Or (InStr(l, "ohne schweinefleisch") > 0) Or l = "moslem" Then cats.Add "Moslem"
    If (InStr(l, "vegetar") > 0) Then cats.Add "Vegetar."
    If (InStr(l, "laktose") > 0) Or (InStr(l, "lactose") > 0) Then cats.Add "Laktose"
    If (InStr(l, "fructos") > 0) Or (InStr(l, "fruktos") > 0) Then cats.Add "Fructose"
    If (InStr(l, "gluten") > 0) Then cats.Add "Gluten"
    If (InStr(l, "kein ei") > 0) Or (InStr(l, "ohne ei") > 0) Then cats.Add "Kein Ei"
    If (InStr(l, "nuss") > 0) Or (InStr(l, "nüss") > 0) Or (InStr(l, "cashew") > 0) Or (InStr(l, "pistaz") > 0) Then cats.Add "Nüsse"
    If (InStr(l, "rind") > 0) Then cats.Add "Rind"
    If (InStr(l, "möhre") > 0) Or (InStr(l, "möhren") > 0) Or (InStr(l, "karotte") > 0) Then cats.Add "Möhre"
    If (InStr(l, "soja") > 0) Then cats.Add "Soja"
    If (InStr(l, "hülsen") > 0) Then cats.Add "Hülsenf."
    If (InStr(l, "fisch") > 0) Then cats.Add "Fisch"
    If InStr(l, "südfr") > 0 Then cats.Add "Südfrüchte"
    If InStr(l, "apfel") > 0 Or InStr(l, "äpfel") > 0 Then cats.Add "Apfel"
    If InStr(l, "pfirs") > 0 Or InStr(l, "pfirsch") > 0 Then cats.Add "Pfirsich"
    If InStr(l, "erdbeer") > 0 Then cats.Add "Erdbeer"
    If InStr(l, "kiwi") > 0 Then cats.Add "Kiwi"
    If InStr(l, "paprika") > 0 Then cats.Add "Paprika"
    If InStr(l, "grünes gemüse") > 0 Then cats.Add "Grünes Gemüse"
    Set Categories = cats
End Function

Private Function KnownLabel(ByVal raw As String) As Boolean
    Dim re As Object, remaining As String
    Set re = CreateObject("VBScript.RegExp"): re.Global = True: re.IgnoreCase = True
    re.Pattern = "\b(?:vegetarisch|vegetarier|moslem|laktose|lactose|fructos(?:e)?|fruktos(?:e)?|gluten|erdnuss|haselnuss|nüsse|nuss|cashew|pistazie|soja|südfrüchte|hülsenfrüchte|schweinefleisch|fisch|rind|eier|ei|möhren|möhre|karotten|karotte|äpfel|apfel|pfirsich|pfirsch|erdbeere|erdbeer|kiwi|paprika|grünes gemüse)(?:allergie|intoleranz|unverträglichkeit|frei)?\b"
    remaining = re.Replace(LCase(raw), "")
    re.Pattern = "\b(?:allergie|auf|und|keine|kein|ohne|alle|insb)\b"
    remaining = re.Replace(remaining, "")
    re.Pattern = "[a-zäöüß]"
    KnownLabel = Not re.Test(remaining)
End Function

Private Function Offset(ByVal cat As String) As Long
    Select Case cat
        Case "Moslem": Offset = 1
        Case "Vegetar.": Offset = 2
        Case "Laktose": Offset = 3
        Case "Fructose": Offset = 4
        Case "Gluten": Offset = 5
        Case "Kein Ei": Offset = 6
        Case "Nüsse": Offset = 7
        Case "Rind": Offset = 8
        Case "Möhre": Offset = 9
        Case "Soja": Offset = 10
        Case "Hülsenf.": Offset = 11
        Case Else: Offset = -1
    End Select
End Function

Private Function ResolveCustomer(ByVal id As String, ByVal context As String) As String
    Select Case id
        Case "59432"
            If KC_V20.RxTest(context, "^Cluster\s*[1-4]$") Then ResolveCustomer = Replace(context, " ", "")
        Case "59426"
            If context = "Gruppe 1" Then ResolveCustomer = "Kita1"
            If context = "Gruppe 2" Or context = "Gruppe 3" Or context = "Gruppe 4" Then ResolveCustomer = "Kita2_4"
        Case "59430"
            If context = KC_V20.LueText() Or context = KC_V20.LueText() & "2" Then ResolveCustomer = context
        Case "59431": ResolveCustomer = "Liedbach"
        Case "59500": ResolveCustomer = "Strolche"
    End Select
End Function

Private Function ComboColumn(ByVal customer As String, ByVal raw As String, cats As Collection) As Long
    Dim l As String
    l = LCase(raw)
    ' Count exact known feature sets; extra restrictions never silently disappear.
    If customer = "Kita2_4" And cats.Count = 3 Then
        If (InStr(l, "schweinefleisch") > 0) And (InStr(l, "hülsen") > 0) And ((InStr(l, "nuss") > 0) Or (InStr(l, "nüss") > 0)) Then ComboColumn = 29
    End If
    If customer <> "Strolche" And Left(customer, 7) <> "Cluster" Then Exit Function
    If cats.Count = 2 Then
        If ((InStr(l, "laktose") > 0) Or (InStr(l, "lactose") > 0)) And ((InStr(l, "nuss") > 0) Or (InStr(l, "nüss") > 0)) Then ComboColumn = 18
        If ((InStr(l, "laktose") > 0) Or (InStr(l, "lactose") > 0)) And ((InStr(l, "fructos") > 0) Or (InStr(l, "fruktos") > 0)) Then ComboColumn = 19
    ElseIf cats.Count = 3 Then
        If ((InStr(l, "laktose") > 0) Or (InStr(l, "lactose") > 0)) And ((InStr(l, "fructos") > 0) Or (InStr(l, "fruktos") > 0)) And (InStr(l, "gluten") > 0) And Left(customer, 7) = "Cluster" Then ComboColumn = 30
        If ((InStr(l, "nuss") > 0) Or (InStr(l, "nüss") > 0)) And (InStr(l, "soja") > 0) And (InStr(l, "südfr") > 0) Then ComboColumn = 20
    End If
End Function

Private Function CustomerRow(ByVal customer As String) As Long
    Dim ws As Worksheet, r As Long, label As String
    Set ws = KC_Sheet("Sonderkostformen")
    For r = 5 To 21
        label = CStr(ws.Cells(r, 1).Value)
        If label = customer Or (customer = "Liedbach" And label = "Liedbach1") Then
            If CustomerRow > 0 Then CustomerRow = 0: Exit Function
            CustomerRow = r
        End If
    Next r
End Function

Private Sub KC_SpecialPlan(e As Object, ByVal seg As String, ByVal id As String, ByVal dt As Date, kw As Worksheet, ByVal base As Long, labels As Variant)
    Dim sk As String, p As Variant, context As String, raw As String, customer As String, count As Long
    Dim cats As Collection, off As Long, col As Long, combCol As Long, cr As Long, learnedRow As Long, fr As Long, registryRow As Long
    Dim plan As Collection, sums As Object, key As String, keys As Variant, k As Variant, a As Variant
    Dim sourceMeals As Long, plannedMeals As Long, declared As Long, re As Object, matches As Object, m As Object
    Dim gx As Long, i As Long, currentDate As Variant, canMatrix As Boolean, skSeg As String
    Set plan = e("plan"): Set sums = CreateObject("Scripting.Dictionary")
    skSeg = seg
    If id = "59500" Then
        Set re = CreateObject("VBScript.RegExp"): re.Global = True: re.MultiLine = True
        re.pattern = "^[AB]-[^\r\n]+Anzahl[^\r\n]*"
        skSeg = re.Replace(skSeg, "Strolche Anzahl")
    End If
    sk = KC_V20.SonderkostPreview(skSeg)
    declared = 0
    Set re = CreateObject("VBScript.RegExp"): re.IgnoreCase = True: re.Global = True
    re.pattern = "Davon[ \t]+Ern.hrungsbesonderheiten![ \t]*Details[^\r\n]*?[ \t]+(\d+)[ \t]*(?:\r|\n|$)"
    Set matches = re.Execute(seg)
    If matches.Count <> 1 Then Block e, "Sonderkost-Kontrollsumme der Tageszusammenfassung fehlt/mehrdeutig": Exit Sub
    declared = CLng(matches(0).SubMatches(0))
    sourceMeals = KC_V20.SonderkostMealCount(sk)
    If declared <> sourceMeals Then Block e, "Sonderkost-Summe " & sourceMeals & " <> Tageskontrollsumme " & declared: Exit Sub
    currentDate = KC_Sheet("Produktionsplan").Range("L4").Value
    If IsDate(currentDate) Then canMatrix = (DateValue(currentDate) = dt)
    ' Complete replacement includes explicit zeros for standard fields.
    For gx = 0 To UBound(labels)
        col = KC_V20.HeaderCol(kw, labels(gx))
        For i = 1 To 11
            key = kw.name & "|" & kw.Cells(base + i, col).Address(False, False)
            sums(key) = 0
        Next i
    Next gx
    If Len(sk) > 0 Then
        For Each p In Split(sk, " / ")
            context = Left(CStr(p), InStr(CStr(p), ":") - 1)
            raw = Mid(CStr(p), InStr(CStr(p), ":") + 1)
            raw = Trim(Left(raw, InStrRev(raw, "[") - 1))
            count = CLng(KC_V20.Rx1(CStr(p), "\[(\d+)\]\s*$"))
            customer = ResolveCustomer(id, context)
            If customer = "" Then Warn e, "Sonderkost-Kunde nicht eindeutig: " & p, False: GoTo NextForm
            learnedRow = KC_FormRow(raw)
            If learnedRow > 0 Then
                col = MatrixColumn(customer)
                If col = 0 Then Block e, "Sonderkost-Matrixkunde fehlt": Exit Sub
                key = "Sonderkostformen|" & KC_Sheet("Sonderkostformen").Cells(learnedRow, col).Address(False, False)
                Warn e, "Manuell bestätigte neue Sonderkostform (nur manuell verbuchen): " & raw, True
                GoTo CountForm
            End If
            If Not KnownLabel(raw) Then OfferForm e, raw, customer, dt, count: GoTo NextForm
            Set cats = Categories(raw)
            If cats.Count = 0 Then OfferForm e, raw, customer, dt, count: GoTo NextForm
            col = KC_V20.HeaderCol(kw, customer)
            If col = 0 Then Block e, "Sonderkost-Zielspalte fehlt": Exit Sub
            If cats.Count > 1 Then
                combCol = ComboColumn(customer, raw, cats)
                If combCol = 0 Then OfferForm e, raw, customer, dt, count: GoTo NextForm
                cr = CustomerRow(customer)
                If cr = 0 Then Block e, "Kombinations-Kundenzeile nicht eindeutig": Exit Sub
                key = "Sonderkostformen|" & KC_Sheet("Sonderkostformen").Cells(cr, combCol).Address(False, False)
            Else
                off = Offset(cats(1))
                If off > 0 Then
                    key = kw.name & "|" & kw.Cells(base + off, col).Address(False, False)
                ElseIf cats(1) = "Fisch" Then
                    col = MatrixColumn(customer): cr = MatrixRow("Fisch")
                    If col = 0 Or cr = 0 Then Block e, "Fisch-Matrixziel fehlt": Exit Sub
                    key = "Sonderkostformen|" & KC_Sheet("Sonderkostformen").Cells(cr, col).Address(False, False)
                Else
                    OfferForm e, raw, customer, dt, count: GoTo NextForm
                End If
            End If
CountForm:
            If Not sums.Exists(key) Then sums.Add key, 0
            sums(key) = CLng(sums(key)) + count
            plannedMeals = plannedMeals + count
NextForm:
        Next p
    End If
    If True Then
        For gx = 0 To UBound(labels)
            cr = CustomerRow(labels(gx))
            If cr > 0 Then
                If labels(gx) = "Strolche" Or Left(labels(gx), 7) = "Cluster" Then
                    For i = 18 To 20
                        key = "Sonderkostformen|" & KC_Sheet("Sonderkostformen").Cells(cr, i).Address(False, False)
                        If Not sums.Exists(key) Then sums.Add key, 0
                    Next i
                    If Left(labels(gx), 7) = "Cluster" Then
                        key = "Sonderkostformen|" & KC_Sheet("Sonderkostformen").Cells(cr, 30).Address(False, False)
                        If Not sums.Exists(key) Then sums.Add key, 0
                    End If
                ElseIf labels(gx) = "Kita2_4" Then
                    key = "Sonderkostformen|" & KC_Sheet("Sonderkostformen").Cells(cr, 29).Address(False, False)
                    If Not sums.Exists(key) Then sums.Add key, 0
                End If
            End If
            col = MatrixColumn(labels(gx)): cr = MatrixRow("Fisch")
            If col > 0 And cr > 0 Then
                key = "Sonderkostformen|" & KC_Sheet("Sonderkostformen").Cells(cr, col).Address(False, False)
                If Not sums.Exists(key) Then sums.Add key, 0
            End If
        Next gx
    End If
    ' Explicit per-date zeros also remove a newly learned form from a replacement order.
    For registryRow = 2 To KC_Sheet("_KC_Forms").Cells(KC_Sheet("_KC_Forms").Rows.Count, 1).End(xlUp).Row
        fr = KC_FormRow(CStr(KC_Sheet("_KC_Forms").Cells(registryRow, 2).Value2))
        For gx = 0 To UBound(labels)
            col = MatrixColumn(CStr(labels(gx)))
            If col = 0 Then Block e, "Sonderkost-Matrixkunde fehlt": Exit Sub
            key = "Sonderkostformen|" & KC_Sheet("Sonderkostformen").Cells(fr, col).Address(False, False)
            If Not sums.Exists(key) Then sums.Add key, 0
        Next gx
    Next registryRow
    For Each k In sums.Keys
        a = Split(CStr(k), "|")
        If a(0) = "Sonderkostformen" Then
            Dim index As Long
            index = KC_DatedIndex(CStr(a(1)))
            plan.Add Array("_KC_Dated", KC_Sheet("_KC_Dated").Cells(CLng(dt), index).Address(False, False), sums(k), "Sonderkost (datiert)", dt, CStr(a(1)))
            If canMatrix Then
                plan.Add Array(a(0), a(1), sums(k), "Sonderkost (Plantag)", dt, "")
                plan.Add Array("_KC_Dated", KC_Sheet("_KC_Dated").Cells(2, index).Address(False, False), sums(k), "Projektionsnachweis", dt, "")
                plan.Add Array("_KC_Dated", KC_Sheet("_KC_Dated").Cells(3, index).Address(False, False), CLng(dt), "Projektionsnachweis", dt, "")
            End If
        Else
            plan.Add Array(a(0), a(1), sums(k), "Sonderkost", dt, "")
        End If
    Next k
    If sourceMeals <> plannedMeals Then Warn e, "Sonderkost-Schreibplan unvollständig: " & plannedMeals & "/" & sourceMeals, False
End Sub

Private Function MatrixColumn(ByVal customer As String) As Long
    Dim ws As Worksheet, c As Long, s As String
    Set ws = KC_Sheet("Sonderkostformen")
    For c = 38 To 50
        s = CStr(ws.Cells(4, c).Value)
        If s = customer Or (customer = "Liedbach" And s = "Liedbach1") Then MatrixColumn = c: Exit Function
    Next c
End Function

Private Function MatrixRow(ByVal feature As String) As Long
    Dim ws As Worksheet, r As Long
    Set ws = KC_Sheet("Sonderkostformen")
    For r = 5 To 28
        If CStr(ws.Cells(r, 37).Value) = feature Then MatrixRow = r: Exit Function
    Next r
End Function



Private Sub OfferForm(e As Object, ByVal raw As String, ByVal customer As String, ByVal dt As Date, ByVal count As Long)
    Dim unknown As Collection
    If MatrixColumn(customer) = 0 Then Block e, "Sonderkost-Matrixkunde fehlt": Exit Sub
    Set unknown = e("unknown")
    unknown.Add Array(raw, customer, dt, count)
    Warn e, "Unbekannte Sonderkostform/Kombination – manuelle Übernahme erforderlich: " & raw, False
End Sub
