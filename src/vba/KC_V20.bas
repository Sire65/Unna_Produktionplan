Attribute VB_Name = "KC_V20"
Option Explicit

Public Function Rx1(t,p)
    Dim r,m
    Set r=CreateObject("VBScript.RegExp")
    r.IgnoreCase=True:r.Global=False:r.Pattern=p
    If r.Test(t) Then Set m=r.Execute(t)(0):Rx1=m.SubMatches(0) Else Rx1=""
End Function

Public Function AddN(a,b)
    If Trim(CStr(a))="" Then AddN=b Else AddN=a  &  "; "  &  b
End Function

Public Function FlatText(ByVal s As Variant)
    s=Replace(CStr(s),vbCr," ")
    s=Replace(s,vbLf," ")
    s=Replace(s,vbTab," ")
    Do While InStr(s,"  ")>0
        s=Replace(s,"  "," ")
    Loop
    FlatText=Trim(s)
End Function

Public Function GermanDate(t)
    Dim r,m,d,y,mo,mn
    Set r=CreateObject("VBScript.RegExp")
    r.IgnoreCase=True:r.Global=True
    r.Pattern="(\d{1,2})\.\s*(Januar|Februar|März|Maerz|April|Mai|Juni|Juli|August|September|Oktober|November|Dezember)\s*(\d{4})"
    If Not r.Test(t) Then GermanDate="":Exit Function
    Set m=r.Execute(t)(0)
    d=CLng(m.SubMatches(0)):mo=LCase(m.SubMatches(1)):y=CLng(m.SubMatches(2))
    mn=MonthNo(mo):GermanDate=DateSerial(y,mn,d)
End Function

Public Function SecondGermanDate(t)
    Dim r,ms,m,d,y,mo,mn
    Set r=CreateObject("VBScript.RegExp")
    r.IgnoreCase=True:r.Global=True
    r.Pattern="(\d{1,2})\.\s*(Januar|Februar|März|Maerz|April|Mai|Juni|Juli|August|September|Oktober|November|Dezember)\s*(\d{4})"
    If Not r.Test(t) Then SecondGermanDate="":Exit Function
    Set ms=r.Execute(t)
    If ms.Count<2 Then SecondGermanDate="":Exit Function
    Set m=ms(1)
    d=CLng(m.SubMatches(0)):mo=LCase(m.SubMatches(1)):y=CLng(m.SubMatches(2))
    mn=MonthNo(mo):SecondGermanDate=DateSerial(y,mn,d)
End Function

Public Function MonthNo(mo)
    Select Case LCase(mo)
        Case "januar":MonthNo=1
        Case "februar":MonthNo=2
        Case "märz","maerz":MonthNo=3
        Case "april":MonthNo=4
        Case "mai":MonthNo=5
        Case "juni":MonthNo=6
        Case "juli":MonthNo=7
        Case "august":MonthNo=8
        Case "september":MonthNo=9
        Case "oktober":MonthNo=10
        Case "november":MonthNo=11
        Case "dezember":MonthNo=12
    End Select
End Function

Public Function DateText(d)
    If IsDate(d) Then
        DateText=Right("0" & Day(d),2) & "." & Right("0" & Month(d),2) & "." & Year(d)
    Else
        DateText=""
    End If
End Function

Public Function DetectOrderType(subj)
    Dim a,b
    a=GermanDate(subj):b=SecondGermanDate(subj)
    If IsDate(a) And IsDate(b) Then
        DetectOrderType="WOCHENBESTELLUNG"
    ElseIf IsDate(a) Then
        DetectOrderType="TAGESBESTELLUNG"
    Else
        DetectOrderType="UNBEKANNT"
    End If
End Function

Public Function DetectOrderRange(subj)
    Dim a,b
    a=GermanDate(subj):b=SecondGermanDate(subj)
    If IsDate(a) And IsDate(b) Then
        DetectOrderRange=DateText(a) & " - " & DateText(b)
    ElseIf IsDate(a) Then
        DetectOrderRange=DateText(a)
    Else
        DetectOrderRange=""
    End If
End Function

Public Function NumberFromCell(v)
    Dim r,m,s
    s=CStr(v):Set r=CreateObject("VBScript.RegExp")
    r.Global=False:r.Pattern="(\d+)"
    If r.Test(s) Then Set m=r.Execute(s)(0):NumberFromCell=CLng(m.SubMatches(0)) Else NumberFromCell=0
End Function

Public Function IsWeekCopySheet(sheetName)
    Dim s
    s=LCase(Trim(CStr(sheetName)))
    IsWeekCopySheet=False
    If Right(s,2)=" a" Then IsWeekCopySheet=True:Exit Function
    If InStr(1,s,"kopie",vbTextCompare)>0 Then IsWeekCopySheet=True:Exit Function
    If InStr(1,s,"backup",vbTextCompare)>0 Then IsWeekCopySheet=True:Exit Function
    If InStr(1,s,"tempprint",vbTextCompare)>0 Then IsWeekCopySheet=True:Exit Function
    If InStr(1,s,"tmp_print",vbTextCompare)>0 Then IsWeekCopySheet=True:Exit Function
    If InStr(1,s,"tmpprint",vbTextCompare)>0 Then IsWeekCopySheet=True:Exit Function
    If Left(s,4)="tmp_" Then IsWeekCopySheet=True:Exit Function
    If Left(s,5)="temp_" Then IsWeekCopySheet=True:Exit Function
    If InStr(1,s,"druck",vbTextCompare)>0 Then IsWeekCopySheet=True:Exit Function
    If InStr(1,s,"print",vbTextCompare)>0 Then IsWeekCopySheet=True:Exit Function
    If Not IsRealWeekSheetName(s) Then IsWeekCopySheet=True
End Function

Public Function IsRealWeekSheetName(s)
    Dim r
    Set r=CreateObject("VBScript.RegExp")
    r.IgnoreCase=True:r.Global=False:r.Pattern="^kw\s*\d*$"
    IsRealWeekSheetName=r.Test(Trim(CStr(s)))
End Function

Public Function WeekStartFromSheet(ws)
    Dim v
    WeekStartFromSheet = ""
    v = ws.Range("A8").Value
    If IsDate(v) Then
        If Weekday(CDate(v), vbMonday) = 1 Then WeekStartFromSheet = DateValue(v)
    End If
End Function

Public Function FindWeekSheetByDate(wb,d,ByRef info)
    Dim ws,wsFound,n,mon,candidates,kwName,kwNo
    Set FindWeekSheetByDate=Nothing:Set wsFound=Nothing
    n=0:candidates=""
    If Not IsDate(d) Then info="Datum ungültig":Exit Function

    ' Primär: Original-Wochenblatt über ISO-KW im Blattnamen finden.
    kwNo=DatePart("ww",DateValue(d),vbMonday,vbFirstFourDays)
    kwName="KW "  &  CStr(kwNo)
    For Each ws In wb.Worksheets
        If Not IsWeekCopySheet(ws.Name) Then
            If StrComp(Trim(CStr(ws.Name)),kwName,vbTextCompare)=0 Then
                n=n+1:Set wsFound=ws
                If candidates<>"" Then candidates=candidates  &  ", "
                candidates=candidates  &  ws.Name
            End If
        End If
    Next
    If n=1 Then
        mon=WeekStartFromSheet(wsFound)
        If Not IsDate(mon) Then info="Wochenstart fehlt": Exit Function
        If DateValue(d)<DateValue(mon) Or DateValue(d)>DateAdd("d",4,DateValue(mon)) Then info="KW-Jahr/Datum passt nicht": Exit Function
        Set FindWeekSheetByDate=wsFound
        If IsDate(mon) Then
            info="1 Treffer: "  &  wsFound.Name  &  "["  &  DateText(mon)  &  "-"  &  DateText(DateAdd("d",4,mon))  &  "]"
        Else
            info="1 Treffer: "  &  wsFound.Name  &  " [KW-Namensuche]"
        End If
        Exit Function
    ElseIf n>1 Then
        info=CStr(n)  &  " Treffer per KW-Name: "  &  candidates
        Exit Function
    End If

    ' Fallback: Datumsbereich im Blattkopf prüfen.
    n=0:candidates="":Set wsFound=Nothing
    For Each ws In wb.Worksheets
        If Not IsWeekCopySheet(ws.Name) Then
            mon=WeekStartFromSheet(ws)
            If IsDate(mon) Then
                If DateValue(d)>=DateValue(mon) And DateValue(d)<=DateAdd("d",4,DateValue(mon)) Then
                    n=n+1:Set wsFound=ws
                    If candidates<>"" Then candidates=candidates  &  ", "
                    candidates=candidates  &  ws.Name  &  "["  &  DateText(mon)  &  "-"  &  DateText(DateAdd("d",4,mon))  &  "]"
                End If
            End If
        End If
    Next

    If n=1 Then
        Set FindWeekSheetByDate=wsFound
        info="1 Treffer: "  &  candidates
    ElseIf n=0 Then
        info="0 Treffer (erwartet: "  &  kwName  &  ")"
    Else
        info=CStr(n)  &  " Treffer: "  &  candidates
    End If
End Function

Public Function HeaderCol(ws,want)
    Dim c,h
    HeaderCol=0
    For c=3 To 15
        h=Trim(CStr(ws.Cells(3,c).Value))
        If StrComp(h,want,vbTextCompare)=0 Then HeaderCol=c:Exit Function
        If StrComp(want,"Liedbach1",vbTextCompare)=0 And StrComp(h,"Liedbach",vbTextCompare)=0 Then HeaderCol=c:Exit Function
    Next
End Function

Public Function DayBaseRow(ws,d)
    Dim mon,off
    DayBaseRow=0
    mon=WeekStartFromSheet(ws)
    If Not IsDate(mon) Or Not IsDate(d) Then Exit Function
    off=DateDiff("d",DateValue(mon),DateValue(d))
    If off>=0 And off<=4 Then DayBaseRow=4+(off*13)
End Function

Public Function ColLetter(colNo)
    Dim n,s
    n=CLng(colNo):s=""
    Do While n>0
        n=n-1:s=Chr(65+(n Mod 26)) & s:n=Int(n/26)
    Loop
    ColLetter=s
End Function

Public Function LueText()
    LueText="L"  &  ChrW(252)  &  "nern"
End Function

Public Function InstitutionById(id)
    Select Case CStr(id)
        Case "59432":InstitutionById="Ganztag am Hertinger Tor Unna"
        Case "59431":InstitutionById="Ganztag Liedbachschule Unna"
        Case "59500":InstitutionById="Familienzentrum Vorstadtstrolche Unna"
        Case "59430":InstitutionById="OGS Grundschule "  &  LueText()  &  " Unna"
        Case "59426":InstitutionById="Familienzentrum Hertinger Tor"
        Case Else:InstitutionById=""
    End Select
End Function

Public Function TargetLabel(id)
    Select Case CStr(id)
        Case "59432":TargetLabel="Cluster1-4"
        Case "59431":TargetLabel="Liedbach1"
        Case "59500":TargetLabel="Strolche"
        Case "59430":TargetLabel=LueText()  &  "/"  &  LueText()  &  "2"
        Case "59426":TargetLabel="Kita1/Kita2_4"
        Case Else:TargetLabel=""
    End Select
End Function


Public Function WeekDaySegment(body,d)
    Dim a,b,p1,p2,mark1,mark2
    a=Replace(Replace(body,vbCrLf,vbLf),vbCr,vbLf)
    mark1="Zusammenfassung " & Day(d) & ". " & GermanMonthName(Month(d)) & " " & Year(d)
    p1=InStr(1,a,mark1,vbTextCompare)
    If p1=0 Then WeekDaySegment="":Exit Function
    If Weekday(d,2)<5 Then
        mark2="Zusammenfassung " & Day(DateAdd("d",1,d)) & ". " & GermanMonthName(Month(DateAdd("d",1,d))) & " " & Year(DateAdd("d",1,d))
        p2=InStr(p1+Len(mark1),a,mark2,vbTextCompare)
    Else
        ' FIX5L: Freitag am fruehesten Footer-Marker beenden.
        Dim fp,ft
        p2=0
        For Each ft In Array("[Dies ist","Dies ist eine automatisch generierte","automatisch generierte E-Mail","kitafino GmbH","kontakt@kitafino","Donaustra","Amtsgericht")
            fp=InStr(p1+Len(mark1),a,CStr(ft),vbTextCompare)
            If fp>0 Then
                If p2=0 Or fp<p2 Then p2=fp
            End If
        Next
    End If
    If p2=0 Then p2=Len(a)+1
    WeekDaySegment=Mid(a,p1,p2-p1)
End Function

Public Function GermanMonthName(m)
    Dim a:a=Array("","Januar","Februar","Maerz","April","Mai","Juni","Juli","August","September","Oktober","November","Dezember")
    GermanMonthName=a(m)
End Function

Public Function ClusterQtySum(seg,n)
    Dim lines,i,line,norm,inCluster,totalNums,davonQty,pendingDavon,r,curRe,nextRe,tailRe,davonRe,m,v
    lines=Split(Replace(Replace(seg,vbCrLf,vbLf),vbCr,vbLf),vbLf)
    inCluster=False:totalNums=0:davonQty=0:pendingDavon=False
    Set curRe=CreateObject("VBScript.RegExp"):curRe.IgnoreCase=True:curRe.Global=False:curRe.Pattern="^\s*Cluster\s*" & n & "(\s+Anzahl)?\s*$"
    Set nextRe=CreateObject("VBScript.RegExp"):nextRe.IgnoreCase=True:nextRe.Global=False:nextRe.Pattern="^\s*Cluster\s*[1-4](\s+Anzahl)?\s*$"
    Set tailRe=CreateObject("VBScript.RegExp"):tailRe.IgnoreCase=True:tailRe.Global=False:tailRe.Pattern="(?:^|[^0-9])(\d{1,4})\s*$"
    Set davonRe=CreateObject("VBScript.RegExp"):davonRe.IgnoreCase=True:davonRe.Global=False:davonRe.Pattern="Davon\s+Ern.hrungsbesonderheiten.*?:\s*(\d+)\s*$"
    For i=0 To UBound(lines)
        line=Trim(CStr(lines(i)))
        norm=Replace(line,vbTab," ")
        Do While InStr(norm,"  ")>0:norm=Replace(norm,"  "," "):Loop
        If Not inCluster Then
            If curRe.Test(norm) Or InStr(1,norm,"Cluster " & n & " Anzahl",vbTextCompare)=1 Then inCluster=True
        Else
            If nextRe.Test(norm) And Not curRe.Test(norm) Then Exit For
            If InStr(1,norm,"automatisch generierte",vbTextCompare)>0 _
               Or InStr(1,norm,"kitafino GmbH",vbTextCompare)>0 _
               Or InStr(1,norm,"kontakt@kitafino",vbTextCompare)>0 _
               Or InStr(1,norm,"Donaustra",vbTextCompare)>0 _
               Or InStr(1,norm,"Amtsgericht",vbTextCompare)>0 Then Exit For
            If InStr(1,norm,"Davon Ern",vbTextCompare)>0 Then
                If davonRe.Test(norm) Then
                    Set m=davonRe.Execute(norm)(0):davonQty=CLng(m.SubMatches(0)):pendingDavon=False
                Else
                    pendingDavon=True
                End If
            End If
            If tailRe.Test(norm) Then
                Set m=tailRe.Execute(norm)(0):v=CLng(m.SubMatches(0)):totalNums=totalNums+v
                If pendingDavon Then davonQty=v:pendingDavon=False
            End If
        End If
    Next
    If Not inCluster Then
        ClusterQtySum=""
    ElseIf totalNums-(2*davonQty)<0 Then
        ClusterQtySum=""
    Else
        ClusterQtySum=CStr(totalNums-(2*davonQty))
    End If
End Function

Public Function GroupQtySum(seg,n)
    GroupQtySum=BlockQtySum(seg,"^\s*Gruppe\s*" & n & "(\s+Anzahl)?\s*$","^\s*Gruppe\s*[1-4](\s+Anzahl)?\s*$")
End Function

Public Function LuenernWeekQtySum(seg,n)
    LuenernWeekQtySum=BlockQtySum(seg,"^.*nern\s*" & n & ".*Anzahl\s*$","^.*nern\s*[12].*Anzahl\s*$")
End Function

Public Function BlockQtySum(seg,startPattern,boundaryPattern)
    Dim lines,i,line,norm,inside,totalNums,davonQty,pendingDavon,startRe,boundRe,tailRe,davonRe,m,v
    lines=Split(Replace(Replace(seg,vbCrLf,vbLf),vbCr,vbLf),vbLf)
    inside=False:totalNums=0:davonQty=0:pendingDavon=False
    Set startRe=CreateObject("VBScript.RegExp"):startRe.IgnoreCase=True:startRe.Global=False:startRe.Pattern=startPattern
    Set boundRe=CreateObject("VBScript.RegExp"):boundRe.IgnoreCase=True:boundRe.Global=False:boundRe.Pattern=boundaryPattern
    Set tailRe=CreateObject("VBScript.RegExp"):tailRe.IgnoreCase=True:tailRe.Global=False:tailRe.Pattern="(?:^|[^0-9])(\d{1,4})\s*$"
    Set davonRe=CreateObject("VBScript.RegExp"):davonRe.IgnoreCase=True:davonRe.Global=False:davonRe.Pattern="Davon\s+Ern.hrungsbesonderheiten.*?:\s*(\d+)\s*$"
    For i=0 To UBound(lines)
        line=Trim(CStr(lines(i))):norm=Replace(line,vbTab," ")
        Do While InStr(norm,"  ")>0:norm=Replace(norm,"  "," "):Loop
        If Not inside Then
            If startRe.Test(norm) Then inside=True
        Else
            If boundRe.Test(norm) And Not startRe.Test(norm) Then Exit For
            If InStr(1,norm,"automatisch generierte",vbTextCompare)>0 Or InStr(1,norm,"kitafino GmbH",vbTextCompare)>0 Or InStr(1,norm,"Amtsgericht",vbTextCompare)>0 Then Exit For
            If InStr(1,norm,"Davon Ern",vbTextCompare)>0 Then
                If davonRe.Test(norm) Then
                    Set m=davonRe.Execute(norm)(0):davonQty=CLng(m.SubMatches(0)):pendingDavon=False
                Else
                    pendingDavon=True
                End If
            End If
            If tailRe.Test(norm) Then
                Set m=tailRe.Execute(norm)(0):v=CLng(m.SubMatches(0)):totalNums=totalNums+v
                If pendingDavon Then davonQty=v:pendingDavon=False
            End If
        End If
    Next
    If Not inside Or totalNums-(2*davonQty)<0 Then BlockQtySum="" Else BlockQtySum=CStr(totalNums-(2*davonQty))
End Function

Public Function SummaryQtySum(seg)
    Dim lines,i,line,norm,total,tailRe,m
    lines=Split(Replace(Replace(seg,vbCrLf,vbLf),vbCr,vbLf),vbLf)
    total=0
    Set tailRe=CreateObject("VBScript.RegExp"):tailRe.IgnoreCase=True:tailRe.Global=False:tailRe.Pattern="(?:^|[^0-9])(\d{1,4})\s*$"
    For i=1 To UBound(lines)
        line=Trim(CStr(lines(i))):norm=Replace(line,vbTab," ")
        If InStr(1,norm,"Davon Ern",vbTextCompare)>0 Then Exit For
        If InStr(1,norm,"Gruppe ",vbTextCompare)=1 Or InStr(1,norm,"Liedbachschule",vbTextCompare)>0 Then Exit For
        If tailRe.Test(norm) Then
            Set m=tailRe.Execute(norm)(0):total=total+CLng(m.SubMatches(0))
        End If
    Next
    If total<=0 Then SummaryQtySum="" Else SummaryQtySum=CStr(total)
End Function

Public Function QtyAfterLabel(body,labelPattern)
    Dim r,m,x
    x=FlatText(body)
    Set r=CreateObject("VBScript.RegExp")
    r.IgnoreCase=True:r.Global=False
    r.Pattern=labelPattern  &  "\s+Anzahl.*?(\d{1,4})(?=\s+(?:Davon|Cluster|Gruppe|L[üu]nern|A-|B-|Ganztag|$))"
    If r.Test(x) Then Set m=r.Execute(x)(0):QtyAfterLabel=m.SubMatches(0) Else QtyAfterLabel=""
End Function

Public Function ClusterDetails(body)
    Dim q1,q2,q3,q4
    q1=QtyAfterLabel(body,"Cluster\s*1")
    q2=QtyAfterLabel(body,"Cluster\s*2")
    q3=QtyAfterLabel(body,"Cluster\s*3")
    q4=QtyAfterLabel(body,"Cluster\s*4")
    If IsNumeric(q1) And IsNumeric(q2) And IsNumeric(q3) And IsNumeric(q4) Then
        ClusterDetails="Cluster1=" & q1 & "; Cluster2=" & q2 & "; Cluster3=" & q3 & "; Cluster4=" & q4
    Else
        ClusterDetails=""
    End If
End Function

Public Function LuenernDetails(body)
    Dim q1,q2
    q1=QtyAfterLabel(body,"L[üu]nern\s*1(?:\s*\(.*?\))?")
    q2=QtyAfterLabel(body,"L[üu]nern\s*2(?:\s*\(.*?\))?")
    If q1="" Then q1=QtyAfterLabel(body,"nern\s*1(?:\s*\(.*?\))?")
    If q2="" Then q2=QtyAfterLabel(body,"nern\s*2(?:\s*\(.*?\))?")
    If IsNumeric(q1) And IsNumeric(q2) Then LuenernDetails=LueText() & "=" & q1 & "; " & LueText() & "2=" & q2 Else LuenernDetails=""
End Function

Public Function SonderkostPreview(body)
    Dim lines,i,line,norm,ctx,re,m,label,qty,key,d,keys,k,out
    Set d=CreateObject("Scripting.Dictionary")
    lines=Split(Replace(Replace(body,vbCrLf,vbLf),vbCr,vbLf),vbLf)
    ctx=""

    Set re=CreateObject("VBScript.RegExp")
    re.IgnoreCase=True:re.Global=False
    re.Pattern="^(.*?)[,;]?\s+(\d{1,3})\s*$"

    For i=0 To UBound(lines)
        line=Trim(CStr(lines(i)))
        norm=Replace(line,vbTab," ")

        If RxTest(norm,"^Cluster\s*[1-4]\b") Then
            ctx=Rx1(norm,"^(Cluster\s*[1-4])")
        ElseIf RxTest(norm,"^Gruppe\s*[1-4]\b") Then
            ctx=Rx1(norm,"^(Gruppe\s*[1-4])")
        ElseIf RxTest(norm,"^L[üu]nern\s*2\b") Then
            ctx=LueText() & "2"
        ElseIf RxTest(norm,"^L[üu]nern(?:\s*1)?\b") Then
            ctx=LueText()
        ElseIf RxTest(norm,"^Liedbach") Then
            ctx="Liedbach"
        ElseIf RxTest(norm,"^Strolche") Then
            ctx="Strolche"
        End If

        ' FIX5ZF: Speiseplan-Zeilen strikt ausschließen.
        ' Nur typische Sonderkost-/Allergieformulierungen dürfen durch.
        If IsStrictSonderkostLine(norm) And re.Test(norm) Then
            Set m=re.Execute(norm)(0)
            label=Trim(CStr(m.SubMatches(0)))
            qty=CLng(m.SubMatches(1))
            If ctx="" Then ctx="?"
            key=ctx & "|" & label
            If d.Exists(key) Then
                d(key)=CLng(d(key))+qty
            Else
                d.Add key,qty
            End If
        End If
    Next

    out=""
    keys=d.Keys
    For Each k In keys
        If out<>"" Then out=out & " / "
        out=out & Replace(CStr(k),"|",": ") & " [" & CStr(d(k)) & "]"
    Next
    SonderkostPreview=out
End Function

Public Function IsStrictSonderkostLine(s)
    Dim x
    x=LCase(Trim(CStr(s)))

    ' Positive Sonderkost-/Allergie-Signale. Zutaten im normalen Speiseplan
    ' (z.B. Möhrensalat, Bio-Bulgur, Fischgericht) reichen ausdrücklich NICHT.
    IsStrictSonderkostLine = _
        (RxTest(x,"^laktosefrei[,! ]*(\d+)?$") Or _
         InStr(x,"allerg")>0 Or _
         InStr(x,"intoler")>0 Or _
         InStr(x,"unvertr")>0 Or _
         InStr(x,"kein schweine")>0 Or _
         InStr(x,"ohne schweine")>0 Or _
         InStr(x,"vegetar")>0 Or _
         InStr(x,"kein ei")>0 Or InStr(x,"ohne ei")>0 Or _
         InStr(x,"keine n")>0 Or InStr(x,"ohne n")>0 Or _
         InStr(x,"ohne rind")>0 Or _
         InStr(x,"ohne möhr")>0 Or InStr(x,"ohne karott")>0 Or _
         InStr(x,"ohne soja")>0 Or _
         InStr(x,"ohne hülsen")>0 Or _
         InStr(x,"ohne kiwi")>0)
End Function

Public Function SonderkostPretty(sk)
    Dim parts,p,out
    out=""
    parts=Split(CStr(sk)," / ")
    For Each p In parts
        If Trim(CStr(p))<>"" Then
            If out<>"" Then out=out & vbCrLf
            out=out & "- " & Trim(CStr(p))
        End If
    Next
    SonderkostPretty=out
End Function

Public Function SonderkostFormCount(sk)
    If Trim(CStr(sk))="" Then
        SonderkostFormCount=0
    Else
        SonderkostFormCount=UBound(Split(CStr(sk)," / "))+1
    End If
End Function

Public Function SonderkostMealCount(sk)
    Dim parts,p,re,m,total
    total=0
    parts=Split(CStr(sk)," / ")
    Set re=CreateObject("VBScript.RegExp")
    re.Global=False:re.IgnoreCase=True
    re.Pattern="\[(\d+)\]\s*$"
    For Each p In parts
        If re.Test(CStr(p)) Then
            Set m=re.Execute(CStr(p))(0)
            total=total+CLng(m.SubMatches(0))
        End If
    Next
    SonderkostMealCount=total
End Function

Public Function IsSonderkostLine(s)
    Dim x
    x=LCase(Trim(CStr(s)))
    IsSonderkostLine = _
        (InStr(x,"schweine")>0 Or InStr(x,"vegetar")>0 Or _
         InStr(x,"lakt")>0 Or InStr(x,"lact")>0 Or _
         InStr(x,"fruct")>0 Or InStr(x,"frukt")>0 Or _
         InStr(x,"gluten")>0 Or InStr(x,"nuss")>0 Or InStr(x,"nüsse")>0 Or _
         InStr(x,"soja")>0 Or InStr(x,"möh")>0 Or InStr(x,"karott")>0 Or _
         InStr(x,"hülsen")>0 Or InStr(x,"apfel")>0 Or InStr(x,"pfir")>0 Or _
         InStr(x,"zitr")>0 Or InStr(x,"kiwi")>0 Or InStr(x,"fisch")>0 Or _
         InStr(x,"paprika")>0 Or InStr(x,"erdbeer")>0 Or _
         InStr(x,"südfr")>0 Or InStr(x,"grünes gemüse")>0)
End Function

Public Function RxTest(s,pat)
    Dim r
    Set r=CreateObject("VBScript.RegExp")
    r.IgnoreCase=True:r.Global=False:r.Pattern=pat
    RxTest=r.Test(CStr(s))
End Function



' Jede geöffnete Wochenmappe liest nur Bestellungen für ihre eigenen Lieferdaten.
Public Function KC_TargetWeekMatches(wb As Object, ByVal subject As String, ByRef reason As String) As Boolean
    Dim firstDate As Variant, lastDate As Variant, ws As Object, info As Variant, mon As Variant
    On Error GoTo Failed
    firstDate = GermanDate(subject): lastDate = SecondGermanDate(subject)
    If Not IsDate(firstDate) Then reason = "Lieferdatum fehlt": Exit Function
    info = "": Set ws = FindWeekSheetByDate(wb, firstDate, info)
    If ws Is Nothing Then reason = CStr(info): Exit Function
    mon = WeekStartFromSheet(ws)
    If Not IsDate(mon) Then reason = "Wochenstart fehlt": Exit Function
    If NumberFromCell(ws.Name) <> DatePart("ww", CDate(firstDate), vbMonday, vbFirstFourDays) Then
        reason = "KW im Blattnamen widerspricht dem Lieferdatum": Exit Function
    End If
    If Weekday(CDate(firstDate), vbMonday) > 5 Then reason = "Lieferdatum am Wochenende": Exit Function
    If IsDate(lastDate) Then
        If Weekday(CDate(firstDate), vbMonday) <> 1 Or DateDiff("d", firstDate, lastDate) <> 4 Then
            reason = "Wochenbestellung muss Montag bis Freitag derselben Woche umfassen": Exit Function
        End If
        If DateValue(lastDate) > DateAdd("d", 4, CDate(mon)) Then reason = "Enddatum außerhalb der Zielwoche": Exit Function
    End If
    KC_TargetWeekMatches = True
    reason = ws.Name & " / " & DateText(mon)
    Exit Function
Failed:
    reason = "Zielwoche nicht sicher bestimmbar: " & Err.Description
End Function
