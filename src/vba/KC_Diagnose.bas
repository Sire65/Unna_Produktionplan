Attribute VB_Name = "KC_Diagnose"
Option Explicit
Public Function KC_DiagnoseMail(ByVal subject As String, ByVal body As String) As String
    Dim e As Object, p As Collection, item As Variant
    Set e = KC_Evaluate(subject, body): Set p = e("plan")
    KC_DiagnoseMail = e("status") & vbTab & e("bookable") & vbTab & KC_V20.SonderkostMealCount(KC_V20.SonderkostPreview(body)) & vbTab & e("forms") & vbTab & e("note")
    For Each item In p
        KC_DiagnoseMail = KC_DiagnoseMail & vbLf & item(0) & vbTab & item(1) & vbTab & item(2) & vbTab & item(3)
    Next item
End Function
