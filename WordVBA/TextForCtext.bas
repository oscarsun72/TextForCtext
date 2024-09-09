Attribute VB_Name = "TextForCtext"
Option Explicit
Rem TextForCtext¬ÛÃö¾Þ§@
Property Get tx() As String
    tx = "TextForCtext"
End Property

Sub Hanchi_CTP_SearchingKeywordsYijing()
' Alt + shift + ,
' Alt + <
' Ctrl + Alt + F9
' Ctrl + Alt + F5
    If Not word.Tasks.Exists(tx) Then Exit Sub
    Dim dt As Date
    dt = VBA.Now
    Do While DateDiff("s", dt, VBA.Now) < 0.2
        DoEvents
    Loop

    AppActivate tx
    DoEvents
    SendKeys "%{F9}"
    DoEvents
End Sub
