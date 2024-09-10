Attribute VB_Name = "TextForCtext"
Option Explicit
Rem TextForCtext相關操作
Private Property Get tx() As String
    tx = "TextForCtext"
End Property
Property Get TextForCtextExist()
    TextForCtextExist = word.Tasks.Exists(TextForCtext.tx)
End Property
Rem 檢查TextForCtext是否啟動 20240910
Private Function examToRun() As Boolean
    If Not word.Tasks.Exists(TextForCtext.tx) Then Exit Function
    'SystemSetup.wait 0.3
    Dim dt As Date
    dt = VBA.Now
    Do While DateDiff("s", dt, VBA.Now) < 0.3
        DoEvents
    Loop
    examToRun = True
End Function

Sub Hanchi_CTP_SearchingKeywordsYijing()
' Alt + shift + ,
' Alt + <
' Alt + shift + F5
' Ctrl + Alt + F9

    If Not examToRun Then Exit Sub
    
    AppActivate tx
    DoEvents
    SendKeys "%{F9}"
    DoEvents
End Sub
Rem 送交《古籍酷》自動標點。先複製好是要處理的純文字。將結果讀回至剪貼簿中
Function GjcoolPunct() As Boolean
    
    If Not examToRun Then Exit Function
    
    AppActivate tx
    DoEvents
    SendKeys "^a"
    DoEvents
    AppActivate tx
    SystemSetup.wait 0.1
    DoEvents
    SendKeys "+{INSERT}"
    DoEvents
    AppActivate tx
    DoEvents
    SendKeys "^%{F10}"
    DoEvents
    Dim dt As Date, x As String
    Dim containsPunctuation As Boolean, punctuation As String, i As Byte, noErrOccured As Boolean
    dt = VBA.Now
    Do While DateDiff("s", dt, VBA.Now) < 30
        If DateDiff("s", dt, VBA.Now) Mod 1 = 0 Then
            GoSub puncted
            If containsPunctuation Then
                noErrOccured = True
                Exit Do
            End If
        End If
        DoEvents
        
    Loop
    If Not noErrOccured Then
        GjcoolPunct = False
        Exit Function
    End If
    DoEvents
    AppActivate tx
    DoEvents
    SystemSetup.wait 0.1
    SendKeys "^a"
    DoEvents
    'SendKeys "^x" '複製文本的功能寫在C#中
    SystemSetup.wait 0.05
    SendKeys "{delete}"
    DoEvents
    GjcoolPunct = True
    Exit Function
puncted:
    
    punctuation = "。，"
    For i = 1 To Len(punctuation)
        If InStr(ClipBoardObject.GetClipboard, Mid(punctuation, i, 1)) > 0 Then
            containsPunctuation = True
            Exit For
        End If
    Next i
Return
End Function

