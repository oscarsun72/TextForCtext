VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventClassModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'https://learn.microsoft.com/zh-tw/office/vba/word/concepts/objects-properties-methods/using-events-with-the-application-object-word
'https://learn.microsoft.com/zh-tw/office/vba/word/concepts/objects-properties-methods/using-events-with-the-application-object-word
Public WithEvents App As word.Application
Attribute App.VB_VarHelpID = -1

'Public WithEvents EventClassModule As Word.Application
'Private Sub EventClassModule_WindowActivate _
'        (ByVal Wn As Word.Window)
'    Wn.WindowState = wdWindowStateMaximize
'    隱藏字表
'End Sub

'20230120 和 chatGPT大菩薩討論的結果，在Word物件類別模組中是不能加入 App_Startup() 此程序的，只能用 AutoExec 模組的方法來達成啟動 Word時執行特定程序
'Private Sub App_Startup()
'    '在這裡添加您想要在Word啟動時執行的程序
'    MsgBox "Word啟動了"
'End Sub
'
'Private Sub App_Startup()
''20230120 12:10 初入小年夜 creedit chatGPT大菩薩：VBA  如何結束 MS word 時處理程序呢：
'    '在這裡添加您想要執行的程序
'    SystemSetup.UserProfilePath = SystemSetup.取得使用者路徑_含反斜線()
'End Sub

Private Sub app_Quit()
    '自動備份Normal.dotm
    SystemSetup.backupNormal_dotm
    
    
    '清除Word Selenium所開啟的chromedriver
    '20230119 chatGPT菩薩：VBA  如何結束 MS word 時處理程序呢：
    'https://learn.microsoft.com/zh-tw/office/vba/word/concepts/objects-properties-methods/using-events-with-the-application-object-word
    SystemSetup.killchromedriverFromHere
    Dim dt As Date
    dt = VBA.Now
    Do While DateDiff("s", dt, VBA.Now) < 1
        DoEvents
    Loop
    If Chromedrivers.Count > 0 Then
        If MsgBox("是否要清除剩下的chromedriver？", vbOKCancel + vbExclamation) = vbOK Then
            killchromedrivers
        End If
    End If
End Sub


Private Sub App_DocumentBeforeSave(ByVal doc As Document, SaveAsUI As Boolean, Cancel As Boolean)
    '隱藏字表
End Sub

Private Sub App_DocumentChange() '寫在此處較有效!2010/10/20
'Register_Event_Handler
'原來沒效是因為我在"在另一文件中尋找選取字串"使用End使註冊程序Register_Event_Handler失效了!
隱藏字表
'If InStr(ActiveDocument.Name, "字表7") = 0 Then
'    if Documents("字表7.2.doc"). .Visible = False
'End If
End Sub



Private Sub App_WindowActivate(ByVal doc As Document, ByVal Wn As Window)
''If InStr(Doc.Name, "字表7") = 0 Then Wn.Visible = False
''    'Wn.WindowState = wdWindowStateMaximize
''    隱藏字表
'If VBA.Left(Wn.Caption, 2) = "排印" And VBA.IsNumeric(VBA.Mid(Wn.Caption, 3, 4)) Then
'    If Tasks.Exists(VBA.Mid(Wn.Caption, 3) & ".pdf") Then
'       Tasks(VBA.Mid(Wn.Caption, 3) & ".pdf").Activate
''    If Tasks("Adobe Reader").Exists Then
''       Tasks(VBA.Mid(Wn.Caption, 3, 4) & "*pdf*").Activate
'       Tasks(Wn.Caption & " - Microsoft Word").Activate
'    End If
'End If
End Sub

Private Sub App_WindowDeactivate(ByVal doc As Document, ByVal Wn As Window) '寫在App_DocumentChange較有效!2010/10/20
'Register_Event_Handler
'If Documents.Count > 1 Then
'    If InStr(Doc.Name, "字表7") Then Wn.Visible = False
'End If
隱藏字表
End Sub


Private Sub App_WindowSelectionChange(ByVal Sel As Selection)
    If Not Docs.AutoCopy Then
        Exit Sub
    End If
    If Selection.Type = wdSelectionIP Then
        Exit Sub
    End If
'    Stop
    On Error GoTo eH
    'Dim clipboardTxt As String
'    clipboardTxt = SystemSetup.GetClipboardText
    
    If Selection.InlineShapes.Count = 0 And Selection.ShapeRange.Count = 0 Then
        If code.IsChineseString(Selection.text) Then
            If Selection.Type = wdSelectionNormal Then
                Selection.Copy
                SystemSetup.playSound 0.96
            End If
        End If
    End If

    Exit Sub
eH:
Select Case Err.number
    Case 4605, -2147467259 'Copy' 方法 ('Range' 物件) 失敗
        '4605此方法或屬性無法使用，因為未選取任何文字。
        Exit Sub
    Case Else
        Debug.Print Err.number & Err.description
        MsgBox Err.number & Err.description
        'Resume
End Select
End Sub
