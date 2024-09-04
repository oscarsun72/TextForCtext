Attribute VB_Name = "AutoExec"
Option Explicit


'Sub AutoExec()
''    Stop
'    UserProfilePath = SystemSetup.取得使用者路徑_含反斜線()
'    '在這裡添加您想要執行的程序
''    SystemSetup.ShortcutKeys
'    Stop
'    Register_Event_Handler
'
'End Sub

Rem 20240903 Copilot大菩薩：Word VBA 事件處理程序註冊：https://sl.bing.net/jKWC0wBWaFo
'要讓 AutoExec 巨集在 Normal.dotm 中運行並參照 Startup 路徑下的 TextForCtextWordVBA.dotm 範本裡的專案，您可以使用 Application.Run 方法來調用 TextForCtextWordVBA.dotm 中的程序。以下是修改後的範例：
'在 Normal.dotm 中的 AutoExec 巨集：
Sub AutoExec()
'    Stop
    ' 確保 TextForCtextWordVBA.dotm 已加載
    AddInLoad "TextForCtextWordVBA.dotm"
    
    ' 調用 TextForCtextWordVBA.dotm 中的程序（須是 sub 才能執行）
    'Application.Run "TextForCtextWordVBA.SystemSetup.取得使用者路徑_含反斜線"
    
    ' 在這裡添加您想要執行的程序
    'Application.Run "TextForCtextWordVBA.SystemSetup.ShortcutKeys"
    
    ' 註冊事件處理程序
    'Register_Event_Handler
    Application.Run "TextForCtextWordVBA.Docs.Register_Event_Handler"
End Sub

Sub AddInLoad(addInName As String)
    Dim addIn As Template
    On Error Resume Next
    Set addIn = Application.Templates(addInName)
    If addIn Is Nothing Then
        Set addIn = Application.AddIns.Add(FileName:=Application.StartupPath & "\" & addInName, Install:=True)
    End If
    On Error GoTo 0
End Sub
