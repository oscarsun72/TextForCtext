Attribute VB_Name = "Transferkit"
Option Explicit

Sub hperLinkTextToDisplayShorten()
Dim rng As Range, h As Hyperlink, ur As UndoRecord, d As Document
Set d = Documents.Add()
Set rng = d.Range
'Set ur = SystemSetup.stopUndo("hperLinkTextToDisplayShorten")
SystemSetup.stopUndo ur, "hperLinkTextToDisplayShorten"
rng.Paste
'd.ActiveWindow.Visible = True
For Each h In rng.Hyperlinks
    'h.TextToDisplay = Mid(h.TextToDisplay, InStrRev(h.TextToDisplay, "/") + 1)
    h.TextToDisplay = code.URLDecode(Mid(h.Address, InStrRev(h.Address, "/") + 1))
Next h
rng.Find.Execute "^p", , , , , , True, wdFindContinue, , "^t", wdReplaceAll

If d.Characters.Count > 10000 Then
    userProfilePath = SystemSetup.取得使用者路徑_含反斜線()
    If Dir(userProfilePath & "Dropbox\") <> "" Then
    d.SaveAs2 userProfilePath & "Dropbox\hperLinkTextToDisplayShorten.docx", , , , False
    End If
    d.Activate
    d.ActiveWindow.Visible = True
Else
    rng.Cut
    d.Application.WindowState = wdWindowStateMinimize
    d.Close wdDoNotSaveChanges
End If
SystemSetup.contiUndo ur
Beep
End Sub

