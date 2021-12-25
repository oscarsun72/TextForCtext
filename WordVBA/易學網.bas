Attribute VB_Name = "易學網"
Option Explicit
'https://www.eee-learning.com/book/
Sub 周易本義()
Dim d As Document, a, noteFlag As Boolean, x As String
Set d = ActiveDocument
If d.path <> "" Then Set d = Documents.Add()
d.Range.Paste
d.Range.Find.Execute "^p", , , , , , , wdFindContinue, , "", wdReplaceAll
For Each a In d.Characters
    x = x & a
    If Not a.Previous Is Nothing And Not a.Next Is Nothing Then
        If a.Font.Bold = True Then
            noteFlag = False
        Else '注文
            noteFlag = True
        End If
        If Not noteFlag And a.Next.Font.Bold = False Then
            x = x + "{{"
        End If
        If noteFlag And a.Next.Font.Bold = True Then
            x = x + "}}"
        End If
    End If
Next a
SystemSetup.ClipboardPutIn x
Beep
End Sub
