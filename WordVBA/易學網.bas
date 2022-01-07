Attribute VB_Name = "易學網"
Option Explicit
'https://www.eee-learning.com/book/
Sub 周易本義()
Dim d As Document, a, noteFlag As Boolean, x As String
Set d = ActiveDocument
'If d.path <> "" Then
Set d = Documents.Add()
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
SystemSetup.ClipboardPutIn Replace(Replace(x, " ", ""), ChrW(160), "")
d.Close wdDoNotSaveChanges
Beep
End Sub

Sub 周易本義_象傳()
Dim x As String, s As Integer, rng As Range, p As Paragraph, pn As Paragraph, px As String
Set rng = Documents.Add().Range
rng.Paste
For Each p In rng.Paragraphs
    If Left(p.Range, 4) = "《象》曰" Then
        px = p.Range.Text
        s = InStr(px, "本文取自易學網。") - 1
        If s = -1 Then s = Len(px) - 1
        px = Mid(px, 1, s)
        x = x & Replace(px, Chr(13), "")
        Set pn = p.Next
        If pn.Range.Characters.Count > 2 Then
            px = pn.Range.Text
            s = InStr(px, "本文取自易學網。") - 1
            If s = -1 Then s = Len(px) - 1
            px = Mid(px, 1, s)
            x = x & "{{" & Replace(px, Chr(13), "") & "}}"
            Set p = p.Next
        End If
    End If
Next p
SystemSetup.ClipboardPutIn Replace(Replace(Replace(x, " ", ""), "《象》曰：", ""), ChrW(160), "")
rng.Document.Close wdDoNotSaveChanges
Beep
End Sub
 
