Attribute VB_Name = "���Ǻ�"
Option Explicit
'https://www.eee-learning.com/book/
Sub �P�����q()
Dim d As Document, a, noteFlag As Boolean, x As String, omitPara As Paragraph
Set d = ActiveDocument
'If d.path <> "" Then
Set d = Documents.Add()
d.Range.Paste
'd.Range.Find.Execute "^p", , , , , , , wdFindContinue, , "", wdReplaceAll
For Each a In d.Characters
    If InStr(a.Paragraphs(1).Range, "�mν�n") > 0 Or _
        InStr(a.Paragraphs(1).Range, "�m�H�n") > 0 Then
        Set omitPara = a.Paragraphs(1)
        If Not omitPara.Next.Next Is Nothing Then _
            Set a = omitPara.Next.Next.Range.Characters(1)
    End If
    Do While InStr(Chr(13) & Chr(160) & Chr(32) & Chr(30), a) > 0
        If a.Next Is Nothing Then Exit Do
        Set a = a.Next
    Loop
    x = x & a
    If Not a.Previous Is Nothing And Not a.Next Is Nothing Then
        If a.Font.Bold = True Then
            noteFlag = False
        Else '�`��
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

Sub �P�����q_�H��()
Dim x As String, s As Integer, rng As Range, p As Paragraph, pn As Paragraph, px As String
Set rng = Documents.Add().Range
rng.Paste
For Each p In rng.Paragraphs
    If Left(p.Range, 4) = "�m�H�n��" Then
        px = p.Range.Text
        s = InStr(px, "������۩��Ǻ��C") - 1
        If s = -1 Then s = Len(px) - 1
        px = Mid(px, 1, s)
        x = x & Replace(px, Chr(13), "")
        Set pn = p.Next
        If pn.Range.Characters.Count > 2 Then
            px = pn.Range.Text
            s = InStr(px, "������۩��Ǻ��C") - 1
            If s = -1 Then s = Len(px) - 1
            px = Mid(px, 1, s)
            x = x & "{{" & Replace(px, Chr(13), "") & "}}"
            Set p = p.Next
        End If
    End If
Next p
SystemSetup.ClipboardPutIn Replace(Replace(Replace(x, " ", ""), "�m�H�n��G", ""), ChrW(160), "")
rng.Document.Close wdDoNotSaveChanges
Beep
End Sub
 
