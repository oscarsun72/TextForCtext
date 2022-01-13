Attribute VB_Name = "易學網"
Option Explicit
Function isOmitChar(char) As Boolean
Dim omitChar As String
omitChar = ChrW(160) & ChrW(13)
If InStr(omitChar, char) > 0 Then isOmitChar = True
End Function


'https://www.eee-learning.com/book/
Sub 周易本義()
Dim d As Document, a, i As Long, noteFlag As Boolean, x As String, _
    omitPara As Paragraph, acnt As Long, omitAcnt As Long, omitParaNext As Paragraph _
    , omitParaNextRng As Range, omitParaRng As Range, omitParaNext1Rng As Range, omitParaNext1 As Paragraph _
    , noteFlgPrevious As Boolean, openCnt As Long, closeCnt As Long
Set d = Documents.Add()
d.Range.Paste
acnt = d.Characters.Count
'd.Range.Find.Execute "^p", , , , , , , wdFindContinue, , "", wdReplaceAll
For i = 1 To acnt
    Set a = d.Characters(i)
    Do While InStr(ChrW(13) & ChrW(160) & ChrW(32) & ChrW(30), a) > 0
        If a.Next Is Nothing Then Exit Do
        Set a = a.Next
        i = i + 1
    Loop
    If InStr(a.Paragraphs(1).Range, "《彖》") > 0 Or _
        InStr(a.Paragraphs(1).Range, "《象》") > 0 Then
        Set omitPara = a.Paragraphs(1)
        Set omitParaRng = omitPara.Range
        Set omitParaNext = omitPara.Next
        Set omitParaNextRng = omitParaNext.Range
        Set omitParaNext1 = omitPara.Next.Next
        If Not omitParaNext1 Is Nothing Then
            Set omitParaNext1Rng = omitParaNext1.Range
            Set a = omitParaNext1Rng.Characters(1)
            i = i + omitParaRng.Characters.Count + omitParaNextRng.Characters.Count
        End If
    End If

    If Not a.Previous Is Nothing And Not a.Next Is Nothing Then
        If a.Font.Bold = True Then
            noteFlag = False
        Else '注文
            noteFlag = True
        End If
        If Not noteFlag And noteFlgPrevious Then
            x = x + "}}" + a
            noteFlgPrevious = noteFlag
            closeCnt = closeCnt + 1
        Else
            x = x & a
            If Not noteFlag And a.Next.Font.Bold = False Then
                x = x + "{{"
                noteFlgPrevious = True
                openCnt = openCnt + 1
            End If
            If noteFlag And a.Next.Font.Bold = True Then
                x = x + "}}"
                closeCnt = closeCnt + 1
            End If
        End If
    End If
Next i
If openCnt > closeCnt Then x = x + "}}"
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
 

Sub 易程傳_伊川易傳()
Dim rng As Range, a, x As String, noteFlg As Boolean, preNoteFlg As Boolean, rngEnd As Range
Set rng = Documents.Add.Range
rng.Paste
Set rngEnd = rng
If rngEnd.Find.Execute("本文取自易學網。") Then rng.SetRange 0, rngEnd.start
For Each a In rng.Characters
    If isOmitChar(a) Then
'        If preNoteFlg Then
'            x = x & "}}"
'        End If
    Else
        If a.Font.Bold Then
            noteFlg = False
        Else
            noteFlg = True
        End If
        If preNoteFlg And Not noteFlg Then
            x = x & "}}" & a
        ElseIf Not preNoteFlg And noteFlg Then
            x = x & "{{" & a
        Else
            x = x & a
        End If
    End If
    preNoteFlg = noteFlg
Next a
rng.SetRange rng.Document.Range.start, rng.Document.Range.End
rng.Text = Replace(x, " ", "")
rng.Cut
rng.Document.Close wdDoNotSaveChanges
Beep
End Sub
