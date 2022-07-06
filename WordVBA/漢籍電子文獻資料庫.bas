Attribute VB_Name = "漢籍電子文獻資料庫"
Option Explicit
Sub clearResultMarked()
Dim a As Range, d As Document, rng As Range
Set d = Documents.Add
Set rng = d.Range
With rng
    .Paste
    .Find.Font.Shading.BackgroundPatternColor = 65535
    Do While .Find.Execute()
        If Asc(.Text) > 13 Or Asc(.Text) < 0 Then
            .Shading.BackgroundPatternColor = wdColorAutomatic
            With .Font
                .Color = 0
                .Bold = 0
            End With
        End If
        rng.SetRange rng.End, d.Range.End - 1
    Loop
    .Find.ClearFormatting
End With

'For Each a In d.Characters
'    If a.Shading.BackgroundPatternColor = 65535 Then '<> wdColorAutomatic
'        With a
'            .Shading.BackgroundPatternColor = wdColorAutomatic
'            With .Font
'                .Color = 0
'                .Bold = 0
'            End With
'        End With
'    End If
'Next a
d.Range.Cut
DoEvents
End Sub

Sub 漢籍電子文獻資料庫文本整理_以轉貼到中國哲學書電子化計劃()
clearResultMarked
文字處理.漢籍電子文獻資料庫文本整理_以轉貼到中國哲學書電子化計劃
SystemSetup.playSound 2
End Sub
Sub 漢籍電子文獻資料庫文本整理_十三經注疏()
clearResultMarked
文字處理.漢籍電子文獻資料庫文本整理_以轉貼到中國哲學書電子化計劃 True
漢籍電子文獻資料庫文本整理_十三經注疏_sub
On Error Resume Next
AppActivate "TextForCtext" '"EmEditor"
End Sub
Sub 漢籍電子文獻資料庫文本整理_十三經注疏_sub()
Dim d As Document, a, i As Integer, ub As Integer
a = Array("^p" & ChrW(12310) & "疏" & ChrW(12311), ChrW(12310) & "疏" & ChrW(12311) & "{{", _
    "．", "", "釋曰", "《釋》曰：", "正義曰", "《正義》曰：", "○", ChrW(12295), "^p彖曰", "<p>〈彖〉曰：", "^p象曰", "<p>〈象〉曰：", _
    "^p", "}}<p>^p", "^p" & ChrW(12295), "}}<p>" & ChrW(12295), ChrW(12295) & "^p", ChrW(12295) & "}}<p>", _
    "}}", "。}}", "。}}<p>^p。}}<p>", "。}}<p>", "。}}<p>。}}<p>", "。}}<p>", "{{注。}}", "○《注》：")
ub = UBound(a) - 1
Set d = ActiveDocument
If d.path <> "" Then
    Set d = Documents.Add
    d.Range.Paste
End If
For i = 0 To ub
    d.Range.Find.Execute a(i), , , , , , True, wdFindContinue, , a(i + 1), wdReplaceAll
    i = i + 1
Next i

文字處理.書名號篇名號標注
d.Range.Cut
d.Close wdDoNotSaveChanges
End Sub

Sub get阮元挍勘記()
'其實本有，在原書各頁下，不必做。
Dim rng As Range, a, noteFlag As Boolean, x As String
Set rng = Documents.Add().Range
rng.Paste
For Each a In rng.Characters
    If a.Font.Color = 255 Then
        If a.Font.Size = 10 Then
            noteFlag = False
        Else
            noteFlag = True
        End If
        If noteFlag = False Then
            If a.Next.Font.Size = 10 Then
                x = x & a
            Else
                x = x & a & "：{{"
            End If
        Else
            If a.Next.Font.Size > 7.5 Or a.Next.Font.Color <> 255 Then
                x = x & a & "}}<p>" & Chr(13) & Chr(10)
            Else
                x = x & a
            End If
        End If
    End If
Next a
rng.Text = x
文字處理.書名號篇名號標注
rng.Cut
rng.Document.Close wdDoNotSaveChanges
'SystemSetup.ClipboardPutIn x
Beep
End Sub
