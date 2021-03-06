Attribute VB_Name = "中國哲學書電子化計劃"
Option Explicit
Sub 新頁面()
'the page begin
Dim start As Integer
' the page end
Dim e As Integer
' the book
Dim fileID As Long
'https://ctext.org/library.pl?if=gb&file=1000081&page=2621

Dim x As String ', data As New MSForms.DataObject
Dim i As Integer, rng As Range
Set rng = ActiveDocument.Range
start = CInt(Replace(rng.Paragraphs(1).Range, Chr(13), ""))
e = CInt(Replace(rng.Paragraphs(2).Range, Chr(13), ""))
fileID = CLng(Replace(rng.Paragraphs(3).Range, Chr(13), ""))
For i = start To e
    If i = 1 Then
        x = x & "<scanbegin file=""" & fileID & """ page=""" & i & """ />●" & Chr(9) & "<scanend file=""" & fileID & """ page=""" & i & """ />"
    Else
        x = x & "<scanbegin file=""" & fileID & """ page=""" & i & """ />" & Chr(9) & "<scanend file=""" & fileID & """ page=""" & i & """ />" '若中間沒有任何內容，頁面最後便不能成一段落。若剛好一個段落，會與下一頁黏合在一起
    End If
Next i

rng.Paragraphs(3).Range = CLng(Replace(rng.Paragraphs(3).Range, Chr(13), "")) + 1
'For Each e In Selection.Value
'    x = x & e
'Next e
''x = Replace(x, Chr(13), "")
'data.SetText Replace(x, "/>", "/>●", 1, 1)
'data.PutInClipboard
SystemSetup.SetClipboard x
DoEvents
Network.AppActivateDefaultBrowser
SendKeys "^v"

End Sub
Sub 清除頁前的分段符號()
Dim d As Document, rng As Range, e As Long, s As Long
Set d = Documents.Add
DoEvents
將星號前的分段符號移置前段之末 d
DoEvents
Set rng = d.Range
'd.ActiveWindow.Visible = True
'rng.Paste
rng.Find.ClearFormatting
Do While rng.Find.Execute("<scanbegin ") '<scanbegin file="80564" page="13" y="4" />
    rng.MoveEndUntil ">"
    rng.MoveEnd
'    rng.Select
    rng.SetRange rng.End, rng.End + 2
    If rng.Text = Chr(13) & Chr(13) Then
'        rng.Select
        e = rng.End
        rng.Delete
        Set rng = d.Range(e, d.Range.End)
    Else
    rng.SetRange rng.End, d.Range.End
    End If
Loop

playSound 1

Set rng = d.Range
rng.Find.ClearFormatting
Do While rng.Find.Execute("<scanend file=") ', , , , , , True, wdFindStop)
    s = rng.start
    rng.MoveEndUntil ">"
    rng.MoveEnd
'    rng.Select
    rng.SetRange rng.End, rng.End + 2
    If rng.Text = Chr(13) & Chr(13) Then
'        e = rng.End
'        rng.Select
        rng.Cut
        rng.SetRange s, s
        rng.Paste
        Set rng = d.Range(e, d.Range.End)
    Else
        rng.SetRange rng.End, d.Range.End
    End If
Loop


DoEvents
d.Range.Cut
DoEvents
playSound 1
DoEvents
pastetoEditBox "將星號前的分段符號移置前段之末 & 清除頁前的分段符號"
d.Close wdDoNotSaveChanges

End Sub

Sub 將星號前的分段符號移置前段之末(ByRef d As Document) '20220522
Dim rng As Range, e As Long, s As Long, rngP As Range
'd As Document,Set d = Documents.Add
Set rng = d.Range
rng.Paste
rng.Find.ClearFormatting
Do While rng.Find.Execute("*")
    e = rng.End
    If rng.start > 0 Then
        If rng.Previous = Chr(13) Then
            Set rng = rng.Previous
            If rng.Previous = Chr(13) Then
                Set rng = rng.Previous
                If rng.Previous = ">" Then
                    rng.SetRange rng.start, e - 1
                    s = rng.start
                    Set rngP = d.Range(s, s)
                    rng.Delete
                    Do Until rngP.Next = "<"
                        If rngP.start = 0 Then GoTo NextOne
                        rngP.move wdCharacter, -1
                    Loop
                    rngP.move
                    rngP.InsertAfter Chr(13) & Chr(13)
                End If
            End If
        End If
    End If
NextOne:
    Set rng = d.Range(e, d.Range.End)
Loop
'd.Range.Cut
'd.Close wdDoNotSaveChanges
playSound 1
'pastetoEditBox "將星號前的分段符號移置前段之末"

End Sub

Sub 將每頁間的分段符號清除()
Dim d As Document, rng As Range, s As Long, e As Long, rngCheck As Range
Const pageStart As String = "<scanbegin file="
Const pageEnd As String = "<scanend file="
Set d = ActiveDocument
Set rng = d.Range(Len(pageStart), d.Range.End)
Do While rng.Find.Execute(pageStart)
    e = rng.start: s = e - 2
    Set rngCheck = d.Range(s, e)
    rngCheck.Select
    If rngCheck.Previous = ">" Then rngCheck.Delete
    rng.SetRange rng.End + 1, d.Range.End
Loop
End Sub

Sub pastetoEditBox(Description_from_ClipBoard As String)
word.Application.WindowState = wdWindowStateMinimize
AppActivateDefaultBrowser
DoEvents
SendKeys "^v", True
DoEvents ' DoEvents: DoEvents
Beep
SystemSetup.Wait 0.5
DoEvents:
SendKeys "{tab}"
'SystemSetup.ClipboardPutIn Description_from_ClipBoard
DoEvents
'SendKeys "^v"
SendKeys Description_from_ClipBoard
SendKeys "{tab 2}~"
End Sub

Sub 轉成黑豆以作行字數長度判斷用()
Dim p As Paragraph, a, i As Byte, cntr As Byte, ur As UndoRecord
Set ur = SystemSetup.stopUndo("轉成黑豆以作行字數長度判斷用")
Set p = Selection.Paragraphs(1)
cntr = p.Range.Characters.Count - 1
For i = 1 To cntr
    Set a = p.Range.Characters(i)
    If a.Text <> Chr(13) Then a.Text = "●"
Next i
p.Range.Cut
SystemSetup.contiUndo ur
Set ur = Nothing
End Sub
Sub 清除所有符號_分段注文符號例外()
Dim f, i As Integer
f = Array("。", "」", Chr(-24152), "：", "，", "；", _
    "、", "「", ".", Chr(34), ":", ",", ";", _
    "……", "...", "．", "【", "】", " ", "《", "》", "〈", "〉", "？" _
    , "！", "﹝", "﹞", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0" _
    , "『", "』", ChrW(9312), ChrW(9313), ChrW(9314), ChrW(9315), ChrW(9316) _
    , ChrW(9317), ChrW(9318), ChrW(9319), ChrW(9320), ChrW(9321), ChrW(9322), ChrW(9323) _
    , ChrW(9324), ChrW(9325), ChrW(9326), ChrW(9327), ChrW(9328), ChrW(9329), ChrW(9330) _
    , ChrW(9331), ChrW(8221), """") '先設定標點符號陣列以備用
    '全形圓括弧暫不取代！
    For i = 0 To UBound(f)
        ActiveDocument.Range.Find.Execute f(i), True, , , , , , wdFindContinue, True, "", wdReplaceAll
    Next
End Sub

Sub 撤掉與書圖的對應_脫鉤() '20220210
Dim rng As Range, angleRng As Range
word.Application.ScreenUpdating = False
Set rng = Documents.Add().Range
Set angleRng = rng
rng.Paste
Do While rng.Find.Execute("<")
    rng.MoveEndUntil ">"
    rng.SetRange rng.start, rng.End + 1
    angleRng.SetRange rng.start, rng.End
    If InStr(angleRng.Text, "file") > 0 Then angleRng.Delete
Loop
rng.Document.Range.Cut
rng.Document.Close wdDoNotSaveChanges
pastetoEditBox "與原本書圖不合，圖文脫鉤。另依《維基文庫》本輔以末學自製軟件TextForCtext對應錄入。感恩感恩　南無阿彌陀佛"
word.Application.ScreenUpdating = True
End Sub

Sub formatter() '為《經典釋文》春秋三傳等格式用，日後可改成其他需要格式化的文本
Dim d As Document, rng As Range, a As Range, s As Long, e As Long
Const spcs As String = "　"

Set d = Documents.Add: Set rng = d.Range
rng.Paste
For Each a In d.Characters
    If a = spcs Then
        If a.Next = spcs Then
            If InStr(a.Paragraphs(1).Range.Text, "*") = 0 Then
                a.Select
                s = a.start
                Do Until Selection.Next <> spcs
                    Selection.MoveRight , , wdExtend
                Loop
                e = Selection.End
                Set a = Selection.Next
                If a.Next = spcs Then
                    a.Select
                    Do Until Selection.Next <> spcs
                        Selection.Next.Delete
                    Loop
                
                    rng.SetRange s, e
                    'rng.Select
                    rng.Text = Replace(rng.Text, "　", ChrW(-9217) & ChrW(-8195))
                    Set a = rng.Characters(rng.Characters.Count)
                End If
            End If
        End If
    End If
Next a
d.Range.Cut
d.Close wdDoNotSaveChanges
SystemSetup.playSound 2
End Sub

Sub formatter年前加分段符號() '為《經典釋文》春秋三傳等格式用，日後可改成其他需要格式化的文本
Dim d As Document, rng As Range, a As Range, s As Long, e As Long, i As Integer, yi As Byte, ok As Boolean, yStr As String
Const y As String = "年"
Set d = Documents.Add: Set rng = d.Range
'd.ActiveWindow.Visible = True
rng.Paste
rng.Find.ClearFormatting
Do While rng.Find.Execute("^p")
    If rng.End = d.Range.End - 1 Then Exit Do
    Set a = d.Range
    For i = 4 To 2 Step -1
        a.SetRange rng.End, rng.End + i
'        a.Select
        If Right(a, 1) = y Then
            If a.Previous.Previous <> ">" Then
                For yi = 1 To 99
                    yStr = 文字轉換.數字轉漢字2位數(yi) + y
                    If a.Text = yStr Then
                        rng.InsertBefore "<p>"
                        ok = True: Exit For
                    End If
                Next yi
                If ok Then
                    ok = False
                    Exit For
                End If
            End If
        End If
    Next i
    rng.SetRange rng.End, d.Range.End
Loop
SystemSetup.playSound 2
d.Range.Cut
d.Close wdDoNotSaveChanges
End Sub
Sub 維基文庫四部叢刊本轉來()
Dim d As Document, a, i, p As Paragraph, xP As String, acP As Integer, space As String, rng As Range
On Error GoTo eh
a = Array(ChrW(12296), "{{", ChrW(12297), "}}", "〈", "{{", "〉", "}}", _
    "○", ChrW(12295))
'《容齋三筆》等小注作正文省版面者 https://ctext.org/library.pl?if=gb&file=89545&page=24
'a = Array("〈", "", "〉", "", _
    "○", ChrW(12295))


Set d = Documents.Add()
d.Range.Paste
維基文庫造字圖取代為文字 d.Range
For i = 0 To UBound(a) - 1
    d.Range.Find.Execute a(i), , , , , , True, wdFindContinue, , a(i + 1), wdReplaceAll
    i = i + 1
Next i
For Each p In d.Range.Paragraphs
    xP = p.Range
    If Left(xP, 2) = "{{" And Right(xP, 3) = "}}" & Chr(13) Then
        xP = Mid(p.Range, 3, Len(xP) - 5)
        If InStr(xP, "{{") = 0 And InStr(xP, "}}") = 0 Then
            acP = p.Range.Characters.Count - 1
            If acP Mod 2 = 0 Then
                acP = CInt(acP / 2)
            Else
                acP = CInt((acP + 1) / 2)
            End If
            If p.Range.Characters(acP).InlineShapes.Count = 0 Then
                p.Range.Characters(acP).InsertParagraphAfter
            Else
                p.Range.Characters(acP).Select
                Selection.Delete
                Selection.TypeText " "
                p.Range.Characters(acP).InsertParagraphAfter
            End If
        End If
    ElseIf Left(xP, 1) = "　" Then '前有空格的
        i = InStr(xP, "{{")
        If i > 0 And Right(xP, 3) = "}}" & Chr(13) Then
            space = Mid(xP, 1, i - 1)
            If Replace(space, "　", "") = "" Then
                xP = Mid(xP, i + 2, Len(xP) - 3 - (i + 2))
                If InStr(xP, "{{") = 0 And InStr(xP, "}}") = 0 Then
                    Set rng = p.Range
                    rng.SetRange rng.Characters(1).start, rng.Characters(i + 1).End
                    rng.Text = "{{" & space
                    acP = p.Range.Characters.Count - 1 - Len(space)
                    If acP Mod 2 = 0 Then
                        acP = CInt(acP / 2) + Len(space) + 1
                    Else
                        acP = CInt((acP + 1) / 2) + Len(space) + 1
                    End If
                    If p.Range.Characters(acP).InlineShapes.Count = 0 Then
                        p.Range.Characters(acP).InsertBefore Chr(13) & space
                    Else
                        p.Range.Characters(acP).Select
                        Selection.Delete
                        Selection.TypeText " "
                        p.Range.Characters(acP).InsertBefore Chr(13) & space
                    End If
                    
                End If
            End If
        End If
    End If
Next p
維基文庫等欲直接抽換之字 d
文字處理.書名號篇名號標注
d.Range.Cut
d.Close wdDoNotSaveChanges
SystemSetup.playSound 2
Exit Sub
eh:
Select Case Err.Number
    Case 5904 '無法編輯 [範圍]。
        If p.Range.Characters(acP).Hyperlinks.Count > 0 Then p.Range.Characters(acP).Hyperlinks(1).Delete
        Resume
    Case Else
        MsgBox Err.Number & Err.Description
End Select
End Sub

Sub 維基文庫四部叢刊本轉來_early()
Dim d As Document, a, i

a = Array("^p^p", "@", "〈", "{{", "〉", "}}", "^p", "", "}}{{", "^p", "@", "^p", _
    "○", ChrW(12295))
Set d = Documents.Add()
d.Range.Paste
維基文庫造字圖取代為文字 d.Range
For i = 0 To UBound(a) - 1
    d.Range.Find.Execute a(i), , , , , , True, wdFindContinue, , a(i + 1), wdReplaceAll
    i = i + 1
Next i
文字處理.書名號篇名號標注
d.Range.Cut
d.Close wdDoNotSaveChanges
Beep
End Sub

Sub 史記三家注()
'從2858頁起，20210920:0817之後，改用臺師大附中同學吳恆昇先生《中華文化網》所錄中研院《瀚典》初本，雖或仍未精，然至少免有簡化字轉換訛窘或造字亂碼之困擾，原文字檔棄置。根據初作比對，格式完全一樣！根本就是從這裡出來的，再轉簡化字，再又反正，造成之紊亂。悔當初沒想到用此本也。阿彌陀佛。佛弟子孫守真任真甫謹識於2021年9月20日
Dim d As Document, a, i, p As Paragraph, px As String, rng As Range, e As Long, pRng As Range, pa
'Const corTxt As String = "＝詳點校本校勘記＝"'該網站圖文對照排版功能未能配合，故今不採用。其格式只對文本版有效。https://ctext.org/instructions/wiki-formatting/zh
'a = Array(" ", "", "　　","","　", ChrW(-9217) & ChrW(-8195), "^p", "<p>^p",
'a = Array(" ", "", "　　", "", "^p^p", "<p>^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195),
a = Array("　　", "", "^p", "^p^p", "^p^p^p", "^p^p", "「^p^p", "「", "『^p^p", "『", "〔^p^p", "〔", "（^p^p", "（", _
    "^p^p", "<p>^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), _
    "^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195) & "〔", _
    "^p{{" & ChrW(-9217) & ChrW(-8195) & "{{{〈", _
    "「<p>^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), "「", _
    "〔<p>^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), "〔", _
    "『<p>^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), "『", _
    "（<p>^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), "（", _
    "集解", "《集解》：", "索隱", "《索隱》：", "【《索隱》：述贊】", "【《索隱》述贊】：", "正義", "《正義》：", _
    "九州島", "九州", "齊愍", "齊湣", "愍王", "湣王", "安厘王", "安釐王", _
    "塚", "冢", "慚", ChrW(24921), "啟", ChrW(21843), _
     ChrW(-30641), ChrW(-25066), _
     "群", ChrW(32675), "即", ChrW(21373), "眾", ChrW(-30650), _
     "既", ChrW(26083), "概", ChrW(27114), "溉", ChrW(28433), _
     "衛", ChrW(-30626), _
     "真", ChrW(30494), "填", ChrW(22625), "清", ChrW(28152), "青", ChrW(-26799), "教", ChrW(25934), _
    "鄉", ChrW(-28395), "鎮", ChrW(-27731), "慎", ChrW(24892), _
    "并", ChrW(24183), "屏", ChrW(23643), "荊", ChrW(-31930), "邢", ChrW(-28471), "笄", ChrW(31571), _
    "犁", ChrW(29314), "鬥", ChrW(-25811), "綿", ChrW(32220), _
    "冉", ChrW(20868), "腳", ChrW(-32486), _
    ChrW(25995), ChrW(-24956))
Set d = Documents.Add()
d.Range.Paste
維基文庫造字圖取代為文字 d.Range
d.Range.Cut
d.Range.PasteAndFormat wdFormatPlainText
d.Range.Text = VBA.Replace(d.Range.Text, " ", "")
For i = 0 To UBound(a) - 1
    If a(i) = "^p^p^p" Then
        px = d.Range.Text
        Do While InStr(px, Chr(13) & Chr(13) & Chr(13))
            px = Replace(px, Chr(13) & Chr(13) & Chr(13), Chr(13) & Chr(13))
        Loop
        d.Range.Text = px
        'Set rng = d.Range
'        Do While rng.Find.Execute(a(i), , , , , , True, wdFindContinue, , a(i + 1), wdReplaceAll)
'            If rng.End = d.Range.End Then Exit Do
'        Loop
    Else
        d.Range.Find.Execute a(i), , , , , , True, wdFindContinue, , a(i + 1), wdReplaceAll
    End If
    i = i + 1
Next i
文字處理.書名號篇名號標注
Set rng = Selection.Range
For Each p In d.Paragraphs
    px = p.Range.Text
    If Left(px, 7) = "{{" & ChrW(-9217) & ChrW(-8195) & "{{{" Then '注腳段落
        e = p.Range.Characters(1).End
        rng.SetRange e, e
        rng.MoveEndUntil "〕"
        If rng.Next.Next = "　" Then rng.Next.Next.Delete
        If InStr(p.Range.Text, "　") Then
            For Each pa In p.Range.Characters
                If pa = "　" Then
                    pa.Text = ChrW(-9217) & ChrW(-8195)
                End If
            Next
'            p.Range.text = VBA.Replace(p.Range.text, "　", ChrW(-9217) & ChrW(-8195))
'            'replace the text of paragraph the paragraph will be move to next one
'            Set p = p.Previous
'            e = p.Range.Characters(1).End
'            rng.SetRange e, e
'            rng.MoveEndUntil "〕"
        End If
        'rng.Select
        rng.Collapse wdCollapseEnd
        rng.Select
        Selection.MoveRight wdCharacter, 1, wdExtend
        Selection.TypeText "〉}}}" '將注腳編號〔一〕的右邊〕改成}}}
        px = p.Range.Text
        If InStr(Right(px, 4), "<p>") Then
            e = p.Range.Characters(p.Range.Characters.Count - 4).End
        Else
            e = p.Range.Characters(p.Range.Characters.Count - 1).End
        End If
        rng.SetRange e, e
        rng.InsertAfter "}}"
    Else '正文段落
        e = p.Range.Characters(1).start
        Set pRng = p.Range
        Do While InStr(pRng.Text, "〔")
            rng.SetRange e, e
            rng.MoveEndUntil "〔"
            If rng.Characters(rng.Characters.Count) <> "）" Then  ' if not correction
                rng.Collapse wdCollapseEnd
                rng.move , 1
                rng.MoveEnd wdCharacter, 1
                If rng.Text Like "[一二三四五六七八九]" Then  ' is footnote No.
                    e = rng.start
                    'rng.Collapse wdCollapseEnd
                    rng.SetRange e - 1, e
                    rng.Text = "　{{{〈"
                    rng.MoveEndUntil "〕"
                    rng.Collapse wdCollapseEnd
                    rng.MoveEnd wdCharacter, 1
                    rng.Text = "〉}}}"
                Else 'is correction to insert words
'                    rng.MoveEndUntil "〕"
'                    rng.SetRange rng.End + 2, rng.End + 2
'                    rng.InsertAfter corTxt
                End If
                e = rng.End
            Else 'is correction
'                If rng.Characters(rng.Characters.Count).Next = "〔" Then ' delete and insert words
'                    rng.MoveEndUntil "〕"
'                    rng.SetRange rng.End + 2, rng.End + 2
'                End If
'                rng.InsertAfter corTxt
               e = rng.End + 1
            End If
            'e = rng.End
            pRng.SetRange e, p.Range.End
            'pRng.SetRange rng.End, p.Range.End
            
        Loop
    End If
    If VBA.Left(p.Range.Text, 9) = ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195) & "【《索隱》" Then
        Set rng = p.Range
        p.Range.Characters(1).Delete
        rng.SetRange p.Range.start, p.Range.start
        rng.InsertAfter "{{"
        rng.SetRange p.Range.Characters(p.Range.Characters.Count - 4).End, p.Range.Characters(p.Range.Characters.Count - 4).End
        rng.InsertAfter "}}"
        If Len(rng.Paragraphs(1).Next.Range.Text) = 1 Then rng.Paragraphs(1).Next.Range.Delete
    End If
    
    If Len(p.Range) < 20 Then
        If (InStr(p.Range, "《史記》卷") Or VBA.Left(p.Range.Text, 3) = "史記卷") And InStr(p.Range, "*") = 0 Then
            rng.SetRange p.Range.start, p.Range.start
            rng.InsertAfter "*"
            For Each pa In p.Range.Characters
                    If pa Like "[〈《》〉]" Or StrComp(pa, ChrW(-9217) & ChrW(-8195)) = 0 Then pa.Delete
            Next pa
            '以下方式會造成p 值被設定為下一個段落
'            p.Range.text = VBA.Replace(p.Range.text, ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), "")
'            p.Range.text = VBA.Replace(VBA.Replace(p.Range.text, "《", ""), "》", "")
        End If
    End If
    If Len(p.Range) < 25 Then
        If VBA.InStr(p.Range.Text, "第") And InStr(p.Range, "*") = 0 _
                And (InStr(p.Range, "本紀") Or InStr(p.Range, "書") Or InStr(p.Range, "表") _
                Or InStr(p.Range, "世家") Or InStr(p.Range, "列傳")) Then
            rng.SetRange p.Range.start, p.Range.start
            rng.InsertAfter "　*"
            For Each pa In p.Range.Characters
                If pa Like "[〈《》〉]" Or StrComp(pa, ChrW(-9217) & ChrW(-8195)) = 0 Then pa.Delete
            Next pa
   
'            p.Range.text = VBA.Replace(p.Range.text, ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), "　*")
'            p.Range.text = VBA.Replace(VBA.Replace(p.Range.text, "〈", ""), "〉", "")
        End If
    End If

Next p
If VBA.Left(d.Paragraphs(1).Range.Text, 3) = "史記卷" And InStr(d.Paragraphs(1).Range.Text, "*") = 0 Then
    Set p = d.Paragraphs(1)
    rng.SetRange p.Range.start, p.Range.start
    rng.InsertAfter "*"
'    rng.SetRange p.Range.Characters(p.Range.Characters.Count - 1).End, p.Range.Characters(p.Range.Characters.Count - 1).End
'    rng.InsertAfter "<p>"
End If
If VBA.InStr(d.Paragraphs(2).Range.Text, "第") And InStr(d.Paragraphs(2).Range.Text, "*") = 0 Then
    Set p = d.Paragraphs(2)
'    rng.SetRange p.Range.start, p.Range.start
'    rng.InsertAfter "　*"
''    rng.SetRange p.Range.Characters(p.Range.Characters.Count - 1).End, p.Range.Characters(p.Range.Characters.Count - 1).End
''    rng.InsertAfter "<p>"
    p.Range.Text = VBA.Replace(p.Range.Text, ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), "　*")
    Set p = d.Paragraphs(2)
    p.Range.Text = VBA.Replace(VBA.Replace(p.Range.Text, "〈", ""), "〉", "")
End If

'Set rng = d.Range
'Do While rng.Find.Execute("〕", , , , , , True, wdFindStop)
'    If rng.Characters(1).Next <> "＝" Then rng.InsertAfter corTxt
'Loop
'Set rng = d.Range
'Do While rng.Find.Execute("）", , , , , , True, wdFindStop)
'    If InStr("＝〔", rng.Characters(1).Next) = 0 Then rng.InsertAfter corTxt
'Loop
d.Range.Cut
d.Close wdDoNotSaveChanges
Beep
word.Application.ActiveWindow.WindowState = wdWindowStateMinimize
End Sub
Sub 史記三家注2old()
'從2858頁起，20210920:0817之後，改用臺師大附中同學吳恆昇先生《中華文化網》所錄中研院《瀚典》初本，雖或仍未精，然至少免有簡化字轉換訛窘或造字亂碼之困擾，原文字檔棄置。根據初作比對，格式完全一樣！根本就是從這裡出來的，再轉簡化字，再又反正，造成之紊亂。悔當初沒想到用此本也。阿彌陀佛。佛弟子孫守真任真甫謹識於2021年9月20日
Dim d As Document, a, i, p As Paragraph, px As String, rng As Range, e As Long, pRng As Range
'Const corTxt As String = "＝詳點校本校勘記＝"'該網站圖文對照排版功能未能配合，故今不採用。其格式只對文本版有效。https://ctext.org/instructions/wiki-formatting/zh
a = Array("^p^p", "<p>^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), _
    "^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195) & "〔", _
    "^p{{" & ChrW(-9217) & ChrW(-8195) & "{{{〈", _
    "「<p>^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), "「", _
    "〔<p>^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), "〔", _
    "『<p>^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), "『", _
    "（<p>^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), "（", _
    "集解", "《集解》：", "索隱", "《索隱》：", "【《索隱》：述贊】", "【《索隱》述贊】：", "正義", "《正義》：", _
    "九州島", "九州", "齊愍", "齊湣", "愍王", "湣王", "安厘王", "安釐王", _
    "塚", "冢", _
     "群", ChrW(32675), "即", ChrW(21373), "眾", ChrW(-30650), "既", ChrW(26083), "衛", ChrW(-30626), _
     "真", ChrW(30494), "填", ChrW(22625), "清", ChrW(28152), "青", ChrW(-26799), "教", ChrW(25934), _
    "鄉", ChrW(-28395), "鎮", ChrW(-27731), "慎", ChrW(24892), "屏", ChrW(23643), "概", ChrW(27114), _
    "荊", ChrW(-31930), "邢", ChrW(-28471))
Set d = Documents.Add()
d.Range.Paste
For i = 0 To UBound(a) - 1
    d.Range.Find.Execute a(i), , , , , , True, wdFindContinue, , a(i + 1), wdReplaceAll
    i = i + 1
Next i
文字處理.書名號篇名號標注
Set rng = Selection.Range
For Each p In d.Paragraphs
    px = p.Range.Text
    If Left(px, 7) = "{{" & ChrW(-9217) & ChrW(-8195) & "{{{" Then '注腳段落
        e = p.Range.Characters(1).End
        rng.SetRange e, e
        rng.MoveEndUntil "〕"
        'rng.Select
        rng.Collapse wdCollapseEnd
        rng.Select
        Selection.MoveRight wdCharacter, 1, wdExtend
        Selection.TypeText "〉}}}" '將注腳編號〔一〕的右邊〕改成}}}
        px = p.Range.Text
        If InStr(Right(px, 4), "<p>") Then
            e = p.Range.Characters(p.Range.Characters.Count - 4).End
        Else
            e = p.Range.Characters(p.Range.Characters.Count - 1).End
        End If
        rng.SetRange e, e
        rng.InsertAfter "}}"
    Else '正文段落
        e = p.Range.Characters(1).start
        Set pRng = p.Range
        Do While InStr(pRng.Text, "〔")
            rng.SetRange e, e
            rng.MoveEndUntil "〔"
            If rng.Characters(rng.Characters.Count) <> "）" Then  ' if not correction
                rng.Collapse wdCollapseEnd
                rng.move , 1
                rng.MoveEnd wdCharacter, 1
                If rng.Text Like "[一二三四五六七八九]" Then  ' is footnote No.
                    e = rng.start
                    'rng.Collapse wdCollapseEnd
                    rng.SetRange e - 1, e
                    rng.Text = "　{{{〈"
                    rng.MoveEndUntil "〕"
                    rng.Collapse wdCollapseEnd
                    rng.MoveEnd wdCharacter, 1
                    rng.Text = "〉}}}"
                Else 'is correction to insert words
'                    rng.MoveEndUntil "〕"
'                    rng.SetRange rng.End + 2, rng.End + 2
'                    rng.InsertAfter corTxt
                End If
                e = rng.End
            Else 'is correction
'                If rng.Characters(rng.Characters.Count).Next = "〔" Then ' delete and insert words
'                    rng.MoveEndUntil "〕"
'                    rng.SetRange rng.End + 2, rng.End + 2
'                End If
'                rng.InsertAfter corTxt
               e = rng.End + 1
            End If
            'e = rng.End
            pRng.SetRange e, p.Range.End
            'pRng.SetRange rng.End, p.Range.End
            
        Loop
    End If
    If VBA.Left(p.Range.Text, 9) = ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195) & "【《索隱》" Then
        Set rng = p.Range
        p.Range.Characters(1).Delete
        rng.SetRange p.Range.start, p.Range.start
        rng.InsertAfter "{{"
        rng.SetRange p.Range.Characters(p.Range.Characters.Count).End, p.Range.Characters(p.Range.Characters.Count).End
        rng.InsertAfter "}}"
    End If
Next p
If VBA.Left(d.Paragraphs(1).Range.Text, 3) = "史記卷" Then
    Set p = d.Paragraphs(1)
    rng.SetRange p.Range.start, p.Range.start
    rng.InsertAfter "*"
    rng.SetRange p.Range.Characters(p.Range.Characters.Count - 1).End, p.Range.Characters(p.Range.Characters.Count - 1).End
    rng.InsertAfter "<p>"
End If
If VBA.InStr(d.Paragraphs(2).Range.Text, "第") Then
    Set p = d.Paragraphs(2)
    rng.SetRange p.Range.start, p.Range.start
    rng.InsertAfter "　*"
'    rng.SetRange p.Range.Characters(p.Range.Characters.Count - 1).End, p.Range.Characters(p.Range.Characters.Count - 1).End
'    rng.InsertAfter "<p>"
    p.Range.Text = VBA.Replace(VBA.Replace(p.Range.Text, "〈", ""), "〉", "")
End If


'Set rng = d.Range
'Do While rng.Find.Execute("〕", , , , , , True, wdFindStop)
'    If rng.Characters(1).Next <> "＝" Then rng.InsertAfter corTxt
'Loop
'Set rng = d.Range
'Do While rng.Find.Execute("）", , , , , , True, wdFindStop)
'    If InStr("＝〔", rng.Characters(1).Next) = 0 Then rng.InsertAfter corTxt
'Loop
d.Range.Cut
d.Close wdDoNotSaveChanges
Beep
End Sub

Sub 史記三家注1old()
Dim d As Document, a, i, p As Paragraph, px As String, rng As Range, e As Long
a = Array("<p>{{{", "<p>^p{{" & ChrW(-9217) & ChrW(-8195) & "{{{", _
        "<p>", "<p>^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), _
        ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195) & "^p{{" & ChrW(-9217) & ChrW(-8195), _
        "{{" & ChrW(-9217) & ChrW(-8195))
Set d = Documents.Add()
d.Range.Paste
For i = 0 To UBound(a) - 1
    d.Range.Find.Execute a(i), , , , , , True, wdFindContinue, , a(i + 1), wdReplaceAll
    i = i + 1
Next i
文字處理.書名號篇名號標注
d.Range.Find.Execute "《《", , , , , , True, wdFindContinue, , "《", wdReplaceAll
d.Range.Find.Execute "》》", , , , , , True, wdFindContinue, , "》", wdReplaceAll
d.Range.Find.Execute "〈〈", , , , , , True, wdFindContinue, , "〈", wdReplaceAll
d.Range.Find.Execute "〉〉", , , , , , True, wdFindContinue, , "〉", wdReplaceAll
Set rng = Selection.Range
For Each p In d.Paragraphs
    px = p.Range.Text
    If Left(p.Range.Text, 7) = "{{" & ChrW(-9217) & ChrW(-8195) & "{{{" Then
        If InStr(Right(px, 4), "<p>") Then
            e = p.Range.Characters(p.Range.Characters.Count - 4).End
        Else
            e = p.Range.Characters(p.Range.Characters.Count - 1).End
        End If
        rng.SetRange e, e
        rng.InsertAfter "}}"
    End If
    
Next p

d.Range.Cut
d.Close wdDoNotSaveChanges
End Sub
Sub 表sub()
Dim p As Paragraph, d As Document, rng As Range, s As Long, e As Long
Set d = Documents.Add(): Set rng = d.Range
d.Range.Paste
For Each p In d.Paragraphs
    If InStr(p.Range, "《索隱》：") Or _
        InStr(p.Range, "《正義》：") Or _
        InStr(p.Range, "《集解》：") Then
        If InStr(p.Range, "{{") = 0 Then
            s = p.Range.Characters(1).start
            rng.SetRange s, s
            rng.InsertBefore "{{"
            e = p.Range.Characters(p.Range.Characters.Count - 4).End
            rng.SetRange e, e
            rng.InsertAfter "}}"
        End If
    End If
Next p
d.Range.Cut
d.Close wdDoNotSaveChanges
Beep
End Sub

Sub 表sub1()
Dim d As Document, rng As Range, rngLast As Range, s As Long, e As Long
Set d = ActiveDocument
Set rng = d.Range: Set rngLast = rng
With rng.Find
    .Font.Color = 10092543
    .Font.Size = 10
    .Forward = True
    Do
        .Execute , , , , , , , wdFindStop
        If InStr(rng, "}}") Then
            .Execute , , , , , , , wdFindStop
            If InStr(rng, "}}") Then Exit Do
        End If
        s = rng.Characters(1).start
        e = rng.Characters(rng.Characters.Count - 1).End
        rngLast.SetRange e - 1, e
        rngLast.InsertAfter "}}"
        rngLast.SetRange s, s
        rngLast.InsertBefore "{{" & ChrW(-9217) & ChrW(-8195)
'        rng.SetRange rng.End + 222, d.Range.End
        
    Loop 'Until InStr(rng, "{{")
    .ClearFormatting
End With
Beep
End Sub

Sub 讀史記三家注()
Dim d As Document, t As Table
Set d = Documents.Add
d.Range.Paste
Set t = d.Tables(1)
With t
    .Columns(1).Delete
    .ConvertToText wdSeparateByParagraphs
End With
d.Range.Cut
d.Close wdDoNotSaveChanges
If word.Application.Windows.Count > 0 Then word.Application.ActiveWindow.WindowState = wdWindowStateMinimize
End Sub

Sub 戰國策_四部叢刊_維基文庫本() '《戰國策》格式者皆適用（即主文首行頂格，而其餘內容降一格者）
'https://ctext.org/library.pl?if=gb&res=77385
Dim a, rng As Range, rngDoc As Range, p As Paragraph, i As Long, rngCnt As Integer, ok As Boolean
Dim omits As String
omits = "《》〈〉「」『』·" & Chr(13)
Set rngDoc = Documents.Add.Range
re:
rngDoc.Paste
維基文庫造字圖取代為文字 rngDoc
For Each p In rngDoc.Paragraphs
    Set a = p.Range.Characters(1)
    If a <> "　" Then a.InsertBefore "　"
Next p
For Each a In rngDoc.Characters
    If Not a.Next Is Nothing And Not a.Previous Is Nothing Then
        If a = "　" And a.Next <> "　" And a.Previous <> "　" Then
            If a.Previous <> Chr(13) Then a.InsertBefore Chr(13)
            Set a = a.Next
        End If
    End If
Next a

For Each p In rngDoc.Paragraphs
    Set rng = p.Range
    If StrComp(rng.Characters(1), "　") = 0 And InStr(rng, "}") > 0 Then
        If rng.Characters(1) = "　" And rng.Characters(2) = "{" And rng.Characters(3) = "{" Then
            rng.Characters(1) = "{": rng.Characters(2) = "{": rng.Characters(3) = "　"
            For Each a In rng.Characters
               i = i + 1
               If rng.Characters(i) = "}" Then Exit For
               If rng.Characters(i) = Chr(13) Then
                    i = 0
                    Exit For
               End If
            Next a
        Else
            For Each a In rng.Characters
               i = i + 1
               If rng.Characters(i) = "}" Then Exit For
               If rng.Characters(i) = Chr(13) Or rng.Characters(i) = "{" Then
                    i = 0
                    Exit For
               End If
            Next a
        End If
        If i <> 0 Then
            If rng.Characters(1) = "{" And rng.Characters(2) = "{" And rng.Characters(3) = "　" Then
                rng.SetRange rng.Characters(3).End, rng.Characters(i).start
            Else
                rng.SetRange rng.Characters(1).End, rng.Characters(i).start
            End If
'            rng.Select
'            Stop
            rngCnt = rng.Characters.Count
            If rngCnt > 1 Then
                i = 0
                For Each a In rng.Characters
                    If InStr(omits, a) = 0 Then i = i + 1
                Next a
                rngCnt = i: i = 0
                If rngCnt Mod 2 = 1 Then
                    rngCnt = (rngCnt + 1) / 2
                Else
                    rngCnt = rngCnt / 2
                End If
                For Each a In rng.Characters
                    If InStr(omits, a) = 0 Then i = i + 1
                    If i = rngCnt Then
                        a.InsertAfter "　"
                        Exit For
                    End If
                Next a
'                If rngCnt Mod 2 = 1 Then
'                    If rng.Characters((rngCnt - rngCnt Mod 2) / 2 + 1).Next <> "　" _
'                        Then rng.Characters((rngCnt - rngCnt Mod 2) / 2 + 1).InsertAfter "　"
'
'                Else
'                    If rng.Characters((rngCnt - rngCnt Mod 2) / 2).Next <> "　" _
'                        Then rng.Characters((rngCnt - rngCnt Mod 2) / 2).InsertAfter "　"
'                End If
            Else
                rng.Characters(1).InsertAfter "　"
            End If
        End If
        i = 0
    End If
Next
If ok Then
    For Each p In rngDoc.Paragraphs
        If Left(p.Range.Text, 3) = "{{　" And p.Range.Characters(p.Range.Characters.Count - 1) = "}" Then
            a = p.Range.Text
            a = Mid(a, 4, Len(a) - 6)
            If InStr(a, "　") > 0 And InStr(a, "{") = 0 And InStr(a, "}") = 0 Then
                rngCnt = p.Range.Characters.Count
                For i = 4 To rngCnt
                    Set a = p.Range.Characters(i)
                    If a = "　" Then
                        a.InsertParagraphBefore
                        Exit For
                    End If
                Next i
            End If
        End If
    Next p
    '以下3行《戰國策》本身才需要
'    rngDoc.Find.Execute "正曰", , , , , , , wdFindContinue, , "【正曰】", wdReplaceAll
'    rngDoc.Find.Execute ChrW(-10155) & ChrW(-8585) & "曰", , , , , , , wdFindContinue, , "【" & ChrW(-10155) & ChrW(-8585) & "曰】", wdReplaceAll
'    rngDoc.Find.Execute "補曰", , , , , , , wdFindContinue, , "【" & ChrW(-10155) & ChrW(-8585) & "曰】", wdReplaceAll
End If
If ok Then 文字處理.書名號篇名號標注
rngDoc.Cut
If Not ok Then
    DoEvents
    rngDoc.PasteAndFormat wdFormatPlainText
    rngDoc.Find.Execute "〈", , , , , , , wdFindContinue, , "{{", wdReplaceAll
    rngDoc.Find.Execute "〉", , , , , , , wdFindContinue, , "}}", wdReplaceAll
    rngDoc.Cut
    ok = True
    GoTo re
End If
rngDoc.Document.Close wdDoNotSaveChanges
On Error Resume Next
AppActivate "TextForCtext"
SendKeys "%{insert}", True
SystemSetup.playSound 4
End Sub
Sub 維基文庫造字圖取代為文字(rng As Range)
Dim inlnsp As InlineShape, aLtTxt As String
Dim dictMdb As New dBase, cnt As New ADODB.Connection, rst As New ADODB.Recordset
dictMdb.cnt查字 cnt
For Each inlnsp In rng.InlineShapes
    aLtTxt = inlnsp.AlternativeText
    If Len(aLtTxt) < 3 Then
        'inlnsp.Delete
    Else
        If aLtTxt Like "?酉?? -- 醢" Then
            aLtTxt = "醢"
        ElseIf aLtTxt Like "揚 --（『昜』上『旦』之『日』與『一』相連）" Then
            aLtTxt = "揚"
        ElseIf aLtTxt Like "（??石）" Then
            aLtTxt = "若"
        ElseIf aLtTxt Like "??皿" Then
            aLtTxt = "盟"
        ElseIf aLtTxt Like "??? -- 溥" Then
            aLtTxt = "溥"
        ElseIf aLtTxt Like "場 --（『昜』上『旦』之『日』與『一』相連）" Then
            aLtTxt = "場"
        ElseIf aLtTxt Like "?????禾 -- 蘇" Then
            aLtTxt = "蘇"
        ElseIf aLtTxt Like ChrW(12272) & ChrW(-10155) & ChrW(-8696) & ChrW(31860) Then
            aLtTxt = "隸"
        ElseIf aLtTxt Like "彎（?弓爪）-- 弧莫不投" Then
            aLtTxt = "弧"
        ElseIf aLtTxt Like "?土? -- 坳" Then
            aLtTxt = "坳"
        ElseIf aLtTxt Like "????口?欠 -- " & ChrW(-10111) & ChrW(-8620) Then
            aLtTxt = ChrW(-10111) & ChrW(-8620)
        ElseIf aLtTxt Like "???? -- 攙" Then
            aLtTxt = "攙"
        ElseIf aLtTxt Like "?????? -- 掾" Then
            aLtTxt = "掾"
        ElseIf aLtTxt Like "???? -- 詣" Then
            aLtTxt = "詣"
        ElseIf aLtTxt Like "????? --狖" Then
            aLtTxt = "狖"
        ElseIf aLtTxt Like "?馬? -- 驂" Then
            aLtTxt = "驂"
        ElseIf aLtTxt Like "?日?? -- 暝" Then
            aLtTxt = "暝"
        ElseIf aLtTxt Like "???? -- 溟" Then
            aLtTxt = "溟"
        ElseIf aLtTxt Like "（?厂雝）" Then
            aLtTxt = "廱"
        ElseIf aLtTxt Like "??乃??皿 -- 盈" Then
            aLtTxt = "盈"
        ElseIf aLtTxt Like "叟 -- 臾 ?" Then
            aLtTxt = ChrW(-10114) & ChrW(-9161)
        ElseIf aLtTxt Like "愓 --（『昜』上『旦』之『日』與『一』相連）" Then
            aLtTxt = "愓"
        ElseIf aLtTxt Like "場 --（『昜』上『旦』之『日』與『一』相連）" Then
            aLtTxt = "場"
        ElseIf aLtTxt Like "暘 --（『昜』上『旦』之『日』與『一』相連）" Then
            aLtTxt = "暘"
        ElseIf aLtTxt Like "煬(「旦」改為「??」)" Then
            aLtTxt = "煬"
        ElseIf aLtTxt Like "錫 --（右上『日』字下一?長出，類似『旦』字的『日』與『一』相連）" Then
            aLtTxt = "錫"
        ElseIf aLtTxt Like ChrW(24298) & "（" & ChrW(8220) & ChrW(13357) & ChrW(8221) & "換為" & ChrW(8220) & "面" & ChrW(8221) & "）" Then
            aLtTxt = "廩"
        ElseIf aLtTxt Like ChrW(12273) & ChrW(11966) & ChrW(30464) Then
            aLtTxt = "萌"
        ElseIf aLtTxt Like "?彳? -- 徊" Then
            aLtTxt = "徊"
        ElseIf aLtTxt Like ChrW(12272) & ChrW(-10145) & ChrW(-8265) & "變" Then
            aLtTxt = "●＝" & aLtTxt & "＝"
        ElseIf aLtTxt Like "? -- or ?? ?" Then
            aLtTxt = ChrW(-32119)
        ElseIf aLtTxt Like "輕" Then
            aLtTxt = ChrW(18518)
        ElseIf aLtTxt Like "能" Then
            aLtTxt = ChrW(17403)
        ElseIf aLtTxt Like ChrW(12272) & ChrW(-10145) & ChrW(-8265) & ChrW(25908) Then
            aLtTxt = ChrW(-10109) & ChrW(-8699)
        ElseIf aLtTxt Like "??八 -- " & ChrW(-10170) & ChrW(-8693) Then
            aLtTxt = ChrW(-10124) & ChrW(-9097)
        ElseIf aLtTxt Like ChrW(12282) & ChrW(-28746) & "商" Then
            aLtTxt = "適"
        ElseIf aLtTxt Like "??？ -- 狐" Then
            aLtTxt = "狐"
        ElseIf aLtTxt Like "??戔 -- 殘" Then
            aLtTxt = "殘"
        ElseIf aLtTxt Like "?????匹 -- 繼" Then
            aLtTxt = "繼"
        ElseIf aLtTxt Like "???么 -- " & ChrW(31762) Then
            aLtTxt = "篡"
        ElseIf aLtTxt Like "????凡 -- 彘" Then
            aLtTxt = "彘"
        ElseIf aLtTxt Like "?麻止 -- ?" Then
            aLtTxt = "歷"
        ElseIf aLtTxt Like ChrW(12282) & ChrW(-28746) & ChrW(17807) Then
            aLtTxt = "遽"
        ElseIf aLtTxt Like "?至支 -- ??" Then
            aLtTxt = "致"
        ElseIf aLtTxt Like "（???女）" Then
            aLtTxt = "嫈"
        ElseIf aLtTxt Like "（???力）" Then
            aLtTxt = ChrW(-10174) & ChrW(-9072)
        ElseIf aLtTxt Like "??? -- 懈" Then
            aLtTxt = "懈"
        ElseIf aLtTxt Like "（???）-- 釵" Then
            aLtTxt = "釵"
        ElseIf aLtTxt Like "?目兆 -- 晁" Then
            aLtTxt = "晁"
        ElseIf aLtTxt Like "???? -- " & ChrW(-10161) & ChrW(-8272) Then
            aLtTxt = "漆"
        ElseIf aLtTxt Like "?口? -- 噦" Then
            aLtTxt = "噦"
        ElseIf aLtTxt Like "?口? -- 呦" Then
            aLtTxt = "呦"
        ElseIf aLtTxt Like "???? -- 指" Then
            aLtTxt = "指"
        ElseIf aLtTxt Like "?夸?? -- 瓠" Then
            aLtTxt = ChrW(-10158) & ChrW(-8444)
        ElseIf aLtTxt Like "*page2700-20px-SKQSfont.pdf.jpg*" Then
            aLtTxt = "劇"
        ElseIf aLtTxt Like ChrW(12273) & ChrW(11966) & ChrW(12272) & ChrW(27701) & ChrW(20158) Then
            aLtTxt = ChrW(-10161) & ChrW(-8915)
        ElseIf aLtTxt Like "???止自匕?儿? -- 夔" Then
            aLtTxt = "夔"
        ElseIf aLtTxt Like "?穴之 -- 窆" Then
            aLtTxt = "窆"
        ElseIf aLtTxt Like ChrW(12272) & "目" & ChrW(-10170) & ChrW(-8693) Then
            aLtTxt = ChrW(-10121) & ChrW(-8228)
        ElseIf aLtTxt Like "??? -- 潤" Then
            aLtTxt = "潤"
        ElseIf aLtTxt Like "??? -- 靦" Then
            aLtTxt = "靦"
        ElseIf aLtTxt Like "??向 -- " & ChrW(-28664) Then
            aLtTxt = "迥"
        ElseIf aLtTxt Like "?日黽 -- " & ChrW(-24830) Then
            aLtTxt = ChrW(-24830)
        ElseIf aLtTxt Like "???????友-- 擾" Then
            aLtTxt = "擾"
        ElseIf aLtTxt Like "??? -- 癩" Then
            aLtTxt = "癩"
        ElseIf aLtTxt Like "（?血?）" Then
            aLtTxt = ChrW(-30654)
        ElseIf aLtTxt Like "SKchar" Then
            GoTo nxt
'            aLtTxt = "疾,優,虢,曷,姬,鮑,徑,梓,死（2DB7E）,鬼,灌,瓘,鸛,毓,褭,舁"'餘詳 查字.mdb
        ElseIf aLtTxt Like "SKchar2" Then
            GoTo nxt
'            aLtTxt = "纏（7E92）,丑,"'餘詳 查字.mdb
        Else
            Select Case aLtTxt
                Case ChrW(12280) & ChrW(30098) & ChrW(-28523)
                    aLtTxt = "●＝" & aLtTxt & "＝"
                    '缺字則直接插入字圖替代文字
                    GoTo replaceIt
                Case Else
                    Dim rp As Boolean
                    rst.Open "select * from 維基文庫造字圖取代對照表 where (strcomp(find, """ & aLtTxt & """)=0 " & _
                        "and not find like ""SKchar*"") ", cnt, adOpenStatic, adLockReadOnly
'                    If rst.RecordCount > 0 Then
                    Do Until rst.EOF
                        aLtTxt = rst.Fields("replace").Value
                        rp = True
                        Exit Do
                    Loop
'                    Else
                        rst.Close
                        If Not rp Then
                            GoTo nxt
                        Else
                            rp = False
                        End If
'                    End If
'                    rst.Close
            End Select
        End If
    End If
replaceIt:
    inlnsp.Select
    Selection.TypeText aLtTxt
    inlnsp.Delete
nxt:
Next inlnsp
cnt.Close
End Sub
Sub 國學大師_四庫全書本轉來()
Dim rng As Range, noteRng As Range
Set rng = Documents.Add().Range
rng.Paste
With rng.Find
    .ClearAllFuzzyOptions
    .ClearFormatting
    .MatchWildcards = True
    .Execute "[[]*[]]  ", , , True, , , True, wdFindContinue, , "", wdReplaceAll
    .ClearAllFuzzyOptions
    .ClearFormatting
End With
rng.Find.Font.Color = 16711935
Do While rng.Find.Execute("", , , False, , , True, wdFindStop)
    Set noteRng = rng
    Do While noteRng.Next.Font.Color = 16711935
        noteRng.SetRange noteRng.start, noteRng.Next.End
    Loop
    noteRng.Text = "{{" & Replace(noteRng, "/", "") & "}}"
Loop
With rng.Document
    .Range.Cut
    .Close wdDoNotSaveChanges
End With
End Sub

Sub mdb開發_千慮一得齋Export()
Dim cnt As New ADODB.Connection, db As New dBase, rst As New ADODB.Recordset, exportStr As String, preTitle As String, title As String
Const bookName As String = "原抄本日知錄" '執行前請先指定書名
db.cnt_開發_千慮一得齋 cnt
rst.Open "SELECT 篇.篇名, 札.札記, 書.書名, 篇.卷, 篇.頁, 篇.末頁, 札.篇ID, 札.頁, 札.札ID, 札.類ID, 類別主題.類別主題" & _
        " FROM 類別主題 INNER JOIN ((書 INNER JOIN 篇 ON 書.書ID = 篇.書ID) INNER JOIN 札 ON 篇.篇ID = 札.篇ID) ON 類別主題.類ID = 札.類ID" & _
        " WHERE (((書.書名)=""" & bookName & """) AND ((類別主題.類別主題) Not Like "" * 真按 * "" Or (類別主題.類別主題) Is Null))" & _
        " ORDER BY 篇.卷, 篇.頁, 篇.末頁, 札.篇ID, 札.頁, 札.札ID;", cnt, adOpenKeyset, adLockReadOnly
Do Until rst.EOF
    title = rst.Fields(0).Value
    If preTitle <> title Then
        exportStr = exportStr & Chr(13) & "*" & title & Chr(13)
    End If
    preTitle = title
    exportStr = exportStr & rst.Fields(1).Value
    rst.MoveNext
Loop
rst.Close
cnt.Close
Documents.Add.Range = exportStr
End Sub
Sub 清除所有符號_加上井號_作為網址後綴()
Dim rng As Range, e, sybol 'Alt + l
sybol = Array("(", ")", "（", "）")
Set rng = Documents.Add().Range
rng.Paste
Docs.清除所有符號
For Each e In sybol
    rng.Text = Replace(rng, e, "")
Next e
rng.Text = "#" & rng.Text
rng.Cut
rng.Document.Close wdDoNotSaveChanges
DoEvents
AppActivateDefaultBrowser
SendKeys "^v~"
DoEvents
SendKeys "^l^c"
DoEvents
SendKeys "{F5}"
End Sub

Sub 維基文庫等欲直接抽換之字(d As Document)
Dim rst As New ADODB.Recordset, cnt As New ADODB.Connection, db As New dBase
db.cnt查字 cnt
rst.Open "select * from 維基文庫等欲直接抽換之字 where doIt = true", cnt, adOpenForwardOnly, adLockReadOnly
Do Until rst.EOF
    d.Range.Find.Execute rst.Fields("replaced").Value, , , , , , True, wdFindContinue, , rst.Fields("replacewith").Value, wdReplaceAll
    rst.MoveNext
Loop
rst.Close: cnt.Close: Set db = Nothing
End Sub

Sub dbSBCKWordtoReplace() '四部叢刊造字對照表 Alt+5
Dim rng As Range, ur As UndoRecord
Set ur = stopUndo("《四部叢刊》資料庫造字取代為系統字")
If ActiveDocument.Name = "《四部叢刊資料庫》補入《中國哲學書電子化計劃》.docm" Then
    Set rng = ActiveDocument.Range
Else
    Set rng = Documents.Add.Range
    rng.Paste
End If
dbSBCKWordtoReplaceSub rng
If Not ActiveDocument.Name = "《四部叢刊資料庫》補入《中國哲學書電子化計劃》.docm" Then
    rng.Cut
    If rng.Application.Documents.Count = 1 Then
        rng.Application.Quit wdDoNotSaveChanges
    Else
        rng.Document.Close wdDoNotSaveChanges
    End If
Else
    ActiveDocument.Save
End If
contiUndo ur
End Sub
Sub dbSBCKWordtoReplaceSub(ByRef rng As Range)
Const tbName As String = "四部叢刊造字對照表"
Dim rst As New ADODB.Recordset, cnt As New ADODB.Connection, db As New dBase
rng.Find.ClearFormatting
db.cnt查字 cnt
rst.Open tbName, cnt, adOpenForwardOnly, adLockReadOnly
Do Until rst.EOF
    If InStr(rng.Text, rst.Fields(0).Value) Then _
        rng.Find.Execute rst.Fields(0).Value, , , , , , True, wdFindContinue, , rst.Fields(1).Value, wdReplaceAll
    rst.MoveNext
Loop
rst.Close: cnt.Close: Set db = Nothing
End Sub

Sub dbSBCKWordtoReplace_AddNewOne() '四部叢刊造字對照表 Alt+4
Const tbName As String = "四部叢刊造字對照表"
Dim rst As New ADODB.Recordset, cnt As New ADODB.Connection, db As New dBase
Dim rng As Range
Set rng = Selection.Range
db.cnt查字 cnt
rst.Open "select * from " + tbName + " where strcomp(造字, """ + rng.Characters(1) + """)=0", cnt, adOpenKeyset, adLockOptimistic
If rst.RecordCount = 0 Then
    If rng.Characters.Count = 2 Then
        'rst.Open tbName, cnt, adOpenKeyset, adLockOptimistic
        rst.AddNew
        rst.Fields(0) = rng.Characters(1)
        rst.Fields(1) = rng.Characters(2)
        rst.Update
        rng.Characters(1).Delete
    Else
        MsgBox "plz input the replace word next the one"
        Selection.move
    End If
Else
    If rng.Characters.Count = 2 Then If rng.Characters(2) = rst.Fields(1).Value Then rng.Characters(2).Delete
    rng.Characters(1) = rst.Fields(1).Value
End If
rst.Close: cnt.Close: Set db = Nothing
dbSBCKWordtoReplaceSub rng.Document.Range
End Sub

Sub entity_Markup_edit_via_API_Annotate_Reverting()
Dim rng As Range, rngMark As Range, d As Document, ay(), e, i As Long, DoctoMarked As Document
Set d = ActiveDocument
Const markStrOpen As String = "<entity ", markStrClose As String = "</entity>"
If InStr(d.Range, markStrOpen) = 0 Then
    MsgBox "plz paste the marked text in active doc First thx"
    Exit Sub
End If
Set rng = d.Range
'get the terms which were marked
Do While rng.Find.Execute(markStrOpen)
    'rng.SetRange rng.start, rng.End + rng.MoveEndUntil(markStrClose)
    rng.SetRange rng.start, rng.End + rng.MoveEndUntil("/")
    If d.Range(rng.End, rng.End + 7) = "entity>" Then
        rng.SetRange rng.start, rng.End - 2
        ReDim Preserve ay(i)
        ay(i) = VBA.Split(rng.Text, ">")
        i = i + 1
    End If
    
    rng.SetRange rng.End, d.Range.End
Loop
'got the terms which were marked already
'mark the text
'Stop
If MsgBox("if NOT text to be marked already copied then push CANCEL button", vbOKCancel + vbExclamation) = vbCancel Then Exit Sub
Set DoctoMarked = Documents.Add
Set rng = DoctoMarked.Range: Set rngMark = DoctoMarked.Range
rng.Paste
For Each e In ay
reFind:
    If rng.Find.Execute(e(1)) Then
        If rng.Characters(1).Previous = ">" Then
            rngMark.SetRange rng.start - 1, rng.start
            'rngMark.MoveStartUntil "<"
            Do Until DoctoMarked.Range(rngMark.start, rngMark.start + 1) = "<"
                rngMark.move wdCharacter, -1
            Loop
            rngMark.SetRange rngMark.start, rng.start
            If Left(rngMark.Text, 8) <> "<entity " Then
                GoSub mark
            Else
                rng.SetRange rng.End, DoctoMarked.Range.End
                GoTo reFind
            End If
        Else
            GoSub mark
        End If
    Else
        SystemSetup.ClipboardPutIn CStr(e(1))
        MsgBox "plz check out why the " + e(1) + "dosen't exist !!", vbExclamation
        Stop
        Set rng = DoctoMarked.Range
        GoTo reFind
        'Exit Sub
    End If
    rng.SetRange rng.End, d.Range.End
Next e
Beep
Exit Sub
mark:
    rng.InsertAfter "</entity>"
    rng.InsertBefore e(0) + ">"
Return
End Sub

Sub checkEditingOfPreviousVersion()
Dim d As Document, rng As Range
Set d = Documents.Add()
Set rng = d.Range
rng.Paste
GoSub fontColor
GoSub punctuations
If d.Application.Documents.Count = 1 Then
    d.Application.Quit wdDoNotSaveChanges
Else
    d.Close wdDoNotSaveChanges
End If
Exit Sub
 
 
fontColor:

    rng.Find.ClearFormatting
    rng.Find.Font.Color = 8912896 '{{{}}}語法下的文字
    rng.Find.Replacement.ClearFormatting
    With rng.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    If (rng.Find.Execute) Then GoSub CheckOut
Return

punctuations:
    rng.Find.ClearFormatting
    rng.Find.Replacement.ClearFormatting
    Dim punctus, e
    punctus = Array("，", "。", "「", "·", "：", "（")  '檢查幾個具代表者即可
    For Each e In punctus
        If InStr(rng.Text, e) > 0 Then
            rng.Find.Execute e
            GoTo CheckOut
        End If
    Next e
Return

CheckOut:
    rng.Select
    d.ActiveWindow.Visible = True
    d.ActiveWindow.ScrollIntoView rng
    MsgBox "plz check it out !", vbExclamation
End Sub

Sub EditMakeup()
Const differPageNum  As Integer = 161 '頁數差
Dim rng As Range, pageNum As Range, d As Document, ur As UndoRecord
Set d = ActiveDocument
Set rng = d.Range
Set ur = SystemSetup.stopUndo("EditMakeupCtext")
Do While rng.Find.Execute(" page=""", , , , , , True, wdFindStop)
    Set pageNum = rng
    pageNum.SetRange rng.End, rng.End + 1
    pageNum.MoveEndUntil """"
    pageNum.Text = CStr(CInt(pageNum.Text) - differPageNum)
    rng.SetRange pageNum.End, d.Range.End
Loop
SystemSetup.contiUndo ur
End Sub


Sub tempReplaceTxtforCtextEdit()
Dim a, d As Document, i As Integer, x As String
a = Array("{{（", "{{", "）}}", "}}", "（", "{{", "）", "}}", "○", ChrW(12295))
Set d = Documents.Add
d.Range.Paste
x = d.Range
For i = 0 To UBound(a)
    x = Replace(x, a(i), a(i + 1))
    i = i + 1
Next i
d.Range = x
d.Range.Cut
d.Close wdDoNotSaveChanges
AppActivateDefaultBrowser
SendKeys "^v"
End Sub


Sub tempReplaceTxtforCtext() 'for Quick edit only
Dim a, d As Document, i As Integer
a = Array("{{（", "{{", "）}}", "}}", "（", "{{", "）", "}}", "○", ChrW(12295))
Set d = Documents.Add
d.Range.Paste
For i = 0 To UBound(a)
    d.Range.Find.Execute a(i), , , , , , , wdFindContinue, , a(i + 1), wdReplaceAll
    i = i + 1
Next i
d.Range.Cut
d.Application.WindowState = wdWindowStateMinimize
d.Close wdDoNotSaveChanges
AppActivateDefaultBrowser
SendKeys "^v"
SendKeys "{tab}~"

End Sub

