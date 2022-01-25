Attribute VB_Name = "中國哲學書電子化計劃"

Sub 新頁面()
'the page begin
Const start As Integer = 2375
' the page end
Const e As Integer = 2380
' the book
Const fileID As Long = 1000081
'https://ctext.org/library.pl?if=gb&file=1000081&page=2621

Dim x As String, data As New MSForms.DataObject
Dim i As Integer
For i = start To e
    x = x & "<scanbegin file=""" & fileID & """ page=""" & i & """ />" & Chr(9) & "<scanend file=""" & fileID & """ page=""" & i & """ />" '若中間沒有任何內容，頁面最後便不能成一段落。若剛好一個段落，會與下一頁黏合在一起
Next i


'For Each e In Selection.Value
'    x = x & e
'Next e
''x = Replace(x, Chr(13), "")
data.SetText Replace(x, "/>", "/>●", 1, 1)
data.PutInClipboard
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


Sub 維基文庫四部叢刊本轉來()
Dim d As Document, a, i

a = Array("^p^p", "@", "〈", "{{", "〉", "}}", "^p", "", "}}{{", "^p", "@", "^p", _
    "○", ChrW(12295))
Set d = Documents.Add()
d.Range.Paste
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

Sub 戰國策_四部叢刊_維基文庫本()
'https://ctext.org/library.pl?if=gb&res=77385
Dim a, rng As Range, rngDoc As Range, p As Paragraph, i As Long, rngCnt As Integer
Set rngDoc = Documents.Add.Range
rngDoc.Paste
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
        For Each a In rng.Characters
           i = i + 1
           If rng.Characters(i) = "}" Then Exit For
           If rng.Characters(i) = Chr(13) Or rng.Characters(i) = "{" Then
                i = 0
                Exit For
           End If
        Next a
        If i <> 0 Then
            rng.SetRange rng.Characters(1).End, rng.Characters(i).start
'            rng.Select
'            Stop
            rngCnt = rng.Characters.Count
            If rngCnt > 1 Then
                If rngCnt Mod 2 = 1 Then
                    If rng.Characters((rngCnt - rngCnt Mod 2) / 2 + 1).Next <> "　" _
                        Then rng.Characters((rngCnt - rngCnt Mod 2) / 2).InsertAfter "　"

                Else
                    If rng.Characters((rngCnt - rngCnt Mod 2) / 2).Next <> "　" _
                        Then rng.Characters((rngCnt - rngCnt Mod 2) / 2).InsertAfter "　"
                End If
            End If
        End If
        i = 0
    End If
Next
rngDoc.Cut
rngDoc.Document.Close wdDoNotSaveChanges
AppActivate "TextForCtext"
End Sub

Sub tempReplaceTxtforCtext()
Dim a, d As Document, i As Integer
a = Array("（", "", "）", "", "○", ChrW(12295))
Set d = Documents.Add
d.Range.Paste
For i = 0 To UBound(a)
    d.Range.Find.Execute a(i), , , , , , , wdFindContinue, , a(i + 1), wdReplaceAll
    i = i + 1
Next i
d.Range.Cut
d.Close wdDoNotSaveChanges
AppActivate "google chrome"
SendKeys "^v"
SendKeys "{tab}~"

End Sub

