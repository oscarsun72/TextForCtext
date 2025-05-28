Attribute VB_Name = "中國哲學書電子化計劃"
Option Explicit
Dim ChapterSelector As String
'Const description As String = "將星號前的分段符號移置前段之末 & 清除頁前的分段符號"
'Const description As String = "將星號前的分段符號移置前段之末 & 清除頁前的分段符號{佛弟子文獻學者孫守真任真甫按：仁者志士義民菩薩賢友請多利用賢超法師《古籍酷AI》或《看典古籍》OCR事半功倍也。如蒙不棄，可利用末學於GitHub開源免費免安裝之TextForCtext 應用程式，加速輸入與排版。討論區與末學YouTube頻道有演示影片可資參考。感恩感恩　讚歎讚歎　南無阿彌陀佛"
'Const description As String = "將星號前的分段符號移置前段之末 & 清除頁前的分段符號{據Kanripo.org或《國學大師》所藏本輔以末學自製於GitHub開源免費免安裝之TextForCtext排版對應錄入。討論區與末學YouTube頻道有實境演示影片可資參考。感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主}"
'Const description As String = "將星號前的分段符號移置前段之末 & 清除頁前的分段符號{據《國學大師》或北京元引科技有限公司《元引科技引得數字人文資源平臺·中國歷代文獻》所藏本輔以末學自製於GitHub開源免費免安裝之TextForCtext排版對應錄入。討論區與末學YouTube頻道有實境演示影片可資參考。感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主}"
Const description As String = "將星號前的分段符號移置前段之末 & 清除頁前的分段符號{據北京元引科技有限公司《元引科技引得數字人文資源平臺·中國歷代文獻》所藏本輔以末學自製於GitHub開源免費免安裝之TextForCtext排版對應錄入。討論區與末學YouTube頻道有實境演示影片可資參考。感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主}"

'Const description_Edit_textbox_新頁面 As String = "據《國學大師》或《Kanripo》所收本輔以末學自製於GitHub開源免費免安裝之TextForCtext軟件排版對應錄入；討論區及末學YouTube頻道有實境演示影片。感恩感恩　讚歎讚歎　南無阿彌陀佛"
Const description_Edit_textbox_新頁面 As String = "據北京元引科技有限公司《元引科技引得數字人文資源平臺·中國歷代文獻》所收本輔以末學自製於GitHub開源免費免安裝之TextForCtext軟件排版對應錄入；討論區及末學YouTube頻道有實境演示影片。感恩感恩　讚歎讚歎　南無阿彌陀佛"

Sub 分行分段()
    
    Dim lineLength As Byte, d As Document, rng As Range, si As New StringInfo, firstLineIndentValue As Single, leadSpaceCount As Byte, p As Paragraph, leadSpaces As String, i As Long, t As table
    
    lineLength = 21 ''第一行指定正常行長度: d.Paragraphs(1).Range.Characters.Count - 1
    'd.Paragraphs(1).Range.text = vbNullString
    
    Set d = Documents.Add
    d.Range.Paste
    
    
    For Each p In d.Paragraphs
        If p.Style = "內文" And p.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft Then
            firstLineIndentValue = p.Range.ParagraphFormat.FirstLineIndent
            If firstLineIndentValue <> 0 Then
                leadSpaceCount = VBA.Abs(firstLineIndentValue) / d.Paragraphs(1).Range.Characters(1).font.Size
            End If
            Exit For
        End If
    Next p
    'firstLineIndentValue = d.Paragraphs(1).Range.ParagraphFormat.FirstLineIndent
    
    
    For Each t In d.tables
        t.Delete
    Next t
    '清除 【圖】（《漢籍全文資料庫》文本，以其複製文字功能）
    If VBA.InStr(d.Range.text, "【圖】") Then d.Range.Find.Execute "【圖】", , , , , , , wdFindContinue, , vbNullString, wdReplaceAll
    d.Range.Find.Execute "^l", , , , , , , wdFindContinue, , vbNullString, wdReplaceAll
    For Each p In d.Paragraphs
        If p.Style = "內文" And p.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft Then
            If Not p.Next Is Nothing Then
                If p.Next.Range.text <> "．　．　．　．　．　．　．　．　．　．　．　．　．　．　．　．　．　．" & Chr(13) Then
                    p.Range.Characters(p.Range.Characters.Count).text = vbNullString
                    Set p = p.Previous
                Else
                    p.Next.Range.text = vbNullString
                End If
            End If
        End If
    Next p
'    d.Range.Find.Execute "．　．　．　．　．　．　．　．　．　．　．　．　．　．　．　．　．　．^p", , , , , , , wdFindContinue, , vbNullString, wdReplaceAll

    Dim lineCntr As Byte, noteCntr As Long
    For Each p In d.Paragraphs
        If p.Style = "內文" And p.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft Then
        
'            If InStr(p.Range, "一作原武") Then
'                p.Range.Select
'                Stop
'            End If

            If p.Range.Characters.Count - 1 > lineLength Then
                i = 0 ''i =  1 'lineLength
                Set rng = p.Range.Characters(1)
                Do While i + lineLength < p.Range.Characters.Count - 1 'Step lineLength 'p.Range.Characters.Count Step lineLength
lastfew:
                    
                    lineCntr = 0
                    Do While lineCntr < lineLength
                        If rng.font.Size = 7.5 Then '小注
                            noteCntr = noteCntr + 1
                            rng.Move wdCharacter, 1
                            i = i + 1
                            If rng.font.Size = 7.5 Then
                                noteCntr = noteCntr + 1
                                rng.Move wdCharacter, 1
                                i = i + 1
                            End If
                            'i = i + 2
                        Else
                            If noteCntr > 0 Then
                                If noteCntr Mod 2 = 1 Then lineCntr = lineCntr + 1
                                noteCntr = 0
                            End If
                            rng.Move wdCharacter, 1
                            i = i + 1
                        End If
                        lineCntr = lineCntr + 1
                    Loop
                    
'                    rng.Select
                    '如果有凸排
                    If leadSpaceCount > 0 Then
                        If VBA.InStr(p.Range.text, Chr(11)) Then
                            'p.Range.Characters(i - leadSpaceCount).InsertAfter Chr(11)
                            rng.Move wdCharacter, -leadSpaceCount
                            rng.InsertAfter Chr(11)
                            rng.Collapse wdCollapseEnd
                            i = i - leadSpaceCount
                        Else
                            'p.Range.Characters(i).InsertAfter Chr(11)
                            rng.InsertAfter Chr(11)
                            rng.Collapse wdCollapseEnd
                        End If
                    Else '沒有凸排
                        'p.Range.Characters(i).InsertAfter Chr(11)
                        rng.InsertAfter Chr(11)
                        rng.Collapse wdCollapseEnd
                    End If
                    
                    i = i + 1
'
                    'i = i + lineLength
                Loop
            End If
        End If
    Next p
    
    Set rng = d.Range
    With rng.Find
        .ClearFormatting
        .font.Size = 7.5
    End With
    
    Do While rng.Find.Execute()
        rng.InsertAfter "}}"
        rng.InsertBefore "{{"
        rng.Collapse wdCollapseEnd
    Loop
    
    d.Range.Find.ClearFormatting
    
    leadSpaces = VBA.StrConv(VBA.space(leadSpaceCount), vbWide)
    d.Range.Find.Execute "^l", , , , , , , , , "^p" & leadSpaces, wdReplaceAll
    d.Range.Cut
    d.Close wdDoNotSaveChanges
    AppActivate "TextForCtext"
    DoEvents
    SendKeys "^v"
    DoEvents
End Sub

Sub 集杜詩_文山先生全集_四部叢刊_維基文庫本_去掉中間誤空的格() '《集杜詩》格式者皆適用（中間誤空的格） 20221112
    Dim rng As Range, d As Document, p As Paragraph, a As Range, i As Integer, ur As UndoRecord
    Set d = ActiveDocument
    If d.path <> "" Then Set d = Documents.Add
    SystemSetup.stopUndo ur, "集杜詩_文山先生全集_四部叢刊_維基文庫本_去掉中間誤空的格"
    For Each p In d.Paragraphs
        For Each a In p.Range.Characters
            i = i + 1
            If i > 3 Then '此與標題、縮排等前空幾格之條件有關
                If Not a.Next Is Nothing And Not a.Previous Is Nothing Then
                    If a <> "　" And a.Next = "　" And a.Previous = "　" Then '單字前後皆空格者才處理
                        Set rng = d.Range(a.End, a.End)
                        rng.MoveEndWhile "　"
        '                rng.Select
        '                Stop
                        rng.Delete
                    End If
                End If
            End If
        Next
        i = 0
    Next p
    DoEvents
    d.Range.Copy
    DoEvents
    SystemSetup.contiUndo ur
    SystemSetup.playSound 2
End Sub

Rem 在新文件上操作：第一段為始頁碼、第二段為終頁碼、第三段為書ID
Sub 新頁面()
    'the page begin
    Dim start As Integer, ur As UndoRecord
    ' the page end
    Dim e As Integer
    ' the book
    Dim fileID As Long
    'https://ctext.org/library.pl?if=gb&file=1000081&page=2621
    
    Dim x As String ', data As New MSForms.DataObject
    Dim i As Integer, rng As Range, d As Document
    SystemSetup.stopUndo ur, "新頁面"
    Set d = ActiveDocument
    If d.path <> "" Then Exit Sub
    Set rng = d.Range
    start = CInt(Replace(rng.Paragraphs(1).Range, VBA.Chr(13), ""))
    e = CInt(Replace(rng.Paragraphs(2).Range, VBA.Chr(13), ""))
    fileID = CLng(Replace(rng.Paragraphs(3).Range, VBA.Chr(13), ""))
    For i = start To e
        If i = 1 Then
            x = x & "<scanbegin file=""" & fileID & """ page=""" & i & """ />●" & VBA.Chr(9) & "<scanend file=""" & fileID & """ page=""" & i & """ />"
        Else
            x = x & "<scanbegin file=""" & fileID & """ page=""" & i & """ />" & VBA.Chr(9) & "<scanend file=""" & fileID & """ page=""" & i & """ />" '若中間沒有任何內容，頁面最後便不能成一段落。若剛好一個段落，會與下一頁黏合在一起
        End If
    Next i
    
    rng.Document.Range(d.Paragraphs(3).Range.start, d.Paragraphs(3).Range.End - 1).text = CLng(Replace(rng.Paragraphs(3).Range, VBA.Chr(13), "")) + 1
    'rng.Document.Paragraphs(3).Range.text = VBA.CStr(VBA.CLng(VBA.Left(rng.Document.Paragraphs(3).Range.text, VBA.Len(rng.Document.Paragraphs(3).Range.text) - 1)) + 1)
    
    'For Each e In Selection.Value
    '    x = x & e
    'Next e
    ''x = Replace(x, vba.Chr(13), "")
    'data.SetText Replace(x, "/>", "/>●", 1, 1)
    'data.PutInClipboard
    'SystemSetup.SetClipboard x
    'SystemSetup.CopyText x
    SystemSetup.SetClipboard x
    If SystemSetup.GetClipboardText <> x Then
        rng.SetRange d.Range.End - 1, d.Range.End - 1
        rng.InsertAfter x
        rng.Cut
    End If
    rng.Document.ActiveWindow.windowState = wdWindowStateMinimize
    DoEvents
    'Network.AppActivateDefaultBrowser
    ActivateChrome
'    SendKeys "^a"
'    SendKeys "^v"
    
    SystemSetup.contiUndo ur
End Sub
Sub setPage1Code() '(ByRef d As Document)
    Dim xd As String
    xd = SystemSetup.GetClipboardText
    If InStr(xd, "page=""1""") = 0 Then
        Dim bID As String, s As Byte, pge As String
        s = InStr(xd, "page=""")
        pge = VBA.Mid(xd, s + Len("page="""), InStr(s + Len("page="""), xd, """") - s - Len("page="""))
        If CInt(pge) < 10 Then
            s = InStr(xd, """")
            bID = VBA.Mid(xd, s + 1, InStr(s + 1, xd, """") - s - 1)
            xd = "<scanbegin file=""" & bID & """ page=""1"" />●<scanend file=""" & bID & """ page=""1"" />" + xd
            SystemSetup.ClipboardPutIn xd
        End If
    End If
End Sub

Sub clearRedundantCode()
Dim xd As String, s As Long, e As Long
xd = SystemSetup.GetClipboardText
s = InStr(xd, "<scanend ") 'end 和 begin間不當有任何文字
Do Until s = 0
    e = InStr(s, xd, ">")
    s = InStr(e, xd, "<scanbegin ")
    If s - e > 1 Then
        xd = VBA.Mid(xd, 1, e) + VBA.Mid(xd, s)
    End If
    s = InStr(e, xd, "<scanend ")
Loop
SystemSetup.ClipboardPutIn xd
End Sub
Sub clearRedundantText()
'清除誤判的注文
clearWrongNoteText
End Sub
Sub clearWrongNoteText() '有些有評點或斷句的版本，OCR時分行切字，乃至誤將其右傍圈點者判讀為字。此則清除之。
Dim d As Document, p As Paragraph
Set d = Documents.Add(, , , False)
d.Range.Paste
For Each p In d.Range.Paragraphs
    If InStr(p.Range, "{{") Or InStr(p.Range, "}}") Then
        p.Range.Delete
    End If
Next p
DoEvents
d.Range.Copy
d.Close wdDoNotSaveChanges
AppActivateChrome
'SendKeys "+{insert}{tab}~"
SendKeys "+{insert}"

End Sub


Sub formatTitleCode() '標題格式設定
Dim rng As Range, d As Document, y As Byte, s As Long, ur As UndoRecord
Set d = ActiveDocument: SystemSetup.stopUndo ur, "formatTitleCode標題格式設定"
Set rng = d.Range
rng.Find.ClearFormatting
For y = 2 To 4
    Do While rng.Find.Execute("y=""" & y & """ />", , , , , , True, wdFindStop)
        GoSub code
    Loop
    Set rng = d.Range
Next y
SystemSetup.contiUndo ur
SystemSetup.playSound 1.469
Exit Sub
code:
    rng.text = rng.text + "*"
    s = rng.End + 1
    rng.Collapse wdCollapseStart
    rng.SetRange rng.start, rng.start
    'rng.MoveStartUntil ">"
    Do Until rng.Next.text = "<"
        rng.Move wdCharacter, -1
    Loop
    rng.Move
    rng.text = rng.text + VBA.Chr(13) + VBA.Chr(13)
    rng.SetRange s, d.Range.End
    Return
End Sub

Sub 清除頁前的分段符號()
    Dim d As Document, rng As Range, e As Long, s As Long, xd As String
    Dim iwe As SeleniumBasic.IWebElement
    Set d = Documents.Add
    DoEvents
    'If (MsgBox("add page 1 code?", vbExclamation + vbOKCancel) = vbOK) Then setPage1Code
    中國哲學書電子化計劃.setPage1Code:  clearRedundantCode
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
        If rng.text = VBA.Chr(13) & VBA.Chr(13) Then
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
        If rng.text = VBA.Chr(13) & VBA.Chr(13) Then
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
    xd = d.Range.text
    'If d.Characters.Count < 50000 Then ' 147686
    '    d.Range.Cut '原來是Word的 cut 到剪貼簿裡有問題
    'Else
        'SystemSetup.SetClipboard d.Range.Text
        SystemSetup.ClipboardPutIn xd
    'End If
    DoEvents
    playSound 1, 0
    DoEvents
    
    pastetoEditBox description
    d.Close wdDoNotSaveChanges

End Sub

Sub 將星號前的分段符號移置前段之末(ByRef d As Document) '20220522
    Dim rng As Range, e As Long, s As Long, rngP As Range
    'd As Document,Set d = Documents.Add
    Set rng = d.Range
    DoEvents
    On Error GoTo eH
    rng.Paste
    rng.Find.ClearFormatting
    Do While rng.Find.Execute("*")
        e = rng.End
        If rng.start > 0 Then
            If rng.Previous = VBA.Chr(13) Then
                Set rng = rng.Previous
                If rng.Previous = VBA.Chr(13) Then
                    Set rng = rng.Previous
                    If rng.Previous = ">" Then
                        rng.SetRange rng.start, e - 1
                        s = rng.start
                        Set rngP = d.Range(s, s)
                        rng.Delete
                        Do Until rngP.Next = "<"
                            If rngP.start = 0 Then GoTo NextOne
                            rngP.Move wdCharacter, -1
                        Loop
                        '檢查是否正在跨頁處 20230811
                        If d.Range(rngP.start, rngP.start + 11) = "><scanbegin" Then
                            rngP.Move Count:=-1
                            Do Until rngP.Next = "<"
                                If rngP.start = 0 Then GoTo NextOne
                                rngP.Move wdCharacter, -1
                            Loop
                        End If
                        '以上 檢查是否正在跨頁處 20230811
                        rngP.Move
                        rngP.InsertAfter VBA.Chr(13) & VBA.Chr(13)
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
    Exit Sub
eH:
    Select Case Err.number
        Case 4605, 13 '此方法或屬性無法使用，因為[剪貼簿] 是空的或無效的。
            SystemSetup.wait 0.8
            Resume
        Case Else
            MsgBox Err.number + Err.description
     End Select
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

Private Sub pastetoEditBox(Optional Description_from_ClipBoard As String = vbNullString)
    word.Application.windowState = wdWindowStateMinimize
    'MsgBox "ready to paste", vbInformation
    AppActivateDefaultBrowser
    DoEvents
    'SystemSetup.Wait 0.5 '關鍵在這行！否則大容量貼上會失效。20220809'根本還是沒用！實際上是在Word的剪下傳送到剪貼簿的資料是空的
    SendKeys "+{INSERT}" '"(^v)" ', True'恐怕要去掉這個才是；都不是！實際上問題是出在Word的剪下傳送到剪貼簿的資料是空的
    DoEvents ' DoEvents: DoEvents
    Beep
    SystemSetup.wait 0.3
    DoEvents:
    SendKeys "{tab}"
    AppActivateDefaultBrowser
    If Description_from_ClipBoard <> vbNullString Then
        SystemSetup.ClipboardPutIn Description_from_ClipBoard
        DoEvents
        SendKeys "^v"
'        SendKeys Description_from_ClipBoard
    End If
    SendKeys "{tab 2}~" '按下 Submit changes
End Sub

Sub 金石錄_四部叢刊_維基文庫本() '《金石錄》格式者皆適用（即注文單行，而換行前的不單行） 20221110
    Dim rng As Range, d As Document, s As Long, e As Long, rngDel As Range, ur As UndoRecord
    Set d = ActiveDocument
    If d.path <> "" Then Set d = Documents.Add
    DoEvents
    d.Range.Paste
    DoEvents
    Set rng = d.Range: Set rngDel = rng
    rng.Find.ClearFormatting
    SystemSetup.stopUndo ur, "金石錄_四部叢刊_維基文庫本"
    Do While rng.Find.Execute("}}|" & VBA.Chr(13) & "{{", , , , , , True, wdFindStop)
        s = rng.start - 1: e = rng.start
        Do Until d.Range(s, e) <> "　" '清除其前空格
            s = s - 1: e = e - 1
        Loop
        rngDel.SetRange s + 1, rng.start
        'rngDel.Select
        If rngDel.text <> "" Then If Replace(rngDel, "　", "") = "" Then rngDel.Delete
        rng.SetRange s + Len("}}|" & VBA.Chr(13) & "{{"), d.Range.End
        
        'Set rng = d.Range
    Loop
    d.Range.text = Replace(Replace(d.Range.text, "|" & VBA.Chr(13) & "　", ""), "}}|" & VBA.Chr(13) & "{{", VBA.Chr(13))
    d.Range.Copy
    SystemSetup.contiUndo ur
    SystemSetup.playSound 2
    word.Application.windowState = wdWindowStateMinimize
    On Error Resume Next
    AppActivate "TextForCtext", True
End Sub

Sub 轉成黑豆以作行字數長度判斷用()
    Dim p As Paragraph, a, i As Byte, cntr As Byte, ur As UndoRecord
    If ActiveDocument.path <> "" Then Exit Sub
    SystemSetup.stopUndo ur, "轉成黑豆以作行字數長度判斷用"
    Set p = Selection.Paragraphs(1)
    cntr = p.Range.Characters.Count - 1
    For i = 1 To cntr
        Set a = p.Range.Characters(i)
        If a.text <> VBA.Chr(13) Then a.text = "●"
    Next i
    p.Range.Cut
    SystemSetup.contiUndo ur
    Set ur = Nothing
End Sub
Sub 清除所有符號_分段注文符號例外()
    Dim f, i As Integer
    f = Array("。", "」", VBA.Chr(-24152), "：", "，", "；", _
        "、", "「", ".", VBA.Chr(34), ":", ",", ";", _
        "……", "...", "．", "【", "】", " ", "《", "》", "〈", "〉", "？" _
        , "！", "﹝", "﹞", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0" _
        , "『", "』", VBA.ChrW(9312), VBA.ChrW(9313), VBA.ChrW(9314), VBA.ChrW(9315), VBA.ChrW(9316) _
        , VBA.ChrW(9317), VBA.ChrW(9318), VBA.ChrW(9319), VBA.ChrW(9320), VBA.ChrW(9321), VBA.ChrW(9322), VBA.ChrW(9323) _
        , VBA.ChrW(9324), VBA.ChrW(9325), VBA.ChrW(9326), VBA.ChrW(9327), VBA.ChrW(9328), VBA.ChrW(9329), VBA.ChrW(9330) _
        , VBA.ChrW(9331), VBA.ChrW(8221), """") '先設定標點符號陣列以備用
        '全形圓括弧暫不取代！
        For i = 0 To UBound(f)
            ActiveDocument.Range.Find.Execute f(i), True, , , , , , wdFindContinue, True, "", wdReplaceAll
        Next
End Sub

Sub 撤掉與書圖的對應_脫鉤() '20220210
    Dim rng As Range, angleRng As Range, cntr As Long
    word.Application.ScreenUpdating = False
    Set rng = Documents.Add().Range
    Set angleRng = rng
    rng.Paste
    Do While rng.Find.Execute("<")
        rng.MoveEndUntil ">"
        rng.SetRange rng.start, rng.End + 1
        angleRng.SetRange rng.start, rng.End
        If InStr(angleRng.text, "file") > 0 Then
            angleRng.Delete
        Else
            rng.SetRange rng.End, rng.Document.Range.End - 1
        End If
        If InStr(rng.Document.Range, " file=") = 0 Then Exit Do '若有上標籤「<entity entityid=」，則判斷會失誤
        cntr = cntr + 1
        If cntr > 2300 Then Stop
    Loop
    SystemSetup.playSound 1
    rng.Document.Range.Cut
    rng.Document.Close wdDoNotSaveChanges
    word.Application.ScreenUpdating = True
    pastetoEditBox "與原本書圖不合，圖文脫鉤。另依《維基文庫》本輔以末學自製軟件TextForCtext對應錄入。感恩感恩　南無阿彌陀佛"
End Sub

Sub formatter() '為《經典釋文》春秋三傳等格式用，日後可改成其他需要格式化的文本
    Dim d As Document, rng As Range, a As Range, s As Long, e As Long
    Const spcs As String = "　"
    
    Set d = Documents.Add: Set rng = d.Range
    rng.Paste
    For Each a In d.Characters
        If a = spcs Then
            If a.Next = spcs Then
                If InStr(a.Paragraphs(1).Range.text, "*") = 0 Then
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
                        rng.text = Replace(rng.text, "　", VBA.ChrW(-9217) & VBA.ChrW(-8195))
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
            If VBA.Right(a, 1) = y Then
                If a.Previous.Previous <> ">" Then
                    For yi = 1 To 99
                        yStr = 文字轉換.數字轉漢字2位數(yi) + y
                        If a.text = yStr Then
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
    On Error GoTo eH
    a = Array(VBA.ChrW(12296), "{{", VBA.ChrW(12297), "}}", "〈", "{{", "〉", "}}", _
        "○", VBA.ChrW(12295))
    '《容齋三筆》等小注作正文省版面者 https://ctext.org/library.pl?if=gb&file=89545&page=24
    'a = Array("〈", "", "〉", "", _
        "○", vba.Chrw(12295))
    
    
    Set d = Documents.Add()
    d.Range.Paste
    '提示貼上無礙
    SystemSetup.playSound 1
    維基文庫造字圖取代為文字 d.Range
    For i = 0 To UBound(a) - 1
        d.Range.Find.Execute a(i), , , , , , True, wdFindContinue, , a(i + 1), wdReplaceAll
        i = i + 1
    Next i
    For Each p In d.Range.Paragraphs
        xP = p.Range
        If VBA.Left(xP, 2) = "{{" And VBA.Right(xP, 3) = "}}" & VBA.Chr(13) Then
            xP = VBA.Mid(p.Range, 3, Len(xP) - 5)
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
        ElseIf VBA.Left(xP, 1) = "　" Then '前有空格的
            i = InStr(xP, "{{")
            If i > 0 And VBA.Right(xP, 3) = "}}" & VBA.Chr(13) Then
                space = VBA.Mid(xP, 1, i - 1)
                If Replace(space, "　", "") = "" Then
                    xP = VBA.Mid(xP, i + 2, Len(xP) - 3 - (i + 2))
                    If InStr(xP, "{{") = 0 And InStr(xP, "}}") = 0 Then
                        Set rng = p.Range
                        rng.SetRange rng.Characters(1).start, rng.Characters(i + 1).End
                        rng.text = "{{" & space
                        acP = p.Range.Characters.Count - 1 - Len(space)
                        If acP Mod 2 = 0 Then
                            acP = CInt(acP / 2) + Len(space) + 1
                        Else
                            acP = CInt((acP + 1) / 2) + Len(space) + 1
                        End If
                        If p.Range.Characters(acP).InlineShapes.Count = 0 Then
                            p.Range.Characters(acP).InsertBefore VBA.Chr(13) & space
                        Else
                            p.Range.Characters(acP).Select
                            Selection.Delete
                            Selection.TypeText " "
                            p.Range.Characters(acP).InsertBefore VBA.Chr(13) & space
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
eH:
    Select Case Err.number
        Case 5904 '無法編輯 [範圍]。
            If p.Range.Characters(acP).Hyperlinks.Count > 0 Then p.Range.Characters(acP).Hyperlinks(1).Delete
            Resume
        Case Else
            MsgBox Err.number & Err.description
    End Select
End Sub

Sub 維基文庫四部叢刊本轉來_early()
    Dim d As Document, a, i
    
    a = Array("^p^p", "@", "〈", "{{", "〉", "}}", "^p", "", "}}{{", "^p", "@", "^p", _
        "○", VBA.ChrW(12295))
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

Sub searchuCtext()
    ' Alt + Shift + ,
    ' Alt + <
    SystemSetup.playSound 0.484
    Select Case Selection.text
        Case "", VBA.Chr(13), VBA.Chr(9), VBA.Chr(7), VBA.Chr(10), " ", "　"
            MsgBox "no selected text for search !", vbCritical: Exit Sub
    End Select
    Static bookID
    Dim searchedTerm, e, addressHyper As String, bID As String, cndn As String
    'Const site As String = "https://ctext.org/wiki.pl?if=gb&res="
    Const site As String = "https://ctext.org/wiki.pl?if=gb"
    bID = VBA.Left(ActiveDocument.Paragraphs(1).Range, Len(ActiveDocument.Paragraphs(1).Range) - 1)
    If Not VBA.IsNumeric(bID) Then
        If InStr(bID, site) = 0 Then
            bookID = InputBox("plz input the book id ", , bookID)
        Else
            bookID = bID
        End If
    Else
        bookID = bID
    End If
    If InStr(bookID, "https") > 0 Then
        If InStr(bookID, "&res=") = 0 And InStr(bookID, "&chapter=") = 0 Then MsgBox "error . not the proper bookID ref ", vbCritical: Exit Sub
        If InStr(bookID, "&res=") > 0 Then
            cndn = "&res="
        ElseIf InStr(bookID, "&chapter=") > 0 Then
            cndn = "&chapter="
        Else
            MsgBox "error . not the proper bookID ref ", vbCritical: Exit Sub
        End If
        bookID = VBA.Mid(bookID, InStr(bookID, cndn) + Len(cndn))
        If Not VBA.IsNumeric(bookID) Then
            bookID = VBA.Mid(bookID, 0, InStr(bookID, "&searchu"))
        End If
    End If
    If Not VBA.IsNumeric(bookID) Then
        MsgBox "error . not the proper bookID ref ", vbCritical: Exit Sub
    End If
    文字處理.ResetSelectionAvoidSymbols
    e = code.UrlEncode(Selection.text)
    'searchedTerm = 'Array("卦", "爻", "周易", "易經", "系辭", "繫辭", "擊辭", "說卦", "序卦", "卦序", "敘卦", "雜卦", "文言", "乾坤", "無咎", vba.Chrw(26080) & "咎", "天咎", "元亨", "利貞", "易") ', "", "", "", "")
    ''https://ctext.org/wiki.pl?if=gb&res=757381&searchu=%E5%8D%A6
    'For Each e In searchedTerm
        addressHyper = addressHyper + " " + site + cndn + bookID + "&searchu=" + e
    'Next e
    Shell Network.getDefaultBrowserFullname + addressHyper + " --remote-debugging-port=9222 "
    
    Selection.Hyperlinks.Add Selection.Range, addressHyper
End Sub

Sub 史記三家注()
'從2858頁起，20210920:0817之後，改用臺師大附中同學吳恆昇先生《中華文化網》所錄中研院《瀚典》初本，雖或仍未精，然至少免有簡化字轉換訛窘或造字亂碼之困擾，原文字檔棄置。根據初作比對，格式完全一樣！根本就是從這裡出來的，再轉簡化字，再又反正，造成之紊亂。悔當初沒想到用此本也。阿彌陀佛。佛弟子孫守真任真甫謹識於2021年9月20日
Dim d As Document, a, i, p As Paragraph, px As String, rng As Range, e As Long, pRng As Range, pa
'Const corTxt As String = "＝詳點校本校勘記＝"'該網站圖文對照排版功能未能配合，故今不採用。其格式只對文本版有效。https://ctext.org/instructions/wiki-formatting/zh
'a = Array(" ", "", "　　","","　", vba.Chrw(-9217) & vba.Chrw(-8195), "^p", "<p>^p",
'a = Array(" ", "", "　　", "", "^p^p", "<p>^p" & vba.Chrw(-9217) & vba.Chrw(-8195) & vba.Chrw(-9217) & vba.Chrw(-8195),
a = Array("　　", "", "^p", "^p^p", "^p^p^p", "^p^p", "「^p^p", "「", "『^p^p", "『", "〔^p^p", "〔", "（^p^p", "（", _
    "^p^p", "<p>^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195), _
    "^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195) & "〔", _
    "^p{{" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & "{{{〈", _
    "「<p>^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195), "「", _
    "〔<p>^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195), "〔", _
    "『<p>^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195), "『", _
    "（<p>^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195), "（", _
    "集解", "《集解》：", "索隱", "《索隱》：", "【《索隱》：述贊】", "【《索隱》述贊】：", "正義", "《正義》：", _
    "九州島", "九州", "齊愍", "齊湣", "愍王", "湣王", "安厘王", "安釐王", _
    "塚", "冢", "慚", VBA.ChrW(24921), "啟", VBA.ChrW(21843), _
     VBA.ChrW(-30641), VBA.ChrW(-25066), _
     "群", VBA.ChrW(32675), "即", VBA.ChrW(21373), "眾", VBA.ChrW(-30650), _
     "既", VBA.ChrW(26083), "概", VBA.ChrW(27114), "溉", VBA.ChrW(28433), _
     "衛", VBA.ChrW(-30626), _
     "真", VBA.ChrW(30494), "填", VBA.ChrW(22625), "清", VBA.ChrW(28152), "青", VBA.ChrW(-26799), "教", VBA.ChrW(25934), _
    "鄉", VBA.ChrW(-28395), "鎮", VBA.ChrW(-27731), "慎", VBA.ChrW(24892), _
    "并", VBA.ChrW(24183), "屏", VBA.ChrW(23643), "荊", VBA.ChrW(-31930), "邢", VBA.ChrW(-28471), "笄", VBA.ChrW(31571), _
    "犁", VBA.ChrW(29314), "鬥", VBA.ChrW(-25811), "綿", VBA.ChrW(32220), _
    "冉", VBA.ChrW(20868), "腳", VBA.ChrW(-32486), _
    VBA.ChrW(25995), VBA.ChrW(-24956))
Set d = Documents.Add()
d.Range.Paste
維基文庫造字圖取代為文字 d.Range
d.Range.Cut
d.Range.PasteAndFormat wdFormatPlainText
d.Range.text = VBA.Replace(d.Range.text, " ", "")
For i = 0 To UBound(a) - 1
    If a(i) = "^p^p^p" Then
        px = d.Range.text
        Do While InStr(px, VBA.Chr(13) & VBA.Chr(13) & VBA.Chr(13))
            px = Replace(px, VBA.Chr(13) & VBA.Chr(13) & VBA.Chr(13), VBA.Chr(13) & VBA.Chr(13))
        Loop
        d.Range.text = px
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
    px = p.Range.text
    If VBA.Left(px, 7) = "{{" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & "{{{" Then '注腳段落
        e = p.Range.Characters(1).End
        rng.SetRange e, e
        rng.MoveEndUntil "〕"
        If rng.Next.Next = "　" Then rng.Next.Next.Delete
        If InStr(p.Range.text, "　") Then
            For Each pa In p.Range.Characters
                If pa = "　" Then
                    pa.text = VBA.ChrW(-9217) & VBA.ChrW(-8195)
                End If
            Next
'            p.Range.text = VBA.Replace(p.Range.text, "　", vba.Chrw(-9217) & vba.Chrw(-8195))
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
        px = p.Range.text
        If InStr(VBA.Right(px, 4), "<p>") Then
            e = p.Range.Characters(p.Range.Characters.Count - 4).End
        Else
            e = p.Range.Characters(p.Range.Characters.Count - 1).End
        End If
        rng.SetRange e, e
        rng.InsertAfter "}}"
    Else '正文段落
        e = p.Range.Characters(1).start
        Set pRng = p.Range
        Do While InStr(pRng.text, "〔")
            rng.SetRange e, e
            rng.MoveEndUntil "〔"
            If rng.Characters(rng.Characters.Count) <> "）" Then  ' if not correction
                rng.Collapse wdCollapseEnd
                rng.Move , 1
                rng.MoveEnd wdCharacter, 1
                If rng.text Like "[一二三四五六七八九]" Then  ' is footnote No.
                    e = rng.start
                    'rng.Collapse wdCollapseEnd
                    rng.SetRange e - 1, e
                    rng.text = "　{{{〈"
                    rng.MoveEndUntil "〕"
                    rng.Collapse wdCollapseEnd
                    rng.MoveEnd wdCharacter, 1
                    rng.text = "〉}}}"
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
    If VBA.Left(p.Range.text, 9) = VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195) & "【《索隱》" Then
        Set rng = p.Range
        p.Range.Characters(1).Delete
        rng.SetRange p.Range.start, p.Range.start
        rng.InsertAfter "{{"
        rng.SetRange p.Range.Characters(p.Range.Characters.Count - 4).End, p.Range.Characters(p.Range.Characters.Count - 4).End
        rng.InsertAfter "}}"
        If Len(rng.Paragraphs(1).Next.Range.text) = 1 Then rng.Paragraphs(1).Next.Range.Delete
    End If
    
    If Len(p.Range) < 20 Then
        If (InStr(p.Range, "《史記》卷") Or VBA.Left(p.Range.text, 3) = "史記卷") And InStr(p.Range, "*") = 0 Then
            rng.SetRange p.Range.start, p.Range.start
            rng.InsertAfter "*"
            For Each pa In p.Range.Characters
                    If pa Like "[〈《》〉]" Or StrComp(pa, VBA.ChrW(-9217) & VBA.ChrW(-8195)) = 0 Then pa.Delete
            Next pa
            '以下方式會造成p 值被設定為下一個段落
'            p.Range.text = VBA.Replace(p.Range.text, vba.Chrw(-9217) & vba.Chrw(-8195) & vba.Chrw(-9217) & vba.Chrw(-8195), "")
'            p.Range.text = VBA.Replace(VBA.Replace(p.Range.text, "《", ""), "》", "")
        End If
    End If
    If Len(p.Range) < 25 Then
        If VBA.InStr(p.Range.text, "第") And InStr(p.Range, "*") = 0 _
                And (InStr(p.Range, "本紀") Or InStr(p.Range, "書") Or InStr(p.Range, "表") _
                Or InStr(p.Range, "世家") Or InStr(p.Range, "列傳")) Then
            rng.SetRange p.Range.start, p.Range.start
            rng.InsertAfter "　*"
            For Each pa In p.Range.Characters
                If pa Like "[〈《》〉]" Or StrComp(pa, VBA.ChrW(-9217) & VBA.ChrW(-8195)) = 0 Then pa.Delete
            Next pa
   
'            p.Range.text = VBA.Replace(p.Range.text, vba.Chrw(-9217) & vba.Chrw(-8195) & vba.Chrw(-9217) & vba.Chrw(-8195), "　*")
'            p.Range.text = VBA.Replace(VBA.Replace(p.Range.text, "〈", ""), "〉", "")
        End If
    End If

Next p
If VBA.Left(d.Paragraphs(1).Range.text, 3) = "史記卷" And InStr(d.Paragraphs(1).Range.text, "*") = 0 Then
    Set p = d.Paragraphs(1)
    rng.SetRange p.Range.start, p.Range.start
    rng.InsertAfter "*"
'    rng.SetRange p.Range.Characters(p.Range.Characters.Count - 1).End, p.Range.Characters(p.Range.Characters.Count - 1).End
'    rng.InsertAfter "<p>"
End If
If VBA.InStr(d.Paragraphs(2).Range.text, "第") And InStr(d.Paragraphs(2).Range.text, "*") = 0 Then
    Set p = d.Paragraphs(2)
'    rng.SetRange p.Range.start, p.Range.start
'    rng.InsertAfter "　*"
''    rng.SetRange p.Range.Characters(p.Range.Characters.Count - 1).End, p.Range.Characters(p.Range.Characters.Count - 1).End
''    rng.InsertAfter "<p>"
    p.Range.text = VBA.Replace(p.Range.text, VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195), "　*")
    Set p = d.Paragraphs(2)
    p.Range.text = VBA.Replace(VBA.Replace(p.Range.text, "〈", ""), "〉", "")
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
word.Application.ActiveWindow.windowState = wdWindowStateMinimize
End Sub
Sub 史記三家注2old()
'從2858頁起，20210920:0817之後，改用臺師大附中同學吳恆昇先生《中華文化網》所錄中研院《瀚典》初本，雖或仍未精，然至少免有簡化字轉換訛窘或造字亂碼之困擾，原文字檔棄置。根據初作比對，格式完全一樣！根本就是從這裡出來的，再轉簡化字，再又反正，造成之紊亂。悔當初沒想到用此本也。阿彌陀佛。佛弟子孫守真任真甫謹識於2021年9月20日
Dim d As Document, a, i, p As Paragraph, px As String, rng As Range, e As Long, pRng As Range
'Const corTxt As String = "＝詳點校本校勘記＝"'該網站圖文對照排版功能未能配合，故今不採用。其格式只對文本版有效。https://ctext.org/instructions/wiki-formatting/zh
a = Array("^p^p", "<p>^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195), _
    "^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195) & "〔", _
    "^p{{" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & "{{{〈", _
    "「<p>^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195), "「", _
    "〔<p>^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195), "〔", _
    "『<p>^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195), "『", _
    "（<p>^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195), "（", _
    "集解", "《集解》：", "索隱", "《索隱》：", "【《索隱》：述贊】", "【《索隱》述贊】：", "正義", "《正義》：", _
    "九州島", "九州", "齊愍", "齊湣", "愍王", "湣王", "安厘王", "安釐王", _
    "塚", "冢", _
     "群", VBA.ChrW(32675), "即", VBA.ChrW(21373), "眾", VBA.ChrW(-30650), "既", VBA.ChrW(26083), "衛", VBA.ChrW(-30626), _
     "真", VBA.ChrW(30494), "填", VBA.ChrW(22625), "清", VBA.ChrW(28152), "青", VBA.ChrW(-26799), "教", VBA.ChrW(25934), _
    "鄉", VBA.ChrW(-28395), "鎮", VBA.ChrW(-27731), "慎", VBA.ChrW(24892), "屏", VBA.ChrW(23643), "概", VBA.ChrW(27114), _
    "荊", VBA.ChrW(-31930), "邢", VBA.ChrW(-28471))
Set d = Documents.Add()
d.Range.Paste
For i = 0 To UBound(a) - 1
    d.Range.Find.Execute a(i), , , , , , True, wdFindContinue, , a(i + 1), wdReplaceAll
    i = i + 1
Next i
文字處理.書名號篇名號標注
Set rng = Selection.Range
For Each p In d.Paragraphs
    px = p.Range.text
    If VBA.Left(px, 7) = "{{" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & "{{{" Then '注腳段落
        e = p.Range.Characters(1).End
        rng.SetRange e, e
        rng.MoveEndUntil "〕"
        'rng.Select
        rng.Collapse wdCollapseEnd
        rng.Select
        Selection.MoveRight wdCharacter, 1, wdExtend
        Selection.TypeText "〉}}}" '將注腳編號〔一〕的右邊〕改成}}}
        px = p.Range.text
        If InStr(VBA.Right(px, 4), "<p>") Then
            e = p.Range.Characters(p.Range.Characters.Count - 4).End
        Else
            e = p.Range.Characters(p.Range.Characters.Count - 1).End
        End If
        rng.SetRange e, e
        rng.InsertAfter "}}"
    Else '正文段落
        e = p.Range.Characters(1).start
        Set pRng = p.Range
        Do While InStr(pRng.text, "〔")
            rng.SetRange e, e
            rng.MoveEndUntil "〔"
            If rng.Characters(rng.Characters.Count) <> "）" Then  ' if not correction
                rng.Collapse wdCollapseEnd
                rng.Move , 1
                rng.MoveEnd wdCharacter, 1
                If rng.text Like "[一二三四五六七八九]" Then  ' is footnote No.
                    e = rng.start
                    'rng.Collapse wdCollapseEnd
                    rng.SetRange e - 1, e
                    rng.text = "　{{{〈"
                    rng.MoveEndUntil "〕"
                    rng.Collapse wdCollapseEnd
                    rng.MoveEnd wdCharacter, 1
                    rng.text = "〉}}}"
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
    If VBA.Left(p.Range.text, 9) = VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195) & "【《索隱》" Then
        Set rng = p.Range
        p.Range.Characters(1).Delete
        rng.SetRange p.Range.start, p.Range.start
        rng.InsertAfter "{{"
        rng.SetRange p.Range.Characters(p.Range.Characters.Count).End, p.Range.Characters(p.Range.Characters.Count).End
        rng.InsertAfter "}}"
    End If
Next p
If VBA.Left(d.Paragraphs(1).Range.text, 3) = "史記卷" Then
    Set p = d.Paragraphs(1)
    rng.SetRange p.Range.start, p.Range.start
    rng.InsertAfter "*"
    rng.SetRange p.Range.Characters(p.Range.Characters.Count - 1).End, p.Range.Characters(p.Range.Characters.Count - 1).End
    rng.InsertAfter "<p>"
End If
If VBA.InStr(d.Paragraphs(2).Range.text, "第") Then
    Set p = d.Paragraphs(2)
    rng.SetRange p.Range.start, p.Range.start
    rng.InsertAfter "　*"
'    rng.SetRange p.Range.Characters(p.Range.Characters.Count - 1).End, p.Range.Characters(p.Range.Characters.Count - 1).End
'    rng.InsertAfter "<p>"
    p.Range.text = VBA.Replace(VBA.Replace(p.Range.text, "〈", ""), "〉", "")
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
a = Array("<p>{{{", "<p>^p{{" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & "{{{", _
        "<p>", "<p>^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195), _
        VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195) & "^p{{" & VBA.ChrW(-9217) & VBA.ChrW(-8195), _
        "{{" & VBA.ChrW(-9217) & VBA.ChrW(-8195))
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
    px = p.Range.text
    If VBA.Left(p.Range.text, 7) = "{{" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & "{{{" Then
        If InStr(VBA.Right(px, 4), "<p>") Then
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
    .font.Color = 10092543
    .font.Size = 10
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
        rngLast.InsertBefore "{{" & VBA.ChrW(-9217) & VBA.ChrW(-8195)
'        rng.SetRange rng.End + 222, d.Range.End
        
    Loop 'Until InStr(rng, "{{")
    .ClearFormatting
End With
Beep
End Sub

Rem 回傳網址
Function Search(searchWhatsUrl As String) As String
    Dim d As Document, encode As String
    Set d = ActiveDocument
    If d.path <> "" Then If d.Saved = False Then d.Save
    文字處理.ResetSelectionAvoidSymbols
    If Selection.Type = wdSelectionNormal Then
        Selection.Copy
    End If
    encode = code.UrlEncode(Selection.text)
    'Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe https://ctext.org/wiki.pl?if=gb&res=384378&searchu=" & Selection.text
    'Shell Normal.SystemSetup.getChrome & searchWhatsUrl & Selection.Text
    Shell TextForCtextWordVBA.Network.GetDefaultBrowserEXE & searchWhatsUrl & encode
    Search = searchWhatsUrl & encode
End Function
Rem 檢索CTP特定之書 成功則傳回true
Function Searchu(res As String, undoName As String) As Boolean
    Dim url As String, ur As UndoRecord, d As Document
    SystemSetup.stopUndo ur, undoName
    'SystemSetup.playSound 0.484
    Set d = Selection.Document
    If d.path <> "" Then If d.Saved = False Then d.Save
    
    文字處理.ResetSelectionAvoidSymbols
    If Selection.Type = wdSelectionNormal Then
        Selection.Copy
    End If
    
    Dim iwe As SeleniumBasic.IWebElement, key As New SeleniumBasic.keys
    If Not SeleniumOP.OpenChrome("https://ctext.org/wiki.pl?if=gb&res=" & res) Then Exit Function
    SeleniumOP.ActivateChrome
    word.Application.windowState = wdWindowStateMinimize
    '檢索框
    Set iwe = SeleniumOP.WD.FindElementByCssSelector("#content > div.wikibox > table > tbody > tr.mobilesearch > td > form > input[type=text]:nth-child(3)")
    If iwe Is Nothing Then Exit Function
    SeleniumOP.SetIWebElementValueProperty iwe, Selection.text
    
    On Error GoTo eH
    iwe.SendKeys key.enter
    '檢索結果
    Set iwe = SeleniumOP.WD.FindElementByCssSelector("#content > table.searchsummary > tbody > tr:nth-child(4) > th > b")
    If iwe Is Nothing Then Exit Function
    If iwe.GetAttribute("textContent") <> "Total 0" Then url = SeleniumOP.WD.url
    If url <> vbNullString Then
        If Selection.Type = wdSelectionIP Then Selection.MoveRight wdCharacter, 1, wdExtend
        ActiveDocument.Hyperlinks.Add Selection.Range, url
    End If
    SystemSetup.contiUndo ur
    Searchu = True
    Exit Function
eH:
    Select Case Err.number
        Case -2146233088
            If VBA.InStr(Err.description, "element not interactable") = 1 Then '(Session info: chrome=130.0.6723.117)
                Set iwe = SeleniumOP.WD.FindElementByCssSelector("#searchform > input.searchbox")
                SeleniumOP.SetIWebElementValueProperty iwe, Selection.text
                Resume
            Else
                GoTo elses
            End If
        Case Else
elses:
            Debug.Print Err.number & Err.description
            MsgBox Err.number & Err.description
    End Select
End Function

Rem 20241006 以Google檢索《中國哲學書電子化計劃》 Alt + t
Sub SearchSite()
    SeleniumOP.GoogleSearch "site:https://ctext.org/ """ + Selection.text + """"
End Sub
Rem Alt + m ： 以選取文字 search史記三家注並於於選取處插入檢索結果之超連結 （m=司馬遷的馬 ma） 20241014;20241005
'原為 Ctrl + s,j 因這樣的指定會取消掉內建的 Ctrl + s ，故改定 20241014
Sub search史記三家注()
    Searchu "384378", "search史記三家注"
'    Dim ur As UndoRecord
'    SystemSetup.stopUndo ur, "search史記三家注"
'    ActiveDocument.Hyperlinks.Add Selection.Range, Search(" https://ctext.org/wiki.pl?if=gb&res=384378&searchu=")
'    SystemSetup.contiUndo ur
End Sub
Rem Ctrl + Alt + = ： 以選取的文字檢索 CTP 所收阮元《十三經注疏·周易正義》並在選取文字上加上該檢索結果頁面之超連結
Sub search周易正義_阮元十三經注疏()
    Searchu "315747", "search周易正義_阮元十三經注疏"
    'url = 中國哲學書電子化計劃.Search(" https://ctext.org/wiki.pl?if=gb&res=315747&searchu=")
    
End Sub
Rem Ctrl + shift + y ： 以選取文字 search《四部叢刊》本《周易》並於於選取處插入檢索結果之超連結(y:yi 易) 20241005
Sub search周易_四部叢刊本()
    Searchu "129518", "search周易_四部叢刊本"
    'ActiveDocument.Hyperlinks.Add Selection.Range, Search(" https://ctext.org/wiki.pl?if=gb&res=129518&searchu=")
End Sub
Sub 讀史記三家注()
    Dim d As Document, t As table
    Set d = Documents.Add
    d.Range.Paste
    Set t = d.tables(1)
    With t
        .Columns(1).Delete
        .ConvertToText wdSeparateByParagraphs
    End With
    d.Range.Cut
    d.Close wdDoNotSaveChanges
    If word.Application.Windows.Count > 0 Then word.Application.ActiveWindow.windowState = wdWindowStateMinimize
End Sub

Sub 戰國策_四部叢刊_維基文庫本() '《戰國策》格式者皆適用（即主文首行頂格，而其餘內容降一格者）
'https://ctext.org/library.pl?if=gb&res=77385
    Dim a, rng As Range, rngDoc As Range, p As Paragraph, i As Long, rngCnt As Integer, ok As Boolean
    Dim omits As String
    omits = "《》〈〉「」『』·" & VBA.Chr(13)
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
                If a.Previous <> VBA.Chr(13) Then a.InsertBefore VBA.Chr(13)
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
                   If rng.Characters(i) = VBA.Chr(13) Then
                        i = 0
                        Exit For
                   End If
                Next a
            Else
                For Each a In rng.Characters
                   i = i + 1
                   If rng.Characters(i) = "}" Then Exit For
                   If rng.Characters(i) = VBA.Chr(13) Or rng.Characters(i) = "{" Then
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
            If VBA.Left(p.Range.text, 3) = "{{　" And p.Range.Characters(p.Range.Characters.Count - 1) = "}" Then
                a = p.Range.text
                a = VBA.Mid(a, 4, Len(a) - 6)
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
    '    rngDoc.Find.Execute vba.Chrw(-10155) & vba.Chrw(-8585) & "曰", , , , , , , wdFindContinue, , "【" & vba.Chrw(-10155) & vba.Chrw(-8585) & "曰】", wdReplaceAll
    '    rngDoc.Find.Execute "補曰", , , , , , , wdFindContinue, , "【" & vba.Chrw(-10155) & vba.Chrw(-8585) & "曰】", wdReplaceAll
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
Sub 楚辭集注縮排N格雙行小注格式_四庫全書_國學大師()
    Dim d As Document, p As Paragraph, px As String, rng As Range, a As Range, ur As UndoRecord, s As Long, e As Long, sx As String
    Set d = ActiveDocument: Set rng = d.Range
    SystemSetup.stopUndo ur, "楚辭集注縮排N格雙行小注格式_四庫全書_國學大師"
    For Each p In d.Paragraphs
        px = p.Range.text
        s = VBA.InStr(px, "{{"): e = VBA.InStr(px, "}}" & VBA.Chr(13))
        If e > 0 Then sx = VBA.Mid(px, s + 2, e - s - 2)
        If e > 0 And s > 0 And VBA.InStr(sx, "{{") = 0 And VBA.InStr(sx, "}}") = 0 Then '前後有{{}}，但些中間不能再有{{}}
            If s = 1 Then '如果前無縮排
                rng.SetRange p.Range.start + 2, p.Range.End - 3
                rng.Characters(Int(rng.Characters.Count / 2)).InsertAfter VBA.Chr(13)
            Else
                If VBA.InStr(px, "　{{") > 0 Then
                    sx = VBA.Mid(px, 1, s - 1)
                    If Replace(sx, "　", vbNullString) = vbNullString Then '如前前綴都是全形空格；即縮排
                        rng.SetRange p.Range.start + VBA.Len(sx) + 2, p.Range.End - 3
                        rng.Characters(Int(rng.Characters.Count / 2)).InsertAfter VBA.Chr(13) & VBA.Mid(px, 1, s - 1)
                    End If
                End If
            End If
        End If
        
    Next p
    SystemSetup.contiUndo ur
End Sub

Sub 本草綱目縮排1格雙行小注格式_四庫全書_國學大師()
    Dim d As Document, p As Paragraph, px As String, rng As Range, a As Range, ur As UndoRecord
    Set d = ActiveDocument: Set rng = d.Range
    SystemSetup.stopUndo ur, "本草綱目縮排一格雙行小注格式_四庫全書_國學大師"
    For Each p In d.Paragraphs
        px = p.Range.text
        If (VBA.Left(px, 3) = "　{{" Or VBA.Left(px, 3) = "{{　") And VBA.Right(px, 3) = "}}" & VBA.Chr(13) Then
            rng.SetRange p.Range.start + 3, p.Range.End - 3
            If VBA.InStr(rng.text, "}") = 0 Then
                If rng.Characters.Count > 1 Then
                    rng.Characters(Int(rng.Characters.Count / 2)).InsertAfter VBA.Chr(13) & "　"
                Else
                    rng.Characters(1).InsertAfter VBA.Chr(13) & "　"
                End If
            End If
        ElseIf VBA.Left(px, 3) = "{{　" And VBA.Right(px, 6) = "}}<p>" & VBA.Chr(13) Then
            rng.SetRange p.Range.start + 3, p.Range.End - 6
            If VBA.InStr(rng.text, "}") = 0 Then
                If InStr(rng.text, "　") Then
                    For Each a In rng.Characters
                        If a.text = "　" Then
                            a.InsertBefore VBA.Chr(13)
                        End If
                    Next a
                Else
                    If rng.Characters.Count > 1 Then
                        'Skype Copilot大菩薩 20240519
                        rng.Characters(-Int(-(rng.Characters.Count / 2))).InsertAfter VBA.Chr(13) & "　"
                    Else
                        rng.Characters(1).InsertAfter VBA.Chr(13) & "　"
                    End If
                End If
            End If
        End If
    Next p
    SystemSetup.contiUndo ur
End Sub

Sub 補括弧()
    Dim d As Document, rng As Range, p As Paragraph, ur As UndoRecord
    Set d = ActiveDocument: SystemSetup.stopUndo ur, "補括弧"
    Set rng = d.Range
    For Each p In d.Paragraphs
        If VBA.Left(p.Range.text, 2) = "{{" And VBA.Right(p.Range.text, 3) <> "}}" & VBA.Chr(13) Then
            If VBA.Right(p.Next.Range.text, 3) = "}}" & VBA.Chr(13) Then
                rng.SetRange p.Range.start, p.Range.End - 1
                rng.text = VBA.Left(p.Range.text, VBA.Len(p.Range.text) - 1) & "}}"
                p.Next.Range.text = "{{" & p.Next.Range.text
            End If
        End If
    Next p
    SystemSetup.contiUndo ur
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
        ElseIf aLtTxt Like VBA.ChrW(12272) & VBA.ChrW(-10155) & VBA.ChrW(-8696) & VBA.ChrW(31860) Then
            aLtTxt = "隸"
        ElseIf aLtTxt Like "彎（?弓爪）-- 弧莫不投" Then
            aLtTxt = "弧"
        ElseIf aLtTxt Like "?土? -- 坳" Then
            aLtTxt = "坳"
        ElseIf aLtTxt Like "????口?欠 -- " & VBA.ChrW(-10111) & VBA.ChrW(-8620) Then
            aLtTxt = VBA.ChrW(-10111) & VBA.ChrW(-8620)
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
            aLtTxt = VBA.ChrW(-10114) & VBA.ChrW(-9161)
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
        ElseIf aLtTxt Like VBA.ChrW(24298) & "（" & VBA.ChrW(8220) & VBA.ChrW(13357) & VBA.ChrW(8221) & "換為" & VBA.ChrW(8220) & "面" & VBA.ChrW(8221) & "）" Then
            aLtTxt = "廩"
        ElseIf aLtTxt Like VBA.ChrW(12273) & VBA.ChrW(11966) & VBA.ChrW(30464) Then
            aLtTxt = "萌"
        ElseIf aLtTxt Like "?彳? -- 徊" Then
            aLtTxt = "徊"
        ElseIf aLtTxt Like VBA.ChrW(12272) & VBA.ChrW(-10145) & VBA.ChrW(-8265) & "變" Then
            aLtTxt = "●＝" & aLtTxt & "＝"
        ElseIf aLtTxt Like "? -- or ?? ?" Then
            aLtTxt = VBA.ChrW(-32119)
        ElseIf aLtTxt Like "輕" Then
            aLtTxt = VBA.ChrW(18518)
        ElseIf aLtTxt Like "能" Then
            aLtTxt = VBA.ChrW(17403)
        ElseIf aLtTxt Like VBA.ChrW(12272) & VBA.ChrW(-10145) & VBA.ChrW(-8265) & VBA.ChrW(25908) Then
            aLtTxt = VBA.ChrW(-10109) & VBA.ChrW(-8699)
        ElseIf aLtTxt Like "??八 -- " & VBA.ChrW(-10170) & VBA.ChrW(-8693) Then
            aLtTxt = VBA.ChrW(-10124) & VBA.ChrW(-9097)
        ElseIf aLtTxt Like VBA.ChrW(12282) & VBA.ChrW(-28746) & "商" Then
            aLtTxt = "適"
        ElseIf aLtTxt Like "??？ -- 狐" Then
            aLtTxt = "狐"
        ElseIf aLtTxt Like "??戔 -- 殘" Then
            aLtTxt = "殘"
        ElseIf aLtTxt Like "?????匹 -- 繼" Then
            aLtTxt = "繼"
        ElseIf aLtTxt Like "???么 -- " & VBA.ChrW(31762) Then
            aLtTxt = "篡"
        ElseIf aLtTxt Like "????凡 -- 彘" Then
            aLtTxt = "彘"
        ElseIf aLtTxt Like "?麻止 -- ?" Then
            aLtTxt = "歷"
        ElseIf aLtTxt Like VBA.ChrW(12282) & VBA.ChrW(-28746) & VBA.ChrW(17807) Then
            aLtTxt = "遽"
        ElseIf aLtTxt Like "?至支 -- ??" Then
            aLtTxt = "致"
        ElseIf aLtTxt Like "（???女）" Then
            aLtTxt = "嫈"
        ElseIf aLtTxt Like "（???力）" Then
            aLtTxt = VBA.ChrW(-10174) & VBA.ChrW(-9072)
        ElseIf aLtTxt Like "??? -- 懈" Then
            aLtTxt = "懈"
        ElseIf aLtTxt Like "（???）-- 釵" Then
            aLtTxt = "釵"
        ElseIf aLtTxt Like "?目兆 -- 晁" Then
            aLtTxt = "晁"
        ElseIf aLtTxt Like "???? -- " & VBA.ChrW(-10161) & VBA.ChrW(-8272) Then
            aLtTxt = "漆"
        ElseIf aLtTxt Like "?口? -- 噦" Then
            aLtTxt = "噦"
        ElseIf aLtTxt Like "?口? -- 呦" Then
            aLtTxt = "呦"
        ElseIf aLtTxt Like "???? -- 指" Then
            aLtTxt = "指"
        ElseIf aLtTxt Like "?夸?? -- 瓠" Then
            aLtTxt = VBA.ChrW(-10158) & VBA.ChrW(-8444)
        ElseIf aLtTxt Like "*page2700-20px-SKQSfont.pdf.jpg*" Then
            aLtTxt = "劇"
        ElseIf aLtTxt Like VBA.ChrW(12273) & VBA.ChrW(11966) & VBA.ChrW(12272) & VBA.ChrW(27701) & VBA.ChrW(20158) Then
            aLtTxt = VBA.ChrW(-10161) & VBA.ChrW(-8915)
        ElseIf aLtTxt Like "???止自匕?儿? -- 夔" Then
            aLtTxt = "夔"
        ElseIf aLtTxt Like "?穴之 -- 窆" Then
            aLtTxt = "窆"
        ElseIf aLtTxt Like VBA.ChrW(12272) & "目" & VBA.ChrW(-10170) & VBA.ChrW(-8693) Then
            aLtTxt = VBA.ChrW(-10121) & VBA.ChrW(-8228)
        ElseIf aLtTxt Like "??? -- 潤" Then
            aLtTxt = "潤"
        ElseIf aLtTxt Like "??? -- 靦" Then
            aLtTxt = "靦"
        ElseIf aLtTxt Like "??向 -- " & VBA.ChrW(-28664) Then
            aLtTxt = "迥"
        ElseIf aLtTxt Like "?日黽 -- " & VBA.ChrW(-24830) Then
            aLtTxt = VBA.ChrW(-24830)
        ElseIf aLtTxt Like "???????友-- 擾" Then
            aLtTxt = "擾"
        ElseIf aLtTxt Like "??? -- 癩" Then
            aLtTxt = "癩"
        ElseIf aLtTxt Like "（?血?）" Then
            aLtTxt = VBA.ChrW(-30654)
        ElseIf aLtTxt Like "SKchar" Then
            GoTo nxt
'            aLtTxt = "疾,優,虢,曷,姬,鮑,徑,梓,死（2DB7E）,鬼,灌,瓘,鸛,毓,褭,舁"'餘詳 查字.mdb
        ElseIf aLtTxt Like "SKchar2" Then
            GoTo nxt
'            aLtTxt = "纏（7E92）,丑,"'餘詳 查字.mdb
        Else
            Select Case aLtTxt
                Case VBA.ChrW(12280) & VBA.ChrW(30098) & VBA.ChrW(-28523)
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
Sub 千慮一得齋匯出() '20250318
    Dim db As New dBase, cnt As New ADODB.Connection, rst As New ADODB.Recordset, rstNote As New ADODB.Recordset, note As String, noteMark As String
    Dim d As Document, stPageNum As Integer, endPageNum As Integer, rng As Range, p As Paragraph, followWords As String, rngDup As Range, si As New StringInfo, ur As UndoRecord
    Rem 第1段是始頁，第2段指定末頁
    Set d = ActiveDocument
    stPageNum = VBA.CInt(d.Range(d.Paragraphs(1).Range.start, d.Paragraphs(1).Range.End - 1).text)
    endPageNum = VBA.CInt(d.Range(d.Paragraphs(2).Range.start, d.Paragraphs(2).Range.End - 1).text)
    
    SystemSetup.stopUndo ur, "千慮一得齋匯出"
    d.Range.text = vbNullString
    
    db.cnt_開發_千慮一得齋 cnt
    rst.Open "SELECT 札.札ID, 札.札記 FROM (書 LEFT JOIN 篇 ON 書.書ID = 篇.書ID) LEFT JOIN 札 ON 篇.篇ID = 札.篇ID " & _
                    "WHERE (((書.書ID)=9327) AND ((札.頁) Between " & stPageNum & " And " & endPageNum & ")) " & _
                    "ORDER BY 篇.頁, 篇.篇ID, 札.頁", cnt, adOpenForwardOnly, adLockReadOnly
    Do Until rst.EOF
        Set p = d.Paragraphs.Add
        Set rng = d.Range(p.Range.start, p.Range.End - 1)
        rng.InsertAfter rst.Fields("札記").Value '在套用此方法之後，該範圍就會展開成包含新的文字。
        Set rngDup = rng.Duplicate
        '19323:校,校勘記,真按    36171:注
        rstNote.Open "SELECT 札_類.類ID,札箋.札箋, 札箋.後續字元, 札箋.備註 FROM 札_類 INNER JOIN 札箋 ON 札_類.類_ID = 札箋.類_ID " & _
                        "WHERE (((札_類.札ID)=" & rst.Fields("札ID") & ") AND ((札_類.類ID)=36171 Or (札_類.類ID)=19323)) " & _
                        " order by st", cnt, adOpenKeyset, adLockReadOnly
        Do Until rstNote.EOF
            noteMark = rstNote.Fields("札箋").Value
findnext:
            If rng.Find.Execute(noteMark) Then
                followWords = VBA.Replace(VBA.IIf(VBA.IsNull(rstNote.Fields("後續字元").Value), vbNullString, rstNote.Fields("後續字元").Value), VBA.Chr(13) & VBA.Chr(10), VBA.Chr(13))
                If followWords <> vbNullString Then
                    Do Until d.Range(rng.End, rng.End + VBA.Len(followWords)).text = followWords
                        If Not rng.Find.Execute(rstNote.Fields("札箋").Value) Then Exit Do
                        If rng.End + VBA.Len(followWords) >= d.content.End Then GoTo nextRecord
                    Loop
                End If
            End If
            Select Case rstNote.Fields("類ID").Value
                Case 36171 '注
                    If rng.start > 0 Then
                        If rng.Previous(wdCharacter, 1) = "{" Then
                            rng.SetRange rng.End, rngDup.End
                            GoTo findnext
                        End If
                        rng.text = "{{" & rng.text & "}}"
                    End If
                Case 19323 '校,校勘記,真
                    note = rstNote.Fields("備註").Value
                    If VBA.InStr(note, "不復一一出校") = 0 Then
                        If rng.start > 0 Then
                            Do Until VBA.InStr("。，" & VBA.Chr(13), rng.Previous.text)
                                If rng.Previous.text <> VBA.Chr(13) Then rng.Move wdCharacter, 1
                                If rng.End = rng.Document.content.End - 1 Then Exit Do
                            Loop
                            si.Create noteMark
                            If si.LengthInTextElements > 1 Then
                                rng.InsertAfter "{{{孫守真按：" & "「" & noteMark & "」：" & note & "}}}"
                            Else
                                rng.InsertAfter "{{{孫守真按：" & noteMark & "，" & note & "}}}"
                            End If
                        End If
                    End If
            End Select
nextRecord:
            rng.SetRange rngDup.start, rngDup.End
            rstNote.MoveNext
        Loop
        rstNote.Close
        
        rst.MoveNext
    Loop
    
    rst.Close: cnt.Close
    
    文字處理.書名號篇名號標注
    rng.Document.content.Cut '剪下準備貼到TextForCtext的textBox1中
    SystemSetup.contiUndo ur
    
    d.ActiveWindow.windowState = wdWindowStateMinimize
    d.Range.InsertParagraphAfter
    d.Range(d.Paragraphs(1).Range.start, d.Paragraphs(1).Range.End - 1).text = endPageNum + 1
    d.Range(d.Paragraphs(2).Range.start, d.Paragraphs(2).Range.End - 1).text = endPageNum + 9
    
    On Error Resume Next
    AppActivate "TextForCtext"
    VBA.DoEvents
    SendKeys "^v", True
    VBA.DoEvents
    
End Sub
Rem 現在多用Kanripo.org者 20250202大年初五
Sub 元引科技引得數字人文資源平臺_北京元引科技有限公司轉來()
    Dim rng As Range, noteRng As Range, aNext As Range, aPre As Range, ur As UndoRecord, midNoteRngPos As Byte, midNoteRng As Range, aX As String, a As Range, aSt As Long, aEd As Long
    Dim noteFont As font '記下注文格式以備用
    Dim insertX As String, counter As Byte
    Set rng = Documents.Add().Range
    SystemSetup.stopUndo ur, "國學大師_Kanripo_四庫全書本轉來"
    SystemSetup.playSound 1
    rng.Paste
    '提示貼上無礙
    SystemSetup.playSound 1 '光貼上耗時就很久了，後面這一大堆式子反而快 20230211
    
    With rng.Find
        .font.ColorIndex = 6
    End With
    Set rng = rng.Document.Range
    '清除頁碼
    Do While rng.Find.Execute("P", , , , , , True, wdFindContinue)
       rng.Paragraphs(1).Range.Delete
    Loop
    rng.Find.ClearFormatting
    
    Set rng = rng.Document.Range
'    rng.Find.Execute "^p^p", , , , , , , wdFindContinue, , "^p", wdReplaceAll
'    If VBA.InStr(rng.text, VBA.ChrW(160) & "/" & VBA.Chr(11)) Then _
'        rng.Find.Execute VBA.ChrW(160) & "^g" & VBA.Chr(11), , , , , , , wdFindContinue, , VBA.Chr(11), wdReplaceAll 'chr(11)分行符號
'    If VBA.InStr(rng.text, VBA.ChrW(160) & "/" & VBA.Chr(13)) Then _
'        rng.Find.Execute VBA.ChrW(160) & "^g" & VBA.Chr(13), , , , , , , wdFindContinue, , VBA.Chr(13), wdReplaceAll
    
    rng.Find.Execute VBA.Chr(13), , , , , , , wdFindContinue, , VBA.Chr(11), wdReplaceAll
    rng.Find.Execute "^p/", , , , , , , wdFindContinue, , "^p", wdReplaceAll
        
    rng.Find.font.Color = 1310883
    Do While rng.Find.Execute(vbNullString, , , False, , , True, wdFindStop)
        If noteFont Is Nothing Then Set noteFont = rng.font
        Set noteRng = rng '.Document.Range(rng.start, rng.End)
        Do While noteRng.Next.font.Color = 1310883
            noteRng.SetRange noteRng.start, noteRng.Next.End
        Loop
        
'        If InStr(noteRng, "萁草之句") Then Stop 'just for test
        
        Set aNext = noteRng.Characters(noteRng.Characters.Count).Next
        Set aPre = noteRng.Characters(1).Previous
        midNoteRngPos = Excel.RoundUpCustom(noteRng.Characters.Count / 2)
        
        Set midNoteRng = noteRng.Document.Range(noteRng.Characters(VBA.IIf(midNoteRngPos - 1 < 1, 1, midNoteRngPos - 1)).start _
            , noteRng.Characters(VBA.IIf(midNoteRngPos + 1 > noteRng.Characters.Count, noteRng.Characters.Count, midNoteRngPos + 1)).End)
        If midNoteRng.start = noteRng.start And midNoteRng.End = noteRng.End Then
            Set midNoteRng = noteRng
        End If
'        If (aNext.text = VBA.Chr(11) And aPre.text = VBA.Chr(11)) Then
'            If aNext.Previous = "/" Then
'                midNoteRng.text = VBA.Replace(midNoteRng, "/", vbNullString, 1, 1)
'                noteRng.text = "{{" & noteRng.text & "}}"
'            Else
'                midNoteRng.text = VBA.Replace(midNoteRng, "/", VBA.Chr(11), 1, 1)
'                noteRng.text = "{{" & noteRng.text & "}}"
'            End If
'        ElseIf aNext.text = VBA.Chr(13) And aPre.text = VBA.Chr(13) Then
'            If aNext.Previous = "/" Then
'                midNoteRng.text = VBA.Replace(midNoteRng, "/", vbNullString, 1, 1)
'                noteRng.text = "{{" & noteRng.text & "}}"
'            Else
'                midNoteRng.text = VBA.Replace(midNoteRng, "/", VBA.Chr(13), 1, 1)
'                noteRng.text = "{{" & noteRng.text & "}}"
'            End If
'        Else
'            If aNext.text = VBA.Chr(11) Then


'        If InStr(noteRng, "適/") Then Stop


                '判斷有無縮排
                If Not aPre Is Nothing Then
                    Set a = aPre.Document.Range(aPre.start, aPre.End) '記下aPre原來的位置
                    If aPre.start > 0 And aPre.text <> VBA.Chr(11) Then
                        Do Until aPre.Previous = VBA.Chr(11)
                            aPre.Move wdCharacter, -1
                            If aPre.start <= 0 Then Exit Do
                        Loop
                    End If
                    If a.start > aPre.start Then 'a =aPre原來的位置
                        a.SetRange aPre.start, a.End
                        aX = a.text '縮排的空格
                    Else
                        If a.text = aPre.text Then
                            If aPre.text = "　" Then '有縮排
                                aX = a.text
                            Else
                                aX = vbNullString
'                                SystemSetup.playSound 12, 0
'                                Stop
                            End If
                        Else
                            aX = vbNullString
                        End If
                    End If
                End If
                
'                Dim line As New LineChr11
                
                '如果有縮排('aX=縮排的空格)
                If aX <> vbNullString And VBA.Replace(aX, "　", vbNullString) = vbNullString Then
                    If noteRng.Next Is Nothing Then '怕在文件最末端，與下一年判斷並不重複
'                    If line.LineRange(noteRng).start = noteRng.start And line.LineRange(noteRng).End = noteRng.End Then
                        insertX = VBA.Chr(11) & aX
                    ElseIf noteRng.Next = VBA.Chr(11) Then 'ax=縮排的空格 ●●●●●●●●●●●●●
                        insertX = VBA.Chr(11) & aX  'VBA.Chr(11) 後面 a.text = "}}" & VBA.Replace(insertX, VBA.Chr(11), VBA.Chr(11) & "{{") 要參照
                    Else
                        If VBA.InStr(midNoteRng.text, "/") _
                            And noteRng.Next.font.Size > 11.5 _
                            And (noteRng.Next.text <> VBA.Chr(11) Or noteRng.Next.text = "　") Then  '若是夾注(通常是標題下的夾注（則後面有空格），如 https://ctext.org/library.pl?if=en&file=55677&page=6） 20250205
                            'noteRng.Next.text <> VBA.Chr(11):後面還有文字，則為夾注 20250223補
                            insertX = aX '補空格以縮排
                        Else
                            insertX = vbNullString
                        End If
                    End If
                Else '沒有縮排
                    If aX = vbNullString And Not noteRng.Previous(wdCharacter, 1) Is Nothing And Not noteRng.Next(wdCharacter, 1) Is Nothing Then
                        If noteRng.Previous(wdCharacter, 1) = VBA.Chr(11) And noteRng.Next(wdCharacter, 1) = VBA.Chr(11) Then
                            insertX = VBA.Chr(11)
                        Else
'                            SystemSetup.playSound 7, 0
'                            Stop
                            insertX = vbNullString
                        End If
                    Else
                        insertX = vbNullString
                    End If
                End If
                
                
                For Each a In noteRng.Characters '找到/（夾注換行）的位置
                    If a = "/" And a.InlineShapes.Count = 0 Then
                        If a.font.Color = noteFont.Color And a.font.Size = noteFont.Size Then
                            aSt = a.start
                            aEd = a.End
                            
                            Do Until VBA.Abs(noteRng.Document.Range(noteRng.start, a.start).Characters.Count - VBA.IIf(a.End = noteRng.End, 0, noteRng.Document.Range(a.End, noteRng.End).Characters.Count)) < 2
                               'noteRng.Document.Range(a.End, noteRng.End).text = noteRng.Document.Range(a.End, noteRng.End).text & "　"
                               noteRng.text = noteRng.text & "　"
                               a.SetRange aSt, aEd
                               If rng.End + 3 >= rng.Document.Range.End Then Exit Do
                               counter = counter + 1
                               If counter > 50 Then Exit Do
                            Loop
                            counter = 0
                            If a.Next = VBA.Chr(11) Then '如果斜線/後面即換行
                                If aX = vbNullString Or VBA.Replace(aX, "　", vbNullString) <> vbNullString Then '若無縮排，則清除掉斜線/
                                    a.text = vbNullString
                                Else '有縮排時
                                    a.text = insertX '●●●●●●●●●●●●●再觀察
                                    noteRng.SetRange aSt, aEd + VBA.Len(insertX) - 1 '「/」（ a = "/" ）拿掉了故減1
                                End If
                            Else
                                If noteRng.Next = VBA.Chr(11) And aX <> vbNullString And VBA.Replace(aX, "　", vbNullString) = vbNullString Then
                                'If noteRng.Next = VBA.Chr(11) And VBA.Replace(aPre.text, "　", vbNullString) = vbNullString Then
                                    If aPre.Previous = VBA.Chr(11) Then
                                        noteRng.SetRange aPre.start, noteRng.End
                                        a.text = "}}" & VBA.Replace(insertX, VBA.Chr(11), VBA.Chr(11) & "{{")
                                    Else
                                        SystemSetup.playSound 12, 0
                                        Stop
                                    End If
                                Else
                                    a.text = insertX
                                End If
                            End If
                            Exit For
                        End If
                    End If
                Next a
                If insertX <> vbNullString And VBA.Replace(insertX, "　", vbNullString) <> vbNullString Then '如果置換「/」的字符不是空字串也不是縮排用的空格
                    If aX <> vbNullString Then
                        aSt = noteRng.start
                        noteRng.SetRange aPre.start, noteRng.End
                    End If
                    noteRng.text = "{{" & noteRng.text & "}}"
                    noteRng.Collapse wdCollapseEnd
                Else
'                   midNoteRng.text = VBA.Replace(midNoteRng, "/", vbNullString, 1, 1)
                    If aX <> vbNullString And VBA.Replace(aX, "　", vbNullString) = vbNullString Then '●●●●●●●●●●●●
                        '如果有縮排，則擴展noteRng至前後全形空格的兩端
                        noteRng.MoveStartWhile "　", -50
                        '如果夾注中沒有縮排補上的空格
                        If a.text <> "　" Then
                            noteRng.MoveEndWhile "　", 50
                        End If
                        noteRng.InsertBefore "{{"
                        noteRng.InsertAfter "}}"
                        rng.SetRange rng.End, rng.End
                    Else
                        noteRng.text = "{{" & noteRng.text & "}}"
                    End If
                    
                End If
'            Else
'                midNoteRng.text = VBA.Replace(midNoteRng, "/", vbNullString, 1, 1)
'                noteRng.text = "{{" & noteRng.text & "}}"
'            End If
'        End If
    Loop
    
'    'word.Application.Activate'在背景執行Word（即不見Word）時不能如此，會出錯
'    SystemSetup.playSound 3
'    If VBA.MsgBox("空格轉成空白？", vbOKCancel + vbExclamation) = vbOK Then
'        國學大師_Kanripo_四庫全書本轉來_Sub rng.Document.Content
'    End If
    
    SystemSetup.playSound 1
    '文字處理.書名號篇名號標注 '擬交給TextForCtext C#來標 20250312
    
    With rng.Document
        With .Range.Find
            .ClearFormatting
    '        .Text = vba.Chrw(9675)
            .text = "}}{{　}}"
    '        .Replacement.Text = vba.Chrw(12295)
            .Replacement.text = "　}}"
            .Execute , , , , , , True, wdFindContinue, , , wdReplaceAll
        End With
        .Range.text = Replace(Replace(.Range.text, Chr(11), Chr(13) & Chr(10)), "?", "/")
        .Range.Cut
        
'        If VBA.InStr(.Range.text, "{{}}") Then
'            SystemSetup.playSound 12, 0
'        End If

'        SystemSetup.ClipboardPutIn Replace(Replace(.Range.text, Chr(11), Chr(13) & Chr(10)), "?", "/")
        DoEvents
        If .Application.Visible Then .Application.windowState = wdWindowStateMinimize
        .Close wdDoNotSaveChanges
        
    End With
    SystemSetup.playSound 1.921
    SystemSetup.contiUndo ur
    
    
'    AppActivate "TextForCtext"
'    DoEvents
'    SendKeys "^v"
'    DoEvents
End Sub
Rem 現在多用Kanripo.org者 20250202大年初五
Sub 國學大師_Kanripo_四庫全書本轉來()
    Dim rng As Range, noteRng As Range, aNext As Range, aPre As Range, ur As UndoRecord, midNoteRngPos As Byte, midNoteRng As Range, aX As String, a As Range, aSt As Long, aEd As Long
    Dim noteFont As font '記下注文格式以備用
    Dim insertX As String
    Set rng = Documents.Add().Range
    SystemSetup.stopUndo ur, "國學大師_Kanripo_四庫全書本轉來"
    SystemSetup.playSound 1
    
    'P 乃「北京元引科技有限公司《元引科技引得數字人文資源平臺·中國歷代文獻》」的文本特徵
    If VBA.InStr(SystemSetup.GetClipboard, "P") Then
        元引科技引得數字人文資源平臺_北京元引科技有限公司轉來
        Exit Sub
    End If
    
    rng.Paste
    '提示貼上無礙
    SystemSetup.playSound 1 '光貼上耗時就很久了，後面這一大堆式子反而快 20230211
    
    'With rng.Find
    '    .ClearAllFuzzyOptions
    '    .ClearFormatting
    '    .Execute "^l", , , True, , , True, wdFindContinue, , "^p", wdReplaceAll
    'End With
    With rng.Find
        .ClearAllFuzzyOptions
        .ClearFormatting
        .MatchWildcards = True
        .Execute "[[]*[]]  ", , , True, , , True, wdFindContinue, , vbNullString, wdReplaceAll
        .ClearAllFuzzyOptions
        .ClearFormatting
    End With
    Do While rng.Find.Execute("[[]", , , , , , True, wdFindContinue)
       rng.MoveEndUntil "]"
       rng.SetRange rng.start, rng.End + 1
       rng.Delete
    Loop
    Set rng = rng.Document.Range
    rng.Find.Execute "^p^p", , , , , , , wdFindContinue, , "^p", wdReplaceAll
    If VBA.InStr(rng.text, VBA.ChrW(160) & "/" & VBA.Chr(11)) Then _
        rng.Find.Execute VBA.ChrW(160) & "^g" & VBA.Chr(11), , , , , , , wdFindContinue, , VBA.Chr(11), wdReplaceAll 'chr(11)分行符號
    If VBA.InStr(rng.text, VBA.ChrW(160) & "/" & VBA.Chr(13)) Then _
        rng.Find.Execute VBA.ChrW(160) & "^g" & VBA.Chr(13), , , , , , , wdFindContinue, , VBA.Chr(13), wdReplaceAll
        
    rng.Find.ClearFormatting
    
    
    rng.Find.font.Color = 16711935
    Do While rng.Find.Execute(vbNullString, , , False, , , True, wdFindStop)
        If noteFont Is Nothing Then Set noteFont = rng.font
        Set noteRng = rng '.Document.Range(rng.start, rng.End)
        Do While noteRng.Next.font.Color = 16711935
            noteRng.SetRange noteRng.start, noteRng.Next.End
        Loop
        
'        If InStr(noteRng, "萁草之句") Then Stop 'just for test
        
        Set aNext = noteRng.Characters(noteRng.Characters.Count).Next
        Set aPre = noteRng.Characters(1).Previous
        midNoteRngPos = Excel.RoundUpCustom(noteRng.Characters.Count / 2)
        
        Set midNoteRng = noteRng.Document.Range(noteRng.Characters(VBA.IIf(midNoteRngPos - 1 < 1, 1, midNoteRngPos - 1)).start _
            , noteRng.Characters(VBA.IIf(midNoteRngPos + 1 > noteRng.Characters.Count, noteRng.Characters.Count, midNoteRngPos + 1)).End)
        If midNoteRng.start = noteRng.start And midNoteRng.End = noteRng.End Then
            Set midNoteRng = noteRng
        End If
'        If (aNext.text = VBA.Chr(11) And aPre.text = VBA.Chr(11)) Then
'            If aNext.Previous = "/" Then
'                midNoteRng.text = VBA.Replace(midNoteRng, "/", vbNullString, 1, 1)
'                noteRng.text = "{{" & noteRng.text & "}}"
'            Else
'                midNoteRng.text = VBA.Replace(midNoteRng, "/", VBA.Chr(11), 1, 1)
'                noteRng.text = "{{" & noteRng.text & "}}"
'            End If
'        ElseIf aNext.text = VBA.Chr(13) And aPre.text = VBA.Chr(13) Then
'            If aNext.Previous = "/" Then
'                midNoteRng.text = VBA.Replace(midNoteRng, "/", vbNullString, 1, 1)
'                noteRng.text = "{{" & noteRng.text & "}}"
'            Else
'                midNoteRng.text = VBA.Replace(midNoteRng, "/", VBA.Chr(13), 1, 1)
'                noteRng.text = "{{" & noteRng.text & "}}"
'            End If
'        Else
'            If aNext.text = VBA.Chr(11) Then


'        If InStr(noteRng, "適/") Then Stop


                '判斷有無縮排
                If Not aPre Is Nothing Then
                    Set a = aPre.Document.Range(aPre.start, aPre.End) '記下aPre原來的位置
                    If aPre.start > 0 And aPre.text <> VBA.Chr(11) Then
                        Do Until aPre.Previous = VBA.Chr(11)
                            aPre.Move wdCharacter, -1
                            If aPre.start <= 0 Then Exit Do
                        Loop
                    End If
                    If a.start > aPre.start Then 'a =aPre原來的位置
                        a.SetRange aPre.start, a.End
                        aX = a.text '縮排的空格
                    Else
                        If a.text = aPre.text Then
                            If aPre.text = "　" Then '有縮排
                                aX = a.text
                            Else
                                aX = vbNullString
'                                SystemSetup.playSound 12, 0
'                                Stop
                            End If
                        Else
                            aX = vbNullString
                        End If
                    End If
                End If
                
'                Dim line As New LineChr11
                
                '如果有縮排('aX=縮排的空格)
                If aX <> vbNullString And VBA.Replace(aX, "　", vbNullString) = vbNullString Then
                    If noteRng.Next Is Nothing Then '怕在文件最末端，與下一年判斷並不重複
'                    If line.LineRange(noteRng).start = noteRng.start And line.LineRange(noteRng).End = noteRng.End Then
                        insertX = VBA.Chr(11) & aX
                    ElseIf noteRng.Next = VBA.Chr(11) Then 'ax=縮排的空格 ●●●●●●●●●●●●●
                        insertX = VBA.Chr(11) & aX  'VBA.Chr(11) 後面 a.text = "}}" & VBA.Replace(insertX, VBA.Chr(11), VBA.Chr(11) & "{{") 要參照
                    Else
                        If VBA.InStr(midNoteRng.text, "/") _
                            And noteRng.Next.font.Size > 11.5 _
                            And (noteRng.Next.text <> VBA.Chr(11) Or noteRng.Next.text = "　") Then  '若是夾注(通常是標題下的夾注（則後面有空格），如 https://ctext.org/library.pl?if=en&file=55677&page=6） 20250205
                            'noteRng.Next.text <> VBA.Chr(11):後面還有文字，則為夾注 20250223補
                            insertX = aX '補空格以縮排
                        Else
                            insertX = vbNullString
                        End If
                    End If
                Else '沒有縮排
                    If aX = vbNullString And Not noteRng.Previous(wdCharacter, 1) Is Nothing And Not noteRng.Next(wdCharacter, 1) Is Nothing Then
                        If noteRng.Previous(wdCharacter, 1) = VBA.Chr(11) And noteRng.Next(wdCharacter, 1) = VBA.Chr(11) Then
                            insertX = VBA.Chr(11)
                        Else
'                            SystemSetup.playSound 7, 0
'                            Stop
                            insertX = vbNullString
                        End If
                    Else
                        insertX = vbNullString
                    End If
                End If
                
                
                For Each a In noteRng.Characters '找到/（夾注換行）的位置
                    If a = "/" And a.InlineShapes.Count = 0 Then
                        If a.font.Color = noteFont.Color And a.font.Size = noteFont.Size Then
                            aSt = a.start
                            aEd = a.End
                            
                            Do Until VBA.Abs(noteRng.Document.Range(noteRng.start, a.start).Characters.Count - VBA.IIf(a.End = noteRng.End, 0, noteRng.Document.Range(a.End, noteRng.End).Characters.Count)) < 2
                               'noteRng.Document.Range(a.End, noteRng.End).text = noteRng.Document.Range(a.End, noteRng.End).text & "　"
                               noteRng.text = noteRng.text & "　"
                               a.SetRange aSt, aEd
                            Loop
                            If a.Next = VBA.Chr(11) Then '如果斜線/後面即換行
                                If aX = vbNullString Or VBA.Replace(aX, "　", vbNullString) <> vbNullString Then '若無縮排，則清除掉斜線/
                                    a.text = vbNullString
                                Else '有縮排時
                                    a.text = insertX '●●●●●●●●●●●●●再觀察
                                    noteRng.SetRange aSt, aEd + VBA.Len(insertX) - 1 '「/」（ a = "/" ）拿掉了故減1
                                End If
                            Else
                                If noteRng.Next = VBA.Chr(11) And aX <> vbNullString And VBA.Replace(aX, "　", vbNullString) = vbNullString Then
                                'If noteRng.Next = VBA.Chr(11) And VBA.Replace(aPre.text, "　", vbNullString) = vbNullString Then
                                    If aPre.Previous = VBA.Chr(11) Then
                                        noteRng.SetRange aPre.start, noteRng.End
                                        a.text = "}}" & VBA.Replace(insertX, VBA.Chr(11), VBA.Chr(11) & "{{")
                                    Else
                                        SystemSetup.playSound 12, 0
                                        Stop
                                    End If
                                Else
                                    a.text = insertX
                                End If
                            End If
                            Exit For
                        End If
                    End If
                Next a
                If insertX <> vbNullString And VBA.Replace(insertX, "　", vbNullString) <> vbNullString Then '如果置換「/」的字符不是空字串也不是縮排用的空格
                    If aX <> vbNullString Then
                        aSt = noteRng.start
                        noteRng.SetRange aPre.start, noteRng.End
                    End If
                    noteRng.text = "{{" & noteRng.text & "}}"
                    noteRng.Collapse wdCollapseEnd
                Else
'                   midNoteRng.text = VBA.Replace(midNoteRng, "/", vbNullString, 1, 1)
                    If aX <> vbNullString And VBA.Replace(aX, "　", vbNullString) = vbNullString Then '●●●●●●●●●●●●
                        '如果有縮排，則擴展noteRng至前後全形空格的兩端
                        noteRng.MoveStartWhile "　", -50
                        If Not a Is Nothing Then
                        '如果夾注中沒有縮排補上的空格
                            If a.text <> "　" Then
                                noteRng.MoveEndWhile "　", 50
                            End If
                        End If
                        noteRng.InsertBefore "{{"
                        noteRng.InsertAfter "}}"
                        rng.SetRange rng.End, rng.End
                    Else
                        noteRng.text = "{{" & noteRng.text & "}}"
                    End If
                    
                End If
'            Else
'                midNoteRng.text = VBA.Replace(midNoteRng, "/", vbNullString, 1, 1)
'                noteRng.text = "{{" & noteRng.text & "}}"
'            End If
'        End If
    Loop
    
'    'word.Application.Activate'在背景執行Word（即不見Word）時不能如此，會出錯
'    SystemSetup.playSound 3
'    If VBA.MsgBox("空格轉成空白？", vbOKCancel + vbExclamation) = vbOK Then
'        國學大師_Kanripo_四庫全書本轉來_Sub rng.Document.Content
'    End If
    
    SystemSetup.playSound 1
    '文字處理.書名號篇名號標注 '擬交給TextForCtext C#來標 20250312
    
    With rng.Document
        With .Range.Find
            .ClearFormatting
    '        .Text = vba.Chrw(9675)
            .text = "}}{{　}}"
    '        .Replacement.Text = vba.Chrw(12295)
            .Replacement.text = "　}}"
            .Execute , , , , , , True, wdFindContinue, , , wdReplaceAll
        End With
        '.Range.Cut
        
'        If VBA.InStr(.Range.text, "{{}}") Then
'            SystemSetup.playSound 12, 0
'        End If

        SystemSetup.ClipboardPutIn .Range.text
        DoEvents
        .Close wdDoNotSaveChanges
    End With
    SystemSetup.playSound 1.921
    SystemSetup.contiUndo ur
End Sub
Rem 作為 國學大師_四庫全書本轉來()的子程序:1.取代空格為空白
Sub 國學大師_Kanripo_四庫全書本轉來_Sub(rng As Range)
    Dim rngEd As Long, rngChk As Range, rngChkX As String, rngChkPre As Range
    Do While rng.Find.Execute("　")
        rngEd = rng.End
        Set rngChk = rng.Document.Range(rng.start, rng.End)
        rngChk.SetRange rng.start + rngChk.MoveStartUntil(VBA.Chr(11), -(rng.End - 1)) + 1, rngEd
        rngChkX = VBA.Replace(rngChk.text, "　", vbNullString)
        Set rngChkPre = rng.Previous
        If rngChkX <> vbNullString And rngChkX <> "{{" Then
            If rng.Previous.text <> VBA.Chr(11) Then
                GoSub replaceSpaceWithBlank:
            End If
        ElseIf Not rngChkPre Is Nothing Then
            If rngChkPre.text <> VBA.Chr(11) And VBA.Left(rngChk, 2) <> "{{" Then GoSub replaceSpaceWithBlank:
        End If
        rng.SetRange rngEd, rng.Document.content.End
    Loop
    
Exit Sub
replaceSpaceWithBlank:
    rngChk.SetRange rng.start, rng.End
    Dim line As New LineChr11
    If VBA.InStr(rngChk.Document.Range(rngChk.End, line.EndPosition(rngChk)).text, "}") Then
        'rngChk.MoveEndUntil "}", line.LineRange(rngChk).End - rng.End
        rngChk.MoveEndUntil "}", line.EndPosition(rngChk) - rng.End
        rngChkX = rngChk.text
        If VBA.Replace(rngChkX, "　", vbNullString) <> vbNullString Then
            rngEd = rng.End
            rng.text = VBA.ChrW(-9217) & VBA.ChrW(-8195) '取代空格為空白
        End If
    Else
        rngEd = rng.End
        rng.text = VBA.ChrW(-9217) & VBA.ChrW(-8195) '取代空格為空白
    End If
    Return
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
            exportStr = exportStr & VBA.Chr(13) & "*" & title & VBA.Chr(13)
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
        rng.text = Replace(rng, e, "")
    Next e
    rng.text = "#" & rng.text
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
Sub 插入超連結_將顯示之編碼改為中文()
Const keys As String = "&searchu=" 'Alt + j
Dim rng As Range, lnk As String, cde As String, s As Long, d As Document, ur As UndoRecord
Set rng = Selection.Range: Set d = ActiveDocument
lnk = SystemSetup.GetClipboardText
cde = VBA.Mid(lnk, InStr(lnk, keys) + Len(keys))
cde = code.URLDecode(cde)
s = Selection.start
SystemSetup.stopUndo ur, "插入超連結_將顯示之編碼改為中文"
With Selection
    .Hyperlinks.Add Selection.Range, lnk, , , VBA.Left(lnk, InStr(lnk, keys) + Len(keys) - 1) + cde
    'd.Range(Selection.End, Selection.End + Len(cde)).Select
    'Selection.Collapse
    .MoveLeft wdCharacter, Len(cde)
    .MoveRight wdCharacter, Len(cde) - 1, wdExtend
    .Range.HighlightColorIndex = wdYellow
    .Move , 2
    .InsertParagraphAfter
    .InsertParagraphAfter
    .Collapse
End With
SystemSetup.contiUndo ur
End Sub
Sub 只保留正文注文_且注文前後加括弧(d As Document)
    Dim ur As UndoRecord, slRng As Range
    SystemSetup.stopUndo ur, "中國哲學書電子化計劃_只保留正文注文_且注文前後加括弧"
'    Set d = Docs.空白的新文件()
'    Set d = ActiveDocument
'    d.Activate
    d.Range.Paste
    'If Selection.Type = wdSelectionIP Then ActiveDocument.Select
'    Set slRng = Selection.Range
    Set slRng = d.Range
    '圖文對照，Quict edit，單頁頁面 20240912
    If InStr(slRng.text, "<p>") Or (InStr(slRng.text, "{") And InStr(slRng.text, "}")) Then
        If InStr(slRng.text, "{{{") Then
            slRng.Find.ClearAllFuzzyOptions: slRng.Find.ClearFormatting
            slRng.Find.Execute "{{{*}}}", , , True, , , True, wdFindContinue, , vbNullString, wdReplaceAll
        End If
        slRng.Find.ClearAllFuzzyOptions: slRng.Find.ClearFormatting
        slRng.Find.Execute "^p", , , , , , , , , vbNullString, wdReplaceAll
        slRng.Find.Execute "<p>", , , , , , , , , vbNullString, wdReplaceAll
        slRng.Find.Execute "{{", , , , , , , , , "（", wdReplaceAll
        slRng.Find.Execute "}}", , , , , , , , , "）", wdReplaceAll
    Else 'Edit、View，篇章節單位頁面
        清除文本頁中的編號儲存格 slRng
        中國哲學書電子化計劃_表格轉文字 slRng
        Dim ay, e
        ay = Array(254, 8912896)
        With d.Range.Find
            .ClearFormatting
        End With
        For Each e In ay
            With d.Range.Find
                .font.Color = e
                .Execute "", , , , , , True, wdFindContinue, , "", wdReplaceAll
            End With
        Next e
        Set slRng = d.Range
        With slRng.Find
            .ClearFormatting
            .font.Color = 34816
        End With
        Do While slRng.Find.Execute(, , , , , , True, wdFindStop)
            If InStr(VBA.Chr(13) & VBA.Chr(11) & VBA.Chr(7) & VBA.Chr(8) & VBA.Chr(9) & VBA.Chr(10), slRng) = 0 Then
            slRng.text = "（" + slRng.text + "）"
            'slRng.SetRange slRng.End, d.Range.End
            End If
        Loop
    End If
    SystemSetup.contiUndo ur
End Sub

Sub 清除文本頁中的編號儲存格(rng As Range)
    Dim c As cell, cx As String, t As table
    For Each t In rng.tables
        For Each c In t.Range.cells
            c.Select
            cx = c.Range.text
            If VBA.IsNumeric(VBA.Left(cx, 1)) And VBA.InStr(cx, VBA.ChrW(160) & VBA.ChrW(47)) > 0 And c.Range.InlineShapes.Count = 1 And VBA.Len(cx) < 13 Then
                If VBA.InStr(cx, VBA.Val(cx) & VBA.ChrW(160) & VBA.ChrW(47)) = 1 Then
                    c.Delete
                End If
            End If
        Next c
    Next t
End Sub

Sub 維基文庫等欲直接抽換之字(d As Document)
    Dim rst As New ADODB.Recordset, cnt As New ADODB.Connection, db As New dBase
    db.cnt查字 cnt
    rst.Open "select * from 維基文庫等欲直接抽換之字 where doIt = true order by len(replaced) desc", cnt, adOpenForwardOnly, adLockReadOnly
    Do Until rst.EOF
        d.Range.Find.Execute rst.Fields("replaced").Value, , , , , , True, wdFindContinue, , rst.Fields("replacewith").Value, wdReplaceAll
        rst.MoveNext
    Loop
    rst.Close: cnt.Close: Set db = Nothing
End Sub

Sub dbSBCKWordtoReplace() '四部叢刊造字對照表 Alt+5
    Dim rng As Range, ur As UndoRecord
    'Set ur = stopUndo("《四部叢刊》資料庫造字取代為系統字")
    SystemSetup.stopUndo ur, "《四部叢刊》資料庫造字取代為系統字"
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
        If InStr(rng.text, rst.Fields(0).Value) Then _
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
            Selection.Move
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
        ay(i) = VBA.Split(rng.text, ">")
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
                rngMark.Move wdCharacter, -1
            Loop
            rngMark.SetRange rngMark.start, rng.start
            If VBA.Left(rngMark.text, 8) <> "<entity " Then
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
        rng.Find.font.Color = 8912896 '{{{}}}語法下的文字
        rng.Find.Replacement.ClearFormatting
        With rng.Find
            .text = ""
            .Replacement.text = ""
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
            If InStr(rng.text, e) > 0 Then
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

Sub EditModeMakeup_changeFile_Page() '同版本文本帶入置換file id 和 頁數
    Dim rng As Range, pageNum As Range, d As Document, ur As UndoRecord
    Set d = ActiveDocument
    
    '文件前3段分別是以下資訊,執行完會清除
    'If Not VBA.IsNumeric(VBA.Replace(d.Range.Paragraphs(1).Range.text, vba.Chr(13), "")) then
    If VBA.Replace(d.Paragraphs(1).Range + d.Paragraphs(2).Range + d.Paragraphs(3).Range, VBA.Chr(13), "") = "" _
        Or Not IsNumeric(VBA.Replace(d.Paragraphs(1).Range, VBA.Chr(13), "")) Then
        If Not IsNumeric(VBA.Replace(VBA.Replace(d.Paragraphs(1).Range + d.Paragraphs(2).Range + d.Paragraphs(3).Range, VBA.Chr(13), ""), "-", vbNullString, 1, 1)) Then
            MsgBox "請在文件前3段分別是以下資訊（皆是數字）,執行完會清除" & vbCr & vbCr & _
                "1. 頁數差(來源-(減去)目的）。無頁差則為0，省略則預設為0" & vbCr & vbCr & _
                 "也可輸入「來源-目的」這樣的格式，如來源114頁，目的是69頁，可以「114-69」表示" & vbCr & _
                "2. 目的的 file number。要置換成的；不取代則為0，省略則預設為0" & vbCr & _
                "3. 來源的 file number，要被取代的,省略（仍要空其段落=空行）則取文件中的file=後的值"
            Exit Sub
        End If
    End If
    Dim differPageNum  As Integer '頁數差(來源-(減去)目的）
    Dim numRngDashPost As Byte, numRng As Range
    numRngDashPost = VBA.InStr(d.Paragraphs(1).Range.text, "-")
    If numRngDashPost > 1 Then '=1 是負數標識
        Set numRng = d.Range(d.Paragraphs(1).Range.Characters(1).start, d.Paragraphs(1).Range.Characters(d.Paragraphs(1).Range.Characters.Count).start)
        numRng.text = VBA.CInt(VBA.Left(numRng.text, numRngDashPost - 1)) - VBA.CInt(Mid(numRng.text, numRngDashPost + 1))
    End If
    differPageNum = VBA.IIf(d.Paragraphs(1).Range.Characters.Count = 1, 0, VBA.Replace(d.Paragraphs(1).Range.text, VBA.Chr(13), "")) '頁數差(來源-(減去)目的）
    Dim file
    file = VBA.Replace(d.Paragraphs(2).Range.text, VBA.Chr(13), "") ' 目的。不取代則為0
    If file = "" Then file = 0
    Dim fileFrom As String
    fileFrom = VBA.Replace(d.Paragraphs(3).Range.text, VBA.Chr(13), "") ' '來源
    If fileFrom = "" Then
        Dim s As String: s = VBA.InStr(d.Range.text, "<scanbegin file="): s = s + VBA.Len("<scanbegin file=")
        fileFrom = VBA.Mid(d.Range.text, s + 1, InStr(s + 1, d.Range.text, """") - s - 1)
    End If
    Set rng = d.Range
    'Set ur = SystemSetup.stopUndo("EditMakeupCtext")
    SystemSetup.stopUndo ur, "EditMakeupCtext"
    If file > 0 Then
        'rng.Find.Execute " file=""77991""", True, True, , , , True, wdFindContinue, , " file=""" & file & """", wdReplaceAll
        rng.text = Replace(rng.text, " file=""" & fileFrom & """", " file=""" & file & """")
    End If

    Do While rng.Find.Execute(" page=""", , , , , , True, wdFindStop)
        Set pageNum = rng
        pageNum.SetRange rng.End, rng.End + 1
        pageNum.MoveEndUntil """"
        pageNum.text = CStr(CInt(pageNum.text) - differPageNum)
        rng.SetRange pageNum.End, d.Range.End
    Loop
    rng.SetRange d.Range.Paragraphs(1).Range.start, d.Range.Paragraphs(3).Range.End
    rng.Delete
    'd.Range.Cut
    SystemSetup.SetClipboard d.Range.text
    SystemSetup.contiUndo ur
    SystemSetup.playSound 1
'    d.Application.Activate
End Sub
Property Get Div_generic_IncludePathAndEndPageNum() As SeleniumBasic.IWebElement
    Dim iwe As SeleniumBasic.IWebElement
    'if Form1.IsValidUrl＿ImageTextComparisonPage(ActiveForm1.textBox3Text))
    Set iwe = WD.FindElementByCssSelector("#content > div:nth-child(3)")
    Set Div_generic_IncludePathAndEndPageNum = iwe
End Property
Rem 取得某書冊頁之上限
Property Get pageUBound() As Integer
    Dim iwe  As SeleniumBasic.IWebElement, str As String
    Set iwe = Div_generic_IncludePathAndEndPageNum
    If iwe Is Nothing Then pageUBound = 0
    str = iwe.GetAttribute("textContent") '"線上圖書館 -> 松煙小錄 -> 松煙小錄三  /117 ";
    pageUBound = VBA.CInt(VBA.Mid(str, VBA.InStr(str, "/") + 1, VBA.Len(str) - 1 - VBA.InStr(str, "/")))
End Property
Function CurrentChapterNum_Selector() As String
    Dim selector As String
    Dim match As Object
    Dim regex As Object
    
    ' 設定選擇器字符串
    selector = ChapterSelector '"#content > div:nth-child(6) > table > tbody > tr:nth-child(2) > td:nth-child(1) > a"
    
    ' 建立正則表達式對象
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "tr:nth-child\((\d+)\)"
    regex.Global = False
    
    ' 進行匹配
    Set match = regex.Execute(selector)
    If match.Count > 0 Then
        ' 取得匹配的群組值
        CurrentChapterNum_Selector = match(0).SubMatches(0)
    Else
        ' 若無匹配，返回空字串
        CurrentChapterNum_Selector = ""
    End If
End Function
Function IncrementNthChild(selector As String) As String
    Dim regex As Object
    Dim match As Object
    Dim number As Integer
    
    ' 建立正則表達式對象
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "tr:nth-child\((\d+)\)"
    regex.Global = False
    
    ' 執行匹配
    Set match = regex.Execute(selector)
    If match.Count > 0 Then
        ' 取得群組中的數字，並轉為整數
        number = CInt(match(0).SubMatches(0))
        number = number + 1
        
        ' 使用正則表達式替換為更新後的值
        IncrementNthChild = regex.Replace(selector, "tr:nth-child(" & number & ")")
    Else
        ' 如果匹配失敗，返回原始字符串
        IncrementNthChild = selector
    End If
End Function

Function NextChapterSelector(ChapterSelector As String) As String
    'Static ChapterSelector As String
    Dim selector As String
    Dim newSelector As String
    
    ' 檢查 ChapterSelector 是否為空
    If IsEmpty(ChapterSelector) Or ChapterSelector = "" Then
        NextChapterSelector = "" ' 如果為空，返回空字串
        Exit Function
    End If
    
    ' 設定當前的選擇器
    selector = ChapterSelector ' 範例: "#content > div:nth-child(6) > table > tbody > tr:nth-child(2) > td:nth-child(1) > a"
    
    ' 使用 IncrementNthChild 函式來更新選擇器
    newSelector = IncrementNthChild(selector)
    
    ' 更新靜態變數 ChapterSelector
    ChapterSelector = newSelector
    
    ' 返回新的選擇器
    NextChapterSelector = newSelector
End Function
Property Get Head_Edit_textbox() As SeleniumBasic.IWebElement
    If VBA.InStr(WD.url, "&action") Then '"&action=newchapter" 或 action=editchapter
        Set Head_Edit_textbox = WD.FindElementByCssSelector("#content > h2")
    End If
End Property
Property Get Title_Edit_textbox() As SeleniumBasic.IWebElement
    If VBA.InStr(WD.url, "&action=") Then '"&action=newchapter" 或 action=editchapter
        Set Title_Edit_textbox = WD.FindElementByCssSelector("#title")
    End If
End Property
Property Get Sequence_data_Edit_textbox() As SeleniumBasic.IWebElement
    If VBA.InStr(WD.url, "&action=") Then '"&action=newchapter" 或 action=editchapter
        Set Sequence_data_Edit_textbox = WD.FindElementByCssSelector("#sequence")
    End If
End Property
Property Get Textarea_data_Edit_textbox() As SeleniumBasic.IWebElement
    If VBA.InStr(WD.url, "&action=") Then '"&action=newchapter" 或 action=editchapter
        Set Textarea_data_Edit_textbox = WD.FindElementByCssSelector("#data")
    End If
End Property
Property Get description_Edit_textbox() As SeleniumBasic.IWebElement
    If VBA.InStr(WD.url, "&action=") Then '"&action=newchapter" 或 action=editchapter
        Set description_Edit_textbox = WD.FindElementByCssSelector("#description")
    End If
End Property
Property Get Commit_Edit_textbox() As SeleniumBasic.IWebElement
    If VBA.InStr(WD.url, "&action=") Then '"&action=newchapter" 或 action=editchapter
        Set Commit_Edit_textbox = WD.FindElementByCssSelector("#commit")
    End If
End Property
Sub 新頁面Auto_get_argument()
'    Rem 未完成
'    Rem 自動取得首頁、末頁及file ID num 3個引數
'    '文件第一段貼上首頁網址，如： https://ctext.org/library.pl?if=en&file=3918&page=1
'    Dim url As String, d As Document, p As Paragraph, iwe As SeleniumBasic.IWebElement
'    Set d = ActiveDocument
'    Set p = d.Paragraphs(1)
'    url = d.Range(p.Range.start, p.Range.End - 1).text
'    d.Range(p.Range.start, p.Range.End - 1).text = 1 '第一段為首頁頁碼=1
'    p.Range.InsertParagraphAfter
'    Set p = d.Paragraphs(2)
'    d.Range(p.Range.start, p.Range.End - 1).text = pageUBound '第2段為末頁頁碼
End Sub
Sub 新頁面Auto_action_newchapter()
    Dim d As Document, chapterNum As Integer, iwe As SeleniumBasic.IWebElement, newchapterUrl As String, title As String
    Set d = ActiveDocument
    '文件第4段輸入要開啟的書首頁面，如https://ctext.org/library.pl?if=gb&res=4925
    
    If IsWDInvalid Then
        If Not OpenChrome(VBA.Left(d.Paragraphs(4).Range.text, VBA.Len(d.Paragraphs(4).Range.text) - 1)) Then Exit Sub
    Else
        If Not Commit_Edit_textbox Is Nothing Then
            Commit_Edit_textbox.Click '送出
        End If
        WD.url = VBA.Left(d.Paragraphs(4).Range.text, VBA.Len(d.Paragraphs(4).Range.text) - 1)
    End If
    WD.SwitchTo.Window WD.CurrentWindowHandle
    '第5段輸入現在要新增單位的冊chapter序號，如第1冊則為2（冊序號+1）
    chapterNum = VBA.CInt(VBA.Left(d.Paragraphs(5).Range.text, VBA.Len(d.Paragraphs(5).Range.text) - 1))
    '"#content > div:nth-child(6) > table > tbody > tr:nth-child(2) > td:nth-child(1) > a"
    ChapterSelector = "#content > div:nth-child(6) > table > tbody > tr:nth-child(" & chapterNum & ") > td:nth-child(1) > a"
    '第6段為新增單位的頁面網址：
    'https://ctext.org/wiki.pl?if=gb&res=350225&action=newchapter
    newchapterUrl = VBA.Left(d.Paragraphs(6).Range.text, VBA.Len(d.Paragraphs(6).Range.text) - 1)
    
    '在書首資訊頁面中點擊相對應的chapter（冊）
    Set iwe = WD.FindElementByCssSelector(ChapterSelector)
    If iwe Is Nothing Then
        MsgBox "done!", vbInformation
        Exit Sub
    End If
    title = iwe.GetAttribute("text")
    iwe.Click
    Do Until VBA.InStr(WD.url, "&page")
        DoEvents
        Set iwe = WD.FindElementByCssSelector(ChapterSelector)
        If Not iwe Is Nothing Then
            iwe.Click
        End If
    Loop
    'Set iwe = WD.FindElementByCssSelector(Div_generic_IncludePathAndEndPageNum)
    d.Range(d.Paragraphs(1).Range.start, d.Paragraphs(1).Range.End - 1).text = 1 '首頁
    d.Range(d.Paragraphs(2).Range.start, d.Paragraphs(2).Range.End - 1).text = pageUBound '末頁
    'file ID
    'https://ctext.org/library.pl?if=gb&file=76754&page=1
    d.Range(d.Paragraphs(3).Range.start, d.Paragraphs(3).Range.End - 1).text = VBA.Trim(VBA.Mid(WD.url, VBA.InStr(WD.url, "&file=") + VBA.Len("&file="), VBA.InStr(WD.url, "&page=") - (VBA.InStr(WD.url, "&file=") + VBA.Len("&file="))))
    d.Activate
    WD.url = newchapterUrl
    WD.SwitchTo.Window WD.CurrentWindowHandle
    ActivateChrome
    Textarea_data_Edit_textbox.Click
    ActivateChrome
    VBA.DoEvents
    新頁面
    d.Undo
    If VBA.Len(Textarea_data_Edit_textbox.GetAttribute("value")) < 4 Then
        SetIWebElementValueProperty Textarea_data_Edit_textbox, GetClipboardText
    End If
    '輸入title值：
    Dim head As String
    head = Head_Edit_textbox.GetAttribute("outerText")
    head = VBA.Mid(head, 2, VBA.InStr(head, "》") - 2)
    
    SetIWebElementValueProperty Title_Edit_textbox, VBA.Replace(title, head, vbNullString)
    '輸入Sequence值：
    SetIWebElementValueProperty Sequence_data_Edit_textbox, VBA.CStr(chapterNum) & "0"
    '輸入修改摘要:
    SetIWebElementValueProperty description_Edit_textbox, description_Edit_textbox_新頁面 '"據《國學大師》或《Kanripo》所收本輔以末學自製於GitHub開源免費免安裝之TextForCtext軟件排版對應錄入；討論區及末學YouTube頻道有實境演示影片。感恩感恩　讚歎讚歎　南無阿彌陀佛"
'    SetIWebElementValueProperty Description_Edit_textbox, "據北京元引科技有限公司《元引科技引得數字人文資源平臺·中國歷代文獻》所收本輔以末學自製於GitHub開源免費免安裝之TextForCtext軟件排版對應錄入；討論區及末學YouTube頻道有實境演示影片。感恩感恩　讚歎讚歎　南無阿彌陀佛"
    'Commit_Edit_textbox.Click '送出
    
    Title_Edit_textbox.Click
    d.Range(d.Paragraphs(5).Range.start, d.Paragraphs(5).Range.End - 1).text = chapterNum + 1
    
    '按下【保存編輯】
    WD.FindElementByCssSelector("#commit").Click
    
    d.Application.Activate
    d.Application.windowState = wdWindowStateNormal
    d.Activate
    
End Sub


Rem 在序號欄位補零以調整章節及其次序用。因《御定佩文韻府》須調整章節單位長度而設（其原有290個單位故！！） https://ctext.org/wiki.pl?if=en&res=589161 20241214
Sub Add0toSequenceField()
    Dim w, url As String, iwe As SeleniumBasic.IWebElement, add0 As String
    If SeleniumOP.IsWDInvalid() Then
        If WD Is Nothing Then
            SeleniumOP.OpenChrome "https://ctext.org/"
        Else
            'WD.SwitchTo.Window SeleniumOP.WindowHandles()(SeleniumOP.WindowHandlesCount - 1)
'            WD.SwitchTo.Window SeleniumOP.WindowHandles()(0)
            WD.SwitchTo.Window SeleniumOP.WD.WindowHandles()(0)
        End If
    End If
    If Selection.Type = wdSelectionIP Then
        add0 = VBA.InputBox("輸入要補0的值，如要補兩個0，則輸入「00」。感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主")
    Else
        ResetSelectionAvoidSymbols
        add0 = Selection.text
    End If
    For Each w In WD.WindowHandles
        WD.SwitchTo.Window w
        url = WD.url
        If VBA.InStr(url, "https://ctext.org/wiki.pl") = 1 And VBA.InStr(url, "&chapter=") Then
            'Edit Link
            Set iwe = WD.FindElementByCssSelector("#content > h2 > span > a:nth-child(2)")
            iwe.Click
            'sequence Box
            Set iwe = WD.FindElementByCssSelector("#sequence")
            SeleniumOP.SetIWebElementValueProperty iwe, iwe.GetAttribute("value") & add0
            'Submit changes
            Set iwe = WD.FindElementByCssSelector("#commit")
            iwe.Click
            VBA.Interaction.DoEvents
            Do While VBA.InStr(WD.url, "&action=editchapter")
                SystemSetup.wait 0.3
            Loop
            WD.Close
        End If
    Next w
End Sub

Sub tempReplaceTxtforCtextEdit()
Dim a, d As Document, i As Integer, x As String
a = Array("{{（", "{{", "）}}", "}}", "（", "{{", "）", "}}", "○", VBA.ChrW(12295))
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
a = Array("{{（", "{{", "）}}", "}}", "（", "{{", "）", "}}", "○", VBA.ChrW(12295))
Set d = Documents.Add
d.Range.Paste
For i = 0 To UBound(a)
    d.Range.Find.Execute a(i), , , , , , , wdFindContinue, , a(i + 1), wdReplaceAll
    i = i + 1
Next i
d.Range.Cut
d.Application.windowState = wdWindowStateMinimize
d.Close wdDoNotSaveChanges
AppActivateDefaultBrowser
SendKeys "^v"
SendKeys "{tab}~"

End Sub



