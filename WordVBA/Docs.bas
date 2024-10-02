Attribute VB_Name = "Docs"
Option Explicit
Public d字表 As Document, x As New EventClassModule   '這才是所謂的建立"新"的類別模組--實際上是建立對它的參照.'原照線上說明乃Dim也.
'https://learn.microsoft.com/en-us/office/vba/word/concepts/objects-properties-methods/using-events-with-the-application-object-word
Public Sub Register_Event_Handler() '使自設物件類別模組有效的登錄程序.見「使用 Application 物件 (Application Object) 的事件」
    If x Is Nothing Or Not x.App Is word.Application Then
        SystemSetup.playSound 4
        Set x.App = word.Application '此即使新建的物件與Word.Application物件作上關聯
    End If
End Sub

Function 空白的新文件(Optional newDocVisible As Boolean = True) As Document '20210209
    Dim a As Document, flg As Boolean
    word.Application.ScreenUpdating = False
    If Documents.Count = 0 Then GoTo a:
    If ActiveDocument.Characters.Count = 1 And VBA.InStr(ActiveDocument.Name, "dotm") = 0 Then
        Set a = ActiveDocument
    ElseIf ActiveDocument.Characters.Count > 1 Then
        For Each a In Documents
            If (a.path = "" Or a.Characters.Count = 1) And VBA.InStr(a.Name, "dotm") = 0 Then
    '            a.Range.Paste'原來都有貼上，現在不要，單純找+開新文件就好
    '            a.Activate
    '            a.ActiveWindow.Activate
                flg = True
                Exit For
            End If
        Next a
        If flg = False Then GoTo a
    Else
a:     Set a = Documents.Add(Visible:=newDocVisible)
    End If
    Set 空白的新文件 = a
End Function


Sub 在本文件中尋找選取字串_迅速() '指定鍵:Alt+Ctrl+Down 2015/11/1
Static x As String
With Selection
    If .Type = wdSelectionNormal Then
        x = 文字處理.trimStrForSearch(.text, Selection)
        .Copy
        .Collapse wdCollapseEnd
        .Find.ClearAllFuzzyOptions
        .Find.ClearFormatting
    End If
    If x = "" Then Exit Sub
    If .Find.Execute(x, True, True, , , , True, wdFindContinue) = False Then MsgBox "沒了!", vbExclamation
End With
End Sub

Sub 隱藏字表()
'Dim 字表 As Boolean
''On Error Resume Next
'For Each d字表 In Application.Documents
'    'If d.Name = "字表7.2.doc" Then 字表 = True
'    If d字表.Name Like "字表*" Then 字表 = True: Exit For
'Next d字表
''If 字表 Then If InStr(ActiveDocument.Name, "字表7") = 0 Then Documents("字表7.2.doc").Windows(1).Visible = False
'If 字表 Then If InStr(ActiveDocument.Name, "字表") = 0 Then d字表.Windows(1).Visible = False
End Sub
Sub 插入札記註腳() '2002/11/15由圖書管理OLE.插入註腳應用而來
Dim h As word.Document, sectionct As Integer, rst As Recordset '2002/11/15
Dim d As Object
On Error GoTo OVER
'以下二行實可省略，今省之！2003/12/14
'Set d = GetObject("d:\千慮一得齋\書籍資料\圖書管理.mdb") '檢查圖書管理有無開啟!
'Set d = Nothing
''AppActivate "圖書管理"
If Not blog.myaccess.CurrentProject.AllForms("札記_類查詢").IsLoaded Then
     MsgBox "[札記_類查詢]表單沒開啟,不能作業!", vbCritical
    End
End If
Set h = ActiveDocument
Set rst = blog.myaccess.Forms("札記_類查詢").RecordsetClone
sectionct = h.Sections.Count
Dim a, a1, i As Integer, j As Integer
'h.Application.Visible = True '檢查用
a = Array("冊名", "書名", "篇名", "卷", "頁")
ReDim a1(0 To UBound(a)) As String
For i = 1 To sectionct '執行逐節插入註腳(每節為不重複的一筆記錄!)
    For j = 0 To UBound(a)
        '由頁與札記來比對相應的冊書篇名卷...等資料
        If Not VBA.Left(h.Sections(i).Range.Paragraphs(2).Range.text, Len(h.Sections(i).Range.Paragraphs(2).Range.text) - 1) Like "..." Then
            rst.FindFirst "頁 = " & VBA.Left(h.Sections(i).Range.Paragraphs(1).Range.text, Len(h.Sections(i).Range.Paragraphs(1).Range.text) - 1) _
                & "and 札記 like '" & "*" & VBA.Left(h.Sections(i).Range.Paragraphs(2).Range.text, _
                    Len(h.Sections(i).Range.Paragraphs(2).Range.text) - 1) & "*'"
    '            & "and 札記 like '" & "*" & Replace(h.Sections(i).Range.Paragraphs(2).Range.Text, vba.Chr(12), "") & "*'"
                'vba.Chr(12)暫不知是什麼(應是分節符號!),但會影響比對,故予取代為空字串 _
                因為最後一節沒有分節符號(vba.Chr(12))而是段落符號(vba.Chr(13),故若以Replace函數須分別處理 _
                為免麻煩,一律用Left函數不取最右方之字元即可(不管是vba.Chr(12)或vba.Chr(13))
        Else '札記為""時的處理
            rst.FindFirst "頁 = " & VBA.Left(h.Sections(i).Range.Paragraphs(1).Range.text, Len(h.Sections(i).Range.Paragraphs(1).Range.text) - 1) _
                & "and 札記 = """"" '在此,不用CSng型態轉換是可以的! _
                因為頁有小數點,故在作頁比對時,不能用Words物件(會將小數點之數字分開算成不同的Word),如果頁沒有小數點的話就可以了! _
                減一,同札記,是為了剔除最右方的vba.Chr(10)(換行符號、段落符號)
        End If
        a1(j) = blog.myaccess.Nz(rst(a(j)), 0) '卷次會有Null值!
    Next j
    h.Footnotes.Add h.Sections(i).Range.words(h.Sections(i).Range.words.Count), _
        , "《" & a1(LBound(a1)) & "》，《" & a1(LBound(a1) + 1) & "》，〈" & a1(LBound(a1) + 2) & "〉，卷" _
        & a1(LBound(a1) + 3) & "，頁" & a1(LBound(a1) + 4) & "。"
Next i
Set rst = Nothing: Set h = Nothing
MsgBox "完成" & sectionct & "項註腳插入! "
End
OVER:
    MsgBox "圖書管理沒開啟,不能作業!", vbCritical
End Sub

Sub 文件內容比對_校勘用() '_以字元為單位() '2004/10/20:指定鍵(快速鍵):Ctrl+Alt+Return
Dim s1 As Range, s2 As Range, d1 As Document, d2 As Document, j As Long, k As Long
'Static i As Long, DN As String, MarkTimes As Byte
Dim i As Long, dn As String, MarkTimes As Byte
Select Case Documents.Count '先檢查文件數
    Case Is > 2
        MsgBox "只能一次核對兩分文件！請將不必要的文件關閉後，再操作一次！", vbExclamation: Exit Sub
    Case Is = 0
        Exit Sub
    Case Is = 1
        MsgBox "目前只有一份文件！請將要與之校對的文件打開，然後再操作一次！", vbExclamation: Exit Sub
End Select
'再檢查視窗數.
If Windows.Count > 2 Then MsgBox "請將多餘的視窗關後，再操作一次！", vbExclamation: Exit Sub
'要先置入插入點,以插入點位置開始往後處理!
Set d1 = Documents(1)
Set d2 = Documents(2)
Windows.Arrange 'wdTiled'排列視窗
If MsgBox("是否要清除符號?", vbQuestion + vbOKCancel) = vbOK Then
    d1.Activate: 清除所有符號
    d2.Activate: 清除所有符號
End If
If dn = "" Then dn = d1.Name
If Not d1.Name Like dn And Not d1.Name Like dn Then i = 0: dn = d1.Name
If Selection.start + 1 = ActiveDocument.Content.End Then Selection.HomeKey wdStory, wdMove '如果插入點為文件末時...
If i = 0 Then i = Selection.start 'ActiveDocument.Range.Start
If d1.Characters.Count >= d2.Characters.Count Then
    j = d1.Characters.Count
    k = d2.Characters.Count
Else
    j = d2.Characters.Count
    k = d1.Characters.Count
End If
For i = i + 1 To j
    StatusBar = i
    If i > k Then MsgBox "比對完畢!": End ': Exit For  '到了較少字文件的末端時
    Set s1 = d1.Characters(i): Set s2 = d2.Characters(i)
'    If Asc(S1) = 2 Or Asc(S1) = 5 Or Asc(S2) = 5 Or Asc(S1) = 2 Then Stop '註腳(2)或註解(5)時
'    If AscW(S1) = 63 Or AscW(S2) = 63 Then Stop '半形問號時
    If Not s1 Like s2 Then
        MarkTimes = MarkTimes + 1
        If MarkTimes > 20 Then MsgBox "恐有整段不同的情況或缺漏刪省之異文，請自行校對，再將插入點置於適當的比對初始位置繼續執行即可！", vbExclamation: Exit Sub
        s1.Select ': D1.Windows(1).ScrollIntoView S1, True 'Selection.Range, True
'        Options.DefaultHighlightColorIndex = wdBrightGreen
        s1.HighlightColorIndex = wdBrightGreen '標示為螢光綠
        s2.Select ': D1.Windows(1).ScrollIntoView S2, True 'Selection.Range, True
        ActiveDocument.Windows(1).ScrollIntoView ActiveDocument.Characters(i), True
'        j = MsgBox("請校對！" & vbCr & vbCr & _
            "【 " & S1 & " ←→ " & S2 & " 】" & vbCr & vbCr & _
            "要重來請按﹝取消﹞鍵!", vbExclamation + vbOKCancel)
            s2.HighlightColorIndex = wdBrightGreen '標示為螢光綠
'        If j = vbOK Then
'            ActiveWindow.Next.Activate
''            ActiveDocument.Windows(1).ScrollIntoView ActiveDocument.Characters(i), True
            ActiveWindow.ScrollIntoView ActiveDocument.Characters(i), True
''            Dim x As Long'自動計時瀏覽用
''            For x = 1 To 50000000
''            Next
'            Exit For
'        Else
'            End
'        End If
    End If
Next i
MsgBox "比對完畢!"
End Sub

Sub 文件內容比對_校勘用_unfinished() '_以字元為單位() '2004/10/20:指定鍵(快速鍵):Ctrl+Alt+Return
Dim s1 As Range, s2 As Range, Dw1 As Document, Dw2 As Document, j As Long, k As Long
Static i As Long, DwN As String
Select Case Documents.Count '先檢查文件數
    Case Is > 2
        MsgBox "只能一次核對兩分文件！請將不必要的文件關閉後，再操作一次！", vbExclamation: Exit Sub
    Case Is = 0
        Exit Sub
    Case Is = 1
        MsgBox "目前只有一份文件！請將要與之校對的文件打開，然後再操作一次！", vbExclamation: Exit Sub
End Select
'再檢查視窗數.
If Windows.Count > 2 Then MsgBox "請將多餘的視窗關後，再操作一次！", vbExclamation: Exit Sub
'要先置入插入點,以插入點位置開始往後處理!
Set Dw1 = Windows(1)
Set Dw2 = Windows(2)
Windows.Arrange 'wdTiled'排列視窗
If MsgBox("是否要清除符號?", vbQuestion + vbOKCancel) = vbOK Then
    Dw1.Activate: 清除所有符號
    Dw2.Activate: 清除所有符號
End If
If DwN = "" Then DwN = Dw1.Name
If Not Dw1.Name Like DwN And Not Dw1.Name Like DwN Then i = 0: DwN = Dw1.Name
If Dw1.Selection.start + 1 = ActiveDocument.Content.End Then Selection.HomeKey wdStory, wdMove '如果插入點為文件末時...
If i = 0 Then i = Dw1.Selection.start 'ActiveDocument.Range.Start
If Dw1.Characters.Count >= Dw2.Characters.Count Then
    j = Dw1.Characters.Count
    k = Dw2.Characters.Count
Else
    j = Dw2.Characters.Count
    k = Dw1.Characters.Count
End If
For i = i + 1 To j
    If i > k Then MsgBox "比對完畢!": End ': Exit For  '到了較少字文件的末端時
    Set s1 = Dw1.Characters(i): Set s2 = Dw2.Characters(i)
'    If Asc(S1) = 2 Or Asc(S1) = 5 Or Asc(S2) = 5 Or Asc(S1) = 2 Then Stop '註腳(2)或註解(5)時
'    If AscW(S1) = 63 Or AscW(S2) = 63 Then Stop '半形問號時
    If Not s1 Like s2 Then
        s1.Select ': D1.Windows(1).ScrollIntoView S1, True 'Selection.Range, True
'        Options.DefaultHighlightColorIndex = wdBrightGreen
        s1.HighlightColorIndex = wdBrightGreen '標示為螢光綠
        s2.Select ': D1.Windows(1).ScrollIntoView S2, True 'Selection.Range, True
        ActiveDocument.Windows(1).ScrollIntoView ActiveDocument.Characters(i), True
'        j = MsgBox("請校對！" & vbCr & vbCr & _
            "【 " & S1 & " ←→ " & S2 & " 】" & vbCr & vbCr & _
            "要重來請按﹝取消﹞鍵!", vbExclamation + vbOKCancel)
            s2.HighlightColorIndex = wdBrightGreen '標示為螢光綠
'        If j = vbOK Then
'            ActiveWindow.Next.Activate
''            ActiveDocument.Windows(1).ScrollIntoView ActiveDocument.Characters(i), True
            ActiveWindow.ScrollIntoView ActiveDocument.Characters(i), True
''            Dim x As Long'自動計時瀏覽用
''            For x = 1 To 50000000
''            Next
'            Exit For
'        Else
'            End
'        End If
    End If
Next i
MsgBox "比對完畢!"
End Sub

Sub 清除所有符號() '由圖書管理symbles模組清除標點符號改編'包括註腳、數字
'Dim F, a As String, i As Integer
Dim f, i As Integer, ur As UndoRecord
SystemSetup.stopUndo ur, "清除所有符號"
f = Array("·", "‧", "。", "」", VBA.Chr(-24152), "：", "，", "；", _
    "、", "「", ".", VBA.Chr(34), ":", ",", ";", _
    "……", "...", "．", "【", "】", " ", "《", "》", "〈", "〉", "？" _
    , "！", "﹝", "﹞", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0" _
    , "『", "』", VBA.Chr(13), VBA.ChrW(9312), VBA.ChrW(9313), VBA.ChrW(9314), VBA.ChrW(9315), VBA.ChrW(9316) _
    , VBA.ChrW(9317), VBA.ChrW(9318), VBA.ChrW(9319), VBA.ChrW(9320), VBA.ChrW(9321), VBA.ChrW(9322), VBA.ChrW(9323) _
    , VBA.ChrW(9324), VBA.ChrW(9325), VBA.ChrW(9326), VBA.ChrW(9327), VBA.ChrW(9328), VBA.ChrW(9329), VBA.ChrW(9330) _
    , VBA.ChrW(9331), VBA.ChrW(8221), """") '先設定標點符號陣列以備用
    '全形圓括弧暫不取代！
    'a = ActiveDocument.Content
'    Set a = ActiveDocument.Range.FormattedText '包含格式化的資訊
    For i = 0 To UBound(f)
        'a = Replace(a, F(i), "")
        ActiveDocument.Range.Find.Execute f(i), True, , , , , , wdFindContinue, True, "", wdReplaceAll
    Next
    'ActiveDocument.Content = a
SystemSetup.contiUndo ur
End Sub
Sub 注腳符號() '注釋符號、註釋符號、註腳符號
Dim i As Integer
For i = 9312 To 9331
    With Selection.Range.Find
        .Replacement.font.Name = "Arial Unicode MS"
        .Execute VBA.ChrW(i), , , , , , , wdFindContinue, , VBA.ChrW(i), wdReplaceAll
    End With
Next i
'Dim i As Long, w As Characters, wc As Long
'Set w = ActiveDocument.Range.Characters
'wc = ActiveDocument.Range.Characters.Count
'For i = 1 To wc
'    If AscW(w(i)) > 9311 And AscW(w(i)) < 9332 Then
'        w(i).Font.Name = "Arial Unicode MS"
'    End If
'Next i
End Sub

Sub 貼上引文() '將已複製到剪貼簿的內容貼成引文
    Dim s As Long, e As Long, r  As Range
    If Selection.Type = wdSelectionNormal And VBA.Right(Selection, 1) Like VBA.Chr(13) Then _
                Selection.MoveLeft wdCharacter, 1, wdExtend '不要包含分段符號!
    If Selection.Style <> "引文" Then Selection.Style = "引文" '如果不是引文樣式時,則改成引文樣式
    s = Selection.start '記下起始位置
    Selection.PasteSpecial , , , , wdPasteText '貼上純文字
    e = Selection.End '記下貼上後的結束位置
    Selection.SetRange s, e
    Set r = Selection.Range
    With r
        r.Find.Execute VBA.Chr(13), , , , , , , wdFindStop, , VBA.Chr(11), wdReplaceAll '將分行符號改成手動分行符號
    End With
    r.Footnotes.Add r '插入註腳!
End Sub
Sub 貼上純文字() 'shift+insert 2016/7/20
    Dim hl, s As Long, r As Range
    On Error GoTo ErrHandler
    hl = Selection.Range.HighlightColorIndex
    
    s = Selection.start
    Set r = Selection.Range
'    '如果有選取則清除
    If Selection.Flags <> 24 And Selection.Flags <> 25 Or Selection.Flags = 9 Then
        If s < Selection.End Then Selection.text = vbNullString
    End If
    'Selection.PasteSpecial , , , , wdPasteText '貼上純文字
    Selection.PasteAndFormat (wdFormatPlainText)
    r.SetRange s, Selection.End
    If hl <> 9999999 Then r.HighlightColorIndex = hl '9999999 is multi-color 多重高亮色彩則無顯示9999999（7位數9）的值
    Exit Sub
ErrHandler:
    Select Case Err.Number
        Case 5342 '指定的資料類型無法取得。
            
        Case Else
            MsgBox Err.Number & Err.Description
    End Select
End Sub
Sub 貼上簡化字文本轉正()
    Dim rng As Range, ur As UndoRecord
    SystemSetup.stopUndo ur, "貼上簡化字文本轉正"
    Set rng = Selection.Range
    rng.PasteAndFormat (wdFormatPlainText)
    標點符號置換 rng: 清除半形空格 rng: 半形括號轉全形 rng
    If MsgBox("是否簡轉正？", vbOKCancel) = vbOK Then
        'rng.Select
        rng.TCSCConverter wdTCSCConverterDirectionAuto
        'Selection.Range.TCSCConverter wdTCSCConverterDirectionAuto
    End If
    SystemSetup.contiUndo ur
    End Sub
Sub 簡化字文本轉正()
    Dim rng As Range, ur As UndoRecord
    SystemSetup.stopUndo ur, "簡化字文本轉正"
    Set rng = Selection.Range
    標點符號置換 rng: 清除半形空格 rng: 半形括號轉全形 rng
    rng.TCSCConverter wdTCSCConverterDirectionAuto
    SystemSetup.contiUndo ur
    SystemSetup.playSound 1
End Sub
Function 標點符號置換(Optional rng As Range)
    Dim ay, i As Integer
    ay = Array(VBA.ChrW(8220), "「", VBA.ChrW(8221), "」", VBA.ChrW(-431), "、", VBA.ChrW(-432), "，" _
        , VBA.ChrW(58), "：", VBA.ChrW(8216), "『", VBA.ChrW(8217), "』", _
        VBA.ChrW(-428), "；", "·", "‧", ",", "，", ";", "；" _
        , "?", "？", ":", "：", "﹕", "：")
    For i = 0 To UBound(ay)
        rng.Find.Execute ay(i), , , , , , , wdFindContinue, , ay(i + 1), wdReplaceAll
        i = i + 1
    Next i
End Function
Function 清除半形空格(Optional rng As Range)
rng.Find.Execute " ", , , , , , , wdFindContinue, , "", wdReplaceAll
End Function
Function 半形括號轉全形(Optional rng As Range)
rng.Find.Execute "(", , , , , , , wdFindContinue, , "（", wdReplaceAll
rng.Find.Execute ")", , , , , , , wdFindContinue, , "）", wdReplaceAll
End Function


Sub 一字一段()
With Selection
    .HomeKey wdStory
    Do Until .End = .Document.Range.End - 1
        .MoveRight
        .TypeText VBA.Chr(13)
    Loop
End With
End Sub
Sub OCR表格處理()
Dim a, b, i As Byte
a = Array("┴", VBA.Chr(13) & "│", " ", "─", "┼", "├", "┤", "┐", "┌", "└", "┘", "┬", "│", VBA.Chr(9) & VBA.Chr(13), VBA.Chr(13) & VBA.Chr(13), VBA.Chr(13) & VBA.Chr(9))
b = Array("", VBA.Chr(13), "", "", "", "", "", "", "", "", "", "", VBA.Chr(9), VBA.Chr(13), VBA.Chr(13), VBA.Chr(13))
With ActiveDocument
    If .path = "" Then
        For i = 0 To UBound(a) - 1
            With .Range.Find
                .text = a(i)
                With .Replacement
                    .text = b(i)
                End With
                .Execute , , , , , , , wdFindContinue, , , wdReplaceAll
            End With
        Next
    End If
End With
End Sub

Sub 在同目錄下尋找符合關鍵字之文本() '2009/8/4'Alt+shift+F3
Dim x As String, d As Document, i As Integer
Set d = ActiveDocument
x = Selection
With word.Application.FileSearch
    .NewSearch
    If d.path = "" Then
        .LookIn = "D:\千慮一得齋\論文資料夾\博士論文\論文稿"
    Else
        .LookIn = d.path
    End If
    .SearchSubFolders = True
    '.FileName = "Run"
    .MatchTextExactly = True
    .FileType = msoFileTypeAllFiles
    .TextOrProperty = x
     If .Execute() > 0 Then
        MsgBox "There were " & .FoundFiles.Count & _
            " file(s) found."
        For i = 1 To .FoundFiles.Count
            x = .FoundFiles(i)
            x = VBA.Mid(x, InStrRev(x, "\") + 1)
            If x <> d.Name Then
                If MsgBox(x, vbOKCancel) = vbOK Then
                     Documents.Open .FoundFiles(i)
                End If
            End If
        Next i
    Else
        MsgBox "There were no files found."
    End If

End With
End Sub

Sub 清除所有註解()
Dim e
With ActiveDocument
    For Each e In .Comments
        e.Range.Select
        e.Delete
    Next e
End With
End Sub

Sub 清除多餘不必要的分段()
Dim p As Paragraph, rng As Range, ur As UndoRecord
SystemSetup.stopUndo ur, "清除多餘不必要的分段"
Set rng = ActiveDocument.Range
For Each p In ActiveDocument.Paragraphs
    If p.Range.Characters.Count > 2 Then
        If Not p.Range.Characters(p.Range.Characters.Count - 1) Like "[》」』。（" & VBA.ChrW(-197) & "0-9a-zA-Z]" Then
            If p.Range.End < ActiveDocument.Range.End - 1 Then
'                p.Range.Characters(p.Range.Characters.Count).Select
                p.Range.Characters(p.Range.Characters.Count).Delete
            End If
        End If
    End If
Next
SystemSetup.contiUndo ur
SystemSetup.playSound 2
End Sub

Sub 開新視窗() '快速鍵:alt+shift+w-原為OLE至備忘欄()指定鍵  '2011/6/23''2012/5/20 2003不能設定Alt+w 原設於"字形轉換_華康儷粗黑"
Dim l As Long, s As Long, YwdInFootnoteEndnotePane
    YwdInFootnoteEndnotePane = Selection.Information(wdInFootnoteEndnotePane) '記下開新視窗前註腳窗格狀態
    l = Selection.End '.Information(wdActiveEndPageNumber)
    s = Selection.start '記下原位置
    If CommandBars("web").Visible Then CommandBars("web").Visible = False
    NewWindow
    'ActiveWindow.Document.Range.Characters(l).Select
    If YwdInFootnoteEndnotePane Then '如果在註腳窗格中
        ActiveWindow.View.SplitSpecial = wdPaneFootnotes '2011/8/13
    End If
    Selection.End = l ' 'Selection.GoTo wdGoToObject, wdGoToAbsolute, l
    Selection.start = s '到原位置
End Sub

Sub 文件引導模式切換() ' Alt+M 2011/6/26
' 巨集7 巨集
' 巨集錄製於 2011/6/26，錄製者 Oscar Sun
'
Dim s As Long, e As Long, YwdInFootnoteEndnotePane
YwdInFootnoteEndnotePane = Selection.Information(wdInFootnoteEndnotePane) '記下開新視窗前註腳窗格狀態
If YwdInFootnoteEndnotePane Then '如果在註腳窗格中
    ActiveWindow.ActivePane.Close '2011/9/24
End If
s = Selection.start
e = Selection.End
If ActiveWindow.DocumentMap Then
    ActiveWindow.DocumentMap = False
ElseIf ActiveWindow.DocumentMap = False Then
    ActiveWindow.DocumentMap = True
End If
Selection.start = s
Selection.End = e
End Sub

Sub 目錄字形更正() '2011/8/23'蓋若標題改用非標題大小之字形,會bug
Dim e As Long
With ActiveDocument
    e = .Range.End
'    With .ActiveWindow.Selection.Find
''        .Font.Size>12
'        .Execute
'
    Do Until Selection.End = e - 1
        Selection.MoveRight
            If Selection.Next.font.Size > 12 Then '目錄預設為12字形
                Selection.MoveRight
                Do Until Selection.Next.font.Size = 12
                    Selection.MoveRight , , wdExtend
                Loop
                If MsgBox("是否要縮小為10號字?", vbQuestion + vbOKCancel) = vbOK Then Selection.font.Size = 10
        End If
    Loop
End With
End Sub

Sub 清除螢光黃() '2012/2/3 待測試
Dim h
h = wdYellow
'無為0,淡藍為3,.......
'If h = "" Then h = Selection.Range.HighlightColorIndex
Do Until Selection.Range.HighlightColorIndex <> h

    Exit Do 'Selection.Range.HighlightColorIndex = wdAuto
Loop
'MsgBox "完成!", vbInformation
End Sub

Sub 關閉其餘的文件()
'Ctrl+Alt+W
Dim d As Document, dn As String
dn = ActiveDocument.FullName
For Each d In Documents
    If d.FullName <> dn Then d.Close wdDoNotSaveChanges
Next
End Sub

Sub 在本文件中尋找選取字串() 'Ctrl+Alt+Down 2020/10/4改用 Ctrl+Shift+PageDown
    'CheckSavedNoClear
    If ActiveDocument.path <> "" Then If ActiveDocument.Saved = False Then ActiveDocument.Save
    Dim ins(4) As Long, MnText As String, FnText As String, FdText As String, st As Long, ed As Long
    On Error GoTo errHH
    With Selection '快速鍵：Alt+Ctrl+Down
    'If Not .Text Like "" Then '快速鍵：Alt+Ctrl+Down
    If .Type = wdSelectionIP Then MsgBox "請選取想要尋找之文字", vbExclamation: Exit Sub
    If .Type = wdSelectionNormal Then ' <> wdNoSelection OR wdSelectionIP Then '不為插入點
    '    If InStr(ActiveDocument.Content, .Text) = InStrRev(ActiveDocument.Content, .Text) Then MsgBox "本文只有此處!", vbInformation: Exit Sub
        FdText = 文字處理.trimStrForSearch(.text, Selection)
        st = .start: ed = .End
        .Collapse wdCollapseEnd
        MnText = .Document.StoryRanges(wdMainTextStory) '變數化處理較快2003/4/8
    '    MnText = ActiveDocument.Range '2010/2/5
        ins(1) = InStr(MnText, FdText)
        ins(2) = InStrRev(MnText, FdText)
        
         '有註腳才檢查2003/4/3
        If .Document.Footnotes.Count > 0 Then
            FnText = .Document.StoryRanges(wdFootnotesStory)
            ins(3) = InStr(FnText, FdText)
            ins(4) = InStrRev(FnText, FdText)
        End If
        
        If ins(1) = ins(2) And ins(3) = ins(4) Then
            If ins(1) <> 0 And .Information(wdInFootnote) Then
                If MsgBox("註腳只有此處!　　正文還有.." & vbCr & vbCr & _
                    "要尋找嗎?", vbInformation + vbOKCancel, "尋找：「" & FdText & "」") = vbCancel Then
                    Exit Sub
                Else
    '                FdText = .Text
    '                .Document.ActiveWindow.ActivePane.Previous.Activate
    '                .Document.Select '此法可將焦點轉移到正文
                    With .Document.Range.Find
                        .ClearFormatting
                        .ClearAllFuzzyOptions
                        .text = FdText
                        .Execute
                        .Parent.Select
                    End With
                End If
            ElseIf ins(3) <> 0 And Not .Information(wdInFootnote) Then
                If MsgBox("正文只有此處!　　註腳還有.." & vbCr & vbCr & _
                    "要繼續尋找嗎？", vbInformation + vbOKCancel, "尋找：「" & _
                        FdText & "」") = vbCancel Then
                    Exit Sub
                Else
    '                FdText = .Text
                    With .Document.ActiveWindow
                        If .Panes.Count = 1 Then
                            '開啟註腳視窗
                            If .View.Type = wdNormalView Then _
                               .View.SplitSpecial = wdPaneFootnotes
                        Else
                            .ActivePane.Next.Activate
                        End If
                        With .ActivePane.Selection.Find
                            .ClearFormatting
                            .ClearAllFuzzyOptions
                            .text = FdText
                            .Forward = True
                            .Wrap = wdFindContinue '要有這行才能正確尋找
                            .Execute
                        End With
    '                    .ScrollIntoView .ActivePane.Selection, True
    '                    .ActivePane.SmallScroll
                    End With
                End If
            ElseIf ins(1) = ins(2) And ins(3) <> 0 And ins(1) = 0 And ins(3) = ins(4) Then
                MsgBox "本文只有此處!  正文無!", vbInformation, "尋找：「" & FdText & "」": Exit Sub
            ElseIf ins(1) = ins(2) And ins(1) <> 0 And ins(3) = 0 And ins(3) = ins(4) Then
                MsgBox "本文只有此處!  註腳無!", vbExclamation, "尋找：「" & FdText & "」"
                .start = st
                .End = ed
                Exit Sub
            End If
        Else
    '        If ins(1) <> 0 Then
    '            ins(1) = wdMainTextStory
    '        Else
    '            ins(1) = wdFootnotesStory
    '        End If
            If ins(3) = ins(4) And .Information(wdInFootnote) = True Then _
                MsgBox "本文只有註腳此處有!", vbInformation, "尋找：「" & FdText & "」": Exit Sub
    '        With .Document.StoryRanges(ins(1)).Find
            If ins(1) < ins(2) Then .HomeKey wdStory 'ins(2)是文件本文最後出現的位置故 20241002
            With .Find
                .ClearFormatting
                .ClearAllFuzzyOptions
                .Replacement.ClearFormatting '這也要清除才行
                .Forward = True
                .Wrap = wdFindAsk
                .MatchCase = True
                .text = FdText '.Parent.Text
                .Execute
    '            .Parent.Select'用Range物件得用此方法才能改變選取
            End With
        End If
    End If
    End With
    Exit Sub
errHH:
    Select Case Err.Number
        Case 7 '記憶體不足
            ActiveDocument.ActiveWindow.Selection.Find.Execute Selection.text
        Case Else
            MsgBox Err.Number & Err.Description
            Resume
    End Select
End Sub


Sub 書籤_以選取文字作為書籤() 'ALT+SHIFT+B

' 巨集錄製於 2015/9/20，錄製者 王觀如
    With ActiveDocument.bookmarks
        .Add Range:=Selection.Range, Name:=Replace(Selection.text, VBA.Chr(13), "")
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
End Sub
Sub 小小輸入法詞庫cj5_ftzk_3字以上詞彙隱藏()
If Not ActiveDocument.Name = "cj5-ftzk.txt" Then Exit Sub
Dim d As Document, flg As Boolean, s As Byte, prngTxt As String
Set d = ActiveDocument
Dim p As Paragraph
Const x As String = "ahysy 易於"
For Each p In d.Paragraphs
    If InStr(p.Range, x) Then flg = True
    If Not flg Then
        p.Range.font.Hidden = True
    Else
        prngTxt = p.Range.text
        s = InStr(prngTxt, " ")
        If VBA.Len(VBA.Mid(prngTxt, s + 1)) < 5 Then
            p.Range.font.Hidden = True
        End If
    End If
Next p
d.ActiveWindow.View.ShowHiddenText = False
Beep
End Sub
Sub 小小輸入法詞庫cj5_ftzk_3字以上詞彙刪存至cj5_ftzk_other()
If Not ActiveDocument.Name = "cj5-ftzk.txt" Then Exit Sub
Dim d As Document
小小輸入法詞庫cj5_ftzk_3字以上詞彙隱藏
Set d = ActiveDocument
Dim p As Paragraph, s As Byte, prngTxt As String, msgResult As Integer
For Each p In d.Paragraphs
    If Not p.Range.font.Hidden Then
        prngTxt = p.Range.text
        s = InStr(prngTxt, VBA.Chr(32))
        If VBA.Len(VBA.Mid(prngTxt, s + 1)) > 4 Then
            DoEvents
            p.Range.Select
            ActiveWindow.ScrollIntoView p.Range
            msgResult = MsgBox("是否移到other去？", vbYesNoCancel)
            Select Case msgResult
                Case vbYes
                    小小輸入法詞庫cj5_ftzk刪存至cj5_ftzk_other
                Case vbNo
                    Debug.Print p.Next.Range.text
                    Exit Sub
            End Select
        End If
    End If
Next p
End Sub
Sub 小小輸入法詞庫cj5_ftzk刪存至cj5_ftzk_other()
'Dim p As Paragraph'Alt+1 Alt+2 ctrl+q
Dim rng As Range, p As Paragraph, pc As Integer, pSelRng As Range, x As String
If ActiveDocument.Name <> "cj5-ftzk.txt" Then Exit Sub
word.Application.ScreenUpdating = False
If Selection.Type = wdSelectionIP Then
    If Not Selection.Range.TextRetrievalMode.IncludeHiddenText Then
        Selection.Paragraphs(1).Range.Select
    End If
Else
    ActiveWindow.ActivePane.View.ShowAll = Not ActiveWindow.ActivePane.View. _
        ShowAll
End If
If Not Selection.Range.TextRetrievalMode.IncludeHiddenText Then
    x = Selection.text
    Selection.Delete
    'Selection.Cut
        GoSub subP
'        ActiveWindow.ActivePane.View.ShowAll = Not ActiveWindow.ActivePane.View. _
        ShowAll
        word.Application.ScreenUpdating = True
        Exit Sub
Else
    Set pSelRng = Selection.Range
    For Each p In pSelRng.Paragraphs
        If Not p.Range.font.Hidden Then 'prepare to delete
            'p.Range.Cut
            x = p.Range.text
            p.Range.Delete
            GoSub subP
        End If
    Next p
End If
ActiveWindow.ActivePane.View.ShowAll = Not ActiveWindow.ActivePane.View. _
        ShowAll
word.Application.ScreenUpdating = True
Exit Sub
subP:
    With Documents("cj5-ftzk-other.txt").Range
'        If .Paragraphs(.Paragraphs.Count).Range <> vba.Chr(13) Then .InsertParagraphAfter
'        .Paragraphs(.Paragraphs.Count).Range.Paste
'        .Document.ActiveWindow.ScrollIntoView .Paragraphs(.Paragraphs.Count).Range
        .InsertAfter x
        .Document.ActiveWindow.ScrollIntoView .Parent.Range(.End - 1, .End)
    End With
    Return
End Sub
Function 樣式取代()
Const styleSrc As String = "純文字"
Const styleDest As String = "易經原文"
Dim d As Document, p As Paragraph
For Each p In d.Paragraphs
    If p.Style = styleSrc Then p.Style = styleDest
Next p
End Function

Function 樣式add_沛榮按等樣式()
Const styleHprAn As String = "沛榮按"
Const styleShengDiao As String = "聲調"
Dim d As Document, myStyle  As Style, doNotAdd As Boolean
Set d = ActiveDocument
For Each myStyle In d.Styles
    If myStyle = styleHprAn Then doNotAdd = True: Exit For
Next myStyle
If Not doNotAdd Then
    Set myStyle = d.Styles.Add(styleHprAn, wdStyleTypeCharacter)
    With myStyle
        With .font
            .NameFarEast = "標楷體"
            .Color = 12611584
            .Size = 11
            .Spacing = -0.4
        End With
        .Visibility = True
        .Priority = 1
        .UnhideWhenUsed = True
    End With
    doNotAdd = False
End If
For Each myStyle In d.Styles
    If myStyle = styleShengDiao Then doNotAdd = True: Exit For
Next myStyle
If Not doNotAdd Then
    Set myStyle = d.Styles.Add(styleShengDiao, wdStyleTypeCharacter) 'https://docs.microsoft.com/zh-tw/office/vba/api/word.wdstyletype
    With myStyle
        .BaseStyle = d.Styles(styleHprAn)
        With .font
            .NameFarEast = "標楷體"
            .Name = "標楷體"
            .Position = 3
        End With
        .Visibility = True
        .Priority = 1
        .UnhideWhenUsed = True
    End With
    doNotAdd = False
End If
End Function

Sub closeDocs關閉未儲存的文件檔案() 'Alt+w
Dim d As Document
word.Application.ScreenUpdating = False
For Each d In Documents
    If d.path = "" Then d.ActiveWindow.Visible = False
Next d
For Each d In Documents
    If d.path = "" Then d.Close wdDoNotSaveChanges
Next d
word.Application.ScreenUpdating = True
If word.Windows.Count = 0 Then
    word.Documents.Add
    ActiveWindow.Visible = True
End If
End Sub
Sub DocBackgroundFillColor() '頁面色彩
    ActiveDocument.ActiveWindow.View.Type = wdPrintView 'https://docs.microsoft.com/en-us/office/vba/api/word.document.background
'    ActiveDocument.Background.Fill.ForeColor.RGB = RGB(192, 192, 192) 'RGB(146, 208, 80)
'    ActiveDocument.Background.Fill.Visible = True
'    ActiveDocument.Background.Fill.Solid
    ActiveDocument.Background.Fill.Visible = True
    ActiveDocument.Background.Fill.ForeColor.RGB = RGB(0, 102, 102)
    ActiveDocument.Background.Fill.BackColor.RGB = RGB(0, 102, 102)
    ActiveDocument.Background.Fill.Solid
End Sub

Sub 內文前空二格() 'Alt+n
    With Selection.ParagraphFormat
        .Style = "內文"
        .CharacterUnitFirstLineIndent = 2
    End With
End Sub
Sub 易學關鍵字()
    ' Alt + `
    mark易學關鍵字
End Sub
Rem 執行完成才傳回true，否則為false
Function mark易學關鍵字(Optional pasteRange As Range, Optional doNotMark As Boolean) As Boolean
    ' Alt + `
    Dim searchedTerm, e, ur As UndoRecord, d As Document, clipBTxt As String, flgPaste As Boolean, dSource As Document
    Dim strAutoCorrection, endDocOld As Long, rng As Range, returnVaule As Boolean
    Dim punc As New punctuation
    SystemSetup.playSound 0.484
    strAutoCorrection = Array("，〉", "〉，", "〈、", "〈", "〈。", "〈", "。〉", "〉", "〈：", "〈", "：〉", "〉", "〈，", "〈", "、〉", "〉")
    If InStr(ActiveDocument.path, "易學雜著文本") = 0 Then
        If MsgBox("目前文件為" + ActiveDocument.Name + "是否繼續？", vbExclamation + vbOKCancel) = vbCancel Then Exit Function
    End If
    Set dSource = ActiveDocument: If Not dSource.Saved Then dSource.Save
    Set rng = dSource.Range
    With rng.Find
        .ClearAllFuzzyOptions
        .ClearFormatting
        '255是上限，但也可能包含了標點斷句異文而導致文本有異不能比對，還不如縮減到可以識別的長度即可
        'If .Execute(VBA.Trim(VBA.Left(SystemSetup.GetClipboard, 255)), , , , , , True, wdFindContinue) Then
        If .Execute(VBA.Trim(VBA.Left(SystemSetup.GetClipboard, 25)), , , , , , True, wdFindContinue) Then
            rng.Select
            rng.Document.ActiveWindow.ScrollIntoView rng, True
            Exit Function
        End If
    End With
    'If Documents.Count = 0 Then Documents.Add
    If Documents.Count = 0 Then Set d = Docs.空白的新文件(True)
    If ClipBoardOp.Is_ClipboardContainCtext_Note_InlinecommentColor Then
        Set d = Docs.空白的新文件(False)
        中國哲學書電子化計劃.只保留正文注文_且注文前後加括弧 d
        'Set d = ActiveDocument
        On Error GoTo eH:
        DoEvents
        d.Range.Cut
        d.Close wdDoNotSaveChanges
    End If
    
    'Set d = ActiveDocument
    Set d = dSource
    Rem 因為前面尚有「中國哲學書電子化計劃.只保留正文注文_且注文前後加括弧」會用到UndoRecord物件，且會關閉其文件，故以下此行所寫位置就很關鍵，否則會隨文件關閉而隨之無效。20230201癸卯年十一
    SystemSetup.stopUndo ur, "mark易學關鍵字"
    Set rng = d.Range
    endDocOld = d.Range.End
    '    If InStr(d.Range.text, vba.Chr(13) & vba.Chr(13) & vba.Chr(13) & vba.Chr(13)) > 0 Then
    ''        d.Range.Text = Replace(d.Range.Text, vba.Chr(13) & vba.Chr(13) & vba.Chr(13) & vba.Chr(13), vba.Chr(13) & vba.Chr(13) & vba.Chr(13))
    '    '保留格式，故用以下，不用以上
    '        With d.Range.Find
    '            If InStr(.Parent.text, vba.Chr(13) & vba.Chr(13) & vba.Chr(13) & vba.Chr(13)) > 1 Then
    '                .ClearFormatting
    '                '.Execute vba.Chr(13) & vba.Chr(13) & vba.Chr(13) & vba.Chr(13), , , , , , True, wdFindContinue, , vba.Chr(13) & vba.Chr(13) & vba.Chr(13), wdReplaceAll
                    Rem 此行會造成Word crash
    '                .Execute "^p^p^p^p", , , , , , True, wdFindContinue, , "^p^p^p", wdReplaceAll
    '            End If
    '            .ClearFormatting
    '        End With
    '    End If
    
    Rem 將剪貼簿內擬加入的文本規範化
    '"．　", ""), "　．"為各大段落界格，不宜清除，在成卷或大批複製時，很重要，免得各大段落文本都連在一起了 20240925
    'clipBTxt = Replace(Replace(Replace(Replace(Replace(VBA.Trim(SystemSetup.GetClipboardText), VBA.Chr(13) + VBA.Chr(10) + "空句子" + VBA.Chr(13) + VBA.Chr(10), VBA.Chr(13) + VBA.Chr(10) + VBA.Chr(13) + VBA.Chr(10)), VBA.Chr(9), ""), "．　", ""), "　．", ""), " ", vbNullString)
    clipBTxt = Replace(Replace(Replace(VBA.Trim(SystemSetup.GetClipboardText), VBA.Chr(13) + VBA.Chr(10) + "空句子" + VBA.Chr(13) + VBA.Chr(10), VBA.Chr(13) + VBA.Chr(10) + VBA.Chr(13) + VBA.Chr(10)), VBA.Chr(9), ""), " ", vbNullString)
    clipBTxt = 文字處理.trimStrForSearch_PlainText(clipBTxt)
    clipBTxt = 漢籍電子文獻資料庫.CleanTextPicPageMark(clipBTxt)
    For e = 0 To UBound(strAutoCorrection)
        clipBTxt = Replace(clipBTxt, strAutoCorrection(e), strAutoCorrection(e + 1))
        e = e + 1
    Next e
    searchedTerm = Keywords.易學Keywords_ToMark
        
    'If Selection.Type = wdSelectionIP Then
        Rem 判斷是否已含有該文本
        '如果不含其文本
        If Not Docs.isDocumentContainClipboardText_IgnorePunctuation(d, clipBTxt) Then
            Rem 文本相似度比對
            Dim similarCompare As New Collection
            Set similarCompare = Docs.similarTextCheckInSpecificDocument(d, clipBTxt)
            If similarCompare.item(1) Then
                word.Application.Activate
'                AppActivate word.ActiveWindow.Caption
                If MsgBox("文本相似度為 " & vbCr & similarCompare.item(3) _
                    & VBA.vbCr & vbCr & VBA.vbTab & "相似段落為：" & VBA.vbCr & VBA.vbCr & VBA.IIf(VBA.Len(similarCompare.item(2)) > 255, VBA.Left(similarCompare.item(2), 255) & "……", similarCompare.item(2)) & vbCr & vbCr & vbCr & _
                    "按下「確定」將會選取類似段落，請自行檢查是否仍要再貼入" & vbCr & vbCr & "按下「取消」則忽略檢查，將繼續執行", vbExclamation + vbOKCancel, "要貼入的文本在原文件中有類似的段落!!!") _
                        = vbOK Then
                    Set rng = d.Range
                    If rng.Find.Execute(VBA.Left(similarCompare.item(2), 255), , , , , , , wdFindContinue) Then
                        If VBA.Len(similarCompare.item(2)) > 255 Then
                            rng.Paragraphs(1).Range.Select                  '標示相似文本
                            d.ActiveWindow.ScrollIntoView Selection.Characters(1), True
                        Else
                            rng.Select
                        End If
                    End If
                    Set similarCompare = Nothing
                    GoTo exitSub
                End If
            End If
            Set similarCompare = Nothing
            Rem end 文本相似度比對
            
            Rem 含有必須的關鍵字才貼上
            For Each e In searchedTerm
                If InStr(clipBTxt, e) > 0 Then
                    flgPaste = True '如果含有必須的關鍵字
                    Exit For
                End If
            Next e
            If Not flgPaste Then
                'vba.Chrw() & vba.Chrw() &'vba.Chrw() & vba.Chrw() &
                Dim guaKeyword
                guaKeyword = Keywords.易學Keywords_ToCheck
                For Each e In guaKeyword
                    If InStr(clipBTxt, e) > 0 Then
                        flgPaste = True
                        Exit For
                    End If
                Next e
            End If
            
            If flgPaste Then
pasteAnyway:
                d.Activate
                If Selection.Document.FullName <> d.FullName Then
                    Stop
                End If
                
                On Error GoTo eH
                If pasteRange Is Nothing Then
                    Selection.EndKey wdStory
                    Selection.InsertParagraphAfter
        '            Selection.InsertParagraphAfter
                Else
                    pasteRange.Select
                End If
                Selection.Collapse wdCollapseEnd
                Selection.TypeText clipBTxt
                'SystemSetup.SetClipboard clipBTxt
                'Docs.貼上純文字
                If pasteRange Is Nothing Then
                    Selection.InsertParagraphAfter: Selection.InsertParagraphAfter: Selection.InsertParagraphAfter
                    Selection.Collapse wdCollapseEnd
                End If
                ActiveWindow.ScrollIntoView Selection
            Else
                Dim noneYijingKeyword As Boolean
                noneYijingKeyword = True
            End If
        Else '如果文件中已有文本，則顯示其所在處
            Dim sx As String
            If InStr(d.Content, clipBTxt) Then
                'rng.Find.Execute VBA.Left(clipBTxt, 255), , , , , , , wdFindContinue
                Dim ps As Integer
                ps = InStr(clipBTxt, VBA.Chr(13)) '如有本來要貼入的文本中有段落，則止到其段落前為止；若沒有，則取能尋找的最大值255個字元長的內容作搜尋
                sx = VBA.IIf(ps > 0, VBA.Left(VBA.Mid(clipBTxt, 1, VBA.IIf(ps > 0, ps, 2) - 1), 255), VBA.Left(clipBTxt, 255))
            Else '標點符號處理：確定文本已有只是標點符號不同者
                punc.clearPunctuations clipBTxt
                punc.restoreOriginalTextPunctuations d.Range.text, clipBTxt
                Set punc = Nothing
                sx = 文字處理.trimStrForSearch_PlainText(clipBTxt)
                SystemSetup.SetClipboard sx
                sx = VBA.Left(sx, 255)
            End If
            rng.Find.Execute sx, , , , , , , wdFindContinue
            endDocOld = rng.End
    
        End If
    'End If
    
    If flgPaste Then
        word.Application.ScreenUpdating = False
        If d.path <> "" Then
            If InStr(ActiveDocument.path, "易學雜著文本") = 0 Then
                Set d = dSource
            End If
            If Not d.Saved Then d.Save
        End If
        
        Rem 標識關鍵字
        If Not doNotMark Then
        '    xd = d.Range.text
            Dim rngMark As Range
            
            Set rngMark = d.Range(IIf(endDocOld >= d.Range.End, d.Range.End - 1, endDocOld), d.Range.End)
            
            marking易學關鍵字 rngMark, searchedTerm, word.wdYellow, wdRed, False
            
        End If
        Rem 以上標識關鍵字
        
        GoSub refres
        SystemSetup.playSound 1.921
        Rem https://en.wikipedia.org/wiki/CJK_Unified_Ideographs
        Rem 兼容字
        'https://en.wikipedia.org/wiki/CJK_Compatibility_Ideographs
    '    Docs.ChangeFontOfSurrogatePairs_Range "HanaMinA", d.Range(selection.Paragraphs(1).Range.start, d.Range.End), CJK_Compatibility_Ideographs
        'https://en.wikipedia.org/wiki/CJK_Compatibility_Ideographs_Supplement
        Dim rngChangeFontName As Range
        'Set rngChangeFontName = d.Range(Selection.Paragraphs(1).Range.start, d.Range.End)
        Set rngChangeFontName = d.Range(rngMark.start, d.Range.End)
        Dim fontName As String '20240920 creedit_with_Copilot大菩薩:https://sl.bing.net/9KC0PtODtI
        fontName = "全宋體-2"
        If Fonts.IsFontInstalled(fontName) Then
            'MsgBox fontName & " 已安裝在系統中。"
        Else    'MsgBox fontName & " 未安裝在系統中。"
            fontName = "HanaMinA"
            If Fonts.IsFontInstalled("HanaMinA") Then
            ElseIf Fonts.IsFontInstalled("TH-Tshyn-P2") Then
                fontName = "TH-Tshyn-P2"
            Else
                fontName = vbNullString
            End If
        End If
        If Not fontName = vbNullString Then
            'Docs.ChangeFontOfSurrogatePairs_Range "HanaMinA", rngChangeFontName, CJK_Compatibility_Ideographs_Supplement
            Docs.ChangeFontOfSurrogatePairs_Range fontName, rngChangeFontName, CJK_Compatibility_Ideographs_Supplement
        End If
        
        Rem 擴充字集
        'HanaMinB還不支援G以後的
        fontName = "HanaMinB"
        Docs.ChangeFontOfSurrogatePairs_Range fontName, rngChangeFontName, CJK_Unified_Ideographs_Extension_E
        Docs.ChangeFontOfSurrogatePairs_Range fontName, rngChangeFontName, CJK_Unified_Ideographs_Extension_F
        returnVaule = True
        
    Else '文件內已有內容時
        GoSub refres
        SystemSetup.playSound 1.294
        If noneYijingKeyword Then
            If MsgBox("要貼上的文本並不含有易學關鍵字哦！" + vbCr + vbCr + _
                "請再檢查所複製到剪貼簿的內容是否正確。感恩感恩　南無阿彌陀佛" & _
                "○是否仍要貼上？" + vbCr + vbCr + clipBTxt, vbOKCancel + vbExclamation + vbDefaultButton2) _
                = vbOK Then
                noneYijingKeyword = False
                GoTo pasteAnyway
            End If
        End If
    End If
    
exitSub:
    SystemSetup.contiUndo ur
    Set ur = Nothing
    'word.Application.ScreenUpdating = True
    'word.Application.ScreenRefresh
    mark易學關鍵字 = returnVaule
    Exit Function
    
    
refres:
        word.Application.ScreenUpdating = True
        If flgPaste Then
            Rem 先省略，免得每次貼入都做一次，文字處理.書名號篇名號標注，應當等做完時要關閉檔案前再做
            '文字處理.書名號篇名號標注
            'If flgPaste Then'測試無礙後可刪此行
            '顯示新貼上的文本頂端
            rng.SetRange endDocOld, endDocOld
            Do Until rng.font.ColorIndex = wdRed Or rng.End = d.Range.End - 1
                rng.Move
            Loop
            e = rng.End
            rng.SetRange endDocOld, e
        Else
            rng.SetRange endDocOld - Len(sx), endDocOld
        End If
        rng.Select
        Static cntr  As Byte

        '如果選取範圍是文件末端，顯係沒找到，故以剪貼簿內前25字再找一次

        If rng.End = rng.Document.Range.End Then

            SystemSetup.SetClipboard VBA.Left(SystemSetup.GetClipboard, 25)

            If cntr < 2 Then
                cntr = cntr + 1
                If VBA.vbOK = MsgBox("沒找到位置，是否重試？感恩感恩　南無阿彌陀佛", vbOKCancel + vbExclamation) Then
                    mark易學關鍵字
                Else
                    cntr = 2
                End If
            Else
                cntr = 0
            End If

        End If
    '    word.Application.ScreenRefresh
        ActiveWindow.ScrollIntoView Selection.Characters(1) ', False
    Return
    
eH:
    Select Case Err.Number
        Case 5825 '物件已被刪除。
            GoTo exitSub
        Case Else
            MsgBox Err.Number & Err.Description
            Resume
    End Select
End Function
Rem 根據選取文字高亮文件中所有此文字出現的部分，所高亮者並不隨文件儲存，僅作顯示爾 20240922
Sub HitHighlightBySelecton()
    If Selection.Type = wdSelectionIP Then Exit Sub
    Dim rng As Range
    Set rng = Selection.Document.Range
    rng.Find.HitHighlight Selection.text, wdColorYellow
    Debug.Print rng.start; rng.End; rng.Document.Range.start; rng.Document.Range.End
End Sub

Rem 20240922 Copilot大菩薩根據我的改良，試試 https://sl.bing.net/dtWVmyauIFw https://sl.bing.net/glAQGL0KKCO
Rem 邏輯錯誤，完全不行！
Sub marking易學關鍵字_BAD_Copilot大菩薩(rng As Range, arr As Variant, Optional defaultHighlightColorIndex As word.WdColorIndex = word.wdYellow, _
        Optional fontColor As word.WdColorIndex = word.wdRed, Optional allDoc As Boolean = False)
    
    Dim regex As Object, matches As Object, match As Object
    Dim startRng As Long, endRng As Long
    Dim e As Variant
    Dim examOK As Boolean, rngExam As Range
    Dim isInPhrasesAvoid As Boolean, isFollowedAvoid As Boolean, isPrecededAvoid As Boolean
    Dim dictCoordinatesPhrase As New Scripting.Dictionary, key
    
    On Error GoTo eH
    word.options.defaultHighlightColorIndex = defaultHighlightColorIndex
    If allDoc Then Set rng = rng.Document.Range
    startRng = rng.start
    endRng = rng.End
    Set rngExam = rng.Document.Range
    
    ' 建立正則表達式對象
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    
    For Each e In arr
        'Copilot大菩薩：這樣的正則表達式模式確保了我們只匹配整個單詞 e，而不是單詞的一部分。例如，如果 e 是 “word”，這個模式會匹配 “word”，但不會匹配 “wording” 或 “sword”。
        'regex.Pattern = "\b" & e & "\b" 'Copilot大菩薩：b 在正則表達式中代表的是「單詞邊界」（word boundary），可以理解為「bound」的縮寫。它用來匹配單詞的開始或結束位置，確保我們只匹配整個單詞而不是單詞的一部分。 https://sl.bing.net/kIwNUNbDloO
        regex.Pattern = "." & e & "."
        Set matches = regex.Execute(rng.text)
        If matches.Count > 0 Then Stop
        For Each match In matches
            examOK = True
            rng.SetRange startRng + match.FirstIndex, startRng + match.FirstIndex + Len(match.Value)
            
            ' 檢查是否需要前後文檢查
            isFollowedAvoid = Keywords.易學KeywordsToMark_Exam_Followed_Avoid.Exists(e)
            isPrecededAvoid = Keywords.易學KeywordsToMark_Exam_Preceded_Avoid.Exists(e)
            isInPhrasesAvoid = Keywords.易學KeywordsToMark_Exam_InPhrase_Avoid.Exists(e)
            
            If isFollowedAvoid Or isPrecededAvoid Or isInPhrasesAvoid Then
                ' 後綴檢查
                If Not rng.Next Is Nothing And rng.Next.Characters.Count > 0 Then
                    If isFollowedAvoid Then
                        For Each key In Keywords.易學KeywordsToMark_Exam_Followed_Avoid(e)
                            If rng.End + Len(key) <= endRng Then
                                rngExam.SetRange rng.End, rng.End + Len(key)
                                If StrComp(rngExam.text, key) = 0 Then
                                    examOK = False
                                    Exit For
                                End If
                            End If
                        Next key
                    End If
                End If
                
                ' 前綴檢查
                If examOK And Not rng.Previous Is Nothing And rng.Previous.Characters.Count > 0 Then
                    If isPrecededAvoid Then
                        For Each key In Keywords.易學KeywordsToMark_Exam_Preceded_Avoid(e)
                            If rng.start - Len(key) > -1 Then
                                rngExam.SetRange rng.start - Len(key), rng.start
                                If StrComp(rngExam.text, key) = 0 Then
                                    examOK = False
                                    Exit For
                                End If
                            End If
                        Next key
                    End If
                End If
                
                ' 內嵌於檢查
                If examOK And isInPhrasesAvoid Then
                    If dictCoordinatesPhrase.Count = 0 Then
                        For Each key In Keywords.易學KeywordsToMark_Exam_InPhrase_Avoid(e)
                            rngExam.SetRange startRng, endRng
                            Do While rngExam.Find.Execute(key, , , , , , True, wdFindStop)
                                dictCoordinatesPhrase.Add rngExam.start, rngExam.End
                            Loop
                        Next key
                    End If
                    
                    For Each key In dictCoordinatesPhrase
                        If rng.start >= key And rng.End <= dictCoordinatesPhrase(key) Then
                            examOK = False
                            Exit For
                        End If
                    Next key
                End If
            End If
            
            ' 標識關鍵字
            If examOK Then
                With rng
                    .HighlightColorIndex = defaultHighlightColorIndex
                    If .font.ColorIndex = wdAuto Then .font.ColorIndex = fontColor
                End With
            Else
                ' 不需要前後文檢查的情形
                Do While regex.Execute(e, , , , , , True, wdFindStop, True)
                    With rng
                        .HighlightColorIndex = defaultHighlightColorIndex
                        If .font.ColorIndex = wdAuto Then .font.ColorIndex = fontColor
                    End With
                Loop
            End If
        Next match
        
        If dictCoordinatesPhrase.Count > 0 Then dictCoordinatesPhrase.RemoveAll
    Next e

finish:
    rng.SetRange startRng, endRng
    Exit Sub

eH:
    MsgBox Err.Number & Err.Description
    Resume finish
End Sub

Rem rng 要處理的範圍 ,arr 要處理的關鍵字 （預設為字串陣列）
Sub marking易學關鍵字(rng As Range, arr As Variant, Optional defaultHighlightColorIndex As word.WdColorIndex = word.wdYellow, _
        Optional fontColor As word.WdColorIndex = word.wdRed, Optional allDoc As Boolean = False)
    '優化的建議：Copilot大菩薩 20240922 Word VBA 中的 Find 物件屬性： https://sl.bing.net/fMV3NYyXdLg https://sl.bing.net/kosqk2rrnFc
    Dim xd As String, a As Range, e, eArrKey, arrKey, startRng As Long, endRng As Long ', dict As Scripting.Dictionary
    Dim examOK As Boolean, rngExam As Range, processCntr As Long, dictCoordinatesPhrase As New Scripting.Dictionary, key, isInPhrasesAvoid As Boolean, isFollowedAvoid As Boolean, isPrecededAvoid As Boolean

    On Error GoTo eH
    word.options.defaultHighlightColorIndex = defaultHighlightColorIndex
    examOK = True
    If allDoc Then Set rng = rng.Document.Range
    startRng = rng.start
    endRng = rng.End
    Set rngExam = rng.Document.Range

    xd = rng.text

    With rng.Find
        .ClearFormatting
'        .HitHighlight
        
'        With .Replacement '現在不用 wdReplaceAll 引數的方法了
'            .font.ColorIndex = fontColor 'wdRed
'            .Highlight = True
'        End With
        For Each e In arr '遍歷每個要標識的關鍵字
        
'            If e = "臨卦" Then Stop 'just for test
            
            If InStr(xd, e) > 0 Then '在有超連結等功能變數、隱藏文字時可能會miss，今測試並不會，待再測試。
            
                rng.SetRange startRng, endRng
                'If e = "豫" Then
                isFollowedAvoid = Keywords.易學KeywordsToMark_Exam_Followed_Avoid.Exists(e)
                isPrecededAvoid = Keywords.易學KeywordsToMark_Exam_Preceded_Avoid.Exists(e)
                isInPhrasesAvoid = Keywords.易學KeywordsToMark_Exam_InPhrase_Avoid.Exists(e)
                If isFollowedAvoid Or isPrecededAvoid Or isInPhrasesAvoid Then
'                    '若有內嵌於的片語語詞（包含關鍵的片語語句）需查驗的話，就先記下要比對的位置集合
'                    '若後、前檢查都未過就不必了，故今再改移至內嵌之檢查，並以 dictCoordinatesPhrase.Count 來作判斷
'                    If isInPhrasesAvoid Then
'                        arrKey = Keywords.易學KeywordsToMark_Exam_InPhrase_Avoid(e)
'                        For Each eArrKey In arrKey '遍歷每個該避開的詞語，蒐集其在文件中的所在位置，以供後續比對
'                            'Set rngExam = rngExam.Document.Range '弄錯了，已有allDoc參數控制是否操作整份文件
'                            rngExam.SetRange startRng, endRng '不需處理整份文件，就只就操作範圍內處置就可。感恩感恩　讚歎讚歎　南無阿彌陀佛 20240915
'                            With rngExam.Find
'                                Do While .Execute(eArrKey, , , , , , True, wdFindStop)
'                                    '記下含有目前關鍵字的詞語片語語彙片段在範圍中的位置
'                                    dictCoordinatesPhrase.Add rngExam.start, rngExam.End
'                                Loop
'                            End With
'                        Next eArrKey
'                    End If
                    
                    Do While .Execute(e, , , , , , True, wdFindStop, True) '在範圍中尋遍關鍵字出現的位置
                        examOK = True '歸零
                        
                        'rng.Select '偵錯用
                        
                        If Not rng.Next Is Nothing Then
                            If rng.Next.Characters.Count > 0 Then
                                '後綴檢查
                                
                                'If rng.Document.Range(rng.End, rng.End + 4).text = "坤之清氣" Then Stop 'just for test
                                
                                'If UBound(VBA.Filter(Keywords.易學KeywordsToMark_Exam_Followed_Avoid(e), rng.Next.Characters(1).text)) < 0 Then
                                If isFollowedAvoid Then
                                    arrKey = Keywords.易學KeywordsToMark_Exam_Followed_Avoid(e)
                                    For Each eArrKey In arrKey
                                        If rng.End + VBA.Len(eArrKey) <= endRng Then '這樣的寫法，如果內含超連結等功能變數，恐怕就會失誤了！
                                            rngExam.SetRange rng.End, rng.End + VBA.Len(eArrKey)
                                            If VBA.StrComp(rngExam.text, eArrKey) = 0 Then '找到要避開的關鍵字
                                                examOK = False '檢測不通過
                                                Exit For
                                            End If
                                        End If
                                    Next eArrKey
                                End If
                                If examOK Then
checkPrevious:
                                    If Not rng.Previous Is Nothing Then
                                        '前綴檢查
                                        If rng.Previous.Characters.Count > 0 Then
                                            'If UBound(VBA.Filter(Keywords.易學KeywordsToMark_Exam_Preceded_Avoid(e), rng.Previous.Characters(rng.Previous.Characters.Count).text)) < 0 Then
                                            If isPrecededAvoid Then
                                                arrKey = Keywords.易學KeywordsToMark_Exam_Preceded_Avoid(e)
                                                For Each eArrKey In arrKey
                                                    If rng.start - VBA.Len(eArrKey) > -1 Then '這樣的寫法，如果內含超連結等功能變數，恐怕就會失誤了！
                                                        rngExam.SetRange rng.start - VBA.Len(eArrKey), rng.start
                                                        If VBA.StrComp(rngExam.text, eArrKey) = 0 Then '找到要避開的關鍵字
                                                            examOK = False '檢測不通過
                                                            Exit For
                                                        End If
                                                    End If
                                                Next eArrKey
                                            End If
                                            If examOK Then
checkPhrases:                                   '內嵌於檢查：關鍵字含在該避免的片語詞句檢查
                                                'If Keywords.易學KeywordsToMark_Exam_InPhrase_Avoid.Exists(e) Then
                                                If isInPhrasesAvoid Then
                                                    If dictCoordinatesPhrase.Count = 0 Then
                                                        GoSub buildDictCoordinatesPhrase
                                                    End If
                                                
                                                    For Each key In dictCoordinatesPhrase
                                                     '遍歷每個該避開的詞語片語座標
                                                        '若目前關鍵字內含於要避開的詞語片語語彙片段
                                                        If rng.start >= key And rng.End <= dictCoordinatesPhrase(key) Then

                                                            'rng.Select 'just for test

                                                            examOK = False '檢測不通過
                                                            Exit For
                                                        End If
                                                    Next key

'                                                    arrKey = Keywords.易學KeywordsToMark_Exam_InPhrase_Avoid(e)
'                                                    For Each eArrKey In arrKey '遍歷每個該避開的詞語，當先蒐集其位置在文件中所在位置，以供後續比對
'                                                        Set rngExam = rngExam.Document.Range
'                                                        With rngExam.Find
'                                                            Do While .Execute(eArrKey, , , , , , True, wdFindStop)
'                                                                '目前關鍵字內含於要避開的詞語片語語彙片段
'                                                                If rng.start >= rngExam.start And rng.End <= rngExam.End Then
'
''                                                                    rng.Select 'just for test
'
'                                                                    examOK = False '檢測不通過
'                                                                    Exit For
'                                                                End If
'                                                            Loop
'                                                        End With
'                                                    Next eArrKey
                                                    
                                                End If
                                                '後、前、中三關都檢驗合格了
                                                If examOK Then '合格才標識
                                                    With rng
                                                        processCntr = processCntr + 1
                                                        If processCntr Mod 35 = 0 Then SystemSetup.playSound 1 '播放音效以免誤以為當了
                                                        Rem 此檔可供效能測試，跑起來不知何故特別久！ file:///H:\我的雲端硬碟\黃老師遠端工作\1易學雜著文本\非清人已初步標點\何良俊＠四友齋叢說.docx
                                                        
                                                        .HighlightColorIndex = defaultHighlightColorIndex
                                                        For Each a In rng.Characters
                                                            If a.font.ColorIndex = wdAuto Then a.font.ColorIndex = fontColor
                                                        Next a
                                                    End With
                                                Else '片語詞句檢查未過
                                                    If rng.HighlightColorIndex = defaultHighlightColorIndex Then
                                                        With rng
                                                            '.Select 'just for test
                                                            
                                                            .HighlightColorIndex = wdNoHighlight
                                                            .font.ColorIndex = wdAuto
                                                        End With
                                                    End If
                                                End If
                                            Else '前綴檢查未過
                                                'If allDoc Then
                                                If rng.HighlightColorIndex = defaultHighlightColorIndex Then
                                                    With rng
                                                        .HighlightColorIndex = wdNoHighlight
                                                        .font.ColorIndex = wdAuto
                                                    End With
                                                End If
                                            End If
                                        End If
                                    Else '如果前也無後也無
                                        GoTo checkPhrases:
                                    End If
                                Else '後綴檢查未過
                                    'If allDoc Then
                                    If rng.HighlightColorIndex = defaultHighlightColorIndex Then
                                        With rng
                                            .HighlightColorIndex = wdNoHighlight
                                            .font.ColorIndex = wdAuto
                                        End With
                                    End If
                                End If
                            End If
                        Else '如果沒有下文
                            GoTo checkPrevious
                        End If
                    Loop 'Do While .Execute(e, , , , , , True, wdFindStop, True) '在範圍中尋遍關鍵字出現的位置

                    If dictCoordinatesPhrase.Count > 0 Then
                        dictCoordinatesPhrase.RemoveAll '歸零供下一個關鍵字使用
                    End If
                        'If Keywords.易學KeywordsToMark_Exam_InPhrase_Avoid.Exists(e) Then
                         '       ……
                          '          dictCoordinatesPhrase.Add rngExam.start, rngExam.End
                          
                Else '不用檢查直接取代（逕行標識）者
                    Do While .Execute(e, , , , , , True, wdFindStop, True) '此效能也不會比 wdReplaceAll 引數者慢，可見其內當亦係用類似之迴圈實作者也 20240919 感恩感恩　讚歎讚歎　南無阿彌陀佛
'                        .Parent.HighlightColorIndex = defaultHighlightColorIndex
'                        .Parent.Font.ColorIndex = fontColor
                        Rem 若寫成以下會有49DLL呼叫規格錯誤，這次重寫又不會了，可見是 VBE編譯器故障 20240920
                        With rng
                            .HighlightColorIndex = defaultHighlightColorIndex
                            For Each a In rng.Characters
                                If a.font.ColorIndex = wdAuto Then a.font.ColorIndex = fontColor
                            Next a
                            processCntr = processCntr + 1
                            If processCntr Mod 35 = 0 Then SystemSetup.playSound 1 '播放音效以免誤以為當了
                        End With
                    Loop
'                    .Execute e, , , , , , True, wdFindStop, True, e, Replace:=wdReplaceAll '在含有超連結等格式化文字時會失靈
                    'rng.SetRange startRng, endRng'前已有
                End If
            End If
        Next e '下一個打算要標識的關鍵字
    End With

finish:
    rng.SetRange startRng, endRng '設回原來的樣子才不會改變，於呼叫端才不會出錯
    Set dictCoordinatesPhrase = Nothing
    
    Exit Sub
    
buildDictCoordinatesPhrase:
        '若有內嵌於的片語語詞（包含關鍵的片語語句）需查驗的話，就先記下要比對的位置集合
'        若後、前檢查都未過就不必了，故今再改移至內嵌之檢查，並以 dictCoordinatesPhrase.Count 來作判斷
        If isInPhrasesAvoid Then
            arrKey = Keywords.易學KeywordsToMark_Exam_InPhrase_Avoid(e)
            For Each eArrKey In arrKey '遍歷每個該避開的詞語，蒐集其在文件中的所在位置，以供後續比對
                    '已有allDoc參數控制是否操作整份文件
                rngExam.SetRange startRng, endRng '故不需處理整份文件，就只就操作範圍內處置就可。感恩感恩　讚歎讚歎　南無阿彌陀佛 20240915
                With rngExam.Find
                    Do While .Execute(eArrKey, , , , , , True, wdFindStop)
'                        記下含有目前關鍵字的詞語片語語彙片段在範圍中的位置
                        dictCoordinatesPhrase.Add rngExam.start, rngExam.End
                    Loop
                End With
            Next eArrKey
        End If
        
        Return

eH:
    Select Case Err.Number
        Case Else
            MsgBox Err.Number & Err.Description
            Debug.Print Err.Number & Err.Description
            Resume
    End Select
End Sub
Rem 整個文件重新標識易學關鍵字
Sub mark易學關鍵字Doc()
    Dim ur As word.UndoRecord
    SystemSetup.playSound 0.484
    SystemSetup.stopUndo ur, "mark易學關鍵字Doc"
    word.Application.ScreenUpdating = False
    marking易學關鍵字 ActiveDocument.Range, Keywords.易學Keywords_ToMark, word.wdYellow
    SystemSetup.contiUndo ur
    SystemSetup.playSound 2
    word.Application.ScreenUpdating = True
End Sub

Rem 判斷剪貼簿裡的純文字(或指定的文字)內容是否在文件中已存在
Function isDocumentContainClipboardText_IgnorePunctuation(d As Document, Optional chkClipboardText As String) As Boolean
    Dim xd As String
    xd = d.Range.text
    If VBA.Len(xd) = 1 Then Exit Function
    
    If chkClipboardText = "" Then chkClipboardText = SystemSetup.GetClipboardText
    Rem 剪貼簿裡的換行符號值是vba.Chr(13)&vba.Chr(10)而在Word文件中是只有 vba.Chr(13)
    chkClipboardText = VBA.Replace(chkClipboardText, VBA.Chr(13) & VBA.Chr(10), VBA.Chr(13))
    
    If VBA.InStr(xd, chkClipboardText) > 0 Then
        isDocumentContainClipboardText_IgnorePunctuation = True
    Else '忽略標點符號的比對
        Dim punc As New punctuation
        If punc.inStrIgnorePunctuation(xd, chkClipboardText) Then
            isDocumentContainClipboardText_IgnorePunctuation = True
        Else
            If isDocumentContainClipboardText_IgnorePunctuation Then isDocumentContainClipboardText_IgnorePunctuation = False
        End If
        Set punc = Nothing
    End If
End Function

Function similarTextCheckInSpecificDocument(d As Document, text As String) As Collection 'item1 as Boolean(文本是否相似),item2 as string(找到的相似文本段落),item3 as String from Dictionary SimilarityResult(相似度名&相似度)
Rem 文本相似度比對
Dim similarText As New similarText, dClearPunctuation As String, textClearPunctuation As String, dCleanParagraphs() As String, punc As New punctuation, e, Similarity As Boolean, result As New Collection
dClearPunctuation = d.Content.text
textClearPunctuation = text
'清除標點符號
punc.clearPunctuations textClearPunctuation: punc.clearPunctuations dClearPunctuation
dCleanParagraphs = VBA.Split(dClearPunctuation, VBA.Chr(13))
Dim cntr As Long
For Each e In dCleanParagraphs
    cntr = cntr + 1
    If cntr Mod 20 = 0 Then SystemSetup.playSound 1
    If e <> "" Then
'        If e = "易" Then Stop
        If similarText.Similarity(e, textClearPunctuation) Then
            Similarity = True: Exit For
        ElseIf similarText.SimilarityPercent(e, textClearPunctuation) > 80 Then
            Similarity = True: Exit For
        End If
    End If
Next e
'If Similarity = True Then Stop 'for test
Rem index   Required. An expression that specifies the position of a member of the collection. If a numeric expression, index must be a number from 1 to the value of the collection's Count property. If a string expression, index must correspond to the key argument specified when the member referred to was added to the collection.
result.Add Similarity 'item1:文本是否相似'https://learn.microsoft.com/en-us/office/vba/Language/Reference/User-Interface-Help/item-method-visual-basic-for-applications
dClearPunctuation = e
punc.restoreOriginalTextPunctuations d.Content.text, dClearPunctuation
result.Add dClearPunctuation 'item2:找到的相似文本段落
result.Add similarText.SimilarityResultsString 'item3:相似度名&相似度
Set similarText = Nothing
Set similarTextCheckInSpecificDocument = result
Rem end 文本相似度比對
End Function
Sub 文件比對_抓抄襲()
Dim d1 As Document, d2 As Document, p As Paragraph, x As String, i As Byte, rng As Range, pc As Long, d1RngTxt, px As String, rng2 As Range
Static pi As Long
Set d1 = Documents(1) '來源
d1RngTxt = d1.Range.text
Set d2 = Documents(2) '抄襲或引用(須先將。，的句子單位拆成各段文字）
pc = d2.Paragraphs.Count
If pi = 0 Then pi = 1
For pi = pi To pc
    Set p = d2.Paragraphs(pi)
    If p.Range.font.NameFarEast <> "標楷體" And p.Range.HighlightColorIndex = 0 Then
        px = p.Range
        x = VBA.Trim(VBA.Left(px, Len(px) - 1)) '去掉分段符號
        If Len(x) > 2 Then
            x = VBA.Left(x, Len(x) - 1) '去掉端後標點。，等
            If VBA.InStr(d1RngTxt, x) Then
                i = i + 1
            Else
                i = 0
            End If
        End If
        If i > 1 And Len(x) > 2 Then
            Set rng = d1.Range
            rng.Find.Execute x
            rng.Select
            rng.Copy
            d1.Activate
            Set rng2 = d2.Range
            rng2.Find.Execute x
            rng2.Select
            If d2.ActiveWindow.Selection.Range.HighlightColorIndex = 0 Then
                SystemSetup.playSound 2
                Exit Sub
            End If
        End If
    End If
Next
pi = 0
SystemSetup.playSound 3
End Sub

Sub 抽取超連結位址()
Dim hplnk As Hyperlink, x As String, d As Document
For Each hplnk In ActiveDocument.Hyperlinks
    x = x & hplnk.Address & VBA.Chr(13)
Next hplnk
Set d = Documents.Add
d.Range.text = x
d.Range.Cut
d.Close wdDoNotSaveChanges
End Sub


Sub 插入超連結_文件中的位置_標題() 'Alt+P 原是「引詩」樣式'2021/11/27
Dim d As Document, title As String, p As Paragraph, pTxt As String, subAddrs As String, flg As Boolean
Set d = ActiveDocument
title = Selection.text
title = 文字處理.trimStrForSearch(title, Selection)
For Each p In d.Paragraphs
    If VBA.Left(p.Style.NameLocal, 2) = "標題" Then
        pTxt = p.Range.text
        pTxt = VBA.Left(pTxt, Len(pTxt) - 1)
        If StrComp(pTxt, title) = 0 Then
            subAddrs = title
            flg = True
            Exit For
        ElseIf InStr(pTxt, title) > 0 Then
            subAddrs = "_" & VBA.Mid(pTxt, 1, InStrRev(pTxt, " ") - 1)
            subAddrs = Replace(subAddrs, " ", "_")
            flg = True
            Exit For
        End If
    End If
Next p
'
'    'Selection.MoveLeft Unit:=wdCharacter, Count:=4, Extend:=wdExtend
'    'ChangeFileOpenDirectory d.path & "\"  ''userProfilePath & "Dropbox\"
'
If flg Then
    d.Hyperlinks.Add Anchor:=Selection.Range, Address:="", _
        SubAddress:=subAddrs, ScreenTip:="", TextToDisplay:=title
Else
    MsgBox "請手動插入！", vbExclamation
End If
End Sub
Rem 有選取時，即以選取的文字處理
Sub 中國哲學書電子化計劃_只保留正文注文_且注文前後加括弧_貼到古籍酷自動標點()
    Dim ur As UndoRecord, d As Document, x As String, i As Long
    Dim SelectionRange As Range
    SystemSetup.playSound 0.484
    If (ActiveDocument.path <> "" And Not ActiveDocument.Saved) Then ActiveDocument.Save
    
    Rem 這行要寫在不用的文件關閉後才有效，蓋其與文件併走也（雖UndoRecord為Application的屬性，但在文件被關閉時，其所載之復原記錄也會隨之清除，故須寫在文件關閉後才有效）
    SystemSetup.stopUndo ur, "中國哲學書電子化計劃_註文前後加括弧_貼到古籍酷自動標點"
    
    If Selection.Type = wdSelectionNormal Then
        Selection.Cut
        Set SelectionRange = Selection.Range
    End If
    'If Documents.Count = 0 Then
    '    Set d = Docs.空白的新文件()
    'Else
    '    Set d = ActiveDocument
    'End If
    word.Application.ScreenUpdating = False
    Set d = Docs.空白的新文件()
    VBA.DoEvents
    中國哲學書電子化計劃.只保留正文注文_且注文前後加括弧 d
    
    If d.path <> "" Then
        MsgBox "本文檔已儲存，不能操作！", vbCritical
        Exit Sub
    End If
    If Len(d.Range) = 1 Then Exit Sub '空白文件不處理
    
    '以下2行已不必，待觀察 20240716
    '先要複製到剪貼簿,純文字操作即可
    'd.Range.Cut
    
    x = 文字處理.trimStrForSearch_PlainText(d.Range.text)
    x = 漢籍電子文獻資料庫.CleanTextPicPageMark(x)
    SystemSetup.SetClipboard VBA.Replace(x, "·", "") '以《古籍酷》自動標點不會清除「·」，造成書名號標點機制不正確，故於此先清除之。
    DoEvents
    'If d.path = "" Then '前已作判斷 If d.path <> "" Then Exit Sub
    d.Close wdDoNotSaveChanges
    
    Rem 這行要寫在不用的文件關閉後才有效，蓋其與文件併走也（雖UndoRecord為Application的屬性，但在文件被關閉時，其所載之復原記錄也會隨之清除，故須寫在文件關閉後才有效）
    Rem SystemSetup.stopUndo ur, "中國哲學書電子化計劃_註文前後加括弧_貼到古籍酷自動標點"
    
    '將剪貼簿中的文本內容，送交古籍酷自動標點
    If 貼到古籍酷自動標點() = True Then
        If Documents.Count = 0 Then GoTo exitSub
        ActiveDocument.Application.Activate
        '自動執行易學關鍵字標識
        If Documents.Count > 0 Then
            If InStr(ActiveDocument.path, "已初步標點") > 0 Then
            On Error GoTo eH:
                If Not SeleniumOP.WD Is Nothing Then
                    Dim ws() As String
                    ws = SeleniumOP.WindowHandles
                    If Not VBA.IsEmpty(ws) Then
                        WD.SwitchTo.Window ws(UBound(ws))
                        SeleniumOP.WD.Manage.Window.Minimize
                    End If
                End If
mark:
                ActiveDocument.Application.Activate
                mark易學關鍵字 SelectionRange
            End If
        End If
    End If
exitSub:
    SystemSetup.contiUndo ur
    word.Application.ScreenUpdating = True
    Exit Sub
eH:
        Select Case Err.Number
            Case 9
                If InStr(Err.Description, "陣列索引超出範圍") Then
                    GoTo mark
                Else
                    GoTo msg:
                End If
            Case -2146233088 'unknown error: unhandled inspector error: {"code":-32000,"message":"Browser window not found"}
                              '(Session info: chrome=126.0.6478.127)
                If VBA.InStr(Err.Description, "unknown error: unhandled inspector error:") > 0 Then
                    GoTo mark
                End If
            Case Else
msg:
                MsgBox Err.Number & Err.Description
        End Select
End Sub

'先要複製到剪貼簿
Function 貼到古籍酷自動標點() As Boolean
    Dim x As String, result As String, resumeTimer As Byte
    On Error GoTo Err1
    x = SystemSetup.GetClipboard
    x = Replace(x, VBA.Chr(0), "")
    If x = "" Then x = Selection
    result = SeleniumOP.grabGjCoolPunctResult(x, result)
    If result = "" Or result = x Then
        DoEvents
        貼到古籍酷自動標點SendKeys
    Else
        '寫到剪貼簿
        SystemSetup.SetClipboard result
        '完成放音效
        SystemSetup.playSound 1.469
        貼到古籍酷自動標點 = True
    End If
    
    Exit Function
Err1:
        Select Case Err.Number
            Case 49 'DLL 呼叫規格錯誤
                resumeTimer = resumeTimer + 1
                If resumeTimer > 2 Then
                    MsgBox Err.Description, vbCritical
                    SystemSetup.killchromedriverFromHere
                Else
                    Resume
                End If
            Case 5 'https://www.google.com/search?q=vba+Err.Number+5&oq=vba+Err.Number+5&aqs=chrome..69i57j0i10i30j0i30l2j0i5i30.4768j0j7&sourceid=chrome&ie=UTF-8
                SystemSetup.wait 1.5
                resumeTimer = resumeTimer + 1
                If resumeTimer > 2 Then
                    MsgBox Err.Description, vbCritical
                    SystemSetup.killchromedriverFromHere
                Else
                    Resume
                End If
            Case 13
                If InStr(Err.Description, "型態不符合") Then
                    SystemSetup.killchromedriverFromHere
    '                Stop
                    resumeTimer = resumeTimer + 1
                    If resumeTimer > 2 Then
                        MsgBox Err.Description, vbCritical
                        SystemSetup.killchromedriverFromHere
                    Else
                        Resume
                    End If
                Else
                    MsgBox Err.Description, vbCritical
                    Stop
        '           Resume
                End If
            Case -2146233088
                If InStr(Err.Description, "disconnected: not connected to DevTools") Then '(failed to check if window was closed: disconnected: not connected to DevTools)
                                                                                            '(Session info: chrome=110.0.5481.178)
                    SystemSetup.killchromedriverFromHere
    '                Stop
                    resumeTimer = resumeTimer + 1
                    If resumeTimer > 2 Then
                        MsgBox Err.Description, vbCritical
                        SystemSetup.killchromedriverFromHere
                    Else
                        Resume
                    End If
    
                Else
                    MsgBox Err.Description, vbCritical
                    SystemSetup.killchromedriverFromHere
        '           Resume
                End If
            Case Else
                If InStr(Err.Description, "no such window") Then
                    If Not WD Is Nothing Then Resume
                Else
                    MsgBox Err.Description, vbCritical
                    SystemSetup.killchromedriverFromHere
        '           Resume
                End If
        End Select

End Function
Sub 貼到古籍酷自動標點SendKeys()
'Dim d As Document
'Set d = ActiveDocument
'If d.path <> "" Then Exit Sub
'If SystemSetup.GetClipboard = "" Then
'    If Len(d.Range) = 1 Then Exit Sub '空白文件不處理
'    d.Range.Cut
'End If
On Error GoTo App
AppActivate "古籍酷"
DoEvents
'SendKeys "{TAB 16}", True'舊版
SendKeys "{TAB 15}", True
Dim x As String
x = SystemSetup.GetClipboardText
If InStr(x, VBA.Chr(13)) > 0 And InStr(x, VBA.Chr(13) & VBA.Chr(10)) = 0 Then
    x = VBA.Replace(x, VBA.Chr(13), VBA.Chr(13) & VBA.Chr(10) & VBA.Chr(13) & VBA.Chr(10))
    DoEvents
    SystemSetup.ClipboardPutIn x
End If
SystemSetup.wait 0.5
SendKeys "^v"
DoEvents
'SendKeys "+{TAB 2}~", True '舊版
SendKeys "+{TAB 1}~", True
wait 2
SendKeys "+{TAB 1} ", True
'If d.path = "" Then d.Close wdDoNotSaveChanges
Exit Sub
App:
Select Case Err.Number
    Case 5
        'Shell (Network.getDefaultBrowserFullname + " https://old.gj.cool/gjcool/index")'舊版
        Shell (Network.getDefaultBrowserFullname + " https://gj.cool/punct")
        AppActivate Network.getDefaultBrowserNameAppActivate '"古籍酷"
        DoEvents
        SystemSetup.wait 2.9 '2.5 打開網頁 等待載入完畢
        'SendKeys "{TAB 16}", True
        Resume Next
    Case Else
        MsgBox Err.Number & Err.Description
End Select
End Sub

Rem 20230224 creedit with  Bing菩薩：
Sub ChangeFontOfSurrogatePairs_ActiveDocument(fontName As String, Optional whatCJKBlock As CJKBlockName)
    Dim rng         As Range
    Dim C           As String
    Dim i           As Long
    Dim ur As UndoRecord
    SystemSetup.stopUndo ur, "ChangeFontOfSurrogatePairs_ActiveDocument"
    ' Loop through each character in the document
    For Each rng In ActiveDocument.Characters
        C = rng.text
        ' Check if the character is a high surrogate
        If AscW(C) >= &HD800 And AscW(C) <= &HDBFF Then
            ' Check if the next character is a low surrogate
            If rng.End < ActiveDocument.Content.End Then
                i = rng.End + 1        ' The index of the next character
                If i < ActiveDocument.Range.End Then
                    C = C & ActiveDocument.Range(i, i).text        ' The combined character
                End If
                If AscW(VBA.Right(C, 1)) >= &HDC00 And AscW(VBA.Right(C, 1)) <= &HDFFF Then
                    ' Check if the combined character is in CJK extension B or later
                    'If AscW(vba.Left(c, 1)) >= &HD840 Then
                    If AscW(VBA.Left(C, 1)) >= SurrogateCodePoint.HighStart Then '前導代理 (lead surrogates)，介於 D800 至 DBFF 之間，第二個被稱為 後尾代理 (trail surrogates)，介於 DC00 至 DFFF 之間
                        Dim change As Boolean
                        change = True
'                        rng.Select
                        Select Case whatCJKBlock
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_B
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_B)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_C
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_C)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_D
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_D)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_E
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_E)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_F
                                'change = isCJK_ExtF(c)
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_F)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_G
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_G)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_H
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_H)
                            Case Else
                            ' Change the font name to HanaMinB
                            ' Change the font name to fontName
                        End Select
                        If change Then rng.font.Name = fontName '"HanaMinB"
                    End If
                End If
            End If
        End If
    Next rng
    SystemSetup.contiUndo ur
End Sub
Sub ChangeFontOfSurrogatePairs_Range(fontName As String, rngtoChange As Range, Optional whatCJKBlock As CJKBlockName)
    Dim rng         As Range
    Dim C           As String
    Dim i           As Long
    Dim ur As UndoRecord
    SystemSetup.stopUndo ur, "ChangeFontOfSurrogatePairs_Range"
    For Each rng In rngtoChange.Characters
        C = rng.text
        
        Rem forDebugText
'        If c = vba.Chrw(-10122) & vba.Chrw(-8820) Or c = vba.Chrw(-10119) & vba.Chrw(-8987) Then Stop
        
        ' Check if the character is a high surrogate
        If AscW(C) >= &HD800 And AscW(C) <= &HDBFF Then
'            ' Check if the next character is a low surrogate
'            'If rng.End < ActiveDocument.Content.End Then
'            If rng.End < rngtoChange.End Then
'                i = rng.End + 1        ' The index of the next character
'                'If i < ActiveDocument.Range.End Then
'                If i < rngtoChange.End Then
'                    'c = c & ActiveDocument.Range(i, i).text        ' The combined character
'                    c = c & VBA.Mid(rngtoChange, i, 1).text        ' The combined character
'                End If
                If AscW(VBA.Right(C, 1)) >= &HDC00 And AscW(VBA.Right(C, 1)) <= &HDFFF Then
                    ' Check if the combined character is in CJK extension B or later
                    'If AscW(vba.Left(c, 1)) >= &HD840 Then
                    If AscW(VBA.Left(C, 1)) >= SurrogateCodePoint.HighStart Then '前導代理 (lead surrogates)，介於 D800 至 DBFF 之間，第二個被稱為 後尾代理 (trail surrogates)，介於 DC00 至 DFFF 之間
                        Dim change As Boolean, isCjkResult As Collection
                        change = True
'                        rng.Select
                        Select Case whatCJKBlock
                            Case CJKBlockName.CJK_Compatibility_Ideographs
                                 Set isCjkResult = IsCJK(C)
                                 If isCjkResult.item(1) Then
                                    If isCjkResult.item(2) <> CJKBlockName.CJK_Compatibility_Ideographs Then change = False
                                 End If
                            Case CJKBlockName.CJK_Compatibility_Ideographs_Supplement
                                 Set isCjkResult = IsCJK(C)
                                 If isCjkResult.item(1) Then
                                    If isCjkResult.item(2) <> CJKBlockName.CJK_Compatibility_Ideographs_Supplement Then change = False
                                 End If
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_B
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_B)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_C
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_C)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_D
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_D)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_E
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_E)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_F
                                'change = isCJK_ExtF(c)
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_F)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_G
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_G)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_H
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_H)
                            Case Else
                            ' Change the font name to HanaMinB
                            ' Change the font name to fontName
                        End Select
                        If change Then rng.font.Name = fontName '"HanaMinB"
                    End If
                End If
'            End If
        End If
    Next rng
    SystemSetup.contiUndo ur
End Sub
Sub ChangeCharacterFontName(character As String, fontName As String, d As Document, Optional fontNameFarEast As String)
With d.Range
    With .Find
        With .Replacement.font
            .Name = fontName
            .NameFarEast = fontNameFarEast
        End With
        .Execute character, , , , , , True, wdFindContinue, , , wdReplaceAll
    End With
End With
End Sub

Sub ChangeCharacterFontNameAccordingSelection()
Dim fontName As String, fontNameFarEast As String
With Selection
    fontName = .font.Name
    fontNameFarEast = .font.NameFarEast
    ChangeCharacterFontName .text, fontName, .Document, fontNameFarEast
End With
End Sub

Rem 20230224 chatGPT大菩薩或Bing in Skype 菩薩:
Sub FindMissingCharacters() '這應該只是找文件中的字不能以新細明體、標楷體來顯示者
    Dim Doc As Document
    Set Doc = ActiveDocument
    
    '定義新細明體和標楷體字型的集合
    Dim nmf As font
    Set nmf = Doc.Styles("Normal").font
    Dim kff As font
    Set kff = Doc.Styles("段落").font
    
    Dim p As Paragraph
    Dim r As Range
    Dim C As Variant
    
    ' 遍歷文檔中的每個段落和字符
    For Each p In Doc.Paragraphs
        For Each r In p.Range.Characters
            
            ' 判斷字符是否在新細明體或標楷體字型中
            C = r.text
            If Len(C) > 0 Then
                If (AscW(VBA.Left(C, 1)) >= &H4E00 And AscW(VBA.Left(C, 1)) <= &H9FFF) _
                    Or (AscW(VBA.Left(C, 1)) >= &H3400 And AscW(VBA.Left(C, 1)) <= &H4DBF) _
                    Or (AscW(VBA.Left(C, 1)) >= &H20000 And AscW(VBA.Left(C, 1)) <= &H2A6DF) _
                    Or (AscW(VBA.Left(C, 1)) >= &H2A700 And AscW(VBA.Left(C, 1)) <= &H2B73F) _
                    Or (AscW(VBA.Left(C, 1)) >= &H2B740 And AscW(VBA.Left(C, 1)) <= &H2B81F) _
                    Or (AscW(VBA.Left(C, 1)) >= &H2B820 And AscW(VBA.Left(C, 1)) <= &H2CEAF) _
                    Or (AscW(VBA.Left(C, 1)) >= &HF900 And AscW(VBA.Left(C, 1)) <= &HFAFF) _
                    Or (AscW(VBA.Left(C, 1)) >= &H2F800 And AscW(VBA.Left(C, 1)) <= &H2FA1F) Then '這裡沒取碼點，必定有誤，待改寫！！！！！！！！
                    If Not r.font.Name = nmf.Name And Not r.font.Name = kff.Name Then '運用之原理在此行！！！！
                        ' 如果字符不在新細明體或標楷體字型中，則將其字體更改為HanaMinB
                        r.font.Name = "HanaMinB"
                    End If
                End If
            End If
        Next r
    Next p
End Sub

Sub updateURL() '更新超連結網址
Dim site As String
Dim lnk As New Links
site = InputBox("what site to update?", , "漢語大詞典=1;國語辭典=2;國學大師=3")
If site = "" Then Exit Sub
Select Case site
    Case 1 '"漢語大詞典"
        lnk.updateURL漢語大詞典 ActiveDocument
    Case 2 '"國語辭典"
        lnk.updateURL國語辭典 ActiveDocument
    Case 3 '"國學大師"
        lnk.updateURL國學大師 ActiveDocument
        
End Select
End Sub



