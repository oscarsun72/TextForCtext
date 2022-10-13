Attribute VB_Name = "Docs"
Option Explicit
Public d字表 As Document, x As New EventClassModule '這才是所謂的建立"新"的類別模組--實際上是建立對它的參照.'原照線上說明乃Dim也.

Public Sub Register_Event_Handler() '使自設物件類別模組有效的登錄程序.見「使用 Application 物件 (Application Object) 的事件」
    Set x.app = word.Application '此即使新建的物件與Word.Application物件作上關聯
End Sub
Sub 空白的新文件() '20210209
Dim a As Document, flg As Boolean
If Documents.Count = 0 Then GoTo a:
If ActiveDocument.Characters.Count = 1 Then
    Selection.Paste
ElseIf ActiveDocument.Characters.Count > 1 Then
    For Each a In Documents
        If a.path = "" Or a.Characters.Count = 1 Then
            a.Range.Paste
            a.Activate
            a.ActiveWindow.Activate
            flg = True
            Exit For
        End If
    Next a
    If flg = False Then GoTo a
Else
a: Documents.Add
    Selection.Paste
End If
End Sub


Sub 在本文件中尋找選取字串_迅速() '指定鍵:Alt+Ctrl+Down 2015/11/1
Static x As String
With Selection
    If .Type = wdSelectionNormal Then
        x = 文字處理.trimStrForSearch(.Text, Selection)
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
        If Not Left(h.Sections(i).Range.Paragraphs(2).Range.Text, Len(h.Sections(i).Range.Paragraphs(2).Range.Text) - 1) Like "..." Then
            rst.FindFirst "頁 = " & Left(h.Sections(i).Range.Paragraphs(1).Range.Text, Len(h.Sections(i).Range.Paragraphs(1).Range.Text) - 1) _
                & "and 札記 like '" & "*" & Left(h.Sections(i).Range.Paragraphs(2).Range.Text, _
                    Len(h.Sections(i).Range.Paragraphs(2).Range.Text) - 1) & "*'"
    '            & "and 札記 like '" & "*" & Replace(h.Sections(i).Range.Paragraphs(2).Range.Text, Chr(12), "") & "*'"
                'chr(12)暫不知是什麼(應是分節符號!),但會影響比對,故予取代為空字串 _
                因為最後一節沒有分節符號(Chr(12))而是段落符號(Chr(13),故若以Replace函數須分別處理 _
                為免麻煩,一律用Left函數不取最右方之字元即可(不管是Chr(12)或Chr(13))
        Else '札記為""時的處理
            rst.FindFirst "頁 = " & Left(h.Sections(i).Range.Paragraphs(1).Range.Text, Len(h.Sections(i).Range.Paragraphs(1).Range.Text) - 1) _
                & "and 札記 = """"" '在此,不用CSng型態轉換是可以的! _
                因為頁有小數點,故在作頁比對時,不能用Words物件(會將小數點之數字分開算成不同的Word),如果頁沒有小數點的話就可以了! _
                減一,同札記,是為了剔除最右方的Chr(10)(換行符號、段落符號)
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
Dim f, i As Integer
f = Array("·", "‧", "。", "」", Chr(-24152), "：", "，", "；", _
    "、", "「", ".", Chr(34), ":", ",", ";", _
    "……", "...", "．", "【", "】", " ", "《", "》", "〈", "〉", "？" _
    , "！", "﹝", "﹞", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0" _
    , "『", "』", Chr(13), ChrW(9312), ChrW(9313), ChrW(9314), ChrW(9315), ChrW(9316) _
    , ChrW(9317), ChrW(9318), ChrW(9319), ChrW(9320), ChrW(9321), ChrW(9322), ChrW(9323) _
    , ChrW(9324), ChrW(9325), ChrW(9326), ChrW(9327), ChrW(9328), ChrW(9329), ChrW(9330) _
    , ChrW(9331), ChrW(8221), """") '先設定標點符號陣列以備用
    '全形圓括弧暫不取代！
    'a = ActiveDocument.Content
'    Set a = ActiveDocument.Range.FormattedText '包含格式化的資訊
    For i = 0 To UBound(f)
        'a = Replace(a, F(i), "")
        ActiveDocument.Range.Find.Execute f(i), True, , , , , , wdFindContinue, True, "", wdReplaceAll
    Next
    'ActiveDocument.Content = a
End Sub
Sub 注腳符號() '注釋符號、註釋符號、註腳符號
Dim i As Integer
For i = 9312 To 9331
    With Selection.Range.Find
        .Replacement.Font.Name = "Arial Unicode MS"
        .Execute ChrW(i), , , , , , , wdFindContinue, , ChrW(i), wdReplaceAll
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
If Selection.Type = wdSelectionNormal And Right(Selection, 1) Like Chr(13) Then _
            Selection.MoveLeft wdCharacter, 1, wdExtend '不要包含分段符號!
If Selection.Style <> "引文" Then Selection.Style = "引文" '如果不是引文樣式時,則改成引文樣式
s = Selection.start '記下起始位置
Selection.PasteSpecial , , , , wdPasteText '貼上純文字
e = Selection.End '記下貼上後的結束位置
Selection.SetRange s, e
Set r = Selection.Range
With r
    r.Find.Execute Chr(13), , , , , , , wdFindStop, , Chr(11), wdReplaceAll '將分行符號改成手動分行符號
End With
r.Footnotes.Add r '插入註腳!
End Sub
Sub 貼上純文字() 'shift+insert 2016/7/20
Dim hl, s As Long, r As Range
On Error GoTo ErrHandler
hl = Selection.Range.HighlightColorIndex

s = Selection.start
Set r = Selection.Range
'Selection.PasteSpecial , , , , wdPasteText '貼上純文字
Selection.PasteAndFormat (wdFormatPlainText)
r.SetRange s, Selection.End
If hl <> 9999999 Then r.HighlightColorIndex = hl
Exit Sub
ErrHandler:
Select Case Err.Number
    Case 5342 '指定的資料類型無法取得。
        
    Case Else
        MsgBox Err.Number & Err.Description
End Select
End Sub
Sub 一字一段()
With Selection
    .HomeKey wdStory
    Do Until .End = .Document.Range.End - 1
        .MoveRight
        .TypeText Chr(13)
    Loop
End With
End Sub
Sub OCR表格處理()
Dim a, b, i As Byte
a = Array("┴", Chr(13) & "│", " ", "─", "┼", "├", "┤", "┐", "┌", "└", "┘", "┬", "│", Chr(9) & Chr(13), Chr(13) & Chr(13), Chr(13) & Chr(9))
b = Array("", Chr(13), "", "", "", "", "", "", "", "", "", "", Chr(9), Chr(13), Chr(13), Chr(13))
With ActiveDocument
    If .path = "" Then
        For i = 0 To UBound(a) - 1
            With .Range.Find
                .Text = a(i)
                With .Replacement
                    .Text = b(i)
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
            x = Mid(x, InStrRev(x, "\") + 1)
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
            If Selection.Next.Font.Size > 12 Then '目錄預設為12字形
                Selection.MoveRight
                Do Until Selection.Next.Font.Size = 12
                    Selection.MoveRight , , wdExtend
                Loop
                If MsgBox("是否要縮小為10號字?", vbQuestion + vbOKCancel) = vbOK Then Selection.Font.Size = 10
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
    FdText = 文字處理.trimStrForSearch(.Text, Selection)
    st = .start: ed = .End
    .Collapse wdCollapseEnd
    MnText = .Document.StoryRanges(wdMainTextStory) '變數化處理較快2003/4/8
'    MnText = ActiveDocument.Range '2010/2/5
    ins(1) = InStr(MnText, FdText)
    ins(2) = InStrRev(MnText, FdText)
    If .Document.Footnotes.Count > 0 Then '有註腳才檢查2003/4/3
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
                    .Text = FdText
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
                        .Text = FdText
'                        .Forward = True
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
        With .Find
            .ClearFormatting
            .Replacement.ClearFormatting '這也要清除才行
            .Forward = True
            .Wrap = wdFindAsk
            .MatchCase = True
            .Text = FdText '.Parent.Text
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
        ActiveDocument.ActiveWindow.Selection.Find.Execute Selection.Text
    Case Else
        MsgBox Err.Number & Err.Description
        Resume
End Select
End Sub


Sub 書籤_以選取文字作為書籤() 'ALT+SHIFT+B

' 巨集錄製於 2015/9/20，錄製者 王觀如
    With ActiveDocument.bookmarks
        .Add Range:=Selection.Range, Name:=Replace(Selection.Text, Chr(13), "")
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
        p.Range.Font.Hidden = True
    Else
        prngTxt = p.Range.Text
        s = InStr(prngTxt, " ")
        If VBA.Len(VBA.Mid(prngTxt, s + 1)) < 5 Then
            p.Range.Font.Hidden = True
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
    If Not p.Range.Font.Hidden Then
        prngTxt = p.Range.Text
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
                    Debug.Print p.Next.Range.Text
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
    x = Selection.Text
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
        If Not p.Range.Font.Hidden Then 'prepare to delete
            'p.Range.Cut
            x = p.Range.Text
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
'        If .Paragraphs(.Paragraphs.Count).Range <> Chr(13) Then .InsertParagraphAfter
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
        With .Font
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
        With .Font
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

Sub mark易學關鍵字()
Dim searchedTerm, e, ur As UndoRecord, d As Document, clipBTxt As String, flgPaste As Boolean, xd As String
'Set ur = SystemSetup.stopUndo("mark易學關鍵字")
SystemSetup.stopUndo ur, "mark易學關鍵字"
Set d = ActiveDocument
clipBTxt = VBA.Trim(SystemSetup.GetClipboardText)
searchedTerm = Array("卦", "爻", "易", "周易", "易經", "系辭", "繫辭", "擊辭", "擊詞", "繫詞", "說卦", "序卦", "卦序", "敘卦", "雜卦", "文言", "乾坤", "無咎", ChrW(26080) & "咎", "天咎", "元亨", "利貞" _
    , "史記", "九五", "六二", "上九", "上六", "九二", "筮") ', "", "", "", "")

'If Selection.Type = wdSelectionIP Then
    If InStr(d.Range, clipBTxt) = 0 Then
        For Each e In searchedTerm
            If InStr(clipBTxt, e) > 0 Then
                flgPaste = True
                Exit For
            End If
        Next e
        If Not flgPaste Then
            'ChrW() & ChrW() &'ChrW() & ChrW() &
            Dim gua
            gua = Array(ChrW(-10119), ChrW(-8742), ChrW(-30233), ChrW(-10164), ChrW(-8698), ChrW(-31827), ChrW(-10132), ChrW(-8313), ChrW(20810), ChrW(-10167), ChrW(-8698), ChrW(-26587), ChrW(21093), ChrW(14615), ChrW(20089), ChrW(26080), "妄", ChrW(26083), "濟" _
                        , "遘", "遁", ChrW(20089), "离", "乾", "小畜", "履", "臨", "觀", "大過", "坤", "泰", "否", "噬嗑", "賁", "坎", "屯", "蒙", "同人", "大有", "剝", "復", "離", "需", "訟", "謙", "豫", "無妄", "大畜", "師", "比", "隨", "蠱", "頤", "咸", "恒", "損", "益", "震", "艮", "中孚", "遯", "大壯", "夬", "姤", "漸", "歸妹", "小過", "晉", "明夷", "萃", "升", "豐", "旅", "既濟", "未濟", "家人", "睽", "困", "井", "巽", "兌", "蹇", "解", "革", "鼎", "渙", "節")
            For Each e In gua
                If InStr(clipBTxt, e) > 0 Then
                    flgPaste = True
                    Exit For
                End If
            Next e
        End If
        
        If flgPaste Then
            Selection.EndKey wdStory
            Selection.InsertParagraphAfter
            Selection.InsertParagraphAfter
            Selection.Collapse wdCollapseEnd
            貼上純文字
            Selection.InsertParagraphAfter: Selection.InsertParagraphAfter: Selection.InsertParagraphAfter
            Selection.Collapse wdCollapseEnd
            ActiveWindow.ScrollIntoView Selection
        End If
    Else
        Dim rng As Range
        Set rng = d.Range
        rng.Find.Execute VBA.Left(clipBTxt, 255), , , , , , , wdFindContinue
    End If
'End If
If flgPaste Then
    word.Application.ScreenUpdating = False
    If d.path <> "" And Not d.Saved Then d.Save
    xd = d.Range.Text
    For Each e In searchedTerm
        If InStr(xd, e) > 0 Then
            With d.Range.Find
                .ClearFormatting
                .Text = e
                With .Replacement
                    .Text = e
                    .Font.ColorIndex = wdRed
                    .Highlight = True
                End With
                .Execute , , , , , , True, wdFindContinue, , , wdReplaceAll
            End With
        End If
    Next e
    d.Range.Find.ClearFormatting

    GoSub refres
    SystemSetup.playSound 1.921
Else
    GoSub refres
    SystemSetup.playSound 1.469
End If
SystemSetup.contiUndo ur
Set ur = Nothing
Exit Sub
refres:
    word.Application.ScreenUpdating = True
    If Not rng Is Nothing Then rng.Select
    word.Application.ScreenRefresh
    ActiveWindow.ScrollIntoView Selection, True

Return
End Sub


Sub 抽取超連結位址()
Dim hplnk As Hyperlink, x As String, d As Document
For Each hplnk In ActiveDocument.Hyperlinks
    x = x & hplnk.Address & Chr(13)
Next hplnk
Set d = Documents.Add
d.Range.Text = x
d.Range.Cut
d.Close wdDoNotSaveChanges
End Sub


Sub 插入超連結_文件中的位置_標題() 'Alt+P 原是「引詩」樣式'2021/11/27
Dim d As Document, title As String, p As Paragraph, pTxt As String, subAddrs As String, flg As Boolean
Set d = ActiveDocument
title = Selection.Text
title = 文字處理.trimStrForSearch(title, Selection)
For Each p In d.Paragraphs
    If Left(p.Style.NameLocal, 2) = "標題" Then
        pTxt = p.Range.Text
        pTxt = Left(pTxt, Len(pTxt) - 1)
        If StrComp(pTxt, title) = 0 Then
            subAddrs = title
            flg = True
            Exit For
        ElseIf InStr(pTxt, title) > 0 Then
            subAddrs = "_" & Mid(pTxt, 1, InStrRev(pTxt, " ") - 1)
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

Sub 中國哲學書電子化計劃_只保留正文注文_且注文前後加括弧_貼到古籍酷自動標點()
Dim ur As UndoRecord
SystemSetup.stopUndo ur, "中國哲學書電子化計劃_註文前後加括弧_貼到古籍酷自動標點"
中國哲學書電子化計劃.只保留正文注文_且注文前後加括弧
貼到古籍酷自動標點
SystemSetup.contiUndo ur
End Sub

Sub 貼到古籍酷自動標點()
Dim d As Document
Set d = ActiveDocument
If d.path <> "" Then Exit Sub
d.Range.Cut
On Error GoTo app
AppActivate "古籍酷"
DoEvents
SendKeys "{TAB 16}", True
SendKeys "^v"
DoEvents
SendKeys "+{TAB 2}~", True
If d.path = "" Then d.Close wdDoNotSaveChanges
Exit Sub
app:
Select Case Err.Number
    Case 5
        Shell (Network.getDefaultBrowserFullname + " https://old.gj.cool/gjcool/index")
        AppActivate Network.getDefaultBrowserNameAppActivate '"古籍酷"
        DoEvents
        SystemSetup.Wait 2.5
        'SendKeys "{TAB 16}", True
        Resume Next
    Case Else
        MsgBox Err.Number & Err.Description
End Select
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



