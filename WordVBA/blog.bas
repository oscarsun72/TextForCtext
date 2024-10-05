Attribute VB_Name = "blog"
Option Explicit '奇摩部落格專用模組
Public OX As Object, htmFilename As String, myaccess As Object 'As Access.Application
Dim es As Byte '記下頁差2007/11/1
Sub setOX()
On Error Resume Next
If OX Is Nothing Then Set OX = CreateObject("AutoItX3.Control")
End Sub
Sub 儲存供google索引並備分(dp As Document)  '2008/12/24
    htmFilename = InputBox("請輸入檔名", , "htmfilename")
    If htmFilename = "" Then Exit Sub
    On Error GoTo eH:
    dp.SaveAs fileName:= _
        "P:\我的部落格\5160_\" & VBA.Left(htmFilename, 235) & ".html", _
        FileFormat:=wdFormatUnicodeText, LockComments:=False, Password:="", _
        AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
        EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
        :=False, SaveAsAOCELetter:=False
'    Windows("復初齋詩集（一）(588頁)-卷25(秘閣直廬集上（壬寅三月至十二月）壬寅 乾隆47年.1782年.先生年50歲)"). _
'        Activate
'    Windows("復初齋詩集（一）(588頁)-卷25(秘閣直廬集上（壬寅三月至十二月）壬寅 乾隆47年.1782年.先生年50歲)"). _
'        Activate
    setOX
    OX.WinActivate "iexplorer"
    'AppActivate "iexplorer"
Exit Sub
eH:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " - " & Err.Description
        htmFilename = InputBox("請輸入檔名", , htmFilename)
        Resume
End Select
End Sub


Sub 插入相簿圖片連結網址() '部落格用'2006/4/4
Dim x As String, i As Long 'Integer
x = InputBox("請輸入第1個圖片的連結網址")
'If x = "" Then End 'Exit Sub
If x = "" Or InStr(x, "http") = 0 Then End
If InStr(x, "&prev") > 0 Then x = VBA.Left(x, InStr(x, "&prev") - 1)
i = Int(VBA.Mid(x, InStrRev(x, "=") + 1))
x = VBA.Left(x, InStrRev(x, "="))
options.AutoFormatAsYouTypeReplaceQuotes = False '關掉智慧引號,如此"才不會被自動置換為”（在自動校正裡依照您的輸入自動格式頁籤裡）
With ActiveDocument.Range
    With .Find
        .MatchWildcards = False
        .ClearFormatting
        .text = "<img src="
        .Forward = True
        .Wrap = wdFindStop
        .Replacement.text = "<a href=""" & x & i & """><img src="
        Do While .Execute(, , , , , , , , , , wdReplaceOne)
            .Parent.Move
            i = i + 1
            With .Replacement
                .text = "<a href=""" & x & i & """><img src="
            End With
        Loop
        .Parent.Move wdStory
        .Forward = False
        .text = ".jpg"" />"
        .Replacement.text = ".jpg"" /></a>"
        .Execute , , , , , , , , , , wdReplaceAll
    End With
End With
options.AutoFormatAsYouTypeReplaceQuotes = True '恢復智慧引號
End Sub


Sub 圖片插入頁後指定位置()
Dim x As Range, s As String, e As String, t As String, sy As Boolean
Static sn As Long '記下末頁
Dim d As Document ', dp As Document
Dim ts As Byte
Set d = ActiveDocument
'Set dp = d.Windows(1).Previous.Document
On Error GoTo errhan:
With d.Windows(1).Previous.Document '自動關閉之前打開的新文件(表示部落格上傳已成功)2007/10/28
    If .path = "" Then .Close wdDoNotSaveChanges
'    If d.Application.Documents.Count > 2 Then
'        If .Path = "" Then
'            If .Range = vba.Chr(13) Then .Undo  '還原剪下前
'            If InStr(.Range, "<p><a href=""http://tw.myblog.yahoo.com/") Then
'                Set dp = d.Windows(1).Previous.Document
'                儲存供google索引並備分 dp
'            End If
'            .Close wdDoNotSaveChanges
'        End If
'    End If
End With
d.Activate
s = InputBox("請輸入第1頁", , sn + 1)
If s = "" Then Exit Sub '以10頁為單位
If InStr(s, ".") Then sy = True: ts = 1 '由起始判斷是否為"葉" 2009/11/13
's = CInt(s)
s = CSng(s)
'e = s + 10
If es = 0 Then es = 9
e = InputBox("請輸入最後1頁", , s + es) ' + 9)
If sy = False Then If InStr(e, ".") Then sy = True: ts = 1 '由起始判斷是否為"葉" 2009/11/13
If e = "" Then Exit Sub
'e = CInt(e)
e = CSng(e)
If e <> 0 Then sn = e: es = e - s
'If e = 0 Then e = s + 9 '省略即以10頁為單位
's = 261: e = 270
With d.Range
    Do Until s >= e + 0.1 '+ 1
'    s = e '倒序時請用這兩行
'    Do Until s = e - 10
'        .move wdStory
        Set x = d.Range(InStr(d.Range, "<img src=""") - 1, InStr(d.Range, ".jpg"" />") + Len(".jpg"" />") - 1)
        t = x.text
        .SetRange InStr(d.Range, "<img src=""") - 1, InStr(d.Range, ".jpg"" /><br>") + Len(".jpg"" /><br>") - 1
        .Delete
        With Selection
            .HomeKey wdStory
            .Find.ClearFormatting
            .Find.MatchWildcards = False
            .Find.Forward = True
            'If .Find.Execute(">" & s & "<") Then
            If .Find.Execute("<p class=MsoNormal style=""MARGIN: 0cm 0cm 0pt""><strong><span lang=EN-US style=""FONT-SIZE: 16pt; BACKGROUND: #efef5a""><font face=""Times New Roman"">" _
                    & s & "</font></span></strong></p>") Then
                .Move wdParagraph
                .InsertAfter "<p>" & t & "</p>"
            ElseIf .Find.Execute("<p class=MsoNormal style=""MARGIN: 0cm 0cm 0pt""><strong><span lang=EN-US style=""FONT-SIZE: 16pt;  COLOR: blue""><font face=""Times New Roman"">" _
                    & s & "</font></span></strong></p>") Then
                .Move wdParagraph
                .InsertAfter "<p>" & t & "</p>"
            ElseIf .Find.Execute("<p class=""MsoNormal"" style=""MARGIN:0cm 0cm 0pt;""><strong><span style=""FONT-SIZE:16pt;COLOR:blue;""><font face=""Times New Roman"">" _
                    & s & "</font></span></strong></p>") Then
                .Move wdParagraph
                .InsertAfter "<p>" & t & "</p>"
            ElseIf .Find.Execute("<p class=MsoNormal style=""MARGIN: 0cm 0cm 0pt""><strong><span lang=EN-US style=""FONT-SIZE: 16pt; COLOR: blue""><font face=""Times New Roman"">" _
                    & s & "</font></span></strong></p>") Then
                .Move wdParagraph
                .InsertAfter "<p>" & t & "</p>"
            ElseIf .Find.Execute(">" & s & "<o:p></o:p></font></span></b></p>") Then
'            If .Find.Execute("<span lang=EN-US style=""FONT-SIZE: 16pt; COLOR: blue; mso-bidi-font-size: 12.0pt; mso-text-animation: ants-red""><font face=""Times New Roman"">" & s & "<") Then '前行原式會在有出現數字頁時錯亂!2006/8/15
                .Move wdParagraph
                .InsertAfter "<p>" & t & "</p>"
            ElseIf .Find.Execute(">" & s & "<?xml:namespace prefix = o") Then
                .Move wdParagraph
                .InsertAfter "<p>" & t & "</p>"
            ElseIf .Find.Execute(">" & s & "<") Then
                .Move wdParagraph
                .InsertAfter "<p>" & t & "</p>"
'            Else
'                Exit Do
            End If
        End With
        If Not sy Then
            s = s + 1 '非葉時
        Else
'葉時:       If ts Mod 2 = 1 Then '葉時
'                s = s - 0.1 + 1
'            Else
'                s = s + 0.1
'            End If
'葉時end:        ts = ts + 1
            If InStr(s, ".") = 0 Then
                s = s + 0.1
            Else
                s = s - 0.1 + 1
            End If
        End If
'        s = s - 1 '倒序時請用這一行
    Loop
End With
Exit Sub
errhan:
Select Case Err.Number
    Case 91 '沒有設定物件變數或 With 區塊變數
        Resume Next '表示之前沒有開啟的新文件.
End Select
End Sub
Sub 排列圖片後插入連結()
If ActiveDocument.path <> "" Then Documents.Add DocumentType:=wdNewBlankDocument
options.AutoFormatAsYouTypeReplaceQuotes = False '恢復智慧引號
With ActiveDocument.Range
    .Select
    .Paste
End With
圖片插入頁後指定位置
插入相簿圖片連結網址
ActiveDocument.Range.Copy
options.AutoFormatAsYouTypeReplaceQuotes = True '恢復智慧引號
End Sub
Sub 取代為新的圖片網址()
With ActiveDocument.Range
'    With .Find
'        .Text = "src="""
'        .Forward = True
'        Do Until .Execute = False
'
'            With .Parent
'                .move
'                .MoveUntil .Text = "h", wdExtend
'            End With
'        Loop
'    End With
End With
End Sub

Sub 輸入上傳圖片位址()
Dim x As String, i As Byte
Static sn As Long '記下末頁
x = InputBox("請輸入第一個圖片的頁碼", , sn + 1)
If x = "" Then Exit Sub
If IsNumeric(x) Then sn = x + es
'AppActivate "avant browser"
AppActivate "explorer"
'AppActivate "mozilla firefox"'mozilla firefox不能用
DoEvents
SendKeys "{tab 3}" & 取得桌面路徑 & "\測試用\變更檔名用\" & Format(x, "_000000") & ".jpg"
For i = 1 To 9 '從第2個圖片到第10個
    DoEvents
    SendKeys "{tab 4}" & 取得桌面路徑 & "\測試用\變更檔名用\" & Format(x + i, "_000000") & ".jpg"
Next i
DoEvents
'SendKeys "{tab 3}{right}"'隱藏
SendKeys "{tab 3}"
End Sub
Sub 全選剪下後關閉文件() '以利貼上部落格也'2007/10/30-貼上上頁下頁書首回總目的html碼.
Dim Dnow As Document, bt As String, hide As Boolean
With Selection
    If ActiveDocument.path = "" Then
        Set Dnow = ActiveDocument
        If InStr(Dnow.Range, "<a href=") Then '表示要貼上html碼了!
            Dim o As Boolean, d, h As String
            For Each d In Documents
                If d.Name = "暫存.doc" Then o = True: Exit For
            Next
            If o = False Then
                Documents.Open 取得桌面路徑 & "\暫存.doc"
            Else
                Documents("暫存").Activate
            End If
            o = True '記下已是在處理html碼,以供下面用.2007/11/4
            With ActiveWindow.Selection
                If Len(.text) = 1 Then .GoTo wdGoToBookmark, , , "游標_暫存"
                bt = .Range
                h = InputBox("請輸入上文網址")
                'If h = "" Or InStr(h, "http") = 0 Then Exit Sub
                If h = "" Then Exit Sub
                If InStr(h, "http") <> 0 Then
                    If InStr(h, "&prev") > 0 Then h = VBA.Left(h, InStr(h, "&prev") - 1)
                    bt = Replace(bt, "上文", "<a href=""" & h & """>" & "上文</a>") '插入上文網址
                End If
                '.Parent.WindowState wdWindowStateMinimize
                ActiveWindow.Visible = False
            End With
            With Dnow
                .Activate
                .Range = bt & .Range & bt
            End With
        End If
'        If InStr(.Document.Range, ": 標楷體""") Then .Document.Range = 去標楷體字(.Document.Range)
        If InStr(.Document.Range, "<img") Then '判斷式不可省,否則會在非源碼時執行2009/4/17
            .Document.Range = 去標楷體字(.Document.Range)
            .EndKey wdStory, wdMove
            Do Until Asc(.Previous) <> 13
                .TypeBackspace
            Loop
             '在頁首插入在線人數
            .Document.Range = "<p><a href=""http://whos.amung.us/stats/s5z4puepm2vb/""><img title=""Click to see how many people are online"" src=""http://whos.amung.us/widget/s5z4puepm2vb.png"" border=""0"" height=""29"" width=""81"" /></a></p> " & .Document.Range
            '.Document.Range = .Document.Range & "<p style=""text-align: right;""><a href=""""><span style=""color: red;"">請多指教(comments) <span style=""color: rgb(51, 102, 255);""><br><font size=1>不須註冊,不必登入,<br>可匿名留言</font></span></span></a></p>"
            '.Document.Range = .Document.Range & "<p style=""text-align: right;""><a href=""http://www.blogger.com/comment.g?blogID=37481082&amp;postID=116317965718083160""><span style=""color: red;"">歡迎指教(comments) <span style=""color: rgb(51, 102, 255);""><br><font size=1>不須註冊,不必登入,<br>可匿名留言</font></span></span></a></p>"
            .Document.Range = .Document.Range & "<p style=""text-align: right;""><a href=""http://www.blogger.com/comment.g?blogID=37481082&amp;postID=116317965718083160""><span style=""color: red;"">沒有東西是平白得來的.質能守恆,請不要虧欠太多.<br>這本來就是給有心讀書,且真讀書的好學者修學者交流的平台,不是供學術或學生取巧功名的苟且.<br>這是助人,而不是害人的帖子.更不是我來放高利貸的功德. <br>請多給予指教或打氣(comments) 誠心交流,互不相欠,衡情斟酌.<br>否則寧願您去付費的所在,銀貨兩訖,也好償贖. <br>感謝你,也救贖了自己.坦蕩平夷,生生世世平安.<br>當初真沒想到,這種教育的話還要留給我今天來說.何況是有閒情看覽此頁者.<br>我們無須爭辯神存不存在,但是,死,存在<span style=""color: rgb(51, 102, 255);""><br><font size=1>不須註冊,不必登入,<br>點擊此處即可匿名留言<br>留言前,請仰望天,或看看自己將去的冥年<br>不留言,只有你會忘了,我會不知,事實會永存</font></span></span></a></p>"
        End If
        .WholeStory
        .Cut
'        .Document.Close wdDoNotSaveChanges'2007/11/2因為奇摩常當,故不關閉文件以便還原
'        If o Then
'            If MsgBox("是否隱藏?", vbOKCancel) = vbOK Then
'                hide = True
'            Else
                hide = False
'            End If
'        End If
'        AppActivate "Avant Browser"
        AppActivate "explorer"
'        AppActivate "mozilla firefox"
'        SendKeys "+{insert}"
        SendKeys "2"
        If o Then '2007/11/4
            DoEvents
            If Not hide Then
                SendKeys "{tab 3}" '公開貼子
            Else
                SendKeys "{tab 3}{right}" '隱藏貼子
            End If
            SendKeys "{tab 4}{enter}" '發表貼子
        End If '2007/11/4
    Else
'        .Document.Close wdDoNotSaveChanges
    End If
End With

End Sub

Sub 複製貼上書首資訊()
Dim o As Boolean, d
For Each d In Documents
    If d.Name = "暫存" Then o = True: Exit For
Next
If o = False Then
    Documents.Open 取得桌面路徑 & "\暫存.doc"
Else
    Documents("暫存").Activate
End If
With ActiveWindow.Selection
    If Len(.text) = 1 Then .GoTo wdGoToBookmark, , , "游標_暫存"
    .Copy
    '.Parent.WindowState wdWindowStateMinimize
    ActiveWindow.Visible = False
'    AppActivate "Avant Browser"
    AppActivate "explorer"
    DoEvents
    SendKeys "+{insert}"
End With
End Sub

Sub 產生書首頁碼()
Static a As String, e As String ', s As String
Dim i As Long, d As Document
a = InputBox("請輸入間隔頁數", "產生書首頁碼", 10)
If a = "" Then Exit Sub
i = InputBox("請輸入起始頁碼", "產生書首頁碼", 1)
'If i = "" Then Exit Sub
e = InputBox("請輸入結束頁碼", "產生書首頁碼")
If e = "" Then Exit Sub
e = VBA.StrConv(e, vbNarrow)
'i = VBA.StrConv(i, vbNarrow)
'a = VBA.StrConv(a, vbNarrow)
Set d = Documents.Add
With d
    Do Until i > e
        If i + a > e Then
            .Range = .Range & i & "-" & e & "書末 "
        Else
            .Range = .Range & i & "-" & i + a - 1 & " "
        End If
        i = i + a
    Loop

.Range = Replace(d.Range, VBA.Chr(13), "")
.Range.font.Size = 8
End With
End Sub

Sub 重排圖序() '2008/7/7 將第1張置于最後一張,依此類推
Dim i As String, x As String, d As Document, r As Range
Dim s As String, e As String, sp As Long, eP As Long
Set d = Documents.Add
s = "<img src="
e = "/><br>"
With d
    .Range.Paste
    Do
        x = .Range.text
        eP = InStr(x, e) + Len(e) - 1
        sp = InStr(x, s)
        If sp = 0 Then Exit Do
        Set r = d.Range(0, eP)
        i = r & i
        r.Delete
    Loop
    .Range = i
    .Range.Select
    .Range.Cut
    .Close wdDoNotSaveChanges
End With
'setOX
'OX.WinActivate "explorer"
AppActivate "explorer"
End Sub

Sub 開啟超連結()
'Ctrl + i   系統預設是 italic（即斜體字），此配合 ExcelVBA設定
Dim rng As Range
Set rng = Selection.Range
If rng.Hyperlinks.Count = 0 Then '如果所在位置沒有超連結，則看其前有否；若又無，則再看其後有否；若都無則不執行 2022/12/20
    If rng.start = 0 Then
        If Selection.Type = wdSelectionNormal Then
            GoTo Selected_Range
        Else
            GoSub nxt
        End If
    ElseIf rng.End = rng.Document.Range.End - 1 Then
        GoSub pre
    ElseIf rng.End < rng.Document.Range.End - 1 Then
Selected_Range:
        '為 「生難字加上國語辭典注音」 而設
        Dim rngNext As Range
        Set rngNext = rng.Next
        If rngNext.text = "（" Then
            If rngNext.Next.Hyperlinks.Count > 0 Then
                rng.SetRange rngNext.Next.start, rngNext.Next.End
                GoSub nxt
            End If
        Else
            GoSub Position
        End If
    Else
        GoSub Position
    End If
End If
If rng.Hyperlinks.Count > 0 Then
    Dim strLnk As String, lnk As Hyperlink
    Set lnk = rng.Hyperlinks(1)
    If rng.Hyperlinks(1).SubAddress <> "" Then
        Dim subAdrs As String
        subAdrs = rng.Hyperlinks(1).SubAddress
        strLnk = rng.Hyperlinks(1).Address + "#" + _
            VBA.IIf(VBA.InStr(subAdrs, "%"), subAdrs, _
            IIf(code.IsSurrogate(subAdrs), UrlEncode(subAdrs), code.UrlEncode_Big5UnicodOLNLY(subAdrs)))
    Else
        strLnk = rng.Hyperlinks(1).Address
    End If
    SystemSetup.playSound 0.484
    Shell getDefaultBrowserFullname + " " + strLnk + " --remote-debugging-port=9222 "
End If
Exit Sub
Position:
    If rng.Previous.Hyperlinks.Count > 0 Then
pre:        Set rng = rng.Previous
    ElseIf rng.Next.Hyperlinks.Count > 0 Then
nxt:        Set rng = rng.Next
    End If
Return
End Sub
Sub 插入超連結() '2008/9/1 指定鍵(快捷鍵) Ctrl+shift+K(原系統指定在smallcaps為)
'Alt+k
'    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
'    Selection.Range.Hyperlinks(1).Range.Fields(1).Result.Select
'    Selection.Range.Hyperlinks(1).Delete
'setOX
'    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:= _
        VBA.StrConv(OX.ClipGet, vbNarrow) _
        , SubAddress:=""
'    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:= _
        VBA.StrConv(GetClipboard, vbNarrow) _
        , SubAddress:="" '    Selection.Collapse Direction:=wdCollapseEnd
        
    Dim lnk As String
    'lnk = UrlEncode(SystemSetup.GetClipboardText)
    lnk = SystemSetup.GetClipboardText
    If VBA.InStr(lnk, "http") = 0 Or VBA.InStr(lnk, "http") > 1 Then MsgBox "剪貼簿中非有效網址！": Exit Sub
            
    Dim rng As Range, b As Boolean, ur As UndoRecord ', wndo As Window ', d As Document ', sty As String
    
    
    Set rng = Selection.Range ': Set wndo = ActiveWindow
    If rng.Information(wdInFootnote) Then
        If rng.Document.Windows.Count > 1 Then '因為開多視窗時若在第1個以外的視窗的註腳中執行「rng.Hyperlinks.Add」則會誤插到第1個視窗體中
            Dim wnd As Window, ww, i As Byte
            Set wnd = ActiveWindow
            If CByte(VBA.Right(wnd.Caption, 1)) > 1 Then
                Dim wnds() As Object
                For Each ww In rng.Document.Windows
                    ReDim Preserve wnds(i)
                    Set wnds(i) = ww
                    i = i + 1
                Next
            End If
        End If
    End If
    'Set ur = SystemSetup.stopUndo("")
     SystemSetup.stopUndo ur, ""
    'Set d = ActiveDocument
'    If d.path <> "" Then d.Save
    b = rng.Bold ': sty = rng.Style
    'wndo.Activate 在所開多視窗中似乎只會插到第一個視窗的文件位置
    If i > 0 Then
        Dim slRng() As Object
        i = 0
        For Each ww In wnds
            If Not ww Is wnd Then '不能用 ww.Caption <> wnd.Caption 因為下面有ww.Close 視窗一旦關閉，則Caption屬性也會異動
                ReDim Preserve slRng(i)
                Set slRng(i) = ww.Selection.Range
                ww.Close '既然不能改變視窗前後，就只能先關閉，且記下其游標所在位置了
                i = i + 1
            End If
        Next
    End If
    Dim ssharp As String
    ssharp = InStr(lnk, "#")
    If ssharp > 0 Then
        Dim w As String
        lnk = VBA.Replace(lnk, VBA.ChrW(-9217) & VBA.ChrW(-8195), "　")
        w = VBA.Mid(lnk, ssharp + 1, Len(lnk) - ssharp)
        w = code.UrlEncode(w)   'byRef
        lnk = VBA.Mid(lnk, 1, ssharp) + w
    End If
    rng.Hyperlinks.Add Anchor:=rng, Address:= _
        VBA.StrConv(lnk, vbNarrow) _
        , SubAddress:="", Target:="_blank" '    Selection.Collapse Direction:=wdCollapseEnd
    'wndo.Activate
    If rng.Bold <> b Then rng.Bold = b
    'If rng.Style <> sty Then rng.Style = sty
    SystemSetup.contiUndo ur
    Set ur = Nothing
    If i > 0 Then
        i = 0
        For Each ww In wnds
            If Not ww Is wnd Then
                Dim wwp As Window
                Set wwp = rng.Document.Windows.Add()
                wwp.Activate
                If slRng(i).Information(wdInFootnote) Then
                    With wwp
                        If .Panes.Count = 1 Then
                            '開啟註腳視窗
                            If .View.Type = wdNormalView Then _
                               .View.SplitSpecial = wdPaneFootnotes
                        Else
                            .ActivePane.Next.Activate
                        End If
    '                    .ScrollIntoView .ActivePane.Selection, True
    '                    .ActivePane.SmallScroll
                    End With
                End If
                slRng(i).Select
                i = i + 1
            End If
        Next
        wnd.Activate
    End If
    If rng.Document.path <> "" Then rng.Document.Save
End Sub

Sub insertHydzdLink()
Dim lk As New Links, db As New dBase
db.setWordControlValue (文字處理.trimStrForSearch(Selection.text, Selection))
db.setDictControlValue 3
lk.insertLinktoHydzd
Set lk = Nothing: Set db = Nothing
End Sub
Sub insertHydcdLink()
Dim lk As New Links, db As New dBase
db.setWordControlValue (文字處理.trimStrForSearch(Selection.text, Selection))
db.setDictControlValue 4
lk.insertLinktoHydcd
Set lk = Nothing: Set db = Nothing
End Sub
Sub updateURL國語辭典()
Dim lnks As New Links
lnks.updateURL國語辭典 ActiveDocument
'SystemSetup.playSound 7
MsgBox "done!", vbInformation
End Sub
Sub saveV5URL()
Dim ac As Object, lnk As String
Dim dbFullName As String
dbFullName = UserProfilePath & "Dropbox\《重編國語辭典修訂本》資料庫.mdb"
If Selection.Hyperlinks.Count > 0 Then
    lnk = Selection.Hyperlinks(1).Address
Else
    Selection.MoveRight wdCharacter, 1, wdExtend
    lnk = Selection.Hyperlinks(1).Address
End If
Set ac = GetObject(dbFullName).Application
ac.Run "saveV5URL", lnk
AppActivate "access"
End Sub
Sub updateURL國學大師()
Dim lnks As New Links
lnks.updateURL國學大師 ActiveDocument
'SystemSetup.playSound 7
End Sub
Sub 標題文字()
With Selection.font
    .Size = 20
    .Bold = True
End With
End Sub

Function 去標楷體字(r As Range)
'If InStr(r, ": 標楷體""") Then
    r = Replace(r, "<span style=""FONT-FAMILY: 標楷體"">", "<span>")
    r = Replace(r, "; FONT-FAMILY: 標楷體"">", """>")
    r = Replace(r, "FONT-FAMILY: 標楷體;", "")
    r = Replace(r, "; mso-fareast-font-family: 標楷體", "")
    r = Replace(r, "mso-fareast-font-family: 標楷體", "")
    r = Replace(r, "font-family: 標楷體; ", "") '2009/4/17補此行
    
    r = Replace(r, "<span style=""FONT-FAMILY: 新細明體"">", "<span>")
    r = Replace(r, "; FONT-FAMILY: 新細明體"">", """>")
    r = Replace(r, "FONT-FAMILY: 新細明體;", "")
    r = Replace(r, "; mso-fareast-font-family: 新細明體", "")
    r = Replace(r, "mso-fareast-font-family: 新細明體", "")
    
    r = Replace(r, "<span style=""FONT-FAMILY: 新細明體; mso-ascii-font-family: 'Times New Roman'; mso-hansi-font-family: 'Times New Roman'"">", "<span>")
    r = Replace(r, "; FONT-FAMILY: 新細明體", "")
    r = Replace(r, "; mso-bidi-font-size: 12.0pt; mso-ascii-font-family: 'Times New Roman'; mso-hansi-font-family: 'Times New Roman'", "")
    r = Replace(r, "; mso-bidi-font-size: 12.0pt", "")
    r = Replace(r, "<span lang=EN-US>", "<span>")
    r = Replace(r, " lang=EN-US", "")
    r = Replace(r, "<span style=""FONT-SIZE: 8pt; COLOR: navy""><font face=""Times New Roman"">", "<span style=""COLOR: navy""><font face=""Times New Roman"" size=2>")
    r = Replace(r, "; mso-text-animation: ants-red", "")
    r = Replace(r, "</span><span style=""COLOR: navy"">", "")
    r = Replace(r, " size=2>", ">")
    r = Replace(r, "<span style=""FONT-SIZE: 8pt; COLOR: navy;  mso-bidi-font-size: 12.0pt; mso-ascii-font-family: 'Times New Roman'; mso-hansi-font-family: 'Times New Roman'"">", "<span style=""FONT-SIZE: 8pt; COLOR: navy"">")
    r = Replace(r, "<span style="" mso-ascii-font-family: 'Times New Roman'; mso-hansi-font-family: 'Times New Roman'"">", "<span>")
    r = Replace(r, "<span style=""mso-tab-count: 1"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>", "")
    r = Replace(r, "<span style=""mso-tab-count: 1""><font face=""Times New Roman"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font></span>", "")
    r = Replace(r, "<span style=""mso-tab-count: 1"">&nbsp;&nbsp;&nbsp;&nbsp; </span>", "")
    r = Replace(r, ";  mso-ascii-font-family: 'Times New Roman'; mso-hansi-font-family: 'Times New Roman'", "")
    r = Replace(r, "<span style=""mso-tab-count: 1"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>", "")
    r = Replace(r, "<span style=""mso-tab-count: 3"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>", "")
    r = Replace(r, "<font face=""Times New Roman"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font>", "")
    r = Replace(r, "<span style=""mso-tab-count: 1"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>", "")
    r = Replace(r, "<span style=""mso-tab-count: 1""><font face=""Times New Roman"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font></span>", "")
    r = Replace(r, "</span><span style=""FONT-SIZE: 8pt; COLOR: navy""><span style=""mso-spacerun: yes""><font face=""Times New Roman"">&nbsp; </font></span></span><span style=""FONT-SIZE: 8pt; COLOR: navy"">", "&nbsp; ")
    
    r = Replace(r, "<font face=""Times New Roman""> </font></span><span style=""FONT-SIZE: 8pt; COLOR: navy"">", " ")
    
    r = Replace(r, ".</font></span><span style=""FONT-SIZE: 8pt; COLOR: navy"">", ".</font>", , , vbBinaryCompare)
    r = Replace(r, "\</font></span><span style=""FONT-SIZE: 8pt; COLOR: navy"">", "\</font>", , , vbBinaryCompare)
    r = Replace(r, "_</font></span><span style=""FONT-SIZE: 8pt; COLOR: navy"">", "_</font>", , , vbBinaryCompare)
    r = Replace(r, "<font face=""Times New Roman"">-</font></span><span style=""FONT-SIZE: 8pt; COLOR: navy"">", "<font face=""Times New Roman"">-</font>", , , vbBinaryCompare)
    
    r = Replace(r, "&nbsp;&nbsp;&nbsp;&nbsp;", "&nbsp;&nbsp;")
    r = Replace(r, "<font face=""Times New Roman"">&nbsp;&nbsp;&nbsp; </font>", "&nbsp;&nbsp;&nbsp; ")
    r = Replace(r, "<font face=""Times New Roman"">&nbsp;&nbsp; </font>", "&nbsp;&nbsp; ")
    r = Replace(r, "<font face=""Times New Roman"">&nbsp; </font>", "&nbsp; ")
    r = Replace(r, "<span style=""mso-tab-count: 1"">&nbsp;&nbsp;&nbsp; </span>", "&nbsp;&nbsp;&nbsp; ")
    r = Replace(r, "<span style=""mso-tab-count: 1"">&nbsp;&nbsp; </span>", "&nbsp;&nbsp; ")
    r = Replace(r, "<span style=""mso-tab-count: 1"">&nbsp; </span>", "&nbsp; ")
    r = Replace(r, "<span style=""mso-tab-count: 2"">&nbsp;&nbsp;&nbsp; </span>", "&nbsp;&nbsp;&nbsp; ")
    r = Replace(r, "<span style=""mso-tab-count: 2"">&nbsp;&nbsp; </span>", "&nbsp;&nbsp; ")
    r = Replace(r, "<span style=""mso-tab-count: 2"">&nbsp; </span>", "&nbsp; ")

    r = Replace(r, "&nbsp;&nbsp;&nbsp;", "　　　")
    r = Replace(r, "&nbsp;&nbsp;", "　　")
    r = Replace(r, "&nbsp;", " ")
    r = Replace(r, "<font face=""Times New Roman""> </font>", " ")
    r = Replace(r, "<p class=MsoNormal style=""MARGIN: 0cm 0cm 0pt; TEXT-INDENT: 24pt"">", "<p>")
    
    r = Replace(r, "<span style=""FONT-SIZE: 8pt; COLOR: navy""><span style=""mso-tab-count: 1""> </span></span>", " ")
    r = Replace(r, "<span style=""mso-spacerun: yes"">  </span>", "")
    r = Replace(r, ";  mso-bidi-font-size: 12.0pt; mso-ascii-font-family: 'Times New Roman'; mso-hansi-font-family: 'Times New Roman'"">", """>")
    
'    r = Replace(r, "真按：</span></strong><span style=""FONT-SIZE: 8pt; COLOR: navy"">", "真按：</strong>")
    r = Replace(r, "<strong><span style=""FONT-SIZE: 8pt; COLOR: navy"">真按：</span></strong><span style=""FONT-SIZE: 8pt; COLOR: navy"">" _
            , "<span style=""FONT-SIZE: 8pt; COLOR: navy""><strong>真按：</strong>")
    
    
    
    r = Replace(r, "; mso-bidi-font-family: 'Times New Roman'; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA", "")
    r = Replace(r, "; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA", "")
    r = Replace(r, "; FONT-FAMILY: 'Times New Roman'", "")
    
    
    
    去標楷體字 = r
'End If
End Function
Sub 去標楷體字s()
Dim r As Range
Set r = ActiveDocument.Range
'If InStr(r, ": 標楷體""") Then
'    r = Replace(r, "<span style=""FONT-FAMILY: 標楷體"">", "<span>")
'    r = Replace(r, "; FONT-FAMILY: 標楷體"">", """>")
'    r = Replace(r, "FONT-FAMILY: 標楷體;", "")
'    r = Replace(r, "; mso-fareast-font-family: 標楷體", "")
'    r = Replace(r, "mso-fareast-font-family: 標楷體", "")
'
'    r = Replace(r, "<span style=""FONT-FAMILY: 新細明體"">", "<span>")
'    r = Replace(r, "; FONT-FAMILY: 新細明體"">", """>")
'    r = Replace(r, "FONT-FAMILY: 新細明體;", "")
'    r = Replace(r, "; mso-fareast-font-family: 新細明體", "")
'    r = Replace(r, "mso-fareast-font-family: 新細明體", "")
    
    r = 去標楷體字(r)
    'r = Replace(r, "", "")
    
    With ActiveDocument
        .Range = r
        With .Windows(1)
            .Selection.EndKey wdStory, wdMove
            Do Until .Selection.Previous <> VBA.Chr(13)
                .Selection.TypeBackspace
            Loop
        End With
        .Range.WholeStory
        .Range.Cut
    End With
'End If
End Sub

Sub 檢查亂碼問號() '2008/8/19防止檔名裡有亂碼?以礙燒錄備分也.
Dim d  As Document
Set d = Documents.Add
'd.Range.Paste
d.Range.PasteAndFormat (wdFormatPlainText)
If InStr(d.Range, "?") Then
    MsgBox "有亂碼!!", vbCritical
    With d.Windows(1).Selection.Find
        .ClearFormatting
        .Execute "?"
    End With
Else
    d.Close wdDoNotSaveChanges
End If
End Sub

Sub 由程式碼中取出圖片暨連結並排序()
Dim d, x As String, a, s As Long, e As Long, p As String
Set d = ActiveDocument
x = d.Range
s = InStr(x, "<p><a href=""http://tw.myblog.yahoo.com/jw%21ob4NscCdAxS_yWJbxTvlgfR./photo?pid=")
If s Then
    e = InStr(x, ".jpg"" /></a></p>")
    With d
        Do Until s = 0
            p = p & VBA.Mid(x, s, (e - (s - 1)) + 16) '16=len(".jpg"" /></a></p>")
            s = InStr(s + 1, x, "<p><a href=""http://tw.myblog.yahoo.com/jw%21ob4NscCdAxS_yWJbxTvlgfR./photo?pid=")
            e = InStr(e + 1, x, ".jpg"" /></a></p>")
        Loop
    End With
End If
Debug.Print p
End Sub
Sub 由程式碼中取出圖片並排序()
Dim d, x As String, a, s As Long, e As Long, p As String, pe As String, ps As String, pL As Byte
setOX
'Set d = ActiveDocument
d = OX.ClipGet
'x = d.Range
x = d
If InStr(x, "<a name") Then MsgBox "本頁有書籤,請檢查!!", vbCritical: Exit Sub
s = InStr(x, """><img src=""http://tw.blog.yahoo.com/photo/photo.php?id=ob4NscCdAxS_yWJbxTvlgfR.&amp;photo=ap_")
If s Then
    ps = """><img src=""http://tw.blog.yahoo.com/photo/photo.php?id=ob4NscCdAxS_yWJbxTvlgfR.&amp;photo=ap_"
ElseIf InStr(x, """><img alt="""""""" src=""http://tw.blog.yahoo.com/photo/photo.php?id=ob4NscCdAxS_yWJbxTvlgfR.&amp;photo=ap_") Then
    ps = """><img alt="""""""" src=""http://tw.blog.yahoo.com/photo/photo.php?id=ob4NscCdAxS_yWJbxTvlgfR.&amp;photo=ap_"
    s = InStr(x, ps)
End If
If s Then
    e = InStr(s, x, ".jpg"" /></a></p>")
    If e Then
        pe = ".jpg"" /></a></p>"
    ElseIf InStr(s, x, ".jpg"" /></a></font>") Then
        pe = ".jpg"" /></a></font>"
    ElseIf InStr(s, x, ".jpg"" border=""0"" /></a><") Then
        pe = ".jpg"" border=""0"" /></a><"
        pL = 11 'pe比對長度有變,故要設此參數 2009/6/15 _
        是以pe=".jpg"" /></a></p>"作基準參照的,故要看jpg與>間又多出多少
    End If
    e = InStr(s, x, pe)
    With d
        Do Until s = 0
            p = p & VBA.Mid(x, s + 2, (e - (s - 1)) + 5 + pL) & "<br>" '16=len(".jpg"" /></a></p>")
            s = InStr(s + 1, x, ps)
            e = InStr(s + 1, x, pe) 'e = InStr(s + 1, x, ".jpg"" /></a></p>")
        Loop
    End With
End If
'Debug.Print p
OX.ClipPut p
On Error Resume Next
AppActivate "avant browser"
End Sub


Sub 圖片公開隱藏屬性()
'AppActivate "opera"
AppActivate 2312
DoEvents
'SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{tab}", True
'SendKeys " ", True
'SendKeys "{tab}", True: SendKeys " ", True
'SendKeys "{tab}", True: SendKeys " ", True
'SendKeys "{tab}", True: SendKeys " ", True
'SendKeys "{tab}", True: SendKeys " ", True
'SendKeys "{tab}", True: SendKeys " ", True
'SendKeys "{tab}", True: SendKeys " ", True
'SendKeys "{tab}", True: SendKeys " ", True
'SendKeys "{tab}", True: SendKeys " ", True
SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{left}", True
SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{left}", True
SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{left}", True
SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{left}", True
SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{left}", True
SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{left}", True
SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{left}", True
SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{left}", True
SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{left}", True
SendKeys "{tab}", True: SendKeys "{tab}", True:: SendKeys " ", True
End Sub

Sub 查詢奇摩我的部落格blog() 'Alt+Q
If Selection.Type = wdSelectionIP Then Exit Sub
If ActiveDocument.path <> "" And ActiveDocument.Saved = False Then ActiveDocument.Save
If myaccess Is Nothing Then
    Set myaccess = GetObject("C:\千慮一得齋\書籍資料\圖書管理(C槽版).mdb")
End If
myaccess.Run "查詢奇摩我的部落格blog_word參照", Selection
Selection.Copy
myaccess.UserControl = True
Set myaccess = Nothing
End Sub

Sub 異體字字典插入超連結()
Dim st As String, h As String, d As Document
If ActiveDocument.path <> "" Then Exit Sub
Set d = ActiveDocument
If InStr(d.Range, "http") = 0 Then Exit Sub
With d.Application.Selection
    .HomeKey wdStory, wdMove
    .Find.ClearFormatting
    Do
1   If .Find.Execute("http", , , , , , True, wdFindStop) = False Then Exit Sub
    'If d.Range(.Start - 2, .Start) = "異." Then '插入超連結－
    If d.Range(.start - 2, .start) = "異." Then
        Do Until st = ".htm"
            .MoveRight wdCharacter, 1, wdExtend
            st = VBA.Right(.text, 4)
        Loop
        h = .text
        .Delete
        h = VBA.StrConv(h, vbNarrow) '全形轉半形
        .Hyperlinks.Add d.Range(.start - 2, .start - 1), h '.Range, h
        st = ""
    ElseIf d.Range(.start - 1, .start) = VBA.Chr(9) Or d.Range(.start - 1, .start) = "』" Then '以tab鍵定位字元為判斷,蓋在資料庫vba.Chr(13)皆會被轉成此字元故也.
        Do Until st = VBA.Chr(9) Or st = " " Or .Next.font.Size > 8
            .MoveRight wdCharacter, 1, wdExtend
            st = VBA.Right(.text, 1)
        Loop
        .MoveLeft wdCharacter, 1, wdExtend
        h = .text
        .Delete
        h = VBA.StrConv(h, vbNarrow) '全形轉半形
        .Hyperlinks.Add d.Range(.start - 1, .start), h      '.Range, h
        st = ""
        
    Else
        GoTo 1
    End If
    Loop
End With

End Sub

Sub 檢查尚未發布之eMule清單() '2011/6/19
Dim Dnow As Document, Dold As Document, p As Paragraph, x, l
If Documents(1).path <> "" Or Documents(2).path <> "" Then Exit Sub
Set Dnow = Documents(1) '最後一個文件為目前複製自emule的清單,前一個文件則為blog清單帖複製來的
Set Dold = Documents(2)
With Dold
    For Each p In .Paragraphs
        x = VBA.Left(p.Range, Len(p.Range) - 1)
        l = InStr(Dnow.Range, x)
        If l = 0 Then
            p.Range.Select
            Exit For
        Else
            Dnow.Characters(l).Paragraphs(1).Range.Delete
        End If
    Next p
End With
Dold.Activate
End Sub


Sub 抽出書首各頁連結以便搜尋引擎索引() '2011/7/14'http://www.webconfs.com/search-engine-spider-simulator.php
Dim i As Long, j As Long, x As String, l As Long
With ActiveDocument
    If .path <> "" Then Exit Sub
    .Range.Paste
    x = .Range
    l = Len(x)
    i = 1
    Do Until i > l
    Select Case VBA.Mid(x, i, 1) & VBA.Mid(x, i + 1, 1)
        Case """>"
            j = i + 2
            Do Until VBA.Mid(x, j, 1) & VBA.Mid(x, j + 1, 1) & VBA.Mid(x, j + 2, 1) = "</a"
                x = VBA.Left(x, j - 1) & VBA.Mid(x, j + 1) ' Replace(x, VBA.Mid(x, j, 1), "", j, 1)
                'VBA.Mid(x, j, 1) = ""
                'j = j + 1
                l = l - 1
            Loop
            i = j
        Case "a>"
            j = i + 2
            Do Until VBA.Mid(x, j, 1) & VBA.Mid(x, j + 1, 1) & VBA.Mid(x, j + 2, 1) = "<a "
                x = VBA.Left(x, j - 1) & VBA.Mid(x, j + 1)
                l = l - 1
                If VBA.Mid(x, j, 1) & VBA.Mid(x, j + 1, 1) & VBA.Mid(x, j + 2, 1) = "" Then Exit Do
            Loop
            i = j
        
    End Select
    i = i + 1
    Loop
    x = Replace(x, "> <", "><")
    x = Replace(x, " &nbsp;", "")
    x = Replace(x, "<br>", "")
    x = Replace(x, "<div>", "")
    x = Replace(x, "</div>", "")
    x = Replace(x, "<p>", "")
    x = Replace(x, "</p>", "")
    .Range = x
    .Range.Cut
End With

End Sub

Sub 插入頁圖() '20161105
Dim p  As Paragraph, s As Long, lnk As String
Const x As String = "\\VBOXSVR\d_drive\千慮一得齋\資料庫\掃描資料庫\書藏\2487_語言與人生\"
s = Selection.End
For Each p In ActiveDocument.Paragraphs
    If p.Range.End > s Then
        If IsNumeric(p.Range) And p.Range.font.Size = 16 Then
            If Dir(x & Format(p.Range, "_000000") & ".tif") = "" Then
                If Dir(x & Format(p.Range, "_000000") & ".jpg") <> "" Then
                    lnk = x & Format(p.Range, "_000000") & ".jpg"
                Else
                    GoTo nt
                End If
            ElseIf Dir(x & Format(p.Range, "_000000") & ".jpg") = "" Then
                If Dir(x & Format(p.Range, "_000000") & ".tif") <> "" Then
                    lnk = x & Format(p.Range, "_000000") & ".tif"
                Else
                    GoTo nt
                End If
            Else
                GoTo nt
            End If
            p.Range.Hyperlinks.Add p.Range, lnk
        End If
    End If
nt:
Next p
MsgBox "done!", vbInformation
End Sub
