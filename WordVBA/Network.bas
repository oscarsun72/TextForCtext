Attribute VB_Name = "Network"
Option Explicit
Dim DefaultBrowserNameAppActivate As String

Sub 查詢國語辭典() '指定鍵:Ctrl+F12'2010/10/18修訂
    ''    If ActiveDocument.Path <> "" Then ActiveDocument.Save '怕word當掉忘了儲存
    ''    If GetUserAddress = True Then
    '''        MsgBox "成功的跟隨超連結。"
    ''    Else
    ''        MsgBox "無法跟隨超連結。"
    ''    End If
    '    Selection.Copy
    '    Shell "W:\!! for hpr\VB\查詢國語辭典\查詢國語辭典\bin\Debug\查詢國語辭典.EXE"
    Const st As String = "C:\Program Files\孫守真\查詢國語辭典等\"
    Const f As String = "查詢國語辭典.EXE"
    Dim funame As String
    If Selection.Type = wdSelectionNormal Then
        文字處理.ResetSelectionAvoidSymbols
        Selection.Copy
        If Dir(st & f) <> "" Then
            funame = st & f
        ElseIf Dir("C:\Program Files (x86)\孫守真\查詢國語辭典等\" & f) <> "" Then
            funame = "C:\Program Files (x86)\孫守真\查詢國語辭典等\" & f
        ElseIf Dir("W:\!! for hpr\VB\查詢國語辭典\查詢國語辭典\bin\Debug\" & f) <> "" Then
            funame = "W:\!! for hpr\VB\查詢國語辭典\查詢國語辭典\bin\Debug\" & f
        ElseIf Dir("C:\查詢國語辭典\查詢國語辭典\bin\Debug\" & f) <> "" Then
            funame = "C:\查詢國語辭典\查詢國語辭典\bin\Debug\" & f
        ElseIf Dir(UserProfilePath & "Dropbox\VS\VB\查詢國語辭典\查詢國語辭典\bin\Debug\" & f) <> "" Then
            funame = UserProfilePath & "Dropbox\VS\VB\查詢國語辭典\查詢國語辭典\bin\Debug\" & f
        ElseIf Dir("A:\", vbVolume) <> "" Then
            If Dir("A:\Users\oscar\Dropbox\VS\VB\查詢國語辭典\查詢國語辭典\bin\Debug\" & f) <> "" Then _
            funame = "A:\Users\oscar\Dropbox\VS\VB\查詢國語辭典\查詢國語辭典\bin\Debug\" & f
        ElseIf Dir(UserProfilePath & "Dropbox\VS\VB\查詢國語辭典\查詢國語辭典\bin\Debug\" & f) <> "" Then
            funame = UserProfilePath & "Dropbox\VS\VB\查詢國語辭典\查詢國語辭典\bin\Debug\" & f
        Else
            Exit Sub
        End If
        Shell funame
    End If
    查國語辭典
End Sub

Sub A速檢網路字辭典() '指定鍵:Alt+F12'2010/10/18修訂
Const f As String = "速檢網路字辭典.EXE"
Const st As String = "C:\Program Files\孫守真\速檢網路字辭典\"
Dim funame As String
If Selection.Type = wdSelectionNormal Then
    文字處理.ResetSelectionAvoidSymbols
    Selection.Copy
    If Dir(st & f) <> "" Then
        funame = st & f
    ElseIf Dir("C:\Users\oscar\Dropbox\VS\VB\速檢網路字辭典\速檢網路字辭典\bin\Debug\" & f) <> "" Then
        funame = "C:\Users\oscar\Dropbox\VS\VB\速檢網路字辭典\速檢網路字辭典\bin\Debug\" & f
    ElseIf Dir("C:\Program Files (x86)\孫守真\速檢網路字辭典\" & f) <> "" Then
        funame = "C:\Program Files (x86)\孫守真\速檢網路字辭典\" & f
    ElseIf Dir("W:\!! for hpr\VB\速檢網路字辭典\速檢網路字辭典\bin\Debug\" & f) <> "" Then
        funame = "W:\!! for hpr\VB\速檢網路字辭典\速檢網路字辭典\bin\Debug\" & f
    ElseIf Dir("C:\速檢網路字辭典\速檢網路字辭典\bin\Debug\" & f) <> "" Then
        funame = "C:\速檢網路字辭典\速檢網路字辭典\bin\Debug\" & f
    ElseIf Dir(UserProfilePath & "Dropbox\VS\VB\速檢網路字辭典\速檢網路字辭典\bin\Debug\" & f) <> "" Then
        funame = "A:\Users\oscar\Dropbox\VS\VB\速檢網路字辭典\速檢網路字辭典\bin\Debug\" & f
    ElseIf Dir(UserProfilePath & "Dropbox\VS\VB\速檢網路字辭典\速檢網路字辭典\bin\Debug\" & f) <> "" Then
        funame = "A:\Users\oscar\Dropbox\VS\VB\速檢網路字辭典\速檢網路字辭典\bin\Debug\" & f
    ElseIf Dir(UserProfilePath & "Dropbox\VS\VB\速檢網路字辭典\速檢網路字辭典\bin\Debug\" & f) <> "" Then
        funame = UserProfilePath & "Dropbox\VS\VB\速檢網路字辭典\速檢網路字辭典\bin\Debug\" & f
    Else
        Exit Sub
    End If
    Shell funame
End If

End Sub

Sub 查國語辭典()
    文字處理.ResetSelectionAvoidSymbols
    SeleniumOP.dictRevisedSearch VBA.Replace(Selection, VBA.Chr(13), "")
End Sub

'Sub 擷取國語辭典詞條網址()
'SeleniumOP.grabDictRevisedUrl VBA.Replace(Selection, vba.Chr(13), "")
'End Sub
Sub 查Google()
    Rem Alt + g
    文字處理.ResetSelectionAvoidSymbols
    SeleniumOP.GoogleSearch Selection.text
End Sub
Sub 查百度()
    Rem Alt b
    文字處理.ResetSelectionAvoidSymbols
    SeleniumOP.BaiduSearch Selection.text
End Sub
Sub 查字統網()
    Rem Alt + z
    文字處理.ResetSelectionAvoidSymbols
    If Selection.Characters.Count > 1 Then
        MsgBox "限查1字", vbExclamation ', vbError
        Exit Sub
    End If
    SeleniumOP.LookupZitools Selection.text
End Sub
Sub 查異體字字典()
    Rem Alt + F12
    文字處理.ResetSelectionAvoidSymbols
    If Selection.Characters.Count > 1 Then
        MsgBox "限查1字", vbExclamation ', vbError
        Exit Sub
    End If
    SeleniumOP.LookupDictionary_of_ChineseCharacterVariants Selection.text
End Sub
Sub 查康熙字典網上版()
    Rem Ctrl + Alt + x
    文字處理.ResetSelectionAvoidSymbols
    If Selection.Characters.Count > 1 Then
        MsgBox "限查1字", vbExclamation ', vbError
        Exit Sub
    End If
    SeleniumOP.LookupKangxizidian Selection.text
End Sub
Sub 查國語辭典_到網頁去看()
    Rem Ctrl + Alt + F12
    文字處理.ResetSelectionAvoidSymbols
    SeleniumOP.LookupDictRevised Selection.text
End Sub
Sub 查漢語大詞典()
    Rem Alt + c
    文字處理.ResetSelectionAvoidSymbols
    If Selection.Characters.Count < 2 Then
        MsgBox "要2字以上才能檢索！", vbExclamation ', vbError
        Exit Sub
    End If
    SeleniumOP.LookupHYDCD Selection.text
End Sub
Sub 查國學大師()
    Rem Ctrl + d + s （ds：大師）
    文字處理.ResetSelectionAvoidSymbols
    SeleniumOP.LookupGXDS Selection.text
End Sub
Sub 查白雲深處人家說文解字圖像查閱_藤花榭本優先()
    Rem  Alt + s （說文的說） Alt + j （解字的解）
    文字處理.ResetSelectionAvoidSymbols
    If Selection.Characters.Count > 1 Then
        MsgBox "限查1字", vbExclamation ', vbError
        Exit Sub
    End If
    Dim ar 'As Variant
    ar = SeleniumOP.LookupHomeinmistsShuowenImageAccess_VineyardHall(Selection.text)
    If ar(0) = vbNullString Then
        MsgBox "找不到，或網頁當了或改版了！", vbExclamation
'    Else
'        word.Application.Activate
'        If ar(1) = "" Then MsgBox "找出結果不止1條，請手動自行操作！", vbInformation
    End If
End Sub
Sub 查白雲深處人家說文解字圖文檢索WFG版_解說檢索()
    Rem  Alt + shift + s （說文的說） Alt + Shift + j （解字的解）
    文字處理.ResetSelectionAvoidSymbols
    SeleniumOP.LookupHomeinmistsShuowenImageTextSearchWFG_Interpretation Selection.text
End Sub
Sub 查漢語多功能字庫並取回其說文解釋欄位之值插入至插入點位置()
    Rem  Alt + n （n= 能 neng）
    文字處理.ResetSelectionAvoidSymbols
    If Selection.Characters.Count > 1 Then
        MsgBox "限查1字", vbExclamation ', vbError
        Exit Sub
    End If
    Dim ar 'As Variant
    Dim windowState As word.WdWindowState     '記下原來的視窗模式
    windowState = word.Application.windowState '記下原來的視窗模式
    ar = SeleniumOP.LookupMultiFunctionChineseCharacterDatabase(Selection.text)
    If ar(0) = vbNullString Then
        word.Application.Activate
        MsgBox "找不到，或網頁當了或改版了！", vbExclamation
        With Selection.Application
            .Activate
            With .ActiveWindow
                If .windowState = wdWindowStateMinimize Then
                    .windowState = windowState
                    .Activate
                End If
            End With
        End With
    Else 'ar(0)不為空時
        Dim ur As UndoRecord, fontsize As Single
        SystemSetup.stopUndo ur, "查漢語多功能字庫並取回其說文解釋欄位之值插入至插入點位置"
        With Selection
            fontsize = VBA.IIf(.font.Size = 9999999, 12, .font.Size) - 4
            If fontsize < 0 Then fontsize = 10
            If .Type = wdSelectionIP Then
                .Move
            Else
                .Collapse wdCollapseEnd
            End If
            '插入取回的《說文》內容
            .TypeText "，《說文》云：「"
            .InsertAfter ar(0) & "」" & VBA.Chr(13) 'ar(0)=《說文》內容
            .Collapse wdCollapseEnd
            If Selection.End = Selection.Document.Range.End - 1 Then
                Selection.Document.Range.InsertParagraphAfter
            End If
            .font.Size = fontsize
            .InsertAfter ar(1) '植入網址
            SystemSetup.contiUndo ur
            .Collapse wdCollapseStart
            With .Application
                .Activate
                With .ActiveWindow
                    If .windowState = wdWindowStateMinimize Then
                        .windowState = windowState
                        .Activate
                    End If
                End With
            End With
        End With
    End If
End Sub
Sub 查說文解字並取回其解釋欄位及網址值插入至插入點位置()
    Rem  Alt + o （o= 說文解字 ShuoWen.ORG 的 O）
    文字處理.ResetSelectionAvoidSymbols
    If Selection.Characters.Count > 1 Then
        MsgBox "限查1字", vbExclamation ', vbError
        Exit Sub
    End If
    Dim ar 'As Variant
    Dim windowState As word.WdWindowState      '記下原來的視窗模式
    windowState = word.Application.windowState '記下原來的視窗模式
    ar = SeleniumOP.LookupShuowenOrg(Selection.text)
    If ar(0) = vbNullString Then
        word.Application.Activate
        MsgBox "找不到，或網頁當了或改版了！", vbExclamation
        With Selection.Application
            .Activate
            With .ActiveWindow
                If .windowState = wdWindowStateMinimize Then
                    .windowState = windowState
                    .Activate
                End If
            End With
        End With
    Else 'ar(0)不為空時
        Dim ur As UndoRecord, fontsize As Single
        SystemSetup.stopUndo ur, "查說文解字並取回其解釋欄位及網址值插入至插入點位置"
        With Selection
            fontsize = VBA.IIf(.font.Size = 9999999, 12, .font.Size) - 4
            If fontsize < 0 Then fontsize = 10
            If .Type = wdSelectionIP Then
                .Move
            Else
                .Collapse wdCollapseEnd
            End If
            .TypeText "，《說文》云：「"
            .InsertAfter ar(0) & "」" & VBA.Chr(13) 'ar(0)=《說文》內容
            .Collapse wdCollapseEnd
            If Selection.End = Selection.Document.Range.End - 1 Then
                Selection.Document.Range.InsertParagraphAfter
            End If
            .font.Size = fontsize
            .InsertAfter ar(1) '插入網址
            SystemSetup.contiUndo ur
            .Collapse wdCollapseStart
            With .Application
                .Activate
                With .ActiveWindow
                    If .windowState = wdWindowStateMinimize Then
                        .windowState = windowState
                        .Activate
                    End If
                End With
            End With
        End With
    End If
End Sub
Sub 查說文解字並取回其解釋欄位段注及網址值插入至插入點位置()
    Rem  Ctrl+ Shift + Alt + o （o= 說文解字 ShuoWen.ORG 的 O）
    文字處理.ResetSelectionAvoidSymbols
    If Selection.Characters.Count > 1 Then
        MsgBox "限查1字", vbExclamation ', vbError
        Exit Sub
    End If
    Dim ar 'As Variant
    Dim windowState As word.WdWindowState      '記下原來的視窗模式
    windowState = word.Application.windowState '記下原來的視窗模式
    
    ar = SeleniumOP.LookupShuowenOrg(Selection.text, True)
    
    If ar(0) = vbNullString Then
        word.Application.Activate
        MsgBox "找不到，或網頁當了或改版了！", vbExclamation
        With Selection.Application
            .Activate
            With .ActiveWindow
                If .windowState = wdWindowStateMinimize Then
                    .windowState = windowState
                    .Activate
                End If
            End With
        End With
    Else 'ar(0)不為空時
        Dim ur As UndoRecord, fontsize As Single, st As Long
        SystemSetup.stopUndo ur, "查說文解字並取回其解釋欄位及網址值插入至插入點位置"
        With Selection
            fontsize = VBA.IIf(.font.Size = 9999999, 12, .font.Size) - 4
            If fontsize < 0 Then fontsize = 10
            If .Type = wdSelectionIP Then
                .Move
            Else
                .Collapse wdCollapseEnd
            End If
            st = .start
            .TypeText "，《說文》云："
            .InsertAfter ar(0) & VBA.Chr(13) 'ar(0)=《說文》內容
            .Collapse wdCollapseEnd
            If Selection.End = Selection.Document.Range.End - 1 Then
                Selection.Document.Range.InsertParagraphAfter
            End If
            If ar(2) <> vbNullString Then
                '插入段注內容
                .InsertAfter "段注本：" & VBA.IIf(VBA.Asc(VBA.Left(ar(2), 1)) = 13, vbNullString, VBA.Chr(13)) & ar(2) & VBA.Chr(13)
                Dim p As Paragraph, s As Byte, sDuan As Byte
                s = VBA.Len("                                ") '段注本的說文
                sDuan = VBA.Len("                ") '段注本的段注文
                .Paragraphs(1).Range.font.Bold = True '粗體： "段注本："
reCheck:
                For Each p In .Paragraphs
                    If VBA.InStr(p.Range.text, "清代 段玉裁《說文解字注》") Then
                        p.Range.Delete
                        GoTo reCheck:
                    ElseIf VBA.Replace(p.Range.text, " ", "") = VBA.Chr(13) Then
                        p.Range.Delete
                        GoTo reCheck:
                    ElseIf VBA.Left(p.Range.text, s) = VBA.space(s) Then '段注本的說文
                        p.Range.text = VBA.Mid(p.Range.text, s + 1)
                    ElseIf VBA.Left(p.Range.text, sDuan) = VBA.space(sDuan) Then '段注本的段注文
                        With p.Range
                            .text = VBA.Mid(p.Range.text, sDuan + 1)
                            With .font
                                .Size = fontsize + 2
                                .ColorIndex = 11 '.Font.Color= 34816
                            End With
                        End With
                    End If
                Next p
                .Collapse wdCollapseEnd
            End If
                        
            '網址格式設定
            讀入網路資料後_於其後植入網址及設定格式 .Range, VBA.CStr(ar(1)), fontsize
'            .font.Size = fontsize
'            .InsertAfter ar(1) '插入網址
'            .Collapse wdCollapseStart
            SystemSetup.contiUndo ur
            讀入網路資料後_還原視窗狀態 .Application.ActiveWindow, windowState
'            With .Application
'                .Activate
'                With .ActiveWindow
'                    If .windowState = wdWindowStateMinimize Then
'                        VBA.Interaction.DoEvents
'                        .windowState = windowState
'                        .Activate
'                        VBA.Interaction.DoEvents
'                    End If
'                End With
'            End With
            .SetRange st, st
        End With
    End If
End Sub
Sub 查異體字字典並取回其說文釋形欄位及網址值插入至插入點位置()
    Rem  Alt + v （v= 異體字 variants 的 v）
    文字處理.ResetSelectionAvoidSymbols
    If Selection.Characters.Count > 1 Then
        MsgBox "限查1字", vbExclamation ', vbError
        Exit Sub
    End If
    'ar(1) as String
    Dim ar As Variant, x As String, windowState As word.WdWindowState     '記下原來的視窗模式

    x = Selection.text
    windowState = word.Application.windowState '記下原來的視窗模式

    ar = SeleniumOP.LookupDictionary_of_ChineseCharacterVariants_RetrieveShuoWenData(x)
    
    If ar(0) = vbNullString Then
        word.Application.Activate
        MsgBox "找不到，或網頁當了或改版了！", vbExclamation
        With Selection.Application
            .Activate
            With .ActiveWindow
                If .windowState = wdWindowStateMinimize Then
                    .windowState = windowState
                    .Activate
                End If
            End With
        End With
    Else '如果ar(0)非空字串（空值）
        Dim ur As UndoRecord, fontsize As Single, st As Long ', ed As Long
        SystemSetup.stopUndo ur, "查異體字字典並取回其說文釋形欄位及網址值插入至插入點位置"
        With Selection
            st = .start
            fontsize = VBA.IIf(.font.Size = 9999999, 12, .font.Size) - 4
            If fontsize < 0 Then fontsize = 10
            If .Type = wdSelectionIP Then
                .Move
            Else
                .Collapse wdCollapseEnd
            End If
            Dim s As Byte
            s = VBA.InStr(ar(0), "《說文》不錄。")
            If s = 0 Then
                If ar(0) = "說文釋形沒有資料！" Then
                    .TypeText VBA.Chr(13)
                Else
                    .TypeText "，《說文》：" & VBA.Chr(13)
                End If
            Else
                 .TypeText "，" & VBA.Mid(ar(0), s) & VBA.Chr(13)
            End If
            Dim shuowen As String
            shuowen = VBA.Replace(VBA.Replace(ar(0), "：，", "：" & x & "，"), "段注本：", VBA.Chr(13) & "段注本：")
            If VBA.Left(shuowen, 1) = "，" Then
                shuowen = x & shuowen
            End If
            If s = 0 And ar(0) <> "說文釋形沒有資料！" Then
                If VBA.InStr(shuowen, "<img ") Then
                    word.Application.ScreenUpdating = False
                    Dim rngHtml As Range
                    Set rngHtml = .Document.Range(.start, .start)
                    '.TypeText shuoWen & VBA.Chr(13)
                    '字太長時TypeText會反應不及,會無效
                    .text = shuowen & VBA.Chr(13)
                    'ed = Selection.Range.End '插入文字後，即Selection改變後， 用 With 區塊未能及時反應！20221010
                    'Set rngHtml = .Document.Range(st, ed)
                    rngHtml.End = Selection.End
                    讀入網路資料後_還原視窗狀態 .Application.ActiveWindow, windowState
                    
                    'InnerHTML_Convert_to_WordDocumentContent Selection.Range ', vbNullString
                    InnerHTML_Convert_to_WordDocumentContent rngHtml ', vbNullString
                    rngHtml.Collapse wdCollapseEnd
                    rngHtml.Select
                Else
                    .InsertAfter shuowen & VBA.Chr(13)  'ar(0)=《說文》內容
                    .Collapse wdCollapseEnd
                End If
            End If
            
            If Selection.End = Selection.Document.Range.End - 1 Then
                Selection.Document.Range.InsertParagraphAfter
            End If
            讀入網路資料後_於其後植入網址及設定格式 .Range, VBA.CStr(ar(1)), fontsize
            讀入網路資料後_還原視窗狀態 .Application.ActiveWindow, windowState
            SystemSetup.contiUndo ur
            word.Application.ScreenUpdating = True
            .SetRange st, st
        End With
    End If
End Sub
Rem 1.指定卦名再操作 20241004 Alt + Shift + y (y:易) 。2.若游標所在為《易學網》的網址，則將其內容讀入到文件（於該連結段落後插入）
Sub 查易學網易經周易原文指定卦名文本_並取回其純文字值及網址值插入至插入點位置()
    Dim linkInput As Boolean, rngLink As Range, rngHtml As Range
    If Selection.Type = wdSelectionIP Then
        '若游標所在為《易學網》的網址，則將其內容讀入到文件
        Set rngLink = Selection.Range
        If rngLink.End + 1 = rngLink.Paragraphs(1).Range.End Then GoTo previousLink
        If rngLink.Hyperlinks.Count = 1 Then
            If VBA.Left(rngLink.Hyperlinks(1).Address, VBA.Len("https://www.eee-learning.com/")) = "https://www.eee-learning.com/" Then
                linkInput = True
            End If
        Else
            Set rngLink = Selection.Range.Next
            If Not rngLink Is Nothing Then
                If rngLink.Hyperlinks.Count = 1 Then
                    If VBA.Left(rngLink.Hyperlinks(1).Address, VBA.Len("https://www.eee-learning.com/")) = "https://www.eee-learning.com/" Then
                        linkInput = True
                    End If
                Else
previousLink:
                    Set rngLink = Selection.Range.Previous
                    If Not rngLink Is Nothing Then
                        If rngLink.Hyperlinks.Count = 1 Then
                            If VBA.Left(rngLink.Hyperlinks(1).Address, VBA.Len("https://www.eee-learning.com/")) = "https://www.eee-learning.com/" Then
                                linkInput = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        文字處理.ResetSelectionAvoidSymbols
        
        If Selection.Characters.Count > 2 Then
errExit:
            word.Application.Activate
            VBA.MsgBox "卦名有誤! 請重新選取。", vbExclamation
            Exit Sub
        End If
    End If
    Dim gua As String
    gua = Selection.text
    
    Dim ur As UndoRecord, s As Long, ed As Long
    SystemSetup.stopUndo ur, "查易學網易經周易原文指定卦名文本_並取回其純文字值及網址值插入至插入點位置"
    word.Application.ScreenUpdating = False

    Dim result(1) As String, iwe As SeleniumBasic.IWebElement

    If linkInput Then
        If SeleniumOP.OpenChrome(rngLink.Hyperlinks(1).Address) Then
            Set iwe = WD.FindElementByCssSelector("#block-bartik-content > div > article > div > div.clearfix.text-formatted.field.field--name-body.field--type-text-with-summary.field--label-hidden.field__item")
            If Not iwe Is Nothing Then
                Selection.MoveUntil Chr(13)
                Selection.TypeText Chr(13)
                Selection.Style = word.wdStyleNormal '"內文"
                
                If SeleniumOP.IslinkImageIncluded內容部分包含超連結或圖片(iwe) Then '有圖片時取 "innerHTML" 屬性值
'                    Dim Links() As SeleniumBasic.IWebElement, images() As SeleniumBasic.IWebElement
'                    Links = SeleniumOP.Links
'                    images = SeleniumOP.images
                    s = Selection.start
                    Selection.TypeText iwe.GetAttribute("innerHTML")
                    ed = Selection.End
                                        
                    Set rngHtml = Selection.Document.Range(s, ed)
                    
                    
                    
                    InnerHTML_Convert_to_WordDocumentContent rngHtml, "https://www.eee-learning.com"
                    'SeleniumOP.inputElementContentAll插入網頁元件所有的內容 iwe
                    
                    
                    

'                    Stop 'just for test
                    GoTo finish 'just for test


                Else '沒有圖片時取 "textContent" 屬性值
                    result(0) = iwe.GetAttribute("textContent")
                    result(1) = rngLink.Hyperlinks(1).Address
                End If
                
                GoTo insertText:
            End If
        End If
    Else

        If Selection.Characters.Count = 2 Then
            If Selection = "習坎" Then
                If Selection.Characters(2) = "坎" Then
                    gua = "坎"
                End If
            End If
        End If
        
        On Error GoTo eH:
        If Keywords.周易卦名_卦形_卦序.Exists(gua) = False Then
            If Keywords.易學異體字典.Exists(gua) = False Then
                GoTo errExit
            Else
                gua = Keywords.易學異體字典(gua)
            End If
        End If
    End If
    '以上防呆檢查
    
    '以下基本檢查通過後
    Dim windowState As word.WdWindowState      '記下原來的視窗模式
    windowState = word.Application.windowState '記下原來的視窗模式
    
    gua = Keywords.周易卦名_卦形_卦序(gua)(1)

    Dim fontsize As Single, rngBooks As Range
    fontsize = VBA.IIf(Selection.font.Size = 9999999, 12, Selection.font.Size) * 0.6
    If fontsize < 0 Then fontsize = 10
    
    If Not SeleniumOP.grabEeeLearning_IChing_ZhouYi_originalText(gua, result, iwe) Then
        word.Application.Activate
        VBA.MsgBox "找不到，或網頁改了或掛了……", vbInformation
        Exit Sub
    End If
    
    
'        Set iwe = WD.FindElementByCssSelector("#block-bartik-content > div > article > div > div.clearfix.text-formatted.field.field--name-body.field--type-text-with-summary.field--label-hidden.field__item")
'        If Not iwe Is Nothing Then
            If SeleniumOP.IslinkImageIncluded內容部分包含超連結或圖片(iwe) Then '有圖片時取 "innerHTML" 屬性值
            'If SeleniumOP.IsImageIncluded內容部分包含圖片(iwe) Then '有圖片時取 "innerHTML" 屬性值
                If Selection.Style <> word.wdStyleNormal Then
                    Selection.MoveUntil Chr(13)
                    Selection.TypeText Chr(13)
                    Selection.Style = word.wdStyleNormal '"內文"
                End If
                s = Selection.start
                
                Selection.TypeText iwe.GetAttribute("innerHTML")
                ed = Selection.End
                
                
                
                Set rngHtml = Selection.Document.Range(s, ed)
                
                InnerHTML_Convert_to_WordDocumentContent rngHtml, "https://www.eee-learning.com"
                'SeleniumOP.inputElementContentAll插入網頁元件所有的內容 iwe
                
                
                
                Rem 歷代注本：
                Set rngBooks = Selection.Document.Range(rngHtml.start, rngHtml.End)
                rngBooks.Find.ClearFormatting
                If rngBooks.Find.Execute("歷代注本：", , , , , , True, wdFindStop) Then
                    With rngBooks '"歷代注本："所在段落範圍
                        .Style = wdStyleHeading1
                        .font.Size = 22
                        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle '單行間距
                    End With
                    ed = 讀入網路資料後_於其後植入網址及設定格式(rngHtml, result(1), fontsize)
                    讀入網路資料後_還原視窗狀態 Selection.Document.ActiveWindow, windowState
                End If

        '                    Stop 'just for test
                GoTo finish 'just for test
        
        
        '    Else '沒有圖片時取 "textContent" 屬性值
        '        result(0) = iwe.GetAttribute("textContent")
        '        result(1) = rngLink.Hyperlinks(1).Address
            End If
'        End If
    
insertText:
    word.Application.Activate
    If VBA.InStr(result(0), "歷代注本：") Then
        If VBA.vbOK = MsgBox("是否清除「歷代注本：」以後的文字？", VBA.vbQuestion + VBA.vbOKCancel) Then
            result(0) = VBA.Left(result(0), VBA.InStr(result(0), "歷代注本：") - 1)
        End If
    End If
        
    Dim p As Paragraph, book As String, iwes() As IWebElement, e, x As String
    
    With Selection
'        fontsize = VBA.IIf(.font.Size = 9999999, 12, .font.Size) - 4
'        If fontsize < 0 Then fontsize = 10
        If .Type = wdSelectionIP And .text <> Chr(13) Then
            .Delete
        End If
        s = .start
        .TypeText VBA.Replace(result(0), ChrW(160), vbNullString)
        
        ed = Selection.End 'Selection值變動後，似乎用 With區塊無法及時回應，好像要用本尊「再次呼叫」才能反映及時真相 20241009
        Set rngBooks = .Document.Range(s, ed)
        文字處理.FixFontname rngBooks
        
        '設定粗體字
        iwes = WD.FindElementsByTagName("STRONG")
        For Each e In iwes
            Set iwe = e
            x = iwe.text '.GetAttribute("textContent")
            Do While rngBooks.Find.Execute(x, , , , , , , wdFindStop)
                Set p = rngBooks.Paragraphs(1)
                If p.Range.text = x & Chr(13) Then
                    p.Range.font.Bold = True
                    Exit Do
                End If
            Loop
            rngBooks.SetRange s, ed
        Next e
        
        rngBooks.SetRange s, ed
        '*小注字
        For Each p In rngBooks.Paragraphs
            If VBA.Left(p.Range.text, 1) = "*" Then
                With p.Range.font
                    .Size = .Size - 2
                    .ColorIndex = 11 '.Font.Color= 34816
                End With
            ElseIf VBA.InStr(p.Range.text, "*") Then
                With p.Range.Find
                    .ClearFormatting
                    With .Replacement
                        .font.ColorIndex = 11 '.Font.Color= 34816
                        .font.Bold = True
                    End With
                    .Execute "*", , , , , , , , , "*", wdReplaceAll
                    .ClearFormatting
                End With
            End If
        Next p
        
        ed = 讀入網路資料後_於其後植入網址及設定格式(Selection.Range, result(1), fontsize)
'        .Document.Range(s, s).Select
'        讀入網路資料後_還原視窗狀態 .Document.ActiveWindow, windowState
        
    End With
    
    '保留歷代注本及其超連結
    Set rngBooks = Selection.Document.Range
    If VBA.InStr(result(0), "歷代注本：") Then
        With rngBooks
            With .Find
                .ClearFormatting
                If .Execute("歷代注本：", , , , , , True, wdFindStop) Then
                    With rngBooks '"歷代注本："所在段落範圍
                        .Style = wdStyleHeading1
                        .font.Size = 22
                        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle '單行間距
                    End With
                    rngBooks.End = ed '設定為全部歷代注本的範圍
                    For Each p In rngBooks.Paragraphs
                        If p.Range.text <> Chr(13) And VBA.Left(p.Range.text, 4) <> "http" And VBA.Left(p.Range.text, 5) <> "歷代注本：" Then
                            book = VBA.Left(p.Range.text, VBA.Len(p.Range.text) - 1)
                            Set iwe = SeleniumOP.WD.FindElementByLinkText(book)
                            If Not iwe Is Nothing Then
                                Set rngLink = p.Range.Document.Range(p.Range.start, p.Range.End - 1)
                                With rngLink
                                    .Hyperlinks.Add rngLink, iwe.GetAttribute("href")
                                    .Style = wdStyleHeading2 '標題 2
                                    .font.Size = 18
                                    .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle '單行間距
                                End With
                            End If
                        End If
                    Next p
                End If
            End With
        End With
    End If
    
    '閱讀古書    '　 閱讀古圖書
    Dim chkBook As String
    If VBA.InStr(result(0), "閱讀古書") Then
        chkBook = "閱讀古書,"
    ElseIf VBA.InStr(result(0), "閱讀古圖書") Then
        chkBook = "閱讀古圖書"
    Else
'        Stop 'for check
        GoTo finish
    End If
    If chkBook <> VBA.vbNullString Then
        Set rngBooks = Selection.Document.Range(s, ed)
        With rngBooks.Find
            .ClearFormatting
            If .Execute(chkBook, , , , , , True, wdFindStop) Then
'                rngBooks.Paragraphs(1).Range.text = "　" & chkBook & Chr(13)
'                Set iwe = WD.FindElementByPartialLinkText("閱讀古書")
'                If Not iwe Is Nothing Then
                    rngBooks.SetRange rngBooks.Paragraphs(1).Range.start, rngBooks.Paragraphs(1).Range.End - 1
                    rngBooks.Hyperlinks.Add rngBooks, result(1) 'iwe.GetAttribute("href")
'                End If
            End If
        End With
    End If

finish:

    word.Application.ScreenUpdating = True
    Selection.Document.Range(s, s).Select
    讀入網路資料後_還原視窗狀態 Selection.Document.ActiveWindow, windowState
    SystemSetup.contiUndo ur
    playSound 2
    Exit Sub

eH:
    Select Case Err.Number
        Case Else
            Debug.Print Err.Number & Err.Description
            Stop 'just for test
            Resume
    End Select
End Sub
Rem 20241009 將HTML轉成Word文件內文。creedit_with_Copilot大菩薩：https://sl.bing.net/jij3PK59Rka
Sub InnerHTML_Convert_to_WordDocumentContent(rngHtml As Range, Optional domainUrlPrefix As String)
    If VBA.InStr(rngHtml.text, "<") = 0 Then Exit Sub
    
     SystemSetup.playSound 1
    
    Dim htmlStr As String, rng As Range, rngClose As Range, p As Paragraph, url As String, stRngHTML As Long, pCntr As Long
    Dim s As Integer '作為 InStr() 記下結果值用
    Dim l As Integer '作為 Len() 記下結果值用
    '作為通用變數用，或陣列記住用
    Dim arr, arr1, e '作為通用一般變數用，或陣列元素記住用
    Dim obj As Object '作為通用物件變數用
    
    'dim w As Single, h As Single, textPart As String
    
    'Dim ur As UndoRecord  'just for test
    
'    GoTo finish 'just for test
    
    '取得網址前綴的網域值（不含尾斜線）
    If domainUrlPrefix = vbNullString Then
        If Not SeleniumOP.IsWDInvalid() Then ' domainUrlPrefix = "https://www.eee-learning.com"
            domainUrlPrefix = getDomainUrlPrefix(WD.url)
        End If
    End If
    'SystemSetup.stopUndo ur, "InnerHTML_DocContent"
    stRngHTML = rngHtml.start
    htmlStr = rngHtml.text '記下起始位置
    
    Rem 前置整理文本
    rngHtml.text = VBA.Replace(VBA.Replace(VBA.Replace(htmlStr, "</p>", vbNullString), "<p>", vbNullString), "&nbsp;", ChrW(160))
    htmlStr = rngHtml.text
    
    If VBA.InStr(htmlStr, "<sup>") Then
        Set rng = rngHtml.Document.Range(rngHtml.start, rngHtml.End)
        HTML2Doc.ConvertHTMLSupToWordSup rng
    End If
    If VBA.InStr(htmlStr, "<sub>") Then
        Set rng = rngHtml.Document.Range(rngHtml.start, rngHtml.End)
        HTML2Doc.ConvertHTMLSubToWordSub rng
    End If
    With rngHtml.Find
        .ClearFormatting
        '置換
        If VBA.InStr(htmlStr, "<br>") Then .Execute "<br>", , , , , , , , , "^l", wdReplaceAll
        If VBA.InStr(htmlStr, "<a style=""line-height:1.5;"" href=") Then .Execute "<a style=""line-height:1.5;"" href=", , , , , , , , , "<a href=", wdReplaceAll ' " 會置換成 “
        If VBA.InStr(htmlStr, "&lt;") Then .Execute "&lt;", , , , , , , , , "＜", wdReplaceAll
        If VBA.InStr(htmlStr, "&gt;") Then .Execute "&gt;", , , , , , , , , "＞", wdReplaceAll
        '清除
'        If VBA.InStr(htmlStr, "<div>" & ChrW(160) & "</div>") Then .Execute "<div>" & ChrW(160) & "</div>", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, ChrW(160)) Then .Execute ChrW(160), , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, " class=""colorbox cboxElement""") Then .Execute " class=""colorbox cboxElement""", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, " class=""colorbox colorbox-insert-image cboxElement""") Then .Execute " class=""colorbox colorbox-insert-image cboxElement""", , , , , , , , , vbNullString, wdReplaceAll '
        'If VBA.InStr(htmlStr, "<a class=""colorbox colorbox-insert-image cboxElement"" href=") Then .Execute "<a class=""colorbox colorbox-insert-image cboxElement"" href=", , , , , , , , , "<a href=", wdReplaceAll ' " 會置換成 “
        If VBA.InStr(htmlStr, " rel=""group-all""") Then .Execute " rel=""group-all""", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "<o:p></o:p>") Then .Execute "<o:p></o:p>", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "<span></span>") Then .Execute "<span></span>", , , , , , , , , vbNullString, wdReplaceAll
        Rem 原網頁蓋用諸如Word等編輯貼上，故多有殘碼、亂碼
        If VBA.InStr(htmlStr, "<o:p>") Then .Execute "<o:p>", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "</o:p>") Then .Execute "</o:p>", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "<span style=""color:#ffffff;"">ppp</span>") Then .Execute "<span style=""color:#ffffff;"">ppp</span>", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "<!--EndFragment-->") Then .Execute "<!--EndFragment-->", , , , , , , , , vbNullString, wdReplaceAll
    End With
    
    SystemSetup.playSound 1
    Set rng = rngHtml.Document.Range(rngHtml.start, rngHtml.End)
    '清除空標籤
    RemoveEmptyTags rngHtml
    
    Rem 表格處理 https://sl.bing.net/fQ5lVr8PLye
    Do While VBA.InStr(rngHtml.text, "<table")
        Set rng = rngHtml.Document.Range(rngHtml.start, rngHtml.End)
        With rng.Find
            .ClearFormatting
            .text = "<table "
            .Execute
            Set rngClose = rng.Document.Range(rng.End, rngHtml.End)
            With rngClose.Find
                .text = "</table>"
                .Execute
            End With
            InsertHTMLTable rngHtml.Document.Range(rng.start, rngClose.End), domainUrlPrefix
        End With
    Loop
    
    Set rng = rngHtml.Document.Range(rngHtml.start, rngHtml.End)
    Rem 無序清單的處理
    unorderedListPorc_HTML2Word rng
    
    For Each p In rngHtml.Paragraphs
        pCntr = pCntr + 1
        If pCntr Mod 20 = 0 Then SystemSetup.playSound 1
        
        Set rng = p.Range '用set 會歸零，用 setRange 不會，只是調整
        With rng
            
'            If VBA.InStr(rng.text, "六五，貞疾，恒不") Then
'                Stop 'check
'            End If

'            If VBA.InStr(rng.text, "潛龍勿用，陽在下也") Then
'                Stop 'check
'            End If
            
            With .Find
                .ClearFormatting
                .text = "<b>"
                Do While .Execute()
                    Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                    rngClose.Find.ClearFormatting
                    If Not rngClose.Find.Execute("</b>") Then Stop 'to check
                    rng.Document.Range(rng.End, rngClose.start).font.Bold = True
                    rng.text = vbNullString: rngClose.text = vbNullString
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                Loop
                .text = "<b "
                Do While .Execute()
                    rng.MoveEndUntil ">"
                    rng.End = rng.End + 1
                    Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                    rngClose.Find.ClearFormatting
                    If Not rngClose.Find.Execute("</b>") Then Stop 'to check
                    rng.Document.Range(rng.End, rngClose.start).font.Bold = True
                    rng.text = vbNullString: rngClose.text = vbNullString
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                Loop
                .text = "<span lang=""EN-US"">"
                Do While .Execute()
                    Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                    rngClose.Find.ClearFormatting
                    If Not rngClose.Find.Execute("</span>") Then Stop 'to check
                    rng.Document.Range(rng.End, rngClose.start).font.Name = "Calibri"
                    rng.text = vbNullString: rngClose.text = vbNullString
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                Loop
                .text = "<st1:chmetcnv "
                Do While .Execute()
                    rng.MoveEndUntil ">"
                    rng.End = rng.End + 1
                    Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                    rngClose.Find.ClearFormatting
                    If Not rngClose.Find.Execute("</st1:chmetcnv>") Then Stop 'to check
                    rng.text = vbNullString: rngClose.text = vbNullString
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                Loop
                .text = "<span class=""Apple-style-span"""
                Do While .Execute()
                    rng.MoveEndUntil ">"
                    rng.End = rng.End + 1
                    Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                    rngClose.Find.ClearFormatting
                    If Not rngClose.Find.Execute("</span>") Then Stop 'to check
                    rng.text = vbNullString: rngClose.text = vbNullString
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                Loop
                .text = "<blockquote>"
                Do While .Execute()
                    Set rngClose = rng.Document.Range(rng.End, p.Next.Range.End)
                    rngClose.Find.ClearFormatting
                    If Not rngClose.Find.Execute("</blockquote>") Then Stop 'to check
                    rng.text = vbNullString
                    If rngClose.Paragraphs(1).Range.text = "</blockquote>" & Chr(13) Then
                        rngClose.Paragraphs(1).Range.text = vbNullString
                    Else
                        Stop 'for check
                        rngClose.text = vbNullString
                    End If
                    rng.ParagraphFormat.CharacterUnitLeftIndent = 3
                    rng.Paragraphs(1).Range.font.Name = "標楷體"
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                Loop
                .text = "<hr>"
                If .Execute() Then
                    rng.text = vbNullString
                    'rng.InsertBreak WdBreakType.wdLineBreak
                    With p.Borders(wdBorderBottom)
                        .LineStyle = wdLineStyleSingle ' 插入實線  插入雙線：wdLineStyleDouble 插入虛線：wdLineStyleDot
                        .LineWidth = wdLineWidth050pt
                        .Color = wdColorAutomatic
                    End With
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                End If
                .text = "<hr " 'ex: <hr style="padding-left: 30px;">
                If .Execute() Then
                    rng.MoveEndUntil ">"
                    rng.End = rng.End + 1
                    '借用 url 變數
                    url = rng.text
                    rng.text = vbNullString
                    'rng.InsertBreak WdBreakType.wdLineBreak
                    With p.Borders(wdBorderBottom)
                        .LineStyle = wdLineStyleSingle ' 插入實線  插入雙線：wdLineStyleDouble 插入虛線：wdLineStyleDot
                        .LineWidth = wdLineWidth050pt
                        .Color = wdColorAutomatic
                    End With
                    url = getHTML_AttributeValue("style", url)
                    arr = VBA.Split(url, ";")
                    For Each e In arr
                        If e <> vbNullString Then
                            e = VBA.Trim(e)
                            l = VBA.Len("padding-left: ")
                            If VBA.Left(e, l) = "padding-left: " Then
                                rng.ParagraphFormat.LeftIndent = PixelsToPoints(VBA.Replace(VBA.Mid(e, l + 1), "px", vbNullString))
                            Else
                                playSound 12
                                Stop 'for check
                            End If
                        End If
                    Next e

                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                End If
                '處理圖片
                .text = "<img "
                Do While .Execute()
                    rng.MoveEndUntil ">" 'ex: <img style="float:right;margin-left:15px;margin-right:15px;" src="/image/3.jpg" width="200" height="297"
                    '借用變數
                    url = rng.text
                    rng.End = rng.End + 1 '包含 ">"
                    rng.text = vbNullString
                    'pCntr + VBA.Abs(10 - pCntr) '下載圖片需要時間
                    If Not insert_ImageHTML(url, rng, domainUrlPrefix) Is Nothing Then
                        p.Range.ParagraphFormat.BaseLineAlignment = wdBaselineAlignCenter
                    End If
'                    If rng.Paragraphs(1).Range.ShapeRange.Count > 0 Then
'                        Stop
'                    End If
                    
                    rng.SetRange rng.End, p.Range.End
                    
                Loop
                If VBA.Len(p.Range.text) > VBA.Len("<strong></strong>") Then
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                    .text = "<strong>"
                    Do While .Execute()
                        Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                        rngClose.Find.text = "</strong>"
                        If Not rngClose.Find.Execute() Then
                            rngClose.SetRange rngClose.End, rngClose.Paragraphs(1).Next.Range.End
                            If Not rngClose.Find.Execute() Then Stop 'for check
                        End If
                        rng.Document.Range(rng.End, rngClose.start).font.Bold = True
                        rng.text = vbNullString: rngClose.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                If VBA.Len(p.Range.text) > VBA.Len("<strong style=""; ;""></strong>") Then
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                    .text = "<strong style="""
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                        rngClose.Find.text = "</strong>"
                        rngClose.Find.Execute
                        rng.Document.Range(rng.End, rngClose.start).font.Bold = True
                        rng.text = vbNullString: rngClose.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                '處理字型樣式
                If VBA.Len(p.Range.text) > VBA.Len("<span style=""></span>") Then
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                    .text = "<span style"
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                        rngClose.Find.text = "</span>"
                        rngClose.Find.Execute
                        
'                        If InStr(p.Range.text, "陽湖　孫星衍　淵如纂") Then Stop 'just for test
                        
                        '借用url變數
                        url = VBA.Replace(getHTML_AttributeValue("span style", p.Range.text), "font-family:", vbNullString)
                        url = VBA.Left(url, VBA.Len(url) - 1)
                        Select Case url
                            Case "font-size: x-large", "font-size:x-large"
                                rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size * 2
                            Case "font-size: small", "font-size:small"
                                rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size * (5 / 6)
                            Case "font-size: x-small", "font-size:x-small"
                                rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size * (2 / 3)
                            Case "font-size: xx-small", "font-size:xx-small"
                                rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size * (1 / 2)
                            Case "text-decoration:underline"
                                rng.Document.Range(rng.End, rngClose.start).font.Underline = wdUnderlineSingle
                            Case Else
                            
                                If VBA.InStr(url, ";") = 0 And VBA.InStr(url, "; ") = 0 And VBA.InStr(url, "font-size:") <> 1 And VBA.InStr(url, "line-height:") = 0 And VBA.InStr(url, "font-family") = 0 And VBA.InStr(url, "Mso") = 0 And VBA.InStr(url, "mso-") = 0 And VBA.InStr(url, "標楷體") = 0 And VBA.InStr(url, "letter-spacing:0pt") = 0 And VBA.InStr(url, "新細明體") = 0 And VBA.InStr(url, "background-color: ") = 0 And VBA.InStr(url, "color: ") = 0 Then
                                    
                                    rng.Select
                                    Debug.Print url
                                    Stop 'for check
                                End If
                                
                                'FontName
                                If VBA.Left(url, 3) = "標楷體" Then url = "標楷體"
                                If Fonts.IsFontInstalled(VBA.Trim(url)) Then
                                    If rng.Document.Range(rng.End, rngClose.start).font.Name <> VBA.Trim(url) Then
                                        rng.Document.Range(rng.End, rngClose.start).font.Name = VBA.Trim(url)
                                    End If
                                End If
                                'FontSzie
                                If VBA.InStr(url, "font-size:") = 1 Then
                                    l = VBA.Len("font-size:")
                                    If VBA.Right(url, 2) = "em" Then ' em 是一個相對單位，用於設置字體大小。它相對於父元素的字體大小。例如，如果父元素的字體大小是16像素，則 1em 等於16像素，1.5em 等於24像素。20241011 https://sl.bing.net/bVzA9JEh8VM
                                        l = VBA.Len("font-size:")
                                        rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size _
                                                * VBA.CSng(VBA.Trim(VBA.Mid(url, l + 1, VBA.Len(url) - l - VBA.Len("em"))))
                                    ElseIf VBA.IsNumeric(VBA.Mid(url, l + 1)) Then
                                        rng.Document.Range(rng.End, rngClose.start).font.Size = VBA.CSng(IIf(VBA.Mid(url, l + 1) < 1, VBA.Mid(url, l + 1) * 10, VBA.Mid(url, l + 1)))
                                    Else
                                        
                                        If VBA.InStr(url, "font-size: medium") = 0 And VBA.InStr(url, "font-size:medium") = 0 Then
                                            Stop
                                        End If
                                    End If
                                End If
                                '字型段落其他格式化雜項
                                If VBA.InStr(url, "; ") Or VBA.InStr(url, ";") Then
                                    arr = VBA.Split(url, ";")
                                    For Each e In arr
                                        e = VBA.Trim(e)
                                        If VBA.Left(e, 17) = "background-color:" Then
                                            arr1 = colorCodetoRGB(VBA.LTrim(VBA.Mid(e, VBA.Len("background-color:") + 1)))
                                            rng.Document.Range(rng.End, rngClose.start).font.Shading.BackgroundPatternColor = VBA.RGB(arr1(0), arr1(1), arr1(2))
                                        ElseIf VBA.Left(e, 6) = "color:" Then
                                            arr1 = colorCodetoRGB(VBA.LTrim(VBA.Mid(e, VBA.Len("color:") + 1)))
                                            rng.Document.Range(rng.End, rngClose.start).font.Color = VBA.RGB(arr1(0), arr1(1), arr1(2))
                                        ElseIf VBA.Left(e, 12) = "line-height:" Then
                                            arr1 = VBA.LTrim(VBA.Mid(e, VBA.Len("line-height:") + 1))
                                            If Not VBA.IsNumeric(arr1) Then
                                                If VBA.InStr(arr1, "px") Then
                                                    arr1 = VBA.Replace(arr1, "px", vbNullString)
                                                Else
                                                    playSound 12 'for check
                                                    Stop
                                                End If
                                            End If
                                            If arr1 < 10 Then
                                                rng.Document.Range(rng.End, rngClose.start).ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
                                                rng.Document.Range(rng.End, rngClose.start).ParagraphFormat.LineSpacing = VBA.CSng(arr1)
                                            Else
                                                rng.Document.Range(rng.End, rngClose.start).ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
                                                rng.Document.Range(rng.End, rngClose.start).ParagraphFormat.LineSpacing = VBA.CSng(arr1)
                                            End If
                                        ElseIf VBA.Left(e, 10) = "font-size:" Then
                                            arr1 = VBA.Replace(VBA.LTrim(VBA.Mid(e, VBA.Len("font-size:") + 1)), "px", vbNullString)
                                            If VBA.IsNumeric(arr1) Then
                                                rng.Document.Range(rng.End, rngClose.start).font.Size = VBA.CSng(arr1)
                                            Else
                                                If arr1 = "x-small" Then
                                                    rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size * (2 / 3)
                                                ElseIf arr1 = "medium" Then
                                                    'Stop
                                                    '不處理，即預設大小
                                                Else
                                                    playSound 12
                                                    Stop 'to check
                                                End If
                                            End If
                                        Else
                                            SystemSetup.playSound 12
                                            rng.Select
                                            Debug.Print e
                                            Stop 'to check
                                        End If
                                    Next e
                                End If
                        End Select
                        rng.text = vbNullString: rngClose.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                
                '處理超連結
                If VBA.Len(p.Range.text) > VBA.Len("<a href=""></a>") Then
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                    '.text = "<a href="""
                    .text = "<a "
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        url = rng.text: e = rng.text
                        Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                        rngClose.Find.Execute "</a>"
                        url = getHTML_AttributeValue("href", url)
                        'url = getHTML_AttributeValue("<a href", p.Range.text)
                        e = getHTML_AttributeValue("title", VBA.CStr(e))
                        Select Case VBA.Left(url, 1)
                            Case "#"
                                If Not SeleniumOP.IsWDInvalid() Then
                                    url = WD.url & url
                                End If
                            Case "/"
                                url = domainUrlPrefix & url '路徑中多一個斜線（/）也是可以的，沒差 20241012
                            Case Else
                                If Not VBA.Left(url, 4) = "http" Then
                                    Stop 'check
                                    url = domainUrlPrefix & url
                                End If
                        End Select
                        
                        Set obj = rng.Document.Range(rng.start, rngClose.End).ShapeRange
                        rng.text = vbNullString: rngClose.text = vbNullString
                        If Not obj Is Nothing Then
                            Select Case obj.Count
                                Case 0
                                    If rng.Document.Range(rng.End, rngClose.start).text <> vbNullString Then
                                        rng.Document.Range(rng.End, rngClose.start).Hyperlinks.Add rng.Document.Range(rng.End, rngClose.start), url, , e
                                    End If
                                Case 1
                                    rng.Document.Range(rng.End, rngClose.start).Hyperlinks.Add obj(1), url, , e
                                Case Else
                                    playSound 12 'for check
                                    Stop
                            End Select
                            
                            Set obj = Nothing
                        Else
                            playSound 12 'for check
                            Stop
                        End If
                        rng.SetRange rngClose.End, p.Range.End
                    Loop
                End If

                If VBA.Len(p.Range.text) > VBA.Len("<p style=""padding-left:;>") Then
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                    .text = "<p style=""padding-left:"
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        p.Range.ParagraphFormat.IndentCharWidth 3
                        rng.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                If VBA.Len(p.Range.text) > VBA.Len("<span size=""></span>") Then
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                    .text = "<span size="
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                        rngClose.Find.text = "</span>"
                        rngClose.Find.Execute
                        
                        rng.text = vbNullString: rngClose.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                If VBA.Len(p.Range.text) > VBA.Len("<span></span>") Then
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                    .text = "<span>"
                    Do While .Execute()
                        Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                        rngClose.Find.text = "</span>"
                        rngClose.Find.Execute
                        rng.text = vbNullString: rngClose.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                If VBA.Len(p.Range.text) > VBA.Len("<p id=") Then
                    .text = "<p id="
                    rng.SetRange p.Range.start, p.Range.End
                    If .Execute() Then
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        'rng.Select
                        rng.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                    End If
                End If
                If VBA.Len(p.Range.text) > VBA.Len("<p style=""line-height:px;"">>") Then
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                    '.text = " style=""line-height: "
                    .text = "<p style="""
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        '借用url變數
                        url = getHTML_AttributeValue("style", p.Range.text)
                        arr = VBA.Split(url, ";")
                        For Each e In arr
                            e = VBA.Trim(e)
                            If VBA.Left(e, 13) = "line-height: " Then
                                rng.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
                                rng.ParagraphFormat.LineSpacing = CSng(VBA.Replace(VBA.Mid(e, VBA.Len("line-height: ") + 1), "px", vbNullString))
                            ElseIf VBA.Left(e, 11) = "font-size: " Then
                                rng.Paragraphs(1).Range.font.Size = VBA.CSng(VBA.Replace(VBA.Mid(e, VBA.Len("font-size: ") + 1), "px", vbNullString))
                            ElseIf VBA.Left(e, 11) = "margin-top:" Then
                                '不處理
                            Else
                                If e <> vbNullString Then
                                    playSound 12
                                    rng.Select
                                    Stop 'for check
                                End If
                            End If
                        Next e
                        'url = VBA.Replace(VBA.Replace(url, "line-height: ", vbNullString), "px;", vbNullString)
                        'rng.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
                        'rng.ParagraphFormat.LineSpacing = VBA.CSng(url)
                        rng.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                If VBA.Len(p.Range.text) > VBA.Len("<p dir="""">") Then
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                    .text = "<p dir="""
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        '借用url變數
                        If VBA.InStr(rng.text, "ltr") Then rng.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                If VBA.Len(p.Range.text) > VBA.Len("<p class="";"">") Then
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                    .text = "<p class="""
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        rng.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                If VBA.Len(p.Range.text) > VBA.Len("<st1:personname ></st1:personname>") Then
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                    .text = "<st1:personname "
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                        rngClose.Find.ClearFormatting
                        rngClose.Find.Execute "</st1:personname>"
                        rng.text = vbNullString: rngClose.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                'rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                
            End With 'rng.Find
            
            
            If .Paragraphs(1).Range.text = "<br class=""Apple-interchange-newline""> " & Chr(13) Then
                .Paragraphs(1).Range.text = vbNullString
            End If
        End With 'rng
    Next p
    
    Set rng = rngHtml.Document.Range(rngHtml.start, rngHtml.End)
    文字處理.FixFontname rng

    
finish:
    rngHtml.Document.Range(stRngHTML, stRngHTML).Select '回到起始位置
    
    Rem just for check
    With rngHtml.Find
        .ClearFormatting
        If .Execute("[<>&;]", , , True) Then
            rngHtml.Select
            SystemSetup.playSound 3
        End If
    End With
End Sub
Rem 20241011 HTML 無序清單的處理.Porc=Porcess
Private Sub unorderedListPorc_HTML2Word(rngHtml As Range)
    Rem 無序清單的處理
    Dim rngUnorderedList As Range, st As Long, ed As Long, rngUnorderedListSub As Range, p As Paragraph
    Do
        Set rngUnorderedList = GetRangeFromULToUL_UnorderedListRange(rngHtml)
        If Not rngUnorderedList Is Nothing Then
            st = rngUnorderedList.start
            Set p = rngUnorderedList.Paragraphs(1).Previous
            If Not p Is Nothing Then
                '如果是易學網的「歷代注本：」
                If VBA.InStr(p.Range.text, "歷代注本：") Then
                    With rngUnorderedList.Find
                        .Execute "<li>", , , , , , , , , vbNullString, wdReplaceAll
                        .Execute "</li>", , , , , , , , , vbNullString, wdReplaceAll
                        .Execute "</ul>", , , , , , , , , vbNullString, wdReplaceAll
                         ed = rngUnorderedList.End
                    End With
                    Set rngUnorderedListSub = rngUnorderedList.Document.Range(rngUnorderedList.start, rngUnorderedList.End)
                    rngUnorderedListSub.Find.ClearFormatting
                    If rngUnorderedListSub.Find.Execute("<ul ") Then
                        rngUnorderedListSub.MoveEndUntil ">"
                        rngUnorderedListSub.End = rngUnorderedListSub.End + 2
                        If rngUnorderedListSub.Characters(rngUnorderedListSub.Characters.Count) <> Chr(13) Then
                            rngUnorderedListSub.End = rngUnorderedListSub.End - 1
                        End If
                        rngUnorderedListSub.text = vbNullString

                    Else
                        rngUnorderedListSub.SetRange rngUnorderedList.start, rngUnorderedList.End
                        If rngUnorderedListSub.Find.Execute("<ul>") Then
                            If rngUnorderedListSub.Paragraphs(1).Range.text = rngUnorderedListSub & Chr(13) Then
                                rngUnorderedListSub.Paragraphs(1).Range.text = vbNullString
                            Else
                                rngUnorderedListSub.text = vbNullString
                            End If
                        End If
                    End If
                    If rngUnorderedList.Characters(rngUnorderedList.Characters.Count) = Chr(13) Then
                        rngUnorderedList.End = rngUnorderedList.End - 1
                    End If
                    With rngUnorderedList
                        '.Hyperlinks.Add rngLink, iwe.GetAttribute("href")'在前面已經插入超連結了
                        .Style = wdStyleHeading2 '標題 2
                        .font.Size = 18
                        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle '單行間距
                    End With
                Else
                    GoTo UnorderedListRange
                End If
            Else
UnorderedListRange:
                
                rngUnorderedList.Select 'for chect
                'Stop
                
                'Set rngUnorderedList = Nothing
                'InsertHTMLList rngUnorderedList.text
                
                If VBA.Left(rngUnorderedList, 5) = "<ul>" & Chr(13) And VBA.Right(rngUnorderedList, 6) = Chr(13) & "</ul>" Then
                    With rngUnorderedList
                        With .Find
                            .Execute "<li>", , , , , , , , , vbNullString, wdReplaceAll
                            .Execute "</li>", , , , , , , , , vbNullString, wdReplaceAll
                            .Execute "<ul>^p", , , , , , , , , vbNullString, wdReplaceAll
                            .Execute "^p</ul>", , , , , , , , , vbNullString, wdReplaceAll
                        End With
                        .ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
                            ListGalleries(wdBulletGallery).ListTemplates(1), ContinuePreviousList:= _
                            False, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
                            wdWord10ListBehavior
                        
                    End With
                Else
                    Stop 'for chect
                
                End If
            End If
        End If
    Loop Until rngUnorderedList Is Nothing
End Sub
Rem 解析HTML並插入清單 20241011 creedit_with_Copilot大菩薩：https://sl.bing.net/gbeqh0TAks8：HTML表格轉換和屬性設置
Rem 解析HTML內容，提取清單項目，然後在Word中插入相應的清單樣式。https://sl.bing.net/bhFU3zNMSom
Sub InsertHTMLList(html As String)
    Dim doc As Document
    Dim listItems As Collection
    Dim listItem As Variant
    Dim rng As Range
    
    ' 解析HTML
    Set listItems = ParseHTMLList(html)
    
    ' 插入清單
    Set doc = ActiveDocument
    Set rng = doc.Range(start:=doc.Content.End - 1, End:=doc.Content.End - 1)
    
    ' 開始清單
    rng.ListFormat.ApplyBulletDefault
    
    ' 填充清單內容
    For Each listItem In listItems
        rng.text = StripHTMLTags(VBA.CStr(listItem))
        rng.ParagraphFormat.Alignment = wdAlignParagraphLeft
        rng.font.Name = "標楷體"
        rng.InsertParagraphAfter
        Set rng = rng.Next(wdParagraph, 1) '.Range
    Next listItem
End Sub

Function ParseHTMLList(html As String) As Collection
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim listItems As New Collection
    
    ' 初始化正則表達式對象
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "<li.*?>(.*?)</li>"
    
    Set matches = regex.Execute(html)
    For Each match In matches
        listItems.Add match.SubMatches(0)
    Next match
    
    Set ParseHTMLList = listItems
End Function


Rem 將HTML文本置換成圖片，成功則傳回一個有效了 InlineShape物件 20241011 textPart:要解析的HTML文本，rng：要插入圖片的位置；domainUrlPrefix 是否圖片網址要加域名前綴
Private Function insert_ImageHTML(textPart As String, rng As Range, Optional domainUrlPrefix As String) As word.inlineShape
    Dim url As String, w As Single, h As Single, align As String, hspace As String
    'url = getImageUrl(textPart)
    url = getHTML_AttributeValue("src", textPart)
    If VBA.InStr(textPart, "width") Then
        w = VBA.CSng(getHTML_AttributeValue("width", textPart))
    End If
    If VBA.InStr(textPart, "height") Then
        h = VBA.CSng(getHTML_AttributeValue("height", textPart))
    End If
    If VBA.InStr(textPart, "align") Then
        align = getHTML_AttributeValue("align", textPart)
    End If
    If VBA.InStr(textPart, "hspace") Then
        hspace = getHTML_AttributeValue("hspace", textPart)
    End If
    
    If VBA.InStr(url, "http") <> 1 Then
'        If domainUrlPrefix = vbNullString Then
'            'msgbox "須帶入網域前綴才行"
'            'If domainUrlPrefix = vbNullString Then domainUrlPrefix = "https://www.eee-learning.com"
'
'            'If Not SeleniumOP.IsWDInvalid() Then
'                'domainUrlPrefix = getDomainUrlPrefix(SeleniumOP.WD.url)
'            'End If
'
'        End If
        If Not IsBase64Image(url) Then 'base64編碼的圖片
            url = domainUrlPrefix & url '路徑中多一個斜線（/）也是可以的，沒差 20241012
        Else
            If Base64ToImage(url, VBA.Environ("TEMP") & "\" & "tempImage.png") = False Then
                Stop
'                GoTo finish
                Exit Function
            End If
        End If
    End If
    Dim inlsp As inlineShape
    
    If Not IsBase64Image(url) Then 'VBA.InStr(url, "data:image/png;base64") = 0 Then
            'rng.InlineShapes.AddPicture fileName:=url, _
                            LinkToFile:=False, SaveWithDocument:=True
        If w > 0 And h > 0 Then
            Set inlsp = rng.InlineShapes.AddPicture(fileName:=url, _
                            LinkToFile:=False, SaveWithDocument:=True)
            resizePicture rng, inlsp, url, w, h
        Else
            On Error Resume Next
            Set inlsp = rng.InlineShapes.AddPicture(fileName:=url, _
                            LinkToFile:=False, SaveWithDocument:=True)
            On Error GoTo 0
            If Not inlsp Is Nothing Then
                If inlsp.Range.tables.Count > 0 Then
                    'resizePicture rng, inlsp, url, inlsp.Range.tables(1).PreferredWidth, inlsp.height * (inlsp.width / inlsp.Range.tables(1).PreferredWidth)
                    Rem 先插表格並處理其中的圖片，應該預設就是表格大小
                Else
                    resizePicture rng, inlsp, url
                End If
            Else
                Exit Function
            End If
        End If
    Else 'base64編碼的圖片
        
        ' 插入base64編碼的圖片
        Set inlsp = InsertBase64Image(url, "tempImage.png", rng)
        resizePicture rng, inlsp, url
        
    End If
    
    Rem 設定圖片格式
    Rem inlineShape格式
    Dim shp As Shape
    If align <> vbNullString And hspace <> vbNullString Then
        Select Case align
            Case "right"
                Set shp = inlsp.ConvertToShape
                With shp.WrapFormat
                    .Type = wdWrapSquare
                    .Side = wdWrapBoth
                    '.DistanceTop = CentimetersToPoints(0.5)
                    .DistanceLeft = CentimetersToPoints(0.5)
                    .Parent.RelativeHorizontalPosition = wdRelativeHorizontalPositionRightMarginArea
                    .Parent.RelativeVerticalPosition = wdRelativeVerticalPositionTopMarginArea
                    .Parent.Left = WdShapePosition.wdShapeRight
                    'shp.Top = WdShapePosition.wdShapeTop
    '                shp.Left = ActiveDocument.PageSetup.PageWidth - shp.width - CentimetersToPoints(1) ' 設定右邊距離
    '                shp.Top = CentimetersToPoints(1) ' 設定上邊距離
                End With
            Case "left"
                Set shp = inlsp.ConvertToShape
                With shp.WrapFormat
                    .Type = wdWrapSquare
                    .Side = wdWrapBoth
                    '.DistanceTop = CentimetersToPoints(0.5)
                    .DistanceRight = CentimetersToPoints(0.5)
                
                    .Parent.RelativeHorizontalPosition = wdRelativeHorizontalPositionLeftMarginArea
                    .Parent.RelativeVerticalPosition = wdRelativeVerticalPositionTopMarginArea
                    .Parent.Left = WdShapePosition.wdShapeLeft
    '                shp.Top = wdShapeTop
                End With
            Case "absbottom"
            Case Else
                playSound 12
                Stop 'for check
        End Select
    End If
    Rem Shape文繞圖格式
    Dim imgStyle As String, float As String, marginLeft, marginRight
    'ex: float:right;margin-left:10px;margin-right:10px;"
    imgStyle = getHTML_AttributeValue("style", textPart)
    If imgStyle <> vbNullString Then
        If inlsp.Range.tables.Count = 0 Then
            If InStr(imgStyle, "float:") Then
                float = VBA.Mid(imgStyle, VBA.InStr(imgStyle, "float:") + VBA.Len("float:"), VBA.InStr(VBA.InStr(imgStyle, "float:"), imgStyle, ";") - (VBA.InStr(imgStyle, "float:") + VBA.Len("float:")))
            End If
            If InStr(imgStyle, "margin-left:") Then
                marginLeft = VBA.Val(VBA.Mid(imgStyle, VBA.InStr(imgStyle, "margin-left:") + VBA.Len("margin-left:"), VBA.InStr(VBA.InStr(imgStyle, "margin-left:"), imgStyle, ";") - (VBA.InStr(imgStyle, "margin-left:") + VBA.Len("margin-left:"))))
            End If
            If InStr(imgStyle, "margin-right:") Then
                marginRight = VBA.Val(VBA.Mid(imgStyle, VBA.InStr(imgStyle, "margin-right:") + VBA.Len("margin-right:"), VBA.InStr(VBA.InStr(imgStyle, "margin-right:"), imgStyle, ";") - (VBA.InStr(imgStyle, "margin-right:") + VBA.Len("margin-right:"))))
            End If
            If float <> "" And VBA.IsEmpty(marginLeft) = False And VBA.IsEmpty(marginRight) = False Then
                ' 設置圖片的文繞圖方式和對齊方式
                Set shp = inlsp.ConvertToShape
                With shp
                    .WrapFormat.Type = WdWrapType.wdWrapTight ' wdWrapSquare
                    Select Case float
                        Case vbNullString
                        Case "left"
                            .Left = WdShapePosition.wdShapeLeft
                            '.WrapFormat.Side = WdWrapSideType.wdWrapLeft
                        Case "right"
                            .Left = WdShapePosition.wdShapeRight
                            '.WrapFormat.Side = WdWrapSideType.wdWrapRight ' 對應於float:right
                        Case Else
                            Stop ' check
                    End Select
                    If marginLeft <> 0 Then
                        .WrapFormat.DistanceLeft = marginLeft ' 對應於margin-left:10px
                    End If
                    If marginRight <> 0 Then
                        .WrapFormat.DistanceRight = marginRight ' 對應於margin-right:10px
                    End If
                End With
            End If
        End If
    End If
    
    Set insert_ImageHTML = inlsp
    SystemSetup.playSound 0.411
End Function
Rem 解析HTML內容，提取表格、行、單元格、圖片和文字 20241011 creedit_with_Copilot大菩薩：https://sl.bing.net/fQ5lVr8PLye
Function ParseHTMLTable(html As String) As Collection
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim tables As New Collection
    Dim rows As New Collection
    Dim cells As New Collection
    Dim table, row
    
    ' 初始化正則表達式對象
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    
    ' 匹配表格
    regex.Pattern = "<table.*?>(.*?)</table>"
    Set matches = regex.Execute(html)
    For Each match In matches
        tables.Add match.SubMatches(0)
    Next match
    
    ' 匹配行/列
    regex.Pattern = "<tr.*?>(.*?)</tr>"
    For Each table In tables
        Set matches = regex.Execute(table)
        For Each match In matches
            rows.Add match.SubMatches(0)
        Next match
    Next table
    
    ' 匹配單元格
    regex.Pattern = "<td.*?>(.*?)</td>"
    For Each row In rows
        Set matches = regex.Execute(row)
        For Each match In matches
            cells.Add match.SubMatches(0)
        Next match
    Next row
    
    Set ParseHTMLTable = cells
End Function
Rem 接下來，您可以在Word中創建表格並插入相應的內容 creedit_with_Copilot大菩薩 20241011
Sub InsertHTMLTable(rngHtml As Range, Optional domainUrlPrefix As String)
    Dim html As String
    Dim tbl As word.table
    Dim cells As Collection
    Dim cell As Variant
    Dim row As Integer
    Dim col As Integer
    Dim img As inlineShape
    Dim rngTxt As Range
    Dim c As cell
    Dim align As String
    Dim bgcolor As String
    Dim tblWidth As Single
'    Dim imgSrc As String
'    Dim imgWidth As Single
'    Dim imgHeight As Single
    
    
'    Dim ur As UndoRecord
'    SystemSetup.stopUndo ur, "InsertHTMLTable"
    
    html = rngHtml.text
    ' 解析HTML
    Set cells = ParseHTMLTable(html)
    
    ' 插入表格
    rngHtml.text = vbNullString
    
    Set tbl = rngHtml.tables.Add(Range:=Selection.Range, NumRows:=1, NumColumns:=1)
    
     ' 設置表格屬性
'    align = getHTML_AttributeValue("align", html)
'    bgcolor = getHTML_AttributeValue("bgcolor", html)
'    tblWidth = CSng(getHTML_AttributeValue("width", html))
    align = getHTML_AttributeValue("align", html)
    bgcolor = getHTML_AttributeValue("bgcolor", html)
    tblWidth = VBA.CSng(VBA.Val((getHTML_AttributeValue("width", html, ":"))))
    
    
    If align = "left" Then
        tbl.rows.Alignment = wdAlignRowLeft
    ElseIf align = "center" Then
        tbl.rows.Alignment = wdAlignRowCenter
    ElseIf align = "right" Then
        tbl.rows.Alignment = wdAlignRowRight
    End If
    
    If bgcolor <> "" Then
        If VBA.Left(bgcolor, 1) = "#" Then
            Dim arr
            arr = colorCodetoRGB(bgcolor)
            tbl.Shading.BackgroundPatternColor = RGB(arr(0), arr(1), arr(2))
        Else
            If bgcolor = "white" Then
                tbl.Shading.BackgroundPatternColor = RGB(255, 255, 255)
            Else
                playSound 12 'for check
                Stop
            End If
        End If
    End If
    
'    Dim shp As Shape
    ' 將表格轉換為Shape對象
'    Set shp = tbl.ConvertToShape
    tbl.rows.WrapAroundText = True
    ' 設置文繞圖方式
'    shp.WrapFormat.Type = wdWrapSquare
'    shp.WrapFormat.Side = wdWrapBoth
'    shp.WrapFormat.DistanceTop = 0
'    shp.WrapFormat.DistanceBottom = 0
'    shp.WrapFormat.DistanceLeft = 0
'    shp.WrapFormat.DistanceRight = 0
    
    tbl.PreferredWidthType = wdPreferredWidthPoints
    tbl.PreferredWidth = tblWidth
    
    ' 填充表格內容
    row = 1
    col = 1
    For Each cell In cells
        ' 檢查是否包含圖片
        If InStr(cell, "<img") > 0 Then
            
            Set c = tbl.cell(row, col)
            Set img = insert_ImageHTML(html, c.Range, domainUrlPrefix)
'            imgSrc = getHTML_AttributeValue("src", VBA.CStr(cell))  'Mid(cell, InStr(cell, "src=") + 5, InStr(cell, """", InStr(cell, "src=") + 5) - InStr(cell, "src=") - 5)
'            imgWidth = getHTML_AttributeValue("width", VBA.CStr(cell)) 'CSng(Mid(cell, InStr(cell, "width=") + 7, InStr(cell, """", InStr(cell, "width=") + 7) - InStr(cell, "width=") - 7))
'            imgHeight = getHTML_AttributeValue("height", VBA.CStr(cell)) 'CSng(Mid(cell, InStr(cell, "height=") + 8, InStr(cell, """", InStr(cell, "height=") + 8) - InStr(cell, "height=") - 8))
'            tbl.cell(row, col).Range.InlineShapes.AddPicture fileName:=imgSrc, LinkToFile:=False, SaveWithDocument:=True
            c.Range.InlineShapes(1).width = img.width 'imgWidth
            c.Range.InlineShapes(1).height = img.height 'imgHeight
            Set rngTxt = c.Range.Document.Range(c.Range.End - 1, c.Range.End - 1)
            rngTxt.text = StripHTMLTags(VBA.CStr(cell))
        Else
            tbl.cell(row, col).Range.text = StripHTMLTags(VBA.CStr(cell))
        End If
        col = col + 1
        If col > tbl.Columns.Count Then
            tbl.rows.Add
            row = row + 1
            col = 1
        End If
    Next cell
    
'    SystemSetup.contiUndo ur
End Sub
Rem 顏色碼轉換成RGB
Private Function colorCodetoRGB(colorCode As String) As Long()
    ' 將bgcolor轉換為RGB顏色
    'Dim r As Integer, g As Integer, b As Integer
    If VBA.Left(colorCode, 1) <> "#" Then Exit Function
    Dim arr(2) As Long
    arr(0) = CLng("&H" & Mid(colorCode, 2, 2))
    arr(1) = CLng("&H" & Mid(colorCode, 4, 2))
    arr(2) = CLng("&H" & Mid(colorCode, 6, 2))
    colorCodetoRGB = arr
End Function


Rem 取得HTML中表格的屬性值 20241011 creedit_with_Copilot大菩薩：HTML表格轉換和屬性設置：HTML表格轉換和屬性設置
Function getHTMLAttributeValue(attributeName As String, html As String) As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    
    ' 初始化正則表達式對象
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.Pattern = attributeName & "=""'[""']"
    
    Set matches = regex.Execute(html)
    If matches.Count > 0 Then
        getHTMLAttributeValue = matches(0).SubMatches(0)
    Else
        getHTMLAttributeValue = ""
    End If
End Function
Rem 清除一切的html tags HTML標籤
Function StripHTMLTags(html As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "<.*?>"
    regex.Global = True
    StripHTMLTags = regex.Replace(html, "")
End Function

Rem 20241010國慶日 清除在標籤間沒有任何內容的HTML空標籤
Sub RemoveEmptyTags(rng As Range)
    Dim rngOriginal As Range, arr, e
    Set rngOriginal = rng.Document.Range(rng.start, rng.End)
    arr = Array(">" & VBA.Chr(11) & "<", "><", ">" & VBA.Chr(13) & "<")
    rng.Find.MatchWildcards = False
    With rng.Find
        .ClearFormatting
        For Each e In arr
            .text = e
            .Wrap = wdFindStop
            Do While .Execute()
                Do Until rng.Characters(1) = "<"
                    rng.MoveStart , -1
                Loop
'                rng.Select 'for check
                rng.MoveEndUntil ">"
                rng.MoveEnd 1
'                rng.Select 'for check
                If Not VBA.Left(rng.text, 2) = "</" And (VBA.InStr(rng.text, e & "/") Or VBA.InStr(rng.text, ">" & VBA.Chr(13) & "/>")) _
                        And VBA.Mid(rng.text, VBA.InStr(rng.text, "/") + 1, VBA.Len(rng.text) - VBA.InStr(rng.text, "/") - 1) _
                            = rng.Document.Range(rng.start + 1, rng.start + 1 + VBA.Len(rng.text) - VBA.InStr(rng.text, "/") - 1) Then
                    rng.text = vbNullString
                End If
                If rng.Characters.Count = 1 And rng.Characters(1).text = VBA.Chr(13) And rng.Paragraphs(1).Range.Characters.Count = 1 Then
                    rng.Characters(1).text = vbNullString
                End If
                rng.Collapse wdCollapseEnd
                'rng.SetRange rng.End, rngOriginal.End
            Loop
            rng.SetRange rngOriginal.start, rngOriginal.End
        Next e
    End With
End Sub
Rem 取得無序列表（<ul></ur>）的範圍 20241010creedit_with_Copilot大菩薩：HTML超連結轉換成Word VBA：https://sl.bing.net/bXsbFqI2cz6
Function GetRangeFromULToUL_UnorderedListRange(rng As Range) As Range
    Dim startRange As Range
    Dim endRange As Range
    
    ' 查找 <ul> 標籤
    Set startRange = rng.Document.Range(rng.start, rng.End)
    With startRange.Find
        .ClearFormatting
        .text = "<ul"
        If .Execute Then
            startRange.Collapse Direction:=wdCollapseStart
        End If
    End With
    
    ' 查找 </ul> 標籤
    Set endRange = rng.Document.Range(startRange.End, rng.End)
    With endRange.Find
        .ClearFormatting
        .text = "</ul>"
        If .Execute Then
            endRange.Collapse Direction:=wdCollapseEnd
        End If
    End With
    
    ' 設定範圍
    If Not (startRange.start = rng.start And endRange.End = rng.End) Then
        Set GetRangeFromULToUL_UnorderedListRange = rng.Document.Range(startRange.start, endRange.End)
    End If
End Function


Rem 判斷圖檔是否有效 creedit_with_Copilot大菩薩：20241010 https://sl.bing.net/emhYXUvuos8 https://sl.bing.net/cG4Jn2MciZ2
Function IsValidImage_LoadPicture(filePath As String) As Boolean
    On Error Resume Next
    Dim img As Object
    Set img = stdole.LoadPicture(filePath) 'err:481+圖片不正確，仍可能是有效圖檔！
    '如 Set img = stdole.LoadPicture("C:\Users\oscar\Documents\CtextTempFiles\Ctext_Page_Image●.png") 'err:481+圖片不正確，仍是有效圖檔！
    Rem 大概僅支援 jpg！！20241010
    IsValidImage_LoadPicture = Not img Is Nothing
    On Error GoTo 0
End Function
Rem creedit_with_Copilot大菩薩：WordVBA+SeleniumBasic讀入網頁內容圖片與超連結：https://sl.bing.net/fWOLN5PwHsG
Rem  啟動Chrome瀏覽器並導航到圖片URL,失敗則傳回false。這可用，但須取得Chrome瀏覽器下載目錄才行
Function DownloadImage_chromedriverExecuteScript(url As String, filePath As String) As Boolean
'    Dim driver As New SeleniumBasic.ChromeDriver
'    driver.start "Chrome"
'    driver.Get url
    Dim driver As SeleniumBasic.IWebDriver, currentWin As String
    If Not SeleniumOP.IsWDInvalid Then
        Set driver = SeleniumOP.WD
        currentWin = driver.CurrentWindowHandle
    Else
        If SeleniumOP.OpenChrome(url) Then
            Set driver = SeleniumOP.WD
        Else
            Exit Function
        End If
    End If
    
    ' 等待圖片加載完成  'Application是我自行做的Excel模組中的物件。本專案並沒引用 Excel
    'Excel.Application.wait (Now + TimeValue("0:00:05"))
    SystemSetup.wait (Now + TimeValue("0:00:02"))
    
    ' 下載圖片 rem 可以正常下載，只是要取得Chrome瀏覽器的下載路徑才能供後續使用！20241010
    driver.ExecuteScript "var link = document.createElement('a'); link.href = arguments[0]; link.download = arguments[1]; document.body.appendChild(link); link.click();", url, filePath
    ' 等待下載完成
    'Excel.Application.wait (Now + TimeValue("0:00:02"))
    SystemSetup.wait (Now + TimeValue("0:00:02"))
    If VBA.Dir(filePath) = vbNullString Or IsValidImage_LoadPicture(url) Then
        Stop
    Else
        DownloadImage_chromedriverExecuteScript = True
    End If
    
    'driver.Quit
    driver.Close
    If currentWin <> vbNullString Then
        If IsWDInvalid() Then
            driver.SwitchTo.Window driver.CurrentWindowHandle
        Else
            WD.SwitchTo.Window currentWin
        End If
    End If
End Function
Rem 將使用 XMLHTTP 來下載圖片，然後將其保存為暫存文件.若失敗傳回false 20241010 creedit_with_Copilot大菩薩：https://sl.bing.net/caezeDQDlfg
Function DownloadImage_XMLHTTP_url(url As String, filePath As String) As Boolean
    Dim xmlHttp As Object
    Dim stream As Object
    
    ' 創建 XMLHTTP 對象
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP")
    xmlHttp.Open "GET", url, False
    xmlHttp.send '-2147467259 無法指出的錯誤。可能是由於您使用的是 base64 編碼的 URL。XMLHTTP 無法直接處理 base64 編碼的圖像數據 https://sl.bing.net/dd1AOLdKBaK
    
    ' 創建 ADODB.Stream 對象
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    stream.Write xmlHttp.responseBody
    stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    stream.Close
    If VBA.Dir(filePath) <> vbNullString And IsValidImage_LoadPicture(filePath) Then
        DownloadImage_XMLHTTP_url = True
    End If
End Function
Rem 20241010 Copilot大菩薩：使用 ServerXMLHTTP: 有時候 MSXML2.XMLHTTP 會有問題，您可以嘗試使用 MSXML2.ServerXMLHTTP 來代替。https://sl.bing.net/dvbNeuzNEjc
Rem XMLHTTP 無法直接處理 base64 編碼的圖像數據。您需要先將 base64 編碼的數據解碼，然後再將其保存為圖像文件。https://sl.bing.net/b9gYh5mICbc
Function DownloadImage_XMLHTTP(url As String, filePath As String) As Boolean
    On Error GoTo ErrorHandler
    Dim xmlHttp As Object
    Dim stream As Object
    
    ' 創建 ServerXMLHTTP 對象
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    xmlHttp.Open "GET", url, False
    xmlHttp.send
    
    ' 創建 ADODB.Stream 對象
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    stream.Write xmlHttp.responseBody
    stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    stream.Close
    If VBA.Dir(filePath) <> vbNullString And IsValidImage_LoadPicture(filePath) Then
        DownloadImage_XMLHTTP = True
    End If
    Exit Function
    
ErrorHandler:
    Debug.Print Err.Number & Err.Description
    MsgBox "Error: " & Err.Description
    DownloadImage_XMLHTTP = False
End Function

Rem  將下載的圖片插入到Word文件中
Sub InsertDownloadedImage(url As String, filePath As String, rng As Range)
    ' 下載圖片
    DownloadImage_chromedriverExecuteScript url, filePath
    
    ' 插入圖片
    rng.InlineShapes.AddPicture fileName:=filePath, LinkToFile:=False, SaveWithDocument:=True
    
    ' 刪除暫存文件
    Kill filePath
End Sub


Rem 20241010 creedit_with_Copilot大菩薩：WordVBA+SeleniumBasic讀入網頁內容圖片與超連結：https://sl.bing.net/htsW1HREBOe
'Base64ToBinary：將base64編碼的圖片數據轉換為二進制數據。
'SaveBinaryAsFile：將二進制數據保存為臨時圖片文件。
'InsertBase64Image：將臨時圖片文件插入到Word文件中，並設置圖片的寬度和高度。
'主程式：調用上述方法來插入base64編碼的圖片。
'這樣，您就可以將base64編碼的圖片插入到Word文件中，並保持其格式和內容。
Rem Copilot大菩薩：使用正則表達式來判斷URL是否包含base64編碼的圖片數據的範例：https://sl.bing.net/c6eMOTPP4wK
Function IsBase64Image(url As String) As Boolean '應該是解析時出錯了，無效！20241010（國慶日）
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp") 'If VBA.InStr(url, "data:image/png;base64")  Then 'base64編碼的圖片
    regex.Pattern = "^data:image\/(png|jpg|jpeg|gif);base64," '可以更靈活地判斷是否是base64編碼的圖片
    regex.IgnoreCase = True
    IsBase64Image = regex.test(url)
End Function
Rem 20241010 creedit_with_Copilot大菩薩：解決WordVBA + Selenium下載Chrome瀏覽器網頁中的圖片問題：https://sl.bing.net/dJlhQRbUOHI
Rem 解碼 base64 編碼的圖像數據並將其保存為文件：https://sl.bing.net/c3GitspZG8G
Function Base64ToImage(base64String As String, filePath As String) As Boolean
    On Error GoTo ErrorHandler
    Dim binaryData() As Byte
    Dim stream As Object
    
    If VBA.InStr(base64String, "data:image/png;base64,") Then
        ' 去掉 base64 頭部的 "data:image/png;base64,"
        base64String = Replace(base64String, "data:image/png;base64,", "")
    Else
        Stop
    End If
    
    ' 解碼 base64 字符串
    binaryData = Base64Decode(base64String)
    
    ' 創建 ADODB.Stream 對象
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    stream.Write binaryData
    stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    stream.Close
    
    If VBA.Dir(filePath) <> vbNullString Then
        If VBA.Right(filePath, 3) <> "png" Then
            Base64ToImage = IsValidImage_LoadPicture(filePath)
        Else
            Base64ToImage = True
        End If
    End If
    If Base64ToImage = True Then base64String = filePath
    Exit Function
    
ErrorHandler:
    MsgBox "Error: " & Err.Description
    Base64ToImage = False
End Function

Function Base64Decode(base64String As String) As Byte()
    Dim xml As Object
    Dim node As Object
    
    ' 創建 MSXML2.DOMDocument 對象
    Set xml = CreateObject("MSXML2.DOMDocument")
    Set node = xml.createElement("base64")
    node.DataType = "bin.base64"
    node.text = base64String
    Base64Decode = node.nodeTypedValue
End Function

Rem 解析base64編碼'失敗！
Function Base64ToBinary(base64String As String) As Byte()
    Dim xmlObj As Object
    Dim base64Data As String
    
    ' 去掉前綴部分
    base64Data = Mid(base64String, InStr(base64String, ",") + 1)
    
    ' 解析base64編碼
    Set xmlObj = CreateObject("MSXML2.DOMDocument.6.0")
    xmlObj.LoadXML "<root><binary>" & base64Data & "</binary></root>"
    Base64ToBinary = xmlObj.DocumentElement.ChildNodes(0).nodeTypedValue
End Function
Rem 保存為臨時文件 將二進制數據保存為臨時圖片文件。'失敗！
Function SaveBinaryAsFile(binaryData() As Byte, filePath As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    stream.Write binaryData
    stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    stream.Close
End Function
Rem 插入base64編碼圖片,將臨時圖片文件插入到Word文件中
Function InsertBase64Image(base64String As String, filePath As String, rng As Range) As inlineShape
    Dim binaryData() As Byte
    Dim tempFilePath As String
    
    ' 解析base64編碼
    binaryData = Base64ToBinary(base64String)
    
    ' 保存為臨時文件
    tempFilePath = Environ("TEMP") & "\" & filePath
    SaveBinaryAsFile binaryData, tempFilePath
    
    ' 插入圖片
    Set InsertBase64Image = rng.InlineShapes.AddPicture(fileName:=tempFilePath, LinkToFile:=False, SaveWithDocument:=True)
    base64String = tempFilePath
    ' 刪除臨時文件
    Kill tempFilePath
End Function


Rem 20241009 取得HTML中的屬性之值 pro 不包含「="」
Private Function getHTML_AttributeValue(atrb As String, textIncludingAttribute As String, Optional marker As String)
    Dim lenatrb As Byte
    Select Case marker
        Case vbNullString
            atrb = atrb & "="""
        Case ":"
        atrb = atrb & ": "
    End Select
    If VBA.InStr(textIncludingAttribute, atrb) > 0 Then
        lenatrb = VBA.Len(atrb)
        getHTML_AttributeValue = VBA.Mid(textIncludingAttribute, VBA.InStr(textIncludingAttribute, atrb) + lenatrb, _
            VBA.InStr(VBA.InStr(textIncludingAttribute, atrb) + lenatrb, textIncludingAttribute, """") - (VBA.InStr(textIncludingAttribute, atrb) + lenatrb))
    End If
End Function

Rem 插入圖片後，根據前後字型大小自動調整圖片大小 20241009 creedit_with_Copilot大菩薩：WordVBA 圖片自動調整大小：https://sl.bing.net/e1S3H59hvI4
Private Function getImageUrl(textIncludingSrc As String)
    getImageUrl = VBA.Mid(textIncludingSrc, VBA.InStr(textIncludingSrc, "src=""") + 5, _
        VBA.InStr(VBA.InStr(textIncludingSrc, "src=""") + 5, textIncludingSrc, """") - (VBA.InStr(textIncludingSrc, "src=""") + 5))
End Function
Function getDomainUrlPrefix(url As String)
    getDomainUrlPrefix = VBA.Left(url, VBA.InStr(url, "//")) & "/" & VBA.Mid(url, VBA.InStr(url, "//") + 2, _
                VBA.InStr(VBA.InStr(url, "//") + 2, url, "/") - (VBA.InStr(url, "//") + 2))
End Function
Rem 重新調整圖片大小，若無指定 width與height 則參考前後文字型大小平均值設定
Private Sub resizePicture(rng As Range, pic As inlineShape, url As String, Optional width As Single = 0, Optional height As Single = 0)
    If width > 0 And height > 0 Then
        pic.width = width
        pic.height = height
    Else
    
        Dim fontSizeBefore As Single
        Dim fontSizeAfter As Single
        Dim avgFontSize As Single
        ' 獲取前後字型大小
        If rng.start > 1 Then
            fontSizeBefore = rng.Characters.First.Previous.font.Size
        Else
            fontSizeBefore = rng.Characters.First.font.Size
        End If
    
        If rng.End < rng.Document.Content.End Then
            fontSizeAfter = rng.Characters.Last.Next.font.Size
        Else
            fontSizeAfter = rng.Characters.Last.font.Size
        End If
    
        ' 計算平均字型大小
        avgFontSize = (fontSizeBefore + fontSizeAfter) / 2
    
        ' 調整圖片大小
        pic.LockAspectRatio = msoTrue
        If Not IsValidImage_LoadPicture(url) Then
            pic.height = avgFontSize
            If Not SeleniumOP.IsWDInvalid() Then
                pic.Range.Hyperlinks.Add pic.Range, WD.url
            End If
        Else
            pic.height = avgFontSize * 2 ' 根據需要調整比例
            pic.width = pic.height * pic.width / pic.height
        End If
    End If
End Sub

Rem 20241006 《看典古籍·古籍全文檢索》 Ctrl + k,d
Sub 查看典古籍古籍全文檢索()
    文字處理.ResetSelectionAvoidSymbols
    SeleniumOP.KandiangujiSearchAll Selection.text
End Sub
Rem 20241006 檢索《漢籍全文資料庫》 Alt + Shfit + h
Sub 查漢籍全文資料庫()
    文字處理.ResetSelectionAvoidSymbols
    SeleniumOP.HanchiSearch Selection.text
End Sub
Rem 20241006 以Google檢索《中國哲學書電子化計劃》Alt + t
Sub 查中國哲學書電子化計劃網域()
    文字處理.ResetSelectionAvoidSymbols
    中國哲學書電子化計劃.SearchSite
End Sub
Rem 20241006 rng 要處理的範圍，傳回結束的位置
Private Function 讀入網路資料後_於其後植入網址及設定格式(rng As Range, url As String, fontsize As Single) As Long
    Dim rngNote As Range
    Set rngNote = rng.Document.Range(rng.start, rng.End)
    With rngNote
        '網址格式設定
        If VBA.Len(.Paragraphs.Last.Range.text) > 1 Then .InsertParagraphAfter
        If .Paragraphs.Count > 1 Then
            rngNote.SetRange rng.Paragraphs.Last.Range.start, rng.Paragraphs.Last.Range.End
        End If
        .InsertAfter url '插入網址
        .InsertParagraphAfter
        .End = .End - 1
        If .Characters(1) = Chr(13) Then .start = .start + 1
        .font.Size = fontsize
        .Collapse wdCollapseEnd
        讀入網路資料後_於其後植入網址及設定格式 = rng.End 'Range或Selection值變動後，似乎用 With區塊無法及時回應，好像要用本尊「再次呼叫」才能反映及時真相 20241009
    End With
End Function
Rem 20241006 rng 要處理的範圍
Private Sub 讀入網路資料後_還原視窗狀態(win As word.Window, windowState As word.WdWindowState)
    
    With win.Application
        .Activate
        With win
            If .windowState = wdWindowStateMinimize Then
                VBA.Interaction.DoEvents
                .windowState = windowState
                .Activate
                VBA.Interaction.DoEvents
            End If
        End With
    End With

End Sub
Sub 送交古籍酷自動標點()
    'Alt + F10(此快速鍵待確認！）
    Dim ur As UndoRecord
    If Selection.Characters.Count < 10 Then
        MsgBox "字數太少，有必要嗎？請至少大於10字", vbExclamation
        Exit Sub
    End If
    Selection.Copy
    TextForCtext.GjcoolPunct
    Selection.Document.Activate
    Selection.Document.Application.Activate
    SystemSetup.stopUndo ur, "送交《古籍酷》自動標點"
    Selection.text = SystemSetup.GetClipboardText
    SystemSetup.contiUndo ur
End Sub
Sub 讀入古籍酷自動標點結果()
    'Ctrl + Alt + F10 或 Ctrl + Alt + F11
    If inputGjcoolPunctResult = False Then MsgBox "請重試！", vbCritical
End Sub
Rem 20241008 失敗則傳回false
Function inputGjcoolPunctResult() As Boolean
        Dim ur As UndoRecord, result As String
    文字處理.ResetSelectionAvoidSymbols
    If Selection.Characters.Count < 10 Then
        MsgBox "字數太少，有必要嗎？請至少大於10字", vbExclamation
        Exit Function
    End If
    word.Application.ScreenUpdating = False
    Const ignoreMarker = "《》〈〉「」『』" '書名號、篇名號、引號不處理（由前面的程式碼處理）
    result = Selection.text
    Rem 書名號、引號之處理
    result = VBA.Replace(VBA.Replace(result, "《", "〔"), "》", "〕") '書名號亦會被自動標點清除故,以備還原 20241001
    result = VBA.Replace(VBA.Replace(result, "「", "〔"), "」", "〕") '引號亦會被自動標點清除故,以備還原 20241001
    
    If SeleniumOP.grabGjCoolPunctResult(result, result, False) = vbNullString Then
        Selection.Document.Activate
        Selection.Document.Application.Activate
        Exit Function
    End If
    Selection.Document.Activate
    Selection.Document.Application.Activate
    Rem 書名號、引號之處理
    result = VBA.Replace(VBA.Replace(result, "〔", "《"), "〕", "》") '書名號亦會被自動標點清除故,以備還原 20241001
    result = VBA.Replace(VBA.Replace(result, "〔", "「"), "〕", "」") '引號亦會被自動標點清除故,以備還原 20241001
    result = VBA.Replace(result, VBA.Chr(13) & VBA.Chr(10), VBA.Chr(13)) '讀回來的自動標點結果會將chr(13)轉成VBA.Chr(13) & VBA.Chr(10)
    SystemSetup.stopUndo ur, "讀入《古籍酷》自動標點結果"
    Rem Selection.text = result'純文字處理
    Dim puncts As New punctuation, cln As New VBA.Collection, e, rng As Range '適應於格式化文字
    Set cln = puncts.CreateContextPunctuationCollection(result)
    Rem 清除原來的標點符號，以利比對與插入
    For Each e In Selection.Characters
        'If e = "。" Then Stop 'just for test
        If e.text = "　" Then '空格要清除（《古籍酷》自動標點會清除空格）
            e.text = vbNullString
        Else
            If VBA.InStr(ignoreMarker, e.text) = 0 Then '書名號、引號不處理（由前面的程式碼處理）
                If puncts.PunctuationDictionary.Exists(e.text) Then
                    e.text = vbNullString
                End If
            End If
        End If
    Next e
    Set rng = Selection.Document.Range(Selection.start, Selection.End)
    rng.Find.ClearAllFuzzyOptions: rng.Find.ClearFormatting
    For Each e In cln
'        If e(1) = Chr(13) Then Stop 'just for test
        If e(0) <> vbNullString Then
            If rng.Find.Execute(e(0), , , , , , True, wdFindStop) = False Then
                If rng.text = e(0) Then '最後一個
                    If VBA.InStr(ignoreMarker, e(1)) = 0 Then '書名號、引號不處理（由前面的程式碼處理）
                        rng.InsertAfter e(1)
                    End If
'                Else
'                    Stop 'just for test
                End If
            Else
                If VBA.InStr(ignoreMarker, e(1)) = 0 Then '書名號、引號不處理（由前面的程式碼處理）
                    rng.InsertAfter e(1)
                Else
                    rng.SetRange rng.start, rng.End + 1
                End If
            End If
        Else
            If VBA.InStr(ignoreMarker & VBA.Chr(13), e(1)) = 0 Then '書名號、引號不處理（由前面的程式碼處理）
                rng.Collapse wdCollapseStart
                rng.InsertAfter e(1)
            Else
                rng.SetRange rng.start, rng.start + 1
            End If
        End If
        If rng.End <= Selection.End Then '最後一個
            Set rng = Selection.Document.Range(rng.End, Selection.End)
        Else
            Selection.End = rng.End
        End If
    Next e
    word.Application.ScreenUpdating = True
    SystemSetup.contiUndo ur
    inputGjcoolPunctResult = True
End Function
Function GetUserAddress() As Boolean
    Dim x As String, a As Object 'Access.Application
    On Error GoTo Error_GetUserAddress
    x = Selection.text
    Set a = GetObject("D:\千慮一得齋\書籍資料\圖書管理.mdb") '2010/10/18修訂
    If x = "" Then x = InputBox("請輸入欲查詢的字串")
    x = a.Run("查詢字串轉換_國語會碼", x)
''    'ActiveDocument.FollowHyperlink "http://140.111.34.46/cgi-bin/dict/newsearch.cgi", , False, , "Database=dict&GraphicWord=yes&QueryString=^" & X & "$", msoMethodGet
'    FollowHyperlink "http://dict.revised.moe.edu.tw/cgi-bin/newDict/dict.sh?", , False, , "=dict.idx&cond=^" & x & "$&pieceLen=50&fld=1&cat=&imgFont=1", msoMethodGet
    Shell Replace(GetDefaultBrowserEXE, """%1", "http://dict.revised.moe.edu.tw/cgi-bin/newDict/dict.sh?cond=^" & x & "$&pieceLen=50&fld=1&cat=&imgFont=1")
    'AppActivate GetDefaultBrowser'無效
'    'FollowHyperlink "http://dict.revised.moe.edu.tw/cgi-bin/newDict/dict.sh?", , False, , "=dict.idx&cond=^" & X & "$&pieceLen=50&fld=1&cat=&imgFont=1", msoMethodGet
    
'    If Len(Selection.Text) = 1 Then _
        FollowHyperlink "http://www.nlcsearch.moe.gov.tw/EDMS/admin/dict3/search.php", , False, , "qstr=" & x & "&dictlist=47,46,51,18,16,13,20,19,53,12,14,17,48,57,24,25,26,29,30,31,32,33,34,35,36,37,39,38,41,42,43,45,50,&searchFlag=A&hdnCheckAll=checked", msoMethodGet '2009/1/10'教育部-國家語文綜合連結檢索系統-語文綜合檢索
        If a.Visible = False Then
            a.Visible = True
            a.UserControl = True
        End If
'        a.Quit acQuitSaveNone
'        Set a = Nothing
    GetUserAddress = True
Exit_GetUserAddress:
    Exit Function

Error_GetUserAddress:
    MsgBox Err & ": " & Err.Description
    GetUserAddress = False
    Resume Exit_GetUserAddress
End Function


    
Function GetDefaultBrowser() '2010/10/18由http://chijanzen.net/wp/?p=156#comment-1303(取得預設瀏覽器(default web browser)的名稱? chijanzen 雜貨舖)而來.
    Dim objShell
    Set objShell = CreateObject("WScript.Shell")
    'HKEY_CLASSES_ROOT\HTTP\shell\open\ddeexec\Application
    '取得註冊表中的值
    GetDefaultBrowser = objShell.RegRead _
            ("HKCR\http\shell\open\ddeexec\Application\")
    'GetDefaultBrowser = objShell.RegRead _
            ("HKEY_CLASSES_ROOT\http\shell\open\ddeexec\Application\")
End Function


Function GetDefaultBrowserEXE() '2010/10/18由http://chijanzen.net/wp/?p=156#comment-1303(取得預設瀏覽器(default web browser)的名稱? chijanzen 雜貨舖)而來.
Dim deflBrowser As String
deflBrowser = getDefaultBrowserNameAppActivate
Select Case deflBrowser
    Case "iexplore":
        GetDefaultBrowserEXE = "C:\Program Files\Internet Explorer\iexplore.exe"
    Case "firefox":
        If Dir("W:\PortableApps\PortableApps\FirefoxPortable\App\Firefox64\firefox.exe") = "" Then
            GetDefaultBrowserEXE = "C:\Program Files\Mozilla Firefox\firefox.exe"
        Else
            GetDefaultBrowserEXE = "W:\PortableApps\PortableApps\FirefoxPortable\App\Firefox64\firefox.exe"
        End If
    Case "brave":
        If Dir(UserProfilePath & "\AppData\Local\BraveSoftware\Brave-Browser\Application\brave.exe") = "" Then
            GetDefaultBrowserEXE = "C:\Program Files (x86)\BraveSoftware\Brave-Browser\Application\brave.exe"
        Else
            GetDefaultBrowserEXE = UserProfilePath & "\AppData\Local\BraveSoftware\Brave-Browser\Application\brave.exe"
        End If
    Case "vivaldi":
        GetDefaultBrowserEXE = UserProfilePath & "\AppData\Local\Vivaldi\Application\vivaldi.exe"
    Case "Opera":
        GetDefaultBrowserEXE = ""
    Case "Safari":
        GetDefaultBrowserEXE = ""
    Case "edge":
        GetDefaultBrowserEXE = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" '"msedge"
    Case "ChromeHTML", "google chrome": '"chrome"
        GetDefaultBrowserEXE = SystemSetup.getChrome
'
'        If Dir("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe") = "" Then
'            GetDefaultBrowserEXE = "W:\PortableApps\PortableApps\GoogleChromePortable\GoogleChromePortable.exe"
'        Else
'            GetDefaultBrowserEXE = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
'        End If
    Case Else:
        Dim objShell
        Set objShell = CreateObject("WScript.Shell")
        'HKEY_CLASSES_ROOT\HTTP\shell\open\ddeexec\Application
        '取得註冊表中的值
        deflBrowser = objShell.RegRead _
                ("HKCR\http\shell\open\command\")
        GetDefaultBrowserEXE = VBA.Mid(deflBrowser, 2, InStr(deflBrowser, ".exe") + Len(".exe") - 2)

End Select
    
    
End Function

Function getDefaultBrowserFullname()
Dim appFullname As String
appFullname = GetDefaultBrowserEXE
'appFullname = VBA.Mid(appFullname, 2, InStr(appFullname, ".exe") + Len(".exe") - 2)
getDefaultBrowserFullname = appFullname
'DefaultBrowserNameAppActivate = VBA.Replace(VBA.Mid(appFullname, InStrRev(appFullname, "\") + 1), ".exe", "")
End Function


Function getDefaultBrowserNameAppActivate() As String
Dim objShell, ProgID As String: Set objShell = CreateObject("WScript.Shell")
ProgID = objShell.RegRead _
            ("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice\ProgID")
ProgID = VBA.Mid(ProgID, 1, IIf(InStr(ProgID, ".") = 0, Len(ProgID), InStr(ProgID, ".") - 1))
Select Case ProgID
    Case "IE.HTTP":
        DefaultBrowserNameAppActivate = "iexplore"
    Case "FirefoxURL":
        DefaultBrowserNameAppActivate = "firefox"
    Case "ChromeHTML":
        DefaultBrowserNameAppActivate = "google chrome"
    Case "BraveHTML":
        DefaultBrowserNameAppActivate = "brave"
    Case "VivaldiHTM":
        DefaultBrowserNameAppActivate = "vivaldi"
    Case "OperaStable":
        DefaultBrowserNameAppActivate = "Opera"
    Case "SafariHTML":
        DefaultBrowserNameAppActivate = "Safari"
    Case "AppXq0fevzme2pys62n3e0fbqa7peapykr8v", "MSEdgeHTM":
        'browser = BrowserApplication.Edge;
        DefaultBrowserNameAppActivate = "edge" '"msedge"
    Case Else:
        DefaultBrowserNameAppActivate = "google chrome" '"chrome"
End Select
getDefaultBrowserNameAppActivate = DefaultBrowserNameAppActivate
End Function


Sub AppActivateDefaultBrowser()
    On Error GoTo eH
    Dim i As Byte, a
    a = Array("google chrome", "brave", "edge")
    DoEvents
    If DefaultBrowserNameAppActivate = "" Then getDefaultBrowserNameAppActivate
    AppActivate DefaultBrowserNameAppActivate
    DoEvents
    Exit Sub
eH:
        Select Case Err.Number
            Case 5
                DefaultBrowserNameAppActivate = a(i)
                i = i + 1
                If i > UBound(a) Then
                    MsgBox Err.Number & Err.Description
                    Exit Sub
                End If
                Resume
            Case Else
                MsgBox Err.Number + Err.Description
        End Select
    'AppActivate ""
End Sub



