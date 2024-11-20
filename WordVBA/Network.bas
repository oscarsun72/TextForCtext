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
    If Not code.IsChineseCharacter(Selection.text) Then
        MsgBox "限中文字", vbCritical
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
    If Not code.IsChineseCharacter(Selection.text) Then
        MsgBox "限中文字", vbCritical
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
    If Not code.IsChineseCharacter(Selection.text) Then
        MsgBox "限中文字", vbCritical
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
Sub 查韻典網()
    Rem Ctrl + Shift + Alt + y  y=yun（韻）的y
    文字處理.ResetSelectionAvoidSymbols
    SeleniumOP.LookupYtenx Selection.text
End Sub
Sub 查國學大師()
    Rem Alt + Shift + g ：檢索《國學大師》（g=guo 國學大師的國）；原為：Ctrl + d + s （ds：大師）
    文字處理.ResetSelectionAvoidSymbols
    SeleniumOP.LookupGXDS Selection.text
End Sub
Sub 查中文大辭典()
    '《國學大師》所收轉愚所掃全彩為黑白版 20241020
    Rem Alt + Shift + z （z：中（zh）的 z）
    文字處理.ResetSelectionAvoidSymbols
    SeleniumOP.LookupZWDCD Selection.text
End Sub
Sub 查古音小鏡_訓詁工具書查詢()
    Rem Ctrl + Shift + Alt + U u=xun（訓）的u
    文字處理.ResetSelectionAvoidSymbols
    SeleniumOP.LookupBook_Xungu_kaom Selection.text
End Sub
Sub 查古音小鏡_漢語大詞典()
    Rem Ctrl + Shift + Alt + i i=ci（詞）的i
    文字處理.ResetSelectionAvoidSymbols
    SeleniumOP.LookupHYDCD_kaom Selection.text
End Sub
Sub 查白雲深處人家漢語大詞典()
    Rem Ctrl + Shift + Alt + c c=ci（詞）的c
    Rem  Ctrl + Alt + b （b=bai(白，白雲深處人家的白)）
    文字處理.ResetSelectionAvoidSymbols
    LookupHomeinmistsHYDCD Selection.text
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
    If Not code.IsChineseCharacter(Selection.text) Then
        MsgBox "限中文字", vbCritical
        Exit Sub
    End If
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
    If Not code.IsChineseCharacter(Selection.text) Then
        MsgBox "限中文字", vbCritical
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
    If Not code.IsChineseCharacter(Selection.text) Then
        MsgBox "限中文字", vbCritical
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
                    innerHTML_Convert_to_WordDocumentContent rngHtml ', vbNullString
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
    SystemSetup.playSound 0.484
    Dim linkInput As Boolean, rngLink As Range, rngHtml As Range, ss As Long, x As String, gua As String, e, arrYi
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
                
        If Selection.Characters.Count > 3 Then
errExit:
            word.Application.Activate
            VBA.MsgBox "卦名有誤! 請重新選取。", vbExclamation
            Exit Sub
        End If
    End If
    
    Dim ur As UndoRecord, s As Long, ed As Long
    Dim fontsize As Single, rngBooks As Range
    fontsize = VBA.IIf(Selection.font.Size = 9999999, 12, Selection.font.Size) * 0.6
    If fontsize < 0 Then fontsize = 10

    SystemSetup.stopUndo ur, "查易學網易經周易原文指定卦名文本_並取回其純文字值及網址值插入至插入點位置"
    word.Application.ScreenUpdating = False
    Dim windowState As word.WdWindowState      '記下原來的視窗模式
    windowState = word.Application.windowState '記下原來的視窗模式
    If Selection.Document.path <> vbNullString And Selection.Document.Saved = False Then Selection.Document.Save
    
    s = Selection.start: ss = s
    
    Dim result(1) As String, iwe As SeleniumBasic.IWebElement

    If linkInput Then
        If SeleniumOP.OpenChrome(rngLink.Hyperlinks(1).Address) Then
            Set iwe = WD.FindElementByCssSelector("#block-bartik-content > div > article > div > div.clearfix.text-formatted.field.field--name-body.field--type-text-with-summary.field--label-hidden.field__item")
            If Not iwe Is Nothing Then
                Selection.MoveUntil Chr(13)
                Selection.TypeText Chr(13)
                Selection.Style = word.wdStyleNormal '"內文"
                Selection.ClearFormatting
                Selection.Collapse wdCollapseEnd
                s = Selection.start
                Rem 現在一律以html取內容，會更優（尤其在諸格式表現上，可以幾乎完全對應原網頁樣式） 20241015
'                If SeleniumOP.IslinkImageIncluded內容部分包含超連結或圖片(iwe) Then '有圖片時取 "innerHTML" 屬性值
'                    Dim Links() As SeleniumBasic.IWebElement, images() As SeleniumBasic.IWebElement
'                    Links = SeleniumOP.Links
'                    images = SeleniumOP.images
                    Selection.InsertAfter iwe.GetAttribute("innerHTML")
                    ed = Selection.End
                                        
                    Set rngHtml = Selection.Document.Range(s, ed)
                    
                    
                    
                    innerHTML_Convert_to_WordDocumentContent rngHtml, "https://www.eee-learning.com"
                    'SeleniumOP.inputElementContentAll插入網頁元件所有的內容 iwe
                    
                    
                    

'                    Stop 'just for test
                    GoTo finish 'just for test

                Rem 現在一律以html取內容，會更優（尤其在諸格式表現上，可以幾乎完全對應原網頁樣式） 20241015
'                Else '沒有圖片時取 "textContent" 屬性值
'                    result(0) = iwe.GetAttribute("textContent")
'                    result(1) = rngLink.Hyperlinks(1).Address
'                End If Rem 現在一律以html取內容，會更優（尤其在諸格式表現上，可以幾乎完全對應原網頁樣式） 20241015
                
                Rem 現在一律以html取內容，會更優（尤其在諸格式表現上，可以幾乎完全對應原網頁樣式） 20241015
'                GoTo insertText:
            Else 'If iwe Is Nothing Then
                MsgBox "網頁結構不同，沒有所需內容，可能是連結至各章節的超連結，【請貼上含有超連結的各章節標題或文字後再繼續】。感恩感恩　南無阿彌陀佛　讚美主", vbExclamation
                GoTo finish
            End If
        End If
    Else 'If Not linkInput Then
        Rem 自動檢查下一個字
            Selection.Collapse wdCollapseStart
'        If Selection.Characters.Count = 1 Then
            x = Selection.text
            If Not Keywords.周易卦名_卦形_卦序.Exists(x) And Not Keywords.易學異體字典.Exists(x) Then
                x = Selection.Characters(1).text & Selection.Characters(1).Next(wdCharacter, 1).text
                If Keywords.周易卦名_卦形_卦序.Exists(x) Or Keywords.易學異體字典.Exists(x) Then
                    Selection.MoveRight wdCharacter, 2, wdExtend
                Else
                    arrYi = Array("習坎", "繫辭", "說卦", "序卦", "雜卦")
                    If VBA.IsArray(VBA.Filter(arrYi, x)) Then
                        If (UBound(VBA.Filter(arrYi, x)) > -1) Then
                            Select Case x
                                Case "習坎"
                                    gua = "29"
                                    GoTo grab:
                                Case "繫辭"
                                    Select Case Selection.Next(wdCharacter, 2).text
                                        Case "上"
                                            gua = "65"
                                            GoTo grab:
                                        Case "下"
                                            gua = "66"
                                            GoTo grab:
                                    End Select
                                Case "說卦"
                                    gua = "67" '"說卦傳"
                                    GoTo grab:
                                Case "序卦"
                                    gua = "68" '"序卦傳"
                                    GoTo grab:
                                Case "雜卦"
                                    gua = "69" '"雜卦傳"
                                    GoTo grab:
                            End Select
                        End If
                    End If
                End If
            End If
'        End If
        gua = Selection.text
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
    
    gua = Keywords.周易卦名_卦形_卦序(gua)(1)

grab:
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
                    Selection.Collapse wdCollapseEnd
                End If
                
                s = Selection.start
                
                Selection.InsertAfter iwe.GetAttribute("innerHTML")
                ed = Selection.End
                
                
                
                Set rngHtml = Selection.Document.Range(s, ed)
                
                innerHTML_Convert_to_WordDocumentContent rngHtml, "https://www.eee-learning.com"
                'SeleniumOP.inputElementContentAll插入網頁元件所有的內容 iwe
                
                
'                Dim refs As Boolean
'                Rem 歷代注本： or 各家註解：
'                Set rngBooks = Selection.Document.Range(rngHtml.start, rngHtml.End)
'                rngBooks.Find.ClearFormatting
'                If rngBooks.Find.Execute("歷代注本：", , , , , , True, wdFindStop) Then
'                    refs = True
'                Else
'                    rngBooks.SetRange rngHtml.start, rngHtml.End
'                    If rngBooks.Find.Execute("各家註解：", , , , , , True, wdFindStop) Then
'                        refs = True
'                    End If
'                End If
'                If refs Then
'                    With rngBooks '"歷代注本："所在段落範圍
'                        .Style = wdStyleHeading1
'                        .font.Size = 22
'                        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle '單行間距
'                    End With
'                    ed = 讀入網路資料後_於其後植入網址及設定格式(rngHtml, result(1), fontsize)
'                    讀入網路資料後_還原視窗狀態 Selection.Document.ActiveWindow, windowState
'                End If

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
        
    Dim p As Paragraph, book As String, iwes() As IWebElement
    
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
    If linkInput And Not rngHtml Is Nothing Then
        Dim refs As Boolean
        Rem 歷代注本： or 各家註解：
        Set rngBooks = Selection.Document.Range(rngHtml.start, rngHtml.End)
        rngBooks.Find.ClearFormatting
        If rngBooks.Find.Execute("歷代注本：", , , , , , True, wdFindStop) Then
            refs = True
        Else
            rngBooks.SetRange rngHtml.start, rngHtml.End
            If rngBooks.Find.Execute("各家註解：", , , , , , True, wdFindStop) Then
                refs = True
            End If
        End If
        If refs Then
            With rngBooks '"歷代注本："所在段落範圍
                .Style = wdStyleHeading1
                .font.Size = 22
                .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle '單行間距
            End With
            rngBooks.SetRange rngBooks.Paragraphs(1).Range.End + 1, rngHtml.End '設定為全部歷代注本的範圍
            For Each p In rngBooks.Paragraphs
                If p.Range.text <> Chr(13) And VBA.Left(p.Range.text, 4) <> "http" And VBA.Left(p.Range.text, 5) <> "歷代注本：" Then
                    If p.Range.Hyperlinks.Count = 0 Then
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
                    Else
                        If p.Range.text <> Chr(13) And p.Style <> "標題 2" And p.Style <> wdStyleHeading2 Or p.Style = "清單段落" Or p.Style = wdStyleListParagraph Then
                            p.Style = wdStyleHeading2 '標題 2
                            p.Range.font.Size = 18
                            p.Range.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle '單行間距
                        End If
                    End If
                End If
            Next p
            
        End If
        If result(1) <> vbNullString Then
            ed = 讀入網路資料後_於其後植入網址及設定格式(rngHtml, result(1), fontsize)
        End If
        讀入網路資料後_還原視窗狀態 Selection.Document.ActiveWindow, windowState
        
        Rem just for check
        With rngHtml.Find
            .ClearFormatting
            .MatchWholeWord = True
            If .Execute("[<>&;]", , , True) Then
                rngHtml.Select
                SystemSetup.playSound 3
            End If
            .MatchWholeWord = False
            .ClearFormatting
        End With
        Rem just for check
    Else 'If linkInput=false Then
        If result(1) <> vbNullString Then
            ed = 讀入網路資料後_於其後植入網址及設定格式(rngHtml, result(1), fontsize)
        End If
    End If 'If linkInput Then

    word.Application.ScreenUpdating = True
    Selection.Document.Range(ss, ss).Select
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
    Dim xmlhttp As Object
    Dim stream As Object
    
    ' 創建 XMLHTTP 對象
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    xmlhttp.Open "GET", url, False
    xmlhttp.send '-2147467259 無法指出的錯誤。可能是由於您使用的是 base64 編碼的 URL。XMLHTTP 無法直接處理 base64 編碼的圖像數據 https://sl.bing.net/dd1AOLdKBaK
    
    ' 創建 ADODB.Stream 對象
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    stream.Write xmlhttp.responseBody
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
    Dim xmlhttp As Object
    Dim stream As Object
    
    ' 創建 ServerXMLHTTP 對象
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
    xmlhttp.Open "GET", url, False
    xmlhttp.send
    
    ' 創建 ADODB.Stream 對象
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    stream.Write xmlhttp.responseBody
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
Function InsertBase64Image(base64String As String, filePath As String, rng As Range) As InlineShape
    Dim binaryData() As Byte
    Dim tempFilePath As String
    
    ' 解析base64編碼
    binaryData = Base64ToBinary(base64String)
    
    ' 保存為臨時文件
    tempFilePath = Environ("TEMP") & "\" & filePath
    SaveBinaryAsFile binaryData, tempFilePath
    
    ' 插入圖片
    Set InsertBase64Image = rng.InlineShapes.AddPicture(fileName:=tempFilePath, LinkToFile:=False, SaveWithDocument:=True, Range:=rng)
    base64String = tempFilePath
    ' 刪除臨時文件
    Kill tempFilePath
End Function


Function GetDomainUrlPrefix(url As String)
    GetDomainUrlPrefix = VBA.Left(url, VBA.InStr(url, "//")) & "/" & VBA.Mid(url, VBA.InStr(url, "//") + 2, _
                VBA.InStr(VBA.InStr(url, "//") + 2, url, "/") - (VBA.InStr(url, "//") + 2))
End Function
Rem 20241006 《看典古籍·古籍全文檢索》 Ctrl + Alt + j （j=籍 ji） 或 Alt + d （d：典 dian）
'原為 Ctrl + k,d  ，因會使內建的 Ctrl + k （插入超連結）失效，故改定 20241014
Sub 查看典古籍古籍全文檢索()
    文字處理.ResetSelectionAvoidSymbols
    SeleniumOP.KandiangujiSearchAll Selection.text
End Sub
Rem 20241006 檢索《漢籍全文資料庫》 Alt + Shfit + h 或 Alt + h
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
Rem Alt + F10(此快速鍵待確認！）
Sub 送交古籍酷自動標點()
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
Rem Ctrl + Alt + a
Sub 讀入AI太炎標點結果()
    playSound 0.484
    If inputAITShenShenWikiPunctResult = False Then MsgBox "請重試！", vbCritical
End Sub
Rem Ctrl + Alt + F10 或 Ctrl + Alt + F11
Sub 讀入古籍酷自動標點結果()
    playSound 0.484
    If inputGjcoolPunctResult = False Then MsgBox "請重試！", vbCritical
End Sub
Rem 20241008 失敗則傳回false
Function inputGjcoolPunctResult() As Boolean
    Dim ur As UndoRecord, result As String, d As Document
    文字處理.ResetSelectionAvoidSymbols
    If Selection.Characters.Count < 10 Then
        MsgBox "字數太少，有必要嗎？請至少大於10字", vbExclamation
        Exit Function
    End If
    word.Application.ScreenUpdating = False
    Set d = Selection.Document
    If d.path <> vbNullString And Not d.Saved Then d.Save
    Const ignoreMarker = "《》〈〉「」『』" '書名號、篇名號、引號不處理（由前面的程式碼處理）
    result = Selection.text
    Rem 書名號、引號之處理
    result = VBA.Replace(VBA.Replace(result, "《", "〔"), "》", "〕") '書名號亦會被自動標點清除故,以備還原 20241001
    result = VBA.Replace(VBA.Replace(result, "「", "〔"), "」", "〕") '引號亦會被自動標點清除故,以備還原 20241001
    
    Rem 送去《古籍酷》自動標點

    If SeleniumOP.grabGjCoolPunctResult(result, result, False) = vbNullString Then
        d.Activate
        d.Application.Activate
        Exit Function
    End If
    d.Activate
    d.Application.Activate
    
    GoSub 標點校正
    
    
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
    Set rng = d.Range(Selection.start, Selection.End)
    rng.Find.ClearFormatting
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
            Set rng = d.Range(rng.End, Selection.End)
        Else
            Selection.End = rng.End
        End If
    Next e
    word.Application.ScreenUpdating = True
    SystemSetup.contiUndo ur
    inputGjcoolPunctResult = True
    Exit Function
    
標點校正:
    Dim arrPunct, arrPunctCorrector, iarr As Byte, arrUb As Byte
    arrPunct = Array(VBA.ChrW(-10148) & VBA.ChrW(-9010) & "：文" & VBA.ChrW(28152) & "公")
    arrPunctCorrector = Array(VBA.ChrW(-10148) & VBA.ChrW(-9010) & "文" & VBA.ChrW(28152) & "公")
    arrUb = UBound(arrPunct)
    For iarr = 0 To arrUb
        If VBA.InStr(result, arrPunct(iarr)) Then
            result = VBA.Replace(result, arrPunct(iarr), arrPunctCorrector(iarr))
        End If
    Next iarr
    
    Return
    
End Function
Rem 20241008 失敗則傳回false
Function inputAITShenShenWikiPunctResult() As Boolean
    Dim ur As UndoRecord, result As String, d As Document
    文字處理.ResetSelectionAvoidSymbols
    If Selection.Characters.Count < 10 Then
        MsgBox "字數太少，有必要嗎？請至少大於10字", vbExclamation
        Exit Function
    End If
    word.Application.ScreenUpdating = False
    Set d = Selection.Document
    If d.path <> vbNullString And Not d.Saved Then d.Save
'    Const ignoreMarker = "《》〈〉「」『』" '書名號、篇名號、引號不處理（由前面的程式碼處理）
    Const ignoreMarker = "《》〈〉（）·「」『』" '書名號、篇名號、括號、音節號、引號不處理（由前面的程式碼處理）
    result = Selection.text
    Rem 括號、篇名號之處理：〔、〕、〔、〕 會被清除掉
    result = VBA.Replace(VBA.Replace(result, "《", "＜"), "》", "＞") '書名號亦會被自動標點清除故,以備還原 20241106
    result = VBA.Replace(VBA.Replace(result, "（", "【"), "）", "】") '括號亦會被自動標點清除故,以備還原 202411106
    result = VBA.Replace(VBA.Replace(result, "〈", VBA.ChrW(12310)), "〉", VBA.ChrW(12311))     '篇名號亦會被自動標點清除故,以備還原 20241001
    result = VBA.Replace(result, "·", "⊙")     '音節號亦會被自動標點清除故,以備還原 20241001
'
    If SeleniumOP.grabAITShenShenWikiPunctResult(result, result, False) = vbNullString Then
        d.Activate
        d.Application.Activate
        Exit Function
    Else
        result = VBA.Replace(VBA.Replace(VBA.Replace(VBA.Replace(result, VBA.ChrW(8220), "「"), VBA.ChrW(8221), "」"), VBA.ChrW(8216), "『"), VBA.ChrW(8217), "』")
        result = VBA.Replace(VBA.Replace(result, VBA.ChrW(12310) & "《", VBA.ChrW(12310)), "》" & VBA.ChrW(12311), VBA.ChrW(12311))      '書名號亦會被自動標點清除故,以備還原 20241106
        result = VBA.Replace(VBA.Replace(result, "《", "＜"), "》", "＞") '書名號亦會被自動標點清除故,以備還原 20241106
        result = VBA.Replace(result, "·", "⊙")     '音節號亦會被自動標點清除故,以備還原 20241001
    End If
    
    d.Application.Activate
    d.Activate
    Rem 括號之處理'標點會在（處停止
    result = VBA.Replace(VBA.Replace(result, "＜", "《"), "＞", "》") '書名號亦會被自動標點清除故,以備還原 20241001
    result = VBA.Replace(result, "⊙", "·")  '音節號亦會被自動標點清除故,以備還原 20241001
    result = VBA.Replace(VBA.Replace(result, "【", "（"), "】", "）")  '括號亦會被自動標點清除故,以備還原 20241001
    result = VBA.Replace(VBA.Replace(result, VBA.ChrW(12310), "〈"), VBA.ChrW(12311), "〉")     '篇名號亦會被自動標點清除故,以備還原 20241001
    'result = VBA.Replace(result, VBA.Chr(13) & VBA.Chr(10), VBA.Chr(13)) '讀回來的自動標點結果會將chr(13)轉成VBA.Chr(13) & VBA.Chr(10)
    
    result = VBA.Replace(result, VBA.ChrW(12295), "○")  '"○"會被置換！
    Debug.Print result
    SystemSetup.stopUndo ur, "讀入AI太炎標點結果"
    Rem Selection.text = result'純文字處理
    Dim puncts As New punctuation, cln As New VBA.Collection, e, rng As Range '適應於格式化文字
    Set cln = puncts.CreateContextPunctuationCollection(result)
    Rem 清除原來的標點符號，以利比對與插入
    Set rng = d.Range(Selection.start, Selection.End)
    For Each e In rng.Characters
'        'If e = "。" Then Stop 'just for test
'        If e.text = "　" Then '空格要清除（《古籍酷》自動標點會清除空格）
'            e.text = vbNullString
'        Else
            If VBA.InStr(ignoreMarker, e.text) = 0 Then '書名號、引號不處理（由前面的程式碼處理）
                If puncts.PunctuationDictionary.Exists(e.text) Then
                    e.text = vbNullString
                End If
            End If
'        End If
    Next e
    Set rng = d.Range(Selection.start, Selection.End)
    rng.Find.ClearFormatting
    For Each e In cln
'        If e(1) = Chr(13) Then Stop 'just for test
        If e(0) <> vbNullString Then
            If rng.Find.Execute(e(0), , , , , , True, wdFindStop) = False Then
                If rng.text = e(0) Then '最後一個
                    If VBA.InStr(ignoreMarker, e(1)) = 0 Then '篇名號、括號不處理（由前面的程式碼處理）
                        rng.InsertAfter e(1)
                    End If
'                Else
'                    Stop 'just for test
                End If
            Else
                If VBA.InStr(ignoreMarker, e(1)) = 0 Then '括號、篇名號不處理（由前面的程式碼處理）
                    rng.InsertAfter e(1)
                Else
                    If rng.Document.Range(rng.End, rng.End + 1) <> e(1) Then
                        rng.Collapse wdCollapseEnd
                        rng.InsertAfter e(1)
                    Else
                        rng.SetRange rng.start, rng.End + 1
                    End If
                End If
            End If
        Else
            If VBA.InStr(ignoreMarker & VBA.Chr(13), e(1)) = 0 Then '篇名號、括號不處理（由前面的程式碼處理）
                rng.Collapse wdCollapseStart
                rng.InsertAfter e(1)
            Else
                If rng.Document.Range(rng.start, rng.start + 1) <> e(1) Then
                    rng.Collapse wdCollapseStart
                    rng.InsertAfter e(1)
                Else
                    rng.SetRange rng.start, rng.start + 1
                End If
            End If
        End If
        If rng.End <= Selection.End Then '最後一個
'            If rng.End = d.Range.End Then
'                Set rng = d.Range(rng.End - 1, Selection.End)
'            Else
                Set rng = d.Range(rng.End, Selection.End)
'            End If
        Else
            Selection.End = rng.End
        End If
    Next e
    SystemSetup.contiUndo ur
    word.Application.ScreenUpdating = True
    inputAITShenShenWikiPunctResult = True
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

Rem AppActivate方法常會當掉！！20241020
Sub AppActivateDefaultBrowser()
    On Error GoTo eH
    Dim i As Byte, a
    a = Array("google chrome", "brave", "edge")
    DoEvents
    If DefaultBrowserNameAppActivate = "" Then getDefaultBrowserNameAppActivate
    VBA.Interaction.AppActivate DefaultBrowserNameAppActivate
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



