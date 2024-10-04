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
            .font.Size = fontsize
            .InsertAfter ar(1) '插入網址
            SystemSetup.contiUndo ur
            .Collapse wdCollapseStart
            With .Application
                .Activate
                With .ActiveWindow
                    If .windowState = wdWindowStateMinimize Then
                        VBA.Interaction.DoEvents
                        .windowState = windowState
                        .Activate
                        VBA.Interaction.DoEvents
                    End If
                End With
            End With
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
        Dim ur As UndoRecord, fontsize As Single
        SystemSetup.stopUndo ur, "查異體字字典並取回其說文釋形欄位及網址值插入至插入點位置"
        With Selection
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
            Dim shuoWen As String
            shuoWen = VBA.Replace(VBA.Replace(ar(0), "：，", "：" & x & "，"), "段注本：", VBA.Chr(13) & "段注本：")
            If VBA.Left(shuoWen, 1) = "，" Then
                shuoWen = x & shuoWen
            End If
            If s = 0 And ar(0) <> "說文釋形沒有資料！" Then
                .InsertAfter shuoWen & VBA.Chr(13)  'ar(0)=《說文》內容
                .Collapse wdCollapseEnd
            End If
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
                        VBA.Interaction.DoEvents
                        .windowState = windowState
                        .Activate
                        VBA.Interaction.DoEvents
                    End If
                End With
            End With
        End With
    End If
End Sub
Rem 指定卦名再操作 20241004 Alt + Shift + y (y:易)
Sub 查易學網易經周易原文指定卦名文本_並取回其純文字值及網址值插入至插入點位置()
    文字處理.ResetSelectionAvoidSymbols
    Dim gua As String
    If Selection.Characters.Count > 2 Then
errExit:
        word.Application.Activate
        VBA.MsgBox "卦名有誤! 請重新選取。", vbExclamation
        Exit Sub
    End If
    
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
    
    gua = Keywords.周易卦名_卦形_卦序(gua)(1)

    Dim result As String
    If SeleniumOP.grabEeeLearning_IChing_ZhouYi_originalText(gua, result) = False Then
        word.Application.Activate
        VBA.MsgBox "找不到，或網頁改了或掛了……", vbInformation
        Exit Sub
    End If
    word.Application.Activate
    If VBA.InStr(result, "歷代注本：") Then
        If VBA.vbOK = MsgBox("是否清除「歷代注本：」以後的文字？", VBA.vbQuestion + VBA.vbOKCancel) Then
            result = VBA.Left(result, VBA.InStr(result, "歷代注本：") - 1)
        End If
    End If
    Dim ur As UndoRecord, s As Long
    SystemSetup.stopUndo ur, "查易學網易經周易原文指定卦名文本_並取回其純文字值及網址值插入至插入點位置"
    With Selection
        If .Type = wdSelectionIP Then .Delete
        s = .start
        .TypeText result
                
        文字處理.FixFontname .Document.Range(s, Selection.End)
        
        .Document.Range(s, s).Select
    End With
    SystemSetup.contiUndo ur
    Exit Sub
eH:
    Select Case Err.Number
        Case Else
            Debug.Print Err.Number & Err.Description
            Stop 'just for test
            Resume
    End Select
End Sub
Sub 送交古籍酷自動標點()
    'Alt + F10
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
    'Ctrl + Alt + F10
    Dim ur As UndoRecord, result As String
    文字處理.ResetSelectionAvoidSymbols
    If Selection.Characters.Count < 10 Then
        MsgBox "字數太少，有必要嗎？請至少大於10字", vbExclamation
        Exit Sub
    End If
    Const ignoreMarker = "《》〈〉「」『』" '書名號、篇名號、引號不處理（由前面的程式碼處理）
    result = Selection.text
    Rem 書名號、引號之處理
    result = VBA.Replace(VBA.Replace(result, "《", "〔"), "》", "〕") '書名號亦會被自動標點清除故,以備還原 20241001
    result = VBA.Replace(VBA.Replace(result, "「", "〔"), "」", "〕") '引號亦會被自動標點清除故,以備還原 20241001
    
    SeleniumOP.grabGjCoolPunctResult result, result, False
    
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
        If VBA.InStr(ignoreMarker, e.text) = 0 Then '書名號、引號不處理（由前面的程式碼處理）
            If puncts.PunctuationDictionary.Exists(e.text) Then
                e.text = vbNullString
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
    SystemSetup.contiUndo ur
End Sub
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



