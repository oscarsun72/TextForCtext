Attribute VB_Name = "NewMacros"
Option Compare Text
Option Explicit
Public DonotSave As Boolean
'快速鍵筆記：Alt+P =Shift+F5'2003/4/1
'            原ANSI65294（．）指定為Alt+.,今改為ANSI8231（‧）

Sub HideWebBar()
If CommandBars("web").Visible Then CommandBars("web").Visible = False
End Sub

Sub CheckSaved()
Dim bkIdx As Integer
With ActiveDocument '是舊檔案（文件）才檢查
'    If Not Dir(.FullName) = "" Then '有路徑（已儲存之舊檔），方儲存2003/3/27
    If Not .path = "" Then '判斷是不是新文件的辦法有此二種
        If .Saved = False Then
'            If .Bookmarks.Count > 1 Then '有書籤時才做
'                For Each bk In .Bookmarks                '將兩天前的標地書籤刪除'2003/3/28
'                    bkIdx = bkIdx + 1 '記下書籤索引
'                    With bk '如果是編輯處才處理
'                        If InStr(1, .Name, "Edit_", vbTextCompare) > 0 Then
'                            '如果是兩天前的
'                            Do While InStr(1, bk, "Edit_", vbTextCompare) > 0 And _
'                                    CDate(Replace(Mid(bk, 6, 10), "_", "/")) <= Date - 2
'                                bk.Delete '殺掉以後索引值會向前遞補
'                                Set bk = ActiveDocument.Bookmarks(bkIdx)
'                            Loop
'                            Exit For
'    '                    If InStr(.Name, "Edit_" & Format(Now() - 2, "yyyy_mm_dd")) Then
'    '                        .Delete
'                        End If
'                    End With
'                Next bk
'            End If
            If ActiveDocument.Application.Templates(1).Name = "論文.dot" Then '論文範本才執行記錄編輯處標籤2004/2/7
                For bkIdx = 1 To .bookmarks.Count
                    With .bookmarks(bkIdx)
                        If .End >= Selection.Range.End _
                            And .start <= Selection.Range.start _
                            And InStr(.Name, Format(Date, "yyyy_mm_dd")) Then  '不是今天建立的
                            bkIdx = 0
                            Exit For
                        End If
                    End With
                Next bkIdx
                If bkIdx <> 0 Then '避免同一區塊設定太多書籤
                    With .bookmarks '新增今天的標地書籤'2003/3/28
                        .DefaultSorting = wdSortByName
                        '12小時制：
        '                .Add Range:=Selection.Range, Name:="Edit_" & Format(Now(), "yyyy_mm_dd_AM/PM_hh_nn_ss")
                        '24小時制：
                        .Add Range:=Selection.Range, Name:="Edit_" & _
                                Format(Format(Date, "short date"), "yyyy_mm_dd__") _
                                    & Format(Format(Time, "Short Time"), "__hh_mm_dd")
                        .ShowHidden = False
                    End With
                End If
            End If
            .Save
            .UndoClear '清除還原清單也許如此可省記憶體
        End If
    End If
End With
End Sub

Sub CheckSavedNoClear() '2003/4/3
Dim bkIdx As Integer
With ActiveDocument '是舊檔案（文件）才檢查
'    If Not Dir(.FullName) = "" Then '有路徑（已儲存之舊檔），方儲存2003/3/27
    If Not .path = "" Then '判斷是不是新文件的辦法有此二種
        If .Saved = False Then
            For bkIdx = 1 To .bookmarks.Count
                With .bookmarks(bkIdx)
                    If .End >= Selection.Range.End _
                        And .start <= Selection.Range.start _
                        And InStr(.Name, Format(Date, "yyyy_mm_dd")) Then  '不是今天建立的
                        bkIdx = 0
                        Exit For
                    End If
                End With
            Next bkIdx
            If bkIdx <> 0 Then '避免同一區塊設定太多書籤
                With .bookmarks '新增今天的標地書籤'2003/3/28
                    .DefaultSorting = wdSortByName
                    '12小時制：
    '                .Add Range:=Selection.Range, Name:="Edit_" & Format(Now(), "yyyy_mm_dd_AM/PM_hh_nn_ss")
                    '24小時制：
                    .Add Range:=Selection.Range, Name:="Edit_" & _
                            Format(Format(Date, "short date"), "yyyy_mm_dd__") _
                                & Format(Format(Time, "Short Time"), "__hh_mm_dd")
                    .ShowHidden = False
                End With
            End If
            .Save
            'NoClear
'            .UndoClear '清除還原清單也許如此可省記憶體
        End If
    End If
End With
End Sub


Sub ClearTodayBookmarks() '將今天的標地書籤刪除'2003/3/28
Dim bk As Bookmark, bkIdx As Integer
With ActiveDocument '是舊檔案（文件）才處理
'    If Not Dir(.FullName) = "" Then '有路徑（已儲存之舊檔），方處理
    If Not .path = "" Then '判斷是不是新文件的辦法有此二種
        If .bookmarks.Count > 1 Then '有書籤時才做
            For Each bk In .bookmarks
                bkIdx = bkIdx + 1 '記下書籤索引
                With bk '如果是編輯處才處理
                    If InStr(1, .Name, "Edit_", vbTextCompare) > 0 Then
                        '刪除今天之標地書籤
                        Do While InStr(1, bk, Format(Date, "yyyy_mm_dd"), vbTextCompare)
                            bk.Delete '殺掉以後索引值會向前遞補
                            Set bk = ActiveDocument.bookmarks(bkIdx)
                        Loop
'                        Exit For
                    End If
                End With
            Next bk
        End If
    End If
End With
End Sub

Sub DeleteSelBookmarks()
'指定鍵:Alt+Del
With Selection
    On Error GoTo 5941
    If MsgBox(.bookmarks.item(1) & "書籤，確定刪除？" & _
            vbCr & vbCr & "其內容為：" & .bookmarks(1).Range _
            , vbExclamation + vbOKCancel, "BookmarkID = " & .BookmarkID) = vbOK Then
        .bookmarks.item(1).Delete
    End If
Exit Sub
5941 '集合中的成員不存在！
    Select Case Err.Number
        Case 5941
            MsgBox "此處沒有書籤！", vbExclamation
        Case Else
            MsgBox Err.Number & Err.Description
    End Select
End With
End Sub

Sub BopomofoOnlyDirect()
'
' BopomofoOnlyDirect 巨集
' 巨集建立於 2001/11/12，建立者 孫守真
'

End Sub
Sub 離開() '不存檔,也不記錄所在位置
'Sub 開始編輯OLE物件()
'
' 開始編輯OLE物件 巨集
' 巨集錄製於 2001/11/12，錄製者 孫守真
'指定鍵:Ctrl+Alt+Q
Dim i As Byte
    On Error Resume Next
    For i = 1 To Documents.Count
        QuitClose
    Next i
    ActiveWindow.Close wdDoNotSaveChanges
    word.Application.Quit wdDoNotSaveChanges '2003/3/22
'    Selection.WholeStory
'    Selection.Delete Unit:=wdCharacter, Count:=1
End Sub
Sub OLE至備忘欄()
Attribute OLE至備忘欄.VB_Description = "巨集錄製於 2001/11/13，錄製者 孫守真"
Attribute OLE至備忘欄.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.OLE至備忘欄"
'
' OLE至備忘欄 巨集
' 巨集錄製於 2001/11/13，錄製者 孫守真
'快速鍵:alt+w--今改為開新視窗指定鍵
Dim ActCtl As Control, s As Integer, l As Integer
With ActiveDocument '2003/3/27
    If .path <> "" Then MsgBox "此文件不能操作", vbExclamation: Exit Sub
'    With .Content '全選
'    '     .Selection.WholeStory
'        'Selection.Copy
'        .Cut '剪下
'    '    Selection.Cut
'    End With
    On Error GoTo ErrH:
    AppActivate "圖書管理", True
    blog.myaccess.screen.ActiveControl.SetFocus
    Set ActCtl = blog.myaccess.screen.ActiveControl
    Dim cName As String
    Select Case ActCtl.Parent.Name
        Case "札", "札記_藏", "札_查詢", "札記_類查詢"
            cName = "〈" & blog.myaccess.DLookup("篇名", "篇", "篇ID = " & _
                ActCtl.Parent.Recordset("篇ID")) & "〉"
        Case "劄記"
            cName = "〈" & ActCtl.Parent.Recordset("篇名") & "〉"
        Case Else
            cName = "【無篇名!】"
    End Select
    If MsgBox("現在作用中的控制項是--〔" & ActCtl.Name & "〕！" & vbCr & vbCr _
        & "篇名是： " & cName & vbCr & vbCr & _
            "要由系統自動尋找,請按〔取消〕", vbOKCancel + vbExclamation _
            , "作用中的表單是: " & ActCtl.Parent.Name) = vbCancel Then
'        Stop
        For Each ActCtl In blog.myaccess.screen.activeform.Controls '2003/3/２7
        If TypeName(ActCtl) = "textbox" Then
'            If ActCtl.ControlSource = "札記" Then
'                 '因為札記_GotFocus有程式碼，不宜用SetFocus!
                If MsgBox("現在作用中的控制項是--〔" & ActCtl.Name & "〕！" & vbCr & vbCr _
                        & "要繼續看下一個控制項請按〔取消〕", vbOKCancel + vbExclamation) = vbOK Then
                    With .Content '全選
                        .Cut '剪下
                    End With
                    With .Application.CommandBars("論文＿札記瀏覽")
                        If .Visible = True Then .Visible = False
                    End With
                    .ActiveWindow.Close wdDoNotSaveChanges
                    ActCtl.SetFocus '到這裡再SetFocus！
                    Exit For
                End If
'            End If
        End If
        Next ActCtl
        If ActCtl Is Nothing Then
            MsgBox "沒有適合的控制項，請自行決定！", vbExclamation
            End ' AppActivate .Application.Name
        End If
    Else
        With .ActiveWindow.Selection
            s = .start '記下插入點位置'2003/3/30
            l = Len(.text)
        End With
        With .Content '全選
            .Cut '剪下
        End With
        With .Application.CommandBars("論文＿札記瀏覽")
            If .Visible = True Then .Visible = False
        End With
        .ActiveWindow.Close wdDoNotSaveChanges
    End If
End With
'    Application.Visible = False
    AppActivate "圖書管理"
'    Access.Application.SetOption "Behavior Entering Field", 0
    With ActCtl
        If .Parent.DefaultView = 0 Then 'Single Form
            .Parent.AllowEdits = True '設定此時,若非單一表單檢視則會將記錄移動至第一筆!2002/11/28
        End If
    '    Else
    '        Screen.ActiveForm.ActiveControl.form.ActiveControl.SetFocus
    ''        DoCmd.GoToRecord Screen.ActiveForm.CurrentRecord
'        If .Name = "札記" Then
'            .Locked = False
'        End If
        If .Locked = True Then .Locked = False
        If .Parent.AllowEdits = False Then .Parent.AllowEdits = True
        If .Name <> "札記" Then
            .SetFocus
            .SelStart = 0 '全選
            .SelLength = Len(.text)
        Else
             .Value = Null
        End If
        blog.myaccess.docmd.RunCommand blog.myaccess.acCmdPaste
        .SelStart = s 's + 1 '設定插入點位置
        .SelLength = l
    End With
    If Windows.Count = 0 And Documents.Count = 0 Then word.Application.Quit ' wdDotNotSaveChanges    '如果沒有視窗和文件開啟,才關掉2003/3/27
Exit Sub
ErrH:
Select Case Err.Number
    Case 5 '程序呼叫或引數不正確(即 AppActivate的引數有誤!)
        On Error GoTo ErrH1
'        AppActivate "圖書管理 - [" & Screen.ActiveControl.Parent.Caption & "]"
'        DoCmd.Restore
        AppActivate blog.myaccess.CurrentObjectName
        Resume Next
    Case Else
Shows:  MsgBox Err.Number & Err.Description
End Select
Exit Sub
ErrH1:
Select Case Err.Number
    Case 5 '程序呼叫或引數不正確(即 AppActivate的引數有誤!)
        AppActivate "圖書管理 - [" & blog.myaccess.screen.ActiveControl.Parent.Caption & "]"
        Resume Next
    Case Else
        GoTo Shows
End Select
End Sub
Sub OLE至備忘欄1() '複製到圖書管理而待繼續編輯，不予關閉者2003/12/25
'快速鍵:alt+w
Dim ActCtl As Control, s As Integer, l As Integer
With ActiveDocument '2003/3/27
    If .path <> "" Then MsgBox "此文件不能操作", vbExclamation: Exit Sub
'    With .Content '全選
'    '     .Selection.WholeStory
'        'Selection.Copy
'        .Cut '剪下
'    '    Selection.Cut
'    End With
    AppActivate "圖書管理", True
    blog.myaccess.screen.ActiveControl.SetFocus
    Set ActCtl = blog.myaccess.screen.ActiveControl
    Dim cName As String
    Select Case ActCtl.Parent.Name
        Case "札", "札記_藏", "札_查詢", "札記_類查詢"
            cName = "〈" & blog.myaccess.DLookup("篇名", "篇", "篇ID = " & _
                ActCtl.Parent.Recordset("篇ID")) & "〉"
        Case "劄記"
            cName = "〈" & ActCtl.Parent.Recordset("篇名") & "〉"
        Case Else
            cName = "【無篇名!】"
    End Select
    If MsgBox("現在作用中的控制項是--〔" & ActCtl.Name & "〕！" & vbCr & vbCr _
        & "篇名是： " & cName & vbCr & vbCr & _
            "要由系統自動尋找,請按〔取消〕", vbOKCancel + vbExclamation _
            , "作用中的表單是: " & ActCtl.Parent.Name) = vbCancel Then
'        Stop
        For Each ActCtl In blog.myaccess.screen.activeform.Controls '2003/3/２7
        If TypeName(ActCtl) = "textbox" Then
'            If ActCtl.ControlSource = "札記" Then
'                 '因為札記_GotFocus有程式碼，不宜用SetFocus!
                If MsgBox("現在作用中的控制項是--〔" & ActCtl.Name & "〕！" & vbCr & vbCr _
                        & "要繼續看下一個控制項請按〔取消〕", vbOKCancel + vbExclamation) = vbOK Then
                    With .Content '全選
                        .Cut '剪下
                    End With
                    With .Application.CommandBars("論文＿札記瀏覽")
                        If .Visible = True Then .Visible = False
                    End With
                    .ActiveWindow.Close wdDoNotSaveChanges
                    ActCtl.SetFocus '到這裡再SetFocus！
                    Exit For
                End If
'            End If
        End If
        Next ActCtl
        If ActCtl Is Nothing Then
            MsgBox "沒有適合的控制項，請自行決定！", vbExclamation
            End ' AppActivate .Application.Name
        End If
    Else
        With .ActiveWindow.Selection
            s = .start '記下插入點位置'2003/3/30
            l = Len(.text)
        End With
        With .Content '全選
            .Copy '複製
        End With
        With .Application.CommandBars("論文＿札記瀏覽")
            If .Visible = True Then .Visible = False
        End With
    End If
End With
'    Application.Visible = False
    AppActivate "圖書管理"
'    Access.Application.SetOption "Behavior Entering Field", 0
    With ActCtl
        If .Parent.DefaultView = 0 Then 'Single Form
            .Parent.AllowEdits = True '設定此時,若非單一表單檢視則會將記錄移動至第一筆!2002/11/28
        End If
    '    Else
    '        Screen.ActiveForm.ActiveControl.form.ActiveControl.SetFocus
    ''        DoCmd.GoToRecord Screen.ActiveForm.CurrentRecord
'        If .Name = "札記" Then
'            .Locked = False
'        End If
        If .Locked = True Then .Locked = False
        If .Name <> "札記" Then
            .SetFocus
            .SelStart = 0 '全選
            .SelLength = Len(.text)
        Else
             .Value = Null
        End If
        blog.myaccess.docmd.RunCommand blog.myaccess.acCmdPaste
        .SelStart = s 's + 1 '設定插入點位置
        .SelLength = l
    End With
End Sub

Sub Access說明()
Attribute Access說明.VB_Description = "巨集錄製於 2001/11/28，錄製者 孫守真"
Attribute Access說明.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Access說明"
'
' Access說明 巨集
' 巨集錄製於 2001/11/28，錄製者 孫守真
'
    Selection.Paste
    Selection.WholeStory
    word.Application.Keyboard (1033)
    With Selection.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceAtLeast
        .LineSpacing = 0.4
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .WordWrap = True
    End With
End Sub
Sub 游標所在位置書籤()
Attribute 游標所在位置書籤.VB_Description = "巨集錄製於 2002/3/10，錄製者 孫守真"
Attribute 游標所在位置書籤.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.游標所在位置書籤"
'
' 游標所在位置書籤 巨集
' 巨集錄製於 2002/3/10，錄製者 孫守真
'指定鍵:F5(原指定給指令:EditGoTo)2004/12/14
On Error GoTo ErrH
HideWebBar

    With ActiveDocument.bookmarks
        .Add Range:=Selection.Range, Name:="游標_" & Replace(Replace(Replace(Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 4), "-", "_"), "(", "（"), ")", "）") '以因應主控文件模式之多重子文件(要扣除副檔名)2003/3/16
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
    ActiveDocument.Save
'    按下掃描鍵
Exit Sub
ErrH:
Select Case Err.Number
    Case 5828 '不正確的書籤名稱。
        If MsgBox("不正確的書籤名稱,是否略過？", vbOKCancel) = vbCancel Then
            Stop
            Resume
        Else
            Resume Next
        End If
    Case Else
        MsgBox Err.Number & Err.Description
End Select
End Sub
Sub 編輯處書籤()
' 2003/3/10--剛好距草創時屆一年矣！
    With ActiveDocument.bookmarks
        .Add Range:=Selection.Range, Name:="游標１_" & Replace(Replace(Replace(Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 4), "-", "_"), "(", "（"), ")", "）")
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
    ActiveDocument.Save
End Sub

Sub 到上一次停佇之游標處() '指定于(快速鍵)Alt+shift+F5 2009/5/6
Attribute 到上一次停佇之游標處.VB_Description = "巨集錄製於 2002/3/12，錄製者 孫守真"
Attribute 到上一次停佇之游標處.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.到上一次停佇之游標處"
'
' 到上一次停佇之游標處 巨集
' 巨集錄製於 2002/3/12，錄製者 孫守真
'
CheckSaved
HideWebBar
'On Error GoTo Ftnote
'    Selection.GoTo What:=wdGoToBookmark, Name:="游標_" & Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 4)
    With ActiveDocument.bookmarks
        '如此寫便不須有錯誤處理函式了
        .item("游標_" & Replace(Replace(Replace(Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 4), "-", "_"), "(", "（"), ")", "）")).Select
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
'Exit Sub
'Ftnote:
'Select Case Err.Number
'    Case 5678
'        With ActiveDocument.ActiveWindow.View
'            If .Type = wdNormalView Then
'            On Error GoTo Comts
'                .SplitSpecial = wdPaneFootnotes '註腳檢視,wdPaneComments 註解檢視
'            Else
'                .SeekView = wdSeekFootnotes
'            End If
'        End With
'        Resume
'    Case Else
'        MsgBox Err.Number & Err.Description
'End Select
'Exit Sub
'Comts:
'Select Case Err.Number
'    Case 4198
'        With ActiveDocument.ActiveWindow.View
'            If .Type = wdNormalView Then
'                .SplitSpecial = wdPaneComments
'            Else
'                .SeekView = wdSeekFootnotes
'            End If
'        End With
'        Resume
'    Case Else
'        MsgBox Err.Number & Err.Description
'End Select
End Sub

Sub 到編輯處_游標1()
CheckSaved
'On Error GoTo Ftnote
'    Selection.GoTo What:=wdGoToBookmark, Name:="游標１_" & Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 4)
    With ActiveDocument.bookmarks
        '如此寫便不須有錯誤處理函式了
        .item("游標１_" & Replace(Replace(Replace(Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 4), "-", "_"), "(", "（"), ")", "）")).Select
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
'Exit Sub
'Ftnote:
'Select Case Err.Number
'    Case 5678
'        With ActiveDocument.ActiveWindow.View
'            If .Type = wdNormalView Then
'            On Error GoTo Comts
'                .SplitSpecial = wdPaneFootnotes '註腳檢視,wdPaneComments 註解檢視
'            Else
'                .SeekView = wdSeekFootnotes
'            End If
'        End With
'        Resume
'    Case Else
'        MsgBox Err.Number & Err.Description
'End Select
'Exit Sub
'Comts:
'Select Case Err.Number
'    Case 4198
'        With ActiveDocument.ActiveWindow.View
'            If .Type = wdNormalView Then
'                .SplitSpecial = wdPaneComments
'            Else
'                .SeekView = wdSeekFootnotes
'            End If
'        End With
'        Resume
'    Case Else
'        MsgBox Err.Number & Err.Description
'End Select
End Sub

Sub QuitClose()
If Documents.Count = 0 Then word.Application.Quit wdDoNotSaveChanges: Exit Sub
With ActiveDocument '2003/3/25改良
If Left(.Name, 2) = "文件" And _
    IsNumeric(Mid(.Name, 3)) And _
        .AttachedTemplate.Name = "Normal.dot" Then
        If .Windows.Count = 1 Then
            DonotSave = True
            .Close wdDoNotSaveChanges
        Else
            .ActiveWindow.Close
        End If
        With word.Application
            If Documents.Count = 0 Then
                .CommandBars("論文＿札記瀏覽").Visible = False
'                .Position = msoBarTop
                .Quit wdPromptToSaveChanges
                End 'Exit Sub
            End If
        End With
Else
    If .Windows.Count = 1 Then
            DonotSave = True
            .Close wdDoNotSaveChanges
    Else
        If .Saved = False Then
            Select Case MsgBox("文件已變更，是否要在本視窗儲存？" & vbCr & vbCr _
                & "不儲存請按否!", vbExclamation + vbYesNoCancel)
                Case vbYes
                    .Save
                    .ActiveWindow.Close
                Case vbCancel
                    .ActiveWindow.Close
                    .Save
            End Select
        Else
            DonotSave = True
            .ActiveWindow.Close wdDoNotSaveChanges
        End If
    End If
End If
DonotSave = False
End With
按下掃描鍵
End Sub

Sub 在圖書管理中尋找選取字串() '原名「尋找選取字串」
Dim Mystr As String, ctl As Control, ctlSourceName As String ', f As Byte '快速鍵:Alt+Z
Dim C As Integer
CheckSaved

With Selection
If .Type = wdSelectionIP Then MsgBox "請選取想要尋找之文字", vbExclamation: Exit Sub
If .Type = wdSelectionNormal Then ' <> wdNoSelection OR wdSelectionIP Then '不為插入點
'If .Text <> "" Then
    If VBA.right(.Range, 1) Like Chr(13) Then
        Mystr = Mid(.Range, 1, Len(.Range) - 1)
    Else
        Mystr = .Range
    End If
    Mystr = Replace(Mystr, Chr(13), Chr(13) & Chr(10)) '.Text'因為Access與Word換行所存的值不同!
    Mystr = Replace(Mystr, Chr(11), Chr(13) & Chr(10)) '.Text'因為Access與Word換行所存的值不同!
'    .Font.Color = wdColorRed
'    .Collapse wdCollapseEnd
    On Error GoTo 備註
'    setOX
'    OX.WinActivate "圖書管理"
    AppActivate "圖書管理"
    If myaccess Is Nothing Then
        Set myaccess = GetObject("D:\千慮一得齋\書籍資料\圖書管理.mdb")
    End If
'    AppActivate "圖書管理" ', True '因為有時要儲存文件，可能須先待Word得到焦點處理完後。
    If myaccess.CurrentObjectName = "原文" Then myaccess.docmd.RunCommand blog.myaccess.acCmdWindowHide ' Screen.ActiveForm.Visible = False
'cl: For Each Ctl In Screen.ActiveForm.Controls '2003/3/17
'        If TypeName(Ctl) = "textbox" Then
'            If Ctl.ControlSource Like "[札劄]記" Then '= "札記" Then
'                Ctl.SetFocus
'                ctlSourceName = Ctl.ControlSource
'                Exit For
'            End If
'        End If
'    Next Ctl
cl: For C = 0 To myaccess.screen.activeform.Controls.Count - 1 '2006/4/21
        Set ctl = myaccess.screen.activeform.Controls(C)
        If TypeName(ctl) = "textbox" Then
            If ctl.ControlSource Like "[札劄]記" Then '= "札記" Then
                ctl.SetFocus
                ctlSourceName = ctl.ControlSource
                Exit For
            End If
        End If
    Next C

'    If TypeName(Ctl) = "Nothing" Then
    If ctl Is Nothing Then
'    If TypeName(Ctl) = "textbox" Then
'        If Ctl.ControlSource <> "札記" Then
            If MsgBox("沒有[札記]欄位資料!" & vbCr & vbCr & "是 否 要 看 其 他 表 單？" _
                , vbExclamation + vbYesNo) = vbYes Then
                'f = 1
                GoTo 其他表單
            Else
                End
    '            Exit Sub
            End If
    End If
'    DoEvents
    With ctl.Parent
        If .Dirty = True Then 'Ctl.Parent.Refresh
            If .AllowEdits = False Then
                .AllowEdits = True
                myaccess.docmd.RunCommand blog.myaccess.acCmdSaveRecord
                .AllowEdits = False
            Else
                myaccess.docmd.RunCommand blog.myaccess.acCmdSaveRecord
            End If
        End If
''        '.RecordsetClone.FindFirst ctlSourceName & " like " & Chr$(34) & "*" _
''         & Mystr & "*" & Chr$(34) & ""
'        If InStr(.Recordset.Fields(ctlSourceName), Mystr) = 0 Then
'            With .RecordsetClone
'                Do
'                    .MoveNext
'                    If .EOF Then Exit Do
'                Loop While InStr(.Fields(ctlSourceName), Mystr) = 0
'            End With
'        End If
        Dim ff As Boolean
        ff = myaccess.Run("尋找字串_ole用", .Recordset, ctlSourceName, Mystr) '.Recordset, ctlSourceName, Mystr)
'    If .RecordsetClone.NoMatch Then
    If ff Then
        Select Case MsgBox("找不到!!" & vbCr & vbCr & "是 否 要 看 其 他 表 單？" _
                & vbCr & vbCr & "※要結束請按〔取消〕!", vbInformation _
                    + vbYesNoCancel, "目前表單是： " & myaccess.screen.activeform.Name)
            Case Is = vbYes
            'f = 2
                GoTo 其他表單
            Case vbCancel
'                AppActivate .Application.Caption, True
'                .Application.ActiveWindow.Activate
                End
            Case vbNo
                Selection.Copy '複製以備貼上用!2003/3/28
                AppActivate "圖書管理"
                myaccess.Forms(0).SetFocus
                End
        End Select
    Else
'        .Recordset.Bookmark = .RecordsetClone.Bookmark
''        Ctl.Parent.Recordset.Bookmark = Ctl.Parent.RecordsetClone.Bookmark
'        If Ctl.Parent.CurrentRecord > Ctl.Parent.RecordsetClone.AbsolutePosition + 1 Then
'            myaccess.DoCmd.FindRecord Mystr, acAnywhere, True, acUp, True, , False
''        '    On Error Resume Next
'        ElseIf Ctl.Parent.CurrentRecord < Ctl.Parent.RecordsetClone.AbsolutePosition + 1 Then
'            myaccess.DoCmd.FindRecord Mystr, acAnywhere, True, acDown, True, , False
'        End If
'        With myaccess.Forms("札").札記
'            .SelStart = InStr(.Value, Mystr) - 1
'            .SelLength = Len(Mystr)
'        End With
        With ctl
            If Not .SelText Like Mystr Then
                If InStr(.Value, Mystr) <> 0 Then
                    .SelStart = InStr(.Value, Mystr) - 1
                    .SelLength = Len(Mystr)
                End If
            End If
        End With
        Beep '尋找完音響提示
    End If
    End With
End If
End With
Exit Sub
備註:
'Stop '檢查用
Select Case Err.Number
    Case 5 '呼叫引數不正確'程序呼叫或引數不正確(即 AppActivate的引數有誤!)
'        MsgBox Err.Number & Err.Description
        If myaccess.CurrentProject.AllForms("原文").IsLoaded Then blog.myaccess.Forms("原文").Visible = False 'docmd.Close acForm,"原文",acSaveNo
        AppActivate myaccess.CurrentObjectName
'        AppActivate "圖書管理 - [" & Screen.ActiveControl.Parent.Caption & "]"
'        AppActivate "圖書管理"
        Resume Next
'        MsgBox "圖書管理資料庫沒有開啟！", vbExclamation
'        End
    Case 2475
        Select Case MsgBox("圖書管理資料庫沒有作用中的表單！" & vbCr & "是否要由系統來尋找？", vbExclamation + vbOKCancel)
            Case Is = vbOK
'                Raise
'其他表單:        Dim frm As form
其他表單:       Dim i As Integer '要依開啟先後倒著找才合操作習慣2003/3/18
'                For Each frm In access.Forms
'                    If MsgBox("目前作用中的表單是： " & frm.Name & Space(3) & vbCr & vbCr _
                        & "要 看 下 一 個 表 單 嗎？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                        
'                For i = 0 To Access1.CurrentProject.AllForms.Count
'                    If CurrentProject.AllForms(i).IsLoaded Then Forms(i).SetFocus
'                    Exit For
'                Next
                For i = myaccess.Forms.Count - 1 To 0 Step -1
                    If blog.myaccess.Forms(i).Name <> blog.myaccess.CurrentObjectName And blog.myaccess.Forms(i).Visible Then  ' Screen.ActiveForm.Name Then
                       If myaccess.Forms(i).Name <> "掃描" Then
                            If MsgBox("目前作用中的表單是： " & myaccess.Forms(i).Name & space(3) & vbCr & vbCr _
                                & "要 看 下 一 個 表 單 嗎？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                                AppActivate "圖書管理"
    '                            frm.SetFocus
                                myaccess.Forms(i).SetFocus
                                myaccess.docmd.Restore
        '                        If Ctl Is Nothing Then Set Ctl = frm.ActiveControl
                                If ctl Is Nothing Then Set ctl = myaccess.Forms(i).ActiveControl
                                If ctl.Parent.Dirty = True Then ctl.Parent.Refresh
        ''                        If Not IsEmpty(f) Then GoTo cl
                                If Err.Number = 0 Then GoTo cl '沒錯誤，表示在尋找表單，找到表單後，要重新找控制項（Cl)2003/3/22
                                Exit For
                            End If
                        End If
                    End If
                    If i = 0 Then MsgBox "已瀏覽完畢，沒有您適合的表單，請自行擇選..", vbExclamation: End
                Next
                GoTo cl '找完表單後重新尋找!2005/3/24
'                    If frm Is Nothing Then MsgBox "已瀏覽完畢，沒有您適合的表單，請自行擇選..", vbExclamation: End
            Case Is = vbCancel
                End
        End Select
    Case 20 '回復且無錯誤！
        Resume Next
    Case 2137 '目前不能尋找,則再試!2005/3/17
        DoEvents
'        AppActivate ActiveWindow.Application
'        If MsgBox("目前不能搜尋,是否再試?", vbOKCancel + vbExclamation) = vbOK Then
            AppActivate myaccess.CurrentObjectName
            myaccess.screen.ActiveControl.Parent.SetFocus
            With myaccess.screen.ActiveControl
                If .ControlSource = "札記" Then
                    .SetFocus: Resume
                Else
                    Stop
                End If
            End With
'        End If
    Case 2110 '不能移動到札記欄位'2005/4/1
'        Dim c As Control
'        For Each c In Screen.ActiveControl.Parent
'            If TypeName(c) = "Textbox" Then '文字方塊
'                If c.ControlSource = "札記" Then c.SetFocus
'            End If
'        Next
        Resume
    Case Else
        MsgBox Err.Number & Err.Description
        End
End Select
'Resume
End Sub
Sub 在本文件中取代選取字串格式_取代成紅字()
With ActiveWindow.Selection '快速鍵：
If .Type = wdSelectionIP Then MsgBox "請選取想要尋找之文字", vbExclamation: Exit Sub
If .Type = wdSelectionNormal Then ' <> wdNoSelection OR wdSelectionIP Then '不為插入點
    If InStr(ActiveDocument.Content, .text) = InStrRev(ActiveDocument.Content, .text) Then
        MsgBox "本文只有此處!", vbInformation
        .Font.Color = wdColorRed
        .Font.Bold = True
        Exit Sub
    End If
    .Find.ClearFormatting
    .Find.Replacement.Font.Color = wdColorRed
    .Find.Replacement.Font.Bold = True
    .Find.Execute FindText:=.text, MatchCase:=True, Replace:=wdReplaceAll, Replacewith:=.text, Wrap:=wdFindContinue
    '一定要有Wrap:=wdFindContinue否則單向找尋時,預設值為Wrap:=wdFindStop
End If
End With
End Sub
Sub 在本文件中取代選取字串格式()
With Selection '快速鍵：
If .Font.Color = wdColorAutomatic Or .Font.Color = wdColorBlack Then _
    If MsgBox("請先指定字形色彩!" & vbCr & _
        "若要保留黑字請按〔取消〕", vbExclamation + vbOKCancel) = vbOK Then Exit Sub
If .Type = wdSelectionIP Then MsgBox "請選取想要尋找之文字", vbExclamation: Exit Sub
If .Type = wdSelectionNormal Then ' <> wdNoSelection OR wdSelectionIP Then '不為插入點
    If InStr(ActiveDocument.Content, .text) = InStrRev(ActiveDocument.Content, .text) Then MsgBox "本文只有此處!", vbInformation: Exit Sub
    .Find.ClearFormatting
    .Find.Replacement.ClearFormatting
'    Dim aFont
'    Set aFont = .Words.First.Font.Duplicate
'    .Find.Replacement.Font = aFont
    
    .Find.Replacement.Font.Color = .Font.Color
    .Find.Replacement.Font.Bold = .Font.Bold
    .Find.Replacement.Font.Italic = .Font.Italic
    .Find.Replacement.Font.Size = .Font.Size
    .Find.Replacement.Font.Name = .Font.Name
    .Find.Replacement.Font.NameAscii = .Font.NameAscii
    .Find.Replacement.Font.Underline = .Font.Underline
    .Find.Replacement.Font.Borders = .Font.Borders
    .Find.Replacement.Font.Outline = .Font.Outline
    .Find.Replacement.Font.position = .Font.position
    .Find.Replacement.Font.Animation = .Font.Animation
    .Find.Replacement.Font.Spacing = .Font.Spacing
    .Find.Replacement.Font.EmphasisMark = .Font.EmphasisMark
    .Find.Replacement.Font.Emboss = .Font.Emboss
    .Find.Replacement.Font.Engrave = .Font.Engrave
    .Find.Replacement.Font.Hidden = .Font.Hidden
    .Find.Replacement.Font.ItalicBi = .Font.ItalicBi
    .Find.Replacement.Font.Kerning = .Font.Kerning
    .Find.Replacement.Font.NameFarEast = .Font.NameFarEast
    .Find.Replacement.Font.NameOther = .Font.NameOther
    .Find.Replacement.Font.Scaling = .Font.Scaling
'    .Find.Replacement.Font.Shading = .Font.Shading
    .Find.Replacement.Font.Shadow = .Font.Shadow
    .Find.Replacement.Font.SizeBi = .Font.SizeBi
    .Find.Replacement.Font.Subscript = .Font.Subscript
    .Find.Replacement.Font.Superscript = .Font.Superscript
    .Find.Replacement.Font.UnderlineColor = .Font.UnderlineColor
    .Find.Execute FindText:=.text, MatchCase:=True, Replace:=wdReplaceAll, Replacewith:=.text, Wrap:=wdFindContinue
    '一定要有Wrap:=wdFindContinue否則單向找尋時,預設值為Wrap:=wdFindStop
End If
End With
End Sub
Sub BopomofoWithBlankCharDirect()
'
' BopomofoWithBlankCharDirect 巨集
' 巨集建立於 2002/11/10，建立者 孫守真
'

End Sub
Sub 字形轉換_華康儷粗黑()
Attribute 字形轉換_華康儷粗黑.VB_Description = "巨集錄製於 2003/1/10，錄製者 孫守真"
Attribute 字形轉換_華康儷粗黑.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.字形轉換_華康儷粗黑"
'指定鍵：Shift+Alt+W
' 字形轉換_華康儷粗黑 巨集
' 巨集錄製於 2003/1/10，錄製者 孫守真
'
Static fontName As String

CheckSaved

If Selection.Font.Name <> "華康儷粗黑" Then
    fontName = Selection.Font.Name
    Selection.Font.Name = "華康儷粗黑"
Else '復原為原先字形2003/1/11(或按指定鍵:Ctrl+Spacebar--見線上說明)2003/1/12
    Selection.Font.Name = fontName
End If
End Sub

Sub Timer()
'Application.OnTime When:=Now + TimeValue("00:00:10"), _
    Name:="Timer"
word.Application.OnTime When:=Now + TimeValue("00:10:00"), _
    Name:="游標所在位置書籤" '"Project1.Module1.Macro1"
Stop '檢查用
End Sub

Sub 檢視資料庫資料() '2003/2/10
Dim SearchedText As String '指定鍵：Alt+Q '2009.9.17資料庫已龐大,不適用矣!
On Error GoTo errs
'因為開啟資料庫耗系統資源，往往導致當機，因此先予儲存！2003/3/26札
CheckSaved

With Selection
If .Type = wdSelectionIP Then MsgBox "請選取想要尋找之文字", vbExclamation: Exit Sub
If .Type = wdSelectionNormal Then ' <> wdNoSelection OR wdSelectionIP Then '不為插入點
'If SearchedText <> "" Then
'    SearchedText = .Text'今改以下式:2004/11/11
    SearchedText = Replace(.Range, Chr(13), Chr(13) & Chr(10)) '.Text'因為Access與Word換行所存的值不同!
    SearchedText = Replace(SearchedText, Chr(11), Chr(13) & Chr(10)) '.Text'因為Access與Word換行所存的值不同!
    Dim Access As Object
    Set Access = CreateObject("access.application")
'    '以上一行及原來下面的一行.OpenCurrentDatabase可以改成如下一行,蓋如下一行只會開啟一個Access(不管執行幾次); _
    若照原來上面那行再打開資料庫,則每次即會開啟一個2003/12/14
'    Set access = GetObject("d:\千慮一得齋\書籍資料\圖書管理_查詢版.mdb")
    Access.UserControl = True
''    access.UserControl = False '如果用False還會使尋找下一個至末筆時，不會顯示訊息方塊2003/3/25
'    access.Visible = True
    Access.OpenCurrentDatabase "d:\千慮一得齋\書籍資料\圖書管理_查詢版.mdb"
    Access.docmd.OpenForm "札_查詢", , , , , , "Word"
''    access.DoCmd.Close acForm, "主表單", acSaveNo
''    access.Forms("札_查詢").RecordSource = "札_查詢"
    Access.Forms("札_查詢").Controls("關鍵字").SetFocus
''    SendKeys Selection.Text & "{tab}"
    Access.Forms("札_查詢").Controls("關鍵字").text = SearchedText
''    access.Forms("札_查詢").Controls("關鍵字") = SearchedText
''    SendKeys "~"
    Set Access = Nothing
End If
Exit Sub
errs:
    Select Case Err.Number
        Case 7866 '已開啟
            Select Case MsgBox("已開啟，是否要關閉再繼續？", vbYesNoCancel + vbExclamation)
                Case vbYes
                    Dim Access1 As Object
                    Set Access1 = GetObject("d:\千慮一得齋\書籍資料\圖書管理_查詢版.mdb")
'                    Set access = Access1
'                    If Access1.Visible = False Then Access1.Visible = True
'                    If Access1.CurrentProject.AllForms("札_查詢").IsLoaded Then Access1.DoCmd.Close acForm, "札_查詢", acSaveNo
'                    If Access1.CurrentProject.AllForms("主_子題檢視").IsLoaded Then Access1.DoCmd.Close acForm, "主_子題檢視", acSaveNo
'                    Access1.CloseCurrentDatabase
                    Access1.Application.Quit blog.myaccess.acExit
'                    Access1.Quit
                    Set Access1 = GetObject(, "access.application")
                    Access1.Quit blog.myaccess.acExit
                    Set Access1 = Nothing
                    Set Access = CreateObject("access.application")
                    Access.UserControl = True
                    Access.OpenCurrentDatabase "d:\千慮一得齋\書籍資料\圖書管理_查詢版.mdb"
                    Resume Next
                Case vbNo
                    Access.Quit blog.myaccess.acExit '把新開的Access關掉
'                    重近取得參照
                    Set Access = GetObject("d:\千慮一得齋\書籍資料\圖書管理_查詢版.mdb")
'                    Set access = CreateObject("d:\千慮一得齋\書籍資料\圖書管理_查詢版.mdb")
'                    access.UserControl = True
'                    If access.Visible = False Then access.Visible = True
'                    If Access1.CurrentProject.AllForms("札_查詢").IsLoaded Then Access1.DoCmd.Close acForm, "札_查詢", acSaveNo
'                    If Access1.CurrentProject.AllForms("主_子題檢視").IsLoaded Then Access1.DoCmd.Close acForm, "主_子題檢視", acSaveNo
                    Resume Next
                Case Else
    '                Set access = GetObject("access.application")
    '                access.CloseCurrentDatabase
                    Access.Quit   '2003/2/22
                    Set Access = Nothing
    '                Stop
                    blog.myaccess.Application.DDETerminateAll
                    End
            End Select
        Case 4605 '此方法或屬性無法使用，因為 這部份的主文件在編輯鎖定狀態中.
            On Error Resume Next
            If Selection.Information(wdInMasterDocument) Then  '如果是主控文件
            '或者寫成:
'            If ActiveDocument.IsMasterDocument = True Then
                Dim subdoc, wins
                For Each subdoc In ActiveDocument.Subdocuments
                    If subdoc.Locked Then
                        If InStr(subdoc.Name, "自動回復") Then
                            Documents(Mid(subdoc.Name, 7, Len(subdoc.Name) - 10)).Activate
'                            subdoc.Locked = False
'                        Else
'                            If MsgBox("子文件已開啟,請先關閉子文件,再操作!", vbExclamation + vbOKCancel) = vbOK Then
'                            Documents(Mid(subdoc.Name, 7, Len(subdoc.Name) - 10)).Close
'                            End If
                        Else
                            For Each wins In Windows
                                If wins.Caption = Left(subdoc.Name, Len(subdoc.Name) - 4) Then
                                    Documents(subdoc.Name).Activate
                                Else
                                    subdoc.Locked = False
'                                    subdoc.Parent.Undo
                                End If
                            Next wins
                        End If
                    End If
                Next subdoc
                Resume
'            If ActiveDocument.Subdocuments(3).Locked Then ActiveDocument.Subdocuments(1).Locked = False
'            if ActiveDocument.Subdocuments
            Else
                MsgBox Err.Number & ":" & Err.Description, vbExclamation
            End If
        Case Else
            MsgBox Err.Number & ":" & Err.Description, vbExclamation
    End Select
End With
End Sub
Sub 檢視資料庫主子題() '2003/2/10
CheckSaved

If Selection.Type <> wdSelectionIP Or wdNoSelection Then
'If Selection.Text <> "" Then
    Dim Access As Object
    Set Access = CreateObject("access.application")
    Access.UserControl = True
    Access.OpenCurrentDatabase "d:\千慮一得齋\書籍資料\圖書管理_查詢版.mdb"
    Access.docmd.OpenForm "主_子題檢視", , , , , , "Word"
'    access.DoCmd.Close acForm, "主表單", acSaveNo
'    access.Forms("主_子題檢視").SetFocus
    Access.Forms("主_子題檢視").Controls("Text18").SetFocus
'    access.Forms("主_子題檢視").Controls("Text18").Text = Selection.Text
    SendKeys Selection.text '上一行無效
    Set Access = Nothing
End If
End Sub

Sub 新增札箋()
CheckSaved

If Selection.Type <> wdSelectionIP Or wdNoSelection Then
在圖書管理中尋找選取字串
0   Select Case blog.myaccess.screen.activeform.Name '2003/3/30
        Case "札", "札記_藏", "札_查詢"
            blog.myaccess.Forms(blog.myaccess.screen.activeform.Name).Label19_Click
        Case "札記_類查詢"
            blog.myaccess.docmd.RunMacro "記錄處理.新增札箋"
'    Else
'        If MsgBox("沒有 [札記] 表單!   無法新增札箋..." & vbCr & vbCr & "是 否 要 看 其 他 表 單？" _
'                    , vbExclamation + vbYesNo) = vbNo Then End 'Exit Sub
'        Dim i As Integer '要依開啟先後倒著找才合操作習慣2003/3/18
'    '    For Each frm In Forms
'        For i = Forms.Count - 1 To 0 Step -1
'            If MsgBox("目前作用中的表單是： " & vbCr & Forms(i).Name & Space(3) & vbCr & vbCr _
'                & "要 看 下 一 個 表 單 嗎？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
'    '                    If CurrentProject.AllForms(Forms(i).Name).IsLoaded Then'forms本身則已限在已開啟的表單了！
'                AppActivate "圖書管理"
'                Forms(i).SetFocus
'                DoCmd.Restore '
'                GoTo 0
'                Exit For
'            End If
'            If i = 0 Then MsgBox "已瀏覽完畢，沒有您適合的表單，請自行擇選..", vbExclamation: End
'        Next i
    End Select
    On Error GoTo e
    AppActivate "圖書管理"

    
''If Selection.Text <> "" Then
''    Dim access As Object
''    Set access = CreateObject("access.application")
''    access.UserControl = True
''    access.OpenCurrentDatabase "d:\千慮一得齋\書籍資料\圖書管理_查詢版.mdb"
''    Set access = GetObject("d:\千慮一得齋\書籍資料\圖書管理_查詢版.mdb")
'    AppActivate "圖書管理"
'    If Not CurrentProject.AllForms("札").IsLoaded And Not CurrentProject.AllForms("札記_藏").IsLoaded Then Exit Sub
'    If CurrentProject.AllForms("札").IsLoaded Then
'        Forms("札").SetFocus
'        Forms("札").label19_Click
'    End If
'    If CurrentProject.AllForms("札記_藏").IsLoaded Then
'        Forms("札記_藏").SetFocus
'        Forms("札記_藏").label19_Click
'    End If
''            access.Forms("札").Controls("Text18").SetFocus
'''    access.Forms("主_子題檢視").Controls("Text18").Text = Selection.Text
''    SendKeys Selection.Text '上一行無效
''    Set access = Nothing
End If
Exit Sub
e:
Select Case Err.Number
    Case 5 '程序呼叫或引數不正確(即 AppActivate的引數有誤!)
        AppActivate "圖書管理 - [" & blog.myaccess.screen.ActiveControl.Parent.Caption & "]"
        Resume Next
    Case Else
        MsgBox Err.Number & Err.Description
End Select
End Sub

Public Sub 刪除已閱札記()
'指定鍵：Ctrl+Shift+Del'或刪除文件內所有選取範圍的文字
With Selection '如果沒選取則先選取
    If Not .Type = wdSelectionNormal Then
        .HomeKey unit:=wdStory, Extend:=wdExtend
        .Delete
    Else
        ActiveDocument.Range.Find.Execute Selection.text, , , , , , True, wdFindContinue, , "", wdReplaceAll
    End If
End With
End Sub

Public Sub 剪下以貼上札記()
'指定鍵：Ctrl+Num 0

Dim p As Byte
With Selection '如果沒選取則先選取
    If .Document.path <> "" Then MsgBox "此文件不能操作", vbExclamation: Exit Sub
    If .Type = wdSelectionIP Then
        p = MsgBox("確定「跨頁」嗎？", vbQuestion + vbYesNoCancel)
        If p = vbNo Then End 'Exit Sub'2003/11/20
        .HomeKey unit:=wdStory, Extend:=wdExtend
        .Cut
        Do While Asc(.text) = 10 Or Asc(.text) = 13
        If .End + 1 = .Document.Content.End Then Exit Do
        .Delete
        Loop
'        AppActivate "圖書管理", True
        AppActivate blog.myaccess.CurrentObjectName, True
        With blog.myaccess.screen.activeform
            If .RecordSource <> "札_" Then Exit Sub
            If Not .NewRecord Then Exit Sub
            If IsNull(.Controls("頁").DefaultValue) Or .Controls("頁").DefaultValue = "" Then Exit Sub
            .Controls("頁") = .Controls("頁").DefaultValue '札記表單會設定SetFocus故須如此
'            If .ActiveControl.ControlSource <> "札記" Then _
'                .Controls("札記").SetFocus
            Do Until .ActiveControl.ControlSource = "札記"
                .Controls("札記").SetFocus
            Loop '2003/11/22
            blog.myaccess.docmd.RunCommand blog.myaccess.acCmdPaste
            If p = vbYes Then .Controls("跨頁") = True
            blog.myaccess.docmd.GoToRecord blog.myaccess.acDataForm, .Name, blog.myaccess.acNewRec
'            Do While .ActiveControl.ControlSource <> "札記"'因為貼上前已有檢查，故不再聚焦了2004/8/27
'                Screen.PreviousControl.SetFocus
'            Loop
        End With
    Else
        MsgBox "請置入插入點", vbExclamation
    End If
'    If Len(.Document.Content) = 1 Then QuitClose
    AppActivate .Application
    '紅樓夢專用：開啟舊檔'2003/11/23
    If Len(.Document.Content) > 1 Then
'        SendKeys "^f", True     '開啟尋找方塊
    Else
''        SendKeys "^o", True
'        With Forms("篇")
'            .SetFocus
'            If IsNull(.Controls("卷")) Then .Controls("卷") = InputBox("請輸入卷數!!")
'        End With
'        With Selection
'            .Application.Documents.Open Dir("D:\千慮一得齋\資料庫\文字檔資料庫\小說\紅樓夢" & "\*" & Screen.ActiveForm.Controls("卷") + 1 & "*")
'            .Application.Documents(1).Activate
'            With Selection
'                .EndKey wdStory, wdMove
'                .HomeKey wdStory, wdExtend
'                .Copy
'                .Document.Close
'            End With
'            .Paste
'            With .Find
'                .Text = "【回前一頁】 【紅樓首頁】 【紅樓全文】 【上一章回】 【下一章回】"
'                Do While .Execute(, , , , , , , wdFindContinue)
'                    .Parent.Paragraphs(1).Range.Delete
'                Loop
'            End With
'            .Range.Find.Execute " ", True, , , , , , wdFindContinue, , "　", wdReplaceAll
'            .HomeKey wdStory, wdMove
'        End With
'        'Dir("D:\*文字檔資料庫\小說\紅樓夢" & Screen.ActiveForm.Controls("卷") + 1 & "*")
    End If
End With
End Sub

Public Sub 清除分行符號()
'2003/4/3指定鍵:Shift+Backspace
Dim rp As String, p(2) As String, i As Byte
p(1) = Chr(10): p(2) = Chr(13)
With Selection '.Find '如此才不會更改字形格式！
    If VBA.right(.Range, 1) Like p(1) Or VBA.right(.Range, 1) Like p(2) Then .MoveLeft wdCharacter, 1, wdExtend
''        .Range.Select(.Range.Words.Count=
'    .ClearFormatting
    rp = .Range
    For i = 1 To 2
         rp = Replace(rp, p(i), "")
    Next i
    .Range = rp
    
'    .Execute findtext:="^p", Replacewith:="", Wrap:=wdFindContinue, Replace:=wdReplaceAll
'    .ClearFormatting
End With
End Sub

Public Sub 轉換分行符號為手動分行符號()
'2003/3/25
With Selection.Find '如此才不會更改字形格式！
    .ClearFormatting
    .Execute FindText:="^p", Replacewith:="^l", Wrap:=wdFindContinue, Replace:=wdReplaceAll
'    .ClearFormatting
End With
'2003/3/23
'Dim i As Integer, s As Integer, p As Byte, InStrs As Long, InStrRevs As Long
's = Selection.Information(wdActiveEndSectionNumber) '傳回所在之末節數！
'With ActiveDocument.Sections(s) '處理所在之節中的分行
'    If IsNumeric(.Range.Paragraphs(1).Range.Text) Then p = 1 '判斷首段是否為數字
'    InStrs = InStr(.Range.Text, Chr(13)): InStrRevs = InStrRev(.Range.Text, Chr(13))
'    If InStrs = InStrRevs And InStrs <> 0 Then _
'    If MsgBox("本節只有一個分行符號! 是否要取代？", vbInformation + vbYesNo) = vbNo Then .Range.Words(.Range.Words.Count).Select: Exit Sub
'    For i = .Range.Paragraphs.Count - 1 To 1 + p Step -1 '第一段若為頁碼則不要,最末一段分節符號取代後，成了分頁符號，故亦不要！
'        With .Range.Paragraphs(i)  '快速鍵：
'            .Range.Text = Replace(.Range.Text, Chr(13), Chr(11)) '轉換分行符號為手動分行符號
'        '    .Find.ClearFormatting
'        '    .Find.Replacement.ClearFormatting
'        '    .Find.Execute findtext:=.Text, MatchCase:=False, Replace:=wdReplaceAll, replacewith:=.Text, Wrap:=wdFindContinue
'        End With
'    Next i
'End With
End Sub


Sub 檢視這兩天內的編輯處() '2003/3/28
Static bk As Bookmark, ps As Byte, dt As String, ThisDoc As String, dtbefore As Integer
Dim dts As String, dtbeforeStr As String
With ActiveDocument
    If ThisDoc <> .Name Then Set bk = Nothing: ps = 0: dt = ""
    ThisDoc = .Name '換文件,則重設
    If bk Is Nothing Then
        Select Case MsgBox("要瀏覽「今天」的編輯處嗎?", vbQuestion + vbYesNoCancel)
            Case vbCancel
                End
            Case vbYes
                dt = "Edit_" & Format(Now(), "yyyy_mm_dd")
                dtbefore = 0
            Case vbNo
Again:          dtbeforeStr = InputBox("要看" & Chr(-24153) & "幾" & Chr(-24152) & _
                        "天以前的編輯處?", "瀏覽編輯處書籤", "1")
                If Not IsNumeric(dtbeforeStr) Then
                    If Not dtbeforeStr Like "" Then
                        MsgBox "請輸入數字！": GoTo Again
                    Else
                       End
                    End If
                End If
                dtbefore = CInt(dtbeforeStr)
                dt = "Edit_" & Format(Now() - dtbefore, "yyyy_mm_dd")
        End Select
        For Each bk In .bookmarks
            ps = ps + 1 '記下索引值
            With bk
                If InStr(.Name, dt) Then
                    bk.Select
    '                .GoTo wdGoToBookmark, wdGoToAbsolute, , bk.Name
                    GoTo Repeats
                    Exit For
                End If
            End With
        Next bk
        If bk Is Nothing Then MsgBox "沒有符合的書籤" & vbCr & _
            "(沒有編輯處記錄，或記錄已刪除)", vbExclamation, "瀏覽編輯處": End
    Else
Repeats: ps = ps + 1 '因為書籤預設之順序乃照名稱排序, _
                        故可如此用靜態變數ps來設計參照書籤的索引值
        If ps > .bookmarks.Count Then GoTo e
        Set bk = .bookmarks(ps)
'        If dt <> "Edit_" & Format(Now(), "yyyy_mm_dd") Then
        Select Case dtbefore
            Case Is > 2
                dts = "昔日（" & FormatDateTime(Date - dtbefore, vbLongDate) & "）"
            Case Is = 2
                dts = "前天"
            Case Is = 1
                dts = "昨天"
            Case Else
                dts = "今天"
        End Select
        If InStr(bk.Name, dt) = 0 Then
            '瀏覽完畢後
e:          .bookmarks(ps - 1).Select '選取（到）最後一個檢視的書籤
            MsgBox dts & "編輯處已瀏覽完畢!" & vbCr & vbCr & _
                "目前頁碼：" & Selection.Information(wdActiveEndAdjustedPageNumber) _
                , vbInformation
            End '用End 即可初始化一切變數
    '        Set bk = Nothing: ps = 0 '重新初始化
        Else
            Dim a As Paragraph '取得書籤所在位置標題'2003/4/26
            With Selection
                If Not .Information(wdInFootnote) Then
                    Set a = .Paragraphs(1)
                Else '書籤在註腳時的處理
                    Set a = .Footnotes(1).Reference.Paragraphs(1)
                End If
                Do
                    Set a = a.Previous
                   If Left(a.Style, 2) = "標題" Then _
                    Exit Do
                Loop 'a.Range會包括段落字元,要去除可用：Left(a.Range, Len(a.Range) - 1)
            End With
            Select Case MsgBox("要重新開始請按〔否〕!" & vbCr & vbCr & _
                "目前標題為：" & a.Range & _
                "；目前頁碼：" & Selection.Information(wdActiveEndAdjustedPageNumber), _
                vbQuestion + vbYesNoCancel, _
                "檢視「" & dts & "」的編輯處...要繼續嗎?")
                Case vbCancel
                    ps = ps - 1
                    Exit Sub
                Case vbNo
                    End
                Case vbYes
        '            .ActiveWindow.ScrollIntoView .Range, True
                    .bookmarks(ps).Select
                    GoTo Repeats
            End Select
        End If
    End If
End With
End Sub

Private Sub 瞭解書籤排序()
Dim bk As Bookmark, bkIdx As Integer
With ActiveDocument.bookmarks
    For bkIdx = 1 To .Count
'     .DefaultSorting = wdSortByName
        Debug.Print .item(bkIdx)
    Next bkIdx
End With
End Sub

Sub 插入交互參照() '2003/3/28'指定鍵：Ctrl+Shift+Insert
Dim CrossReference, i As Integer, CrossReferenceID As String
Static doinsert As Boolean, WinID As Byte, DocWin As Byte
Dim Winview As Byte, s As Long '2003/3/31
With ActiveDocument
    If Selection.Type = 2 Then
        WinID = .ActiveWindow.WindowNumber
        DocWin = .ActiveWindow.Previous.WindowNumber
        GoTo 1
    End If
    If doinsert = False Then
'        DocWin = .ActiveWindow.Index '記下使用中文件之將要插入交互參照之視窗索引值
        DocWin = .ActiveWindow.WindowNumber 'Index是全部視窗的編號，此乃本份文件之編號，二者不同！2003/4/1
        '如果是在註腳檢視(等特殊檢視區塊）處，則記下'2003/3/31
        With .ActiveWindow.View 'wdPaneNone=0
            If .SplitSpecial <> wdPaneNone Then
                Winview = .SplitSpecial
                s = .Application.Selection.start '記下插入點位置
            End If
        End With
        .Windows.Add
         With .ActiveWindow.View
            If Winview <> wdPaneNone Then 'wdPaneNone=0
                .SplitSpecial = Winview '設定特殊檢視視窗（如註腳...等）
                .Application.Selection.start = s '設定插入點位置
            End If
        End With
'        .ActiveWindow.Application.GoBack
        WinID = .ActiveWindow.WindowNumber
         MsgBox "請在本視窗中選取欲插入之交互參照物件" & vbCr & vbCr _
                & "選好後，再按一次，即可插入！", vbExclamation
        doinsert = True
    Else
1       Select Case .ActiveWindow.Selection.Range.StoryType
            Case wdMainTextStory
                Select Case Selection.Style
                    Case "標題 1", "標題 2", "標題 3", "標題 4", "標題 5", "標題 6" _
                        , "標題 7", "標題 8", "標題 9"
                        'wdStyleHeader  'Left(Selection.Style, 2) = "標題" '如果是標題
'                        CrossReferenceID = .ActiveWindow.Selection.HeaderFooter .Footnotes (1).Index
                        CrossReference = .GetCrossReferenceItems(wdRefTypeHeading)
                        For i = 1 To UBound(CrossReference)
'                            If Trim(Left(CrossReference(i), Len(CrossReferenceID))) _
                                = CrossReferenceID Then
                            If Trim(CrossReference(i)) Like Selection Then
                                If MsgBox("要插入註腳所在之「頁碼」而非「標題文字」，" _
                                        & "請按〔取消〕", vbQuestion + vbOKCancel, "插入標題:" & _
                                        .ActiveWindow.Selection.Style & .ActiveWindow.Selection) _
                                        = vbOK Then
                                    .Windows(DocWin).Activate '以原編輯視窗為準
                                    Selection.Range.InsertCrossReference _
                                        ReferenceType:=wdRefTypeHeading _
                                        , ReferenceKind:=wdContentText, _
                                            ReferenceItem:=i _
                                                , InsertAsHyperlink:=True
                                Else
                                    .Windows(DocWin).Activate '以原編輯視窗為準
                                    Selection.Range.InsertCrossReference _
                                        ReferenceType:=wdRefTypeHeading _
                                        , ReferenceKind:=wdPageNumber, _
                                            ReferenceItem:=i _
                                                , InsertAsHyperlink:=True
                                End If
                                Exit For
                            End If
                        Next i
                        If i > UBound(CrossReference) Then MsgBox "沒有合適的註腳參照，請手動操作！", vbExclamation: End
                        .Windows(WinID).Close
                        doinsert = False: WinID = 0: DocWin = 0
                    On Error GoTo ErrHs
                    Case "註腳參照" 'wdStyleFootnoteReference, wdStyleFootnoteText
                        CrossReferenceID = .ActiveWindow.Selection.Range.Footnotes(1).index
                        CrossReference = .GetCrossReferenceItems(wdRefTypeFootnote)
                        For i = 1 To UBound(CrossReference)
                            If Trim(Left(CrossReference(i), Len(CrossReferenceID))) _
                                = CrossReferenceID Then
                                If MsgBox("要插入註腳所在之「頁碼」而非「註腳編號」，" _
                                        & "請按〔取消〕", vbQuestion + vbOKCancel, "插入註腳:" & _
                                        .ActiveWindow.Selection.Range.Footnotes(1).index) _
                                        = vbOK Then
                                    .Windows(DocWin).Activate '以原編輯視窗為準
                            '        .ActiveWindow.Selection.Range.Paste
                                    Selection.Range.InsertCrossReference _
                                        ReferenceType:=wdRefTypeFootnote _
                                        , ReferenceKind:=wdFootnoteNumber, _
                                            ReferenceItem:=i _
                                                , InsertAsHyperlink:=True
                                Else
                                    .Windows(DocWin).Activate '以原編輯視窗為準
                                    Selection.Range.InsertCrossReference _
                                        ReferenceType:=wdRefTypeFootnote _
                                        , ReferenceKind:=wdPageNumber, _
                                            ReferenceItem:=i _
                                                , InsertAsHyperlink:=True
                                End If
                                Exit For
                            End If
                        Next i
                        If i > UBound(CrossReference) Then MsgBox "沒有合適的註腳參照，請手動操作！", vbExclamation: End
                        .Windows(WinID).Close
                        doinsert = False: WinID = 0: DocWin = 0
                    Case Else '插入書籤
                        For Each CrossReference In .bookmarks
                            If CrossReference.Range Like ActiveWindow.Selection Then
                                If MsgBox("插入書籤: " & CrossReference.Name, vbQuestion + vbOKCancel, "插入交互參照") = vbOK Then
                                    .Windows(DocWin).Activate '以原編輯視窗為準
                                    Selection.Range.InsertCrossReference _
                                        ReferenceType:=wdRefTypeBookmark _
                                        , ReferenceKind:=wdPageNumber, _
                                            ReferenceItem:=CrossReference.Name, _
                                            InsertAsHyperlink:=True
                                    .Windows(WinID).Close
                                    doinsert = False: WinID = 0: DocWin = 0
                                    Exit For
                                End If
                            End If
                        Next CrossReference
                        If IsObject(CrossReference) Then
                            If CrossReference Is Nothing Then MsgBox "沒有合適的書籤參照，請手動操作！", vbExclamation: End
                        Else
                            If IsEmpty(CrossReference) Then MsgBox "沒有合適的書籤參照，請手動操作！", vbExclamation: End
                        End If
                End Select
            Case wdFootnotesStory
                .ActiveWindow.Selection.Range.Copy
                .ActiveWindow.Close
                .Windows(DocWin).Activate
                .ActiveWindow.Selection.Range.Paste
        End Select
    End If
Exit Sub
'End With
ErrHs:
'With ActiveDocument
Select Case Err.Number
'    Case 5941 '集合中的成員不存在（即沒有註腳）,則插入書籤2003/3/29
'        For Each CrossReference In .Bookmarks
'            If CrossReference.Range Like ActiveWindow.Selection Then
'                If MsgBox("插入書籤: " & CrossReference.Name, vbQuestion + vbOKCancel, "插入交互參照") = vbOK Then
'                    .Windows(DocWin).Activate '以原編輯視窗為準
'                    Selection.Range.InsertCrossReference _
'                        ReferenceType:=wdRefTypeBookmark _
'                        , ReferenceKind:=wdPageNumber, _
'                            ReferenceItem:=CrossReference.Name, _
'                            InsertAsHyperLink:=True
'                    .Windows(WinID).Close
'                    doinsert = False: WinID = 0: DocWin = 0
'                    Exit For
'                End If
'            End If
'        Next CrossReference
'        If IsObject(CrossReference) Then
'            If CrossReference Is Nothing Then MsgBox "沒有合適的書籤參照，請手動操作！", vbExclamation: End
'        Else
'            If IsEmpty(CrossReference) Then MsgBox "沒有合適的書籤參照，請手動操作！", vbExclamation: End
'        End If
    Case Else
        MsgBox Err.Number & Err.Description: End
End Select
End With
End Sub
Public Sub 在另一文件中尋找選取字串()
Static winNum As Byte, preR As String
Dim r As String, ins(4) As Long, MnText As String, FnText As String
Dim d As Document, winINdex As Byte, startD As Document
'CheckSavedNoClear
With Selection '指定鍵：Alt+Ctrl+Up
'If Not .Text Like "" Then '快速鍵：Alt+Ctrl+Down
    If .Type = wdSelectionIP Then MsgBox "請選取想要尋找之文字", vbExclamation: Exit Sub
    If .Type = wdSelectionNormal Then '不為插入點
        If preR <> r Then winNum = 0
        r = .text
        preR = r
        Set startD = .Document
        On Error GoTo Previews
        For Each d In .Application.Documents
            winINdex = winINdex + 1
            If Not d Is startD And winINdex > winNum Then
            'With .Application.ActiveWindow.Document ' ActiveDocument
                With d
                    
                    MnText = .StoryRanges(wdMainTextStory) '用變數代對長篇文件來說較快！2003/4/8
                    ins(1) = InStr(MnText, r)
                    ins(2) = InStrRev(MnText, r)
                    If .Footnotes.Count > 1 Then
                        FnText = .StoryRanges(wdFootnotesStory)
                        ins(3) = InStr(FnText, r)
                        ins(4) = InStrRev(FnText, r)
                    End If
                    If ins(1) = 0 And ins(3) = 0 Then
                        Select Case MsgBox("沒有符合文字!" & vbCr & vbCr & _
                            "是否要找下一份文件？", vbExclamation + vbOKCancel)
                            Case vbOK
                                
                            Case vbCancel
                                winNum = winINdex
                                Exit Sub
                        End Select
                    End If
                    If ins(1) = ins(2) And ins(3) = ins(4) Then
                        d.Activate
                        MsgBox "本文只有此處!", vbInformation ': Exit Sub
                    End If
                    If ins(1) <> 0 Then
                        ins(1) = wdMainTextStory
                    Else
                        ins(1) = wdFootnotesStory
                    End If
                    With .StoryRanges(ins(1)).Find
        '            With Selection.Find
                        .ClearFormatting
                        .Replacement.ClearFormatting '這也要清除才行
                        .Forward = True
                        .Wrap = wdFindAsk
                        .MatchCase = True
                        .text = r
                        .Execute
                        .Parent.Select
                        d.Activate
                        With d.ActiveWindow
                            .ScrollIntoView Selection
                            If .WindowState = wdWindowStateMinimize Then
                                .WindowState = wdWindowStateNormal
                            End If
                        End With
'                        With .Application.ActiveWindow
'                            If .WindowState = wdWindowStateMinimize Then .WindowState = wdWindowStateMaximize
'                        End With
                        winNum = winINdex
                        Exit Sub
                    End With
                End With
            End If
        Next d
        winNum = 0
    End If
End With

Exit Sub
Previews:
Select Case Err.Number
'    Case 91
'        On Error Resume Next
'        Dim d As Byte
'        d = Documents.Count
'        If d > 1 Then
'            If Documents(d - 1) <> ActiveDocument Then
'                Documents(d - 1).Activate
'            Else
'                Documents(d).Activate
'            End If
'        End If
''        ActiveWindow.Previous.Document.ActiveWindow.Activate
'        Resume Next
'    Case 5941
'        Resume Next
    Case Else
        MsgBox Err.Number & Err.Description, vbExclamation
End Select
End Sub

Public Sub 在另一文件中尋找選取字串_old()
Static winNum As Byte
Dim r As String, ins(4) As Long, MnText As String, FnText As String
CheckSavedNoClear
With Selection
'If Not .Text Like "" Then '快速鍵：Alt+Ctrl+Down
    If .Type = wdSelectionIP Then MsgBox "請選取想要尋找之文字", vbExclamation: Exit Sub
    If .Type = wdSelectionNormal Then ' <> wdNoSelection OR wdSelectionIP Then '不為插入點
        r = .text
        On Error GoTo Previews
Again: If winNum <> 0 Then .Application.Windows(winNum).Activate
        If .Application.Documents.Count = 1 And .Document.Windows.Count > 1 Then
            If .Document.ActiveWindow.WindowNumber < .Document.Windows.Count Then
                .Document.ActiveWindow.Next.Activate
            Else
                .Document.ActiveWindow.Previous.Activate
            End If
        Else
            With .Application.ActiveWindow.Next.Document
                If .Windows.Count > 1 Then
                    .ActiveWindow.Activate
                Else
                    .Activate
                End If
            End With
        End If
        'If InStr(ActiveDocument.Name, "字表7") Then Register_Event_Handler: Documents("字表7.2.doc").Windows(1).Visible = True
        If InStr(ActiveDocument.Name, "字表") Then Register_Event_Handler: d字表.Windows(1).Visible = True
        With .Application.ActiveWindow.Document ' ActiveDocument
            MnText = .StoryRanges(wdMainTextStory) '用變數代對長篇文件來說較快！2003/4/8
            FnText = .StoryRanges(wdFootnotesStory)
            ins(1) = InStr(MnText, r)
            ins(2) = InStrRev(MnText, r)
            ins(3) = InStr(FnText, r)
            ins(4) = InStrRev(FnText, r)
            If ins(1) = 0 And ins(3) = 0 Then
                Select Case MsgBox("沒有符合文字!" & vbCr & vbCr & _
                    "是否要找下一份文件？", vbExclamation + vbYesNoCancel)
                    Case vbYes
                        '記下目前文件視窗：
                        winNum = .ActiveWindow.index '.WindowNumber
                        GoTo Again
                    Case vbNo
                        .Application.ActiveWindow.Previous.Activate
                         Exit Sub 'End
                    Case vbCancel
                        Exit Sub 'End'end 會重設所有變數及設定值,包括使用application的Register_Event_Handler
                End Select
            End If
            If winNum <> Empty Then winNum = Empty
            If ins(1) = ins(2) And ins(3) = ins(4) Then _
                MsgBox "本文只有此處!", vbInformation ': Exit Sub
            If ins(1) <> 0 Then
                ins(1) = wdMainTextStory
            Else
                ins(1) = wdFootnotesStory
            End If
            With .StoryRanges(ins(1)).Find
'            With Selection.Find
                .ClearFormatting
                .Replacement.ClearFormatting '這也要清除才行
                .Forward = True
                .Wrap = wdFindAsk
                .MatchCase = True
                .text = r
                .Execute
                .Parent.Select
                With .Application.ActiveWindow
                    If .WindowState = wdWindowStateMinimize Then .WindowState = wdWindowStateMaximize
                End With
            End With
        End With
    End If
End With
Exit Sub
Previews:
Select Case Err.Number
    Case 91
        On Error Resume Next
        Dim d As Byte
        d = Documents.Count
        If d > 1 Then
            If Documents(d - 1) <> ActiveDocument Then
                Documents(d - 1).Activate
            Else
                Documents(d).Activate
            End If
        End If
'        ActiveWindow.Previous.Document.ActiveWindow.Activate
        Resume Next
    Case 5941
        Resume Next
    Case Else
        MsgBox Err.Number & Err.Description, vbExclamation
End Select
End Sub

Public Sub 比對選取文字()  '2003/4/4(不包括任一符號)
'指定鍵: Atl+Ctrl+Shift+Up(↑)
CheckSavedNoClear
'字元表：between -24667 and 19968
'Selection = ChrW(字元表)
Dim r As String, ins(4) As Long, f, i As Long, rCompMain As String, rCompFootnote As String, R1 As String
f = Array("。", "」", Chr(-24152), "：", "，", "；", _
    "、", "「", ".", Chr(34), ":", ",", ";", _
            "……", "...", "）", ")", "-", "．", "『", "』" _
            , "《", "》", "〉", "〈", "（", "）", "--", _
            ChrW(8212), "－", "？", ChrW(2), Chr(13), Chr(10), Chr(8), Chr(9), _
            "　", " ")
            'ChrW (2)為註腳符號
With Selection '指定鍵：Alt+Ctrl+Up
'If Not .Text Like "" Then '快速鍵：Alt+Ctrl+Down
    If .Type = wdSelectionIP Then MsgBox "請選取想要尋找之文字", vbExclamation: Exit Sub
    If .Type = wdSelectionNormal Then ' <> wdNoSelection OR wdSelectionIP Then '不為插入點
        
        r = .text
        For i = 0 To UBound(f)
            If InStr(r, f(i)) Then
                r = Replace(r, f(i), "")
            End If
        Next i
'        Debug.Print r
        On Error GoTo Previews
'        If .Application.ActiveWindow.Next.Document.Name = .Document.Name Then _
'            .Application.ActiveWindow.Next.Document.Activate
        With .Application.ActiveWindow.Next.Document
            If .Windows.Count > 1 Then
                .ActiveWindow.Activate
            Else
                .Activate
            End If
        End With
        With .Application.ActiveWindow.Document ' ActiveDocument
            rCompMain = .StoryRanges(wdMainTextStory)
            If .Footnotes.Count > 0 Then _
                rCompFootnote = .StoryRanges(wdFootnotesStory)
            For i = 0 To UBound(f)
                If InStr(rCompMain, f(i)) Then
                    rCompMain = Replace(rCompMain, f(i), "")
                End If
                If Not rCompFootnote Like "" And InStr(rCompFootnote, f(i)) Then _
                    rCompFootnote = Replace(rCompFootnote, f(i), "")
            Next i
            f = Empty '釋放記憶體
            ins(1) = InStr(rCompMain, r)
            ins(2) = InStrRev(rCompMain, r)
            ins(3) = InStr(rCompFootnote, r)
            ins(4) = InStrRev(rCompFootnote, r)
            If ins(1) = 0 And ins(3) = 0 Then _
                MsgBox "沒有符合文字", vbExclamation: _
                    .Application.ActiveWindow.Previous.Activate: End
            If ins(1) = ins(2) And ins(3) = ins(4) Then _
                MsgBox "本文只有此處!", vbInformation ': Exit Sub
            If ins(1) <> 0 Then
                ins(1) = wdMainTextStory
            Else
                ins(1) = wdFootnotesStory
            End If
'            rCompFootnote = Empty '用不著的字串變數歸零
            rCompMain = r '重新使用字串變數
            For i = 1 To Len(rCompMain)
                If InStr(.StoryRanges(ins(1)), Left(rCompMain, i)) = 0 Then Exit For
            Next i
            rCompFootnote = rCompMain '重新使用字串變數
            rCompMain = Left(rCompMain, i - 1)
            For i = 1 To Len(rCompFootnote)
                If InStrRev(.StoryRanges(ins(1)), right(rCompFootnote, i), -1, vbTextCompare) = 0 Then Exit For
            Next i
            rCompFootnote = right(rCompFootnote, i - 1)
            R1 = rCompMain
            Beep
            ins(2) = 1: ins(3) = 0 '重新使用變數
            Do While ins(2) > ins(3)
                ins(2) = InStr(ins(2) + Len(rCompMain), .StoryRanges(ins(1)), rCompMain, vbTextCompare) - 1
                ins(3) = InStrRev(.StoryRanges(ins(1)), rCompFootnote, ins(3) - 1, vbTextCompare) - 1 + Len(rCompFootnote)
            Loop
            Selection.SetRange ins(2), ins(3)
            .ActiveWindow.ScrollIntoView Selection.Range, True
            .ActiveWindow.ScrollIntoView Selection.Range, False
            Beep
            For i = ins(3) To ins(2) Step -1
'                Application.System
                If rCompMain Like Selection.Range Then
                    Exit For
                Else
                    rCompMain = Selection.Range
                End If
'                If Right(rCompMain, Len(r)) Like "雅，則" Then Stop
                If Len(rCompMain) >= Len(r) And _
                InStrRev(rCompMain, rCompFootnote) = 0 Then
'                    ins(3) = ins(3) + 1 '當縮短一單位長度沒有時，則復原原長度（表示最後符合者，即減一長度前的字串矣）
                    ins(3) = ins(3) + Len(rCompFootnote) '當縮短一單位長度沒有時，則復原原長度（表示最後符合者，即減一長度前的字串矣）
                    Exit For
                End If
'                ins(3) = ins(3) - 1 '縮短一單位長度再找
                If InStr(right(rCompMain, Len(rCompMain) - Len(R1)), R1) > 0 Then
                    ins(2) = InStr(right(rCompMain, Len(rCompMain) - Len(R1)), R1) - 1 + Len(R1) + ins(2) '縮短一單位長度再找
'                Else
'                    Exit For
                End If
                Selection.SetRange start:=ins(2), End:=ins(3)
                .ActiveWindow.ScrollIntoView Selection.Range, True '顯示選取範圍
                If InStrRev(Left(rCompMain, Len(rCompMain) - Len(rCompFootnote)), rCompFootnote, -1, vbTextCompare) > 0 Then
                    ins(3) = InStrRev(Left(rCompMain, Len(rCompMain) - Len(rCompFootnote)), rCompFootnote, -1, vbTextCompare) - 1 + Len(rCompFootnote) + ins(2) '縮短一單位長度再找
'                Else
'                    Exit For
                End If
                Selection.SetRange start:=ins(2), End:=ins(3)
'                .StoryRanges(ins(1)).SetRange Start:=ins(2), End:=ins(3)
                .ActiveWindow.ScrollIntoView Selection.Range, False '顯示選取範圍
                
            Next i
            Beep
            '設定選取範圍
            Selection.SetRange ins(2), ins(3)
            .ActiveWindow.ScrollIntoView Selection.Range, False
'            Selection.SetRange InStr(.StoryRanges(ins(1)), rCompMain), _
'                InStrRev(.StoryRanges(ins(1)), rCompFootnote, -1, vbTextCompare)
'            With .StoryRanges(ins(1)).Find
'                .ClearFormatting
'                .Replacement.ClearFormatting '這也要清除才行
'                .Forward = True
'                .Wrap = wdFindAsk
'                .MatchCase = True
'                .Text = r
'                .Execute
'                .Parent.Select
'                With .Application.ActiveWindow
'                    If .WindowState = wdWindowStateMinimize Then .WindowState = wdWindowStateMaximize
'                    .ScrollIntoView Selection.Range, True '顯示選取範圍
'                End With
'            End With
        End With
    End If
End With
Exit Sub
Previews:
Select Case Err.Number
    Case 91
        On Error Resume Next
        Dim d As Byte
        d = Documents.Count
        If d > 1 Then
            If Documents(d - 1) <> ActiveDocument Then
                Documents(d - 1).Activate
            Else
                Documents(d).Activate
            End If
        End If
'        ActiveWindow.Previous.Document.ActiveWindow.Activate
        Resume Next
    Case 5941
        Resume Next
    Case Else
        MsgBox Err.Number & Err.Description, vbExclamation
End Select
End Sub

Public Sub 比對選取文字1() '2003/4/4(不包括任一符號)
'指定鍵: Atl+Ctrl+Shift+Up(↑)
CheckSavedNoClear
'字元表：between -24667 and 19968
'Selection = ChrW(字元表)
Dim r As String, ins(4) As Long, f
Dim rLeft As String, rRight As String, rComp As String, i As Long, rCompMain As String, rCompFootnote As String
f = Array("。", "」", Chr(-24152), "：", "，", "；", _
    "、", "「", ".", Chr(34), ":", ",", ";", _
            "……", "...", "）", ")", "-", "．", "『", "』" _
            , "《", "》", "〉", "〈", "（", "）", "--", _
            ChrW(8212), "－", "？", ChrW(2), Chr(13), Chr(10), Chr(8), Chr(9), _
            "　", " ")
            'ChrW (2)為註腳符號
With Selection '指定鍵：Alt+Ctrl+Up
'If Not .Text Like "" Then '快速鍵：Alt+Ctrl+Down
    If .Type = wdSelectionIP Then MsgBox "請選取想要尋找之文字", vbExclamation: Exit Sub
    If .Type = wdSelectionNormal Then ' <> wdNoSelection OR wdSelectionIP Then '不為插入點

        r = .text
        For i = 0 To UBound(f)
            If InStr(r, f(i)) Then
                r = Replace(r, f(i), "")
            End If
        Next i
'        Debug.Print r
        On Error GoTo Previews
        With .Application.ActiveWindow.Next.Document
            If .Windows.Count > 1 Then
                .ActiveWindow.Activate
            Else
                .Activate
            End If
        End With
        With .Application.ActiveWindow.Document ' ActiveDocument
            rCompMain = .StoryRanges(wdMainTextStory)
            If .Footnotes.Count > 0 Then _
                rCompFootnote = .StoryRanges(wdFootnotesStory)
            '取消符號：
            For i = 0 To UBound(f)
                If InStr(rCompMain, f(i)) Then
                    'rCompMain=沒有符號之正文
                    rCompMain = Replace(rCompMain, f(i), "")
                End If
                If Not rCompFootnote Like "" And InStr(rCompFootnote, f(i)) Then _
                    rCompFootnote = Replace(rCompFootnote, f(i), "")
                    'rCompFootnote=沒有符號的註腳
            Next i
            f = Empty '釋放記憶體
            ins(1) = InStr(rCompMain, r)
            ins(2) = InStrRev(rCompMain, r)
            ins(3) = InStr(rCompFootnote, r)
            ins(4) = InStrRev(rCompFootnote, r)
            If ins(1) = 0 And ins(3) = 0 Then _
                MsgBox "沒有符合文字", vbExclamation: _
                    .Application.ActiveWindow.Previous.Activate: End
            If ins(1) = ins(2) And ins(3) = ins(4) Then _
                MsgBox "本文只有此處!", vbInformation ': Exit Sub
            If ins(1) <> 0 Then
'                ins(1) = wdMainTextStory
                rComp = rCompMain '.StoryRanges(wdMainTextStory)
            Else
'                ins(1) = wdFootnotesStory
                rComp = rCompFootnote '.StoryRanges(wdFootnotesStory)
            End If
'            rCompFootnote = Empty '用不著的字串變數歸零
            '由左方取得第一個吻合正文的詞
            For i = 1 To Len(r)
                If InStr(rComp, Left(r, i)) = 0 Then Exit For
            Next i
            rLeft = Left(r, i - 1)
            '由右方取得第一個吻合正文的詞
            For i = 1 To Len(r)
                If InStrRev(rComp, right(r, i), -1, vbTextCompare) = 0 Then Exit For
            Next i
            rRight = right(r, i - 1)

            ins(2) = InStr(rComp, rLeft) '- 1
            ins(3) = InStrRev(rComp, rRight, -1, vbTextCompare) + Len(rRight) '- 1
'            Selection.SetRange ins(2), ins(3)
'            rComp = Mid(rComp, ins(2), ins(3) - ins(2))
            For i = ins(2) To ins(3) Step 1
'                Application.System
'                r = Selection.Range
'                If Right(r, Len(r)) Like "雅，則" Then Stop
                ins(2) = ins(2) + 1
                If InStr(right(rComp, Len(rComp) - Len(rRight)), rRight) > 0 Then _
                    ins(2) = InStr(right(rComp, Len(rComp) - Len(rRight)), rRight) - 1 + Len(rRight) + ins(2) '縮短一單位長度再找
                If Len(rComp) >= Len(r) Then
'                 InStrRev(rComp, rRight) = 0
'                    ins(3) = ins(3) + 1 '當縮短一單位長度沒有時，則復原原長度（表示最後符合者，即減一長度前的字串矣）
                    ins(3) = ins(3) + Len(rRight) '當縮短一單位長度沒有時，則復原原長度（表示最後符合者，即減一長度前的字串矣）
                    Exit For
                End If
'                ins(3) = ins(3) - 1 '縮短一單位長度再找
                If InStr(right(rComp, Len(rComp) - Len(rRight)), rRight) > 0 Then
                    ins(2) = InStr(right(rComp, Len(rComp) - Len(rRight)), rRight) - 1 + Len(rRight) + ins(2) '縮短一單位長度再找
'                Else
'                    Exit For
                End If
                Selection.SetRange start:=ins(2), End:=ins(3)
                .ActiveWindow.ScrollIntoView Selection.Range, True '顯示選取範圍
                If InStrRev(Left(rComp, Len(rComp) - Len(rRight)), rRight, -1, vbTextCompare) > 0 Then
                    ins(3) = InStrRev(Left(rComp, Len(rComp) - Len(rRight)), rRight, -1, vbTextCompare) - 1 + Len(rRight) + ins(2) '縮短一單位長度再找
'                Else
'                    Exit For
                End If
                Selection.SetRange start:=ins(2), End:=ins(3)
'                .StoryRanges(ins(1)).SetRange Start:=ins(2), End:=ins(3)
                .ActiveWindow.ScrollIntoView Selection.Range, False '顯示選取範圍
            Next i
            Beep
            '設定選取範圍
            Selection.SetRange ins(2), ins(3)
'            Selection.SetRange InStr(.StoryRanges(ins(1)), rcomp), _
'                InStrRev(.StoryRanges(ins(1)), rright, -1, vbTextCompare)
'            With .StoryRanges(ins(1)).Find
'                .ClearFormatting
'                .Replacement.ClearFormatting '這也要清除才行
'                .Forward = True
'                .Wrap = wdFindAsk
'                .MatchCase = True
'                .Text = r
'                .Execute
'                .Parent.Select
'                With .Application.ActiveWindow
'                    If .WindowState = wdWindowStateMinimize Then .WindowState = wdWindowStateMaximize
'                    .ScrollIntoView Selection.Range, True '顯示選取範圍
'                End With
'            End With
        End With
    End If
End With
Exit Sub
Previews:
Select Case Err.Number
    Case 91
        On Error Resume Next
        Dim d As Byte
        d = Documents.Count
        If d > 1 Then
            If Documents(d - 1) <> ActiveDocument Then
                Documents(d - 1).Activate
            Else
                Documents(d).Activate
            End If
        End If
'        ActiveWindow.Previous.Document.ActiveWindow.Activate
        Resume Next
    Case 5941
        Resume Next
    Case Else
        MsgBox Err.Number & Err.Description, vbExclamation
End Select
End Sub

Public Sub 瀏覽瀚典() '2003/4/6
With Selection
    If IsNumeric(.Range) Then
        Dim r As Integer
        r = CInt(.Range)
        If .Document.Range(.End, .End + 1) Like "." Then r = r + 1
        With .Find
            .ClearFormatting
'            .ClearAllFuzzyOptions
            .text = CStr(r)
            .Execute Forward:=True, Wrap:=wdFindContinue ', Wrap:=wdFindAsk
        End With
        If .Range = r And Not .Document.Range(.start - 1, .start) Like Chr(13) _
            And .Document.Range(.End, .End + 1) Like "." Then
            .Range = Chr(13) & r
            .MoveRight
            '要轉成字串，所得長度方為字串長度，數字者，Len()則得半長爾
            .SetRange .start, End:=.start + Len(CStr(r))
        End If
    End If
End With
End Sub

Sub 瀏覽瀚典_自動() '2003/4/6
With Selection
    If IsNumeric(.Range) Then
        Dim r As Integer, C As Integer, p1 As Long, p2 As Long
        Do
            '如果倒著找或原地找，則表示找完了，須+1再繼續找。因為Find設定為Wrap:=wdFindContinue
            If p1 >= .start Then
                r = r + 1
            Else
                r = CInt(.Range)
            End If
            If .Document.Range(.End, .End + 1) Like "." Then r = r + 1
            p1 = .start '記下尋找下一個時的位置
            With .Find
                .ClearFormatting
                .ClearAllFuzzyOptions
                .text = r
                .Execute Forward:=True, Wrap:=wdFindContinue ' Wrap:=wdFindAsk
            End With
            If .Document.Range(.start - 1, .start) Like Chr(13) Or p2 >= p1 Then
                MsgBox "已完成" & C & "次替換！", vbInformation
                Exit Do
            End If
            If .Range = r And .Document.Range(.End, .End + 1) Like "." Then
                .Range = Chr(13) & r
                C = C + 1
                p2 = .start
                .MoveRight
            '要轉成字串，所得長度方為字串長度，數字者，Len()則得半長爾
                .SetRange .start, End:=.start + Len(CStr(r))
            End If
        Loop
    End If
End With
End Sub

Sub 瀏覽瀚典_檢查() '2003/4/6
With Selection
    If IsNumeric(.Range) Then
        Dim r As Integer, R1 As Integer, C As Integer, p1 As Long
        Do
            If p1 >= .start Then GoTo Out ' Exit Do
            r = CInt(.Range)
            R1 = CInt(.Range)
            If .start = 0 Then .MoveRight
            r = r + 1
            p1 = .start
'            .MoveRight
            With .Find
                .ClearFormatting
                .ClearAllFuzzyOptions
                .text = r 'CStr(r)
                .Execute Forward:=True, Wrap:=wdFindContinue       ', Wrap:=wdFindContinue ' Wrap:=wdFindAsk
            End With
            C = C + 1
            If R1 = CInt(.Range) And _
                (Not .Document.Range(.start - 1, .start) Like Chr(10) _
                    Or Not .Document.Range(.start - 1, .start) Like Chr(13)) _
                    And .Document.Range(.End, .End + 1) Like "." Then
Out:            MsgBox "已檢查" & C & "次！", vbExclamation
                Exit Do
            End If
        Loop
    End If
End With
End Sub

Sub 清除文句中斷行() '2003/4/7
Dim a As String, b As String
Dim C As Integer, p As Integer, d As Long, StepByStep As Byte
Const NoArrange = 255
StepByStep = MsgBox("要逐處檢視嗎？", vbYesNoCancel + vbDefaultButton2 + vbQuestion)
If StepByStep = vbCancel Then End
With Selection
    d = Len(.Document.Content)
    If .End >= d Then
        If MsgBox("要從頭開始嗎？", vbQuestion + vbOKCancel) = vbOK Then
           .HomeKey wdStory, wdMove
            If .text Like Chr(13) Then .Delete
        Else
            Exit Sub
        End If
    End If
    If .Type <> wdSelectionIP Then .MoveRight '有選取範圍時會取代掉選取範圍, 故須先檢查!
    Do
        .Find.ClearFormatting
        .Find.Execute FindText:="^p", Forward:=True
        C = C + 1
        a = .Range.Previous '此法較快
'        a = .Document.Range(.Start - 1, .Start)
'        If .End + 1 > Len(.Document.Content) Then Exit Do
        If .End + 1 >= d Then Exit Do
        b = .Range.Next
'        b = .Document.Range(.End, .End + 1)
        If InStr(.Paragraphs(1).Range, "勘記") Then
            Stop
'            If StepByStep = NoArrange Then StepByStep = vbNo'校勘記完後(此廿四史格式不同,暫緩)
            StepByStep = NoArrange '校勘記的格式不同
        End If
        If (Not Asc(a) = 13 And Not Asc(a) = 10 And Not IsNumeric(a) _
                    And Not a Like "-") And Not Asc(a) = 46 _
            And (Not Asc(b) = 13 And Not Asc(b) = 10 _
                And Not Asc(b) = 91 And Not Asc(b) = 93 _
                    And Not IsNumeric(b) _
               And Not b Like "　" And Not b Like "-" And Not b Like "〔") And Not b Like "【" Then
            .Document.ActiveWindow.LargeScroll 1, 0, 0, 0
            .Document.ActiveWindow.ScrollIntoView Selection.Range, True
            If StepByStep = vbNo Then
'                If Len(.Paragraphs(1).Range) > 28 And _
                    Left(VBA.Right(.Paragraphs(1).Range, 3), 1) <> "。" And _
                    Left(VBA.Right(.Paragraphs(1).Range, 4), 2) <> "。」" Then      '如此多字數而段行者才處理,否則太繁複了!(連標題等也算入,就太瑣碎了!)2003/11/30
               '以樂府詩集之格式等，暫加如此IF條件式！2003/11/30
'                If Len(.Paragraphs(1).Range) < 40 Then
'                    If MsgBox("要清除嗎?", vbQuestion + vbOKCancel) = vbOK Then .Range = ""
'                Else
                   .Range = ""
'                End If
                p = p + 1
'                End If
            Else '>28,以樂府詩集之格式等，暫加如此IF條件式！2003/11/30
                If Len(.Paragraphs(1).Range) > 36 And Left(VBA.right(.Paragraphs(1).Range, 3), 1) <> "。" _
                        And Left(VBA.right(.Paragraphs(1).Range, 3), 1) <> "﹂" _
                        And Left(VBA.right(.Paragraphs(1).Range, 3), 1) <> "」" _
                        Then '如此多字數而段行者才處理,否則太繁複了!(連標題等也算入,就太瑣碎了!)2003/11/30
                    Select Case MsgBox("要清除嗎?" & vbCr & vbCr & "要終止請按〔否〕！" _
                        , vbYesNoCancel + vbQuestion)
                        Case vbYes  '2003/4/17
                            .Range = ""
                            p = p + 1
                        Case vbCancel
                            If StepByStep = vbYes Then
                                .Range = Chr(13) & .Range '插入新段落以區隔開來!
                                p = p + 1
                            ElseIf StepByStep = NoArrange Then '不處理
                            Else
                                Stop
                            End If
                        Case vbNo
                            End
                    End Select
                End If
            End If
'            p = p + 1
            d = d - 2 '消除換行符號(chr(13)會併復位符號(Chr(10)也取消掉,故須減二
        End If
    Loop
    MsgBox "完成" & C & "次檢查，" & p & "次置換！", vbInformation
End With
End Sub

Sub 瀏覽瀚典_清除標題行()
Dim a As String
With Selection
If .Document.path <> "" Then MsgBox "此文件不能操作", vbExclamation: Exit Sub
    a = .text
    With .Find
        .text = a
        .Parent.Paragraphs(1).Range.Delete
        .Forward = True
        Do While .Execute
            .Parent.Paragraphs(1).Range.Delete
        Loop
    End With
End With
End Sub
Sub 瀏覽瀚典_清除頁碼()
Dim a As Paragraph
With Selection
    If .Document.path <> "" Then MsgBox "此文件不能操作", vbExclamation: Exit Sub
    For Each a In .Document.Paragraphs
        If Len(a.Range) > 2 Then
            If IsNumeric(Mid(a.Range, 2, Len(a.Range) - 4)) Then
                a.Range.Delete
            End If
        End If
    Next a
End With
End Sub

Sub a() '宋明理學期末作業用2004/5/2
Attribute a.VB_Description = "巨集錄製於 2004/1/13，錄製者 孫守真"
Attribute a.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.巨集1"
'
' 巨集1 巨集
' 巨集錄製於 2004/1/13，錄製者 孫守真
'
Dim i As Long
If MsgBox("請先檢查出處是否已獨立成段落！", vbExclamation + vbOKCancel) = vbOK Then Exit Sub
With Selection
    If .Type = wdSelectionNormal Then .move wdLine, -1
    For i = 1 To .Document.Paragraphs.Count
        Select Case .Paragraphs(1).Range.Font.Name
            Case "新細明體"
1               .MoveDown wdParagraph, 1, wdExtend
                If IsNumeric(Left(LTrim(.Paragraphs(1).Range), 1)) And (InStr(1, .Sections(1).Range, _
                    Left(LTrim(.Paragraphs(1).Range), 1), vbBinaryCompare) < InStrRev(.Sections(1).Range, _
                    Left(LTrim(.Paragraphs(1).Range), 1), , vbBinaryCompare)) Then
                    '要避免原文註腳被刪!
                    GoTo 2
                End If
                .Application.ScreenRefresh
'                Select Case MsgBox("確定刪除？", vbYesNoCancel + vbQuestion)
'                    Case vbYes
                        If InStr(.Paragraphs(1).Range, "---") Then If MsgBox("確定刪除註腳分隔線？", vbExclamation + vbOKCancel + vbDefaultButton2) = vbCancel Then GoTo 2
'                        If IsNumeric(.Paragraphs(1).Range.Words(1)) _
                            And (InStr(1, .Sections(1).Range, _
                            .Paragraphs(1).Range.Words(1), vbBinaryCompare) = InStrRev(.Sections(1).Range, _
                            .Paragraphs(1).Range.Words(1), , vbBinaryCompare)) Then
                        If IsNumeric(Left(LTrim(.Paragraphs(1).Range), 1)) _
                            And (InStr(1, .Sections(1).Range, _
                            Left(LTrim(.Paragraphs(1).Range), 1), vbBinaryCompare) = InStrRev(.Sections(1).Range, _
                            Left(LTrim(.Paragraphs(1).Range), 1), , vbBinaryCompare)) Then
                            '要原文註腳刪光後清除分隔線!(有時註腳編號前會空一格,故改上式）
                            .Paragraphs(1).Range.Delete
                            If Not IsNumeric(Left(LTrim(.Paragraphs(1).Range), 1)) And InStr(.Paragraphs(1).Previous.Range, "---") Then
                                .Paragraphs(1).Previous.Range.Delete
                            End If
                        Else
                            .Paragraphs(1).Range.Delete
                        End If
'                    Case vbNo
'                        .MoveDown wdParagraph, 1
'                    Case vbCancel
'                        Exit For
'                End Select
            Case "Times New Roman"
                If InStr(.Paragraphs(1).Range, "---") Then
                    GoTo 1
                Else
                    .MoveDown wdParagraph, 1
                End If
            Case Else
'            If d = 0 Then
2           .MoveDown wdParagraph, 1
        End Select
        If .End + 1 = .Document.Range.End Then MsgBox "恭喜完成!", vbInformation: Exit Sub
    Next i
End With
End Sub
Sub a1()
Selection.Range.Find.Execute "為", , , , , , , , , ChrW(29234), wdReplaceAll
End Sub

Sub 清除頁碼標記()
Dim p As Paragraph
For Each p In Documents(1).Paragraphs
    If IsNumeric(p.Range) Then
'        Select Case MsgBox("確定刪除？", vbYesNoCancel + vbQuestion)
'            Case vbYes
                p.Range.Select
                word.Application.ScreenRefresh
                p.Range.Delete
'            Case Else
'                Exit For
'        End Select
    End If
Next p
End Sub

Sub 貼上漢語大詞典()
Attribute 貼上漢語大詞典.VB_Description = "巨集錄製於 2005/2/24，錄製者 孫守真"
Attribute 貼上漢語大詞典.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.貼上漢語大詞典"
'
' 貼上漢語大詞典 巨集
' 巨集錄製於 2005/2/24，錄製者 孫守真
'
    Documents.Add DocumentType:=wdNewBlankDocument
    Selection.Paste
    ActiveDocument.SaveAs fileName:="漢語大詞典.htm", FileFormat:=wdFormatHTML, _
        LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword _
        :="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
        SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
        False
    ActiveWindow.View.Type = wdWebView
    ActiveWindow.Close
End Sub
Sub 校槁列印()
Attribute 校槁列印.VB_Description = "巨集錄製於 2005/4/13，錄製者 Oscar Sun"
Attribute 校槁列印.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.校槁列印"
'
' 校槁列印 巨集
' 巨集錄製於 2005/4/13，錄製者 Oscar Sun
'
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(1.54)
        .BottomMargin = CentimetersToPoints(1.54)
        .LeftMargin = CentimetersToPoints(1.17)
        .RightMargin = CentimetersToPoints(1.17)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1.5)
        .FooterDistance = CentimetersToPoints(1.75)
        .PageWidth = CentimetersToPoints(21)
        .PageHeight = CentimetersToPoints(29.7)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .GutterPos = wdGutterPosLeft
        .LayoutMode = wdLayoutModeLineGrid
    End With
    Selection.Sections(1).Footers(1).PageNumbers.Add PageNumberAlignment:= _
        wdAlignPageNumberCenter, FirstPage:=True
    With ActiveDocument.Styles("內文")
        .AutomaticallyUpdate = False
        .BaseStyle = ""
        .NextParagraphStyle = "內文"
    End With
    With ActiveDocument.Styles("內文").Font
        .NameFarEast = "標楷體"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Name = "Times New Roman"
        .Size = 12
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 1
        .Animation = wdAnimationNone
        .DisableCharacterSpaceGrid = False
        .EmphasisMark = wdEmphasisMarkNone
    End With
    With ActiveDocument.Styles("內文").ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 14
        .Alignment = wdAlignParagraphLeft
        .WidowControl = False
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .AutoAdjustRightIndent = True
        .DisableLineHeightGrid = False
        .FarEastLineBreakControl = True
        .WordWrap = True
        .HangingPunctuation = True
        .HalfWidthPunctuationOnTopOfLine = False
        .AddSpaceBetweenFarEastAndAlpha = True
        .AddSpaceBetweenFarEastAndDigit = True
        .BaseLineAlignment = wdBaselineAlignAuto
    End With
    Selection.Style = ActiveDocument.Styles("內文")
    word.Application.PrintOut fileName:="", Range:=wdPrintAllDocument, item:= _
        wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
        ManualDuplexPrint:=False, Collate:=True, Background:=False, PrintToFile:= _
        False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
        PrintZoomPaperHeight:=0
End Sub



Sub 巨集1()
Attribute 巨集1.VB_Description = "巨集錄製於 2008/12/24，錄製者 Oscar Sun"
Attribute 巨集1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.巨集1"
'
' 巨集1 巨集
' 巨集錄製於 2008/12/24，錄製者 Oscar Sun
'
    ActiveDocument.SaveAs fileName:= _
        "復初齋詩集（一）(588頁)-卷25(秘閣直廬集上（壬寅三月至十二月）壬寅 乾隆47年.1782年.先生年50歲).html", _
        FileFormat:=wdFormatText, LockComments:=False, Password:="", _
        AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
        EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
        :=False, SaveAsAOCELetter:=False
End Sub

Sub 巨集2()
Attribute 巨集2.VB_Description = "巨集錄製於 2010/10/28，錄製者 Oscar Sun"
Attribute 巨集2.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.巨集2"
'
' 巨集2 巨集
' 巨集錄製於 2010/10/28，錄製者 Oscar Sun
'
    With Selection.ParagraphFormat
        .RightIndent = CentimetersToPoints(8.74)
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
    End With
End Sub
Sub 巨集3()
Attribute 巨集3.VB_Description = "巨集錄製於 2010/10/28，錄製者 Oscar Sun"
Attribute 巨集3.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.巨集3"
'
' 巨集3 巨集
' 巨集錄製於 2010/10/28，錄製者 Oscar Sun
'
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(2.54)
        .BottomMargin = CentimetersToPoints(2.54)
        .LeftMargin = CentimetersToPoints(3.17)
        .RightMargin = CentimetersToPoints(12.17)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1.5)
        .FooterDistance = CentimetersToPoints(1.75)
        .PageWidth = CentimetersToPoints(21)
        .PageHeight = CentimetersToPoints(29.7)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .GutterPos = wdGutterPosLeft
        .LayoutMode = wdLayoutModeLineGrid
    End With
End Sub
Sub 巨集4()
Attribute 巨集4.VB_Description = "巨集錄製於 2010/10/28，錄製者 Oscar Sun"
Attribute 巨集4.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.巨集4"
'
' 巨集4 巨集
' 巨集錄製於 2010/10/28，錄製者 Oscar Sun
'
End Sub
Sub 巨集5()
Attribute 巨集5.VB_Description = "巨集錄製於 2010/10/28，錄製者 Oscar Sun"
Attribute 巨集5.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.巨集5"
'
' 巨集5 巨集
' 巨集錄製於 2010/10/28，錄製者 Oscar Sun
'
    ActiveWindow.ActivePane.View.Zoom.Percentage = 75
End Sub
    



Sub 巨集6()
Attribute 巨集6.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.巨集6"
'
' 巨集6 巨集
'
'
    Selection.MoveLeft unit:=wdCharacter, Count:=1, Extend:=wdExtend
    ActiveDocument.DefaultTargetFrame = ""
    Selection.Range.Hyperlinks(1).Range.Fields(1).result.Select
    Selection.Range.Hyperlinks(1).Delete
    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:= _
        "https://oscarsun72.blogspot.com/2021/02/blog-post_25.html", SubAddress:= _
        "", ScreenTip:="", TextToDisplay:="", Target:="_blank"
    Selection.Collapse Direction:=wdCollapseEnd
End Sub
Sub 巨集7()
'
' 巨集7 巨集
'
'
    'ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:="", _
        SubAddress:="_彌留之際", ScreenTip:="", TextToDisplay:=Selection '"他傳令"
    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:="", _
        SubAddress:="_鳴金收兵", ScreenTip:="", TextToDisplay:=Selection '"他傳令"
    'ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:="", _
        SubAddress:="_" & Selection, ScreenTip:="", TextToDisplay:=Selection '"他傳令"

End Sub
Sub 巨集8()
Attribute 巨集8.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.巨集8"
'
' 巨集8 巨集
'
'
'    Selection.Range.Hyperlinks(1).Range.Fields(1).Result.Select
'    Selection.Range.Hyperlinks(1).Delete
'     ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:="", _
        SubAddress:="_鳴金收兵（ㄇ" & ChrW(20008) & "ㄥˊ_ㄐ" & ChrW(20008) & "ㄣ_ㄕㄡ"
    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:="", _
        SubAddress:="鳴金收兵（ㄇ" & ChrW(20008) & "ㄥˊ_ㄐ" & ChrW(20008) & "ㄣ_ㄕㄡ_ㄅ" & ChrW(20008) & "ㄥ）"
        
'    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:="", _
        SubAddress:="鳴金收兵"
    
End Sub
Sub 巨集9()
Attribute 巨集9.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.巨集9"
'
' 巨集9 巨集
'
'
    Selection.InsertCrossReference ReferenceType:="標題", ReferenceKind:= _
        wdContentText, ReferenceItem:="241", InsertAsHyperlink:=True, _
        IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
End Sub
Sub 巨集10()
Attribute 巨集10.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.巨集10"
'
' 巨集10 巨集
'
'
    Selection.InsertCrossReference ReferenceType:="標題", ReferenceKind:= _
        wdContentText, ReferenceItem:="6", InsertAsHyperlink:=True, _
        IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
End Sub

Sub 貼上ut內容()
Attribute 貼上ut內容.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.巨集11"
    Dim tb As Table, s As Long, ur As UndoRecord, rng As Range
    SystemSetup.stopUndo ur, "貼上ut內容"
    word.Application.ScreenUpdating = False
    s = Selection.start
    If Selection.Type = wdSelectionIP Then
        With Selection.Document
            'If .path = "" Then .Range.Select
            If .path = "" Then Set rng = .Range
            Selection.Paste
            '.Range(s, Selection.End).Select
            Set rng = .Range(s, Selection.End)
        End With
    End If
    'For Each tb In Selection.Document.Tables
    For Each tb In rng.Tables
        tb.Rows.ConvertToText Separator:=wdSeparateByParagraphs, _
            NestedTables:=True
    Next tb
'    Selection.Find.ClearFormatting
'    Selection.Find.Replacement.ClearFormatting
    rng.Find.ClearFormatting
    rng.Find.Replacement.ClearFormatting
    'With Selection.Find
    With rng.Find
        .text = "^p^p"
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
'    Selection.Find.Execute Replace:=wdReplaceAll
'    Selection.Find.Execute Replace:=wdReplaceAll
    rng.Find.Execute Replace:=wdReplaceAll
    rng.Find.Execute Replace:=wdReplaceAll
'    Selection.Copy
    SystemSetup.contiUndo ur
    word.Application.ScreenUpdating = True
    rng.Document.ActiveWindow.ScrollIntoView rng, False
End Sub
