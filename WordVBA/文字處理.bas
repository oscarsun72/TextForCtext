Attribute VB_Name = "文字處理"
Option Explicit
Dim rst As Recordset, d As Object
Dim db As Database 'set db=CurrentDb _
只能在已開啟之Access中參照一次 , 二次以上的參照 _
,須以Set db = DBEngine.Workspaces(0).OpenDatabase _
    ("d:\千慮一得齋\書籍資料\詞頻.mdb")!的形式參照! _
    參考: _
    Dim dbsCurrent As Database, dbsContacts As Database'由 CurrentDb 的線上說明複製 _
    Set dbsCurrent = CurrentDb _
    Set dbsContacts = DBEngine.Workspaces(0).OpenDatabase("Contacts.mdb")


Function isNum(x As String) As Boolean
If Len(x) > 1 Then Exit Function
x = StrConv(x, vbNarrow)
If x Like "[0-9]" Then isNum = True
End Function
Function isLetter(x As String) As Boolean
If Len(x) > 1 Then Exit Function
x = StrConv(x, vbNarrow)
If x Like "[a-z]" Then isLetter = True
End Function

Sub 字頻() '2002/11/10要Sub才能在Word中執行!
On Error GoTo 錯誤處理
Dim ch, wrong As Long
'Dim chct As Long
Dim StTime As Date, EndTime As Date
'Dim x As Long, firstword As String '亂碼檢查!2002/11/13
Options.SaveInterval = 0 '取消自動儲存
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\千慮一得齋\書籍資料\詞頻.mdb", False
d.docmd.SelectObject blog.myaccess.acTable, "字頻表", True
'd.Visible = True '檢查用
Set db = d.CurrentDb '一定要加〝d〞!!寫成以下亦可!
'以上可併成下二式即可!但不會顯示在營幕上,只能作幕後計算用!(見OpenCurrentDatabase的線上說明)
'Set db = d.DBEngine.OpenDatabase("d:\千慮一得齋\書籍資料\詞頻.mdb")
'Set db = d.DBEngine.Workspaces(0).OpenDatabase("d:\千慮一得齋\書籍資料\詞頻.mdb")
Set rst = db.OpenRecordset("字頻表", dbOpenDynaset)
If rst.RecordCount > 0 Then '要獲得全部的筆數須用MoveLast但此只需判斷有沒有原先的記錄即可!
'rst打開以後只會取得第一筆記錄!
'    db.Execute "DELETE 字頻表.* FROM 字頻表"
    db.Execute "DELETE * FROM 字頻表"
End If
StTime = Time
With ActiveDocument
    For Each ch In .Characters '有亂碼字時ch會傳回"?"變成了運算用符號
        wrong = wrong + 1 '檢視用!
'        If wrong = 373 Then MsgBox "Check!!" '檢查用!
        If wrong Mod 27250 = 0 Then 'If wrong Mod 29688 = 0 Then '到29688時會產生OLE沒有回應的錯誤,故在此歇會兒
            MsgBox "因系統負荷達到極限,請務必切換至Access打開資料表後關閉,再回來按下確定按鈕繼續!!" _
                , vbExclamation, "★系統重要資訊★"
'        ElseIf wrong = 49761 Then
'            MsgBox "請檢查!!"
        End If
'        If wrong Mod 1000 = 0 Then Debug.Print wrong
'        Debug.Print ch & vbCr & "--------"
        '換行字元、復位字元不計!
'        If Right(ch, 1) <> Chr(10) Or Left(ch, 1) <> Chr(13) Then
        Select Case Asc(ch)
            Case Is <> 13, 10
        With rst
11          .FindFirst "字彙 like '" & ch & "'"
12          If .NoMatch Then
                .AddNew
                rst("字彙") = ch
                rst("次數") = 1
                rst("Asc") = Asc(ch)
                rst("AscW") = AscW(ch)
    '            On Error GoTo 次數
                .Update
            Else '當有亂碼字時,會成為比較運算元"?"(Asc(ch)=63),則可能在文件中第一次出現的字會誤增次數
                '此外如"鶴"字等(在Word中插入→符號內最後一些)字,亦會與同形字同字元碼(Asc), _
                但在符號表中卻有不同位置,代表不同字!在統計時,系統亦會誤算在一起! _
                這點還須要克服!2002/11/13測試時,有時又會分開!(但Asc則相同!)
'                If .AbsolutePosition < 1 And ch Like "?" And Not rst("字彙") = "?" Then
'                    'If x = 1 Then MsgBox "有亂碼字,次數將加入第一個出現的字中!!"
'                    MsgBox "有亂碼字,次數將加入第一個出現的字中!!"
'                    AppActivate "Microsoft Word"
'                    Selection.Collapse
'                    Selection.SetRange wrong + ActiveDocument.Paragraphs.Count / 2, wrong + 1 '將該亂碼字選取
'                    x = x + 1
'                End If
                .Edit
                rst("次數") = rst("次數") + 1
                .Update
            End If
        End With
        End Select
'        chct = .Characters.Count
'        chct = Selection.StoryLength
'        instr(1+
'        .Select
retry:  Next ch
'    rst.Requery
'    rst.MoveFirst
'    If x > 0 Then
'        firstword = "◎◎亂碼字加入第一字:「" & rst("字彙") & "」中共有" & x & "次!!"
'    Else
'        firstword = "★放心吧!亂碼字亦統計正確!!★"
'    End If
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "統計完成!!" & vbCr & "(※共執行了" & wrong & "次的檢查※)" _
        & vbCr & "※耗時:" & Format(EndTime - StTime, "n分s秒") & "※" _
        & vbCr & "字元數=" & .Characters.Count & vbCr '_
'        & firstword
'    MsgBox "統計完成!!" & vbCr & "(※共執行了" & wrong & "次的檢查※)" _
'        & vbCr & "※耗時:" & DateDiff("n", StTime, EndTime) & "分鐘※" _
'        & vbCr & "字元數=" & .Characters.Count
End With
If MsgBox("要即刻檢視結果嗎?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\千慮一得齋\書籍資料\詞頻.mdb")
    AppActivate "Microsoft access"
End If
d.docmd.OpenTable "字頻表", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '恢復自動儲存
End '用Exit Sub無法每次關閉Access
錯誤處理:
Select Case Err.Number
    Case Is = 91, 3078 '參照不到DataBase內物件時
        MsgBox "請再按一次!", vbCritical
        'access.Application.Quit
        d.Quit
        End
'        d.CurrentDb.Close
'        Set db = DBEngine.Workspaces(0).OpenDatabase("d:\千慮一得齋\書籍資料\詞頻.mdb")
''        Debug.Print Err.Description '檢查用!
'        Resume
'    Case Is = 3163 '換行字元、復位字元不計!
'        If Right(ch, 1) = Chr(10) Then
'            ch = Left(ch, Len(ch) - 1)
'        ElseIf Left(ch, 1) = Chr(13) Then
'            ch = Right(ch, Len(ch) - 1) '或If Asc(ch)=13
'        End If
'        Resume 11
    Case Is = 93 '為[]等運算式特殊字元所設之比較式
        rst.FindFirst "asc(字彙) = " & Asc(ch)
        Resume 12
'    Case Is = -2147023170
'        MsgBox Err.Number & ":" & Err.Description
'        MsgBox Err.LastDllError & "." & Err.Source
'        Set d = CreateObject("access.application")
'        d.OpenCurrentDatabase "d:\千慮一得齋\書籍資料\詞頻.mdb", False
'        d.UserControl = True
'        Resume
'    Case Is = 462 '"遠端伺服器不存在或無法使用"
'        'd.OpenCurrentDatabase "d:\千慮一得齋\書籍資料\詞頻.mdb", False
''        Set d = GetObject("d:\千慮一得齋\書籍資料\詞頻.mdb")
'        Set db = d.CurrentDb
'        Set rst = db.OpenRecordset("詞頻表", dbOpenDynaset)
'        Resume
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub 詞頻() '2002/11/10
On Error GoTo 錯誤處理
Dim Wd, wrong As Long
Dim wrongmark As Integer ', wdct As Long
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '取消自動儲存
Set d = CreateObject("access.application")
d.UserControl = True '如果為False則db.close會關閉資料庫!
'd.UserControl = False
d.OpenCurrentDatabase "d:\千慮一得齋\書籍資料\詞頻.mdb", False
d.docmd.SelectObject d.acTable, "詞頻表", True
'd.Visible = True '檢查用UserControl=True則有此反會致誤!
Set db = d.CurrentDb
Set rst = db.OpenRecordset("詞頻表", dbOpenDynaset)
If rst.RecordCount > 0 Then db.Execute "DELETE * FROM 詞頻表"
StTime = Time
With ActiveDocument
    For Each Wd In .words
        wrong = wrong + 1 '檢視用!
'        If wrong Mod 1000 = 0 Then Debug.Print wrong
'        Debug.Print wd & vbCr & "--------"
        If Len(Wd) > 1 And Right(Wd, 1) = " " Then
            hfspace = hfspace + 1 '計次
            GoTo retry '字串右邊是半形空格時,AccessUpdate時會銷去,且於詞彙亦無意意,故不計!
        End If
        rst.FindFirst "詞彙 like '" & Wd & "'"
        If rst.NoMatch Then
            rst.AddNew
            rst("詞彙") = Wd
'            On Error GoTo 次數
            rst.Update
        Else
            rst.Edit
            rst("次數") = rst("次數") + 1
            rst.Update
        End If
'        wrong = 1
'        wdct = .Words.Count
'        wdct = Selection.StoryLength
'        instr(1+
'        .Select
retry:  Next Wd
End With
EndTime = Time
AppActivate "Microsoft word"
MsgBox "統計完成!!" & vbCr & "(※共執行了" & wrong & "次的檢查※)" _
    & "詞彙右邊半形空格凡" & hfspace & "次,忽略不計!" _
    & vbCr & "※耗時:" & Format(EndTime - StTime, "n分s秒") & "※"
If MsgBox("要即刻檢視結果嗎?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\千慮一得齋\書籍資料\詞頻.mdb")
    AppActivate "Microsoft access"
End If
d.docmd.OpenTable "詞頻表", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '恢復自動儲存
End
錯誤處理:
Select Case Err.Number '主索引值重複
    Case Is = 91, 3078
        MsgBox "請再按一次!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
'次數:
'    wrongmark = Err.Number
''    Err.Description = wd
'    If wrongmark = 3022 Then '重複了
''        wrong = wrong + 1
''        rst.Seek "=", "詞彙"
'        rst.FindFirst "詞彙 like '" & wd & "'"
'        rst.Edit
'        rst("次數") = rst("次數") + 1
'        rst.Update
'        Resume retry
'    Else
'        MsgBox "有錯誤,請檢查!!" & Err.Description, vbExclamation
'    End If
End Sub
Sub 進階詞頻() '2002/11/10要Sub才能在Word中執行!'2005/4/21此法在跑大檔案時太沒效率了!!跑了3天3夜300頁的文件檔取1-3字詞跑不完!
On Error GoTo 錯誤處理
Dim wrong As Long, phra As Long, phras As String, phralh As Byte
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Dim length As Byte 'As String
Dim Dw As String, dwL As Long
length = InputBox("請指定分析詞彙之上限,最多五個字", , "5")
If length = "" Or Not IsNumeric(length) Then End
If CByte(length) < 1 Or CByte(length) > 5 Then End
Options.SaveInterval = 0 '取消自動儲存
StTime = Time
Set d = CreateObject("access.application")
'或Set d = CreateObject("Access.Application.9")
d.UserControl = True
d.OpenCurrentDatabase "d:\千慮一得齋\書籍資料\詞頻.mdb", False
d.docmd.SelectObject d.acTable, "詞頻表", True
'd.Visible = True '檢查用
Set db = d.CurrentDb
Set rst = db.OpenRecordset("詞頻表", dbOpenDynaset)
If rst.RecordCount > 0 Then '要獲得全部的筆數須用MoveLast但此只需判斷有沒有原先的記錄即可!
'rst打開以後只會取得第一筆記錄!
'    db.Execute "DELETE 字頻表.* FROM 字頻表"
    db.Execute "DELETE * FROM 詞頻表"
End If
'With ActiveDocument
With ActiveDocument
    Dw = .Content '文件內容
    dwL = Len(Dw) '文件長度
    .Close
End With
    For phralh = 1 To length 'CByte(length)
'    For phralh = 1 To 5 '暫定最長為5個字構成的詞(仍可改作變數)
        For phra = 1 To dwL '.Characters.Count
            Select Case phralh
                Case Is = 1
                    If Err.LastDllError <> 0 Then
                        MsgBox Err.LastDllError & ":" & Err.Description & "Err.Number:" & Err.Number
                        GoTo 錯誤處理
                    End If
'                    phras = .Characters(phra)'此法太慢!
                    phras = Mid(Dw, phra, 1)
                Case Is = 2
                    If Err.LastDllError <> 0 Then
                        MsgBox Err.LastDllError & ":" & Err.Description & "Err.Number:" & Err.Number
                        GoTo 錯誤處理
                    End If
'                    If phra + 1 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1)
                    If phra + 1 <= dwL Then phras = Mid(Dw, phra, 2)
                Case Is = 3
                    If Err.LastDllError <> 0 Then
                        MsgBox Err.LastDllError & ":" & Err.Description & "Err.Number:" & Err.Number
                        GoTo 錯誤處理
                    End If
'                    If phra + 2 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1) & _
                            .Characters(phra + 2)
                    If phra + 2 <= dwL Then phras = Mid(Dw, phra, 3)
                Case Is = 4
                    On Error GoTo 錯誤處理
                    If Err.LastDllError <> 0 Then
                        MsgBox Err.LastDllError & ":" & Err.Description & "Err.Number:" & Err.Number
                        GoTo 錯誤處理
                    End If
'                    If phra + 3 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1) & _
                            .Characters(phra + 2) & .Characters(phra + 3)
                    If phra + 3 <= dwL Then phras = Mid(Dw, phra, 3)
                Case Is = 5
                    On Error GoTo 錯誤處理
                    If Err.LastDllError <> 0 Then
                        MsgBox Err.LastDllError & ":" & Err.Description & "Err.Number:" & Err.Number
                        GoTo 錯誤處理
                    End If
'                    If phra + 4 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1) & _
                            .Characters(phra + 2) & .Characters(phra + 3) & _
                            .Characters(phra + 4)
                    If phra + 4 <= dwL Then phras = Mid(Dw, phra, 3)
            End Select
            If Len(phras) > 1 And Right(phras, 1) = " " Then
                hfspace = hfspace + 1 '計次
                GoTo 11 '字串右邊是半形空格時,AccessUpdate時會銷去,且於詞彙亦無意意,故不計!
            End If
            '直接進入下一個字串比對
            wrong = wrong + 1 '檢視用!
            If wrong Mod 29688 = 0 Then '到29688時會產生OLE沒有回應的錯誤,故在此歇會兒
                DoEvents 'MsgBox "請檢查!!"
    '        ElseIf wrong = 49761 Then
    '            MsgBox "請檢查!!"
            End If
'            if rst Set rst = CurrentDb.OpenRecordset("SELECT  詞頻表.* FROM 詞頻表 WHERE (((詞頻表.詞彙) like '" & phras & "'));")
            With rst
'                If .RecordCount = 0 Then
                .FindFirst "詞彙 like '" & phras & "'"
                If .NoMatch Then
'                    .MoveLast
                    .AddNew
                    rst("詞彙") = phras
'                    rst("次數") = 1'預設值已為1
                    On Error GoTo 錯誤處理
                    .Update 'dbUpdateBatch, True
                Else
1                   .Edit
                    rst("次數") = rst("次數") + 1
                    .Update
                End If
'                .Close
            End With
11      Next phra
2   Next phralh
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "統計完成!!" & vbCr & "(※共執行了" & wrong & "次的檢查※)" _
        & "詞彙右邊半形空格凡" & hfspace & "次,忽略不計!" _
        & vbCr & "※耗時:" & Format(EndTime - StTime, "n分s秒") & "※" _
        & vbCr & "字元數=" & dwL '.Characters.Count
'End With
'd.Visible = True
If MsgBox("要即刻檢視結果嗎?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\千慮一得齋\書籍資料\詞頻.mdb")
    AppActivate "Microsoft access"
End If
d.docmd.OpenTable "詞頻表", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '恢復自動儲存
End '用Exit Sub無法每次關閉Access'2002/11/15
錯誤處理:
Select Case Err.Number '主索引值重複
    Case Is = 3022
        rst.Requery
        rst.FindFirst "詞彙 like '" & Trim(phras) & "'"
        GoTo 1
    Case Is = 5941 '集合中的成員不存在(指超過文件長度!)
        GoTo 2
    Case Is = 91, 3078
        MsgBox "請再按一次!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub 進階詞頻1() '2002/11/15要Sub才能在Word中執行!
On Error GoTo 錯誤處理
Dim wrong As Long, phra As Long, phras As String, phralh As Byte
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Dim length As String
Dim i As Byte, j As Byte
length = InputBox("請指定分析詞彙之上限,最多255個字", , "5")
If length = "" Or Not IsNumeric(length) Then End
If CByte(length) < 1 Or CByte(length) > 255 Then End
Options.SaveInterval = 0 '取消自動儲存
StTime = Time
Set d = CreateObject("access.application")
'或Set d = CreateObject("Access.Application.9")
d.UserControl = True
d.OpenCurrentDatabase "d:\千慮一得齋\書籍資料\詞頻.mdb", False
d.docmd.SelectObject d.acTable, "詞頻表", True
'd.Visible = True '檢查用
Set db = d.CurrentDb
Set rst = db.OpenRecordset("詞頻表", dbOpenDynaset)
If rst.RecordCount > 0 Then '要獲得全部的筆數須用MoveLast但此只需判斷有沒有原先的記錄即可!
'rst打開以後只會取得第一筆記錄!
'    db.Execute "DELETE 字頻表.* FROM 字頻表"
    db.Execute "DELETE * FROM 詞頻表"
End If
j = CByte(length)
With ActiveDocument
    For phralh = 1 To j
'    原暫定最長為5個字構成的詞,今改作變數j,則限於Byte大小耳!
        For phra = 1 To .Characters.Count
            If phra + (phralh - 1) <= .Characters.Count Then
                phras = ""
                For i = 0 To phralh - 1
                    phras = phras & .Characters(phra + i)
                Next i
            End If
            If Len(phras) > 1 And Right(phras, 1) = " " Then
                hfspace = hfspace + 1 '計次
                GoTo 11 '字串右邊是半形空格時,AccessUpdate時會銷去,且於詞彙亦無意意,故不計!
            End If
            '直接進入下一個字串比對
            wrong = wrong + 1 '檢視用!
            If wrong Mod 29688 = 0 Then '到29688時會產生OLE沒有回應的錯誤,故在此歇會兒
                MsgBox "請檢查!!"
    '        ElseIf wrong = 49761 Then
    '            MsgBox "請檢查!!"
            End If
            With rst
                .FindFirst "詞彙 like '" & phras & "'"
                If .NoMatch Then
    '                .MoveLast
                    .AddNew
                    rst("詞彙") = phras
                    rst("次數") = 1
                    On Error GoTo 錯誤處理
                    .Update 'dbUpdateBatch, True
                Else
1                   .Edit
                    rst("次數") = rst("次數") + 1
                    .Update
                End If
            End With
11      Next phra
2   Next phralh
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "統計完成!!" & vbCr & "(※共執行了" & wrong & "次的檢查※)" _
        & "詞彙右邊半形空格凡" & hfspace & "次,忽略不計!" _
        & vbCr & "※耗時:" & Format(EndTime - StTime, "n分s秒") & "※" _
        & vbCr & "字元數=" & .Characters.Count
End With
'd.Visible = True
If MsgBox("要即刻檢視結果嗎?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\千慮一得齋\書籍資料\詞頻.mdb")
    AppActivate "Microsoft access"
End If
d.docmd.OpenTable "詞頻表", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '恢復自動儲存
End '用Exit Sub無法每次關閉Access
錯誤處理:
Select Case Err.Number '主索引值重複
    Case Is = 3022
        rst.Requery
        rst.FindFirst "詞彙 like '" & Trim(phras) & "'"
        GoTo 1
    Case Is = 5941 '集合中的成員不存在(指超過文件長度!)
        GoTo 2
    Case Is = 91, 3078
        MsgBox "請再按一次!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub 指定字數詞頻() '2002/11/11
On Error GoTo 錯誤處理
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
phralh = InputBox("請用阿拉伯數字指定詞的組成字數,最多字數為「11」!", "指定詞彙字數", "2")
If phralh = "" Or Not IsNumeric(phralh) Then Exit Sub
If CByte(phralh) > 11 Or CByte(phralh) < 1 Then Exit Sub
Options.SaveInterval = 0 '取消自動儲存
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\千慮一得齋\書籍資料\詞頻.mdb", False
d.docmd.SelectObject d.acTable, "詞頻表", True
'd.Visible = True '檢查用
Set db = d.CurrentDb
Set rst = db.OpenRecordset("詞頻表", dbOpenDynaset)
If rst.RecordCount > 0 Then '要獲得全部的筆數須用MoveLast但此只需判斷有沒有原先的記錄即可!
'rst打開以後只會取得第一筆記錄!
'    db.Execute "DELETE 字頻表.* FROM 字頻表"
    db.Execute "DELETE * FROM 詞頻表"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        Select Case CByte(phralh)
            Case Is = 1
                phras = .Characters(phra)
            Case Is = 2
                If phra + 1 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1)
            Case Is = 3
                If phra + 2 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2)
            Case Is = 4
                If phra + 3 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2) & .Characters(phra + 3)
            Case Is = 5
                If phra + 4 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2) & .Characters(phra + 3) & _
                        .Characters(phra + 4)
            Case Is = 6
                If phra + 5 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2) & .Characters(phra + 3) & _
                        .Characters(phra + 4) & .Characters(phra + 5)
            Case Is = 7
                If phra + 6 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2) & .Characters(phra + 3) & _
                        .Characters(phra + 4) & .Characters(phra + 5) & _
                        .Characters(phra + 6)
            Case Is = 8
                If phra + 7 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2) & .Characters(phra + 3) & _
                        .Characters(phra + 4) & .Characters(phra + 5) & _
                        .Characters(phra + 6) & .Characters(phra + 7)
            Case Is = 9
                If phra + 8 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2) & .Characters(phra + 3) & _
                        .Characters(phra + 4) & .Characters(phra + 5) & _
                        .Characters(phra + 6) & .Characters(phra + 7) & _
                        .Characters(phra + 8)
            Case Is = 10
                If phra + 9 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2) & .Characters(phra + 3) & _
                        .Characters(phra + 4) & .Characters(phra + 5) & _
                        .Characters(phra + 6) & .Characters(phra + 7) & _
                        .Characters(phra + 8) & .Characters(phra + 9)
            Case Is = 11
                If phra + 10 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2) & .Characters(phra + 3) & _
                        .Characters(phra + 4) & .Characters(phra + 5) & _
                        .Characters(phra + 6) & .Characters(phra + 7) & _
                        .Characters(phra + 8) & .Characters(phra + 9) & _
                        .Characters(phra + 10)
        End Select
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '計次
            GoTo 11 '字串右邊是半形空格時,AccessUpdate時會銷去,且於詞彙亦無意意,故不計!
        End If
        '直接進入下一個字串比對
        wrong = wrong + 1 '檢視用!
'        If wrong Mod 29688 = 0 Then '到29688時會產生OLE沒有回應的錯誤,故在此歇會兒
'            MsgBox "請檢查!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "請檢查!!"
'        End If
        With rst
            .FindFirst "詞彙 like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("詞彙") = phras
'                rst("次數") = 1'預設值已定為1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("次數") = rst("次數") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "統計完成!!" & vbCr & "(※共執行了" & wrong & "次的檢查※)" _
        & "詞彙右邊半形空格凡" & hfspace & "次,忽略不計!" _
        & vbCr & "※耗時:" & Format(EndTime - StTime, "n分s秒") & "※" _
        & vbCr & "字元數=" & .Characters.Count
End With
If MsgBox("要即刻檢視結果嗎?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\千慮一得齋\書籍資料\詞頻.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "詞頻表", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "詞頻表", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '恢復自動儲存
End '用Exit Sub無法每次關閉Access
錯誤處理:
Select Case Err.Number '主索引值重複
    Case Is = 91, 3078
        MsgBox "請再按一次!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub 指定11字數詞頻()     '2002/11/15'以此為例,可作為預先限定字數的各個程序(本例為11個字的查詢)
On Error GoTo 錯誤處理
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
'phralh = InputBox("請用阿拉伯數字指定詞的組成字數,最多字數為「11」!", "指定詞彙字數", "2")
'If phralh = "" Or Not IsNumeric(phralh) Then Exit Sub
'If CByte(phralh) > 11 Or CByte(phralh) < 1 Then Exit Sub
Options.SaveInterval = 0 '取消自動儲存
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\千慮一得齋\書籍資料\詞頻.mdb", False
d.docmd.SelectObject d.acTable, "詞頻表", True
'd.Visible = True '檢查用
Set db = d.CurrentDb
Set rst = db.OpenRecordset("詞頻表", dbOpenDynaset)
If rst.RecordCount > 0 Then '要獲得全部的筆數須用MoveLast但此只需判斷有沒有原先的記錄即可!
'rst打開以後只會取得第一筆記錄!
'    db.Execute "DELETE 字頻表.* FROM 字頻表"
    db.Execute "DELETE * FROM 詞頻表"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 10 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2) & .Characters(phra + 3) & _
                    .Characters(phra + 4) & .Characters(phra + 5) & _
                    .Characters(phra + 6) & .Characters(phra + 7) & _
                    .Characters(phra + 8) & .Characters(phra + 9) & _
                    .Characters(phra + 10)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '計次
            GoTo 11 '字串右邊是半形空格時,AccessUpdate時會銷去,且於詞彙亦無意意,故不計!
        End If
        '直接進入下一個字串比對
        wrong = wrong + 1 '檢視用!
'        If wrong Mod 29688 = 0 Then '到29688時會產生OLE沒有回應的錯誤,故在此歇會兒
'            MsgBox "請檢查!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "請檢查!!"
'        End If
        With rst
            .FindFirst "詞彙 like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("詞彙") = phras
'                rst("次數") = 1'預設值已定為1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("次數") = rst("次數") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "統計完成!!" & vbCr & "(※共執行了" & wrong & "次的檢查※)" _
        & "詞彙右邊半形空格凡" & hfspace & "次,忽略不計!" _
        & vbCr & "※耗時:" & Format(EndTime - StTime, "n分s秒") & "※" _
        & vbCr & "字元數=" & .Characters.Count
End With
If MsgBox("要即刻檢視結果嗎?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\千慮一得齋\書籍資料\詞頻.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "詞頻表", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "詞頻表", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '恢復自動儲存
End '用Exit Sub無法每次關閉Access
錯誤處理:
Select Case Err.Number '主索引值重複
    Case Is = 91, 3078
        MsgBox "請再按一次!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub 指定10字數詞頻() '2002/11/15
On Error GoTo 錯誤處理
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '取消自動儲存
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\千慮一得齋\書籍資料\詞頻.mdb", False
d.docmd.SelectObject d.acTable, "詞頻表", True
'd.Visible = True '檢查用
Set db = d.CurrentDb
Set rst = db.OpenRecordset("詞頻表", dbOpenDynaset)
If rst.RecordCount > 0 Then '要獲得全部的筆數須用MoveLast但此只需判斷有沒有原先的記錄即可!
'rst打開以後只會取得第一筆記錄!
'    db.Execute "DELETE 字頻表.* FROM 字頻表"
    db.Execute "DELETE * FROM 詞頻表"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 9 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2) & .Characters(phra + 3) & _
                    .Characters(phra + 4) & .Characters(phra + 5) & _
                    .Characters(phra + 6) & .Characters(phra + 7) & _
                    .Characters(phra + 8) & .Characters(phra + 9)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '計次
            GoTo 11 '字串右邊是半形空格時,AccessUpdate時會銷去,且於詞彙亦無意意,故不計!
        End If
        '直接進入下一個字串比對
        wrong = wrong + 1 '檢視用!
'        If wrong Mod 29688 = 0 Then '到29688時會產生OLE沒有回應的錯誤,故在此歇會兒
'            MsgBox "請檢查!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "請檢查!!"
'        End If
        With rst
            .FindFirst "詞彙 like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("詞彙") = phras
'                rst("次數") = 1'預設值已定為1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("次數") = rst("次數") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "統計完成!!" & vbCr & "(※共執行了" & wrong & "次的檢查※)" _
        & "詞彙右邊半形空格凡" & hfspace & "次,忽略不計!" _
        & vbCr & "※耗時:" & Format(EndTime - StTime, "n分s秒") & "※" _
        & vbCr & "字元數=" & .Characters.Count
End With
If MsgBox("要即刻檢視結果嗎?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\千慮一得齋\書籍資料\詞頻.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "詞頻表", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "詞頻表", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '恢復自動儲存
End '用Exit Sub無法每次關閉Access
錯誤處理:
Select Case Err.Number '主索引值重複
    Case Is = 91, 3078
        MsgBox "請再按一次!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub 指定9字數詞頻()  '2002/11/15
On Error GoTo 錯誤處理
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '取消自動儲存
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\千慮一得齋\書籍資料\詞頻.mdb", False
d.docmd.SelectObject d.acTable, "詞頻表", True
'd.Visible = True '檢查用
Set db = d.CurrentDb
Set rst = db.OpenRecordset("詞頻表", dbOpenDynaset)
If rst.RecordCount > 0 Then '要獲得全部的筆數須用MoveLast但此只需判斷有沒有原先的記錄即可!
'rst打開以後只會取得第一筆記錄!
'    db.Execute "DELETE 字頻表.* FROM 字頻表"
    db.Execute "DELETE * FROM 詞頻表"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 8 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2) & .Characters(phra + 3) & _
                    .Characters(phra + 4) & .Characters(phra + 5) & _
                    .Characters(phra + 6) & .Characters(phra + 7) & _
                    .Characters(phra + 8)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '計次
            GoTo 11 '字串右邊是半形空格時,AccessUpdate時會銷去,且於詞彙亦無意意,故不計!
        End If
        '直接進入下一個字串比對
        wrong = wrong + 1 '檢視用!
'        If wrong Mod 29688 = 0 Then '到29688時會產生OLE沒有回應的錯誤,故在此歇會兒
'            MsgBox "請檢查!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "請檢查!!"
'        End If
        With rst
            .FindFirst "詞彙 like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("詞彙") = phras
'                rst("次數") = 1'預設值已定為1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("次數") = rst("次數") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "統計完成!!" & vbCr & "(※共執行了" & wrong & "次的檢查※)" _
        & "詞彙右邊半形空格凡" & hfspace & "次,忽略不計!" _
        & vbCr & "※耗時:" & Format(EndTime - StTime, "n分s秒") & "※" _
        & vbCr & "字元數=" & .Characters.Count
End With
If MsgBox("要即刻檢視結果嗎?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\千慮一得齋\書籍資料\詞頻.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "詞頻表", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "詞頻表", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '恢復自動儲存
End '用Exit Sub無法每次關閉Access
錯誤處理:
Select Case Err.Number '主索引值重複
    Case Is = 91, 3078
        MsgBox "請再按一次!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub


Sub 指定8字數詞頻()   '2002/11/15
On Error GoTo 錯誤處理
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '取消自動儲存
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\千慮一得齋\書籍資料\詞頻.mdb", False
d.docmd.SelectObject d.acTable, "詞頻表", True
'd.Visible = True '檢查用
Set db = d.CurrentDb
Set rst = db.OpenRecordset("詞頻表", dbOpenDynaset)
If rst.RecordCount > 0 Then '要獲得全部的筆數須用MoveLast但此只需判斷有沒有原先的記錄即可!
'rst打開以後只會取得第一筆記錄!
'    db.Execute "DELETE 字頻表.* FROM 字頻表"
    db.Execute "DELETE * FROM 詞頻表"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 7 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2) & .Characters(phra + 3) & _
                    .Characters(phra + 4) & .Characters(phra + 5) & _
                    .Characters(phra + 6) & .Characters(phra + 7)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '計次
            GoTo 11 '字串右邊是半形空格時,AccessUpdate時會銷去,且於詞彙亦無意意,故不計!
        End If
        '直接進入下一個字串比對
        wrong = wrong + 1 '檢視用!
'        If wrong Mod 29688 = 0 Then '到29688時會產生OLE沒有回應的錯誤,故在此歇會兒
'            MsgBox "請檢查!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "請檢查!!"
'        End If
        With rst
            .FindFirst "詞彙 like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("詞彙") = phras
'                rst("次數") = 1'預設值已定為1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("次數") = rst("次數") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "統計完成!!" & vbCr & "(※共執行了" & wrong & "次的檢查※)" _
        & "詞彙右邊半形空格凡" & hfspace & "次,忽略不計!" _
        & vbCr & "※耗時:" & Format(EndTime - StTime, "n分s秒") & "※" _
        & vbCr & "字元數=" & .Characters.Count
End With
If MsgBox("要即刻檢視結果嗎?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\千慮一得齋\書籍資料\詞頻.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "詞頻表", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "詞頻表", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '恢復自動儲存
End '用Exit Sub無法每次關閉Access
錯誤處理:
Select Case Err.Number '主索引值重複
    Case Is = 91, 3078
        MsgBox "請再按一次!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub 指定6字數詞頻()    '2002/11/15
On Error GoTo 錯誤處理
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '取消自動儲存
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\千慮一得齋\書籍資料\詞頻.mdb", False
d.docmd.SelectObject d.acTable, "詞頻表", True
'd.Visible = True '檢查用
Set db = d.CurrentDb
Set rst = db.OpenRecordset("詞頻表", dbOpenDynaset)
If rst.RecordCount > 0 Then '要獲得全部的筆數須用MoveLast但此只需判斷有沒有原先的記錄即可!
'rst打開以後只會取得第一筆記錄!
'    db.Execute "DELETE 字頻表.* FROM 字頻表"
    db.Execute "DELETE * FROM 詞頻表"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 5 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2) & .Characters(phra + 3) & _
                    .Characters(phra + 4) & .Characters(phra + 5)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '計次
            GoTo 11 '字串右邊是半形空格時,AccessUpdate時會銷去,且於詞彙亦無意意,故不計!
        End If
        '直接進入下一個字串比對
        wrong = wrong + 1 '檢視用!
'        If wrong Mod 29688 = 0 Then '到29688時會產生OLE沒有回應的錯誤,故在此歇會兒
'            MsgBox "請檢查!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "請檢查!!"
'        End If
        With rst
            .FindFirst "詞彙 like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("詞彙") = phras
'                rst("次數") = 1'預設值已定為1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("次數") = rst("次數") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "統計完成!!" & vbCr & "(※共執行了" & wrong & "次的檢查※)" _
        & "詞彙右邊半形空格凡" & hfspace & "次,忽略不計!" _
        & vbCr & "※耗時:" & Format(EndTime - StTime, "n分s秒") & "※" _
        & vbCr & "字元數=" & .Characters.Count
End With
If MsgBox("要即刻檢視結果嗎?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\千慮一得齋\書籍資料\詞頻.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "詞頻表", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "詞頻表", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '恢復自動儲存
End '用Exit Sub無法每次關閉Access
錯誤處理:
Select Case Err.Number '主索引值重複
    Case Is = 91, 3078
        MsgBox "請再按一次!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub 指定5字數詞頻()     '2002/11/15
On Error GoTo 錯誤處理
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '取消自動儲存
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\千慮一得齋\書籍資料\詞頻.mdb", False
d.docmd.SelectObject d.acTable, "詞頻表", True
'd.Visible = True '檢查用
Set db = d.CurrentDb
Set rst = db.OpenRecordset("詞頻表", dbOpenDynaset)
If rst.RecordCount > 0 Then '要獲得全部的筆數須用MoveLast但此只需判斷有沒有原先的記錄即可!
'rst打開以後只會取得第一筆記錄!
'    db.Execute "DELETE 字頻表.* FROM 字頻表"
    db.Execute "DELETE * FROM 詞頻表"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 4 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2) & .Characters(phra + 3) & _
                    .Characters(phra + 4)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '計次
            GoTo 11 '字串右邊是半形空格時,AccessUpdate時會銷去,且於詞彙亦無意意,故不計!
        End If
        '直接進入下一個字串比對
        wrong = wrong + 1 '檢視用!
'        If wrong Mod 29688 = 0 Then '到29688時會產生OLE沒有回應的錯誤,故在此歇會兒
'            MsgBox "請檢查!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "請檢查!!"
'        End If
        With rst
            .FindFirst "詞彙 like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("詞彙") = phras
'                rst("次數") = 1'預設值已定為1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("次數") = rst("次數") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "統計完成!!" & vbCr & "(※共執行了" & wrong & "次的檢查※)" _
        & "詞彙右邊半形空格凡" & hfspace & "次,忽略不計!" _
        & vbCr & "※耗時:" & Format(EndTime - StTime, "n分s秒") & "※" _
        & vbCr & "字元數=" & .Characters.Count
End With
If MsgBox("要即刻檢視結果嗎?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\千慮一得齋\書籍資料\詞頻.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "詞頻表", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "詞頻表", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '恢復自動儲存
End '用Exit Sub無法每次關閉Access
錯誤處理:
Select Case Err.Number '主索引值重複
    Case Is = 91, 3078
        MsgBox "請再按一次!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub 指定4字數詞頻()       '2002/11/15
On Error GoTo 錯誤處理
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '取消自動儲存
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\千慮一得齋\書籍資料\詞頻.mdb", False
d.docmd.SelectObject d.acTable, "詞頻表", True
'd.Visible = True '檢查用
Set db = d.CurrentDb
Set rst = db.OpenRecordset("詞頻表", dbOpenDynaset)
If rst.RecordCount > 0 Then '要獲得全部的筆數須用MoveLast但此只需判斷有沒有原先的記錄即可!
'rst打開以後只會取得第一筆記錄!
'    db.Execute "DELETE 字頻表.* FROM 字頻表"
    db.Execute "DELETE * FROM 詞頻表"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 3 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2) & .Characters(phra + 3)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '計次
            GoTo 11 '字串右邊是半形空格時,AccessUpdate時會銷去,且於詞彙亦無意意,故不計!
        End If
        '直接進入下一個字串比對
        wrong = wrong + 1 '檢視用!
'        If wrong Mod 29688 = 0 Then '到29688時會產生OLE沒有回應的錯誤,故在此歇會兒
'            MsgBox "請檢查!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "請檢查!!"
'        End If
        With rst
            .FindFirst "詞彙 like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("詞彙") = phras
'                rst("次數") = 1'預設值已定為1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("次數") = rst("次數") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "統計完成!!" & vbCr & "(※共執行了" & wrong & "次的檢查※)" _
        & "詞彙右邊半形空格凡" & hfspace & "次,忽略不計!" _
        & vbCr & "※耗時:" & Format(EndTime - StTime, "n分s秒") & "※" _
        & vbCr & "字元數=" & .Characters.Count
End With
If MsgBox("要即刻檢視結果嗎?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\千慮一得齋\書籍資料\詞頻.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "詞頻表", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "詞頻表", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '恢復自動儲存
End '用Exit Sub無法每次關閉Access
錯誤處理:
Select Case Err.Number '主索引值重複
    Case Is = 91, 3078
        MsgBox "請再按一次!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub 指定3字數詞頻()      '2002/11/15
On Error GoTo 錯誤處理
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '取消自動儲存
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\千慮一得齋\書籍資料\詞頻.mdb", False
d.docmd.SelectObject d.acTable, "詞頻表", True
'd.Visible = True '檢查用
Set db = d.CurrentDb
Set rst = db.OpenRecordset("詞頻表", dbOpenDynaset)
If rst.RecordCount > 0 Then '要獲得全部的筆數須用MoveLast但此只需判斷有沒有原先的記錄即可!
'rst打開以後只會取得第一筆記錄!
'    db.Execute "DELETE 字頻表.* FROM 字頻表"
    db.Execute "DELETE * FROM 詞頻表"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 2 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '計次
            GoTo 11 '字串右邊是半形空格時,AccessUpdate時會銷去,且於詞彙亦無意意,故不計!
        End If
        '直接進入下一個字串比對
        wrong = wrong + 1 '檢視用!
'        If wrong Mod 29688 = 0 Then '到29688時會產生OLE沒有回應的錯誤,故在此歇會兒
'            MsgBox "請檢查!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "請檢查!!"
'        End If
        With rst
            .FindFirst "詞彙 like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("詞彙") = phras
'                rst("次數") = 1'預設值已定為1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("次數") = rst("次數") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "統計完成!!" & vbCr & "(※共執行了" & wrong & "次的檢查※)" _
        & "詞彙右邊半形空格凡" & hfspace & "次,忽略不計!" _
        & vbCr & "※耗時:" & Format(EndTime - StTime, "n分s秒") & "※" _
        & vbCr & "字元數=" & .Characters.Count
End With
If MsgBox("要即刻檢視結果嗎?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\千慮一得齋\書籍資料\詞頻.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "詞頻表", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "詞頻表", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '恢復自動儲存
End '用Exit Sub無法每次關閉Access
錯誤處理:
Select Case Err.Number '主索引值重複
    Case Is = 91, 3078
        MsgBox "請再按一次!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub 指定2字數詞頻()       '2002/11/15
On Error GoTo 錯誤處理
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '取消自動儲存
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\千慮一得齋\書籍資料\詞頻.mdb", False
d.docmd.SelectObject d.acTable, "詞頻表", True
'd.Visible = True '檢查用
Set db = d.CurrentDb
Set rst = db.OpenRecordset("詞頻表", dbOpenDynaset)
If rst.RecordCount > 0 Then '要獲得全部的筆數須用MoveLast但此只需判斷有沒有原先的記錄即可!
'rst打開以後只會取得第一筆記錄!
'    db.Execute "DELETE 字頻表.* FROM 字頻表"
    db.Execute "DELETE * FROM 詞頻表"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 1 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '計次
            GoTo 11 '字串右邊是半形空格時,AccessUpdate時會銷去,且於詞彙亦無意意,故不計!
        End If
        '直接進入下一個字串比對
        wrong = wrong + 1 '檢視用!
'        If wrong Mod 29688 = 0 Then '到29688時會產生OLE沒有回應的錯誤,故在此歇會兒
'            MsgBox "請檢查!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "請檢查!!"
'        End If
        With rst
            .FindFirst "詞彙 like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("詞彙") = phras
'                rst("次數") = 1'預設值已定為1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("次數") = rst("次數") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "統計完成!!" & vbCr & "(※共執行了" & wrong & "次的檢查※)" _
        & "詞彙右邊半形空格凡" & hfspace & "次,忽略不計!" _
        & vbCr & "※耗時:" & Format(EndTime - StTime, "n分s秒") & "※" _
        & vbCr & "字元數=" & .Characters.Count
End With
If MsgBox("要即刻檢視結果嗎?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\千慮一得齋\書籍資料\詞頻.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "詞頻表", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "詞頻表", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '恢復自動儲存
End '用Exit Sub無法每次關閉Access
錯誤處理:
Select Case Err.Number '主索引值重複
    Case Is = 91, 3078
        MsgBox "請再按一次!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub 指定1字數詞頻()        '2002/11/15
On Error GoTo 錯誤處理
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '取消自動儲存
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\千慮一得齋\書籍資料\詞頻.mdb", False
d.docmd.SelectObject d.acTable, "詞頻表", True
'd.Visible = True '檢查用
Set db = d.CurrentDb
Set rst = db.OpenRecordset("詞頻表", dbOpenDynaset)
If rst.RecordCount > 0 Then '要獲得全部的筆數須用MoveLast但此只需判斷有沒有原先的記錄即可!
'rst打開以後只會取得第一筆記錄!
'    db.Execute "DELETE 字頻表.* FROM 字頻表"
    db.Execute "DELETE * FROM 詞頻表"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
            phras = .Characters(phra)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '計次
            GoTo 11 '字串右邊是半形空格時,AccessUpdate時會銷去,且於詞彙亦無意意,故不計!
        End If
        '直接進入下一個字串比對
        wrong = wrong + 1 '檢視用!
'        If wrong Mod 29688 = 0 Then '到29688時會產生OLE沒有回應的錯誤,故在此歇會兒
'            MsgBox "請檢查!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "請檢查!!"
'        End If
        With rst
            .FindFirst "詞彙 like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("詞彙") = phras
'                rst("次數") = 1'預設值已定為1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("次數") = rst("次數") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "統計完成!!" & vbCr & "(※共執行了" & wrong & "次的檢查※)" _
        & "詞彙右邊半形空格凡" & hfspace & "次,忽略不計!" _
        & vbCr & "※耗時:" & Format(EndTime - StTime, "n分s秒") & "※" _
        & vbCr & "字元數=" & .Characters.Count
End With
If MsgBox("要即刻檢視結果嗎?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\千慮一得齋\書籍資料\詞頻.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "詞頻表", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "詞頻表", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '恢復自動儲存
End '用Exit Sub無法每次關閉Access
錯誤處理:
Select Case Err.Number '主索引值重複
    Case Is = 91, 3078
        MsgBox "請再按一次!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub 指定7字數詞頻()      '2002/11/15'以此為例,可作為預先限定字數的各個程序(本例為7個字的查詢)
On Error GoTo 錯誤處理
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
'phralh = InputBox("請用阿拉伯數字指定詞的組成字數,最多字數為「11」!", "指定詞彙字數", "2")
'If phralh = "" Or Not IsNumeric(phralh) Then Exit Sub
'If CByte(phralh) > 11 Or CByte(phralh) < 1 Then Exit Sub
Options.SaveInterval = 0 '取消自動儲存
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\千慮一得齋\書籍資料\詞頻.mdb", False
d.docmd.SelectObject d.acTable, "詞頻表", True
'd.Visible = True '檢查用
Set db = d.CurrentDb
Set rst = db.OpenRecordset("詞頻表", dbOpenDynaset)
If rst.RecordCount > 0 Then '要獲得全部的筆數須用MoveLast但此只需判斷有沒有原先的記錄即可!
'rst打開以後只會取得第一筆記錄!
'    db.Execute "DELETE 字頻表.* FROM 字頻表"
    db.Execute "DELETE * FROM 詞頻表"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 6 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2) & .Characters(phra + 3) & _
                    .Characters(phra + 4) & .Characters(phra + 5) & _
                    .Characters(phra + 6)
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '計次
            GoTo 11 '字串右邊是半形空格時,AccessUpdate時會銷去,且於詞彙亦無意意,故不計!
        End If
        '直接進入下一個字串比對
        wrong = wrong + 1 '檢視用!
'        If wrong Mod 29688 = 0 Then '到29688時會產生OLE沒有回應的錯誤,故在此歇會兒
'            MsgBox "請檢查!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "請檢查!!"
'        End If
        With rst
            .FindFirst "詞彙 like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("詞彙") = phras
'                rst("次數") = 1'預設值已定為1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("次數") = rst("次數") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "統計完成!!" & vbCr & "(※共執行了" & wrong & "次的檢查※)" _
        & "詞彙右邊半形空格凡" & hfspace & "次,忽略不計!" _
        & vbCr & "※耗時:" & Format(EndTime - StTime, "n分s秒") & "※" _
        & vbCr & "字元數=" & .Characters.Count
End With
If MsgBox("要即刻檢視結果嗎?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\千慮一得齋\書籍資料\詞頻.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "詞頻表", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "詞頻表", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '恢復自動儲存
End '用Exit Sub無法每次關閉Access
錯誤處理:
Select Case Err.Number '主索引值重複
    Case Is = 91, 3078
        MsgBox "請再按一次!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub 指定字數詞頻1() '2002/11/15'效能較慢!
On Error GoTo 錯誤處理
Dim wrong As Long, phra As Long, phras, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Dim a1, i As Byte, j As Byte
phralh = InputBox("請用阿拉伯數字指定詞的組成字數,最多字數為「255」!", "指定詞彙字數", "2")
If phralh = "" Or Not IsNumeric(phralh) Then Exit Sub
If CByte(phralh) > 255 Or CByte(phralh) < 1 Then Exit Sub
Options.SaveInterval = 0 '取消自動儲存
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\千慮一得齋\書籍資料\詞頻.mdb", False
d.docmd.SelectObject d.acTable, "詞頻表", True
'd.Visible = True '檢查用
Set db = d.CurrentDb
Set rst = db.OpenRecordset("詞頻表", dbOpenDynaset)
If rst.RecordCount > 0 Then '要獲得全部的筆數須用MoveLast但此只需判斷有沒有原先的記錄即可!
'rst打開以後只會取得第一筆記錄!
'    db.Execute "DELETE 字頻表.* FROM 字頻表"
    db.Execute "DELETE * FROM 詞頻表"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        j = CByte(phralh)
        ReDim a1(1 To j) As String
        If j > 1 Then
            If phra + (phralh - 1) <= .Characters.Count Then
                For j = 1 To j
                    For i = 0 To j - 1
                            a1(j) = a1(j) & .Characters(phra + i)
                    Next i
    '                    Debug.Print a1(j)
                Next j
                phras = a1(j - 1)
            End If
        Else
            phras = .Characters(phra)
        End If
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '計次
            GoTo 11 '字串右邊是半形空格時,AccessUpdate時會銷去,且於詞彙亦無意意,故不計!
        End If
        '直接進入下一個字串比對
        wrong = wrong + 1 '檢視用!
'        If wrong Mod 29688 = 0 Then '到29688時會產生OLE沒有回應的錯誤,故在此歇會兒
'            MsgBox "請檢查!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "請檢查!!"
'        End If
        With rst
            .FindFirst "詞彙 like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("詞彙") = phras
'                rst("次數") = 1'預設值已定為1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("次數") = rst("次數") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "統計完成!!" & vbCr & "(※共執行了" & wrong & "次的檢查※)" _
        & "詞彙右邊半形空格凡" & hfspace & "次,忽略不計!" _
        & vbCr & "※耗時:" & Format(EndTime - StTime, "n分s秒") & "※" _
        & vbCr & "字元數=" & .Characters.Count
End With
If MsgBox("要即刻檢視結果嗎?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\千慮一得齋\書籍資料\詞頻.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "詞頻表", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "詞頻表", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '恢復自動儲存
End '用Exit Sub無法每次關閉Access
錯誤處理:
Select Case Err.Number '主索引值重複
    Case Is = 91, 3078
        MsgBox "請再按一次!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub 指定字數詞頻2() '2002/11/15效能與原設計差不多,但可變數化!
On Error GoTo 錯誤處理
Dim wrong As Long, phra As Long, phras, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Dim i As Byte, j As Byte
phralh = InputBox("請用阿拉伯數字指定詞的組成字數,最多字數為「255」!", "指定詞彙字數", "2")
If phralh = "" Or Not IsNumeric(phralh) Then Exit Sub
If CByte(phralh) > 255 Or CByte(phralh) < 1 Then Exit Sub
Options.SaveInterval = 0 '取消自動儲存
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\千慮一得齋\書籍資料\詞頻.mdb", False
d.docmd.SelectObject d.acTable, "詞頻表", True
'd.Visible = True '檢查用
Set db = d.CurrentDb
Set rst = db.OpenRecordset("詞頻表", dbOpenDynaset)
If rst.RecordCount > 0 Then '要獲得全部的筆數須用MoveLast但此只需判斷有沒有原先的記錄即可!
'rst打開以後只會取得第一筆記錄!
'    db.Execute "DELETE 字頻表.* FROM 字頻表"
    db.Execute "DELETE * FROM 詞頻表"
End If
StTime = Time
j = CByte(phralh)
With ActiveDocument
    For phra = 1 To .Characters.Count
'        If j > 1 Then'即使是單字也不須分別處理了!!
            If phra + (phralh - 1) <= .Characters.Count Then
                phras = ""
                For i = 0 To j - 1
                    phras = phras & .Characters(phra + i)
                Next i
            End If
'        Else
'            phras = .Characters(phra)
'        End If
        If Len(phras) > 1 And Right(phras, 1) = " " Then
            hfspace = hfspace + 1 '計次
            GoTo 11 '字串右邊是半形空格時,AccessUpdate時會銷去,且於詞彙亦無意意,故不計!
        End If
        '直接進入下一個字串比對
        wrong = wrong + 1 '檢視用!
'        If wrong Mod 29688 = 0 Then '到29688時會產生OLE沒有回應的錯誤,故在此歇會兒
'            MsgBox "請檢查!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "請檢查!!"
'        End If
        With rst
            .FindFirst "詞彙 like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("詞彙") = phras
'                rst("次數") = 1'預設值已定為1
                .Update 'dbUpdateBatch, True
            Else
                .Edit
                rst("次數") = rst("次數") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "統計完成!!" & vbCr & "(※共執行了" & wrong & "次的檢查※)" _
        & "詞彙右邊半形空格凡" & hfspace & "次,忽略不計!" _
        & vbCr & "※耗時:" & Format(EndTime - StTime, "n分s秒") & "※" _
        & vbCr & "字元數=" & .Characters.Count
End With
If MsgBox("要即刻檢視結果嗎?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\千慮一得齋\書籍資料\詞頻.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "詞頻表", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "詞頻表", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '恢復自動儲存
End '用Exit Sub無法每次關閉Access
錯誤處理:
Select Case Err.Number '主索引值重複
    Case Is = 91, 3078
        MsgBox "請再按一次!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub 文件字頻_old()
Dim DR As Range, d As Document, Char, charText As String, preChar As String _
    , x() As String, xT() As Long, i As Long, j As Long, ExcelSheet  As Object, _
    ds As Date, de As Date '
Static xlsp As String
On Error GoTo ErrH:
'xlsp = "C:\Documents and Settings\Superwings\桌面\"
Set d = ActiveDocument
xlsp = 取得桌面路徑 & "\" 'GetDeskDir() & "\"
If Dir(xlsp) = "" Then xlsp = 取得桌面路徑 'GetDeskDir ' "C:\Users\Wong\Desktop\" '& Replace(ActiveDocument.Name, ".doc", "") & "字頻.XLS"
'If Dir(xlsp) = "" Then xlsp = "C:\Documents and Settings\Superwings\桌面\" & Replace(ActiveDocument.Name, ".doc", "") & "字頻.XLS"
'xlsp = "C:\Documents and Settings\Superwings\桌面\" & Replace(ActiveDocument.Name, ".doc", "") & "字頻.XLS"
xlsp = InputBox("請輸入存檔路徑及檔名(全檔名,含副檔名)!" & vbCr & vbCr & _
        "預設將以此word文件檔名 + ""字頻.XLSX""字綴,存於桌面上", "字頻調查", xlsp & Replace(ActiveDocument.Name, ".doc", "") & "字頻" & StrConv(Time, vbWide) & ".XLSX")
If xlsp = "" Then Exit Sub

ds = VBA.Timer

With d
    For Each Char In d.Characters
        charText = Char
        If Not charText = Chr(13) And charText <> "-" And Not charText Like "[a-zA-Z0-9０-９]" Then
            'If Not charText Like "[a-z1-9]" & Chr(-24153) & Chr(-24152) & " 　、'""「」『』（）－？！]" Then
'            If InStr(Chr(-24153) & Chr(-24152) & Chr(2) & "•[]〔〕﹝﹞…；,，.。． 　、'""‘’`\{}｛｝「」『』（）《》〈〉－？！]", charText) = 0 Then
            If InStr(ChrW(-24153) & ChrW(-24152) & Chr(2) & "•[]〔〕﹝﹞…；,，.。． 　、'""‘’`\{}｛｝「」『』（）《》〈〉－？！]", charText) = 0 Then
            'chr(2)可能是註腳標記
                If preChar <> charText Then
                    'If UBound(X) > 0 Then
                        If preChar = "" Then 'If IsEmpty(X) Then'如果是一開始
                            GoTo 1
                        ElseIf UBound(Filter(x, charText)) Then ' <> charText Then  '如果尚無此字
1                           ReDim Preserve x(i)
                            ReDim Preserve xT(i)
                            x(i) = charText
                            xT(i) = xT(i) + 1
                            i = i + 1
                        Else
                            GoSub 字頻加一
                        End If
                    'End If
                Else
                    GoSub 字頻加一
                End If
                preChar = Char
            End If
        End If
    Next Char
End With

Dim Doc As New Document, Xsort() As String, u As Long ', xTsort() As Integer, k As Long, so As Long, ww As String
'ReDim Xsort(i) As String ', xtsort(i) as Integer
'ReDim Xsort(d.Characters.Count) As String
If u = 0 Then u = 1 '若無執行「字頻加一:」副程序,若無超過1次的字頻，則　Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "、" & x(j - 1) & _
                                會出錯：陣列索引超出範圍 2015/11/5

ReDim Xsort(u) As String
Set ExcelSheet = CreateObject("Excel.Sheet")
With ExcelSheet.Application
    For j = 1 To i
        .Cells(j, 1) = x(j - 1)
        .Cells(j, 2) = xT(j - 1)
        Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "、" & x(j - 1) 'Xsort(xT(j - 1)) & ww '陣列排序'2010/10/29
    Next j
End With
'Doc.ActiveWindow.Visible = False
'U = UBound(Xsort)
For j = u To 0 Step -1 '陣列排序'2010/10/29
    If Xsort(j) <> "" Then
        With Doc
            If Len(.Range) = 1 Then '尚未輸入內容
                .Range.InsertAfter "字頻 = " & j & "次：（" & Len(Replace(Xsort(j), "、", "")) & "字）"
                .Range.Paragraphs(1).Range.Font.Size = 12
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "新細明體"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
                '.Range.Paragraphs(1).Range.Font.Bold = True
            Else
                .Range.InsertParagraphAfter
                .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
                .Range.InsertAfter "字頻 = " & j & "次：（" & Len(Replace(Xsort(j), "、", "")) & "字）"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Size = 12
                '.Range.Paragraphs(.Paragraphs.Count).Range.Bold = True
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "新細明體"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
            End If
            .Range.InsertParagraphAfter
            .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
            .Range.Paragraphs(.Paragraphs.Count).Range.Font.Size = 12
'            .Range.Paragraphs(.Paragraphs.Count).Range.Bold = False
            .Range.InsertAfter Replace(Xsort(j), "、", Chr(9), 1, 1) 'chr(9)為定位字元(Tab鍵值)
            .Range.InsertParagraphAfter
            If InStr(.Range.Paragraphs(.Paragraphs.Count).Range, "字頻") = 0 Then
                .Range.Paragraphs(.Paragraphs.Count - 1).Range.Font.Name = "標楷體"
            Else
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "新細明體"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
            End If
        End With
    End If
Next j

With Doc.Paragraphs(1).Range
     .InsertParagraphBefore
     .Font.NameAscii = "times new roman"
    Doc.Paragraphs(1).Range.InsertParagraphAfter
    Doc.Paragraphs(1).Range.InsertParagraphAfter
    Doc.Paragraphs(1).Range.InsertAfter "你提供的文本共使用了" & i & "個不同的字（傳統字與簡化字不予合併）"
End With

Doc.ActiveWindow.Visible = True
'

'U = UBound(xT)
'ReDim Xsort(U) As String, xTsort(U) As Long
'
'i = d.Characters
'For j = 1 To i '用數字相比
'    For k = 0 To U 'xT陣列中每個元素都與j比
'        If xT(k) = j Then
'            Xsort(so) = x(k)
'            xTsort(so) = xT(k)
'            so = so + 1
'        End If
'    Next k
'Next j

'With doc
'    .Range.InsertAfter "字頻=0001"
'    .Range.InsertParagraphAfter
'End With


' Cells.Select
'    Selection.Sort Key1:=Range("B1"), Order1:=xlDescending, Header:=xlGuess, _
'        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom


'Set ExcelSheet = Nothing'此行會使消失
'Set d = Nothing
de = VBA.Timer
MsgBox "完成！" & vbCr & vbCr & "費時" & Left(de - ds, 5) & "秒!"
ExcelSheet.Application.Visible = True
ExcelSheet.Application.UserControl = True
ExcelSheet.SaveAs xlsp '"C:\Macros\守真TEST.XLS"
Doc.SaveAs Replace(xlsp, "XLS", "doc") '分大小寫
'Doc.SaveAs "c:\test1.doc"
AppActivate "microsoft excel"
Exit Sub
字頻加一:
For j = 0 To UBound(x)
    If x(j) = charText Then
        xT(j) = xT(j) + 1
        If u < xT(j) Then u = xT(j) '記下最高字頻,以便排序(將欲排序之陣列最高元素值設為此,則不會超出陣列.
        '多此一行因為要重複判斷計算好幾次,故效能不增反減''效能還是差不多啦.
        Exit For
    End If
Next j

Return
ErrH:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & Err.Description, vbCritical 'STOP: Resume
'        Resume
        End
    
End Select
End Sub

Function lEnglish() '英文大寫字母
Dim Wd, wdct As Long, i As Byte
For i = 65 To 90
    Debug.Print Chr(i) & vbCr
Next
End Function
Function sEnglish() '英文小寫字母
Dim i As Byte
For i = 97 To 122
    Debug.Print Chr(i) & vbCr
Next
End Function

Function trimStrForSearch(x As String, sl As word.Selection) As String
'https://docs.microsoft.com/zh-tw/dotnet/visual-basic/programming-guide/language-features/procedures/passing-arguments-by-value-and-by-reference
Dim ayToTrim As Variant, a As Variant, rng As Range, slTxtR As String
On Error GoTo eH
slTxtR = sl.Characters(sl.Characters.Count)
ayToTrim = Array(Chr(13), Chr(9), Chr(10), Chr(11), Chr(13) & Chr(7), Chr(13) & Chr(10))
x = VBA.Trim(x)
For Each a In ayToTrim
    x = VBA.Replace(x, a, "")
Next a
trimStrForSearch = x
If sl.Type <> wdSelectionIP Then
    If UBound(VBA.Strings.Filter(ayToTrim, slTxtR)) > -1 Then
    'If sl.Characters(sl.Characters.Count) = Chr(13) Then
        Set rng = sl.Range
        rng.SetRange sl.start, sl.End - Len(slTxtR)
        rng.Select
    End If
End If
Exit Function
eH:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & Err.Description
'        Resume
End Select
End Function


Function Symbol() '標點符號表
Dim f As Variant
f = Array("。", "」", Chr(-24152), "：", "，", "；", _
    "、", "「", ".", Chr(34), ":", ",", ";", _
    "……", "...", "）", ")", "-")  '先設定標點符號陣列以備用
                                'Chr(-24152)是「”」,由Asc函數在選取(.SelText)「”」時取得;Chr(34):「"」
End Function
Function isSymbol(ByVal a As String) As Boolean
Dim f As String
 f = "（。「」『』[]【】〔〕《》〈〉-－"",  ：，；" _
    & "、. :,;" _
    & "……...）!)-" & Chr(34) & Chr(-24152)
If InStr(1, f, a, vbTextCompare) Then
    isSymbol = True
End If
End Function
Function is注音符號(ByVal a As String, Optional rng As Variant) As Boolean
Dim f As String
On Error GoTo eH
If Len(a) > 1 Then Exit Function
f = "ㄅㄆㄇㄈㄉㄊㄋㄌㄍㄎㄏㄐㄑㄒㄓㄔㄕㄖㄗㄘㄙㄧㄨㄩㄚㄛㄜㄝㄞㄟㄠㄡㄢㄣㄤㄥㄦˊ  ˇ  ˋ  ˙"
If a = ChrW(20008) Then
    If Not rng Is Nothing Then
        If rng.start = 0 Then
            If InStr(f, rng.Next.Characters(1)) Then
                is注音符號 = True
                Exit Function
            End If
        ElseIf rng.End = rng.Document.Range.End - 1 Then
            If InStr(f, rng.Previous.Characters(1)) Then
                is注音符號 = True
                Exit Function
            End If
        End If
    End If
Else
    If InStr(f, a) Then is注音符號 = True
End If
Exit Function
eH:
Select Case Err.Number
    Case 424 '此處需要物件
        Set rng = Nothing
        Resume
    Case Else
        MsgBox Err.Number & Err.Description
        Debug.Print Err.Number & Err.Description
End Select
End Function

Sub 選取段落符號()
'第1段的最後()
'    With ActiveDocument.Paragraphs(1).Range
'        ActiveDocument.Range(.End - 1, .End).Select
'    End With
Dim i As Integer
For i = 1 To ActiveDocument.Paragraphs.Count
    With ActiveDocument.Paragraphs(i).Range
        ActiveDocument.Range(.End - 1, .End).Select
    End With
Next i
End Sub


Sub 造字字元檢查() '非細明體檢查,2004/8/23
Dim ch
For Each ch In ActiveDocument.Characters
'    If AscW(ch) < -1491 Or AscW(ch) > 19968 Then
    If Asc(ch) < -24256 Or (0 > Asc(ch) And Asc(ch) >= -1468) Then
        ch.Select
        ch.Font.Name = "EUDC"
    End If
Next ch
End Sub

Sub 注腳符號置換() '2004/10/17
Dim Wd As Range 'As Range 'Words物件即表一個Range物件,見線上說明!
'Dim i As Long ' Integer
'要先執行全形轉半形,這樣words才能正確判斷為數字
全形數字轉換成半形數字
With Selection '原以整份文件(ActiveDocument),今但以選取範圍整理,但因更改值而影響,作廢!
    If .Type = wdSelectionIP Then .Document.Select '如果沒有選取範圍(為插入點)則處理整份文件
    If .Document.path = "" Then
        For Each Wd In .words
            '要是數字且前後不能加﹝﹞或〔〕才執行！
            If Not Wd.Text Like "﹝" And Not Wd.Text Like "〔" And Not Wd Like "[[]" And Not Wd Like "[]]" Then
                If IsNumeric(Wd) Then
                    If Wd.End = .Document.Content.StoryLength Or Wd.start = 0 Then GoTo w '文件之首尾另外處理
                    If Not Wd.Previous Like "﹝" And Not Wd.Previous Like "〔" And Not Wd.Previous Like "[[]" _
                        And Not Wd.Next Like "﹞" And Not Wd.Next Like "〕" And Not Wd.Next Like "]" Then
w:                      If Wd <= 20 Then 'Arial Unicode MS[種類]裡"括號文數字"只有二十個!
                            With Wd
                                '選取會改變Selection的範圍,故今取消!
'                                .Select 'Words物件即表一個Range物件,見線上說明!
                                .Font.Name = "Arial Unicode MS"
                                Wd.Text = ChrW((9312 - 1) + Wd)
                            End With
                        Else '超過20號的註腳時
                            With Wd
                                .Text = "﹝" & Wd.Text & "﹞" '加括號
                            End With
        '                    MsgBox "有超過20號的註腳,不能執行！", vbCritical
        '                    Do Until .Undo(i) = False '還原直至不能還原（還原所有動作）
        '                    i = i + 1
        '                    Loop
        '                    StatusBar = "Undo was successful " & i & " times!!" '在狀態列顯示文字！
        '                    Exit Sub
                        End If
                    End If
                End If
            End If
        Next
        MsgBox "執行完畢！", vbInformation
    Else
        MsgBox "本文件不能操作!", vbCritical
    End If
End With
End Sub

Sub 全形數字轉換成半形數字() '2004/10/17-由圖書管理複製改撰的原式－不好，會影響字形
Dim FNumArray, HNumArray, i As Byte, e As Range
FNumArray = Array("０", "１", "２", "３", "４", "５", "６", "７", "８", "９")
HNumArray = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9")
With ActiveDocument
    For Each e In .Characters
        For i = 1 To UBound(FNumArray) + 1
            If e.Text Like FNumArray(i - 1) Then
                e.Text = HNumArray(i - 1)
        End If
        Next i
    Next e
End With
End Sub

Sub 全形轉半形()
With Selection
    .Range = StrConv(.Range, vbNarrow)
End With
End Sub
Sub 圓括號改篇名號()
If Selection.Type = wdSelectionIP Then Selection.HomeKey wdStory: Selection.EndKey wdStory, wdExtend
Selection.Text = Replace(Replace(Selection.Text, "（", "〈"), "）", "〉")
End Sub


Sub 校勘文字標色() '2009/8/23
Register_Event_Handler
'指定鍵F2
' 巨集2 巨集
' 巨集錄製於 2009/8/23，錄製者 Oscar Sun
'
'    Selection.MoveDown Unit:=wdLine, Count:=2
'    Selection.EndKey Unit:=wdLine
'    Selection.MoveLeft Unit:=wdCharacter, Count:=1
'    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
If Selection.Type = wdSelectionIP Then Exit Sub
    With Selection.Font.Shading
        If InStr(ActiveDocument.Name, "排印") Then
            .Parent.Color = wdColorRed
            .Texture = wdTextureNone
        Else
            If .Texture = wdTextureNone Then '字元網底
                .Texture = wdTexture15Percent
                .ForegroundPatternColor = wdColorBlack
                .BackgroundPatternColor = wdColorWhite
                .Parent.Color = wdColorRed
            Else
                .Texture = wdTextureNone '字元網底
                .Parent.Color = wdColorAutomatic
            End If
        End If
    End With
    If InStr(ActiveDocument.Name, "排印") Then
        ActiveDocument.save
'        setOX
'        OX.WinActivate "Microsoft Excel"
        'Dim e As New Excel.Application
        Dim e
        Set e = Excel.Application
        Dim r As Long, i As Byte
        With Selection
            Set e = GetObject(, "Excel.application")
            AppActivate "microsoft excel"
            With e
                '.ActiveWorkbook.Save
                r = .ActiveCell.Row
                For i = 1 To 7
                    If .Cells(r, i).Value <> "" Then
                        MsgBox "請到新記錄列！！", vbExclamation
                        Exit Sub
                    End If
                Next i
                .Cells(r, 1).Activate
                DoEvents
                .activesheet.Paste
                .Cells(r, 2).Value = Selection
                .Cells(r, 2).Font.Color = wdColorRed
                If Not Selection Like "*[����������������������������☆★｜　]*" Then
                    .Cells(r, 5) = Len(Selection)
                ElseIf Selection Like "*　*" Then
                    .Cells(r, 5) = Len(Selection) - 1
                Else
                    .Cells(r, 5) = 1
                End If
                .ActiveWorkbook.save
                .Cells(.ActiveCell.Row + 1, .ActiveCell.Column).Activate
            End With
        End With
        游標所在位置書籤
        OX.WinActivate "Adobe Reader"
        AppActivate "microsoft word"
    End If
End Sub

Sub 註腳編號前後加方括號()
With Selection
    Do

        Selection.GoTo What:=wdGoToFootnote, Which:=wdGoToNext, Count:=1, Name:=""
'        Selection.GoTo What:=wdGoToFootnote, Which:=wdGoToNext, Count:=1, Name:=""
        Selection.Find.ClearFormatting
'        With Selection.Find
'            .Text = ""
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindStop
'            .Format = False
'            .MatchCase = False
'            .MatchWholeWord = False
'            .MatchByte = True
'            .MatchWildcards = False
'            .MatchSoundsLike = False
'            .MatchAllWordForms = False
'        End With
'        If .Find.Execute() = False Then Exit Do
        'Application.Browser.Next
        .TypeText Text:="["
        .MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
        .Font.Superscript = wdToggle
'        Selection.Copy
'        Selection.MoveRight Unit:=wdCharacter, Count:=3
'        Selection.MoveLeft Unit:=wdCharacter, Count:=1
'        Selection.Paste
'        Selection.MoveLeft Unit:=wdCharacter, Count:=1
'        Selection.Delete Unit:=wdCharacter, Count:=1
'        Selection.TypeText Text:="》"
'        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.MoveRight Unit:=wdCharacter, Count:=2
        'Selection.TypeBackspace
        Selection.TypeText Text:="]"
        'Selection.MoveRight Unit:=wdCharacter, Count:=1
    Loop 'While .Find.Execute()
End With
End Sub

Sub 大陸引號換台灣引號()
Dim a, b, i
a = Array(-24153, -24152, -24155, -24154)  '“,”,‘,”
b = Array("「", "」", "『", "』")

With ActiveDocument.Range.Find
    For i = 0 To 3
        '.Text = a(i)
         '.Replacement.Text = b(i)
         .ClearFormatting
         .Execute Chr(a(i)), , , , , , , , , b(i), wdReplaceAll
    Next i
End With
End Sub


Sub 文件字頻()
Dim d As Document, Char, charText As String, preChar As String _
    , x() As String, xT() As Long, i As Long, j As Long, ds As Date, de As Date     '
'Dim ExcelSheet  As New Excel.Worksheet 'As Object,
'這是之前以先期引用的方式，在設定引用項目中手動加入的寫法:https://hankvba.blogspot.com/2018/03/vba.html  、 http://markc0826.blogspot.com/2012/07/blog-post.html
'Dim xlApp As Excel.Application, xlBook As Excel.Workbook, xlSheet As Excel.Worksheet
''這就是後期引用，以自訂新仿Excel類別的方法來實作(如此寫的緣故是原來要改寫的程式碼就會比較少，變動較小，且也不必再New出一個執行個體才能執行：
Dim xlApp, xlBook, xlSheet
Set xlApp = Excel.Application
Set xlBook = Excel.Workbook
Set xlSheet = Excel.Worksheet
Dim ReadingLayoutB As Boolean
Static xlsp As String
On Error GoTo ErrH:
'xlsp = "C:\Documents and Settings\Superwings\桌面\"
Set d = ActiveDocument
xlsp = 取得桌面路徑 & "\" 'GetDeskDir() & "\"
If Dir(xlsp) = "" Then xlsp = 取得桌面路徑 'GetDeskDir ' "C:\Users\Wong\Desktop\" '& Replace(ActiveDocument.Name, ".doc", "") & "字頻.XLS"
'If Dir(xlsp) = "" Then xlsp = "C:\Documents and Settings\Superwings\桌面\" & Replace(ActiveDocument.Name, ".doc", "") & "字頻.XLS"
'xlsp = "C:\Documents and Settings\Superwings\桌面\" & Replace(ActiveDocument.Name, ".doc", "") & "字頻.XLS"
xlsp = InputBox("請輸入存檔路徑及檔名(全檔名,含副檔名)!" & vbCr & vbCr & _
        "預設將以此word文件檔名 + ""字頻.XLSX""字綴,存於桌面上", "字頻調查", xlsp & Replace(ActiveDocument.Name, ".doc", "") & "字頻" & StrConv(Time, vbWide) & ".XLSX")
If xlsp = "" Then Exit Sub

ds = VBA.Timer

With d
    For Each Char In d.Characters
        charText = Char
        If InStr("()：>" & Chr(13) & Chr(9) & Chr(10) & Chr(11) & ChrW(12), charText) = 0 And charText <> "-" And Not charText Like "[a-zA-Z0-9０-９]" Then
            'If Not charText Like "[a-z1-9]" & Chr(-24153) & Chr(-24152) & " 　、'""「」『』（）－？！]" Then
'            If InStr(Chr(-24153) & Chr(-24152) & Chr(2) & "•[]〔〕﹝﹞…；,，.。． 　、'""‘’`\{}｛｝「」『』（）《》〈〉－？！]", charText) = 0 Then
            If InStr(ChrW(9312) & ChrW(-24153) & ChrW(-24152) & Chr(2) & "•[]〔〕﹝﹞…；,，.。． 　、'""‘’`\{}｛｝「」『』（）《》〈〉－？！]▽□】【~/︵—" & Chr(-24152) & Chr(-24153), charText) = 0 Then
            'chr(2)可能是註腳標記
                If preChar <> charText Then
                    'If UBound(X) > 0 Then
                        If preChar = "" Then 'If IsEmpty(X) Then'如果是一開始
                            GoTo 1
                        ElseIf UBound(Filter(x, charText)) Then ' <> charText Then  '如果尚無此字
1                           ReDim Preserve x(i)
                            ReDim Preserve xT(i)
                            x(i) = charText
                            xT(i) = xT(i) + 1
                            i = i + 1
                        Else
                            GoSub 字頻加一
                        End If
                    'End If
                Else
                    GoSub 字頻加一
                End If
                preChar = Char
            End If
        End If
    Next Char
End With

Dim Doc As New Document, Xsort() As String, u As Long ', xTsort() As Integer, k As Long, so As Long, ww As String
'ReDim Xsort(i) As String ', xtsort(i) as Integer
'ReDim Xsort(d.Characters.Count) As String
If u = 0 Then u = 1 '若無執行「字頻加一:」副程序,若無超過1次的字頻，則　Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "、" & x(j - 1) & _
                                會出錯：陣列索引超出範圍 2015/11/5

ReDim Xsort(u) As String
'Set ExcelSheet = CreateObject("Excel.Sheet")
'Set xlApp = CreateObject("Excel.Application")
'Set xlBook = xlApp.workbooks.Add
'Set xlSheet = xlBook.Worksheets(1)
With xlSheet.Application
    For j = 1 To i
        .Cells(j, 1) = x(j - 1)
        .Cells(j, 2) = xT(j - 1)
        Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "、" & x(j - 1) 'Xsort(xT(j - 1)) & ww '陣列排序'2010/10/29
    Next j
End With
'Doc.ActiveWindow.Visible = False
'U = UBound(Xsort)
For j = u To 0 Step -1 '陣列排序'2010/10/29
    If Xsort(j) <> "" Then
        With Doc
            If Len(.Range) = 1 Then '尚未輸入內容
                .Range.InsertAfter "字頻 = " & j & "次：（" & Len(Replace(Xsort(j), "、", "")) & "字）"
                .Range.Paragraphs(1).Range.Font.Size = 12
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "新細明體"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
                '.Range.Paragraphs(1).Range.Font.Bold = True
            Else
                .Range.InsertParagraphAfter
                .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
                .Range.InsertAfter "字頻 = " & j & "次：（" & Len(Replace(Xsort(j), "、", "")) & "字）"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Size = 12
                '.Range.Paragraphs(.Paragraphs.Count).Range.Bold = True
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "新細明體"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
            End If
            .Range.InsertParagraphAfter
            .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
            .Range.Paragraphs(.Paragraphs.Count).Range.Font.Size = 12
'            .Range.Paragraphs(.Paragraphs.Count).Range.Bold = False
            .Range.InsertAfter Replace(Xsort(j), "、", Chr(9), 1, 1) 'chr(9)為定位字元(Tab鍵值)
            .Range.InsertParagraphAfter
            If InStr(.Range.Paragraphs(.Paragraphs.Count).Range, "字頻") = 0 Then
                .Range.Paragraphs(.Paragraphs.Count - 1).Range.Font.Name = "標楷體"
            Else
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "新細明體"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
            End If
        End With
    End If
Next j

With Doc.Paragraphs(1).Range
     .InsertParagraphBefore
     .Font.NameAscii = "times new roman"
    Doc.Paragraphs(1).Range.InsertParagraphAfter
    Doc.Paragraphs(1).Range.InsertParagraphAfter
    Doc.Paragraphs(1).Range.InsertAfter "你提供的文本共使用了" & i & "個不同的字（傳統字與簡化字不予合併）"
End With

Doc.ActiveWindow.Visible = True
'

'U = UBound(xT)
'ReDim Xsort(U) As String, xTsort(U) As Long
'
'i = d.Characters
'For j = 1 To i '用數字相比
'    For k = 0 To U 'xT陣列中每個元素都與j比
'        If xT(k) = j Then
'            Xsort(so) = x(k)
'            xTsort(so) = xT(k)
'            so = so + 1
'        End If
'    Next k
'Next j

'With doc
'    .Range.InsertAfter "字頻=0001"
'    .Range.InsertParagraphAfter
'End With


' Cells.Select
'    Selection.Sort Key1:=Range("B1"), Order1:=xlDescending, Header:=xlGuess, _
'        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom


'Set ExcelSheet = Nothing'此行會使消失
'Set d = Nothing
de = VBA.Timer
If ReadingLayoutB Then d.ActiveWindow.View.ReadingLayout = Not d.ActiveWindow.View.ReadingLayout
MsgBox "完成！" & vbCr & vbCr & "費時" & Left(de - ds, 5) & "秒!", vbInformation
xlSheet.Application.Visible = True
xlSheet.Application.UserControl = True
xlSheet.SaveAs xlsp '"C:\Macros\守真TEST.XLS"
Doc.SaveAs Replace(xlsp, "XLS", "doc") '分大小寫
Set Excel.Application = Nothing
Exit Sub
字頻加一:
For j = 0 To UBound(x)
    If x(j) = charText Then
        xT(j) = xT(j) + 1
        If u < xT(j) Then u = xT(j) '記下最高字頻,以便排序(將欲排序之陣列最高元素值設為此,則不會超出陣列.
        '多此一行因為要重複判斷計算好幾次,故效能不增反減''效能還是差不多啦.
        Exit For
    End If
Next j

Return
ErrH:
Select Case Err.Number
    Case 4605 '閱讀模式不能編輯'此方法或屬性無法使用，因為此命令無法在閱讀中使用。
    '    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
    '        ActiveWindow.ActivePane.View.Type = wdNormalView
    '    Else
    '        ActiveWindow.View.Type = wdNormalView
    '    End If
    '    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
    '        ActiveWindow.ActivePane.View.Type = wdPrintView
    '    Else
    '        ActiveWindow.View.Type = wdPrintView
    '    End If
        'Doc.Application.ActiveWindow.View.ReadingLayout
        d.ActiveWindow.View.ReadingLayout = Not d.ActiveWindow.View.ReadingLayout
        Doc.ActiveWindow.View.ReadingLayout = False
        Doc.ActiveWindow.Visible = False
        ReadingLayoutB = True
        Resume
    Case Else
        MsgBox Err.Number & Err.Description, vbCritical 'STOP: Resume
        'Resume
        End
    
End Select
End Sub

Sub 文件詞頻() '由文件字頻改來'2015/11/28
Dim d As Document, Char, charText As String, preChar As String _
    , x() As String, xT() As Long, i As Long, j As Long, ds As Date, de As Date     '
'Dim ExcelSheet  As New Excel.Worksheet 'As Object,
'Dim xlApp As Excel.Application, xlBook As Excel.Workbook, xlSheet As Excel.Worksheet
Dim xlApp, xlBook, xlSheet
Set xlApp = Excel.Application
Set xlBook = Excel.Workbook
Set xlSheet = Excel.Worksheet
Dim ReadingLayoutB As Boolean
Static Ln
Dim xlsp As String
On Error GoTo ErrH:
Set d = ActiveDocument
'If xlsp = "" Then xlsp = 取得桌面路徑 & "\" 'GetDeskDir() & "\"
'If Dir(xlsp) = "" Then xlsp = 取得桌面路徑 'GetDeskDir
'xlsp = InputBox("請輸入存檔路徑及檔名(全檔名,含副檔名)!" & vbCr & vbCr & _
        "預設將以此word文件檔名 + ""詞頻.XLSX""字綴,存於桌面上", "詞頻調查", xlsp & Replace(d.Name, ".doc", "") & "詞頻" & StrConv(Time, vbWide) & ".XLSX")
'If xlsp = "" Then Exit Sub
xlsp = 取得桌面路徑 & "\" & Replace(d.Name, ".doc", "") & "_詞頻" & StrConv(Time, vbWide) & ".XLSX"
If Ln = "" Then Ln = 1
Ln = InputBox("請指定詞彙長度" & vbCr & vbCr & "檔案會存在桌面上名為:" & vbCr & vbCr & Replace(d.Name, ".doc", "") & "_詞頻" & StrConv(Time, vbWide) & ".XLSX" & _
                vbCr & vbCr & "的檔案", , Ln + 1)
If Ln = "" Then Exit Sub
If Not IsNumeric(Ln) Then Exit Sub
If Ln > 11 Or Ln < 2 Then Exit Sub


ds = VBA.Timer

With d
    For Each Char In d.Characters
        Select Case Ln
            Case 2
                charText = Char & Char.Next
            Case 3
                charText = Char & Char.Next & Char.Next.Next
            Case 4
                charText = Char & Char.Next & Char.Next.Next & Char.Next.Next.Next
            Case 5
                charText = Char & Char.Next & Char.Next.Next & Char.Next.Next.Next & Char.Next.Next.Next.Next
            Case 6
                charText = Char & Char.Next & Char.Next.Next & Char.Next.Next.Next & Char.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next
            Case 7
                charText = Char & Char.Next & Char.Next.Next & Char.Next.Next.Next & Char.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next
            Case 8
                charText = Char & Char.Next & Char.Next.Next & Char.Next.Next.Next & Char.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next
            Case 9
                charText = Char & Char.Next & Char.Next.Next & Char.Next.Next.Next & Char.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next.Next
            Case 10
                charText = Char & Char.Next & Char.Next.Next & Char.Next.Next.Next & Char.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next.Next.Next
            Case 11
                charText = Char & Char.Next & Char.Next.Next & Char.Next.Next.Next & Char.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next.Next.Next.Next
        End Select
        If Not charText Like "*[-'　 。，、；：？:,;,〈〉《》 ''「」『』（）▽△？！（）【】—""()<>" _
            & ChrW(9312) & Chr(-24153) & Chr(-24152) & ChrW(8218) & Chr(13) & Chr(10) & Chr(11) & ChrW(12) & Chr(63) & Chr(9) & Chr(-24152) & Chr(-24153) & "▽□】【~/︵—]*" _
            And Not charText Like "*[a-zA-Z0-9０-９]*" And InStr(charText, ChrW(-243)) = 0 And InStr(charText, Chr(91)) = 0 And InStr(charText, Chr(93)) = 0 Then
            'If Not charText Like "[a-z1-9]" & Chr(-24153) & Chr(-24152) & " 　、'""「」『』（）－？！]" Then
'            If InStr(Chr(-24153) & Chr(-24152) & Chr(2) & "•[]〔〕﹝﹞…；,，.。． 　、'""‘’`\{}｛｝「」『』（）《》〈〉－？！]", charText) = 0 Then
            If Not charText Like "*[" & ChrW(-24153) & ChrW(-24152) & Chr(2) & "•[]〔〕﹝﹞…；,，.。． 　、'""‘’`\{}｛｝「」『』（）《》〈〉－？！‘｛｝]*" Then
            'chr(2)可能是註腳標記
                If preChar <> charText Then
                    'If UBound(X) > 0 Then
                        If preChar = "" Then 'If IsEmpty(X) Then'如果是一開始
                            GoTo 1
                        ElseIf UBound(Filter(x, charText)) Then ' <> charText Then  '如果尚無此字
1                           ReDim Preserve x(i)
                            ReDim Preserve xT(i)
                            x(i) = charText
                            xT(i) = xT(i) + 1
                            i = i + 1
                        Else
                            GoSub 詞頻加一
                        End If
                    'End If
                Else
                    GoSub 詞頻加一
                End If
                preChar = charText
            End If
        End If
    Next
End With
12
Dim Doc As New Document, Xsort() As String, u As Long ', xTsort() As Integer, k As Long, so As Long, ww As String
If u = 0 Then u = 1 '若無執行「詞頻加一:」副程序,若無超過1次的詞頻，則　Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "、" & x(j - 1) & _
                                會出錯：陣列索引超出範圍 2015/11/5

ReDim Xsort(u) As String
Set xlApp = CreateObject("Excel.Application")
Set xlBook = xlApp.workbooks.Add
Set xlSheet = xlBook.Worksheets(1)
With xlSheet.Application
    For j = 1 To i
        .Cells(j, 1) = x(j - 1)
        .Cells(j, 2) = xT(j - 1)
        Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "、" & x(j - 1) 'Xsort(xT(j - 1)) & ww '陣列排序'2010/10/29
    Next j
End With
Doc.ActiveWindow.Visible = False
If d.ActiveWindow.View.ReadingLayout Then ReadingLayoutB = True: d.ActiveWindow.View.ReadingLayout = Not d.ActiveWindow.View.ReadingLayout
'U = UBound(Xsort)
For j = u To 0 Step -1 '陣列排序'2010/10/29
    If Xsort(j) <> "" Then
        With Doc
            If Len(.Range) = 1 Then '尚未輸入內容
                .Range.InsertAfter "詞頻 = " & j & "次：（" & Len(Replace(Xsort(j), "、", "")) / Ln & "個）"
                .Range.Paragraphs(1).Range.Font.Size = 12
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "新細明體"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
                '.Range.Paragraphs(1).Range.Font.Bold = True
            Else
                .Range.InsertParagraphAfter
                .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
                .Range.InsertAfter "詞頻 = " & j & "次：（" & Len(Replace(Xsort(j), "、", "")) / Ln & "個）"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Size = 12
                '.Range.Paragraphs(.Paragraphs.Count).Range.Bold = True
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "新細明體"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
            End If
            .Range.InsertParagraphAfter
            .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
            .Range.Paragraphs(.Paragraphs.Count).Range.Font.Size = 12
'            .Range.Paragraphs(.Paragraphs.Count).Range.Bold = False
            .Range.InsertAfter Replace(Xsort(j), "、", Chr(9), 1, 1) 'chr(9)為定位字元(Tab鍵值)
            .Range.InsertParagraphAfter
            If InStr(.Range.Paragraphs(.Paragraphs.Count).Range, "詞頻") = 0 Then
                .Range.Paragraphs(.Paragraphs.Count - 1).Range.Font.Name = "標楷體"
            Else
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "新細明體"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
            End If
        End With
    End If
Next j

With Doc.Paragraphs(1).Range
     .InsertParagraphBefore
     .Font.NameAscii = "times new roman"
    Doc.Paragraphs(1).Range.InsertParagraphAfter
    Doc.Paragraphs(1).Range.InsertParagraphAfter
    Doc.Paragraphs(1).Range.InsertAfter "你提供的文本共使用了" & i & "個不同的詞彙（傳統字與簡化字不予合併）"
End With

Doc.ActiveWindow.Visible = True

de = VBA.Timer
Doc.SaveAs Replace(xlsp, "XLS", "doc") '分大小寫
If ReadingLayoutB Then d.ActiveWindow.View.ReadingLayout = Not d.ActiveWindow.View.ReadingLayout
Set d = Nothing ' ActiveDocument.Close wdDoNotSaveChanges

Debug.Print Now

MsgBox "完成！" & vbCr & vbCr & "費時" & Left(de - ds, 5) & "秒!", vbInformation
xlSheet.Application.Visible = True
xlSheet.Application.UserControl = True
xlSheet.SaveAs xlsp
Exit Sub
詞頻加一:
For j = 0 To UBound(x)
    If x(j) = charText Then
        xT(j) = xT(j) + 1
        If u < xT(j) Then u = xT(j) '記下最高詞頻,以便排序(將欲排序之陣列最高元素值設為此,則不會超出陣列.
        '多此一行因為要重複判斷計算好幾次,故效能不增反減''效能還是差不多啦.
        Exit For
    End If
Next j

Return
ErrH:
Select Case Err.Number
    Case 4605 '閱讀模式不能編輯'此方法或屬性無法使用，因為此命令無法在閱讀中使用。
    '    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
    '        ActiveWindow.ActivePane.View.Type = wdNormalView
    '    Else
    '        ActiveWindow.View.Type = wdNormalView
    '    End If
    '    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
    '        ActiveWindow.ActivePane.View.Type = wdPrintView
    '    Else
    '        ActiveWindow.View.Type = wdPrintView
    '    End If
        'Doc.Application.ActiveWindow.View.ReadingLayout
        d.ActiveWindow.View.ReadingLayout = False ' Not d.ActiveWindow.View.ReadingLayout
        Doc.ActiveWindow.View.ReadingLayout = False
        Doc.ActiveWindow.Visible = False
        ReadingLayoutB = True
        Resume
    
    Case 91, 5941 '沒有設定物件變數或 With 區塊變數,集合中所需的成員不存在
        GoTo 12
    Case Else
        MsgBox Err.Number & Err.Description, vbCritical 'STOP: Resume
        Resume
        End
    
End Select
End Sub


Sub 書名號篇名號檢查()
Dim s As Long, rng As Range, e, trm As String, ans
Static x() As String, i As Integer
On Error GoTo eH
Do
    Selection.Find.Execute "〈", , , , , , True, wdFindAsk
    Set rng = Selection.Range
    rng.MoveEndUntil "〉"
    trm = Mid(rng, 2)
    
    For Each e In x()
        If StrComp(e, trm) = 0 Then GoTo 1
    Next e
2   ans = MsgBox("是否略過「" & trm & "」？" & vbCr & vbCr & vbCr & "結束請按 NO[否]", vbExclamation + vbYesNoCancel)
    Select Case ans
        Case vbYes
            ReDim Preserve x(i) As String
            x(i) = trm
            i = i + 1
        Case vbNo
            Exit Sub
    End Select
1
Loop
Exit Sub
eH:
Select Case Err.Number
    Case 92 '沒有設定 For 迴圈的初始值 陣列尚未有值
        GoTo 2
End Select
End Sub

Sub 時間軸單位轉換() '2017/5/13 因應YOUKU與YOUTUBE時間軸單位不同而設
'Debug.Print Len(ActiveDocument.Range)
Dim a, aM, aMM, s As Long, e As Long
Dim myRng As Range, chRng As Range
Set myRng = ActiveDocument.Range
Set chRng = ActiveDocument.Range
s = -1
For Each a In ActiveDocument.Characters
    If a.Font.Name = "Times New Roman" Then
        If s = -1 Then s = a.start
        If a = Chr(13) Then GoTo 1
    Else
1       If s > -1 Then
            e = a.Previous.End
            myRng.SetRange s, e
            If InStr(myRng, "http") = 0 Then
                If InStr(Replace(myRng, ":", "", 1, 1), ":") Then 'if find : * 2
                    If InStr(Trim(myRng), " ") Then '如果有2個以上時間軸
                        For Each aMM In myRng.Characters
                            If aMM.Next = " " Then
                                e = aMM.End
                                chRng.SetRange s, e
'                                chRng.Select
                                If InStr(Replace(chRng, ":", "", 1, 1), ":") Then 'if find : * 2
                                    GoSub chng
                                End If
                                s = chRng.End + 1
                            End If
                        Next
                    Else '如果只有1個時間軸
                        chRng.SetRange myRng.start, myRng.End
                        GoSub chng
                    End If
                End If
            End If
            s = -1
        End If
    End If
Next
ActiveDocument.Range.Find.Execute "  ", True, , , , , , wdFindContinue, , " ", wdReplaceAll
Exit Sub
chng:
                    For Each aM In chRng.Characters
                        If aM.Next = ":" Then
                            aM.Next.Next.Text = Str((CInt(aM.Next.Next) * 10 + CInt(aM) * 60) / 10)
                            aM.Next.Delete
                            aM.Delete
                            Exit For
                        End If
                    Next
Return
End Sub
Sub 中國哲學書電子化計劃_表格轉文字(ByRef r As Range)
On Error GoTo eH
Dim lngTemp As Long '因為誤按到追蹤修訂，才會引發訊息提示刪除儲存格不會有標識
'Dim d As Document
Dim tb As Table, c As Cell ', ci As Long
'Set d = ActiveDocument
lngTemp = word.Application.DisplayAlerts
If r.Tables.Count > 0 Then
    For Each tb In r.Tables
        'tb.Columns(1).Delete
        Err.Raise 5992
        Set r = tb.ConvertToText()
    Next tb
End If
'word.Application.DisplayAlerts = lngTemp
Exit Sub
eH:
Select Case Err.Number
    Case 5992 '無法個別存取此集合中的各欄，因為表格中有混合的儲存格寬度。
        For Each c In tb.Range.Cells
'            ci = ci + 1
'            If ci Mod 3 = 2 Then
                'If VBA.IsNumeric(VBA.Left(c.Range.text, VBA.InStr(c.Range.text, "?") - 1)) Then
                If VBA.InStr(c.Range.Text, ChrW(160) & ChrW(47)) > 0 Then
'                    word.Application.DisplayAlerts = False
                    c.Delete  '刪除編號之儲存格
                End If
'            End If
        Next c
        Resume Next
    Case Else
        MsgBox Err.Number & Err.Description
        End
End Select
End Sub

Sub 中國哲學書電子化計劃_註文變小正文回大()
Dim slRng As Range, a
Set slRng = Selection.Range
中國哲學書電子化計劃_表格轉文字 slRng
For Each a In slRng.Characters
    Select Case a.Font.Color
        Case 34816, 8912896
            a.Font.Size = 14
        Case 0
            a.Font.Size = 30
    End Select
Next a
End Sub
Sub 中國哲學書電子化計劃_去掉註文保留正文()
Dim slRng As Range, a
Docs.空白的新文件
If ActiveDocument.Characters.Count = 1 Then Selection.Paste
If Selection.Type = wdSelectionIP Then ActiveDocument.Select
Set slRng = Selection.Range
中國哲學書電子化計劃_表格轉文字 slRng
For Each a In slRng.Characters
    Select Case a.Font.Color
        Case 34816, 8912896
            If a.Font.Size <> 12 Then Stop
            a.Delete
        Case 254
            If a.Font.Size = 9 Then a.Delete
    End Select
Next a
If MsgBox("是否取代異體字？", vbOKCancel) = vbOK Then 文字轉換.異體字轉正
Beep 'MsgBox "done!", vbInformation
End Sub
Sub 中國哲學書電子化計劃_註文前後加括弧()
Dim slRng As Range, a, flg As Boolean 'Alt+1
Docs.空白的新文件
If Selection.Type = wdSelectionIP Then ActiveDocument.Select
Set slRng = Selection.Range
中國哲學書電子化計劃_表格轉文字 slRng
For Each a In slRng.Document.Paragraphs 'for漢籍電子文獻資料庫
    If VBA.Left(a.Range, 3) = "[疏]" Then
        slRng.SetRange a.Range.Characters(4).start _
            , a.Range.End
        slRng.Font.Size = 7.5
    End If
Next a
If Selection.Type = wdSelectionIP Then ActiveDocument.Select
Set slRng = Selection.Range
For Each a In slRng.Characters
    Select Case a.Font.Color
        Case 34816, 8912896, 15776152 '34816:綠色小注
p:          If flg = False Then
                a.Select
                Selection.Range.InsertBefore "（"
                Selection.Range.SetRange Selection.start, Selection.start + 1
                Selection.Range.Font.Size = a.Characters(2).Font.Size
                Selection.Range.Font.Color = a.Characters(2).Font.Color
'                a.Font.Size = a.Next.Font.Size
'                a.Font.Color = a.Next.Font.Color
                flg = True
            Else
                If a.Font.Color = 8912896 And a.Previous.Font.Color = 34816 Then '8912896藍字小注
                    a.InsertBefore "）（"
                    a.SetRange a.start, a.start + 2
                    a.Font.Size = a.Characters(2).Next.Font.Size
                    a.Font.Color = a.Characters(2).Next.Font.Color
'                    a.Characters(1).Font.Color = a.Characters(1).Previous.Font.Color
                End If
            End If
'        Case 8912896 '8912896藍字小注
            
        Case 0, 15595002, 15649962
            If a.Font.Color = 0 Then 'black'漢籍電子文獻資料庫
                If a.Font.Size = 7.5 And Not flg Then
                    GoTo p
                ElseIf a.Font.Size > 7.5 And flg Then
                    GoTo b
                End If
            'End If
            ElseIf flg Then
b:
'                a.Select
'                Selection.Range.InsertBefore "）"
                If a.Previous = Chr(13) Then
                    a.Previous.Previous.Select
                Else
                    a.Previous.Select
                End If
                Selection.Range.InsertAfter "）"
                flg = False
            End If
        Case -16777216 'black'漢籍電子文獻資料庫
            If a.Font.Size = 7.5 And Not flg Then
                GoTo p
            ElseIf a.Font.Size > 7.5 And flg Then
                GoTo b
            End If
        Case 255 'red'漢籍電子文獻資料庫
            Select Case a.Font.Size
                Case 7.5, 10
                    a.Delete
            End Select
    End Select
Next a
slRng.Find.Execute "（（", True, , , , , , , , "（", wdReplaceAll
slRng.Find.Execute "））", True, , , , , , , , "）", wdReplaceAll
Beep
Selection.EndKey wdStory
Do
   Selection.MoveLeft
   If Selection = Chr(13) Then Selection.Delete
Loop While Selection = Chr(13)
'MsgBox "done!", vbInformation
End Sub
Sub 漢籍電子文獻資料庫文本整理_以轉貼到中國哲學書電子化計劃(Optional doNotCloseDoc As Boolean)
Dim rng As Range, d As Document, a
Dim rp As Variant, i As Byte
If Documents.Count = 0 Then Documents.Add
Set d = ActiveDocument
If d.path <> "" Or d.Content.Text <> Chr(13) Then
    Set d = Documents.Add()
    'Exit Sub
End If
rp = Array("(", "{{", ")", "}}", ChrW(160), "", "【圖】", "", _
     "^p^p", "^p", _
     ChrW(13) & ChrW(45) & ChrW(13) & ChrW(13) & ChrW(11), "^p", _
     ChrW(13) & ChrW(45) & ChrW(13), "^p", "{{ }}", "", "[", ChrW(12310), _
     "]", ChrW(12311), " ", "", "○", ChrW(12295), _
     "^p" & ChrW(12310) & "疏" & ChrW(12311), ChrW(12310) & "疏" & ChrW(12311) & "{{", _
     "}}" & Chr(13) & "^#" & Chr(13) & "{{", "", _
     "．．．．．．．．．．．．．．．．．．" & Chr(13), "", _
     Chr(13) & "^#" & Chr(13), "", _
     "}}" & Chr(13) & "^#" & Chr(13), "}}", _
     "}}" & Chr(13) & "{{", "", _
     "-", "", "^#", "", "。。", "。") ', "。}}<p>。}}<p>", "。}}<p>")
     '原來「ChrW(13) & ChrW(45) & ChrW(13) & ChrW(13) & ChrW(11)」是其中有表格啊
Set rng = d.Range
rng.Paste
漢籍電子文獻資料庫文本整理_注文前後加括號
For Each a In rng.Characters
    If a.Font.Size = 10 Then
        Select Case a.Font.Color
            Case 255, 9915136
                a.Delete
        End Select
    End If
Next a
rng.Cut
rng.PasteAndFormat wdFormatPlainText
rng.Find.ClearFormatting
For i = 0 To UBound(rp)
    rng.Find.Execute rp(i), , , , , , , wdFindContinue, , rp(i + 1), wdReplaceAll
    i = i + 1
Next i
文字處理.書名號篇名號標注
Beep
If Not doNotCloseDoc Then
    d.Range.Cut
    d.Close wdDoNotSaveChanges
End If
End Sub
Sub 漢籍電子文獻資料庫文本整理_注文前後加括號()
Dim rng As Range, fColor As Long, flg As Boolean
Const fSize As Byte = 10
Set rng = ActiveDocument.Range
rng.Collapse wdCollapseStart
fColor = rng.Font.Color
Do While rng.End < rng.Document.Range.End - 1
    rng.move wdCharacter, 1
    If rng.Font.Color = 204 And rng.Font.Size = 11 Then
        rng.Delete
    ElseIf rng.Font.Color = 0 And rng.Font.Size = 7.5 Then
        GoTo mark
    ElseIf (rng.Font.Color <> fColor Or rng.Font.Size = fSize) And _
                (rng.Font.Color <> 234 And rng.Font.Bold = False) Then '紅字+粗體為檢索結果
mark:
        If flg = False Then
            If rng.Font.Color <> -16777216 Then
                rng.InsertBefore "("
                rng.Characters(1).Font.Color = rng.Next.Next.Font.Color
                rng.Characters(1).Font.Size = rng.Next.Next.Font.Size
                flg = True
            End If
        End If
    ElseIf rng.Font.Color = fColor And flg = True Then
        rng.Previous.InsertAfter ")"
        flg = False
    End If
Loop
Beep
End Sub
Sub 詩句分行()
Dim slRng As Range, a
Set slRng = Selection.Range
For Each a In slRng.Characters
    If a Like "[。，；？！「」『』]" Then
        a.Select
        Selection.move
        Selection.TypeText Chr(11)
    End If
Next a
End Sub

Sub 刪除校案語()
Dim rng As Range, e, d As Document
Set d = ActiveDocument
Set rng = d.Range
e = rng.End
With rng.Find
    .Style = "超連結"
    .Execute , , , , , , , wdFindStop ', , "" ', wdReplaceAll
    Do
        If InStr(rng.Characters(rng.Characters.Count).Next.Style, "校案") _
            Or InStr(rng.Characters(1).Previous.Style, "校案") Then
            rng.Select
            Selection.Delete
            rng.SetRange Selection.start, e
        End If
    Loop While .Execute(, , , , , , , wdFindStop)  ', , "" ', wdReplaceAll
End With

With rng.Find
    .Style = "校案"
    .Execute , , , , , , , wdFindContinue, , "", wdReplaceAll
End With
With rng.Find
    .Style = "校案引文"
    .Execute , , , , , , , wdFindContinue, , "", wdReplaceAll
End With
Beep
End Sub

Function 國語辭典注音文字處理(x As String)
Dim ay, i As Byte
ay = Array("ㄧ", ChrW(20008), "　", " ", "（又音）", "又音 ", "（讀音）", "讀音 ", "（語音）", "語音 ", _
        "(一)", "", "(二)", "", "(三)", "", "(四)", "", "(五)", "", "(六)", "", "）", "", "（", "")
For i = 0 To UBound(ay)
    x = Replace(x, ay(i), ay(i + 1))
    i = i + 1
Next i
國語辭典注音文字處理 = x
End Function
Sub 生難字加上國語辭典注音()
Dim rng As Range, x, rst As New ADODB.Recordset, st As WdSelectionType, words As String
Dim cnt As New ADODB.Connection, id As Long, sty As word.Style, url As String
Dim frmDict As New Form_DictsURL, lnks As New Links, db As New dBase ', frm As New MSForms.DataObject
Static cntStr As String, chromePath As String
st = Selection.Type
If st = wdSelectionIP Then
    If Selection.start = 0 Then Exit Sub
    x = Selection.Previous.Characters(Selection.Previous.Characters.Count).Text
    If InStr("。，；「」『』〈〉《》？.,;""?－-──--（）()【】〔〕<>[]…! 　！", x) Then Exit Sub
    If 文字處理.isSymbol(CStr(x)) Or 文字處理.is注音符號(CStr(x)) Or 文字處理.isLetter(CStr(x)) Or 文字處理.isNum(CStr(x)) Then Exit Sub
'    Selection.Previous.Copy
Else
    x = trimStrForSearch(VBA.CStr(Selection.Text), Selection)
    'Selection.Copy
    SystemSetup.ClipboardPutIn "=" & Selection.Text
End If
Set rng = Selection.Range
words = x
db.setWordControlValue (words)
On Error GoTo eH
Dim ur As UndoRecord
Set ur = SystemSetup.stopUndo("生難字加上國語辭典注音")

'If Not Selection.Document.path = "" Then If Not Selection.Document.Saved Then Selection.Document.save
If cntStr = "" Then
    Dim dbp As New Paths
    cntStr = dbp.getdb_重編國語辭典修訂本_資料庫fullName
End If

If chromePath = "" Then
    chromePath = SystemSetup.getChrome
End If

'Dim ay, i As Byte
'ay = Array("ㄧ", ChrW(20008), "　", " ", "（又音）", "又音 ", "（讀音）", "讀音 ", "（語音）", "語音 ", _
'        "(一)", "", "(二)", "", "(三)", "", "(四)", "", "(五)", "", "(六)", "", "）", "", "（", "")

    cnt.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & cntStr
'Exit Sub
'cntt:
    rst.Open "select 注音一式,釋義,url,ID from [《重編國語辭典修訂本》 總表] where strcomp(字詞名,""" & x & """)=0 order by 注音一式", cnt, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        GoSub list
    Else
        生難字加上國語辭典注音nextTable rst, cnt, x, "《重編國語辭典修訂本》 總表-20210928以前", True
        If rst.RecordCount > 0 Then
            GoSub list
        Else
2
            If Selection.Characters.Count = 1 Then 'words  單字
                frmDict.getDictVariantsRecS words, rst
                If rst.RecordCount > 0 Then
                    GoSub list
                Else
                    frmDict.getDictHydzdRecS words, rst
                    If rst.RecordCount > 0 Then
                        'GoSub list
                        Set sty = rng.Style
                        rng.Hyperlinks.Add rng, lnks.trimLinks(rst.Fields(2).Value), , , , "_blank"
                        lnks.setStylewithHyperlinkMark sty, rng
                    Else
                        GoSub notFound
                    End If
                End If
            Else 'terms 詞彙
                frmDict.getDictHydcdRecS words, rst
                If rst.RecordCount > 0 Then
                    If Not VBA.IsNull(rst.Fields(0)) Then
                        GoSub list
                    Else
                        Set sty = rng.Style
                        rng.Hyperlinks.Add rng, lnks.trimLinks(rst.Fields(2).Value), , , , "_blank"
                        lnks.setStylewithHyperlinkMark sty, rng
                    End If
                Else
                    GoSub notFound
                End If
            End If
        End If
    End If
endS:
    SystemSetup.contiUndo ur
    Set ur = Nothing
    If rst.State <> adStateClosed Then rst.Close
    If cnt.State <> adStateClosed Then cnt.Close
    Set rst = Nothing: Set cnt = Nothing: Set frmDict = Nothing ': Set frm = Nothing
    Set lnks = Nothing: Set db = Nothing: Set rng = Nothing
Exit Sub

notFound:
                If st = wdSelectionIP Then
                    Selection.Previous.Copy
                    'Selection.Document.FollowHyperlink "https://dict.variants.moe.edu.tw/variants/rbt/query_by_standard_tiles.rbt?command=clear"
                    x = frmDict.add1URLTo1異體字字典(words)
                    If x = "" Then GoTo endS
                    GoTo 2
                Else
                    rst.Close
                    rst.Open "select 注音一式,釋義,url,ID from [《重編國語辭典修訂本》 總表] where instr(字詞名,""" & x & """)>0 order by 注音一式", cnt, adOpenKeyset, adLockOptimistic
                    Selection.Copy
                    If rst.RecordCount > 0 Then
                        Beep
                        'Selection.Document.FollowHyperlink "https://www.zdic.net/hans/" & x, , True
                        Shell chromePath & " https://www.zdic.net/hans/" & x
                        GoSub list
                        'Selection.Document.FollowHyperlink "http://dict.revised.moe.edu.tw/cbdic/search.htm", , True
'                        Shell chromePath & " http://dict.revised.moe.edu.tw/cbdic/search.htm"
                    Else
                            生難字加上國語辭典注音nextTable rst, cnt, x, "《重編國語辭典修訂本》 總表-20210928以前", False
                            If rst.RecordCount > 0 Then
                                Beep
                                GoSub list
                            Else
                            'Selection.Document.FollowHyperlink "https://www.zdic.net/hans/" & x, , True
                            Shell chromePath & " https://www.zdic.net/hans/" & x
                            End If
                    End If
                End If
Return

list:
'        Dim ur As UndoRecord
'        Set ur = SystemSetup.stopUndo("沛榮按")
'        Docs.樣式add_沛榮按等樣式
        rng.Collapse wdCollapseEnd
        If rng.Style <> "沛榮按" Then
            rng.InsertAfter "（）"
            rng.Style = "沛榮按"
            rng.SetRange rng.End - 1, rng.End - 1
        End If
        Do Until rst.EOF
            x = ""
            If VBA.IsNull(rst.Fields(0).Value) Then
                x = rst.Fields(1).Value '釋義
            Else
                x = rst.Fields(0).Value '注音
            End If
            GoSub typeTexts
            rst.MoveNext
        Loop
        If rng.Previous = "，" Then rng.Previous.Delete
'        SystemSetup.contiUndo ur
'        Set ur = Nothing:  'Set frm = Nothing: Set frmDict = Nothing

Return

typeTexts:
        If x = "" Or VBA.IsNull(x) Then GoTo 2
'        X = Mid(X, 1, Len(X) - 1)
        x = 國語辭典注音文字處理(CStr(x))
'        If sT <> wdSelectionIP Then
'            rng.SetRange Selection.End, Selection.End
'        End If
'        rng.SetRange rng.End - 1, rng.End - 1
        rng.InsertAfter x 'insert ZhuYin
        For Each x In rng.Characters 'format ZhuYin
            If InStr("ˊˇˋ", x) Then
                x.Style = "聲調"
            ElseIf InStr("˙", x) Then
                x.Font.Name = "標楷體"
            End If
        Next x
        x = rst.Fields(2).Value 'URL  'frmDict.get1URLfor1(words)
        If VBA.IsNull(x) Then
                If st = wdSelectionIP Then
                    If Selection.Previous.Characters(Selection.Previous.Characters.Count).Hyperlinks.Count > 0 Then
                        Dim rngW As Range
                        Set rngW = Selection.Range
                        rngW.SetRange Selection.Previous.Characters(Selection.Previous.Characters.Count).start, Selection.Previous.Characters(Selection.Previous.Characters.Count).End
                        SystemSetup.ClipboardPutIn "=" & rngW.Text '"^" & rngW.text & "$" 'version 6's new settings
                        Set rngW = Nothing
                    Else
                        Set rngW = Selection.Previous.Characters(Selection.Previous.Characters.Count)
                        SystemSetup.ClipboardPutIn "=" & rngW.Text
                        'Selection.Previous.Characters(Selection.Previous.Characters.Count).Copy
                    End If
                End If
'                Shell chromePath & " http://dict.revised.moe.edu.tw/cbdic/search.htm"
'            frm.Clear
'            frm.SetText words, 1
'            frm.PutInClipboard
            'add new url
rePt:
            If rst.RecordCount = 1 Then
                If Not SystemSetup.appActivatedYet("chrome") Then
                'If Not word.Tasks.Exists("google chrome") Then
                    Shell SystemSetup.getChrome & " https://dict.revised.moe.edu.tw/search.jsp?md=1"
                Else
'                    AppActivate "google chrome" 'https://docs.microsoft.com/zh-tw/sql/ado/reference/ado-api/absoluteposition-property-ado?view=sql-server-ver15
                        'try looking for both Chrome_WidgetWin_1 and Chrome_RenderWidgetHostHWND
                    SystemSetup.apicShowWindow "Chrome_WidgetWin_1" _
                        , vbNullString, 3 'https://zechs.taipei/?p=146
                        'https://docs.microsoft.com/zh-tw/visualstudio/debugger/how-to-start-spy-increment?view=vs-2022
                    'https://stackoverflow.com/questions/19705797/find-the-window-handle-for-a-chrome-browser
                End If
            Else
                Beep
            End If
            x = InputBox("plz putin the url", , IIf(VBA.IsNull(rst.Fields(0).Value), "", rst.Fields(0).Value)) 'frmDict.add1URLTo1國語辭典(words)
            If x = "" Then GoTo endS
            If Left(x, 4) <> "http" Then GoTo rePt
            x = lnks.trimLinks_http_Dicts_toAddZhuYin_RevisedMoeEdu(CStr(x), rst.Fields(0))
            url = VBA.CStr(x)
            If lnks.chkLinks_http_Dicts_toAddZhuYin(url, words, 1, id, rst.Fields(0)) Then
                x = url
                rst.Fields(2).Value = x
                If id <> 0 Then
                    rst.Fields("ID") = id
                    id = 0
                End If
                rst.Update
                db.setURLControlValue VBA.CStr(x)
            Else
                GoTo endS 'Exit Sub
            End If
        End If
        Set sty = rng.Style
        rng.Hyperlinks.Add rng, lnks.trimLinks(VBA.CStr(x)), , , , "_blank"
        lnks.setStylewithHyperlinkMark sty, rng
        rng.Collapse wdCollapseEnd
        'rng.SetRange rng.End, rng.End
        rng.Next.InsertBefore "，"
'        rng.Style = "沛榮按"
        'rng.Hyperlinks.Item(1).Delete
        'rng.Collapse wdCollapseEnd
        rng.SetRange rng.End + 2, rng.End + 2
Return


eH:
    Select Case Err.Number
        Case 4198 '指令失敗 'Google Drive的問題
            Resume Next
        Case 5834 '指定名稱的項目不存在
            Docs.樣式add_沛榮按等樣式
            Resume
        Case Else
            MsgBox Err.Number & Err.Description
'            Resume
            GoTo endS
            'If cnt.State <> adStateClosed Then cnt.Close
    End Select
End Sub
Sub 生難字加上國語辭典注音nextTable(ByRef rst As ADODB.Recordset, ByRef cnt As ADODB.Connection, x, tbName As String, precise As Boolean)
    If rst.State = adStateOpen Then rst.Close
    Dim src As String
    Dim srcs As String
    srcs = "select 注音一式,釋義,url,ID from [" & tbName & "] where "
    If precise Then
        src = "strcomp(字詞名,""" & x & """)=0"
    Else
        src = "instr(字詞名,""" & x & """)>0"
    End If
    rst.Open srcs & src, cnt, adOpenKeyset
End Sub


Sub 書名號篇名號標注()
Dim cnt As New ADODB.Connection, rst As New ADODB.Recordset
Dim cntStr As String, d As Document, dx As String, rngF As Range
Dim db As New dBase
db.cnt查字 cnt
'If Dir("H:\我的雲端硬碟\私人\千慮一得齋(C槽版)\書籍資料\圖書管理附件", vbDirectory) <> "" Then
'    cntStr = "Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=H:\我的雲端硬碟\私人\千慮一得齋(C槽版)\書籍資料\圖書管理附件\查字.mdb;"
'ElseIf Dir("D:\千慮一得齋\書籍資料\圖書管理附件", vbDirectory) <> "" Then
'    cntStr = "Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=D:\千慮一得齋\書籍資料\圖書管理附件\查字.mdb;"
'Else
'    MsgBox "路徑不存在！", vbCritical: Exit Sub
'End If
Set d = ActiveDocument: dx = d.Range.Text: Set rngF = d.Range
'cnt.Open cntStr

GoSub bookmarks '標點符號_書名號_自動加上用
rst.Open "select * from 標點符號_篇名號_自動加上用 order by 排序", cnt
Set rngF = d.Range: dx = d.Range.Text
Do Until rst.EOF
    If VBA.InStr(dx, rst("篇名").Value) Then 'if found
        Do While rngF.Find.Execute(rst("篇名").Value, , , , , , True, wdFindStop)
            If InStr("》〉·•", IIf(rngF.Characters(rngF.Characters.Count).Next Is Nothing, "", rngF.Characters(rngF.Characters.Count).Next)) = 0 And _
                InStr("《〈·•", IIf(rngF.Characters(1).Previous Is Nothing, "", rngF.Characters(1).Previous)) = 0 Then
                If VBA.IsNull(rst("取代為").Value) Then
                    rngF.Text = "〈" & rst("篇名").Value & "〉"
                              'd.Range.Find.Execute rst("篇名").Value, , , , , , True, wdFindContinue, , "〈" & rst("篇名").Value & "〉", wdReplaceAll
                Else
                    rngF.Text = rst("取代為").Value
                    'd.Range.Find.Execute rst("篇名").Value, , , , , , True, wdFindContinue, , rst("取代為").Value, wdReplaceAll
                End If
            End If
        Loop
    End If
    Set rngF = d.Range: dx = d.Range.Text
    rst.MoveNext
Loop
d.Range.Find.Execute "《《", , , , , , True, wdFindContinue, , "《", wdReplaceAll
d.Range.Find.Execute "》》", , , , , , True, wdFindContinue, , "》", wdReplaceAll
d.Range.Find.Execute "〈〈", , , , , , True, wdFindContinue, , "〈", wdReplaceAll
d.Range.Find.Execute "〉〉", , , , , , True, wdFindContinue, , "〉", wdReplaceAll

'GoSub bookmarks 'do again to check and correct SHOULD BE use another table to do this
rst.Close: cnt.Close
Exit Sub
bookmarks:
If rst.State = adStateOpen Then rst.Close
rst.Open "select * from 標點符號_書名號_自動加上用 order by 排序", cnt
Do Until rst.EOF
    If VBA.InStr(dx, rst("書名").Value) Then 'if found
        Do While rngF.Find.Execute(rst("書名").Value, , , , , , True, wdFindStop)
            If InStr("》〉·•", IIf(rngF.Characters(rngF.Characters.Count).Next Is Nothing, "", rngF.Characters(rngF.Characters.Count).Next)) = 0 And _
                InStr("《〈·•", IIf(rngF.Characters(1).Previous Is Nothing, "", rngF.Characters(1).Previous)) = 0 Then
                If VBA.IsNull(rst("取代為").Value) Then
                    rngF.Text = "《" & rst("書名").Value & "》"
        '            d.Range.Find.Execute rst("書名").Value, , , , , , True, wdFindContinue, , "《" & rst("書名").Value & "》", wdReplaceAll
                Else
                    rngF.Text = rst("取代為").Value
        '            d.Range.Find.Execute rst("書名").Value, , , , , , True, wdFindContinue, , rst("取代為").Value, wdReplaceAll
                End If
            End If
        Loop
    End If
    Set rngF = d.Range: dx = d.Range.Text
    rst.MoveNext
Loop
rst.Close
Return

End Sub

Sub 分行分段_根據第1行的字數長度來作切割()
Dim wordCount As Byte, d As Document, rng As Range, i As Integer, dx As String, a, p As Paragraph, j As Byte, wl
Dim omitStr As String
omitStr = "{}<p>《》〈〉：，。「」『』　·0123456789-" & ChrW(8231) & ChrW(183) & Chr(13)
If word.Documents.Count = 0 Then
    Set d = Documents.Add()
ElseIf ActiveDocument.path <> "" Then
    Set d = Documents.Add() 'ActiveDocument
Else
    Set d = ActiveDocument
End If
Set rng = d.Range
rng.Paste
Set p = rng.Paragraphs(1)
'wordCount = p.Range.Characters.Count - 1
For Each a In p.Range.Characters
    If InStr(omitStr, a) = 0 Then wordCount = wordCount + 1
Next a
dx = rng.Text
wl = InStr(dx, Chr(13))
rng.Text = Left(dx, wl) & Replace(dx, Chr(13), "", wl)

i = 1
Do Until rng.Paragraphs(rng.Paragraphs.Count).Range.Characters.Count < wordCount
    i = i + 1
    If i > rng.Paragraphs.Count Then Exit Do
    Set p = rng.Paragraphs(i)
    For Each a In p.Range.Characters
        If InStr(omitStr, a) = 0 Then j = j + 1
        If j = wordCount Then
            a.InsertAfter Chr(13)
            j = 0
            Exit For
        End If
    Next a
'    rng.Paragraphs(i).Range.Characters(wordCount).InsertAfter Chr(13)
Loop
rng.Cut
rng.Document.Close wdDoNotSaveChanges
If word.Documents.Count = 0 Then
    word.Application.Quit
Else
    word.ActiveWindow.WindowState = wdWindowStateMinimize
End If
Beep
End Sub

Sub 國語辭典網址及ID尚缺者列出()
Dim db As New dBase
db.國語辭典網址及ID尚缺者列出
SystemSetup.playSound 12
End Sub
Sub 國語辭典網址及ID尚缺者填入()
Dim i As Long
ActiveDocument.Range.Find.Execute Chr(13), , , , , , , wdFindContinue, , "", wdReplaceAll
Do Until Selection.End = ActiveDocument.Range.End - 1
    Selection.move
    If Selection.Previous <> ChrW(20008) And Selection.Hyperlinks.Count = 0 Then
        生難字加上國語辭典注音
        ActiveWindow.ScrollIntoView Selection, False
        i = i + 1
    End If
    If i = 40 Then Exit Sub
Loop
Selection.HomeKey wdStory, wdExtend
End Sub
