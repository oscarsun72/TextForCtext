VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Dim db As DAO.Database
Sub 在整頁模式中不顯示頁面間的空白區域()
    If word.Documents.Count = 0 Then Exit Sub
    If word.Documents(1).Windows.Count = 0 Then Exit Sub
    If word.Documents(1).Windows(1).Visible Then
        If word.ActiveWindow.View.ReadingLayout Then _
            word.ActiveWindow.View.ReadingLayout = Not word.ActiveWindow.View.ReadingLayout
        DoEvents
        word.ActiveWindow.View.DisplayPageBoundaries = False
    End If
End Sub

Sub Word資料匯入Access(wcontent As String, Terms As String, beforeTerm As String, acpath As String, before部首, selRange As Range) '2010/10/20修訂
'Dim w As Word.Document
Dim 部首ID As Long '2010/10/29
Dim rst As DAO.Recordset, rst1 As Recordset, ss As String, 字ID As Long ', 字new As Boolean
Dim 不輸入注音 As Boolean
Static 不輸入字序 As Byte, 不覆蓋原資料 As Byte
On Error GoTo ErrH:
'Set w = GetObject("C:\Documents and Settings\Superwings\桌面\0001（定稿）_詞典.mdb參照.doc")
'wcontent = w.Application.Run(ThisDocument.詞典條目內容擷取)
ss = VBA.right(wcontent, 4)
wcontent = left(wcontent, Len(wcontent) - 4)
If db Is Nothing Then
    Set db = DBEngine.Workspaces(0).OpenDatabase(acpath) '("D:\千慮一得齋\書信\圖書管理\黃沛榮老師助理\詞典.mdb")
End If
'Terms = Mid(wcontent, InStr(wcontent, "【") + 1, InStr(wcontent, "】") - (InStr(wcontent, "【") + 1))
Set rst = db.OpenRecordset("字") 'If beforeTerm = "" Then beforeTerm = Terms已在所呼叫之"詞典條目內容擷取"程序判斷
If ss <> "XXX部" Then  '取得字ID值'ss <> "XXX字"
    If Len(beforeTerm) > 1 Then MsgBox "字的長度大於1,請程式撰寫者檢查!", vbExclamation: Stop
    With rst
        .Index = "字"
        .Seek "=", beforeTerm
        If .NoMatch Then
            'MsgBox "請程式撰寫者檢查!", vbExclamation: Stop
            'GoTo a
            .AddNew
            .Fields("字") = beforeTerm
            .Fields("釋文") = wcontent
            If before部首 <> "" Then
                Set rst1 = db.OpenRecordset("部首")
                With rst1
                    .Index = "部首"
                    .Seek "=", before部首
                    If .NoMatch Then
                        MsgBox "請程式撰寫者檢查!", vbExclamation: Stop '因為214部首已入,若有找不到,可疑!
                    Else
                        rst.Fields("部首ID") = .Fields("部首ID") ' DLookup("部首ID", "部首", "部首 = """ & before部首 & """")
                        部首ID = .Fields("部首ID")
                    End If
                    .Close
                End With
            End If
            .Update
            .Index = "字" '新增後一定要再找才有,或重開資料表才行.
            .Seek "=", beforeTerm '為了取得.Fields("編輯時間")及.Fields("字ID")的值
            'Set rst = DB.OpenRecordset("select 字ID from 字 where 字 = """ & Terms & """")'今有前式,不必再此式了
            字ID = .Fields("字ID") 'rst.Fields("字ID")
        Else
            字ID = .Fields("字ID") 'DLookup("字ID", "字", "字 like """ & beforeTerm & """")
        End If
        '檢查編輯更新情況
        If ss <> "XXX詞" Then '詞的處理在下面!2010/10/29
            If Not IsNull(.Fields("編輯時間")) Then
                Select Case 不覆蓋原資料
                    Case 0
                        不覆蓋原資料 = MsgBox("『" & .Fields("字") & "』字在" & .Fields("編輯時間") & "曾經編輯更新過" & _
                            "是否匯入?--匯入將會覆蓋<詞典>資料內原『" & .Fields("字") & "』字已經編輯過之資料" _
                            & vbCr & vbCr & "若略過,可記下此字,在此次匯入完成後," & _
                            "再回<詞典>資料庫檢索檢視" & vbCr & vbCr & "要匯入覆蓋原資料,請按[是]!! 要略過" _
                            & "請按[否]!  這次匯入一律不覆蓋請按[取消]", vbYesNoCancel + vbExclamation, "覆蓋原資料確認!!!")
                        Select Case 不覆蓋原資料
                            Case vbCancel
                                不覆蓋原資料 = vbCancel
                                GoTo 不匯入
                            Case vbNo
                                不覆蓋原資料 = vbNo
                                GoTo 不匯入
                            Case vbYes
                                不覆蓋原資料 = vbYes
                        End Select
                    Case vbNo
                        不覆蓋原資料 = MsgBox("『" & .Fields("字") & "』字在" & .Fields("編輯時間") & "曾經編輯更新過" & _
                            "是否匯入?--匯入將會覆蓋<詞典>資料內原『" & .Fields("字") & "』字已經編輯過之資料" _
                            & vbCr & vbCr & "若略過,可記下此字,在此次匯入完成後," & _
                            "再回<詞典>資料庫檢索檢視" & vbCr & vbCr & "要匯入覆蓋原資料,請按[是]!! 要略過" _
                            & "請按[否]!  這次匯入一律不覆蓋請按[取消]", vbYesNoCancel + vbExclamation, "覆蓋原資料確認!!!")
                        If 不覆蓋原資料 <> vbYes Then GoTo 不匯入
                    Case vbCancel
                        GoTo 不匯入
                End Select
            Else
            End If
        End If
    End With
End If
'If Terms = "數" Then Stop '有三聲（多音）尚未抓到
'If Terms = "吧" Then Stop '有輕聲尚未抓到
Select Case ss
    Case "XXX字"
        Set rst = db.OpenRecordset("字") '原此行為a:標籤
        Dim selrangeP As Paragraph, zhu, Uzhu As Byte, ZhuYin, char As Long, zhuyinText As String, pinyinText As String, charItem As Integer '輸入注音與拼音'2010/10/23
        Dim Zhuyins() As String, Pinyins() As String, zps As Byte '一字多音時, 用陳列儲存
        Dim 注音拼音rst As Recordset, 注音ID As Long, 拼音ID As Long
        zhu = 注音符號: Uzhu = UBound(zhu): char = 1
        For Each selrangeP In selRange.Paragraphs
            charItem = 1: char = charItem
            For Each ZhuYin In zhu
'                If ZhuYin = ChrW(20008) Then Stop
                If char > selrangeP.Range.Characters.Count Then Exit For '如果找完本段,即不再找,換下一段
                If InStr(selrangeP.Range, ZhuYin) Then  '找出要處理的段落-即段落中含注音的(必含拼音)
                    selrangeP.Range.SetRange char, selrangeP.Range.Characters.Count '重新取得range範圍--在找到前一個值得(即注音、拼音陣列的下一個索引之元素)
'                    charItem = InStr(selrangeP.Range, ZhuYin) '取得此段落第一個注音符號的位置--非也,乃比對有注音段落時,第一個比對到的注音符號(與陣列中次序有關也)!
                    For char = charItem To selrangeP.Range.Characters.Count '要從找到注音的段落從第一字,逐字檢查,凡係注音符號,即寫入陣列(此因注音在拼音前故可如是寫也)
                    '所以此順序與注音符號陣列之順序極相關也.輕聲符號必置於最前才能不漏掉,否則要再加一行,繼續搜尋此段裡是否還有注音陣列裡的元素,且還在注音陣列中找到的,在原段落更前的位置.此才能抓到真正的第一注音符號位置!
                    '然據下找到注音符號位置,並無關注音陣列內順序也.
                        'If Not IsEmpty(Filter(zhu, selrangeP.Range.Characters(Char), False, vbBinaryCompare)) Then  '找到注音符號
                        '原用上行,但因找不到依然傳回陣列(與線上說明所言有異!!),唯其找到一個則陣列索引值會減一,故改寫成下式:
                        If UBound(Filter(zhu, selrangeP.Range.Characters(char), False, vbBinaryCompare)) < Uzhu Then '找到注音符號
                            If selrangeP.Range.Characters(char) = "．" And selrangeP.Range.Characters(char).Font.Name <> "標楷體" Then GoTo nextP '"．"此輕聲符號與數字編號符號相同,故只能以字形,或其後接字元是否為注音符號判斷了.今暫以字形耳.
                            charItem = char 'InStr(selrangeP, selrangeP.Range.Characters(Char)) '取得第一個注音符號的位置
                            If selrangeP.Range.Characters(char) <> ChrW(20008) Or (selrangeP.Range.Characters(char) <> ChrW(20008) And InStr(selrangeP.Range, "】") = 0) Then
11                              Do Until selrangeP.Range.Characters(charItem).Font.Name = "SimSun" Or selrangeP.Range.Characters(charItem).Font.Name = "Pintone Times" Or selrangeP.Range.Characters(charItem) = Chr(13)
                                    If UBound(Filter(zhu, selrangeP.Range.Characters(char), False, vbBinaryCompare)) < Uzhu Then
                                        zhuyinText = zhuyinText & selrangeP.Range.Characters(charItem)
                                    End If
                                    charItem = charItem + 1
                                Loop
                                zhuyinText = Trim(zhuyinText) '取得注音
                                Do Until (selrangeP.Range.Characters(charItem).Font.Name <> "SimSun" And selrangeP.Range.Characters(charItem).Font.Name <> "Pintone Times") Or selrangeP.Range.Characters(charItem) = Chr(13)
                                    'If UBound(Filter(zhu, selrangeP.Range.Characters(Char), False, vbBinaryCompare)) < Uzhu Then'接在注音後必為拼音,故不可如是判斷
                                        pinyinText = pinyinText & selrangeP.Range.Characters(charItem)
                                    'End If
                                    charItem = charItem + 1
                                Loop
                                pinyinText = Trim(pinyinText) '取得拼音
                                'GoTo nextP '取完
                                ReDim Preserve Zhuyins(zps)
                                ReDim Preserve Pinyins(zps)
'                                ReDim Zhuyins(zps)'還是要Preserve才能保陣列前一索引值
'                                ReDim Pinyins(zps)
                                Zhuyins(zps) = zhuyinText
                                Pinyins(zps) = pinyinText
                                zps = zps + 1
                                zhuyinText = ""
                                pinyinText = ""
                                char = charItem
                            Else
'                                If InStr(selrangeP.Range, Char) > InStr(selrangeP.Range, "】") Then
                                If char > InStr(selrangeP.Range, "】") Then
                                    GoTo 11
                                End If
                            End If
'                        Else
'                            Stop
                        End If
                    Next char
                End If
            Next
nextP: Next
        'If AscW(Terms) = 25908 Then Stop
'        If Terms = "斁" Then Stop
'        If Terms = "斖" Then Stop
'        If Terms = "曳" Then Stop
        With rst
            .Index = "字"
            .Seek "=", Terms
            If .NoMatch Then
                .AddNew
            Else
                .edit
            End If
            .Fields("字") = Terms
            .Fields("釋文") = wcontent
            If IsNumeric(left(wcontent, 5)) Then
                Set rst1 = db.OpenRecordset("字")
                With rst1
                    .Index = "字序"
                    .Seek "=", left(wcontent, 5)
                    If .NoMatch = False Then
                        If .Fields("字ID") <> 字ID Then
                            If 不輸入字序 <> vbNo Then
                                不輸入字序 = MsgBox("字序" & left(wcontent, 5) & "重複了!要繼續,則不會匯入字序!" & vbCr & vbCr _
                                    & "  要終止這次匯入,以改正原檔案,請按[取消]!!" & _
                                    "要繼續,請按[是].   不再檢覈字序,直接略過(不輸入字序)則按[否]   " & _
                                    "按下〔取消〕後,此字序有誤之字,其條目內容將會反白顯示!", vbExclamation + vbYesNoCancel, "字序重複了!!!")
                                Select Case 不輸入字序
                                    Case vbCancel
                                        'ActiveDocument.Windows(1).Visible = True
                                        Documents.Open ActiveDocument.FullName, , False
                                        End
                                End Select
                            End If
                        Else
                            rst.Fields("字序") = left(wcontent, 5)
                        End If
                    Else
                        rst.Fields("字序") = left(wcontent, 5)
                    End If
                End With
            End If
            If before部首 <> "" Then '取得字的部首ID
                Set rst1 = db.OpenRecordset("部首")
                With rst1
                    .Index = "部首"
                    .Seek "=", before部首
                    If .NoMatch Then
                        MsgBox "請程式撰寫者檢查!", vbExclamation: Stop '因為214部首已入,若有找不到,可疑!
                    Else
                        If Not IsNull(rst.Fields("部首ID")) Then
                            If .Fields("部首ID") <> rst.Fields("部首ID") Then
                                Dim rst1clone As Recordset
                                Set rst1clone = rst1.Clone
                                With rst1clone
                                    .Index = "部首ID"
                                    .Seek "=", rst.Fields("部首ID")
                                End With
                                If MsgBox("『" & rst.Fields("字") & "』字部首已為『" _
                                    & rst1clone.Fields("部首") & "』部,是否更改為『" & .Fields("部首") & _
                                    "』部？", vbExclamation + vbYesNo, "部首變更確認!") = vbYes Then
                                    rst.Fields("部首ID") = .Fields("部首ID")
                                    rst1clone.Close
                                    Set rst1clone = Nothing
                                End If
                            End If
                        Else
                            rst.Fields("部首ID") = .Fields("部首ID") ' DLookup("部首ID", "部首", "部首 = """ & before部首 & """")
                        End If
                    End If
                    .Close
                End With
            End If
            If zps = 0 Then
                If .Fields("不輸入注音") = False Then
                    If MsgBox(Terms & "字缺注音!是否改定原稿後再重新執行一次匯入？" _
                        & vbCr & vbCr & "確定終止這次匯入，請按〔確定〕鈕！" & vbCr & vbCr & _
                        "按下〔確定〕後,此無注音之字,其條目內容將會反白顯示!", vbCritical + vbOKCancel) = vbOK Then
                        'ActiveDocument.Windows(1).Visible = True
                        Documents.Open ActiveDocument.FullName, , False
                        End
                    Else
                        不輸入注音 = True
                    End If
                Else
                    不輸入注音 = True
                End If
            End If
            If zps > 0 Then zps = zps - 1
            .Fields("編輯時間") = Null '清除編輯標記,以免再次詢問覆蓋確認2010/10/27
            If 不輸入注音 Then .Fields("不輸入注音") = True '記下不輸入字音之字
            .Update
        End With
        If 不輸入注音 = False Then
            Set rst = db.OpenRecordset("字_注音") '輸入字_注音
            With rst
    '            If zps > 0 Then Stop
                For char = 0 To zps  'UBound(Zhuyins(zps))
                    Set 注音拼音rst = db.OpenRecordset("注音") '檢查注音有無
                    With 注音拼音rst
                        If InStr(1, "．˙‧" & ChrW(12539), left(Zhuyins(char), 1), vbBinaryCompare) Then 'Debug.Print Terms: Stop
                            If MsgBox("注音「" & Zhuyins(char) & "」是否輸入？" & vbCr & vbCr & "因為輕聲符號有與此黑點符號相同者，故程式可能會抓錯。請判斷！", vbExclamation + vbOKCancel + vbDefaultButton2, "輕聲注音輸入確認！！") = vbCancel Then GoTo exitZhuyin
                        End If
                        Zhuyins(char) = Replace(Replace(Replace(Zhuyins(char), "˙", "‧"), "．", "‧"), ChrW(12539), "‧")
21                      .Index = "注音"
                        .Seek "=", Zhuyins(char)
                        If .NoMatch Then
                            .AddNew
                            .Fields("注音") = Zhuyins(char)
                            .Update
                            '.Close '果然不能關掉,關掉則無效了,得重新set
                            .OpenRecordset
                            GoTo 21 '新增後再去找注音ID值
                        Else
                            注音ID = .Fields("注音ID")
                        End If
                    End With
                    Set 注音拼音rst = db.OpenRecordset("拼音") '檢查拼音有無
                    With 注音拼音rst
22                      .Index = "拼音"
                        .Seek "=", Pinyins(char)
                        If .NoMatch Then
                            .AddNew
                            .Fields("拼音") = Pinyins(char)
                            .Update
                            .OpenRecordset
                            GoTo 22 '新增後再去找注音ID值
                        Else
                            拼音ID = .Fields("拼音ID")
                        End If
                    End With
                    
                    .Index = "PrimaryKey"
                    .Seek "=", 字ID, 注音ID
                    If .NoMatch Then
                        .AddNew
                    Else
                        .edit
                    End If
                    .Fields("字ID") = 字ID
                    .Fields("注音ID") = 注音ID
                    .Fields("拼音ID") = 拼音ID
                    .Update
exitZhuyin:    Next char
            End With
        End If
        'beforeTerm = Terms
    Case "XXX詞"
        Set rst = db.OpenRecordset("詞")
        With rst
            .Index = "詞"
            .Seek "=", Terms
            If .NoMatch Then
                .AddNew
            Else
                '檢查編輯更新情況
                If Not IsNull(.Fields("編輯時間")) Then
                    If MsgBox("『" & .Fields("詞") & "』一詞在" & .Fields("編輯時間") & "曾經編輯更新過" & _
                        "是否匯入?--匯入將會覆蓋<詞典>資料內原『" & .Fields("詞") & "』詞條已經編輯過之資料" _
                        & vbCr & vbCr & "若略過,可記下此字,在此次匯入完成後," & _
                        "再回<詞典>資料庫檢索檢視" & vbCr & vbCr & "要匯入覆蓋原資料,請按[確認]!! 要略過" _
                        & "請按[取消]!", vbOKCancel + vbExclamation, "覆蓋原資料確認!!!") = vbCancel Then _
                            GoTo 不匯入
                End If
                .edit
            End If
            .Fields("字ID") = 字ID
            .Fields("詞") = Terms
            .Fields("釋文") = wcontent
            .Fields("編輯時間") = Null
            .Update
        End With

    Case "XXX衍"
        Set rst = db.OpenRecordset("衍")
        With rst
            .Index = "字ID" '原作"字ID1"須改!
            .Seek "=", 字ID
            If .NoMatch Then
                .AddNew
            Else
                .edit
            End If
            .Fields("字ID") = 字ID
            .Fields("衍") = wcontent
            .Update
        End With
    Case "XXX部"
        Set rst = db.OpenRecordset("部首")
        With rst
            .Index = "部首"
            .Seek "=", Terms
            If .NoMatch Then
                .AddNew
                .Fields("部首") = Terms
                .Update
'            Else
'                .Edit
            End If
        End With
    Case Else
        MsgBox "有bug,請通知程式撰寫人!", vbCritical
End Select
不匯入:
Exit Sub
ErrH:
    MsgBox Err.Number & Err.Description: Stop: Resume
End Sub


Function 注音符號()
注音符號 = Array("ㄅ", "ㄆ", "ㄇ", "ㄈ", "ㄉ", "ㄊ", "ㄋ", "ㄌ", "ㄍ", "ㄎ", "ㄏ", "ㄐ", "ㄑ", "ㄒ", "ㄓ", "ㄔ", "ㄕ", "ㄖ", "ㄗ", "ㄘ", "ㄙ", "ㄚ", "ㄛ", "ㄜ", "ㄝ", "ㄞ", "ㄟ", "ㄠ", "ㄡ", "ㄢ", "ㄣ", "ㄤ", "ㄥ", "ㄦ", "ㄧ", "ㄨ", "ㄩ", "‧", "ˊ", "ˇ", "ˋ", ChrW(20008), "˙", "．", ChrW(12539)) '輕聲符號即有四種!!'以此「‧」為準!!
'所以此順序與注音符號陣列之順序極相關也.輕聲符號必置於最前才能不漏掉,否則要再加一行,繼續搜尋此段裡是否還有注音陣列裡的元素,且還在注音陣列中找到的,在原段落更前的位置.此才能抓到真正的第一注音符號位置!
'悠關第一個找到的位置,則注音符號的順序即極重要!!
'但是如腰等音抓來確正常呀???
End Function



Sub 詞典條目內容擷取(w As Document, acpath As String)
Dim p As Paragraph, pselect ' As Range
Dim a As Object 'Access.Application
Dim beforeTerm As String, Terms As String, wcontent As String, before部首 As String
Dim i As Long, e As Boolean
Dim docEnd As Long
On Error GoTo ErrH
'Set A = GetObject("C:\Documents and Settings\Superwings\桌面\詞典.MDB")
Set a = GetObject(acpath)
With w.ActiveWindow.Selection
    docEnd = .Document.Range.End
    'If MsgBox("是否從頭(文件首)開始?", vbYesNo + vbQuestion) = vbYes Then .HomeKey wdStory, wdMove
    '改寫到詞典.mdb中
    Do
        If .Paragraphs(1).Range.Font.Hidden = False And .Paragraphs(1).Range.Font.Hidden <> 9999999 Then
                            '9999999乃包括隱藏與不隱藏也.
'            If InStr(.Paragraphs(1).Range, "◆") Then Stop
'            If InStr(.Paragraphs(1).Range, "橖") Then Stop
            If left(.Paragraphs(1).Range, 1) = "◆" Then ''字之部首處理
'                If InStr(.Paragraphs(1).Range, "木部") Then Stop
'                If InStr(.Paragraphs(1).Range, "日部") Then Stop
                GoSub 選取
                wcontent = .Range
                wcontent = wcontent & "XXX部"
                Terms = Mid(wcontent, 2, 1)
                If beforeTerm = "" Then beforeTerm = Terms
                before部首 = Terms
            ElseIf InStr(Selection.Paragraphs(1).Range, "【") Then
                GoSub 選取
    '            .Paragraphs(1).Range.Select
    '            Do Until InStr(.Next.Paragraphs(1).Range, "【") Or Left(.Next.Paragraphs(1).Range, 1) = "◇"
    '                .MoveDown wdParagraph, 1, wdExtend
    '            Loop
    '            Terms = Mid(wcontent, InStr(wcontent, "【") + 1, InStr(wcontent, "】") - (InStr(wcontent, "【") + 1))
                wcontent = Replace(.Range, Chr(13), Chr(13) & Chr(10))
                wcontent = Replace(wcontent, "〕", "〕" & Chr(13) & Chr(10))
                wcontent = Replace(wcontent, "〔", Chr(13) & Chr(10) & "〔")
                Terms = Mid(wcontent, InStr(wcontent, "【") + 1, InStr(wcontent, "】") - (InStr(wcontent, "【") + 1))
                Select Case InStr(Selection, "】") - InStr(Selection, "【")
                    Case Is = 2 '字
                        wcontent = wcontent & "XXX字"
                        'A.Run "Word資料匯入_字', 詞典條目內容擷取, Terms"
    '                    Word資料匯入Access wcontent, Terms, beforeTerm
                        beforeTerm = Terms
                    Case Is > 2
                        wcontent = wcontent & "XXX詞"
                        'A.Run "Word資料匯入_詞", 詞典條目內容擷取, Terms, beforeTerm
    '                    Word資料匯入Access wcontent, Terms, beforeTerm
                End Select
            ElseIf InStr(.Paragraphs(1).Range, "◇") Then ''字之衍生詞處理
                GoSub 選取
                wcontent = .Range
                wcontent = wcontent & "XXX衍"
                'A.Run "Word資料匯入_衍", 詞典條目內容擷取, beforeTerm
    '            Word資料匯入Access wcontent, Terms, beforeTerm
            Else
                GoTo NextOne '不處理
            End If
            If wcontent <> "" Then Word資料匯入Access wcontent, Terms, beforeTerm, acpath, before部首, .Range
        End If
        'If .Range.End <> .Document.Range.End Then .MoveDown wdParagraph, 1, wdMove
NextOne: If e = False Then .MoveDown wdParagraph, 1, wdMove
        i = i + 1
        w.Application.StatusBar = "正在處理第" & i & "筆"
        a.Forms("主表單").Controls("Word資料匯入").Caption = "正在處理第" & i & "筆"
        a.Forms("主表單").Repaint
        
    Loop Until e '.Range.End = .Document.Range.End
'    MsgBox "匯入完成!", vbInformation' 2010/10/20'今改寫到Access裡面!
'    a.Forms("主表單").Controls("Word資料匯入").Caption = "Word資料匯入"
    If w.Application.Documents.Count = 0 Then
        w.word.Application.Quit wdDoNotSaveChanges
    Else
        w.Close wdDoNotSaveChanges
    End If
    AppActivate Replace(a.CurrentProject.Name, ".mdb", "")
    With a.Forms("主表單")
        .Controls("Word資料匯入").FontSize = 8
        .Controls("Word資料匯入").Caption = "第" & i & "筆處理完畢,正在匯入詞首、詞尾與成語,請稍候..."
        .Repaint
    End With
    db.Close
    Set db = Nothing
Exit Sub
選取:
    .Paragraphs(1).Range.Select
    If .Range.End = docEnd Then e = True: Return
    Do Until InStr(.Next.Paragraphs(1).Range, "【") Or _
            left(.Next.Paragraphs(1).Range, 1) = "◇" Or _
            left(.Next.Paragraphs(1).Range, 1) = "◆" Or _
            .Next.Paragraphs(1).Range.Font.Hidden = (True Or 9999999)
        .MoveDown wdParagraph, 1, wdExtend
        If .Range.End = docEnd Then e = True: Exit Do
    Loop
    Return
End With
'With ActiveDocument
'    For Each p In .Paragraphs
'        If InStr(p, "【") Then
'            If Left(p, 1) <> "◇" Then
'                'p.Range.SetRange
'                p.Range.Select
'                'Selection.MoveEndUntil Chr(13) & Chr(13), wdForward    '(Chr(13) & Chr(13)) ' "【"
'    '            .Range.MoveEndUntil Chr(13) & Chr(13), wdForward
'                Selection.MoveEndUntil "【", wdForward
'    '            p.Range.MoveEndUntil "【", wdForward
'                If InStr(Selection, "◇") Then
'                    'Selection.MoveStartUntil "◇", wdForward
'                    'Selection.MoveLeft wdParagraph, 2, wdExtend
'                    Selection.MoveUp wdParagraph, 3, wdExtend '  "◇"
'                Else
'                    If Selection.Next.Font.Size > 12 Then '=20
'                        Selection.MoveLeft wdCharacter, 6, wdExtend
'                    End If
'                End If
'    '            p.Range.Copy
''                .Windows(1).ScrollIntoView Selection, False
'                詞典條目內容擷取 = Selection.Text
'            Else
'                '字之衍生詞處理
'                Selection.Paragraphs (1)
'            End If
'        ElseIf Left(p, 1) <> "◇" Then
'
'        End If
'    Next
'End With
ErrH:
Select Case Err.Number
    Case Else
        If MsgBox("出錯了,請通知系統撰寫員!" & vbCr & _
        "錯誤敘述:" & Err.Number & Err.Description & vbCr & vbCr & _
        "按確定→結束.按取消→偵錯.", vbExclamation + vbOKCancel) = vbCancel Then Stop: Resume
        
End Select
End Sub










Sub 漢語拼音置換() 'PinTone轉SimSun
Dim db As Database, p As Paragraph, chra As word.Characters, m As Long, pinL As Long
Dim rst As Recordset, rstReplace As Recordset, Paragraph As String, ParagraphNew As String, pL As Integer, ps As Integer
Dim zhu, Uzhu, i As Byte, zhuL As Integer ', zhuYes As Boolean
Dim tble, t As Byte, b As Byte
Dim ds As Date, de As Date, times As Long
Static wpath As String
wpath = "D:\千慮一得齋\書信\圖書管理\黃沛榮老師助理\詞典.mdb"
If Dir(wpath) = "" Then wpath = "C:\Documents and Settings\Superwings\桌面\詞典.mdb"
'wpath = InputBox("請輸入欲匯入檔案的全路徑檔名(含副檔名)", , _
    IIf(Nz(wpath, "") = "", "C:\Documents and Settings\Superwings\桌面\詞典.mdb", wpath))
wpath = InputBox("請輸入欲匯入檔案的全路徑檔名(含副檔名)", , wpath)
If wpath = "" Then Exit Sub

Set db = DBEngine.OpenDatabase(wpath)
ds = Now
'zhu = Array("ㄅ", "ㄆ", "ㄇ", "ㄈ", "ㄉ", "ㄊ", "ㄋ", "ㄌ", "ㄍ", "ㄎ", "ㄏ", "ㄐ", "ㄑ", "ㄒ", "ㄓ", "ㄔ", "ㄕ", "ㄖ", "ㄗ", "ㄘ", "ㄙ", "ㄚ", "ㄛ", "ㄜ", "ㄝ", "ㄞ", "ㄟ", "ㄠ", "ㄡ", "ㄢ", "ㄣ", "ㄤ", "ㄥ", "ㄦ", "ㄧ", "ㄨ", "ㄩ", "˙", "ˊ", "ˇ", "ˋ", ChrW(20008))
'Uzhu = UBound(zhu)
Set rstReplace = db.OpenRecordset("Pintone取代為拼音用查詢")
With rstReplace
    Do Until .EOF
    '                If .Fields("拼音Pintone") = "y9" Then Stop
        Paragraph = .Fields("拼音Pintone")
        ParagraphNew = .Fields("拼音")
        With Selection.Range.Find
            .text = Paragraph
            .ClearFormatting
'            .Font.Name = "Pintone Times"
            .Font.NameAscii = "Pintone Times"
'            .Font.NameFarEast = "Pintone Times"
            .Font.NameOther = "Pintone Times"
            If .Execute(, , , , , , True, wdFindContinue) Then
                .Replacement.Font.Name = "SimSun" '"Times New Roman"
                .Replacement.Font.NameAscii = "SimSun"
                .Replacement.Font.NameFarEast = "SimSun"
                .Replacement.Font.NameOther = "SimSun"
                .Replacement.text = ParagraphNew
                .Execute , True, False, , , , True, wdFindContinue, , , wdReplaceAll
                times = times + 1
            End If
'            .Parent.Collapse
            
            'Selection.Collapse
'            Selection.HomeKey
        End With
        
        .MoveNext
    Loop
'    .MoveFirst
End With
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting

'MsgBox "完成!", vbInformation
'Exit Sub
'
''
'For Each p In ActiveDocument.Paragraphs
'    For Each chra In p.Range.Characters
'        m = m + 1
'        If chra.Range.Font.Name = "Pintone Times" Then
'            Do Until chra.Range.Font.Name <> "Pintone Times" Or chra.Range Like Chr(13)
'                pinL = m
'            Loop
'
'        End If
'
'    Next chra
'        Paragraph = p.Range 'Nz(.Fields("釋文"))
        'If InStr(firstParagraph, Chr(13)) <> 0 Then
        'pL = InStr(Paragraph, Chr(13))
''        Do Until pL = 0
'''            Do
''                Paragraph = Mid(Paragraph, pS + 1, pL)
'''            Loop Until Paragraph <> ""
'            For i = 0 To Uzhu
'                zhuL = InStr(Paragraph, zhu(i))
'                If zhuL Then Exit For
'            Next i
'            If zhuL Then
'                If InStr(Paragraph, "【") Then zhuL = InStr(Paragraph, "【")
'                Paragraph = Mid(Paragraph, zhuL)
'                ParagraphNew = Paragraph
'
'                With rstReplace
'                    Do Until .EOF
'        '                If .Fields("拼音Pintone") = "y9" Then Stop
'                        ParagraphNew = Replace(ParagraphNew, .Fields("拼音Pintone"), .Fields("拼音"), , , vbBinaryCompare)
'                        .MoveNext
'                    Loop
'                    .MoveFirst
'                End With
''                .Edit
''                .Fields("釋文") = Replace(.Fields("釋文"), Paragraph, ParagraphNew)
''                .Update
'                p.Range.Text = Replace(p.Range.Text, Paragraph, ParagraphNew)
'                zhuL = 0 ''歸零
'            End If
'            'pS = pL + 2
'            'pL = InStr(pS, .Fields("釋文"), Chr(13))
'            'Paragraph = Nz(.Fields("釋文"))
''        Loop
'        'pS = 0 ''歸零
'        'End If
''        .MoveNext
'    'Loop
'    Next p
''End With
'Next t
de = Now
MsgBox "完成!" & vbCr & "費時" & Minute(de - ds) & "分" & Second(de - ds) & "秒" & vbCr _
    & "共" & times & "次置換", vbInformation

End Sub

Private Sub Document_Close()

''閉閉文件時以書籤方式記住游標位置
'With ThisDocument
'If Not .path = "" Then '舊檔才檢查2003/4/1
'    If .Saved = False Then
'        If .Name = "Normal.dotm" Then
'            If MsgBox("要儲存""Normal.dotm""範本嗎？", vbQuestion + vbOKCancel) = vbOK Then .Save
'            Exit Sub
'        End If
'        If DonotSave = False Then
'            Select Case MsgBox("要記下游標最後位置嗎?", vbYesNoCancel + vbInformation)
'                Case Is = vbYes
'                    游標所在位置書籤
'            '        ActiveDocument.Save
'                Case vbNo
'                    If MsgBox("文件已修改,是否要儲存", vbExclamation + vbYesNo) = vbYes Then
'                        .Save
'
'                    Else
'                        .Saved = True '如此設定Word將不再詢問是否要儲存的對話框!2003/3/22
'                    End If
'                Case vbCancel
'                    Stop
'                    .Undo
'                    .Save
'            End Select
'        End If
'    End If
'End If
'End With
''If Tasks.Exists("詞典") Then
''    If Tasks("詞典").Visible = False Then Tasks("詞典").Close
''End If
End Sub

Sub combars()
Dim a As CommandBar
For Each a In CommandBars
    Debug.Print a.Name
Next a
End Sub

Private Sub Document_New()
On Error Resume Next
Register_Event_Handler '原來會失效是因為我在"在另一文件中尋找選取字串"用了END啦!
'If ActiveDocument.AttachedTemplate.Name = "Normal.dot" Then
'    CommandBars.Item("blog").Visible = True
'Else
'    CommandBars.Item("blog").Visible = False
'End If
'Docs.DocBackgroundFillColor'原來頁面會設置暗墨綠色是由於這行啊。哈哈，現在才發現，一直傻傻地在設計那裡找與設定也無效，剛才靈機一動，該不會是用程式設定的吧，一查，果然！感恩感恩　讚歎讚歎　南無阿彌陀佛
'上面也都是舊版的了，除了事件處理程式掛鉤要留外，其餘均註解掉了 20230111
End Sub


Private Sub Document_Open()
'Public Sub Document_Open() '其他範本參照用!故不用Private'2004/10/17
'On Error GoTo eH

Dim bk As Bookmark, bkIdx As Integer, tp As Template '2015/11/4
Register_Event_Handler
With ThisDocument
    If left(.Name, 2) = "排印" Then .ActiveWindow.DisplayVerticalRuler = False
    With .Application.CommandBars
        If Application.Templates.Count > 1 Then
            For Each tp In Application.Templates
                If tp.Name = "mywork.dot" Then If .item("得士派").Visible Then .item("得士派").Visible = False
            Next
        End If
        If left(ThisDocument, 7) Like "Access_" _
            Or InStr(ThisDocument.FullName, "文字檔資料庫") Then _
            .item("論文＿札記瀏覽").Visible = True
        If .item("Web").Visible = True Then .item("web").Visible = False
    End With
    If .bookmarks.Count > 0 Then '有書籤時才做
        '        For bkIdx = 1 To .Bookmarks.Count
                For Each bk In .bookmarks '將兩天前的標地書籤刪除'2003/3/28
                    bkIdx = bkIdx + 1 '記下書籤索引
                    With bk '如果是編輯處才處理
                        If InStr(1, .Name, "Edit_", vbTextCompare) > 0 _
                            And InStr(1, bk, "編輯", vbTextCompare) = 0 Then
                            '如果是兩天前的
                            Do While InStr(1, bk, "Edit_", vbTextCompare) > 0 And _
                                    CDate(Replace(Mid(bk, 6, 10), "_", "/")) <= Date - 2 _
                                    And InStr(1, bk, "編輯", vbTextCompare) = 0
                                    If ThisDocument.Saved Then
                                        If MsgBox("要刪除前天以前編輯之位置書籤嗎?", vbYesNo + vbExclamation) = vbYes Then
                                            bk.Delete '殺掉以後索引值會向前遞補'留下待編輯處2003/4/2
                                        Else
                                            Set bk = Nothing
                                            Exit For
                                        End If
                                    Else
                                        bk.Delete
                                    End If
                                    If ThisDocument.bookmarks.Count = 0 Then Exit Sub
                                    Set bk = ThisDocument.bookmarks(bkIdx)
                                    If InStr(1, bk, "Edit_", vbTextCompare) = 0 Then _
                                            Exit For
                            Loop
                        End If
                        If InStr(1, bk, "Edit_", vbTextCompare) = 0 Then _
                            Exit For
                    End With
                Next bk
        '        Next bkIdx
    End If
    With .Application.CommandBars
        If .item("web").Visible = True Then .item("web").Visible = False
    End With
'    .ActiveWindow.WindowState = wdWindowStateMaximize
'    .ActiveWindow.Application.GoBack '到最後編輯處,作用等同於按 Shift+F5 鍵
    StatusBar = .AttachedTemplate.Name & " template is attached to the active document"
    Options.CheckGrammarAsYouType = False

End With
111
'If MsgBox("是否要進行造字字元檢查?", vbQuestion + vbOKCancel + vbDefaultButton2) = vbOK Then 造字字元檢查
'在狀態列顯示本文件所使用的範本名稱
'Exit Sub
'eH:
'Select Case Err.Number
'    Case 4248 '因為沒有開啟文件，所以無法使用這個指令。
'        GoTo 111 'Resume Next'91沒有設定物件變數或 With 區塊變數
'    Case Else
'        MsgBox Err.Number & vbTab & Err.Description
'End Select
DoEvents
Me.在整頁模式中不顯示頁面間的空白區域
If UserProfilePath = "" Then UserProfilePath = SystemSetup.取得使用者路徑_含反斜線()
End Sub

Private Sub Application_Quit()
'20230119 YouChat菩薩:'根本就不會觸動啟動此程序！
  ' Perform the event handling procedure here.
  SystemSetup.killchromedriverFromHere
End Sub
