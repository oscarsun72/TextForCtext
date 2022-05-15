Attribute VB_Name = "文字轉換"
Option Explicit
Public UserForm1TextBox1Value As String
Sub 漢字轉拼音()
'F2

On Error Resume Next
Dim pt As String
pt = Replace(取得桌面路徑, "Desktop", "Dropbox")
'Const fpath As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\ssz3\Google 雲端硬碟\私人\VB\詞典.mdb" ';Mode=Share Deny None;Extended Properties="";Jet OLEDB:System database=C:\Users\ssz3\AppData\Roaming\Microsoft\Access\System.mdw;Jet OLEDB:Registry Path=Software\Microsoft\Office\16.0\Access\Access Connectivity Engine;Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=True;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False"
Dim fpath As String
fpath = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & pt & "\VS\詞典.mdb" ';Mode=Share Deny None;Extended Properties="";Jet OLEDB:System database=C:\Users\ssz3\AppData\Roaming\Microsoft\Access\System.mdw;Jet OLEDB:Registry Path=Software\Microsoft\Office\16.0\Access\Access Connectivity Engine;Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=True;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False"
Dim rst As New ADODB.Recordset, rstPinyin As New ADODB.Recordset
Dim cnt As New ADODB.Connection
'Dim p As New ADODB.Parameter
Dim x As String, y As String, z As String
Dim cmd As New ADODB.Command
If Selection.Type = wdSelectionNormal Then
    x = Selection
    Selection.Copy
Else
    x = Selection.Previous
End If
cnt.Open fpath
Set cmd.ActiveConnection = cnt
cmd.CommandText = "SELECT 字.字, 拼音.拼音 FROM 拼音 INNER JOIN (字 INNER JOIN 字_注音 ON 字.字ID = 字_注音.字ID) ON 拼音.拼音ID = 字_注音.拼音ID" _
        & " WHERE (((字.字)=""" & x & """) ) order by 字_注音.字_注音ID;"
cmd.CommandType = adCmdText
rstPinyin.Open cmd
If rstPinyin.EOF Then
    Beep
    MsgBox "大菩薩好：尚無「" & x & "」此字之拼音、注音，請手動補充。感恩感恩 南無阿彌陀佛", vbExclamation
Else
    Do Until rstPinyin.EOF
        'y = "（" & rst.Fields("拼音").Value & "）"
        y = rstPinyin.Fields("拼音").Value
        cmd.CommandText = "SELECT 字.字 FROM 拼音 INNER JOIN (注音 INNER JOIN (字 INNER JOIN 字_注音 ON 字.字ID = 字_注音.字ID) ON 注音.注音ID = 字_注音.注音ID) ON 拼音.拼音ID = 字_注音.拼音ID " & _
                        "WHERE (((拼音.拼音) = """ & y & """) And ((字.罕用字) = False) And ((字_注音.先後)= 0)) ORDER BY 字.字, 字.字序;"
    '    rst.Close
        rst.Open cmd
        z = Replace(rst.GetString(, , , " "), x & " ", "")
        With Selection
            .Collapse wdCollapseEnd
            .TypeText y
            .MoveLeft wdCharacter, Len(y), wdExtend
            .Font.Name = "simsun" '"NSimSun" simhei'"Verdana"
            .Font.ColorIndex = wdRed
            .Font.Size = 14 '18
            .Collapse wdCollapseEnd
            .TypeText z
            .MoveLeft wdCharacter, Len(z), wdExtend
            .Font.Name = "標楷體"
            .Font.ColorIndex = wdRed
            .Font.Size = 14
            .MoveLeft wdCharacter, Len(y), wdExtend
            .Range.HighlightColorIndex = wdYellow
            .Font.Bold = False
            .Collapse wdCollapseEnd
        End With
    '    Selection.Document.Save
        rst.Close
        rstPinyin.MoveNext
    Loop
End If
rstPinyin.Close
cnt.Close
End Sub
Sub 異體字轉正() '20210228
Dim cnt As New ADODB.Connection, rst As New ADODB.Recordset, rng As Range, a, docAs As String, x As String, rngCC As Integer, flaSel As Byte, flgsel, drv As String, db As New dBase, ur As UndoRecord
'Alt+7
If Selection.Type = wdSelectionIP Then
    Set rng = ActiveDocument.Range
    flgsel = wdFindContinue
Else
    Set rng = Selection.Range
    flgsel = wdFindStop
End If
Set ur = SystemSetup.stopUndo("異體字轉正")
'For Each a In rng.Characters
'    If InStr(docAs, a) = 0 Then docAs = docAs & a '記下不重複的用字集
'Next
rngCC = rng.Characters.Count
'If VBA.Dir("H:\", vbVolume) = "" Then
'    drv = "D:\千慮一得齋\"
'Else
'    drv = "H:\我的雲端硬碟\私人\千慮一得齋(C槽版)\"
'End If
'cnt.Open _
    "Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=" & drv & "書籍資料\圖書管理附件\查字.mdb;"
db.cnt查字 cnt

For Each a In rng.Characters
    If InStr(docAs, a) = 0 Then
        docAs = docAs & a '記下不重複的用字集
        If rngCC = 1 Then
            rst.Open "select * from 異體字轉正 " & _
                "where strcomp(異體字,""" & a & """)=0 " & _
                "order by 排序", cnt, adOpenKeyset, adLockReadOnly
        Else
            rst.Open "select * from 異體字轉正 " & _
                "where (strcomp(異體字,""" & a & """)=0 " & _
                "and 手挍=false) order by 排序", cnt, adOpenKeyset, adLockReadOnly
        End If
        If rst.RecordCount > 0 Then
            x = 異體字轉正_取得取代資料(rst)
            With rng.Find
                .Text = a 'rst.Fields("異體字")
                .Replacement.Text = x 'rst.Fields("正字")
                .Execute , , , , , , True, flaSel, , , wdReplaceAll
            End With
        End If
        rst.Close
    End If
Next
'保留他是異體字形也好，可作識別
'For Each a In rng.Characters
'    Select Case a.Font.NameFarEast
'        Case "微軟正黑體", "hanaminb", "kaixinsongb"
'            rng.Font.NameFarEast = "新細明體-extb"
'    End Select
'Next a
endS:
If Not rst.State = adStateClosed Then rst.Close
cnt.Close: Set rst = Nothing: Set cnt = Nothing: Set a = Nothing: Set rng = Nothing
SystemSetup.playSound 2
SystemSetup.contiUndo ur
End Sub
Sub 異體字轉正_新增資料() '20210228
Dim cnt As New ADODB.Connection, rst As New ADODB.Recordset, rng As Range, rngx As String, bz As Boolean, flaSel As Byte, drv As String, db As New dBase
Static x As String
If Selection.Type = wdSelectionIP Then
    Set rng = Selection.Characters(1)
    flaSel = wdFindContinue
Else
    Set rng = Selection.Range
    flaSel = wdFindStop
End If
If VBA.InStr(Chr(13) & Chr(7) & Chr(8) & Chr(9) _
        , rng) Then Exit Sub
If Len(rng) > 2 Then MsgBox "非單字，請檢查！", vbExclamation: Exit Sub
If rng = "" Then Exit Sub

x = InputBox("請輸入正字", , x)
If InStr(x, "?") Then
    UserForm1.Show
    x = UserForm1TextBox1Value
End If
If x = "" Then Exit Sub
If x Like "*[a-z0-9A-Z]*" Then MsgBox "非中文，請檢查！", vbExclamation:        Exit Sub
If StrComp(x, rng) = 0 Then MsgBox "與要轉換之字相同，請檢查！", vbExclamation: Exit Sub
x = VBA.Trim(x)
rngx = VBA.Trim(rng.Text)
If MsgBox("是否兼正字？", vbQuestion + vbOKCancel + vbDefaultButton2) = vbOK Then bz = True
'If VBA.Dir("H:\", vbVolume) = "" Then
'    drv = "D:\千慮一得齋\"
'Else
'    drv = "H:\我的雲端硬碟\私人\千慮一得齋(C槽版)\"
'End If
'cnt.Open _
'    "Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=" & drv & "書籍資料\圖書管理附件\查字.mdb;"
db.cnt查字 cnt
rst.Open "select * from 異體字轉正" & _
    " where strcomp(異體字,""" & rngx & """)=0 and " & _
        "strcomp(正字,""" & x & """)=0 " & _
    "order by 排序", cnt, adOpenKeyset, adLockOptimistic
If rst.RecordCount = 0 Then
    With rst
        .AddNew
        .Fields("異體字").Value = rngx
        .Fields("正字").Value = x
        .Fields("兼正字").Value = bz
        .Update
        .Requery
    End With
End If
x = 異體字轉正_取得取代資料(rst)
rng.Document.Range.Find.Execute _
    rngx, , , , , , True, flaSel, , x, wdReplaceAll
rst.Close: cnt.Close
End Sub
Sub 異體字轉正_訂正資料() '20210228
Dim cnt As New ADODB.Connection, rst As New ADODB.Recordset, rng As Range, x As String, drv As String, db As New dBase
If Selection.Type = wdSelectionIP Then
    Set rng = Selection.Characters(1)
Else
    Set rng = Selection.Range
End If
If Len(rng) > 2 Then MsgBox "非單字，請檢查！", vbExclamation: Exit Sub
'If VBA.Dir("H:\", vbVolume) = "" Then
'    drv = "D:\千慮一得齋\"
'Else
'    drv = "H:\我的雲端硬碟\私人\千慮一得齋(C槽版)\"
'End If
'cnt.Open _
    "Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=" & drv & "書籍資料\圖書管理附件\查字.mdb;"
db.cnt查字 cnt
rst.Open "select 異體字,正字 from 異體字轉正" & _
    " where strcomp(異體字,""" & rng & """)=0", cnt, adOpenKeyset, adLockOptimistic
If rst.RecordCount = 1 Then
    x = InputBox("請輸入正字")
    If x = "" Then rst.Close: Exit Sub
    If x Like "*[a-z0-9A-Z]*" Then
        MsgBox "非中文，請檢查！", vbExclamation
        rst.Close: Exit Sub
    End If
    If StrComp(x, rng) = 0 Then MsgBox "與要轉換之字相同，請檢查！", vbExclamation: rst.Close: cnt.Close: Exit Sub
    With rst
        .Fields("正字").Value = x
        .Update
    End With
Else
    MsgBox "找不著，請檢查！", vbExclamation: rst.Close: cnt.Close: Exit Sub
End If
rng.Document.Range.Find.Execute _
    rng, , , , , , True, wdFindContinue, , x, wdReplaceAll
rst.Close: cnt.Close
End Sub
Function 異體字轉正_取得取代資料(ByRef rst As ADODB.Recordset) As String
Dim x As String
If rst.RecordCount > 1 Then
    Do Until rst.EOF
        x = x & rst.Fields("正字").Value
        rst.MoveNext
    Loop
    異體字轉正_取得取代資料 = "★" & x & "◆"
ElseIf rst.RecordCount = 1 Then
    異體字轉正_取得取代資料 = rst.Fields("正字").Value
End If
End Function

Function 插入同音字()

End Function

Sub 漢字轉注音_國語辭典()
'Alt+Shift+z,Alt+8
Dim cnt As New ADODB.Connection, rst As New ADODB.Recordset, rng As word.Range, a, db As New dBase, zy As String, zys As String, ur As UndoRecord ', ed As Long
Dim rngZhuYin As word.Range
Set rng = Selection.Range: Set rngZhuYin = Selection.Range
rngZhuYin.SetRange rng.End, rng.End
db.cnt_重編國語辭典修訂本_資料庫 cnt

For Each a In rng.Characters
    rst.Open "select 注音一式  from [《重編國語辭典修訂本》 總表] where strcomp(字詞名,""" & a & """)=0 order by 注音一式", cnt, adOpenKeyset
    Do Until rst.EOF
        zy = zy & rst.Fields(0).Value & "，"
        rst.MoveNext
    Loop
    If zy <> "" Then
        zy = 文字處理.國語辭典注音文字處理(zy)
        zy = Left(zy, Len(zy) - 1)
        zys = zys & " " & zy
        zy = ""
    End If
    rst.Close
Next a
zys = Mid(zys, 2)
Set ur = SystemSetup.stopUndo()
If ActiveDocument.path = "" Then
    rng.Text = zys
Else
'    ed = rng.End
    rng.InsertAfter Chr(9) & zys
'    rng.SetRange ed, rng.End
End If
With rngZhuYin
    .SetRange rngZhuYin.start, rng.End
    .Font.Name = "標楷體"
    .Bold = False
End With
Set rngZhuYin = Nothing
rng.Select
SystemSetup.contiUndo ur
If rst.State <> adStateClosed Then rst.Close
cnt.Close: Set rng = Nothing: Set ur = Nothing
End Sub

Function 數字轉漢字2位數(yi As Byte)
Const digit As Byte = 10
Dim q As Byte, r As Byte, ay
ay = Array("元", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十")
    r = yi Mod digit: q = (yi - r) / digit
    If q = 0 Then
        If yi = 1 Then
            數字轉漢字2位數 = ay(q)
        Else
            數字轉漢字2位數 = ay(r)
        End If
    Else
        If r = 0 Then
            If q = 1 Then
                數字轉漢字2位數 = ay(10)
            Else
                數字轉漢字2位數 = ay(q) + ay(10)
            End If
        Else
            If q = 1 Then
                數字轉漢字2位數 = ay(10) + ay(r)
            Else
                數字轉漢字2位數 = ay(q) + ay(10) + ay(r)
            End If
        End If
    End If
End Function
