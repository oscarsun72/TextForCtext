VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_DictsURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbFullName As String, cnt As New ADODB.Connection

Private Sub addNewRecords(URL As String, w As String, dictID, rst As ADODB.Recordset, cnt As ADODB.Connection, Optional ZhuYin)
Dim lnks As New Links
If Not lnks.chkLinks_http_Dicts_toAddZhuYin(URL, w, dictID) Then Exit Sub
    rst.AddNew
    rst.Fields(0) = URL
    rst.Fields(1) = w
    rst.Fields(2) = dictID
    rst.Fields(3) = ZhuYin
    rst.Update
End Sub

Function add1URLTo1異體字字典(words As String) As String
Dim URL As String, rst As New ADODB.Recordset, lnks As New Links
Shell SystemSetup.getChrome() & " https://dict.variants.moe.edu.tw/variants/rbt/query_by_standard_tiles.rbt?command=clear"
1
    AppActivateDefaultBrowser
    URL = InputBox("plz  input the URL!")
    If URL = "" Then Exit Function
    If Left(URL, 4) <> "http" Then GoTo 1
    URL = lnks.trimLinks(URL)
        Dim ZhuYin
rep:
        AppActivateDefaultBrowser
reZY:
        ZhuYin = VBA.Replace(VBA.Trim(InputBox("plz input ZhuYin")), Chr(9), "")
        If ZhuYin = "" Then
            If MsgBox("none of  ZhuYin(注音) input , sure to leave?", vbOKCancel + vbDefaultButton2, "注音") = vbOK Then
                Exit Function
            Else
                GoTo reZY
            End If
        Else
            If Not 文字處理.is注音符號(Right(ZhuYin, 1)) Then
                If MsgBox("Check ZhuYin(注音) to reInput ?", vbOKCancel, "注音") = vbOK Then
                    GoTo reZY
                Else
                    Exit Function
                End If
            End If
        End If
rst.Open "Select URL,word,dict,注音 from dicturl where strcomp(word,""" & words & """)=0 and url=""" & URL & """", cnt, adOpenKeyset, adLockOptimistic
If rst.RecordCount = 0 Then addNewRecords URL, words, 2, rst, cnt, ZhuYin
add1URLTo1異體字字典 = URL
rst.Close
End Function

'Function add1URLTo1國語辭典(words As String) As String
'Dim URL As String, rst As New ADODB.Recordset
'rst.Open "Select URL,word,dict from dicturl", cnt, adOpenKeyset, adLockOptimistic
'Shell SystemSetup.getChrome & " http://dict.revised.moe.edu.tw/cbdic/search.htm"
'1
'    URL = InputBox("plz  input the URL!")
'    If URL = "" Then Exit Function
'    If Left(URL, 4) <> "http" Then GoTo 1
'addNewRecords URL, words, 1, rst
'add1URLTo1國語辭典 = URL
'rst.Close
'End Function
Function add1URLTo1Dict(words As String, dictID As Byte) As String
Dim URL As String, rst As New ADODB.Recordset, lnks As New Links
rst.Open "select dicturl from DictsNameList where dictid  =" & dictID, cnt, adOpenDynamic
URL = rst.Fields(0).Value
rst.Close
'open in the same windows
'Shell SystemSetup.getChrome() & " " & lnks.trimLinks(URL)
'open in the mew windows
Shell SystemSetup.getChrome() & " "
Shell SystemSetup.getChrome() & " " & lnks.trimLinks(URL)

1
    AppActivateDefaultBrowser
    URL = InputBox("plz  input the URL!")
    If URL = "" Then AppActivate word.ActiveDocument.ActiveWindow.Caption & " - word": Exit Function
    If Left(URL, 4) <> "http" Then GoTo 1
    URL = lnks.trimLinks(URL)
'        Dim zhuYin
'rep:
'        zhuYin = VBA.Trim(InputBox("plz input ZhuYin"))
'        If zhuYin = "" Then
'            rst.Close
'            Exit Function
'        End If
rst.Open "Select URL,word,dict,注音 from dicturl where strcomp(word,""" & words & """)=0 and url=""" & URL & """", cnt, adOpenKeyset, adLockOptimistic
If rst.RecordCount = 0 Then addNewRecords URL, words, dictID, rst, cnt ', zhuYin
add1URLTo1Dict = URL
rst.Close
End Function

Sub Word_AfterUpdate(words As String)
Dim dictID, rst As New ADODB.Recordset, w As String, URL As String, u As String, lnks As New Links
If words = "" Then words = Selection.Text
If words = "" Then Exit Sub
w = words
dictID = InputBox("plz  input the Dictionary ID!" & vbCr & vbLf & "國語辭典=1" & vbCr & vbLf & "異體字字典=2" _
        & vbCr & vbLf & "漢語大字典=3" & vbCr & vbLf & "漢語大詞典=4", , 1)
If dictID <> "" Then
    dictID = CLng(dictID)
    rst.Open "Select URL,word,dict,注音 from dicturl where strcomp(word , """ & w & """)=0 and dict=" & dictID, cnt, adOpenKeyset, adLockOptimistic
Else
    rst.Open "Select URL,word,dict,注音 from dicturl where strcomp(word , """ & w & """)=0", cnt, adOpenKeyset, adLockOptimistic
End If
If rst.RecordCount > 0 Then
    Do Until rst.EOF
        u = lnks.trimLinks(rst.Fields(0).Value)
        URL = URL & " " & u
        rst.MoveNext
    Loop
    'FollowHyperlink URL
    Shell getChrome & URL
Else
1
    AppActivateDefaultBrowser
    URL = InputBox("plz  input the URL!")
    If URL = "" Then rst.Close: Exit Sub
    If Left(URL, 4) <> "http" Then GoTo 1
    dictID = CLng(InputBox("plz  input the Dictionary ID!" & vbCr & vbLf & "國語辭典=1" & vbCr & vbLf & "異體字字典=2" _
        & vbCr & vbLf & "漢語大字典=3" & vbCr & vbLf & "漢語大詞典=4", , 1))
    URL = lnks.trimLinks(URL)
    addNewRecords URL, w, dictID, rst, cnt
End If
rst.Close
End Sub

Private Sub Class_Initialize()
Dim pth As New Paths
dbFullName = pth.getdb_查字FullName()
cnt.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbFullName
End Sub

Private Sub Class_Terminate()
cnt.Close
End Sub

Sub getDictRecS(w As String, ByRef rst As ADODB.Recordset, dictID As Byte)
If rst.State <> adStateClosed Then rst.Close
'rst.Open "Select 注音,dict,URL from dicturl where strcomp(word , """ & w & """)=0 and dict=" & DictID, cnt, adOpenKeyset, adLockOptimistic
rst.Open "Select 注音,釋義,URL,dict from dicturl where strcomp(word , """ & w & """)=0 and dict=" & dictID, cnt, adOpenKeyset, adLockOptimistic
End Sub

Sub getDictVariantsRecS(w As String, ByRef rst As ADODB.Recordset)
getDictRecS w, rst, 2
End Sub

Sub getDictHydzdRecS(w As String, ByRef rst As ADODB.Recordset)
getDictRecS w, rst, 3
End Sub

Sub getDictHydcdRecS(w As String, ByRef rst As ADODB.Recordset)
getDictRecS w, rst, 4
End Sub

'Function get1URLfor1(words As String, zhuYin As String) As String
'Dim u As String, rst As New ADODB.Recordset '國語辭典 dict.revised.
'Dim dbp As New Paths, cnt As New ADODB.Connection
'cnt.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbp.getdb_重編國語辭典修訂本_資料庫fullName()
'rst.Open "Select URL,字詞名,注音一式 from [《重編國語辭典修訂本》 總表] where strcomp(字詞名 , """ & words & """)=0" _
'             , cnt, adOpenKeyset, adLockOptimistic
'If rst.RecordCount > 0 Then
'    u = rst.Fields(0).Value
'    If Left(u, 1) = "#" Then u = Mid(u, 2, Len(u) - 2)
'    get1URLfor1 = u
'End If
'End Function
