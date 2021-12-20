VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const cntStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="

Sub cnt查字(ByRef cnt As ADODB.Connection)
Dim pat As New Paths, p As String
p = pat.getdb_查字FullName()
If p = "" Then
    MsgBox "找不到查字.mdb", vbCritical
    End
Else
    cnt.Open cntStr & p
End If
Set pat = Nothing
End Sub
Sub cnt_重編國語辭典修訂本_資料庫(ByRef cnt As ADODB.Connection)
Dim pat As New Paths
cnt.Open cntStr & pat.getdb_重編國語辭典修訂本_資料庫fullName
End Sub
Property Get 重編國語辭典修訂本_資料庫_總表Name() As String
重編國語辭典修訂本_資料庫_總表Name = "[《重編國語辭典修訂本》 總表]"
End Property

Sub 漢語大詞典加注音(x As String)
dict加注音 x, 4
End Sub
Sub dict加注音(x As String, dictID As Byte)
Dim cnt As New ADODB.Connection, rst As New ADODB.Recordset, w As String, zy As String, ay
If VBA.InStr(x, Chr(9)) Then
    ay = VBA.Split(x, Chr(9))
ElseIf VBA.InStr(x, "（") Then
    ay = VBA.Split(x, "（")
Else
    Exit Sub
End If
w = ay(0): zy = ay(1)
cnt查字 cnt
rst.Open "select 注音 from dicturl where strcomp(word, """ & w & """)=0 and dict=" & dictID, cnt, adOpenKeyset, adLockOptimistic
If rst.RecordCount = 0 Then
    rst.Close: cnt.Close
    Exit Sub
End If
With rst
    Do Until rst.EOF
        .Fields(0).Value = zy
        .Update
        .MoveNext
    Loop
    .Close
End With
End Sub

Sub 國語辭典網址及ID尚缺者列出()
Dim rng As word.Range, a, x As String, cnt As New ADODB.Connection, rst As New ADODB.Recordset, ws As String
If Selection.Type = wdSelectionIP Then
    Set rng = ActiveDocument.Range
    rng.Find.Execute Chr(13), , , , , , , wdFindContinue, , "", wdReplaceAll
Else
    Set rng = Selection.Range
End If
cnt_重編國語辭典修訂本_資料庫 cnt
For Each a In rng.Characters
    If InStr(x, a) = 0 Then
        If 文字處理.isSymbol(a) = False And 文字處理.is注音符號(a, a) = False Then
            x = x & a
            rst.Open "select url,ID from " & 重編國語辭典修訂本_資料庫_總表Name & " where strcomp(字詞名,""" & a & """)=0", cnt, adOpenKeyset
            If rst.RecordCount > 0 Then
                If VBA.IsNull(rst.Fields(0)) And VBA.IsNull(rst.Fields(1)) Then
                    ws = ws & a
                End If
            End If
            rst.Close
        End If
    End If
Next
If ws <> "" Then
    Dim d As Document
    Set d = Documents.Add
    d.Range.Text = ws
End If
Beep
End Sub
Function ctl(ctlName As String) As Object
Dim acs As Object, pt As New Paths, frmDictsurl As Object
Set acs = GetObject(pt.getdb_查字forinputFullName()).Application 'https://stackoverflow.com/questions/28179066/access-vba-getobject
Set frmDictsurl = acs.screen.activeform
If frmDictsurl.Name <> "DictsURL" Then
   'Set ctl = Nothing
   Exit Function '區分大小寫（case-sensitive）
End If
Set ctl = frmDictsurl.Controls(ctlName)
Set acs = Nothing
End Function

Sub setWordControlValue(w As String)
Const cz As String = "查字forinput - [dictsurl]"
On Error GoTo eH
If Not SystemSetup.appActivatedYet("msaccess") Then Exit Sub
'If Not word.Tasks.Exists(cz) Then Exit Sub ' 用 word.Tasks 有時會當掉
If w = "" Then Exit Sub
On Error Resume Next
'do until acs.screen.activeform.ActiveControl.Name <> "word" Then
'    For Each ctl In acs.screen.activeform.Controls
'
'    Next ctl
With Me.ctl("word")
    .Value = w
    .Requery
    If Len(w) > 1 Then
        .SelStart = 1
        .SelLength = Len(w) - 1
    End If
End With
Exit Sub
eH:
Select Case Err.Number
    Case 5 '程序呼叫或引數不正確
    Case Else
        MsgBox Err.Number & Err.Description
        'resume
End Select
End Sub

Sub setURLControlValue(url As String)
Dim lnk As New Links
If Not word.Tasks.Exists("查字forinput - [dictsurl]") Then Exit Sub
If Not lnk.checkURL_single(url) Then Exit Sub

On Error Resume Next
Me.ctl("Text8").Value = url
Me.ctl("url").Value = url

End Sub
Sub setDictControlValue(dictID As Byte)
If Not word.Tasks.Exists("查字forinput - [dictsurl]") Then Exit Sub

On Error Resume Next
Me.ctl("Combo4").Value = dictID
End Sub

