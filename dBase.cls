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

Sub cnt�d�r(ByRef cnt As ADODB.Connection)
Dim pat As New Paths, p As String
p = pat.getdb_�d�rforinputFullName()
If p = "" Then
    MsgBox "�䤣��d�r.mdb", vbCritical
    End
Else
    cnt.Open cntStr & p
End If
Set pat = Nothing
End Sub
Sub cnt_���s��y���׭q��_��Ʈw(ByRef cnt As ADODB.Connection)
Dim pat As New Paths
cnt.Open cntStr & pat.getdb_���s��y���׭q��_��ƮwfullName
End Sub
Property Get ���s��y���׭q��_��Ʈw_�`��Name() As String
���s��y���׭q��_��Ʈw_�`��Name = "[�m���s��y���׭q���n �`��]"
End Property

Sub �~�y�j����[�`��(x As String)
dict�[�`�� x, 4
End Sub
Sub dict�[�`��(x As String, dictID As Byte)
Dim cnt As New ADODB.Connection, rst As New ADODB.Recordset, w As String, zy As String, ay
If VBA.InStr(x, Chr(9)) Then
    ay = VBA.Split(x, Chr(9))
ElseIf VBA.InStr(x, "�]") Then
    ay = VBA.Split(x, "�]")
Else
    Exit Sub
End If
w = ay(0): zy = ay(1)
cnt�d�r cnt
rst.Open "select �`�� from dicturl where strcomp(word, """ & w & """)=0 and dict=" & dictID, cnt, adOpenKeyset, adLockOptimistic
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

Sub ��y�����}��ID�|�ʪ̦C�X()
Dim rng As word.Range, a, x As String, cnt As New ADODB.Connection, rst As New ADODB.Recordset, ws As String
If Selection.Type = wdSelectionIP Then
    Set rng = ActiveDocument.Range
    rng.Find.Execute Chr(13), , , , , , , wdFindContinue, , "", wdReplaceAll
Else
    Set rng = Selection.Range
End If
cnt_���s��y���׭q��_��Ʈw cnt
For Each a In rng.Characters
    If InStr(x, a) = 0 Then
        If ��r�B�z.isSymbol(a) = False And ��r�B�z.is�`���Ÿ�(a, a) = False Then
            x = x & a
            rst.Open "select url,ID from " & ���s��y���׭q��_��Ʈw_�`��Name & " where strcomp(�r���W,""" & a & """)=0", cnt, adOpenKeyset
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
Set acs = GetObject(pt.getdb_�d�rforinputFullName()).Application 'https://stackoverflow.com/questions/28179066/access-vba-getobject
Set frmDictsurl = acs.screen.activeform
If frmDictsurl.Name <> "DictsURL" Then
   'Set ctl = Nothing
   Exit Function '�Ϥ��j�p�g�]case-sensitive�^
End If
Set ctl = frmDictsurl.Controls(ctlName)
Set acs = Nothing
End Function

Sub setWordControlValue(w As String)
If Not word.Tasks.Exists("�d�rforinput - [dictsurl]") Then Exit Sub
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
End Sub

Sub setURLControlValue(url As String)
Dim lnk As New Links
If Not word.Tasks.Exists("�d�rforinput - [dictsurl]") Then Exit Sub
If Not lnk.checkURL_single(url) Then Exit Sub

On Error Resume Next
Me.ctl("Text8").Value = url
Me.ctl("url").Value = url

End Sub
Sub setDictControlValue(dictID As Byte)
If Not word.Tasks.Exists("�d�rforinput - [dictsurl]") Then Exit Sub

On Error Resume Next
Me.ctl("Combo4").Value = dictID
End Sub
