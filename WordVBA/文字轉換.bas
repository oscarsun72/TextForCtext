Attribute VB_Name = "��r�ഫ"
Option Explicit
Public UserForm1TextBox1Value As String
Sub �~�r�����()
'F2

On Error Resume Next
Dim pt As String
pt = Replace(���o�ୱ���|, "Desktop", "Dropbox")
'Const fpath As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\ssz3\Google ���ݵw��\�p�H\VB\����.mdb" ';Mode=Share Deny None;Extended Properties="";Jet OLEDB:System database=C:\Users\ssz3\AppData\Roaming\Microsoft\Access\System.mdw;Jet OLEDB:Registry Path=Software\Microsoft\Office\16.0\Access\Access Connectivity Engine;Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=True;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False"
Dim fpath As String
fpath = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & pt & "\VS\����.mdb" ';Mode=Share Deny None;Extended Properties="";Jet OLEDB:System database=C:\Users\ssz3\AppData\Roaming\Microsoft\Access\System.mdw;Jet OLEDB:Registry Path=Software\Microsoft\Office\16.0\Access\Access Connectivity Engine;Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=True;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False"
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
cmd.CommandText = "SELECT �r.�r, ����.���� FROM ���� INNER JOIN (�r INNER JOIN �r_�`�� ON �r.�rID = �r_�`��.�rID) ON ����.����ID = �r_�`��.����ID" _
        & " WHERE (((�r.�r)=""" & x & """) ) order by �r_�`��.�r_�`��ID;"
cmd.CommandType = adCmdText
rstPinyin.Open cmd
If rstPinyin.EOF Then
    Beep
    MsgBox "�j���Ħn�G�|�L�u" & x & "�v���r�������B�`���A�Ф�ʸɥR�C�P���P�� �n�L��������", vbExclamation
Else
    Do Until rstPinyin.EOF
        'y = "�]" & rst.Fields("����").Value & "�^"
        y = rstPinyin.Fields("����").Value
        cmd.CommandText = "SELECT �r.�r FROM ���� INNER JOIN (�`�� INNER JOIN (�r INNER JOIN �r_�`�� ON �r.�rID = �r_�`��.�rID) ON �`��.�`��ID = �r_�`��.�`��ID) ON ����.����ID = �r_�`��.����ID " & _
                        "WHERE (((����.����) = """ & y & """) And ((�r.�u�Φr) = False) And ((�r_�`��.����)= 0)) ORDER BY �r.�r, �r.�r��;"
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
            .Font.Name = "�з���"
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
Sub ����r�ॿ() '20210228
Dim cnt As New ADODB.Connection, rst As New ADODB.Recordset, rng As Range, a, docAs As String, x As String, rngCC As Integer, flaSel As Byte, flgsel, drv As String, db As New dBase, ur As UndoRecord
'Alt+7
If Selection.Type = wdSelectionIP Then
    Set rng = ActiveDocument.Range
    flgsel = wdFindContinue
Else
    Set rng = Selection.Range
    flgsel = wdFindStop
End If
Set ur = SystemSetup.stopUndo("����r�ॿ")
'For Each a In rng.Characters
'    If InStr(docAs, a) = 0 Then docAs = docAs & a '�O�U�����ƪ��Φr��
'Next
rngCC = rng.Characters.Count
'If VBA.Dir("H:\", vbVolume) = "" Then
'    drv = "D:\�d�{�@�o�N\"
'Else
'    drv = "H:\�ڪ����ݵw��\�p�H\�d�{�@�o�N(C�Ѫ�)\"
'End If
'cnt.Open _
    "Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=" & drv & "���y���\�ϮѺ޲z����\�d�r.mdb;"
db.cnt�d�r cnt

For Each a In rng.Characters
    If InStr(docAs, a) = 0 Then
        docAs = docAs & a '�O�U�����ƪ��Φr��
        If rngCC = 1 Then
            rst.Open "select * from ����r�ॿ " & _
                "where strcomp(����r,""" & a & """)=0 " & _
                "order by �Ƨ�", cnt, adOpenKeyset, adLockReadOnly
        Else
            rst.Open "select * from ����r�ॿ " & _
                "where (strcomp(����r,""" & a & """)=0 " & _
                "and ����=false) order by �Ƨ�", cnt, adOpenKeyset, adLockReadOnly
        End If
        If rst.RecordCount > 0 Then
            x = ����r�ॿ_���o���N���(rst)
            With rng.Find
                .Text = a 'rst.Fields("����r")
                .Replacement.Text = x 'rst.Fields("���r")
                .Execute , , , , , , True, flaSel, , , wdReplaceAll
            End With
        End If
        rst.Close
    End If
Next
'�O�d�L�O����r�Τ]�n�A�i�@�ѧO
'For Each a In rng.Characters
'    Select Case a.Font.NameFarEast
'        Case "�L�n������", "hanaminb", "kaixinsongb"
'            rng.Font.NameFarEast = "�s�ө���-extb"
'    End Select
'Next a
endS:
If Not rst.State = adStateClosed Then rst.Close
cnt.Close: Set rst = Nothing: Set cnt = Nothing: Set a = Nothing: Set rng = Nothing
SystemSetup.playSound 2
SystemSetup.contiUndo ur
End Sub
Sub ����r�ॿ_�s�W���() '20210228
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
If Len(rng) > 2 Then MsgBox "�D��r�A���ˬd�I", vbExclamation: Exit Sub
If rng = "" Then Exit Sub

x = InputBox("�п�J���r", , x)
If InStr(x, "?") Then
    UserForm1.Show
    x = UserForm1TextBox1Value
End If
If x = "" Then Exit Sub
If x Like "*[a-z0-9A-Z]*" Then MsgBox "�D����A���ˬd�I", vbExclamation:        Exit Sub
If StrComp(x, rng) = 0 Then MsgBox "�P�n�ഫ���r�ۦP�A���ˬd�I", vbExclamation: Exit Sub
x = VBA.Trim(x)
rngx = VBA.Trim(rng.Text)
If MsgBox("�O�_�ݥ��r�H", vbQuestion + vbOKCancel + vbDefaultButton2) = vbOK Then bz = True
'If VBA.Dir("H:\", vbVolume) = "" Then
'    drv = "D:\�d�{�@�o�N\"
'Else
'    drv = "H:\�ڪ����ݵw��\�p�H\�d�{�@�o�N(C�Ѫ�)\"
'End If
'cnt.Open _
'    "Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=" & drv & "���y���\�ϮѺ޲z����\�d�r.mdb;"
db.cnt�d�r cnt
rst.Open "select * from ����r�ॿ" & _
    " where strcomp(����r,""" & rngx & """)=0 and " & _
        "strcomp(���r,""" & x & """)=0 " & _
    "order by �Ƨ�", cnt, adOpenKeyset, adLockOptimistic
If rst.RecordCount = 0 Then
    With rst
        .AddNew
        .Fields("����r").Value = rngx
        .Fields("���r").Value = x
        .Fields("�ݥ��r").Value = bz
        .Update
        .Requery
    End With
End If
x = ����r�ॿ_���o���N���(rst)
rng.Document.Range.Find.Execute _
    rngx, , , , , , True, flaSel, , x, wdReplaceAll
rst.Close: cnt.Close
End Sub
Sub ����r�ॿ_�q�����() '20210228
Dim cnt As New ADODB.Connection, rst As New ADODB.Recordset, rng As Range, x As String, drv As String, db As New dBase
If Selection.Type = wdSelectionIP Then
    Set rng = Selection.Characters(1)
Else
    Set rng = Selection.Range
End If
If Len(rng) > 2 Then MsgBox "�D��r�A���ˬd�I", vbExclamation: Exit Sub
'If VBA.Dir("H:\", vbVolume) = "" Then
'    drv = "D:\�d�{�@�o�N\"
'Else
'    drv = "H:\�ڪ����ݵw��\�p�H\�d�{�@�o�N(C�Ѫ�)\"
'End If
'cnt.Open _
    "Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=" & drv & "���y���\�ϮѺ޲z����\�d�r.mdb;"
db.cnt�d�r cnt
rst.Open "select ����r,���r from ����r�ॿ" & _
    " where strcomp(����r,""" & rng & """)=0", cnt, adOpenKeyset, adLockOptimistic
If rst.RecordCount = 1 Then
    x = InputBox("�п�J���r")
    If x = "" Then rst.Close: Exit Sub
    If x Like "*[a-z0-9A-Z]*" Then
        MsgBox "�D����A���ˬd�I", vbExclamation
        rst.Close: Exit Sub
    End If
    If StrComp(x, rng) = 0 Then MsgBox "�P�n�ഫ���r�ۦP�A���ˬd�I", vbExclamation: rst.Close: cnt.Close: Exit Sub
    With rst
        .Fields("���r").Value = x
        .Update
    End With
Else
    MsgBox "�䤣�ۡA���ˬd�I", vbExclamation: rst.Close: cnt.Close: Exit Sub
End If
rng.Document.Range.Find.Execute _
    rng, , , , , , True, wdFindContinue, , x, wdReplaceAll
rst.Close: cnt.Close
End Sub
Function ����r�ॿ_���o���N���(ByRef rst As ADODB.Recordset) As String
Dim x As String
If rst.RecordCount > 1 Then
    Do Until rst.EOF
        x = x & rst.Fields("���r").Value
        rst.MoveNext
    Loop
    ����r�ॿ_���o���N��� = "��" & x & "��"
ElseIf rst.RecordCount = 1 Then
    ����r�ॿ_���o���N��� = rst.Fields("���r").Value
End If
End Function

Function ���J�P���r()

End Function

Sub �~�r��`��_��y���()
'Alt+Shift+z,Alt+8
Dim cnt As New ADODB.Connection, rst As New ADODB.Recordset, rng As word.Range, a, db As New dBase, zy As String, zys As String, ur As UndoRecord ', ed As Long
Dim rngZhuYin As word.Range
Set rng = Selection.Range: Set rngZhuYin = Selection.Range
rngZhuYin.SetRange rng.End, rng.End
db.cnt_���s��y���׭q��_��Ʈw cnt

For Each a In rng.Characters
    rst.Open "select �`���@��  from [�m���s��y���׭q���n �`��] where strcomp(�r���W,""" & a & """)=0 order by �`���@��", cnt, adOpenKeyset
    Do Until rst.EOF
        zy = zy & rst.Fields(0).Value & "�A"
        rst.MoveNext
    Loop
    If zy <> "" Then
        zy = ��r�B�z.��y���`����r�B�z(zy)
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
    .Font.Name = "�з���"
    .Bold = False
End With
Set rngZhuYin = Nothing
rng.Select
SystemSetup.contiUndo ur
If rst.State <> adStateClosed Then rst.Close
cnt.Close: Set rng = Nothing: Set ur = Nothing
End Sub

Function �Ʀr��~�r2���(yi As Byte)
Const digit As Byte = 10
Dim q As Byte, r As Byte, ay
ay = Array("��", "�@", "�G", "�T", "�|", "��", "��", "�C", "�K", "�E", "�Q")
    r = yi Mod digit: q = (yi - r) / digit
    If q = 0 Then
        If yi = 1 Then
            �Ʀr��~�r2��� = ay(q)
        Else
            �Ʀr��~�r2��� = ay(r)
        End If
    Else
        If r = 0 Then
            If q = 1 Then
                �Ʀr��~�r2��� = ay(10)
            Else
                �Ʀr��~�r2��� = ay(q) + ay(10)
            End If
        Else
            If q = 1 Then
                �Ʀr��~�r2��� = ay(10) + ay(r)
            Else
                �Ʀr��~�r2��� = ay(q) + ay(10) + ay(r)
            End If
        End If
    End If
End Function
