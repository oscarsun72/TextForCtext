Attribute VB_Name = "��r�B�z"
Option Explicit
Dim punctuationStr As String ' ���I�Ÿ��r��
Dim rst As Recordset, d As Object
Dim db As Database 'set db=CurrentDb _
�u��b�w�}�Ҥ�Access���ѷӤ@�� , �G���H�W���ѷ� _
,���HSet db = DBEngine.Workspaces(0).OpenDatabase _
    ("d:\�d�{�@�o�N\���y���\���W.mdb")!���Φ��ѷ�! _
    �Ѧ�: _
    Dim dbsCurrent As Database, dbsContacts As Database'�� CurrentDb ���u�W�����ƻs _
    Set dbsCurrent = CurrentDb _
    Set dbsContacts = DBEngine.Workspaces(0).OpenDatabase("Contacts.mdb")

Rem ���I�Ÿ��r��
Public Static Property Get PunctuationString() As String
If punctuationStr = "" Then _
    punctuationStr = "�]�^�C�u�v�y�z[]�i�j�e�f�m�n�q�r-��"",  �G�A�F�I�H?" _
        & "�B. :,;" _
        & "�K�K...!()-�P�E" & Chr(34) & Chr(-24153) & Chr(-24152) & Chr(-24155) & Chr(-24154) & ChrW(8218) '34�G���޸��C�j�����I�Ÿ��W�U���޸��B�W�U��޼ơB�r��
PunctuationString = punctuationStr
End Property
'Public Static Property Let Punctionn(ByVal vNewValue As Variant)
'
'End Property


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

Sub �r�W() '2002/11/10�nSub�~��bWord������!
On Error GoTo ���~�B�z
Dim ch, wrong As Long
'Dim chct As Long
Dim StTime As Date, EndTime As Date
'Dim x As Long, firstword As String '�ýX�ˬd!2002/11/13
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.docmd.SelectObject blog.myaccess.acTable, "�r�W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb '�@�w�n�[��d��!!�g���H�U��i!
'�H�W�i�֦��U�G���Y�i!�����|��ܦb����W,�u��@����p���!(��OpenCurrentDatabase���u�W����)
'Set db = d.DBEngine.OpenDatabase("d:\�d�{�@�o�N\���y���\���W.mdb")
'Set db = d.DBEngine.Workspaces(0).OpenDatabase("d:\�d�{�@�o�N\���y���\���W.mdb")
Set rst = db.OpenRecordset("�r�W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM �r�W��"
End If
StTime = Time
With ActiveDocument
    For Each ch In .Characters '���ýX�r��ch�|�Ǧ^"?"�ܦ��F�B��βŸ�
        wrong = wrong + 1 '�˵���!
'        If wrong = 373 Then MsgBox "Check!!" '�ˬd��!
        If wrong Mod 27250 = 0 Then 'If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
            MsgBox "�]�t�έt���F�췥��,�аȥ�������Access���}��ƪ������,�A�^�ӫ��U�T�w���s�~��!!" _
                , vbExclamation, "���t�έ��n��T��"
'        ElseIf wrong = 49761 Then
'            MsgBox "���ˬd!!"
        End If
'        If wrong Mod 1000 = 0 Then Debug.Print wrong
'        Debug.Print ch & vbCr & "--------"
        '����r���B�_��r�����p!
'        If Right(ch, 1) <> Chr(10) Or Left(ch, 1) <> Chr(13) Then
        Select Case Asc(ch)
            Case Is <> 13, 10
        With rst
11          .FindFirst "�r�J like '" & ch & "'"
12          If .NoMatch Then
                .AddNew
                rst("�r�J") = ch
                rst("����") = 1
                rst("Asc") = Asc(ch)
                rst("AscW") = AscW(ch)
    '            On Error GoTo ����
                .Update
            Else '���ýX�r��,�|��������B�⤸"?"(Asc(ch)=63),�h�i��b��󤤲Ĥ@���X�{���r�|�~�W����
                '���~�p"�b"�r��(�bWord�����J���Ÿ����̫�@��)�r,��|�P�P�Φr�P�r���X(Asc), _
                ���b�Ÿ����o�����P��m,�N���P�r!�b�έp��,�t�Υ�|�~��b�@�_! _
                �o�I�ٶ��n�J�A!2002/11/13���ծ�,���ɤS�|���}!(��Asc�h�ۦP!)
'                If .AbsolutePosition < 1 And ch Like "?" And Not rst("�r�J") = "?" Then
'                    'If x = 1 Then MsgBox "���ýX�r,���ƱN�[�J�Ĥ@�ӥX�{���r��!!"
'                    MsgBox "���ýX�r,���ƱN�[�J�Ĥ@�ӥX�{���r��!!"
'                    AppActivate "Microsoft Word"
'                    Selection.Collapse
'                    Selection.SetRange wrong + ActiveDocument.Paragraphs.Count / 2, wrong + 1 '�N�ӶýX�r���
'                    x = x + 1
'                End If
                .edit
                rst("����") = rst("����") + 1
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
'        firstword = "�����ýX�r�[�J�Ĥ@�r:�u" & rst("�r�J") & "�v���@��" & x & "��!!"
'    Else
'        firstword = "����ߧa!�ýX�r��έp���T!!��"
'    End If
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count & vbCr '_
'        & firstword
'    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
'        & vbCr & "���Ӯ�:" & DateDiff("n", StTime, EndTime) & "������" _
'        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
End If
d.docmd.OpenTable "�r�W��", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number
    Case Is = 91, 3078 '�ѷӤ���DataBase�������
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
'        d.CurrentDb.Close
'        Set db = DBEngine.Workspaces(0).OpenDatabase("d:\�d�{�@�o�N\���y���\���W.mdb")
''        Debug.Print Err.Description '�ˬd��!
'        Resume
'    Case Is = 3163 '����r���B�_��r�����p!
'        If Right(ch, 1) = Chr(10) Then
'            ch = Left(ch, Len(ch) - 1)
'        ElseIf Left(ch, 1) = Chr(13) Then
'            ch = Right(ch, Len(ch) - 1) '��If Asc(ch)=13
'        End If
'        Resume 11
    Case Is = 93 '��[]���B�⦡�S��r���ҳ]�������
        rst.FindFirst "asc(�r�J) = " & Asc(ch)
        Resume 12
'    Case Is = -2147023170
'        MsgBox Err.Number & ":" & Err.Description
'        MsgBox Err.LastDllError & "." & Err.Source
'        Set d = CreateObject("access.application")
'        d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
'        d.UserControl = True
'        Resume
'    Case Is = 462 '"���ݦ��A�����s�b�εL�k�ϥ�"
'        'd.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
''        Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
'        Set db = d.CurrentDb
'        Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
'        Resume
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub ���W() '2002/11/10
On Error GoTo ���~�B�z
Dim WD, wrong As Long
Dim wrongmark As Integer ', wdct As Long
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True '�p�G��False�hdb.close�|������Ʈw!
'd.UserControl = False
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.docmd.SelectObject d.acTable, "���W��", True
'd.Visible = True '�ˬd��UserControl=True�h�����Ϸ|�P�~!
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then db.Execute "DELETE * FROM ���W��"
StTime = Time
With ActiveDocument
    For Each WD In .words
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 1000 = 0 Then Debug.Print wrong
'        Debug.Print wd & vbCr & "--------"
        If Len(WD) > 1 And right(WD, 1) = " " Then
            hfspace = hfspace + 1 '�p��
            GoTo retry '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        rst.FindFirst "���J like '" & WD & "'"
        If rst.NoMatch Then
            rst.AddNew
            rst("���J") = WD
'            On Error GoTo ����
            rst.Update
        Else
            rst.edit
            rst("����") = rst("����") + 1
            rst.Update
        End If
'        wrong = 1
'        wdct = .Words.Count
'        wdct = Selection.StoryLength
'        instr(1+
'        .Select
retry:  Next WD
End With
EndTime = Time
AppActivate "Microsoft word"
MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
    & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
    & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��"
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
End If
d.docmd.OpenTable "���W��", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
'����:
'    wrongmark = Err.Number
''    Err.Description = wd
'    If wrongmark = 3022 Then '���ƤF
''        wrong = wrong + 1
''        rst.Seek "=", "���J"
'        rst.FindFirst "���J like '" & wd & "'"
'        rst.Edit
'        rst("����") = rst("����") + 1
'        rst.Update
'        Resume retry
'    Else
'        MsgBox "�����~,���ˬd!!" & Err.Description, vbExclamation
'    End If
End Sub
Sub �i�����W() '2002/11/10�nSub�~��bWord������!'2005/4/21���k�b�]�j�ɮ׮ɤӨS�Ĳv�F!!�]�F3��3�]300��������ɨ�1-3�r���]����!
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As Byte
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Dim length As Byte 'As String
Dim Dw As String, dwL As Long
length = InputBox("�Ы��w���R���J���W��,�̦h���Ӧr", , "5")
If length = "" Or Not IsNumeric(length) Then End
If CByte(length) < 1 Or CByte(length) > 5 Then End
Options.SaveInterval = 0 '�����۰��x�s
StTime = Time
Set d = CreateObject("access.application")
'��Set d = CreateObject("Access.Application.9")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.docmd.SelectObject d.acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
End If
'With ActiveDocument
With ActiveDocument
    Dw = .Content '��󤺮e
    dwL = Len(Dw) '������
    .Close
End With
    For phralh = 1 To length 'CByte(length)
'    For phralh = 1 To 5 '�ȩw�̪���5�Ӧr�c������(���i��@�ܼ�)
        For phra = 1 To dwL '.Characters.Count
            Select Case phralh
                Case Is = 1
                    If Err.LastDllError <> 0 Then
                        MsgBox Err.LastDllError & ":" & Err.Description & "Err.Number:" & Err.Number
                        GoTo ���~�B�z
                    End If
'                    phras = .Characters(phra)'���k�ӺC!
                    phras = Mid(Dw, phra, 1)
                Case Is = 2
                    If Err.LastDllError <> 0 Then
                        MsgBox Err.LastDllError & ":" & Err.Description & "Err.Number:" & Err.Number
                        GoTo ���~�B�z
                    End If
'                    If phra + 1 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1)
                    If phra + 1 <= dwL Then phras = Mid(Dw, phra, 2)
                Case Is = 3
                    If Err.LastDllError <> 0 Then
                        MsgBox Err.LastDllError & ":" & Err.Description & "Err.Number:" & Err.Number
                        GoTo ���~�B�z
                    End If
'                    If phra + 2 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1) & _
                            .Characters(phra + 2)
                    If phra + 2 <= dwL Then phras = Mid(Dw, phra, 3)
                Case Is = 4
                    On Error GoTo ���~�B�z
                    If Err.LastDllError <> 0 Then
                        MsgBox Err.LastDllError & ":" & Err.Description & "Err.Number:" & Err.Number
                        GoTo ���~�B�z
                    End If
'                    If phra + 3 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1) & _
                            .Characters(phra + 2) & .Characters(phra + 3)
                    If phra + 3 <= dwL Then phras = Mid(Dw, phra, 3)
                Case Is = 5
                    On Error GoTo ���~�B�z
                    If Err.LastDllError <> 0 Then
                        MsgBox Err.LastDllError & ":" & Err.Description & "Err.Number:" & Err.Number
                        GoTo ���~�B�z
                    End If
'                    If phra + 4 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1) & _
                            .Characters(phra + 2) & .Characters(phra + 3) & _
                            .Characters(phra + 4)
                    If phra + 4 <= dwL Then phras = Mid(Dw, phra, 3)
            End Select
            If Len(phras) > 1 And right(phras, 1) = " " Then
                hfspace = hfspace + 1 '�p��
                GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
            End If
            '�����i�J�U�@�Ӧr����
            wrong = wrong + 1 '�˵���!
            If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
                DoEvents 'MsgBox "���ˬd!!"
    '        ElseIf wrong = 49761 Then
    '            MsgBox "���ˬd!!"
            End If
'            if rst Set rst = CurrentDb.OpenRecordset("SELECT  ���W��.* FROM ���W�� WHERE (((���W��.���J) like '" & phras & "'));")
            With rst
'                If .RecordCount = 0 Then
                .FindFirst "���J like '" & phras & "'"
                If .NoMatch Then
'                    .MoveLast
                    .AddNew
                    rst("���J") = phras
'                    rst("����") = 1'�w�]�Ȥw��1
                    On Error GoTo ���~�B�z
                    .Update 'dbUpdateBatch, True
                Else
1                   .edit
                    rst("����") = rst("����") + 1
                    .Update
                End If
'                .Close
            End With
11      Next phra
2   Next phralh
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & dwL '.Characters.Count
'End With
'd.Visible = True
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
End If
d.docmd.OpenTable "���W��", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access'2002/11/15
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 3022
        rst.Requery
        rst.FindFirst "���J like '" & Trim(phras) & "'"
        GoTo 1
    Case Is = 5941 '���X�����������s�b(���W�L������!)
        GoTo 2
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub �i�����W1() '2002/11/15�nSub�~��bWord������!
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As Byte
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Dim length As String
Dim i As Byte, j As Byte
length = InputBox("�Ы��w���R���J���W��,�̦h255�Ӧr", , "5")
If length = "" Or Not IsNumeric(length) Then End
If CByte(length) < 1 Or CByte(length) > 255 Then End
Options.SaveInterval = 0 '�����۰��x�s
StTime = Time
Set d = CreateObject("access.application")
'��Set d = CreateObject("Access.Application.9")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.docmd.SelectObject d.acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
End If
j = CByte(length)
With ActiveDocument
    For phralh = 1 To j
'    ��ȩw�̪���5�Ӧr�c������,����@�ܼ�j,�h����Byte�j�p��!
        For phra = 1 To .Characters.Count
            If phra + (phralh - 1) <= .Characters.Count Then
                phras = ""
                For i = 0 To phralh - 1
                    phras = phras & .Characters(phra + i)
                Next i
            End If
            If Len(phras) > 1 And right(phras, 1) = " " Then
                hfspace = hfspace + 1 '�p��
                GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
            End If
            '�����i�J�U�@�Ӧr����
            wrong = wrong + 1 '�˵���!
            If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
                MsgBox "���ˬd!!"
    '        ElseIf wrong = 49761 Then
    '            MsgBox "���ˬd!!"
            End If
            With rst
                .FindFirst "���J like '" & phras & "'"
                If .NoMatch Then
    '                .MoveLast
                    .AddNew
                    rst("���J") = phras
                    rst("����") = 1
                    On Error GoTo ���~�B�z
                    .Update 'dbUpdateBatch, True
                Else
1                   .edit
                    rst("����") = rst("����") + 1
                    .Update
                End If
            End With
11      Next phra
2   Next phralh
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
'd.Visible = True
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
End If
d.docmd.OpenTable "���W��", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 3022
        rst.Requery
        rst.FindFirst "���J like '" & Trim(phras) & "'"
        GoTo 1
    Case Is = 5941 '���X�����������s�b(���W�L������!)
        GoTo 2
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub ���w�r�Ƶ��W() '2002/11/11
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
phralh = InputBox("�ХΪ��ԧB�Ʀr���w�����զ��r��,�̦h�r�Ƭ��u11�v!", "���w���J�r��", "2")
If phralh = "" Or Not IsNumeric(phralh) Then Exit Sub
If CByte(phralh) > 11 Or CByte(phralh) < 1 Then Exit Sub
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.docmd.SelectObject d.acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
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
        If Len(phras) > 1 And right(phras, 1) = " " Then
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "���W��", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "���W��", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub ���w11�r�Ƶ��W()     '2002/11/15'�H������,�i�@���w�����w�r�ƪ��U�ӵ{��(���Ҭ�11�Ӧr���d��)
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
'phralh = InputBox("�ХΪ��ԧB�Ʀr���w�����զ��r��,�̦h�r�Ƭ��u11�v!", "���w���J�r��", "2")
'If phralh = "" Or Not IsNumeric(phralh) Then Exit Sub
'If CByte(phralh) > 11 Or CByte(phralh) < 1 Then Exit Sub
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.docmd.SelectObject d.acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
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
        If Len(phras) > 1 And right(phras, 1) = " " Then
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "���W��", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "���W��", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub ���w10�r�Ƶ��W() '2002/11/15
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.docmd.SelectObject d.acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
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
        If Len(phras) > 1 And right(phras, 1) = " " Then
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "���W��", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "���W��", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub ���w9�r�Ƶ��W()  '2002/11/15
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.docmd.SelectObject d.acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
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
        If Len(phras) > 1 And right(phras, 1) = " " Then
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "���W��", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "���W��", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub


Sub ���w8�r�Ƶ��W()   '2002/11/15
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.docmd.SelectObject d.acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 7 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2) & .Characters(phra + 3) & _
                    .Characters(phra + 4) & .Characters(phra + 5) & _
                    .Characters(phra + 6) & .Characters(phra + 7)
        If Len(phras) > 1 And right(phras, 1) = " " Then
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "���W��", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "���W��", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub ���w6�r�Ƶ��W()    '2002/11/15
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.docmd.SelectObject d.acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 5 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2) & .Characters(phra + 3) & _
                    .Characters(phra + 4) & .Characters(phra + 5)
        If Len(phras) > 1 And right(phras, 1) = " " Then
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "���W��", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "���W��", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub ���w5�r�Ƶ��W()     '2002/11/15
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.docmd.SelectObject d.acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 4 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2) & .Characters(phra + 3) & _
                    .Characters(phra + 4)
        If Len(phras) > 1 And right(phras, 1) = " " Then
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "���W��", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "���W��", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub ���w4�r�Ƶ��W()       '2002/11/15
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.docmd.SelectObject d.acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 3 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2) & .Characters(phra + 3)
        If Len(phras) > 1 And right(phras, 1) = " " Then
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "���W��", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "���W��", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub ���w3�r�Ƶ��W()      '2002/11/15
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.docmd.SelectObject d.acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 2 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2)
        If Len(phras) > 1 And right(phras, 1) = " " Then
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "���W��", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "���W��", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub ���w2�r�Ƶ��W()       '2002/11/15
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.docmd.SelectObject d.acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 1 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1)
        If Len(phras) > 1 And right(phras, 1) = " " Then
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "���W��", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "���W��", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub ���w1�r�Ƶ��W()        '2002/11/15
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.docmd.SelectObject d.acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
            phras = .Characters(phra)
        If Len(phras) > 1 And right(phras, 1) = " " Then
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "���W��", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "���W��", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub ���w7�r�Ƶ��W()      '2002/11/15'�H������,�i�@���w�����w�r�ƪ��U�ӵ{��(���Ҭ�7�Ӧr���d��)
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras As String, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
'phralh = InputBox("�ХΪ��ԧB�Ʀr���w�����զ��r��,�̦h�r�Ƭ��u11�v!", "���w���J�r��", "2")
'If phralh = "" Or Not IsNumeric(phralh) Then Exit Sub
'If CByte(phralh) > 11 Or CByte(phralh) < 1 Then Exit Sub
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.docmd.SelectObject d.acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
End If
StTime = Time
With ActiveDocument
    For phra = 1 To .Characters.Count
        If phra + 6 <= .Characters.Count Then _
            phras = .Characters(phra) & .Characters(phra + 1) & _
                    .Characters(phra + 2) & .Characters(phra + 3) & _
                    .Characters(phra + 4) & .Characters(phra + 5) & _
                    .Characters(phra + 6)
        If Len(phras) > 1 And right(phras, 1) = " " Then
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "���W��", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "���W��", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub ���w�r�Ƶ��W1() '2002/11/15'�į���C!
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Dim a1, i As Byte, j As Byte
phralh = InputBox("�ХΪ��ԧB�Ʀr���w�����զ��r��,�̦h�r�Ƭ��u255�v!", "���w���J�r��", "2")
If phralh = "" Or Not IsNumeric(phralh) Then Exit Sub
If CByte(phralh) > 255 Or CByte(phralh) < 1 Then Exit Sub
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.docmd.SelectObject d.acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
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
        If Len(phras) > 1 And right(phras, 1) = " " Then
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "���W��", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "���W��", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub
Sub ���w�r�Ƶ��W2() '2002/11/15�į�P��]�p�t���h,���i�ܼƤ�!
On Error GoTo ���~�B�z
Dim wrong As Long, phra As Long, phras, phralh As String
Dim StTime As Date, EndTime As Date
Dim hfspace As Long
Dim i As Byte, j As Byte
phralh = InputBox("�ХΪ��ԧB�Ʀr���w�����զ��r��,�̦h�r�Ƭ��u255�v!", "���w���J�r��", "2")
If phralh = "" Or Not IsNumeric(phralh) Then Exit Sub
If CByte(phralh) > 255 Or CByte(phralh) < 1 Then Exit Sub
Options.SaveInterval = 0 '�����۰��x�s
Set d = CreateObject("access.application")
d.UserControl = True
d.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\���W.mdb", False
d.docmd.SelectObject d.acTable, "���W��", True
'd.Visible = True '�ˬd��
Set db = d.CurrentDb
Set rst = db.OpenRecordset("���W��", dbOpenDynaset)
If rst.RecordCount > 0 Then '�n��o���������ƶ���MoveLast�����u�ݧP�_���S��������O���Y�i!
'rst���}�H��u�|���o�Ĥ@���O��!
'    db.Execute "DELETE �r�W��.* FROM �r�W��"
    db.Execute "DELETE * FROM ���W��"
End If
StTime = Time
j = CByte(phralh)
With ActiveDocument
    For phra = 1 To .Characters.Count
'        If j > 1 Then'�Y�ϬO��r�]�������O�B�z�F!!
            If phra + (phralh - 1) <= .Characters.Count Then
                phras = ""
                For i = 0 To j - 1
                    phras = phras & .Characters(phra + i)
                Next i
            End If
'        Else
'            phras = .Characters(phra)
'        End If
        If Len(phras) > 1 And right(phras, 1) = " " Then
            hfspace = hfspace + 1 '�p��
            GoTo 11 '�r��k��O�b�ΪŮ��,AccessUpdate�ɷ|�P�h,�B����J��L�N�N,�G���p!
        End If
        '�����i�J�U�@�Ӧr����
        wrong = wrong + 1 '�˵���!
'        If wrong Mod 29688 = 0 Then '��29688�ɷ|����OLE�S���^�������~,�G�b�����|��
'            MsgBox "���ˬd!!"
''        ElseIf wrong = 49761 Then
''            MsgBox "���ˬd!!"
'        End If
        With rst
            .FindFirst "���J like '" & phras & "'"
            If .NoMatch Then
                .AddNew
                rst("���J") = phras
'                rst("����") = 1'�w�]�Ȥw�w��1
                .Update 'dbUpdateBatch, True
            Else
                .edit
                rst("����") = rst("����") + 1
                .Update
            End If
        End With
11  Next phra
    EndTime = Time
    AppActivate "Microsoft word"
    MsgBox "�έp����!!" & vbCr & "(���@����F" & wrong & "�����ˬd��)" _
        & "���J�k��b�ΪŮ�Z" & hfspace & "��,�������p!" _
        & vbCr & "���Ӯ�:" & Format(EndTime - StTime, "n��s��") & "��" _
        & vbCr & "�r����=" & .Characters.Count
End With
If MsgBox("�n�Y���˵����G��?", vbYesNo + vbQuestion) = vbYes Then
'    Set d = GetObject("d:\�d�{�@�o�N\���y���\���W.mdb")
    AppActivate "Microsoft access"
    d.docmd.OpenTable "���W��", , d.acReadOnly
    d.docmd.Maximize
End If
d.docmd.OpenTable "���W��", , d.acReadOnly
d.docmd.Maximize
rst.Close: db.Close: Set d = Nothing
Options.SaveInterval = 10 '��_�۰��x�s
End '��Exit Sub�L�k�C������Access
���~�B�z:
Select Case Err.Number '�D���ޭȭ���
    Case Is = 91, 3078
        MsgBox "�ЦA���@��!", vbCritical
        'access.Application.Quit
        d.Quit
        End
    Case Else
        MsgBox Err.Number & ":" & Err.Description, vbExclamation
        Resume
End Select
End Sub

Sub ���r�W_old()
Dim DR As Range, d As Document, char, charText As String, preChar As String _
    , x() As String, xT() As Long, i As Long, j As Long, ExcelSheet  As Object, _
    ds As Date, de As Date '
Static xlsp As String
On Error GoTo ErrH:
'xlsp = "C:\Documents and Settings\Superwings\�ୱ\"
Set d = ActiveDocument
xlsp = ���o�ୱ���| & "\" 'GetDeskDir() & "\"
If Dir(xlsp) = "" Then xlsp = ���o�ୱ���| 'GetDeskDir ' "C:\Users\Wong\Desktop\" '& Replace(ActiveDocument.Name, ".doc", "") & "�r�W.XLS"
'If Dir(xlsp) = "" Then xlsp = "C:\Documents and Settings\Superwings\�ୱ\" & Replace(ActiveDocument.Name, ".doc", "") & "�r�W.XLS"
'xlsp = "C:\Documents and Settings\Superwings\�ୱ\" & Replace(ActiveDocument.Name, ".doc", "") & "�r�W.XLS"
xlsp = InputBox("�п�J�s�ɸ��|���ɦW(���ɦW,�t���ɦW)!" & vbCr & vbCr & _
        "�w�]�N�H��word����ɦW + ""�r�W.XLSX""�r��,�s��ୱ�W", "�r�W�լd", xlsp & Replace(ActiveDocument.Name, ".doc", "") & "�r�W" & StrConv(Time, vbWide) & ".XLSX")
If xlsp = "" Then Exit Sub

ds = VBA.Timer

With d
    For Each char In d.Characters
        charText = char
        If Not charText = Chr(13) And charText <> "-" And Not charText Like "[a-zA-Z0-9��-��]" Then
            'If Not charText Like "[a-z1-9]" & Chr(-24153) & Chr(-24152) & " �@�B'""�u�v�y�z�]�^�СH�I]" Then
'            If InStr(Chr(-24153) & Chr(-24152) & Chr(2) & "�E[]�e�f�����K�F,�A.�C�D �@�B'""����`\{}�a�b�u�v�y�z�]�^�m�n�q�r�СH�I]", charText) = 0 Then
            If InStr(ChrW(-24153) & ChrW(-24152) & Chr(2) & "�E[]�e�f�����K�F,�A.�C�D �@�B'""����`\{}�a�b�u�v�y�z�]�^�m�n�q�r�СH�I]", charText) = 0 Then
            'chr(2)�i��O���}�аO
                If preChar <> charText Then
                    'If UBound(X) > 0 Then
                        If preChar = "" Then 'If IsEmpty(X) Then'�p�G�O�@�}�l
                            GoTo 1
                        ElseIf UBound(Filter(x, charText)) Then ' <> charText Then  '�p�G�|�L���r
1                           ReDim Preserve x(i)
                            ReDim Preserve xT(i)
                            x(i) = charText
                            xT(i) = xT(i) + 1
                            i = i + 1
                        Else
                            GoSub �r�W�[�@
                        End If
                    'End If
                Else
                    GoSub �r�W�[�@
                End If
                preChar = char
            End If
        End If
    Next char
End With

Dim doc As New Document, Xsort() As String, u As Long ', xTsort() As Integer, k As Long, so As Long, ww As String
'ReDim Xsort(i) As String ', xtsort(i) as Integer
'ReDim Xsort(d.Characters.Count) As String
If u = 0 Then u = 1 '�Y�L����u�r�W�[�@:�v�Ƶ{��,�Y�L�W�L1�����r�W�A�h�@Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "�B" & x(j - 1) & _
                                �|�X���G�}�C���޶W�X�d�� 2015/11/5

ReDim Xsort(u) As String
Set ExcelSheet = CreateObject("Excel.Sheet")
With ExcelSheet.Application
    For j = 1 To i
        .Cells(j, 1) = x(j - 1)
        .Cells(j, 2) = xT(j - 1)
        Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "�B" & x(j - 1) 'Xsort(xT(j - 1)) & ww '�}�C�Ƨ�'2010/10/29
    Next j
End With
'Doc.ActiveWindow.Visible = False
'U = UBound(Xsort)
For j = u To 0 Step -1 '�}�C�Ƨ�'2010/10/29
    If Xsort(j) <> "" Then
        With doc
            If Len(.Range) = 1 Then '�|����J���e
                .Range.InsertAfter "�r�W = " & j & "���G�]" & Len(Replace(Xsort(j), "�B", "")) & "�r�^"
                .Range.Paragraphs(1).Range.Font.Size = 12
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "�s�ө���"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
                '.Range.Paragraphs(1).Range.Font.Bold = True
            Else
                .Range.InsertParagraphAfter
                .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
                .Range.InsertAfter "�r�W = " & j & "���G�]" & Len(Replace(Xsort(j), "�B", "")) & "�r�^"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Size = 12
                '.Range.Paragraphs(.Paragraphs.Count).Range.Bold = True
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "�s�ө���"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
            End If
            .Range.InsertParagraphAfter
            .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
            .Range.Paragraphs(.Paragraphs.Count).Range.Font.Size = 12
'            .Range.Paragraphs(.Paragraphs.Count).Range.Bold = False
            .Range.InsertAfter Replace(Xsort(j), "�B", Chr(9), 1, 1) 'chr(9)���w��r��(Tab���)
            .Range.InsertParagraphAfter
            If InStr(.Range.Paragraphs(.Paragraphs.Count).Range, "�r�W") = 0 Then
                .Range.Paragraphs(.Paragraphs.Count - 1).Range.Font.Name = "�з���"
            Else
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "�s�ө���"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
            End If
        End With
    End If
Next j

With doc.Paragraphs(1).Range
     .InsertParagraphBefore
     .Font.NameAscii = "times new roman"
    doc.Paragraphs(1).Range.InsertParagraphAfter
    doc.Paragraphs(1).Range.InsertParagraphAfter
    doc.Paragraphs(1).Range.InsertAfter "�A���Ѫ��奻�@�ϥΤF" & i & "�Ӥ��P���r�]�ǲΦr�P²�Ʀr�����X�֡^"
End With

doc.ActiveWindow.Visible = True
'

'U = UBound(xT)
'ReDim Xsort(U) As String, xTsort(U) As Long
'
'i = d.Characters
'For j = 1 To i '�μƦr�ۤ�
'    For k = 0 To U 'xT�}�C���C�Ӥ������Pj��
'        If xT(k) = j Then
'            Xsort(so) = x(k)
'            xTsort(so) = xT(k)
'            so = so + 1
'        End If
'    Next k
'Next j

'With doc
'    .Range.InsertAfter "�r�W=0001"
'    .Range.InsertParagraphAfter
'End With


' Cells.Select
'    Selection.Sort Key1:=Range("B1"), Order1:=xlDescending, Header:=xlGuess, _
'        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom


'Set ExcelSheet = Nothing'����|�Ϯ���
'Set d = Nothing
de = VBA.Timer
MsgBox "�����I" & vbCr & vbCr & "�O��" & left(de - ds, 5) & "��!"
ExcelSheet.Application.Visible = True
ExcelSheet.Application.UserControl = True
ExcelSheet.SaveAs xlsp '"C:\Macros\�u�uTEST.XLS"
doc.SaveAs Replace(xlsp, "XLS", "doc") '���j�p�g
'Doc.SaveAs "c:\test1.doc"
AppActivate "microsoft excel"
Exit Sub
�r�W�[�@:
For j = 0 To UBound(x)
    If x(j) = charText Then
        xT(j) = xT(j) + 1
        If u < xT(j) Then u = xT(j) '�O�U�̰��r�W,�H�K�Ƨ�(�N���ƧǤ��}�C�̰������ȳ]����,�h���|�W�X�}�C.
        '�h���@��]���n���ƧP�_�p��n�X��,�G�įण�W�ϴ�''�į��٬O�t���h��.
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

Function lEnglish() '�^��j�g�r��
Dim WD, wdct As Long, i As Byte
For i = 65 To 90
    Debug.Print Chr(i) & vbCr
Next
End Function
Function sEnglish() '�^��p�g�r��
Dim i As Byte
For i = 97 To 122
    Debug.Print Chr(i) & vbCr
Next
End Function

Function trimStrForSearch_PlainText(x As String) As String
Rem 20230128 �ѥf�~��C �]�u�u��chatGPT�j���ġGVBA Overload Functionality�G
'chatGPT�j���ķs�~�N���G �Q�а� VBA �O���O���๳ C# �@�˨禡��k�i�H �h���B�����]overload�^�H
'VBA (Visual Basic for Applications) �O�@�طL�n���{���y���A�D�n�Ω�۰ʤ� Microsoft Office ���ε{�����CVBA ���䴩�禡���h���M�����C�o�N���ۡA�z����b VBA ���w�q�㦳�ۦP�W�٦��ѼƤ��P���h�Ө禡�C

Dim ayToTrim As Variant, a As Variant
On Error GoTo eH
ayToTrim = Array(Chr(13), Chr(9), Chr(10), Chr(11), Chr(13) & Chr(7), Chr(13) & Chr(10))
x = VBA.Trim(x)
For Each a In ayToTrim
    'x = VBA.Replace(x, a, "")
    Do While VBA.left(x, Len(a)) = a
        x = VBA.Mid(x, Len(a) + 1)
    Loop
    Do While VBA.right(x, Len(a)) = a
        x = VBA.Mid(x, 1, Len(x) - Len(a))
    Loop
Next a
trimStrForSearch_PlainText = x
Exit Function
eH:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & Err.Description
'        Resume
End Select
End Function

Function trimStrForSearch(x As String, sl As word.Selection) As String
'https://docs.microsoft.com/zh-tw/dotnet/visual-basic/programming-guide/language-features/procedures/passing-arguments-by-value-and-by-reference
Dim ayToTrim As Variant, a As Variant, rng As Range, slTxtR As String
On Error GoTo eH
slTxtR = sl.Characters(sl.Characters.Count)
ayToTrim = Array(Chr(13), Chr(9), Chr(10), Chr(11), Chr(13) & Chr(7), Chr(13) & Chr(10))
x = VBA.Trim(x)
For Each a In ayToTrim
    'x = VBA.Replace(x, a, "")
    Do While VBA.left(x, Len(a)) = a
        x = VBA.Mid(x, Len(a))
    Loop
    Do While VBA.right(x, Len(a)) = a
        x = VBA.Mid(x, 1, Len(x) - Len(a))
    Loop
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


'Function Symbol() '���I�Ÿ���
'Dim f As Variant
'f = Array("�C", "�v", Chr(-24152), "�G", "�A", "�F", _
'    "�B", "�u", ".", Chr(34), ":", ",", ";", _
'    "�K�K", "...", "�^", ")", "-")  '���]�w���I�Ÿ��}�C�H�ƥ�
'                                'Chr(-24152)�O�u���v,��Asc��Ʀb���(.SelText)�u���v�ɨ��o;Chr(34):�u"�v
'End Function
Function isSymbol(ByVal a As String) As Boolean
Dim f As String
f = punctuationStr
If InStr(1, f, a, vbTextCompare) Then
    isSymbol = True
End If
End Function

Sub �M������B���Ҧ��Ÿ�() '�ѹϮѺ޲zsymbles�ҲղM�����I�Ÿ���s'�]�A���}�B�Ʀr
'Dim F, a As String, i As Integer
Dim f, i As Integer, ur As UndoRecord
SystemSetup.stopUndo ur, "�M������B���Ҧ��Ÿ�"
f = Array("-", "�P", "�E", "�C", "�v", Chr(-24152), "�G", "�A", "�F", _
    "�B", "�u", ".", Chr(34), ":", ",", ";", _
    "�K�K", "...", "�D", "�i", "�j", " ", "�m", "�n", "�q", "�r", "�H" _
    , "�I", "��", "��", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0" _
    , "�y", "�z", Chr(13), ChrW(9312), ChrW(9313), ChrW(9314), ChrW(9315), ChrW(9316) _
    , ChrW(9317), ChrW(9318), ChrW(9319), ChrW(9320), ChrW(9321), ChrW(9322), ChrW(9323) _
    , ChrW(9324), ChrW(9325), ChrW(9326), ChrW(9327), ChrW(9328), ChrW(9329), ChrW(9330) _
    , ChrW(9331), ChrW(8221), """") '���]�w���I�Ÿ��}�C�H�ƥ�
    '���ζ�A���Ȥ����N�I
    'a = ActiveDocument.Content
'    Set a = ActiveDocument.Range.FormattedText '�]�t�榡�ƪ���T
    For i = 0 To UBound(f)
        If InStr(Selection.Range.text, f(i)) Then
            'a = Replace(a, F(i), "")
            Selection.Range.Find.Execute f(i), True, , , , , , wdFindStop, True, "", wdReplaceAll
        End If
    Next
    'ActiveDocument.Content = a
SystemSetup.contiUndo ur
End Sub

Function is�`���Ÿ�(ByVal a As String, Optional rng As Variant) As Boolean
Dim f As String
On Error GoTo eH
If Len(a) > 1 Then Exit Function
f = "�t�u�v�w�x�y�z�{�|�}�~������������������������������������������������������  ��  ��  ��"
If a = ChrW(20008) Then
    If Not rng Is Nothing Then
        If rng.start = 0 Then
            If InStr(f, rng.Next.Characters(1)) Then
                is�`���Ÿ� = True
                Exit Function
            End If
        ElseIf rng.End = rng.Document.Range.End - 1 Then
            If InStr(f, rng.Previous.Characters(1)) Then
                is�`���Ÿ� = True
                Exit Function
            End If
        End If
    End If
Else
    If InStr(f, a) Then is�`���Ÿ� = True
End If
Exit Function
eH:
Select Case Err.Number
    Case 424 '���B�ݭn����
        Set rng = Nothing
        Resume
    Case Else
        MsgBox Err.Number & Err.Description
        Debug.Print Err.Number & Err.Description
End Select
End Function

Sub ����q���Ÿ�()
'��1�q���̫�()
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


Sub �y�r�r���ˬd() '�D�ө����ˬd,2004/8/23
Dim ch
For Each ch In ActiveDocument.Characters
'    If AscW(ch) < -1491 Or AscW(ch) > 19968 Then
    If Asc(ch) < -24256 Or (0 > Asc(ch) And Asc(ch) >= -1468) Then
        ch.Select
        ch.Font.Name = "EUDC"
    End If
Next ch
End Sub

Sub �`�}�Ÿ��m��() '2004/10/17
Dim WD As Range 'As Range 'Words����Y��@��Range����,���u�W����!
'Dim i As Long ' Integer
'�n�����������b��,�o��words�~�ॿ�T�P�_���Ʀr
���μƦr�ഫ���b�μƦr
With Selection '��H������(ActiveDocument),�����H����d���z,���]���ȦӼv�T,�@�o!
    If .Type = wdSelectionIP Then .Document.Select '�p�G�S������d��(�����J�I)�h�B�z������
    If .Document.path = "" Then
        For Each WD In .words
            '�n�O�Ʀr�B�e�ᤣ��[�����Ρe�f�~����I
            If Not WD.text Like "��" And Not WD.text Like "�e" And Not WD Like "[[]" And Not WD Like "[]]" Then
                If IsNumeric(WD) Then
                    If WD.End = .Document.Content.StoryLength Or WD.start = 0 Then GoTo w '��󤧭����t�~�B�z
                    If Not WD.Previous Like "��" And Not WD.Previous Like "�e" And Not WD.Previous Like "[[]" _
                        And Not WD.Next Like "��" And Not WD.Next Like "�f" And Not WD.Next Like "]" Then
w:                      If WD <= 20 Then 'Arial Unicode MS[����]��"�A����Ʀr"�u���G�Q��!
                            With WD
                                '����|����Selection���d��,�G������!
'                                .Select 'Words����Y��@��Range����,���u�W����!
                                .Font.Name = "Arial Unicode MS"
                                WD.text = ChrW((9312 - 1) + WD)
                            End With
                        Else '�W�L20�������}��
                            With WD
                                .text = "��" & WD.text & "��" '�[�A��
                            End With
        '                    MsgBox "���W�L20�������},�������I", vbCritical
        '                    Do Until .Undo(i) = False '�٭쪽�ܤ����٭�]�٭�Ҧ��ʧ@�^
        '                    i = i + 1
        '                    Loop
        '                    StatusBar = "Undo was successful " & i & " times!!" '�b���A�C��ܤ�r�I
        '                    Exit Sub
                        End If
                    End If
                End If
            End If
        Next
        MsgBox "���槹���I", vbInformation
    Else
        MsgBox "����󤣯�ާ@!", vbCritical
    End If
End With
End Sub

Sub ���μƦr�ഫ���b�μƦr() '2004/10/17-�ѹϮѺ޲z�ƻs�Ｖ���즡�Ф��n�A�|�v�T�r��
Dim FNumArray, HNumArray, i As Byte, e As Range
FNumArray = Array("��", "��", "��", "��", "��", "��", "��", "��", "��", "��")
HNumArray = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9")
With ActiveDocument
    For Each e In .Characters
        For i = 1 To UBound(FNumArray) + 1
            If e.text Like FNumArray(i - 1) Then
                e.text = HNumArray(i - 1)
        End If
        Next i
    Next e
End With
End Sub

Sub ������b��()
With Selection
    .Range = StrConv(.Range, vbNarrow)
End With
End Sub
Sub ��A����g�W��()
If Selection.Type = wdSelectionIP Then Selection.HomeKey wdStory: Selection.EndKey wdStory, wdExtend
Selection.text = Replace(Replace(Selection.text, "�]", "�q"), "�^", "�r")
End Sub


Sub �հɤ�r�Ц�() '2009/8/23
Register_Event_Handler
'���w��F2
' ����2 ����
' �������s�� 2009/8/23�A���s�� Oscar Sun
'
'    Selection.MoveDown Unit:=wdLine, Count:=2
'    Selection.EndKey Unit:=wdLine
'    Selection.MoveLeft Unit:=wdCharacter, Count:=1
'    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
If Selection.Type = wdSelectionIP Then Exit Sub
    With Selection.Font.Shading
        If InStr(ActiveDocument.Name, "�ƦL") Then
            .Parent.Color = wdColorRed
            .Texture = wdTextureNone
        Else
            If .Texture = wdTextureNone Then '�r������
                .Texture = wdTexture15Percent
                .ForegroundPatternColor = wdColorBlack
                .BackgroundPatternColor = wdColorWhite
                .Parent.Color = wdColorRed
            Else
                .Texture = wdTextureNone '�r������
                .Parent.Color = wdColorAutomatic
            End If
        End If
    End With
    If InStr(ActiveDocument.Name, "�ƦL") Then
        ActiveDocument.Save
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
                        MsgBox "�Ш�s�O���C�I�I", vbExclamation
                        Exit Sub
                    End If
                Next i
                .Cells(r, 1).Activate
                DoEvents
                .activesheet.Paste
                .Cells(r, 2).Value = Selection
                .Cells(r, 2).Font.Color = wdColorRed
                If Not Selection Like "*[�����������������������������U�@]*" Then
                    .Cells(r, 5) = Len(Selection)
                ElseIf Selection Like "*�@*" Then
                    .Cells(r, 5) = Len(Selection) - 1
                Else
                    .Cells(r, 5) = 1
                End If
                .ActiveWorkbook.Save
                .Cells(.ActiveCell.Row + 1, .ActiveCell.Column).Activate
            End With
        End With
        ��ЩҦb��m����
        OX.WinActivate "Adobe Reader"
        AppActivate "microsoft word"
    End If
End Sub

Sub ���}�s���e��[��A��()
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
        .TypeText text:="["
        .MoveLeft unit:=wdCharacter, Count:=1, Extend:=wdExtend
        .Font.Superscript = wdToggle
'        Selection.Copy
'        Selection.MoveRight Unit:=wdCharacter, Count:=3
'        Selection.MoveLeft Unit:=wdCharacter, Count:=1
'        Selection.Paste
'        Selection.MoveLeft Unit:=wdCharacter, Count:=1
'        Selection.Delete Unit:=wdCharacter, Count:=1
'        Selection.TypeText Text:="�n"
'        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.MoveRight unit:=wdCharacter, Count:=2
        'Selection.TypeBackspace
        Selection.TypeText text:="]"
        'Selection.MoveRight Unit:=wdCharacter, Count:=1
    Loop 'While .Find.Execute()
End With
End Sub

Sub �j���޸����x�W�޸�()
Dim a, b, i
a = Array(-24153, -24152, -24155, -24154)  '��,��,��,��
b = Array("�u", "�v", "�y", "�z")

With ActiveDocument.Range.Find
    For i = 0 To 3
        '.Text = a(i)
         '.Replacement.Text = b(i)
         .ClearFormatting
         .Execute Chr(a(i)), , , , , , , , , b(i), wdReplaceAll
    Next i
End With
End Sub


Sub ���r�W()
Dim d As Document, char, charText As String, preChar As String _
    , x() As String, xT() As Long, i As Long, j As Long, ds As Date, de As Date     '
'Dim ExcelSheet  As New Excel.Worksheet 'As Object,
'�o�O���e�H�����ޥΪ��覡�A�b�]�w�ޥζ��ؤ���ʥ[�J���g�k:https://hankvba.blogspot.com/2018/03/vba.html  �B http://markc0826.blogspot.com/2012/07/blog-post.html
'Dim xlApp As Excel.Application, xlBook As Excel.Workbook, xlSheet As Excel.Worksheet
''�o�N�O����ޥΡA�H�ۭq�s��Excel���O����k�ӹ�@(�p���g���t�G�O��ӭn��g���{���X�N�|����֡A�ܰʸ��p�A�B�]�����ANew�X�@�Ӱ������~�����G
Dim xlApp, xlBook, xlSheet
Set xlApp = Excel.Application
Set xlBook = Excel.Workbook
Set xlSheet = Excel.Worksheet
Dim ReadingLayoutB As Boolean
Static xlsp As String
On Error GoTo ErrH:
'xlsp = "C:\Documents and Settings\Superwings\�ୱ\"
Set d = ActiveDocument
xlsp = ���o�ୱ���| & "\" 'GetDeskDir() & "\"
If Dir(xlsp) = "" Then xlsp = ���o�ୱ���| 'GetDeskDir ' "C:\Users\Wong\Desktop\" '& Replace(ActiveDocument.Name, ".doc", "") & "�r�W.XLS"
'If Dir(xlsp) = "" Then xlsp = "C:\Documents and Settings\Superwings\�ୱ\" & Replace(ActiveDocument.Name, ".doc", "") & "�r�W.XLS"
'xlsp = "C:\Documents and Settings\Superwings\�ୱ\" & Replace(ActiveDocument.Name, ".doc", "") & "�r�W.XLS"
xlsp = InputBox("�п�J�s�ɸ��|���ɦW(���ɦW,�t���ɦW)!" & vbCr & vbCr & _
        "�w�]�N�H��word����ɦW + ""�r�W.XLSX""�r��,�s��ୱ�W", "�r�W�լd", xlsp & Replace(ActiveDocument.Name, ".doc", "") & "�r�W" & StrConv(Time, vbWide) & ".XLSX")
If xlsp = "" Then Exit Sub

ds = VBA.Timer

With d
    For Each char In d.Characters
        charText = char
        If InStr("()�G>" & Chr(13) & Chr(9) & Chr(10) & Chr(11) & ChrW(12), charText) = 0 And charText <> "-" And Not charText Like "[a-zA-Z0-9��-��]" Then
            'If Not charText Like "[a-z1-9]" & Chr(-24153) & Chr(-24152) & " �@�B'""�u�v�y�z�]�^�СH�I]" Then
'            If InStr(Chr(-24153) & Chr(-24152) & Chr(2) & "�E[]�e�f�����K�F,�A.�C�D �@�B'""����`\{}�a�b�u�v�y�z�]�^�m�n�q�r�СH�I]", charText) = 0 Then
            If InStr(ChrW(9312) & ChrW(-24153) & ChrW(-24152) & Chr(2) & "�E[]�e�f�����K�F,�A.�C�D �@�B'""����`\{}�a�b�u�v�y�z�]�^�m�n�q�r�СH�I]�����j�i~/�_�X" & Chr(-24152) & Chr(-24153), charText) = 0 Then
            'chr(2)�i��O���}�аO
                If preChar <> charText Then
                    'If UBound(X) > 0 Then
                        If preChar = "" Then 'If IsEmpty(X) Then'�p�G�O�@�}�l
                            GoTo 1
                        ElseIf UBound(Filter(x, charText)) Then ' <> charText Then  '�p�G�|�L���r
1                           ReDim Preserve x(i)
                            ReDim Preserve xT(i)
                            x(i) = charText
                            xT(i) = xT(i) + 1
                            i = i + 1
                        Else
                            GoSub �r�W�[�@
                        End If
                    'End If
                Else
                    GoSub �r�W�[�@
                End If
                preChar = char
            End If
        End If
    Next char
End With

Dim doc As New Document, Xsort() As String, u As Long ', xTsort() As Integer, k As Long, so As Long, ww As String
'ReDim Xsort(i) As String ', xtsort(i) as Integer
'ReDim Xsort(d.Characters.Count) As String
If u = 0 Then u = 1 '�Y�L����u�r�W�[�@:�v�Ƶ{��,�Y�L�W�L1�����r�W�A�h�@Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "�B" & x(j - 1) & _
                                �|�X���G�}�C���޶W�X�d�� 2015/11/5

ReDim Xsort(u) As String
'Set ExcelSheet = CreateObject("Excel.Sheet")
'Set xlApp = CreateObject("Excel.Application")
'Set xlBook = xlApp.workbooks.Add
'Set xlSheet = xlBook.Worksheets(1)
With xlSheet.Application
    For j = 1 To i
        .Cells(j, 1) = x(j - 1)
        .Cells(j, 2) = xT(j - 1)
        Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "�B" & x(j - 1) 'Xsort(xT(j - 1)) & ww '�}�C�Ƨ�'2010/10/29
    Next j
End With
'Doc.ActiveWindow.Visible = False
'U = UBound(Xsort)
For j = u To 0 Step -1 '�}�C�Ƨ�'2010/10/29
    If Xsort(j) <> "" Then
        With doc
            If Len(.Range) = 1 Then '�|����J���e
                .Range.InsertAfter "�r�W = " & j & "���G�]" & Len(Replace(Xsort(j), "�B", "")) & "�r�^"
                .Range.Paragraphs(1).Range.Font.Size = 12
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "�s�ө���"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
                '.Range.Paragraphs(1).Range.Font.Bold = True
            Else
                .Range.InsertParagraphAfter
                .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
                .Range.InsertAfter "�r�W = " & j & "���G�]" & Len(Replace(Xsort(j), "�B", "")) & "�r�^"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Size = 12
                '.Range.Paragraphs(.Paragraphs.Count).Range.Bold = True
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "�s�ө���"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
            End If
            .Range.InsertParagraphAfter
            .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
            .Range.Paragraphs(.Paragraphs.Count).Range.Font.Size = 12
'            .Range.Paragraphs(.Paragraphs.Count).Range.Bold = False
            .Range.InsertAfter Replace(Xsort(j), "�B", Chr(9), 1, 1) 'chr(9)���w��r��(Tab���)
            .Range.InsertParagraphAfter
            If InStr(.Range.Paragraphs(.Paragraphs.Count).Range, "�r�W") = 0 Then
                .Range.Paragraphs(.Paragraphs.Count - 1).Range.Font.Name = "�з���"
            Else
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "�s�ө���"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
            End If
        End With
    End If
Next j

With doc.Paragraphs(1).Range
     .InsertParagraphBefore
     .Font.NameAscii = "times new roman"
    doc.Paragraphs(1).Range.InsertParagraphAfter
    doc.Paragraphs(1).Range.InsertParagraphAfter
    doc.Paragraphs(1).Range.InsertAfter "�A���Ѫ��奻�@�ϥΤF" & i & "�Ӥ��P���r�]�ǲΦr�P²�Ʀr�����X�֡^"
End With

doc.ActiveWindow.Visible = True
'

'U = UBound(xT)
'ReDim Xsort(U) As String, xTsort(U) As Long
'
'i = d.Characters
'For j = 1 To i '�μƦr�ۤ�
'    For k = 0 To U 'xT�}�C���C�Ӥ������Pj��
'        If xT(k) = j Then
'            Xsort(so) = x(k)
'            xTsort(so) = xT(k)
'            so = so + 1
'        End If
'    Next k
'Next j

'With doc
'    .Range.InsertAfter "�r�W=0001"
'    .Range.InsertParagraphAfter
'End With


' Cells.Select
'    Selection.Sort Key1:=Range("B1"), Order1:=xlDescending, Header:=xlGuess, _
'        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom


'Set ExcelSheet = Nothing'����|�Ϯ���
'Set d = Nothing
de = VBA.Timer
If ReadingLayoutB Then d.ActiveWindow.View.ReadingLayout = Not d.ActiveWindow.View.ReadingLayout
MsgBox "�����I" & vbCr & vbCr & "�O��" & left(de - ds, 5) & "��!", vbInformation
xlSheet.Application.Visible = True
xlSheet.Application.UserControl = True
xlSheet.SaveAs xlsp '"C:\Macros\�u�uTEST.XLS"
doc.SaveAs Replace(xlsp, "XLS", "doc") '���j�p�g
Set Excel.Application = Nothing
Exit Sub
�r�W�[�@:
For j = 0 To UBound(x)
    If x(j) = charText Then
        xT(j) = xT(j) + 1
        If u < xT(j) Then u = xT(j) '�O�U�̰��r�W,�H�K�Ƨ�(�N���ƧǤ��}�C�̰������ȳ]����,�h���|�W�X�}�C.
        '�h���@��]���n���ƧP�_�p��n�X��,�G�įण�W�ϴ�''�į��٬O�t���h��.
        Exit For
    End If
Next j

Return
ErrH:
Select Case Err.Number
    Case 4605 '�\Ū�Ҧ�����s��'����k���ݩʵL�k�ϥΡA�]�����R�O�L�k�b�\Ū���ϥΡC
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
        doc.ActiveWindow.View.ReadingLayout = False
        doc.ActiveWindow.Visible = False
        ReadingLayoutB = True
        Resume
    Case Else
        MsgBox Err.Number & Err.Description, vbCritical 'STOP: Resume
        'Resume
        End
    
End Select
End Sub

Sub �����W() '�Ѥ��r�W���'2015/11/28
Dim d As Document, char, charText As String, preChar As String _
    , x() As String, xT() As Long, i As Long, j As Long, ds As Date, de As Date     '
'Dim ExcelSheet  As New Excel.Worksheet 'As Object,
'Dim xlApp As Excel.Application, xlBook As Excel.Workbook, xlSheet As Excel.Worksheet
Dim xlApp, xlBook, xlSheet
Set xlApp = Excel.Application
Set xlBook = Excel.Workbook
Set xlSheet = Excel.Worksheet
Dim ReadingLayoutB As Boolean
Static ln
Dim xlsp As String
On Error GoTo ErrH:
Set d = ActiveDocument
'If xlsp = "" Then xlsp = ���o�ୱ���| & "\" 'GetDeskDir() & "\"
'If Dir(xlsp) = "" Then xlsp = ���o�ୱ���| 'GetDeskDir
'xlsp = InputBox("�п�J�s�ɸ��|���ɦW(���ɦW,�t���ɦW)!" & vbCr & vbCr & _
        "�w�]�N�H��word����ɦW + ""���W.XLSX""�r��,�s��ୱ�W", "���W�լd", xlsp & Replace(d.Name, ".doc", "") & "���W" & StrConv(Time, vbWide) & ".XLSX")
'If xlsp = "" Then Exit Sub
xlsp = ���o�ୱ���| & "\" & Replace(d.Name, ".doc", "") & "_���W" & StrConv(Time, vbWide) & ".XLSX"
If ln = "" Then ln = 1
ln = InputBox("�Ы��w���J����" & vbCr & vbCr & "�ɮ׷|�s�b�ୱ�W�W��:" & vbCr & vbCr & Replace(d.Name, ".doc", "") & "_���W" & StrConv(Time, vbWide) & ".XLSX" & _
                vbCr & vbCr & "���ɮ�", , ln + 1)
If ln = "" Then Exit Sub
If Not IsNumeric(ln) Then Exit Sub
If ln > 11 Or ln < 2 Then Exit Sub


ds = VBA.Timer

With d
    For Each char In d.Characters
        Select Case ln
            Case 2
                charText = char & char.Next
            Case 3
                charText = char & char.Next & char.Next.Next
            Case 4
                charText = char & char.Next & char.Next.Next & char.Next.Next.Next
            Case 5
                charText = char & char.Next & char.Next.Next & char.Next.Next.Next & char.Next.Next.Next.Next
            Case 6
                charText = char & char.Next & char.Next.Next & char.Next.Next.Next & char.Next.Next.Next.Next & char.Next.Next.Next.Next.Next
            Case 7
                charText = char & char.Next & char.Next.Next & char.Next.Next.Next & char.Next.Next.Next.Next & char.Next.Next.Next.Next.Next & char.Next.Next.Next.Next.Next.Next
            Case 8
                charText = char & char.Next & char.Next.Next & char.Next.Next.Next & char.Next.Next.Next.Next & char.Next.Next.Next.Next.Next & char.Next.Next.Next.Next.Next.Next & char.Next.Next.Next.Next.Next.Next.Next
            Case 9
                charText = char & char.Next & char.Next.Next & char.Next.Next.Next & char.Next.Next.Next.Next & char.Next.Next.Next.Next.Next & char.Next.Next.Next.Next.Next.Next & char.Next.Next.Next.Next.Next.Next.Next & char.Next.Next.Next.Next.Next.Next.Next.Next
            Case 10
                charText = char & char.Next & char.Next.Next & char.Next.Next.Next & char.Next.Next.Next.Next & char.Next.Next.Next.Next.Next & char.Next.Next.Next.Next.Next.Next & char.Next.Next.Next.Next.Next.Next.Next & char.Next.Next.Next.Next.Next.Next.Next.Next & char.Next.Next.Next.Next.Next.Next.Next.Next.Next
            Case 11
                charText = char & char.Next & char.Next.Next & char.Next.Next.Next & char.Next.Next.Next.Next & char.Next.Next.Next.Next.Next & char.Next.Next.Next.Next.Next.Next & char.Next.Next.Next.Next.Next.Next.Next & char.Next.Next.Next.Next.Next.Next.Next.Next & char.Next.Next.Next.Next.Next.Next.Next.Next.Next & char.Next.Next.Next.Next.Next.Next.Next.Next.Next.Next
        End Select
        If Not charText Like "*[-'�@ �C�A�B�F�G�H:,;,�q�r�m�n ''�u�v�y�z�]�^�����H�I�]�^�i�j�X""()<>" _
            & ChrW(9312) & Chr(-24153) & Chr(-24152) & ChrW(8218) & Chr(13) & Chr(10) & Chr(11) & ChrW(12) & Chr(63) & Chr(9) & Chr(-24152) & Chr(-24153) & "�����j�i~/�_�X]*" _
            And Not charText Like "*[a-zA-Z0-9��-��]*" And InStr(charText, ChrW(-243)) = 0 And InStr(charText, Chr(91)) = 0 And InStr(charText, Chr(93)) = 0 Then
            'If Not charText Like "[a-z1-9]" & Chr(-24153) & Chr(-24152) & " �@�B'""�u�v�y�z�]�^�СH�I]" Then
'            If InStr(Chr(-24153) & Chr(-24152) & Chr(2) & "�E[]�e�f�����K�F,�A.�C�D �@�B'""����`\{}�a�b�u�v�y�z�]�^�m�n�q�r�СH�I]", charText) = 0 Then
            If Not charText Like "*[" & ChrW(-24153) & ChrW(-24152) & Chr(2) & "�E[]�e�f�����K�F,�A.�C�D �@�B'""����`\{}�a�b�u�v�y�z�]�^�m�n�q�r�СH�I���a�b]*" Then
            'chr(2)�i��O���}�аO
                If preChar <> charText Then
                    'If UBound(X) > 0 Then
                        If preChar = "" Then 'If IsEmpty(X) Then'�p�G�O�@�}�l
                            GoTo 1
                        ElseIf UBound(Filter(x, charText)) Then ' <> charText Then  '�p�G�|�L���r
1                           ReDim Preserve x(i)
                            ReDim Preserve xT(i)
                            x(i) = charText
                            xT(i) = xT(i) + 1
                            i = i + 1
                        Else
                            GoSub ���W�[�@
                        End If
                    'End If
                Else
                    GoSub ���W�[�@
                End If
                preChar = charText
            End If
        End If
    Next
End With
12
Dim doc As New Document, Xsort() As String, u As Long ', xTsort() As Integer, k As Long, so As Long, ww As String
If u = 0 Then u = 1 '�Y�L����u���W�[�@:�v�Ƶ{��,�Y�L�W�L1�������W�A�h�@Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "�B" & x(j - 1) & _
                                �|�X���G�}�C���޶W�X�d�� 2015/11/5

ReDim Xsort(u) As String
Set xlApp = CreateObject("Excel.Application")
Set xlBook = xlApp.workbooks.Add
Set xlSheet = xlBook.Worksheets(1)
With xlSheet.Application
    For j = 1 To i
        .Cells(j, 1) = x(j - 1)
        .Cells(j, 2) = xT(j - 1)
        Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "�B" & x(j - 1) 'Xsort(xT(j - 1)) & ww '�}�C�Ƨ�'2010/10/29
    Next j
End With
doc.ActiveWindow.Visible = False
If d.ActiveWindow.View.ReadingLayout Then ReadingLayoutB = True: d.ActiveWindow.View.ReadingLayout = Not d.ActiveWindow.View.ReadingLayout
'U = UBound(Xsort)
For j = u To 0 Step -1 '�}�C�Ƨ�'2010/10/29
    If Xsort(j) <> "" Then
        With doc
            If Len(.Range) = 1 Then '�|����J���e
                .Range.InsertAfter "���W = " & j & "���G�]" & Len(Replace(Xsort(j), "�B", "")) / ln & "�ӡ^"
                .Range.Paragraphs(1).Range.Font.Size = 12
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "�s�ө���"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
                '.Range.Paragraphs(1).Range.Font.Bold = True
            Else
                .Range.InsertParagraphAfter
                .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
                .Range.InsertAfter "���W = " & j & "���G�]" & Len(Replace(Xsort(j), "�B", "")) / ln & "�ӡ^"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Size = 12
                '.Range.Paragraphs(.Paragraphs.Count).Range.Bold = True
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "�s�ө���"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
            End If
            .Range.InsertParagraphAfter
            .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
            .Range.Paragraphs(.Paragraphs.Count).Range.Font.Size = 12
'            .Range.Paragraphs(.Paragraphs.Count).Range.Bold = False
            .Range.InsertAfter Replace(Xsort(j), "�B", Chr(9), 1, 1) 'chr(9)���w��r��(Tab���)
            .Range.InsertParagraphAfter
            If InStr(.Range.Paragraphs(.Paragraphs.Count).Range, "���W") = 0 Then
                .Range.Paragraphs(.Paragraphs.Count - 1).Range.Font.Name = "�з���"
            Else
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.Name = "�s�ө���"
                .Range.Paragraphs(.Paragraphs.Count).Range.Font.NameAscii = "Times New Roman"
            End If
        End With
    End If
Next j

With doc.Paragraphs(1).Range
     .InsertParagraphBefore
     .Font.NameAscii = "times new roman"
    doc.Paragraphs(1).Range.InsertParagraphAfter
    doc.Paragraphs(1).Range.InsertParagraphAfter
    doc.Paragraphs(1).Range.InsertAfter "�A���Ѫ��奻�@�ϥΤF" & i & "�Ӥ��P�����J�]�ǲΦr�P²�Ʀr�����X�֡^"
End With

doc.ActiveWindow.Visible = True

de = VBA.Timer
doc.SaveAs Replace(xlsp, "XLS", "doc") '���j�p�g
If ReadingLayoutB Then d.ActiveWindow.View.ReadingLayout = Not d.ActiveWindow.View.ReadingLayout
Set d = Nothing ' ActiveDocument.Close wdDoNotSaveChanges

Debug.Print Now

MsgBox "�����I" & vbCr & vbCr & "�O��" & left(de - ds, 5) & "��!", vbInformation
xlSheet.Application.Visible = True
xlSheet.Application.UserControl = True
xlSheet.SaveAs xlsp
Exit Sub
���W�[�@:
For j = 0 To UBound(x)
    If x(j) = charText Then
        xT(j) = xT(j) + 1
        If u < xT(j) Then u = xT(j) '�O�U�̰����W,�H�K�Ƨ�(�N���ƧǤ��}�C�̰������ȳ]����,�h���|�W�X�}�C.
        '�h���@��]���n���ƧP�_�p��n�X��,�G�įण�W�ϴ�''�į��٬O�t���h��.
        Exit For
    End If
Next j

Return
ErrH:
Select Case Err.Number
    Case 4605 '�\Ū�Ҧ�����s��'����k���ݩʵL�k�ϥΡA�]�����R�O�L�k�b�\Ū���ϥΡC
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
        doc.ActiveWindow.View.ReadingLayout = False
        doc.ActiveWindow.Visible = False
        ReadingLayoutB = True
        Resume
    
    Case 91, 5941 '�S���]�w�����ܼƩ� With �϶��ܼ�,���X���һݪ��������s�b
        GoTo 12
    Case Else
        MsgBox Err.Number & Err.Description, vbCritical 'STOP: Resume
        Resume
        End
    
End Select
End Sub


Sub �ѦW���g�W���ˬd()
Dim s As Long, rng As Range, e, trm As String, ans
Static x() As String, i As Integer
On Error GoTo eH
Do
    Selection.Find.Execute "�q", , , , , , True, wdFindAsk
    Set rng = Selection.Range
    rng.MoveEndUntil "�r"
    trm = Mid(rng, 2)
    
    For Each e In x()
        If StrComp(e, trm) = 0 Then GoTo 1
    Next e
2   ans = MsgBox("�O�_���L�u" & trm & "�v�H" & vbCr & vbCr & vbCr & "�����Ы� NO[�_]", vbExclamation + vbYesNoCancel)
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
    Case 92 '�S���]�w For �j�骺��l�� �}�C�|������
        GoTo 2
End Select
End Sub

Sub �ɶ��b����ഫ() '2017/5/13 �]��YOUKU�PYOUTUBE�ɶ��b��줣�P�ӳ]
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
                    If InStr(Trim(myRng), " ") Then '�p�G��2�ӥH�W�ɶ��b
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
                    Else '�p�G�u��1�Ӯɶ��b
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
                            aM.Next.Next.text = str((CInt(aM.Next.Next) * 10 + CInt(aM) * 60) / 10)
                            aM.Next.Delete
                            aM.Delete
                            Exit For
                        End If
                    Next
Return
End Sub
Sub ������Ǯѹq�l�ƭp��_������r(ByRef r As Range)
On Error GoTo eH
Dim lngTemp As Long '�]���~����l�ܭ׭q�A�~�|�޵o�T�����ܧR���x�s�椣�|������
'Dim d As Document
Dim tb As Table, C As Cell ', ci As Long
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
    Case 5992 '�L�k�ӧO�s�������X�����U��A�]����椤���V�X���x�s��e�סC
        For Each C In tb.Range.Cells
'            ci = ci + 1
'            If ci Mod 3 = 2 Then
                'If VBA.IsNumeric(VBA.Left(c.Range.text, VBA.InStr(c.Range.text, "?") - 1)) Then
                If VBA.InStr(C.Range.text, ChrW(160) & ChrW(47)) > 0 Then
'                    word.Application.DisplayAlerts = False
                    C.Delete  '�R���s�����x�s��
                End If
'            End If
        Next C
        Resume Next
    Case Else
        MsgBox Err.Number & Err.Description
        End
End Select
End Sub

Sub ������Ǯѹq�l�ƭp��_�����ܤp����^�j()
Dim slRng As Range, a
Set slRng = Selection.Range
������Ǯѹq�l�ƭp��_������r slRng
For Each a In slRng.Characters
    Select Case a.Font.Color
        Case 34816, 8912896
            a.Font.Size = 14
        Case 0
            a.Font.Size = 30
    End Select
Next a
End Sub
Sub ������Ǯѹq�l�ƭp��_�h������O�d����()
Dim slRng As Range, a, ur As UndoRecord
'Set ur = SystemSetup.stopUndo("������Ǯѹq�l�ƭp��_�h������O�d����")
SystemSetup.stopUndo ur, "������Ǯѹq�l�ƭp��_�h������O�d����"
Docs.�ťժ��s���
If ActiveDocument.Characters.Count = 1 Then Selection.Paste
If Selection.Type = wdSelectionIP Then ActiveDocument.Select
Set slRng = Selection.Range
������Ǯѹq�l�ƭp��_������r slRng
For Each a In slRng.Characters
    Select Case a.Font.Color
        Case 34816, 8912896
            If a.Font.Size <> 12 Then Stop
            a.Delete
        Case 254
            If a.Font.Size = 9 Then a.Delete
    End Select
Next a
If MsgBox("�O�_���N����r�H", vbOKCancel) = vbOK Then ��r�ഫ.����r�ॿ
Beep 'MsgBox "done!", vbInformation
SystemSetup.contiUndo ur
End Sub
Sub ������Ǯѹq�l�ƭp��_����e��[�A��()
Dim slRng As Range, a, flg As Boolean, ur As UndoRecord 'Alt+1
'Set ur = SystemSetup.stopUndo("������Ǯѹq�l�ƭp��_����e��[�A��")
SystemSetup.stopUndo ur, "������Ǯѹq�l�ƭp��_����e��[�A��"
Docs.�ťժ��s���
If Selection.Type = wdSelectionIP Then ActiveDocument.Select
Set slRng = Selection.Range
������Ǯѹq�l�ƭp��_������r slRng
For Each a In slRng.Document.Paragraphs 'for�~�y�q�l���m��Ʈw
    If VBA.left(a.Range, 3) = "[��]" Then
        slRng.SetRange a.Range.Characters(4).start _
            , a.Range.End
        slRng.Font.Size = 7.5
    End If
Next a
If Selection.Type = wdSelectionIP Then ActiveDocument.Select
Set slRng = Selection.Range
For Each a In slRng.Characters
    Select Case a.Font.Color
        Case 34816, 8912896, 15776152 '34816:���p�`
p:          If flg = False Then
                a.Select
                Selection.Range.InsertBefore "�]"
                Selection.Range.SetRange Selection.start, Selection.start + 1
                Selection.Range.Font.Size = a.Characters(2).Font.Size
                Selection.Range.Font.Color = a.Characters(2).Font.Color
'                a.Font.Size = a.Next.Font.Size
'                a.Font.Color = a.Next.Font.Color
                flg = True
            Else
                If a.Font.Color = 8912896 And a.Previous.Font.Color = 34816 Then '8912896�Ŧr�p�`
                    a.InsertBefore "�^�]"
                    a.SetRange a.start, a.start + 2
                    a.Font.Size = a.Characters(2).Next.Font.Size
                    a.Font.Color = a.Characters(2).Next.Font.Color
'                    a.Characters(1).Font.Color = a.Characters(1).Previous.Font.Color
                End If
            End If
'        Case 8912896 '8912896�Ŧr�p�`
            
        Case 0, 15595002, 15649962
            If a.Font.Color = 0 Then 'black'�~�y�q�l���m��Ʈw
                If a.Font.Size = 7.5 And Not flg Then
                    GoTo p
                ElseIf a.Font.Size > 7.5 And flg Then
                    GoTo b
                End If
            'End If
            ElseIf flg Then
b:
'                a.Select
'                Selection.Range.InsertBefore "�^"
                If a.Previous = Chr(13) Then
                    a.Previous.Previous.Select
                Else
                    a.Previous.Select
                End If
                Selection.Range.InsertAfter "�^"
                flg = False
            End If
        Case -16777216 'black'�~�y�q�l���m��Ʈw
            If a.Font.Size = 7.5 And Not flg Then
                GoTo p
            ElseIf a.Font.Size > 7.5 And flg Then
                GoTo b
            End If
        Case 255 'red'�~�y�q�l���m��Ʈw
            Select Case a.Font.Size
                Case 7.5, 10
                    a.Delete
            End Select
    End Select
Next a
slRng.Find.Execute "�]�]", True, , , , , , , , "�]", wdReplaceAll
slRng.Find.Execute "�^�^", True, , , , , , , , "�^", wdReplaceAll
Beep
Selection.EndKey wdStory
Do
   Selection.MoveLeft
   If Selection = Chr(13) Then Selection.Delete
Loop While Selection = Chr(13)
'MsgBox "done!", vbInformation
SystemSetup.contiUndo ur
End Sub
Sub �~�y�q�l���m��Ʈw�奻��z_�H��K�줤����Ǯѹq�l�ƭp��(Optional doNotCloseDoc As Boolean)
Dim rng As Range, d As Document, a, ur As UndoRecord
Dim rp As Variant, i As Byte
'Set ur = SystemSetup.stopUndo("�~�y�q�l���m��Ʈw�奻��z_�H��K�줤����Ǯѹq�l�ƭp��")
SystemSetup.stopUndo ur, "�~�y�q�l���m��Ʈw�奻��z_�H��K�줤����Ǯѹq�l�ƭp��"
If Documents.Count = 0 Then Documents.Add
Set d = ActiveDocument
If d.path <> "" Or d.Content.text <> Chr(13) Then
    Set d = Documents.Add()
    'Exit Sub
End If
rp = Array("(", "{{", ")", "}}", ChrW(160), "", "�i�ϡj", "", _
     "^p^p", "^p", _
     ChrW(13) & ChrW(45) & ChrW(13) & ChrW(13) & ChrW(11), "^p", _
     ChrW(13) & ChrW(45) & ChrW(13), "^p", "{{ }}", "", "[", ChrW(12310), _
     "]", ChrW(12311), " ", "", "��", ChrW(12295), _
     "^p" & ChrW(12310) & "��" & ChrW(12311), ChrW(12310) & "��" & ChrW(12311) & "{{", _
     "}}" & Chr(13) & "^#" & Chr(13) & "{{", "", _
     "�D�D�D�D�D�D�D�D�D�D�D�D�D�D�D�D�D�D" & Chr(13), "", _
     Chr(13) & "^#" & Chr(13), "", _
     "}}" & Chr(13) & "^#" & Chr(13), "}}", _
     "}}" & Chr(13) & "{{", "", _
     "-", "", "^#", "", "�C�C", "�C") ', "�C}}<p>�C}}<p>", "�C}}<p>")
     '��ӡuChrW(13) & ChrW(45) & ChrW(13) & ChrW(13) & ChrW(11)�v�O�䤤������
Set rng = d.Range
rng.Paste
�~�y�q�l���m��Ʈw�奻��z_�`��e��[�A��
For Each a In rng.Characters
    If a.Font.Size = 10 Then
        Select Case a.Font.Color
            Case 255, 9915136
                a.Delete
        End Select
    End If
Next a
rng.Cut
On Error GoTo eH:
rng.PasteAndFormat wdFormatPlainText
rng.Find.ClearFormatting
For i = 0 To UBound(rp)
    If InStr(rng.text, rp(i)) > 0 Then
        rng.Find.Execute rp(i), , , , , , , wdFindContinue, , rp(i + 1), wdReplaceAll
    End If
    i = i + 1
Next i
������Ǯѹq�l�ƭp��.�����w���������⴫���r d
��r�B�z.�ѦW���g�W���Ъ`
Beep
If Not doNotCloseDoc Then
    d.Range.Cut
    d.Close wdDoNotSaveChanges
End If
SystemSetup.contiUndo ur
Exit Sub
eH:
Select Case Err.Number
    Case 4198 '���O����
        SystemSetup.wait 900
        Resume
    Case Else
        MsgBox Err.Number + Err.Description
End Select
End Sub
Sub �~�y�q�l���m��Ʈw�奻��z_�`��e��[�A��()
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
                (rng.Font.Color <> 234 And rng.Font.Bold = False) Then '���r+���鬰�˯����G
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
Sub �֥y����()
Dim slRng As Range, a
Set slRng = Selection.Range
For Each a In slRng.Characters
    If a Like "[�C�A�F�H�I�u�v�y�z]" Then
        a.Select
        Selection.move
        Selection.TypeText Chr(11)
    End If
Next a
End Sub

Sub �R���ծ׻y()
Dim rng As Range, e, d As Document
Set d = ActiveDocument
Set rng = d.Range
e = rng.End
With rng.Find
    .Style = "�W�s��"
    .Execute , , , , , , , wdFindStop ', , "" ', wdReplaceAll
    Do
        If InStr(rng.Characters(rng.Characters.Count).Next.Style, "�ծ�") _
            Or InStr(rng.Characters(1).Previous.Style, "�ծ�") Then
            rng.Select
            Selection.Delete
            rng.SetRange Selection.start, e
        End If
    Loop While .Execute(, , , , , , , wdFindStop)  ', , "" ', wdReplaceAll
End With

With rng.Find
    .Style = "�ծ�"
    .Execute , , , , , , , wdFindContinue, , "", wdReplaceAll
End With
With rng.Find
    .Style = "�ծפޤ�"
    .Execute , , , , , , , wdFindContinue, , "", wdReplaceAll
End With
Beep
End Sub

Function ��y���`����r�B�z(x As String)
Dim ay, i As Byte
ay = Array("��", ChrW(20008), "�@", " ", "�]�S���^", "�S�� ", "�]Ū���^", "Ū�� ", "�]�y���^", "�y�� ", _
        "(�@)", "", "(�G)", "", "(�T)", "", "(�|)", "", "(��)", "", "(��)", "", "�^", "", "�]", "")
For i = 0 To UBound(ay)
    x = Replace(x, ay(i), ay(i + 1))
    i = i + 1
Next i
��y���`����r�B�z = x
End Function
Sub �����r�[�W��y���`��()
Dim rng As Range, x, rst As New ADODB.Recordset, st As WdSelectionType, words As String
Dim cnt As New ADODB.Connection, id As Long, sty As word.Style, url As String
Dim frmDict As New Form_DictsURL, lnks As New Links, db As New dBase ', frm As New MSForms.DataObject
Static cntStr As String, chromePath As String
st = Selection.Type
If st = wdSelectionIP Then
    If Selection.start = 0 Then Exit Sub
    x = Selection.Previous.Characters(Selection.Previous.Characters.Count).text
    If InStr("�C�A�F�u�v�y�z�q�r�m�n�H.,;""?��-�w�w--�]�^()�i�j�e�f<>[]�K! �@�I", x) Then Exit Sub
'    Selection.Previous.Copy
Else
    x = trimStrForSearch(VBA.CStr(Selection.text), Selection)
    'Selection.Copy
    SystemSetup.ClipboardPutIn "=" & Selection.text
End If
    If ��r�B�z.isSymbol(CStr(x)) Or ��r�B�z.is�`���Ÿ�(CStr(x)) Or ��r�B�z.isLetter(CStr(x)) Or ��r�B�z.isNum(CStr(x)) Then Exit Sub
Set rng = Selection.Range
words = x
db.setWordControlValue (words)
On Error GoTo eH
Dim ur As UndoRecord
'Set ur = SystemSetup.stopUndo("�����r�[�W��y���`��")
SystemSetup.stopUndo ur, "�����r�[�W��y���`��"

'If Not Selection.Document.path = "" Then If Not Selection.Document.Saved Then Selection.Document.save
If cntStr = "" Then
    Dim dbp As New Paths
    cntStr = dbp.getdb_���s��y���׭q��_��ƮwfullName
End If

If chromePath = "" Then
    chromePath = SystemSetup.getChrome
End If

'Dim ay, i As Byte
'ay = Array("��", ChrW(20008), "�@", " ", "�]�S���^", "�S�� ", "�]Ū���^", "Ū�� ", "�]�y���^", "�y�� ", _
'        "(�@)", "", "(�G)", "", "(�T)", "", "(�|)", "", "(��)", "", "(��)", "", "�^", "", "�]", "")

    cnt.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & cntStr
'Exit Sub
'cntt:
    rst.Open "select �`���@��,���q,url,ID,�h���Ƨ� from [�m���s��y���׭q���n �`��] where strcomp(�r���W,""" & x & """)=0 order by �h���Ƨ�", cnt, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        GoSub list
    Else
        �����r�[�W��y���`��nextTable rst, cnt, x, "�m���s��y���׭q���n �`��-20210928�H�e", True
        If rst.RecordCount > 0 Then
            GoSub list
        Else
2
            If Selection.Characters.Count = 1 Then 'words  ��r
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
            Else 'terms ���J
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
                    x = frmDict.add1URLTo1����r�r��(words)
                    If x = "" Then GoTo endS
                    GoTo 2
                Else
                    rst.Close
                    rst.Open "select �`���@��,���q,url,ID,�h���Ƨ� from [�m���s��y���׭q���n �`��] where instr(�r���W,""" & x & """)>0 order by �h���Ƨ�", cnt, adOpenKeyset, adLockOptimistic
                    Selection.Copy
                    If rst.RecordCount > 0 Then
                        Beep
                        'Selection.Document.FollowHyperlink "https://www.zdic.net/hans/" & x, , True
                        Shell chromePath & " https://www.zdic.net/hans/" & x
                        GoSub list
                        'Selection.Document.FollowHyperlink "http://dict.revised.moe.edu.tw/cbdic/search.htm", , True
'                        Shell chromePath & " http://dict.revised.moe.edu.tw/cbdic/search.htm"
                    Else
                            �����r�[�W��y���`��nextTable rst, cnt, x, "�m���s��y���׭q���n �`��-20210928�H�e", False
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
'        Set ur = SystemSetup.stopUndo("�K�a��")
'        Docs.�˦�add_�K�a�����˦�
        rng.Collapse wdCollapseEnd
        If rng.Style <> "�K�a��" Then
            rng.InsertAfter "�]�^"
            rng.Style = "�K�a��"
            rng.SetRange rng.End - 1, rng.End - 1
        End If
        Do Until rst.EOF
            x = ""
            If VBA.IsNull(rst.Fields(0).Value) Then
                x = rst.Fields(1).Value '���q
            Else
                x = rst.Fields(0).Value '�`��
            End If
            GoSub typeTexts
            rst.MoveNext
        Loop
        If rng.Previous = "�A" Then rng.Previous.Delete
'        SystemSetup.contiUndo ur
'        Set ur = Nothing:  'Set frm = Nothing: Set frmDict = Nothing

Return

typeTexts:
        If x = "" Or VBA.IsNull(x) Then GoTo 2
'        X = Mid(X, 1, Len(X) - 1)
        x = ��y���`����r�B�z(CStr(x))
'        If sT <> wdSelectionIP Then
'            rng.SetRange Selection.End, Selection.End
'        End If
'        rng.SetRange rng.End - 1, rng.End - 1
        rng.InsertAfter x 'insert ZhuYin
        For Each x In rng.Characters 'format ZhuYin
            If InStr("������", x) Then
                x.Style = "�n��"
            ElseIf InStr("��", x) Then
                x.Font.Name = "�з���"
            End If
        Next x
        x = rst.Fields(2).Value 'URL  'frmDict.get1URLfor1(words)
        If VBA.IsNull(x) Then
                If st = wdSelectionIP Then
                    If Selection.Previous.Characters(Selection.Previous.Characters.Count).Hyperlinks.Count > 0 Then
                        Dim rngW As Range
                        Set rngW = Selection.Range
                        rngW.SetRange Selection.Previous.Characters(Selection.Previous.Characters.Count).start, Selection.Previous.Characters(Selection.Previous.Characters.Count).End
                        SystemSetup.ClipboardPutIn "=" & rngW.text '"^" & rngW.text & "$" 'version 6's new settings
                        Set rngW = Nothing
                    Else
                        Set rngW = Selection.Previous.Characters(Selection.Previous.Characters.Count)
                        SystemSetup.ClipboardPutIn "=" & rngW.text
                        'Selection.Previous.Characters(Selection.Previous.Characters.Count).Copy
                    End If
                End If
'                Shell chromePath & " http://dict.revised.moe.edu.tw/cbdic/search.htm"
'            frm.Clear
'            frm.SetText words, 1
'            frm.PutInClipboard
            'add new url
            Dim repeated As Boolean '�˯����G����@�Ӯɷ|����
rePt:
            If repeated = False Then x = SeleniumOP.grabDictRevisedUrl_OnlyOneResult(words) '���k�u�A�Ω��1����Ʈ�,�S���Φh��1���h��^""�Ŧr��
            rng.Document.ActiveWindow.Application.Activate
            If rst.RecordCount = 1 Then '��y����Ʈw�̥u��1���k�X���
                If x = "" Then '���G����1�Ӯ�
                    Shell Network.getDefaultBrowserFullname & " https://dict.revised.moe.edu.tw/search.jsp?md=1"
                End If
''                If Not SystemSetup.appActivatedYet("chrome") Then
''                'If Not word.Tasks.Exists("google chrome") Then
''                    Shell SystemSetup.getChrome & " https://dict.revised.moe.edu.tw/search.jsp?md=1"
''                Else
''                    SystemSetup.appActivateChrome
''                End If
            Else
                If VBA.IsNull(x) Then x = ""
                Beep
            End If
            If x = "" Then '���G����1�Ӯ�
                If repeated = False Then
                    If SeleniumOP.ActiveXComponentsCanNotBeCreated Then
                        SystemSetup.playSound 2
                        Shell Network.getDefaultBrowserFullname & " https://dict.revised.moe.edu.tw/search.jsp?md=1&word=" & words
                    Else
                        Shell Network.getDefaultBrowserFullname & " https://dict.revised.moe.edu.tw/search.jsp?md=1"
                    End If
                Else
                
                End If
                x = InputBox("plz putin the url", , IIf(VBA.IsNull(rst.Fields(0).Value), "", rst.Fields(0).Value)) 'frmDict.add1URLTo1��y���(words)
                If repeated Then
                    SystemSetup.wait 1 '���T�w�n��J���ӵ����A�A�N�s�����m�e
                    appActivateChrome
                End If
                repeated = True
            End If
            If x = "" Then GoTo endS
            If left(x, 4) <> "http" Then GoTo rePt
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
                '�H�U�����h
                '�b�d�rforInPut��Ʈw����椤�]�w�������
                'db.setURLControlValue VBA.CStr(x)
            Else
                GoTo endS 'Exit Sub
            End If
        End If
        Set sty = rng.Style
        rng.Hyperlinks.Add rng, lnks.trimLinks(VBA.CStr(x)), , , , "_blank"
        lnks.setStylewithHyperlinkMark sty, rng
        rng.Collapse wdCollapseEnd
        'rng.SetRange rng.End, rng.End
        rng.Next.InsertBefore "�A"
'        rng.Style = "�K�a��"
        'rng.Hyperlinks.Item(1).Delete
        'rng.Collapse wdCollapseEnd
        rng.SetRange rng.End + 2, rng.End + 2
Return


eH:
    Select Case Err.Number
        Case 4198 '���O���� 'Google Drive�����D
            Resume Next
        Case 5834 '���w�W�٪����ؤ��s�b
            Docs.�˦�add_�K�a�����˦�
            Resume
        Case 5 '�{�ǩI�s�Τ޼Ƥ����T
            SystemSetup.wait 3 'http://vbcity.com/forums/t/81315.aspx
            'Application.Wait (Now + TimeValue("0:00:10")) '<~~ Waits ten seconds.
            Resume 'https://stackoverflow.com/questions/21937053/appactivate-to-return-to-excel
        Case Else
            MsgBox Err.Number & Err.Description
'            Resume
            GoTo endS
            'If cnt.State <> adStateClosed Then cnt.Close
    End Select
End Sub

Sub �����r�[�W��y���`��nextTable(ByRef rst As ADODB.Recordset, ByRef cnt As ADODB.Connection, x, tbName As String, precise As Boolean)
    If rst.State = adStateOpen Then rst.Close
    Dim src As String
    Dim srcs As String
    srcs = "select �`���@��,���q,url,ID from [" & tbName & "] where "
    If precise Then
        src = "strcomp(�r���W,""" & x & """)=0"
    Else
        src = "instr(�r���W,""" & x & """)>0"
    End If
    rst.Open srcs & src, cnt, adOpenKeyset
End Sub

Rem �ŧi���ѡG�b�udx = regEx.Replace(dx, rw)�v�o��|�X�{�G 5017�G���ε{���Ϊ���w�q�W�����~
Rem 20230309 creedit with chatGPT�j���ġG�ѦW�����I�P���h��F��ADO.NET�BLINQ�G
Sub �ѦW���g�W���Ъ`_���h��F��RegularExpression_Plaintext()
Dim cnt As New ADODB.Connection, rst As New ADODB.Recordset
Dim cntStr As String, d As Document, dx As String, w As String, rw As String
Dim db As New dBase
db.cnt�d�r cnt
Dim regEx As Object
'Dim regEx As New RegExp
    Set regEx = CreateObject("VBScript.RegExp")
Dim replacedText As String
Set d = ActiveDocument: dx = d.Range.text
rst.Open "select * from ���I�Ÿ�_�ѦW��_�۰ʥ[�W�� order by �Ƨ�", cnt, adOpenForwardOnly, adLockReadOnly
Do Until rst.EOF
    w = rst("�ѦW").Value
    If VBA.InStr(dx, w) Then 'if found
        If VBA.IsNull(rst("���N��").Value) Then
            rw = "�m" & rst("�ѦW").Value & "�n"
        Else
            rw = rst("���N��").Value
        End If
        With regEx
            '.Pattern = "(?<!�m)(?<!�q)(?<![\\p{P}&&[^�n�r]]+)" + regEx.Escape(w) + "(?!�n)(?!�r)"
            .Pattern = "(?<!�m)(?<!�q)(?<![\\p{P}&&[^�n�r]]+)" + Replace(Replace(w, "\", "\\"), ".", "\.") + "(?!�n)(?!�r)"
            Rem �b Word VBA ���ARegExp ���� Escape ��k�O���Q�䴩���C�ҥH�z�ݭn��o�Ӥ�k�令�ϥ� Replace ��k�N�S��r���ഫ�����h��F������q�r�šC
            Rem �o�ӡuregEx.Escape(w)�v�쩳�O����N�q�H
            Rem regEx.Escape(w) �O�N�r�� w ���Ҧ������h��F�����r�� (�Ҧp *, ?, [, ], \, (, ), {, }, +, ^, $, ., |) �ഫ���¤�r�A�H�קK�o�Ǥ��r�ųQ��@���h��F���������ӥX�{���~�C
            Rem �Ҧp�A�p�G w �� test*�A�h regEx.Escape(w) �|�^�� test\*�A�o�˥��h��F�������N�|�� * �����@��r���A�Ӥ��O���h��F�������q��
            .Global = True
        End With
        dx = regEx.Replace(dx, rw)
    End If
    rst.MoveNext
Loop
Documents.Add.Range.text = dx
rst.Close
'rst.Open "select * from ���I�Ÿ�_�g�W��_�۰ʥ[�W�� order by �Ƨ�", cnt, adOpenForwardOnly, adLockReadOnly
rst.Close: cnt.Close
Exit Sub
Rem �o�˪��g�k�]�i�H���ΦbWord VBA�� �p���@
Rem �O���A�o�˪����h��F���g�k�]�i�H�b Word VBA �����ΡC�H�U�O�@��²�檺�ܨ� VBA �{�ǡA�i�H�N�襤�奻�����ѦW���M�g�W�����������媺���μ��I�Ÿ��G
Rem
Rem vba
Rem Copy code
Rem Sub ReplacePunctuation()
Rem
Rem     Dim regEx As Object
Rem     Set regEx = CreateObject("VBScript.RegExp")
Rem
Rem     With regEx
Rem         .Pattern = "(?<!�m)(?<!�q)[,.;:!?()]*(�m.*?�n|�q.*?�r)[,.;:!?()]*"
Rem         .Global = True
Rem     End With
Rem
Rem     Dim selection As Range
Rem     Set selection = Selection.Range
Rem
Rem     Dim replacedText As String
Rem     replacedText = regEx.Replace(selection.Text, "�m$1�n")
Rem
Rem     selection.Text = replacedText
Rem
Rem End Sub
Rem �o�ӵ{���|�ϥλP�e�� Python �{�ǬۦP�����h��F���Ӷi������C���|���襤�@�q��r�A�M��N�䤤���ѦW���M�g�W�����������媺���μ��I�Ÿ��C�{�����ϥΤF Word VBA ���ت� Range �M Selection ��H�A�H�� VBScript.RegExp ���O�� CreateObject ��k�ӳЫإ��h��F����H�C
Rem
Rem
Rem
Rem
Rem �]�u�u
Rem .Global=true �O����N��
Rem .Global=true �O���h��F����H���@���ݩʡA�Ω���w�b�奻���ǰt�Ҧ��ŦX���󪺼Ҧ��A�Ӥ��ȶȬO�Ĥ@�Ӥǰt���C�q�{���p�U�A���h��F���u�ǰt�Ĥ@�ӲŦX���󪺼Ҧ��A�M��N����ǰt�C�]���A�N .Global �ݩʳ]�m�� true �i�H�O�Ҥǰt��Ҧ����Ҧ��C
Rem
Rem �b�W�����Ҥl���Aregex ��H�Q�ЫخɡA.Global �ݩʳQ�]�m�� true�A�o�˴N�i�H�ǰt�奻���Ҧ������I�Ÿ��F�C�b Word VBA ���A�z�i�H�q�L�]�m Match.Global �ݩʨӹ�{�������\��C�Ҧp�G
Rem
Rem vbnet
Rem Copy code
Rem Sub ReplacePunctuation()
Rem
Rem     Dim regEx As New RegExp
Rem     regEx.Pattern = "\\p{P}"
Rem     regEx.Global = True
Rem
Rem     ActiveDocument.Content.Select
Rem     Selection.Find.ClearFormatting
Rem     Selection.Find.Replacement.ClearFormatting
Rem     With Selection.Find
Rem         .Text = regEx.Pattern
Rem         .Replacement.Text = "�A"
Rem         .Forward = True
Rem         .Wrap = wdFindContinue
Rem         .Format = False
Rem         .MatchCase = False
Rem         .MatchWholeWord = False
Rem         .MatchAllWordForms = False
Rem         .MatchSoundsLike = False
Rem         .MatchWildcards = True
Rem         .Execute Replace:=wdReplaceAll
Rem     End With
Rem
Rem End Sub
Rem �b�o�ӨҤl���A�ڭ̳ЫؤF�@�� RegExp ��H regEx�A�ñN�� .Pattern �ݩʳ]�m�� \\p{P}�A�N .Global �ݩʳ]�m�� True�C�M��ڭ̿襤���ɤ����Ҧ����e�A����d��M�����ާ@�A�N�Ҧ��ŦX�Ҧ������I�Ÿ�����������r���C.MatchWildcards �ݩʳQ�]�m�� True�A�o�˦b�d��ɴN�i�H�ϥΥ��h��F���Ҧ��C
Rem
Rem
Rem
Rem
Rem �]�u�u
Rem �ҥH Find ���� .Text �ݩʬO�i�H�]�����h��F���o�H ���n�N .MatchWildcards �]�� true �N�O�F�H �_�h�N�|�Q��@�@���r�d��A�Ӥ��O�H���h��F�����W�h����o�H �O�o�˶ܡH�Ѯv����
Rem �O���A�z���o�D�`���T�C�b Word VBA ���A�n�N Find ���� .Text �ݩʳ]�����h��F���A�������N .MatchWildcards �]�� true�A�_�h Word �|�N������@���r�d��C�]���A�p�G�z�Q�n�ϥΥ��h��F���i��d��M�����ާ@�A�аO�o�N .MatchWildcards �]�� true�C
Rem
End Sub
'Function �ѦW���g�W���Ъ`PreExamOK(d As Document, term As String, whatMark As String, startPos_term As Long, Optional endPos_term As Long) As Boolean
Function �ѦW���g�W���Ъ`PreExamOK(d As Document, term As String, startPos_term As Long, Optional endPos_term As Long) As Boolean
    Dim rngChk As Range, xChk As String
    On Error GoTo eH:
    Set rngChk = d.Range(0, startPos_term)
    xChk = rngChk.text
    'If term = "��v�qŲ" Then Stop
    If InStrRev(xChk, "�m") <= InStrRev(xChk, "�n") And InStrRev(xChk, "�q") <= InStrRev(xChk, "�r") Then �ѦW���g�W���Ъ`PreExamOK = True
    
    Exit Function
eH:
        Select Case Err.Number
            'Case 4608 '�ƭȶW�X�d��
                'Resume
            Case Else
                MsgBox Err.Number + Err.Description
    '            Resume
        End Select
    
    
    'Dim result As Boolean
    'If whatMark = "�m" Then ' = �H �p�G���ɷ|�u=�v�GIf InStr(xChk, "�m") = 0 And InStr(xChk, "�n") = 0 And InStr(xChk, "�q") = 0 And InStr(xChk, "�r") = 0 Then 20230312 �u�ơC�P���P���@�g���g�ۡ@�n�L��������C�S������ĥ[���A�ڮ]�u�u�i��ܡH
    '    If InStrRev(xChk, "�m") <= InStrRev(xChk, "�n") Then result = True
    'Else
    '    If InStrRev(xChk, "�q") <= InStrRev(xChk, "�r") Then result = True
    'End If
    
    ''�e�����S�m�n�q�r��
    'If InStr(xChk, "�m") = 0 And InStr(xChk, "�n") = 0 And InStr(xChk, "�q") = 0 And InStr(xChk, "�r") = 0 Then
    '    result = True
    ''�e�����m�q�b�n�r���e��
    'Else
    '    'If InStrRev(xChk, "�m") < InStrRev(xChk, "�n") Or InStrRev(xChk, "�q") < InStrRev(xChk, "�r") Then result = True
    '    If whatMark = "�m" Then
    '        If InStr(xChk, "�m") = 0 And InStr(xChk, "�n") = 0 Then
    '            result = True
    '        Else
    '            If InStrRev(xChk, "�m") < InStrRev(xChk, "�n") Then result = True
    '        End If
    '    ElseIf whatMark = "�q" Then
    '        If InStr(xChk, "�q") = 0 And InStr(xChk, "�r") = 0 Then
    '            result = True
    '        Else
    '            If InStrRev(xChk, "�q") < InStrRev(xChk, "�r") Then result = True
    '        End If
    '    End If
    'End If
    '�ѦW���g�W���Ъ`PreExamOK = result
End Function
Sub �ѦW���g�W���Ъ`()
    Dim cnt As New ADODB.Connection, rst As New ADODB.Recordset
    Dim cntStr As String, d As Document, dx As String, rngF As Range, title As String
    Dim db As New dBase
    Dim ur As UndoRecord
    On Error GoTo eH:
    SystemSetup.stopUndo ur, "�ѦW���g�W���Ъ`"
    db.cnt�d�r cnt
    'If Dir("H:\�ڪ����ݵw��\�p�H\�d�{�@�o�N(C�Ѫ�)\���y���\�ϮѺ޲z����", vbDirectory) <> "" Then
    '    cntStr = "Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=H:\�ڪ����ݵw��\�p�H\�d�{�@�o�N(C�Ѫ�)\���y���\�ϮѺ޲z����\�d�r.mdb;"
    'ElseIf Dir("D:\�d�{�@�o�N\���y���\�ϮѺ޲z����", vbDirectory) <> "" Then
    '    cntStr = "Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=D:\�d�{�@�o�N\���y���\�ϮѺ޲z����\�d�r.mdb;"
    'Else
    '    MsgBox "���|���s�b�I", vbCritical: Exit Sub
    'End If
    Set d = ActiveDocument: dx = d.Range.text: Set rngF = d.Range
    'cnt.Open cntStr
    word.Application.ScreenUpdating = False
    
    GoSub bookmarks '���I�Ÿ�_�ѦW��_�۰ʥ[�W��
    rst.Open "select * from ���I�Ÿ�_�g�W��_�۰ʥ[�W�� order by �Ƨ�", cnt, adOpenForwardOnly, adLockReadOnly
    Set rngF = d.Range: dx = d.Range.text
    Do Until rst.EOF
        title = rst("�g�W").Value
        If VBA.InStr(dx, title) Then 'if found
            Do While rngF.Find.Execute(title, , , , , , True, wdFindStop)
    '            If InStr("�n�r�P�E", IIf(rngF.Characters(rngF.Characters.Count).Next Is Nothing, "", rngF.Characters(rngF.Characters.Count).Next)) = 0 And _
    '                InStr("�m�q�P�E", IIf(rngF.Characters(1).Previous Is Nothing, "", rngF.Characters(1).Previous)) = 0 Then
                    If �ѦW���g�W���Ъ`PreExamOK(d, title, rngF.start) Then
                        If VBA.IsNull(rst("���N��").Value) Then
                            rngF.text = "�q" & title & "�r"
                                      'd.Range.Find.Execute title, , , , , , True, wdFindContinue, , "�q" & title & "�r", wdReplaceAll
                        Else
                            rngF.text = rst("���N��").Value
                            'd.Range.Find.Execute title, , , , , , True, wdFindContinue, , rst("���N��").Value, wdReplaceAll
                        End If
                        rngF.SetRange rngF.End, d.Range.End
                    End If
    '            End If
            Loop
            Set rngF = d.Range: dx = d.Range.text
        End If
        
        rst.MoveNext
    Loop
    d.Range.Find.Execute "�m�m", , , , , , True, wdFindContinue, , "�m", wdReplaceAll
    d.Range.Find.Execute "�n�n", , , , , , True, wdFindContinue, , "�n", wdReplaceAll
    d.Range.Find.Execute "�q�q", , , , , , True, wdFindContinue, , "�q", wdReplaceAll
    d.Range.Find.Execute "�r�r", , , , , , True, wdFindContinue, , "�r", wdReplaceAll
    
    'GoSub bookmarks 'do again to check and correct SHOULD BE use another table to do this
    If ur.CustomRecordLevel > 0 Then
        SystemSetup.playSound 1.921
    Else
        SystemSetup.playSound 1
    End If
    rst.Close: cnt.Close: SystemSetup.contiUndo ur
    word.Application.ScreenUpdating = True
    
    
    Exit Sub
    
    
bookmarks:
    If rst.State = adStateOpen Then rst.Close
    rst.Open "select * from ���I�Ÿ�_�ѦW��_�۰ʥ[�W�� order by �Ƨ�", cnt, adOpenForwardOnly, adLockReadOnly
    Do Until rst.EOF
        title = rst("�ѦW").Value
        
    '    If title = "��v�qŲ" Then Stop
        
        If VBA.InStr(dx, title) Then 'if found
            Do While rngF.Find.Execute(title, , , , , , True, wdFindStop)
    '            If InStr("�n�r�P�E", IIf(rngF.Characters(rngF.Characters.Count).Next Is Nothing, "", rngF.Characters(rngF.Characters.Count).Next)) = 0 And _
    '                InStr("�m�q�P�E", IIf(rngF.Characters(1).Previous Is Nothing, "", rngF.Characters(1).Previous)) = 0 Then
                    If �ѦW���g�W���Ъ`PreExamOK(d, title, rngF.start) Then
                        
'                        If title = "��v�qŲ" Then Stop 'just for test
                        
                        If VBA.IsNull(rst("���N��").Value) Then
                            rngF.text = "�m" & title & "�n"
                '            d.Range.Find.Execute title, , , , , , True, wdFindContinue, , "�m" & title & "�n", wdReplaceAll
                        Else
                            rngF.text = rst("���N��").Value
                '            d.Range.Find.Execute title, , , , , , True, wdFindContinue, , rst("���N��").Value, wdReplaceAll
                        End If
                        rngF.SetRange rngF.End, d.Range.End
                    End If
    '            End If
            Loop
            Set rngF = d.Range: dx = d.Range.text
        End If
        
        rst.MoveNext
    Loop
    rst.Close
    Return
    
eH:
        Select Case Err.Number
            Case Else
                MsgBox Err.Number + Err.Description
    '            Resume
        End Select
End Sub


Sub ������q_�ھڲ�1�檺�r�ƪ��רӧ@����()
Dim wordCount As Byte, d As Document, rng As Range, i As Integer, dx As String, a, p As Paragraph, j As Byte, wl
Dim omitStr As String
omitStr = "{}<p>�m�n�q�r�G�A�C�u�v�y�z�@�P0123456789-" & ChrW(8231) & ChrW(183) & Chr(13)
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
dx = rng.text
wl = InStr(dx, Chr(13))
rng.text = left(dx, wl) & Replace(dx, Chr(13), "", wl)

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
Sub replaceWithNextChararcter() 'Alt+Shift+h
Dim s As Integer, chars 'As Characters
Dim f As String, r As String
Set chars = Selection.Characters
If chars.Count < 2 And InStr(Selection, Chr(9)) = 0 Then Exit Sub
If chars.Count > 2 Then
    s = InStr(Selection, Chr(9))
    If s > 0 Then
        If InStr(Mid(Selection.text, s + 1), Chr(9)) = 0 Then
            chars = VBA.Split(Selection.text, Chr(9))
            Selection.text = left(Selection.text, s - 1)
            s = 0
            f = chars(s): r = chars(s + 1) 'VBA.IIf(chars(s + 1) = Chr(9), "", chars(s + 1))
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If
Else
    s = 1
    f = chars(s)
    r = VBA.IIf(chars(s + 1) = Chr(9), "", chars(s + 1))
    Selection.Characters(s + 1) = ""
End If
Selection.Find.Execute f, , , , , , True, wdFindContinue, , r, wdReplaceAll
End Sub

Sub ��y�����}��ID�|�ʪ̦C�X()
Dim db As New dBase
db.��y�����}��ID�|�ʪ̦C�X
SystemSetup.playSound 12
End Sub
Sub ��y�����}��ID�|�ʪ̶�J()
Dim i As Long
ActiveDocument.Range.Find.Execute Chr(13), , , , , , , wdFindContinue, , "", wdReplaceAll
Do Until Selection.End = ActiveDocument.Range.End - 1
    Selection.move
    If Selection.Previous <> ChrW(20008) And Selection.Hyperlinks.Count = 0 Then
        �����r�[�W��y���`��
        ActiveWindow.ScrollIntoView Selection, False
        i = i + 1
    End If
    If i = 40 Then Exit Sub
Loop
Selection.HomeKey wdStory, wdExtend
End Sub

Rem 20230707 Bing�j���ġG �P�_���Υb�Φr
Public Function FullOrHalf(ByVal str As String) As Integer
    Dim strLocal As String
    Debug.Assert Len(str) = 1
    If Len(str) <> 1 Then
        FullOrHalf = -1
        Exit Function
    End If
    strLocal = StrConv(str, vbFromUnicode)
    If Len(str) * 2 = LenB(strLocal) Then
        FullOrHalf = 2 ' wide
    ElseIf Len(str) = LenB(strLocal) Then
        FullOrHalf = 1 ' narrow
    Else
        FullOrHalf = 0 ' error
    End If
End Function
Rem Bing�j���ġG
'�o�Ө�Ʊ����@�Ӧr�Ŧ�@����J�A��^�@�Ӿ�ƭȡC�p�G��^�Ȭ� 2�A�h��ܿ�J���r�ŬO�����F�p�G��^�Ȭ� 1�A�h��ܿ�J���r�ŬO�b���F�p�G��^�Ȭ� -1 �� 0�A�h��ܥX�{���~�C
'
'�P�_����z�O�N�s�X�q Unicode �ର���a�s�X�A�M�����ഫ�e��r�Ŧꪺ���סC�p�G�ഫ�e��r�Ŧ���׬۵��A�h��ܿ�J���r�ŬO�b���F�p�G�ഫ��r�Ŧ���׬O�ഫ�e�r�Ŧ���ת��⭿�A�h��ܿ�J���r�ŬO����(1)�C
'
'�ӷ�: �P Bing ����͡A 2023/7/7
'(1) �@������d�wvba�B�z�����b�� - ���G. https://zhuanlan.zhihu.com/p/600306305.
'(2) WordVBA�G�b���r���ର�����r�š]���X�d���k�^_word�b���Ÿ��אּ�����Ÿ���_VBA-�u�Ԫ��ի�-CSDN�ի�. https://blog.csdn.net/qq_64613735/article/details/124760907.
'(3) office�n��word���ɤ��p���O�b���M���� - �ʫת��D. https://zhidao.baidu.com/question/347564125.html.
'(4) VBA��Ƨ�q�N�N�r�ťѥ����ର�b���A�Υѥb���ର����-�P�ɾA��Excel Access - Excel��Ƥ��� - Office��y��. https://www.office-cn.net/excel-func/297.html.
'(5) �p�����word�峹�������I�O�����٬O�b���H_�ʫת��D. https://zhidao.baidu.com/question/45987987.html.


Rem �r����r��}�C creedit with chatGPT�j����
Function SplitWithoutDelimiter_StringToStringArray(str As String) As String()
Dim lenStr  As Long, arr() As String, i As Long, ch As String, eCount As Long
lenStr = VBA.Len(str)

'str =�n�ഫ���}�C���r��

' �N�r���ഫ���}�C
For i = 1 To lenStr
    ch = Mid(str, i, 1)
    eCount = eCount + 1
    If code.IsHighSurrogate(ch) Then
        ch = Mid(str, i, 2): i = i + 1
    End If
    ReDim Preserve arr(eCount - 1) ' �վ�}�C�j�p�A�Ϩ�P�r����׬ۦP
    arr(eCount - 1) = ch
Next i
SplitWithoutDelimiter_StringToStringArray = arr
End Function

