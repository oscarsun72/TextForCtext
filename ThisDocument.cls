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
Sub �b�㭶�Ҧ�������ܭ��������ťհϰ�()
    If word.ActiveWindow.View.ReadingLayout Then _
        ActiveWindow.View.ReadingLayout = Not ActiveWindow.View.ReadingLayout
    DoEvents
    word.ActiveWindow.View.DisplayPageBoundaries = False
End Sub

Sub Word��ƶפJAccess(wcontent As String, Terms As String, beforeTerm As String, acpath As String, before����, selRange As Range) '2010/10/20�׭q
'Dim w As Word.Document
Dim ����ID As Long '2010/10/29
Dim rst As DAO.Recordset, rst1 As Recordset, ss As String, �rID As Long ', �rnew As Boolean
Dim ����J�`�� As Boolean
Static ����J�r�� As Byte, ���л\���� As Byte
On Error GoTo ErrH:
'Set w = GetObject("C:\Documents and Settings\Superwings\�ୱ\0001�]�w�Z�^_����.mdb�ѷ�.doc")
'wcontent = w.Application.Run(ThisDocument.������ؤ��e�^��)
ss = VBA.Right(wcontent, 4)
wcontent = Left(wcontent, Len(wcontent) - 4)
If db Is Nothing Then
    Set db = DBEngine.Workspaces(0).OpenDatabase(acpath) '("D:\�d�{�@�o�N\�ѫH\�ϮѺ޲z\���K�a�Ѯv�U�z\����.mdb")
End If
'Terms = Mid(wcontent, InStr(wcontent, "�i") + 1, InStr(wcontent, "�j") - (InStr(wcontent, "�i") + 1))
Set rst = db.OpenRecordset("�r") 'If beforeTerm = "" Then beforeTerm = Terms�w�b�ҩI�s��"������ؤ��e�^��"�{�ǧP�_
If ss <> "XXX��" Then  '���o�rID��'ss <> "XXX�r"
    If Len(beforeTerm) > 1 Then MsgBox "�r�����פj��1,�е{�����g���ˬd!", vbExclamation: Stop
    With rst
        .Index = "�r"
        .Seek "=", beforeTerm
        If .NoMatch Then
            'MsgBox "�е{�����g���ˬd!", vbExclamation: Stop
            'GoTo a
            .AddNew
            .Fields("�r") = beforeTerm
            .Fields("����") = wcontent
            If before���� <> "" Then
                Set rst1 = db.OpenRecordset("����")
                With rst1
                    .Index = "����"
                    .Seek "=", before����
                    If .NoMatch Then
                        MsgBox "�е{�����g���ˬd!", vbExclamation: Stop '�]��214�����w�J,�Y���䤣��,�i��!
                    Else
                        rst.Fields("����ID") = .Fields("����ID") ' DLookup("����ID", "����", "���� = """ & before���� & """")
                        ����ID = .Fields("����ID")
                    End If
                    .Close
                End With
            End If
            .Update
            .Index = "�r" '�s�W��@�w�n�A��~��,�έ��}��ƪ��~��.
            .Seek "=", beforeTerm '���F���o.Fields("�s��ɶ�")��.Fields("�rID")����
            'Set rst = DB.OpenRecordset("select �rID from �r where �r = """ & Terms & """")'�����e��,�����A�����F
            �rID = .Fields("�rID") 'rst.Fields("�rID")
        Else
            �rID = .Fields("�rID") 'DLookup("�rID", "�r", "�r like """ & beforeTerm & """")
        End If
        '�ˬd�s���s���p
        If ss <> "XXX��" Then '�����B�z�b�U��!2010/10/29
            If Not IsNull(.Fields("�s��ɶ�")) Then
                Select Case ���л\����
                    Case 0
                        ���л\���� = MsgBox("�y" & .Fields("�r") & "�z�r�b" & .Fields("�s��ɶ�") & "���g�s���s�L" & _
                            "�O�_�פJ?--�פJ�N�|�л\<����>��Ƥ���y" & .Fields("�r") & "�z�r�w�g�s��L�����" _
                            & vbCr & vbCr & "�Y���L,�i�O�U���r,�b�����פJ������," & _
                            "�A�^<����>��Ʈw�˯��˵�" & vbCr & vbCr & "�n�פJ�л\����,�Ы�[�O]!! �n���L" _
                            & "�Ы�[�_]!  �o���פJ�@�ߤ��л\�Ы�[����]", vbYesNoCancel + vbExclamation, "�л\���ƽT�{!!!")
                        Select Case ���л\����
                            Case vbCancel
                                ���л\���� = vbCancel
                                GoTo ���פJ
                            Case vbNo
                                ���л\���� = vbNo
                                GoTo ���פJ
                            Case vbYes
                                ���л\���� = vbYes
                        End Select
                    Case vbNo
                        ���л\���� = MsgBox("�y" & .Fields("�r") & "�z�r�b" & .Fields("�s��ɶ�") & "���g�s���s�L" & _
                            "�O�_�פJ?--�פJ�N�|�л\<����>��Ƥ���y" & .Fields("�r") & "�z�r�w�g�s��L�����" _
                            & vbCr & vbCr & "�Y���L,�i�O�U���r,�b�����פJ������," & _
                            "�A�^<����>��Ʈw�˯��˵�" & vbCr & vbCr & "�n�פJ�л\����,�Ы�[�O]!! �n���L" _
                            & "�Ы�[�_]!  �o���פJ�@�ߤ��л\�Ы�[����]", vbYesNoCancel + vbExclamation, "�л\���ƽT�{!!!")
                        If ���л\���� <> vbYes Then GoTo ���פJ
                    Case vbCancel
                        GoTo ���פJ
                End Select
            Else
            End If
        End If
    End With
End If
'If Terms = "��" Then Stop '���T�n�]�h���^�|�����
'If Terms = "�a" Then Stop '�����n�|�����
Select Case ss
    Case "XXX�r"
        Set rst = db.OpenRecordset("�r") '�즹�欰a:����
        Dim selrangeP As Paragraph, zhu, Uzhu As Byte, ZhuYin, Char As Long, zhuyinText As String, pinyinText As String, charItem As Integer '��J�`���P����'2010/10/23
        Dim Zhuyins() As String, Pinyins() As String, zps As Byte '�@�r�h����, �γ��C�x�s
        Dim �`������rst As Recordset, �`��ID As Long, ����ID As Long
        zhu = �`���Ÿ�: Uzhu = UBound(zhu): Char = 1
        For Each selrangeP In selRange.Paragraphs
            charItem = 1: Char = charItem
            For Each ZhuYin In zhu
'                If ZhuYin = ChrW(20008) Then Stop
                If Char > selrangeP.Range.Characters.Count Then Exit For '�p�G�䧹���q,�Y���A��,���U�@�q
                If InStr(selrangeP.Range, ZhuYin) Then  '��X�n�B�z���q��-�Y�q�����t�`����(���t����)
                    selrangeP.Range.SetRange Char, selrangeP.Range.Characters.Count '���s���orange�d��--�b���e�@�ӭȱo(�Y�`���B�����}�C���U�@�ӯ��ޤ�����)
'                    charItem = InStr(selrangeP.Range, ZhuYin) '���o���q���Ĥ@�Ӫ`���Ÿ�����m--�D�],�D��靈�`���q����,�Ĥ@�Ӥ��쪺�`���Ÿ�(�P�}�C�����Ǧ����])!
                    For Char = charItem To selrangeP.Range.Characters.Count '�n�q���`�����q���q�Ĥ@�r,�v�r�ˬd,�Z�Y�`���Ÿ�,�Y�g�J�}�C(���]�`���b�����e�G�i�p�O�g�])
                    '�ҥH�����ǻP�`���Ÿ��}�C�����Ƿ������].���n�Ÿ����m��̫e�~�ण�|��,�_�h�n�A�[�@��,�~��j�M���q�̬O�_�٦��`���}�C�̪�����,�B�٦b�`���}�C����쪺,�b��q����e����m.���~����u�����Ĥ@�`���Ÿ���m!
                    '�M�ڤU���`���Ÿ���m,�õL���`���}�C�����Ǥ].
                        'If Not IsEmpty(Filter(zhu, selrangeP.Range.Characters(Char), False, vbBinaryCompare)) Then  '���`���Ÿ�
                        '��ΤW��,���]�䤣��̵M�Ǧ^�}�C(�P�u�W�����Ҩ�����!!),�ߨ���@�ӫh�}�C���ޭȷ|��@,�G��g���U��:
                        If UBound(Filter(zhu, selrangeP.Range.Characters(Char), False, vbBinaryCompare)) < Uzhu Then '���`���Ÿ�
                            If selrangeP.Range.Characters(Char) = "�D" And selrangeP.Range.Characters(Char).Font.Name <> "�з���" Then GoTo nextP '"�D"�����n�Ÿ��P�Ʀr�s���Ÿ��ۦP,�G�u��H�r��,�Ψ�ᱵ�r���O�_���`���Ÿ��P�_�F.���ȥH�r�Φ�.
                            charItem = Char 'InStr(selrangeP, selrangeP.Range.Characters(Char)) '���o�Ĥ@�Ӫ`���Ÿ�����m
                            If selrangeP.Range.Characters(Char) <> ChrW(20008) Or (selrangeP.Range.Characters(Char) <> ChrW(20008) And InStr(selrangeP.Range, "�j") = 0) Then
11                              Do Until selrangeP.Range.Characters(charItem).Font.Name = "SimSun" Or selrangeP.Range.Characters(charItem).Font.Name = "Pintone Times" Or selrangeP.Range.Characters(charItem) = Chr(13)
                                    If UBound(Filter(zhu, selrangeP.Range.Characters(Char), False, vbBinaryCompare)) < Uzhu Then
                                        zhuyinText = zhuyinText & selrangeP.Range.Characters(charItem)
                                    End If
                                    charItem = charItem + 1
                                Loop
                                zhuyinText = Trim(zhuyinText) '���o�`��
                                Do Until (selrangeP.Range.Characters(charItem).Font.Name <> "SimSun" And selrangeP.Range.Characters(charItem).Font.Name <> "Pintone Times") Or selrangeP.Range.Characters(charItem) = Chr(13)
                                    'If UBound(Filter(zhu, selrangeP.Range.Characters(Char), False, vbBinaryCompare)) < Uzhu Then'���b�`���ᥲ������,�G���i�p�O�P�_
                                        pinyinText = pinyinText & selrangeP.Range.Characters(charItem)
                                    'End If
                                    charItem = charItem + 1
                                Loop
                                pinyinText = Trim(pinyinText) '���o����
                                'GoTo nextP '����
                                ReDim Preserve Zhuyins(zps)
                                ReDim Preserve Pinyins(zps)
'                                ReDim Zhuyins(zps)'�٬O�nPreserve�~��O�}�C�e�@���ޭ�
'                                ReDim Pinyins(zps)
                                Zhuyins(zps) = zhuyinText
                                Pinyins(zps) = pinyinText
                                zps = zps + 1
                                zhuyinText = ""
                                pinyinText = ""
                                Char = charItem
                            Else
'                                If InStr(selrangeP.Range, Char) > InStr(selrangeP.Range, "�j") Then
                                If Char > InStr(selrangeP.Range, "�j") Then
                                    GoTo 11
                                End If
                            End If
'                        Else
'                            Stop
                        End If
                    Next Char
                End If
            Next
nextP: Next
        'If AscW(Terms) = 25908 Then Stop
'        If Terms = "��" Then Stop
'        If Terms = "��" Then Stop
'        If Terms = "��" Then Stop
        With rst
            .Index = "�r"
            .Seek "=", Terms
            If .NoMatch Then
                .AddNew
            Else
                .Edit
            End If
            .Fields("�r") = Terms
            .Fields("����") = wcontent
            If IsNumeric(Left(wcontent, 5)) Then
                Set rst1 = db.OpenRecordset("�r")
                With rst1
                    .Index = "�r��"
                    .Seek "=", Left(wcontent, 5)
                    If .NoMatch = False Then
                        If .Fields("�rID") <> �rID Then
                            If ����J�r�� <> vbNo Then
                                ����J�r�� = MsgBox("�r��" & Left(wcontent, 5) & "���ƤF!�n�~��,�h���|�פJ�r��!" & vbCr & vbCr _
                                    & "  �n�פ�o���פJ,�H�勵���ɮ�,�Ы�[����]!!" & _
                                    "�n�~��,�Ы�[�O].   ���A����r��,�������L(����J�r��)�h��[�_]   " & _
                                    "���U�e�����f��,���r�Ǧ��~���r,����ؤ��e�N�|�ϥ����!", vbExclamation + vbYesNoCancel, "�r�ǭ��ƤF!!!")
                                Select Case ����J�r��
                                    Case vbCancel
                                        'ActiveDocument.Windows(1).Visible = True
                                        Documents.Open ActiveDocument.FullName, , False
                                        End
                                End Select
                            End If
                        Else
                            rst.Fields("�r��") = Left(wcontent, 5)
                        End If
                    Else
                        rst.Fields("�r��") = Left(wcontent, 5)
                    End If
                End With
            End If
            If before���� <> "" Then '���o�r������ID
                Set rst1 = db.OpenRecordset("����")
                With rst1
                    .Index = "����"
                    .Seek "=", before����
                    If .NoMatch Then
                        MsgBox "�е{�����g���ˬd!", vbExclamation: Stop '�]��214�����w�J,�Y���䤣��,�i��!
                    Else
                        If Not IsNull(rst.Fields("����ID")) Then
                            If .Fields("����ID") <> rst.Fields("����ID") Then
                                Dim rst1clone As Recordset
                                Set rst1clone = rst1.Clone
                                With rst1clone
                                    .Index = "����ID"
                                    .Seek "=", rst.Fields("����ID")
                                End With
                                If MsgBox("�y" & rst.Fields("�r") & "�z�r�����w���y" _
                                    & rst1clone.Fields("����") & "�z��,�O�_��אּ�y" & .Fields("����") & _
                                    "�z���H", vbExclamation + vbYesNo, "�����ܧ�T�{!") = vbYes Then
                                    rst.Fields("����ID") = .Fields("����ID")
                                    rst1clone.Close
                                    Set rst1clone = Nothing
                                End If
                            End If
                        Else
                            rst.Fields("����ID") = .Fields("����ID") ' DLookup("����ID", "����", "���� = """ & before���� & """")
                        End If
                    End If
                    .Close
                End With
            End If
            If zps = 0 Then
                If .Fields("����J�`��") = False Then
                    If MsgBox(Terms & "�r�ʪ`��!�O�_��w��Z��A���s����@���פJ�H" _
                        & vbCr & vbCr & "�T�w�פ�o���פJ�A�Ы��e�T�w�f�s�I" & vbCr & vbCr & _
                        "���U�e�T�w�f��,���L�`�����r,����ؤ��e�N�|�ϥ����!", vbCritical + vbOKCancel) = vbOK Then
                        'ActiveDocument.Windows(1).Visible = True
                        Documents.Open ActiveDocument.FullName, , False
                        End
                    Else
                        ����J�`�� = True
                    End If
                Else
                    ����J�`�� = True
                End If
            End If
            If zps > 0 Then zps = zps - 1
            .Fields("�s��ɶ�") = Null '�M���s��аO,�H�K�A���߰��л\�T�{2010/10/27
            If ����J�`�� Then .Fields("����J�`��") = True '�O�U����J�r�����r
            .Update
        End With
        If ����J�`�� = False Then
            Set rst = db.OpenRecordset("�r_�`��") '��J�r_�`��
            With rst
    '            If zps > 0 Then Stop
                For Char = 0 To zps  'UBound(Zhuyins(zps))
                    Set �`������rst = db.OpenRecordset("�`��") '�ˬd�`�����L
                    With �`������rst
                        If InStr(1, "�D���E" & ChrW(12539), Left(Zhuyins(Char), 1), vbBinaryCompare) Then 'Debug.Print Terms: Stop
                            If MsgBox("�`���u" & Zhuyins(Char) & "�v�O�_��J�H" & vbCr & vbCr & "�]�����n�Ÿ����P�����I�Ÿ��ۦP�̡A�G�{���i��|����C�ЧP�_�I", vbExclamation + vbOKCancel + vbDefaultButton2, "���n�`����J�T�{�I�I") = vbCancel Then GoTo exitZhuyin
                        End If
                        Zhuyins(Char) = Replace(Replace(Replace(Zhuyins(Char), "��", "�E"), "�D", "�E"), ChrW(12539), "�E")
21                      .Index = "�`��"
                        .Seek "=", Zhuyins(Char)
                        If .NoMatch Then
                            .AddNew
                            .Fields("�`��") = Zhuyins(Char)
                            .Update
                            '.Close '�G�M��������,�����h�L�ĤF,�o���sset
                            .OpenRecordset
                            GoTo 21 '�s�W��A�h��`��ID��
                        Else
                            �`��ID = .Fields("�`��ID")
                        End If
                    End With
                    Set �`������rst = db.OpenRecordset("����") '�ˬd�������L
                    With �`������rst
22                      .Index = "����"
                        .Seek "=", Pinyins(Char)
                        If .NoMatch Then
                            .AddNew
                            .Fields("����") = Pinyins(Char)
                            .Update
                            .OpenRecordset
                            GoTo 22 '�s�W��A�h��`��ID��
                        Else
                            ����ID = .Fields("����ID")
                        End If
                    End With
                    
                    .Index = "PrimaryKey"
                    .Seek "=", �rID, �`��ID
                    If .NoMatch Then
                        .AddNew
                    Else
                        .Edit
                    End If
                    .Fields("�rID") = �rID
                    .Fields("�`��ID") = �`��ID
                    .Fields("����ID") = ����ID
                    .Update
exitZhuyin:    Next Char
            End With
        End If
        'beforeTerm = Terms
    Case "XXX��"
        Set rst = db.OpenRecordset("��")
        With rst
            .Index = "��"
            .Seek "=", Terms
            If .NoMatch Then
                .AddNew
            Else
                '�ˬd�s���s���p
                If Not IsNull(.Fields("�s��ɶ�")) Then
                    If MsgBox("�y" & .Fields("��") & "�z�@���b" & .Fields("�s��ɶ�") & "���g�s���s�L" & _
                        "�O�_�פJ?--�פJ�N�|�л\<����>��Ƥ���y" & .Fields("��") & "�z�����w�g�s��L�����" _
                        & vbCr & vbCr & "�Y���L,�i�O�U���r,�b�����פJ������," & _
                        "�A�^<����>��Ʈw�˯��˵�" & vbCr & vbCr & "�n�פJ�л\����,�Ы�[�T�{]!! �n���L" _
                        & "�Ы�[����]!", vbOKCancel + vbExclamation, "�л\���ƽT�{!!!") = vbCancel Then _
                            GoTo ���פJ
                End If
                .Edit
            End If
            .Fields("�rID") = �rID
            .Fields("��") = Terms
            .Fields("����") = wcontent
            .Fields("�s��ɶ�") = Null
            .Update
        End With

    Case "XXX�l"
        Set rst = db.OpenRecordset("�l")
        With rst
            .Index = "�rID" '��@"�rID1"����!
            .Seek "=", �rID
            If .NoMatch Then
                .AddNew
            Else
                .Edit
            End If
            .Fields("�rID") = �rID
            .Fields("�l") = wcontent
            .Update
        End With
    Case "XXX��"
        Set rst = db.OpenRecordset("����")
        With rst
            .Index = "����"
            .Seek "=", Terms
            If .NoMatch Then
                .AddNew
                .Fields("����") = Terms
                .Update
'            Else
'                .Edit
            End If
        End With
    Case Else
        MsgBox "��bug,�гq���{�����g�H!", vbCritical
End Select
���פJ:
Exit Sub
ErrH:
    MsgBox Err.Number & Err.Description: Stop: Resume
End Sub


Function �`���Ÿ�()
�`���Ÿ� = Array("�t", "�u", "�v", "�w", "�x", "�y", "�z", "�{", "�|", "�}", "�~", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "�E", "��", "��", "��", ChrW(20008), "��", "�D", ChrW(12539)) '���n�Ÿ��Y���|��!!'�H���u�E�v����!!
'�ҥH�����ǻP�`���Ÿ��}�C�����Ƿ������].���n�Ÿ����m��̫e�~�ण�|��,�_�h�n�A�[�@��,�~��j�M���q�̬O�_�٦��`���}�C�̪�����,�B�٦b�`���}�C����쪺,�b��q����e����m.���~����u�����Ĥ@�`���Ÿ���m!
'�y���Ĥ@�ӧ�쪺��m,�h�`���Ÿ������ǧY�����n!!
'���O�p�y������ӽT���`�r???
End Function



Sub ������ؤ��e�^��(w As Document, acpath As String)
Dim p As Paragraph, pselect ' As Range
Dim a As Object 'Access.Application
Dim beforeTerm As String, Terms As String, wcontent As String, before���� As String
Dim i As Long, e As Boolean
Dim docEnd As Long
On Error GoTo ErrH
'Set A = GetObject("C:\Documents and Settings\Superwings\�ୱ\����.MDB")
Set a = GetObject(acpath)
With w.ActiveWindow.Selection
    docEnd = .Document.Range.End
    'If MsgBox("�O�_�q�Y(���)�}�l?", vbYesNo + vbQuestion) = vbYes Then .HomeKey wdStory, wdMove
    '��g�����.mdb��
    Do
        If .Paragraphs(1).Range.Font.Hidden = False And .Paragraphs(1).Range.Font.Hidden <> 9999999 Then
                            '9999999�D�]�A���ûP�����ä].
'            If InStr(.Paragraphs(1).Range, "��") Then Stop
'            If InStr(.Paragraphs(1).Range, "��") Then Stop
            If Left(.Paragraphs(1).Range, 1) = "��" Then ''�r�������B�z
'                If InStr(.Paragraphs(1).Range, "�쳡") Then Stop
'                If InStr(.Paragraphs(1).Range, "�鳡") Then Stop
                GoSub ���
                wcontent = .Range
                wcontent = wcontent & "XXX��"
                Terms = Mid(wcontent, 2, 1)
                If beforeTerm = "" Then beforeTerm = Terms
                before���� = Terms
            ElseIf InStr(Selection.Paragraphs(1).Range, "�i") Then
                GoSub ���
    '            .Paragraphs(1).Range.Select
    '            Do Until InStr(.Next.Paragraphs(1).Range, "�i") Or Left(.Next.Paragraphs(1).Range, 1) = "��"
    '                .MoveDown wdParagraph, 1, wdExtend
    '            Loop
    '            Terms = Mid(wcontent, InStr(wcontent, "�i") + 1, InStr(wcontent, "�j") - (InStr(wcontent, "�i") + 1))
                wcontent = Replace(.Range, Chr(13), Chr(13) & Chr(10))
                wcontent = Replace(wcontent, "�f", "�f" & Chr(13) & Chr(10))
                wcontent = Replace(wcontent, "�e", Chr(13) & Chr(10) & "�e")
                Terms = Mid(wcontent, InStr(wcontent, "�i") + 1, InStr(wcontent, "�j") - (InStr(wcontent, "�i") + 1))
                Select Case InStr(Selection, "�j") - InStr(Selection, "�i")
                    Case Is = 2 '�r
                        wcontent = wcontent & "XXX�r"
                        'A.Run "Word��ƶפJ_�r', ������ؤ��e�^��, Terms"
    '                    Word��ƶפJAccess wcontent, Terms, beforeTerm
                        beforeTerm = Terms
                    Case Is > 2
                        wcontent = wcontent & "XXX��"
                        'A.Run "Word��ƶפJ_��", ������ؤ��e�^��, Terms, beforeTerm
    '                    Word��ƶפJAccess wcontent, Terms, beforeTerm
                End Select
            ElseIf InStr(.Paragraphs(1).Range, "��") Then ''�r���l�͵��B�z
                GoSub ���
                wcontent = .Range
                wcontent = wcontent & "XXX�l"
                'A.Run "Word��ƶפJ_�l", ������ؤ��e�^��, beforeTerm
    '            Word��ƶפJAccess wcontent, Terms, beforeTerm
            Else
                GoTo NextOne '���B�z
            End If
            If wcontent <> "" Then Word��ƶפJAccess wcontent, Terms, beforeTerm, acpath, before����, .Range
        End If
        'If .Range.End <> .Document.Range.End Then .MoveDown wdParagraph, 1, wdMove
NextOne: If e = False Then .MoveDown wdParagraph, 1, wdMove
        i = i + 1
        w.Application.StatusBar = "���b�B�z��" & i & "��"
        a.Forms("�D����").Controls("Word��ƶפJ").Caption = "���b�B�z��" & i & "��"
        a.Forms("�D����").Repaint
        
    Loop Until e '.Range.End = .Document.Range.End
'    MsgBox "�פJ����!", vbInformation' 2010/10/20'����g��Access�̭�!
'    a.Forms("�D����").Controls("Word��ƶפJ").Caption = "Word��ƶפJ"
    If w.Application.Documents.Count = 0 Then
        w.word.Application.Quit wdDoNotSaveChanges
    Else
        w.Close wdDoNotSaveChanges
    End If
    AppActivate Replace(a.CurrentProject.Name, ".mdb", "")
    With a.Forms("�D����")
        .Controls("Word��ƶפJ").FontSize = 8
        .Controls("Word��ƶפJ").Caption = "��" & i & "���B�z����,���b�פJ�����B�����P���y,�еy��..."
        .Repaint
    End With
    db.Close
    Set db = Nothing
Exit Sub
���:
    .Paragraphs(1).Range.Select
    If .Range.End = docEnd Then e = True: Return
    Do Until InStr(.Next.Paragraphs(1).Range, "�i") Or _
            Left(.Next.Paragraphs(1).Range, 1) = "��" Or _
            Left(.Next.Paragraphs(1).Range, 1) = "��" Or _
            .Next.Paragraphs(1).Range.Font.Hidden = (True Or 9999999)
        .MoveDown wdParagraph, 1, wdExtend
        If .Range.End = docEnd Then e = True: Exit Do
    Loop
    Return
End With
'With ActiveDocument
'    For Each p In .Paragraphs
'        If InStr(p, "�i") Then
'            If Left(p, 1) <> "��" Then
'                'p.Range.SetRange
'                p.Range.Select
'                'Selection.MoveEndUntil Chr(13) & Chr(13), wdForward    '(Chr(13) & Chr(13)) ' "�i"
'    '            .Range.MoveEndUntil Chr(13) & Chr(13), wdForward
'                Selection.MoveEndUntil "�i", wdForward
'    '            p.Range.MoveEndUntil "�i", wdForward
'                If InStr(Selection, "��") Then
'                    'Selection.MoveStartUntil "��", wdForward
'                    'Selection.MoveLeft wdParagraph, 2, wdExtend
'                    Selection.MoveUp wdParagraph, 3, wdExtend '  "��"
'                Else
'                    If Selection.Next.Font.Size > 12 Then '=20
'                        Selection.MoveLeft wdCharacter, 6, wdExtend
'                    End If
'                End If
'    '            p.Range.Copy
''                .Windows(1).ScrollIntoView Selection, False
'                ������ؤ��e�^�� = Selection.Text
'            Else
'                '�r���l�͵��B�z
'                Selection.Paragraphs (1)
'            End If
'        ElseIf Left(p, 1) <> "��" Then
'
'        End If
'    Next
'End With
ErrH:
Select Case Err.Number
    Case Else
        If MsgBox("�X���F,�гq���t�μ��g��!" & vbCr & _
        "���~�ԭz:" & Err.Number & Err.Description & vbCr & vbCr & _
        "���T�w������.������������.", vbExclamation + vbOKCancel) = vbCancel Then Stop: Resume
        
End Select
End Sub










Sub �~�y�����m��() 'PinTone��SimSun
Dim db As Database, p As Paragraph, chra As word.Characters, m As Long, pinL As Long
Dim rst As Recordset, rstReplace As Recordset, Paragraph As String, ParagraphNew As String, pL As Integer, pS As Integer
Dim zhu, Uzhu, i As Byte, zhuL As Integer ', zhuYes As Boolean
Dim tble, t As Byte, b As Byte
Dim ds As Date, de As Date, times As Long
Static wpath As String
wpath = "D:\�d�{�@�o�N\�ѫH\�ϮѺ޲z\���K�a�Ѯv�U�z\����.mdb"
If Dir(wpath) = "" Then wpath = "C:\Documents and Settings\Superwings\�ୱ\����.mdb"
'wpath = InputBox("�п�J���פJ�ɮת������|�ɦW(�t���ɦW)", , _
    IIf(Nz(wpath, "") = "", "C:\Documents and Settings\Superwings\�ୱ\����.mdb", wpath))
wpath = InputBox("�п�J���פJ�ɮת������|�ɦW(�t���ɦW)", , wpath)
If wpath = "" Then Exit Sub

Set db = DBEngine.OpenDatabase(wpath)
ds = Now
'zhu = Array("�t", "�u", "�v", "�w", "�x", "�y", "�z", "�{", "�|", "�}", "�~", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", ChrW(20008))
'Uzhu = UBound(zhu)
Set rstReplace = db.OpenRecordset("Pintone���N�������άd��")
With rstReplace
    Do Until .EOF
    '                If .Fields("����Pintone") = "y9" Then Stop
        Paragraph = .Fields("����Pintone")
        ParagraphNew = .Fields("����")
        With Selection.Range.Find
            .Text = Paragraph
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
                .Replacement.Text = ParagraphNew
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

'MsgBox "����!", vbInformation
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
'        Paragraph = p.Range 'Nz(.Fields("����"))
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
'                If InStr(Paragraph, "�i") Then zhuL = InStr(Paragraph, "�i")
'                Paragraph = Mid(Paragraph, zhuL)
'                ParagraphNew = Paragraph
'
'                With rstReplace
'                    Do Until .EOF
'        '                If .Fields("����Pintone") = "y9" Then Stop
'                        ParagraphNew = Replace(ParagraphNew, .Fields("����Pintone"), .Fields("����"), , , vbBinaryCompare)
'                        .MoveNext
'                    Loop
'                    .MoveFirst
'                End With
''                .Edit
''                .Fields("����") = Replace(.Fields("����"), Paragraph, ParagraphNew)
''                .Update
'                p.Range.Text = Replace(p.Range.Text, Paragraph, ParagraphNew)
'                zhuL = 0 ''�k�s
'            End If
'            'pS = pL + 2
'            'pL = InStr(pS, .Fields("����"), Chr(13))
'            'Paragraph = Nz(.Fields("����"))
''        Loop
'        'pS = 0 ''�k�s
'        'End If
''        .MoveNext
'    'Loop
'    Next p
''End With
'Next t
de = Now
MsgBox "����!" & vbCr & "�O��" & Minute(de - ds) & "��" & Second(de - ds) & "��" & vbCr _
    & "�@" & times & "���m��", vbInformation

End Sub

Private Sub Document_Close()
''�������ɥH���Ҥ覡�O����Ц�m
'With ThisDocument
'If Not .path = "" Then '���ɤ~�ˬd2003/4/1
'    If .Saved = False Then
'        If .Name = "Normal.dotm" Then
'            If MsgBox("�n�x�s""Normal.dotm""�d���ܡH", vbQuestion + vbOKCancel) = vbOK Then .Save
'            Exit Sub
'        End If
'        If DonotSave = False Then
'            Select Case MsgBox("�n�O�U��г̫��m��?", vbYesNoCancel + vbInformation)
'                Case Is = vbYes
'                    ��ЩҦb��m����
'            '        ActiveDocument.Save
'                Case vbNo
'                    If MsgBox("���w�ק�,�O�_�n�x�s", vbExclamation + vbYesNo) = vbYes Then
'                        .Save
'
'                    Else
'                        .Saved = True '�p���]�wWord�N���A�߰ݬO�_�n�x�s����ܮ�!2003/3/22
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
''If Tasks.Exists("����") Then
''    If Tasks("����").Visible = False Then Tasks("����").Close
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
Register_Event_Handler '��ӷ|���ĬO�]���ڦb"�b�t�@��󤤴M�����r��"�ΤFEND��!
If ActiveDocument.AttachedTemplate.Name = "Normal.dot" Then
    CommandBars.Item("blog").Visible = True
Else
    CommandBars.Item("blog").Visible = False
End If
Docs.DocBackgroundFillColor
End Sub


Private Sub Document_Open()
'Public Sub Document_Open() '��L�d���ѷӥ�!�G����Private'2004/10/17
'On Error GoTo eH

Dim bk As bookmark, bkIdx As Integer, tp As Template '2015/11/4
Register_Event_Handler
With ThisDocument
    If Left(.Name, 2) = "�ƦL" Then .ActiveWindow.DisplayVerticalRuler = False
    With .Application.CommandBars
        If Application.Templates.Count > 1 Then
            For Each tp In Application.Templates
                If tp.Name = "mywork.dot" Then If .Item("�o�h��").Visible Then .Item("�o�h��").Visible = False
            Next
        End If
        If Left(ThisDocument, 7) Like "Access_" _
            Or InStr(ThisDocument.FullName, "��r�ɸ�Ʈw") Then _
            .Item("�פ�ĥ��O�s��").Visible = True
        If .Item("Web").Visible = True Then .Item("web").Visible = False
    End With
    If .bookmarks.Count > 0 Then '�����Үɤ~��
        '        For bkIdx = 1 To .Bookmarks.Count
                For Each bk In .bookmarks '�N��ѫe���Цa���ҧR��'2003/3/28
                    bkIdx = bkIdx + 1 '�O�U���ү���
                    With bk '�p�G�O�s��B�~�B�z
                        If InStr(1, .Name, "Edit_", vbTextCompare) > 0 _
                            And InStr(1, bk, "�s��", vbTextCompare) = 0 Then
                            '�p�G�O��ѫe��
                            Do While InStr(1, bk, "Edit_", vbTextCompare) > 0 And _
                                    CDate(Replace(Mid(bk, 6, 10), "_", "/")) <= Date - 2 _
                                    And InStr(1, bk, "�s��", vbTextCompare) = 0
                                    If ThisDocument.Saved Then
                                        If MsgBox("�n�R���e�ѥH�e�s�褧��m���Ҷ�?", vbYesNo + vbExclamation) = vbYes Then
                                            bk.Delete '�����H����ޭȷ|�V�e����'�d�U�ݽs��B2003/4/2
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
        If .Item("web").Visible = True Then .Item("web").Visible = False
    End With
'    .ActiveWindow.WindowState = wdWindowStateMaximize
'    .ActiveWindow.Application.GoBack '��̫�s��B,�@�ε��P��� Shift+F5 ��
    StatusBar = .AttachedTemplate.Name & " template is attached to the active document"
    Options.CheckGrammarAsYouType = False

End With
111
'If MsgBox("�O�_�n�i��y�r�r���ˬd?", vbQuestion + vbOKCancel + vbDefaultButton2) = vbOK Then �y�r�r���ˬd
'�b���A�C��ܥ����ҨϥΪ��d���W��
'Exit Sub
'eH:
'Select Case Err.Number
'    Case 4248 '�]���S���}�Ҥ��A�ҥH�L�k�ϥγo�ӫ��O�C
'        GoTo 111 'Resume Next'91�S���]�w�����ܼƩ� With �϶��ܼ�
'    Case Else
'        MsgBox Err.Number & vbTab & Err.Description
'End Select
DoEvents
Me.�b�㭶�Ҧ�������ܭ��������ťհϰ�
If userProfilePath = "" Then userProfilePath = SystemSetup.���o�ϥΪ̸��|_�t�ϱ׽u()
End Sub
