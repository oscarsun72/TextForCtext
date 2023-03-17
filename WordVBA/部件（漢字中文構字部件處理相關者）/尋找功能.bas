Attribute VB_Name = "�M��\��"
Option Explicit
Public ThisDocument�`��5032�~�r�����Dictionary As New Scripting.Dictionary

Sub �M��e�ϫ�r()
Dim isp As InlineShape
Static x As String, n As String, nw As Long
x = InputBox("�п�J�Ϥ��R�W�ԭz", , x)
If x = "" Then Exit Sub
n = InputBox("�п�J�Ϥ�����䤧��r,�Ψ�Ascw��", , nw)
If n = "" Then Exit Sub
If IsNumeric(n) Then
    nw = n
    n = ChrW(CLng(n))
Else
    nw = n
End If
For Each isp In ActiveDocument.InlineShapes
    If isp.AlternativeText = x Then
        If isp.Range.Characters(1).Next = n Then
            ActiveDocument.Range(isp.Range.Start, isp.Range.Characters(1).Next.End).Select
            Exit Sub
        End If
    End If
Next
MsgBox "�S���!", vbExclamation
End Sub

Sub �˯�������s()
Dim d As Document, t As Table, c As Cell, a, i As Integer, w As String, cln As Byte, clnSearch As Long, flg As Boolean, ck As Boolean, cInPut As Cell
Set d = ThisDocument
Set t = d.Tables(1)
For Each c In t.Rows(1).Cells
    clnSearch = clnSearch + 1
    If InStr(c.Range, "�˯������") > 0 Then
        flg = True
        Exit For
    End If
Next c
If flg = False Then
    MsgBox "�䤣��""�˯������""���", vbCritical
    Exit Sub
Else
    flg = False
End If
For Each c In t.Rows(1).Cells
    cln = cln + 1
'    If InStr(c.Range, "����]�X��") > 0 Then
    If InStr(c.Range, "����]���") > 0 Then
        flg = True
        Exit For
    End If
Next c
If flg = False Then
'    MsgBox "�䤣��""����]�X��""""���", vbCritical
    MsgBox "�䤣��""����]���""""���", vbCritical
    Exit Sub
End If
For Each c In t.Columns(cln).Cells
    i = i + 1
    If i > 1 Then
        For Each a In c.Range.Characters
            If a <> Chr(13) & Chr(7) Then
                If a.InlineShapes.Count > 0 Then
                    w = w & a.InlineShapes(1).AlternativeText
                Else
                    w = w & a
                End If
            End If
        Next a
'        t.Cell(c.Row.Index, clnSearch).Range.Text = w'�ӺC�A�G���.Next
        Set cInPut = c.Next.Next.Next
        If Not ck Then
            If t.Cell(1, cInPut.ColumnIndex).Range.Text <> "�˯������" & Chr(13) & Chr(7) Then
                MsgBox "�{�����~�A��J�����D�u�˯������v", vbCritical
                cInPut.Select
                Exit Sub
            Else
                ck = True
            End If
        End If
        cInPut.Range.Text = w
        w = ""
    End If
Next c
MsgBox "done!", vbInformation
End Sub


Rem 20230316 creedit with Bing�j����
'�d�ݥu�]�����ܧΦӺc�����P�~�r����Ҧ��������~�r�]�Y���F�����ܧΥ~�A��L������ҦP�C�|���p��ƦC�覡�^
Rem �糡�ܧγ������W�U�̡A�������A�u�����ҡA���i�H�H�u��X
Sub ��X�Ҧ��u���@���ܧγ������󤣦P���rDictionary()
Dim t As Table, d As Document, a As Range, c As Cell, result As String
Dim dictComponents As New Scripting.Dictionary, components As New components, componentsArray() As String, w As String, wList As String, key As String
Dim dictVariantRadicals As Scripting.Dictionary, variantRadical As String, w_ofVariantRedical As String, w_ofVariantRedical_componentsCollection As VBA.Collection
Dim sameComposeDict As New Scripting.Dictionary '�߰��ܧγ����~���P����զ����M�Ӧh�A����� Dictionary sameComposeDict �x�s key=���� value= �r �H�K��X�P�����P�Ӧ���ӥH�W���~�r�~��J
Const columnComponents As Byte = 2 '�������
Const columnChar As Byte = 1 '�~�r���

'���o�~�r�Ψ䳡����
Set d = ThisDocument
Set t = d.Tables(1)
'���o�ܧγ������
Set dictVariantRadicals = components.VariantRadicalsDictionary
For Each c In t.Columns(columnComponents).Cells
    If c.RowIndex > 1 Then '��1�C�O���D��
        '���o�~�r
        Set a = t.Cell(c.RowIndex, columnChar).Range
        w = VBA.Mid(a.Text, 1, Len(a.Text) - 2) '�x�s��аO�D���� chr(13) & chr(7),�n����
        '���o����}�C componentsArray �]�޼ƥH�ǧ}�]pass by reference�^�覡�ǻ��^
        components.getComponentsArray componentsArray, c.Range
        '�ƦC����}�C�]�Y�������󪺱ƦC�覡�]�b��Ҳզ����~�r������m�^�^
        Call components.SortStringArray(componentsArray)
        If UBound(componentsArray) > 0 Then
            Dim arr, earr, iarr As Long, subsetExcludingVariantRadicalsDict As Scripting.Dictionary, e_earr
            
'            If w = "��" Then sndPlaySound32 "C:\Windows\Media\Ring10.wav", 1: Stop 'for debug
'            If w = "�s" Then sndPlaySound32 "C:\Windows\Media\Ring10.wav", 1: Stop 'for debug
'            If w = "��" Then Beep: Stop 'for debug
'            If w = "��" Then Beep: Stop 'for debug
'            If w = "��" Then sndPlaySound32 "C:\Windows\Media\Ring10.wav", 1: Stop 'for debug
'            If w = "�E" Then Beep: Stop 'for debug
            
            '���o�������������󶰦X��
            Set subsetExcludingVariantRadicalsDict = components.subsetExcludingVariantRadicals(componentsArray)
            If Not subsetExcludingVariantRadicalsDict Is Nothing Then
                Dim cln As VBA.Collection
                Dim clnKeysArr
                Dim eClnKeysArr
                Dim radical_variantRadical As String, radical_in_dictComponents As String
                '���o�����ܧγ���������զ����X�A�O�@�ӤG���}�C�]�bCollection�̭�
                clnKeysArr = subsetExcludingVariantRadicalsDict.Keys
                For Each eClnKeysArr In clnKeysArr
                    Set cln = eClnKeysArr
                    radical_variantRadical = dictVariantRadicals(subsetExcludingVariantRadicalsDict(cln))
                    Rem Bing�j���ġG
'                    �ھں����j�M�����G�AVBA �� Dictionary �������O����ȡ]key�^�i�H�O�����������A�]�A Collection12�C���O�A�p�G�n�ϥ� Collection �@����ȡA�ݭn�`�N�H�U�X�I3�G
                    'Collection �@����ȮɡA�����O�@�����C
                    'Collection �@����ȮɡA�������ۦP�������ӼƩM���Ǥ~��Q�����ۦP����ȡC
                    'Collection �@����ȮɡA���ઽ���ί��ީ� Keys ��k�Ӧs���A�ݭn���ഫ�� Variant �� String �����C�K�K
                    '�ӷ�: �P Bing ����͡A 2023/3/17(1) Keys method (Visual Basic for Applications) | Microsoft Learn. https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/keys-method �w�s�� 2023/3/17.
                    '(2) Dictionary ���� | Microsoft Learn. https://learn.microsoft.com/zh-tw/office/vba/language/reference/user-interface-help/dictionary-object �w�s�� 2023/3/17.
                    '(3) excel - VBA get key by value in Collection - Stack Overflow. https://stackoverflow.com/questions/12561539/vba-get-key-by-value-in-collection �w�s�� 2023/3/17.
                    arr = cln.Item(1)
                    '�p�G�O�}�C
                    If TypeName(arr) = "Variant()" Then
                        For iarr = LBound(arr) To UBound(arr)
                            '���o�G���}�C�����@���}�C�����s�J�ܼ� earr
                            earr = Excel.Application.Index(arr, iarr, 0)
                            '�p�G���o���O�}�C�]�@���}�C�^
                            If TypeName(earr) = "Variant()" Then
                                '�}�C�ন��r�r��]�H�걵 Join ����k�^�A�@�����Ϊ� key ������
                                key = VBA.Trim(VBA.Join(earr))
                                '�p�Gkey���O�b�ΪŮ� rem �]��subsetExcludingVariantRadicals()���w�q���O�G���}�C�A�S�|���L�D�ܧγ������󤣳B�z�A�ҥH�^�Ǫ��|�]�t�\�h�Ū������]�}�C�j�p�O�T�w�����ܩΤ��i�ܪ��^
                                If key <> " " And key <> "" Then
                                    '�p�G���ۦP�����L�ܧγ���������զ��A
                                    If dictComponents.Exists(key) Then
                                        '���ֹ�G�̳���ƬO�_�@�P
                                        '�����o�w�O�����󱡪����~�r�A�n�ӻP�ثe���~�r�@����ƪ����
                                        wList = components.JoinDictionaryValuesWithComma(dictComponents(key))  '���o�����ܧγ����ᦳ�ۦP����զ����~�r�O����F
                                        '�p�G���h�ӡA�H��1�Ӭ��N��Ӥ���Y�i�C�]�����O�����F�@�˪��ܧγ���
                                        Dim commaPos As Integer
                                        commaPos = VBA.InStr(wList, ",")
                                        w_ofVariantRedical = VBA.IIf(commaPos > 0, VBA.Mid(wList, 1, IIf(commaPos = 0, 0, commaPos - 1)), wList)
                                        
                                        '���G�̳���ƬO�_�@�P�A�p�G�@�P�A
                                        If components.componentsCountofChar(w_ofVariantRedical) = components.componentsCountofChar(w) Then
                                        
                                            '���۴N�n���өҩ������p���ܧγ����O�_�O�P�ݤ@�ӳ����A�O�P�@�������ܧγ������~���N�q�F�Y����k���D��
                                            variantRadical = subsetExcludingVariantRadicalsDict(cln) '���o�{�b�n��諸�~�rw�ҩ������ܧγ����O����
                                            '���o�n��諸�t�ܧγ������~�r�䳡�󤸯������X�C�o�̤���Φr��A�]���|�������ƪ�����A�p�u�ӡB�šv������ӡu�H�v�C�ΰ}�C�S���n�M�������A�G��Collection���X�ӧ@
                                            Set w_ofVariantRedical_componentsCollection = components.DecomposeCollection(w_ofVariantRedical)
                                            Dim i_w_ofVariantRedical_componentsCollection As Byte
                                            '�H�{��������P�n��諸����@���
                                            For Each e_earr In earr
                                                For i_w_ofVariantRedical_componentsCollection = 1 To w_ofVariantRedical_componentsCollection.Count
                                                    If VBA.StrComp(e_earr, w_ofVariantRedical_componentsCollection.Item(i_w_ofVariantRedical_componentsCollection)) = 0 Then
                                                        w_ofVariantRedical_componentsCollection.Remove i_w_ofVariantRedical_componentsCollection '�v�@�����ۦP���A�ѤU���B�ߤ@���N�O�����ܧγ����F
                                                        Exit For
                                                    End If
                                                Next i_w_ofVariantRedical_componentsCollection
                                            Next
                                            '�즹 w_ofVariantRedical_componentsCollection ����u�Ѥ@�Ӥ����F
                                            
                                            Rem �H���o���ܧγ����P�{�b�n�B�z���~�r���ܧγ����@���A�O�P�ݤ@�����~�����B�z
                                            
                                            '���o�w�O�����ܧγ���radical_in_dictComponents�B�β{�b�n�Ӥ�諸�ܧγ���radical_variantRadical���O�O���򳡭�
                                            Rem ���O���ܧγ����~���P����զ����M�Ӧh�A����� Dictionary sameComposeDict �x�s key=���� value= �r �H�K��X�P�����P�Ӧ���ӥH�W���~�r�~��J
                                            'Dim radical_variantRadical As String, radical_in_dictComponents As String
                                            radical_in_dictComponents = dictVariantRadicals(w_ofVariantRedical_componentsCollection.Item(1))
                                            radical_variantRadical = dictVariantRadicals(variantRadical)
                                            Rem �����P���P���o�B�z�A�]�����ܧγ����~�P����զ������M�Ӧh�A�ҥH���Dictionary�x�s��~�rvalue�]�ȡ^�P���k������key�]��^����ȹ�]��-�ȹ�^�A�H�K�d����P�R��
                                            '�p�G�����@�P
                                            If VBA.StrComp(radical_variantRadical, radical_in_dictComponents) = 0 Then
                                                                              
                                                wList = dictComponents(key)(radical_variantRadical) '���o�����ܧγ����ᦳ�ۦP����զ����~�r�r��M��A�ΨӤ��O�_�w���Ӧr�s�b�F�Y���h���A���ƲK�J
                                                If VBA.InStr(wList, w) = 0 Then '�|�L�~�[�J�A�]�������P������զX�n�v�@�P�����w�x�s������դ���ˬd
'                                                    sndPlaySound32 "C:\Windows\Media\Alarm10.wav", 1'�]���Ӧh�A�K�����񴣥ܭ��ĤF
                                                    If wList <> "" Then wList = wList & "," '�H�u,�v�ϧO�U�Ӻ~�r�A�i�ΨӧP�_�O�_����@�Ӻ~�r�C�U�P�C
                                                    dictComponents(key)(radical_variantRadical) = wList & w
                                                End If
                                            
                                            '�p�G�������P�A�h�@�k���A�H���o�C�ӳ������һ⪺�~�r�s�ռƳ��b1�ӥH�W�~�e��resultDict��X
                                            Else
                                                
                                                Rem �ܧγ������P�����ӳ���զ��ۦP�̤Ӧh�A�G���A�[�z��
                                                wList = dictComponents(key)(radical_variantRadical) '���o�����ܧγ����ᦳ�ۦP����զ����~�r�r��M��A�ΨӤ��O�_�w���Ӧr�s�b�F�Y���h���A���ƲK�J
                                                If VBA.InStr(wList, w) = 0 Then
                                                    If wList <> "" Then wList = wList & ","
                                                    dictComponents(key)(radical_variantRadical) = wList & w
                                                End If
'                                                dictComponents(key) = wList & "," & w'�����즡�A���H�r��w�ӫD�r�嫬�O��sameComposeDict���x�s�A�T���ť�
                                            End If '�H���o���ܧγ����P�{�b�n�B�z���~�r���ܧγ����@���A�O�P�ݤ@�����~�B�z�]�~���N�q�^
'                                        Else
'                                            dictComponents(key) = wList & "," & w
                                        End If '�H�W���G�Ӻ~�r��������ƬO�_�@�P�A�p�G�@�P�K�K
                                        
                                    '�p�G�S�����ۦP�����L�ܧγ���������զ��A�h�s�W�O���]�����^��dictComponents���]��c�y�G��G������զ����r��A�ȡG���@�Ө�Ȭ��ŦX����ȳ��󤧺~�r�B�䬰�Ӧr���������r�� sameComposeDict�x�s�^
                                    Else
                                        'dictComponents.Add key, w'��ȥΦr���x�s�Ҧ��ŦX���󪺺~�r���G�A���������ΡA�G���w�אּDictionary sameComposeDict�x�s
'                                        dictComponents(key) = w
                                        sameComposeDict.Add radical_variantRadical, w '�إ߷s������-�~�r�s��-�ȹ�r��H�@��dictComponents��value��
                                        dictComponents.Add key, sameComposeDict '�[�J�s������զ������~�r�s��^�ӷǳƦ^�Ǫ�dictComponents�r�夤
                                        Set sameComposeDict = Nothing '����PCollection��z�]�Ѩ�components�������O�Ҳդ���subsetExcludingVariantRadicals��k�^�A�o�ˤ~��M�ųƤU���ϥΡ]�귽�A�Q�Ρ^�C��
                                    End If
                                End If
                            End If
                        Next iarr
                    End If '�H�W end If: TypeName(arr) = "Variant()" Then
                Next eClnKeysArr
            End If '�H�W end If: Not subsetExcludingVariantRadicalsDict is Nothing Then
        End If
    End If
Next c

Rem ���A�Ψ쪺�ܼ�(�p e_earr�BsameComposeDict�BeClnKeysArr�C�u�n���O�X��Y�i)�b���@�귽�A�Q�ΡA�G���A�ŧi�s�ܼƨӾާ@
'�z��P��������r�H�W������
For Each e_earr In dictComponents.Keys
    If e_earr <> "" Then
        Set sameComposeDict = dictComponents(e_earr)
        For Each eClnKeysArr In sameComposeDict.Keys '���d�C�ӳ���
            If InStr(sameComposeDict(eClnKeysArr), ",") = 0 Then '�p�G�S����Ӻ~�r�H�W�C�H�u,�v�ϧO�U�Ӻ~�r�A�G�i�ΨӧP�_
                sameComposeDict.Remove eClnKeysArr '�N�����ӳ���
            End If
        Next
    End If
Next e_earr
Dim resultDict As New Scripting.Dictionary '�@���̫�Ψӿ�X�d�ﵲ�G���r��]�Y�@��JoinDictionaryValues()���޼ơ^
'�u�ѱa��ӳ����H�W��,�N���^�̫᪺���G�r��
'Set sameComposeDict = Nothing�F�]���U���w�� Set sameComposeDict = dictComponents(e_earr) ���z���A�G���B�K���Υ��M���A�ΡC = �B��l�۷|���¥H���s
For Each e_earr In dictComponents.Keys
    If e_earr <> "" Then
        Set sameComposeDict = dictComponents(e_earr) '���o���G���r��
        For Each eClnKeysArr In sameComposeDict.Keys '���d�C�ӧt��r�H�W������
            resultDict(e_earr) = sameComposeDict(eClnKeysArr) '���o�̫��X�����G
        Next
    End If
Next e_earr


result = components.JoinDictionaryValues(resultDict) '��O�ǤJ dictComponents �@�޼ơA����O�C
Documents.Add().Range.Text = result '�u���@���ܧγ������󤣦P���r
sndPlaySound32 "C:\Windows\Media\Ring05.wav", 1
Rem ����O����
Set subsetExcludingVariantRadicalsDict = Nothing: Set components = Nothing: Set dictComponents = Nothing: Set dictVariantRadicals = Nothing: Set w_ofVariantRedical_componentsCollection = Nothing: Set sameComposeDict = Nothing: Set resultDict = Nothing
Rem �Ʊ�b�@�~�t�οE������Word�����A�Ϩ䦨���̫e�ݪ��]����򦳥Ρ^
Application.ActiveDocument.ActiveWindow.Activate
Application.Activate
End Sub


Sub ��X�Ҧ��u���@�ӳ��󤣦P���rDictionary()
Dim t As Table, d As Document, a As Range, c As Cell, result As String
Dim dictComponents As New Scripting.Dictionary, components As New components, componentsArray() As String, w As String, key As String, arrCom()
Const columnComponents As Byte = 2 '�������
Const columnChar As Byte = 1 '�~�r���

'���o�~�r�Ψ䳡����
Set d = ThisDocument
Set t = d.Tables(1)
For Each c In t.Columns(columnComponents).Cells
    If c.RowIndex > 1 Then '��1�C�O���D��
        '���o�~�r
        Set a = t.Cell(c.RowIndex, columnChar).Range
        w = VBA.Mid(a.Text, 1, Len(a.Text) - 2) '�x�s��аO�D���� chr(13) & chr(7),�n����
        '���o����}�C componentsArray
        components.getComponentsArray componentsArray, c.Range
        '�ƦC����}�C
                Call components.SortStringArray(componentsArray)
                If UBound(componentsArray) > 0 Then
'                    Dim clnComponets As VBA.Collection, eClnComponets
'                    Set clnComponets = components.Subset(componentsArray)
                    Dim arr, earr, iarr As Long
                    arr = components.Subset(componentsArray)
'                    For Each eClnComponets In clnComponets
                    For iarr = LBound(arr) To UBound(arr)
                        earr = Excel.Application.Index(arr, iarr, 0)

'                        earr = arr(iarr)
'                        If Not earr = Empty Then
'                            If TypeName(earr) = "String" Then
'                                key = earr
'                            Else
                                key = VBA.Join(earr)
'                            End If
    '                        key = eClnComponets
                            If dictComponents.Exists(key) Then
                                sndPlaySound32 "C:\Windows\Media\Alarm10.wav", 1
                                If InStr(dictComponents(key), w) = 0 Then
                                    dictComponents(key) = dictComponents(key) & "," & w
                                End If
                            Else
                                'dictComponents.Add key, w
                                dictComponents(key) = w
                            End If
'                        End If
'                    Next eClnComponets
                    Next iarr
                End If
    End If
Next c

result = components.JoinDictionaryValues(dictComponents)
Documents.Add().Range.Text = result '�u���@�ӳ��󤣦P���r
sndPlaySound32 "C:\Windows\Media\Ring05.wav", 1
Application.Activate
End Sub

Sub ����c���r��_5032�r�Ѩ�ӳ���H�W�c����()
Dim t As Table, d As Document, a As Range, c As Cell, result As String
Dim dictComponents As New Scripting.Dictionary, components As New components, componentsArray() As String, w As String, key As String, arrCom()
Const columnComponents As Byte = 2 '�������
Const columnChar As Byte = 1 '�~�r���

'���o�~�r�Ψ䳡����
Set d = ThisDocument
Set t = d.Tables(1)
For Each c In t.Columns(columnComponents).Cells
    If c.RowIndex > 1 Then '��1�C�O���D��
        '���o�~�r
        Set a = t.Cell(c.RowIndex, columnChar).Range
        w = VBA.Mid(a.Text, 1, Len(a.Text) - 2) '�x�s��аO�D���� chr(13) & chr(7),�n����
        '���o����}�C componentsArray
        components.getComponentsArray componentsArray, c.Range
        '�ƦC����}�C
                Call components.SortStringArray(componentsArray)
                If UBound(componentsArray) > 0 Then
'                    Dim clnComponets As VBA.Collection, eClnComponets
'                    Set clnComponets = components.Subset(componentsArray)
                    Dim arr, earr
                    arr = components.Subset(componentsArray)
'                    For Each eClnComponets In clnComponets
                    For Each earr In arr '�o�˷|�C�X�Ҧ��������ȦӤ��O�G���}�C���@���}�C����
                        If Not earr = Empty Then
                            If TypeName(earr) = "String" Then
                                key = earr
                            Else
                                key = VBA.Join(earr)
                            End If
    '                        key = eClnComponets
                            If dictComponents.Exists(key) Then
                                sndPlaySound32 "C:\Windows\Media\Alarm10.wav", 1
                                If InStr(dictComponents(key), w) = 0 Then
                                    dictComponents(key) = dictComponents(key) & "," & w
                                End If
                            Else
                                'dictComponents.Add key, w
                                dictComponents(key) = w
                            End If
                        End If
'                    Next eClnComponets
                    Next earr
                End If
    End If
Next c

result = components.JoinDictionaryValues(dictComponents)
Documents.Add().Range.Text = result
sndPlaySound32 "C:\Windows\Media\Ring05.wav", 1
Application.Activate
End Sub


Sub ��X�Ҧ��㦳�ۦP����զ����rDictionary()
Dim t As Table, d As Document, a As Range, c As Cell, result As String
Dim dictComponents As New Scripting.Dictionary, components As New components, componentsArray() As String, w As String, key As String
Const columnComponents As Byte = 2 '�������
Const columnChar As Byte = 1 '�~�r���

'���o�~�r�Ψ䳡����
Set d = ThisDocument
Set t = d.Tables(1)
For Each c In t.Columns(columnComponents).Cells
    If c.RowIndex > 1 Then '��1�C�O���D��
        '���o�~�r
        Set a = t.Cell(c.RowIndex, columnChar).Range
        w = VBA.Mid(a.Text, 1, Len(a.Text) - 2) '�x�s��аO�D���� chr(13) & chr(7),�n����
        '���o����}�C componentsArray
        components.getComponentsArray componentsArray, c.Range
        
        Call components.SortStringArray(componentsArray)
        key = VBA.Join(componentsArray)
        If dictComponents.Exists(key) Then

            sndPlaySound32 "C:\Windows\Media\Alarm10.wav", 1
            dictComponents(key) = dictComponents(key) & "," & w
        Else
            dictComponents.Add key, w
        End If
    End If
Next c

result = components.JoinDictionaryValues(dictComponents)
Documents.Add().Range.Text = result
sndPlaySound32 "C:\Windows\Media\Ring05.wav", 1

End Sub
Sub ��X�Ҧ��㦳�ۦP����զ����rCollection()
Dim t As Table, d As Document, charCl As New Collection, componentsCl As New Collection, c As Cell, inlsp As InlineShape, i As Long, result As String, cl, e, ee, cll, j As Long, cnt As Byte, flag As Boolean
Dim components As New components
'���o�~�r�Ψ䳡����
Set d = ThisDocument
Set t = d.Tables(1)
For Each c In t.Range.Cells
    If c.RowIndex > 1 Then '��1�C�O���D��
        Select Case c.ColumnIndex
            Case 1 '�~�r
                    charCl.Add VBA.Left(c.Range, 1)
            Case 2 '����
                    components.getComponentsCollection componentsCl, c.Range
                
'                    '�t�m����}�C�H�ƥ[�Jcolleciton �e��
'                    ReDim componets(c.Range.Characters.Count - 2)
'                    '�p�G�S���Ϥ�
'                    If c.Range.InlineShapes.Count = 0 Then
'                        For Each a In c.Range.Characters
'                            ' '�ư��x�s��r��
'                            If InStr(Chr(13) & Chr(7), a) = 0 Then
'                                componets(i) = a.Text
'                                i = i + 1
'                            End If
'                        Next
'                    '�p�G���Ϥ�
'                    Else
'                        For Each a In c.Range.Characters
'                            '�ư��x�s��r��
'                            If InStr(Chr(13) & Chr(7), a) = 0 Then
'                                If a.InlineShapes.Count = 0 Then '�D�Ϥ�
'                                    componets(i) = a.Text
'                                Else '�Ϥ�
'                                    componets(i) = a.InlineShapes(1).AlternativeText
'                                End If
'                                i = i + 1
'                            End If
'                        Next a
'                    End If
'                    componentsCl.Add componets
        End Select
        i = 0
    End If
Next c

reset:
i = 0: j = 0
'��ﳡ��ۦP��
For Each cl In componentsCl
    i = i + 1
    If i > componentsCl.Count Then
        Exit For
    End If
    '���o�����
    cnt = UBound(cl) + 1
        For Each cll In componentsCl
            j = j + 1
            If i <> j Then '����M�ۤv��
                If UBound(cll) + 1 = cnt Then '�p�G����Ƥ@�P�~�i����
'                    For Each e In cl
'                        For Each ee In cll
'                            If VBA.StrComp(ee, e) = 0 Then
'                                flag = True
'                                acl.Remove
'                                Exit For
'                            Else
'                                flag = False
'                            End If
'                        Next ee
'                        If flag = False Then Exit For
'                        flag = False
'                    Next e
                Rem creedit with chatGPT�j����
                    flag = components.CompareArrays(cl, cll)
                End If
                If flag Then
                    VBA.Beep
                    sndPlaySound32 "C:\Windows\Media\Alarm10.wav", 1
                    'If VBA.InStr(result, charCl(j)) = 0 Then '�٨S��쪺�r�~��--����o�˰�
                        result = result & charCl(i) & VBA.vbTab & VBA.Join(componentsCl(i)) & VBA.vbTab & charCl(j) & VBA.vbTab & VBA.Join(componentsCl(j)) & VBA.vbNewLine
                            charCl.Remove j: componentsCl.Remove j
                            flag = False
                        GoTo reset
'                    End If
                End If
            End If
            
        Next cll
        j = 0: flag = False
    
Next cl
Documents.Add().Range.Text = result
'Documents(1).Activate
'Documents(1).Application.Activate
sndPlaySound32 "C:\Windows\Media\Ring05.wav", 1
End Sub
