Attribute VB_Name = "Network"
Option Explicit
Dim DefaultBrowserNameAppActivate As String


Sub �d�߰�y���() '���w��:Ctrl+F12'2010/10/18�׭q
    ''    If ActiveDocument.Path <> "" Then ActiveDocument.Save '��word���ѤF�x�s
    ''    If GetUserAddress = True Then
    '''        MsgBox "���\�����H�W�s���C"
    ''    Else
    ''        MsgBox "�L�k���H�W�s���C"
    ''    End If
    '    Selection.Copy
    '    Shell "W:\!! for hpr\VB\�d�߰�y���\�d�߰�y���\bin\Debug\�d�߰�y���.EXE"
    Const st As String = "C:\Program Files\�]�u�u\�d�߰�y��嵥\"
    Const f As String = "�d�߰�y���.EXE"
    Dim funame As String
    If Selection.Type = wdSelectionNormal Then
        ��r�B�z.ResetSelectionAvoidSymbols
        Selection.Copy
        If Dir(st & f) <> "" Then
            funame = st & f
        ElseIf Dir("C:\Program Files (x86)\�]�u�u\�d�߰�y��嵥\" & f) <> "" Then
            funame = "C:\Program Files (x86)\�]�u�u\�d�߰�y��嵥\" & f
        ElseIf Dir("W:\!! for hpr\VB\�d�߰�y���\�d�߰�y���\bin\Debug\" & f) <> "" Then
            funame = "W:\!! for hpr\VB\�d�߰�y���\�d�߰�y���\bin\Debug\" & f
        ElseIf Dir("C:\�d�߰�y���\�d�߰�y���\bin\Debug\" & f) <> "" Then
            funame = "C:\�d�߰�y���\�d�߰�y���\bin\Debug\" & f
        ElseIf Dir(UserProfilePath & "Dropbox\VS\VB\�d�߰�y���\�d�߰�y���\bin\Debug\" & f) <> "" Then
            funame = UserProfilePath & "Dropbox\VS\VB\�d�߰�y���\�d�߰�y���\bin\Debug\" & f
        ElseIf Dir("A:\", vbVolume) <> "" Then
            If Dir("A:\Users\oscar\Dropbox\VS\VB\�d�߰�y���\�d�߰�y���\bin\Debug\" & f) <> "" Then _
            funame = "A:\Users\oscar\Dropbox\VS\VB\�d�߰�y���\�d�߰�y���\bin\Debug\" & f
        ElseIf Dir(UserProfilePath & "Dropbox\VS\VB\�d�߰�y���\�d�߰�y���\bin\Debug\" & f) <> "" Then
            funame = UserProfilePath & "Dropbox\VS\VB\�d�߰�y���\�d�߰�y���\bin\Debug\" & f
        Else
            Exit Sub
        End If
        Shell funame
    End If
    �d��y���
End Sub

Sub A�t�˺����r���() '���w��:Alt+F12'2010/10/18�׭q
Const f As String = "�t�˺����r���.EXE"
Const st As String = "C:\Program Files\�]�u�u\�t�˺����r���\"
Dim funame As String
If Selection.Type = wdSelectionNormal Then
    ��r�B�z.ResetSelectionAvoidSymbols
    Selection.Copy
    If Dir(st & f) <> "" Then
        funame = st & f
    ElseIf Dir("C:\Users\oscar\Dropbox\VS\VB\�t�˺����r���\�t�˺����r���\bin\Debug\" & f) <> "" Then
        funame = "C:\Users\oscar\Dropbox\VS\VB\�t�˺����r���\�t�˺����r���\bin\Debug\" & f
    ElseIf Dir("C:\Program Files (x86)\�]�u�u\�t�˺����r���\" & f) <> "" Then
        funame = "C:\Program Files (x86)\�]�u�u\�t�˺����r���\" & f
    ElseIf Dir("W:\!! for hpr\VB\�t�˺����r���\�t�˺����r���\bin\Debug\" & f) <> "" Then
        funame = "W:\!! for hpr\VB\�t�˺����r���\�t�˺����r���\bin\Debug\" & f
    ElseIf Dir("C:\�t�˺����r���\�t�˺����r���\bin\Debug\" & f) <> "" Then
        funame = "C:\�t�˺����r���\�t�˺����r���\bin\Debug\" & f
    ElseIf Dir(UserProfilePath & "Dropbox\VS\VB\�t�˺����r���\�t�˺����r���\bin\Debug\" & f) <> "" Then
        funame = "A:\Users\oscar\Dropbox\VS\VB\�t�˺����r���\�t�˺����r���\bin\Debug\" & f
    ElseIf Dir(UserProfilePath & "Dropbox\VS\VB\�t�˺����r���\�t�˺����r���\bin\Debug\" & f) <> "" Then
        funame = "A:\Users\oscar\Dropbox\VS\VB\�t�˺����r���\�t�˺����r���\bin\Debug\" & f
    ElseIf Dir(UserProfilePath & "Dropbox\VS\VB\�t�˺����r���\�t�˺����r���\bin\Debug\" & f) <> "" Then
        funame = UserProfilePath & "Dropbox\VS\VB\�t�˺����r���\�t�˺����r���\bin\Debug\" & f
    Else
        Exit Sub
    End If
    Shell funame
End If

End Sub

Sub �d��y���()
    ��r�B�z.ResetSelectionAvoidSymbols
    SeleniumOP.dictRevisedSearch VBA.Replace(Selection, VBA.Chr(13), "")
End Sub

'Sub �^����y���������}()
'SeleniumOP.grabDictRevisedUrl VBA.Replace(Selection, vba.Chr(13), "")
'End Sub
Sub �dGoogle()
    Rem Alt + g
    ��r�B�z.ResetSelectionAvoidSymbols
    SeleniumOP.GoogleSearch Selection.text
End Sub
Sub �d�ʫ�()
    Rem Alt b
    ��r�B�z.ResetSelectionAvoidSymbols
    SeleniumOP.BaiduSearch Selection.text
End Sub
Sub �d�r�κ�()
    Rem Alt + z
    ��r�B�z.ResetSelectionAvoidSymbols
    If Selection.Characters.Count > 1 Then
        MsgBox "���d1�r", vbExclamation ', vbError
        Exit Sub
    End If
    If Not code.IsChineseCharacter(Selection.text) Then
        MsgBox "������r", vbCritical
        Exit Sub
    End If
    SeleniumOP.LookupZitools Selection.text
End Sub
Sub �d����r�r��()
    Rem Alt + F12
    ��r�B�z.ResetSelectionAvoidSymbols
    If Selection.Characters.Count > 1 Then
        MsgBox "���d1�r", vbExclamation ', vbError
        Exit Sub
    End If
    If Not code.IsChineseCharacter(Selection.text) Then
        MsgBox "������r", vbCritical
        Exit Sub
    End If
    SeleniumOP.LookupDictionary_of_ChineseCharacterVariants Selection.text
End Sub
Sub �d�d���r����W��()
    Rem Ctrl + Alt + x
    ��r�B�z.ResetSelectionAvoidSymbols
    If Selection.Characters.Count > 1 Then
        MsgBox "���d1�r", vbExclamation ', vbError
        Exit Sub
    End If
    If Not code.IsChineseCharacter(Selection.text) Then
        MsgBox "������r", vbCritical
        Exit Sub
    End If
    SeleniumOP.LookupKangxizidian Selection.text
End Sub
Sub �d��y���_������h��()
    Rem Ctrl + Alt + F12
    ��r�B�z.ResetSelectionAvoidSymbols
    SeleniumOP.LookupDictRevised Selection.text
End Sub
Sub �d�Ш|�ʬ�_�Ш|���u�W�r��()
    Rem Ctrl + Shift + Alt + j (j=jiao(��))
    ��r�B�z.ResetSelectionAvoidSymbols
    SeleniumOP.LookupPediaCloudEduTw Selection.text
End Sub
Sub �d�~�y�j����()
    Rem Alt + c
    ��r�B�z.ResetSelectionAvoidSymbols
    If Selection.Characters.Count < 2 Then
        MsgBox "�n2�r�H�W�~���˯��I", vbExclamation ', vbError
        Exit Sub
    End If
    SeleniumOP.LookupHYDCD Selection.text
End Sub
Sub �d�����()
    Rem Ctrl + Shift + Alt + y  y=yun�]���^��y
    ��r�B�z.ResetSelectionAvoidSymbols
    SeleniumOP.LookupYtenx Selection.text
End Sub
Sub �d��Ǥj�v()
    Rem Alt + Shift + g �G�˯��m��Ǥj�v�n�]g=guo ��Ǥj�v����^�F�쬰�GCtrl + d + s �]ds�G�j�v�^
    ��r�B�z.ResetSelectionAvoidSymbols
    SeleniumOP.LookupGXDS Selection.text
End Sub
Sub �d����j���()
    '�m��Ǥj�v�n�Ҧ���M�ұ����m���¥ժ� 20241020
    Rem Ctrl + Alt + Shift + z �]z�G���]zh�^�� z�^
    ��r�B�z.ResetSelectionAvoidSymbols
    SeleniumOP.LookupZWDCD Selection.text
End Sub
Sub �d�~�y�j�r��()
    Rem Alt + Shift + z �]z�G�r�]zi�^�� z�^ 20250108
    Rem Ctrl + Shift + Alt + h (h:�~ (han))
    ��r�B�z.ResetSelectionAvoidSymbols
    SeleniumOP.LookupHYDZD Selection.text
End Sub
Sub �d�j���p��_�V���u��Ѭd��()
    Rem Ctrl + Shift + Alt + U u=xun�]�V�^��u
    ��r�B�z.ResetSelectionAvoidSymbols
    SeleniumOP.LookupBook_Xungu_kaom Selection.text
End Sub
Sub �d�j���p��_�~�y�j����()
    Rem Ctrl + Shift + Alt + i i=ci�]���^��i
    ��r�B�z.ResetSelectionAvoidSymbols
    SeleniumOP.LookupHYDCD_kaom Selection.text
End Sub
Sub �d�ն��`�B�H�a�~�y�j����()
    Rem Ctrl + Shift + Alt + c c=ci�]���^��c
    Rem  Ctrl + Alt + b �]b=bai(�աA�ն��`�B�H�a����)�^
    ��r�B�z.ResetSelectionAvoidSymbols
    LookupHomeinmistsHYDCD Selection.text
End Sub
Sub �d�ն��`�B�H�a�~�y�j�r�����q���˯�()
    Rem Ctrl + Shift + Alt + d d=da�]�j�^��d
    ��r�B�z.ResetSelectionAvoidSymbols
    If Not LookupHomeinmistsHYDZDTextSearch(Selection.text) Then
    End If
End Sub
Sub �d�ն��`�B�H�a����Ѧr�Ϲ��d�\_�ê�k���u��()
    Rem Alt + j �]�Ѧr���ѡ^
    ��r�B�z.ResetSelectionAvoidSymbols
    If Selection.Characters.Count > 1 Then
        MsgBox "���d1�r", vbExclamation ', vbError
        Exit Sub
    End If
    Dim ar 'As Variant
    ar = SeleniumOP.LookupHomeinmistsShuowenImageAccess_VineyardHall(Selection.text)
    If ar(0) = vbNullString Then
        MsgBox "�䤣��A�κ�����F�Χ睊�F�I", vbExclamation
'    Else
'        word.Application.Activate
'        If ar(1) = "" Then MsgBox "��X���G����1���A�Ф�ʦۦ�ާ@�I", vbInformation
    End If
End Sub
Sub �d�ն��`�B�H�a����Ѧr�Ϥ��˯�WFG��_�ѻ��˯�()
    Rem  Alt + shift + s �]���媺���^ Alt + Shift + j �]�Ѧr���ѡ^
    ��r�B�z.ResetSelectionAvoidSymbols
    SeleniumOP.LookupHomeinmistsShuowenImageTextSearchWFG_Interpretation Selection.text
End Sub
Sub �d�ն��`�B�H�a����Ѧr�`()
Rem  Alt + s �]���媺���^
    ��r�B�z.ResetSelectionAvoidSymbols
    If Not SeleniumOP.LookupHomeinmistsShuowenJieZiZhu(Selection.text) Then
        MsgBox "�Э���", vbCritical
    End If
End Sub
Sub �d�~�y�h�\��r�w�è��^�仡�������줧�ȴ��J�ܴ��J�I��m()
    Rem  Alt + n �]n= �� neng�^
    ��r�B�z.ResetSelectionAvoidSymbols
    If Selection.Characters.Count > 1 Then
        MsgBox "���d1�r", vbExclamation ', vbError
        Exit Sub
    End If
    Dim ar 'As Variant
    Dim windowState As word.WdWindowState     '�O�U��Ӫ������Ҧ�
    windowState = word.Application.windowState '�O�U��Ӫ������Ҧ�
    ar = SeleniumOP.LookupMultiFunctionChineseCharacterDatabase(Selection.text)
    If ar(0) = vbNullString Then
        word.Application.Activate
        MsgBox "�䤣��A�κ�����F�Χ睊�F�I", vbExclamation
        With Selection.Application
            .Activate
            With .ActiveWindow
                If .windowState = wdWindowStateMinimize Then
                    .windowState = windowState
                    .Activate
                End If
            End With
        End With
    Else 'ar(0)�����Ů�
        Dim ur As UndoRecord, fontsize As Single
        SystemSetup.stopUndo ur, "�d�~�y�h�\��r�w�è��^�仡�������줧�ȴ��J�ܴ��J�I��m"
        With Selection
            fontsize = VBA.IIf(.font.Size = 9999999, 12, .font.Size) - 4
            If fontsize < 0 Then fontsize = 10
            If .Type = wdSelectionIP Then
                .Move
            Else
                .Collapse wdCollapseEnd
            End If
            '���J���^���m����n���e
            .TypeText "�A�m����n���G�u"
            .InsertAfter ar(0) & "�v" & VBA.Chr(13) 'ar(0)=�m����n���e
            .Collapse wdCollapseEnd
            If Selection.End = Selection.Document.Range.End - 1 Then
                Selection.Document.Range.InsertParagraphAfter
            End If
            .font.Size = fontsize
            .InsertAfter ar(1) '�ӤJ���}
            SystemSetup.contiUndo ur
            .Collapse wdCollapseStart
            With .Application
                .Activate
                With .ActiveWindow
                    If .windowState = wdWindowStateMinimize Then
                        .windowState = windowState
                        .Activate
                    End If
                End With
            End With
        End With
    End If
End Sub
Sub �d����Ѧr�è��^��������κ��}�ȴ��J�ܴ��J�I��m()
    Rem  Alt + o �]o= ����Ѧr ShuoWen.ORG �� O�^
    ��r�B�z.ResetSelectionAvoidSymbols
    If Selection.Characters.Count > 1 Then
        MsgBox "���d1�r", vbExclamation ', vbError
        Exit Sub
    End If
    Dim ar 'As Variant
    Dim windowState As word.WdWindowState      '�O�U��Ӫ������Ҧ�
    windowState = word.Application.windowState '�O�U��Ӫ������Ҧ�
    If Not code.IsChineseCharacter(Selection.text) Then
        MsgBox "������r", vbCritical
        Exit Sub
    End If
    ar = SeleniumOP.LookupShuowenOrg(Selection.text)
    If ar(0) = vbNullString Then
        word.Application.Activate
        MsgBox "�䤣��A�κ�����F�Χ睊�F�I", vbExclamation
        With Selection.Application
            .Activate
            With .ActiveWindow
                If .windowState = wdWindowStateMinimize Then
                    .windowState = windowState
                    .Activate
                End If
            End With
        End With
    Else 'ar(0)�����Ů�
        Dim ur As UndoRecord, fontsize As Single
        SystemSetup.stopUndo ur, "�d����Ѧr�è��^��������κ��}�ȴ��J�ܴ��J�I��m"
        With Selection
            fontsize = VBA.IIf(.font.Size = 9999999, 12, .font.Size) - 4
            If fontsize < 0 Then fontsize = 10
            If .Type = wdSelectionIP Then
                .Move
            Else
                .Collapse wdCollapseEnd
            End If
            .TypeText "�A�m����n���G�u"
            .InsertAfter ar(0) & "�v" & VBA.Chr(13) 'ar(0)=�m����n���e
            .Collapse wdCollapseEnd
            If Selection.End = Selection.Document.Range.End - 1 Then
                Selection.Document.Range.InsertParagraphAfter
            End If
            .font.Size = fontsize
            .InsertAfter ar(1) '���J���}
            SystemSetup.contiUndo ur
            .Collapse wdCollapseStart
            With .Application
                .Activate
                With .ActiveWindow
                    If .windowState = wdWindowStateMinimize Then
                        .windowState = windowState
                        .Activate
                    End If
                End With
            End With
        End With
    End If
End Sub
Sub �d����Ѧr�è��^��������q�`�κ��}�ȴ��J�ܴ��J�I��m()
    Rem  Ctrl+ Shift + Alt + o �]o= ����Ѧr ShuoWen.ORG �� O�^
    ��r�B�z.ResetSelectionAvoidSymbols
    If Selection.Characters.Count > 1 Then
        MsgBox "���d1�r", vbExclamation ', vbError
        Exit Sub
    End If
    If Not code.IsChineseCharacter(Selection.text) Then
        MsgBox "������r", vbCritical
        Exit Sub
    End If
    Dim ar 'As Variant
    Dim windowState As word.WdWindowState      '�O�U��Ӫ������Ҧ�
    windowState = word.Application.windowState '�O�U��Ӫ������Ҧ�
    
    ar = SeleniumOP.LookupShuowenOrg(Selection.text, True)
    
    If ar(0) = vbNullString Then
        word.Application.Activate
        MsgBox "�䤣��A�κ�����F�Χ睊�F�I", vbExclamation
        With Selection.Application
            .Activate
            With .ActiveWindow
                If .windowState = wdWindowStateMinimize Then
                    .windowState = windowState
                    .Activate
                End If
            End With
        End With
    Else 'ar(0)�����Ů�
        Dim ur As UndoRecord, fontsize As Single, st As Long
        SystemSetup.stopUndo ur, "�d����Ѧr�è��^��������κ��}�ȴ��J�ܴ��J�I��m"
        With Selection
            fontsize = VBA.IIf(.font.Size = 9999999, 12, .font.Size) - 4
            If fontsize < 0 Then fontsize = 10
            If .Type = wdSelectionIP Then
                .Move
            Else
                .Collapse wdCollapseEnd
            End If
            st = .start
            .TypeText "�A�m����n���G"
            .InsertAfter ar(0) & VBA.Chr(13) 'ar(0)=�m����n���e
            .Collapse wdCollapseEnd
            If Selection.End = Selection.Document.Range.End - 1 Then
                Selection.Document.Range.InsertParagraphAfter
            End If
            If ar(2) <> vbNullString Then
                '���J�q�`���e
                .InsertAfter "�q�`���G" & VBA.IIf(VBA.Asc(VBA.Left(ar(2), 1)) = 13, vbNullString, VBA.Chr(13)) & ar(2) & VBA.Chr(13)
                Dim p As Paragraph, s As Byte, sDuan As Byte
                s = VBA.Len("                                ") '�q�`��������
                sDuan = VBA.Len("                ") '�q�`�����q�`��
                .Paragraphs(1).Range.font.Bold = True '����G "�q�`���G"
reCheck:
                For Each p In .Paragraphs
                    If VBA.InStr(p.Range.text, "�M�N �q�ɵ��m����Ѧr�`�n") Then
                        p.Range.Delete
                        GoTo reCheck:
                    ElseIf VBA.Replace(p.Range.text, " ", "") = VBA.Chr(13) Then
                        p.Range.Delete
                        GoTo reCheck:
                    ElseIf VBA.Left(p.Range.text, s) = VBA.space(s) Then '�q�`��������
                        p.Range.text = VBA.Mid(p.Range.text, s + 1)
                    ElseIf VBA.Left(p.Range.text, sDuan) = VBA.space(sDuan) Then '�q�`�����q�`��
                        With p.Range
                            .text = VBA.Mid(p.Range.text, sDuan + 1)
                            With .font
                                .Size = fontsize + 2
                                .ColorIndex = 11 '.Font.Color= 34816
                            End With
                        End With
                    End If
                Next p
                .Collapse wdCollapseEnd
            End If
                        
            '���}�榡�]�w
            Ū�J������ƫ�_����ӤJ���}�γ]�w�榡 .Range, VBA.CStr(ar(1)), fontsize
'            .font.Size = fontsize
'            .InsertAfter ar(1) '���J���}
'            .Collapse wdCollapseStart
            SystemSetup.contiUndo ur
            Ū�J������ƫ�_�٭�������A .Application.ActiveWindow, windowState
'            With .Application
'                .Activate
'                With .ActiveWindow
'                    If .windowState = wdWindowStateMinimize Then
'                        VBA.Interaction.DoEvents
'                        .windowState = windowState
'                        .Activate
'                        VBA.Interaction.DoEvents
'                    End If
'                End With
'            End With
            .SetRange st, st
        End With
    End If
End Sub
Sub �d����r�r��è��^�仡���������κ��}�ȴ��J�ܴ��J�I��m()
    Rem  Alt + v �]v= ����r variants �� v�^
    ��r�B�z.ResetSelectionAvoidSymbols
    If Selection.Characters.Count > 1 Then
        MsgBox "���d1�r", vbExclamation ', vbError
        Exit Sub
    End If
    If Not code.IsChineseCharacter(Selection.text) Then
        MsgBox "������r", vbCritical
        Exit Sub
    End If
    'ar(1) as String
    Dim ar As Variant, x As String, windowState As word.WdWindowState     '�O�U��Ӫ������Ҧ�

    x = Selection.text
    windowState = word.Application.windowState '�O�U��Ӫ������Ҧ�

    ar = SeleniumOP.LookupDictionary_of_ChineseCharacterVariants_RetrieveShuoWenData(x)
    
    If ar(0) = vbNullString Then
        word.Application.Activate
        MsgBox "�䤣��A�κ�����F�Χ睊�F�I", vbExclamation
        With Selection.Application
            .Activate
            With .ActiveWindow
                If .windowState = wdWindowStateMinimize Then
                    .windowState = windowState
                    .Activate
                End If
            End With
        End With
    Else '�p�Gar(0)�D�Ŧr��]�ŭȡ^
        Dim ur As UndoRecord, fontsize As Single, st As Long ', ed As Long
        SystemSetup.stopUndo ur, "�d����r�r��è��^�仡���������κ��}�ȴ��J�ܴ��J�I��m"
        With Selection
            st = .start
            fontsize = VBA.IIf(.font.Size = 9999999, 12, .font.Size) - 4
            If fontsize < 0 Then fontsize = 10
            If .Type = wdSelectionIP Then
                .Move
            Else
                .Collapse wdCollapseEnd
            End If
            Dim s As Byte
            s = VBA.InStr(ar(0), "�m����n�����C")
            If s = 0 Then
                If ar(0) = "�������ΨS����ơI" Then
                    .TypeText VBA.Chr(13)
                Else
                    .TypeText "�A�m����n�G" & VBA.Chr(13)
                End If
            Else
                 .TypeText "�A" & VBA.Mid(ar(0), s) & VBA.Chr(13)
            End If
            Dim shuowen As String
            shuowen = VBA.Replace(VBA.Replace(ar(0), "�G�A", "�G" & x & "�A"), "�q�`���G", VBA.Chr(13) & "�q�`���G")
            If VBA.Left(shuowen, 1) = "�A" Then
                shuowen = x & shuowen
            End If
            If s = 0 And ar(0) <> "�������ΨS����ơI" Then
                If VBA.InStr(shuowen, "<img ") Then
                    word.Application.ScreenUpdating = False
                    Dim rngHtml As Range
                    Set rngHtml = .Document.Range(.start, .start)
                    '.TypeText shuoWen & VBA.Chr(13)
                    '�r�Ӫ���TypeText�|��������,�|�L��
                    .text = shuowen & VBA.Chr(13)
                    'ed = Selection.Range.End '���J��r��A�YSelection���ܫ�A �� With �϶�����ήɤ����I20221010
                    'Set rngHtml = .Document.Range(st, ed)
                    rngHtml.End = Selection.End
                    Ū�J������ƫ�_�٭�������A .Application.ActiveWindow, windowState
                    
                    'InnerHTML_Convert_to_WordDocumentContent Selection.Range ', vbNullString
                    innerHTML_Convert_to_WordDocumentContent rngHtml ', vbNullString
                    rngHtml.Collapse wdCollapseEnd
                    rngHtml.Select
                Else
                    .InsertAfter shuowen & VBA.Chr(13)  'ar(0)=�m����n���e
                    .Collapse wdCollapseEnd
                End If
            End If
            
            If Selection.End = Selection.Document.Range.End - 1 Then
                Selection.Document.Range.InsertParagraphAfter
            End If
            Ū�J������ƫ�_����ӤJ���}�γ]�w�榡 .Range, VBA.CStr(ar(1)), fontsize
            Ū�J������ƫ�_�٭�������A .Application.ActiveWindow, windowState
            SystemSetup.contiUndo ur
            word.Application.ScreenUpdating = True
            .SetRange st, st
        End With
    End If
End Sub
Rem 1.���w���W�A�ާ@ 20241004 Alt + Shift + y (y:��) �C2.�Y��ЩҦb���m���Ǻ��n�����}�A�h�N�䤺�eŪ�J����]��ӳs���q���ᴡ�J�^
Sub �d���Ǻ����g�P�������w���W�奻_�è��^��¤�r�Ȥκ��}�ȴ��J�ܴ��J�I��m()
    SystemSetup.playSound 0.484
    Dim linkInput As Boolean, rngLink As Range, rngHtml As Range, ss As Long, x As String, gua As String, e, arrYi
    If Selection.Type = wdSelectionIP Then
        '�Y��ЩҦb���m���Ǻ��n�����}�A�h�N�䤺�eŪ�J����
        Set rngLink = Selection.Range
        If rngLink.End + 1 = rngLink.Paragraphs(1).Range.End Then GoTo previousLink
        If rngLink.Hyperlinks.Count = 1 Then
            If VBA.Left(rngLink.Hyperlinks(1).Address, VBA.Len("https://www.eee-learning.com/")) = "https://www.eee-learning.com/" Then
                linkInput = True
            End If
        Else
            Set rngLink = Selection.Range.Next
            If Not rngLink Is Nothing Then
                If rngLink.Hyperlinks.Count = 1 Then
                    If VBA.Left(rngLink.Hyperlinks(1).Address, VBA.Len("https://www.eee-learning.com/")) = "https://www.eee-learning.com/" Then
                        linkInput = True
                    End If
                Else
previousLink:
                    Set rngLink = Selection.Range.Previous
                    If Not rngLink Is Nothing Then
                        If rngLink.Hyperlinks.Count = 1 Then
                            If VBA.Left(rngLink.Hyperlinks(1).Address, VBA.Len("https://www.eee-learning.com/")) = "https://www.eee-learning.com/" Then
                                linkInput = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        ��r�B�z.ResetSelectionAvoidSymbols
                
        If Selection.Characters.Count > 3 Then
errExit:
            word.Application.Activate
            VBA.MsgBox "���W���~! �Э��s����C", vbExclamation
            Exit Sub
        End If
    End If
    
    Dim ur As UndoRecord, s As Long, ed As Long
    Dim fontsize As Single, rngBooks As Range
    fontsize = VBA.IIf(Selection.font.Size = 9999999, 12, Selection.font.Size) * 0.6
    If fontsize < 0 Then fontsize = 10

    SystemSetup.stopUndo ur, "�d���Ǻ����g�P�������w���W�奻_�è��^��¤�r�Ȥκ��}�ȴ��J�ܴ��J�I��m"
    word.Application.ScreenUpdating = False
    Dim windowState As word.WdWindowState      '�O�U��Ӫ������Ҧ�
    windowState = word.Application.windowState '�O�U��Ӫ������Ҧ�
    If Selection.Document.path <> vbNullString And Selection.Document.Saved = False Then Selection.Document.Save
    
    s = Selection.start: ss = s
    
    Dim result(1) As String, iwe As SeleniumBasic.IWebElement

    If linkInput Then
        If SeleniumOP.OpenChrome(rngLink.Hyperlinks(1).Address) Then
            Set iwe = WD.FindElementByCssSelector("#block-bartik-content > div > article > div > div.clearfix.text-formatted.field.field--name-body.field--type-text-with-summary.field--label-hidden.field__item")
            If Not iwe Is Nothing Then
                Selection.MoveUntil Chr(13)
                Selection.TypeText Chr(13)
                Selection.Style = word.wdStyleNormal '"����"
                Selection.ClearFormatting
                Selection.Collapse wdCollapseEnd
                s = Selection.start
                Rem �{�b�@�ߥHhtml�����e�A�|���u�]�ר�b�Ѯ榡��{�W�A�i�H�X�G��������������˦��^ 20241015
'                If SeleniumOP.IslinkImageIncluded���e�����]�t�W�s���ιϤ�(iwe) Then '���Ϥ��ɨ� "innerHTML" �ݩʭ�
'                    Dim Links() As SeleniumBasic.IWebElement, images() As SeleniumBasic.IWebElement
'                    Links = SeleniumOP.Links
'                    images = SeleniumOP.images
                    Selection.InsertAfter iwe.GetAttribute("innerHTML")
                    ed = Selection.End
                                        
                    Set rngHtml = Selection.Document.Range(s, ed)
                    
                    
                    
                    HTML2Doc.innerHTML_Convert_to_WordDocumentContent rngHtml, "https://www.eee-learning.com"
                    'SeleniumOP.inputElementContentAll���J��������Ҧ������e iwe
                    
                    
                    

'                    Stop 'just for test
                    GoTo finish 'just for test

                Rem �{�b�@�ߥHhtml�����e�A�|���u�]�ר�b�Ѯ榡��{�W�A�i�H�X�G��������������˦��^ 20241015
'                Else '�S���Ϥ��ɨ� "textContent" �ݩʭ�
'                    result(0) = iwe.GetAttribute("textContent")
'                    result(1) = rngLink.Hyperlinks(1).Address
'                End If Rem �{�b�@�ߥHhtml�����e�A�|���u�]�ר�b�Ѯ榡��{�W�A�i�H�X�G��������������˦��^ 20241015
                
                Rem �{�b�@�ߥHhtml�����e�A�|���u�]�ר�b�Ѯ榡��{�W�A�i�H�X�G��������������˦��^ 20241015
'                GoTo insertText:
            Else 'If iwe Is Nothing Then
                MsgBox "�������c���P�A�S���һݤ��e�A�i��O�s���ܦU���`���W�s���A�i�жK�W�t���W�s�����U���`���D�Τ�r��A�~��j�C�P���P���@�n�L��������@�g���D", vbExclamation
                GoTo finish
            End If
        End If
    Else 'If Not linkInput Then
        Rem �۰��ˬd�U�@�Ӧr
            Selection.Collapse wdCollapseStart
'        If Selection.Characters.Count = 1 Then
            x = Selection.text
            If Not Keywords.�P�����W_����_����.Exists(x) And Not Keywords.���ǲ���r��.Exists(x) Then
                x = Selection.Characters(1).text & Selection.Characters(1).Next(wdCharacter, 1).text
                If Keywords.�P�����W_����_����.Exists(x) Or Keywords.���ǲ���r��.Exists(x) Then
                    Selection.MoveRight wdCharacter, 2, wdExtend
                Else
                    arrYi = Array("�ߧ�", "ô��", "����", "�Ǩ�", "����")
                    If VBA.IsArray(VBA.Filter(arrYi, x)) Then
                        If (UBound(VBA.Filter(arrYi, x)) > -1) Then
                            Select Case x
                                Case "�ߧ�"
                                    gua = "29"
                                    GoTo grab:
                                Case "ô��"
                                    Select Case Selection.Next(wdCharacter, 2).text
                                        Case "�W"
                                            gua = "65"
                                            GoTo grab:
                                        Case "�U"
                                            gua = "66"
                                            GoTo grab:
                                    End Select
                                Case "����"
                                    gua = "67" '"������"
                                    GoTo grab:
                                Case "�Ǩ�"
                                    gua = "68" '"�Ǩ���"
                                    GoTo grab:
                                Case "����"
                                    gua = "69" '"������"
                                    GoTo grab:
                            End Select
                        End If
                    End If
                End If
            End If
'        End If
        gua = Selection.text
        If Selection.Characters.Count = 2 Then
            If Selection = "�ߧ�" Then
                If Selection.Characters(2) = "��" Then
                    gua = "��"
                End If
            End If
        End If
        
        On Error GoTo eH:
        If Keywords.�P�����W_����_����.Exists(gua) = False Then
            If Keywords.���ǲ���r��.Exists(gua) = False Then
                GoTo errExit
            Else
                gua = Keywords.���ǲ���r��(gua)
            End If
        End If
    End If
    '�H�W���b�ˬd
    
    '�H�U���ˬd�q�L��
    
    gua = Keywords.�P�����W_����_����(gua)(1)

grab:
    If Not SeleniumOP.GrabEeeLearning_IChing_ZhouYi_originalText(gua, result, iwe) Then
        word.Application.Activate
        VBA.MsgBox "�䤣��A�κ�����F�α��F�K�K", vbInformation
        Exit Sub
    End If
    
    
'        Set iwe = WD.FindElementByCssSelector("#block-bartik-content > div > article > div > div.clearfix.text-formatted.field.field--name-body.field--type-text-with-summary.field--label-hidden.field__item")
'        If Not iwe Is Nothing Then
            If SeleniumOP.IslinkImageIncluded���e�����]�t�W�s���ιϤ�(iwe) Then '���Ϥ��ɨ� "innerHTML" �ݩʭ�
            'If SeleniumOP.IsImageIncluded���e�����]�t�Ϥ�(iwe) Then '���Ϥ��ɨ� "innerHTML" �ݩʭ�
                If Selection.Style <> word.wdStyleNormal Then
                    Selection.MoveUntil Chr(13)
                    Selection.TypeText Chr(13)
                    Selection.Style = word.wdStyleNormal '"����"
                    Selection.Collapse wdCollapseEnd
                End If
                
                s = Selection.start
                
                Selection.InsertAfter iwe.GetAttribute("innerHTML")
                ed = Selection.End
                
                
                
                Set rngHtml = Selection.Document.Range(s, ed)
                
                HTML2Doc.innerHTML_Convert_to_WordDocumentContent rngHtml, "https://www.eee-learning.com"
                'SeleniumOP.inputElementContentAll���J��������Ҧ������e iwe
                
                
'                Dim refs As Boolean
'                Rem ���N�`���G or �U�a���ѡG
'                Set rngBooks = Selection.Document.Range(rngHtml.start, rngHtml.End)
'                rngBooks.Find.ClearFormatting
'                If rngBooks.Find.Execute("���N�`���G", , , , , , True, wdFindStop) Then
'                    refs = True
'                Else
'                    rngBooks.SetRange rngHtml.start, rngHtml.End
'                    If rngBooks.Find.Execute("�U�a���ѡG", , , , , , True, wdFindStop) Then
'                        refs = True
'                    End If
'                End If
'                If refs Then
'                    With rngBooks '"���N�`���G"�Ҧb�q���d��
'                        .Style = wdStyleHeading1
'                        .font.Size = 22
'                        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle '��涡�Z
'                    End With
'                    ed = Ū�J������ƫ�_����ӤJ���}�γ]�w�榡(rngHtml, result(1), fontsize)
'                    Ū�J������ƫ�_�٭�������A Selection.Document.ActiveWindow, windowState
'                End If

        '                    Stop 'just for test
                GoTo finish 'just for test
        
        
        '    Else '�S���Ϥ��ɨ� "textContent" �ݩʭ�
        '        result(0) = iwe.GetAttribute("textContent")
        '        result(1) = rngLink.Hyperlinks(1).Address
            End If
'        End If
    
insertText:
    word.Application.Activate
    If VBA.InStr(result(0), "���N�`���G") Then
        If VBA.vbOK = MsgBox("�O�_�M���u���N�`���G�v�H�᪺��r�H", VBA.vbQuestion + VBA.vbOKCancel) Then
            result(0) = VBA.Left(result(0), VBA.InStr(result(0), "���N�`���G") - 1)
        End If
    End If
        
    Dim p As Paragraph, book As String, iwes() As IWebElement
    
    With Selection
'        fontsize = VBA.IIf(.font.Size = 9999999, 12, .font.Size) - 4
'        If fontsize < 0 Then fontsize = 10
        If .Type = wdSelectionIP And .text <> Chr(13) Then
            .Delete
        End If
        s = .start
        .TypeText VBA.Replace(result(0), ChrW(160), vbNullString)
        
        ed = Selection.End 'Selection���ܰʫ�A���G�� With�϶��L�k�ήɦ^���A�n���n�Υ��L�u�A���I�s�v�~��ϬM�ήɯu�� 20241009
        Set rngBooks = .Document.Range(s, ed)
        ��r�B�z.FixFontname rngBooks
        
        '�]�w����r
        iwes = WD.FindElementsByTagName("STRONG")
        For Each e In iwes
            Set iwe = e
            x = iwe.text '.GetAttribute("textContent")
            Do While rngBooks.Find.Execute(x, , , , , , , wdFindStop)
                Set p = rngBooks.Paragraphs(1)
                If p.Range.text = x & Chr(13) Then
                    p.Range.font.Bold = True
                    Exit Do
                End If
            Loop
            rngBooks.SetRange s, ed
        Next e
        
        rngBooks.SetRange s, ed
        '*�p�`�r
        For Each p In rngBooks.Paragraphs
            If VBA.Left(p.Range.text, 1) = "*" Then
                With p.Range.font
                    .Size = .Size - 2
                    .ColorIndex = 11 '.Font.Color= 34816
                End With
            ElseIf VBA.InStr(p.Range.text, "*") Then
                With p.Range.Find
                    .ClearFormatting
                    With .Replacement
                        .font.ColorIndex = 11 '.Font.Color= 34816
                        .font.Bold = True
                    End With
                    .Execute "*", , , , , , , , , "*", wdReplaceAll
                    .ClearFormatting
                End With
            End If
        Next p
        
        ed = Ū�J������ƫ�_����ӤJ���}�γ]�w�榡(Selection.Range, result(1), fontsize)
'        .Document.Range(s, s).Select
'        Ū�J������ƫ�_�٭�������A .Document.ActiveWindow, windowState
        
    End With
    
    '�O�d���N�`���Ψ�W�s��
    Set rngBooks = Selection.Document.Range
    If VBA.InStr(result(0), "���N�`���G") Then
        With rngBooks
            With .Find
                .ClearFormatting
                If .Execute("���N�`���G", , , , , , True, wdFindStop) Then
                    With rngBooks '"���N�`���G"�Ҧb�q���d��
                        .Style = wdStyleHeading1
                        .font.Size = 22
                        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle '��涡�Z
                    End With
                    rngBooks.End = ed '�]�w���������N�`�����d��
                    For Each p In rngBooks.Paragraphs
                        If p.Range.text <> Chr(13) And VBA.Left(p.Range.text, 4) <> "http" And VBA.Left(p.Range.text, 5) <> "���N�`���G" Then
                            book = VBA.Left(p.Range.text, VBA.Len(p.Range.text) - 1)
                            Set iwe = SeleniumOP.WD.FindElementByLinkText(book)
                            If Not iwe Is Nothing Then
                                Set rngLink = p.Range.Document.Range(p.Range.start, p.Range.End - 1)
                                With rngLink
                                    .Hyperlinks.Add rngLink, iwe.GetAttribute("href")
                                    .Style = wdStyleHeading2 '���D 2
                                    .font.Size = 18
                                    .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle '��涡�Z
                                End With
                            End If
                        End If
                    Next p
                End If
            End With
        End With
    End If
    
    '�\Ū�j��    '�@ �\Ū�j�Ϯ�
    Dim chkBook As String
    If VBA.InStr(result(0), "�\Ū�j��") Then
        chkBook = "�\Ū�j��,"
    ElseIf VBA.InStr(result(0), "�\Ū�j�Ϯ�") Then
        chkBook = "�\Ū�j�Ϯ�"
    Else
'        Stop 'for check
        GoTo finish
    End If
    If chkBook <> VBA.vbNullString Then
        Set rngBooks = Selection.Document.Range(s, ed)
        With rngBooks.Find
            .ClearFormatting
            If .Execute(chkBook, , , , , , True, wdFindStop) Then
'                rngBooks.Paragraphs(1).Range.text = "�@" & chkBook & Chr(13)
'                Set iwe = WD.FindElementByPartialLinkText("�\Ū�j��")
'                If Not iwe Is Nothing Then
                    rngBooks.SetRange rngBooks.Paragraphs(1).Range.start, rngBooks.Paragraphs(1).Range.End - 1
                    rngBooks.Hyperlinks.Add rngBooks, result(1) 'iwe.GetAttribute("href")
'                End If
            End If
        End With
    End If

finish:
    If linkInput And Not rngHtml Is Nothing Then
        Dim refs As Boolean
        Rem ���N�`���G or �U�a���ѡG
        Set rngBooks = Selection.Document.Range(rngHtml.start, rngHtml.End)
        rngBooks.Find.ClearFormatting
        If rngBooks.Find.Execute("���N�`���G", , , , , , True, wdFindStop) Then
            refs = True
        Else
            rngBooks.SetRange rngHtml.start, rngHtml.End
            If rngBooks.Find.Execute("�U�a���ѡG", , , , , , True, wdFindStop) Then
                refs = True
            End If
        End If
        If refs Then
            With rngBooks '"���N�`���G"�Ҧb�q���d��
                .Style = wdStyleHeading1
                .font.Size = 22
                .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle '��涡�Z
            End With
            rngBooks.SetRange rngBooks.Paragraphs(1).Range.End + 1, rngHtml.End '�]�w���������N�`�����d��
            For Each p In rngBooks.Paragraphs
                If p.Range.text <> Chr(13) And VBA.Left(p.Range.text, 4) <> "http" And VBA.Left(p.Range.text, 5) <> "���N�`���G" Then
                    If p.Range.Hyperlinks.Count = 0 Then
                        book = VBA.Left(p.Range.text, VBA.Len(p.Range.text) - 1)
                        Set iwe = SeleniumOP.WD.FindElementByLinkText(book)
                        If Not iwe Is Nothing Then
                            Set rngLink = p.Range.Document.Range(p.Range.start, p.Range.End - 1)
                            With rngLink
                                .Hyperlinks.Add rngLink, iwe.GetAttribute("href")
                                .Style = wdStyleHeading2 '���D 2
                                .font.Size = 18
                                .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle '��涡�Z
                            End With
                        End If
                    Else
                        If p.Range.text <> Chr(13) And p.Style <> "���D 2" And p.Style <> wdStyleHeading2 Or p.Style = "�M��q��" Or p.Style = wdStyleListParagraph Then
                            p.Style = wdStyleHeading2 '���D 2
                            p.Range.font.Size = 18
                            p.Range.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle '��涡�Z
                        End If
                    End If
                End If
            Next p
            
        End If
        If result(1) <> vbNullString Then
            ed = Ū�J������ƫ�_����ӤJ���}�γ]�w�榡(rngHtml, result(1), fontsize)
        End If
        Ū�J������ƫ�_�٭�������A Selection.Document.ActiveWindow, windowState
        
        Rem just for check
        With rngHtml.Find
            .ClearFormatting
            .MatchWholeWord = True
            If .Execute("[<>&;]", , , True) Then
                rngHtml.Select
                SystemSetup.playSound 3
            End If
            .MatchWholeWord = False
            .ClearFormatting
        End With
        Rem just for check
    Else 'If linkInput=false Then
        If result(1) <> vbNullString Then
            ed = Ū�J������ƫ�_����ӤJ���}�γ]�w�榡(rngHtml, result(1), fontsize)
        End If
    End If 'If linkInput Then

    word.Application.ScreenUpdating = True
    Selection.Document.Range(ss, ss).Select
    Ū�J������ƫ�_�٭�������A Selection.Document.ActiveWindow, windowState
    SystemSetup.contiUndo ur
    playSound 2
    Exit Sub

eH:
    Select Case Err.Number
        Case Else
            Debug.Print Err.Number & Err.Description
            Stop 'just for test
            Resume
    End Select
End Sub
Rem �ѫ��w�ֳt��� Ctrl + Shift + Alt + x�]x:�p(xiao)��x)
Sub �d�p�ǰ�W�j����Ū�J�䵲�GSub()
    �d�p�ǰ�W�j����Ū�J�䵲�G
End Sub
Rem �d�p�ǰ�W�j�� �]��@�ɮ��ٷ|�եΡA�i�H�P�_���榨�\�P�_�C���ѶǦ^false
Function �d�p�ǰ�W�j����Ū�J�䵲�G() As Boolean
    
    'https://xiaoxue.iis.sinica.edu.tw/shangguyin/
    '#PageResult
    'innerHTML
    'outerHTML
    Dim rng As Range, selInfo As New StringInfo, w As String, innerHtml As String, p As Paragraph, iwe As SeleniumBasic.IWebElement
    Dim ur As UndoRecord, s As Long, e As Long, result As Boolean
    w = Selection.text
    If Selection.Type <> wdSelectionIP Then
        selInfo.Create w
        If selInfo.LengthInTextElements > 1 Or Not code.IsChineseCharacter(w) Then
            MsgBox "�u��d��1�Ӻ~�r", vbCritical
            Exit Function
        End If
        s = Selection.Characters(1).End
    Else
        If Selection.End = Selection.Document.Range.End - 1 Then
            Selection.MoveLeft wdCharacter, 1, wdMove
        ElseIf Selection.text = VBA.Chr(13) Then
            Selection.MoveLeft wdCharacter, 1, wdMove
        End If
        w = Selection.text
        If Not code.IsChineseCharacter(w) Then
            MsgBox "�u��d��1�Ӻ~�r", vbCritical
            Exit Function
        End If
        If Not Selection.Next(wdCharacter, 1) Is Nothing Then
            s = Selection.Next(wdCharacter, 1).start
        Else
            s = Selection.Document.Range.End - 1
        End If
    End If
    
    SystemSetup.stopUndo ur, "�d�p�ǰ�W�j����Ū�J�䵲�G"
    word.Application.ScreenUpdating = False

    On Error GoTo eH:

    Set rng = Selection.Document.Range(s, s)
    
    innerHtml = SeleniumOP.GrabXiaoxueShangGuYin(w)
�S:
    If innerHtml = vbNullString Then
        VBA.DoEvents
        word.Application.Activate
        MsgBox "�d�L��ơF�� �Э��աI", vbCritical
        GoTo finish
    Else
         Rem �]�t �����s�� �H�U������
         'innerHtml = VBA.Replace(innerHtml, "</form>" & VBA.Chr(13), vbNullString)
         Rem �簣 �����s�� �H�U������
         innerHtml = VBA.Mid(innerHtml, 1, VBA.InStr(innerHtml, "</form>" & VBA.Chr(13)) - 1)
         
    End If

    e = VBA.InStr(innerHtml, "�]�d�L��ơ^")
    Do While e
        innerHtml = VBA.Mid(innerHtml, 1, VBA.InStrRev(innerHtml, "<table", e) - 1) & "�d�L���" & VBA.Mid(innerHtml, VBA.InStr(e, innerHtml, "</table>") + VBA.Len("</table>"))
        e = VBA.InStr(innerHtml, "�]�d�L��ơ^")
    Loop
    innerHtml = VBA.Replace(innerHtml, "�d�L���", "�]�d�L��ơ^")
    e = VBA.InStr(innerHtml, "�]��ƼW�ɤ��^")
    Do While e
        innerHtml = VBA.Mid(innerHtml, 1, VBA.InStrRev(innerHtml, "<table", e) - 1) & "��ƼW�ɤ�" & VBA.Mid(innerHtml, VBA.InStr(e, innerHtml, "</table>") + VBA.Len("</table>"))
        e = VBA.InStr(innerHtml, "�]��ƼW�ɤ��^")
    Loop
    innerHtml = VBA.Replace(innerHtml, "��ƼW�ɤ�", "�]��ƼW�ɤ��^")
    e = VBA.InStr(innerHtml, "colspan")
    Do While e
        innerHtml = VBA.Mid(innerHtml, 1, VBA.InStrRev(innerHtml, "<tr>", e) - 1) & VBA.Mid(innerHtml, VBA.InStr(e, innerHtml, "</tr>") + VBA.Len("</tr>"))
        e = VBA.InStr(innerHtml, "colspan")
    Loop
    
    rng.InsertParagraphAfter
    rng.Collapse wdCollapseEnd
    If Not IsStyleExists("�p�ǰ�W�j�����", rng.Document) Then
        Dim styl As word.Style
        Set styl = rng.Document.Styles.Add("�p�ǰ�W�j�����", wdStyleTypeParagraph)
        With styl
            .font.Name = "Lucida Sans Unicode"
            .Priority = 13
        End With
    End If
    rng.ParagraphFormat.Style = "�p�ǰ�W�j�����"

    rng.text = innerHtml
    e = rng.End
    
    rng.Find.Execute ">���j��<"
    rng.SetRange rng.Paragraphs(1).Range.start, e
    rng.Document.Range(s + 1, rng.start).text = vbNullString
    
    rng.Find.Execute "        ", , , , , , , , , vbNullString, wdReplaceAll
    rng.Find.Execute "      ", , , , , , , , , vbNullString, wdReplaceAll
    rng.Find.Execute "    ", , , , , , , , , vbNullString, wdReplaceAll
    s = rng.start: e = rng.End
    Do While rng.Find.Execute("MessageTitle")
        rng.ParagraphFormat.Style = wdStyleHeading6
    Loop
    rng.SetRange s, e
    
    HTML2Doc.innerHTML_Convert_to_WordDocumentContent rng, vbNullString ', "Lucida Sans Unicode"'��μ˦��ӱ���
    
    
    '�M���Ŧ�/�q
    Do While rng.Document.Range(rng.start, rng.End).Find.Execute("^p^p")
        rng.Find.Execute "^p^p", , , , , , , , , "^p", wdReplaceAll
    Loop
    
    For Each p In rng.Paragraphs
        If p.Range.text = VBA.Chr(13) Or p.Range.text = " " & VBA.Chr(13) Then
            p.Range.text = vbNullString
        End If
    Next p
    '�H�W'�M���Ŧ�/�q
    
    Dim tb As word.table, c As word.cell, a As Range
    For Each tb In rng.tables
'        tb.AllowAutoFit = True
'        tb.AutoFitBehavior wdAutoFitContent
        
        For Each c In tb.Range.cells
            For Each a In c.Range.Characters
                Select Case a.text
                    Case VBA.ChrW(232), VBA.ChrW(566), VBA.ChrW(545), VBA.ChrW(224)
                        If a.font.Name <> "Lucida Sans Unicode" Then
                            a.font.Name = "Lucida Sans Unicode"
                        End If
                        If a.font.Name <> "Lucida Sans Unicode" Then
                            a.Select
                            MsgBox "�Ф�ʳ]�w�r���� ""Lucida Sans Unicode""", vbExclamation
                        End If
                    Case Else
                        If a.font.Name = "Calibri" Then
                            a.font.Name = "Lucida Sans Unicode"
                            If a.font.Name = "Calibri" Then
                                a.Select
                                Stop
                            End If
                        End If
                End Select
            Next a
        Next c

        With tb
            If .Range.Style Is Nothing Then
                .Range.Style = "�p�ǰ�W�j�����"
            ElseIf .Range.Style <> "�p�ǰ�W�j�����" Then
                .Range.Style = "�p�ǰ�W�j�����"
            End If
            '.Range.font.Name = "Times New Roman"
            '.Range.font.Name = "Lucida Sans Unicode"
            '.Range.font.NameOther = "Lucida Sans Unicode"
            '.Range.font.NameAscii = "Lucida Sans Unicode"
'            .Range.font.NameFarEast = "�s�ө���"
            Set c = .cell(2, 1)
            If .cell(1, 1).Range.text = "��y�`��" & VBA.Chr(13) & VBA.Chr(7) Then
                If VBA.InStr(c.Range.text, "��") Then
                    c.Range.text = VBA.Replace(VBA.Replace(c.Range.text, "��", VBA.ChrW(20008)), VBA.Chr(13) & VBA.Chr(7), vbNullString)
                End If
                c.Range.font.Name = "�з���"
            End If
            Set c = .cell(2, 2)
            If .cell(1, 2).Range.text = "�~�y����" & VBA.Chr(13) & VBA.Chr(7) Then
                If VBA.InStr(c.Range.text, "g") Then
                    c.Range.text = VBA.Replace(VBA.Replace(c.Range.text, "g", VBA.ChrW(609)), VBA.Chr(13) & VBA.Chr(7), vbNullString)
                End If
                With c.Range.font
                    .Name = "SimHei"
                    .Size = .Size + .Size * (2 / 12)
                End With
            End If
            
            With .Borders(wdBorderLeft)
                .LineStyle = wdLineStyleDot
                '.LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            With .Borders(wdBorderRight)
                .LineStyle = wdLineStyleDot
                '.LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            With .Borders(wdBorderTop)
                .LineStyle = wdLineStyleDot
                '.LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            With .Borders(wdBorderBottom)
                .LineStyle = wdLineStyleDot
                '.LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            With .Borders(wdBorderHorizontal)
                .LineStyle = wdLineStyleDot
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            With .Borders(wdBorderVertical)
                .LineStyle = wdLineStyleDot
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
            .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
            .Borders.Shadow = False

            .TopPadding = CentimetersToPoints(0.1)
            .BottomPadding = CentimetersToPoints(0.1)
            .LeftPadding = CentimetersToPoints(0.1)
            .RightPadding = CentimetersToPoints(0.1)
            '.Spacing = CentimetersToPoints(0.04)
            .AllowPageBreaks = True
            .AllowAutoFit = True
        
            .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Range.cells.VerticalAlignment = wdCellAlignVerticalCenter
            
        End With
        With options
            .DefaultBorderLineStyle = wdLineStyleDot
            .DefaultBorderLineWidth = wdLineWidth050pt
            .DefaultBorderColor = wdColorAutomatic
        End With
        
    Next tb
    
    rng.tables(1).Range.Find.Execute " ", , , , , , , , , vbNullString, wdReplaceAll
    
'    rng.font.Name = "Times New Roman"
'    rng.font.Name = "Lucida Sans Unicode"
'
'    rng.Select
'    Selection.font.Name = "Times New Roman"
'    Selection.font.Name = "Lucida Sans Unicode"
'    'Selection.font.Name = "Times New Roman"
'    'rng.SetRange rng.tables(2).cell(2, 2).Range.start, rng.End '�@�w�|����@��C
'    rng.tables(3).Select
    
    Dim dict As New Scripting.Dictionary, ziorder, tdi As Byte
    '�r��
    ziorder = WD.FindElementByCssSelector("#ZiOrder").GetAttribute("value")
    If Not dict.Exists(ziorder) Then dict.Add ziorder, vbNullString
    If ziorder <> vbNullString Then
        
        If Not VBA.IsNumeric(VBA.Replace(rng.Paragraphs(1).Previous.Range, VBA.Chr(13), vbNullString)) Then
            rng.Document.Range(rng.start, rng.start).InsertParagraphBefore
            rng.Paragraphs(1).Range.text = ziorder & VBA.Chr(13)
            rng.Paragraphs(1).Style = wdStyleHeading5
        End If
        
        tdi = tdi + 1
        
        
        '�r�Y�I�@�U����
        'Set iwe = WD.FindElementByXPath("/html/body/table[2]/tbody/tr/td[3]/form/table[1]/tbody/tr/td[1]/a")
        'If iwe Is Nothing Then
            Set iwe = WD.FindElementByXPath("/html/body/table[2]/tbody/tr/td[3]/form/table[1]/tbody/tr/td[1]/a[" & tdi & "]")
        'End If
        If Not iwe Is Nothing Then
            ziorder = iwe.GetAttribute("textContent")
            If Not dict.Exists(ziorder) Then
                iwe.Click
                dict.Add ziorder, vbNullString
                
                Dim dt As Date
                dt = VBA.Now
                '�����s��
                Do While WD.FindElementByCssSelector("#PageResult > p:nth-child(2)") Is Nothing
                    If VBA.DateDiff("s", dt, VBA.Now) > 2 Then
                        MsgBox "�SŪ�r���Ū�J���ѡA�Э��աI", vbCritical
                        GoTo finish
                    End If
                Loop
                Do While WD.FindElementByCssSelector("#StartOrder").Displayed = False
                    If VBA.DateDiff("s", dt, VBA.Now) > 2 Then
                        MsgBox "�SŪ�r���Ū�J���ѡA�Э��աI", vbCritical
                        GoTo finish
                    End If
                Loop
                
                SystemSetup.wait 2
                Do Until ziorder = WD.FindElementByCssSelector("#StartOrder").GetAttribute("textContent")
                    If VBA.DateDiff("s", dt, VBA.Now) > 2 Then
                        MsgBox "�SŪ�r���Ū�J���ѡA�Э��աI", vbCritical
                        GoTo finish
                    End If
                Loop
        
                Set iwe = WD.FindElementByCssSelector("#PageResult")
                innerHtml = iwe.GetAttribute("innerHTML") 'outerHTML �]�i�H
                
                Set rng = Selection.Document.Range(rng.End, rng.End)
                rng.InsertParagraphAfter
                rng.Collapse wdCollapseEnd
                rng.InsertAfter ziorder
                rng.ParagraphFormat.Style = wdStyleHeading5
                'rng.InsertParagraphAfter
                rng.Collapse wdCollapseEnd
                Set rng = rng.Document.Range(rng.start, rng.start)
                s = rng.start
                SystemSetup.playSound 0.411
                GoTo �S
            End If
        End If
    End If
    
    result = True
    
finish:
    VBA.DoEvents
    word.Application.Activate
    
    Set dict = Nothing
    
    SystemSetup.contiUndo ur
    word.Application.ScreenUpdating = True
    �d�p�ǰ�W�j����Ū�J�䵲�G = result
    Exit Function
eH:
    Select Case Err.Number
        Case -2146233088
            If VBA.InStr(Err.Description, "stale element reference: stale element not found in the current frame") = 1 Then '(Session info: chrome=131.0.6778.86)
                Dim retryCntr As Byte
                If retryCntr < 3 Then
                    SystemSetup.wait 1
                    retryCntr = retryCntr + 1
                    Resume
                Else
                    Debug.Print Err.Number & Err.Description
                    MsgBox Err.Number & Err.Description, vbCritical
                    GoTo finish
                End If
            Else
                GoTo caseElse
            End If
        Case Else
caseElse:
            Debug.Print Err.Number & Err.Description
            If MsgBox(Err.Number & Err.Description & vbCr & vbCr & "�O�_�n�����H", vbCritical + vbOKCancel) = vbCancel Then
                GoTo finish
            Else
                Stop
                Resume
            End If
    End Select
End Function


Rem �P�_���ɬO�_���� creedit_with_Copilot�j���ġG20241010 https://sl.bing.net/emhYXUvuos8 https://sl.bing.net/cG4Jn2MciZ2
Function IsValidImage_LoadPicture(filePath As String) As Boolean
    On Error Resume Next
    Dim img As Object
    Set img = stdole.LoadPicture(filePath) 'err:481+�Ϥ������T�A���i��O���Ĺ��ɡI
    '�p Set img = stdole.LoadPicture("C:\Users\oscar\Documents\CtextTempFiles\Ctext_Page_Image��.png") 'err:481+�Ϥ������T�A���O���Ĺ��ɡI
    Rem �j���Ȥ䴩 jpg�I�I20241010
    IsValidImage_LoadPicture = Not img Is Nothing
    On Error GoTo 0
End Function
Rem creedit_with_Copilot�j���ġGWordVBA+SeleniumBasicŪ�J�������e�Ϥ��P�W�s���Ghttps://sl.bing.net/fWOLN5PwHsG
Rem  �Ұ�Chrome�s�����þɯ��Ϥ�URL,���ѫh�Ǧ^false�C�o�i�ΡA�������oChrome�s�����U���ؿ��~��
Function DownloadImage_chromedriverExecuteScript(url As String, filePath As String) As Boolean
'    Dim driver As New SeleniumBasic.ChromeDriver
'    driver.start "Chrome"
'    driver.Get url
    Dim driver As SeleniumBasic.IWebDriver, currentWin As String
    If Not SeleniumOP.IsWDInvalid Then
        Set driver = SeleniumOP.WD
        currentWin = driver.CurrentWindowHandle
    Else
        If SeleniumOP.OpenChrome(url) Then
            Set driver = SeleniumOP.WD
        Else
            Exit Function
        End If
    End If
    
    ' ���ݹϤ��[������  'Application�O�ڦۦ氵��Excel�Ҳդ�������C���M�רèS�ޥ� Excel
    'Excel.Application.wait (Now + TimeValue("0:00:05"))
    SystemSetup.wait (Now + TimeValue("0:00:02"))
    
    ' �U���Ϥ� rem �i�H���`�U���A�u�O�n���oChrome�s�������U�����|�~��ѫ���ϥΡI20241010
    driver.ExecuteScript "var link = document.createElement('a'); link.href = arguments[0]; link.download = arguments[1]; document.body.appendChild(link); link.click();", url, filePath
    ' ���ݤU������
    'Excel.Application.wait (Now + TimeValue("0:00:02"))
    SystemSetup.wait (Now + TimeValue("0:00:02"))
    If VBA.Dir(filePath) = vbNullString Or IsValidImage_LoadPicture(url) Then
        Stop
    Else
        DownloadImage_chromedriverExecuteScript = True
    End If
    
    'driver.Quit
    driver.Close
    If currentWin <> vbNullString Then
        If IsWDInvalid() Then
            driver.SwitchTo.Window driver.CurrentWindowHandle
        Else
            WD.SwitchTo.Window currentWin
        End If
    End If
End Function
Rem �N�ϥ� XMLHTTP �ӤU���Ϥ��A�M��N��O�s���Ȧs���.�Y���ѶǦ^false 20241010 creedit_with_Copilot�j���ġGhttps://sl.bing.net/caezeDQDlfg
Function DownloadImage_XMLHTTP_url(url As String, filePath As String) As Boolean
    Dim xmlhttp As Object
    Dim stream As Object
    
    ' �Ы� XMLHTTP ��H
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    xmlhttp.Open "GET", url, False
    xmlhttp.send '-2147467259 �L�k���X�����~�C�i��O�ѩ�z�ϥΪ��O base64 �s�X�� URL�CXMLHTTP �L�k�����B�z base64 �s�X���Ϲ��ƾ� https://sl.bing.net/dd1AOLdKBaK
    
    ' �Ы� ADODB.Stream ��H
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    stream.Write xmlhttp.responseBody
    stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    stream.Close
    If VBA.Dir(filePath) <> vbNullString And IsValidImage_LoadPicture(filePath) Then
        DownloadImage_XMLHTTP_url = True
    End If
End Function
Rem 20241010 Copilot�j���ġG�ϥ� ServerXMLHTTP: ���ɭ� MSXML2.XMLHTTP �|�����D�A�z�i�H���ըϥ� MSXML2.ServerXMLHTTP �ӥN���Chttps://sl.bing.net/dvbNeuzNEjc
Rem XMLHTTP �L�k�����B�z base64 �s�X���Ϲ��ƾڡC�z�ݭn���N base64 �s�X���ƾڸѽX�A�M��A�N��O�s���Ϲ����Chttps://sl.bing.net/b9gYh5mICbc
Function DownloadImage_XMLHTTP(url As String, filePath As String) As Boolean
    On Error GoTo ErrorHandler
    Dim xmlhttp As Object
    Dim stream As Object
    
    ' �Ы� ServerXMLHTTP ��H
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
    xmlhttp.Open "GET", url, False
    xmlhttp.send
    
    ' �Ы� ADODB.Stream ��H
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    stream.Write xmlhttp.responseBody
    stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    stream.Close
    If VBA.Dir(filePath) <> vbNullString And IsValidImage_LoadPicture(filePath) Then
        DownloadImage_XMLHTTP = True
    End If
    Exit Function
    
ErrorHandler:
    Debug.Print Err.Number & Err.Description
    MsgBox "Error: " & Err.Description
    DownloadImage_XMLHTTP = False
End Function

Rem  �N�U�����Ϥ����J��Word���
Sub InsertDownloadedImage(url As String, filePath As String, rng As Range)
    ' �U���Ϥ�
    DownloadImage_chromedriverExecuteScript url, filePath
    
    ' ���J�Ϥ�
    rng.InlineShapes.AddPicture fileName:=filePath, LinkToFile:=False, SaveWithDocument:=True
    
    ' �R���Ȧs���
    Kill filePath
End Sub


Rem 20241010 creedit_with_Copilot�j���ġGWordVBA+SeleniumBasicŪ�J�������e�Ϥ��P�W�s���Ghttps://sl.bing.net/htsW1HREBOe
'Base64ToBinary�G�Nbase64�s�X���Ϥ��ƾ��ഫ���G�i��ƾڡC
'SaveBinaryAsFile�G�N�G�i��ƾګO�s���{�ɹϤ����C
'InsertBase64Image�G�N�{�ɹϤ���󴡤J��Word��󤤡A�ó]�m�Ϥ����e�שM���סC
'�D�{���G�եΤW�z��k�Ӵ��Jbase64�s�X���Ϥ��C
'�o�ˡA�z�N�i�H�Nbase64�s�X���Ϥ����J��Word��󤤡A�ëO����榡�M���e�C
Rem Copilot�j���ġG�ϥΥ��h��F���ӧP�_URL�O�_�]�tbase64�s�X���Ϥ��ƾڪ��d�ҡGhttps://sl.bing.net/c6eMOTPP4wK
Function IsBase64Image(url As String) As Boolean '���ӬO�ѪR�ɥX���F�A�L�ġI20241010�]��y��^
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp") 'If VBA.InStr(url, "data:image/png;base64")  Then 'base64�s�X���Ϥ�
    regex.Pattern = "^data:image\/(png|jpg|jpeg|gif);base64," '�i�H���F���a�P�_�O�_�Obase64�s�X���Ϥ�
    regex.IgnoreCase = True
    IsBase64Image = regex.test(url)
End Function
Rem 20241010 creedit_with_Copilot�j���ġG�ѨMWordVBA + Selenium�U��Chrome�s�������������Ϥ����D�Ghttps://sl.bing.net/dJlhQRbUOHI
Rem �ѽX base64 �s�X���Ϲ��ƾڨñN��O�s�����Ghttps://sl.bing.net/c3GitspZG8G
Function Base64ToImage(base64String As String, filePath As String) As Boolean
    On Error GoTo ErrorHandler
    Dim binaryData() As Byte
    Dim stream As Object
    
    If VBA.InStr(base64String, "data:image/png;base64,") Then
        ' �h�� base64 �Y���� "data:image/png;base64,"
        base64String = Replace(base64String, "data:image/png;base64,", "")
    Else
        Stop
    End If
    
    ' �ѽX base64 �r�Ŧ�
    binaryData = Base64Decode(base64String)
    
    ' �Ы� ADODB.Stream ��H
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    stream.Write binaryData
    stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    stream.Close
    
    If VBA.Dir(filePath) <> vbNullString Then
        If VBA.Right(filePath, 3) <> "png" Then
            Base64ToImage = IsValidImage_LoadPicture(filePath)
        Else
            Base64ToImage = True
        End If
    End If
    If Base64ToImage = True Then base64String = filePath
    Exit Function
    
ErrorHandler:
    MsgBox "Error: " & Err.Description
    Base64ToImage = False
End Function

Function Base64Decode(base64String As String) As Byte()
    Dim xml As Object
    Dim node As Object
    
    ' �Ы� MSXML2.DOMDocument ��H
    Set xml = CreateObject("MSXML2.DOMDocument")
    Set node = xml.createElement("base64")
    node.DataType = "bin.base64"
    node.text = base64String
    Base64Decode = node.nodeTypedValue
End Function

Rem �ѪRbase64�s�X'���ѡI
Function Base64ToBinary(base64String As String) As Byte()
    Dim xmlObj As Object
    Dim base64Data As String
    
    ' �h���e�󳡤�
    base64Data = Mid(base64String, InStr(base64String, ",") + 1)
    
    ' �ѪRbase64�s�X
    Set xmlObj = CreateObject("MSXML2.DOMDocument.6.0")
    xmlObj.LoadXML "<root><binary>" & base64Data & "</binary></root>"
    Base64ToBinary = xmlObj.DocumentElement.ChildNodes(0).nodeTypedValue
End Function
Rem �O�s���{�ɤ�� �N�G�i��ƾګO�s���{�ɹϤ����C'���ѡI
Function SaveBinaryAsFile(binaryData() As Byte, filePath As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    stream.Write binaryData
    stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    stream.Close
End Function
Rem ���Jbase64�s�X�Ϥ�,�N�{�ɹϤ���󴡤J��Word���
Function InsertBase64Image(base64String As String, filePath As String, rng As Range) As InlineShape
    Dim binaryData() As Byte
    Dim tempFilePath As String
    
    ' �ѪRbase64�s�X
    binaryData = Base64ToBinary(base64String)
    
    ' �O�s���{�ɤ��
    tempFilePath = Environ("TEMP") & "\" & filePath
    SaveBinaryAsFile binaryData, tempFilePath
    
    ' ���J�Ϥ�
    Set InsertBase64Image = rng.InlineShapes.AddPicture(fileName:=tempFilePath, LinkToFile:=False, SaveWithDocument:=True, Range:=rng)
    base64String = tempFilePath
    ' �R���{�ɤ��
    Kill tempFilePath
End Function


Function GetDomainUrlPrefix(url As String)
    GetDomainUrlPrefix = VBA.Left(url, VBA.InStr(url, "//")) & "/" & VBA.Mid(url, VBA.InStr(url, "//") + 2, _
                VBA.InStr(VBA.InStr(url, "//") + 2, url, "/") - (VBA.InStr(url, "//") + 2))
End Function
Rem 20241006 �m�ݨ�j�y�P�j�y�����˯��n Ctrl + Alt + j �]j=�y ji�^ �� Alt + d �]d�G�� dian�^
'�쬰 Ctrl + k,d  �A�]�|�Ϥ��ت� Ctrl + k �]���J�W�s���^���ġA�G��w 20241014
Sub �d�ݨ�j�y�j�y�����˯�()
    ��r�B�z.ResetSelectionAvoidSymbols
    SeleniumOP.KandiangujiSearchAll Selection.text
End Sub
Rem 20241006 �˯��m�~�y�����Ʈw�n Alt + Shfit + h �� Alt + h
Sub �d�~�y�����Ʈw()
    ��r�B�z.ResetSelectionAvoidSymbols
    SeleniumOP.HanchiSearch Selection.text
End Sub
Rem 20241006 �HGoogle�˯��m������Ǯѹq�l�ƭp���nAlt + t
Sub �d������Ǯѹq�l�ƭp������()
    ��r�B�z.ResetSelectionAvoidSymbols
    ������Ǯѹq�l�ƭp��.SearchSite
End Sub
Rem 20241006 rng �n�B�z���d��A�Ǧ^��������m
Private Function Ū�J������ƫ�_����ӤJ���}�γ]�w�榡(rng As Range, url As String, fontsize As Single) As Long
    Dim rngNote As Range
    Set rngNote = rng.Document.Range(rng.start, rng.End)
    With rngNote
        '���}�榡�]�w
        If VBA.Len(.Paragraphs.Last.Range.text) > 1 Then .InsertParagraphAfter
        If .Paragraphs.Count > 1 Then
            rngNote.SetRange rng.Paragraphs.Last.Range.start, rng.Paragraphs.Last.Range.End
        End If
        .InsertAfter url '���J���}
        .InsertParagraphAfter
        .End = .End - 1
        If .Characters(1) = Chr(13) Then .start = .start + 1
        .font.Size = fontsize
        .Collapse wdCollapseEnd
        Ū�J������ƫ�_����ӤJ���}�γ]�w�榡 = rng.End 'Range��Selection���ܰʫ�A���G�� With�϶��L�k�ήɦ^���A�n���n�Υ��L�u�A���I�s�v�~��ϬM�ήɯu�� 20241009
    End With
End Function
Rem 20241006 rng �n�B�z���d��
Private Sub Ū�J������ƫ�_�٭�������A(win As word.Window, windowState As word.WdWindowState)
    VBA.DoEvents '�����L�k���`���������D�ѨM�F �I �P���P���@�g���g�ۡ@�n�L�������� 20241127
    With win.Application
        .Activate
        With win
            If .windowState = wdWindowStateMinimize Then
                VBA.Interaction.DoEvents
                .windowState = windowState
                .Activate
                VBA.Interaction.DoEvents
            End If
        End With
    End With

End Sub
Rem Alt + F10(���ֳt��ݽT�{�I�^
Sub �e��j�y�Ŧ۰ʼ��I()
    Dim ur As UndoRecord
    If Selection.Characters.Count < 10 Then
        MsgBox "�r�ƤӤ֡A�����n�ܡH�Цܤ֤j��10�r", vbExclamation
        Exit Sub
    End If
    Selection.Copy
    TextForCtext.GjcoolPunct
    Selection.Document.Activate
    Selection.Document.Application.Activate
    SystemSetup.stopUndo ur, "�e��m�j�y�šn�۰ʼ��I"
    Selection.text = SystemSetup.GetClipboardText
    SystemSetup.contiUndo ur
End Sub
Rem Ctrl + Alt + a
Sub Ū�JAI�Ӫ����I���G()
    playSound 0.484
    Dim startCharacters, endCharacters, d As Document, winState As WdWindowState, rng As Range
'    startCharacters = VBA.Chr(13) & VBA.ChrW(9711) & "���G�@�C�F�A�B}�v�z"
    startCharacters = VBA.Chr(13) & VBA.ChrW(9711) & "���@"
    endCharacters = VBA.Chr(13) & "�v�z�C"
    Set d = Selection.Document:  winState = d.ActiveWindow.windowState
    If Selection.Type = wdSelectionIP Then
        Set rng = Selection.Document.Range(Selection.start, Selection.End)
        '    word.Application.ScreenUpdating = False
        If Not rng.Previous Is Nothing Then
            Do Until VBA.InStr(startCharacters, rng.Previous(wdCharacter).text)
                rng.MoveStart Count:=-1
                If rng.Previous Is Nothing Then Exit Do
            Loop
        End If
        If Not rng.Next Is Nothing Then
            Do Until VBA.InStr(VBA.Chr(13) & VBA.ChrW(9711), rng.Next(wdCharacter).text) Or VBA.InStr(endCharacters, rng.Characters(rng.Characters.Count).text)
                rng.MoveEnd Count:=1
                If rng.Next Is Nothing Then Exit Do
            Loop
        End If
        rng.Select
    '    word.Application.ScreenUpdating = True
    '    SystemSetup.wait 0.4
        word.Application.ScreenRefresh
        
        Set rng = Nothing
    End If
    If inputAITShenShenWikiPunctResult = False Then
        If Selection.Characters.Count > 500 Then
            If MsgBox("�O�_��ѡm�j�y�šn�۰ʼ��I�H", vbOKCancel + vbQuestion) = vbOK Then
                Ū�J�j�y�Ŧ۰ʼ��I���G
            Else
                MsgBox "�Э��աI", vbCritical
            End If
        ElseIf Selection.Characters.Count > 10 Then
            If MsgBox("�O�_��Ρm�j�y�šnAI�۰ʼ��I�H", vbOKCancel + vbQuestion) = vbOK Then
                Ū�J�j�y�Ŧ۰ʼ��I���G
            Else
                MsgBox "�Э��աI", vbCritical
            End If
        Else
            If MsgBox("�O�_��ѡm�j�y�šn�۰ʼ��I�H", vbOKCancel + vbQuestion) = vbOK Then
                Ū�J�j�y�Ŧ۰ʼ��I���G
            End If
        End If
    End If
    DoEvents '���F�o��A�����L�k���`���������D�ѨM�F �I �P���P���@�g���g�ۡ@�n�L�������� 20241127
    d.Activate
    d.ActiveWindow.windowState = winState
    d.ActiveWindow.ScrollIntoView Selection, False
End Sub

Rem Ctrl + Alt + F10 �� Ctrl + Alt + F11
Sub Ū�J�j�y�Ŧ۰ʼ��I���G()
    playSound 0.484
    Dim startCharacters, endCharacters, d As Document, winState As WdWindowState, rng As Range
    startCharacters = VBA.Chr(13) & VBA.ChrW(9711) & "���G�@�C�F�A�B}�v�z"
    endCharacters = VBA.Chr(13) & "�v�z�C"
    Set d = Selection.Document:  winState = d.ActiveWindow.windowState
    If Selection.Type = wdSelectionIP Then
        Set rng = Selection.Document.Range(Selection.start, Selection.End)
        '    word.Application.ScreenUpdating = False
        If Not rng.Previous Is Nothing Then
            Do Until VBA.InStr(startCharacters, rng.Previous(wdCharacter).text)
                rng.MoveStart Count:=-1
            If rng.Previous Is Nothing Then Exit Do
            Loop
        End If
        If Not rng.Next Is Nothing Then
            Do Until VBA.InStr(VBA.Chr(13) & VBA.ChrW(9711), rng.Next(wdCharacter).text) Or VBA.InStr(endCharacters, rng.Characters(rng.Characters.Count).text)
                rng.MoveEnd Count:=1
                If rng.Next Is Nothing Then Exit Do
            Loop
        End If
        rng.Select
    '    word.Application.ScreenUpdating = True
    '    SystemSetup.wait 0.4
        word.Application.ScreenRefresh
        
        Set rng = Nothing
    End If
    If inputGjcoolPunctResult = False Then
        MsgBox "�Э��աI", vbCritical
        killchromedriverFromHere
        Set WD = Nothing
    End If
    DoEvents
    d.Activate
    d.ActiveWindow.windowState = winState
    d.ActiveWindow.ScrollIntoView Selection, False
    
End Sub
Rem 20241008 ���ѫh�Ǧ^false
Function inputGjcoolPunctResult() As Boolean
    Dim ur As UndoRecord, result As String, d As Document
    ��r�B�z.ResetSelectionAvoidSymbols
    If Selection.Characters.Count < 10 Then
        If MsgBox("�r�ƤӤ֡A�����n�ܡH�Цܤ֤j��10�r" & vbCr & vbCr & "�Y����n���I�A�Ы��U�i�T�w�j", vbExclamation + vbOKCancel + vbDefaultButton2) = vbCancel Then
            Exit Function
        End If
    End If
    word.Application.ScreenUpdating = False
    Set d = Selection.Document
    If d.path <> vbNullString And Not d.Saved Then d.Save
    Const ignoreMarker = "�m�n�q�r�u�v�y�z" '�ѦW���B�g�W���B�޸����B�z�]�ѫe�����{���X�B�z�^
    result = Selection.text
    Rem �ѦW���B�޸����B�z
    result = VBA.Replace(VBA.Replace(result, "�m", "�e"), "�n", "�f") '�ѦW����|�Q�۰ʼ��I�M���G,�H���٭� 20241001
    result = VBA.Replace(VBA.Replace(result, "�u", "�e"), "�v", "�f") '�޸���|�Q�۰ʼ��I�M���G,�H���٭� 20241001
    
    Rem �e�h�m�j�y�šn�۰ʼ��I

    If SeleniumOP.grabGjCoolPunctResult_New(result, result) = vbNullString Then
        d.Activate
        d.Application.Activate
        Exit Function
    End If
    d.Activate
    d.Application.Activate
    
    GoSub ���I�ե�
    
    Rem �ѦW���B�޸����B�z
    result = VBA.Replace(VBA.Replace(result, "�e", "�m"), "�f", "�n") '�ѦW����|�Q�۰ʼ��I�M���G,�H���٭� 20241001
    result = VBA.Replace(VBA.Replace(result, "�e", "�u"), "�f", "�v") '�޸���|�Q�۰ʼ��I�M���G,�H���٭� 20241001
    result = VBA.Replace(result, VBA.Chr(13) & VBA.Chr(10), VBA.Chr(13)) 'Ū�^�Ӫ��۰ʼ��I���G�|�Nchr(13)�নVBA.Chr(13) & VBA.Chr(10)
    SystemSetup.stopUndo ur, "Ū�J�m�j�y�šn�۰ʼ��I���G"
    Rem Selection.text = result'�¤�r�B�z
    Dim puncts As New punctuation, cln As New VBA.Collection, e, rng As Range '�A����榡�Ƥ�r
    Set cln = puncts.CreateContextPunctuationCollection(result)
    Rem �M����Ӫ����I�Ÿ��A�H�Q���P���J
    For Each e In Selection.Characters
        'If e = "�C" Then Stop 'just for test
        If e.text = "�@" Then '�Ů�n�M���]�m�j�y�šn�۰ʼ��I�|�M���Ů�^
            e.text = vbNullString
        Else
            If VBA.InStr(ignoreMarker, e.text) = 0 Then '�ѦW���B�޸����B�z�]�ѫe�����{���X�B�z�^
                If puncts.PunctuationDictionary.Exists(e.text) Then
                    e.text = vbNullString
                End If
            End If
        End If
    Next e
    Set rng = d.Range(Selection.start, Selection.End)
    rng.Find.ClearFormatting
    For Each e In cln
'        If e(1) = Chr(13) Then Stop 'just for test
        If e(0) <> vbNullString Then
            If VBA.Len(e(0)) > 255 Then
                MsgBox "�������I�����e�I", vbExclamation
                GoTo finish
            End If
           'If rng.text = e(0) Then'�̫�@��/��
            If VBA.StrComp(rng.text, e(0)) = 0 Then
                rng.InsertAfter e(1)
                rng.Collapse wdCollapseEnd
            ElseIf rng.Find.Execute(e(0), , , , , , True, wdFindStop) = False Then
                If rng.text = e(0) Then '�̫�@��
                    If VBA.InStr(ignoreMarker, e(1)) = 0 Then '�ѦW���B�޸����B�z�]�ѫe�����{���X�B�z�^
                        rng.InsertAfter e(1)
                    End If
'                Else
'                    Stop 'just for test
                End If
            Else
                If VBA.InStr(ignoreMarker, e(1)) = 0 Then '�ѦW���B�޸����B�z�]�ѫe�����{���X�B�z�^
                    rng.InsertAfter e(1)
                Else
                    rng.SetRange rng.start, rng.End + 1
                End If
            End If
        Else
            If VBA.InStr(ignoreMarker & VBA.Chr(13), e(1)) = 0 Then '�ѦW���B�޸����B�z�]�ѫe�����{���X�B�z�^
                rng.Collapse wdCollapseStart
                rng.InsertAfter e(1)
            Else
                rng.SetRange rng.start, rng.start + 1
            End If
        End If
        If rng.End <= Selection.End Then '�̫�@��
            Set rng = d.Range(rng.End, Selection.End)
        Else
            Selection.End = rng.End
        End If
    Next e
    GoSub ���I�ե�_���
finish:
    word.Application.ScreenUpdating = True
    SystemSetup.contiUndo ur
    inputGjcoolPunctResult = True
    Exit Function
    
���I�ե�:
    Dim arrPunct, arrPunctCorrector, iarr As Byte, arrUb As Byte
    arrPunct = Array(VBA.ChrW(-10148) & VBA.ChrW(-9010) & "�G��" & VBA.ChrW(28152) & "��", "�G�r", "�n�r", "�q�G", "�r�G�u�P", "�G�A", "�q�C", "�q�B", "�q�A")
    arrPunctCorrector = Array(VBA.ChrW(-10148) & VBA.ChrW(-9010) & "��" & VBA.ChrW(28152) & "��", "�r�G", "�r", "�q", "�P", "�G", "�C�q", "�B�q", "�q")
    arrUb = UBound(arrPunct)
    For iarr = 0 To arrUb
        If VBA.InStr(result, arrPunct(iarr)) Then
            result = VBA.Replace(result, arrPunct(iarr), arrPunctCorrector(iarr))
        End If
    Next iarr
    
    Return
    
���I�ե�_���:
    rng.SetRange Selection.start, Selection.End
    For iarr = 0 To arrUb
        If VBA.InStr(rng.text, arrPunct(iarr)) Then
            rng.Find.Execute arrPunct(iarr), , , , , , , , , arrPunctCorrector(iarr), wdReplaceAll
        End If
    Next iarr

    Return
End Function
Rem 20241008 ���ѫh�Ǧ^false
Function inputAITShenShenWikiPunctResult() As Boolean
    Dim ur As UndoRecord, result As String, d As Document
    ��r�B�z.ResetSelectionAvoidSymbols
    If Selection.Characters.Count < 10 Then
        MsgBox "�r�ƤӤ֡A�����n�ܡH�Цܤ֤j��10�r" & vbCr & vbCr & "�п�J�������W�U��A�H�K�ҫ��z�Ѥ奻�A�ܤֻݭn10�Ӧr�C", vbCritical
            Exit Function
    End If
    word.Application.ScreenUpdating = False
    Set d = Selection.Document
    If d.path <> vbNullString And Not d.Saved Then d.Save
'    Const ignoreMarker = "�m�n�q�r�u�v�y�z" '�ѦW���B�g�W���B�޸����B�z�]�ѫe�����{���X�B�z�^
    Const ignoreMarker = "�m�n�q�r�]�^�P�u�v�y�z" '�ѦW���B�g�W���B�A���B���`���B�޸����B�z�]�ѫe�����{���X�B�z�^
    result = Selection.text
    Rem �A���B�g�W�����B�z�G�e�B�f�B�e�B�f �|�Q�M����
    result = VBA.Replace(VBA.Replace(result, "�m", "��"), "�n", "��") '�ѦW����|�Q�۰ʼ��I�M���G,�H���٭� 20241106
    result = VBA.Replace(VBA.Replace(result, "�]", "�i"), "�^", "�j") '�A����|�Q�۰ʼ��I�M���G,�H���٭� 202411106
    result = VBA.Replace(VBA.Replace(result, "�q", VBA.ChrW(12310)), "�r", VBA.ChrW(12311))     '�g�W����|�Q�۰ʼ��I�M���G,�H���٭� 20241001
    'result = VBA.Replace(result, "�P", "��")     '���`����|�Q�۰ʼ��I�M���G,�H���٭� 20241001 ��E�h�|�Q���N�����A�L�q�A`�h�`�Q�~��A.�|�Q�M�� 20241124
    result = VBA.Replace(result, "�P", "&")     '���`����|�Q�۰ʼ��I�M���G,�H���٭� 20241124
'
    If SeleniumOP.grabAITShenShenWikiPunctResult(result, result, False) = vbNullString Then
        d.Activate
        d.Application.Activate
        Exit Function
    Else
        result = VBA.Replace(VBA.Replace(VBA.Replace(VBA.Replace(result, VBA.ChrW(8220), "�u"), VBA.ChrW(8221), "�v"), VBA.ChrW(8216), "�y"), VBA.ChrW(8217), "�z")
        result = VBA.Replace(VBA.Replace(result, VBA.ChrW(12310) & "�m", VBA.ChrW(12310)), "�n" & VBA.ChrW(12311), VBA.ChrW(12311))      '�ѦW����|�Q�۰ʼ��I�M���G,�H���٭� 20241106
        result = VBA.Replace(VBA.Replace(result, "�m", "��"), "�n", "��") '�ѦW����|�Q�۰ʼ��I�M���G,�H���٭� 20241106
        result = VBA.Replace(result, "�P", "&")     '���`����|�Q�۰ʼ��I�M���G,�H���٭� 20241001
        result = VBA.Replace(result, "&�F", "�P")
        Rem �X�G�����{�o�����|�ন�u���v
        result = VBA.Replace(result, "��", VBA.ChrW(9711))
    End If
    
    d.Activate
    d.Application.Activate
    Rem �A�����B�z'���I�|�b�]�B����
    result = VBA.Replace(VBA.Replace(result, "��", "�m"), "��", "�n") '�ѦW����|�Q�۰ʼ��I�M���G,�H���٭� 20241001
    result = VBA.Replace(result, "&", "�P")  '���`����|�Q�۰ʼ��I�M���G,�H���٭� 20241001
    result = VBA.Replace(VBA.Replace(result, "�i", "�]"), "�j", "�^")  '�A����|�Q�۰ʼ��I�M���G,�H���٭� 20241001
    result = VBA.Replace(VBA.Replace(result, VBA.ChrW(12310), "�q"), VBA.ChrW(12311), "�r")     '�g�W����|�Q�۰ʼ��I�M���G,�H���٭� 20241001
    'result = VBA.Replace(result, VBA.Chr(13) & VBA.Chr(10), VBA.Chr(13)) 'Ū�^�Ӫ��۰ʼ��I���G�|�Nchr(13)�নVBA.Chr(13) & VBA.Chr(10)
    
    result = VBA.Replace(result, VBA.ChrW(12295), "��")  '"��"�|�Q�m���I
    Debug.Print result
    SystemSetup.stopUndo ur, "Ū�JAI�Ӫ����I���G"
    Rem Selection.text = result'�¤�r�B�z
    Dim puncts As New punctuation, cln As New VBA.Collection, e, rng As Range, rngSub As Range '�A����榡�Ƥ�r
    Set cln = puncts.CreateContextPunctuationCollection(result)
    Rem �M����Ӫ����I�Ÿ��A�H�Q���P���J
    Set rng = d.Range(Selection.start, Selection.End)
    For Each e In rng.Characters
'        'If e = "�C" Then Stop 'just for test
'        If e.text = "�@" Then '�Ů�n�M���]�m�j�y�šn�۰ʼ��I�|�M���Ů�^
'            e.text = vbNullString
'        Else
            If VBA.InStr(ignoreMarker, e.text) = 0 Then '�ѦW���B�޸����B�z�]�ѫe�����{���X�B�z�^
                If puncts.PunctuationDictionary.Exists(e.text) Then
                    e.text = vbNullString
                End If
            End If
'        End If
    Next e
    Set rng = d.Range(Selection.start, Selection.End)
    rng.Find.ClearFormatting
    For Each e In cln
        Set rngSub = rng.Document.Range(rng.start, rng.End)
'        If e(1) = Chr(13) Then Stop 'just for test
        If e(0) <> vbNullString Then
            'If rng.text = e(0) Then'�̫�@��/��
            If VBA.StrComp(rng.text, e(0)) = 0 Then
                rng.InsertAfter e(1)
                rng.Collapse wdCollapseEnd
            ElseIf rng.Find.Execute(e(0), , , , , , True, wdFindStop) = False Then
                'If rng.text = e(0) Then '�̫�@��
                If VBA.StrComp(rng.text, e(0)) = 0 Then '�̫�@��
                    If VBA.InStr(ignoreMarker, e(1)) = 0 Then '�g�W���B�A�����B�z�]�ѫe�����{���X�B�z�^
                        rng.InsertAfter e(1)
                    End If
                Else
                    'Stop 'just for test
                    'rng.Select
                    SystemSetup.ClipboardPutIn VBA.Replace(VBA.CStr(e(0)), VBA.ChrW(9711), "��")
                    If MsgBox("��O���Q�y��A���ˬd�C" & vbCr & vbCr & "�ثe�n���J���I����r���q�O�G " & vbCr & vbCr & "�i" & e(0) & "�j" & vbCr & vbCr & "�n�ץ���󤤪��奻���~��Ы��i�����j", vbExclamation + vbOKCancel) = vbCancel Then '�mAI�Ӫ��n�|�y����A����r�令�q��r�A�ƦܭS��令�d��!! https://www.facebook.com/groups/chineseandme/posts/8754662361296223/
                        '�ثe�O�ѦW���B���`���|�Q�M��ġ��G 20241122
                        Set rng = rng.Document.Range(rngSub.start, rngSub.End)
                        rngSub.Find.Execute "�m", , , , , , , , , vbNullString, wdReplaceAll
                        rngSub.Find.Execute "�n", , , , , , , , , vbNullString, wdReplaceAll
                        rngSub.Find.Execute "�P", , , , , , , , , vbNullString, wdReplaceAll
                        If Selection.Type = wdSelectionIP Then
                            Selection.start = rng.start: Selection.End = rng.End
                        End If
                        If rng.Find.Execute(e(0), , , , , , True, wdFindStop) = False Then
                            If MsgBox("��O���Q�y��A���ˬd�C" & vbCr & vbCr & "�ثe�n���J���I����r���q�O�G " & vbCr & vbCr & "�i" & e(0) & "�j" & vbCr & vbCr & "�n�ץ���󤤪��奻���~��Ы��i�����j", vbExclamation + vbOKCancel) = vbCancel Then '�mAI�Ӫ��n�|�y����A����r�令�q��r�A�ƦܭS��令�d��!! https://www.facebook.com/groups/chineseandme/posts/8754662361296223/
                                If VBA.MsgBox("�n��e�h�m�j�y�šnAI �۰ʼ��I�ܡH", vbQuestion + vbOKCancel) = vbOK Then
                                    Ū�J�j�y�Ŧ۰ʼ��I���G
                                    inputAITShenShenWikiPunctResult = True
                                    GoTo finish
                                Else
                                    GoTo finish
                                End If
                            Else
                                GoTo finish
                            End If
                        Else
                            If VBA.InStr(ignoreMarker, e(1)) = 0 Then '�A���B�g�W�����B�z�]�ѫe�����{���X�B�z�^
                                rng.InsertAfter e(1)
                            Else
                                If rng.Document.Range(rng.End, rng.End + 1) <> e(1) Then
                                    rng.Collapse wdCollapseEnd
                                    rng.InsertAfter e(1)
                                Else
                                    rng.SetRange rng.start, rng.End + 1
                                End If
                            End If
                        End If
                    Else
                        GoTo finish
                    End If
                End If
            Else
                If VBA.InStr(ignoreMarker, e(1)) = 0 Then '�A���B�g�W�����B�z�]�ѫe�����{���X�B�z�^
                    rng.InsertAfter e(1)
                Else
                    If rng.Document.Range(rng.End, rng.End + 1) <> e(1) Then
                        rng.Collapse wdCollapseEnd
                        rng.InsertAfter e(1)
                    Else
                        rng.SetRange rng.start, rng.End + 1
                    End If
                End If
            End If
        Else
            If VBA.InStr(ignoreMarker & VBA.Chr(13), e(1)) = 0 Then '�g�W���B�A�����B�z�]�ѫe�����{���X�B�z�^
                rng.Collapse wdCollapseStart
                rng.InsertAfter e(1)
            Else
                If rng.Document.Range(rng.start, rng.start + 1) <> e(1) Then
                    rng.Collapse wdCollapseStart
                    rng.InsertAfter e(1)
                Else
                    rng.SetRange rng.start, rng.start + 1
                End If
            End If
        End If
        If rng.End <= Selection.End Then '�̫�@��
'            If rng.End = d.Range.End Then
'                Set rng = d.Range(rng.End - 1, Selection.End)
'            Else
                Set rng = d.Range(rng.End, Selection.End)
'            End If
        Else
            Selection.End = rng.End
        End If
        
    Next e
    inputAITShenShenWikiPunctResult = True
    
finish:
    SystemSetup.contiUndo ur
    word.Application.ScreenUpdating = True
    
End Function
Function GetUserAddress() As Boolean
    Dim x As String, a As Object 'Access.Application
    On Error GoTo Error_GetUserAddress
    x = Selection.text
    Set a = GetObject("D:\�d�{�@�o�N\���y���\�ϮѺ޲z.mdb") '2010/10/18�׭q
    If x = "" Then x = InputBox("�п�J���d�ߪ��r��")
    x = a.Run("�d�ߦr���ഫ_��y�|�X", x)
''    'ActiveDocument.FollowHyperlink "http://140.111.34.46/cgi-bin/dict/newsearch.cgi", , False, , "Database=dict&GraphicWord=yes&QueryString=^" & X & "$", msoMethodGet
'    FollowHyperlink "http://dict.revised.moe.edu.tw/cgi-bin/newDict/dict.sh?", , False, , "=dict.idx&cond=^" & x & "$&pieceLen=50&fld=1&cat=&imgFont=1", msoMethodGet
    Shell Replace(GetDefaultBrowserEXE, """%1", "http://dict.revised.moe.edu.tw/cgi-bin/newDict/dict.sh?cond=^" & x & "$&pieceLen=50&fld=1&cat=&imgFont=1")
    'AppActivate GetDefaultBrowser'�L��
'    'FollowHyperlink "http://dict.revised.moe.edu.tw/cgi-bin/newDict/dict.sh?", , False, , "=dict.idx&cond=^" & X & "$&pieceLen=50&fld=1&cat=&imgFont=1", msoMethodGet
    
'    If Len(Selection.Text) = 1 Then _
        FollowHyperlink "http://www.nlcsearch.moe.gov.tw/EDMS/admin/dict3/search.php", , False, , "qstr=" & x & "&dictlist=47,46,51,18,16,13,20,19,53,12,14,17,48,57,24,25,26,29,30,31,32,33,34,35,36,37,39,38,41,42,43,45,50,&searchFlag=A&hdnCheckAll=checked", msoMethodGet '2009/1/10'�Ш|��-��a�y���X�s���˯��t��-�y���X�˯�
        If a.Visible = False Then
            a.Visible = True
            a.UserControl = True
        End If
'        a.Quit acQuitSaveNone
'        Set a = Nothing
    GetUserAddress = True
Exit_GetUserAddress:
    Exit Function

Error_GetUserAddress:
    MsgBox Err & ": " & Err.Description
    GetUserAddress = False
    Resume Exit_GetUserAddress
End Function


    
Function GetDefaultBrowser() '2010/10/18��http://chijanzen.net/wp/?p=156#comment-1303(���o�w�]�s����(default web browser)���W��? chijanzen ���f�E)�Ө�.
    Dim objShell
    Set objShell = CreateObject("WScript.Shell")
    'HKEY_CLASSES_ROOT\HTTP\shell\open\ddeexec\Application
    '���o���U������
    GetDefaultBrowser = objShell.RegRead _
            ("HKCR\http\shell\open\ddeexec\Application\")
    'GetDefaultBrowser = objShell.RegRead _
            ("HKEY_CLASSES_ROOT\http\shell\open\ddeexec\Application\")
End Function


Function GetDefaultBrowserEXE() '2010/10/18��http://chijanzen.net/wp/?p=156#comment-1303(���o�w�]�s����(default web browser)���W��? chijanzen ���f�E)�Ө�.
Dim deflBrowser As String
deflBrowser = getDefaultBrowserNameAppActivate
Select Case deflBrowser
    Case "iexplore":
        GetDefaultBrowserEXE = "C:\Program Files\Internet Explorer\iexplore.exe"
    Case "firefox":
        If Dir("W:\PortableApps\PortableApps\FirefoxPortable\App\Firefox64\firefox.exe") = "" Then
            GetDefaultBrowserEXE = "C:\Program Files\Mozilla Firefox\firefox.exe"
        Else
            GetDefaultBrowserEXE = "W:\PortableApps\PortableApps\FirefoxPortable\App\Firefox64\firefox.exe"
        End If
    Case "brave":
        If Dir(UserProfilePath & "\AppData\Local\BraveSoftware\Brave-Browser\Application\brave.exe") = "" Then
            GetDefaultBrowserEXE = "C:\Program Files (x86)\BraveSoftware\Brave-Browser\Application\brave.exe"
        Else
            GetDefaultBrowserEXE = UserProfilePath & "\AppData\Local\BraveSoftware\Brave-Browser\Application\brave.exe"
        End If
    Case "vivaldi":
        GetDefaultBrowserEXE = UserProfilePath & "\AppData\Local\Vivaldi\Application\vivaldi.exe"
    Case "Opera":
        GetDefaultBrowserEXE = ""
    Case "Safari":
        GetDefaultBrowserEXE = ""
    Case "edge":
        GetDefaultBrowserEXE = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" '"msedge"
    Case "ChromeHTML", "google chrome": '"chrome"
        GetDefaultBrowserEXE = SystemSetup.getChrome
'
'        If Dir("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe") = "" Then
'            GetDefaultBrowserEXE = "W:\PortableApps\PortableApps\GoogleChromePortable\GoogleChromePortable.exe"
'        Else
'            GetDefaultBrowserEXE = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
'        End If
    Case Else:
        Dim objShell
        Set objShell = CreateObject("WScript.Shell")
        'HKEY_CLASSES_ROOT\HTTP\shell\open\ddeexec\Application
        '���o���U������
        deflBrowser = objShell.RegRead _
                ("HKCR\http\shell\open\command\")
        GetDefaultBrowserEXE = VBA.Mid(deflBrowser, 2, InStr(deflBrowser, ".exe") + Len(".exe") - 2)

End Select
    
    
End Function

Function getDefaultBrowserFullname()
Dim appFullname As String
appFullname = GetDefaultBrowserEXE
'appFullname = VBA.Mid(appFullname, 2, InStr(appFullname, ".exe") + Len(".exe") - 2)
getDefaultBrowserFullname = appFullname
'DefaultBrowserNameAppActivate = VBA.Replace(VBA.Mid(appFullname, InStrRev(appFullname, "\") + 1), ".exe", "")
End Function


Function getDefaultBrowserNameAppActivate() As String
Dim objShell, ProgID As String: Set objShell = CreateObject("WScript.Shell")
ProgID = objShell.RegRead _
            ("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice\ProgID")
ProgID = VBA.Mid(ProgID, 1, IIf(InStr(ProgID, ".") = 0, Len(ProgID), InStr(ProgID, ".") - 1))
Select Case ProgID
    Case "IE.HTTP":
        DefaultBrowserNameAppActivate = "iexplore"
    Case "FirefoxURL":
        DefaultBrowserNameAppActivate = "firefox"
    Case "ChromeHTML":
        DefaultBrowserNameAppActivate = "google chrome"
    Case "BraveHTML":
        DefaultBrowserNameAppActivate = "brave"
    Case "VivaldiHTM":
        DefaultBrowserNameAppActivate = "vivaldi"
    Case "OperaStable":
        DefaultBrowserNameAppActivate = "Opera"
    Case "SafariHTML":
        DefaultBrowserNameAppActivate = "Safari"
    Case "AppXq0fevzme2pys62n3e0fbqa7peapykr8v", "MSEdgeHTM":
        'browser = BrowserApplication.Edge;
        DefaultBrowserNameAppActivate = "edge" '"msedge"
    Case Else:
        DefaultBrowserNameAppActivate = "google chrome" '"chrome"
End Select
getDefaultBrowserNameAppActivate = DefaultBrowserNameAppActivate
End Function

Rem AppActivate��k�`�|���I�I20241020
Sub AppActivateDefaultBrowser()
    On Error GoTo eH
    Dim i As Byte, a
    a = Array("google chrome", "brave", "edge")
    DoEvents
    If DefaultBrowserNameAppActivate = "" Then getDefaultBrowserNameAppActivate
    VBA.Interaction.AppActivate DefaultBrowserNameAppActivate
    DoEvents
    Exit Sub
eH:
        Select Case Err.Number
            Case 5
                DefaultBrowserNameAppActivate = a(i)
                i = i + 1
                If i > UBound(a) Then
                    MsgBox Err.Number & Err.Description
                    Exit Sub
                End If
                Resume
            Case Else
                MsgBox Err.Number + Err.Description
        End Select
    'AppActivate ""
End Sub



