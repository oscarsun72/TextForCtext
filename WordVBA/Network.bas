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
    SeleniumOP.LookupZitools Selection.text
End Sub
Sub �d����r�r��()
    Rem Alt + F12
    ��r�B�z.ResetSelectionAvoidSymbols
    If Selection.Characters.Count > 1 Then
        MsgBox "���d1�r", vbExclamation ', vbError
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
    SeleniumOP.LookupKangxizidian Selection.text
End Sub
Sub �d��y���_������h��()
    Rem Ctrl + Alt + F12
    ��r�B�z.ResetSelectionAvoidSymbols
    SeleniumOP.LookupDictRevised Selection.text
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
Sub �d��Ǥj�v()
    Rem Ctrl + d + s �]ds�G�j�v�^
    ��r�B�z.ResetSelectionAvoidSymbols
    SeleniumOP.LookupGXDS Selection.text
End Sub
Sub �d�ն��`�B�H�a����Ѧr�Ϲ��d�\_�ê�k���u��()
    Rem  Alt + s �]���媺���^ Alt + j �]�Ѧr���ѡ^
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
                    InnerHTML_Convert_to_WordDocumentContent rngHtml ', vbNullString
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
    Dim linkInput As Boolean, rngLink As Range, rngHtml As Range
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
        
        If Selection.Characters.Count > 2 Then
errExit:
            word.Application.Activate
            VBA.MsgBox "���W���~! �Э��s����C", vbExclamation
            Exit Sub
        End If
    End If
    Dim gua As String
    gua = Selection.text
    
    Dim ur As UndoRecord, s As Long, ed As Long
    SystemSetup.stopUndo ur, "�d���Ǻ����g�P�������w���W�奻_�è��^��¤�r�Ȥκ��}�ȴ��J�ܴ��J�I��m"
    word.Application.ScreenUpdating = False

    Dim result(1) As String, iwe As SeleniumBasic.IWebElement

    If linkInput Then
        If SeleniumOP.OpenChrome(rngLink.Hyperlinks(1).Address) Then
            Set iwe = WD.FindElementByCssSelector("#block-bartik-content > div > article > div > div.clearfix.text-formatted.field.field--name-body.field--type-text-with-summary.field--label-hidden.field__item")
            If Not iwe Is Nothing Then
                Selection.MoveUntil Chr(13)
                Selection.TypeText Chr(13)
                Selection.Style = word.wdStyleNormal '"����"
                
                If SeleniumOP.IslinkImageIncluded���e�����]�t�W�s���ιϤ�(iwe) Then '���Ϥ��ɨ� "innerHTML" �ݩʭ�
'                    Dim Links() As SeleniumBasic.IWebElement, images() As SeleniumBasic.IWebElement
'                    Links = SeleniumOP.Links
'                    images = SeleniumOP.images
                    s = Selection.start
                    Selection.TypeText iwe.GetAttribute("innerHTML")
                    ed = Selection.End
                                        
                    Set rngHtml = Selection.Document.Range(s, ed)
                    
                    
                    
                    InnerHTML_Convert_to_WordDocumentContent rngHtml, "https://www.eee-learning.com"
                    'SeleniumOP.inputElementContentAll���J��������Ҧ������e iwe
                    
                    
                    

'                    Stop 'just for test
                    GoTo finish 'just for test


                Else '�S���Ϥ��ɨ� "textContent" �ݩʭ�
                    result(0) = iwe.GetAttribute("textContent")
                    result(1) = rngLink.Hyperlinks(1).Address
                End If
                
                GoTo insertText:
            End If
        End If
    Else

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
    Dim windowState As word.WdWindowState      '�O�U��Ӫ������Ҧ�
    windowState = word.Application.windowState '�O�U��Ӫ������Ҧ�
    
    gua = Keywords.�P�����W_����_����(gua)(1)

    Dim fontsize As Single, rngBooks As Range
    fontsize = VBA.IIf(Selection.font.Size = 9999999, 12, Selection.font.Size) * 0.6
    If fontsize < 0 Then fontsize = 10
    
    If Not SeleniumOP.grabEeeLearning_IChing_ZhouYi_originalText(gua, result, iwe) Then
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
                End If
                s = Selection.start
                
                Selection.TypeText iwe.GetAttribute("innerHTML")
                ed = Selection.End
                
                
                
                Set rngHtml = Selection.Document.Range(s, ed)
                
                InnerHTML_Convert_to_WordDocumentContent rngHtml, "https://www.eee-learning.com"
                'SeleniumOP.inputElementContentAll���J��������Ҧ������e iwe
                
                
                
                Rem ���N�`���G
                Set rngBooks = Selection.Document.Range(rngHtml.start, rngHtml.End)
                rngBooks.Find.ClearFormatting
                If rngBooks.Find.Execute("���N�`���G", , , , , , True, wdFindStop) Then
                    With rngBooks '"���N�`���G"�Ҧb�q���d��
                        .Style = wdStyleHeading1
                        .font.Size = 22
                        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle '��涡�Z
                    End With
                    ed = Ū�J������ƫ�_����ӤJ���}�γ]�w�榡(rngHtml, result(1), fontsize)
                    Ū�J������ƫ�_�٭�������A Selection.Document.ActiveWindow, windowState
                End If

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
        
    Dim p As Paragraph, book As String, iwes() As IWebElement, e, x As String
    
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

    word.Application.ScreenUpdating = True
    Selection.Document.Range(s, s).Select
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
Rem 20241009 �NHTML�নWord��󤺤�Ccreedit_with_Copilot�j���ġGhttps://sl.bing.net/jij3PK59Rka
Sub InnerHTML_Convert_to_WordDocumentContent(rngHtml As Range, Optional domainUrlPrefix As String)
    If VBA.InStr(rngHtml.text, "<") = 0 Then Exit Sub
    
     SystemSetup.playSound 1
    
    Dim htmlStr As String, rng As Range, rngClose As Range, p As Paragraph, url As String, stRngHTML As Long, pCntr As Long
    Dim s As Integer '�@�� InStr() �O�U���G�ȥ�
    Dim l As Integer '�@�� Len() �O�U���G�ȥ�
    '�@���q���ܼƥΡA�ΰ}�C�O���
    Dim arr, arr1, e '�@���q�Τ@���ܼƥΡA�ΰ}�C�����O���
    Dim obj As Object '�@���q�Ϊ����ܼƥ�
    
    'dim w As Single, h As Single, textPart As String
    
    'Dim ur As UndoRecord  'just for test
    
'    GoTo finish 'just for test
    
    '���o���}�e�󪺺���ȡ]���t���׽u�^
    If domainUrlPrefix = vbNullString Then
        If Not SeleniumOP.IsWDInvalid() Then ' domainUrlPrefix = "https://www.eee-learning.com"
            domainUrlPrefix = getDomainUrlPrefix(WD.url)
        End If
    End If
    'SystemSetup.stopUndo ur, "InnerHTML_DocContent"
    stRngHTML = rngHtml.start
    htmlStr = rngHtml.text '�O�U�_�l��m
    
    Rem �e�m��z�奻
    rngHtml.text = VBA.Replace(VBA.Replace(VBA.Replace(htmlStr, "</p>", vbNullString), "<p>", vbNullString), "&nbsp;", ChrW(160))
    htmlStr = rngHtml.text
    
    If VBA.InStr(htmlStr, "<sup>") Then
        Set rng = rngHtml.Document.Range(rngHtml.start, rngHtml.End)
        HTML2Doc.ConvertHTMLSupToWordSup rng
    End If
    If VBA.InStr(htmlStr, "<sub>") Then
        Set rng = rngHtml.Document.Range(rngHtml.start, rngHtml.End)
        HTML2Doc.ConvertHTMLSubToWordSub rng
    End If
    With rngHtml.Find
        .ClearFormatting
        '�m��
        If VBA.InStr(htmlStr, "<br>") Then .Execute "<br>", , , , , , , , , "^l", wdReplaceAll
        If VBA.InStr(htmlStr, "<a style=""line-height:1.5;"" href=") Then .Execute "<a style=""line-height:1.5;"" href=", , , , , , , , , "<a href=", wdReplaceAll ' " �|�m���� ��
        If VBA.InStr(htmlStr, "&lt;") Then .Execute "&lt;", , , , , , , , , "��", wdReplaceAll
        If VBA.InStr(htmlStr, "&gt;") Then .Execute "&gt;", , , , , , , , , "��", wdReplaceAll
        '�M��
'        If VBA.InStr(htmlStr, "<div>" & ChrW(160) & "</div>") Then .Execute "<div>" & ChrW(160) & "</div>", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, ChrW(160)) Then .Execute ChrW(160), , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, " class=""colorbox cboxElement""") Then .Execute " class=""colorbox cboxElement""", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, " class=""colorbox colorbox-insert-image cboxElement""") Then .Execute " class=""colorbox colorbox-insert-image cboxElement""", , , , , , , , , vbNullString, wdReplaceAll '
        'If VBA.InStr(htmlStr, "<a class=""colorbox colorbox-insert-image cboxElement"" href=") Then .Execute "<a class=""colorbox colorbox-insert-image cboxElement"" href=", , , , , , , , , "<a href=", wdReplaceAll ' " �|�m���� ��
        If VBA.InStr(htmlStr, " rel=""group-all""") Then .Execute " rel=""group-all""", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "<o:p></o:p>") Then .Execute "<o:p></o:p>", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "<span></span>") Then .Execute "<span></span>", , , , , , , , , vbNullString, wdReplaceAll
        Rem ������\�νѦpWord���s��K�W�A�G�h���ݽX�B�ýX
        If VBA.InStr(htmlStr, "<o:p>") Then .Execute "<o:p>", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "</o:p>") Then .Execute "</o:p>", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "<span style=""color:#ffffff;"">ppp</span>") Then .Execute "<span style=""color:#ffffff;"">ppp</span>", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "<!--EndFragment-->") Then .Execute "<!--EndFragment-->", , , , , , , , , vbNullString, wdReplaceAll
    End With
    
    SystemSetup.playSound 1
    Set rng = rngHtml.Document.Range(rngHtml.start, rngHtml.End)
    '�M���ż���
    RemoveEmptyTags rngHtml
    
    Rem ���B�z https://sl.bing.net/fQ5lVr8PLye
    Do While VBA.InStr(rngHtml.text, "<table")
        Set rng = rngHtml.Document.Range(rngHtml.start, rngHtml.End)
        With rng.Find
            .ClearFormatting
            .text = "<table "
            .Execute
            Set rngClose = rng.Document.Range(rng.End, rngHtml.End)
            With rngClose.Find
                .text = "</table>"
                .Execute
            End With
            InsertHTMLTable rngHtml.Document.Range(rng.start, rngClose.End), domainUrlPrefix
        End With
    Loop
    
    Set rng = rngHtml.Document.Range(rngHtml.start, rngHtml.End)
    Rem �L�ǲM�檺�B�z
    unorderedListPorc_HTML2Word rng
    
    For Each p In rngHtml.Paragraphs
        pCntr = pCntr + 1
        If pCntr Mod 20 = 0 Then SystemSetup.playSound 1
        
        Set rng = p.Range '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
        With rng
            
'            If VBA.InStr(rng.text, "�����A�s�e�A�ڤ�") Then
'                Stop 'check
'            End If

'            If VBA.InStr(rng.text, "���s�ťΡA���b�U�]") Then
'                Stop 'check
'            End If
            
            With .Find
                .ClearFormatting
                .text = "<b>"
                Do While .Execute()
                    Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                    rngClose.Find.ClearFormatting
                    If Not rngClose.Find.Execute("</b>") Then Stop 'to check
                    rng.Document.Range(rng.End, rngClose.start).font.Bold = True
                    rng.text = vbNullString: rngClose.text = vbNullString
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                Loop
                .text = "<b "
                Do While .Execute()
                    rng.MoveEndUntil ">"
                    rng.End = rng.End + 1
                    Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                    rngClose.Find.ClearFormatting
                    If Not rngClose.Find.Execute("</b>") Then Stop 'to check
                    rng.Document.Range(rng.End, rngClose.start).font.Bold = True
                    rng.text = vbNullString: rngClose.text = vbNullString
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                Loop
                .text = "<span lang=""EN-US"">"
                Do While .Execute()
                    Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                    rngClose.Find.ClearFormatting
                    If Not rngClose.Find.Execute("</span>") Then Stop 'to check
                    rng.Document.Range(rng.End, rngClose.start).font.Name = "Calibri"
                    rng.text = vbNullString: rngClose.text = vbNullString
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                Loop
                .text = "<st1:chmetcnv "
                Do While .Execute()
                    rng.MoveEndUntil ">"
                    rng.End = rng.End + 1
                    Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                    rngClose.Find.ClearFormatting
                    If Not rngClose.Find.Execute("</st1:chmetcnv>") Then Stop 'to check
                    rng.text = vbNullString: rngClose.text = vbNullString
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                Loop
                .text = "<span class=""Apple-style-span"""
                Do While .Execute()
                    rng.MoveEndUntil ">"
                    rng.End = rng.End + 1
                    Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                    rngClose.Find.ClearFormatting
                    If Not rngClose.Find.Execute("</span>") Then Stop 'to check
                    rng.text = vbNullString: rngClose.text = vbNullString
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                Loop
                .text = "<blockquote>"
                Do While .Execute()
                    Set rngClose = rng.Document.Range(rng.End, p.Next.Range.End)
                    rngClose.Find.ClearFormatting
                    If Not rngClose.Find.Execute("</blockquote>") Then Stop 'to check
                    rng.text = vbNullString
                    If rngClose.Paragraphs(1).Range.text = "</blockquote>" & Chr(13) Then
                        rngClose.Paragraphs(1).Range.text = vbNullString
                    Else
                        Stop 'for check
                        rngClose.text = vbNullString
                    End If
                    rng.ParagraphFormat.CharacterUnitLeftIndent = 3
                    rng.Paragraphs(1).Range.font.Name = "�з���"
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                Loop
                .text = "<hr>"
                If .Execute() Then
                    rng.text = vbNullString
                    'rng.InsertBreak WdBreakType.wdLineBreak
                    With p.Borders(wdBorderBottom)
                        .LineStyle = wdLineStyleSingle ' ���J��u  ���J���u�GwdLineStyleDouble ���J��u�GwdLineStyleDot
                        .LineWidth = wdLineWidth050pt
                        .Color = wdColorAutomatic
                    End With
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                End If
                .text = "<hr " 'ex: <hr style="padding-left: 30px;">
                If .Execute() Then
                    rng.MoveEndUntil ">"
                    rng.End = rng.End + 1
                    '�ɥ� url �ܼ�
                    url = rng.text
                    rng.text = vbNullString
                    'rng.InsertBreak WdBreakType.wdLineBreak
                    With p.Borders(wdBorderBottom)
                        .LineStyle = wdLineStyleSingle ' ���J��u  ���J���u�GwdLineStyleDouble ���J��u�GwdLineStyleDot
                        .LineWidth = wdLineWidth050pt
                        .Color = wdColorAutomatic
                    End With
                    url = getHTML_AttributeValue("style", url)
                    arr = VBA.Split(url, ";")
                    For Each e In arr
                        If e <> vbNullString Then
                            e = VBA.Trim(e)
                            l = VBA.Len("padding-left: ")
                            If VBA.Left(e, l) = "padding-left: " Then
                                rng.ParagraphFormat.LeftIndent = PixelsToPoints(VBA.Replace(VBA.Mid(e, l + 1), "px", vbNullString))
                            Else
                                playSound 12
                                Stop 'for check
                            End If
                        End If
                    Next e

                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                End If
                '�B�z�Ϥ�
                .text = "<img "
                Do While .Execute()
                    rng.MoveEndUntil ">" 'ex: <img style="float:right;margin-left:15px;margin-right:15px;" src="/image/3.jpg" width="200" height="297"
                    '�ɥ��ܼ�
                    url = rng.text
                    rng.End = rng.End + 1 '�]�t ">"
                    rng.text = vbNullString
                    'pCntr + VBA.Abs(10 - pCntr) '�U���Ϥ��ݭn�ɶ�
                    If Not insert_ImageHTML(url, rng, domainUrlPrefix) Is Nothing Then
                        p.Range.ParagraphFormat.BaseLineAlignment = wdBaselineAlignCenter
                    End If
'                    If rng.Paragraphs(1).Range.ShapeRange.Count > 0 Then
'                        Stop
'                    End If
                    
                    rng.SetRange rng.End, p.Range.End
                    
                Loop
                If VBA.Len(p.Range.text) > VBA.Len("<strong></strong>") Then
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                    .text = "<strong>"
                    Do While .Execute()
                        Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                        rngClose.Find.text = "</strong>"
                        If Not rngClose.Find.Execute() Then
                            rngClose.SetRange rngClose.End, rngClose.Paragraphs(1).Next.Range.End
                            If Not rngClose.Find.Execute() Then Stop 'for check
                        End If
                        rng.Document.Range(rng.End, rngClose.start).font.Bold = True
                        rng.text = vbNullString: rngClose.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                If VBA.Len(p.Range.text) > VBA.Len("<strong style=""; ;""></strong>") Then
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                    .text = "<strong style="""
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                        rngClose.Find.text = "</strong>"
                        rngClose.Find.Execute
                        rng.Document.Range(rng.End, rngClose.start).font.Bold = True
                        rng.text = vbNullString: rngClose.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                '�B�z�r���˦�
                If VBA.Len(p.Range.text) > VBA.Len("<span style=""></span>") Then
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                    .text = "<span style"
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                        rngClose.Find.text = "</span>"
                        rngClose.Find.Execute
                        
'                        If InStr(p.Range.text, "����@�]�P�l�@�W�pġ") Then Stop 'just for test
                        
                        '�ɥ�url�ܼ�
                        url = VBA.Replace(getHTML_AttributeValue("span style", p.Range.text), "font-family:", vbNullString)
                        url = VBA.Left(url, VBA.Len(url) - 1)
                        Select Case url
                            Case "font-size: x-large", "font-size:x-large"
                                rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size * 2
                            Case "font-size: small", "font-size:small"
                                rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size * (5 / 6)
                            Case "font-size: x-small", "font-size:x-small"
                                rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size * (2 / 3)
                            Case "font-size: xx-small", "font-size:xx-small"
                                rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size * (1 / 2)
                            Case "text-decoration:underline"
                                rng.Document.Range(rng.End, rngClose.start).font.Underline = wdUnderlineSingle
                            Case Else
                            
                                If VBA.InStr(url, ";") = 0 And VBA.InStr(url, "; ") = 0 And VBA.InStr(url, "font-size:") <> 1 And VBA.InStr(url, "line-height:") = 0 And VBA.InStr(url, "font-family") = 0 And VBA.InStr(url, "Mso") = 0 And VBA.InStr(url, "mso-") = 0 And VBA.InStr(url, "�з���") = 0 And VBA.InStr(url, "letter-spacing:0pt") = 0 And VBA.InStr(url, "�s�ө���") = 0 And VBA.InStr(url, "background-color: ") = 0 And VBA.InStr(url, "color: ") = 0 Then
                                    
                                    rng.Select
                                    Debug.Print url
                                    Stop 'for check
                                End If
                                
                                'FontName
                                If VBA.Left(url, 3) = "�з���" Then url = "�з���"
                                If Fonts.IsFontInstalled(VBA.Trim(url)) Then
                                    If rng.Document.Range(rng.End, rngClose.start).font.Name <> VBA.Trim(url) Then
                                        rng.Document.Range(rng.End, rngClose.start).font.Name = VBA.Trim(url)
                                    End If
                                End If
                                'FontSzie
                                If VBA.InStr(url, "font-size:") = 1 Then
                                    l = VBA.Len("font-size:")
                                    If VBA.Right(url, 2) = "em" Then ' em �O�@�Ӭ۹���A�Ω�]�m�r��j�p�C���۹����������r��j�p�C�Ҧp�A�p�G���������r��j�p�O16�����A�h 1em ����16�����A1.5em ����24�����C20241011 https://sl.bing.net/bVzA9JEh8VM
                                        l = VBA.Len("font-size:")
                                        rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size _
                                                * VBA.CSng(VBA.Trim(VBA.Mid(url, l + 1, VBA.Len(url) - l - VBA.Len("em"))))
                                    ElseIf VBA.IsNumeric(VBA.Mid(url, l + 1)) Then
                                        rng.Document.Range(rng.End, rngClose.start).font.Size = VBA.CSng(IIf(VBA.Mid(url, l + 1) < 1, VBA.Mid(url, l + 1) * 10, VBA.Mid(url, l + 1)))
                                    Else
                                        
                                        If VBA.InStr(url, "font-size: medium") = 0 And VBA.InStr(url, "font-size:medium") = 0 Then
                                            Stop
                                        End If
                                    End If
                                End If
                                '�r���q����L�榡������
                                If VBA.InStr(url, "; ") Or VBA.InStr(url, ";") Then
                                    arr = VBA.Split(url, ";")
                                    For Each e In arr
                                        e = VBA.Trim(e)
                                        If VBA.Left(e, 17) = "background-color:" Then
                                            arr1 = colorCodetoRGB(VBA.LTrim(VBA.Mid(e, VBA.Len("background-color:") + 1)))
                                            rng.Document.Range(rng.End, rngClose.start).font.Shading.BackgroundPatternColor = VBA.RGB(arr1(0), arr1(1), arr1(2))
                                        ElseIf VBA.Left(e, 6) = "color:" Then
                                            arr1 = colorCodetoRGB(VBA.LTrim(VBA.Mid(e, VBA.Len("color:") + 1)))
                                            rng.Document.Range(rng.End, rngClose.start).font.Color = VBA.RGB(arr1(0), arr1(1), arr1(2))
                                        ElseIf VBA.Left(e, 12) = "line-height:" Then
                                            arr1 = VBA.LTrim(VBA.Mid(e, VBA.Len("line-height:") + 1))
                                            If Not VBA.IsNumeric(arr1) Then
                                                If VBA.InStr(arr1, "px") Then
                                                    arr1 = VBA.Replace(arr1, "px", vbNullString)
                                                Else
                                                    playSound 12 'for check
                                                    Stop
                                                End If
                                            End If
                                            If arr1 < 10 Then
                                                rng.Document.Range(rng.End, rngClose.start).ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
                                                rng.Document.Range(rng.End, rngClose.start).ParagraphFormat.LineSpacing = VBA.CSng(arr1)
                                            Else
                                                rng.Document.Range(rng.End, rngClose.start).ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
                                                rng.Document.Range(rng.End, rngClose.start).ParagraphFormat.LineSpacing = VBA.CSng(arr1)
                                            End If
                                        ElseIf VBA.Left(e, 10) = "font-size:" Then
                                            arr1 = VBA.Replace(VBA.LTrim(VBA.Mid(e, VBA.Len("font-size:") + 1)), "px", vbNullString)
                                            If VBA.IsNumeric(arr1) Then
                                                rng.Document.Range(rng.End, rngClose.start).font.Size = VBA.CSng(arr1)
                                            Else
                                                If arr1 = "x-small" Then
                                                    rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size * (2 / 3)
                                                ElseIf arr1 = "medium" Then
                                                    'Stop
                                                    '���B�z�A�Y�w�]�j�p
                                                Else
                                                    playSound 12
                                                    Stop 'to check
                                                End If
                                            End If
                                        Else
                                            SystemSetup.playSound 12
                                            rng.Select
                                            Debug.Print e
                                            Stop 'to check
                                        End If
                                    Next e
                                End If
                        End Select
                        rng.text = vbNullString: rngClose.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                
                '�B�z�W�s��
                If VBA.Len(p.Range.text) > VBA.Len("<a href=""></a>") Then
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                    '.text = "<a href="""
                    .text = "<a "
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        url = rng.text: e = rng.text
                        Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                        rngClose.Find.Execute "</a>"
                        url = getHTML_AttributeValue("href", url)
                        'url = getHTML_AttributeValue("<a href", p.Range.text)
                        e = getHTML_AttributeValue("title", VBA.CStr(e))
                        Select Case VBA.Left(url, 1)
                            Case "#"
                                If Not SeleniumOP.IsWDInvalid() Then
                                    url = WD.url & url
                                End If
                            Case "/"
                                url = domainUrlPrefix & url '���|���h�@�ӱ׽u�]/�^�]�O�i�H���A�S�t 20241012
                            Case Else
                                If Not VBA.Left(url, 4) = "http" Then
                                    Stop 'check
                                    url = domainUrlPrefix & url
                                End If
                        End Select
                        
                        Set obj = rng.Document.Range(rng.start, rngClose.End).ShapeRange
                        rng.text = vbNullString: rngClose.text = vbNullString
                        If Not obj Is Nothing Then
                            Select Case obj.Count
                                Case 0
                                    If rng.Document.Range(rng.End, rngClose.start).text <> vbNullString Then
                                        rng.Document.Range(rng.End, rngClose.start).Hyperlinks.Add rng.Document.Range(rng.End, rngClose.start), url, , e
                                    End If
                                Case 1
                                    rng.Document.Range(rng.End, rngClose.start).Hyperlinks.Add obj(1), url, , e
                                Case Else
                                    playSound 12 'for check
                                    Stop
                            End Select
                            
                            Set obj = Nothing
                        Else
                            playSound 12 'for check
                            Stop
                        End If
                        rng.SetRange rngClose.End, p.Range.End
                    Loop
                End If

                If VBA.Len(p.Range.text) > VBA.Len("<p style=""padding-left:;>") Then
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                    .text = "<p style=""padding-left:"
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        p.Range.ParagraphFormat.IndentCharWidth 3
                        rng.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                If VBA.Len(p.Range.text) > VBA.Len("<span size=""></span>") Then
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                    .text = "<span size="
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                        rngClose.Find.text = "</span>"
                        rngClose.Find.Execute
                        
                        rng.text = vbNullString: rngClose.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                If VBA.Len(p.Range.text) > VBA.Len("<span></span>") Then
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                    .text = "<span>"
                    Do While .Execute()
                        Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                        rngClose.Find.text = "</span>"
                        rngClose.Find.Execute
                        rng.text = vbNullString: rngClose.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                If VBA.Len(p.Range.text) > VBA.Len("<p id=") Then
                    .text = "<p id="
                    rng.SetRange p.Range.start, p.Range.End
                    If .Execute() Then
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        'rng.Select
                        rng.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                    End If
                End If
                If VBA.Len(p.Range.text) > VBA.Len("<p style=""line-height:px;"">>") Then
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                    '.text = " style=""line-height: "
                    .text = "<p style="""
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        '�ɥ�url�ܼ�
                        url = getHTML_AttributeValue("style", p.Range.text)
                        arr = VBA.Split(url, ";")
                        For Each e In arr
                            e = VBA.Trim(e)
                            If VBA.Left(e, 13) = "line-height: " Then
                                rng.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
                                rng.ParagraphFormat.LineSpacing = CSng(VBA.Replace(VBA.Mid(e, VBA.Len("line-height: ") + 1), "px", vbNullString))
                            ElseIf VBA.Left(e, 11) = "font-size: " Then
                                rng.Paragraphs(1).Range.font.Size = VBA.CSng(VBA.Replace(VBA.Mid(e, VBA.Len("font-size: ") + 1), "px", vbNullString))
                            ElseIf VBA.Left(e, 11) = "margin-top:" Then
                                '���B�z
                            Else
                                If e <> vbNullString Then
                                    playSound 12
                                    rng.Select
                                    Stop 'for check
                                End If
                            End If
                        Next e
                        'url = VBA.Replace(VBA.Replace(url, "line-height: ", vbNullString), "px;", vbNullString)
                        'rng.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
                        'rng.ParagraphFormat.LineSpacing = VBA.CSng(url)
                        rng.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                If VBA.Len(p.Range.text) > VBA.Len("<p dir="""">") Then
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                    .text = "<p dir="""
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        '�ɥ�url�ܼ�
                        If VBA.InStr(rng.text, "ltr") Then rng.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                If VBA.Len(p.Range.text) > VBA.Len("<p class="";"">") Then
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                    .text = "<p class="""
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        rng.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                If VBA.Len(p.Range.text) > VBA.Len("<st1:personname ></st1:personname>") Then
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                    .text = "<st1:personname "
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                        rngClose.Find.ClearFormatting
                        rngClose.Find.Execute "</st1:personname>"
                        rng.text = vbNullString: rngClose.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                'rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                
            End With 'rng.Find
            
            
            If .Paragraphs(1).Range.text = "<br class=""Apple-interchange-newline""> " & Chr(13) Then
                .Paragraphs(1).Range.text = vbNullString
            End If
        End With 'rng
    Next p
    
    Set rng = rngHtml.Document.Range(rngHtml.start, rngHtml.End)
    ��r�B�z.FixFontname rng

    
finish:
    rngHtml.Document.Range(stRngHTML, stRngHTML).Select '�^��_�l��m
    
    Rem just for check
    With rngHtml.Find
        .ClearFormatting
        If .Execute("[<>&;]", , , True) Then
            rngHtml.Select
            SystemSetup.playSound 3
        End If
    End With
End Sub
Rem 20241011 HTML �L�ǲM�檺�B�z.Porc=Porcess
Private Sub unorderedListPorc_HTML2Word(rngHtml As Range)
    Rem �L�ǲM�檺�B�z
    Dim rngUnorderedList As Range, st As Long, ed As Long, rngUnorderedListSub As Range, p As Paragraph
    Do
        Set rngUnorderedList = GetRangeFromULToUL_UnorderedListRange(rngHtml)
        If Not rngUnorderedList Is Nothing Then
            st = rngUnorderedList.start
            Set p = rngUnorderedList.Paragraphs(1).Previous
            If Not p Is Nothing Then
                '�p�G�O���Ǻ����u���N�`���G�v
                If VBA.InStr(p.Range.text, "���N�`���G") Then
                    With rngUnorderedList.Find
                        .Execute "<li>", , , , , , , , , vbNullString, wdReplaceAll
                        .Execute "</li>", , , , , , , , , vbNullString, wdReplaceAll
                        .Execute "</ul>", , , , , , , , , vbNullString, wdReplaceAll
                         ed = rngUnorderedList.End
                    End With
                    Set rngUnorderedListSub = rngUnorderedList.Document.Range(rngUnorderedList.start, rngUnorderedList.End)
                    rngUnorderedListSub.Find.ClearFormatting
                    If rngUnorderedListSub.Find.Execute("<ul ") Then
                        rngUnorderedListSub.MoveEndUntil ">"
                        rngUnorderedListSub.End = rngUnorderedListSub.End + 2
                        If rngUnorderedListSub.Characters(rngUnorderedListSub.Characters.Count) <> Chr(13) Then
                            rngUnorderedListSub.End = rngUnorderedListSub.End - 1
                        End If
                        rngUnorderedListSub.text = vbNullString

                    Else
                        rngUnorderedListSub.SetRange rngUnorderedList.start, rngUnorderedList.End
                        If rngUnorderedListSub.Find.Execute("<ul>") Then
                            If rngUnorderedListSub.Paragraphs(1).Range.text = rngUnorderedListSub & Chr(13) Then
                                rngUnorderedListSub.Paragraphs(1).Range.text = vbNullString
                            Else
                                rngUnorderedListSub.text = vbNullString
                            End If
                        End If
                    End If
                    If rngUnorderedList.Characters(rngUnorderedList.Characters.Count) = Chr(13) Then
                        rngUnorderedList.End = rngUnorderedList.End - 1
                    End If
                    With rngUnorderedList
                        '.Hyperlinks.Add rngLink, iwe.GetAttribute("href")'�b�e���w�g���J�W�s���F
                        .Style = wdStyleHeading2 '���D 2
                        .font.Size = 18
                        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle '��涡�Z
                    End With
                Else
                    GoTo UnorderedListRange
                End If
            Else
UnorderedListRange:
                
                rngUnorderedList.Select 'for chect
                'Stop
                
                'Set rngUnorderedList = Nothing
                'InsertHTMLList rngUnorderedList.text
                
                If VBA.Left(rngUnorderedList, 5) = "<ul>" & Chr(13) And VBA.Right(rngUnorderedList, 6) = Chr(13) & "</ul>" Then
                    With rngUnorderedList
                        With .Find
                            .Execute "<li>", , , , , , , , , vbNullString, wdReplaceAll
                            .Execute "</li>", , , , , , , , , vbNullString, wdReplaceAll
                            .Execute "<ul>^p", , , , , , , , , vbNullString, wdReplaceAll
                            .Execute "^p</ul>", , , , , , , , , vbNullString, wdReplaceAll
                        End With
                        .ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
                            ListGalleries(wdBulletGallery).ListTemplates(1), ContinuePreviousList:= _
                            False, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
                            wdWord10ListBehavior
                        
                    End With
                Else
                    Stop 'for chect
                
                End If
            End If
        End If
    Loop Until rngUnorderedList Is Nothing
End Sub
Rem �ѪRHTML�ô��J�M�� 20241011 creedit_with_Copilot�j���ġGhttps://sl.bing.net/gbeqh0TAks8�GHTML����ഫ�M�ݩʳ]�m
Rem �ѪRHTML���e�A�����M�涵�ءA�M��bWord�����J�������M��˦��Chttps://sl.bing.net/bhFU3zNMSom
Sub InsertHTMLList(html As String)
    Dim doc As Document
    Dim listItems As Collection
    Dim listItem As Variant
    Dim rng As Range
    
    ' �ѪRHTML
    Set listItems = ParseHTMLList(html)
    
    ' ���J�M��
    Set doc = ActiveDocument
    Set rng = doc.Range(start:=doc.Content.End - 1, End:=doc.Content.End - 1)
    
    ' �}�l�M��
    rng.ListFormat.ApplyBulletDefault
    
    ' ��R�M�椺�e
    For Each listItem In listItems
        rng.text = StripHTMLTags(VBA.CStr(listItem))
        rng.ParagraphFormat.Alignment = wdAlignParagraphLeft
        rng.font.Name = "�з���"
        rng.InsertParagraphAfter
        Set rng = rng.Next(wdParagraph, 1) '.Range
    Next listItem
End Sub

Function ParseHTMLList(html As String) As Collection
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim listItems As New Collection
    
    ' ��l�ƥ��h��F����H
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "<li.*?>(.*?)</li>"
    
    Set matches = regex.Execute(html)
    For Each match In matches
        listItems.Add match.SubMatches(0)
    Next match
    
    Set ParseHTMLList = listItems
End Function


Rem �NHTML�奻�m�����Ϥ��A���\�h�Ǧ^�@�Ӧ��ĤF InlineShape���� 20241011 textPart:�n�ѪR��HTML�奻�Arng�G�n���J�Ϥ�����m�FdomainUrlPrefix �O�_�Ϥ����}�n�[��W�e��
Private Function insert_ImageHTML(textPart As String, rng As Range, Optional domainUrlPrefix As String) As word.inlineShape
    Dim url As String, w As Single, h As Single, align As String, hspace As String
    'url = getImageUrl(textPart)
    url = getHTML_AttributeValue("src", textPart)
    If VBA.InStr(textPart, "width") Then
        w = VBA.CSng(getHTML_AttributeValue("width", textPart))
    End If
    If VBA.InStr(textPart, "height") Then
        h = VBA.CSng(getHTML_AttributeValue("height", textPart))
    End If
    If VBA.InStr(textPart, "align") Then
        align = getHTML_AttributeValue("align", textPart)
    End If
    If VBA.InStr(textPart, "hspace") Then
        hspace = getHTML_AttributeValue("hspace", textPart)
    End If
    
    If VBA.InStr(url, "http") <> 1 Then
'        If domainUrlPrefix = vbNullString Then
'            'msgbox "���a�J����e��~��"
'            'If domainUrlPrefix = vbNullString Then domainUrlPrefix = "https://www.eee-learning.com"
'
'            'If Not SeleniumOP.IsWDInvalid() Then
'                'domainUrlPrefix = getDomainUrlPrefix(SeleniumOP.WD.url)
'            'End If
'
'        End If
        If Not IsBase64Image(url) Then 'base64�s�X���Ϥ�
            url = domainUrlPrefix & url '���|���h�@�ӱ׽u�]/�^�]�O�i�H���A�S�t 20241012
        Else
            If Base64ToImage(url, VBA.Environ("TEMP") & "\" & "tempImage.png") = False Then
                Stop
'                GoTo finish
                Exit Function
            End If
        End If
    End If
    Dim inlsp As inlineShape
    
    If Not IsBase64Image(url) Then 'VBA.InStr(url, "data:image/png;base64") = 0 Then
            'rng.InlineShapes.AddPicture fileName:=url, _
                            LinkToFile:=False, SaveWithDocument:=True
        If w > 0 And h > 0 Then
            Set inlsp = rng.InlineShapes.AddPicture(fileName:=url, _
                            LinkToFile:=False, SaveWithDocument:=True)
            resizePicture rng, inlsp, url, w, h
        Else
            On Error Resume Next
            Set inlsp = rng.InlineShapes.AddPicture(fileName:=url, _
                            LinkToFile:=False, SaveWithDocument:=True)
            On Error GoTo 0
            If Not inlsp Is Nothing Then
                If inlsp.Range.tables.Count > 0 Then
                    'resizePicture rng, inlsp, url, inlsp.Range.tables(1).PreferredWidth, inlsp.height * (inlsp.width / inlsp.Range.tables(1).PreferredWidth)
                    Rem �������óB�z�䤤���Ϥ��A���ӹw�]�N�O���j�p
                Else
                    resizePicture rng, inlsp, url
                End If
            Else
                Exit Function
            End If
        End If
    Else 'base64�s�X���Ϥ�
        
        ' ���Jbase64�s�X���Ϥ�
        Set inlsp = InsertBase64Image(url, "tempImage.png", rng)
        resizePicture rng, inlsp, url
        
    End If
    
    Rem �]�w�Ϥ��榡
    Rem inlineShape�榡
    Dim shp As Shape
    If align <> vbNullString And hspace <> vbNullString Then
        Select Case align
            Case "right"
                Set shp = inlsp.ConvertToShape
                With shp.WrapFormat
                    .Type = wdWrapSquare
                    .Side = wdWrapBoth
                    '.DistanceTop = CentimetersToPoints(0.5)
                    .DistanceLeft = CentimetersToPoints(0.5)
                    .Parent.RelativeHorizontalPosition = wdRelativeHorizontalPositionRightMarginArea
                    .Parent.RelativeVerticalPosition = wdRelativeVerticalPositionTopMarginArea
                    .Parent.Left = WdShapePosition.wdShapeRight
                    'shp.Top = WdShapePosition.wdShapeTop
    '                shp.Left = ActiveDocument.PageSetup.PageWidth - shp.width - CentimetersToPoints(1) ' �]�w�k��Z��
    '                shp.Top = CentimetersToPoints(1) ' �]�w�W��Z��
                End With
            Case "left"
                Set shp = inlsp.ConvertToShape
                With shp.WrapFormat
                    .Type = wdWrapSquare
                    .Side = wdWrapBoth
                    '.DistanceTop = CentimetersToPoints(0.5)
                    .DistanceRight = CentimetersToPoints(0.5)
                
                    .Parent.RelativeHorizontalPosition = wdRelativeHorizontalPositionLeftMarginArea
                    .Parent.RelativeVerticalPosition = wdRelativeVerticalPositionTopMarginArea
                    .Parent.Left = WdShapePosition.wdShapeLeft
    '                shp.Top = wdShapeTop
                End With
            Case "absbottom"
            Case Else
                playSound 12
                Stop 'for check
        End Select
    End If
    Rem Shape��¶�Ϯ榡
    Dim imgStyle As String, float As String, marginLeft, marginRight
    'ex: float:right;margin-left:10px;margin-right:10px;"
    imgStyle = getHTML_AttributeValue("style", textPart)
    If imgStyle <> vbNullString Then
        If inlsp.Range.tables.Count = 0 Then
            If InStr(imgStyle, "float:") Then
                float = VBA.Mid(imgStyle, VBA.InStr(imgStyle, "float:") + VBA.Len("float:"), VBA.InStr(VBA.InStr(imgStyle, "float:"), imgStyle, ";") - (VBA.InStr(imgStyle, "float:") + VBA.Len("float:")))
            End If
            If InStr(imgStyle, "margin-left:") Then
                marginLeft = VBA.Val(VBA.Mid(imgStyle, VBA.InStr(imgStyle, "margin-left:") + VBA.Len("margin-left:"), VBA.InStr(VBA.InStr(imgStyle, "margin-left:"), imgStyle, ";") - (VBA.InStr(imgStyle, "margin-left:") + VBA.Len("margin-left:"))))
            End If
            If InStr(imgStyle, "margin-right:") Then
                marginRight = VBA.Val(VBA.Mid(imgStyle, VBA.InStr(imgStyle, "margin-right:") + VBA.Len("margin-right:"), VBA.InStr(VBA.InStr(imgStyle, "margin-right:"), imgStyle, ";") - (VBA.InStr(imgStyle, "margin-right:") + VBA.Len("margin-right:"))))
            End If
            If float <> "" And VBA.IsEmpty(marginLeft) = False And VBA.IsEmpty(marginRight) = False Then
                ' �]�m�Ϥ�����¶�Ϥ覡�M����覡
                Set shp = inlsp.ConvertToShape
                With shp
                    .WrapFormat.Type = WdWrapType.wdWrapTight ' wdWrapSquare
                    Select Case float
                        Case vbNullString
                        Case "left"
                            .Left = WdShapePosition.wdShapeLeft
                            '.WrapFormat.Side = WdWrapSideType.wdWrapLeft
                        Case "right"
                            .Left = WdShapePosition.wdShapeRight
                            '.WrapFormat.Side = WdWrapSideType.wdWrapRight ' ������float:right
                        Case Else
                            Stop ' check
                    End Select
                    If marginLeft <> 0 Then
                        .WrapFormat.DistanceLeft = marginLeft ' ������margin-left:10px
                    End If
                    If marginRight <> 0 Then
                        .WrapFormat.DistanceRight = marginRight ' ������margin-right:10px
                    End If
                End With
            End If
        End If
    End If
    
    Set insert_ImageHTML = inlsp
    SystemSetup.playSound 0.411
End Function
Rem �ѪRHTML���e�A�������B��B�椸��B�Ϥ��M��r 20241011 creedit_with_Copilot�j���ġGhttps://sl.bing.net/fQ5lVr8PLye
Function ParseHTMLTable(html As String) As Collection
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim tables As New Collection
    Dim rows As New Collection
    Dim cells As New Collection
    Dim table, row
    
    ' ��l�ƥ��h��F����H
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    
    ' �ǰt���
    regex.Pattern = "<table.*?>(.*?)</table>"
    Set matches = regex.Execute(html)
    For Each match In matches
        tables.Add match.SubMatches(0)
    Next match
    
    ' �ǰt��/�C
    regex.Pattern = "<tr.*?>(.*?)</tr>"
    For Each table In tables
        Set matches = regex.Execute(table)
        For Each match In matches
            rows.Add match.SubMatches(0)
        Next match
    Next table
    
    ' �ǰt�椸��
    regex.Pattern = "<td.*?>(.*?)</td>"
    For Each row In rows
        Set matches = regex.Execute(row)
        For Each match In matches
            cells.Add match.SubMatches(0)
        Next match
    Next row
    
    Set ParseHTMLTable = cells
End Function
Rem ���U�ӡA�z�i�H�bWord���Ыت��ô��J���������e creedit_with_Copilot�j���� 20241011
Sub InsertHTMLTable(rngHtml As Range, Optional domainUrlPrefix As String)
    Dim html As String
    Dim tbl As word.table
    Dim cells As Collection
    Dim cell As Variant
    Dim row As Integer
    Dim col As Integer
    Dim img As inlineShape
    Dim rngTxt As Range
    Dim c As cell
    Dim align As String
    Dim bgcolor As String
    Dim tblWidth As Single
'    Dim imgSrc As String
'    Dim imgWidth As Single
'    Dim imgHeight As Single
    
    
'    Dim ur As UndoRecord
'    SystemSetup.stopUndo ur, "InsertHTMLTable"
    
    html = rngHtml.text
    ' �ѪRHTML
    Set cells = ParseHTMLTable(html)
    
    ' ���J���
    rngHtml.text = vbNullString
    
    Set tbl = rngHtml.tables.Add(Range:=Selection.Range, NumRows:=1, NumColumns:=1)
    
     ' �]�m����ݩ�
'    align = getHTML_AttributeValue("align", html)
'    bgcolor = getHTML_AttributeValue("bgcolor", html)
'    tblWidth = CSng(getHTML_AttributeValue("width", html))
    align = getHTML_AttributeValue("align", html)
    bgcolor = getHTML_AttributeValue("bgcolor", html)
    tblWidth = VBA.CSng(VBA.Val((getHTML_AttributeValue("width", html, ":"))))
    
    
    If align = "left" Then
        tbl.rows.Alignment = wdAlignRowLeft
    ElseIf align = "center" Then
        tbl.rows.Alignment = wdAlignRowCenter
    ElseIf align = "right" Then
        tbl.rows.Alignment = wdAlignRowRight
    End If
    
    If bgcolor <> "" Then
        If VBA.Left(bgcolor, 1) = "#" Then
            Dim arr
            arr = colorCodetoRGB(bgcolor)
            tbl.Shading.BackgroundPatternColor = RGB(arr(0), arr(1), arr(2))
        Else
            If bgcolor = "white" Then
                tbl.Shading.BackgroundPatternColor = RGB(255, 255, 255)
            Else
                playSound 12 'for check
                Stop
            End If
        End If
    End If
    
'    Dim shp As Shape
    ' �N����ഫ��Shape��H
'    Set shp = tbl.ConvertToShape
    tbl.rows.WrapAroundText = True
    ' �]�m��¶�Ϥ覡
'    shp.WrapFormat.Type = wdWrapSquare
'    shp.WrapFormat.Side = wdWrapBoth
'    shp.WrapFormat.DistanceTop = 0
'    shp.WrapFormat.DistanceBottom = 0
'    shp.WrapFormat.DistanceLeft = 0
'    shp.WrapFormat.DistanceRight = 0
    
    tbl.PreferredWidthType = wdPreferredWidthPoints
    tbl.PreferredWidth = tblWidth
    
    ' ��R��椺�e
    row = 1
    col = 1
    For Each cell In cells
        ' �ˬd�O�_�]�t�Ϥ�
        If InStr(cell, "<img") > 0 Then
            
            Set c = tbl.cell(row, col)
            Set img = insert_ImageHTML(html, c.Range, domainUrlPrefix)
'            imgSrc = getHTML_AttributeValue("src", VBA.CStr(cell))  'Mid(cell, InStr(cell, "src=") + 5, InStr(cell, """", InStr(cell, "src=") + 5) - InStr(cell, "src=") - 5)
'            imgWidth = getHTML_AttributeValue("width", VBA.CStr(cell)) 'CSng(Mid(cell, InStr(cell, "width=") + 7, InStr(cell, """", InStr(cell, "width=") + 7) - InStr(cell, "width=") - 7))
'            imgHeight = getHTML_AttributeValue("height", VBA.CStr(cell)) 'CSng(Mid(cell, InStr(cell, "height=") + 8, InStr(cell, """", InStr(cell, "height=") + 8) - InStr(cell, "height=") - 8))
'            tbl.cell(row, col).Range.InlineShapes.AddPicture fileName:=imgSrc, LinkToFile:=False, SaveWithDocument:=True
            c.Range.InlineShapes(1).width = img.width 'imgWidth
            c.Range.InlineShapes(1).height = img.height 'imgHeight
            Set rngTxt = c.Range.Document.Range(c.Range.End - 1, c.Range.End - 1)
            rngTxt.text = StripHTMLTags(VBA.CStr(cell))
        Else
            tbl.cell(row, col).Range.text = StripHTMLTags(VBA.CStr(cell))
        End If
        col = col + 1
        If col > tbl.Columns.Count Then
            tbl.rows.Add
            row = row + 1
            col = 1
        End If
    Next cell
    
'    SystemSetup.contiUndo ur
End Sub
Rem �C��X�ഫ��RGB
Private Function colorCodetoRGB(colorCode As String) As Long()
    ' �Nbgcolor�ഫ��RGB�C��
    'Dim r As Integer, g As Integer, b As Integer
    If VBA.Left(colorCode, 1) <> "#" Then Exit Function
    Dim arr(2) As Long
    arr(0) = CLng("&H" & Mid(colorCode, 2, 2))
    arr(1) = CLng("&H" & Mid(colorCode, 4, 2))
    arr(2) = CLng("&H" & Mid(colorCode, 6, 2))
    colorCodetoRGB = arr
End Function


Rem ���oHTML����檺�ݩʭ� 20241011 creedit_with_Copilot�j���ġGHTML����ഫ�M�ݩʳ]�m�GHTML����ഫ�M�ݩʳ]�m
Function getHTMLAttributeValue(attributeName As String, html As String) As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    
    ' ��l�ƥ��h��F����H
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.Pattern = attributeName & "=""'[""']"
    
    Set matches = regex.Execute(html)
    If matches.Count > 0 Then
        getHTMLAttributeValue = matches(0).SubMatches(0)
    Else
        getHTMLAttributeValue = ""
    End If
End Function
Rem �M���@����html tags HTML����
Function StripHTMLTags(html As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "<.*?>"
    regex.Global = True
    StripHTMLTags = regex.Replace(html, "")
End Function

Rem 20241010��y�� �M���b���Ҷ��S�����󤺮e��HTML�ż���
Sub RemoveEmptyTags(rng As Range)
    Dim rngOriginal As Range, arr, e
    Set rngOriginal = rng.Document.Range(rng.start, rng.End)
    arr = Array(">" & VBA.Chr(11) & "<", "><", ">" & VBA.Chr(13) & "<")
    rng.Find.MatchWildcards = False
    With rng.Find
        .ClearFormatting
        For Each e In arr
            .text = e
            .Wrap = wdFindStop
            Do While .Execute()
                Do Until rng.Characters(1) = "<"
                    rng.MoveStart , -1
                Loop
'                rng.Select 'for check
                rng.MoveEndUntil ">"
                rng.MoveEnd 1
'                rng.Select 'for check
                If Not VBA.Left(rng.text, 2) = "</" And (VBA.InStr(rng.text, e & "/") Or VBA.InStr(rng.text, ">" & VBA.Chr(13) & "/>")) _
                        And VBA.Mid(rng.text, VBA.InStr(rng.text, "/") + 1, VBA.Len(rng.text) - VBA.InStr(rng.text, "/") - 1) _
                            = rng.Document.Range(rng.start + 1, rng.start + 1 + VBA.Len(rng.text) - VBA.InStr(rng.text, "/") - 1) Then
                    rng.text = vbNullString
                End If
                If rng.Characters.Count = 1 And rng.Characters(1).text = VBA.Chr(13) And rng.Paragraphs(1).Range.Characters.Count = 1 Then
                    rng.Characters(1).text = vbNullString
                End If
                rng.Collapse wdCollapseEnd
                'rng.SetRange rng.End, rngOriginal.End
            Loop
            rng.SetRange rngOriginal.start, rngOriginal.End
        Next e
    End With
End Sub
Rem ���o�L�ǦC��]<ul></ur>�^���d�� 20241010creedit_with_Copilot�j���ġGHTML�W�s���ഫ��Word VBA�Ghttps://sl.bing.net/bXsbFqI2cz6
Function GetRangeFromULToUL_UnorderedListRange(rng As Range) As Range
    Dim startRange As Range
    Dim endRange As Range
    
    ' �d�� <ul> ����
    Set startRange = rng.Document.Range(rng.start, rng.End)
    With startRange.Find
        .ClearFormatting
        .text = "<ul"
        If .Execute Then
            startRange.Collapse Direction:=wdCollapseStart
        End If
    End With
    
    ' �d�� </ul> ����
    Set endRange = rng.Document.Range(startRange.End, rng.End)
    With endRange.Find
        .ClearFormatting
        .text = "</ul>"
        If .Execute Then
            endRange.Collapse Direction:=wdCollapseEnd
        End If
    End With
    
    ' �]�w�d��
    If Not (startRange.start = rng.start And endRange.End = rng.End) Then
        Set GetRangeFromULToUL_UnorderedListRange = rng.Document.Range(startRange.start, endRange.End)
    End If
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
    Dim xmlHttp As Object
    Dim stream As Object
    
    ' �Ы� XMLHTTP ��H
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP")
    xmlHttp.Open "GET", url, False
    xmlHttp.send '-2147467259 �L�k���X�����~�C�i��O�ѩ�z�ϥΪ��O base64 �s�X�� URL�CXMLHTTP �L�k�����B�z base64 �s�X���Ϲ��ƾ� https://sl.bing.net/dd1AOLdKBaK
    
    ' �Ы� ADODB.Stream ��H
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    stream.Write xmlHttp.responseBody
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
    Dim xmlHttp As Object
    Dim stream As Object
    
    ' �Ы� ServerXMLHTTP ��H
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    xmlHttp.Open "GET", url, False
    xmlHttp.send
    
    ' �Ы� ADODB.Stream ��H
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    stream.Write xmlHttp.responseBody
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
Function InsertBase64Image(base64String As String, filePath As String, rng As Range) As inlineShape
    Dim binaryData() As Byte
    Dim tempFilePath As String
    
    ' �ѪRbase64�s�X
    binaryData = Base64ToBinary(base64String)
    
    ' �O�s���{�ɤ��
    tempFilePath = Environ("TEMP") & "\" & filePath
    SaveBinaryAsFile binaryData, tempFilePath
    
    ' ���J�Ϥ�
    Set InsertBase64Image = rng.InlineShapes.AddPicture(fileName:=tempFilePath, LinkToFile:=False, SaveWithDocument:=True)
    base64String = tempFilePath
    ' �R���{�ɤ��
    Kill tempFilePath
End Function


Rem 20241009 ���oHTML�����ݩʤ��� pro ���]�t�u="�v
Private Function getHTML_AttributeValue(atrb As String, textIncludingAttribute As String, Optional marker As String)
    Dim lenatrb As Byte
    Select Case marker
        Case vbNullString
            atrb = atrb & "="""
        Case ":"
        atrb = atrb & ": "
    End Select
    If VBA.InStr(textIncludingAttribute, atrb) > 0 Then
        lenatrb = VBA.Len(atrb)
        getHTML_AttributeValue = VBA.Mid(textIncludingAttribute, VBA.InStr(textIncludingAttribute, atrb) + lenatrb, _
            VBA.InStr(VBA.InStr(textIncludingAttribute, atrb) + lenatrb, textIncludingAttribute, """") - (VBA.InStr(textIncludingAttribute, atrb) + lenatrb))
    End If
End Function

Rem ���J�Ϥ���A�ھګe��r���j�p�۰ʽվ�Ϥ��j�p 20241009 creedit_with_Copilot�j���ġGWordVBA �Ϥ��۰ʽվ�j�p�Ghttps://sl.bing.net/e1S3H59hvI4
Private Function getImageUrl(textIncludingSrc As String)
    getImageUrl = VBA.Mid(textIncludingSrc, VBA.InStr(textIncludingSrc, "src=""") + 5, _
        VBA.InStr(VBA.InStr(textIncludingSrc, "src=""") + 5, textIncludingSrc, """") - (VBA.InStr(textIncludingSrc, "src=""") + 5))
End Function
Function getDomainUrlPrefix(url As String)
    getDomainUrlPrefix = VBA.Left(url, VBA.InStr(url, "//")) & "/" & VBA.Mid(url, VBA.InStr(url, "//") + 2, _
                VBA.InStr(VBA.InStr(url, "//") + 2, url, "/") - (VBA.InStr(url, "//") + 2))
End Function
Rem ���s�վ�Ϥ��j�p�A�Y�L���w width�Pheight �h�Ѧҫe���r���j�p�����ȳ]�w
Private Sub resizePicture(rng As Range, pic As inlineShape, url As String, Optional width As Single = 0, Optional height As Single = 0)
    If width > 0 And height > 0 Then
        pic.width = width
        pic.height = height
    Else
    
        Dim fontSizeBefore As Single
        Dim fontSizeAfter As Single
        Dim avgFontSize As Single
        ' ����e��r���j�p
        If rng.start > 1 Then
            fontSizeBefore = rng.Characters.First.Previous.font.Size
        Else
            fontSizeBefore = rng.Characters.First.font.Size
        End If
    
        If rng.End < rng.Document.Content.End Then
            fontSizeAfter = rng.Characters.Last.Next.font.Size
        Else
            fontSizeAfter = rng.Characters.Last.font.Size
        End If
    
        ' �p�⥭���r���j�p
        avgFontSize = (fontSizeBefore + fontSizeAfter) / 2
    
        ' �վ�Ϥ��j�p
        pic.LockAspectRatio = msoTrue
        If Not IsValidImage_LoadPicture(url) Then
            pic.height = avgFontSize
            If Not SeleniumOP.IsWDInvalid() Then
                pic.Range.Hyperlinks.Add pic.Range, WD.url
            End If
        Else
            pic.height = avgFontSize * 2 ' �ھڻݭn�վ���
            pic.width = pic.height * pic.width / pic.height
        End If
    End If
End Sub

Rem 20241006 �m�ݨ�j�y�P�j�y�����˯��n Ctrl + k,d
Sub �d�ݨ�j�y�j�y�����˯�()
    ��r�B�z.ResetSelectionAvoidSymbols
    SeleniumOP.KandiangujiSearchAll Selection.text
End Sub
Rem 20241006 �˯��m�~�y�����Ʈw�n Alt + Shfit + h
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
Sub �e��j�y�Ŧ۰ʼ��I()
    'Alt + F10(���ֳt��ݽT�{�I�^
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
Sub Ū�J�j�y�Ŧ۰ʼ��I���G()
    'Ctrl + Alt + F10 �� Ctrl + Alt + F11
    If inputGjcoolPunctResult = False Then MsgBox "�Э��աI", vbCritical
End Sub
Rem 20241008 ���ѫh�Ǧ^false
Function inputGjcoolPunctResult() As Boolean
        Dim ur As UndoRecord, result As String
    ��r�B�z.ResetSelectionAvoidSymbols
    If Selection.Characters.Count < 10 Then
        MsgBox "�r�ƤӤ֡A�����n�ܡH�Цܤ֤j��10�r", vbExclamation
        Exit Function
    End If
    word.Application.ScreenUpdating = False
    Const ignoreMarker = "�m�n�q�r�u�v�y�z" '�ѦW���B�g�W���B�޸����B�z�]�ѫe�����{���X�B�z�^
    result = Selection.text
    Rem �ѦW���B�޸����B�z
    result = VBA.Replace(VBA.Replace(result, "�m", "�e"), "�n", "�f") '�ѦW����|�Q�۰ʼ��I�M���G,�H���٭� 20241001
    result = VBA.Replace(VBA.Replace(result, "�u", "�e"), "�v", "�f") '�޸���|�Q�۰ʼ��I�M���G,�H���٭� 20241001
    
    If SeleniumOP.grabGjCoolPunctResult(result, result, False) = vbNullString Then
        Selection.Document.Activate
        Selection.Document.Application.Activate
        Exit Function
    End If
    Selection.Document.Activate
    Selection.Document.Application.Activate
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
    Set rng = Selection.Document.Range(Selection.start, Selection.End)
    rng.Find.ClearAllFuzzyOptions: rng.Find.ClearFormatting
    For Each e In cln
'        If e(1) = Chr(13) Then Stop 'just for test
        If e(0) <> vbNullString Then
            If rng.Find.Execute(e(0), , , , , , True, wdFindStop) = False Then
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
            Set rng = Selection.Document.Range(rng.End, Selection.End)
        Else
            Selection.End = rng.End
        End If
    Next e
    word.Application.ScreenUpdating = True
    SystemSetup.contiUndo ur
    inputGjcoolPunctResult = True
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


Sub AppActivateDefaultBrowser()
    On Error GoTo eH
    Dim i As Byte, a
    a = Array("google chrome", "brave", "edge")
    DoEvents
    If DefaultBrowserNameAppActivate = "" Then getDefaultBrowserNameAppActivate
    AppActivate DefaultBrowserNameAppActivate
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



