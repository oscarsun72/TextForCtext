Attribute VB_Name = "NewMacros"
Option Compare Text
Option Explicit
Public DonotSave As Boolean
'�ֳt�䵧�O�GAlt+P =Shift+F5'2003/4/1
'            ��ANSI65294�]�D�^���w��Alt+.,���אּANSI8231�]�E�^

Sub HideWebBar()
If CommandBars("web").Visible Then CommandBars("web").Visible = False
End Sub

Sub CheckSaved()
Dim bkIdx As Integer
With ActiveDocument '�O���ɮס]���^�~�ˬd
'    If Not Dir(.FullName) = "" Then '�����|�]�w�x�s�����ɡ^�A���x�s2003/3/27
    If Not .path = "" Then '�P�_�O���O�s��󪺿�k�����G��
        If .Saved = False Then
'            If .Bookmarks.Count > 1 Then '�����Үɤ~��
'                For Each bk In .Bookmarks                '�N��ѫe���Цa���ҧR��'2003/3/28
'                    bkIdx = bkIdx + 1 '�O�U���ү���
'                    With bk '�p�G�O�s��B�~�B�z
'                        If InStr(1, .Name, "Edit_", vbTextCompare) > 0 Then
'                            '�p�G�O��ѫe��
'                            Do While InStr(1, bk, "Edit_", vbTextCompare) > 0 And _
'                                    CDate(Replace(Mid(bk, 6, 10), "_", "/")) <= Date - 2
'                                bk.Delete '�����H����ޭȷ|�V�e����
'                                Set bk = ActiveDocument.Bookmarks(bkIdx)
'                            Loop
'                            Exit For
'    '                    If InStr(.Name, "Edit_" & Format(Now() - 2, "yyyy_mm_dd")) Then
'    '                        .Delete
'                        End If
'                    End With
'                Next bk
'            End If
            If ActiveDocument.Application.Templates(1).Name = "�פ�.dot" Then '�פ�d���~����O���s��B����2004/2/7
                For bkIdx = 1 To .bookmarks.Count
                    With .bookmarks(bkIdx)
                        If .End >= Selection.Range.End _
                            And .start <= Selection.Range.start _
                            And InStr(.Name, Format(Date, "yyyy_mm_dd")) Then  '���O���ѫإߪ�
                            bkIdx = 0
                            Exit For
                        End If
                    End With
                Next bkIdx
                If bkIdx <> 0 Then '�קK�P�@�϶��]�w�Ӧh����
                    With .bookmarks '�s�W���Ѫ��Цa����'2003/3/28
                        .DefaultSorting = wdSortByName
                        '12�p�ɨ�G
        '                .Add Range:=Selection.Range, Name:="Edit_" & Format(Now(), "yyyy_mm_dd_AM/PM_hh_nn_ss")
                        '24�p�ɨ�G
                        .Add Range:=Selection.Range, Name:="Edit_" & _
                                Format(Format(Date, "short date"), "yyyy_mm_dd__") _
                                    & Format(Format(Time, "Short Time"), "__hh_mm_dd")
                        .ShowHidden = False
                    End With
                End If
            End If
            .Save
            .UndoClear '�M���٭�M��]�\�p���i�ٰO����
        End If
    End If
End With
End Sub

Sub CheckSavedNoClear() '2003/4/3
Dim bkIdx As Integer
With ActiveDocument '�O���ɮס]���^�~�ˬd
'    If Not Dir(.FullName) = "" Then '�����|�]�w�x�s�����ɡ^�A���x�s2003/3/27
    If Not .path = "" Then '�P�_�O���O�s��󪺿�k�����G��
        If .Saved = False Then
            For bkIdx = 1 To .bookmarks.Count
                With .bookmarks(bkIdx)
                    If .End >= Selection.Range.End _
                        And .start <= Selection.Range.start _
                        And InStr(.Name, Format(Date, "yyyy_mm_dd")) Then  '���O���ѫإߪ�
                        bkIdx = 0
                        Exit For
                    End If
                End With
            Next bkIdx
            If bkIdx <> 0 Then '�קK�P�@�϶��]�w�Ӧh����
                With .bookmarks '�s�W���Ѫ��Цa����'2003/3/28
                    .DefaultSorting = wdSortByName
                    '12�p�ɨ�G
    '                .Add Range:=Selection.Range, Name:="Edit_" & Format(Now(), "yyyy_mm_dd_AM/PM_hh_nn_ss")
                    '24�p�ɨ�G
                    .Add Range:=Selection.Range, Name:="Edit_" & _
                            Format(Format(Date, "short date"), "yyyy_mm_dd__") _
                                & Format(Format(Time, "Short Time"), "__hh_mm_dd")
                    .ShowHidden = False
                End With
            End If
            .Save
            'NoClear
'            .UndoClear '�M���٭�M��]�\�p���i�ٰO����
        End If
    End If
End With
End Sub


Sub ClearTodayBookmarks() '�N���Ѫ��Цa���ҧR��'2003/3/28
Dim bk As Bookmark, bkIdx As Integer
With ActiveDocument '�O���ɮס]���^�~�B�z
'    If Not Dir(.FullName) = "" Then '�����|�]�w�x�s�����ɡ^�A��B�z
    If Not .path = "" Then '�P�_�O���O�s��󪺿�k�����G��
        If .bookmarks.Count > 1 Then '�����Үɤ~��
            For Each bk In .bookmarks
                bkIdx = bkIdx + 1 '�O�U���ү���
                With bk '�p�G�O�s��B�~�B�z
                    If InStr(1, .Name, "Edit_", vbTextCompare) > 0 Then
                        '�R�����Ѥ��Цa����
                        Do While InStr(1, bk, Format(Date, "yyyy_mm_dd"), vbTextCompare)
                            bk.Delete '�����H����ޭȷ|�V�e����
                            Set bk = ActiveDocument.bookmarks(bkIdx)
                        Loop
'                        Exit For
                    End If
                End With
            Next bk
        End If
    End If
End With
End Sub

Sub DeleteSelBookmarks()
'���w��:Alt+Del
With Selection
    On Error GoTo 5941
    If MsgBox(.bookmarks.item(1) & "���ҡA�T�w�R���H" & _
            vbCr & vbCr & "�䤺�e���G" & .bookmarks(1).Range _
            , vbExclamation + vbOKCancel, "BookmarkID = " & .BookmarkID) = vbOK Then
        .bookmarks.item(1).Delete
    End If
Exit Sub
5941 '���X�����������s�b�I
    Select Case Err.Number
        Case 5941
            MsgBox "���B�S�����ҡI", vbExclamation
        Case Else
            MsgBox Err.Number & Err.Description
    End Select
End With
End Sub

Sub BopomofoOnlyDirect()
'
' BopomofoOnlyDirect ����
' �����إߩ� 2001/11/12�A�إߪ� �]�u�u
'

End Sub
Sub ���}() '���s��,�]���O���Ҧb��m
'Sub �}�l�s��OLE����()
'
' �}�l�s��OLE���� ����
' �������s�� 2001/11/12�A���s�� �]�u�u
'���w��:Ctrl+Alt+Q
Dim i As Byte
    On Error Resume Next
    For i = 1 To Documents.Count
        QuitClose
    Next i
    ActiveWindow.Close wdDoNotSaveChanges
    word.Application.Quit wdDoNotSaveChanges '2003/3/22
'    Selection.WholeStory
'    Selection.Delete Unit:=wdCharacter, Count:=1
End Sub
Sub OLE�ܳƧ���()
Attribute OLE�ܳƧ���.VB_Description = "�������s�� 2001/11/13�A���s�� �]�u�u"
Attribute OLE�ܳƧ���.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.OLE�ܳƧ���"
'
' OLE�ܳƧ��� ����
' �������s�� 2001/11/13�A���s�� �]�u�u
'�ֳt��:alt+w--���אּ�}�s�������w��
Dim ActCtl As Control, s As Integer, l As Integer
With ActiveDocument '2003/3/27
    If .path <> "" Then MsgBox "����󤣯�ާ@", vbExclamation: Exit Sub
'    With .Content '����
'    '     .Selection.WholeStory
'        'Selection.Copy
'        .Cut '�ŤU
'    '    Selection.Cut
'    End With
    On Error GoTo ErrH:
    AppActivate "�ϮѺ޲z", True
    blog.myaccess.screen.ActiveControl.SetFocus
    Set ActCtl = blog.myaccess.screen.ActiveControl
    Dim cName As String
    Select Case ActCtl.Parent.Name
        Case "��", "���O_��", "��_�d��", "���O_���d��"
            cName = "�q" & blog.myaccess.DLookup("�g�W", "�g", "�gID = " & _
                ActCtl.Parent.Recordset("�gID")) & "�r"
        Case "�Z�O"
            cName = "�q" & ActCtl.Parent.Recordset("�g�W") & "�r"
        Case Else
            cName = "�i�L�g�W!�j"
    End Select
    If MsgBox("�{�b�@�Τ�������O--�e" & ActCtl.Name & "�f�I" & vbCr & vbCr _
        & "�g�W�O�G " & cName & vbCr & vbCr & _
            "�n�Ѩt�Φ۰ʴM��,�Ы��e�����f", vbOKCancel + vbExclamation _
            , "�@�Τ������O: " & ActCtl.Parent.Name) = vbCancel Then
'        Stop
        For Each ActCtl In blog.myaccess.screen.activeform.Controls '2003/3/��7
        If TypeName(ActCtl) = "textbox" Then
'            If ActCtl.ControlSource = "���O" Then
'                 '�]�����O_GotFocus���{���X�A���y��SetFocus!
                If MsgBox("�{�b�@�Τ�������O--�e" & ActCtl.Name & "�f�I" & vbCr & vbCr _
                        & "�n�~��ݤU�@�ӱ���Ы��e�����f", vbOKCancel + vbExclamation) = vbOK Then
                    With .Content '����
                        .Cut '�ŤU
                    End With
                    With .Application.CommandBars("�פ�ĥ��O�s��")
                        If .Visible = True Then .Visible = False
                    End With
                    .ActiveWindow.Close wdDoNotSaveChanges
                    ActCtl.SetFocus '��o�̦ASetFocus�I
                    Exit For
                End If
'            End If
        End If
        Next ActCtl
        If ActCtl Is Nothing Then
            MsgBox "�S���A�X������A�Цۦ�M�w�I", vbExclamation
            End ' AppActivate .Application.Name
        End If
    Else
        With .ActiveWindow.Selection
            s = .start '�O�U���J�I��m'2003/3/30
            l = Len(.text)
        End With
        With .Content '����
            .Cut '�ŤU
        End With
        With .Application.CommandBars("�פ�ĥ��O�s��")
            If .Visible = True Then .Visible = False
        End With
        .ActiveWindow.Close wdDoNotSaveChanges
    End If
End With
'    Application.Visible = False
    AppActivate "�ϮѺ޲z"
'    Access.Application.SetOption "Behavior Entering Field", 0
    With ActCtl
        If .Parent.DefaultView = 0 Then 'Single Form
            .Parent.AllowEdits = True '�]�w����,�Y�D��@����˵��h�|�N�O�����ʦܲĤ@��!2002/11/28
        End If
    '    Else
    '        Screen.ActiveForm.ActiveControl.form.ActiveControl.SetFocus
    ''        DoCmd.GoToRecord Screen.ActiveForm.CurrentRecord
'        If .Name = "���O" Then
'            .Locked = False
'        End If
        If .Locked = True Then .Locked = False
        If .Parent.AllowEdits = False Then .Parent.AllowEdits = True
        If .Name <> "���O" Then
            .SetFocus
            .SelStart = 0 '����
            .SelLength = Len(.text)
        Else
             .Value = Null
        End If
        blog.myaccess.docmd.RunCommand blog.myaccess.acCmdPaste
        .SelStart = s 's + 1 '�]�w���J�I��m
        .SelLength = l
    End With
    If Windows.Count = 0 And Documents.Count = 0 Then word.Application.Quit ' wdDotNotSaveChanges    '�p�G�S�������M���}��,�~����2003/3/27
Exit Sub
ErrH:
Select Case Err.Number
    Case 5 '�{�ǩI�s�Τ޼Ƥ����T(�Y AppActivate���޼Ʀ��~!)
        On Error GoTo ErrH1
'        AppActivate "�ϮѺ޲z - [" & Screen.ActiveControl.Parent.Caption & "]"
'        DoCmd.Restore
        AppActivate blog.myaccess.CurrentObjectName
        Resume Next
    Case Else
Shows:  MsgBox Err.Number & Err.Description
End Select
Exit Sub
ErrH1:
Select Case Err.Number
    Case 5 '�{�ǩI�s�Τ޼Ƥ����T(�Y AppActivate���޼Ʀ��~!)
        AppActivate "�ϮѺ޲z - [" & blog.myaccess.screen.ActiveControl.Parent.Caption & "]"
        Resume Next
    Case Else
        GoTo Shows
End Select
End Sub
Sub OLE�ܳƧ���1() '�ƻs��ϮѺ޲z�ӫ��~��s��A����������2003/12/25
'�ֳt��:alt+w
Dim ActCtl As Control, s As Integer, l As Integer
With ActiveDocument '2003/3/27
    If .path <> "" Then MsgBox "����󤣯�ާ@", vbExclamation: Exit Sub
'    With .Content '����
'    '     .Selection.WholeStory
'        'Selection.Copy
'        .Cut '�ŤU
'    '    Selection.Cut
'    End With
    AppActivate "�ϮѺ޲z", True
    blog.myaccess.screen.ActiveControl.SetFocus
    Set ActCtl = blog.myaccess.screen.ActiveControl
    Dim cName As String
    Select Case ActCtl.Parent.Name
        Case "��", "���O_��", "��_�d��", "���O_���d��"
            cName = "�q" & blog.myaccess.DLookup("�g�W", "�g", "�gID = " & _
                ActCtl.Parent.Recordset("�gID")) & "�r"
        Case "�Z�O"
            cName = "�q" & ActCtl.Parent.Recordset("�g�W") & "�r"
        Case Else
            cName = "�i�L�g�W!�j"
    End Select
    If MsgBox("�{�b�@�Τ�������O--�e" & ActCtl.Name & "�f�I" & vbCr & vbCr _
        & "�g�W�O�G " & cName & vbCr & vbCr & _
            "�n�Ѩt�Φ۰ʴM��,�Ы��e�����f", vbOKCancel + vbExclamation _
            , "�@�Τ������O: " & ActCtl.Parent.Name) = vbCancel Then
'        Stop
        For Each ActCtl In blog.myaccess.screen.activeform.Controls '2003/3/��7
        If TypeName(ActCtl) = "textbox" Then
'            If ActCtl.ControlSource = "���O" Then
'                 '�]�����O_GotFocus���{���X�A���y��SetFocus!
                If MsgBox("�{�b�@�Τ�������O--�e" & ActCtl.Name & "�f�I" & vbCr & vbCr _
                        & "�n�~��ݤU�@�ӱ���Ы��e�����f", vbOKCancel + vbExclamation) = vbOK Then
                    With .Content '����
                        .Cut '�ŤU
                    End With
                    With .Application.CommandBars("�פ�ĥ��O�s��")
                        If .Visible = True Then .Visible = False
                    End With
                    .ActiveWindow.Close wdDoNotSaveChanges
                    ActCtl.SetFocus '��o�̦ASetFocus�I
                    Exit For
                End If
'            End If
        End If
        Next ActCtl
        If ActCtl Is Nothing Then
            MsgBox "�S���A�X������A�Цۦ�M�w�I", vbExclamation
            End ' AppActivate .Application.Name
        End If
    Else
        With .ActiveWindow.Selection
            s = .start '�O�U���J�I��m'2003/3/30
            l = Len(.text)
        End With
        With .Content '����
            .Copy '�ƻs
        End With
        With .Application.CommandBars("�פ�ĥ��O�s��")
            If .Visible = True Then .Visible = False
        End With
    End If
End With
'    Application.Visible = False
    AppActivate "�ϮѺ޲z"
'    Access.Application.SetOption "Behavior Entering Field", 0
    With ActCtl
        If .Parent.DefaultView = 0 Then 'Single Form
            .Parent.AllowEdits = True '�]�w����,�Y�D��@����˵��h�|�N�O�����ʦܲĤ@��!2002/11/28
        End If
    '    Else
    '        Screen.ActiveForm.ActiveControl.form.ActiveControl.SetFocus
    ''        DoCmd.GoToRecord Screen.ActiveForm.CurrentRecord
'        If .Name = "���O" Then
'            .Locked = False
'        End If
        If .Locked = True Then .Locked = False
        If .Name <> "���O" Then
            .SetFocus
            .SelStart = 0 '����
            .SelLength = Len(.text)
        Else
             .Value = Null
        End If
        blog.myaccess.docmd.RunCommand blog.myaccess.acCmdPaste
        .SelStart = s 's + 1 '�]�w���J�I��m
        .SelLength = l
    End With
End Sub

Sub Access����()
Attribute Access����.VB_Description = "�������s�� 2001/11/28�A���s�� �]�u�u"
Attribute Access����.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Access����"
'
' Access���� ����
' �������s�� 2001/11/28�A���s�� �]�u�u
'
    Selection.Paste
    Selection.WholeStory
    word.Application.Keyboard (1033)
    With Selection.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceAtLeast
        .LineSpacing = 0.4
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .WordWrap = True
    End With
End Sub
Sub ��ЩҦb��m����()
Attribute ��ЩҦb��m����.VB_Description = "�������s�� 2002/3/10�A���s�� �]�u�u"
Attribute ��ЩҦb��m����.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.��ЩҦb��m����"
'
' ��ЩҦb��m���� ����
' �������s�� 2002/3/10�A���s�� �]�u�u
'���w��:F5(����w�����O:EditGoTo)2004/12/14
On Error GoTo ErrH
HideWebBar

    With ActiveDocument.bookmarks
        .Add Range:=Selection.Range, Name:="���_" & Replace(Replace(Replace(Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 4), "-", "_"), "(", "�]"), ")", "�^") '�H�]���D�����Ҧ����h���l���(�n�������ɦW)2003/3/16
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
    ActiveDocument.Save
'    ���U���y��
Exit Sub
ErrH:
Select Case Err.Number
    Case 5828 '�����T�����ҦW�١C
        If MsgBox("�����T�����ҦW��,�O�_���L�H", vbOKCancel) = vbCancel Then
            Stop
            Resume
        Else
            Resume Next
        End If
    Case Else
        MsgBox Err.Number & Err.Description
End Select
End Sub
Sub �s��B����()
' 2003/3/10--��n�Z��Юɩ��@�~�o�I
    With ActiveDocument.bookmarks
        .Add Range:=Selection.Range, Name:="��Т�_" & Replace(Replace(Replace(Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 4), "-", "_"), "(", "�]"), ")", "�^")
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
    ActiveDocument.Save
End Sub

Sub ��W�@�������гB() '���w�_(�ֳt��)Alt+shift+F5 2009/5/6
Attribute ��W�@�������гB.VB_Description = "�������s�� 2002/3/12�A���s�� �]�u�u"
Attribute ��W�@�������гB.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.��W�@�������гB"
'
' ��W�@�������гB ����
' �������s�� 2002/3/12�A���s�� �]�u�u
'
CheckSaved
HideWebBar
'On Error GoTo Ftnote
'    Selection.GoTo What:=wdGoToBookmark, Name:="���_" & Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 4)
    With ActiveDocument.bookmarks
        '�p���g�K���������~�B�z�禡�F
        .item("���_" & Replace(Replace(Replace(Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 4), "-", "_"), "(", "�]"), ")", "�^")).Select
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
'Exit Sub
'Ftnote:
'Select Case Err.Number
'    Case 5678
'        With ActiveDocument.ActiveWindow.View
'            If .Type = wdNormalView Then
'            On Error GoTo Comts
'                .SplitSpecial = wdPaneFootnotes '���}�˵�,wdPaneComments �����˵�
'            Else
'                .SeekView = wdSeekFootnotes
'            End If
'        End With
'        Resume
'    Case Else
'        MsgBox Err.Number & Err.Description
'End Select
'Exit Sub
'Comts:
'Select Case Err.Number
'    Case 4198
'        With ActiveDocument.ActiveWindow.View
'            If .Type = wdNormalView Then
'                .SplitSpecial = wdPaneComments
'            Else
'                .SeekView = wdSeekFootnotes
'            End If
'        End With
'        Resume
'    Case Else
'        MsgBox Err.Number & Err.Description
'End Select
End Sub

Sub ��s��B_���1()
CheckSaved
'On Error GoTo Ftnote
'    Selection.GoTo What:=wdGoToBookmark, Name:="��Т�_" & Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 4)
    With ActiveDocument.bookmarks
        '�p���g�K���������~�B�z�禡�F
        .item("��Т�_" & Replace(Replace(Replace(Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 4), "-", "_"), "(", "�]"), ")", "�^")).Select
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
'Exit Sub
'Ftnote:
'Select Case Err.Number
'    Case 5678
'        With ActiveDocument.ActiveWindow.View
'            If .Type = wdNormalView Then
'            On Error GoTo Comts
'                .SplitSpecial = wdPaneFootnotes '���}�˵�,wdPaneComments �����˵�
'            Else
'                .SeekView = wdSeekFootnotes
'            End If
'        End With
'        Resume
'    Case Else
'        MsgBox Err.Number & Err.Description
'End Select
'Exit Sub
'Comts:
'Select Case Err.Number
'    Case 4198
'        With ActiveDocument.ActiveWindow.View
'            If .Type = wdNormalView Then
'                .SplitSpecial = wdPaneComments
'            Else
'                .SeekView = wdSeekFootnotes
'            End If
'        End With
'        Resume
'    Case Else
'        MsgBox Err.Number & Err.Description
'End Select
End Sub

Sub QuitClose()
If Documents.Count = 0 Then word.Application.Quit wdDoNotSaveChanges: Exit Sub
With ActiveDocument '2003/3/25��}
If Left(.Name, 2) = "���" And _
    IsNumeric(Mid(.Name, 3)) And _
        .AttachedTemplate.Name = "Normal.dot" Then
        If .Windows.Count = 1 Then
            DonotSave = True
            .Close wdDoNotSaveChanges
        Else
            .ActiveWindow.Close
        End If
        With word.Application
            If Documents.Count = 0 Then
                .CommandBars("�פ�ĥ��O�s��").Visible = False
'                .Position = msoBarTop
                .Quit wdPromptToSaveChanges
                End 'Exit Sub
            End If
        End With
Else
    If .Windows.Count = 1 Then
            DonotSave = True
            .Close wdDoNotSaveChanges
    Else
        If .Saved = False Then
            Select Case MsgBox("���w�ܧ�A�O�_�n�b�������x�s�H" & vbCr & vbCr _
                & "���x�s�Ы��_!", vbExclamation + vbYesNoCancel)
                Case vbYes
                    .Save
                    .ActiveWindow.Close
                Case vbCancel
                    .ActiveWindow.Close
                    .Save
            End Select
        Else
            DonotSave = True
            .ActiveWindow.Close wdDoNotSaveChanges
        End If
    End If
End If
DonotSave = False
End With
���U���y��
End Sub

Sub �b�ϮѺ޲z���M�����r��() '��W�u�M�����r��v
Dim Mystr As String, ctl As Control, ctlSourceName As String ', f As Byte '�ֳt��:Alt+Z
Dim C As Integer
CheckSaved

With Selection
If .Type = wdSelectionIP Then MsgBox "�п���Q�n�M�䤧��r", vbExclamation: Exit Sub
If .Type = wdSelectionNormal Then ' <> wdNoSelection OR wdSelectionIP Then '�������J�I
'If .Text <> "" Then
    If VBA.right(.Range, 1) Like Chr(13) Then
        Mystr = Mid(.Range, 1, Len(.Range) - 1)
    Else
        Mystr = .Range
    End If
    Mystr = Replace(Mystr, Chr(13), Chr(13) & Chr(10)) '.Text'�]��Access�PWord����Ҧs���Ȥ��P!
    Mystr = Replace(Mystr, Chr(11), Chr(13) & Chr(10)) '.Text'�]��Access�PWord����Ҧs���Ȥ��P!
'    .Font.Color = wdColorRed
'    .Collapse wdCollapseEnd
    On Error GoTo �Ƶ�
'    setOX
'    OX.WinActivate "�ϮѺ޲z"
    AppActivate "�ϮѺ޲z"
    If myaccess Is Nothing Then
        Set myaccess = GetObject("D:\�d�{�@�o�N\���y���\�ϮѺ޲z.mdb")
    End If
'    AppActivate "�ϮѺ޲z" ', True '�]�����ɭn�x�s���A�i�භ����Word�o��J�I�B�z����C
    If myaccess.CurrentObjectName = "���" Then myaccess.docmd.RunCommand blog.myaccess.acCmdWindowHide ' Screen.ActiveForm.Visible = False
'cl: For Each Ctl In Screen.ActiveForm.Controls '2003/3/17
'        If TypeName(Ctl) = "textbox" Then
'            If Ctl.ControlSource Like "[���Z]�O" Then '= "���O" Then
'                Ctl.SetFocus
'                ctlSourceName = Ctl.ControlSource
'                Exit For
'            End If
'        End If
'    Next Ctl
cl: For C = 0 To myaccess.screen.activeform.Controls.Count - 1 '2006/4/21
        Set ctl = myaccess.screen.activeform.Controls(C)
        If TypeName(ctl) = "textbox" Then
            If ctl.ControlSource Like "[���Z]�O" Then '= "���O" Then
                ctl.SetFocus
                ctlSourceName = ctl.ControlSource
                Exit For
            End If
        End If
    Next C

'    If TypeName(Ctl) = "Nothing" Then
    If ctl Is Nothing Then
'    If TypeName(Ctl) = "textbox" Then
'        If Ctl.ControlSource <> "���O" Then
            If MsgBox("�S��[���O]�����!" & vbCr & vbCr & "�O �_ �n �� �� �L �� ��H" _
                , vbExclamation + vbYesNo) = vbYes Then
                'f = 1
                GoTo ��L���
            Else
                End
    '            Exit Sub
            End If
    End If
'    DoEvents
    With ctl.Parent
        If .Dirty = True Then 'Ctl.Parent.Refresh
            If .AllowEdits = False Then
                .AllowEdits = True
                myaccess.docmd.RunCommand blog.myaccess.acCmdSaveRecord
                .AllowEdits = False
            Else
                myaccess.docmd.RunCommand blog.myaccess.acCmdSaveRecord
            End If
        End If
''        '.RecordsetClone.FindFirst ctlSourceName & " like " & Chr$(34) & "*" _
''         & Mystr & "*" & Chr$(34) & ""
'        If InStr(.Recordset.Fields(ctlSourceName), Mystr) = 0 Then
'            With .RecordsetClone
'                Do
'                    .MoveNext
'                    If .EOF Then Exit Do
'                Loop While InStr(.Fields(ctlSourceName), Mystr) = 0
'            End With
'        End If
        Dim ff As Boolean
        ff = myaccess.Run("�M��r��_ole��", .Recordset, ctlSourceName, Mystr) '.Recordset, ctlSourceName, Mystr)
'    If .RecordsetClone.NoMatch Then
    If ff Then
        Select Case MsgBox("�䤣��!!" & vbCr & vbCr & "�O �_ �n �� �� �L �� ��H" _
                & vbCr & vbCr & "���n�����Ы��e�����f!", vbInformation _
                    + vbYesNoCancel, "�ثe���O�G " & myaccess.screen.activeform.Name)
            Case Is = vbYes
            'f = 2
                GoTo ��L���
            Case vbCancel
'                AppActivate .Application.Caption, True
'                .Application.ActiveWindow.Activate
                End
            Case vbNo
                Selection.Copy '�ƻs�H�ƶK�W��!2003/3/28
                AppActivate "�ϮѺ޲z"
                myaccess.Forms(0).SetFocus
                End
        End Select
    Else
'        .Recordset.Bookmark = .RecordsetClone.Bookmark
''        Ctl.Parent.Recordset.Bookmark = Ctl.Parent.RecordsetClone.Bookmark
'        If Ctl.Parent.CurrentRecord > Ctl.Parent.RecordsetClone.AbsolutePosition + 1 Then
'            myaccess.DoCmd.FindRecord Mystr, acAnywhere, True, acUp, True, , False
''        '    On Error Resume Next
'        ElseIf Ctl.Parent.CurrentRecord < Ctl.Parent.RecordsetClone.AbsolutePosition + 1 Then
'            myaccess.DoCmd.FindRecord Mystr, acAnywhere, True, acDown, True, , False
'        End If
'        With myaccess.Forms("��").���O
'            .SelStart = InStr(.Value, Mystr) - 1
'            .SelLength = Len(Mystr)
'        End With
        With ctl
            If Not .SelText Like Mystr Then
                If InStr(.Value, Mystr) <> 0 Then
                    .SelStart = InStr(.Value, Mystr) - 1
                    .SelLength = Len(Mystr)
                End If
            End If
        End With
        Beep '�M�䧹���T����
    End If
    End With
End If
End With
Exit Sub
�Ƶ�:
'Stop '�ˬd��
Select Case Err.Number
    Case 5 '�I�s�޼Ƥ����T'�{�ǩI�s�Τ޼Ƥ����T(�Y AppActivate���޼Ʀ��~!)
'        MsgBox Err.Number & Err.Description
        If myaccess.CurrentProject.AllForms("���").IsLoaded Then blog.myaccess.Forms("���").Visible = False 'docmd.Close acForm,"���",acSaveNo
        AppActivate myaccess.CurrentObjectName
'        AppActivate "�ϮѺ޲z - [" & Screen.ActiveControl.Parent.Caption & "]"
'        AppActivate "�ϮѺ޲z"
        Resume Next
'        MsgBox "�ϮѺ޲z��Ʈw�S���}�ҡI", vbExclamation
'        End
    Case 2475
        Select Case MsgBox("�ϮѺ޲z��Ʈw�S���@�Τ������I" & vbCr & "�O�_�n�Ѩt�ΨӴM��H", vbExclamation + vbOKCancel)
            Case Is = vbOK
'                Raise
'��L���:        Dim frm As form
��L���:       Dim i As Integer '�n�̶}�ҥ���˵ۧ�~�X�ާ@�ߺD2003/3/18
'                For Each frm In access.Forms
'                    If MsgBox("�ثe�@�Τ������O�G " & frm.Name & Space(3) & vbCr & vbCr _
                        & "�n �� �U �@ �� �� �� �ܡH", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                        
'                For i = 0 To Access1.CurrentProject.AllForms.Count
'                    If CurrentProject.AllForms(i).IsLoaded Then Forms(i).SetFocus
'                    Exit For
'                Next
                For i = myaccess.Forms.Count - 1 To 0 Step -1
                    If blog.myaccess.Forms(i).Name <> blog.myaccess.CurrentObjectName And blog.myaccess.Forms(i).Visible Then  ' Screen.ActiveForm.Name Then
                       If myaccess.Forms(i).Name <> "���y" Then
                            If MsgBox("�ثe�@�Τ������O�G " & myaccess.Forms(i).Name & space(3) & vbCr & vbCr _
                                & "�n �� �U �@ �� �� �� �ܡH", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                                AppActivate "�ϮѺ޲z"
    '                            frm.SetFocus
                                myaccess.Forms(i).SetFocus
                                myaccess.docmd.Restore
        '                        If Ctl Is Nothing Then Set Ctl = frm.ActiveControl
                                If ctl Is Nothing Then Set ctl = myaccess.Forms(i).ActiveControl
                                If ctl.Parent.Dirty = True Then ctl.Parent.Refresh
        ''                        If Not IsEmpty(f) Then GoTo cl
                                If Err.Number = 0 Then GoTo cl '�S���~�A��ܦb�M����A������A�n���s�䱱��]Cl)2003/3/22
                                Exit For
                            End If
                        End If
                    End If
                    If i = 0 Then MsgBox "�w�s�������A�S���z�A�X�����A�Цۦ�ܿ�..", vbExclamation: End
                Next
                GoTo cl '�䧹���᭫�s�M��!2005/3/24
'                    If frm Is Nothing Then MsgBox "�w�s�������A�S���z�A�X�����A�Цۦ�ܿ�..", vbExclamation: End
            Case Is = vbCancel
                End
        End Select
    Case 20 '�^�_�B�L���~�I
        Resume Next
    Case 2137 '�ثe����M��,�h�A��!2005/3/17
        DoEvents
'        AppActivate ActiveWindow.Application
'        If MsgBox("�ثe����j�M,�O�_�A��?", vbOKCancel + vbExclamation) = vbOK Then
            AppActivate myaccess.CurrentObjectName
            myaccess.screen.ActiveControl.Parent.SetFocus
            With myaccess.screen.ActiveControl
                If .ControlSource = "���O" Then
                    .SetFocus: Resume
                Else
                    Stop
                End If
            End With
'        End If
    Case 2110 '���ಾ�ʨ쥾�O���'2005/4/1
'        Dim c As Control
'        For Each c In Screen.ActiveControl.Parent
'            If TypeName(c) = "Textbox" Then '��r���
'                If c.ControlSource = "���O" Then c.SetFocus
'            End If
'        Next
        Resume
    Case Else
        MsgBox Err.Number & Err.Description
        End
End Select
'Resume
End Sub
Sub �b����󤤨��N����r��榡_���N�����r()
With ActiveWindow.Selection '�ֳt��G
If .Type = wdSelectionIP Then MsgBox "�п���Q�n�M�䤧��r", vbExclamation: Exit Sub
If .Type = wdSelectionNormal Then ' <> wdNoSelection OR wdSelectionIP Then '�������J�I
    If InStr(ActiveDocument.Content, .text) = InStrRev(ActiveDocument.Content, .text) Then
        MsgBox "����u�����B!", vbInformation
        .Font.Color = wdColorRed
        .Font.Bold = True
        Exit Sub
    End If
    .Find.ClearFormatting
    .Find.Replacement.Font.Color = wdColorRed
    .Find.Replacement.Font.Bold = True
    .Find.Execute FindText:=.text, MatchCase:=True, Replace:=wdReplaceAll, Replacewith:=.text, Wrap:=wdFindContinue
    '�@�w�n��Wrap:=wdFindContinue�_�h��V��M��,�w�]�Ȭ�Wrap:=wdFindStop
End If
End With
End Sub
Sub �b����󤤨��N����r��榡()
With Selection '�ֳt��G
If .Font.Color = wdColorAutomatic Or .Font.Color = wdColorBlack Then _
    If MsgBox("�Х����w�r�Φ�m!" & vbCr & _
        "�Y�n�O�d�¦r�Ы��e�����f", vbExclamation + vbOKCancel) = vbOK Then Exit Sub
If .Type = wdSelectionIP Then MsgBox "�п���Q�n�M�䤧��r", vbExclamation: Exit Sub
If .Type = wdSelectionNormal Then ' <> wdNoSelection OR wdSelectionIP Then '�������J�I
    If InStr(ActiveDocument.Content, .text) = InStrRev(ActiveDocument.Content, .text) Then MsgBox "����u�����B!", vbInformation: Exit Sub
    .Find.ClearFormatting
    .Find.Replacement.ClearFormatting
'    Dim aFont
'    Set aFont = .Words.First.Font.Duplicate
'    .Find.Replacement.Font = aFont
    
    .Find.Replacement.Font.Color = .Font.Color
    .Find.Replacement.Font.Bold = .Font.Bold
    .Find.Replacement.Font.Italic = .Font.Italic
    .Find.Replacement.Font.Size = .Font.Size
    .Find.Replacement.Font.Name = .Font.Name
    .Find.Replacement.Font.NameAscii = .Font.NameAscii
    .Find.Replacement.Font.Underline = .Font.Underline
    .Find.Replacement.Font.Borders = .Font.Borders
    .Find.Replacement.Font.Outline = .Font.Outline
    .Find.Replacement.Font.position = .Font.position
    .Find.Replacement.Font.Animation = .Font.Animation
    .Find.Replacement.Font.Spacing = .Font.Spacing
    .Find.Replacement.Font.EmphasisMark = .Font.EmphasisMark
    .Find.Replacement.Font.Emboss = .Font.Emboss
    .Find.Replacement.Font.Engrave = .Font.Engrave
    .Find.Replacement.Font.Hidden = .Font.Hidden
    .Find.Replacement.Font.ItalicBi = .Font.ItalicBi
    .Find.Replacement.Font.Kerning = .Font.Kerning
    .Find.Replacement.Font.NameFarEast = .Font.NameFarEast
    .Find.Replacement.Font.NameOther = .Font.NameOther
    .Find.Replacement.Font.Scaling = .Font.Scaling
'    .Find.Replacement.Font.Shading = .Font.Shading
    .Find.Replacement.Font.Shadow = .Font.Shadow
    .Find.Replacement.Font.SizeBi = .Font.SizeBi
    .Find.Replacement.Font.Subscript = .Font.Subscript
    .Find.Replacement.Font.Superscript = .Font.Superscript
    .Find.Replacement.Font.UnderlineColor = .Font.UnderlineColor
    .Find.Execute FindText:=.text, MatchCase:=True, Replace:=wdReplaceAll, Replacewith:=.text, Wrap:=wdFindContinue
    '�@�w�n��Wrap:=wdFindContinue�_�h��V��M��,�w�]�Ȭ�Wrap:=wdFindStop
End If
End With
End Sub
Sub BopomofoWithBlankCharDirect()
'
' BopomofoWithBlankCharDirect ����
' �����إߩ� 2002/11/10�A�إߪ� �]�u�u
'

End Sub
Sub �r���ഫ_�رd�ײʶ�()
Attribute �r���ഫ_�رd�ײʶ�.VB_Description = "�������s�� 2003/1/10�A���s�� �]�u�u"
Attribute �r���ഫ_�رd�ײʶ�.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.�r���ഫ_�رd�ײʶ�"
'���w��GShift+Alt+W
' �r���ഫ_�رd�ײʶ� ����
' �������s�� 2003/1/10�A���s�� �]�u�u
'
Static fontName As String

CheckSaved

If Selection.Font.Name <> "�رd�ײʶ�" Then
    fontName = Selection.Font.Name
    Selection.Font.Name = "�رd�ײʶ�"
Else '�_�쬰����r��2003/1/11(�Ϋ����w��:Ctrl+Spacebar--���u�W����)2003/1/12
    Selection.Font.Name = fontName
End If
End Sub

Sub Timer()
'Application.OnTime When:=Now + TimeValue("00:00:10"), _
    Name:="Timer"
word.Application.OnTime When:=Now + TimeValue("00:10:00"), _
    Name:="��ЩҦb��m����" '"Project1.Module1.Macro1"
Stop '�ˬd��
End Sub

Sub �˵���Ʈw���() '2003/2/10
Dim SearchedText As String '���w��GAlt+Q '2009.9.17��Ʈw�w�e�j,���A�Ψo!
On Error GoTo errs
'�]���}�Ҹ�Ʈw�Өt�θ귽�A�����ɭP����A�]�������x�s�I2003/3/26��
CheckSaved

With Selection
If .Type = wdSelectionIP Then MsgBox "�п���Q�n�M�䤧��r", vbExclamation: Exit Sub
If .Type = wdSelectionNormal Then ' <> wdNoSelection OR wdSelectionIP Then '�������J�I
'If SearchedText <> "" Then
'    SearchedText = .Text'����H�U��:2004/11/11
    SearchedText = Replace(.Range, Chr(13), Chr(13) & Chr(10)) '.Text'�]��Access�PWord����Ҧs���Ȥ��P!
    SearchedText = Replace(SearchedText, Chr(11), Chr(13) & Chr(10)) '.Text'�]��Access�PWord����Ҧs���Ȥ��P!
    Dim Access As Object
    Set Access = CreateObject("access.application")
'    '�H�W�@��έ�ӤU�����@��.OpenCurrentDatabase�i�H�令�p�U�@��,�\�p�U�@��u�|�}�Ҥ@��Access(���ް���X��); _
    �Y�ӭ�ӤW������A���}��Ʈw,�h�C���Y�|�}�Ҥ@��2003/12/14
'    Set access = GetObject("d:\�d�{�@�o�N\���y���\�ϮѺ޲z_�d�ߪ�.mdb")
    Access.UserControl = True
''    access.UserControl = False '�p�G��False�ٷ|�ϴM��U�@�Ӧܥ����ɡA���|��ܰT�����2003/3/25
'    access.Visible = True
    Access.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\�ϮѺ޲z_�d�ߪ�.mdb"
    Access.docmd.OpenForm "��_�d��", , , , , , "Word"
''    access.DoCmd.Close acForm, "�D���", acSaveNo
''    access.Forms("��_�d��").RecordSource = "��_�d��"
    Access.Forms("��_�d��").Controls("����r").SetFocus
''    SendKeys Selection.Text & "{tab}"
    Access.Forms("��_�d��").Controls("����r").text = SearchedText
''    access.Forms("��_�d��").Controls("����r") = SearchedText
''    SendKeys "~"
    Set Access = Nothing
End If
Exit Sub
errs:
    Select Case Err.Number
        Case 7866 '�w�}��
            Select Case MsgBox("�w�}�ҡA�O�_�n�����A�~��H", vbYesNoCancel + vbExclamation)
                Case vbYes
                    Dim Access1 As Object
                    Set Access1 = GetObject("d:\�d�{�@�o�N\���y���\�ϮѺ޲z_�d�ߪ�.mdb")
'                    Set access = Access1
'                    If Access1.Visible = False Then Access1.Visible = True
'                    If Access1.CurrentProject.AllForms("��_�d��").IsLoaded Then Access1.DoCmd.Close acForm, "��_�d��", acSaveNo
'                    If Access1.CurrentProject.AllForms("�D_�l�D�˵�").IsLoaded Then Access1.DoCmd.Close acForm, "�D_�l�D�˵�", acSaveNo
'                    Access1.CloseCurrentDatabase
                    Access1.Application.Quit blog.myaccess.acExit
'                    Access1.Quit
                    Set Access1 = GetObject(, "access.application")
                    Access1.Quit blog.myaccess.acExit
                    Set Access1 = Nothing
                    Set Access = CreateObject("access.application")
                    Access.UserControl = True
                    Access.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\�ϮѺ޲z_�d�ߪ�.mdb"
                    Resume Next
                Case vbNo
                    Access.Quit blog.myaccess.acExit '��s�}��Access����
'                    ������o�ѷ�
                    Set Access = GetObject("d:\�d�{�@�o�N\���y���\�ϮѺ޲z_�d�ߪ�.mdb")
'                    Set access = CreateObject("d:\�d�{�@�o�N\���y���\�ϮѺ޲z_�d�ߪ�.mdb")
'                    access.UserControl = True
'                    If access.Visible = False Then access.Visible = True
'                    If Access1.CurrentProject.AllForms("��_�d��").IsLoaded Then Access1.DoCmd.Close acForm, "��_�d��", acSaveNo
'                    If Access1.CurrentProject.AllForms("�D_�l�D�˵�").IsLoaded Then Access1.DoCmd.Close acForm, "�D_�l�D�˵�", acSaveNo
                    Resume Next
                Case Else
    '                Set access = GetObject("access.application")
    '                access.CloseCurrentDatabase
                    Access.Quit   '2003/2/22
                    Set Access = Nothing
    '                Stop
                    blog.myaccess.Application.DDETerminateAll
                    End
            End Select
        Case 4605 '����k���ݩʵL�k�ϥΡA�]�� �o�������D���b�s����w���A��.
            On Error Resume Next
            If Selection.Information(wdInMasterDocument) Then  '�p�G�O�D�����
            '�Ϊ̼g��:
'            If ActiveDocument.IsMasterDocument = True Then
                Dim subdoc, wins
                For Each subdoc In ActiveDocument.Subdocuments
                    If subdoc.Locked Then
                        If InStr(subdoc.Name, "�۰ʦ^�_") Then
                            Documents(Mid(subdoc.Name, 7, Len(subdoc.Name) - 10)).Activate
'                            subdoc.Locked = False
'                        Else
'                            If MsgBox("�l���w�}��,�Х������l���,�A�ާ@!", vbExclamation + vbOKCancel) = vbOK Then
'                            Documents(Mid(subdoc.Name, 7, Len(subdoc.Name) - 10)).Close
'                            End If
                        Else
                            For Each wins In Windows
                                If wins.Caption = Left(subdoc.Name, Len(subdoc.Name) - 4) Then
                                    Documents(subdoc.Name).Activate
                                Else
                                    subdoc.Locked = False
'                                    subdoc.Parent.Undo
                                End If
                            Next wins
                        End If
                    End If
                Next subdoc
                Resume
'            If ActiveDocument.Subdocuments(3).Locked Then ActiveDocument.Subdocuments(1).Locked = False
'            if ActiveDocument.Subdocuments
            Else
                MsgBox Err.Number & ":" & Err.Description, vbExclamation
            End If
        Case Else
            MsgBox Err.Number & ":" & Err.Description, vbExclamation
    End Select
End With
End Sub
Sub �˵���Ʈw�D�l�D() '2003/2/10
CheckSaved

If Selection.Type <> wdSelectionIP Or wdNoSelection Then
'If Selection.Text <> "" Then
    Dim Access As Object
    Set Access = CreateObject("access.application")
    Access.UserControl = True
    Access.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\�ϮѺ޲z_�d�ߪ�.mdb"
    Access.docmd.OpenForm "�D_�l�D�˵�", , , , , , "Word"
'    access.DoCmd.Close acForm, "�D���", acSaveNo
'    access.Forms("�D_�l�D�˵�").SetFocus
    Access.Forms("�D_�l�D�˵�").Controls("Text18").SetFocus
'    access.Forms("�D_�l�D�˵�").Controls("Text18").Text = Selection.Text
    SendKeys Selection.text '�W�@��L��
    Set Access = Nothing
End If
End Sub

Sub �s�W����()
CheckSaved

If Selection.Type <> wdSelectionIP Or wdNoSelection Then
�b�ϮѺ޲z���M�����r��
0   Select Case blog.myaccess.screen.activeform.Name '2003/3/30
        Case "��", "���O_��", "��_�d��"
            blog.myaccess.Forms(blog.myaccess.screen.activeform.Name).Label19_Click
        Case "���O_���d��"
            blog.myaccess.docmd.RunMacro "�O���B�z.�s�W����"
'    Else
'        If MsgBox("�S�� [���O] ���!   �L�k�s�W����..." & vbCr & vbCr & "�O �_ �n �� �� �L �� ��H" _
'                    , vbExclamation + vbYesNo) = vbNo Then End 'Exit Sub
'        Dim i As Integer '�n�̶}�ҥ���˵ۧ�~�X�ާ@�ߺD2003/3/18
'    '    For Each frm In Forms
'        For i = Forms.Count - 1 To 0 Step -1
'            If MsgBox("�ثe�@�Τ������O�G " & vbCr & Forms(i).Name & Space(3) & vbCr & vbCr _
'                & "�n �� �U �@ �� �� �� �ܡH", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
'    '                    If CurrentProject.AllForms(Forms(i).Name).IsLoaded Then'forms�����h�w���b�w�}�Ҫ����F�I
'                AppActivate "�ϮѺ޲z"
'                Forms(i).SetFocus
'                DoCmd.Restore '
'                GoTo 0
'                Exit For
'            End If
'            If i = 0 Then MsgBox "�w�s�������A�S���z�A�X�����A�Цۦ�ܿ�..", vbExclamation: End
'        Next i
    End Select
    On Error GoTo e
    AppActivate "�ϮѺ޲z"

    
''If Selection.Text <> "" Then
''    Dim access As Object
''    Set access = CreateObject("access.application")
''    access.UserControl = True
''    access.OpenCurrentDatabase "d:\�d�{�@�o�N\���y���\�ϮѺ޲z_�d�ߪ�.mdb"
''    Set access = GetObject("d:\�d�{�@�o�N\���y���\�ϮѺ޲z_�d�ߪ�.mdb")
'    AppActivate "�ϮѺ޲z"
'    If Not CurrentProject.AllForms("��").IsLoaded And Not CurrentProject.AllForms("���O_��").IsLoaded Then Exit Sub
'    If CurrentProject.AllForms("��").IsLoaded Then
'        Forms("��").SetFocus
'        Forms("��").label19_Click
'    End If
'    If CurrentProject.AllForms("���O_��").IsLoaded Then
'        Forms("���O_��").SetFocus
'        Forms("���O_��").label19_Click
'    End If
''            access.Forms("��").Controls("Text18").SetFocus
'''    access.Forms("�D_�l�D�˵�").Controls("Text18").Text = Selection.Text
''    SendKeys Selection.Text '�W�@��L��
''    Set access = Nothing
End If
Exit Sub
e:
Select Case Err.Number
    Case 5 '�{�ǩI�s�Τ޼Ƥ����T(�Y AppActivate���޼Ʀ��~!)
        AppActivate "�ϮѺ޲z - [" & blog.myaccess.screen.ActiveControl.Parent.Caption & "]"
        Resume Next
    Case Else
        MsgBox Err.Number & Err.Description
End Select
End Sub

Public Sub �R���w�\���O()
'���w��GCtrl+Shift+Del'�ΧR����󤺩Ҧ�����d�򪺤�r
With Selection '�p�G�S����h�����
    If Not .Type = wdSelectionNormal Then
        .HomeKey unit:=wdStory, Extend:=wdExtend
        .Delete
    Else
        ActiveDocument.Range.Find.Execute Selection.text, , , , , , True, wdFindContinue, , "", wdReplaceAll
    End If
End With
End Sub

Public Sub �ŤU�H�K�W���O()
'���w��GCtrl+Num 0

Dim p As Byte
With Selection '�p�G�S����h�����
    If .Document.path <> "" Then MsgBox "����󤣯�ާ@", vbExclamation: Exit Sub
    If .Type = wdSelectionIP Then
        p = MsgBox("�T�w�u�󭶡v�ܡH", vbQuestion + vbYesNoCancel)
        If p = vbNo Then End 'Exit Sub'2003/11/20
        .HomeKey unit:=wdStory, Extend:=wdExtend
        .Cut
        Do While Asc(.text) = 10 Or Asc(.text) = 13
        If .End + 1 = .Document.Content.End Then Exit Do
        .Delete
        Loop
'        AppActivate "�ϮѺ޲z", True
        AppActivate blog.myaccess.CurrentObjectName, True
        With blog.myaccess.screen.activeform
            If .RecordSource <> "��_" Then Exit Sub
            If Not .NewRecord Then Exit Sub
            If IsNull(.Controls("��").DefaultValue) Or .Controls("��").DefaultValue = "" Then Exit Sub
            .Controls("��") = .Controls("��").DefaultValue '���O���|�]�wSetFocus�G���p��
'            If .ActiveControl.ControlSource <> "���O" Then _
'                .Controls("���O").SetFocus
            Do Until .ActiveControl.ControlSource = "���O"
                .Controls("���O").SetFocus
            Loop '2003/11/22
            blog.myaccess.docmd.RunCommand blog.myaccess.acCmdPaste
            If p = vbYes Then .Controls("��") = True
            blog.myaccess.docmd.GoToRecord blog.myaccess.acDataForm, .Name, blog.myaccess.acNewRec
'            Do While .ActiveControl.ControlSource <> "���O"'�]���K�W�e�w���ˬd�A�G���A�E�J�F2004/8/27
'                Screen.PreviousControl.SetFocus
'            Loop
        End With
    Else
        MsgBox "�иm�J���J�I", vbExclamation
    End If
'    If Len(.Document.Content) = 1 Then QuitClose
    AppActivate .Application
    '���ӹڱM�ΡG�}������'2003/11/23
    If Len(.Document.Content) > 1 Then
'        SendKeys "^f", True     '�}�ҴM����
    Else
''        SendKeys "^o", True
'        With Forms("�g")
'            .SetFocus
'            If IsNull(.Controls("��")) Then .Controls("��") = InputBox("�п�J����!!")
'        End With
'        With Selection
'            .Application.Documents.Open Dir("D:\�d�{�@�o�N\��Ʈw\��r�ɸ�Ʈw\�p��\���ӹ�" & "\*" & Screen.ActiveForm.Controls("��") + 1 & "*")
'            .Application.Documents(1).Activate
'            With Selection
'                .EndKey wdStory, wdMove
'                .HomeKey wdStory, wdExtend
'                .Copy
'                .Document.Close
'            End With
'            .Paste
'            With .Find
'                .Text = "�i�^�e�@���j �i���ӭ����j �i���ӥ���j �i�W�@���^�j �i�U�@���^�j"
'                Do While .Execute(, , , , , , , wdFindContinue)
'                    .Parent.Paragraphs(1).Range.Delete
'                Loop
'            End With
'            .Range.Find.Execute " ", True, , , , , , wdFindContinue, , "�@", wdReplaceAll
'            .HomeKey wdStory, wdMove
'        End With
'        'Dir("D:\*��r�ɸ�Ʈw\�p��\���ӹ�" & Screen.ActiveForm.Controls("��") + 1 & "*")
    End If
End With
End Sub

Public Sub �M������Ÿ�()
'2003/4/3���w��:Shift+Backspace
Dim rp As String, p(2) As String, i As Byte
p(1) = Chr(10): p(2) = Chr(13)
With Selection '.Find '�p���~���|���r�ή榡�I
    If VBA.right(.Range, 1) Like p(1) Or VBA.right(.Range, 1) Like p(2) Then .MoveLeft wdCharacter, 1, wdExtend
''        .Range.Select(.Range.Words.Count=
'    .ClearFormatting
    rp = .Range
    For i = 1 To 2
         rp = Replace(rp, p(i), "")
    Next i
    .Range = rp
    
'    .Execute findtext:="^p", Replacewith:="", Wrap:=wdFindContinue, Replace:=wdReplaceAll
'    .ClearFormatting
End With
End Sub

Public Sub �ഫ����Ÿ�����ʤ���Ÿ�()
'2003/3/25
With Selection.Find '�p���~���|���r�ή榡�I
    .ClearFormatting
    .Execute FindText:="^p", Replacewith:="^l", Wrap:=wdFindContinue, Replace:=wdReplaceAll
'    .ClearFormatting
End With
'2003/3/23
'Dim i As Integer, s As Integer, p As Byte, InStrs As Long, InStrRevs As Long
's = Selection.Information(wdActiveEndSectionNumber) '�Ǧ^�Ҧb�����`�ơI
'With ActiveDocument.Sections(s) '�B�z�Ҧb���`��������
'    If IsNumeric(.Range.Paragraphs(1).Range.Text) Then p = 1 '�P�_���q�O�_���Ʀr
'    InStrs = InStr(.Range.Text, Chr(13)): InStrRevs = InStrRev(.Range.Text, Chr(13))
'    If InStrs = InStrRevs And InStrs <> 0 Then _
'    If MsgBox("���`�u���@�Ӥ���Ÿ�! �O�_�n���N�H", vbInformation + vbYesNo) = vbNo Then .Range.Words(.Range.Words.Count).Select: Exit Sub
'    For i = .Range.Paragraphs.Count - 1 To 1 + p Step -1 '�Ĥ@�q�Y�����X�h���n,�̥��@�q���`�Ÿ����N��A���F�����Ÿ��A�G�礣�n�I
'        With .Range.Paragraphs(i)  '�ֳt��G
'            .Range.Text = Replace(.Range.Text, Chr(13), Chr(11)) '�ഫ����Ÿ�����ʤ���Ÿ�
'        '    .Find.ClearFormatting
'        '    .Find.Replacement.ClearFormatting
'        '    .Find.Execute findtext:=.Text, MatchCase:=False, Replace:=wdReplaceAll, replacewith:=.Text, Wrap:=wdFindContinue
'        End With
'    Next i
'End With
End Sub


Sub �˵��o��Ѥ����s��B() '2003/3/28
Static bk As Bookmark, ps As Byte, dt As String, ThisDoc As String, dtbefore As Integer
Dim dts As String, dtbeforeStr As String
With ActiveDocument
    If ThisDoc <> .Name Then Set bk = Nothing: ps = 0: dt = ""
    ThisDoc = .Name '�����,�h���]
    If bk Is Nothing Then
        Select Case MsgBox("�n�s���u���ѡv���s��B��?", vbQuestion + vbYesNoCancel)
            Case vbCancel
                End
            Case vbYes
                dt = "Edit_" & Format(Now(), "yyyy_mm_dd")
                dtbefore = 0
            Case vbNo
Again:          dtbeforeStr = InputBox("�n��" & Chr(-24153) & "�X" & Chr(-24152) & _
                        "�ѥH�e���s��B?", "�s���s��B����", "1")
                If Not IsNumeric(dtbeforeStr) Then
                    If Not dtbeforeStr Like "" Then
                        MsgBox "�п�J�Ʀr�I": GoTo Again
                    Else
                       End
                    End If
                End If
                dtbefore = CInt(dtbeforeStr)
                dt = "Edit_" & Format(Now() - dtbefore, "yyyy_mm_dd")
        End Select
        For Each bk In .bookmarks
            ps = ps + 1 '�O�U���ޭ�
            With bk
                If InStr(.Name, dt) Then
                    bk.Select
    '                .GoTo wdGoToBookmark, wdGoToAbsolute, , bk.Name
                    GoTo Repeats
                    Exit For
                End If
            End With
        Next bk
        If bk Is Nothing Then MsgBox "�S���ŦX������" & vbCr & _
            "(�S���s��B�O���A�ΰO���w�R��)", vbExclamation, "�s���s��B": End
    Else
Repeats: ps = ps + 1 '�]�����ҹw�]�����ǤD�ӦW�ٱƧ�, _
                        �G�i�p�����R�A�ܼ�ps�ӳ]�p�ѷӮ��Ҫ����ޭ�
        If ps > .bookmarks.Count Then GoTo e
        Set bk = .bookmarks(ps)
'        If dt <> "Edit_" & Format(Now(), "yyyy_mm_dd") Then
        Select Case dtbefore
            Case Is > 2
                dts = "����]" & FormatDateTime(Date - dtbefore, vbLongDate) & "�^"
            Case Is = 2
                dts = "�e��"
            Case Is = 1
                dts = "�Q��"
            Case Else
                dts = "����"
        End Select
        If InStr(bk.Name, dt) = 0 Then
            '�s��������
e:          .bookmarks(ps - 1).Select '����]��^�̫�@���˵�������
            MsgBox dts & "�s��B�w�s������!" & vbCr & vbCr & _
                "�ثe���X�G" & Selection.Information(wdActiveEndAdjustedPageNumber) _
                , vbInformation
            End '��End �Y�i��l�Ƥ@���ܼ�
    '        Set bk = Nothing: ps = 0 '���s��l��
        Else
            Dim a As Paragraph '���o���ҩҦb��m���D'2003/4/26
            With Selection
                If Not .Information(wdInFootnote) Then
                    Set a = .Paragraphs(1)
                Else '���Ҧb���}�ɪ��B�z
                    Set a = .Footnotes(1).Reference.Paragraphs(1)
                End If
                Do
                    Set a = a.Previous
                   If Left(a.Style, 2) = "���D" Then _
                    Exit Do
                Loop 'a.Range�|�]�A�q���r��,�n�h���i�ΡGLeft(a.Range, Len(a.Range) - 1)
            End With
            Select Case MsgBox("�n���s�}�l�Ы��e�_�f!" & vbCr & vbCr & _
                "�ثe���D���G" & a.Range & _
                "�F�ثe���X�G" & Selection.Information(wdActiveEndAdjustedPageNumber), _
                vbQuestion + vbYesNoCancel, _
                "�˵��u" & dts & "�v���s��B...�n�~���?")
                Case vbCancel
                    ps = ps - 1
                    Exit Sub
                Case vbNo
                    End
                Case vbYes
        '            .ActiveWindow.ScrollIntoView .Range, True
                    .bookmarks(ps).Select
                    GoTo Repeats
            End Select
        End If
    End If
End With
End Sub

Private Sub �A�Ѯ��ұƧ�()
Dim bk As Bookmark, bkIdx As Integer
With ActiveDocument.bookmarks
    For bkIdx = 1 To .Count
'     .DefaultSorting = wdSortByName
        Debug.Print .item(bkIdx)
    Next bkIdx
End With
End Sub

Sub ���J�椬�ѷ�() '2003/3/28'���w��GCtrl+Shift+Insert
Dim CrossReference, i As Integer, CrossReferenceID As String
Static doinsert As Boolean, WinID As Byte, DocWin As Byte
Dim Winview As Byte, s As Long '2003/3/31
With ActiveDocument
    If Selection.Type = 2 Then
        WinID = .ActiveWindow.WindowNumber
        DocWin = .ActiveWindow.Previous.WindowNumber
        GoTo 1
    End If
    If doinsert = False Then
'        DocWin = .ActiveWindow.Index '�O�U�ϥΤ���󤧱N�n���J�椬�ѷӤ��������ޭ�
        DocWin = .ActiveWindow.WindowNumber 'Index�O�����������s���A���D������󤧽s���A�G�̤��P�I2003/4/1
        '�p�G�O�b���}�˵�(���S���˵��϶��^�B�A�h�O�U'2003/3/31
        With .ActiveWindow.View 'wdPaneNone=0
            If .SplitSpecial <> wdPaneNone Then
                Winview = .SplitSpecial
                s = .Application.Selection.start '�O�U���J�I��m
            End If
        End With
        .Windows.Add
         With .ActiveWindow.View
            If Winview <> wdPaneNone Then 'wdPaneNone=0
                .SplitSpecial = Winview '�]�w�S���˵������]�p���}...���^
                .Application.Selection.start = s '�]�w���J�I��m
            End If
        End With
'        .ActiveWindow.Application.GoBack
        WinID = .ActiveWindow.WindowNumber
         MsgBox "�Цb����������������J���椬�ѷӪ���" & vbCr & vbCr _
                & "��n��A�A���@���A�Y�i���J�I", vbExclamation
        doinsert = True
    Else
1       Select Case .ActiveWindow.Selection.Range.StoryType
            Case wdMainTextStory
                Select Case Selection.Style
                    Case "���D 1", "���D 2", "���D 3", "���D 4", "���D 5", "���D 6" _
                        , "���D 7", "���D 8", "���D 9"
                        'wdStyleHeader  'Left(Selection.Style, 2) = "���D" '�p�G�O���D
'                        CrossReferenceID = .ActiveWindow.Selection.HeaderFooter .Footnotes (1).Index
                        CrossReference = .GetCrossReferenceItems(wdRefTypeHeading)
                        For i = 1 To UBound(CrossReference)
'                            If Trim(Left(CrossReference(i), Len(CrossReferenceID))) _
                                = CrossReferenceID Then
                            If Trim(CrossReference(i)) Like Selection Then
                                If MsgBox("�n���J���}�Ҧb���u���X�v�ӫD�u���D��r�v�A" _
                                        & "�Ы��e�����f", vbQuestion + vbOKCancel, "���J���D:" & _
                                        .ActiveWindow.Selection.Style & .ActiveWindow.Selection) _
                                        = vbOK Then
                                    .Windows(DocWin).Activate '�H��s���������
                                    Selection.Range.InsertCrossReference _
                                        ReferenceType:=wdRefTypeHeading _
                                        , ReferenceKind:=wdContentText, _
                                            ReferenceItem:=i _
                                                , InsertAsHyperlink:=True
                                Else
                                    .Windows(DocWin).Activate '�H��s���������
                                    Selection.Range.InsertCrossReference _
                                        ReferenceType:=wdRefTypeHeading _
                                        , ReferenceKind:=wdPageNumber, _
                                            ReferenceItem:=i _
                                                , InsertAsHyperlink:=True
                                End If
                                Exit For
                            End If
                        Next i
                        If i > UBound(CrossReference) Then MsgBox "�S���X�A�����}�ѷӡA�Ф�ʾާ@�I", vbExclamation: End
                        .Windows(WinID).Close
                        doinsert = False: WinID = 0: DocWin = 0
                    On Error GoTo ErrHs
                    Case "���}�ѷ�" 'wdStyleFootnoteReference, wdStyleFootnoteText
                        CrossReferenceID = .ActiveWindow.Selection.Range.Footnotes(1).index
                        CrossReference = .GetCrossReferenceItems(wdRefTypeFootnote)
                        For i = 1 To UBound(CrossReference)
                            If Trim(Left(CrossReference(i), Len(CrossReferenceID))) _
                                = CrossReferenceID Then
                                If MsgBox("�n���J���}�Ҧb���u���X�v�ӫD�u���}�s���v�A" _
                                        & "�Ы��e�����f", vbQuestion + vbOKCancel, "���J���}:" & _
                                        .ActiveWindow.Selection.Range.Footnotes(1).index) _
                                        = vbOK Then
                                    .Windows(DocWin).Activate '�H��s���������
                            '        .ActiveWindow.Selection.Range.Paste
                                    Selection.Range.InsertCrossReference _
                                        ReferenceType:=wdRefTypeFootnote _
                                        , ReferenceKind:=wdFootnoteNumber, _
                                            ReferenceItem:=i _
                                                , InsertAsHyperlink:=True
                                Else
                                    .Windows(DocWin).Activate '�H��s���������
                                    Selection.Range.InsertCrossReference _
                                        ReferenceType:=wdRefTypeFootnote _
                                        , ReferenceKind:=wdPageNumber, _
                                            ReferenceItem:=i _
                                                , InsertAsHyperlink:=True
                                End If
                                Exit For
                            End If
                        Next i
                        If i > UBound(CrossReference) Then MsgBox "�S���X�A�����}�ѷӡA�Ф�ʾާ@�I", vbExclamation: End
                        .Windows(WinID).Close
                        doinsert = False: WinID = 0: DocWin = 0
                    Case Else '���J����
                        For Each CrossReference In .bookmarks
                            If CrossReference.Range Like ActiveWindow.Selection Then
                                If MsgBox("���J����: " & CrossReference.Name, vbQuestion + vbOKCancel, "���J�椬�ѷ�") = vbOK Then
                                    .Windows(DocWin).Activate '�H��s���������
                                    Selection.Range.InsertCrossReference _
                                        ReferenceType:=wdRefTypeBookmark _
                                        , ReferenceKind:=wdPageNumber, _
                                            ReferenceItem:=CrossReference.Name, _
                                            InsertAsHyperlink:=True
                                    .Windows(WinID).Close
                                    doinsert = False: WinID = 0: DocWin = 0
                                    Exit For
                                End If
                            End If
                        Next CrossReference
                        If IsObject(CrossReference) Then
                            If CrossReference Is Nothing Then MsgBox "�S���X�A�����ҰѷӡA�Ф�ʾާ@�I", vbExclamation: End
                        Else
                            If IsEmpty(CrossReference) Then MsgBox "�S���X�A�����ҰѷӡA�Ф�ʾާ@�I", vbExclamation: End
                        End If
                End Select
            Case wdFootnotesStory
                .ActiveWindow.Selection.Range.Copy
                .ActiveWindow.Close
                .Windows(DocWin).Activate
                .ActiveWindow.Selection.Range.Paste
        End Select
    End If
Exit Sub
'End With
ErrHs:
'With ActiveDocument
Select Case Err.Number
'    Case 5941 '���X�����������s�b�]�Y�S�����}�^,�h���J����2003/3/29
'        For Each CrossReference In .Bookmarks
'            If CrossReference.Range Like ActiveWindow.Selection Then
'                If MsgBox("���J����: " & CrossReference.Name, vbQuestion + vbOKCancel, "���J�椬�ѷ�") = vbOK Then
'                    .Windows(DocWin).Activate '�H��s���������
'                    Selection.Range.InsertCrossReference _
'                        ReferenceType:=wdRefTypeBookmark _
'                        , ReferenceKind:=wdPageNumber, _
'                            ReferenceItem:=CrossReference.Name, _
'                            InsertAsHyperLink:=True
'                    .Windows(WinID).Close
'                    doinsert = False: WinID = 0: DocWin = 0
'                    Exit For
'                End If
'            End If
'        Next CrossReference
'        If IsObject(CrossReference) Then
'            If CrossReference Is Nothing Then MsgBox "�S���X�A�����ҰѷӡA�Ф�ʾާ@�I", vbExclamation: End
'        Else
'            If IsEmpty(CrossReference) Then MsgBox "�S���X�A�����ҰѷӡA�Ф�ʾާ@�I", vbExclamation: End
'        End If
    Case Else
        MsgBox Err.Number & Err.Description: End
End Select
End With
End Sub
Public Sub �b�t�@��󤤴M�����r��()
Static winNum As Byte, preR As String
Dim r As String, ins(4) As Long, MnText As String, FnText As String
Dim d As Document, winINdex As Byte, startD As Document
'CheckSavedNoClear
With Selection '���w��GAlt+Ctrl+Up
'If Not .Text Like "" Then '�ֳt��GAlt+Ctrl+Down
    If .Type = wdSelectionIP Then MsgBox "�п���Q�n�M�䤧��r", vbExclamation: Exit Sub
    If .Type = wdSelectionNormal Then '�������J�I
        If preR <> r Then winNum = 0
        r = .text
        preR = r
        Set startD = .Document
        On Error GoTo Previews
        For Each d In .Application.Documents
            winINdex = winINdex + 1
            If Not d Is startD And winINdex > winNum Then
            'With .Application.ActiveWindow.Document ' ActiveDocument
                With d
                    
                    MnText = .StoryRanges(wdMainTextStory) '���ܼƥN����g���ӻ����֡I2003/4/8
                    ins(1) = InStr(MnText, r)
                    ins(2) = InStrRev(MnText, r)
                    If .Footnotes.Count > 1 Then
                        FnText = .StoryRanges(wdFootnotesStory)
                        ins(3) = InStr(FnText, r)
                        ins(4) = InStrRev(FnText, r)
                    End If
                    If ins(1) = 0 And ins(3) = 0 Then
                        Select Case MsgBox("�S���ŦX��r!" & vbCr & vbCr & _
                            "�O�_�n��U�@�����H", vbExclamation + vbOKCancel)
                            Case vbOK
                                
                            Case vbCancel
                                winNum = winINdex
                                Exit Sub
                        End Select
                    End If
                    If ins(1) = ins(2) And ins(3) = ins(4) Then
                        d.Activate
                        MsgBox "����u�����B!", vbInformation ': Exit Sub
                    End If
                    If ins(1) <> 0 Then
                        ins(1) = wdMainTextStory
                    Else
                        ins(1) = wdFootnotesStory
                    End If
                    With .StoryRanges(ins(1)).Find
        '            With Selection.Find
                        .ClearFormatting
                        .Replacement.ClearFormatting '�o�]�n�M���~��
                        .Forward = True
                        .Wrap = wdFindAsk
                        .MatchCase = True
                        .text = r
                        .Execute
                        .Parent.Select
                        d.Activate
                        With d.ActiveWindow
                            .ScrollIntoView Selection
                            If .WindowState = wdWindowStateMinimize Then
                                .WindowState = wdWindowStateNormal
                            End If
                        End With
'                        With .Application.ActiveWindow
'                            If .WindowState = wdWindowStateMinimize Then .WindowState = wdWindowStateMaximize
'                        End With
                        winNum = winINdex
                        Exit Sub
                    End With
                End With
            End If
        Next d
        winNum = 0
    End If
End With

Exit Sub
Previews:
Select Case Err.Number
'    Case 91
'        On Error Resume Next
'        Dim d As Byte
'        d = Documents.Count
'        If d > 1 Then
'            If Documents(d - 1) <> ActiveDocument Then
'                Documents(d - 1).Activate
'            Else
'                Documents(d).Activate
'            End If
'        End If
''        ActiveWindow.Previous.Document.ActiveWindow.Activate
'        Resume Next
'    Case 5941
'        Resume Next
    Case Else
        MsgBox Err.Number & Err.Description, vbExclamation
End Select
End Sub

Public Sub �b�t�@��󤤴M�����r��_old()
Static winNum As Byte
Dim r As String, ins(4) As Long, MnText As String, FnText As String
CheckSavedNoClear
With Selection
'If Not .Text Like "" Then '�ֳt��GAlt+Ctrl+Down
    If .Type = wdSelectionIP Then MsgBox "�п���Q�n�M�䤧��r", vbExclamation: Exit Sub
    If .Type = wdSelectionNormal Then ' <> wdNoSelection OR wdSelectionIP Then '�������J�I
        r = .text
        On Error GoTo Previews
Again: If winNum <> 0 Then .Application.Windows(winNum).Activate
        If .Application.Documents.Count = 1 And .Document.Windows.Count > 1 Then
            If .Document.ActiveWindow.WindowNumber < .Document.Windows.Count Then
                .Document.ActiveWindow.Next.Activate
            Else
                .Document.ActiveWindow.Previous.Activate
            End If
        Else
            With .Application.ActiveWindow.Next.Document
                If .Windows.Count > 1 Then
                    .ActiveWindow.Activate
                Else
                    .Activate
                End If
            End With
        End If
        'If InStr(ActiveDocument.Name, "�r��7") Then Register_Event_Handler: Documents("�r��7.2.doc").Windows(1).Visible = True
        If InStr(ActiveDocument.Name, "�r��") Then Register_Event_Handler: d�r��.Windows(1).Visible = True
        With .Application.ActiveWindow.Document ' ActiveDocument
            MnText = .StoryRanges(wdMainTextStory) '���ܼƥN����g���ӻ����֡I2003/4/8
            FnText = .StoryRanges(wdFootnotesStory)
            ins(1) = InStr(MnText, r)
            ins(2) = InStrRev(MnText, r)
            ins(3) = InStr(FnText, r)
            ins(4) = InStrRev(FnText, r)
            If ins(1) = 0 And ins(3) = 0 Then
                Select Case MsgBox("�S���ŦX��r!" & vbCr & vbCr & _
                    "�O�_�n��U�@�����H", vbExclamation + vbYesNoCancel)
                    Case vbYes
                        '�O�U�ثe�������G
                        winNum = .ActiveWindow.index '.WindowNumber
                        GoTo Again
                    Case vbNo
                        .Application.ActiveWindow.Previous.Activate
                         Exit Sub 'End
                    Case vbCancel
                        Exit Sub 'End'end �|���]�Ҧ��ܼƤγ]�w��,�]�A�ϥ�application��Register_Event_Handler
                End Select
            End If
            If winNum <> Empty Then winNum = Empty
            If ins(1) = ins(2) And ins(3) = ins(4) Then _
                MsgBox "����u�����B!", vbInformation ': Exit Sub
            If ins(1) <> 0 Then
                ins(1) = wdMainTextStory
            Else
                ins(1) = wdFootnotesStory
            End If
            With .StoryRanges(ins(1)).Find
'            With Selection.Find
                .ClearFormatting
                .Replacement.ClearFormatting '�o�]�n�M���~��
                .Forward = True
                .Wrap = wdFindAsk
                .MatchCase = True
                .text = r
                .Execute
                .Parent.Select
                With .Application.ActiveWindow
                    If .WindowState = wdWindowStateMinimize Then .WindowState = wdWindowStateMaximize
                End With
            End With
        End With
    End If
End With
Exit Sub
Previews:
Select Case Err.Number
    Case 91
        On Error Resume Next
        Dim d As Byte
        d = Documents.Count
        If d > 1 Then
            If Documents(d - 1) <> ActiveDocument Then
                Documents(d - 1).Activate
            Else
                Documents(d).Activate
            End If
        End If
'        ActiveWindow.Previous.Document.ActiveWindow.Activate
        Resume Next
    Case 5941
        Resume Next
    Case Else
        MsgBox Err.Number & Err.Description, vbExclamation
End Select
End Sub

Public Sub �������r()  '2003/4/4(���]�A���@�Ÿ�)
'���w��: Atl+Ctrl+Shift+Up(��)
CheckSavedNoClear
'�r����Gbetween -24667 and 19968
'Selection = ChrW(�r����)
Dim r As String, ins(4) As Long, f, i As Long, rCompMain As String, rCompFootnote As String, R1 As String
f = Array("�C", "�v", Chr(-24152), "�G", "�A", "�F", _
    "�B", "�u", ".", Chr(34), ":", ",", ";", _
            "�K�K", "...", "�^", ")", "-", "�D", "�y", "�z" _
            , "�m", "�n", "�r", "�q", "�]", "�^", "--", _
            ChrW(8212), "��", "�H", ChrW(2), Chr(13), Chr(10), Chr(8), Chr(9), _
            "�@", " ")
            'ChrW (2)�����}�Ÿ�
With Selection '���w��GAlt+Ctrl+Up
'If Not .Text Like "" Then '�ֳt��GAlt+Ctrl+Down
    If .Type = wdSelectionIP Then MsgBox "�п���Q�n�M�䤧��r", vbExclamation: Exit Sub
    If .Type = wdSelectionNormal Then ' <> wdNoSelection OR wdSelectionIP Then '�������J�I
        
        r = .text
        For i = 0 To UBound(f)
            If InStr(r, f(i)) Then
                r = Replace(r, f(i), "")
            End If
        Next i
'        Debug.Print r
        On Error GoTo Previews
'        If .Application.ActiveWindow.Next.Document.Name = .Document.Name Then _
'            .Application.ActiveWindow.Next.Document.Activate
        With .Application.ActiveWindow.Next.Document
            If .Windows.Count > 1 Then
                .ActiveWindow.Activate
            Else
                .Activate
            End If
        End With
        With .Application.ActiveWindow.Document ' ActiveDocument
            rCompMain = .StoryRanges(wdMainTextStory)
            If .Footnotes.Count > 0 Then _
                rCompFootnote = .StoryRanges(wdFootnotesStory)
            For i = 0 To UBound(f)
                If InStr(rCompMain, f(i)) Then
                    rCompMain = Replace(rCompMain, f(i), "")
                End If
                If Not rCompFootnote Like "" And InStr(rCompFootnote, f(i)) Then _
                    rCompFootnote = Replace(rCompFootnote, f(i), "")
            Next i
            f = Empty '����O����
            ins(1) = InStr(rCompMain, r)
            ins(2) = InStrRev(rCompMain, r)
            ins(3) = InStr(rCompFootnote, r)
            ins(4) = InStrRev(rCompFootnote, r)
            If ins(1) = 0 And ins(3) = 0 Then _
                MsgBox "�S���ŦX��r", vbExclamation: _
                    .Application.ActiveWindow.Previous.Activate: End
            If ins(1) = ins(2) And ins(3) = ins(4) Then _
                MsgBox "����u�����B!", vbInformation ': Exit Sub
            If ins(1) <> 0 Then
                ins(1) = wdMainTextStory
            Else
                ins(1) = wdFootnotesStory
            End If
'            rCompFootnote = Empty '�Τ��۪��r���ܼ��k�s
            rCompMain = r '���s�ϥΦr���ܼ�
            For i = 1 To Len(rCompMain)
                If InStr(.StoryRanges(ins(1)), Left(rCompMain, i)) = 0 Then Exit For
            Next i
            rCompFootnote = rCompMain '���s�ϥΦr���ܼ�
            rCompMain = Left(rCompMain, i - 1)
            For i = 1 To Len(rCompFootnote)
                If InStrRev(.StoryRanges(ins(1)), right(rCompFootnote, i), -1, vbTextCompare) = 0 Then Exit For
            Next i
            rCompFootnote = right(rCompFootnote, i - 1)
            R1 = rCompMain
            Beep
            ins(2) = 1: ins(3) = 0 '���s�ϥ��ܼ�
            Do While ins(2) > ins(3)
                ins(2) = InStr(ins(2) + Len(rCompMain), .StoryRanges(ins(1)), rCompMain, vbTextCompare) - 1
                ins(3) = InStrRev(.StoryRanges(ins(1)), rCompFootnote, ins(3) - 1, vbTextCompare) - 1 + Len(rCompFootnote)
            Loop
            Selection.SetRange ins(2), ins(3)
            .ActiveWindow.ScrollIntoView Selection.Range, True
            .ActiveWindow.ScrollIntoView Selection.Range, False
            Beep
            For i = ins(3) To ins(2) Step -1
'                Application.System
                If rCompMain Like Selection.Range Then
                    Exit For
                Else
                    rCompMain = Selection.Range
                End If
'                If Right(rCompMain, Len(r)) Like "���A�h" Then Stop
                If Len(rCompMain) >= Len(r) And _
                InStrRev(rCompMain, rCompFootnote) = 0 Then
'                    ins(3) = ins(3) + 1 '���Y�u�@�����רS���ɡA�h�_�����ס]��̫ܳ�ŦX�̡A�Y��@���׫e���r��o�^
                    ins(3) = ins(3) + Len(rCompFootnote) '���Y�u�@�����רS���ɡA�h�_�����ס]��̫ܳ�ŦX�̡A�Y��@���׫e���r��o�^
                    Exit For
                End If
'                ins(3) = ins(3) - 1 '�Y�u�@�����צA��
                If InStr(right(rCompMain, Len(rCompMain) - Len(R1)), R1) > 0 Then
                    ins(2) = InStr(right(rCompMain, Len(rCompMain) - Len(R1)), R1) - 1 + Len(R1) + ins(2) '�Y�u�@�����צA��
'                Else
'                    Exit For
                End If
                Selection.SetRange start:=ins(2), End:=ins(3)
                .ActiveWindow.ScrollIntoView Selection.Range, True '��ܿ���d��
                If InStrRev(Left(rCompMain, Len(rCompMain) - Len(rCompFootnote)), rCompFootnote, -1, vbTextCompare) > 0 Then
                    ins(3) = InStrRev(Left(rCompMain, Len(rCompMain) - Len(rCompFootnote)), rCompFootnote, -1, vbTextCompare) - 1 + Len(rCompFootnote) + ins(2) '�Y�u�@�����צA��
'                Else
'                    Exit For
                End If
                Selection.SetRange start:=ins(2), End:=ins(3)
'                .StoryRanges(ins(1)).SetRange Start:=ins(2), End:=ins(3)
                .ActiveWindow.ScrollIntoView Selection.Range, False '��ܿ���d��
                
            Next i
            Beep
            '�]�w����d��
            Selection.SetRange ins(2), ins(3)
            .ActiveWindow.ScrollIntoView Selection.Range, False
'            Selection.SetRange InStr(.StoryRanges(ins(1)), rCompMain), _
'                InStrRev(.StoryRanges(ins(1)), rCompFootnote, -1, vbTextCompare)
'            With .StoryRanges(ins(1)).Find
'                .ClearFormatting
'                .Replacement.ClearFormatting '�o�]�n�M���~��
'                .Forward = True
'                .Wrap = wdFindAsk
'                .MatchCase = True
'                .Text = r
'                .Execute
'                .Parent.Select
'                With .Application.ActiveWindow
'                    If .WindowState = wdWindowStateMinimize Then .WindowState = wdWindowStateMaximize
'                    .ScrollIntoView Selection.Range, True '��ܿ���d��
'                End With
'            End With
        End With
    End If
End With
Exit Sub
Previews:
Select Case Err.Number
    Case 91
        On Error Resume Next
        Dim d As Byte
        d = Documents.Count
        If d > 1 Then
            If Documents(d - 1) <> ActiveDocument Then
                Documents(d - 1).Activate
            Else
                Documents(d).Activate
            End If
        End If
'        ActiveWindow.Previous.Document.ActiveWindow.Activate
        Resume Next
    Case 5941
        Resume Next
    Case Else
        MsgBox Err.Number & Err.Description, vbExclamation
End Select
End Sub

Public Sub �������r1() '2003/4/4(���]�A���@�Ÿ�)
'���w��: Atl+Ctrl+Shift+Up(��)
CheckSavedNoClear
'�r����Gbetween -24667 and 19968
'Selection = ChrW(�r����)
Dim r As String, ins(4) As Long, f
Dim rLeft As String, rRight As String, rComp As String, i As Long, rCompMain As String, rCompFootnote As String
f = Array("�C", "�v", Chr(-24152), "�G", "�A", "�F", _
    "�B", "�u", ".", Chr(34), ":", ",", ";", _
            "�K�K", "...", "�^", ")", "-", "�D", "�y", "�z" _
            , "�m", "�n", "�r", "�q", "�]", "�^", "--", _
            ChrW(8212), "��", "�H", ChrW(2), Chr(13), Chr(10), Chr(8), Chr(9), _
            "�@", " ")
            'ChrW (2)�����}�Ÿ�
With Selection '���w��GAlt+Ctrl+Up
'If Not .Text Like "" Then '�ֳt��GAlt+Ctrl+Down
    If .Type = wdSelectionIP Then MsgBox "�п���Q�n�M�䤧��r", vbExclamation: Exit Sub
    If .Type = wdSelectionNormal Then ' <> wdNoSelection OR wdSelectionIP Then '�������J�I

        r = .text
        For i = 0 To UBound(f)
            If InStr(r, f(i)) Then
                r = Replace(r, f(i), "")
            End If
        Next i
'        Debug.Print r
        On Error GoTo Previews
        With .Application.ActiveWindow.Next.Document
            If .Windows.Count > 1 Then
                .ActiveWindow.Activate
            Else
                .Activate
            End If
        End With
        With .Application.ActiveWindow.Document ' ActiveDocument
            rCompMain = .StoryRanges(wdMainTextStory)
            If .Footnotes.Count > 0 Then _
                rCompFootnote = .StoryRanges(wdFootnotesStory)
            '�����Ÿ��G
            For i = 0 To UBound(f)
                If InStr(rCompMain, f(i)) Then
                    'rCompMain=�S���Ÿ�������
                    rCompMain = Replace(rCompMain, f(i), "")
                End If
                If Not rCompFootnote Like "" And InStr(rCompFootnote, f(i)) Then _
                    rCompFootnote = Replace(rCompFootnote, f(i), "")
                    'rCompFootnote=�S���Ÿ������}
            Next i
            f = Empty '����O����
            ins(1) = InStr(rCompMain, r)
            ins(2) = InStrRev(rCompMain, r)
            ins(3) = InStr(rCompFootnote, r)
            ins(4) = InStrRev(rCompFootnote, r)
            If ins(1) = 0 And ins(3) = 0 Then _
                MsgBox "�S���ŦX��r", vbExclamation: _
                    .Application.ActiveWindow.Previous.Activate: End
            If ins(1) = ins(2) And ins(3) = ins(4) Then _
                MsgBox "����u�����B!", vbInformation ': Exit Sub
            If ins(1) <> 0 Then
'                ins(1) = wdMainTextStory
                rComp = rCompMain '.StoryRanges(wdMainTextStory)
            Else
'                ins(1) = wdFootnotesStory
                rComp = rCompFootnote '.StoryRanges(wdFootnotesStory)
            End If
'            rCompFootnote = Empty '�Τ��۪��r���ܼ��k�s
            '�ѥ�����o�Ĥ@�ӧk�X���媺��
            For i = 1 To Len(r)
                If InStr(rComp, Left(r, i)) = 0 Then Exit For
            Next i
            rLeft = Left(r, i - 1)
            '�ѥk����o�Ĥ@�ӧk�X���媺��
            For i = 1 To Len(r)
                If InStrRev(rComp, right(r, i), -1, vbTextCompare) = 0 Then Exit For
            Next i
            rRight = right(r, i - 1)

            ins(2) = InStr(rComp, rLeft) '- 1
            ins(3) = InStrRev(rComp, rRight, -1, vbTextCompare) + Len(rRight) '- 1
'            Selection.SetRange ins(2), ins(3)
'            rComp = Mid(rComp, ins(2), ins(3) - ins(2))
            For i = ins(2) To ins(3) Step 1
'                Application.System
'                r = Selection.Range
'                If Right(r, Len(r)) Like "���A�h" Then Stop
                ins(2) = ins(2) + 1
                If InStr(right(rComp, Len(rComp) - Len(rRight)), rRight) > 0 Then _
                    ins(2) = InStr(right(rComp, Len(rComp) - Len(rRight)), rRight) - 1 + Len(rRight) + ins(2) '�Y�u�@�����צA��
                If Len(rComp) >= Len(r) Then
'                 InStrRev(rComp, rRight) = 0
'                    ins(3) = ins(3) + 1 '���Y�u�@�����רS���ɡA�h�_�����ס]��̫ܳ�ŦX�̡A�Y��@���׫e���r��o�^
                    ins(3) = ins(3) + Len(rRight) '���Y�u�@�����רS���ɡA�h�_�����ס]��̫ܳ�ŦX�̡A�Y��@���׫e���r��o�^
                    Exit For
                End If
'                ins(3) = ins(3) - 1 '�Y�u�@�����צA��
                If InStr(right(rComp, Len(rComp) - Len(rRight)), rRight) > 0 Then
                    ins(2) = InStr(right(rComp, Len(rComp) - Len(rRight)), rRight) - 1 + Len(rRight) + ins(2) '�Y�u�@�����צA��
'                Else
'                    Exit For
                End If
                Selection.SetRange start:=ins(2), End:=ins(3)
                .ActiveWindow.ScrollIntoView Selection.Range, True '��ܿ���d��
                If InStrRev(Left(rComp, Len(rComp) - Len(rRight)), rRight, -1, vbTextCompare) > 0 Then
                    ins(3) = InStrRev(Left(rComp, Len(rComp) - Len(rRight)), rRight, -1, vbTextCompare) - 1 + Len(rRight) + ins(2) '�Y�u�@�����צA��
'                Else
'                    Exit For
                End If
                Selection.SetRange start:=ins(2), End:=ins(3)
'                .StoryRanges(ins(1)).SetRange Start:=ins(2), End:=ins(3)
                .ActiveWindow.ScrollIntoView Selection.Range, False '��ܿ���d��
            Next i
            Beep
            '�]�w����d��
            Selection.SetRange ins(2), ins(3)
'            Selection.SetRange InStr(.StoryRanges(ins(1)), rcomp), _
'                InStrRev(.StoryRanges(ins(1)), rright, -1, vbTextCompare)
'            With .StoryRanges(ins(1)).Find
'                .ClearFormatting
'                .Replacement.ClearFormatting '�o�]�n�M���~��
'                .Forward = True
'                .Wrap = wdFindAsk
'                .MatchCase = True
'                .Text = r
'                .Execute
'                .Parent.Select
'                With .Application.ActiveWindow
'                    If .WindowState = wdWindowStateMinimize Then .WindowState = wdWindowStateMaximize
'                    .ScrollIntoView Selection.Range, True '��ܿ���d��
'                End With
'            End With
        End With
    End If
End With
Exit Sub
Previews:
Select Case Err.Number
    Case 91
        On Error Resume Next
        Dim d As Byte
        d = Documents.Count
        If d > 1 Then
            If Documents(d - 1) <> ActiveDocument Then
                Documents(d - 1).Activate
            Else
                Documents(d).Activate
            End If
        End If
'        ActiveWindow.Previous.Document.ActiveWindow.Activate
        Resume Next
    Case 5941
        Resume Next
    Case Else
        MsgBox Err.Number & Err.Description, vbExclamation
End Select
End Sub

Public Sub �s���v��() '2003/4/6
With Selection
    If IsNumeric(.Range) Then
        Dim r As Integer
        r = CInt(.Range)
        If .Document.Range(.End, .End + 1) Like "." Then r = r + 1
        With .Find
            .ClearFormatting
'            .ClearAllFuzzyOptions
            .text = CStr(r)
            .Execute Forward:=True, Wrap:=wdFindContinue ', Wrap:=wdFindAsk
        End With
        If .Range = r And Not .Document.Range(.start - 1, .start) Like Chr(13) _
            And .Document.Range(.End, .End + 1) Like "." Then
            .Range = Chr(13) & r
            .MoveRight
            '�n�ন�r��A�ұo���פ謰�r����סA�Ʀr�̡ALen()�h�o�b����
            .SetRange .start, End:=.start + Len(CStr(r))
        End If
    End If
End With
End Sub

Sub �s���v��_�۰�() '2003/4/6
With Selection
    If IsNumeric(.Range) Then
        Dim r As Integer, C As Integer, p1 As Long, p2 As Long
        Do
            '�p�G�˵ۧ�έ�a��A�h��ܧ䧹�F�A��+1�A�~���C�]��Find�]�w��Wrap:=wdFindContinue
            If p1 >= .start Then
                r = r + 1
            Else
                r = CInt(.Range)
            End If
            If .Document.Range(.End, .End + 1) Like "." Then r = r + 1
            p1 = .start '�O�U�M��U�@�Ӯɪ���m
            With .Find
                .ClearFormatting
                .ClearAllFuzzyOptions
                .text = r
                .Execute Forward:=True, Wrap:=wdFindContinue ' Wrap:=wdFindAsk
            End With
            If .Document.Range(.start - 1, .start) Like Chr(13) Or p2 >= p1 Then
                MsgBox "�w����" & C & "�������I", vbInformation
                Exit Do
            End If
            If .Range = r And .Document.Range(.End, .End + 1) Like "." Then
                .Range = Chr(13) & r
                C = C + 1
                p2 = .start
                .MoveRight
            '�n�ন�r��A�ұo���פ謰�r����סA�Ʀr�̡ALen()�h�o�b����
                .SetRange .start, End:=.start + Len(CStr(r))
            End If
        Loop
    End If
End With
End Sub

Sub �s���v��_�ˬd() '2003/4/6
With Selection
    If IsNumeric(.Range) Then
        Dim r As Integer, R1 As Integer, C As Integer, p1 As Long
        Do
            If p1 >= .start Then GoTo Out ' Exit Do
            r = CInt(.Range)
            R1 = CInt(.Range)
            If .start = 0 Then .MoveRight
            r = r + 1
            p1 = .start
'            .MoveRight
            With .Find
                .ClearFormatting
                .ClearAllFuzzyOptions
                .text = r 'CStr(r)
                .Execute Forward:=True, Wrap:=wdFindContinue       ', Wrap:=wdFindContinue ' Wrap:=wdFindAsk
            End With
            C = C + 1
            If R1 = CInt(.Range) And _
                (Not .Document.Range(.start - 1, .start) Like Chr(10) _
                    Or Not .Document.Range(.start - 1, .start) Like Chr(13)) _
                    And .Document.Range(.End, .End + 1) Like "." Then
Out:            MsgBox "�w�ˬd" & C & "���I", vbExclamation
                Exit Do
            End If
        Loop
    End If
End With
End Sub

Sub �M����y���_��() '2003/4/7
Dim a As String, b As String
Dim C As Integer, p As Integer, d As Long, StepByStep As Byte
Const NoArrange = 255
StepByStep = MsgBox("�n�v�B�˵��ܡH", vbYesNoCancel + vbDefaultButton2 + vbQuestion)
If StepByStep = vbCancel Then End
With Selection
    d = Len(.Document.Content)
    If .End >= d Then
        If MsgBox("�n�q�Y�}�l�ܡH", vbQuestion + vbOKCancel) = vbOK Then
           .HomeKey wdStory, wdMove
            If .text Like Chr(13) Then .Delete
        Else
            Exit Sub
        End If
    End If
    If .Type <> wdSelectionIP Then .MoveRight '������d��ɷ|���N������d��, �G�����ˬd!
    Do
        .Find.ClearFormatting
        .Find.Execute FindText:="^p", Forward:=True
        C = C + 1
        a = .Range.Previous '���k����
'        a = .Document.Range(.Start - 1, .Start)
'        If .End + 1 > Len(.Document.Content) Then Exit Do
        If .End + 1 >= d Then Exit Do
        b = .Range.Next
'        b = .Document.Range(.End, .End + 1)
        If InStr(.Paragraphs(1).Range, "�ɰO") Then
            Stop
'            If StepByStep = NoArrange Then StepByStep = vbNo'�հɰO����(���ܥ|�v�榡���P,�Ƚw)
            StepByStep = NoArrange '�հɰO���榡���P
        End If
        If (Not Asc(a) = 13 And Not Asc(a) = 10 And Not IsNumeric(a) _
                    And Not a Like "-") And Not Asc(a) = 46 _
            And (Not Asc(b) = 13 And Not Asc(b) = 10 _
                And Not Asc(b) = 91 And Not Asc(b) = 93 _
                    And Not IsNumeric(b) _
               And Not b Like "�@" And Not b Like "-" And Not b Like "�e") And Not b Like "�i" Then
            .Document.ActiveWindow.LargeScroll 1, 0, 0, 0
            .Document.ActiveWindow.ScrollIntoView Selection.Range, True
            If StepByStep = vbNo Then
'                If Len(.Paragraphs(1).Range) > 28 And _
                    Left(VBA.Right(.Paragraphs(1).Range, 3), 1) <> "�C" And _
                    Left(VBA.Right(.Paragraphs(1).Range, 4), 2) <> "�C�v" Then      '�p���h�r�ƦӬq��̤~�B�z,�_�h���c�ƤF!(�s���D���]��J,�N�Ӻ��H�F!)2003/11/30
               '�H�֩��ֶ����榡���A�ȥ[�p��IF���󦡡I2003/11/30
'                If Len(.Paragraphs(1).Range) < 40 Then
'                    If MsgBox("�n�M����?", vbQuestion + vbOKCancel) = vbOK Then .Range = ""
'                Else
                   .Range = ""
'                End If
                p = p + 1
'                End If
            Else '>28,�H�֩��ֶ����榡���A�ȥ[�p��IF���󦡡I2003/11/30
                If Len(.Paragraphs(1).Range) > 36 And Left(VBA.right(.Paragraphs(1).Range, 3), 1) <> "�C" _
                        And Left(VBA.right(.Paragraphs(1).Range, 3), 1) <> "�x" _
                        And Left(VBA.right(.Paragraphs(1).Range, 3), 1) <> "�v" _
                        Then '�p���h�r�ƦӬq��̤~�B�z,�_�h���c�ƤF!(�s���D���]��J,�N�Ӻ��H�F!)2003/11/30
                    Select Case MsgBox("�n�M����?" & vbCr & vbCr & "�n�פ�Ы��e�_�f�I" _
                        , vbYesNoCancel + vbQuestion)
                        Case vbYes  '2003/4/17
                            .Range = ""
                            p = p + 1
                        Case vbCancel
                            If StepByStep = vbYes Then
                                .Range = Chr(13) & .Range '���J�s�q���H�Ϲj�}��!
                                p = p + 1
                            ElseIf StepByStep = NoArrange Then '���B�z
                            Else
                                Stop
                            End If
                        Case vbNo
                            End
                    End Select
                End If
            End If
'            p = p + 1
            d = d - 2 '��������Ÿ�(chr(13)�|�ִ_��Ÿ�(Chr(10)�]������,�G����G
        End If
    Loop
    MsgBox "����" & C & "���ˬd�A" & p & "���m���I", vbInformation
End With
End Sub

Sub �s���v��_�M�����D��()
Dim a As String
With Selection
If .Document.path <> "" Then MsgBox "����󤣯�ާ@", vbExclamation: Exit Sub
    a = .text
    With .Find
        .text = a
        .Parent.Paragraphs(1).Range.Delete
        .Forward = True
        Do While .Execute
            .Parent.Paragraphs(1).Range.Delete
        Loop
    End With
End With
End Sub
Sub �s���v��_�M�����X()
Dim a As Paragraph
With Selection
    If .Document.path <> "" Then MsgBox "����󤣯�ާ@", vbExclamation: Exit Sub
    For Each a In .Document.Paragraphs
        If Len(a.Range) > 2 Then
            If IsNumeric(Mid(a.Range, 2, Len(a.Range) - 4)) Then
                a.Range.Delete
            End If
        End If
    Next a
End With
End Sub

Sub a() '�����z�Ǵ����@�~��2004/5/2
Attribute a.VB_Description = "�������s�� 2004/1/13�A���s�� �]�u�u"
Attribute a.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.����1"
'
' ����1 ����
' �������s�� 2004/1/13�A���s�� �]�u�u
'
Dim i As Long
If MsgBox("�Х��ˬd�X�B�O�_�w�W�ߦ��q���I", vbExclamation + vbOKCancel) = vbOK Then Exit Sub
With Selection
    If .Type = wdSelectionNormal Then .move wdLine, -1
    For i = 1 To .Document.Paragraphs.Count
        Select Case .Paragraphs(1).Range.Font.Name
            Case "�s�ө���"
1               .MoveDown wdParagraph, 1, wdExtend
                If IsNumeric(Left(LTrim(.Paragraphs(1).Range), 1)) And (InStr(1, .Sections(1).Range, _
                    Left(LTrim(.Paragraphs(1).Range), 1), vbBinaryCompare) < InStrRev(.Sections(1).Range, _
                    Left(LTrim(.Paragraphs(1).Range), 1), , vbBinaryCompare)) Then
                    '�n�קK�����}�Q�R!
                    GoTo 2
                End If
                .Application.ScreenRefresh
'                Select Case MsgBox("�T�w�R���H", vbYesNoCancel + vbQuestion)
'                    Case vbYes
                        If InStr(.Paragraphs(1).Range, "---") Then If MsgBox("�T�w�R�����}���j�u�H", vbExclamation + vbOKCancel + vbDefaultButton2) = vbCancel Then GoTo 2
'                        If IsNumeric(.Paragraphs(1).Range.Words(1)) _
                            And (InStr(1, .Sections(1).Range, _
                            .Paragraphs(1).Range.Words(1), vbBinaryCompare) = InStrRev(.Sections(1).Range, _
                            .Paragraphs(1).Range.Words(1), , vbBinaryCompare)) Then
                        If IsNumeric(Left(LTrim(.Paragraphs(1).Range), 1)) _
                            And (InStr(1, .Sections(1).Range, _
                            Left(LTrim(.Paragraphs(1).Range), 1), vbBinaryCompare) = InStrRev(.Sections(1).Range, _
                            Left(LTrim(.Paragraphs(1).Range), 1), , vbBinaryCompare)) Then
                            '�n�����}�R����M�����j�u!(���ɵ��}�s���e�|�Ť@��,�G��W���^
                            .Paragraphs(1).Range.Delete
                            If Not IsNumeric(Left(LTrim(.Paragraphs(1).Range), 1)) And InStr(.Paragraphs(1).Previous.Range, "---") Then
                                .Paragraphs(1).Previous.Range.Delete
                            End If
                        Else
                            .Paragraphs(1).Range.Delete
                        End If
'                    Case vbNo
'                        .MoveDown wdParagraph, 1
'                    Case vbCancel
'                        Exit For
'                End Select
            Case "Times New Roman"
                If InStr(.Paragraphs(1).Range, "---") Then
                    GoTo 1
                Else
                    .MoveDown wdParagraph, 1
                End If
            Case Else
'            If d = 0 Then
2           .MoveDown wdParagraph, 1
        End Select
        If .End + 1 = .Document.Range.End Then MsgBox "���ߧ���!", vbInformation: Exit Sub
    Next i
End With
End Sub
Sub a1()
Selection.Range.Find.Execute "��", , , , , , , , , ChrW(29234), wdReplaceAll
End Sub

Sub �M�����X�аO()
Dim p As Paragraph
For Each p In Documents(1).Paragraphs
    If IsNumeric(p.Range) Then
'        Select Case MsgBox("�T�w�R���H", vbYesNoCancel + vbQuestion)
'            Case vbYes
                p.Range.Select
                word.Application.ScreenRefresh
                p.Range.Delete
'            Case Else
'                Exit For
'        End Select
    End If
Next p
End Sub

Sub �K�W�~�y�j����()
Attribute �K�W�~�y�j����.VB_Description = "�������s�� 2005/2/24�A���s�� �]�u�u"
Attribute �K�W�~�y�j����.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.�K�W�~�y�j����"
'
' �K�W�~�y�j���� ����
' �������s�� 2005/2/24�A���s�� �]�u�u
'
    Documents.Add DocumentType:=wdNewBlankDocument
    Selection.Paste
    ActiveDocument.SaveAs fileName:="�~�y�j����.htm", FileFormat:=wdFormatHTML, _
        LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword _
        :="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
        SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
        False
    ActiveWindow.View.Type = wdWebView
    ActiveWindow.Close
End Sub
Sub �պ`�C�L()
Attribute �պ`�C�L.VB_Description = "�������s�� 2005/4/13�A���s�� Oscar Sun"
Attribute �պ`�C�L.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.�պ`�C�L"
'
' �պ`�C�L ����
' �������s�� 2005/4/13�A���s�� Oscar Sun
'
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(1.54)
        .BottomMargin = CentimetersToPoints(1.54)
        .LeftMargin = CentimetersToPoints(1.17)
        .RightMargin = CentimetersToPoints(1.17)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1.5)
        .FooterDistance = CentimetersToPoints(1.75)
        .PageWidth = CentimetersToPoints(21)
        .PageHeight = CentimetersToPoints(29.7)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .GutterPos = wdGutterPosLeft
        .LayoutMode = wdLayoutModeLineGrid
    End With
    Selection.Sections(1).Footers(1).PageNumbers.Add PageNumberAlignment:= _
        wdAlignPageNumberCenter, FirstPage:=True
    With ActiveDocument.Styles("����")
        .AutomaticallyUpdate = False
        .BaseStyle = ""
        .NextParagraphStyle = "����"
    End With
    With ActiveDocument.Styles("����").Font
        .NameFarEast = "�з���"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Name = "Times New Roman"
        .Size = 12
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 1
        .Animation = wdAnimationNone
        .DisableCharacterSpaceGrid = False
        .EmphasisMark = wdEmphasisMarkNone
    End With
    With ActiveDocument.Styles("����").ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 14
        .Alignment = wdAlignParagraphLeft
        .WidowControl = False
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .AutoAdjustRightIndent = True
        .DisableLineHeightGrid = False
        .FarEastLineBreakControl = True
        .WordWrap = True
        .HangingPunctuation = True
        .HalfWidthPunctuationOnTopOfLine = False
        .AddSpaceBetweenFarEastAndAlpha = True
        .AddSpaceBetweenFarEastAndDigit = True
        .BaseLineAlignment = wdBaselineAlignAuto
    End With
    Selection.Style = ActiveDocument.Styles("����")
    word.Application.PrintOut fileName:="", Range:=wdPrintAllDocument, item:= _
        wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
        ManualDuplexPrint:=False, Collate:=True, Background:=False, PrintToFile:= _
        False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
        PrintZoomPaperHeight:=0
End Sub



Sub ����1()
Attribute ����1.VB_Description = "�������s�� 2008/12/24�A���s�� Oscar Sun"
Attribute ����1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.����1"
'
' ����1 ����
' �������s�� 2008/12/24�A���s�� Oscar Sun
'
    ActiveDocument.SaveAs fileName:= _
        "�_���N�ֶ��]�@�^(588��)-��25(���ժ��f���W�]�бG�T��ܤQ�G��^�бG ����47�~.1782�~.���ͦ~50��).html", _
        FileFormat:=wdFormatText, LockComments:=False, Password:="", _
        AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
        EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
        :=False, SaveAsAOCELetter:=False
End Sub

Sub ����2()
Attribute ����2.VB_Description = "�������s�� 2010/10/28�A���s�� Oscar Sun"
Attribute ����2.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.����2"
'
' ����2 ����
' �������s�� 2010/10/28�A���s�� Oscar Sun
'
    With Selection.ParagraphFormat
        .RightIndent = CentimetersToPoints(8.74)
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
    End With
End Sub
Sub ����3()
Attribute ����3.VB_Description = "�������s�� 2010/10/28�A���s�� Oscar Sun"
Attribute ����3.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.����3"
'
' ����3 ����
' �������s�� 2010/10/28�A���s�� Oscar Sun
'
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(2.54)
        .BottomMargin = CentimetersToPoints(2.54)
        .LeftMargin = CentimetersToPoints(3.17)
        .RightMargin = CentimetersToPoints(12.17)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1.5)
        .FooterDistance = CentimetersToPoints(1.75)
        .PageWidth = CentimetersToPoints(21)
        .PageHeight = CentimetersToPoints(29.7)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .GutterPos = wdGutterPosLeft
        .LayoutMode = wdLayoutModeLineGrid
    End With
End Sub
Sub ����4()
Attribute ����4.VB_Description = "�������s�� 2010/10/28�A���s�� Oscar Sun"
Attribute ����4.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.����4"
'
' ����4 ����
' �������s�� 2010/10/28�A���s�� Oscar Sun
'
End Sub
Sub ����5()
Attribute ����5.VB_Description = "�������s�� 2010/10/28�A���s�� Oscar Sun"
Attribute ����5.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.����5"
'
' ����5 ����
' �������s�� 2010/10/28�A���s�� Oscar Sun
'
    ActiveWindow.ActivePane.View.Zoom.Percentage = 75
End Sub
    



Sub ����6()
Attribute ����6.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.����6"
'
' ����6 ����
'
'
    Selection.MoveLeft unit:=wdCharacter, Count:=1, Extend:=wdExtend
    ActiveDocument.DefaultTargetFrame = ""
    Selection.Range.Hyperlinks(1).Range.Fields(1).result.Select
    Selection.Range.Hyperlinks(1).Delete
    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:= _
        "https://oscarsun72.blogspot.com/2021/02/blog-post_25.html", SubAddress:= _
        "", ScreenTip:="", TextToDisplay:="", Target:="_blank"
    Selection.Collapse Direction:=wdCollapseEnd
End Sub
Sub ����7()
'
' ����7 ����
'
'
    'ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:="", _
        SubAddress:="_���d����", ScreenTip:="", TextToDisplay:=Selection '"�L�ǥO"
    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:="", _
        SubAddress:="_������L", ScreenTip:="", TextToDisplay:=Selection '"�L�ǥO"
    'ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:="", _
        SubAddress:="_" & Selection, ScreenTip:="", TextToDisplay:=Selection '"�L�ǥO"

End Sub
Sub ����8()
Attribute ����8.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.����8"
'
' ����8 ����
'
'
'    Selection.Range.Hyperlinks(1).Range.Fields(1).Result.Select
'    Selection.Range.Hyperlinks(1).Delete
'     ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:="", _
        SubAddress:="_������L�]�v" & ChrW(20008) & "����_��" & ChrW(20008) & "��_����"
    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:="", _
        SubAddress:="������L�]�v" & ChrW(20008) & "����_��" & ChrW(20008) & "��_����_�t" & ChrW(20008) & "���^"
        
'    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:="", _
        SubAddress:="������L"
    
End Sub
Sub ����9()
Attribute ����9.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.����9"
'
' ����9 ����
'
'
    Selection.InsertCrossReference ReferenceType:="���D", ReferenceKind:= _
        wdContentText, ReferenceItem:="241", InsertAsHyperlink:=True, _
        IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
End Sub
Sub ����10()
Attribute ����10.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.����10"
'
' ����10 ����
'
'
    Selection.InsertCrossReference ReferenceType:="���D", ReferenceKind:= _
        wdContentText, ReferenceItem:="6", InsertAsHyperlink:=True, _
        IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=" "
End Sub

Sub �K�Wut���e()
Attribute �K�Wut���e.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.����11"
    Dim tb As Table, s As Long, ur As UndoRecord, rng As Range
    SystemSetup.stopUndo ur, "�K�Wut���e"
    word.Application.ScreenUpdating = False
    s = Selection.start
    If Selection.Type = wdSelectionIP Then
        With Selection.Document
            'If .path = "" Then .Range.Select
            If .path = "" Then Set rng = .Range
            Selection.Paste
            '.Range(s, Selection.End).Select
            Set rng = .Range(s, Selection.End)
        End With
    End If
    'For Each tb In Selection.Document.Tables
    For Each tb In rng.Tables
        tb.Rows.ConvertToText Separator:=wdSeparateByParagraphs, _
            NestedTables:=True
    Next tb
'    Selection.Find.ClearFormatting
'    Selection.Find.Replacement.ClearFormatting
    rng.Find.ClearFormatting
    rng.Find.Replacement.ClearFormatting
    'With Selection.Find
    With rng.Find
        .text = "^p^p"
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
'    Selection.Find.Execute Replace:=wdReplaceAll
'    Selection.Find.Execute Replace:=wdReplaceAll
    rng.Find.Execute Replace:=wdReplaceAll
    rng.Find.Execute Replace:=wdReplaceAll
'    Selection.Copy
    SystemSetup.contiUndo ur
    word.Application.ScreenUpdating = True
    rng.Document.ActiveWindow.ScrollIntoView rng, False
End Sub
