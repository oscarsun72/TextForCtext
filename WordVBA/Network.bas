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
    SeleniumOP.dictRevisedSearch VBA.Replace(Selection, VBA.Chr(13), "")
End Sub

'Sub �^����y���������}()
'SeleniumOP.grabDictRevisedUrl VBA.Replace(Selection, vba.Chr(13), "")
'End Sub
Sub �dGoogle()
    Rem Alt + g
    SeleniumOP.GoogleSearch Selection.text
End Sub
Sub �d�ʫ�()
    Rem Alt b
    SeleniumOP.BaiduSearch Selection
End Sub
Sub �d�r�κ�()
    Rem Alt + z
    If Selection.Characters.Count > 1 Then
        MsgBox "���d1�r", vbExclamation ', vbError
        Exit Sub
    End If
    SeleniumOP.LookupZitools Selection.text
End Sub
Sub �d����r�r��()
    Rem Alt + F12
    If Selection.Characters.Count > 1 Then
        MsgBox "���d1�r", vbExclamation ', vbError
        Exit Sub
    End If
    SeleniumOP.LookupDictionary_of_ChineseCharacterVariants Selection.text
End Sub
Sub �d�d���r����W��()
    Rem Ctrl + Alt + x
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
    If Selection.Characters.Count < 2 Then
        MsgBox "�n2�r�H�W�~���˯��I", vbExclamation ', vbError
        Exit Sub
    End If
    ��r�B�z.ResetSelectionAvoidSymbols
    SeleniumOP.LookupHYDCD Selection.text
End Sub
Sub �d��Ǥj�v()
    Rem Ctrl + d + s �]ds�G�j�v�^
    ��r�B�z.ResetSelectionAvoidSymbols
    SeleniumOP.LookupGXDS Selection.text
End Sub
Sub �d�ն��`�B�H�a����Ѧr�Ϲ��d�\_�ê�k���u��()
    Rem  Alt + s �]���媺���^ Alt + j �]�Ѧr���ѡ^
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
                    ElseIf VBA.Replace(p.Range.text, " ", "") = Chr(13) Then
                        p.Range.Delete
                        GoTo reCheck:
                    ElseIf VBA.Left(p.Range.text, s) = VBA.space(s) Then '�q�`��������
                        p.Range.text = Mid(p.Range.text, s + 1)
                    ElseIf VBA.Left(p.Range.text, sDuan) = VBA.space(sDuan) Then '�q�`�����q�`��
                        With p.Range
                            .text = Mid(p.Range.text, sDuan + 1)
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
            .font.Size = fontsize
            .InsertAfter ar(1) '���J���}
            SystemSetup.contiUndo ur
            .Collapse wdCollapseStart
            With .Application
                .Activate
                With .ActiveWindow
                    If .windowState = wdWindowStateMinimize Then
                        VBA.Interaction.DoEvents
                        .windowState = windowState
                        .Activate
                        VBA.Interaction.DoEvents
                    End If
                End With
            End With
        End With
    End If
End Sub
Sub �d����r�r��è��^�仡���������κ��}�ȴ��J�ܴ��J�I��m()
    Rem  Alt + v �]v= ����r variants �� v�^
    If Selection.Characters.Count > 1 Then
        MsgBox "���d1�r", vbExclamation ', vbError
        Exit Sub
    End If
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
        Dim ur As UndoRecord, fontsize As Single
        SystemSetup.stopUndo ur, "�d����r�r��è��^�仡���������κ��}�ȴ��J�ܴ��J�I��m"
        With Selection
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
            Dim shuoWen As String
            shuoWen = VBA.Replace(VBA.Replace(ar(0), "�G�A", "�G" & x & "�A"), "�q�`���G", VBA.Chr(13) & "�q�`���G")
            If VBA.Left(shuoWen, 1) = "�A" Then
                shuoWen = x & shuoWen
            End If
            If s = 0 And ar(0) <> "�������ΨS����ơI" Then
                .InsertAfter shuoWen & VBA.Chr(13)  'ar(0)=�m����n���e
                .Collapse wdCollapseEnd
            End If
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
                        VBA.Interaction.DoEvents
                        .windowState = windowState
                        .Activate
                        VBA.Interaction.DoEvents
                    End If
                End With
            End With
        End With
    End If
End Sub
Sub �e��j�y�Ŧ۰ʼ��I()
    'Alt + F10
    Dim ur As UndoRecord
    If Selection.Characters.Count < 10 Then
        MsgBox "�r�ƤӤ֡A�����n�ܡH�Цܤ֤j��10�r", vbExclamation
        Exit Sub
    End If
    Selection.Copy
    TextForCtext.GjcoolPunct
    Selection.Document.Activate
    Selection.Document.Application.Activate
    SystemSetup.stopUndo ur, "�e��j�y�Ŧ۰ʼ��I"
    Selection.text = SystemSetup.GetClipboardText
    SystemSetup.contiUndo ur
End Sub

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



