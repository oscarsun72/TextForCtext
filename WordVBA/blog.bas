Attribute VB_Name = "blog"
Option Explicit '�_��������M�μҲ�
Public OX As Object, htmFilename As String, myaccess As Object 'As Access.Application
Dim es As Byte '�O�U���t2007/11/1
Sub setOX()
On Error Resume Next
If OX Is Nothing Then Set OX = CreateObject("AutoItX3.Control")
End Sub
Sub �x�s��google���ިóƤ�(dp As Document)  '2008/12/24
    htmFilename = InputBox("�п�J�ɦW", , "htmfilename")
    If htmFilename = "" Then Exit Sub
    On Error GoTo eH:
    dp.SaveAs fileName:= _
        "P:\�ڪ�������\5160_\" & VBA.Left(htmFilename, 235) & ".html", _
        FileFormat:=wdFormatUnicodeText, LockComments:=False, Password:="", _
        AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
        EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
        :=False, SaveAsAOCELetter:=False
'    Windows("�_���N�ֶ��]�@�^(588��)-��25(���ժ��f���W�]�бG�T��ܤQ�G��^�бG ����47�~.1782�~.���ͦ~50��)"). _
'        Activate
'    Windows("�_���N�ֶ��]�@�^(588��)-��25(���ժ��f���W�]�бG�T��ܤQ�G��^�бG ����47�~.1782�~.���ͦ~50��)"). _
'        Activate
    setOX
    OX.WinActivate "iexplorer"
    'AppActivate "iexplorer"
Exit Sub
eH:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & " - " & Err.Description
        htmFilename = InputBox("�п�J�ɦW", , htmFilename)
        Resume
End Select
End Sub


Sub ���J��ï�Ϥ��s�����}() '�������'2006/4/4
Dim x As String, i As Long 'Integer
x = InputBox("�п�J��1�ӹϤ����s�����}")
'If x = "" Then End 'Exit Sub
If x = "" Or InStr(x, "http") = 0 Then End
If InStr(x, "&prev") > 0 Then x = VBA.Left(x, InStr(x, "&prev") - 1)
i = Int(VBA.Mid(x, InStrRev(x, "=") + 1))
x = VBA.Left(x, InStrRev(x, "="))
options.AutoFormatAsYouTypeReplaceQuotes = False '�������z�޸�,�p��"�~���|�Q�۰ʸm�������]�b�۰ʮե��̷̨ӱz����J�۰ʮ榡���Ҹ̡^
With ActiveDocument.Range
    With .Find
        .MatchWildcards = False
        .ClearFormatting
        .text = "<img src="
        .Forward = True
        .Wrap = wdFindStop
        .Replacement.text = "<a href=""" & x & i & """><img src="
        Do While .Execute(, , , , , , , , , , wdReplaceOne)
            .Parent.Move
            i = i + 1
            With .Replacement
                .text = "<a href=""" & x & i & """><img src="
            End With
        Loop
        .Parent.Move wdStory
        .Forward = False
        .text = ".jpg"" />"
        .Replacement.text = ".jpg"" /></a>"
        .Execute , , , , , , , , , , wdReplaceAll
    End With
End With
options.AutoFormatAsYouTypeReplaceQuotes = True '��_���z�޸�
End Sub


Sub �Ϥ����J������w��m()
Dim x As Range, s As String, e As String, t As String, sy As Boolean
Static sn As Long '�O�U����
Dim d As Document ', dp As Document
Dim ts As Byte
Set d = ActiveDocument
'Set dp = d.Windows(1).Previous.Document
On Error GoTo errhan:
With d.Windows(1).Previous.Document '�۰��������e���}���s���(��ܳ�����W�Ǥw���\)2007/10/28
    If .path = "" Then .Close wdDoNotSaveChanges
'    If d.Application.Documents.Count > 2 Then
'        If .Path = "" Then
'            If .Range = vba.Chr(13) Then .Undo  '�٭�ŤU�e
'            If InStr(.Range, "<p><a href=""http://tw.myblog.yahoo.com/") Then
'                Set dp = d.Windows(1).Previous.Document
'                �x�s��google���ިóƤ� dp
'            End If
'            .Close wdDoNotSaveChanges
'        End If
'    End If
End With
d.Activate
s = InputBox("�п�J��1��", , sn + 1)
If s = "" Then Exit Sub '�H10�������
If InStr(s, ".") Then sy = True: ts = 1 '�Ѱ_�l�P�_�O�_��"��" 2009/11/13
's = CInt(s)
s = CSng(s)
'e = s + 10
If es = 0 Then es = 9
e = InputBox("�п�J�̫�1��", , s + es) ' + 9)
If sy = False Then If InStr(e, ".") Then sy = True: ts = 1 '�Ѱ_�l�P�_�O�_��"��" 2009/11/13
If e = "" Then Exit Sub
'e = CInt(e)
e = CSng(e)
If e <> 0 Then sn = e: es = e - s
'If e = 0 Then e = s + 9 '�ٲ��Y�H10�������
's = 261: e = 270
With d.Range
    Do Until s >= e + 0.1 '+ 1
'    s = e '�˧ǮɽХγo���
'    Do Until s = e - 10
'        .move wdStory
        Set x = d.Range(InStr(d.Range, "<img src=""") - 1, InStr(d.Range, ".jpg"" />") + Len(".jpg"" />") - 1)
        t = x.text
        .SetRange InStr(d.Range, "<img src=""") - 1, InStr(d.Range, ".jpg"" /><br>") + Len(".jpg"" /><br>") - 1
        .Delete
        With Selection
            .HomeKey wdStory
            .Find.ClearFormatting
            .Find.MatchWildcards = False
            .Find.Forward = True
            'If .Find.Execute(">" & s & "<") Then
            If .Find.Execute("<p class=MsoNormal style=""MARGIN: 0cm 0cm 0pt""><strong><span lang=EN-US style=""FONT-SIZE: 16pt; BACKGROUND: #efef5a""><font face=""Times New Roman"">" _
                    & s & "</font></span></strong></p>") Then
                .Move wdParagraph
                .InsertAfter "<p>" & t & "</p>"
            ElseIf .Find.Execute("<p class=MsoNormal style=""MARGIN: 0cm 0cm 0pt""><strong><span lang=EN-US style=""FONT-SIZE: 16pt;  COLOR: blue""><font face=""Times New Roman"">" _
                    & s & "</font></span></strong></p>") Then
                .Move wdParagraph
                .InsertAfter "<p>" & t & "</p>"
            ElseIf .Find.Execute("<p class=""MsoNormal"" style=""MARGIN:0cm 0cm 0pt;""><strong><span style=""FONT-SIZE:16pt;COLOR:blue;""><font face=""Times New Roman"">" _
                    & s & "</font></span></strong></p>") Then
                .Move wdParagraph
                .InsertAfter "<p>" & t & "</p>"
            ElseIf .Find.Execute("<p class=MsoNormal style=""MARGIN: 0cm 0cm 0pt""><strong><span lang=EN-US style=""FONT-SIZE: 16pt; COLOR: blue""><font face=""Times New Roman"">" _
                    & s & "</font></span></strong></p>") Then
                .Move wdParagraph
                .InsertAfter "<p>" & t & "</p>"
            ElseIf .Find.Execute(">" & s & "<o:p></o:p></font></span></b></p>") Then
'            If .Find.Execute("<span lang=EN-US style=""FONT-SIZE: 16pt; COLOR: blue; mso-bidi-font-size: 12.0pt; mso-text-animation: ants-red""><font face=""Times New Roman"">" & s & "<") Then '�e��즡�|�b���X�{�Ʀr���ɿ���!2006/8/15
                .Move wdParagraph
                .InsertAfter "<p>" & t & "</p>"
            ElseIf .Find.Execute(">" & s & "<?xml:namespace prefix = o") Then
                .Move wdParagraph
                .InsertAfter "<p>" & t & "</p>"
            ElseIf .Find.Execute(">" & s & "<") Then
                .Move wdParagraph
                .InsertAfter "<p>" & t & "</p>"
'            Else
'                Exit Do
            End If
        End With
        If Not sy Then
            s = s + 1 '�D����
        Else
'����:       If ts Mod 2 = 1 Then '����
'                s = s - 0.1 + 1
'            Else
'                s = s + 0.1
'            End If
'����end:        ts = ts + 1
            If InStr(s, ".") = 0 Then
                s = s + 0.1
            Else
                s = s - 0.1 + 1
            End If
        End If
'        s = s - 1 '�˧ǮɽХγo�@��
    Loop
End With
Exit Sub
errhan:
Select Case Err.Number
    Case 91 '�S���]�w�����ܼƩ� With �϶��ܼ�
        Resume Next '��ܤ��e�S���}�Ҫ��s���.
End Select
End Sub
Sub �ƦC�Ϥ��ᴡ�J�s��()
If ActiveDocument.path <> "" Then Documents.Add DocumentType:=wdNewBlankDocument
options.AutoFormatAsYouTypeReplaceQuotes = False '��_���z�޸�
With ActiveDocument.Range
    .Select
    .Paste
End With
�Ϥ����J������w��m
���J��ï�Ϥ��s�����}
ActiveDocument.Range.Copy
options.AutoFormatAsYouTypeReplaceQuotes = True '��_���z�޸�
End Sub
Sub ���N���s���Ϥ����}()
With ActiveDocument.Range
'    With .Find
'        .Text = "src="""
'        .Forward = True
'        Do Until .Execute = False
'
'            With .Parent
'                .move
'                .MoveUntil .Text = "h", wdExtend
'            End With
'        Loop
'    End With
End With
End Sub

Sub ��J�W�ǹϤ���}()
Dim x As String, i As Byte
Static sn As Long '�O�U����
x = InputBox("�п�J�Ĥ@�ӹϤ������X", , sn + 1)
If x = "" Then Exit Sub
If IsNumeric(x) Then sn = x + es
'AppActivate "avant browser"
AppActivate "explorer"
'AppActivate "mozilla firefox"'mozilla firefox�����
DoEvents
SendKeys "{tab 3}" & ���o�ୱ���| & "\���ե�\�ܧ��ɦW��\" & Format(x, "_000000") & ".jpg"
For i = 1 To 9 '�q��2�ӹϤ����10��
    DoEvents
    SendKeys "{tab 4}" & ���o�ୱ���| & "\���ե�\�ܧ��ɦW��\" & Format(x + i, "_000000") & ".jpg"
Next i
DoEvents
'SendKeys "{tab 3}{right}"'����
SendKeys "{tab 3}"
End Sub
Sub ����ŤU���������() '�H�Q�K�W������]'2007/10/30-�K�W�W���U���ѭ��^�`�ت�html�X.
Dim Dnow As Document, bt As String, hide As Boolean
With Selection
    If ActiveDocument.path = "" Then
        Set Dnow = ActiveDocument
        If InStr(Dnow.Range, "<a href=") Then '��ܭn�K�Whtml�X�F!
            Dim o As Boolean, d, h As String
            For Each d In Documents
                If d.Name = "�Ȧs.doc" Then o = True: Exit For
            Next
            If o = False Then
                Documents.Open ���o�ୱ���| & "\�Ȧs.doc"
            Else
                Documents("�Ȧs").Activate
            End If
            o = True '�O�U�w�O�b�B�zhtml�X,�H�ѤU����.2007/11/4
            With ActiveWindow.Selection
                If Len(.text) = 1 Then .GoTo wdGoToBookmark, , , "���_�Ȧs"
                bt = .Range
                h = InputBox("�п�J�W����}")
                'If h = "" Or InStr(h, "http") = 0 Then Exit Sub
                If h = "" Then Exit Sub
                If InStr(h, "http") <> 0 Then
                    If InStr(h, "&prev") > 0 Then h = VBA.Left(h, InStr(h, "&prev") - 1)
                    bt = Replace(bt, "�W��", "<a href=""" & h & """>" & "�W��</a>") '���J�W����}
                End If
                '.Parent.WindowState wdWindowStateMinimize
                ActiveWindow.Visible = False
            End With
            With Dnow
                .Activate
                .Range = bt & .Range & bt
            End With
        End If
'        If InStr(.Document.Range, ": �з���""") Then .Document.Range = �h�з���r(.Document.Range)
        If InStr(.Document.Range, "<img") Then '�P�_�����i��,�_�h�|�b�D���X�ɰ���2009/4/17
            .Document.Range = �h�з���r(.Document.Range)
            .EndKey wdStory, wdMove
            Do Until Asc(.Previous) <> 13
                .TypeBackspace
            Loop
             '�b�������J�b�u�H��
            .Document.Range = "<p><a href=""http://whos.amung.us/stats/s5z4puepm2vb/""><img title=""Click to see how many people are online"" src=""http://whos.amung.us/widget/s5z4puepm2vb.png"" border=""0"" height=""29"" width=""81"" /></a></p> " & .Document.Range
            '.Document.Range = .Document.Range & "<p style=""text-align: right;""><a href=""""><span style=""color: red;"">�Цh����(comments) <span style=""color: rgb(51, 102, 255);""><br><font size=1>�������U,�����n�J,<br>�i�ΦW�d��</font></span></span></a></p>"
            '.Document.Range = .Document.Range & "<p style=""text-align: right;""><a href=""http://www.blogger.com/comment.g?blogID=37481082&amp;postID=116317965718083160""><span style=""color: red;"">�w�����(comments) <span style=""color: rgb(51, 102, 255);""><br><font size=1>�������U,�����n�J,<br>�i�ΦW�d��</font></span></span></a></p>"
            .Document.Range = .Document.Range & "<p style=""text-align: right;""><a href=""http://www.blogger.com/comment.g?blogID=37481082&amp;postID=116317965718083160""><span style=""color: red;"">�S���F��O���ձo�Ӫ�.���u��,�Ф��n����Ӧh.<br>�o���ӴN�O������Ū��,�B�uŪ�Ѫ��n�Ǫ̭׾Ǫ̥�y�����x,���O�ѾǳN�ξǥͨ����\�W���e�B.<br>�o�O�U�H,�Ӥ��O�`�H�����l.�󤣬O�ڨө񰪧Q�U���\�w. <br>�Цh�������ЩΥ���(comments) �ۤߥ�y,�����ۤ�,�ű��r�u.<br>�_�h���@�z�h�I�O���Ҧb,�ȳf��W,�]�n�vū. <br>�P�§A,�]��ū�F�ۤv.�Z�����i,�ͥͥ@�@���w.<br>���u�S�Q��,�o�رШ|�����٭n�d���ڤ��Ѩӻ�.��p�O����������������.<br>�ڭ̵L�����G���s���s�b,���O,��,�s�b<span style=""color: rgb(51, 102, 255);""><br><font size=1>�������U,�����n�J,<br>�I�����B�Y�i�ΦW�d��<br>�d���e,�Х����,�άݬݦۤv�N�h���ߦ~<br>���d��,�u���A�|�ѤF,�ڷ|����,�ƹ�|�æs</font></span></span></a></p>"
        End If
        .WholeStory
        .Cut
'        .Document.Close wdDoNotSaveChanges'2007/11/2�]���_���`��,�G���������H�K�٭�
'        If o Then
'            If MsgBox("�O�_����?", vbOKCancel) = vbOK Then
'                hide = True
'            Else
                hide = False
'            End If
'        End If
'        AppActivate "Avant Browser"
        AppActivate "explorer"
'        AppActivate "mozilla firefox"
'        SendKeys "+{insert}"
        SendKeys "2"
        If o Then '2007/11/4
            DoEvents
            If Not hide Then
                SendKeys "{tab 3}" '���}�K�l
            Else
                SendKeys "{tab 3}{right}" '���öK�l
            End If
            SendKeys "{tab 4}{enter}" '�o��K�l
        End If '2007/11/4
    Else
'        .Document.Close wdDoNotSaveChanges
    End If
End With

End Sub

Sub �ƻs�K�W�ѭ���T()
Dim o As Boolean, d
For Each d In Documents
    If d.Name = "�Ȧs" Then o = True: Exit For
Next
If o = False Then
    Documents.Open ���o�ୱ���| & "\�Ȧs.doc"
Else
    Documents("�Ȧs").Activate
End If
With ActiveWindow.Selection
    If Len(.text) = 1 Then .GoTo wdGoToBookmark, , , "���_�Ȧs"
    .Copy
    '.Parent.WindowState wdWindowStateMinimize
    ActiveWindow.Visible = False
'    AppActivate "Avant Browser"
    AppActivate "explorer"
    DoEvents
    SendKeys "+{insert}"
End With
End Sub

Sub ���ͮѭ����X()
Static a As String, e As String ', s As String
Dim i As Long, d As Document
a = InputBox("�п�J���j����", "���ͮѭ����X", 10)
If a = "" Then Exit Sub
i = InputBox("�п�J�_�l���X", "���ͮѭ����X", 1)
'If i = "" Then Exit Sub
e = InputBox("�п�J�������X", "���ͮѭ����X")
If e = "" Then Exit Sub
e = VBA.StrConv(e, vbNarrow)
'i = VBA.StrConv(i, vbNarrow)
'a = VBA.StrConv(a, vbNarrow)
Set d = Documents.Add
With d
    Do Until i > e
        If i + a > e Then
            .Range = .Range & i & "-" & e & "�ѥ� "
        Else
            .Range = .Range & i & "-" & i + a - 1 & " "
        End If
        i = i + a
    Loop

.Range = Replace(d.Range, VBA.Chr(13), "")
.Range.font.Size = 8
End With
End Sub

Sub ���ƹϧ�() '2008/7/7 �N��1�i�m�_�̫�@�i,�̦�����
Dim i As String, x As String, d As Document, r As Range
Dim s As String, e As String, sp As Long, eP As Long
Set d = Documents.Add
s = "<img src="
e = "/><br>"
With d
    .Range.Paste
    Do
        x = .Range.text
        eP = InStr(x, e) + Len(e) - 1
        sp = InStr(x, s)
        If sp = 0 Then Exit Do
        Set r = d.Range(0, eP)
        i = r & i
        r.Delete
    Loop
    .Range = i
    .Range.Select
    .Range.Cut
    .Close wdDoNotSaveChanges
End With
'setOX
'OX.WinActivate "explorer"
AppActivate "explorer"
End Sub

Sub �}�ҶW�s��()
'Ctrl + i   �t�ιw�]�O italic�]�Y����r�^�A���t�X ExcelVBA�]�w
Dim rng As Range
Set rng = Selection.Range
If rng.Hyperlinks.Count = 0 Then '�p�G�Ҧb��m�S���W�s���A�h�ݨ�e���_�F�Y�S�L�A�h�A�ݨ�ᦳ�_�F�Y���L�h������ 2022/12/20
    If rng.start = 0 Then
        If Selection.Type = wdSelectionNormal Then
            GoTo Selected_Range
        Else
            GoSub nxt
        End If
    ElseIf rng.End = rng.Document.Range.End - 1 Then
        GoSub pre
    ElseIf rng.End < rng.Document.Range.End - 1 Then
Selected_Range:
        '�� �u�����r�[�W��y���`���v �ӳ]
        Dim rngNext As Range
        Set rngNext = rng.Next
        If rngNext.text = "�]" Then
            If rngNext.Next.Hyperlinks.Count > 0 Then
                rng.SetRange rngNext.Next.start, rngNext.Next.End
                GoSub nxt
            End If
        Else
            GoSub Position
        End If
    Else
        GoSub Position
    End If
End If
If rng.Hyperlinks.Count > 0 Then
    Dim strLnk As String, lnk As Hyperlink
    Set lnk = rng.Hyperlinks(1)
    If rng.Hyperlinks(1).SubAddress <> "" Then
        Dim subAdrs As String
        subAdrs = rng.Hyperlinks(1).SubAddress
        strLnk = rng.Hyperlinks(1).Address + "#" + _
            VBA.IIf(VBA.InStr(subAdrs, "%"), subAdrs, _
            IIf(code.IsSurrogate(subAdrs), UrlEncode(subAdrs), code.UrlEncode_Big5UnicodOLNLY(subAdrs)))
    Else
        strLnk = rng.Hyperlinks(1).Address
    End If
    SystemSetup.playSound 0.484
    Shell getDefaultBrowserFullname + " " + strLnk + " --remote-debugging-port=9222 "
End If
Exit Sub
Position:
    If rng.Previous.Hyperlinks.Count > 0 Then
pre:        Set rng = rng.Previous
    ElseIf rng.Next.Hyperlinks.Count > 0 Then
nxt:        Set rng = rng.Next
    End If
Return
End Sub
Sub ���J�W�s��() '2008/9/1 ���w��(�ֱ���) Ctrl+shift+K(��t�Ϋ��w�bsmallcaps��)
'Alt+k
'    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
'    Selection.Range.Hyperlinks(1).Range.Fields(1).Result.Select
'    Selection.Range.Hyperlinks(1).Delete
'setOX
'    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:= _
        VBA.StrConv(OX.ClipGet, vbNarrow) _
        , SubAddress:=""
'    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:= _
        VBA.StrConv(GetClipboard, vbNarrow) _
        , SubAddress:="" '    Selection.Collapse Direction:=wdCollapseEnd
        
    Dim lnk As String
    'lnk = UrlEncode(SystemSetup.GetClipboardText)
    lnk = SystemSetup.GetClipboardText
    If VBA.InStr(lnk, "http") = 0 Or VBA.InStr(lnk, "http") > 1 Then MsgBox "�ŶKï���D���ĺ��}�I": Exit Sub
            
    Dim rng As Range, b As Boolean, ur As UndoRecord ', wndo As Window ', d As Document ', sty As String
    
    
    Set rng = Selection.Range ': Set wndo = ActiveWindow
    If rng.Information(wdInFootnote) Then
        If rng.Document.Windows.Count > 1 Then '�]���}�h�����ɭY�b��1�ӥH�~�����������}������urng.Hyperlinks.Add�v�h�|�~�����1�ӵ����餤
            Dim wnd As Window, ww, i As Byte
            Set wnd = ActiveWindow
            If CByte(VBA.Right(wnd.Caption, 1)) > 1 Then
                Dim wnds() As Object
                For Each ww In rng.Document.Windows
                    ReDim Preserve wnds(i)
                    Set wnds(i) = ww
                    i = i + 1
                Next
            End If
        End If
    End If
    'Set ur = SystemSetup.stopUndo("")
     SystemSetup.stopUndo ur, ""
    'Set d = ActiveDocument
'    If d.path <> "" Then d.Save
    b = rng.Bold ': sty = rng.Style
    'wndo.Activate �b�Ҷ}�h���������G�u�|����Ĥ@�ӵ���������m
    If i > 0 Then
        Dim slRng() As Object
        i = 0
        For Each ww In wnds
            If Not ww Is wnd Then '����� ww.Caption <> wnd.Caption �]���U����ww.Close �����@�������A�hCaption�ݩʤ]�|����
                ReDim Preserve slRng(i)
                Set slRng(i) = ww.Selection.Range
                ww.Close '�J�M������ܵ����e��A�N�u��������A�B�O�U���ЩҦb��m�F
                i = i + 1
            End If
        Next
    End If
    Dim ssharp As String
    ssharp = InStr(lnk, "#")
    If ssharp > 0 Then
        Dim w As String
        lnk = VBA.Replace(lnk, VBA.ChrW(-9217) & VBA.ChrW(-8195), "�@")
        w = VBA.Mid(lnk, ssharp + 1, Len(lnk) - ssharp)
        w = code.UrlEncode(w)   'byRef
        lnk = VBA.Mid(lnk, 1, ssharp) + w
    End If
    rng.Hyperlinks.Add Anchor:=rng, Address:= _
        VBA.StrConv(lnk, vbNarrow) _
        , SubAddress:="", Target:="_blank" '    Selection.Collapse Direction:=wdCollapseEnd
    'wndo.Activate
    If rng.Bold <> b Then rng.Bold = b
    'If rng.Style <> sty Then rng.Style = sty
    SystemSetup.contiUndo ur
    Set ur = Nothing
    If i > 0 Then
        i = 0
        For Each ww In wnds
            If Not ww Is wnd Then
                Dim wwp As Window
                Set wwp = rng.Document.Windows.Add()
                wwp.Activate
                If slRng(i).Information(wdInFootnote) Then
                    With wwp
                        If .Panes.Count = 1 Then
                            '�}�ҵ��}����
                            If .View.Type = wdNormalView Then _
                               .View.SplitSpecial = wdPaneFootnotes
                        Else
                            .ActivePane.Next.Activate
                        End If
    '                    .ScrollIntoView .ActivePane.Selection, True
    '                    .ActivePane.SmallScroll
                    End With
                End If
                slRng(i).Select
                i = i + 1
            End If
        Next
        wnd.Activate
    End If
    If rng.Document.path <> "" Then rng.Document.Save
End Sub

Sub insertHydzdLink()
Dim lk As New Links, db As New dBase
db.setWordControlValue (��r�B�z.trimStrForSearch(Selection.text, Selection))
db.setDictControlValue 3
lk.insertLinktoHydzd
Set lk = Nothing: Set db = Nothing
End Sub
Sub insertHydcdLink()
Dim lk As New Links, db As New dBase
db.setWordControlValue (��r�B�z.trimStrForSearch(Selection.text, Selection))
db.setDictControlValue 4
lk.insertLinktoHydcd
Set lk = Nothing: Set db = Nothing
End Sub
Sub updateURL��y���()
Dim lnks As New Links
lnks.updateURL��y��� ActiveDocument
'SystemSetup.playSound 7
MsgBox "done!", vbInformation
End Sub
Sub saveV5URL()
Dim ac As Object, lnk As String
Dim dbFullName As String
dbFullName = UserProfilePath & "Dropbox\�m���s��y���׭q���n��Ʈw.mdb"
If Selection.Hyperlinks.Count > 0 Then
    lnk = Selection.Hyperlinks(1).Address
Else
    Selection.MoveRight wdCharacter, 1, wdExtend
    lnk = Selection.Hyperlinks(1).Address
End If
Set ac = GetObject(dbFullName).Application
ac.Run "saveV5URL", lnk
AppActivate "access"
End Sub
Sub updateURL��Ǥj�v()
Dim lnks As New Links
lnks.updateURL��Ǥj�v ActiveDocument
'SystemSetup.playSound 7
End Sub
Sub ���D��r()
With Selection.font
    .Size = 20
    .Bold = True
End With
End Sub

Function �h�з���r(r As Range)
'If InStr(r, ": �з���""") Then
    r = Replace(r, "<span style=""FONT-FAMILY: �з���"">", "<span>")
    r = Replace(r, "; FONT-FAMILY: �з���"">", """>")
    r = Replace(r, "FONT-FAMILY: �з���;", "")
    r = Replace(r, "; mso-fareast-font-family: �з���", "")
    r = Replace(r, "mso-fareast-font-family: �з���", "")
    r = Replace(r, "font-family: �з���; ", "") '2009/4/17�ɦ���
    
    r = Replace(r, "<span style=""FONT-FAMILY: �s�ө���"">", "<span>")
    r = Replace(r, "; FONT-FAMILY: �s�ө���"">", """>")
    r = Replace(r, "FONT-FAMILY: �s�ө���;", "")
    r = Replace(r, "; mso-fareast-font-family: �s�ө���", "")
    r = Replace(r, "mso-fareast-font-family: �s�ө���", "")
    
    r = Replace(r, "<span style=""FONT-FAMILY: �s�ө���; mso-ascii-font-family: 'Times New Roman'; mso-hansi-font-family: 'Times New Roman'"">", "<span>")
    r = Replace(r, "; FONT-FAMILY: �s�ө���", "")
    r = Replace(r, "; mso-bidi-font-size: 12.0pt; mso-ascii-font-family: 'Times New Roman'; mso-hansi-font-family: 'Times New Roman'", "")
    r = Replace(r, "; mso-bidi-font-size: 12.0pt", "")
    r = Replace(r, "<span lang=EN-US>", "<span>")
    r = Replace(r, " lang=EN-US", "")
    r = Replace(r, "<span style=""FONT-SIZE: 8pt; COLOR: navy""><font face=""Times New Roman"">", "<span style=""COLOR: navy""><font face=""Times New Roman"" size=2>")
    r = Replace(r, "; mso-text-animation: ants-red", "")
    r = Replace(r, "</span><span style=""COLOR: navy"">", "")
    r = Replace(r, " size=2>", ">")
    r = Replace(r, "<span style=""FONT-SIZE: 8pt; COLOR: navy;  mso-bidi-font-size: 12.0pt; mso-ascii-font-family: 'Times New Roman'; mso-hansi-font-family: 'Times New Roman'"">", "<span style=""FONT-SIZE: 8pt; COLOR: navy"">")
    r = Replace(r, "<span style="" mso-ascii-font-family: 'Times New Roman'; mso-hansi-font-family: 'Times New Roman'"">", "<span>")
    r = Replace(r, "<span style=""mso-tab-count: 1"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>", "")
    r = Replace(r, "<span style=""mso-tab-count: 1""><font face=""Times New Roman"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font></span>", "")
    r = Replace(r, "<span style=""mso-tab-count: 1"">&nbsp;&nbsp;&nbsp;&nbsp; </span>", "")
    r = Replace(r, ";  mso-ascii-font-family: 'Times New Roman'; mso-hansi-font-family: 'Times New Roman'", "")
    r = Replace(r, "<span style=""mso-tab-count: 1"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>", "")
    r = Replace(r, "<span style=""mso-tab-count: 3"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>", "")
    r = Replace(r, "<font face=""Times New Roman"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font>", "")
    r = Replace(r, "<span style=""mso-tab-count: 1"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span>", "")
    r = Replace(r, "<span style=""mso-tab-count: 1""><font face=""Times New Roman"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font></span>", "")
    r = Replace(r, "</span><span style=""FONT-SIZE: 8pt; COLOR: navy""><span style=""mso-spacerun: yes""><font face=""Times New Roman"">&nbsp; </font></span></span><span style=""FONT-SIZE: 8pt; COLOR: navy"">", "&nbsp; ")
    
    r = Replace(r, "<font face=""Times New Roman""> </font></span><span style=""FONT-SIZE: 8pt; COLOR: navy"">", " ")
    
    r = Replace(r, ".</font></span><span style=""FONT-SIZE: 8pt; COLOR: navy"">", ".</font>", , , vbBinaryCompare)
    r = Replace(r, "\</font></span><span style=""FONT-SIZE: 8pt; COLOR: navy"">", "\</font>", , , vbBinaryCompare)
    r = Replace(r, "_</font></span><span style=""FONT-SIZE: 8pt; COLOR: navy"">", "_</font>", , , vbBinaryCompare)
    r = Replace(r, "<font face=""Times New Roman"">-</font></span><span style=""FONT-SIZE: 8pt; COLOR: navy"">", "<font face=""Times New Roman"">-</font>", , , vbBinaryCompare)
    
    r = Replace(r, "&nbsp;&nbsp;&nbsp;&nbsp;", "&nbsp;&nbsp;")
    r = Replace(r, "<font face=""Times New Roman"">&nbsp;&nbsp;&nbsp; </font>", "&nbsp;&nbsp;&nbsp; ")
    r = Replace(r, "<font face=""Times New Roman"">&nbsp;&nbsp; </font>", "&nbsp;&nbsp; ")
    r = Replace(r, "<font face=""Times New Roman"">&nbsp; </font>", "&nbsp; ")
    r = Replace(r, "<span style=""mso-tab-count: 1"">&nbsp;&nbsp;&nbsp; </span>", "&nbsp;&nbsp;&nbsp; ")
    r = Replace(r, "<span style=""mso-tab-count: 1"">&nbsp;&nbsp; </span>", "&nbsp;&nbsp; ")
    r = Replace(r, "<span style=""mso-tab-count: 1"">&nbsp; </span>", "&nbsp; ")
    r = Replace(r, "<span style=""mso-tab-count: 2"">&nbsp;&nbsp;&nbsp; </span>", "&nbsp;&nbsp;&nbsp; ")
    r = Replace(r, "<span style=""mso-tab-count: 2"">&nbsp;&nbsp; </span>", "&nbsp;&nbsp; ")
    r = Replace(r, "<span style=""mso-tab-count: 2"">&nbsp; </span>", "&nbsp; ")

    r = Replace(r, "&nbsp;&nbsp;&nbsp;", "�@�@�@")
    r = Replace(r, "&nbsp;&nbsp;", "�@�@")
    r = Replace(r, "&nbsp;", " ")
    r = Replace(r, "<font face=""Times New Roman""> </font>", " ")
    r = Replace(r, "<p class=MsoNormal style=""MARGIN: 0cm 0cm 0pt; TEXT-INDENT: 24pt"">", "<p>")
    
    r = Replace(r, "<span style=""FONT-SIZE: 8pt; COLOR: navy""><span style=""mso-tab-count: 1""> </span></span>", " ")
    r = Replace(r, "<span style=""mso-spacerun: yes"">  </span>", "")
    r = Replace(r, ";  mso-bidi-font-size: 12.0pt; mso-ascii-font-family: 'Times New Roman'; mso-hansi-font-family: 'Times New Roman'"">", """>")
    
'    r = Replace(r, "�u���G</span></strong><span style=""FONT-SIZE: 8pt; COLOR: navy"">", "�u���G</strong>")
    r = Replace(r, "<strong><span style=""FONT-SIZE: 8pt; COLOR: navy"">�u���G</span></strong><span style=""FONT-SIZE: 8pt; COLOR: navy"">" _
            , "<span style=""FONT-SIZE: 8pt; COLOR: navy""><strong>�u���G</strong>")
    
    
    
    r = Replace(r, "; mso-bidi-font-family: 'Times New Roman'; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA", "")
    r = Replace(r, "; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA", "")
    r = Replace(r, "; FONT-FAMILY: 'Times New Roman'", "")
    
    
    
    �h�з���r = r
'End If
End Function
Sub �h�з���rs()
Dim r As Range
Set r = ActiveDocument.Range
'If InStr(r, ": �з���""") Then
'    r = Replace(r, "<span style=""FONT-FAMILY: �з���"">", "<span>")
'    r = Replace(r, "; FONT-FAMILY: �з���"">", """>")
'    r = Replace(r, "FONT-FAMILY: �з���;", "")
'    r = Replace(r, "; mso-fareast-font-family: �з���", "")
'    r = Replace(r, "mso-fareast-font-family: �з���", "")
'
'    r = Replace(r, "<span style=""FONT-FAMILY: �s�ө���"">", "<span>")
'    r = Replace(r, "; FONT-FAMILY: �s�ө���"">", """>")
'    r = Replace(r, "FONT-FAMILY: �s�ө���;", "")
'    r = Replace(r, "; mso-fareast-font-family: �s�ө���", "")
'    r = Replace(r, "mso-fareast-font-family: �s�ө���", "")
    
    r = �h�з���r(r)
    'r = Replace(r, "", "")
    
    With ActiveDocument
        .Range = r
        With .Windows(1)
            .Selection.EndKey wdStory, wdMove
            Do Until .Selection.Previous <> VBA.Chr(13)
                .Selection.TypeBackspace
            Loop
        End With
        .Range.WholeStory
        .Range.Cut
    End With
'End If
End Sub

Sub �ˬd�ýX�ݸ�() '2008/8/19�����ɦW�̦��ýX?�Hê�N���Ƥ��].
Dim d  As Document
Set d = Documents.Add
'd.Range.Paste
d.Range.PasteAndFormat (wdFormatPlainText)
If InStr(d.Range, "?") Then
    MsgBox "���ýX!!", vbCritical
    With d.Windows(1).Selection.Find
        .ClearFormatting
        .Execute "?"
    End With
Else
    d.Close wdDoNotSaveChanges
End If
End Sub

Sub �ѵ{���X�����X�Ϥ��[�s���ñƧ�()
Dim d, x As String, a, s As Long, e As Long, p As String
Set d = ActiveDocument
x = d.Range
s = InStr(x, "<p><a href=""http://tw.myblog.yahoo.com/jw%21ob4NscCdAxS_yWJbxTvlgfR./photo?pid=")
If s Then
    e = InStr(x, ".jpg"" /></a></p>")
    With d
        Do Until s = 0
            p = p & VBA.Mid(x, s, (e - (s - 1)) + 16) '16=len(".jpg"" /></a></p>")
            s = InStr(s + 1, x, "<p><a href=""http://tw.myblog.yahoo.com/jw%21ob4NscCdAxS_yWJbxTvlgfR./photo?pid=")
            e = InStr(e + 1, x, ".jpg"" /></a></p>")
        Loop
    End With
End If
Debug.Print p
End Sub
Sub �ѵ{���X�����X�Ϥ��ñƧ�()
Dim d, x As String, a, s As Long, e As Long, p As String, pe As String, ps As String, pL As Byte
setOX
'Set d = ActiveDocument
d = OX.ClipGet
'x = d.Range
x = d
If InStr(x, "<a name") Then MsgBox "����������,���ˬd!!", vbCritical: Exit Sub
s = InStr(x, """><img src=""http://tw.blog.yahoo.com/photo/photo.php?id=ob4NscCdAxS_yWJbxTvlgfR.&amp;photo=ap_")
If s Then
    ps = """><img src=""http://tw.blog.yahoo.com/photo/photo.php?id=ob4NscCdAxS_yWJbxTvlgfR.&amp;photo=ap_"
ElseIf InStr(x, """><img alt="""""""" src=""http://tw.blog.yahoo.com/photo/photo.php?id=ob4NscCdAxS_yWJbxTvlgfR.&amp;photo=ap_") Then
    ps = """><img alt="""""""" src=""http://tw.blog.yahoo.com/photo/photo.php?id=ob4NscCdAxS_yWJbxTvlgfR.&amp;photo=ap_"
    s = InStr(x, ps)
End If
If s Then
    e = InStr(s, x, ".jpg"" /></a></p>")
    If e Then
        pe = ".jpg"" /></a></p>"
    ElseIf InStr(s, x, ".jpg"" /></a></font>") Then
        pe = ".jpg"" /></a></font>"
    ElseIf InStr(s, x, ".jpg"" border=""0"" /></a><") Then
        pe = ".jpg"" border=""0"" /></a><"
        pL = 11 'pe�����צ���,�G�n�]���Ѽ� 2009/6/15 _
        �O�Hpe=".jpg"" /></a></p>"�@��ǰѷӪ�,�G�n��jpg�P>���S�h�X�h��
    End If
    e = InStr(s, x, pe)
    With d
        Do Until s = 0
            p = p & VBA.Mid(x, s + 2, (e - (s - 1)) + 5 + pL) & "<br>" '16=len(".jpg"" /></a></p>")
            s = InStr(s + 1, x, ps)
            e = InStr(s + 1, x, pe) 'e = InStr(s + 1, x, ".jpg"" /></a></p>")
        Loop
    End With
End If
'Debug.Print p
OX.ClipPut p
On Error Resume Next
AppActivate "avant browser"
End Sub


Sub �Ϥ����}�����ݩ�()
'AppActivate "opera"
AppActivate 2312
DoEvents
'SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{tab}", True
'SendKeys " ", True
'SendKeys "{tab}", True: SendKeys " ", True
'SendKeys "{tab}", True: SendKeys " ", True
'SendKeys "{tab}", True: SendKeys " ", True
'SendKeys "{tab}", True: SendKeys " ", True
'SendKeys "{tab}", True: SendKeys " ", True
'SendKeys "{tab}", True: SendKeys " ", True
'SendKeys "{tab}", True: SendKeys " ", True
'SendKeys "{tab}", True: SendKeys " ", True
SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{left}", True
SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{left}", True
SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{left}", True
SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{left}", True
SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{left}", True
SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{left}", True
SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{left}", True
SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{left}", True
SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{tab}", True: SendKeys "{left}", True
SendKeys "{tab}", True: SendKeys "{tab}", True:: SendKeys " ", True
End Sub

Sub �d�ߩ_���ڪ�������blog() 'Alt+Q
If Selection.Type = wdSelectionIP Then Exit Sub
If ActiveDocument.path <> "" And ActiveDocument.Saved = False Then ActiveDocument.Save
If myaccess Is Nothing Then
    Set myaccess = GetObject("C:\�d�{�@�o�N\���y���\�ϮѺ޲z(C�Ѫ�).mdb")
End If
myaccess.Run "�d�ߩ_���ڪ�������blog_word�ѷ�", Selection
Selection.Copy
myaccess.UserControl = True
Set myaccess = Nothing
End Sub

Sub ����r�r�崡�J�W�s��()
Dim st As String, h As String, d As Document
If ActiveDocument.path <> "" Then Exit Sub
Set d = ActiveDocument
If InStr(d.Range, "http") = 0 Then Exit Sub
With d.Application.Selection
    .HomeKey wdStory, wdMove
    .Find.ClearFormatting
    Do
1   If .Find.Execute("http", , , , , , True, wdFindStop) = False Then Exit Sub
    'If d.Range(.Start - 2, .Start) = "��." Then '���J�W�s����
    If d.Range(.start - 2, .start) = "��." Then
        Do Until st = ".htm"
            .MoveRight wdCharacter, 1, wdExtend
            st = VBA.Right(.text, 4)
        Loop
        h = .text
        .Delete
        h = VBA.StrConv(h, vbNarrow) '������b��
        .Hyperlinks.Add d.Range(.start - 2, .start - 1), h '.Range, h
        st = ""
    ElseIf d.Range(.start - 1, .start) = VBA.Chr(9) Or d.Range(.start - 1, .start) = "�z" Then '�Htab��w��r�����P�_,�\�b��Ʈwvba.Chr(13)�ҷ|�Q�ন���r���G�].
        Do Until st = VBA.Chr(9) Or st = " " Or .Next.font.Size > 8
            .MoveRight wdCharacter, 1, wdExtend
            st = VBA.Right(.text, 1)
        Loop
        .MoveLeft wdCharacter, 1, wdExtend
        h = .text
        .Delete
        h = VBA.StrConv(h, vbNarrow) '������b��
        .Hyperlinks.Add d.Range(.start - 1, .start), h      '.Range, h
        st = ""
        
    Else
        GoTo 1
    End If
    Loop
End With

End Sub

Sub �ˬd�|���o����eMule�M��() '2011/6/19
Dim Dnow As Document, Dold As Document, p As Paragraph, x, l
If Documents(1).path <> "" Or Documents(2).path <> "" Then Exit Sub
Set Dnow = Documents(1) '�̫�@�Ӥ�󬰥ثe�ƻs��emule���M��,�e�@�Ӥ��h��blog�M�橫�ƻs�Ӫ�
Set Dold = Documents(2)
With Dold
    For Each p In .Paragraphs
        x = VBA.Left(p.Range, Len(p.Range) - 1)
        l = InStr(Dnow.Range, x)
        If l = 0 Then
            p.Range.Select
            Exit For
        Else
            Dnow.Characters(l).Paragraphs(1).Range.Delete
        End If
    Next p
End With
Dold.Activate
End Sub


Sub ��X�ѭ��U���s���H�K�j�M��������() '2011/7/14'http://www.webconfs.com/search-engine-spider-simulator.php
Dim i As Long, j As Long, x As String, l As Long
With ActiveDocument
    If .path <> "" Then Exit Sub
    .Range.Paste
    x = .Range
    l = Len(x)
    i = 1
    Do Until i > l
    Select Case VBA.Mid(x, i, 1) & VBA.Mid(x, i + 1, 1)
        Case """>"
            j = i + 2
            Do Until VBA.Mid(x, j, 1) & VBA.Mid(x, j + 1, 1) & VBA.Mid(x, j + 2, 1) = "</a"
                x = VBA.Left(x, j - 1) & VBA.Mid(x, j + 1) ' Replace(x, VBA.Mid(x, j, 1), "", j, 1)
                'VBA.Mid(x, j, 1) = ""
                'j = j + 1
                l = l - 1
            Loop
            i = j
        Case "a>"
            j = i + 2
            Do Until VBA.Mid(x, j, 1) & VBA.Mid(x, j + 1, 1) & VBA.Mid(x, j + 2, 1) = "<a "
                x = VBA.Left(x, j - 1) & VBA.Mid(x, j + 1)
                l = l - 1
                If VBA.Mid(x, j, 1) & VBA.Mid(x, j + 1, 1) & VBA.Mid(x, j + 2, 1) = "" Then Exit Do
            Loop
            i = j
        
    End Select
    i = i + 1
    Loop
    x = Replace(x, "> <", "><")
    x = Replace(x, " &nbsp;", "")
    x = Replace(x, "<br>", "")
    x = Replace(x, "<div>", "")
    x = Replace(x, "</div>", "")
    x = Replace(x, "<p>", "")
    x = Replace(x, "</p>", "")
    .Range = x
    .Range.Cut
End With

End Sub

Sub ���J����() '20161105
Dim p  As Paragraph, s As Long, lnk As String
Const x As String = "\\VBOXSVR\d_drive\�d�{�@�o�N\��Ʈw\���y��Ʈw\����\2487_�y���P�H��\"
s = Selection.End
For Each p In ActiveDocument.Paragraphs
    If p.Range.End > s Then
        If IsNumeric(p.Range) And p.Range.font.Size = 16 Then
            If Dir(x & Format(p.Range, "_000000") & ".tif") = "" Then
                If Dir(x & Format(p.Range, "_000000") & ".jpg") <> "" Then
                    lnk = x & Format(p.Range, "_000000") & ".jpg"
                Else
                    GoTo nt
                End If
            ElseIf Dir(x & Format(p.Range, "_000000") & ".jpg") = "" Then
                If Dir(x & Format(p.Range, "_000000") & ".tif") <> "" Then
                    lnk = x & Format(p.Range, "_000000") & ".tif"
                Else
                    GoTo nt
                End If
            Else
                GoTo nt
            End If
            p.Range.Hyperlinks.Add p.Range, lnk
        End If
    End If
nt:
Next p
MsgBox "done!", vbInformation
End Sub
