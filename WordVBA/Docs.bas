Attribute VB_Name = "Docs"
Option Explicit
Public d�r�� As Document, x As New EventClassModule '�o�~�O�ҿת��إ�"�s"�����O�Ҳ�--��ڤW�O�إ߹復���ѷ�.'��ӽu�W�����DDim�].

Public Sub Register_Event_Handler() '�Ϧ۳]�������O�Ҳզ��Ī��n���{��.���u�ϥ� Application ���� (Application Object) ���ƥ�v
    Set x.app = word.Application '���Y�Ϸs�ت�����PWord.Application����@�W���p
End Sub
Sub �ťժ��s���() '20210209
Dim a As Document, flg As Boolean
If Documents.Count = 0 Then GoTo a:
If ActiveDocument.Characters.Count = 1 Then
    Selection.Paste
ElseIf ActiveDocument.Characters.Count > 1 Then
    For Each a In Documents
        If a.path = "" Or a.Characters.Count = 1 Then
            a.Range.Paste
            a.Activate
            a.ActiveWindow.Activate
            flg = True
            Exit For
        End If
    Next a
    If flg = False Then GoTo a
Else
a: Documents.Add
    Selection.Paste
End If
End Sub


Sub �b����󤤴M�����r��_���t() '���w��:Alt+Ctrl+Down 2015/11/1
Static x As String
With Selection
    If .Type = wdSelectionNormal Then
        x = ��r�B�z.trimStrForSearch(.Text, Selection)
        .Copy
        .Collapse wdCollapseEnd
        .Find.ClearAllFuzzyOptions
        .Find.ClearFormatting
    End If
    If x = "" Then Exit Sub
    If .Find.Execute(x, True, True, , , , True, wdFindContinue) = False Then MsgBox "�S�F!", vbExclamation
End With
End Sub

Sub ���ær��()
'Dim �r�� As Boolean
''On Error Resume Next
'For Each d�r�� In Application.Documents
'    'If d.Name = "�r��7.2.doc" Then �r�� = True
'    If d�r��.Name Like "�r��*" Then �r�� = True: Exit For
'Next d�r��
''If �r�� Then If InStr(ActiveDocument.Name, "�r��7") = 0 Then Documents("�r��7.2.doc").Windows(1).Visible = False
'If �r�� Then If InStr(ActiveDocument.Name, "�r��") = 0 Then d�r��.Windows(1).Visible = False
End Sub
Sub ���J���O���}() '2002/11/15�ѹϮѺ޲zOLE.���J���}���ΦӨ�
Dim h As word.Document, sectionct As Integer, rst As Recordset '2002/11/15
Dim d As Object
On Error GoTo OVER
'�H�U�G���i�ٲ��A���٤��I2003/12/14
'Set d = GetObject("d:\�d�{�@�o�N\���y���\�ϮѺ޲z.mdb") '�ˬd�ϮѺ޲z���L�}��!
'Set d = Nothing
''AppActivate "�ϮѺ޲z"
If Not blog.myaccess.CurrentProject.AllForms("���O_���d��").IsLoaded Then
     MsgBox "[���O_���d��]���S�}��,����@�~!", vbCritical
    End
End If
Set h = ActiveDocument
Set rst = blog.myaccess.Forms("���O_���d��").RecordsetClone
sectionct = h.Sections.Count
Dim a, a1, i As Integer, j As Integer
'h.Application.Visible = True '�ˬd��
a = Array("�U�W", "�ѦW", "�g�W", "��", "��")
ReDim a1(0 To UBound(a)) As String
For i = 1 To sectionct '����v�`���J���}(�C�`�������ƪ��@���O��!)
    For j = 0 To UBound(a)
        '�ѭ��P���O�Ӥ��������U�ѽg�W��...�����
        If Not Left(h.Sections(i).Range.Paragraphs(2).Range.Text, Len(h.Sections(i).Range.Paragraphs(2).Range.Text) - 1) Like "..." Then
            rst.FindFirst "�� = " & Left(h.Sections(i).Range.Paragraphs(1).Range.Text, Len(h.Sections(i).Range.Paragraphs(1).Range.Text) - 1) _
                & "and ���O like '" & "*" & Left(h.Sections(i).Range.Paragraphs(2).Range.Text, _
                    Len(h.Sections(i).Range.Paragraphs(2).Range.Text) - 1) & "*'"
    '            & "and ���O like '" & "*" & Replace(h.Sections(i).Range.Paragraphs(2).Range.Text, Chr(12), "") & "*'"
                'chr(12)�Ȥ����O����(���O���`�Ÿ�!),���|�v�T���,�G�����N���Ŧr�� _
                �]���̫�@�`�S�����`�Ÿ�(Chr(12))�ӬO�q���Ÿ�(Chr(13),�G�Y�HReplace��ƶ����O�B�z _
                ���K�·�,�@�ߥ�Left��Ƥ����̥k�褧�r���Y�i(���ެOChr(12)��Chr(13))
        Else '���O��""�ɪ��B�z
            rst.FindFirst "�� = " & Left(h.Sections(i).Range.Paragraphs(1).Range.Text, Len(h.Sections(i).Range.Paragraphs(1).Range.Text) - 1) _
                & "and ���O = """"" '�b��,����CSng���A�ഫ�O�i�H��! _
                �]�������p���I,�G�b�@������,�����Words����(�|�N�p���I���Ʀr���}�⦨���P��Word),�p�G���S���p���I���ܴN�i�H�F! _
                ��@,�P���O,�O���F�簣�̥k�誺Chr(10)(����Ÿ��B�q���Ÿ�)
        End If
        a1(j) = blog.myaccess.Nz(rst(a(j)), 0) '�����|��Null��!
    Next j
    h.Footnotes.Add h.Sections(i).Range.words(h.Sections(i).Range.words.Count), _
        , "�m" & a1(LBound(a1)) & "�n�A�m" & a1(LBound(a1) + 1) & "�n�A�q" & a1(LBound(a1) + 2) & "�r�A��" _
        & a1(LBound(a1) + 3) & "�A��" & a1(LBound(a1) + 4) & "�C"
Next i
Set rst = Nothing: Set h = Nothing
MsgBox "����" & sectionct & "�����}���J! "
End
OVER:
    MsgBox "�ϮѺ޲z�S�}��,����@�~!", vbCritical
End Sub

Sub ��󤺮e���_�հɥ�() '_�H�r�������() '2004/10/20:���w��(�ֳt��):Ctrl+Alt+Return
Dim s1 As Range, s2 As Range, d1 As Document, d2 As Document, j As Long, k As Long
'Static i As Long, DN As String, MarkTimes As Byte
Dim i As Long, dn As String, MarkTimes As Byte
Select Case Documents.Count '���ˬd����
    Case Is > 2
        MsgBox "�u��@���ֹ������I�бN�����n�����������A�A�ާ@�@���I", vbExclamation: Exit Sub
    Case Is = 0
        Exit Sub
    Case Is = 1
        MsgBox "�ثe�u���@�����I�бN�n�P���չ諸��󥴶}�A�M��A�ާ@�@���I", vbExclamation: Exit Sub
End Select
'�A�ˬd������.
If Windows.Count > 2 Then MsgBox "�бN�h�l����������A�A�ާ@�@���I", vbExclamation: Exit Sub
'�n���m�J���J�I,�H���J�I��m�}�l����B�z!
Set d1 = Documents(1)
Set d2 = Documents(2)
Windows.Arrange 'wdTiled'�ƦC����
If MsgBox("�O�_�n�M���Ÿ�?", vbQuestion + vbOKCancel) = vbOK Then
    d1.Activate: �M���Ҧ��Ÿ�
    d2.Activate: �M���Ҧ��Ÿ�
End If
If dn = "" Then dn = d1.Name
If Not d1.Name Like dn And Not d1.Name Like dn Then i = 0: dn = d1.Name
If Selection.start + 1 = ActiveDocument.Content.End Then Selection.HomeKey wdStory, wdMove '�p�G���J�I����󥽮�...
If i = 0 Then i = Selection.start 'ActiveDocument.Range.Start
If d1.Characters.Count >= d2.Characters.Count Then
    j = d1.Characters.Count
    k = d2.Characters.Count
Else
    j = d2.Characters.Count
    k = d1.Characters.Count
End If
For i = i + 1 To j
    StatusBar = i
    If i > k Then MsgBox "��粒��!": End ': Exit For  '��F���֦r��󪺥��ݮ�
    Set s1 = d1.Characters(i): Set s2 = d2.Characters(i)
'    If Asc(S1) = 2 Or Asc(S1) = 5 Or Asc(S2) = 5 Or Asc(S1) = 2 Then Stop '���}(2)�ε���(5)��
'    If AscW(S1) = 63 Or AscW(S2) = 63 Then Stop '�b�ΰݸ���
    If Not s1 Like s2 Then
        MarkTimes = MarkTimes + 1
        If MarkTimes > 20 Then MsgBox "������q���P�����p�ίʺ|�R�٤�����A�Цۦ�չ�A�A�N���J�I�m��A������l��m�~�����Y�i�I", vbExclamation: Exit Sub
        s1.Select ': D1.Windows(1).ScrollIntoView S1, True 'Selection.Range, True
'        Options.DefaultHighlightColorIndex = wdBrightGreen
        s1.HighlightColorIndex = wdBrightGreen '�Хܬ��å���
        s2.Select ': D1.Windows(1).ScrollIntoView S2, True 'Selection.Range, True
        ActiveDocument.Windows(1).ScrollIntoView ActiveDocument.Characters(i), True
'        j = MsgBox("�Юչ�I" & vbCr & vbCr & _
            "�i " & S1 & " ���� " & S2 & " �j" & vbCr & vbCr & _
            "�n���ӽЫ�����������!", vbExclamation + vbOKCancel)
            s2.HighlightColorIndex = wdBrightGreen '�Хܬ��å���
'        If j = vbOK Then
'            ActiveWindow.Next.Activate
''            ActiveDocument.Windows(1).ScrollIntoView ActiveDocument.Characters(i), True
            ActiveWindow.ScrollIntoView ActiveDocument.Characters(i), True
''            Dim x As Long'�۰ʭp���s����
''            For x = 1 To 50000000
''            Next
'            Exit For
'        Else
'            End
'        End If
    End If
Next i
MsgBox "��粒��!"
End Sub

Sub ��󤺮e���_�հɥ�_unfinished() '_�H�r�������() '2004/10/20:���w��(�ֳt��):Ctrl+Alt+Return
Dim s1 As Range, s2 As Range, Dw1 As Document, Dw2 As Document, j As Long, k As Long
Static i As Long, DwN As String
Select Case Documents.Count '���ˬd����
    Case Is > 2
        MsgBox "�u��@���ֹ������I�бN�����n�����������A�A�ާ@�@���I", vbExclamation: Exit Sub
    Case Is = 0
        Exit Sub
    Case Is = 1
        MsgBox "�ثe�u���@�����I�бN�n�P���չ諸��󥴶}�A�M��A�ާ@�@���I", vbExclamation: Exit Sub
End Select
'�A�ˬd������.
If Windows.Count > 2 Then MsgBox "�бN�h�l����������A�A�ާ@�@���I", vbExclamation: Exit Sub
'�n���m�J���J�I,�H���J�I��m�}�l����B�z!
Set Dw1 = Windows(1)
Set Dw2 = Windows(2)
Windows.Arrange 'wdTiled'�ƦC����
If MsgBox("�O�_�n�M���Ÿ�?", vbQuestion + vbOKCancel) = vbOK Then
    Dw1.Activate: �M���Ҧ��Ÿ�
    Dw2.Activate: �M���Ҧ��Ÿ�
End If
If DwN = "" Then DwN = Dw1.Name
If Not Dw1.Name Like DwN And Not Dw1.Name Like DwN Then i = 0: DwN = Dw1.Name
If Dw1.Selection.start + 1 = ActiveDocument.Content.End Then Selection.HomeKey wdStory, wdMove '�p�G���J�I����󥽮�...
If i = 0 Then i = Dw1.Selection.start 'ActiveDocument.Range.Start
If Dw1.Characters.Count >= Dw2.Characters.Count Then
    j = Dw1.Characters.Count
    k = Dw2.Characters.Count
Else
    j = Dw2.Characters.Count
    k = Dw1.Characters.Count
End If
For i = i + 1 To j
    If i > k Then MsgBox "��粒��!": End ': Exit For  '��F���֦r��󪺥��ݮ�
    Set s1 = Dw1.Characters(i): Set s2 = Dw2.Characters(i)
'    If Asc(S1) = 2 Or Asc(S1) = 5 Or Asc(S2) = 5 Or Asc(S1) = 2 Then Stop '���}(2)�ε���(5)��
'    If AscW(S1) = 63 Or AscW(S2) = 63 Then Stop '�b�ΰݸ���
    If Not s1 Like s2 Then
        s1.Select ': D1.Windows(1).ScrollIntoView S1, True 'Selection.Range, True
'        Options.DefaultHighlightColorIndex = wdBrightGreen
        s1.HighlightColorIndex = wdBrightGreen '�Хܬ��å���
        s2.Select ': D1.Windows(1).ScrollIntoView S2, True 'Selection.Range, True
        ActiveDocument.Windows(1).ScrollIntoView ActiveDocument.Characters(i), True
'        j = MsgBox("�Юչ�I" & vbCr & vbCr & _
            "�i " & S1 & " ���� " & S2 & " �j" & vbCr & vbCr & _
            "�n���ӽЫ�����������!", vbExclamation + vbOKCancel)
            s2.HighlightColorIndex = wdBrightGreen '�Хܬ��å���
'        If j = vbOK Then
'            ActiveWindow.Next.Activate
''            ActiveDocument.Windows(1).ScrollIntoView ActiveDocument.Characters(i), True
            ActiveWindow.ScrollIntoView ActiveDocument.Characters(i), True
''            Dim x As Long'�۰ʭp���s����
''            For x = 1 To 50000000
''            Next
'            Exit For
'        Else
'            End
'        End If
    End If
Next i
MsgBox "��粒��!"
End Sub

Sub �M���Ҧ��Ÿ�() '�ѹϮѺ޲zsymbles�ҲղM�����I�Ÿ���s'�]�A���}�B�Ʀr
'Dim F, a As String, i As Integer
Dim f, i As Integer
f = Array("�P", "�E", "�C", "�v", Chr(-24152), "�G", "�A", "�F", _
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
        'a = Replace(a, F(i), "")
        ActiveDocument.Range.Find.Execute f(i), True, , , , , , wdFindContinue, True, "", wdReplaceAll
    Next
    'ActiveDocument.Content = a
End Sub
Sub �`�}�Ÿ�() '�`���Ÿ��B�����Ÿ��B���}�Ÿ�
Dim i As Integer
For i = 9312 To 9331
    With Selection.Range.Find
        .Replacement.Font.Name = "Arial Unicode MS"
        .Execute ChrW(i), , , , , , , wdFindContinue, , ChrW(i), wdReplaceAll
    End With
Next i
'Dim i As Long, w As Characters, wc As Long
'Set w = ActiveDocument.Range.Characters
'wc = ActiveDocument.Range.Characters.Count
'For i = 1 To wc
'    If AscW(w(i)) > 9311 And AscW(w(i)) < 9332 Then
'        w(i).Font.Name = "Arial Unicode MS"
'    End If
'Next i
End Sub

Sub �K�W�ޤ�() '�N�w�ƻs��ŶKï�����e�K���ޤ�
Dim s As Long, e As Long, r  As Range
If Selection.Type = wdSelectionNormal And Right(Selection, 1) Like Chr(13) Then _
            Selection.MoveLeft wdCharacter, 1, wdExtend '���n�]�t���q�Ÿ�!
If Selection.Style <> "�ޤ�" Then Selection.Style = "�ޤ�" '�p�G���O�ޤ�˦���,�h�令�ޤ�˦�
s = Selection.start '�O�U�_�l��m
Selection.PasteSpecial , , , , wdPasteText '�K�W�¤�r
e = Selection.End '�O�U�K�W�᪺������m
Selection.SetRange s, e
Set r = Selection.Range
With r
    r.Find.Execute Chr(13), , , , , , , wdFindStop, , Chr(11), wdReplaceAll '�N����Ÿ��令��ʤ���Ÿ�
End With
r.Footnotes.Add r '���J���}!
End Sub
Sub �K�W�¤�r() 'shift+insert 2016/7/20
Dim hl, s As Long, r As Range
On Error GoTo ErrHandler
hl = Selection.Range.HighlightColorIndex

s = Selection.start
Set r = Selection.Range
'Selection.PasteSpecial , , , , wdPasteText '�K�W�¤�r
Selection.PasteAndFormat (wdFormatPlainText)
r.SetRange s, Selection.End
If hl <> 9999999 Then r.HighlightColorIndex = hl
Exit Sub
ErrHandler:
Select Case Err.Number
    Case 5342 '���w����������L�k���o�C
        
    Case Else
        MsgBox Err.Number & Err.Description
End Select
End Sub
Sub �@�r�@�q()
With Selection
    .HomeKey wdStory
    Do Until .End = .Document.Range.End - 1
        .MoveRight
        .TypeText Chr(13)
    Loop
End With
End Sub
Sub OCR���B�z()
Dim a, b, i As Byte
a = Array("�r", Chr(13) & "�x", " ", "�w", "�q", "�u", "�t", "�{", "�z", "�|", "�}", "�s", "�x", Chr(9) & Chr(13), Chr(13) & Chr(13), Chr(13) & Chr(9))
b = Array("", Chr(13), "", "", "", "", "", "", "", "", "", "", Chr(9), Chr(13), Chr(13), Chr(13))
With ActiveDocument
    If .path = "" Then
        For i = 0 To UBound(a) - 1
            With .Range.Find
                .Text = a(i)
                With .Replacement
                    .Text = b(i)
                End With
                .Execute , , , , , , , wdFindContinue, , , wdReplaceAll
            End With
        Next
    End If
End With
End Sub

Sub �b�P�ؿ��U�M��ŦX����r���奻() '2009/8/4'Alt+shift+F3
Dim x As String, d As Document, i As Integer
Set d = ActiveDocument
x = Selection
With word.Application.FileSearch
    .NewSearch
    If d.path = "" Then
        .LookIn = "D:\�d�{�@�o�N\�פ��Ƨ�\�դh�פ�\�פ�Z"
    Else
        .LookIn = d.path
    End If
    .SearchSubFolders = True
    '.FileName = "Run"
    .MatchTextExactly = True
    .FileType = msoFileTypeAllFiles
    .TextOrProperty = x
     If .Execute() > 0 Then
        MsgBox "There were " & .FoundFiles.Count & _
            " file(s) found."
        For i = 1 To .FoundFiles.Count
            x = .FoundFiles(i)
            x = Mid(x, InStrRev(x, "\") + 1)
            If x <> d.Name Then
                If MsgBox(x, vbOKCancel) = vbOK Then
                     Documents.Open .FoundFiles(i)
                End If
            End If
        Next i
    Else
        MsgBox "There were no files found."
    End If

End With
End Sub

Sub �M���Ҧ�����()
Dim e
With ActiveDocument
    For Each e In .Comments
        e.Range.Select
        e.Delete
    Next e
End With
End Sub

Sub �}�s����() '�ֳt��:alt+shift+w-�쬰OLE�ܳƧ���()���w��  '2011/6/23''2012/5/20 2003����]�wAlt+w ��]��"�r���ഫ_�رd�ײʶ�"
Dim l As Long, s As Long, YwdInFootnoteEndnotePane
    YwdInFootnoteEndnotePane = Selection.Information(wdInFootnoteEndnotePane) '�O�U�}�s�����e���}���檬�A
    l = Selection.End '.Information(wdActiveEndPageNumber)
    s = Selection.start '�O�U���m
    If CommandBars("web").Visible Then CommandBars("web").Visible = False
    NewWindow
    'ActiveWindow.Document.Range.Characters(l).Select
    If YwdInFootnoteEndnotePane Then '�p�G�b���}���椤
        ActiveWindow.View.SplitSpecial = wdPaneFootnotes '2011/8/13
    End If
    Selection.End = l ' 'Selection.GoTo wdGoToObject, wdGoToAbsolute, l
    Selection.start = s '����m
End Sub

Sub ���޾ɼҦ�����() ' Alt+M 2011/6/26
' ����7 ����
' �������s�� 2011/6/26�A���s�� Oscar Sun
'
Dim s As Long, e As Long, YwdInFootnoteEndnotePane
YwdInFootnoteEndnotePane = Selection.Information(wdInFootnoteEndnotePane) '�O�U�}�s�����e���}���檬�A
If YwdInFootnoteEndnotePane Then '�p�G�b���}���椤
    ActiveWindow.ActivePane.Close '2011/9/24
End If
s = Selection.start
e = Selection.End
If ActiveWindow.DocumentMap Then
    ActiveWindow.DocumentMap = False
ElseIf ActiveWindow.DocumentMap = False Then
    ActiveWindow.DocumentMap = True
End If
Selection.start = s
Selection.End = e
End Sub

Sub �ؿ��r�Χ�() '2011/8/23'�\�Y���D��ΫD���D�j�p���r��,�|bug
Dim e As Long
With ActiveDocument
    e = .Range.End
'    With .ActiveWindow.Selection.Find
''        .Font.Size>12
'        .Execute
'
    Do Until Selection.End = e - 1
        Selection.MoveRight
            If Selection.Next.Font.Size > 12 Then '�ؿ��w�]��12�r��
                Selection.MoveRight
                Do Until Selection.Next.Font.Size = 12
                    Selection.MoveRight , , wdExtend
                Loop
                If MsgBox("�O�_�n�Y�p��10���r?", vbQuestion + vbOKCancel) = vbOK Then Selection.Font.Size = 10
        End If
    Loop
End With
End Sub

Sub �M���å���() '2012/2/3 �ݴ���
Dim h
h = wdYellow
'�L��0,�H�Ŭ�3,.......
'If h = "" Then h = Selection.Range.HighlightColorIndex
Do Until Selection.Range.HighlightColorIndex <> h

    Exit Do 'Selection.Range.HighlightColorIndex = wdAuto
Loop
'MsgBox "����!", vbInformation
End Sub

Sub ������l�����()
'Ctrl+Alt+W
Dim d As Document, dn As String
dn = ActiveDocument.FullName
For Each d In Documents
    If d.FullName <> dn Then d.Close wdDoNotSaveChanges
Next
End Sub

Sub �b����󤤴M�����r��() 'Ctrl+Alt+Down 2020/10/4��� Ctrl+Shift+PageDown
'CheckSavedNoClear
If ActiveDocument.path <> "" Then If ActiveDocument.Saved = False Then ActiveDocument.Save
Dim ins(4) As Long, MnText As String, FnText As String, FdText As String, st As Long, ed As Long
On Error GoTo errHH
With Selection '�ֳt��GAlt+Ctrl+Down
'If Not .Text Like "" Then '�ֳt��GAlt+Ctrl+Down
If .Type = wdSelectionIP Then MsgBox "�п���Q�n�M�䤧��r", vbExclamation: Exit Sub
If .Type = wdSelectionNormal Then ' <> wdNoSelection OR wdSelectionIP Then '�������J�I
'    If InStr(ActiveDocument.Content, .Text) = InStrRev(ActiveDocument.Content, .Text) Then MsgBox "����u�����B!", vbInformation: Exit Sub
    FdText = ��r�B�z.trimStrForSearch(.Text, Selection)
    st = .start: ed = .End
    .Collapse wdCollapseEnd
    MnText = .Document.StoryRanges(wdMainTextStory) '�ܼƤƳB�z����2003/4/8
'    MnText = ActiveDocument.Range '2010/2/5
    ins(1) = InStr(MnText, FdText)
    ins(2) = InStrRev(MnText, FdText)
    If .Document.Footnotes.Count > 0 Then '�����}�~�ˬd2003/4/3
        FnText = .Document.StoryRanges(wdFootnotesStory)
        ins(3) = InStr(FnText, FdText)
        ins(4) = InStrRev(FnText, FdText)
    End If
    If ins(1) = ins(2) And ins(3) = ins(4) Then
        If ins(1) <> 0 And .Information(wdInFootnote) Then
            If MsgBox("���}�u�����B!�@�@�����٦�.." & vbCr & vbCr & _
                "�n�M���?", vbInformation + vbOKCancel, "�M��G�u" & FdText & "�v") = vbCancel Then
                Exit Sub
            Else
'                FdText = .Text
'                .Document.ActiveWindow.ActivePane.Previous.Activate
'                .Document.Select '���k�i�N�J�I�ಾ�쥿��
                With .Document.Range.Find
                    .Text = FdText
                    .Execute
                    .Parent.Select
                End With
            End If
        ElseIf ins(3) <> 0 And Not .Information(wdInFootnote) Then
            If MsgBox("����u�����B!�@�@���}�٦�.." & vbCr & vbCr & _
                "�n�~��M��ܡH", vbInformation + vbOKCancel, "�M��G�u" & _
                    FdText & "�v") = vbCancel Then
                Exit Sub
            Else
'                FdText = .Text
                With .Document.ActiveWindow
                    If .Panes.Count = 1 Then
                        '�}�ҵ��}����
                        If .View.Type = wdNormalView Then _
                           .View.SplitSpecial = wdPaneFootnotes
                    Else
                        .ActivePane.Next.Activate
                    End If
                    With .ActivePane.Selection.Find
                        .Text = FdText
'                        .Forward = True
                        .Wrap = wdFindContinue '�n���o��~�ॿ�T�M��
                        .Execute
                    End With
'                    .ScrollIntoView .ActivePane.Selection, True
'                    .ActivePane.SmallScroll
                End With
            End If
        ElseIf ins(1) = ins(2) And ins(3) <> 0 And ins(1) = 0 And ins(3) = ins(4) Then
            MsgBox "����u�����B!  ����L!", vbInformation, "�M��G�u" & FdText & "�v": Exit Sub
        ElseIf ins(1) = ins(2) And ins(1) <> 0 And ins(3) = 0 And ins(3) = ins(4) Then
            MsgBox "����u�����B!  ���}�L!", vbExclamation, "�M��G�u" & FdText & "�v"
            .start = st
            .End = ed
            Exit Sub
        End If
    Else
'        If ins(1) <> 0 Then
'            ins(1) = wdMainTextStory
'        Else
'            ins(1) = wdFootnotesStory
'        End If
        If ins(3) = ins(4) And .Information(wdInFootnote) = True Then _
            MsgBox "����u�����}���B��!", vbInformation, "�M��G�u" & FdText & "�v": Exit Sub
'        With .Document.StoryRanges(ins(1)).Find
        With .Find
            .ClearFormatting
            .Replacement.ClearFormatting '�o�]�n�M���~��
            .Forward = True
            .Wrap = wdFindAsk
            .MatchCase = True
            .Text = FdText '.Parent.Text
            .Execute
'            .Parent.Select'��Range����o�Φ���k�~����ܿ��
        End With
    End If
End If
End With
Exit Sub
errHH:
Select Case Err.Number
    Case 7 '�O���餣��
        ActiveDocument.ActiveWindow.Selection.Find.Execute Selection.Text
    Case Else
        MsgBox Err.Number & Err.Description
        Resume
End Select
End Sub


Sub ����_�H�����r�@������() 'ALT+SHIFT+B

' �������s�� 2015/9/20�A���s�� ���[�p
    With ActiveDocument.bookmarks
        .Add Range:=Selection.Range, Name:=Replace(Selection.Text, Chr(13), "")
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
End Sub
Sub �p�p��J�k���wcj5_ftzk_3�r�H�W���J����()
If Not ActiveDocument.Name = "cj5-ftzk.txt" Then Exit Sub
Dim d As Document, flg As Boolean, s As Byte, prngTxt As String
Set d = ActiveDocument
Dim p As Paragraph
Const x As String = "ahysy ����"
For Each p In d.Paragraphs
    If InStr(p.Range, x) Then flg = True
    If Not flg Then
        p.Range.Font.Hidden = True
    Else
        prngTxt = p.Range.Text
        s = InStr(prngTxt, " ")
        If VBA.Len(VBA.Mid(prngTxt, s + 1)) < 5 Then
            p.Range.Font.Hidden = True
        End If
    End If
Next p
d.ActiveWindow.View.ShowHiddenText = False
Beep
End Sub
Sub �p�p��J�k���wcj5_ftzk_3�r�H�W���J�R�s��cj5_ftzk_other()
If Not ActiveDocument.Name = "cj5-ftzk.txt" Then Exit Sub
Dim d As Document
�p�p��J�k���wcj5_ftzk_3�r�H�W���J����
Set d = ActiveDocument
Dim p As Paragraph, s As Byte, prngTxt As String, msgResult As Integer
For Each p In d.Paragraphs
    If Not p.Range.Font.Hidden Then
        prngTxt = p.Range.Text
        s = InStr(prngTxt, VBA.Chr(32))
        If VBA.Len(VBA.Mid(prngTxt, s + 1)) > 4 Then
            DoEvents
            p.Range.Select
            ActiveWindow.ScrollIntoView p.Range
            msgResult = MsgBox("�O�_����other�h�H", vbYesNoCancel)
            Select Case msgResult
                Case vbYes
                    �p�p��J�k���wcj5_ftzk�R�s��cj5_ftzk_other
                Case vbNo
                    Debug.Print p.Next.Range.Text
                    Exit Sub
            End Select
        End If
    End If
Next p
End Sub
Sub �p�p��J�k���wcj5_ftzk�R�s��cj5_ftzk_other()
'Dim p As Paragraph'Alt+1 Alt+2 ctrl+q
Dim rng As Range, p As Paragraph, pc As Integer, pSelRng As Range, x As String
If ActiveDocument.Name <> "cj5-ftzk.txt" Then Exit Sub
word.Application.ScreenUpdating = False
If Selection.Type = wdSelectionIP Then
    If Not Selection.Range.TextRetrievalMode.IncludeHiddenText Then
        Selection.Paragraphs(1).Range.Select
    End If
Else
    ActiveWindow.ActivePane.View.ShowAll = Not ActiveWindow.ActivePane.View. _
        ShowAll
End If
If Not Selection.Range.TextRetrievalMode.IncludeHiddenText Then
    x = Selection.Text
    Selection.Delete
    'Selection.Cut
        GoSub subP
'        ActiveWindow.ActivePane.View.ShowAll = Not ActiveWindow.ActivePane.View. _
        ShowAll
        word.Application.ScreenUpdating = True
        Exit Sub
Else
    Set pSelRng = Selection.Range
    For Each p In pSelRng.Paragraphs
        If Not p.Range.Font.Hidden Then 'prepare to delete
            'p.Range.Cut
            x = p.Range.Text
            p.Range.Delete
            GoSub subP
        End If
    Next p
End If
ActiveWindow.ActivePane.View.ShowAll = Not ActiveWindow.ActivePane.View. _
        ShowAll
word.Application.ScreenUpdating = True
Exit Sub
subP:
    With Documents("cj5-ftzk-other.txt").Range
'        If .Paragraphs(.Paragraphs.Count).Range <> Chr(13) Then .InsertParagraphAfter
'        .Paragraphs(.Paragraphs.Count).Range.Paste
'        .Document.ActiveWindow.ScrollIntoView .Paragraphs(.Paragraphs.Count).Range
        .InsertAfter x
        .Document.ActiveWindow.ScrollIntoView .Parent.Range(.End - 1, .End)
    End With
    Return
End Sub
Function �˦����N()
Const styleSrc As String = "�¤�r"
Const styleDest As String = "���g���"
Dim d As Document, p As Paragraph
For Each p In d.Paragraphs
    If p.Style = styleSrc Then p.Style = styleDest
Next p
End Function

Function �˦�add_�K�a�����˦�()
Const styleHprAn As String = "�K�a��"
Const styleShengDiao As String = "�n��"
Dim d As Document, myStyle  As Style, doNotAdd As Boolean
Set d = ActiveDocument
For Each myStyle In d.Styles
    If myStyle = styleHprAn Then doNotAdd = True: Exit For
Next myStyle
If Not doNotAdd Then
    Set myStyle = d.Styles.Add(styleHprAn, wdStyleTypeCharacter)
    With myStyle
        With .Font
            .NameFarEast = "�з���"
            .Color = 12611584
            .Size = 11
            .Spacing = -0.4
        End With
        .Visibility = True
        .Priority = 1
        .UnhideWhenUsed = True
    End With
    doNotAdd = False
End If
For Each myStyle In d.Styles
    If myStyle = styleShengDiao Then doNotAdd = True: Exit For
Next myStyle
If Not doNotAdd Then
    Set myStyle = d.Styles.Add(styleShengDiao, wdStyleTypeCharacter) 'https://docs.microsoft.com/zh-tw/office/vba/api/word.wdstyletype
    With myStyle
        .BaseStyle = d.Styles(styleHprAn)
        With .Font
            .NameFarEast = "�з���"
            .Name = "�з���"
            .Position = 3
        End With
        .Visibility = True
        .Priority = 1
        .UnhideWhenUsed = True
    End With
    doNotAdd = False
End If
End Function

Sub closeDocs�������x�s������ɮ�() 'Alt+w
Dim d As Document
word.Application.ScreenUpdating = False
For Each d In Documents
    If d.path = "" Then d.ActiveWindow.Visible = False
Next d
For Each d In Documents
    If d.path = "" Then d.Close wdDoNotSaveChanges
Next d
word.Application.ScreenUpdating = True
If word.Windows.Count = 0 Then
    word.Documents.Add
    ActiveWindow.Visible = True
End If
End Sub
Sub DocBackgroundFillColor() '������m
    ActiveDocument.ActiveWindow.View.Type = wdPrintView 'https://docs.microsoft.com/en-us/office/vba/api/word.document.background
'    ActiveDocument.Background.Fill.ForeColor.RGB = RGB(192, 192, 192) 'RGB(146, 208, 80)
'    ActiveDocument.Background.Fill.Visible = True
'    ActiveDocument.Background.Fill.Solid
    ActiveDocument.Background.Fill.Visible = True
    ActiveDocument.Background.Fill.ForeColor.RGB = RGB(0, 102, 102)
    ActiveDocument.Background.Fill.BackColor.RGB = RGB(0, 102, 102)
    ActiveDocument.Background.Fill.Solid
End Sub

Sub ����e�ŤG��() 'Alt+n
With Selection.ParagraphFormat
    .Style = "����"
    .CharacterUnitFirstLineIndent = 2
End With
End Sub

Sub mark��������r()
Dim searchedTerm, e, ur As UndoRecord, d As Document, clipBTxt As String, flgPaste As Boolean, xd As String
'Set ur = SystemSetup.stopUndo("mark��������r")
SystemSetup.stopUndo ur, "mark��������r"
Set d = ActiveDocument
clipBTxt = VBA.Trim(SystemSetup.GetClipboardText)
searchedTerm = Array("��", "��", "��", "�P��", "���g", "�t��", "ô��", "����", "����", "ô��", "����", "�Ǩ�", "����", "�Ԩ�", "����", "�娥", "���[", "�L�S", ChrW(26080) & "�S", "�ѩS", "����", "�Q�s" _
    , "�v�O", "�E��", "���G", "�W�E", "�W��", "�E�G", "�b") ', "", "", "", "")

'If Selection.Type = wdSelectionIP Then
    If InStr(d.Range, clipBTxt) = 0 Then
        For Each e In searchedTerm
            If InStr(clipBTxt, e) > 0 Then
                flgPaste = True
                Exit For
            End If
        Next e
        If Not flgPaste Then
            'ChrW() & ChrW() &'ChrW() & ChrW() &
            Dim gua
            gua = Array(ChrW(-10119), ChrW(-8742), ChrW(-30233), ChrW(-10164), ChrW(-8698), ChrW(-31827), ChrW(-10132), ChrW(-8313), ChrW(20810), ChrW(-10167), ChrW(-8698), ChrW(-26587), ChrW(21093), ChrW(14615), ChrW(20089), ChrW(26080), "�k", ChrW(26083), "��" _
                        , "��", "�P", ChrW(20089), "��", "��", "�p�b", "�i", "�{", "�[", "�j�L", "�[", "��", "�_", "����", "�N", "��", "��", "�X", "�P�H", "�j��", "��", "�_", "��", "��", "�^", "��", "��", "�L�k", "�j�b", "�v", "��", "�H", "��", "�[", "�w", "��", "�l", "�q", "�_", "��", "����", "�Q", "�j��", "�[", "�l", "��", "�k�f", "�p�L", "��", "���i", "��", "��", "��", "��", "�J��", "����", "�a�H", "��", "�x", "��", "�S", "�I", "�", "��", "��", "��", "�A", "�`")
            For Each e In gua
                If InStr(clipBTxt, e) > 0 Then
                    flgPaste = True
                    Exit For
                End If
            Next e
        End If
        
        If flgPaste Then
            Selection.EndKey wdStory
            Selection.InsertParagraphAfter
            Selection.InsertParagraphAfter
            Selection.Collapse wdCollapseEnd
            �K�W�¤�r
            Selection.InsertParagraphAfter: Selection.InsertParagraphAfter: Selection.InsertParagraphAfter
            Selection.Collapse wdCollapseEnd
            ActiveWindow.ScrollIntoView Selection
        End If
    Else
        Dim rng As Range
        Set rng = d.Range
        rng.Find.Execute VBA.Left(clipBTxt, 255), , , , , , , wdFindContinue
    End If
'End If
If flgPaste Then
    word.Application.ScreenUpdating = False
    If d.path <> "" And Not d.Saved Then d.Save
    xd = d.Range.Text
    For Each e In searchedTerm
        If InStr(xd, e) > 0 Then
            With d.Range.Find
                .ClearFormatting
                .Text = e
                With .Replacement
                    .Text = e
                    .Font.ColorIndex = wdRed
                    .Highlight = True
                End With
                .Execute , , , , , , True, wdFindContinue, , , wdReplaceAll
            End With
        End If
    Next e
    d.Range.Find.ClearFormatting

    GoSub refres
    SystemSetup.playSound 1.921
Else
    GoSub refres
    SystemSetup.playSound 1.469
End If
SystemSetup.contiUndo ur
Set ur = Nothing
Exit Sub
refres:
    word.Application.ScreenUpdating = True
    If Not rng Is Nothing Then rng.Select
    word.Application.ScreenRefresh
    ActiveWindow.ScrollIntoView Selection, True

Return
End Sub


Sub ����W�s����}()
Dim hplnk As Hyperlink, x As String, d As Document
For Each hplnk In ActiveDocument.Hyperlinks
    x = x & hplnk.Address & Chr(13)
Next hplnk
Set d = Documents.Add
d.Range.Text = x
d.Range.Cut
d.Close wdDoNotSaveChanges
End Sub


Sub ���J�W�s��_��󤤪���m_���D() 'Alt+P ��O�u�޸֡v�˦�'2021/11/27
Dim d As Document, title As String, p As Paragraph, pTxt As String, subAddrs As String, flg As Boolean
Set d = ActiveDocument
title = Selection.Text
title = ��r�B�z.trimStrForSearch(title, Selection)
For Each p In d.Paragraphs
    If Left(p.Style.NameLocal, 2) = "���D" Then
        pTxt = p.Range.Text
        pTxt = Left(pTxt, Len(pTxt) - 1)
        If StrComp(pTxt, title) = 0 Then
            subAddrs = title
            flg = True
            Exit For
        ElseIf InStr(pTxt, title) > 0 Then
            subAddrs = "_" & Mid(pTxt, 1, InStrRev(pTxt, " ") - 1)
            subAddrs = Replace(subAddrs, " ", "_")
            flg = True
            Exit For
        End If
    End If
Next p
'
'    'Selection.MoveLeft Unit:=wdCharacter, Count:=4, Extend:=wdExtend
'    'ChangeFileOpenDirectory d.path & "\"  ''userProfilePath & "Dropbox\"
'
If flg Then
    d.Hyperlinks.Add Anchor:=Selection.Range, Address:="", _
        SubAddress:=subAddrs, ScreenTip:="", TextToDisplay:=title
Else
    MsgBox "�Ф�ʴ��J�I", vbExclamation
End If
End Sub

Sub ������Ǯѹq�l�ƭp��_�u�O�d����`��_�B�`��e��[�A��_�K��j�y�Ŧ۰ʼ��I()
Dim ur As UndoRecord
SystemSetup.stopUndo ur, "������Ǯѹq�l�ƭp��_����e��[�A��_�K��j�y�Ŧ۰ʼ��I"
������Ǯѹq�l�ƭp��.�u�O�d����`��_�B�`��e��[�A��
�K��j�y�Ŧ۰ʼ��I
SystemSetup.contiUndo ur
End Sub

Sub �K��j�y�Ŧ۰ʼ��I()
Dim d As Document
Set d = ActiveDocument
If d.path <> "" Then Exit Sub
d.Range.Cut
On Error GoTo app
AppActivate "�j�y��"
DoEvents
SendKeys "{TAB 16}", True
SendKeys "^v"
DoEvents
SendKeys "+{TAB 2}~", True
If d.path = "" Then d.Close wdDoNotSaveChanges
Exit Sub
app:
Select Case Err.Number
    Case 5
        Shell (Network.getDefaultBrowserFullname + " https://old.gj.cool/gjcool/index")
        AppActivate Network.getDefaultBrowserNameAppActivate '"�j�y��"
        DoEvents
        SystemSetup.Wait 2.5
        'SendKeys "{TAB 16}", True
        Resume Next
    Case Else
        MsgBox Err.Number & Err.Description
End Select
End Sub

Sub updateURL() '��s�W�s�����}
Dim site As String
Dim lnk As New Links
site = InputBox("what site to update?", , "�~�y�j����=1;��y���=2;��Ǥj�v=3")
If site = "" Then Exit Sub
Select Case site
    Case 1 '"�~�y�j����"
        lnk.updateURL�~�y�j���� ActiveDocument
    Case 2 '"��y���"
        lnk.updateURL��y��� ActiveDocument
    Case 3 '"��Ǥj�v"
        lnk.updateURL��Ǥj�v ActiveDocument
        
End Select
End Sub



