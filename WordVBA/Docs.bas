Attribute VB_Name = "Docs"
Option Explicit
Public d�r�� As Document, x As New EventClassModule   '�o�~�O�ҿת��إ�"�s"�����O�Ҳ�--��ڤW�O�إ߹復���ѷ�.'��ӽu�W�����DDim�].
'https://learn.microsoft.com/en-us/office/vba/word/concepts/objects-properties-methods/using-events-with-the-application-object-word
Public Sub Register_Event_Handler() '�Ϧ۳]�������O�Ҳզ��Ī��n���{��.���u�ϥ� Application ���� (Application Object) ���ƥ�v
    Set x.App = word.Application '���Y�Ϸs�ت�����PWord.Application����@�W���p
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
        If Not left(h.Sections(i).Range.Paragraphs(2).Range.Text, Len(h.Sections(i).Range.Paragraphs(2).Range.Text) - 1) Like "..." Then
            rst.FindFirst "�� = " & left(h.Sections(i).Range.Paragraphs(1).Range.Text, Len(h.Sections(i).Range.Paragraphs(1).Range.Text) - 1) _
                & "and ���O like '" & "*" & left(h.Sections(i).Range.Paragraphs(2).Range.Text, _
                    Len(h.Sections(i).Range.Paragraphs(2).Range.Text) - 1) & "*'"
    '            & "and ���O like '" & "*" & Replace(h.Sections(i).Range.Paragraphs(2).Range.Text, Chr(12), "") & "*'"
                'chr(12)�Ȥ����O����(���O���`�Ÿ�!),���|�v�T���,�G�����N���Ŧr�� _
                �]���̫�@�`�S�����`�Ÿ�(Chr(12))�ӬO�q���Ÿ�(Chr(13),�G�Y�HReplace��ƶ����O�B�z _
                ���K�·�,�@�ߥ�Left��Ƥ����̥k�褧�r���Y�i(���ެOChr(12)��Chr(13))
        Else '���O��""�ɪ��B�z
            rst.FindFirst "�� = " & left(h.Sections(i).Range.Paragraphs(1).Range.Text, Len(h.Sections(i).Range.Paragraphs(1).Range.Text) - 1) _
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
Dim f, i As Integer, ur As UndoRecord
SystemSetup.stopUndo ur, "�M���Ҧ��Ÿ�"
f = Array("�P", "�E", "�C", "�v", chr(-24152), "�G", "�A", "�F", _
    "�B", "�u", ".", chr(34), ":", ",", ";", _
    "�K�K", "...", "�D", "�i", "�j", " ", "�m", "�n", "�q", "�r", "�H" _
    , "�I", "��", "��", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0" _
    , "�y", "�z", chr(13), ChrW(9312), ChrW(9313), ChrW(9314), ChrW(9315), ChrW(9316) _
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
SystemSetup.contiUndo ur
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
    If Selection.Type = wdSelectionNormal And right(Selection, 1) Like chr(13) Then _
                Selection.MoveLeft wdCharacter, 1, wdExtend '���n�]�t���q�Ÿ�!
    If Selection.Style <> "�ޤ�" Then Selection.Style = "�ޤ�" '�p�G���O�ޤ�˦���,�h�令�ޤ�˦�
    s = Selection.start '�O�U�_�l��m
    Selection.PasteSpecial , , , , wdPasteText '�K�W�¤�r
    e = Selection.End '�O�U�K�W�᪺������m
    Selection.SetRange s, e
    Set r = Selection.Range
    With r
        r.Find.Execute chr(13), , , , , , , wdFindStop, , chr(11), wdReplaceAll '�N����Ÿ��令��ʤ���Ÿ�
    End With
    r.Footnotes.Add r '���J���}!
End Sub
Sub �K�W�¤�r() 'shift+insert 2016/7/20
    Dim hl, s As Long, r As Range
    On Error GoTo ErrHandler
    hl = Selection.Range.HighlightColorIndex
    
    s = Selection.start
    Set r = Selection.Range
'    '�p�G������h�M��
    If Selection.Flags <> 24 And Selection.Flags <> 25 Or Selection.Flags = 9 Then
        If s < Selection.End Then Selection.Text = vbNullString
    End If
    'Selection.PasteSpecial , , , , wdPasteText '�K�W�¤�r
    Selection.PasteAndFormat (wdFormatPlainText)
    r.SetRange s, Selection.End
    If hl <> 9999999 Then r.HighlightColorIndex = hl '9999999 is multi-color �h�����G��m�h�L���9999999�]7���9�^����
    Exit Sub
ErrHandler:
    Select Case Err.Number
        Case 5342 '���w����������L�k���o�C
            
        Case Else
            MsgBox Err.Number & Err.Description
    End Select
End Sub
Sub �K�W²�Ʀr�奻�ॿ()
    Dim rng As Range, ur As UndoRecord
    SystemSetup.stopUndo ur, "�K�W²�Ʀr�奻�ॿ"
    Set rng = Selection.Range
    rng.PasteAndFormat (wdFormatPlainText)
    ���I�Ÿ��m�� rng: �M���b�ΪŮ� rng: �b�άA������� rng
    If MsgBox("�O�_²�ॿ�H", vbOKCancel) = vbOK Then
        'rng.Select
        rng.TCSCConverter wdTCSCConverterDirectionAuto
        'Selection.Range.TCSCConverter wdTCSCConverterDirectionAuto
    End If
    SystemSetup.contiUndo ur
    End Sub
Sub ²�Ʀr�奻�ॿ()
    Dim rng As Range, ur As UndoRecord
    SystemSetup.stopUndo ur, "²�Ʀr�奻�ॿ"
    Set rng = Selection.Range
    ���I�Ÿ��m�� rng: �M���b�ΪŮ� rng: �b�άA������� rng
    rng.TCSCConverter wdTCSCConverterDirectionAuto
    SystemSetup.contiUndo ur
    SystemSetup.playSound 1
End Sub
Function ���I�Ÿ��m��(Optional rng As Range)
    Dim ay, i As Integer
    ay = Array(ChrW(8220), "�u", ChrW(8221), "�v", ChrW(-431), "�B", ChrW(-432), "�A" _
        , ChrW(58), "�G", ChrW(8216), "�y", ChrW(8217), "�z", _
        ChrW(-428), "�F", "�P", "�E", ",", "�A", ";", "�F" _
        , "?", "�H", ":", "�G", "�R", "�G")
    For i = 0 To UBound(ay)
        rng.Find.Execute ay(i), , , , , , , wdFindContinue, , ay(i + 1), wdReplaceAll
        i = i + 1
    Next i
End Function
Function �M���b�ΪŮ�(Optional rng As Range)
rng.Find.Execute " ", , , , , , , wdFindContinue, , "", wdReplaceAll
End Function
Function �b�άA�������(Optional rng As Range)
rng.Find.Execute "(", , , , , , , wdFindContinue, , "�]", wdReplaceAll
rng.Find.Execute ")", , , , , , , wdFindContinue, , "�^", wdReplaceAll
End Function


Sub �@�r�@�q()
With Selection
    .HomeKey wdStory
    Do Until .End = .Document.Range.End - 1
        .MoveRight
        .TypeText chr(13)
    Loop
End With
End Sub
Sub OCR���B�z()
Dim a, b, i As Byte
a = Array("�r", chr(13) & "�x", " ", "�w", "�q", "�u", "�t", "�{", "�z", "�|", "�}", "�s", "�x", chr(9) & chr(13), chr(13) & chr(13), chr(13) & chr(9))
b = Array("", chr(13), "", "", "", "", "", "", "", "", "", "", chr(9), chr(13), chr(13), chr(13))
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

Sub �M���h�l�����n�����q()
Dim p As Paragraph, rng As Range, ur As UndoRecord
SystemSetup.stopUndo ur, "�M���h�l�����n�����q"
Set rng = ActiveDocument.Range
For Each p In ActiveDocument.Paragraphs
    If p.Range.Characters.Count > 2 Then
        If Not p.Range.Characters(p.Range.Characters.Count - 1) Like "[�n�v�z�C�]" & ChrW(-197) & "0-9a-zA-Z]" Then
            If p.Range.End < ActiveDocument.Range.End - 1 Then
'                p.Range.Characters(p.Range.Characters.Count).Select
                p.Range.Characters(p.Range.Characters.Count).Delete
            End If
        End If
    End If
Next
SystemSetup.contiUndo ur
SystemSetup.playSound 2
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
        .Add Range:=Selection.Range, Name:=Replace(Selection.Text, chr(13), "")
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
        s = InStr(prngTxt, VBA.chr(32))
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
            .position = 3
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
Dim strAutoCorrection, endDocOld As Long, rng As Range
Dim punc As New punctuation
strAutoCorrection = Array("�A�r", "�r�A", "�q�B", "�q", "�q�C", "�q", "�C�r", "�r", "�q�G", "�q", "�G�r", "�r", "�q�A", "�q", "�B�r", "�r")
'If Documents.Count = 0 Then Documents.Add
If Documents.Count = 0 Then Docs.�ťժ��s���
If ClipBoardOp.Is_ClipboardContainCtext_Note_InlinecommentColor Then
    ������Ǯѹq�l�ƭp��.�u�O�d����`��_�B�`��e��[�A��
    Set d = ActiveDocument
    On Error GoTo eH:
    DoEvents
    d.Range.Cut
    d.Close wdDoNotSaveChanges
End If
Set d = ActiveDocument
Rem �]���e���|���u������Ǯѹq�l�ƭp��.�u�O�d����`��_�B�`��e��[�A���v�|�Ψ�UndoRecord����A�B�|��������A�G�H�U����Ҽg��m�N������A�_�h�|�H����������H���L�ġC20230201�ѥf�~�Q�@
SystemSetup.stopUndo ur, "mark��������r"
Set rng = d.Range
endDocOld = d.Range.End
'    If InStr(d.Range.text, Chr(13) & Chr(13) & Chr(13) & Chr(13)) > 0 Then
''        d.Range.Text = Replace(d.Range.Text, Chr(13) & Chr(13) & Chr(13) & Chr(13), Chr(13) & Chr(13) & Chr(13))
'    '�O�d�榡�A�G�ΥH�U�A���ΥH�W
'        With d.Range.Find
'            If InStr(.Parent.text, Chr(13) & Chr(13) & Chr(13) & Chr(13)) > 1 Then
'                .ClearFormatting
'                '.Execute Chr(13) & Chr(13) & Chr(13) & Chr(13), , , , , , True, wdFindContinue, , Chr(13) & Chr(13) & Chr(13), wdReplaceAll
                Rem ����|�y��Word crash
'                .Execute "^p^p^p^p", , , , , , True, wdFindContinue, , "^p^p^p", wdReplaceAll
'            End If
'            .ClearFormatting
'        End With
'    End If

Rem �N�ŶKï�����[�J���奻�W�d��
clipBTxt = Replace(Replace(Replace(Replace(VBA.Trim(SystemSetup.GetClipboardText), chr(13) + chr(10) + "�ťy�l" + chr(13) + chr(10), chr(13) + chr(10) + chr(13) + chr(10)), chr(9), ""), "�D�@", ""), "�@�D", "")
clipBTxt = ��r�B�z.trimStrForSearch_PlainText(clipBTxt)
clipBTxt = �~�y�q�l���m��Ʈw.CleanTextPicPageMark(clipBTxt)
For e = 0 To UBound(strAutoCorrection)
    clipBTxt = Replace(clipBTxt, strAutoCorrection(e), strAutoCorrection(e + 1))
    e = e + 1
Next e
searchedTerm = Array("��", "��", "��", "�P��", "���g", "�t��", "ô��", "����", "����", "ô��", "����", "�Ǩ�", _
    "����", "�Ԩ�", "����", "�娥", "���[", "�L�S", ChrW(26080) & "�S", "�ѩS", "����", "�Q�s", "�v�O", "�E��", _
    "���G", "�W�E", "�W��", "�E�G", "�E�T", "�b", "�[", "�q���r", "�q�[�r", "���B�[", "�q���B�[�r", "�H��", "ν��", _
    ChrW(26080) & ChrW(-10171) & ChrW(-8522))  ', "", "", "", "" )

'If Selection.Type = wdSelectionIP Then
    Rem �P�_�O�_�w�t���Ӥ奻
    '�p�G���t��奻
    If Not Docs.isDocumentContainClipboardText_IgnorePunctuation(d, clipBTxt) Then
        Rem �奻�ۦ��פ��
        Dim similarCompare As New Collection
        Set similarCompare = Docs.similarTextCheckInSpecificDocument(d, clipBTxt)
        If similarCompare.item(1) Then
            If MsgBox("�奻�ۦ��׬� " & vbCr & similarCompare.item(3) _
                & vbCr & VBA.vbTab & "�ۦ��q�����G" & VBA.IIf(VBA.Len(similarCompare.item(2)) > 255, VBA.left(similarCompare.item(2), 255) & "�K�K", similarCompare.item(2)) _
                & vbCr & vbCr & "���U�u�T�w�v�N�|��������q���A�Цۦ��ˬd�O�_���n�A�K�J" & vbCr & vbCr & "���U�u�����v�h�����ˬd�A�N�~�����", vbExclamation + vbOKCancel, "�n�K�J���奻�b���󤤦��������q��!!!") = vbOK Then
                Set rng = d.Range
                If rng.Find.Execute(VBA.left(similarCompare.item(2), 255), , , , , , , wdFindContinue) Then
                    If VBA.Len(similarCompare.item(2)) > 255 Then
                        rng.Paragraphs(1).Range.Select                  '�Хܬۦ��奻
                        d.ActiveWindow.ScrollIntoView Selection.Characters(1), True
                    Else
                        rng.Select
                    End If
                End If
                Set similarCompare = Nothing
                GoTo exitSub
            End If
        End If
        Set similarCompare = Nothing
        Rem end �奻�ۦ��פ��
        
        Rem �t������������r�~�K�W
        For Each e In searchedTerm
            If InStr(clipBTxt, e) > 0 Then
                flgPaste = True '�p�G�t������������r
                Exit For
            End If
        Next e
        If Not flgPaste Then
            'ChrW() & ChrW() &'ChrW() & ChrW() &
            Dim gua
            gua = Array(ChrW(-10119), ChrW(-8742), ChrW(-30233), ChrW(-10164), ChrW(-8698), ChrW(-31827), ChrW(-10132), ChrW(-8313), ChrW(20810), ChrW(-10167), ChrW(-8698), ChrW(-26587), ChrW(21093), ChrW(14615), ChrW(20089), ChrW(26080), "�k", ChrW(26083), "��" _
                        , "��", "�P", ChrW(20089), "��", "��", "�p�b", "�i", "�{", "�[", "�j�L", "�[", "��", "�_", "����", "�N", "��", "��", "�X", "�P�H", "�j��", "��", "�_", "��", "��", "�^", "��", "��", "�L�k", "�j�b", "�v", "��", "�H", "��", "�[", "�w", "��", "�l", "�q", "�_", "��", "����", "�Q", "�j��", "�[", "�l", "��", "�k�f", "�p�L", "��", "���i", "��", "��", "��", "��", "�J��", "����", "�a�H", "��", "�x", "��", "�S", "�I", "�", "��", "��", "��", "�A", "�`", "�ӷ�", "����", "���", "�H", "ν")
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
'            Selection.InsertParagraphAfter
            Selection.Collapse wdCollapseEnd
            Selection.TypeText clipBTxt
            'SystemSetup.SetClipboard clipBTxt
            On Error GoTo eH
            'Docs.�K�W�¤�r
            
            Selection.InsertParagraphAfter: Selection.InsertParagraphAfter: Selection.InsertParagraphAfter
            Selection.Collapse wdCollapseEnd
            ActiveWindow.ScrollIntoView Selection
        Else
            Dim noneYijingKeyword As Boolean
            noneYijingKeyword = True
        End If
    Else '�p�G��󤤤w���奻�A�h��ܨ�Ҧb�B
        Dim sx As String
        If InStr(d.Content, clipBTxt) Then
            'rng.Find.Execute VBA.Left(clipBTxt, 255), , , , , , , wdFindContinue
            Dim ps As Integer
            ps = InStr(clipBTxt, chr(13)) '�p�����ӭn�K�J���奻�����q���A�h����q���e����F�Y�S���A�h����M�䪺�̤j��255�Ӧr���������e�@�j�M
            sx = VBA.IIf(ps > 0, VBA.left(VBA.Mid(clipBTxt, 1, VBA.IIf(ps > 0, ps, 2) - 1), 255), VBA.left(clipBTxt, 255))
        Else '���I�Ÿ��B�z�G�T�w�奻�w���u�O���I�Ÿ����P��
            punc.clearPunctuations clipBTxt
            punc.restoreOriginalTextPunctuations d.Range.Text, clipBTxt
            Set punc = Nothing
            sx = ��r�B�z.trimStrForSearch_PlainText(clipBTxt)
            SystemSetup.SetClipboard sx
            sx = VBA.left(sx, 255)
        End If
        rng.Find.Execute sx, , , , , , , wdFindContinue
        endDocOld = rng.End

    End If
'End If
If flgPaste Then
    Rem ��������r
    word.Application.ScreenUpdating = False
    If d.path <> "" And Not d.Saved Then d.Save
'    xd = d.Range.text
    Dim rngMark As Range
    For Each e In searchedTerm
        Set rngMark = d.Range(endDocOld, d.Range.End)
        xd = rngMark.Text
        If InStr(xd, e) > 0 Then
'            With d.Range.Find
            With rngMark.Find
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
    GoSub refres
    SystemSetup.playSound 1.921
    Rem https://en.wikipedia.org/wiki/CJK_Unified_Ideographs
    Rem �ݮe�r
    'https://en.wikipedia.org/wiki/CJK_Compatibility_Ideographs
'    Docs.ChangeFontOfSurrogatePairs_Range "HanaMinA", d.Range(selection.Paragraphs(1).Range.start, d.Range.End), CJK_Compatibility_Ideographs
    'https://en.wikipedia.org/wiki/CJK_Compatibility_Ideographs_Supplement
    Dim rngChangeFontName As Range
    Set rngChangeFontName = d.Range(Selection.Paragraphs(1).Range.start, d.Range.End)
    Docs.ChangeFontOfSurrogatePairs_Range "HanaMinA", rngChangeFontName, CJK_Compatibility_Ideographs_Supplement
    Rem �X�R�r��
    'HanaMinB�٤��䴩G�H�᪺
    Docs.ChangeFontOfSurrogatePairs_Range "HanaMinB", rngChangeFontName, CJK_Unified_Ideographs_Extension_E
    Docs.ChangeFontOfSurrogatePairs_Range "HanaMinB", rngChangeFontName, CJK_Unified_Ideographs_Extension_F
Else '��󤺤w�����e��
    GoSub refres
    SystemSetup.playSound 1.294
    If noneYijingKeyword Then MsgBox "�n�K�W���奻�ä��t����������r�@�I" + vbCr + vbCr + "�ЦA�ˬd�ҽƻs��ŶKï�����e�O�_���T�C�P���P���@�n�L��������"
End If

exitSub:
SystemSetup.contiUndo ur
Set ur = Nothing
'word.Application.ScreenUpdating = True
'word.Application.ScreenRefresh
Exit Sub


refres:
    word.Application.ScreenUpdating = True
    If flgPaste Then
        Rem ���ٲ��A�K�o�C���K�J�����@���A��r�B�z.�ѦW���g�W���Ъ`�A���������ɭn�����ɮ׫e�A��
        '��r�B�z.�ѦW���g�W���Ъ`
        'If flgPaste Then'���յLê��i�R����
        '��ܷs�K�W���奻����
        rng.SetRange endDocOld, endDocOld
        Do Until rng.Font.ColorIndex = wdRed Or rng.End = d.Range.End - 1
            rng.move
        Loop
        e = rng.End
        rng.SetRange endDocOld, e
    Else
        rng.SetRange endDocOld - Len(sx), endDocOld
    End If
    rng.Select
'    word.Application.ScreenRefresh
    ActiveWindow.ScrollIntoView Selection.Characters(1) ', False
Return

eH:
Select Case Err.Number
    Case 5825 '����w�Q�R���C
        GoTo exitSub
    Case Else
        MsgBox Err.Number + Err.Description
End Select
End Sub



Rem �P�_�ŶKï�̪��¤�r(�Ϋ��w����r)���e�O�_�b��󤤤w�s�b
Function isDocumentContainClipboardText_IgnorePunctuation(d As Document, Optional chkClipboardText As String) As Boolean
    Dim xd As String
    If chkClipboardText = "" Then chkClipboardText = SystemSetup.GetClipboardText
    Rem �ŶKï�̪�����Ÿ��ȬOchr(13)&chr(10)�ӦbWord��󤤬O�u�� chr(13)
    chkClipboardText = VBA.Replace(chkClipboardText, chr(13) & chr(10), chr(13))
    xd = d.Range.Text
    If VBA.InStr(xd, chkClipboardText) > 0 Then
        isDocumentContainClipboardText_IgnorePunctuation = True
    Else '�������I�Ÿ������
        Dim punc As New punctuation
        If punc.inStrIgnorePunctuation(xd, chkClipboardText) Then
            isDocumentContainClipboardText_IgnorePunctuation = True
        Else
            If isDocumentContainClipboardText_IgnorePunctuation Then isDocumentContainClipboardText_IgnorePunctuation = False
        End If
        Set punc = Nothing
    End If
End Function

Function similarTextCheckInSpecificDocument(d As Document, Text As String) As Collection 'item1 as Boolean(�奻�O�_�ۦ�),item2 as string(��쪺�ۦ��奻�q��),item3 as String from Dictionary SimilarityResult(�ۦ��צW&�ۦ���)
Rem �奻�ۦ��פ��
Dim similarText As New similarText, dClearPunctuation As String, textClearPunctuation As String, dCleanParagraphs() As String, punc As New punctuation, e, Similarity As Boolean, result As New Collection
dClearPunctuation = d.Content.Text
textClearPunctuation = Text
'�M�����I�Ÿ�
punc.clearPunctuations textClearPunctuation: punc.clearPunctuations dClearPunctuation
dCleanParagraphs = VBA.Split(dClearPunctuation, chr(13))
For Each e In dCleanParagraphs
    If e <> "" Then
'        If e = "��" Then Stop
        If similarText.Similarity(e, textClearPunctuation) Then
            Similarity = True: Exit For
        ElseIf similarText.SimilarityPercent(e, textClearPunctuation) > 80 Then
            Similarity = True: Exit For
        End If
    End If
Next e
'If Similarity = True Then Stop 'for test
Rem index   Required. An expression that specifies the position of a member of the collection. If a numeric expression, index must be a number from 1 to the value of the collection's Count property. If a string expression, index must correspond to the key argument specified when the member referred to was added to the collection.
result.Add Similarity 'item1:�奻�O�_�ۦ�'https://learn.microsoft.com/en-us/office/vba/Language/Reference/User-Interface-Help/item-method-visual-basic-for-applications
dClearPunctuation = e
punc.restoreOriginalTextPunctuations d.Content.Text, dClearPunctuation
result.Add dClearPunctuation 'item2:��쪺�ۦ��奻�q��
result.Add similarText.SimilarityResultsString 'item3:�ۦ��צW&�ۦ���
Set similarText = Nothing
Set similarTextCheckInSpecificDocument = result
Rem end �奻�ۦ��פ��
End Function
Sub �����_���ŧ()
Dim d1 As Document, d2 As Document, p As Paragraph, x As String, i As Byte, rng As Range, pc As Long, d1RngTxt, px As String, rng2 As Range
Static pi As Long
Set d1 = Documents(1) '�ӷ�
d1RngTxt = d1.Range.Text
Set d2 = Documents(2) '��ŧ�Τޥ�(�����N�C�A���y�l����U�q��r�^
pc = d2.Paragraphs.Count
If pi = 0 Then pi = 1
For pi = pi To pc
    Set p = d2.Paragraphs(pi)
    If p.Range.Font.NameFarEast <> "�з���" And p.Range.HighlightColorIndex = 0 Then
        px = p.Range
        x = VBA.Trim(left(px, Len(px) - 1)) '�h�����q�Ÿ�
        If Len(x) > 2 Then
            x = VBA.left(x, Len(x) - 1) '�h���ݫ���I�C�A��
            If VBA.InStr(d1RngTxt, x) Then
                i = i + 1
            Else
                i = 0
            End If
        End If
        If i > 1 And Len(x) > 2 Then
            Set rng = d1.Range
            rng.Find.Execute x
            rng.Select
            rng.Copy
            d1.Activate
            Set rng2 = d2.Range
            rng2.Find.Execute x
            rng2.Select
            If d2.ActiveWindow.Selection.Range.HighlightColorIndex = 0 Then
                SystemSetup.playSound 2
                Exit Sub
            End If
        End If
    End If
Next
pi = 0
SystemSetup.playSound 3
End Sub

Sub ����W�s����}()
Dim hplnk As Hyperlink, x As String, d As Document
For Each hplnk In ActiveDocument.Hyperlinks
    x = x & hplnk.Address & chr(13)
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
    If left(p.Style.NameLocal, 2) = "���D" Then
        pTxt = p.Range.Text
        pTxt = left(pTxt, Len(pTxt) - 1)
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
If Documents.Count = 0 Then Docs.�ťժ��s���
word.Application.ScreenUpdating = False
If (ActiveDocument.path <> "" And Not ActiveDocument.Saved) Then ActiveDocument.Save
VBA.DoEvents
������Ǯѹq�l�ƭp��.�u�O�d����`��_�B�`��e��[�A��

Dim d As Document, x As String, i As Long
Set d = ActiveDocument
If d.path <> "" Then Exit Sub
If Len(d.Range) = 1 Then Exit Sub '�ťդ�󤣳B�z

'���n�ƻs��ŶKï,�¤�r�ާ@�Y�i
'd.Range.Cut
x = ��r�B�z.trimStrForSearch_PlainText(d.Range)
x = �~�y�q�l���m��Ʈw.CleanTextPicPageMark(x)
SystemSetup.SetClipboard VBA.Replace(x, "�P", "") '�H�m�j�y�šn�۰ʼ��I���|�M���u�P�v�A�y���ѦW�����I������T�A�G�󦹥��M�����C
DoEvents
'If d.path = "" Then '�e�w�@�P�_ If d.path <> "" Then Exit Sub
d.Close wdDoNotSaveChanges

Rem �o��n�g�b���Ϊ����������~���ġA�\��P���֨��]�]��UndoRecord��Application���ݩʡA���b���Q�����ɡA��Ҹ����_��O���]�|�H���M���A�G���g�b���������~���ġ^
SystemSetup.stopUndo ur, "������Ǯѹq�l�ƭp��_����e��[�A��_�K��j�y�Ŧ۰ʼ��I"

'�N�ŶKï�����奻���e�A�e��j�y�Ŧ۰ʼ��I
If �K��j�y�Ŧ۰ʼ��I() = True Then
    If Documents.Count = 0 Then GoTo exitSub
    ActiveDocument.Application.Activate
    '�۰ʰ����������r����
    If Documents.Count > 0 Then
        If InStr(ActiveDocument.path, "�w��B���I") > 0 Then
        On Error GoTo eH:
            If Not SeleniumOP.WD Is Nothing Then
                Dim ws() As String
                ws = SeleniumOP.WindowHandles
                If Not VBA.IsEmpty(ws) Then
                    WD.SwitchTo.Window ws(UBound(ws))
                    SeleniumOP.WD.Manage.Window.Minimize
                End If
            End If
mark:
            ActiveDocument.Application.Activate
            mark��������r
        End If
    End If
End If
exitSub:
SystemSetup.contiUndo ur
word.Application.ScreenUpdating = True
Exit Sub
eH:
    Select Case Err.Number
        Case 9
            If InStr(Err.Description, "�}�C���޶W�X�d��") Then
                GoTo mark
            Else
                GoTo msg:
            End If
        Case Else
msg:
            MsgBox Err.Number + Err.Description
    End Select
End Sub

'���n�ƻs��ŶKï
Function �K��j�y�Ŧ۰ʼ��I() As Boolean
Dim x As String, result As String
On Error GoTo Err1
x = SystemSetup.GetClipboard
x = Replace(x, chr(0), "")
If x = "" Then x = Selection
result = SeleniumOP.grabGjCoolPunctResult(x)
If result = "" Or result = x Then
    DoEvents
    �K��j�y�Ŧ۰ʼ��ISendKeys
Else
    '�g��ŶKï
    SystemSetup.SetClipboard result
    '�����񭵮�
    SystemSetup.playSound 1.469
    �K��j�y�Ŧ۰ʼ��I = True
End If

Exit Function
Err1:
    Select Case Err.Number
        Case 49 'DLL �I�s�W����~
            Resume
        Case 5 'https://www.google.com/search?q=vba+Err.Number+5&oq=vba+Err.Number+5&aqs=chrome..69i57j0i10i30j0i30l2j0i5i30.4768j0j7&sourceid=chrome&ie=UTF-8
            SystemSetup.wait 1.5
            Resume
        Case 13
            If InStr(Err.Description, "���A���ŦX") Then
                SystemSetup.killchromedriverFromHere
'                Stop
                Resume
            Else
                MsgBox Err.Description, vbCritical
                Stop
    '           Resume
            End If
        Case -2146233088
            If InStr(Err.Description, "disconnected: not connected to DevTools") Then '(failed to check if window was closed: disconnected: not connected to DevTools)
                                                                                        '(Session info: chrome=110.0.5481.178)
                SystemSetup.killchromedriverFromHere
'                Stop
                Resume
            Else
                MsgBox Err.Description, vbCritical
                SystemSetup.killchromedriverFromHere
    '           Resume
            End If
        Case Else
            If InStr(Err.Description, "no such window") Then
                If Not WD Is Nothing Then Resume
            Else
                MsgBox Err.Description, vbCritical
                SystemSetup.killchromedriverFromHere
    '           Resume
            End If
    End Select

End Function
Sub �K��j�y�Ŧ۰ʼ��ISendKeys()
'Dim d As Document
'Set d = ActiveDocument
'If d.path <> "" Then Exit Sub
'If SystemSetup.GetClipboard = "" Then
'    If Len(d.Range) = 1 Then Exit Sub '�ťդ�󤣳B�z
'    d.Range.Cut
'End If
On Error GoTo App
AppActivate "�j�y��"
DoEvents
'SendKeys "{TAB 16}", True'�ª�
SendKeys "{TAB 15}", True
Dim x As String
x = SystemSetup.GetClipboardText
If InStr(x, chr(13)) > 0 And InStr(x, chr(13) & chr(10)) = 0 Then
    x = VBA.Replace(x, chr(13), chr(13) & chr(10) & chr(13) & chr(10))
    DoEvents
    SystemSetup.ClipboardPutIn x
End If
SystemSetup.wait 0.5
SendKeys "^v"
DoEvents
'SendKeys "+{TAB 2}~", True '�ª�
SendKeys "+{TAB 1}~", True
wait 2
SendKeys "+{TAB 1} ", True
'If d.path = "" Then d.Close wdDoNotSaveChanges
Exit Sub
App:
Select Case Err.Number
    Case 5
        'Shell (Network.getDefaultBrowserFullname + " https://old.gj.cool/gjcool/index")'�ª�
        Shell (Network.getDefaultBrowserFullname + " https://gj.cool/punct")
        AppActivate Network.getDefaultBrowserNameAppActivate '"�j�y��"
        DoEvents
        SystemSetup.wait 2.9 '2.5 ���}���� ���ݸ��J����
        'SendKeys "{TAB 16}", True
        Resume Next
    Case Else
        MsgBox Err.Number & Err.Description
End Select
End Sub

Rem 20230224 creedit with  Bing���ġG
Sub ChangeFontOfSurrogatePairs_ActiveDocument(fontName As String, Optional whatCJKBlock As CJKBlockName)
    Dim rng         As Range
    Dim c           As String
    Dim i           As Long
    Dim ur As UndoRecord
    SystemSetup.stopUndo ur, "ChangeFontOfSurrogatePairs_ActiveDocument"
    ' Loop through each character in the document
    For Each rng In ActiveDocument.Characters
        c = rng.Text
        ' Check if the character is a high surrogate
        If AscW(c) >= &HD800 And AscW(c) <= &HDBFF Then
            ' Check if the next character is a low surrogate
            If rng.End < ActiveDocument.Content.End Then
                i = rng.End + 1        ' The index of the next character
                If i < ActiveDocument.Range.End Then
                    c = c & ActiveDocument.Range(i, i).Text        ' The combined character
                End If
                If AscW(right(c, 1)) >= &HDC00 And AscW(right(c, 1)) <= &HDFFF Then
                    ' Check if the combined character is in CJK extension B or later
                    'If AscW(Left(c, 1)) >= &HD840 Then
                    If AscW(left(c, 1)) >= SurrogateCodePoint.HighStart Then '�e�ɥN�z (lead surrogates)�A���� D800 �� DBFF �����A�ĤG�ӳQ�٬� ����N�z (trail surrogates)�A���� DC00 �� DFFF ����
                        Dim change As Boolean
                        change = True
'                        rng.Select
                        Select Case whatCJKBlock
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_B
                                change = isCJK_Ext(c, CJK_Unified_Ideographs_Extension_B)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_C
                                change = isCJK_Ext(c, CJK_Unified_Ideographs_Extension_C)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_D
                                change = isCJK_Ext(c, CJK_Unified_Ideographs_Extension_D)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_E
                                change = isCJK_Ext(c, CJK_Unified_Ideographs_Extension_E)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_F
                                'change = isCJK_ExtF(c)
                                change = isCJK_Ext(c, CJK_Unified_Ideographs_Extension_F)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_G
                                change = isCJK_Ext(c, CJK_Unified_Ideographs_Extension_G)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_H
                                change = isCJK_Ext(c, CJK_Unified_Ideographs_Extension_H)
                            Case Else
                            ' Change the font name to HanaMinB
                            ' Change the font name to fontName
                        End Select
                        If change Then rng.Font.Name = fontName '"HanaMinB"
                    End If
                End If
            End If
        End If
    Next rng
    SystemSetup.contiUndo ur
End Sub
Sub ChangeFontOfSurrogatePairs_Range(fontName As String, rngtoChange As Range, Optional whatCJKBlock As CJKBlockName)
    Dim rng         As Range
    Dim c           As String
    Dim i           As Long
    Dim ur As UndoRecord
    SystemSetup.stopUndo ur, "ChangeFontOfSurrogatePairs_Range"
    For Each rng In rngtoChange.Characters
        c = rng.Text
        
        Rem forDebugText
'        If c = ChrW(-10122) & ChrW(-8820) Or c = ChrW(-10119) & ChrW(-8987) Then Stop
        
        ' Check if the character is a high surrogate
        If AscW(c) >= &HD800 And AscW(c) <= &HDBFF Then
'            ' Check if the next character is a low surrogate
'            'If rng.End < ActiveDocument.Content.End Then
'            If rng.End < rngtoChange.End Then
'                i = rng.End + 1        ' The index of the next character
'                'If i < ActiveDocument.Range.End Then
'                If i < rngtoChange.End Then
'                    'c = c & ActiveDocument.Range(i, i).text        ' The combined character
'                    c = c & Mid(rngtoChange, i, 1).text        ' The combined character
'                End If
                If AscW(right(c, 1)) >= &HDC00 And AscW(right(c, 1)) <= &HDFFF Then
                    ' Check if the combined character is in CJK extension B or later
                    'If AscW(Left(c, 1)) >= &HD840 Then
                    If AscW(left(c, 1)) >= SurrogateCodePoint.HighStart Then '�e�ɥN�z (lead surrogates)�A���� D800 �� DBFF �����A�ĤG�ӳQ�٬� ����N�z (trail surrogates)�A���� DC00 �� DFFF ����
                        Dim change As Boolean, isCjkResult As Collection
                        change = True
'                        rng.Select
                        Select Case whatCJKBlock
                            Case CJKBlockName.CJK_Compatibility_Ideographs
                                 Set isCjkResult = IsCJK(c)
                                 If isCjkResult.item(1) Then
                                    If isCjkResult.item(2) <> CJKBlockName.CJK_Compatibility_Ideographs Then change = False
                                 End If
                            Case CJKBlockName.CJK_Compatibility_Ideographs_Supplement
                                 Set isCjkResult = IsCJK(c)
                                 If isCjkResult.item(1) Then
                                    If isCjkResult.item(2) <> CJKBlockName.CJK_Compatibility_Ideographs_Supplement Then change = False
                                 End If
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_B
                                change = isCJK_Ext(c, CJK_Unified_Ideographs_Extension_B)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_C
                                change = isCJK_Ext(c, CJK_Unified_Ideographs_Extension_C)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_D
                                change = isCJK_Ext(c, CJK_Unified_Ideographs_Extension_D)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_E
                                change = isCJK_Ext(c, CJK_Unified_Ideographs_Extension_E)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_F
                                'change = isCJK_ExtF(c)
                                change = isCJK_Ext(c, CJK_Unified_Ideographs_Extension_F)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_G
                                change = isCJK_Ext(c, CJK_Unified_Ideographs_Extension_G)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_H
                                change = isCJK_Ext(c, CJK_Unified_Ideographs_Extension_H)
                            Case Else
                            ' Change the font name to HanaMinB
                            ' Change the font name to fontName
                        End Select
                        If change Then rng.Font.Name = fontName '"HanaMinB"
                    End If
                End If
'            End If
        End If
    Next rng
    SystemSetup.contiUndo ur
End Sub
Sub ChangeCharacterFontName(character As String, fontName As String, d As Document, Optional fontNameFarEast As String)
With d.Range
    With .Find
        With .Replacement.Font
            .Name = fontName
            .NameFarEast = fontNameFarEast
        End With
        .Execute character, , , , , , True, wdFindContinue, , , wdReplaceAll
    End With
End With
End Sub

Sub ChangeCharacterFontNameAccordingSelection()
Dim fontName As String, fontNameFarEast As String
With Selection
    fontName = .Font.Name
    fontNameFarEast = .Font.NameFarEast
    ChangeCharacterFontName .Text, fontName, .Document, fontNameFarEast
End With
End Sub

Rem 20230224 chatGPT�j���ĩ�Bing in Skype ����:
Sub FindMissingCharacters() '�o���ӥu�O���󤤪��r����H�s�ө���B�з������ܪ�
    Dim doc As Document
    Set doc = ActiveDocument
    
    '�w�q�s�ө���M�з���r�������X
    Dim nmf As Font
    Set nmf = doc.Styles("Normal").Font
    Dim kff As Font
    Set kff = doc.Styles("�q��").Font
    
    Dim p As Paragraph
    Dim r As Range
    Dim c As Variant
    
    ' �M�����ɤ����C�Ӭq���M�r��
    For Each p In doc.Paragraphs
        For Each r In p.Range.Characters
            
            ' �P�_�r�ŬO�_�b�s�ө���μз���r����
            c = r.Text
            If Len(c) > 0 Then
                If (AscW(left(c, 1)) >= &H4E00 And AscW(left(c, 1)) <= &H9FFF) _
                    Or (AscW(left(c, 1)) >= &H3400 And AscW(left(c, 1)) <= &H4DBF) _
                    Or (AscW(left(c, 1)) >= &H20000 And AscW(left(c, 1)) <= &H2A6DF) _
                    Or (AscW(left(c, 1)) >= &H2A700 And AscW(left(c, 1)) <= &H2B73F) _
                    Or (AscW(left(c, 1)) >= &H2B740 And AscW(left(c, 1)) <= &H2B81F) _
                    Or (AscW(left(c, 1)) >= &H2B820 And AscW(left(c, 1)) <= &H2CEAF) _
                    Or (AscW(left(c, 1)) >= &HF900 And AscW(left(c, 1)) <= &HFAFF) _
                    Or (AscW(left(c, 1)) >= &H2F800 And AscW(left(c, 1)) <= &H2FA1F) Then '�o�̨S���X�I�A���w���~�A�ݧ�g�I�I�I�I�I�I�I�I
                    If Not r.Font.Name = nmf.Name And Not r.Font.Name = kff.Name Then '�B�Τ���z�b����I�I�I�I
                        ' �p�G�r�Ť��b�s�ө���μз���r�����A�h�N��r���אּHanaMinB
                        r.Font.Name = "HanaMinB"
                    End If
                End If
            End If
        Next r
    Next p
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



