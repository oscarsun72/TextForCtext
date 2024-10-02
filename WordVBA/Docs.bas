Attribute VB_Name = "Docs"
Option Explicit
Public d�r�� As Document, x As New EventClassModule   '�o�~�O�ҿת��إ�"�s"�����O�Ҳ�--��ڤW�O�إ߹復���ѷ�.'��ӽu�W�����DDim�].
'https://learn.microsoft.com/en-us/office/vba/word/concepts/objects-properties-methods/using-events-with-the-application-object-word
Public Sub Register_Event_Handler() '�Ϧ۳]�������O�Ҳզ��Ī��n���{��.���u�ϥ� Application ���� (Application Object) ���ƥ�v
    If x Is Nothing Or Not x.App Is word.Application Then
        SystemSetup.playSound 4
        Set x.App = word.Application '���Y�Ϸs�ت�����PWord.Application����@�W���p
    End If
End Sub

Function �ťժ��s���(Optional newDocVisible As Boolean = True) As Document '20210209
    Dim a As Document, flg As Boolean
    word.Application.ScreenUpdating = False
    If Documents.Count = 0 Then GoTo a:
    If ActiveDocument.Characters.Count = 1 And VBA.InStr(ActiveDocument.Name, "dotm") = 0 Then
        Set a = ActiveDocument
    ElseIf ActiveDocument.Characters.Count > 1 Then
        For Each a In Documents
            If (a.path = "" Or a.Characters.Count = 1) And VBA.InStr(a.Name, "dotm") = 0 Then
    '            a.Range.Paste'��ӳ����K�W�A�{�b���n�A��§�+�}�s���N�n
    '            a.Activate
    '            a.ActiveWindow.Activate
                flg = True
                Exit For
            End If
        Next a
        If flg = False Then GoTo a
    Else
a:     Set a = Documents.Add(Visible:=newDocVisible)
    End If
    Set �ťժ��s��� = a
End Function


Sub �b����󤤴M�����r��_���t() '���w��:Alt+Ctrl+Down 2015/11/1
Static x As String
With Selection
    If .Type = wdSelectionNormal Then
        x = ��r�B�z.trimStrForSearch(.text, Selection)
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
        If Not VBA.Left(h.Sections(i).Range.Paragraphs(2).Range.text, Len(h.Sections(i).Range.Paragraphs(2).Range.text) - 1) Like "..." Then
            rst.FindFirst "�� = " & VBA.Left(h.Sections(i).Range.Paragraphs(1).Range.text, Len(h.Sections(i).Range.Paragraphs(1).Range.text) - 1) _
                & "and ���O like '" & "*" & VBA.Left(h.Sections(i).Range.Paragraphs(2).Range.text, _
                    Len(h.Sections(i).Range.Paragraphs(2).Range.text) - 1) & "*'"
    '            & "and ���O like '" & "*" & Replace(h.Sections(i).Range.Paragraphs(2).Range.Text, vba.Chr(12), "") & "*'"
                'vba.Chr(12)�Ȥ����O����(���O���`�Ÿ�!),���|�v�T���,�G�����N���Ŧr�� _
                �]���̫�@�`�S�����`�Ÿ�(vba.Chr(12))�ӬO�q���Ÿ�(vba.Chr(13),�G�Y�HReplace��ƶ����O�B�z _
                ���K�·�,�@�ߥ�Left��Ƥ����̥k�褧�r���Y�i(���ެOvba.Chr(12)��vba.Chr(13))
        Else '���O��""�ɪ��B�z
            rst.FindFirst "�� = " & VBA.Left(h.Sections(i).Range.Paragraphs(1).Range.text, Len(h.Sections(i).Range.Paragraphs(1).Range.text) - 1) _
                & "and ���O = """"" '�b��,����CSng���A�ഫ�O�i�H��! _
                �]�������p���I,�G�b�@������,�����Words����(�|�N�p���I���Ʀr���}�⦨���P��Word),�p�G���S���p���I���ܴN�i�H�F! _
                ��@,�P���O,�O���F�簣�̥k�誺vba.Chr(10)(����Ÿ��B�q���Ÿ�)
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
f = Array("�P", "�E", "�C", "�v", VBA.Chr(-24152), "�G", "�A", "�F", _
    "�B", "�u", ".", VBA.Chr(34), ":", ",", ";", _
    "�K�K", "...", "�D", "�i", "�j", " ", "�m", "�n", "�q", "�r", "�H" _
    , "�I", "��", "��", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0" _
    , "�y", "�z", VBA.Chr(13), VBA.ChrW(9312), VBA.ChrW(9313), VBA.ChrW(9314), VBA.ChrW(9315), VBA.ChrW(9316) _
    , VBA.ChrW(9317), VBA.ChrW(9318), VBA.ChrW(9319), VBA.ChrW(9320), VBA.ChrW(9321), VBA.ChrW(9322), VBA.ChrW(9323) _
    , VBA.ChrW(9324), VBA.ChrW(9325), VBA.ChrW(9326), VBA.ChrW(9327), VBA.ChrW(9328), VBA.ChrW(9329), VBA.ChrW(9330) _
    , VBA.ChrW(9331), VBA.ChrW(8221), """") '���]�w���I�Ÿ��}�C�H�ƥ�
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
        .Replacement.font.Name = "Arial Unicode MS"
        .Execute VBA.ChrW(i), , , , , , , wdFindContinue, , VBA.ChrW(i), wdReplaceAll
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
    If Selection.Type = wdSelectionNormal And VBA.Right(Selection, 1) Like VBA.Chr(13) Then _
                Selection.MoveLeft wdCharacter, 1, wdExtend '���n�]�t���q�Ÿ�!
    If Selection.Style <> "�ޤ�" Then Selection.Style = "�ޤ�" '�p�G���O�ޤ�˦���,�h�令�ޤ�˦�
    s = Selection.start '�O�U�_�l��m
    Selection.PasteSpecial , , , , wdPasteText '�K�W�¤�r
    e = Selection.End '�O�U�K�W�᪺������m
    Selection.SetRange s, e
    Set r = Selection.Range
    With r
        r.Find.Execute VBA.Chr(13), , , , , , , wdFindStop, , VBA.Chr(11), wdReplaceAll '�N����Ÿ��令��ʤ���Ÿ�
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
        If s < Selection.End Then Selection.text = vbNullString
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
    ay = Array(VBA.ChrW(8220), "�u", VBA.ChrW(8221), "�v", VBA.ChrW(-431), "�B", VBA.ChrW(-432), "�A" _
        , VBA.ChrW(58), "�G", VBA.ChrW(8216), "�y", VBA.ChrW(8217), "�z", _
        VBA.ChrW(-428), "�F", "�P", "�E", ",", "�A", ";", "�F" _
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
        .TypeText VBA.Chr(13)
    Loop
End With
End Sub
Sub OCR���B�z()
Dim a, b, i As Byte
a = Array("�r", VBA.Chr(13) & "�x", " ", "�w", "�q", "�u", "�t", "�{", "�z", "�|", "�}", "�s", "�x", VBA.Chr(9) & VBA.Chr(13), VBA.Chr(13) & VBA.Chr(13), VBA.Chr(13) & VBA.Chr(9))
b = Array("", VBA.Chr(13), "", "", "", "", "", "", "", "", "", "", VBA.Chr(9), VBA.Chr(13), VBA.Chr(13), VBA.Chr(13))
With ActiveDocument
    If .path = "" Then
        For i = 0 To UBound(a) - 1
            With .Range.Find
                .text = a(i)
                With .Replacement
                    .text = b(i)
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
            x = VBA.Mid(x, InStrRev(x, "\") + 1)
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
        If Not p.Range.Characters(p.Range.Characters.Count - 1) Like "[�n�v�z�C�]" & VBA.ChrW(-197) & "0-9a-zA-Z]" Then
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
            If Selection.Next.font.Size > 12 Then '�ؿ��w�]��12�r��
                Selection.MoveRight
                Do Until Selection.Next.font.Size = 12
                    Selection.MoveRight , , wdExtend
                Loop
                If MsgBox("�O�_�n�Y�p��10���r?", vbQuestion + vbOKCancel) = vbOK Then Selection.font.Size = 10
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
        FdText = ��r�B�z.trimStrForSearch(.text, Selection)
        st = .start: ed = .End
        .Collapse wdCollapseEnd
        MnText = .Document.StoryRanges(wdMainTextStory) '�ܼƤƳB�z����2003/4/8
    '    MnText = ActiveDocument.Range '2010/2/5
        ins(1) = InStr(MnText, FdText)
        ins(2) = InStrRev(MnText, FdText)
        
         '�����}�~�ˬd2003/4/3
        If .Document.Footnotes.Count > 0 Then
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
                        .ClearFormatting
                        .ClearAllFuzzyOptions
                        .text = FdText
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
                            .ClearFormatting
                            .ClearAllFuzzyOptions
                            .text = FdText
                            .Forward = True
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
            If ins(1) < ins(2) Then .HomeKey wdStory 'ins(2)�O��󥻤�̫�X�{����m�G 20241002
            With .Find
                .ClearFormatting
                .ClearAllFuzzyOptions
                .Replacement.ClearFormatting '�o�]�n�M���~��
                .Forward = True
                .Wrap = wdFindAsk
                .MatchCase = True
                .text = FdText '.Parent.Text
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
            ActiveDocument.ActiveWindow.Selection.Find.Execute Selection.text
        Case Else
            MsgBox Err.Number & Err.Description
            Resume
    End Select
End Sub


Sub ����_�H�����r�@������() 'ALT+SHIFT+B

' �������s�� 2015/9/20�A���s�� ���[�p
    With ActiveDocument.bookmarks
        .Add Range:=Selection.Range, Name:=Replace(Selection.text, VBA.Chr(13), "")
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
        p.Range.font.Hidden = True
    Else
        prngTxt = p.Range.text
        s = InStr(prngTxt, " ")
        If VBA.Len(VBA.Mid(prngTxt, s + 1)) < 5 Then
            p.Range.font.Hidden = True
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
    If Not p.Range.font.Hidden Then
        prngTxt = p.Range.text
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
                    Debug.Print p.Next.Range.text
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
    x = Selection.text
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
        If Not p.Range.font.Hidden Then 'prepare to delete
            'p.Range.Cut
            x = p.Range.text
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
'        If .Paragraphs(.Paragraphs.Count).Range <> vba.Chr(13) Then .InsertParagraphAfter
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
        With .font
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
        With .font
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
Sub ��������r()
    ' Alt + `
    mark��������r
End Sub
Rem ���槹���~�Ǧ^true�A�_�h��false
Function mark��������r(Optional pasteRange As Range, Optional doNotMark As Boolean) As Boolean
    ' Alt + `
    Dim searchedTerm, e, ur As UndoRecord, d As Document, clipBTxt As String, flgPaste As Boolean, dSource As Document
    Dim strAutoCorrection, endDocOld As Long, rng As Range, returnVaule As Boolean
    Dim punc As New punctuation
    SystemSetup.playSound 0.484
    strAutoCorrection = Array("�A�r", "�r�A", "�q�B", "�q", "�q�C", "�q", "�C�r", "�r", "�q�G", "�q", "�G�r", "�r", "�q�A", "�q", "�B�r", "�r")
    If InStr(ActiveDocument.path, "�������ۤ奻") = 0 Then
        If MsgBox("�ثe���" + ActiveDocument.Name + "�O�_�~��H", vbExclamation + vbOKCancel) = vbCancel Then Exit Function
    End If
    Set dSource = ActiveDocument: If Not dSource.Saved Then dSource.Save
    Set rng = dSource.Range
    With rng.Find
        .ClearAllFuzzyOptions
        .ClearFormatting
        '255�O�W���A���]�i��]�t�F���I�_�y����ӾɭP�奻����������A�٤��p�Y���i�H�ѧO�����קY�i
        'If .Execute(VBA.Trim(VBA.Left(SystemSetup.GetClipboard, 255)), , , , , , True, wdFindContinue) Then
        If .Execute(VBA.Trim(VBA.Left(SystemSetup.GetClipboard, 25)), , , , , , True, wdFindContinue) Then
            rng.Select
            rng.Document.ActiveWindow.ScrollIntoView rng, True
            Exit Function
        End If
    End With
    'If Documents.Count = 0 Then Documents.Add
    If Documents.Count = 0 Then Set d = Docs.�ťժ��s���(True)
    If ClipBoardOp.Is_ClipboardContainCtext_Note_InlinecommentColor Then
        Set d = Docs.�ťժ��s���(False)
        ������Ǯѹq�l�ƭp��.�u�O�d����`��_�B�`��e��[�A�� d
        'Set d = ActiveDocument
        On Error GoTo eH:
        DoEvents
        d.Range.Cut
        d.Close wdDoNotSaveChanges
    End If
    
    'Set d = ActiveDocument
    Set d = dSource
    Rem �]���e���|���u������Ǯѹq�l�ƭp��.�u�O�d����`��_�B�`��e��[�A���v�|�Ψ�UndoRecord����A�B�|��������A�G�H�U����Ҽg��m�N������A�_�h�|�H����������H���L�ġC20230201�ѥf�~�Q�@
    SystemSetup.stopUndo ur, "mark��������r"
    Set rng = d.Range
    endDocOld = d.Range.End
    '    If InStr(d.Range.text, vba.Chr(13) & vba.Chr(13) & vba.Chr(13) & vba.Chr(13)) > 0 Then
    ''        d.Range.Text = Replace(d.Range.Text, vba.Chr(13) & vba.Chr(13) & vba.Chr(13) & vba.Chr(13), vba.Chr(13) & vba.Chr(13) & vba.Chr(13))
    '    '�O�d�榡�A�G�ΥH�U�A���ΥH�W
    '        With d.Range.Find
    '            If InStr(.Parent.text, vba.Chr(13) & vba.Chr(13) & vba.Chr(13) & vba.Chr(13)) > 1 Then
    '                .ClearFormatting
    '                '.Execute vba.Chr(13) & vba.Chr(13) & vba.Chr(13) & vba.Chr(13), , , , , , True, wdFindContinue, , vba.Chr(13) & vba.Chr(13) & vba.Chr(13), wdReplaceAll
                    Rem ����|�y��Word crash
    '                .Execute "^p^p^p^p", , , , , , True, wdFindContinue, , "^p^p^p", wdReplaceAll
    '            End If
    '            .ClearFormatting
    '        End With
    '    End If
    
    Rem �N�ŶKï�����[�J���奻�W�d��
    '"�D�@", ""), "�@�D"���U�j�q���ɮ�A���y�M���A�b�����Τj��ƻs�ɡA�ܭ��n�A�K�o�U�j�q���奻���s�b�@�_�F 20240925
    'clipBTxt = Replace(Replace(Replace(Replace(Replace(VBA.Trim(SystemSetup.GetClipboardText), VBA.Chr(13) + VBA.Chr(10) + "�ťy�l" + VBA.Chr(13) + VBA.Chr(10), VBA.Chr(13) + VBA.Chr(10) + VBA.Chr(13) + VBA.Chr(10)), VBA.Chr(9), ""), "�D�@", ""), "�@�D", ""), " ", vbNullString)
    clipBTxt = Replace(Replace(Replace(VBA.Trim(SystemSetup.GetClipboardText), VBA.Chr(13) + VBA.Chr(10) + "�ťy�l" + VBA.Chr(13) + VBA.Chr(10), VBA.Chr(13) + VBA.Chr(10) + VBA.Chr(13) + VBA.Chr(10)), VBA.Chr(9), ""), " ", vbNullString)
    clipBTxt = ��r�B�z.trimStrForSearch_PlainText(clipBTxt)
    clipBTxt = �~�y�q�l���m��Ʈw.CleanTextPicPageMark(clipBTxt)
    For e = 0 To UBound(strAutoCorrection)
        clipBTxt = Replace(clipBTxt, strAutoCorrection(e), strAutoCorrection(e + 1))
        e = e + 1
    Next e
    searchedTerm = Keywords.����Keywords_ToMark
        
    'If Selection.Type = wdSelectionIP Then
        Rem �P�_�O�_�w�t���Ӥ奻
        '�p�G���t��奻
        If Not Docs.isDocumentContainClipboardText_IgnorePunctuation(d, clipBTxt) Then
            Rem �奻�ۦ��פ��
            Dim similarCompare As New Collection
            Set similarCompare = Docs.similarTextCheckInSpecificDocument(d, clipBTxt)
            If similarCompare.item(1) Then
                word.Application.Activate
'                AppActivate word.ActiveWindow.Caption
                If MsgBox("�奻�ۦ��׬� " & vbCr & similarCompare.item(3) _
                    & VBA.vbCr & vbCr & VBA.vbTab & "�ۦ��q�����G" & VBA.vbCr & VBA.vbCr & VBA.IIf(VBA.Len(similarCompare.item(2)) > 255, VBA.Left(similarCompare.item(2), 255) & "�K�K", similarCompare.item(2)) & vbCr & vbCr & vbCr & _
                    "���U�u�T�w�v�N�|��������q���A�Цۦ��ˬd�O�_���n�A�K�J" & vbCr & vbCr & "���U�u�����v�h�����ˬd�A�N�~�����", vbExclamation + vbOKCancel, "�n�K�J���奻�b���󤤦��������q��!!!") _
                        = vbOK Then
                    Set rng = d.Range
                    If rng.Find.Execute(VBA.Left(similarCompare.item(2), 255), , , , , , , wdFindContinue) Then
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
                'vba.Chrw() & vba.Chrw() &'vba.Chrw() & vba.Chrw() &
                Dim guaKeyword
                guaKeyword = Keywords.����Keywords_ToCheck
                For Each e In guaKeyword
                    If InStr(clipBTxt, e) > 0 Then
                        flgPaste = True
                        Exit For
                    End If
                Next e
            End If
            
            If flgPaste Then
pasteAnyway:
                d.Activate
                If Selection.Document.FullName <> d.FullName Then
                    Stop
                End If
                
                On Error GoTo eH
                If pasteRange Is Nothing Then
                    Selection.EndKey wdStory
                    Selection.InsertParagraphAfter
        '            Selection.InsertParagraphAfter
                Else
                    pasteRange.Select
                End If
                Selection.Collapse wdCollapseEnd
                Selection.TypeText clipBTxt
                'SystemSetup.SetClipboard clipBTxt
                'Docs.�K�W�¤�r
                If pasteRange Is Nothing Then
                    Selection.InsertParagraphAfter: Selection.InsertParagraphAfter: Selection.InsertParagraphAfter
                    Selection.Collapse wdCollapseEnd
                End If
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
                ps = InStr(clipBTxt, VBA.Chr(13)) '�p�����ӭn�K�J���奻�����q���A�h����q���e����F�Y�S���A�h����M�䪺�̤j��255�Ӧr���������e�@�j�M
                sx = VBA.IIf(ps > 0, VBA.Left(VBA.Mid(clipBTxt, 1, VBA.IIf(ps > 0, ps, 2) - 1), 255), VBA.Left(clipBTxt, 255))
            Else '���I�Ÿ��B�z�G�T�w�奻�w���u�O���I�Ÿ����P��
                punc.clearPunctuations clipBTxt
                punc.restoreOriginalTextPunctuations d.Range.text, clipBTxt
                Set punc = Nothing
                sx = ��r�B�z.trimStrForSearch_PlainText(clipBTxt)
                SystemSetup.SetClipboard sx
                sx = VBA.Left(sx, 255)
            End If
            rng.Find.Execute sx, , , , , , , wdFindContinue
            endDocOld = rng.End
    
        End If
    'End If
    
    If flgPaste Then
        word.Application.ScreenUpdating = False
        If d.path <> "" Then
            If InStr(ActiveDocument.path, "�������ۤ奻") = 0 Then
                Set d = dSource
            End If
            If Not d.Saved Then d.Save
        End If
        
        Rem ��������r
        If Not doNotMark Then
        '    xd = d.Range.text
            Dim rngMark As Range
            
            Set rngMark = d.Range(IIf(endDocOld >= d.Range.End, d.Range.End - 1, endDocOld), d.Range.End)
            
            marking��������r rngMark, searchedTerm, word.wdYellow, wdRed, False
            
        End If
        Rem �H�W��������r
        
        GoSub refres
        SystemSetup.playSound 1.921
        Rem https://en.wikipedia.org/wiki/CJK_Unified_Ideographs
        Rem �ݮe�r
        'https://en.wikipedia.org/wiki/CJK_Compatibility_Ideographs
    '    Docs.ChangeFontOfSurrogatePairs_Range "HanaMinA", d.Range(selection.Paragraphs(1).Range.start, d.Range.End), CJK_Compatibility_Ideographs
        'https://en.wikipedia.org/wiki/CJK_Compatibility_Ideographs_Supplement
        Dim rngChangeFontName As Range
        'Set rngChangeFontName = d.Range(Selection.Paragraphs(1).Range.start, d.Range.End)
        Set rngChangeFontName = d.Range(rngMark.start, d.Range.End)
        Dim fontName As String '20240920 creedit_with_Copilot�j����:https://sl.bing.net/9KC0PtODtI
        fontName = "������-2"
        If Fonts.IsFontInstalled(fontName) Then
            'MsgBox fontName & " �w�w�˦b�t�Τ��C"
        Else    'MsgBox fontName & " ���w�˦b�t�Τ��C"
            fontName = "HanaMinA"
            If Fonts.IsFontInstalled("HanaMinA") Then
            ElseIf Fonts.IsFontInstalled("TH-Tshyn-P2") Then
                fontName = "TH-Tshyn-P2"
            Else
                fontName = vbNullString
            End If
        End If
        If Not fontName = vbNullString Then
            'Docs.ChangeFontOfSurrogatePairs_Range "HanaMinA", rngChangeFontName, CJK_Compatibility_Ideographs_Supplement
            Docs.ChangeFontOfSurrogatePairs_Range fontName, rngChangeFontName, CJK_Compatibility_Ideographs_Supplement
        End If
        
        Rem �X�R�r��
        'HanaMinB�٤��䴩G�H�᪺
        fontName = "HanaMinB"
        Docs.ChangeFontOfSurrogatePairs_Range fontName, rngChangeFontName, CJK_Unified_Ideographs_Extension_E
        Docs.ChangeFontOfSurrogatePairs_Range fontName, rngChangeFontName, CJK_Unified_Ideographs_Extension_F
        returnVaule = True
        
    Else '��󤺤w�����e��
        GoSub refres
        SystemSetup.playSound 1.294
        If noneYijingKeyword Then
            If MsgBox("�n�K�W���奻�ä��t����������r�@�I" + vbCr + vbCr + _
                "�ЦA�ˬd�ҽƻs��ŶKï�����e�O�_���T�C�P���P���@�n�L��������" & _
                "���O�_���n�K�W�H" + vbCr + vbCr + clipBTxt, vbOKCancel + vbExclamation + vbDefaultButton2) _
                = vbOK Then
                noneYijingKeyword = False
                GoTo pasteAnyway
            End If
        End If
    End If
    
exitSub:
    SystemSetup.contiUndo ur
    Set ur = Nothing
    'word.Application.ScreenUpdating = True
    'word.Application.ScreenRefresh
    mark��������r = returnVaule
    Exit Function
    
    
refres:
        word.Application.ScreenUpdating = True
        If flgPaste Then
            Rem ���ٲ��A�K�o�C���K�J�����@���A��r�B�z.�ѦW���g�W���Ъ`�A���������ɭn�����ɮ׫e�A��
            '��r�B�z.�ѦW���g�W���Ъ`
            'If flgPaste Then'���յLê��i�R����
            '��ܷs�K�W���奻����
            rng.SetRange endDocOld, endDocOld
            Do Until rng.font.ColorIndex = wdRed Or rng.End = d.Range.End - 1
                rng.Move
            Loop
            e = rng.End
            rng.SetRange endDocOld, e
        Else
            rng.SetRange endDocOld - Len(sx), endDocOld
        End If
        rng.Select
        Static cntr  As Byte

        '�p�G����d��O��󥽺ݡA��Y�S���A�G�H�ŶKï���e25�r�A��@��

        If rng.End = rng.Document.Range.End Then

            SystemSetup.SetClipboard VBA.Left(SystemSetup.GetClipboard, 25)

            If cntr < 2 Then
                cntr = cntr + 1
                If VBA.vbOK = MsgBox("�S����m�A�O�_���աH�P���P���@�n�L��������", vbOKCancel + vbExclamation) Then
                    mark��������r
                Else
                    cntr = 2
                End If
            Else
                cntr = 0
            End If

        End If
    '    word.Application.ScreenRefresh
        ActiveWindow.ScrollIntoView Selection.Characters(1) ', False
    Return
    
eH:
    Select Case Err.Number
        Case 5825 '����w�Q�R���C
            GoTo exitSub
        Case Else
            MsgBox Err.Number & Err.Description
            Resume
    End Select
End Function
Rem �ھڿ����r���G��󤤩Ҧ�����r�X�{�������A�Ұ��G�̨ä��H����x�s�A�ȧ@��ܺ� 20240922
Sub HitHighlightBySelecton()
    If Selection.Type = wdSelectionIP Then Exit Sub
    Dim rng As Range
    Set rng = Selection.Document.Range
    rng.Find.HitHighlight Selection.text, wdColorYellow
    Debug.Print rng.start; rng.End; rng.Document.Range.start; rng.Document.Range.End
End Sub

Rem 20240922 Copilot�j���Įھڧڪ���}�A�ո� https://sl.bing.net/dtWVmyauIFw https://sl.bing.net/glAQGL0KKCO
Rem �޿���~�A��������I
Sub marking��������r_BAD_Copilot�j����(rng As Range, arr As Variant, Optional defaultHighlightColorIndex As word.WdColorIndex = word.wdYellow, _
        Optional fontColor As word.WdColorIndex = word.wdRed, Optional allDoc As Boolean = False)
    
    Dim regex As Object, matches As Object, match As Object
    Dim startRng As Long, endRng As Long
    Dim e As Variant
    Dim examOK As Boolean, rngExam As Range
    Dim isInPhrasesAvoid As Boolean, isFollowedAvoid As Boolean, isPrecededAvoid As Boolean
    Dim dictCoordinatesPhrase As New Scripting.Dictionary, key
    
    On Error GoTo eH
    word.options.defaultHighlightColorIndex = defaultHighlightColorIndex
    If allDoc Then Set rng = rng.Document.Range
    startRng = rng.start
    endRng = rng.End
    Set rngExam = rng.Document.Range
    
    ' �إߥ��h��F����H
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    
    For Each e In arr
        'Copilot�j���ġG�o�˪����h��F���Ҧ��T�O�F�ڭ̥u�ǰt��ӳ�� e�A�Ӥ��O������@�����C�Ҧp�A�p�G e �O ��word���A�o�ӼҦ��|�ǰt ��word���A�����|�ǰt ��wording�� �� ��sword���C
        'regex.Pattern = "\b" & e & "\b" 'Copilot�j���ġGb �b���h��F�����N���O�u�����ɡv�]word boundary�^�A�i�H�z�Ѭ��ubound�v���Y�g�C���ΨӤǰt������}�l�ε�����m�A�T�O�ڭ̥u�ǰt��ӳ���Ӥ��O������@�����C https://sl.bing.net/kIwNUNbDloO
        regex.Pattern = "." & e & "."
        Set matches = regex.Execute(rng.text)
        If matches.Count > 0 Then Stop
        For Each match In matches
            examOK = True
            rng.SetRange startRng + match.FirstIndex, startRng + match.FirstIndex + Len(match.Value)
            
            ' �ˬd�O�_�ݭn�e����ˬd
            isFollowedAvoid = Keywords.����KeywordsToMark_Exam_Followed_Avoid.Exists(e)
            isPrecededAvoid = Keywords.����KeywordsToMark_Exam_Preceded_Avoid.Exists(e)
            isInPhrasesAvoid = Keywords.����KeywordsToMark_Exam_InPhrase_Avoid.Exists(e)
            
            If isFollowedAvoid Or isPrecededAvoid Or isInPhrasesAvoid Then
                ' ����ˬd
                If Not rng.Next Is Nothing And rng.Next.Characters.Count > 0 Then
                    If isFollowedAvoid Then
                        For Each key In Keywords.����KeywordsToMark_Exam_Followed_Avoid(e)
                            If rng.End + Len(key) <= endRng Then
                                rngExam.SetRange rng.End, rng.End + Len(key)
                                If StrComp(rngExam.text, key) = 0 Then
                                    examOK = False
                                    Exit For
                                End If
                            End If
                        Next key
                    End If
                End If
                
                ' �e���ˬd
                If examOK And Not rng.Previous Is Nothing And rng.Previous.Characters.Count > 0 Then
                    If isPrecededAvoid Then
                        For Each key In Keywords.����KeywordsToMark_Exam_Preceded_Avoid(e)
                            If rng.start - Len(key) > -1 Then
                                rngExam.SetRange rng.start - Len(key), rng.start
                                If StrComp(rngExam.text, key) = 0 Then
                                    examOK = False
                                    Exit For
                                End If
                            End If
                        Next key
                    End If
                End If
                
                ' ���O���ˬd
                If examOK And isInPhrasesAvoid Then
                    If dictCoordinatesPhrase.Count = 0 Then
                        For Each key In Keywords.����KeywordsToMark_Exam_InPhrase_Avoid(e)
                            rngExam.SetRange startRng, endRng
                            Do While rngExam.Find.Execute(key, , , , , , True, wdFindStop)
                                dictCoordinatesPhrase.Add rngExam.start, rngExam.End
                            Loop
                        Next key
                    End If
                    
                    For Each key In dictCoordinatesPhrase
                        If rng.start >= key And rng.End <= dictCoordinatesPhrase(key) Then
                            examOK = False
                            Exit For
                        End If
                    Next key
                End If
            End If
            
            ' ��������r
            If examOK Then
                With rng
                    .HighlightColorIndex = defaultHighlightColorIndex
                    If .font.ColorIndex = wdAuto Then .font.ColorIndex = fontColor
                End With
            Else
                ' ���ݭn�e����ˬd������
                Do While regex.Execute(e, , , , , , True, wdFindStop, True)
                    With rng
                        .HighlightColorIndex = defaultHighlightColorIndex
                        If .font.ColorIndex = wdAuto Then .font.ColorIndex = fontColor
                    End With
                Loop
            End If
        Next match
        
        If dictCoordinatesPhrase.Count > 0 Then dictCoordinatesPhrase.RemoveAll
    Next e

finish:
    rng.SetRange startRng, endRng
    Exit Sub

eH:
    MsgBox Err.Number & Err.Description
    Resume finish
End Sub

Rem rng �n�B�z���d�� ,arr �n�B�z������r �]�w�]���r��}�C�^
Sub marking��������r(rng As Range, arr As Variant, Optional defaultHighlightColorIndex As word.WdColorIndex = word.wdYellow, _
        Optional fontColor As word.WdColorIndex = word.wdRed, Optional allDoc As Boolean = False)
    '�u�ƪ���ĳ�GCopilot�j���� 20240922 Word VBA ���� Find �����ݩʡG https://sl.bing.net/fMV3NYyXdLg https://sl.bing.net/kosqk2rrnFc
    Dim xd As String, a As Range, e, eArrKey, arrKey, startRng As Long, endRng As Long ', dict As Scripting.Dictionary
    Dim examOK As Boolean, rngExam As Range, processCntr As Long, dictCoordinatesPhrase As New Scripting.Dictionary, key, isInPhrasesAvoid As Boolean, isFollowedAvoid As Boolean, isPrecededAvoid As Boolean

    On Error GoTo eH
    word.options.defaultHighlightColorIndex = defaultHighlightColorIndex
    examOK = True
    If allDoc Then Set rng = rng.Document.Range
    startRng = rng.start
    endRng = rng.End
    Set rngExam = rng.Document.Range

    xd = rng.text

    With rng.Find
        .ClearFormatting
'        .HitHighlight
        
'        With .Replacement '�{�b���� wdReplaceAll �޼ƪ���k�F
'            .font.ColorIndex = fontColor 'wdRed
'            .Highlight = True
'        End With
        For Each e In arr '�M���C�ӭn���Ѫ�����r
        
'            If e = "�{��" Then Stop 'just for test
            
            If InStr(xd, e) > 0 Then '�b���W�s�����\���ܼơB���ä�r�ɥi��|miss�A�����ըä��|�A�ݦA���աC
            
                rng.SetRange startRng, endRng
                'If e = "��" Then
                isFollowedAvoid = Keywords.����KeywordsToMark_Exam_Followed_Avoid.Exists(e)
                isPrecededAvoid = Keywords.����KeywordsToMark_Exam_Preceded_Avoid.Exists(e)
                isInPhrasesAvoid = Keywords.����KeywordsToMark_Exam_InPhrase_Avoid.Exists(e)
                If isFollowedAvoid Or isPrecededAvoid Or isInPhrasesAvoid Then
'                    '�Y�����O�󪺤��y�y���]�]�t���䪺���y�y�y�^�ݬd�窺�ܡA�N���O�U�n��諸��m���X
'                    '�Y��B�e�ˬd�����L�N�����F�A�G���A�ﲾ�ܤ��O���ˬd�A�åH dictCoordinatesPhrase.Count �ӧ@�P�_
'                    If isInPhrasesAvoid Then
'                        arrKey = Keywords.����KeywordsToMark_Exam_InPhrase_Avoid(e)
'                        For Each eArrKey In arrKey '�M���C�Ӹ��׶}�����y�A�`����b��󤤪��Ҧb��m�A�H�ѫ�����
'                            'Set rngExam = rngExam.Document.Range '�˿��F�A�w��allDoc�ѼƱ���O�_�ާ@������
'                            rngExam.SetRange startRng, endRng '���ݳB�z������A�N�u�N�ާ@�d�򤺳B�m�N�i�C�P���P���@�g���g�ۡ@�n�L�������� 20240915
'                            With rngExam.Find
'                                Do While .Execute(eArrKey, , , , , , True, wdFindStop)
'                                    '�O�U�t���ثe����r�����y���y�y�J���q�b�d�򤤪���m
'                                    dictCoordinatesPhrase.Add rngExam.start, rngExam.End
'                                Loop
'                            End With
'                        Next eArrKey
'                    End If
                    
                    Do While .Execute(e, , , , , , True, wdFindStop, True) '�b�d�򤤴M�M����r�X�{����m
                        examOK = True '�k�s
                        
                        'rng.Select '������
                        
                        If Not rng.Next Is Nothing Then
                            If rng.Next.Characters.Count > 0 Then
                                '����ˬd
                                
                                'If rng.Document.Range(rng.End, rng.End + 4).text = "�[���M��" Then Stop 'just for test
                                
                                'If UBound(VBA.Filter(Keywords.����KeywordsToMark_Exam_Followed_Avoid(e), rng.Next.Characters(1).text)) < 0 Then
                                If isFollowedAvoid Then
                                    arrKey = Keywords.����KeywordsToMark_Exam_Followed_Avoid(e)
                                    For Each eArrKey In arrKey
                                        If rng.End + VBA.Len(eArrKey) <= endRng Then '�o�˪��g�k�A�p�G���t�W�s�����\���ܼơA���ȴN�|���~�F�I
                                            rngExam.SetRange rng.End, rng.End + VBA.Len(eArrKey)
                                            If VBA.StrComp(rngExam.text, eArrKey) = 0 Then '���n�׶}������r
                                                examOK = False '�˴����q�L
                                                Exit For
                                            End If
                                        End If
                                    Next eArrKey
                                End If
                                If examOK Then
checkPrevious:
                                    If Not rng.Previous Is Nothing Then
                                        '�e���ˬd
                                        If rng.Previous.Characters.Count > 0 Then
                                            'If UBound(VBA.Filter(Keywords.����KeywordsToMark_Exam_Preceded_Avoid(e), rng.Previous.Characters(rng.Previous.Characters.Count).text)) < 0 Then
                                            If isPrecededAvoid Then
                                                arrKey = Keywords.����KeywordsToMark_Exam_Preceded_Avoid(e)
                                                For Each eArrKey In arrKey
                                                    If rng.start - VBA.Len(eArrKey) > -1 Then '�o�˪��g�k�A�p�G���t�W�s�����\���ܼơA���ȴN�|���~�F�I
                                                        rngExam.SetRange rng.start - VBA.Len(eArrKey), rng.start
                                                        If VBA.StrComp(rngExam.text, eArrKey) = 0 Then '���n�׶}������r
                                                            examOK = False '�˴����q�L
                                                            Exit For
                                                        End If
                                                    End If
                                                Next eArrKey
                                            End If
                                            If examOK Then
checkPhrases:                                   '���O���ˬd�G����r�t�b���קK�����y���y�ˬd
                                                'If Keywords.����KeywordsToMark_Exam_InPhrase_Avoid.Exists(e) Then
                                                If isInPhrasesAvoid Then
                                                    If dictCoordinatesPhrase.Count = 0 Then
                                                        GoSub buildDictCoordinatesPhrase
                                                    End If
                                                
                                                    For Each key In dictCoordinatesPhrase
                                                     '�M���C�Ӹ��׶}�����y���y�y��
                                                        '�Y�ثe����r���t��n�׶}�����y���y�y�J���q
                                                        If rng.start >= key And rng.End <= dictCoordinatesPhrase(key) Then

                                                            'rng.Select 'just for test

                                                            examOK = False '�˴����q�L
                                                            Exit For
                                                        End If
                                                    Next key

'                                                    arrKey = Keywords.����KeywordsToMark_Exam_InPhrase_Avoid(e)
'                                                    For Each eArrKey In arrKey '�M���C�Ӹ��׶}�����y�A����`�����m�b��󤤩Ҧb��m�A�H�ѫ�����
'                                                        Set rngExam = rngExam.Document.Range
'                                                        With rngExam.Find
'                                                            Do While .Execute(eArrKey, , , , , , True, wdFindStop)
'                                                                '�ثe����r���t��n�׶}�����y���y�y�J���q
'                                                                If rng.start >= rngExam.start And rng.End <= rngExam.End Then
'
''                                                                    rng.Select 'just for test
'
'                                                                    examOK = False '�˴����q�L
'                                                                    Exit For
'                                                                End If
'                                                            Loop
'                                                        End With
'                                                    Next eArrKey
                                                    
                                                End If
                                                '��B�e�B���T��������X��F
                                                If examOK Then '�X��~����
                                                    With rng
                                                        processCntr = processCntr + 1
                                                        If processCntr Mod 35 = 0 Then SystemSetup.playSound 1 '���񭵮ĥH�K�~�H����F
                                                        Rem ���ɥi�Ѯį���աA�]�_�Ӥ�����G�S�O�[�I file:///H:\�ڪ����ݵw��\���Ѯv���ݤu�@\1�������ۤ奻\�D�M�H�w��B���I\��}�T�I�|���N�O��.docx
                                                        
                                                        .HighlightColorIndex = defaultHighlightColorIndex
                                                        For Each a In rng.Characters
                                                            If a.font.ColorIndex = wdAuto Then a.font.ColorIndex = fontColor
                                                        Next a
                                                    End With
                                                Else '���y���y�ˬd���L
                                                    If rng.HighlightColorIndex = defaultHighlightColorIndex Then
                                                        With rng
                                                            '.Select 'just for test
                                                            
                                                            .HighlightColorIndex = wdNoHighlight
                                                            .font.ColorIndex = wdAuto
                                                        End With
                                                    End If
                                                End If
                                            Else '�e���ˬd���L
                                                'If allDoc Then
                                                If rng.HighlightColorIndex = defaultHighlightColorIndex Then
                                                    With rng
                                                        .HighlightColorIndex = wdNoHighlight
                                                        .font.ColorIndex = wdAuto
                                                    End With
                                                End If
                                            End If
                                        End If
                                    Else '�p�G�e�]�L��]�L
                                        GoTo checkPhrases:
                                    End If
                                Else '����ˬd���L
                                    'If allDoc Then
                                    If rng.HighlightColorIndex = defaultHighlightColorIndex Then
                                        With rng
                                            .HighlightColorIndex = wdNoHighlight
                                            .font.ColorIndex = wdAuto
                                        End With
                                    End If
                                End If
                            End If
                        Else '�p�G�S���U��
                            GoTo checkPrevious
                        End If
                    Loop 'Do While .Execute(e, , , , , , True, wdFindStop, True) '�b�d�򤤴M�M����r�X�{����m

                    If dictCoordinatesPhrase.Count > 0 Then
                        dictCoordinatesPhrase.RemoveAll '�k�s�ѤU�@������r�ϥ�
                    End If
                        'If Keywords.����KeywordsToMark_Exam_InPhrase_Avoid.Exists(e) Then
                         '       �K�K
                          '          dictCoordinatesPhrase.Add rngExam.start, rngExam.End
                          
                Else '�����ˬd�������N�]�w����ѡ^��
                    Do While .Execute(e, , , , , , True, wdFindStop, True) '���į�]���|�� wdReplaceAll �޼ƪ̺C�A�i���䤺���Y���������j���@�̤] 20240919 �P���P���@�g���g�ۡ@�n�L��������
'                        .Parent.HighlightColorIndex = defaultHighlightColorIndex
'                        .Parent.Font.ColorIndex = fontColor
                        Rem �Y�g���H�U�|��49DLL�I�s�W����~�A�o�����g�S���|�F�A�i���O VBE�sĶ���G�� 20240920
                        With rng
                            .HighlightColorIndex = defaultHighlightColorIndex
                            For Each a In rng.Characters
                                If a.font.ColorIndex = wdAuto Then a.font.ColorIndex = fontColor
                            Next a
                            processCntr = processCntr + 1
                            If processCntr Mod 35 = 0 Then SystemSetup.playSound 1 '���񭵮ĥH�K�~�H����F
                        End With
                    Loop
'                    .Execute e, , , , , , True, wdFindStop, True, e, Replace:=wdReplaceAll '�b�t���W�s�����榡�Ƥ�r�ɷ|���F
                    'rng.SetRange startRng, endRng'�e�w��
                End If
            End If
        Next e '�U�@�ӥ���n���Ѫ�����r
    End With

finish:
    rng.SetRange startRng, endRng '�]�^��Ӫ��ˤl�~���|���ܡA��I�s�ݤ~���|�X��
    Set dictCoordinatesPhrase = Nothing
    
    Exit Sub
    
buildDictCoordinatesPhrase:
        '�Y�����O�󪺤��y�y���]�]�t���䪺���y�y�y�^�ݬd�窺�ܡA�N���O�U�n��諸��m���X
'        �Y��B�e�ˬd�����L�N�����F�A�G���A�ﲾ�ܤ��O���ˬd�A�åH dictCoordinatesPhrase.Count �ӧ@�P�_
        If isInPhrasesAvoid Then
            arrKey = Keywords.����KeywordsToMark_Exam_InPhrase_Avoid(e)
            For Each eArrKey In arrKey '�M���C�Ӹ��׶}�����y�A�`����b��󤤪��Ҧb��m�A�H�ѫ�����
                    '�w��allDoc�ѼƱ���O�_�ާ@������
                rngExam.SetRange startRng, endRng '�G���ݳB�z������A�N�u�N�ާ@�d�򤺳B�m�N�i�C�P���P���@�g���g�ۡ@�n�L�������� 20240915
                With rngExam.Find
                    Do While .Execute(eArrKey, , , , , , True, wdFindStop)
'                        �O�U�t���ثe����r�����y���y�y�J���q�b�d�򤤪���m
                        dictCoordinatesPhrase.Add rngExam.start, rngExam.End
                    Loop
                End With
            Next eArrKey
        End If
        
        Return

eH:
    Select Case Err.Number
        Case Else
            MsgBox Err.Number & Err.Description
            Debug.Print Err.Number & Err.Description
            Resume
    End Select
End Sub
Rem ��Ӥ�󭫷s���ѩ�������r
Sub mark��������rDoc()
    Dim ur As word.UndoRecord
    SystemSetup.playSound 0.484
    SystemSetup.stopUndo ur, "mark��������rDoc"
    word.Application.ScreenUpdating = False
    marking��������r ActiveDocument.Range, Keywords.����Keywords_ToMark, word.wdYellow
    SystemSetup.contiUndo ur
    SystemSetup.playSound 2
    word.Application.ScreenUpdating = True
End Sub

Rem �P�_�ŶKï�̪��¤�r(�Ϋ��w����r)���e�O�_�b��󤤤w�s�b
Function isDocumentContainClipboardText_IgnorePunctuation(d As Document, Optional chkClipboardText As String) As Boolean
    Dim xd As String
    xd = d.Range.text
    If VBA.Len(xd) = 1 Then Exit Function
    
    If chkClipboardText = "" Then chkClipboardText = SystemSetup.GetClipboardText
    Rem �ŶKï�̪�����Ÿ��ȬOvba.Chr(13)&vba.Chr(10)�ӦbWord��󤤬O�u�� vba.Chr(13)
    chkClipboardText = VBA.Replace(chkClipboardText, VBA.Chr(13) & VBA.Chr(10), VBA.Chr(13))
    
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

Function similarTextCheckInSpecificDocument(d As Document, text As String) As Collection 'item1 as Boolean(�奻�O�_�ۦ�),item2 as string(��쪺�ۦ��奻�q��),item3 as String from Dictionary SimilarityResult(�ۦ��צW&�ۦ���)
Rem �奻�ۦ��פ��
Dim similarText As New similarText, dClearPunctuation As String, textClearPunctuation As String, dCleanParagraphs() As String, punc As New punctuation, e, Similarity As Boolean, result As New Collection
dClearPunctuation = d.Content.text
textClearPunctuation = text
'�M�����I�Ÿ�
punc.clearPunctuations textClearPunctuation: punc.clearPunctuations dClearPunctuation
dCleanParagraphs = VBA.Split(dClearPunctuation, VBA.Chr(13))
Dim cntr As Long
For Each e In dCleanParagraphs
    cntr = cntr + 1
    If cntr Mod 20 = 0 Then SystemSetup.playSound 1
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
punc.restoreOriginalTextPunctuations d.Content.text, dClearPunctuation
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
d1RngTxt = d1.Range.text
Set d2 = Documents(2) '��ŧ�Τޥ�(�����N�C�A���y�l����U�q��r�^
pc = d2.Paragraphs.Count
If pi = 0 Then pi = 1
For pi = pi To pc
    Set p = d2.Paragraphs(pi)
    If p.Range.font.NameFarEast <> "�з���" And p.Range.HighlightColorIndex = 0 Then
        px = p.Range
        x = VBA.Trim(VBA.Left(px, Len(px) - 1)) '�h�����q�Ÿ�
        If Len(x) > 2 Then
            x = VBA.Left(x, Len(x) - 1) '�h���ݫ���I�C�A��
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
    x = x & hplnk.Address & VBA.Chr(13)
Next hplnk
Set d = Documents.Add
d.Range.text = x
d.Range.Cut
d.Close wdDoNotSaveChanges
End Sub


Sub ���J�W�s��_��󤤪���m_���D() 'Alt+P ��O�u�޸֡v�˦�'2021/11/27
Dim d As Document, title As String, p As Paragraph, pTxt As String, subAddrs As String, flg As Boolean
Set d = ActiveDocument
title = Selection.text
title = ��r�B�z.trimStrForSearch(title, Selection)
For Each p In d.Paragraphs
    If VBA.Left(p.Style.NameLocal, 2) = "���D" Then
        pTxt = p.Range.text
        pTxt = VBA.Left(pTxt, Len(pTxt) - 1)
        If StrComp(pTxt, title) = 0 Then
            subAddrs = title
            flg = True
            Exit For
        ElseIf InStr(pTxt, title) > 0 Then
            subAddrs = "_" & VBA.Mid(pTxt, 1, InStrRev(pTxt, " ") - 1)
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
Rem ������ɡA�Y�H�������r�B�z
Sub ������Ǯѹq�l�ƭp��_�u�O�d����`��_�B�`��e��[�A��_�K��j�y�Ŧ۰ʼ��I()
    Dim ur As UndoRecord, d As Document, x As String, i As Long
    Dim SelectionRange As Range
    SystemSetup.playSound 0.484
    If (ActiveDocument.path <> "" And Not ActiveDocument.Saved) Then ActiveDocument.Save
    
    Rem �o��n�g�b���Ϊ����������~���ġA�\��P���֨��]�]��UndoRecord��Application���ݩʡA���b���Q�����ɡA��Ҹ����_��O���]�|�H���M���A�G���g�b���������~���ġ^
    SystemSetup.stopUndo ur, "������Ǯѹq�l�ƭp��_����e��[�A��_�K��j�y�Ŧ۰ʼ��I"
    
    If Selection.Type = wdSelectionNormal Then
        Selection.Cut
        Set SelectionRange = Selection.Range
    End If
    'If Documents.Count = 0 Then
    '    Set d = Docs.�ťժ��s���()
    'Else
    '    Set d = ActiveDocument
    'End If
    word.Application.ScreenUpdating = False
    Set d = Docs.�ťժ��s���()
    VBA.DoEvents
    ������Ǯѹq�l�ƭp��.�u�O�d����`��_�B�`��e��[�A�� d
    
    If d.path <> "" Then
        MsgBox "�����ɤw�x�s�A����ާ@�I", vbCritical
        Exit Sub
    End If
    If Len(d.Range) = 1 Then Exit Sub '�ťդ�󤣳B�z
    
    '�H�U2��w�����A���[�� 20240716
    '���n�ƻs��ŶKï,�¤�r�ާ@�Y�i
    'd.Range.Cut
    
    x = ��r�B�z.trimStrForSearch_PlainText(d.Range.text)
    x = �~�y�q�l���m��Ʈw.CleanTextPicPageMark(x)
    SystemSetup.SetClipboard VBA.Replace(x, "�P", "") '�H�m�j�y�šn�۰ʼ��I���|�M���u�P�v�A�y���ѦW�����I������T�A�G�󦹥��M�����C
    DoEvents
    'If d.path = "" Then '�e�w�@�P�_ If d.path <> "" Then Exit Sub
    d.Close wdDoNotSaveChanges
    
    Rem �o��n�g�b���Ϊ����������~���ġA�\��P���֨��]�]��UndoRecord��Application���ݩʡA���b���Q�����ɡA��Ҹ����_��O���]�|�H���M���A�G���g�b���������~���ġ^
    Rem SystemSetup.stopUndo ur, "������Ǯѹq�l�ƭp��_����e��[�A��_�K��j�y�Ŧ۰ʼ��I"
    
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
                mark��������r SelectionRange
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
            Case -2146233088 'unknown error: unhandled inspector error: {"code":-32000,"message":"Browser window not found"}
                              '(Session info: chrome=126.0.6478.127)
                If VBA.InStr(Err.Description, "unknown error: unhandled inspector error:") > 0 Then
                    GoTo mark
                End If
            Case Else
msg:
                MsgBox Err.Number & Err.Description
        End Select
End Sub

'���n�ƻs��ŶKï
Function �K��j�y�Ŧ۰ʼ��I() As Boolean
    Dim x As String, result As String, resumeTimer As Byte
    On Error GoTo Err1
    x = SystemSetup.GetClipboard
    x = Replace(x, VBA.Chr(0), "")
    If x = "" Then x = Selection
    result = SeleniumOP.grabGjCoolPunctResult(x, result)
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
                resumeTimer = resumeTimer + 1
                If resumeTimer > 2 Then
                    MsgBox Err.Description, vbCritical
                    SystemSetup.killchromedriverFromHere
                Else
                    Resume
                End If
            Case 5 'https://www.google.com/search?q=vba+Err.Number+5&oq=vba+Err.Number+5&aqs=chrome..69i57j0i10i30j0i30l2j0i5i30.4768j0j7&sourceid=chrome&ie=UTF-8
                SystemSetup.wait 1.5
                resumeTimer = resumeTimer + 1
                If resumeTimer > 2 Then
                    MsgBox Err.Description, vbCritical
                    SystemSetup.killchromedriverFromHere
                Else
                    Resume
                End If
            Case 13
                If InStr(Err.Description, "���A���ŦX") Then
                    SystemSetup.killchromedriverFromHere
    '                Stop
                    resumeTimer = resumeTimer + 1
                    If resumeTimer > 2 Then
                        MsgBox Err.Description, vbCritical
                        SystemSetup.killchromedriverFromHere
                    Else
                        Resume
                    End If
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
                    resumeTimer = resumeTimer + 1
                    If resumeTimer > 2 Then
                        MsgBox Err.Description, vbCritical
                        SystemSetup.killchromedriverFromHere
                    Else
                        Resume
                    End If
    
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
If InStr(x, VBA.Chr(13)) > 0 And InStr(x, VBA.Chr(13) & VBA.Chr(10)) = 0 Then
    x = VBA.Replace(x, VBA.Chr(13), VBA.Chr(13) & VBA.Chr(10) & VBA.Chr(13) & VBA.Chr(10))
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
    Dim C           As String
    Dim i           As Long
    Dim ur As UndoRecord
    SystemSetup.stopUndo ur, "ChangeFontOfSurrogatePairs_ActiveDocument"
    ' Loop through each character in the document
    For Each rng In ActiveDocument.Characters
        C = rng.text
        ' Check if the character is a high surrogate
        If AscW(C) >= &HD800 And AscW(C) <= &HDBFF Then
            ' Check if the next character is a low surrogate
            If rng.End < ActiveDocument.Content.End Then
                i = rng.End + 1        ' The index of the next character
                If i < ActiveDocument.Range.End Then
                    C = C & ActiveDocument.Range(i, i).text        ' The combined character
                End If
                If AscW(VBA.Right(C, 1)) >= &HDC00 And AscW(VBA.Right(C, 1)) <= &HDFFF Then
                    ' Check if the combined character is in CJK extension B or later
                    'If AscW(vba.Left(c, 1)) >= &HD840 Then
                    If AscW(VBA.Left(C, 1)) >= SurrogateCodePoint.HighStart Then '�e�ɥN�z (lead surrogates)�A���� D800 �� DBFF �����A�ĤG�ӳQ�٬� ����N�z (trail surrogates)�A���� DC00 �� DFFF ����
                        Dim change As Boolean
                        change = True
'                        rng.Select
                        Select Case whatCJKBlock
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_B
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_B)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_C
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_C)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_D
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_D)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_E
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_E)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_F
                                'change = isCJK_ExtF(c)
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_F)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_G
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_G)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_H
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_H)
                            Case Else
                            ' Change the font name to HanaMinB
                            ' Change the font name to fontName
                        End Select
                        If change Then rng.font.Name = fontName '"HanaMinB"
                    End If
                End If
            End If
        End If
    Next rng
    SystemSetup.contiUndo ur
End Sub
Sub ChangeFontOfSurrogatePairs_Range(fontName As String, rngtoChange As Range, Optional whatCJKBlock As CJKBlockName)
    Dim rng         As Range
    Dim C           As String
    Dim i           As Long
    Dim ur As UndoRecord
    SystemSetup.stopUndo ur, "ChangeFontOfSurrogatePairs_Range"
    For Each rng In rngtoChange.Characters
        C = rng.text
        
        Rem forDebugText
'        If c = vba.Chrw(-10122) & vba.Chrw(-8820) Or c = vba.Chrw(-10119) & vba.Chrw(-8987) Then Stop
        
        ' Check if the character is a high surrogate
        If AscW(C) >= &HD800 And AscW(C) <= &HDBFF Then
'            ' Check if the next character is a low surrogate
'            'If rng.End < ActiveDocument.Content.End Then
'            If rng.End < rngtoChange.End Then
'                i = rng.End + 1        ' The index of the next character
'                'If i < ActiveDocument.Range.End Then
'                If i < rngtoChange.End Then
'                    'c = c & ActiveDocument.Range(i, i).text        ' The combined character
'                    c = c & VBA.Mid(rngtoChange, i, 1).text        ' The combined character
'                End If
                If AscW(VBA.Right(C, 1)) >= &HDC00 And AscW(VBA.Right(C, 1)) <= &HDFFF Then
                    ' Check if the combined character is in CJK extension B or later
                    'If AscW(vba.Left(c, 1)) >= &HD840 Then
                    If AscW(VBA.Left(C, 1)) >= SurrogateCodePoint.HighStart Then '�e�ɥN�z (lead surrogates)�A���� D800 �� DBFF �����A�ĤG�ӳQ�٬� ����N�z (trail surrogates)�A���� DC00 �� DFFF ����
                        Dim change As Boolean, isCjkResult As Collection
                        change = True
'                        rng.Select
                        Select Case whatCJKBlock
                            Case CJKBlockName.CJK_Compatibility_Ideographs
                                 Set isCjkResult = IsCJK(C)
                                 If isCjkResult.item(1) Then
                                    If isCjkResult.item(2) <> CJKBlockName.CJK_Compatibility_Ideographs Then change = False
                                 End If
                            Case CJKBlockName.CJK_Compatibility_Ideographs_Supplement
                                 Set isCjkResult = IsCJK(C)
                                 If isCjkResult.item(1) Then
                                    If isCjkResult.item(2) <> CJKBlockName.CJK_Compatibility_Ideographs_Supplement Then change = False
                                 End If
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_B
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_B)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_C
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_C)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_D
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_D)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_E
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_E)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_F
                                'change = isCJK_ExtF(c)
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_F)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_G
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_G)
                            Case CJKBlockName.CJK_Unified_Ideographs_Extension_H
                                change = isCJK_Ext(C, CJK_Unified_Ideographs_Extension_H)
                            Case Else
                            ' Change the font name to HanaMinB
                            ' Change the font name to fontName
                        End Select
                        If change Then rng.font.Name = fontName '"HanaMinB"
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
        With .Replacement.font
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
    fontName = .font.Name
    fontNameFarEast = .font.NameFarEast
    ChangeCharacterFontName .text, fontName, .Document, fontNameFarEast
End With
End Sub

Rem 20230224 chatGPT�j���ĩ�Bing in Skype ����:
Sub FindMissingCharacters() '�o���ӥu�O���󤤪��r����H�s�ө���B�з������ܪ�
    Dim Doc As Document
    Set Doc = ActiveDocument
    
    '�w�q�s�ө���M�з���r�������X
    Dim nmf As font
    Set nmf = Doc.Styles("Normal").font
    Dim kff As font
    Set kff = Doc.Styles("�q��").font
    
    Dim p As Paragraph
    Dim r As Range
    Dim C As Variant
    
    ' �M�����ɤ����C�Ӭq���M�r��
    For Each p In Doc.Paragraphs
        For Each r In p.Range.Characters
            
            ' �P�_�r�ŬO�_�b�s�ө���μз���r����
            C = r.text
            If Len(C) > 0 Then
                If (AscW(VBA.Left(C, 1)) >= &H4E00 And AscW(VBA.Left(C, 1)) <= &H9FFF) _
                    Or (AscW(VBA.Left(C, 1)) >= &H3400 And AscW(VBA.Left(C, 1)) <= &H4DBF) _
                    Or (AscW(VBA.Left(C, 1)) >= &H20000 And AscW(VBA.Left(C, 1)) <= &H2A6DF) _
                    Or (AscW(VBA.Left(C, 1)) >= &H2A700 And AscW(VBA.Left(C, 1)) <= &H2B73F) _
                    Or (AscW(VBA.Left(C, 1)) >= &H2B740 And AscW(VBA.Left(C, 1)) <= &H2B81F) _
                    Or (AscW(VBA.Left(C, 1)) >= &H2B820 And AscW(VBA.Left(C, 1)) <= &H2CEAF) _
                    Or (AscW(VBA.Left(C, 1)) >= &HF900 And AscW(VBA.Left(C, 1)) <= &HFAFF) _
                    Or (AscW(VBA.Left(C, 1)) >= &H2F800 And AscW(VBA.Left(C, 1)) <= &H2FA1F) Then '�o�̨S���X�I�A���w���~�A�ݧ�g�I�I�I�I�I�I�I�I
                    If Not r.font.Name = nmf.Name And Not r.font.Name = kff.Name Then '�B�Τ���z�b����I�I�I�I
                        ' �p�G�r�Ť��b�s�ө���μз���r�����A�h�N��r���אּHanaMinB
                        r.font.Name = "HanaMinB"
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



