Attribute VB_Name = "������Ǯѹq�l�ƭp��"
Option Explicit
Sub ������_��s���ͥ���_�|���O�Z_�����w��_�h�������~�Ū���() '�m�����֡n�榡�̬ҾA�Ρ]�����~�Ū���^ 20221112
Dim rng As Range, d As Document, p As Paragraph, a As Range, i As Integer, ur As UndoRecord
Set d = ActiveDocument
If d.path <> "" Then Set d = Documents.Add
SystemSetup.stopUndo ur, "������_��s���ͥ���_�|���O�Z_�����w��_�h�������~�Ū���"
For Each p In d.Paragraphs
    For Each a In p.Range.Characters
        i = i + 1
        If i > 3 Then '���P���D�B�Y�Ƶ��e�ŴX�椧������
            If Not a.Next Is Nothing And Not a.Previous Is Nothing Then
                If a <> "�@" And a.Next = "�@" And a.Previous = "�@" Then '��r�e��ҪŮ�̤~�B�z
                    Set rng = d.Range(a.End, a.End)
                    rng.MoveEndWhile "�@"
    '                rng.Select
    '                Stop
                    rng.Delete
                End If
            End If
        End If
    Next
    i = 0
Next p
DoEvents
d.Range.Copy
DoEvents
SystemSetup.contiUndo ur
SystemSetup.playSound 2
End Sub

Rem �b�s���W�ާ@�G�Ĥ@�q���l���X�B�ĤG�q���׭��X�B�ĤT�q����ID
Sub �s����()
'the page begin
Dim start As Integer, ur As UndoRecord
' the page end
Dim e As Integer
' the book
Dim fileID As Long
'https://ctext.org/library.pl?if=gb&file=1000081&page=2621

Dim x As String ', data As New MSForms.DataObject
Dim i As Integer, rng As Range, d As Document
SystemSetup.stopUndo ur, "�s����"
Set d = ActiveDocument
Set rng = d.Range
start = CInt(Replace(rng.Paragraphs(1).Range, chr(13), ""))
e = CInt(Replace(rng.Paragraphs(2).Range, chr(13), ""))
fileID = CLng(Replace(rng.Paragraphs(3).Range, chr(13), ""))
For i = start To e
    If i = 1 Then
        x = x & "<scanbegin file=""" & fileID & """ page=""" & i & """ />��" & chr(9) & "<scanend file=""" & fileID & """ page=""" & i & """ />"
    Else
        x = x & "<scanbegin file=""" & fileID & """ page=""" & i & """ />" & chr(9) & "<scanend file=""" & fileID & """ page=""" & i & """ />" '�Y�����S�����󤺮e�A�����̫�K���ন�@�q���C�Y��n�@�Ӭq���A�|�P�U�@���H�X�b�@�_
    End If
Next i

rng.Paragraphs(3).Range = CLng(Replace(rng.Paragraphs(3).Range, chr(13), "")) + 1
'For Each e In Selection.Value
'    x = x & e
'Next e
''x = Replace(x, Chr(13), "")
'data.SetText Replace(x, "/>", "/>��", 1, 1)
'data.PutInClipboard
'SystemSetup.SetClipboard x
'SystemSetup.CopyText x
SystemSetup.SetClipboard x
If SystemSetup.GetClipboardText <> x Then
    rng.SetRange d.Range.End - 1, d.Range.End - 1
    rng.InsertAfter x
    rng.Cut
End If
rng.Document.ActiveWindow.WindowState = wdWindowStateMinimize
DoEvents
Network.AppActivateDefaultBrowser
SendKeys "^v"
SystemSetup.contiUndo ur
End Sub
Sub setPage1Code() '(ByRef d As Document)
Dim xd As String
xd = SystemSetup.GetClipboardText
If InStr(xd, "page=""1""") = 0 Then
    Dim bID As String, s As Byte, pge As String
    s = InStr(xd, "page=""")
    pge = Mid(xd, s + Len("page="""), InStr(s + Len("page="""), xd, """") - s - Len("page="""))
    If CInt(pge) < 10 Then
        s = InStr(xd, """")
        bID = Mid(xd, s + 1, InStr(s + 1, xd, """") - s - 1)
        xd = "<scanbegin file=""" & bID & """ page=""1"" />��<scanend file=""" & bID & """ page=""1"" />" + xd
        SystemSetup.ClipboardPutIn xd
    End If
End If
End Sub

Sub clearRedundantCode()
Dim xd As String, s As Long, e As Long
xd = SystemSetup.GetClipboardText
s = InStr(xd, "<scanend ") 'end �M begin�����������r
Do Until s = 0
    e = InStr(s, xd, ">")
    s = InStr(e, xd, "<scanbegin ")
    If s - e > 1 Then
        xd = Mid(xd, 1, e) + Mid(xd, s)
    End If
    s = InStr(e, xd, "<scanend ")
Loop
SystemSetup.ClipboardPutIn xd
End Sub
Sub clearRedundantText()
'�M���~�P���`��
clearWrongNoteText
End Sub
Sub clearWrongNoteText() '���Ǧ����I���_�y�������AOCR�ɤ�����r�A�D�ܻ~�N��k�İ��I�̧PŪ���r�C���h�M�����C
Dim d As Document, p As Paragraph
Set d = Documents.Add(, , , False)
d.Range.Paste
For Each p In d.Range.Paragraphs
    If InStr(p.Range, "{{") Or InStr(p.Range, "}}") Then
        p.Range.Delete
    End If
Next p
DoEvents
d.Range.Copy
d.Close wdDoNotSaveChanges
appActivateChrome
'SendKeys "+{insert}{tab}~"
SendKeys "+{insert}"

End Sub


Sub formatTitleCode() '���D�榡�]�w
Dim rng As Range, d As Document, y As Byte, s As Long, ur As UndoRecord
Set d = ActiveDocument: SystemSetup.stopUndo ur, "formatTitleCode���D�榡�]�w"
Set rng = d.Range
rng.Find.ClearFormatting
For y = 2 To 4
    Do While rng.Find.Execute("y=""" & y & """ />", , , , , , True, wdFindStop)
        GoSub code
    Loop
    Set rng = d.Range
Next y
SystemSetup.contiUndo ur
SystemSetup.playSound 1.469
Exit Sub
code:
    rng.Text = rng.Text + "*"
    s = rng.End + 1
    rng.Collapse wdCollapseStart
    rng.SetRange rng.start, rng.start
    'rng.MoveStartUntil ">"
    Do Until rng.Next.Text = "<"
        rng.move wdCharacter, -1
    Loop
    rng.move
    rng.Text = rng.Text + chr(13) + chr(13)
    rng.SetRange s, d.Range.End
    Return
End Sub

Sub �M�����e�����q�Ÿ�()
Dim d As Document, rng As Range, e As Long, s As Long, xd As String
Set d = Documents.Add
DoEvents
'If (MsgBox("add page 1 code?", vbExclamation + vbOKCancel) = vbOK) Then setPage1Code
������Ǯѹq�l�ƭp��.setPage1Code:  clearRedundantCode
�N�P���e�����q�Ÿ����m�e�q���� d
DoEvents
Set rng = d.Range
'd.ActiveWindow.Visible = True
'rng.Paste
rng.Find.ClearFormatting
Do While rng.Find.Execute("<scanbegin ") '<scanbegin file="80564" page="13" y="4" />
    rng.MoveEndUntil ">"
    rng.MoveEnd
'    rng.Select
    rng.SetRange rng.End, rng.End + 2
    If rng.Text = chr(13) & chr(13) Then
'        rng.Select
        e = rng.End
        rng.Delete
        Set rng = d.Range(e, d.Range.End)
    Else
    rng.SetRange rng.End, d.Range.End
    End If
Loop

playSound 1

Set rng = d.Range
rng.Find.ClearFormatting
Do While rng.Find.Execute("<scanend file=") ', , , , , , True, wdFindStop)
    s = rng.start
    rng.MoveEndUntil ">"
    rng.MoveEnd
'    rng.Select
    rng.SetRange rng.End, rng.End + 2
    If rng.Text = chr(13) & chr(13) Then
'        e = rng.End
'        rng.Select
        rng.Cut
        rng.SetRange s, s
        rng.Paste
        Set rng = d.Range(e, d.Range.End)
    Else
        rng.SetRange rng.End, d.Range.End
    End If
Loop


DoEvents
xd = d.Range.Text
'If d.Characters.Count < 50000 Then ' 147686
'    d.Range.Cut '��ӬOWord�� cut ��ŶKï�̦����D
'Else
    'SystemSetup.SetClipboard d.Range.Text
    SystemSetup.ClipboardPutIn xd
'End If
DoEvents
playSound 1, 0
DoEvents
pastetoEditBox "�N�P���e�����q�Ÿ����m�e�q���� & �M�����e�����q�Ÿ�"
d.Close wdDoNotSaveChanges

End Sub

Sub �N�P���e�����q�Ÿ����m�e�q����(ByRef d As Document) '20220522
Dim rng As Range, e As Long, s As Long, rngP As Range
'd As Document,Set d = Documents.Add
Set rng = d.Range
DoEvents
On Error GoTo eH
rng.Paste
rng.Find.ClearFormatting
Do While rng.Find.Execute("*")
    e = rng.End
    If rng.start > 0 Then
        If rng.Previous = chr(13) Then
            Set rng = rng.Previous
            If rng.Previous = chr(13) Then
                Set rng = rng.Previous
                If rng.Previous = ">" Then
                    rng.SetRange rng.start, e - 1
                    s = rng.start
                    Set rngP = d.Range(s, s)
                    rng.Delete
                    Do Until rngP.Next = "<"
                        If rngP.start = 0 Then GoTo NextOne
                        rngP.move wdCharacter, -1
                    Loop
                    '�ˬd�O�_���b�󭶳B 20230811
                    If d.Range(rngP.start, rngP.start + 11) = "><scanbegin" Then
                        rngP.move Count:=-1
                        Do Until rngP.Next = "<"
                            If rngP.start = 0 Then GoTo NextOne
                            rngP.move wdCharacter, -1
                        Loop
                    End If
                    '�H�W �ˬd�O�_���b�󭶳B 20230811
                    rngP.move
                    rngP.InsertAfter chr(13) & chr(13)
                End If
            End If
        End If
    End If
NextOne:
    Set rng = d.Range(e, d.Range.End)
Loop
'd.Range.Cut
'd.Close wdDoNotSaveChanges
playSound 1
'pastetoEditBox "�N�P���e�����q�Ÿ����m�e�q����"
Exit Sub
eH:
Select Case Err.Number
    Case 4605, 13 '����k���ݩʵL�k�ϥΡA�]��[�ŶKï] �O�Ū��εL�Ī��C
        SystemSetup.wait 0.8
        Resume
    Case Else
        MsgBox Err.Number + Err.Description
 End Select
End Sub

Sub �N�C���������q�Ÿ��M��()
Dim d As Document, rng As Range, s As Long, e As Long, rngCheck As Range
Const pageStart As String = "<scanbegin file="
Const pageEnd As String = "<scanend file="
Set d = ActiveDocument
Set rng = d.Range(Len(pageStart), d.Range.End)
Do While rng.Find.Execute(pageStart)
    e = rng.start: s = e - 2
    Set rngCheck = d.Range(s, e)
    rngCheck.Select
    If rngCheck.Previous = ">" Then rngCheck.Delete
    rng.SetRange rng.End + 1, d.Range.End
Loop
End Sub

Sub pastetoEditBox(Description_from_ClipBoard As String)
word.Application.WindowState = wdWindowStateMinimize
'MsgBox "ready to paste", vbInformation
AppActivateDefaultBrowser
DoEvents
'SystemSetup.Wait 0.5 '����b�o��I�_�h�j�e�q�K�W�|���ġC20220809'�ڥ��٬O�S�ΡI��ڤW�O�bWord���ŤU�ǰe��ŶKï����ƬO�Ū�
SendKeys "+{INSERT}" '"(^v)" ', True'���ȭn�h���o�Ӥ~�O�F�����O�I��ڤW���D�O�X�bWord���ŤU�ǰe��ŶKï����ƬO�Ū�
DoEvents ' DoEvents: DoEvents
Beep
SystemSetup.wait 0.3
DoEvents:
SendKeys "{tab}"
AppActivateDefaultBrowser
'SystemSetup.ClipboardPutIn Description_from_ClipBoard
DoEvents
'SendKeys "^v"
SendKeys Description_from_ClipBoard
SendKeys "{tab 2}~"
End Sub

Sub ���ۿ�_�|���O�Z_�����w��() '�m���ۿ��n�榡�̬ҾA�Ρ]�Y�`����A�Ӵ���e�������^ 20221110
Dim rng As Range, d As Document, s As Long, e As Long, rngDel As Range, ur As UndoRecord
Set d = ActiveDocument
If d.path <> "" Then Set d = Documents.Add
DoEvents
d.Range.Paste
DoEvents
Set rng = d.Range: Set rngDel = rng
rng.Find.ClearFormatting
SystemSetup.stopUndo ur, "���ۿ�_�|���O�Z_�����w��"
Do While rng.Find.Execute("}}|" & chr(13) & "{{", , , , , , True, wdFindStop)
    s = rng.start - 1: e = rng.start
    Do Until d.Range(s, e) <> "�@" '�M����e�Ů�
        s = s - 1: e = e - 1
    Loop
    rngDel.SetRange s + 1, rng.start
    'rngDel.Select
    If rngDel.Text <> "" Then If Replace(rngDel, "�@", "") = "" Then rngDel.Delete
    rng.SetRange s + Len("}}|" & chr(13) & "{{"), d.Range.End
    
    'Set rng = d.Range
Loop
d.Range.Text = Replace(Replace(d.Range.Text, "|" & chr(13) & "�@", ""), "}}|" & chr(13) & "{{", chr(13))
d.Range.Copy
SystemSetup.contiUndo ur
SystemSetup.playSound 2
word.Application.WindowState = wdWindowStateMinimize
On Error Resume Next
AppActivate "TextForCtext", True
End Sub

Sub �ন�¨��H�@��r�ƪ��קP�_��()
Dim p As Paragraph, a, i As Byte, cntr As Byte, ur As UndoRecord
If ActiveDocument.path <> "" Then Exit Sub
SystemSetup.stopUndo ur, "�ন�¨��H�@��r�ƪ��קP�_��"
Set p = Selection.Paragraphs(1)
cntr = p.Range.Characters.Count - 1
For i = 1 To cntr
    Set a = p.Range.Characters(i)
    If a.Text <> chr(13) Then a.Text = "��"
Next i
p.Range.Cut
SystemSetup.contiUndo ur
Set ur = Nothing
End Sub
Sub �M���Ҧ��Ÿ�_���q�`��Ÿ��ҥ~()
Dim f, i As Integer
f = Array("�C", "�v", chr(-24152), "�G", "�A", "�F", _
    "�B", "�u", ".", chr(34), ":", ",", ";", _
    "�K�K", "...", "�D", "�i", "�j", " ", "�m", "�n", "�q", "�r", "�H" _
    , "�I", "��", "��", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0" _
    , "�y", "�z", ChrW(9312), ChrW(9313), ChrW(9314), ChrW(9315), ChrW(9316) _
    , ChrW(9317), ChrW(9318), ChrW(9319), ChrW(9320), ChrW(9321), ChrW(9322), ChrW(9323) _
    , ChrW(9324), ChrW(9325), ChrW(9326), ChrW(9327), ChrW(9328), ChrW(9329), ChrW(9330) _
    , ChrW(9331), ChrW(8221), """") '���]�w���I�Ÿ��}�C�H�ƥ�
    '���ζ�A���Ȥ����N�I
    For i = 0 To UBound(f)
        ActiveDocument.Range.Find.Execute f(i), True, , , , , , wdFindContinue, True, "", wdReplaceAll
    Next
End Sub

Sub �M���P�ѹϪ�����_��_() '20220210
Dim rng As Range, angleRng As Range, cntr As Long
word.Application.ScreenUpdating = False
Set rng = Documents.Add().Range
Set angleRng = rng
rng.Paste
Do While rng.Find.Execute("<")
    rng.MoveEndUntil ">"
    rng.SetRange rng.start, rng.End + 1
    angleRng.SetRange rng.start, rng.End
    If InStr(angleRng.Text, "file") > 0 Then
        angleRng.Delete
    Else
        rng.SetRange rng.End, rng.Document.Range.End - 1
    End If
    If InStr(rng.Document.Range, " file=") = 0 Then Exit Do '�Y���W���ҡu<entity entityid=�v�A�h�P�_�|���~
    cntr = cntr + 1
    If cntr > 2300 Then Stop
Loop
SystemSetup.playSound 1
rng.Document.Range.Cut
rng.Document.Close wdDoNotSaveChanges
word.Application.ScreenUpdating = True
pastetoEditBox "�P�쥻�ѹϤ��X�A�Ϥ��_�C�t�̡m�����w�n�����H���Ǧۻs�n��TextForCtext�������J�C�P���P���@�n�L��������"
End Sub

Sub formatter() '���m�g������n�K��T�ǵ��榡�ΡA���i�令��L�ݭn�榡�ƪ��奻
Dim d As Document, rng As Range, a As Range, s As Long, e As Long
Const spcs As String = "�@"

Set d = Documents.Add: Set rng = d.Range
rng.Paste
For Each a In d.Characters
    If a = spcs Then
        If a.Next = spcs Then
            If InStr(a.Paragraphs(1).Range.Text, "*") = 0 Then
                a.Select
                s = a.start
                Do Until Selection.Next <> spcs
                    Selection.MoveRight , , wdExtend
                Loop
                e = Selection.End
                Set a = Selection.Next
                If a.Next = spcs Then
                    a.Select
                    Do Until Selection.Next <> spcs
                        Selection.Next.Delete
                    Loop
                
                    rng.SetRange s, e
                    'rng.Select
                    rng.Text = Replace(rng.Text, "�@", ChrW(-9217) & ChrW(-8195))
                    Set a = rng.Characters(rng.Characters.Count)
                End If
            End If
        End If
    End If
Next a
d.Range.Cut
d.Close wdDoNotSaveChanges
SystemSetup.playSound 2
End Sub

Sub formatter�~�e�[���q�Ÿ�() '���m�g������n�K��T�ǵ��榡�ΡA���i�令��L�ݭn�榡�ƪ��奻
Dim d As Document, rng As Range, a As Range, s As Long, e As Long, i As Integer, yi As Byte, ok As Boolean, yStr As String
Const y As String = "�~"
Set d = Documents.Add: Set rng = d.Range
'd.ActiveWindow.Visible = True
rng.Paste
rng.Find.ClearFormatting
Do While rng.Find.Execute("^p")
    If rng.End = d.Range.End - 1 Then Exit Do
    Set a = d.Range
    For i = 4 To 2 Step -1
        a.SetRange rng.End, rng.End + i
'        a.Select
        If right(a, 1) = y Then
            If a.Previous.Previous <> ">" Then
                For yi = 1 To 99
                    yStr = ��r�ഫ.�Ʀr��~�r2���(yi) + y
                    If a.Text = yStr Then
                        rng.InsertBefore "<p>"
                        ok = True: Exit For
                    End If
                Next yi
                If ok Then
                    ok = False
                    Exit For
                End If
            End If
        End If
    Next i
    rng.SetRange rng.End, d.Range.End
Loop
SystemSetup.playSound 2
d.Range.Cut
d.Close wdDoNotSaveChanges
End Sub
Sub �����w�|���O�Z�����()
Dim d As Document, a, i, p As Paragraph, xP As String, acP As Integer, space As String, rng As Range
On Error GoTo eH
a = Array(ChrW(12296), "{{", ChrW(12297), "}}", "�q", "{{", "�r", "}}", _
    "��", ChrW(12295))
'�m�e�N�T���n���p�`�@����٪����� https://ctext.org/library.pl?if=gb&file=89545&page=24
'a = Array("�q", "", "�r", "", _
    "��", ChrW(12295))


Set d = Documents.Add()
d.Range.Paste
'���ܶK�W�Lê
SystemSetup.playSound 1
�����w�y�r�Ϩ��N����r d.Range
For i = 0 To UBound(a) - 1
    d.Range.Find.Execute a(i), , , , , , True, wdFindContinue, , a(i + 1), wdReplaceAll
    i = i + 1
Next i
For Each p In d.Range.Paragraphs
    xP = p.Range
    If left(xP, 2) = "{{" And right(xP, 3) = "}}" & chr(13) Then
        xP = Mid(p.Range, 3, Len(xP) - 5)
        If InStr(xP, "{{") = 0 And InStr(xP, "}}") = 0 Then
            acP = p.Range.Characters.Count - 1
            If acP Mod 2 = 0 Then
                acP = CInt(acP / 2)
            Else
                acP = CInt((acP + 1) / 2)
            End If
            If p.Range.Characters(acP).InlineShapes.Count = 0 Then
                p.Range.Characters(acP).InsertParagraphAfter
            Else
                p.Range.Characters(acP).Select
                Selection.Delete
                Selection.TypeText " "
                p.Range.Characters(acP).InsertParagraphAfter
            End If
        End If
    ElseIf left(xP, 1) = "�@" Then '�e���Ů檺
        i = InStr(xP, "{{")
        If i > 0 And right(xP, 3) = "}}" & chr(13) Then
            space = Mid(xP, 1, i - 1)
            If Replace(space, "�@", "") = "" Then
                xP = Mid(xP, i + 2, Len(xP) - 3 - (i + 2))
                If InStr(xP, "{{") = 0 And InStr(xP, "}}") = 0 Then
                    Set rng = p.Range
                    rng.SetRange rng.Characters(1).start, rng.Characters(i + 1).End
                    rng.Text = "{{" & space
                    acP = p.Range.Characters.Count - 1 - Len(space)
                    If acP Mod 2 = 0 Then
                        acP = CInt(acP / 2) + Len(space) + 1
                    Else
                        acP = CInt((acP + 1) / 2) + Len(space) + 1
                    End If
                    If p.Range.Characters(acP).InlineShapes.Count = 0 Then
                        p.Range.Characters(acP).InsertBefore chr(13) & space
                    Else
                        p.Range.Characters(acP).Select
                        Selection.Delete
                        Selection.TypeText " "
                        p.Range.Characters(acP).InsertBefore chr(13) & space
                    End If
                    
                End If
            End If
        End If
    End If
Next p
�����w���������⴫���r d
��r�B�z.�ѦW���g�W���Ъ`
d.Range.Cut
d.Close wdDoNotSaveChanges
SystemSetup.playSound 2
Exit Sub
eH:
Select Case Err.Number
    Case 5904 '�L�k�s�� [�d��]�C
        If p.Range.Characters(acP).Hyperlinks.Count > 0 Then p.Range.Characters(acP).Hyperlinks(1).Delete
        Resume
    Case Else
        MsgBox Err.Number & Err.Description
End Select
End Sub

Sub �����w�|���O�Z�����_early()
Dim d As Document, a, i

a = Array("^p^p", "@", "�q", "{{", "�r", "}}", "^p", "", "}}{{", "^p", "@", "^p", _
    "��", ChrW(12295))
Set d = Documents.Add()
d.Range.Paste
�����w�y�r�Ϩ��N����r d.Range
For i = 0 To UBound(a) - 1
    d.Range.Find.Execute a(i), , , , , , True, wdFindContinue, , a(i + 1), wdReplaceAll
    i = i + 1
Next i
��r�B�z.�ѦW���g�W���Ъ`
d.Range.Cut
d.Close wdDoNotSaveChanges
Beep
End Sub

Sub searchuCtext()
' Alt+,
SystemSetup.playSound 0.484
Select Case Selection.Text
    Case "", chr(13), chr(9), chr(7), chr(10), " ", "�@"
        MsgBox "no selected text for search !", vbCritical: Exit Sub
End Select
Static bookID
Dim searchedTerm, e, addressHyper As String, bID As String, cndn As String
'Const site As String = "https://ctext.org/wiki.pl?if=gb&res="
Const site As String = "https://ctext.org/wiki.pl?if=gb"
bID = left(ActiveDocument.Paragraphs(1).Range, Len(ActiveDocument.Paragraphs(1).Range) - 1)
If Not VBA.IsNumeric(bID) Then
    If InStr(bID, site) = 0 Then
        bookID = InputBox("plz input the book id ", , bookID)
    Else
        bookID = bID
    End If
Else
    bookID = bID
End If
If InStr(bookID, "https") > 0 Then
    If InStr(bookID, "&res=") = 0 And InStr(bookID, "&chapter=") = 0 Then MsgBox "error . not the proper bookID ref ", vbCritical: Exit Sub
    If InStr(bookID, "&res=") > 0 Then
        cndn = "&res="
    ElseIf InStr(bookID, "&chapter=") > 0 Then
        cndn = "&chapter="
    Else
        MsgBox "error . not the proper bookID ref ", vbCritical: Exit Sub
    End If
    bookID = Mid(bookID, InStr(bookID, cndn) + Len(cndn))
    If Not VBA.IsNumeric(bookID) Then
        bookID = Mid(bookID, 0, InStr(bookID, "&searchu"))
    End If
End If
If Not VBA.IsNumeric(bookID) Then
    MsgBox "error . not the proper bookID ref ", vbCritical: Exit Sub
End If
e = Selection.Text
'searchedTerm = 'Array("��", "��", "�P��", "���g", "�t��", "ô��", "����", "����", "�Ǩ�", "����", "�Ԩ�", "����", "�娥", "���[", "�L�S", ChrW(26080) & "�S", "�ѩS", "����", "�Q�s", "��") ', "", "", "", "")
''https://ctext.org/wiki.pl?if=gb&res=757381&searchu=%E5%8D%A6
'For Each e In searchedTerm
    addressHyper = addressHyper + " " + site + cndn + bookID + "&searchu=" + e
'Next e
Shell Network.getDefaultBrowserFullname + addressHyper

Selection.Hyperlinks.Add Selection.Range, addressHyper
End Sub

Sub �v�O�T�a�`()
'�q2858���_�A20210920:0817����A��λO�v�j�����P�ǧd��@���͡m���ؤ�ƺ��n�ҿ�����|�m�v��n�쥻�A���Τ�����A�M�ܤ֧K��²�Ʀr�ഫ�_�~�γy�r�ýX���x�Z�A���r�ɱ�m�C�ھڪ�@���A�榡�����@�ˡI�ڥ��N�O�q�o�̥X�Ӫ��A�A��²�Ʀr�A�A�S�ϥ��A�y�������áC�����S�Q��Φ����]�C��������C��̤l�]�u�u���u�j���ѩ�2021�~9��20��
Dim d As Document, a, i, p As Paragraph, px As String, rng As Range, e As Long, pRng As Range, pa
'Const corTxt As String = "�׸��I�ե��հɰO��"'�Ӻ����Ϥ��ӱƪ��\�ॼ��t�X�A�G�����ĥΡC��榡�u��奻�����ġChttps://ctext.org/instructions/wiki-formatting/zh
'a = Array(" ", "", "�@�@","","�@", ChrW(-9217) & ChrW(-8195), "^p", "<p>^p",
'a = Array(" ", "", "�@�@", "", "^p^p", "<p>^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195),
a = Array("�@�@", "", "^p", "^p^p", "^p^p^p", "^p^p", "�u^p^p", "�u", "�y^p^p", "�y", "�e^p^p", "�e", "�]^p^p", "�]", _
    "^p^p", "<p>^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), _
    "^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195) & "�e", _
    "^p{{" & ChrW(-9217) & ChrW(-8195) & "{{{�q", _
    "�u<p>^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), "�u", _
    "�e<p>^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), "�e", _
    "�y<p>^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), "�y", _
    "�]<p>^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), "�]", _
    "����", "�m���ѡn�G", "����", "�m�����n�G", "�i�m�����n�G�z�١j", "�i�m�����n�z�١j�G", "���q", "�m���q�n�G", _
    "�E�{�q", "�E�{", "���]", "���C", "�]��", "�C��", "�w����", "�w���", _
    "��", "��", "�F", ChrW(24921), "��", ChrW(21843), _
     ChrW(-30641), ChrW(-25066), _
     "�s", ChrW(32675), "�Y", ChrW(21373), "��", ChrW(-30650), _
     "�J", ChrW(26083), "��", ChrW(27114), "�@", ChrW(28433), _
     "��", ChrW(-30626), _
     "�u", ChrW(30494), "��", ChrW(22625), "�M", ChrW(28152), "�C", ChrW(-26799), "��", ChrW(25934), _
    "�m", ChrW(-28395), "��", ChrW(-27731), "�V", ChrW(24892), _
    "�}", ChrW(24183), "��", ChrW(23643), "��", ChrW(-31930), "��", ChrW(-28471), "�@", ChrW(31571), _
    "�p", ChrW(29314), "��", ChrW(-25811), "��", ChrW(32220), _
    "�T", ChrW(20868), "�}", ChrW(-32486), _
    ChrW(25995), ChrW(-24956))
Set d = Documents.Add()
d.Range.Paste
�����w�y�r�Ϩ��N����r d.Range
d.Range.Cut
d.Range.PasteAndFormat wdFormatPlainText
d.Range.Text = VBA.Replace(d.Range.Text, " ", "")
For i = 0 To UBound(a) - 1
    If a(i) = "^p^p^p" Then
        px = d.Range.Text
        Do While InStr(px, chr(13) & chr(13) & chr(13))
            px = Replace(px, chr(13) & chr(13) & chr(13), chr(13) & chr(13))
        Loop
        d.Range.Text = px
        'Set rng = d.Range
'        Do While rng.Find.Execute(a(i), , , , , , True, wdFindContinue, , a(i + 1), wdReplaceAll)
'            If rng.End = d.Range.End Then Exit Do
'        Loop
    Else
        d.Range.Find.Execute a(i), , , , , , True, wdFindContinue, , a(i + 1), wdReplaceAll
    End If
    i = i + 1
Next i
��r�B�z.�ѦW���g�W���Ъ`
Set rng = Selection.Range
For Each p In d.Paragraphs
    px = p.Range.Text
    If left(px, 7) = "{{" & ChrW(-9217) & ChrW(-8195) & "{{{" Then '�`�}�q��
        e = p.Range.Characters(1).End
        rng.SetRange e, e
        rng.MoveEndUntil "�f"
        If rng.Next.Next = "�@" Then rng.Next.Next.Delete
        If InStr(p.Range.Text, "�@") Then
            For Each pa In p.Range.Characters
                If pa = "�@" Then
                    pa.Text = ChrW(-9217) & ChrW(-8195)
                End If
            Next
'            p.Range.text = VBA.Replace(p.Range.text, "�@", ChrW(-9217) & ChrW(-8195))
'            'replace the text of paragraph the paragraph will be move to next one
'            Set p = p.Previous
'            e = p.Range.Characters(1).End
'            rng.SetRange e, e
'            rng.MoveEndUntil "�f"
        End If
        'rng.Select
        rng.Collapse wdCollapseEnd
        rng.Select
        Selection.MoveRight wdCharacter, 1, wdExtend
        Selection.TypeText "�r}}}" '�N�`�}�s���e�@�f���k��f�令}}}
        px = p.Range.Text
        If InStr(right(px, 4), "<p>") Then
            e = p.Range.Characters(p.Range.Characters.Count - 4).End
        Else
            e = p.Range.Characters(p.Range.Characters.Count - 1).End
        End If
        rng.SetRange e, e
        rng.InsertAfter "}}"
    Else '����q��
        e = p.Range.Characters(1).start
        Set pRng = p.Range
        Do While InStr(pRng.Text, "�e")
            rng.SetRange e, e
            rng.MoveEndUntil "�e"
            If rng.Characters(rng.Characters.Count) <> "�^" Then  ' if not correction
                rng.Collapse wdCollapseEnd
                rng.move , 1
                rng.MoveEnd wdCharacter, 1
                If rng.Text Like "[�@�G�T�|�����C�K�E]" Then  ' is footnote No.
                    e = rng.start
                    'rng.Collapse wdCollapseEnd
                    rng.SetRange e - 1, e
                    rng.Text = "�@{{{�q"
                    rng.MoveEndUntil "�f"
                    rng.Collapse wdCollapseEnd
                    rng.MoveEnd wdCharacter, 1
                    rng.Text = "�r}}}"
                Else 'is correction to insert words
'                    rng.MoveEndUntil "�f"
'                    rng.SetRange rng.End + 2, rng.End + 2
'                    rng.InsertAfter corTxt
                End If
                e = rng.End
            Else 'is correction
'                If rng.Characters(rng.Characters.Count).Next = "�e" Then ' delete and insert words
'                    rng.MoveEndUntil "�f"
'                    rng.SetRange rng.End + 2, rng.End + 2
'                End If
'                rng.InsertAfter corTxt
               e = rng.End + 1
            End If
            'e = rng.End
            pRng.SetRange e, p.Range.End
            'pRng.SetRange rng.End, p.Range.End
            
        Loop
    End If
    If VBA.left(p.Range.Text, 9) = ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195) & "�i�m�����n" Then
        Set rng = p.Range
        p.Range.Characters(1).Delete
        rng.SetRange p.Range.start, p.Range.start
        rng.InsertAfter "{{"
        rng.SetRange p.Range.Characters(p.Range.Characters.Count - 4).End, p.Range.Characters(p.Range.Characters.Count - 4).End
        rng.InsertAfter "}}"
        If Len(rng.Paragraphs(1).Next.Range.Text) = 1 Then rng.Paragraphs(1).Next.Range.Delete
    End If
    
    If Len(p.Range) < 20 Then
        If (InStr(p.Range, "�m�v�O�n��") Or VBA.left(p.Range.Text, 3) = "�v�O��") And InStr(p.Range, "*") = 0 Then
            rng.SetRange p.Range.start, p.Range.start
            rng.InsertAfter "*"
            For Each pa In p.Range.Characters
                    If pa Like "[�q�m�n�r]" Or StrComp(pa, ChrW(-9217) & ChrW(-8195)) = 0 Then pa.Delete
            Next pa
            '�H�U�覡�|�y��p �ȳQ�]�w���U�@�Ӭq��
'            p.Range.text = VBA.Replace(p.Range.text, ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), "")
'            p.Range.text = VBA.Replace(VBA.Replace(p.Range.text, "�m", ""), "�n", "")
        End If
    End If
    If Len(p.Range) < 25 Then
        If VBA.InStr(p.Range.Text, "��") And InStr(p.Range, "*") = 0 _
                And (InStr(p.Range, "����") Or InStr(p.Range, "��") Or InStr(p.Range, "��") _
                Or InStr(p.Range, "�@�a") Or InStr(p.Range, "�C��")) Then
            rng.SetRange p.Range.start, p.Range.start
            rng.InsertAfter "�@*"
            For Each pa In p.Range.Characters
                If pa Like "[�q�m�n�r]" Or StrComp(pa, ChrW(-9217) & ChrW(-8195)) = 0 Then pa.Delete
            Next pa
   
'            p.Range.text = VBA.Replace(p.Range.text, ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), "�@*")
'            p.Range.text = VBA.Replace(VBA.Replace(p.Range.text, "�q", ""), "�r", "")
        End If
    End If

Next p
If VBA.left(d.Paragraphs(1).Range.Text, 3) = "�v�O��" And InStr(d.Paragraphs(1).Range.Text, "*") = 0 Then
    Set p = d.Paragraphs(1)
    rng.SetRange p.Range.start, p.Range.start
    rng.InsertAfter "*"
'    rng.SetRange p.Range.Characters(p.Range.Characters.Count - 1).End, p.Range.Characters(p.Range.Characters.Count - 1).End
'    rng.InsertAfter "<p>"
End If
If VBA.InStr(d.Paragraphs(2).Range.Text, "��") And InStr(d.Paragraphs(2).Range.Text, "*") = 0 Then
    Set p = d.Paragraphs(2)
'    rng.SetRange p.Range.start, p.Range.start
'    rng.InsertAfter "�@*"
''    rng.SetRange p.Range.Characters(p.Range.Characters.Count - 1).End, p.Range.Characters(p.Range.Characters.Count - 1).End
''    rng.InsertAfter "<p>"
    p.Range.Text = VBA.Replace(p.Range.Text, ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), "�@*")
    Set p = d.Paragraphs(2)
    p.Range.Text = VBA.Replace(VBA.Replace(p.Range.Text, "�q", ""), "�r", "")
End If

'Set rng = d.Range
'Do While rng.Find.Execute("�f", , , , , , True, wdFindStop)
'    If rng.Characters(1).Next <> "��" Then rng.InsertAfter corTxt
'Loop
'Set rng = d.Range
'Do While rng.Find.Execute("�^", , , , , , True, wdFindStop)
'    If InStr("�סe", rng.Characters(1).Next) = 0 Then rng.InsertAfter corTxt
'Loop
d.Range.Cut
d.Close wdDoNotSaveChanges
Beep
word.Application.ActiveWindow.WindowState = wdWindowStateMinimize
End Sub
Sub �v�O�T�a�`2old()
'�q2858���_�A20210920:0817����A��λO�v�j�����P�ǧd��@���͡m���ؤ�ƺ��n�ҿ�����|�m�v��n�쥻�A���Τ�����A�M�ܤ֧K��²�Ʀr�ഫ�_�~�γy�r�ýX���x�Z�A���r�ɱ�m�C�ھڪ�@���A�榡�����@�ˡI�ڥ��N�O�q�o�̥X�Ӫ��A�A��²�Ʀr�A�A�S�ϥ��A�y�������áC�����S�Q��Φ����]�C��������C��̤l�]�u�u���u�j���ѩ�2021�~9��20��
Dim d As Document, a, i, p As Paragraph, px As String, rng As Range, e As Long, pRng As Range
'Const corTxt As String = "�׸��I�ե��հɰO��"'�Ӻ����Ϥ��ӱƪ��\�ॼ��t�X�A�G�����ĥΡC��榡�u��奻�����ġChttps://ctext.org/instructions/wiki-formatting/zh
a = Array("^p^p", "<p>^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), _
    "^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195) & "�e", _
    "^p{{" & ChrW(-9217) & ChrW(-8195) & "{{{�q", _
    "�u<p>^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), "�u", _
    "�e<p>^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), "�e", _
    "�y<p>^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), "�y", _
    "�]<p>^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), "�]", _
    "����", "�m���ѡn�G", "����", "�m�����n�G", "�i�m�����n�G�z�١j", "�i�m�����n�z�١j�G", "���q", "�m���q�n�G", _
    "�E�{�q", "�E�{", "���]", "���C", "�]��", "�C��", "�w����", "�w���", _
    "��", "��", _
     "�s", ChrW(32675), "�Y", ChrW(21373), "��", ChrW(-30650), "�J", ChrW(26083), "��", ChrW(-30626), _
     "�u", ChrW(30494), "��", ChrW(22625), "�M", ChrW(28152), "�C", ChrW(-26799), "��", ChrW(25934), _
    "�m", ChrW(-28395), "��", ChrW(-27731), "�V", ChrW(24892), "��", ChrW(23643), "��", ChrW(27114), _
    "��", ChrW(-31930), "��", ChrW(-28471))
Set d = Documents.Add()
d.Range.Paste
For i = 0 To UBound(a) - 1
    d.Range.Find.Execute a(i), , , , , , True, wdFindContinue, , a(i + 1), wdReplaceAll
    i = i + 1
Next i
��r�B�z.�ѦW���g�W���Ъ`
Set rng = Selection.Range
For Each p In d.Paragraphs
    px = p.Range.Text
    If left(px, 7) = "{{" & ChrW(-9217) & ChrW(-8195) & "{{{" Then '�`�}�q��
        e = p.Range.Characters(1).End
        rng.SetRange e, e
        rng.MoveEndUntil "�f"
        'rng.Select
        rng.Collapse wdCollapseEnd
        rng.Select
        Selection.MoveRight wdCharacter, 1, wdExtend
        Selection.TypeText "�r}}}" '�N�`�}�s���e�@�f���k��f�令}}}
        px = p.Range.Text
        If InStr(right(px, 4), "<p>") Then
            e = p.Range.Characters(p.Range.Characters.Count - 4).End
        Else
            e = p.Range.Characters(p.Range.Characters.Count - 1).End
        End If
        rng.SetRange e, e
        rng.InsertAfter "}}"
    Else '����q��
        e = p.Range.Characters(1).start
        Set pRng = p.Range
        Do While InStr(pRng.Text, "�e")
            rng.SetRange e, e
            rng.MoveEndUntil "�e"
            If rng.Characters(rng.Characters.Count) <> "�^" Then  ' if not correction
                rng.Collapse wdCollapseEnd
                rng.move , 1
                rng.MoveEnd wdCharacter, 1
                If rng.Text Like "[�@�G�T�|�����C�K�E]" Then  ' is footnote No.
                    e = rng.start
                    'rng.Collapse wdCollapseEnd
                    rng.SetRange e - 1, e
                    rng.Text = "�@{{{�q"
                    rng.MoveEndUntil "�f"
                    rng.Collapse wdCollapseEnd
                    rng.MoveEnd wdCharacter, 1
                    rng.Text = "�r}}}"
                Else 'is correction to insert words
'                    rng.MoveEndUntil "�f"
'                    rng.SetRange rng.End + 2, rng.End + 2
'                    rng.InsertAfter corTxt
                End If
                e = rng.End
            Else 'is correction
'                If rng.Characters(rng.Characters.Count).Next = "�e" Then ' delete and insert words
'                    rng.MoveEndUntil "�f"
'                    rng.SetRange rng.End + 2, rng.End + 2
'                End If
'                rng.InsertAfter corTxt
               e = rng.End + 1
            End If
            'e = rng.End
            pRng.SetRange e, p.Range.End
            'pRng.SetRange rng.End, p.Range.End
            
        Loop
    End If
    If VBA.left(p.Range.Text, 9) = ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195) & "�i�m�����n" Then
        Set rng = p.Range
        p.Range.Characters(1).Delete
        rng.SetRange p.Range.start, p.Range.start
        rng.InsertAfter "{{"
        rng.SetRange p.Range.Characters(p.Range.Characters.Count).End, p.Range.Characters(p.Range.Characters.Count).End
        rng.InsertAfter "}}"
    End If
Next p
If VBA.left(d.Paragraphs(1).Range.Text, 3) = "�v�O��" Then
    Set p = d.Paragraphs(1)
    rng.SetRange p.Range.start, p.Range.start
    rng.InsertAfter "*"
    rng.SetRange p.Range.Characters(p.Range.Characters.Count - 1).End, p.Range.Characters(p.Range.Characters.Count - 1).End
    rng.InsertAfter "<p>"
End If
If VBA.InStr(d.Paragraphs(2).Range.Text, "��") Then
    Set p = d.Paragraphs(2)
    rng.SetRange p.Range.start, p.Range.start
    rng.InsertAfter "�@*"
'    rng.SetRange p.Range.Characters(p.Range.Characters.Count - 1).End, p.Range.Characters(p.Range.Characters.Count - 1).End
'    rng.InsertAfter "<p>"
    p.Range.Text = VBA.Replace(VBA.Replace(p.Range.Text, "�q", ""), "�r", "")
End If


'Set rng = d.Range
'Do While rng.Find.Execute("�f", , , , , , True, wdFindStop)
'    If rng.Characters(1).Next <> "��" Then rng.InsertAfter corTxt
'Loop
'Set rng = d.Range
'Do While rng.Find.Execute("�^", , , , , , True, wdFindStop)
'    If InStr("�סe", rng.Characters(1).Next) = 0 Then rng.InsertAfter corTxt
'Loop
d.Range.Cut
d.Close wdDoNotSaveChanges
Beep
End Sub

Sub �v�O�T�a�`1old()
Dim d As Document, a, i, p As Paragraph, px As String, rng As Range, e As Long
a = Array("<p>{{{", "<p>^p{{" & ChrW(-9217) & ChrW(-8195) & "{{{", _
        "<p>", "<p>^p" & ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195), _
        ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195) & "^p{{" & ChrW(-9217) & ChrW(-8195), _
        "{{" & ChrW(-9217) & ChrW(-8195))
Set d = Documents.Add()
d.Range.Paste
For i = 0 To UBound(a) - 1
    d.Range.Find.Execute a(i), , , , , , True, wdFindContinue, , a(i + 1), wdReplaceAll
    i = i + 1
Next i
��r�B�z.�ѦW���g�W���Ъ`
d.Range.Find.Execute "�m�m", , , , , , True, wdFindContinue, , "�m", wdReplaceAll
d.Range.Find.Execute "�n�n", , , , , , True, wdFindContinue, , "�n", wdReplaceAll
d.Range.Find.Execute "�q�q", , , , , , True, wdFindContinue, , "�q", wdReplaceAll
d.Range.Find.Execute "�r�r", , , , , , True, wdFindContinue, , "�r", wdReplaceAll
Set rng = Selection.Range
For Each p In d.Paragraphs
    px = p.Range.Text
    If left(p.Range.Text, 7) = "{{" & ChrW(-9217) & ChrW(-8195) & "{{{" Then
        If InStr(right(px, 4), "<p>") Then
            e = p.Range.Characters(p.Range.Characters.Count - 4).End
        Else
            e = p.Range.Characters(p.Range.Characters.Count - 1).End
        End If
        rng.SetRange e, e
        rng.InsertAfter "}}"
    End If
    
Next p

d.Range.Cut
d.Close wdDoNotSaveChanges
End Sub
Sub ��sub()
Dim p As Paragraph, d As Document, rng As Range, s As Long, e As Long
Set d = Documents.Add(): Set rng = d.Range
d.Range.Paste
For Each p In d.Paragraphs
    If InStr(p.Range, "�m�����n�G") Or _
        InStr(p.Range, "�m���q�n�G") Or _
        InStr(p.Range, "�m���ѡn�G") Then
        If InStr(p.Range, "{{") = 0 Then
            s = p.Range.Characters(1).start
            rng.SetRange s, s
            rng.InsertBefore "{{"
            e = p.Range.Characters(p.Range.Characters.Count - 4).End
            rng.SetRange e, e
            rng.InsertAfter "}}"
        End If
    End If
Next p
d.Range.Cut
d.Close wdDoNotSaveChanges
Beep
End Sub

Sub ��sub1()
Dim d As Document, rng As Range, rngLast As Range, s As Long, e As Long
Set d = ActiveDocument
Set rng = d.Range: Set rngLast = rng
With rng.Find
    .Font.Color = 10092543
    .Font.Size = 10
    .Forward = True
    Do
        .Execute , , , , , , , wdFindStop
        If InStr(rng, "}}") Then
            .Execute , , , , , , , wdFindStop
            If InStr(rng, "}}") Then Exit Do
        End If
        s = rng.Characters(1).start
        e = rng.Characters(rng.Characters.Count - 1).End
        rngLast.SetRange e - 1, e
        rngLast.InsertAfter "}}"
        rngLast.SetRange s, s
        rngLast.InsertBefore "{{" & ChrW(-9217) & ChrW(-8195)
'        rng.SetRange rng.End + 222, d.Range.End
        
    Loop 'Until InStr(rng, "{{")
    .ClearFormatting
End With
Beep
End Sub

Rem �^�Ǻ��}
Function Search(searchWhatsUrl As String) As String
    Dim d As Document
    Set d = ActiveDocument
    If d.path <> "" Then If d.Saved = False Then d.Save
    If Selection.Type = wdSelectionNormal Then
        Selection.Copy
    End If
    'Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe https://ctext.org/wiki.pl?if=gb&res=384378&searchu=" & Selection.text
    'Shell Normal.SystemSetup.getChrome & searchWhatsUrl & Selection.Text
    Shell Normal.Network.GetDefaultBrowserEXE & searchWhatsUrl & Selection.Text
    Search = searchWhatsUrl & Selection.Text
End Function

Sub search�v�O�T�a�`()
    ActiveDocument.Hyperlinks.Add Selection.Range, Search(" https://ctext.org/wiki.pl?if=gb&res=384378&searchu=")
End Sub

Sub search�P�����q_�����Q�T�g�`��()
    ActiveDocument.Hyperlinks.Add Selection.Range, Search(" https://ctext.org/wiki.pl?if=gb&res=315747&searchu=")
End Sub


Sub Ū�v�O�T�a�`()
Dim d As Document, t As Table
Set d = Documents.Add
d.Range.Paste
Set t = d.Tables(1)
With t
    .Columns(1).Delete
    .ConvertToText wdSeparateByParagraphs
End With
d.Range.Cut
d.Close wdDoNotSaveChanges
If word.Application.Windows.Count > 0 Then word.Application.ActiveWindow.WindowState = wdWindowStateMinimize
End Sub

Sub �԰굦_�|���O�Z_�����w��() '�m�԰굦�n�榡�̬ҾA�Ρ]�Y�D�孺�泻��A�Ө�l���e���@��̡^
'https://ctext.org/library.pl?if=gb&res=77385
Dim a, rng As Range, rngDoc As Range, p As Paragraph, i As Long, rngCnt As Integer, ok As Boolean
Dim omits As String
omits = "�m�n�q�r�u�v�y�z�P" & chr(13)
Set rngDoc = Documents.Add.Range
re:
rngDoc.Paste
�����w�y�r�Ϩ��N����r rngDoc
For Each p In rngDoc.Paragraphs
    Set a = p.Range.Characters(1)
    If a <> "�@" Then a.InsertBefore "�@"
Next p
For Each a In rngDoc.Characters
    If Not a.Next Is Nothing And Not a.Previous Is Nothing Then
        If a = "�@" And a.Next <> "�@" And a.Previous <> "�@" Then
            If a.Previous <> chr(13) Then a.InsertBefore chr(13)
            Set a = a.Next
        End If
    End If
Next a

For Each p In rngDoc.Paragraphs
    Set rng = p.Range
    If StrComp(rng.Characters(1), "�@") = 0 And InStr(rng, "}") > 0 Then
        If rng.Characters(1) = "�@" And rng.Characters(2) = "{" And rng.Characters(3) = "{" Then
            rng.Characters(1) = "{": rng.Characters(2) = "{": rng.Characters(3) = "�@"
            For Each a In rng.Characters
               i = i + 1
               If rng.Characters(i) = "}" Then Exit For
               If rng.Characters(i) = chr(13) Then
                    i = 0
                    Exit For
               End If
            Next a
        Else
            For Each a In rng.Characters
               i = i + 1
               If rng.Characters(i) = "}" Then Exit For
               If rng.Characters(i) = chr(13) Or rng.Characters(i) = "{" Then
                    i = 0
                    Exit For
               End If
            Next a
        End If
        If i <> 0 Then
            If rng.Characters(1) = "{" And rng.Characters(2) = "{" And rng.Characters(3) = "�@" Then
                rng.SetRange rng.Characters(3).End, rng.Characters(i).start
            Else
                rng.SetRange rng.Characters(1).End, rng.Characters(i).start
            End If
'            rng.Select
'            Stop
            rngCnt = rng.Characters.Count
            If rngCnt > 1 Then
                i = 0
                For Each a In rng.Characters
                    If InStr(omits, a) = 0 Then i = i + 1
                Next a
                rngCnt = i: i = 0
                If rngCnt Mod 2 = 1 Then
                    rngCnt = (rngCnt + 1) / 2
                Else
                    rngCnt = rngCnt / 2
                End If
                For Each a In rng.Characters
                    If InStr(omits, a) = 0 Then i = i + 1
                    If i = rngCnt Then
                        a.InsertAfter "�@"
                        Exit For
                    End If
                Next a
'                If rngCnt Mod 2 = 1 Then
'                    If rng.Characters((rngCnt - rngCnt Mod 2) / 2 + 1).Next <> "�@" _
'                        Then rng.Characters((rngCnt - rngCnt Mod 2) / 2 + 1).InsertAfter "�@"
'
'                Else
'                    If rng.Characters((rngCnt - rngCnt Mod 2) / 2).Next <> "�@" _
'                        Then rng.Characters((rngCnt - rngCnt Mod 2) / 2).InsertAfter "�@"
'                End If
            Else
                rng.Characters(1).InsertAfter "�@"
            End If
        End If
        i = 0
    End If
Next
If ok Then
    For Each p In rngDoc.Paragraphs
        If left(p.Range.Text, 3) = "{{�@" And p.Range.Characters(p.Range.Characters.Count - 1) = "}" Then
            a = p.Range.Text
            a = Mid(a, 4, Len(a) - 6)
            If InStr(a, "�@") > 0 And InStr(a, "{") = 0 And InStr(a, "}") = 0 Then
                rngCnt = p.Range.Characters.Count
                For i = 4 To rngCnt
                    Set a = p.Range.Characters(i)
                    If a = "�@" Then
                        a.InsertParagraphBefore
                        Exit For
                    End If
                Next i
            End If
        End If
    Next p
    '�H�U3��m�԰굦�n�����~�ݭn
'    rngDoc.Find.Execute "����", , , , , , , wdFindContinue, , "�i����j", wdReplaceAll
'    rngDoc.Find.Execute ChrW(-10155) & ChrW(-8585) & "��", , , , , , , wdFindContinue, , "�i" & ChrW(-10155) & ChrW(-8585) & "��j", wdReplaceAll
'    rngDoc.Find.Execute "�ɤ�", , , , , , , wdFindContinue, , "�i" & ChrW(-10155) & ChrW(-8585) & "��j", wdReplaceAll
End If
If ok Then ��r�B�z.�ѦW���g�W���Ъ`
rngDoc.Cut
If Not ok Then
    DoEvents
    rngDoc.PasteAndFormat wdFormatPlainText
    rngDoc.Find.Execute "�q", , , , , , , wdFindContinue, , "{{", wdReplaceAll
    rngDoc.Find.Execute "�r", , , , , , , wdFindContinue, , "}}", wdReplaceAll
    rngDoc.Cut
    ok = True
    GoTo re
End If
rngDoc.Document.Close wdDoNotSaveChanges
On Error Resume Next
AppActivate "TextForCtext"
SendKeys "%{insert}", True
SystemSetup.playSound 4
End Sub
Sub ��������Y�Ƥ@������p�`�榡_�|�w����_��Ǥj�v()
    Dim d As Document, p As Paragraph, px As String, rng As Range, a As Range, ur As UndoRecord
    Set d = ActiveDocument: Set rng = d.Range
    SystemSetup.stopUndo ur, "��������Y�Ƥ@������p�`�榡_�|�w����_��Ǥj�v"
    For Each p In d.Paragraphs
        px = p.Range.Text
        If VBA.left(px, 3) = "�@{{" And VBA.right(px, 3) = "}}" & chr(13) Then
            rng.SetRange p.Range.start + 3, p.Range.End - 3
            rng.Characters(Int(rng.Characters.Count / 2)).InsertAfter chr(13) & "�@"
        ElseIf VBA.left(px, 3) = "{{�@" And VBA.right(px, 6) = "}}<p>" & chr(13) Then
            rng.SetRange p.Range.start + 3, p.Range.End - 6
            For Each a In rng.Characters
                If a.Text = "�@" Then
                    a.InsertBefore chr(13)
                End If
            Next a
        End If
    Next p
    SystemSetup.contiUndo ur
End Sub


Sub �����w�y�r�Ϩ��N����r(rng As Range)
Dim inlnsp As InlineShape, aLtTxt As String
Dim dictMdb As New dBase, cnt As New ADODB.Connection, rst As New ADODB.Recordset
dictMdb.cnt�d�r cnt
For Each inlnsp In rng.InlineShapes
    aLtTxt = inlnsp.AlternativeText
    If Len(aLtTxt) < 3 Then
        'inlnsp.Delete
    Else
        If aLtTxt Like "?��?? -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "�� --�]�y��z�W�y���z���y��z�P�y�@�z�۳s�^" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "�]??�ۡ^" Then
            aLtTxt = "�Y"
        ElseIf aLtTxt Like "??��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "??? -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "�� --�]�y��z�W�y���z���y��z�P�y�@�z�۳s�^" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "?????�� -- Ĭ" Then
            aLtTxt = "Ĭ"
        ElseIf aLtTxt Like ChrW(12272) & ChrW(-10155) & ChrW(-8696) & ChrW(31860) Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "�s�]?�}���^-- ��������" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "?�g? -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "????�f?�� -- " & ChrW(-10111) & ChrW(-8620) Then
            aLtTxt = ChrW(-10111) & ChrW(-8620)
        ElseIf aLtTxt Like "???? -- �e" Then
            aLtTxt = "�e"
        ElseIf aLtTxt Like "?????? -- �d" Then
            aLtTxt = "�d"
        ElseIf aLtTxt Like "???? -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "????? --�\" Then
            aLtTxt = "�\"
        ElseIf aLtTxt Like "?��? -- �|" Then
            aLtTxt = "�|"
        ElseIf aLtTxt Like "?��?? -- �\" Then
            aLtTxt = "�\"
        ElseIf aLtTxt Like "???? -- �H" Then
            aLtTxt = "�H"
        ElseIf aLtTxt Like "�]?�D�u�^" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "??�D??�� -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "�� -- �� ?" Then
            aLtTxt = ChrW(-10114) & ChrW(-9161)
        ElseIf aLtTxt Like "�E --�]�y��z�W�y���z���y��z�P�y�@�z�۳s�^" Then
            aLtTxt = "�E"
        ElseIf aLtTxt Like "�� --�]�y��z�W�y���z���y��z�P�y�@�z�۳s�^" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "�z --�]�y��z�W�y���z���y��z�P�y�@�z�۳s�^" Then
            aLtTxt = "�z"
        ElseIf aLtTxt Like "��(�u���v�אּ�u??�v)" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "�� --�]�k�W�y��z�r�U�@?���X�A�����y���z�r���y��z�P�y�@�z�۳s�^" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like ChrW(24298) & "�]" & ChrW(8220) & ChrW(13357) & ChrW(8221) & "����" & ChrW(8220) & "��" & ChrW(8221) & "�^" Then
            aLtTxt = "�o"
        ElseIf aLtTxt Like ChrW(12273) & ChrW(11966) & ChrW(30464) Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "?�L? -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like ChrW(12272) & ChrW(-10145) & ChrW(-8265) & "��" Then
            aLtTxt = "����" & aLtTxt & "��"
        ElseIf aLtTxt Like "? -- or ?? ?" Then
            aLtTxt = ChrW(-32119)
        ElseIf aLtTxt Like "��" Then
            aLtTxt = ChrW(18518)
        ElseIf aLtTxt Like "��" Then
            aLtTxt = ChrW(17403)
        ElseIf aLtTxt Like ChrW(12272) & ChrW(-10145) & ChrW(-8265) & ChrW(25908) Then
            aLtTxt = ChrW(-10109) & ChrW(-8699)
        ElseIf aLtTxt Like "??�K -- " & ChrW(-10170) & ChrW(-8693) Then
            aLtTxt = ChrW(-10124) & ChrW(-9097)
        ElseIf aLtTxt Like ChrW(12282) & ChrW(-28746) & "��" Then
            aLtTxt = "�A"
        ElseIf aLtTxt Like "??�H -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "??̱ -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "?????�� -- �~" Then
            aLtTxt = "�~"
        ElseIf aLtTxt Like "???�\ -- " & ChrW(31762) Then
            aLtTxt = "�y"
        ElseIf aLtTxt Like "????�Z -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "?�¤� -- ?" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like ChrW(12282) & ChrW(-28746) & ChrW(17807) Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "?�ܤ� -- ??" Then
            aLtTxt = "�P"
        ElseIf aLtTxt Like "�]???�k�^" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "�]???�O�^" Then
            aLtTxt = ChrW(-10174) & ChrW(-9072)
        ElseIf aLtTxt Like "??? -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "�]???�^-- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "?�إ� -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "???? -- " & ChrW(-10161) & ChrW(-8272) Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "?�f? -- �A" Then
            aLtTxt = "�A"
        ElseIf aLtTxt Like "?�f? -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "???? -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "?�j?? -- �}" Then
            aLtTxt = ChrW(-10158) & ChrW(-8444)
        ElseIf aLtTxt Like "*page2700-20px-SKQSfont.pdf.jpg*" Then
            aLtTxt = "�@"
        ElseIf aLtTxt Like ChrW(12273) & ChrW(11966) & ChrW(12272) & ChrW(27701) & ChrW(20158) Then
            aLtTxt = ChrW(-10161) & ChrW(-8915)
        ElseIf aLtTxt Like "???��ۤP?�I? -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "?�ޤ� -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like ChrW(12272) & "��" & ChrW(-10170) & ChrW(-8693) Then
            aLtTxt = ChrW(-10121) & ChrW(-8228)
        ElseIf aLtTxt Like "??? -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "??? -- �S" Then
            aLtTxt = "�S"
        ElseIf aLtTxt Like "??�V -- " & ChrW(-28664) Then
            aLtTxt = "�~"
        ElseIf aLtTxt Like "?���� -- " & ChrW(-24830) Then
            aLtTxt = ChrW(-24830)
        ElseIf aLtTxt Like "???????��-- �Z" Then
            aLtTxt = "�Z"
        ElseIf aLtTxt Like "??? -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "�]?��?�^" Then
            aLtTxt = ChrW(-30654)
        ElseIf aLtTxt Like "SKchar" Then
            GoTo nxt
'            aLtTxt = "�e,�u,�~,�T,�V,�j,�|,��,���]2DB7E�^,��,��,��,�|,��,�s,�p"'�l�� �d�r.mdb
        ElseIf aLtTxt Like "SKchar2" Then
            GoTo nxt
'            aLtTxt = "��]7E92�^,��,"'�l�� �d�r.mdb
        Else
            Select Case aLtTxt
                Case ChrW(12280) & ChrW(30098) & ChrW(-28523)
                    aLtTxt = "����" & aLtTxt & "��"
                    '�ʦr�h�������J�r�ϴ��N��r
                    GoTo replaceIt
                Case Else
                    Dim rp As Boolean
                    rst.Open "select * from �����w�y�r�Ϩ��N��Ӫ� where (strcomp(find, """ & aLtTxt & """)=0 " & _
                        "and not find like ""SKchar*"") ", cnt, adOpenStatic, adLockReadOnly
'                    If rst.RecordCount > 0 Then
                    Do Until rst.EOF
                        aLtTxt = rst.Fields("replace").Value
                        rp = True
                        Exit Do
                    Loop
'                    Else
                        rst.Close
                        If Not rp Then
                            GoTo nxt
                        Else
                            rp = False
                        End If
'                    End If
'                    rst.Close
            End Select
        End If
    End If
replaceIt:
    inlnsp.Select
    Selection.TypeText aLtTxt
    inlnsp.Delete
nxt:
Next inlnsp
cnt.Close
End Sub
Sub ��Ǥj�v_�|�w���ѥ����()
Dim rng As Range, noteRng As Range
Set rng = Documents.Add().Range
SystemSetup.playSound 1
rng.Paste
'���ܶK�W�Lê
SystemSetup.playSound 1 '���K�W�ӮɴN�ܤ[�F�A�᭱�o�@�j�說�l�Ϧӧ� 20230211
'With rng.Find
'    .ClearAllFuzzyOptions
'    .ClearFormatting
'    .Execute "^l", , , True, , , True, wdFindContinue, , "^p", wdReplaceAll
'End With
With rng.Find
    .ClearAllFuzzyOptions
    .ClearFormatting
    .MatchWildcards = True
    .Execute "[[]*[]]  ", , , True, , , True, wdFindContinue, , "", wdReplaceAll
    .ClearAllFuzzyOptions
    .ClearFormatting
End With
Do While rng.Find.Execute("[[]", , , , , , True, wdFindContinue)
   rng.MoveEndUntil "]"
   rng.SetRange rng.start, rng.End + 1
   rng.Delete
Loop
Set rng = rng.Document.Range
rng.Find.Execute "^p^p", , , , , , , wdFindContinue, , "^p", wdReplaceAll
rng.Find.Font.Color = 16711935
Do While rng.Find.Execute("", , , False, , , True, wdFindStop)
    Set noteRng = rng
    Do While noteRng.Next.Font.Color = 16711935
        noteRng.SetRange noteRng.start, noteRng.Next.End
    Loop
    noteRng.Text = "{{" & Replace(noteRng, "/", "") & "}}"
Loop

��r�B�z.�ѦW���g�W���Ъ`

With rng.Document
'    With .Range.Find
'        .ClearFormatting
'        .Text = ChrW(9675)
'        .Replacement.Text = ChrW(12295)
'        .Execute , , , , , , True, wdFindContinue, , , wdReplaceAll
'    End With
    '.Range.Cut
    SystemSetup.ClipboardPutIn .Range.Text
    DoEvents
    .Close wdDoNotSaveChanges
End With
SystemSetup.playSound 1.921
End Sub

Sub mdb�}�o_�d�{�@�o�NExport()
Dim cnt As New ADODB.Connection, db As New dBase, rst As New ADODB.Recordset, exportStr As String, preTitle As String, title As String
Const bookName As String = "��ۥ��骾��" '����e�Х����w�ѦW
db.cnt_�}�o_�d�{�@�o�N cnt
rst.Open "SELECT �g.�g�W, ��.���O, ��.�ѦW, �g.��, �g.��, �g.����, ��.�gID, ��.��, ��.��ID, ��.��ID, ���O�D�D.���O�D�D" & _
        " FROM ���O�D�D INNER JOIN ((�� INNER JOIN �g ON ��.��ID = �g.��ID) INNER JOIN �� ON �g.�gID = ��.�gID) ON ���O�D�D.��ID = ��.��ID" & _
        " WHERE (((��.�ѦW)=""" & bookName & """) AND ((���O�D�D.���O�D�D) Not Like "" * �u�� * "" Or (���O�D�D.���O�D�D) Is Null))" & _
        " ORDER BY �g.��, �g.��, �g.����, ��.�gID, ��.��, ��.��ID;", cnt, adOpenKeyset, adLockReadOnly
Do Until rst.EOF
    title = rst.Fields(0).Value
    If preTitle <> title Then
        exportStr = exportStr & chr(13) & "*" & title & chr(13)
    End If
    preTitle = title
    exportStr = exportStr & rst.Fields(1).Value
    rst.MoveNext
Loop
rst.Close
cnt.Close
Documents.Add.Range = exportStr
End Sub
Sub �M���Ҧ��Ÿ�_�[�W����_�@�����}���()
    Dim rng As Range, e, sybol 'Alt + l
    sybol = Array("(", ")", "�]", "�^")
    Set rng = Documents.Add().Range
    rng.Paste
    Docs.�M���Ҧ��Ÿ�
    For Each e In sybol
        rng.Text = Replace(rng, e, "")
    Next e
    rng.Text = "#" & rng.Text
    rng.Cut
    rng.Document.Close wdDoNotSaveChanges
    DoEvents
    AppActivateDefaultBrowser
    SendKeys "^v~"
    DoEvents
    SendKeys "^l^c"
    DoEvents
    SendKeys "{F5}"
End Sub
Sub ���J�W�s��_�N��ܤ��s�X�אּ����()
Const keys As String = "&searchu=" 'Alt + j
Dim rng As Range, lnk As String, cde As String, s As Long, d As Document, ur As UndoRecord
Set rng = Selection.Range: Set d = ActiveDocument
lnk = SystemSetup.GetClipboardText
cde = Mid(lnk, InStr(lnk, keys) + Len(keys))
cde = code.URLDecode(cde)
s = Selection.start
SystemSetup.stopUndo ur, "���J�W�s��_�N��ܤ��s�X�אּ����"
With Selection
    .Hyperlinks.Add Selection.Range, lnk, , , left(lnk, InStr(lnk, keys) + Len(keys) - 1) + cde
    'd.Range(Selection.End, Selection.End + Len(cde)).Select
    'Selection.Collapse
    .MoveLeft wdCharacter, Len(cde)
    .MoveRight wdCharacter, Len(cde) - 1, wdExtend
    .Range.HighlightColorIndex = wdYellow
    .move , 2
    .InsertParagraphAfter
    .InsertParagraphAfter
    .Collapse
End With
SystemSetup.contiUndo ur
End Sub
Sub �u�O�d����`��_�B�`��e��[�A��()
Dim d As Document, ur As UndoRecord, slRng As Range
SystemSetup.stopUndo ur, "������Ǯѹq�l�ƭp��_�u�O�d����`��_�B�`��e��[�A��"
Docs.�ťժ��s���
Set d = ActiveDocument
If Selection.Type = wdSelectionIP Then ActiveDocument.Select
Set slRng = Selection.Range
������Ǯѹq�l�ƭp��_������r slRng
Dim ay, e
ay = Array(254, 8912896)
With d.Range.Find
    .ClearFormatting
End With
For Each e In ay
    With d.Range.Find
        .Font.Color = e
        .Execute "", , , , , , True, wdFindContinue, , "", wdReplaceAll
    End With
Next e
Set slRng = d.Range
With slRng.Find
    .ClearFormatting
    .Font.Color = 34816
End With
Do While slRng.Find.Execute(, , , , , , True, wdFindStop)
    If InStr(chr(13) & chr(11) & chr(7) & chr(8) & chr(9) & chr(10), slRng) = 0 Then
    slRng.Text = "�]" + slRng.Text + "�^"
    'slRng.SetRange slRng.End, d.Range.End
    End If
Loop
SystemSetup.contiUndo ur
End Sub

Sub �����w���������⴫���r(d As Document)
Dim rst As New ADODB.Recordset, cnt As New ADODB.Connection, db As New dBase
db.cnt�d�r cnt
rst.Open "select * from �����w���������⴫���r where doIt = true order by len(replaced) desc", cnt, adOpenForwardOnly, adLockReadOnly
Do Until rst.EOF
    d.Range.Find.Execute rst.Fields("replaced").Value, , , , , , True, wdFindContinue, , rst.Fields("replacewith").Value, wdReplaceAll
    rst.MoveNext
Loop
rst.Close: cnt.Close: Set db = Nothing
End Sub

Sub dbSBCKWordtoReplace() '�|���O�Z�y�r��Ӫ� Alt+5
Dim rng As Range, ur As UndoRecord
'Set ur = stopUndo("�m�|���O�Z�n��Ʈw�y�r���N���t�Φr")
SystemSetup.stopUndo ur, "�m�|���O�Z�n��Ʈw�y�r���N���t�Φr"
If ActiveDocument.Name = "�m�|���O�Z��Ʈw�n�ɤJ�m������Ǯѹq�l�ƭp���n.docm" Then
    Set rng = ActiveDocument.Range
Else
    Set rng = Documents.Add.Range
    rng.Paste
End If
dbSBCKWordtoReplaceSub rng
If Not ActiveDocument.Name = "�m�|���O�Z��Ʈw�n�ɤJ�m������Ǯѹq�l�ƭp���n.docm" Then
    rng.Cut
    If rng.Application.Documents.Count = 1 Then
        rng.Application.Quit wdDoNotSaveChanges
    Else
        rng.Document.Close wdDoNotSaveChanges
    End If
Else
    ActiveDocument.Save
End If
contiUndo ur
End Sub
Sub dbSBCKWordtoReplaceSub(ByRef rng As Range)
Const tbName As String = "�|���O�Z�y�r��Ӫ�"
Dim rst As New ADODB.Recordset, cnt As New ADODB.Connection, db As New dBase
rng.Find.ClearFormatting
db.cnt�d�r cnt
rst.Open tbName, cnt, adOpenForwardOnly, adLockReadOnly
Do Until rst.EOF
    If InStr(rng.Text, rst.Fields(0).Value) Then _
        rng.Find.Execute rst.Fields(0).Value, , , , , , True, wdFindContinue, , rst.Fields(1).Value, wdReplaceAll
    rst.MoveNext
Loop
rst.Close: cnt.Close: Set db = Nothing
End Sub

Sub dbSBCKWordtoReplace_AddNewOne() '�|���O�Z�y�r��Ӫ� Alt+4
Const tbName As String = "�|���O�Z�y�r��Ӫ�"
Dim rst As New ADODB.Recordset, cnt As New ADODB.Connection, db As New dBase
Dim rng As Range
Set rng = Selection.Range
db.cnt�d�r cnt
rst.Open "select * from " + tbName + " where strcomp(�y�r, """ + rng.Characters(1) + """)=0", cnt, adOpenKeyset, adLockOptimistic
If rst.RecordCount = 0 Then
    If rng.Characters.Count = 2 Then
        'rst.Open tbName, cnt, adOpenKeyset, adLockOptimistic
        rst.AddNew
        rst.Fields(0) = rng.Characters(1)
        rst.Fields(1) = rng.Characters(2)
        rst.Update
        rng.Characters(1).Delete
    Else
        MsgBox "plz input the replace word next the one"
        Selection.move
    End If
Else
    If rng.Characters.Count = 2 Then If rng.Characters(2) = rst.Fields(1).Value Then rng.Characters(2).Delete
    rng.Characters(1) = rst.Fields(1).Value
End If
rst.Close: cnt.Close: Set db = Nothing
dbSBCKWordtoReplaceSub rng.Document.Range
End Sub

Sub entity_Markup_edit_via_API_Annotate_Reverting()
Dim rng As Range, rngMark As Range, d As Document, ay(), e, i As Long, DoctoMarked As Document
Set d = ActiveDocument
Const markStrOpen As String = "<entity ", markStrClose As String = "</entity>"
If InStr(d.Range, markStrOpen) = 0 Then
    MsgBox "plz paste the marked text in active doc First thx"
    Exit Sub
End If
Set rng = d.Range
'get the terms which were marked
Do While rng.Find.Execute(markStrOpen)
    'rng.SetRange rng.start, rng.End + rng.MoveEndUntil(markStrClose)
    rng.SetRange rng.start, rng.End + rng.MoveEndUntil("/")
    If d.Range(rng.End, rng.End + 7) = "entity>" Then
        rng.SetRange rng.start, rng.End - 2
        ReDim Preserve ay(i)
        ay(i) = VBA.Split(rng.Text, ">")
        i = i + 1
    End If
    
    rng.SetRange rng.End, d.Range.End
Loop
'got the terms which were marked already
'mark the text
'Stop
If MsgBox("if NOT text to be marked already copied then push CANCEL button", vbOKCancel + vbExclamation) = vbCancel Then Exit Sub
Set DoctoMarked = Documents.Add
Set rng = DoctoMarked.Range: Set rngMark = DoctoMarked.Range
rng.Paste
For Each e In ay
reFind:
    If rng.Find.Execute(e(1)) Then
        If rng.Characters(1).Previous = ">" Then
            rngMark.SetRange rng.start - 1, rng.start
            'rngMark.MoveStartUntil "<"
            Do Until DoctoMarked.Range(rngMark.start, rngMark.start + 1) = "<"
                rngMark.move wdCharacter, -1
            Loop
            rngMark.SetRange rngMark.start, rng.start
            If left(rngMark.Text, 8) <> "<entity " Then
                GoSub mark
            Else
                rng.SetRange rng.End, DoctoMarked.Range.End
                GoTo reFind
            End If
        Else
            GoSub mark
        End If
    Else
        SystemSetup.ClipboardPutIn CStr(e(1))
        MsgBox "plz check out why the " + e(1) + "dosen't exist !!", vbExclamation
        Stop
        Set rng = DoctoMarked.Range
        GoTo reFind
        'Exit Sub
    End If
    rng.SetRange rng.End, d.Range.End
Next e
Beep
Exit Sub
mark:
    rng.InsertAfter "</entity>"
    rng.InsertBefore e(0) + ">"
Return
End Sub

Sub checkEditingOfPreviousVersion()
    Dim d As Document, rng As Range
    Set d = Documents.Add()
    Set rng = d.Range
    rng.Paste
    GoSub fontColor
    GoSub punctuations
    If d.Application.Documents.Count = 1 Then
        d.Application.Quit wdDoNotSaveChanges
    Else
        d.Close wdDoNotSaveChanges
    End If
    Exit Sub
     
     
fontColor:
    
        rng.Find.ClearFormatting
        rng.Find.Font.Color = 8912896 '{{{}}}�y�k�U����r
        rng.Find.Replacement.ClearFormatting
        With rng.Find
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchByte = True
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        If (rng.Find.Execute) Then GoSub CheckOut
    Return
    
punctuations:
        rng.Find.ClearFormatting
        rng.Find.Replacement.ClearFormatting
        Dim punctus, e
        punctus = Array("�A", "�C", "�u", "�P", "�G", "�]")  '�ˬd�X�Ө�N��̧Y�i
        For Each e In punctus
            If InStr(rng.Text, e) > 0 Then
                rng.Find.Execute e
                GoTo CheckOut
            End If
        Next e
    Return
    
CheckOut:
        rng.Select
        d.ActiveWindow.Visible = True
        d.ActiveWindow.ScrollIntoView rng
        MsgBox "plz check it out !", vbExclamation
End Sub

Sub EditModeMakeup_changeFile_Page() '�P�����奻�a�J�m��file id �M ����
    Dim rng As Range, pageNum As Range, d As Document, ur As UndoRecord
    Set d = ActiveDocument
    
    '���e3�q���O�O�H�U��T,���槹�|�M��
    'If Not VBA.IsNumeric(VBA.Replace(d.Range.Paragraphs(1).Range.text, Chr(13), "")) then
    If Replace(d.Paragraphs(1).Range + d.Paragraphs(2).Range + d.Paragraphs(3).Range, chr(13), "") = "" _
        Or Not IsNumeric(Replace(d.Paragraphs(1).Range + d.Paragraphs(2).Range + d.Paragraphs(3).Range, chr(13), "")) Then
        MsgBox "�Цb���e3�q���O�O�H�U��T�]�ҬO�Ʀr�^,���槹�|�M��" & vbCr & vbCr & _
            "1. ���Ʈt(�ӷ�-(��h)�ت��^�C�L���t�h��0�A�ٲ��h�w�]��0" & vbCr & _
            "2. �ت��� file number�C�n�m�������F�����N�h��0�A�ٲ��h�w�]��0" & vbCr & _
            "3. �ӷ��� file number�A�n�Q���N��,�ٲ��]���n�Ũ�q��=�Ŧ�^�h����󤤪�file=�᪺��"
        Exit Sub
    End If
    Dim differPageNum  As Integer '���Ʈt(�ӷ�-(��h)�ت��^
    differPageNum = VBA.IIf(d.Paragraphs(1).Range.Characters.Count = 1, 0, VBA.Replace(d.Paragraphs(1).Range.Text, chr(13), "")) '���Ʈt(�ӷ�-(��h)�ت��^
    Dim file
    file = VBA.Replace(d.Paragraphs(2).Range.Text, chr(13), "") ' �ت��C�����N�h��0
    If file = "" Then file = 0
    Dim fileFrom As String
    fileFrom = VBA.Replace(d.Paragraphs(3).Range.Text, chr(13), "") ' '�ӷ�
    If fileFrom = "" Then
        Dim s As String: s = VBA.InStr(d.Range.Text, "<scanbegin file="): s = s + VBA.Len("<scanbegin file=")
        fileFrom = Mid(d.Range.Text, s + 1, InStr(s + 1, d.Range.Text, """") - s - 1)
    End If
    Set rng = d.Range
    'Set ur = SystemSetup.stopUndo("EditMakeupCtext")
    SystemSetup.stopUndo ur, "EditMakeupCtext"
    If file > 0 Then
        'rng.Find.Execute " file=""77991""", True, True, , , , True, wdFindContinue, , " file=""" & file & """", wdReplaceAll
        rng.Text = Replace(rng.Text, " file=""" & fileFrom & """", " file=""" & file & """")
    End If
    
    Do While rng.Find.Execute(" page=""", , , , , , True, wdFindStop)
        Set pageNum = rng
        pageNum.SetRange rng.End, rng.End + 1
        pageNum.MoveEndUntil """"
        pageNum.Text = CStr(CInt(pageNum.Text) - differPageNum)
        rng.SetRange pageNum.End, d.Range.End
    Loop
    rng.SetRange d.Range.Paragraphs(1).Range.start, d.Range.Paragraphs(3).Range.End
    rng.Delete
    'd.Range.Cut
    SystemSetup.SetClipboard d.Range.Text
    SystemSetup.contiUndo ur
    SystemSetup.playSound 1
    d.Application.Activate
End Sub


Sub tempReplaceTxtforCtextEdit()
Dim a, d As Document, i As Integer, x As String
a = Array("{{�]", "{{", "�^}}", "}}", "�]", "{{", "�^", "}}", "��", ChrW(12295))
Set d = Documents.Add
d.Range.Paste
x = d.Range
For i = 0 To UBound(a)
    x = Replace(x, a(i), a(i + 1))
    i = i + 1
Next i
d.Range = x
d.Range.Cut
d.Close wdDoNotSaveChanges
AppActivateDefaultBrowser
SendKeys "^v"
End Sub


Sub tempReplaceTxtforCtext() 'for Quick edit only
Dim a, d As Document, i As Integer
a = Array("{{�]", "{{", "�^}}", "}}", "�]", "{{", "�^", "}}", "��", ChrW(12295))
Set d = Documents.Add
d.Range.Paste
For i = 0 To UBound(a)
    d.Range.Find.Execute a(i), , , , , , , wdFindContinue, , a(i + 1), wdReplaceAll
    i = i + 1
Next i
d.Range.Cut
d.Application.WindowState = wdWindowStateMinimize
d.Close wdDoNotSaveChanges
AppActivateDefaultBrowser
SendKeys "^v"
SendKeys "{tab}~"

End Sub



