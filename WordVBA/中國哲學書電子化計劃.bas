Attribute VB_Name = "������Ǯѹq�l�ƭp��"
Option Explicit
Dim ChapterSelector As String
'Const description As String = "�N�P���e�����q�Ÿ����m�e�q���� & �M�����e�����q�Ÿ�"
'Const description As String = "�N�P���e�����q�Ÿ����m�e�q���� & �M�����e�����q�Ÿ�{��̤l���m�Ǫ̮]�u�u���u�j���G���̧Ӥh�q�����Ľ�ͽЦh�Q�ν�W�k�v�m�j�y��AI�n�Ρm�ݨ�j�y�nOCR�ƥb�\���]�C�p�X����A�i�Q�Υ��ǩ�GitHub�}���K�O�K�w�ˤ�TextForCtext ���ε{���A�[�t��J�P�ƪ��C�Q�װϻP����YouTube�W�D���t�ܼv���i��ѦҡC�P���P���@�g���g�ۡ@�n�L��������"
'Const description As String = "�N�P���e�����q�Ÿ����m�e�q���� & �M�����e�����q�Ÿ�{��Kanripo.org�Ρm��Ǥj�v�n���å����H���Ǧۻs��GitHub�}���K�O�K�w�ˤ�TextForCtext�ƪ��������J�C�Q�װϻP����YouTube�W�D����Һt�ܼv���i��ѦҡC�P���P���@�g���g�ۡ@�n�L��������@�g���D}"
'Const description As String = "�N�P���e�����q�Ÿ����m�e�q���� & �M�����e�����q�Ÿ�{�ڡm��Ǥj�v�n�Υ_�ʤ��ެ�ަ������q�m���ެ�ޤޱo�Ʀr�H��귽���O�P������N���m�n���å����H���Ǧۻs��GitHub�}���K�O�K�w�ˤ�TextForCtext�ƪ��������J�C�Q�װϻP����YouTube�W�D����Һt�ܼv���i��ѦҡC�P���P���@�g���g�ۡ@�n�L��������@�g���D}"
Const description As String = "�N�P���e�����q�Ÿ����m�e�q���� & �M�����e�����q�Ÿ�{�ڥ_�ʤ��ެ�ަ������q�m���ެ�ޤޱo�Ʀr�H��귽���O�P������N���m�n���å����H���Ǧۻs��GitHub�}���K�O�K�w�ˤ�TextForCtext�ƪ��������J�C�Q�װϻP����YouTube�W�D����Һt�ܼv���i��ѦҡC�P���P���@�g���g�ۡ@�n�L��������@�g���D}"

'Const description_Edit_textbox_�s���� As String = "�ڡm��Ǥj�v�n�ΡmKanripo�n�Ҧ������H���Ǧۻs��GitHub�}���K�O�K�w�ˤ�TextForCtext�n��ƪ��������J�F�Q�װϤΥ���YouTube�W�D����Һt�ܼv���C�P���P���@�g���g�ۡ@�n�L��������"
Const description_Edit_textbox_�s���� As String = "�ڥ_�ʤ��ެ�ަ������q�m���ެ�ޤޱo�Ʀr�H��귽���O�P������N���m�n�Ҧ������H���Ǧۻs��GitHub�}���K�O�K�w�ˤ�TextForCtext�n��ƪ��������J�F�Q�װϤΥ���YouTube�W�D����Һt�ܼv���C�P���P���@�g���g�ۡ@�n�L��������"

Sub ������q()
    
    Dim lineLength As Byte, d As Document, rng As Range, si As New StringInfo, firstLineIndentValue As Single, leadSpaceCount As Byte, p As Paragraph, leadSpaces As String, i As Long, t As table
    
    lineLength = 21 ''�Ĥ@����w���`�����: d.Paragraphs(1).Range.Characters.Count - 1
    'd.Paragraphs(1).Range.text = vbNullString
    
    Set d = Documents.Add
    d.Range.Paste
    
    
    For Each p In d.Paragraphs
        If p.Style = "����" And p.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft Then
            firstLineIndentValue = p.Range.ParagraphFormat.FirstLineIndent
            If firstLineIndentValue <> 0 Then
                leadSpaceCount = VBA.Abs(firstLineIndentValue) / d.Paragraphs(1).Range.Characters(1).font.Size
            End If
            Exit For
        End If
    Next p
    'firstLineIndentValue = d.Paragraphs(1).Range.ParagraphFormat.FirstLineIndent
    
    
    For Each t In d.tables
        t.Delete
    Next t
    '�M�� �i�ϡj�]�m�~�y�����Ʈw�n�奻�A�H��ƻs��r�\��^
    If VBA.InStr(d.Range.text, "�i�ϡj") Then d.Range.Find.Execute "�i�ϡj", , , , , , , wdFindContinue, , vbNullString, wdReplaceAll
    d.Range.Find.Execute "^l", , , , , , , wdFindContinue, , vbNullString, wdReplaceAll
    For Each p In d.Paragraphs
        If p.Style = "����" And p.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft Then
            If Not p.Next Is Nothing Then
                If p.Next.Range.text <> "�D�@�D�@�D�@�D�@�D�@�D�@�D�@�D�@�D�@�D�@�D�@�D�@�D�@�D�@�D�@�D�@�D�@�D" & Chr(13) Then
                    p.Range.Characters(p.Range.Characters.Count).text = vbNullString
                    Set p = p.Previous
                Else
                    p.Next.Range.text = vbNullString
                End If
            End If
        End If
    Next p
'    d.Range.Find.Execute "�D�@�D�@�D�@�D�@�D�@�D�@�D�@�D�@�D�@�D�@�D�@�D�@�D�@�D�@�D�@�D�@�D�@�D^p", , , , , , , wdFindContinue, , vbNullString, wdReplaceAll

    Dim lineCntr As Byte, noteCntr As Long
    For Each p In d.Paragraphs
        If p.Style = "����" And p.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft Then
        
'            If InStr(p.Range, "�@�@��Z") Then
'                p.Range.Select
'                Stop
'            End If

            If p.Range.Characters.Count - 1 > lineLength Then
                i = 0 ''i =  1 'lineLength
                Set rng = p.Range.Characters(1)
                Do While i + lineLength < p.Range.Characters.Count - 1 'Step lineLength 'p.Range.Characters.Count Step lineLength
lastfew:
                    
                    lineCntr = 0
                    Do While lineCntr < lineLength
                        If rng.font.Size = 7.5 Then '�p�`
                            noteCntr = noteCntr + 1
                            rng.Move wdCharacter, 1
                            i = i + 1
                            If rng.font.Size = 7.5 Then
                                noteCntr = noteCntr + 1
                                rng.Move wdCharacter, 1
                                i = i + 1
                            End If
                            'i = i + 2
                        Else
                            If noteCntr > 0 Then
                                If noteCntr Mod 2 = 1 Then lineCntr = lineCntr + 1
                                noteCntr = 0
                            End If
                            rng.Move wdCharacter, 1
                            i = i + 1
                        End If
                        lineCntr = lineCntr + 1
                    Loop
                    
'                    rng.Select
                    '�p�G���Y��
                    If leadSpaceCount > 0 Then
                        If VBA.InStr(p.Range.text, Chr(11)) Then
                            'p.Range.Characters(i - leadSpaceCount).InsertAfter Chr(11)
                            rng.Move wdCharacter, -leadSpaceCount
                            rng.InsertAfter Chr(11)
                            rng.Collapse wdCollapseEnd
                            i = i - leadSpaceCount
                        Else
                            'p.Range.Characters(i).InsertAfter Chr(11)
                            rng.InsertAfter Chr(11)
                            rng.Collapse wdCollapseEnd
                        End If
                    Else '�S���Y��
                        'p.Range.Characters(i).InsertAfter Chr(11)
                        rng.InsertAfter Chr(11)
                        rng.Collapse wdCollapseEnd
                    End If
                    
                    i = i + 1
'
                    'i = i + lineLength
                Loop
            End If
        End If
    Next p
    
    Set rng = d.Range
    With rng.Find
        .ClearFormatting
        .font.Size = 7.5
    End With
    
    Do While rng.Find.Execute()
        rng.InsertAfter "}}"
        rng.InsertBefore "{{"
        rng.Collapse wdCollapseEnd
    Loop
    
    d.Range.Find.ClearFormatting
    
    leadSpaces = VBA.StrConv(VBA.space(leadSpaceCount), vbWide)
    d.Range.Find.Execute "^l", , , , , , , , , "^p" & leadSpaces, wdReplaceAll
    d.Range.Cut
    d.Close wdDoNotSaveChanges
    AppActivate "TextForCtext"
    DoEvents
    SendKeys "^v"
    DoEvents
End Sub

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
    If d.path <> "" Then Exit Sub
    Set rng = d.Range
    start = CInt(Replace(rng.Paragraphs(1).Range, VBA.Chr(13), ""))
    e = CInt(Replace(rng.Paragraphs(2).Range, VBA.Chr(13), ""))
    fileID = CLng(Replace(rng.Paragraphs(3).Range, VBA.Chr(13), ""))
    For i = start To e
        If i = 1 Then
            x = x & "<scanbegin file=""" & fileID & """ page=""" & i & """ />��" & VBA.Chr(9) & "<scanend file=""" & fileID & """ page=""" & i & """ />"
        Else
            x = x & "<scanbegin file=""" & fileID & """ page=""" & i & """ />" & VBA.Chr(9) & "<scanend file=""" & fileID & """ page=""" & i & """ />" '�Y�����S�����󤺮e�A�����̫�K���ন�@�q���C�Y��n�@�Ӭq���A�|�P�U�@���H�X�b�@�_
        End If
    Next i
    
    rng.Document.Range(d.Paragraphs(3).Range.start, d.Paragraphs(3).Range.End - 1).text = CLng(Replace(rng.Paragraphs(3).Range, VBA.Chr(13), "")) + 1
    'rng.Document.Paragraphs(3).Range.text = VBA.CStr(VBA.CLng(VBA.Left(rng.Document.Paragraphs(3).Range.text, VBA.Len(rng.Document.Paragraphs(3).Range.text) - 1)) + 1)
    
    'For Each e In Selection.Value
    '    x = x & e
    'Next e
    ''x = Replace(x, vba.Chr(13), "")
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
    rng.Document.ActiveWindow.windowState = wdWindowStateMinimize
    DoEvents
    'Network.AppActivateDefaultBrowser
    ActivateChrome
'    SendKeys "^a"
'    SendKeys "^v"
    
    SystemSetup.contiUndo ur
End Sub
Sub setPage1Code() '(ByRef d As Document)
    Dim xd As String
    xd = SystemSetup.GetClipboardText
    If InStr(xd, "page=""1""") = 0 Then
        Dim bID As String, s As Byte, pge As String
        s = InStr(xd, "page=""")
        pge = VBA.Mid(xd, s + Len("page="""), InStr(s + Len("page="""), xd, """") - s - Len("page="""))
        If CInt(pge) < 10 Then
            s = InStr(xd, """")
            bID = VBA.Mid(xd, s + 1, InStr(s + 1, xd, """") - s - 1)
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
        xd = VBA.Mid(xd, 1, e) + VBA.Mid(xd, s)
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
AppActivateChrome
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
    rng.text = rng.text + "*"
    s = rng.End + 1
    rng.Collapse wdCollapseStart
    rng.SetRange rng.start, rng.start
    'rng.MoveStartUntil ">"
    Do Until rng.Next.text = "<"
        rng.Move wdCharacter, -1
    Loop
    rng.Move
    rng.text = rng.text + VBA.Chr(13) + VBA.Chr(13)
    rng.SetRange s, d.Range.End
    Return
End Sub

Sub �M�����e�����q�Ÿ�()
    Dim d As Document, rng As Range, e As Long, s As Long, xd As String
    Dim iwe As SeleniumBasic.IWebElement
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
        If rng.text = VBA.Chr(13) & VBA.Chr(13) Then
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
        If rng.text = VBA.Chr(13) & VBA.Chr(13) Then
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
    xd = d.Range.text
    'If d.Characters.Count < 50000 Then ' 147686
    '    d.Range.Cut '��ӬOWord�� cut ��ŶKï�̦����D
    'Else
        'SystemSetup.SetClipboard d.Range.Text
        SystemSetup.ClipboardPutIn xd
    'End If
    DoEvents
    playSound 1, 0
    DoEvents
    
    pastetoEditBox description
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
            If rng.Previous = VBA.Chr(13) Then
                Set rng = rng.Previous
                If rng.Previous = VBA.Chr(13) Then
                    Set rng = rng.Previous
                    If rng.Previous = ">" Then
                        rng.SetRange rng.start, e - 1
                        s = rng.start
                        Set rngP = d.Range(s, s)
                        rng.Delete
                        Do Until rngP.Next = "<"
                            If rngP.start = 0 Then GoTo NextOne
                            rngP.Move wdCharacter, -1
                        Loop
                        '�ˬd�O�_���b�󭶳B 20230811
                        If d.Range(rngP.start, rngP.start + 11) = "><scanbegin" Then
                            rngP.Move Count:=-1
                            Do Until rngP.Next = "<"
                                If rngP.start = 0 Then GoTo NextOne
                                rngP.Move wdCharacter, -1
                            Loop
                        End If
                        '�H�W �ˬd�O�_���b�󭶳B 20230811
                        rngP.Move
                        rngP.InsertAfter VBA.Chr(13) & VBA.Chr(13)
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
    Select Case Err.number
        Case 4605, 13 '����k���ݩʵL�k�ϥΡA�]��[�ŶKï] �O�Ū��εL�Ī��C
            SystemSetup.wait 0.8
            Resume
        Case Else
            MsgBox Err.number + Err.description
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

Private Sub pastetoEditBox(Optional Description_from_ClipBoard As String = vbNullString)
    word.Application.windowState = wdWindowStateMinimize
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
    If Description_from_ClipBoard <> vbNullString Then
        SystemSetup.ClipboardPutIn Description_from_ClipBoard
        DoEvents
        SendKeys "^v"
'        SendKeys Description_from_ClipBoard
    End If
    SendKeys "{tab 2}~" '���U Submit changes
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
    Do While rng.Find.Execute("}}|" & VBA.Chr(13) & "{{", , , , , , True, wdFindStop)
        s = rng.start - 1: e = rng.start
        Do Until d.Range(s, e) <> "�@" '�M����e�Ů�
            s = s - 1: e = e - 1
        Loop
        rngDel.SetRange s + 1, rng.start
        'rngDel.Select
        If rngDel.text <> "" Then If Replace(rngDel, "�@", "") = "" Then rngDel.Delete
        rng.SetRange s + Len("}}|" & VBA.Chr(13) & "{{"), d.Range.End
        
        'Set rng = d.Range
    Loop
    d.Range.text = Replace(Replace(d.Range.text, "|" & VBA.Chr(13) & "�@", ""), "}}|" & VBA.Chr(13) & "{{", VBA.Chr(13))
    d.Range.Copy
    SystemSetup.contiUndo ur
    SystemSetup.playSound 2
    word.Application.windowState = wdWindowStateMinimize
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
        If a.text <> VBA.Chr(13) Then a.text = "��"
    Next i
    p.Range.Cut
    SystemSetup.contiUndo ur
    Set ur = Nothing
End Sub
Sub �M���Ҧ��Ÿ�_���q�`��Ÿ��ҥ~()
    Dim f, i As Integer
    f = Array("�C", "�v", VBA.Chr(-24152), "�G", "�A", "�F", _
        "�B", "�u", ".", VBA.Chr(34), ":", ",", ";", _
        "�K�K", "...", "�D", "�i", "�j", " ", "�m", "�n", "�q", "�r", "�H" _
        , "�I", "��", "��", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0" _
        , "�y", "�z", VBA.ChrW(9312), VBA.ChrW(9313), VBA.ChrW(9314), VBA.ChrW(9315), VBA.ChrW(9316) _
        , VBA.ChrW(9317), VBA.ChrW(9318), VBA.ChrW(9319), VBA.ChrW(9320), VBA.ChrW(9321), VBA.ChrW(9322), VBA.ChrW(9323) _
        , VBA.ChrW(9324), VBA.ChrW(9325), VBA.ChrW(9326), VBA.ChrW(9327), VBA.ChrW(9328), VBA.ChrW(9329), VBA.ChrW(9330) _
        , VBA.ChrW(9331), VBA.ChrW(8221), """") '���]�w���I�Ÿ��}�C�H�ƥ�
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
        If InStr(angleRng.text, "file") > 0 Then
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
                If InStr(a.Paragraphs(1).Range.text, "*") = 0 Then
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
                        rng.text = Replace(rng.text, "�@", VBA.ChrW(-9217) & VBA.ChrW(-8195))
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
            If VBA.Right(a, 1) = y Then
                If a.Previous.Previous <> ">" Then
                    For yi = 1 To 99
                        yStr = ��r�ഫ.�Ʀr��~�r2���(yi) + y
                        If a.text = yStr Then
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
    a = Array(VBA.ChrW(12296), "{{", VBA.ChrW(12297), "}}", "�q", "{{", "�r", "}}", _
        "��", VBA.ChrW(12295))
    '�m�e�N�T���n���p�`�@����٪����� https://ctext.org/library.pl?if=gb&file=89545&page=24
    'a = Array("�q", "", "�r", "", _
        "��", vba.Chrw(12295))
    
    
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
        If VBA.Left(xP, 2) = "{{" And VBA.Right(xP, 3) = "}}" & VBA.Chr(13) Then
            xP = VBA.Mid(p.Range, 3, Len(xP) - 5)
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
        ElseIf VBA.Left(xP, 1) = "�@" Then '�e���Ů檺
            i = InStr(xP, "{{")
            If i > 0 And VBA.Right(xP, 3) = "}}" & VBA.Chr(13) Then
                space = VBA.Mid(xP, 1, i - 1)
                If Replace(space, "�@", "") = "" Then
                    xP = VBA.Mid(xP, i + 2, Len(xP) - 3 - (i + 2))
                    If InStr(xP, "{{") = 0 And InStr(xP, "}}") = 0 Then
                        Set rng = p.Range
                        rng.SetRange rng.Characters(1).start, rng.Characters(i + 1).End
                        rng.text = "{{" & space
                        acP = p.Range.Characters.Count - 1 - Len(space)
                        If acP Mod 2 = 0 Then
                            acP = CInt(acP / 2) + Len(space) + 1
                        Else
                            acP = CInt((acP + 1) / 2) + Len(space) + 1
                        End If
                        If p.Range.Characters(acP).InlineShapes.Count = 0 Then
                            p.Range.Characters(acP).InsertBefore VBA.Chr(13) & space
                        Else
                            p.Range.Characters(acP).Select
                            Selection.Delete
                            Selection.TypeText " "
                            p.Range.Characters(acP).InsertBefore VBA.Chr(13) & space
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
    Select Case Err.number
        Case 5904 '�L�k�s�� [�d��]�C
            If p.Range.Characters(acP).Hyperlinks.Count > 0 Then p.Range.Characters(acP).Hyperlinks(1).Delete
            Resume
        Case Else
            MsgBox Err.number & Err.description
    End Select
End Sub

Sub �����w�|���O�Z�����_early()
    Dim d As Document, a, i
    
    a = Array("^p^p", "@", "�q", "{{", "�r", "}}", "^p", "", "}}{{", "^p", "@", "^p", _
        "��", VBA.ChrW(12295))
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
    ' Alt + Shift + ,
    ' Alt + <
    SystemSetup.playSound 0.484
    Select Case Selection.text
        Case "", VBA.Chr(13), VBA.Chr(9), VBA.Chr(7), VBA.Chr(10), " ", "�@"
            MsgBox "no selected text for search !", vbCritical: Exit Sub
    End Select
    Static bookID
    Dim searchedTerm, e, addressHyper As String, bID As String, cndn As String
    'Const site As String = "https://ctext.org/wiki.pl?if=gb&res="
    Const site As String = "https://ctext.org/wiki.pl?if=gb"
    bID = VBA.Left(ActiveDocument.Paragraphs(1).Range, Len(ActiveDocument.Paragraphs(1).Range) - 1)
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
        bookID = VBA.Mid(bookID, InStr(bookID, cndn) + Len(cndn))
        If Not VBA.IsNumeric(bookID) Then
            bookID = VBA.Mid(bookID, 0, InStr(bookID, "&searchu"))
        End If
    End If
    If Not VBA.IsNumeric(bookID) Then
        MsgBox "error . not the proper bookID ref ", vbCritical: Exit Sub
    End If
    ��r�B�z.ResetSelectionAvoidSymbols
    e = code.UrlEncode(Selection.text)
    'searchedTerm = 'Array("��", "��", "�P��", "���g", "�t��", "ô��", "����", "����", "�Ǩ�", "����", "�Ԩ�", "����", "�娥", "���[", "�L�S", vba.Chrw(26080) & "�S", "�ѩS", "����", "�Q�s", "��") ', "", "", "", "")
    ''https://ctext.org/wiki.pl?if=gb&res=757381&searchu=%E5%8D%A6
    'For Each e In searchedTerm
        addressHyper = addressHyper + " " + site + cndn + bookID + "&searchu=" + e
    'Next e
    Shell Network.getDefaultBrowserFullname + addressHyper + " --remote-debugging-port=9222 "
    
    Selection.Hyperlinks.Add Selection.Range, addressHyper
End Sub

Sub �v�O�T�a�`()
'�q2858���_�A20210920:0817����A��λO�v�j�����P�ǧd��@���͡m���ؤ�ƺ��n�ҿ�����|�m�v��n�쥻�A���Τ�����A�M�ܤ֧K��²�Ʀr�ഫ�_�~�γy�r�ýX���x�Z�A���r�ɱ�m�C�ھڪ�@���A�榡�����@�ˡI�ڥ��N�O�q�o�̥X�Ӫ��A�A��²�Ʀr�A�A�S�ϥ��A�y�������áC�����S�Q��Φ����]�C��������C��̤l�]�u�u���u�j���ѩ�2021�~9��20��
Dim d As Document, a, i, p As Paragraph, px As String, rng As Range, e As Long, pRng As Range, pa
'Const corTxt As String = "�׸��I�ե��հɰO��"'�Ӻ����Ϥ��ӱƪ��\�ॼ��t�X�A�G�����ĥΡC��榡�u��奻�����ġChttps://ctext.org/instructions/wiki-formatting/zh
'a = Array(" ", "", "�@�@","","�@", vba.Chrw(-9217) & vba.Chrw(-8195), "^p", "<p>^p",
'a = Array(" ", "", "�@�@", "", "^p^p", "<p>^p" & vba.Chrw(-9217) & vba.Chrw(-8195) & vba.Chrw(-9217) & vba.Chrw(-8195),
a = Array("�@�@", "", "^p", "^p^p", "^p^p^p", "^p^p", "�u^p^p", "�u", "�y^p^p", "�y", "�e^p^p", "�e", "�]^p^p", "�]", _
    "^p^p", "<p>^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195), _
    "^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195) & "�e", _
    "^p{{" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & "{{{�q", _
    "�u<p>^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195), "�u", _
    "�e<p>^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195), "�e", _
    "�y<p>^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195), "�y", _
    "�]<p>^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195), "�]", _
    "����", "�m���ѡn�G", "����", "�m�����n�G", "�i�m�����n�G�z�١j", "�i�m�����n�z�١j�G", "���q", "�m���q�n�G", _
    "�E�{�q", "�E�{", "���]", "���C", "�]��", "�C��", "�w����", "�w���", _
    "��", "��", "�F", VBA.ChrW(24921), "��", VBA.ChrW(21843), _
     VBA.ChrW(-30641), VBA.ChrW(-25066), _
     "�s", VBA.ChrW(32675), "�Y", VBA.ChrW(21373), "��", VBA.ChrW(-30650), _
     "�J", VBA.ChrW(26083), "��", VBA.ChrW(27114), "�@", VBA.ChrW(28433), _
     "��", VBA.ChrW(-30626), _
     "�u", VBA.ChrW(30494), "��", VBA.ChrW(22625), "�M", VBA.ChrW(28152), "�C", VBA.ChrW(-26799), "��", VBA.ChrW(25934), _
    "�m", VBA.ChrW(-28395), "��", VBA.ChrW(-27731), "�V", VBA.ChrW(24892), _
    "�}", VBA.ChrW(24183), "��", VBA.ChrW(23643), "��", VBA.ChrW(-31930), "��", VBA.ChrW(-28471), "�@", VBA.ChrW(31571), _
    "�p", VBA.ChrW(29314), "��", VBA.ChrW(-25811), "��", VBA.ChrW(32220), _
    "�T", VBA.ChrW(20868), "�}", VBA.ChrW(-32486), _
    VBA.ChrW(25995), VBA.ChrW(-24956))
Set d = Documents.Add()
d.Range.Paste
�����w�y�r�Ϩ��N����r d.Range
d.Range.Cut
d.Range.PasteAndFormat wdFormatPlainText
d.Range.text = VBA.Replace(d.Range.text, " ", "")
For i = 0 To UBound(a) - 1
    If a(i) = "^p^p^p" Then
        px = d.Range.text
        Do While InStr(px, VBA.Chr(13) & VBA.Chr(13) & VBA.Chr(13))
            px = Replace(px, VBA.Chr(13) & VBA.Chr(13) & VBA.Chr(13), VBA.Chr(13) & VBA.Chr(13))
        Loop
        d.Range.text = px
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
    px = p.Range.text
    If VBA.Left(px, 7) = "{{" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & "{{{" Then '�`�}�q��
        e = p.Range.Characters(1).End
        rng.SetRange e, e
        rng.MoveEndUntil "�f"
        If rng.Next.Next = "�@" Then rng.Next.Next.Delete
        If InStr(p.Range.text, "�@") Then
            For Each pa In p.Range.Characters
                If pa = "�@" Then
                    pa.text = VBA.ChrW(-9217) & VBA.ChrW(-8195)
                End If
            Next
'            p.Range.text = VBA.Replace(p.Range.text, "�@", vba.Chrw(-9217) & vba.Chrw(-8195))
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
        px = p.Range.text
        If InStr(VBA.Right(px, 4), "<p>") Then
            e = p.Range.Characters(p.Range.Characters.Count - 4).End
        Else
            e = p.Range.Characters(p.Range.Characters.Count - 1).End
        End If
        rng.SetRange e, e
        rng.InsertAfter "}}"
    Else '����q��
        e = p.Range.Characters(1).start
        Set pRng = p.Range
        Do While InStr(pRng.text, "�e")
            rng.SetRange e, e
            rng.MoveEndUntil "�e"
            If rng.Characters(rng.Characters.Count) <> "�^" Then  ' if not correction
                rng.Collapse wdCollapseEnd
                rng.Move , 1
                rng.MoveEnd wdCharacter, 1
                If rng.text Like "[�@�G�T�|�����C�K�E]" Then  ' is footnote No.
                    e = rng.start
                    'rng.Collapse wdCollapseEnd
                    rng.SetRange e - 1, e
                    rng.text = "�@{{{�q"
                    rng.MoveEndUntil "�f"
                    rng.Collapse wdCollapseEnd
                    rng.MoveEnd wdCharacter, 1
                    rng.text = "�r}}}"
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
    If VBA.Left(p.Range.text, 9) = VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195) & "�i�m�����n" Then
        Set rng = p.Range
        p.Range.Characters(1).Delete
        rng.SetRange p.Range.start, p.Range.start
        rng.InsertAfter "{{"
        rng.SetRange p.Range.Characters(p.Range.Characters.Count - 4).End, p.Range.Characters(p.Range.Characters.Count - 4).End
        rng.InsertAfter "}}"
        If Len(rng.Paragraphs(1).Next.Range.text) = 1 Then rng.Paragraphs(1).Next.Range.Delete
    End If
    
    If Len(p.Range) < 20 Then
        If (InStr(p.Range, "�m�v�O�n��") Or VBA.Left(p.Range.text, 3) = "�v�O��") And InStr(p.Range, "*") = 0 Then
            rng.SetRange p.Range.start, p.Range.start
            rng.InsertAfter "*"
            For Each pa In p.Range.Characters
                    If pa Like "[�q�m�n�r]" Or StrComp(pa, VBA.ChrW(-9217) & VBA.ChrW(-8195)) = 0 Then pa.Delete
            Next pa
            '�H�U�覡�|�y��p �ȳQ�]�w���U�@�Ӭq��
'            p.Range.text = VBA.Replace(p.Range.text, vba.Chrw(-9217) & vba.Chrw(-8195) & vba.Chrw(-9217) & vba.Chrw(-8195), "")
'            p.Range.text = VBA.Replace(VBA.Replace(p.Range.text, "�m", ""), "�n", "")
        End If
    End If
    If Len(p.Range) < 25 Then
        If VBA.InStr(p.Range.text, "��") And InStr(p.Range, "*") = 0 _
                And (InStr(p.Range, "����") Or InStr(p.Range, "��") Or InStr(p.Range, "��") _
                Or InStr(p.Range, "�@�a") Or InStr(p.Range, "�C��")) Then
            rng.SetRange p.Range.start, p.Range.start
            rng.InsertAfter "�@*"
            For Each pa In p.Range.Characters
                If pa Like "[�q�m�n�r]" Or StrComp(pa, VBA.ChrW(-9217) & VBA.ChrW(-8195)) = 0 Then pa.Delete
            Next pa
   
'            p.Range.text = VBA.Replace(p.Range.text, vba.Chrw(-9217) & vba.Chrw(-8195) & vba.Chrw(-9217) & vba.Chrw(-8195), "�@*")
'            p.Range.text = VBA.Replace(VBA.Replace(p.Range.text, "�q", ""), "�r", "")
        End If
    End If

Next p
If VBA.Left(d.Paragraphs(1).Range.text, 3) = "�v�O��" And InStr(d.Paragraphs(1).Range.text, "*") = 0 Then
    Set p = d.Paragraphs(1)
    rng.SetRange p.Range.start, p.Range.start
    rng.InsertAfter "*"
'    rng.SetRange p.Range.Characters(p.Range.Characters.Count - 1).End, p.Range.Characters(p.Range.Characters.Count - 1).End
'    rng.InsertAfter "<p>"
End If
If VBA.InStr(d.Paragraphs(2).Range.text, "��") And InStr(d.Paragraphs(2).Range.text, "*") = 0 Then
    Set p = d.Paragraphs(2)
'    rng.SetRange p.Range.start, p.Range.start
'    rng.InsertAfter "�@*"
''    rng.SetRange p.Range.Characters(p.Range.Characters.Count - 1).End, p.Range.Characters(p.Range.Characters.Count - 1).End
''    rng.InsertAfter "<p>"
    p.Range.text = VBA.Replace(p.Range.text, VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195), "�@*")
    Set p = d.Paragraphs(2)
    p.Range.text = VBA.Replace(VBA.Replace(p.Range.text, "�q", ""), "�r", "")
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
word.Application.ActiveWindow.windowState = wdWindowStateMinimize
End Sub
Sub �v�O�T�a�`2old()
'�q2858���_�A20210920:0817����A��λO�v�j�����P�ǧd��@���͡m���ؤ�ƺ��n�ҿ�����|�m�v��n�쥻�A���Τ�����A�M�ܤ֧K��²�Ʀr�ഫ�_�~�γy�r�ýX���x�Z�A���r�ɱ�m�C�ھڪ�@���A�榡�����@�ˡI�ڥ��N�O�q�o�̥X�Ӫ��A�A��²�Ʀr�A�A�S�ϥ��A�y�������áC�����S�Q��Φ����]�C��������C��̤l�]�u�u���u�j���ѩ�2021�~9��20��
Dim d As Document, a, i, p As Paragraph, px As String, rng As Range, e As Long, pRng As Range
'Const corTxt As String = "�׸��I�ե��հɰO��"'�Ӻ����Ϥ��ӱƪ��\�ॼ��t�X�A�G�����ĥΡC��榡�u��奻�����ġChttps://ctext.org/instructions/wiki-formatting/zh
a = Array("^p^p", "<p>^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195), _
    "^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195) & "�e", _
    "^p{{" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & "{{{�q", _
    "�u<p>^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195), "�u", _
    "�e<p>^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195), "�e", _
    "�y<p>^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195), "�y", _
    "�]<p>^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195), "�]", _
    "����", "�m���ѡn�G", "����", "�m�����n�G", "�i�m�����n�G�z�١j", "�i�m�����n�z�١j�G", "���q", "�m���q�n�G", _
    "�E�{�q", "�E�{", "���]", "���C", "�]��", "�C��", "�w����", "�w���", _
    "��", "��", _
     "�s", VBA.ChrW(32675), "�Y", VBA.ChrW(21373), "��", VBA.ChrW(-30650), "�J", VBA.ChrW(26083), "��", VBA.ChrW(-30626), _
     "�u", VBA.ChrW(30494), "��", VBA.ChrW(22625), "�M", VBA.ChrW(28152), "�C", VBA.ChrW(-26799), "��", VBA.ChrW(25934), _
    "�m", VBA.ChrW(-28395), "��", VBA.ChrW(-27731), "�V", VBA.ChrW(24892), "��", VBA.ChrW(23643), "��", VBA.ChrW(27114), _
    "��", VBA.ChrW(-31930), "��", VBA.ChrW(-28471))
Set d = Documents.Add()
d.Range.Paste
For i = 0 To UBound(a) - 1
    d.Range.Find.Execute a(i), , , , , , True, wdFindContinue, , a(i + 1), wdReplaceAll
    i = i + 1
Next i
��r�B�z.�ѦW���g�W���Ъ`
Set rng = Selection.Range
For Each p In d.Paragraphs
    px = p.Range.text
    If VBA.Left(px, 7) = "{{" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & "{{{" Then '�`�}�q��
        e = p.Range.Characters(1).End
        rng.SetRange e, e
        rng.MoveEndUntil "�f"
        'rng.Select
        rng.Collapse wdCollapseEnd
        rng.Select
        Selection.MoveRight wdCharacter, 1, wdExtend
        Selection.TypeText "�r}}}" '�N�`�}�s���e�@�f���k��f�令}}}
        px = p.Range.text
        If InStr(VBA.Right(px, 4), "<p>") Then
            e = p.Range.Characters(p.Range.Characters.Count - 4).End
        Else
            e = p.Range.Characters(p.Range.Characters.Count - 1).End
        End If
        rng.SetRange e, e
        rng.InsertAfter "}}"
    Else '����q��
        e = p.Range.Characters(1).start
        Set pRng = p.Range
        Do While InStr(pRng.text, "�e")
            rng.SetRange e, e
            rng.MoveEndUntil "�e"
            If rng.Characters(rng.Characters.Count) <> "�^" Then  ' if not correction
                rng.Collapse wdCollapseEnd
                rng.Move , 1
                rng.MoveEnd wdCharacter, 1
                If rng.text Like "[�@�G�T�|�����C�K�E]" Then  ' is footnote No.
                    e = rng.start
                    'rng.Collapse wdCollapseEnd
                    rng.SetRange e - 1, e
                    rng.text = "�@{{{�q"
                    rng.MoveEndUntil "�f"
                    rng.Collapse wdCollapseEnd
                    rng.MoveEnd wdCharacter, 1
                    rng.text = "�r}}}"
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
    If VBA.Left(p.Range.text, 9) = VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195) & "�i�m�����n" Then
        Set rng = p.Range
        p.Range.Characters(1).Delete
        rng.SetRange p.Range.start, p.Range.start
        rng.InsertAfter "{{"
        rng.SetRange p.Range.Characters(p.Range.Characters.Count).End, p.Range.Characters(p.Range.Characters.Count).End
        rng.InsertAfter "}}"
    End If
Next p
If VBA.Left(d.Paragraphs(1).Range.text, 3) = "�v�O��" Then
    Set p = d.Paragraphs(1)
    rng.SetRange p.Range.start, p.Range.start
    rng.InsertAfter "*"
    rng.SetRange p.Range.Characters(p.Range.Characters.Count - 1).End, p.Range.Characters(p.Range.Characters.Count - 1).End
    rng.InsertAfter "<p>"
End If
If VBA.InStr(d.Paragraphs(2).Range.text, "��") Then
    Set p = d.Paragraphs(2)
    rng.SetRange p.Range.start, p.Range.start
    rng.InsertAfter "�@*"
'    rng.SetRange p.Range.Characters(p.Range.Characters.Count - 1).End, p.Range.Characters(p.Range.Characters.Count - 1).End
'    rng.InsertAfter "<p>"
    p.Range.text = VBA.Replace(VBA.Replace(p.Range.text, "�q", ""), "�r", "")
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
a = Array("<p>{{{", "<p>^p{{" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & "{{{", _
        "<p>", "<p>^p" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195), _
        VBA.ChrW(-9217) & VBA.ChrW(-8195) & VBA.ChrW(-9217) & VBA.ChrW(-8195) & "^p{{" & VBA.ChrW(-9217) & VBA.ChrW(-8195), _
        "{{" & VBA.ChrW(-9217) & VBA.ChrW(-8195))
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
    px = p.Range.text
    If VBA.Left(p.Range.text, 7) = "{{" & VBA.ChrW(-9217) & VBA.ChrW(-8195) & "{{{" Then
        If InStr(VBA.Right(px, 4), "<p>") Then
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
    .font.Color = 10092543
    .font.Size = 10
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
        rngLast.InsertBefore "{{" & VBA.ChrW(-9217) & VBA.ChrW(-8195)
'        rng.SetRange rng.End + 222, d.Range.End
        
    Loop 'Until InStr(rng, "{{")
    .ClearFormatting
End With
Beep
End Sub

Rem �^�Ǻ��}
Function Search(searchWhatsUrl As String) As String
    Dim d As Document, encode As String
    Set d = ActiveDocument
    If d.path <> "" Then If d.Saved = False Then d.Save
    ��r�B�z.ResetSelectionAvoidSymbols
    If Selection.Type = wdSelectionNormal Then
        Selection.Copy
    End If
    encode = code.UrlEncode(Selection.text)
    'Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe https://ctext.org/wiki.pl?if=gb&res=384378&searchu=" & Selection.text
    'Shell Normal.SystemSetup.getChrome & searchWhatsUrl & Selection.Text
    Shell TextForCtextWordVBA.Network.GetDefaultBrowserEXE & searchWhatsUrl & encode
    Search = searchWhatsUrl & encode
End Function
Rem �˯�CTP�S�w���� ���\�h�Ǧ^true
Function Searchu(res As String, undoName As String) As Boolean
    Dim url As String, ur As UndoRecord, d As Document
    SystemSetup.stopUndo ur, undoName
    'SystemSetup.playSound 0.484
    Set d = Selection.Document
    If d.path <> "" Then If d.Saved = False Then d.Save
    
    ��r�B�z.ResetSelectionAvoidSymbols
    If Selection.Type = wdSelectionNormal Then
        Selection.Copy
    End If
    
    Dim iwe As SeleniumBasic.IWebElement, key As New SeleniumBasic.keys
    If Not SeleniumOP.OpenChrome("https://ctext.org/wiki.pl?if=gb&res=" & res) Then Exit Function
    SeleniumOP.ActivateChrome
    word.Application.windowState = wdWindowStateMinimize
    '�˯���
    Set iwe = SeleniumOP.WD.FindElementByCssSelector("#content > div.wikibox > table > tbody > tr.mobilesearch > td > form > input[type=text]:nth-child(3)")
    If iwe Is Nothing Then Exit Function
    SeleniumOP.SetIWebElementValueProperty iwe, Selection.text
    
    On Error GoTo eH
    iwe.SendKeys key.enter
    '�˯����G
    Set iwe = SeleniumOP.WD.FindElementByCssSelector("#content > table.searchsummary > tbody > tr:nth-child(4) > th > b")
    If iwe Is Nothing Then Exit Function
    If iwe.GetAttribute("textContent") <> "Total 0" Then url = SeleniumOP.WD.url
    If url <> vbNullString Then
        If Selection.Type = wdSelectionIP Then Selection.MoveRight wdCharacter, 1, wdExtend
        ActiveDocument.Hyperlinks.Add Selection.Range, url
    End If
    SystemSetup.contiUndo ur
    Searchu = True
    Exit Function
eH:
    Select Case Err.number
        Case -2146233088
            If VBA.InStr(Err.description, "element not interactable") = 1 Then '(Session info: chrome=130.0.6723.117)
                Set iwe = SeleniumOP.WD.FindElementByCssSelector("#searchform > input.searchbox")
                SeleniumOP.SetIWebElementValueProperty iwe, Selection.text
                Resume
            Else
                GoTo elses
            End If
        Case Else
elses:
            Debug.Print Err.number & Err.description
            MsgBox Err.number & Err.description
    End Select
End Function

Rem 20241006 �HGoogle�˯��m������Ǯѹq�l�ƭp���n Alt + t
Sub SearchSite()
    SeleniumOP.GoogleSearch "site:https://ctext.org/ """ + Selection.text + """"
End Sub
Rem Alt + m �G �H�����r search�v�O�T�a�`�é�����B���J�˯����G���W�s�� �]m=�q���E���� ma�^ 20241014;20241005
'�쬰 Ctrl + s,j �]�o�˪����w�|���������ت� Ctrl + s �A�G��w 20241014
Sub search�v�O�T�a�`()
    Searchu "384378", "search�v�O�T�a�`"
'    Dim ur As UndoRecord
'    SystemSetup.stopUndo ur, "search�v�O�T�a�`"
'    ActiveDocument.Hyperlinks.Add Selection.Range, Search(" https://ctext.org/wiki.pl?if=gb&res=384378&searchu=")
'    SystemSetup.contiUndo ur
End Sub
Rem Ctrl + Alt + = �G �H�������r�˯� CTP �Ҧ������m�Q�T�g�`���P�P�����q�n�æb�����r�W�[�W���˯����G�������W�s��
Sub search�P�����q_�����Q�T�g�`��()
    Searchu "315747", "search�P�����q_�����Q�T�g�`��"
    'url = ������Ǯѹq�l�ƭp��.Search(" https://ctext.org/wiki.pl?if=gb&res=315747&searchu=")
    
End Sub
Rem Ctrl + shift + y �G �H�����r search�m�|���O�Z�n���m�P���n�é�����B���J�˯����G���W�s��(y:yi ��) 20241005
Sub search�P��_�|���O�Z��()
    Searchu "129518", "search�P��_�|���O�Z��"
    'ActiveDocument.Hyperlinks.Add Selection.Range, Search(" https://ctext.org/wiki.pl?if=gb&res=129518&searchu=")
End Sub
Sub Ū�v�O�T�a�`()
    Dim d As Document, t As table
    Set d = Documents.Add
    d.Range.Paste
    Set t = d.tables(1)
    With t
        .Columns(1).Delete
        .ConvertToText wdSeparateByParagraphs
    End With
    d.Range.Cut
    d.Close wdDoNotSaveChanges
    If word.Application.Windows.Count > 0 Then word.Application.ActiveWindow.windowState = wdWindowStateMinimize
End Sub

Sub �԰굦_�|���O�Z_�����w��() '�m�԰굦�n�榡�̬ҾA�Ρ]�Y�D�孺�泻��A�Ө�l���e���@��̡^
'https://ctext.org/library.pl?if=gb&res=77385
    Dim a, rng As Range, rngDoc As Range, p As Paragraph, i As Long, rngCnt As Integer, ok As Boolean
    Dim omits As String
    omits = "�m�n�q�r�u�v�y�z�P" & VBA.Chr(13)
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
                If a.Previous <> VBA.Chr(13) Then a.InsertBefore VBA.Chr(13)
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
                   If rng.Characters(i) = VBA.Chr(13) Then
                        i = 0
                        Exit For
                   End If
                Next a
            Else
                For Each a In rng.Characters
                   i = i + 1
                   If rng.Characters(i) = "}" Then Exit For
                   If rng.Characters(i) = VBA.Chr(13) Or rng.Characters(i) = "{" Then
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
            If VBA.Left(p.Range.text, 3) = "{{�@" And p.Range.Characters(p.Range.Characters.Count - 1) = "}" Then
                a = p.Range.text
                a = VBA.Mid(a, 4, Len(a) - 6)
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
    '    rngDoc.Find.Execute vba.Chrw(-10155) & vba.Chrw(-8585) & "��", , , , , , , wdFindContinue, , "�i" & vba.Chrw(-10155) & vba.Chrw(-8585) & "��j", wdReplaceAll
    '    rngDoc.Find.Execute "�ɤ�", , , , , , , wdFindContinue, , "�i" & vba.Chrw(-10155) & vba.Chrw(-8585) & "��j", wdReplaceAll
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
Sub ���㶰�`�Y��N������p�`�榡_�|�w����_��Ǥj�v()
    Dim d As Document, p As Paragraph, px As String, rng As Range, a As Range, ur As UndoRecord, s As Long, e As Long, sx As String
    Set d = ActiveDocument: Set rng = d.Range
    SystemSetup.stopUndo ur, "���㶰�`�Y��N������p�`�榡_�|�w����_��Ǥj�v"
    For Each p In d.Paragraphs
        px = p.Range.text
        s = VBA.InStr(px, "{{"): e = VBA.InStr(px, "}}" & VBA.Chr(13))
        If e > 0 Then sx = VBA.Mid(px, s + 2, e - s - 2)
        If e > 0 And s > 0 And VBA.InStr(sx, "{{") = 0 And VBA.InStr(sx, "}}") = 0 Then '�e�ᦳ{{}}�A���Ǥ�������A��{{}}
            If s = 1 Then '�p�G�e�L�Y��
                rng.SetRange p.Range.start + 2, p.Range.End - 3
                rng.Characters(Int(rng.Characters.Count / 2)).InsertAfter VBA.Chr(13)
            Else
                If VBA.InStr(px, "�@{{") > 0 Then
                    sx = VBA.Mid(px, 1, s - 1)
                    If Replace(sx, "�@", vbNullString) = vbNullString Then '�p�e�e�󳣬O���ΪŮ�F�Y�Y��
                        rng.SetRange p.Range.start + VBA.Len(sx) + 2, p.Range.End - 3
                        rng.Characters(Int(rng.Characters.Count / 2)).InsertAfter VBA.Chr(13) & VBA.Mid(px, 1, s - 1)
                    End If
                End If
            End If
        End If
        
    Next p
    SystemSetup.contiUndo ur
End Sub

Sub ��������Y��1������p�`�榡_�|�w����_��Ǥj�v()
    Dim d As Document, p As Paragraph, px As String, rng As Range, a As Range, ur As UndoRecord
    Set d = ActiveDocument: Set rng = d.Range
    SystemSetup.stopUndo ur, "��������Y�Ƥ@������p�`�榡_�|�w����_��Ǥj�v"
    For Each p In d.Paragraphs
        px = p.Range.text
        If (VBA.Left(px, 3) = "�@{{" Or VBA.Left(px, 3) = "{{�@") And VBA.Right(px, 3) = "}}" & VBA.Chr(13) Then
            rng.SetRange p.Range.start + 3, p.Range.End - 3
            If VBA.InStr(rng.text, "}") = 0 Then
                If rng.Characters.Count > 1 Then
                    rng.Characters(Int(rng.Characters.Count / 2)).InsertAfter VBA.Chr(13) & "�@"
                Else
                    rng.Characters(1).InsertAfter VBA.Chr(13) & "�@"
                End If
            End If
        ElseIf VBA.Left(px, 3) = "{{�@" And VBA.Right(px, 6) = "}}<p>" & VBA.Chr(13) Then
            rng.SetRange p.Range.start + 3, p.Range.End - 6
            If VBA.InStr(rng.text, "}") = 0 Then
                If InStr(rng.text, "�@") Then
                    For Each a In rng.Characters
                        If a.text = "�@" Then
                            a.InsertBefore VBA.Chr(13)
                        End If
                    Next a
                Else
                    If rng.Characters.Count > 1 Then
                        'Skype Copilot�j���� 20240519
                        rng.Characters(-Int(-(rng.Characters.Count / 2))).InsertAfter VBA.Chr(13) & "�@"
                    Else
                        rng.Characters(1).InsertAfter VBA.Chr(13) & "�@"
                    End If
                End If
            End If
        End If
    Next p
    SystemSetup.contiUndo ur
End Sub

Sub �ɬA��()
    Dim d As Document, rng As Range, p As Paragraph, ur As UndoRecord
    Set d = ActiveDocument: SystemSetup.stopUndo ur, "�ɬA��"
    Set rng = d.Range
    For Each p In d.Paragraphs
        If VBA.Left(p.Range.text, 2) = "{{" And VBA.Right(p.Range.text, 3) <> "}}" & VBA.Chr(13) Then
            If VBA.Right(p.Next.Range.text, 3) = "}}" & VBA.Chr(13) Then
                rng.SetRange p.Range.start, p.Range.End - 1
                rng.text = VBA.Left(p.Range.text, VBA.Len(p.Range.text) - 1) & "}}"
                p.Next.Range.text = "{{" & p.Next.Range.text
            End If
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
        ElseIf aLtTxt Like VBA.ChrW(12272) & VBA.ChrW(-10155) & VBA.ChrW(-8696) & VBA.ChrW(31860) Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "�s�]?�}���^-- ��������" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "?�g? -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "????�f?�� -- " & VBA.ChrW(-10111) & VBA.ChrW(-8620) Then
            aLtTxt = VBA.ChrW(-10111) & VBA.ChrW(-8620)
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
            aLtTxt = VBA.ChrW(-10114) & VBA.ChrW(-9161)
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
        ElseIf aLtTxt Like VBA.ChrW(24298) & "�]" & VBA.ChrW(8220) & VBA.ChrW(13357) & VBA.ChrW(8221) & "����" & VBA.ChrW(8220) & "��" & VBA.ChrW(8221) & "�^" Then
            aLtTxt = "�o"
        ElseIf aLtTxt Like VBA.ChrW(12273) & VBA.ChrW(11966) & VBA.ChrW(30464) Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "?�L? -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like VBA.ChrW(12272) & VBA.ChrW(-10145) & VBA.ChrW(-8265) & "��" Then
            aLtTxt = "����" & aLtTxt & "��"
        ElseIf aLtTxt Like "? -- or ?? ?" Then
            aLtTxt = VBA.ChrW(-32119)
        ElseIf aLtTxt Like "��" Then
            aLtTxt = VBA.ChrW(18518)
        ElseIf aLtTxt Like "��" Then
            aLtTxt = VBA.ChrW(17403)
        ElseIf aLtTxt Like VBA.ChrW(12272) & VBA.ChrW(-10145) & VBA.ChrW(-8265) & VBA.ChrW(25908) Then
            aLtTxt = VBA.ChrW(-10109) & VBA.ChrW(-8699)
        ElseIf aLtTxt Like "??�K -- " & VBA.ChrW(-10170) & VBA.ChrW(-8693) Then
            aLtTxt = VBA.ChrW(-10124) & VBA.ChrW(-9097)
        ElseIf aLtTxt Like VBA.ChrW(12282) & VBA.ChrW(-28746) & "��" Then
            aLtTxt = "�A"
        ElseIf aLtTxt Like "??�H -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "??̱ -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "?????�� -- �~" Then
            aLtTxt = "�~"
        ElseIf aLtTxt Like "???�\ -- " & VBA.ChrW(31762) Then
            aLtTxt = "�y"
        ElseIf aLtTxt Like "????�Z -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "?�¤� -- ?" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like VBA.ChrW(12282) & VBA.ChrW(-28746) & VBA.ChrW(17807) Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "?�ܤ� -- ??" Then
            aLtTxt = "�P"
        ElseIf aLtTxt Like "�]???�k�^" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "�]???�O�^" Then
            aLtTxt = VBA.ChrW(-10174) & VBA.ChrW(-9072)
        ElseIf aLtTxt Like "??? -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "�]???�^-- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "?�إ� -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "???? -- " & VBA.ChrW(-10161) & VBA.ChrW(-8272) Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "?�f? -- �A" Then
            aLtTxt = "�A"
        ElseIf aLtTxt Like "?�f? -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "???? -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "?�j?? -- �}" Then
            aLtTxt = VBA.ChrW(-10158) & VBA.ChrW(-8444)
        ElseIf aLtTxt Like "*page2700-20px-SKQSfont.pdf.jpg*" Then
            aLtTxt = "�@"
        ElseIf aLtTxt Like VBA.ChrW(12273) & VBA.ChrW(11966) & VBA.ChrW(12272) & VBA.ChrW(27701) & VBA.ChrW(20158) Then
            aLtTxt = VBA.ChrW(-10161) & VBA.ChrW(-8915)
        ElseIf aLtTxt Like "???��ۤP?�I? -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "?�ޤ� -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like VBA.ChrW(12272) & "��" & VBA.ChrW(-10170) & VBA.ChrW(-8693) Then
            aLtTxt = VBA.ChrW(-10121) & VBA.ChrW(-8228)
        ElseIf aLtTxt Like "??? -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "??? -- �S" Then
            aLtTxt = "�S"
        ElseIf aLtTxt Like "??�V -- " & VBA.ChrW(-28664) Then
            aLtTxt = "�~"
        ElseIf aLtTxt Like "?���� -- " & VBA.ChrW(-24830) Then
            aLtTxt = VBA.ChrW(-24830)
        ElseIf aLtTxt Like "???????��-- �Z" Then
            aLtTxt = "�Z"
        ElseIf aLtTxt Like "??? -- ��" Then
            aLtTxt = "��"
        ElseIf aLtTxt Like "�]?��?�^" Then
            aLtTxt = VBA.ChrW(-30654)
        ElseIf aLtTxt Like "SKchar" Then
            GoTo nxt
'            aLtTxt = "�e,�u,�~,�T,�V,�j,�|,��,���]2DB7E�^,��,��,��,�|,��,�s,�p"'�l�� �d�r.mdb
        ElseIf aLtTxt Like "SKchar2" Then
            GoTo nxt
'            aLtTxt = "��]7E92�^,��,"'�l�� �d�r.mdb
        Else
            Select Case aLtTxt
                Case VBA.ChrW(12280) & VBA.ChrW(30098) & VBA.ChrW(-28523)
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
Sub �d�{�@�o�N�ץX() '20250318
    Dim db As New dBase, cnt As New ADODB.Connection, rst As New ADODB.Recordset, rstNote As New ADODB.Recordset, note As String, noteMark As String
    Dim d As Document, stPageNum As Integer, endPageNum As Integer, rng As Range, p As Paragraph, followWords As String, rngDup As Range, si As New StringInfo, ur As UndoRecord
    Rem ��1�q�O�l���A��2�q���w����
    Set d = ActiveDocument
    stPageNum = VBA.CInt(d.Range(d.Paragraphs(1).Range.start, d.Paragraphs(1).Range.End - 1).text)
    endPageNum = VBA.CInt(d.Range(d.Paragraphs(2).Range.start, d.Paragraphs(2).Range.End - 1).text)
    
    SystemSetup.stopUndo ur, "�d�{�@�o�N�ץX"
    d.Range.text = vbNullString
    
    db.cnt_�}�o_�d�{�@�o�N cnt
    rst.Open "SELECT ��.��ID, ��.���O FROM (�� LEFT JOIN �g ON ��.��ID = �g.��ID) LEFT JOIN �� ON �g.�gID = ��.�gID " & _
                    "WHERE (((��.��ID)=9327) AND ((��.��) Between " & stPageNum & " And " & endPageNum & ")) " & _
                    "ORDER BY �g.��, �g.�gID, ��.��", cnt, adOpenForwardOnly, adLockReadOnly
    Do Until rst.EOF
        Set p = d.Paragraphs.Add
        Set rng = d.Range(p.Range.start, p.Range.End - 1)
        rng.InsertAfter rst.Fields("���O").Value '�b�M�Φ���k����A�ӽd��N�|�i�}���]�t�s����r�C
        Set rngDup = rng.Duplicate
        '19323:��,�հɰO,�u��    36171:�`
        rstNote.Open "SELECT ��_��.��ID,����.����, ����.����r��, ����.�Ƶ� FROM ��_�� INNER JOIN ���� ON ��_��.��_ID = ����.��_ID " & _
                        "WHERE (((��_��.��ID)=" & rst.Fields("��ID") & ") AND ((��_��.��ID)=36171 Or (��_��.��ID)=19323)) " & _
                        " order by st", cnt, adOpenKeyset, adLockReadOnly
        Do Until rstNote.EOF
            noteMark = rstNote.Fields("����").Value
findnext:
            If rng.Find.Execute(noteMark) Then
                followWords = VBA.Replace(VBA.IIf(VBA.IsNull(rstNote.Fields("����r��").Value), vbNullString, rstNote.Fields("����r��").Value), VBA.Chr(13) & VBA.Chr(10), VBA.Chr(13))
                If followWords <> vbNullString Then
                    Do Until d.Range(rng.End, rng.End + VBA.Len(followWords)).text = followWords
                        If Not rng.Find.Execute(rstNote.Fields("����").Value) Then Exit Do
                        If rng.End + VBA.Len(followWords) >= d.content.End Then GoTo nextRecord
                    Loop
                End If
            End If
            Select Case rstNote.Fields("��ID").Value
                Case 36171 '�`
                    If rng.start > 0 Then
                        If rng.Previous(wdCharacter, 1) = "{" Then
                            rng.SetRange rng.End, rngDup.End
                            GoTo findnext
                        End If
                        rng.text = "{{" & rng.text & "}}"
                    End If
                Case 19323 '��,�հɰO,�u
                    note = rstNote.Fields("�Ƶ�").Value
                    If VBA.InStr(note, "���_�@�@�X��") = 0 Then
                        If rng.start > 0 Then
                            Do Until VBA.InStr("�C�A" & VBA.Chr(13), rng.Previous.text)
                                If rng.Previous.text <> VBA.Chr(13) Then rng.Move wdCharacter, 1
                                If rng.End = rng.Document.content.End - 1 Then Exit Do
                            Loop
                            si.Create noteMark
                            If si.LengthInTextElements > 1 Then
                                rng.InsertAfter "{{{�]�u�u���G" & "�u" & noteMark & "�v�G" & note & "}}}"
                            Else
                                rng.InsertAfter "{{{�]�u�u���G" & noteMark & "�A" & note & "}}}"
                            End If
                        End If
                    End If
            End Select
nextRecord:
            rng.SetRange rngDup.start, rngDup.End
            rstNote.MoveNext
        Loop
        rstNote.Close
        
        rst.MoveNext
    Loop
    
    rst.Close: cnt.Close
    
    ��r�B�z.�ѦW���g�W���Ъ`
    rng.Document.content.Cut '�ŤU�ǳƶK��TextForCtext��textBox1��
    SystemSetup.contiUndo ur
    
    d.ActiveWindow.windowState = wdWindowStateMinimize
    d.Range.InsertParagraphAfter
    d.Range(d.Paragraphs(1).Range.start, d.Paragraphs(1).Range.End - 1).text = endPageNum + 1
    d.Range(d.Paragraphs(2).Range.start, d.Paragraphs(2).Range.End - 1).text = endPageNum + 9
    
    On Error Resume Next
    AppActivate "TextForCtext"
    VBA.DoEvents
    SendKeys "^v", True
    VBA.DoEvents
    
End Sub
Rem �{�b�h��Kanripo.org�� 20250202�j�~�줭
Sub ���ެ�ޤޱo�Ʀr�H��귽���O_�_�ʤ��ެ�ަ������q���()
    Dim rng As Range, noteRng As Range, aNext As Range, aPre As Range, ur As UndoRecord, midNoteRngPos As Byte, midNoteRng As Range, aX As String, a As Range, aSt As Long, aEd As Long
    Dim noteFont As font '�O�U�`��榡�H�ƥ�
    Dim insertX As String, counter As Byte
    Set rng = Documents.Add().Range
    SystemSetup.stopUndo ur, "��Ǥj�v_Kanripo_�|�w���ѥ����"
    SystemSetup.playSound 1
    rng.Paste
    '���ܶK�W�Lê
    SystemSetup.playSound 1 '���K�W�ӮɴN�ܤ[�F�A�᭱�o�@�j�說�l�Ϧӧ� 20230211
    
    With rng.Find
        .font.ColorIndex = 6
    End With
    Set rng = rng.Document.Range
    '�M�����X
    Do While rng.Find.Execute("P", , , , , , True, wdFindContinue)
       rng.Paragraphs(1).Range.Delete
    Loop
    rng.Find.ClearFormatting
    
    Set rng = rng.Document.Range
'    rng.Find.Execute "^p^p", , , , , , , wdFindContinue, , "^p", wdReplaceAll
'    If VBA.InStr(rng.text, VBA.ChrW(160) & "/" & VBA.Chr(11)) Then _
'        rng.Find.Execute VBA.ChrW(160) & "^g" & VBA.Chr(11), , , , , , , wdFindContinue, , VBA.Chr(11), wdReplaceAll 'chr(11)����Ÿ�
'    If VBA.InStr(rng.text, VBA.ChrW(160) & "/" & VBA.Chr(13)) Then _
'        rng.Find.Execute VBA.ChrW(160) & "^g" & VBA.Chr(13), , , , , , , wdFindContinue, , VBA.Chr(13), wdReplaceAll
    
    rng.Find.Execute VBA.Chr(13), , , , , , , wdFindContinue, , VBA.Chr(11), wdReplaceAll
    rng.Find.Execute "^p/", , , , , , , wdFindContinue, , "^p", wdReplaceAll
        
    rng.Find.font.Color = 1310883
    Do While rng.Find.Execute(vbNullString, , , False, , , True, wdFindStop)
        If noteFont Is Nothing Then Set noteFont = rng.font
        Set noteRng = rng '.Document.Range(rng.start, rng.End)
        Do While noteRng.Next.font.Color = 1310883
            noteRng.SetRange noteRng.start, noteRng.Next.End
        Loop
        
'        If InStr(noteRng, "�m�󤧥y") Then Stop 'just for test
        
        Set aNext = noteRng.Characters(noteRng.Characters.Count).Next
        Set aPre = noteRng.Characters(1).Previous
        midNoteRngPos = Excel.RoundUpCustom(noteRng.Characters.Count / 2)
        
        Set midNoteRng = noteRng.Document.Range(noteRng.Characters(VBA.IIf(midNoteRngPos - 1 < 1, 1, midNoteRngPos - 1)).start _
            , noteRng.Characters(VBA.IIf(midNoteRngPos + 1 > noteRng.Characters.Count, noteRng.Characters.Count, midNoteRngPos + 1)).End)
        If midNoteRng.start = noteRng.start And midNoteRng.End = noteRng.End Then
            Set midNoteRng = noteRng
        End If
'        If (aNext.text = VBA.Chr(11) And aPre.text = VBA.Chr(11)) Then
'            If aNext.Previous = "/" Then
'                midNoteRng.text = VBA.Replace(midNoteRng, "/", vbNullString, 1, 1)
'                noteRng.text = "{{" & noteRng.text & "}}"
'            Else
'                midNoteRng.text = VBA.Replace(midNoteRng, "/", VBA.Chr(11), 1, 1)
'                noteRng.text = "{{" & noteRng.text & "}}"
'            End If
'        ElseIf aNext.text = VBA.Chr(13) And aPre.text = VBA.Chr(13) Then
'            If aNext.Previous = "/" Then
'                midNoteRng.text = VBA.Replace(midNoteRng, "/", vbNullString, 1, 1)
'                noteRng.text = "{{" & noteRng.text & "}}"
'            Else
'                midNoteRng.text = VBA.Replace(midNoteRng, "/", VBA.Chr(13), 1, 1)
'                noteRng.text = "{{" & noteRng.text & "}}"
'            End If
'        Else
'            If aNext.text = VBA.Chr(11) Then


'        If InStr(noteRng, "�A/") Then Stop


                '�P�_���L�Y��
                If Not aPre Is Nothing Then
                    Set a = aPre.Document.Range(aPre.start, aPre.End) '�O�UaPre��Ӫ���m
                    If aPre.start > 0 And aPre.text <> VBA.Chr(11) Then
                        Do Until aPre.Previous = VBA.Chr(11)
                            aPre.Move wdCharacter, -1
                            If aPre.start <= 0 Then Exit Do
                        Loop
                    End If
                    If a.start > aPre.start Then 'a =aPre��Ӫ���m
                        a.SetRange aPre.start, a.End
                        aX = a.text '�Y�ƪ��Ů�
                    Else
                        If a.text = aPre.text Then
                            If aPre.text = "�@" Then '���Y��
                                aX = a.text
                            Else
                                aX = vbNullString
'                                SystemSetup.playSound 12, 0
'                                Stop
                            End If
                        Else
                            aX = vbNullString
                        End If
                    End If
                End If
                
'                Dim line As New LineChr11
                
                '�p�G���Y��('aX=�Y�ƪ��Ů�)
                If aX <> vbNullString And VBA.Replace(aX, "�@", vbNullString) = vbNullString Then
                    If noteRng.Next Is Nothing Then '�Ȧb���̥��ݡA�P�U�@�~�P�_�ä�����
'                    If line.LineRange(noteRng).start = noteRng.start And line.LineRange(noteRng).End = noteRng.End Then
                        insertX = VBA.Chr(11) & aX
                    ElseIf noteRng.Next = VBA.Chr(11) Then 'ax=�Y�ƪ��Ů� ��������������������������
                        insertX = VBA.Chr(11) & aX  'VBA.Chr(11) �᭱ a.text = "}}" & VBA.Replace(insertX, VBA.Chr(11), VBA.Chr(11) & "{{") �n�ѷ�
                    Else
                        If VBA.InStr(midNoteRng.text, "/") _
                            And noteRng.Next.font.Size > 11.5 _
                            And (noteRng.Next.text <> VBA.Chr(11) Or noteRng.Next.text = "�@") Then  '�Y�O���`(�q�`�O���D�U�����`�]�h�᭱���Ů�^�A�p https://ctext.org/library.pl?if=en&file=55677&page=6�^ 20250205
                            'noteRng.Next.text <> VBA.Chr(11):�᭱�٦���r�A�h�����` 20250223��
                            insertX = aX '�ɪŮ�H�Y��
                        Else
                            insertX = vbNullString
                        End If
                    End If
                Else '�S���Y��
                    If aX = vbNullString And Not noteRng.Previous(wdCharacter, 1) Is Nothing And Not noteRng.Next(wdCharacter, 1) Is Nothing Then
                        If noteRng.Previous(wdCharacter, 1) = VBA.Chr(11) And noteRng.Next(wdCharacter, 1) = VBA.Chr(11) Then
                            insertX = VBA.Chr(11)
                        Else
'                            SystemSetup.playSound 7, 0
'                            Stop
                            insertX = vbNullString
                        End If
                    Else
                        insertX = vbNullString
                    End If
                End If
                
                
                For Each a In noteRng.Characters '���/�]���`����^����m
                    If a = "/" And a.InlineShapes.Count = 0 Then
                        If a.font.Color = noteFont.Color And a.font.Size = noteFont.Size Then
                            aSt = a.start
                            aEd = a.End
                            
                            Do Until VBA.Abs(noteRng.Document.Range(noteRng.start, a.start).Characters.Count - VBA.IIf(a.End = noteRng.End, 0, noteRng.Document.Range(a.End, noteRng.End).Characters.Count)) < 2
                               'noteRng.Document.Range(a.End, noteRng.End).text = noteRng.Document.Range(a.End, noteRng.End).text & "�@"
                               noteRng.text = noteRng.text & "�@"
                               a.SetRange aSt, aEd
                               If rng.End + 3 >= rng.Document.Range.End Then Exit Do
                               counter = counter + 1
                               If counter > 50 Then Exit Do
                            Loop
                            counter = 0
                            If a.Next = VBA.Chr(11) Then '�p�G�׽u/�᭱�Y����
                                If aX = vbNullString Or VBA.Replace(aX, "�@", vbNullString) <> vbNullString Then '�Y�L�Y�ơA�h�M�����׽u/
                                    a.text = vbNullString
                                Else '���Y�Ʈ�
                                    a.text = insertX '���������������������������A�[��
                                    noteRng.SetRange aSt, aEd + VBA.Len(insertX) - 1 '�u/�v�] a = "/" �^�����F�G��1
                                End If
                            Else
                                If noteRng.Next = VBA.Chr(11) And aX <> vbNullString And VBA.Replace(aX, "�@", vbNullString) = vbNullString Then
                                'If noteRng.Next = VBA.Chr(11) And VBA.Replace(aPre.text, "�@", vbNullString) = vbNullString Then
                                    If aPre.Previous = VBA.Chr(11) Then
                                        noteRng.SetRange aPre.start, noteRng.End
                                        a.text = "}}" & VBA.Replace(insertX, VBA.Chr(11), VBA.Chr(11) & "{{")
                                    Else
                                        SystemSetup.playSound 12, 0
                                        Stop
                                    End If
                                Else
                                    a.text = insertX
                                End If
                            End If
                            Exit For
                        End If
                    End If
                Next a
                If insertX <> vbNullString And VBA.Replace(insertX, "�@", vbNullString) <> vbNullString Then '�p�G�m���u/�v���r�Ť��O�Ŧr��]���O�Y�ƥΪ��Ů�
                    If aX <> vbNullString Then
                        aSt = noteRng.start
                        noteRng.SetRange aPre.start, noteRng.End
                    End If
                    noteRng.text = "{{" & noteRng.text & "}}"
                    noteRng.Collapse wdCollapseEnd
                Else
'                   midNoteRng.text = VBA.Replace(midNoteRng, "/", vbNullString, 1, 1)
                    If aX <> vbNullString And VBA.Replace(aX, "�@", vbNullString) = vbNullString Then '������������������������
                        '�p�G���Y�ơA�h�X�inoteRng�ܫe����ΪŮ檺���
                        noteRng.MoveStartWhile "�@", -50
                        '�p�G���`���S���Y�ƸɤW���Ů�
                        If a.text <> "�@" Then
                            noteRng.MoveEndWhile "�@", 50
                        End If
                        noteRng.InsertBefore "{{"
                        noteRng.InsertAfter "}}"
                        rng.SetRange rng.End, rng.End
                    Else
                        noteRng.text = "{{" & noteRng.text & "}}"
                    End If
                    
                End If
'            Else
'                midNoteRng.text = VBA.Replace(midNoteRng, "/", vbNullString, 1, 1)
'                noteRng.text = "{{" & noteRng.text & "}}"
'            End If
'        End If
    Loop
    
'    'word.Application.Activate'�b�I������Word�]�Y����Word�^�ɤ���p���A�|�X��
'    SystemSetup.playSound 3
'    If VBA.MsgBox("�Ů��ন�ťաH", vbOKCancel + vbExclamation) = vbOK Then
'        ��Ǥj�v_Kanripo_�|�w���ѥ����_Sub rng.Document.Content
'    End If
    
    SystemSetup.playSound 1
    '��r�B�z.�ѦW���g�W���Ъ` '���浹TextForCtext C#�Ӽ� 20250312
    
    With rng.Document
        With .Range.Find
            .ClearFormatting
    '        .Text = vba.Chrw(9675)
            .text = "}}{{�@}}"
    '        .Replacement.Text = vba.Chrw(12295)
            .Replacement.text = "�@}}"
            .Execute , , , , , , True, wdFindContinue, , , wdReplaceAll
        End With
        .Range.text = Replace(Replace(.Range.text, Chr(11), Chr(13) & Chr(10)), "?", "/")
        .Range.Cut
        
'        If VBA.InStr(.Range.text, "{{}}") Then
'            SystemSetup.playSound 12, 0
'        End If

'        SystemSetup.ClipboardPutIn Replace(Replace(.Range.text, Chr(11), Chr(13) & Chr(10)), "?", "/")
        DoEvents
        If .Application.Visible Then .Application.windowState = wdWindowStateMinimize
        .Close wdDoNotSaveChanges
        
    End With
    SystemSetup.playSound 1.921
    SystemSetup.contiUndo ur
    
    
'    AppActivate "TextForCtext"
'    DoEvents
'    SendKeys "^v"
'    DoEvents
End Sub
Rem �{�b�h��Kanripo.org�� 20250202�j�~�줭
Sub ��Ǥj�v_Kanripo_�|�w���ѥ����()
    Dim rng As Range, noteRng As Range, aNext As Range, aPre As Range, ur As UndoRecord, midNoteRngPos As Byte, midNoteRng As Range, aX As String, a As Range, aSt As Long, aEd As Long
    Dim noteFont As font '�O�U�`��榡�H�ƥ�
    Dim insertX As String
    Set rng = Documents.Add().Range
    SystemSetup.stopUndo ur, "��Ǥj�v_Kanripo_�|�w���ѥ����"
    SystemSetup.playSound 1
    
    'P �D�u�_�ʤ��ެ�ަ������q�m���ެ�ޤޱo�Ʀr�H��귽���O�P������N���m�n�v���奻�S�x
    If VBA.InStr(SystemSetup.GetClipboard, "P") Then
        ���ެ�ޤޱo�Ʀr�H��귽���O_�_�ʤ��ެ�ަ������q���
        Exit Sub
    End If
    
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
        .Execute "[[]*[]]  ", , , True, , , True, wdFindContinue, , vbNullString, wdReplaceAll
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
    If VBA.InStr(rng.text, VBA.ChrW(160) & "/" & VBA.Chr(11)) Then _
        rng.Find.Execute VBA.ChrW(160) & "^g" & VBA.Chr(11), , , , , , , wdFindContinue, , VBA.Chr(11), wdReplaceAll 'chr(11)����Ÿ�
    If VBA.InStr(rng.text, VBA.ChrW(160) & "/" & VBA.Chr(13)) Then _
        rng.Find.Execute VBA.ChrW(160) & "^g" & VBA.Chr(13), , , , , , , wdFindContinue, , VBA.Chr(13), wdReplaceAll
        
    rng.Find.ClearFormatting
    
    
    rng.Find.font.Color = 16711935
    Do While rng.Find.Execute(vbNullString, , , False, , , True, wdFindStop)
        If noteFont Is Nothing Then Set noteFont = rng.font
        Set noteRng = rng '.Document.Range(rng.start, rng.End)
        Do While noteRng.Next.font.Color = 16711935
            noteRng.SetRange noteRng.start, noteRng.Next.End
        Loop
        
'        If InStr(noteRng, "�m�󤧥y") Then Stop 'just for test
        
        Set aNext = noteRng.Characters(noteRng.Characters.Count).Next
        Set aPre = noteRng.Characters(1).Previous
        midNoteRngPos = Excel.RoundUpCustom(noteRng.Characters.Count / 2)
        
        Set midNoteRng = noteRng.Document.Range(noteRng.Characters(VBA.IIf(midNoteRngPos - 1 < 1, 1, midNoteRngPos - 1)).start _
            , noteRng.Characters(VBA.IIf(midNoteRngPos + 1 > noteRng.Characters.Count, noteRng.Characters.Count, midNoteRngPos + 1)).End)
        If midNoteRng.start = noteRng.start And midNoteRng.End = noteRng.End Then
            Set midNoteRng = noteRng
        End If
'        If (aNext.text = VBA.Chr(11) And aPre.text = VBA.Chr(11)) Then
'            If aNext.Previous = "/" Then
'                midNoteRng.text = VBA.Replace(midNoteRng, "/", vbNullString, 1, 1)
'                noteRng.text = "{{" & noteRng.text & "}}"
'            Else
'                midNoteRng.text = VBA.Replace(midNoteRng, "/", VBA.Chr(11), 1, 1)
'                noteRng.text = "{{" & noteRng.text & "}}"
'            End If
'        ElseIf aNext.text = VBA.Chr(13) And aPre.text = VBA.Chr(13) Then
'            If aNext.Previous = "/" Then
'                midNoteRng.text = VBA.Replace(midNoteRng, "/", vbNullString, 1, 1)
'                noteRng.text = "{{" & noteRng.text & "}}"
'            Else
'                midNoteRng.text = VBA.Replace(midNoteRng, "/", VBA.Chr(13), 1, 1)
'                noteRng.text = "{{" & noteRng.text & "}}"
'            End If
'        Else
'            If aNext.text = VBA.Chr(11) Then


'        If InStr(noteRng, "�A/") Then Stop


                '�P�_���L�Y��
                If Not aPre Is Nothing Then
                    Set a = aPre.Document.Range(aPre.start, aPre.End) '�O�UaPre��Ӫ���m
                    If aPre.start > 0 And aPre.text <> VBA.Chr(11) Then
                        Do Until aPre.Previous = VBA.Chr(11)
                            aPre.Move wdCharacter, -1
                            If aPre.start <= 0 Then Exit Do
                        Loop
                    End If
                    If a.start > aPre.start Then 'a =aPre��Ӫ���m
                        a.SetRange aPre.start, a.End
                        aX = a.text '�Y�ƪ��Ů�
                    Else
                        If a.text = aPre.text Then
                            If aPre.text = "�@" Then '���Y��
                                aX = a.text
                            Else
                                aX = vbNullString
'                                SystemSetup.playSound 12, 0
'                                Stop
                            End If
                        Else
                            aX = vbNullString
                        End If
                    End If
                End If
                
'                Dim line As New LineChr11
                
                '�p�G���Y��('aX=�Y�ƪ��Ů�)
                If aX <> vbNullString And VBA.Replace(aX, "�@", vbNullString) = vbNullString Then
                    If noteRng.Next Is Nothing Then '�Ȧb���̥��ݡA�P�U�@�~�P�_�ä�����
'                    If line.LineRange(noteRng).start = noteRng.start And line.LineRange(noteRng).End = noteRng.End Then
                        insertX = VBA.Chr(11) & aX
                    ElseIf noteRng.Next = VBA.Chr(11) Then 'ax=�Y�ƪ��Ů� ��������������������������
                        insertX = VBA.Chr(11) & aX  'VBA.Chr(11) �᭱ a.text = "}}" & VBA.Replace(insertX, VBA.Chr(11), VBA.Chr(11) & "{{") �n�ѷ�
                    Else
                        If VBA.InStr(midNoteRng.text, "/") _
                            And noteRng.Next.font.Size > 11.5 _
                            And (noteRng.Next.text <> VBA.Chr(11) Or noteRng.Next.text = "�@") Then  '�Y�O���`(�q�`�O���D�U�����`�]�h�᭱���Ů�^�A�p https://ctext.org/library.pl?if=en&file=55677&page=6�^ 20250205
                            'noteRng.Next.text <> VBA.Chr(11):�᭱�٦���r�A�h�����` 20250223��
                            insertX = aX '�ɪŮ�H�Y��
                        Else
                            insertX = vbNullString
                        End If
                    End If
                Else '�S���Y��
                    If aX = vbNullString And Not noteRng.Previous(wdCharacter, 1) Is Nothing And Not noteRng.Next(wdCharacter, 1) Is Nothing Then
                        If noteRng.Previous(wdCharacter, 1) = VBA.Chr(11) And noteRng.Next(wdCharacter, 1) = VBA.Chr(11) Then
                            insertX = VBA.Chr(11)
                        Else
'                            SystemSetup.playSound 7, 0
'                            Stop
                            insertX = vbNullString
                        End If
                    Else
                        insertX = vbNullString
                    End If
                End If
                
                
                For Each a In noteRng.Characters '���/�]���`����^����m
                    If a = "/" And a.InlineShapes.Count = 0 Then
                        If a.font.Color = noteFont.Color And a.font.Size = noteFont.Size Then
                            aSt = a.start
                            aEd = a.End
                            
                            Do Until VBA.Abs(noteRng.Document.Range(noteRng.start, a.start).Characters.Count - VBA.IIf(a.End = noteRng.End, 0, noteRng.Document.Range(a.End, noteRng.End).Characters.Count)) < 2
                               'noteRng.Document.Range(a.End, noteRng.End).text = noteRng.Document.Range(a.End, noteRng.End).text & "�@"
                               noteRng.text = noteRng.text & "�@"
                               a.SetRange aSt, aEd
                            Loop
                            If a.Next = VBA.Chr(11) Then '�p�G�׽u/�᭱�Y����
                                If aX = vbNullString Or VBA.Replace(aX, "�@", vbNullString) <> vbNullString Then '�Y�L�Y�ơA�h�M�����׽u/
                                    a.text = vbNullString
                                Else '���Y�Ʈ�
                                    a.text = insertX '���������������������������A�[��
                                    noteRng.SetRange aSt, aEd + VBA.Len(insertX) - 1 '�u/�v�] a = "/" �^�����F�G��1
                                End If
                            Else
                                If noteRng.Next = VBA.Chr(11) And aX <> vbNullString And VBA.Replace(aX, "�@", vbNullString) = vbNullString Then
                                'If noteRng.Next = VBA.Chr(11) And VBA.Replace(aPre.text, "�@", vbNullString) = vbNullString Then
                                    If aPre.Previous = VBA.Chr(11) Then
                                        noteRng.SetRange aPre.start, noteRng.End
                                        a.text = "}}" & VBA.Replace(insertX, VBA.Chr(11), VBA.Chr(11) & "{{")
                                    Else
                                        SystemSetup.playSound 12, 0
                                        Stop
                                    End If
                                Else
                                    a.text = insertX
                                End If
                            End If
                            Exit For
                        End If
                    End If
                Next a
                If insertX <> vbNullString And VBA.Replace(insertX, "�@", vbNullString) <> vbNullString Then '�p�G�m���u/�v���r�Ť��O�Ŧr��]���O�Y�ƥΪ��Ů�
                    If aX <> vbNullString Then
                        aSt = noteRng.start
                        noteRng.SetRange aPre.start, noteRng.End
                    End If
                    noteRng.text = "{{" & noteRng.text & "}}"
                    noteRng.Collapse wdCollapseEnd
                Else
'                   midNoteRng.text = VBA.Replace(midNoteRng, "/", vbNullString, 1, 1)
                    If aX <> vbNullString And VBA.Replace(aX, "�@", vbNullString) = vbNullString Then '������������������������
                        '�p�G���Y�ơA�h�X�inoteRng�ܫe����ΪŮ檺���
                        noteRng.MoveStartWhile "�@", -50
                        If Not a Is Nothing Then
                        '�p�G���`���S���Y�ƸɤW���Ů�
                            If a.text <> "�@" Then
                                noteRng.MoveEndWhile "�@", 50
                            End If
                        End If
                        noteRng.InsertBefore "{{"
                        noteRng.InsertAfter "}}"
                        rng.SetRange rng.End, rng.End
                    Else
                        noteRng.text = "{{" & noteRng.text & "}}"
                    End If
                    
                End If
'            Else
'                midNoteRng.text = VBA.Replace(midNoteRng, "/", vbNullString, 1, 1)
'                noteRng.text = "{{" & noteRng.text & "}}"
'            End If
'        End If
    Loop
    
'    'word.Application.Activate'�b�I������Word�]�Y����Word�^�ɤ���p���A�|�X��
'    SystemSetup.playSound 3
'    If VBA.MsgBox("�Ů��ন�ťաH", vbOKCancel + vbExclamation) = vbOK Then
'        ��Ǥj�v_Kanripo_�|�w���ѥ����_Sub rng.Document.Content
'    End If
    
    SystemSetup.playSound 1
    '��r�B�z.�ѦW���g�W���Ъ` '���浹TextForCtext C#�Ӽ� 20250312
    
    With rng.Document
        With .Range.Find
            .ClearFormatting
    '        .Text = vba.Chrw(9675)
            .text = "}}{{�@}}"
    '        .Replacement.Text = vba.Chrw(12295)
            .Replacement.text = "�@}}"
            .Execute , , , , , , True, wdFindContinue, , , wdReplaceAll
        End With
        '.Range.Cut
        
'        If VBA.InStr(.Range.text, "{{}}") Then
'            SystemSetup.playSound 12, 0
'        End If

        SystemSetup.ClipboardPutIn .Range.text
        DoEvents
        .Close wdDoNotSaveChanges
    End With
    SystemSetup.playSound 1.921
    SystemSetup.contiUndo ur
End Sub
Rem �@�� ��Ǥj�v_�|�w���ѥ����()���l�{��:1.���N�Ů欰�ť�
Sub ��Ǥj�v_Kanripo_�|�w���ѥ����_Sub(rng As Range)
    Dim rngEd As Long, rngChk As Range, rngChkX As String, rngChkPre As Range
    Do While rng.Find.Execute("�@")
        rngEd = rng.End
        Set rngChk = rng.Document.Range(rng.start, rng.End)
        rngChk.SetRange rng.start + rngChk.MoveStartUntil(VBA.Chr(11), -(rng.End - 1)) + 1, rngEd
        rngChkX = VBA.Replace(rngChk.text, "�@", vbNullString)
        Set rngChkPre = rng.Previous
        If rngChkX <> vbNullString And rngChkX <> "{{" Then
            If rng.Previous.text <> VBA.Chr(11) Then
                GoSub replaceSpaceWithBlank:
            End If
        ElseIf Not rngChkPre Is Nothing Then
            If rngChkPre.text <> VBA.Chr(11) And VBA.Left(rngChk, 2) <> "{{" Then GoSub replaceSpaceWithBlank:
        End If
        rng.SetRange rngEd, rng.Document.content.End
    Loop
    
Exit Sub
replaceSpaceWithBlank:
    rngChk.SetRange rng.start, rng.End
    Dim line As New LineChr11
    If VBA.InStr(rngChk.Document.Range(rngChk.End, line.EndPosition(rngChk)).text, "}") Then
        'rngChk.MoveEndUntil "}", line.LineRange(rngChk).End - rng.End
        rngChk.MoveEndUntil "}", line.EndPosition(rngChk) - rng.End
        rngChkX = rngChk.text
        If VBA.Replace(rngChkX, "�@", vbNullString) <> vbNullString Then
            rngEd = rng.End
            rng.text = VBA.ChrW(-9217) & VBA.ChrW(-8195) '���N�Ů欰�ť�
        End If
    Else
        rngEd = rng.End
        rng.text = VBA.ChrW(-9217) & VBA.ChrW(-8195) '���N�Ů欰�ť�
    End If
    Return
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
            exportStr = exportStr & VBA.Chr(13) & "*" & title & VBA.Chr(13)
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
        rng.text = Replace(rng, e, "")
    Next e
    rng.text = "#" & rng.text
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
cde = VBA.Mid(lnk, InStr(lnk, keys) + Len(keys))
cde = code.URLDecode(cde)
s = Selection.start
SystemSetup.stopUndo ur, "���J�W�s��_�N��ܤ��s�X�אּ����"
With Selection
    .Hyperlinks.Add Selection.Range, lnk, , , VBA.Left(lnk, InStr(lnk, keys) + Len(keys) - 1) + cde
    'd.Range(Selection.End, Selection.End + Len(cde)).Select
    'Selection.Collapse
    .MoveLeft wdCharacter, Len(cde)
    .MoveRight wdCharacter, Len(cde) - 1, wdExtend
    .Range.HighlightColorIndex = wdYellow
    .Move , 2
    .InsertParagraphAfter
    .InsertParagraphAfter
    .Collapse
End With
SystemSetup.contiUndo ur
End Sub
Sub �u�O�d����`��_�B�`��e��[�A��(d As Document)
    Dim ur As UndoRecord, slRng As Range
    SystemSetup.stopUndo ur, "������Ǯѹq�l�ƭp��_�u�O�d����`��_�B�`��e��[�A��"
'    Set d = Docs.�ťժ��s���()
'    Set d = ActiveDocument
'    d.Activate
    d.Range.Paste
    'If Selection.Type = wdSelectionIP Then ActiveDocument.Select
'    Set slRng = Selection.Range
    Set slRng = d.Range
    '�Ϥ��ӡAQuict edit�A�歶���� 20240912
    If InStr(slRng.text, "<p>") Or (InStr(slRng.text, "{") And InStr(slRng.text, "}")) Then
        If InStr(slRng.text, "{{{") Then
            slRng.Find.ClearAllFuzzyOptions: slRng.Find.ClearFormatting
            slRng.Find.Execute "{{{*}}}", , , True, , , True, wdFindContinue, , vbNullString, wdReplaceAll
        End If
        slRng.Find.ClearAllFuzzyOptions: slRng.Find.ClearFormatting
        slRng.Find.Execute "^p", , , , , , , , , vbNullString, wdReplaceAll
        slRng.Find.Execute "<p>", , , , , , , , , vbNullString, wdReplaceAll
        slRng.Find.Execute "{{", , , , , , , , , "�]", wdReplaceAll
        slRng.Find.Execute "}}", , , , , , , , , "�^", wdReplaceAll
    Else 'Edit�BView�A�g���`��쭶��
        �M���奻�������s���x�s�� slRng
        ������Ǯѹq�l�ƭp��_������r slRng
        Dim ay, e
        ay = Array(254, 8912896)
        With d.Range.Find
            .ClearFormatting
        End With
        For Each e In ay
            With d.Range.Find
                .font.Color = e
                .Execute "", , , , , , True, wdFindContinue, , "", wdReplaceAll
            End With
        Next e
        Set slRng = d.Range
        With slRng.Find
            .ClearFormatting
            .font.Color = 34816
        End With
        Do While slRng.Find.Execute(, , , , , , True, wdFindStop)
            If InStr(VBA.Chr(13) & VBA.Chr(11) & VBA.Chr(7) & VBA.Chr(8) & VBA.Chr(9) & VBA.Chr(10), slRng) = 0 Then
            slRng.text = "�]" + slRng.text + "�^"
            'slRng.SetRange slRng.End, d.Range.End
            End If
        Loop
    End If
    SystemSetup.contiUndo ur
End Sub

Sub �M���奻�������s���x�s��(rng As Range)
    Dim c As cell, cx As String, t As table
    For Each t In rng.tables
        For Each c In t.Range.cells
            c.Select
            cx = c.Range.text
            If VBA.IsNumeric(VBA.Left(cx, 1)) And VBA.InStr(cx, VBA.ChrW(160) & VBA.ChrW(47)) > 0 And c.Range.InlineShapes.Count = 1 And VBA.Len(cx) < 13 Then
                If VBA.InStr(cx, VBA.Val(cx) & VBA.ChrW(160) & VBA.ChrW(47)) = 1 Then
                    c.Delete
                End If
            End If
        Next c
    Next t
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
        If InStr(rng.text, rst.Fields(0).Value) Then _
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
            Selection.Move
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
        ay(i) = VBA.Split(rng.text, ">")
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
                rngMark.Move wdCharacter, -1
            Loop
            rngMark.SetRange rngMark.start, rng.start
            If VBA.Left(rngMark.text, 8) <> "<entity " Then
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
        rng.Find.font.Color = 8912896 '{{{}}}�y�k�U����r
        rng.Find.Replacement.ClearFormatting
        With rng.Find
            .text = ""
            .Replacement.text = ""
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
            If InStr(rng.text, e) > 0 Then
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
    'If Not VBA.IsNumeric(VBA.Replace(d.Range.Paragraphs(1).Range.text, vba.Chr(13), "")) then
    If VBA.Replace(d.Paragraphs(1).Range + d.Paragraphs(2).Range + d.Paragraphs(3).Range, VBA.Chr(13), "") = "" _
        Or Not IsNumeric(VBA.Replace(d.Paragraphs(1).Range, VBA.Chr(13), "")) Then
        If Not IsNumeric(VBA.Replace(VBA.Replace(d.Paragraphs(1).Range + d.Paragraphs(2).Range + d.Paragraphs(3).Range, VBA.Chr(13), ""), "-", vbNullString, 1, 1)) Then
            MsgBox "�Цb���e3�q���O�O�H�U��T�]�ҬO�Ʀr�^,���槹�|�M��" & vbCr & vbCr & _
                "1. ���Ʈt(�ӷ�-(��h)�ت��^�C�L���t�h��0�A�ٲ��h�w�]��0" & vbCr & vbCr & _
                 "�]�i��J�u�ӷ�-�ت��v�o�˪��榡�A�p�ӷ�114���A�ت��O69���A�i�H�u114-69�v���" & vbCr & _
                "2. �ت��� file number�C�n�m�������F�����N�h��0�A�ٲ��h�w�]��0" & vbCr & _
                "3. �ӷ��� file number�A�n�Q���N��,�ٲ��]���n�Ũ�q��=�Ŧ�^�h����󤤪�file=�᪺��"
            Exit Sub
        End If
    End If
    Dim differPageNum  As Integer '���Ʈt(�ӷ�-(��h)�ت��^
    Dim numRngDashPost As Byte, numRng As Range
    numRngDashPost = VBA.InStr(d.Paragraphs(1).Range.text, "-")
    If numRngDashPost > 1 Then '=1 �O�t�Ƽ���
        Set numRng = d.Range(d.Paragraphs(1).Range.Characters(1).start, d.Paragraphs(1).Range.Characters(d.Paragraphs(1).Range.Characters.Count).start)
        numRng.text = VBA.CInt(VBA.Left(numRng.text, numRngDashPost - 1)) - VBA.CInt(Mid(numRng.text, numRngDashPost + 1))
    End If
    differPageNum = VBA.IIf(d.Paragraphs(1).Range.Characters.Count = 1, 0, VBA.Replace(d.Paragraphs(1).Range.text, VBA.Chr(13), "")) '���Ʈt(�ӷ�-(��h)�ت��^
    Dim file
    file = VBA.Replace(d.Paragraphs(2).Range.text, VBA.Chr(13), "") ' �ت��C�����N�h��0
    If file = "" Then file = 0
    Dim fileFrom As String
    fileFrom = VBA.Replace(d.Paragraphs(3).Range.text, VBA.Chr(13), "") ' '�ӷ�
    If fileFrom = "" Then
        Dim s As String: s = VBA.InStr(d.Range.text, "<scanbegin file="): s = s + VBA.Len("<scanbegin file=")
        fileFrom = VBA.Mid(d.Range.text, s + 1, InStr(s + 1, d.Range.text, """") - s - 1)
    End If
    Set rng = d.Range
    'Set ur = SystemSetup.stopUndo("EditMakeupCtext")
    SystemSetup.stopUndo ur, "EditMakeupCtext"
    If file > 0 Then
        'rng.Find.Execute " file=""77991""", True, True, , , , True, wdFindContinue, , " file=""" & file & """", wdReplaceAll
        rng.text = Replace(rng.text, " file=""" & fileFrom & """", " file=""" & file & """")
    End If

    Do While rng.Find.Execute(" page=""", , , , , , True, wdFindStop)
        Set pageNum = rng
        pageNum.SetRange rng.End, rng.End + 1
        pageNum.MoveEndUntil """"
        pageNum.text = CStr(CInt(pageNum.text) - differPageNum)
        rng.SetRange pageNum.End, d.Range.End
    Loop
    rng.SetRange d.Range.Paragraphs(1).Range.start, d.Range.Paragraphs(3).Range.End
    rng.Delete
    'd.Range.Cut
    SystemSetup.SetClipboard d.Range.text
    SystemSetup.contiUndo ur
    SystemSetup.playSound 1
'    d.Application.Activate
End Sub
Property Get Div_generic_IncludePathAndEndPageNum() As SeleniumBasic.IWebElement
    Dim iwe As SeleniumBasic.IWebElement
    'if Form1.IsValidUrl��ImageTextComparisonPage(ActiveForm1.textBox3Text))
    Set iwe = WD.FindElementByCssSelector("#content > div:nth-child(3)")
    Set Div_generic_IncludePathAndEndPageNum = iwe
End Property
Rem ���o�Y�ѥU�����W��
Property Get pageUBound() As Integer
    Dim iwe  As SeleniumBasic.IWebElement, str As String
    Set iwe = Div_generic_IncludePathAndEndPageNum
    If iwe Is Nothing Then pageUBound = 0
    str = iwe.GetAttribute("textContent") '"�u�W�Ϯ��] -> �Q�Ϥp�� -> �Q�Ϥp���T  /117 ";
    pageUBound = VBA.CInt(VBA.Mid(str, VBA.InStr(str, "/") + 1, VBA.Len(str) - 1 - VBA.InStr(str, "/")))
End Property
Function CurrentChapterNum_Selector() As String
    Dim selector As String
    Dim match As Object
    Dim regex As Object
    
    ' �]�w��ܾ��r�Ŧ�
    selector = ChapterSelector '"#content > div:nth-child(6) > table > tbody > tr:nth-child(2) > td:nth-child(1) > a"
    
    ' �إߥ��h��F����H
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "tr:nth-child\((\d+)\)"
    regex.Global = False
    
    ' �i��ǰt
    Set match = regex.Execute(selector)
    If match.Count > 0 Then
        ' ���o�ǰt���s�խ�
        CurrentChapterNum_Selector = match(0).SubMatches(0)
    Else
        ' �Y�L�ǰt�A��^�Ŧr��
        CurrentChapterNum_Selector = ""
    End If
End Function
Function IncrementNthChild(selector As String) As String
    Dim regex As Object
    Dim match As Object
    Dim number As Integer
    
    ' �إߥ��h��F����H
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "tr:nth-child\((\d+)\)"
    regex.Global = False
    
    ' ����ǰt
    Set match = regex.Execute(selector)
    If match.Count > 0 Then
        ' ���o�s�դ����Ʀr�A���ର���
        number = CInt(match(0).SubMatches(0))
        number = number + 1
        
        ' �ϥΥ��h��F����������s�᪺��
        IncrementNthChild = regex.Replace(selector, "tr:nth-child(" & number & ")")
    Else
        ' �p�G�ǰt���ѡA��^��l�r�Ŧ�
        IncrementNthChild = selector
    End If
End Function

Function NextChapterSelector(ChapterSelector As String) As String
    'Static ChapterSelector As String
    Dim selector As String
    Dim newSelector As String
    
    ' �ˬd ChapterSelector �O�_����
    If IsEmpty(ChapterSelector) Or ChapterSelector = "" Then
        NextChapterSelector = "" ' �p�G���šA��^�Ŧr��
        Exit Function
    End If
    
    ' �]�w��e����ܾ�
    selector = ChapterSelector ' �d��: "#content > div:nth-child(6) > table > tbody > tr:nth-child(2) > td:nth-child(1) > a"
    
    ' �ϥ� IncrementNthChild �禡�ӧ�s��ܾ�
    newSelector = IncrementNthChild(selector)
    
    ' ��s�R�A�ܼ� ChapterSelector
    ChapterSelector = newSelector
    
    ' ��^�s����ܾ�
    NextChapterSelector = newSelector
End Function
Property Get Head_Edit_textbox() As SeleniumBasic.IWebElement
    If VBA.InStr(WD.url, "&action") Then '"&action=newchapter" �� action=editchapter
        Set Head_Edit_textbox = WD.FindElementByCssSelector("#content > h2")
    End If
End Property
Property Get Title_Edit_textbox() As SeleniumBasic.IWebElement
    If VBA.InStr(WD.url, "&action=") Then '"&action=newchapter" �� action=editchapter
        Set Title_Edit_textbox = WD.FindElementByCssSelector("#title")
    End If
End Property
Property Get Sequence_data_Edit_textbox() As SeleniumBasic.IWebElement
    If VBA.InStr(WD.url, "&action=") Then '"&action=newchapter" �� action=editchapter
        Set Sequence_data_Edit_textbox = WD.FindElementByCssSelector("#sequence")
    End If
End Property
Property Get Textarea_data_Edit_textbox() As SeleniumBasic.IWebElement
    If VBA.InStr(WD.url, "&action=") Then '"&action=newchapter" �� action=editchapter
        Set Textarea_data_Edit_textbox = WD.FindElementByCssSelector("#data")
    End If
End Property
Property Get description_Edit_textbox() As SeleniumBasic.IWebElement
    If VBA.InStr(WD.url, "&action=") Then '"&action=newchapter" �� action=editchapter
        Set description_Edit_textbox = WD.FindElementByCssSelector("#description")
    End If
End Property
Property Get Commit_Edit_textbox() As SeleniumBasic.IWebElement
    If VBA.InStr(WD.url, "&action=") Then '"&action=newchapter" �� action=editchapter
        Set Commit_Edit_textbox = WD.FindElementByCssSelector("#commit")
    End If
End Property
Sub �s����Auto_get_argument()
'    Rem ������
'    Rem �۰ʨ��o�����B������file ID num 3�Ӥ޼�
'    '���Ĥ@�q�K�W�������}�A�p�G https://ctext.org/library.pl?if=en&file=3918&page=1
'    Dim url As String, d As Document, p As Paragraph, iwe As SeleniumBasic.IWebElement
'    Set d = ActiveDocument
'    Set p = d.Paragraphs(1)
'    url = d.Range(p.Range.start, p.Range.End - 1).text
'    d.Range(p.Range.start, p.Range.End - 1).text = 1 '�Ĥ@�q���������X=1
'    p.Range.InsertParagraphAfter
'    Set p = d.Paragraphs(2)
'    d.Range(p.Range.start, p.Range.End - 1).text = pageUBound '��2�q���������X
End Sub
Sub �s����Auto_action_newchapter()
    Dim d As Document, chapterNum As Integer, iwe As SeleniumBasic.IWebElement, newchapterUrl As String, title As String
    Set d = ActiveDocument
    '����4�q��J�n�}�Ҫ��ѭ������A�phttps://ctext.org/library.pl?if=gb&res=4925
    
    If IsWDInvalid Then
        If Not OpenChrome(VBA.Left(d.Paragraphs(4).Range.text, VBA.Len(d.Paragraphs(4).Range.text) - 1)) Then Exit Sub
    Else
        If Not Commit_Edit_textbox Is Nothing Then
            Commit_Edit_textbox.Click '�e�X
        End If
        WD.url = VBA.Left(d.Paragraphs(4).Range.text, VBA.Len(d.Paragraphs(4).Range.text) - 1)
    End If
    WD.SwitchTo.Window WD.CurrentWindowHandle
    '��5�q��J�{�b�n�s�W��쪺�Uchapter�Ǹ��A�p��1�U�h��2�]�U�Ǹ�+1�^
    chapterNum = VBA.CInt(VBA.Left(d.Paragraphs(5).Range.text, VBA.Len(d.Paragraphs(5).Range.text) - 1))
    '"#content > div:nth-child(6) > table > tbody > tr:nth-child(2) > td:nth-child(1) > a"
    ChapterSelector = "#content > div:nth-child(6) > table > tbody > tr:nth-child(" & chapterNum & ") > td:nth-child(1) > a"
    '��6�q���s�W��쪺�������}�G
    'https://ctext.org/wiki.pl?if=gb&res=350225&action=newchapter
    newchapterUrl = VBA.Left(d.Paragraphs(6).Range.text, VBA.Len(d.Paragraphs(6).Range.text) - 1)
    
    '�b�ѭ���T�������I���۹�����chapter�]�U�^
    Set iwe = WD.FindElementByCssSelector(ChapterSelector)
    If iwe Is Nothing Then
        MsgBox "done!", vbInformation
        Exit Sub
    End If
    title = iwe.GetAttribute("text")
    iwe.Click
    Do Until VBA.InStr(WD.url, "&page")
        DoEvents
        Set iwe = WD.FindElementByCssSelector(ChapterSelector)
        If Not iwe Is Nothing Then
            iwe.Click
        End If
    Loop
    'Set iwe = WD.FindElementByCssSelector(Div_generic_IncludePathAndEndPageNum)
    d.Range(d.Paragraphs(1).Range.start, d.Paragraphs(1).Range.End - 1).text = 1 '����
    d.Range(d.Paragraphs(2).Range.start, d.Paragraphs(2).Range.End - 1).text = pageUBound '����
    'file ID
    'https://ctext.org/library.pl?if=gb&file=76754&page=1
    d.Range(d.Paragraphs(3).Range.start, d.Paragraphs(3).Range.End - 1).text = VBA.Trim(VBA.Mid(WD.url, VBA.InStr(WD.url, "&file=") + VBA.Len("&file="), VBA.InStr(WD.url, "&page=") - (VBA.InStr(WD.url, "&file=") + VBA.Len("&file="))))
    d.Activate
    WD.url = newchapterUrl
    WD.SwitchTo.Window WD.CurrentWindowHandle
    ActivateChrome
    Textarea_data_Edit_textbox.Click
    ActivateChrome
    VBA.DoEvents
    �s����
    d.Undo
    If VBA.Len(Textarea_data_Edit_textbox.GetAttribute("value")) < 4 Then
        SetIWebElementValueProperty Textarea_data_Edit_textbox, GetClipboardText
    End If
    '��Jtitle�ȡG
    Dim head As String
    head = Head_Edit_textbox.GetAttribute("outerText")
    head = VBA.Mid(head, 2, VBA.InStr(head, "�n") - 2)
    
    SetIWebElementValueProperty Title_Edit_textbox, VBA.Replace(title, head, vbNullString)
    '��JSequence�ȡG
    SetIWebElementValueProperty Sequence_data_Edit_textbox, VBA.CStr(chapterNum) & "0"
    '��J�ק�K�n:
    SetIWebElementValueProperty description_Edit_textbox, description_Edit_textbox_�s���� '"�ڡm��Ǥj�v�n�ΡmKanripo�n�Ҧ������H���Ǧۻs��GitHub�}���K�O�K�w�ˤ�TextForCtext�n��ƪ��������J�F�Q�װϤΥ���YouTube�W�D����Һt�ܼv���C�P���P���@�g���g�ۡ@�n�L��������"
'    SetIWebElementValueProperty Description_Edit_textbox, "�ڥ_�ʤ��ެ�ަ������q�m���ެ�ޤޱo�Ʀr�H��귽���O�P������N���m�n�Ҧ������H���Ǧۻs��GitHub�}���K�O�K�w�ˤ�TextForCtext�n��ƪ��������J�F�Q�װϤΥ���YouTube�W�D����Һt�ܼv���C�P���P���@�g���g�ۡ@�n�L��������"
    'Commit_Edit_textbox.Click '�e�X
    
    Title_Edit_textbox.Click
    d.Range(d.Paragraphs(5).Range.start, d.Paragraphs(5).Range.End - 1).text = chapterNum + 1
    
    '���U�i�O�s�s��j
    WD.FindElementByCssSelector("#commit").Click
    
    d.Application.Activate
    d.Application.windowState = wdWindowStateNormal
    d.Activate
    
End Sub


Rem �b�Ǹ����ɹs�H�վ㳹�`�Ψ䦸�ǥΡC�]�m�s�w�ؤ������n���վ㳹�`�����צӳ]�]��즳290�ӳ��G�I�I�^ https://ctext.org/wiki.pl?if=en&res=589161 20241214
Sub Add0toSequenceField()
    Dim w, url As String, iwe As SeleniumBasic.IWebElement, add0 As String
    If SeleniumOP.IsWDInvalid() Then
        If WD Is Nothing Then
            SeleniumOP.OpenChrome "https://ctext.org/"
        Else
            'WD.SwitchTo.Window SeleniumOP.WindowHandles()(SeleniumOP.WindowHandlesCount - 1)
'            WD.SwitchTo.Window SeleniumOP.WindowHandles()(0)
            WD.SwitchTo.Window SeleniumOP.WD.WindowHandles()(0)
        End If
    End If
    If Selection.Type = wdSelectionIP Then
        add0 = VBA.InputBox("��J�n��0���ȡA�p�n�ɨ��0�A�h��J�u00�v�C�P���P���@�g���g�ۡ@�n�L��������@�g���D")
    Else
        ResetSelectionAvoidSymbols
        add0 = Selection.text
    End If
    For Each w In WD.WindowHandles
        WD.SwitchTo.Window w
        url = WD.url
        If VBA.InStr(url, "https://ctext.org/wiki.pl") = 1 And VBA.InStr(url, "&chapter=") Then
            'Edit Link
            Set iwe = WD.FindElementByCssSelector("#content > h2 > span > a:nth-child(2)")
            iwe.Click
            'sequence Box
            Set iwe = WD.FindElementByCssSelector("#sequence")
            SeleniumOP.SetIWebElementValueProperty iwe, iwe.GetAttribute("value") & add0
            'Submit changes
            Set iwe = WD.FindElementByCssSelector("#commit")
            iwe.Click
            VBA.Interaction.DoEvents
            Do While VBA.InStr(WD.url, "&action=editchapter")
                SystemSetup.wait 0.3
            Loop
            WD.Close
        End If
    Next w
End Sub

Sub tempReplaceTxtforCtextEdit()
Dim a, d As Document, i As Integer, x As String
a = Array("{{�]", "{{", "�^}}", "}}", "�]", "{{", "�^", "}}", "��", VBA.ChrW(12295))
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
a = Array("{{�]", "{{", "�^}}", "}}", "�]", "{{", "�^", "}}", "��", VBA.ChrW(12295))
Set d = Documents.Add
d.Range.Paste
For i = 0 To UBound(a)
    d.Range.Find.Execute a(i), , , , , , , wdFindContinue, , a(i + 1), wdReplaceAll
    i = i + 1
Next i
d.Range.Cut
d.Application.windowState = wdWindowStateMinimize
d.Close wdDoNotSaveChanges
AppActivateDefaultBrowser
SendKeys "^v"
SendKeys "{tab}~"

End Sub



