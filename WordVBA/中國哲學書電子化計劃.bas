Attribute VB_Name = "������Ǯѹq�l�ƭp��"

Sub �s����()
'the page begin
Const start As Integer = 2375
' the page end
Const e As Integer = 2380
' the book
Const fileID As Long = 1000081
'https://ctext.org/library.pl?if=gb&file=1000081&page=2621

Dim x As String, data As New MSForms.DataObject
Dim i As Integer
For i = start To e
    x = x & "<scanbegin file=""" & fileID & """ page=""" & i & """ />" & Chr(9) & "<scanend file=""" & fileID & """ page=""" & i & """ />" '�Y�����S�����󤺮e�A�����̫�K���ন�@�q���C�Y��n�@�Ӭq���A�|�P�U�@���H�X�b�@�_
Next i


'For Each e In Selection.Value
'    x = x & e
'Next e
''x = Replace(x, Chr(13), "")
data.SetText Replace(x, "/>", "/>��", 1, 1)
data.PutInClipboard
End Sub
Sub �M���Ҧ��Ÿ�_���q�`��Ÿ��ҥ~()
Dim f, i As Integer
f = Array("�C", "�v", Chr(-24152), "�G", "�A", "�F", _
    "�B", "�u", ".", Chr(34), ":", ",", ";", _
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


Sub �����w�|���O�Z�����()
Dim d As Document, a, i

a = Array("^p^p", "@", "�q", "{{", "�r", "}}", "^p", "", "}}{{", "^p", "@", "^p", _
    "��", ChrW(12295))
Set d = Documents.Add()
d.Range.Paste
For i = 0 To UBound(a) - 1
    d.Range.Find.Execute a(i), , , , , , True, wdFindContinue, , a(i + 1), wdReplaceAll
    i = i + 1
Next i
��r�B�z.�ѦW���g�W���Ъ`
d.Range.Cut
d.Close wdDoNotSaveChanges
Beep
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
d.Range.PasteAndFormat wdFormatPlainText
d.Range.Text = VBA.Replace(d.Range.Text, " ", "")
For i = 0 To UBound(a) - 1
    If a(i) = "^p^p^p" Then
        px = d.Range.Text
        Do While InStr(px, Chr(13) & Chr(13) & Chr(13))
            px = Replace(px, Chr(13) & Chr(13) & Chr(13), Chr(13) & Chr(13))
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
    If Left(px, 7) = "{{" & ChrW(-9217) & ChrW(-8195) & "{{{" Then '�`�}�q��
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
        If InStr(Right(px, 4), "<p>") Then
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
    If VBA.Left(p.Range.Text, 9) = ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195) & "�i�m�����n" Then
        Set rng = p.Range
        p.Range.Characters(1).Delete
        rng.SetRange p.Range.start, p.Range.start
        rng.InsertAfter "{{"
        rng.SetRange p.Range.Characters(p.Range.Characters.Count - 4).End, p.Range.Characters(p.Range.Characters.Count - 4).End
        rng.InsertAfter "}}"
        If Len(rng.Paragraphs(1).Next.Range.Text) = 1 Then rng.Paragraphs(1).Next.Range.Delete
    End If
    
    If Len(p.Range) < 20 Then
        If (InStr(p.Range, "�m�v�O�n��") Or VBA.Left(p.Range.Text, 3) = "�v�O��") And InStr(p.Range, "*") = 0 Then
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
If VBA.Left(d.Paragraphs(1).Range.Text, 3) = "�v�O��" And InStr(d.Paragraphs(1).Range.Text, "*") = 0 Then
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
    If Left(px, 7) = "{{" & ChrW(-9217) & ChrW(-8195) & "{{{" Then '�`�}�q��
        e = p.Range.Characters(1).End
        rng.SetRange e, e
        rng.MoveEndUntil "�f"
        'rng.Select
        rng.Collapse wdCollapseEnd
        rng.Select
        Selection.MoveRight wdCharacter, 1, wdExtend
        Selection.TypeText "�r}}}" '�N�`�}�s���e�@�f���k��f�令}}}
        px = p.Range.Text
        If InStr(Right(px, 4), "<p>") Then
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
    If VBA.Left(p.Range.Text, 9) = ChrW(-9217) & ChrW(-8195) & ChrW(-9217) & ChrW(-8195) & "�i�m�����n" Then
        Set rng = p.Range
        p.Range.Characters(1).Delete
        rng.SetRange p.Range.start, p.Range.start
        rng.InsertAfter "{{"
        rng.SetRange p.Range.Characters(p.Range.Characters.Count).End, p.Range.Characters(p.Range.Characters.Count).End
        rng.InsertAfter "}}"
    End If
Next p
If VBA.Left(d.Paragraphs(1).Range.Text, 3) = "�v�O��" Then
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
    If Left(p.Range.Text, 7) = "{{" & ChrW(-9217) & ChrW(-8195) & "{{{" Then
        If InStr(Right(px, 4), "<p>") Then
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

Sub �԰굦_�|���O�Z_�����w��()
'https://ctext.org/library.pl?if=gb&res=77385
Dim a, rng As Range, rngDoc As Range, p As Paragraph, i As Long, rngCnt As Integer
Set rngDoc = Documents.Add.Range
rngDoc.Paste
For Each a In rngDoc.Characters
    If Not a.Next Is Nothing And Not a.Previous Is Nothing Then
        If a = "�@" And a.Next <> "�@" And a.Previous <> "�@" Then
            If a.Previous <> Chr(13) Then a.InsertBefore Chr(13)
            Set a = a.Next
        End If
    End If
Next a

For Each p In rngDoc.Paragraphs
    Set rng = p.Range
    If StrComp(rng.Characters(1), "�@") = 0 And InStr(rng, "}") > 0 Then
        For Each a In rng.Characters
           i = i + 1
           If rng.Characters(i) = "}" Then Exit For
           If rng.Characters(i) = Chr(13) Or rng.Characters(i) = "{" Then
                i = 0
                Exit For
           End If
        Next a
        If i <> 0 Then
            rng.SetRange rng.Characters(1).End, rng.Characters(i).start
'            rng.Select
'            Stop
            rngCnt = rng.Characters.Count
            If rngCnt > 1 Then
                If rngCnt Mod 2 = 1 Then
                    If rng.Characters((rngCnt - rngCnt Mod 2) / 2 + 1).Next <> "�@" _
                        Then rng.Characters((rngCnt - rngCnt Mod 2) / 2).InsertAfter "�@"

                Else
                    If rng.Characters((rngCnt - rngCnt Mod 2) / 2).Next <> "�@" _
                        Then rng.Characters((rngCnt - rngCnt Mod 2) / 2).InsertAfter "�@"
                End If
            End If
        End If
        i = 0
    End If
Next
rngDoc.Cut
rngDoc.Document.Close wdDoNotSaveChanges
AppActivate "TextForCtext"
End Sub

Sub tempReplaceTxtforCtext()
Dim a, d As Document, i As Integer
a = Array("�]", "", "�^", "", "��", ChrW(12295))
Set d = Documents.Add
d.Range.Paste
For i = 0 To UBound(a)
    d.Range.Find.Execute a(i), , , , , , , wdFindContinue, , a(i + 1), wdReplaceAll
    i = i + 1
Next i
d.Range.Cut
d.Close wdDoNotSaveChanges
AppActivate "google chrome"
SendKeys "^v"
SendKeys "{tab}~"

End Sub

