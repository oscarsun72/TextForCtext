Attribute VB_Name = "�~�y�q�l���m��Ʈw"
Option Explicit
Sub clearResultMarked(d As Document)
    Dim a As Range, rng As Range
    'Dim posRng As Long '���եΩΥH�ƸU�@
    'Set d = Documents.Add
    Set rng = d.Range
    With rng
        SystemSetup.playSound 1
        .Paste '�m�~�y�����Ʈw�n�睊��ŶKï�į�ܮt�A�n���q�קK �P���P���@�g���g�ۡ@�n�L�������� 20240825
        SystemSetup.playSound 1 '���ĥi�H���նK�W�n����h�[�I20240825
        .Find.Font.Shading.BackgroundPatternColor = 65535
        Do While .Find.Execute()
            If Asc(.text) > 13 Or Asc(.text) < 0 Then
                .Shading.BackgroundPatternColor = wdColorAutomatic
                With .Font
                    .Color = 0
                    .Bold = 0
                End With
            End If
            rng.SetRange rng.End, d.Range.End - 1
            
'            If posRng = rng.End Then
'                SystemSetup.playSound 1.294
'                Exit Do
'            End If
'            posRng = rng.End

        Loop
        .Find.ClearFormatting
    End With
    
    'For Each a In d.Characters
    '    If a.Shading.BackgroundPatternColor = 65535 Then '<> wdColorAutomatic
    '        With a
    '            .Shading.BackgroundPatternColor = wdColorAutomatic
    '            With .Font
    '                .Color = 0
    '                .Bold = 0
    '            End With
    '        End With
    '    End If
    'Next a
    'd.Range.Cut '�m�~�y�����Ʈw�n�睊��ŶKï�į�ܮt�A�n���q�קK �P���P���@�g���g�ۡ@�n�L�������� 20240825
    'DoEvents
End Sub

Sub �~�y�q�l���m��Ʈw�奻��z_�H��K�줤����Ǯѹq�l�ƭp��()
    Dim d As Document
    Set d = Documents.Add
    clearResultMarked d
    If Not ActiveDocument Is d Then
        d.Activate
    End If
    ��r�B�z.�~�y�q�l���m��Ʈw�奻��z_�H��K�줤����Ǯѹq�l�ƭp��
    SystemSetup.playSound 2
End Sub
Sub �~�y�q�l���m��Ʈw�奻��z_�Q�T�g�`��()
    Dim d As Document
    Set d = Documents.Add
    clearResultMarked d
    SystemSetup.playSound 1.469
    If Not ActiveDocument Is d Then
        d.Activate
    End If
    word.Application.ScreenUpdating = False
    ��r�B�z.�~�y�q�l���m��Ʈw�奻��z_�H��K�줤����Ǯѹq�l�ƭp�� True
    SystemSetup.playSound 1.469
    �~�y�q�l���m��Ʈw�奻��z_�Q�T�g�`��_sub
    SystemSetup.playSound 1.469
    word.Application.ScreenUpdating = True
    On Error Resume Next
    If d.ActiveWindow.Visible = True And word.Application.Visible = True Then
        AppActivate "TextForCtext" '"EmEditor"
        SendKeys "%{insert}", True
    End If
End Sub
Sub �~�y�q�l���m��Ʈw�奻��z_�Q�T�g�`��_sub()
    Dim d As Document, a, i As Integer, ub As Integer
    a = Array("^p" & VBA.ChrW(12310) & "��" & VBA.ChrW(12311), VBA.ChrW(12310) & "��" & VBA.ChrW(12311) & "{{", _
        "�D", "", "����", "�m���n��G", "���q��", "�m���q�n��G", "��", VBA.ChrW(12295), "^pν��", "<p>�qν�r��G", "^p�H��", "<p>�q�H�r��G", _
        "^p", "}}<p>^p", "^p" & VBA.ChrW(12295), "}}<p>" & VBA.ChrW(12295), VBA.ChrW(12295) & "^p", VBA.ChrW(12295) & "}}<p>", _
        "}}", "�C}}", "�C}}<p>^p�C}}<p>", "�C}}<p>", "�C}}<p>�C}}<p>", "�C}}<p>", "{{�`�C}}", "���m�`�n�G")
    ub = UBound(a) - 1
    Set d = ActiveDocument
    If d.path <> "" Then
        Set d = Documents.Add
        d.Range.Paste
    End If
    For i = 0 To ub
        d.Range.Find.Execute a(i), , , , , , True, wdFindContinue, , a(i + 1), wdReplaceAll
        i = i + 1
    Next i
    
    ��r�B�z.�ѦW���g�W���Ъ`
    d.Range.Cut
    d.Close wdDoNotSaveChanges
End Sub

Sub get�����հɰO()
    '��ꥻ���A�b��ѦU���U�A�������C
    Dim rng As Range, a, noteFlag As Boolean, x As String
    Set rng = Documents.Add().Range
    rng.Paste
    For Each a In rng.Characters
        If a.Font.Color = 255 Then
            If a.Font.Size = 10 Then
                noteFlag = False
            Else
                noteFlag = True
            End If
            If noteFlag = False Then
                If a.Next.Font.Size = 10 Then
                    x = x & a
                Else
                    x = x & a & "�G{{"
                End If
            Else
                If a.Next.Font.Size > 7.5 Or a.Next.Font.Color <> 255 Then
                    x = x & a & "}}<p>" & VBA.Chr(13) & VBA.Chr(10)
                Else
                    x = x & a
                End If
            End If
        End If
    Next a
    rng.text = x
    ��r�B�z.�ѦW���g�W���Ъ`
    rng.Cut
    rng.Document.Close wdDoNotSaveChanges
    'SystemSetup.ClipboardPutIn x
    Beep
End Sub

Rem 20230610 YouChat�j���ġGhttps://you.com/search?q=%E6%89%80%E4%BB%A5+Python%E8%A3%A1%E9%A0%AD%E7%9A%84+re.Sub+%E5%B0%B1%E9%A1%9E%E4%BC%BC++VBA%E4%B8%AD%E7%9A%84+re.Replace+%E5%9B%89&cid=c1_fb622a50-b65c-41dd-8f1d-0ad276074e80&tbm=youchat
'Function CleanTextPicPageMark1(text As String)
'  Dim re As New RegExp
'  'Dim text As String
'  Dim cleanedText As String
'
'  'Set re = New RegExp
'  Set re = CreateObject("vbscript.regexp")
'  re.Pattern = "\d+-\d+\s*�i�ϡj?\s*"
'  re.Global = True
'  'cleanedText = VBA.Replace(re.Replace(text, vbNullString), vba.Chr(13), vbNullString)
'  cleanedText = re.Replace(text, vbNullString)
'  CleanTextPicPageMark = cleanedText
'  Debug.Print cleanedText
'End Function
Rem 20230610 Bing�j���ġG
Function CleanTextPicPageMark(text As String)
    Dim re As Object 'New RegExp
    Dim cleanedText As String

    Set re = CreateObject("vbscript.regexp")
    re.Pattern = "\s*\d+-\d+\s*�i�ϡj\s*" '�u\s�v�Gspace �M ���q�Ÿ���
    're.Pattern = "\d+-\d+\s*�i�ϡj?\s*"
     '�M���Ѧp�G
        '�u
        '7-2
        '�i�ϡj
        '�v
'        ���奻
    re.Global = True
    cleanedText = re.Replace(text, vbNullString)
    CleanTextPicPageMark = cleanedText
End Function


