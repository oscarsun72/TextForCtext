Attribute VB_Name = "�~�y�q�l���m��Ʈw"
Option Explicit
Sub clearResultMarked()
Dim a As Range, d As Document, rng As Range
Set d = Documents.Add
Set rng = d.Range
With rng
    .Paste
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
d.Range.Cut
DoEvents
End Sub

Sub �~�y�q�l���m��Ʈw�奻��z_�H��K�줤����Ǯѹq�l�ƭp��()
clearResultMarked
��r�B�z.�~�y�q�l���m��Ʈw�奻��z_�H��K�줤����Ǯѹq�l�ƭp��
SystemSetup.playSound 2
End Sub
Sub �~�y�q�l���m��Ʈw�奻��z_�Q�T�g�`��()
clearResultMarked
��r�B�z.�~�y�q�l���m��Ʈw�奻��z_�H��K�줤����Ǯѹq�l�ƭp�� True
�~�y�q�l���m��Ʈw�奻��z_�Q�T�g�`��_sub
On Error Resume Next
AppActivate "TextForCtext" '"EmEditor"
End Sub
Sub �~�y�q�l���m��Ʈw�奻��z_�Q�T�g�`��_sub()
Dim d As Document, a, i As Integer, ub As Integer
a = Array("^p" & ChrW(12310) & "��" & ChrW(12311), ChrW(12310) & "��" & ChrW(12311) & "{{", _
    "�D", "", "����", "�m���n��G", "���q��", "�m���q�n��G", "��", ChrW(12295), "^pν��", "<p>�qν�r��G", "^p�H��", "<p>�q�H�r��G", _
    "^p", "}}<p>^p", "^p" & ChrW(12295), "}}<p>" & ChrW(12295), ChrW(12295) & "^p", ChrW(12295) & "}}<p>", _
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
                x = x & a & "}}<p>" & Chr(13) & Chr(10)
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
'  'cleanedText = VBA.Replace(re.Replace(text, vbNullString), Chr(13), vbNullString)
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


