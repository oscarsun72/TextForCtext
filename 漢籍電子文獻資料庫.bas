Attribute VB_Name = "�~�y�q�l���m��Ʈw"
Option Explicit
Sub �~�y�q�l���m��Ʈw�奻��z_�H��K�줤����Ǯѹq�l�ƭp��()
��r�B�z.�~�y�q�l���m��Ʈw�奻��z_�H��K�줤����Ǯѹq�l�ƭp��
End Sub
Sub �~�y�q�l���m��Ʈw�奻��z_�Q�T�g�`��()
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
