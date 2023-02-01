Attribute VB_Name = "ClipBoardOp"
Option Explicit
Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal format As Integer) As Long
Declare Function GetClipboardData Lib "user32" (ByVal uFormat As Integer) As Long
Const CF_HTML = &HC3&

'�ƻs�����奻��ŶKï�A�@�˨S�ΡF�u����o�¤�r���e�A�o����HTML���榡�ƼаO
Function ClipboardGetHTML() As String
If IsClipboardFormatAvailable(CF_HTML) = 0 Then
MsgBox "The Clipboard does not contain HTML data."
Exit Function
End If
Dim hClipboardData As Long
hClipboardData = GetClipboardData(CF_HTML)
Dim strHTML As String
strHTML = StrConv(hClipboardData, vbUnicode)
ClipboardGetHTML = strHTML
End Function

Function Is_ClipboardContainCtext_Note_InlinecommentColor() As Boolean
    Dim TextRange As word.Range
    Dim d As Document, a As Range
    DoEvents
    word.Application.ScreenUpdating = False
    DoEvents
    Set d = Documents.Add(, , , False)
    '���ö}�Ҥ��K���|�������s�b chatGPT���ĤS���F
'    d.Windows(1).Visible = False
    ' �N�ŶKï�����e�[�JWord�ɮ�
    Set TextRange = d.Range
    TextRange.Paste

    ' �ˬd�r���C��O�_�����
    For Each a In TextRange.Characters
        If a.Font.Color = 34816 Then
            'MsgBox "�ŶKï��������r"
            Is_ClipboardContainCtext_Note_InlinecommentColor = True
            d.Close wdDoNotSaveChanges
            word.Application.ScreenUpdating = True
            Exit Function
'        Else
'            MsgBox "�ŶKï���S������r"
        End If
    Next a
    d.Close wdDoNotSaveChanges
    word.Application.ScreenUpdating = True
End Function
