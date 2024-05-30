Attribute VB_Name = "ClipBoardOp"
Option Explicit
'
'Rem 20230407 Bing�j���ġG�z�n�A�p�G�z�Q�b 64 �쪩���� Office ���B�榹�N�X�A�h�ݭn�N hClipMemory �M lpClipMemory �ܶq��������אּ LongPtr �Ӥ��O Long�C���~�A�z�ٻݭn�T�O�Ҧ��Ω�P�ŶK�O�M�������s�椬����Ƴ����T�n���èϥΤF PtrSafe ����r�C
'Rem �H�U�O�ק�᪺�N�X:
'#If VBA7 Then
''    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
''    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
''    Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
''    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
''    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
'    Private Declare PtrSafe Function EmptyClipboard Lib "user32.dll" () As Long
'    Private Declare PtrSafe Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
'    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As LongPtr
'#Else
''    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
''    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
''    Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
''    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
''    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
'    Private Declare PtrSafe Function EmptyClipboard Lib "user32.dll" () As Long
'    Private Declare PtrSafe Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
'    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
'#End If
'
#If VBA7 Then
'    Private Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
'    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
'    Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As Long
'    Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
'    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
'    Private Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
#Else
'    Private Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hWnd As Long) As Long
'    Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "User32" (ByVal wFormat As Long) As Long
    Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
'    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
'    Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
'    Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
'    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
'    Private Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
#End If
'Public Const GHND = &H42
'Public Const CF_TEXT = 1
'Public Const MAXSIZE = 4096
'
'Rem 20230407 Bing�j���ġG
'Rem �z�n�A�p�G�z�b�ϥ� 64 �쪩���� Office�A�h�ݭn�N iLock �M iStrPtr �ܶq��������אּ LongPtr �Ӥ��O Long�C�o�˥i�H�T�O�N�X�b 64 �쪩���� Office �����T�B��C
'Rem �й��ձN�N�X��אּ�H�U���e:



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
    On Error GoTo eH
    TextRange.Paste
    If (TextRange.Tables.Count > 0) Then
        ������Ǯѹq�l�ƭp��.�M���奻�������s���x�s�� TextRange
        TextRange.Copy
    End If
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
exitFunction:
    d.Close wdDoNotSaveChanges
    word.Application.ScreenUpdating = True
Exit Function
eH:
    Select Case Err.Number
        Case 4605
            MsgBox Err.Description '����k���ݩʵL�k�ϥΡA�]��[�ŶKï] �O�Ū��εL�Ī��C
            GoTo exitFunction
        Case Else
            MsgBox Err.Number + Err.Description
    End Select
End Function




