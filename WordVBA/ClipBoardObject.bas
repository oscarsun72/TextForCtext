Attribute VB_Name = "ClipBoardObject"
Option Explicit
Rem creedit_with_Copilot�j���� 20240910
Rem https://www.facebook.com/oscarsun72/posts/pfbid02fCK6wJNrTJSo2Br4zKFrTxWoGd3pYdQhC1D3cxrHKFB7sVoV6LSL1XusVs45Q7EQl
Rem https://www.facebook.com/oscarsun72/posts/pfbid02VYcQ4dMtZVZNcuiA3AAykxyj9pnspALVa6f7mf3CcP7Y44LE6NZMiGsj7R9TRJMwl
Rem ���Ѹѻ��G
'SetClipboard�G�]�m�ŶKï���e�C�ϥ� OpenClipboard ���}�ŶKï�AEmptyClipboard �M�ŰŶKï�A�M��ϥ� GlobalAlloc ���t���s�AGlobalLock ��w���s�Alstrcpy �ƻs�r�Ŧ�A�̫�ϥ� SetClipboardData �]�m�ŶKï���e�A�èϥ� CloseClipboard �����ŶKï�C
'GetClipboard�GŪ���ŶKï���e�C�ϥ� OpenClipboard ���}�ŶKï�A�ˬd�O�_���i�Ϊ��ŶKï�榡�A�M��ϥ� GetClipboardData ����ŶKï�ƾڡA�ϥ� GlobalLock ��w���s�Alstrcpy �ƻs�r�Ŧ�A�̫�ϥ� CloseClipboard �����ŶKï�C
'ClearClipboard�G�M�ŰŶKï�C�ϥ� OpenClipboard ���}�ŶKï�AEmptyClipboard �M�ŰŶKï�A�èϥ� CloseClipboard �����ŶKï�C
#If VBA7 Then
    #If Win64 Then
        ' 64�줸����
        Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
        Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
        Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
        Public Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hMem As LongPtr) As LongPtr
        Public Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal uFormat As Long) As LongPtr
        Public Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
        Public Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
        Public Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
        Public Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
        Public Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
        Public Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)
    #Else
        ' 32�줸����
        Public Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
        Public Declare Function EmptyClipboard Lib "user32" () As Long
        Public Declare Function CloseClipboard Lib "user32" () As Long
        Public Declare Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hMem As Long) As Long
        Public Declare Function GetClipboardData Lib "user32" (ByVal uFormat As Long) As Long
        Public Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
        Public Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
        Public Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
        Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
        Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
        Public Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
    #End If
#Else
    ' Office 2010�ΥH�U����
#End If

' �]�m�ŶKï���e
Public Sub SetClipboard(sUniText As String)
    Dim iStrPtr As LongPtr
    Dim iLen As LongPtr
    Dim iLock As LongPtr
    Const GMEM_MOVEABLE As Long = &H2
    Const GMEM_ZEROINIT As Long = &H40
    Const CF_UNICODETEXT As Long = &HD
    OpenClipboard 0&
    EmptyClipboard
    iLen = LenB(sUniText) + 2&
    iStrPtr = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, iLen)
    iLock = GlobalLock(iStrPtr)
    RtlMoveMemory iLock, StrPtr(sUniText), iLen
    GlobalUnlock iStrPtr
    SetClipboardData CF_UNICODETEXT, iStrPtr
    CloseClipboard
End Sub

' Ū���ŶKï���e
Public Function GetClipboard() As String
    Dim iStrPtr As LongPtr
    Dim iLen As LongPtr
    Dim iLock As LongPtr
    Dim sUniText As String
    Const CF_UNICODETEXT As Long = 13&
    OpenClipboard 0&
    If IsClipboardFormatAvailable(CF_UNICODETEXT) Then
        iStrPtr = GetClipboardData(CF_UNICODETEXT)
        If iStrPtr Then
            iLock = GlobalLock(iStrPtr)
            iLen = GlobalSize(iStrPtr)
            sUniText = VBA.String$(iLen \ 2& - 1&, vbNullChar)
            RtlMoveMemory ByVal StrPtr(sUniText), ByVal iLock, iLen - 2
            GlobalUnlock iStrPtr
        End If
        GetClipboard = sUniText
    End If
    CloseClipboard
End Function



' �M�ŰŶKï
Public Sub ClearClipboard()
    OpenClipboard 0&
    EmptyClipboard
    CloseClipboard
End Sub
