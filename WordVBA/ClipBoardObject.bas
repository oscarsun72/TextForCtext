Attribute VB_Name = "ClipBoardObject"
Option Explicit
Rem creedit_with_Copilot大菩薩 20240910
Rem https://www.facebook.com/oscarsun72/posts/pfbid02fCK6wJNrTJSo2Br4zKFrTxWoGd3pYdQhC1D3cxrHKFB7sVoV6LSL1XusVs45Q7EQl
Rem https://www.facebook.com/oscarsun72/posts/pfbid02VYcQ4dMtZVZNcuiA3AAykxyj9pnspALVa6f7mf3CcP7Y44LE6NZMiGsj7R9TRJMwl
Rem 註解解說：
'SetClipboard：設置剪貼簿內容。使用 OpenClipboard 打開剪貼簿，EmptyClipboard 清空剪貼簿，然後使用 GlobalAlloc 分配內存，GlobalLock 鎖定內存，lstrcpy 複製字符串，最後使用 SetClipboardData 設置剪貼簿內容，並使用 CloseClipboard 關閉剪貼簿。
'GetClipboard：讀取剪貼簿內容。使用 OpenClipboard 打開剪貼簿，檢查是否有可用的剪貼簿格式，然後使用 GetClipboardData 獲取剪貼簿數據，使用 GlobalLock 鎖定內存，lstrcpy 複製字符串，最後使用 CloseClipboard 關閉剪貼簿。
'ClearClipboard：清空剪貼簿。使用 OpenClipboard 打開剪貼簿，EmptyClipboard 清空剪貼簿，並使用 CloseClipboard 關閉剪貼簿。
#If VBA7 Then
    #If Win64 Then
        ' 64位元環境
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
        ' 32位元環境
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
    ' Office 2010及以下版本
#End If

' 設置剪貼簿內容
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

' 讀取剪貼簿內容
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



' 清空剪貼簿
Public Sub ClearClipboard()
    OpenClipboard 0&
    EmptyClipboard
    CloseClipboard
End Sub
