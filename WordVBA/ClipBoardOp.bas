Attribute VB_Name = "ClipBoardOp"
Option Explicit
'
'Rem 20230407 Bing大菩薩：您好，如果您想在 64 位版本的 Office 中運行此代碼，則需要將 hClipMemory 和 lpClipMemory 變量的類型更改為 LongPtr 而不是 Long。此外，您還需要確保所有用於與剪貼板和全局內存交互的函數都正確聲明並使用了 PtrSafe 關鍵字。
'Rem 以下是修改後的代碼:
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
'Rem 20230407 Bing大菩薩：
'Rem 您好，如果您在使用 64 位版本的 Office，則需要將 iLock 和 iStrPtr 變量的類型更改為 LongPtr 而不是 Long。這樣可以確保代碼在 64 位版本的 Office 中正確運行。
'Rem 請嘗試將代碼更改為以下內容:



Const CF_HTML = &HC3&

'複製網頁文本到剪貼簿，一樣沒用；只能取得純文字內容，得不到HTML等格式化標記
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
    '隱藏開啟文件便不會有視窗存在 chatGPT菩薩又錯了
'    d.Windows(1).Visible = False
    ' 將剪貼簿的內容加入Word檔案
    Set TextRange = d.Range
    On Error GoTo eH
    TextRange.Paste
    If (TextRange.Tables.Count > 0) Then
        中國哲學書電子化計劃.清除文本頁中的編號儲存格 TextRange
        TextRange.Copy
    End If
    ' 檢查字型顏色是否為綠色
    For Each a In TextRange.Characters
        If a.Font.Color = 34816 Then
            'MsgBox "剪貼簿中有綠色文字"
            Is_ClipboardContainCtext_Note_InlinecommentColor = True
            d.Close wdDoNotSaveChanges
            word.Application.ScreenUpdating = True
            Exit Function
'        Else
'            MsgBox "剪貼簿中沒有綠色文字"
        End If
    Next a
exitFunction:
    d.Close wdDoNotSaveChanges
    word.Application.ScreenUpdating = True
Exit Function
eH:
    Select Case Err.Number
        Case 4605
            MsgBox Err.Description '此方法或屬性無法使用，因為[剪貼簿] 是空的或無效的。
            GoTo exitFunction
        Case Else
            MsgBox Err.Number + Err.Description
    End Select
End Function




