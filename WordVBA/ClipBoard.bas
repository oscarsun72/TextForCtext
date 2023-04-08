Attribute VB_Name = "ClipBoard"
Option Explicit
Rem 20230408 Bing大菩薩：
#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function lstrcpy Lib "kernel32" Alias "lstrcpyW" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
#Else
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare PtrSafe Function lstrcpy Lib "kernel32" Alias "lstrcpyW" (ByVal lpString1 As Any, ByVal lpString2 Any) As Long
#End If

Public Const GHND = &H42
Public Const CF_UNICODETEXT = 13&

Public Sub SetClipboard(sUniText As String)
   #If VBA7 Then
       Dim iStrPtr As LongPtr
       Dim iLock As LongPtr
       Dim ILEN As LongPtr
   #Else
       Dim iStrPtr As Long
       Dim iLock As Long
       Dim ILEN As Long
   #End If
   
   OpenClipboard 0&
   
   ILEN = Len(sUniText)
   
   iStrPtr = GlobalAlloc(GHND, ILEN * 2 + 2)
   
   If iStrPtr Then
   
        iLock = GlobalLock(iStrPtr)
        
        lstrcpy iLock, StrPtr(sUniText)
        
        GlobalUnlock iStrPtr
        
        SetClipboardData CF_UNICODETEXT, iStrPtr
        
   End If
   
   CloseClipboard
   
End Sub

Public Function GetClipboard() As String

   #If VBA7 Then
       Dim iStrPtr As LongPtr
       Dim ILEN As LongPtr
       Dim iLock As LongPtr
   #Else
       Dim iStrPtr As Long
       Dim ILEN As Long
       Dim iLock As Long
   #End If
   
   Dim sUniText As String
   
   OpenClipboard 0&
   
   If IsClipboardFormatAvailable(CF_UNICODETEXT) Then
   
        iStrPtr = GetClipboardData(CF_UNICODETEXT)
        
        If iStrPtr Then
        
            iLock = GlobalLock(iStrPtr)
            
            ILEN = GlobalSize(iStrPtr)
            
            sUniText = String$(CLng(ILEN \ 2&) - 1&, vbNullChar)
            
            lstrcpy StrPtr(sUniText), iLock
            
            GlobalUnlock iStrPtr
            
        End If
        
        GetClipboard = sUniText
        
   End If
   
   CloseClipboard
   
End Function
' 您可以將上面提供的代碼複製並粘貼到您的 VBA 項目中。這些代碼定義了兩個函數：SetClipboard 和 GetClipboard。
' SetClipboard 函數接受一個字符串參數，並將其設置為剪貼板的內容。
' GetClipboard 函數返回剪貼板中的文本內容。
' 您可以在您的 VBA 代碼中調用這兩個函數來操作剪貼板。
' 例如，要將剪貼板設置為 “Hello World”，您可以使用以下代碼：
' SetClipboard "Hello World"
' 要獲取剪貼板中的文本內容，您可以使用以下代碼：
' Dim sText As String
' sText = GetClipboard()
' 希望這些信息對您有所幫助。如果您有其他問題，請隨時告訴我。
' 如果您想要將圖形或其他非文本數據設置為剪貼板的內容，您需要使用不同的剪貼板格式。在上面提供的代碼中，我們使用了 `CF_UNICODETEXT` 常量來指定剪貼板格式為 Unicode 文本。
'對於圖形數據，您可以使用 `CF_BITMAP`、`CF_DIB` 或 `CF_ENHMETAFILE` 等剪貼板格式。您需要根據您要設置的圖形數據類型選擇合適的剪貼板格式。
'此外，您還需要修改 `SetClipboard` 函數以接受圖形數據作為參數，並使用相應的 Windows API 函數來將圖形數據設置為剪貼板的內容。
'由於這涉及到更複雜的 Windows API 編程，因此我無法提供完整的代碼示例。但是，您可以在網上搜索相關資料來瞭解如何使用 VBA 和 Windows API 來操作剪貼板中的圖形數據。
'是的，使用 VBA 操作剪貼板中的文本數據相對容易實現。這是因為 VBA 本身就提供了對字符串的良好支持，而且 Windows API 中也有許多用於操作文本數據的函數。
''對於其他類型的數據，如圖形、文件等，操作起來會相對複雜一些，因為您需要更多地使用 Windows API 函數，並且需要更多地瞭解剪貼板格式和數據類型。
''但是，只要您願意學習並掌握相關知識，使用 VBA 操作剪貼板中的其他類型數據也是完全可行的。
