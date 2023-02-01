Attribute VB_Name = "ClipBoardOp"
Option Explicit
Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal format As Integer) As Long
Declare Function GetClipboardData Lib "user32" (ByVal uFormat As Integer) As Long
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
    TextRange.Paste

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
    d.Close wdDoNotSaveChanges
    word.Application.ScreenUpdating = True
End Function
