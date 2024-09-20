Attribute VB_Name = "Fonts"
Option Explicit

Rem 主管所有與字型有關的操作

Rem 20240920 Copilot大菩薩：https://sl.bing.net/bfIsi0Ehd48 Word VBA 字型檢查
Function IsFontInstalled(fontName As String) As Boolean
    Dim font As Variant
    Dim fontInstalled As Boolean
    fontInstalled = False
    
    ' 遍歷所有已安裝的字型
    For Each font In word.Application.FontNames
        If font = fontName Then
            fontInstalled = True
            Exit For
        End If
    Next font
    
    IsFontInstalled = fontInstalled
End Function


