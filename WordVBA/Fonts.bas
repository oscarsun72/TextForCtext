Attribute VB_Name = "Fonts"
Option Explicit

Rem �D�ީҦ��P�r���������ާ@

Rem 20240920 Copilot�j���ġGhttps://sl.bing.net/bfIsi0Ehd48 Word VBA �r���ˬd
Function IsFontInstalled(fontName As String) As Boolean
    Dim font As Variant
    Dim fontInstalled As Boolean
    fontInstalled = False
    
    ' �M���Ҧ��w�w�˪��r��
    For Each font In word.Application.FontNames
        If font = fontName Then
            fontInstalled = True
            Exit For
        End If
    Next font
    
    IsFontInstalled = fontInstalled
End Function


