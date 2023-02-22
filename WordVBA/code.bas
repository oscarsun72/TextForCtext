Attribute VB_Name = "code"
Option Explicit



Public Function UrlEncode(ByRef szString As String) As String '以下函數可以編碼中文的URL： VBA與Unicode Ansi URL編碼解碼等相關的代碼集錦 - 成功需要自律的文章 - 知乎 https://zhuanlan.zhihu.com/p/435181691
Dim szChar As String
Dim szTemp As String
Dim szCode As String
Dim szHex As String
Dim szBin As String
Dim iCount1 As Integer
Dim iCount2 As Integer
Dim iStrLen1 As Integer
Dim iStrLen2 As Integer
Dim lResult As Long
Dim lAscVal As Long
szString = Trim$(szString)
iStrLen1 = Len(szString)
For iCount1 = 1 To iStrLen1
szChar = Mid$(szString, iCount1, 1)
lAscVal = AscW(szChar)
If lAscVal >= &H0 And lAscVal <= &HFF Then
If (lAscVal >= &H30 And lAscVal <= &H39) Or _
(lAscVal >= &H41 And lAscVal <= &H5A) Or _
(lAscVal >= &H61 And lAscVal <= &H7A) Then
szCode = szCode & szChar
Else
szCode = szCode & "%" & Hex(AscW(szChar))
End If
Else
szHex = Hex(AscW(szChar))
iStrLen2 = Len(szHex)
For iCount2 = 1 To iStrLen2
szChar = Mid$(szHex, iCount2, 1)
Select Case szChar
Case Is = "0"
szBin = szBin & "0000"
Case Is = "1"
szBin = szBin & "0001"
Case Is = "2"
szBin = szBin & "0010"
Case Is = "3"
szBin = szBin & "0011"
Case Is = "4"
szBin = szBin & "0100"
Case Is = "5"
szBin = szBin & "0101"
Case Is = "6"
szBin = szBin & "0110"
Case Is = "7"
szBin = szBin & "0111"
Case Is = "8"
szBin = szBin & "1000"
Case Is = "9"
szBin = szBin & "1001"
Case Is = "A"
szBin = szBin & "1010"
Case Is = "B"
szBin = szBin & "1011"
Case Is = "C"
szBin = szBin & "1100"
Case Is = "D"
szBin = szBin & "1101"
Case Is = "E"
szBin = szBin & "1110"
Case Is = "F"
szBin = szBin & "1111"
Case Else
End Select
Next iCount2
szTemp = "1110" & Left$(szBin, 4) & "10" & Mid$(szBin, 5, 6) & "10" & Right$(szBin, 6)
For iCount2 = 1 To 24
If Mid$(szTemp, iCount2, 1) = "1" Then
lResult = lResult + 1 * 2 ^ (24 - iCount2)
Else: lResult = lResult + 0 * 2 ^ (24 - iCount2)
End If
Next iCount2
szTemp = Hex(lResult)
szCode = szCode & "%" & Left$(szTemp, 2) & "%" & Mid$(szTemp, 3, 2) & "%" & Right$(szTemp, 2)
End If
szBin = vbNullString
lResult = 0
Next iCount1
UrlEncode = szCode

End Function




Function URLDecode(ByVal strIn) ' 五、Excel-VBA-UTF-8 地址解碼 編碼 函數 （作者：時鵬亮） 以下函數可以解碼UTF-8地址的中文關鍵詞。

URLDecode = ""

Dim sl: sl = 1

Dim tl: tl = 1

Dim key: key = "%"

Dim kl: kl = Len(key)

sl = InStr(sl, strIn, key, 1)

Do While sl > 0

If (tl = 1 And sl <> 1) Or tl < sl Then

URLDecode = URLDecode & Mid(strIn, tl, sl - tl)

End If

Dim hh, hi, hl

Dim a

Select Case UCase(Mid(strIn, sl + kl, 1))

Case "U" 'Unicode URLEncode

a = Mid(strIn, sl + kl + 1, 4)

URLDecode = URLDecode & ChrW("&H" & a)

sl = sl + 6

Case "E" 'UTF-8 URLEncode

hh = Mid(strIn, sl + kl, 2)

a = Int("&H" & hh) 'ascii?

If Abs(a) < 128 Then

sl = sl + 3

URLDecode = URLDecode & Chr(a)

Else

hi = Mid(strIn, sl + 3 + kl, 2)

hl = Mid(strIn, sl + 6 + kl, 2)

a = ("&H" & hh And &HF) * 2 ^ 12 Or ("&H" & hi And &H3F) * 2 ^ 6 Or ("&H" & hl And &H3F)

If a < 0 Then a = a + 65536

URLDecode = URLDecode & ChrW(a)

sl = sl + 9

End If

Case Else 'Asc URLEncode

hh = Mid(strIn, sl + kl, 2) '高位

a = Int("&H" & hh) 'ascii?

If Abs(a) < 128 Then

sl = sl + 3

Else

hi = Mid(strIn, sl + 3 + kl, 2) '低位

a = Int("&H" & hh & hi) '非ascii?

sl = sl + 6

End If

URLDecode = URLDecode & Chr(a)

End Select

tl = sl

sl = InStr(sl, strIn, key, 1)

Loop

URLDecode = URLDecode & Mid(strIn, tl)

End Function



'https://narkive.com/t730ls1c
'https://microsoft.public.tw.excel.narkive.com/t730ls1c/big-5
'是否有函數可將文字之內碼(BIG-5)顯示出來
'(??太久?法回复)
'robert788417 years ago
'Permalink例如 =code('A') 結果為(65)
'=???(心) 結果為(A4DF) <------BIG-5碼
'
'或許是笨問題 但誠心求教 謝謝!!
'璉璉17 years ago
'PermalinkVBA:
'Print Hex(Asc("心")) ' Big5
'A4DF
'Print Hex(AscW("心")) ' Unicode
'5 FC3
'
'所以用 VBA 包一個函數：
'Function MyAsc(ByVal strChar As String) As String
'MyAsc = Hex(Asc(strChar))
'End Function
'
'在工作表用
'=MyAsc(A1)


Sub 清除所有程式碼註解()
Dim ur As UndoRecord
SystemSetup.stopUndo ur, "清除所有程式碼註解"
With ActiveDocument.Range.Find
    .ClearFormatting
    .Font.ColorIndex = 11
    .Execute "", , , , , , , wdFindContinue, , "", wdReplaceAll
    .ClearFormatting
End With
If InStr(ActiveDocument.Range, "//") Or InStr(ActiveDocument.Range, Chr(39)) > 0 Then
    Dim p As Paragraph
    For Each p In ActiveDocument.Paragraphs
        If InStr(p.Range, "//") > 0 Or InStr(1, p.Range, Chr(39), vbTextCompare) > 0 Then p.Range.Delete
    Next p
End If
ActiveDocument.Range = VBA.Replace(ActiveDocument.Range, Chr(13) & Chr(13), Chr(13))
SystemSetup.contiUndo ur
End Sub

Rem 20230215 chatGPT大菩薩：
Rem 這段代碼中的 IsChineseCharacter 函數用於判斷單個字符是否是CJK或CJK擴展字符集中的漢字，而 IsChineseString 函數則用於判斷一個字符串是否全部由CJK或CJK擴展字符集中的漢字組成。
Rem 在VBA中，我們使用了 AscW 函數來獲取字符的Unicode編碼值。然後，我們就可以使用和C#中類似的方式來判斷字符是否屬於CJK或CJK擴展字符集中的漢字。
' 判斷一個字符是否是CJK或CJK擴展字符集中的漢字
Public Function IsChineseCharacter(c As String) As Boolean
    ' Unicode範圍: CJK字符集範圍：4E00–9FFF，CJK擴展字符集範圍：20000–2A6DF
    Dim unicodeVal As Long
    unicodeVal = AscW(c)
    IsChineseCharacter = (unicodeVal >= &H4E00 And unicodeVal <= &H9FFF) Or (unicodeVal >= &H20000 And unicodeVal <= &H2A6DF)
End Function

' 判斷一個字符串是否全部由CJK或CJK擴展字符集中的漢字組成
Public Function IsChineseString(s As String) As Boolean
    Dim i As Long
    For i = 1 To Len(s)
        If Not IsChineseCharacter(Mid(s, i, 1)) Then
            IsChineseString = False
            Exit Function
        End If
    Next i
    IsChineseString = True
End Function

Rem 20230221 chatGPT大菩薩: VBA檢查surrogate字符：
Rem 在 VBA 中，您可以使用 AscW 函數將一個字符轉換為 Unicode 編碼。然後，您可以檢查該編碼是否在代理對範圍內。
Rem Unicode 中的代理對範圍為 U+D800 到 U+DFFF，共有 2048 個代碼點。代理對是一種特殊的編碼形式，由兩個 Unicode 編碼組成，用於表示較大的字符集，如 Emoji。
Rem 下面是一個示例函數，該函數接受一個字符並返回一個布爾值，指示該字符是否為代理對中的字符：
Function IsSurrogate(ByVal ch As String) As Boolean
    Dim code As Long
    code = AscW(ch)
    IsSurrogate = (code >= &HD800 And code <= &HDFFF)
    Rem 這個函數將字符轉換為 Unicode 編碼，並檢查該編碼是否在代理對範圍內。如果是，則函數返回 True，否則返回 False。請注意，AscW 函數只能用於 Unicode 字符串，如果您要處理 ANSI 字符串，則需要使用 Asc 函數。
End Function

