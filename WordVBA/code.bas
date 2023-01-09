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

With ActiveDocument.Range.Find
    .ClearFormatting
    .Font.ColorIndex = 11
    .Execute "", , , , , , , wdFindContinue, , "", wdReplaceAll
    .ClearFormatting
End With
End Sub
