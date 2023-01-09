Attribute VB_Name = "code"
Option Explicit



Public Function UrlEncode(ByRef szString As String) As String '�H�U��ƥi�H�s�X���媺URL�G VBA�PUnicode Ansi URL�s�X�ѽX���������N�X���A - ���\�ݭn�۫ߪ��峹 - ���G https://zhuanlan.zhihu.com/p/435181691
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




Function URLDecode(ByVal strIn) ' ���BExcel-VBA-UTF-8 �a�}�ѽX �s�X ��� �]�@�̡G���P�G�^ �H�U��ƥi�H�ѽXUTF-8�a�}������������C

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

hh = Mid(strIn, sl + kl, 2) '����

a = Int("&H" & hh) 'ascii?

If Abs(a) < 128 Then

sl = sl + 3

Else

hi = Mid(strIn, sl + 3 + kl, 2) '�C��

a = Int("&H" & hh & hi) '�Dascii?

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
'�O�_����ƥi�N��r�����X(BIG-5)��ܥX��
'(??�Ӥ[?�k�^�`)
'robert788417 years ago
'Permalink�Ҧp =code('A') ���G��(65)
'=???(��) ���G��(A4DF) <------BIG-5�X
'
'�γ\�O�°��D ���ۤߨD�� ����!!
'��17 years ago
'PermalinkVBA:
'Print Hex(Asc("��")) ' Big5
'A4DF
'Print Hex(AscW("��")) ' Unicode
'5 FC3
'
'�ҥH�� VBA �]�@�Ө�ơG
'Function MyAsc(ByVal strChar As String) As String
'MyAsc = Hex(Asc(strChar))
'End Function
'
'�b�u�@���
'=MyAsc(A1)


Sub �M���Ҧ��{���X����()

With ActiveDocument.Range.Find
    .ClearFormatting
    .Font.ColorIndex = 11
    .Execute "", , , , , , , wdFindContinue, , "", wdReplaceAll
    .ClearFormatting
End With
End Sub
