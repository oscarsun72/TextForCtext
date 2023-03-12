Attribute VB_Name = "code"
Option Explicit

Enum CJKBlockName 'https://en.wikipedia.org/wiki/CJK_characters
    CJK_Unified_Ideographs 'Enum�]�C�|�^�S�����w�`�ȡA�h�q���0�}�l
    CJK_Unified_Ideographs_Extension_A
    CJK_Unified_Ideographs_Extension_B
    CJK_Unified_Ideographs_Extension_C
    CJK_Unified_Ideographs_Extension_D
    CJK_Unified_Ideographs_Extension_E
    CJK_Unified_Ideographs_Extension_F
    CJK_Unified_Ideographs_Extension_G
    CJK_Unified_Ideographs_Extension_H
    CJK_Radicals_Supplement
    Kangxi_Radicals
    Ideographic_Description_Characters
    CJK_Symbols_and_Punctuation
    CJK_Strokes
    Enclosed_CJK_Letters_and_Months
    CJK_Compatibility
    CJK_Compatibility_Ideographs
    CJK_Compatibility_Forms
    Enclosed_Ideographic_Supplement
    CJK_Compatibility_Ideographs_Supplement
End Enum

Enum CJKChartRange 'https://en.wikipedia.org/wiki/CJK_characters
CJK_Unified_Ideographs_start = &H4E00
CJK_Unified_Ideographs_Extension_A_start = &H3400
CJK_Unified_Ideographs_Extension_B_start = &H20000
CJK_Unified_Ideographs_Extension_C_start = &H2A700
CJK_Unified_Ideographs_Extension_D_start = &H2B740
CJK_Unified_Ideographs_Extension_E_start = &H2B820
CJK_Unified_Ideographs_Extension_F_start = &H2CEB0
CJK_Unified_Ideographs_Extension_G_start = &H30000
CJK_Unified_Ideographs_Extension_H_start = &H31350
CJK_Radicals_Supplement_start = &H2E80
Kangxi_Radicals_start = &H2F00
Ideographic_Description_Characters_start = &H2FF0
CJK_Symbols_and_Punctuation_start = &H3000
CJK_Strokes_start = &H31C0
Enclosed_CJK_Letters_and_Months_start = &H3200
CJK_Compatibility_start = &H3300
CJK_Compatibility_Ideographs_start = &HF900
CJK_Compatibility_Forms_start = &HFE30
Enclosed_Ideographic_Supplement_start = &H1F200
CJK_Compatibility_Ideographs_Supplement_start = &H2F800
CJK_Unified_Ideographs_end = &H9FFF
CJK_Unified_Ideographs_Extension_A_end = &H4DBF
CJK_Unified_Ideographs_Extension_B_end = &H2A6DF
CJK_Unified_Ideographs_Extension_C_end = &H2B73F
CJK_Unified_Ideographs_Extension_D_end = &H2B81F
CJK_Unified_Ideographs_Extension_E_end = &H2CEAF
CJK_Unified_Ideographs_Extension_F_end = &H2EBEF
CJK_Unified_Ideographs_Extension_G_end = &H3134F
CJK_Unified_Ideographs_Extension_H_end = &H323AF
CJK_Radicals_Supplement_end = &H2EFF
Kangxi_Radicals_end = &H2FDF
Ideographic_Description_Characters_end = &H2FFF
CJK_Symbols_and_Punctuation_end = &H303F
CJK_Strokes_end = &H31EF
Enclosed_CJK_Letters_and_Months_end = &H32FF
CJK_Compatibility_end = &H33FF
CJK_Compatibility_Ideographs_end = &HFAFF
CJK_Compatibility_Forms_end = &HFE4F
Enclosed_Ideographic_Supplement_end = &H1F2FF
CJK_Compatibility_Ideographs_Supplement_end = &H2FA1F

End Enum
Enum CJKChartRangeString 'https://en.wikipedia.org/wiki/CJK_characters
CJK_Unified_Ideographs_start = "&H4E00"
CJK_Unified_Ideographs_Extension_A_start = "&H3400"
CJK_Unified_Ideographs_Extension_B_start = "&H20000"
CJK_Unified_Ideographs_Extension_C_start = "&H2A700"
CJK_Unified_Ideographs_Extension_D_start = "&H2B740"
CJK_Unified_Ideographs_Extension_E_start = "&H2B820"
CJK_Unified_Ideographs_Extension_F_start = "&H2CEB0"
CJK_Unified_Ideographs_Extension_G_start = "&H30000"
CJK_Unified_Ideographs_Extension_H_start = "&H31350"
CJK_Radicals_Supplement_start = "&H2E80"
Kangxi_Radicals_start = "&H2F00"
Ideographic_Description_Characters_start = "&H2FF0"
CJK_Symbols_and_Punctuation_start = "&H3000"
CJK_Strokes_start = "&H31C0"
Enclosed_CJK_Letters_and_Months_start = "&H3200"
CJK_Compatibility_start = "&H3300"
CJK_Compatibility_Ideographs_start = "&HF900"
CJK_Compatibility_Forms_start = "&HFE30"
Enclosed_Ideographic_Supplement_start = "&H1F200"
CJK_Compatibility_Ideographs_Supplement_start = "&H2F800"
CJK_Unified_Ideographs_end = "&H9FFF"
CJK_Unified_Ideographs_Extension_A_end = "&H4DBF"
CJK_Unified_Ideographs_Extension_B_end = "&H2A6DF"
CJK_Unified_Ideographs_Extension_C_end = "&H2B73F"
CJK_Unified_Ideographs_Extension_D_end = "&H2B81F"
CJK_Unified_Ideographs_Extension_E_end = "&H2CEAF"
CJK_Unified_Ideographs_Extension_F_end = "&H2EBEF"
CJK_Unified_Ideographs_Extension_G_end = "&H3134F"
CJK_Unified_Ideographs_Extension_H_end = "&H323AF"
CJK_Radicals_Supplement_end = "&H2EFF"
Kangxi_Radicals_end = "&H2FDF"
Ideographic_Description_Characters_end = "&H2FFF"
CJK_Symbols_and_Punctuation_end = "&H303F"
CJK_Strokes_end = "&H31EF"
Enclosed_CJK_Letters_and_Months_end = "&H32FF"
CJK_Compatibility_end = "&H33FF"
CJK_Compatibility_Ideographs_end = "&HFAFF"
CJK_Compatibility_Forms_end = "&HFE4F"
Enclosed_Ideographic_Supplement_end = "&H1F2FF"
CJK_Compatibility_Ideographs_Supplement_end = "&H2FA1F"

End Enum

Enum SurrogateCodePoint 'https://zhuanlan.zhihu.com/p/147339588
    HighStart = &HD800 'UTF-16 �i�H�x�s U+0000 �� U+10FFFF �������r�X�AU+FFFF �H�U���r�X�H 2 �� byte �x�s�A�� U+10000 �H�W���r�X�A�|�Q���Ӥ��� D800 �� DFFF ��������ơA�Ĥ@�ӳQ�٬� �e�ɥN�z (lead surrogates)�A���� D800 �� DBFF �����A�ĤG�ӳQ�٬� ����N�z (trail surrogates)�A���� DC00 �� DFFF �����AUTF-16 �N�O�Q�γo��ӥN�z��Ӫ�� FFFF ���~�A��L���U��������r�C
    HighEnd = &HDBFF 'https://ithelp.ithome.com.tw/articles/10198444#_=_
    LowStart = &HDC00 'https://zh.wikipedia.org/zh-tw/UTF-16
    LowEnd = &HDFFF
End Enum

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
Dim ur As UndoRecord
SystemSetup.stopUndo ur, "�M���Ҧ��{���X����"
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

Rem 20230215 chatGPT�j���ġG
Rem �o�q�N�X���� IsChineseCharacter ��ƥΩ�P�_��Ӧr�ŬO�_�OCJK��CJK�X�i�r�Ŷ������~�r�A�� IsChineseString ��ƫh�Ω�P�_�@�Ӧr�Ŧ�O�_������CJK��CJK�X�i�r�Ŷ������~�r�զ��C
Rem �bVBA���A�ڭ̨ϥΤF AscW ��ƨ�����r�Ū�Unicode�s�X�ȡC�M��A�ڭ̴N�i�H�ϥΩMC#���������覡�ӧP�_�r�ŬO�_�ݩ�CJK��CJK�X�i�r�Ŷ������~�r�C
' �P�_�@�Ӧr�ŬO�_�OCJK��CJK�X�i�r�Ŷ������~�r
Public Function IsChineseCharacter(c As String) As Boolean
'    chatGPT�j���ġG Unicode�d��: CJK�r�Ŷ��d��G4E00�V9FFF�ACJK�X�i�r�Ŷ��d��G20000�V2A6DF �]�u�u���G�o�ˮڥ������A�u�� CJK�Τ@��N�Ÿ��MCJK�X�iB
'    Dim unicodeVal As Long
'    unicodeVal = AscW(c)
'    IsChineseCharacter = (unicodeVal >= &H4E00 And unicodeVal <= &H9FFF) Or (unicodeVal >= &H20000 And unicodeVal <= &H2A6DF)
    IsChineseCharacter = IsCJK(c)(1)
End Function

' �P�_�@�Ӧr�Ŧ�O�_������CJK��CJK�X�i�r�Ŷ������~�r�զ�
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

Rem 20230221 chatGPT�j����: VBA�ˬdsurrogate�r�šG
Rem �b VBA ���A�z�i�H�ϥ� AscW ��ƱN�@�Ӧr���ഫ�� Unicode �s�X�C�M��A�z�i�H�ˬd�ӽs�X�O�_�b�N�z��d�򤺡C
Rem Unicode �����N�z��d�� U+D800 �� U+DFFF�A�@�� 2048 �ӥN�X�I�C�N�z��O�@�دS���s�X�Φ��A�Ѩ�� Unicode �s�X�զ��A�Ω��ܸ��j���r�Ŷ��A�p Emoji�C
Rem �U���O�@�ӥܨҨ�ơA�Ө�Ʊ����@�Ӧr�Ũê�^�@�ӥ����ȡA���ܸӦr�ŬO�_���N�z�襤���r�šG
Function IsSurrogate(ByVal ch As String) As Boolean
    Dim code As Long
    code = AscW(ch)
    'IsSurrogate = (code >= &HD800 And code <= &HDFFF) 'https://ithelp.ithome.com.tw/articles/10198444#_=_
    IsSurrogate = (code >= &HD800 And code <= &HDBFF) _
                Or (code >= &HDC00 And code <= &HDFFF) 'https://ithelp.ithome.com.tw/articles/10198444#_=_
    'UTF-16 �i�H�x�s U+0000 �� U+10FFFF �������r�X�AU+FFFF �H�U���r�X�H 2 �� byte �x�s�A�� U+10000 �H�W���r�X�A�|�Q���Ӥ��� D800 �� DFFF ��������ơA
    '�Ĥ@�ӳQ�٬� �e�ɥN�z (lead surrogates)�A���� D800 �� DBFF �����A�ĤG�ӳQ�٬� ����N�z (trail surrogates)�A���� DC00 �� DFFF �����AUTF-16 �N�O�Q�γo��ӥN�z��Ӫ�� FFFF ���~�A��L���U��������r�C
    Rem �o�Ө�ƱN�r���ഫ�� Unicode �s�X�A���ˬd�ӽs�X�O�_�b�N�z��d�򤺡C�p�G�O�A�h��ƪ�^ True�A�_�h��^ False�C�Ъ`�N�AAscW ��ƥu��Ω� Unicode �r�Ŧ�A�p�G�z�n�B�z ANSI �r�Ŧ�A�h�ݭn�ϥ� Asc ��ơC
End Function
Function IsHighSurrogate(ByVal ch As String) As Boolean
    Dim code As Long
    code = AscW(ch)
    IsHighSurrogate = (code >= &HD800 And code <= &HDBFF) 'https://ithelp.ithome.com.tw/articles/10198444#_=_
    'UTF-16 �i�H�x�s U+0000 �� U+10FFFF �������r�X�AU+FFFF �H�U���r�X�H 2 �� byte �x�s�A�� U+10000 �H�W���r�X�A�|�Q���Ӥ��� D800 �� DFFF ��������ơA
    '�Ĥ@�ӳQ�٬� �e�ɥN�z (lead surrogates)�A���� D800 �� DBFF ����
    
End Function
Function IsLowSurrogate(ByVal ch As String) As Boolean
    Dim code As Long
    code = AscW(ch)
    IsLowSurrogate = (code >= &HDC00 And code <= &HDFFF) 'https://ithelp.ithome.com.tw/articles/10198444#_=_
    'UTF-16 �i�H�x�s U+0000 �� U+10FFFF �������r�X�AU+FFFF �H�U���r�X�H 2 �� byte �x�s�A�� U+10000 �H�W���r�X�A�|�Q���Ӥ��� D800 �� DFFF ��������ơA
    '�ĤG�ӳQ�٬� ����N�z (trail surrogates)�A���� DC00 �� DFFF �����AUTF-16 �N�O�Q�γo��ӥN�z��Ӫ�� FFFF ���~�A��L���U��������r�C
    
End Function

Rem 20230224 chatGPT�j���ġG���surrogate pair�r�šA���Өϥ�Unicode�зǤ��ҭz����k�N���ഫ����ӥN�z�r�šC����ӻ��A�N�N�z��]surrogate pair�^����Ӥ������O�٬�high surrogate�Mlow surrogate�C
Rem �H�U�O�N�N�z���ഫ���N�z�r�Ū���k:
Private Function CombineSurrogatePair(ByVal highSurrogate As String, ByVal lowSurrogate As String) As String
    CombineSurrogatePair = ChrW((AscW(highSurrogate) - &HD800&) * &H400& + (AscW(lowSurrogate) - &HDC00&) + &H10000)
End Function
Rem �ϥγo�Ө�ơA�z�i�H�q�L�b�`�����B�z��Ӧr�šA�èϥΤW�����d��ӧP�_�r�ŬO�_�bCJK���r���d�򤺡C �p�G���N�z�r�šA�h�i�H�ϥθӨ�ƱN���ഫ��Unicode�r�šC

Function IsCJK(c As String) As Collection 'Boolean,CJKBlockName
    Dim code As Long, cjk As Boolean, cjkBlackName As CJKBlockName, result As New Collection
'    Dim code
    Rem chatGPT�j���ġG�O���A�z���o�S���C�b VBA ���A�ϥ� AscW �禡���o Unicode �r������ƭȮɡA�p�G�ǤJ���r��O surrogate pair�A����禡�u�|�p�� pair ���Ĥ@�Ӧr���]�Y High surrogate�^���ȡC�]���A�i�H�����ϥ� AscW(c) �ӭp�� c ����ƭȡA�Ӥ����A�ϥ� Left �禡�Ө��o�Ĥ@�Ӧr���C
    'code = AscW(Left(c, 1))
    'code = AscW(c)
    If Len(c) = 1 Then
        code = AscW(c) 'AscW_IncludeSurrogatePairUnicodecode(c)
    Else
        getCodePoint c, code
    End If
    Rem https://en.wikipedia.org/wiki/CJK_characters
    'CJK Unified Ideographs
    'If code >= CLng("&H4E00") And code <= CLng("&H9FFF") Then'�@�w�n�uCLng("&H9FFF")�v ���� �uCLng(&H9FFF)�v
    If code >= CLng(CJKChartRangeString.CJK_Unified_Ideographs_start) And code <= CLng(CJKChartRangeString.CJK_Unified_Ideographs_end) Then
        cjk = True: cjkBlackName = CJKBlockName.CJK_Unified_Ideographs
    'CJK Compatibility Ideographs
    'ElseIf code >= CLng("&HF900") And code <= CLng("&HFAFF") Then
    ElseIf code >= CLng(CJKChartRangeString.CJK_Compatibility_Ideographs_start) And code <= CLng(CJKChartRangeString.CJK_Compatibility_Ideographs_end) Then
        cjk = True: cjkBlackName = CJKBlockName.CJK_Compatibility_Ideographs
    'CJK Unified Ideographs Extension A
    'ElseIf code >= CLng(&H3400") And code <= CLng("&H4DBF") Then
    ElseIf code >= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_A_start) And code <= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_A_end) Then
        cjk = True: cjkBlackName = CJKBlockName.CJK_Unified_Ideographs_Extension_A
    'CJK Unified Ideographs Extension B
    'ElseIf code >= CLng("&H20000") And code <= CLng("&H2A6DF") Then
    ElseIf code >= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_B_start) And code <= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_B_end) Then
        cjk = True: cjkBlackName = CJKBlockName.CJK_Unified_Ideographs_Extension_B
    'CJK Unified Ideographs Extension C
'    ElseIf code >= CLng("&H2A700") And code <= CLng("&H2B73F") Then
    ElseIf code >= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_C_start) And code <= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_C_end) Then
        cjk = True: cjkBlackName = CJKBlockName.CJK_Unified_Ideographs_Extension_C
    'CJK Unified Ideographs Extension D
'    ElseIf code >= CLng("&H2B740") And code <= CLng("&H2B81F") Then
    ElseIf code >= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_D_start) And code <= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_D_end) Then
        cjk = True: cjkBlackName = CJKBlockName.CJK_Unified_Ideographs_Extension_D
    'CJK Unified Ideographs Extension E
    'ElseIf code >= CLng("&H2B820") And code <= CLng("&H2CEAF") Then
    ElseIf code >= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_E_start) And code <= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_E_end) Then
        cjk = True: cjkBlackName = CJKBlockName.CJK_Unified_Ideographs_Extension_E
    'CJK Unified Ideographs Extension F
    'ElseIf code >= CLng("&H2CEB0") And code <= CLng("&H2EBEF") Then
    ElseIf code >= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_F_start) And code <= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_F_end) Then
        cjk = True: cjkBlackName = CJKBlockName.CJK_Unified_Ideographs_Extension_F
    'CJK Unified Ideographs Extension G
    ElseIf code >= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_G_start) And code <= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_G_end) Then
        cjk = True: cjkBlackName = CJKBlockName.CJK_Unified_Ideographs_Extension_G
    'CJK Unified Ideographs Extension H
    ElseIf code >= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_H_start) And code <= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_H_end) Then
        cjk = True: cjkBlackName = CJKBlockName.CJK_Unified_Ideographs_Extension_H
    'CJK Radicals Supplement
    ElseIf code >= CLng(CJKChartRangeString.CJK_Radicals_Supplement_start) And code <= CLng(CJKChartRangeString.CJK_Radicals_Supplement_end) Then
        cjk = True: cjkBlackName = CJKBlockName.CJK_Radicals_Supplement
    'Kangxi Radicals
    ElseIf code >= CLng(CJKChartRangeString.Kangxi_Radicals_start) And code <= CLng(CJKChartRangeString.Kangxi_Radicals_end) Then
        cjk = True: cjkBlackName = CJKBlockName.Kangxi_Radicals
    'Ideographic Description Characters
    ElseIf code >= CLng(CJKChartRangeString.Ideographic_Description_Characters_start) And code <= CLng(CJKChartRangeString.Ideographic_Description_Characters_end) Then
        cjk = True: cjkBlackName = CJKBlockName.Ideographic_Description_Characters
    'CJK Symbols And punctuation
    ElseIf code >= CLng(CJKChartRangeString.CJK_Symbols_and_Punctuation_start) And code <= CLng(CJKChartRangeString.CJK_Symbols_and_Punctuation_end) Then
        cjk = True: cjkBlackName = CJKBlockName.CJK_Symbols_and_Punctuation
    'CJK Strokes
    ElseIf code >= CLng(CJKChartRangeString.CJK_Strokes_start) And code <= CLng(CJKChartRangeString.CJK_Strokes_end) Then
        cjk = True: cjkBlackName = CJKBlockName.CJK_Strokes
    'Enclosed CJK Letters and Months
    ElseIf code >= CLng(CJKChartRangeString.Enclosed_CJK_Letters_and_Months_start) And code <= CLng(CJKChartRangeString.Enclosed_CJK_Letters_and_Months_end) Then
        cjk = True: cjkBlackName = CJKBlockName.Enclosed_CJK_Letters_and_Months
    'CJK Compatibility
    ElseIf code >= CLng(CJKChartRangeString.CJK_Compatibility_start) And code <= CLng(CJKChartRangeString.CJK_Compatibility_end) Then
        cjk = True: cjkBlackName = CJKBlockName.CJK_Compatibility
    'CJK Compatibility Forms
    ElseIf code >= CLng(CJKChartRangeString.CJK_Compatibility_Forms_start) And code <= CLng(CJKChartRangeString.CJK_Compatibility_Forms_end) Then
        cjk = True: cjkBlackName = CJKBlockName.CJK_Compatibility_Forms
    'Enclosed Ideographic Supplement
    ElseIf code >= CLng(CJKChartRangeString.Enclosed_Ideographic_Supplement_start) And code <= CLng(CJKChartRangeString.Enclosed_Ideographic_Supplement_end) Then
        cjk = True: cjkBlackName = CJKBlockName.Enclosed_Ideographic_Supplement
    'CJK Compatibility Ideographs Supplement
    'ElseIf code >= CLng("&H2F800") And code <= CLng("&H2FA1F") Then
    ElseIf code >= CLng(CJKChartRangeString.CJK_Compatibility_Ideographs_Supplement_start) And code <= CLng(CJKChartRangeString.CJK_Compatibility_Ideographs_Supplement_end) Then
        cjk = True: cjkBlackName = CJKBlockName.CJK_Compatibility_Ideographs_Supplement
    Else
        cjk = False
    End If
    result.Add cjk
    result.Add cjkBlackName
    Set IsCJK = result
Rem chatGPT�j���ġG��p�A�ڤ��e�^�������~�C�z���쪺�u���v�r��Unicode�X�T��O5143�A�ݩ�CJK�򥻶��d�򤺡C
Rem �t�~�A�ڤ��e���p��O���~���A�]���N16�i���ର10�i��ɻݭn�`�N���t���C���T���d�������G
Rem CJK�򥻶��G4E00�]19968�^��9FFF�]40959�^
Rem CJK�X�iA�G3400�]13312�^��4DBF�]19871�^
Rem CJK�X�iB�G20000�]131072�^��2A6DF�]173791�^
Rem CJK�X�iC�G2A700�]173824�^��2B73F�]177983�^
Rem CJK�X�iD�G2B740�]177984�^��2B81F�]178207�^
Rem CJK�X�iE�G2B820�]178208�^��2CEAF�]235519�^
Rem CJK�X�iF�G2CEB0�]235520�^��2EBEF�]303231�^
Rem ���� &H9FFF �ন�t�ƪ����D�A�O�]���bVBA���A����������̰��쬰�Ÿ���A�p�G�̰��쬰1�A�h��ܭt�ơC�]���A&H9FFF �N�Q��@�t�ƳB�z�A���ڭȬ� -24577�C
End Function
Function HextoLng(hexValue As String) As Long
    'HextoLng = CLng(hexValue) And &HFFFF 'Val("&H" & Right("0000" & hexValue, 4))
    HextoLng = CLng(hexValue)
End Function

Function AscW_IncludeSurrogatePairUnicodecode(ByVal str As String) As Long
    Dim utf16 As String
    utf16 = StrConv(str, vbUnicode)
    Dim code As Long
    If Len(utf16) = 2 Then ' surrogate pair
        code = (CLng(AscW(Mid(utf16, 1, 1))) - &HD800&) * &H400& + (CLng(AscW(Mid(utf16, 2, 1))) - &HDC00&) + &H10000
    Else
        code = AscW(utf16)
    End If
    AscW_IncludeSurrogatePairUnicodecode = code
End Function
Sub getCodePoint(character As String, codepoint As Long)
' ����r�Ŧꪺ high surrogate �M low surrogate �� AscW() ��
codepoint = ((CLng(AscW(Left(character, 1))) - &HD800) * &H400) + (CLng(AscW(Right(character, 1))) - &HDC00) + &H10000
Rem �S���uCLng�v�૬�|����A�Y�̦p isCJK_Ext()�禡�����覡�A�H���O�� Long ���ܼ��x�s��ȡA��|���t�૬
End Sub


Function isCJK_Ext(str As String, whatBlockNameInExt As CJKBlockName) As Boolean
Dim codepoint As Long
Dim highSurrogate As Long
Dim lowSurrogate As Long

' ����r�Ŧꪺ high surrogate �M low surrogate �� AscW() ��
highSurrogate = AscW(Left(str, 1))
lowSurrogate = AscW(Right(str, 1))

If (highSurrogate >= SurrogateCodePoint.HighStart And highSurrogate <= SurrogateCodePoint.HighEnd) _
    And (lowSurrogate >= SurrogateCodePoint.LowStart And lowSurrogate <= SurrogateCodePoint.LowEnd) Then
    ' �p��r�Ū��X�I��!!!!!!!!!!!!!!!!!
'    codepoint = ((highSurrogate - &HD800) * &H400) + (lowSurrogate - &HDC00) + &H10000
    getCodePoint str, codepoint '�Y�S�H�uCLng()�v�૬�|����A�H���O�� Long ���ܼ��x�s��ȡA�Y�|���t�૬
        
        Rem forDebugText
'    If codepoint = &H2E4E5 Then Stop
'    If Hex(codepoint) = "2E4E5" Then Stop

    Select Case whatBlockNameInExt
        Case CJKBlockName.CJK_Unified_Ideographs_Extension_A
            If codepoint >= CJKChartRange.CJK_Unified_Ideographs_Extension_A_start And codepoint <= CJKChartRange.CJK_Unified_Ideographs_Extension_A_end Then isCJK_Ext = True
        Case CJKBlockName.CJK_Unified_Ideographs_Extension_B
            If codepoint >= CJKChartRange.CJK_Unified_Ideographs_Extension_B_start And codepoint <= CJKChartRange.CJK_Unified_Ideographs_Extension_B_end Then isCJK_Ext = True
        Case CJKBlockName.CJK_Unified_Ideographs_Extension_C
            If codepoint >= CJKChartRange.CJK_Unified_Ideographs_Extension_C_start And codepoint <= CJKChartRange.CJK_Unified_Ideographs_Extension_C_end Then isCJK_Ext = True
        Case CJKBlockName.CJK_Unified_Ideographs_Extension_D
            If codepoint >= CJKChartRange.CJK_Unified_Ideographs_Extension_D_start And codepoint <= CJKChartRange.CJK_Unified_Ideographs_Extension_D_end Then isCJK_Ext = True
        Case CJKBlockName.CJK_Unified_Ideographs_Extension_E
            If codepoint >= CJKChartRange.CJK_Unified_Ideographs_Extension_E_start And codepoint <= CJKChartRange.CJK_Unified_Ideographs_Extension_E_end Then isCJK_Ext = True
        Case CJKBlockName.CJK_Unified_Ideographs_Extension_F
            If codepoint >= CJKChartRange.CJK_Unified_Ideographs_Extension_F_start And codepoint <= CJKChartRange.CJK_Unified_Ideographs_Extension_F_end Then isCJK_Ext = True
        Case CJKBlockName.CJK_Unified_Ideographs_Extension_G
            If codepoint >= CJKChartRange.CJK_Unified_Ideographs_Extension_G_start And codepoint <= CJKChartRange.CJK_Unified_Ideographs_Extension_G_end Then isCJK_Ext = True
        Case CJKBlockName.CJK_Unified_Ideographs_Extension_H
            If codepoint >= CJKChartRange.CJK_Unified_Ideographs_Extension_H_start And codepoint <= CJKChartRange.CJK_Unified_Ideographs_Extension_H_end Then isCJK_Ext = True
    End Select
End If
End Function

Rem 20230225 chatGPT�j���ġGCJK-ext F high surrogate.�G�P�_ Unicode �r�ŬO�_�b CJK-Ext F �d�򤺡A�åB�p��X�r�Ū��X�I�ȡG
Function isCJK_ExtF(str As String) As Boolean
'https://ithelp.ithome.com.tw/articles/10198444#_=_
'�Ĥ@�ӳQ�٬� �e�ɥN�z (lead surrogates)�A���� D800 �� DBFF ����
'�ĤG�ӳQ�٬� ����N�z (trail surrogates)�A���� DC00 �� DFFF ����

Dim codepoint As Long
Dim highSurrogate As Long
Dim lowSurrogate As Long

' ����r�Ŧꪺ high surrogate �M low surrogate �� AscW() ��
highSurrogate = AscW(Left(str, 1))
lowSurrogate = AscW(Right(str, 1))

If (highSurrogate >= &HD84D And highSurrogate <= &HDBFF) And (lowSurrogate >= &HDC00 And lowSurrogate <= &HDFFF) Then
    ' �p��r�Ū��X�I��
    codepoint = ((highSurrogate - &HD800) * &H400) + (lowSurrogate - &HDC00) + &H10000
    
    If codepoint >= &H2CEB0 And codepoint <= &H2EBEF Then
        ' �r�Ŧb CJK-Ext F �d��
        isCJK_ExtF = True
    Else
        ' �r�Ť��b CJK-Ext F �d��
    End If
Else
    ' �r�Ť��b CJK-Ext F �d��
End If
'�N�X�޿�p�U:
'
'������r�Ŧꪺ high surrogate �M low surrogate �� AscW() �ȡC
'�p�G high surrogate �M low surrogate �� AscW() �ȳ��b CJK-Ext F �d�򤺡A�h�p��r�Ū��X�I�ȡC
'�P�_�r�Ū��X�I�ȬO�_�b CJK-Ext F �d�򤺡A�p�G�b�A�h�����r�Ŧb CJK-Ext F �d�򤺡F�p�G���b�A�h�����r�Ť��b CJK-Ext F �d�򤺡C
'�p��r�Ū��X�I�Ȫ������p�U:
'
'codePoint = ((highSurrogate - &HD800) * &H400) + (lowSurrogate - &HDC00) + &H10000
'
'�䤤�A&HD800 �M &HDC00 ���O�O high surrogate �M low surrogate ����ǭȡA&H400 �O surrogate pair �������q�A&H10000 �O Unicode �s�X����ǭȡC

End Function


Rem chatGPT�j����:WordVBA�ʦr���:�b Word ���A���U Alt + X ��i�H�N�ҿ��r�ഫ��������� Unicode �X�I�A�o�ӥ\��٬� Unicode �r�ſ�J�C
Rem �b VBA ���A�i�H�ϥ� Selection.Range.Text �� Range.Text �ݩʨ�����ҿ��r�νd�򪺤��e�A�M��ϥ� Selection.Range.Text = ChrW(unicode_code) �� Range.Text = ChrW(unicode_code) �ӱN���ഫ�� Unicode �X�I�ҹ������r�šC
Rem �H�U�O�@�ӥܨҡA�i�ܤF�p��ϥ� VBA �b Word ���N��w�d�򪺤��e�ഫ���� Unicode �X�I�G
Sub ConvertToUnicode_SelectionToggleCharacterCode() '������@ Selection.ToggleCharacterCode ��k
    Dim selectedText As String
    Dim unicodeCode As Long
    
    selectedText = selection.Range.text
    
    If Len(selectedText) = 1 Then
        unicodeCode = AscW(selectedText)
        selection.Range.text = Hex(unicodeCode)
    ElseIf Len(selectedText) = 2 Then
        unicodeCode = (AscW(Mid(selectedText, 1, 1)) - &HD800&) * &H400& + (AscW(Mid(selectedText, 2, 1)) - &HDC00&) + &H10000 '
        getCodePoint selectedText, unicodeCode
        selection.Range.text = Hex(unicodeCode)
    Else
        MsgBox "Invalid selection"
        Exit Sub
    End If
    
'    Selection.Range.text = ChrW(unicodeCode)
    Rem chatGPT���ġG�`�N�A�b�B�z surrogate pair �ɡA�ݭn�N��ӥN�z�諸 Unicode �X�I�ഫ����ڪ� Unicode �X�I�C�W�z�ܨҤ����N�X�N�O�N surrogate pair �ഫ����ڪ� Unicode �X�I���d�ҡC
End Sub
Rem creedit with chatGPT�j���ġG
Function ConvertToUnicode(chartoConvert As String) As Long
    Dim unicodeCode As Long
    If Len(chartoConvert) = 1 Then
        unicodeCode = AscW(chartoConvert)
    ElseIf Len(chartoConvert) = 2 Then
        'unicodeCode = (CLng(AscW(Mid(chartoConvert, 1, 1))) - &HD800&) * &H400& + (CLng(AscW(Mid(chartoConvert, 2, 1))) - &HDC00&) + &H10000
        'unicodeCode = ((CLng(AscW(Mid(chartoConvert, 1, 1))) - &HD800)) * &H400 + (CLng(AscW(Mid(chartoConvert, 2, 1))) - &HDC00) + &H10000
        getCodePoint chartoConvert, unicodeCode
    Else
        MsgBox "Invalid character"
        Exit Function
    End If
    ConvertToUnicode = unicodeCode
    
End Function


