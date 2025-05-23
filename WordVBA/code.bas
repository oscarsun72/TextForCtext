Attribute VB_Name = "code"
Option Explicit

Enum CJKBlockName 'https://en.wikipedia.org/wiki/CJK_characters
    CJK_Unified_Ideographs 'Enum（列舉）沒有指定常值，則從整數0開始
    CJK_Unified_Ideographs_Extension_A
    CJK_Unified_Ideographs_Extension_B
    CJK_Unified_Ideographs_Extension_C
    CJK_Unified_Ideographs_Extension_D
    CJK_Unified_Ideographs_Extension_E
    CJK_Unified_Ideographs_Extension_F
    CJK_Unified_Ideographs_Extension_G
    CJK_Unified_Ideographs_Extension_H
    CJK_Unified_Ideographs_Extension_I
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
    CJK_Unified_Ideographs_Extension_I_start = &H2EBF0
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
    CJK_Unified_Ideographs_Extension_I_end = &H2EE5D
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
    CJK_Unified_Ideographs_Extension_I_start = "&H2EBF0"
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
    CJK_Unified_Ideographs_Extension_I_end = "&H2EE5D"
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
    HighStart = &HD800 'UTF-16 可以儲存 U+0000 至 U+10FFFF 之間的字碼，U+FFFF 以下的字碼以 2 個 byte 儲存，而 U+10000 以上的字碼，會被拆成兩個介於 D800 至 DFFF 之間的整數，第一個被稱為 前導代理 (lead surrogates)，介於 D800 至 DBFF 之間，第二個被稱為 後尾代理 (trail surrogates)，介於 DC00 至 DFFF 之間，UTF-16 就是利用這兩個代理對來表示 FFFF 之外，其他輔助平面的文字。
    HighEnd = &HDBFF 'https://ithelp.ithome.com.tw/articles/10198444#_=_
    LowStart = &HDC00 'https://zh.wikipedia.org/zh-tw/UTF-16
    LowEnd = &HDFFF
End Enum
Rem 20240106 StackOverflow ai & Bing大菩薩: 建置C# 程式庫成dll檔案
Rem 須【設定引用項目】UrlEncodingDLL.tlb
Public Function UrlEncode_建置CSharp程式庫成dll檔案者(ByRef szString As String) As String
''Public Function UrlEncode(ByRef szString As String) As String
'    If InStr(szString, "%") Then UrlEncode_建置CSharp程式庫成dll檔案者 = szString: Exit Function
'    'If InStr(szString, "%") Then UrlEncode = szString: Exit Function
'    Dim encoder As New UrlEncodingDLL.UrlEncoder
'    Dim encodedUrl As String
'    encodedUrl = encoder.UrlEncode(szString) ' Chinese text
'    'Debug.Print encodedUrl ' Output: "%E4%BD%A0%E5%A5%BD%E4%B8%96%E7%95%8C"
'    UrlEncode_建置CSharp程式庫成dll檔案者 = encodedUrl
''    UrlEncode = encodedUrl
End Function
Rem 20241003 Copilot大菩薩：https://sl.bing.net/jCIXUMLIFTE
Rem VBA URL Encoding Issue ： https://sl.bing.net/ctZR9HBRFBc ： 看來 CreateObject("HTMLFile").parentWindow.encodeURIComponent(szString) 這行在 VBA 中可能會出現物件不支援此屬性或方法的錯誤。這是因為 VBA 中的 CreateObject("HTMLFile") 無法正確地創建 HTML 文件對象。
Public Function UrlEncode(ByRef szString As String) As String
    'Dim scriptControl As Object
    Dim scriptControl As New MSScriptControl.scriptControl '項目->添加引用->COM->類型庫中勾選Microsoft Script Control 1.0組件: https://blog.csdn.net/DXCyber409/article/details/79976104
                     'msscript.ocx       '阿彌陀佛！剛才在網路上Google 「如何讓 Windows 10 登錄 MSScriptControl.ScriptControl 類別」結果看到這一篇 https://blog.csdn.net/DXCyber409/article/details/79976104 裡面提到了 「Microsoft Script Control 1.0」 結果我到VBE的「工具→設定引用項目」中找，就找到了，勾選後再改寫如下，一樣可以。為什麼　Copilot大菩薩您沒提到這個物件引用呢？如果這是有效的（因為我現在離開那個有問題的電腦沒法實測，再下禮拜去才能實測）請您記下來、學習起來，幫助更多和我一樣茫然助的人。感恩感恩　南無阿彌陀佛
    On Error GoTo eH:
    'Set scriptControl = CreateObject("MSScriptControl.ScriptControl")
    scriptControl.Language = "JScript"
    UrlEncode = scriptControl.Eval("encodeURIComponent('" & szString & "')") '在 VBA 中使用 JavaScript 的 encodeURIComponent 方法來進行 URL 編碼了。
    Exit Function
eH:
    Select Case Err.Number
        Case -2147221164
            If VBA.InStr(Err.Description, "類別未登錄") = 1 Then
'                Stop
                playSound 1.294
                UrlEncode = szString
            Else
                GoTo default
            End If
        Case Else
default:
            Debug.Print Err.Number & Err.Description
    End Select
End Function



Public Function UrlEncode_Big5UnicodOLNLY(ByRef szString As String) As String '以下函數可以編碼中文的URL： VBA與Unicode Ansi URL編碼解碼等相關的代碼集錦 - 成功需要自律的文章 - 知乎 https://zhuanlan.zhihu.com/p/435181691
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
    UrlEncode_Big5UnicodOLNLY = szCode

End Function

Public Function UrlEncodeWorksheetFunction(ByRef szString As String) As String '20240103Bing大菩薩：VBA代碼中的問題：
'您好，這是Bing。您的VBA代碼中的問題在於，當您嘗試對超出基本多語言平面（BMP）的Unicode字符（如「??」）進行編碼時，AscW函數會返回一個代理對，而不是實際的Unicode碼點。這就是為什麼您看到的結果是「%E5%B4%87%ED%A1%80%ED%B0%89」而不是「%E5%B4%87%F0%A0%80%89」。
'要解決這個問題，您需要對代理對進行特殊處理。以下是一個修改後的UrlEncode函數，它可以正確處理這種情況：
'這個修改後的函數會檢查每個字符是否為代理對的高代理項，如果是，則會將其與下一個低代理項組合以形成完整的Unicode碼點。然後，它將這個碼點轉換為十六進制並添加到輸出字符串中。希望這對您有所幫助！
    Dim szChar As String
    Dim szCode As String
    Dim iCount1 As Integer
    Dim iStrLen1 As Integer
    Dim lAscVal As Long
    Dim lAscVal2 As Long
    szString = Trim$(szString)
    iStrLen1 = Len(szString)
    For iCount1 = 1 To iStrLen1
        szChar = Mid$(szString, iCount1, 1)
        lAscVal = AscW(szChar)
        If lAscVal >= &HD800 And lAscVal <= &HDBFF And iCount1 < iStrLen1 Then
            lAscVal2 = AscW(Mid$(szString, iCount1 + 1, 1))
            If lAscVal2 >= &HDC00 And lAscVal2 <= &HDFFF Then
                lAscVal = (lAscVal - &HD800) * &H400 + (lAscVal2 - &HDC00) + &H10000
                iCount1 = iCount1 + 1
            End If
        End If
        If lAscVal > &H7F Then
        '須先安裝Excel 並引用參考才行！
'            szCode = szCode & "%" & WorksheetFunction.Dec2Hex(lAscVal, 2)
        Else
            szCode = szCode & szChar
        End If
    Next iCount1
    UrlEncodeWorksheetFunction = szCode
End Function



Function IsAlphaNumeric(ByVal asciiCode As Integer) As Boolean
    IsAlphaNumeric = (asciiCode >= 48 And asciiCode <= 57) Or _
                     (asciiCode >= 65 And asciiCode <= 90) Or _
                     (asciiCode >= 97 And asciiCode <= 122)
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

URLDecode = URLDecode & VBA.Mid(strIn, tl, sl - tl)

End If

Dim hh, hi, hl

Dim a

Select Case UCase(VBA.Mid(strIn, sl + kl, 1))

Case "U" 'Unicode URLEncode

a = VBA.Mid(strIn, sl + kl + 1, 4)

URLDecode = URLDecode & VBA.ChrW("&H" & a)

sl = sl + 6

Case "E" 'UTF-8 URLEncode

hh = VBA.Mid(strIn, sl + kl, 2)

a = Int("&H" & hh) 'ascii?

If Abs(a) < 128 Then

sl = sl + 3

URLDecode = URLDecode & VBA.Chr(a)

Else

hi = VBA.Mid(strIn, sl + 3 + kl, 2)

hl = VBA.Mid(strIn, sl + 6 + kl, 2)

a = ("&H" & hh And &HF) * 2 ^ 12 Or ("&H" & hi And &H3F) * 2 ^ 6 Or ("&H" & hl And &H3F)

If a < 0 Then a = a + 65536

URLDecode = URLDecode & VBA.ChrW(a)

sl = sl + 9

End If

Case Else 'Asc URLEncode

hh = VBA.Mid(strIn, sl + kl, 2) '高位

a = Int("&H" & hh) 'ascii?

If Abs(a) < 128 Then

sl = sl + 3

Else

hi = VBA.Mid(strIn, sl + 3 + kl, 2) '低位

a = Int("&H" & hh & hi) '非ascii?

sl = sl + 6

End If

URLDecode = URLDecode & VBA.Chr(a)

End Select

tl = sl

sl = InStr(sl, strIn, key, 1)

Loop

URLDecode = URLDecode & VBA.Mid(strIn, tl)

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
    .font.ColorIndex = 11
    .Execute "", , , , , , , wdFindContinue, , "", wdReplaceAll
    .ClearFormatting
End With
If InStr(ActiveDocument.Range, "//") Or InStr(ActiveDocument.Range, VBA.Chr(39)) > 0 Then
    Dim p As Paragraph
    For Each p In ActiveDocument.Paragraphs
        If InStr(p.Range, "//") > 0 Or InStr(1, p.Range, VBA.Chr(39), vbTextCompare) > 0 Then p.Range.Delete
    Next p
End If
ActiveDocument.Range = VBA.Replace(ActiveDocument.Range, VBA.Chr(13) & VBA.Chr(13), VBA.Chr(13))
SystemSetup.contiUndo ur
End Sub

Rem 20230215 chatGPT大菩薩：
Rem 這段代碼中的 IsChineseCharacter 函數用於判斷單個字符是否是CJK或CJK擴展字符集中的漢字，而 IsChineseString 函數則用於判斷一個字符串是否全部由CJK或CJK擴展字符集中的漢字組成。
Rem 在VBA中，我們使用了 AscW 函數來獲取字符的Unicode編碼值。然後，我們就可以使用和C#中類似的方式來判斷字符是否屬於CJK或CJK擴展字符集中的漢字。
' 判斷一個字符是否是CJK或CJK擴展字符集中的漢字
Public Function IsChineseCharacter(c As String) As Boolean
'    chatGPT大菩薩： Unicode範圍: CJK字符集範圍：4E00–9FFF，CJK擴展字符集範圍：20000–2A6DF 孫守真按：這樣根本不夠，只有 CJK統一表意符號和CJK擴展B
'    Dim unicodeVal As Long
'    unicodeVal = AscW(c)
'    IsChineseCharacter = (unicodeVal >= &H4E00 And unicodeVal <= &H9FFF) Or (unicodeVal >= &H20000 And unicodeVal <= &H2A6DF)
    IsChineseCharacter = IsCJK(c)(1) And IsCJK(c)(2) <> CJKBlockName.CJK_Symbols_and_Punctuation
End Function

' 判斷一個字符串是否全部由CJK或CJK擴展字符集中的漢字組成
Public Function IsChineseString(s As String) As Boolean
    Dim i As Long
    For i = 1 To Len(s)
        If IsHighSurrogate(VBA.Mid(s, i, 1)) Then
        
        ElseIf IsLowSurrogate(VBA.Mid(s, i, 1)) Then
            If Not IsChineseCharacter(VBA.Mid(s, i - 1, 2)) Then
                IsChineseString = False
                Exit Function
            End If
        Else
            If Not IsChineseCharacter(VBA.Mid(s, i, 1)) Then
                IsChineseString = False
                Exit Function
            End If
        End If
    Next i
    IsChineseString = True
End Function

Rem 20240122 Bing大菩薩：Excel中提取surrogate字元
Rem 我明白您的問題了。在處理含有代理對的字串時，確實需要特別小心，以避免將代理對的字元錯誤地切割開來。在VBA中，我們可以使用一些特殊的方法來處理這種情況。
Rem 一種可能的解決方案是使用一個自定義的函數來檢查每個字元是否為代理對的一部分。以下是一個可能的實現：
Function IsSurrogatePair(str As String, pos As Integer) As Boolean
    Dim c As Integer
    c = AscW(VBA.Mid(str, pos, 1))
    IsSurrogatePair = c >= &HD800 And c <= &HDFFF
End Function
Rem 這個函數會檢查字串中指定位置的字元是否為代理對的一部分。然後，您可以在逐字處理字串時使用這個函數來確保不會將代理對的字元切割開來。
Rem 請注意，這只是一種可能的解決方案，並且可能需要根據您的具體需求進行調整。希望這對您有所幫助！如果您有其他問題，請隨時告訴我。南無阿彌陀佛。

Rem 20230221 chatGPT大菩薩: VBA檢查surrogate字符：
Rem 在 VBA 中，您可以使用 AscW 函數將一個字符轉換為 Unicode 編碼。然後，您可以檢查該編碼是否在代理對範圍內。
Rem Unicode 中的代理對範圍為 U+D800 到 U+DFFF，共有 2048 個代碼點。代理對是一種特殊的編碼形式，由兩個 Unicode 編碼組成，用於表示較大的字符集，如 Emoji。
Rem 下面是一個示例函數，該函數接受一個字符並返回一個布爾值，指示該字符是否為代理對中的字符：
Function IsSurrogate(ByVal ch As String) As Boolean
    Dim code As Long
    code = AscW(ch)
    'IsSurrogate = (code >= &HD800 And code <= &HDFFF) 'https://ithelp.ithome.com.tw/articles/10198444#_=_
    IsSurrogate = (code >= &HD800 And code <= &HDBFF) _
                Or (code >= &HDC00 And code <= &HDFFF) 'https://ithelp.ithome.com.tw/articles/10198444#_=_
    'UTF-16 可以儲存 U+0000 至 U+10FFFF 之間的字碼，U+FFFF 以下的字碼以 2 個 byte 儲存，而 U+10000 以上的字碼，會被拆成兩個介於 D800 至 DFFF 之間的整數，
    '第一個被稱為 前導代理 (lead surrogates)，介於 D800 至 DBFF 之間，第二個被稱為 後尾代理 (trail surrogates)，介於 DC00 至 DFFF 之間，UTF-16 就是利用這兩個代理對來表示 FFFF 之外，其他輔助平面的文字。
    Rem 這個函數將字符轉換為 Unicode 編碼，並檢查該編碼是否在代理對範圍內。如果是，則函數返回 True，否則返回 False。請注意，AscW 函數只能用於 Unicode 字符串，如果您要處理 ANSI 字符串，則需要使用 Asc 函數。
End Function

Function IsHighSurrogate(ByVal ch As String) As Boolean
    Dim code As Long
    code = AscW(ch)
    IsHighSurrogate = (code >= &HD800 And code <= &HDBFF) 'https://ithelp.ithome.com.tw/articles/10198444#_=_
    'UTF-16 可以儲存 U+0000 至 U+10FFFF 之間的字碼，U+FFFF 以下的字碼以 2 個 byte 儲存，而 U+10000 以上的字碼，會被拆成兩個介於 D800 至 DFFF 之間的整數，
    '第一個被稱為 前導代理 (lead surrogates)，介於 D800 至 DBFF 之間
    
End Function
Function IsLowSurrogate(ByVal ch As String) As Boolean
    Dim code As Long
    code = AscW(ch)
    IsLowSurrogate = (code >= &HDC00 And code <= &HDFFF) 'https://ithelp.ithome.com.tw/articles/10198444#_=_
    'UTF-16 可以儲存 U+0000 至 U+10FFFF 之間的字碼，U+FFFF 以下的字碼以 2 個 byte 儲存，而 U+10000 以上的字碼，會被拆成兩個介於 D800 至 DFFF 之間的整數，
    '第二個被稱為 後尾代理 (trail surrogates)，介於 DC00 至 DFFF 之間，UTF-16 就是利用這兩個代理對來表示 FFFF 之外，其他輔助平面的文字。
    
End Function

Rem 20230224 chatGPT大菩薩：對於surrogate pair字符，應該使用Unicode標準中所述的方法將其轉換為單個代理字符。具體來說，將代理對（surrogate pair）的兩個元素分別稱為high surrogate和low surrogate。
Rem 以下是將代理對轉換為代理字符的方法:
Private Function CombineSurrogatePair(ByVal highSurrogate As String, ByVal lowSurrogate As String) As String
    CombineSurrogatePair = VBA.ChrW((AscW(highSurrogate) - &HD800&) * &H400& + (AscW(lowSurrogate) - &HDC00&) + &H10000)
End Function
Rem 使用這個函數，您可以通過在循環中處理單個字符，並使用上面的範圍來判斷字符是否在CJK全字集範圍內。 如果找到代理字符，則可以使用該函數將其轉換為Unicode字符。

Rem 計算x的字數長度
Function CharactersCount(x As String) As Long
    Dim i As Long, cntr As Long
    For i = 1 To VBA.Len(x)
        If Not IsLowSurrogate(VBA.Mid(x, i, 1)) Then cntr = cntr + 1
    Next i
    CharactersCount = cntr
End Function

Function IsCJK(c As String) As Collection
    Dim code As Long, cjk As Boolean, cjkBName As CJKBlockName, result As New Collection
    Dim unsignedCode As Long
    
    If Len(c) = 1 Then
        code = AscW(c)
    Else
        getCodePoint c, code
    End If
    
    ' 將code轉換為無符號整數 rem 20241102 Copilot大菩薩：是的，不僅在基本區域，擴充區域也會遇到類似的問題。許多擴充的CJK字符也會因為符號位的緣故被當作負數處理。為了確保這些字符在所有區域都能正確判斷，可以對所有區域的比較進行無符號處理。
                                                '這種處理方式會避免符號位影響判斷結果，不論字符在基本區域還是擴充區域。
    If code < 0 Then
        unsignedCode = code + 65536 'Copilot大菩薩 ： 在VBA中，您可以使用無符號整數來處理這些帶負號影響比較範圍的16進位值。通過以下方式可以避免負號問題：
                                                '使用 AscW 獲取字符的Unicode碼。
                                                '檢查其是否為負數，然後轉換為無符號整數。
                                                '這樣可以確保即使在處理基本區域內的字符時，也能正確地比較範圍，避免負號問題。
    Else
        unsignedCode = code
    End If
    
    If unsignedCode >= CLng(CJKChartRangeString.CJK_Unified_Ideographs_start) And unsignedCode <= CLng(CJKChartRangeString.CJK_Unified_Ideographs_end) Then
        cjk = True: cjkBName = CJKBlockName.CJK_Unified_Ideographs
    ElseIf unsignedCode >= CLng(CJKChartRangeString.CJK_Compatibility_Ideographs_start) And unsignedCode <= CLng(CJKChartRangeString.CJK_Compatibility_Ideographs_end) Then
        cjk = True: cjkBName = CJKBlockName.CJK_Compatibility_Ideographs
    ElseIf unsignedCode >= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_A_start) And unsignedCode <= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_A_end) Then
        cjk = True: cjkBName = CJKBlockName.CJK_Unified_Ideographs_Extension_A
    ElseIf unsignedCode >= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_B_start) And unsignedCode <= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_B_end) Then
        cjk = True: cjkBName = CJKBlockName.CJK_Unified_Ideographs_Extension_B
    ElseIf unsignedCode >= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_C_start) And unsignedCode <= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_C_end) Then
        cjk = True: cjkBName = CJKBlockName.CJK_Unified_Ideographs_Extension_C
    ElseIf unsignedCode >= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_D_start) And unsignedCode <= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_D_end) Then
        cjk = True: cjkBName = CJKBlockName.CJK_Unified_Ideographs_Extension_D
    ElseIf unsignedCode >= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_E_start) And unsignedCode <= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_E_end) Then
        cjk = True: cjkBName = CJKBlockName.CJK_Unified_Ideographs_Extension_E
    ElseIf unsignedCode >= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_F_start) And unsignedCode <= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_F_end) Then
        cjk = True: cjkBName = CJKBlockName.CJK_Unified_Ideographs_Extension_F
    ElseIf unsignedCode >= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_G_start) And unsignedCode <= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_G_end) Then
        cjk = True: cjkBName = CJKBlockName.CJK_Unified_Ideographs_Extension_G
    ElseIf unsignedCode >= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_H_start) And unsignedCode <= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_H_end) Then
        cjk = True: cjkBName = CJKBlockName.CJK_Unified_Ideographs_Extension_H
    ElseIf unsignedCode >= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_I_start) And unsignedCode <= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_I_end) Then
        cjk = True: cjkBName = CJKBlockName.CJK_Unified_Ideographs_Extension_I
    ElseIf unsignedCode >= CLng(CJKChartRangeString.CJK_Radicals_Supplement_start) And unsignedCode <= CLng(CJKChartRangeString.CJK_Radicals_Supplement_end) Then
        cjk = True: cjkBName = CJKBlockName.CJK_Radicals_Supplement
    ElseIf unsignedCode >= CLng(CJKChartRangeString.Kangxi_Radicals_start) And unsignedCode <= CLng(CJKChartRangeString.Kangxi_Radicals_end) Then
        cjk = True: cjkBName = CJKBlockName.Kangxi_Radicals
    ElseIf unsignedCode >= CLng(CJKChartRangeString.Ideographic_Description_Characters_start) And unsignedCode <= CLng(CJKChartRangeString.Ideographic_Description_Characters_end) Then
        cjk = True: cjkBName = CJKBlockName.Ideographic_Description_Characters
    ElseIf unsignedCode >= CLng(CJKChartRangeString.CJK_Symbols_and_Punctuation_start) And unsignedCode <= CLng(CJKChartRangeString.CJK_Symbols_and_Punctuation_end) Then
        cjk = True: cjkBName = CJKBlockName.CJK_Symbols_and_Punctuation
    ElseIf unsignedCode >= CLng(CJKChartRangeString.CJK_Strokes_start) And unsignedCode <= CLng(CJKChartRangeString.CJK_Strokes_end) Then
        cjk = True: cjkBName = CJKBlockName.CJK_Strokes
    ElseIf unsignedCode >= CLng(CJKChartRangeString.Enclosed_CJK_Letters_and_Months_start) And unsignedCode <= CLng(CJKChartRangeString.Enclosed_CJK_Letters_and_Months_end) Then
        cjk = True: cjkBName = CJKBlockName.Enclosed_CJK_Letters_and_Months
    ElseIf unsignedCode >= CLng(CJKChartRangeString.CJK_Compatibility_start) And unsignedCode <= CLng(CJKChartRangeString.CJK_Compatibility_end) Then
        cjk = True: cjkBName = CJKBlockName.CJK_Compatibility
    ElseIf unsignedCode >= CLng(CJKChartRangeString.CJK_Compatibility_Forms_start) And unsignedCode <= CLng(CJKChartRangeString.CJK_Compatibility_Forms_end) Then
        cjk = True: cjkBName = CJKBlockName.CJK_Compatibility_Forms
    ElseIf unsignedCode >= CLng(CJKChartRangeString.Enclosed_Ideographic_Supplement_start) And unsignedCode <= CLng(CJKChartRangeString.Enclosed_Ideographic_Supplement_end) Then
        cjk = True: cjkBName = CJKBlockName.Enclosed_Ideographic_Supplement
    ElseIf unsignedCode >= CLng(CJKChartRangeString.CJK_Compatibility_Ideographs_Supplement_start) And unsignedCode <= CLng(CJKChartRangeString.CJK_Compatibility_Ideographs_Supplement_end) Then
        cjk = True: cjkBName = CJKBlockName.CJK_Compatibility_Ideographs_Supplement
    Else
        cjk = False
    End If
    
    result.Add cjk
    result.Add cjkBName
    Set IsCJK = result
End Function

Function IsCJK_old(c As String) As Collection 'Boolean,CJKBlockName
    Dim code As Long, cjk As Boolean, cjkBlackName As CJKBlockName, result As New Collection
    Dim codeHex As String
'    Dim code
    Rem chatGPT大菩薩：是的，您說得沒錯。在 VBA 中，使用 AscW 函式取得 Unicode 字元的整數值時，如果傳入的字串是 surrogate pair，那麼函式只會計算 pair 的第一個字元（即 High surrogate）的值。因此，可以直接使用 AscW(c) 來計算 c 的整數值，而不必再使用 Left 函式來取得第一個字元。
    'code = AscW(vba.Left(c, 1))
    'code = AscW(c)
    If Len(c) = 1 Then
        code = AscW(c) 'AscW_IncludeSurrogatePairUnicodecode(c)
        If code < 0 Then 'Bing大菩薩：您好，這是Bing。關於您的問題，AscW 函數在 VBA 中用於獲取字符的 Unicode 編碼。然而，對於某些字符（特別是一些中文字符），AscW 可能會返回負值。這是因為 AscW 返回的是一個 16 位的有符號整數，範圍是 -32768 到 327671。當字符的 Unicode 編碼超過 32767 時，AscW 會返回一個負數23。
                        '解決這個問題的一種方法是對 AscW 返回的負值進行處理。如果 AscW 返回一個負數，您可以將該數值加上 65536 來獲得正確的 Unicode 編碼23。以下是一個修改過的函數：
            code = code + 65536
        End If
    Else
        getCodePoint c, code
    End If
    Rem https://en.wikipedia.org/wiki/CJK_characters
    'CJK Unified Ideographs
'    If code >= CLng("&H4E00") And code <= CLng("&H9FFF") Then '一定要「CLng("&H9FFF")」 不能 「CLng(&H9FFF)」
'        cjk = True: cjkBlackName = CJKBlockName.CJK_Unified_Ideographs
'    ElseIf code >= CLng("&H6300") And code <= CLng("&H77FF") Then
'        cjk = True: cjkBlackName = CJKBlockName.CJK_Unified_Ideographs
'    ElseIf code >= CLng("&H7800") And code <= CLng("&H8CFF") Then
'        cjk = True: cjkBlackName = CJKBlockName.CJK_Unified_Ideographs
'    ElseIf code >= CLng("&H8D00") And code <= CLng("&H9FFF") Then
'        cjk = True: cjkBlackName = CJKBlockName.CJK_Unified_Ideographs
        

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
    'CJK Unified Ideographs Extension I
    ElseIf code >= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_I_start) And code <= CLng(CJKChartRangeString.CJK_Unified_Ideographs_Extension_I_end) Then
        cjk = True: cjkBlackName = CJKBlockName.CJK_Unified_Ideographs_Extension_I
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
    Set IsCJK_old = result
Rem chatGPT大菩薩：抱歉，我之前回答的有誤。您提到的「元」字的Unicode碼確實是5143，屬於CJK基本集範圍內。
Rem 另外，我之前的計算是有誤的，因為將16進制轉為10進制時需要注意正負號。正確的範圍應為：
Rem CJK基本集：4E00（19968）到9FFF（40959）
Rem CJK擴展A：3400（13312）到4DBF（19871）
Rem CJK擴展B：20000（131072）到2A6DF（173791）
Rem CJK擴展C：2A700（173824）到2B73F（177983）
Rem CJK擴展D：2B740（177984）到2B81F（178207）
Rem CJK擴展E：2B820（178208）到2CEAF（235519）
Rem CJK擴展F：2CEB0（235520）到2EBEF（303231）
Rem 關於 &H9FFF 轉成負數的問題，是因為在VBA中，整數類型的最高位為符號位，如果最高位為1，則表示負數。因此，&H9FFF 將被當作負數處理，其實際值為 -24577。
End Function
Function HextoLng(hexValue As String) As Long
    'HextoLng = CLng(hexValue) And &HFFFF 'Val("&H" & VBA.Right("0000" & hexValue, 4))
    HextoLng = CLng(hexValue)
End Function

Function AscW_IncludeSurrogatePairUnicodecode(ByVal str As String) As Long
    Dim utf16 As String
    utf16 = VBA.StrConv(str, vbUnicode)
    Dim code As Long
    If Len(utf16) = 2 Then ' surrogate pair
        code = (CLng(AscW(VBA.Mid(utf16, 1, 1))) - &HD800&) * &H400& + (CLng(AscW(VBA.Mid(utf16, 2, 1))) - &HDC00&) + &H10000
    Else
        code = AscW(utf16)
    End If
    AscW_IncludeSurrogatePairUnicodecode = code
End Function
Sub getCodePoint(character As String, codePoint As Long)
' 獲取字符串的 high surrogate 和 low surrogate 的 AscW() 值
codePoint = ((CLng(AscW(VBA.Left(character, 1))) - &HD800) * &H400) + (CLng(AscW(VBA.Right(character, 1))) - &HDC00) + &H10000
Rem 沒有「CLng」轉型會溢位，若者如 isCJK_Ext()函式中的方式，以型別為 Long 的變數儲存其值，亦會隱含轉型
End Sub


Function isCJK_Ext(str As String, whatBlockNameInExt As CJKBlockName) As Boolean
Dim codePoint As Long
Dim highSurrogate As Long
Dim lowSurrogate As Long

' 獲取字符串的 high surrogate 和 low surrogate 的 AscW() 值
highSurrogate = AscW(VBA.Left(str, 1))
lowSurrogate = AscW(VBA.Right(str, 1))

If (highSurrogate >= SurrogateCodePoint.HighStart And highSurrogate <= SurrogateCodePoint.HighEnd) _
    And (lowSurrogate >= SurrogateCodePoint.LowStart And lowSurrogate <= SurrogateCodePoint.LowEnd) Then
    ' 計算字符的碼點值!!!!!!!!!!!!!!!!!
'    codepoint = ((highSurrogate - &HD800) * &H400) + (lowSurrogate - &HDC00) + &H10000
    getCodePoint str, codePoint '若沒以「CLng()」轉型會溢位，以型別為 Long 的變數儲存其值，即會隱含轉型
        
        Rem forDebugText
'    If codepoint = &H2E4E5 Then Stop
'    If Hex(codepoint) = "2E4E5" Then Stop

    Select Case whatBlockNameInExt
        Case CJKBlockName.CJK_Unified_Ideographs_Extension_A
            If codePoint >= CJKChartRange.CJK_Unified_Ideographs_Extension_A_start And codePoint <= CJKChartRange.CJK_Unified_Ideographs_Extension_A_end Then isCJK_Ext = True
        Case CJKBlockName.CJK_Unified_Ideographs_Extension_B
            If codePoint >= CJKChartRange.CJK_Unified_Ideographs_Extension_B_start And codePoint <= CJKChartRange.CJK_Unified_Ideographs_Extension_B_end Then isCJK_Ext = True
        Case CJKBlockName.CJK_Unified_Ideographs_Extension_C
            If codePoint >= CJKChartRange.CJK_Unified_Ideographs_Extension_C_start And codePoint <= CJKChartRange.CJK_Unified_Ideographs_Extension_C_end Then isCJK_Ext = True
        Case CJKBlockName.CJK_Unified_Ideographs_Extension_D
            If codePoint >= CJKChartRange.CJK_Unified_Ideographs_Extension_D_start And codePoint <= CJKChartRange.CJK_Unified_Ideographs_Extension_D_end Then isCJK_Ext = True
        Case CJKBlockName.CJK_Unified_Ideographs_Extension_E
            If codePoint >= CJKChartRange.CJK_Unified_Ideographs_Extension_E_start And codePoint <= CJKChartRange.CJK_Unified_Ideographs_Extension_E_end Then isCJK_Ext = True
        Case CJKBlockName.CJK_Unified_Ideographs_Extension_F
            If codePoint >= CJKChartRange.CJK_Unified_Ideographs_Extension_F_start And codePoint <= CJKChartRange.CJK_Unified_Ideographs_Extension_F_end Then isCJK_Ext = True
        Case CJKBlockName.CJK_Unified_Ideographs_Extension_G
            If codePoint >= CJKChartRange.CJK_Unified_Ideographs_Extension_G_start And codePoint <= CJKChartRange.CJK_Unified_Ideographs_Extension_G_end Then isCJK_Ext = True
        Case CJKBlockName.CJK_Unified_Ideographs_Extension_H
            If codePoint >= CJKChartRange.CJK_Unified_Ideographs_Extension_H_start And codePoint <= CJKChartRange.CJK_Unified_Ideographs_Extension_H_end Then isCJK_Ext = True
        Case CJKBlockName.CJK_Unified_Ideographs_Extension_I
            If codePoint >= CJKChartRange.CJK_Unified_Ideographs_Extension_I_start And codePoint <= CJKChartRange.CJK_Unified_Ideographs_Extension_I_end Then isCJK_Ext = True
    End Select
End If
End Function

Rem 20230225 chatGPT大菩薩：CJK-ext F high surrogate.：判斷 Unicode 字符是否在 CJK-Ext F 範圍內，並且計算出字符的碼點值：
Function isCJK_ExtF(str As String) As Boolean
'https://ithelp.ithome.com.tw/articles/10198444#_=_
'第一個被稱為 前導代理 (lead surrogates)，介於 D800 至 DBFF 之間
'第二個被稱為 後尾代理 (trail surrogates)，介於 DC00 至 DFFF 之間

Dim codePoint As Long
Dim highSurrogate As Long
Dim lowSurrogate As Long

' 獲取字符串的 high surrogate 和 low surrogate 的 AscW() 值
highSurrogate = AscW(VBA.Left(str, 1))
lowSurrogate = AscW(VBA.Right(str, 1))

If (highSurrogate >= &HD84D And highSurrogate <= &HDBFF) And (lowSurrogate >= &HDC00 And lowSurrogate <= &HDFFF) Then
    ' 計算字符的碼點值
    codePoint = ((highSurrogate - &HD800) * &H400) + (lowSurrogate - &HDC00) + &H10000
    
    If codePoint >= &H2CEB0 And codePoint <= &H2EBEF Then
        ' 字符在 CJK-Ext F 範圍內
        isCJK_ExtF = True
    Else
        ' 字符不在 CJK-Ext F 範圍內
    End If
Else
    ' 字符不在 CJK-Ext F 範圍內
End If
'代碼邏輯如下:
'
'先獲取字符串的 high surrogate 和 low surrogate 的 AscW() 值。
'如果 high surrogate 和 low surrogate 的 AscW() 值都在 CJK-Ext F 範圍內，則計算字符的碼點值。
'判斷字符的碼點值是否在 CJK-Ext F 範圍內，如果在，則說明字符在 CJK-Ext F 範圍內；如果不在，則說明字符不在 CJK-Ext F 範圍內。
'計算字符的碼點值的公式如下:
'
'codePoint = ((highSurrogate - &HD800) * &H400) + (lowSurrogate - &HDC00) + &H10000
'
'其中，&HD800 和 &HDC00 分別是 high surrogate 和 low surrogate 的基準值，&H400 是 surrogate pair 的偏移量，&H10000 是 Unicode 編碼的基準值。

End Function


Rem chatGPT大菩薩:WordVBA缺字顯示:在 Word 中，按下 Alt + X 鍵可以將所選文字轉換為其對應的 Unicode 碼點，這個功能稱為 Unicode 字符輸入。
Rem 在 VBA 中，可以使用 Selection.Range.Text 或 Range.Text 屬性來獲取所選文字或範圍的內容，然後使用 Selection.Range.Text = vba.Chrw(unicode_code) 或 Range.Text = vba.Chrw(unicode_code) 來將其轉換為 Unicode 碼點所對應的字符。
Rem 以下是一個示例，展示了如何使用 VBA 在 Word 中將選定範圍的內容轉換為其 Unicode 碼點：
Sub ConvertToUnicode_SelectionToggleCharacterCode() '類似實作 Selection.ToggleCharacterCode 方法
    Dim selectedText As String
    Dim unicodeCode As Long
    
    selectedText = Selection.Range.text
    
    If Len(selectedText) = 1 Then
        unicodeCode = AscW(selectedText)
        Selection.Range.text = Hex(unicodeCode)
    ElseIf Len(selectedText) = 2 Then
        unicodeCode = (AscW(VBA.Mid(selectedText, 1, 1)) - &HD800&) * &H400& + (AscW(VBA.Mid(selectedText, 2, 1)) - &HDC00&) + &H10000 '
        getCodePoint selectedText, unicodeCode
        Selection.Range.text = Hex(unicodeCode)
    Else
        MsgBox "Invalid selection"
        Exit Sub
    End If
    
'    Selection.Range.text = vba.Chrw(unicodeCode)
    Rem chatGPT菩薩：注意，在處理 surrogate pair 時，需要將兩個代理對的 Unicode 碼點轉換為實際的 Unicode 碼點。上述示例中的代碼就是將 surrogate pair 轉換為實際的 Unicode 碼點的範例。
End Sub
Rem creedit with chatGPT大菩薩：
Function ConvertToUnicode(chartoConvert As String) As Long
    Dim unicodeCode As Long
    If Len(chartoConvert) = 1 Then
        unicodeCode = AscW(chartoConvert)
    ElseIf Len(chartoConvert) = 2 Then
        'unicodeCode = (CLng(AscW(VBA.Mid(chartoConvert, 1, 1))) - &HD800&) * &H400& + (CLng(AscW(VBA.Mid(chartoConvert, 2, 1))) - &HDC00&) + &H10000
        'unicodeCode = ((CLng(AscW(VBA.Mid(chartoConvert, 1, 1))) - &HD800)) * &H400 + (CLng(AscW(VBA.Mid(chartoConvert, 2, 1))) - &HDC00) + &H10000
        getCodePoint chartoConvert, unicodeCode
    Else
        MsgBox "Invalid character"
        Exit Function
    End If
    ConvertToUnicode = unicodeCode
    
End Function
Rem 是否是八卦、六十四卦卦形字
Function IsGuaShape(c As String) As Boolean
    IsGuaShape = (VBA.AscW(c) >= &H2630 And VBA.AscW(c) <= &H2637) Or (VBA.AscW(c) >= &H4DC0 And VBA.AscW(c) <= &H4DFF)  '易經六十四卦符號 及八卦 https://zh.wikipedia.org/wiki/%E6%98%93%E7%B6%93%E5%85%AD%E5%8D%81%E5%9B%9B%E5%8D%A6%E7%AC%A6%E8%99%9F_(Unicode%E5%8D%80%E6%AE%B5)
                                                                                                                         '八卦 https://zh.wikipedia.org/wiki/%E5%85%AB%E5%8D%A6
End Function

Rem 判斷是否是Big5碼 20240926 creedit_with_Copilot大菩薩 ：Word VBA 判斷 Big5 程式：https://sl.bing.net/cAEk03kEN4e
Function IsBig5(text As String) As Boolean
    Dim stream As Object
    Dim encodedText As String
    
    On Error GoTo ErrorHandler
    
    Set stream = CreateObject("ADODB.Stream") '這個函數使用ADODB.Stream來將字串轉換為Big5編碼，然後檢查轉換後的字串是否與原始字串相同。如果相同，則表示該字串是Big5碼。
    stream.Type = 2 ' Text
    stream.Mode = 3 ' Read/Write
    stream.Charset = "big5"
    stream.Open
    stream.WriteText text
    stream.Position = 0
    encodedText = stream.ReadText(-1)
    
    IsBig5 = (encodedText = text)
    
    stream.Close
    Set stream = Nothing
    Exit Function
    
ErrorHandler:
    IsBig5 = False
    If Not stream Is Nothing Then
        stream.Close
        Set stream = Nothing
    End If
End Function



Rem 選取文字轉成ChrW表示式 20240926
Function ConvertToChrwCode_IfNotBig5() As String
    Dim i As Byte, rng As Range, a As Range, resultA As String, result As String, big5Flag As Boolean, quoteMark As String, big5FlagPrevisou As Boolean
    Dim ur As UndoRecord, combination As String
    Set rng = Selection.Range
    For Each a In rng.Characters
'        If a.text = Chr(13) Then
'            result = result & Chr(13)
'            GoTo continue
'        End If
        If big5FlagPrevisou Then
            quoteMark = vbNullString
        Else
            quoteMark = """"
        End If
        big5Flag = IsBig5(a.text)
        If Not big5Flag Then
            If big5FlagPrevisou Then
                quoteMark = """"
            Else
                quoteMark = vbNullString
            End If
            For i = 1 To VBA.Len(a)
                resultA = resultA & "VBA.ChrW(" & VBA.AscW(VBA.Mid(a.text, i, 1)) & ")" & " & "
            Next i
            If result = vbNullString Then
                combination = vbNullString
            Else
                If big5FlagPrevisou Then
                    combination = " & "
                Else
                    combination = vbNullString
                End If
            End If
            result = result & quoteMark & combination & resultA
        Else '是Big5的字
            If result = vbNullString Then quoteMark = """"
            If big5FlagPrevisou = False Then
                result = result & """" & a.text
            Else
                result = quoteMark & result & a.text
            End If
        End If
        resultA = vbNullString
        big5FlagPrevisou = big5Flag
continue:
    Next a
    
    If big5FlagPrevisou Then result = result & """"
    If VBA.Right(result, 3) = " & " Then result = VBA.Left(result, VBA.Len(result) - 3)
    
    SystemSetup.stopUndo ur, "ConvertToChrwCode"
    SystemSetup.ClipboardPutIn result
    If Selection.Document.path = "" Then
        rng.text = result
        rng.Cut
    End If
    ConvertToChrwCode_IfNotBig5 = result
    SystemSetup.contiUndo ur
End Function
Rem 20240826 Copilot大菩薩 ： Word VBA 私人造字碼區字符搜尋 ： https://sl.bing.net/hahIGJ4sxX2
Rem BAD!!!!不能用，要再改！
Sub FindPrivateUseCharacters()
    Dim rng As Range
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .text = "[\uE000-\uF8FF]" ' 私人造字碼區的範圍
        .MatchWildcards = True
        Do While .Execute(Forward:=True) = True
            rng.Select
            MsgBox "找到私人造字碼區的字符: " & rng.text
            rng.Collapse Direction:=wdCollapseEnd
        Loop
    End With
End Sub

Rem 20240826 creedit_with_Copilot大菩薩：要遍歷這三個私人造字碼區塊來檢查某一字符是否為私人造字，可以使用 VBA 來檢查字符的 Unicode 值是否在這些區塊範圍內。
Rem ： Word VBA 私人造字碼區字符搜尋 ： https://sl.bing.net/hahIGJ4sxX2
'這段程式碼包含兩個部分:
'IsPrivateUseCharacter 函數：檢查字符是否在私人造字碼區塊範圍內。
'CheckPrivateUseCharacters 子程序：遍歷文件中的每個字符，並使用 IsPrivateUseCharacter 函數檢查是否為私人造字碼區的字符。
Function IsPrivateUseCharacter(ch As String) As Boolean
    Dim codePoint As Long
    codePoint = AscW(ch)
    
    If (codePoint >= &HE000 And codePoint <= &HF8FF) Or _
       (codePoint >= &HF0000 And codePoint <= &HFFFFF) Or _
       (codePoint >= &H100000 And codePoint <= &H10FFFF) Then
        IsPrivateUseCharacter = True
    Else
        IsPrivateUseCharacter = False
    End If
End Function

Sub CheckPrivateUseCharacters_Less10000()
    Static notPrivateUseCharacters As String
    Static rngCharactersCount As Long
    Dim rng As Range
    Dim ch As String, i As Long
    If Selection.Type = wdSelectionIP _
        Or Selection.Characters.Count = 1 Then '如果沒選取則以插入點以後的文件內容
        Set rng = ActiveDocument.Range(Selection.End, ActiveDocument.Content.End)
    Else
        Set rng = Selection.Range
    End If
    
    If notPrivateUseCharacters = vbNullString Then notPrivateUseCharacters = VBA.Chr(13)
    
    If rngCharactersCount = 0 Then rngCharactersCount = rng.Characters.Count
    
    For i = 1 To rngCharactersCount
        If i Mod 1000 = 0 Then SystemSetup.playSound 1
        If i = 2000 Then Stop
        ch = rng.Characters(i).text
        If VBA.InStr(notPrivateUseCharacters, ch) = 0 Then
            If IsPrivateUseCharacter(ch) Then
                rng.Characters(i).Select
    '            MsgBox "找到私人造字碼區的字符: " & ch
                SystemSetup.playSound 2
                Exit Sub
            Else
                notPrivateUseCharacters = notPrivateUseCharacters & ch
                'rng.text = VBA.Replace(rng.text, ch, vbNullString)
            End If
        End If
    Next i
    SystemSetup.playSound 7
End Sub
Rem 這是不保留已檢查過的字。文件大時一樣不適用
Sub CheckPrivateUseCharacters_text()
    Static processedCharacters As String
    Dim privateUseCharacters As String
    Dim i As Long, d As Document
    Dim a As Range
    Dim rng As Range
    Dim ch As String
    If Selection.Type = wdSelectionIP _
        Or Selection.Characters.Count = 1 Then '如果沒選取則以插入點以後的文件內容
        Set rng = ActiveDocument.Range(Selection.End, ActiveDocument.Content.End)
    Else
        Set rng = Selection.Range
    End If
    
    If processedCharacters = vbNullString Then
        processedCharacters = VBA.Chr(13)
    End If
    
reStart:
    For Each a In rng.Characters
        i = i + 1
        If i Mod 10 = 0 Then
            SystemSetup.playSound 1
        End If
        ch = a.text
        If VBA.InStr(processedCharacters, ch) = 0 Then
            If IsPrivateUseCharacter(ch) Then
                'a.Select
                privateUseCharacters = privateUseCharacters & ch
                processedCharacters = processedCharacters & ch
    '            MsgBox "找到私人造字碼區的字符: " & ch
                SystemSetup.playSound 2
                rng.text = VBA.Replace(rng.text, ch, vbNullString)
                GoTo reStart
'                Exit Sub
            Else
                processedCharacters = processedCharacters & ch
                rng.text = VBA.Replace(rng.text, ch, vbNullString)
                GoTo reStart
            End If
        Else
            
            rng.text = VBA.Replace(rng.text, ch, vbNullString)
            GoTo reStart
            
        End If
    
    Next a
    SystemSetup.playSound 7
    processedCharacters = vbNullString
    Set d = Documents.Add
    d.Range.text = privateUseCharacters
    d.SaveAs2 rng.Document.path & "\" & "PrivateUseCharacters" & ".docx"
    rng.Document.Close wdDoNotSaveChanges
End Sub

Rem 逐字瀏覽檢視所造字 20240829
Sub CheckPrivateUseCharacters()
    Static processedCharacters As String
    Dim i As Long
    Dim a As Range
    Dim rng As Range
    Dim ch As String
    If Selection.Type = wdSelectionIP _
        Or Selection.Characters.Count = 1 Then '如果沒選取則以插入點以後的文件內容
        Set rng = ActiveDocument.Range(Selection.End, ActiveDocument.Content.End)
    Else
        Set rng = Selection.Range
    End If
    
    If processedCharacters = vbNullString Then
        processedCharacters = VBA.Chr(13)
    End If
    
    For Each a In rng.Characters
        i = i + 1
        If i Mod 60000 = 0 Then
            SystemSetup.playSound 1
        End If
        ch = a.text
        If VBA.InStr(processedCharacters, ch) = 0 Then
            If IsPrivateUseCharacter(ch) Then
                a.Select
                a.HighlightColorIndex = wdYellow
                a.Document.ActiveWindow.ScrollIntoView a
'                MsgBox "找到私人造字碼區的字符: " & ch
                SystemSetup.playSound 1
                Exit Sub
            Else
                processedCharacters = processedCharacters & ch
            End If
            
        End If
    
    Next a
    SystemSetup.playSound 7
    processedCharacters = vbNullString
End Sub
Rem 找出造字 - 找出使用中的文件裡的造字並匯出成新檔，檔名為 原檔名+"PrivateUseCharacters.docx"
Sub PrivateUseCharactersOutput()
    Static processedCharacters As String
    Dim privateUseCharacters As String
    Dim i As Long, d As Document
    Dim a As Range
    Dim rng As Range
    Dim ch As String
    If ActiveDocument.path = "" Then
        MsgBox "請先儲存文件再繼續！", vbCritical
        Exit Sub
    End If
    If Selection.Type = wdSelectionIP _
        Or Selection.Characters.Count = 1 Then '如果沒選取則以插入點以後的文件內容
        Set rng = ActiveDocument.Range(Selection.End, ActiveDocument.Content.End)
    Else
        Set rng = Selection.Range
    End If
    
    If processedCharacters = vbNullString Then
        processedCharacters = VBA.Chr(13)
    End If
    
    For Each a In rng.Characters
        i = i + 1
        If i Mod 100000 = 0 Then
            SystemSetup.playSound 1
        End If
        ch = a.text
        If VBA.InStr(processedCharacters, ch) = 0 Then
            If IsPrivateUseCharacter(ch) Then
                'a.Select
                privateUseCharacters = privateUseCharacters & ch & vbTab & vbCr
    '            MsgBox "找到私人造字碼區的字符: " & ch
'                SystemSetup.playSound 2
'                Exit Sub
            End If
            processedCharacters = processedCharacters & ch
        End If
    
    Next a
    SystemSetup.playSound 7
    processedCharacters = vbNullString
    If privateUseCharacters <> vbNullString Then
        Set d = Documents.Add
        d.Range.text = privateUseCharacters
        d.SaveAs2 rng.Document.path & "\" & VBA.Left(rng.Document.Name, VBA.InStr(rng.Document.Name, ".") - 1) & "PrivateUseCharacters.docx"
    '    rng.Document.Close wdDoNotSaveChanges
        d.Activate
        d.Application.Activate
    Else
        MsgBox rng.Document.Name & "文件中並沒有造字！", vbInformation
        rng.Document.Application.Activate
    End If
End Sub

Sub ReplacePrivateUserCharactersWithCJK()
    Dim dReplaced As Document, dTable As Document, r As row
    Dim wFind As String, wReplace As String, ur As UndoRecord
    Set dReplaced = Documents("201010說文A造字") ' .docx")
    Set dTable = Documents("文件3")
    
    SystemSetup.stopUndo ur, "ReplacePrivateUserCharactersWithCJK"
    word.Application.ScreenUpdating = False
    For Each r In dTable.tables(1).rows
        wFind = r.cells(1).Range.Characters(1).text
        wReplace = r.cells(2).Range.Characters(1).text
        dReplaced.Range.Find.Execute wFind, , , , , , , wdFindContinue, _
            , wReplace, wdReplaceAll
    Next
    SystemSetup.contiUndo ur
    word.Application.ScreenUpdating = True
    Set dReplaced = Nothing
    Set dTable = Nothing
End Sub
