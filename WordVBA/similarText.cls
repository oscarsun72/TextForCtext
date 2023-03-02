VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "similarText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim dict1 As New Scripting.Dictionary, dict2 As New Scripting.Dictionary, s1 As Variant, s2 As Variant
Dim similarTextCopy As String '被比較文本的副本（即比對目標物）
Dim similarWithTextCopy As String '用來比較的文本的副本（用來比對的來源）
Dim SimilarityResult As New Scripting.Dictionary '<相似度名,相似度>
Dim currentProcedureName As String '正在執行的程序名稱
Dim ProceduresNameAliasPair As New Scripting.Dictionary '記下所有執行過的程序名稱及其別名<currentProcedureName,procedureNameAlias>
Private Function Initialize(ByVal text1, ByVal text2) As similartext
If s1 = text1 And s2 = text2 Then Exit Function
Dim w, str1, str2
If s1 <> text1 Then
    If dict1.Count > 0 Then dict1.RemoveAll
    s1 = text1
    If VBA.VarType(s1) = vbString Then
        Rem 對於surrogate 已修正，待測試，觀察中 20230227
        str1 = 文字處理.SplitWithoutDelimiter_StringToStringArray((s1)) '強制以傳值方式傳遞引數 https://learn.microsoft.com/zh-tw/dotnet/visual-basic/programming-guide/language-features/procedures/how-to-force-an-argument-to-be-passed-by-value
    Else
        str1 = s1
    End If
    For Each w In str1
        If Not dict1.Exists(w) Then
            dict1.Add w, 1
        Else
            dict1.item(w) = dict1.item(w) + 1
        End If
    Next w
End If
If s2 <> text2 Then
    If dict2.Count > 0 Then dict2.RemoveAll
    s2 = text2
    If VBA.VarType(s2) = vbString Then
        str2 = 文字處理.SplitWithoutDelimiter_StringToStringArray((s2))
    Else
        str2 = s2
    End If
    For Each w In str2
        If Not dict2.Exists(w) Then
            dict2.Add w, 1
        Else
            dict2.item(w) = dict2.item(w) + 1
        End If
    Next w
End If
End Function
'判斷文本間是否相似
Function Similarity(ByVal text1, ByVal text2) As Boolean 'text1 被比對者 text2 拿來比對者
If IsEmpty(text2) Or IsEmpty(text1) Then Exit Function
If SimilarityResult.Count > 0 Then SimilarityResult.RemoveAll
similarWithTextCopy = text2
Initialize text1, text2
If Me.CompareDictionaries(dict1, dict2) Then
    Similarity = True
ElseIf Me.CompareDictionarySimilarity(dict1, dict2) > 0.8 Then
    Similarity = True
ElseIf jaccard_Similarity() > 0.8 Then
    Similarity = True
ElseIf cosine_Similarity() > 0.92 Then 'cosine餘弦方法好像在文字多時會失誤，待觀察！ 感恩感恩　南無阿彌陀佛 20230227
    Similarity = True
End If
If Similarity Then
    similarTextCopy = text1
Else
    similarTextCopy = ""
End If
End Function

'回傳文本間相似的百分比
Function SimilarityPercent(ByVal text1, ByVal text2) As Variant 'text1 被比對者 text2 拿來比對者
If IsEmpty(text2) Or IsEmpty(text1) Then Exit Function
If SimilarityResult.Count > 0 Then SimilarityResult.RemoveAll
similarWithTextCopy = text2
Dim sim As Double
Initialize text1, text2
If Me.CompareDictionaries(dict1, dict2) Then
    sim = 1
Else
    sim = (Me.CompareDictionarySimilarity(dict1, dict2) + cosine_Similarity() + jaccard_Similarity()) / 3
End If
Rem chatGPT大菩薩：在 SimilarityPercent 函數中，應使用 Round 函數的第二個參數，以指定小數點後保留的位數。例如，要保留一位小數，可以使用 VBA.Round(sim * 100, 1)。
SimilarityPercent = VBA.Round(sim * 100, 2)
End Function
Rem 20230222 creedit with chatGPT大菩薩：比較 Dictionary 對象：
Rem 在 VBA 中，比較兩個 Dictionary 對象的相似度或完全相同，需要分別比較它們的鍵和值。以下是一些比較方法：
Rem 1.  比較兩個 Dictionary 是否完全相同:
'Function CompareDictionaries(ByVal dict1 As Object, ByVal dict2 As Object) As Boolean
Function CompareDictionaries(ByVal dict1 As Scripting.Dictionary, ByVal dict2 As Scripting.Dictionary) As Boolean
currentProcedureName = "CompareDictionaries"
ProceduresNameAliasPair(currentProcedureName) = "用字相似度"
    If dict1.Count = 0 Or dict2.Count = 0 Then Exit Function
    If dict1.Count <> dict2.Count Then
        CompareDictionaries = False
        Exit Function
    End If
    
    Dim key As Variant
    For Each key In dict1.keys
        If Not dict2.Exists(key) Then
            CompareDictionaries = False
            Exit Function
        End If
        If Not dict1(key) = dict2(key) Then
            CompareDictionaries = False
            Exit Function
        End If
    Next
    
    CompareDictionaries = True
    Rem 上述代碼會檢查兩個 Dictionary 的鍵值對數量是否相同，以及它們的鍵和值是否完全匹配。如果兩個 Dictionary 完全相同，則返回 True，否則返回 False。
End Function

Rem 2.  比較兩個 Dictionary 的相似度：
'Function CompareDictionarySimilarity(ByVal dict1 As Object, ByVal dict2 As Object) As Double
Function CompareDictionarySimilarity(ByVal dict1 As Scripting.Dictionary, ByVal dict2 As Scripting.Dictionary) As Double
currentProcedureName = "CompareDictionarySimilarity"
ProceduresNameAliasPair(currentProcedureName) = "用字相似度"
    If dict1.Count = 0 Or dict2.Count = 0 Then
        SimilarityResult(ProceduresNameAliasPair(currentProcedureName)) = 0 ': CompareDictionarySimilarity = 0 'The default value of Double is 0. https://learn.microsoft.com/en-us/dotnet/visual-basic/language-reference/data-types/double-data-type
        Exit Function
    End If
    Dim totalKeys As Long
    Dim matchedKeys As Long
    
    totalKeys = dict1.Count
    Rem 前已作判斷，以離開函式
'    If totalKeys = 0 Then '如果是空字串，則無從比對、亦無所謂相似度
'        CompareDictionarySimilarity = 0: similarityResult = 0 '原為 1 ，今更正。待測試後即刪除此注解
'        Exit Function
'    End If
    
    Dim key As Variant
    For Each key In dict1.keys
        If dict2.Exists(key) Then '若 dict中沒有該鍵值key，則在執行 dict(key)時，會自動添入該鍵值對之元素，更動dict物件
            If dict1(key) = dict2(key) Then matchedKeys = matchedKeys + 1
        End If
    Next
    SimilarityResult(ProceduresNameAliasPair(currentProcedureName)) = matchedKeys / totalKeys
    CompareDictionarySimilarity = SimilarityResult(ProceduresNameAliasPair(currentProcedureName))
    Rem 上述代碼會計算兩個 Dictionary 的相似度。它會檢查兩個 Dictionary 中有多少鍵值對是相同的，然後計算它們的比例。如果兩個 Dictionary 完全相同，則返回 1。如果兩個 Dictionary 沒有任何相同的鍵值對，則返回 0。
End Function

Rem 20230227用字及其字距相似度(unfinished)
Function CompareDictionarySimilarityDistance(ByVal dict1 As Scripting.Dictionary, ByVal dict2 As Scripting.Dictionary) As Double
currentProcedureName = "CompareDictionarySimilarityDistance"
ProceduresNameAliasPair(currentProcedureName) = "用字及其字距相似度"
    If dict1.Count = 0 Or dict2.Count = 0 Then
        SimilarityResult(ProceduresNameAliasPair(currentProcedureName)) = 0 ': CompareDictionarySimilarity = 0 'The default value of Double is 0. https://learn.microsoft.com/en-us/dotnet/visual-basic/language-reference/data-types/double-data-type
        Exit Function
    End If
    Dim totalKeys As Long
    Dim matchedKeys As Long

    totalKeys = dict1.Count
    Rem 前已作判斷，以離開函式
'    If totalKeys = 0 Then '如果是空字串，則無從比對、亦無所謂相似度
'        CompareDictionarySimilarity = 0: similarityResult = 0 '原為 1 ，今更正。待測試後即刪除此注解
'        Exit Function
'    End If

    Dim key As Variant
    For Each key In dict1.keys
        If dict2.Exists(key) Then '若 dict中沒有該鍵值key，則在執行 dict(key)時，會自動添入該鍵值對之元素，更動dict物件
            If dict1(key) = dict2(key) Then matchedKeys = matchedKeys + 1
        End If
    Next
    SimilarityResult(ProceduresNameAliasPair(currentProcedureName)) = matchedKeys / totalKeys
    CompareDictionarySimilarityDistance = SimilarityResult(ProceduresNameAliasPair(currentProcedureName))
    Rem 上述代碼會計算兩個 Dictionary 的相似度。它會檢查兩個 Dictionary 中有多少鍵值對是相同的，然後計算它們的比例。如果兩個 Dictionary 完全相同，則返回 1。如果兩個 Dictionary 沒有任何相同的鍵值對，則返回 0。
End Function

Rem 比較兩個文本的相似度可以使用餘弦相似度或者Jaccard相似度等算法。
Rem "餘弦相似度"和"Jaccard相似度"是常用的文本相似度算法，用於比較兩個文本之間的相似程度。
Rem "餘弦相似度"是通過計算兩個文本向量的夾角餘弦值來衡量它們之間的相似度。向量夾角餘弦值越大，表示兩個向量越接近，相似度也就越高。具體實作上，可以先將兩個文本表示成詞頻向量，再計算兩個向量的夾角餘弦值。相似度值的範圍在0到1之間，值越大表示相似度越高。
Rem "Jaccard相似度"是通過計算兩個文本的交集與並集之間的比例來衡量它們之間的相似度。相似度值的範圍也在0到1之間，值越大表示相似度越高。具體實作上，可以先將兩個文本表示成詞集合，然後計算兩個集合的交集與並集，最後用交集除以並集來得到相似度值。
Rem 以下是VBA程式碼示例 , 用於計算兩個文本的餘弦相似度和Jaccard相似度:
Function CosineSimilarity(s1 As Variant, s2 As Variant) As Double
    Dim dict1 As Object, dict2 As Object
    Set dict1 = CreateObject("Scripting.Dictionary")
    Set dict2 = CreateObject("Scripting.Dictionary")
    
    Dim w As Variant
    For Each w In s1
        If Not dict1.Exists(w) Then
            dict1.Add w, 1
        Else
            dict1.item(w) = dict1.item(w) + 1
        End If
    Next w
    
    For Each w In s2
        If Not dict2.Exists(w) Then
            dict2.Add w, 1
        Else
            dict2.item(w) = dict2.item(w) + 1
        End If
    Next w
    
    Dim dotProduct As Double, mag1 As Double, mag2 As Double, word As Variant
    For Each word In dict1
        If dict2.Exists(word) Then
            dotProduct = dotProduct + dict1.item(word) * dict2.item(word)
        End If
        mag1 = mag1 + dict1.item(word) ^ 2
    Next word
    
    For Each word In dict2
        mag2 = mag2 + dict2.item(word) ^ 2
    Next word
    
    mag1 = Sqr(mag1)
    mag2 = Sqr(mag2)
    
    If mag1 * mag2 = 0 Then
        CosineSimilarity = 0
    Else
        CosineSimilarity = dotProduct / (mag1 * mag2)
    End If
End Function


Function JaccardSimilarity(s1 As Variant, s2 As Variant) As Double
    Dim dict1 As Object, dict2 As Object
    Set dict1 = CreateObject("Scripting.Dictionary")
    Set dict2 = CreateObject("Scripting.Dictionary")
    
    Dim w As Variant
    For Each w In s1
        If Not dict1.Exists(w) Then
            dict1.Add w, 1
        End If
    Next w
    
    For Each w In s2
        If Not dict2.Exists(w) Then
            dict2.Add w, 1
        End If
    Next w
    
    Dim intersection As Double, union As Double, word As Variant
    For Each word In dict1
        If dict2.Exists(word) Then
            intersection = intersection + 1
        End If
        union = union + 1
    Next word
    
    For Each word In dict2
        union = union + 1
    Next word
    
    If union = 0 Then
        JaccardSimilarity = 0
    Else
        JaccardSimilarity = intersection / union
    End If
End Function

Private Function cosine_Similarity() As Double
If dict1.Count = 0 Or dict2.Count = 0 Then Exit Function
currentProcedureName = "cosine_Similarity"
ProceduresNameAliasPair(currentProcedureName) = "餘弦相似度"
'Private Function Cosine_Similarity(s1 As Variant, s2 As Variant) As Double
'    Dim dict1 As Object, dict2 As Object
'    Set dict1 = CreateObject("Scripting.Dictionary")
'    Set dict2 = CreateObject("Scripting.Dictionary")
'
'    Dim w As Variant
'    For Each w In s1
'        If Not dict1.Exists(w) Then
'            dict1.Add w, 1
'        Else
'            dict1.item(w) = dict1.item(w) + 1
'        End If
'    Next w
'
'    For Each w In s2
'        If Not dict2.Exists(w) Then
'            dict2.Add w, 1
'        Else
'            dict2.item(w) = dict2.item(w) + 1
'        End If
'    Next w
    
    Dim dotProduct As Double, mag1 As Double, mag2 As Double, word As Variant
    For Each word In dict1
        If dict2.Exists(word) Then
            dotProduct = dotProduct + dict1.item(word) * dict2.item(word)
        End If
        mag1 = mag1 + dict1.item(word) ^ 2
    Next word
    
    For Each word In dict2
        mag2 = mag2 + dict2.item(word) ^ 2
    Next word
    
    mag1 = Sqr(mag1)
    mag2 = Sqr(mag2)
    
    If mag1 * mag2 = 0 Then
         SimilarityResult(ProceduresNameAliasPair(currentProcedureName)) = 0 ':cosine_Similarity = 0' The default value of Double is 0.
    Else
        SimilarityResult(ProceduresNameAliasPair(currentProcedureName)) = dotProduct / (mag1 * mag2)
        cosine_Similarity = SimilarityResult(ProceduresNameAliasPair(currentProcedureName))
    End If
End Function
Private Function jaccard_Similarity() As Double
If dict1.Count = 0 Or dict2.Count = 0 Then Exit Function
currentProcedureName = "jaccard_Similarity"
ProceduresNameAliasPair(currentProcedureName) = "jaccard相似度"
'Private Function jaccard_Similarity(s1 As Variant, s2 As Variant) As Double
'    Dim dict1 As Object, dict2 As Object
'    Set dict1 = CreateObject("Scripting.Dictionary")
'    Set dict2 = CreateObject("Scripting.Dictionary")
'
'    Dim w As Variant
'    For Each w In s1
'        If Not dict1.Exists(w) Then
'            dict1.Add w, 1
'        End If
'    Next w
'
'    For Each w In s2
'        If Not dict2.Exists(w) Then
'            dict2.Add w, 1
'        End If
'    Next w
'
    Dim intersection As Double, union As Double, word As Variant
    For Each word In dict1
        If dict2.Exists(word) Then
            intersection = intersection + 1
        End If
        union = union + 1
    Next word
    
    For Each word In dict2
        union = union + 1
    Next word
    
    If union = 0 Then
        SimilarityResult(ProceduresNameAliasPair(currentProcedureName)) = 0 ':jaccard_Similarity = 0'The default value of Double is 0.
    Else
        SimilarityResult(ProceduresNameAliasPair(currentProcedureName)) = intersection / union
        jaccard_Similarity = SimilarityResult(ProceduresNameAliasPair(currentProcedureName))
    End If
End Function
Rem 20230222 chatGPT大菩薩：
Rem 這個實作看起來很不錯！以下是我建議的一些小改進：
Rem 在 Initialize 函數的開頭，應將傳入的 text1 和 text2 轉換為小寫，以便在比較單詞時忽略大小寫。
Rem 在 SimilarityPercent 函數中，應將變數 sim 的數據類型聲明為 Double，以便更準確地計算相似度。
Rem 在 SimilarityPercent 函數中，應使用 Round 函數的第二個參數，以指定小數點後保留的位數。例如，要保留一位小數，可以使用 VBA.Round(sim * 100, 1)。
Rem 這些改進只是一些小細節 , 你的實作已經非常好了!


Public Property Get SimilarityResults() As Scripting.Dictionary
    Set SimilarityResults = SimilarityResult
End Property

Public Property Get SimilarityResultsString() As String
Dim resultStr As String, key
For Each key In SimilarityResult.keys
    resultStr = resultStr & key & space(2) & VBA.vbTab & VBA.CStr(VBA.Round(SimilarityResult(key), 2) * 100) & "％" & vbCr
Next key
SimilarityResultsString = VBA.Left(resultStr, Len(resultStr) - Len(vbCr))
End Property



Rem(unfinished)
'Rem 20230227 Bing菩薩：文本相似度比對：在VBA中實現編輯距離算法需要使用一些循環和陣列操作 (blog.csdn.net)。以下是一個可能的代碼示例：
'Rem https://blog.csdn.net/dyfgs/article/details/7177088
'Private Function EditDistance() As Integer
'Dim len_s1 As Integer 'the length of string 1
'Dim len_s2 As Integer 'the length of string 2
'Dim i As Integer 'the loop variable for string 1
'Dim j As Integer 'the loop variable for string 2
'Dim cost As Integer 'the cost of substitution
'Dim matrix As Integer 'the matrix to store the edit distance values
'
'len_s1 = Len(len_s1)
'len_s2 = Len(len_s2)
'
'If len_s1 = 0 Then 'if string 1 is empty, return the length of string 2 as edit distance
'EditDistance = len_s2
'Exit Function
'End If
'
'If len_s2 = 0 Then 'if string 2 is empty, return the length of string 1 as edit distance
'EditDistance = len_s1
'Exit Function
'End If
'
'ReDim matrix(0 To len_s1, 0 To len_s2) 'initialize the matrix with appropriate size
'
'For i = 0 To len_s1 'fill in the first column with i values
'matrix(i, 0) = i
'Next i
'
'For j = 0 To len_s2 'fill in the first row with j values
'matrix(0, j) = j
'Next j
'
'For i = 1 To len_s1
'For j = 1 To len_s2
'If Mid(s1, i, 1) = Mid(s2, j, 1) Then 'if the characters are equal, no cost is needed
'cost = 0
'Else
'cost = 1 'otherwise, a substitution cost is needed
'End If
'
'matrix(i, j) = min(matrix(i - 1, j) + 1, matrix(i, j - 1) + 1, matrix(i - 1, j - 1) + cost) 'find the minimum value among insertion, deletion and substitution
'Next j
'Next i
'
''EditDistance=matrix(len_ s_ _ _ _ _ _ _ _ _ _ _) '_return_the_final_value_in_the_bottom_right_corner_of_the_matrix_
'
'End Function
'
'Private Function min(a, b, c) As Integer '_a_helper_function_to_find_the_minimum_among_three_values_
'min = a '_assume_a_is_the_minimum_
'If b < min Then min = b '_compare_with_b_and_update_if_needed_
'If c < min Then min = c '_compare_with_c_and_update_if_needed_
'End Function
Rem end 20230227 Bing菩薩：

Rem 待處理：https://blog.csdn.net/dyfgs/article/details/7177088
Private Function min(one As Integer, two As Integer, three As Integer)
    min = one
    If (two < min) Then
     min = two
    End If
    If (three < min) Then
     min = three
    End If
End Function
 
Private Function ld(str1 As String, str2 As String)
Dim n, m, i, j As Integer
Dim ch1, ch2 As String
    n = Len(str1)
    m = Len(str2)
    Dim temp As Integer
    If (n = 0) Then
        ld = m
    End If
    If (m = 0) Then
        ld = n
    End If
Dim d As Variant
ReDim d(n + 1, m + 1) As Variant
    For i = 0 To n
        d(i, 0) = i
    Next i
    For j = 0 To m
        d(0, j) = j
    Next j
    For i = 1 To n
        ch1 = Mid(str1, i, 1)
        For j = 1 To m
            ch2 = Mid(str2, j, 1)
            If (ch1 = ch2) Then
            temp = 0
            Else
                temp = 1
            End If
            d(i, j) = min(d(i - 1, j) + 1, d(i, j - 1) + 1, d(i - 1, j - 1) + temp)
        Next j
    Next i
    ld = d(n, m)
End Function
 
Public Function sim(str1 As String, str2 As String)
    Dim ldint As Integer
    ldint = ld(str1, str2)
    Dim strlen As Integer
    If (Len(str1) >= Len(str2)) Then
        strlen = Len(str1)
    Else
        strlen = Len(str2)
    End If
    sim = 1 - ldint / strlen
End Function
Rem end 以上尚待處理：https://blog.csdn.net/dyfgs/article/details/7177088 感恩感恩　南無阿彌陀佛



Private Sub Class_Terminate()
Set SimilarityResult = Nothing
Set dict1 = Nothing
Set dict2 = Nothing
s1 = Empty: s2 = Empty
similarTextCopy = Empty
similarWithTextCopy = Empty
Set ProceduresNameAliasPair = Nothing
currentProcedureName = Empty
End Sub
