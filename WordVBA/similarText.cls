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
Private Function Initialize(ByVal text1, ByVal text2) As similartext
If s1 = text1 And s2 = text2 Then Exit Function
Dim w, str1, str2
If s1 <> text1 Then
    If dict1.Count > 0 Then dict1.RemoveAll
    s1 = text1
    If VBA.VarType(s1) = vbString Then
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
Function Similarity(ByVal text1, ByVal text2) As Boolean
If IsEmpty(text2) Or IsEmpty(text1) Then Exit Function
Initialize text1, text2
If Me.CompareDictionaries(dict1, dict2) Then
    Similarity = True: Exit Function
End If
If Me.CompareDictionarySimilarity(dict1, dict2) > 0.8 Then
    Similarity = True: Exit Function
End If
If cosine_Similarity() > 0.8 Then
    Similarity = True: Exit Function
End If
If jaccard_Similarity() > 0.8 Then
    Similarity = True: Exit Function
End If
End Function

'回傳文本間相似的百分比
Function SimilarityPercent(ByVal text1, ByVal text2) As Variant
If IsEmpty(text2) Or IsEmpty(text1) Then Exit Function
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
    If dict1.Count = 0 Or dict2.Count = 0 Then Exit Function
    Dim totalKeys As Integer
    Dim matchedKeys As Integer
    
    totalKeys = dict1.Count
    If totalKeys = 0 Then
        CompareDictionarySimilarity = 1
        Exit Function
    End If
    
    Dim key As Variant
    For Each key In dict1.keys
        If dict2.Exists(key) Then '若 dict中沒有該鍵值key，則在執行 dict(key)時，會自動添入該鍵值對之元素，更動dict物件
            If dict1(key) = dict2(key) Then matchedKeys = matchedKeys + 1
        End If
    Next
    
    CompareDictionarySimilarity = matchedKeys / totalKeys
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
        cosine_Similarity = 0
    Else
        cosine_Similarity = dotProduct / (mag1 * mag2)
    End If
End Function
Private Function jaccard_Similarity() As Double
If dict1.Count = 0 Or dict2.Count = 0 Then Exit Function
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
        jaccard_Similarity = 0
    Else
        jaccard_Similarity = intersection / union
    End If
End Function


Rem 20230222 chatGPT大菩薩：
Rem 這個實作看起來很不錯！以下是我建議的一些小改進：
Rem 在 Initialize 函數的開頭，應將傳入的 text1 和 text2 轉換為小寫，以便在比較單詞時忽略大小寫。
Rem 在 SimilarityPercent 函數中，應將變數 sim 的數據類型聲明為 Double，以便更準確地計算相似度。
Rem 在 SimilarityPercent 函數中，應使用 Round 函數的第二個參數，以指定小數點後保留的位數。例如，要保留一位小數，可以使用 VBA.Round(sim * 100, 1)。
Rem 這些改進只是一些小細節 , 你的實作已經非常好了!
