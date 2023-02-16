VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "components"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Sub getComponentsCollection(ByRef componentsCl As Collection, c As Range)
Dim components() As String, a, i As Byte
                    '配置部件陣列以備加入colleciton 容器
                    getComponentsArray components, c
                    componentsCl.Add components
End Sub


Sub getComponentsArray(ByRef components() As String, c As Range)
Dim a, i As Byte
        ReDim components(c.Characters.Count - 2)
        '如果沒有圖片
        If c.InlineShapes.Count = 0 Then
            For Each a In c.Characters
                ' '排除儲存格字元
                If InStr(Chr(13) & Chr(7), a) = 0 Then
                    components(i) = a.Text
                    i = i + 1
                End If
            Next
        '如果有圖片
        Else
            For Each a In c.Characters
                '排除儲存格字元
                If InStr(Chr(13) & Chr(7), a) = 0 Then
                    If a.InlineShapes.Count = 0 Then '非圖片
                        components(i) = a.Text
                    Else '圖片
                        components(i) = a.InlineShapes(1).AlternativeText
                    End If
                    i = i + 1
                End If
            Next a
        End If
                    
End Sub

Rem creedit with chatGPT大菩薩
Function JoinDictionaryValues(dict As Scripting.Dictionary) As String
    Dim result As String
    Dim key As Variant
    For Each key In dict.Keys
        If InStr(dict(key), ",") > 0 Then '以逗號為元素間的間隔符
            result = result & key & VBA.vbTab & dict(key) & VBA.vbNewLine
        End If
    Next key
    If Len(result) > 0 Then
        JoinDictionaryValues = VBA.Replace(Left(result, Len(result) - 1), ",", "") '清除作為元素間間隔符的逗號
    End If
End Function

Rem creedit with chatGPT大菩薩
Sub SortStringArray(ByRef arr() As String)
'arr = Array("apple", "banana", "cherry")
QuickSort arr, LBound(arr), UBound(arr) 'chatGPT大菩薩：LBound 是一個 VBA 函數，它會回傳陣列的下界（Lower Bound），也就是陣列的第一個元素的索引。在大多數情況下，陣列的下界是 0，但有時也可以定義成其他數字。例如，如果定義了一個索引為 1 到 10 的陣列，那麼 LBound 的值就是 1。
'For Each s In arr
'Debug.Print s
'Next s
End Sub

Private Sub QuickSort(ByRef arr() As String, ByVal l As Long, ByVal r As Long) 'l=left,r=right chatGPT大菩薩：是的，在這個 QuickSort 函數中， l 參數代表左邊的索引位置，而 r 參數代表右邊的索引位置。這些參數是快速排序算法的主要部分，用於指定排序的邊界。在這個函數中， arr 數組是要進行排序的數組， l 和 r 指定了要進行排序的數組的區間。
If l >= r Then Exit Sub
Dim i As Long, j As Long, x As String
i = l: j = r: x = arr((l + r) \ 2)
'Do
'    While arr(i) < x
'        i = i + 1
'    Wend
'    While x < arr(j)
'    j = j - 1
'    Wend
'    If i <= j Then
'    Swap arr(i), arr(j)
'    i = i + 1
'    j = j - 1
'    End If
'Loop Until i > j
Do
    While StrComp(arr(i), x, vbTextCompare) < 0
    i = i + 1
    Wend
    While StrComp(x, arr(j), vbTextCompare) < 0
    j = j - 1
    Wend
    If i <= j Then
        Swap arr(i), arr(j)
        i = i + 1
        j = j - 1
    End If
Loop Until i > j
QuickSort arr, l, j
QuickSort arr, i, r
End Sub

Private Sub Swap(ByRef a As String, ByRef b As String)
Dim temp As String
temp = a
a = b
b = temp
End Sub

Rem creedit with chatGPT大菩薩
Function CompareArrays(Array1 As Variant, Array2 As Variant) As Boolean
    Dim Dict1 As Object, Dict2 As Object, element, key
    Set Dict1 = CreateObject("Scripting.Dictionary")
    Set Dict2 = CreateObject("Scripting.Dictionary")

    For Each element In Array1
        If Not Dict1.Exists(element) Then
            Dict1.Add element, 1
        Else
            Dict1(element) = Dict1(element) + 1
        End If
    Next

    For Each element In Array2
        If Not Dict2.Exists(element) Then
            Dict2.Add element, 1
        Else
            Dict2(element) = Dict2(element) + 1
        End If
    Next

    If Dict1.Count <> Dict2.Count Then
        CompareArrays = False
        Exit Function
    End If

    For Each key In Dict1.Keys
        'If Not Dict2.Exists(Key) Or Dict2(Key) <> Dict1(Key) Then
        If Not Dict2.Exists(key) Or VBA.StrComp(Dict2(key), Dict1(key)) <> 0 Then
            CompareArrays = False
            Exit Function
        End If
    Next

    CompareArrays = True
End Function

