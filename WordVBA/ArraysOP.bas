Attribute VB_Name = "ArraysOP"
Option Explicit

Rem 20230328 Adrenaline ：
Rem 漢字會照部首再筆畫排序
Public Sub SortArray_QuickSort(arrayToSort As Variant) 'VBA引數預設為傳址（pass by reference）
'    Dim arrayToSort As Variant
'    Dim i As Integer
'
'    ' 取得陣列
'    arrayToSort = Application.Transpose(ExistedNumColumnRange.value)
'
    ' 用 QuickSort 排序
    'Call QuickSortArray(arrayToSort, 1, UBound(arrayToSort))
    Call QuickSortArray(arrayToSort, LBound(arrayToSort), UBound(arrayToSort))
    
'    ' 輸出排序後的結果
'    Debug.Print "排序後的結果："
'    For i = 1 To UBound(arrayToSort)
'        Debug.Print arrayToSort(i)
'    Next i
End Sub

Private Sub QuickSortArray(ByRef arr As Variant, ByVal left As Long, ByVal right As Long)
    Dim i As Long
    Dim j As Long
    Dim pivot As Variant
    Dim temp As Variant
    
    i = left
    j = right
    pivot = arr((left + right) \ 2)
    
    While i <= j
        While arr(i) < pivot And i < right
            i = i + 1
        Wend
        
        While pivot < arr(j) And j > left
            j = j - 1
        Wend
        
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Wend
    
    If left < j Then
        Call QuickSortArray(arr, left, j)
    End If
    
    If i < right Then
        Call QuickSortArray(arr, i, right)
    End If
End Sub

Rem creedit with chatGPT大菩薩
Rem 漢字會照筆畫再部首排序
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

Rem Bing大菩薩'https://www.notion.so/Characters-76ccb4ff823e4a82b0d0af042e5a650b?pvs=4#d7f45c8d4863487db4d92e4cb7787525
'如果只保留漢字中文排序，則 hanOnly=true
Function CharactersToArray(myRange As Range, Optional hanOnly As Boolean = False) As String()

    Dim myArray() As String, arr, e, xRng As String
    Dim i As Long

    If hanOnly Then
        arr = SplitWithoutDelimiter_StringToStringArray(PunctuationString)
        xRng = myRange.text
        For Each e In arr
            xRng = VBA.Replace(xRng, e, "")
        Next e
        myRange.text = VBA.Replace(xRng, Chr(13), "")
    End If
        
    ReDim myArray(1 To myRange.Characters.Count)
    
    For i = 1 To myRange.Characters.Count
        myArray(i) = myRange.Characters(i)
    Next i
    
    CharactersToArray = myArray
End Function
