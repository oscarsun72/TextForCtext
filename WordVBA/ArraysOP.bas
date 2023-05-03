Attribute VB_Name = "ArraysOP"
Option Explicit

Rem 20230328 Adrenaline �G
Rem �~�r�|�ӳ����A���e�Ƨ�
Public Sub SortArray_QuickSort(arrayToSort As Variant) 'VBA�޼ƹw�]���ǧ}�]pass by reference�^
'    Dim arrayToSort As Variant
'    Dim i As Integer
'
'    ' ���o�}�C
'    arrayToSort = Application.Transpose(ExistedNumColumnRange.value)
'
    ' �� QuickSort �Ƨ�
    'Call QuickSortArray(arrayToSort, 1, UBound(arrayToSort))
    Call QuickSortArray(arrayToSort, LBound(arrayToSort), UBound(arrayToSort))
    
'    ' ��X�Ƨǫ᪺���G
'    Debug.Print "�Ƨǫ᪺���G�G"
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

Rem creedit with chatGPT�j����
Rem �~�r�|�ӵ��e�A�����Ƨ�
Sub SortStringArray(ByRef arr() As String)
'arr = Array("apple", "banana", "cherry")
QuickSort arr, LBound(arr), UBound(arr) 'chatGPT�j���ġGLBound �O�@�� VBA ��ơA���|�^�ǰ}�C���U�ɡ]Lower Bound�^�A�]�N�O�}�C���Ĥ@�Ӥ��������ޡC�b�j�h�Ʊ��p�U�A�}�C���U�ɬO 0�A�����ɤ]�i�H�w�q����L�Ʀr�C�Ҧp�A�p�G�w�q�F�@�ӯ��ެ� 1 �� 10 ���}�C�A���� LBound ���ȴN�O 1�C
'For Each s In arr
'Debug.Print s
'Next s
End Sub

Private Sub QuickSort(ByRef arr() As String, ByVal l As Long, ByVal r As Long) 'l=left,r=right chatGPT�j���ġG�O���A�b�o�� QuickSort ��Ƥ��A l �ѼƥN���䪺���ަ�m�A�� r �ѼƥN��k�䪺���ަ�m�C�o�ǰѼƬO�ֳt�ƧǺ�k���D�n�����A�Ω���w�ƧǪ���ɡC�b�o�Ө�Ƥ��A arr �ƲլO�n�i��ƧǪ��ƲաA l �M r ���w�F�n�i��ƧǪ��Ʋժ��϶��C
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

Rem Bing�j����'https://www.notion.so/Characters-76ccb4ff823e4a82b0d0af042e5a650b?pvs=4#d7f45c8d4863487db4d92e4cb7787525
'�p�G�u�O�d�~�r����ƧǡA�h hanOnly=true
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
