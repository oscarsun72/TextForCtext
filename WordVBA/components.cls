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
                    '�t�m����}�C�H�ƥ[�Jcolleciton �e��
                    getComponentsArray components, c
                    componentsCl.Add components
End Sub


Sub getComponentsArray(ByRef components() As String, c As Range)
Dim a, i As Byte
        ReDim components(c.Characters.Count - 2)
        '�p�G�S���Ϥ�
        If c.InlineShapes.Count = 0 Then
            For Each a In c.Characters
                ' '�ư��x�s��r��
                If InStr(Chr(13) & Chr(7), a) = 0 Then
                    components(i) = a.Text
                    i = i + 1
                End If
            Next
        '�p�G���Ϥ�
        Else
            For Each a In c.Characters
                '�ư��x�s��r��
                If InStr(Chr(13) & Chr(7), a) = 0 Then
                    If a.InlineShapes.Count = 0 Then '�D�Ϥ�
                        components(i) = a.Text
                    Else '�Ϥ�
                        components(i) = a.InlineShapes(1).AlternativeText
                    End If
                    i = i + 1
                End If
            Next a
        End If
                    
End Sub

Rem creedit with chatGPT�j����
Function JoinDictionaryValues(dict As Scripting.Dictionary) As String
    Dim result As String
    Dim key As Variant
    For Each key In dict.Keys
        If InStr(dict(key), ",") > 0 Then '�H�r���������������j��
            result = result & key & VBA.vbTab & dict(key) & VBA.vbNewLine
        End If
    Next key
    If Len(result) > 0 Then
        JoinDictionaryValues = VBA.Replace(Left(result, Len(result) - 1), ",", "") '�M���@�����������j�Ū��r��
    End If
End Function

Rem creedit with chatGPT�j����
Sub SortStringArray(ByRef arr() As String)
'arr = Array("apple", "banana", "cherry")
QuickSort arr, LBound(arr), UBound(arr) 'chatGPT�j���ġGLBound �O�@�� VBA ��ơA���|�^�ǰ}�C���U�ɡ]Lower Bound�^�A�]�N�O�}�C���Ĥ@�Ӥ��������ޡC�b�j�h�Ʊ��p�U�A�}�C���U�ɬO 0�A�����ɤ]�i�H�w�q����L�Ʀr�C�Ҧp�A�p�G�w�q�F�@�ӯ��ެ� 1 �� 10 ���}�C�A���� LBound ���ȴN�O 1�C
'For Each s In arr
'Debug.Print s
'Next s
End Sub

Private Sub QuickSort(ByRef arr() As String, ByVal l As Long, ByVal r As Long) 'l=left,r=right chatGPT�j���ġG�O���A�b�o�� QuickSort ��Ƥ��A l �ѼƥN�����䪺���ަ�m�A�� r �ѼƥN���k�䪺���ަ�m�C�o�ǰѼƬO�ֳt�ƧǺ�k���D�n�����A�Ω���w�ƧǪ���ɡC�b�o�Ө�Ƥ��A arr �ƲլO�n�i��ƧǪ��ƲաA l �M r ���w�F�n�i��ƧǪ��Ʋժ��϶��C
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

Rem creedit with chatGPT�j����
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
