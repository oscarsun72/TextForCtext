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
Dim dict1 As New Scripting.Dictionary, dict2 As New Scripting.Dictionary, s1 As Variant, s2 As Variant, similarTextCopy As String, similarWithTextCopy As String
Private Function Initialize(ByVal text1, ByVal text2) As similartext
If s1 = text1 And s2 = text2 Then Exit Function
Dim w, str1, str2
If s1 <> text1 Then
    If dict1.Count > 0 Then dict1.RemoveAll
    s1 = text1
    If VBA.VarType(s1) = vbString Then
        str1 = ��r�B�z.SplitWithoutDelimiter_StringToStringArray((s1)) '�j��H�ǭȤ覡�ǻ��޼� https://learn.microsoft.com/zh-tw/dotnet/visual-basic/programming-guide/language-features/procedures/how-to-force-an-argument-to-be-passed-by-value
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
        str2 = ��r�B�z.SplitWithoutDelimiter_StringToStringArray((s2))
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
'�P�_�奻���O�_�ۦ�
Function Similarity(ByVal text1, ByVal text2) As Boolean 'text1 �Q���� text2 ���Ӥ���
If IsEmpty(text2) Or IsEmpty(text1) Then Exit Function
similarWithTextCopy = text2
Initialize text1, text2
If Me.CompareDictionaries(dict1, dict2) Then
    Similarity = True
ElseIf Me.CompareDictionarySimilarity(dict1, dict2) > 0.8 Then
    Similarity = True
ElseIf cosine_Similarity() > 0.8 Then
    Similarity = True
ElseIf jaccard_Similarity() > 0.8 Then
    Similarity = True
End If
If Similarity Then
    similarTextCopy = text1
Else
    similarTextCopy = ""
End If
End Function

'�^�Ǥ奻���ۦ����ʤ���
Function SimilarityPercent(ByVal text1, ByVal text2) As Variant 'text1 �Q���� text2 ���Ӥ���
If IsEmpty(text2) Or IsEmpty(text1) Then Exit Function
similarWithTextCopy = text2
Dim sim As Double
Initialize text1, text2
If Me.CompareDictionaries(dict1, dict2) Then
    sim = 1
Else
    sim = (Me.CompareDictionarySimilarity(dict1, dict2) + cosine_Similarity() + jaccard_Similarity()) / 3
End If
Rem chatGPT�j���ġG�b SimilarityPercent ��Ƥ��A���ϥ� Round ��ƪ��ĤG�ӰѼơA�H���w�p���I��O�d����ơC�Ҧp�A�n�O�d�@��p�ơA�i�H�ϥ� VBA.Round(sim * 100, 1)�C
SimilarityPercent = VBA.Round(sim * 100, 2)
End Function
Rem 20230222 creedit with chatGPT�j���ġG��� Dictionary ��H�G
Rem �b VBA ���A������ Dictionary ��H���ۦ��שΧ����ۦP�A�ݭn���O������̪���M�ȡC�H�U�O�@�Ǥ����k�G
Rem 1.  ������ Dictionary �O�_�����ۦP:
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
    Rem �W�z�N�X�|�ˬd��� Dictionary ����ȹ�ƶq�O�_�ۦP�A�H�Υ��̪���M�ȬO�_�����ǰt�C�p�G��� Dictionary �����ۦP�A�h��^ True�A�_�h��^ False�C
End Function

Rem 2.  ������ Dictionary ���ۦ��סG
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
        If dict2.Exists(key) Then '�Y dict���S�������key�A�h�b���� dict(key)�ɡA�|�۰ʲK�J����ȹ蠟�����A���dict����
            If dict1(key) = dict2(key) Then matchedKeys = matchedKeys + 1
        End If
    Next
    
    CompareDictionarySimilarity = matchedKeys / totalKeys
    Rem �W�z�N�X�|�p���� Dictionary ���ۦ��סC���|�ˬd��� Dictionary �����h����ȹ�O�ۦP���A�M��p�⥦�̪���ҡC�p�G��� Dictionary �����ۦP�A�h��^ 1�C�p�G��� Dictionary �S������ۦP����ȹ�A�h��^ 0�C
End Function


Rem �����Ӥ奻���ۦ��ץi�H�ϥξl���ۦ��שΪ�Jaccard�ۦ��׵���k�C
Rem "�l���ۦ���"�M"Jaccard�ۦ���"�O�`�Ϊ��奻�ۦ��׺�k�A�Ω�����Ӥ奻�������ۦ��{�סC
Rem "�l���ۦ���"�O�q�L�p���Ӥ奻�V�q�������l���ȨӿŶq���̤������ۦ��סC�V�q�����l���ȶV�j�A���ܨ�ӦV�q�V����A�ۦ��פ]�N�V���C�����@�W�A�i�H���N��Ӥ奻���ܦ����W�V�q�A�A�p���ӦV�q�������l���ȡC�ۦ��׭Ȫ��d��b0��1�����A�ȶV�j���ܬۦ��׶V���C
Rem "Jaccard�ۦ���"�O�q�L�p���Ӥ奻���涰�P�ö���������ҨӿŶq���̤������ۦ��סC�ۦ��׭Ȫ��d��]�b0��1�����A�ȶV�j���ܬۦ��׶V���C�����@�W�A�i�H���N��Ӥ奻���ܦ������X�A�M��p���Ӷ��X���涰�P�ö��A�̫�Υ涰���H�ö��ӱo��ۦ��׭ȡC
Rem �H�U�OVBA�{���X�ܨ� , �Ω�p���Ӥ奻���l���ۦ��שMJaccard�ۦ���:
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


Rem 20230222 chatGPT�j���ġG
Rem �o�ӹ�@�ݰ_�ӫܤ����I�H�U�O�ګ�ĳ���@�Ǥp��i�G
Rem �b Initialize ��ƪ��}�Y�A���N�ǤJ�� text1 �M text2 �ഫ���p�g�A�H�K�b�������ɩ����j�p�g�C
Rem �b SimilarityPercent ��Ƥ��A���N�ܼ� sim ���ƾ������n���� Double�A�H�K��ǽT�a�p��ۦ��סC
Rem �b SimilarityPercent ��Ƥ��A���ϥ� Round ��ƪ��ĤG�ӰѼơA�H���w�p���I��O�d����ơC�Ҧp�A�n�O�d�@��p�ơA�i�H�ϥ� VBA.Round(sim * 100, 1)�C
Rem �o�ǧ�i�u�O�@�Ǥp�Ӹ` , �A����@�w�g�D�`�n�F!