VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "punctuation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim punctuationStr As String, punctuation_Array, punctuation_Collection As VBA.Collection, punctuation_Dictionary As Scripting.Dictionary

Private Sub Class_Initialize()
    punctuationStr = ��r�B�z.PunctuationString
End Sub

Rem 20230221 chatGPT�j���ġG�������I�Ÿ����j�M�G�b VBA ���A�i�H�ϥ� Dictionary �� Collection �o��Ӫ���ӹ�{�M�g�����\��A���@�w�D�n�ΰ}�C�C
Rem �H Dictionary ���ҡA�A�i�H�N���I�Ÿ�����m�@�� key�A���������I�Ÿ��@�� value�A�s�J Dictionary ���A�o�˴N�i�H��{�b�٭쪺�L�{���ֳt�d��ݭn�٭쪺���I�Ÿ��F�C
Rem �H�U�O�@��²�檺�d�ҵ{���X�A�i�ܦp��ϥ� Dictionary �ӫإ߼��I�Ÿ���m�P�����Ÿ����M�g���G
Sub CreatePunctuationMap()
    Dim text As String
    text = "This is a sample text, with some punctuations! Do you like it?"
    
    Dim map As New Scripting.Dictionary
    Dim i As Long
    For i = 1 To Len(text)
        Dim c As String
        c = Mid(text, i, 1)
        If c Like "[!A-Za-z0-9 ]" Then
            map.Add i, c
        End If
    Next i
    
    ' �N�إߪ��M�g����X�� Immediate Window
    Dim k As Variant
    For Each k In map.keys
        Debug.Print k, map(k)
    Next k
    Rem �b�o�ӽd�Ҥ��A�ڭ̨ϥΤF Scripting.Dictionary �o�Ӫ���ӳЫجM�g���C�ѩ� VBA �w�]�S���ޥγo�Ӫ���A�ҥH�ݭn�b�{���X�������ޤJ�������ѷӡ]Reference�^�A�b VBE ���i��ܡu�u��]Tools�^�v-�u�ѷӡ]References�^�v�ӷs�W�� Scripting Runtime ?���ޥΡC
    Rem �b�o�ӵ{���X���A�ڭ̥��w�q�@�Ӥ奻�r�Ŧ�A�M��ϥΤ@�Ӱj��M���Ӧr�Ŧꤤ���C�Ӧr�šC�p�G�r�ŬO���I�Ÿ��A�h�N���m�P�r�ť����@�� key �M value �s�J Dictionary ���C
    Rem �̫�A�ڭ̨ϥΤ@�Ӱj��M�� Dictionary ���Ҧ� key�A�ñN key �M������ value ��X�� VBE �� Immediate Window ���C
End Sub

Function CreateContextPunctuationCollection(x As String) As Collection
Dim i As Long, xLen As Long, dict As Scripting.Dictionary, a As String, text As String, contextCollection As New Collection
Set dict = PunctuationDictionary
xLen = VBA.Len(x)
For i = 1 To xLen
    a = Mid(x, i, 1)
    If dict.Exists(a) Or a = Chr(13) Then '�H���I�Ÿ��Τ��q���_�y�B�A�[�J���X����������
        ReDim pare(1) As String '�H������2�Ӫ��}�C�x�s���T��
        pare(0) = text '�O���r���q
        pare(1) = a '�O�U��r���q�᪺���I�Ÿ�
        contextCollection.Add pare
        text = ""
    Else
        text = text + a
    End If
Next i
Set CreateContextPunctuationCollection = contextCollection
End Function

Rem 20230221 creedit with YouChat���ġG
Rem vba �p��N�H�U�r���ন�r��}�C�� collection �� dictionary ? punctuationStr = "�]�C�u�v�K�K
Rem You can use the Split function in VBA to convert the string into an array or collection. For example:
Property Get PunctuationArray()
If VBA.IsEmpty(punctuation_Array) Then
    Dim i As Long, lenpunctuationStr As Long
    lenpunctuationStr = Len(punctuationStr)
    ReDim punc(lenpunctuationStr - 1) As String
    For i = 1 To lenpunctuationStr
         punc(i - 1) = Mid(punctuationStr, i, 1)
    Next i
    'PunctuationArray = VBA.Split(punctuationStr)
    punctuation_Array = punc
End If
PunctuationArray = punctuation_Array
End Property

Property Get PunctuationCollection() As Collection
    If punctuation_Collection Is Nothing Then
        Dim item As Variant
        Set PunctuationCollection = New Collection
        For Each item In PunctuationArray
            PunctuationCollection.Add item
        Next
        Set punctuation_Collection = PunctuationCollection
    End If
    Set PunctuationCollection = punctuation_Collection
End Property

Property Get PunctuationDictionary() As Scripting.Dictionary
    If punctuation_Dictionary Is Nothing Then
        Dim item As Variant
        Set PunctuationDictionary = New Scripting.Dictionary
        For Each item In PunctuationArray
            If Not PunctuationDictionary.Exists(item) Then
                PunctuationDictionary.Add item, item
            End If
        Next
        Set punctuation_Dictionary = PunctuationDictionary
    End If
    Set PunctuationDictionary = punctuation_Dictionary
End Property

Rem �M��x�奻�������I�Ÿ�
Sub clearPunctuations(ByRef x As String)
Dim punc, e
punc = PunctuationArray
For Each e In punc
    If InStr(x, e) Then x = VBA.Replace(x, e, "")
Next e
End Sub

Rem �������I�Ÿ�����r��
Function inStrIgnorePunctuation(ByVal xDomain As String, ByVal chkClipboardText As String) As Long
'Dim x As String, chkx As String
'x = xDomain: chkx = chkClipboardText
Me.clearPunctuations chkClipboardText '�j��H�ǭȤ覡�ǻ��޼� :https://learn.microsoft.com/zh-tw/dotnet/visual-basic/programming-guide/language-features/procedures/how-to-force-an-argument-to-be-passed-by-value
Me.clearPunctuations xDomain
inStrIgnorePunctuation = VBA.InStr(xDomain, chkClipboardText)
End Function

Sub restoreOriginalTextPunctuations(ByRef OriginalText As String, ByRef RemovedPunctuationsText As String)
Dim RemovedPunctuationsTextSubStr As String
RemovedPunctuationsTextSubStr = OriginalText
Me.clearPunctuations RemovedPunctuationsTextSubStr
'�Y��妳���I�Ÿ����奻 OriginalText �ä��]�t�n�٭����I�Ÿ����奻 �h���B�z
If InStr(RemovedPunctuationsTextSubStr, RemovedPunctuationsText) = 0 Then Exit Sub

Dim contextCollection As Collection, iCntcontextCollection As Long, jCntcontextCollection As Long, cntcontextCollection As Long
Set contextCollection = CreateContextPunctuationCollection(OriginalText)
cntcontextCollection = contextCollection.Count
For iCntcontextCollection = 1 To cntcontextCollection
    If contextCollection.item(iCntcontextCollection)(0) <> "" And VBA.InStr(RemovedPunctuationsText, contextCollection.item(iCntcontextCollection)(0)) = 1 Then '���ŦX�n�[�J�ӭ��ƪ���r���q
        jCntcontextCollection = iCntcontextCollection + 1
        RemovedPunctuationsTextSubStr = Mid(RemovedPunctuationsText, Len(contextCollection.item(iCntcontextCollection)(0)) + 1)
        Do While VBA.InStr(RemovedPunctuationsTextSubStr, contextCollection.item(jCntcontextCollection)(0)) = 1 And RemovedPunctuationsTextSubStr <> ""
            RemovedPunctuationsTextSubStr = Mid(RemovedPunctuationsTextSubStr, Len(contextCollection.item(jCntcontextCollection)(0)) + 1)
            jCntcontextCollection = jCntcontextCollection + 1
            If jCntcontextCollection > contextCollection.Count Then Exit Do
        Loop
        If RemovedPunctuationsTextSubStr = "" Then '���k�X�A���}�j��A���o�t���I�Ÿ��������q
            Exit For
        ElseIf VBA.Replace(RemovedPunctuationsTextSubStr, Chr(13), "") = "" Then '�u�ѤU���q�Ÿ�
            Exit For
        End If
    End If
Next iCntcontextCollection
'���o�t���I�Ÿ��������q
If RemovedPunctuationsTextSubStr = "" Or VBA.Replace(RemovedPunctuationsTextSubStr, Chr(13), "") = "" Then
    For iCntcontextCollection = iCntcontextCollection To jCntcontextCollection - 1
        RemovedPunctuationsTextSubStr = RemovedPunctuationsTextSubStr + contextCollection.item(iCntcontextCollection)(0) + contextCollection.item(iCntcontextCollection)(1)
    Next iCntcontextCollection
    RemovedPunctuationsText = RemovedPunctuationsTextSubStr
End If
End Sub