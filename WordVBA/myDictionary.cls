VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "myDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Rem creedit with chatGPT�j����
'Rem chatGPT�j���ġG�w�q�@�����O�Ӫ��ܤ@����ȹ�G
'Public Key As String
'Public Value As Variant
Rem chatGPT�j���ġG �b VBA ���A�z�i�H�ϥΤU�����N�X�w�q KeyValuePair ���O�G
Private Type KeyValuePair
    key As Variant
    value As Variant
End Type



Rem chatGPT�j���ġG�w�q�@�����O�Ӫ��ܦr�媫��:
Private dict As Collection

Private Sub Class_Initialize()
    Set dict = New Collection
End Sub
Rem chatGPT�j���ġG���Ѥ@�Ӥ�k�ӲK�[��ȹ�G
Public Sub Add(key As String, value As Variant)
    Dim pair As KeyValuePair
    'Set pair = New KeyValuePair
    pair.key = key
    pair.value = value
    dict.Add pair.value, key
End Sub
Rem chatGPT�j���ġG���Ѥ@�Ӥ�k������r�夤���@�ӭȡG
Public Function GetValue(key As String) As Variant
    Dim pair As KeyValuePair
    On Error Resume Next
    pair.value = dict(key)
    GetValue = pair.value
End Function

Rem chatGPT�j���ġG���Ѥ@�Ӥ�k�ӧP�_�r�夤�O�_�s�b�@����:
Public Function ContainsKey(key As String) As Boolean
    On Error Resume Next
    Dim value As String
    value = dict(key)
    If Err.Number = 0 Then
        ContainsKey = True
    Else
        ContainsKey = False
    End If
End Function
