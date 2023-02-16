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
Rem creedit with chatGPT大菩薩
'Rem chatGPT大菩薩：定義一個類別來表示一個鍵值對：
'Public Key As String
'Public Value As Variant
Rem chatGPT大菩薩： 在 VBA 中，您可以使用下面的代碼定義 KeyValuePair 型別：
Private Type KeyValuePair
    key As Variant
    value As Variant
End Type



Rem chatGPT大菩薩：定義一個類別來表示字典物件:
Private dict As Collection

Private Sub Class_Initialize()
    Set dict = New Collection
End Sub
Rem chatGPT大菩薩：提供一個方法來添加鍵值對：
Public Sub Add(key As String, value As Variant)
    Dim pair As KeyValuePair
    'Set pair = New KeyValuePair
    pair.key = key
    pair.value = value
    dict.Add pair.value, key
End Sub
Rem chatGPT大菩薩：提供一個方法來獲取字典中的一個值：
Public Function GetValue(key As String) As Variant
    Dim pair As KeyValuePair
    On Error Resume Next
    pair.value = dict(key)
    GetValue = pair.value
End Function

Rem chatGPT大菩薩：提供一個方法來判斷字典中是否存在一個鍵:
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

