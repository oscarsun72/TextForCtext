VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private clnString As VBA.Collection

Public Function Create(str As String) As StringInfo
    Dim si As New StringInfo
    Set clnString = ¤å¦r³B²z.CharactersStr(str)
    Set si.ClnStringSet = clnString
    Set Create = si
End Function

Public Property Get LengthInTextElements() As Long
    LengthInTextElements = clnString.Count
End Property

Property Set ClnStringSet(c As VBA.Collection)
    Set clnString = c
End Property


