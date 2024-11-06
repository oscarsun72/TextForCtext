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

Public Sub Create(str As String)
    Set clnString = ¤å¦r³B²z.CharactersStr(str)
End Sub

Public Property Get LengthInTextElements() As Long
    LengthInTextElements = clnString.Count
End Property
