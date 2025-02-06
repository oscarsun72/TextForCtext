VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Document_Close()
    Keywords.ClearDicts_YiKeywords
End Sub

Private Sub Document_New()
    Docs.Register_Event_Handler

End Sub

Private Sub Document_Open()
'    Stop
    Docs.Register_Event_Handler
End Sub
