Attribute VB_Name = "db"
Option Explicit

Sub �~�y�j����[�`��()
'Alt+6
Dim db As New dBase, rng As word.Range
Set rng = Selection.Range
db.�~�y�j����[�`�� VBA.Replace(VBA.Replace(rng.Text, Chr(13) & Chr(7), ""), Chr(13), "")
rng.MoveStartUntil Chr(9)
rng.Select
End Sub

