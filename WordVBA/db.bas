Attribute VB_Name = "db"
Option Explicit

Sub 漢語大詞典加注音()
'Alt+6
Dim db As New dBase, rng As word.Range
Set rng = Selection.Range
db.漢語大詞典加注音 VBA.Replace(VBA.Replace(rng.Text, Chr(13) & Chr(7), ""), Chr(13), "")
rng.MoveStartUntil Chr(9)
rng.Select
End Sub

