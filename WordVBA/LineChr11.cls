VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LineChr11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Rem 與手動分行有關的所有物件與操作 20250202 大年初五
Dim endPos As Long, line As Range, rng As Range

Rem 取得範圍所在位置的行
Property Get LineRange(rng As Range) As Range
    If line Is Nothing Then
        Create rng
    ElseIf rng.start < line.start Or rng.End > line.End Then
        Create rng
    End If
    Set LineRange = line
End Property

Property Get EndPosition(rng As Range) As Long
    If line Is Nothing Then
        Set line = LineRange(rng)
    ElseIf rng.start < line.start Or rng.End > line.End Then
        Create rng
    End If
    
    EndPosition = line.End
End Property


Private Sub Create(rng As Range)
    Dim rngLine As Range, d As Document ', rngLineSt As Long, rngLineEd As Long
    Set d = rng.Document
    Set rngLine = rng.Document.Range(rng.start, rng.End)
    '指定範圍的前面有手動分行符號時
    If VBA.InStr(d.Range(1, rng.start).text, VBA.Chr(11)) Then
        rngLine.MoveStartUntil VBA.Chr(11), -rng.start
        '後面有分行符號時
        If VBA.InStr(d.Range(rng.End, d.Range.End).text, VBA.Chr(11)) Then
            rngLine.MoveEndUntil VBA.Chr(11), d.Range.End - rng.End
        Else '後面沒有分行符號時
            rngLine.End = d.Range.End - 1
        End If
        
        
    Else '指定範圍的前面沒有手動分行符號時
        rngLine.start = 0
        If VBA.InStr(d.Range(rng.End, d.Range.End), VBA.Chr(11)) Then '後面有分行符號
            rngLine.MoveEndUntil VBA.Chr(11), d.Range.End - rng.End
        Else
            Set line = Nothing
        End If
    End If
    Set line = rngLine
End Sub
