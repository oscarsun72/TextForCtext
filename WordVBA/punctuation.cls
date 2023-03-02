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
    punctuationStr = 文字處理.PunctuationString
End Sub

Rem 20230221 chatGPT大菩薩：忽略標點符號的搜尋：在 VBA 中，可以使用 Dictionary 或 Collection 這兩個物件來實現映射表的功能，不一定非要用陣列。
Rem 以 Dictionary 為例，你可以將標點符號的位置作為 key，對應的標點符號作為 value，存入 Dictionary 中，這樣就可以實現在還原的過程中快速查找需要還原的標點符號了。
Rem 以下是一個簡單的範例程式碼，展示如何使用 Dictionary 來建立標點符號位置與對應符號的映射表：
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
    
    ' 將建立的映射表輸出到 Immediate Window
    Dim k As Variant
    For Each k In map.keys
        Debug.Print k, map(k)
    Next k
    Rem 在這個範例中，我們使用了 Scripting.Dictionary 這個物件來創建映射表。由於 VBA 預設沒有引用這個物件，所以需要在程式碼的頂部引入相關的參照（Reference），在 VBE 中可選擇「工具（Tools）」-「參照（References）」來新增對 Scripting Runtime ?的引用。
    Rem 在這個程式碼中，我們先定義一個文本字符串，然後使用一個迴圈遍歷該字符串中的每個字符。如果字符是標點符號，則將其位置與字符本身作為 key 和 value 存入 Dictionary 中。
    Rem 最後，我們使用一個迴圈遍歷 Dictionary 的所有 key，並將 key 和對應的 value 輸出到 VBE 的 Immediate Window 中。
End Sub

Function CreateContextPunctuationCollection(x As String) As Collection
Dim i As Long, xLen As Long, dict As Scripting.Dictionary, a As String, text As String, contextCollection As New Collection
Set dict = PunctuationDictionary
xLen = VBA.Len(x)
For i = 1 To xLen
    a = Mid(x, i, 1)
    If dict.Exists(a) Or a = Chr(13) Then '以標點符號及分段為斷句處，加入集合中成為元素
        ReDim pare(1) As String '以元素為2個的陣列儲存原文訊息
        pare(0) = text '記文文字片段
        pare(1) = a '記下文字片段後的標點符號
        contextCollection.Add pare
        text = ""
    Else
        text = text + a
    End If
Next i
Set CreateContextPunctuationCollection = contextCollection
End Function

Rem 20230221 creedit with YouChat菩薩：
Rem vba 如何將以下字串轉成字串陣列或 collection 或 dictionary ? punctuationStr = "（。「」……
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

Rem 清除x文本中的標點符號
Sub clearPunctuations(ByRef x As String)
Dim punc, e
punc = PunctuationArray
For Each e In punc
    If InStr(x, e) Then x = VBA.Replace(x, e, "")
Next e
End Sub

Rem 忽略標點符號的找字串
Function inStrIgnorePunctuation(ByVal xDomain As String, ByVal chkClipboardText As String) As Long
'Dim x As String, chkx As String
'x = xDomain: chkx = chkClipboardText
Me.clearPunctuations chkClipboardText '強制以傳值方式傳遞引數 :https://learn.microsoft.com/zh-tw/dotnet/visual-basic/programming-guide/language-features/procedures/how-to-force-an-argument-to-be-passed-by-value
Me.clearPunctuations xDomain
inStrIgnorePunctuation = VBA.InStr(xDomain, chkClipboardText)
End Function

Sub restoreOriginalTextPunctuations(ByRef OriginalText As String, ByRef RemovedPunctuationsText As String)
Dim RemovedPunctuationsTextSubStr As String
RemovedPunctuationsTextSubStr = OriginalText
Me.clearPunctuations RemovedPunctuationsTextSubStr
'若原文有標點符號之文本 OriginalText 並不包含要還原其標點符號之文本 則不處理
If InStr(RemovedPunctuationsTextSubStr, RemovedPunctuationsText) = 0 Then Exit Sub

Dim contextCollection As Collection, iCntcontextCollection As Long, jCntcontextCollection As Long, cntcontextCollection As Long
Set contextCollection = CreateContextPunctuationCollection(OriginalText)
cntcontextCollection = contextCollection.Count
For iCntcontextCollection = 1 To cntcontextCollection
    If contextCollection.item(iCntcontextCollection)(0) <> "" And VBA.InStr(RemovedPunctuationsText, contextCollection.item(iCntcontextCollection)(0)) = 1 Then '找到符合要加入而重複的文字片段
        jCntcontextCollection = iCntcontextCollection + 1
        RemovedPunctuationsTextSubStr = Mid(RemovedPunctuationsText, Len(contextCollection.item(iCntcontextCollection)(0)) + 1)
        Do While VBA.InStr(RemovedPunctuationsTextSubStr, contextCollection.item(jCntcontextCollection)(0)) = 1 And RemovedPunctuationsTextSubStr <> ""
            RemovedPunctuationsTextSubStr = Mid(RemovedPunctuationsTextSubStr, Len(contextCollection.item(jCntcontextCollection)(0)) + 1)
            jCntcontextCollection = jCntcontextCollection + 1
            If jCntcontextCollection > contextCollection.Count Then Exit Do
        Loop
        If RemovedPunctuationsTextSubStr = "" Then '比對吻合，離開迴圈，取得含標點符號之原文片段
            Exit For
        ElseIf VBA.Replace(RemovedPunctuationsTextSubStr, Chr(13), "") = "" Then '只剩下分段符號
            Exit For
        End If
    End If
Next iCntcontextCollection
'取得含標點符號之原文片段
If RemovedPunctuationsTextSubStr = "" Or VBA.Replace(RemovedPunctuationsTextSubStr, Chr(13), "") = "" Then
    For iCntcontextCollection = iCntcontextCollection To jCntcontextCollection - 1
        RemovedPunctuationsTextSubStr = RemovedPunctuationsTextSubStr + contextCollection.item(iCntcontextCollection)(0) + contextCollection.item(iCntcontextCollection)(1)
    Next iCntcontextCollection
    RemovedPunctuationsText = RemovedPunctuationsTextSubStr
End If
End Sub

Private Sub Class_Terminate()
punctuationStr = Empty
punctuation_Array = Empty
Set punctuation_Collection = Nothing
Set punctuation_Dictionary = Nothing
End Sub
