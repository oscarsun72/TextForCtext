Attribute VB_Name = "HTML2Doc"
Option Explicit
Enum TagNameHTML
    Sup = 0
    Subscript = 1
End Enum
Rem 與HTML文本轉成Word文件的操作今集中在此，先作練習用!!! 20241012 creedit_with_Copilot大菩薩：
'https://forkful.ai/zh/vba/html-and-the-web/parsing-html/
'https://www.msofficeforums.com/word-vba/48539-how-i-covert-html-documents-word-using.html
'https://www.youtube.com/watch?v=bcjKYdJa7nI&ab_channel=VBAbyMBA
Sub ConvertHtmlToWord() '20241012 creedit_with_Copilot大菩薩：
    Dim objWordApp As New word.Application
    Dim objWordDoc As word.Document
    Dim strFile As String
    Dim strFolder As String

    ' 設定 HTML 文件所在的文件夾
    strFolder = "C:\path\to\your\html\folder\"
    strFile = Dir(strFolder & "*.html")

    ' 開啟 Word 應用程序
    With objWordApp
        ' 開啟 HTML 文件
        Set objWordDoc = .Documents.Open(fileName:=strFolder & strFile, ConfirmConversions:=False)
        ' 將 HTML 內容儲存為 Word 文件
        objWordDoc.SaveAs2 fileName:=strFolder & Replace(strFile, ".html", ".docx"), FileFormat:=wdFormatDocumentDefault
        ' 關閉文件
        objWordDoc.Close
        ' 關閉 Word 應用程序
        .Quit
    End With
End Sub
'https://forkful.ai/zh/vba/html-and-the-web/parsing-html/
Sub ParseHTML()
    Dim htmlDoc As MSHTML.HTMLDocument
    Dim htmlElement As MSHTML.IHTMLElement
    Dim htmlElements As MSHTML.IHTMLElementCollection
    Dim htmlFile As String
    Dim fileContent As String
    
    ' 從文件加載 HTML 內容
    htmlFile = "C:\path\to\your\file.html"
    Open htmlFile For Input As #1
    fileContent = Input$(LOF(1), 1)
    Close #1
    
    ' 初始化 HTML 文檔
    Set htmlDoc = New MSHTML.HTMLDocument
    htmlDoc.body.innerHTML = fileContent
    
    ' 獲取所有錨標籤
    Set htmlElements = htmlDoc.getElementsByTagName("a")

    ' 循環遍歷所有錨元素並打印 href 屬性
    For Each htmlElement In htmlElements
        Debug.Print htmlElement.GetAttribute("href")
    Next htmlElement
End Sub


Rem 上標格式 20241012 creedit_with_Copilot大菩薩：
Sub ConvertHTMLSupToWordSup(rng As Range)
    ConvertHTMLTagToWord rng, Sup
End Sub
Rem 下標格式
Sub ConvertHTMLSubToWordSub(rng As Range)
    ConvertHTMLTagToWord rng, Subscript
End Sub
Private Sub ConvertHTMLTagToWord(rng As Range, tagname As TagNameHTML)
    ' 查找所有標籤
    Dim tag As String
    rng.Find.ClearFormatting
    Select Case tagname
        Case TagNameHTML.Sup
            tag = "sup"
'            rng.Find.Replacement.font.Superscript = True ' 設定文字為上標格式
        Case TagNameHTML.Subscript
            tag = "sub"
'            rng.Find.Replacement.font.Subscript = True ' 設定文字為下標格式
    End Select
    
    With rng.Find
        .text = "\<" & tag & "\>(*)\</" & tag & "\>"
'        With .Replacement'這些都沒用
''            .text = "^&"
'        End With
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = True
        .Execute 'Replace:=wdReplaceAll
        Do While rng.Find.Found
            Select Case tagname
                Case TagNameHTML.Sup
                    rng.font.Superscript = True ' 設定文字為上標格式
                Case TagNameHTML.Subscript
                    rng.font.Subscript = True ' 設定文字為下標格式
                Case Else
                
            End Select
            rng.Collapse Direction:=wdCollapseEnd
            rng.Find.Execute
        Loop
        
        With rng.Document.Range.Find
            .ClearFormatting
            .Execute findtext:="<" & tag & ">", Replace:=wdReplaceAll, replaceWith:=vbNullString
            .Execute findtext:="</" & tag & ">", Replace:=wdReplaceAll, replaceWith:=vbNullString
            .MatchWildcards = False
        End With
            
    End With
End Sub
