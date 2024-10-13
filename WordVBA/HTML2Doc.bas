Attribute VB_Name = "HTML2Doc"
Option Explicit
Enum TagNameHTML
    Sup = 0
    Subscript = 1
End Enum
Rem �PHTML�奻�নWord��󪺾ާ@�������b���A���@�m�ߥ�!!! 20241012 creedit_with_Copilot�j���ġG
'https://forkful.ai/zh/vba/html-and-the-web/parsing-html/
'https://www.msofficeforums.com/word-vba/48539-how-i-covert-html-documents-word-using.html
'https://www.youtube.com/watch?v=bcjKYdJa7nI&ab_channel=VBAbyMBA
Sub ConvertHtmlToWord() '20241012 creedit_with_Copilot�j���ġG
    Dim objWordApp As New word.Application
    Dim objWordDoc As word.Document
    Dim strFile As String
    Dim strFolder As String

    ' �]�w HTML ���Ҧb�����
    strFolder = "C:\path\to\your\html\folder\"
    strFile = Dir(strFolder & "*.html")

    ' �}�� Word ���ε{��
    With objWordApp
        ' �}�� HTML ���
        Set objWordDoc = .Documents.Open(fileName:=strFolder & strFile, ConfirmConversions:=False)
        ' �N HTML ���e�x�s�� Word ���
        objWordDoc.SaveAs2 fileName:=strFolder & Replace(strFile, ".html", ".docx"), FileFormat:=wdFormatDocumentDefault
        ' �������
        objWordDoc.Close
        ' ���� Word ���ε{��
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
    
    ' �q���[�� HTML ���e
    htmlFile = "C:\path\to\your\file.html"
    Open htmlFile For Input As #1
    fileContent = Input$(LOF(1), 1)
    Close #1
    
    ' ��l�� HTML ����
    Set htmlDoc = New MSHTML.HTMLDocument
    htmlDoc.body.innerHTML = fileContent
    
    ' ����Ҧ������
    Set htmlElements = htmlDoc.getElementsByTagName("a")

    ' �`���M���Ҧ��㤸���å��L href �ݩ�
    For Each htmlElement In htmlElements
        Debug.Print htmlElement.GetAttribute("href")
    Next htmlElement
End Sub


Rem �W�Ю榡 20241012 creedit_with_Copilot�j���ġG
Sub ConvertHTMLSupToWordSup(rng As Range)
    ConvertHTMLTagToWord rng, Sup
End Sub
Rem �U�Ю榡
Sub ConvertHTMLSubToWordSub(rng As Range)
    ConvertHTMLTagToWord rng, Subscript
End Sub
Private Sub ConvertHTMLTagToWord(rng As Range, tagname As TagNameHTML)
    ' �d��Ҧ�����
    Dim tag As String
    rng.Find.ClearFormatting
    Select Case tagname
        Case TagNameHTML.Sup
            tag = "sup"
'            rng.Find.Replacement.font.Superscript = True ' �]�w��r���W�Ю榡
        Case TagNameHTML.Subscript
            tag = "sub"
'            rng.Find.Replacement.font.Subscript = True ' �]�w��r���U�Ю榡
    End Select
    
    With rng.Find
        .text = "\<" & tag & "\>(*)\</" & tag & "\>"
'        With .Replacement'�o�ǳ��S��
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
                    rng.font.Superscript = True ' �]�w��r���W�Ю榡
                Case TagNameHTML.Subscript
                    rng.font.Subscript = True ' �]�w��r���U�Ю榡
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
