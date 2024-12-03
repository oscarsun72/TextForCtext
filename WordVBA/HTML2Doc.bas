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
Rem '�B�z�W�s�� 20241016
Private Sub InsertHTMLLinks(rngHtml As Range, Optional domainUrlPrefix As String)
    Dim e As Variant '�@���q�Τ@���ܼ�
    Dim obj As Object '�@���q�Ϊ����ܼ�
    Dim rng As Range, rngClose As Range, url As String
    Set rng = rngHtml.Document.Range(rngHtml.start, rngHtml.End)
    If VBA.InStr(rng.text, "</a>") = 0 Then
        playSound 12 'for check
        rng.Select
        Stop
    End If
    If VBA.InStr(rng.text, "<a ") Then 'And VBA.Len(rng.text) > VBA.Len("<a href=""></a>") Then
        With rng.Find
            .ClearFormatting
            .MatchWildcards = False
            Do While .Execute(findText:="<a ")
                rng.MoveEndUntil ">"
                rng.End = rng.End + 1
                url = rng.text: e = rng.text
                Set rngClose = rng.Document.Range(rng.End, rngHtml.End)
                
'                If InStr(rngHtml.text, "id") Then
'                    rng.Select
'                    Stop 'just for test
'                End If
                
                If Not rngClose.Find.Execute("</a>") Then
                    playSound 12 'for check
                    rng.Select
                    Stop
                End If
                url = getHTML_AttributeValue("href", url)
                'url = getHTML_AttributeValue("<a href", p.Range.text)
                e = getHTML_AttributeValue("title", VBA.CStr(e))
                Select Case VBA.Left(url, 1)
                    Case "#"
                            If Not SeleniumOP.IsWDInvalid() Then
                                url = WD.url & url
                            End If
                    Case "/"
                        url = domainUrlPrefix & url '���|���h�@�ӱ׽u�]/�^�]�O�i�H���A�S�t 20241012
                    Case vbNullString
                        If rng.Document.Range(rng.End, rngClose.start).InlineShapes.Count > 0 Then '�᭱�|�ˬd�Grng.Document.Range(rng.End, rngClose.start).ShapeRange.Count > 0
                            playSound 12
                            Stop 'for check
                        End If
                        '�Ū��W�s���A���B�z�A�����M��
                    Case "."
                        If VBA.Left(url, VBA.Len("../../")) = "../../" Then
                            url = domainUrlPrefix & VBA.Mid(url, VBA.Len("../../"))
                        Else
                            playSound 12
                            rng.Select
                            Debug.Print url
                            Stop 'check
                        End If
                    Case Else
                        If Not VBA.Left(url, 4) = "http" Then
                            playSound 12
                            rng.Select
                            Debug.Print url
                            Stop 'check
                            url = domainUrlPrefix & url
                        End If
                End Select
                If url <> vbNullString Then
                    Set obj = rng.Document.Range(rng.start, rngClose.End).ShapeRange
                    rng.text = vbNullString: rngClose.text = vbNullString
                    If Not obj Is Nothing Then
                        Select Case obj.Count
                            Case 0
                                If rng.Document.Range(rng.End, rngClose.start).text <> vbNullString Then
                                    rng.Document.Range(rng.End, rngClose.start).Hyperlinks.Add rng.Document.Range(rng.End, rngClose.start), url, , e
                                End If
                            Case 1
                                rng.Document.Range(rng.End, rngClose.start).Hyperlinks.Add obj(1), url, , e
                            Case Else
                                playSound 12 'for check
                                Stop
                        End Select
                        
                        Set obj = Nothing
                    Else
                        playSound 12 'for check
                        Stop
                    End If
                End If
                If rng.text <> vbNullString Or rngClose.text = vbNullString Then
                    rng.text = vbNullString: rngClose.text = vbNullString
                End If
                If rng.Paragraphs(1).Range.text = Chr(13) Then
                    rng.Paragraphs(1).Range.text = vbNullString
                End If
                rng.SetRange rngHtml.start, rngHtml.End
            Loop
        End With 'With rng.Find
    End If '�B�z�W�s��
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
    htmlDoc.body.innerHtml = fileContent
    
    ' ����Ҧ������
    Set htmlElements = htmlDoc.getElementsByTagName("a")

    ' �`���M���Ҧ��㤸���å��L href �ݩ�
    For Each htmlElement In htmlElements
        Debug.Print htmlElement.GetAttribute("href")
    Next htmlElement
End Sub

Rem �C��X�ഫ��RGB
Function ColorCodetoRGB(colorCode As String) As Long()
    ' �Nbgcolor�ഫ��RGB�C��
    'Dim r As Integer, g As Integer, b As Integer
    If VBA.InStr(colorCode, " ") Then colorCode = VBA.Trim(colorCode)
    If VBA.Left(colorCode, 1) <> "#" Then Exit Function
    Dim arr(2) As Long
    arr(0) = CLng("&H" & Mid(colorCode, 2, 2))
    arr(1) = CLng("&H" & Mid(colorCode, 4, 2))
    arr(2) = CLng("&H" & Mid(colorCode, 6, 2))
    ColorCodetoRGB = arr
End Function

Rem �C��X�ഫ��RGB
Function RGBFormColorCode(colorCode As String) As Long
    ' �Nbgcolor�ഫ��RGB�C��
    'Dim r As Integer, g As Integer, b As Integer
    If VBA.InStr(colorCode, " ") Then colorCode = VBA.Trim(colorCode)
    If VBA.Left(colorCode, 1) <> "#" Then Exit Function
    Dim arr(2) As Long
    arr(0) = CLng("&H" & Mid(colorCode, 2, 2))
    arr(1) = CLng("&H" & Mid(colorCode, 4, 2))
    arr(2) = CLng("&H" & Mid(colorCode, 6, 2))
    RGBFormColorCode = VBA.RGB(arr(0), arr(1), arr(2))
End Function

Rem �W�Ю榡 20241012 creedit_with_Copilot�j���ġG
Sub ConvertHTMLSupToWordSup(rng As Range)
    ConvertHTMLTagToWord rng, TagNameHTML.Sup
End Sub
Rem �U�Ю榡
Sub ConvertHTMLSubToWordSub(rng As Range)
    ConvertHTMLTagToWord rng, TagNameHTML.Subscript
End Sub
Private Sub ConvertHTMLTagToWord(rng As Range, ByVal tagname As TagNameHTML)
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
            .Execute findText:="<" & tag & ">", Replace:=wdReplaceAll, replaceWith:=vbNullString
            .Execute findText:="</" & tag & ">", Replace:=wdReplaceAll, replaceWith:=vbNullString
            .MatchWildcards = False
        End With
            
    End With
End Sub
'Rem 20241011 HTML ���B�z.Porc=Porcess
'Private Sub tablePorc_HTML2Word(rng As Range)
'    Dim rngClose
'    'Do While VBA.InStr(rngHtml.text, "<table")
'        With rng.Find
'            .ClearFormatting
'            .text = "<table "
'            .Execute
'            Set rngClose = rng.Document.Range(rng.End, rngHtml.End)
'            With rngClose.Find
'                .text = "</table>"
'                .Execute
'            End With
'            Set rng = rngHtml.Document.Range(rng.start, rngClose.End)
'            insertHTMLTable rng, domainUrlPrefix
'        End With
'    'Loop
'End Sub
Rem 20241011 HTML �L�ǲM�檺�B�z.Porc=Porcess
Private Sub unorderedListPorc_HTML2Word(rngHtml As Range)
    Rem �L�ǲM�檺�B�z
    Dim rngUnorderedList As Range, st As Long, ed As Long, rngUnorderedListSub As Range, p As Paragraph
    If VBA.InStr(rngHtml.text, "<ul") Then
        Do
            Set rngUnorderedList = getRangeFromULToUL_UnorderedListRange(rngHtml)
            If Not rngUnorderedList Is Nothing Then
                st = rngUnorderedList.start
                Set p = rngUnorderedList.Paragraphs(1).Previous
                If Not p Is Nothing Then
                    '�p�G�O���Ǻ����u���N�`���G�v
                    If VBA.InStr(p.Range.text, "���N�`���G") Then
                        With rngUnorderedList.Find
                            .Execute "<li>", , , , , , , , , vbNullString, wdReplaceAll
                            .Execute "</li>", , , , , , , , , vbNullString, wdReplaceAll
                            .Execute "</ul>", , , , , , , , , vbNullString, wdReplaceAll
                             ed = rngUnorderedList.End
                        End With
                        Set rngUnorderedListSub = rngUnorderedList.Document.Range(rngUnorderedList.start, rngUnorderedList.End)
                        rngUnorderedListSub.Find.ClearFormatting
                        If rngUnorderedListSub.Find.Execute("<ul ") Then
                            rngUnorderedListSub.MoveEndUntil ">"
                            rngUnorderedListSub.End = rngUnorderedListSub.End + 2
                            If rngUnorderedListSub.Characters(rngUnorderedListSub.Characters.Count) <> Chr(13) Then
                                rngUnorderedListSub.End = rngUnorderedListSub.End - 1
                            End If
                            rngUnorderedListSub.text = vbNullString
    
                        Else
                            rngUnorderedListSub.SetRange rngUnorderedList.start, rngUnorderedList.End
                            If rngUnorderedListSub.Find.Execute("<ul>") Then
                                If rngUnorderedListSub.Paragraphs(1).Range.text = rngUnorderedListSub & Chr(13) Then
                                    rngUnorderedListSub.Paragraphs(1).Range.text = vbNullString
                                Else
                                    rngUnorderedListSub.text = vbNullString
                                End If
                            End If
                        End If
                        If rngUnorderedList.Characters(rngUnorderedList.Characters.Count) = Chr(13) Then
                            rngUnorderedList.End = rngUnorderedList.End - 1
                        End If
                        With rngUnorderedList
                            '.Hyperlinks.Add rngLink, iwe.GetAttribute("href")'�b�e���w�g���J�W�s���F
                            .Style = wdStyleHeading2 '���D 2
                            .font.Size = 18
                            .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle '��涡�Z
                        End With
                    Else
                        GoTo UnorderedListRange
                    End If
                Else
UnorderedListRange:
                    
                    rngUnorderedList.Select 'for chect
                    'Stop
                    
                    'Set rngUnorderedList = Nothing
                    'InsertHTMLList rngUnorderedList.text
                    
                    'If (VBA.Left(rngUnorderedList, 5) = "<ul>" & Chr(13) Or VBA.Left(rngUnorderedList, 4) = "<ul ") And VBA.Right(rngUnorderedList, 6) = Chr(13) & "</ul>" Then
                    'If (VBA.Left(rngUnorderedList, 4) = "<ul " Or VBA.Left(rngUnorderedList, 5) = "<ul>" & Chr(13) Or VBA.Left(rngUnorderedList, 4) = "<ul ") And VBA.Right(rngUnorderedList, 5) = "</ul>" Then
                    If (VBA.Left(rngUnorderedList, 4) = "<ul " Or VBA.Left(rngUnorderedList, 5) = "<ul>" & Chr(13) Or VBA.Left(rngUnorderedList, 4) = "<ul ") And (VBA.Right(rngUnorderedList, 6) = " </ul>" Or VBA.Right(rngUnorderedList, 6) = Chr(13) & "</ul>") Then
                        With rngUnorderedList
                            With .Find
                                .Execute "<li>", , , , , , , , , vbNullString, wdReplaceAll
                                .Execute "</li>", , , , , , , , , vbNullString, wdReplaceAll
                                If VBA.InStr(rngUnorderedList.text, "<ul>" & Chr(13)) Then .Execute "<ul>^p", , , , , , , , , vbNullString, wdReplaceAll
                                .Execute "^p</ul>", , , , , , , , , vbNullString, wdReplaceAll
                            End With
                            If VBA.Left(rngUnorderedList, 4) = "<ul " Then
                                Dim rngClear As Range
                                Set rngClear = rngUnorderedList.Document.Range(rngUnorderedList.start, rngUnorderedList.End)
                                rngClear.Find.text = "<ul "
                                rngClear.Find.Execute
                                rngClear.MoveEndUntil ">"
                                rngClear.End = rngClear.End + 1
                                If rngClear.Paragraphs(1).Range = rngClear.text & Chr(13) Then
                                    rngClear.Paragraphs(1).Range.text = vbNullString
                                Else
                                    rngClear.text = vbNullString
                                End If
                            End If
                            .ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
                                ListGalleries(wdBulletGallery).ListTemplates(1), ContinuePreviousList:= _
                                False, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
                                wdWord10ListBehavior
                            
                        End With
                    Else
                        Stop 'for chect
                    
                    End If
                End If
            End If
        Loop Until rngUnorderedList Is Nothing
    End If 'If VBA.InStr(rngHtml.text, "<ul") Then
End Sub
Rem �P�_�O�_�O��¶�� creedit_with_Copilot�j���� 20241016
Function IsWrapTextImage(imgTag As String) As Boolean
    Dim startPos As Long, endPos As Long, tagContent As String
    
    ' ��� <img ���_�l��m
    startPos = InStr(imgTag, "<img")
    If startPos = 0 Then
        IsWrapTextImage = False
        Exit Function
    End If
    
    ' ��� > ��������m
    endPos = InStr(startPos, imgTag, ">")
    If endPos = 0 Then endPos = VBA.Len(imgTag)
'    If endPos = 0 Then
'        IsWrapTextImage = False
'        Exit Function
'    End If
    
    ' ���� <img> ���Ҥ��e
    tagContent = Mid(imgTag, startPos, endPos - startPos + 1)
    
    ' �ˬd���Ҥ��e�O�_�]�t��¶�Ϫ� CSS �˦�
    If InStr(tagContent, "float:") > 0 Or InStr(tagContent, "class=") > 0 Then '_
'        Or (InStr(tagContent, "vertical-align:") > 0 _
'            And InStr(tagContent, "vertical-align: bottom") = 0 _
'            And VBA.InStr(tagContent, "vertical-align: baseline") = 0 _
'            And InStr(tagContent, "vertical-align:bottom") = 0 _
'            And VBA.InStr(tagContent, "vertical-align:baseline") = 0) Then

'                   vertical-align���������ӬO�Ginlsp.Range.ParagraphFormat.BaseLineAlignment �ݩʡI20241016
    
        IsWrapTextImage = True
    Else
        IsWrapTextImage = False
    End If
End Function

Rem �NHTML�奻�m�����Ϥ��A���\�h�Ǧ^�@�Ӧ��ĤF InlineShape���� 20241011 textPart:�n�ѪR��HTML�奻�Arng�G�n���J�Ϥ�����m�FdomainUrlPrefix �O�_�Ϥ����}�n�[��W�e��
Private Function insertHTMLImage(textPart As String, rng As Range, Optional domainUrlPrefix As String) As word.InlineShape
    Dim url As String, w As Single, h As Single, align As String, hspace As String
    Dim arr, e, attr As String, l, attrSetting
    'url = getImageUrl(textPart)
    url = VBA.Replace(getHTML_AttributeValue("src", textPart), "../..", vbNullString)
    If VBA.InStr(textPart, "width") Then
        w = VBA.CSng(getHTML_AttributeValue("width", textPart))
    End If
    If VBA.InStr(textPart, "height") Then
        h = VBA.CSng(getHTML_AttributeValue("height", textPart))
    End If
    If VBA.InStr(textPart, "align") Then
        align = getHTML_AttributeValue("align", textPart)
    End If
    If VBA.InStr(textPart, "hspace") Then
        hspace = getHTML_AttributeValue("hspace", textPart)
    End If
    
    
    If VBA.InStr(url, "http") <> 1 Then
'        If domainUrlPrefix = vbNullString Then
'            'msgbox "���a�J����e��~��"
'            'If domainUrlPrefix = vbNullString Then domainUrlPrefix = "https://www.eee-learning.com"
'
'            'If Not SeleniumOP.IsWDInvalid() Then
'                'domainUrlPrefix = getDomainUrlPrefix(SeleniumOP.WD.url)
'            'End If
'
'        End If
        If Not IsBase64Image(url) Then 'base64�s�X���Ϥ�
            url = domainUrlPrefix & url '���|���h�@�ӱ׽u�]/�^�]�O�i�H���A�S�t 20241012
        Else
            If Base64ToImage(url, VBA.Environ("TEMP") & "\" & "tempImage.png") = False Then
                Stop
'                GoTo finish
                Exit Function
            End If
        End If
    End If
    Dim inlsp As InlineShape
    
    If Not IsBase64Image(url) Then 'VBA.InStr(url, "data:image/png;base64") = 0 Then
            'rng.InlineShapes.AddPicture fileName:=url, _
                            LinkToFile:=False, SaveWithDocument:=True
        If w > 0 Or h > 0 Then
            On Error Resume Next
            Set inlsp = rng.InlineShapes.AddPicture(fileName:=url, _
                            LinkToFile:=False, SaveWithDocument:=True, Range:=rng)
            On Error GoTo 0
            If Not inlsp Is Nothing Then
                If inlsp.Range.tables.Count > 0 Then
                Else
'                    If Not IsWrapTextImage(textPart) Then
                        If w = 0 Or h = 0 Then
                            inlsp.LockAspectRatio = msoTrue
                        End If
                        resizePicture rng, inlsp, url, w, h
'                    Else
'                        playSound 12
'                        rng.Select
'                        Debug.Print w & vbTab & h
'                        Stop
'                    End If
                End If
            Else
                Exit Function
            End If
        Else
            On Error Resume Next
            Set inlsp = rng.InlineShapes.AddPicture(fileName:=url, _
                            LinkToFile:=False, SaveWithDocument:=True, Range:=rng)
            On Error GoTo 0
            If Not inlsp Is Nothing Then
                If inlsp.Range.tables.Count > 0 Then
                    'resizePicture rng, inlsp, url, inlsp.Range.tables(1).PreferredWidth, inlsp.height * (inlsp.width / inlsp.Range.tables(1).PreferredWidth)
                    Rem �������óB�z�䤤���Ϥ��A���ӹw�]�N�O���j�p
                Else
                    If Not IsWrapTextImage(textPart) Then
                        resizePicture rng, inlsp, url
                    End If
                End If
            Else
                Exit Function
            End If
        End If
    Else 'base64�s�X���Ϥ�
        
        ' ���Jbase64�s�X���Ϥ�
        Set inlsp = InsertBase64Image(url, "tempImage.png", rng)
        If Not IsWrapTextImage(textPart) Then
            resizePicture rng, inlsp, url
        End If
        
    End If
    
    Rem �]�w�Ϥ��榡
    Rem inlineShape�榡

    Dim shp As Shape
    If align <> vbNullString And hspace <> vbNullString Then
        Select Case align
            Case "right"
                Set shp = inlsp.ConvertToShape
                With shp.WrapFormat
                    .Type = wdWrapSquare
                    .Side = wdWrapBoth
                    '.DistanceTop = CentimetersToPoints(0.5)
                    .DistanceLeft = CentimetersToPoints(0.5)
                    .Parent.RelativeHorizontalPosition = wdRelativeHorizontalPositionRightMarginArea
                    .Parent.RelativeVerticalPosition = wdRelativeVerticalPositionTopMarginArea
                    .Parent.Left = WdShapePosition.wdShapeRight
                    'shp.Top = WdShapePosition.wdShapeTop
    '                shp.Left = ActiveDocument.PageSetup.PageWidth - shp.width - CentimetersToPoints(1) ' �]�w�k��Z��
    '                shp.Top = CentimetersToPoints(1) ' �]�w�W��Z��
                End With
            Case "left"
                Set shp = inlsp.ConvertToShape
                With shp.WrapFormat
                    .Type = wdWrapSquare
                    .Side = wdWrapBoth
                    '.DistanceTop = CentimetersToPoints(0.5)
                    .DistanceRight = CentimetersToPoints(0.5)
                
                    .Parent.RelativeHorizontalPosition = wdRelativeHorizontalPositionLeftMarginArea
                    .Parent.RelativeVerticalPosition = wdRelativeVerticalPositionTopMarginArea
                    .Parent.Left = WdShapePosition.wdShapeLeft
    '                shp.Top = wdShapeTop
                End With
            Case "absbottom"
            Case Else
                playSound 12
                Stop 'for check
        End Select
    End If
    Rem Shape��¶�Ϯ榡
    Dim imgStyle As String, float As String, marginLeft, marginRight ', border As String

    'ex: float:right;margin-left:10px;margin-right:10px;"
    imgStyle = getHTML_AttributeValue("style", textPart)
    If imgStyle <> vbNullString Then
        If inlsp.Range.tables.Count = 0 Then
            If InStr(imgStyle, "float:") Then
                float = VBA.Trim(VBA.Mid(imgStyle, VBA.InStr(imgStyle, "float:") + VBA.Len("float:"), VBA.InStr(VBA.InStr(imgStyle, "float:"), imgStyle, ";") - (VBA.InStr(imgStyle, "float:") + VBA.Len("float:"))))
            End If
            If InStr(imgStyle, "margin-left:") Then
                marginLeft = VBA.Val(VBA.Mid(imgStyle, VBA.InStr(imgStyle, "margin-left:") + VBA.Len("margin-left:"), VBA.InStr(VBA.InStr(imgStyle, "margin-left:"), imgStyle, ";") - (VBA.InStr(imgStyle, "margin-left:") + VBA.Len("margin-left:"))))
            End If
            If InStr(imgStyle, "margin-right:") Then
                marginRight = VBA.Val(VBA.Mid(imgStyle, VBA.InStr(imgStyle, "margin-right:") + VBA.Len("margin-right:"), VBA.InStr(VBA.InStr(imgStyle, "margin-right:"), imgStyle, ";") - (VBA.InStr(imgStyle, "margin-right:") + VBA.Len("margin-right:"))))
            End If
            If float <> "" Or VBA.IsEmpty(marginLeft) = False Or VBA.IsEmpty(marginRight) = False Then
                ' �]�m�Ϥ�����¶�Ϥ覡�M����覡
                Set shp = inlsp.ConvertToShape
                With shp
                    .WrapFormat.Type = WdWrapType.wdWrapTight ' wdWrapSquare
                    Select Case float
                        Case vbNullString
                        Case "left"
                            .Left = WdShapePosition.wdShapeLeft
                            '.WrapFormat.Side = WdWrapSideType.wdWrapLeft
                        Case "right"
                            .Left = WdShapePosition.wdShapeRight
                            '.WrapFormat.Side = WdWrapSideType.wdWrapRight ' ������float:right
                        Case Else
                            Stop ' check
                    End Select
                    If marginLeft <> 0 Then
                        .WrapFormat.DistanceLeft = marginLeft ' ������margin-left:10px
                    End If
                    If marginRight <> 0 Then
                        .WrapFormat.DistanceRight = marginRight ' ������margin-right:10px
                    End If
                End With
            End If 'If float <> "" Or VBA.IsEmpty(marginLeft) = False Or VBA.IsEmpty(marginRight) = False Then
            
        End If 'inlsp.Range.tables.Count = 0 Then

        '��LStyle�ݩʳ]�w
        arr = VBA.Split(imgStyle, ";")
        For Each e In arr
            If e <> vbNullString Then
                e = VBA.Trim(e)
                If VBA.Left(e, VBA.Len("border-style:")) = "border-style:" Then
                    l = VBA.Len("border-style:")
                    attr = VBA.Trim(VBA.Mid(e, l + 1))
                    Select Case attr
                        Case "initial"
                            inlsp.Line.Visible = msoFalse ' �L���
                        Case "solid"
                            inlsp.Line.Visible = msoTrue ' ���
                            inlsp.Line.Style = msoLineSolid
                        Case Else
                            playSound 12
                            Debug.Print e & vbTab & attr
                            rng.Select
                            Stop
                    End Select
                ElseIf VBA.Left(e, VBA.Len("border-color:")) = "border-color:" Then
                    l = VBA.Len("border-color:")
                    attr = VBA.Trim(VBA.Mid(e, l + 1))
                    Select Case attr
                        Case "initial", "black"
                            inlsp.Line.ForeColor.RGB = RGB(0, 0, 0) ' ����C�⬰�¦�
                        Case Else
                            playSound 12
                            Debug.Print e & vbTab & attr
                            rng.Select
                            Stop
                    End Select
                ElseIf VBA.Left(e, VBA.Len("vertical-align:")) = "vertical-align:" Then
                    l = VBA.Len("vertical-align:")
                    attr = VBA.Trim(VBA.Mid(e, l + 1))
                    Select Case attr
                        Case "baseline"
                            inlsp.Range.ParagraphFormat.BaseLineAlignment = wdBaselineAlignBaseline
                            'inlsp.Range.ParagraphFormat.BaseLineAlignment = wdBaselineAlignCenter
                        Case "middle", "bottom", "text-bottom"
                            inlsp.Range.ParagraphFormat.BaseLineAlignment = wdBaselineAlignCenter
                        Case Else
                            playSound 12
                            Debug.Print e & vbTab & attr
                            rng.Select
                            Stop
                    End Select
                ElseIf VBA.Left(e, VBA.Len("border-width:")) = "border-width:" Then
                    l = VBA.Len("border-width:")
                    attr = VBA.Trim(VBA.Mid(e, l + 1))
                    If VBA.Right(attr, 2) = "px" Then
                        attrSetting = VBA.Val(attr)
                        If VBA.IsNumeric(attrSetting) Then
                            If attrSetting <> 0 Then
                                inlsp.Line.Visible = msoTrue
                                inlsp.Line.Weight = VBA.CSng(attrSetting)   ' ��ؼe��
                            Else
                                inlsp.Line.Visible = msoFalse ' �L���
                            End If
                        Else
                            playSound 12
                            Debug.Print e & vbTab & attr
                            rng.Select
                            Stop
                        End If
                    Else
                        playSound 12
                        Debug.Print e & vbTab & attr
                        rng.Select
                        Stop
                    End If
                ElseIf VBA.Left(e, VBA.Len("padding:")) = "padding:" Then
                    l = VBA.Len("padding:")
                    attr = VBA.Trim(VBA.Mid(e, l + 1))
                    Select Case attr
                        Case "0px"
                            inlsp.Range.ParagraphFormat.SpaceBefore = 0 '�q�e���Z
                        Case Else
                            playSound 12
                            Debug.Print e & vbTab & attr
                            rng.Select
                            Stop
                    End Select
                ElseIf VBA.Left(e, VBA.Len("margin:")) = "margin:" Then
                    l = VBA.Len("margin:")
                    attr = VBA.Trim(VBA.Mid(e, l + 1))
                    Select Case attr
                        Case "0px"
                            inlsp.Range.ParagraphFormat.SpaceAfter = 0 '�q�ᶡ�Z
                        Case Else
                            playSound 12
                            Debug.Print e & vbTab & attr
                            rng.Select
                            Stop
                    End Select
                ElseIf VBA.Left(e, VBA.Len("border:")) = "border:" Then
                    l = VBA.Len("border:")
                    attr = VBA.Trim(VBA.Mid(e, l + 1))
                    Select Case attr
                        Case "1px solid black"
                            With inlsp.Line
                                .Visible = msoTrue
                                .ForeColor.RGB = RGB(0, 0, 0) ' ����C�⬰�¦�
                                .Weight = 1 ' ��ؼe��
                                .Style = msoLineSolid ' ��ؼ˦�����u
                            End With
                        Case "0"
                            inlsp.Line.Visible = msoFalse
                        Case Else
                            playSound 12
                            Debug.Print e & vbTab & attr
                            rng.Select
                            Stop
                    End Select
                ElseIf VBA.Left(e, VBA.Len("color:")) = "color:" Then
                    '���B�z �Ϥ���� color�ݩʡH
                ElseIf VBA.Left(e, VBA.Len("font-size:")) = "font-size:" Then
                    '���B�z �Ϥ���� font size�ݩʡH
                ElseIf VBA.Left(e, VBA.Len("background-color:")) = "background-color:" Then
                    l = VBA.Len("background-color:")
                    attr = VBA.Trim(VBA.Mid(e, l + 1))
                    If VBA.Left(attr, 1) = "#" Then
                        inlsp.Fill.BackColor.RGB = RGBFormColorCode(attr)
                    Else
                        playSound 12
                        Debug.Print e & vbTab & attr
                        rng.Select
                        Stop
                    End If
                ElseIf VBA.Left(e, VBA.Len("display:")) = "display:" Then 'Copilot�j���ġGdisplay: block; �O�Ψӳ]�w��������ܤ覡�A�b�o�̹Ϥ��Q�]�w�����Ť����]block�^�A�o�˹Ϥ��|�W�e�@��A�����q���CWord VBA ���èS�������������ݩʡA���i�H�q�L�վ�q���M�Ϥ����G���Ӽ����o�ӮĪG�C20241017
                    l = VBA.Len("display:")
                    attr = VBA.Trim(VBA.Mid(e, l + 1))
                    If attr = "block" Then
                        If inlsp.Range.Paragraphs(1).Range.text <> Chr(13) Then
                            inlsp.Range.InsertParagraphBefore 'display: block; �V �N�Ϥ���m�b�q�����A�T�O�Ϥ��W�e�@��C
                            inlsp.Range.InsertParagraphAfter
                        End If
                    End If
                    inlsp.Fill.BackColor.RGB = RGBFormColorCode(attr)
                ElseIf VBA.Left(e, VBA.Len("margin-left:")) = "margin-left:" Then
                    l = VBA.Len("margin-left:")
                    attr = VBA.Trim(VBA.Mid(e, l + 1))
                    Select Case attr
                        Case "auto"
                            If shp Is Nothing Then
                                inlsp.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                            Else
                                shp.Left = WdShapePosition.wdShapeCenter
                            End If
                        Case Else
                            playSound 12
                            Debug.Print e & vbTab & attr
                            rng.Select
                            Stop
                    End Select
                ElseIf VBA.Left(e, VBA.Len("margin-right:")) = "margin-right:" Then
                    l = VBA.Len("margin-right:")
                    attr = VBA.Trim(VBA.Mid(e, l + 1))
                    If attr = "auto" Then
                        If shp Is Nothing Then
                            inlsp.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                        Else
                            shp.Left = WdShapePosition.wdShapeCenter
                        End If
                    End If
                ElseIf VBA.Left(e, VBA.Len("float:")) = "float:" Then
                    l = VBA.Len("float:")
                    attr = VBA.Trim(VBA.Mid(e, l + 1))
                    ' �]�m�Ϥ�����¶�Ϥ覡�M����覡
                    Set shp = inlsp.ConvertToShape
                    With shp
                        .LockAspectRatio = msoTrue ' ��w�Ϥ����
                        .WrapFormat.Type = WdWrapType.wdWrapTight ' wdWrapSquare
                        Select Case float
                            Case vbNullString
                            Case "left"
                                .Left = WdShapePosition.wdShapeLeft
                                '.WrapFormat.Side = WdWrapSideType.wdWrapLeft
                            Case "right"
                                .Left = WdShapePosition.wdShapeRight
                                '.WrapFormat.Side = WdWrapSideType.wdWrapRight ' ������float:right
                            Case Else
                                Stop ' check
                        End Select
                    End With
                ElseIf VBA.Left(e, VBA.Len("line-height:")) = "line-height:" Then
                    l = VBA.Len("line-height:")
                    attr = VBA.Trim(VBA.Mid(e, l + 1))
                    inlsp.Range.ParagraphFormat.LineSpacingRule = wdLineSpaceAtLeast
                    inlsp.Range.ParagraphFormat.LineSpacing = word.LinesToPoints(VBA.CSng(attr)) ' ����HTML���� line-height

'                ElseIf VBA.Left(e, VBA.Len(":")) = ":" Then
'                    l = VBA.Len(":")
'                    attr = VBA.Trim(VBA.Mid(e, l + 1))
'                    inlsp
                Else
                        playSound 12
                        Debug.Print e & vbTab & attr
                        rng.Select
                        Stop
                    
                'If VBA.InStr(imgStyle, "border") Then
                    'border = VBA.Trim(VBA.Mid(imgStyle, VBA.InStr(imgStyle, "border:") + VBA.Len("border:"), VBA.InStr(VBA.InStr(imgStyle, "border:"), imgStyle, ";") - (VBA.InStr(imgStyle, "border:") + VBA.Len("border:"))))
                    'If border <> "" Then
                        'If VBA.Right(border, 1) = ";" Then border = VBA.Left(border, VBA.Len(border) - 1)
'                                Select Case border
'                                    Case "1px solid black"
'                                        With inlsp.Line
'                                            .Visible = msoTrue
'                                            .ForeColor.RGB = RGB(0, 0, 0) ' ����C�⬰�¦�
'                                            .Weight = 1 ' ��ؼe��
'                                            .Style = msoLineSolid ' ��ؼ˦�����u
'                                        End With
'                                    Case "0"
'                                        inlsp.Line.Visible = msoFalse
'                                    Case Else
'                                        inlsp.Select
'                                        playSound 12 'for check
'                                        Stop
'                                End Select
                        'End If
                    'End If
                'End If
                End If
            End If 'If e <> vbNullString Then
        Next e
    End If
    
    If shp Is Nothing Then
        If IsWrapTextImage(textPart) Then
'            inlsp.Range.Select
'            playSound 12 'for check
'            Stop
            Set shp = inlsp.ConvertToShape
            shp.WrapFormat.Type = wdWrapTopBottom
        End If
    End If
    
    If shp Is Nothing Then
        rng.ParagraphFormat.BaseLineAlignment = wdBaselineAlignCenter
    End If
    Set insertHTMLImage = inlsp
    SystemSetup.playSound 0.411
End Function



Rem ���oHTML����檺�ݩʭ� 20241011 creedit_with_Copilot�j���ġGHTML����ഫ�M�ݩʳ]�m�GHTML����ഫ�M�ݩʳ]�m
Function GetHTMLAttributeValue(attributeName As String, html As String) As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    
    ' ��l�ƥ��h��F����H
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
'    regex.Pattern = attributeName & "=""'[""']"
    regex.Pattern = attributeName & "="".*?[""']"
    
    Set matches = regex.Execute(html)
    If matches.Count > 0 Then
        'GetHTMLAttributeValue = matches(0).SubMatches(0)
        GetHTMLAttributeValue = VBA.Replace(VBA.Replace(matches(0).Value, "href=""", vbNullString), """", vbNullString)
    Else
        GetHTMLAttributeValue = ""
    End If
End Function

Rem �M���@����html tags HTML����
Function StripHTMLTags(html As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "<.*?>"
    regex.Global = True
    StripHTMLTags = regex.Replace(html, "")
End Function
Function ParseHTMLList(html As String) As Collection
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim listItems As New Collection
    
    ' ��l�ƥ��h��F����H
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "<li.*?>(.*?)</li>"
    
    Set matches = regex.Execute(html)
    For Each match In matches
        listItems.Add match.SubMatches(0)
    Next match
    
    Set ParseHTMLList = listItems
End Function



Rem 20241010��y�� �M���b���Ҷ��S�����󤺮e��HTML�ż���
Sub RemoveEmptyTags(rng As Range)
    Dim rngOriginal As Range, arr, e
    Set rngOriginal = rng.Document.Range(rng.start, rng.End)
    arr = Array(">" & VBA.Chr(11) & "<", "><", ">" & VBA.Chr(13) & "<")
    rng.Find.MatchWildcards = False
    With rng.Find
        .ClearFormatting
        For Each e In arr
            .text = e
            .Wrap = wdFindStop
            Do While .Execute()
                Do Until rng.Characters(1) = "<"
                    rng.MoveStart , -1
                Loop
'                rng.Select 'for check
                rng.MoveEndUntil ">"
                rng.MoveEnd 1
'                rng.Select 'for check
                If Not VBA.Left(rng.text, 2) = "</" And (VBA.InStr(rng.text, e & "/") Or VBA.InStr(rng.text, ">" & VBA.Chr(13) & "/>")) _
                        And VBA.Mid(rng.text, VBA.InStr(rng.text, "/") + 1, VBA.Len(rng.text) - VBA.InStr(rng.text, "/") - 1) _
                            = rng.Document.Range(rng.start + 1, rng.start + 1 + VBA.Len(rng.text) - VBA.InStr(rng.text, "/") - 1) Then
                    If VBA.InStr(rng.text, "</td>") = 0 And VBA.InStr(rng.text, "<td>") = 0 Then '�x�s�椣��M���I20241016
                        rng.text = vbNullString
                    End If
                End If
                If rng.Characters.Count = 1 And rng.Characters(1).text = VBA.Chr(13) And rng.Paragraphs(1).Range.Characters.Count = 1 Then
                    rng.Characters(1).text = vbNullString
                End If
                rng.Collapse wdCollapseEnd
                'rng.SetRange rng.End, rngOriginal.End
            Loop
            rng.SetRange rngOriginal.start, rngOriginal.End
        Next e
    End With
End Sub
Rem ���o�L�ǦC��]<ul></ur>�^���d�� 20241010creedit_with_Copilot�j���ġGHTML�W�s���ഫ��Word VBA�Ghttps://sl.bing.net/bXsbFqI2cz6
Private Function getRangeFromULToUL_UnorderedListRange(rng As Range) As Range
    Dim startRange As Range
    Dim endRange As Range
    If VBA.InStr(rng.text, "<ul") Then
    ' �d�� <ul> ����
        Set startRange = rng.Document.Range(rng.start, rng.End)
        With startRange.Find
            .ClearFormatting
            .text = "<ul"
            If .Execute Then
                startRange.Collapse Direction:=wdCollapseStart
            End If
        End With
        
        ' �d�� </ul> ����
        Set endRange = rng.Document.Range(startRange.End, rng.End)
        With endRange.Find
            .ClearFormatting
            .text = "</ul>"
            If .Execute Then
                endRange.Collapse Direction:=wdCollapseEnd
            End If
        End With
        
        ' �]�w�d��
        If Not (startRange.start = rng.start And endRange.End = rng.End) Then
            Set getRangeFromULToUL_UnorderedListRange = rng.Document.Range(startRange.start, endRange.End)
        End If
    End If
End Function



Rem 20241009 ���oHTML�����ݩʤ��� pro ���]�t�u="�v,start �j�M���_�l��m
Private Function getHTML_AttributeValue(atrb As String, textIncludingAttribute As String, Optional marker As String)
    Dim lenatrb As Byte
    Select Case marker
        Case vbNullString
            atrb = atrb & "="""
            'atrb = atrb & "="
        Case ":"
            atrb = atrb & ": "
            'atrb = atrb & ":"
    End Select
retest:
    If VBA.InStr(textIncludingAttribute, atrb) > 0 Then
        lenatrb = VBA.Len(atrb)
        getHTML_AttributeValue = VBA.Mid(textIncludingAttribute, VBA.InStr(textIncludingAttribute, atrb) + lenatrb, _
            VBA.InStr(VBA.InStr(textIncludingAttribute, atrb) + lenatrb, textIncludingAttribute, """") - (VBA.InStr(textIncludingAttribute, atrb) + lenatrb))
'        getHTML_AttributeValue = VBA.Mid(textIncludingAttribute, VBA.InStr(start, textIncludingAttribute, atrb) + lenatrb, _
            VBA.InStr(start, VBA.InStr(start, textIncludingAttribute, atrb) + lenatrb, textIncludingAttribute, """") _
                - (VBA.InStr(start, textIncludingAttribute, atrb) + lenatrb))
    Else
        If VBA.Right(atrb, 2) = ": " Then
            atrb = VBA.Mid(atrb, 1, VBA.Len(atrb) - 1)
            GoTo retest
        End If
    End If
End Function

Rem ���J�Ϥ���A�ھګe��r���j�p�۰ʽվ�Ϥ��j�p 20241009 creedit_with_Copilot�j���ġGWordVBA �Ϥ��۰ʽվ�j�p�Ghttps://sl.bing.net/e1S3H59hvI4
Private Function getImageUrl(textIncludingSrc As String)
    getImageUrl = VBA.Mid(textIncludingSrc, VBA.InStr(textIncludingSrc, "src=""") + 5, _
        VBA.InStr(VBA.InStr(textIncludingSrc, "src=""") + 5, textIncludingSrc, """") - (VBA.InStr(textIncludingSrc, "src=""") + 5))
End Function
Rem ���s�վ�Ϥ��j�p�A�Y�L���w width�Pheight �h�Ѧҫe���r���j�p�����ȳ]�w
Private Sub resizePicture(rng As Range, pic As InlineShape, url As String, Optional width As Single = 0, Optional height As Single = 0) ', Optional imgHtml As String)
    If width > 0 And height > 0 Then
        pic.width = width
        pic.height = height
    ElseIf width > 0 Then
        pic.LockAspectRatio = msoTrue
        pic.width = width
    ElseIf height > 0 Then
        pic.LockAspectRatio = msoTrue
        pic.height = height
    Else
'        If Not IsWrapTextImage(imgHtml) Then
            Dim fontSizeBefore As Single
            Dim fontSizeAfter As Single
            Dim avgFontSize As Single
            ' ����e��r���j�p
            If rng.start > 1 Then
                fontSizeBefore = rng.Characters.First.Previous.font.Size
            Else
                fontSizeBefore = rng.Characters.First.font.Size
            End If
        
            If rng.End < rng.Document.Content.End Then
                fontSizeAfter = rng.Characters.Last.Next.font.Size
            Else
                fontSizeAfter = rng.Characters.Last.font.Size
            End If
        
            ' �p�⥭���r���j�p
            avgFontSize = (fontSizeBefore + fontSizeAfter) / 2
        
            ' �վ�Ϥ��j�p
            pic.LockAspectRatio = msoTrue
            If Not IsValidImage_LoadPicture(url) Then
                pic.height = avgFontSize
                If Not SeleniumOP.IsWDInvalid() Then
                    pic.Range.Hyperlinks.Add pic.Range, WD.url
                End If
            Else
                pic.height = avgFontSize * 2 ' �ھڻݭn�վ���
                pic.width = pic.height * pic.width / pic.height
            End If
'        Else
'            playSound 12
'            pic.Range.Select
'            Stop
'        End If
    End If
End Sub

Rem �ѪRHTML�ô��J�M�� 20241011 creedit_with_Copilot�j���ġGhttps://sl.bing.net/gbeqh0TAks8�GHTML����ഫ�M�ݩʳ]�m
Rem �ѪRHTML���e�A�����M�涵�ءA�M��bWord�����J�������M��˦��Chttps://sl.bing.net/bhFU3zNMSom
Sub InsertHTMLList(html As String)
    Dim doc As Document
    Dim listItems As Collection
    Dim listItem As Variant
    Dim rng As Range
    
    ' �ѪRHTML
    Set listItems = ParseHTMLList(html)
    
    ' ���J�M��
    Set doc = ActiveDocument
    Set rng = doc.Range(start:=doc.Content.End - 1, End:=doc.Content.End - 1)
    
    ' �}�l�M��
    rng.ListFormat.ApplyBulletDefault
    
    ' ��R�M�椺�e
    For Each listItem In listItems
        rng.text = StripHTMLTags(VBA.CStr(listItem))
        rng.ParagraphFormat.Alignment = wdAlignParagraphLeft
        rng.font.Name = "�з���"
        rng.InsertParagraphAfter
        Set rng = rng.Next(wdParagraph, 1) '.Range
    Next listItem
End Sub

Rem �ѪRHTML���e�A�������B��B�椸��B�Ϥ��M��r 20241011 creedit_with_Copilot�j���ġGhttps://sl.bing.net/fQ5lVr8PLye
Private Function parseHTMLTable(html As String) As Collection
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim tables As New Collection
    Dim rows As New Collection
    Dim cells As New Collection
    Dim table, row
    
    ' ��l�ƥ��h��F����H
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    
    ' �ǰt���
    regex.Pattern = "<table.*?>(.*?)</table>"
    Set matches = regex.Execute(html)
    For Each match In matches
        tables.Add match.SubMatches(0)
    Next match
    
    ' �ǰt��/�C
    regex.Pattern = "<tr.*?>(.*?)</tr>"
    For Each table In tables
        Set matches = regex.Execute(table)
        For Each match In matches
            rows.Add match.SubMatches(0)
        Next match
    Next table
    
    ' �ǰt�椸��
    regex.Pattern = "<td.*?>(.*?)</td>"
    For Each row In rows
        Set matches = regex.Execute(row)
        For Each match In matches
            cells.Add match.SubMatches(0)
        Next match
    Next row
    
    Set parseHTMLTable = cells
End Function

Rem ���U�ӡA�z�i�H�bWord���Ыت��ô��J���������e creedit_with_Copilot�j���� 20241011
Private Sub insertHTMLTable(rngHtml As Range, Optional domainUrlPrefix As String, Optional fontName As String)
    Dim html As String
    Dim tbl As word.table
    Dim cells As Collection
    Dim cell As Variant
    Dim row As Integer
    Dim col As Integer
    Dim img As InlineShape
    Dim rng As Range, rngClose As Range, c As cell, arr, e, obj As Object, attr As String, l
    Dim align As String
    Dim bgcolorTbl As String
    Dim bgcolorTd As String
    Dim tblWidth As Single
    Dim float As String
    Dim colCount As Byte
    Dim rowCount As Long
    
'    Dim st As Long
'    Dim rngTbl As Range
    
'    Dim imgSrc As String
'    Dim imgWidth As Single
'    Dim imgHeight As Single
    
    
'    Dim ur As UndoRecord
'    SystemSetup.stopUndo ur, "InsertHTMLTable"
    
    With rngHtml
        html = .text
'        st = .start

        Rem ��~���դ~�o�{�A�p�G�ڦb�ন���e���奻���]�w�n Range.font.Name = "Lucida Sans Unicode" ���b�ন����A�N�i�H�b�Q�n�O "Lucida Sans Unicode" �r�������Цr���W�]�w���o�Ӧr���F�C�P���P���@�g���g�ۡ@�n�L�������� �ҥH����b����ഫ��]�w�A�n���b��r����e�����w ��������
        '.font.Name = "Lucida Sans Unicode"
        If fontName <> vbNullString Then
            .font.Name = fontName
        End If
        
    End With
    ' �ѪRHTML
    Set cells = parseHTMLTable(html)
    
'    Set rngTbl = rngHtml.Document.Range(st, st)
    'rngHtml.text = vbNullString
    
    ' �p�����
    rowCount = UBound(Split(html, "<tr")) '- 1
    If rowCount = 0 Then
        Exit Sub
    Else
        colCount = cells.Count / rowCount 'UBound(Split(html, "<td")) ' - 1'creedit_with_Copilot�j���� 20241013
    End If
    
    ' ���J���
    Set tbl = rngHtml.tables.Add(Range:=rngHtml, NumRows:=rowCount, NumColumns:=colCount)
    
    
     ' �]�m����ݩ�
'    align = getHTML_AttributeValue("align", html)
'    bgcolorTbl = getHTML_AttributeValue("bgcolor", html)
'    tblWidth = CSng(getHTML_AttributeValue("width", html))
    align = getHTML_AttributeValue("align", html)
    bgcolorTbl = getHTML_AttributeValue("bgcolor", VBA.Mid(html, VBA.InStr(html, "<table"), VBA.InStr(VBA.InStr(html, "<table"), html, ">") - VBA.InStr(html, "<table") + 1))
    bgcolorTd = getHTML_AttributeValue("bgcolor", VBA.Mid(html, VBA.InStr(html, "<td"), VBA.InStr(VBA.InStr(html, "<td"), html, ">") - VBA.InStr(html, "<td") + 1))
    tblWidth = VBA.CSng(VBA.Val((getHTML_AttributeValue("width", html, ":"))))
    float = getHTML_AttributeValue("float", html, ":")
    If float <> vbNullString Then
        float = VBA.Mid(float, 1, VBA.InStr(float, ";") - 1)
        If VBA.Right(float, 1) = ";" Then
            float = VBA.Left(float, VBA.Len(float) - 1)
        End If
    End If
    
    If align = "left" Then
        tbl.rows.Alignment = wdAlignRowLeft
    ElseIf align = "center" Then
        tbl.rows.Alignment = wdAlignRowCenter
    ElseIf align = "right" Then
        tbl.rows.Alignment = wdAlignRowRight
    End If
    
    If bgcolorTbl <> "" Then
        If VBA.Left(bgcolorTbl, 1) = "#" Then
            arr = HTML2Doc.ColorCodetoRGB(bgcolorTbl)
            tbl.Shading.BackgroundPatternColor = RGB(arr(0), arr(1), arr(2))
        Else
            If bgcolorTbl = "white" Then
                tbl.Shading.BackgroundPatternColor = RGB(255, 255, 255)
            Else
                playSound 12 'for check
                Stop
            End If
        End If
    Else
        With tbl.Shading
            .Texture = wdTexture10Percent
            .ForegroundPatternColor = wdColorAutomatic
            .BackgroundPatternColor = wdColorAutomatic
        End With
    End If
    

    tbl.PreferredWidthType = wdPreferredWidthPoints
    tbl.PreferredWidth = tblWidth
    
    ' ��R��椺�e
    row = 1
    col = 1
    For Each cell In cells 'cell�Ghtml text in a cell
        If cell <> vbNullString Then
            Set c = tbl.cell(row, col)
            'c.Range.text = cell
            Set rng = c.Range.Document.Range(c.Range.start, c.Range.End - 1) '1�GChr(13) & Chr(7)�]�bWordVBA���o�ө�strat end �ݩʭȥu�@ 1�I�^20241016
            rng.text = cell
            ' �ˬd�O�_�]�t�Ϥ�
            If InStr(cell, "<img") Then
                'c.Range.text = cell & Chr(13) & Chr(7)
                ''''''Set img = insertHTMLImage(html, c.Range, domainUrlPrefix)
                'Set rng = c.Range.Document.Range(c.Range.start, c.Range.End - 2) '2=Len(Chr(13) & Chr(7))
                    '�B�z�Ϥ�
                    With rng.Find
                        Do While .Execute(findText:="<img ")
                            rng.MoveEndUntil ">" 'ex: <img style="float:right;margin-left:15px;margin-right:15px;" src="/image/3.jpg" width="200" height="297"
    '                        '�ɥ��ܼ�
    '                        url = rng.text
                            rng.End = rng.End + 1 '�]�t ">"
                            'rng.text = vbNullString
                            'Set img = insertHTMLImage(c.Range.text, rng, domainUrlPrefix)
                            Set img = insertHTMLImage(rng.text, rng, domainUrlPrefix)
                            If Not img Is Nothing Then
                                If rng.InlineShapes.Count > 0 Then
                                    rng.InlineShapes(1).width = img.width 'imgWidth
                                    rng.InlineShapes(1).height = img.height 'imgHeight
                                ElseIf rng.ShapeRange.Count > 0 Then
                                    rng.ShapeRange(1).width = img.width 'imgWidth
                                    rng.ShapeRange(1).height = img.height 'imgHeight
                                End If
                                rng.SetRange rng.start + 1, rng.End '1= "/"�Ҵ��J�Ϥ����w��Ŧr��
                                rng.text = vbNullString
                            Else
                                playSound 12
                                rng.Select
                                Stop 'for check
                            End If
                            'pCntr + VBA.Abs(10 - pCntr) '�U���Ϥ��ݭn�ɶ�
                            rng.SetRange rng.End, c.Range.End - 2
                        Loop '�B�z�Ϥ�
                    End With 'rng.Find
                'Set img = insertHTMLImage(c.Range.text, c.Range, domainUrlPrefix)
    ''''            imgSrc = getHTML_AttributeValue("src", VBA.CStr(cell))  'Mid(cell, InStr(cell, "src=") + 5, InStr(cell, """", InStr(cell, "src=") + 5) - InStr(cell, "src=") - 5)
    ''''            imgWidth = getHTML_AttributeValue("width", VBA.CStr(cell)) 'CSng(Mid(cell, InStr(cell, "width=") + 7, InStr(cell, """", InStr(cell, "width=") + 7) - InStr(cell, "width=") - 7))
    ''''            imgHeight = getHTML_AttributeValue("height", VBA.CStr(cell)) 'CSng(Mid(cell, InStr(cell, "height=") + 8, InStr(cell, """", InStr(cell, "height=") + 8) - InStr(cell, "height=") - 8))
    ''''            tbl.cell(row, col).Range.InlineShapes.AddPicture fileName:=imgSrc, LinkToFile:=False, SaveWithDocument:=True
    '            c.Range.InlineShapes(1).width = img.width 'imgWidth
    '            c.Range.InlineShapes(1).height = img.height 'imgHeight
    '            Set rng = c.Range.Document.Range(c.Range.End - 1, c.Range.End - 1)
    '            rng.text = StripHTMLTags(VBA.CStr(cell))
            Else
                c.Range.text = VBA.CStr(cell) 'StripHTMLTags(VBA.CStr(cell))
                'tbl.cell(row, col).Range.text = VBA.CStr(cell) 'StripHTMLTags(VBA.CStr(cell))
            End If
            
            '�ˬd�O�_�]�t�W�s��
            If VBA.InStr(cell, "<a ") > 0 Then
                '�B�z�W�s�� 20241016
                InsertHTMLLinks c.Range, domainUrlPrefix
            End If
            '�M�z<�K�K>����
            Set rng = c.Range.Document.Range(c.Range.start, c.Range.End - 1)
            Do While VBA.InStr(c.Range.text, "<")
                Do While rng.Find.Execute(findText:="<span ")
                    rng.MoveEndUntil ">"
                    rng.End = rng.End + 1
                    Set rngClose = rng.Document.Range(rng.End, c.Range.End - 1)
                    If Not rngClose.Find.Execute(findText:="</span>") Then
                        playSound 12 'for check
                        rng.Select
                        Stop
                    End If
                    Do While VBA.InStr(rng.Document.Range(rng.End, rngClose.start).text, "<span>")
                        rng.SetRange rng.End, c.Range.End - 1
                        rng.Find.Execute "<span>"
                        rngClose.SetRange rng.End, c.Range.End - 1
                        rngClose.Find.Execute "</span>"
                        If rngClose.End + 1 = c.Range.End Or VBA.InStr(rngClose.Document.Range(rngClose.End, c.Range.End - 1), "</span>") = 0 Then
                            Exit Do
                        End If
                    Loop
                    arr = VBA.Split(getHTML_AttributeValue("style", rng.text), ";")
                    For Each e In arr
                        If e <> vbNullString Then
                            e = VBA.Trim(e)
                            If VBA.Left(e, VBA.Len("font-size:")) = "font-size:" Then
                                l = VBA.Len("font-size:")
                                attr = VBA.Trim(VBA.Mid(e, l + 1))
                                Select Case attr
                                    Case "x-large", "x-large"
                                        rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size * 1.3
                                    Case "large", "large"
                                        rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size * 1.2
                                    Case "medium", "medium" '���B�z
                                    Case "small", "small"
                                        rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size * (5 / 6)
                                    Case "x-small", "x-small"
                                        rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size * (2 / 3)
                                    Case "xx-small", "xx-small"
                                        rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size * (1 / 2)
                                    Case Else
                                        playSound 12 'for check
                                        Debug.Print e
                                        rng.Select
                                        Stop
                                End Select
                            ElseIf VBA.Left(e, VBA.Len("text-decoration:")) = "text-decoration:" Then
                                l = VBA.Len("text-decoration:")
                                attr = VBA.Trim(VBA.Mid(e, l + 1))
                                Select Case attr
                                    Case "underline"
                                        rng.Document.Range(rng.End, rngClose.start).font.Underline = wdUnderlineSingle
                                    Case Else
                                        playSound 12 'for check
                                        Debug.Print e
                                        rng.Select
                                        Stop
                                End Select
                            ElseIf VBA.Left(e, VBA.Len("color:")) = "color:" Then
                                attr = VBA.Trim(VBA.Mid(e, VBA.Len("color:") + 1))
                                If VBA.Left(attr, 1) = "#" Then
                                    rng.Document.Range(rng.End, rngClose.start).font.Color = RGBFormColorCode(attr)
                                Else
                                    If attr = "red" Then
                                        rng.Document.Range(rng.End, rngClose.start).font.ColorIndex = wdRed
                                    Else
                                        playSound 12
                                        rng.Select
                                        Debug.Print e
                                        Stop 'for check
                                    End If
                                End If
                            Else
                                playSound 12 'for check
                                Debug.Print e
                                rng.Select
                                Stop
                            End If
                        End If 'if e<> vbNullString then
                    Next e 'For Each e In arr
                    rng.text = vbNullString: rngClose.text = vbNullString
                    rng.SetRange c.Range.start, c.Range.End - 1
                    rng.SetRange c.Range.start, c.Range.End - 1
                Loop 'Do While rng.Find.Execute(findText:="<span ")
                '��m<b></b>
                Do While rng.Find.Execute(findText:="<b>")
                    Set rngClose = rng.Document.Range(rng.End, c.Range.End - 1)
                    If Not rngClose.Find.Execute(findText:="</b>") Then
                        playSound 12 'for check
                        rng.Select
                        Stop
                    End If
                    rng.Document.Range(rng.End, rngClose.start).font.Bold = True
                    rng.text = vbNullString: rngClose.text = vbNullString
                    rng.SetRange c.Range.start, c.Range.End - 1
                Loop 'Do While rng.Find.Execute(findText:="<b>")
                '<strong></strong>
                Do While rng.Find.Execute(findText:="<strong>")
                    Set rngClose = rng.Document.Range(rng.End, c.Range.End - 1)
                    If Not rngClose.Find.Execute(findText:="</strong>") Then
                        playSound 12 'for check
                        rng.Select
                        Stop
                    End If
                    rng.Document.Range(rng.End, rngClose.start).font.Bold = True
                    rng.text = vbNullString: rngClose.text = vbNullString
                    rng.SetRange c.Range.start, c.Range.End - 1
                Loop
                '<p �K�K
                Do While VBA.InStr(c.Range.text, "<p ")
                    rng.Find.Execute findText:="<p "
                    rng.MoveEndUntil ">"
                    rng.End = rng.End + 1
                    arr = VBA.Split(getHTML_AttributeValue("style", rng.text), ";")
                    For Each e In arr
                        If e <> vbNullString Then
                        e = VBA.Trim(e)
                        If VBA.Left(e, VBA.Len("line-height:")) = "line-height:" Then
                            l = VBA.Len("line-height:")
                            attr = VBA.Replace(VBA.Trim(VBA.Mid(e, l + 1)), "px", vbNullString)
                            rng.ParagraphFormat.LineSpacing = VBA.CSng(attr)
                        ElseIf VBA.Left(e, VBA.Len("margin-top:")) = "margin-top:" Then
                            l = VBA.Len("margin-top:")
                            attr = VBA.Replace(VBA.Trim(VBA.Mid(e, l + 1)), "px", vbNullString)
                            rng.ParagraphFormat.SpaceBefore = VBA.CSng(attr)
                        ElseIf VBA.Left(e, VBA.Len("margin-bottom:")) = "margin-bottom:" Then
                            l = VBA.Len("margin-bottom:")
                            attr = VBA.Replace(VBA.Trim(VBA.Mid(e, l + 1)), "px", vbNullString)
                            rng.ParagraphFormat.SpaceAfter = VBA.CSng(attr)
                        ElseIf VBA.Left(e, VBA.Len("text-align:")) = "text-align:" Then
                            l = VBA.Len("text-align:")
                            attr = VBA.Trim(VBA.Mid(e, l + 1))
                            Select Case attr
                                Case "center"
                                    rng.ParagraphFormat.Alignment = wdAlignParagraphCenter
                                Case "left"
                                    rng.ParagraphFormat.Alignment = wdAlignParagraphLeft
                                Case "right"
                                    rng.ParagraphFormat.Alignment = wdAlignParagraphRight
                                Case Else
                                    playSound 12 'for check
                                    rng.Select
                                    Debug.Print e
                                    Stop
                            End Select
'                        ElseIf VBA.Left(e, VBA.Len(":")) = ":" Then
'                            l = VBA.Len(":")
'                            attr = VBA.Replace(VBA.Mid(e, l + 1), "px", vbNullString)
'                            rng.=
                        Else
                            playSound 12 'for check
                            rng.Select
                            Debug.Print e
                            Stop
                        End If
                        End If
                    Next e
                    arr = Empty
                    arr = VBA.Split(getHTML_AttributeValue("dir", rng.text), ";")
                    For Each e In arr
                        e = VBA.Trim(e)
                        If e = "ltr" Then
                            rng.ParagraphFormat.ReadingOrder = wdReadingOrderLtr
                        Else
                            playSound 12 'for check
                            rng.Select
                            Debug.Print e
                            Stop
                        End If
                    Next e
                    rng.text = vbNullString
                    rng.SetRange c.Range.start, c.Range.End - 1
                Loop 'VBA.InStr(c.Range.text, "<p style=")
                
                If VBA.InStr(c.Range.text, "<") Then
                    playSound 12 'for check
                    c.Range.Select
                    Debug.Print c.Range.text
                    Stop
                    c.Range.text = StripHTMLTags(VBA.CStr(cell))
                End If
                
                rng.SetRange c.Range.start, c.Range.End - 1
            Loop 'Do While VBA.InStr(c.Range.text, "<")
        
        End If 'If cell <> vbNullString Then
        
        '�]�w�U�@���x�s��y�СA�ǳƲ��ʨ�U�@���x�s��ǳƶ�J���e20241016
        col = col + 1
        If col > tbl.Columns.Count Then
            'tbl.rows.Add
            row = row + 1
            col = 1
        End If
        
    Next cell
    
    If float <> vbNullString Then
    '    'Dim shp As Shape
    '     '�N����ഫ��Shape��H
    '    Set shp = tbl.ConvertToShape
    '
    '    tbl.rows.WrapAroundText = True
    '     �]�m��¶�Ϥ覡
    '    shp.WrapFormat.Type = wdWrapSquare
    '    shp.WrapFormat.Side = wdWrapBoth
    '    shp.WrapFormat.DistanceTop = 0
    '    shp.WrapFormat.DistanceBottom = 0
    '    shp.WrapFormat.DistanceLeft = 0
    '    shp.WrapFormat.DistanceRight = 0
        With tbl.rows
            .WrapAroundText = True
            .RelativeHorizontalPosition = wdRelativeHorizontalPositionColumn
            .DistanceLeft = CentimetersToPoints(0.32)
            .DistanceRight = CentimetersToPoints(0.32)
            .VerticalPosition = CentimetersToPoints(0)
            .RelativeVerticalPosition = wdRelativeVerticalPositionParagraph
            .DistanceTop = CentimetersToPoints(0)
            .DistanceBottom = CentimetersToPoints(0)
            .AllowOverlap = False
            Select Case VBA.Trim(float)
                Case "left"
                    .HorizontalPosition = wdTableLeft
                Case "right"
                    .HorizontalPosition = wdTableRight
                Case Else
                    tbl.Select
                    playSound 12
                    Stop
            End Select
        End With
    End If
    
'    SystemSetup.contiUndo ur
End Sub

Rem 20241009 �NHTML�নWord��󤺤�Ccreedit_with_Copilot�j���ġGhttps://sl.bing.net/jij3PK59Rka
Sub innerHTML_Convert_to_WordDocumentContent(rngHtml As Range, Optional domainUrlPrefix As String, Optional fontName As String)
    If VBA.InStr(rngHtml.text, "<") = 0 Then Exit Sub
    
     SystemSetup.playSound 1
    
    Dim htmlStr As String, rng As Range, rngClose As Range, p As Paragraph, url As String, stRngHTML As Long, pCntr As Long
    Dim s As Integer '�@�� InStr() �O�U���G�ȥ�
    Dim l As Integer '�@�� Len() �O�U���G�ȥ�
    '�@���q���ܼƥΡA�ΰ}�C�O���
    Dim arr, arr1, e '�@���q�Τ@���ܼƥΡA�ΰ}�C�����O���
    Dim obj As Object '�@���q�Ϊ����ܼƥ�
    
    
'
'    Rem just for check
'    SystemSetup.ClipboardPutIn rngHtml.text
'    Stop
'
'    GoTo finish 'just for test
'
    
    
    '���o���}�e�󪺺���ȡ]���t���׽u�^
    If domainUrlPrefix = vbNullString Then
        If Not SeleniumOP.IsWDInvalid() Then ' domainUrlPrefix = "https://www.eee-learning.com"
            domainUrlPrefix = GetDomainUrlPrefix(WD.url)
        End If
    End If
    'SystemSetup.stopUndo ur, "InnerHTML_DocContent"
    stRngHTML = rngHtml.start
    htmlStr = rngHtml.text '�O�U�_�l��m
    
    Rem �e�m��z�奻
    'rngHtml.text = VBA.Replace(VBA.Replace(VBA.Replace(htmlStr, "</p>", vbNullString), "<p>", vbNullString), "&nbsp;", ChrW(160))
    htmlStr = rngHtml.text
    
    If VBA.InStr(htmlStr, "<sup>") Then
        Set rng = rngHtml.Document.Range(rngHtml.start, rngHtml.End)
        HTML2Doc.ConvertHTMLSupToWordSup rng
    End If
    If VBA.InStr(htmlStr, "<sub>") Then
        Set rng = rngHtml.Document.Range(rngHtml.start, rngHtml.End)
        HTML2Doc.ConvertHTMLSubToWordSub rng
    End If
    With rngHtml.Find
        .ClearFormatting
        '�m��
        If VBA.InStr(htmlStr, "<br>" & VBA.Chr(13)) Then .Execute "<br>^p", , , , , , , , , "^l", wdReplaceAll
        If VBA.InStr(htmlStr, "<br>") Then .Execute "<br>", , , , , , , , , "^l", wdReplaceAll
        If VBA.InStr(htmlStr, "<a style=""line-height:1.5;"" href=") Then .Execute "<a style=""line-height:1.5;"" href=", , , , , , , , , "<a href=", wdReplaceAll ' " �|�m���� ��
        If VBA.InStr(htmlStr, "&lt;") Then .Execute "&lt;", , , , , , , , , "��", wdReplaceAll
        If VBA.InStr(htmlStr, "&gt;") Then .Execute "&gt;", , , , , , , , , "��", wdReplaceAll
        '�M��
        If VBA.InStr(htmlStr, "<div>") Then .Execute "<div>", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "</div>") Then .Execute "</div>", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "&nbsp;") Then .Execute "&nbsp;", , , , , , , , , vbNullString, wdReplaceAll '"&nbsp;" = ChrW(160)
        If VBA.InStr(htmlStr, "<p>") Then .Execute "<p>", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "</p>") Then .Execute "</p>", , , , , , , , , vbNullString, wdReplaceAll
'        If VBA.InStr(htmlStr, "<div>" & ChrW(160) & "</div>") Then .Execute "<div>" & ChrW(160) & "</div>", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, VBA.ChrW(160)) Then .Execute VBA.ChrW(160), , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "<wbr>") Then .Execute "<wbr>", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, " class=""colorbox cboxElement""") Then .Execute " class=""colorbox cboxElement""", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, " class=""colorbox colorbox-insert-image cboxElement""") Then .Execute " class=""colorbox colorbox-insert-image cboxElement""", , , , , , , , , vbNullString, wdReplaceAll '
        'If VBA.InStr(htmlStr, "<a class=""colorbox colorbox-insert-image cboxElement"" href=") Then .Execute "<a class=""colorbox colorbox-insert-image cboxElement"" href=", , , , , , , , , "<a href=", wdReplaceAll ' " �|�m���� ��
        If VBA.InStr(htmlStr, " rel=""group-all""") Then .Execute " rel=""group-all""", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "<o:p></o:p>") Then .Execute "<o:p></o:p>", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "<span></span>") Then .Execute "<span></span>", , , , , , , , , vbNullString, wdReplaceAll
        Rem ������\�νѦpWord���s��K�W�A�G�h���ݽX�B�ýX
        If VBA.InStr(htmlStr, "<o:p>") Then .Execute "<o:p>", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "</o:p>") Then .Execute "</o:p>", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "<span style=""color:#ffffff;"">ppp</span>") Then .Execute "<span style=""color:#ffffff;"">ppp</span>", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "<!--EndFragment-->") Then .Execute "<!--EndFragment-->", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, " face=""�з���""") Then .Execute " face=""�з���""", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "<span class=""Apple-tab-span"" style=""white-space: pre;""> </span>") Then .Execute "<span class=""Apple-tab-span"" style=""white-space: pre;""> </span>", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "<!--break-->") Then .Execute "<!--break-->", , , , , , , , , vbNullString, wdReplaceAll
    End With
    
    SystemSetup.playSound 1
    Set rng = rngHtml.Document.Range(rngHtml.start, rngHtml.End)
    '�M���ż���
    RemoveEmptyTags rngHtml
    
    Set rng = rngHtml.Document.Range(rngHtml.start, rngHtml.End)
    Rem �L�ǲM�檺�B�z
    unorderedListPorc_HTML2Word rng
    
    For Each p In rngHtml.Paragraphs
        pCntr = pCntr + 1
        If pCntr Mod 30 = 0 Then SystemSetup.playSound 1
        
        If p.Range.text = "/" & Chr(13) Or p.Range.text = Chr(13) Or p.Range.tables.Count > 0 Then
            GoTo nextP
        End If
        Set rng = p.Range '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
        With rng
            

            Rem just for test
'            If VBA.InStr(rng.text, "<a id=""ch44"">") Then
'                rng.Select
'                Stop 'check
'            End If
''            GoTo finish
            Rem just for test

            
            
            With .Find
                .ClearFormatting
                Rem ���B�z https://sl.bing.net/fQ5lVr8PLye
                .text = "<table"
                If .Execute() Then
                    Set rngClose = rng.Document.Range(rng.End, rngHtml.End)
                    With rngClose.Find
                        .text = "</table>"
                        .Execute
                    End With
                    Set rng = rngHtml.Document.Range(rng.start, rngClose.End)
                    insertHTMLTable rng, domainUrlPrefix, fontName
                    'rng.text = vbNullString
                    GoTo nextP
                End If
                .text = "<b>"
                Do While .Execute()
                    Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                    rngClose.Find.ClearFormatting
                    If Not rngClose.Find.Execute("</b>") Then Stop 'to check
                    rng.Document.Range(rng.End, rngClose.start).font.Bold = True
                    rng.text = vbNullString: rngClose.text = vbNullString
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                Loop
                .text = "<b "
                Do While .Execute()
                    rng.MoveEndUntil ">"
                    rng.End = rng.End + 1
                    Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                    rngClose.Find.ClearFormatting
                    If Not rngClose.Find.Execute("</b>") Then Stop 'to check
                    rng.Document.Range(rng.End, rngClose.start).font.Bold = True
                    rng.text = vbNullString: rngClose.text = vbNullString
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                Loop
                .text = "<span lang=""EN-US"">"
                Do While .Execute()
                    Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                    rngClose.Find.ClearFormatting
                    If Not rngClose.Find.Execute("</span>") Then Stop 'to check
                    rng.Document.Range(rng.End, rngClose.start).font.Name = "Calibri"
                    rng.text = vbNullString: rngClose.text = vbNullString
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                Loop
                .text = "<st1:chmetcnv "
                Do While .Execute()
                    rng.MoveEndUntil ">"
                    rng.End = rng.End + 1
                    Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                    rngClose.Find.ClearFormatting
                    If Not rngClose.Find.Execute("</st1:chmetcnv>") Then Stop 'to check
                    rng.text = vbNullString: rngClose.text = vbNullString
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                Loop
                .text = "<span class=""Apple-style-span"""
                Do While .Execute()
                    rng.MoveEndUntil ">"
                    rng.End = rng.End + 1
                    Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                    rngClose.Find.ClearFormatting
                    If Not rngClose.Find.Execute("</span>") Then Stop 'to check
                    rng.text = vbNullString: rngClose.text = vbNullString
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                Loop
                .text = "<span id="""
                Do While .Execute()
                    rng.MoveEndUntil ">"
                    rng.End = rng.End + 1
                    If VBA.InStr(rng.text, "data-llen=") Then
                        Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                        rngClose.Find.ClearFormatting
                        If Not rngClose.Find.Execute("</span>") Then Stop 'to check
                        rng.text = vbNullString: rngClose.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                    Else
                        rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                        Exit Do
                    End If
                Loop
                .text = "<small>"
                Do While .Execute()
                    Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                    rngClose.Find.ClearFormatting
                    If Not rngClose.Find.Execute("</small>") Then Stop 'to check
                    rng.text = vbNullString: rngClose.text = vbNullString
                    rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size * 0.84
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                Loop
                .text = "<blockquote>"
                If .Execute() Then
                    Set rngClose = rng.Document.Range(rng.End, p.Next.Range.End)
                    rngClose.Find.ClearFormatting
                    If Not rngClose.Find.Execute("</blockquote>") Then
                        rngClose.SetRange rngClose.End, rngHtml.End
                        If Not rngClose.Find.Execute("</blockquote>") Then
                            rng.Select
                            playSound 12
                            Stop 'to check
                        End If
                    End If
                    rng.text = vbNullString
                    If rngClose.Paragraphs(1).Range.text = "</blockquote>" & Chr(13) Then
                        rngClose.Paragraphs(1).Range.text = vbNullString
                    Else
                        Stop 'for check
                        rngClose.text = vbNullString
                    End If
                    
                    rng.Document.Range(rng.End, rngClose.start).ParagraphFormat.CharacterUnitLeftIndent = 3
                    rng.Document.Range(rng.End, rngClose.start).font.Name = "�з���"
                    'GoTo nextP
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                End If
                .text = "<h2>"
                If .Execute() Then
                    Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                    If rngClose.Find.Execute("</h2>") = False Then
                        rng.Select
                        playSound 12
                        Stop
                    End If
                    rng.Style = wdStyleHeading2
                    With rng.Document.Range(rng.End, rngClose.start)
                        .text = VBA.Replace(VBA.Replace(.text, "<strong>", vbNullString), "</strong>", vbNullString)
                    End With
                    rngClose.Find.Execute "</h2>"
                    rng.text = vbNullString: rngClose.text = vbNullString
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                End If
                .text = "<hr>"
                If .Execute() Then
                    rng.text = vbNullString
                    'rng.InsertBreak WdBreakType.wdLineBreak
                    With p.Borders(wdBorderBottom)
                        .LineStyle = wdLineStyleSingle ' ���J��u  ���J���u�GwdLineStyleDouble ���J��u�GwdLineStyleDot
                        .LineWidth = wdLineWidth050pt
                        .Color = wdColorAutomatic
                    End With
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                End If
                .text = "<hr " 'ex: <hr style="padding-left: 30px;">
                If .Execute() Then
                    rng.MoveEndUntil ">"
                    rng.End = rng.End + 1
                    '�ɥ� url �ܼ�
                    url = rng.text
                    rng.text = vbNullString
                    'rng.InsertBreak WdBreakType.wdLineBreak
                    With p.Borders(wdBorderBottom)
                        .LineStyle = wdLineStyleSingle ' ���J��u  ���J���u�GwdLineStyleDouble ���J��u�GwdLineStyleDot
                        .LineWidth = wdLineWidth050pt
                        .Color = wdColorAutomatic
                    End With
                    url = getHTML_AttributeValue("style", url)
                    arr = VBA.Split(url, ";")
                    For Each e In arr
                        If e <> vbNullString Then
                            e = VBA.Trim(e)
                            l = VBA.Len("padding-left: ")
                            If VBA.Left(e, l) = "padding-left: " Then
                                rng.ParagraphFormat.LeftIndent = PixelsToPoints(VBA.Replace(VBA.Mid(e, l + 1), "px", vbNullString))
                            Else
                                playSound 12
                                Stop 'for check
                            End If
                        End If
                    Next e

                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                End If
                
                '�B�z�Ϥ�
                .text = "<img "
                Do While .Execute()
                    rng.MoveEndUntil ">" 'ex: <img style="float:right;margin-left:15px;margin-right:15px;" src="/image/3.jpg" width="200" height="297"
                    '�ɥ��ܼ�
                    url = rng.text
                    rng.End = rng.End + 1 '�]�t ">"
                    rng.text = vbNullString
                    'pCntr + VBA.Abs(10 - pCntr) '�U���Ϥ��ݭn�ɶ�
                    If Not insertHTMLImage(url, rng, domainUrlPrefix) Is Nothing Then
                        'p.Range.ParagraphFormat.BaseLineAlignment = wdBaselineAlignCenter
                    End If
'                    If rng.Paragraphs(1).Range.ShapeRange.Count > 0 Then
'                        Stop
'                    End If
                    
                    rng.SetRange rng.End, p.Range.End
                    
                Loop '�B�z�Ϥ�
                
                If VBA.Len(p.Range.text) > VBA.Len("<strong></strong>") Then
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                    .text = "<strong>"
                    Do While .Execute()
                        Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                        rngClose.Find.text = "</strong>"
                        If Not rngClose.Find.Execute() Then
                            rngClose.SetRange rngClose.End, rngClose.Paragraphs(1).Next.Range.End
                            If Not rngClose.Find.Execute() Then Stop 'for check
                        End If
                        rng.Document.Range(rng.End, rngClose.start).font.Bold = True
                        rng.text = vbNullString: rngClose.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                If VBA.Len(p.Range.text) > VBA.Len("<strong style=""; ;""></strong>") Then
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                    .text = "<strong style="""
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                        rngClose.Find.text = "</strong>"
                        rngClose.Find.Execute
                        rng.Document.Range(rng.End, rngClose.start).font.Bold = True
                        rng.text = vbNullString: rngClose.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                '�B�z�r���˦�
                If VBA.Len(p.Range.text) > VBA.Len("<span style=""></span>") Then
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
'                    .text = "<span style"
                    Do While .Execute("<span style")
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                        rngClose.Find.text = "</span>"
                        rngClose.Find.Execute
                        Do While VBA.InStr(rng.Document.Range(rng.End, rngClose.start).text, "<span>")
                            rng.SetRange rng.End, p.Range.End
                            rng.Find.Execute "<span>"
                            rngClose.SetRange rng.End, p.Range.End
                            rngClose.Find.Execute "</span>"
                            If rngClose.End + 1 = p.Range.End Or VBA.InStr(rngClose.Document.Range(rngClose.End, p.Range.End), "</span>") = 0 Then
                                Exit Do
                            End If
                        Loop
                        
                        '�ɥ�url�ܼ�
                        'url = VBA.Replace(getHTML_AttributeValue("span style", rng.text), "font-family:", vbNullString)
                        url = getHTML_AttributeValue("style", rng.text)
                        If url <> "" Then
                            If VBA.Right(url, 1) = ";" Then
                                url = VBA.Left(url, VBA.Len(url) - 1)
                            End If
                            
                            '��@�ݩʭȡGhtml�ݩʭȤ��S�������];�^��
                            If VBA.InStr(url, ";") = 0 Then
                                Select Case url
                                    Case "font-size: x-large", "font-size:x-large"
                                        rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size * 1.3
                                    Case "font-size: large", "font-size:large"
                                        rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size * 1.2
                                    Case "font-size: medium", "font-size:medium" '���B�z
                                    Case "font-size: small", "font-size:small"
                                        rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size * (5 / 6)
                                    Case "font-size: x-small", "font-size:x-small"
                                        rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size * (2 / 3)
                                    Case "font-size: xx-small", "font-size:xx-small"
                                        rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size * (1 / 2)
                                    Case "text-decoration:underline"
                                        rng.Document.Range(rng.End, rngClose.start).font.Underline = wdUnderlineSingle
                                    Case Else
                                        If VBA.Left(url, 6) = "color:" Then
                                            url = VBA.LTrim(VBA.Mid(url, VBA.Len("color:") + 1))
                                            If VBA.Left(url, 1) = "#" Then
                                                arr1 = HTML2Doc.ColorCodetoRGB(url)
                                                rng.Document.Range(rng.End, rngClose.start).font.Color = VBA.RGB(arr1(0), arr1(1), arr1(2))
                                            Else
                                                If url = "red" Then
                                                    rng.Document.Range(rng.End, rngClose.start).font.ColorIndex = wdRed
                                                Else
                                                    playSound 12
                                                    rng.Select
                                                    Debug.Print url
                                                    Stop 'for check
                                                End If
                                            End If
                                        ElseIf VBA.InStr(url, ";") = 0 And VBA.InStr(url, "; ") = 0 And VBA.InStr(url, "font-size:") <> 1 And VBA.InStr(url, "line-height:") = 0 And VBA.InStr(url, "font-family") = 0 And VBA.InStr(url, "Mso") = 0 And VBA.InStr(url, "mso-") = 0 And VBA.InStr(url, "�з���") = 0 And VBA.InStr(url, "letter-spacing:0pt") = 0 And VBA.InStr(url, "�s�ө���") = 0 And VBA.InStr(url, "background-color: ") = 0 And VBA.InStr(url, "color: ") = 0 Then
                                            playSound 12
                                            rng.Select
                                            Debug.Print url
                                            Stop 'for check
                                        ElseIf VBA.InStr(url, "font-family:") = 1 Or VBA.InStr(url, "fontname=") Then
                                            Select Case VBA.Trim(Mid(url, VBA.Len("font-family:") + 1))
                                                Case "�s�ө���", "Verdana, Arial, Helvetica, sans-serif"
                                                    '���B�z
                                                Case "�з���"
                                                    If Fonts.IsFontInstalled("�з���") Then
                                                        If rng.Document.Range(rng.End, rngClose.start).font.Name <> "�з���" Then
                                                            rng.Document.Range(rng.End, rngClose.start).font.Name = "�з���"
                                                        End If
                                                    End If
                                                Case Else
                                                    playSound 12
                                                    rng.Select
                                                    Debug.Print url
                                                    Stop 'for check
                                                    'url = "�з���"
    '                                                If Fonts.IsFontInstalled(VBA.Trim(url)) Then
    '                                                    If rng.Document.Range(rng.End, rngClose.start).font.Name <> VBA.Trim(url) Then
    '                                                        rng.Document.Range(rng.End, rngClose.start).font.Name = VBA.Trim(url)
    '                                                    End If
    '                                                End If
                                            End Select
                                        ElseIf VBA.InStr(url, "font-size:") = 1 Then
                                            l = VBA.Len("font-size:")
                                            If VBA.Right(url, 2) = "em" Then ' em �O�@�Ӭ۹���A�Ω�]�m�r��j�p�C���۹����������r��j�p�C�Ҧp�A�p�G���������r��j�p�O16�����A�h 1em ����16�����A1.5em ����24�����C20241011 https://sl.bing.net/bVzA9JEh8VM
                                                l = VBA.Len("font-size:")
                                                rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size _
                                                        * VBA.CSng(VBA.Trim(VBA.Mid(url, l + 1, VBA.Len(url) - l - VBA.Len("em"))))
                                            ElseIf VBA.IsNumeric(VBA.Mid(url, l + 1)) Then
                                                rng.Document.Range(rng.End, rngClose.start).font.Size = VBA.CSng(IIf(VBA.Mid(url, l + 1) < 1, VBA.Mid(url, l + 1) * 10, VBA.Mid(url, l + 1)))
                                            Else
                                                playSound 12
                                                Debug.Print url
                                                rng.Select 'for check
                                                Stop
    '                                            If VBA.InStr(url, ";") Then
    '                                                If Not IsEmpty(arr) Then
    '                                                    If IsArray(arr) Then
    '                                                        ' �� ReDim �M�Ű}�C
    '                                                        If VBA.IsArray(arr) Then
    '                                                            ' �� ReDim �M�Ű}�C
    '                                                            ReDim arr(0)
    '                                                        End If
    '                                                    End If
    '                                                    arr = Empty
    '                                                End If
    '                                                arr = VBA.Split(url, ";")
    '                                                For Each e In arr
    '                                                    e = VBA.Trim(e)
    '                                                    If VBA.Left(e, 10) = "font-size:" Then
    '                                                        If VBA.IsNumeric(VBA.Replace(VBA.Trim(VBA.Mid(e, 11)), "px", vbNullString)) Then
    '                                                            rng.font.Size = VBA.CSng(VBA.Replace(VBA.Trim(VBA.Mid(e, 11)), "px", vbNullString))
    '                                                        Else
    '                                                            Select Case VBA.Replace(VBA.Trim(VBA.Mid(e, 11)), "px", vbNullString)
    '                                                                Case "x-small"
    '                                                                    rng.font.Size = rng.font.Size * (2 / 3)
    '                                                                Case "medium"
    '                                                                '���B�z
    '                                                                Case Else
    '                                                                    playSound 12
    '                                                                    rng.Select
    '                                                                    Debug.Print e
    '                                                                    Stop 'to check
    '                                                            End Select
    '                                                        End If
    '                                                Next e
    '                                            Else
'                                                    If VBA.InStr(url, "font-size: medium") = 0 And VBA.InStr(url, "font-size:medium") = 0 Then
'                                                        Stop
'                                                    End If
    '                                            End If
                                            End If 'If VBA.Right(url, 2) <> "em" and VBA.IsNumeric(VBA.Mid(url, l + 1))=false �K�K Then
                                        ElseIf VBA.Left(url, 12) = "line-height:" Then
                                            If VBA.InStr(url, "px") Then
                                                rng.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
                                                rng.ParagraphFormat.LineSpacing = VBA.CSng(VBA.Replace(VBA.Trim(VBA.Mid(url, 13)), "px", vbNullString))
                                            Else
                                                rng.ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
                                                rng.ParagraphFormat.LineSpacing = VBA.CSng(VBA.Replace(VBA.Trim(VBA.Mid(url, 13)), "px", vbNullString))
                                            End If
                                        Else 'If VBA.InStr(url, "font-size:") <> 1 Then
                                            playSound 12
                                            Debug.Print url
                                            rng.Select 'for check
                                            Stop
                                        End If 'If VBA.InStr(url, "font-size:") = 1 Then
    
                                End Select 'Select Case url
                            '�h���ݩʭȡGhtml�ݩʭȦ������];�^���j��
                            Else 'If VBA.InStr(url, "; ") Or VBA.InStr(url, ";") Then '�r���q����L�榡������
                                arr = VBA.Split(url, ";")
                                For Each e In arr
                                    e = VBA.Trim(e)
                                    If VBA.Left(e, 17) = "background-color:" Then
                                        arr1 = HTML2Doc.ColorCodetoRGB(VBA.LTrim(VBA.Mid(e, VBA.Len("background-color:") + 1)))
                                        rng.Document.Range(rng.End, rngClose.start).font.Shading.BackgroundPatternColor = VBA.RGB(arr1(0), arr1(1), arr1(2))
                                    ElseIf VBA.Left(e, 6) = "color:" Then
                                        arr1 = HTML2Doc.ColorCodetoRGB(VBA.LTrim(VBA.Mid(e, VBA.Len("color:") + 1)))
                                        rng.Document.Range(rng.End, rngClose.start).font.Color = VBA.RGB(arr1(0), arr1(1), arr1(2))
                                    ElseIf VBA.Left(e, 12) = "line-height:" Then
                                        arr1 = VBA.LTrim(VBA.Mid(e, VBA.Len("line-height:") + 1))
                                        If Not VBA.IsNumeric(arr1) Then
                                            If VBA.InStr(arr1, "px") Then
                                                arr1 = VBA.Replace(arr1, "px", vbNullString)
                                            ElseIf arr1 = "normal" Then
                                                '���B�z
                                            Else
                                                playSound 12 'for check
                                                Debug.Print e
                                                rng.Select
                                                Stop
                                            End If
                                        End If
                                        If VBA.IsNumeric(arr) Then
                                            If arr1 < 10 Then
                                                rng.Document.Range(rng.End, rngClose.start).ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
                                                rng.Document.Range(rng.End, rngClose.start).ParagraphFormat.LineSpacing = VBA.CSng(arr1)
                                            Else
                                                rng.Document.Range(rng.End, rngClose.start).ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
                                                rng.Document.Range(rng.End, rngClose.start).ParagraphFormat.LineSpacing = VBA.CSng(arr1)
                                            End If
                                        End If
                                    ElseIf VBA.Left(e, 10) = "font-size:" Then
                                        arr1 = VBA.Replace(VBA.LTrim(VBA.Mid(e, VBA.Len("font-size:") + 1)), "px", vbNullString)
                                        If VBA.IsNumeric(arr1) Then
                                            rng.Document.Range(rng.End, rngClose.start).font.Size = VBA.CSng(arr1)
                                        Else
                                            Select Case arr1
                                                Case "x-small"
                                                    rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size * (2 / 3)
                                                Case "medium"
                                                    '���B�z�A�Y�w�]�j�p
                                                Case Else
                                                    playSound 12
                                                    Debug.Print e
                                                    rng.Select
                                                    Stop 'to check
                                            End Select
                                        End If
                                    ElseIf e = "'Helvetica Neue', Helvetica, Arial, sans-serif" Then
                                        '���B�z
                                    ElseIf e = "�ө���" Then
                                    ElseIf e = "font-family: �s�ө���" Then
                                    ElseIf e = "font-family:�s�ө���" Then
                                    ElseIf e = "�L�n������, 'Helvetica Neue', Helvetica, sans-serif, �s�ө���" Then
                                    ElseIf e = "mso-ascii-font-family: 'Times New Roman'" Then
                                    ElseIf e = "mso-hansi-font-family: 'Times New Roman'" Then
                                    ElseIf e = "mso-ascii- 'Times New Roman'" Then
                                    ElseIf e = "mso-hansi- 'Times New Roman'" Then
                                    ElseIf e = "letter-spacing:0pt" Then
                                    ElseIf e = "�з���" Then
                                        rng.Document.Range(rng.End, rngClose.start).font.Name = "�з���"
                                    Else
                                        SystemSetup.playSound 12
                                        rng.Select
                                        Debug.Print e
                                        Stop 'to check
                                    End If
                                Next e
                            End If
                        End If 'If url <> "" Then
                        
                        rng.text = vbNullString: rngClose.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                If VBA.Len(p.Range.text) > VBA.Len("<font size=""""></font>") Then
                    .text = "<font size="
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                        If rngClose.Find.Execute("</font>") = False Then
                            playSound 12
                            rng.Select
                            Stop 'for check
                        End If
                        rng.Document.Range(rng.End, rngClose.start).font.Size = VBA.CSng(getHTML_AttributeValue("size", rng.text))
                        rng.text = vbNullString: rngClose.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                '�B�z�W�s��
                If VBA.InStr(p.Range.text, "<a ") Then
                    Set rngClose = rng.Document.Range(p.Range.start, p.Range.End - 1)
                    If VBA.InStr(rngClose.text, "</a>") = 0 Then
                        rngClose.SetRange p.Next.Range.start, p.Range.Document.Content.End
                        If Not rngClose.Find.Execute(findText:="</a>") Then
                            playSound 12
                            p.Range.Select
                            Stop
                        End If
                    End If
                    If VBA.InStr(rngClose.text, "</a>") = 0 Then
                        playSound 12
                        p.Range.Select
                        Stop
                    End If
                    InsertHTMLLinks p.Range.Document.Range(p.Range.start, rngClose.End), domainUrlPrefix
'                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
'                    '.text = "<a href="""
'                    .text = "<a "
'                    Do While .Execute()
'                        rng.MoveEndUntil ">"
'                        rng.End = rng.End + 1
'                        url = rng.text: e = rng.text
'                        Set rngClose = rng.Document.Range(rng.End, p.Range.End)
'                        rngClose.Find.Execute "</a>"
'                        url = getHTML_AttributeValue("href", url)
'                        'url = getHTML_AttributeValue("<a href", p.Range.text)
'                        e = getHTML_AttributeValue("title", VBA.CStr(e))
'                        Select Case VBA.Left(url, 1)
'                            Case "#"
'                                If Not SeleniumOP.IsWDInvalid() Then
'                                    url = WD.url & url
'                                End If
'                            Case "/"
'                                url = domainUrlPrefix & url '���|���h�@�ӱ׽u�]/�^�]�O�i�H���A�S�t 20241012
'                            Case vbNullString
'                                If rng.Document.Range(rng.End, rngClose.start).InlineShapes.Count > 0 Then '�᭱�|�ˬd�Grng.Document.Range(rng.End, rngClose.start).ShapeRange.Count > 0
'                                    playSound 12
'                                    Stop 'for check
'                                End If
'                                '�Ū��W�s���A���B�z�A�����M��
'                            Case Else
'                                If Not VBA.Left(url, 4) = "http" Then
'                                    Stop 'check
'                                    url = domainUrlPrefix & url
'                                End If
'                        End Select
'
'                        Set obj = rng.Document.Range(rng.start, rngClose.End).ShapeRange
'                        rng.text = vbNullString: rngClose.text = vbNullString
'                        If Not obj Is Nothing Then
'                            Select Case obj.Count
'                                Case 0
'                                    If rng.Document.Range(rng.End, rngClose.start).text <> vbNullString Then
'                                        rng.Document.Range(rng.End, rngClose.start).Hyperlinks.Add rng.Document.Range(rng.End, rngClose.start), url, , e
'                                    End If
'                                Case 1
'                                    rng.Document.Range(rng.End, rngClose.start).Hyperlinks.Add obj(1), url, , e
'                                Case Else
'                                    playSound 12 'for check
'                                    Stop
'                            End Select
'
'                            Set obj = Nothing
'                        Else
'                            playSound 12 'for check
'                            Stop
'                        End If
'                        rng.SetRange rngClose.End, p.Range.End
'                    Loop
                End If '�B�z�W�s��

'                If VBA.Len(p.Range.text) > VBA.Len("<p style=""padding-left:;>") Then
'                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
'                    .text = "<p style=""padding-left:"
'                    Do While .Execute()
'                        rng.MoveEndUntil ">"
'                        rng.End = rng.End + 1
'                        p.Range.ParagraphFormat.IndentCharWidth 3
'                        rng.text = vbNullString
'                        rng.SetRange p.Range.start, p.Range.End
'                    Loop
'                End If
                If VBA.Len(p.Range.text) > VBA.Len("<span size=""></span>") Then
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                    .text = "<span size="
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                        rngClose.Find.text = "</span>"
                        rngClose.Find.Execute
                        
                        rng.text = vbNullString: rngClose.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                If VBA.Len(p.Range.text) > VBA.Len("<span color=""></span>") Then
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                    .text = "<span color="
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                        rngClose.Find.text = "</span>"
                        rngClose.Find.Execute
                        e = rng.text
                        e = getHTML_AttributeValue("color", VBA.CStr(e))
                        arr = HTML2Doc.ColorCodetoRGB(VBA.CStr(e))
                        rng.Document.Range(rng.End, rngClose.start).font.Color = VBA.RGB(arr(0), arr(1), arr(2))
                        rng.text = vbNullString: rngClose.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                If VBA.Len(p.Range.text) > VBA.Len("<span></span>") Then
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                    .text = "<span>"
                    Do While .Execute()
                        Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                        rngClose.Find.text = "</span>"
                        rngClose.Find.Execute
                        rng.text = vbNullString: rngClose.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                If VBA.Len(p.Range.text) > VBA.Len("<center>") Then
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                    .text = "<center>"
                    If .Execute() Then
                        Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                        If VBA.InStr(rngClose.text, "</center") = 0 Then
                            rngClose.SetRange rngClose.End, rngClose.Document.Range.End - 1
                        End If
                        If Not rngClose.Find.Execute(findText:="</center>") Then
                            rngClose.Select
                            playSound 12
                            Stop 'for check
                        End If
                        rng.Document.Range(rng.End, rngClose.start).ParagraphFormat.Alignment = wdAlignParagraphCenter
                        rng.text = vbNullString: rngClose.text = vbNullString
                        If rngClose.Paragraphs(1).Range.text = Chr(13) Then rngClose.Paragraphs(1).Range.text = vbNullString
                        If rng.Paragraphs(1).Range.text = Chr(13) Then rng.Paragraphs(1).Range.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    End If
                End If
                
                If VBA.Len(p.Range.text) > VBA.Len("<p id=") Then
                    .text = "<p id="
                    rng.SetRange p.Range.start, p.Range.End
                    If .Execute() Then
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        'rng.Select
                        rng.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                    End If
                End If
                If VBA.InStr(p.Range.text, "<p style=") Then
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                    '.text = " style=""line-height: "
                    .text = "<p style="""
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        '�ɥ�url�ܼ�
                        url = getHTML_AttributeValue("style", p.Range.text)
                        arr = VBA.Split(url, ";")
                        For Each e In arr
                            e = VBA.Trim(e)
                            If VBA.Left(e, 12) = "line-height:" Then
                                rng.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
                                rng.ParagraphFormat.LineSpacing = CSng(VBA.Replace(VBA.LTrim(VBA.Mid(e, VBA.Len("line-height:") + 1)), "px", vbNullString))
                            ElseIf VBA.Left(e, 10) = "font-size:" Then
                                rng.Paragraphs(1).Range.font.Size = VBA.CSng(VBA.Replace(VBA.LTrim(VBA.Mid(e, VBA.Len("font-size:") + 1)), "px", vbNullString))
                            ElseIf VBA.Left(e, 11) = "margin-top:" Then
                                '���B�z
                            ElseIf VBA.Left(e, 12) = "margin-left:" Then
                            ElseIf VBA.Left(e, 13) = "margin-right:" Then
                            ElseIf VBA.Left(e, 6) = "color:" Then
                                arr1 = VBA.LTrim(VBA.Mid(e, VBA.Len("color:") + 1))
                                If VBA.Left(arr1, 1) = "#" Then
                                    'arr1 = HTML2Doc.ColorCodetoRGB(url)
                                    p.Range.font.Color = RGBFormColorCode(VBA.CStr(arr1))  'VBA.RGB(arr1(0), arr1(1), arr1(2))
                                Else
                                    If arr1 = "red" Then
                                        p.Range.font.ColorIndex = wdRed
                                    Else
                                        playSound 12
                                        rng.Select
                                        Debug.Print e
                                        Stop 'for check
                                    End If
                                End If
                            ElseIf VBA.Left(e, 17) = "background-color:" Then
                                arr1 = VBA.LTrim(VBA.Mid(e, VBA.Len("background-color:") + 1))
                                If VBA.Left(arr1, 1) = "#" Then
                                    '20241018 Copilot�j���ġG �]�m�q���I���� (����HTML���� background-color: #ffffff)
                                    p.Range.ParagraphFormat.Shading.BackgroundPatternColor = RGBFormColorCode(VBA.CStr(arr1))
                                Else
                                    If arr1 = "red" Then
                                        p.Range.ParagraphFormat.Shading.BackgroundPatternColorIndex = wdRed
                                    Else
                                        playSound 12
                                        rng.Select
                                        Debug.Print e
                                        Stop 'for check
                                    End If
                                End If
                            ElseIf VBA.Left(e, 11) = "text-align:" Then
                                l = VBA.Len("text-align:") + 1
                                If VBA.LTrim(VBA.Mid(e, l)) = "right" Then
                                    rng.ParagraphFormat.Alignment = wdAlignParagraphRight
                                ElseIf VBA.LTrim(VBA.Mid(e, l)) = "left" Then
                                    rng.ParagraphFormat.Alignment = wdAlignParagraphLeft
                                ElseIf VBA.LTrim(VBA.Mid(e, l)) = "center" Then
                                    rng.ParagraphFormat.Alignment = wdAlignRowCenter
                                Else
                                    playSound 12
                                    rng.Select
                                    Debug.Print e
                                    Stop 'for check
                                End If
                            ElseIf VBA.Left(e, 13) = "padding-left:" Then
                                arr1 = Empty
                                arr1 = getHTML_AttributeValue("padding-left", rng.text, ":")
                                If VBA.Right(arr1, 1) = ";" Then arr1 = VBA.Left(arr1, VBA.Len(arr1) - 1)
                                If VBA.Right(VBA.Trim(VBA.Left(arr1, 4)), 2) = "px" Then
                                    'If VBA.Trim(VBA.Left(arr1, 4)) = "30px" Then
                                    If VBA.IsNumeric(VBA.Val(VBA.Trim(VBA.Replace(arr1, "px", vbNullString)))) Then
                                        rng.ParagraphFormat.IndentCharWidth 3 * (VBA.CSng(VBA.Val(VBA.Trim(VBA.Replace(arr1, "px", vbNullString)))) / 30)
                                    Else
                                        playSound 12
                                        rng.Select
                                        Stop 'for check
                                    End If
                                Else
                                    playSound 12
                                    rng.Select
                                    Stop 'for check
                                End If
                            Else
                                If e <> vbNullString Then
                                    playSound 12
                                    rng.Select
                                    Debug.Print e
                                    Stop 'for check
                                End If
                            End If
                        Next e
                        'url = VBA.Replace(VBA.Replace(url, "line-height: ", vbNullString), "px;", vbNullString)
                        'rng.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
                        'rng.ParagraphFormat.LineSpacing = VBA.CSng(url)
                        rng.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                If VBA.Len(p.Range.text) > VBA.Len("<p dir="""">") Then
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                    .text = "<p dir="""
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        '�ɥ�url�ܼ�
                        If VBA.InStr(rng.text, "ltr") Then rng.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                If VBA.Len(p.Range.text) > VBA.Len("�mp align"""">") Then
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                    .text = "<p align="""
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        '�ɥ�url�ܼ�
                        url = getHTML_AttributeValue("align", rng.text)
                        If url = "left" Then
                            rng.ParagraphFormat.Alignment = wdAlignParagraphLeft
                        ElseIf url = "right" Then
                            rng.ParagraphFormat.Alignment = wdAlignParagraphRight
                        ElseIf url = "center" Then
                            rng.ParagraphFormat.Alignment = wdAlignParagraphCenter
                        Else
                            playSound 12 'for check
                            Debug.Print url
                            rng.Select
                            Stop
                        End If
                        rng.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                If VBA.Len(p.Range.text) > VBA.Len("<p class="";"">") Then
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                    .text = "<p class="""
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        rng.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                If VBA.Len(p.Range.text) > VBA.Len("<st1:personname ></st1:personname>") Then
                    rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                    .text = "<st1:personname "
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                        rngClose.Find.ClearFormatting
                        rngClose.Find.Execute "</st1:personname>"
                        rng.text = vbNullString: rngClose.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                'rng.SetRange p.Range.start, p.Range.End '��set �|�k�s�A�� setRange ���|�A�u�O�վ�
                
            End With 'rng.Find
            
            
            If .Paragraphs(1).Range.text = "<br class=""Apple-interchange-newline""> " & Chr(13) Then
                .Paragraphs(1).Range.text = vbNullString
            End If
        End With 'rng
nextP:
    Next p
    
    Set rng = rngHtml.Document.Range(rngHtml.start, rngHtml.End)
    ��r�B�z.FixFontname rng

    
finish:
  
    rngHtml.Document.Range(stRngHTML, stRngHTML).Select '�^��_�l��m
    
End Sub
