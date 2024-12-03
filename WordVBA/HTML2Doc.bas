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
Rem '處理超連結 20241016
Private Sub InsertHTMLLinks(rngHtml As Range, Optional domainUrlPrefix As String)
    Dim e As Variant '作為通用一般變數
    Dim obj As Object '作為通用物件變數
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
                        url = domainUrlPrefix & url '路徑中多一個斜線（/）也是可以的，沒差 20241012
                    Case vbNullString
                        If rng.Document.Range(rng.End, rngClose.start).InlineShapes.Count > 0 Then '後面會檢查：rng.Document.Range(rng.End, rngClose.start).ShapeRange.Count > 0
                            playSound 12
                            Stop 'for check
                        End If
                        '空的超連結，不處理，直接清除
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
    End If '處理超連結
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
    htmlDoc.body.innerHtml = fileContent
    
    ' 獲取所有錨標籤
    Set htmlElements = htmlDoc.getElementsByTagName("a")

    ' 循環遍歷所有錨元素並打印 href 屬性
    For Each htmlElement In htmlElements
        Debug.Print htmlElement.GetAttribute("href")
    Next htmlElement
End Sub

Rem 顏色碼轉換成RGB
Function ColorCodetoRGB(colorCode As String) As Long()
    ' 將bgcolor轉換為RGB顏色
    'Dim r As Integer, g As Integer, b As Integer
    If VBA.InStr(colorCode, " ") Then colorCode = VBA.Trim(colorCode)
    If VBA.Left(colorCode, 1) <> "#" Then Exit Function
    Dim arr(2) As Long
    arr(0) = CLng("&H" & Mid(colorCode, 2, 2))
    arr(1) = CLng("&H" & Mid(colorCode, 4, 2))
    arr(2) = CLng("&H" & Mid(colorCode, 6, 2))
    ColorCodetoRGB = arr
End Function

Rem 顏色碼轉換成RGB
Function RGBFormColorCode(colorCode As String) As Long
    ' 將bgcolor轉換為RGB顏色
    'Dim r As Integer, g As Integer, b As Integer
    If VBA.InStr(colorCode, " ") Then colorCode = VBA.Trim(colorCode)
    If VBA.Left(colorCode, 1) <> "#" Then Exit Function
    Dim arr(2) As Long
    arr(0) = CLng("&H" & Mid(colorCode, 2, 2))
    arr(1) = CLng("&H" & Mid(colorCode, 4, 2))
    arr(2) = CLng("&H" & Mid(colorCode, 6, 2))
    RGBFormColorCode = VBA.RGB(arr(0), arr(1), arr(2))
End Function

Rem 上標格式 20241012 creedit_with_Copilot大菩薩：
Sub ConvertHTMLSupToWordSup(rng As Range)
    ConvertHTMLTagToWord rng, TagNameHTML.Sup
End Sub
Rem 下標格式
Sub ConvertHTMLSubToWordSub(rng As Range)
    ConvertHTMLTagToWord rng, TagNameHTML.Subscript
End Sub
Private Sub ConvertHTMLTagToWord(rng As Range, ByVal tagname As TagNameHTML)
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
            .Execute findText:="<" & tag & ">", Replace:=wdReplaceAll, replaceWith:=vbNullString
            .Execute findText:="</" & tag & ">", Replace:=wdReplaceAll, replaceWith:=vbNullString
            .MatchWildcards = False
        End With
            
    End With
End Sub
'Rem 20241011 HTML 表格處理.Porc=Porcess
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
Rem 20241011 HTML 無序清單的處理.Porc=Porcess
Private Sub unorderedListPorc_HTML2Word(rngHtml As Range)
    Rem 無序清單的處理
    Dim rngUnorderedList As Range, st As Long, ed As Long, rngUnorderedListSub As Range, p As Paragraph
    If VBA.InStr(rngHtml.text, "<ul") Then
        Do
            Set rngUnorderedList = getRangeFromULToUL_UnorderedListRange(rngHtml)
            If Not rngUnorderedList Is Nothing Then
                st = rngUnorderedList.start
                Set p = rngUnorderedList.Paragraphs(1).Previous
                If Not p Is Nothing Then
                    '如果是易學網的「歷代注本：」
                    If VBA.InStr(p.Range.text, "歷代注本：") Then
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
                            '.Hyperlinks.Add rngLink, iwe.GetAttribute("href")'在前面已經插入超連結了
                            .Style = wdStyleHeading2 '標題 2
                            .font.Size = 18
                            .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle '單行間距
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
Rem 判斷是否是文繞圖 creedit_with_Copilot大菩薩 20241016
Function IsWrapTextImage(imgTag As String) As Boolean
    Dim startPos As Long, endPos As Long, tagContent As String
    
    ' 找到 <img 的起始位置
    startPos = InStr(imgTag, "<img")
    If startPos = 0 Then
        IsWrapTextImage = False
        Exit Function
    End If
    
    ' 找到 > 的結束位置
    endPos = InStr(startPos, imgTag, ">")
    If endPos = 0 Then endPos = VBA.Len(imgTag)
'    If endPos = 0 Then
'        IsWrapTextImage = False
'        Exit Function
'    End If
    
    ' 提取 <img> 標籤內容
    tagContent = Mid(imgTag, startPos, endPos - startPos + 1)
    
    ' 檢查標籤內容是否包含文繞圖的 CSS 樣式
    If InStr(tagContent, "float:") > 0 Or InStr(tagContent, "class=") > 0 Then '_
'        Or (InStr(tagContent, "vertical-align:") > 0 _
'            And InStr(tagContent, "vertical-align: bottom") = 0 _
'            And VBA.InStr(tagContent, "vertical-align: baseline") = 0 _
'            And InStr(tagContent, "vertical-align:bottom") = 0 _
'            And VBA.InStr(tagContent, "vertical-align:baseline") = 0) Then

'                   vertical-align對應的應該是：inlsp.Range.ParagraphFormat.BaseLineAlignment 屬性！20241016
    
        IsWrapTextImage = True
    Else
        IsWrapTextImage = False
    End If
End Function

Rem 將HTML文本置換成圖片，成功則傳回一個有效了 InlineShape物件 20241011 textPart:要解析的HTML文本，rng：要插入圖片的位置；domainUrlPrefix 是否圖片網址要加域名前綴
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
'            'msgbox "須帶入網域前綴才行"
'            'If domainUrlPrefix = vbNullString Then domainUrlPrefix = "https://www.eee-learning.com"
'
'            'If Not SeleniumOP.IsWDInvalid() Then
'                'domainUrlPrefix = getDomainUrlPrefix(SeleniumOP.WD.url)
'            'End If
'
'        End If
        If Not IsBase64Image(url) Then 'base64編碼的圖片
            url = domainUrlPrefix & url '路徑中多一個斜線（/）也是可以的，沒差 20241012
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
                    Rem 先插表格並處理其中的圖片，應該預設就是表格大小
                Else
                    If Not IsWrapTextImage(textPart) Then
                        resizePicture rng, inlsp, url
                    End If
                End If
            Else
                Exit Function
            End If
        End If
    Else 'base64編碼的圖片
        
        ' 插入base64編碼的圖片
        Set inlsp = InsertBase64Image(url, "tempImage.png", rng)
        If Not IsWrapTextImage(textPart) Then
            resizePicture rng, inlsp, url
        End If
        
    End If
    
    Rem 設定圖片格式
    Rem inlineShape格式

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
    '                shp.Left = ActiveDocument.PageSetup.PageWidth - shp.width - CentimetersToPoints(1) ' 設定右邊距離
    '                shp.Top = CentimetersToPoints(1) ' 設定上邊距離
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
    Rem Shape文繞圖格式
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
                ' 設置圖片的文繞圖方式和對齊方式
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
                            '.WrapFormat.Side = WdWrapSideType.wdWrapRight ' 對應於float:right
                        Case Else
                            Stop ' check
                    End Select
                    If marginLeft <> 0 Then
                        .WrapFormat.DistanceLeft = marginLeft ' 對應於margin-left:10px
                    End If
                    If marginRight <> 0 Then
                        .WrapFormat.DistanceRight = marginRight ' 對應於margin-right:10px
                    End If
                End With
            End If 'If float <> "" Or VBA.IsEmpty(marginLeft) = False Or VBA.IsEmpty(marginRight) = False Then
            
        End If 'inlsp.Range.tables.Count = 0 Then

        '其他Style屬性設定
        arr = VBA.Split(imgStyle, ";")
        For Each e In arr
            If e <> vbNullString Then
                e = VBA.Trim(e)
                If VBA.Left(e, VBA.Len("border-style:")) = "border-style:" Then
                    l = VBA.Len("border-style:")
                    attr = VBA.Trim(VBA.Mid(e, l + 1))
                    Select Case attr
                        Case "initial"
                            inlsp.Line.Visible = msoFalse ' 無邊框
                        Case "solid"
                            inlsp.Line.Visible = msoTrue ' 邊框
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
                            inlsp.Line.ForeColor.RGB = RGB(0, 0, 0) ' 邊框顏色為黑色
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
                                inlsp.Line.Weight = VBA.CSng(attrSetting)   ' 邊框寬度
                            Else
                                inlsp.Line.Visible = msoFalse ' 無邊框
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
                            inlsp.Range.ParagraphFormat.SpaceBefore = 0 '段前間距
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
                            inlsp.Range.ParagraphFormat.SpaceAfter = 0 '段後間距
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
                                .ForeColor.RGB = RGB(0, 0, 0) ' 邊框顏色為黑色
                                .Weight = 1 ' 邊框寬度
                                .Style = msoLineSolid ' 邊框樣式為實線
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
                    '不處理 圖片怎麼有 color屬性？
                ElseIf VBA.Left(e, VBA.Len("font-size:")) = "font-size:" Then
                    '不處理 圖片怎麼有 font size屬性？
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
                ElseIf VBA.Left(e, VBA.Len("display:")) = "display:" Then 'Copilot大菩薩：display: block; 是用來設定元素的顯示方式，在這裡圖片被設定成塊級元素（block），這樣圖片會獨占一行，類似段落。Word VBA 中並沒有直接對應的屬性，但可以通過調整段落和圖片的佈局來模擬這個效果。20241017
                    l = VBA.Len("display:")
                    attr = VBA.Trim(VBA.Mid(e, l + 1))
                    If attr = "block" Then
                        If inlsp.Range.Paragraphs(1).Range.text <> Chr(13) Then
                            inlsp.Range.InsertParagraphBefore 'display: block; – 將圖片放置在段落中，確保圖片獨占一行。
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
                    ' 設置圖片的文繞圖方式和對齊方式
                    Set shp = inlsp.ConvertToShape
                    With shp
                        .LockAspectRatio = msoTrue ' 鎖定圖片比例
                        .WrapFormat.Type = WdWrapType.wdWrapTight ' wdWrapSquare
                        Select Case float
                            Case vbNullString
                            Case "left"
                                .Left = WdShapePosition.wdShapeLeft
                                '.WrapFormat.Side = WdWrapSideType.wdWrapLeft
                            Case "right"
                                .Left = WdShapePosition.wdShapeRight
                                '.WrapFormat.Side = WdWrapSideType.wdWrapRight ' 對應於float:right
                            Case Else
                                Stop ' check
                        End Select
                    End With
                ElseIf VBA.Left(e, VBA.Len("line-height:")) = "line-height:" Then
                    l = VBA.Len("line-height:")
                    attr = VBA.Trim(VBA.Mid(e, l + 1))
                    inlsp.Range.ParagraphFormat.LineSpacingRule = wdLineSpaceAtLeast
                    inlsp.Range.ParagraphFormat.LineSpacing = word.LinesToPoints(VBA.CSng(attr)) ' 對應HTML中的 line-height

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
'                                            .ForeColor.RGB = RGB(0, 0, 0) ' 邊框顏色為黑色
'                                            .Weight = 1 ' 邊框寬度
'                                            .Style = msoLineSolid ' 邊框樣式為實線
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



Rem 取得HTML中表格的屬性值 20241011 creedit_with_Copilot大菩薩：HTML表格轉換和屬性設置：HTML表格轉換和屬性設置
Function GetHTMLAttributeValue(attributeName As String, html As String) As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    
    ' 初始化正則表達式對象
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

Rem 清除一切的html tags HTML標籤
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
    
    ' 初始化正則表達式對象
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "<li.*?>(.*?)</li>"
    
    Set matches = regex.Execute(html)
    For Each match In matches
        listItems.Add match.SubMatches(0)
    Next match
    
    Set ParseHTMLList = listItems
End Function



Rem 20241010國慶日 清除在標籤間沒有任何內容的HTML空標籤
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
                    If VBA.InStr(rng.text, "</td>") = 0 And VBA.InStr(rng.text, "<td>") = 0 Then '儲存格不能清掉！20241016
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
Rem 取得無序列表（<ul></ur>）的範圍 20241010creedit_with_Copilot大菩薩：HTML超連結轉換成Word VBA：https://sl.bing.net/bXsbFqI2cz6
Private Function getRangeFromULToUL_UnorderedListRange(rng As Range) As Range
    Dim startRange As Range
    Dim endRange As Range
    If VBA.InStr(rng.text, "<ul") Then
    ' 查找 <ul> 標籤
        Set startRange = rng.Document.Range(rng.start, rng.End)
        With startRange.Find
            .ClearFormatting
            .text = "<ul"
            If .Execute Then
                startRange.Collapse Direction:=wdCollapseStart
            End If
        End With
        
        ' 查找 </ul> 標籤
        Set endRange = rng.Document.Range(startRange.End, rng.End)
        With endRange.Find
            .ClearFormatting
            .text = "</ul>"
            If .Execute Then
                endRange.Collapse Direction:=wdCollapseEnd
            End If
        End With
        
        ' 設定範圍
        If Not (startRange.start = rng.start And endRange.End = rng.End) Then
            Set getRangeFromULToUL_UnorderedListRange = rng.Document.Range(startRange.start, endRange.End)
        End If
    End If
End Function



Rem 20241009 取得HTML中的屬性之值 pro 不包含「="」,start 搜尋的起始位置
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

Rem 插入圖片後，根據前後字型大小自動調整圖片大小 20241009 creedit_with_Copilot大菩薩：WordVBA 圖片自動調整大小：https://sl.bing.net/e1S3H59hvI4
Private Function getImageUrl(textIncludingSrc As String)
    getImageUrl = VBA.Mid(textIncludingSrc, VBA.InStr(textIncludingSrc, "src=""") + 5, _
        VBA.InStr(VBA.InStr(textIncludingSrc, "src=""") + 5, textIncludingSrc, """") - (VBA.InStr(textIncludingSrc, "src=""") + 5))
End Function
Rem 重新調整圖片大小，若無指定 width與height 則參考前後文字型大小平均值設定
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
            ' 獲取前後字型大小
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
        
            ' 計算平均字型大小
            avgFontSize = (fontSizeBefore + fontSizeAfter) / 2
        
            ' 調整圖片大小
            pic.LockAspectRatio = msoTrue
            If Not IsValidImage_LoadPicture(url) Then
                pic.height = avgFontSize
                If Not SeleniumOP.IsWDInvalid() Then
                    pic.Range.Hyperlinks.Add pic.Range, WD.url
                End If
            Else
                pic.height = avgFontSize * 2 ' 根據需要調整比例
                pic.width = pic.height * pic.width / pic.height
            End If
'        Else
'            playSound 12
'            pic.Range.Select
'            Stop
'        End If
    End If
End Sub

Rem 解析HTML並插入清單 20241011 creedit_with_Copilot大菩薩：https://sl.bing.net/gbeqh0TAks8：HTML表格轉換和屬性設置
Rem 解析HTML內容，提取清單項目，然後在Word中插入相應的清單樣式。https://sl.bing.net/bhFU3zNMSom
Sub InsertHTMLList(html As String)
    Dim doc As Document
    Dim listItems As Collection
    Dim listItem As Variant
    Dim rng As Range
    
    ' 解析HTML
    Set listItems = ParseHTMLList(html)
    
    ' 插入清單
    Set doc = ActiveDocument
    Set rng = doc.Range(start:=doc.Content.End - 1, End:=doc.Content.End - 1)
    
    ' 開始清單
    rng.ListFormat.ApplyBulletDefault
    
    ' 填充清單內容
    For Each listItem In listItems
        rng.text = StripHTMLTags(VBA.CStr(listItem))
        rng.ParagraphFormat.Alignment = wdAlignParagraphLeft
        rng.font.Name = "標楷體"
        rng.InsertParagraphAfter
        Set rng = rng.Next(wdParagraph, 1) '.Range
    Next listItem
End Sub

Rem 解析HTML內容，提取表格、行、單元格、圖片和文字 20241011 creedit_with_Copilot大菩薩：https://sl.bing.net/fQ5lVr8PLye
Private Function parseHTMLTable(html As String) As Collection
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim tables As New Collection
    Dim rows As New Collection
    Dim cells As New Collection
    Dim table, row
    
    ' 初始化正則表達式對象
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    
    ' 匹配表格
    regex.Pattern = "<table.*?>(.*?)</table>"
    Set matches = regex.Execute(html)
    For Each match In matches
        tables.Add match.SubMatches(0)
    Next match
    
    ' 匹配行/列
    regex.Pattern = "<tr.*?>(.*?)</tr>"
    For Each table In tables
        Set matches = regex.Execute(table)
        For Each match In matches
            rows.Add match.SubMatches(0)
        Next match
    Next table
    
    ' 匹配單元格
    regex.Pattern = "<td.*?>(.*?)</td>"
    For Each row In rows
        Set matches = regex.Execute(row)
        For Each match In matches
            cells.Add match.SubMatches(0)
        Next match
    Next row
    
    Set parseHTMLTable = cells
End Function

Rem 接下來，您可以在Word中創建表格並插入相應的內容 creedit_with_Copilot大菩薩 20241011
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

        Rem 剛才測試才發現，如果我在轉成表格前的文本先設定好 Range.font.Name = "Lucida Sans Unicode" 那在轉成表格後，就可以在想要是 "Lucida Sans Unicode" 字型的音標字元上設定成這個字型了。感恩感恩　讚歎讚歎　南無阿彌陀佛 所以不能在表格轉換後設定，要先在文字轉表格前先指定 阿彌陀佛
        '.font.Name = "Lucida Sans Unicode"
        If fontName <> vbNullString Then
            .font.Name = fontName
        End If
        
    End With
    ' 解析HTML
    Set cells = parseHTMLTable(html)
    
'    Set rngTbl = rngHtml.Document.Range(st, st)
    'rngHtml.text = vbNullString
    
    ' 計算欄數
    rowCount = UBound(Split(html, "<tr")) '- 1
    If rowCount = 0 Then
        Exit Sub
    Else
        colCount = cells.Count / rowCount 'UBound(Split(html, "<td")) ' - 1'creedit_with_Copilot大菩薩 20241013
    End If
    
    ' 插入表格
    Set tbl = rngHtml.tables.Add(Range:=rngHtml, NumRows:=rowCount, NumColumns:=colCount)
    
    
     ' 設置表格屬性
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
    
    ' 填充表格內容
    row = 1
    col = 1
    For Each cell In cells 'cell：html text in a cell
        If cell <> vbNullString Then
            Set c = tbl.cell(row, col)
            'c.Range.text = cell
            Set rng = c.Range.Document.Range(c.Range.start, c.Range.End - 1) '1：Chr(13) & Chr(7)（在WordVBA中這個於strat end 屬性值只作 1！）20241016
            rng.text = cell
            ' 檢查是否包含圖片
            If InStr(cell, "<img") Then
                'c.Range.text = cell & Chr(13) & Chr(7)
                ''''''Set img = insertHTMLImage(html, c.Range, domainUrlPrefix)
                'Set rng = c.Range.Document.Range(c.Range.start, c.Range.End - 2) '2=Len(Chr(13) & Chr(7))
                    '處理圖片
                    With rng.Find
                        Do While .Execute(findText:="<img ")
                            rng.MoveEndUntil ">" 'ex: <img style="float:right;margin-left:15px;margin-right:15px;" src="/image/3.jpg" width="200" height="297"
    '                        '借用變數
    '                        url = rng.text
                            rng.End = rng.End + 1 '包含 ">"
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
                                rng.SetRange rng.start + 1, rng.End '1= "/"所插入圖片的定位符字元
                                rng.text = vbNullString
                            Else
                                playSound 12
                                rng.Select
                                Stop 'for check
                            End If
                            'pCntr + VBA.Abs(10 - pCntr) '下載圖片需要時間
                            rng.SetRange rng.End, c.Range.End - 2
                        Loop '處理圖片
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
            
            '檢查是否包含超連結
            If VBA.InStr(cell, "<a ") > 0 Then
                '處理超連結 20241016
                InsertHTMLLinks c.Range, domainUrlPrefix
            End If
            '清理<……>標籤
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
                                    Case "medium", "medium" '不處理
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
                '轉置<b></b>
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
                '<p ……
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
        
        '設定下一個儲存格座標，準備移動到下一個儲存格準備填入內容20241016
        col = col + 1
        If col > tbl.Columns.Count Then
            'tbl.rows.Add
            row = row + 1
            col = 1
        End If
        
    Next cell
    
    If float <> vbNullString Then
    '    'Dim shp As Shape
    '     '將表格轉換為Shape對象
    '    Set shp = tbl.ConvertToShape
    '
    '    tbl.rows.WrapAroundText = True
    '     設置文繞圖方式
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

Rem 20241009 將HTML轉成Word文件內文。creedit_with_Copilot大菩薩：https://sl.bing.net/jij3PK59Rka
Sub innerHTML_Convert_to_WordDocumentContent(rngHtml As Range, Optional domainUrlPrefix As String, Optional fontName As String)
    If VBA.InStr(rngHtml.text, "<") = 0 Then Exit Sub
    
     SystemSetup.playSound 1
    
    Dim htmlStr As String, rng As Range, rngClose As Range, p As Paragraph, url As String, stRngHTML As Long, pCntr As Long
    Dim s As Integer '作為 InStr() 記下結果值用
    Dim l As Integer '作為 Len() 記下結果值用
    '作為通用變數用，或陣列記住用
    Dim arr, arr1, e '作為通用一般變數用，或陣列元素記住用
    Dim obj As Object '作為通用物件變數用
    
    
'
'    Rem just for check
'    SystemSetup.ClipboardPutIn rngHtml.text
'    Stop
'
'    GoTo finish 'just for test
'
    
    
    '取得網址前綴的網域值（不含尾斜線）
    If domainUrlPrefix = vbNullString Then
        If Not SeleniumOP.IsWDInvalid() Then ' domainUrlPrefix = "https://www.eee-learning.com"
            domainUrlPrefix = GetDomainUrlPrefix(WD.url)
        End If
    End If
    'SystemSetup.stopUndo ur, "InnerHTML_DocContent"
    stRngHTML = rngHtml.start
    htmlStr = rngHtml.text '記下起始位置
    
    Rem 前置整理文本
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
        '置換
        If VBA.InStr(htmlStr, "<br>" & VBA.Chr(13)) Then .Execute "<br>^p", , , , , , , , , "^l", wdReplaceAll
        If VBA.InStr(htmlStr, "<br>") Then .Execute "<br>", , , , , , , , , "^l", wdReplaceAll
        If VBA.InStr(htmlStr, "<a style=""line-height:1.5;"" href=") Then .Execute "<a style=""line-height:1.5;"" href=", , , , , , , , , "<a href=", wdReplaceAll ' " 會置換成 “
        If VBA.InStr(htmlStr, "&lt;") Then .Execute "&lt;", , , , , , , , , "＜", wdReplaceAll
        If VBA.InStr(htmlStr, "&gt;") Then .Execute "&gt;", , , , , , , , , "＞", wdReplaceAll
        '清除
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
        'If VBA.InStr(htmlStr, "<a class=""colorbox colorbox-insert-image cboxElement"" href=") Then .Execute "<a class=""colorbox colorbox-insert-image cboxElement"" href=", , , , , , , , , "<a href=", wdReplaceAll ' " 會置換成 “
        If VBA.InStr(htmlStr, " rel=""group-all""") Then .Execute " rel=""group-all""", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "<o:p></o:p>") Then .Execute "<o:p></o:p>", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "<span></span>") Then .Execute "<span></span>", , , , , , , , , vbNullString, wdReplaceAll
        Rem 原網頁蓋用諸如Word等編輯貼上，故多有殘碼、亂碼
        If VBA.InStr(htmlStr, "<o:p>") Then .Execute "<o:p>", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "</o:p>") Then .Execute "</o:p>", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "<span style=""color:#ffffff;"">ppp</span>") Then .Execute "<span style=""color:#ffffff;"">ppp</span>", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "<!--EndFragment-->") Then .Execute "<!--EndFragment-->", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, " face=""標楷體""") Then .Execute " face=""標楷體""", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "<span class=""Apple-tab-span"" style=""white-space: pre;""> </span>") Then .Execute "<span class=""Apple-tab-span"" style=""white-space: pre;""> </span>", , , , , , , , , vbNullString, wdReplaceAll
        If VBA.InStr(htmlStr, "<!--break-->") Then .Execute "<!--break-->", , , , , , , , , vbNullString, wdReplaceAll
    End With
    
    SystemSetup.playSound 1
    Set rng = rngHtml.Document.Range(rngHtml.start, rngHtml.End)
    '清除空標籤
    RemoveEmptyTags rngHtml
    
    Set rng = rngHtml.Document.Range(rngHtml.start, rngHtml.End)
    Rem 無序清單的處理
    unorderedListPorc_HTML2Word rng
    
    For Each p In rngHtml.Paragraphs
        pCntr = pCntr + 1
        If pCntr Mod 30 = 0 Then SystemSetup.playSound 1
        
        If p.Range.text = "/" & Chr(13) Or p.Range.text = Chr(13) Or p.Range.tables.Count > 0 Then
            GoTo nextP
        End If
        Set rng = p.Range '用set 會歸零，用 setRange 不會，只是調整
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
                Rem 表格處理 https://sl.bing.net/fQ5lVr8PLye
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
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
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
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                Loop
                .text = "<span lang=""EN-US"">"
                Do While .Execute()
                    Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                    rngClose.Find.ClearFormatting
                    If Not rngClose.Find.Execute("</span>") Then Stop 'to check
                    rng.Document.Range(rng.End, rngClose.start).font.Name = "Calibri"
                    rng.text = vbNullString: rngClose.text = vbNullString
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                Loop
                .text = "<st1:chmetcnv "
                Do While .Execute()
                    rng.MoveEndUntil ">"
                    rng.End = rng.End + 1
                    Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                    rngClose.Find.ClearFormatting
                    If Not rngClose.Find.Execute("</st1:chmetcnv>") Then Stop 'to check
                    rng.text = vbNullString: rngClose.text = vbNullString
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                Loop
                .text = "<span class=""Apple-style-span"""
                Do While .Execute()
                    rng.MoveEndUntil ">"
                    rng.End = rng.End + 1
                    Set rngClose = rng.Document.Range(rng.End, p.Range.End)
                    rngClose.Find.ClearFormatting
                    If Not rngClose.Find.Execute("</span>") Then Stop 'to check
                    rng.text = vbNullString: rngClose.text = vbNullString
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
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
                        rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                    Else
                        rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
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
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
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
                    rng.Document.Range(rng.End, rngClose.start).font.Name = "標楷體"
                    'GoTo nextP
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
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
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                End If
                .text = "<hr>"
                If .Execute() Then
                    rng.text = vbNullString
                    'rng.InsertBreak WdBreakType.wdLineBreak
                    With p.Borders(wdBorderBottom)
                        .LineStyle = wdLineStyleSingle ' 插入實線  插入雙線：wdLineStyleDouble 插入虛線：wdLineStyleDot
                        .LineWidth = wdLineWidth050pt
                        .Color = wdColorAutomatic
                    End With
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                End If
                .text = "<hr " 'ex: <hr style="padding-left: 30px;">
                If .Execute() Then
                    rng.MoveEndUntil ">"
                    rng.End = rng.End + 1
                    '借用 url 變數
                    url = rng.text
                    rng.text = vbNullString
                    'rng.InsertBreak WdBreakType.wdLineBreak
                    With p.Borders(wdBorderBottom)
                        .LineStyle = wdLineStyleSingle ' 插入實線  插入雙線：wdLineStyleDouble 插入虛線：wdLineStyleDot
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

                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                End If
                
                '處理圖片
                .text = "<img "
                Do While .Execute()
                    rng.MoveEndUntil ">" 'ex: <img style="float:right;margin-left:15px;margin-right:15px;" src="/image/3.jpg" width="200" height="297"
                    '借用變數
                    url = rng.text
                    rng.End = rng.End + 1 '包含 ">"
                    rng.text = vbNullString
                    'pCntr + VBA.Abs(10 - pCntr) '下載圖片需要時間
                    If Not insertHTMLImage(url, rng, domainUrlPrefix) Is Nothing Then
                        'p.Range.ParagraphFormat.BaseLineAlignment = wdBaselineAlignCenter
                    End If
'                    If rng.Paragraphs(1).Range.ShapeRange.Count > 0 Then
'                        Stop
'                    End If
                    
                    rng.SetRange rng.End, p.Range.End
                    
                Loop '處理圖片
                
                If VBA.Len(p.Range.text) > VBA.Len("<strong></strong>") Then
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
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
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
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
                '處理字型樣式
                If VBA.Len(p.Range.text) > VBA.Len("<span style=""></span>") Then
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
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
                        
                        '借用url變數
                        'url = VBA.Replace(getHTML_AttributeValue("span style", rng.text), "font-family:", vbNullString)
                        url = getHTML_AttributeValue("style", rng.text)
                        If url <> "" Then
                            If VBA.Right(url, 1) = ";" Then
                                url = VBA.Left(url, VBA.Len(url) - 1)
                            End If
                            
                            '單一屬性值：html屬性值中沒有分號（;）時
                            If VBA.InStr(url, ";") = 0 Then
                                Select Case url
                                    Case "font-size: x-large", "font-size:x-large"
                                        rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size * 1.3
                                    Case "font-size: large", "font-size:large"
                                        rng.Document.Range(rng.End, rngClose.start).font.Size = rng.Document.Range(rng.End, rngClose.start).font.Size * 1.2
                                    Case "font-size: medium", "font-size:medium" '不處理
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
                                        ElseIf VBA.InStr(url, ";") = 0 And VBA.InStr(url, "; ") = 0 And VBA.InStr(url, "font-size:") <> 1 And VBA.InStr(url, "line-height:") = 0 And VBA.InStr(url, "font-family") = 0 And VBA.InStr(url, "Mso") = 0 And VBA.InStr(url, "mso-") = 0 And VBA.InStr(url, "標楷體") = 0 And VBA.InStr(url, "letter-spacing:0pt") = 0 And VBA.InStr(url, "新細明體") = 0 And VBA.InStr(url, "background-color: ") = 0 And VBA.InStr(url, "color: ") = 0 Then
                                            playSound 12
                                            rng.Select
                                            Debug.Print url
                                            Stop 'for check
                                        ElseIf VBA.InStr(url, "font-family:") = 1 Or VBA.InStr(url, "fontname=") Then
                                            Select Case VBA.Trim(Mid(url, VBA.Len("font-family:") + 1))
                                                Case "新細明體", "Verdana, Arial, Helvetica, sans-serif"
                                                    '不處理
                                                Case "標楷體"
                                                    If Fonts.IsFontInstalled("標楷體") Then
                                                        If rng.Document.Range(rng.End, rngClose.start).font.Name <> "標楷體" Then
                                                            rng.Document.Range(rng.End, rngClose.start).font.Name = "標楷體"
                                                        End If
                                                    End If
                                                Case Else
                                                    playSound 12
                                                    rng.Select
                                                    Debug.Print url
                                                    Stop 'for check
                                                    'url = "標楷體"
    '                                                If Fonts.IsFontInstalled(VBA.Trim(url)) Then
    '                                                    If rng.Document.Range(rng.End, rngClose.start).font.Name <> VBA.Trim(url) Then
    '                                                        rng.Document.Range(rng.End, rngClose.start).font.Name = VBA.Trim(url)
    '                                                    End If
    '                                                End If
                                            End Select
                                        ElseIf VBA.InStr(url, "font-size:") = 1 Then
                                            l = VBA.Len("font-size:")
                                            If VBA.Right(url, 2) = "em" Then ' em 是一個相對單位，用於設置字體大小。它相對於父元素的字體大小。例如，如果父元素的字體大小是16像素，則 1em 等於16像素，1.5em 等於24像素。20241011 https://sl.bing.net/bVzA9JEh8VM
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
    '                                                        ' 用 ReDim 清空陣列
    '                                                        If VBA.IsArray(arr) Then
    '                                                            ' 用 ReDim 清空陣列
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
    '                                                                '不處理
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
                                            End If 'If VBA.Right(url, 2) <> "em" and VBA.IsNumeric(VBA.Mid(url, l + 1))=false …… Then
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
                            '多重屬性值：html屬性值有分號（;）間隔者
                            Else 'If VBA.InStr(url, "; ") Or VBA.InStr(url, ";") Then '字型段落其他格式化雜項
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
                                                '不處理
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
                                                    '不處理，即預設大小
                                                Case Else
                                                    playSound 12
                                                    Debug.Print e
                                                    rng.Select
                                                    Stop 'to check
                                            End Select
                                        End If
                                    ElseIf e = "'Helvetica Neue', Helvetica, Arial, sans-serif" Then
                                        '不處理
                                    ElseIf e = "細明體" Then
                                    ElseIf e = "font-family: 新細明體" Then
                                    ElseIf e = "font-family:新細明體" Then
                                    ElseIf e = "微軟正黑體, 'Helvetica Neue', Helvetica, sans-serif, 新細明體" Then
                                    ElseIf e = "mso-ascii-font-family: 'Times New Roman'" Then
                                    ElseIf e = "mso-hansi-font-family: 'Times New Roman'" Then
                                    ElseIf e = "mso-ascii- 'Times New Roman'" Then
                                    ElseIf e = "mso-hansi- 'Times New Roman'" Then
                                    ElseIf e = "letter-spacing:0pt" Then
                                    ElseIf e = "標楷體" Then
                                        rng.Document.Range(rng.End, rngClose.start).font.Name = "標楷體"
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
                '處理超連結
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
'                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
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
'                                url = domainUrlPrefix & url '路徑中多一個斜線（/）也是可以的，沒差 20241012
'                            Case vbNullString
'                                If rng.Document.Range(rng.End, rngClose.start).InlineShapes.Count > 0 Then '後面會檢查：rng.Document.Range(rng.End, rngClose.start).ShapeRange.Count > 0
'                                    playSound 12
'                                    Stop 'for check
'                                End If
'                                '空的超連結，不處理，直接清除
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
                End If '處理超連結

'                If VBA.Len(p.Range.text) > VBA.Len("<p style=""padding-left:;>") Then
'                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
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
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
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
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
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
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
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
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
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
                        rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                    End If
                End If
                If VBA.InStr(p.Range.text, "<p style=") Then
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                    '.text = " style=""line-height: "
                    .text = "<p style="""
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        '借用url變數
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
                                '不處理
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
                                    '20241018 Copilot大菩薩： 設置段落背景色 (對應HTML中的 background-color: #ffffff)
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
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                    .text = "<p dir="""
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        '借用url變數
                        If VBA.InStr(rng.text, "ltr") Then rng.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                If VBA.Len(p.Range.text) > VBA.Len("《p align"""">") Then
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                    .text = "<p align="""
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        '借用url變數
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
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                    .text = "<p class="""
                    Do While .Execute()
                        rng.MoveEndUntil ">"
                        rng.End = rng.End + 1
                        rng.text = vbNullString
                        rng.SetRange p.Range.start, p.Range.End
                    Loop
                End If
                If VBA.Len(p.Range.text) > VBA.Len("<st1:personname ></st1:personname>") Then
                    rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
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
                'rng.SetRange p.Range.start, p.Range.End '用set 會歸零，用 setRange 不會，只是調整
                
            End With 'rng.Find
            
            
            If .Paragraphs(1).Range.text = "<br class=""Apple-interchange-newline""> " & Chr(13) Then
                .Paragraphs(1).Range.text = vbNullString
            End If
        End With 'rng
nextP:
    Next p
    
    Set rng = rngHtml.Document.Range(rngHtml.start, rngHtml.End)
    文字處理.FixFontname rng

    
finish:
  
    rngHtml.Document.Range(stRngHTML, stRngHTML).Select '回到起始位置
    
End Sub
