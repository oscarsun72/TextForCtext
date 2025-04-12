Attribute VB_Name = "FileSystemOP"
Option Explicit

Rem Copilot大菩薩 20250409：
'Copilot大菩薩吉祥：我想寫一個WordVBA程式，在指定的資料夾路徑下中遍歷所有的txt當，找到其中包含指定中文關鍵字的檔案，條件其全檔名這樣的結果在一個新的word文件中的各段（一段一個全檔名） 請幫我完成，好嗎？我再進行測試。感恩感恩　南無阿彌陀佛
Sub FindTxtFilesWithKeyword()
    Dim folderPath As String
    Dim keyword As String
    Dim fileName As String
    Dim fileContent As String
    Dim doc As Document
    Dim rng As Range
    Dim fso As Object
    Dim file As Object
    Dim d As Document
    
    Set d = ActiveDocument
    ' 設定資料夾路徑和關鍵字在第1段
    'folderPath = "C:\YourFolderPath\" ' 請替換成您的資料夾路徑
    folderPath = d.Range(d.Paragraphs(1).Range.start, d.Paragraphs(1).Range.End - 1).text
    
    'keyword = "指定中文關鍵字" ' 請替換成您的關鍵字在第2段
    keyword = d.Range(d.Paragraphs(2).Range.start, d.Paragraphs(2).Range.End - 1).text
    
    ' 建立新的 Word 文件
    Set doc = Documents.Add

    ' 使用 FileSystemObject 瀏覽資料夾
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        MsgBox "資料夾不存在：" & folderPath, vbExclamation
        Exit Sub
    End If

    For Each file In fso.GetFolder(folderPath).Files
        If LCase(fso.GetExtensionName(file.Name)) = "txt" Then
            fileName = file.path
            ' 開啟 .txt 檔案並檢查內容
            fileContent = GetFileContent(fileName)
            If InStr(fileContent, keyword) > 0 Then
                ' 將檔案名添加到 Word 文件中
                Set rng = doc.content
                rng.Collapse wdCollapseEnd
                rng.InsertAfter fileName & vbCrLf
                rng.InsertParagraphAfter
                rng.Hyperlinks.Add rng, fileName
            End If
        End If
    Next file

    MsgBox "完成！檔案名已添加到 Word 文件中。", vbInformation
End Sub

Function GetFileContent(filePath As String) As String
    Dim stream As Object
    Dim content As String

    ' 使用 ADODB.Stream 來讀取 UTF-8 編碼的文字檔案
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' 設置為文字模式
    stream.Mode = 3 ' 設置為讀寫模式
    stream.Charset = "UTF-8" ' 指定編碼為 UTF-8
    stream.Open
    stream.LoadFromFile filePath
    content = stream.ReadText
    stream.Close

    GetFileContent = content
End Function
