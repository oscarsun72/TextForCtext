Attribute VB_Name = "Excel"
Option Explicit '設定引用項目-VBA-引用Excel避免版本不合的問題，原理就是做一個叫做Excel的類別（模組）來仿真
Dim App As Object, wb As Object, sht As Object   '用Dim才能兼顧保留其生命週期與封裝性
'後期綁定(後期繫結） late bound
'https://dotblogs.com.tw/regionbbs/2016/10/13/concepts-in-late-binding
'https://docs.microsoft.com/zh-tw/previous-versions/office/troubleshoot/office-developer/binding-type-available-to-automation-clients
'https://docs.microsoft.com/zh-tw/dotnet/visual-basic/programming-guide/language-features/early-late-binding/
Enum XlFindLookIn
    xlComments = -4144
    xlFormulas = -4123
    xlValues = -4163
End Enum
Enum XlLookAt
    xlPart = 2
    xlWhole = 1
End Enum
Enum XlSearchOrder
    xlByColumns = 2
    xlByRows = 1
End Enum
Enum XlSearchDirection
    xlNext = 1
    xlPrevious = 2
End Enum


Property Get Application()
    'Stop
    'If VBA.IsEmpty(app) Then Class_Initialize
    If App Is Nothing Then Class_Initialize
    Set Application = App
End Property
Property Set Application(appOrNothing)
    Set App = appOrNothing 'can't be Empty! because once the app was set to Application object it'll be object type no longer be a variant type. Only the variant type could be the value of Empty.
End Property

Property Get Workbook()
    Set Workbook = wb
End Property
Property Get Worksheet()
    Set Worksheet = sht
End Property

Private Sub Class_Initialize()
    'Stop
    Set App = VBA.CreateObject("Excel.Application")
    App.UserControl = False 'for closing the app by user,this must be set to false or it will end until the word application close. https://docs.microsoft.com/zh-tw/office/vba/api/excel.application.usercontrol
    Set wb = App.Workbooks.Add() 'https://docs.microsoft.com/zh-tw/office/vba/api/excel.workbooks.add
    Set sht = wb.Sheets.Add()
End Sub
Rem 在Excel檔案中找到要找的文字
Rem 現用中的文件第1行指定Excel檔案全檔名
Sub FindPrivateUseCharactersInExcel()
    Static xlApp As Object
    Static xlBook As Object
    Static xlSheet As Object, myExcelFileFullnamePrevious
    Dim foundCell As Object, w As String, myExcelFileFullname As String
    On Error GoTo eH:
    'Static searchRange As Object
'    Dim firstAddress As String
    
    
    w = Selection.text
    'If Not code.IsPrivateUseCharacter(w) Then Exit Sub
        
    If Selection.Type = wdSelectionIP Then
        Selection.Characters(1).Copy
    Else
        Selection.Copy
    End If
'    Selection.Document.Save
    
    With Selection.Document
        myExcelFileFullname = _
            .Range(.Paragraphs(1).Range.Characters(1).start _
                , _
            .Paragraphs(1).Range.Characters( _
            .Paragraphs(1).Range.Characters.Count - 1).End).text
    End With
    If myExcelFileFullname = "" Then Exit Sub
    If myExcelFileFullname = VBA.Chr(13) Then Exit Sub
    If VBA.Dir(myExcelFileFullname) = "" Then
        MsgBox "所指定的全檔名有誤！", vbCritical
        Exit Sub
    End If
    
    SystemSetup.playSound 0.484
    
openWorkBook:
    ' 開啟Excel應用程式
    'Set xlApp = CreateObject("Excel.Application")
    If xlBook Is Nothing Then
        Set xlBook = GetObject(myExcelFileFullname)
        myExcelFileFullnamePrevious = myExcelFileFullname
        Set xlApp = xlBook.Application
        xlApp.UserControl = True
    
        ' 開啟指定的Excel檔案
        'Set xlBook = xlApp.Workbooks.Open("H:\我的雲端硬碟\黃老師遠端工作\3詞學\＃＃@@詞學韻律資料庫20240121@@●●.xlsm")
        Set xlSheet = xlBook.Sheets(1) ' 假設搜尋第一個工作表
    Else
        If myExcelFileFullnamePrevious <> myExcelFileFullname Then
            xlBook.Close SaveChanges:=False
            xlApp.Quit
            Set xlBook = Nothing
            GoTo openWorkBook
        End If
        
        Debug.Print xlBook.path
        
    End If
    
    SystemSetup.playSound 1
    
    If xlApp.Visible = False Then
        xlApp.Visible = True
        xlBook.Activate
        'xlSheet.Visible = True
        DoEvents
    End If
    DoEvents
    If xlBook.Windows(1).Visible = False Then
        xlBook.Windows(1).Visible = True
'        xlApp.Windows(1).Visible = True
        DoEvents
    End If
    
    
    ' 使用Find方法搜尋特定字元
'    Set foundCell = xlSheet.Cells.Find(What:=w, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    ' 設定搜尋範圍為G、H、I和M欄
    'Set searchRange = xlSheet.Range("G:G,H:H,I:I,M:M")
    ' 使用Find方法搜尋特定字元
    'Set foundCell = searchRange.Find(What:=w, After:=xlSheet.Application.ActiveCell, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
'    Set foundCell = xlSheet.Application.ActiveCell.Find _
        (What:=w, LookIn:=xlApp.XlFindLookIn.xlValues, _
            LookAt:=xlApp.XlLookAt.xlPart, SearchOrder:=xlApp.XlSearchOrder.xlByRows _
                , SearchDirection:=xlApp.XlSearchDirection.xlNext, MatchCase:=False)
    Set foundCell = xlSheet.Application.ActiveCell.Find(What:=w, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    
    If Not foundCell Is Nothing Then
        foundCell.Select
'        MsgBox "找到了字元 w！"
    Else
        MsgBox "未找到!", vbExclamation
    End If
    
    AppActivate xlBook.Application.Caption
'    ' 關閉Excel檔案
'    xlBook.Close SaveChanges:=False
'    xlApp.Quit
    
    ' 釋放物件
'    Set xlSheet = Nothing
'    Set xlBook = Nothing
'    Set xlApp = Nothing
Exit Sub
eH:
    Select Case Err.Number
        Case 7
            If VBA.InStr(Err.Description, "記憶體不足") Then
                Resume Next
            Else
                MsgBox Err.Number & Err.Description, vbCritical
            End If
        Case 9
            If VBA.InStr(Err.Description, "陣列索引超出範圍") Then
                DoEvents
                SystemSetup.playSound 1
                Resume
            Else
                MsgBox Err.Number & Err.Description, vbCritical
            End If
        Case 462
            If VBA.InStr(Err.Description, "遠端伺服器不存在或無法使用") Then
                DoEvents
                SystemSetup.playSound 1
                Set xlBook = Nothing
                GoTo openWorkBook
            Else
                MsgBox Err.Number & Err.Description, vbCritical
            End If
        Case -2147417848
            If VBA.InStr(Err.Description, "Automation 錯誤") Then
                Set xlBook = Nothing
                GoTo openWorkBook:
            Else
                MsgBox Err.Number & Err.Description, vbCritical
            End If
        Case Else
            MsgBox Err.Number & Err.Description, vbCritical
    End Select
End Sub
Rem 20240902 取代私人造字為系統用字 creedit_with_Copilot大菩薩：Excel 私人造字替換程式：https://sl.bing.net/cQSdvLaFCZo
Rem 現用中的文件第1行指定Excel檔案全檔名
Rem 現用中的文件第一個1表格須是私人造字（第1欄）與系統用字（第2欄）的對照表
Sub ReplacePrivateUseCharactersInExcel()
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim privateUseChar As String
    Dim replacementChar As String, myExcelFileFullname As String
    Dim tb As Table
    
    
    With Selection.Document
        If .Tables.Count < 1 Then
            MsgBox "現用中的文件第一個1表格須是私人造字（第1欄）與系統用字（第2欄）的對照表", vbCritical
            Exit Sub
        Else
            Set tb = .Tables(1)
        End If
        myExcelFileFullname = _
            .Range(.Paragraphs(1).Range.Characters(1).start _
                , _
            .Paragraphs(1).Range.Characters( _
            .Paragraphs(1).Range.Characters.Count - 1).End).text
    End With
    If myExcelFileFullname = "" Then Exit Sub
    If myExcelFileFullname = VBA.Chr(13) Then Exit Sub
    If VBA.Dir(myExcelFileFullname) = "" Then
        MsgBox "所指定的全檔名有誤！", vbCritical
        Exit Sub
    End If
    
    SystemSetup.playSound 0.484
    
    ' 開啟 Excel 應用程式
'    Set xlApp = CreateObject("Excel.Application")
    Set xlApp = Excel.Application
    xlApp.UserControl = True
    xlApp.Visible = True
    
    ' 開啟 Excel 工作簿
'    Set xlBook = xlApp.Workbooks.Open("C:\path\to\your\file.xlsx")
    Set xlBook = xlApp.Workbooks.Open(myExcelFileFullname)
    Set xlSheet = xlBook.Sheets(1) ' 根據需要修改工作表索引
    
    xlApp.DisplayAlerts = False
    
    Dim r As row, i As Long
    For Each r In tb.Rows
        i = i + 1
        ' 設定私人造字區的範圍
        'privateUseChar = "[\uE000-\uF8FF]" ' 這是私人造字區的範圍
        privateUseChar = r.Cells(1).Range.Characters(1).text
        'replacementChar = "?" ' 替換字元，可以根據需要修改
        replacementChar = r.Cells(2).Range.Characters(1).text
        If privateUseChar <> vbNullString And replacementChar <> vbNullString Then
    
            ' 使用 Replace 方法替換私人造字區的字元
            xlSheet.Cells.Replace What:=privateUseChar, Replacement:=replacementChar, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
            If i Mod 10 = 0 Then
                'Beep
                SystemSetup.playSound 1
            End If
            
        End If
    Next r
    xlApp.DisplayAlerts = True
    ' 儲存並關閉工作簿
    If MsgBox("私人造字已替換完成！是否儲存？", vbInformation + vbOKCancel) = vbOK Then
        xlBook.Save
    End If
'    xlBook.Close
'    xlApp.Quit
'
'    ' 釋放物件
'    Set xlSheet = Nothing
'    Set xlBook = Nothing
'    Set xlApp = Nothing
    
    
End Sub

