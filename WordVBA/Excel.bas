Attribute VB_Name = "Excel"
Option Explicit '設定引用項目-VBA-引用Excel避免版本不合的問題，原理就是做一個叫做Excel的類別（模組）來仿真
Dim App As Object, wb As Object, sht As Object   '用Dim才能兼顧保留其生命週期與封裝性
'後期綁定(後期繫結） late bound
'https://dotblogs.com.tw/regionbbs/2016/10/13/concepts-in-late-binding
'https://docs.microsoft.com/zh-tw/previous-versions/office/troubleshoot/office-developer/binding-type-available-to-automation-clients
'https://docs.microsoft.com/zh-tw/dotnet/visual-basic/programming-guide/language-features/early-late-binding/

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
    Set wb = App.workbooks.Add() 'https://docs.microsoft.com/zh-tw/office/vba/api/excel.workbooks.add
    Set sht = wb.Sheets.Add()
End Sub

Sub FindPrivateUseCharactersInExcel()
    Static xlApp As Object
    Static xlBook As Object, w As String
    Static xlSheet As Object, myExcelFileFullname As String
    Dim foundCell As Object
    'Static searchRange As Object
'    Dim firstAddress As String
    
    w = Selection.text
    'If Not code.IsPrivateUseCharacter(w) Then Exit Sub
        
    If Selection.Type = wdSelectionIP Then Selection.Characters(1).Copy
'    Selection.Document.Save
    
    With Selection.Document
        myExcelFileFullname = _
            .Range(.Paragraphs(1).Range.Characters(1).start _
                , _
            .Paragraphs(1).Range.Characters( _
            .Paragraphs(1).Range.Characters.Count - 1).End).text
    End With
    If myExcelFileFullname = "" Then Exit Sub
    If myExcelFileFullname = Chr(13) Then Exit Sub
    If VBA.Dir(myExcelFileFullname) = "" Then
        MsgBox "所指定的全檔名有誤！", vbCritical
        Exit Sub
    End If
    
    
    ' 開啟Excel應用程式
    'Set xlApp = CreateObject("Excel.Application")
    If xlBook Is Nothing Then
        Set xlBook = GetObject(myExcelFileFullname)
        Set xlApp = xlBook.Application
    
        ' 開啟指定的Excel檔案
        'Set xlBook = xlApp.Workbooks.Open("H:\我的雲端硬碟\黃老師遠端工作\3詞學\＃＃@@詞學韻律資料庫20240121@@●●.xlsm")
        Set xlSheet = xlBook.Sheets(1) ' 假設搜尋第一個工作表
    
    End If
    
    If xlApp.Visible = False Then
'        xlApp.Visible = True
        xlApp.Windows(1).Visible = True
    End If
    
    ' 使用Find方法搜尋特定字元
    'Set foundCell = xlSheet.Cells.Find(What:=w, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    ' 設定搜尋範圍為G、H、I和M欄
    'Set searchRange = xlSheet.Range("G:G,H:H,I:I,M:M")
    ' 使用Find方法搜尋特定字元
    'Set foundCell = searchRange.Find(What:=w, After:=xlSheet.Application.ActiveCell, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
'    Set foundCell = xlSheet.Application.ActiveCell.Find _
        (What:=w, LookIn:=xlApp.XlFindLookIn.xlValues, _
            LookAt:=xlApp.XlLookAt.xlPart, SearchOrder:=xlApp.XlSearchOrder.xlByRows _
                , SearchDirection:=xlApp.XlSearchDirection.xlNext, MatchCase:=False)
    Set foundCell = xlSheet.Application.ActiveCell.Find(What:=w, LookIn:=-4163, LookAt:=2, SearchOrder:=1, SearchDirection:=1, MatchCase:=False)
    
    If Not foundCell Is Nothing Then
        foundCell.Select
'        MsgBox "找到了字元 w！"
    Else
        MsgBox "未找到!"
    End If
    
    AppActivate xlBook.Application.Caption
'    ' 關閉Excel檔案
'    xlBook.Close SaveChanges:=False
'    xlApp.Quit
    
    ' 釋放物件
'    Set xlSheet = Nothing
'    Set xlBook = Nothing
'    Set xlApp = Nothing
End Sub


