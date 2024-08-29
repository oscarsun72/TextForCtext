Attribute VB_Name = "Excel"
Option Explicit '�]�w�ޥζ���-VBA-�ޥ�Excel�קK�������X�����D�A��z�N�O���@�ӥs��Excel�����O�]�Ҳա^�ӥ�u
Dim App As Object, wb As Object, sht As Object   '��Dim�~����U�O�d��ͩR�g���P�ʸ˩�
'����j�w(���ô���^ late bound
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
        MsgBox "�ҫ��w�����ɦW���~�I", vbCritical
        Exit Sub
    End If
    
    
    ' �}��Excel���ε{��
    'Set xlApp = CreateObject("Excel.Application")
    If xlBook Is Nothing Then
        Set xlBook = GetObject(myExcelFileFullname)
        Set xlApp = xlBook.Application
    
        ' �}�ҫ��w��Excel�ɮ�
        'Set xlBook = xlApp.Workbooks.Open("H:\�ڪ����ݵw��\���Ѯv���ݤu�@\3����\����@@�������߸�Ʈw20240121@@����.xlsm")
        Set xlSheet = xlBook.Sheets(1) ' ���]�j�M�Ĥ@�Ӥu�@��
    
    End If
    
    If xlApp.Visible = False Then
'        xlApp.Visible = True
        xlApp.Windows(1).Visible = True
    End If
    
    ' �ϥ�Find��k�j�M�S�w�r��
    'Set foundCell = xlSheet.Cells.Find(What:=w, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    ' �]�w�j�M�d��G�BH�BI�MM��
    'Set searchRange = xlSheet.Range("G:G,H:H,I:I,M:M")
    ' �ϥ�Find��k�j�M�S�w�r��
    'Set foundCell = searchRange.Find(What:=w, After:=xlSheet.Application.ActiveCell, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
'    Set foundCell = xlSheet.Application.ActiveCell.Find _
        (What:=w, LookIn:=xlApp.XlFindLookIn.xlValues, _
            LookAt:=xlApp.XlLookAt.xlPart, SearchOrder:=xlApp.XlSearchOrder.xlByRows _
                , SearchDirection:=xlApp.XlSearchDirection.xlNext, MatchCase:=False)
    Set foundCell = xlSheet.Application.ActiveCell.Find(What:=w, LookIn:=-4163, LookAt:=2, SearchOrder:=1, SearchDirection:=1, MatchCase:=False)
    
    If Not foundCell Is Nothing Then
        foundCell.Select
'        MsgBox "���F�r�� w�I"
    Else
        MsgBox "�����!"
    End If
    
    AppActivate xlBook.Application.Caption
'    ' ����Excel�ɮ�
'    xlBook.Close SaveChanges:=False
'    xlApp.Quit
    
    ' ���񪫥�
'    Set xlSheet = Nothing
'    Set xlBook = Nothing
'    Set xlApp = Nothing
End Sub


