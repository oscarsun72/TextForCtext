Attribute VB_Name = "Excel"
Option Explicit '�]�w�ޥζ���-VBA-�ޥ�Excel�קK�������X�����D�A��z�N�O���@�ӥs��Excel�����O�]�Ҳա^�ӥ�u
Dim App As Object, wb As Object, sht As Object   '��Dim�~����U�O�d��ͩR�g���P�ʸ˩�
'����j�w(���ô���^ late bound
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
Rem �bExcel�ɮפ����n�䪺��r
Rem �{�Τ�������1����wExcel�ɮץ��ɦW
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
        MsgBox "�ҫ��w�����ɦW���~�I", vbCritical
        Exit Sub
    End If
    
    SystemSetup.playSound 0.484
    
openWorkBook:
    ' �}��Excel���ε{��
    'Set xlApp = CreateObject("Excel.Application")
    If xlBook Is Nothing Then
        Set xlBook = GetObject(myExcelFileFullname)
        myExcelFileFullnamePrevious = myExcelFileFullname
        Set xlApp = xlBook.Application
        xlApp.UserControl = True
    
        ' �}�ҫ��w��Excel�ɮ�
        'Set xlBook = xlApp.Workbooks.Open("H:\�ڪ����ݵw��\���Ѯv���ݤu�@\3����\����@@�������߸�Ʈw20240121@@����.xlsm")
        Set xlSheet = xlBook.Sheets(1) ' ���]�j�M�Ĥ@�Ӥu�@��
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
    
    
    ' �ϥ�Find��k�j�M�S�w�r��
'    Set foundCell = xlSheet.Cells.Find(What:=w, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    ' �]�w�j�M�d��G�BH�BI�MM��
    'Set searchRange = xlSheet.Range("G:G,H:H,I:I,M:M")
    ' �ϥ�Find��k�j�M�S�w�r��
    'Set foundCell = searchRange.Find(What:=w, After:=xlSheet.Application.ActiveCell, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
'    Set foundCell = xlSheet.Application.ActiveCell.Find _
        (What:=w, LookIn:=xlApp.XlFindLookIn.xlValues, _
            LookAt:=xlApp.XlLookAt.xlPart, SearchOrder:=xlApp.XlSearchOrder.xlByRows _
                , SearchDirection:=xlApp.XlSearchDirection.xlNext, MatchCase:=False)
    Set foundCell = xlSheet.Application.ActiveCell.Find(What:=w, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    
    If Not foundCell Is Nothing Then
        foundCell.Select
'        MsgBox "���F�r�� w�I"
    Else
        MsgBox "�����!", vbExclamation
    End If
    
    AppActivate xlBook.Application.Caption
'    ' ����Excel�ɮ�
'    xlBook.Close SaveChanges:=False
'    xlApp.Quit
    
    ' ���񪫥�
'    Set xlSheet = Nothing
'    Set xlBook = Nothing
'    Set xlApp = Nothing
Exit Sub
eH:
    Select Case Err.Number
        Case 7
            If VBA.InStr(Err.Description, "�O���餣��") Then
                Resume Next
            Else
                MsgBox Err.Number & Err.Description, vbCritical
            End If
        Case 9
            If VBA.InStr(Err.Description, "�}�C���޶W�X�d��") Then
                DoEvents
                SystemSetup.playSound 1
                Resume
            Else
                MsgBox Err.Number & Err.Description, vbCritical
            End If
        Case 462
            If VBA.InStr(Err.Description, "���ݦ��A�����s�b�εL�k�ϥ�") Then
                DoEvents
                SystemSetup.playSound 1
                Set xlBook = Nothing
                GoTo openWorkBook
            Else
                MsgBox Err.Number & Err.Description, vbCritical
            End If
        Case -2147417848
            If VBA.InStr(Err.Description, "Automation ���~") Then
                Set xlBook = Nothing
                GoTo openWorkBook:
            Else
                MsgBox Err.Number & Err.Description, vbCritical
            End If
        Case Else
            MsgBox Err.Number & Err.Description, vbCritical
    End Select
End Sub
Rem 20240902 ���N�p�H�y�r���t�ΥΦr creedit_with_Copilot�j���ġGExcel �p�H�y�r�����{���Ghttps://sl.bing.net/cQSdvLaFCZo
Rem �{�Τ�������1����wExcel�ɮץ��ɦW
Rem �{�Τ������Ĥ@��1��涷�O�p�H�y�r�]��1��^�P�t�ΥΦr�]��2��^����Ӫ�
Sub ReplacePrivateUseCharactersInExcel()
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim privateUseChar As String
    Dim replacementChar As String, myExcelFileFullname As String
    Dim tb As Table
    
    
    With Selection.Document
        If .Tables.Count < 1 Then
            MsgBox "�{�Τ������Ĥ@��1��涷�O�p�H�y�r�]��1��^�P�t�ΥΦr�]��2��^����Ӫ�", vbCritical
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
        MsgBox "�ҫ��w�����ɦW���~�I", vbCritical
        Exit Sub
    End If
    
    SystemSetup.playSound 0.484
    
    ' �}�� Excel ���ε{��
'    Set xlApp = CreateObject("Excel.Application")
    Set xlApp = Excel.Application
    xlApp.UserControl = True
    xlApp.Visible = True
    
    ' �}�� Excel �u�@ï
'    Set xlBook = xlApp.Workbooks.Open("C:\path\to\your\file.xlsx")
    Set xlBook = xlApp.Workbooks.Open(myExcelFileFullname)
    Set xlSheet = xlBook.Sheets(1) ' �ھڻݭn�ק�u�@�����
    
    xlApp.DisplayAlerts = False
    
    Dim r As row, i As Long
    For Each r In tb.Rows
        i = i + 1
        ' �]�w�p�H�y�r�Ϫ��d��
        'privateUseChar = "[\uE000-\uF8FF]" ' �o�O�p�H�y�r�Ϫ��d��
        privateUseChar = r.Cells(1).Range.Characters(1).text
        'replacementChar = "?" ' �����r���A�i�H�ھڻݭn�ק�
        replacementChar = r.Cells(2).Range.Characters(1).text
        If privateUseChar <> vbNullString And replacementChar <> vbNullString Then
    
            ' �ϥ� Replace ��k�����p�H�y�r�Ϫ��r��
            xlSheet.Cells.Replace What:=privateUseChar, Replacement:=replacementChar, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
            If i Mod 10 = 0 Then
                'Beep
                SystemSetup.playSound 1
            End If
            
        End If
    Next r
    xlApp.DisplayAlerts = True
    ' �x�s�������u�@ï
    If MsgBox("�p�H�y�r�w���������I�O�_�x�s�H", vbInformation + vbOKCancel) = vbOK Then
        xlBook.Save
    End If
'    xlBook.Close
'    xlApp.Quit
'
'    ' ���񪫥�
'    Set xlSheet = Nothing
'    Set xlBook = Nothing
'    Set xlApp = Nothing
    
    
End Sub

