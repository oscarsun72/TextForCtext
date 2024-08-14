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
Set sht = wb.sheets.Add()
End Sub


