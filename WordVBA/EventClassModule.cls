VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventClassModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'https://learn.microsoft.com/zh-tw/office/vba/word/concepts/objects-properties-methods/using-events-with-the-application-object-word
'https://learn.microsoft.com/zh-tw/office/vba/word/concepts/objects-properties-methods/using-events-with-the-application-object-word
Public WithEvents App As word.Application
Attribute App.VB_VarHelpID = -1

'Public WithEvents EventClassModule As Word.Application
'Private Sub EventClassModule_WindowActivate _
'        (ByVal Wn As Word.Window)
'    Wn.WindowState = wdWindowStateMaximize
'    ���ær��
'End Sub

'20230120 �M chatGPT�j���İQ�ת����G�A�bWord�������O�Ҳդ��O����[�J App_Startup() ���{�Ǫ��A�u��� AutoExec �Ҳժ���k�ӹF���Ұ� Word�ɰ���S�w�{��
'Private Sub App_Startup()
'    '�b�o�̲K�[�z�Q�n�bWord�Ұʮɰ��檺�{��
'    MsgBox "Word�ҰʤF"
'End Sub
'
'Private Sub App_Startup()
''20230120 12:10 ��J�p�~�] creedit chatGPT�j���ġGVBA  �p�󵲧� MS word �ɳB�z�{�ǩO�G
'    '�b�o�̲K�[�z�Q�n���檺�{��
'    SystemSetup.UserProfilePath = SystemSetup.���o�ϥΪ̸��|_�t�ϱ׽u()
'End Sub

Private Sub app_Quit()

'�۰ʳƥ�Normal.dotm
backupNormal_dotm

'�M��Word Selenium�Ҷ}�Ҫ�chromedriver
'20230119 chatGPT���ġGVBA  �p�󵲧� MS word �ɳB�z�{�ǩO�G
'https://learn.microsoft.com/zh-tw/office/vba/word/concepts/objects-properties-methods/using-events-with-the-application-object-word
SystemSetup.killchromedriverFromHere
End Sub


Private Sub App_DocumentBeforeSave(ByVal Doc As Document, SaveAsUI As Boolean, Cancel As Boolean)
'���ær��
End Sub

Private Sub App_DocumentChange() '�g�b���B������!2010/10/20
'Register_Event_Handler
'��ӨS�ĬO�]���ڦb"�b�t�@��󤤴M�����r��"�ϥ�End�ϵ��U�{��Register_Event_Handler���ĤF!
���ær��
'If InStr(ActiveDocument.Name, "�r��7") = 0 Then
'    if Documents("�r��7.2.doc"). .Visible = False
'End If
End Sub



Private Sub App_WindowActivate(ByVal Doc As Document, ByVal Wn As Window)
''If InStr(Doc.Name, "�r��7") = 0 Then Wn.Visible = False
''    'Wn.WindowState = wdWindowStateMaximize
''    ���ær��
'If VBA.Left(Wn.Caption, 2) = "�ƦL" And VBA.IsNumeric(Mid(Wn.Caption, 3, 4)) Then
'    If Tasks.Exists(Mid(Wn.Caption, 3) & ".pdf") Then
'       Tasks(Mid(Wn.Caption, 3) & ".pdf").Activate
''    If Tasks("Adobe Reader").Exists Then
''       Tasks(Mid(Wn.Caption, 3, 4) & "*pdf*").Activate
'       Tasks(Wn.Caption & " - Microsoft Word").Activate
'    End If
'End If
End Sub

Private Sub App_WindowDeactivate(ByVal Doc As Document, ByVal Wn As Window) '�g�bApp_DocumentChange������!2010/10/20
'Register_Event_Handler
'If Documents.Count > 1 Then
'    If InStr(Doc.Name, "�r��7") Then Wn.Visible = False
'End If
���ær��
End Sub