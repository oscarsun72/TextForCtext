Attribute VB_Name = "AutoExec"
Option Explicit


'Sub AutoExec()
''    Stop
'    UserProfilePath = SystemSetup.���o�ϥΪ̸��|_�t�ϱ׽u()
'    '�b�o�̲K�[�z�Q�n���檺�{��
''    SystemSetup.ShortcutKeys
'    Stop
'    Register_Event_Handler
'
'End Sub

Rem 20240903 Copilot�j���ġGWord VBA �ƥ�B�z�{�ǵ��U�Ghttps://sl.bing.net/jKWC0wBWaFo
'�n�� AutoExec �����b Normal.dotm ���B��ðѷ� Startup ���|�U�� TextForCtextWordVBA.dotm �d���̪��M�סA�z�i�H�ϥ� Application.Run ��k�ӽե� TextForCtextWordVBA.dotm �����{�ǡC�H�U�O�ק�᪺�d�ҡG
'�b Normal.dotm ���� AutoExec �����G
Sub AutoExec()
'    Stop
    ' �T�O TextForCtextWordVBA.dotm �w�[��
    AddInLoad "TextForCtextWordVBA.dotm"
    
    ' �ե� TextForCtextWordVBA.dotm �����{�ǡ]���O sub �~�����^
    'Application.Run "TextForCtextWordVBA.SystemSetup.���o�ϥΪ̸��|_�t�ϱ׽u"
    
    ' �b�o�̲K�[�z�Q�n���檺�{��
    'Application.Run "TextForCtextWordVBA.SystemSetup.ShortcutKeys"
    
    ' ���U�ƥ�B�z�{��
    'Register_Event_Handler
    Application.Run "TextForCtextWordVBA.Docs.Register_Event_Handler"
End Sub

Sub AddInLoad(addInName As String)
    Dim addIn As Template
    On Error Resume Next
    Set addIn = Application.Templates(addInName)
    If addIn Is Nothing Then
        Set addIn = Application.AddIns.Add(FileName:=Application.StartupPath & "\" & addInName, Install:=True)
    End If
    On Error GoTo 0
End Sub
