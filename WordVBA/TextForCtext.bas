Attribute VB_Name = "TextForCtext"
Option Explicit
Rem TextForCtext�����ާ@
Private Property Get tx() As String
    tx = "TextForCtext"
End Property
Property Get TextForCtextExist()
    TextForCtextExist = word.Tasks.Exists(TextForCtext.tx)
End Property
Rem �ˬdTextForCtext�O�_�Ұ� 20240910
Private Function examToRun() As Boolean
    If Not word.Tasks.Exists(TextForCtext.tx) Then Exit Function
    'SystemSetup.wait 0.3
    Dim dt As Date
    dt = VBA.Now
    Do While DateDiff("s", dt, VBA.Now) < 0.3
        DoEvents
    Loop
    examToRun = True
End Function

Sub Hanchi_CTP_SearchingKeywordsYijing()
    ' Alt + shift + ,
    ' Alt + ,
    ' Alt + shift + F5
    ' Ctrl + Alt + F9
    SystemSetup.playSound 0.484

    If Not examToRun Then Exit Sub
    
    On Error GoTo eH:
    AppActivate tx
    DoEvents
    SendKeys "%{F9}"
    DoEvents
    Exit Sub
eH:
    Select Case Err.Number
        Case 5 '�{�ǩI�s�Τ޼Ƥ����T
            If vbOK = MsgBox("�Ы�_TextForCtext�������A���T�w�~��", vbOKCancel + vbExclamation) Then
                Resume
            End If
        Case Else
            MsgBox Err.Number + Err.Description
    End Select
End Sub
Rem �e��m�j�y�šn�۰ʼ��I�C���ƻs�n�O�n�B�z���¤�r�C�N���GŪ�^�ܰŶKï��
Function GjcoolPunct() As Boolean
    
    If Not examToRun Then Exit Function
    
    On Error GoTo eH:
    AppActivate tx
    DoEvents
    SystemSetup.wait 0.05
    SendKeys "^a"
    DoEvents
    SystemSetup.wait 0.02
    SendKeys "{delete}"
    DoEvents
    AppActivate tx
    SystemSetup.wait 0.1
    DoEvents
    
    Rem �K�W
    'SendKeys "+{INSERT}"'�]��TextForCtext�� textBox1_TextChanged ���p�U���l�A�ҥH������U shift�A�G��� ctrl+v
                                    ' �K�K   {//�b��ʿ�J�Ҧ��U
                                    '    if (mk != Keys.None)
                                    '    {//�i����UShift+Delete �ŤUtextBox1�����e��
                                    '        hideToNICo(); �K�K
    SendKeys "^v" 'Ctrl + v
    DoEvents
    
    
    AppActivate tx
    DoEvents
    SendKeys "^%{F10}"
    DoEvents
    Dim dt As Date, x As String
    Dim containsPunctuation As Boolean, punctuation As String, i As Byte, noErrOccured As Boolean
    dt = VBA.Now
    Do While DateDiff("s", dt, VBA.Now) < 30
        If DateDiff("s", dt, VBA.Now) Mod 1 = 0 Then
            GoSub puncted
            If containsPunctuation Then
                noErrOccured = True
                Exit Do
            End If
        End If
        DoEvents
        
    Loop
    If Not noErrOccured Then
        GjcoolPunct = False
        Exit Function
    End If
    DoEvents
    AppActivate tx
    DoEvents
    SystemSetup.wait 0.1
    SendKeys "^a"
    DoEvents
    'SendKeys "^x" '�ƻs�奻���\��g�bC#��
    SystemSetup.wait 0.05
    SendKeys "{delete}"
    DoEvents
    GjcoolPunct = True
    Exit Function
puncted:
    
    punctuation = "�C�A"
    For i = 1 To Len(punctuation)
        If InStr(SystemSetup.GetClipboard, VBA.Mid(punctuation, i, 1)) > 0 Then
            containsPunctuation = True
            Exit For
        End If
    Next i
Return

eH:
    Select Case Err.Number
        Case 5 '�{�ǩI�s�Τ޼Ƥ����T
            If vbOK = MsgBox("�Ы�_TextForCtext�������A���T�w�~��", vbOKCancel + vbExclamation) Then
                Resume
            End If
        Case Else
            MsgBox Err.Number + Err.Description
    End Select
    
End Function

