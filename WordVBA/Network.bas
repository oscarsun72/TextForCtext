Attribute VB_Name = "Network"
Option Explicit

Sub �d�߰�y���() '���w��:Ctrl+F12'2010/10/18�׭q
''    If ActiveDocument.Path <> "" Then ActiveDocument.Save '��word�����ѤF�x�s
''    If GetUserAddress = True Then
'''        MsgBox "���\�����H�W�s���C"
''    Else
''        MsgBox "�L�k���H�W�s���C"
''    End If
'    Selection.Copy
'    Shell "W:\!! for hpr\VB\�d�߰�y���\�d�߰�y���\bin\Debug\�d�߰�y���.EXE"
Const st As String = "C:\Program Files\�]�u�u\�d�߰�y��嵥\"
Const f As String = "�d�߰�y���.EXE"
Dim funame As String
If Selection.Type = wdSelectionNormal Then
    Selection.Copy
    If Dir(st & f) <> "" Then
        funame = st & f
    ElseIf Dir("C:\Program Files (x86)\�]�u�u\�d�߰�y��嵥\" & f) <> "" Then
        funame = "C:\Program Files (x86)\�]�u�u\�d�߰�y��嵥\" & f
    ElseIf Dir("W:\!! for hpr\VB\�d�߰�y���\�d�߰�y���\bin\Debug\" & f) <> "" Then
        funame = "W:\!! for hpr\VB\�d�߰�y���\�d�߰�y���\bin\Debug\" & f
    ElseIf Dir("C:\�d�߰�y���\�d�߰�y���\bin\Debug\" & f) <> "" Then
        funame = "C:\�d�߰�y���\�d�߰�y���\bin\Debug\" & f
    ElseIf Dir(userProfilePath & "Dropbox\VS\VB\�d�߰�y���\�d�߰�y���\bin\Debug\" & f) <> "" Then
        funame = userProfilePath & "Dropbox\VS\VB\�d�߰�y���\�d�߰�y���\bin\Debug\" & f
    ElseIf Dir("A:\", vbVolume) <> "" Then
        If Dir("A:\Users\oscar\Dropbox\VS\VB\�d�߰�y���\�d�߰�y���\bin\Debug\" & f) <> "" Then _
        funame = "A:\Users\oscar\Dropbox\VS\VB\�d�߰�y���\�d�߰�y���\bin\Debug\" & f
    ElseIf Dir(userProfilePath & "Dropbox\VS\VB\�d�߰�y���\�d�߰�y���\bin\Debug\" & f) <> "" Then
        funame = userProfilePath & "Dropbox\VS\VB\�d�߰�y���\�d�߰�y���\bin\Debug\" & f
    Else
        Exit Sub
    End If
    Shell funame
End If
End Sub

Sub A�t�˺����r���() '���w��:Alt+F12'2010/10/18�׭q
Const f As String = "�t�˺����r���.EXE"
Const st As String = "C:\Program Files\�]�u�u\�t�˺����r���\"
Dim funame As String
If Selection.Type = wdSelectionNormal Then
    Selection.Copy
    If Dir(st & f) <> "" Then
        funame = st & f
    ElseIf Dir("C:\Program Files (x86)\�]�u�u\�t�˺����r���\" & f) <> "" Then
        funame = "C:\Program Files (x86)\�]�u�u\�t�˺����r���\" & f
    ElseIf Dir("W:\!! for hpr\VB\�t�˺����r���\�t�˺����r���\bin\Debug\" & f) <> "" Then
        funame = "W:\!! for hpr\VB\�t�˺����r���\�t�˺����r���\bin\Debug\" & f
    ElseIf Dir("C:\�t�˺����r���\�t�˺����r���\bin\Debug\" & f) <> "" Then
        funame = "C:\�t�˺����r���\�t�˺����r���\bin\Debug\" & f
    ElseIf Dir(userProfilePath & "Dropbox\VS\VB\�t�˺����r���\�t�˺����r���\bin\Debug\" & f) <> "" Then
        funame = "A:\Users\oscar\Dropbox\VS\VB\�t�˺����r���\�t�˺����r���\bin\Debug\" & f
    ElseIf Dir(userProfilePath & "Dropbox\VS\VB\�t�˺����r���\�t�˺����r���\bin\Debug\" & f) <> "" Then
        funame = "A:\Users\oscar\Dropbox\VS\VB\�t�˺����r���\�t�˺����r���\bin\Debug\" & f
    ElseIf Dir(userProfilePath & "Dropbox\VS\VB\�t�˺����r���\�t�˺����r���\bin\Debug\" & f) <> "" Then
        funame = userProfilePath & "Dropbox\VS\VB\�t�˺����r���\�t�˺����r���\bin\Debug\" & f
    Else
        Exit Sub
    End If
    Shell funame
End If

End Sub

Function GetUserAddress() As Boolean
    Dim x As String, a As Object 'Access.Application
    On Error GoTo Error_GetUserAddress
    x = Selection.Text
    Set a = GetObject("D:\�d�{�@�o�N\���y���\�ϮѺ޲z.mdb") '2010/10/18�׭q
    If x = "" Then x = InputBox("�п�J���d�ߪ��r��")
    x = a.Run("�d�ߦr���ഫ_��y�|�X", x)
''    'ActiveDocument.FollowHyperlink "http://140.111.34.46/cgi-bin/dict/newsearch.cgi", , False, , "Database=dict&GraphicWord=yes&QueryString=^" & X & "$", msoMethodGet
'    FollowHyperlink "http://dict.revised.moe.edu.tw/cgi-bin/newDict/dict.sh?", , False, , "=dict.idx&cond=^" & x & "$&pieceLen=50&fld=1&cat=&imgFont=1", msoMethodGet
    Shell Replace(GetDefaultBrowserEXE, """%1", "http://dict.revised.moe.edu.tw/cgi-bin/newDict/dict.sh?cond=^" & x & "$&pieceLen=50&fld=1&cat=&imgFont=1")
    'AppActivate GetDefaultBrowser'�L��
'    'FollowHyperlink "http://dict.revised.moe.edu.tw/cgi-bin/newDict/dict.sh?", , False, , "=dict.idx&cond=^" & X & "$&pieceLen=50&fld=1&cat=&imgFont=1", msoMethodGet
    
'    If Len(Selection.Text) = 1 Then _
        FollowHyperlink "http://www.nlcsearch.moe.gov.tw/EDMS/admin/dict3/search.php", , False, , "qstr=" & x & "&dictlist=47,46,51,18,16,13,20,19,53,12,14,17,48,57,24,25,26,29,30,31,32,33,34,35,36,37,39,38,41,42,43,45,50,&searchFlag=A&hdnCheckAll=checked", msoMethodGet '2009/1/10'�Ш|��-��a�y���X�s���˯��t��-�y���X�˯�
        If a.Visible = False Then
            a.Visible = True
            a.UserControl = True
        End If
'        a.Quit acQuitSaveNone
'        Set a = Nothing
    GetUserAddress = True
Exit_GetUserAddress:
    Exit Function

Error_GetUserAddress:
    MsgBox Err & ": " & Err.Description
    GetUserAddress = False
    Resume Exit_GetUserAddress
End Function


    
Function GetDefaultBrowser() '2010/10/18��http://chijanzen.net/wp/?p=156#comment-1303(���o�w�]�s����(default web browser)���W��? chijanzen ���f�E)�Ө�.
    Dim objShell
    Set objShell = CreateObject("WScript.Shell")
    'HKEY_CLASSES_ROOT\HTTP\shell\open\ddeexec\Application
    '���o���U��������
    GetDefaultBrowser = objShell.RegRead _
            ("HKCR\http\shell\open\ddeexec\Application\")
    'GetDefaultBrowser = objShell.RegRead _
            ("HKEY_CLASSES_ROOT\http\shell\open\ddeexec\Application\")
End Function


Function GetDefaultBrowserEXE() '2010/10/18��http://chijanzen.net/wp/?p=156#comment-1303(���o�w�]�s����(default web browser)���W��? chijanzen ���f�E)�Ө�.
    Dim objShell
    Set objShell = CreateObject("WScript.Shell")
    'HKEY_CLASSES_ROOT\HTTP\shell\open\ddeexec\Application
    '���o���U��������
    GetDefaultBrowserEXE = objShell.RegRead _
            ("HKCR\http\shell\open\command\")
    
End Function

Sub AppActivateDefaultBrowser()
Dim DefaultBrowserName As String
On Error GoTo eh
DefaultBrowserName = "google chrome"
DoEvents
AppActivate DefaultBrowserName
DoEvents
Exit Sub
eh:
DefaultBrowserName = "brave"
Resume
'AppActivate ""
End Sub


