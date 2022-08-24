Attribute VB_Name = "SystemSetup"
Option Explicit
Public fso As Object
Public userProfilePath As String
Public Declare PtrSafe Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'https://msdn.microsoft.com/zh-tw/library/office/ff192913.aspx
Private Declare PtrSafe Function OpenClipboard Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32.dll" () As Long
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare PtrSafe Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare PtrSafe Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long

Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
  ByVal lpParameters As String, ByVal lpDirectory As String, _
  ByVal nShowCmd As Long) As Long 'https://www.mrexcel.com/board/threads/vba-api-call-issues-with-show-window-activation.920147/
Public Declare PtrSafe Function ShowWindow Lib "user32" _
  (ByVal hWnd As Long, ByVal nCmdSHow As Long) As Long
Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
  (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  
Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Boolean
  
  
  
'https://msdn.microsoft.com/zh-tw/library/office/ff194373.aspx
'Declare Function OpenClipboard Lib "User32" (ByVal hWnd As Long) _
'   As Long
'Declare Function CloseClipboard Lib "User32" () As Long
'Declare Function GetClipboardData Lib "User32" (ByVal wFormat As _
'   Long) As Long
'Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags&, ByVal _
'   dwBytes As Long) As Long
'Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) _
'   As Long
'Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) _
'   As Long
'Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) _
'   As Long
'Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
'   ByVal lpString2 As Any) As Long
 
Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096
Function ClipBoard_GetData()
   Dim hClipMemory As Long
   Dim lpClipMemory As Long
   Dim MyString As String
   Dim RetVal As Long
 
   If OpenClipboard(0&) = 0 Then
      MsgBox "Cannot open Clipboard. Another app. may have it open"
      Exit Function
   End If
          
   ' Obtain the handle to the global memory
   ' block that is referencing the text.
   hClipMemory = GetClipboardData(CF_TEXT)
   If IsNull(hClipMemory) Then
      MsgBox "Could not allocate memory"
      GoTo OutOfHere
   End If
 
   ' Lock Clipboard memory so we can reference
   ' the actual data string.
   lpClipMemory = GlobalLock(hClipMemory)
 
   If Not IsNull(lpClipMemory) Then
      MyString = space$(MAXSIZE)
      RetVal = lstrcpy(MyString, lpClipMemory)
      RetVal = GlobalUnlock(hClipMemory)
       
      ' Peel off the null terminating character.
      MyString = Mid(MyString, 1, InStr(1, MyString, Chr$(0), 0) - 1)
   Else
      MsgBox "Could not lock memory to copy string from."
   End If
 
OutOfHere:
 
   RetVal = CloseClipboard()
   ClipBoard_GetData = MyString
 
End Function
Public Sub SetClipboard(sUniText As String)
    Dim iStrPtr As Long
    Dim iLen As Long
    Dim iLock As Long
    Const GMEM_MOVEABLE As Long = &H2
    Const GMEM_ZEROINIT As Long = &H40
    Const CF_UNICODETEXT As Long = &HD
    OpenClipboard 0&
    EmptyClipboard
    iLen = LenB(sUniText) + 2&
    iStrPtr = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, iLen)
    iLock = GlobalLock(iStrPtr)
    lstrcpy iLock, CLng(StrPtr(sUniText)) 'http://forum.slime.com.tw/thread152795.html
    GlobalUnlock iStrPtr
    SetClipboardData CF_UNICODETEXT, iStrPtr
    CloseClipboard
End Sub

Public Function GetClipboard() As String
    Dim iStrPtr As Long
    Dim iLen As Long
    Dim iLock As Long
    Dim sUniText As String
    Const CF_UNICODETEXT As Long = 13&
    OpenClipboard 0&
    If IsClipboardFormatAvailable(CF_UNICODETEXT) Then
        iStrPtr = GetClipboardData(CF_UNICODETEXT)
        If iStrPtr Then
            iLock = GlobalLock(iStrPtr)
            iLen = GlobalSize(iStrPtr)
            sUniText = String$(iLen \ 2& - 1&, vbNullChar)
            lstrcpy CLng(StrPtr(sUniText)), iLock 'http://forum.slime.com.tw/thread152795.html
            GlobalUnlock iStrPtr
        End If
        GetClipboard = sUniText
    End If
    CloseClipboard
End Function

Sub CopyText(Text As String) 'https://stackoverflow.com/questions/14219455/excel-vba-code-to-copy-a-specific-string-to-clipboard
    'VBA Macro using late binding to copy text to clipboard.
    'By Justin Kay, 8/15/2014
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
'    Set MSForms_DataObject = New MSForms.DataObject �P�W���ۦP
'    MSForms_DataObject.Clear
    MSForms_DataObject.SetText Text
    MSForms_DataObject.PutInClipboard
    Set MSForms_DataObject = Nothing
End Sub

Function ClipboardPutIn(Optional StoreText As String) As String 'https://www.thespreadsheetguru.com/blog/2015/1/13/how-to-use-vba-code-to-copy-text-to-the-clipboard
'PURPOSE: Read/Write to Clipboard
'Source: ExcelHero.com (Daniel Ferry)

Dim x As Variant

'Store as variant for 64-bit VBA support
  x = StoreText

'Create HTMLFile Object
  With CreateObject("htmlfile")
    DoEvents
    With .parentWindow.clipboardData
      Select Case True
        Case Len(StoreText)
          'Write to the clipboard
            .setData "text", x
        Case Else
          'Read from the clipboard (no variable passed through)
            ClipboardPutIn = .GetData("text")
      End Select
    End With
  End With

End Function
Sub ���U���y��() 'ctrl+1 2008/7/23 F7'�쬰ToolsProofing
On Error Resume Next
    setOX
    OX.ControlSend "ScanGear CS-U", "", "Button2", "!S"
'    OX.WinActivate "�ϮѺ޲z"
'    OX.WinGetState "ScanGear CS-U"
    OX.WinSetState "ScanGear CS-U", "", OX.SW_MINIMIZE
    DoEvents
    OX.WinSetState "ScanGear CS-U", "", OX.SW_MINIMIZE
    'AppActivate "�ϮѺ޲z"
End Sub

Sub �d�ߩ_��() 'Ctrl+Shift+Y
On Error GoTo ErrMsg '�u�dgoogle
'FollowHyperlink "http://tw.search.yahoo.com/search", , , , "fr=slv1-ptec&p=" & Screen.ActiveControl.seltext
Selection.Copy
'FollowHyperlink "http://tw.search.yahoo.com/search", , , , "p=" & Selection, msoMethodGet
'If Tasks.Exists("skqs professional version") Then
    Shell Replace(GetDefaultBrowserEXE, """%1", "http://tw.search.yahoo.com/search?p=" & Selection)
'Else
'    Shell "C:\Program Files\Opera\opera.exe" & " http://tw.search.yahoo.com/search?p=" & Selection, vbNormalFocus
'End If
'���U���y��
'ActiveDocument.Save
Exit Sub
ErrMsg:
MsgBox Err & " : " & Err.Description
End Sub

Sub �d��Google()
'�ֳt��'Ctrl+shift+g'2011/8/11'2021/4/15�����w��w���r�Ʋέp�ΡA������w��Alt+Shift+g�BAlt+g
On Error GoTo ErrMsg
Const f As String = "�����j�M_���j�M-�P�ɷj�h�Ӥ���.EXE"
Const st As String = "C:\Program Files\�]�u�u\�����j�M_���j�M-�P�ɷj�h�Ӥ���\"
Dim funame As String
'FollowHyperlink "http://tw.search.yahoo.com/search", , , , "fr=slv1-ptec&p=" & Screen.ActiveControl.seltext
'FollowHyperlink "http://www.google.com.tw/search", , , , "q=" & Screen.ActiveControl.seltext, msoMethodGet
If Selection.Type = wdSelectionNormal Then
    Selection.Copy
    If ActiveDocument.Saved = False And ActiveDocument.path <> "" Then ActiveDocument.Save: DoEvents
    If Tasks.Exists("skqs professional version") Then
        Shell Replace(GetDefaultBrowserEXE, """%1", "http://www.google.com.tw/search?q=" & Selection)
    Else
        'Shell "C:\Program Files\Opera\opera.exe" & " http://www.google.com.tw/search?q=" & Selection, vbNormalFocus
        If Dir(st & f) <> "" Then
            funame = st & f
        ElseIf Dir("C:\Program Files (x86)\�]�u�u\�����j�M_���j�M-�P�ɷj�h�Ӥ���\" & f) <> "" Then
            funame = "C:\Program Files (x86)\�]�u�u\�����j�M_���j�M-�P�ɷj�h�Ӥ���\" & f
        ElseIf Dir("W:\!! for hpr\VB\�����j�M_���j�M-�P�ɷj�h�Ӥ���\�����j�M_���j�M-�P�ɷj�h�Ӥ���\bin\Debug\" & f) <> "" Then
            funame = "W:\!! for hpr\VB\�����j�M_���j�M-�P�ɷj�h�Ӥ���\�����j�M_���j�M-�P�ɷj�h�Ӥ���\bin\Debug\" & f
        ElseIf Dir("C:\�����j�M_���j�M-�P�ɷj�h�Ӥ���\�����j�M_���j�M-�P�ɷj�h�Ӥ���\bin\Debug\" & f) <> "" Then
            funame = "C:\�����j�M_���j�M-�P�ɷj�h�Ӥ���\�����j�M_���j�M-�P�ɷj�h�Ӥ���\bin\Debug\" & f
        ElseIf Dir("A:\", vbVolume) <> "" Then
            If Dir("A:\Users\oscar\Dropbox\VS\VB\�����j�M_���j�M-�P�ɷj�h�Ӥ���\�����j�M_���j�M-�P�ɷj�h�Ӥ���\bin\Debug\" & f) <> "" Then _
                funame = "A:\Users\oscar\Dropbox\VS\VB\�����j�M_���j�M-�P�ɷj�h�Ӥ���\�����j�M_���j�M-�P�ɷj�h�Ӥ���\bin\Debug\" & f
        ElseIf Dir(userProfilePath & "Dropbox\VS\VB\�����j�M_���j�M-�P�ɷj�h�Ӥ���\�����j�M_���j�M-�P�ɷj�h�Ӥ���\bin\Debug\" & f) <> "" Then
            funame = userProfilePath & "Dropbox\VS\VB\�����j�M_���j�M-�P�ɷj�h�Ӥ���\�����j�M_���j�M-�P�ɷj�h�Ӥ���\bin\Debug\" & f
        ElseIf Dir(userProfilePath & "Dropbox\VS\VB\�����j�M_���j�M-�P�ɷj�h�Ӥ���\�����j�M_���j�M-�P�ɷj�h�Ӥ���\bin\Debug\" & f) <> "" Then
            funame = userProfilePath & "Dropbox\VS\VB\�����j�M_���j�M-�P�ɷj�h�Ӥ���\�����j�M_���j�M-�P�ɷj�h�Ӥ���\bin\Debug\" & f
        Else
            Exit Sub
        End If
        Shell funame
        'Shell "W:\!! for hpr\VB\�����j�M_���j�M-�P�ɷj�h�Ӥ���\�����j�M_���j�M-�P�ɷj�h�Ӥ���\bin\Debug\�����j�M_���j�M-�P�ɷj�h�Ӥ���.exe"
    End If
End If
'���U���y��
Exit Sub

ErrMsg:
MsgBox Err & " : " & Err.Description
End Sub


Function ���o�ୱ���|() 'WshEnvironment.Item'2012/6/3

'GetDeskDir() '���o�ୱ
    'Dim wshshell As Object '�ŧiwshshell���@��Object
    Dim strDesktop As String 'strDesktop�ܼ��x�swshshell.regread���Ǧ^��
    'Set wshshell = CreateObject("wscript.shell") '�N"wscript.shell"���J��wshshell��
    'strDesktop = wshshell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Desktop") '���o�ୱ���|
    strDesktop = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    
'    Print "�ୱ���|���G"; strDesktop
���o�ୱ���| = strDesktop
'End Sub
'http://it-easy.tw/vb-get-path/#4

'Dim wshshell As Object
'Dim strDesktop
'Set wshshell = CreateObject("wscript.shell")
'strDesktop = wshshell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\desktop")
'http://www.accessoft.com/blog/article-show.asp?userid=32&Id=97
End Function
Function ���o�ϥΪ̸��|_�t�ϱ׽u() '2021/11/3
'https://www.796t.com/post/M2ExcmU=.html
'https://stackoverflow.com/questions/42091960/userprofile-environ-on-vba
Dim a As String
a = VBA.Environ("AppData")
a = VBA.Replace(a, "AppData\Roaming", "")
���o�ϥΪ̸��|_�t�ϱ׽u = a
End Function
Function GetClipboardText()
Dim clipboard As New MSForms.DataObject
DoEvents
clipboard.GetFromClipboard
GetClipboardText = clipboard.GetText
End Function

Sub insertNowTime()
With Selection.Range 'Alt+t
    .InsertAfter Now
    .Font.Subscript = True
End With
End Sub
Sub ���Ҥp�p��J�k() 'Alt+q
Shell Replace(SystemSetup.���o�ୱ���|, "Desktop", "Dropbox") & "\VS\bat\���Ҥp�p��J�k.bat"
End Sub

Sub shortcutKeys() '���w�ֳt��
CustomizationContext = NormalTemplate
'KeyBindings.Add _
    KeyCategory:=wdKeyCategoryCommand, _
    Command:="Docs.�b����󤤴M�����r��", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyPageDown)
KeyBindings.Add _
    KeyCategory:=wdKeyCategoryCommand, _
    Command:="Docs.�K�W�¤�r", _
    KeyCode:=BuildKeyCode(wdKeyShift, wdKeyInsert)
End Sub


'https://analystcave.com/vba-status-bar-progress-bar-sounds-emails-alerts-vba/#:~:text=The%20VBA%20Status%20Bar%20is%20a%20panel%20that,Bar%20we%20need%20to%20Enable%20it%20using%20Application.DisplayStatusBar%3A
Sub playSound(longShort As Single) 'Public Declare Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
    '�����n���B���ġB����
    Select Case longShort
        Case 1
            sndPlaySound32 "C:\Windows\Media\Chimes.wav", &H0
        Case 1.469
            sndPlaySound32 "C:\Windows\Media\Windows Message Nudge.wav", &H0
        Case 1.921
            sndPlaySound32 "C:\Windows\Media\Windows Notify System Generic.wav", &H0 '�H PotPlayer ����Y�i��M�椤�˵��^���ɦW
        Case 2
            sndPlaySound32 "C:\Windows\Media\Windows Notify Calendar.wav", &H0
        Case 3
            sndPlaySound32 "C:\Windows\Media\Alarm10.wav", &H0
        Case 4
            sndPlaySound32 "C:\Windows\Media\Alarm03.wav", &H0
        Case 7
            sndPlaySound32 "C:\Windows\Media\Ring10.wav", &H0
        Case 12
            sndPlaySound32 "C:\Windows\Media\Ring05.wav", &H0
    End Select
    
End Sub


Function getChrome()
Dim chromePath As String
If fso Is Nothing Then Set fso = CreateObject("scripting.filesystemobject")
If fso.fileexists("W:\PortableApps\PortableApps\GoogleChromePortable\App\Chrome-bin\chrome.exe") Then
    chromePath = "W:\PortableApps\PortableApps\GoogleChromePortable\App\Chrome-bin\chrome.exe"
ElseIf Dir("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe") <> "" Then
    chromePath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
ElseIf fso.fileexists("C:\Program Files\Google\Chrome\Application\chrome.exe") Then
    chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
End If
getChrome = chromePath
End Function

Sub stopUndo(ByRef ur As UndoRecord, Optional undoName As String)
'https://www.google.com/search?q=word+vba+stop+undo&rlz=1C1GCEU_zh-TWTW967TW967&oq=word+vba+stop+undo&aqs=chrome..69i57j69i64.10026j0j7&sourceid=chrome&ie=UTF-8
'https://docs.microsoft.com/en-us/office/vba/word/concepts/working-with-word/working-with-the-undorecord-object
'https://stackoverflow.com/questions/28051381/how-to-disable-the-changes-made-by-vba-in-undo-list-in-ms-word#_=_
'Dim ur As UndoRecord
Set ur = word.Application.UndoRecord
ur.StartCustomRecord undoName
'Set stopUndo = ur
End Sub

Sub contiUndo(ByRef ur As UndoRecord)
ur.EndCustomRecord
End Sub


Public Function appActivatedYet(exeName As String) As Boolean
On Error GoTo eh:
      exeName = exeName & ".exe": exeName = StrConv(exeName, vbUpperCase)
'https://stackoverflow.com/questions/44075292/determine-process-id-with-vba
'https://stackoverflow.com/questions/26277214/vba-getting-program-names-and-task-id-of-running-processes
    Dim objServices As Object, objProcessSet As Object, Process As Object

    Set objServices = GetObject("winmgmts:\\.\root\CIMV2")
    Set objProcessSet = objServices.ExecQuery("SELECT ProcessID, name FROM Win32_Process WHERE name = """ & exeName & """", , 48)

    'you may find more than one processid depending on your search/program
    For Each Process In objProcessSet
       'Debug.Print Process.ProcessID, Process.Name
       'If Process.Name = exeName Then 'processName Then
       If Not StrComp(Process.Name, exeName, vbTextCompare) Then 'processName Then
        appActivatedYet = True
        Exit Function
       End If
    Next
'    If objProcessSet.pri.Count > 0 Then appActivatedYet = True
    
    Set objProcessSet = Nothing
Exit Function
eh:
Select Case Err.Number
    Case 5 '�{�ǩI�s�Τ޼Ƥ����T
    Case Else
        MsgBox Err.Number & Err.Description
        'resume
End Select
End Function

Function apicShowWindow(strClassName As String, strWindowName As String, lngState As Long)
  'https://www.mrexcel.com/board/threads/vba-api-call-issues-with-show-window-activation.920147/
  'Declare variables
  Dim lngWnd As Long
  Dim intRet As Integer
  
  lngWnd = FindWindow(strClassName, strWindowName)
  apicShowWindow = ShowWindow(lngWnd, lngState)
  'Spy + + :https://docs.microsoft.com/zh-tw/visualstudio/debugger/how-to-start-spy-increment?view=vs-2022
  SetForegroundWindow lngWnd 'https://zechs.taipei/?p=146
End Function

Sub Wait(sec As Single)
'http://vbcity.com/forums/t/81315.aspx
Dim WaitDt As Date
WaitDt = DateAdd("s", sec, Now())
Do While Now < WaitDt
Loop
End Sub

Sub appActivateChrome()
    AppActivateDefaultBrowser 'https://docs.microsoft.com/zh-tw/sql/ado/reference/ado-api/absoluteposition-property-ado?view=sql-server-ver15
        'try looking for both Chrome_WidgetWin_1 and Chrome_RenderWidgetHostHWND
'    SystemSetup.apicShowWindow "Chrome_WidgetWin_1" _
        , vbNullString, 3 'https://zechs.taipei/?p=146
        'https://docs.microsoft.com/zh-tw/visualstudio/debugger/how-to-start-spy-increment?view=vs-2022
    'https://stackoverflow.com/questions/19705797/find-the-window-handle-for-a-chrome-browser

End Sub
