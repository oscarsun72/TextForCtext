Attribute VB_Name = "SystemSetup"
Option Explicit
Public FsO As Object, UserProfilePath As String
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

Public Property Get FileSystemObject() As Object
If FsO Is Nothing Then Set FsO = CreateObject("scripting.filesystemobject")
Set FileSystemObject = FsO
End Property

Public Property Get UserProfilePathIncldBackSlash() As String
    UserProfilePathIncldBackSlash = ���o�ϥΪ̸��|_�t�ϱ׽u
End Property
Public Property Get DropBoxPathIncldBackSlash() As String
    DropBoxPathIncldBackSlash = UserProfilePathIncldBackSlash & "Dropbox\"
End Property
Public Property Get UserAppDataRoamingPathIncldBackSlash() As String
    UserAppDataRoamingPathIncldBackSlash = UserProfilePathIncldBackSlash + "AppData\Roaming\"
End Property
Public Property Get WordTemplatesPathIncldBackSlash() As String
    WordTemplatesPathIncldBackSlash = UserProfilePathIncldBackSlash + "AppData\Roaming\Microsoft\Templates\"
End Property
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

'���o�ŶKï����r
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

Sub CopyText(text As String) 'https://stackoverflow.com/questions/14219455/excel-vba-code-to-copy-a-specific-string-to-clipboard
    'VBA Macro using late binding to copy text to clipboard.
    'By Justin Kay, 8/15/2014
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
'    Set MSForms_DataObject = New MSForms.DataObject �P�W���ۦP
'    MSForms_DataObject.Clear
    MSForms_DataObject.SetText text
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
selection.Copy
'FollowHyperlink "http://tw.search.yahoo.com/search", , , , "p=" & Selection, msoMethodGet
'If Tasks.Exists("skqs professional version") Then
    Shell Replace(GetDefaultBrowserEXE, """%1", "http://tw.search.yahoo.com/search?p=" & selection)
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
If selection.Type = wdSelectionNormal Then
    selection.Copy
    If ActiveDocument.Saved = False And ActiveDocument.path <> "" Then ActiveDocument.Save: DoEvents
'    If Tasks.Exists("skqs professional version") Then
'        Shell Replace(GetDefaultBrowserEXE, """%1", "http://www.google.com.tw/search?q=" & Selection)
'    Else
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
        ElseIf Dir(UserProfilePath & "Dropbox\VS\VB\�����j�M_���j�M-�P�ɷj�h�Ӥ���\�����j�M_���j�M-�P�ɷj�h�Ӥ���\bin\Debug\" & f) <> "" Then
            funame = UserProfilePath & "Dropbox\VS\VB\�����j�M_���j�M-�P�ɷj�h�Ӥ���\�����j�M_���j�M-�P�ɷj�h�Ӥ���\bin\Debug\" & f
        ElseIf Dir(UserProfilePath & "Dropbox\VS\VB\�����j�M_���j�M-�P�ɷj�h�Ӥ���\�����j�M_���j�M-�P�ɷj�h�Ӥ���\bin\Debug\" & f) <> "" Then
            funame = UserProfilePath & "Dropbox\VS\VB\�����j�M_���j�M-�P�ɷj�h�Ӥ���\�����j�M_���j�M-�P�ɷj�h�Ӥ���\bin\Debug\" & f
        Else
            Exit Sub
        End If
        Shell funame
        'Shell "W:\!! for hpr\VB\�����j�M_���j�M-�P�ɷj�h�Ӥ���\�����j�M_���j�M-�P�ɷj�h�Ӥ���\bin\Debug\�����j�M_���j�M-�P�ɷj�h�Ӥ���.exe"
'    End If
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
On Error GoTo eh
Dim clipboard As New MSForms.DataObject
DoEvents
clipboard.GetFromClipboard
GetClipboardText = clipboard.GetText
Exit Function
eh:
    Select Case Err.Number
        Case -2147221040 'DataObject:GetFromClipboard OpenClipboard ����
            SystemSetup.wait 0.8
            Resume
        Case Else
            MsgBox Err.Number + Err.Description
    End Select
End Function

Sub insertNowTime()
With selection.Range 'Alt+t
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
Sub playSound(longShort As Single, Optional waittoPlay As Byte = 1) 'Public Declare Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
    '�����n���B���ġB����'https://blog.csdn.net/xuemanqianshan/article/details/113485233
    Select Case longShort
        Case 1
            sndPlaySound32 "C:\Windows\Media\Chimes.wav", waittoPlay '&H1 '&H0:�������~���椧�᪺�{���X�A&H1=1�A�@����A���������A�Y���汵�U�Ӫ��{���X
        Case 1.294 'https://learn.microsoft.com/en-us/previous-versions/dd798676(v=vs.85)
            sndPlaySound32 "C:\Windows\Media\notify.wav", waittoPlay
        Case 1.469
            sndPlaySound32 "C:\Windows\Media\Windows Message Nudge.wav", waittoPlay
        Case 1.921
            sndPlaySound32 "C:\Windows\Media\Windows Notify System Generic.wav", waittoPlay '�H PotPlayer ����Y�i��M�椤�˵��^���ɦW
        Case 2
            sndPlaySound32 "C:\Windows\Media\Windows Notify Calendar.wav", waittoPlay
        Case 3
            sndPlaySound32 "C:\Windows\Media\Alarm10.wav", waittoPlay
        Case 4
            sndPlaySound32 "C:\Windows\Media\Alarm03.wav", waittoPlay
        Case 7
            sndPlaySound32 "C:\Windows\Media\Ring10.wav", waittoPlay
        Case 12
            sndPlaySound32 "C:\Windows\Media\Ring05.wav", waittoPlay
    End Select
    
End Sub

Property Get getChromePathIncludeBackslash() As String
Dim chromeFullname As String
chromeFullname = getChrome
getChromePathIncludeBackslash = VBA.Replace(chromeFullname, Dir(chromeFullname, vbDirectory), "")
End Property


Function getChrome() As String
Dim chromePath As String
If FsO Is Nothing Then Set FsO = CreateObject("scripting.filesystemobject")
If FsO.fileexists("W:\PortableApps\PortableApps\GoogleChromePortable\GoogleChromePortable.exe") Then '�γo�Ӥ~���|�M Selenium ���[
    chromePath = "W:\PortableApps\PortableApps\GoogleChromePortable\GoogleChromePortable.exe"
ElseIf FsO.fileexists("W:\PortableApps\PortableApps\GoogleChromePortable\App\Chrome-bin\chrome.exe") Then
    chromePath = "W:\PortableApps\PortableApps\GoogleChromePortable\App\Chrome-bin\chrome.exe"
ElseIf Dir("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe") <> "" Then
    chromePath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
ElseIf FsO.fileexists("C:\Program Files\Google\Chrome\Application\chrome.exe") Then
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
Set ur = Nothing
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

Sub wait(sec As Single)
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

Sub backupNormal_dotm() '�۰ʳƥ�Normal.dotm
'If ActiveDocument.path = "" Then Exit Sub
Dim source As String, destination As String
source = SystemSetup.DropBoxPathIncldBackSlash + "Normal.dotm"
destination = SystemSetup.WordTemplatesPathIncldBackSlash + "Normal.dotm"
On Error GoTo eh
With SystemSetup.FileSystemObject
If (.getfile(source).DateLastModified < _
    .getfile(destination).DateLastModified) Then _
        .CopyFile source, destination
End With
Exit Sub
eh:
Select Case Err.Number
    Case 70
        MsgBox "Normal.dotm�٥��ƥ�", vbExclamation
    Case Else
        MsgBox Err.Number & Err.Description
End Select
End Sub
Sub seleniumDllReference()
'20230119 creedit chatGPT�j���ġG
' �]�m dll �ɮשҦb�����|
Dim dllPath As String
dllPath = "C:\Users\oscar\AppData\Local\SeleniumBasic\Selenium.dll"

' �N dll ���[�J�Ѧ�
Dim ref 'As Reference
'Set ref = ThisWorkbook.VBProject.References.AddFromFile(dllPath)
Set ref = Templates("Normal").VBProject.References.AddFromFile(dllPath)
End Sub

Sub killProcessByName(imageName As String) 'ex: ChromeDriver.exe
Shell "taskkill /im " + imageName + " /f", vbHide
End Sub
Rem 20230119�GchatGPT�j���ġGVBA �ޱ� Chrome �s�����G
                Rem �b�z�����ϥ� Selenium Basic �w�ާ@ Chrome �s��������A�z�i�H�ϥ� VBA �{���X�ӵ��� chromedriver.exe �{�ǡC�z�i�H�ϥ� VBA �� "Shell" ��ƨӹF��o�@�ت��C
                Rem ����Ϊk�O:
                Rem Shell "taskkill /im chromedriver.exe /f", vbHide
                Rem �o�q�{���X�|�������b�B�檺 chromedriver.exe �{�ǡC
                Rem �䤤 /im �O���uimage name�v�A/f �O���u�j��v�C
Rem �]�i�H�ϥ� WMI �ӵ����B�椤��chromedriver.exe
'' WMI(Windows Management Instrumentation) �O Windows ���Ѫ��@�պ޲z�u��A�i�H�ΨӺ޲z�t�ΡB�����B�n��B�w�󵥡C�b�W�����{���X���A�ڭ̨ϥΤF WMI �Ӭd�ߩҦ��B�椤�� chromedriver.exe �{�ǡA�õ������̡C
'' �����ŧi�@�� WMI ����A�èϥ�ExecQuery��k�Ӭd�ߩҦ��B�椤�� chromedriver.exe �{�ǡC�M��ϥΰj��N�C�ӵ{�Ǫ�Terminate��k����A�o�˴N�i�H�����Ҧ��B�椤�� chromedriver.exe �{�ǤF�C
'' �z�i�H�ϥ� Windows Task Manager �Ӭd�ݭ��� chromedriver.exe �{�ǬO�ѱz�� VBA �{�ǱҰʪ��C�b Task Manager ���A�z�i�H�ݨ�C�ӹB�椤���{�Ǫ��ԲӸ�T�A�]�A�{�Ǫ��W�١BPID (?�{ ID) �H�Ω������ε{�ǡC
'' �z�i�H�bVBA���ϥ� WMI �Ӭd��Ҧ��B�椤�� chromedriver.exe �{�ǡA�ìd�ݥ��̪�PID�A�M��P�_���ǬO�ѱz�� VBA �{�ǱҰʪ��C
Sub killProcessesByName(imageName As String, pid As Long)
Dim objWMIService, objProcess, colProcess
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
'Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'chromedriver.exe'")
Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name = '" + imageName + "'")
For Each objProcess In colProcess
    'If objProcess.ProcessId = �AVBA�{���ұҰʪ�PID Then
    If objProcess.ProcessId = pid Then
        objProcess.Terminate
    End If
Next
End Sub
Sub killchromedriverFromHere()
Dim objWMIService, objProcess, colProcess, pid
If chromedriversPIDcntr = 0 Then
    ReDim chromedriversPID(0)
    If Not SeleniumOP.WD Is Nothing Then
        On Error GoTo eh:
        SeleniumOP.WD.Quit
        Set SeleniumOP.WD = Nothing
        Exit Sub
    End If
End If
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'chromedriver.exe'")
For Each objProcess In colProcess
    For Each pid In SeleniumOP.chromedriversPID
        If objProcess.ProcessId = pid And objProcess.ProcessId <> 0 Then objProcess.Terminate
    Next pid
Next
'���]�x�schromedriver�{��ID���}�C
ReDim chromedriversPID(0): chromedriversPIDcntr = 0: Set SeleniumOP.WD = Nothing

Exit Sub
eh:
Select Case Err.Number
    Case -2147467261 '�å��N����Ѧҳ]�w�����󪺰������C
        Resume Next
    Case Else
        MsgBox Err.Number + Err.Description
End Select
Rem 20230119 YouChat���ġG
Rem ���S���{�����禡��k�i�H�M�ΩO�H �u��� loop through �ܡH
Rem Yes, there are existing functions and methods that can be used to check if an element is included in an array. One such method is to use the InStr() function, which takes two arguments: the array and the element you are looking for. It then returns a non-zero value if the element is found in the array, or zero if it is not. However, this method is not recommended for large arrays, since it can slow down the execution of the code. The best way to check for an element in an array is to use a loop, as shown in the example above.

Rem     vba �������w���}�C����
Rem �A�i�H�ϥ� VBA ���� Erase ��ƨӲ������w���}�C�����A�Ҧp�G
Rem Dim myArray As String = {"A", "B", "C", "D"} Erase myArray(1), myArray(3)
Rem �o�� myArray �N�|�Q���s�w�q�� {"A", "C"}�C
Rem VBA �������w���ʺA�}�C����
Rem �p�G�A�Q�����ʺA�}�C�������w�����A�A�i�H�ϥ� VBA ���� Filter ��ơA�Ҧp�G
Rem Dim myDynamicArray As Variant myDynamicArray = Array("A", "B", "C", "D") myDynamicArray = Filter(myDynamicArray, "B", False)
Rem �o�� myDynamicArray �N�|�Q���s�w�q�� {"A", "C", "D"}�A�Ӥ��]�t "B" �o�Ӥ����C
Rem �p�G�ʺA�}�C�������O long ���O���O
Rem �p�G�ʺA�}�C�������O Long ���O���ܡA�A�i�H�ϥ� VBA ���� Erase ��ƨӲ������w�������A�Ҧp�G
Rem Dim myDynamicArray As Variant myDynamicArray = Array(100, 200, 300, 400) Erase myDynamicArray(2)
Rem �o�� myDynamicArray �N�|�Q���s�w�q�� {100, 200, 400}�A�Ӥ��]�t 300 �o�� Long ���O�������C
Rem VBA ������ C# ��List ���O ���e���]container�^ �i�H�ζ�
Rem ��AVBA ���]��?�� C# ���e���]container�^�i�H�ϥΡA���̥]�A Collection�BDictionary �M Array ���O�CCollection ���O�i�H�Ψ��x�s��������������A�� Dictionary ���O�h�i�H�Ψ��x�s��ȹ�AArray ���O�h�i�H�Ψ��x�s�s�򪺼ƭȡC�Ҧp�A�A�i�H�ϥ� Array ���O���x�s�@�ӦW�� myArray ���ƲաA�䤤�]�t 1�B2 �M 3�G
Rem Dim myArray As Variant
Rem �·нХܽd�@�U�bVBA�� collection �n���s�W�P��������
Rem �A�i�H�ϥ� VBA ���� Add ��k�ӦV Collection ���O�s�W�����A�Ҧp�G
Rem Dim myCollection As New Collection myCollection.Add "A" myCollection.Add "B"
Rem �P�˦a�A�A�]�i�H�ϥ� Remove ��k�Ӳ��� Collection ���O�������w�����A�Ҧp�G
Rem myCollection.Remove 1
Rem �o�� myCollection �N�|�Q���s�w�q�� {"B"}�A���]�t "A" �o�Ӥ����C
Rem ���b�䤤�d��@�ӯS�w�����A�����O
Rem �A�i�H�ϥ� Find ��k�Ӭd�� Collection ���O�������w�����A�Ҧp�G
Rem Dim myCollection As New Collection myCollection.Add "A" myCollection.Add "B"
Rem Dim myElement As Variant myElement = myCollection.Find("B")
Rem �p�G�����w�����AmyElement ���ȴN�|�� "B"�A���ۧA�N�i�H�ϥ� Remove ��k�Ӳ����o�Ӥ����G
Rem myCollection.Remove myElement
Rem �o�� myCollection �N�|�Q���s�w�q�� {"A"}�A���]�t "B" �o�Ӥ����C
Rem �u���D�`�P��YouChat���ız �P���P���@�g���g�ۡ@�n�L��������
Rem �P�±z�����y�I�n�L��������I

End Sub

Sub killchromedrivers()
Dim objWMIService, objProcess, colProcess
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'chromedriver.exe'")
For Each objProcess In colProcess
    objProcess.Terminate
Next
End Sub


Rem VBA ���o�t�������ܼƤίS���Ƨ����| �pC#���� Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) �^�ǭ�
Rem ���i�H�ϥ� Environ ��ƨӨ��o�t�������ܼƪ��ȡC�Ҧp�A�n���o�������ε{����Ƹ��|�A�i�H�ϥΤU�����{���X: localAppData = Environ("LOCALAPPDATA")
'' �p�G�n���o��L�S���Ƨ������| , �i�H�ϥΤU�����{���X:
'' desktop = Environ("USERPROFILE") & "\Desktop"
'' �i�Ϊ��S���Ƨ����|�i�H�bMSDN�W���C



