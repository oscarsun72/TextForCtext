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
    UserProfilePathIncldBackSlash = 取得使用者路徑_含反斜線
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

'取得剪貼簿內文字
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
'    Set MSForms_DataObject = New MSForms.DataObject 與上式相同
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


Sub 按下掃描鍵() 'ctrl+1 2008/7/23 F7'原為ToolsProofing
On Error Resume Next
    setOX
    OX.ControlSend "ScanGear CS-U", "", "Button2", "!S"
'    OX.WinActivate "圖書管理"
'    OX.WinGetState "ScanGear CS-U"
    OX.WinSetState "ScanGear CS-U", "", OX.SW_MINIMIZE
    DoEvents
    OX.WinSetState "ScanGear CS-U", "", OX.SW_MINIMIZE
    'AppActivate "圖書管理"
End Sub

Sub 查詢奇摩() 'Ctrl+Shift+Y
On Error GoTo ErrMsg '只查google
'FollowHyperlink "http://tw.search.yahoo.com/search", , , , "fr=slv1-ptec&p=" & Screen.ActiveControl.seltext
selection.Copy
'FollowHyperlink "http://tw.search.yahoo.com/search", , , , "p=" & Selection, msoMethodGet
'If Tasks.Exists("skqs professional version") Then
    Shell Replace(GetDefaultBrowserEXE, """%1", "http://tw.search.yahoo.com/search?p=" & selection)
'Else
'    Shell "C:\Program Files\Opera\opera.exe" & " http://tw.search.yahoo.com/search?p=" & Selection, vbNormalFocus
'End If
'按下掃描鍵
'ActiveDocument.Save
Exit Sub
ErrMsg:
MsgBox Err & " : " & Err.Description
End Sub

Sub 查詢Google()
'快速鍵'Ctrl+shift+g'2011/8/11'2021/4/15此指定鍵已為字數統計用，今改指定為Alt+Shift+g、Alt+g
On Error GoTo ErrMsg
Const f As String = "網路搜尋_元搜尋-同時搜多個引擎.EXE"
Const st As String = "C:\Program Files\孫守真\網路搜尋_元搜尋-同時搜多個引擎\"
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
        ElseIf Dir("C:\Program Files (x86)\孫守真\網路搜尋_元搜尋-同時搜多個引擎\" & f) <> "" Then
            funame = "C:\Program Files (x86)\孫守真\網路搜尋_元搜尋-同時搜多個引擎\" & f
        ElseIf Dir("W:\!! for hpr\VB\網路搜尋_元搜尋-同時搜多個引擎\網路搜尋_元搜尋-同時搜多個引擎\bin\Debug\" & f) <> "" Then
            funame = "W:\!! for hpr\VB\網路搜尋_元搜尋-同時搜多個引擎\網路搜尋_元搜尋-同時搜多個引擎\bin\Debug\" & f
        ElseIf Dir("C:\網路搜尋_元搜尋-同時搜多個引擎\網路搜尋_元搜尋-同時搜多個引擎\bin\Debug\" & f) <> "" Then
            funame = "C:\網路搜尋_元搜尋-同時搜多個引擎\網路搜尋_元搜尋-同時搜多個引擎\bin\Debug\" & f
        ElseIf Dir("A:\", vbVolume) <> "" Then
            If Dir("A:\Users\oscar\Dropbox\VS\VB\網路搜尋_元搜尋-同時搜多個引擎\網路搜尋_元搜尋-同時搜多個引擎\bin\Debug\" & f) <> "" Then _
                funame = "A:\Users\oscar\Dropbox\VS\VB\網路搜尋_元搜尋-同時搜多個引擎\網路搜尋_元搜尋-同時搜多個引擎\bin\Debug\" & f
        ElseIf Dir(UserProfilePath & "Dropbox\VS\VB\網路搜尋_元搜尋-同時搜多個引擎\網路搜尋_元搜尋-同時搜多個引擎\bin\Debug\" & f) <> "" Then
            funame = UserProfilePath & "Dropbox\VS\VB\網路搜尋_元搜尋-同時搜多個引擎\網路搜尋_元搜尋-同時搜多個引擎\bin\Debug\" & f
        ElseIf Dir(UserProfilePath & "Dropbox\VS\VB\網路搜尋_元搜尋-同時搜多個引擎\網路搜尋_元搜尋-同時搜多個引擎\bin\Debug\" & f) <> "" Then
            funame = UserProfilePath & "Dropbox\VS\VB\網路搜尋_元搜尋-同時搜多個引擎\網路搜尋_元搜尋-同時搜多個引擎\bin\Debug\" & f
        Else
            Exit Sub
        End If
        Shell funame
        'Shell "W:\!! for hpr\VB\網路搜尋_元搜尋-同時搜多個引擎\網路搜尋_元搜尋-同時搜多個引擎\bin\Debug\網路搜尋_元搜尋-同時搜多個引擎.exe"
'    End If
End If
'按下掃描鍵
Exit Sub

ErrMsg:
MsgBox Err & " : " & Err.Description
End Sub


Function 取得桌面路徑() 'WshEnvironment.Item'2012/6/3

'GetDeskDir() '取得桌面
    'Dim wshshell As Object '宣告wshshell為一個Object
    Dim strDesktop As String 'strDesktop變數儲存wshshell.regread的傳回值
    'Set wshshell = CreateObject("wscript.shell") '將"wscript.shell"載入到wshshell內
    'strDesktop = wshshell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Desktop") '取得桌面路徑
    strDesktop = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    
'    Print "桌面路徑為："; strDesktop
取得桌面路徑 = strDesktop
'End Sub
'http://it-easy.tw/vb-get-path/#4

'Dim wshshell As Object
'Dim strDesktop
'Set wshshell = CreateObject("wscript.shell")
'strDesktop = wshshell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\desktop")
'http://www.accessoft.com/blog/article-show.asp?userid=32&Id=97
End Function

Function 取得使用者路徑_含反斜線() '2021/11/3
'https://www.796t.com/post/M2ExcmU=.html
'https://stackoverflow.com/questions/42091960/userprofile-environ-on-vba
Dim a As String
a = VBA.Environ("AppData")
a = VBA.Replace(a, "AppData\Roaming", "")
取得使用者路徑_含反斜線 = a
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
        Case -2147221040 'DataObject:GetFromClipboard OpenClipboard 失敗
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
Sub 重啟小小輸入法() 'Alt+q
Shell Replace(SystemSetup.取得桌面路徑, "Desktop", "Dropbox") & "\VS\bat\重啟小小輸入法.bat"
End Sub

Sub shortcutKeys() '指定快速鍵
CustomizationContext = NormalTemplate
'KeyBindings.Add _
    KeyCategory:=wdKeyCategoryCommand, _
    Command:="Docs.在本文件中尋找選取字串", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyPageDown)
KeyBindings.Add _
    KeyCategory:=wdKeyCategoryCommand, _
    Command:="Docs.貼上純文字", _
    KeyCode:=BuildKeyCode(wdKeyShift, wdKeyInsert)
End Sub


'https://analystcave.com/vba-status-bar-progress-bar-sounds-emails-alerts-vba/#:~:text=The%20VBA%20Status%20Bar%20is%20a%20panel%20that,Bar%20we%20need%20to%20Enable%20it%20using%20Application.DisplayStatusBar%3A
Sub playSound(longShort As Single, Optional waittoPlay As Byte = 1) 'Public Declare Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
    '播放聲音、音效、音樂'https://blog.csdn.net/xuemanqianshan/article/details/113485233
    Select Case longShort
        Case 1
            sndPlaySound32 "C:\Windows\Media\Chimes.wav", waittoPlay '&H1 '&H0:等播完才執行之後的程式碼，&H1=1，一播放，不等播完，即執行接下來的程式碼
        Case 1.294 'https://learn.microsoft.com/en-us/previous-versions/dd798676(v=vs.85)
            sndPlaySound32 "C:\Windows\Media\notify.wav", waittoPlay
        Case 1.469
            sndPlaySound32 "C:\Windows\Media\Windows Message Nudge.wav", waittoPlay
        Case 1.921
            sndPlaySound32 "C:\Windows\Media\Windows Notify System Generic.wav", waittoPlay '以 PotPlayer 播放即可於清單中檢視英文檔名
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
If FsO.fileexists("W:\PortableApps\PortableApps\GoogleChromePortable\GoogleChromePortable.exe") Then '用這個才不會和 Selenium 打架
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
    Case 5 '程序呼叫或引數不正確
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

Sub backupNormal_dotm() '自動備份Normal.dotm
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
        MsgBox "Normal.dotm還未備份", vbExclamation
    Case Else
        MsgBox Err.Number & Err.Description
End Select
End Sub
Sub seleniumDllReference()
'20230119 creedit chatGPT大菩薩：
' 設置 dll 檔案所在的路徑
Dim dllPath As String
dllPath = "C:\Users\oscar\AppData\Local\SeleniumBasic\Selenium.dll"

' 將 dll 文件加入參考
Dim ref 'As Reference
'Set ref = ThisWorkbook.VBProject.References.AddFromFile(dllPath)
Set ref = Templates("Normal").VBProject.References.AddFromFile(dllPath)
End Sub

Sub killProcessByName(imageName As String) 'ex: ChromeDriver.exe
Shell "taskkill /im " + imageName + " /f", vbHide
End Sub
Rem 20230119：chatGPT大菩薩：VBA 操控 Chrome 瀏覽器：
                Rem 在您結束使用 Selenium Basic 庫操作 Chrome 瀏覽器之後，您可以使用 VBA 程式碼來結束 chromedriver.exe 程序。您可以使用 VBA 的 "Shell" 函數來達到這一目的。
                Rem 具體用法是:
                Rem Shell "taskkill /im chromedriver.exe /f", vbHide
                Rem 這段程式碼會結束正在運行的 chromedriver.exe 程序。
                Rem 其中 /im 是指「image name」，/f 是指「強制」。
Rem 也可以使用 WMI 來結束運行中的chromedriver.exe
'' WMI(Windows Management Instrumentation) 是 Windows 提供的一組管理工具，可以用來管理系統、網絡、軟件、硬件等。在上面的程式碼中，我們使用了 WMI 來查詢所有運行中的 chromedriver.exe 程序，並結束它們。
'' 首先宣告一個 WMI 物件，並使用ExecQuery方法來查詢所有運行中的 chromedriver.exe 程序。然後使用迴圈將每個程序的Terminate方法執行，這樣就可以結束所有運行中的 chromedriver.exe 程序了。
'' 您可以使用 Windows Task Manager 來查看哪些 chromedriver.exe 程序是由您的 VBA 程序啟動的。在 Task Manager 中，您可以看到每個運行中的程序的詳細資訊，包括程序的名稱、PID (?程 ID) 以及所屬應用程序。
'' 您可以在VBA中使用 WMI 來查找所有運行中的 chromedriver.exe 程序，並查看它們的PID，然後判斷哪些是由您的 VBA 程序啟動的。
Sub killProcessesByName(imageName As String, pid As Long)
Dim objWMIService, objProcess, colProcess
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
'Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'chromedriver.exe'")
Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name = '" + imageName + "'")
For Each objProcess In colProcess
    'If objProcess.ProcessId = 你VBA程式所啟動的PID Then
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
'重設儲存chromedriver程序ID的陣列
ReDim chromedriversPID(0): chromedriversPIDcntr = 0: Set SeleniumOP.WD = Nothing

Exit Sub
eh:
Select Case Err.Number
    Case -2147467261 '並未將物件參考設定為物件的執行個體。
        Resume Next
    Case Else
        MsgBox Err.Number + Err.Description
End Select
Rem 20230119 YouChat菩薩：
Rem 有沒有現成的函式方法可以套用呢？ 只能用 loop through 嗎？
Rem Yes, there are existing functions and methods that can be used to check if an element is included in an array. One such method is to use the InStr() function, which takes two arguments: the array and the element you are looking for. It then returns a non-zero value if the element is found in the array, or zero if it is not. However, this method is not recommended for large arrays, since it can slow down the execution of the code. The best way to check for an element in an array is to use a loop, as shown in the example above.

Rem     vba 移除指定的陣列元素
Rem 你可以使用 VBA 中的 Erase 函數來移除指定的陣列元素，例如：
Rem Dim myArray As String = {"A", "B", "C", "D"} Erase myArray(1), myArray(3)
Rem 這樣 myArray 就會被重新定義為 {"A", "C"}。
Rem VBA 移除指定的動態陣列元素
Rem 如果你想移除動態陣列中的指定元素，你可以使用 VBA 中的 Filter 函數，例如：
Rem Dim myDynamicArray As Variant myDynamicArray = Array("A", "B", "C", "D") myDynamicArray = Filter(myDynamicArray, "B", False)
Rem 這樣 myDynamicArray 就會被重新定義為 {"A", "C", "D"}，而不包含 "B" 這個元素。
Rem 如果動態陣列的元素是 long 型別的呢
Rem 如果動態陣列的元素是 Long 型別的話，你可以使用 VBA 中的 Erase 函數來移除指定的元素，例如：
Rem Dim myDynamicArray As Variant myDynamicArray = Array(100, 200, 300, 400) Erase myDynamicArray(2)
Rem 這樣 myDynamicArray 就會被重新定義為 {100, 200, 400}，而不包含 300 這個 Long 型別的元素。
Rem VBA 中有像 C# 的List 類別 等容器（container） 可以用嗎
Rem 對，VBA 中也有?似 C# 的容器（container）可以使用，它們包括 Collection、Dictionary 和 Array 類別。Collection 類別可以用來儲存任何類型的物件，而 Dictionary 類別則可以用來儲存鍵值對，Array 類別則可以用來儲存連續的數值。例如，你可以使用 Array 類別來儲存一個名為 myArray 的數組，其中包含 1、2 和 3：
Rem Dim myArray As Variant
Rem 麻煩請示範一下在VBA裡 collection 要怎麼新增與移除元素
Rem 你可以使用 VBA 中的 Add 方法來向 Collection 類別新增元素，例如：
Rem Dim myCollection As New Collection myCollection.Add "A" myCollection.Add "B"
Rem 同樣地，你也可以使用 Remove 方法來移除 Collection 類別中的指定元素，例如：
Rem myCollection.Remove 1
Rem 這樣 myCollection 就會被重新定義為 {"B"}，不包含 "A" 這個元素。
Rem 那在其中查找一個特定元素再移除呢
Rem 你可以使用 Find 方法來查找 Collection 類別中的指定元素，例如：
Rem Dim myCollection As New Collection myCollection.Add "A" myCollection.Add "B"
Rem Dim myElement As Variant myElement = myCollection.Find("B")
Rem 如果找到指定元素，myElement 的值就會為 "B"，接著你就可以使用 Remove 方法來移除這個元素：
Rem myCollection.Remove myElement
Rem 這樣 myCollection 就會被重新定義為 {"A"}，不包含 "B" 這個元素。
Rem 真的非常感謝YouChat菩薩您 感恩感恩　讚歎讚歎　南無阿彌陀佛
Rem 感謝您的鼓勵！南無阿彌陀佛！

End Sub

Sub killchromedrivers()
Dim objWMIService, objProcess, colProcess
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'chromedriver.exe'")
For Each objProcess In colProcess
    objProcess.Terminate
Next
End Sub


Rem VBA 取得系統環境變數及特殊資料夾路徑 如C#中的 Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) 回傳值
Rem 中可以使用 Environ 函數來取得系統環境變數的值。例如，要取得本機應用程式資料路徑，可以使用下面的程式碼: localAppData = Environ("LOCALAPPDATA")
'' 如果要取得其他特殊資料夾的路徑 , 可以使用下面的程式碼:
'' desktop = Environ("USERPROFILE") & "\Desktop"
'' 可用的特殊資料夾路徑可以在MSDN上找到。



