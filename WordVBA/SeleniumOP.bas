Attribute VB_Name = "SeleniumOP"
Option Explicit
Public wd As SeleniumBasic.IWebDriver
Public chromedriversPID() As Long '�x�schromedriver�{��ID���}�C
Public chromedriversPIDcntr As Integer 'chromedriversPID���U�Э�
Public ActiveXComponentsCanNotBeCreated As Boolean

'Sub tesSeleniumBasic() 'https://github.com/florentbr/SeleniumBasic
''20230119 creedit chatGPT�j����
'
'    Dim driver As New Selenium.WebDriver
'    'driver.start "chrome", "https://www.google.com"
'    driver.SetBinary getChromePathIncludeBackslash
'    driver.start getChromePathIncludeBackslash + "chrome.exe", "https://www.google.com"
'    driver.Get "/"
'
'End Sub
Rem ���ѮɶǦ^false
Function openNewTabWhenTabAlreadyExit(ByVal wd As SeleniumBasic.IWebDriver) As Boolean
    On Error GoTo eH
    Dim iw As Byte, ew, ii As Byte
reOpenChrome:
    For Each ew In wd.WindowHandles
        iw = iw + 1
    Next ew
    If iw > 0 Then
          wd.ExecuteScript "window.open('about:blank','_blank');"
          For Each ew In wd.WindowHandles
                ii = ii + 1
                If ii = iw + 1 Then Exit For
          Next ew
          wd.SwitchTo().Window (ew)
    End If
    Exit Function
eH:
    Select Case Err.Number
        Case -2146233088
            If InStr(Err.Description, "no such window: target window already closed") Then
                If iw > 0 Then
                    For Each ew In wd.WindowHandles
                        Exit For
                    Next ew
                    wd.SwitchTo.Window (ew)
                    Resume
                Else
                    Stop
                End If
            ElseIf InStr(Err.Description, "ot connected to DevTools") Then
'                disconnected: not connected to DevTools
'                (failed to check if window was closed: disconnected: not connected to DevTools)
'                (Session info: chrome=127.0.6533.120)
                If Not wd Is Nothing Then wd.Quit
                Set wd = Nothing
                killchromedriverFromHere
                MsgBox "�YChrome�s�����w�}�ҡA������Chrome�s������A���T�w", vbExclamation
                openChrome "https://www.google.com"
                Resume 'GoTo reOpenChrome:
            ElseIf InStr(Err.Description, "A exception with a null response was thrown sending an HTTP") Then
'                A exception with a null response was thrown sending an HTTP request to the remote WebDriver server for URL http://localhost:1760/session/ed5864479325c154783256563f97e610/window/handles. The status of the exception was ConnectFailure, and the message was: �L�k�s���ܻ��ݦ��A��
                Set wd = Nothing
                killchromedriverFromHere
                MsgBox "�YChrome�s�����w�}�ҡA������Chrome�s������A���T�w", vbExclamation
'                openChrome "https://www.google.com"
'                Resume 'GoTo reOpenChrome:
                openNewTabWhenTabAlreadyExit = False
            Else
                MsgBox Err.Number & Err.Description
                Stop
            End If
        Case -2147467261
            If Err.Description = "�å��N����Ѧҳ]�w�����󪺰������C" Then
                Set wd = Nothing
                killchromedriverFromHere
                MsgBox "�YChrome�s�����w�}�ҡA������Chrome�s������A����@��", vbExclamation
                openNewTabWhenTabAlreadyExit = False
            Else
                Stop
            End If
        Case 91
            If Err.Description = "�S���]�w�����ܼƩ� With �϶��ܼ�" Then
                Set wd = Nothing
                killchromedriverFromHere
                MsgBox "�YChrome�s�����w�}�ҡA������Chrome�s������A����@��", vbExclamation
                openNewTabWhenTabAlreadyExit = False
            Else
                Stop
            End If
        Case Else
            MsgBox Err.Description, vbCritical
            wd.Quit
            SystemSetup.killchromedriverFromHere
    '           Resume
    End Select
End Function
Rem �Ұ�Chrome�s�����Τw�Ұʫ�}�ҷs�����s���C���ѮɶǦ^false
Function openChrome(Optional URL As String) As Boolean
reStart:
        'Dim WD As SeleniumBasic.IWebDriver
        On Error GoTo ErrH
        Dim Service As SeleniumBasic.ChromeDriverService
        Dim Options As SeleniumBasic.ChromeOptions
        Dim pid As Long
    
    '����chromedriver.exe
    '�ϥ� WMI �M�W���ҭz����k
    '�P�_PID�O�_����pid
    
        If wd Is Nothing Then
            Set wd = New SeleniumBasic.IWebDriver
            Set Service = New SeleniumBasic.ChromeDriverService
                
                Dim chromePath As String
                chromePath = getChromePathIncludeBackslash
                If InStr(chromePath, "GoogleChromePortable") Then
                    chromePath = "W:\PortableApps\PortableApps\GoogleChromePortable\App\Chrome-bin\"
                End If
    
            With Service
                .CreateDefaultService driverPath:=chromePath 'getChromePathIncludeBackslash
                '.CreateDefaultService driverPath:="E:\Selenium\Drivers"
                .HideCommandPromptWindow = True '����ܩR�O���ܦr������
            End With
            Set Options = New SeleniumBasic.ChromeOptions
            With Options
                .BinaryLocation = chromePath + "chrome.exe"
                .AddExcludedArgument "enable-automation" '�T�ΡuChrome ���b�Q�۰ʤƳn�鱱��v��ĵ�i����
                
                'C#�Goptions.AddArgument("user-data-dir=" + Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\Google\\Chrome\\User Data\\");
                .AddArgument "user-data-dir=" + VBA.Environ("LOCALAPPDATA") + _
                    "\Google\Chrome\User Data\"
    '            .AddArgument "--new-window"
                '.AddArgument "--start-maximized"
                '.DebuggerAddress = "127.0.0.1:9999" '���n�O��L�L�ӲV��
            End With
            wd.New_ChromeDriver Service:=Service, Options:=Options
            Docs.Register_Event_Handler '���M��chromedriver�@�ǳ�
            pid = Service.ProcessId 'Chrome�s�����S���}���\�N�|�O0
            If pid <> 0 Then
                ReDim Preserve chromedriversPID(chromedriversPIDcntr)
                chromedriversPID(chromedriversPIDcntr) = pid
                chromedriversPIDcntr = chromedriversPIDcntr + 1
            End If
            wd.ExecuteScript "window.open('about:blank','_blank');" 'openNewTabWhenTabAlreadyExit WD
            wd.SwitchTo.Window WindowHandlesItem(WindowHandlesCount - 1)
            wd.URL = URL
        Else
            wd.ExecuteScript "window.open('about:blank','_blank');" 'openNewTabWhenTabAlreadyExit WD
            wd.SwitchTo.Window WindowHandlesItem(WindowHandlesCount - 1)
            wd.URL = URL
        End If
        If ActiveXComponentsCanNotBeCreated Then ActiveXComponentsCanNotBeCreated = False
        openChrome = True
    Exit Function
ErrH:
    Select Case Err.Number
        Case 49
            If Err.Description = "DLL �I�s�W����~" Then
'                WD.Quit
'                killchromedriverFromHere
                Stop
                Resume
            End If
        Case -2146233079
            If VBA.Left(Err.Description, Len("session not created: Chrome failed to start: exited normally.")) = "session not created: Chrome failed to start: exited normally." Then
                wd.Quit
                SystemSetup.killchromedriverFromHere
                Set wd = Nothing
                Stop
                If MsgBox("������Chrome�s�����A�~��I" & vbCr & vbCr & _
                    vbTab & "�O�_�n�{���۰����z�����B�ҰʡC�P���P���@�n�L��������", vbCritical + vbOKCancel) _
                        = vbOK Then
                    SystemSetup.killProcessesByName "chrome.exe"
                    GoTo reStart
                Else
                    openChrome = False
                End If
                
                Exit Function
            End If
        Case -2146233088 '**'
            'Debug.Print Err.Description
            If InStr(Err.Description, "Chrome failed to start: exited normally.") Then
                '' err.Descriptionunknown error: Chrome failed to start: exited normally.
                ''  (unknown error: DevToolsActivePort file doesn't exist)
                '' (The process started from chrome location W:\PortableApps\PortableApps\GoogleChromePortable\App\Chrome-bin\chrome.exe is no longer running, so ChromeDriver is assuming that Chrome has crashed.)
                If MsgBox("���������e�}�Ҫ�Chrome�s�����A�~��", vbExclamation + vbOKCancel) = vbOK Then
                        'killProcessesByName "ChromeDriver.exe", pid
                        killchromedriverFromHere
                    GoTo reStart
                Else
        '            WD.Quit
                    killchromedriverFromHere
                End If
            ElseIf InStr(Err.Description, "no such window: No target with given id found") Then
                killchromedriverFromHere
                GoTo reStart
            ElseIf InStr(Err.Description, "disconnected: received Inspector.detached event") Then '(failed to check if window was closed: disconnected: not connected to DevTools)
                                                                                                    '(Session info: chrome=110.0.5481.178)
                killchromedriverFromHere
                GoTo reStart
            ElseIf InStr(Err.Description, "no such window: target window already closed") Then 'no such window: target window already closed
                                                                                                        'from unknown error: web view not found
                                                                                                         ' (Session info: chrome=128.0.6613.85)
'                Stop
                wd.SwitchTo.Window WindowHandlesItem(WindowHandlesCount - 1)
                Resume
                '�^�� wd.ExecuteScript "window.open('about:blank','_blank');" 'openNewTabWhenTabAlreadyExit WD
                     'wd.SwitchTo.Window WindowHandlesItem(WindowHandlesCount - 1)
            Else
                
                MsgBox Err.Description, vbCritical
                Stop
            End If
        Case 429 'ActiveX ����L�k���ͪ���'
            ActiveXComponentsCanNotBeCreated = True
            Exit Function
        Case -2147467261
            If InStr(Err.Description, "�å��N����Ѧҳ]�w�����󪺰������C") Then
                SystemSetup.killchromedriverFromHere
                Set wd = Nothing
'                Stop
                If MsgBox("������Chrome�s�����A�~��I" & vbCr & vbCr & _
                    vbTab & "�O�_�n�{���۰����z�����B�ҰʡC�P���P���@�n�L��������", vbCritical + vbOKCancel) _
                        = vbOK Then
                    SystemSetup.killProcessesByName "chrome.exe"
                    GoTo reStart
                Else
                    openChrome = False
                End If
                Exit Function
            Else
                MsgBox Err.Description, vbCritical
                Stop
            End If
        Case Else
            MsgBox Err.Description, vbCritical
            Resume
    End Select

End Function

Function openChromeBackground(URL As String) As SeleniumBasic.IWebDriver
reStart:
    'Dim WD As SeleniumBasic.IWebDriver
    On Error GoTo ErrH
    Dim wd As SeleniumBasic.IWebDriver
    Dim Service As SeleniumBasic.ChromeDriverService
    Dim Options As SeleniumBasic.ChromeOptions
    Dim pid As Long
    
        Set wd = New SeleniumBasic.IWebDriver
        Set Service = New SeleniumBasic.ChromeDriverService
        
        Dim chromePath As String
        chromePath = getChromePathIncludeBackslash
        If InStr(chromePath, "GoogleChromePortable") Then
            chromePath = "W:\PortableApps\PortableApps\GoogleChromePortable\App\Chrome-bin\"
        End If
        
        With Service
            .CreateDefaultService driverPath:=chromePath 'getChromePathIncludeBackslash
            .HideCommandPromptWindow = True '����ܩR�O���ܦr������
            If chromedriversPIDcntr = 0 Then chromedriversPIDcntr = 1
            ReDim chromedriversPID(chromedriversPIDcntr - 1)
'            chromedriversPID(chromedriversPIDcntr - 1) = Service.ProcessId'�٥��Ұ�=0
        End With
        
        Set Options = New SeleniumBasic.ChromeOptions
        With Options
            '.BinaryLocation = getChromePathIncludeBackslash + "chrome.exe"
            .BinaryLocation = chromePath + "chrome.exe"
            
            .AddExcludedArgument "enable-automation" '�T�ΡuChrome ���b�Q�۰ʤƳn�鱱��v��ĵ�i����
            
            'C#�Goptions.AddArgument("user-data-dir=" + Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\Google\\Chrome\\User Data\\");
            .AddArgument "user-data-dir=" + VBA.Environ("LOCALAPPDATA") + _
                "\Google\Chrome\User Data\"
            .AddArgument "--headless" '����ܹ���A�Y�ݤ���Chrome�s�����A�L�k��ʾާ@�κʱ�
            .AddArgument "--disable-gpu"
            .AddArgument "--disable-infobars"
            .AddArgument "--disable-extensions"
            .AddArgument "--disable-dev-shm-usage"
            '.AddArgument "--start-maximized"
            '.DebuggerAddress = "127.0.0.1:9999" '���n�O��L�L�ӲV��
        End With
        wd.New_ChromeDriver Service:=Service, Options:=Options
        'WD.Quit �|�۰ʲM��chromedriver�A�N���ΰO�U�}�L���ǤF
'        pid = Service.ProcessId 'Chrome�s�����S���}���\�N�|�O0
'        If pid <> 0 Then
'            ReDim Preserve chromedriversPID(chromedriversPIDcntr)
'            chromedriversPID(chromedriversPIDcntr) = pid
'            chromedriversPIDcntr = chromedriversPIDcntr + 1
'        End If

        wd.URL = URL
        Set openChromeBackground = wd
    
Exit Function
ErrH:
Select Case Err.Number

    Case -2146233088 '**'
        'Debug.Print Err.Description
        '' err.Descriptionunknown error: Chrome failed to start: exited normally.
        ''  (unknown error: DevToolsActivePort file doesn't exist)
        '' (The process started from chrome location W:\PortableApps\PortableApps\GoogleChromePortable\App\Chrome-bin\chrome.exe is no longer running, so ChromeDriver is assuming that Chrome has crashed.)
        If chromedriversPIDcntr = 0 Then chromedriversPIDcntr = 1
        ReDim chromedriversPID(chromedriversPIDcntr - 1)
        chromedriversPID(chromedriversPIDcntr - 1) = Service.ProcessId
        If InStr(Err.Description, "/session timed out after 60 seconds.") Then
            killchromedriverFromHere
            Set openChromeBackground = Nothing
        Else
            If MsgBox("���������e�}�Ҫ�Chrome�s�����A�~��", vbExclamation + vbOKCancel) = vbOK Then
                'killProcessesByName "ChromeDriver.exe", pid
                killchromedriverFromHere
            GoTo reStart
            End If
        End If
    Case Else
        MsgBox Err.Description, vbCritical
'        Resume
End Select

'20230119 creedit chatGPT�j����

'    Dim driver As New Selenium.WebDriver
'    'driver.start "chrome", "https://www.google.com"
'    driver.SetBinary getChromePathIncludeBackslash
'    driver.start getChromePathIncludeBackslash + "chrome.exe", "https://www.google.com"
'    driver.Get "/"
End Function


'https://www.cnblogs.com/ryueifu-VBA/p/13661128.html
Sub Search(URL As String, frmID As String, keywdID As String, btnID As String, Optional searchStr As String)
    On Error GoTo Err1
    'If searchStr = "" And Selection = "" Then Exit Sub
    'If wd Is Nothing Then
        openChrome (URL)
    'End If
        wd.URL = URL
        Dim form As SeleniumBasic.IWebElement
        Dim keyword As SeleniumBasic.IWebElement
        Dim button As SeleniumBasic.IWebElement
        Set form = wd.FindElementById(frmID)
        Set keyword = form.FindElementById(keywdID)
        Set button = form.FindElementById(btnID)
        word.Application.WindowState = wdWindowStateMinimize
        If searchStr <> "" Then
            keyword.SendKeys searchStr
            '�W�@���J�Y�˯��F�A�G�i�����U�@��;���Y���Q��ܤU�ԲM��A�B�T�w�i��ܵ��G�A�h�٬O�ݭn�U�@��
            button.Click
        End If
    '    Debug.Print WD.title, WD.url
    '    Debug.Print WD.PageSource
    '    MsgBox "�U���h�X�s�����C"
    '    WD.Quit
        Exit Sub
Err1:
        Select Case Err.Number
            Case 49 'DLL �I�s�W����~
                Resume
            Case Else
                MsgBox Err.Description, vbCritical
                SystemSetup.killchromedriverFromHere
    '           Resume
    End Select
End Sub

'��ʫ� �G https://www.cnblogs.com/ryueifu-VBA/p/13661128.html
Sub BaiduSearch(Optional searchStr As String)
    On Error GoTo Err1
    Search "https://www.baidu.com", "form", "kw", "su", searchStr
        Exit Sub
Err1:
        Select Case Err.Number
            Case 49 'DLL �I�s�W����~
                Resume
            Case Else
                MsgBox Err.Description, vbCritical
                SystemSetup.killchromedriverFromHere
    '           Resume
    End Select
End Sub

'�d�߰�y���
Sub dictRevisedSearch(Optional searchStr As String)
    On Error GoTo Err1
    'If searchStr = "" And Selection = "" Then Exit Sub
    Const URL As String = "https://dict.revised.moe.edu.tw/search.jsp?md=1"
    If wd Is Nothing Then
        openChrome (URL)
    End If
        If wd.URL <> URL Then wd.URL = URL
        Dim form As SeleniumBasic.IWebElement
        Dim keyword As SeleniumBasic.IWebElement
        Dim button As SeleniumBasic.IWebElement
        Set form = wd.FindElementById("searchF")
        Set keyword = form.FindElementByName("word")
        Set button = form.FindElementByClassName("submit")
        If searchStr <> "" Then
            keyword.SendKeys searchStr
            If Not button Is Nothing Then
                button.Click
            Else
                keyword.Submit '�o��Ӥ�k���i
    '            Dim k As New SeleniumBasic.keys
    '            keyword.SendKeys k.Enter
            End If
        End If
    '   �h�X�s�����C"
    '    WD.Quit
        Exit Sub
Err1:
        Select Case Err.Number
            Case 49 'DLL �I�s�W����~
                Resume
            Case Else
                MsgBox Err.Description, vbCritical
                SystemSetup.killchromedriverFromHere
    '           Resume
        End Select
End Sub

'�^����y���������}
Function grabDictRevisedUrl_OnlyOneResult(searchStr As String, Optional Background As Boolean) As String
    'If searchStr = "" And Selection = "" Then Exit Sub
    If searchStr = "" Then Exit Function
    If VBA.Left(searchStr, 1) <> "=" Then searchStr = "=" + searchStr '��T�j�M�r����O
    Const notFoundOrMultiKey As String = "&qMd=0&qCol=1" '�d�L��ƩΦp�G����@���ɡA���}��󳣦�������r
    Dim URL As String, retryTime As Byte
    URL = "https://dict.revised.moe.edu.tw/search.jsp?md=1"
    
    On Error GoTo Err1
    
    Dim wdB As SeleniumBasic.IWebDriver, WBQuit As Boolean '=true �h�i�H��Chrome�s����
    
    If Background Then
        WBQuit = True '�]���b�I������A�w�]�n�i�H��
        Set wdB = openChromeBackground(URL)
        If wdB Is Nothing Then
            If wd Is Nothing Then
                openChrome URL
            Else
                WBQuit = False
            End If
            Set wdB = wd
        End If
    Else
        WBQuit = False
            If wd Is Nothing Then
                openChrome URL
            Else
                Set wdB = wd
            End If
            If ActiveXComponentsCanNotBeCreated Then
                Exit Function
            End If
    End If
retry:
        If wdB.URL <> URL Then wd.Navigate.GoToUrl URL ' wdB.url = url
        Dim form As SeleniumBasic.IWebElement
        Dim keyword As SeleniumBasic.IWebElement
        Dim button As SeleniumBasic.IWebElement
        Set form = wdB.FindElementById("searchF")
        Set keyword = form.FindElementByName("word")
        Set button = form.FindElementByClassName("submit")
        
        SetIWebElementValueProperty keyword, searchStr
'        If keyword.text <> searchStr Then 20240914�@�o
'            keyword.Clear
'            keyword.SendKeys searchStr
'        End If
        
        Rem �b headless �ѼƳ]�w�U�}�Ҫ�Chrome�s�����A�O�L�k�ϥΨt�ζK�W�\�઺
        Rem Dim key As New SeleniumBasic.keys
        Rem     keyword.SendKeys key.Control + "v"
        Rem     keyword.SendKeys key.LeftShift + key.Insert
        Rem ���Chrome�s���������\����K�W�\��ո� 20230121 �]����G
        Rem <stale element reference: element is not attached to the page document(Session info: headless chrome=109.0.5414.75)>
        Rem �]���u��ޱ������A���O�s��������
        Rem With keyword
    '        .Click
    '        .SendKeys key.Alt + "e"
    '        .SendKeys "l"
    '        .SendKeys key.Escape
    '        .SendKeys key.Down: .SendKeys key.Down: .SendKeys key.Down
    '        .SendKeys key.Enter
    '    End With
    
        If Not button Is Nothing Then
            button.Click
        Else
            keyword.Submit '�o��Ӥ�k���i
    '            Dim k As New SeleniumBasic.keys
    '            keyword.SendKeys k.Enter
        End If
        URL = wdB.URL
        If InStr(URL, notFoundOrMultiKey) = 0 Then
            grabDictRevisedUrl_OnlyOneResult = URL '�����h�Ǧ^���}
        Else
            grabDictRevisedUrl_OnlyOneResult = "" '�S�����Ǧ^�Ŧr��
        End If
        If WBQuit Then
            '�h�X�s����
            wdB.Quit
            If Not Background Then Set wd = Nothing
        Else
            wdB.Close
        End If
        Exit Function
Err1:
        Select Case Err.Number
            Case 49 'DLL �I�s�W����~
                Resume
            Case 91 '�S���]�w�����ܼƩ� With �϶��ܼ�
                If retryTime > 1 Then
                    MsgBox Err.Number + Err.Description
                Else
    '                SystemSetup.wait 0.5
    '                Resume
    '                Set WD = Nothing
    '                openChrome url
                    Set wdB = wd
    '                WBQuit = True
                    retryTime = retryTime + 1
                    GoTo retry
                End If
            Case -2147467261 '�å��N����Ѧҳ]�w�����󪺰������C
                Set wd = Nothing
                killchromedriverFromHere
                openChrome URL
                Set wdB = wd
                WBQuit = True
                Resume
            Case -2146233088 'unknown error: ChromeDriver only supports characters in the BMP  (Session info: chrome=109.0.5414.75)
                If InStr(Err.Description, "/session timed out after 60 seconds.") Then
                    If wd Is Nothing Then openChrome (URL)
                    Set wdB = wd
                ElseIf InStr(Err.Description, "no such window: target window already closed") Or InStr(Err.Description, "invalid session id") Then
                    wd.Quit: Set wd = Nothing: killchromedriverFromHere: openChrome (URL)
                    Set wdB = wd
                Else
                    'textbox.SendKeys key.LeftShift + key.Insert
                    WBQuit = pasteWhenOutBMP(wdB, URL, "word", searchStr, keyword, Background)
                End If
                Resume Next
            Case Else
                MsgBox Err.Description, vbCritical
                wdB.Quit
                SystemSetup.killchromedriverFromHere
    '           Resume
        End Select

End Function
Rem x �n�d���r,Variants �n���n�ݲ���r ���榨�\�Ǧ^true  20240828.
Function LookupZitools(x As String, Optional Variants As Boolean = False) As Boolean
    On Error GoTo eH
    If Not code.IsChineseCharacter(x) Then
        LookupZitools = False
        Exit Function
    End If
    
    If Not openChrome("https://zi.tools/zi/" + x) Then
        If Not openChrome("https://zi.tools/zi/" + x) Then
            Stop
        End If
    End If
    word.Application.WindowState = wdWindowStateMinimize
    Dim iwe As SeleniumBasic.IWebElement
    Rem �Y�������d�ݲ���r
    If Variants Then
        Dim dt As Date
        dt = VBA.Now
        Do While iwe Is Nothing
            Set iwe = wd.FindElementByCssSelector("#mainContent > span > div.content > div > div.sidebar_navigation > div > div:nth-child(11)")
            If DateDiff("s", dt, VBA.Now) > 3 Then
                Exit Do '�䤣������r������
            End If
        Loop
        If Not iwe Is Nothing Then iwe.Click
    End If
    
    wd.SwitchTo.Window (wd.CurrentWindowHandle)
'    AppActivate "chrome"
    LookupZitools = True
    Exit Function
eH:
Select Case Err.Number
        Case -2146233088
            If InStr(Err.Description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                            '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                            '  (Session info: chrome=128.0.6613.85)
                'Set wd = Nothing
                SystemSetup.killchromedriverFromHere
                Set wd = Nothing
                Resume
            Else
                MsgBox Err.Number & Err.Description, vbExclamation
            End If
        Case Else
            MsgBox "������Chrome�s������A����@���I" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
    End Select
End Function
Rem �d�m����r�r��n�Gx �n�d���r�C�Ǧ^�@�Ӧr��}�C�A��1�Ӥ����O�Ҭd�ߪ��r��A��2�Ӥ����O�d�ߵ��G���}�C�Y�S���A�h�Ǧ^�Ŧr�� ""
Function LookupDictionary_of_ChineseCharacterVariants(x As String) As String()
    On Error GoTo eH
    Dim result(1) As String '1=���ޭȤW���]�̤j�ȡ^
    LookupDictionary_of_ChineseCharacterVariants = result
    If Not code.IsChineseCharacter(x) Then
        Exit Function
    End If
    SystemSetup.SetClipboard x
'    If wd Is Nothing Then
        openChrome "https://dict.variants.moe.edu.tw/"
'    Else
'        openNewTabWhenTabAlreadyExit wd
'        wd.Navigate.GoToUrl "https://dict.variants.moe.edu.tw/"
'    End If

    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '�˯���J��
    Do While iwe Is Nothing
        Set iwe = wd.FindElementByCssSelector("#header > div > flex > div:nth-child(3) > div.quick > form > input[type=text]:nth-child(2)")
        If DateDiff("s", dt, VBA.Now) > 3 Then
            Exit Function
        End If
    Loop
    
    word.Application.WindowState = wdWindowStateMinimize
    wd.SwitchTo.Window (wd.CurrentWindowHandle)
'    VBA.AppActivate "chrome"

    
    If Not iwe Is Nothing Then
        Dim keys As New SeleniumBasic.keys
        iwe.SendKeys keys.Shift + keys.Insert
        iwe.SendKeys keys.Enter
        '�d�ߵ��G�T���ءA�p[ �] ]�A �d�ߵ��G�G���� 1 �r�A�����r 3 �r
        Set iwe = wd.FindElementByCssSelector("body > main > div > flex > div:nth-child(1) > red:nth-child(1)")
        If Not iwe Is Nothing Then
            Dim zhengWen As String
            zhengWen = iwe.text
            Set iwe = wd.FindElementByCssSelector("body > main > div > flex > div:nth-child(1) > red:nth-child(2)")
            If zhengWen <> "0" Or iwe.text <> 0 Then
                result(0) = x
                result(1) = wd.URL
                SystemSetup.SetClipboard result(1)
            End If
        Else
            '�p�G������ܸӦr�����A�D�d�ߵ��G���A�p�G https://dict.variants.moe.edu.tw/dictView.jsp?ID=5565
            '�r�Y����
            Set iwe = wd.FindElementByCssSelector("#header > section > h2 > span > a")
            If iwe Is Nothing = False Then
                result(0) = x
                result(1) = wd.URL
                SystemSetup.SetClipboard result(1)
            End If
        End If
    End If
    
    LookupDictionary_of_ChineseCharacterVariants = result
    Exit Function
eH:
Select Case Err.Number
        Case -2146233088
            If InStr(Err.Description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                            '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                            '  (Session info: chrome=128.0.6613.85)
                'Set wd = Nothing
                SystemSetup.killchromedriverFromHere
                Set wd = Nothing
                Resume
            Else
                MsgBox Err.Number & Err.Description, vbExclamation
            End If
        Case Else
            MsgBox "������Chrome�s������A����@���I" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
    End Select
End Function
Rem �d�m��y���n�Gx �n�d���r���C�Ǧ^�@�Ӧr��}�C�A��1�Ӥ����O�Ҭd�ߪ��r��A��2�Ӥ����O�d�ߵ��G���}�C�Y�S���A�h�Ǧ^�Ŧr�� ""
Function LookupDictRevised(x As String) As String()
    On Error GoTo eH
    Dim result(1) As String '1=���ޭȤW���]�̤j�ȡ^
    LookupDictRevised = result

    If Not code.IsChineseString(x) Then
        MsgBox "�u���˯�����C���ˬd�˯��r��A���s�}�l�C", vbExclamation
        Exit Function
    End If
    SystemSetup.SetClipboard x
    
    openChrome "https://dict.revised.moe.edu.tw/search.jsp?md=1"
    
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '�˯���J��
    Do While iwe Is Nothing
        Set iwe = wd.FindElementByCssSelector("#searchF > div.line > input[type=text]:nth-child(1)")
        If DateDiff("s", dt, VBA.Now) > 3 Then
            Exit Function
        End If
    Loop
    
    wd.SwitchTo.Window (wd.CurrentWindowHandle)
    word.Application.WindowState = wdWindowStateMinimize
'    VBA.AppActivate "chrome"

    '����˯��ؤ���
    If Not iwe Is Nothing Then
        Dim keys As New SeleniumBasic.keys
        'iwe.SendKeys keys.Shift + keys.Insert
        iwe.SendKeys keys.Control + "v"
        iwe.SendKeys keys.Enter
        '�d�ߵ��G�T���ءA�p �d�L���
        Set iwe = wd.FindElementByCssSelector("#searchL > tbody > tr > td")
        '�d�ߦ����G�ɡG
        If iwe Is Nothing Then
            result(0) = x
            result(1) = wd.URL
            SystemSetup.SetClipboard result(1)
        End If
    End If
    LookupDictRevised = result
    Exit Function
eH:
Select Case Err.Number
        Case -2146233088
            If InStr(Err.Description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                            '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                            '  (Session info: chrome=128.0.6613.85)
                'Set wd = Nothing
                SystemSetup.killchromedriverFromHere
                Set wd = Nothing
                Resume
            Else
                MsgBox Err.Number & Err.Description, vbExclamation
            End If
        Case Else
            MsgBox "������Chrome�s������A����@���I" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
    End Select
End Function
Rem �d�m�~�y�j����n�Gx �n�d���r���C�Ǧ^�@�Ӧr��}�C�A��1�Ӥ����O�Ҭd�ߪ��r��A��2�Ӥ����O�d�ߵ��G���}�C�Y�S���A�h�Ǧ^�Ŧr�� ""
Function LookupHYDCD(x As String) As String()
    On Error GoTo eH
    Dim result(1) As String '1=���ޭȤW���]�̤j�ȡ^
    LookupHYDCD = result
    If Not code.IsChineseString(x) Then
        MsgBox "�u���˯�����C���ˬd�˯��r��A���s�}�l�C", vbCritical
        Exit Function
    End If
    SystemSetup.SetClipboard x
    
    If openChrome("https://ivantsoi.myds.me/web/hydcd/search.html") = False Then
        openChrome ("https://ivantsoi.myds.me/web/hydcd/search.html")
        
    End If
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '�˯���J��
    Do While iwe Is Nothing
        Set iwe = wd.FindElementByCssSelector("#SearchBox")
        If DateDiff("s", dt, VBA.Now) > 3 Then
            Exit Function
        End If
    Loop
    
    wd.SwitchTo.Window (wd.CurrentWindowHandle)
    word.Application.WindowState = wdWindowStateMinimize
'    VBA.AppActivate "chrome"

    '����˯��ؤ���
    If Not iwe Is Nothing Then
        Dim keys As New SeleniumBasic.keys
        iwe.SendKeys keys.Shift + keys.Insert
        'iwe.SendKeys keys.Control + "v"
        iwe.SendKeys keys.Enter
        '�d�ߵ��G�T���ءA�p ��p�A�L�����y�C
                        '�����y������L�k�d��²��r�A�]�L�k�w����r�C
                        '�Y�n�d��r�A�i�d�ߥH�Ӧr�}�Y�����y�A�A���u�W�@���v����ӳ�r�X�{�A
                        '�ΨϥΤU�������r�d�ߪ��m�~�y�j����n�s��
                        '�ΨϥΡm�~�y�j�r��n�C
        Set iwe = wd.FindElementByCssSelector("#SearchResult > font")
        '�d�ߦ����G�ɡG
        If iwe Is Nothing Then
            '�d�ߵ��G���W�s����
            Set iwe = wd.FindElementByCssSelector("#SearchResult > p > a > font")
            If Not iwe Is Nothing Then
                iwe.Click
                result(0) = x
                wd.SwitchTo.Window WindowHandlesItem(WindowHandlesCount - 1)
                result(1) = wd.URL
                SystemSetup.SetClipboard result(1)
            Else
                MsgBox "���ˬd", vbCritical
                Stop
            End If
        End If
    End If
    LookupHYDCD = result
    Exit Function
eH:
Select Case Err.Number
        Case -2146233088
            If InStr(Err.Description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                            '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                            '  (Session info: chrome=128.0.6613.85)
                'Set wd = Nothing
                SystemSetup.killchromedriverFromHere
                Set wd = Nothing
                Resume
            Else
                MsgBox Err.Number & Err.Description, vbExclamation
            End If
        Case Else
            MsgBox "������Chrome�s������A����@���I" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
    End Select
End Function
Rem �d�m��Ǥj�v�n�Gx �n�d���r���C�Ǧ^�@�Ӧr��}�C�A��1�Ӥ����O�Ҭd�ߪ��r��A��2�Ӥ����O�d�ߵ��G���}�C�Y�S���A�h�Ǧ^�Ŧr�� ""
Function LookupGXDS(x As String) As String()
    On Error GoTo eH
    Dim result(1) As String '1=���ޭȤW���]�̤j�ȡ^
    LookupGXDS = result
    If Not code.IsChineseString(x) Then
        MsgBox "�u���˯�����C���ˬd�˯��r��A���s�}�l�C", vbCritical
        Exit Function
    End If
    SystemSetup.SetClipboard x
    
    If openChrome("https://www.guoxuedashi.net/zidian/bujian/") = False Then
        If openChrome("https://www.guoxuedashi.net/zidian/bujian/") = False Then
            Stop
        End If
    End If
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '�˯���J��
    Do While iwe Is Nothing
        Set iwe = wd.FindElementByCssSelector("#sokeyzi")
        If DateDiff("s", dt, VBA.Now) > 3 Then
            Exit Function
        End If
    Loop
    
    wd.SwitchTo.Window (wd.CurrentWindowHandle)
    word.Application.WindowState = wdWindowStateMinimize
'    VBA.AppActivate "chrome"

    '����˯��ؤ���
    If Not iwe Is Nothing Then
        Dim keys As New SeleniumBasic.keys
        iwe.SendKeys keys.Shift + keys.Insert
        'iwe.SendKeys keys.Control + "v"
        iwe.SendKeys keys.Enter
        
        '�d�ߵ��G�T���ءA�p �i���̡j�覡�d�K�K�A����ϥΡi�ҽk�j�Ρi�����j�覡�d��C
        Set iwe = wd.FindElementByCssSelector("body > div:nth-child(3) > div.info.l > div.info_content.zj.clearfix > div.info_txt2.clearfix")
        '�d�ߦ����G�ɡG
        If iwe Is Nothing Or VBA.InStr(iwe.text, "�i���̡j�覡�d") = 0 Then
            result(0) = x
            result(1) = wd.URL
            SystemSetup.SetClipboard result(1)
        End If
    End If
    LookupGXDS = result
    Exit Function
eH:
Select Case Err.Number
        Case -2146233088
            If InStr(Err.Description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                            '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                            '  (Session info: chrome=128.0.6613.85)
                'Set wd = Nothing
                SystemSetup.killchromedriverFromHere
                Set wd = Nothing
                Resume
            Else
                MsgBox Err.Number & Err.Description, vbExclamation
            End If
        Case Else
            MsgBox "������Chrome�s������A����@���I" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
    End Select
End Function
Rem �d�m�d���r����W���n�Gx �n�d���r���C�Ǧ^�@�Ӧr��}�C�A��1�Ӥ����O�Ҭd�ߪ��r��A��2�Ӥ����O�d�ߵ��G���}�C�Y�S���A�h�Ǧ^�Ŧr�� ""
Function LookupKangxizidian(x As String) As String()
    On Error GoTo eH
    Dim result(1) As String '1=���ޭȤW���]�̤j�ȡ^
    LookupKangxizidian = result
    If Not code.IsChineseCharacter(x) Then
        Exit Function
    End If
    SystemSetup.SetClipboard x
    
    If Not openChrome("https://www.kangxizidian.com/search/index.php?stype=Word") Then
        If Not openChrome("https://www.kangxizidian.com/search/index.php?stype=Word") Then
            Stop
        End If
    End If
    
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '�˯���J��
    Do While iwe Is Nothing
        Set iwe = wd.FindElementByCssSelector("#cornermenubody1 > font18 > input[type=search]:nth-child(2)")
        If DateDiff("s", dt, VBA.Now) > 3 Then
            Exit Function
        End If
    Loop
    
    word.Application.WindowState = wdWindowStateMinimize
    wd.SwitchTo.Window (wd.CurrentWindowHandle)
'    VBA.AppActivate "chrome"

    '����˯���J��
    If Not iwe Is Nothing Then
        Dim keys As New SeleniumBasic.keys
        iwe.Clear
        iwe.SendKeys keys.Shift + keys.Insert
        iwe.SendKeys keys.Enter
        '�d�ߵ��G�T���ءA�p�G ��p�A�d�L��ơK�K�Э��d�I
                                '�νЬd��H�U��L�r��:
        Set iwe = wd.FindElementByCssSelector("body > center:nth-child(10) > center > table.td0 > tbody > tr > td.td1 > center > font22 > font > p:nth-child(1)")
        If iwe Is Nothing Then
            result(0) = x
            result(1) = wd.URL
            SystemSetup.SetClipboard result(1)
        End If
    End If
    
    LookupKangxizidian = result
    Exit Function
eH:
Select Case Err.Number
        Case -2146233088
            If InStr(Err.Description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                            '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                            '  (Session info: chrome=128.0.6613.85)
                'Set wd = Nothing
                SystemSetup.killchromedriverFromHere
                Set wd = Nothing
                Resume
            Else
                MsgBox Err.Number & Err.Description, vbExclamation
            End If
        Case Else
            MsgBox "������Chrome�s������A����@���I" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
    End Select
End Function
Rem �d�m�ն��`�B�H�a�P����Ѧr�P�Ϲ��d�\�P�ê�k���n�Gx �n�d���r�C�Ǧ^�@�Ӧr��}�C�A��1�Ӥ����O�Ҭd�ߪ��r��A��2�Ӥ����O�d�ߵ��G���}�C�Y�S���Χ��h���A�h�Ǧ^�Ŧr��
Function LookupHomeinmistsShuowenImageAccess_VineyardHall(x As String) As String()
    On Error GoTo eH
    Dim result(1) As String '1=���ޭȤW���]�̤j�ȡ^'�w�]�ȧY�O�Ű}�C
    LookupHomeinmistsShuowenImageAccess_VineyardHall = result '�w�]�ȬO�}�C�ŧi�ɰt�m���A�i�O�|���⦹�t�m��������Ұ}�C�ᤩ�γ]�w�����禡�@���Ǧ^�ȡC�G����U���i�ٲ��A���D�I�s�ݤ��ݭn�Ǧ^�ȨѳB�z
    If Not code.IsChineseCharacter(x) Then
        Exit Function
    End If
    SystemSetup.SetClipboard x
    
    If Not openChrome("https://homeinmists.ilotus.org/shuowen/find.php") Then
        If Not openChrome("https://homeinmists.ilotus.org/shuowen/find.php") Then
            Stop
        End If
    End If
    
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '�˯���J��
    Do While iwe Is Nothing
        Set iwe = wd.FindElementByCssSelector("#queryString1")
        If DateDiff("s", dt, VBA.Now) > 5 Then
            Exit Function
        End If
    Loop
    
'    GoSub iweNothingExitFunction:
    
    word.Application.WindowState = wdWindowStateMinimize
    wd.SwitchTo.Window (wd.CurrentWindowHandle)
'    VBA.AppActivate "chrome"

    '����˯���J��
    Dim keys As New SeleniumBasic.keys
    iwe.Clear
    iwe.SendKeys keys.Shift + keys.Insert
'        iwe.SendKeys keys.Enter'���B��Enter�S�@�ΡA�����˯����s
    '�˯����s
    Set iwe = wd.FindElementByCssSelector("body > div.search-block > table > tbody > tr > td > input[type=button]")
    GoSub iweNothingExitFunction:
    iwe.Click
    
    '�d�ߵ��G�T���ءA�p�G�S�����C�Э��s�˯��C�����²�ƺ~�r�˯��C
    Set iwe = wd.FindElementByCssSelector("#searchedResults")
    GoSub iweNothingExitFunction:
    If VBA.InStr(iwe.text, "�S�����C�Э��s�˯��C�����²�ƺ~�r�˯��C") = 1 Then
        Exit Function
    End If
            
    '�˥X n ��
'    Set iwe = wd.FindElementByCssSelector("#searchedResults > span")
    Dim n As Byte '�p�G�˥X 6 ��
    n = VBA.CByte(VBA.IIf(VBA.IsNumeric(VBA.Trim(VBA.Replace(VBA.Replace(iwe.text, "�˥X", vbNullString), "��", vbNullString))), VBA.Trim(VBA.Replace(VBA.Replace(iwe.text, "�˥X", vbNullString), "��", vbNullString)), "0"))
    If n = 0 Then '�����T�������~�A���ˬd(�]���䤣��ɩ���ܪ��O�G�u�S�����C�Э��s�˯��C�����²�ƺ~�r�˯��C�v
        Exit Function
    End If
    '�˥X���G�u���@���~�۰ʶ}�Ҩ䵲�G�s���A�_�h��ʶ}��
    If n > 1 Then
        result(0) = x
        '�ê�k������2��
        Set iwe = wd.FindElementByCssSelector("#searchTableOut > tr:nth-child(3) > td:nth-child(15)")
        GoSub iweNothingExitFunction
        
        If iwe.text <> vbNullString Then
            LookupHomeinmistsShuowenImageAccess_VineyardHall = result
            Exit Function
        End If
    Else
        result(0) = x
    End If
    
    Set iwe = wd.FindElementByCssSelector("#searchTableOut > tr:nth-child(2) > td:nth-child(15) > a")
    GoSub iweNothingExitFunction
    
    iwe.Click
    wd.SwitchTo.Window WindowHandlesItem(WindowHandlesCount - 1)
    
    result(1) = wd.URL
    SystemSetup.SetClipboard result(1)

    LookupHomeinmistsShuowenImageAccess_VineyardHall = result
    Exit Function
    
iweNothingExitFunction:
    If iwe Is Nothing Then
        LookupHomeinmistsShuowenImageAccess_VineyardHall = result
        Exit Function
    End If
    Return
eH:
Select Case Err.Number
        Case -2146233088
            If InStr(Err.Description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                            '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                            '  (Session info: chrome=128.0.6613.85)
                'Set wd = Nothing
                SystemSetup.killchromedriverFromHere
                Set wd = Nothing
                Resume
            Else
                MsgBox Err.Number & Err.Description, vbExclamation
            End If
        Case Else
            MsgBox "������Chrome�s������A����@���I" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
    End Select
End Function

Rem �d�m�ն��`�B�H�a�P����Ѧr�P�Ϥ��˯�WFG���n�Gx �n�d���r�C�Y�˯������G�A�h�Ǧ^true�A�Y�S���A�Υ��ѡB�G�١A�h�Ǧ^false 20240903
Function LookupHomeinmistsShuowenImageTextSearchWFG_Interpretation(x As String) As Boolean
    On Error GoTo eH
    LookupHomeinmistsShuowenImageTextSearchWFG_Interpretation = False
    If Not code.IsChineseString(x) Then
        Exit Function
    End If
    SystemSetup.SetClipboard x '���˯�����ƻs��ŶKï�H�ƥ�
    
    If Not openChrome("https://homeinmists.ilotus.org/shuowen/WFG2.php") Then
        If Not openChrome("https://homeinmists.ilotus.org/shuowen/WFG2.php") Then
            Stop
        End If
    End If
    
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '�˯��u�ѻ��v���e��������J��
    Do While iwe Is Nothing
        Set iwe = wd.FindElementByCssSelector("#queryString2")
        If DateDiff("s", dt, VBA.Now) > 5 Then
            Exit Function
        End If
    Loop
    
    word.Application.WindowState = wdWindowStateMinimize
    wd.SwitchTo.Window (wd.CurrentWindowHandle)
'    VBA.AppActivate "chrome"

    '����˯���J�ؤ���
    Dim keys As New SeleniumBasic.keys
    iwe.Clear
    iwe.SendKeys keys.Shift + keys.Insert '�K�W�˯����e = x
'        iwe.SendKeys keys.Enter'���B��Enter�S�@�ΡA�����˯����s
    '�˯����s
    Set iwe = wd.FindElementByCssSelector("body > div.search-block > table > tbody > tr > td > input[type=button]:nth-child(4)")
    GoSub iweNothingExitFunction:
    iwe.Click
    
    '�d�ߵ��G�T���ءA�p�G�S�����C�����²�ƺ~�r�˯��C
    Set iwe = wd.FindElementByCssSelector("#searchedResults")
    GoSub iweNothingExitFunction:
    If VBA.InStr(iwe.text, "�S�����C�����²�ƺ~�r�˯��C") = 1 Then
        Exit Function
    End If
    
    LookupHomeinmistsShuowenImageTextSearchWFG_Interpretation = True
    Exit Function
    
iweNothingExitFunction:
    If iwe Is Nothing Then
        Exit Function
    End If
    Return
eH:
Select Case Err.Number
        Case -2146233088
            If InStr(Err.Description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                            '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                            '  (Session info: chrome=128.0.6613.85)
                'Set wd = Nothing
                SystemSetup.killchromedriverFromHere
                Set wd = Nothing
                Resume
            Else
                MsgBox Err.Number & Err.Description, vbExclamation
            End If
        Case Else
            MsgBox "������Chrome�s������A����@���I" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
    End Select
End Function

Rem �d�m�~�y�h�\��r�w�n���^��m����n�u�����v��쪺���e�Gx �n�d���r�C�Ǧ^�@�Ӧr��}�C�A��1�Ӥ����O�m����n�u�����v�����e�r��A��2�Ӥ����O�d�ߵ��G���}�C�Y�S���A�h�Ǧ^�Ŧr��
Function LookupMultiFunctionChineseCharacterDatabase(x As String, Optional backgroundStartChrome As Boolean) As String()
    On Error GoTo eH
    Dim result(1) As String '1=���ޭȤW���]�̤j�ȡ^
    LookupMultiFunctionChineseCharacterDatabase = result
    If Not code.IsChineseCharacter(x) Then
        Exit Function
    End If
    SystemSetup.SetClipboard x
    
    If backgroundStartChrome Then
        Set wd = openChromeBackground("https://humanum.arts.cuhk.edu.hk/Lexis/lexi-mf/")
        If wd Is Nothing Then Exit Function
    Else
        If Not openChrome("https://humanum.arts.cuhk.edu.hk/Lexis/lexi-mf/") Then
            If Not openChrome("https://humanum.arts.cuhk.edu.hk/Lexis/lexi-mf/") Then
                Stop
            End If
        End If
    End If
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '�˯���J��
    Do While iwe Is Nothing
        Set iwe = wd.FindElementByCssSelector("#search_input")
        If DateDiff("s", dt, VBA.Now) > 5 Then
            If backgroundStartChrome Then wd.Quit
            Exit Function
        End If
    Loop
    
    word.Application.WindowState = wdWindowStateMinimize
    wd.SwitchTo.Window (wd.CurrentWindowHandle)
'    VBA.AppActivate "chrome"

    '����˯���J�ؤ���
    Dim keys As New SeleniumBasic.keys
    iwe.Clear
    iwe.SendKeys keys.Shift + keys.Insert '�K�W�˯�����
    iwe.SendKeys keys.Enter
    
    '�����˯����G
    Set iwe = Nothing
    '�������e������
    Set iwe = wd.FindElementByCssSelector("#shuoWenTable > tbody > tr:nth-child(2) > td:nth-child(2)")
    dt = VBA.Now
    Do While iwe Is Nothing
        Set iwe = wd.FindElementByCssSelector("#shuoWenTable > tbody > tr:nth-child(2) > td:nth-child(2)")
        If DateDiff("s", dt, VBA.Now) > 1.5 Then
            If backgroundStartChrome Then wd.Quit
            Exit Function
        End If
    Loop
    
    result(0) = iwe.text
    result(1) = wd.URL
    LookupMultiFunctionChineseCharacterDatabase = result
    If backgroundStartChrome Then wd.Quit
    Exit Function
    
iweNothingExitFunction:
    If iwe Is Nothing Then
        LookupMultiFunctionChineseCharacterDatabase = result
        If backgroundStartChrome Then wd.Quit
        Exit Function
    End If
    Return
eH:
Select Case Err.Number
        Case -2146233088
            If InStr(Err.Description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                            '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                            '  (Session info: chrome=128.0.6613.85)
                'Set wd = Nothing
                SystemSetup.killchromedriverFromHere
                Set wd = Nothing
                Resume
            Else
                MsgBox Err.Number & Err.Description, vbExclamation
            End If
        Case Else
            MsgBox "������Chrome�s������A����@���I" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
    End Select
End Function

Rem �d�m����Ѧr�n���^��m����n�u�����v��쪺���e�Gx �n�d���r,includingDuan �O�_�]�Ǧ^�q�`���e�C�Ǧ^�@�Ӧr��}�C�A��1�Ӥ����O�m����n�]�j�}���^�u�����v�����e�r��A��2�Ӥ����O�d�ߵ��G���}�A��3�ӫh�O�q�`�����e�C�Y�S���A�h�Ǧ^�Ŧr��}�C
Function LookupShuowenOrg(x As String, Optional includingDuan As Boolean) As String()
    On Error GoTo eH
    Dim result(2) As String '2=���ޭȤW���]�̤j�� = UBound �Ǧ^�ȡ^
    LookupShuowenOrg = result '���]�w�n�n�Ǧ^���r��}�C�A��S�ᤩ�ȮɴN�O�Ǧ^�Ŧr�ꪺ�}�C
    If Not code.IsChineseCharacter(x) Then
        Exit Function
    End If
    SystemSetup.SetClipboard x
    
    If Not openChrome("https://www.shuowen.org/") Then
        If Not openChrome("https://www.shuowen.org/") Then
            Stop
        End If
    End If
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '�˯���J��
    Do While iwe Is Nothing
        Set iwe = wd.FindElementByCssSelector("#inputKaishu")
        If DateDiff("s", dt, VBA.Now) > 5 Then
            Exit Function
        End If
    Loop
    
    word.Application.WindowState = wdWindowStateMinimize
    wd.SwitchTo.Window (wd.CurrentWindowHandle)
'    VBA.AppActivate "chrome"

    '����˯���J�ؤ���
    Dim keys As New SeleniumBasic.keys
    iwe.Clear
    iwe.SendKeys keys.Shift + keys.Insert '�K�W�˯�����
    iwe.SendKeys keys.Enter
    
    '�����˯����G
    Set iwe = wd.FindElementByCssSelector("body > div.container.main > div > div.col-md-9.main-content.pull-right > table > tbody > tr > td")
    If Not iwe Is Nothing Then
        If iwe.text = "�S���O��" Then
            Exit Function
        End If
    End If
    '�����檺���e
    Set iwe = wd.FindElementByCssSelector("body > div.container.main > div > div.col-md-9.main-content.pull-right > div.row.summary > div.col-md-9.pull-left.info-container > div.media.info-body > div.media-body")
    GoSub iweNothingExitFunction
    result(0) = iwe.text
    result(1) = wd.URL
    '���o�q�`�����e
    If includingDuan Then
        Dim i As Byte
        i = 1
        'Dim duanCommentary As String
        '���o�q�`�����e�ت�����
        Set iwe = wd.FindElementByCssSelector("body > div.container.main > div > div.col-md-9.main-content.pull-right > div:nth-child(" & i & ") > div")
        Do
            If i > 30 Then Exit Do
            If Not iwe Is Nothing Then
                If VBA.InStr(iwe.GetAttribute("textContent"), "�M�N �q�ɵ��m����Ѧr�`�n") Then Exit Do
            End If
            Set iwe = wd.FindElementByCssSelector("body > div.container.main > div > div.col-md-9.main-content.pull-right > div:nth-child(" & i & ") > div")
            i = i + 1
        Loop
        GoSub iweNothingExitFunction
        result(2) = iwe.GetAttribute("textContent") '=duanCommentary
    End If
    
    LookupShuowenOrg = result
    Exit Function
    
iweNothingExitFunction:
    If iwe Is Nothing Then
        LookupShuowenOrg = result
        Exit Function
    End If
    Return
eH:
Select Case Err.Number
        Case -2146233088
            If InStr(Err.Description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                            '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                            '  (Session info: chrome=128.0.6613.85)
                'Set wd = Nothing
                SystemSetup.killchromedriverFromHere
                Set wd = Nothing
                Resume
            Else
                MsgBox Err.Number & Err.Description, vbExclamation
            End If
        Case Else
            MsgBox "������Chrome�s������A����@���I" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
    End Select
End Function

Rem �d�m����r�r��n���^��u�������Ρv��쪺���e�Gx �n�d���r�C�Ǧ^�@�Ӧr��}�C�A��1�Ӥ����O�u�������Ρv�����e�r��A��2�Ӥ����O�d�ߵ��G���}�C�Y�S���A�h�Ǧ^�Ŧr��}�C 20240916
Function LookupDictionary_of_ChineseCharacterVariants_RetrieveShuoWenData(x As String) As String()
    On Error GoTo eH
    Dim result(1) As String '1=���ޭȤW���]�̤j�ȡ^
    LookupDictionary_of_ChineseCharacterVariants_RetrieveShuoWenData = result
    If Not code.IsChineseCharacter(x) Then
        Exit Function
    End If
    SystemSetup.SetClipboard x
    
    If Not openChrome("https://dict.variants.moe.edu.tw/") Then
        If Not openChrome("https://dict.variants.moe.edu.tw/") Then
            Stop
        End If
    End If
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '�d�߿�J��
    Do While iwe Is Nothing
        Set iwe = wd.FindElementByCssSelector("#header > div > flex > div:nth-child(3) > div.quick > form > input[type=text]:nth-child(2)")
        If DateDiff("s", dt, VBA.Now) > 5 Then
            Exit Function
        End If
    Loop
    
    word.Application.WindowState = wdWindowStateMinimize
    wd.SwitchTo.Window (wd.CurrentWindowHandle)
'    VBA.AppActivate "chrome"

    '���d�߿�J�ؤ���
    Dim keys As New SeleniumBasic.keys
    iwe.Clear
    iwe.SendKeys keys.Shift + keys.Insert '�K�W�˯�����
    iwe.SendKeys keys.Enter
    
    '�d�ߵ��G�T���ءA�p�i[ �] ]�A �d�ߵ��G�G���� 1 �r�A�����r 3 �r �j�����u1�v�o�Ӥ���A�H������ӧP�_
    Set iwe = wd.FindElementByCssSelector("body > main > div > flex > div:nth-child(1) > red:nth-child(1)")
    Rem ��X�Ӫ����G�������G�G�@�O�C�X����B�����r�U�r�C�������A�G�O�����i�H�Ӧr���r�Y������
    If Not iwe Is Nothing Then
        Dim zhengWen As String
        zhengWen = iwe.text '�e�Ҫ��u1�v
        '�e�Ҫ��u3�v
        Set iwe = wd.FindElementByCssSelector("body > main > div > flex > div:nth-child(1) > red:nth-child(2)")
        If zhengWen <> "0" Or iwe.text <> "0" Then
            '�C�X����B�����r�U�r�C������
            Set iwe = wd.FindElementByCssSelector("#searchL > a")
            If Not iwe Is Nothing Then
                If VBA.InStr(iwe.GetAttribute("outerHTML"), " data-tp=") = 0 Then
                GoTo plural
                Else
                    Do Until VBA.InStr(iwe.GetAttribute("outerHTML"), " data-tp=""��"" ")
                    Loop
                End If
            Else
plural: '��d�ߵ��G����@�ӡu�r�v�ɡA�p�u�h�{�v�r
'                Stop
                
                Dim ai As Byte
                ai = 2 '#searchL > a:nth-child(4)'#searchL > a:nth-child(3)'#searchL > a:nth-child(2)
                Set iwe = wd.FindElementByCssSelector("#searchL > a:nth-child(" & ai & ")")
                Do Until VBA.InStr(iwe.GetAttribute("outerHTML"), " data-tp=""��"" ")
                    ai = ai + 1
                    Set iwe = wd.FindElementByCssSelector("#searchL > a:nth-child(" & ai & ")")
                Loop
            End If
            iwe.Click
            '���ˬd �������� �x�s�� ������r�O�_�O�u�������Ρv
            Set iwe = wd.FindElementByCssSelector("#view > tbody > tr:nth-child(2) > th")
            GoSub iweNothingExitFunction
            If iwe.GetAttribute("textContent") <> "��������" Then
                Set iwe = Nothing
                result(0) = "�������ΨS����ơI"
                result(1) = wd.URL
                GoSub iweNothingExitFunction
            End If
            '�������� �x�s�椸��k�䪺�x�s��
            Set iwe = wd.FindElementByCssSelector("#view > tbody > tr:nth-child(2) > td")
            GoSub iweNothingExitFunction
            result(0) = iwe.GetAttribute("textContent")
            result(1) = wd.URL
            SystemSetup.SetClipboard result(1)
        End If
    Else
        '�p�G������ܸӦr�����A�D�d�ߵ��G���A�p�G https://dict.variants.moe.edu.tw/dictView.jsp?ID=5565
        '�r�Y����
        Set iwe = wd.FindElementByCssSelector("#header > section > h2 > span > a")
        If iwe Is Nothing = False Then
        
            '���ˬd �������� �x�s�� ������r�O�_�O�u�������Ρv
            Set iwe = wd.FindElementByCssSelector("#view > tbody > tr:nth-child(2) > th")
            GoSub iweNothingExitFunction
            If iwe.GetAttribute("textContent") <> "��������" Then
                Set iwe = Nothing
                result(0) = "�������ΨS����ơI"
                result(1) = wd.URL
                GoSub iweNothingExitFunction
            End If
            '�������� �x�s�椸��k�䪺�x�s��
            Set iwe = wd.FindElementByCssSelector("#view > tbody > tr:nth-child(2) > td")
            GoSub iweNothingExitFunction
            result(0) = iwe.GetAttribute("textContent")
            result(1) = wd.URL
            SystemSetup.SetClipboard result(1)
        End If
    End If
''''    '�����˯����G
''''    Set iwe = wd.FindElementByCssSelector("body > div.container.main > div > div.col-md-9.main-content.pull-right > table > tbody > tr > td")
''''    If Not iwe Is Nothing Then
''''        If iwe.text = "�S���O��" Then
''''            Exit Function
''''        End If
''''    End If
''''
''''    Set iwe = wd.FindElementByCssSelector("body > div.container.main > div > div.col-md-9.main-content.pull-right > div.row.summary > div.col-md-9.pull-left.info-container > div.media.info-body > div.media-body")
''''    GoSub iweNothingExitFunction
''''
''''    result(0) = iwe.text
''''    result(1) = wd.URL
    LookupDictionary_of_ChineseCharacterVariants_RetrieveShuoWenData = result
    Exit Function
    
iweNothingExitFunction:
    If iwe Is Nothing Then
        LookupDictionary_of_ChineseCharacterVariants_RetrieveShuoWenData = result
        Exit Function
    End If
    Return
eH:
Select Case Err.Number
        Case -2146233088
            If InStr(Err.Description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                            '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                            '  (Session info: chrome=128.0.6613.85)
                'Set wd = Nothing
                SystemSetup.killchromedriverFromHere
                Set wd = Nothing
                Resume
            Else
                MsgBox Err.Number & Err.Description, vbExclamation
            End If
        Case Else
            MsgBox "������Chrome�s������A����@���I" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
    End Select
End Function

Rem �˯�Google

Sub GoogleSearch(Optional searchStr As String)
    On Error GoTo Err1
    If searchStr = "" And Selection = "" Then Exit Sub
   
    SystemSetup.SetClipboard searchStr
    
    'Dim wd As SeleniumBasic.IWebDriver
    'Set wd = openChrome("https://www.baidu.com")
    If Not openChrome("https://www.google.com") Then Exit Sub
    word.Application.WindowState = wdWindowStateMinimize
    Dim iwe As SeleniumBasic.IWebElement
    Dim keys As New SeleniumBasic.keys
    Set iwe = wd.FindElementByCssSelector("#APjFqb")
    If Not iwe Is Nothing Then
        iwe.Clear
        SystemSetup.SetClipboard searchStr
        iwe.SendKeys keys.Shift + keys.Insert
        iwe.SendKeys keys.Enter
    End If
    Exit Sub
    '    Dim form As SeleniumBasic.IWebElement
    '    Dim keyword As SeleniumBasic.IWebElement
    '    Dim button As SeleniumBasic.IWebElement
    '    Set form = wd.FindElementById("form")
    '    Set keyword = form.FindElementById("kw")
    '    Set button = form.FindElementById("su")
    '    keyword.SendKeys VBA.IIf(searchStr = "", Selection, searchStr)
    '    '�W�@���J�Y�˯��F�A�G�i�����U�@��;���Y���Q��ܤU�ԲM��A�B�T�w�i��ܵ��G�A�h�٬O�ݭn�U�@��
    '    button.Click
    ''    Debug.Print WD.title, WD.url
    ''    Debug.Print WD.PageSource
    ''    MsgBox "�U���h�X�s�����C"
    ''    WD.Quit
    '    Exit Sub
Err1:
        Select Case Err.Number
'            Case 49 'DLL �I�s�W����~
'                Resume
            Case -2146233088
                If InStr(Err.Description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                                '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                                '  (Session info: chrome=128.0.6613.85)
                    'Set wd = Nothing
                    SystemSetup.killchromedriverFromHere
                    Set wd = Nothing
                    Resume
                Else
                    MsgBox Err.Number & Err.Description, vbExclamation
                End If
            Case Else
                MsgBox Err.Description, vbCritical
                SystemSetup.killchromedriverFromHere
    '           Resume
        End Select

End Sub

'�K��j�y�Ŧ۰ʼ��I()
Function grabGjCoolPunctResult(text As String, resultText As String, Optional Background As Boolean) As String
    Const URL = "https://gj.cool/punct"
    Dim wdB As SeleniumBasic.IWebDriver, WBQuit As Boolean '=true �h�i�H��Chrome�s����
    Dim textBox As SeleniumBasic.IWebElement, btn As SeleniumBasic.IWebElement, btn2 As SeleniumBasic.IWebElement, item As SeleniumBasic.IWebElement
    Dim timeOut As Byte '�̦h�� timeOut ��
    On Error GoTo Err1
    
    If Background Then
        Rem ����
        Set wdB = openChromeBackground(URL)
        WBQuit = True '�]���b�I������A�w�]�n�i�H��
        If wdB Is Nothing Then
            If wd Is Nothing Then openChrome ("https://gj.cool/punct")
            Set wdB = wd
        End If
    Else
        Rem ���
        If wd Is Nothing Then openChrome ("https://gj.cool/punct")
        Set wdB = wd
    End If
    If wdB Is Nothing Then Exit Function
    If wdB.URL <> URL Then wdB.Navigate.GoToUrl URL
    
    '��z�奻
    Dim chkStr As String: chkStr = VBA.Chr(13) & VBA.Chr(10) & VBA.Chr(7) & VBA.Chr(9) & VBA.Chr(8)
    text = VBA.Trim(text)
    Do While VBA.InStr(chkStr, VBA.Left(text, 1)) > 0
        text = VBA.Mid(text, 2)
    Loop
    Do While VBA.InStr(chkStr, VBA.Right(text, 1)) > 0
        text = VBA.Left(text, Len(text) - 1)
    Loop
    
    
    '�K�W�奻
    Set textBox = wdB.FindElementById("PunctArea")
    Dim key As New SeleniumBasic.keys

'    textBox.Click 20240914�@�o
'    textBox.Clear

    'textbox.SendKeys key.LeftShift + key.Insert
    'textbox.SendKeys VBA.KeyCodeConstants.vbKeyControl & VBA.KeyCodeConstants.vbKeyV
    
    '�p�G�u��vba.Chr(13)�ӨS��vba.Chr(13)&vba.Chr(10)�h�o��|�Ϥ��q�Ÿ������F�]���U�����I���s�@���A���|�Ϥ@�դ��q�Ÿ������A����������աA�~��O�d�@��
    If InStr(text, VBA.Chr(13) & VBA.Chr(10)) = 0 And InStr(text, VBA.Chr(13)) > 0 Then text = Replace(text, VBA.Chr(13), VBA.Chr(13) & VBA.Chr(10) & VBA.Chr(13) & VBA.Chr(10))
    
    SetIWebElement_textContent_Property textBox, text
'    If Background Then 20240914�@�o
'        textBox.SendKeys text 'SystemSetup.GetClipboardText
'
'    Else
'        systemsetup.SetClipboard text
'        textBox.SendKeys key.Control + "v"
'    End If
    
    '�K�W�����h�h�X
    Dim WaitDt As Date, chkTxtTime As Date, nx As String, xl As Integer
    
    nx = textBox.text
    text = nx
    SystemSetup.playSound 1.294
    If nx = "" Then
        grabGjCoolPunctResult = ""
        If WBQuit Then wdB.Quit
        Exit Function
    End If
    
    '���I
    'Set btn = wdB.FindElementByCssSelector("#main > div.my-4 > div.p-1.p-md-3.d-flex.justify-content-end > div.ms-2 > button")
    Set btn = wdB.FindElementByCssSelector("#main > div > div.p-1.p-md-3.d-flex.justify-content-end > div:nth-child(6) > button") '20240710
    '�Y�K�O��vba.Chr(13)&vba.Chr(10)�H�U�o�椴�|�Ϥ��q�Ÿ�����,�G�Y�n�O���q���A�����uvba.Chr(13) & vba.Chr(10) & vba.Chr(13) & vba.Chr(10)�v�G�դ��q�Ÿ��A����u���@��
    If btn Is Nothing Then Stop
    DoEvents
    wdB.SwitchTo().Window (wdB.CurrentWindowHandle)
    SystemSetup.wait 0.9
    'btn.Click
    Dim k As New SeleniumBasic.keys
    btn.SendKeys k.Enter
    SystemSetup.playSound 1.469
    '���ݼ��I����
    'SystemSetup.Wait 3.6
    
    If VBA.Len(text) < 3000 Then
        timeOut = 10
    Else
        timeOut = 20
    End If
    '�̦h�� timeOut ��
    WaitDt = DateAdd("s", timeOut, Now()) '����10��
    xl = VBA.Len(text)
    chkTxtTime = VBA.Now
    Do
        If VBA.DateDiff("s", chkTxtTime, VBA.Now) > 1.5 Then
            nx = textBox.text
            SystemSetup.playSound 1
            '�ˬd�p�G�S������u���I�v���s�A�N�A�����U 20240725 �H�X�{���ݹϥܱ�����P�_
            If wdB.FindElementByCssSelector("#waitingSpinner") Is Nothing Then
                btn.SendKeys k.Enter
            Else
                If wdB.FindElementByCssSelector("#waitingSpinner").Displayed = False And nx = text Then
                    btn.SendKeys k.Enter
                    SystemSetup.playSound 1.469
                End If
            End If
            chkTxtTime = Now
            'VBA.StrComp(text, nx) <> 0
            If nx <> text Then Exit Do
            If InStr(nx, "�A") > 0 And InStr(nx, "�C") > 0 And Len(nx) > xl Then Exit Do
        End If
        If Now > WaitDt Then
            'Exit Do '�W�L���w�ɶ������}
            grabGjCoolPunctResult = ""
            'wdB.Quit
            wdB.Close
            SystemSetup.playSound 1.469
            Exit Function
        End If
    Loop
    'Set btn2 = WDB.FindElementById("dropdownMenuButton2")
    'btn2.Click
    '
    ''�ƻs
    'Set item = WDB.FindElementByCssSelector("#main > div > div.p-1.p-md-3.d-flex.justify-content-end > div.dropdown > ul > li:nth-child(4) > a")
    'item.Click
    '
    ''Ū���ŶKï�@���^�ǭ�
    'SystemSetup.Wait 0.3
    'systemsetup.SetClipboard textbox.text
    'grabGjCoolPunctResult = SystemSetup.GetClipboardText
    grabGjCoolPunctResult = textBox.text
    resultText = grabGjCoolPunctResult
    If WBQuit = False Then
        wdB.Close
    Else
        wdB.Quit
        If Not Background Then Set wd = Nothing
    End If
    'Debug.Print grabGjCoolPunctResult
    Exit Function
    
Err1:
        Select Case Err.Number
            Case 49 'DLL �I�s�W����~
                Resume
            Case 91 '�S���]�w�����ܼƩ� With �϶��ܼ�
                    killchromedriverFromHere
                    openChrome URL
                    Set wdB = wd
                Resume
            Case -2146233088 'unknown error: ChromeDriver only supports characters in the BMP  (Session info: chrome=109.0.5414.75)
                Rem �����L�@��
                Rem systemsetup.SetClipboard text
                Rem SystemSetup.Wait 0.3
                Rem textBox.SendKeys key.Control + "v"
                Rem textBox.SendKeys key.LeftShift + key.Insert
                If InStr(Err.Description, "ChromeDriver only supports characters in the BMP") Then
                    WBQuit = pasteWhenOutBMP(wdB, URL, "PunctArea", text, textBox, Background)
                    Resume Next
                ElseIf InStr(Err.Description, "invalid session id") Or InStr(Err.Description, "A exception with a null response was thrown sending an HTTP request to the remote WebDriver server for URL http://localhost:4609/session/455865a54d3f64364cf76b41fe7953a3/url. The status of the exception was ConnectFailure, and the message was: �L�k�s���ܻ��ݦ��A��") Then 'Or InStr(Err.Description, "no such window: target window already closed") Then
                    killchromedriverFromHere
                    openChrome URL
                    Set wdB = wd: WBQuit = True
                    Resume
                ElseIf InStr(Err.Description, "no such window: target window already closed") Then
                    openNewTabWhenTabAlreadyExit wdB
                    wdB.Navigate.GoToUrl URL
                    Resume
                Else
                    MsgBox Err.Number + Err.Description
                    Stop
                End If
            Case -2147467261 '�å��N����Ѧҳ]�w�����󪺰������C
                If InStr(Err.Description, "�å��N����Ѧҳ]�w�����󪺰������C") Then
                    killchromedriverFromHere 'WD.Quit: Set WD = Nothing:
                     openChrome URL
                    Set wdB = wd
                    Resume
                Else
                    MsgBox Err.Description, vbCritical
                    Stop
                End If
            Case Else
                MsgBox Err.Description, vbCritical
                wdB.Quit
                SystemSetup.killchromedriverFromHere
    '           Resume
        End Select
    
End Function
Rem 20240914 creedit_with_Copilot�j���ġGhttps://sl.bing.net/gCpH6nC61Cu
' �]�w���� IWebElement��value�ݩʭ�  20240913
Private Function SetIWebElementValueProperty(iwe As IWebElement, txt As String) As Boolean
    If Not iwe Is Nothing Then
        'driver.ExecuteScript "arguments[0].value = arguments[1];", element, valueToSet
        wd.ExecuteScript "arguments[0].value = arguments[1];", iwe, txt
        SetIWebElementValueProperty = True
    End If
End Function
Rem 20240914 creedit_with_Copilot�j���ġGhttps://sl.bing.net/gCpH6nC61Cu
' �]�w���� IWebElement��value�ݩʭ�  20240913
Private Function SetIWebElement_textContent_Property(iwe As IWebElement, txt As String) As Boolean
    If Not iwe Is Nothing Then
        'driver.ExecuteScript "arguments[0].value = arguments[1];", element, valueToSet
        wd.ExecuteScript "arguments[0].textContent = arguments[1];", iwe, txt
        SetIWebElement_textContent_Property = True
    End If
End Function


Private Function pasteWhenOutBMP(ByRef iwd As SeleniumBasic.IWebDriver, URL, textBoxToPastedID, pastedTxt As String, ByRef textBox As SeleniumBasic.IWebElement, Background As Boolean) As Boolean ''unknown error: ChromeDriver only supports characters in the BMP  (Session info: chrome=109.0.5414.75)
Rem creedit chatGPT�j���ġG�z���쪺�T��O Selenium �� SendKeys ��k����K�W BMP �~���r�����D�C
On Error GoTo Err1
Dim retryTimes As Byte
DoEvents
'systemsetup.SetClipboard pastedTxt
'SystemSetup.Wait 0.2
If Background Then iwd.Quit
retry:
If iwd Is Nothing Then
    If wd Is Nothing Then
        openChrome (URL)
        pasteWhenOutBMP = True
    End If
    Set iwd = wd
End If
If iwd.URL <> URL Then iwd.Navigate.GoToUrl (URL)
Dim key As New SeleniumBasic.keys
Set textBox = iwd.FindElementById(textBoxToPastedID)
If textBox Is Nothing Then Set textBox = iwd.FindElementByName(textBoxToPastedID)
textBox.Click

'�K�W
'SystemSetup.Wait 1.5
'textbox.SendKeys key.LeftShift + key.Insert
textBox.SendKeys key.Control + "v"

Exit Function
Err1:
    Select Case Err.Number
        Case 49 'DLL �I�s�W����~
            Resume
        Case 91 '���]�w�����ܼ�
            If retryTimes > 1 Then
                MsgBox Err.Number + Err.Description
            Else
                retryTimes = retryTimes + 1
                GoTo retry
            End If
        Case -2146233088 '��ӿ��~�����X�O�@�˪��A�u��δy�z�ӧP�_�F
        'unknown error: ChromeDriver only supports characters in the BMP  (Session info: chrome=109.0.5414.75)
'            textbox.SendKeys key.LeftShift + key.Insert
'            usePaste WD, url
'            Resume Next
            If InStr(Err.Description, "timed out after 60 seconds") Or InStr(Err.Description, "�L�k�s���ܻ��ݦ��A��") Then
                'The HTTP request to the remote WebDriver server for URL http://localhost:1944/session/d83a0c74803e25f1e7f48999b87a6b7d/element/69589515-4189-4db6-8655-80e30fc05ee0/value timed out after 60 seconds.
                'A exception with a null response was thrown sending an HTTP request to the remote WebDriver server for URL http://localhost:1921/session//element/a9208c93-91ae-4956-9455-d42f51719f23/text. The status of the exception was ConnectFailure, and the message was: �L�k�s���ܻ��ݦ��A��
                iwd.Close
                SystemSetup.killchromedriverFromHere
            End If
        Case -2147467261 '�å��N����Ѧҳ]�w�����󪺰������C
'            If Not WD Is Nothing Then WD.Quit
            Set wd = Nothing
            GoTo retry
        Case Else
            If wd Is Nothing Then
                openChrome (URL)
                pasteWhenOutBMP = True
                Resume Next
            Else
                MsgBox Err.Description, vbCritical
    '            WD.Quit
                iwd.Close
                SystemSetup.killchromedriverFromHere
               Resume
            End If
    End Select
End Function

Function WindowHandlesItem(index As Long) As String
    Dim windowHandle, i As Long
    For Each windowHandle In wd.WindowHandles
        If i = index Then
            WindowHandlesItem = windowHandle
            Exit Function
        End If
        i = i + 1
    Next
End Function

Public Property Get WindowHandlesCount() As Long
    WindowHandlesCount = UBound(wd.WindowHandles) + 1
End Property
Public Property Get WindowHandles() As String()
On Error GoTo eH:
If Not wd Is Nothing Then WindowHandles = wd.WindowHandles
Exit Property
eH:
Select Case Err.Number
    Case -2146233088
        If InStr(Err.Description, "invalid session id") Then
            SystemSetup.killchromedriverFromHere
        Else
            GoTo msg
        End If
    Case Else
msg:
        MsgBox Err.Number + Err.Description
End Select
End Property

