Attribute VB_Name = "SeleniumOP"
Option Explicit
Public WD As SeleniumBasic.IWebDriver
'Private Const timeoutsImplicitWait As Long = 0 '�w�]�Ȭ�0 ��������X���A�������@���ĭȨ��٭�]�w
Private Const timeoutsPageLoad As Long = 300 ''�w�]�Ȭ�300    20241020
Public chromedriversPID() As Long '�x�schromedriver�{��ID���}�C
Public chromedriversPIDcntr As Integer 'chromedriversPID���U�Э�
Public ActiveXComponentsCanNotBeCreated As Boolean

'' �ŧi Windows API ��� 20241003 creedit_with_Copilot�j���ġG��iWordVBA+SeleniumBasic �}��Chrome�s�����s��������k�Ghttps://sl.bing.net/iqY5XH1MVci
'Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
'Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
#End If

Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Private last_ValidWindow As String
Private images_arrayIWebElement() As SeleniumBasic.IWebElement
Private links_arrayIWebElement() As SeleniumBasic.IWebElement
Rem 20241005 Copilot�j���ġGWordVBA �P�_ Word �O�_���̫e�ݵ����Ghttps://sl.bing.net/b21Z8KIK3Ua
Function IsWordActive() As Boolean
    Dim hWnd As LongPtr
    Dim title As String * 255
    Dim Length As Long
    
    ' �����e���ʵ������y�`
    hWnd = GetForegroundWindow()
    
    ' ����������D
    Length = GetWindowText(hWnd, title, Len(title))
    title = Left(title, Length)
    
    ' �ˬd���D�O�_�]�t "Microsoft Word"
    If InStr(title, "Microsoft Word") > 0 Then
        IsWordActive = True
    Else
        IsWordActive = False
    End If
End Function
' �N Chrome �s�����]�m���e�ݵ��f
Sub ActivateChrome()
    Dim hWnd As LongPtr
    hWnd = FindWindow("Chrome_WidgetWin_1", vbNullString)
    If hWnd <> 0 Then
        SetForegroundWindow hWnd
    Else
        hWnd = FindWindow(vbNullString, "Google Chrome")
        If hWnd <> 0 Then
            SetForegroundWindow hWnd
        Else
            MsgBox "Could not find Chrome window", vbCritical
        End If
    End If
End Sub

Property Get LastValidWindow() As String
    LastValidWindow = last_ValidWindow
End Property
Rem 20241008 Copilot�j���ġGhttps://sl.bing.net/cmQuvtGT28O
Rem Gemini�j���ĴN����F�Ihttps://sl.bing.net/cmQuvtGT28O
Property Let LastValidWindow(validWindowHandle As String)
    last_ValidWindow = validWindowHandle
End Property
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
Function openNewTabWhenTabAlreadyExit(ByVal WD As SeleniumBasic.IWebDriver) As Boolean
    On Error GoTo eH
    Dim iw As Byte, ew, ii As Byte
reOpenChrome:
    For Each ew In WD.WindowHandles
        iw = iw + 1
    Next ew
    If iw > 0 Then
          WD.ExecuteScript "window.open('about:blank','_blank');"
          If Not IsNewBlankPageTab(WD) Then
'            Stop 'just for test
            OpenNewTab WD
          End If
          For Each ew In WD.WindowHandles
                ii = ii + 1
                If ii = iw + 1 Then Exit For
          Next ew
          WD.SwitchTo().Window (ew)
    End If
    openNewTabWhenTabAlreadyExit = True
    Exit Function
eH:
    Select Case Err.Number
        Case -2146233088
            If InStr(Err.Description, "no such window: target window already closed") Then
                If iw > 0 Then
                    For Each ew In WD.WindowHandles
                        Exit For
                    Next ew
                    WD.SwitchTo.Window (ew)
                    Resume
                Else
                    Stop
                End If
            ElseIf InStr(Err.Description, "ot connected to DevTools") Then
'                disconnected: not connected to DevTools
'                (failed to check if window was closed: disconnected: not connected to DevTools)
'                (Session info: chrome=127.0.6533.120)
                If Not WD Is Nothing Then WD.Quit
                Set WD = Nothing
                killchromedriverFromHere
                MsgBox "�YChrome�s�����w�}�ҡA������Chrome�s������A���T�w", vbExclamation
                OpenChrome "https://www.google.com"
                Resume 'GoTo reOpenChrome:
            ElseIf InStr(Err.Description, "A exception with a null response was thrown sending an HTTP") Then
'                A exception with a null response was thrown sending an HTTP request to the remote WebDriver server for URL http://localhost:1760/session/ed5864479325c154783256563f97e610/window/handles. The status of the exception was ConnectFailure, and the message was: �L�k�s���ܻ��ݦ��A��
                Set WD = Nothing
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
                Set WD = Nothing
                killchromedriverFromHere
                MsgBox "�YChrome�s�����w�}�ҡA������Chrome�s������A����@��", vbExclamation
                openNewTabWhenTabAlreadyExit = False
            Else
                Stop
            End If
        Case 91
            If Err.Description = "�S���]�w�����ܼƩ� With �϶��ܼ�" Then
                Set WD = Nothing
                killchromedriverFromHere
                MsgBox "�YChrome�s�����w�}�ҡA������Chrome�s������A����@��", vbExclamation
                openNewTabWhenTabAlreadyExit = False
            Else
                Stop
            End If
        Case Else
            MsgBox Err.Description, vbCritical
            WD.Quit
            SystemSetup.killchromedriverFromHere
    '           Resume
    End Select
End Function

Rem �ˬd driver �O�_���� 20241002
Function IsDriverInvalid(ByRef driver As IWebDriver) As Boolean
    On Error Resume Next
    Dim url As String
    url = driver.url
    IsDriverInvalid = (url = vbNullString Or (driver Is Nothing))
End Function
Rem �ˬd wd �O�_���� 20241002
Function IsWDInvalid() As Boolean
    On Error Resume Next
    Dim url As String
    url = WD.url
    IsWDInvalid = (url = vbNullString Or (WD Is Nothing))
End Function

Rem �ˬd�O�_���s���ťխ� �}�Ҫ��s���� 20241003
Function IsNewBlankPageTab(ByRef driver As IWebDriver) As Boolean
    'On Error Resume Next
    IsNewBlankPageTab = (driver.url = "about:blank" Or WD.title = vbNullString) _
                Or (WD.title = "�s����" Or WD.url = "chrome://new-tab-page/")
End Function
Rem �Ұ�Chrome�s�����Τw�Ұʫ�}�ҷs�����s���C���ѮɶǦ^false
Function OpenChrome(Optional url As String) As Boolean
reStart:
        'Dim WD As SeleniumBasic.IWebDriver
        On Error GoTo ErrH
        Dim Service As SeleniumBasic.ChromeDriverService
        Dim options As SeleniumBasic.ChromeOptions
        Dim pid As Long
    
    '����chromedriver.exe
    '�ϥ� WMI �M�W���ҭz����k
    '�P�_PID�O�_����pid
    
        If WD Is Nothing Then
        
            If IsChromeRunning Then '20241002
                If Not OpenChrome_NEW_Get Then Exit Function
                If WD Is Nothing Then
                    'Stop ' for test
                    If MsgBox("������Chrome�s������A���u�T�w�v�~��C�_�h�Ы��u�����v�H�����@�~�C", vbExclamation + vbOKCancel) = VBA.vbCancel Then Exit Function
                    GoTo reStart
                Else
                    WD.url = url
                End If
            Else
                Set WD = New SeleniumBasic.IWebDriver
                
                If WD Is Nothing Then Stop 'just for test
                
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
                Set options = New SeleniumBasic.ChromeOptions
                With options
                    .BinaryLocation = chromePath + "chrome.exe"
                    .AddExcludedArgument "enable-automation" '�T�ΡuChrome ���b�Q�۰ʤƳn�鱱��v��ĵ�i����
                    
                    'C#�Goptions.AddArgument("user-data-dir=" + Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\Google\\Chrome\\User Data\\");
                    .AddArgument "user-data-dir=" + VBA.Environ("LOCALAPPDATA") + _
                        "\Google\Chrome\User Data\"
                    .AddArgument "--new-window"
                    '.AddArgument "--start-maximized"
                    '.DebuggerAddress = "127.0.0.1:9999" '���n�O��L�L�ӲV��
                    
                    .AddArgument "--remote-debugging-port=9222" '20241002 Copilot�j���ġGWord VBA ���� Selenium �ާ@�Ghttps://sl.bing.net/SMTsa6sktU
                                        
                End With
                WD.New_ChromeDriver Service:=Service, options:=options
                Docs.Register_Event_Handler '���M��chromedriver�@�ǳ�
                pid = Service.ProcessId 'Chrome�s�����S���}���\�N�|�O0
                If pid <> 0 Then
                    ReDim Preserve chromedriversPID(chromedriversPIDcntr)
                    chromedriversPID(chromedriversPIDcntr) = pid
                    chromedriversPIDcntr = chromedriversPIDcntr + 1
                End If
                OpenNewTab WD '�e��.AddArgument "--new-window" 20241005 ���O window ���O tab !!
'                WD.ExecuteScript "window.open('about:blank','_blank');" 'openNewTabWhenTabAlreadyExit WD
'                WD.SwitchTo.Window WindowHandlesItem(WindowHandlesCount - 1)
                WD.url = url
            End If
        Else
            If IsWDInvalid() Then OpenNewTab WD
'            WD.ExecuteScript "window.open('about:blank','_blank');" 'openNewTabWhenTabAlreadyExit WD
'            WD.SwitchTo.Window WindowHandlesItem(WindowHandlesCount - 1)
            WD.url = url
        End If
        If ActiveXComponentsCanNotBeCreated Then ActiveXComponentsCanNotBeCreated = False
        OpenChrome = True
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
                WD.Quit
                SystemSetup.killchromedriverFromHere
                Set WD = Nothing
                Stop
                If MsgBox("������Chrome�s�����A�~��I" & vbCr & vbCr & _
                    vbTab & "�O�_�n�{���۰����z�����B�ҰʡC�P���P���@�n�L��������", vbCritical + vbOKCancel) _
                        = vbOK Then
                    SystemSetup.killProcessesByName "chrome.exe"
                    GoTo reStart
                Else
                    OpenChrome = False
                End If
                
                Exit Function
            End If
        Case -2146233088 '**'
            Debug.Print Err.Number & Err.Description
            If VBA.InStr(Err.Description, "invalid session id") = 1 Then '-2146233088 invalid session id
                killchromedriverFromHere
                Set WD = Nothing
                GoTo reStart
            ElseIf InStr(Err.Description, "Chrome failed to start: exited normally.") Then
                '' err.Descriptionunknown error: Chrome failed to start: exited normally.
                ''  (unknown error: DevToolsActivePort file doesn't exist)
                '' (The process started from chrome location W:\PortableApps\PortableApps\GoogleChromePortable\App\Chrome-bin\chrome.exe is no longer running, so ChromeDriver is assuming that Chrome has crashed.)
                If MsgBox("���������e�}�Ҫ�Chrome�s�����A�~��", vbExclamation + vbOKCancel) = vbOK Then
                        'killProcessesByName "ChromeDriver.exe", pid
                        killchromedriverFromHere
                        Set WD = Nothing
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
'                Dim urlCheck As String
'                On Error Resume Next
'                urlCheck = Wd.url
'                If urlCheck = vbNullString Then
'                    killchromedriverFromHere
'                    Set Wd = Nothing
'                    GoTo reStart
'                End If
'                On Error GoTo 0
                WD.SwitchTo.Window WindowHandlesItem(WindowHandlesCount - 1)
                Resume
                '�^�� wd.ExecuteScript "window.open('about:blank','_blank');" 'openNewTabWhenTabAlreadyExit WD
                     'wd.SwitchTo.Window WindowHandlesItem(WindowHandlesCount - 1)
            ElseIf InStr(Err.Description, "Unexpected error. System.Net.WebException: �L�k�s���ܻ��ݦ��A��") Then 'Unexpected error. System.Net.WebException: �L�k�s���ܻ��ݦ��A�� ---> System.Net.Sockets.SocketException: �L�k�s�u�A�]���ؼйq���ڵ��s�u�C 127.0.0.1:6579
                                                                                                '   �� System.Net.Sockets.Socket.DoConnect(EndPoint endPointSnapshot, SocketAddress socketAddress)
                                                                                                '   �� System.Net.ServicePoint.ConnectSocketInternal(Boolean connectFailure, Socket s4, Socket s6, Socket& socket, IPAddress& address, ConnectSocketState state, IAsyncResult asyncResult, Exception& exception)
                                                                                                '   --- �����ҥ~���p���|�l�ܪ����� ---
                                                                                                '   �� System.Net.HttpWebRequest.GetRequestStream(TransportContext& context)
                                                                                                '   �� System.Net.HttpWebRequest.GetRequestStream()
                                                                                                '   �� OpenQA.Selenium.Remote.HttpCommandExecutor.MakeHttpRequest(HttpRequestInfo requestInfo)
                                                                                                '   �� OpenQA.Selenium.Remote.HttpCommandExecutor.Execute(Command commandToExecute)
                                                                                                '   �� OpenQA.Selenium.Remote.DriverServiceCommandExecutor.Execute(Command commandToExecute)
                                                                                                '   �� OpenQA.Selenium.Remote.RemoteWebDriver.Execute(String driverCommandToExecute, Dictionary`2 parameters)

                SystemSetup.killchromedriverFromHere
                Set SeleniumOP.WD = Nothing
'                Stop 'just for test 20240924
                Resume
            ElseIf VBA.InStr(Err.Description, "chromedriver.exe does not exist") Then 'The file C:\Program Files\Google\Chrome\Application\chromedriver.exe does not exist. The driver can be downloaded at http://chromedriver.storage.googleapis.com/index.html
                Set WD = Nothing
                MsgBox "�Цb�u" & getChromePathIncludeBackslash & "�v���|�U�ƻschromedriver.exe�ɮצA�~��I", vbCritical
                OpenChrome = False
                SystemSetup.OpenExplorerAtPath getChromePathIncludeBackslash
                Exit Function
            ElseIf VBA.InStr(Err.Description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                                                                    '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                                                                    '  (Session info: chrome=129.0.6668.60)
                killchromedriverFromHere
                Set WD = Nothing
                GoTo reStart
            ElseIf VBA.InStr(Err.Description, "no such window") Then 'no such window
                                                                    '  (Session info: chrome=129.0.6668.59)
                OpenNewTab WD
                Resume
            ElseIf VBA.InStr(Err.Description, "timeout: Timed out receiving message from renderer:") = 1 Then '-2146233088 timeout: Timed out receiving message from renderer: 2.972
                                                                    '(Session info: chrome=130.0.6723.69)
                If Not IsWDInvalid() Then
                    WD.Manage.Timeouts.PageLoad = timeoutsPageLoad
                    Resume
                Else
                    GoTo 2146233088
                End If
            Else
2146233088:
                Debug.Print Err.Number; Err.Description
                MsgBox Err.Description, vbCritical
                Stop
            End If
        Case 429 'ActiveX ����L�k���ͪ���'
            ActiveXComponentsCanNotBeCreated = True
            Exit Function
        Case -2147467261
            If InStr(Err.Description, "�å��N����Ѧҳ]�w�����󪺰������C") Then
                SystemSetup.killchromedriverFromHere
                Set WD = Nothing
'                Stop
                If MsgBox("������Chrome�s�����A�~��I" & vbCr & vbCr & _
                    vbTab & "�O�_�n�{���۰����z�����B�ҰʡC�P���P���@�n�L��������", vbCritical + vbOKCancel) _
                        = vbOK Then
                    SystemSetup.killProcessesByName "chrome.exe"
                    GoTo reStart
                Else
                    killchromedriverFromHere
                    Set WD = Nothing
                    OpenChrome = False
                End If
                Exit Function
            Else
                MsgBox Err.Description, vbCritical
                Stop
            End If
        Case Else
            MsgBox Err.Description, vbCritical
            If Err.Description = "�S���]�w�����ܼƩ� With �϶��ܼ�" Then
                killchromedriverFromHere
                Set WD = Nothing
                GoTo reStart
            End If
            
            Resume
    End Select

End Function

Function openChromeBackground(url As String) As SeleniumBasic.IWebDriver
reStart:
    'Dim WD As SeleniumBasic.IWebDriver
    On Error GoTo ErrH
    Dim WD As SeleniumBasic.IWebDriver
    Dim Service As SeleniumBasic.ChromeDriverService
    Dim options As SeleniumBasic.ChromeOptions
    Dim pid As Long
    
        Set WD = New SeleniumBasic.IWebDriver
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
        
        Set options = New SeleniumBasic.ChromeOptions
        With options
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
'            .AddArgument "--remote-debugging-port=9222"
        End With
        WD.New_ChromeDriver Service:=Service, options:=options
        'WD.Quit �|�۰ʲM��chromedriver�A�N���ΰO�U�}�L���ǤF
'        pid = Service.ProcessId 'Chrome�s�����S���}���\�N�|�O0
'        If pid <> 0 Then
'            ReDim Preserve chromedriversPID(chromedriversPIDcntr)
'            chromedriversPID(chromedriversPIDcntr) = pid
'            chromedriversPIDcntr = chromedriversPIDcntr + 1
'        End If
        
'        OpenNewTab WD
        WD.url = url
        Set openChromeBackground = WD
    
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

Rem 20241002 Copilot�j���ġGWord VBA ���� Selenium �ާ@�G
Rem �b VBA ���s���� ChromeDriver�Ghttps://sl.bing.net/ib6ZEOurJ4S
Rem �ϥ� SeleniumBasic �b VBA ���s����w�Ұʪ� ChromeDriver�C�Ҧp
Sub SeleniumGet()
'    Dim driver As New WebDriver
'    Dim options As New ChromeOptions
'
'    options.AddArgument "--remote-debugging-port=9222"
'    driver.start "chrome", options
'
'    driver.Get "http://localhost:9222"
'    ' �i��i�@�B���ާ@
End Sub
Sub SeleniumGetTest()
    Dim driver As New IWebDriver
    Dim options As New ChromeOptions
    Dim Service As New SeleniumBasic.ChromeDriverService
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

    options.AddArgument "--remote-debugging-port=9222"
    On Error Resume Next
    driver.New_ChromeDriver Service, options

'    driver.Get "http://localhost:9222"
'    driver.Navigate.GoToUrl "http://localhost:9222"
    If Err.Number = 0 Then
        driver.SwitchTo.Window driver.CurrentWindowHandle
        VBA.Interaction.DoEvents
        SendKeys "%{F4}"
        SystemSetup.playSound 1.469
        VBA.Interaction.DoEvents
    End If
    If Not IsDriverInvalid(driver) Then
        Set WD = driver
    Else
        OpenChrome_NEW_Get
    End If
    
    On Error GoTo 0
'    ' �i��i�@�B���ާ@
End Sub
Rem 20241002 �ѫe�� SeleniumGet �o�쪺�F�P creedit_with_Copilot�j���ġGhttps://sl.bing.net/hwtm2YPAfdY
Function OpenChrome_NEW_Get() As Boolean
    On Error GoTo eH
    If Not WD Is Nothing Then
        If Not IsWDInvalid() Then
            Exit Function
        End If
    End If
    'Dim driver As New WebDriver
    Dim driver As New IWebDriver
    Dim options As New ChromeOptions
    Dim Service As SeleniumBasic.ChromeDriverService
    Dim closeNewOpen As Boolean
    Dim pid As Long
    closeNewOpen = IsChromeRunning
    
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
    With options
'        If Not IsChromeRunning Then '�Y�w�ΦP�@�ϥΪ̳]�w�ɶ}�ҫh�L�k�A�}�s��Chrome�s�����F
                                                '   session not created: Chrome failed to start: exited normally.
                                                '  (chrome not reachable)
                                                '  (The process started from chrome location W:\PortableApps\PortableApps\GoogleChromePortable\App\Chrome-bin\chrome.exe is no longer running, so ChromeDriver is assuming that Chrome has crashed.) (SessionNotCreated)

            .AddArgument "user-data-dir=" + VBA.Environ("LOCALAPPDATA") + _
                        "\Google\Chrome\User Data\"
'        End If
        .AddArgument "--remote-debugging-port=9222"
        .AddArgument "--start-maximized"
    End With
    'driver.start "chrome", options
    On Error Resume Next
    driver.New_ChromeDriver Service:=Service, options:=options
    Docs.Register_Event_Handler '���M��chromedriver�@�ǳ�
    pid = Service.ProcessId 'Chrome�s�����S���}���\�N�|�O0
    If pid <> 0 Then
        ReDim Preserve chromedriversPID(chromedriversPIDcntr)
        chromedriversPID(chromedriversPIDcntr) = pid
        chromedriversPIDcntr = chromedriversPIDcntr + 1
    End If
    Set WD = driver
    If IsWDInvalid() Then
        
'        Dim urlCheck As String
'        urlCheck = wd.url
'        If urlCheck = vbNullString Then
'            WD.SwitchTo.Window WD.CurrentWindowHandle
            
            Rem �ȷ|�~����L���}�Ҫ�Chrome�s����
            'ActivateChrome
'            SystemSetup.wait 2
'            VBA.Interaction.DoEvents
            
            Debug.Print "Word is active = " & VBA.CStr(IsWordActive())
                        
            If VBA.InStr(Err.Description, "from disconnected: unable to connect to renderer (SessionNotCreated)") = 0 Then
                If IsWordActive() Then
                    MsgBox "������Chrome�s������A�~��C", vbExclamation
                    'Stop 'just for test
                    
    '                SendKeys "%{F4}", True '�����w�}�ҦӵL�k���\��Chrome�s����
    '                SystemSetup.playSound 1.469
    '                VBA.Interaction.DoEvents
                Else
    '                Stop 'just for test
'                    ActivateChrome
'                    SendKeys "^{F4}", True '�����w�}�Ҫ�Chrome�s��������
'                    SystemSetup.playSound 1.469
'                    VBA.Interaction.DoEvents
                    
                End If
            End If
            Set options = New SeleniumBasic.ChromeOptions
            With options
                .AddArgument "--remote-debugging-port=9222"
                '�n������Q Get ���ʱҰʪ�Chrome�s�����A�h�b��ʱҰ�Chrome�s���������|�u�ؼ�(T)�v��줺���ȫ�n�[�W �u--remote-debugging-port=9222�v�A�p�G "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 '20241002
                .AddArgument "--start-maximized"
            End With
'            If Not WD Is Nothing Then
'                'WD.Close
'                WD.Quit
'                Set WD = Nothing
'                killchromedriverFromHere
'            End If
            driver.New_ChromeDriver Service:=Service, options:=options
            Docs.Register_Event_Handler '���M��chromedriver�@�ǳ�
            pid = Service.ProcessId 'Chrome�s�����S���}���\�N�|�O0
            If pid <> 0 Then
                ReDim Preserve chromedriversPID(chromedriversPIDcntr)
                chromedriversPID(chromedriversPIDcntr) = pid
                chromedriversPIDcntr = chromedriversPIDcntr + 1
            End If

            If IsDriverInvalid(driver) Then
'                Stop  'just for test
                killchromedriverFromHere
                Set WD = Nothing
                Exit Function
'                OpenChrome_NEW_Get
'                MsgBox "�ЦA����@���C�P���P���@�n�L��������", vbInformation
                'End
            End If
            Set WD = driver
'            urlCheck = wd.url
'            If urlCheck = vbNullString Then Stop 'just for test
'        End If
    Else
        If closeNewOpen Then
'            driver.SwitchTo.Window WD.CurrentWindowHandle
'            SystemSetup.wait 1.3
'            VBA.Interaction.DoEvents
'            SendKeys "%{F4}" '�������\�Ұʫ�W�ߪ�����
'            VBA.Interaction.DoEvents
            If UBound(driver.WindowHandles) > 0 Then
                Dim wh
                For Each wh In WD.WindowHandles
                    WD.SwitchTo.Window wh
                    If WD.title = "�s����" Then 'WD.url="chrome://new-tab-page/"
                        WD.Close '�������\�Ұʫ�W�ߪ�����
                        SystemSetup.playSound 1.469
                        VBA.Interaction.DoEvents
                        Exit For
                    End If
                Next wh
                If IsWDInvalid() Then
                    WD.SwitchTo.Window UBound(WD.WindowHandles)
                End If
                openNewTabWhenTabAlreadyExit WD
'                OpenNewTab WD '�A�}�Ҥ@�ӷs�����A�ѫ���{���ާ@�ΡA�קK�v�T��w�}�Ҫ�����
            End If
        End If
    End If
    On Error GoTo 0
    Rem 20241008 Gemini�j���ġG�������~�B�z�G ��{������� On Error GoTo 0 �o��ɡA���e�]�w��������~�B�z���|�Q�����C
    Rem ��_�w�]�欰�G �������~�B�z��A�p�G�{���A���J����~�A�N�|���ӹw�]���欰�A��������������ܿ��~�T���Chttps://g.co/gemini/share/7359ab0a85e3
    
    Docs.Register_Event_Handler '���M��chromedriver�@�ǳ�
    'driver.Navigate.GoToUrl "https://github.com/oscarsun72/TextForCtext/blob/master/WordVBA/SeleniumOP.bas"
    'driver.Get "http://localhost:9222"
'    Dim Wd As New SeleniumBasic.IWebDriver
'    'wd.Get "http://localhost:9222"
'    Wd.Navigate "http://localhost:9222" 'https://github.com/GCuser99/SeleniumVBA/discussions/74
    ' �i��i�@�B���ާ@
    
    If closeNewBlankPageTabs() Then OpenNewTab WD
    OpenChrome_NEW_Get = True
    Exit Function
eH:
    Select Case Err.Number
        Case -2146233088
            If VBA.InStr(Err.Description, "chromedriver.exe does not exist") Then 'The file C:\Program Files\Google\Chrome\Application\chromedriver.exe does not exist. The driver can be downloaded at http://chromedriver.storage.googleapis.com/index.html
                Set WD = Nothing
                MsgBox "�Цb�u" & getChromePathIncludeBackslash & "�v���|�U�ƻschromedriver.exe�ɮצA�~��I", vbCritical
                SystemSetup.OpenExplorerAtPath getChromePathIncludeBackslash
                Exit Function
            Else
                GoTo caseElse
            End If
caseElse:
        Case Else
            Debug.Print Err.Number & vbTab & Err.Description
            MsgBox Err.Number & Err.Description
            'Resume
    End Select
End Function
Sub CloseNewBlankPagesTabs()
    closeNewBlankPageTabs
End Sub
Rem �Y�S���s���ťխ��n�����h�Ǧ^false,�Y�u�Ѥ@�Ӥ����h���������B�Ǧ^false�ѫ���ϥ�
Private Function closeNewBlankPageTabs() As Boolean
    Dim w, result As Boolean
    For Each w In WD.WindowHandles
        WD.SwitchTo.Window w
        If SeleniumOP.IsNewBlankPageTab(WD) Then
            If WindowHandlesCount > 1 Then
                WD.Close
                If Not result Then result = True
            Else
                If result Then
                    result = False
                End If
                Exit Function
            End If
        End If
        
    Next w
    closeNewBlankPageTabs = result
End Function
Rem �}�ҷs���� �Y���ѫh�Ǧ^false'��iWordVBA+SeleniumBasic �}��Chrome�s�����s��������k creedit_with_Copilot�j���ġG https://sl.bing.net/bcfc14PWlFc
Function OpenNewTab(ByVal driver As SeleniumBasic.IWebDriver) As Boolean
    Dim result As Boolean
    result = True
    On Error GoTo eH
    driver.ExecuteScript "window.open('about:blank','_blank');" 'openNewTabWhenTabAlreadyExit WD
    VBA.Interaction.DoEvents
    SwitchToLastWindowHandleWindow driver
    If Not IsNewBlankPageTab(driver) Then
        Dim key As New SeleniumBasic.keys, iwe As SeleniumBasic.IWebElement
'        Set iwe = driver.FindElementByTagName("body")
        Set iwe = driver.FindElementByCssSelector("body")
        If iwe Is Nothing Then Stop 'just for test
        iwe.SendKeys key.Control + "t"
        VBA.Interaction.DoEvents
        SystemSetup.playSound 1
        SwitchToLastWindowHandleWindow driver
        If Not IsNewBlankPageTab(driver) Then '�Y�S���\�}��
            '�إ� Actions ���� �C Copilot�j���ġG�b�o�q��i���{���X���A�ڨϥ� CreateObject ��k�ӫإ� Actions ����A�åB�����I�s Perform ��k�Ӱ���ʧ@�C�o�˥i�H�T�O Actions ���󥿽T�إߨð���C
            Dim actions As New SeleniumBasic.actions
            actions.Create driver
            actions.MoveToElement(iwe).Click().Perform
            actions.SendKeys(key.Control + "t").Perform
            actions.SendKeys key.Control + "t"
            actions.SendKeys "^t"
            VBA.Interaction.DoEvents
            SystemSetup.playSound 1
            SwitchToLastWindowHandleWindow driver
            Dim wh
            For Each wh In driver.WindowHandles
                driver.SwitchTo.Window wh
                If IsNewBlankPageTab(driver) Then Exit For
            Next wh
            If Not IsNewBlankPageTab(driver) Then '�Y�S���\�}��
                driver.SwitchTo().Window driver.CurrentWindowHandle
                VBA.Interaction.DoEvents
                
                ActivateChrome
                VBA.Interaction.DoEvents
                
                VBA.Interaction.SendKeys "^t", True
                VBA.Interaction.DoEvents
                SwitchToLastWindowHandleWindow driver
                VBA.Interaction.DoEvents
                SystemSetup.playSound 1 'for test
                If Not IsNewBlankPageTab(driver) Then '�Y�S���\�}��
                    For Each wh In driver.WindowHandles
                        driver.SwitchTo.Window wh
                        If IsNewBlankPageTab(driver) Then Exit For
                    Next wh
                    If Not IsNewBlankPageTab(driver) Then '�Y�S���\�}��
                    'Stop 'for debug
                    ActivateChrome
                    word.Application.Activate
                        If VBA.vbOK = MsgBox("�Y�n�}�ҷs���������Ф�ʶ}�ҫ�A���U�u�T�w�v���s�A�_�h�Y�b�������~�����C" & vbCr & vbCr _
                            & "�Y���Q�b����������A�аȥ��ۦ�}�ҷs�����ηs�����A�A���U�u�T�w�v���s�C�P���P���@�n�L��������", VBA.vbOKCancel + VBA.vbExclamation) Then
                            SwitchToLastWindowHandleWindow driver
                        Else
                            result = False
                            driver.SwitchTo.Window driver.CurrentWindowHandle
                        End If
                    End If
                End If
            End If
        End If
    Else
        SystemSetup.playSound 0.484 'for test
    End If
    
    OpenNewTab = result
    
    Exit Function
eH:
    Select Case Err.Number
        Case -2146233088
            If VBA.InStr(Err.Description, "no such window: target window already closed") = 1 Then 'no such window: target window already closed
                driver.SwitchTo.Window driver.WindowHandles()(UBound(driver.WindowHandles))
                Resume
            Else
                GoTo caseElse
            End If
caseElse:
        Case Else
            Debug.Print Err.Number & Err.Description
            MsgBox Err.Number & Err.Description
'            Resume
    End Select
'    On Error GoTo eH
'    driver.ExecuteScript "window.open('about:blank','_blank');" 'openNewTabWhenTabAlreadyExit WD
'    If Not IsNewBlankPageTab(driver) Then
'        Dim key As New SeleniumBasic.keys, iwe As SeleniumBasic.IWebElement
''        driver.FindElementByTagName("body").SendKeys "^t" ' Ctrl + t to open a new tab '20241003creedit_with_Copilot�j���ġG�ѨMWordVBA + SeleniumBasic�}�s�������D�Ghttps://sl.bing.net/gehCkm98JRA
'        Set iwe = driver.FindElementByTagName("body")
'        If iwe Is Nothing Then Stop 'just for test
'        iwe.SendKeys key.Control + "t"
'
'        If Not IsNewBlankPageTab(driver) Then '�Y�S���\�}��
'            '�إ� Actions ����
'            Dim actions As New SeleniumBasic.actions
''            actions.MoveToElement(iwe).Click().Perform
''            actions.SendKeys(key.Control + "t").Build '().Perform
'
'            If Not IsNewBlankPageTab(driver) Then '�Y�S���\�}��
'                word.Application.windowState = wdWindowStateMinimize
'                driver.SwitchTo().Window driver.CurrentWindowHandle
'                VBA.Interaction.DoEvents
'                SendKeys "^t", True
'                VBA.Interaction.DoEvents
'                If Not IsNewBlankPageTab(driver) Then '�Y�S���\�}��
'                    Stop 'for debug
'                End If
'            End If
'        End If
'
'    End If
'    SwitchToLastWindowHandleWindow driver
'    Exit Sub
'eH:
'    Select Case Err.Number
'        Case -2146233088
'            If VBA.InStr(Err.Description, "no such window: target window already closed") = 1 Then 'no such window: target window already closed
''                                                from unknown error: web view not found
''                                                  (Session info: chrome=129.0.6668.60)
'                driver.SwitchTo.Window driver.WindowHandles()(UBound(driver.WindowHandles))
'                Resume
'            Else
'                GoTo CaseElse
'            End If
'CaseElse:
'        Case Else
'            Debug.Print Err.Number & Err.Description
'            MsgBox Err.Number & Err.Description
''            Resume
'    End Select
End Function
Sub SwitchToLastWindowHandleWindow(driver As SeleniumBasic.IWebDriver)
    driver.SwitchTo().Window driver.WindowHandles()(UBound(driver.WindowHandles))
End Sub

Rem 20241002 Copilot�j���ġGWord VBA ���� Selenium �ާ@: https://sl.bing.net/jH2j6GzDiQm
Rem �ϥ� Word VBA ���o���{�ոպݤf
Rem �ˬd Chrome �s�����O�_�w�ҰʡG �ϥ� WMI (Windows Management Instrumentation) ���ˬd Chrome �s�����O�_���b�B��C
Rem Ū�����{�ոպݤf�G ���]�A�w�g�N���{�ոպݤf�g�J���A�i�H�q�Ӥ��Ū���ݤf���C
Sub CheckChromeAndGetPort()
    Dim chromeRunning As Boolean
    Dim port As String
    
    chromeRunning = IsChromeRunning()
    
    If chromeRunning Then
        port = GetPortFromFile("C:\Temp\port.txt")
        MsgBox "Chrome is running on port: " & port
    Else
        MsgBox "Chrome is not running."
    End If
End Sub

Function IsChromeRunning() As Boolean
    Dim objWMIService As Object
    Dim colProcesses As Object
    Dim objProcess As Object
    Dim chromeRunning As Boolean
    
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'chrome.exe'")
    
    chromeRunning = (colProcesses.Count > 0)
    
    IsChromeRunning = chromeRunning
End Function

Function GetPortFromFile(filePath As String) As String
    Dim fileNum As Integer
    Dim port As String
    
    fileNum = FreeFile
    Open filePath For Input As fileNum
    Input #fileNum, port
    Close fileNum
    
    GetPortFromFile = port
End Function

'https://www.cnblogs.com/ryueifu-VBA/p/13661128.html
Sub Search(url As String, frmID As String, keywdID As String, btnID As String, Optional searchStr As String)
    On Error GoTo Err1
    'If searchStr = "" And Selection = "" Then Exit Sub
    'If wd Is Nothing Then
        If Not OpenChrome(url) Then Exit Sub
    'End If
'        wd.url = url'�e1��w��
        word.Application.windowState = wdWindowStateMinimize
        WD.SwitchTo.Window (WD.CurrentWindowHandle)
        VBA.Interaction.DoEvents
        Dim form As SeleniumBasic.IWebElement
        Dim keyword As SeleniumBasic.IWebElement
        Dim button As SeleniumBasic.IWebElement
        Set form = WD.FindElementById(frmID)
        Set keyword = form.FindElementById(keywdID)
        Set button = form.FindElementById(btnID)
        If searchStr <> "" Then
            'keyword.SendKeys searchStr
            SetIWebElementValueProperty keyword, searchStr
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
'    wd.SwitchTo.Window (wd.CurrentWindowHandle)'Search �̤w�� 20240930
'    VBA.Interaction.DoEvents
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
    Const url As String = "https://dict.revised.moe.edu.tw/search.jsp?md=1"
    If WD Is Nothing Then
        OpenChrome (url)
    End If
        If WD.url <> url Then WD.url = url
        Dim form As SeleniumBasic.IWebElement
        Dim keyword As SeleniumBasic.IWebElement
        Dim button As SeleniumBasic.IWebElement
        Set form = WD.FindElementById("searchF")
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
    Dim url As String, retryTime As Byte
    url = "https://dict.revised.moe.edu.tw/search.jsp?md=1"
    
    On Error GoTo Err1
    
    Dim wdB As SeleniumBasic.IWebDriver, WBQuit As Boolean '=true �h�i�H��Chrome�s����
    
    If Background Then
        WBQuit = True '�]���b�I������A�w�]�n�i�H��
        Set wdB = openChromeBackground(url)
        If wdB Is Nothing Then
            If WD Is Nothing Then
                OpenChrome url
            Else
                WBQuit = False
            End If
            Set wdB = WD
        End If
    Else
        WBQuit = False
            If WD Is Nothing Then
                OpenChrome url
            Else
                Set wdB = WD
            End If
            If ActiveXComponentsCanNotBeCreated Then
                Exit Function
            End If
    End If
retry:
        If wdB.url <> url Then WD.Navigate.GoToUrl url ' wdB.url = url
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
        url = wdB.url
        If InStr(url, notFoundOrMultiKey) = 0 Then
            grabDictRevisedUrl_OnlyOneResult = url '�����h�Ǧ^���}
        Else
            grabDictRevisedUrl_OnlyOneResult = "" '�S�����Ǧ^�Ŧr��
        End If
        If WBQuit Then
            '�h�X�s����
            wdB.Quit
            If Not Background Then Set WD = Nothing
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
                    Set wdB = WD
    '                WBQuit = True
                    retryTime = retryTime + 1
                    GoTo retry
                End If
            Case -2147467261 '�å��N����Ѧҳ]�w�����󪺰������C
                Set WD = Nothing
                killchromedriverFromHere
                OpenChrome url
                Set wdB = WD
                WBQuit = True
                Resume
            Case -2146233088 'unknown error: ChromeDriver only supports characters in the BMP  (Session info: chrome=109.0.5414.75)
                If InStr(Err.Description, "/session timed out after 60 seconds.") Then
                    If WD Is Nothing Then OpenChrome (url)
                    Set wdB = WD
                ElseIf InStr(Err.Description, "no such window: target window already closed") Or InStr(Err.Description, "invalid session id") Then
                    WD.Quit: Set WD = Nothing: killchromedriverFromHere: OpenChrome (url)
                    Set wdB = WD
                Else
                    'textbox.SendKeys key.LeftShift + key.Insert
                    WBQuit = pasteWhenOutBMP(wdB, url, "word", searchStr, keyword, Background)
                End If
                Resume Next
            Case Else
                MsgBox Err.Description, vbCritical
                wdB.Quit
                SystemSetup.killchromedriverFromHere
    '           Resume
        End Select

End Function
Rem 20241006 �m�ݨ�j�y�P�j�y�����˯��n�A���\�h�Ǧ^true
Function KandiangujiSearchAll(searchTxt As String) As Boolean
    Dim exact As Boolean
    Const url = "https://kandianguji.com/search_all"
    SystemSetup.SetClipboard searchTxt
    If VBA.vbOK = VBA.MsgBox("�O�_�n�i��T�˯��j�H", vbQuestion + vbOKCancel) Then exact = True

    If Not IsWDInvalid() Then
        If WD.url <> url Then WD.url = url
    Else
        If Not OpenChrome(url) Then
            Exit Function
        End If
    End If
    
    WD.SwitchTo().Window (WD.CurrentWindowHandle)
    ActivateChrome
    word.Application.windowState = wdWindowStateMinimize
    Dim iwe As SeleniumBasic.IWebElement ', key As New SeleniumBasic.keys
    Set iwe = WD.FindElementByCssSelector("#keyword")
    If iwe Is Nothing Then Exit Function
    SetIWebElementValueProperty iwe, searchTxt
'    iwe.SendKeys key.Enter'���UEnter��õL�@��
    If exact Then
        Set iwe = WD.FindElementByCssSelector("body > div > div > div.form-inline > button.btn.btn-info.btn-lg.ml-2")
    Else
        Set iwe = WD.FindElementByCssSelector("body > div > div > div.form-inline > button.btn.btn-danger.btn-lg")
    End If
    If iwe Is Nothing Then Exit Function
    iwe.Click
    KandiangujiSearchAll = True
End Function
Rem 20241006 �˯��m�~�y�����Ʈw�n�A���\�h�Ǧ^true
Function HanchiSearch(searchTxt As String) As Boolean
    Dim free As Boolean, inside As Boolean
    SystemSetup.SetClipboard searchTxt
    If Not IsWDInvalid() Then
        If VBA.Left(WD.url, VBA.Len("https://hanchi.ihp.sinica.edu.tw/")) <> "https://hanchi.ihp.sinica.edu.tw/" Then
            If VBA.vbCancel = MsgBox("�O�_�O�i���v�ϥΡj�H", vbQuestion + vbOKCancel) Then free = True
        Else
            inside = True
        End If
    Else
        If VBA.vbCancel = MsgBox("�O�_�O�i���v�ϥΡj�H", vbQuestion + vbOKCancel) Then free = True
    End If
    
    Const url = "https://hanchi.ihp.sinica.edu.tw/ihp/hanji.htm"

    If Not IsWDInvalid() Then
        If VBA.Left(WD.url, VBA.Len("https://hanchi.ihp.sinica.edu.tw/")) <> "https://hanchi.ihp.sinica.edu.tw/" Then WD.url = url
    Else
        If Not OpenChrome(url) Then
            Exit Function
        End If
    End If
    
    WD.SwitchTo().Window (WD.CurrentWindowHandle)
    ActivateChrome
    Dim iwe As SeleniumBasic.IWebElement, key As New SeleniumBasic.keys
    If Not inside Then
        If free Then
            Set iwe = WD.FindElementByCssSelector("body > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td > table > tbody > tr:nth-child(4) > td > a:nth-child(8) > img")
        Else
            Set iwe = WD.FindElementByCssSelector("body > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td > table > tbody > tr:nth-child(4) > td > a:nth-child(9) > img")
        End If
        If iwe Is Nothing Then Exit Function
        iwe.Click
    End If
    'keyword
    Set iwe = WD.FindElementByCssSelector("#frmTitle > table > tbody > tr:nth-child(2) > td > table > tbody > tr:nth-child(1) > td > input[type=text]:nth-child(2)")
    If iwe Is Nothing Then Exit Function
    SetIWebElementValueProperty iwe, searchTxt
    iwe.SendKeys key.Enter
    HanchiSearch = True
End Function
Rem x �n�d���r,Variants �n���n�ݲ���r ���榨�\�Ǧ^true  20240828.
Function LookupZitools(x As String, Optional Variants As Boolean = False) As Boolean
    On Error GoTo eH
    If Not code.IsChineseCharacter(x) Then
        LookupZitools = False
        Exit Function
    End If
    
    If Not OpenChrome("https://zi.tools/zi/" + x) Then
        If Not OpenChrome("https://zi.tools/zi/" + x) Then
            Stop
        End If
    End If
    word.Application.windowState = wdWindowStateMinimize
    WD.SwitchTo.Window (WD.CurrentWindowHandle)
    VBA.Interaction.DoEvents
'    AppActivate "chrome"
    Dim iwe As SeleniumBasic.IWebElement
    Rem �Y�������d�ݲ���r
    If Variants Then
        Dim dt As Date
        dt = VBA.Now
        Do While iwe Is Nothing
            Set iwe = WD.FindElementByCssSelector("#mainContent > span > div.content > div > div.sidebar_navigation > div > div:nth-child(11)")
            If DateDiff("s", dt, VBA.Now) > 3 Then
                Exit Do '�䤣������r������
            End If
        Loop
        If Not iwe Is Nothing Then iwe.Click
    End If
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
                Set WD = Nothing
                Resume
            Else
                MsgBox Err.Number & Err.Description, vbExclamation
            End If
        Case Else
            MsgBox "������Chrome�s������A����@���I" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
    End Select
End Function
Rem �d�m�j���p��P�V���u��Ѭd�ߡn,���\�h�Ǧ^true 20241020
Function LookupBook_Xungu_kaom(x As String) As Boolean
    If Not code.IsChineseString(x) Then
        MsgBox "�u������I", vbCritical
        Exit Function
    End If
    Dim iwe As SeleniumBasic.IWebElement, dt As Date, key As New SeleniumBasic.keys
    If Not OpenChrome("http://www.kaom.net/book_xungu.php") Then Exit Function
    word.Application.windowState = wdWindowStateMinimize
    WD.SwitchTo.Window WD.CurrentWindowHandle
    ActivateChrome
    dt = VBA.Now
    '�˯���
    Set iwe = WD.FindElementByCssSelector("body > table > tbody > tr > td > form > input.form_1")
    Do While iwe Is Nothing
        Set iwe = WD.FindElementByCssSelector("body > table > tbody > tr > td > form > input.form_1")
        If VBA.DateDiff("s", dt, VBA.Now) > 5 Then
            Exit Function
        End If
    Loop
    SystemSetup.wait 3.3
    SetIWebElementValueProperty iwe, x
'    iwe.SendKeys key.Enter
'    '�d�߫��s
    SystemSetup.wait 3.3
    Set iwe = WD.FindElementByCssSelector("body > table > tbody > tr > td > form > input.form_2")
    If iwe Is Nothing Then Exit Function
    iwe.Click
    LookupBook_Xungu_kaom = True
End Function
Rem �d�m�j���p��n�~�y�j����,���\�h�Ǧ^true 20241020
Function LookupHYDCD_kaom(x As String) As Boolean
    If Not code.IsChineseString(x) Then
        MsgBox "�u������I", vbCritical
        Exit Function
    End If
    Dim iwe As SeleniumBasic.IWebElement, dt As Date ', key As New SeleniumBasic.keys
    If Not OpenChrome("http://www.kaom.net/book_hanyudacidian.php") Then Exit Function
    word.Application.windowState = wdWindowStateMinimize
    WD.SwitchTo.Window WD.CurrentWindowHandle
    ActivateChrome
    '�˯���
    dt = VBA.Now
    Set iwe = WD.FindElementByCssSelector("body > table > tbody > tr > td > form > input.form_1")
    Do While iwe Is Nothing
        Set iwe = WD.FindElementByCssSelector("body > table > tbody > tr > td > form > input.form_1")
        If VBA.DateDiff("s", dt, VBA.Now) > 5 Then
            Exit Function
        End If
    Loop
    SystemSetup.wait 3.3
    SetIWebElementValueProperty iwe, x
    'iwe.SendKeys key.Enter
    '�d�߫��s
    VBA.Interaction.DoEvents
    SystemSetup.wait 3.3
    Set iwe = WD.FindElementByCssSelector("body > table > tbody > tr > td > form > input.form_2")
    If iwe Is Nothing Then Exit Function
    iwe.Click
    LookupHYDCD_kaom = True
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
        If OpenChrome("https://dict.variants.moe.edu.tw/") = False Then Exit Function
'    Else
'        openNewTabWhenTabAlreadyExit wd
'        wd.Navigate.GoToUrl "https://dict.variants.moe.edu.tw/"
'    End If

    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '�˯���J��
    Do While iwe Is Nothing
        Set iwe = WD.FindElementByCssSelector("#header > div > flex > div:nth-child(3) > div.quick > form > input[type=text]:nth-child(2)")
        If DateDiff("s", dt, VBA.Now) > 3 Then
            Exit Function
        End If
    Loop
    
    word.Application.windowState = wdWindowStateMinimize
    WD.SwitchTo.Window (WD.CurrentWindowHandle)
    VBA.Interaction.DoEvents
'    VBA.AppActivate "chrome"
    ActivateChrome

    
    If Not iwe Is Nothing Then
        Dim keys As New SeleniumBasic.keys
        iwe.SendKeys keys.Shift + keys.Insert
        iwe.SendKeys keys.Enter
        '�d�ߵ��G�T���ءA�p[ �] ]�A �d�ߵ��G�G���� 1 �r�A�����r 3 �r
        Set iwe = WD.FindElementByCssSelector("body > main > div > flex > div:nth-child(1) > red:nth-child(1)")
        If Not iwe Is Nothing Then
            Dim zhengWen As String
            zhengWen = iwe.text
            Set iwe = WD.FindElementByCssSelector("body > main > div > flex > div:nth-child(1) > red:nth-child(2)")
            If zhengWen <> "0" Or iwe.text <> 0 Then
                result(0) = x
                result(1) = WD.url
                SystemSetup.SetClipboard result(1)
            End If
        Else
            '�p�G������ܸӦr�����A�D�d�ߵ��G���A�p�G https://dict.variants.moe.edu.tw/dictView.jsp?ID=5565
            '�r�Y����
            Set iwe = WD.FindElementByCssSelector("#header > section > h2 > span > a")
            If iwe Is Nothing = False Then
                result(0) = x
                result(1) = WD.url
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
                Set WD = Nothing
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
    
    If Not OpenChrome("https://dict.revised.moe.edu.tw/search.jsp?md=1") Then Exit Function
    
    
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '�˯���J��
    Do While iwe Is Nothing
        Set iwe = WD.FindElementByCssSelector("#searchF > div.line > input[type=text]:nth-child(1)")
        If DateDiff("s", dt, VBA.Now) > 3 Then
            Exit Function
        End If
    Loop
    
    word.Application.windowState = wdWindowStateMinimize
    WD.SwitchTo.Window (WD.CurrentWindowHandle)
    VBA.Interaction.DoEvents
'    VBA.AppActivate "chrome"
    ActivateChrome

    '����˯��ؤ���
    If Not iwe Is Nothing Then
        Dim keys As New SeleniumBasic.keys
        'iwe.SendKeys keys.Shift + keys.Insert
        iwe.SendKeys keys.Control + "v"
        iwe.SendKeys keys.Enter
        '�d�ߵ��G�T���ءA�p �d�L���
        Set iwe = WD.FindElementByCssSelector("#searchL > tbody > tr > td")
        '�d�ߦ����G�ɡG
        If iwe Is Nothing Then
            result(0) = x
            result(1) = WD.url
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
                Set WD = Nothing
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
    
    If OpenChrome("https://ivantsoi.myds.me/web/hydcd/search.html") = False Then
        OpenChrome ("https://ivantsoi.myds.me/web/hydcd/search.html")
        
    End If
    word.Application.windowState = wdWindowStateMinimize
    WD.SwitchTo.Window (WD.CurrentWindowHandle)
'    VBA.Interaction.DoEvents
'    VBA.AppActivate "chrome"
    'AppActivateChrome
    SeleniumOP.ActivateChrome
    
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '�˯���J��
    Do While iwe Is Nothing
        Set iwe = WD.FindElementByCssSelector("#SearchBox")
        If DateDiff("s", dt, VBA.Now) > 3 Then
            Exit Function
        End If
    Loop
    

    
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
        Set iwe = WD.FindElementByCssSelector("#SearchResult > font")
        '�d�ߦ����G�ɡG
        If iwe Is Nothing Then
            '�d�ߵ��G���W�s����
            Set iwe = WD.FindElementByCssSelector("#SearchResult > p > a > font")
            If Not iwe Is Nothing Then
                iwe.Click
                result(0) = x
                WD.SwitchTo.Window WindowHandlesItem(WindowHandlesCount - 1)
                result(1) = WD.url
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
                Set WD = Nothing
                Resume
            Else
                MsgBox Err.Number & Err.Description, vbExclamation
            End If
        Case Else
            MsgBox "������Chrome�s������A����@���I" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
    End Select
End Function
Rem �d����j���]�m��Ǥj�v�n�N�کұ����ন�¥ժ��^�A���\�Ǧ^true
Function LookupZWDCD(x As String) As Boolean
    On Error GoTo eH
    If Not code.IsChineseString(x) Then
        MsgBox "�u������I", vbCritical
        Exit Function
    End If
    Dim key As New SeleniumBasic.keys
    Dim iwe As SeleniumBasic.IWebElement, dt As Date, tds() As SeleniumBasic.IWebElement, i As Integer, actions As New SeleniumBasic.actions, flag As Boolean, attr As String, retyrCntr As Byte
    
    If Not IsWDInvalid() Then WD.Manage.Timeouts.PageLoad = 3
    
    If Not OpenChrome("https://www.guoxuedashi.net/zidian/bujian/") Then
        If Not IsWDInvalid() Then WD.Manage.Timeouts.PageLoad = timeoutsPageLoad
        Exit Function
    End If
    'WD.Manage.timeouts.ImplicitWait = 2
    WD.Manage.Timeouts.PageLoad = 2 '�]�m�������J�W��3�� creedit_with_Copilot�j����
    WD.SwitchTo.Window WD.CurrentWindowHandle
    ActivateChrome
    word.Application.windowState = wdWindowStateMinimize
    '�˯���
    dt = VBA.Now
    Set iwe = WD.FindElementByCssSelector("#sokeyzi")
    Do While iwe Is Nothing
        Set iwe = WD.FindElementByCssSelector("#sokeyzi")
        If VBA.DateDiff("s", dt, VBA.Now) > 5 Then
            WD.Manage.Timeouts.PageLoad = timeoutsPageLoad
            Exit Function
        End If
    Loop
    SetIWebElementValueProperty iwe, x
    WD.Manage.Timeouts.PageLoad = 10 '�]�m�������J�W�ɬ�� creedit_with_Copilot�j����
    On Error Resume Next
    iwe.SendKeys key.Enter
    
    On Error GoTo 0
'    '�d�߫��s
'    Set iwe = WD.FindElementByCssSelector("")
'    If iwe Is Nothing Then Exit Function
'    iwe.Click
    actions.Create WD
    dt = VBA.Now
    '�ѥئC��
    Set iwe = WD.FindElementById("shupage")
    Do While iwe Is Nothing And VBA.DateDiff("s", dt, VBA.Now) < 10
        Set iwe = WD.FindElementById("shupage")
        actions.SendKeys(key.End).Perform
    Loop
    actions.SendKeys(key.End).Perform
    '�ϥ� JavaScript �P�_����`��ƬO�_����
    Dim prevRowCount As Long, currRowCount As Long
scroll:
    prevRowCount = 0
    dt = VBA.Now
    Do
        currRowCount = WD.ExecuteScript("return document.getElementById('shupage').rows.length") '20241020creedit_with_Copilot�j����
        If currRowCount > prevRowCount Then
            prevRowCount = currRowCount
            dt = VBA.Now '���m�ɶ�
        End If
        SystemSetup.wait 0.4
        actions.SendKeys(key.End).Perform
        '�p���y
        SystemSetup.wait 0.4 ' 1000 �@���� 1 ��'Application.wait (Now + TimeValue("0:00:01"))
        actions.SendKeys(key.End).Perform
        If VBA.DateDiff("s", dt, VBA.Now) > 20 Then '�ɶ��i�H�վ�
            WD.Manage.Timeouts.PageLoad = timeoutsPageLoad
            Exit Function
        Else
            actions.SendKeys(key.End).Perform
            SystemSetup.wait 1
        End If
    Loop While currRowCount > prevRowCount

    '���ؼм���
    Set iwe = WD.FindElementById("shupage")
    tds = iwe.FindElementsByTagName("td")
    

    dt = VBA.Now
    Do While UBound(tds) = 0
        tds = WD.FindElementsByTagName("td")
        If VBA.DateDiff("s", dt, VBA.Now) > 3 Then
            WD.Manage.Timeouts.PageLoad = timeoutsPageLoad
            Exit Function
        End If
    Loop
        
    On Error GoTo eH:
    
    For i = 0 To UBound(tds)
        If tds(i).GetAttribute("textContent") = "����j" & VBA.ChrW(-28770) & "��" Then
            attr = tds(i + 1).GetAttribute("innerHTML")
            flag = True
            Exit For
        End If
    Next i
    If Not flag Then
        If retyrCntr > 1 Then
            MsgBox "���r�S���m����j���n�T���C�P���P���@�n�L��������@�g���D", vbExclamation
        Else
            retyrCntr = retyrCntr + 1
            GoTo scroll
        End If
    End If
    'WD.Manage.timeouts.ImplicitWait = 3 ' ����3��
    On Error GoTo 0
    On Error GoTo eH:
    WD.Manage.Timeouts.PageLoad = 4 '�]�m�������J�W��x�� creedit_with_Copilot�j����
    WD.url = "https://www.guoxuedashi.net" & HTML2Doc.GetHTMLAttributeValue("href", VBA.Replace(attr, "amp;", vbNullString))
    'Set iwe = WD.FindElementByCssSelector("body > div:nth-child(3) > center:nth-child(2) > img")
    'iwe.Click
    WD.SwitchTo.Window WD.CurrentWindowHandle
    ActivateChrome
    retyrCntr = 0
    Do Until isImageLoaded("body > div:nth-child(3) > center:nth-child(2) > img")
        WD.Manage.Timeouts.PageLoad = WD.Manage.Timeouts.PageLoad + 2
        WD.Navigate.Refresh
        retyrCntr = retyrCntr + 1
        Debug.Print "reload image" & retyrCntr
        If retyrCntr > 3 Then Exit Do
    Loop

    word.Application.Activate
finish:
    LookupZWDCD = True
    playSound 0.411
    WD.SwitchTo.Window WD.CurrentWindowHandle
    ActivateChrome
    WD.FindElementByCssSelector("body > div:nth-child(3) > center:nth-child(2) > img").Click
'    WD.Manage.timeouts.ImplicitWait = timeoutsImplicitWait '�w�]�Ȭ�0
    WD.Manage.Timeouts.PageLoad = timeoutsPageLoad '�w�]�Ȭ�300
    Exit Function
eH:
    Select Case Err.Number
        Case -2146233088
            If VBA.InStr(Err.Description, "stale element reference: stale element not found in the current frame") = 1 Then 'stale element reference: stale element not found in the current frame
'                                                (Session info: chrome=129.0.6668.101)
                actions.SendKeys(key.End).Perform
                GoTo scroll
            ElseIf VBA.InStr(Err.Description, "timeout: Timed out receiving message from renderer:") = 1 Then 'timeout: Timed out receiving message from renderer: 3.000
                                        '  (Session info: chrome=129.0.6668.101)
                If VBA.InStr(WD.url, "zwdcd") Then
    '                WD.Manage.timeouts.ImplicitWait = WD.Manage.timeouts.ImplicitWait + 3 ' ����3��
    '                WD.Manage.timeouts.PageLoad = WD.Manage.timeouts.PageLoad + 3
    '                Resume
                    If Not isImageLoaded("body > div:nth-child(3) > center:nth-child(2) > img") Then
                        WD.Manage.Timeouts.PageLoad = WD.Manage.Timeouts.PageLoad + 5
                        playSound 1
                        On Error Resume Next
                        WD.Navigate.Refresh
                        On Error GoTo 0
                        Debug.Print "reload image" & retyrCntr
                        If isImageLoaded("body > div:nth-child(3) > center:nth-child(2) > img") Then
                            LookupZWDCD = True
                            WD.Manage.Timeouts.PageLoad = timeoutsPageLoad
                            Exit Function
                        Else
                            Resume Next
                        End If
                    Else
                        LookupZWDCD = True
                        Debug.Print "okok..."
                        playSound 0.484
                        WD.Manage.Timeouts.PageLoad = timeoutsPageLoad
                        Exit Function
                    End If
                Else
                    WD.Manage.Timeouts.PageLoad = WD.Manage.Timeouts.PageLoad + 2
                    Resume
                End If
            ElseIf VBA.InStr(Err.Description, "javascript error: Cannot read properties of null (reading 'rows')") Then '-2146233088javascript error: Cannot read properties of null (reading 'rows')
                                                                                                                    '(Session info: chrome=129.0.6668.101)
                word.Application.Activate
                MsgBox "�����G�١A�Ш����@�~�έ��աC�P���P���@�n�L��������@�g���D", vbCritical
                If WD.Manage.Timeouts.PageLoad <> timeoutsPageLoad Then WD.Manage.Timeouts.PageLoad = timeoutsPageLoad
                Exit Function
            Else
                GoTo caseElse
            End If
        Case Else
caseElse:
            Debug.Print Err.Number & Err.Description
            word.Application.Activate
            MsgBox Err.Number & Err.Description, vbCritical
    End Select
End Function
'�j���������J 20241020 creedit_with_Copilot�j����
Sub StopLoadPage()
    WD.ExecuteScript "window.stop();"
End Sub
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
    
    If OpenChrome("https://www.guoxuedashi.net/zidian/bujian/") = False Then
        If OpenChrome("https://www.guoxuedashi.net/zidian/bujian/") = False Then
            Stop
        End If
    End If
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '�˯���J��
    Do While iwe Is Nothing
        Set iwe = WD.FindElementByCssSelector("#sokeyzi")
        If DateDiff("s", dt, VBA.Now) > 3 Then
            Exit Function
        End If
    Loop
    
    word.Application.windowState = wdWindowStateMinimize
    WD.SwitchTo.Window (WD.CurrentWindowHandle)
    VBA.Interaction.DoEvents
'    VBA.AppActivate "chrome"

    '����˯��ؤ���
    If Not iwe Is Nothing Then
        Dim keys As New SeleniumBasic.keys
        iwe.SendKeys keys.Shift + keys.Insert
        'iwe.SendKeys keys.Control + "v"
        iwe.SendKeys keys.Enter
        
        '�d�ߵ��G�T���ءA�p �i���̡j�覡�d�K�K�A����ϥΡi�ҽk�j�Ρi�����j�覡�d��C
        Set iwe = WD.FindElementByCssSelector("body > div:nth-child(3) > div.info.l > div.info_content.zj.clearfix > div.info_txt2.clearfix")
        '�d�ߦ����G�ɡG
        If iwe Is Nothing Or VBA.InStr(iwe.text, "�i���̡j�覡�d") = 0 Then
            result(0) = x
            result(1) = WD.url
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
                Set WD = Nothing
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
    
    If Not OpenChrome("https://www.kangxizidian.com/search/index.php?stype=Word") Then
        If Not OpenChrome("https://www.kangxizidian.com/search/index.php?stype=Word") Then
            Stop
        End If
    End If
    
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '�˯���J��
    Do While iwe Is Nothing
        Set iwe = WD.FindElementByCssSelector("#cornermenubody1 > font18 > input[type=search]:nth-child(2)")
        If DateDiff("s", dt, VBA.Now) > 3 Then
            Exit Function
        End If
    Loop
    
    word.Application.windowState = wdWindowStateMinimize
    WD.SwitchTo.Window (WD.CurrentWindowHandle)
    VBA.Interaction.DoEvents
'    VBA.AppActivate "chrome"

    '����˯���J��
    If Not iwe Is Nothing Then
        Dim keys As New SeleniumBasic.keys
        iwe.Clear
        iwe.SendKeys keys.Shift + keys.Insert
        iwe.SendKeys keys.Enter
        '�d�ߵ��G�T���ءA�p�G ��p�A�d�L��ơK�K�Э��d�I
                                '�νЬd��H�U��L�r��:
        Set iwe = WD.FindElementByCssSelector("body > center:nth-child(10) > center > table.td0 > tbody > tr > td.td1 > center > font22 > font > p:nth-child(1)")
        If iwe Is Nothing Then
            result(0) = x
            result(1) = WD.url
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
                Set WD = Nothing
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
    
    If Not OpenChrome("https://homeinmists.ilotus.org/shuowen/find.php") Then
        If Not OpenChrome("https://homeinmists.ilotus.org/shuowen/find.php") Then
            Stop
        End If
    End If
    
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '�˯���J��
    Do While iwe Is Nothing
        Set iwe = WD.FindElementByCssSelector("#queryString1")
        If DateDiff("s", dt, VBA.Now) > 5 Then
            Exit Function
        End If
    Loop
    
'    GoSub iweNothingExitFunction:
    
    word.Application.windowState = wdWindowStateMinimize
    WD.SwitchTo.Window (WD.CurrentWindowHandle)
    VBA.Interaction.DoEvents
'    VBA.AppActivate "chrome"

    '����˯���J��
    Dim keys As New SeleniumBasic.keys
    iwe.Clear
    iwe.SendKeys keys.Shift + keys.Insert
'        iwe.SendKeys keys.Enter'���B��Enter�S�@�ΡA�����˯����s
    '�˯����s
    Set iwe = WD.FindElementByCssSelector("body > div.search-block > table > tbody > tr > td > input[type=button]")
    GoSub iweNothingExitFunction:
    iwe.Click
    
    '�d�ߵ��G�T���ءA�p�G�S�����C�Э��s�˯��C�����²�ƺ~�r�˯��C
    Set iwe = WD.FindElementByCssSelector("#searchedResults")
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
        Set iwe = WD.FindElementByCssSelector("#searchTableOut > tr:nth-child(3) > td:nth-child(15)")
        GoSub iweNothingExitFunction
        
        If iwe.text <> vbNullString Then
            LookupHomeinmistsShuowenImageAccess_VineyardHall = result
            Exit Function
        End If
    Else
        result(0) = x
    End If
    
    Set iwe = WD.FindElementByCssSelector("#searchTableOut > tr:nth-child(2) > td:nth-child(15) > a")
    GoSub iweNothingExitFunction
    
    iwe.Click
    WD.SwitchTo.Window WindowHandlesItem(WindowHandlesCount - 1)
    
    result(1) = WD.url
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
                Set WD = Nothing
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
    
    If Not OpenChrome("https://homeinmists.ilotus.org/shuowen/WFG2.php") Then
        If Not OpenChrome("https://homeinmists.ilotus.org/shuowen/WFG2.php") Then
            Stop
        End If
    End If
    
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '�˯��u�ѻ��v���e��������J��
    Do While iwe Is Nothing
        Set iwe = WD.FindElementByCssSelector("#queryString2")
        If DateDiff("s", dt, VBA.Now) > 5 Then
            Exit Function
        End If
    Loop
    
    word.Application.windowState = wdWindowStateMinimize
    WD.SwitchTo.Window (WD.CurrentWindowHandle)
    VBA.Interaction.DoEvents
'    VBA.AppActivate "chrome"

    '����˯���J�ؤ���
    Dim keys As New SeleniumBasic.keys
    iwe.Clear
    SetIWebElementValueProperty iwe, x
    'iwe.SendKeys keys.Shift + keys.Insert '�K�W�˯����e = x
'        iwe.SendKeys keys.Enter'���B��Enter�S�@�ΡA�����˯����s
    '�˯����s
    Set iwe = WD.FindElementByCssSelector("body > div.search-block > div > div:nth-child(2) > input[type=button]") '"body > div.search-block > table > tbody > tr > td > input[type=button]:nth-child(4)")
    GoSub iweNothingExitFunction:
    iwe.Click
    
    '�d�ߵ��G�T���ءA�p�G�S�����C�����²�ƺ~�r�˯��C
    Set iwe = WD.FindElementByCssSelector("#searchedResults")
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
                Set WD = Nothing
                Resume
            Else
                MsgBox Err.Number & Err.Description, vbExclamation
            End If
        Case Else
            MsgBox "������Chrome�s������A����@���I" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
    End Select
End Function
Rem �d�m�ն��`�B�H�a�n���m�~�y�j����n 20241020
Function LookupHomeinmistsHYDCD(x As String) As Boolean
    If Not code.IsChineseString(x) Then
        MsgBox "�u������I", vbCritical
        Exit Function
    End If
    Dim iwe As SeleniumBasic.IWebElement, dt As Date, key As New SeleniumBasic.keys
    If Not OpenChrome("https://homeinmists.ilotus.org/hd/hydcd.php") Then Exit Function
    word.Application.windowState = wdWindowStateMinimize
    WD.SwitchTo.Window WD.CurrentWindowHandle
    ActivateChrome
    '�˯���
    dt = VBA.Now
    Set iwe = WD.FindElementByCssSelector("#keywords")
    Do While iwe Is Nothing
        Set iwe = WD.FindElementByCssSelector("#keywords")
        If VBA.DateDiff("s", dt, VBA.Now) > 5 Then
            Exit Function
        End If
    Loop
    SetIWebElementValueProperty iwe, x
    iwe.SendKeys key.Enter
'    '�d�߫��s
'    Set iwe = WD.FindElementByCssSelector("")
'    If iwe Is Nothing Then Exit Function
'    iwe.Click
    LookupHomeinmistsHYDCD = True
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
        Set WD = openChromeBackground("https://humanum.arts.cuhk.edu.hk/Lexis/lexi-mf/")
        If WD Is Nothing Then Exit Function
    Else
        If Not OpenChrome("https://humanum.arts.cuhk.edu.hk/Lexis/lexi-mf/") Then
            If Not OpenChrome("https://humanum.arts.cuhk.edu.hk/Lexis/lexi-mf/") Then
                Stop
            End If
        End If
    End If
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '�˯���J��
    Do While iwe Is Nothing
        Set iwe = WD.FindElementByCssSelector("#search_input")
        If DateDiff("s", dt, VBA.Now) > 5 Then
            If backgroundStartChrome Then WD.Quit
            Exit Function
        End If
    Loop
    
    word.Application.windowState = wdWindowStateMinimize
    WD.SwitchTo.Window (WD.CurrentWindowHandle)
    VBA.Interaction.DoEvents
'    VBA.AppActivate "chrome"

    '����˯���J�ؤ���
    Dim keys As New SeleniumBasic.keys
    iwe.Clear
    iwe.SendKeys keys.Shift + keys.Insert '�K�W�˯�����
    iwe.SendKeys keys.Enter
    
    '�����˯����G
    Set iwe = Nothing
    '�������e������
    Set iwe = WD.FindElementByCssSelector("#shuoWenTable > tbody > tr:nth-child(2) > td:nth-child(2)")
    dt = VBA.Now
    Do While iwe Is Nothing
        Set iwe = WD.FindElementByCssSelector("#shuoWenTable > tbody > tr:nth-child(2) > td:nth-child(2)")
        If DateDiff("s", dt, VBA.Now) > 1.5 Then
            If backgroundStartChrome Then WD.Quit
            Exit Function
        End If
    Loop
    
    result(0) = iwe.text
    result(1) = WD.url
    LookupMultiFunctionChineseCharacterDatabase = result
    If backgroundStartChrome Then WD.Quit
    Exit Function
    
iweNothingExitFunction:
    If iwe Is Nothing Then
        LookupMultiFunctionChineseCharacterDatabase = result
        If backgroundStartChrome Then WD.Quit
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
                Set WD = Nothing
                Resume
            Else
                MsgBox Err.Number & Err.Description, vbExclamation
            End If
        Case Else
            MsgBox "������Chrome�s������A����@���I" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
    End Select
End Function

Rem �d�m������n�A���\�h�Ǧ^true
Function LookupYtenx(x As String) As Boolean
    If Not code.IsChineseString(x) Then
        MsgBox "�u������I", vbCritical
        Exit Function
    End If
    Dim iwe As SeleniumBasic.IWebElement, dt As Date, key As New SeleniumBasic.keys
    If Not OpenChrome("https://ytenx.org/") Then Exit Function
    word.Application.windowState = wdWindowStateMinimize
    WD.SwitchTo.Window WD.CurrentWindowHandle
    ActivateChrome
    '�˯���
    dt = VBA.Now
    Set iwe = WD.FindElementByCssSelector("#search-form > input.search-query.span3")
    Do While iwe Is Nothing
        Set iwe = WD.FindElementByCssSelector("#search-form > input.search-query.span3")
        If VBA.DateDiff("s", dt, VBA.Now) > 5 Then
            Exit Function
        End If
    Loop
    SetIWebElementValueProperty iwe, x
    iwe.SendKeys key.Enter
'    '�d�߫��s
'    Set iwe = WD.FindElementByCssSelector("")
'    If iwe Is Nothing Then Exit Function
'    iwe.Click
    LookupYtenx = True
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
    
    If Not OpenChrome("https://www.shuowen.org/") Then
        If Not OpenChrome("https://www.shuowen.org/") Then
            SystemSetup.killchromedriverFromHere
            Set SeleniumOP.WD = Nothing
            MsgBox "�ЦA���դ@�M�C���իe�A�нT�OChrome�s�����w�������C�P���P���@�n�L��������@�g���D", vbExclamation
            Stop
            Exit Function
        End If
    End If
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '�˯���J��
    Do While iwe Is Nothing
        Set iwe = WD.FindElementByCssSelector("#inputKaishu")
        If DateDiff("s", dt, VBA.Now) > 5 Then
            Exit Function
        End If
    Loop
    
    word.Application.windowState = wdWindowStateMinimize
    WD.SwitchTo.Window (WD.CurrentWindowHandle)
    VBA.Interaction.DoEvents
'    VBA.AppActivate "chrome"

    '����˯���J�ؤ���
    Dim keys As New SeleniumBasic.keys
    iwe.Clear
    iwe.SendKeys keys.Shift + keys.Insert '�K�W�˯�����
    iwe.SendKeys keys.Enter
    
    '�����˯����G
    Set iwe = WD.FindElementByCssSelector("body > div.container.main > div > div.col-md-9.main-content.pull-right > table > tbody > tr > td")
    If Not iwe Is Nothing Then
        If iwe.text = "�S���O��" Then
            Exit Function
        Else '�p�˯��u���v�r�A�]�ҿ��˦����A�G 20240924
            '�H�˯����G�M�椤�u���ѡv��W����ӧP�_
            'Set iwe = wd.FindElementByCssSelector("body > div.container.main > div > div.col-md-9.main-content.pull-right > table > thead > tr > th:nth-child(1)")
            '���˯����G�T���بӧP�_
            Set iwe = WD.FindElementByCssSelector("body > div.container.main > div > div.col-md-9.main-content.pull-right > div.row.paginator > div.col-md-4.info")
            GoSub iweNothingExitFunction
            'If iwe.Name = "����" Then
            Dim msg As String
            msg = iwe.GetAttribute("textContent")
            If VBA.IsNumeric(VBA.Left(msg, 1)) Then '�˯����G�T���ز�1�Ӧr�O�Ʀr
                '�q�`�i��|�H�˯����G�M�椤����1�����O�A�p�u���v�����G��2�Ӧr�u�x�v�A��Y²�Ʀr�G
                '���X�R�����ϥΪ̿�J��ƥH���ܭnŪ�J���C�A�u�n�w�]�Ȭ� 1 �Y�����Ī��ĪG 20240928 creedit_with_Copilot�j����
                Dim tb As SeleniumBasic.IWebElement
                Set tb = WD.FindElementByCssSelector("body > div.container.main > div > div.col-md-9.main-content.pull-right > table")

                Dim rows 'WebElements
                'Dim rows As SeleniumBasic.IWebElement
                rows = tb.FindElementsByTagName("tr")
                
                Dim cells 'IWebElements
                Dim r 'As Integer
'                For r = 1 To rows.Count
'                    Set cell = rows.item(i).FindElementByTag("td")
'                    Debug.Print cell.text
'                Next i

                word.Application.Activate
'                If VBA.vbOK = MsgBox(msg + vbCr + vbCr + "�˯����G����@���A�O�_�n���J�Ĥ@���������ơH", vbExclamation + vbOKCancel) Then
'                    '�˯����G�M�椤��1������������--�Y�r�Y
'                    Set iwe = wd.FindElementByCssSelector("body > div.container.main > div > div.col-md-9.main-content.pull-right > table > tbody > tr:nth-child(1) > td:nth-child(1) > a")
'                    GoSub iweNothingExitFunction
'                    iwe.Click
'                Else
'                    Exit Function
'                End If
reInput:
                r = VBA.InputBox(msg + vbCr + vbCr + _
                    "�˯����G����@���A�п�J�n���J�ĴX���������ơH�]����ơ^", "�нT�{�nŪ�J�ĴX�����m����n���e�C�w�]�Ȭ� 1 ", "1")
                If VBA.IsNumeric(r) = False Then
                    Exit Function
                ElseIf r > UBound(rows) Or r < 0 Then
                    If VBA.vbOK = MsgBox("��J���Ʀr����A�O�_�n���s��J�H", vbExclamation + vbOKCancel) Then
                        GoTo reInput
                    Else
                        Exit Function
                    End If
                Else
                    cells = rows(r).FindElementsByTagName("td")
                    Set iwe = cells(0)
                    GoSub iweNothingExitFunction
                    'iwe.Click '�L�@�Ρ]��쪺�A���u�O���G�Ȥ���A���O�u�������W������^
                    Dim outerHTML As String
                    outerHTML = iwe.GetAttribute("outerHTML")
                    If openNewTabWhenTabAlreadyExit(WD) Then
                        WD.Navigate.GoToUrl "https://www.shuowen.org" & VBA.Mid(outerHTML, VBA.InStr(outerHTML, "/"), VBA.InStr(outerHTML, " title=") - 1 - VBA.InStr(outerHTML, "/"))
                    End If
                End If
            End If
        End If
    End If
    '�����檺���e
    Set iwe = WD.FindElementByCssSelector("body > div.container.main > div > div.col-md-9.main-content.pull-right > div.row.summary > div.col-md-9.pull-left.info-container > div.media.info-body > div.media-body")
    GoSub iweNothingExitFunction
    result(0) = iwe.text
    result(1) = WD.url
    '���o�q�`�����e
    If includingDuan Then
        Dim i As Byte
        i = 1
        'Dim duanCommentary As String
        '���o�q�`�����e�ت�����
        Set iwe = WD.FindElementByCssSelector("body > div.container.main > div > div.col-md-9.main-content.pull-right > div:nth-child(" & i & ") > div")
        Do
            If i > 30 Then Exit Do
            If Not iwe Is Nothing Then
                If VBA.InStr(iwe.GetAttribute("textContent"), "�M�N �q�ɵ��m����Ѧr�`�n") Then Exit Do
            End If
            Set iwe = WD.FindElementByCssSelector("body > div.container.main > div > div.col-md-9.main-content.pull-right > div:nth-child(" & i & ") > div")
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
                SystemSetup.killchromedriverFromHere
                Set WD = Nothing
                Resume
            ElseIf InStr(Err.Description, "chromedriver.exe does not exist") Then 'The file C:\Program Files\Google\Chrome\Application\chromedriver.exe does not exist. The driver can be downloaded at http://chromedriver.storage.googleapis.com/index.html
                Set WD = Nothing
                MsgBox "�Цb�u" & getChromePathIncludeBackslash & "�v���|�U�ƻschromedriver.exe�ɮצA�~��I", vbCritical
                SystemSetup.OpenExplorerAtPath getChromePathIncludeBackslash
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
    
    If Not OpenChrome("https://dict.variants.moe.edu.tw/") Then
        If Not OpenChrome("https://dict.variants.moe.edu.tw/") Then
            Stop
        End If
    End If
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '�d�߿�J��
    Do While iwe Is Nothing
        Set iwe = WD.FindElementByCssSelector("#header > div > flex > div:nth-child(3) > div.quick > form > input[type=text]:nth-child(2)")
        If DateDiff("s", dt, VBA.Now) > 5 Then
            Exit Function
        End If
    Loop
    
    word.Application.windowState = wdWindowStateMinimize
    WD.SwitchTo.Window (WD.CurrentWindowHandle)
    VBA.Interaction.DoEvents
'    VBA.AppActivate "chrome"

    '���d�߿�J�ؤ���
    Dim keys As New SeleniumBasic.keys
    iwe.Clear
    SetIWebElementValueProperty iwe, x
    'iwe.SendKeys keys.Shift + keys.Insert '�K�W�˯�����
    iwe.SendKeys keys.Enter '���ɷ|����
    '�u�d�ߡv���s
    SystemSetup.wait 1
    Set iwe = WD.FindElementByCssSelector("#header > div > flex > div:nth-child(3) > div.quick > form > input[type=submit]:nth-child(5)")
    If Not iwe Is Nothing Then
        iwe.Submit
    End If
    
    dt = VBA.Now
    Set iwe = Nothing
    Do While iwe Is Nothing
        '�d�ߵ��G�T���ءA�p�i[ �] ]�A �d�ߵ��G�G���� 1 �r�A�����r 3 �r �j�����u1�v�o�Ӥ���A�H������ӧP�_
        Set iwe = WD.FindElementByCssSelector("body > main > div > flex > div:nth-child(1) > red:nth-child(1)")
        '�����������
        If Not WD.FindElementByCssSelector("#view > tbody > tr:nth-child(2) > th") Is Nothing Then
            Set iwe = WD.FindElementByCssSelector("#view > tbody > tr:nth-child(2) > th")
            If iwe.GetAttribute("textContent") = "��������" Then
                GoTo shuowenField
            End If
        End If
        Rem ��X�Ӫ����G�������G�G�@�O�C�X����B�����r�U�r�C�������A�G�O�����i�H�Ӧr���r�Y������
        If DateDiff("s", dt, VBA.Now) > 5 Then
            Exit Function
        End If
    Loop
    If Not iwe Is Nothing Then
        Dim zhengWen As String
        zhengWen = iwe.text '�e�Ҫ��u1�v
        '�e�Ҫ��u3�v
    
        dt = VBA.Now
        Set iwe = Nothing
        Do While iwe Is Nothing '��u���v�r
            Set iwe = WD.FindElementByCssSelector("body > main > div > flex > div:nth-child(1) > red:nth-child(2)")
            If DateDiff("s", dt, VBA.Now) > 5 Then
                Exit Function
            End If
        Loop

        If zhengWen <> "0" Or iwe.text <> "0" Then
            '�C�X����B�����r�U�r�C������
            Set iwe = WD.FindElementByCssSelector("#searchL > a")
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
                Set iwe = WD.FindElementByCssSelector("#searchL > a:nth-child(" & ai & ")")
                Do Until VBA.InStr(iwe.GetAttribute("outerHTML"), " data-tp=""��"" ")
                    ai = ai + 1
                    Set iwe = WD.FindElementByCssSelector("#searchL > a:nth-child(" & ai & ")")
                Loop
            End If
            iwe.Click
            '���ˬd �������� �x�s�� ������r�O�_�O�u�������Ρv
            Set iwe = WD.FindElementByCssSelector("#view > tbody > tr:nth-child(2) > th")
            GoSub iweNothingExitFunction
            If iwe.GetAttribute("textContent") <> "��������" Then
                Set iwe = Nothing
                result(0) = "�������ΨS����ơI"
                result(1) = WD.url
                GoSub iweNothingExitFunction
            End If
shuowenField:
            '�������� �x�s�椸��k�䪺�x�s��
            Set iwe = WD.FindElementByCssSelector("#view > tbody > tr:nth-child(2) > td")
            GoSub iweNothingExitFunction
            If IslinkImageIncluded���e�����]�t�W�s���ιϤ�(iwe) Then
                result(0) = iwe.GetAttribute("innerHTML")
            Else
                result(0) = iwe.GetAttribute("textContent")
            End If
            result(1) = WD.url
            SystemSetup.SetClipboard result(1)
        End If
    Else
        '�p�G������ܸӦr�����A�D�d�ߵ��G���A�p�G https://dict.variants.moe.edu.tw/dictView.jsp?ID=5565
        '�r�Y����
        Set iwe = WD.FindElementByCssSelector("#header > section > h2 > span > a")
        If iwe Is Nothing = False Then
        
            '���ˬd �������� �x�s�� ������r�O�_�O�u�������Ρv
            Set iwe = WD.FindElementByCssSelector("#view > tbody > tr:nth-child(2) > th")
            GoSub iweNothingExitFunction
            If iwe.GetAttribute("textContent") <> "��������" Then
                Set iwe = Nothing
                result(0) = "�������ΨS����ơI"
                result(1) = WD.url
                GoSub iweNothingExitFunction
            End If
            '�������� �x�s�椸��k�䪺�x�s��
            Set iwe = WD.FindElementByCssSelector("#view > tbody > tr:nth-child(2) > td")
            GoSub iweNothingExitFunction
            result(0) = iwe.GetAttribute("textContent")
            result(1) = WD.url
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
                Set WD = Nothing
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
    If Not OpenChrome("https://www.google.com") Then Exit Sub
    word.Application.windowState = wdWindowStateMinimize
    WD.SwitchTo.Window (WD.CurrentWindowHandle)
'    AppActivateChrome
    SeleniumOP.ActivateChrome
    VBA.Interaction.DoEvents
    Dim iwe As SeleniumBasic.IWebElement
    Dim keys As New SeleniumBasic.keys
    Set iwe = WD.FindElementByCssSelector("#APjFqb")
    If Not iwe Is Nothing Then
        iwe.Clear
'        SystemSetup.SetClipboard searchStr
'        iwe.SendKeys keys.Shift + keys.Insert
        SetIWebElementValueProperty iwe, searchStr
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
                    Set WD = Nothing
                    Resume
                Else
                    MsgBox Err.Number & Err.Description, vbExclamation
                End If
            Case Else
                MsgBox Err.Description, vbCritical
                SystemSetup.killchromedriverFromHere
                Set WD = Nothing
    '           Resume
        End Select

End Sub

'�K��j�y�Ŧ۰ʼ��I,�^���䵲�G�C�Y���ѡA�h�Ǧ^�Ŧr�� vbnullstring
Function grabGjCoolPunctResult(text As String, resultText As String, Optional Background As Boolean) As String
    Const url = "https://gj.cool/punct"
    Dim wdB As SeleniumBasic.IWebDriver, WBQuit As Boolean '=true �h�i�H��Chrome�s����
    Dim textBox As SeleniumBasic.IWebElement, btn As SeleniumBasic.IWebElement, btn2 As SeleniumBasic.IWebElement, item As SeleniumBasic.IWebElement
    Dim timeOut As Byte '�̦h�� timeOut ��
    On Error GoTo Err1
    
    If Background Then
        Rem ����
        Set wdB = openChromeBackground(url)
        WBQuit = True '�]���b�I������A�w�]�n�i�H��'�{�b�� .AddArgument "--remote-debugging-port=9222"  �ݮe���L�Ҷ}�Ҫ̡A�G�����A�I���F 20241003
        If wdB Is Nothing Then
            If WD Is Nothing Then
                If OpenChrome("https://gj.cool/punct") Then
                    Exit Function
                End If
            End If
            Set wdB = WD
        End If
    Else
        Rem ���
        If WD Is Nothing Then
            If Not OpenChrome("https://gj.cool/punct") Then
                Exit Function
            End If
        Else
            If IsWDInvalid() Then
                OpenNewTab WD
            Else
                'Stop 'just for test
                killchromedriverFromHere
                Set WD = Nothing
                If Not OpenChrome("https://gj.cool/punct") Then
                    Exit Function
                End If
            End If
        End If
        Set wdB = WD
        
    End If
    If wdB Is Nothing Or IsDriverInvalid(wdB) Then Exit Function
    If wdB.url <> url Then
        If Not IsNewBlankPageTab(wdB) Then OpenNewTab wdB
        wdB.Navigate.GoToUrl url
    End If
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
        If VBA.DateDiff("s", chkTxtTime, VBA.Now) > 1.8 Then
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
        If Not Background Then Set WD = Nothing
    End If
    'Debug.Print grabGjCoolPunctResult
    Exit Function
    
Err1:
        Select Case Err.Number
            Case 49 'DLL �I�s�W����~
                Resume
            Case 91 '�S���]�w�����ܼƩ� With �϶��ܼ�
                    killchromedriverFromHere
                    OpenChrome url
                    Set wdB = WD
                Resume
            Case -2146233088 'unknown error: ChromeDriver only supports characters in the BMP  (Session info: chrome=109.0.5414.75)
                Rem �����L�@��
                Rem systemsetup.SetClipboard text
                Rem SystemSetup.Wait 0.3
                Rem textBox.SendKeys key.Control + "v"
                Rem textBox.SendKeys key.LeftShift + key.Insert
                If InStr(Err.Description, "ChromeDriver only supports characters in the BMP") Then
                    WBQuit = pasteWhenOutBMP(wdB, url, "PunctArea", text, textBox, Background)
                    Resume Next
                ElseIf InStr(Err.Description, "invalid session id") Or InStr(Err.Description, "A exception with a null response was thrown sending an HTTP request to the remote WebDriver server for URL http://localhost:4609/session/455865a54d3f64364cf76b41fe7953a3/url. The status of the exception was ConnectFailure, and the message was: �L�k�s���ܻ��ݦ��A��") Then 'Or InStr(Err.Description, "no such window: target window already closed") Then
                    killchromedriverFromHere
                    OpenChrome url
                    Set wdB = WD: WBQuit = True
                    Resume
                ElseIf InStr(Err.Description, "no such window: target window already closed") Then
                    openNewTabWhenTabAlreadyExit wdB
                    wdB.Navigate.GoToUrl url
                    Resume
                ElseIf InStr(Err.Description, "disconnected: not connected to DevTools") = 1 Then 'disconnected: not connected to DevTools
                                                                                        '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                                                                        '  (Session info: chrome=130.0.6723.117)
                    killchromedriverFromHere 'WD.Quit: Set WD = Nothing:
                    OpenChrome url
                    Set wdB = WD
                    Resume
                    
                Else
                    MsgBox Err.Number & Err.Description
                    Stop
                End If
            Case -2147467261 '�å��N����Ѧҳ]�w�����󪺰������C
                If InStr(Err.Description, "�å��N����Ѧҳ]�w�����󪺰������C") Then
                    killchromedriverFromHere 'WD.Quit: Set WD = Nothing:
                     OpenChrome url
                    Set wdB = WD
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
Rem Ctrl + Alt + a : [AI�Ӫ�](https://t.shenshen.wiki/)���I 20241105
Function grabAITShenShenWikiPunctResult(text As String, resultText As String, Optional Background As Boolean) As String
        '��500�r
    Dim strInfo As New StringInfo, iwe As IWebElement, winState As WdWindowState
    strInfo.Create text
    If strInfo.LengthInTextElements > 500 Then
        MsgBox "��500�r", vbCritical
        Exit Function
    End If
    If IsWDInvalid() Then
        If WD Is Nothing Then
            If Not OpenChrome("https://t.shenshen.wiki/") Then Exit Function
        Else
            If IsChromeRunning Then
                WD.SwitchTo.Window (WD.WindowHandles()(UBound(WD.WindowHandles)))
            Else
                If Not OpenChrome("https://t.shenshen.wiki/") Then Exit Function
            End If
        End If
    Else
        LastValidWindow = WD.CurrentWindowHandle
    End If
    winState = word.Application.windowState
    WD.Navigate.GoToUrl "https://t.shenshen.wiki/"
    WD.SwitchTo.Window WD.CurrentWindowHandle
    ActivateChrome
    word.Application.windowState = wdWindowStateMinimize
    '���I
    Set iwe = WD.FindElementByCssSelector("#nav-biaodian-tab")
    Dim dt As Date
    dt = DateTime.Now
    Do While iwe Is Nothing
        Set iwe = WD.FindElementByCssSelector("#nav-biaodian-tab")
        If VBA.DateDiff("s", dt, DateTime.Now) > 5 Then Exit Function
    Loop
    iwe.Click
    '��J��
    Set iwe = WD.FindElementByCssSelector("#textarea-biaodian")
    If iwe Is Nothing Then Exit Function
    SetIWebElementValueProperty iwe, text
    '����
    Set iwe = WD.FindElementByCssSelector("#button-submit")
    If iwe Is Nothing Then Exit Function
    iwe.Click
    dt = DateTime.Now
    '���G��\�ˡH
    Set iwe = WD.FindElementByCssSelector("#feedback > div.feedback-button.feedback-tip")
    Do While Not iwe.Displayed ' Is Nothing
        If VBA.DateDiff("s", dt, DateTime.Now) > 36 Then Exit Function
        Set iwe = WD.FindElementByCssSelector("#feedback > div.feedback-button.feedback-tip")
    Loop
    '���G
    Set iwe = WD.FindElementByCssSelector("#output-content")
    If iwe Is Nothing Then Exit Function
    resultText = iwe.GetAttribute("textContent")
    If UBound(WD.WindowHandles) > 1 Then WD.Close '�������A�H��ʵ��q����I�}��
    If LastValidWindow <> vbNullString Then WD.SwitchTo().Window (LastValidWindow)
    grabAITShenShenWikiPunctResult = resultText
    word.Application.Activate
    word.Application.windowState = winState
End Function
Rem ���o�m�~�y�����Ʈw�P�_�y�Q�T�g�g��P�P���n�奻 �G gua ���W �C���\�h�Ǧ^ true 20241004
Function grabHanchiZhouYi_TheOriginalText_ThirteenSutras(gua As String, resultText As String) As Boolean

End Function

Rem ���o�m���Ǻ��P���g�e�P���f���n�奻�C���\�h�Ǧ^ true 20241004.20241006 resultText�O�Ӷ��X�A��1�Ӥ����O���������e�r��A��2�Ӥ����O�d�ߵ��G���}�C�Y�S���A�h�Ǧ^�����O�Ŧr�ꪺ�}�C
Function grabEeeLearning_IChing_ZhouYi_originalText(guaSequence As String, resultText As Variant, Optional iwe As SeleniumBasic.IWebElement) As Boolean
'    If Not OpenChrome("https://www.eee-learning.com/article/571") Then Exit Function
    If Not VBA.IsArray(resultText) Then
        MsgBox "��2�Ӥ޼ƥ����O�r��}�C", vbCritical
        'grabEeeLearning_IChing_ZhouYi_originalText = False'�w�]�Y��false
    Else
        If UBound(resultText) <> 1 Then
            MsgBox "��2�Ӥ޼ƥ����O2�Ӥ������r��}�C", vbCritical
        End If
    End If
    Dim e2 As String
    e2 = "https://www.eee-learning.com/book/eee" & guaSequence
    If Not OpenChrome(e2) Then Exit Function
    
    grabEeeLearning_IChing_ZhouYi_originalText = True
    
    'Dim iwe As SeleniumBasic.IWebElement
    Set iwe = WD.FindElementByCssSelector("#block-bartik-content > div > article > div > div.clearfix.text-formatted.field.field--name-body.field--type-text-with-summary.field--label-hidden.field__item")
    If iwe Is Nothing Then
        grabEeeLearning_IChing_ZhouYi_originalText = False
        Exit Function
    End If
    
    resultText(0) = iwe.GetAttribute("textContent")
    resultText(1) = e2
    
End Function
Rem 20240914 creedit_with_Copilot�j���ġGhttps://sl.bing.net/gCpH6nC61Cu
' �]�w���� IWebElement��value�ݩʭ�  20240913
Function SetIWebElementValueProperty(iwe As IWebElement, txt As String) As Boolean
    If Not iwe Is Nothing Then
        'driver.ExecuteScript "arguments[0].value = arguments[1];", element, valueToSet
        WD.ExecuteScript "arguments[0].value = arguments[1];", iwe, txt
        SetIWebElementValueProperty = True
    End If
End Function
Rem 20240914 creedit_with_Copilot�j���ġGhttps://sl.bing.net/gCpH6nC61Cu
' �]�w���� IWebElement��value�ݩʭ�  20240913
Function SetIWebElement_textContent_Property(iwe As IWebElement, txt As String) As Boolean
    If Not iwe Is Nothing Then
        'driver.ExecuteScript "arguments[0].value = arguments[1];", element, valueToSet
        WD.ExecuteScript "arguments[0].textContent = arguments[1];", iwe, txt
        SetIWebElement_textContent_Property = True
    End If
End Function


Private Function pasteWhenOutBMP(ByRef iwd As SeleniumBasic.IWebDriver, url, textBoxToPastedID, pastedTxt As String, ByRef textBox As SeleniumBasic.IWebElement, Background As Boolean) As Boolean ''unknown error: ChromeDriver only supports characters in the BMP  (Session info: chrome=109.0.5414.75)
Rem creedit chatGPT�j���ġG�z���쪺�T��O Selenium �� SendKeys ��k����K�W BMP �~���r�����D�C
On Error GoTo Err1
Dim retryTimes As Byte
DoEvents
'systemsetup.SetClipboard pastedTxt
'SystemSetup.Wait 0.2
If Background Then iwd.Quit
retry:
If iwd Is Nothing Then
    If WD Is Nothing Then
        OpenChrome (url)
        pasteWhenOutBMP = True
    End If
    Set iwd = WD
End If
If iwd.url <> url Then iwd.Navigate.GoToUrl (url)
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
            Set WD = Nothing
            GoTo retry
        Case Else
            If WD Is Nothing Then
                OpenChrome (url)
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
    For Each windowHandle In WD.WindowHandles
        If i = index Then
            WindowHandlesItem = windowHandle
            Exit Function
        End If
        i = i + 1
    Next
End Function

Public Property Get WindowHandlesCount() As Long
    WindowHandlesCount = UBound(WD.WindowHandles) + 1
End Property
Public Property Get WindowHandles() As String()
    On Error GoTo eH:
    If Not WD Is Nothing Then WindowHandles = WD.WindowHandles
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

Rem 20241009 creedit_with_Copilot�j���ġGWordVBA+SeleniumBasicŪ�J�������e�Ϥ��P�W�s���Ghttps://sl.bing.net/hxRfMU08232
Function IslinkImageIncluded���e�����]�t�W�s���ιϤ�(iwe As SeleniumBasic.IWebElement) As Boolean
    Dim hasLinks As Boolean, arr
    Dim hasImages As Boolean
    
    
    ' �P�_�O�_�]�t�W�s��
    arr = iwe.FindElementsByTagName("a")
    If VBA.IsArray(arr) Then
        links_arrayIWebElement = arr
    End If
    hasLinks = UBound(links_arrayIWebElement) > -1
    
    ' �P�_�O�_�]�t�Ϥ�
    arr = iwe.FindElementsByTagName("img")
    If VBA.IsArray(arr) Then
        images_arrayIWebElement = arr
    End If
    hasImages = UBound(images_arrayIWebElement) > -1
    
    ' ��^���G
    IslinkImageIncluded���e�����]�t�W�s���ιϤ� = hasLinks Or hasImages
End Function
Rem 20241009 creedit_with_Copilot�j���ġGWordVBA+SeleniumBasicŪ�J�������e�Ϥ��P�W�s���Ghttps://sl.bing.net/hxRfMU08232
Function IslinkIncluded���e�����]�t�W�s��(iwe As SeleniumBasic.IWebElement) As Boolean
    Dim hasLinks As Boolean, links
        
    ' �P�_�O�_�]�t�W�s��
    links = iwe.FindElementsByTagName("a")
    If VBA.IsArray(links) Then
        links_arrayIWebElement = links
    End If
    hasLinks = UBound(links_arrayIWebElement) > -1
    ' ��^���G
    IslinkIncluded���e�����]�t�W�s�� = hasLinks
End Function
Rem 20241009 creedit_with_Copilot�j���ġGWordVBA+SeleniumBasicŪ�J�������e�Ϥ��P�W�s���Ghttps://sl.bing.net/hxRfMU08232
Function IsImageIncluded���e�����]�t�Ϥ�(iwe As SeleniumBasic.IWebElement) As Boolean

    Dim hasImages As Boolean, imgs
    ' �P�_�O�_�]�t�Ϥ�
    imgs = iwe.FindElementsByTagName("img")
    If VBA.IsArray(imgs) Then
        images_arrayIWebElement = imgs
    End If
    'On Error Resume Next
    hasImages = UBound(images_arrayIWebElement) > -1
    'On Error GoTo 0
    ' ��^���G
    IsImageIncluded���e�����]�t�Ϥ� = hasImages
End Function
Property Get images() As SeleniumBasic.IWebElement()
    images = images_arrayIWebElement
End Property
Property Get links()
    links = links_arrayIWebElement
End Property

Rem ����������e�ðO����m 20241009 creedit_with_Copilot�j���ġGhttps://sl.bing.net/gGHK9dMCbNQ
Sub grabPageContent����������e�ðO����m()
    'Dim wd As New SeleniumBasic.ChromeDriver
    Dim elements() As SeleniumBasic.IWebElement 'Object
    Dim e
    Dim element As SeleniumBasic.IWebElement 'As Object
    Dim elementPositions As Collection
    Set elementPositions = New Collection
    
    '' ���}����
    'wd.start "Chrome"
    'wd.Get "https://www.eee-learning.com/article/3694"
    
    ' ����Ҧ����e�ðO����m
    elements = WD.FindElementsByCssSelector("body *")
    
    For Each e In elements
        Dim elementInfo As New Collection
        Set element = e
        elementInfo.Add element.tagname
        elementInfo.Add element.text
        elementInfo.Add element.GetAttribute("src")
        elementInfo.Add element.GetAttribute("href")
        elementPositions.Add elementInfo
    Next e
    
    'wd.Quit
    
    ' �N����쪺���e���J��Word���
    inputElementContent���J�������󳡤����e elementPositions
End Sub
Sub inputElementContent���J�������󳡤����e(elementPositions As Collection)
    Dim elementInfo As Collection, e
    Dim rng As Range
    
    For Each e In elementPositions
        Set elementInfo = e
        Select Case elementInfo(1)
            Case "P", "DIV"
                ' ���J��r
                Set rng = Selection.Range
                rng.text = elementInfo(2)
                Selection.MoveRight Unit:=wdCharacter, Count:=1
            Case "IMG"
                ' ���J�Ϥ�
                Set rng = Selection.Range
                rng.InlineShapes.AddPicture fileName:=elementInfo(3), _
                                            LinkToFile:=False, SaveWithDocument:=True
                Selection.MoveRight Unit:=wdCharacter, Count:=1
            Case "A"
                ' ���J�W�s��
                Set rng = Selection.Range
                ActiveDocument.Hyperlinks.Add Anchor:=rng, _
                                              Address:=elementInfo(4), _
                                              TextToDisplay:=elementInfo(2)
                Selection.MoveRight Unit:=wdCharacter, Count:=1
        End Select
    Next e
End Sub

Rem �M���S�w�d�򤺪��Ҧ������A�æb�S�w��m���J�Ϥ��M�W�s�� https://sl.bing.net/f4Mv2PVPse4 20241009 creedit_with_Copilot�j���ġG
Sub inputElementContentAll���J��������Ҧ������e(iwe As SeleniumBasic.IWebElement)
    'Dim wd As New SeleniumBasic.ChromeDriver
    'Dim iwe As SeleniumBasic.IWebElement
    Dim elements() As SeleniumBasic.IWebElement  'Object
    Dim e, element As SeleniumBasic.IWebElement 'Object
    Dim rng As Range
    
    ' ���}����
'    wd.start "Chrome"
'    wd.Get "https://www.eee-learning.com/article/3694"
    
    ' ����S�w���e����
    'Set iwe = wd.FindElementByCssSelector("#block-bartik-content > div > article > div > div.clearfix.text-formatted.field.field--name-body.field--type-text-with-summary.field--label-hidden.field__item")
    
    ' �M���S�w�d�򤺪��Ҧ�����
    elements = iwe.FindElementsByTagName("*")
    
    For Each e In elements
        Set element = e
        
'        Stop
        
'        If SeleniumOP.IsImageIncluded���e�����]�t�Ϥ�(element) Then
'            insertImageInline���J�]�t�Ϥ����q�� element
'        End If
        Select Case element.tagname
            Case "p", "div", "span" '"P", "DIV", "STRONG", "SPAN"
                ' "strong" ���� .Bold=true ��ѩI�s�ݳB�z
                ' ���J��r
                Set rng = Selection.Range
                If element.tagname = "strong" Then
                    
                Else
                    rng.text = element.GetAttribute("textContent") 'element.text
                End If
                If SeleniumOP.IsImageIncluded���e�����]�t�Ϥ�(element) Then
                    'insertImageInline���J�]�t�Ϥ����q�� element
                    rng.Find.Execute (ChrW(160))
'                    If rng.Find.Execute(ChrW(160)) Then
                        rng.InlineShapes.AddPicture fileName:=images_arrayIWebElement()(0).GetAttribute("src"), LinkToFile:=False, SaveWithDocument:=True
                        If rng.Find.Found Then
                            Selection.MoveDown
                        Else
                            Selection.MoveRight Unit:=wdCharacter, Count:=1
                        End If
'                    End If
                End If
                
                'Selection.MoveRight Unit:=wdCharacter, Count:=1
            Case "img"
                ' ���J�Ϥ�
                If UBound(images_arrayIWebElement) > -1 Then
                    'If Not e.GetAttribute("src") = images_arrayIWebElement()(0).GetAttribute("src") Then
                    If Not (e.GetAttribute("x") = images_arrayIWebElement()(0).GetAttribute("x") And e.GetAttribute("y") = images_arrayIWebElement()(0).GetAttribute("y")) Then
                        Set rng = Selection.Range
                        rng.InlineShapes.AddPicture fileName:=element.GetAttribute("src"), _
                                                    LinkToFile:=False, SaveWithDocument:=True
                        Selection.MoveRight Unit:=wdCharacter, Count:=1
                    End If
                Else
                    Set rng = Selection.Range
                    rng.InlineShapes.AddPicture fileName:=element.GetAttribute("src"), LinkToFile:=False, SaveWithDocument:=True
                    Selection.MoveRight Unit:=wdCharacter, Count:=1
                End If
                
            Case "a"
'                ' ���J�W�s�� rem ���浹�I�s�ݰ�
'                Set rng = Selection.Range
'                ActiveDocument.Hyperlinks.Add Anchor:=rng, _
'                                              Address:=element.GetAttribute("href"), _
'                                              TextToDisplay:=element.text
'                Selection.MoveRight Unit:=wdCharacter, Count:=1
        End Select
    Next e
    
'    wd.Quit
End Sub

Rem �B�z�]�t�Ϥ����q���A���]�t�Ϥ����q���A�z�ݭn������Ӭq������r���e�A�M��b���J�Ϥ��ɽվ㴡�J�I����m�C 20241009 creedit_with_Copilot�j���ġGWordVBA+SeleniumBasicŪ�J�������e�Ϥ��P�W�s���Ghttps://sl.bing.net/2k4A3xjuh2
Private Sub insertImageInline���J�]�t�Ϥ����q��(iwe As SeleniumBasic.IWebElement)
    'Dim wd As New SeleniumBasic.ChromeDriver
    'Dim
    Dim elements As Object
    Dim element As Object
    Dim rng As Range
    Dim textParts() As String
    Dim i As Integer
    
'    ' ���}����
'    wd.start "Chrome"
'    wd.Get "https://www.eee-learning.com/article/3694"
    
'    ' ����S�w�q��
    'Set iwe = wd.FindElementByCssSelector("#block-bartik-content > div > article > div > div.clearfix.text-formatted.field.field--name-body.field--type-text-with-summary.field--label-hidden.field__item > p:nth-child(2)")
    
    ' ���άq����r���e
    textParts = Split(iwe.GetAttribute("innerHTML"), "<img")
    '�p �G 3 &nbsp;�@<img style="border-width:0;" src="/image/yi03b.png" width="28" height="28" align="absbottom" border="0">�@<strong>�٨�</strong>�@���p��
    
    ' ���J��r�M�Ϥ�
    For i = LBound(textParts) To UBound(textParts)
        If i > 0 Then
            ' ���J�Ϥ�
            Set rng = Selection.Range
            rng.InlineShapes.AddPicture images_arrayIWebElement()(0).GetAttribute("src"), LinkToFile:=False, SaveWithDocument:=True
'            rng.InlineShapes.AddPicture fileName:=Mid(textParts(i), InStr(textParts(i), "src=") + 5, InStr(textParts(i), """", InStr(textParts(i), "src=") + 5) - InStr(textParts(i), "src=") - 5), _
                                        LinkToFile:=False, SaveWithDocument:=True
            Selection.MoveRight Unit:=wdCharacter, Count:=1
        End If
        ' ���J��r
        Set rng = Selection.Range
        rng.text = iwe.GetAttribute("textContent") 'StripHTMLTags(textParts(i))
        Selection.MoveRight Unit:=wdCharacter, Count:=1
    Next i
    
'    wd.Quit
End Sub

Function StripHTMLTags(html As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "<.*?>"
    regex.Global = True
    StripHTMLTags = regex.Replace(html, "")
End Function
Rem �P�_�Ϥ����J�����_ 20241021creedit_with_Copilot�j����:�n�P�_�Ϥ��O�_���J���\�A�i�H�z�L�ˬd�Ϥ������� complete �ݩʩΪ̬O��ť�Ϥ��� load �ƥ�C�o�̦��@��²�檺��k�A�Q�� SeleniumBasic �� JavaScript ����\����ˬd�Ϥ��O�_�w�g���J���\�G
Private Function isImageLoaded(CssSelector As String) As Boolean
    Dim script As String
    Dim imgLoaded As Boolean
    
    ' JavaScript �N�X���ˬd�Ϥ��O�_���J���\
    script = "return document.querySelector('" & CssSelector & "').complete;" '�o�q�{���X�|�ˬd���w���Ϥ������� complete �ݩʡA�p�G�Ϥ��w�g���J�Acomplete �ݩʷ|�O true�A�_�h�|�O false�C
'    script = "return document.querySelector('body > div:nth-child(3) > center:nth-child(2) > img').complete;"
    ' ���� JavaScript �è��o���G
    imgLoaded = WD.ExecuteScript(script)
    '�p�G�Ϥ���ܪ��O���N��r�A�q�`�N���۹Ϥ��s�����ѩθ��J���~�C�b�o�ر��p�U�A�i�H�ˬd�Ϥ��� naturalWidth �M naturalHeight �ݩʡC�p�G�o����ݩʳ��j�� 0�A�h�Ϥ����J���\�F�_�h�A���J���ѡC
    '�o�q�{���X�|�ˬd���w�Ϥ��� naturalWidth �M naturalHeight �ݩʡA�p�G�o����ݩʳ��j�� 0�A��ܹϤ����J���\�F�_�h�A�Ϥ����J���ѡC�o�����ӯ���T�a�P�_�Ϥ��O�_���J���\�C
    If imgLoaded Then
        ' JavaScript �N�X���ˬd�Ϥ��� naturalWidth �M naturalHeight �ݩ�
        script = "var img = document.querySelector('body > div:nth-child(3) > center:nth-child(2) > img');return (img.naturalWidth > 0 && img.naturalHeight > 0);"
        ' ���� JavaScript �è��o���G
        imgLoaded = WD.ExecuteScript(script)
    End If
    If imgLoaded Then
        isImageLoaded = True 'MsgBox "�Ϥ��w���J���\"
    Else
        isImageLoaded = False 'MsgBox "�Ϥ��|�����J"
    End If
End Function
