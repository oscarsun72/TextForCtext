Attribute VB_Name = "SeleniumOP"
Option Explicit
Public WD As SeleniumBasic.IWebDriver
'Private Const timeoutsImplicitWait As Long = 0 '預設值為0 但實測為出錯，必須予一有效值供還原設定
Private Const timeoutsPageLoad As Long = 300 ''預設值為300    20241020
Public chromedriversPID() As Long '儲存chromedriver程序ID的陣列
Public chromedriversPIDcntr As Integer 'chromedriversPID的下標值
Public ActiveXComponentsCanNotBeCreated As Boolean

Const SW_RESTORE = 9
Const SW_MAXIMIZE = 3
Const WM_SYSCOMMAND = &H112
Const SC_RESTORE = &HF120
Const VK_CONTROL = &H11
Const VK_MENU = &H12
Const VK_SPACE = &H20
Const VK_DOWN = &H28
Const VK_RETURN = &HD
Const KEYEVENTF_KEYUP = &H2

'' 宣告 Windows API 函數 20241003 creedit_with_Copilot大菩薩：改進WordVBA+SeleniumBasic 開啟Chrome瀏覽器新分頁的方法：https://sl.bing.net/iqY5XH1MVci
'Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
'Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function ShowWindowAsync Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As LongPtr)

#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function ShowWindowAsync Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare Function GetForegroundWindow Lib "user32" () As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

#End If

'Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
'Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Private last_ValidWindow As String
Private images_arrayIWebElement() As SeleniumBasic.IWebElement
Private links_arrayIWebElement() As SeleniumBasic.IWebElement
Rem 20241005 Copilot大菩薩：WordVBA 判斷 Word 是否為最前端視窗：https://sl.bing.net/b21Z8KIK3Ua
Function IsWordActive() As Boolean
    Dim hWnd As LongPtr
    Dim title As String * 255
    Dim Length As Long
    
    ' 獲取當前活動視窗的句柄
    hWnd = GetForegroundWindow()
    
    ' 獲取視窗標題
    Length = GetWindowText(hWnd, title, Len(title))
    title = Left(title, Length)
    
    ' 檢查標題是否包含 "Microsoft Word"
    If InStr(title, "Microsoft Word") > 0 Then
        IsWordActive = True
    Else
        IsWordActive = False
    End If
End Function

' 模擬按鍵操作
Sub SimulateKeyStroke(vk As Byte)
    keybd_event vk, 0, 0, 0
    keybd_event vk, 0, KEYEVENTF_KEYUP, 0
End Sub
' 將 Chrome 瀏覽器設置為前端窗口並恢復最小化或隱藏到托盤的窗口
Sub ActivateChrome()
    VBA.DoEvents
    Dim hWnd As LongPtr
    Dim title As String * 256
    
    hWnd = FindWindow("Chrome_WidgetWin_1", vbNullString)
    If hWnd <> 0 Then
        GetWindowText hWnd, title, 256
        If InStr(title, "Chrome") > 0 Then ' 確保窗口標題包含 "Chrome"
            If IsIconic(hWnd) Or Not IsWindowVisible(hWnd) Then
                SetForegroundWindow hWnd
'                SimulateKeyStroke VK_MENU ' 模擬按下 Alt 鍵
'                SimulateKeyStroke VK_SPACE ' 模擬按下 Space 鍵
'                SimulateKeyStroke VK_DOWN ' 模擬按下箭頭向下鍵
'                SimulateKeyStroke VK_DOWN ' 再次模擬按下箭頭向下鍵
'                SimulateKeyStroke VK_RETURN ' 模擬按下 Enter 鍵
                VBA.DoEvents ' 確保窗口已經顯示
                ShowWindow hWnd, SW_MAXIMIZE ' 最大化窗口
            End If
        End If
    Else
        hWnd = FindWindow(vbNullString, "Google Chrome")
        If hWnd <> 0 Then
            GetWindowText hWnd, title, 256
            If InStr(title, "Chrome") > 0 Then ' 確保窗口標題包含 "Chrome"
                If IsIconic(hWnd) Or Not IsWindowVisible(hWnd) Then
                    SetForegroundWindow hWnd
                    SimulateKeyStroke VK_MENU ' 模擬按下 Alt 鍵
                    SimulateKeyStroke VK_SPACE ' 模擬按下 Space 鍵
                    SimulateKeyStroke VK_DOWN ' 模擬按下箭頭向下鍵
                    SimulateKeyStroke VK_DOWN ' 再次模擬按下箭頭向下鍵
                    SimulateKeyStroke VK_RETURN ' 模擬按下 Enter 鍵
                    VBA.DoEvents ' 確保窗口已經顯示
                    ShowWindow hWnd, SW_MAXIMIZE ' 最大化窗口
                End If
            End If
        Else
            MsgBox "Could not find Chrome window", vbCritical
        End If
    End If
End Sub

' 將 Chrome 瀏覽器設置為前端窗口並恢復最小化的窗口 20241204 Copilot大菩薩：這段程式碼現在同時檢查窗口是否最小化（IsIconic）以及是否可見（IsWindowVisible），並使用 ShowWindow 將其恢復。這樣即使 Chrome 瀏覽器在任務欄中或處於最小化狀態，程式碼都能將其恢復並設置為前台窗口。
Sub ActivateChromeOK()
    VBA.DoEvents
    Dim hWnd As LongPtr
    Dim title As String * 256

    hWnd = FindWindow("Chrome_WidgetWin_1", vbNullString)
    If hWnd <> 0 Then
        GetWindowText hWnd, title, 256 '在這段程式碼中，我添加了 GetWindowText 函數來獲取窗口標題，然後使用 InStr 函數來檢查窗口標題是否包含 "Chrome"。這樣可以確保找到的窗口確實是 Chrome 瀏覽器，而不是其他應用程序，例如 Dropbox。
                                            '可能是因為 FindWindow 函數在查找窗口時，偶爾會找到非目標應用程序的窗口，例如 Dropbox。為了解決這個問題，您可以嘗試通過窗口標題進一步篩選，以確保找到的是正確的 Chrome 瀏覽器窗口。
        If InStr(title, "Chrome") > 0 Then ' 確保窗口標題包含 "Chrome"
            If IsIconic(hWnd) Or Not IsWindowVisible(hWnd) Then
                ShowWindowAsync hWnd, SW_RESTORE ' 恢復最小化或隱藏的窗口
                VBA.DoEvents ' 確保窗口已經顯示
                ShowWindow hWnd, SW_MAXIMIZE ' 最大化窗口
            End If
            SetForegroundWindow hWnd
        End If
    Else
        hWnd = FindWindow(vbNullString, "Google Chrome")
        If hWnd <> 0 Then
            GetWindowText hWnd, title, 256
            If InStr(title, "Chrome") > 0 Then ' 確保窗口標題包含 "Chrome"
                If IsIconic(hWnd) Or Not IsWindowVisible(hWnd) Then
                    ShowWindowAsync hWnd, SW_RESTORE ' 恢復最小化或隱藏的窗口
                    VBA.DoEvents ' 確保窗口已經顯示
                    ShowWindow hWnd, SW_MAXIMIZE ' 最大化窗口
                End If
                SetForegroundWindow hWnd
            End If
        Else
            MsgBox "Could not find Chrome window", vbCritical
        End If
    End If
End Sub

' 將 Chrome 瀏覽器設置為前端窗口並恢復最小化的窗口
Sub ActivateChrome_minim()
    VBA.DoEvents
    Dim hWnd As LongPtr
    hWnd = FindWindow("Chrome_WidgetWin_1", vbNullString)
    If hWnd <> 0 Then
        If IsIconic(hWnd) Then ShowWindow hWnd, SW_RESTORE ' 恢復最小化的窗口
        SetForegroundWindow hWnd
    Else
        hWnd = FindWindow(vbNullString, "Google Chrome")
        If hWnd <> 0 Then
            If IsIconic(hWnd) Then ShowWindow hWnd, SW_RESTORE ' 恢復最小化的窗口
            SetForegroundWindow hWnd
        Else
            MsgBox "Could not find Chrome window", vbCritical
        End If
    End If
End Sub

' 將 Chrome 瀏覽器設置為前端窗口
Sub ActivateChrome_noRestore()
    VBA.DoEvents
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
Rem 20241008 Copilot大菩薩：https://sl.bing.net/cmQuvtGT28O
Rem Gemini大菩薩就不行了！https://sl.bing.net/cmQuvtGT28O
Property Let LastValidWindow(validWindowHandle As String)
    last_ValidWindow = validWindowHandle
End Property
'Sub tesSeleniumBasic() 'https://github.com/florentbr/SeleniumBasic
''20230119 creedit chatGPT大菩薩
'
'    Dim driver As New Selenium.WebDriver
'    'driver.start "chrome", "https://www.google.com"
'    driver.SetBinary getChromePathIncludeBackslash
'    driver.start getChromePathIncludeBackslash + "chrome.exe", "https://www.google.com"
'    driver.Get "/"
'
'End Sub
Rem 失敗時傳回false
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
    Select Case Err.number
        Case -2146233088
            If InStr(Err.description, "no such window: target window already closed") Then
                If iw > 0 Then
                    For Each ew In WD.WindowHandles
                        Exit For
                    Next ew
                    WD.SwitchTo.Window (ew)
                    Resume
                Else
                    Stop
                End If
            ElseIf InStr(Err.description, "ot connected to DevTools") Then
'                disconnected: not connected to DevTools
'                (failed to check if window was closed: disconnected: not connected to DevTools)
'                (Session info: chrome=127.0.6533.120)
                If Not WD Is Nothing Then WD.Quit
                Set WD = Nothing
                killchromedriverFromHere
                MsgBox "若Chrome瀏覽器已開啟，請關閉Chrome瀏覽器後再按確定", vbExclamation
                OpenChrome "https://www.google.com"
                Resume 'GoTo reOpenChrome:
            ElseIf InStr(Err.description, "A exception with a null response was thrown sending an HTTP") Then
'                A exception with a null response was thrown sending an HTTP request to the remote WebDriver server for URL http://localhost:1760/session/ed5864479325c154783256563f97e610/window/handles. The status of the exception was ConnectFailure, and the message was: 無法連接至遠端伺服器
                Set WD = Nothing
                killchromedriverFromHere
                MsgBox "若Chrome瀏覽器已開啟，請關閉Chrome瀏覽器後再按確定", vbExclamation
'                openChrome "https://www.google.com"
'                Resume 'GoTo reOpenChrome:
                openNewTabWhenTabAlreadyExit = False
            ElseIf InStr(Err.description, "invalid session id") Then
                killchromedriverFromHere
                Set WD = Nothing
                OpenChrome "https://www.google.com.tw/"
                'Resume
            Else
                MsgBox Err.number & Err.description
                Stop
            End If
        Case -2147467261
            If Err.description = "並未將物件參考設定為物件的執行個體。" Then
                Set WD = Nothing
                killchromedriverFromHere
                MsgBox "若Chrome瀏覽器已開啟，請關閉Chrome瀏覽器後再執行一次", vbExclamation
                openNewTabWhenTabAlreadyExit = False
            Else
                Stop
            End If
        Case 91
            If Err.description = "沒有設定物件變數或 With 區塊變數" Then
                Set WD = Nothing
                killchromedriverFromHere
                MsgBox "若Chrome瀏覽器已開啟，請關閉Chrome瀏覽器後再執行一次", vbExclamation
                openNewTabWhenTabAlreadyExit = False
            Else
                Stop
            End If
        Case Else
            MsgBox Err.description, vbCritical
            WD.Quit
            SystemSetup.killchromedriverFromHere
    '           Resume
    End Select
End Function

Rem 檢查 driver 是否有效 20241002
Function IsDriverInvalid(ByRef driver As IWebDriver) As Boolean
    On Error Resume Next
    Dim url As String
    url = driver.url
    IsDriverInvalid = (url = vbNullString Or (driver Is Nothing))
End Function
Rem 檢查 wd 是否有效 20241002
Function IsWDInvalid() As Boolean
    On Error Resume Next
    Dim url As String
    url = WD.url
    IsWDInvalid = (url = vbNullString Or (WD Is Nothing))
End Function

Rem 檢查是否為新的空白頁 開啟的新分頁 20241003
Function IsNewBlankPageTab(ByRef driver As IWebDriver) As Boolean
    'On Error Resume Next
    IsNewBlankPageTab = (driver.url = "about:blank" Or WD.title = vbNullString) _
                Or (WD.title = "新分頁" Or WD.url = "chrome://new-tab-page/")
End Function
Rem 啟動Chrome瀏覽器或已啟動後開啟新分頁瀏覽。失敗時傳回false
Function OpenChrome(Optional url As String) As Boolean
reStart:
        'Dim WD As SeleniumBasic.IWebDriver
        On Error GoTo ErrH
        Dim Service As SeleniumBasic.ChromeDriverService
        Dim options As SeleniumBasic.ChromeOptions
        Dim pid As Long
    
    '結束chromedriver.exe
    '使用 WMI 和上面所述的方法
    '判斷PID是否等於pid
    
        If WD Is Nothing Then
        
            If IsChromeRunning Then '20241002
                If Not OpenChrome_NEW_Get Then Exit Function
                If WD Is Nothing Then
                    'Stop ' for test
                    If MsgBox("請關閉Chrome瀏覽器後再按「確定」繼續。否則請按「取消」以取消作業。", vbExclamation + vbOKCancel) = VBA.vbCancel Then Exit Function
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
                        If VBA.Dir("W:\PortableApps\PortableApps\GoogleChromePortable\App\Chrome-bin\") <> vbNullString Then
                            chromePath = "W:\PortableApps\PortableApps\GoogleChromePortable\App\Chrome-bin\"
                        Else
                            chromePath = "W:\PortableApps\PortableApps\GoogleChromePortable64\App\Chrome-bin\"
                        End If
                    End If
        
                With Service
                    .CreateDefaultService driverPath:=chromePath 'getChromePathIncludeBackslash
                    '.CreateDefaultService driverPath:="E:\Selenium\Drivers"
                    .HideCommandPromptWindow = True '不顯示命令提示字元視窗
                End With
                Set options = New SeleniumBasic.ChromeOptions
                With options
                    .BinaryLocation = chromePath + "chrome.exe"
                    .AddExcludedArgument "enable-automation" '禁用「Chrome 正在被自動化軟體控制」的警告消息
                    
                    'C#：options.AddArgument("user-data-dir=" + Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\Google\\Chrome\\User Data\\");
                    .AddArgument "user-data-dir=" + VBA.Environ("LOCALAPPDATA") + _
                        "\Google\Chrome\User Data\"
                    .AddArgument "--new-window"
                    '.AddArgument "--start-maximized"
                    '.DebuggerAddress = "127.0.0.1:9999" '不要与其他几個混用
                    
                    .AddArgument "--remote-debugging-port=9222" '20241002 Copilot大菩薩：Word VBA 中的 Selenium 操作：https://sl.bing.net/SMTsa6sktU
                                        
                End With
                WD.New_ChromeDriver Service:=Service, options:=options
                Docs.Register_Event_Handler '為清除chromedriver作準備
                pid = Service.ProcessId 'Chrome瀏覽器沒有開成功就會是0
                If pid <> 0 Then
                    ReDim Preserve chromedriversPID(chromedriversPIDcntr)
                    chromedriversPID(chromedriversPIDcntr) = pid
                    chromedriversPIDcntr = chromedriversPIDcntr + 1
                End If
                If url = "https://gj.cool/punct" Then
                    OpenNewTab WD, url '前有.AddArgument "--new-window" 20241005 此是 window 不是 tab !!
                Else
                    OpenNewTab WD
                End If
'                WD.ExecuteScript "window.open('about:blank','_blank');" 'openNewTabWhenTabAlreadyExit WD
'                WD.SwitchTo.Window WindowHandlesItem(WindowHandlesCount - 1)
                WD.url = url
            End If
        Else
            If IsWDInvalid() Then
                If url = "https://gj.cool/punct" Then
                    OpenNewTab WD, url
                Else
                    OpenNewTab WD
                End If
            End If
'            WD.ExecuteScript "window.open('about:blank','_blank');" 'openNewTabWhenTabAlreadyExit WD
'            WD.SwitchTo.Window WindowHandlesItem(WindowHandlesCount - 1)
            WD.url = url
        End If
        If ActiveXComponentsCanNotBeCreated Then ActiveXComponentsCanNotBeCreated = False
        
      
        OpenChrome = True
    Exit Function
ErrH:
    Select Case Err.number
        Case 49
            If Err.description = "DLL 呼叫規格錯誤" Then
'                WD.Quit
'                killchromedriverFromHere
                Stop
                Resume
            End If
        Case -2146233079
            If VBA.Left(Err.description, Len("session not created: Chrome failed to start: exited normally.")) = "session not created: Chrome failed to start: exited normally." Then
                WD.Quit
                SystemSetup.killchromedriverFromHere
                Set WD = Nothing
                Stop
                If MsgBox("須關閉Chrome瀏覽器再繼續！" & vbCr & vbCr & _
                    vbTab & "是否要程式自動幫您關閉、啟動。感恩感恩　南無阿彌陀佛", vbCritical + vbOKCancel) _
                        = vbOK Then
                    SystemSetup.killProcessesByName "chrome.exe"
                    GoTo reStart
                Else
                    OpenChrome = False
                End If
                
                Exit Function
            ElseIf VBA.InStr(Err.description, "session not created: This version of ChromeDriver only supports Chrome version ") = 1 Then 'session not created: This version of ChromeDriver only supports Chrome version 129
                                                                                                                                            'Current browser version is 131.0.6778.70 with binary path C:\Program Files\Google\Chrome\Application\chrome.exe (SessionNotCreated)
                MsgBox "請更新chromedriver再重試！", vbCritical
                VBA.Shell "explorer.exe " & chromePath, vbMaximizedFocus
                VBA.Shell "explorer.exe https://googlechromelabs.github.io/chrome-for-testing/#stable", vbMaximizedFocus
                killchromedriverFromHere
                Set WD = Nothing
                Exit Function
            End If
        Case -2146233088 '**'
            Debug.Print Err.number & Err.description
            If VBA.InStr(Err.description, "invalid session id") = 1 Then '-2146233088 invalid session id
                killchromedriverFromHere
                Set WD = Nothing
                GoTo reStart
            ElseIf InStr(Err.description, "Chrome failed to start: exited normally.") Then
                '' err.Descriptionunknown error: Chrome failed to start: exited normally.
                ''  (unknown error: DevToolsActivePort file doesn't exist)
                '' (The process started from chrome location W:\PortableApps\PortableApps\GoogleChromePortable\App\Chrome-bin\chrome.exe is no longer running, so ChromeDriver is assuming that Chrome has crashed.)
                If MsgBox("請關閉先前開啟的Chrome瀏覽器再繼續", vbExclamation + vbOKCancel) = vbOK Then
                        'killProcessesByName "ChromeDriver.exe", pid
                        killchromedriverFromHere
                        Set WD = Nothing
                    GoTo reStart
                Else
        '            WD.Quit
                    killchromedriverFromHere
                End If
            ElseIf InStr(Err.description, "no such window: No target with given id found") Then
                killchromedriverFromHere
                GoTo reStart
            ElseIf InStr(Err.description, "disconnected: received Inspector.detached event") Then '(failed to check if window was closed: disconnected: not connected to DevTools)
                                                                                                    '(Session info: chrome=110.0.5481.178)
                killchromedriverFromHere
                GoTo reStart
            ElseIf InStr(Err.description, "no such window: target window already closed") Then 'no such window: target window already closed
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
                '回到 wd.ExecuteScript "window.open('about:blank','_blank');" 'openNewTabWhenTabAlreadyExit WD
                     'wd.SwitchTo.Window WindowHandlesItem(WindowHandlesCount - 1)
            ElseIf InStr(Err.description, "Unexpected error. System.Net.WebException: 無法連接至遠端伺服器") Then 'Unexpected error. System.Net.WebException: 無法連接至遠端伺服器 ---> System.Net.Sockets.SocketException: 無法連線，因為目標電腦拒絕連線。 127.0.0.1:6579
                                                                                                '   於 System.Net.Sockets.Socket.DoConnect(EndPoint endPointSnapshot, SocketAddress socketAddress)
                                                                                                '   於 System.Net.ServicePoint.ConnectSocketInternal(Boolean connectFailure, Socket s4, Socket s6, Socket& socket, IPAddress& address, ConnectSocketState state, IAsyncResult asyncResult, Exception& exception)
                                                                                                '   --- 內部例外狀況堆疊追蹤的結尾 ---
                                                                                                '   於 System.Net.HttpWebRequest.GetRequestStream(TransportContext& context)
                                                                                                '   於 System.Net.HttpWebRequest.GetRequestStream()
                                                                                                '   於 OpenQA.Selenium.Remote.HttpCommandExecutor.MakeHttpRequest(HttpRequestInfo requestInfo)
                                                                                                '   於 OpenQA.Selenium.Remote.HttpCommandExecutor.Execute(Command commandToExecute)
                                                                                                '   於 OpenQA.Selenium.Remote.DriverServiceCommandExecutor.Execute(Command commandToExecute)
                                                                                                '   於 OpenQA.Selenium.Remote.RemoteWebDriver.Execute(String driverCommandToExecute, Dictionary`2 parameters)

                SystemSetup.killchromedriverFromHere
                Set SeleniumOP.WD = Nothing
'                Stop 'just for test 20240924
                Resume
            ElseIf VBA.InStr(Err.description, "chromedriver.exe does not exist") Then 'The file C:\Program Files\Google\Chrome\Application\chromedriver.exe does not exist. The driver can be downloaded at http://chromedriver.storage.googleapis.com/index.html
                Set WD = Nothing
                MsgBox "請在「" & getChromePathIncludeBackslash & "」路徑下複製chromedriver.exe檔案再繼續！", vbCritical
                OpenChrome = False
                SystemSetup.OpenExplorerAtPath getChromePathIncludeBackslash
                Exit Function
            ElseIf VBA.InStr(Err.description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                                                                    '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                                                                    '  (Session info: chrome=129.0.6668.60)
                killchromedriverFromHere
                Set WD = Nothing
                GoTo reStart
            ElseIf VBA.InStr(Err.description, "no such window") Then 'no such window
                                                                    '  (Session info: chrome=129.0.6668.59)
                If url = "https://gj.cool/punct" Then
                    OpenNewTab WD, url
                Else
                    OpenNewTab WD
                End If
                Resume
            ElseIf VBA.InStr(Err.description, "timeout: Timed out receiving message from renderer:") = 1 Then '-2146233088 timeout: Timed out receiving message from renderer: 2.972
                                                                    '(Session info: chrome=130.0.6723.69)
                If Not IsWDInvalid() Then
                    WD.Manage.Timeouts.PageLoad = timeoutsPageLoad
                    Resume
                Else
                    GoTo 2146233088
                End If
            Else
2146233088:
                Debug.Print Err.number; Err.description
                MsgBox Err.description, vbCritical
                Stop
            End If
        Case 429 'ActiveX 元件無法產生物件'
            ActiveXComponentsCanNotBeCreated = True
            Exit Function
        Case -2147467261
            If InStr(Err.description, "並未將物件參考設定為物件的執行個體。") Then
                SystemSetup.killchromedriverFromHere
                Set WD = Nothing
'                Stop
                If MsgBox("須關閉Chrome瀏覽器再繼續！" & vbCr & vbCr & _
                    vbTab & "是否要程式自動幫您關閉、啟動。感恩感恩　南無阿彌陀佛", vbCritical + vbOKCancel) _
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
                MsgBox Err.description, vbCritical
                Stop
            End If
        Case Else
            MsgBox Err.description, vbCritical
            If Err.description = "沒有設定物件變數或 With 區塊變數" Then
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
            If VBA.Dir("W:\PortableApps\PortableApps\GoogleChromePortable\App\Chrome-bin\") <> vbNullString Then
                chromePath = "W:\PortableApps\PortableApps\GoogleChromePortable\App\Chrome-bin\"
            Else
                chromePath = "W:\PortableApps\PortableApps\GoogleChromePortable64\App\Chrome-bin\"
            End If
        End If
        
        With Service
            .CreateDefaultService driverPath:=chromePath 'getChromePathIncludeBackslash
            .HideCommandPromptWindow = True '不顯示命令提示字元視窗
            If chromedriversPIDcntr = 0 Then chromedriversPIDcntr = 1
            ReDim chromedriversPID(chromedriversPIDcntr - 1)
'            chromedriversPID(chromedriversPIDcntr - 1) = Service.ProcessId'還未啟動=0
        End With
        
        Set options = New SeleniumBasic.ChromeOptions
        With options
            '.BinaryLocation = getChromePathIncludeBackslash + "chrome.exe"
            .BinaryLocation = chromePath + "chrome.exe"
            
            .AddExcludedArgument "enable-automation" '禁用「Chrome 正在被自動化軟體控制」的警告消息
            
            'C#：options.AddArgument("user-data-dir=" + Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\Google\\Chrome\\User Data\\");
            .AddArgument "user-data-dir=" + VBA.Environ("LOCALAPPDATA") + _
                "\Google\Chrome\User Data\"
            .AddArgument "--headless" '不顯示實體，即看不到Chrome瀏覽器，無法手動操作及監控
            .AddArgument "--disable-gpu"
            .AddArgument "--disable-infobars"
            .AddArgument "--disable-extensions"
            .AddArgument "--disable-dev-shm-usage"
            '.AddArgument "--start-maximized"
            '.DebuggerAddress = "127.0.0.1:9999" '不要与其他几個混用
'            .AddArgument "--remote-debugging-port=9222"
        End With
        WD.New_ChromeDriver Service:=Service, options:=options
        'WD.Quit 會自動清除chromedriver，就不用記下開過哪些了
'        pid = Service.ProcessId 'Chrome瀏覽器沒有開成功就會是0
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
Select Case Err.number

    Case -2146233088 '**'
        'Debug.Print Err.Description
        '' err.Descriptionunknown error: Chrome failed to start: exited normally.
        ''  (unknown error: DevToolsActivePort file doesn't exist)
        '' (The process started from chrome location W:\PortableApps\PortableApps\GoogleChromePortable\App\Chrome-bin\chrome.exe is no longer running, so ChromeDriver is assuming that Chrome has crashed.)
        If chromedriversPIDcntr = 0 Then chromedriversPIDcntr = 1
        ReDim chromedriversPID(chromedriversPIDcntr - 1)
        chromedriversPID(chromedriversPIDcntr - 1) = Service.ProcessId
        If InStr(Err.description, "/session timed out after 60 seconds.") Then
            killchromedriverFromHere
            Set openChromeBackground = Nothing
        Else
            If MsgBox("請關閉先前開啟的Chrome瀏覽器再繼續", vbExclamation + vbOKCancel) = vbOK Then
                'killProcessesByName "ChromeDriver.exe", pid
                killchromedriverFromHere
            GoTo reStart
            End If
        End If
    Case Else
        MsgBox Err.description, vbCritical
'        Resume
End Select

'20230119 creedit chatGPT大菩薩

'    Dim driver As New Selenium.WebDriver
'    'driver.start "chrome", "https://www.google.com"
'    driver.SetBinary getChromePathIncludeBackslash
'    driver.start getChromePathIncludeBackslash + "chrome.exe", "https://www.google.com"
'    driver.Get "/"
End Function

Rem 20241002 Copilot大菩薩：Word VBA 中的 Selenium 操作：
Rem 在 VBA 中連接到 ChromeDriver：https://sl.bing.net/ib6ZEOurJ4S
Rem 使用 SeleniumBasic 在 VBA 中連接到已啟動的 ChromeDriver。例如
Sub SeleniumGet()
'    Dim driver As New WebDriver
'    Dim options As New ChromeOptions
'
'    options.AddArgument "--remote-debugging-port=9222"
'    driver.start "chrome", options
'
'    driver.Get "http://localhost:9222"
'    ' 進行進一步的操作
End Sub
Sub SeleniumGetTest()
    Dim driver As New IWebDriver
    Dim options As New ChromeOptions
    Dim Service As New SeleniumBasic.ChromeDriverService
    Dim chromePath As String
    chromePath = getChromePathIncludeBackslash
    If InStr(chromePath, "GoogleChromePortable") Then
        If VBA.Dir("W:\PortableApps\PortableApps\GoogleChromePortable\App\Chrome-bin\") <> vbNullString Then
            chromePath = "W:\PortableApps\PortableApps\GoogleChromePortable\App\Chrome-bin\"
        Else
            chromePath = "W:\PortableApps\PortableApps\GoogleChromePortable64\App\Chrome-bin\"
        End If
    End If

    With Service
        .CreateDefaultService driverPath:=chromePath 'getChromePathIncludeBackslash
        '.CreateDefaultService driverPath:="E:\Selenium\Drivers"
        .HideCommandPromptWindow = True '不顯示命令提示字元視窗
    End With

    options.AddArgument "--remote-debugging-port=9222"
    On Error Resume Next
    driver.New_ChromeDriver Service, options

'    driver.Get "http://localhost:9222"
'    driver.Navigate.GoToUrl "http://localhost:9222"
    If Err.number = 0 Then
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
'    ' 進行進一步的操作
End Sub
Rem 20241002 由前面 SeleniumGet 得到的靈感 creedit_with_Copilot大菩薩：https://sl.bing.net/hwtm2YPAfdY
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
        If VBA.Dir("W:\PortableApps\PortableApps\GoogleChromePortable\App\Chrome-bin\") <> vbNullString Then
            chromePath = "W:\PortableApps\PortableApps\GoogleChromePortable\App\Chrome-bin\"
        Else
            chromePath = "W:\PortableApps\PortableApps\GoogleChromePortable64\App\Chrome-bin\"
        End If
    End If

    With Service
        .CreateDefaultService driverPath:=chromePath 'getChromePathIncludeBackslash
        '.CreateDefaultService driverPath:="E:\Selenium\Drivers"
        .HideCommandPromptWindow = True '不顯示命令提示字元視窗
    End With
    With options
'        If Not IsChromeRunning Then '若已用同一使用者設定檔開啟則無法再開新的Chrome瀏覽器了
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
    Docs.Register_Event_Handler '為清除chromedriver作準備
    pid = Service.ProcessId 'Chrome瀏覽器沒有開成功就會是0
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
            
            Rem 怕會誤關其他先開啟的Chrome瀏覽器
            'ActivateChrome
'            SystemSetup.wait 2
'            VBA.Interaction.DoEvents
            
            Debug.Print "Word is active = " & VBA.CStr(IsWordActive())
                        
            If VBA.InStr(Err.description, "from disconnected: unable to connect to renderer (SessionNotCreated)") = 0 Then
                If IsWordActive() Then
                    MsgBox "請關閉Chrome瀏覽器後再繼續。", vbExclamation
                    'Stop 'just for test
                    
    '                SendKeys "%{F4}", True '關閉已開啟而無法成功的Chrome瀏覽器
    '                SystemSetup.playSound 1.469
    '                VBA.Interaction.DoEvents
                Else
    '                Stop 'just for test
'                    ActivateChrome
'                    SendKeys "^{F4}", True '關閉已開啟的Chrome瀏覽器分頁
'                    SystemSetup.playSound 1.469
'                    VBA.Interaction.DoEvents
                    
                End If
            End If
            Set options = New SeleniumBasic.ChromeOptions
            With options
                .AddArgument "--remote-debugging-port=9222"
                '要能夠順利 Get 到手動啟動的Chrome瀏覽器，則在手動啟動Chrome瀏覽器的捷徑「目標(T)」欄位內的值後要加上 「--remote-debugging-port=9222」，如： "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 '20241002
                .AddArgument "--start-maximized"
            End With
'            If Not WD Is Nothing Then
'                'WD.Close
'                WD.Quit
'                Set WD = Nothing
'                killchromedriverFromHere
'            End If
            driver.New_ChromeDriver Service:=Service, options:=options
            Docs.Register_Event_Handler '為清除chromedriver作準備
            pid = Service.ProcessId 'Chrome瀏覽器沒有開成功就會是0
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
'                MsgBox "請再執行一次。感恩感恩　南無阿彌陀佛", vbInformation
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
'            SendKeys "%{F4}" '關掉成功啟動後獨立的分頁
'            VBA.Interaction.DoEvents
            If UBound(driver.WindowHandles) > 0 Then
                Dim wh
                For Each wh In WD.WindowHandles
                    WD.SwitchTo.Window wh
                    If WD.title = "新分頁" Then 'WD.url="chrome://new-tab-page/"
                        WD.Close '關掉成功啟動後獨立的分頁
                        SystemSetup.playSound 1.469
                        VBA.Interaction.DoEvents
                        Exit For
                    End If
                Next wh
                If IsWDInvalid() Then
                    WD.SwitchTo.Window UBound(WD.WindowHandles)
                End If

                openNewTabWhenTabAlreadyExit WD
'                OpenNewTab WD '再開啟一個新分頁，供後續程式操作用，避免影響原已開啟的分頁
            End If
        End If
    End If
    On Error GoTo 0
    Rem 20241008 Gemini大菩薩：關閉錯誤處理： 當程式執行到 On Error GoTo 0 這行時，之前設定的任何錯誤處理都會被關閉。
    Rem 恢復預設行為： 關閉錯誤處理後，如果程式再次遇到錯誤，就會按照預設的行為，直接停止執行並顯示錯誤訊息。https://g.co/gemini/share/7359ab0a85e3
    
    Docs.Register_Event_Handler '為清除chromedriver作準備
    'driver.Navigate.GoToUrl "https://github.com/oscarsun72/TextForCtext/blob/master/WordVBA/SeleniumOP.bas"
    'driver.Get "http://localhost:9222"
'    Dim Wd As New SeleniumBasic.IWebDriver
'    'wd.Get "http://localhost:9222"
'    Wd.Navigate "http://localhost:9222" 'https://github.com/GCuser99/SeleniumVBA/discussions/74
    ' 進行進一步的操作
    
    If closeNewBlankPageTabs() Then OpenNewTab WD
    OpenChrome_NEW_Get = True
    Exit Function
eH:
    Select Case Err.number
        Case -2146233088
            If VBA.InStr(Err.description, "chromedriver.exe does not exist") Then 'The file C:\Program Files\Google\Chrome\Application\chromedriver.exe does not exist. The driver can be downloaded at http://chromedriver.storage.googleapis.com/index.html
                Set WD = Nothing
                MsgBox "請在「" & getChromePathIncludeBackslash & "」路徑下複製chromedriver.exe檔案再繼續！", vbCritical
                SystemSetup.OpenExplorerAtPath getChromePathIncludeBackslash
                Exit Function
            Else
                GoTo caseElse
            End If
caseElse:
        Case Else
            Debug.Print Err.number & vbTab & Err.description
            MsgBox Err.number & Err.description
            'Resume
    End Select
End Function
Sub CloseNewBlankPagesTabs()
    closeNewBlankPageTabs
End Sub
Rem 若沒有新的空白頁要關閉則傳回false,若只剩一個分頁則不予關閉且傳回false供後續使用
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


Rem 開啟新分頁 若失敗則傳回false'改進WordVBA+SeleniumBasic 開啟Chrome瀏覽器新分頁的方法 creedit_with_Copilot大菩薩： https://sl.bing.net/bcfc14PWlFc
Function OpenNewTab(ByVal driver As SeleniumBasic.IWebDriver, Optional url As String) As Boolean
    Dim result As Boolean, currentWinhdl As String
    
    result = True
    On Error GoTo eH
    
'    SeleniumOP.ActivateChrome
'    SendKeys "^t"
'
    If Not IsWDInvalid Then currentWinhdl = driver.CurrentWindowHandle
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
        If Not IsNewBlankPageTab(driver) Then '若沒成功開啟
            '建立 Actions 物件 。 Copilot大菩薩：在這段改進的程式碼中，我使用 CreateObject 方法來建立 Actions 物件，並且直接呼叫 Perform 方法來執行動作。這樣可以確保 Actions 物件正確建立並執行。
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
            If Not IsNewBlankPageTab(driver) Then '若沒成功開啟
                driver.SwitchTo().Window driver.CurrentWindowHandle
                VBA.Interaction.DoEvents
                
                ActivateChrome
                VBA.Interaction.DoEvents
                
                SystemSetup.playSound 0.411
                
                SystemSetup.wait 1.7 '有這行就可以了
                
                VBA.Interaction.SendKeys "^t" ', True
                VBA.Interaction.DoEvents
                
                SystemSetup.wait 1.9 '有這行就可以了
                
                SwitchToLastWindowHandleWindow driver
                VBA.Interaction.DoEvents
                SystemSetup.playSound 1 'for test
                If Not IsNewBlankPageTab(driver) Then '若沒成功開啟
                    For Each wh In driver.WindowHandles
                        driver.SwitchTo.Window wh
                        If IsNewBlankPageTab(driver) Then Exit For
                    Next wh
                    If Not IsNewBlankPageTab(driver) Then '若沒成功開啟
                        If currentWinhdl <> vbNullString Then
                            driver.SwitchTo.Window driver.CurrentWindowHandle
                            SystemSetup.playSound 1
                            Exit Function
                        ElseIf url <> vbNullString Then
                            For wh = UBound(driver.WindowHandles) To 0 Step -1
                                driver.SwitchTo.Window driver.WindowHandles()(wh)
                                If driver.url = url Then
                                    SystemSetup.playSound 0.96
                                    Exit Function
                                End If
                            Next wh
                        End If
                        'Stop 'for debug
                        ActivateChrome
                        word.Application.Activate
                        If VBA.vbOK = MsgBox("若要開啟新分頁視窗請手動開啟後再按下「確定」按鈕，否則即在此分頁繼續執行。" & vbCr & vbCr _
                            & "若不想在此分頁執行，請務必自行開啟新分頁或新視窗，再按下「確定」按鈕。感恩感恩　南無阿彌陀佛", VBA.vbOKCancel + VBA.vbExclamation) Then
                            SwitchToLastWindowHandleWindow driver
                        Else
                            result = False
                            driver.SwitchTo.Window driver.CurrentWindowHandle
                        End If
                        result = False
                        driver.SwitchTo.Window driver.CurrentWindowHandle
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
    Select Case Err.number
        Case -2146233088
            If VBA.InStr(Err.description, "no such window: target window already closed") = 1 Then 'no such window: target window already closed
                driver.SwitchTo.Window driver.WindowHandles()(UBound(driver.WindowHandles))
                Resume
            Else
                GoTo caseElse
            End If
caseElse:
        Case Else
            Debug.Print Err.number & Err.description
            MsgBox Err.number & Err.description
'            Resume
    End Select
'    On Error GoTo eH
'    driver.ExecuteScript "window.open('about:blank','_blank');" 'openNewTabWhenTabAlreadyExit WD
'    If Not IsNewBlankPageTab(driver) Then
'        Dim key As New SeleniumBasic.keys, iwe As SeleniumBasic.IWebElement
''        driver.FindElementByTagName("body").SendKeys "^t" ' Ctrl + t to open a new tab '20241003creedit_with_Copilot大菩薩：解決WordVBA + SeleniumBasic開新分頁問題：https://sl.bing.net/gehCkm98JRA
'        Set iwe = driver.FindElementByTagName("body")
'        If iwe Is Nothing Then Stop 'just for test
'        iwe.SendKeys key.Control + "t"
'
'        If Not IsNewBlankPageTab(driver) Then '若沒成功開啟
'            '建立 Actions 物件
'            Dim actions As New SeleniumBasic.actions
''            actions.MoveToElement(iwe).Click().Perform
''            actions.SendKeys(key.Control + "t").Build '().Perform
'
'            If Not IsNewBlankPageTab(driver) Then '若沒成功開啟
'                word.Application.windowState = wdWindowStateMinimize
'                driver.SwitchTo().Window driver.CurrentWindowHandle
'                VBA.Interaction.DoEvents
'                SendKeys "^t", True
'                VBA.Interaction.DoEvents
'                If Not IsNewBlankPageTab(driver) Then '若沒成功開啟
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

Rem 20241002 Copilot大菩薩：Word VBA 中的 Selenium 操作: https://sl.bing.net/jH2j6GzDiQm
Rem 使用 Word VBA 取得遠程調試端口
Rem 檢查 Chrome 瀏覽器是否已啟動： 使用 WMI (Windows Management Instrumentation) 來檢查 Chrome 瀏覽器是否正在運行。
Rem 讀取遠程調試端口： 假設你已經將遠程調試端口寫入文件，可以從該文件中讀取端口號。
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
'        wd.url = url'前1行已有
        word.Application.windowState = wdWindowStateMinimize
        WD.SwitchTo.Window (WD.CurrentWindowHandle)
        VBA.Interaction.DoEvents
        SeleniumOP.ActivateChrome
        
        Dim form As SeleniumBasic.IWebElement
        Dim keyword As SeleniumBasic.IWebElement
        Dim button As SeleniumBasic.IWebElement
        Set form = WD.FindElementById(frmID)
        Set keyword = form.FindElementById(keywdID)
        Set button = form.FindElementById(btnID)
        If searchStr <> "" Then
            'keyword.SendKeys searchStr
            SetIWebElementValueProperty keyword, searchStr
            '上一行輸入即檢索了，故可不必下一行;但若不想顯示下拉清單，且確定可顯示結果，則還是需要下一行
            button.Click
        End If
    '    Debug.Print WD.title, WD.url
    '    Debug.Print WD.PageSource
    '    MsgBox "下面退出瀏覽器。"
    '    WD.Quit
        Exit Sub
Err1:
        Select Case Err.number
            Case 49 'DLL 呼叫規格錯誤
                Resume
            Case Else
                MsgBox Err.description, vbCritical
                SystemSetup.killchromedriverFromHere
    '           Resume
    End Select
End Sub

'找百度 ： https://www.cnblogs.com/ryueifu-VBA/p/13661128.html
Sub BaiduSearch(Optional searchStr As String)
    On Error GoTo Err1
    Search "https://www.baidu.com", "form", "kw", "su", searchStr
'    wd.SwitchTo.Window (wd.CurrentWindowHandle)'Search 裡已有 20240930
'    VBA.Interaction.DoEvents
        Exit Sub
Err1:
        Select Case Err.number
            Case 49 'DLL 呼叫規格錯誤
                Resume
            Case Else
                MsgBox Err.description, vbCritical
                SystemSetup.killchromedriverFromHere
    '           Resume
    End Select
End Sub

'查詢國語辭典
Sub dictRevisedSearch(Optional searchStr As String)
    On Error GoTo Err1
    'If searchStr = "" And Selection = "" Then Exit Sub
    Const url As String = "https://dict.revised.moe.edu.tw/search.jsp?md=1"
    If WD Is Nothing Or IsDriverInvalid(WD) Then
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
                keyword.Submit '這兩個方法都可
    '            Dim k As New SeleniumBasic.keys
    '            keyword.SendKeys k.Enter
            End If
        End If
    '   退出瀏覽器。"
    '    WD.Quit
        Exit Sub
Err1:
        Select Case Err.number
            Case 49 'DLL 呼叫規格錯誤
                Resume
            Case Else
                MsgBox Err.description, vbCritical
                SystemSetup.killchromedriverFromHere
    '           Resume
        End Select
End Sub

'擷取國語辭典詞條網址
Function grabDictRevisedUrl_OnlyOneResult(searchStr As String, Optional Background As Boolean) As String
    'If searchStr = "" And Selection = "" Then Exit Sub
    If searchStr = "" Then Exit Function
    If VBA.Left(searchStr, 1) <> "=" Then searchStr = "=" + searchStr '精確搜尋字串指令
    Const notFoundOrMultiKey As String = "&qMd=0&qCol=1" '查無資料或如果不止一條時，網址後綴都有此關鍵字
    Dim url As String, retryTime As Byte
    url = "https://dict.revised.moe.edu.tw/search.jsp?md=1"
    
    On Error GoTo Err1
    
    Dim wdB As SeleniumBasic.IWebDriver, WBQuit As Boolean '=true 則可以關Chrome瀏覽器
    
    If Background Then
        WBQuit = True '因為在背景執行，預設要可以關
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
                If Not IsWDInvalid Then
                    Set wdB = WD
                Else
                    OpenChrome url
                End If
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
'        If keyword.text <> searchStr Then 20240914作廢
'            keyword.Clear
'            keyword.SendKeys searchStr
'        End If
        
        Rem 在 headless 參數設定下開啟的Chrome瀏覽器，是無法使用系統貼上功能的
        Rem Dim key As New SeleniumBasic.keys
        Rem     keyword.SendKeys key.Control + "v"
        Rem     keyword.SendKeys key.LeftShift + key.Insert
        Rem 改用Chrome瀏覽器介面功能表的貼上功能試試 20230121 也不行：
        Rem <stale element reference: element is not attached to the page document(Session info: headless chrome=109.0.5414.75)>
        Rem 因為只能操控網頁，不是瀏覽器介面
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
            keyword.Submit '這兩個方法都可
    '            Dim k As New SeleniumBasic.keys
    '            keyword.SendKeys k.Enter
        End If
        url = wdB.url
        If InStr(url, notFoundOrMultiKey) = 0 Then
            grabDictRevisedUrl_OnlyOneResult = url '有找到則傳回網址
        Else
            grabDictRevisedUrl_OnlyOneResult = "" '沒有找到傳回空字串
        End If
        If WBQuit Then
            '退出瀏覽器
            wdB.Quit
            If Not Background Then Set WD = Nothing
        Else
            wdB.Close
        End If
        Exit Function
Err1:
        Select Case Err.number
            Case 49 'DLL 呼叫規格錯誤
                Resume
            Case 91 '沒有設定物件變數或 With 區塊變數
                If retryTime > 1 Then
                    MsgBox Err.number + Err.description
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
            Case -2147467261 '並未將物件參考設定為物件的執行個體。
                Set WD = Nothing
                killchromedriverFromHere
                OpenChrome url
                Set wdB = WD
                WBQuit = True
                Resume
            Case -2146233088 'unknown error: ChromeDriver only supports characters in the BMP  (Session info: chrome=109.0.5414.75)
                If InStr(Err.description, "/session timed out after 60 seconds.") Then
                    If WD Is Nothing Then OpenChrome (url)
                    Set wdB = WD
                ElseIf InStr(Err.description, "no such window: target window already closed") Or InStr(Err.description, "invalid session id") Then
                    WD.Quit: Set WD = Nothing: killchromedriverFromHere: OpenChrome (url)
                    Set wdB = WD
                Else
                    'textbox.SendKeys key.LeftShift + key.Insert
                    WBQuit = pasteWhenOutBMP(wdB, url, "word", searchStr, keyword, Background)
                End If
                Resume Next
            Case Else
                MsgBox Err.description, vbCritical
                wdB.Quit
                SystemSetup.killchromedriverFromHere
    '           Resume
        End Select

End Function
Rem 20241006 《看典古籍·古籍全文檢索》，成功則傳回true
Function KandiangujiSearchAll(searchTxt As String) As Boolean
    Dim exact As Boolean, key As New SeleniumBasic.keys
    Const url = "https://kandianguji.com/search"
    SystemSetup.SetClipboard searchTxt
    If VBA.vbOK = VBA.MsgBox("是否要【精確檢索】？", vbQuestion + vbOKCancel) Then exact = True

    If Not IsWDInvalid() Then
        LastValidWindow = WD.CurrentWindowHandle
        If WD.url <> url Then
            If Not OpenChrome(url) Then Exit Function
        End If
    Else
        If Not OpenChrome(url) Then
            Exit Function
        End If
    End If
    
    WD.SwitchTo().Window (WD.CurrentWindowHandle)
    ActivateChrome
    word.Application.windowState = wdWindowStateMinimize
    
    Dim iwe As SeleniumBasic.IWebElement ', key As New SeleniumBasic.keys
    
    Dim dt As Date
    dt = VBA.Now

    '   繁簡同檢
    Set iwe = WD.FindElementByCssSelector("#search_select")
    Do While iwe Is Nothing
        Set iwe = WD.FindElementByCssSelector("#search_select")
        If VBA.DateDiff("s", dt, VBA.Now) > 3 And iwe Is Nothing Then
            Exit Function
        End If
    Loop
    iwe.Click
    'SetIWebElementValueProperty iwe, "jianfan"
    Set iwe = WD.FindElementByCssSelector("#search_select > option:nth-child(2)")
    iwe.Click
    
    ' 設定匹配模式
    Set iwe = WD.FindElementByCssSelector("#search_mode")
    iwe.Click
    If exact Then
        Set iwe = WD.FindElementByCssSelector("#search_mode > option:nth-child(2)")
    Else
        Set iwe = WD.FindElementByCssSelector("#search_mode > option:nth-child(1)")
    End If
    If iwe Is Nothing Then Exit Function
    iwe.Click
    
    Set iwe = WD.FindElementByCssSelector("#search_input")
    If iwe Is Nothing Then Exit Function
    
    iwe.Clear
    'SetIWebElementValueProperty iwe, searchTxt'須用Sendkeys才行
    iwe.SendKeys key.Shift + key.Insert
    
    'iwe.SendKeys key.enter '按下Enter鍵並無作用
    ' 「檢索」按鈕。好奇怪的檢索按鈕與機制！
    Set iwe = WD.FindElementByCssSelector("#search_button")
    If iwe Is Nothing Then
        Exit Function
    End If
    iwe.Click
    
    KandiangujiSearchAll = True
End Function
Rem 20241006 檢索《漢籍全文資料庫》，成功則傳回true
Function HanchiSearch(searchTxt As String) As Boolean
    Dim free As Boolean, inside As Boolean
    SystemSetup.SetClipboard searchTxt
    If Not IsWDInvalid() Then
        If VBA.Left(WD.url, VBA.Len("https://hanchi.ihp.sinica.edu.tw/")) <> "https://hanchi.ihp.sinica.edu.tw/" Then
            If VBA.vbCancel = MsgBox("是否是【授權使用】？", vbQuestion + vbOKCancel) Then free = True
        Else
            inside = True
        End If
    Else
        If VBA.vbCancel = MsgBox("是否是【授權使用】？", vbQuestion + vbOKCancel) Then free = True
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
    word.Application.windowState = wdWindowStateMinimize
    Dim iwe As SeleniumBasic.IWebElement, key As New SeleniumBasic.keys
    If Not inside Then
        If free Then
            Set iwe = WD.FindElementByCssSelector("body > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td > table > tbody > tr:nth-child(4) > td > a:nth-child(8) > img")
            If iwe Is Nothing Then
                Set iwe = WD.FindElementByXPath("/html/body/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[4]/td/a[1]/img")
            End If
        Else
            Set iwe = WD.FindElementByCssSelector("body > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td > table > tbody > tr:nth-child(4) > td > a:nth-child(9) > img")
            If iwe Is Nothing Then
                Set iwe = WD.FindElementByXPath("/html/body/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[4]/td/a[2]/img")
            End If
        End If
        
        If iwe Is Nothing Then Exit Function
        iwe.Click
    End If
    'keyword
    Set iwe = WD.FindElementByCssSelector("#frmTitle > table > tbody > tr:nth-child(2) > td > table > tbody > tr:nth-child(1) > td > input[type=text]:nth-child(2)")
    Dim dt As Date
    dt = VBA.Now
    Do While iwe Is Nothing
        iwe = WD.FindElementByCssSelector("#frmTitle > table > tbody > tr:nth-child(2) > td > table > tbody > tr:nth-child(1) > td > input[type=text]:nth-child(2)")
        If VBA.DateDiff("s", dt, VBA.Now) > 2 Then Exit Do
    Loop
    If iwe Is Nothing Then Exit Function
    
    SetIWebElementValueProperty iwe, searchTxt
    iwe.SendKeys key.enter
    HanchiSearch = True
    
End Function
Rem x 要查的字,Variants 要不要看異體字 執行成功傳回true  20240828.
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
    Rem 若須直接查看異體字
    If Variants Then
        Dim dt As Date
        dt = VBA.Now
        Do While iwe Is Nothing
            Set iwe = WD.FindElementByCssSelector("#mainContent > span > div.content > div > div.sidebar_navigation > div > div:nth-child(11)")
            If DateDiff("s", dt, VBA.Now) > 3 Then
                Exit Do '找不到相關字的元件
            End If
        Loop
        If Not iwe Is Nothing Then iwe.Click
    End If
    LookupZitools = True
    Exit Function
eH:
Select Case Err.number
        Case -2146233088
            If InStr(Err.description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                            '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                            '  (Session info: chrome=128.0.6613.85)
                'Set wd = Nothing
                SystemSetup.killchromedriverFromHere
                Set WD = Nothing
                Resume
            Else
                MsgBox Err.number & Err.description, vbExclamation
            End If
        Case Else
            MsgBox "請關閉Chrome瀏覽器後再執行一次！" & vbCr & vbCr & Err.number & Err.description, vbExclamation
    End Select
End Function
Rem 查《古音小鏡·訓詁工具書查詢》,成功則傳回true 20241020
Function LookupBook_Xungu_kaom(x As String) As Boolean
    If Not code.IsChineseString(x) Then
        MsgBox "只限中文！", vbCritical
        Exit Function
    End If
    Dim iwe As SeleniumBasic.IWebElement, dt As Date, key As New SeleniumBasic.keys
    If Not OpenChrome("http://www.kaom.net/book_xungu.php") Then Exit Function
    word.Application.windowState = wdWindowStateMinimize
    WD.SwitchTo.Window WD.CurrentWindowHandle
    ActivateChrome
    dt = VBA.Now
    '檢索框
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
'    '查詢按鈕
    SystemSetup.wait 3.3
    Set iwe = WD.FindElementByCssSelector("body > table > tbody > tr > td > form > input.form_2")
    If iwe Is Nothing Then Exit Function
    iwe.Click
    LookupBook_Xungu_kaom = True
End Function
Rem 查《古音小鏡》漢語大詞典,成功則傳回true 20241020
Function LookupHYDCD_kaom(x As String) As Boolean
    If Not code.IsChineseString(x) Then
        MsgBox "只限中文！", vbCritical
        Exit Function
    End If
    Dim iwe As SeleniumBasic.IWebElement, dt As Date ', key As New SeleniumBasic.keys
    If Not OpenChrome("http://www.kaom.net/book_hanyudacidian.php") Then Exit Function
    word.Application.windowState = wdWindowStateMinimize
    WD.SwitchTo.Window WD.CurrentWindowHandle
    ActivateChrome
    '檢索框
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
    '查詢按鈕
    VBA.Interaction.DoEvents
    SystemSetup.wait 3.3
    Set iwe = WD.FindElementByCssSelector("body > table > tbody > tr > td > form > input.form_2")
    If iwe Is Nothing Then Exit Function
    iwe.Click
    LookupHYDCD_kaom = True
End Function
Rem 查《異體字字典》：x 要查的字。傳回一個字串陣列，第1個元素是所查詢的字串，第2個元素是查詢結果網址。若沒找到，則傳回空字串 ""
Function LookupDictionary_of_ChineseCharacterVariants(x As String) As String()
    On Error GoTo eH
    Dim result(1) As String '1=索引值上限（最大值）
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
    '檢索輸入框
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
        iwe.SendKeys keys.enter
        '查詢結果訊息框，如[ 孫 ]， 查詢結果：正文 1 字，附收字 3 字
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
            '如果直接顯示該字頁面，非查詢結果頁，如： https://dict.variants.moe.edu.tw/dictView.jsp?ID=5565
            '字頭元件
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
Select Case Err.number
        Case -2146233088
            If InStr(Err.description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                            '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                            '  (Session info: chrome=128.0.6613.85)
                'Set wd = Nothing
                SystemSetup.killchromedriverFromHere
                Set WD = Nothing
                Resume
            Else
                MsgBox Err.number & Err.description, vbExclamation
            End If
        Case Else
            MsgBox "請關閉Chrome瀏覽器後再執行一次！" & vbCr & vbCr & Err.number & Err.description, vbExclamation
    End Select
End Function
Rem 查《國語辭典》：x 要查的字詞。傳回一個字串陣列，第1個元素是所查詢的字串，第2個元素是查詢結果網址。若沒找到，則傳回空字串 ""
Function LookupDictRevised(x As String) As String()
    On Error GoTo eH
    Dim result(1) As String '1=索引值上限（最大值）
    LookupDictRevised = result

    If Not code.IsChineseString(x) Then
        MsgBox "只能檢索中文。請檢查檢索字串，重新開始。", vbExclamation
        Exit Function
    End If
    SystemSetup.SetClipboard x
    
    If IsDriverInvalid(WD) Then
        If Not OpenChrome("https://dict.revised.moe.edu.tw/search.jsp?md=1") Then Exit Function
    Else
        WD.SwitchTo.Window WD.CurrentWindowHandle
'        If VBA.InStr(WD.url, "https://dict.revised.moe.edu.tw/") <> 1 Then WD.url = "https://dict.revised.moe.edu.tw/search.jsp?md=1"
        WD.Navigate.GoToUrl "https://dict.revised.moe.edu.tw/search.jsp?md=1"
    End If
    
    
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '檢索輸入框
    Do While iwe Is Nothing
        Set iwe = WD.FindElementByCssSelector("#searchF > div.line > input[type=text]:nth-child(1)")
        If DateDiff("s", dt, VBA.Now) > 3 Then
            Exit Function
        End If
    Loop
    
'    VBA.AppActivate "chrome"
    ActivateChrome
    word.Application.windowState = wdWindowStateMinimize
    WD.SwitchTo.Window (WD.CurrentWindowHandle)
    VBA.Interaction.DoEvents

    '找到檢索框之後
    If Not iwe Is Nothing Then
        Dim keys As New SeleniumBasic.keys
        'iwe.SendKeys keys.Shift + keys.Insert
        iwe.SendKeys keys.Control + "v"
        iwe.SendKeys keys.enter
        '查詢結果訊息框，如 查無資料
        Set iwe = WD.FindElementByCssSelector("#searchL > tbody > tr > td")
        '查詢有結果時：
        If iwe Is Nothing Then
            result(0) = x
            result(1) = WD.url
            SystemSetup.SetClipboard result(1)
        End If
    End If
    LookupDictRevised = result
    Exit Function
eH:
Select Case Err.number
        Case -2146233088
            If InStr(Err.description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                            '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                            '  (Session info: chrome=128.0.6613.85)
                'Set wd = Nothing
                SystemSetup.killchromedriverFromHere
                Set WD = Nothing
                Resume
            Else
                MsgBox Err.number & Err.description, vbExclamation
            End If
        Case Else
            MsgBox "請關閉Chrome瀏覽器後再執行一次！" & vbCr & vbCr & Err.number & Err.description, vbExclamation
    End Select
End Function
Rem 查《漢語大詞典》：x 要查的字詞。傳回一個字串陣列，第1個元素是所查詢的字串，第2個元素是查詢結果網址。若沒找到，則傳回空字串 ""
Function LookupHYDCD(x As String) As String()
    On Error GoTo eH
    Dim result(1) As String '1=索引值上限（最大值）
    LookupHYDCD = result
    If Not code.IsChineseString(x) Then
        MsgBox "只能檢索中文。請檢查檢索字串，重新開始。", vbCritical
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
    '檢索輸入框
    Do While iwe Is Nothing
        Set iwe = WD.FindElementByCssSelector("#SearchBox")
        If DateDiff("s", dt, VBA.Now) > 3 Then
            Exit Function
        End If
    Loop
    

    
    '找到檢索框之後
    If Not iwe Is Nothing Then
        Dim keys As New SeleniumBasic.keys
        iwe.SendKeys keys.Shift + keys.Insert
        'iwe.SendKeys keys.Control + "v"
        iwe.SendKeys keys.enter
        '查詢結果訊息框，如 抱歉，無此詞語。
                        '本掃描版詞典無法查詢簡體字，也無法定位到單字。
                        '若要查單字，可查詢以該字開頭的詞語，再按「上一頁」直到該單字出現，
                        '或使用下面支持單字查詢的《漢語大詞典》連結
                        '或使用《漢語大字典》。
        Set iwe = WD.FindElementByCssSelector("#SearchResult > font")
        '查詢有結果時：
        If iwe Is Nothing Then
            '查詢結果的超連結框
            Set iwe = WD.FindElementByCssSelector("#SearchResult > p > a > font")
            If Not iwe Is Nothing Then
                iwe.Click
                result(0) = x
                WD.SwitchTo.Window WindowHandlesItem(WindowHandlesCount - 1)
                result(1) = WD.url
                SystemSetup.SetClipboard result(1)
            Else
                MsgBox "請檢查", vbCritical
                Stop
            End If
        End If
    End If
    LookupHYDCD = result
    Exit Function
eH:
Select Case Err.number
        Case -2146233088
            If InStr(Err.description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                            '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                            '  (Session info: chrome=128.0.6613.85)
                'Set wd = Nothing
                SystemSetup.killchromedriverFromHere
                Set WD = Nothing
                Resume
            Else
                MsgBox Err.number & Err.description, vbExclamation
            End If
        Case Else
            MsgBox "請關閉Chrome瀏覽器後再執行一次！" & vbCr & vbCr & Err.number & Err.description, vbExclamation
    End Select
End Function
Rem 查《漢語大字典》傳入檢索字，失敗傳回fasle，若成功，則複製其網址。
Function LookupHYDZD(x As String)
    Dim a As New StringInfo, e As SeleniumBasic.IWebElement, key As New SeleniumBasic.keys
    a.Create x
    If a.LengthInTextElements > 1 Then
        word.Application.Activate
        MsgBox "僅限查一字！", vbCritical
        Exit Function
        'LookupZWDZD = False
    End If
    If IsDriverInvalid(WD) Then
        OpenChrome "https://homeinmists.ilotus.org/hd/hydzd.php"
    Else
        SeleniumOP.ActivateChrome
        WD.SwitchTo.Window WD.CurrentWindowHandle
        WD.Navigate.GoToUrl "https://homeinmists.ilotus.org/hd/hydzd.php"
    End If
    'Set e = WD.FindElementByXPath("/html/body/div[2]/div[2]/form/div[3]/div[1]/input")
    Set e = WD.FindElementByCssSelector("#keyword")
    SeleniumOP.SetIWebElementValueProperty e, x
    'e.SendKeys key.enter
    Set e = WD.FindElementByCssSelector("#main_form > div.form-row > div.col-3 > button")
    e.Click
    
    
    word.Application.ActiveWindow.windowState = wdWindowStateMinimize
    
'    SystemSetup.wait 0.8
    '檢索結果第2個頁碼的連結儲存格
'    Set e = WD.FindElementByClassName("#searchTableOut > tbody > tr:nth-child(3) > td:nth-child(3) > a")
    Set e = WD.FindElementByXPath("/html/body/div[2]/div[4]/div/table/tbody/tr[3]/td[3]/a")
    
    '若檢索結果不止一個
    If Not e Is Nothing Then Exit Function
    
    'Set e = WD.FindElementByClassName("#searchTableOut > tbody > tr:nth-child(2) > td:nth-child(3) > a")
    Set e = WD.FindElementByXPath("/html/body/div[2]/div[4]/div/table/tbody/tr[2]/td[3]/a")
    
    If e Is Nothing Then
            'word.Application.Activate
            'MsgBox "有錯誤，請檢索，或手動執行！", vbCritical
            'Exit Function
            
            GoSub enter
            
    End If
    Dim dt As Date
    dt = VBA.Now
    Do While e Is Nothing
        GoSub enter
'        Set e = WD.FindElementByClassName("#searchTableOut > tbody > tr:nth-child(2) > td:nth-child(3) > a")
        Set e = WD.FindElementByXPath("/html/body/div[2]/div[4]/div/table/tbody/tr[2]/td[3]/a")

        If Not WD.FindElementByClassName("/html/body/div[1]/div[1]/div[2]/h1/span") Is Nothing Then
            If WD.FindElementByClassName("/html/body/div[1]/div[1]/div[2]/h1/span").GetAttribute("textContent") = "這個網頁無法正常運作" Then
                GoTo exits
            End If
        End If
        If VBA.DateDiff("s", dt, VBA.Now) > 2 And e Is Nothing Then
exits:
            word.Application.Activate
            MsgBox "有錯誤，請檢索，或手動執行！", vbCritical
            Exit Function
        End If
    Loop
    
    e.Click
    
    LookupHYDZD = True
    Exit Function

enter:
    '還在檢索頁面 'Set e = WD.FindElementByCssSelector("#keyword")
    If Not WD.FindElementByXPath("/html/body/div[2]/div[2]/form/div[3]/div[1]/input") Is Nothing Then
        WD.FindElementByXPath("/html/body/div[2]/div[2]/form/div[3]/div[1]/input").SendKeys key.enter
        SystemSetup.playSound 1
    End If
    Return
    
End Function

Rem 查中文大辭典（《國學大師》將我所掃的轉成黑白版），成功傳回true
Function LookupZWDCD(x As String) As Boolean
    On Error GoTo eH
    If Not code.IsChineseString(x) Then
        MsgBox "只限中文！", vbCritical
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
    WD.Manage.Timeouts.PageLoad = 2 '設置頁面載入超時3秒 creedit_with_Copilot大菩薩
    WD.SwitchTo.Window WD.CurrentWindowHandle
    ActivateChrome
    word.Application.windowState = wdWindowStateMinimize
    '檢索框
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
    WD.Manage.Timeouts.PageLoad = 10 '設置頁面載入超時秒數 creedit_with_Copilot大菩薩
    On Error Resume Next
    iwe.SendKeys key.enter
    
    On Error GoTo 0
'    '查詢按鈕
'    Set iwe = WD.FindElementByCssSelector("")
'    If iwe Is Nothing Then Exit Function
'    iwe.Click
    actions.Create WD
    dt = VBA.Now
    '書目列表
    Set iwe = WD.FindElementById("shupage")
    Do While iwe Is Nothing And VBA.DateDiff("s", dt, VBA.Now) < 10
        Set iwe = WD.FindElementById("shupage")
        actions.SendKeys(key.End).Perform
    Loop
    actions.SendKeys(key.End).Perform
    '使用 JavaScript 判斷表格總行數是否改變
    Dim prevRowCount As Long, currRowCount As Long
scroll:
    prevRowCount = 0
    dt = VBA.Now
    Do
        currRowCount = WD.ExecuteScript("return document.getElementById('shupage').rows.length") '20241020creedit_with_Copilot大菩薩
        If currRowCount > prevRowCount Then
            prevRowCount = currRowCount
            dt = VBA.Now '重置時間
        End If
        SystemSetup.wait 0.4
        actions.SendKeys(key.End).Perform
        '小停頓
        SystemSetup.wait 0.4 ' 1000 毫秒等於 1 秒'Application.wait (Now + TimeValue("0:00:01"))
        actions.SendKeys(key.End).Perform
        If VBA.DateDiff("s", dt, VBA.Now) > 20 Then '時間可以調整
            WD.Manage.Timeouts.PageLoad = timeoutsPageLoad
            Exit Function
        Else
            actions.SendKeys(key.End).Perform
            SystemSetup.wait 1
        End If
    Loop While currRowCount > prevRowCount

    '找到目標標籤
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
        If tds(i).GetAttribute("textContent") = "中文大" & VBA.ChrW(-28770) & "典" Then
            attr = tds(i + 1).GetAttribute("innerHTML")
            flag = True
            Exit For
        End If
    Next i
    If Not flag Then
        If retyrCntr > 1 Then
            MsgBox "此字沒有《中文大辭典》訊息。感恩感恩　南無阿彌陀佛　讚美主", vbExclamation
        Else
            retyrCntr = retyrCntr + 1
            GoTo scroll
        End If
    End If
    'WD.Manage.timeouts.ImplicitWait = 3 ' 等待3秒
    On Error GoTo 0
    On Error GoTo eH:
    WD.Manage.Timeouts.PageLoad = 4 '設置頁面載入超時x秒 creedit_with_Copilot大菩薩
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
'    WD.Manage.timeouts.ImplicitWait = timeoutsImplicitWait '預設值為0
    WD.Manage.Timeouts.PageLoad = timeoutsPageLoad '預設值為300
    Exit Function
eH:
    Select Case Err.number
        Case -2146233088
            If VBA.InStr(Err.description, "stale element reference: stale element not found in the current frame") = 1 Then 'stale element reference: stale element not found in the current frame
'                                                (Session info: chrome=129.0.6668.101)
                actions.SendKeys(key.End).Perform
                GoTo scroll
            ElseIf VBA.InStr(Err.description, "timeout: Timed out receiving message from renderer:") = 1 Then 'timeout: Timed out receiving message from renderer: 3.000
                                        '  (Session info: chrome=129.0.6668.101)
                If VBA.InStr(WD.url, "zwdcd") Then
    '                WD.Manage.timeouts.ImplicitWait = WD.Manage.timeouts.ImplicitWait + 3 ' 等待3秒
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
            ElseIf VBA.InStr(Err.description, "javascript error: Cannot read properties of null (reading 'rows')") Then '-2146233088javascript error: Cannot read properties of null (reading 'rows')
                                                                                                                    '(Session info: chrome=129.0.6668.101)
                word.Application.Activate
                MsgBox "網站故障，請取消作業或重試。感恩感恩　南無阿彌陀佛　讚美主", vbCritical
                If WD.Manage.Timeouts.PageLoad <> timeoutsPageLoad Then WD.Manage.Timeouts.PageLoad = timeoutsPageLoad
                Exit Function
            Else
                GoTo caseElse
            End If
        Case Else
caseElse:
            Debug.Print Err.number & Err.description
            word.Application.Activate
            MsgBox Err.number & Err.description, vbCritical
    End Select
End Function
Rem 查教育百科_教育雲線上字典 失敗傳回false
Function LookupPediaCloudEduTw(x As String) As Boolean
    Dim e As SeleniumBasic.IWebElement, key As New SeleniumBasic.keys
    If IsDriverInvalid(WD) Then
        If Not OpenChrome("https://pedia.cloud.edu.tw/") Then Exit Function
    End If
    WD.SwitchTo.Window WD.CurrentWindowHandle
    If VBA.InStr(WD.url, "https://pedia.cloud.edu.tw/") <> 1 Then WD.url = "https://pedia.cloud.edu.tw/"
    
    ActivateChrome
    word.Application.windowState = wdWindowStateMinimize
    
    Set e = WD.FindElementByCssSelector("#searchInput")
    If e Is Nothing Then Exit Function
    
    SetIWebElementValueProperty e, x
    e.SendKeys key.enter
    
    LookupPediaCloudEduTw = True
End Function

'強制停止網頁載入 20241020 creedit_with_Copilot大菩薩
Sub StopLoadPage()
    WD.ExecuteScript "window.stop();"
End Sub
Rem 查《國學大師》：x 要查的字詞。傳回一個字串陣列，第1個元素是所查詢的字串，第2個元素是查詢結果網址。若沒找到，則傳回空字串 ""
Function LookupGXDS(x As String) As String()
    On Error GoTo eH
    Dim result(1) As String '1=索引值上限（最大值）
    LookupGXDS = result
    If Not code.IsChineseString(x) Then
        MsgBox "只能檢索中文。請檢查檢索字串，重新開始。", vbCritical
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
    '檢索輸入框
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

    '找到檢索框之後
    If Not iwe Is Nothing Then
        Dim keys As New SeleniumBasic.keys
        iwe.SendKeys keys.Shift + keys.Insert
        'iwe.SendKeys keys.Control + "v"
        iwe.SendKeys keys.enter
        
        '查詢結果訊息框，如 【精确】方式查……，推荐使用【模糊】或【詞首】方式查找。
        Set iwe = WD.FindElementByCssSelector("body > div:nth-child(3) > div.info.l > div.info_content.zj.clearfix > div.info_txt2.clearfix")
        '查詢有結果時：
        If iwe Is Nothing Or VBA.InStr(iwe.text, "【精确】方式查") = 0 Then
            result(0) = x
            result(1) = WD.url
            SystemSetup.SetClipboard result(1)
        End If
    End If
    LookupGXDS = result
    Exit Function
eH:
Select Case Err.number
        Case -2146233088
            If InStr(Err.description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                            '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                            '  (Session info: chrome=128.0.6613.85)
                'Set wd = Nothing
                SystemSetup.killchromedriverFromHere
                Set WD = Nothing
                Resume
            Else
                MsgBox Err.number & Err.description, vbExclamation
            End If
        Case Else
            MsgBox "請關閉Chrome瀏覽器後再執行一次！" & vbCr & vbCr & Err.number & Err.description, vbExclamation
    End Select
End Function
Rem 查《康熙字典網上版》：x 要查的字詞。傳回一個字串陣列，第1個元素是所查詢的字串，第2個元素是查詢結果網址。若沒找到，則傳回空字串 ""
Function LookupKangxizidian(x As String) As String()
    On Error GoTo eH
    Dim result(1) As String '1=索引值上限（最大值）
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
    '檢索輸入框
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

    '找到檢索輸入框
    If Not iwe Is Nothing Then
        Dim keys As New SeleniumBasic.keys
        iwe.Clear
        iwe.SendKeys keys.Shift + keys.Insert
        iwe.SendKeys keys.enter
        '查詢結果訊息框，如： 抱歉，查無資料……請重查！
                                '或請查找以下其他字典:
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
Select Case Err.number
        Case -2146233088
            If InStr(Err.description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                            '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                            '  (Session info: chrome=128.0.6613.85)
                'Set wd = Nothing
                SystemSetup.killchromedriverFromHere
                Set WD = Nothing
                Resume
            Else
                MsgBox Err.number & Err.description, vbExclamation
            End If
        Case Else
            MsgBox "請關閉Chrome瀏覽器後再執行一次！" & vbCr & vbCr & Err.number & Err.description, vbExclamation
    End Select
End Function
Rem 查《白雲深處人家·說文解字·圖像查閱·藤花榭本》：x 要查的字。傳回一個字串陣列，第1個元素是所查詢的字串，第2個元素是查詢結果網址。若沒找到或找到多條，則傳回空字串
Function LookupHomeinmistsShuowenImageAccess_VineyardHall(x As String) As String()
    On Error GoTo eH
    Dim result(1) As String '1=索引值上限（最大值）'預設值即是空陣列
    LookupHomeinmistsShuowenImageAccess_VineyardHall = result '預設值是陣列宣告時配置的，可是尚未把此配置完成的實例陣列賦予或設定給本函式作為傳回值。故本行萬不可省略，除非呼叫端不需要傳回值供處理
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
    '檢索輸入框
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

    '找到檢索輸入框
    Dim keys As New SeleniumBasic.keys
    iwe.Clear
    iwe.SendKeys keys.Shift + keys.Insert
'        iwe.SendKeys keys.Enter'此處按Enter沒作用，須按檢索按鈕
    '檢索按鈕
    Set iwe = WD.FindElementByCssSelector("body > div.search-block > table > tbody > tr > td > input[type=button]")
    GoSub iweNothingExitFunction:
    iwe.Click
    
    '查詢結果訊息框，如：沒有找到。請重新檢索。不支持簡化漢字檢索。
    Set iwe = WD.FindElementByCssSelector("#searchedResults")
    GoSub iweNothingExitFunction:
    If VBA.InStr(iwe.text, "沒有找到。請重新檢索。不支持簡化漢字檢索。") = 1 Then
        Exit Function
    End If
            
    '檢出 n 條
'    Set iwe = wd.FindElementByCssSelector("#searchedResults > span")
    Dim n As Byte '如：檢出 6 條
    n = VBA.CByte(VBA.IIf(VBA.IsNumeric(VBA.Trim(VBA.Replace(VBA.Replace(iwe.text, "檢出", vbNullString), "條", vbNullString))), VBA.Trim(VBA.Replace(VBA.Replace(iwe.text, "檢出", vbNullString), "條", vbNullString)), "0"))
    If n = 0 Then '網頁訊息有錯誤，須檢查(因為找不到時所顯示的是：「沒有找到。請重新檢索。不支持簡化漢字檢索。」
        Exit Function
    End If
    '檢出結果只有一筆才自動開啟其結果連結，否則手動開啟
    If n > 1 Then
        result(0) = x
        '藤花榭本的第2條
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
Select Case Err.number
        Case -2146233088
            If InStr(Err.description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                            '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                            '  (Session info: chrome=128.0.6613.85)
                'Set wd = Nothing
                SystemSetup.killchromedriverFromHere
                Set WD = Nothing
                Resume
            Else
                MsgBox Err.number & Err.description, vbExclamation
            End If
        Case Else
            MsgBox "請關閉Chrome瀏覽器後再執行一次！" & vbCr & vbCr & Err.number & Err.description, vbExclamation
    End Select
End Function

Rem 查《白雲深處人家·說文解字·圖文檢索WFG版》：x 要查的字。若檢索有結果，則傳回true，若沒有，或失敗、故障，則傳回false 20240903
Function LookupHomeinmistsShuowenImageTextSearchWFG_Interpretation(x As String) As Boolean
    On Error GoTo eH
    LookupHomeinmistsShuowenImageTextSearchWFG_Interpretation = False
    If Not code.IsChineseString(x) Then
        Exit Function
    End If
    SystemSetup.SetClipboard x '把檢索條件複製到剪貼簿以備用
    
    If Not OpenChrome("https://homeinmists.ilotus.org/shuowen/WFG2.php") Then
        If Not OpenChrome("https://homeinmists.ilotus.org/shuowen/WFG2.php") Then
            Stop
        End If
    End If
    
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '檢索「解說」內容部分的輸入框
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

    '找到檢索輸入框之後
    Dim keys As New SeleniumBasic.keys
    iwe.Clear
    SetIWebElementValueProperty iwe, x
    'iwe.SendKeys keys.Shift + keys.Insert '貼上檢索內容 = x
'        iwe.SendKeys keys.Enter'此處按Enter沒作用，須按檢索按鈕
    '檢索按鈕
    Set iwe = WD.FindElementByCssSelector("body > div.search-block > div > div:nth-child(2) > input[type=button]") '"body > div.search-block > table > tbody > tr > td > input[type=button]:nth-child(4)")
    GoSub iweNothingExitFunction:
    iwe.Click
    
    '查詢結果訊息框，如：沒有找到。不支持簡化漢字檢索。
    Set iwe = WD.FindElementByCssSelector("#searchedResults")
    GoSub iweNothingExitFunction:
    If VBA.InStr(iwe.text, "沒有找到。不支持簡化漢字檢索。") = 1 Then
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
Select Case Err.number
        Case -2146233088
            If InStr(Err.description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                            '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                            '  (Session info: chrome=128.0.6613.85)
                'Set wd = Nothing
                SystemSetup.killchromedriverFromHere
                Set WD = Nothing
                Resume
            Else
                MsgBox Err.number & Err.description, vbExclamation
            End If
        Case Else
            MsgBox "請關閉Chrome瀏覽器後再執行一次！" & vbCr & vbCr & Err.number & Err.description, vbExclamation
    End Select
End Function
Rem 查白雲深處人家漢語大字典釋義版檢索
Function LookupHomeinmistsHYDZDTextSearch(x As String) As Boolean
    Dim e As SeleniumBasic.IWebElement, si As New StringInfo
    If Not code.IsChineseString(x) Then
        MsgBox "只限中文！", vbCritical
        Exit Function
    End If
    si.Create x
    If si.LengthInTextElements > 1 Then
        MsgBox "限 1 個中文字！", vbCritical
        Exit Function
    End If
    If Not OpenChrome("https://homeinmists.ilotus.org/hd/search.php") Then
        Exit Function
    End If

    WD.Navigate.GoToUrl "https://homeinmists.ilotus.org/hd/search.php"
    '字頭
    Set e = WD.FindElementByCssSelector("#queryString1")
    If e Is Nothing Then
       Exit Function
    End If
    
    ActivateChrome
    WD.SwitchTo.Window WD.CurrentWindowHandle
    word.Application.windowState = wdWindowStateMinimize
    
    SetIWebElementValueProperty e, x
    
    Dim dt As Date
    dt = VBA.Now
    '檢索
    Set e = WD.FindElementByCssSelector("body > div.search-block > div > div:nth-child(1) > input[type=button]")
    Do While e Is Nothing
        Set e = WD.FindElementByCssSelector("body > div.search-block > div > div:nth-child(1) > input[type=button]")
        If VBA.DateDiff("s", dt, VBA.Now) > 2 Then
            LookupHomeinmistsHYDZDTextSearch = False
            Exit Function
        End If
    Loop
    If WD.FindElementByCssSelector("#queryString1").GetAttribute("value") = vbNullString Then
        SetIWebElementValueProperty WD.FindElementByCssSelector("#queryString1"), x
    End If
    e.Click
    
    LookupHomeinmistsHYDZDTextSearch = True
    
End Function
Rem 查《白雲深處人家·說文解字注》：x 要查的字。若檢索有結果，則傳回true，若沒有，或失敗、故障，則傳回false 20250115
Function LookupHomeinmistsShuowenJieZiZhu(x As String) As Boolean
    Dim e As SeleniumBasic.IWebElement, si As New StringInfo
    If Not code.IsChineseString(x) Then
        MsgBox "只限中文！", vbCritical
        Exit Function
    End If
    si.Create x
    If si.LengthInTextElements > 1 Then
        MsgBox "限 1 個中文字！", vbCritical
        Exit Function
    End If
    If Not OpenChrome("https://homeinmists.ilotus.org/shuowen/dyc.php") Then
        Exit Function
    End If
        
    '字頭
    Set e = WD.FindElementByCssSelector("#queryString1")
    If e Is Nothing Then
       Exit Function
    End If
    
    ActivateChrome
    WD.SwitchTo.Window WD.CurrentWindowHandle
    word.Application.windowState = wdWindowStateMinimize
    
    SetIWebElementValueProperty e, x
    '檢索
    Set e = WD.FindElementByCssSelector("body > div:nth-child(8) > input[type=button]:nth-child(2)")
    e.Click
    '圖像欄位的第1個值的儲存格
    Set e = WD.FindElementByCssSelector("#searchTableOut > tr:nth-child(3) > td:nth-child(10) > a")
    
    If Not e Is Nothing Then
        '圖像欄位的第2個值的儲存格
        If Not WD.FindElementByCssSelector("#searchTableOut > tr:nth-child(4) > td:nth-child(10) > a") Is Nothing Then
            LookupHomeinmistsShuowenJieZiZhu = True
            Exit Function
        End If
        
        e.Click
    Else
        Exit Function
    End If
    LookupHomeinmistsShuowenJieZiZhu = True
    
End Function
Rem 查《白雲深處人家》的《漢語大詞典》 20241020
Function LookupHomeinmistsHYDCD(x As String) As Boolean
    If Not code.IsChineseString(x) Then
        MsgBox "只限中文！", vbCritical
        Exit Function
    End If
    Dim iwe As SeleniumBasic.IWebElement, dt As Date, key As New SeleniumBasic.keys
    If Not OpenChrome("https://homeinmists.ilotus.org/hd/hydcd.php") Then Exit Function
    word.Application.windowState = wdWindowStateMinimize
    WD.SwitchTo.Window WD.CurrentWindowHandle
    ActivateChrome
    '檢索框
    dt = VBA.Now
    'Set iwe = WD.FindElementByCssSelector("#keywords")
    Set iwe = WD.FindElementByName("kw")
    Do While iwe Is Nothing
        Set iwe = WD.FindElementByCssSelector("#keywords")
        If VBA.DateDiff("s", dt, VBA.Now) > 5 Then
            Exit Function
        End If
    Loop
    SetIWebElementValueProperty iwe, x
    'iwe.SendKeys key.enter
    '檢索按鈕
    Set iwe = WD.FindElementByCssSelector("body > div.search-block > div > div:nth-child(2) > input[type=button]:nth-child(2)")
    If iwe Is Nothing Then Exit Function
    iwe.Click
    LookupHomeinmistsHYDCD = True
    
End Function

Rem 查《漢語多功能字庫》取回其《說文》「解釋」欄位的內容：x 要查的字。傳回一個字串陣列，第1個元素是《說文》「解釋」的內容字串，第2個元素是查詢結果網址。若沒找到，則傳回空字串
Function LookupMultiFunctionChineseCharacterDatabase(x As String, Optional backgroundStartChrome As Boolean) As String()
    On Error GoTo eH
    Dim result(1) As String '1=索引值上限（最大值）
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
    '檢索輸入框
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

    '找到檢索輸入框之後
    Dim keys As New SeleniumBasic.keys
    iwe.Clear
    iwe.SendKeys keys.Shift + keys.Insert '貼上檢索條件
    iwe.SendKeys keys.enter
    
    '等待檢索結果
    Set iwe = Nothing
    '解釋內容的元件
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
Select Case Err.number
        Case -2146233088
            If InStr(Err.description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                            '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                            '  (Session info: chrome=128.0.6613.85)
                'Set wd = Nothing
                SystemSetup.killchromedriverFromHere
                Set WD = Nothing
                Resume
            Else
                MsgBox Err.number & Err.description, vbExclamation
            End If
        Case Else
            MsgBox "請關閉Chrome瀏覽器後再執行一次！" & vbCr & vbCr & Err.number & Err.description, vbExclamation
    End Select
End Function

Rem 查《韻典網》，成功則傳回true
Function LookupYtenx(x As String) As Boolean
    If Not code.IsChineseString(x) Then
        MsgBox "只限中文！", vbCritical
        Exit Function
    End If
    Dim iwe As SeleniumBasic.IWebElement, dt As Date, key As New SeleniumBasic.keys
    If Not OpenChrome("https://ytenx.org/") Then Exit Function
    word.Application.windowState = wdWindowStateMinimize
    WD.SwitchTo.Window WD.CurrentWindowHandle
    ActivateChrome
    '檢索框
    dt = VBA.Now
    Set iwe = WD.FindElementByCssSelector("#search-form > input.search-query.span3")
    Do While iwe Is Nothing
        Set iwe = WD.FindElementByCssSelector("#search-form > input.search-query.span3")
        If VBA.DateDiff("s", dt, VBA.Now) > 5 Then
            Exit Function
        End If
    Loop
    SetIWebElementValueProperty iwe, x
    iwe.SendKeys key.enter
'    '查詢按鈕
'    Set iwe = WD.FindElementByCssSelector("")
'    If iwe Is Nothing Then Exit Function
'    iwe.Click
    LookupYtenx = True
End Function

Rem 查《說文解字》取回其《說文》「解釋」欄位的內容：x 要查的字,includingDuan 是否也傳回段注內容。傳回一個字串陣列，第1個元素是《說文》（大徐本）「解釋」的內容字串，第2個元素是查詢結果網址，第3個則是段注之內容。若沒找到，則傳回空字串陣列
Function LookupShuowenOrg(x As String, Optional includingDuan As Boolean) As String()
    On Error GoTo eH
    Dim result(2) As String '2=索引值上限（最大值 = UBound 傳回值）
    LookupShuowenOrg = result '先設定好要傳回的字串陣列，當沒賦予值時就是傳回空字串的陣列
    If Not code.IsChineseCharacter(x) Then
        Exit Function
    End If
    SystemSetup.SetClipboard x
    
    If Not OpenChrome("https://www.shuowen.org/") Then
        If Not OpenChrome("https://www.shuowen.org/") Then
            SystemSetup.killchromedriverFromHere
            Set SeleniumOP.WD = Nothing
            MsgBox "請再重試一遍。重試前，請確保Chrome瀏覽器已都關閉。感恩感恩　南無阿彌陀佛　讚美主", vbExclamation
            Stop
            Exit Function
        End If
    End If
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '檢索輸入框
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

    '找到檢索輸入框之後
    Dim keys As New SeleniumBasic.keys
    iwe.Clear
    iwe.SendKeys keys.Shift + keys.Insert '貼上檢索條件
    iwe.SendKeys keys.enter
    
    '等待檢索結果
    Set iwe = WD.FindElementByCssSelector("body > div.container.main > div > div.col-md-9.main-content.pull-right > table > tbody > tr > td")
    If Not iwe Is Nothing Then
        If iwe.text = "沒有記錄" Then
            Exit Function
        Else '如檢索「征」字，因所錄樣有異，故 20240924
            '以檢索結果清單中「楷書」欄名元件來判斷
            'Set iwe = wd.FindElementByCssSelector("body > div.container.main > div > div.col-md-9.main-content.pull-right > table > thead > tr > th:nth-child(1)")
            '來檢索結果訊息框來判斷
            Set iwe = WD.FindElementByCssSelector("body > div.container.main > div > div.col-md-9.main-content.pull-right > div.row.paginator > div.col-md-4.info")
            GoSub iweNothingExitFunction
            'If iwe.Name = "楷書" Then
            Dim Msg As String
            Msg = iwe.GetAttribute("textContent")
            If VBA.IsNumeric(VBA.Left(Msg, 1)) Then '檢索結果訊息框第1個字是數字
                '通常可能會以檢索結果清單中的第1筆為是，如「征」的結果第2個字「徵」，當係簡化字故
                '今擴充為讓使用者輸入整數以指示要讀入的列，只要預設值為 1 即有等效的效果 20240928 creedit_with_Copilot大菩薩
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
'                If VBA.vbOK = MsgBox(msg + vbCr + vbCr + "檢索結果不止一筆，是否要插入第一筆的說文資料？", vbExclamation + vbOKCancel) Then
'                    '檢索結果清單中第1筆的楷書欄位值--即字頭
'                    Set iwe = wd.FindElementByCssSelector("body > div.container.main > div > div.col-md-9.main-content.pull-right > table > tbody > tr:nth-child(1) > td:nth-child(1) > a")
'                    GoSub iweNothingExitFunction
'                    iwe.Click
'                Else
'                    Exit Function
'                End If
reInput:
                r = VBA.InputBox(Msg + vbCr + vbCr + _
                    "檢索結果不止一筆，請輸入要插入第幾筆的說文資料？（正整數）", "請確認要讀入第幾筆的《說文》內容。預設值為 1 ", "1")
                If VBA.IsNumeric(r) = False Then
                    Exit Function
                ElseIf r > UBound(rows) Or r < 0 Then
                    If VBA.vbOK = MsgBox("輸入的數字不對，是否要重新輸入？", vbExclamation + vbOKCancel) Then
                        GoTo reInput
                    Else
                        Exit Function
                    End If
                Else
                    cells = rows(r).FindElementsByTagName("td")
                    Set iwe = cells(0)
                    GoSub iweNothingExitFunction
                    'iwe.Click '無作用（抓到的，似只是結果值元件，不是真正網頁上的元件）
                    Dim outerHTML As String
                    outerHTML = iwe.GetAttribute("outerHTML")
                    If openNewTabWhenTabAlreadyExit(WD) Then
                        WD.Navigate.GoToUrl "https://www.shuowen.org" & VBA.Mid(outerHTML, VBA.InStr(outerHTML, "/"), VBA.InStr(outerHTML, " title=") - 1 - VBA.InStr(outerHTML, "/"))
                    End If
                End If
            End If
        End If
    End If
    '釋文欄的內容
    Set iwe = WD.FindElementByCssSelector("body > div.container.main > div > div.col-md-9.main-content.pull-right > div.row.summary > div.col-md-9.pull-left.info-container > div.media.info-body > div.media-body")
    GoSub iweNothingExitFunction
    result(0) = iwe.text
    result(1) = WD.url
    '取得段注本內容
    If includingDuan Then
        Dim i As Byte
        i = 1
        'Dim duanCommentary As String
        '取得段注本內容框的元件
        Set iwe = WD.FindElementByCssSelector("body > div.container.main > div > div.col-md-9.main-content.pull-right > div:nth-child(" & i & ") > div")
        Do
            If i > 30 Then Exit Do
            If Not iwe Is Nothing Then
                If VBA.InStr(iwe.GetAttribute("textContent"), "清代 段玉裁《說文解字注》") Then Exit Do
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
Select Case Err.number
        Case -2146233088
            If InStr(Err.description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                            '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                            '  (Session info: chrome=128.0.6613.85)
                SystemSetup.killchromedriverFromHere
                Set WD = Nothing
                Resume
            ElseIf InStr(Err.description, "chromedriver.exe does not exist") Then 'The file C:\Program Files\Google\Chrome\Application\chromedriver.exe does not exist. The driver can be downloaded at http://chromedriver.storage.googleapis.com/index.html
                Set WD = Nothing
                MsgBox "請在「" & getChromePathIncludeBackslash & "」路徑下複製chromedriver.exe檔案再繼續！", vbCritical
                SystemSetup.OpenExplorerAtPath getChromePathIncludeBackslash
            Else
                MsgBox Err.number & Err.description, vbExclamation
            End If
        Case Else
            MsgBox "請關閉Chrome瀏覽器後再執行一次！" & vbCr & vbCr & Err.number & Err.description, vbExclamation
    End Select
End Function

Rem 查《異體字字典》取回其「說文釋形」欄位的內容：x 要查的字。傳回一個字串陣列，第1個元素是「說文釋形」的內容字串，第2個元素是查詢結果網址。若沒找到，則傳回空字串陣列 20240916
Function LookupDictionary_of_ChineseCharacterVariants_RetrieveShuoWenData(x As String) As String()
    On Error GoTo eH
    Dim result(1) As String '1=索引值上限（最大值）
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
    '查詢輸入框
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

    '找到查詢輸入框之後
    Dim keys As New SeleniumBasic.keys
    iwe.Clear
    SetIWebElementValueProperty iwe, x
    'iwe.SendKeys keys.Shift + keys.Insert '貼上檢索條件
    iwe.SendKeys keys.enter '有時會失效
    '「查詢」按鈕
    SystemSetup.wait 1
    Set iwe = WD.FindElementByCssSelector("#header > div > flex > div:nth-child(3) > div.quick > form > input[type=submit]:nth-child(5)")
    If Not iwe Is Nothing Then
        iwe.Submit
    End If
    
    dt = VBA.Now
    Set iwe = Nothing
    Do While iwe Is Nothing
        '查詢結果訊息框，如【[ 孫 ]， 查詢結果：正文 1 字，附收字 3 字 】中的「1」這個元件，以此元件來判斷
        Set iwe = WD.FindElementByCssSelector("body > main > div > flex > div:nth-child(1) > red:nth-child(1)")
        '說文釋形欄位
        If Not WD.FindElementByCssSelector("#view > tbody > tr:nth-child(2) > th") Is Nothing Then
            Set iwe = WD.FindElementByCssSelector("#view > tbody > tr:nth-child(2) > th")
            If iwe.GetAttribute("textContent") = "說文釋形" Then
                GoTo shuowenField
            End If
        End If
        Rem 找出來的結果頁面有二：一是列出正文、附收字各字列表的網頁，二是直接進以該字為字頭的網頁
        If DateDiff("s", dt, VBA.Now) > 5 Then
            Exit Function
        End If
    Loop
    If Not iwe Is Nothing Then
        Dim zhengWen As String
        zhengWen = iwe.text '前例的「1」
        '前例的「3」
    
        dt = VBA.Now
        Set iwe = Nothing
        Do While iwe Is Nothing '找「正」字
            Set iwe = WD.FindElementByCssSelector("body > main > div > flex > div:nth-child(1) > red:nth-child(2)")
            If DateDiff("s", dt, VBA.Now) > 5 Then
                Exit Function
            End If
        Loop

        If zhengWen <> "0" Or iwe.text <> "0" Then
            '列出正文、附收字各字列表的網頁
            Set iwe = WD.FindElementByCssSelector("#searchL > a")
            If Not iwe Is Nothing Then
                If VBA.InStr(iwe.GetAttribute("outerHTML"), " data-tp=") = 0 Then
                GoTo plural
                Else
                    Do Until VBA.InStr(iwe.GetAttribute("outerHTML"), " data-tp=""正"" ")
                    Loop
                End If
            Else
plural: '當查詢結果不止一個「字」時，如「去廾」字
'                Stop
                
                Dim ai As Byte
                ai = 2 '#searchL > a:nth-child(4)'#searchL > a:nth-child(3)'#searchL > a:nth-child(2)
                Set iwe = WD.FindElementByCssSelector("#searchL > a:nth-child(" & ai & ")")
                Do Until VBA.InStr(iwe.GetAttribute("outerHTML"), " data-tp=""正"" ")
                    ai = ai + 1
                    Set iwe = WD.FindElementByCssSelector("#searchL > a:nth-child(" & ai & ")")
                Loop
            End If
            iwe.Click
            '先檢查 說文釋形 儲存格 內的文字是否是「說文釋形」
            Set iwe = WD.FindElementByCssSelector("#view > tbody > tr:nth-child(2) > th")
            GoSub iweNothingExitFunction
            If iwe.GetAttribute("textContent") <> "說文釋形" Then
                Set iwe = Nothing
                result(0) = "說文釋形沒有資料！"
                result(1) = WD.url
                GoSub iweNothingExitFunction
            End If
shuowenField:
            '說文釋形 儲存格元件右邊的儲存格
            Set iwe = WD.FindElementByCssSelector("#view > tbody > tr:nth-child(2) > td")
            GoSub iweNothingExitFunction
            If IslinkImageIncluded內容部分包含超連結或圖片(iwe) Then
                result(0) = iwe.GetAttribute("innerHTML")
            Else
                result(0) = iwe.GetAttribute("textContent")
            End If
            result(1) = WD.url
            SystemSetup.SetClipboard result(1)
        End If
    Else
        '如果直接顯示該字頁面，非查詢結果頁，如： https://dict.variants.moe.edu.tw/dictView.jsp?ID=5565
        '字頭元件
        Set iwe = WD.FindElementByCssSelector("#header > section > h2 > span > a")
        If iwe Is Nothing = False Then
        
            '先檢查 說文釋形 儲存格 內的文字是否是「說文釋形」
            Set iwe = WD.FindElementByCssSelector("#view > tbody > tr:nth-child(2) > th")
            GoSub iweNothingExitFunction
            If iwe.GetAttribute("textContent") <> "說文釋形" Then
                Set iwe = Nothing
                result(0) = "說文釋形沒有資料！"
                result(1) = WD.url
                GoSub iweNothingExitFunction
            End If
            '說文釋形 儲存格元件右邊的儲存格
            Set iwe = WD.FindElementByCssSelector("#view > tbody > tr:nth-child(2) > td")
            GoSub iweNothingExitFunction
            result(0) = iwe.GetAttribute("textContent")
            result(1) = WD.url
            SystemSetup.SetClipboard result(1)
        End If
    End If
''''    '等待檢索結果
''''    Set iwe = wd.FindElementByCssSelector("body > div.container.main > div > div.col-md-9.main-content.pull-right > table > tbody > tr > td")
''''    If Not iwe Is Nothing Then
''''        If iwe.text = "沒有記錄" Then
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
Select Case Err.number
        Case -2146233088
            If InStr(Err.description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                            '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                            '  (Session info: chrome=128.0.6613.85)
                'Set wd = Nothing
                SystemSetup.killchromedriverFromHere
                Set WD = Nothing
                Resume
            Else
                MsgBox Err.number & Err.description, vbExclamation
            End If
        Case Else
            MsgBox "請關閉Chrome瀏覽器後再執行一次！" & vbCr & vbCr & Err.number & Err.description, vbExclamation
    End Select
End Function

Rem 檢索Google

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
        iwe.SendKeys keys.enter
    End If
    Exit Sub
    '    Dim form As SeleniumBasic.IWebElement
    '    Dim keyword As SeleniumBasic.IWebElement
    '    Dim button As SeleniumBasic.IWebElement
    '    Set form = wd.FindElementById("form")
    '    Set keyword = form.FindElementById("kw")
    '    Set button = form.FindElementById("su")
    '    keyword.SendKeys VBA.IIf(searchStr = "", Selection, searchStr)
    '    '上一行輸入即檢索了，故可不必下一行;但若不想顯示下拉清單，且確定可顯示結果，則還是需要下一行
    '    button.Click
    ''    Debug.Print WD.title, WD.url
    ''    Debug.Print WD.PageSource
    ''    MsgBox "下面退出瀏覽器。"
    ''    WD.Quit
    '    Exit Sub
Err1:
        Select Case Err.number
'            Case 49 'DLL 呼叫規格錯誤
'                Resume
            Case -2146233088
                If InStr(Err.description, "disconnected: not connected to DevTools") Then 'disconnected: not connected to DevTools
                                                '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                                '  (Session info: chrome=128.0.6613.85)
                    'Set wd = Nothing
                    SystemSetup.killchromedriverFromHere
                    Set WD = Nothing
                    Resume
                Else
                    MsgBox Err.number & Err.description, vbExclamation
                End If
            Case Else
                MsgBox Err.description, vbCritical
                SystemSetup.killchromedriverFromHere
                Set WD = Nothing
    '           Resume
        End Select

End Sub

'貼到古籍酷自動標點,擷取其結果。若失敗，則傳回空字串 vbnullstring
Function grabGjCoolPunctResult(text As String, resultText As String, Optional Background As Boolean) As String
    Const url = "https://gj.cool/punct"
    Dim wdB As SeleniumBasic.IWebDriver, WBQuit As Boolean '=true 則可以關Chrome瀏覽器
    Dim textBox As SeleniumBasic.IWebElement, btn As SeleniumBasic.IWebElement, btn2 As SeleniumBasic.IWebElement, item As SeleniumBasic.IWebElement
    Dim timeout As Byte '最多等 timeOut 秒
    On Error GoTo Err1
    
    If Background Then
        Rem 隱藏
        Set wdB = openChromeBackground(url)
        WBQuit = True '因為在背景執行，預設要可以關'現在用 .AddArgument "--remote-debugging-port=9222"  兼容於其他所開啟者，故不必再背景了 20241003
        If wdB Is Nothing Then
            If WD Is Nothing Then
                If OpenChrome("https://gj.cool/punct") Then
                    Exit Function
                End If
            End If
            Set wdB = WD
        End If
    Else
        Rem 顯示
        If WD Is Nothing Then
            If Not OpenChrome("https://gj.cool/punct") Then
                Exit Function
            End If
        Else
            If Not IsWDInvalid() Then
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
    'If wdB Is Nothing Or IsDriverInvalid(wdB) Then Exit Function
    If wdB Is Nothing Or IsWDInvalid() Then
        Exit Function
    End If
    If wdB.url <> url Then
        '''''''''''If Not IsNewBlankPageTab(wdB) Then OpenNewTab wdB
        wdB.Navigate.GoToUrl url
    End If
    '整理文本
    Dim chkStr As String: chkStr = VBA.Chr(13) & VBA.Chr(10) & VBA.Chr(7) & VBA.Chr(9) & VBA.Chr(8)
    text = VBA.Trim(text)
    Do While VBA.InStr(chkStr, VBA.Left(text, 1)) > 0
        text = VBA.Mid(text, 2)
    Loop
    Do While VBA.InStr(chkStr, VBA.Right(text, 1)) > 0
        text = VBA.Left(text, Len(text) - 1)
    Loop
    
    
    '貼上文本
    Set textBox = wdB.FindElementById("PunctArea")
    If textBox Is Nothing Then
        If IsWDInvalid() Then
            wdB.SwitchTo.Window wdB.WindowHandles()(UBound(wdB.WindowHandles))
        End If
        Set textBox = wdB.FindElementByCssSelector("#PunctArea")
        If textBox Is Nothing Then
            wdB.url = url
            Set textBox = wdB.FindElementByCssSelector("#PunctArea")
        End If
    End If
    Dim key As New SeleniumBasic.keys

'    textBox.Click 20240914作廢
'    textBox.Clear

    'textbox.SendKeys key.LeftShift + key.Insert
    'textbox.SendKeys VBA.KeyCodeConstants.vbKeyControl & VBA.KeyCodeConstants.vbKeyV
    
    '如果只有vba.Chr(13)而沒有vba.Chr(13)&vba.Chr(10)則這行會使分段符號消失；因為下面標點按鈕一按，仍會使一組分段符號消失，必須換成兩組，才能保留一組
    If InStr(text, VBA.Chr(13) & VBA.Chr(10)) = 0 And InStr(text, VBA.Chr(13)) > 0 Then text = Replace(text, VBA.Chr(13), VBA.Chr(13) & VBA.Chr(10) & VBA.Chr(13) & VBA.Chr(10))
    
    SetIWebElement_textContent_Property textBox, text
'    If Background Then 20240914作廢
'        textBox.SendKeys text 'SystemSetup.GetClipboardText
'
'    Else
'        systemsetup.SetClipboard text
'        textBox.SendKeys key.Control + "v"
'    End If
    
    '貼上不成則退出
    Dim WaitDt As Date, chkTxtTime As Date, nx As String, xl As Integer
    
    nx = textBox.text
    text = nx
    SystemSetup.playSound 1.294
    If nx = "" Then
        grabGjCoolPunctResult = ""
        If WBQuit Then wdB.Quit
        Exit Function
    End If
    
    '標點
    'Set btn = wdB.FindElementByCssSelector("#main > div.my-4 > div.p-1.p-md-3.d-flex.justify-content-end > div.ms-2 > button")
    Set btn = wdB.FindElementByCssSelector("#main > div > div.p-1.p-md-3.d-flex.justify-content-end > div:nth-child(6) > button") '20240710
    '即便是有vba.Chr(13)&vba.Chr(10)以下這行仍會使分段符號消失,故若要保持段落，仍須「vba.Chr(13) & vba.Chr(10) & vba.Chr(13) & vba.Chr(10)」二組分段符號，不能只有一個
    If btn Is Nothing Then Stop
    DoEvents
    wdB.SwitchTo().Window (wdB.CurrentWindowHandle)
    SystemSetup.wait 0.9
    'btn.Click
    Dim k As New SeleniumBasic.keys
    btn.SendKeys k.enter
    SystemSetup.playSound 1.469
    '等待標點完成
    'SystemSetup.Wait 3.6
    
    If VBA.Len(text) < 3000 Then
        timeout = 10
    Else
        timeout = 20
    End If
    '最多等 timeOut 秒
    WaitDt = DateAdd("s", timeout, Now()) '極限10秒
    xl = VBA.Len(text)
    chkTxtTime = VBA.Now
    Do
        If VBA.DateDiff("s", chkTxtTime, VBA.Now) > 1.8 Then
            nx = textBox.text
            SystemSetup.playSound 1
            '檢查如果沒有按到「標點」按鈕，就再次按下 20240725 以出現等待圖示控制項為判斷
            If wdB.FindElementByCssSelector("#waitingSpinner") Is Nothing Then
                btn.SendKeys k.enter
            Else
                If wdB.FindElementByCssSelector("#waitingSpinner").Displayed = False And nx = text Then
                    btn.SendKeys k.enter
                    SystemSetup.playSound 1.469
                End If
            End If
            chkTxtTime = Now
            'VBA.StrComp(text, nx) <> 0
            If nx <> text Then Exit Do
            If InStr(nx, "，") > 0 And InStr(nx, "。") > 0 And Len(nx) > xl Then Exit Do
        End If
        If Now > WaitDt Then
            'Exit Do '超過指定時間後離開
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
    ''複製
    'Set item = WDB.FindElementByCssSelector("#main > div > div.p-1.p-md-3.d-flex.justify-content-end > div.dropdown > ul > li:nth-child(4) > a")
    'item.Click
    '
    ''讀取剪貼簿作為回傳值
    'SystemSetup.Wait 0.3
    'systemsetup.SetClipboard textbox.text
    'grabGjCoolPunctResult = SystemSetup.GetClipboardText
    grabGjCoolPunctResult = textBox.text
    resultText = grabGjCoolPunctResult
    If WBQuit = False Then
        'wdB.Close
    Else
        wdB.Quit
        If Not Background Then Set WD = Nothing
    End If
    'Debug.Print grabGjCoolPunctResult
    Exit Function
    
Err1:
        Select Case Err.number
            Case 49 'DLL 呼叫規格錯誤
                Resume
            Case 91 '沒有設定物件變數或 With 區塊變數
                    killchromedriverFromHere
                    OpenChrome url
                    Set wdB = WD
                    If textBox Is Nothing Then Set textBox = wdB.FindElementByCssSelector("#PunctArea")
                    Resume
            Case -2146233088 'unknown error: ChromeDriver only supports characters in the BMP  (Session info: chrome=109.0.5414.75)
                Rem 完全無作用
                Rem systemsetup.SetClipboard text
                Rem SystemSetup.Wait 0.3
                Rem textBox.SendKeys key.Control + "v"
                Rem textBox.SendKeys key.LeftShift + key.Insert
                If InStr(Err.description, "ChromeDriver only supports characters in the BMP") Then
                    WBQuit = pasteWhenOutBMP(wdB, url, "PunctArea", text, textBox, Background)
                    Resume Next
                ElseIf InStr(Err.description, "invalid session id") Or InStr(Err.description, "A exception with a null response was thrown sending an HTTP request to the remote WebDriver server for URL http://localhost:4609/session/455865a54d3f64364cf76b41fe7953a3/url. The status of the exception was ConnectFailure, and the message was: 無法連接至遠端伺服器") Then 'Or InStr(Err.Description, "no such window: target window already closed") Then
                    killchromedriverFromHere
                    OpenChrome url
                    Set wdB = WD: WBQuit = True
                    Resume
                ElseIf InStr(Err.description, "no such window: target window already closed") Then
                    openNewTabWhenTabAlreadyExit wdB
                    wdB.Navigate.GoToUrl url
                    Resume
                ElseIf InStr(Err.description, "disconnected: not connected to DevTools") = 1 Then 'disconnected: not connected to DevTools
                                                                                        '  (failed to check if window was closed: disconnected: not connected to DevTools)
                                                                                        '  (Session info: chrome=130.0.6723.117)
                    killchromedriverFromHere 'WD.Quit: Set WD = Nothing:
                    OpenChrome url
                    Set wdB = WD
                    Resume
                    
                Else
                    MsgBox Err.number & Err.description
                    Stop
                End If
            Case -2147467261 '並未將物件參考設定為物件的執行個體。
                If InStr(Err.description, "並未將物件參考設定為物件的執行個體。") Then
                    killchromedriverFromHere 'WD.Quit: Set WD = Nothing:
                     OpenChrome url
                    Set wdB = WD
                    Resume
                Else
                    MsgBox Err.description, vbCritical
                    Stop
                End If
            Case Else
                MsgBox Err.description, vbCritical
                wdB.Quit
                SystemSetup.killchromedriverFromHere
    '           Resume
        End Select
    
End Function
Function grabGjCoolPunctResult_New(text As String, resultText As String) As String
    Dim iwe As IWebElement, winState As WdWindowState, btn As SeleniumBasic.IWebElement, k As New SeleniumBasic.keys
    If WD Is Nothing Then
'            If Not OpenChrome("https://t.shenshen.wiki/") Then Exit Function
'        Else
'            If IsChromeRunning Then
'                WD.SwitchTo.Window (WD.WindowHandles()(UBound(WD.WindowHandles)))
'            Else
'                If Not OpenChrome("https://t.shenshen.wiki/") Then Exit Function
'            End If
'        End If
    Else
        
        On Error Resume Next
        LastValidWindow = WD.CurrentWindowHandle
        If VBA.InStr(Err.description, "no such window: target window already closed") = 1 Then
            LastValidWindow = WD.WindowHandles()(UBound(WD.WindowHandles))
        End If
        On Error GoTo 0
        
    End If
    If Not OpenChrome("https://gj.cool/punct") Then Exit Function
    winState = word.ActiveDocument.ActiveWindow.windowState
'    WD.Navigate.GoToUrl "https://t.shenshen.wiki/"
    WD.SwitchTo.Window WD.CurrentWindowHandle
    ActivateChrome
    word.ActiveDocument.ActiveWindow.windowState = wdWindowStateMinimize
    '標點區方塊
    Set iwe = WD.FindElementByCssSelector("#PunctArea")
    Dim dt As Date
    dt = DateTime.Now
    Do While iwe Is Nothing
        Set iwe = WD.FindElementByCssSelector("#PunctArea")
        If VBA.DateDiff("s", dt, DateTime.Now) > 5 Then Exit Function
    Loop
    
    If text = vbNullString Or VBA.Replace(text, VBA.Chr(13), vbNullString) = vbNullString Then Exit Function
    SetIWebElement_textContent_Property iwe, text
    
    '執行按鈕
    Set btn = WD.FindElementByCssSelector("#main > div > div.p-1.p-md-3.d-flex.justify-content-end > div:nth-child(6) > button")
    If btn Is Nothing Then Exit Function
    btn.SendKeys k.enter
    
    '執行中的圖示
    Do While WD.FindElementByCssSelector("#waitingSpinner").Displayed = False
        SystemSetup.wait 1
        btn.SendKeys k.enter
        If VBA.DateDiff("s", dt, VBA.Now) > 5 Then
            Exit Do
        End If
    Loop
    
    dt = DateTime.Now
    '結果
'    Set iwe = WD.FindElementByCssSelector("#PunctArea")
'    Do While iwe Is Nothing
'        Set iwe = WD.FindElementByCssSelector("#PunctArea")
'        If VBA.DateDiff("s", dt, VBA.Now) > 5 Then
'            Stop
'            Exit Function
'        End If
'    Loop
'    If iwe Is Nothing Then Set iwe = WD.FindElementByCssSelector("#PunctArea")
    Dim timeout As Byte, textInfo As New StringInfo
    textInfo.Create text
    Select Case textInfo.LengthInTextElements
        Case Is > 2000
            timeout = 36
        Case Is > 1000, Is < 2001
            timeout = 23
        Case Is > 500, Is < 1001
            timeout = 15
        Case Else
            timeout = 10
    End Select
    Do While WD.FindElementByCssSelector("#PunctArea").GetAttribute("textContent") = text
        If VBA.DateDiff("s", dt, DateTime.Now) > timeout Then Exit Function
    Loop
    resultText = WD.FindElementByCssSelector("#PunctArea").GetAttribute("textContent")
    'If UBound(WD.WindowHandles) > 1 Then WD.Close '不關閉，以手動評量其標點良窳
    'If LastValidWindow <> vbNullString Then WD.SwitchTo().Window (LastValidWindow)
    If IsValidWindowHandle(LastValidWindow, WD) Then
        WD.SwitchTo().Window (LastValidWindow)
    Else
        LastValidWindow = WD.CurrentWindowHandle
    End If
    grabGjCoolPunctResult_New = resultText
    word.ActiveDocument.ActiveWindow.windowState = winState
    'word.Application.Activate
    
'    '==========
'    Const url = "https://gj.cool/punct"
'    Dim wdB As SeleniumBasic.IWebDriver, WBQuit As Boolean '=true 則可以關Chrome瀏覽器
'    Dim textBox As SeleniumBasic.IWebElement, btn As SeleniumBasic.IWebElement, btn2 As SeleniumBasic.IWebElement, item As SeleniumBasic.IWebElement
'    Dim timeOut As Byte '最多等 timeOut 秒
'    On Error GoTo Err1
'
'    If Background Then
'        Rem 隱藏
'        Set wdB = openChromeBackground(url)
'        WBQuit = True '因為在背景執行，預設要可以關'現在用 .AddArgument "--remote-debugging-port=9222"  兼容於其他所開啟者，故不必再背景了 20241003
'        If wdB Is Nothing Then
'            If WD Is Nothing Then
'                If OpenChrome("https://gj.cool/punct") Then
'                    Exit Function
'                End If
'            End If
'            Set wdB = WD
'        End If
'    Else
'        Rem 顯示
'        If WD Is Nothing Then
'            If Not OpenChrome("https://gj.cool/punct") Then
'                Exit Function
'            End If
'        Else
'            If Not IsWDInvalid() Then
'                OpenNewTab WD
'            Else
'                'Stop 'just for test
'                killchromedriverFromHere
'                Set WD = Nothing
'                If Not OpenChrome("https://gj.cool/punct") Then
'                    Exit Function
'                End If
'            End If
'        End If
'        Set wdB = WD
'
'    End If
'    'If wdB Is Nothing Or IsDriverInvalid(wdB) Then Exit Function
'    If wdB Is Nothing Or IsWDInvalid() Then
'        Exit Function
'    End If
'    If wdB.url <> url Then
'        '''''''''''If Not IsNewBlankPageTab(wdB) Then OpenNewTab wdB
'        wdB.Navigate.GoToUrl url
'    End If
'    '整理文本
'    Dim chkStr As String: chkStr = VBA.Chr(13) & VBA.Chr(10) & VBA.Chr(7) & VBA.Chr(9) & VBA.Chr(8)
'    text = VBA.Trim(text)
'    Do While VBA.InStr(chkStr, VBA.Left(text, 1)) > 0
'        text = VBA.Mid(text, 2)
'    Loop
'    Do While VBA.InStr(chkStr, VBA.Right(text, 1)) > 0
'        text = VBA.Left(text, Len(text) - 1)
'    Loop
'
'
'    '貼上文本
'    Set textBox = wdB.FindElementById("PunctArea")
'    If textBox Is Nothing Then
'        If IsWDInvalid() Then
'            wdB.SwitchTo.Window wdB.WindowHandles()(UBound(wdB.WindowHandles))
'        End If
'        Set textBox = wdB.FindElementByCssSelector("#PunctArea")
'        If textBox Is Nothing Then
'            wdB.url = url
'            Set textBox = wdB.FindElementByCssSelector("#PunctArea")
'        End If
'    End If
'    Dim key As New SeleniumBasic.keys
'
''    textBox.Click 20240914作廢
''    textBox.Clear
'
'    'textbox.SendKeys key.LeftShift + key.Insert
'    'textbox.SendKeys VBA.KeyCodeConstants.vbKeyControl & VBA.KeyCodeConstants.vbKeyV
'
'    '如果只有vba.Chr(13)而沒有vba.Chr(13)&vba.Chr(10)則這行會使分段符號消失；因為下面標點按鈕一按，仍會使一組分段符號消失，必須換成兩組，才能保留一組
'    If InStr(text, VBA.Chr(13) & VBA.Chr(10)) = 0 And InStr(text, VBA.Chr(13)) > 0 Then text = Replace(text, VBA.Chr(13), VBA.Chr(13) & VBA.Chr(10) & VBA.Chr(13) & VBA.Chr(10))
'
'    SetIWebElement_textContent_Property textBox, text
''    If Background Then 20240914作廢
''        textBox.SendKeys text 'SystemSetup.GetClipboardText
''
''    Else
''        systemsetup.SetClipboard text
''        textBox.SendKeys key.Control + "v"
''    End If
'
'    '貼上不成則退出
'    Dim WaitDt As Date, chkTxtTime As Date, nx As String, xl As Integer
'
'    nx = textBox.text
'    text = nx
'    SystemSetup.playSound 1.294
'    If nx = "" Then
'        grabGjCoolPunctResult = ""
'        If WBQuit Then wdB.Quit
'        Exit Function
'    End If
'
'    '標點
'    'Set btn = wdB.FindElementByCssSelector("#main > div.my-4 > div.p-1.p-md-3.d-flex.justify-content-end > div.ms-2 > button")
'    Set btn = wdB.FindElementByCssSelector("#main > div > div.p-1.p-md-3.d-flex.justify-content-end > div:nth-child(6) > button") '20240710
'    '即便是有vba.Chr(13)&vba.Chr(10)以下這行仍會使分段符號消失,故若要保持段落，仍須「vba.Chr(13) & vba.Chr(10) & vba.Chr(13) & vba.Chr(10)」二組分段符號，不能只有一個
'    If btn Is Nothing Then Stop
'    DoEvents
'    wdB.SwitchTo().Window (wdB.CurrentWindowHandle)
'    SystemSetup.wait 0.9
'    'btn.Click
'    Dim k As New SeleniumBasic.keys
'    btn.SendKeys k.Enter
'    SystemSetup.playSound 1.469
'    '等待標點完成
'    'SystemSetup.Wait 3.6
'
'    If VBA.Len(text) < 3000 Then
'        timeOut = 10
'    Else
'        timeOut = 20
'    End If
'    '最多等 timeOut 秒
'    WaitDt = DateAdd("s", timeOut, Now()) '極限10秒
'    xl = VBA.Len(text)
'    chkTxtTime = VBA.Now
'    Do
'        If VBA.DateDiff("s", chkTxtTime, VBA.Now) > 1.8 Then
'            nx = textBox.text
'            SystemSetup.playSound 1
'            '檢查如果沒有按到「標點」按鈕，就再次按下 20240725 以出現等待圖示控制項為判斷
'            If wdB.FindElementByCssSelector("#waitingSpinner") Is Nothing Then
'                btn.SendKeys k.Enter
'            Else
'                If wdB.FindElementByCssSelector("#waitingSpinner").Displayed = False And nx = text Then
'                    btn.SendKeys k.Enter
'                    SystemSetup.playSound 1.469
'                End If
'            End If
'            chkTxtTime = Now
'            'VBA.StrComp(text, nx) <> 0
'            If nx <> text Then Exit Do
'            If InStr(nx, "，") > 0 And InStr(nx, "。") > 0 And Len(nx) > xl Then Exit Do
'        End If
'        If Now > WaitDt Then
'            'Exit Do '超過指定時間後離開
'            grabGjCoolPunctResult = ""
'            'wdB.Quit
'            wdB.Close
'            SystemSetup.playSound 1.469
'            Exit Function
'        End If
'    Loop
'    'Set btn2 = WDB.FindElementById("dropdownMenuButton2")
'    'btn2.Click
'    '
'    ''複製
'    'Set item = WDB.FindElementByCssSelector("#main > div > div.p-1.p-md-3.d-flex.justify-content-end > div.dropdown > ul > li:nth-child(4) > a")
'    'item.Click
'    '
'    ''讀取剪貼簿作為回傳值
'    'SystemSetup.Wait 0.3
'    'systemsetup.SetClipboard textbox.text
'    'grabGjCoolPunctResult = SystemSetup.GetClipboardText
'    grabGjCoolPunctResult = textBox.text
'    resultText = grabGjCoolPunctResult
'    If WBQuit = False Then
'        'wdB.Close
'    Else
'        wdB.Quit
'        If Not Background Then Set WD = Nothing
'    End If
'    'Debug.Print grabGjCoolPunctResult
'    Exit Function

End Function
Rem Ctrl + Alt + a : [AI太炎](https://t.shenshen.wiki/)標點 20241105
Function grabAITShenShenWikiPunctResult(text As String, resultText As String, Optional Background As Boolean) As String
        '限500字
    Dim strInfo As New StringInfo, iwe As IWebElement, winState As WdWindowState
    strInfo.Create text
    If strInfo.LengthInTextElements > 500 Then
        MsgBox "限500字", vbCritical
        Exit Function
    End If
    If IsWDInvalid() Then
'        If WD Is Nothing Then
'            If Not OpenChrome("https://t.shenshen.wiki/") Then Exit Function
'        Else
'            If IsChromeRunning Then
'                WD.SwitchTo.Window (WD.WindowHandles()(UBound(WD.WindowHandles)))
'            Else
'                If Not OpenChrome("https://t.shenshen.wiki/") Then Exit Function
'            End If
'        End If
    Else
        
        On Error Resume Next
        LastValidWindow = WD.CurrentWindowHandle
        If VBA.InStr(Err.description, "no such window: target window already closed") = 1 Then
            LastValidWindow = WD.WindowHandles()(UBound(WD.WindowHandles))
        End If
        On Error GoTo 0

    End If
'    If Not OpenChrome("https://t.shenshen.wiki/") Then Exit Function
    If Not OpenChrome("https://t.shenshen.wiki/llm") Then Exit Function
    winState = word.ActiveDocument.ActiveWindow.windowState
'    WD.Navigate.GoToUrl "https://t.shenshen.wiki/"
    WD.SwitchTo.Window WD.CurrentWindowHandle
    ActivateChrome
    word.ActiveDocument.ActiveWindow.windowState = wdWindowStateMinimize
    '標點
    Set iwe = WD.FindElementByCssSelector("#nav-biaodian-tab")
    Dim dt As Date
    dt = DateTime.Now
    Do While iwe Is Nothing
        Set iwe = WD.FindElementByCssSelector("#nav-biaodian-tab")
        If VBA.DateDiff("s", dt, DateTime.Now) > 5 Then Exit Function
    Loop
    iwe.Click
    '輸入框
    Set iwe = WD.FindElementByCssSelector("#textarea-biaodian")
    If iwe Is Nothing Then Exit Function
    SetIWebElementValueProperty iwe, text
    '執行
    Set iwe = WD.FindElementByCssSelector("#button-submit")
    If iwe Is Nothing Then Exit Function
    'iwe.Click'在螢幕解析度過大時，不能按到
    Dim key As New SeleniumBasic.keys
    iwe.SendKeys key.enter
    
    dt = DateTime.Now
    '結果怎么樣？
    Set iwe = WD.FindElementByCssSelector("#feedback > div.feedback-button.feedback-tip")
    Do While Not iwe.Displayed ' Is Nothing
        If VBA.DateDiff("s", dt, DateTime.Now) > 36 Then Exit Function
        Set iwe = WD.FindElementByCssSelector("#feedback > div.feedback-button.feedback-tip")
    Loop
    '結果
    Set iwe = WD.FindElementByCssSelector("#output-content")
    If iwe Is Nothing Then Exit Function
    resultText = iwe.GetAttribute("textContent")
    'If UBound(WD.WindowHandles) > 1 Then WD.Close '不關閉，以手動評量其標點良窳
    'If LastValidWindow <> vbNullString Then WD.SwitchTo().Window (LastValidWindow)
    If IsValidWindowHandle(LastValidWindow, WD) Then
        WD.SwitchTo().Window (LastValidWindow)
    Else
        LastValidWindow = WD.CurrentWindowHandle
    End If
    grabAITShenShenWikiPunctResult = resultText
    word.ActiveDocument.ActiveWindow.windowState = winState
    'word.Application.Activate
End Function
Rem 取得《漢籍全文資料庫·斷句十三經經文·周易》文本 ： gua 卦名 。成功則傳回 true 20241004
Function grabHanchiZhouYi_TheOriginalText_ThirteenSutras(gua As String, resultText As String) As Boolean

End Function

Rem 取得《易學網·易經〔周易〕原文》文本。成功則傳回 true 20241004.20241006 resultText是個集合，第1個元素是易卦的內容字串，第2個元素是查詢結果網址。若沒找到，則傳回元素是空字串的陣列
Function GrabEeeLearning_IChing_ZhouYi_originalText(guaSequence As String, resultText As Variant, Optional iwe As SeleniumBasic.IWebElement) As Boolean
'    If Not OpenChrome("https://www.eee-learning.com/article/571") Then Exit Function
    If Not VBA.IsArray(resultText) Then
        MsgBox "第2個引數必須是字串陣列", vbCritical
        'grabEeeLearning_IChing_ZhouYi_originalText = False'預設即為false
    Else
        If UBound(resultText) <> 1 Then
            MsgBox "第2個引數必須是2個元素的字串陣列", vbCritical
        End If
    End If
    Dim e2 As String
    e2 = "https://www.eee-learning.com/book/eee" & guaSequence
    If Not OpenChrome(e2) Then Exit Function
    
    GrabEeeLearning_IChing_ZhouYi_originalText = True
    
    'Dim iwe As SeleniumBasic.IWebElement
    Set iwe = WD.FindElementByCssSelector("#block-bartik-content > div > article > div > div.clearfix.text-formatted.field.field--name-body.field--type-text-with-summary.field--label-hidden.field__item")
    If iwe Is Nothing Then
        GrabEeeLearning_IChing_ZhouYi_originalText = False
        Exit Function
    End If
    
    resultText(0) = iwe.GetAttribute("textContent")
    resultText(1) = e2
    
End Function
Rem 小學堂上古音，傳回 innerHTML 屬性值;出錯則傳回空字串;w:要查的字
Function GrabXiaoxueShangGuYin(w As String) As String
    Const url = "https://xiaoxue.iis.sinica.edu.tw/shangguyin/"
    Dim iwe As SeleniumBasic.IWebElement, key As New SeleniumBasic.keys
    If Not IsWDInvalid Then
        If WD.url <> url Then
            WD.url = url
        Else
            WD.Navigate.Refresh
        End If
    Else
        If Not OpenChrome("https://xiaoxue.iis.sinica.edu.tw/shangguyin/") Then
            Exit Function
        End If
    End If
    
    SeleniumOP.ActivateChrome
    WD.SwitchTo.Window WD.CurrentWindowHandle
    
    '「字形」輸入框
    Set iwe = WD.FindElementByCssSelector("#EudcFontChar")
    SetIWebElementValueProperty iwe, Selection.text
    iwe.SendKeys key.enter
    
    Dim dt As Date
    dt = VBA.Now
    '相關連結
    Do While WD.FindElementByCssSelector("#PageResult > p:nth-child(2)") Is Nothing
        If VBA.DateDiff("s", dt, VBA.Now) > 2 Then Exit Function
    Loop
    'SystemSetup.wait 2.65
    
    Set iwe = WD.FindElementByCssSelector("#PageResult")
    GrabXiaoxueShangGuYin = iwe.GetAttribute("innerHTML") 'outerHTML 也可以
    
End Function
Rem 20240914 creedit_with_Copilot大菩薩：https://sl.bing.net/gCpH6nC61Cu
' 設定元件 IWebElement的value屬性值  20240913
Function SetIWebElementValueProperty(iwe As IWebElement, txt As String) As Boolean
    If Not iwe Is Nothing Then
        'driver.ExecuteScript "arguments[0].value = arguments[1];", element, valueToSet
        WD.ExecuteScript "arguments[0].value = arguments[1];", iwe, txt
        SetIWebElementValueProperty = True
    End If
End Function
Rem 20240914 creedit_with_Copilot大菩薩：https://sl.bing.net/gCpH6nC61Cu
' 設定元件 IWebElement的value屬性值  20240913
Function SetIWebElement_textContent_Property(iwe As IWebElement, txt As String) As Boolean
    If Not iwe Is Nothing Then
        'driver.ExecuteScript "arguments[0].value = arguments[1];", element, valueToSet
        WD.ExecuteScript "arguments[0].textContent = arguments[1];", iwe, txt
        SetIWebElement_textContent_Property = True
    End If
End Function


Private Function pasteWhenOutBMP(ByRef iwd As SeleniumBasic.IWebDriver, url, textBoxToPastedID, pastedTxt As String, ByRef textBox As SeleniumBasic.IWebElement, Background As Boolean) As Boolean ''unknown error: ChromeDriver only supports characters in the BMP  (Session info: chrome=109.0.5414.75)
Rem creedit chatGPT大菩薩：您提到的確實是 Selenium 的 SendKeys 方法不能貼上 BMP 外的字的問題。
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

'貼上
'SystemSetup.Wait 1.5
'textbox.SendKeys key.LeftShift + key.Insert
textBox.SendKeys key.Control + "v"

Exit Function
Err1:
    Select Case Err.number
        Case 49 'DLL 呼叫規格錯誤
            Resume
        Case 91 '未設定物件變數
            If retryTimes > 1 Then
                MsgBox Err.number + Err.description
            Else
                retryTimes = retryTimes + 1
                GoTo retry
            End If
        Case -2146233088 '兩個錯誤的號碼是一樣的，只能用描述來判斷了
        'unknown error: ChromeDriver only supports characters in the BMP  (Session info: chrome=109.0.5414.75)
'            textbox.SendKeys key.LeftShift + key.Insert
'            usePaste WD, url
'            Resume Next
            If InStr(Err.description, "timed out after 60 seconds") Or InStr(Err.description, "無法連接至遠端伺服器") Then
                'The HTTP request to the remote WebDriver server for URL http://localhost:1944/session/d83a0c74803e25f1e7f48999b87a6b7d/element/69589515-4189-4db6-8655-80e30fc05ee0/value timed out after 60 seconds.
                'A exception with a null response was thrown sending an HTTP request to the remote WebDriver server for URL http://localhost:1921/session//element/a9208c93-91ae-4956-9455-d42f51719f23/text. The status of the exception was ConnectFailure, and the message was: 無法連接至遠端伺服器
                iwd.Close
                SystemSetup.killchromedriverFromHere
            End If
        Case -2147467261 '並未將物件參考設定為物件的執行個體。
'            If Not WD Is Nothing Then WD.Quit
            Set WD = Nothing
            GoTo retry
        Case Else
            If WD Is Nothing Then
                OpenChrome (url)
                pasteWhenOutBMP = True
                Resume Next
            Else
                MsgBox Err.description, vbCritical
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
    Select Case Err.number
        Case -2146233088
            If InStr(Err.description, "invalid session id") Then
                SystemSetup.killchromedriverFromHere
            Else
                GoTo Msg
            End If
        Case Else
Msg:
            MsgBox Err.number & Err.description
'            Resume
    End Select
End Property

Rem 20241009 creedit_with_Copilot大菩薩：WordVBA+SeleniumBasic讀入網頁內容圖片與超連結：https://sl.bing.net/hxRfMU08232
Function IslinkImageIncluded內容部分包含超連結或圖片(iwe As SeleniumBasic.IWebElement) As Boolean
    Dim hasLinks As Boolean, arr
    Dim hasImages As Boolean
    
    
    ' 判斷是否包含超連結
    arr = iwe.FindElementsByTagName("a")
    If VBA.IsArray(arr) Then
        links_arrayIWebElement = arr
    End If
    hasLinks = UBound(links_arrayIWebElement) > -1
    
    ' 判斷是否包含圖片
    arr = iwe.FindElementsByTagName("img")
    If VBA.IsArray(arr) Then
        images_arrayIWebElement = arr
    End If
    hasImages = UBound(images_arrayIWebElement) > -1
    
    ' 返回結果
    IslinkImageIncluded內容部分包含超連結或圖片 = hasLinks Or hasImages
End Function
Rem 20241009 creedit_with_Copilot大菩薩：WordVBA+SeleniumBasic讀入網頁內容圖片與超連結：https://sl.bing.net/hxRfMU08232
Function IslinkIncluded內容部分包含超連結(iwe As SeleniumBasic.IWebElement) As Boolean
    Dim hasLinks As Boolean, links
        
    ' 判斷是否包含超連結
    links = iwe.FindElementsByTagName("a")
    If VBA.IsArray(links) Then
        links_arrayIWebElement = links
    End If
    hasLinks = UBound(links_arrayIWebElement) > -1
    ' 返回結果
    IslinkIncluded內容部分包含超連結 = hasLinks
End Function
Rem 20241009 creedit_with_Copilot大菩薩：WordVBA+SeleniumBasic讀入網頁內容圖片與超連結：https://sl.bing.net/hxRfMU08232
Function IsImageIncluded內容部分包含圖片(iwe As SeleniumBasic.IWebElement) As Boolean

    Dim hasImages As Boolean, imgs
    ' 判斷是否包含圖片
    imgs = iwe.FindElementsByTagName("img")
    If VBA.IsArray(imgs) Then
        images_arrayIWebElement = imgs
    End If
    'On Error Resume Next
    hasImages = UBound(images_arrayIWebElement) > -1
    'On Error GoTo 0
    ' 返回結果
    IsImageIncluded內容部分包含圖片 = hasImages
End Function
Property Get Images() As SeleniumBasic.IWebElement()
    Images = images_arrayIWebElement
End Property
Property Get links()
    links = links_arrayIWebElement
End Property

Rem 抓取網頁內容並記錄位置 20241009 creedit_with_Copilot大菩薩：https://sl.bing.net/gGHK9dMCbNQ
Sub grabPageContent抓取網頁內容並記錄位置()
    'Dim wd As New SeleniumBasic.ChromeDriver
    Dim elements() As SeleniumBasic.IWebElement 'Object
    Dim e
    Dim element As SeleniumBasic.IWebElement 'As Object
    Dim elementPositions As Collection
    Set elementPositions = New Collection
    
    '' 打開網頁
    'wd.start "Chrome"
    'wd.Get "https://www.eee-learning.com/article/3694"
    
    ' 抓取所有內容並記錄位置
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
    
    ' 將抓取到的內容插入到Word文件中
    inputElementContent插入網頁元件部分內容 elementPositions
End Sub
Sub inputElementContent插入網頁元件部分內容(elementPositions As Collection)
    Dim elementInfo As Collection, e
    Dim rng As Range
    
    For Each e In elementPositions
        Set elementInfo = e
        Select Case elementInfo(1)
            Case "P", "DIV"
                ' 插入文字
                Set rng = Selection.Range
                rng.text = elementInfo(2)
                Selection.MoveRight Unit:=wdCharacter, Count:=1
            Case "IMG"
                ' 插入圖片
                Set rng = Selection.Range
                rng.InlineShapes.AddPicture fileName:=elementInfo(3), _
                                            LinkToFile:=False, SaveWithDocument:=True
                Selection.MoveRight Unit:=wdCharacter, Count:=1
            Case "A"
                ' 插入超連結
                Set rng = Selection.Range
                ActiveDocument.Hyperlinks.Add Anchor:=rng, _
                                              Address:=elementInfo(4), _
                                              TextToDisplay:=elementInfo(2)
                Selection.MoveRight Unit:=wdCharacter, Count:=1
        End Select
    Next e
End Sub

Rem 遍歷特定範圍內的所有元素，並在特定位置插入圖片和超連結 https://sl.bing.net/f4Mv2PVPse4 20241009 creedit_with_Copilot大菩薩：
Sub inputElementContentAll插入網頁元件所有的內容(iwe As SeleniumBasic.IWebElement)
    'Dim wd As New SeleniumBasic.ChromeDriver
    'Dim iwe As SeleniumBasic.IWebElement
    Dim elements() As SeleniumBasic.IWebElement  'Object
    Dim e, element As SeleniumBasic.IWebElement 'Object
    Dim rng As Range
    
    ' 打開網頁
'    wd.start "Chrome"
'    wd.Get "https://www.eee-learning.com/article/3694"
    
    ' 抓取特定內容部分
    'Set iwe = wd.FindElementByCssSelector("#block-bartik-content > div > article > div > div.clearfix.text-formatted.field.field--name-body.field--type-text-with-summary.field--label-hidden.field__item")
    
    ' 遍歷特定範圍內的所有元素
    elements = iwe.FindElementsByTagName("*")
    
    For Each e In elements
        Set element = e
        
'        Stop
        
'        If SeleniumOP.IsImageIncluded內容部分包含圖片(element) Then
'            insertImageInline插入包含圖片的段落 element
'        End If
        Select Case element.tagname
            Case "p", "div", "span" '"P", "DIV", "STRONG", "SPAN"
                ' "strong" 粗體 .Bold=true 交由呼叫端處理
                ' 插入文字
                Set rng = Selection.Range
                If element.tagname = "strong" Then
                    
                Else
                    rng.text = element.GetAttribute("textContent") 'element.text
                End If
                If SeleniumOP.IsImageIncluded內容部分包含圖片(element) Then
                    'insertImageInline插入包含圖片的段落 element
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
                ' 插入圖片
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
'                ' 插入超連結 rem 先交給呼叫端做
'                Set rng = Selection.Range
'                ActiveDocument.Hyperlinks.Add Anchor:=rng, _
'                                              Address:=element.GetAttribute("href"), _
'                                              TextToDisplay:=element.text
'                Selection.MoveRight Unit:=wdCharacter, Count:=1
        End Select
    Next e
    
'    wd.Quit
End Sub

Rem 處理包含圖片的段落，對於包含圖片的段落，您需要先抓取該段落的文字內容，然後在插入圖片時調整插入點的位置。 20241009 creedit_with_Copilot大菩薩：WordVBA+SeleniumBasic讀入網頁內容圖片與超連結：https://sl.bing.net/2k4A3xjuh2
Private Sub insertImageInline插入包含圖片的段落(iwe As SeleniumBasic.IWebElement)
    'Dim wd As New SeleniumBasic.ChromeDriver
    'Dim
    Dim elements As Object
    Dim element As Object
    Dim rng As Range
    Dim textParts() As String
    Dim i As Integer
    
'    ' 打開網頁
'    wd.start "Chrome"
'    wd.Get "https://www.eee-learning.com/article/3694"
    
'    ' 抓取特定段落
    'Set iwe = wd.FindElementByCssSelector("#block-bartik-content > div > article > div > div.clearfix.text-formatted.field.field--name-body.field--type-text-with-summary.field--label-hidden.field__item > p:nth-child(2)")
    
    ' 分割段落文字內容
    textParts = Split(iwe.GetAttribute("innerHTML"), "<img")
    '如 ： 3 &nbsp;　<img style="border-width:0;" src="/image/yi03b.png" width="28" height="28" align="absbottom" border="0">　<strong>屯卦</strong>　水雷屯
    
    ' 插入文字和圖片
    For i = LBound(textParts) To UBound(textParts)
        If i > 0 Then
            ' 插入圖片
            Set rng = Selection.Range
            rng.InlineShapes.AddPicture images_arrayIWebElement()(0).GetAttribute("src"), LinkToFile:=False, SaveWithDocument:=True
'            rng.InlineShapes.AddPicture fileName:=Mid(textParts(i), InStr(textParts(i), "src=") + 5, InStr(textParts(i), """", InStr(textParts(i), "src=") + 5) - InStr(textParts(i), "src=") - 5), _
                                        LinkToFile:=False, SaveWithDocument:=True
            Selection.MoveRight Unit:=wdCharacter, Count:=1
        End If
        ' 插入文字
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
Rem 判斷圖片載入完成否 20241021creedit_with_Copilot大菩薩:要判斷圖片是否載入成功，可以透過檢查圖片元素的 complete 屬性或者是監聽圖片的 load 事件。這裡有一個簡單的方法，利用 SeleniumBasic 的 JavaScript 執行功能來檢查圖片是否已經載入成功：
Private Function isImageLoaded(CssSelector As String) As Boolean
    Dim script As String
    Dim imgLoaded As Boolean
    
    ' JavaScript 代碼來檢查圖片是否載入成功
    script = "return document.querySelector('" & CssSelector & "').complete;" '這段程式碼會檢查指定的圖片元素的 complete 屬性，如果圖片已經載入，complete 屬性會是 true，否則會是 false。
'    script = "return document.querySelector('body > div:nth-child(3) > center:nth-child(2) > img').complete;"
    ' 執行 JavaScript 並取得結果
    imgLoaded = WD.ExecuteScript(script)
    '如果圖片顯示的是替代文字，通常意味著圖片連結失敗或載入錯誤。在這種情況下，可以檢查圖片的 naturalWidth 和 naturalHeight 屬性。如果這兩個屬性都大於 0，則圖片載入成功；否則，載入失敗。
    '這段程式碼會檢查指定圖片的 naturalWidth 和 naturalHeight 屬性，如果這兩個屬性都大於 0，表示圖片載入成功；否則，圖片載入失敗。這樣應該能更精確地判斷圖片是否載入成功。
    If imgLoaded Then
        ' JavaScript 代碼來檢查圖片的 naturalWidth 和 naturalHeight 屬性
        script = "var img = document.querySelector('body > div:nth-child(3) > center:nth-child(2) > img');return (img.naturalWidth > 0 && img.naturalHeight > 0);"
        ' 執行 JavaScript 並取得結果
        imgLoaded = WD.ExecuteScript(script)
    End If
    If imgLoaded Then
        isImageLoaded = True 'MsgBox "圖片已載入成功"
    Else
        isImageLoaded = False 'MsgBox "圖片尚未載入"
    End If
End Function
Function IsValidWindowHandle(windowHandle As String, driver As SeleniumBasic.IWebDriver) As Boolean
    If windowHandle = vbNullString Then
        IsValidWindowHandle = False
    ElseIf UBound(VBA.Filter(driver.WindowHandles, windowHandle)) > -1 Then
        IsValidWindowHandle = True
    Else
        IsValidWindowHandle = False
    End If
End Function
