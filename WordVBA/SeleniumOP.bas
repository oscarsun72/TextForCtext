Attribute VB_Name = "SeleniumOP"
Option Explicit
Public wd As SeleniumBasic.IWebDriver
Public chromedriversPID() As Long '儲存chromedriver程序ID的陣列
Public chromedriversPIDcntr As Integer 'chromedriversPID的下標值
Public ActiveXComponentsCanNotBeCreated As Boolean

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
                MsgBox "若Chrome瀏覽器已開啟，請關閉Chrome瀏覽器後再按確定", vbExclamation
                openChrome "https://www.google.com"
                Resume 'GoTo reOpenChrome:
            ElseIf InStr(Err.Description, "A exception with a null response was thrown sending an HTTP") Then
'                A exception with a null response was thrown sending an HTTP request to the remote WebDriver server for URL http://localhost:1760/session/ed5864479325c154783256563f97e610/window/handles. The status of the exception was ConnectFailure, and the message was: 無法連接至遠端伺服器
                Set wd = Nothing
                killchromedriverFromHere
                MsgBox "若Chrome瀏覽器已開啟，請關閉Chrome瀏覽器後再按確定", vbExclamation
'                openChrome "https://www.google.com"
'                Resume 'GoTo reOpenChrome:
                openNewTabWhenTabAlreadyExit = False
            Else
                MsgBox Err.Number & Err.Description
                Stop
            End If
        Case -2147467261
            If Err.Description = "並未將物件參考設定為物件的執行個體。" Then
                Set wd = Nothing
                killchromedriverFromHere
                MsgBox "若Chrome瀏覽器已開啟，請關閉Chrome瀏覽器後再執行一次", vbExclamation
                openNewTabWhenTabAlreadyExit = False
            Else
                Stop
            End If
        Case 91
            If Err.Description = "沒有設定物件變數或 With 區塊變數" Then
                Set wd = Nothing
                killchromedriverFromHere
                MsgBox "若Chrome瀏覽器已開啟，請關閉Chrome瀏覽器後再執行一次", vbExclamation
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
Rem 啟動Chrome瀏覽器或已啟動後開啟新分頁瀏覽。失敗時傳回false
Function openChrome(Optional URL As String) As Boolean
reStart:
        'Dim WD As SeleniumBasic.IWebDriver
        On Error GoTo ErrH
        Dim Service As SeleniumBasic.ChromeDriverService
        Dim Options As SeleniumBasic.ChromeOptions
        Dim pid As Long
    
    '結束chromedriver.exe
    '使用 WMI 和上面所述的方法
    '判斷PID是否等於pid
    
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
                .HideCommandPromptWindow = True '不顯示命令提示字元視窗
            End With
            Set Options = New SeleniumBasic.ChromeOptions
            With Options
                .BinaryLocation = chromePath + "chrome.exe"
                .AddExcludedArgument "enable-automation" '禁用「Chrome 正在被自動化軟體控制」的警告消息
                
                'C#：options.AddArgument("user-data-dir=" + Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\Google\\Chrome\\User Data\\");
                .AddArgument "user-data-dir=" + VBA.Environ("LOCALAPPDATA") + _
                    "\Google\Chrome\User Data\"
    '            .AddArgument "--new-window"
                '.AddArgument "--start-maximized"
                '.DebuggerAddress = "127.0.0.1:9999" '不要与其他几個混用
            End With
            wd.New_ChromeDriver Service:=Service, Options:=Options
            Docs.Register_Event_Handler '為清除chromedriver作準備
            pid = Service.ProcessId 'Chrome瀏覽器沒有開成功就會是0
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
            If Err.Description = "DLL 呼叫規格錯誤" Then
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
                If MsgBox("須關閉Chrome瀏覽器再繼續！" & vbCr & vbCr & _
                    vbTab & "是否要程式自動幫您關閉、啟動。感恩感恩　南無阿彌陀佛", vbCritical + vbOKCancel) _
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
                If MsgBox("請關閉先前開啟的Chrome瀏覽器再繼續", vbExclamation + vbOKCancel) = vbOK Then
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
                '回到 wd.ExecuteScript "window.open('about:blank','_blank');" 'openNewTabWhenTabAlreadyExit WD
                     'wd.SwitchTo.Window WindowHandlesItem(WindowHandlesCount - 1)
            Else
                
                MsgBox Err.Description, vbCritical
                Stop
            End If
        Case 429 'ActiveX 元件無法產生物件'
            ActiveXComponentsCanNotBeCreated = True
            Exit Function
        Case -2147467261
            If InStr(Err.Description, "並未將物件參考設定為物件的執行個體。") Then
                SystemSetup.killchromedriverFromHere
                Set wd = Nothing
'                Stop
                If MsgBox("須關閉Chrome瀏覽器再繼續！" & vbCr & vbCr & _
                    vbTab & "是否要程式自動幫您關閉、啟動。感恩感恩　南無阿彌陀佛", vbCritical + vbOKCancel) _
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
            .HideCommandPromptWindow = True '不顯示命令提示字元視窗
            If chromedriversPIDcntr = 0 Then chromedriversPIDcntr = 1
            ReDim chromedriversPID(chromedriversPIDcntr - 1)
'            chromedriversPID(chromedriversPIDcntr - 1) = Service.ProcessId'還未啟動=0
        End With
        
        Set Options = New SeleniumBasic.ChromeOptions
        With Options
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
        End With
        wd.New_ChromeDriver Service:=Service, Options:=Options
        'WD.Quit 會自動清除chromedriver，就不用記下開過哪些了
'        pid = Service.ProcessId 'Chrome瀏覽器沒有開成功就會是0
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
            If MsgBox("請關閉先前開啟的Chrome瀏覽器再繼續", vbExclamation + vbOKCancel) = vbOK Then
                'killProcessesByName "ChromeDriver.exe", pid
                killchromedriverFromHere
            GoTo reStart
            End If
        End If
    Case Else
        MsgBox Err.Description, vbCritical
'        Resume
End Select

'20230119 creedit chatGPT大菩薩

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
            '上一行輸入即檢索了，故可不必下一行;但若不想顯示下拉清單，且確定可顯示結果，則還是需要下一行
            button.Click
        End If
    '    Debug.Print WD.title, WD.url
    '    Debug.Print WD.PageSource
    '    MsgBox "下面退出瀏覽器。"
    '    WD.Quit
        Exit Sub
Err1:
        Select Case Err.Number
            Case 49 'DLL 呼叫規格錯誤
                Resume
            Case Else
                MsgBox Err.Description, vbCritical
                SystemSetup.killchromedriverFromHere
    '           Resume
    End Select
End Sub

'找百度 ： https://www.cnblogs.com/ryueifu-VBA/p/13661128.html
Sub BaiduSearch(Optional searchStr As String)
    On Error GoTo Err1
    Search "https://www.baidu.com", "form", "kw", "su", searchStr
        Exit Sub
Err1:
        Select Case Err.Number
            Case 49 'DLL 呼叫規格錯誤
                Resume
            Case Else
                MsgBox Err.Description, vbCritical
                SystemSetup.killchromedriverFromHere
    '           Resume
    End Select
End Sub

'查詢國語辭典
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
                keyword.Submit '這兩個方法都可
    '            Dim k As New SeleniumBasic.keys
    '            keyword.SendKeys k.Enter
            End If
        End If
    '   退出瀏覽器。"
    '    WD.Quit
        Exit Sub
Err1:
        Select Case Err.Number
            Case 49 'DLL 呼叫規格錯誤
                Resume
            Case Else
                MsgBox Err.Description, vbCritical
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
    Dim URL As String, retryTime As Byte
    URL = "https://dict.revised.moe.edu.tw/search.jsp?md=1"
    
    On Error GoTo Err1
    
    Dim wdB As SeleniumBasic.IWebDriver, WBQuit As Boolean '=true 則可以關Chrome瀏覽器
    
    If Background Then
        WBQuit = True '因為在背景執行，預設要可以關
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
        URL = wdB.URL
        If InStr(URL, notFoundOrMultiKey) = 0 Then
            grabDictRevisedUrl_OnlyOneResult = URL '有找到則傳回網址
        Else
            grabDictRevisedUrl_OnlyOneResult = "" '沒有找到傳回空字串
        End If
        If WBQuit Then
            '退出瀏覽器
            wdB.Quit
            If Not Background Then Set wd = Nothing
        Else
            wdB.Close
        End If
        Exit Function
Err1:
        Select Case Err.Number
            Case 49 'DLL 呼叫規格錯誤
                Resume
            Case 91 '沒有設定物件變數或 With 區塊變數
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
            Case -2147467261 '並未將物件參考設定為物件的執行個體。
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
Rem x 要查的字,Variants 要不要看異體字 執行成功傳回true  20240828.
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
    Rem 若須直接查看異體字
    If Variants Then
        Dim dt As Date
        dt = VBA.Now
        Do While iwe Is Nothing
            Set iwe = wd.FindElementByCssSelector("#mainContent > span > div.content > div > div.sidebar_navigation > div > div:nth-child(11)")
            If DateDiff("s", dt, VBA.Now) > 3 Then
                Exit Do '找不到相關字的元件
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
            MsgBox "請關閉Chrome瀏覽器後再執行一次！" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
    End Select
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
        openChrome "https://dict.variants.moe.edu.tw/"
'    Else
'        openNewTabWhenTabAlreadyExit wd
'        wd.Navigate.GoToUrl "https://dict.variants.moe.edu.tw/"
'    End If

    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '檢索輸入框
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
        '查詢結果訊息框，如[ 孫 ]， 查詢結果：正文 1 字，附收字 3 字
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
            '如果直接顯示該字頁面，非查詢結果頁，如： https://dict.variants.moe.edu.tw/dictView.jsp?ID=5565
            '字頭元件
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
            MsgBox "請關閉Chrome瀏覽器後再執行一次！" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
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
    
    openChrome "https://dict.revised.moe.edu.tw/search.jsp?md=1"
    
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '檢索輸入框
    Do While iwe Is Nothing
        Set iwe = wd.FindElementByCssSelector("#searchF > div.line > input[type=text]:nth-child(1)")
        If DateDiff("s", dt, VBA.Now) > 3 Then
            Exit Function
        End If
    Loop
    
    wd.SwitchTo.Window (wd.CurrentWindowHandle)
    word.Application.WindowState = wdWindowStateMinimize
'    VBA.AppActivate "chrome"

    '找到檢索框之後
    If Not iwe Is Nothing Then
        Dim keys As New SeleniumBasic.keys
        'iwe.SendKeys keys.Shift + keys.Insert
        iwe.SendKeys keys.Control + "v"
        iwe.SendKeys keys.Enter
        '查詢結果訊息框，如 查無資料
        Set iwe = wd.FindElementByCssSelector("#searchL > tbody > tr > td")
        '查詢有結果時：
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
            MsgBox "請關閉Chrome瀏覽器後再執行一次！" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
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
    
    If openChrome("https://ivantsoi.myds.me/web/hydcd/search.html") = False Then
        openChrome ("https://ivantsoi.myds.me/web/hydcd/search.html")
        
    End If
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '檢索輸入框
    Do While iwe Is Nothing
        Set iwe = wd.FindElementByCssSelector("#SearchBox")
        If DateDiff("s", dt, VBA.Now) > 3 Then
            Exit Function
        End If
    Loop
    
    wd.SwitchTo.Window (wd.CurrentWindowHandle)
    word.Application.WindowState = wdWindowStateMinimize
'    VBA.AppActivate "chrome"

    '找到檢索框之後
    If Not iwe Is Nothing Then
        Dim keys As New SeleniumBasic.keys
        iwe.SendKeys keys.Shift + keys.Insert
        'iwe.SendKeys keys.Control + "v"
        iwe.SendKeys keys.Enter
        '查詢結果訊息框，如 抱歉，無此詞語。
                        '本掃描版詞典無法查詢簡體字，也無法定位到單字。
                        '若要查單字，可查詢以該字開頭的詞語，再按「上一頁」直到該單字出現，
                        '或使用下面支持單字查詢的《漢語大詞典》連結
                        '或使用《漢語大字典》。
        Set iwe = wd.FindElementByCssSelector("#SearchResult > font")
        '查詢有結果時：
        If iwe Is Nothing Then
            '查詢結果的超連結框
            Set iwe = wd.FindElementByCssSelector("#SearchResult > p > a > font")
            If Not iwe Is Nothing Then
                iwe.Click
                result(0) = x
                wd.SwitchTo.Window WindowHandlesItem(WindowHandlesCount - 1)
                result(1) = wd.URL
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
            MsgBox "請關閉Chrome瀏覽器後再執行一次！" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
    End Select
End Function
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
    
    If openChrome("https://www.guoxuedashi.net/zidian/bujian/") = False Then
        If openChrome("https://www.guoxuedashi.net/zidian/bujian/") = False Then
            Stop
        End If
    End If
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '檢索輸入框
    Do While iwe Is Nothing
        Set iwe = wd.FindElementByCssSelector("#sokeyzi")
        If DateDiff("s", dt, VBA.Now) > 3 Then
            Exit Function
        End If
    Loop
    
    wd.SwitchTo.Window (wd.CurrentWindowHandle)
    word.Application.WindowState = wdWindowStateMinimize
'    VBA.AppActivate "chrome"

    '找到檢索框之後
    If Not iwe Is Nothing Then
        Dim keys As New SeleniumBasic.keys
        iwe.SendKeys keys.Shift + keys.Insert
        'iwe.SendKeys keys.Control + "v"
        iwe.SendKeys keys.Enter
        
        '查詢結果訊息框，如 【精确】方式查……，推荐使用【模糊】或【詞首】方式查找。
        Set iwe = wd.FindElementByCssSelector("body > div:nth-child(3) > div.info.l > div.info_content.zj.clearfix > div.info_txt2.clearfix")
        '查詢有結果時：
        If iwe Is Nothing Or VBA.InStr(iwe.text, "【精确】方式查") = 0 Then
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
            MsgBox "請關閉Chrome瀏覽器後再執行一次！" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
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
    
    If Not openChrome("https://www.kangxizidian.com/search/index.php?stype=Word") Then
        If Not openChrome("https://www.kangxizidian.com/search/index.php?stype=Word") Then
            Stop
        End If
    End If
    
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '檢索輸入框
    Do While iwe Is Nothing
        Set iwe = wd.FindElementByCssSelector("#cornermenubody1 > font18 > input[type=search]:nth-child(2)")
        If DateDiff("s", dt, VBA.Now) > 3 Then
            Exit Function
        End If
    Loop
    
    word.Application.WindowState = wdWindowStateMinimize
    wd.SwitchTo.Window (wd.CurrentWindowHandle)
'    VBA.AppActivate "chrome"

    '找到檢索輸入框
    If Not iwe Is Nothing Then
        Dim keys As New SeleniumBasic.keys
        iwe.Clear
        iwe.SendKeys keys.Shift + keys.Insert
        iwe.SendKeys keys.Enter
        '查詢結果訊息框，如： 抱歉，查無資料……請重查！
                                '或請查找以下其他字典:
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
            MsgBox "請關閉Chrome瀏覽器後再執行一次！" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
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
    
    If Not openChrome("https://homeinmists.ilotus.org/shuowen/find.php") Then
        If Not openChrome("https://homeinmists.ilotus.org/shuowen/find.php") Then
            Stop
        End If
    End If
    
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '檢索輸入框
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

    '找到檢索輸入框
    Dim keys As New SeleniumBasic.keys
    iwe.Clear
    iwe.SendKeys keys.Shift + keys.Insert
'        iwe.SendKeys keys.Enter'此處按Enter沒作用，須按檢索按鈕
    '檢索按鈕
    Set iwe = wd.FindElementByCssSelector("body > div.search-block > table > tbody > tr > td > input[type=button]")
    GoSub iweNothingExitFunction:
    iwe.Click
    
    '查詢結果訊息框，如：沒有找到。請重新檢索。不支持簡化漢字檢索。
    Set iwe = wd.FindElementByCssSelector("#searchedResults")
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
            MsgBox "請關閉Chrome瀏覽器後再執行一次！" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
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
    
    If Not openChrome("https://homeinmists.ilotus.org/shuowen/WFG2.php") Then
        If Not openChrome("https://homeinmists.ilotus.org/shuowen/WFG2.php") Then
            Stop
        End If
    End If
    
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '檢索「解說」內容部分的輸入框
    Do While iwe Is Nothing
        Set iwe = wd.FindElementByCssSelector("#queryString2")
        If DateDiff("s", dt, VBA.Now) > 5 Then
            Exit Function
        End If
    Loop
    
    word.Application.WindowState = wdWindowStateMinimize
    wd.SwitchTo.Window (wd.CurrentWindowHandle)
'    VBA.AppActivate "chrome"

    '找到檢索輸入框之後
    Dim keys As New SeleniumBasic.keys
    iwe.Clear
    iwe.SendKeys keys.Shift + keys.Insert '貼上檢索內容 = x
'        iwe.SendKeys keys.Enter'此處按Enter沒作用，須按檢索按鈕
    '檢索按鈕
    Set iwe = wd.FindElementByCssSelector("body > div.search-block > table > tbody > tr > td > input[type=button]:nth-child(4)")
    GoSub iweNothingExitFunction:
    iwe.Click
    
    '查詢結果訊息框，如：沒有找到。不支持簡化漢字檢索。
    Set iwe = wd.FindElementByCssSelector("#searchedResults")
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
            MsgBox "請關閉Chrome瀏覽器後再執行一次！" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
    End Select
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
    '檢索輸入框
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

    '找到檢索輸入框之後
    Dim keys As New SeleniumBasic.keys
    iwe.Clear
    iwe.SendKeys keys.Shift + keys.Insert '貼上檢索條件
    iwe.SendKeys keys.Enter
    
    '等待檢索結果
    Set iwe = Nothing
    '解釋內容的元件
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
            MsgBox "請關閉Chrome瀏覽器後再執行一次！" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
    End Select
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
    
    If Not openChrome("https://www.shuowen.org/") Then
        If Not openChrome("https://www.shuowen.org/") Then
            Stop
        End If
    End If
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '檢索輸入框
    Do While iwe Is Nothing
        Set iwe = wd.FindElementByCssSelector("#inputKaishu")
        If DateDiff("s", dt, VBA.Now) > 5 Then
            Exit Function
        End If
    Loop
    
    word.Application.WindowState = wdWindowStateMinimize
    wd.SwitchTo.Window (wd.CurrentWindowHandle)
'    VBA.AppActivate "chrome"

    '找到檢索輸入框之後
    Dim keys As New SeleniumBasic.keys
    iwe.Clear
    iwe.SendKeys keys.Shift + keys.Insert '貼上檢索條件
    iwe.SendKeys keys.Enter
    
    '等待檢索結果
    Set iwe = wd.FindElementByCssSelector("body > div.container.main > div > div.col-md-9.main-content.pull-right > table > tbody > tr > td")
    If Not iwe Is Nothing Then
        If iwe.text = "沒有記錄" Then
            Exit Function
        End If
    End If
    '釋文欄的內容
    Set iwe = wd.FindElementByCssSelector("body > div.container.main > div > div.col-md-9.main-content.pull-right > div.row.summary > div.col-md-9.pull-left.info-container > div.media.info-body > div.media-body")
    GoSub iweNothingExitFunction
    result(0) = iwe.text
    result(1) = wd.URL
    '取得段注本內容
    If includingDuan Then
        Dim i As Byte
        i = 1
        'Dim duanCommentary As String
        '取得段注本內容框的元件
        Set iwe = wd.FindElementByCssSelector("body > div.container.main > div > div.col-md-9.main-content.pull-right > div:nth-child(" & i & ") > div")
        Do
            If i > 30 Then Exit Do
            If Not iwe Is Nothing Then
                If VBA.InStr(iwe.GetAttribute("textContent"), "清代 段玉裁《說文解字注》") Then Exit Do
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
            MsgBox "請關閉Chrome瀏覽器後再執行一次！" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
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
    
    If Not openChrome("https://dict.variants.moe.edu.tw/") Then
        If Not openChrome("https://dict.variants.moe.edu.tw/") Then
            Stop
        End If
    End If
    Dim iwe As SeleniumBasic.IWebElement
    Dim dt As Date
    dt = VBA.Now
    '查詢輸入框
    Do While iwe Is Nothing
        Set iwe = wd.FindElementByCssSelector("#header > div > flex > div:nth-child(3) > div.quick > form > input[type=text]:nth-child(2)")
        If DateDiff("s", dt, VBA.Now) > 5 Then
            Exit Function
        End If
    Loop
    
    word.Application.WindowState = wdWindowStateMinimize
    wd.SwitchTo.Window (wd.CurrentWindowHandle)
'    VBA.AppActivate "chrome"

    '找到查詢輸入框之後
    Dim keys As New SeleniumBasic.keys
    iwe.Clear
    iwe.SendKeys keys.Shift + keys.Insert '貼上檢索條件
    iwe.SendKeys keys.Enter
    
    '查詢結果訊息框，如【[ 孫 ]， 查詢結果：正文 1 字，附收字 3 字 】中的「1」這個元件，以此元件來判斷
    Set iwe = wd.FindElementByCssSelector("body > main > div > flex > div:nth-child(1) > red:nth-child(1)")
    Rem 找出來的結果頁面有二：一是列出正文、附收字各字列表的網頁，二是直接進以該字為字頭的網頁
    If Not iwe Is Nothing Then
        Dim zhengWen As String
        zhengWen = iwe.text '前例的「1」
        '前例的「3」
        Set iwe = wd.FindElementByCssSelector("body > main > div > flex > div:nth-child(1) > red:nth-child(2)")
        If zhengWen <> "0" Or iwe.text <> "0" Then
            '列出正文、附收字各字列表的網頁
            Set iwe = wd.FindElementByCssSelector("#searchL > a")
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
                Set iwe = wd.FindElementByCssSelector("#searchL > a:nth-child(" & ai & ")")
                Do Until VBA.InStr(iwe.GetAttribute("outerHTML"), " data-tp=""正"" ")
                    ai = ai + 1
                    Set iwe = wd.FindElementByCssSelector("#searchL > a:nth-child(" & ai & ")")
                Loop
            End If
            iwe.Click
            '先檢查 說文釋形 儲存格 內的文字是否是「說文釋形」
            Set iwe = wd.FindElementByCssSelector("#view > tbody > tr:nth-child(2) > th")
            GoSub iweNothingExitFunction
            If iwe.GetAttribute("textContent") <> "說文釋形" Then
                Set iwe = Nothing
                result(0) = "說文釋形沒有資料！"
                result(1) = wd.URL
                GoSub iweNothingExitFunction
            End If
            '說文釋形 儲存格元件右邊的儲存格
            Set iwe = wd.FindElementByCssSelector("#view > tbody > tr:nth-child(2) > td")
            GoSub iweNothingExitFunction
            result(0) = iwe.GetAttribute("textContent")
            result(1) = wd.URL
            SystemSetup.SetClipboard result(1)
        End If
    Else
        '如果直接顯示該字頁面，非查詢結果頁，如： https://dict.variants.moe.edu.tw/dictView.jsp?ID=5565
        '字頭元件
        Set iwe = wd.FindElementByCssSelector("#header > section > h2 > span > a")
        If iwe Is Nothing = False Then
        
            '先檢查 說文釋形 儲存格 內的文字是否是「說文釋形」
            Set iwe = wd.FindElementByCssSelector("#view > tbody > tr:nth-child(2) > th")
            GoSub iweNothingExitFunction
            If iwe.GetAttribute("textContent") <> "說文釋形" Then
                Set iwe = Nothing
                result(0) = "說文釋形沒有資料！"
                result(1) = wd.URL
                GoSub iweNothingExitFunction
            End If
            '說文釋形 儲存格元件右邊的儲存格
            Set iwe = wd.FindElementByCssSelector("#view > tbody > tr:nth-child(2) > td")
            GoSub iweNothingExitFunction
            result(0) = iwe.GetAttribute("textContent")
            result(1) = wd.URL
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
            MsgBox "請關閉Chrome瀏覽器後再執行一次！" & vbCr & vbCr & Err.Number & Err.Description, vbExclamation
    End Select
End Function

Rem 檢索Google

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
    '    '上一行輸入即檢索了，故可不必下一行;但若不想顯示下拉清單，且確定可顯示結果，則還是需要下一行
    '    button.Click
    ''    Debug.Print WD.title, WD.url
    ''    Debug.Print WD.PageSource
    ''    MsgBox "下面退出瀏覽器。"
    ''    WD.Quit
    '    Exit Sub
Err1:
        Select Case Err.Number
'            Case 49 'DLL 呼叫規格錯誤
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

'貼到古籍酷自動標點()
Function grabGjCoolPunctResult(text As String, resultText As String, Optional Background As Boolean) As String
    Const URL = "https://gj.cool/punct"
    Dim wdB As SeleniumBasic.IWebDriver, WBQuit As Boolean '=true 則可以關Chrome瀏覽器
    Dim textBox As SeleniumBasic.IWebElement, btn As SeleniumBasic.IWebElement, btn2 As SeleniumBasic.IWebElement, item As SeleniumBasic.IWebElement
    Dim timeOut As Byte '最多等 timeOut 秒
    On Error GoTo Err1
    
    If Background Then
        Rem 隱藏
        Set wdB = openChromeBackground(URL)
        WBQuit = True '因為在背景執行，預設要可以關
        If wdB Is Nothing Then
            If wd Is Nothing Then openChrome ("https://gj.cool/punct")
            Set wdB = wd
        End If
    Else
        Rem 顯示
        If wd Is Nothing Then openChrome ("https://gj.cool/punct")
        Set wdB = wd
    End If
    If wdB Is Nothing Then Exit Function
    If wdB.URL <> URL Then wdB.Navigate.GoToUrl URL
    
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
    btn.SendKeys k.Enter
    SystemSetup.playSound 1.469
    '等待標點完成
    'SystemSetup.Wait 3.6
    
    If VBA.Len(text) < 3000 Then
        timeOut = 10
    Else
        timeOut = 20
    End If
    '最多等 timeOut 秒
    WaitDt = DateAdd("s", timeOut, Now()) '極限10秒
    xl = VBA.Len(text)
    chkTxtTime = VBA.Now
    Do
        If VBA.DateDiff("s", chkTxtTime, VBA.Now) > 1.5 Then
            nx = textBox.text
            SystemSetup.playSound 1
            '檢查如果沒有按到「標點」按鈕，就再次按下 20240725 以出現等待圖示控制項為判斷
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
        wdB.Close
    Else
        wdB.Quit
        If Not Background Then Set wd = Nothing
    End If
    'Debug.Print grabGjCoolPunctResult
    Exit Function
    
Err1:
        Select Case Err.Number
            Case 49 'DLL 呼叫規格錯誤
                Resume
            Case 91 '沒有設定物件變數或 With 區塊變數
                    killchromedriverFromHere
                    openChrome URL
                    Set wdB = wd
                Resume
            Case -2146233088 'unknown error: ChromeDriver only supports characters in the BMP  (Session info: chrome=109.0.5414.75)
                Rem 完全無作用
                Rem systemsetup.SetClipboard text
                Rem SystemSetup.Wait 0.3
                Rem textBox.SendKeys key.Control + "v"
                Rem textBox.SendKeys key.LeftShift + key.Insert
                If InStr(Err.Description, "ChromeDriver only supports characters in the BMP") Then
                    WBQuit = pasteWhenOutBMP(wdB, URL, "PunctArea", text, textBox, Background)
                    Resume Next
                ElseIf InStr(Err.Description, "invalid session id") Or InStr(Err.Description, "A exception with a null response was thrown sending an HTTP request to the remote WebDriver server for URL http://localhost:4609/session/455865a54d3f64364cf76b41fe7953a3/url. The status of the exception was ConnectFailure, and the message was: 無法連接至遠端伺服器") Then 'Or InStr(Err.Description, "no such window: target window already closed") Then
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
            Case -2147467261 '並未將物件參考設定為物件的執行個體。
                If InStr(Err.Description, "並未將物件參考設定為物件的執行個體。") Then
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
Rem 20240914 creedit_with_Copilot大菩薩：https://sl.bing.net/gCpH6nC61Cu
' 設定元件 IWebElement的value屬性值  20240913
Private Function SetIWebElementValueProperty(iwe As IWebElement, txt As String) As Boolean
    If Not iwe Is Nothing Then
        'driver.ExecuteScript "arguments[0].value = arguments[1];", element, valueToSet
        wd.ExecuteScript "arguments[0].value = arguments[1];", iwe, txt
        SetIWebElementValueProperty = True
    End If
End Function
Rem 20240914 creedit_with_Copilot大菩薩：https://sl.bing.net/gCpH6nC61Cu
' 設定元件 IWebElement的value屬性值  20240913
Private Function SetIWebElement_textContent_Property(iwe As IWebElement, txt As String) As Boolean
    If Not iwe Is Nothing Then
        'driver.ExecuteScript "arguments[0].value = arguments[1];", element, valueToSet
        wd.ExecuteScript "arguments[0].textContent = arguments[1];", iwe, txt
        SetIWebElement_textContent_Property = True
    End If
End Function


Private Function pasteWhenOutBMP(ByRef iwd As SeleniumBasic.IWebDriver, URL, textBoxToPastedID, pastedTxt As String, ByRef textBox As SeleniumBasic.IWebElement, Background As Boolean) As Boolean ''unknown error: ChromeDriver only supports characters in the BMP  (Session info: chrome=109.0.5414.75)
Rem creedit chatGPT大菩薩：您提到的確實是 Selenium 的 SendKeys 方法不能貼上 BMP 外的字的問題。
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

'貼上
'SystemSetup.Wait 1.5
'textbox.SendKeys key.LeftShift + key.Insert
textBox.SendKeys key.Control + "v"

Exit Function
Err1:
    Select Case Err.Number
        Case 49 'DLL 呼叫規格錯誤
            Resume
        Case 91 '未設定物件變數
            If retryTimes > 1 Then
                MsgBox Err.Number + Err.Description
            Else
                retryTimes = retryTimes + 1
                GoTo retry
            End If
        Case -2146233088 '兩個錯誤的號碼是一樣的，只能用描述來判斷了
        'unknown error: ChromeDriver only supports characters in the BMP  (Session info: chrome=109.0.5414.75)
'            textbox.SendKeys key.LeftShift + key.Insert
'            usePaste WD, url
'            Resume Next
            If InStr(Err.Description, "timed out after 60 seconds") Or InStr(Err.Description, "無法連接至遠端伺服器") Then
                'The HTTP request to the remote WebDriver server for URL http://localhost:1944/session/d83a0c74803e25f1e7f48999b87a6b7d/element/69589515-4189-4db6-8655-80e30fc05ee0/value timed out after 60 seconds.
                'A exception with a null response was thrown sending an HTTP request to the remote WebDriver server for URL http://localhost:1921/session//element/a9208c93-91ae-4956-9455-d42f51719f23/text. The status of the exception was ConnectFailure, and the message was: 無法連接至遠端伺服器
                iwd.Close
                SystemSetup.killchromedriverFromHere
            End If
        Case -2147467261 '並未將物件參考設定為物件的執行個體。
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

