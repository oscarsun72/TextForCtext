Attribute VB_Name = "SeleniumOP"
Option Explicit
Public WD As SeleniumBasic.IWebDriver
Public chromedriversPID() As Long '儲存chromedriver程序ID的陣列
Public chromedriversPIDcntr As Integer 'chromedriversPID的下標值
Sub openChrome(Optional url As String)
reStart:
    'Dim WD As SeleniumBasic.IWebDriver
    On Error GoTo ErrH
    Dim Service As SeleniumBasic.ChromeDriverService
    Dim Options As SeleniumBasic.ChromeOptions
    Dim pid As Long

'結束chromedriver.exe
'使用 WMI 和上面所述的方法
'判斷PID是否等於pid

    If WD Is Nothing Then
        Set WD = New SeleniumBasic.IWebDriver
        Set Service = New SeleniumBasic.ChromeDriverService
        With Service
            .CreateDefaultService driverPath:=getChromePathIncludeBackslash
            '.CreateDefaultService driverPath:="E:\Selenium\Drivers"
            .HideCommandPromptWindow = True '不顯示命令提示字元視窗
        End With
        Set Options = New SeleniumBasic.ChromeOptions
        With Options
            '.BinaryLocation = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
            .BinaryLocation = getChromePathIncludeBackslash + "chrome.exe"
            .AddExcludedArgument "enable-automation" '禁用「Chrome 正在被自動化軟體控制」的警告消息
            
            'C#：options.AddArgument("user-data-dir=" + Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\Google\\Chrome\\User Data\\");
            .AddArgument "user-data-dir=" + VBA.Environ("LOCALAPPDATA") + _
                "\Google\Chrome\User Data\"
            
            '.AddArgument "--start-maximized"
            '.DebuggerAddress = "127.0.0.1:9999" '不要与其他几個混用
        End With
        WD.New_ChromeDriver Service:=Service, Options:=Options
        pid = Service.ProcessId 'Chrome瀏覽器沒有開成功就會是0
        If pid <> 0 Then
            ReDim Preserve chromedriversPID(chromedriversPIDcntr)
            chromedriversPID(chromedriversPIDcntr) = pid
            chromedriversPIDcntr = chromedriversPIDcntr + 1
        End If
        WD.url = url
    End If
Exit Sub
ErrH:
Select Case Err.Number

    Case -2146233088 '**'
        'Debug.Print Err.Description
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
End Sub

Function openChromeBackground(url As String) As SeleniumBasic.IWebDriver
reStart:
    'Dim WD As SeleniumBasic.IWebDriver
    On Error GoTo ErrH
    Dim WD As SeleniumBasic.IWebDriver
    Dim Service As SeleniumBasic.ChromeDriverService
    Dim Options As SeleniumBasic.ChromeOptions
    Dim pid As Long
    
        Set WD = New SeleniumBasic.IWebDriver
        Set Service = New SeleniumBasic.ChromeDriverService
        With Service
            .CreateDefaultService driverPath:=getChromePathIncludeBackslash
            .HideCommandPromptWindow = True '不顯示命令提示字元視窗
        End With
        Set Options = New SeleniumBasic.ChromeOptions
        With Options
            .BinaryLocation = getChromePathIncludeBackslash + "chrome.exe"
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
        WD.New_ChromeDriver Service:=Service, Options:=Options
        'WD.Quit 會自動清除chromedriver，就不用記下開過哪些了
'        pid = Service.ProcessId 'Chrome瀏覽器沒有開成功就會是0
'        If pid <> 0 Then
'            ReDim Preserve chromedriversPID(chromedriversPIDcntr)
'            chromedriversPID(chromedriversPIDcntr) = pid
'            chromedriversPIDcntr = chromedriversPIDcntr + 1
'        End If

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
        If MsgBox("請關閉先前開啟的Chrome瀏覽器再繼續", vbExclamation + vbOKCancel) = vbOK Then
                'killProcessesByName "ChromeDriver.exe", pid
                killchromedriverFromHere
            GoTo reStart
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
Sub Search(url As String, frmID As String, keywdID As String, btnID As String, Optional searchStr As String)
On Error GoTo Err1
'If searchStr = "" And Selection = "" Then Exit Sub
If WD Is Nothing Then
    openChrome (url)
End If
    WD.url = url
    Dim form As SeleniumBasic.IWebElement
    Dim keyword As SeleniumBasic.IWebElement
    Dim button As SeleniumBasic.IWebElement
    Set form = WD.FindElementById(frmID)
    Set keyword = form.FindElementById(keywdID)
    Set button = form.FindElementById(btnID)
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
Const url As String = "https://dict.revised.moe.edu.tw/search.jsp?md=1"
If WD Is Nothing Then
    openChrome (url)
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
Function grabDictRevisedUrl_OnlyOneResult(searchStr As String) As String
'If searchStr = "" And Selection = "" Then Exit Sub
If searchStr = "" Then Exit Function
If VBA.Left(searchStr, 1) <> "=" Then searchStr = "=" + searchStr '精確搜尋字串指令
Const notFoundOrMultiKey As String = "&qMd=0&qCol=1" '查無資料或如果不止一條時，網址後綴都有此關鍵字
Dim url As String
url = "https://dict.revised.moe.edu.tw/search.jsp?md=1"

On Error GoTo Err1

Dim WD As SeleniumBasic.IWebDriver
    Set WD = openChromeBackground(url)
'    If WD.url <> url Then WD.url = url
    Dim form As SeleniumBasic.IWebElement
    Dim keyword As SeleniumBasic.IWebElement
    Dim button As SeleniumBasic.IWebElement
    Set form = WD.FindElementById("searchF")
    Set keyword = form.FindElementByName("word")
    Set button = form.FindElementByClassName("submit")
    keyword.SendKeys searchStr
    If Not button Is Nothing Then
        button.Click
    Else
        keyword.Submit '這兩個方法都可
'            Dim k As New SeleniumBasic.keys
'            keyword.SendKeys k.Enter
    End If
    url = WD.url
    If InStr(url, notFoundOrMultiKey) = 0 Then
        grabDictRevisedUrl_OnlyOneResult = url '有找到則傳回網址
    Else
        grabDictRevisedUrl_OnlyOneResult = "" '沒有找到傳回空字串
    End If
    '退出瀏覽器
    WD.Quit
    Exit Function
Err1:
    MsgBox Err.Description, vbCritical
'    Resume

End Function

Sub GoogleSearch(Optional searchStr As String) '有空再完成
On Error GoTo Err1
'If searchStr = "" And Selection = "" Then Exit Sub
'Dim wd As SeleniumBasic.IWebDriver
'Set wd = openChrome("https://www.baidu.com")
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
        Case 49 'DLL 呼叫規格錯誤
            Resume
        Case Else
            MsgBox Err.Description, vbCritical
            SystemSetup.killchromedriverFromHere
'           Resume
    End Select

End Sub

'貼到古籍酷自動標點()
Function grabGjCoolPunctResult(text As String) As String
Dim WD As SeleniumBasic.IWebDriver
Dim textbox As SeleniumBasic.IWebElement, btn As SeleniumBasic.IWebElement, btn2 As SeleniumBasic.IWebElement, item As SeleniumBasic.IWebElement
On Error GoTo Err1
Set WD = openChromeBackground("https://gj.cool/punct")
'If WD Is Nothing Then openChrome ("https://gj.cool/punct")
If WD Is Nothing Then Exit Function

'貼上文本
Set textbox = WD.FindElementById("PunctArea")
Dim key As New SeleniumBasic.keys
textbox.Click
textbox.Clear
'textbox.SendKeys key.LeftShift + key.Insert
'textbox.SendKeys VBA.KeyCodeConstants.vbKeyControl & VBA.KeyCodeConstants.vbKeyV
textbox.SendKeys text 'SystemSetup.GetClipboardText
'textbox.SendKeys sys
'標點
Set btn = WD.FindElementByCssSelector("#main > div.my-4 > div.p-1.p-md-3.d-flex.justify-content-end > div.ms-2 > button")
btn.Click
'等待標點完成
'SystemSetup.Wait 3.6
Dim WaitDt As Date
WaitDt = DateAdd("s", 6, Now()) '極限6秒

Do While VBA.StrComp(text, textbox.text) = 0
    If Now > WaitDt Then
        'Exit Do '超過指定時間後離開
        grabGjCoolPunctResult = ""
        Exit Function
    End If
Loop
'Set btn2 = WD.FindElementById("dropdownMenuButton2")
'btn2.Click
'
''複製
'Set item = WD.FindElementByCssSelector("#main > div > div.p-1.p-md-3.d-flex.justify-content-end > div.dropdown > ul > li:nth-child(4) > a")
'item.Click
'
''讀取剪貼簿作為回傳值
'SystemSetup.Wait 0.3
'SystemSetup.SetClipboard textbox.text
'grabGjCoolPunctResult = SystemSetup.GetClipboardText
grabGjCoolPunctResult = textbox.text
WD.Quit
'Debug.Print grabGjCoolPunctResult
Exit Function

Err1:
    Select Case Err.Number
        Case 49 'DLL 呼叫規格錯誤
            Resume
        Case -2146233088 'unknown error: ChromeDriver only supports characters in the BMP  (Session info: chrome=109.0.5414.75)
            textbox.SendKeys key.LeftShift + key.Insert
            Resume Next
        Case Else
            MsgBox Err.Description, vbCritical
            SystemSetup.killchromedriverFromHere
'           Resume
    End Select

End Function
