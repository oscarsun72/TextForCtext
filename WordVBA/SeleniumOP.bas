Attribute VB_Name = "SeleniumOP"
Option Explicit
Public WD As SeleniumBasic.IWebDriver
Public chromedriversPID() As Long '儲存chromedriver程序ID的陣列
Public chromedriversPIDcntr As Integer 'chromedriversPID的下標值

Sub tesSeleniumBasic() 'https://github.com/florentbr/SeleniumBasic
'20230119 creedit chatGPT大菩薩

    Dim driver As New Selenium.WebDriver
    'driver.start "chrome", "https://www.google.com"
    driver.SetBinary getChromePathIncludeBackslash
    driver.start getChromePathIncludeBackslash + "chrome.exe", "https://www.google.com"
    driver.Get "/"

End Sub

Sub openNewTabWhenTabAlreadyExit(ByRef WD As SeleniumBasic.IWebDriver)
Dim iw As Byte, ew, ii As Byte
For Each ew In WD.WindowHandles
    iw = iw + 1
Next ew
If iw > 0 Then
      WD.ExecuteScript "window.open('about:blank','_blank');"
      For Each ew In WD.WindowHandles
            ii = ii + 1
            If ii = iw + 1 Then Exit For
      Next ew
      WD.SwitchTo().Window (ew)
End If
End Sub
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
'            .AddArgument "--new-window"
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
        openNewTabWhenTabAlreadyExit WD
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

Dim wdB As SeleniumBasic.IWebDriver, WBQuit As Boolean '=true 則可以關Chrome瀏覽器
WBQuit = True '因為在背景執行，預設要可以關

    Set wdB = openChromeBackground(url)
'    If wdB.url <> url Then wdB.url = url
    Dim form As SeleniumBasic.IWebElement
    Dim keyword As SeleniumBasic.IWebElement
    Dim button As SeleniumBasic.IWebElement
    Set form = wdB.FindElementById("searchF")
    Set keyword = form.FindElementByName("word")
    Set button = form.FindElementByClassName("submit")
    keyword.SendKeys searchStr
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
    End If
    Exit Function
Err1:
    Select Case Err.Number
        Case 49 'DLL 呼叫規格錯誤
            Resume
        Case -2146233088 'unknown error: ChromeDriver only supports characters in the BMP  (Session info: chrome=109.0.5414.75)
            'textbox.SendKeys key.LeftShift + key.Insert
            WBQuit = pasteWhenOutBMP(wdB, url, "word", searchStr, keyword)
            Resume Next
        Case Else
            MsgBox Err.Description, vbCritical
            wdB.Quit
            SystemSetup.killchromedriverFromHere
'           Resume
    End Select

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
Function grabGjCoolPunctResult(Text As String) As String
Const url = "https://gj.cool/punct"
Dim wdB As SeleniumBasic.IWebDriver, WBQuit As Boolean '=true 則可以關Chrome瀏覽器
Dim textBox As SeleniumBasic.IWebElement, btn As SeleniumBasic.IWebElement, btn2 As SeleniumBasic.IWebElement, item As SeleniumBasic.IWebElement
On Error GoTo Err1
Set wdB = openChromeBackground(url)
'If WDB Is Nothing Then openChrome ("https://gj.cool/punct")
If wdB Is Nothing Then Exit Function
WBQuit = True '因為在背景執行，預設要可以關
'整理文本
Dim chkStr As String: chkStr = VBA.Chr(13) & Chr(10) & Chr(7) & Chr(9) & Chr(8)
Text = VBA.Trim(Text)
Do While VBA.InStr(chkStr, VBA.Left(Text, 1)) > 0
    Text = Mid(Text, 2)
Loop
Do While VBA.InStr(chkStr, VBA.Right(Text, 1)) > 0
    Text = Left(Text, Len(Text) - 1)
Loop


'貼上文本
Set textBox = wdB.FindElementById("PunctArea")
Dim key As New SeleniumBasic.keys
textBox.Click
textBox.Clear
'textbox.SendKeys key.LeftShift + key.Insert
'textbox.SendKeys VBA.KeyCodeConstants.vbKeyControl & VBA.KeyCodeConstants.vbKeyV

'如果只有chr(13)而沒有chr(13)&chr(10)則這行會使分段符號消失；因為下面標點按鈕一按，仍會使一組分段符號消失，必須換成兩組，才能保留一組
If InStr(Text, Chr(13) & Chr(10)) = 0 And InStr(Text, Chr(13)) > 0 Then Text = Replace(Text, Chr(13), Chr(13) & Chr(10) & Chr(13) & Chr(10))
textBox.SendKeys Text 'SystemSetup.GetClipboardText

'貼上不成則退出
Dim WaitDt As Date, nx As String, xl As Integer

nx = textBox.Text
If nx = "" Then
    grabGjCoolPunctResult = ""
    wdB.Quit
    Exit Function
End If

'標點
Set btn = wdB.FindElementByCssSelector("#main > div.my-4 > div.p-1.p-md-3.d-flex.justify-content-end > div.ms-2 > button")
'即便是有chr(13)&chr(10)以下這行仍會使分段符號消失,故若要保持段落，仍須「Chr(13) & Chr(10) & Chr(13) & Chr(10)」二組分段符號，不能只有一個
btn.Click
'等待標點完成
'SystemSetup.Wait 3.6

WaitDt = DateAdd("s", 10, Now()) '極限10秒
xl = VBA.Len(Text)
Do
    nx = textBox.Text
    'VBA.StrComp(text, nx) <> 0
    If InStr(nx, "，") > 0 And InStr(nx, "。") > 0 And Len(nx) > xl Then Exit Do
    If Now > WaitDt Then
        'Exit Do '超過指定時間後離開
        grabGjCoolPunctResult = ""
        wdB.Quit
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
'SystemSetup.SetClipboard textbox.text
'grabGjCoolPunctResult = SystemSetup.GetClipboardText
grabGjCoolPunctResult = textBox.Text
If WBQuit Then wdB.Close
'Debug.Print grabGjCoolPunctResult
Exit Function

Err1:
    Select Case Err.Number
        Case 49 'DLL 呼叫規格錯誤
            Resume
        Case -2146233088 'unknown error: ChromeDriver only supports characters in the BMP  (Session info: chrome=109.0.5414.75)
            Rem 完全無作用
            Rem SystemSetup.SetClipboard text
            Rem SystemSetup.Wait 0.3
            Rem textBox.SendKeys key.Control + "v"
            Rem textBox.SendKeys key.LeftShift + key.Insert
            WBQuit = pasteWhenOutBMP(wdB, url, "PunctArea", Text, textBox)
            Resume Next
        Case Else
            MsgBox Err.Description, vbCritical
            wdB.Quit
            SystemSetup.killchromedriverFromHere
'           Resume
    End Select

End Function

Private Function pasteWhenOutBMP(ByRef iwd As SeleniumBasic.IWebDriver, url, textBoxToPastedID, pastedTxt As String, ByRef textBox As SeleniumBasic.IWebElement) As Boolean ''unknown error: ChromeDriver only supports characters in the BMP  (Session info: chrome=109.0.5414.75)
Rem creedit chatGPT大菩薩：您提到的確實是 Selenium 的 SendKeys 方法不能貼上 BMP 外的字的問題。
On Error GoTo Err1
DoEvents
'SystemSetup.SetClipboard pastedTxt
'SystemSetup.Wait 0.2
iwd.Quit
retry:
If WD Is Nothing Then
    openChrome (url)
    pasteWhenOutBMP = True
End If
Set iwd = WD
If iwd.url <> url Then iwd.Navigate.GoToUrl (url)
Dim key As New SeleniumBasic.keys
Set textBox = iwd.FindElementById(textBoxToPastedID)
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
            Set WD = Nothing
            GoTo retry
        Case Else
            MsgBox Err.Description, vbCritical
'            WD.Quit
            iwd.Close
            SystemSetup.killchromedriverFromHere
           Resume
    End Select
End Function
