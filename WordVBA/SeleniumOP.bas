Attribute VB_Name = "SeleniumOP"
Option Explicit
Public WD As SeleniumBasic.IWebDriver
Public chromedriversPID() As Long '�x�schromedriver�{��ID���}�C
Public chromedriversPIDcntr As Integer 'chromedriversPID���U�Э�
Sub openChrome(Optional url As String)
reStart:
    'Dim WD As SeleniumBasic.IWebDriver
    On Error GoTo ErrH
    Dim Service As SeleniumBasic.ChromeDriverService
    Dim Options As SeleniumBasic.ChromeOptions
    Dim pid As Long

'����chromedriver.exe
'�ϥ� WMI �M�W���ҭz����k
'�P�_PID�O�_����pid

    If WD Is Nothing Then
        Set WD = New SeleniumBasic.IWebDriver
        Set Service = New SeleniumBasic.ChromeDriverService
        With Service
            .CreateDefaultService driverPath:=getChromePathIncludeBackslash
            '.CreateDefaultService driverPath:="E:\Selenium\Drivers"
            .HideCommandPromptWindow = True '����ܩR�O���ܦr������
        End With
        Set Options = New SeleniumBasic.ChromeOptions
        With Options
            '.BinaryLocation = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
            .BinaryLocation = getChromePathIncludeBackslash + "chrome.exe"
            .AddExcludedArgument "enable-automation" '�T�ΡuChrome ���b�Q�۰ʤƳn�鱱��v��ĵ�i����
            
            'C#�Goptions.AddArgument("user-data-dir=" + Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\Google\\Chrome\\User Data\\");
            .AddArgument "user-data-dir=" + VBA.Environ("LOCALAPPDATA") + _
                "\Google\Chrome\User Data\"
            
            '.AddArgument "--start-maximized"
            '.DebuggerAddress = "127.0.0.1:9999" '���n�O��L�L�ӲV��
        End With
        WD.New_ChromeDriver Service:=Service, Options:=Options
        pid = Service.ProcessId 'Chrome�s�����S���}���\�N�|�O0
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
        If MsgBox("���������e�}�Ҫ�Chrome�s�����A�~��", vbExclamation + vbOKCancel) = vbOK Then
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

'20230119 creedit chatGPT�j����

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
            .HideCommandPromptWindow = True '����ܩR�O���ܦr������
        End With
        Set Options = New SeleniumBasic.ChromeOptions
        With Options
            .BinaryLocation = getChromePathIncludeBackslash + "chrome.exe"
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
        WD.New_ChromeDriver Service:=Service, Options:=Options
        'WD.Quit �|�۰ʲM��chromedriver�A�N���ΰO�U�}�L���ǤF
'        pid = Service.ProcessId 'Chrome�s�����S���}���\�N�|�O0
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
        If MsgBox("���������e�}�Ҫ�Chrome�s�����A�~��", vbExclamation + vbOKCancel) = vbOK Then
                'killProcessesByName "ChromeDriver.exe", pid
                killchromedriverFromHere
            GoTo reStart
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
Function grabDictRevisedUrl_OnlyOneResult(searchStr As String) As String
'If searchStr = "" And Selection = "" Then Exit Sub
If searchStr = "" Then Exit Function
If VBA.Left(searchStr, 1) <> "=" Then searchStr = "=" + searchStr '��T�j�M�r����O
Const notFoundOrMultiKey As String = "&qMd=0&qCol=1" '�d�L��ƩΦp�G����@���ɡA���}��󳣦�������r
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
        keyword.Submit '�o��Ӥ�k���i
'            Dim k As New SeleniumBasic.keys
'            keyword.SendKeys k.Enter
    End If
    url = WD.url
    If InStr(url, notFoundOrMultiKey) = 0 Then
        grabDictRevisedUrl_OnlyOneResult = url '�����h�Ǧ^���}
    Else
        grabDictRevisedUrl_OnlyOneResult = "" '�S�����Ǧ^�Ŧr��
    End If
    '�h�X�s����
    WD.Quit
    Exit Function
Err1:
    MsgBox Err.Description, vbCritical
'    Resume

End Function

Sub GoogleSearch(Optional searchStr As String) '���ŦA����
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
'    '�W�@���J�Y�˯��F�A�G�i�����U�@��;���Y���Q��ܤU�ԲM��A�B�T�w�i��ܵ��G�A�h�٬O�ݭn�U�@��
'    button.Click
''    Debug.Print WD.title, WD.url
''    Debug.Print WD.PageSource
''    MsgBox "�U���h�X�s�����C"
''    WD.Quit
'    Exit Sub
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

'�K��j�y�Ŧ۰ʼ��I()
Function grabGjCoolPunctResult(text As String) As String
Dim WD As SeleniumBasic.IWebDriver
Dim textbox As SeleniumBasic.IWebElement, btn As SeleniumBasic.IWebElement, btn2 As SeleniumBasic.IWebElement, item As SeleniumBasic.IWebElement
On Error GoTo Err1
Set WD = openChromeBackground("https://gj.cool/punct")
'If WD Is Nothing Then openChrome ("https://gj.cool/punct")
If WD Is Nothing Then Exit Function

'�K�W�奻
Set textbox = WD.FindElementById("PunctArea")
Dim key As New SeleniumBasic.keys
textbox.Click
textbox.Clear
'textbox.SendKeys key.LeftShift + key.Insert
'textbox.SendKeys VBA.KeyCodeConstants.vbKeyControl & VBA.KeyCodeConstants.vbKeyV
textbox.SendKeys text 'SystemSetup.GetClipboardText
'textbox.SendKeys sys
'���I
Set btn = WD.FindElementByCssSelector("#main > div.my-4 > div.p-1.p-md-3.d-flex.justify-content-end > div.ms-2 > button")
btn.Click
'���ݼ��I����
'SystemSetup.Wait 3.6
Dim WaitDt As Date
WaitDt = DateAdd("s", 6, Now()) '����6��

Do While VBA.StrComp(text, textbox.text) = 0
    If Now > WaitDt Then
        'Exit Do '�W�L���w�ɶ������}
        grabGjCoolPunctResult = ""
        Exit Function
    End If
Loop
'Set btn2 = WD.FindElementById("dropdownMenuButton2")
'btn2.Click
'
''�ƻs
'Set item = WD.FindElementByCssSelector("#main > div > div.p-1.p-md-3.d-flex.justify-content-end > div.dropdown > ul > li:nth-child(4) > a")
'item.Click
'
''Ū���ŶKï�@���^�ǭ�
'SystemSetup.Wait 0.3
'SystemSetup.SetClipboard textbox.text
'grabGjCoolPunctResult = SystemSetup.GetClipboardText
grabGjCoolPunctResult = textbox.text
WD.Quit
'Debug.Print grabGjCoolPunctResult
Exit Function

Err1:
    Select Case Err.Number
        Case 49 'DLL �I�s�W����~
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
