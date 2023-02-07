Attribute VB_Name = "SeleniumOP"
Option Explicit
Public WD As SeleniumBasic.IWebDriver
Public chromedriversPID() As Long '�x�schromedriver�{��ID���}�C
Public chromedriversPIDcntr As Integer 'chromedriversPID���U�Э�

Sub tesSeleniumBasic() 'https://github.com/florentbr/SeleniumBasic
'20230119 creedit chatGPT�j����

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
'            .AddArgument "--new-window"
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

Dim wdB As SeleniumBasic.IWebDriver, WBQuit As Boolean '=true �h�i�H��Chrome�s����
WBQuit = True '�]���b�I������A�w�]�n�i�H��

    Set wdB = openChromeBackground(url)
'    If wdB.url <> url Then wdB.url = url
    Dim form As SeleniumBasic.IWebElement
    Dim keyword As SeleniumBasic.IWebElement
    Dim button As SeleniumBasic.IWebElement
    Set form = wdB.FindElementById("searchF")
    Set keyword = form.FindElementByName("word")
    Set button = form.FindElementByClassName("submit")
    keyword.SendKeys searchStr
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
    End If
    Exit Function
Err1:
    Select Case Err.Number
        Case 49 'DLL �I�s�W����~
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
Function grabGjCoolPunctResult(Text As String) As String
Const url = "https://gj.cool/punct"
Dim wdB As SeleniumBasic.IWebDriver, WBQuit As Boolean '=true �h�i�H��Chrome�s����
Dim textBox As SeleniumBasic.IWebElement, btn As SeleniumBasic.IWebElement, btn2 As SeleniumBasic.IWebElement, item As SeleniumBasic.IWebElement
On Error GoTo Err1
Set wdB = openChromeBackground(url)
'If WDB Is Nothing Then openChrome ("https://gj.cool/punct")
If wdB Is Nothing Then Exit Function
WBQuit = True '�]���b�I������A�w�]�n�i�H��
'��z�奻
Dim chkStr As String: chkStr = VBA.Chr(13) & Chr(10) & Chr(7) & Chr(9) & Chr(8)
Text = VBA.Trim(Text)
Do While VBA.InStr(chkStr, VBA.Left(Text, 1)) > 0
    Text = Mid(Text, 2)
Loop
Do While VBA.InStr(chkStr, VBA.Right(Text, 1)) > 0
    Text = Left(Text, Len(Text) - 1)
Loop


'�K�W�奻
Set textBox = wdB.FindElementById("PunctArea")
Dim key As New SeleniumBasic.keys
textBox.Click
textBox.Clear
'textbox.SendKeys key.LeftShift + key.Insert
'textbox.SendKeys VBA.KeyCodeConstants.vbKeyControl & VBA.KeyCodeConstants.vbKeyV

'�p�G�u��chr(13)�ӨS��chr(13)&chr(10)�h�o��|�Ϥ��q�Ÿ������F�]���U�����I���s�@���A���|�Ϥ@�դ��q�Ÿ������A����������աA�~��O�d�@��
If InStr(Text, Chr(13) & Chr(10)) = 0 And InStr(Text, Chr(13)) > 0 Then Text = Replace(Text, Chr(13), Chr(13) & Chr(10) & Chr(13) & Chr(10))
textBox.SendKeys Text 'SystemSetup.GetClipboardText

'�K�W�����h�h�X
Dim WaitDt As Date, nx As String, xl As Integer

nx = textBox.Text
If nx = "" Then
    grabGjCoolPunctResult = ""
    wdB.Quit
    Exit Function
End If

'���I
Set btn = wdB.FindElementByCssSelector("#main > div.my-4 > div.p-1.p-md-3.d-flex.justify-content-end > div.ms-2 > button")
'�Y�K�O��chr(13)&chr(10)�H�U�o�椴�|�Ϥ��q�Ÿ�����,�G�Y�n�O���q���A�����uChr(13) & Chr(10) & Chr(13) & Chr(10)�v�G�դ��q�Ÿ��A����u���@��
btn.Click
'���ݼ��I����
'SystemSetup.Wait 3.6

WaitDt = DateAdd("s", 10, Now()) '����10��
xl = VBA.Len(Text)
Do
    nx = textBox.Text
    'VBA.StrComp(text, nx) <> 0
    If InStr(nx, "�A") > 0 And InStr(nx, "�C") > 0 And Len(nx) > xl Then Exit Do
    If Now > WaitDt Then
        'Exit Do '�W�L���w�ɶ������}
        grabGjCoolPunctResult = ""
        wdB.Quit
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
'SystemSetup.SetClipboard textbox.text
'grabGjCoolPunctResult = SystemSetup.GetClipboardText
grabGjCoolPunctResult = textBox.Text
If WBQuit Then wdB.Close
'Debug.Print grabGjCoolPunctResult
Exit Function

Err1:
    Select Case Err.Number
        Case 49 'DLL �I�s�W����~
            Resume
        Case -2146233088 'unknown error: ChromeDriver only supports characters in the BMP  (Session info: chrome=109.0.5414.75)
            Rem �����L�@��
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
Rem creedit chatGPT�j���ġG�z���쪺�T��O Selenium �� SendKeys ��k����K�W BMP �~���r�����D�C
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

'�K�W
'SystemSetup.Wait 1.5
'textbox.SendKeys key.LeftShift + key.Insert
textBox.SendKeys key.Control + "v"

Exit Function
Err1:
    Select Case Err.Number
        Case 49 'DLL �I�s�W����~
            Resume
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
            MsgBox Err.Description, vbCritical
'            WD.Quit
            iwd.Close
            SystemSetup.killchromedriverFromHere
           Resume
    End Select
End Function
