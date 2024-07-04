Attribute VB_Name = "Network"
Option Explicit
Dim DefaultBrowserNameAppActivate As String

Sub 查詢國語辭典() '指定鍵:Ctrl+F12'2010/10/18修訂
''    If ActiveDocument.Path <> "" Then ActiveDocument.Save '怕word當掉忘了儲存
''    If GetUserAddress = True Then
'''        MsgBox "成功的跟隨超連結。"
''    Else
''        MsgBox "無法跟隨超連結。"
''    End If
'    Selection.Copy
'    Shell "W:\!! for hpr\VB\查詢國語辭典\查詢國語辭典\bin\Debug\查詢國語辭典.EXE"
Const st As String = "C:\Program Files\孫守真\查詢國語辭典等\"
Const f As String = "查詢國語辭典.EXE"
Dim funame As String
If Selection.Type = wdSelectionNormal Then
    Selection.Copy
    If Dir(st & f) <> "" Then
        funame = st & f
    ElseIf Dir("C:\Program Files (x86)\孫守真\查詢國語辭典等\" & f) <> "" Then
        funame = "C:\Program Files (x86)\孫守真\查詢國語辭典等\" & f
    ElseIf Dir("W:\!! for hpr\VB\查詢國語辭典\查詢國語辭典\bin\Debug\" & f) <> "" Then
        funame = "W:\!! for hpr\VB\查詢國語辭典\查詢國語辭典\bin\Debug\" & f
    ElseIf Dir("C:\查詢國語辭典\查詢國語辭典\bin\Debug\" & f) <> "" Then
        funame = "C:\查詢國語辭典\查詢國語辭典\bin\Debug\" & f
    ElseIf Dir(UserProfilePath & "Dropbox\VS\VB\查詢國語辭典\查詢國語辭典\bin\Debug\" & f) <> "" Then
        funame = UserProfilePath & "Dropbox\VS\VB\查詢國語辭典\查詢國語辭典\bin\Debug\" & f
    ElseIf Dir("A:\", vbVolume) <> "" Then
        If Dir("A:\Users\oscar\Dropbox\VS\VB\查詢國語辭典\查詢國語辭典\bin\Debug\" & f) <> "" Then _
        funame = "A:\Users\oscar\Dropbox\VS\VB\查詢國語辭典\查詢國語辭典\bin\Debug\" & f
    ElseIf Dir(UserProfilePath & "Dropbox\VS\VB\查詢國語辭典\查詢國語辭典\bin\Debug\" & f) <> "" Then
        funame = UserProfilePath & "Dropbox\VS\VB\查詢國語辭典\查詢國語辭典\bin\Debug\" & f
    Else
        Exit Sub
    End If
    Shell funame
End If
查國語辭典
End Sub

Sub A速檢網路字辭典() '指定鍵:Alt+F12'2010/10/18修訂
Const f As String = "速檢網路字辭典.EXE"
Const st As String = "C:\Program Files\孫守真\速檢網路字辭典\"
Dim funame As String
If Selection.Type = wdSelectionNormal Then
    Selection.Copy
    If Dir(st & f) <> "" Then
        funame = st & f
    ElseIf Dir("C:\Users\oscar\Dropbox\VS\VB\速檢網路字辭典\速檢網路字辭典\bin\Debug\" & f) <> "" Then
        funame = "C:\Users\oscar\Dropbox\VS\VB\速檢網路字辭典\速檢網路字辭典\bin\Debug\" & f
    ElseIf Dir("C:\Program Files (x86)\孫守真\速檢網路字辭典\" & f) <> "" Then
        funame = "C:\Program Files (x86)\孫守真\速檢網路字辭典\" & f
    ElseIf Dir("W:\!! for hpr\VB\速檢網路字辭典\速檢網路字辭典\bin\Debug\" & f) <> "" Then
        funame = "W:\!! for hpr\VB\速檢網路字辭典\速檢網路字辭典\bin\Debug\" & f
    ElseIf Dir("C:\速檢網路字辭典\速檢網路字辭典\bin\Debug\" & f) <> "" Then
        funame = "C:\速檢網路字辭典\速檢網路字辭典\bin\Debug\" & f
    ElseIf Dir(UserProfilePath & "Dropbox\VS\VB\速檢網路字辭典\速檢網路字辭典\bin\Debug\" & f) <> "" Then
        funame = "A:\Users\oscar\Dropbox\VS\VB\速檢網路字辭典\速檢網路字辭典\bin\Debug\" & f
    ElseIf Dir(UserProfilePath & "Dropbox\VS\VB\速檢網路字辭典\速檢網路字辭典\bin\Debug\" & f) <> "" Then
        funame = "A:\Users\oscar\Dropbox\VS\VB\速檢網路字辭典\速檢網路字辭典\bin\Debug\" & f
    ElseIf Dir(UserProfilePath & "Dropbox\VS\VB\速檢網路字辭典\速檢網路字辭典\bin\Debug\" & f) <> "" Then
        funame = UserProfilePath & "Dropbox\VS\VB\速檢網路字辭典\速檢網路字辭典\bin\Debug\" & f
    Else
        Exit Sub
    End If
    Shell funame
End If

End Sub

Sub 查國語辭典()
SeleniumOP.dictRevisedSearch VBA.Replace(Selection, chr(13), "")
End Sub

'Sub 擷取國語辭典詞條網址()
'SeleniumOP.grabDictRevisedUrl VBA.Replace(Selection, Chr(13), "")
'End Sub

Sub 查百度()
SeleniumOP.BaiduSearch Selection
End Sub


Function GetUserAddress() As Boolean
    Dim x As String, a As Object 'Access.Application
    On Error GoTo Error_GetUserAddress
    x = Selection.Text
    Set a = GetObject("D:\千慮一得齋\書籍資料\圖書管理.mdb") '2010/10/18修訂
    If x = "" Then x = InputBox("請輸入欲查詢的字串")
    x = a.Run("查詢字串轉換_國語會碼", x)
''    'ActiveDocument.FollowHyperlink "http://140.111.34.46/cgi-bin/dict/newsearch.cgi", , False, , "Database=dict&GraphicWord=yes&QueryString=^" & X & "$", msoMethodGet
'    FollowHyperlink "http://dict.revised.moe.edu.tw/cgi-bin/newDict/dict.sh?", , False, , "=dict.idx&cond=^" & x & "$&pieceLen=50&fld=1&cat=&imgFont=1", msoMethodGet
    Shell Replace(GetDefaultBrowserEXE, """%1", "http://dict.revised.moe.edu.tw/cgi-bin/newDict/dict.sh?cond=^" & x & "$&pieceLen=50&fld=1&cat=&imgFont=1")
    'AppActivate GetDefaultBrowser'無效
'    'FollowHyperlink "http://dict.revised.moe.edu.tw/cgi-bin/newDict/dict.sh?", , False, , "=dict.idx&cond=^" & X & "$&pieceLen=50&fld=1&cat=&imgFont=1", msoMethodGet
    
'    If Len(Selection.Text) = 1 Then _
        FollowHyperlink "http://www.nlcsearch.moe.gov.tw/EDMS/admin/dict3/search.php", , False, , "qstr=" & x & "&dictlist=47,46,51,18,16,13,20,19,53,12,14,17,48,57,24,25,26,29,30,31,32,33,34,35,36,37,39,38,41,42,43,45,50,&searchFlag=A&hdnCheckAll=checked", msoMethodGet '2009/1/10'教育部-國家語文綜合連結檢索系統-語文綜合檢索
        If a.Visible = False Then
            a.Visible = True
            a.UserControl = True
        End If
'        a.Quit acQuitSaveNone
'        Set a = Nothing
    GetUserAddress = True
Exit_GetUserAddress:
    Exit Function

Error_GetUserAddress:
    MsgBox Err & ": " & Err.Description
    GetUserAddress = False
    Resume Exit_GetUserAddress
End Function


    
Function GetDefaultBrowser() '2010/10/18由http://chijanzen.net/wp/?p=156#comment-1303(取得預設瀏覽器(default web browser)的名稱? chijanzen 雜貨舖)而來.
    Dim objShell
    Set objShell = CreateObject("WScript.Shell")
    'HKEY_CLASSES_ROOT\HTTP\shell\open\ddeexec\Application
    '取得註冊表中的值
    GetDefaultBrowser = objShell.RegRead _
            ("HKCR\http\shell\open\ddeexec\Application\")
    'GetDefaultBrowser = objShell.RegRead _
            ("HKEY_CLASSES_ROOT\http\shell\open\ddeexec\Application\")
End Function


Function GetDefaultBrowserEXE() '2010/10/18由http://chijanzen.net/wp/?p=156#comment-1303(取得預設瀏覽器(default web browser)的名稱? chijanzen 雜貨舖)而來.
Dim deflBrowser As String
deflBrowser = getDefaultBrowserNameAppActivate
Select Case deflBrowser
    Case "iexplore":
        GetDefaultBrowserEXE = "C:\Program Files\Internet Explorer\iexplore.exe"
    Case "firefox":
        If Dir("W:\PortableApps\PortableApps\FirefoxPortable\App\Firefox64\firefox.exe") = "" Then
            GetDefaultBrowserEXE = "C:\Program Files\Mozilla Firefox\firefox.exe"
        Else
            GetDefaultBrowserEXE = "W:\PortableApps\PortableApps\FirefoxPortable\App\Firefox64\firefox.exe"
        End If
    Case "brave":
        If Dir(UserProfilePath & "\AppData\Local\BraveSoftware\Brave-Browser\Application\brave.exe") = "" Then
            GetDefaultBrowserEXE = "C:\Program Files (x86)\BraveSoftware\Brave-Browser\Application\brave.exe"
        Else
            GetDefaultBrowserEXE = UserProfilePath & "\AppData\Local\BraveSoftware\Brave-Browser\Application\brave.exe"
        End If
    Case "vivaldi":
        GetDefaultBrowserEXE = UserProfilePath & "\AppData\Local\Vivaldi\Application\vivaldi.exe"
    Case "Opera":
        GetDefaultBrowserEXE = ""
    Case "Safari":
        GetDefaultBrowserEXE = ""
    Case "edge":
        GetDefaultBrowserEXE = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" '"msedge"
    Case "ChromeHTML", "google chrome": '"chrome"
        GetDefaultBrowserEXE = SystemSetup.getChrome
'
'        If Dir("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe") = "" Then
'            GetDefaultBrowserEXE = "W:\PortableApps\PortableApps\GoogleChromePortable\GoogleChromePortable.exe"
'        Else
'            GetDefaultBrowserEXE = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
'        End If
    Case Else:
        Dim objShell
        Set objShell = CreateObject("WScript.Shell")
        'HKEY_CLASSES_ROOT\HTTP\shell\open\ddeexec\Application
        '取得註冊表中的值
        deflBrowser = objShell.RegRead _
                ("HKCR\http\shell\open\command\")
        GetDefaultBrowserEXE = Mid(deflBrowser, 2, InStr(deflBrowser, ".exe") + Len(".exe") - 2)

End Select
    
    
End Function

Function getDefaultBrowserFullname()
Dim appFullname As String
appFullname = GetDefaultBrowserEXE
'appFullname = Mid(appFullname, 2, InStr(appFullname, ".exe") + Len(".exe") - 2)
getDefaultBrowserFullname = appFullname
'DefaultBrowserNameAppActivate = VBA.Replace(VBA.Mid(appFullname, InStrRev(appFullname, "\") + 1), ".exe", "")
End Function


Function getDefaultBrowserNameAppActivate() As String
Dim objShell, ProgID As String: Set objShell = CreateObject("WScript.Shell")
ProgID = objShell.RegRead _
            ("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice\ProgID")
ProgID = Mid(ProgID, 1, IIf(InStr(ProgID, ".") = 0, Len(ProgID), InStr(ProgID, ".") - 1))
Select Case ProgID
    Case "IE.HTTP":
        DefaultBrowserNameAppActivate = "iexplore"
    Case "FirefoxURL":
        DefaultBrowserNameAppActivate = "firefox"
    Case "ChromeHTML":
        DefaultBrowserNameAppActivate = "google chrome"
    Case "BraveHTML":
        DefaultBrowserNameAppActivate = "brave"
    Case "VivaldiHTM":
        DefaultBrowserNameAppActivate = "vivaldi"
    Case "OperaStable":
        DefaultBrowserNameAppActivate = "Opera"
    Case "SafariHTML":
        DefaultBrowserNameAppActivate = "Safari"
    Case "AppXq0fevzme2pys62n3e0fbqa7peapykr8v", "MSEdgeHTM":
        'browser = BrowserApplication.Edge;
        DefaultBrowserNameAppActivate = "edge" '"msedge"
    Case Else:
        DefaultBrowserNameAppActivate = "google chrome" '"chrome"
End Select
getDefaultBrowserNameAppActivate = DefaultBrowserNameAppActivate
End Function


Sub AppActivateDefaultBrowser()
On Error GoTo eH
Dim i As Byte, a
a = Array("google chrome", "brave", "edge")
DoEvents
If DefaultBrowserNameAppActivate = "" Then getDefaultBrowserNameAppActivate
AppActivate DefaultBrowserNameAppActivate
DoEvents
Exit Sub
eH:
    Select Case Err.Number
        Case 5
            DefaultBrowserNameAppActivate = a(i)
            i = i + 1
            If i > UBound(a) Then
                MsgBox Err.Number & Err.Description
                Exit Sub
            End If
            Resume
        Case Else
            MsgBox Err.Number + Err.Description
    End Select
'AppActivate ""
End Sub



