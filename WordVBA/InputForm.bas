Attribute VB_Name = "InputForm"
Option Explicit

'' 建立全域連線物件
'Dim conn As Object
'
'' 初始化連線
'Sub InitConnection()
'    Set conn = CreateObject("ADODB.Connection")
'    ' 請修改路徑為您的 Access 資料庫檔案
'    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\MyDB\MyData.accdb;"
'End Sub

' 動態建立表單 KanripoGitHub存放庫造字圖碼取代對照表 新增資料 20260127
Function CreateInputForm_KanjiCharacterCodeReplacementReferenceTableinKanripoGitHubRepository(tbFindText As String, conn As ADODB.Connection) As Object
    Dim uf As Object
    Set uf = ThisDocument.VBProject.VBComponents.Add(3) ' 3 = UserForm
    
    With uf
        .Properties("Caption") = "輸入資料到 Access"
        .Properties("Width") = 300
        .Properties("Height") = 200
    End With
    
    ' 新增文字框 find
    Dim tbFind As Object
    Set tbFind = uf.Designer.Controls.Add("Forms.TextBox.1", "txtFind", True)
    tbFind.Left = 20: tbFind.Top = 30: tbFind.width = 200
    'tbFind.text = ""
    tbFind.text = tbFindText
    
    ' 新增文字框 replace
    Dim tbReplace As Object
    Set tbReplace = uf.Designer.Controls.Add("Forms.TextBox.1", "txtReplace", True)
    tbReplace.Left = 20: tbReplace.Top = 70: tbReplace.width = 200
    tbReplace.text = ""
    
    ' 新增按鈕
    Dim btnAdd As Object
    Set btnAdd = uf.Designer.Controls.Add("Forms.CommandButton.1", "btnAdd", True)
    btnAdd.Caption = "新增"
    btnAdd.Left = 20: btnAdd.Top = 110: btnAdd.width = 80
    
    ' 加入事件程式碼
    Dim code As String
    code = ""
    code = code & "Private Sub btnAdd_Click()" & vbCrLf
    code = code & "    Dim sql As String" & vbCrLf
    code = code & "    sql = ""INSERT INTO MyTable (find, replace) VALUES ('"" & txtFind.Text & ""','"" & txtReplace.Text & ""')""" & vbCrLf
    code = code & "    conn.Execute sql" & vbCrLf
    code = code & "    MsgBox ""資料已新增"", vbInformation" & vbCrLf
    code = code & "End Sub"
    
    uf.CodeModule.AddFromString code
    
    VBA.UserForms.Add(uf.name).Show
    
    Set CreateInputForm_KanjiCharacterCodeReplacementReferenceTableinKanripoGitHubRepository = uf
End Function



'' 建立全域連線物件
'Dim conn As Object
'
'' 初始化連線
'Sub InitConnection()
'    Set conn = CreateObject("ADODB.Connection")
'    ' 請修改路徑為您的 Access 資料庫檔案
'    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\MyDB\MyData.accdb;"
'End Sub
'
'' 動態建立表單
'Sub CreateInputForm()
'    Dim uf As Object
'    Set uf = ThisDocument.VBProject.VBComponents.Add(3) ' 3 = UserForm
'
'    With uf
'        .Properties("Caption") = "輸入資料到 Access"
'        .Properties("Width") = 300
'        .Properties("Height") = 200
'    End With
'
'    ' 新增文字框 find
'    Dim tbFind As Object
'    Set tbFind = uf.Designer.Controls.Add("Forms.TextBox.1", "txtFind", True)
'    tbFind.Left = 20: tbFind.Top = 30: tbFind.width = 200
'    tbFind.text = ""
'
'    ' 新增文字框 replace
'    Dim tbReplace As Object
'    Set tbReplace = uf.Designer.Controls.Add("Forms.TextBox.1", "txtReplace", True)
'    tbReplace.Left = 20: tbReplace.Top = 70: tbReplace.width = 200
'    tbReplace.text = ""
'
'    ' 新增按鈕
'    Dim btnAdd As Object
'    Set btnAdd = uf.Designer.Controls.Add("Forms.CommandButton.1", "btnAdd", True)
'    btnAdd.Caption = "新增"
'    btnAdd.Left = 20: btnAdd.Top = 110: btnAdd.width = 80
'
'    ' 加入事件程式碼
'    Dim code As String
'    code = ""
'    code = code & "Private Sub btnAdd_Click()" & vbCrLf
'    code = code & "    Dim sql As String" & vbCrLf
'    code = code & "    sql = ""INSERT INTO MyTable (find, replace) VALUES ('"" & txtFind.Text & ""','"" & txtReplace.Text & ""')""" & vbCrLf
'    code = code & "    conn.Execute sql" & vbCrLf
'    code = code & "    MsgBox ""資料已新增"", vbInformation" & vbCrLf
'    code = code & "End Sub"
'
'    uf.CodeModule.AddFromString code
'
'    VBA.UserForms.Add(uf.name).Show
'End Sub
'

Rem https://copilot.microsoft.com/shares/ECL2ZMUQK8tyyp8LPujp9

Rem https://copilot.microsoft.com/shares/2N2irTn2gH3kMJFo4JjUi
