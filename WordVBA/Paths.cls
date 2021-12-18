VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Paths"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Dim fso As Object
Function getdb_重編國語辭典修訂本_資料庫fullName()
Dim cntStr As String
'On Error GoTo eH
userProfilePath = SystemSetup.取得使用者路徑_含反斜線()
    If Dir(userProfilePath & "dropbox\《重編國語辭典修訂本》資料庫.mdb") <> "" Then
        cntStr = userProfilePath & "dropbox\《重編國語辭典修訂本》資料庫.mdb"
    ElseIf Dir(userProfilePath & "Documents\《重編國語辭典修訂本》資料庫.mdb") <> "" Then
        cntStr = userProfilePath & "Documents\《重編國語辭典修訂本》資料庫.mdb"
    ElseIf fso.fileexists("D:\@@@華語文工具及資料@@@\Macros\《重編國語辭典修訂本》資料庫.mdb") Then
                cntStr = "D:\@@@華語文工具及資料@@@\Macros\《重編國語辭典修訂本》資料庫.mdb"
    ElseIf Dir("c:\@@@華語文工具及資料@@@\Macros\《重編國語辭典修訂本》資料庫.mdb") <> "" Then
        cntStr = "c:\@@@華語文工具及資料@@@\Macros\《重編國語辭典修訂本》資料庫.mdb"
    ElseIf fso.fileexists("G:\我的雲端硬碟\DATA\h\@@@華語文工具及資料@@@\Macros\《重編國語辭典修訂本》資料庫.mdb") Then
            cntStr = "G:\我的雲端硬碟\DATA\h\@@@華語文工具及資料@@@\Macros\《重編國語辭典修訂本》資料庫.mdb"
    End If
getdb_重編國語辭典修訂本_資料庫fullName = cntStr
'Exit Function
'
'eH:
'Select Case Err.Number
'    Case 52 '不正確的名稱或數目
'
'    Case Else
'        MsgBox Err.Number & Err.Description
'        'Resume
'End Select
End Function

Function getdb_查字FullName() As String
If fso.fileexists(userProfilePath & "dropbox\查字.mdb") Then
    getdb_查字FullName = userProfilePath & "dropbox\查字.mdb"
ElseIf fso.fileexists("H:\我的雲端硬碟\私人\千慮一得齋(C槽版)\書籍資料\圖書管理附件\查字.mdb") Then
    getdb_查字FullName = "H:\我的雲端硬碟\私人\千慮一得齋(C槽版)\書籍資料\圖書管理附件\查字.mdb"
ElseIf fso.fileexists("D:\千慮一得齋\書籍資料\圖書管理附件\查字.mdb") Then
    getdb_查字FullName = "D:\千慮一得齋\書籍資料\圖書管理附件\查字.mdb"
ElseIf Dir("c:\@@@華語文工具及資料@@@\Macros\查字.mdb") <> "" Then
    getdb_查字FullName = "c:\@@@華語文工具及資料@@@\Macros\查字.mdb"
'Else
'    getdb_查字FullName = ""
End If
End Function
Function getdb_查字forinputFullName()
If fso.fileexists(userProfilePath & "dropbox\查字forinput.mdb") Then
    getdb_查字forinputFullName = userProfilePath & "dropbox\查字forinput.mdb"
ElseIf fso.fileexists("H:\我的雲端硬碟\私人\千慮一得齋(C槽版)\書籍資料\圖書管理附件\查字forinput.mdb") Then
    getdb_查字forinputFullName = "H:\我的雲端硬碟\私人\千慮一得齋(C槽版)\書籍資料\圖書管理附件\查字forinput.mdb"
ElseIf fso.fileexists("D:\千慮一得齋\書籍資料\圖書管理附件\查字forinput.mdb") Then
    getdb_查字forinputFullName = "D:\千慮一得齋\書籍資料\圖書管理附件\查字forinput.mdb"
ElseIf Dir("c:\@@@華語文工具及資料@@@\Macros\查字forinput.mdb") <> "" Then
    getdb_查字forinputFullName = "c:\@@@華語文工具及資料@@@\Macros\查字forinput.mdb"
End If
End Function

Private Sub Class_Initialize()
Set fso = CreateObject("Scripting.FileSystemObject")
userProfilePath = SystemSetup.取得使用者路徑_含反斜線()
End Sub

Private Sub Class_Terminate()
Set fso = Nothing
End Sub
