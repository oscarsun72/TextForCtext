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
Function getdb_重編國語辭典修訂本_資料庫fullName() As String
Dim cntStr As String
'On Error GoTo eH
UserProfilePath = SystemSetup.取得使用者路徑_含反斜線()
    If Dir(UserProfilePath & "dropbox\《重編國語辭典修訂本》資料庫.mdb") <> "" Then
        cntStr = UserProfilePath & "dropbox\《重編國語辭典修訂本》資料庫.mdb"
    ElseIf Dir(Replace(UserProfilePath, "C:", "A:") & "dropbox\查字.mdb") <> "" Then '虛擬電腦用
        cntStr = Replace(UserProfilePath, "C:", "A:") & "dropbox\《重編國語辭典修訂本》資料庫.mdb"
    ElseIf Dir(UserProfilePath & "Documents\《重編國語辭典修訂本》資料庫.mdb") <> "" Then
        cntStr = UserProfilePath & "Documents\《重編國語辭典修訂本》資料庫.mdb"
    ElseIf FsO.fileexists("D:\@@@華語文工具及資料@@@\Macros\《重編國語辭典修訂本》資料庫.mdb") Then
                cntStr = "D:\@@@華語文工具及資料@@@\Macros\《重編國語辭典修訂本》資料庫.mdb"
    ElseIf Dir("c:\@@@華語文工具及資料@@@\Macros\《重編國語辭典修訂本》資料庫.mdb") <> "" Then
        cntStr = "c:\@@@華語文工具及資料@@@\Macros\《重編國語辭典修訂本》資料庫.mdb"
    ElseIf FsO.fileexists("G:\我的雲端硬碟\DATA\h\@@@華語文工具及資料@@@\Macros\《重編國語辭典修訂本》資料庫.mdb") Then
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
If FsO.fileexists(UserProfilePath & "dropbox\查字.mdb") Then
    getdb_查字FullName = UserProfilePath & "dropbox\查字.mdb"
ElseIf FsO.fileexists(Replace(UserProfilePath, "C:", "A:") & "dropbox\查字.mdb") Then '虛擬電腦用
    getdb_查字FullName = Replace(UserProfilePath, "C:", "A:") & "dropbox\查字.mdb"
ElseIf FsO.fileexists("H:\我的雲端硬碟\私人\千慮一得齋(C槽版)\書籍資料\圖書管理附件\查字.mdb") Then
    getdb_查字FullName = "H:\我的雲端硬碟\私人\千慮一得齋(C槽版)\書籍資料\圖書管理附件\查字.mdb"
ElseIf FsO.fileexists("D:\千慮一得齋\書籍資料\圖書管理附件\查字.mdb") Then
    getdb_查字FullName = "D:\千慮一得齋\書籍資料\圖書管理附件\查字.mdb"
ElseIf Dir("c:\@@@華語文工具及資料@@@\Macros\查字.mdb") <> "" Then
    getdb_查字FullName = "c:\@@@華語文工具及資料@@@\Macros\查字.mdb"
'Else
'    getdb_查字FullName = ""
End If
End Function
Function getdb_查字forinputFullName() As String
If FsO.fileexists(UserProfilePath & "dropbox\查字forinput.mdb") Then
    getdb_查字forinputFullName = UserProfilePath & "dropbox\查字forinput.mdb"
ElseIf FsO.fileexists(Replace(UserProfilePath, "C:", "A:") & "dropbox\查字forinput.mdb") Then '虛擬電腦用
    getdb_查字forinputFullName = Replace(UserProfilePath, "C:", "A:") & "dropbox\查字forinput.mdb"
ElseIf FsO.fileexists("H:\我的雲端硬碟\私人\千慮一得齋(C槽版)\書籍資料\圖書管理附件\查字forinput.mdb") Then
    getdb_查字forinputFullName = "H:\我的雲端硬碟\私人\千慮一得齋(C槽版)\書籍資料\圖書管理附件\查字forinput.mdb"
ElseIf FsO.fileexists("D:\千慮一得齋\書籍資料\圖書管理附件\查字forinput.mdb") Then
    getdb_查字forinputFullName = "D:\千慮一得齋\書籍資料\圖書管理附件\查字forinput.mdb"
ElseIf Dir("c:\@@@華語文工具及資料@@@\Macros\查字forinput.mdb") <> "" Then
    getdb_查字forinputFullName = "c:\@@@華語文工具及資料@@@\Macros\查字forinput.mdb"
End If
End Function

Function getdb_開發_千慮一得齋FullName() As String
If FsO.fileexists("D:\千慮一得齋\書籍資料\開發_千慮一得齋.mdb") Then _
    getdb_開發_千慮一得齋FullName = "D:\千慮一得齋\書籍資料\開發_千慮一得齋.mdb"

End Function
Private Sub Class_Initialize()
Set FsO = CreateObject("Scripting.FileSystemObject")
UserProfilePath = SystemSetup.取得使用者路徑_含反斜線()
End Sub

Function CombineFullName(folderPath As String, fileName As String) As String
Dim fullPath As String, fs, f
If right(folderPath, 1) = "\" Then
    fullPath = folderPath + fileName
Else
    fullPath = folderPath + "\" + fileName
End If
CombineFullName = fullPath
End Function

Private Sub Class_Terminate()
Set FsO = Nothing
End Sub


