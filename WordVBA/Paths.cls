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
Function getdb_���s��y���׭q��_��ƮwfullName()
Dim cntStr As String, pth As String
'On Error GoTo eH
pth = SystemSetup.���o�ϥΪ̸��|_�t�ϱ׽u()
    If Dir(pth & "dropbox\�m���s��y���׭q���n��Ʈw.mdb") <> "" Then
        cntStr = pth & "dropbox\�m���s��y���׭q���n��Ʈw.mdb"
    ElseIf Dir("C:\Users\ssz3\Documents\�m���s��y���׭q���n��Ʈw.mdb") <> "" Then
        cntStr = "C:\Users\ssz3\Documents\�m���s��y���׭q���n��Ʈw.mdb"
    ElseIf fso.fileexists("D:\@@@�ػy��u��θ��@@@\Macros\�m���s��y���׭q���n��Ʈw.mdb") Then
                cntStr = "D:\@@@�ػy��u��θ��@@@\Macros\�m���s��y���׭q���n��Ʈw.mdb"
    ElseIf Dir("c:\@@@�ػy��u��θ��@@@\Macros\�m���s��y���׭q���n��Ʈw.mdb") <> "" Then
        cntStr = "c:\@@@�ػy��u��θ��@@@\Macros\�m���s��y���׭q���n��Ʈw.mdb"
    ElseIf fso.fileexists("G:\�ڪ����ݵw��\DATA\h\@@@�ػy��u��θ��@@@\Macros\�m���s��y���׭q���n��Ʈw.mdb") Then
            cntStr = "G:\�ڪ����ݵw��\DATA\h\@@@�ػy��u��θ��@@@\Macros\�m���s��y���׭q���n��Ʈw.mdb"
    End If
getdb_���s��y���׭q��_��ƮwfullName = cntStr
'Exit Function
'
'eH:
'Select Case Err.Number
'    Case 52 '�����T���W�٩μƥ�
'
'    Case Else
'        MsgBox Err.Number & Err.Description
'        'Resume
'End Select
End Function

Function getdb_�d�rFullName() As String
Dim pth As String
pth = SystemSetup.���o�ϥΪ̸��|_�t�ϱ׽u()
If fso.fileexists(pth & "dropbox\�d�r.mdb") Then
    getdb_�d�rFullName = pth & "dropbox\�d�r.mdb"
ElseIf fso.fileexists("H:\�ڪ����ݵw��\�p�H\�d�{�@�o�N(C�Ѫ�)\���y���\�ϮѺ޲z����\�d�r.mdb") Then
    getdb_�d�rFullName = "H:\�ڪ����ݵw��\�p�H\�d�{�@�o�N(C�Ѫ�)\���y���\�ϮѺ޲z����\�d�r.mdb"
ElseIf fso.fileexists("D:\�d�{�@�o�N\���y���\�ϮѺ޲z����\�d�r.mdb") Then
    getdb_�d�rFullName = "D:\�d�{�@�o�N\���y���\�ϮѺ޲z����\�d�r.mdb"
ElseIf Dir("c:\@@@�ػy��u��θ��@@@\Macros\�d�r.mdb") <> "" Then
    getdb_�d�rFullName = "c:\@@@�ػy��u��θ��@@@\Macros\�d�r.mdb"
'Else
'    getdb_�d�rFullName = ""
End If
End Function
Function getdb_�d�rforinputFullName()
Dim pth As String
pth = SystemSetup.���o�ϥΪ̸��|_�t�ϱ׽u()
If fso.fileexists(pth & "dropbox\�d�rforinput.mdb") Then
    getdb_�d�rforinputFullName = pth & "dropbox\�d�rforinput.mdb"
ElseIf fso.fileexists("H:\�ڪ����ݵw��\�p�H\�d�{�@�o�N(C�Ѫ�)\���y���\�ϮѺ޲z����\�d�rforinput.mdb") Then
    getdb_�d�rforinputFullName = "H:\�ڪ����ݵw��\�p�H\�d�{�@�o�N(C�Ѫ�)\���y���\�ϮѺ޲z����\�d�rforinput.mdb"
ElseIf fso.fileexists("D:\�d�{�@�o�N\���y���\�ϮѺ޲z����\�d�rforinput.mdb") Then
    getdb_�d�rforinputFullName = "D:\�d�{�@�o�N\���y���\�ϮѺ޲z����\�d�rforinput.mdb"
ElseIf Dir("c:\@@@�ػy��u��θ��@@@\Macros\�d�rforinput.mdb") <> "" Then
    getdb_�d�rforinputFullName = "c:\@@@�ػy��u��θ��@@@\Macros\�d�rforinput.mdb"
End If
End Function

Private Sub Class_Initialize()
Set fso = CreateObject("Scripting.FileSystemObject")
End Sub

Private Sub Class_Terminate()
Set fso = Nothing
End Sub