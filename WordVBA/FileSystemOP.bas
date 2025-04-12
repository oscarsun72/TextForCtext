Attribute VB_Name = "FileSystemOP"
Option Explicit

Rem Copilot�j���� 20250409�G
'Copilot�j���ĦN���G�ڷQ�g�@��WordVBA�{���A�b���w����Ƨ����|�U���M���Ҧ���txt��A���䤤�]�t���w��������r���ɮסA�������ɦW�o�˪����G�b�@�ӷs��word��󤤪��U�q�]�@�q�@�ӥ��ɦW�^ �����ڧ����A�n�ܡH�ڦA�i����աC�P���P���@�n�L��������
Sub FindTxtFilesWithKeyword()
    Dim folderPath As String
    Dim keyword As String
    Dim fileName As String
    Dim fileContent As String
    Dim doc As Document
    Dim rng As Range
    Dim fso As Object
    Dim file As Object
    Dim d As Document
    
    Set d = ActiveDocument
    ' �]�w��Ƨ����|�M����r�b��1�q
    'folderPath = "C:\YourFolderPath\" ' �д������z����Ƨ����|
    folderPath = d.Range(d.Paragraphs(1).Range.start, d.Paragraphs(1).Range.End - 1).text
    
    'keyword = "���w��������r" ' �д������z������r�b��2�q
    keyword = d.Range(d.Paragraphs(2).Range.start, d.Paragraphs(2).Range.End - 1).text
    
    ' �إ߷s�� Word ���
    Set doc = Documents.Add

    ' �ϥ� FileSystemObject �s����Ƨ�
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        MsgBox "��Ƨ����s�b�G" & folderPath, vbExclamation
        Exit Sub
    End If

    For Each file In fso.GetFolder(folderPath).Files
        If LCase(fso.GetExtensionName(file.Name)) = "txt" Then
            fileName = file.path
            ' �}�� .txt �ɮר��ˬd���e
            fileContent = GetFileContent(fileName)
            If InStr(fileContent, keyword) > 0 Then
                ' �N�ɮצW�K�[�� Word ���
                Set rng = doc.content
                rng.Collapse wdCollapseEnd
                rng.InsertAfter fileName & vbCrLf
                rng.InsertParagraphAfter
                rng.Hyperlinks.Add rng, fileName
            End If
        End If
    Next file

    MsgBox "�����I�ɮצW�w�K�[�� Word ��󤤡C", vbInformation
End Sub

Function GetFileContent(filePath As String) As String
    Dim stream As Object
    Dim content As String

    ' �ϥ� ADODB.Stream ��Ū�� UTF-8 �s�X����r�ɮ�
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' �]�m����r�Ҧ�
    stream.Mode = 3 ' �]�m��Ū�g�Ҧ�
    stream.Charset = "UTF-8" ' ���w�s�X�� UTF-8
    stream.Open
    stream.LoadFromFile filePath
    content = stream.ReadText
    stream.Close

    GetFileContent = content
End Function
