Attribute VB_Name = "ClipBoard"
Option Explicit
Rem 20230408 Bing�j���ġG
#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function lstrcpy Lib "kernel32" Alias "lstrcpyW" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
#Else
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare PtrSafe Function lstrcpy Lib "kernel32" Alias "lstrcpyW" (ByVal lpString1 As Any, ByVal lpString2 Any) As Long
#End If

Public Const GHND = &H42
Public Const CF_UNICODETEXT = 13&

Public Sub SetClipboard(sUniText As String)
   #If VBA7 Then
       Dim iStrPtr As LongPtr
       Dim iLock As LongPtr
       Dim ILEN As LongPtr
   #Else
       Dim iStrPtr As Long
       Dim iLock As Long
       Dim ILEN As Long
   #End If
   
   OpenClipboard 0&
   
   ILEN = Len(sUniText)
   
   iStrPtr = GlobalAlloc(GHND, ILEN * 2 + 2)
   
   If iStrPtr Then
   
        iLock = GlobalLock(iStrPtr)
        
        lstrcpy iLock, StrPtr(sUniText)
        
        GlobalUnlock iStrPtr
        
        SetClipboardData CF_UNICODETEXT, iStrPtr
        
   End If
   
   CloseClipboard
   
End Sub

Public Function GetClipboard() As String

   #If VBA7 Then
       Dim iStrPtr As LongPtr
       Dim ILEN As LongPtr
       Dim iLock As LongPtr
   #Else
       Dim iStrPtr As Long
       Dim ILEN As Long
       Dim iLock As Long
   #End If
   
   Dim sUniText As String
   
   OpenClipboard 0&
   
   If IsClipboardFormatAvailable(CF_UNICODETEXT) Then
   
        iStrPtr = GetClipboardData(CF_UNICODETEXT)
        
        If iStrPtr Then
        
            iLock = GlobalLock(iStrPtr)
            
            ILEN = GlobalSize(iStrPtr)
            
            sUniText = String$(CLng(ILEN \ 2&) - 1&, vbNullChar)
            
            lstrcpy StrPtr(sUniText), iLock
            
            GlobalUnlock iStrPtr
            
        End If
        
        GetClipboard = sUniText
        
   End If
   
   CloseClipboard
   
End Function
' �z�i�H�N�W�����Ѫ��N�X�ƻs���߶K��z�� VBA ���ؤ��C�o�ǥN�X�w�q�F��Ө�ơGSetClipboard �M GetClipboard�C
' SetClipboard ��Ʊ����@�Ӧr�Ŧ�ѼơA�ñN��]�m���ŶK�O�����e�C
' GetClipboard ��ƪ�^�ŶK�O�����奻���e�C
' �z�i�H�b�z�� VBA �N�X���եγo��Ө�ƨӾާ@�ŶK�O�C
' �Ҧp�A�n�N�ŶK�O�]�m�� ��Hello World���A�z�i�H�ϥΥH�U�N�X�G
' SetClipboard "Hello World"
' �n����ŶK�O�����奻���e�A�z�i�H�ϥΥH�U�N�X�G
' Dim sText As String
' sText = GetClipboard()
' �Ʊ�o�ǫH����z�������U�C�p�G�z����L���D�A���H�ɧi�D�ڡC
' �p�G�z�Q�n�N�ϧΩΨ�L�D�奻�ƾڳ]�m���ŶK�O�����e�A�z�ݭn�ϥΤ��P���ŶK�O�榡�C�b�W�����Ѫ��N�X���A�ڭ̨ϥΤF `CF_UNICODETEXT` �`�q�ӫ��w�ŶK�O�榡�� Unicode �奻�C
'���ϧμƾڡA�z�i�H�ϥ� `CF_BITMAP`�B`CF_DIB` �� `CF_ENHMETAFILE` ���ŶK�O�榡�C�z�ݭn�ھڱz�n�]�m���ϧμƾ�������ܦX�A���ŶK�O�榡�C
'���~�A�z�ٻݭn�ק� `SetClipboard` ��ƥH�����ϧμƾڧ@���ѼơA�èϥά����� Windows API ��ƨӱN�ϧμƾڳ]�m���ŶK�O�����e�C
'�ѩ�o�A�Ψ������� Windows API �s�{�A�]���ڵL�k���ѧ��㪺�N�X�ܨҡC���O�A�z�i�H�b���W�j��������ƨ��A�Ѧp��ϥ� VBA �M Windows API �Ӿާ@�ŶK�O�����ϧμƾڡC
'�O���A�ϥ� VBA �ާ@�ŶK�O�����奻�ƾڬ۹�e����{�C�o�O�]�� VBA �����N���ѤF��r�Ŧꪺ�}�n����A�ӥB Windows API ���]���\�h�Ω�ާ@�奻�ƾڪ���ơC
''����L�������ƾڡA�p�ϧΡB��󵥡A�ާ@�_�ӷ|�۹�����@�ǡA�]���z�ݭn��h�a�ϥ� Windows API ��ơA�åB�ݭn��h�a�A�ѰŶK�O�榡�M�ƾ������C
''���O�A�u�n�z�@�N�ǲߨôx���������ѡA�ϥ� VBA �ާ@�ŶK�O������L�����ƾڤ]�O�����i�檺�C
