VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Links"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub insertLinktoDict(dictID As Byte)
Dim db As New dBase, rst As New ADODB.Recordset ', cnt As New ADODB.Connection
Dim x, frmDict As New Form_DictsURL, rng As Range, lnk As String, lnks As New Links
Dim msdo As New MSForms.DataObject, sty As word.Style
Dim haveZhuYin As Boolean
If Selection.Type = wdSelectionIP Then Selection.MoveRight wdCharacter, 1, wdExtend
Set rng = Selection.Range
x = Selection.Text
1
frmDict.getDictRecS CStr(x), rst, dictID
'frmDict.getDictHydzdRecS CStr(x), rst
If rst.RecordCount > 0 Then
    Dim ur As UndoRecord
    Set ur = SystemSetup.stopUndo("�K�a��")
        
        lnk = rst.Fields(2).Value
        trimLinks lnk
        If rst.RecordCount = 1 Then
            If Not VBA.IsNull(rst.Fields(0).Value) Then
                GoSub list
            Else
3              Set sty = rng.Style
                Selection.Hyperlinks.Add Selection.Range, lnk, , , , "_blank" ' if only with Selection not Selection.Range will fall
                setStylewithHyperlinkMark sty, rng
                rng.Select
            End If
        Else
                GoSub list
                If x = "" Then GoTo 3 'if  ZhuYin is none.
        End If
    
    msdo.Clear
    msdo.SetText lnk
    msdo.PutInClipboard
    SystemSetup.contiUndo ur
    Set ur = Nothing:  Set msdo = Nothing

Else
    Selection.Copy
    db.setURLControlValue frmDict.add1URLTo1Dict(CStr(x), dictID)
    rst.Requery
    If rst.RecordCount > 0 Then GoTo 1
End If
Exit Sub

list:
        Do Until rst.EOF
            x = ""
            If VBA.IsNull(rst.Fields(0).Value) Then
'                x = rst.Fields(1).Value '���q
            Else
                If Not haveZhuYin Then
                    rng.Collapse wdCollapseEnd
                    If rng.Style <> "�K�a��" Then
                        rng.InsertAfter "�]�^"
                        rng.Style = "�K�a��"
                        rng.SetRange rng.End - 1, rng.End - 1
                    End If
                    haveZhuYin = True
                End If
                x = rst.Fields(0).Value '�`��
            End If
            GoSub typeTexts
            If haveZhuYin Then
                rng.Next.InsertBefore "�F"
                rng.SetRange rng.Next.End, rng.Next.End
            End If
            rst.MoveNext
        Loop
        If rng.Previous = "�F" Then rng.Previous.Delete

Return

typeTexts:
        Dim ay, i As Byte
        ay = Array("��", ChrW(20008), "�@", " ", "�]�S���^", "�S�� ", "�]Ū���^", "Ū�� ", "�]�y���^", "�y�� ", _
            "(�@)", "", "(�G)", "", "(�T)", "", "(�|)", "", "(��)", "", "(��)", "", "�^", "", "�]", "")

        If x = "" Then Return 'Exit Sub
'        X = Mid(X, 1, Len(X) - 1)
        For i = 0 To UBound(ay)
            x = Replace(x, ay(i), ay(i + 1))
            i = i + 1
        Next i
'        If sT <> wdSelectionIP Then
'            rng.SetRange Selection.End, Selection.End
'        End If
'        rng.SetRange rng.End - 1, rng.End - 1
        rng.InsertAfter x 'insert ZhuYin
        For Each x In rng.Characters 'format ZhuYin
            If InStr("������", x) Then
                x.Style = "�n��"
            ElseIf InStr("��", x) Then
                x.Font.Name = "�з���"
            End If
        Next x
        x = rst.Fields(2).Value 'URL  'frmDict.get1URLfor1(words)
'        If VBA.IsNull(x) Then
'            If sT = wdSelectionIP Then Selection.Characters(1).Previous.Copy
'rePt:
'            If rst.RecordCount = 1 Then AppActivate "google chrome" 'https://docs.microsoft.com/zh-tw/sql/ado/reference/ado-api/absoluteposition-property-ado?view=sql-server-ver15
'            x = InputBox("plz putin the url", , IIf(VBA.IsNull(rst.Fields(0).Value), "", rst.Fields(0).Value)) 'frmDict.add1URLTo1��y���(words)
'            If x = "" Then GoTo endS
'            If Left(x, 4) <> "http" Then GoTo rePt
'            rst.Fields(2).Value = x
'            rst.Update
'        End If
        Set sty = rng.Style
        rng.Hyperlinks.Add rng, lnks.trimLinks(VBA.CStr(x)), , , , "_blank"
        setStylewithHyperlinkMark sty, rng
        rng.Collapse wdCollapseEnd
        'rng.Next.InsertBefore "�A"
        'rng.SetRange rng.End + 2, rng.End + 2
'        rng.SetRange rng.End + 1, rng.End + 1
Return

End Sub
Sub setStylewithHyperlinkMark(sty As word.Style, rng As word.Range)
'If styName <> "����" Then
If sty.Font.Bold Then
    rng.Style = sty
'    rng.Font.Color = VBA.RGB(0, 0, 255)
'    rng.Font.Underline = word.wdUnderlineSingle
End If
End Sub

Sub insertLinktoHydzd()
insertLinktoDict 3
End Sub

Sub insertLinktoHydcd()
insertLinktoDict 4
End Sub

Function trimLinks(lnk As String)
Dim ay, e
ay = Array(Chr(9))
lnk = VBA.Trim(lnk)
If VBA.Left(lnk, 1) = "#" And VBA.Right(lnk, 1) = "#" Then lnk = Mid(lnk, 2, Len(lnk) - 2)
For Each e In ay
    lnk = VBA.Replace(lnk, e, "")
Next e
trimLinks = lnk
End Function

Function trimLinks_http_Dicts_toAddZhuYin_RevisedMoeEdu(url As String, ZhuYin As Variant) As String
Dim i As Byte, ay, e '��y���
ay = Array("&", "#")
    For Each e In ay
        If InStr(url, e) Then url = Left(url, InStr(url, e) - 1)
    Next
trimLinks_http_Dicts_toAddZhuYin_RevisedMoeEdu = url

ay = Array("�@", "�G", "�T", "�|", "��", "��", "�C", "�K", "�E", "�Q", "�Q�@", "�Q�G")
If Not IsNull(ZhuYin) Then
    If InStr(ZhuYin, "(") Then
        For Each e In ay
            i = i + 1
            If InStr(ZhuYin, e) Then Exit For
        Next
        If i <= UBound(ay) + 1 Then trimLinks_http_Dicts_toAddZhuYin_RevisedMoeEdu = url & "#order" & i
        Beep
    Else
        trimLinks_http_Dicts_toAddZhuYin_RevisedMoeEdu = url
    End If
End If
End Function

Function chkLinks_http_Dicts_toAddZhuYin(ByRef url As String, Optional w As String, Optional dictID, Optional ByRef id As Long, Optional ZhuYin As Variant) As Boolean
If Not checkURL_single(url) Then Exit Function

Dim rstUrl As New ADODB.Recordset, dictIDcrct As String, strAyUrlChk, i As Byte, cnt As New ADODB.Connection, db As New dBase, ur As UndoRecord
db.cnt�d�r cnt
chkUrl:
    rstUrl.Open "select url from dicturl where url=""" & url & """ and strcomp(word,""" & w & """)=0", cnt, adOpenKeyset
'    Select Case dictID
'        Case 1, 2 'one URL only one character
            If rstUrl.RecordCount > 0 Then Exit Function  'avoid repeated rec
'    End Select
    If rstUrl.State <> adStateClosed Then rstUrl.Close
    rstUrl.Open "select chkurlstr from dictsNameList where dictid=" & dictID, cnt, adOpenKeyset
    If rstUrl.RecordCount = 0 Then Exit Function
    strAyUrlChk = VBA.Split(rstUrl.Fields(0).Value, " ")
    For i = 0 To UBound(strAyUrlChk)
    'https://docs.microsoft.com/zh-tw/office/vba/Language/Reference/User-Interface-Help/this-array-is-fixed-or-temporarily-locked-error-10
        If VBA.InStr(url, strAyUrlChk(i)) = 0 Then
            dictIDcrct = InputBox("���w���r���妳�~�A�п�J���T��ID�G��=1�A��=2�A�r=3�A��=4" & vbCr & vbCr _
                    & "���ˬd�m��y���n�����}�O�_���~�C�O�_�O�H" & vbCr & vbCr & _
                    "�uhttp://dict.revised.moe.edu.tw/cgi-bin/cbdic/gsweb.cgi?o=dcbdic&searchid=�v�}�Y�����}", , dictID)
            If dictIDcrct = "" Then Exit Function
            If Left(dictIDcrct, 4) = "http" Then
                url = dictIDcrct
                rstUrl.Close
                trimLinks_http_Dicts_toAddZhuYin_RevisedMoeEdu url, ZhuYin
                GoTo chkUrl
            End If
            dictID = CLng(dictIDcrct)
            If rstUrl.State <> adStateClosed Then rstUrl.Close
            GoTo chkUrl
        Else
            'If Combo4.Value <> dictID Then Combo4 = dictID
'            Combo4 = dictID
            chkLinks_http_Dicts_toAddZhuYin = True
            If dictID = 1 Then
                dictIDcrct = url
                If InStr(dictIDcrct, "&") Then dictIDcrct = Left(dictIDcrct, InStr(dictIDcrct, "&") - 1)
                If InStr(dictIDcrct, "#") Then dictIDcrct = Left(dictIDcrct, InStr(dictIDcrct, "#") - 1)
                If rstUrl.State <> adStateClosed Then rstUrl.Close
                If cnt.State <> adStateClosed Then cnt.Close
                db.cnt_���s��y���׭q��_��Ʈw cnt
                id = CLng(Mid(dictIDcrct, InStr(dictIDcrct, "ID=") + 3))
                rstUrl.Open "select * from " & db.���s��y���׭q��_��Ʈw_�`��Name & " where ID=" & id, cnt, adOpenKeyset
                If rstUrl.RecordCount > 0 Then
                    chkLinks_http_Dicts_toAddZhuYin = False
                    MsgBox "url is wrong!", vbExclamation
                    Set ur = word.Application.UndoRecord
                    SystemSetup.contiUndo ur
                    Set ur = Nothing
                    rstUrl.Close
                    Exit Function
                End If
                rstUrl.Close
            End If
            Exit Function
        End If
    Next i
End Function

Function checkURL_single(url As String) As Boolean
Dim u As String
If Left(url, 4) <> "http" Then Exit Function
If VBA.InStr(url, " ") Then Exit Function
If url = "" Then Exit Function
u = VBA.Replace(url, "http", "", 1, 1)
If VBA.InStr(u, "http") Then Exit Function
checkURL_single = True
End Function
Function updateURL��Ǥj�v(d As word.Document)
Dim lnk As String, hpLnk As word.Hyperlink
If d.Hyperlinks.Count = 0 Then Exit Function
For Each hpLnk In d.Hyperlinks
    lnk = hpLnk.Address
    If VBA.InStr(lnk, "guoxuedashi.com") = 0 Then
        hpLnk.Address = VBA.Replace(lnk, ".com", ".net", 1, 1)
    End If
Next hpLnk
End Function

Function updateURL��y���(d As word.Document)
On Error GoTo eH
Dim db As New dBase, cnt As New ADODB.Connection, lnk As String, hpLnk As word.Hyperlink, rst As New ADODB.Recordset
Dim id As Long, chm As String, w As String
'Set d = ActiveDocument
If d.Hyperlinks.Count = 0 Then Exit Function
db.cnt_���s��y���׭q��_��Ʈw cnt
For Each hpLnk In d.Hyperlinks
    lnk = hpLnk.Address
    If VBA.InStr(lnk, "/dictView.jsp?ID=") = 0 Then
        If VBA.InStr(lnk, "dict.revised.moe.edu.tw") Then
1
            rst.Open "select url,urlv5 from " & db.���s��y���׭q��_��Ʈw_�`��Name & " where instr(urlv5 ,""" & lnk & """)>0", cnt, adOpenKeyset
            If rst.RecordCount = 1 Then
                hpLnk.Address = Me.trimLinks(rst.Fields(0).Value)
            ElseIf rst.RecordCount > 1 Then
                hpLnk.Range.Select
                If InStr(Selection, "������") > 0 And InStr(lnk, "9168") > 0 Then
                    hpLnk.Address = "https://dict.revised.moe.edu.tw/dictView.jsp?ID=11765#order3"
                Else
                    SystemSetup.playSound 2
'                    Stop
                    If Selection <> "��������" Then
                        MsgBox "url is duplicated ! plz check out", vbExclamation
                        GoTo endS
                    End If
                    hpLnk.Address = "https://dict.revised.moe.edu.tw/dictView.jsp?ID=9168#order2"
                End If
            ElseIf rst.RecordCount = 0 Then
                hpLnk.Range.Select
                'SystemSetup.playSound 12
                'MsgBox "not found!", vbExclamation
                Beep
                id = VBA.CLng(Mid(lnk, VBA.InStrRev(lnk, "=") + 2))
                chm = SystemSetup.getChrome()
                Shell chm & " "
                Shell chm & " https://dict.revised.moe.edu.tw/dictView.jsp?ID=" & id
                'Stop
                chm = "https://dict.revised.moe.edu.tw/dictView.jsp?ID=" & id 'InputBox("url plz")
                If chm = "" Then GoTo endS
                w = VBA.Replace(VBA.Trim(InputBox("Words plz")), Chr(9), "")
                If w = "" Then GoTo endS
                If VBA.InStr(w, "?") > 0 Then
                    SystemSetup.playSound 3
                    Stop
                    GoTo endS
                End If
                chm = Me.trimLinks(chm)
                If Me.checkURL_single(chm) Then
                    If Me.chkLinks_http_Dicts_toAddZhuYin(chm, w, 1, id, Selection.Text) Then
2
                        rst.Close
                         rst.Open "select url,urlv5,id,�`���@��,�r���W from " & db.���s��y���׭q��_��Ʈw_�`��Name & _
                            " where strcomp(�r���W,""" & w & """)=0 and strcomp(�`���@�� ,""" & VBA.StrConv(VBA.Replace(Selection.Text, ChrW(20008), "��"), vbWide) & """)=0", cnt, adOpenKeyset, adLockOptimistic
                        If rst.RecordCount = 1 Then
                            If VBA.IsNull(rst.Fields(0).Value) Then rst.Fields(0).Value = chm
                            If VBA.IsNull(rst.Fields(1).Value) Then rst.Fields(1).Value = lnk
                            If VBA.IsNull(rst.Fields(2).Value) Then rst.Fields(2).Value = id
                            rst.Update
                            hpLnk.Address = chm
                        Else
                            SystemSetup.playSound 3
                            'Stop
                            MsgBox "do yourself!", vbExclamation
                            GoTo endS
                        End If
                    Else
'                        Debug.Print chm
'                        Stop
                        GoTo 2
                    End If
                End If
            End If
            rst.Close
            
'            Beep
        End If
    End If
Next hpLnk
endS:
If rst.State <> adStateClosed Then rst.Close
If cnt.State <> adStateClosed Then cnt.Close
Set rst = Nothing: Set cnt = Nothing
Exit Function
eH:
Select Case Err.Number
    Case Else
        MsgBox Err.Number & Err.Description
        'resume
End Select
End Function
