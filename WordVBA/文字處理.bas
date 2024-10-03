Attribute VB_Name = "¤å¦r³B²z"
Option Explicit
Dim punctuationStr As String ' ¼ÐÂI²Å¸¹¦r¦ê
Dim rst As Recordset, d As Object
Dim db As Database 'set db=CurrentDb _
¥u¯à¦b¤w¶}±Ò¤§Access¤¤°Ñ·Ó¤@¦¸ , ¤G¦¸¥H¤Wªº°Ñ·Ó _
,¶·¥HSet db = DBEngine.Workspaces(0).OpenDatabase _
    ("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")!ªº§Î¦¡°Ñ·Ó! _
    °Ñ¦Ò: _
    Dim dbsCurrent As Database, dbsContacts As Database'¥Ñ CurrentDb ªº½u¤W»¡©ú½Æ»s _
    Set dbsCurrent = CurrentDb _
    Set dbsContacts = DBEngine.Workspaces(0).OpenDatabase("Contacts.mdb")

Rem ¼ÐÂI²Å¸¹¦r¦ê
Public Static Property Get PunctuationString() As String
If punctuationStr = "" Then _
    punctuationStr = "¡]¡^¡C¡u¡v¡y¡z[]¡i¡j¡e¡f¡m¡n¡q¡r-¡Ð"",  ¡G¡A¡F¡I¡H?" _
        & "¡B. :,;" _
        & "¡K¡K...!()-¡P¡E" & VBA.Chr(34) & VBA.Chr(-24153) & VBA.Chr(-24152) & VBA.Chr(-24155) & VBA.Chr(-24154) & VBA.ChrW(8218) '34¡GÂù¤Þ¸¹¡C¤j³°¼ÐÂI²Å¸¹¤W¤UÂù¤Þ¸¹¡B¤W¤U³æ¤Þ¼Æ¡B³r¸¹
PunctuationString = punctuationStr
End Property
'Public Static Property Let Punctionn(ByVal vNewValue As Variant)
'
'End Property


Function isNum(x As String) As Boolean
    If Len(x) > 1 Then Exit Function
    x = VBA.StrConv(x, vbNarrow)
    If x Like "[0-9]" Then isNum = True
End Function
Function isLetter(x As String) As Boolean
    If Len(x) > 1 Then Exit Function
    x = VBA.StrConv(x, vbNarrow)
    If x Like "[a-z]" Then isLetter = True
End Function

Sub ¦rÀW() '2002/11/10­nSub¤~¯à¦bWord¤¤°õ¦æ!
    On Error GoTo ¿ù»~³B²z
    Dim ch, wrong As Long
    'Dim chct As Long
    Dim StTime As Date, EndTime As Date
    'Dim x As Long, firstword As String '¶Ã½XÀË¬d!2002/11/13
    options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
    Set d = CreateObject("access.application")
    d.UserControl = True
    d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
    d.docmd.SelectObject blog.myaccess.acTable, "¦rÀWªí", True
    'd.Visible = True 'ÀË¬d¥Î
    Set db = d.CurrentDb '¤@©w­n¥[¡©d¡ª!!¼g¦¨¥H¤U¥ç¥i!
    '¥H¤W¥i¨Ö¦¨¤U¤G¦¡§Y¥i!¦ý¤£·|Åã¥Ü¦bÀç¹õ¤W,¥u¯à§@¹õ«á­pºâ¥Î!(¨£OpenCurrentDatabaseªº½u¤W»¡©ú)
    'Set db = d.DBEngine.OpenDatabase("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
    'Set db = d.DBEngine.Workspaces(0).OpenDatabase("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
    Set rst = db.OpenRecordset("¦rÀWªí", dbOpenDynaset)
    If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
    'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
    '    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
        db.Execute "DELETE * FROM ¦rÀWªí"
    End If
    StTime = VBA.Time
    With ActiveDocument
        For Each ch In .Characters '¦³¶Ã½X¦r®Éch·|¶Ç¦^"?"ÅÜ¦¨¤F¹Bºâ¥Î²Å¸¹
            wrong = wrong + 1 'ÀËµø¥Î!
    '        If wrong = 373 Then MsgBox "Check!!" 'ÀË¬d¥Î!
            If wrong Mod 27250 = 0 Then 'If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
                MsgBox "¦]¨t²Î­t²ü¹F¨ì·¥­­,½Ð°È¥²¤Á´«¦ÜAccess¥´¶}¸ê®Æªí«áÃö³¬,¦A¦^¨Ó«ö¤U½T©w«ö¶sÄ~Äò!!" _
                    , vbExclamation, "¡¹¨t²Î­«­n¸ê°T¡¹"
    '        ElseIf wrong = 49761 Then
    '            MsgBox "½ÐÀË¬d!!"
            End If
    '        If wrong Mod 1000 = 0 Then Debug.Print wrong
    '        Debug.Print ch & vbCr & "--------"
            '´«¦æ¦r¤¸¡B´_¦ì¦r¤¸¤£­p!
    '        If VBA.Right(ch, 1) <> vba.Chr(10) Or vba.Left(ch, 1) <> vba.Chr(13) Then
            Select Case Asc(ch)
                Case Is <> 13, 10
            With rst
11              .FindFirst "¦r·J like '" & ch & "'"
12              If .NoMatch Then
                    .AddNew
                    rst("¦r·J") = ch
                    rst("¦¸¼Æ") = 1
                    rst("Asc") = Asc(ch)
                    rst("AscW") = AscW(ch)
        '            On Error GoTo ¦¸¼Æ
                    .Update
                Else '·í¦³¶Ã½X¦r®É,·|¦¨¬°¤ñ¸û¹Bºâ¤¸"?"(Asc(ch)=63),«h¥i¯à¦b¤å¥ó¤¤²Ä¤@¦¸¥X²{ªº¦r·|»~¼W¦¸¼Æ
                    '¦¹¥~¦p"Åb"¦rµ¥(¦bWord¤¤´¡¤J¡÷²Å¸¹¤º³Ì«á¤@¨Ç)¦r,¥ç·|»P¦P§Î¦r¦P¦r¤¸½X(Asc), _
                    ¦ý¦b²Å¸¹ªí¤¤«o¦³¤£¦P¦ì¸m,¥Nªí¤£¦P¦r!¦b²Î­p®É,¨t²Î¥ç·|»~ºâ¦b¤@°_! _
                    ³oÂIÁÙ¶·­n§JªA!2002/11/13´ú¸Õ®É,¦³®É¤S·|¤À¶}!(¦ýAsc«h¬Û¦P!)
    '                If .AbsolutePosition < 1 And ch Like "?" And Not rst("¦r·J") = "?" Then
    '                    'If x = 1 Then MsgBox "¦³¶Ã½X¦r,¦¸¼Æ±N¥[¤J²Ä¤@­Ó¥X²{ªº¦r¤¤!!"
    '                    MsgBox "¦³¶Ã½X¦r,¦¸¼Æ±N¥[¤J²Ä¤@­Ó¥X²{ªº¦r¤¤!!"
    '                    AppActivate "Microsoft Word"
    '                    Selection.Collapse
    '                    Selection.SetRange wrong + ActiveDocument.Paragraphs.Count / 2, wrong + 1 '±N¸Ó¶Ã½X¦r¿ï¨ú
    '                    x = x + 1
    '                End If
                    .edit
                    rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                    .Update
                End If
            End With
            End Select
    '        chct = .Characters.Count
    '        chct = Selection.StoryLength
    '        instr(1+
    '        .Select
retry:      Next ch
    '    rst.Requery
    '    rst.MoveFirst
    '    If x > 0 Then
    '        firstword = "¡·¡·¶Ã½X¦r¥[¤J²Ä¤@¦r:¡u" & rst("¦r·J") & "¡v¤¤¦@¦³" & x & "¦¸!!"
    '    Else
    '        firstword = "¡¹©ñ¤ß§a!¶Ã½X¦r¥ç²Î­p¥¿½T!!¡¹"
    '    End If
        EndTime = VBA.Time
        AppActivate "Microsoft word"
        MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
            & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
            & vbCr & "¦r¤¸¼Æ=" & .Characters.Count & vbCr '_
    '        & firstword
    '    MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
    '        & vbCr & "¡°¯Ó®É:" & DateDiff("n", StTime, EndTime) & "¤ÀÄÁ¡°" _
    '        & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
    End With
    If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
    '    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
        AppActivate "Microsoft access"
    End If
    d.docmd.OpenTable "¦rÀWªí", , d.acReadOnly
    d.docmd.Maximize
    rst.Close: db.Close: Set d = Nothing
    options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
    End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
    Select Case Err.Number
        Case Is = 91, 3078 '°Ñ·Ó¤£¨ìDataBase¤ºª«¥ó®É
            MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
            'access.Application.Quit
            d.Quit
            End
    '        d.CurrentDb.Close
    '        Set db = DBEngine.Workspaces(0).OpenDatabase("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
    ''        Debug.Print Err.Description 'ÀË¬d¥Î!
    '        Resume
    '    Case Is = 3163 '´«¦æ¦r¤¸¡B´_¦ì¦r¤¸¤£­p!
    '        If VBA.Right(ch, 1) = vba.Chr(10) Then
    '            ch = vba.Left(ch, Len(ch) - 1)
    '        ElseIf vba.Left(ch, 1) = vba.Chr(13) Then
    '            ch = VBA.Right(ch, Len(ch) - 1) '©ÎIf Asc(ch)=13
    '        End If
    '        Resume 11
        Case Is = 93 '¬°[]µ¥¹Bºâ¦¡¯S®í¦r¤¸©Ò³]¤§¤ñ¸û¦¡
            rst.FindFirst "asc(¦r·J) = " & Asc(ch)
            Resume 12
    '    Case Is = -2147023170
    '        MsgBox Err.Number & ":" & Err.Description
    '        MsgBox Err.LastDllError & "." & Err.Source
    '        Set d = CreateObject("access.application")
    '        d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
    '        d.UserControl = True
    '        Resume
    '    Case Is = 462 '"»·ºÝ¦øªA¾¹¤£¦s¦b©ÎµLªk¨Ï¥Î"
    '        'd.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
    ''        Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
    '        Set db = d.CurrentDb
    '        Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
    '        Resume
        Case Else
            MsgBox Err.Number & ":" & Err.Description, vbExclamation
            Resume
    End Select
End Sub
Sub µüÀW() '2002/11/10
    On Error GoTo ¿ù»~³B²z
    Dim WD, wrong As Long
    Dim wrongmark As Integer ', wdct As Long
    Dim StTime As Date, EndTime As Date
    Dim hfspace As Long
    options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
    Set d = CreateObject("access.application")
    d.UserControl = True '¦pªG¬°False«hdb.close·|Ãö³¬¸ê®Æ®w!
    'd.UserControl = False
    d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
    d.docmd.SelectObject d.acTable, "µüÀWªí", True
    'd.Visible = True 'ÀË¬d¥ÎUserControl=True«h¦³¦¹¤Ï·|­P»~!
    Set db = d.CurrentDb
    Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
    If rst.RecordCount > 0 Then db.Execute "DELETE * FROM µüÀWªí"
    StTime = VBA.Time
    With ActiveDocument
        For Each WD In .words
            wrong = wrong + 1 'ÀËµø¥Î!
    '        If wrong Mod 1000 = 0 Then Debug.Print wrong
    '        Debug.Print wd & vbCr & "--------"
            If Len(WD) > 1 And VBA.Right(WD, 1) = " " Then
                hfspace = hfspace + 1 '­p¦¸
                GoTo retry '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
            End If
            rst.FindFirst "µü·J like '" & WD & "'"
            If rst.NoMatch Then
                rst.AddNew
                rst("µü·J") = WD
    '            On Error GoTo ¦¸¼Æ
                rst.Update
            Else
                rst.edit
                rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                rst.Update
            End If
    '        wrong = 1
    '        wdct = .Words.Count
    '        wdct = Selection.StoryLength
    '        instr(1+
    '        .Select
retry:      Next WD
    End With
    EndTime = VBA.Time
    AppActivate "Microsoft word"
    MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
        & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
        & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°"
    If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
    '    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
        AppActivate "Microsoft access"
    End If
    d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
    d.docmd.Maximize
    rst.Close: db.Close: Set d = Nothing
    options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
    End
¿ù»~³B²z:
    Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
        Case Is = 91, 3078
            MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
            'access.Application.Quit
            d.Quit
            End
        Case Else
            MsgBox Err.Number & ":" & Err.Description, vbExclamation
            Resume
    End Select
    '¦¸¼Æ:
    '    wrongmark = Err.Number
    ''    Err.Description = wd
    '    If wrongmark = 3022 Then '­«½Æ¤F
    ''        wrong = wrong + 1
    ''        rst.Seek "=", "µü·J"
    '        rst.FindFirst "µü·J like '" & wd & "'"
    '        rst.Edit
    '        rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
    '        rst.Update
    '        Resume retry
    '    Else
    '        MsgBox "¦³¿ù»~,½ÐÀË¬d!!" & Err.Description, vbExclamation
    '    End If
End Sub
Sub ¶i¶¥µüÀW() '2002/11/10­nSub¤~¯à¦bWord¤¤°õ¦æ!'2005/4/21¦¹ªk¦b¶]¤jÀÉ®×®É¤Ó¨S®Ä²v¤F!!¶]¤F3¤Ñ3©]300­¶ªº¤å¥óÀÉ¨ú1-3¦rµü¶]¤£§¹!
    On Error GoTo ¿ù»~³B²z
    Dim wrong As Long, phra As Long, phras As String, phralh As Byte
    Dim StTime As Date, EndTime As Date
    Dim hfspace As Long
    Dim Length As Byte 'As String
    Dim Dw As String, dwL As Long
    Length = InputBox("½Ð«ü©w¤ÀªRµü·J¤§¤W­­,³Ì¦h¤­­Ó¦r", , "5")
    If Length = "" Or Not IsNumeric(Length) Then End
    If CByte(Length) < 1 Or CByte(Length) > 5 Then End
    options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
    StTime = VBA.Time
    Set d = CreateObject("access.application")
    '©ÎSet d = CreateObject("Access.Application.9")
    d.UserControl = True
    d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
    d.docmd.SelectObject d.acTable, "µüÀWªí", True
    'd.Visible = True 'ÀË¬d¥Î
    Set db = d.CurrentDb
    Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
    If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
    'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
    '    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
        db.Execute "DELETE * FROM µüÀWªí"
    End If
    'With ActiveDocument
    With ActiveDocument
        Dw = .Content '¤å¥ó¤º®e
        dwL = Len(Dw) '¤å¥óªø«×
        .Close
    End With
        For phralh = 1 To Length 'CByte(length)
    '    For phralh = 1 To 5 '¼È©w³Ìªø¬°5­Ó¦rºc¦¨ªºµü(¤´¥i§ï§@ÅÜ¼Æ)
            For phra = 1 To dwL '.Characters.Count
                Select Case phralh
                    Case Is = 1
                        If Err.LastDllError <> 0 Then
                            MsgBox Err.LastDllError & ":" & Err.Description & "Err.Number:" & Err.Number
                            GoTo ¿ù»~³B²z
                        End If
    '                    phras = .Characters(phra)'¦¹ªk¤ÓºC!
                        phras = VBA.Mid(Dw, phra, 1)
                    Case Is = 2
                        If Err.LastDllError <> 0 Then
                            MsgBox Err.LastDllError & ":" & Err.Description & "Err.Number:" & Err.Number
                            GoTo ¿ù»~³B²z
                        End If
    '                    If phra + 1 <= .Characters.Count Then _
                        phras = .Characters(phra) & .Characters(phra + 1)
                        If phra + 1 <= dwL Then phras = VBA.Mid(Dw, phra, 2)
                    Case Is = 3
                        If Err.LastDllError <> 0 Then
                            MsgBox Err.LastDllError & ":" & Err.Description & "Err.Number:" & Err.Number
                            GoTo ¿ù»~³B²z
                        End If
    '                    If phra + 2 <= .Characters.Count Then _
                        phras = .Characters(phra) & .Characters(phra + 1) & _
                                .Characters(phra + 2)
                        If phra + 2 <= dwL Then phras = VBA.Mid(Dw, phra, 3)
                    Case Is = 4
                        On Error GoTo ¿ù»~³B²z
                        If Err.LastDllError <> 0 Then
                            MsgBox Err.LastDllError & ":" & Err.Description & "Err.Number:" & Err.Number
                            GoTo ¿ù»~³B²z
                        End If
    '                    If phra + 3 <= .Characters.Count Then _
                        phras = .Characters(phra) & .Characters(phra + 1) & _
                                .Characters(phra + 2) & .Characters(phra + 3)
                        If phra + 3 <= dwL Then phras = VBA.Mid(Dw, phra, 3)
                    Case Is = 5
                        On Error GoTo ¿ù»~³B²z
                        If Err.LastDllError <> 0 Then
                            MsgBox Err.LastDllError & ":" & Err.Description & "Err.Number:" & Err.Number
                            GoTo ¿ù»~³B²z
                        End If
    '                    If phra + 4 <= .Characters.Count Then _
                        phras = .Characters(phra) & .Characters(phra + 1) & _
                                .Characters(phra + 2) & .Characters(phra + 3) & _
                                .Characters(phra + 4)
                        If phra + 4 <= dwL Then phras = VBA.Mid(Dw, phra, 3)
                End Select
                If Len(phras) > 1 And VBA.Right(phras, 1) = " " Then
                    hfspace = hfspace + 1 '­p¦¸
                    GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
                End If
                'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
                wrong = wrong + 1 'ÀËµø¥Î!
                If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
                    DoEvents 'MsgBox "½ÐÀË¬d!!"
        '        ElseIf wrong = 49761 Then
        '            MsgBox "½ÐÀË¬d!!"
                End If
    '            if rst Set rst = CurrentDb.OpenRecordset("SELECT  µüÀWªí.* FROM µüÀWªí WHERE (((µüÀWªí.µü·J) like '" & phras & "'));")
                With rst
    '                If .RecordCount = 0 Then
                    .FindFirst "µü·J like '" & phras & "'"
                    If .NoMatch Then
    '                    .MoveLast
                        .AddNew
                        rst("µü·J") = phras
    '                    rst("¦¸¼Æ") = 1'¹w³]­È¤w¬°1
                        On Error GoTo ¿ù»~³B²z
                        .Update 'dbUpdateBatch, True
                    Else
1                       .edit
                        rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                        .Update
                    End If
    '                .Close
                End With
11          Next phra
2       Next phralh
        EndTime = VBA.Time
        AppActivate "Microsoft word"
        MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
            & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
            & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
            & vbCr & "¦r¤¸¼Æ=" & dwL '.Characters.Count
    'End With
    'd.Visible = True
    If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
    '    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
        AppActivate "Microsoft access"
    End If
    d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
    d.docmd.Maximize
    rst.Close: db.Close: Set d = Nothing
    options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
    End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access'2002/11/15
¿ù»~³B²z:
    Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
        Case Is = 3022
            rst.Requery
            rst.FindFirst "µü·J like '" & Trim(phras) & "'"
            GoTo 1
        Case Is = 5941 '¶°¦X¤¤ªº¦¨­û¤£¦s¦b(«ü¶W¹L¤å¥óªø«×!)
            GoTo 2
        Case Is = 91, 3078
            MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
            'access.Application.Quit
            d.Quit
            End
        Case Else
            MsgBox Err.Number & ":" & Err.Description, vbExclamation
            Resume
    End Select
End Sub
Sub ¶i¶¥µüÀW1() '2002/11/15­nSub¤~¯à¦bWord¤¤°õ¦æ!
    On Error GoTo ¿ù»~³B²z
    Dim wrong As Long, phra As Long, phras As String, phralh As Byte
    Dim StTime As Date, EndTime As Date
    Dim hfspace As Long
    Dim Length As String
    Dim i As Byte, j As Byte
    Length = InputBox("½Ð«ü©w¤ÀªRµü·J¤§¤W­­,³Ì¦h255­Ó¦r", , "5")
    If Length = "" Or Not IsNumeric(Length) Then End
    If CByte(Length) < 1 Or CByte(Length) > 255 Then End
    options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
    StTime = VBA.Time
    Set d = CreateObject("access.application")
    '©ÎSet d = CreateObject("Access.Application.9")
    d.UserControl = True
    d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
    d.docmd.SelectObject d.acTable, "µüÀWªí", True
    'd.Visible = True 'ÀË¬d¥Î
    Set db = d.CurrentDb
    Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
    If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
    'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
    '    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
        db.Execute "DELETE * FROM µüÀWªí"
    End If
    j = CByte(Length)
    With ActiveDocument
        For phralh = 1 To j
    '    ­ì¼È©w³Ìªø¬°5­Ó¦rºc¦¨ªºµü,¤µ§ï§@ÅÜ¼Æj,«h­­©óByte¤j¤p¦Õ!
            For phra = 1 To .Characters.Count
                If phra + (phralh - 1) <= .Characters.Count Then
                    phras = ""
                    For i = 0 To phralh - 1
                        phras = phras & .Characters(phra + i)
                    Next i
                End If
                If Len(phras) > 1 And VBA.Right(phras, 1) = " " Then
                    hfspace = hfspace + 1 '­p¦¸
                    GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
                End If
                'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
                wrong = wrong + 1 'ÀËµø¥Î!
                If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
                    MsgBox "½ÐÀË¬d!!"
        '        ElseIf wrong = 49761 Then
        '            MsgBox "½ÐÀË¬d!!"
                End If
                With rst
                    .FindFirst "µü·J like '" & phras & "'"
                    If .NoMatch Then
        '                .MoveLast
                        .AddNew
                        rst("µü·J") = phras
                        rst("¦¸¼Æ") = 1
                        On Error GoTo ¿ù»~³B²z
                        .Update 'dbUpdateBatch, True
                    Else
1                       .edit
                        rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                        .Update
                    End If
                End With
11          Next phra
2       Next phralh
        EndTime = VBA.Time
        AppActivate "Microsoft word"
        MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
            & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
            & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
            & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
    End With
    'd.Visible = True
    If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
    '    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
        AppActivate "Microsoft access"
    End If
    d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
    d.docmd.Maximize
    rst.Close: db.Close: Set d = Nothing
    options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
    End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
    Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
        Case Is = 3022
            rst.Requery
            rst.FindFirst "µü·J like '" & Trim(phras) & "'"
            GoTo 1
        Case Is = 5941 '¶°¦X¤¤ªº¦¨­û¤£¦s¦b(«ü¶W¹L¤å¥óªø«×!)
            GoTo 2
        Case Is = 91, 3078
            MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
            'access.Application.Quit
            d.Quit
            End
        Case Else
            MsgBox Err.Number & ":" & Err.Description, vbExclamation
            Resume
    End Select
End Sub
Sub «ü©w¦r¼ÆµüÀW() '2002/11/11
    On Error GoTo ¿ù»~³B²z
    Dim wrong As Long, phra As Long, phras As String, phralh As String
    Dim StTime As Date, EndTime As Date
    Dim hfspace As Long
    phralh = InputBox("½Ð¥Îªü©Ô§B¼Æ¦r«ü©wµüªº²Õ¦¨¦r¼Æ,³Ì¦h¦r¼Æ¬°¡u11¡v!", "«ü©wµü·J¦r¼Æ", "2")
    If phralh = "" Or Not IsNumeric(phralh) Then Exit Sub
    If CByte(phralh) > 11 Or CByte(phralh) < 1 Then Exit Sub
    options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
    Set d = CreateObject("access.application")
    d.UserControl = True
    d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
    d.docmd.SelectObject d.acTable, "µüÀWªí", True
    'd.Visible = True 'ÀË¬d¥Î
    Set db = d.CurrentDb
    Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
    If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
    'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
    '    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
        db.Execute "DELETE * FROM µüÀWªí"
    End If
    StTime = VBA.Time
    With ActiveDocument
        For phra = 1 To .Characters.Count
            Select Case CByte(phralh)
                Case Is = 1
                    phras = .Characters(phra)
                Case Is = 2
                    If phra + 1 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1)
                Case Is = 3
                    If phra + 2 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1) & _
                            .Characters(phra + 2)
                Case Is = 4
                    If phra + 3 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1) & _
                            .Characters(phra + 2) & .Characters(phra + 3)
                Case Is = 5
                    If phra + 4 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1) & _
                            .Characters(phra + 2) & .Characters(phra + 3) & _
                            .Characters(phra + 4)
                Case Is = 6
                    If phra + 5 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1) & _
                            .Characters(phra + 2) & .Characters(phra + 3) & _
                            .Characters(phra + 4) & .Characters(phra + 5)
                Case Is = 7
                    If phra + 6 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1) & _
                            .Characters(phra + 2) & .Characters(phra + 3) & _
                            .Characters(phra + 4) & .Characters(phra + 5) & _
                            .Characters(phra + 6)
                Case Is = 8
                    If phra + 7 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1) & _
                            .Characters(phra + 2) & .Characters(phra + 3) & _
                            .Characters(phra + 4) & .Characters(phra + 5) & _
                            .Characters(phra + 6) & .Characters(phra + 7)
                Case Is = 9
                    If phra + 8 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1) & _
                            .Characters(phra + 2) & .Characters(phra + 3) & _
                            .Characters(phra + 4) & .Characters(phra + 5) & _
                            .Characters(phra + 6) & .Characters(phra + 7) & _
                            .Characters(phra + 8)
                Case Is = 10
                    If phra + 9 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1) & _
                            .Characters(phra + 2) & .Characters(phra + 3) & _
                            .Characters(phra + 4) & .Characters(phra + 5) & _
                            .Characters(phra + 6) & .Characters(phra + 7) & _
                            .Characters(phra + 8) & .Characters(phra + 9)
                Case Is = 11
                    If phra + 10 <= .Characters.Count Then _
                    phras = .Characters(phra) & .Characters(phra + 1) & _
                            .Characters(phra + 2) & .Characters(phra + 3) & _
                            .Characters(phra + 4) & .Characters(phra + 5) & _
                            .Characters(phra + 6) & .Characters(phra + 7) & _
                            .Characters(phra + 8) & .Characters(phra + 9) & _
                            .Characters(phra + 10)
            End Select
            If Len(phras) > 1 And VBA.Right(phras, 1) = " " Then
                hfspace = hfspace + 1 '­p¦¸
                GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
            End If
            'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
            wrong = wrong + 1 'ÀËµø¥Î!
    '        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
    '            MsgBox "½ÐÀË¬d!!"
    ''        ElseIf wrong = 49761 Then
    ''            MsgBox "½ÐÀË¬d!!"
    '        End If
            With rst
                .FindFirst "µü·J like '" & phras & "'"
                If .NoMatch Then
                    .AddNew
                    rst("µü·J") = phras
    '                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                    .Update 'dbUpdateBatch, True
                Else
                    .edit
                    rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                    .Update
                End If
            End With
11      Next phra
        EndTime = VBA.Time
        AppActivate "Microsoft word"
        MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
            & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
            & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
            & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
    End With
    If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
    '    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
        AppActivate "Microsoft access"
        d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
        d.docmd.Maximize
    End If
    d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
    d.docmd.Maximize
    rst.Close: db.Close: Set d = Nothing
    options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
    End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
    Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
        Case Is = 91, 3078
            MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
            'access.Application.Quit
            d.Quit
            End
        Case Else
            MsgBox Err.Number & ":" & Err.Description, vbExclamation
            Resume
    End Select
End Sub
Sub «ü©w11¦r¼ÆµüÀW()     '2002/11/15'¥H¦¹¬°¨Ò,¥i§@¬°¹w¥ý­­©w¦r¼Æªº¦U­Óµ{§Ç(¥»¨Ò¬°11­Ó¦rªº¬d¸ß)
    On Error GoTo ¿ù»~³B²z
    Dim wrong As Long, phra As Long, phras As String, phralh As String
    Dim StTime As Date, EndTime As Date
    Dim hfspace As Long
    'phralh = InputBox("½Ð¥Îªü©Ô§B¼Æ¦r«ü©wµüªº²Õ¦¨¦r¼Æ,³Ì¦h¦r¼Æ¬°¡u11¡v!", "«ü©wµü·J¦r¼Æ", "2")
    'If phralh = "" Or Not IsNumeric(phralh) Then Exit Sub
    'If CByte(phralh) > 11 Or CByte(phralh) < 1 Then Exit Sub
    options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
    Set d = CreateObject("access.application")
    d.UserControl = True
    d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
    d.docmd.SelectObject d.acTable, "µüÀWªí", True
    'd.Visible = True 'ÀË¬d¥Î
    Set db = d.CurrentDb
    Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
    If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
    'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
    '    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
        db.Execute "DELETE * FROM µüÀWªí"
    End If
    StTime = VBA.Time
    With ActiveDocument
        For phra = 1 To .Characters.Count
            If phra + 10 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2) & .Characters(phra + 3) & _
                        .Characters(phra + 4) & .Characters(phra + 5) & _
                        .Characters(phra + 6) & .Characters(phra + 7) & _
                        .Characters(phra + 8) & .Characters(phra + 9) & _
                        .Characters(phra + 10)
            If Len(phras) > 1 And VBA.Right(phras, 1) = " " Then
                hfspace = hfspace + 1 '­p¦¸
                GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
            End If
            'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
            wrong = wrong + 1 'ÀËµø¥Î!
    '        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
    '            MsgBox "½ÐÀË¬d!!"
    ''        ElseIf wrong = 49761 Then
    ''            MsgBox "½ÐÀË¬d!!"
    '        End If
            With rst
                .FindFirst "µü·J like '" & phras & "'"
                If .NoMatch Then
                    .AddNew
                    rst("µü·J") = phras
    '                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                    .Update 'dbUpdateBatch, True
                Else
                    .edit
                    rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                    .Update
                End If
            End With
11      Next phra
        EndTime = VBA.Time
        AppActivate "Microsoft word"
        MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
            & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
            & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
            & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
    End With
    If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
    '    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
        AppActivate "Microsoft access"
        d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
        d.docmd.Maximize
    End If
    d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
    d.docmd.Maximize
    rst.Close: db.Close: Set d = Nothing
    options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
    End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
    Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
        Case Is = 91, 3078
            MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
            'access.Application.Quit
            d.Quit
            End
        Case Else
            MsgBox Err.Number & ":" & Err.Description, vbExclamation
            Resume
    End Select
End Sub
Sub «ü©w10¦r¼ÆµüÀW() '2002/11/15
    On Error GoTo ¿ù»~³B²z
    Dim wrong As Long, phra As Long, phras As String, phralh As String
    Dim StTime As Date, EndTime As Date
    Dim hfspace As Long
    options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
    Set d = CreateObject("access.application")
    d.UserControl = True
    d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
    d.docmd.SelectObject d.acTable, "µüÀWªí", True
    'd.Visible = True 'ÀË¬d¥Î
    Set db = d.CurrentDb
    Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
    If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
    'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
    '    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
        db.Execute "DELETE * FROM µüÀWªí"
    End If
    StTime = VBA.Time
    With ActiveDocument
        For phra = 1 To .Characters.Count
            If phra + 9 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2) & .Characters(phra + 3) & _
                        .Characters(phra + 4) & .Characters(phra + 5) & _
                        .Characters(phra + 6) & .Characters(phra + 7) & _
                        .Characters(phra + 8) & .Characters(phra + 9)
            If Len(phras) > 1 And VBA.Right(phras, 1) = " " Then
                hfspace = hfspace + 1 '­p¦¸
                GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
            End If
            'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
            wrong = wrong + 1 'ÀËµø¥Î!
    '        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
    '            MsgBox "½ÐÀË¬d!!"
    ''        ElseIf wrong = 49761 Then
    ''            MsgBox "½ÐÀË¬d!!"
    '        End If
            With rst
                .FindFirst "µü·J like '" & phras & "'"
                If .NoMatch Then
                    .AddNew
                    rst("µü·J") = phras
    '                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                    .Update 'dbUpdateBatch, True
                Else
                    .edit
                    rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                    .Update
                End If
            End With
11      Next phra
        EndTime = VBA.Time
        AppActivate "Microsoft word"
        MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
            & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
            & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
            & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
    End With
    If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
    '    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
        AppActivate "Microsoft access"
        d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
        d.docmd.Maximize
    End If
    d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
    d.docmd.Maximize
    rst.Close: db.Close: Set d = Nothing
    options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
    End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
    Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
        Case Is = 91, 3078
            MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
            'access.Application.Quit
            d.Quit
            End
        Case Else
            MsgBox Err.Number & ":" & Err.Description, vbExclamation
            Resume
    End Select
End Sub

Sub «ü©w9¦r¼ÆµüÀW()  '2002/11/15
    On Error GoTo ¿ù»~³B²z
    Dim wrong As Long, phra As Long, phras As String, phralh As String
    Dim StTime As Date, EndTime As Date
    Dim hfspace As Long
    options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
    Set d = CreateObject("access.application")
    d.UserControl = True
    d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
    d.docmd.SelectObject d.acTable, "µüÀWªí", True
    'd.Visible = True 'ÀË¬d¥Î
    Set db = d.CurrentDb
    Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
    If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
    'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
    '    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
        db.Execute "DELETE * FROM µüÀWªí"
    End If
    StTime = VBA.Time
    With ActiveDocument
        For phra = 1 To .Characters.Count
            If phra + 8 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2) & .Characters(phra + 3) & _
                        .Characters(phra + 4) & .Characters(phra + 5) & _
                        .Characters(phra + 6) & .Characters(phra + 7) & _
                        .Characters(phra + 8)
            If Len(phras) > 1 And VBA.Right(phras, 1) = " " Then
                hfspace = hfspace + 1 '­p¦¸
                GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
            End If
            'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
            wrong = wrong + 1 'ÀËµø¥Î!
    '        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
    '            MsgBox "½ÐÀË¬d!!"
    ''        ElseIf wrong = 49761 Then
    ''            MsgBox "½ÐÀË¬d!!"
    '        End If
            With rst
                .FindFirst "µü·J like '" & phras & "'"
                If .NoMatch Then
                    .AddNew
                    rst("µü·J") = phras
    '                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                    .Update 'dbUpdateBatch, True
                Else
                    .edit
                    rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                    .Update
                End If
            End With
11      Next phra
        EndTime = VBA.Time
        AppActivate "Microsoft word"
        MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
            & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
            & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
            & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
    End With
    If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
    '    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
        AppActivate "Microsoft access"
        d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
        d.docmd.Maximize
    End If
    d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
    d.docmd.Maximize
    rst.Close: db.Close: Set d = Nothing
    options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
    End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
    Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
        Case Is = 91, 3078
            MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
            'access.Application.Quit
            d.Quit
            End
        Case Else
            MsgBox Err.Number & ":" & Err.Description, vbExclamation
            Resume
    End Select
End Sub


Sub «ü©w8¦r¼ÆµüÀW()   '2002/11/15
    On Error GoTo ¿ù»~³B²z
    Dim wrong As Long, phra As Long, phras As String, phralh As String
    Dim StTime As Date, EndTime As Date
    Dim hfspace As Long
    options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
    Set d = CreateObject("access.application")
    d.UserControl = True
    d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
    d.docmd.SelectObject d.acTable, "µüÀWªí", True
    'd.Visible = True 'ÀË¬d¥Î
    Set db = d.CurrentDb
    Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
    If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
    'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
    '    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
        db.Execute "DELETE * FROM µüÀWªí"
    End If
    StTime = VBA.Time
    With ActiveDocument
        For phra = 1 To .Characters.Count
            If phra + 7 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2) & .Characters(phra + 3) & _
                        .Characters(phra + 4) & .Characters(phra + 5) & _
                        .Characters(phra + 6) & .Characters(phra + 7)
            If Len(phras) > 1 And VBA.Right(phras, 1) = " " Then
                hfspace = hfspace + 1 '­p¦¸
                GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
            End If
            'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
            wrong = wrong + 1 'ÀËµø¥Î!
    '        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
    '            MsgBox "½ÐÀË¬d!!"
    ''        ElseIf wrong = 49761 Then
    ''            MsgBox "½ÐÀË¬d!!"
    '        End If
            With rst
                .FindFirst "µü·J like '" & phras & "'"
                If .NoMatch Then
                    .AddNew
                    rst("µü·J") = phras
    '                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                    .Update 'dbUpdateBatch, True
                Else
                    .edit
                    rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                    .Update
                End If
            End With
11      Next phra
        EndTime = VBA.Time
        AppActivate "Microsoft word"
        MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
            & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
            & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
            & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
    End With
    If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
    '    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
        AppActivate "Microsoft access"
        d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
        d.docmd.Maximize
    End If
    d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
    d.docmd.Maximize
    rst.Close: db.Close: Set d = Nothing
    options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
    End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
    Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
        Case Is = 91, 3078
            MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
            'access.Application.Quit
            d.Quit
            End
        Case Else
            MsgBox Err.Number & ":" & Err.Description, vbExclamation
            Resume
    End Select
End Sub

Sub «ü©w6¦r¼ÆµüÀW()    '2002/11/15
    On Error GoTo ¿ù»~³B²z
    Dim wrong As Long, phra As Long, phras As String, phralh As String
    Dim StTime As Date, EndTime As Date
    Dim hfspace As Long
    options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
    Set d = CreateObject("access.application")
    d.UserControl = True
    d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
    d.docmd.SelectObject d.acTable, "µüÀWªí", True
    'd.Visible = True 'ÀË¬d¥Î
    Set db = d.CurrentDb
    Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
    If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
    'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
    '    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
        db.Execute "DELETE * FROM µüÀWªí"
    End If
    StTime = VBA.Time
    With ActiveDocument
        For phra = 1 To .Characters.Count
            If phra + 5 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2) & .Characters(phra + 3) & _
                        .Characters(phra + 4) & .Characters(phra + 5)
            If Len(phras) > 1 And VBA.Right(phras, 1) = " " Then
                hfspace = hfspace + 1 '­p¦¸
                GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
            End If
            'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
            wrong = wrong + 1 'ÀËµø¥Î!
    '        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
    '            MsgBox "½ÐÀË¬d!!"
    ''        ElseIf wrong = 49761 Then
    ''            MsgBox "½ÐÀË¬d!!"
    '        End If
            With rst
                .FindFirst "µü·J like '" & phras & "'"
                If .NoMatch Then
                    .AddNew
                    rst("µü·J") = phras
    '                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                    .Update 'dbUpdateBatch, True
                Else
                    .edit
                    rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                    .Update
                End If
            End With
11      Next phra
        EndTime = VBA.Time
        AppActivate "Microsoft word"
        MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
            & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
            & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
            & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
    End With
    If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
    '    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
        AppActivate "Microsoft access"
        d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
        d.docmd.Maximize
    End If
    d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
    d.docmd.Maximize
    rst.Close: db.Close: Set d = Nothing
    options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
    End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
    Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
        Case Is = 91, 3078
            MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
            'access.Application.Quit
            d.Quit
            End
        Case Else
            MsgBox Err.Number & ":" & Err.Description, vbExclamation
            Resume
    End Select
End Sub

Sub «ü©w5¦r¼ÆµüÀW()     '2002/11/15
    On Error GoTo ¿ù»~³B²z
    Dim wrong As Long, phra As Long, phras As String, phralh As String
    Dim StTime As Date, EndTime As Date
    Dim hfspace As Long
    options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
    Set d = CreateObject("access.application")
    d.UserControl = True
    d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
    d.docmd.SelectObject d.acTable, "µüÀWªí", True
    'd.Visible = True 'ÀË¬d¥Î
    Set db = d.CurrentDb
    Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
    If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
    'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
    '    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
        db.Execute "DELETE * FROM µüÀWªí"
    End If
    StTime = VBA.Time
    With ActiveDocument
        For phra = 1 To .Characters.Count
            If phra + 4 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2) & .Characters(phra + 3) & _
                        .Characters(phra + 4)
            If Len(phras) > 1 And VBA.Right(phras, 1) = " " Then
                hfspace = hfspace + 1 '­p¦¸
                GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
            End If
            'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
            wrong = wrong + 1 'ÀËµø¥Î!
    '        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
    '            MsgBox "½ÐÀË¬d!!"
    ''        ElseIf wrong = 49761 Then
    ''            MsgBox "½ÐÀË¬d!!"
    '        End If
            With rst
                .FindFirst "µü·J like '" & phras & "'"
                If .NoMatch Then
                    .AddNew
                    rst("µü·J") = phras
    '                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                    .Update 'dbUpdateBatch, True
                Else
                    .edit
                    rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                    .Update
                End If
            End With
11      Next phra
        EndTime = VBA.Time
        AppActivate "Microsoft word"
        MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
            & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
            & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
            & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
    End With
    If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
    '    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
        AppActivate "Microsoft access"
        d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
        d.docmd.Maximize
    End If
    d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
    d.docmd.Maximize
    rst.Close: db.Close: Set d = Nothing
    options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
    End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
    Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
        Case Is = 91, 3078
            MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
            'access.Application.Quit
            d.Quit
            End
        Case Else
            MsgBox Err.Number & ":" & Err.Description, vbExclamation
            Resume
    End Select
End Sub
Sub «ü©w4¦r¼ÆµüÀW()       '2002/11/15
    On Error GoTo ¿ù»~³B²z
    Dim wrong As Long, phra As Long, phras As String, phralh As String
    Dim StTime As Date, EndTime As Date
    Dim hfspace As Long
    options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
    Set d = CreateObject("access.application")
    d.UserControl = True
    d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
    d.docmd.SelectObject d.acTable, "µüÀWªí", True
    'd.Visible = True 'ÀË¬d¥Î
    Set db = d.CurrentDb
    Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
    If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
    'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
    '    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
        db.Execute "DELETE * FROM µüÀWªí"
    End If
    StTime = VBA.Time
    With ActiveDocument
        For phra = 1 To .Characters.Count
            If phra + 3 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2) & .Characters(phra + 3)
            If Len(phras) > 1 And VBA.Right(phras, 1) = " " Then
                hfspace = hfspace + 1 '­p¦¸
                GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
            End If
            'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
            wrong = wrong + 1 'ÀËµø¥Î!
    '        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
    '            MsgBox "½ÐÀË¬d!!"
    ''        ElseIf wrong = 49761 Then
    ''            MsgBox "½ÐÀË¬d!!"
    '        End If
            With rst
                .FindFirst "µü·J like '" & phras & "'"
                If .NoMatch Then
                    .AddNew
                    rst("µü·J") = phras
    '                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                    .Update 'dbUpdateBatch, True
                Else
                    .edit
                    rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                    .Update
                End If
            End With
11      Next phra
        EndTime = VBA.Time
        AppActivate "Microsoft word"
        MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
            & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
            & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
            & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
    End With
    If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
    '    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
        AppActivate "Microsoft access"
        d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
        d.docmd.Maximize
    End If
    d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
    d.docmd.Maximize
    rst.Close: db.Close: Set d = Nothing
    options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
    End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
    Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
        Case Is = 91, 3078
            MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
            'access.Application.Quit
            d.Quit
            End
        Case Else
            MsgBox Err.Number & ":" & Err.Description, vbExclamation
            Resume
    End Select
End Sub

Sub «ü©w3¦r¼ÆµüÀW()      '2002/11/15
    On Error GoTo ¿ù»~³B²z
    Dim wrong As Long, phra As Long, phras As String, phralh As String
    Dim StTime As Date, EndTime As Date
    Dim hfspace As Long
    options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
    Set d = CreateObject("access.application")
    d.UserControl = True
    d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
    d.docmd.SelectObject d.acTable, "µüÀWªí", True
    'd.Visible = True 'ÀË¬d¥Î
    Set db = d.CurrentDb
    Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
    If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
    'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
    '    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
        db.Execute "DELETE * FROM µüÀWªí"
    End If
    StTime = VBA.Time
    With ActiveDocument
        For phra = 1 To .Characters.Count
            If phra + 2 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2)
            If Len(phras) > 1 And VBA.Right(phras, 1) = " " Then
                hfspace = hfspace + 1 '­p¦¸
                GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
            End If
            'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
            wrong = wrong + 1 'ÀËµø¥Î!
    '        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
    '            MsgBox "½ÐÀË¬d!!"
    ''        ElseIf wrong = 49761 Then
    ''            MsgBox "½ÐÀË¬d!!"
    '        End If
            With rst
                .FindFirst "µü·J like '" & phras & "'"
                If .NoMatch Then
                    .AddNew
                    rst("µü·J") = phras
    '                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                    .Update 'dbUpdateBatch, True
                Else
                    .edit
                    rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                    .Update
                End If
            End With
11      Next phra
        EndTime = VBA.Time
        AppActivate "Microsoft word"
        MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
            & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
            & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
            & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
    End With
    If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
    '    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
        AppActivate "Microsoft access"
        d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
        d.docmd.Maximize
    End If
    d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
    d.docmd.Maximize
    rst.Close: db.Close: Set d = Nothing
    options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
    End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
    Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
        Case Is = 91, 3078
            MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
            'access.Application.Quit
            d.Quit
            End
        Case Else
            MsgBox Err.Number & ":" & Err.Description, vbExclamation
            Resume
    End Select
End Sub

Sub «ü©w2¦r¼ÆµüÀW()       '2002/11/15
    On Error GoTo ¿ù»~³B²z
    Dim wrong As Long, phra As Long, phras As String, phralh As String
    Dim StTime As Date, EndTime As Date
    Dim hfspace As Long
    options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
    Set d = CreateObject("access.application")
    d.UserControl = True
    d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
    d.docmd.SelectObject d.acTable, "µüÀWªí", True
    'd.Visible = True 'ÀË¬d¥Î
    Set db = d.CurrentDb
    Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
    If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
    'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
    '    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
        db.Execute "DELETE * FROM µüÀWªí"
    End If
    StTime = VBA.Time
    With ActiveDocument
        For phra = 1 To .Characters.Count
            If phra + 1 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1)
            If Len(phras) > 1 And VBA.Right(phras, 1) = " " Then
                hfspace = hfspace + 1 '­p¦¸
                GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
            End If
            'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
            wrong = wrong + 1 'ÀËµø¥Î!
    '        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
    '            MsgBox "½ÐÀË¬d!!"
    ''        ElseIf wrong = 49761 Then
    ''            MsgBox "½ÐÀË¬d!!"
    '        End If
            With rst
                .FindFirst "µü·J like '" & phras & "'"
                If .NoMatch Then
                    .AddNew
                    rst("µü·J") = phras
    '                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                    .Update 'dbUpdateBatch, True
                Else
                    .edit
                    rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                    .Update
                End If
            End With
11      Next phra
        EndTime = VBA.Time
        AppActivate "Microsoft word"
        MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
            & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
            & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
            & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
    End With
    If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
    '    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
        AppActivate "Microsoft access"
        d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
        d.docmd.Maximize
    End If
    d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
    d.docmd.Maximize
    rst.Close: db.Close: Set d = Nothing
    options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
    End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
    Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
        Case Is = 91, 3078
            MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
            'access.Application.Quit
            d.Quit
            End
        Case Else
            MsgBox Err.Number & ":" & Err.Description, vbExclamation
            Resume
    End Select
End Sub

Sub «ü©w1¦r¼ÆµüÀW()        '2002/11/15
    On Error GoTo ¿ù»~³B²z
    Dim wrong As Long, phra As Long, phras As String, phralh As String
    Dim StTime As Date, EndTime As Date
    Dim hfspace As Long
    options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
    Set d = CreateObject("access.application")
    d.UserControl = True
    d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
    d.docmd.SelectObject d.acTable, "µüÀWªí", True
    'd.Visible = True 'ÀË¬d¥Î
    Set db = d.CurrentDb
    Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
    If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
    'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
    '    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
        db.Execute "DELETE * FROM µüÀWªí"
    End If
    StTime = VBA.Time
    With ActiveDocument
        For phra = 1 To .Characters.Count
                phras = .Characters(phra)
            If Len(phras) > 1 And VBA.Right(phras, 1) = " " Then
                hfspace = hfspace + 1 '­p¦¸
                GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
            End If
            'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
            wrong = wrong + 1 'ÀËµø¥Î!
    '        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
    '            MsgBox "½ÐÀË¬d!!"
    ''        ElseIf wrong = 49761 Then
    ''            MsgBox "½ÐÀË¬d!!"
    '        End If
            With rst
                .FindFirst "µü·J like '" & phras & "'"
                If .NoMatch Then
                    .AddNew
                    rst("µü·J") = phras
    '                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                    .Update 'dbUpdateBatch, True
                Else
                    .edit
                    rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                    .Update
                End If
            End With
11      Next phra
        EndTime = VBA.Time
        AppActivate "Microsoft word"
        MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
            & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
            & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
            & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
    End With
    If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
    '    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
        AppActivate "Microsoft access"
        d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
        d.docmd.Maximize
    End If
    d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
    d.docmd.Maximize
    rst.Close: db.Close: Set d = Nothing
    options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
    End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
    Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
        Case Is = 91, 3078
            MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
            'access.Application.Quit
            d.Quit
            End
        Case Else
            MsgBox Err.Number & ":" & Err.Description, vbExclamation
            Resume
    End Select
End Sub

Sub «ü©w7¦r¼ÆµüÀW()      '2002/11/15'¥H¦¹¬°¨Ò,¥i§@¬°¹w¥ý­­©w¦r¼Æªº¦U­Óµ{§Ç(¥»¨Ò¬°7­Ó¦rªº¬d¸ß)
    On Error GoTo ¿ù»~³B²z
    Dim wrong As Long, phra As Long, phras As String, phralh As String
    Dim StTime As Date, EndTime As Date
    Dim hfspace As Long
    'phralh = InputBox("½Ð¥Îªü©Ô§B¼Æ¦r«ü©wµüªº²Õ¦¨¦r¼Æ,³Ì¦h¦r¼Æ¬°¡u11¡v!", "«ü©wµü·J¦r¼Æ", "2")
    'If phralh = "" Or Not IsNumeric(phralh) Then Exit Sub
    'If CByte(phralh) > 11 Or CByte(phralh) < 1 Then Exit Sub
    options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
    Set d = CreateObject("access.application")
    d.UserControl = True
    d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
    d.docmd.SelectObject d.acTable, "µüÀWªí", True
    'd.Visible = True 'ÀË¬d¥Î
    Set db = d.CurrentDb
    Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
    If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
    'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
    '    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
        db.Execute "DELETE * FROM µüÀWªí"
    End If
    StTime = VBA.Time
    With ActiveDocument
        For phra = 1 To .Characters.Count
            If phra + 6 <= .Characters.Count Then _
                phras = .Characters(phra) & .Characters(phra + 1) & _
                        .Characters(phra + 2) & .Characters(phra + 3) & _
                        .Characters(phra + 4) & .Characters(phra + 5) & _
                        .Characters(phra + 6)
            If Len(phras) > 1 And VBA.Right(phras, 1) = " " Then
                hfspace = hfspace + 1 '­p¦¸
                GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
            End If
            'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
            wrong = wrong + 1 'ÀËµø¥Î!
    '        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
    '            MsgBox "½ÐÀË¬d!!"
    ''        ElseIf wrong = 49761 Then
    ''            MsgBox "½ÐÀË¬d!!"
    '        End If
            With rst
                .FindFirst "µü·J like '" & phras & "'"
                If .NoMatch Then
                    .AddNew
                    rst("µü·J") = phras
    '                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                    .Update 'dbUpdateBatch, True
                Else
                    .edit
                    rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                    .Update
                End If
            End With
11      Next phra
        EndTime = VBA.Time
        AppActivate "Microsoft word"
        MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
            & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
            & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
            & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
    End With
    If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
    '    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
        AppActivate "Microsoft access"
        d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
        d.docmd.Maximize
    End If
    d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
    d.docmd.Maximize
    rst.Close: db.Close: Set d = Nothing
    options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
    End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
    Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
        Case Is = 91, 3078
            MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
            'access.Application.Quit
            d.Quit
            End
        Case Else
            MsgBox Err.Number & ":" & Err.Description, vbExclamation
            Resume
    End Select
End Sub
Sub «ü©w¦r¼ÆµüÀW1() '2002/11/15'®Ä¯à¸ûºC!
    On Error GoTo ¿ù»~³B²z
    Dim wrong As Long, phra As Long, phras, phralh As String
    Dim StTime As Date, EndTime As Date
    Dim hfspace As Long
    Dim a1, i As Byte, j As Byte
    phralh = InputBox("½Ð¥Îªü©Ô§B¼Æ¦r«ü©wµüªº²Õ¦¨¦r¼Æ,³Ì¦h¦r¼Æ¬°¡u255¡v!", "«ü©wµü·J¦r¼Æ", "2")
    If phralh = "" Or Not IsNumeric(phralh) Then Exit Sub
    If CByte(phralh) > 255 Or CByte(phralh) < 1 Then Exit Sub
    options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
    Set d = CreateObject("access.application")
    d.UserControl = True
    d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
    d.docmd.SelectObject d.acTable, "µüÀWªí", True
    'd.Visible = True 'ÀË¬d¥Î
    Set db = d.CurrentDb
    Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
    If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
    'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
    '    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
        db.Execute "DELETE * FROM µüÀWªí"
    End If
    StTime = VBA.Time
    With ActiveDocument
        For phra = 1 To .Characters.Count
            j = CByte(phralh)
            ReDim a1(1 To j) As String
            If j > 1 Then
                If phra + (phralh - 1) <= .Characters.Count Then
                    For j = 1 To j
                        For i = 0 To j - 1
                                a1(j) = a1(j) & .Characters(phra + i)
                        Next i
        '                    Debug.Print a1(j)
                    Next j
                    phras = a1(j - 1)
                End If
            Else
                phras = .Characters(phra)
            End If
            If Len(phras) > 1 And VBA.Right(phras, 1) = " " Then
                hfspace = hfspace + 1 '­p¦¸
                GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
            End If
            'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
            wrong = wrong + 1 'ÀËµø¥Î!
    '        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
    '            MsgBox "½ÐÀË¬d!!"
    ''        ElseIf wrong = 49761 Then
    ''            MsgBox "½ÐÀË¬d!!"
    '        End If
            With rst
                .FindFirst "µü·J like '" & phras & "'"
                If .NoMatch Then
                    .AddNew
                    rst("µü·J") = phras
    '                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                    .Update 'dbUpdateBatch, True
                Else
                    .edit
                    rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                    .Update
                End If
            End With
11      Next phra
        EndTime = VBA.Time
        AppActivate "Microsoft word"
        MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
            & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
            & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
            & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
    End With
    If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
    '    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
        AppActivate "Microsoft access"
        d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
        d.docmd.Maximize
    End If
    d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
    d.docmd.Maximize
    rst.Close: db.Close: Set d = Nothing
    options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
    End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
    Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
        Case Is = 91, 3078
            MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
            'access.Application.Quit
            d.Quit
            End
        Case Else
            MsgBox Err.Number & ":" & Err.Description, vbExclamation
            Resume
    End Select
End Sub
Sub «ü©w¦r¼ÆµüÀW2() '2002/11/15®Ä¯à»P­ì³]­p®t¤£¦h,¦ý¥iÅÜ¼Æ¤Æ!
    On Error GoTo ¿ù»~³B²z
    Dim wrong As Long, phra As Long, phras, phralh As String
    Dim StTime As Date, EndTime As Date
    Dim hfspace As Long
    Dim i As Byte, j As Byte
    phralh = InputBox("½Ð¥Îªü©Ô§B¼Æ¦r«ü©wµüªº²Õ¦¨¦r¼Æ,³Ì¦h¦r¼Æ¬°¡u255¡v!", "«ü©wµü·J¦r¼Æ", "2")
    If phralh = "" Or Not IsNumeric(phralh) Then Exit Sub
    If CByte(phralh) > 255 Or CByte(phralh) < 1 Then Exit Sub
    options.SaveInterval = 0 '¨ú®ø¦Û°ÊÀx¦s
    Set d = CreateObject("access.application")
    d.UserControl = True
    d.OpenCurrentDatabase "d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb", False
    d.docmd.SelectObject d.acTable, "µüÀWªí", True
    'd.Visible = True 'ÀË¬d¥Î
    Set db = d.CurrentDb
    Set rst = db.OpenRecordset("µüÀWªí", dbOpenDynaset)
    If rst.RecordCount > 0 Then '­nÀò±o¥þ³¡ªºµ§¼Æ¶·¥ÎMoveLast¦ý¦¹¥u»Ý§PÂ_¦³¨S¦³­ì¥ýªº°O¿ý§Y¥i!
    'rst¥´¶}¥H«á¥u·|¨ú±o²Ä¤@µ§°O¿ý!
    '    db.Execute "DELETE ¦rÀWªí.* FROM ¦rÀWªí"
        db.Execute "DELETE * FROM µüÀWªí"
    End If
    StTime = VBA.Time
    j = CByte(phralh)
    With ActiveDocument
        For phra = 1 To .Characters.Count
    '        If j > 1 Then'§Y¨Ï¬O³æ¦r¤]¤£¶·¤À§O³B²z¤F!!
                If phra + (phralh - 1) <= .Characters.Count Then
                    phras = ""
                    For i = 0 To j - 1
                        phras = phras & .Characters(phra + i)
                    Next i
                End If
    '        Else
    '            phras = .Characters(phra)
    '        End If
            If Len(phras) > 1 And VBA.Right(phras, 1) = " " Then
                hfspace = hfspace + 1 '­p¦¸
                GoTo 11 '¦r¦ê¥kÃä¬O¥b§ÎªÅ®æ®É,AccessUpdate®É·|¾P¥h,¥B©óµü·J¥çµL·N·N,¬G¤£­p!
            End If
            'ª½±µ¶i¤J¤U¤@­Ó¦r¦ê¤ñ¹ï
            wrong = wrong + 1 'ÀËµø¥Î!
    '        If wrong Mod 29688 = 0 Then '¨ì29688®É·|²£¥ÍOLE¨S¦³¦^À³ªº¿ù»~,¬G¦b¦¹·²·|¨à
    '            MsgBox "½ÐÀË¬d!!"
    ''        ElseIf wrong = 49761 Then
    ''            MsgBox "½ÐÀË¬d!!"
    '        End If
            With rst
                .FindFirst "µü·J like '" & phras & "'"
                If .NoMatch Then
                    .AddNew
                    rst("µü·J") = phras
    '                rst("¦¸¼Æ") = 1'¹w³]­È¤w©w¬°1
                    .Update 'dbUpdateBatch, True
                Else
                    .edit
                    rst("¦¸¼Æ") = rst("¦¸¼Æ") + 1
                    .Update
                End If
            End With
11      Next phra
        EndTime = VBA.Time
        AppActivate "Microsoft word"
        MsgBox "²Î­p§¹¦¨!!" & vbCr & "(¡°¦@°õ¦æ¤F" & wrong & "¦¸ªºÀË¬d¡°)" _
            & "µü·J¥kÃä¥b§ÎªÅ®æ¤Z" & hfspace & "¦¸,©¿²¤¤£­p!" _
            & vbCr & "¡°¯Ó®É:" & Format(EndTime - StTime, "n¤Às¬í") & "¡°" _
            & vbCr & "¦r¤¸¼Æ=" & .Characters.Count
    End With
    If MsgBox("­n§Y¨èÀËµøµ²ªG¶Ü?", vbYesNo + vbQuestion) = vbYes Then
    '    Set d = GetObject("d:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\µüÀW.mdb")
        AppActivate "Microsoft access"
        d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
        d.docmd.Maximize
    End If
    d.docmd.OpenTable "µüÀWªí", , d.acReadOnly
    d.docmd.Maximize
    rst.Close: db.Close: Set d = Nothing
    options.SaveInterval = 10 '«ì´_¦Û°ÊÀx¦s
    End '¥ÎExit SubµLªk¨C¦¸Ãö³¬Access
¿ù»~³B²z:
    Select Case Err.Number '¥D¯Á¤Þ­È­«½Æ
        Case Is = 91, 3078
            MsgBox "½Ð¦A«ö¤@¦¸!", vbCritical
            'access.Application.Quit
            d.Quit
            End
        Case Else
            MsgBox Err.Number & ":" & Err.Description, vbExclamation
            Resume
    End Select
End Sub

Sub ¤å¥ó¦rÀW_old()
    Dim DR As Range, d As Document, Char, charText As String, preChar As String _
        , x() As String, xT() As Long, i As Long, j As Long, ExcelSheet  As Object, _
        ds As Date, de As Date '
    Static xlsp As String
    On Error GoTo ErrH:
    'xlsp = "C:\Documents and Settings\Superwings\®à­±\"
    Set d = ActiveDocument
    xlsp = ¨ú±o®à­±¸ô®| & "\" 'GetDeskDir() & "\"
    If Dir(xlsp) = "" Then xlsp = ¨ú±o®à­±¸ô®| 'GetDeskDir ' "C:\Users\Wong\Desktop\" '& Replace(ActiveDocument.Name, ".doc", "") & "¦rÀW.XLS"
    'If Dir(xlsp) = "" Then xlsp = "C:\Documents and Settings\Superwings\®à­±\" & Replace(ActiveDocument.Name, ".doc", "") & "¦rÀW.XLS"
    'xlsp = "C:\Documents and Settings\Superwings\®à­±\" & Replace(ActiveDocument.Name, ".doc", "") & "¦rÀW.XLS"
    xlsp = InputBox("½Ð¿é¤J¦sÀÉ¸ô®|¤ÎÀÉ¦W(¥þÀÉ¦W,§t°ÆÀÉ¦W)!" & vbCr & vbCr & _
            "¹w³]±N¥H¦¹word¤å¥óÀÉ¦W + ""¦rÀW.XLSX""¦rºó,¦s©ó®à­±¤W", "¦rÀW½Õ¬d", xlsp & Replace(ActiveDocument.Name, ".doc", "") & "¦rÀW" & VBA.StrConv(VBA.Time, vbWide) & ".XLSX")
    If xlsp = "" Then Exit Sub
    
    ds = VBA.Timer
    
    With d
        For Each Char In d.Characters
            charText = Char
            If Not charText = VBA.Chr(13) And charText <> "-" And Not charText Like "[a-zA-Z0-9¢¯-¢¸]" Then
                'If Not charText Like "[a-z1-9]" & VBA.Chr(-24153) & VBA.Chr(-24152) & " ¡@¡B'""¡u¡v¡y¡z¡]¡^¡Ð¡H¡I]" Then
    '            If InStr(VBA.Chr(-24153) & VBA.Chr(-24152) & vba.Chr(2) & "¡E[]¡e¡f¡£¡¤¡K¡F,¡A.¡C¡D ¡@¡B'""¡¥¡¦`\{}¡a¡b¡u¡v¡y¡z¡]¡^¡m¡n¡q¡r¡Ð¡H¡I]", charText) = 0 Then
                If InStr(VBA.ChrW(-24153) & VBA.ChrW(-24152) & VBA.Chr(2) & "¡E[]¡e¡f¡£¡¤¡K¡F,¡A.¡C¡D ¡@¡B'""¡¥¡¦`\{}¡a¡b¡u¡v¡y¡z¡]¡^¡m¡n¡q¡r¡Ð¡H¡I]", charText) = 0 Then
                'vba.Chr(2)¥i¯à¬Oµù¸}¼Ð°O
                    If preChar <> charText Then
                        'If UBound(X) > 0 Then
                            If preChar = "" Then 'If IsEmpty(X) Then'¦pªG¬O¤@¶}©l
                                GoTo 1
                            ElseIf UBound(Filter(x, charText)) Then ' <> charText Then  '¦pªG©|µL¦¹¦r
1                               ReDim Preserve x(i)
                                ReDim Preserve xT(i)
                                x(i) = charText
                                xT(i) = xT(i) + 1
                                i = i + 1
                            Else
                                GoSub ¦rÀW¥[¤@
                            End If
                        'End If
                    Else
                        GoSub ¦rÀW¥[¤@
                    End If
                    preChar = Char
                End If
            End If
        Next Char
    End With
    
    Dim Doc As New Document, Xsort() As String, u As Long ', xTsort() As Integer, k As Long, so As Long, ww As String
    'ReDim Xsort(i) As String ', xtsort(i) as Integer
    'ReDim Xsort(d.Characters.Count) As String
    If u = 0 Then u = 1 '­YµL°õ¦æ¡u¦rÀW¥[¤@:¡v°Æµ{§Ç,­YµL¶W¹L1¦¸ªº¦rÀW¡A«h¡@Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "¡B" & x(j - 1) & _
                                    ·|¥X¿ù¡G°}¦C¯Á¤Þ¶W¥X½d³ò 2015/11/5
    
    ReDim Xsort(u) As String
    Set ExcelSheet = CreateObject("Excel.Sheet")
    With ExcelSheet.Application
        For j = 1 To i
            .cells(j, 1) = x(j - 1)
            .cells(j, 2) = xT(j - 1)
            Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "¡B" & x(j - 1) 'Xsort(xT(j - 1)) & ww '°}¦C±Æ§Ç'2010/10/29
        Next j
    End With
    'Doc.ActiveWindow.Visible = False
    'U = UBound(Xsort)
    For j = u To 0 Step -1 '°}¦C±Æ§Ç'2010/10/29
        If Xsort(j) <> "" Then
            With Doc
                If Len(.Range) = 1 Then '©|¥¼¿é¤J¤º®e
                    .Range.InsertAfter "¦rÀW = " & j & "¦¸¡G¡]" & Len(Replace(Xsort(j), "¡B", "")) & "¦r¡^"
                    .Range.Paragraphs(1).Range.font.Size = 12
                    .Range.Paragraphs(.Paragraphs.Count).Range.font.Name = "·s²Ó©úÅé"
                    .Range.Paragraphs(.Paragraphs.Count).Range.font.NameAscii = "Times New Roman"
                    '.Range.Paragraphs(1).Range.Font.Bold = True
                Else
                    .Range.InsertParagraphAfter
                    .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
                    .Range.InsertAfter "¦rÀW = " & j & "¦¸¡G¡]" & Len(Replace(Xsort(j), "¡B", "")) & "¦r¡^"
                    .Range.Paragraphs(.Paragraphs.Count).Range.font.Size = 12
                    '.Range.Paragraphs(.Paragraphs.Count).Range.Bold = True
                    .Range.Paragraphs(.Paragraphs.Count).Range.font.Name = "·s²Ó©úÅé"
                    .Range.Paragraphs(.Paragraphs.Count).Range.font.NameAscii = "Times New Roman"
                End If
                .Range.InsertParagraphAfter
                .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
                .Range.Paragraphs(.Paragraphs.Count).Range.font.Size = 12
    '            .Range.Paragraphs(.Paragraphs.Count).Range.Bold = False
                .Range.InsertAfter Replace(Xsort(j), "¡B", VBA.Chr(9), 1, 1) 'vba.Chr(9)¬°©w¦ì¦r¤¸(TabÁä­È)
                .Range.InsertParagraphAfter
                If InStr(.Range.Paragraphs(.Paragraphs.Count).Range, "¦rÀW") = 0 Then
                    .Range.Paragraphs(.Paragraphs.Count - 1).Range.font.Name = "¼Ð·¢Åé"
                Else
                    .Range.Paragraphs(.Paragraphs.Count).Range.font.Name = "·s²Ó©úÅé"
                    .Range.Paragraphs(.Paragraphs.Count).Range.font.NameAscii = "Times New Roman"
                End If
            End With
        End If
    Next j
    
    With Doc.Paragraphs(1).Range
         .InsertParagraphBefore
         .font.NameAscii = "times new roman"
        Doc.Paragraphs(1).Range.InsertParagraphAfter
        Doc.Paragraphs(1).Range.InsertParagraphAfter
        Doc.Paragraphs(1).Range.InsertAfter "§A´£¨Ñªº¤å¥»¦@¨Ï¥Î¤F" & i & "­Ó¤£¦Pªº¦r¡]¶Ç²Î¦r»PÂ²¤Æ¦r¤£¤©¦X¨Ö¡^"
    End With
    
    Doc.ActiveWindow.Visible = True
    '
    
    'U = UBound(xT)
    'ReDim Xsort(U) As String, xTsort(U) As Long
    '
    'i = d.Characters
    'For j = 1 To i '¥Î¼Æ¦r¬Û¤ñ
    '    For k = 0 To U 'xT°}¦C¤¤¨C­Ó¤¸¯À³£»Pj¤ñ
    '        If xT(k) = j Then
    '            Xsort(so) = x(k)
    '            xTsort(so) = xT(k)
    '            so = so + 1
    '        End If
    '    Next k
    'Next j
    
    'With doc
    '    .Range.InsertAfter "¦rÀW=0001"
    '    .Range.InsertParagraphAfter
    'End With
    
    
    ' Cells.Select
    '    Selection.Sort Key1:=Range("B1"), Order1:=xlDescending, Header:=xlGuess, _
    '        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    
    
    'Set ExcelSheet = Nothing'¦¹¦æ·|¨Ï®ø¥¢
    'Set d = Nothing
    de = VBA.Timer
    MsgBox "§¹¦¨¡I" & vbCr & vbCr & "¶O®É" & VBA.Left(de - ds, 5) & "¬í!"
    ExcelSheet.Application.Visible = True
    ExcelSheet.Application.UserControl = True
    ExcelSheet.SaveAs xlsp '"C:\Macros\¦u¯uTEST.XLS"
    Doc.SaveAs Replace(xlsp, "XLS", "doc") '¤À¤j¤p¼g
    'Doc.SaveAs "c:\test1.doc"
    AppActivate "microsoft excel"
    Exit Sub
¦rÀW¥[¤@:
    For j = 0 To UBound(x)
        If x(j) = charText Then
            xT(j) = xT(j) + 1
            If u < xT(j) Then u = xT(j) '°O¤U³Ì°ª¦rÀW,¥H«K±Æ§Ç(±N±ý±Æ§Ç¤§°}¦C³Ì°ª¤¸¯À­È³]¬°¦¹,«h¤£·|¶W¥X°}¦C.
            '¦h¦¹¤@¦æ¦]¬°­n­«½Æ§PÂ_­pºâ¦n´X¦¸,¬G®Ä¯à¤£¼W¤Ï´î''®Ä¯àÁÙ¬O®t¤£¦h°Õ.
            Exit For
        End If
    Next j
    
    Return
ErrH:
    Select Case Err.Number
        Case Else
            MsgBox Err.Number & Err.Description, vbCritical 'STOP: Resume
    '        Resume
            End
        
    End Select
End Sub

Function UppercaeEnglishLetter() '­^¤å¤j¼g¦r¥À
    Dim WD, wdct As Long, i As Byte
    For i = 65 To 90
        Debug.Print VBA.Chr(i) & vbCr
    Next
End Function
Function LowercaseEnglishLetter() '­^¤å¤p¼g¦r¥À
    Dim i As Byte
    For i = 97 To 122
        Debug.Print VBA.Chr(i) & vbCr
    Next
End Function

Function trimStrForSearch_PlainText(x As String) As String
    Rem 20230128 ¬Ñ¥f¦~ªì¤C ®]¦u¯u¡ÑchatGPT¤jµÐÂÄ¡GVBA Overload Functionality¡G
    'chatGPT¤jµÐÂÄ·s¦~¦N²»¡G ·Q½Ð°Ý VBA ¬O¤£¬O¤£¯à¹³ C# ¤@¼Ë¨ç¦¡¤èªk¥i¥H ¦h¸ü¡B­«¸ü¡]overload¡^¡H
    'VBA (Visual Basic for Applications) ¬O¤@ºØ·L³nªºµ{¦¡»y¨¥¡A¥D­n¥Î©ó¦Û°Ê¤Æ Microsoft Office À³¥Îµ{¦¡¤¤¡CVBA ¤£¤ä´©¨ç¦¡ªº¦h¸ü©M­«¸ü¡C³o·N¨ýµÛ¡A±z¤£¯à¦b VBA ¤¤©w¸q¨ã¦³¬Û¦P¦WºÙ¦ý°Ñ¼Æ¤£¦Pªº¦h­Ó¨ç¦¡¡C
    
    Dim ayToTrim As Variant, a As Variant
    On Error GoTo eH
    ayToTrim = Array(VBA.Chr(13), VBA.Chr(9), VBA.Chr(10), VBA.Chr(11), VBA.Chr(13) & VBA.Chr(7), VBA.Chr(13) & VBA.Chr(10))
    x = VBA.Trim(x)
    For Each a In ayToTrim
        'x = VBA.Replace(x, a, "")
        Do While VBA.Left(x, Len(a)) = a
            x = VBA.Mid(x, Len(a) + 1)
        Loop
        Do While VBA.Right(x, Len(a)) = a
            x = VBA.Mid(x, 1, Len(x) - Len(a))
        Loop
    Next a
    trimStrForSearch_PlainText = x
    Exit Function
eH:
    Select Case Err.Number
        Case Else
            MsgBox Err.Number & Err.Description
    '        Resume
    End Select
End Function

Function trimStrForSearch(x As String, sl As word.Selection) As String
    'https://docs.microsoft.com/zh-tw/dotnet/visual-basic/programming-guide/language-features/procedures/passing-arguments-by-value-and-by-reference
    Dim ayToTrim As Variant, a As Variant, rng As Range, slTxtR As String
    On Error GoTo eH
    slTxtR = sl.Characters(sl.Characters.Count)
    ayToTrim = Array(VBA.Chr(13), VBA.Chr(9), VBA.Chr(10), VBA.Chr(11), VBA.Chr(13) & VBA.Chr(7), VBA.Chr(13) & VBA.Chr(10))
    x = VBA.Trim(x)
    For Each a In ayToTrim
        'x = VBA.Replace(x, a, "")
        Do While VBA.Left(x, Len(a)) = a
            x = VBA.Mid(x, Len(a))
        Loop
        Do While VBA.Right(x, Len(a)) = a
            x = VBA.Mid(x, 1, Len(x) - Len(a))
        Loop
    Next a
    trimStrForSearch = x
    If sl.Type <> wdSelectionIP Then
        If UBound(VBA.Strings.Filter(ayToTrim, slTxtR)) > -1 Then
        'If sl.Characters(sl.Characters.Count) = vba.Chr(13) Then
            Set rng = sl.Range
            rng.SetRange sl.start, sl.End - Len(slTxtR)
            rng.Select
        End If
    End If
    Exit Function
eH:
    Select Case Err.Number
        Case Else
            MsgBox Err.Number & Err.Description
    '        Resume
    End Select
End Function

Sub ResetSelectionAvoidSymbols()
    Dim rng As Range
    Set rng = Selection.Range.Duplicate
    'Do While Selection.Characters(Selection.Characters.Count) = VBA.Chr(13)
    Do While rng.Characters(rng.Characters.Count) = VBA.Chr(13)
        rng.SetRange rng.start, rng.End - 1
        'Selection.MoveLeft wdCharacter, 1, wdExtend'¿ï¨ú°_¨´¤è¦V¤£¦P·|¦³¼vÅT 20240924
    Loop
    rng.Select
    Do While Selection.Characters(1) = VBA.Chr(13)
    'Do While rng.Characters(1) = VBA.Chr(13)
        Selection.SetRange Selection.start + 1, Selection.End
    Loop
End Sub
'Function Symbol() '¼ÐÂI²Å¸¹ªí
'Dim f As Variant
'f = Array("¡C", "¡v", VBA.Chr(-24152), "¡G", "¡A", "¡F", _
'    "¡B", "¡u", ".", vba.Chr(34), ":", ",", ";", _
'    "¡K¡K", "...", "¡^", ")", "-")  '¥ý³]©w¼ÐÂI²Å¸¹°}¦C¥H³Æ¥Î
'                                'VBA.Chr(-24152)¬O¡u¡¨¡v,¥ÑAsc¨ç¼Æ¦b¿ï¨ú(.SelText)¡u¡¨¡v®É¨ú±o;vba.Chr(34):¡u"¡v
'End Function
Function isSymbol(ByVal a As String) As Boolean
    Dim f As String
    f = PunctuationString
    If InStr(1, f, a, vbTextCompare) Then
        isSymbol = True
    End If
End Function

Sub ²M°£¿ï¨ú³Bªº©Ò¦³²Å¸¹() '¥Ñ¹Ï®ÑºÞ²zsymbles¼Ò²Õ²M°£¼ÐÂI²Å¸¹§ï½s'¥]¬Aµù¸}¡B¼Æ¦r
'Dim F, a As String, i As Integer
Dim f, i As Integer, ur As UndoRecord
SystemSetup.stopUndo ur, "²M°£¿ï¨ú³Bªº©Ò¦³²Å¸¹"
f = Array("-", "¡P", "¡E", "¡C", "¡v", VBA.Chr(-24152), "¡G", "¡A", "¡F", _
    "¡B", "¡u", ".", VBA.Chr(34), ":", ",", ";", _
    "¡K¡K", "...", "¡D", "¡i", "¡j", " ", "¡m", "¡n", "¡q", "¡r", "¡H" _
    , "¡I", "¡£", "¡¤", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0" _
    , "¡y", "¡z", VBA.Chr(13), VBA.ChrW(9312), VBA.ChrW(9313), VBA.ChrW(9314), VBA.ChrW(9315), VBA.ChrW(9316) _
    , VBA.ChrW(9317), VBA.ChrW(9318), VBA.ChrW(9319), VBA.ChrW(9320), VBA.ChrW(9321), VBA.ChrW(9322), VBA.ChrW(9323) _
    , VBA.ChrW(9324), VBA.ChrW(9325), VBA.ChrW(9326), VBA.ChrW(9327), VBA.ChrW(9328), VBA.ChrW(9329), VBA.ChrW(9330) _
    , VBA.ChrW(9331), VBA.ChrW(8221), """") '¥ý³]©w¼ÐÂI²Å¸¹°}¦C¥H³Æ¥Î
    '¥þ§Î¶ê¬A©·¼È¤£¨ú¥N¡I
    'a = ActiveDocument.Content
'    Set a = ActiveDocument.Range.FormattedText '¥]§t®æ¦¡¤Æªº¸ê°T
    For i = 0 To UBound(f)
        If InStr(Selection.Range.text, f(i)) Then
            'a = Replace(a, F(i), "")
            Selection.Range.Find.Execute f(i), True, , , , , , wdFindStop, True, "", wdReplaceAll
        End If
    Next
    'ActiveDocument.Content = a
SystemSetup.contiUndo ur
End Sub

Function isª`­µ²Å¸¹(ByVal a As String, Optional rng As Variant) As Boolean
Dim f As String
On Error GoTo eH
If Len(a) > 1 Then Exit Function
f = "£t£u£v£w£x£y£z£{£|£}£~£¡£¢£££¤£¥£¦£§£¨£©£ª£¸£¹£º£«£¬£­£®£¯£°£±£²£³£´£µ£¶£·£½  £¾  £¿  £»"
If a = VBA.ChrW(20008) Then
    If Not rng Is Nothing Then
        If rng.start = 0 Then
            If InStr(f, rng.Next.Characters(1)) Then
                isª`­µ²Å¸¹ = True
                Exit Function
            End If
        ElseIf rng.End = rng.Document.Range.End - 1 Then
            If InStr(f, rng.Previous.Characters(1)) Then
                isª`­µ²Å¸¹ = True
                Exit Function
            End If
        End If
    End If
Else
    If InStr(f, a) Then isª`­µ²Å¸¹ = True
End If
Exit Function
eH:
Select Case Err.Number
    Case 424 '¦¹³B»Ý­nª«¥ó
        Set rng = Nothing
        Resume
    Case Else
        MsgBox Err.Number & Err.Description
        Debug.Print Err.Number & Err.Description
End Select
End Function

Sub ¿ï¨ú¬q¸¨²Å¸¹()
'²Ä1¬qªº³Ì«á()
'    With ActiveDocument.Paragraphs(1).Range
'        ActiveDocument.Range(.End - 1, .End).Select
'    End With
Dim i As Integer
For i = 1 To ActiveDocument.Paragraphs.Count
    With ActiveDocument.Paragraphs(i).Range
        ActiveDocument.Range(.End - 1, .End).Select
    End With
Next i
End Sub


Sub ³y¦r¦r¤¸ÀË¬d() '«D²Ó©úÅéÀË¬d,2004/8/23
Dim ch
For Each ch In ActiveDocument.Characters
'    If AscW(ch) < -1491 Or AscW(ch) > 19968 Then
    If Asc(ch) < -24256 Or (0 > Asc(ch) And Asc(ch) >= -1468) Then
        ch.Select
        ch.font.Name = "EUDC"
    End If
Next ch
End Sub

Sub ª`¸}²Å¸¹¸m´«() '2004/10/17
Dim WD As Range 'As Range 'Wordsª«¥ó§Yªí¤@­ÓRangeª«¥ó,¨£½u¤W»¡©ú!
'Dim i As Long ' Integer
'­n¥ý°õ¦æ¥þ§ÎÂà¥b§Î,³o¼Ëwords¤~¯à¥¿½T§PÂ_¬°¼Æ¦r
¥þ§Î¼Æ¦rÂà´«¦¨¥b§Î¼Æ¦r
With Selection '­ì¥H¾ã¥÷¤å¥ó(ActiveDocument),¤µ¦ý¥H¿ï¨ú½d³ò¾ã²z,¦ý¦]§ó§ï­È¦Ó¼vÅT,§@¼o!
    If .Type = wdSelectionIP Then .Document.Select '¦pªG¨S¦³¿ï¨ú½d³ò(¬°´¡¤JÂI)«h³B²z¾ã¥÷¤å¥ó
    If .Document.path = "" Then
        For Each WD In .words
            '­n¬O¼Æ¦r¥B«e«á¤£¯à¥[¡£¡¤©Î¡e¡f¤~°õ¦æ¡I
            If Not WD.text Like "¡£" And Not WD.text Like "¡e" And Not WD Like "[[]" And Not WD Like "[]]" Then
                If IsNumeric(WD) Then
                    If WD.End = .Document.Content.StoryLength Or WD.start = 0 Then GoTo w '¤å¥ó¤§­º§À¥t¥~³B²z
                    If Not WD.Previous Like "¡£" And Not WD.Previous Like "¡e" And Not WD.Previous Like "[[]" _
                        And Not WD.Next Like "¡¤" And Not WD.Next Like "¡f" And Not WD.Next Like "]" Then
w:                      If WD <= 20 Then 'Arial Unicode MS[ºØÃþ]¸Ì"¬A¸¹¤å¼Æ¦r"¥u¦³¤G¤Q­Ó!
                            With WD
                                '¿ï¨ú·|§ïÅÜSelectionªº½d³ò,¬G¤µ¨ú®ø!
'                                .Select 'Wordsª«¥ó§Yªí¤@­ÓRangeª«¥ó,¨£½u¤W»¡©ú!
                                .font.Name = "Arial Unicode MS"
                                WD.text = VBA.ChrW((9312 - 1) + WD)
                            End With
                        Else '¶W¹L20¸¹ªºµù¸}®É
                            With WD
                                .text = "¡£" & WD.text & "¡¤" '¥[¬A¸¹
                            End With
        '                    MsgBox "¦³¶W¹L20¸¹ªºµù¸},¤£¯à°õ¦æ¡I", vbCritical
        '                    Do Until .Undo(i) = False 'ÁÙ­ìª½¦Ü¤£¯àÁÙ­ì¡]ÁÙ­ì©Ò¦³°Ê§@¡^
        '                    i = i + 1
        '                    Loop
        '                    StatusBar = "Undo was successful " & i & " times!!" '¦bª¬ºA¦CÅã¥Ü¤å¦r¡I
        '                    Exit Sub
                        End If
                    End If
                End If
            End If
        Next
        MsgBox "°õ¦æ§¹²¦¡I", vbInformation
    Else
        MsgBox "¥»¤å¥ó¤£¯à¾Þ§@!", vbCritical
    End If
End With
End Sub

Sub ¥þ§Î¼Æ¦rÂà´«¦¨¥b§Î¼Æ¦r() '2004/10/17-¥Ñ¹Ï®ÑºÞ²z½Æ»s§ï¼¶ªº­ì¦¡¡Ð¤£¦n¡A·|¼vÅT¦r§Î
Dim FNumArray, HNumArray, i As Byte, e As Range
FNumArray = Array("¢¯", "¢°", "¢±", "¢²", "¢³", "¢´", "¢µ", "¢¶", "¢·", "¢¸")
HNumArray = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9")
With ActiveDocument
    For Each e In .Characters
        For i = 1 To UBound(FNumArray) + 1
            If e.text Like FNumArray(i - 1) Then
                e.text = HNumArray(i - 1)
        End If
        Next i
    Next e
End With
End Sub

Sub ¥þ§ÎÂà¥b§Î()
With Selection
    .Range = VBA.StrConv(.Range, vbNarrow)
End With
End Sub
Sub ¶ê¬A¸¹§ï½g¦W¸¹()
If Selection.Type = wdSelectionIP Then Selection.HomeKey wdStory: Selection.EndKey wdStory, wdExtend
Selection.text = Replace(Replace(Selection.text, "¡]", "¡q"), "¡^", "¡r")
End Sub


Sub ®Õ°É¤å¦r¼Ð¦â() '2009/8/23
Register_Event_Handler
'«ü©wÁäF2
' ¥¨¶°2 ¥¨¶°
' ¥¨¶°¿ý»s©ó 2009/8/23¡A¿ý»sªÌ Oscar Sun
'
'    Selection.MoveDown Unit:=wdLine, Count:=2
'    Selection.EndKey Unit:=wdLine
'    Selection.MoveLeft Unit:=wdCharacter, Count:=1
'    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
If Selection.Type = wdSelectionIP Then Exit Sub
    With Selection.font.Shading
        If InStr(ActiveDocument.Name, "±Æ¦L") Then
            .Parent.Color = wdColorRed
            .Texture = wdTextureNone
        Else
            If .Texture = wdTextureNone Then '¦r¤¸ºô©³
                .Texture = wdTexture15Percent
                .ForegroundPatternColor = wdColorBlack
                .BackgroundPatternColor = wdColorWhite
                .Parent.Color = wdColorRed
            Else
                .Texture = wdTextureNone '¦r¤¸ºô©³
                .Parent.Color = wdColorAutomatic
            End If
        End If
    End With
    If InStr(ActiveDocument.Name, "±Æ¦L") Then
        ActiveDocument.Save
'        setOX
'        OX.WinActivate "Microsoft Excel"
        'Dim e As New Excel.Application
        Dim e
        Set e = Excel.Application
        Dim r As Long, i As Byte
        With Selection
            Set e = GetObject(, "Excel.application")
            AppActivate "microsoft excel"
            With e
                '.ActiveWorkbook.Save
                r = .ActiveCell.row
                For i = 1 To 7
                    If .cells(r, i).Value <> "" Then
                        MsgBox "½Ð¨ì·s°O¿ý¦C¡I¡I", vbExclamation
                        Exit Sub
                    End If
                Next i
                .cells(r, 1).Activate
                DoEvents
                .activesheet.Paste
                .cells(r, 2).Value = Selection
                .cells(r, 2).font.Color = wdColorRed
                If Not Selection Like "*[óòñõôöø÷ùûúüýþ¡¸¡¹¡U¡@]*" Then
                    .cells(r, 5) = Len(Selection)
                ElseIf Selection Like "*¡@*" Then
                    .cells(r, 5) = Len(Selection) - 1
                Else
                    .cells(r, 5) = 1
                End If
                .ActiveWorkbook.Save
                .cells(.ActiveCell.row + 1, .ActiveCell.Column).Activate
            End With
        End With
        ´å¼Ð©Ò¦b¦ì¸m®ÑÅÒ
        OX.WinActivate "Adobe Reader"
        AppActivate "microsoft word"
    End If
End Sub

Sub µù¸}½s¸¹«e«á¥[¤è¬A¸¹()
With Selection
    Do

        Selection.GoTo What:=wdGoToFootnote, Which:=wdGoToNext, Count:=1, Name:=""
'        Selection.GoTo What:=wdGoToFootnote, Which:=wdGoToNext, Count:=1, Name:=""
        Selection.Find.ClearFormatting
'        With Selection.Find
'            .Text = ""
'            .Replacement.Text = ""
'            .Forward = True
'            .Wrap = wdFindStop
'            .Format = False
'            .MatchCase = False
'            .MatchWholeWord = False
'            .MatchByte = True
'            .MatchWildcards = False
'            .MatchSoundsLike = False
'            .MatchAllWordForms = False
'        End With
'        If .Find.Execute() = False Then Exit Do
        'Application.Browser.Next
        .TypeText text:="["
        .MoveLeft unit:=wdCharacter, Count:=1, Extend:=wdExtend
        .font.Superscript = wdToggle
'        Selection.Copy
'        Selection.MoveRight Unit:=wdCharacter, Count:=3
'        Selection.MoveLeft Unit:=wdCharacter, Count:=1
'        Selection.Paste
'        Selection.MoveLeft Unit:=wdCharacter, Count:=1
'        Selection.Delete Unit:=wdCharacter, Count:=1
'        Selection.TypeText Text:="¡n"
'        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.MoveRight unit:=wdCharacter, Count:=2
        'Selection.TypeBackspace
        Selection.TypeText text:="]"
        'Selection.MoveRight Unit:=wdCharacter, Count:=1
    Loop 'While .Find.Execute()
End With
End Sub

Sub ¤j³°¤Þ¸¹´«¥xÆW¤Þ¸¹()
Dim a, b, i
a = Array(-24153, -24152, -24155, -24154)  '¡§,¡¨,¡¥,¡¨
b = Array("¡u", "¡v", "¡y", "¡z")

With ActiveDocument.Range.Find
    For i = 0 To 3
        '.Text = a(i)
         '.Replacement.Text = b(i)
         .ClearFormatting
         .Execute VBA.Chr(a(i)), , , , , , , , , b(i), wdReplaceAll
    Next i
End With
End Sub


Sub ¤å¥ó¦rÀW()
Dim d As Document, Char, charText As String, preChar As String _
    , x() As String, xT() As Long, i As Long, j As Long, ds As Date, de As Date     '
'Dim ExcelSheet  As New Excel.Worksheet 'As Object,
'³o¬O¤§«e¥H¥ý´Á¤Þ¥Îªº¤è¦¡¡A¦b³]©w¤Þ¥Î¶µ¥Ø¤¤¤â°Ê¥[¤Jªº¼gªk:https://hankvba.blogspot.com/2018/03/vba.html  ¡B http://markc0826.blogspot.com/2012/07/blog-post.html
'Dim xlApp As Excel.Application, xlBook As Excel.Workbook, xlSheet As Excel.Worksheet
''³o´N¬O«á´Á¤Þ¥Î¡A¥H¦Û­q·s¥éExcelÃþ§Oªº¤èªk¨Ó¹ê§@(¦p¦¹¼gªº½t¬G¬O­ì¨Ó­n§ï¼gªºµ{¦¡½X´N·|¤ñ¸û¤Ö¡AÅÜ°Ê¸û¤p¡A¥B¤]¤£¥²¦ANew¥X¤@­Ó°õ¦æ­ÓÅé¤~¯à°õ¦æ¡G
Dim xlApp, xlBook, xlSheet
Set xlApp = Excel.Application
Set xlBook = Excel.Workbook
Set xlSheet = Excel.Worksheet
Dim ReadingLayoutB As Boolean
Static xlsp As String
On Error GoTo ErrH:
'xlsp = "C:\Documents and Settings\Superwings\®à­±\"
Set d = ActiveDocument
xlsp = ¨ú±o®à­±¸ô®| & "\" 'GetDeskDir() & "\"
If Dir(xlsp) = "" Then xlsp = ¨ú±o®à­±¸ô®| 'GetDeskDir ' "C:\Users\Wong\Desktop\" '& Replace(ActiveDocument.Name, ".doc", "") & "¦rÀW.XLS"
'If Dir(xlsp) = "" Then xlsp = "C:\Documents and Settings\Superwings\®à­±\" & Replace(ActiveDocument.Name, ".doc", "") & "¦rÀW.XLS"
'xlsp = "C:\Documents and Settings\Superwings\®à­±\" & Replace(ActiveDocument.Name, ".doc", "") & "¦rÀW.XLS"
xlsp = InputBox("½Ð¿é¤J¦sÀÉ¸ô®|¤ÎÀÉ¦W(¥þÀÉ¦W,§t°ÆÀÉ¦W)!" & vbCr & vbCr & _
        "¹w³]±N¥H¦¹word¤å¥óÀÉ¦W + ""¦rÀW.XLSX""¦rºó,¦s©ó®à­±¤W", "¦rÀW½Õ¬d", xlsp & Replace(ActiveDocument.Name, ".doc", "") & "¦rÀW" & VBA.StrConv(VBA.Time, vbWide) & ".XLSX")
If xlsp = "" Then Exit Sub

ds = VBA.Timer

With d
    For Each Char In d.Characters
        charText = Char
        If InStr("()¡G>" & VBA.Chr(13) & VBA.Chr(9) & VBA.Chr(10) & VBA.Chr(11) & VBA.ChrW(12), charText) = 0 And charText <> "-" And Not charText Like "[a-zA-Z0-9¢¯-¢¸]" Then
            'If Not charText Like "[a-z1-9]" & VBA.Chr(-24153) & VBA.Chr(-24152) & " ¡@¡B'""¡u¡v¡y¡z¡]¡^¡Ð¡H¡I]" Then
'            If InStr(VBA.Chr(-24153) & VBA.Chr(-24152) & vba.Chr(2) & "¡E[]¡e¡f¡£¡¤¡K¡F,¡A.¡C¡D ¡@¡B'""¡¥¡¦`\{}¡a¡b¡u¡v¡y¡z¡]¡^¡m¡n¡q¡r¡Ð¡H¡I]", charText) = 0 Then
            If InStr(VBA.ChrW(9312) & VBA.ChrW(-24153) & VBA.ChrW(-24152) & VBA.Chr(2) & "¡E[]¡e¡f¡£¡¤¡K¡F,¡A.¡C¡D ¡@¡B'""¡¥¡¦`\{}¡a¡b¡u¡v¡y¡z¡]¡^¡m¡n¡q¡r¡Ð¡H¡I]¡¾¡¼¡j¡i~/¡_¡X" & VBA.Chr(-24152) & VBA.Chr(-24153), charText) = 0 Then
            'vba.Chr(2)¥i¯à¬Oµù¸}¼Ð°O
                If preChar <> charText Then
                    'If UBound(X) > 0 Then
                        If preChar = "" Then 'If IsEmpty(X) Then'¦pªG¬O¤@¶}©l
                            GoTo 1
                        ElseIf UBound(Filter(x, charText)) Then ' <> charText Then  '¦pªG©|µL¦¹¦r
1                           ReDim Preserve x(i)
                            ReDim Preserve xT(i)
                            x(i) = charText
                            xT(i) = xT(i) + 1
                            i = i + 1
                        Else
                            GoSub ¦rÀW¥[¤@
                        End If
                    'End If
                Else
                    GoSub ¦rÀW¥[¤@
                End If
                preChar = Char
            End If
        End If
    Next Char
End With

Dim Doc As New Document, Xsort() As String, u As Long ', xTsort() As Integer, k As Long, so As Long, ww As String
'ReDim Xsort(i) As String ', xtsort(i) as Integer
'ReDim Xsort(d.Characters.Count) As String
If u = 0 Then u = 1 '­YµL°õ¦æ¡u¦rÀW¥[¤@:¡v°Æµ{§Ç,­YµL¶W¹L1¦¸ªº¦rÀW¡A«h¡@Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "¡B" & x(j - 1) & _
                                ·|¥X¿ù¡G°}¦C¯Á¤Þ¶W¥X½d³ò 2015/11/5

ReDim Xsort(u) As String
'Set ExcelSheet = CreateObject("Excel.Sheet")
'Set xlApp = CreateObject("Excel.Application")
'Set xlBook = xlApp.workbooks.Add
'Set xlSheet = xlBook.Worksheets(1)
With xlSheet.Application
    For j = 1 To i
        .cells(j, 1) = x(j - 1)
        .cells(j, 2) = xT(j - 1)
        Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "¡B" & x(j - 1) 'Xsort(xT(j - 1)) & ww '°}¦C±Æ§Ç'2010/10/29
    Next j
End With
'Doc.ActiveWindow.Visible = False
'U = UBound(Xsort)
For j = u To 0 Step -1 '°}¦C±Æ§Ç'2010/10/29
    If Xsort(j) <> "" Then
        With Doc
            If Len(.Range) = 1 Then '©|¥¼¿é¤J¤º®e
                .Range.InsertAfter "¦rÀW = " & j & "¦¸¡G¡]" & Len(Replace(Xsort(j), "¡B", "")) & "¦r¡^"
                .Range.Paragraphs(1).Range.font.Size = 12
                .Range.Paragraphs(.Paragraphs.Count).Range.font.Name = "·s²Ó©úÅé"
                .Range.Paragraphs(.Paragraphs.Count).Range.font.NameAscii = "Times New Roman"
                '.Range.Paragraphs(1).Range.Font.Bold = True
            Else
                .Range.InsertParagraphAfter
                .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
                .Range.InsertAfter "¦rÀW = " & j & "¦¸¡G¡]" & Len(Replace(Xsort(j), "¡B", "")) & "¦r¡^"
                .Range.Paragraphs(.Paragraphs.Count).Range.font.Size = 12
                '.Range.Paragraphs(.Paragraphs.Count).Range.Bold = True
                .Range.Paragraphs(.Paragraphs.Count).Range.font.Name = "·s²Ó©úÅé"
                .Range.Paragraphs(.Paragraphs.Count).Range.font.NameAscii = "Times New Roman"
            End If
            .Range.InsertParagraphAfter
            .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
            .Range.Paragraphs(.Paragraphs.Count).Range.font.Size = 12
'            .Range.Paragraphs(.Paragraphs.Count).Range.Bold = False
            .Range.InsertAfter Replace(Xsort(j), "¡B", VBA.Chr(9), 1, 1) 'vba.Chr(9)¬°©w¦ì¦r¤¸(TabÁä­È)
            .Range.InsertParagraphAfter
            If InStr(.Range.Paragraphs(.Paragraphs.Count).Range, "¦rÀW") = 0 Then
                .Range.Paragraphs(.Paragraphs.Count - 1).Range.font.Name = "¼Ð·¢Åé"
            Else
                .Range.Paragraphs(.Paragraphs.Count).Range.font.Name = "·s²Ó©úÅé"
                .Range.Paragraphs(.Paragraphs.Count).Range.font.NameAscii = "Times New Roman"
            End If
        End With
    End If
Next j

With Doc.Paragraphs(1).Range
     .InsertParagraphBefore
     .font.NameAscii = "times new roman"
    Doc.Paragraphs(1).Range.InsertParagraphAfter
    Doc.Paragraphs(1).Range.InsertParagraphAfter
    Doc.Paragraphs(1).Range.InsertAfter "§A´£¨Ñªº¤å¥»¦@¨Ï¥Î¤F" & i & "­Ó¤£¦Pªº¦r¡]¶Ç²Î¦r»PÂ²¤Æ¦r¤£¤©¦X¨Ö¡^"
End With

Doc.ActiveWindow.Visible = True
'

'U = UBound(xT)
'ReDim Xsort(U) As String, xTsort(U) As Long
'
'i = d.Characters
'For j = 1 To i '¥Î¼Æ¦r¬Û¤ñ
'    For k = 0 To U 'xT°}¦C¤¤¨C­Ó¤¸¯À³£»Pj¤ñ
'        If xT(k) = j Then
'            Xsort(so) = x(k)
'            xTsort(so) = xT(k)
'            so = so + 1
'        End If
'    Next k
'Next j

'With doc
'    .Range.InsertAfter "¦rÀW=0001"
'    .Range.InsertParagraphAfter
'End With


' Cells.Select
'    Selection.Sort Key1:=Range("B1"), Order1:=xlDescending, Header:=xlGuess, _
'        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom


'Set ExcelSheet = Nothing'¦¹¦æ·|¨Ï®ø¥¢
'Set d = Nothing
de = VBA.Timer
If ReadingLayoutB Then d.ActiveWindow.View.ReadingLayout = Not d.ActiveWindow.View.ReadingLayout
MsgBox "§¹¦¨¡I" & vbCr & vbCr & "¶O®É" & VBA.Left(de - ds, 5) & "¬í!", vbInformation
xlSheet.Application.Visible = True
xlSheet.Application.UserControl = True
xlSheet.SaveAs xlsp '"C:\Macros\¦u¯uTEST.XLS"
Doc.SaveAs Replace(xlsp, "XLS", "doc") '¤À¤j¤p¼g
Set Excel.Application = Nothing
Exit Sub
¦rÀW¥[¤@:
For j = 0 To UBound(x)
    If x(j) = charText Then
        xT(j) = xT(j) + 1
        If u < xT(j) Then u = xT(j) '°O¤U³Ì°ª¦rÀW,¥H«K±Æ§Ç(±N±ý±Æ§Ç¤§°}¦C³Ì°ª¤¸¯À­È³]¬°¦¹,«h¤£·|¶W¥X°}¦C.
        '¦h¦¹¤@¦æ¦]¬°­n­«½Æ§PÂ_­pºâ¦n´X¦¸,¬G®Ä¯à¤£¼W¤Ï´î''®Ä¯àÁÙ¬O®t¤£¦h°Õ.
        Exit For
    End If
Next j

Return
ErrH:
Select Case Err.Number
    Case 4605 '¾\Åª¼Ò¦¡¤£¯à½s¿è'¦¹¤èªk©ÎÄÝ©ÊµLªk¨Ï¥Î¡A¦]¬°¦¹©R¥OµLªk¦b¾\Åª¤¤¨Ï¥Î¡C
    '    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
    '        ActiveWindow.ActivePane.View.Type = wdNormalView
    '    Else
    '        ActiveWindow.View.Type = wdNormalView
    '    End If
    '    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
    '        ActiveWindow.ActivePane.View.Type = wdPrintView
    '    Else
    '        ActiveWindow.View.Type = wdPrintView
    '    End If
        'Doc.Application.ActiveWindow.View.ReadingLayout
        d.ActiveWindow.View.ReadingLayout = Not d.ActiveWindow.View.ReadingLayout
        Doc.ActiveWindow.View.ReadingLayout = False
        Doc.ActiveWindow.Visible = False
        ReadingLayoutB = True
        Resume
    Case Else
        MsgBox Err.Number & Err.Description, vbCritical 'STOP: Resume
        'Resume
        End
    
End Select
End Sub

Sub ¤å¥óµüÀW() '¥Ñ¤å¥ó¦rÀW§ï¨Ó'2015/11/28
Dim d As Document, Char, charText As String, preChar As String _
    , x() As String, xT() As Long, i As Long, j As Long, ds As Date, de As Date     '
'Dim ExcelSheet  As New Excel.Worksheet 'As Object,
'Dim xlApp As Excel.Application, xlBook As Excel.Workbook, xlSheet As Excel.Worksheet
Dim xlApp, xlBook, xlSheet
Set xlApp = Excel.Application
Set xlBook = Excel.Workbook
Set xlSheet = Excel.Worksheet
Dim ReadingLayoutB As Boolean
Static ln
Dim xlsp As String
On Error GoTo ErrH:
Set d = ActiveDocument
'If xlsp = "" Then xlsp = ¨ú±o®à­±¸ô®| & "\" 'GetDeskDir() & "\"
'If Dir(xlsp) = "" Then xlsp = ¨ú±o®à­±¸ô®| 'GetDeskDir
'xlsp = InputBox("½Ð¿é¤J¦sÀÉ¸ô®|¤ÎÀÉ¦W(¥þÀÉ¦W,§t°ÆÀÉ¦W)!" & vbCr & vbCr & _
        "¹w³]±N¥H¦¹word¤å¥óÀÉ¦W + ""µüÀW.XLSX""¦rºó,¦s©ó®à­±¤W", "µüÀW½Õ¬d", xlsp & Replace(d.Name, ".doc", "") & "µüÀW" & VBA.StrConv(VBA.Time, vbWide) & ".XLSX")
'If xlsp = "" Then Exit Sub
xlsp = ¨ú±o®à­±¸ô®| & "\" & Replace(d.Name, ".doc", "") & "_µüÀW" & VBA.StrConv(VBA.Time, vbWide) & ".XLSX"
If ln = "" Then ln = 1
ln = InputBox("½Ð«ü©wµü·Jªø«×" & vbCr & vbCr & "ÀÉ®×·|¦s¦b®à­±¤W¦W¬°:" & vbCr & vbCr & Replace(d.Name, ".doc", "") & "_µüÀW" & VBA.StrConv(VBA.Time, vbWide) & ".XLSX" & _
                vbCr & vbCr & "ªºÀÉ®×", , ln + 1)
If ln = "" Then Exit Sub
If Not IsNumeric(ln) Then Exit Sub
If ln > 11 Or ln < 2 Then Exit Sub


ds = VBA.Timer

With d
    For Each Char In d.Characters
        Select Case ln
            Case 2
                charText = Char & Char.Next
            Case 3
                charText = Char & Char.Next & Char.Next.Next
            Case 4
                charText = Char & Char.Next & Char.Next.Next & Char.Next.Next.Next
            Case 5
                charText = Char & Char.Next & Char.Next.Next & Char.Next.Next.Next & Char.Next.Next.Next.Next
            Case 6
                charText = Char & Char.Next & Char.Next.Next & Char.Next.Next.Next & Char.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next
            Case 7
                charText = Char & Char.Next & Char.Next.Next & Char.Next.Next.Next & Char.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next
            Case 8
                charText = Char & Char.Next & Char.Next.Next & Char.Next.Next.Next & Char.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next
            Case 9
                charText = Char & Char.Next & Char.Next.Next & Char.Next.Next.Next & Char.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next.Next
            Case 10
                charText = Char & Char.Next & Char.Next.Next & Char.Next.Next.Next & Char.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next.Next.Next
            Case 11
                charText = Char & Char.Next & Char.Next.Next & Char.Next.Next.Next & Char.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next.Next.Next & Char.Next.Next.Next.Next.Next.Next.Next.Next.Next.Next
        End Select
        If Not charText Like "*[-'¡@ ¡C¡A¡B¡F¡G¡H:,;,¡q¡r¡m¡n ''¡u¡v¡y¡z¡]¡^¡¾¡µ¡H¡I¡]¡^¡i¡j¡X""()<>" _
            & VBA.ChrW(9312) & VBA.Chr(-24153) & VBA.Chr(-24152) & VBA.ChrW(8218) & VBA.Chr(13) & VBA.Chr(10) & VBA.Chr(11) & VBA.ChrW(12) & VBA.Chr(63) & VBA.Chr(9) & VBA.Chr(-24152) & VBA.Chr(-24153) & "¡¾¡¼¡j¡i~/¡_¡X]*" _
            And Not charText Like "*[a-zA-Z0-9¢¯-¢¸]*" And InStr(charText, VBA.ChrW(-243)) = 0 And InStr(charText, VBA.Chr(91)) = 0 And InStr(charText, VBA.Chr(93)) = 0 Then
            'If Not charText Like "[a-z1-9]" & VBA.Chr(-24153) & VBA.Chr(-24152) & " ¡@¡B'""¡u¡v¡y¡z¡]¡^¡Ð¡H¡I]" Then
'            If InStr(VBA.Chr(-24153) & VBA.Chr(-24152) & vba.Chr(2) & "¡E[]¡e¡f¡£¡¤¡K¡F,¡A.¡C¡D ¡@¡B'""¡¥¡¦`\{}¡a¡b¡u¡v¡y¡z¡]¡^¡m¡n¡q¡r¡Ð¡H¡I]", charText) = 0 Then
            If Not charText Like "*[" & VBA.ChrW(-24153) & VBA.ChrW(-24152) & VBA.Chr(2) & "¡E[]¡e¡f¡£¡¤¡K¡F,¡A.¡C¡D ¡@¡B'""¡¥¡¦`\{}¡a¡b¡u¡v¡y¡z¡]¡^¡m¡n¡q¡r¡Ð¡H¡I¡¥¡a¡b]*" Then
            'vba.Chr(2)¥i¯à¬Oµù¸}¼Ð°O
                If preChar <> charText Then
                    'If UBound(X) > 0 Then
                        If preChar = "" Then 'If IsEmpty(X) Then'¦pªG¬O¤@¶}©l
                            GoTo 1
                        ElseIf UBound(Filter(x, charText)) Then ' <> charText Then  '¦pªG©|µL¦¹¦r
1                           ReDim Preserve x(i)
                            ReDim Preserve xT(i)
                            x(i) = charText
                            xT(i) = xT(i) + 1
                            i = i + 1
                        Else
                            GoSub µüÀW¥[¤@
                        End If
                    'End If
                Else
                    GoSub µüÀW¥[¤@
                End If
                preChar = charText
            End If
        End If
    Next
End With
12
Dim Doc As New Document, Xsort() As String, u As Long ', xTsort() As Integer, k As Long, so As Long, ww As String
If u = 0 Then u = 1 '­YµL°õ¦æ¡uµüÀW¥[¤@:¡v°Æµ{§Ç,­YµL¶W¹L1¦¸ªºµüÀW¡A«h¡@Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "¡B" & x(j - 1) & _
                                ·|¥X¿ù¡G°}¦C¯Á¤Þ¶W¥X½d³ò 2015/11/5

ReDim Xsort(u) As String
Set xlApp = CreateObject("Excel.Application")
Set xlBook = xlApp.Workbooks.Add
Set xlSheet = xlBook.Worksheets(1)
With xlSheet.Application
    For j = 1 To i
        .cells(j, 1) = x(j - 1)
        .cells(j, 2) = xT(j - 1)
        Xsort(xT(j - 1)) = Xsort(xT(j - 1)) & "¡B" & x(j - 1) 'Xsort(xT(j - 1)) & ww '°}¦C±Æ§Ç'2010/10/29
    Next j
End With
Doc.ActiveWindow.Visible = False
If d.ActiveWindow.View.ReadingLayout Then ReadingLayoutB = True: d.ActiveWindow.View.ReadingLayout = Not d.ActiveWindow.View.ReadingLayout
'U = UBound(Xsort)
For j = u To 0 Step -1 '°}¦C±Æ§Ç'2010/10/29
    If Xsort(j) <> "" Then
        With Doc
            If Len(.Range) = 1 Then '©|¥¼¿é¤J¤º®e
                .Range.InsertAfter "µüÀW = " & j & "¦¸¡G¡]" & Len(Replace(Xsort(j), "¡B", "")) / ln & "­Ó¡^"
                .Range.Paragraphs(1).Range.font.Size = 12
                .Range.Paragraphs(.Paragraphs.Count).Range.font.Name = "·s²Ó©úÅé"
                .Range.Paragraphs(.Paragraphs.Count).Range.font.NameAscii = "Times New Roman"
                '.Range.Paragraphs(1).Range.Font.Bold = True
            Else
                .Range.InsertParagraphAfter
                .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
                .Range.InsertAfter "µüÀW = " & j & "¦¸¡G¡]" & Len(Replace(Xsort(j), "¡B", "")) / ln & "­Ó¡^"
                .Range.Paragraphs(.Paragraphs.Count).Range.font.Size = 12
                '.Range.Paragraphs(.Paragraphs.Count).Range.Bold = True
                .Range.Paragraphs(.Paragraphs.Count).Range.font.Name = "·s²Ó©úÅé"
                .Range.Paragraphs(.Paragraphs.Count).Range.font.NameAscii = "Times New Roman"
            End If
            .Range.InsertParagraphAfter
            .ActiveWindow.Selection.Range.Collapse Direction:=wdCollapseEnd
            .Range.Paragraphs(.Paragraphs.Count).Range.font.Size = 12
'            .Range.Paragraphs(.Paragraphs.Count).Range.Bold = False
            .Range.InsertAfter Replace(Xsort(j), "¡B", VBA.Chr(9), 1, 1) 'vba.Chr(9)¬°©w¦ì¦r¤¸(TabÁä­È)
            .Range.InsertParagraphAfter
            If InStr(.Range.Paragraphs(.Paragraphs.Count).Range, "µüÀW") = 0 Then
                .Range.Paragraphs(.Paragraphs.Count - 1).Range.font.Name = "¼Ð·¢Åé"
            Else
                .Range.Paragraphs(.Paragraphs.Count).Range.font.Name = "·s²Ó©úÅé"
                .Range.Paragraphs(.Paragraphs.Count).Range.font.NameAscii = "Times New Roman"
            End If
        End With
    End If
Next j

With Doc.Paragraphs(1).Range
     .InsertParagraphBefore
     .font.NameAscii = "times new roman"
    Doc.Paragraphs(1).Range.InsertParagraphAfter
    Doc.Paragraphs(1).Range.InsertParagraphAfter
    Doc.Paragraphs(1).Range.InsertAfter "§A´£¨Ñªº¤å¥»¦@¨Ï¥Î¤F" & i & "­Ó¤£¦Pªºµü·J¡]¶Ç²Î¦r»PÂ²¤Æ¦r¤£¤©¦X¨Ö¡^"
End With

Doc.ActiveWindow.Visible = True

de = VBA.Timer
Doc.SaveAs Replace(xlsp, "XLS", "doc") '¤À¤j¤p¼g
If ReadingLayoutB Then d.ActiveWindow.View.ReadingLayout = Not d.ActiveWindow.View.ReadingLayout
Set d = Nothing ' ActiveDocument.Close wdDoNotSaveChanges

Debug.Print Now

MsgBox "§¹¦¨¡I" & vbCr & vbCr & "¶O®É" & VBA.Left(de - ds, 5) & "¬í!", vbInformation
xlSheet.Application.Visible = True
xlSheet.Application.UserControl = True
xlSheet.SaveAs xlsp
Exit Sub
µüÀW¥[¤@:
For j = 0 To UBound(x)
    If x(j) = charText Then
        xT(j) = xT(j) + 1
        If u < xT(j) Then u = xT(j) '°O¤U³Ì°ªµüÀW,¥H«K±Æ§Ç(±N±ý±Æ§Ç¤§°}¦C³Ì°ª¤¸¯À­È³]¬°¦¹,«h¤£·|¶W¥X°}¦C.
        '¦h¦¹¤@¦æ¦]¬°­n­«½Æ§PÂ_­pºâ¦n´X¦¸,¬G®Ä¯à¤£¼W¤Ï´î''®Ä¯àÁÙ¬O®t¤£¦h°Õ.
        Exit For
    End If
Next j

Return
ErrH:
Select Case Err.Number
    Case 4605 '¾\Åª¼Ò¦¡¤£¯à½s¿è'¦¹¤èªk©ÎÄÝ©ÊµLªk¨Ï¥Î¡A¦]¬°¦¹©R¥OµLªk¦b¾\Åª¤¤¨Ï¥Î¡C
    '    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
    '        ActiveWindow.ActivePane.View.Type = wdNormalView
    '    Else
    '        ActiveWindow.View.Type = wdNormalView
    '    End If
    '    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
    '        ActiveWindow.ActivePane.View.Type = wdPrintView
    '    Else
    '        ActiveWindow.View.Type = wdPrintView
    '    End If
        'Doc.Application.ActiveWindow.View.ReadingLayout
        d.ActiveWindow.View.ReadingLayout = False ' Not d.ActiveWindow.View.ReadingLayout
        Doc.ActiveWindow.View.ReadingLayout = False
        Doc.ActiveWindow.Visible = False
        ReadingLayoutB = True
        Resume
    
    Case 91, 5941 '¨S¦³³]©wª«¥óÅÜ¼Æ©Î With °Ï¶ôÅÜ¼Æ,¶°¦X¤¤©Ò»Ýªº¦¨­û¤£¦s¦b
        GoTo 12
    Case Else
        MsgBox Err.Number & Err.Description, vbCritical 'STOP: Resume
        Resume
        End
    
End Select
End Sub


Sub ®Ñ¦W¸¹½g¦W¸¹ÀË¬d()
Dim s As Long, rng As Range, e, trm As String, ans
Static x() As String, i As Integer
On Error GoTo eH
Do
    Selection.Find.Execute "¡q", , , , , , True, wdFindAsk
    Set rng = Selection.Range
    rng.MoveEndUntil "¡r"
    trm = VBA.Mid(rng, 2)
    
    For Each e In x()
        If StrComp(e, trm) = 0 Then GoTo 1
    Next e
2   ans = MsgBox("¬O§_²¤¹L¡u" & trm & "¡v¡H" & vbCr & vbCr & vbCr & "µ²§ô½Ð«ö NO[§_]", vbExclamation + vbYesNoCancel)
    Select Case ans
        Case vbYes
            ReDim Preserve x(i) As String
            x(i) = trm
            i = i + 1
        Case vbNo
            Exit Sub
    End Select
1
Loop
Exit Sub
eH:
Select Case Err.Number
    Case 92 '¨S¦³³]©w For °j°éªºªì©l­È °}¦C©|¥¼¦³­È
        GoTo 2
End Select
End Sub

Sub ®É¶¡¶b³æ¦ìÂà´«() '2017/5/13 ¦]À³YOUKU»PYOUTUBE®É¶¡¶b³æ¦ì¤£¦P¦Ó³]
'Debug.Print Len(ActiveDocument.Range)
Dim a, aM, aMM, s As Long, e As Long
Dim myRng As Range, chRng As Range
Set myRng = ActiveDocument.Range
Set chRng = ActiveDocument.Range
s = -1
For Each a In ActiveDocument.Characters
    If a.font.Name = "Times New Roman" Then
        If s = -1 Then s = a.start
        If a = VBA.Chr(13) Then GoTo 1
    Else
1       If s > -1 Then
            e = a.Previous.End
            myRng.SetRange s, e
            If InStr(myRng, "http") = 0 Then
                If InStr(Replace(myRng, ":", "", 1, 1), ":") Then 'if find : * 2
                    If InStr(Trim(myRng), " ") Then '¦pªG¦³2­Ó¥H¤W®É¶¡¶b
                        For Each aMM In myRng.Characters
                            If aMM.Next = " " Then
                                e = aMM.End
                                chRng.SetRange s, e
'                                chRng.Select
                                If InStr(Replace(chRng, ":", "", 1, 1), ":") Then 'if find : * 2
                                    GoSub chng
                                End If
                                s = chRng.End + 1
                            End If
                        Next
                    Else '¦pªG¥u¦³1­Ó®É¶¡¶b
                        chRng.SetRange myRng.start, myRng.End
                        GoSub chng
                    End If
                End If
            End If
            s = -1
        End If
    End If
Next
ActiveDocument.Range.Find.Execute "  ", True, , , , , , wdFindContinue, , " ", wdReplaceAll
Exit Sub
chng:
                    For Each aM In chRng.Characters
                        If aM.Next = ":" Then
                            aM.Next.Next.text = str((CInt(aM.Next.Next) * 10 + CInt(aM) * 60) / 10)
                            aM.Next.Delete
                            aM.Delete
                            Exit For
                        End If
                    Next
Return
End Sub
Sub ¤¤°ê­õ¾Ç®Ñ¹q¤l¤Æ­p¹º_ªí®æÂà¤å¦r(ByRef r As Range)
On Error GoTo eH
Dim lngTemp As Long '¦]¬°»~«ö¨ì°lÂÜ­×­q¡A¤~·|¤Þµo°T®§´£¥Ü§R°£Àx¦s®æ¤£·|¦³¼ÐÃÑ
'Dim d As Document
Dim tb As table, C As cell ', ci As Long
'Set d = ActiveDocument
lngTemp = word.Application.DisplayAlerts
If r.Tables.Count > 0 Then
    '¤¤°ê­õ¾Ç®Ñ¹q¤l¤Æ­p¹º.²M°£¤å¥»­¶¤¤ªº½s¸¹Àx¦s®æ r
    For Each tb In r.Tables
        'tb.Columns(1).Delete
        Err.Raise 5992
        Set r = tb.ConvertToText()
    Next tb
End If
'word.Application.DisplayAlerts = lngTemp
Exit Sub
eH:
Select Case Err.Number
    Case 5992 'µLªk­Ó§O¦s¨ú¦¹¶°¦X¤¤ªº¦UÄæ¡A¦]¬°ªí®æ¤¤¦³²V¦XªºÀx¦s®æ¼e«×¡C
        For Each C In tb.Range.cells
'            ci = ci + 1
'            If ci Mod 3 = 2 Then
                'If VBA.IsNumeric(VBA.Left(c.Range.text, VBA.InStr(c.Range.text, "?") - 1)) Then
                If VBA.InStr(C.Range.text, VBA.ChrW(160) & VBA.ChrW(47)) > 0 Then
'                    word.Application.DisplayAlerts = False
                    C.Delete  '§R°£½s¸¹¤§Àx¦s®æ
                End If
'            End If
        Next C
        Resume Next
    Case Else
        MsgBox Err.Number & Err.Description
        End
End Select
End Sub

Sub ¤¤°ê­õ¾Ç®Ñ¹q¤l¤Æ­p¹º_µù¤åÅÜ¤p¥¿¤å¦^¤j()
Dim slRng As Range, a
Set slRng = Selection.Range
¤¤°ê­õ¾Ç®Ñ¹q¤l¤Æ­p¹º_ªí®æÂà¤å¦r slRng
For Each a In slRng.Characters
    Select Case a.font.Color
        Case 34816, 8912896
            a.font.Size = 14
        Case 0
            a.font.Size = 30
    End Select
Next a
End Sub
Sub ¤¤°ê­õ¾Ç®Ñ¹q¤l¤Æ­p¹º_¥h±¼µù¤å«O¯d¥¿¤å()
Dim slRng As Range, a, ur As UndoRecord, d As Document 'Alt + ]
'Set ur = SystemSetup.stopUndo("¤¤°ê­õ¾Ç®Ñ¹q¤l¤Æ­p¹º_¥h±¼µù¤å«O¯d¥¿¤å")
SystemSetup.stopUndo ur, "¤¤°ê­õ¾Ç®Ñ¹q¤l¤Æ­p¹º_¥h±¼µù¤å«O¯d¥¿¤å"
Set d = Docs.ªÅ¥Õªº·s¤å¥ó(True)
'If ActiveDocument.Characters.Count = 1 Then Selection.Paste
If d.Characters.Count = 1 Then d.ActiveWindow.Selection.Paste
If d.ActiveWindow.Selection.Type = wdSelectionIP Then d.Select
Set slRng = d.ActiveWindow.Selection.Range
¤¤°ê­õ¾Ç®Ñ¹q¤l¤Æ­p¹º_ªí®æÂà¤å¦r slRng
For Each a In slRng.Characters
    Select Case a.font.Color
        Case 34816, 8912896
            If a.font.Size <> 12 Then Stop
            a.Delete
        Case 254
            If a.font.Size = 9 Then a.Delete
    End Select
Next a
If MsgBox("¬O§_¨ú¥N²§Åé¦r¡H", vbOKCancel) = vbOK Then ¤å¦rÂà´«.²§Åé¦rÂà¥¿
Beep 'MsgBox "done!", vbInformation
SystemSetup.contiUndo ur
End Sub
Sub ¤¤°ê­õ¾Ç®Ñ¹q¤l¤Æ­p¹º_µù¤å«e«á¥[¬A©·()
    Dim slRng As Range, a, flg As Boolean, ur As UndoRecord, d As Document 'Alt+1
    'Set ur = SystemSetup.stopUndo("¤¤°ê­õ¾Ç®Ñ¹q¤l¤Æ­p¹º_µù¤å«e«á¥[¬A©·")
    SystemSetup.playSound 0.484
    SystemSetup.stopUndo ur, "¤¤°ê­õ¾Ç®Ñ¹q¤l¤Æ­p¹º_µù¤å«e«á¥[¬A©·"
    Set d = Docs.ªÅ¥Õªº·s¤å¥ó(True)
    'If Selection.Type = wdSelectionIP Then ActiveDocument.Select
    'If Selection.Type = wdSelectionIP Then d.Select
    If d Is Nothing Then Set d = ActiveDocument
    d.Range.Paste
    d.Activate
    Set slRng = Selection.Range
    ¤¤°ê­õ¾Ç®Ñ¹q¤l¤Æ­p¹º_ªí®æÂà¤å¦r slRng
    For Each a In slRng.Document.Paragraphs 'forº~Äy¹q¤l¤åÄm¸ê®Æ®w
        If VBA.Left(a.Range, 3) = "[²¨]" Then
            slRng.SetRange a.Range.Characters(4).start _
                , a.Range.End
            slRng.font.Size = 7.5
        End If
    Next a
    If Selection.Type = wdSelectionIP Then ActiveDocument.Select
    Set slRng = Selection.Range
    For Each a In slRng.Characters
        Select Case a.font.Color
            Case 34816, 8912896, 15776152 '34816:ºñ¦â¤pª`
p:              If flg = False Then
                    a.Select
                    Selection.Range.InsertBefore "¡]"
                    Selection.Range.SetRange Selection.start, Selection.start + 1
                    Selection.Range.font.Size = a.Characters(2).font.Size
                    Selection.Range.font.Color = a.Characters(2).font.Color
    '                a.Font.Size = a.Next.Font.Size
    '                a.Font.Color = a.Next.Font.Color
                    flg = True
                Else
                    If a.font.Color = 8912896 And a.Previous.font.Color = 34816 Then '8912896ÂÅ¦r¤pª`
                        a.InsertBefore "¡^¡]"
                        a.SetRange a.start, a.start + 2
                        a.font.Size = a.Characters(2).Next.font.Size
                        a.font.Color = a.Characters(2).Next.font.Color
    '                    a.Characters(1).Font.Color = a.Characters(1).Previous.Font.Color
                    End If
                End If
    '        Case 8912896 '8912896ÂÅ¦r¤pª`
                
            Case 0, 15595002, 15649962
                If a.font.Color = 0 Then 'black'º~Äy¹q¤l¤åÄm¸ê®Æ®w
                    If a.font.Size = 7.5 And Not flg Then
                        GoTo p
                    ElseIf a.font.Size > 7.5 And flg Then
                        GoTo b
                    End If
                'End If
                ElseIf flg Then
b:
    '                a.Select
    '                Selection.Range.InsertBefore "¡^"
                    If a.Previous = VBA.Chr(13) Then
                        a.Previous.Previous.Select
                    Else
                        a.Previous.Select
                    End If
                    Selection.Range.InsertAfter "¡^"
                    flg = False
                End If
            Case -16777216 'black'º~Äy¹q¤l¤åÄm¸ê®Æ®w
                If a.font.Size = 7.5 And Not flg Then
                    GoTo p
                ElseIf a.font.Size > 7.5 And flg Then
                    GoTo b
                End If
            Case 255 'red'º~Äy¹q¤l¤åÄm¸ê®Æ®w
                Select Case a.font.Size
                    Case 7.5, 10
                        a.Delete
                End Select
        End Select
    Next a
    slRng.Find.Execute "¡]¡]", True, , , , , , , , "¡]", wdReplaceAll
    slRng.Find.Execute "¡^¡^", True, , , , , , , , "¡^", wdReplaceAll
    Beep
    Selection.EndKey wdStory
    Do
       Selection.MoveLeft
       If Selection = VBA.Chr(13) Then Selection.Delete
    Loop While Selection = VBA.Chr(13)
    'MsgBox "done!", vbInformation
    If Not ActiveDocument Is d Then d.Activate
    SystemSetup.contiUndo ur
End Sub
Sub º~Äy¹q¤l¤åÄm¸ê®Æ®w¤å¥»¾ã²z_¥HÂà¶K¨ì¤¤°ê­õ¾Ç®Ñ¹q¤l¤Æ­p¹º(Optional doNotCloseDoc As Boolean, Optional openNewDoc As Boolean = False)
    Dim rng As Range, d As Document, a, ur As UndoRecord
    Dim rp As Variant, i As Byte
    'Set ur = SystemSetup.stopUndo("º~Äy¹q¤l¤åÄm¸ê®Æ®w¤å¥»¾ã²z_¥HÂà¶K¨ì¤¤°ê­õ¾Ç®Ñ¹q¤l¤Æ­p¹º")
    SystemSetup.stopUndo ur, "º~Äy¹q¤l¤åÄm¸ê®Æ®w¤å¥»¾ã²z_¥HÂà¶K¨ì¤¤°ê­õ¾Ç®Ñ¹q¤l¤Æ­p¹º"
    If Documents.Count = 0 Then Documents.Add
    Set d = ActiveDocument
    If (d.path <> "" Or d.Content.text <> VBA.Chr(13)) And openNewDoc Then
        Set d = Documents.Add()
        'Exit Sub
    End If
    rp = Array("(", "{{", ")", "}}", VBA.ChrW(160), "", "¡i¹Ï¡j", "", _
         "^p^p", "^p", _
         VBA.ChrW(13) & VBA.ChrW(45) & VBA.ChrW(13) & VBA.ChrW(13) & VBA.ChrW(11), "^p", _
         VBA.ChrW(13) & VBA.ChrW(45) & VBA.ChrW(13), "^p", "{{ }}", "", "[", VBA.ChrW(12310), _
         "]", VBA.ChrW(12311), " ", "", "¡³", VBA.ChrW(12295), _
         "^p" & VBA.ChrW(12310) & "²¨" & VBA.ChrW(12311), VBA.ChrW(12310) & "²¨" & VBA.ChrW(12311) & "{{", _
         "}}" & VBA.Chr(13) & "^#" & VBA.Chr(13) & "{{", "", _
         "¡D¡D¡D¡D¡D¡D¡D¡D¡D¡D¡D¡D¡D¡D¡D¡D¡D¡D" & VBA.Chr(13), "", _
         VBA.Chr(13) & "^#" & VBA.Chr(13), "", _
         "}}" & VBA.Chr(13) & "^#" & VBA.Chr(13), "}}", _
         "}}" & VBA.Chr(13) & "{{", "", _
         "-", "", "^#", "", "¡C¡C", "¡C") ', "¡C}}<p>¡C}}<p>", "¡C}}<p>")
         '­ì¨Ó¡uvba.Chrw(13) & vba.Chrw(45) & vba.Chrw(13) & vba.Chrw(13) & vba.Chrw(11)¡v¬O¨ä¤¤¦³ªí®æ°Ú
    Set rng = d.Range
    If rng.Characters.Count = 1 Then '¤å¥óµL¤º®e®É¤~¶K¤W
        rng.Paste '¡mº~Äy¥þ¤å¸ê®Æ®w¡n§ïª©«á°Å¶KÃ¯®Ä¯à«Ü®t¡A­n¾¨¶qÁ×§K ·P®¦·P®¦¡@Æg¼ÛÆg¼Û¡@«nµLªüÀ±ªû¦ò 20240825
    End If
    If Not rng.Document Is ActiveDocument Then
        rng.Document.Activate '¨Ñµ¹«á­±ªº¨ç¦¡¥Î
    End If
    º~Äy¹q¤l¤åÄm¸ê®Æ®w¤å¥»¾ã²z_ª`¤å«e«á¥[¬A¸¹_origin
    'º~Äy¹q¤l¤åÄm¸ê®Æ®w¤å¥»¾ã²z_ª`¤å«e«á¥[¬A¸¹
    For Each a In rng.Characters
            If a.font.Size = 10 Then
            Select Case a.font.Color
                Case 255, 9915136
                    a.Delete
            End Select
        End If
    Next a
'    rng.Cut
    On Error GoTo eH:
'    rng.PasteAndFormat wdFormatPlainText
    If rng.Find.Execute("/") Then Stop
     
     '¡mº~Äy¥þ¤å¸ê®Æ®w¡n§ïª©«á°Å¶KÃ¯®Ä¯à«Ü®t¡A³y¦¨µ{§Ç·í±¼¡A§ï¼g¦¨³o¦æ«á´N¸Ñ¨M¤F¡I¥i¨£°ÝÃD¥X¦b°Å¶KÃ¯¡A«o¤£¬O§Ú¼gªºµ{¦¡¡C·P®¦·P®¦¡@Æg¼ÛÆg¼Û¡@«nµLªüÀ±ªû¦ò 20240825
    rng.text = VBA.Replace(VBA.Replace(rng.text, "(/)", vbNullString), "/", vbNullString)
'    rng.text = VBA.Replace(rng.text, "(/)", vbNullString)
'    rng.text = VBA.Replace(rng.text, "/", vbNullString)
    
    If rng.Find.Execute("/") Then Stop
    rng.Find.ClearFormatting
    For i = 0 To UBound(rp)
        If InStr(rng.text, rp(i)) > 0 Then
            rng.Find.Execute rp(i), , , , , , , wdFindContinue, , rp(i + 1), wdReplaceAll
        End If
        i = i + 1
    Next i
    ¤¤°ê­õ¾Ç®Ñ¹q¤l¤Æ­p¹º.ºû°ò¤å®wµ¥±ýª½±µ©â´«¤§¦r d
    ¤å¦r³B²z.®Ñ¦W¸¹½g¦W¸¹¼Ðª`
    Beep
'    SystemSetup.playSound 1.294
    If Not doNotCloseDoc Then
        d.Range.Cut '³o¸Ì¤w¬O¯Â¤å¦rªº¤å¥»¡A©Ò¥H¹ï©ó°Å¶KÃ¯ªº»Ý¨D±ø¥óÁÙ¦n
        d.Close wdDoNotSaveChanges
    End If
    SystemSetup.contiUndo ur
    Exit Sub
eH:
    Select Case Err.Number
        Case 4198 '«ü¥O¥¢±Ñ
            SystemSetup.wait 900
            Resume
        Case Else
            MsgBox Err.Number + Err.Description
    End Select
End Sub
Sub º~Äy¹q¤l¤åÄm¸ê®Æ®w¤å¥»¾ã²z_ª`¤å«e«á¥[¬A¸¹() '³Ì«á°õ¦æ Docs.mark©ö¾ÇÃöÁä¦r(¦b¡u©ö¾ÇÂøµÛ¤å¥»¡v¸ô®|¤U®É¡^
    Rem Alt + Shift + `
    Dim rng As Range, ur As UndoRecord, d As Document, fontsize(1) As Single, fsz
    SystemSetup.playSound 0.484
    Set d = ActiveDocument '°O¤U²{¥Î¤å¥ó¡]­n¶K¤Wµ²ªGªº¥Øªº¦a¤å¥ó¡^
    Set rng = d.Range
    
    '¥ýÀË¬d°Å¶KÃ¯¸Ì«e25¦r¤¸¡A­Y¤å¥ó¤¤¦³¡A´N¥ý¿ï¨ú
    With rng.Find
        .ClearAllFuzzyOptions
        .ClearFormatting
        '255¬O¤W­­¡A¦ý¤]¥i¯à¥]§t¤F¼ÐÂIÂ_¥y²§¤å¦Ó¾É­P¤å¥»¦³²§¤£¯à¤ñ¹ï¡AÁÙ¤£¦pÁY´î¨ì¥i¥HÃÑ§Oªºªø«×§Y¥i
        'If .Execute(VBA.Trim(VBA.Left(SystemSetup.GetClipboard, 255)), , , , , , True, wdFindContinue) Then
        If .Execute(VBA.Trim(VBA.Left(SystemSetup.GetClipboard, 25)), , , , , , True, wdFindContinue) Then
            rng.Select
            rng.Document.ActiveWindow.ScrollIntoView rng, True
            SystemSetup.playSound 2
            Exit Sub
        End If
    End With
    '«e25¦r¤¸ÀË¬d³q¹L¦AÄ~Äò
        
    SystemSetup.stopUndo ur, "º~Äy¹q¤l¤åÄm¸ê®Æ®w¤å¥»¾ã²z_ª`¤å«e«á¥[¬A¸¹"
    word.Application.ScreenUpdating = False
   
    'rng.Document ­n³B²z¡mº~Äy¥þ¤å¸ê®Æ®w¡n¤å¥»ªº¤å¥ó
    If rng.Document.path = "" Then
        If rng.text = "" Then rng.Paste
    Else
        Set rng = Docs.ªÅ¥Õªº·s¤å¥ó(False).Range 'False¡G ¤£Åã¥Ü·s¤å¥ó
    End If
    
    rng.Paste
    
'    If VBA.InStr(rng.text, "¡D¡@¡D¡@¡D¡@¡D¡@¡D¡@¡D¡@¡D¡@¡D¡@¡D¡@¡D¡@¡D¡@¡D¡@¡D¡@¡D¡@¡D¡@¡D¡@¡D¡@¡D") Then Stop 'just for test
    
    '²M°£¥b§ÎªÅ®æ
    If VBA.InStr(rng.text, " ") Then rng.Find.Execute " ", , , , , , , wdFindContinue, , vbNullString, wdReplaceAll
    'ª`¤å«e¡B«á¥[¡]¡B¡^
    fontsize(0) = 10: fontsize(1) = 7.5
    For Each fsz In fontsize
        With rng.Find
            .ClearFormatting
            .font.Size = fsz
            .Forward = True
            .Wrap = wdFindStop
        End With
        Do While rng.Find.Execute()
            If rng.font.Size <> fsz Then Exit Do
            If rng.Characters(1) = "¡]" Then
                Exit Do
            End If
            If rng.End = rng.Document.Range.End Then rng.End = rng.End - 1
            Do While rng.Characters(rng.Characters.Count) = VBA.Chr(13)
                rng.End = rng.End - 1
            Loop
            Do While rng.Characters(1) = VBA.Chr(13) & VBA.Chr(7)
                rng.start = rng.start + 1
            Loop
            rng.text = "¡]" & rng.text & "¡^"
        Loop
        'rng ¦ì¸m¤w²¾°Ê¡A¦b§ä¤U¤@­Ó±ø¥ó«e¶·«ì´_
        Set rng = rng.Document.Range
    Next fsz
            
    Set rng = rng.Document.Range
    rng.Find.ClearFormatting
    If InStr(rng.text, "¡^¡]¡C¡^") Then
        With rng.Find
            .text = "¡^¡]¡C¡^"
            .Replacement.text = "¡C¡^"
            .Execute , , , , , , , wdFindContinue, , , wdReplaceAll
        End With
    End If
    Dim str_to_clear() As String
    ReDim str_to_clear(2)
    str_to_clear(0) = "¡^¡]"
    str_to_clear(1) = "¡i¹Ï¡j" & VBA.Chr(11) ' ^l
    str_to_clear(2) = "¡i¹Ï¡j"
    Set rng = rng.Document.Range
    rng.Find.ClearFormatting
    For Each fsz In str_to_clear
        If InStr(rng.text, fsz) Then
            With rng.Find
                .text = fsz
                .Replacement.text = vbNullString
                .Execute , , , , , , , wdFindContinue, , , wdReplaceAll
            End With
        End If
    Next fsz
    
    SystemSetup.contiUndo ur
        
    '½Æ»s¾ã²z¦nªºµ²ªG¡A·Ç³Æ°e¥æmark©ö¾ÇÃöÁä¦r¤ñ¹ï
    rng.Document.Range.Copy
    Dim dt As Date
    dt = VBA.Now '°Å¶KÃ¯¦ü¥G·|¤ÏÀ³¤£¹L¨Ó¡A¬G¶·µ¥µ¥
    Do While DateDiff("s", dt, VBA.Now) < 2
        DoEvents
    Loop
    
    
    Dim yi As Boolean
    If VBA.InStr(d.path, "©ö¾ÇÂøµÛ¤å¥»") Then yi = True
    If yi Then
        DoEvents
        d.Activate
        DoEvents
        
        Dim dx As String, punct, ePunct, hasPunct As Boolean
        dx = rng.Document.Range.text
        punct = Array("¡C", "¡A", "¡B", "¡F", "¡u¡v", "¡y¡z", "¡I", "¡H", "¡E") '¡u¡E¡v¡G¦p¡mº~Äy¥þ¤å¸ê®Æ®w¡P®ø®L¶~°OºK§Û¡n
        For Each ePunct In punct
            If VBA.InStr(dx, ePunct) Then
                hasPunct = True
                Exit For
            End If
        Next ePunct
        
        Dim pasteAppendedRange As Range
        Set pasteAppendedRange = d.Range
        pasteAppendedRange.start = d.Range.End
        '¨O­«¶K¤W¤º®e³¡¤À¥Ñ mark©ö¾ÇÃöÁä¦r ­t³d¡A¼ÐÃÑÃöÁä¦r«h¥æ¥Ñ°e¥æ¦Û°Ê¼ÐÂI¤º§tªº©I¥s mark©ö¾ÇÃöÁä¦r ¨Ó³B²z
        If hasPunct Then
            Docs.mark©ö¾ÇÃöÁä¦r Nothing, False
            
            GoSub chkRng
            
            If rng.Document.path = vbNullString Then
                rng.Document.Close wdDoNotSaveChanges
            Else
                DoEvents
                d.Activate
            End If
            
        Else
            If Docs.mark©ö¾ÇÃöÁä¦r(Nothing, hasPunct) Then
            '¶K¤W¤§«á¥Ñ¨ä¶K¨ì¤å¥ó¥½ºÝ¡B¤S¹w¯d¤@¨Ç¤À¬q²Å¸¹¦¹¤@¯S¼x¡A¥i¥H§ì¨ì¶K¤WªºRange
                pasteAppendedRange.End = d.Range.End
                pasteAppendedRange.Select '¦]¬°°e¥æ¦Û°Ê¼ÐÂIµ{§Ç¤º¦³ Selection.Cut
                
'                Stop 'just for test
                Network.Åª¤J¥jÄy»Å¦Û°Ê¼ÐÂIµ²ªG
                Docs.marking©ö¾ÇÃöÁä¦r pasteAppendedRange, Keywords.©ö¾ÇKeywords_ToMark
                SystemSetup.playSound 2
                Rem ¥H¤U¬°ÂÂ¦¡
                '¥H¤Uµ{§Ç¤º¤w¦³ mark©ö¾ÇÃöÁä¦r ¬G
                
'                If TextForCtext.TextForCtextExist Then
'                    pasteAppendedRange.Cut
'                    DoEvents
'                    'SystemSetup.wait
'                    If TextForCtext.GjcoolPunct() Then
'                        '¦]¬°¦b¤å¥óªº³Ì¥½ºÝ
'                        pasteAppendedRange.SetRange pasteAppendedRange.End - 1, pasteAppendedRange.End - 1
'                        pasteAppendedRange.text = SystemSetup.GetClipboard
'                        DoEvents
'                        SystemSetup.wait 0.1
'                        Docs.marking©ö¾ÇÃöÁä¦r pasteAppendedRange, Keywords.©ö¾ÇKeywords_ToMark
'                        SystemSetup.playSound 2
'                    Else
'                        word.Application.Activate
'                        MsgBox "µo¥Í¿ù»~¡A½Ð­«¸Õ¡I", vbCritical
'                        GoTo finish
'                    End If
'                Else
'                    Docs.¤¤°ê­õ¾Ç®Ñ¹q¤l¤Æ­p¹º_¥u«O¯d¥¿¤åª`¤å_¥Bª`¤å«e«á¥[¬A©·_¶K¨ì¥jÄy»Å¦Û°Ê¼ÐÂI
'                End If
                Rem ¥H¤W¬°ÂÂ¦¡

                pasteAppendedRange.InsertParagraphAfter
                pasteAppendedRange.InsertParagraphAfter
                pasteAppendedRange.InsertParagraphAfter
                pasteAppendedRange.InsertParagraphAfter
            End If
        End If
        DoEvents
        rng.Document.Close wdDoNotSaveChanges
    Else
        DoEvents
        rng.Document.Activate
    End If
finish:
    word.Application.ScreenUpdating = True
    word.Application.Activate
'    AppActivate word.ActiveDocument.ActiveWindow.Caption
Exit Sub

chkRng:
    'creedit_with_Copilot¤jµÐÂÄ¡G20240918 https://sl.bing.net/btKqsDvwrD2
    Dim r As Range
    On Error Resume Next
    Set r = rng
    On Error GoTo 0

    If r Is Nothing Then
        '¤w³Q§R°£©Î¤£¦s¦b¡C
    Else
        '¦s¦b¡C
    End If
    GoTo finish 'Return
End Sub

Sub º~Äy¹q¤l¤åÄm¸ê®Æ®w¤å¥»¾ã²z_ª`¤å«e«á¥[¬A¸¹_origin()
'Sub º~Äy¹q¤l¤åÄm¸ê®Æ®w¤å¥»¾ã²z_ª`¤å«e«á¥[¬A¸¹()
    
    Dim rng As Range, fColor As Long, flg As Boolean, ur As UndoRecord
    Const fSize As Byte = 10
      
    Set rng = ActiveDocument.Range
    rng.Collapse wdCollapseStart
    fColor = rng.font.Color
    Do While rng.End < rng.Document.Range.End - 1
    
        rng.Move wdCharacter, 1
        
        If rng.font.Color = 204 And rng.font.Size = 11 Then
            rng.Delete
        ElseIf rng.font.Color = 0 And rng.font.Size = 7.5 Then
            GoTo mark
        ElseIf (rng.font.Color <> fColor Or rng.font.Size = fSize) And _
                    (rng.font.Color <> 234 And rng.font.Bold = False) Then '¬õ¦r+²ÊÅé¬°ÀË¯Áµ²ªG
mark:
            If flg = False Then
                If rng.font.Color <> -16777216 Then
                    rng.InsertBefore "("
                    rng.Characters(1).font.Color = rng.Next.Next.font.Color
                    rng.Characters(1).font.Size = rng.Next.Next.font.Size
                    flg = True
                End If
            End If
        ElseIf rng.font.Color = fColor And flg = True Then
            rng.Previous.InsertAfter ")"
            flg = False
        End If
        
    Loop
    Beep
    'SystemSetup.playSound 1.294
End Sub
Sub ¸Ö¥y¤À¦æ()
Dim slRng As Range, a
Set slRng = Selection.Range
For Each a In slRng.Characters
    If a Like "[¡C¡A¡F¡H¡I¡u¡v¡y¡z]" Then
        a.Select
        Selection.Move
        Selection.TypeText VBA.Chr(11)
    End If
Next a
End Sub

Sub §R°£®Õ®×»y()
Dim rng As Range, e, d As Document
Set d = ActiveDocument
Set rng = d.Range
e = rng.End
With rng.Find
    .Style = "¶W³sµ²"
    .Execute , , , , , , , wdFindStop ', , "" ', wdReplaceAll
    Do
        If InStr(rng.Characters(rng.Characters.Count).Next.Style, "®Õ®×") _
            Or InStr(rng.Characters(1).Previous.Style, "®Õ®×") Then
            rng.Select
            Selection.Delete
            rng.SetRange Selection.start, e
        End If
    Loop While .Execute(, , , , , , , wdFindStop)  ', , "" ', wdReplaceAll
End With

With rng.Find
    .Style = "®Õ®×"
    .Execute , , , , , , , wdFindContinue, , "", wdReplaceAll
End With
With rng.Find
    .Style = "®Õ®×¤Þ¤å"
    .Execute , , , , , , , wdFindContinue, , "", wdReplaceAll
End With
Beep
End Sub

Function °ê»yÃã¨åª`­µ¤å¦r³B²z(x As String)
Dim ay, i As Byte
ay = Array("£¸", VBA.ChrW(20008), "¡@", " ", "¡]¤S­µ¡^", "¤S­µ ", "¡]Åª­µ¡^", "Åª­µ ", "¡]»y­µ¡^", "»y­µ ", _
        "(¤@)", "", "(¤G)", "", "(¤T)", "", "(¥|)", "", "(¤­)", "", "(¤»)", "", "¡^", "", "¡]", "")
For i = 0 To UBound(ay)
    x = Replace(x, ay(i), ay(i + 1))
    i = i + 1
Next i
°ê»yÃã¨åª`­µ¤å¦r³B²z = x
End Function
Sub ¥ÍÃø¦r¥[¤W°ê»yÃã¨åª`­µ()
Dim rng As Range, x, rst As New ADODB.Recordset, st As WdSelectionType, words As String
Dim cnt As New ADODB.Connection, id As Long, sty As word.Style, url As String
Dim frmDict As New Form_DictsURL, lnks As New Links, db As New dBase ', frm As New MSForms.DataObject
Static cntStr As String, chromePath As String
st = Selection.Type
If st = wdSelectionIP Then
    If Selection.start = 0 Then Exit Sub
    x = Selection.Previous.Characters(Selection.Previous.Characters.Count).text
    If InStr("¡C¡A¡F¡u¡v¡y¡z¡q¡r¡m¡n¡H.,;""?¡Ð-¢w¢w--¡]¡^()¡i¡j¡e¡f<>[]¡K! ¡@¡I", x) Then Exit Sub
'    Selection.Previous.Copy
Else
    x = trimStrForSearch(VBA.CStr(Selection.text), Selection)
    'Selection.Copy
    SystemSetup.ClipboardPutIn "=" & Selection.text
End If
    If ¤å¦r³B²z.isSymbol(CStr(x)) Or ¤å¦r³B²z.isª`­µ²Å¸¹(CStr(x)) Or ¤å¦r³B²z.isLetter(CStr(x)) Or ¤å¦r³B²z.isNum(CStr(x)) Then Exit Sub
Set rng = Selection.Range
words = x
db.setWordControlValue (words)
On Error GoTo eH
Dim ur As UndoRecord
'Set ur = SystemSetup.stopUndo("¥ÍÃø¦r¥[¤W°ê»yÃã¨åª`­µ")
SystemSetup.stopUndo ur, "¥ÍÃø¦r¥[¤W°ê»yÃã¨åª`­µ"

'If Not Selection.Document.path = "" Then If Not Selection.Document.Saved Then Selection.Document.save
If cntStr = "" Then
    Dim dbp As New Paths
    cntStr = dbp.getdb_­«½s°ê»yÃã¨å­×­q¥»_¸ê®Æ®wfullName
End If

If chromePath = "" Then
    chromePath = SystemSetup.getChrome
End If

'Dim ay, i As Byte
'ay = Array("£¸", vba.Chrw(20008), "¡@", " ", "¡]¤S­µ¡^", "¤S­µ ", "¡]Åª­µ¡^", "Åª­µ ", "¡]»y­µ¡^", "»y­µ ", _
'        "(¤@)", "", "(¤G)", "", "(¤T)", "", "(¥|)", "", "(¤­)", "", "(¤»)", "", "¡^", "", "¡]", "")

    cnt.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & cntStr
'Exit Sub
'cntt:
    rst.Open "select ª`­µ¤@¦¡,ÄÀ¸q,url,ID,¦h­µ±Æ§Ç from [¡m­«½s°ê»yÃã¨å­×­q¥»¡n Á`ªí] where strcomp(¦rµü¦W,""" & x & """)=0 order by ¦h­µ±Æ§Ç", cnt, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        GoSub list
    Else
        ¥ÍÃø¦r¥[¤W°ê»yÃã¨åª`­µnextTable rst, cnt, x, "¡m­«½s°ê»yÃã¨å­×­q¥»¡n Á`ªí-20210928¥H«e", True
        If rst.RecordCount > 0 Then
            GoSub list
        Else
2
            If Selection.Characters.Count = 1 Then 'words  ³æ¦r
                frmDict.getDictVariantsRecS words, rst
                If rst.RecordCount > 0 Then
                    GoSub list
                Else
                    frmDict.getDictHydzdRecS words, rst
                    If rst.RecordCount > 0 Then
                        'GoSub list
                        Set sty = rng.Style
                        rng.Hyperlinks.Add rng, lnks.trimLinks(rst.Fields(2).Value), , , , "_blank"
                        lnks.setStylewithHyperlinkMark sty, rng
                    Else
                        GoSub notFound
                    End If
                End If
            Else 'terms µü·J
                frmDict.getDictHydcdRecS words, rst
                If rst.RecordCount > 0 Then
                    If Not VBA.IsNull(rst.Fields(0)) Then
                        GoSub list
                    Else
                        Set sty = rng.Style
                        rng.Hyperlinks.Add rng, lnks.trimLinks(rst.Fields(2).Value), , , , "_blank"
                        lnks.setStylewithHyperlinkMark sty, rng
                    End If
                Else
                    GoSub notFound
                End If
            End If
        End If
    End If
endS:
    SystemSetup.contiUndo ur
    Set ur = Nothing
    If rst.State <> adStateClosed Then rst.Close
    If cnt.State <> adStateClosed Then cnt.Close
    Set rst = Nothing: Set cnt = Nothing: Set frmDict = Nothing ': Set frm = Nothing
    Set lnks = Nothing: Set db = Nothing: Set rng = Nothing
Exit Sub

notFound:
                If st = wdSelectionIP Then
                    Selection.Previous.Copy
                    'Selection.Document.FollowHyperlink "https://dict.variants.moe.edu.tw/variants/rbt/query_by_standard_tiles.rbt?command=clear"
                    x = frmDict.add1URLTo1²§Åé¦r¦r¨å(words)
                    If x = "" Then GoTo endS
                    GoTo 2
                Else
                    rst.Close
                    rst.Open "select ª`­µ¤@¦¡,ÄÀ¸q,url,ID,¦h­µ±Æ§Ç from [¡m­«½s°ê»yÃã¨å­×­q¥»¡n Á`ªí] where instr(¦rµü¦W,""" & x & """)>0 order by ¦h­µ±Æ§Ç", cnt, adOpenKeyset, adLockOptimistic
                    Selection.Copy
                    If rst.RecordCount > 0 Then
                        Beep
                        'Selection.Document.FollowHyperlink "https://www.zdic.net/hans/" & x, , True
                        Shell chromePath & " https://www.zdic.net/hans/" & x
                        GoSub list
                        'Selection.Document.FollowHyperlink "http://dict.revised.moe.edu.tw/cbdic/search.htm", , True
'                        Shell chromePath & " http://dict.revised.moe.edu.tw/cbdic/search.htm"
                    Else
                            ¥ÍÃø¦r¥[¤W°ê»yÃã¨åª`­µnextTable rst, cnt, x, "¡m­«½s°ê»yÃã¨å­×­q¥»¡n Á`ªí-20210928¥H«e", False
                            If rst.RecordCount > 0 Then
                                Beep
                                GoSub list
                            Else
                            'Selection.Document.FollowHyperlink "https://www.zdic.net/hans/" & x, , True
                            Shell chromePath & " https://www.zdic.net/hans/" & x
                            End If
                    End If
                End If
Return

list:
'        Dim ur As UndoRecord
'        Set ur = SystemSetup.stopUndo("¨Kºa«ö")
'        Docs.¼Ë¦¡add_¨Kºa«öµ¥¼Ë¦¡
        rng.Collapse wdCollapseEnd
        If rng.Style <> "¨Kºa«ö" Then
            rng.InsertAfter "¡]¡^"
            rng.Style = "¨Kºa«ö"
            rng.SetRange rng.End - 1, rng.End - 1
        End If
        Do Until rst.EOF
            x = ""
            If VBA.IsNull(rst.Fields(0).Value) Then
                x = rst.Fields(1).Value 'ÄÀ¸q
            Else
                x = rst.Fields(0).Value 'ª`­µ
            End If
            GoSub typeTexts
            rst.MoveNext
        Loop
        If rng.Previous = "¡A" Then rng.Previous.Delete
'        SystemSetup.contiUndo ur
'        Set ur = Nothing:  'Set frm = Nothing: Set frmDict = Nothing

Return

typeTexts:
        If x = "" Or VBA.IsNull(x) Then GoTo 2
'        X = VBA.Mid(X, 1, Len(X) - 1)
        x = °ê»yÃã¨åª`­µ¤å¦r³B²z(CStr(x))
'        If sT <> wdSelectionIP Then
'            rng.SetRange Selection.End, Selection.End
'        End If
'        rng.SetRange rng.End - 1, rng.End - 1
        rng.InsertAfter x 'insert ZhuYin
        For Each x In rng.Characters 'format ZhuYin
            If InStr("£½£¾£¿", x) Then
                x.Style = "Án½Õ"
            ElseIf InStr("£»", x) Then
                x.font.Name = "¼Ð·¢Åé"
            End If
        Next x
        x = rst.Fields(2).Value 'URL  'frmDict.get1URLfor1(words)
        If VBA.IsNull(x) Then
                If st = wdSelectionIP Then
                    If Selection.Previous.Characters(Selection.Previous.Characters.Count).Hyperlinks.Count > 0 Then
                        Dim rngW As Range
                        Set rngW = Selection.Range
                        rngW.SetRange Selection.Previous.Characters(Selection.Previous.Characters.Count).start, Selection.Previous.Characters(Selection.Previous.Characters.Count).End
                        SystemSetup.ClipboardPutIn "=" & rngW.text '"^" & rngW.text & "$" 'version 6's new settings
                        Set rngW = Nothing
                    Else
                        Set rngW = Selection.Previous.Characters(Selection.Previous.Characters.Count)
                        SystemSetup.ClipboardPutIn "=" & rngW.text
                        'Selection.Previous.Characters(Selection.Previous.Characters.Count).Copy
                    End If
                End If
'                Shell chromePath & " http://dict.revised.moe.edu.tw/cbdic/search.htm"
'            frm.Clear
'            frm.SetText words, 1
'            frm.PutInClipboard
            'add new url
            Dim repeated As Boolean 'ÀË¯Áµ²ªG¤£¤î¤@­Ó®É·|­«½Æ
rePt:
            If repeated = False Then x = SeleniumOP.grabDictRevisedUrl_OnlyOneResult(words) '¦¹ªk¥u¾A¥Î©ó¶È1µ§¸ê®Æ®É,¨S¦³©Î¦h©ó1µ§«hªð¦^""ªÅ¦r¦ê
            rng.Document.ActiveWindow.Application.Activate
            If rst.RecordCount = 1 Then '°ê»yÃã¨å¸ê®Æ®w¸Ì¥u¦³1µ§§k¦X¸ê®Æ
                If x = "" Then 'µ²ªG¤£¤î1­Ó®É
                    Shell Network.getDefaultBrowserFullname & " https://dict.revised.moe.edu.tw/search.jsp?md=1"
                End If
''                If Not SystemSetup.appActivatedYet("chrome") Then
''                'If Not word.Tasks.Exists("google chrome") Then
''                    Shell SystemSetup.getChrome & " https://dict.revised.moe.edu.tw/search.jsp?md=1"
''                Else
''                    SystemSetup.appActivateChrome
''                End If
            Else
                If VBA.IsNull(x) Then x = ""
                Beep
            End If
            If x = "" Then 'µ²ªG¤£¤î1­Ó®É
                If repeated = False Then
                    If SeleniumOP.ActiveXComponentsCanNotBeCreated Then
                        SystemSetup.playSound 2
                        Shell Network.getDefaultBrowserFullname & " https://dict.revised.moe.edu.tw/search.jsp?md=1&word=" & words
                    Else
                        Shell Network.getDefaultBrowserFullname & " https://dict.revised.moe.edu.tw/search.jsp?md=1"
                    End If
                Else
                
                End If
                x = InputBox("plz putin the url", , IIf(VBA.IsNull(rst.Fields(0).Value), "", rst.Fields(0).Value)) 'frmDict.add1URLTo1°ê»yÃã¨å(words)
                If repeated Then
                    SystemSetup.wait 1 '¥ý½T©w­n¿é¤J­þ­Óµü±ø¡A¦A±NÂsÄý¾¹¸m«e
                    appActivateChrome
                End If
                repeated = True
            End If
            If x = "" Then GoTo endS
            If VBA.Left(x, 4) <> "http" Then GoTo rePt
            x = lnks.trimLinks_http_Dicts_toAddZhuYin_RevisedMoeEdu(CStr(x), rst.Fields(0))
            url = VBA.CStr(x)
            If lnks.chkLinks_http_Dicts_toAddZhuYin(url, words, 1, id, rst.Fields(0)) Then
                x = url
                rst.Fields(2).Value = x
                If id <> 0 Then
                    rst.Fields("ID") = id
                    id = 0
                End If
                rst.Update
                '¥H¤U¥ý²¤¥h
                '¦b¬d¦rforInPut¸ê®Æ®wªºªí³æ¤¤³]©w±±¨î¶µªº­È
                'db.setURLControlValue VBA.CStr(x)
            Else
                GoTo endS 'Exit Sub
            End If
        End If
        Set sty = rng.Style
        rng.Hyperlinks.Add rng, lnks.trimLinks(VBA.CStr(x)), , , , "_blank"
        lnks.setStylewithHyperlinkMark sty, rng
        rng.Collapse wdCollapseEnd
        'rng.SetRange rng.End, rng.End
        rng.Next.InsertBefore "¡A"
'        rng.Style = "¨Kºa«ö"
        'rng.Hyperlinks.Item(1).Delete
        'rng.Collapse wdCollapseEnd
        rng.SetRange rng.End + 2, rng.End + 2
Return


eH:
    Select Case Err.Number
        Case 4198 '«ü¥O¥¢±Ñ 'Google Driveªº°ÝÃD
            Resume Next
        Case 5834 '«ü©w¦WºÙªº¶µ¥Ø¤£¦s¦b
            Docs.¼Ë¦¡add_¨Kºa«öµ¥¼Ë¦¡
            Resume
        Case 5 'µ{§Ç©I¥s©Î¤Þ¼Æ¤£¥¿½T
            SystemSetup.wait 3 'http://vbcity.com/forums/t/81315.aspx
            'Application.Wait (Now + TimeValue("0:00:10")) '<~~ Waits ten seconds.
            Resume 'https://stackoverflow.com/questions/21937053/appactivate-to-return-to-excel
        Case Else
            If Err.Number = -2146233088 And Err.Description = "A exception with a null response was thrown sending an HTTP request to the remote WebDriver server for URL http://localhost:9674/session/a3f86609a0c2a502c01fe9cdbef88686/window. The status of the exception was ConnectFailure, and the message was: µLªk³s±µ¦Ü»·ºÝ¦øªA¾¹" Then
                SystemSetup.killchromedriverFromHere
            Else
                MsgBox Err.Number & Err.Description
            End If
'            Resume
            GoTo endS
            'If cnt.State <> adStateClosed Then cnt.Close
    End Select
End Sub

Sub ¥ÍÃø¦r¥[¤W°ê»yÃã¨åª`­µnextTable(ByRef rst As ADODB.Recordset, ByRef cnt As ADODB.Connection, x, tbName As String, precise As Boolean)
    If rst.State = adStateOpen Then rst.Close
    Dim src As String
    Dim srcs As String
    srcs = "select ª`­µ¤@¦¡,ÄÀ¸q,url,ID from [" & tbName & "] where "
    If precise Then
        src = "strcomp(¦rµü¦W,""" & x & """)=0"
    Else
        src = "instr(¦rµü¦W,""" & x & """)>0"
    End If
    rst.Open srcs & src, cnt, adOpenKeyset
End Sub

Rem «Å§i¥¢±Ñ¡G¦b¡udx = regEx.Replace(dx, rw)¡v³o¦æ·|¥X²{¡G 5017¡GÀ³¥Îµ{¦¡©Îª«¥ó©w¸q¤Wªº¿ù»~
Rem 20230309 creedit with chatGPT¤jµÐÂÄ¡G®Ñ¦W¸¹¼ÐÂI»P¥¿«hªí¹F¦¡ADO.NET¡BLINQ¡G
Sub ®Ñ¦W¸¹½g¦W¸¹¼Ðª`_¥¿«hªí¹F¦¡RegularExpression_Plaintext()
Dim cnt As New ADODB.Connection, rst As New ADODB.Recordset
Dim cntStr As String, d As Document, dx As String, w As String, rw As String
Dim db As New dBase
db.cnt¬d¦r cnt
Dim regex As Object
'Dim regEx As New RegExp
    Set regex = CreateObject("VBScript.RegExp")
Dim replacedText As String
Set d = ActiveDocument: dx = d.Range.text
rst.Open "select * from ¼ÐÂI²Å¸¹_®Ñ¦W¸¹_¦Û°Ê¥[¤W¥Î order by ±Æ§Ç", cnt, adOpenForwardOnly, adLockReadOnly
Do Until rst.EOF
    w = rst("®Ñ¦W").Value
    If VBA.InStr(dx, w) Then 'if found
        If VBA.IsNull(rst("¨ú¥N¬°").Value) Then
            rw = "¡m" & rst("®Ñ¦W").Value & "¡n"
        Else
            rw = rst("¨ú¥N¬°").Value
        End If
        With regex
            '.Pattern = "(?<!¡m)(?<!¡q)(?<![\\p{P}&&[^¡n¡r]]+)" + regEx.Escape(w) + "(?!¡n)(?!¡r)"
            .Pattern = "(?<!¡m)(?<!¡q)(?<![\\p{P}&&[^¡n¡r]]+)" + Replace(Replace(w, "\", "\\"), ".", "\.") + "(?!¡n)(?!¡r)"
            Rem ¦b Word VBA ¤¤¡ARegExp ª«¥óªº Escape ¤èªk¬O¤£³Q¤ä´©ªº¡C©Ò¥H±z»Ý­n§â³o­Ó¤èªk§ï¦¨¨Ï¥Î Replace ¤èªk±N¯S®í¦r²ÅÂà´«¬°¥¿«hªí¹F¦¡ªºÂà¸q¦r²Å¡C
            Rem ³o­Ó¡uregEx.Escape(w)¡v¨ì©³¬O¤°»ò·N¸q¡H
            Rem regEx.Escape(w) ¬O±N¦r¦ê w ¤¤©Ò¦³ªº¥¿«hªí¹F¦¡¤¸¦r²Å (¨Ò¦p *, ?, [, ], \, (, ), {, }, +, ^, $, ., |) Âà´«¦¨¯Â¤å¦r¡A¥HÁ×§K³o¨Ç¤¸¦r²Å³Q·í§@¥¿«hªí¹F¦¡ªº¤¸¯À¦Ó¥X²{¿ù»~¡C
            Rem ¨Ò¦p¡A¦pªG w ¬° test*¡A«h regEx.Escape(w) ·|¦^¶Ç test\*¡A³o¼Ë¥¿«hªí¹F¦¡¤ÞÀº´N·|§â * µø¬°¤@¯ë¦r¤¸¡A¦Ó¤£¬O¥¿«hªí¹F¦¡¤¤ªº¶qµü
            .Global = True
        End With
        dx = regex.Replace(dx, rw)
    End If
    rst.MoveNext
Loop
Documents.Add.Range.text = dx
rst.Close
'rst.Open "select * from ¼ÐÂI²Å¸¹_½g¦W¸¹_¦Û°Ê¥[¤W¥Î order by ±Æ§Ç", cnt, adOpenForwardOnly, adLockReadOnly
rst.Close: cnt.Close
Exit Sub
Rem ³o¼Ëªº¼gªk¤]¥i¥HÀ³¥Î¦bWord VBA¶Ü ¦p¦ó¹ê§@
Rem ¬Oªº¡A³o¼Ëªº¥¿«hªí¹F¦¡¼gªk¤]¥i¥H¦b Word VBA ¤¤À³¥Î¡C¥H¤U¬O¤@­ÓÂ²³æªº¥Ü¨Ò VBA µ{§Ç¡A¥i¥H±N¿ï¤¤¤å¥»¤¤ªº®Ñ¦W¸¹©M½g¦W¸¹´À´«¦¨¤¤¤åªº¥þ§Î¼ÐÂI²Å¸¹¡G
Rem
Rem vba
Rem Copy code
Rem Sub ReplacePunctuation()
Rem
Rem     Dim regEx As Object
Rem     Set regEx = CreateObject("VBScript.RegExp")
Rem
Rem     With regEx
Rem         .Pattern = "(?<!¡m)(?<!¡q)[,.;:!?()]*(¡m.*?¡n|¡q.*?¡r)[,.;:!?()]*"
Rem         .Global = True
Rem     End With
Rem
Rem     Dim selection As Range
Rem     Set selection = Selection.Range
Rem
Rem     Dim replacedText As String
Rem     replacedText = regEx.Replace(selection.Text, "¡m$1¡n")
Rem
Rem     selection.Text = replacedText
Rem
Rem End Sub
Rem ³o­Óµ{¦¡·|¨Ï¥Î»P«e­± Python µ{§Ç¬Û¦Pªº¥¿«hªí¹F¦¡¨Ó¶i¦æ´À´«¡C¥¦·|¥ý¿ï¤¤¤@¬q¤å¦r¡AµM«á±N¨ä¤¤ªº®Ñ¦W¸¹©M½g¦W¸¹´À´«¦¨¤¤¤åªº¥þ§Î¼ÐÂI²Å¸¹¡Cµ{¦¡¤¤¨Ï¥Î¤F Word VBA ¤º«Øªº Range ©M Selection ¹ï¶H¡A¥H¤Î VBScript.RegExp Ãþ§Oªº CreateObject ¤èªk¨Ó³Ð«Ø¥¿«hªí¹F¦¡¹ï¶H¡C
Rem
Rem
Rem
Rem
Rem ®]¦u¯u
Rem .Global=true ¬O¤°»ò·N«ä
Rem .Global=true ¬O¥¿«hªí¹F¦¡¹ï¶Hªº¤@­ÓÄÝ©Ê¡A¥Î©ó«ü©w¦b¤å¥»¤¤¤Ç°t©Ò¦³²Å¦X±ø¥óªº¼Ò¦¡¡A¦Ó¤£¶È¶È¬O²Ä¤@­Ó¤Ç°t¶µ¡CÀq»{±¡ªp¤U¡A¥¿«hªí¹F¦¡¥u¤Ç°t²Ä¤@­Ó²Å¦X±ø¥óªº¼Ò¦¡¡AµM«á´N°±¤î¤Ç°t¡C¦]¦¹¡A±N .Global ÄÝ©Ê³]¸m¬° true ¥i¥H«OÃÒ¤Ç°t¨ì©Ò¦³ªº¼Ò¦¡¡C
Rem
Rem ¦b¤W­±ªº¨Ò¤l¤¤¡Aregex ¹ï¶H³Q³Ð«Ø®É¡A.Global ÄÝ©Ê³Q³]¸m¬° true¡A³o¼Ë´N¥i¥H¤Ç°t¤å¥»¤¤©Ò¦³ªº¼ÐÂI²Å¸¹¤F¡C¦b Word VBA ¤¤¡A±z¥i¥H³q¹L³]¸m Match.Global ÄÝ©Ê¨Ó¹ê²{Ãþ¦üªº¥\¯à¡C¨Ò¦p¡G
Rem
Rem vbnet
Rem Copy code
Rem Sub ReplacePunctuation()
Rem
Rem     Dim regEx As New RegExp
Rem     regEx.Pattern = "\\p{P}"
Rem     regEx.Global = True
Rem
Rem     ActiveDocument.Content.Select
Rem     Selection.Find.ClearFormatting
Rem     Selection.Find.Replacement.ClearFormatting
Rem     With Selection.Find
Rem         .Text = regEx.Pattern
Rem         .Replacement.Text = "¡A"
Rem         .Forward = True
Rem         .Wrap = wdFindContinue
Rem         .Format = False
Rem         .MatchCase = False
Rem         .MatchWholeWord = False
Rem         .MatchAllWordForms = False
Rem         .MatchSoundsLike = False
Rem         .MatchWildcards = True
Rem         .Execute Replace:=wdReplaceAll
Rem     End With
Rem
Rem End Sub
Rem ¦b³o­Ó¨Ò¤l¤¤¡A§Ú­Ì³Ð«Ø¤F¤@­Ó RegExp ¹ï¶H regEx¡A¨Ã±N¨ä .Pattern ÄÝ©Ê³]¸m¬° \\p{P}¡A±N .Global ÄÝ©Ê³]¸m¬° True¡CµM«á§Ú­Ì¿ï¤¤¤åÀÉ¤¤ªº©Ò¦³¤º®e¡A°õ¦æ¬d§ä©M´À´«¾Þ§@¡A±N©Ò¦³²Å¦X¼Ò¦¡ªº¼ÐÂI²Å¸¹´À´«¦¨¤¤¤å³r¸¹¡C.MatchWildcards ÄÝ©Ê³Q³]¸m¬° True¡A³o¼Ë¦b¬d§ä®É´N¥i¥H¨Ï¥Î¥¿«hªí¹F¦¡¼Ò¦¡¡C
Rem
Rem
Rem
Rem
Rem ®]¦u¯u
Rem ©Ò¥H Find ª«¥óªº .Text ÄÝ©Ê¬O¥i¥H³]¬°¥¿«hªí¹F¦¡Åo¡H ¦ý­n±N .MatchWildcards ³]¦¨ true ´N¬O¤F¡H §_«h´N·|³Q·í§@¤@¯ë¤å¦r¬d§ä¡A¦Ó¤£¬O¥H¥¿«hªí¹F¦¡ªº³W«h¤ñ¹ïÅo¡H ¬O³o¼Ë¶Ü¡H¦Ñ®vµÐÂÄ
Rem ¬Oªº¡A±z»¡±o«D±`¥¿½T¡C¦b Word VBA ¤¤¡A­n±N Find ª«¥óªº .Text ÄÝ©Ê³]¬°¥¿«hªí¹F¦¡¡A¥²¶·¥ý±N .MatchWildcards ³]¬° true¡A§_«h Word ·|±N¨äµø¬°¤@¯ë¤å¦r¬d§ä¡C¦]¦¹¡A¦pªG±z·Q­n¨Ï¥Î¥¿«hªí¹F¦¡¶i¦æ¬d§ä©M´À´«¾Þ§@¡A½Ð°O±o±N .MatchWildcards ³]¬° true¡C
Rem
End Sub
'Function ®Ñ¦W¸¹½g¦W¸¹¼Ðª`PreExamOK(d As Document, term As String, whatMark As String, startPos_term As Long, Optional endPos_term As Long) As Boolean
Function ®Ñ¦W¸¹½g¦W¸¹¼Ðª`PreExamOK(d As Document, term As String, startPos_term As Long, Optional endPos_term As Long) As Boolean
    Dim rngChk As Range, xChk As String
    On Error GoTo eH:
    Set rngChk = d.Range(0, startPos_term)
    xChk = rngChk.text
    'If term = "¸êªv³qÅ²" Then Stop
    If InStrRev(xChk, "¡m") <= InStrRev(xChk, "¡n") And InStrRev(xChk, "¡q") <= InStrRev(xChk, "¡r") Then ®Ñ¦W¸¹½g¦W¸¹¼Ðª`PreExamOK = True
    
    Exit Function
eH:
        Select Case Err.Number
            'Case 4608 '¼Æ­È¶W¥X½d³ò
                'Resume
            Case Else
                MsgBox Err.Number + Err.Description
    '            Resume
        End Select
    
    
    'Dim result As Boolean
    'If whatMark = "¡m" Then ' = ¡H ¦p¡G¦¹®É·|¡u=¡v¡GIf InStr(xChk, "¡m") = 0 And InStr(xChk, "¡n") = 0 And InStr(xChk, "¡q") = 0 And InStr(xChk, "¡r") = 0 Then 20230312 Àu¤Æ¡C·P®¦·P®¦¡@Æg¼ÛÆg¼Û¡@«nµLªüÀ±ªû¦ò¡C¨S¦³¦òµÐÂÄ¥[«ù¡A§Ú®]¦u¯u¥i¯à¶Ü¡H
    '    If InStrRev(xChk, "¡m") <= InStrRev(xChk, "¡n") Then result = True
    'Else
    '    If InStrRev(xChk, "¡q") <= InStrRev(xChk, "¡r") Then result = True
    'End If
    
    ''«e­±³£¨S¡m¡n¡q¡r®É
    'If InStr(xChk, "¡m") = 0 And InStr(xChk, "¡n") = 0 And InStr(xChk, "¡q") = 0 And InStr(xChk, "¡r") = 0 Then
    '    result = True
    ''«e­±ªº¡m¡q¦b¡n¡rªº«e­±
    'Else
    '    'If InStrRev(xChk, "¡m") < InStrRev(xChk, "¡n") Or InStrRev(xChk, "¡q") < InStrRev(xChk, "¡r") Then result = True
    '    If whatMark = "¡m" Then
    '        If InStr(xChk, "¡m") = 0 And InStr(xChk, "¡n") = 0 Then
    '            result = True
    '        Else
    '            If InStrRev(xChk, "¡m") < InStrRev(xChk, "¡n") Then result = True
    '        End If
    '    ElseIf whatMark = "¡q" Then
    '        If InStr(xChk, "¡q") = 0 And InStr(xChk, "¡r") = 0 Then
    '            result = True
    '        Else
    '            If InStrRev(xChk, "¡q") < InStrRev(xChk, "¡r") Then result = True
    '        End If
    '    End If
    'End If
    '®Ñ¦W¸¹½g¦W¸¹¼Ðª`PreExamOK = result
End Function
Sub ®Ñ¦W¸¹½g¦W¸¹¼Ðª`()
    Dim cnt As New ADODB.Connection, rst As New ADODB.Recordset
    Dim cntStr As String, d As Document, dx As String, rngF As Range, title As String
    Dim db As New dBase
    Dim ur As UndoRecord
    On Error GoTo eH:
    SystemSetup.stopUndo ur, "®Ñ¦W¸¹½g¦W¸¹¼Ðª`"
    db.cnt¬d¦r cnt
    'If Dir("H:\§Úªº¶³ºÝµwºÐ\¨p¤H\¤d¼{¤@±oÂN(C¼Ñª©)\®ÑÄy¸ê®Æ\¹Ï®ÑºÞ²zªþ¥ó", vbDirectory) <> "" Then
    '    cntStr = "Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=H:\§Úªº¶³ºÝµwºÐ\¨p¤H\¤d¼{¤@±oÂN(C¼Ñª©)\®ÑÄy¸ê®Æ\¹Ï®ÑºÞ²zªþ¥ó\¬d¦r.mdb;"
    'ElseIf Dir("D:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\¹Ï®ÑºÞ²zªþ¥ó", vbDirectory) <> "" Then
    '    cntStr = "Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=D:\¤d¼{¤@±oÂN\®ÑÄy¸ê®Æ\¹Ï®ÑºÞ²zªþ¥ó\¬d¦r.mdb;"
    'Else
    '    MsgBox "¸ô®|¤£¦s¦b¡I", vbCritical: Exit Sub
    'End If
    Set d = ActiveDocument: dx = d.Range.text: Set rngF = d.Range
    'cnt.Open cntStr
    word.Application.ScreenUpdating = False
    
    GoSub bookmarks '¼ÐÂI²Å¸¹_®Ñ¦W¸¹_¦Û°Ê¥[¤W¥Î
    rst.Open "select * from ¼ÐÂI²Å¸¹_½g¦W¸¹_¦Û°Ê¥[¤W¥Î order by ±Æ§Ç", cnt, adOpenForwardOnly, adLockReadOnly
    Set rngF = d.Range: dx = d.Range.text
    Do Until rst.EOF
        title = rst("½g¦W").Value
        If VBA.InStr(dx, title) Then 'if found
            Do While rngF.Find.Execute(title, , , , , , True, wdFindStop)
    '            If InStr("¡n¡r¡P¡E", IIf(rngF.Characters(rngF.Characters.Count).Next Is Nothing, "", rngF.Characters(rngF.Characters.Count).Next)) = 0 And _
    '                InStr("¡m¡q¡P¡E", IIf(rngF.Characters(1).Previous Is Nothing, "", rngF.Characters(1).Previous)) = 0 Then
                    If ®Ñ¦W¸¹½g¦W¸¹¼Ðª`PreExamOK(d, title, rngF.start) Then
                        If VBA.IsNull(rst("¨ú¥N¬°").Value) Then
                        Rem «ç»ò¼g³£·|¼vÅT­ì¨Ó¤å¦r®æ¦¡¡A¬G¥u¯à¥Î«e«á´¡¤Jªº¤è¦¡¤F
                            rngF.InsertBefore "¡q"
                            rngF.InsertAfter "¡r"
                            'rngF.text = "¡q" & title & "¡r"
                                      'd.Range.Find.Execute title, , , , , , True, wdFindContinue, , "¡q" & title & "¡r", wdReplaceAll
                        Else
                            rngF.text = rst("¨ú¥N¬°").Value
                            'd.Range.Find.Execute title, , , , , , True, wdFindContinue, , rst("¨ú¥N¬°").Value, wdReplaceAll
                        End If
                        rngF.SetRange rngF.End, d.Range.End
                    End If
    '            End If
            Loop
            Set rngF = d.Range: dx = d.Range.text
        End If
        
        rst.MoveNext
    Loop
    d.Range.Find.Execute "¡m¡m", , , , , , True, wdFindContinue, , "¡m", wdReplaceAll
    d.Range.Find.Execute "¡n¡n", , , , , , True, wdFindContinue, , "¡n", wdReplaceAll
    d.Range.Find.Execute "¡q¡q", , , , , , True, wdFindContinue, , "¡q", wdReplaceAll
    d.Range.Find.Execute "¡r¡r", , , , , , True, wdFindContinue, , "¡r", wdReplaceAll
    
    'GoSub bookmarks 'do again to check and correct SHOULD BE use another table to do this
    If ur.CustomRecordLevel > 0 Then
        SystemSetup.playSound 1.921
    Else
        SystemSetup.playSound 1
    End If
    rst.Close: cnt.Close: SystemSetup.contiUndo ur
    word.Application.ScreenUpdating = True
    
    
    Exit Sub
    
    
bookmarks:
    If rst.State = adStateOpen Then rst.Close
    rst.Open "select * from ¼ÐÂI²Å¸¹_®Ñ¦W¸¹_¦Û°Ê¥[¤W¥Î order by ±Æ§Ç", cnt, adOpenForwardOnly, adLockReadOnly
    Do Until rst.EOF
        title = rst("®Ñ¦W").Value
        
'        If title = "©P©ö" Then Stop 'just for test
        
        If VBA.InStr(dx, title) Then 'if found
            Do While rngF.Find.Execute(title, , , , , , True, wdFindStop)
    '            If InStr("¡n¡r¡P¡E", IIf(rngF.Characters(rngF.Characters.Count).Next Is Nothing, "", rngF.Characters(rngF.Characters.Count).Next)) = 0 And _
    '                InStr("¡m¡q¡P¡E", IIf(rngF.Characters(1).Previous Is Nothing, "", rngF.Characters(1).Previous)) = 0 Then
                    If ®Ñ¦W¸¹½g¦W¸¹¼Ðª`PreExamOK(d, title, rngF.start) Then
                        
''                        If title = "¸êªv³qÅ²" Then Stop 'just for test
'                        'just for test
'                        If title = "©P©ö" And rngF.Previous.text = "¸q" Then 'just for test
'                            rngF.Select
'                            Stop
'                        End If
                        
'                        Dim rngfFormattedText As Range
'                        Set rngfFormattedText = rngF.FormattedText.Duplicate '°O¤U­ì¨Ó®æ¦¡
                        If VBA.IsNull(rst("¨ú¥N¬°").Value) Then
                            Rem «ç»ò¼g³£·|¼vÅT­ì¨Ó¤å¦r®æ¦¡¡A¬G¥u¯à¥Î«e«á´¡¤Jªº¤è¦¡¤F
                            rngF.InsertBefore "¡m"
                            rngF.InsertAfter "¡n"
                            'rngF.text = "¡m" & title & "¡n"
                            
                            'rngF.Find.Execute rngF.text, , , , , , , , , "¡m" & title & "¡n", wdReplaceOne
                            
'                            rngF.FormattedText.text = "¡m" & title & "¡n"

'                            rngF.FormattedText = rngfFormattedText '«ì´_­ì¨Ó®æ¦¡

                '            d.Range.Find.Execute title, , , , , , True, wdFindContinue, , "¡m" & title & "¡n", wdReplaceAll
                        Else
                            Rem ¶Õ¥²¼vÅT­ì¨Óªº¤å¦r®æ¦¡¤F
                            rngF.text = rst("¨ú¥N¬°").Value
                '            d.Range.Find.Execute title, , , , , , True, wdFindContinue, , rst("¨ú¥N¬°").Value, wdReplaceAll
                        End If
                        rngF.SetRange rngF.End, d.Range.End
                    End If
    '            End If
            Loop
            Set rngF = d.Range: dx = d.Range.text
        End If
        
        rst.MoveNext
    Loop
    rst.Close
    Return
    
eH:
        Select Case Err.Number
            Case Else
                MsgBox Err.Number + Err.Description
    '            Resume
        End Select
End Sub


Sub ¤À¦æ¤À¬q_®Ú¾Ú²Ä1¦æªº¦r¼Æªø«×¨Ó§@¤Á³Î()
Dim wordCount As Byte, d As Document, rng As Range, i As Integer, dx As String, a, p As Paragraph, j As Byte, wl
Dim omitStr As String
omitStr = "{}<p>¡m¡n¡q¡r¡G¡A¡C¡u¡v¡y¡z¡@¡P0123456789-" & VBA.ChrW(8231) & VBA.ChrW(183) & VBA.Chr(13)
If word.Documents.Count = 0 Then
    Set d = Documents.Add()
ElseIf ActiveDocument.path <> "" Then
    Set d = Documents.Add() 'ActiveDocument
Else
    Set d = ActiveDocument
End If
Set rng = d.Range
rng.Paste
Set p = rng.Paragraphs(1)
'wordCount = p.Range.Characters.Count - 1
For Each a In p.Range.Characters
    If InStr(omitStr, a) = 0 Then wordCount = wordCount + 1
Next a
dx = rng.text
wl = InStr(dx, VBA.Chr(13))
rng.text = VBA.Left(dx, wl) & Replace(dx, VBA.Chr(13), "", wl)

i = 1
Do Until rng.Paragraphs(rng.Paragraphs.Count).Range.Characters.Count < wordCount
    i = i + 1
    If i > rng.Paragraphs.Count Then Exit Do
    Set p = rng.Paragraphs(i)
    For Each a In p.Range.Characters
        If InStr(omitStr, a) = 0 Then j = j + 1
        If j = wordCount Then
            a.InsertAfter VBA.Chr(13)
            j = 0
            Exit For
        End If
    Next a
'    rng.Paragraphs(i).Range.Characters(wordCount).InsertAfter vba.Chr(13)
Loop
rng.Cut
rng.Document.Close wdDoNotSaveChanges
If word.Documents.Count = 0 Then
    word.Application.Quit
Else
    word.ActiveWindow.windowState = wdWindowStateMinimize
End If
Beep
End Sub
Sub replaceWithNextChararcter() 'Alt+Shift+h
Dim s As Integer, chars 'As Characters
Dim f As String, r As String
Set chars = Selection.Characters
If chars.Count < 2 And InStr(Selection, VBA.Chr(9)) = 0 Then Exit Sub
If chars.Count > 2 Then
    s = InStr(Selection, VBA.Chr(9))
    If s > 0 Then
        If InStr(VBA.Mid(Selection.text, s + 1), VBA.Chr(9)) = 0 Then
            chars = VBA.Split(Selection.text, VBA.Chr(9))
            Selection.text = VBA.Left(Selection.text, s - 1)
            s = 0
            f = chars(s): r = chars(s + 1) 'VBA.IIf(chars(s + 1) = vba.Chr(9), "", chars(s + 1))
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If
Else
    s = 1
    f = chars(s)
    r = VBA.IIf(chars(s + 1) = VBA.Chr(9), "", chars(s + 1))
    Selection.Characters(s + 1) = ""
End If
Selection.Find.Execute f, , , , , , True, wdFindContinue, , r, wdReplaceAll
End Sub

Sub °ê»yÃã¨åºô§}¤ÎID©|¯ÊªÌ¦C¥X()
Dim db As New dBase
db.°ê»yÃã¨åºô§}¤ÎID©|¯ÊªÌ¦C¥X
SystemSetup.playSound 12
End Sub
Sub °ê»yÃã¨åºô§}¤ÎID©|¯ÊªÌ¶ñ¤J()
Dim i As Long
ActiveDocument.Range.Find.Execute VBA.Chr(13), , , , , , , wdFindContinue, , "", wdReplaceAll
Do Until Selection.End = ActiveDocument.Range.End - 1
    Selection.Move
    If Selection.Previous <> VBA.ChrW(20008) And Selection.Hyperlinks.Count = 0 Then
        ¥ÍÃø¦r¥[¤W°ê»yÃã¨åª`­µ
        ActiveWindow.ScrollIntoView Selection, False
        i = i + 1
    End If
    If i = 40 Then Exit Sub
Loop
Selection.HomeKey wdStory, wdExtend
End Sub

Rem 20230707 Bing¤jµÐÂÄ¡G §PÂ_¥þ§Î¥b§Î¦r
Public Function FullOrHalf(ByVal str As String) As Integer
    Dim strLocal As String
    Debug.Assert Len(str) = 1
    If Len(str) <> 1 Then
        FullOrHalf = -1
        Exit Function
    End If
    strLocal = VBA.StrConv(str, vbFromUnicode)
    If Len(str) * 2 = LenB(strLocal) Then
        FullOrHalf = 2 ' wide
    ElseIf Len(str) = LenB(strLocal) Then
        FullOrHalf = 1 ' narrow
    Else
        FullOrHalf = 0 ' error
    End If
End Function
Rem Bing¤jµÐÂÄ¡G
'³o­Ó¨ç¼Æ±µ¨ü¤@­Ó¦r²Å¦ê§@¬°¿é¤J¡Aªð¦^¤@­Ó¾ã¼Æ­È¡C¦pªGªð¦^­È¬° 2¡A«hªí¥Ü¿é¤Jªº¦r²Å¬O¥þ¨¤¡F¦pªGªð¦^­È¬° 1¡A«hªí¥Ü¿é¤Jªº¦r²Å¬O¥b¨¤¡F¦pªGªð¦^­È¬° -1 ©Î 0¡A«hªí¥Ü¥X²{¿ù»~¡C
'
'§PÂ_ªº­ì²z¬O±N½s½X±q Unicode Âà¬°¥»¦a½s½X¡AµM«á¤ñ¸ûÂà´««e«á¦r²Å¦êªºªø«×¡C¦pªGÂà´««e«á¦r²Å¦êªø«×¬Ûµ¥¡A«hªí¥Ü¿é¤Jªº¦r²Å¬O¥b¨¤¡F¦pªGÂà´««á¦r²Å¦êªø«×¬OÂà´««e¦r²Å¦êªø«×ªº¨â­¿¡A«hªí¥Ü¿é¤Jªº¦r²Å¬O¥þ¨¤(1)¡C
'
'¨Ó·½: »P Bing ªº¥æ½Í¡A 2023/7/7
'(1) ¤@¤å¹ý©³·d©wvba³B²z¥þ¨¤¥b¨¤ - ª¾¥G. https://zhuanlan.zhihu.com/p/600306305.
'(2) WordVBA¡G¥b¨¤¦r²ÅÂà¬°¥þ¨¤¦r²Å¡]µ²¦X¬d§ä¤èªk¡^_word¥b¨¤²Å¸¹§ï¬°¥þ¨¤²Å¸¹§»_VBA-¦u­Ôªº³Õ«È-CSDN³Õ«È. https://blog.csdn.net/qq_64613735/article/details/124760907.
'(3) office³n¥óword¤åÀÉ¤¤¦p¦ó¿ë§O¥b¨¤©M¥þ¨¤ - ¦Ê«×ª¾¹D. https://zhidao.baidu.com/question/347564125.html.
'(4) VBA¨ç¼Æ§å¶q±N±N¦r²Å¥Ñ¥þ¨¤Âà¬°¥b¨¤¡A©Î¥Ñ¥b¨¤Âà¬°¥þ¨¤-¦P®É¾A¥ÎExcel Access - Excel¨ç¼Æ¤½¦¡ - Office¥æ¬yºô. https://www.office-cn.net/excel-func/297.html.
'(5) ¦p¦ó¤À¿ëword¤å³¹¤¤ªº¼ÐÂI¬O¥þ¨¤ÁÙ¬O¥b¨¤¡H_¦Ê«×ª¾¹D. https://zhidao.baidu.com/question/45987987.html.


Rem ¦r¦êÂà¦r¦ê°}¦C creedit with chatGPT¤jµÐÂÄ
Function SplitWithoutDelimiter_StringToStringArray(str As String) As String()
Dim lenStr  As Long, arr() As String, i As Long, ch As String, eCount As Long
lenStr = VBA.Len(str)

'str =­nÂà´«¬°°}¦Cªº¦r¦ê

' ±N¦r¦êÂà´«¬°°}¦C
For i = 1 To lenStr
    ch = VBA.Mid(str, i, 1)
    eCount = eCount + 1
    If code.IsHighSurrogate(ch) Then
        ch = VBA.Mid(str, i, 2): i = i + 1
    End If
    ReDim Preserve arr(eCount - 1) ' ½Õ¾ã°}¦C¤j¤p¡A¨Ï¨ä»P¦r¦êªø«×¬Û¦P
    arr(eCount - 1) = ch
Next i
SplitWithoutDelimiter_StringToStringArray = arr
End Function

