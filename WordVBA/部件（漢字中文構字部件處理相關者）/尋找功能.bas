Attribute VB_Name = "尋找功能"
Option Explicit
Public ThisDocument常用5032漢字部件表Dictionary As New Scripting.Dictionary

Sub 尋找前圖後字()
Dim isp As InlineShape
Static x As String, n As String, nw As Long
x = InputBox("請輸入圖片命名敘述", , x)
If x = "" Then Exit Sub
n = InputBox("請輸入圖片後欲找之文字,或其Ascw值", , nw)
If n = "" Then Exit Sub
If IsNumeric(n) Then
    nw = n
    n = ChrW(CLng(n))
Else
    nw = n
End If
For Each isp In ActiveDocument.InlineShapes
    If isp.AlternativeText = x Then
        If isp.Range.Characters(1).Next = n Then
            ActiveDocument.Range(isp.Range.Start, isp.Range.Characters(1).Next.End).Select
            Exit Sub
        End If
    End If
Next
MsgBox "沒找到!", vbExclamation
End Sub

Sub 檢索用欄位更新()
Dim d As Document, t As Table, c As Cell, a, i As Integer, w As String, cln As Byte, clnSearch As Long, flg As Boolean, ck As Boolean, cInPut As Cell
Set d = ThisDocument
Set t = d.Tables(1)
For Each c In t.Rows(1).Cells
    clnSearch = clnSearch + 1
    If InStr(c.Range, "檢索用欄位") > 0 Then
        flg = True
        Exit For
    End If
Next c
If flg = False Then
    MsgBox "找不到""檢索用欄位""欄位", vbCritical
    Exit Sub
Else
    flg = False
End If
For Each c In t.Rows(1).Cells
    cln = cln + 1
'    If InStr(c.Range, "部件（合併") > 0 Then
    If InStr(c.Range, "部件（原形") > 0 Then
        flg = True
        Exit For
    End If
Next c
If flg = False Then
'    MsgBox "找不到""部件（合併""""欄位", vbCritical
    MsgBox "找不到""部件（原形""""欄位", vbCritical
    Exit Sub
End If
For Each c In t.Columns(cln).Cells
    i = i + 1
    If i > 1 Then
        For Each a In c.Range.Characters
            If a <> Chr(13) & Chr(7) Then
                If a.InlineShapes.Count > 0 Then
                    w = w & a.InlineShapes(1).AlternativeText
                Else
                    w = w & a
                End If
            End If
        Next a
'        t.Cell(c.Row.Index, clnSearch).Range.Text = w'太慢，故改用.Next
        Set cInPut = c.Next.Next.Next
        If Not ck Then
            If t.Cell(1, cInPut.ColumnIndex).Range.Text <> "檢索用欄位" & Chr(13) & Chr(7) Then
                MsgBox "程式有誤，輸入之欄位非「檢索用欄位」", vbCritical
                cInPut.Select
                Exit Sub
            Else
                ck = True
            End If
        End If
        cInPut.Range.Text = w
        w = ""
    End If
Next c
MsgBox "done!", vbInformation
End Sub


Rem 20230316 creedit with Bing大菩薩
'查看只因部首變形而構成不同漢字中文所有相關的漢字（即除了部首變形外，其他的部件皆同。尚不計其排列方式）
Rem 衣部變形部首分上下者，先不錯，只有此例，其實可以人工找出
Sub 找出所有只有一個變形部首部件不同的字Dictionary()
Dim t As Table, d As Document, a As Range, c As Cell, result As String
Dim dictComponents As New Scripting.Dictionary, components As New components, componentsArray() As String, w As String, wList As String, key As String
Dim dictVariantRadicals As Scripting.Dictionary, variantRadical As String, w_ofVariantRedical As String, w_ofVariantRedical_componentsCollection As VBA.Collection
Dim sameComposeDict As New Scripting.Dictionary '唯除變形部首外的同部件組成仍然太多，今改用 Dictionary sameComposeDict 儲存 key=部首 value= 字 以便找出同部首同而有兩個以上的漢字才輸入
Const columnComponents As Byte = 2 '部件欄位
Const columnChar As Byte = 1 '漢字欄位

'取得漢字及其部件資料
Set d = ThisDocument
Set t = d.Tables(1)
'取得變形部首資料
Set dictVariantRadicals = components.VariantRadicalsDictionary
For Each c In t.Columns(columnComponents).Cells
    If c.RowIndex > 1 Then '第1列是標題時
        '取得漢字
        Set a = t.Cell(c.RowIndex, columnChar).Range
        w = VBA.Mid(a.Text, 1, Len(a.Text) - 2) '儲存格標記乃尾綴為 chr(13) & chr(7),要扣掉
        '取得部件陣列 componentsArray （引數以傳址（pass by reference）方式傳遞）
        components.getComponentsArray componentsArray, c.Range
        '排列部件陣列（即忽略部件的排列方式（在其所組成的漢字中的位置））
        Call components.SortStringArray(componentsArray)
        If UBound(componentsArray) > 0 Then
            Dim arr, earr, iarr As Long, subsetExcludingVariantRadicalsDict As Scripting.Dictionary, e_earr
            
'            If w = "摹" Then sndPlaySound32 "C:\Windows\Media\Ring10.wav", 1: Stop 'for debug
'            If w = "愀" Then sndPlaySound32 "C:\Windows\Media\Ring10.wav", 1: Stop 'for debug
'            If w = "來" Then Beep: Stop 'for debug
'            If w = "巫" Then Beep: Stop 'for debug
'            If w = "江" Then sndPlaySound32 "C:\Windows\Media\Ring10.wav", 1: Stop 'for debug
'            If w = "汞" Then Beep: Stop 'for debug
            
            '取得忽略部首的部件集合集
            Set subsetExcludingVariantRadicalsDict = components.subsetExcludingVariantRadicals(componentsArray)
            If Not subsetExcludingVariantRadicalsDict Is Nothing Then
                Dim cln As VBA.Collection
                Dim clnKeysArr
                Dim eClnKeysArr
                Dim radical_variantRadical As String, radical_in_dictComponents As String
                '取得忽略變形部首的部件組成集合，是一個二維陣列包在Collection裡面
                clnKeysArr = subsetExcludingVariantRadicalsDict.Keys
                For Each eClnKeysArr In clnKeysArr
                    Set cln = eClnKeysArr
                    radical_variantRadical = dictVariantRadicals(subsetExcludingVariantRadicalsDict(cln))
                    Rem Bing大菩薩：
'                    根據網路搜尋的結果，VBA 的 Dictionary 物件類別的鍵值（key）可以是任何資料類型，包括 Collection12。但是，如果要使用 Collection 作為鍵值，需要注意以下幾點3：
                    'Collection 作為鍵值時，必須是一維的。
                    'Collection 作為鍵值時，必須有相同的元素個數和順序才能被視為相同的鍵值。
                    'Collection 作為鍵值時，不能直接用索引或 Keys 方法來存取，需要先轉換成 Variant 或 String 類型。……
                    '來源: 與 Bing 的交談， 2023/3/17(1) Keys method (Visual Basic for Applications) | Microsoft Learn. https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/keys-method 已存取 2023/3/17.
                    '(2) Dictionary 物件 | Microsoft Learn. https://learn.microsoft.com/zh-tw/office/vba/language/reference/user-interface-help/dictionary-object 已存取 2023/3/17.
                    '(3) excel - VBA get key by value in Collection - Stack Overflow. https://stackoverflow.com/questions/12561539/vba-get-key-by-value-in-collection 已存取 2023/3/17.
                    arr = cln.Item(1)
                    '如果是陣列
                    If TypeName(arr) = "Variant()" Then
                        For iarr = LBound(arr) To UBound(arr)
                            '取得二維陣列內的一維陣列元素存入變數 earr
                            earr = Excel.Application.Index(arr, iarr, 0)
                            '如果取得的是陣列（一維陣列）
                            If TypeName(earr) = "Variant()" Then
                                '陣列轉成文字字串（以串接 Join 的方法），作為比對用的 key 索引鍵
                                key = VBA.Trim(VBA.Join(earr))
                                '如果key不是半形空格 rem 因為subsetExcludingVariantRadicals()內定義的是二維陣列，又會跳過非變形部首部件不處理，所以回傳的會包含許多空的元素（陣列大小是固定不易變或不可變的）
                                If key <> " " And key <> "" Then
                                    '如果找到相同的略過變形部首的部件組成，
                                    If dictComponents.Exists(key) Then
                                        '先核對二者部件數是否一致
                                        '先取得已記錄部件情狀之漢字，好來與目前的漢字作部件數的比對
                                        wList = components.JoinDictionaryValuesWithComma(dictComponents(key))  '取得忽略變形部首後有相同部件組成的漢字是什麼；
                                        '如果有多個，以第1個為代表來比較即可。因為都是忽略了一樣的變形部首
                                        Dim commaPos As Integer
                                        commaPos = VBA.InStr(wList, ",")
                                        w_ofVariantRedical = VBA.IIf(commaPos > 0, VBA.Mid(wList, 1, IIf(commaPos = 0, 0, commaPos - 1)), wList)
                                        
                                        '比對二者部件數是否一致，如果一致，
                                        If components.componentsCountofChar(w_ofVariantRedical) = components.componentsCountofChar(w) Then
                                        
                                            '接著就要比對該所忽略不計的變形部首是否是同屬一個部首，是同一部首的變形部首比對才有意義；即此方法之主旨
                                            variantRadical = subsetExcludingVariantRadicalsDict(cln) '取得現在要比對的漢字w所忽略的變形部首是什麼
                                            '取得要比對的含變形部首的漢字其部件元素的集合。這裡不能用字典，因為會消掉重複的部件，如「來、巫」都有兩個「人」。用陣列又不好清除元素，故用Collection集合來作
                                            Set w_ofVariantRedical_componentsCollection = components.DecomposeCollection(w_ofVariantRedical)
                                            Dim i_w_ofVariantRedical_componentsCollection As Byte
                                            '以現有的部件與要比對的部件作比對
                                            For Each e_earr In earr
                                                For i_w_ofVariantRedical_componentsCollection = 1 To w_ofVariantRedical_componentsCollection.Count
                                                    If VBA.StrComp(e_earr, w_ofVariantRedical_componentsCollection.Item(i_w_ofVariantRedical_componentsCollection)) = 0 Then
                                                        w_ofVariantRedical_componentsCollection.Remove i_w_ofVariantRedical_componentsCollection '逐一移除相同的，剩下的、唯一的就是那個變形部首了
                                                        Exit For
                                                    End If
                                                Next i_w_ofVariantRedical_componentsCollection
                                            Next
                                            '到此 w_ofVariantRedical_componentsCollection 應當只剩一個元素了
                                            
                                            Rem 以取得的變形部首與現在要處理的漢字的變形部首作比對，是同屬一部首才直接處理
                                            
                                            '取得已記錄的變形部首radical_in_dictComponents、及現在要來比對的變形部首radical_variantRadical分別是什麼部首
                                            Rem 但是除變形部首外的同部件組成仍然太多，今改用 Dictionary sameComposeDict 儲存 key=部首 value= 字 以便找出同部首同而有兩個以上的漢字才輸入
                                            'Dim radical_variantRadical As String, radical_in_dictComponents As String
                                            radical_in_dictComponents = dictVariantRadicals(w_ofVariantRedical_componentsCollection.Item(1))
                                            radical_variantRadical = dictVariantRadicals(variantRadical)
                                            Rem 部首同不同都得處理，因為除變形部首外同部件組成的仍然太多，所以改用Dictionary儲存其漢字value（值）與所歸之部首key（鍵）之鍵值對（鍵-值對），以便查找比對與刪除
                                            '如果部首一致
                                            If VBA.StrComp(radical_variantRadical, radical_in_dictComponents) = 0 Then
                                                                              
                                                wList = dictComponents(key)(radical_variantRadical) '取得忽略變形部首後有相同部件組成的漢字字串清單，用來比對是否已有該字存在；若有則不再重複添入
                                                If VBA.InStr(wList, w) = 0 Then '尚無才加入，因為有不同的部件組合要逐一與全部已儲存的部件組比對檢查
'                                                    sndPlaySound32 "C:\Windows\Media\Alarm10.wav", 1'因為太多，便不撥放提示音效了
                                                    If wList <> "" Then wList = wList & "," '以「,」區別各個漢字，可用來判斷是否不止一個漢字。下同。
                                                    dictComponents(key)(radical_variantRadical) = wList & w
                                                End If
                                            
                                            '如果部首不同，則作歸類，以取得每個部首的所領的漢字群組數都在1個以上才送交resultDict輸出
                                            Else
                                                
                                                Rem 變形部首不同部首而部件組成相同者太多，故須再加篩選
                                                wList = dictComponents(key)(radical_variantRadical) '取得忽略變形部首後有相同部件組成的漢字字串清單，用來比對是否已有該字存在；若有則不再重複添入
                                                If VBA.InStr(wList, w) = 0 Then
                                                    If wList <> "" Then wList = wList & ","
                                                    dictComponents(key)(radical_variantRadical) = wList & w
                                                End If
'                                                dictComponents(key) = wList & "," & w'此為原式，但以字串w而非字典型別的sameComposeDict來儲存，固不敷用
                                            End If '以取得的變形部首與現在要處理的漢字的變形部首作比對，是同屬一部首才處理（才有意義）
'                                        Else
'                                            dictComponents(key) = wList & "," & w
                                        End If '以上比對二個漢字間的部件數是否一致，如果一致……
                                        
                                    '如果沒有找到相同的略過變形部首的部件組成，則新增記錄（元素）到dictComponents中（其構造：鍵：為部件組成之字串，值：為一個其值為符合此鍵值部件之漢字、鍵為該字之部首之字典 sameComposeDict儲存）
                                    Else
                                        'dictComponents.Add key, w'原僅用字串儲存所有符合條件的漢字結果，但不敷應用，故今已改為Dictionary sameComposeDict儲存
'                                        dictComponents(key) = w
                                        sameComposeDict.Add radical_variantRadical, w '建立新的部首-漢字群鍵-值對字典以作為dictComponents的value值
                                        dictComponents.Add key, sameComposeDict '加入新的部件組成部首漢字群到回來準備回傳的dictComponents字典中
                                        Set sameComposeDict = Nothing '應當同Collection原理（參見components物件類別模組中的subsetExcludingVariantRadicals方法），這樣才能清空備下次使用（資源再利用）。但
                                    End If
                                End If
                            End If
                        Next iarr
                    End If '以上 end If: TypeName(arr) = "Variant()" Then
                Next eClnKeysArr
            End If '以上 end If: Not subsetExcludingVariantRadicalsDict is Nothing Then
        End If
    End If
Next c

Rem 不再用到的變數(如 e_earr、sameComposeDict、eClnKeysArr。只要型別合格即可)在此作資源再利用，故不再宣告新變數來操作
'篩選同部首有兩字以上的元素
For Each e_earr In dictComponents.Keys
    If e_earr <> "" Then
        Set sameComposeDict = dictComponents(e_earr)
        For Each eClnKeysArr In sameComposeDict.Keys '巡查每個部首
            If InStr(sameComposeDict(eClnKeysArr), ",") = 0 Then '如果沒有兩個漢字以上。以「,」區別各個漢字，故可用來判斷
                sameComposeDict.Remove eClnKeysArr '就移除該部首
            End If
        Next
    End If
Next e_earr
Dim resultDict As New Scripting.Dictionary '作為最後用來輸出查對結果的字典（即作為JoinDictionaryValues()的引數）
'只剩帶兩個部首以上的,就取回最後的結果字典
'Set sameComposeDict = Nothing；因為下面已有 Set sameComposeDict = dictComponents(e_earr) 陳述式，故此處便不用先清除再用。 = 運算子自會除舊以布新
For Each e_earr In dictComponents.Keys
    If e_earr <> "" Then
        Set sameComposeDict = dictComponents(e_earr) '取得結果的字典
        For Each eClnKeysArr In sameComposeDict.Keys '巡查每個含兩字以上的部首
            resultDict(e_earr) = sameComposeDict(eClnKeysArr) '取得最後輸出的結果
        Next
    End If
Next e_earr


result = components.JoinDictionaryValues(resultDict) '原是傳入 dictComponents 作引數，今改是。
Documents.Add().Range.Text = result '只有一個變形部首部件不同的字
sndPlaySound32 "C:\Windows\Media\Ring05.wav", 1
Rem 釋放記憶體
Set subsetExcludingVariantRadicalsDict = Nothing: Set components = Nothing: Set dictComponents = Nothing: Set dictVariantRadicals = Nothing: Set w_ofVariantRedical_componentsCollection = Nothing: Set sameComposeDict = Nothing: Set resultDict = Nothing
Rem 希望在作業系統激活文件及Word視窗，使其成為最前端的（不怎麼有用）
Application.ActiveDocument.ActiveWindow.Activate
Application.Activate
End Sub


Sub 找出所有只有一個部件不同的字Dictionary()
Dim t As Table, d As Document, a As Range, c As Cell, result As String
Dim dictComponents As New Scripting.Dictionary, components As New components, componentsArray() As String, w As String, key As String, arrCom()
Const columnComponents As Byte = 2 '部件欄位
Const columnChar As Byte = 1 '漢字欄位

'取得漢字及其部件資料
Set d = ThisDocument
Set t = d.Tables(1)
For Each c In t.Columns(columnComponents).Cells
    If c.RowIndex > 1 Then '第1列是標題時
        '取得漢字
        Set a = t.Cell(c.RowIndex, columnChar).Range
        w = VBA.Mid(a.Text, 1, Len(a.Text) - 2) '儲存格標記乃尾綴為 chr(13) & chr(7),要扣掉
        '取得部件陣列 componentsArray
        components.getComponentsArray componentsArray, c.Range
        '排列部件陣列
                Call components.SortStringArray(componentsArray)
                If UBound(componentsArray) > 0 Then
'                    Dim clnComponets As VBA.Collection, eClnComponets
'                    Set clnComponets = components.Subset(componentsArray)
                    Dim arr, earr, iarr As Long
                    arr = components.Subset(componentsArray)
'                    For Each eClnComponets In clnComponets
                    For iarr = LBound(arr) To UBound(arr)
                        earr = Excel.Application.Index(arr, iarr, 0)

'                        earr = arr(iarr)
'                        If Not earr = Empty Then
'                            If TypeName(earr) = "String" Then
'                                key = earr
'                            Else
                                key = VBA.Join(earr)
'                            End If
    '                        key = eClnComponets
                            If dictComponents.Exists(key) Then
                                sndPlaySound32 "C:\Windows\Media\Alarm10.wav", 1
                                If InStr(dictComponents(key), w) = 0 Then
                                    dictComponents(key) = dictComponents(key) & "," & w
                                End If
                            Else
                                'dictComponents.Add key, w
                                dictComponents(key) = w
                            End If
'                        End If
'                    Next eClnComponets
                    Next iarr
                End If
    End If
Next c

result = components.JoinDictionaryValues(dictComponents)
Documents.Add().Range.Text = result '只有一個部件不同的字
sndPlaySound32 "C:\Windows\Media\Ring05.wav", 1
Application.Activate
End Sub

Sub 部件構成字表_5032字由兩個部件以上構成者()
Dim t As Table, d As Document, a As Range, c As Cell, result As String
Dim dictComponents As New Scripting.Dictionary, components As New components, componentsArray() As String, w As String, key As String, arrCom()
Const columnComponents As Byte = 2 '部件欄位
Const columnChar As Byte = 1 '漢字欄位

'取得漢字及其部件資料
Set d = ThisDocument
Set t = d.Tables(1)
For Each c In t.Columns(columnComponents).Cells
    If c.RowIndex > 1 Then '第1列是標題時
        '取得漢字
        Set a = t.Cell(c.RowIndex, columnChar).Range
        w = VBA.Mid(a.Text, 1, Len(a.Text) - 2) '儲存格標記乃尾綴為 chr(13) & chr(7),要扣掉
        '取得部件陣列 componentsArray
        components.getComponentsArray componentsArray, c.Range
        '排列部件陣列
                Call components.SortStringArray(componentsArray)
                If UBound(componentsArray) > 0 Then
'                    Dim clnComponets As VBA.Collection, eClnComponets
'                    Set clnComponets = components.Subset(componentsArray)
                    Dim arr, earr
                    arr = components.Subset(componentsArray)
'                    For Each eClnComponets In clnComponets
                    For Each earr In arr '這樣會列出所有的元素值而不是二維陣列的一維陣列元素
                        If Not earr = Empty Then
                            If TypeName(earr) = "String" Then
                                key = earr
                            Else
                                key = VBA.Join(earr)
                            End If
    '                        key = eClnComponets
                            If dictComponents.Exists(key) Then
                                sndPlaySound32 "C:\Windows\Media\Alarm10.wav", 1
                                If InStr(dictComponents(key), w) = 0 Then
                                    dictComponents(key) = dictComponents(key) & "," & w
                                End If
                            Else
                                'dictComponents.Add key, w
                                dictComponents(key) = w
                            End If
                        End If
'                    Next eClnComponets
                    Next earr
                End If
    End If
Next c

result = components.JoinDictionaryValues(dictComponents)
Documents.Add().Range.Text = result
sndPlaySound32 "C:\Windows\Media\Ring05.wav", 1
Application.Activate
End Sub


Sub 找出所有具有相同部件組成的字Dictionary()
Dim t As Table, d As Document, a As Range, c As Cell, result As String
Dim dictComponents As New Scripting.Dictionary, components As New components, componentsArray() As String, w As String, key As String
Const columnComponents As Byte = 2 '部件欄位
Const columnChar As Byte = 1 '漢字欄位

'取得漢字及其部件資料
Set d = ThisDocument
Set t = d.Tables(1)
For Each c In t.Columns(columnComponents).Cells
    If c.RowIndex > 1 Then '第1列是標題時
        '取得漢字
        Set a = t.Cell(c.RowIndex, columnChar).Range
        w = VBA.Mid(a.Text, 1, Len(a.Text) - 2) '儲存格標記乃尾綴為 chr(13) & chr(7),要扣掉
        '取得部件陣列 componentsArray
        components.getComponentsArray componentsArray, c.Range
        
        Call components.SortStringArray(componentsArray)
        key = VBA.Join(componentsArray)
        If dictComponents.Exists(key) Then

            sndPlaySound32 "C:\Windows\Media\Alarm10.wav", 1
            dictComponents(key) = dictComponents(key) & "," & w
        Else
            dictComponents.Add key, w
        End If
    End If
Next c

result = components.JoinDictionaryValues(dictComponents)
Documents.Add().Range.Text = result
sndPlaySound32 "C:\Windows\Media\Ring05.wav", 1

End Sub
Sub 找出所有具有相同部件組成的字Collection()
Dim t As Table, d As Document, charCl As New Collection, componentsCl As New Collection, c As Cell, inlsp As InlineShape, i As Long, result As String, cl, e, ee, cll, j As Long, cnt As Byte, flag As Boolean
Dim components As New components
'取得漢字及其部件資料
Set d = ThisDocument
Set t = d.Tables(1)
For Each c In t.Range.Cells
    If c.RowIndex > 1 Then '第1列是標題時
        Select Case c.ColumnIndex
            Case 1 '漢字
                    charCl.Add VBA.Left(c.Range, 1)
            Case 2 '部件
                    components.getComponentsCollection componentsCl, c.Range
                
'                    '配置部件陣列以備加入colleciton 容器
'                    ReDim componets(c.Range.Characters.Count - 2)
'                    '如果沒有圖片
'                    If c.Range.InlineShapes.Count = 0 Then
'                        For Each a In c.Range.Characters
'                            ' '排除儲存格字元
'                            If InStr(Chr(13) & Chr(7), a) = 0 Then
'                                componets(i) = a.Text
'                                i = i + 1
'                            End If
'                        Next
'                    '如果有圖片
'                    Else
'                        For Each a In c.Range.Characters
'                            '排除儲存格字元
'                            If InStr(Chr(13) & Chr(7), a) = 0 Then
'                                If a.InlineShapes.Count = 0 Then '非圖片
'                                    componets(i) = a.Text
'                                Else '圖片
'                                    componets(i) = a.InlineShapes(1).AlternativeText
'                                End If
'                                i = i + 1
'                            End If
'                        Next a
'                    End If
'                    componentsCl.Add componets
        End Select
        i = 0
    End If
Next c

reset:
i = 0: j = 0
'比對部件相同者
For Each cl In componentsCl
    i = i + 1
    If i > componentsCl.Count Then
        Exit For
    End If
    '取得部件數
    cnt = UBound(cl) + 1
        For Each cll In componentsCl
            j = j + 1
            If i <> j Then '不能和自己比
                If UBound(cll) + 1 = cnt Then '如果部件數一致才進行比對
'                    For Each e In cl
'                        For Each ee In cll
'                            If VBA.StrComp(ee, e) = 0 Then
'                                flag = True
'                                acl.Remove
'                                Exit For
'                            Else
'                                flag = False
'                            End If
'                        Next ee
'                        If flag = False Then Exit For
'                        flag = False
'                    Next e
                Rem creedit with chatGPT大菩薩
                    flag = components.CompareArrays(cl, cll)
                End If
                If flag Then
                    VBA.Beep
                    sndPlaySound32 "C:\Windows\Media\Alarm10.wav", 1
                    'If VBA.InStr(result, charCl(j)) = 0 Then '還沒找到的字才做--不能這樣做
                        result = result & charCl(i) & VBA.vbTab & VBA.Join(componentsCl(i)) & VBA.vbTab & charCl(j) & VBA.vbTab & VBA.Join(componentsCl(j)) & VBA.vbNewLine
                            charCl.Remove j: componentsCl.Remove j
                            flag = False
                        GoTo reset
'                    End If
                End If
            End If
            
        Next cll
        j = 0: flag = False
    
Next cl
Documents.Add().Range.Text = result
'Documents(1).Activate
'Documents(1).Application.Activate
sndPlaySound32 "C:\Windows\Media\Ring05.wav", 1
End Sub
