VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "components"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim dict5032ChineseComponents As New Scripting.Dictionary '常用5032漢字部件表Dictionary : key = 漢字 value =該漢字部件的原來順序（未排序）的一維陣列；即元素為漢字部件的一維陣列
Dim variantRadicalsDict As New Scripting.Dictionary
Sub getComponentsCollection(ByRef componentsCl As Collection, c As Range)
Dim components() As String, a, i As Byte
                    '配置部件陣列以備加入colleciton 容器
                    getComponentsArray components, c
                    componentsCl.Add components
End Sub

Rem 取得部件陣列（Range中的部件文字轉陣列，部件間沒有分隔符號）
Sub getComponentsArray(ByRef components() As String, c As Range) 'VBA引數預設為傳址（pass by reference）方式傳遞
Dim a, i As Byte
        ReDim components(c.Characters.Count - 2)
        '如果沒有圖片
        If c.InlineShapes.Count = 0 Then
            For Each a In c.Characters
                ' '排除儲存格字元
                If InStr(Chr(13) & Chr(7), a) = 0 Then
                    components(i) = a.Text
                    i = i + 1
                End If
            Next
        '如果有圖片
        Else
            For Each a In c.Characters
                '排除儲存格字元
                If InStr(Chr(13) & Chr(7), a) = 0 Then
                    If a.InlineShapes.Count = 0 Then '非圖片，則直接取其字值
                        components(i) = a.Text
                    Else '圖片，則取其替代文字（AlternativeText屬性）之值 ∵ 圖片不能檢索
                        components(i) = a.InlineShapes(1).AlternativeText
                    End If
                    i = i + 1
                End If
            Next a
        End If
                    
End Sub

Rem creedit with chatGPT大菩薩
Function JoinDictionaryValues(dict As Scripting.Dictionary) As String
    Dim result As String
    Dim key As Variant
    For Each key In dict.Keys
        If InStr(dict(key), ",") > 0 Then '以逗號為元素間的間隔符
            result = result & key & VBA.vbTab & dict(key) & VBA.vbNewLine
        End If
    Next key
    If Len(result) > 0 Then
        JoinDictionaryValues = VBA.Replace(Left(result, Len(result) - 1), ",", "") '清除作為元素間間隔符的逗號
    End If
End Function

Rem 將Dictionary物件的值以逗號串接成字串
Function JoinDictionaryValuesWithComma(dict As Scripting.Dictionary) As String
    Dim result As String
    Dim key As Variant
    For Each key In dict.Keys
        result = result & dict(key) & ","
    Next key
    result = Left(result, Len(result) - 1) '去掉尾綴的逗號
    JoinDictionaryValuesWithComma = result
End Function

Rem creedit with chatGPT大菩薩
Sub SortStringArray(ByRef arr() As String)
'arr = Array("apple", "banana", "cherry")
QuickSort arr, LBound(arr), UBound(arr) 'chatGPT大菩薩：LBound 是一個 VBA 函數，它會回傳陣列的下界（Lower Bound），也就是陣列的第一個元素的索引。在大多數情況下，陣列的下界是 0，但有時也可以定義成其他數字。例如，如果定義了一個索引為 1 到 10 的陣列，那麼 LBound 的值就是 1。
'For Each s In arr
'Debug.Print s
'Next s
End Sub

Private Sub QuickSort(ByRef arr() As String, ByVal l As Long, ByVal r As Long) 'l=left,r=right chatGPT大菩薩：是的，在這個 QuickSort 函數中， l 參數代表左邊的索引位置，而 r 參數代表右邊的索引位置。這些參數是快速排序算法的主要部分，用於指定排序的邊界。在這個函數中， arr 數組是要進行排序的數組， l 和 r 指定了要進行排序的數組的區間。
If l >= r Then Exit Sub
Dim i As Long, j As Long, x As String
i = l: j = r: x = arr((l + r) \ 2)
'Do
'    While arr(i) < x
'        i = i + 1
'    Wend
'    While x < arr(j)
'    j = j - 1
'    Wend
'    If i <= j Then
'    Swap arr(i), arr(j)
'    i = i + 1
'    j = j - 1
'    End If
'Loop Until i > j
Do
    While StrComp(arr(i), x, vbTextCompare) < 0
    i = i + 1
    Wend
    While StrComp(x, arr(j), vbTextCompare) < 0
    j = j - 1
    Wend
    If i <= j Then
        Swap arr(i), arr(j)
        i = i + 1
        j = j - 1
    End If
Loop Until i > j
QuickSort arr, l, j
QuickSort arr, i, r
End Sub

Private Sub Swap(ByRef a As String, ByRef b As String)
Dim temp As String
temp = a
a = b
b = temp
End Sub

Rem creedit with chatGPT大菩薩
Function CompareArrays(Array1 As Variant, Array2 As Variant) As Boolean
    Dim Dict1 As Object, Dict2 As Object, element, key
    Set Dict1 = CreateObject("Scripting.Dictionary")
    Set Dict2 = CreateObject("Scripting.Dictionary")

    For Each element In Array1
        If Not Dict1.Exists(element) Then
            Dict1.Add element, 1
        Else
            Dict1(element) = Dict1(element) + 1
        End If
    Next

    For Each element In Array2
        If Not Dict2.Exists(element) Then
            Dict2.Add element, 1
        Else
            Dict2(element) = Dict2(element) + 1
        End If
    Next

    If Dict1.Count <> Dict2.Count Then
        CompareArrays = False
        Exit Function
    End If

    For Each key In Dict1.Keys
        'If Not Dict2.Exists(Key) Or Dict2(Key) <> Dict1(Key) Then
        If Not Dict2.Exists(key) Or VBA.StrComp(Dict2(key), Dict1(key)) <> 0 Then
            CompareArrays = False
            Exit Function
        End If
    Next

    CompareArrays = True
End Function

Rem 取得整個5032漢字與其部件：鍵值漢字，值為該漢字部件的原來順序（未排序）的一維陣列
Property Get 常用5032漢字部件表Dictionary() As Scripting.Dictionary
If dict5032ChineseComponents.Count > 0 Then
        Set 常用5032漢字部件表Dictionary = dict5032ChineseComponents
        Rem 初始化 dict5032ChineseComponents （約需5分鐘）並存到公用變數 常用5032漢字部件表Dictionary 中，則在本檔開啟期間都有效，不必每次重置
        If 尋找功能.ThisDocument常用5032漢字部件表Dictionary.Count = 0 Then Set 尋找功能.ThisDocument常用5032漢字部件表Dictionary = dict5032ChineseComponents

        Exit Property
End If
Dim t As Table, d As Document, a As Range, c As Cell, result As String
Dim components As New components, componentsArray() As String, w As String, key As String
Const columnComponents As Byte = 2 '部件欄位
Const columnChar As Byte = 1 '漢字欄位

'取得漢字及其部件資料
Set d = ThisDocument
Set t = d.Tables(1)
For Each c In t.Columns(columnComponents).Cells
    If c.RowIndex > 1 Then '第1列是標題時
        '取得漢字
        Set a = t.Cell(c.RowIndex, columnChar).Range '漢字欄位儲存格
        w = VBA.Mid(a.Text, 1, Len(a.Text) - 2) '儲存格標記乃尾綴為 chr(13) & chr(7),要扣掉
        '取得部件陣列 componentsArray
        components.getComponentsArray componentsArray, c.Range
        
        '排序部件（此函式方法不予排序）
'        Call components.SortStringArray(componentsArray) '引數以傳址傳遞

        '鍵值為漢字
        key = w
        If dict5032ChineseComponents.Exists(key) Then
            sndPlaySound32 "C:\Windows\Media\Alarm03.wav", 1
            Stop
        Else '值為漢字部件的本來次序
            'dict5032ChineseComponents.Add key, componentsArray
            dict5032ChineseComponents(key) = componentsArray
        End If
    End If
Next c

sndPlaySound32 "C:\Windows\Media\Ring05.wav", 1
Set 常用5032漢字部件表Dictionary = dict5032ChineseComponents
If 尋找功能.ThisDocument常用5032漢字部件表Dictionary.Count = 0 Then Set 尋找功能.ThisDocument常用5032漢字部件表Dictionary = dict5032ChineseComponents

End Property

'定義一個函式來從一個陣列中取出其中少一個元素來組成集合，並將集合作為陣列返回
Function Subset1(arr As Variant) As VBA.Collection
    Dim n As Integer, m As Integer, i As Integer, k As Integer, element ', j As Integer
    Dim result As New Collection
    
    '獲取輸入陣列的元素個數，並將其減一賦給m變量
    n = UBound(arr) - LBound(arr) + 1
    m = n - 1
    
    '檢查輸入陣列是否至少有兩個元素，如果不是，則返回錯誤訊息
    If n < 2 Then
        MsgBox "輸入陣列至少要有兩個元素"
        Exit Function
    End If
    
    '利用組合數函式來計算返回陣列的大小，並將其賦給result變量，result變量是一個二維陣列
'    ReDim result(1 To Combination(n, m), 1 To m)
    
    '利用雙重迴圈來填充result陣列的每一行，每一行代表一種可能的集合
    k = 1 'k變量表示result陣列的行索引
    For i = LBound(arr) To UBound(arr) 'i變量表示要從輸入陣列中排除的元素索引
        
'        j = 1 'j變量表示result陣列的列索引
        
        For Each element In arr 'element變量表示輸入陣列中的每一個元素
            
            If element <> arr(i) Then '如果element不等於要排除的元素
                
'                result(k, j) = element '將element賦值給result陣列
                result.Add "●" + element
'                j = j + 1 '增加j變量
                
            End If
            
        Next element
        
        k = k + 1 '增加k變量
        
    Next i
    
    Set Subset1 = result '將result賦值給Subset
    
End Function

Rem 20230315 creedit with Bing大菩薩
'定義一個函式來從一個陣列中取出其中少一個元素來組成集合，並將集合作為陣列返回。真按：將此二維陣列的集合包在Collection作為Dictionary的鍵值（key）傳回。值（value）則是所忽略掉的變形部首
Function subsetExcludingVariantRadicals(arr As Variant) As Scripting.Dictionary 'Variant
    Dim n As Integer, m As Integer, i As Integer, j As Integer, k As Integer, element, flagVariantRadical As Boolean
    Dim dictVariantRadicals As Scripting.Dictionary
    Dim result As Variant, resultDict As New Scripting.Dictionary '由result變量是一個二維陣列改編、增益。resultDict是要回傳的值（key=忽略變形部首的部件組成集合，鍵值為包在Collection中的二維陣列 result；value=被忽略的變形部首）。
    Dim lineIndexUbound As Long '行索引的上界( lineIndexUbound = Combination(n, m)  )
    Dim cln As New VBA.Collection
    
    Set dictVariantRadicals = VariantRadicalsDictionary
    '獲取輸入陣列的元素個數，並將其減一賦給m變量
    n = UBound(arr) - LBound(arr) + 1
    m = n - 1
    
    '檢查輸入陣列是否至少有兩個元素，如果不是，則返回錯誤訊息
    If n < 2 Then
        MsgBox "輸入陣列至少要有兩個元素"
        Exit Function
    End If
    For Each element In arr
        If dictVariantRadicals.Exists(element) Then '「衣」部分上下者還未搞定
            flagVariantRadical = True
        End If
    Next element
    If flagVariantRadical = False Then
        Exit Function '不含變形部首者不做
    End If
    Rem Bing大菩薩：
    '利用組合數函式來計算返回陣列的大小，並將其賦給result變量，result變量是一個二維陣列 rem 真按：即元素皆為一維陣列的一個陣列（名為二維陣列：陣列內又是陣列，如此下去，為「維」。）
    'ReDim result(1 To Combination(n, m), 1 To m)
    Rem end
    lineIndexUbound = Combination(n, m) '真按：取得行索引的上界
    ReDim result(1 To lineIndexUbound, 1 To m)
    
    '利用雙重迴圈來填充result陣列的每一行，每一行代表一種可能的集合
    k = 1 'k變量表示result陣列的行索引 rem 佛弟子文獻學者孫守真任真甫按：即二維陣列本身的索引，即對其內元素（型別為一維陣列）下標的索引值。行=二維陣列
    For i = LBound(arr) To UBound(arr) 'i變量表示要從輸入陣列中排除的元素索引
            
        Rem 真按：前面排除不含變形部首的漢字，這裡排除雖含變形部首但非忽略不計變形部首時的組合
        If variantRadicalsDict.Exists(arr(i)) Then '分上下之「衣」部部件還未搞定。一樣一樣來
            
            j = 1 'j變量表示result陣列的列索引 rem 孫守真按：即二維陣列內的一維陣列元素本身的索引。列，如資料表記錄之列，則一筆記錄（一列）即二維陣列裡的一個一維陣列元素；或可想像成一個只有一列而多行的資料表，則每欄就是二維陣列裡面的每個一維陣列元素。故名為「列」
            
            Dim iarr As Byte '記下部件位置
            iarr = 0
            For Each element In arr 'element變量表示輸入陣列中的每一個元素
                If element <> arr(i) Or (element = arr(i) And iarr <> i) Then '如果element不等於要排除的元素'真按：即忽略這個變形部首 arr(i)
                    
                    result(k, j) = element '將element賦值給result陣列
                    
                    j = j + 1 '增加j變量 rem 真按：列索引（即二維陣列內的一維陣列元素本身的索引。將二維陣列裡每個一維陣列元素視作一列，故名「列」索引；列=一維陣列元素）
                    
                End If
                iarr = iarr + 1
            Next element
            
            k = k + 1 '增加k變量 rem 真按：列索引（即二維陣列的索引 → 即對其下標存取其內一維陣列元素的索引值）
            cln.Add result 'set cln= nothing 後竟然還可以加
            
            resultDict.Add cln, arr(i)  '將result（忽略變形部首的部件組成集合）二維陣列包在Collection作為鍵值（key）、arr(i)（變形部首）存入要回傳的resultDict值

'            cln.Remove 1'用此會影響加入到 resultDict中的鍵值元素
            Set cln = Nothing '用此才不會影響加入到 resultDict中的鍵值元素
            Rem 20230316實測包在Collection中是可行的，只是記得用 set = nothing 清空以備下一迭代用
            'resultDict.Add result, arr(i)
            Rem Bing大菩薩:
            '根據網路搜尋的結果，VBA 的 Dictionary 物件類別的鍵值（key）可以是任何資料類型，包括陣列1。但是，如果要使用陣列作為鍵值，需要注意以下幾點23：
            '陣列作為鍵值時，必須是一維的。
            '陣列作為鍵值時，必須有相同的元素個數和順序才能被視為相同的鍵值。
            '陣列作為鍵值時，不能直接用索引或 Keys 方法來存取，需要先轉換成 Variant 或 String 類型。
            
       Rem 真按：排除多餘的行索引
        Else
'            lineIndexUbound = lineIndexUbound - 1 '把排除掉的不納入二維陣列中的（忽略變形部首的組合不計入）空位置移除
            Rem 陣列索引超出範圍
'            ReDim Preserve result(1 To lineIndexUbound, 1 To m) '並保留二維陣列內原有已填充(填實)的元素；「充實」，同義複詞，降龍十八掌·九陽神功第2招
            Rem 您好，這是 Bing。根據網路搜尋的結果，您的程式碼在 ReDim Preserve 陳述式中出現「陣列索引超出範圍」的錯誤可能是因為以下原因123：
            '- ReDim Preserve 只支援最後一維變大小，其他維度不能變，否則會顯示越界。
            '- ReDim Preserve 只能單向變大，不應該變小。
            '- ReDim Preserve 不適合和迴圈搭配使用，因為每次迴圈都會重新分配記憶體空間，影響效能。
            '
            '您可以嘗試以下方法來解決或避免這個錯誤?：
            '- 使用 Variant 類型的陣列來儲存不同資料類型的元素。
            '- 使用 Collection 或 Dictionary 物件來動態新增或刪除元素。
            '- 預先計算陣列所需的大小，並一次性分配記憶體空間。
            '
            '希望這些資訊對您有幫助。??
            '
            '來源: 與 Bing 的交談， 2023/3/16(1) VBA ?什么你redim() ??二????出?？因? redim ????不太适合和循?搭配_奔跑的犀牛先生的博客-CSDN博客. https://blog.csdn.net/xuemanqianshan/article/details/103962274 已存取 2023/3/16.
            '(2) ReDim 陳述式 - Visual Basic | Microsoft Learn. https://learn.microsoft.com/zh-tw/dotnet/visual-basic/language-reference/statements/redim-statement 已存取 2023/3/16.
            '(3) ReDim 語句 (VBA) | Microsoft Learn. https://learn.microsoft.com/zh-tw/office/vba/language/reference/user-interface-help/redim-statement 已存取 2023/3/16.
            '(4) 動態陣列（陣列大小調整和動態處理） | 他山教程，只選擇最優質的自學材料. http://www.tastones.com/zh-tw/stackoverflow/vba/arrays/dynamic_arrays_array_resizing_and_dynamic_handling/ 已存取 2023/3/16.
        End If
        
    Next i
    
    Set subsetExcludingVariantRadicals = resultDict '
    
End Function


Rem 20230315 Bing大菩薩
'定義一個函式來從一個陣列中取出其中少一個元素來組成集合，並將集合作為陣列返回
Function Subset(arr As Variant) As Variant
    Dim n As Integer, m As Integer, i As Integer, j As Integer, k As Integer, element
    Dim result As Variant
    
    '獲取輸入陣列的元素個數，並將其減一賦給m變量
    n = UBound(arr) - LBound(arr) + 1
    m = n - 1
    
    '檢查輸入陣列是否至少有兩個元素，如果不是，則返回錯誤訊息
    If n < 2 Then
        MsgBox "輸入陣列至少要有兩個元素"
        Exit Function
    End If
    
    '利用組合數函式來計算返回陣列的大小，並將其賦給result變量，result變量是一個二維陣列
    ReDim result(1 To Combination(n, m), 1 To m)
    
    '利用雙重迴圈來填充result陣列的每一行，每一行代表一種可能的集合
    k = 1 'k變量表示result陣列的行索引
    For i = LBound(arr) To UBound(arr) 'i變量表示要從輸入陣列中排除的元素索引
        
        j = 1 'j變量表示result陣列的列索引
        
        For Each element In arr 'element變量表示輸入陣列中的每一個元素
            
            If element <> arr(i) Then '如果element不等於要排除的元素
                
                result(k, j) = element '將element賦值給result陣列
                
                j = j + 1 '增加j變量
                
            End If
            
        Next element
        
        k = k + 1 '增加k變量
        
    Next i
    
    Subset = result '將result賦值給Subset
    
End Function

'定義一個函式來計算階乘
Function Factorial(n As Integer) As Long
        Dim i As Integer, result As Long
        result = 1
        For i = 1 To n
            result = result * i
        Next i
        Factorial = result
End Function

'定義一個函式來計算組合數
Function Combination(n As Integer, m As Integer) As Long '檢查輸入參數是否合法，如果不是，則返回錯誤訊息
        If n < 0 Or m < 0 Or n < m Then
                MsgBox "輸入參數必須為非負整數且n不小於m"
                Exit Function
        End If
'利用階乘函式來計算組合數公式：C(n,m) = n! / (m! * (n-m)!)
Combination = Factorial(n) / (Factorial(m) * Factorial(n - m))
End Function


'測試程式碼，輸入一個字串，並將其轉換為字元陣列，然後調用Subset函式，並顯示返回值在訊息框中。
Sub Test()
    
    Dim inputString As String, inputArray As Variant, outputArray As Variant
    
    inputString = InputBox("請輸入一個字串")
    
    inputArray = Split(StrConv(inputString, vbUnicode), Chr$(0))
    
    outputArray = Subset(inputArray)
    
End Sub

Rem 20230316 定義一個函式來將漢字中文分解為部件。回傳是字典，鍵值是部件，值是部件所構成的漢字
Function DecomposeDict(ch As String) As Scripting.Dictionary '真按：回傳為所問取之漢字 ch 它部件的字典
    Dim result As New Scripting.Dictionary, arr, e
    arr = Decompose(ch)
    For Each e In arr
        result(e) = ch '這樣重複的部件就會消掉（如果只需要知道有哪些部件時才適用）
    Next e
    Set DecomposeDict = result
End Function

Rem 20230316 定義一個函式來將漢字中文分解為部件。回傳是集合Collection，元素值是部件
Function DecomposeCollection(ch As String) As VBA.Collection  '真按：回傳為所問取之漢字 ch 它部件的字典
    Dim result As New VBA.Collection, arr, e
    arr = Decompose(ch)
    For Each e In arr
        result.Add e '這樣重複的部件才不會被消掉，且是照原字的構成序列排序
    Next e
    Set DecomposeCollection = result
End Function

Rem 20230315~6 creedit with Bing大菩薩
'定義一個函式來將漢字中文分解為部件
Function Decompose(ch As String) As Variant '真按：回傳為所問取之漢字 ch 它部件的一維陣列
    'Dim result As Variant
    
    '利用第三方工具或自己寫的程式碼來實現分解功能
    '假設返回值是一個陣列，每個元素是一個部件
    
    'result = SomeToolOrCode(ch)
    
    'Decompose = result
    If dict5032ChineseComponents.Count = 0 Then Set dict5032ChineseComponents = Me.常用5032漢字部件表Dictionary
    Decompose = dict5032ChineseComponents(ch)
End Function

'定義一個函式來判斷兩個漢字中文是否除了部首外其他部件都相同
Function Compare(ch1 As String, ch2 As String) As Boolean
    Dim arr1 As Variant, arr2 As Variant
    Dim i As Integer, j As Integer
    Dim flag1 As Boolean, flag2 As Boolean
    
    '利用Decompose函式來將兩個漢字中文分解為陣列
    arr1 = Decompose(ch1)
    arr2 = Decompose(ch2)
    
    '如果陣列長度不同，則返回False
    If UBound(arr1) <> UBound(arr2) Then
        Compare = False
        Exit Function
    End If
    
    '利用雙重迴圈來比較兩個陣列是否除了第一個元素（即部首）外其他元素都相同
    
    For i = LBound(arr1) + 1 To UBound(arr1) 'i變量表示arr1陣列的索引
        
        flag1 = False 'flag1變量表示arr1陣列中當前元素是否在arr2陣列中存在
        
        For j = LBound(arr2) + 1 To UBound(arr2) 'j變量表示arr2陣列的索引
            
            If arr1(i) = arr2(j) Then '如果arr1陣列和arr2陣列中有相同的元素
                
                flag1 = True '將flag1變量設為True
                
                Exit For '跳出迴圈
                
            End If
            
        Next j
        
        If flag1 = False Then '如果flag1變量仍然為False
            
            Compare = False '返回False
            
            Exit Function
            
        End If
        
    Next i
    
    For i = LBound(arr2) + 1 To UBound(arr2) 'i變量表示arr2陣列的索引
        
        flag2 = False 'flag2變量表示arr2陣列中當前元素是否在arr1陣列中存在
        
        For j = LBound(arr1) + 1 To UBound(arr1) 'j變量表示arr1陣列的索引
            
            If arr2(i) = arr1(j) Then '如果arr2陣列和arr1陣列中有相同的元素
                
                flag2 = True '將flag2變量設為True
                
                Exit For '跳出迴圈
                
            End If
            
        Next j
        
        If flag2 = False Then '如果flag2變量仍然為False
            
            Compare = False '返回False
            
            Exit Function
            
        End If
        
    Next i
    
    Compare = True '如果迴圈都正常結束，則返回True
    
End Function

'定義一個函式來列出所有符合條件的漢字中文，並且將它們作為函式的返回值
Function List(ch As String) As Variant
    Dim result As Variant
    Dim dict As Variant
    Dim key As Variant
    Dim i As Integer
    
    '利用第三方工具或自己寫的程式碼來獲取漢字中文的字典，或者一個能夠將漢字中文分解為部件的工具。
    '假設返回值是一個字典物件，每個鍵是一個漢字中文，每個值是一個陣列，表示該漢字中文的部件。
    
    Set dict = dict5032ChineseComponents 'SomeToolOrCode()
    
    '利用Compare函式來判斷輸入的漢字中文是否與字典中的其他漢字中文除了部首外其他部件都相同。
    
    For Each key In dict.Keys() 'key變量表示字典物件中的每一個鍵
        
        If Compare(ch, (key)) Then '如果輸入的漢字中文和當前鍵相比除了部首外其他部件都相同
            
            ReDim Preserve result(1 To i + 1) '重新定義result陣列的大小，增加一個元素
            
            result(i + 1) = key '將當前鍵賦值給result陣列
            
            i = i + 1 '增加i變量
            
        End If
        
    Next key
    
    List = result '將result賦值給List
    
End Function

'測試程式碼，輸入一個漢字中文，然後調用List函式，並顯示返回值在訊息框中。
Sub Test1()
    
    Dim inputChar As String, outputArray As Variant
    
    inputChar = InputBox("請輸入一個漢字")
    
    outputArray = List(inputChar)
    
End Sub


Public Property Get VariantRadicalsDictionary() As Scripting.Dictionary
If variantRadicalsDict.Count > 0 Then
    Set VariantRadicalsDictionary = variantRadicalsDict
Else
    '第2段起至第47段止為變形部首字串
    Dim variantRadicalsStr As String, variantRadicalsArrSub, variantRadicalsArr, e, eSub
    variantRadicalsStr = ThisDocument.Range(ThisDocument.Paragraphs(2).Range.Start, ThisDocument.Paragraphs(47).Range.End).Text
    variantRadicalsStr = VBA.Left(variantRadicalsStr, VBA.Len(variantRadicalsStr) - 1) '去掉末尾分段符號
    variantRadicalsArr = VBA.Split(variantRadicalsStr, Chr(13))
    For Each e In variantRadicalsArr
        variantRadicalsArrSub = VBA.Split(e, ",")
        For Each eSub In variantRadicalsArrSub
            variantRadicalsDict(eSub) = variantRadicalsArrSub(0) '首元素為部首
        Next eSub
    Next e
    Set VariantRadicalsDictionary = variantRadicalsDict
End If
End Property

Public Property Let VariantRadicalsDictionary(ByVal vNewValue As Scripting.Dictionary)

End Property

Rem 取得所問漢字之部件數
Function componentsCountofChar(ch As String) As Byte
If dict5032ChineseComponents.Count = 0 Then Set dict5032ChineseComponents = 常用5032漢字部件表Dictionary
componentsCountofChar = UBound(dict5032ChineseComponents(ch)) + 1
End Function

Private Sub Class_Initialize()
If 尋找功能.ThisDocument常用5032漢字部件表Dictionary.Count > 0 Then Set dict5032ChineseComponents = 尋找功能.ThisDocument常用5032漢字部件表Dictionary
End Sub
