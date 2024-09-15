Attribute VB_Name = "Keywords"
Option Explicit
Rem 任何關鍵字檢索、標識相關之屬性、參照記錄

Rem 用以檢查是否為易學範圍之內容用
Property Get 易學KeywordsToCheck() As Variant 'string()
    易學KeywordsToCheck = Array(VBA.ChrW(-10119), VBA.ChrW(-8742), VBA.ChrW(-30233), VBA.ChrW(-10164), VBA.ChrW(-8698), VBA.ChrW(-31827), VBA.ChrW(-10132), VBA.ChrW(-8313), VBA.ChrW(20810), VBA.ChrW(-10167), VBA.ChrW(-8698), VBA.ChrW(-26587), VBA.ChrW(21093), VBA.ChrW(14615), VBA.ChrW(20089), VBA.ChrW(26080), "妄", VBA.ChrW(26083), "濟" _
        , "遘", "遁", VBA.ChrW(20089), "离", "乾", "小畜", "履", "臨", "觀", "大過", "坤", "泰", "否", "噬嗑", "賁", "坎", "屯", "蒙", "同人", "大有", "剝", "復", "離", "需", "訟", "謙", "豫", "無妄", "大畜", "師", "比", "隨", "蠱", "頤", "咸", "恒", "損", "益", "震", "艮", "中孚", "遯", "大壯", "夬", "姤", "漸", "歸妹", "小過", "晉", "明夷", "萃", "升", "豐", "旅", "既濟", "未濟", "家人", "睽", "困", "井", "巽", "兌", "蹇", "解", "革", "鼎", "渙", "節", "太極", "陰陽", "兩儀", "象", "彖", _
        "老陰", "老陽", "少陰", "少陽")
End Property
Rem 用以標識易學關鍵字用
Property Get 易學KeywordsToMark() As Variant 'string()
    易學KeywordsToMark = Array("易", "周易", "易經", "大易", "五經", "六經", "七經", "十三經", _
        "卦", "節卦", "離卦", _
        "爻", _
        "系辭", "繫辭", "擊辭", "擊詞", "繫詞", "說卦", "序卦", _
            "卦序", "敘卦", "雜卦", "文言", "乾坤", "元亨", "利貞", "史記", _
        "筮", "夬", "乾", "〈乾〉", "〈坤〉", "乾、坤", "〈乾、坤〉", "噬嗑", "賁", VBA.ChrW(20089), "既濟", VBA.ChrW(26083) & "濟", "未濟", "十翼", _
        "大" & VBA.ChrW(22766), _
        "初九", "九二", "九三", "九四", "九五", "上九", VBA.ChrW(19972) & "九", "用九", "初六", "六二", "六三", "六四", "六五", "上六", "用六", _
        "河圖", "洛書", "太極", "無極", _
            "象曰", "〈象〉曰", "象日", "象云", "象傳", "〈大象〉", "小象", "象義", "彖", _
            "艮", "頤", "坎", "中孚", "兌", "蠱", "姤", "巽", VBA.ChrW(14514), "剝", "遯", "大壯", "明夷", "明" & VBA.ChrW(-10171) & VBA.ChrW(-8739), "小畜", "大畜", "萃", "蹇", "渙", VBA.ChrW(28067), "睽", "暌", "歸妹", "小過", "大有", "大過", "〈泰〉", "〈否〉", "〈損〉", "〈益〉", "〈屯〉", "豫", "〈旡妄〉", "〈復〉", "〈震〉", _
            "老陰", "老陽", "少陰", "少陽", "繇辭", "繇詞", _
            "咎", "咸恒", "咸恆", _
        "無妄", VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522), VBA.ChrW(26080) & "妄", _
        "無咎", VBA.ChrW(26080) & "咎", "天咎", _
        "隨時之義", "庖有魚", "包有魚", "精義入神", "豶豕", "童牛", "承之羞", "雷在天上", "錫馬", "蕃庶", "晝日", "三接", "懲忿", "窒欲", "懲窒", "敬以直內", "義以方外", "迷後得主", "利西南", "品物咸章", "天下大行", "益動而", "日進無疆", "日進" & VBA.ChrW(26080) & "疆", "頻巽", "頻" & VBA.ChrW(14514), "豚魚", "頻復", "閑邪", "存誠", "乾乾", "悔吝", "憧憧", "類萬物", "柔順利貞", VBA.ChrW(-10163) & VBA.ChrW(-9167) & "順利貞", "比之匪人", "履貞", "貞厲", "履道坦坦", "貞吉", "直方", _
        "悔亡", "悔" & VBA.ChrW(20158), "悔" & VBA.ChrW(20838), "時義", "健順", "內健而外順", "內健外順", "外順而內健", "外順內健", "易簡", "易" & VBA.ChrW(-10153) & VBA.ChrW(-9007), "敦復", "開物成務", "窮神知化", "研幾極深", "極深研幾", "見善則遷", "有過則改", "遷善改過", "夕惕", "惕若", "一陰一陽", _
        "伏羲")
        
End Property
Rem 某關鍵字的前面不能是 20240914
Property Get 易學KeywordsToMark_ExamPrecededAvoid() As Scripting.Dictionary
    
    
    Dim dict As New Scripting.Dictionary, cln As New VBA.Collection
    ' 添加資料到字典 creedit_with_Copilot大菩薩：https://sl.bing.net/goDF239cQVw
    dict.Add "易", Array("移", "光庭、賈", "光庭賈", "驕", "資和", "新陳相", "捷最", "崔伯", "劉", "有不善未嘗不知", "事難慮", "事久則慮", "勢固", "立門戶也", "立門" & ChrW(25143) & "也", "聽之者", "厚和", "誠不為", "可", "人欲", "市", "輒", ChrW(-28903), "遽", "過於和", "平心", "尤", "容", "未", "不", "極", "甚", "貿", "交", "物", "變", "或可", "鄙", "博", "辟", "平", "慢", "俗", "坦", "難", "脫", "流", "樂", "革", "更", "簡", "白居", "居", "淺", "輕", "險", "相", "難行", "世", "易", "所")
    dict.Add "乾", Array("白", "豆", "面自", "擰", "餅", "未", "晾", "肉", "蘿蔔", "葡萄", "龍眼", "口", "枯", "烘", "晒", "曬", "筍", "外強中")
    dict.Add "豫", Array("防患於", "暇", "厎", "底", "不", "劉", "猶", "逸", VBA.ChrW(-10143) & VBA.ChrW(-8996))
    dict.Add "剝", Array("刻", "為之解")
    dict.Add "頤", Array("周敦", "程", "朵")
    dict.Add "大過", Array("可過")
    dict.Add "咎", Array("引", "何", "專")
    dict.Add "賁", Array("諸葛")
    dict.Add "貞吉", Array("曹")
    dict.Add "易簡", Array("蘇")
    dict.Add "易" & VBA.ChrW(-10153) & VBA.ChrW(-9007), Array("蘇")
    
    Set 易學KeywordsToMark_ExamPrecededAvoid = dict

        
End Property
Rem 某關鍵字的後面不能是 20240914
Property Get 易學KeywordsToMark_ExamFollowedAvoid() As Scripting.Dictionary
    Dim dict As New Scripting.Dictionary
    dict.Add "易", Array("肆", "得汩沒", "善作詩", "不能堪", "安居士", "名者", "得鹵", "十姓", "言則難", "搖而難", "昏而難", "下手", "萌", "慢", "得消散", "見者", "厭", "出議論", "差", "陷", "退之節", "看", "積", "忘", "物", "易", "俗", "卜生", "堂", "科", "開罐", "筋", "姓", "玄光", "梓宮", "知由單", "幟", "轍", "手", "守", "水", "州", "順鼎", "如反", "如翻", "如拾", "容", "熔", "子而", "簀", "牙", "肩輿", "事爾")
    dict.Add "卦", Array("陣", "橋", "建築")
    dict.Add "筮", Array("仕")
    dict.Add "乾", Array("枯", "淨", _
        "道初", "道元年")
    dict.Add "豫", Array("先要", _
        "知", "讓", "劇", "備", VBA.ChrW(20675), "防", "州", "章", "聞")
    dict.Add "剝", Array("削", "民")
    dict.Add "煥", Array("然", "散")
    dict.Add "頤", Array("和園")
    dict.Add "萃", Array("此書")
    dict.Add "巽", Array("懦")
    dict.Add VBA.ChrW(14514), Array("懦")
    dict.Add "大過", Array("人者")
    dict.Add "小過", Array("宜寬", "宜" & VBA.ChrW(23515))
    dict.Add "咎", Array("繇", "彼")
    
    
    Set 易學KeywordsToMark_ExamFollowedAvoid = dict

End Property
Rem 某關鍵字不能在某個語句裡面 20240914
Property Get 易學KeywordsToMark_ExamInPhraseAvoid() As Scripting.Dictionary
    Dim dict As New Scripting.Dictionary
    dict.Add "易", Array("李易安", "居易錄", "蘇易簡", _
        "江山易主", "深耕易耨", "豈易逢", "延易府", "名易知", "誠至易", "則易束", "者易窺", "置易制", "難、易相", "事易行", "今易以冠服", "俱易以", "易知也", "市易法", "可以易一飽", "賈易不", "和易之氣", _
        "而易見", "而易起", "而易晦", "最易得", "者易訓", "心易偏", "平心易氣", "故其說易差", "始易為力", "乃易合", "而易彼", "慢而易之", "而易治", "樂易之", _
        "以易心", "以易處之", _
        "而易陵", "儉則易足", "者易犯", "狹則易足", "智易窮", "何以易窮", "無以易此", "更有易見者", _
        "至易事", "敢易也", _
        "疾易作", "病易除", "人易從", "易放而難操", "錢差易 ", "言易墜", "所以易放", "者易直", "是易言也", "須易之", _
        "則易使", _
        "則易入於", _
        "非易事", _
        "豈易說")
    dict.Add "卦", Array("八卦山")
    dict.Add "乾", Array("大乾廟")
    dict.Add "剝", Array("解剝而發明")
    dict.Add "豫", Array("人豫知")
    dict.Add "頤", Array("翁頤昌")
    dict.Add "同人", Array("底同人不", "不同人生", "不同人能")
    dict.Add "無妄", Array("人無妄取", "物無妄費")
    dict.Add VBA.ChrW(26080) & "妄", Array("人" & VBA.ChrW(26080) & "妄取", "物" & VBA.ChrW(26080) & "妄費")
    dict.Add VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522), Array("人" & VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522) & "取", "物" & VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522) & "費")
    dict.Add "初九", Array("虞初九百")
        
    Set 易學KeywordsToMark_ExamInPhraseAvoid = dict

End Property
Rem 檢測關鍵字
Function 易學KeywordsToMark_Exam()
    
End Function
