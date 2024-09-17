Attribute VB_Name = "Keywords"
Option Explicit
Rem 任何關鍵字檢索、標識相關之屬性、參照記錄

Rem 用以檢查是否為易學範圍之內容用
Property Get 易學KeywordsToCheck() As Variant 'string()
    易學KeywordsToCheck = Array(VBA.ChrW(-10119), VBA.ChrW(-8742), VBA.ChrW(-30233), VBA.ChrW(-10164), VBA.ChrW(-8698), VBA.ChrW(-31827), VBA.ChrW(-10132), VBA.ChrW(-8313), VBA.ChrW(20810), VBA.ChrW(-10167), VBA.ChrW(-8698), VBA.ChrW(-26587), VBA.ChrW(21093), VBA.ChrW(14615), VBA.ChrW(20089), VBA.ChrW(26080), "妄", VBA.ChrW(26083), "濟" _
        , "遘", "遁", VBA.ChrW(20089), "离", "乾", "小畜", "履", "臨", "觀", "大過", "坤", "泰", "否", "噬嗑", "賁", "坎", "屯", "蒙", "同人", "大有", "剝", "復", "離", "需", "訟", "謙", "豫", "無妄", "大畜", "師", "比", "隨", "蠱", "頤", "咸", "��", "損", "益", "震", "艮", "中孚", "遯", "大壯", "夬", "姤", "漸", "歸妹", "小過", "晉", "明夷", "萃", "升", "豐", "旅", "既濟", "未濟", "家人", "睽", "困", "井", "巽", "兌", "蹇", "解", "革", "鼎", "渙", "節", "太極", "陰陽", "兩儀", "象", "彖", _
        "老陰", "老陽", "少陰", "少陽", "蓍")
End Property
Rem 用以標識易學關鍵字用
Property Get 易學KeywordsToMark() As Variant 'string()因為 Array Returns a Variant containing an array,所以不能寫成 as string()
    易學KeywordsToMark = Array("易", "周易", "易經", "大易", "五經", "六經", "七經", "十三經", "蓍", _
        "卦", "節卦", "離卦", _
        "爻", _
        "系辭", "繫辭", "擊辭", "擊詞", "繫詞", "說卦", "序卦", _
            "卦序", "敘卦", "雜卦", "文言", "乾坤", "元亨", "利貞", "史記", _
        "筮", "夬", "乾", "〈乾〉", "〈坤〉", "乾、坤", "〈乾、坤〉", "噬嗑", "賁", VBA.ChrW(20089), "既濟", VBA.ChrW(26083) & "濟", "未濟", "十翼", _
        "大" & VBA.ChrW(22766), _
        "初九", "九二", "九三", "九四", "九五", "上九", VBA.ChrW(19972) & "九", "用九", "初六", "六二", "六三", "六四", "六五", "上六", "用六", _
        "河圖", "洛書", "太極", "無極", _
            "象曰", "〈象〉曰", "象日", "象云", "象傳", "〈大象〉", "小象", "象義", "彖", _
            "艮", "頤", "坎", "中孚", "兌", "蠱", "姤", "巽", VBA.ChrW(14514), "剝", VBA.ChrW(21093), "遯", "大壯", "明夷", "明" & VBA.ChrW(-10171) & VBA.ChrW(-8739), "小畜", "大畜", "萃", "蹇", "渙", VBA.ChrW(28067), "睽", "暌", "歸妹", "小過", "大有", "大過", "〈泰〉", "〈否〉", "〈損〉", "〈益〉", "〈屯〉", "豫", "〈旡妄〉", "〈復〉", "〈震〉", _
            "老陰", "老陽", "少陰", "少陽", "繇辭", "繇詞", _
            "咎", "咸��", "咸恆", _
        "無妄", VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522), VBA.ChrW(26080) & "妄", _
        "無咎", VBA.ChrW(26080) & "咎", "天咎", "允升", _
        "隨時之義", "庖有魚", "包有魚", "精義入神", "豶豕", "童牛", "承之羞", "雷在天上", "錫馬", "蕃庶", "晝日", "三接", "懲忿", "窒欲", "懲窒", "敬以直內", "義以方外", "迷後得主", "利西南", "品物咸章", "天下大行", "益動而", "日進無疆", "日進" & VBA.ChrW(26080) & "疆", "頻巽", "頻" & VBA.ChrW(14514), "豚魚", "頻復", "閑邪", "存誠", "乾乾", "悔吝", "憧憧", "類萬物", "柔順利貞", VBA.ChrW(-10163) & VBA.ChrW(-9167) & "順利貞", "比之匪人", "履貞", "貞厲", "履道坦坦", "貞吉", "直方", "木上有水", "勞民勸相", "索而得", _
        "悔亡", "悔" & VBA.ChrW(20158), "悔" & VBA.ChrW(20838), "時義", "健順", "內健而外順", "內健外順", "外順而內健", "外順內健", "易簡", "易" & VBA.ChrW(-10153) & VBA.ChrW(-9007), "敦復", "開物成務", "窮神知化", "研幾極深", "極深研幾", "見善則遷", "有過則改", "遷善改過", "夕惕", "惕若", "一陰一陽", _
        "伏羲")
        
End Property
Rem 某關鍵字的前面不能是 20240914
Property Get 易學KeywordsToMark_ExamPrecededAvoid() As Scripting.Dictionary
    
    
    Dim dict As New Scripting.Dictionary, cln As New VBA.Collection
    ' 添加資料到字典 creedit_with_Copilot大菩薩：https://sl.bing.net/goDF239cQVw
    dict.Add "易", Array("移", "平仄相", "錢", "光庭、賈", "光庭賈", "為人和", "辟", "驕", "一氣一", "資和", "新陳相", "捷最", "崔伯", "劉", "有不善未嘗不知", "事難慮", "事久則慮", "勢固", "立門戶也", "立門" & ChrW(25143) & "也", "聽之者", "厚和", "誠不為", "可", "人欲", "市", "輒", ChrW(-28903), "遽", "過於和", "平心", "尤", "容", "未", "不", "極", "甚", "貿", "交", "物", "變", "或可", "鄙", "博", "辟", "平", "慢", "俗", "坦", "難", "脫", "流", "樂", "革", "更", "簡", "白居", "居", "淺", "輕", "險", "相", "難行", "世", "易", "所")
    dict.Add "乾", Array("白", "豆", "面自", "擰", "餅", "未", "晾", "肉", "蘿蔔", "葡萄", "龍眼", "口", "枯", "烘", "晒", "曬", "筍", "外強中")
    dict.Add "乾坤", Array("搆盡", "于此盜")
    dict.Add "豫", Array("防患於", "暇", "厎", "底", "不", "劉", "猶", "逸", VBA.ChrW(-10143) & VBA.ChrW(-8996))
    dict.Add "剝", Array("刻", "皴", "可", "為之解", "歲蹇")
    dict.Add VBA.ChrW(21093), Array("可", "刻", "為之解", "歲蹇")
    dict.Add "頤", Array("周敦", "程", "朵", "筆支", "頂至", "泗交")
    dict.Add VBA.ChrW(-26587), Array("周敦", "程", "朵", "筆支", "頂至")
    dict.Add "巽", Array("李", "翟公", "傅")
    dict.Add VBA.ChrW(14514), Array("李", "翟公", "傅")
    dict.Add "大過", Array("可過")
    dict.Add "賁", Array("諸葛", "齎」作「", "齎作")
    dict.Add "蹇", Array("偃", "矯", "策", "奇")
    dict.Add "夬", Array("龔")
    dict.Add "中孚", Array("周")
    dict.Add "小過", Array("吏有")
    dict.Add "渙", Array("崔", "蘇")
    dict.Add VBA.ChrW(28067), Array("崔", "蘇")
    dict.Add "蠱", Array("韓", "置", "蟲", "可以解", "蛇", "年之", "下", "之立", "謂水", "每遇")
    dict.Add "萃", Array("辭拔")
    dict.Add "睽", Array("暌」作「", "暌作")
    dict.Add "大有", Array("歲稱", "後弟", "來曰", "花朵")
    dict.Add "既濟", Array("沈", "沉")
    dict.Add "九二", Array("卷四", "卷", "卷二")
    dict.Add "九三", Array("卷", "卷一")
    dict.Add "六二", Array("卷", "卷一", "卷二")
    dict.Add "六三", Array("卷")
    dict.Add "六四", Array("卷", "卷一", "…　一")
    dict.Add "象云", Array("光景氣")
    dict.Add "文言", Array("與上")
    dict.Add "咎", Array("身之休", "晁無咎", "晁丈無咎", "殃", "追", "自", "引", "何", "專", "知休", "卜休", "知人休", "能知休")
    dict.Add "元亨", Array("董")
    dict.Add "貞吉", Array("曹")
    dict.Add "悔吝", Array("平生")
    dict.Add "無咎", Array("晁", "晁丈")
    dict.Add "直方", Array("張")
    dict.Add "敦復", Array("張")
    dict.Add "易簡", Array("蘇")
    dict.Add "易" & VBA.ChrW(-10153) & VBA.ChrW(-9007), Array("蘇")
    
    Set 易學KeywordsToMark_ExamPrecededAvoid = dict

        
End Property
Rem 某關鍵字的後面不能是 20240914
Property Get 易學KeywordsToMark_ExamFollowedAvoid() As Scripting.Dictionary
    Dim dict As New Scripting.Dictionary
    dict.Add "易", Array("棺", "肆", "字希", "渡", "得汩沒", "子所作", "易生粉", "發酸", "視之", "紙易之", "腐敗", "善作詩", "著火", "於噴", "以走險", "不能堪", "破", "安居士", "名者", "名之日", "曉", "得鹵", "十姓", "言則難", "搖而難", "昏而難", "下手", "萌", "慢", "得消散", "見者", "厭", "出議論", "差", "陷", "退之節", "看", "積", "忘", "物", "易", "俗", "卜生", "堂", "科", "開罐", "筋", "姓", "玄光", "梓宮", "知由單", "幟", "轍", "手", "守", "水", "州", "順鼎", "如反", "如翻", "如拾", "容", "熔", "子而", "簀", "牙", "肩輿", "事爾", "但不香")
    dict.Add "卦", Array("陣", "橋", "建築")
    dict.Add "筮", Array("仕")
    dict.Add "乾", Array("枯", "淨", _
        "道初", "道元年", "寧", "祐", _
        "和中")
    dict.Add "乾坤", Array("陷吉人")
    dict.Add "豫", Array("章", "讓", "州", "暇", "知", "劇", "備", VBA.ChrW(20675), "防", "聞", _
        "先要", "豫為言之", "豫" & VBA.ChrW(29234) & "言之")
    dict.Add "剝", Array("春" & VBA.ChrW(-31631), "春蔥", "芋", "削", "民", "蝕", "泐", "落", "去腸", "人面", "婦人衣")
    dict.Add VBA.ChrW(21093), Array("芋", "削", "民", "蝕", "泐", "落", "去腸", "人面", "婦人衣")
    dict.Add "蹇", Array("諤")
    dict.Add "渙", Array("然", "散", "遂踰")
    dict.Add VBA.ChrW(28067), Array("然", "散", "遂踰")
    dict.Add "頤", Array("和園", "正叔", "字正", "茂叔", "字茂", "下有皮")
    dict.Add VBA.ChrW(-26587), Array("和園", "正叔", "字正", "茂叔", "字茂", "下有皮")
    dict.Add "萃", Array("此書", "于一")
    dict.Add "艮", Array("嶽", "岳")
    dict.Add "賁", Array("隅之")
    dict.Add "巽", Array("懦", "字仲", "亦不較", "後仕", "風疏")
    dict.Add VBA.ChrW(14514), Array("懦", "字仲", "亦不較", "後仕", "風疏")
    dict.Add "蠱", Array("自埋", "之詐", "實生子", "毒", "發膨", "主", "之屬", "有鬼", "不絕", "者是也", "嫩", VBA.ChrW(23280))
    dict.Add "暌", Array("」作「睽", "作睽")
    dict.Add "坎", Array("鼓")
    dict.Add "遯", Array("世修真")
    dict.Add "小畜", Array("集")
    dict.Add "大有", Array("父風", "力", "警省")
    dict.Add "大過", Array("人者")
    dict.Add "小過", Array("宜寬", "宜" & VBA.ChrW(23515))
    dict.Add "初九", Array("日")
    dict.Add "初六", Array("日")
    dict.Add "咎", Array("繇", "彼")
    dict.Add "敦復", Array("學士")

    
    Set 易學KeywordsToMark_ExamFollowedAvoid = dict

End Property
Rem 某關鍵字不能在某個語句裡面 20240914
Property Get 易學KeywordsToMark_ExamInPhraseAvoid() As Scripting.Dictionary
    Dim dict As New Scripting.Dictionary
    dict.Add "易", Array("李易安", "居易錄", "蘇易簡", "杜易簡", "張易之", "此易事", "人易老", "人易去", "兩易其任", _
        "江山易主", "深耕易耨", "豈易逢", "延易府", "皮易布", "金易一", "名易知", "雖易得", "誠至易", "則易束", "者易窺", "置易制", "難、易相", "事易行", "今易以冠服", "俱易以", "易知也", "市易法", "可以易一飽", "賈易不", "和易之氣", _
        "而易見", "而易起", "而易晦", "最易得", "亢易招", "以易鹽米", "頓易故", "谿易雨", "者易訓", "心易偏", "平心易氣", "故其說易差", "始易為力", "乃易合", "而易彼", "慢而易之", "而易治", "樂易之", "視而易之", _
        "以易心", "以易處之", "因易其韻", "以縉儒者易之", "因改易本文而", "則易發", "根易發", _
        "而易陵", "儉則易足", "者易犯", "狹則易足", "智易窮", "何以易窮", "無以易此", "更有易見者", "而易散", _
        "至易事", "敢易也", "河東、易定", "易定、魏博", "河東易定", "脆易折", "易定魏博", "柔易治", "最易生", _
        "疾易作", "病易除", "人易從", "易放而難操", "錢差易 ", "言易墜", "所以易放", "者易直", "是易言也", "須易之", "之易以仰測", "蓴絲之易", "皆易黃屋", "皆易" & ChrW(-24892) & "屋", "時易以新", _
        "則易使", "侮易承業", "成易具", "綿布易之", "布相易云", "浴易服", _
        "因易名曰", "不易長", _
        "則易入於", _
        "非易事", _
        "豈易說")
    dict.Add "卦", Array("八卦山")
    dict.Add "乾", Array("大乾廟", "乳乾者")
    dict.Add "剝", Array("解剝而發明", "造剝洛陽")
    dict.Add VBA.ChrW(21093), Array("解" & VBA.ChrW(21093) & "而發明", "造" & VBA.ChrW(21093) & "洛陽")
    dict.Add "豫", Array("人豫知")
    dict.Add "蹇", Array("剛蹇絕", "歲蹇剝")
    dict.Add "頤", Array("翁頤昌", "方頤大口")
    dict.Add VBA.ChrW(-26587), Array("翁" & VBA.ChrW(-26587) & "昌", "方" & VBA.ChrW(-26587) & "大口")
    dict.Add "暌", Array("有暌談笑")
    dict.Add "蠱", Array("以蠱留人", "以蠱而", "而蠱者", "以蠱大", "中蠱者")
    dict.Add "賁", Array("古賁灰")
    dict.Add "巽", Array("即巽也", "東巽泉")
    dict.Add VBA.ChrW(14514), Array("即" & VBA.ChrW(14514) & "也", "東" & VBA.ChrW(14514) & "泉")
    dict.Add "同人", Array("底同人不", "不同人生", "不同人能")
    dict.Add "大有", Array("甚大有朱")
    dict.Add "大過", Array("無大過惡")
    dict.Add "無妄", Array("人無妄取", "物無妄費")
    dict.Add VBA.ChrW(26080) & "妄", Array("人" & VBA.ChrW(26080) & "妄取", "物" & VBA.ChrW(26080) & "妄費")
    dict.Add VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522), Array("人" & VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522) & "取", "物" & VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522) & "費")
    dict.Add "初九", Array("虞初九百")
    dict.Add "九二", Array("一九二０")
    dict.Add "九三", Array("廿九三十")
    dict.Add "存誠", Array("心存誠敬")
        
    Set 易學KeywordsToMark_ExamInPhraseAvoid = dict

End Property
Rem 檢測關鍵字
'Function 易學KeywordsToMark_Exam()
'
'End Function
