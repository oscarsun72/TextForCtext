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
        "卦", "節卦", "離卦", "屯蒙", "屯" & VBA.ChrW(-10132) & VBA.ChrW(-8313), "屯、蒙", "屯、" & VBA.ChrW(-10132) & VBA.ChrW(-8313), _
        "爻", _
        "系辭", "繫辭", "擊辭", "擊詞", "繫詞", "說卦", "序卦", _
            "卦序", "敘卦", "雜卦", "文言", "乾坤", "元亨", "利貞", "史記", _
        "筮", "夬", "乾知大始", "坤作成物", "乾以易知", "坤以簡能", "乾", "〈乾〉", "〈坤〉", "乾、坤", "〈乾、坤〉", "噬嗑", "賁于外", "賁於外", "外賁", "內賁", "賁", VBA.ChrW(20089), "既濟", VBA.ChrW(26083) & "濟", "未濟", "十翼", _
        "大" & VBA.ChrW(22766), _
        "初九", "九二", "九三", "九四", "九五", "上九", VBA.ChrW(19972) & "九", "用九", "初六", "六二", "六三", "六四", "六五", "上六", "用六", _
        "河圖", "洛書", "太極", "無極", "兩儀", _
            "象曰", "〈象〉曰", "象日", "象云", "象傳", "〈大象〉", "小象", "象義", "四象", VBA.ChrW(-10145) & VBA.ChrW(-9156), "象：", "象文", _
            "彖", _
             "艮其背", "艮", "頤", "同人于宗", "同人", "坎", "中孚", "兌", "蠱", "姤", "巽", VBA.ChrW(14514), "剝", VBA.ChrW(21093), "遯世無悶", "遯世" & ChrW(26080) & "悶", "遯", "大壯", "明夷", "明" & VBA.ChrW(-10171) & VBA.ChrW(-8739), "小畜", "大畜", "萃", "蹇", "渙", VBA.ChrW(28067), "睽", "暌", "歸妹", "小過", "大有", "大過", "〈泰〉", "〈否〉", "〈損〉", "〈益〉", "〈屯", "蒙〉", VBA.ChrW(-10132) & VBA.ChrW(-8313) & "〉", "豫", "〈旡妄〉", "〈復〉", "〈震〉", "〈需〉", _
            "老陰", "老陽", "少陰", "少陽", "繇辭", "繇詞", _
            "咎", "咸��", "咸恆", _
        "無妄", VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522), VBA.ChrW(26080) & "妄", _
        "無咎", VBA.ChrW(26080) & "咎", "天咎", "允升", "屯其膏", _
        "隨時之義", "庖有魚", "包有魚", "精義入神", "豶豕", "童牛", "承之羞", "雷在天上", "錫馬", "蕃庶", "晝日", "三接", "懲忿", "窒欲", "懲窒", "敬以直內", "義以方外", "迷後得主", "利西南", "品物咸章", "天下大行", "益動而", "日進無疆", "日進" & VBA.ChrW(26080) & "疆", "頻巽", "頻" & VBA.ChrW(14514), "豚魚", "頻復", "閑邪", "存誠", "乾乾", "悔吝", "憧憧", "類萬物", "柔順利貞", VBA.ChrW(-10163) & VBA.ChrW(-9167) & "順利貞", "比之匪人", "履貞", "貞厲", "履道坦坦", "貞吉", "直方", "木上有水", "勞民勸相", "索而得", _
        "無悶", ChrW(26080) & "悶", "悔亡", "悔" & VBA.ChrW(20158), "悔" & VBA.ChrW(20838), "時義", "健順", "內健而外順", "內健外順", "外順而內健", "外順內健", "易簡", "易" & VBA.ChrW(-10153) & VBA.ChrW(-9007), "敦復", "開物成務", "窮神知化", "研幾極深", "極深研幾", "見善則遷", "有過則改", "遷善改過", "夕惕", "惕若", "一陰一陽", "我有好爵", "言有序", "有聖人之道四", "長子帥師", "弟子輿尸", "日用而不知", "之道鮮", "原始反終", "寂然不動", "感而遂通", "朋從", "朋盍", "容民畜眾", "容民畜" & VBA.ChrW(-30650), "養正", "養賢", "知臨", "臨大君", "默而成之", VBA.ChrW(-24871) & "而成之", "不言而信", "存乎德行", "通天下之志", "履正", "繼之者善", "仁者見之", "知者見之", "智者見之", _
        "大貞", "小貞", "帝出乎震", "帝出於震", "帝出于震", "與時偕行", "盈虛", "盈" & VBA.ChrW(-31142), "盈" & VBA.ChrW(-10119) & VBA.ChrW(-8991), "盈" & VBA.ChrW(-31145), _
        "伏羲", "伏" & VBA.ChrW(-10152) & VBA.ChrW(-8255), "庖" & VBA.ChrW(-10152) & VBA.ChrW(-8255), "庖羲", "宓羲", "宓" & VBA.ChrW(-10152) & VBA.ChrW(-8255), "宓犧", "伏犧", "庖犧")
        
End Property
Rem 某關鍵字的前面不能是 20240914
Property Get 易學KeywordsToMark_ExamPrecededAvoid() As Scripting.Dictionary
    
    
    Dim dict As New Scripting.Dictionary, cln As New VBA.Collection
    ' 添加資料到字典 creedit_with_Copilot大菩薩：https://sl.bing.net/goDF239cQVw
    dict.Add "易", Array("尤", "容", "未", "不", "極", "甚", "貿", "交", "物", "變", "或可", "鄙", "博", "辟", "平", "慢", "俗", "坦", "難", "脫", "流", "樂", "革", "更", "簡", "白居", "居", "淺", "輕", "險", "相", "難行", "世", "易", "所", _
        "移", "平仄相", "遂", "每歲一", "錢", "光庭、賈", "光庭賈", "為人和", "辟", "驕", "一氣一", "資和", _
        "新陳相", "捷最", "崔伯", "劉", "覆墜之", "有不善未嘗不知", "事難慮", "事久則慮", "勢固", "立門戶也", "立門" & VBA.ChrW(25143) & "也", _
        "聽之者", "壞真從", "壞" & VBA.ChrW(30494) & "從", "厚和", "誠不為", "可", "人欲", "市", "輒", VBA.ChrW(-28903), "遽", "過於和", "平心", "大樂必", _
        "知其至", "立節行、", "立節行", "圖難於其")
    dict.Add "乾", Array("白", "豆", "面自", "擰", "餅", "未", "晾", "肉", "蘿蔔", "葡萄", "龍眼", "口", "枯", "烘", "晒", "曬", "筍", "外強中")
    dict.Add "乾坤", Array("搆盡", "于此盜")
    dict.Add "豫", Array("防患於", "暇", "厎", "音", "底", "不", "弗", "劉", "猶", "逸", VBA.ChrW(-10143) & VBA.ChrW(-8996), "道", "南")
    dict.Add "剝", Array("刻", "活", "可", "褫", "皴", "為之解", "歲蹇", "石斷")
    dict.Add VBA.ChrW(21093), Array("刻", "活", "可", "褫", "為之解", "歲蹇", "石斷")
    dict.Add "頤", Array("周敦", "程", "朵", "濬", "期", "面豐", "頂至", "泗交", "張", _
                    "獨支", "筆支", "手支", _
                    "寄藥與")
    dict.Add VBA.ChrW(-26587), Array("周敦", "程", "朵", "濬", "期", "面豐", "頂至", "泗交", "張", _
                    "獨支", "筆支", "手支", _
                    "寄藥與")
    dict.Add "巽", Array("李", "翟公", "傅", "家之", "叔")
    dict.Add VBA.ChrW(14514), Array("李", "翟公", "傅", "家之", "叔")
    dict.Add "兌", Array("李")
    dict.Add VBA.ChrW(20817), Array("李")
    dict.Add VBA.ChrW(20810), Array("李")
    dict.Add "大過", Array("可過")
    dict.Add "賁", Array("孟", "孫", "諸葛", "齎」作「", "齎作")
    dict.Add "蹇", Array("偃", "矯", "策", "奇")
    dict.Add "夬", Array("龔")
    dict.Add "中孚", Array("周", "僧")
    dict.Add "小過", Array("吏有")
    dict.Add "渙", Array("崔", "蘇", "程", "畔", "本作", "士", "黃", VBA.ChrW(-24892), "謁")
    dict.Add VBA.ChrW(28067), Array("崔", "蘇", "程", "畔", "本作", "士", "黃", VBA.ChrW(-24892), "謁")
    dict.Add "蠱", Array("巫", "韓", "置", "蟲", "可以解", "蛇", "年之", "下", "之立", "謂水", "每遇")
    dict.Add "萃", Array("拔", "蓊", "拔乎其", "悉")
    dict.Add "睽", Array("暌」作「", "暌作")
    dict.Add "遯", Array("毅然知肥")
    dict.Add "同人", Array("招", "儲")
    dict.Add "大有", Array("歲稱", "後弟", "來曰", "花朵", "葉")
    dict.Add "噬嗑", Array("令")
    dict.Add "既濟", Array("沈", "沉")
    dict.Add "初九", Array("月")
    dict.Add "九二", Array("卷四", "卷", "卷二", "一一", "一五")
    dict.Add "九三", Array("卷", "卷一", "一五")
    dict.Add "九四", Array("卷", "卷一", "一五", "張")
    dict.Add "九五", Array("卷", "卷一", "一五")
    dict.Add "初六", Array("月")
    dict.Add "六二", Array("卷", "卷一", "卷二")
    dict.Add "六三", Array("卷")
    dict.Add "六四", Array("卷", "卷一", "…　一", "（九")
    dict.Add "彖", Array("張")
    dict.Add "象云", Array("郭", "皇", "光景氣")
    dict.Add "文言", Array("與上")
    dict.Add "筮", Array("初", "再")
    dict.Add "咎", Array("得", "歸", "不", "過", "追", "休", "厥", "自", "引", "殃", "何", "晁無", "晁丈無", "足", "受其", "重其", "任其", _
        "專", "示", "將有", "怨") ', "知休", "卜休", "知人休", "能知休", "身之休"
    dict.Add "元亨", Array("董", "萬")
    dict.Add "貞吉", Array("曹")
    dict.Add "悔吝", Array("平生")
    dict.Add "無咎", Array("晁", "晁丈")
    dict.Add "直方", Array("張", "王")
    dict.Add "敦復", Array("張")
    dict.Add "存誠", Array("游操")
    dict.Add "易簡", Array("蘇")
    dict.Add "易" & VBA.ChrW(-10153) & VBA.ChrW(-9007), Array("蘇")
    dict.Add "索而得", Array("豈窮")
    dict.Add "養正", Array("劉")
    
    Set 易學KeywordsToMark_ExamPrecededAvoid = dict

        
End Property
Rem 某關鍵字的後面不能是 20240914
Property Get 易學KeywordsToMark_ExamFollowedAvoid() As Scripting.Dictionary
    Dim dict As New Scripting.Dictionary
    dict.Add "易", Array("簀", "牙", "之以書契", "安居士", "幟", "轍", "厭", "水", "州", "順鼎", "破", "開罐", "筋", "姓", "名之日", "卜生", "堂", "科", _
        "守", "棺", "看", "積", "容", "萌", "葬", "獄", "其處", "名者", "曉", "得消散", "慢", "熔", "忘", "物", "易", "俗", "差", "陷", "肆", "下手", "渡", "紙易之", "腐敗", _
         "手", "如反", "善作詩", "著火", "於噴", "以走險", "不能堪", "直不足言", "以動人", "得鹵", "十姓", "易其形者為", "言則難", "視之", _
        "搖而難", "昏而難", "字希", "見者", "於近者。非知言者也", "於近者非知言者也", "出議論", "退之節", "玄光", "梓宮", "知由單", "占偶書", _
        "如翻", "如拾", "子而", "肩輿", "事爾", "但不香", "知而不知", "舊榜", "得汩沒", "子所作", "易生粉", "發酸", "如燎毛")
    dict.Add "卦", Array("陣", "橋", "建築")
    dict.Add "筮", Array("仕")
    dict.Add "乾", Array("淨", "隆", "寧", "祐", _
        "道初", "道元年", "道二年", "道三年", "道四年", "道五年", _
        "和中", "枯", "坤清泰", "坤之清氣")
    dict.Add "乾坤", Array("陷吉人", "清泰", "之清氣")
    dict.Add "豫", Array("章", "讓", "王", "瞻", "則立", "州", "暇", "知", "劇", "備", VBA.ChrW(20675), "防", "聞", "樟", _
        "先要", "豫為言之", "豫" & VBA.ChrW(29234) & "言之", "子")
    dict.Add "剝", Array("落", "削", "民", "蝕", "泐", "啄", "棗", "苔", "其皮", "去腸", "人面", "婦人衣", "而取之", "春" & VBA.ChrW(-31631), "春蔥", "芋")
    dict.Add VBA.ChrW(21093), Array("落", "削", "民", "蝕", "泐", "啄", "棗", "苔", "其皮", "去腸", "人面", "婦人衣", "而取之", "春" & VBA.ChrW(-31631), "春蔥", "芋")
    dict.Add "蹇", Array("諤", "驢", "叔", "氏", "周", "材望", "毅然", "已莫", "步", "吃")
    dict.Add "渙", Array("然", "散", "遂踰")
    dict.Add VBA.ChrW(28067), Array("然", "散", "遂踰")
    dict.Add "夬", Array("切")
    dict.Add "頤", Array("和園", "正叔", "字正", "茂叔", "字茂", "指氣使", "庵", "菴", "盦", "下有皮", "所以不")
    dict.Add VBA.ChrW(-26587), Array("和園", "正叔", "字正", "茂叔", "字茂", "指氣使", "庵", "菴", "盦", "下有皮", "所以不")
    dict.Add "萃", Array("於此", "此書", "于一", "古人", "之成", "其家", "江", "諸庫", "為一書", "於一門")
    dict.Add "艮", Array("岳", "嶽", "齋")
    dict.Add "賁", Array("赫", "隅之")
    dict.Add "巽", Array("懦", "字仲", "亦不較", "後仕", "風疏", "博學")
    dict.Add VBA.ChrW(14514), Array("懦", "字仲", "亦不較", "後仕", "風疏", "博學")
    dict.Add "蠱", Array("於心", "自埋", "之詐", "實生子", "毒", "發膨", "主", "之屬", "有鬼", "不絕", "者是也", "嫩", VBA.ChrW(23280))
    dict.Add "暌", Array("」作「睽", "作睽")
    dict.Add "坎", Array("鼓")
    dict.Add "遯", Array("跡", VBA.ChrW(-28679), "齋", "世修真")
    dict.Add "小畜", Array("集")
    dict.Add "同人", Array("醵錢")
    dict.Add "中孚", Array("禪子")
    dict.Add "大有", Array("功", "力", "父風", "警省", "逕庭", "李僧")
    dict.Add "大過", Array("人者")
    dict.Add "小過", Array("宜寬", "宜" & VBA.ChrW(23515))
    dict.Add "初九", Array("日")
    dict.Add "初六", Array("日")
    dict.Add "六四", Array("）")
    dict.Add "上六", Array("十里")
    dict.Add "少陰", Array("雨")
    dict.Add "咎", Array("繇", "陶", "彼", "單", "累")
    dict.Add "敦復", Array("學士")
    dict.Add "知臨", Array("江", "泉")
    dict.Add "直方", Array("殂之")
    dict.Add "晝日", Array("無事", "愈長")

    
    Set 易學KeywordsToMark_ExamFollowedAvoid = dict

End Property
Rem 某關鍵字不能在某個語句裡面 20240914
Property Get 易學KeywordsToMark_ExamInPhraseAvoid() As Scripting.Dictionary
    Dim dict As New Scripting.Dictionary
    dict.Add "易", Array("李易安", "居易錄", "蘇易簡", "杜易簡", "張易之", "此易事", "人易老", "人易去", "兩易其任", _
        "江山易主", "深耕易耨", "曾易占", "豈易得", "豈易逢", "竇易直", "後易為", "此易彼", "得而易失", "人而易私", "人而易" & VBA.ChrW(-10155) & VBA.ChrW(-8352), "以易此", "無以易也", "延易府", "皮易布", "金易一", "名易知", "雖易得", "誠至易", "則易束", "者易窺", "置易制", "難、易相", "事易行", "今易以冠服", "俱易以", "易知也", "市易法", "可以易一飽", "賈易不", "和易之氣", _
        "而易見", "而易起", "而易晦", "不能易也", "最易得", "亢易招", "以易鹽米", "頓易故", "谿易雨", "者易訓", "心易偏", "平心易氣", "故其說易差", "始易為力", "乃易合", "而易彼", "慢而易之", "而易治", "樂易之", "視而易之", _
        "以易心", "以易處之", "因易其韻", "以縉儒者易之", "因改易本文而", "尺寸易以", "則易發", "根易發", "客易位", "主易位", _
        "而易陵", "儉則易足", "者易犯", "狹則易足", "悔易勿輕踵", "智易窮", "何以易窮", "更有易見者", "而易散", _
        "至易事", "敢易也", "河東、易定", "易定、魏博", "河東易定", "脆易折", "易定魏博", "柔易治", "最易生", _
        "疾易作", "病易除", "人易從", "易放而難操", "錢差易 ", "言易墜", "所以易放", "者易直", "是易言也", "須易之", "之易以仰測", "蓴絲之易", "皆易黃屋", "皆易" & ChrW(-24892) & "屋", "時易以新", _
        "則易使", "謀易太子", "侮易承業", "成易具", "綿布易之", "布相易云", "浴易服", "惡易敗", _
        "因易名曰", "以字易名", "不易長", "和易近人", "大樂必易", "以之易業", "以之易用", "不欲易也", _
        "則易入於", "皆易與之", _
        "非易事", _
        "豈易說")
    dict.Add "卦", Array("八卦山", "八卦殿")
    dict.Add "乾", Array("大乾廟", "乳乾者")
    dict.Add "剝", Array("解剝而發明", "造剝洛陽")
    dict.Add VBA.ChrW(21093), Array("解" & VBA.ChrW(21093) & "而發明", "造" & VBA.ChrW(21093) & "洛陽")
    dict.Add "豫", Array("人豫知", "而豫求", "能豫逆", "豫射其")
    dict.Add "蹇", Array("剛蹇絕", "歲蹇剝")
    dict.Add "頤", Array("翁頤昌", "方頤大口", "解頤撫掌", "兩頤間")
    dict.Add VBA.ChrW(-26587), Array("翁" & VBA.ChrW(-26587) & "昌", "方" & VBA.ChrW(-26587) & "大口")
    dict.Add "暌", Array("有暌談笑")
    dict.Add "蠱", Array("以蠱留人", "以蠱而", "而蠱者", "以蠱大", "中蠱者")
    dict.Add "萃", Array("檀萃文")
    dict.Add "賁", Array("古賁灰")
    dict.Add "巽", Array("即巽也", "東巽泉", "邀巽二")
    dict.Add VBA.ChrW(14514), Array("即" & VBA.ChrW(14514) & "也", "東" & VBA.ChrW(14514) & "泉")
    dict.Add "同人", Array("底同人不", "不同人生", "不同人能")
    dict.Add "渙", Array("王渙之")
    dict.Add VBA.ChrW(28067), Array("王" & VBA.ChrW(28067) & "之")
    dict.Add "大有", Array("甚大有朱")
    dict.Add "大過", Array("無大過惡")
    dict.Add "既濟", Array("河既濟真")
    dict.Add VBA.ChrW(26083) & "濟", Array("河" & VBA.ChrW(26083) & "濟真")
    dict.Add "無妄", Array("人無妄取", "物無妄費")
    dict.Add VBA.ChrW(26080) & "妄", Array("人" & VBA.ChrW(26080) & "妄取", "物" & VBA.ChrW(26080) & "妄費")
    dict.Add VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522), Array("人" & VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522) & "取", "物" & VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522) & "費")
    dict.Add "初九", Array("虞初九百")
    dict.Add "九二", Array("一九二０")
    dict.Add "九三", Array("廿九三十")
    dict.Add "用九", Array("欲用九月")
    dict.Add "六二", Array("一六二四")
    dict.Add "上六", Array("以上六事", "已上六事")
    dict.Add "用六", Array("威用六極")
    dict.Add "存誠", Array("心存誠敬")
        
    Set 易學KeywordsToMark_ExamInPhraseAvoid = dict

End Property
Rem 檢測關鍵字
'Function 易學KeywordsToMark_Exam()
'
'End Function
