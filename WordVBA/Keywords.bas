Attribute VB_Name = "Keywords"
Option Explicit
Rem 任何關鍵字檢索、標識相關之屬性、參照記錄
Dim zhouyiguaShapeNameSequence As Scripting.Dictionary '這個筆數是固定的，所以可以如此
Dim zhouyiguaNameShapeSequence As Scripting.Dictionary '這個筆數是固定的，所以可以如此
Dim yiVariants As Scripting.Dictionary '周易異體字字典
Dim preceded_Avoid As Scripting.Dictionary '現在還在隨時新增中，故不宜寫死'現在為求效能，還是先寫，反正重啟Word就會更新，且需要更新時可以在即時運算視窗中輸入指令清除已有的內容 20241019
Dim followed_Avoid As Scripting.Dictionary '現在還在隨時新增中，故不宜寫死
Dim inPhrase_Avoid As Scripting.Dictionary '現在還在隨時新增中，故不宜寫死

Sub ClearDicts_YiKeywords()
    Set preceded_Avoid = Nothing
    Set followed_Avoid = Nothing
    Set inPhrase_Avoid = Nothing

End Sub
Rem 《易》學異體字對照/置換用
Property Get 易學異體字典() As Scripting.Dictionary
    If yiVariants Is Nothing Then
        Set 易學異體字典 = New Scripting.Dictionary
        易學異體字典.Add VBA.ChrW(20089), "乾"
        易學異體字典.Add VBA.ChrW(22531), "坤"
        易學異體字典.Add VBA.ChrW(-10132) & VBA.ChrW(-8313), "蒙"
        易學異體字典.Add VBA.ChrW(-10151) & VBA.ChrW(-9004), "需"
        易學異體字典.Add VBA.ChrW(-29764), "訟"
        易學異體字典.Add VBA.ChrW(24072), "師"
        易學異體字典.Add VBA.ChrW(-29658), "謙"
        易學異體字典.Add VBA.ChrW(-26993), "隨"
        易學異體字典.Add VBA.ChrW(20020), "臨"
        易學異體字典.Add VBA.ChrW(-30270), "觀"
        易學異體字典.Add VBA.ChrW(-29390), "賁"
        易學異體字典.Add VBA.ChrW(21093), "剝"
        易學異體字典.Add "复", "復"
        
        易學異體字典.Add "無妄", VBA.ChrW(26080) & "妄"
        易學異體字典.Add "無" & VBA.ChrW(-10171) & VBA.ChrW(-8522), VBA.ChrW(26080) & "妄"
        易學異體字典.Add VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522), VBA.ChrW(26080) & "妄"
        
        易學異體字典.Add VBA.ChrW(-26587), "頤"
        易學異體字典.Add "大" & VBA.ChrW(-28729), "大過"
        易學異體字典.Add "离", "離"
        易學異體字典.Add "恒", "恆"
        
        易學異體字典.Add VBA.ChrW(26187), "晉"
        易學異體字典.Add VBA.ChrW(-10164) & VBA.ChrW(-8698), "晉"
        
        易學異體字典.Add "明" & VBA.ChrW(-10171) & VBA.ChrW(-8739), "明夷"
        
        易學異體字典.Add "暌", "睽"
        
        易學異體字典.Add VBA.ChrW(-30233), "解"
        易學異體字典.Add VBA.ChrW(25439), "損"
        易學異體字典.Add VBA.ChrW(28176), "漸"
        易學異體字典.Add VBA.ChrW(24402) & "妹", "歸妹"
        易學異體字典.Add "丰", "豐"
        
        易學異體字典.Add VBA.ChrW(-26520), "巽"
        易學異體字典.Add VBA.ChrW(14514), "巽"
        
        易學異體字典.Add VBA.ChrW(20817), "兌"
        易學異體字典.Add VBA.ChrW(20810), "兌"
        
        易學異體字典.Add VBA.ChrW(28067), "渙"
        易學異體字典.Add VBA.ChrW(-32126), "節"
        易學異體字典.Add "小" & VBA.ChrW(-28729), "小過"
        
        易學異體字典.Add "既" & VBA.ChrW(27982), "既濟"
        易學異體字典.Add VBA.ChrW(26083) & "濟", "既濟"
        
        易學異體字典.Add "未" & VBA.ChrW(27982), "未濟"
        
        
        Set yiVariants = 易學異體字典
    Else
        Set 易學異體字典 = yiVariants
    End If
End Property
Rem key ,string()
Property Get 周易卦形_卦名_卦序() As Scripting.Dictionary
    If zhouyiguaShapeNameSequence Is Nothing Then
        Set 周易卦形_卦名_卦序 = New Scripting.Dictionary
        周易卦形_卦名_卦序.Add VBA.ChrW(19904), Array("乾", 1)
        周易卦形_卦名_卦序.Add VBA.ChrW(19905), Array("坤", 2)
        周易卦形_卦名_卦序.Add VBA.ChrW(19906), Array("屯", 3)
        周易卦形_卦名_卦序.Add VBA.ChrW(19907), Array("蒙", 4)
        周易卦形_卦名_卦序.Add VBA.ChrW(19908), Array("需", 5)
        周易卦形_卦名_卦序.Add VBA.ChrW(19909), Array("訟", 6)
        周易卦形_卦名_卦序.Add VBA.ChrW(19910), Array("師", 7)
        周易卦形_卦名_卦序.Add VBA.ChrW(19911), Array("比", 8)
        周易卦形_卦名_卦序.Add VBA.ChrW(19912), Array("小畜", 9)
        周易卦形_卦名_卦序.Add VBA.ChrW(19913), Array("履", 10)
        周易卦形_卦名_卦序.Add VBA.ChrW(19914), Array("泰", 11)
        周易卦形_卦名_卦序.Add VBA.ChrW(19915), Array("否", 12)
        周易卦形_卦名_卦序.Add VBA.ChrW(19916), Array("同人", 13)
        周易卦形_卦名_卦序.Add VBA.ChrW(19917), Array("大有", 14)
        周易卦形_卦名_卦序.Add VBA.ChrW(19918), Array("謙", 15)
        周易卦形_卦名_卦序.Add VBA.ChrW(19919), Array("豫", 16)
        周易卦形_卦名_卦序.Add VBA.ChrW(19920), Array("隨", 17)
        周易卦形_卦名_卦序.Add VBA.ChrW(19921), Array("蠱", 18)
        周易卦形_卦名_卦序.Add VBA.ChrW(19922), Array("臨", 19)
        周易卦形_卦名_卦序.Add VBA.ChrW(19923), Array("觀", 20)
        周易卦形_卦名_卦序.Add VBA.ChrW(19924), Array("噬嗑", 21)
        周易卦形_卦名_卦序.Add VBA.ChrW(19925), Array("賁", 22)
        周易卦形_卦名_卦序.Add VBA.ChrW(19926), Array("剝", 23)
        周易卦形_卦名_卦序.Add VBA.ChrW(19927), Array("復", 24)
        周易卦形_卦名_卦序.Add VBA.ChrW(19928), Array(VBA.ChrW(26080) & "妄", 25)
        周易卦形_卦名_卦序.Add VBA.ChrW(19929), Array("大畜", 26)
        周易卦形_卦名_卦序.Add VBA.ChrW(19930), Array("頤", 27)
        周易卦形_卦名_卦序.Add VBA.ChrW(19931), Array("大過", 28)
        周易卦形_卦名_卦序.Add VBA.ChrW(19932), Array("坎", 29)
        周易卦形_卦名_卦序.Add VBA.ChrW(19933), Array("離", 30)
        周易卦形_卦名_卦序.Add VBA.ChrW(19934), Array("咸", 31)
        周易卦形_卦名_卦序.Add VBA.ChrW(19935), Array("恆", 32)
        周易卦形_卦名_卦序.Add VBA.ChrW(19936), Array("遯", 33)
        周易卦形_卦名_卦序.Add VBA.ChrW(19937), Array("大壯", 34)
        周易卦形_卦名_卦序.Add VBA.ChrW(19938), Array("晉", 35)
        周易卦形_卦名_卦序.Add VBA.ChrW(19939), Array("明夷", 36)
        周易卦形_卦名_卦序.Add VBA.ChrW(19940), Array("家人", 37)
        周易卦形_卦名_卦序.Add VBA.ChrW(19941), Array("睽", 38)
        周易卦形_卦名_卦序.Add VBA.ChrW(19942), Array("蹇", 39)
        周易卦形_卦名_卦序.Add VBA.ChrW(19943), Array("解", 40)
        周易卦形_卦名_卦序.Add VBA.ChrW(19944), Array("損", 41)
        周易卦形_卦名_卦序.Add VBA.ChrW(19945), Array("益", 42)
        周易卦形_卦名_卦序.Add VBA.ChrW(19946), Array("夬", 43)
        周易卦形_卦名_卦序.Add VBA.ChrW(19947), Array("姤", 44)
        周易卦形_卦名_卦序.Add VBA.ChrW(19948), Array("萃", 45)
        周易卦形_卦名_卦序.Add VBA.ChrW(19949), Array("升", 46)
        周易卦形_卦名_卦序.Add VBA.ChrW(19950), Array("困", 47)
        周易卦形_卦名_卦序.Add VBA.ChrW(19951), Array("井", 48)
        周易卦形_卦名_卦序.Add VBA.ChrW(19952), Array("革", 49)
        周易卦形_卦名_卦序.Add VBA.ChrW(19953), Array("鼎", 50)
        周易卦形_卦名_卦序.Add VBA.ChrW(19954), Array("震", 51)
        周易卦形_卦名_卦序.Add VBA.ChrW(19955), Array("艮", 52)
        周易卦形_卦名_卦序.Add VBA.ChrW(19956), Array("漸", 53)
        周易卦形_卦名_卦序.Add VBA.ChrW(19957), Array("歸妹", 54)
        周易卦形_卦名_卦序.Add VBA.ChrW(19958), Array("豐", 55)
        周易卦形_卦名_卦序.Add VBA.ChrW(19959), Array("旅", 56)
        周易卦形_卦名_卦序.Add VBA.ChrW(19960), Array("巽", 57)
        周易卦形_卦名_卦序.Add VBA.ChrW(19961), Array("兌", 58)
        周易卦形_卦名_卦序.Add VBA.ChrW(19962), Array("渙", 59)
        周易卦形_卦名_卦序.Add VBA.ChrW(19963), Array("節", 60)
        周易卦形_卦名_卦序.Add VBA.ChrW(19964), Array("中孚", 61)
        周易卦形_卦名_卦序.Add VBA.ChrW(19965), Array("小過", 62)
        周易卦形_卦名_卦序.Add VBA.ChrW(19966), Array("既濟", 63)
        周易卦形_卦名_卦序.Add VBA.ChrW(19967), Array("未濟", 64)
        Set zhouyiguaShapeNameSequence = 周易卦形_卦名_卦序
    Else
        Set 周易卦形_卦名_卦序 = zhouyiguaShapeNameSequence
    End If
End Property
Rem key ,string()
Property Get 周易卦名_卦形_卦序() As Scripting.Dictionary
    On Error GoTo eH:
    If zhouyiguaNameShapeSequence Is Nothing Then
        Set 周易卦名_卦形_卦序 = New Scripting.Dictionary
        周易卦名_卦形_卦序.Add "乾", Array(VBA.ChrW(19904), 1)
        周易卦名_卦形_卦序.Add "坤", Array(VBA.ChrW(19905), 2)
        周易卦名_卦形_卦序.Add "屯", Array(VBA.ChrW(19906), 3)
        周易卦名_卦形_卦序.Add "蒙", Array(VBA.ChrW(19907), 4)
        周易卦名_卦形_卦序.Add "需", Array(VBA.ChrW(19908), 5)
        周易卦名_卦形_卦序.Add "訟", Array(VBA.ChrW(19909), 6)
        周易卦名_卦形_卦序.Add "師", Array(VBA.ChrW(19910), 7)
        周易卦名_卦形_卦序.Add "比", Array(VBA.ChrW(19911), 8)
        周易卦名_卦形_卦序.Add "小畜", Array(VBA.ChrW(19912), 9)
        周易卦名_卦形_卦序.Add "履", Array(VBA.ChrW(19913), 10)
        周易卦名_卦形_卦序.Add "泰", Array(VBA.ChrW(19914), 11)
        周易卦名_卦形_卦序.Add "否", Array(VBA.ChrW(19915), 12)
        周易卦名_卦形_卦序.Add "同人", Array(VBA.ChrW(19916), 13)
        周易卦名_卦形_卦序.Add "大有", Array(VBA.ChrW(19917), 14)
        周易卦名_卦形_卦序.Add "謙", Array(VBA.ChrW(19918), 15)
        周易卦名_卦形_卦序.Add "豫", Array(VBA.ChrW(19919), 16)
        周易卦名_卦形_卦序.Add "隨", Array(VBA.ChrW(19920), 17)
        周易卦名_卦形_卦序.Add "蠱", Array(VBA.ChrW(19921), 18)
        周易卦名_卦形_卦序.Add "臨", Array(VBA.ChrW(19922), 19)
        周易卦名_卦形_卦序.Add "觀", Array(VBA.ChrW(19923), 20)
        周易卦名_卦形_卦序.Add "噬嗑", Array(VBA.ChrW(19924), 21)
        周易卦名_卦形_卦序.Add "賁", Array(VBA.ChrW(19925), 22)
        周易卦名_卦形_卦序.Add "剝", Array(VBA.ChrW(19926), 23)
        周易卦名_卦形_卦序.Add "復", Array(VBA.ChrW(19927), 24)
        周易卦名_卦形_卦序.Add VBA.ChrW(26080) & "妄", Array(VBA.ChrW(19928), 25)
        周易卦名_卦形_卦序.Add "大畜", Array(VBA.ChrW(19929), 26)
        周易卦名_卦形_卦序.Add "頤", Array(VBA.ChrW(19930), 27)
        周易卦名_卦形_卦序.Add "大過", Array(VBA.ChrW(19931), 28)
        周易卦名_卦形_卦序.Add "坎", Array(VBA.ChrW(19932), 29)
        周易卦名_卦形_卦序.Add "離", Array(VBA.ChrW(19933), 30)
        周易卦名_卦形_卦序.Add "咸", Array(VBA.ChrW(19934), 31)
        周易卦名_卦形_卦序.Add "恆", Array(VBA.ChrW(19935), 32)
        周易卦名_卦形_卦序.Add "遯", Array(VBA.ChrW(19936), 33)
        周易卦名_卦形_卦序.Add "大壯", Array(VBA.ChrW(19937), 34)
        周易卦名_卦形_卦序.Add "晉", Array(VBA.ChrW(19938), 35)
        周易卦名_卦形_卦序.Add "明夷", Array(VBA.ChrW(19939), 36)
        周易卦名_卦形_卦序.Add "家人", Array(VBA.ChrW(19940), 37)
        周易卦名_卦形_卦序.Add "睽", Array(VBA.ChrW(19941), 38)
        周易卦名_卦形_卦序.Add "蹇", Array(VBA.ChrW(19942), 39)
        周易卦名_卦形_卦序.Add "解", Array(VBA.ChrW(19943), 40)
        周易卦名_卦形_卦序.Add "損", Array(VBA.ChrW(19944), 41)
        周易卦名_卦形_卦序.Add "益", Array(VBA.ChrW(19945), 42)
        周易卦名_卦形_卦序.Add "夬", Array(VBA.ChrW(19946), 43)
        周易卦名_卦形_卦序.Add "姤", Array(VBA.ChrW(19947), 44)
        周易卦名_卦形_卦序.Add "萃", Array(VBA.ChrW(19948), 45)
        周易卦名_卦形_卦序.Add "升", Array(VBA.ChrW(19949), 46)
        周易卦名_卦形_卦序.Add "困", Array(VBA.ChrW(19950), 47)
        周易卦名_卦形_卦序.Add "井", Array(VBA.ChrW(19951), 48)
        周易卦名_卦形_卦序.Add "革", Array(VBA.ChrW(19952), 49)
        周易卦名_卦形_卦序.Add "鼎", Array(VBA.ChrW(19953), 50)
        周易卦名_卦形_卦序.Add "震", Array(VBA.ChrW(19954), 51)
        周易卦名_卦形_卦序.Add "艮", Array(VBA.ChrW(19955), 52)
        周易卦名_卦形_卦序.Add "漸", Array(VBA.ChrW(19956), 53)
        周易卦名_卦形_卦序.Add "歸妹", Array(VBA.ChrW(19957), 54)
        周易卦名_卦形_卦序.Add "豐", Array(VBA.ChrW(19958), 55)
        周易卦名_卦形_卦序.Add "旅", Array(VBA.ChrW(19959), 56)
        周易卦名_卦形_卦序.Add "巽", Array(VBA.ChrW(19960), 57)
        周易卦名_卦形_卦序.Add "兌", Array(VBA.ChrW(19961), 58)
        周易卦名_卦形_卦序.Add "渙", Array(VBA.ChrW(19962), 59)
        周易卦名_卦形_卦序.Add "節", Array(VBA.ChrW(19963), 60)
        周易卦名_卦形_卦序.Add "中孚", Array(VBA.ChrW(19964), 61)
        周易卦名_卦形_卦序.Add "小過", Array(VBA.ChrW(19965), 62)
        周易卦名_卦形_卦序.Add "既濟", Array(VBA.ChrW(19966), 63)
        周易卦名_卦形_卦序.Add "未濟", Array(VBA.ChrW(19967), 64)
        Set zhouyiguaNameShapeSequence = 周易卦名_卦形_卦序
    Else
        Set 周易卦名_卦形_卦序 = zhouyiguaNameShapeSequence
    End If
    Exit Property
eH:
    Select Case Err.Number
        Case Else
            Debug.Print Err.Number & Err.Description
            Stop
    End Select
End Property
Rem 用以檢查是否為易學範圍之內容用
Property Get 易學Keywords_ToCheck() As Variant 'string()
    易學Keywords_ToCheck = Array(VBA.ChrW(-10119), VBA.ChrW(-8742), VBA.ChrW(-30233), VBA.ChrW(-10164), VBA.ChrW(-8698), VBA.ChrW(-31827), VBA.ChrW(-10132), VBA.ChrW(-8313), VBA.ChrW(20810), VBA.ChrW(-10167), VBA.ChrW(-8698), VBA.ChrW(-26587), VBA.ChrW(21093), VBA.ChrW(14615), VBA.ChrW(20089), VBA.ChrW(26080), "妄", VBA.ChrW(26083), "濟" _
        , "遘", "遁", VBA.ChrW(20089), "离", "乾", "小畜", "履", "臨", "觀", "大過", "坤", "泰", "否", "噬嗑", "賁", "坎", "屯", "蒙", "同人", "大有", "剝", "復", "離", "需", "訟", "謙", "豫", "無妄", "大畜", "師", "比", "隨", "蠱", "頤", "咸", "恒", "損", "益", "震", "艮", "中孚", "遯", "大壯", "夬", "姤", "漸", "歸妹", "小過", "晉", "明夷", "萃", "升", "豐", "旅", "既濟", "未濟", "家人", "睽", "困", "井", "巽", "兌", "蹇", "解", "革", "鼎", "渙", "節", "太極", "陰陽", "兩儀", "象", VBA.ChrW(-10145) & VBA.ChrW(-9156), "彖", _
        "老陰", "老陽", "少陰", "少陽", "蓍")
End Property
Rem 用以標識易學關鍵字用
Property Get 易學Keywords_ToMark() As Variant 'string()因為 Array Returns a Variant containing an array,所以不能寫成 as string()
    易學Keywords_ToMark = Array("易", "周易", "易經", "大易", "五經", "六經", "七經", "十三經", "蓍", _
        "卦", "節卦", "離卦", "臨卦", "屯蒙", "屯" & VBA.ChrW(-10132) & VBA.ChrW(-8313), "屯、蒙", "屯、" & VBA.ChrW(-10132) & VBA.ChrW(-8313), "乾卦", "坤卦", "訟卦", "師卦", "比卦", "履卦", "泰卦", "否卦", "謙卦", "隨卦", "觀卦", "復卦", "習坎", "離卦", "咸卦", "恆卦", "晉卦", "家人卦", "解卦", "損卦", "益卦", "升卦", "困卦", "井卦", "革卦", "鼎卦", "震卦", "漸卦", "豐卦", "旅卦", "節卦", "爻", "系辭", "繫辭", "擊辭", "擊詞", "繫詞", "說卦", "序卦", "卦序", "敘卦", "雜卦", "文言", "乾坤", "元亨", "利貞", "史記", "筮", "夬", "乾知大始", "坤作成物", "乾以易知", "坤以簡能", "乾", "〈乾〉", "〈坤〉", "乾、坤", "〈乾、坤〉", "噬嗑", "賁于外", "賁於外", "外賁", "內賁", "賁", VBA.ChrW(20089), "既濟", VBA.ChrW(26083) & "濟", "未濟", "十翼", "大" & VBA.ChrW(22766), _
        "初九", "九二", "九三", "九四", "九五", "上九", VBA.ChrW(19972) & "九", "用九", "初六", "六二", "六三", "六四", "六五", "上六", "用六", "河圖", "洛書", "太極", "無極", "兩儀", _
            "象曰", "〈象〉曰", "象日", "象云", "象傳", VBA.ChrW(-10145) & VBA.ChrW(-9156) & "傳", "大象", "大" & VBA.ChrW(-10145) & VBA.ChrW(-9156), "小象", "象義", "四象", "象：", "象文", "彖", _
             "艮其背", "艮", "頤", "同人于宗", "同人", "坎", "中孚", "兌", VBA.ChrW(20817), VBA.ChrW(20810), "蠱", "姤", "巽", VBA.ChrW(14514), VBA.ChrW(-26520), "剝", VBA.ChrW(21093), "遯世無悶", "遯世" & ChrW(26080) & "悶", "遯", "大壯", "明夷", "明" & VBA.ChrW(-10171) & VBA.ChrW(-8739), "小畜", "大畜", "萃", "蹇", "渙", VBA.ChrW(28067), "睽", "暌", "歸妹", "小過", "大有", "大過", "〈泰〉", "〈否〉", "〈損〉", "〈益〉", "〈屯", "蒙〉", VBA.ChrW(-10132) & VBA.ChrW(-8313) & "〉", "豫大", "豫", "〈旡妄〉", "〈復〉", "〈震〉", "〈需〉", "老陰", "老" & VBA.ChrW(-27006), "老陽", "少陰", "少" & VBA.ChrW(-27006), "少陽", "繇辭", "繇詞", _
            "咎", "往遴", VBA.ChrW(24451) & "遴", "往吝", VBA.ChrW(24451) & "吝", "咸恒", "咸恆", "大衍", "象也者", VBA.ChrW(-10145) & VBA.ChrW(-9156) & "也者", "咸之九五", "禴祭", "東鄰", "東" & VBA.ChrW(-26973), "禴" & VBA.ChrW(-10155) & VBA.ChrW(-8630), "禴" & VBA.ChrW(-10131) & VBA.ChrW(-8268), "善不積", "天下雷行", "號咷", "見龍在田", "菑畬", VBA.ChrW(-31656) & "畬", _
        "不家食", "無妄", VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522), VBA.ChrW(26080) & "妄", "無咎", VBA.ChrW(26080) & "咎", "天咎", "允升", "屯其膏", "豐亨", "終難", "天行健", "天行" & VBA.ChrW(24484), "知幾", "知" & VBA.ChrW(-10123) & VBA.ChrW(-8628), "知" & VBA.ChrW(14444), "有子考", "有子攷", "不易乎世", "不易乎" & VBA.ChrW(21323), "不易乎" & VBA.ChrW(19991), "不成乎名", "天一地二", "蒞眾", VBA.ChrW(-31867) & "眾", VBA.ChrW(-31867) & VBA.ChrW(-30650), "蒞" & VBA.ChrW(-30650), "幹父", "裕父", "係遯", "甘臨", "翰音", "鶾音", _
        "隨時之義", "庖有魚", "包有魚", "精義入神", "通乎晝夜", _
        "豶豕", "童牛", "承之羞", "雷在天上", "錫馬", "蕃庶", "晝日", "三接", "懲忿", "窒欲", "窒慾", "懲窒", "敬以" & VBA.ChrW(-10114) & VBA.ChrW(-8896) & "內", "敬以直" & VBA.ChrW(20869), "敬以直內", "義以方外", "迷後得主", "利西南", "品物咸章", "天下大行", "益動而", "日進無疆", "日進" & VBA.ChrW(26080) & "疆", "頻巽", "頻" & VBA.ChrW(14514), "豚魚", "頻復", "閑邪", "存誠", "乾乾", "悔吝", "憧憧", "類萬物", VBA.ChrW(-10139) & VBA.ChrW(-8938) & "萬物", VBA.ChrW(-10139) & VBA.ChrW(-8937) & "萬物", "柔順利貞", VBA.ChrW(-10163) & VBA.ChrW(-9167) & "順利貞", "比之匪人", "履貞", "貞厲", "履道坦坦", "貞吉", "貞凶", "直方", "木上有水", "不事王侯", "不事王" & VBA.ChrW(30694), "高尚其事", "高" & VBA.ChrW(23577) & "其事", VBA.ChrW(-25895) & VBA.ChrW(23577) & "其事", VBA.ChrW(-25895) & "尚其事", "勞民勸相", "索而得", "索而" & VBA.ChrW(-10167) & VBA.ChrW(-8906), "貞不字", "立成器", "與地之", _
        "括囊", "無悶", VBA.ChrW(26080) & "悶", "悔亡", "悔" & VBA.ChrW(20158), "悔" & VBA.ChrW(20838), "時義", "健順", "內健而外順", "內健外順", "外順而內健", "外順內健", "易簡", "易" & VBA.ChrW(-10153) & VBA.ChrW(-9007), "敦復", "開物成務", "窮神知化", "研幾極深", "極深研幾", "研幾", "見善則遷", "有過則改", "遷善改過", "夕惕", "惕若", "一" & VBA.ChrW(-27006) & "一陽", "一陰一陽", "我有好爵", "言有序", "有聖人之道四", "長子帥師", "弟子輿尸", "日用而不知", "日用不知", "之道鮮", "原始反終", "然不動", "感而遂通", "朋從", "朋盍", "容民畜眾", "容民畜" & VBA.ChrW(-30650), "養正", "養賢", "知臨", "臨大君", "變化云", VBA.ChrW(-10164) & VBA.ChrW(-9163) & "化云", "神道設教", "神道設" & VBA.ChrW(25934), "默而成之", VBA.ChrW(-24871) & "而成之", "不言而信", "存乎德行", "通天下之志", "履正", "繼之者善", "仁者見之", "知者見之", "智者見之", "撝謙", VBA.ChrW(-10114) & VBA.ChrW(-9019) & "謙", "理財", "正辭", "禁民為非", _
        "大貞", "小貞", "帝出乎震", "帝出於震", "帝出于震", "與時偕行", "盈虛", "盈" & VBA.ChrW(-31142), "盈" & VBA.ChrW(-10119) & VBA.ChrW(-8991), "盈" & VBA.ChrW(-31145), "履霜", "艮其限", "乃孚", "浚恒", "浚恆", "包蒙", "童蒙", "蒙吉", "包" & VBA.ChrW(-10132) & VBA.ChrW(-8313), "童" & VBA.ChrW(-10132) & VBA.ChrW(-8313), VBA.ChrW(-10132) & VBA.ChrW(-8313) & "吉", "確乎其不可拔", "碻乎其不可拔", _
        "天在山中", "多識前言" & VBA.ChrW(24451) & "行", "多識前言往行", "蹇蹇", "匪躬", "匪" & VBA.ChrW(-29005), "噬膚", "山澤通氣", "其腓", "洗心", "龍德", "慎言語", VBA.ChrW(24892) & "言語", "節飲食", "改命吉", VBA.ChrW(25914) & "命吉", "開國承家", "舊井", VBA.ChrW(26087) & "井", VBA.ChrW(-10149) & VBA.ChrW(-8300) & "井", "井谷", _
        "離麗", "离麗", "為麗", VBA.ChrW(29234) & "麗", "賁於丘園", "賁于丘園", "賁於邱園", "賁于邱園", "賁於" & VBA.ChrW(-10176) & VBA.ChrW(-9207) & "園", "賁于" & VBA.ChrW(-10176) & VBA.ChrW(-9207) & "園", "先天後天", "風行水上", "喪貝", VBA.ChrW(17966) & "貝", VBA.ChrW(-10172) & VBA.ChrW(-9052) & "貝", VBA.ChrW(-10124) & VBA.ChrW(-8660) & "貝", VBA.ChrW(-10173) & VBA.ChrW(-8748) & "貝", VBA.ChrW(20007) & "貝", VBA.ChrW(-10173) & VBA.ChrW(-8650) & "貝", "羝羊", "羝" & VBA.ChrW(-32119), "觸藩", "觸籓", "立其誠", "立誠", "修辭立誠", "脩辭立誠", _
        "賁於" & VBA.ChrW(-10176) & VBA.ChrW(-9204) & "園", "賁于" & VBA.ChrW(-10176) & VBA.ChrW(-9204) & "園", "賁於" & VBA.ChrW(-10143) & VBA.ChrW(-8559) & "園", "賁于" & VBA.ChrW(-10143) & VBA.ChrW(-8559) & "園", "鞏用", "艱貞", "金矢", "利有", "攸往", "中正", _
        "束帛", "戔戔", "損下以益上", "損下益上", "損下而益上", "貳用缶", "納約自牖", "利見大人", "何思何慮", "同歸而殊塗", "一致而百慮", "同歸殊塗", "一致百慮", "精氣為物", "者其辭", "事不密", "事不" & VBA.ChrW(23483), "喪貝", VBA.ChrW(-10173) & VBA.ChrW(-8748) & "貝", "日新", _
        "游魂為變", "遊" & VBA.ChrW(19487) & "為變", "游" & VBA.ChrW(19487) & "為變", "漣如", "焚如", "朋亡", "渙其群", VBA.ChrW(28067) & "其群", "渙其" & VBA.ChrW(32675), VBA.ChrW(28067) & "其" & VBA.ChrW(32675), "甲三日", "庚三日", "升其高陵", "升其" & VBA.ChrW(-25895) & "陵", "天道虧盈", "祗悔", "祇悔", "秖悔", "秪悔", _
        "伏羲", "伏" & VBA.ChrW(-10152) & VBA.ChrW(-8255), "庖" & VBA.ChrW(-10152) & VBA.ChrW(-8255), "庖羲", "宓羲", "宓" & VBA.ChrW(-10152) & VBA.ChrW(-8255), "宓犧", "伏犧", "庖犧")
        
End Property
Rem 某關鍵字的前面不能是 20240914
Property Get 易學KeywordsToMark_Exam_Preceded_Avoid() As Scripting.Dictionary
    If preceded_Avoid Is Nothing Then
        
        Dim dict As New Scripting.Dictionary, cln As New VBA.Collection
        ' 添加資料到字典 creedit_with_Copilot大菩薩：https://sl.bing.net/goDF239cQVw
        dict.Add "易", Array("尤", "容", "未", "不", "極", "甚", "貿", "交", "物", "變", "或可", "鄙", "博", "辟", "平", "慢", "俗", "坦", "難", "脫", "流", "樂", "革", "更", "簡", "白居", "居", "淺", "輕", "險", "相", "難行", "世", "易", "所", _
            "移", "狂", "平仄相", "遂", "每歲一", "錢", "涿", "傳寫之便", "光庭、賈", "光庭賈", "為人和", "辟", "驕", "一氣一", "資和", _
            "新陳相", "捷最", "上下互", "崔伯", "劉", "覆墜之", "知其難", "有不善未嘗不知", "事難慮", "事久則慮", "勢固", "立門戶也", "立門" & VBA.ChrW(25143) & "也", _
            "聽之者", "壞真從", "尋又", "愉和", "壞" & VBA.ChrW(30494) & "從", "厚和", "誠不為", "可", "人欲", "市", "輒", VBA.ChrW(-28903), "遽", "過於和", "平心", "大樂必", _
            "知其至", "立節行、", "立節行", "圖難於其", "成誦", "月日時無", "使聽者之", "嶽號五", "鮮不忽", "而外和", "致終身之", "以天下與人", "詰王與")
            
        dict.Add "乾", Array("白", "豆", "衣", "面自", "擰", "餅", "未", "晾", "肉", "蘿蔔", "葡萄", "龍眼", "口", "枯", "烘", "晒", "曬", "筍", "外強中")
        dict.Add "乾坤", Array("搆盡", "于此盜", "五季")
        dict.Add "豫", Array("防患於", "暇", "厎", "音", "底", "悅", "不", "弗", "劉", "猶", "逸", VBA.ChrW(-10143) & VBA.ChrW(-8996), "道", "南", "美熙", "戲", VBA.ChrW(25135))
        dict.Add "剝", Array("刻", "活", "可", "褫", "皴", "為之解", "歲蹇", "石斷")
        dict.Add VBA.ChrW(21093), Array("刻", "活", "可", "褫", "為之解", "歲蹇", "石斷")
        dict.Add "頤", Array("周敦", "程", "朵", "濬", "期", "面豐", "頂至", "泗交", "張", "楚", _
                        "獨支", "筆支", "手支", _
                        "寄藥與")
        dict.Add VBA.ChrW(-26587), Array("周敦", "程", "朵", "濬", "期", "面豐", "頂至", "泗交", "張", "楚", _
                        "獨支", "筆支", "手支", _
                        "寄藥與")
                        
        dict.Add "巽", Array("李", "翟公", "傅", "家之", "叔", "劉", "朱", "陳問")
        dict.Add VBA.ChrW(14514), Array("李", "翟公", "傅", "家之", "叔", "劉", "朱", "陳問")
        
        dict.Add "兌", Array("李")
        dict.Add VBA.ChrW(20817), Array("李")
        dict.Add VBA.ChrW(20810), Array("李")
        
        dict.Add "大過", Array("可過", "公頗無")
        dict.Add "賁", Array("虎", "孟", "孫", "諸葛", "齎」作「", "齎作", "東海襄")
        dict.Add "蹇", Array("偃", "矯", "策", "奇", "長之飛")
        dict.Add "夬", Array("龔")
        dict.Add "中孚", Array("周", "僧")
        dict.Add "小過", Array("吏有")
        
        dict.Add "渙", Array("崔", "蘇", "程", "畔", "本作", "士", "黃", VBA.ChrW(-24892), "謁", "濉", "王", "滑")
        dict.Add VBA.ChrW(28067), Array("崔", "蘇", "程", "畔", "本作", "士", "黃", VBA.ChrW(-24892), "謁", "濉", "王", "滑")
        
        dict.Add "蠱", Array("巫", "韓", "置", "蟲", "可以解", "蛇", "年之", "下", "之立", "謂水", "每遇", "土", "妖", "音")
        dict.Add "萃", Array("拔", "蓊", "拔乎其", "悉", "雲氣", "森")
        dict.Add "睽", Array("暌」作「", "暌作")
        dict.Add "暌", Array("以分")
        dict.Add "遯", Array("毅然知肥", "子宜速")
        dict.Add "同人", Array("招", "儲", "山西大")
        dict.Add "大有", Array("歲稱", "後弟", "來曰", "花朵", "葉")
        dict.Add "噬嗑", Array("令")
        dict.Add "既濟", Array("沈", "沉")
        dict.Add "初九", Array("月")
        dict.Add "九二", Array("卷四", "卷", "卷二", "一一", "一五")
        dict.Add "九三", Array("卷", "卷一", "一五")
        dict.Add "九四", Array("卷", "卷一", "一五", "張")
        dict.Add "九五", Array("卷", "卷一", "一五", "一六")
        dict.Add "初六", Array("月")
        dict.Add "六二", Array("卷", "卷一", "卷二")
        dict.Add "六三", Array("卷")
        dict.Add "六四", Array("卷", "卷一", "…　一", "（九")
        dict.Add "六五", Array("卷", "卷一", "…　一", "（九")
        dict.Add "彖", Array("張")
        dict.Add "象云", Array("郭", "皇", "光景氣")
        dict.Add "文言", Array("與上", "且上")
        dict.Add "筮", Array("初", "再")
        dict.Add "咎", Array("得", "歸", "不", "過", "追", "休", "厥", "自", "引", "殃", "何", "晁無", "晁丈無", "足", "受其", "重其", "任其", "論者皆", _
            "專", "示", "將有", "怨", "屈宜", "適逢其", "乖忤之", "庸得免") ', "知休", "卜休", "知人休", "能知休", "身之休"
        dict.Add "無妄", Array("誠而")
        dict.Add VBA.ChrW(26080) & "妄", Array("誠而")
        dict.Add VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522), Array("誠而")
        dict.Add "元亨", Array("董", "萬")
        dict.Add "貞吉", Array("曹")
        dict.Add "悔吝", Array("平生")
        dict.Add "無咎", Array("晁", "晁丈")
        dict.Add "直方", Array("張", "王", "字")
        dict.Add "敦復", Array("張", "字")
        dict.Add "允升", Array("賈", "薛", "字")
        
        dict.Add "盈虛", Array("邊塞之", "監世")
        dict.Add "盈" & VBA.ChrW(-31142), Array("邊塞之", "監世")
        dict.Add "盈" & VBA.ChrW(-10119) & VBA.ChrW(-8991), Array("邊塞之", "監世")
        dict.Add "盈" & VBA.ChrW(-31145), Array("邊塞之", "監世")
        
        dict.Add "存誠", Array("游操")
        dict.Add "無極", Array("與天", "作配", "，永")
        dict.Add "易簡", Array("蘇")
        dict.Add "易" & VBA.ChrW(-10153) & VBA.ChrW(-9007), Array("蘇")
        dict.Add "索而得", Array("豈窮")
        dict.Add "索而" & VBA.ChrW(-10167) & VBA.ChrW(-8906), Array("豈窮")
        dict.Add "養正", Array("劉")
        dict.Add "號咷", Array(VBA.ChrW(32675) & "臣")
        
        Set preceded_Avoid = dict
        Set 易學KeywordsToMark_Exam_Preceded_Avoid = dict
    Else
        Set 易學KeywordsToMark_Exam_Preceded_Avoid = preceded_Avoid
    End If
        
End Property
Rem 某關鍵字的後面不能是 20240914
Property Get 易學KeywordsToMark_Exam_Followed_Avoid() As Scripting.Dictionary
    If followed_Avoid Is Nothing Then
        Dim dict As New Scripting.Dictionary
        dict.Add "易", Array("簀", "牙", "之以書契", "安居士", "幟", "轍", "厭", "水", "州", "順鼎", "破", "開罐", "筋", "姓", "名之日", "卜生", "堂", "科", _
            "守", "棺", "看", "積", "容", "衣", "萌", "葬", "獄", "其處", "名者", "曉", "得消散", "慢", "熔", "忘", "物", "易", "俗", "差", "陷", "肆", "下手", "渡", "紙易之", "腐敗", _
             "手", "如反", "善作詩", "學而", "著火", "於噴", "元善", "其田疇", "世之後", "以走險", "不能堪", "直不足言", "以動人", "得鹵", "十姓", "易其形者為", "言則難", "視之", _
            "搖而難", "動難安", "昏而難", "字希", "見者", "代而改", "於混亂", VBA.ChrW(24422) & "章", "彥章", "以挺", "於近者。非知言者也", "於近者非知言者也", "出議論", "退之節", "玄光", "梓宮", "知由單", "占偶書", "世而後", "於怠惰", _
            "如翻", "如拾", "子而", "肩輿", "地皆安", "為求福", "為肉黍", "為黍肉", "事爾", "但不香", "知而不知", "舊榜", "得汩沒", "子所作", "易生粉", "發酸", "如燎毛", "置其次", "至沮喪", "驚也")
        
        dict.Add "周易", Array("癡")
        dict.Add "卦", Array("陣", "橋", "建築")
        dict.Add "筮", Array("仕")
        dict.Add "乾", Array("淨", "隆", "寧", "祐", _
            "道初", "道元年", "道二年", "道三年", "道四年", "道五年", "道七月", _
            "和中", "枯", "坤清泰", "坤之清氣", "鵲", "闕一" & VBA.ChrW(-28146) & "金")
        dict.Add "乾坤", Array("陷吉人", "清泰", "之清氣")
        dict.Add "豫", Array("章", "讓", "王", "瞻", "則立", "州", "暇", "知", "劇", "備", VBA.ChrW(20675), "防", "聞", "樟", "憂思", _
            "先要", "豫為言之", "豫" & VBA.ChrW(29234) & "言之", "子", "卜地")
        dict.Add "剝", Array("落", "削", "民", "蝕", "泐", "啄", "棗", "苔", "其皮", "去腸", "人面", "婦人衣", "而取之", "春" & VBA.ChrW(-31631), "春蔥", "芋")
        dict.Add VBA.ChrW(21093), Array("落", "削", "民", "蝕", "泐", "啄", "棗", "苔", "其皮", "去腸", "人面", "婦人衣", "而取之", "春" & VBA.ChrW(-31631), "春蔥", "芋")
        dict.Add "蹇", Array("諤", "驢", "叔", "氏", "周", "序辰", "材望", "毅然", "已莫", "步", "吃")
        dict.Add "渙", Array("然", "散", "遂踰", "指陳是", "贓污")
        dict.Add VBA.ChrW(28067), Array("然", "散", "遂踰", "指陳是", "贓污")
        dict.Add "夬", Array("切")
        dict.Add "頤", Array("和園", "正叔", "字正", "茂叔", "字茂", "指氣使", "庵", "菴", "盦", "谷", "下有皮", "所以不")
        dict.Add VBA.ChrW(-26587), Array("和園", "正叔", "字正", "茂叔", "字茂", "指氣使", "庵", "菴", "盦", "谷", "下有皮", "所以不")
        dict.Add "萃", Array("於此", "此書", "于一", "古人", "之成", "其家", "江", "諸庫", "為一書", "於一門", "屼", "東壁", "育", "瀛洲之", "百王致治", "成六卷")
        dict.Add "艮", Array("岳", "嶽", "齋")
        dict.Add "賁", Array("赫", "隅之", "軍之將", "讀為僨", "字", "軍")
        
        dict.Add "巽", Array("巖", VBA.ChrW(23891), "懦", "字仲", "亦不較", "後仕", "風疏", "博學", "參政", "前問")
        dict.Add VBA.ChrW(14514), Array("巖", "懦", "字仲", "亦不較", "後仕", "風疏", "博學", "參政", "前問")
        
        dict.Add "蠱", Array("惑人心", "蕩", "於心", "自埋", "之詐", "實生子", "毒", "發膨", "主", "之屬", "有鬼", "不絕", "者是也", "嫩", VBA.ChrW(23280), "皆中人")
        dict.Add "暌", Array("」作「睽", "作睽", "隔", "叟")
        dict.Add "坎", Array(VBA.ChrW(22728), "鼓")
        dict.Add "遯", Array("跡", VBA.ChrW(-28679), "齋", "世修真", "世不見")
        dict.Add "兌", Array("命")
        dict.Add VBA.ChrW(20817), Array("命")
        dict.Add VBA.ChrW(20810), Array("命")
        dict.Add "小畜", Array("集")
        dict.Add "同人", Array("醵錢")
        dict.Add "中孚", Array("禪子")
        dict.Add "大有", Array("功", "力", "父風", "警省", "逕庭", "李僧", "所瀦", "事在", "大積", "稽驗")
        dict.Add "無妄", Array("想")
        dict.Add VBA.ChrW(26080) & "妄", Array("想")
        dict.Add VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522), Array("想")
        dict.Add "大過", Array("人者")
        dict.Add "小過", Array("宜寬", "宜" & VBA.ChrW(23515))
        dict.Add "初九", Array("日")
        dict.Add "初六", Array("日")
        dict.Add "六四", Array("）")
        dict.Add "上六", Array("十里")
        dict.Add "文言", Array("惇懿")
        dict.Add "少陰", Array("雨")
        dict.Add "咎", Array("繇", "陶", "彼", "單", "累", "在人之相", VBA.ChrW(-10172) & VBA.ChrW(-8632) & "人之相", "徵之應")
        dict.Add "敦復", Array("學士")
        dict.Add "知臨", Array("江", "泉")
        dict.Add "豚魚", Array("麵")
        dict.Add "直方", Array("殂之")
        dict.Add "晝日", Array("無事", "愈長")
        dict.Add "大衍", Array("曆")

        
        Set 易學KeywordsToMark_Exam_Followed_Avoid = dict
        Set followed_Avoid = dict
    Else
        Set 易學KeywordsToMark_Exam_Followed_Avoid = followed_Avoid
    End If
End Property
Rem 某關鍵字不能在某個語句裡面 20240914
Property Get 易學KeywordsToMark_Exam_InPhrase_Avoid() As Scripting.Dictionary
    If inPhrase_Avoid Is Nothing Then
        Dim dict As New Scripting.Dictionary
        dict.Add "易", Array("李易安", "居易錄", "蘇易簡", "杜易簡", "張易之", "此易事", "人易老", "人易去", "兩易其任", "最易淆", "以賢易不肖", "以治易亂", _
            "江山易主", "深耕易耨", "周易癡", "豈易得", "豈易逢", "君子易事", "竇易直", "後易為", "此易彼", "不以易其", "得而易失", "人而易私", "人而易" & VBA.ChrW(-10155) & VBA.ChrW(-8352), "以易此", "無以易也", "延易府", "皮易布", "金易一", "名易知", "雖易得", "誠至易", "則易束", "者易窺", "置易制", "難、易相", "事易行", "今易以冠服", "俱易以", "易知也", "市易法", "可以易一飽", "賈易不", "和易之氣", "氣易壅", "心易放", "躁而易遷", "淺而易洩", _
            "而易見", "而易起", "簡而易行", "而易晦", "請易之", "多易之", "公易其介", "不能易也", "最易得", "其情易見", "亢易招", "以易鹽米", "頓易故", "谿易雨", "者易訓", "心易偏", "平心易氣", "故其說易差", "始易為力", "乃易合", "而易彼", "慢而易之", "而易治", "樂易之", "視而易之", "邪正易位", "最易牽引", _
            "以易心", "以易處之", "反易天明", "因易其韻", "以縉儒者易之", "因改易本文而", "尺寸易以", "客易主位", "願亦易愜", "則易發", "根易發", "客易位", "以改易之", "主易位", "必易疑", "易信必易疑", "平實易知", "簡要易行", _
            "而易陵", "皆易之以", "事在易而求", "儉則易足", "曾易占", "不能易焉", "深德易占", "後易占", "將易吾", "者易犯", "狹則易足", "悔易勿輕踵", "智易窮", "何以易窮", "更有易見者", "而易散", _
            "至易事", "敢易也", "河東、易定", "易定、魏博", "豈易堪此", "河東易定", "脆易折", "彰癉易位", "易定魏博", "柔易治", "最易生", _
            "疾易作", "病易除", "互易注文", "無難，是以易也", "人易從", "何以易生", "易放而難操", "氣易乘", "物亦易給", "錢差易 ", "言易墜", "所以易放", "者易直", "是易言也", "須易之", "之易以仰測", "蓴絲之易", "皆易黃屋", "皆易" & ChrW(-24892) & "屋", "時易以新", _
            "思以易天下", "則易使", "不能易君子", "謀易太子", "侮易承業", "經師易獲", "成易具", "綿布易之", "勃然易動", "布相易云", "浴易服", "惡易敗", "一易根", "荄於易地", "三易其音", _
            "因易名曰", "以字易名", "食而易得", "籠易制", "論人易矣", "故易辨也", "不易長", "惡幾易熾", "和易近人", "以日易月", "大樂必易", "以之易業", "以之易用", "不欲易也", "基而易舊", _
            "則易入於", "皆易與之", "死易不義", "原易范", "鐵易土", "兵易進", "亦易退", "實而易虛", "江易舟", "苦于易奪", "傷于易敗", "相信，易以成功", _
            "非易事", "有易於內者", "蠹者易之", VBA.ChrW(-30681) & "者易之", "首卷易之", "而易為待補", "以易種姓", "晚年易名", "木易朽", "歲月易得", VBA.ChrW(27507) & "月易得", _
            "豈易說", "人人易知", "人易知焉", "華飾者易于近名", "則易以興起", "而其易明" & VBA.ChrW(-10175) & VBA.ChrW(-8614))
        
        dict.Add "卦", Array("八卦山", "八卦殿")
        dict.Add "爻", Array("如交爻字")
        dict.Add "乾", Array("保乾圖", "大乾廟", "乳乾者", "雨乾時", "五季乾坤")
        dict.Add "剝", Array("解剝而發明", "造剝洛陽")
        dict.Add VBA.ChrW(21093), Array("解" & VBA.ChrW(21093) & "而發明", "造" & VBA.ChrW(21093) & "洛陽")
        dict.Add "豫", Array("事豫則", "康哉豫矣", "人豫知", "而豫求", "能豫逆", "豫射其", "於豫圖", "豫州。豫。舒也")
        dict.Add "蹇", Array("剛蹇絕", "歲蹇剝", "辭以蹇字", "驕蹇之心")
        
        dict.Add "頤", Array("翁頤昌", "方頤大口", "解頤撫掌", "兩頤間", "解頤而悅")
        dict.Add VBA.ChrW(-26587), Array("翁" & VBA.ChrW(-26587) & "昌", "方" & VBA.ChrW(-26587) & "大口", "解" & VBA.ChrW(-26587) & "而悅")
        
        dict.Add "暌", Array("有暌談笑")
        dict.Add "蠱", Array("以蠱留人", "以蠱而", "而蠱者", "以蠱大", "中蠱者", "內蠱豔")
        dict.Add "萃", Array("檀萃文")
        dict.Add "賁", Array("古賁灰", "隸書賁字", "廣賁之", "苗賁皇")
        
        dict.Add "巽", Array("即巽也", "東巽泉", "邀巽二", "公巽曰", "公巽默")
        dict.Add VBA.ChrW(14514), Array("即" & VBA.ChrW(14514) & "也", "東" & VBA.ChrW(14514) & "泉", "公" & VBA.ChrW(14514) & "曰", "公" & VBA.ChrW(14514) & "默")
        
        dict.Add "同人", Array("底同人不", "不同人生", "不同人能", "皆指同人，未知", "要之以己同人則")
        dict.Add "夬", Array("十七夬部")
        
        dict.Add "渙", Array("王渙之", "令渙往", "叱渙事")
        dict.Add VBA.ChrW(28067), Array("王" & VBA.ChrW(28067) & "之", "令" & VBA.ChrW(28067) & "往", "叱" & VBA.ChrW(28067) & "事")
        
        dict.Add "大有", Array("甚大有朱", "碩大有顒")
        dict.Add "大壯", Array("祖大壯之")
        dict.Add "大過", Array("無大過惡")
        dict.Add "小過", Array("忘小過以成", VBA.ChrW(14592) & "小過以成")
        
        dict.Add "既濟", Array("河既濟真")
        dict.Add VBA.ChrW(26083) & "濟", Array("河" & VBA.ChrW(26083) & "濟真")
        
        dict.Add "無妄", Array("人無妄取", "物無妄費")
        dict.Add VBA.ChrW(26080) & "妄", Array("人" & VBA.ChrW(26080) & "妄取", "物" & VBA.ChrW(26080) & "妄費")
        dict.Add VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522), Array("人" & VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522) & "取", "物" & VBA.ChrW(26080) & VBA.ChrW(-10171) & VBA.ChrW(-8522) & "費")
        dict.Add "初九", Array("虞初九百")
        dict.Add "九二", Array("一九二０")
        dict.Add "九三", Array("廿九三十")
        dict.Add "九五", Array("一九五九")
        dict.Add "用九", Array("欲用九月")
        dict.Add "六二", Array("一六二四", "一六二六", "長六二公")
        dict.Add "六四", Array("一六四八")
        dict.Add "上六", Array("以上六事", "已上六事", "以上六十")
        dict.Add "用六", Array("威用六極")
        dict.Add "文言", Array("承此文言之", "古文言『毌", "古文言毌")
        dict.Add "存誠", Array("心存誠敬")
        dict.Add "童蒙", Array("民童蒙不")
                    
        Set 易學KeywordsToMark_Exam_InPhrase_Avoid = dict
        Set inPhrase_Avoid = dict
    Else
        Set 易學KeywordsToMark_Exam_InPhrase_Avoid = inPhrase_Avoid
    End If
End Property
Rem 檢測關鍵字
'Function 易學KeywordsToMark_Exam()
'
'End Function

