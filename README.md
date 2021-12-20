# TextForCtext
為了《中國哲學書電子化計劃》（簡稱ctext）輸入用

尤其由[中研院史語所《漢籍電子文獻資料庫》](http://hanchi.ihp.sinica.edu.tw/ihp/hanji.htm)輸入《十三經注疏》、[《維基文庫》](https://zh.wikisource.org/zh-hant/)輸入《四部叢刊》本諸書圖文對照時，輔助加速，避免人工之失誤。感恩感恩　南無阿彌陀佛

昨天邊寫程式、測試，邊完成了[《四部叢刊》《南華真經》（《莊子》）](https://ctext.org/wiki.pl?if=gb&chapter=941297#lib77891.114)[第一份文件](https://ctext.org/wiki.pl?if=gb&chapter=941297&action=history)輸入的工作；真是感覺像飛了起來，和之前用手、眼合作判斷分行切割的速度，懸若天壤、判若兩人。感恩感恩　讚歎讚歎　南無阿彌陀佛
20211217在不斷修改增潤的過程中，也將把[此部《莊子》](https://ctext.org/library.pl?if=gb&res=77451)書[維基文本](https://ctext.org/wiki.pl?if=gb&res=393223)建置完畢了。感恩感恩　讚歎讚歎　南無阿彌陀佛 20211218：1951 [建置完畢](https://ctext.org/wiki.pl?if=gb&res=393223) 感恩感恩　南無阿彌陀佛

## 介面簡介：
![操作介面](https://github.com/oscarsun72/TextForCtext/blob/master/TextforCtext%E4%BB%8B%E9%9D%A2%E7%B0%A1%E4%BB%8B.png)

textBox1:文本框

textBox2:尋找文本用

textBox3: URL瀏覽參照用（點一下，貼上要瀏覽的網址）

textBox4:文本取代用

## 基本功能
（一切為加速 ctext 網站圖文對照文本編輯而設。目前不免以本人主觀習慣為主）
- 操作介面之表單視窗預設為最上層顯示，當表單視窗不在作用中時，只要焦點/插入點不在 textBox2 中，即自動隱藏至系統右下方之系統列/任務列中，當滑鼠滑過任務列中的縮圖ico時，即還原/恢復視窗窗體
- 當 textBox1、textBox3 內容為空白時，滑鼠左鍵點一下即讀進剪貼簿內容。
- 當離開 textBox2 或 textBox4 文字方塊時，即自執行在 textBox1 尋找或取代文字功能（增罝透過 button2 的切換，決定是否只在選取區中執行尋找、取代）
　
　若沒有找到 textBox2 內的指定字串時，該方塊會顯示紅色半秒，並且插入點還是保留在該方塊中，待繼續輸入其他尋找條件。
- 自動依據第一分段字數將 textBox1 插入點其後的文本分段。按下左上方的按鈕或按下 Ctrl + q （參見下文）
- 清除 textBox1 插入點後的分段。按下 Ctrl + \ 
- 當 textBox4 取得焦點（插入點）時自行調整其大小，焦點離開時恢復
- 儲存textbox1文本，（快速鍵 Ctrl + s ： 儲存路徑在 DropBox 的預設安裝路徑的根目錄（C:\Users\ **使用者名** \Dropbox\）中，名「cText.txt」）
- 將 textBox1 插入點前或含選取文字前的文本貼入 ctext [簡單修改模式]框中，並自動按下「保存編輯」鈕，且在 Chrome 瀏覽器新分頁[簡單修改模式]下開啟下一頁準備編輯文本。（參見下文，快速鍵「Ctrl + + 」處。執行此項時，自動在背後進行該次頁面文本的備份，儲存路徑在 DropBox 的預設安裝路徑的根目錄中，名「cTextBK.txt」，是以追加的方式備份。）
- 自動文本備份及更正備份功能（在複製到剪貼簿時），蓋剛才辛苦做的《四部叢刊》本《南華真經》（《莊子》）第四冊文本，竟然莫名其妙地遺失了，只殘留最後幾頁，整個半天乃至一天的勞動，化為烏有。20211217
- 預設為最上層顯示，若textBox1值為空，則按下Esc鍵會隱藏到任務列（系統列）中；點一下即恢復

## 快速鍵一覽：

### 在表單（操作介面視窗）任何位置按下：

Ctrl + 1 ：執行 Word VBA Sub 巨集指令「漢籍電子文獻資料庫文本整理_以轉貼到中國哲學書電子化計劃」【 附件即有 Word VBA 相關模組 】

Ctrl + 3 ：執行 Word VBA Sub 巨集指令「漢籍電子文獻資料庫文本整理_十三經注疏」

Ctrl + 4 ：執行 Word VBA Sub 巨集指令「維基文庫四部叢刊本轉來」

Ctrl + f ：移至 textBox2 準備尋找文本

Ctrl + s ：儲存文本至 DropBox 根目錄下的「cText.txt」檔案

F5 ：重新載入所儲存的文本

F12  ： 更新最後的備份頁文本

Ctrl + Shift + c ：將textBox1的文本複製到剪貼簿 以備用

Ctrl + PageUp ：根據 textBox3所載的網址，瀏覽ctext書影的上一頁

Ctrl + PageDown ： 根據 textBox3所載的網址，瀏覽ctext書影的下一頁



### 在 textBox1 中按下以下組合鍵：

Insrt : 如 MS Word ，切換插入/取代文字模式（尚無能力寫，希望能做到）

Ctrl + q 或 Alt + q：據第一行(段)長度來將textBox1中的文本分行分段

Ctrl + \ （反斜線） ： 清除textBox1文本插入點後的分段

按下 F1 鍵：以找到的字串位置**前**分行分段

按下 Pause Break 鍵：以找到的字串位置**後**分行分段

Ctrl + 8 ：如同鍵入「　」1個全形空格，且各個空格間有分段符

Ctrl + 9 ：如同鍵入「　　」2個全形空格，且各個空格間有分段符

Ctrl + 0 ：如同鍵入「　　　　」4個全形空格，且各個空格間有分段符

F2 : 全選/取消全選框裡文字。若原有選取文字則取消選取至其尾端

F3 ： 在textBox1 尋找下一個符合所選取的字串；如果沒有選取，則以 textBox2 的字串為據

Shift + F3 ： 在textBox1 尋找上一個符合所選取的字串；如果沒有選取，則以 textBox2 的字串為據

Ctrl + h ：移至 textBox4 準備取代文本文字

Ctrl + F12 ：就 textBox1 所選之字串，執行「[查詢國語辭典.exe](https://github.com/oscarsun72/lookupChineseWords.git)」以查詢網路詞典

Alt + G ：就 textBox1 所選之字串，執行「[網路搜尋_元搜尋-同時搜多個引擎.exe](https://github.com/oscarsun72/SearchEnginesConsole.git)」以查詢 Google 等網站

Ctrl + + （加號，含函數字鍵盤） 或 Ctrl + -（數字鍵盤）  或 Ctrl + 5 (數字鍵盤） ：將插入點或選取文字（含）之前的文本剪下貼到 ctext 的[簡單修改模式]框中，並按下「保存編輯」鈕，且在[簡單修改模式]下于瀏覽器新頁籤開啟下一頁準備編輯文本，並回到前一頁籤以供檢視所貼上之文本是否無誤。

Ctrl+Shift+↑：從插入點開啟向前選取整段

Ctrl+Shift+↓：從插入點開啟向後選取整段

Alt + 滑鼠左鍵 ： 更新最後的備份頁文本

Ctrl + Shift + Delete ： 將選取文字於文本中全部清除

### 在 textBox2、4 中按下以下鍵：

F2 : 全選/取消全選框裡文字。若原有選取文字則取消選取至其尾端

Ctrl+ 滑鼠左鍵：清除框中所有文字

### 在 textBox2 ：
- 按下 F1 鍵：以找到的字串位置**前**分行分段
- 按下 Pause Break 鍵：以找到的字串位置**後**分行分段
- Ctrl + + （加號，含函數字鍵盤） 或 Ctrl + -（數字鍵盤）  或 Ctrl + 5 (數字鍵盤） ：將插入點或選取文字（含）之前的文本剪下貼到 ctext 的[簡單修改模式]框中，並按下「保存編輯」鈕，且在[簡單修改模式]下于瀏覽器新頁籤開啟下一頁準備編輯文本，並回到前一頁籤以供檢視所貼上之文本是否無誤。


### textBox4 取代文字用方塊框
- 若有指定要取代的文字，進入後會自動填入之前用以取代過的文字以便輸入（即自動填入對應的預設值）
- 指定要被取代的文字方式：1. 在textBox1中選取文字；2. 若按下 button2 切換成「選取文」（背景紅色）狀態，則將以 textBox2 內的文字為被取代的字串。


## 參照資源
（大於 25MB 無法在此上傳的檔案則表列於此。若有疏漏，尚祈提醒末學。感恩感恩　南無阿彌陀佛）
- [查字.mdb](https://www.dropbox.com/s/6vn9hi7i95cbhy4/%E6%9F%A5%E5%AD%97.mdb?dl=0)
- [《重編國語辭典修訂本》資料庫.mdb](https://www.dropbox.com/s/a6t3yhou4smpdv7/%E3%80%8A%E9%87%8D%E7%B7%A8%E5%9C%8B%E8%AA%9E%E8%BE%AD%E5%85%B8%E4%BF%AE%E8%A8%82%E6%9C%AC%E3%80%8B%E8%B3%87%E6%96%99%E5%BA%AB.mdb?dl=0)


