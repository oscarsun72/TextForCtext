# TextForCtext
為了《中國哲學書電子化計劃》（簡稱ctext）輸入用

尤其由[中研院史語所《漢籍電子文獻資料庫》](http://hanchi.ihp.sinica.edu.tw/ihp/hanji.htm)輸入《十三經注疏》、[《維基文庫》](https://zh.wikisource.org/zh-hant/)輸入《四部叢刊》本諸書圖文對照時，輔助加速，避免人工之失誤。感恩感恩　南無阿彌陀佛

昨天邊寫程式、測試，邊完成了[《四部叢刊》《南華真經》（《莊子》）](https://ctext.org/wiki.pl?if=gb&chapter=941297#lib77891.114)[第一份文件](https://ctext.org/wiki.pl?if=gb&chapter=941297&action=history)輸入的工作；真是感覺像飛了起來，和之前用手、眼合作判斷分行切割的速度，懸若天壤、判若兩人。感恩感恩　讚歎讚歎　南無阿彌陀佛

## 介面簡介：
![操作介面](https://github.com/oscarsun72/TextForCtext/blob/master/TextforCtext%E4%BB%8B%E9%9D%A2%E7%B0%A1%E4%BB%8B.png)

textBox1:文本框

textBox2:尋找文本用

textBox3: URL瀏覽參照用（點一下，貼上要瀏覽的網址）

textBox4:文本取代用

## 基本功能
（一切為加速 ctext 網站圖文對照文本編輯而設。目前不免以本人主觀習慣為主）
- 當 textBox1、textBox3 內容為空白時，滑鼠左鍵點一下即讀進剪貼簿內容。
- 當離開 textBox2 或 textBox4 文字方塊時，即自執行在 textBox1 尋找或取代文字功能
- 自動依據第一分段字數將 textBox1 插入點其後的文本分段。按下左上方的按鈕或按下 Ctrl + q （參見下文）
- 清除 textBox1 插入點後的分段。按下 Ctrl + \ 
- 將 textBox1 插入點前或含選取文字前的文本貼入 ctext [簡單修改模式]框中，並自動按下「保存編輯」鈕，且在 Chrome 瀏覽器新分頁[簡單修改模式]下開啟下一頁準備編輯文本。（參見下文）
- 當 textBox4 取得焦點（插入點）時自行調整其大小，焦點離開時恢復


## 快速鍵一覽：

### 在表單（操作介面視窗）任何位置按下：

Ctrl + 1 ：執行 Word Sub 巨集指令「漢籍電子文獻資料庫文本整理_以轉貼到中國哲學書電子化計劃」【 附件即有 Word VBA 相關模組 】

Ctrl + 3 ：執行 Word Sub 巨集指令「漢籍電子文獻資料庫文本整理_十三經注疏」

Ctrl + 4 ：執行 Word Sub 巨集指令「維基文庫四部叢刊本轉來」

Ctrl + f ：移至 textBox2 準備尋找文本

Ctrl + s ：儲存文本至 DropBox 根目錄下的「cText.txt」檔案

F5 ：重新載入所儲存的文本

Ctrl + Shift + c ：將textBox1的文本複製到剪貼簿 以備用

Ctrl + q 或 Alt + q：據第一行(段)長度來將textBox1中的文本分行分段

Ctrl + \ （反斜線） ： 清除textBox1文本插入點後的分段

Ctrl + PageUp ：根據 textBox3所載的網址，瀏覽ctext書影的上一頁

Ctrl + PageDown ： 根據 textBox3所載的網址，瀏覽ctext書影的下一頁

### 在 textBox1 中按下以下組合鍵：


Ctrl + 8 ：如同鍵入「　」1個全形空格，且各個空格間有分段符

Ctrl + 9 ：如同鍵入「　　」2個全形空格，且各個空格間有分段符

Ctrl + 0 ：如同鍵入「　　　　」4個全形空格，且各個空格間有分段符

Ctrl + h ：移至 textBox4 準備取代文本文字

Ctrl + + （加號，含函數字鍵盤） 或 Ctrl + 5 (數字鍵盤） ：將插入點或選取文字（含）之前的文本剪下貼到 ctext 的[簡單修改模式]框中，並按下「保存編輯」鈕，且在[簡單修改模式]下開啟下一頁準備編輯文本。

Ctrl+Shift+↑：從插入點開啟向前選取整段

Ctrl+Shift+↓：從插入點開啟向後選取整段


### 在 textBox2、4 中按下以下鍵：

Ctrl+ 滑鼠左鍵：清除框中所有文字


