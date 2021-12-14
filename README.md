# TextForCtext
為了《中國哲學書電子化計劃》（簡稱ctext）輸入用

尤其由[中研院史語所《漢籍電子文獻資料庫》](http://hanchi.ihp.sinica.edu.tw/ihp/hanji.htm)輸入《十三經注疏》、[《維基文庫》](https://zh.wikisource.org/zh-hant/)輸入《四部叢刊》本諸書圖文對照時，輔助加速，避免人工之失誤。感恩感恩　南無阿彌陀佛

昨天邊寫程式、測試，邊完成了[《四部叢刊》《南華真經》（《莊子》）](https://ctext.org/wiki.pl?if=gb&chapter=941297#lib77891.114)[第一份文件](https://ctext.org/wiki.pl?if=gb&chapter=941297&action=history)輸入的工作；真是感覺像飛了起來，和之前用手、眼合作判斷分行切割的速度，懸若天壤、判若兩人。感恩感恩　讚歎讚歎　南無阿彌陀佛

## 介面簡介：

textBox1:文本框

textBox2:尋找文本用

textBox3: URL瀏覽參照用（點一下，貼上要瀏覽的網址）

textBox4:文本取代用


## 快速鍵一覽：

### 在表單（視窗）任何位置按下：

Ctrl + 1 ：執行 Word Sub 巨集指令「漢籍電子文獻資料庫文本整理_以轉貼到中國哲學書電子化計劃」

Ctrl + 3 ：執行 Word Sub 巨集指令「漢籍電子文獻資料庫文本整理_十三經注疏」

Ctrl + 4 ：執行 Word Sub 巨集指令「維基文庫四部叢刊本轉來」

Ctrl + f ：移至 textBox2 準備尋找文本

Ctrl + s ：儲存文本至 DropBox 根目錄下的「cText.txt」檔案

Ctrl + q 或 Alt + q：據第一行(段)長度來將textBox1中的文本分行分段

Ctrl + \ （反斜線） ： 清除textBox1文本插入點後的分段

Ctrl + PageUp ：根據 textBox3所載的網址，瀏覽ctext書影的上一頁

Ctrl + PageDown ： 根據 textBox3所載的網址，瀏覽ctext書影的下一頁

### 在 textBox1 中按下以下組合鍵：


Ctrl + 8 ：如同鍵入「　」1個全形空格，且各個空格間有分段符

Ctrl + 9 ：如同鍵入「　　」2個全形空格，且各個空格間有分段符

Ctrl + 0 ：如同鍵入「　　　　」4個全形空格，且各個空格間有分段符

Ctrl + + （函數字鍵盤） 或 Ctrl + 5 (數字鍵盤） ：將插入點或選取文字（含）之前的文本剪下貼到 ctext 的[簡單修改模式]框中，並按下「保存編輯」鈕，且在[簡單修改模式]下開啟下一頁準備編輯文本。

Ctrl+Shift+↑：從插入點開啟向前選取整段

Ctrl+Shift+↓：從插入點開啟向後選取整段





