# TextForCtext
為了《中國哲學書電子化計劃》（簡稱ctext）輸入用

尤其由[中研院史語所《漢籍電子文獻資料庫》](http://hanchi.ihp.sinica.edu.tw/ihp/hanji.htm)輸入《十三經注疏》、[《維基文庫》](https://zh.wikisource.org/zh-hant/)輸入《四部叢刊》本諸書圖文對照時，輔助加速，避免人工之失誤。感恩感恩　南無阿彌陀佛

昨天邊寫程式、測試，邊完成了[《四部叢刊》《南華真經》（《莊子》）](https://ctext.org/wiki.pl?if=gb&chapter=941297#lib77891.114)[第一份文件](https://ctext.org/wiki.pl?if=gb&chapter=941297&action=history)輸入的工作；真是感覺像飛了起來，和之前用手、眼合作判斷分行切割的速度，懸若天壤、判若兩人。感恩感恩　讚歎讚歎　南無阿彌陀佛
20211217在不斷修改增潤的過程中，也將把[此部《莊子》](https://ctext.org/library.pl?if=gb&res=77451)書[維基文本](https://ctext.org/wiki.pl?if=gb&res=393223)建置完畢了。感恩感恩　讚歎讚歎　南無阿彌陀佛 20211218：1951 [建置完畢](https://ctext.org/wiki.pl?if=gb&res=393223) 感恩感恩　南無阿彌陀佛

其他最新進度，詳鄙人此帖： [transferkit IPFS 永遠保存的電子文獻-藏富天下 暨《中國哲學書電子化計劃》愚所輸入完竣之諸本-任真的網路書房-千慮一得齋OnLine-觀死書齋原著及電子化文獻(不屑智慧財產權)歡迎多利用共玉于成](https://oscarsun72.blogspot.com/2022/02/transferkit-ipfs.html) 

## 介面簡介：
![操作介面](https://github.com/oscarsun72/TextForCtext/blob/master/TextforCtext%E4%BB%8B%E9%9D%A2%E7%B0%A1%E4%BB%8B.png)

textBox1:文本框

textBox2:尋找文本用

textBox3: URL瀏覽參照用（點一下，貼上要瀏覽的網址）

textBox4:文本取代用

## 基本功能
（一切為加速 ctext 網站圖文對照文本編輯而設。目前不免以本人主觀習慣為主）
- 操作介面之表單視窗預設為最上層顯示，當表單視窗不在作用中時，只要焦點/插入點不在 textBox2 中，即非最上層顯示；若恢復作用中時（取得焦點時），則最上層顯示。（或：即自動隱藏至系統右下方之系統列/任務列中，當滑鼠滑過任務列中的縮圖ico時，即還原/恢復視窗窗體－－此隱藏功能先棄置）
- 當 textBox1、textBox3 內容為空白時，滑鼠左鍵點一下即讀進剪貼簿內容。
- 當離開 textBox2 或 textBox4 文字方塊時，即自執行在 textBox1 尋找或取代文字功能（增罝透過 button2 的切換，決定是否只在選取區中執行尋找、取代）
　
　若沒有找到 textBox2 內的指定字串時，該方塊會顯示紅色半秒，並且插入點還是保留在該方塊中，待繼續輸入其他尋找條件。若要離開，須清空內容
 
  若符合尋找的字串並非獨一無二，則 textBox2 會顯示黃色。
  
  若只有一個符合尋找字串，則 textBox2 會顯示黃綠色，並發出提示音；若再配合按下 F1 鍵（以找到的字串位置前分行分段）、按下 Pause Break 鍵（以找到的字串位置後分行分段）該可加速分行分段
  
- 自動依據第一分段字數將 textBox1 插入點其後的文本分段。按下左上方的按鈕或按下 Ctrl + q （參見下文）
- 清除 textBox1 插入點後的分段。按下 Ctrl + \ 
- 當 textBox4 取得焦點（插入點）時自行調整其大小，焦點離開時恢復
- 儲存textbox1文本，（快速鍵 Ctrl + s ： 儲存路徑在 DropBox 的預設安裝路徑的根目錄（C:\Users\ **使用者名** \Dropbox\）中，名「cText.txt」）
- 將 textBox1 插入點前或含選取文字前的文本貼入 ctext [簡單修改模式]框中，並自動按下「保存編輯」鈕，且在 Chrome 瀏覽器新分頁[簡單修改模式]下開啟下一頁準備編輯文本。（參見下文，快速鍵「Ctrl + + 」處。執行此項時，自動在背後進行該次頁面文本的備份，儲存路徑在 DropBox 的預設安裝路徑的根目錄中，名「cTextBK.txt」，是以追加的方式備份。）
- 自動文本備份及更正備份功能（在複製到剪貼簿時），蓋剛才辛苦做的《四部叢刊》本《南華真經》（《莊子》）第四冊文本，竟然莫名其妙地遺失了，只殘留最後幾頁，整個半天乃至一天的勞動，化為烏有。20211217
- 預設為最上層顯示，則按下Esc鍵或滑鼠中鍵會隱藏到任務列（系統列）中；滑鼠在其 ico 圖示上滑過即恢復
- 要清除所選文字，則選取其字，然後在 textBox4 輸入兩個英文半形雙引號 「""」（即表空字串），則不會取代成「""」，而是清除之。
- Ctrl + z 還原 textBox1 文本功能。支援打字與取代文字後的還原。還原上限為50次。

- isShortLine() 配合「每行字數判斷用」資料表作為每行字數判斷參考 


## 快速鍵一覽：

### 在表單（操作介面視窗）任何位置按下：

F5 ：重新載入所儲存的文本

F9 ：重啟小小輸入法

F12  ： 更新最後的備份頁文本

Esc 則按下Esc鍵會隱藏到任務列（系統列）中；滑鼠在其 ico 圖示上滑過即恢復

Ctrl + 1 ：執行 Word VBA Sub 巨集指令「漢籍電子文獻資料庫文本整理_以轉貼到中國哲學書電子化計劃」【 附件即有 [Word VBA](https://github.com/oscarsun72/TextForCtext/tree/master/WordVBA) 相關模組 】

Ctrl + 3 ：執行 Word VBA Sub 巨集指令「漢籍電子文獻資料庫文本整理_十三經注疏」

Ctrl + 4 ：執行 Word VBA Sub 巨集指令「維基文庫四部叢刊本轉來」

Ctrl + f ：移至 textBox2 準備尋找文本

Ctrl + n ：在新頁籤開啟 google 網頁，以備用（在預設瀏覽器為 Chrome 時）

Ctrl + s 或 Shift + F12：儲存文本至 DropBox 根目錄下的「cText.txt」檔案

Ctrl + w 關閉 Chrome 網頁頁籤

Ctrl + PageUp ：根據 textBox3所載的網址，瀏覽ctext書影的上一頁

Ctrl + PageDown ： 根據 textBox3所載的網址，瀏覽ctext書影的下一頁

Ctrl + + （加號，含函數字鍵盤） 或 Ctrl + -（數字鍵盤）  或 Ctrl + 5 (數字鍵盤） 或 Alt + + ：將插入點或選取文字（含）之前的文本剪下貼到 ctext 的[簡單修改模式]框中，並按下「保存編輯」鈕，且在[簡單修改模式]下于瀏覽器新頁籤開啟下一頁準備編輯文本，並回到前一頁籤以供檢視所貼上之文本是否無誤。

Ctrl + Shift + c ：將textBox1的文本複製到剪貼簿 以備用

Alt + ←：視窗向左移動30dpi（+ Ctrl：徵調；插入點在textBox1時例外）

Alt + →：視窗向右移動30dpi（+ Ctrl：徵調；插入點在textBox1時例外）

Alt + F1 ：切換 textbox1 之字型： 切換支援 CJK - Ext 擴充字集的大字集字型

滑鼠上一頁鍵： 同Ctrl + PageUp

滑鼠下一頁鍵： 同Ctrl + PageDown

按下 Ctrl 、Alt 或 Shift 任一鍵再啟用表單成為現用的（activate form1)則會啟動自動輸入( auto paste to Quick Edit textbox in Ctext).

按下 Ctrl + * （Multiply）設定為將《四部叢刊》資料庫所複製的文本在表單得到焦點時直接貼到 textBox1 的末尾,或反設定

按下 Ctrl + / （Divide） 切換自動連續輸入功能

按下 Ctrl + Shift + / （Divide）  切換 check_the_adjacent_pages 值

Ctrl + Shift + t 同Chrome瀏覽器 --還原最近關閉的頁籤

### 在 textBox1 中按下以下組合鍵：

Insert : 如 MS Word ，切換插入/取代文字模式（hit! 還原機制亦大致成功了。還原上限為50次。Ctrl + z ： 支援打字輸入時及取代文字時的還原。）

Ctrl + q 或 Alt + q：據游標（插入點）所在前1段的長度來將textBox1中的文本分行分段

Ctrl + \ （反斜線） 或 Alt + \ ： 清除textBox1文本插入點後的分段

按下 F1 鍵：以找到的字串位置**前**分行分段（在文字選取內容接的不是newline時；若是，且選取長度等於常數「predictEndofPageSelectedTextLen」則進行自動貼入 Ctext 的 quit edit 方塊中）

按下 Pause Break 鍵：以找到的字串位置**後**分行分段

按下 Scroll Lock 將字數較少的行/段落尾末標上「<p>」符號

Ctrl + 6 ：鍵入「{{」

Ctrl + Shift + 6 ：鍵入「}}」(在前面有「{」時，按下 Alt + i 也可以鍵入此值)

Alt + Shift + 6 或 Alt + s：小注文不換行：notes_a_line()

Ctrl + 7 ：如同鍵入「。}}」，於《周易正義》〈彖、象〉辭時適用

Ctrl + 8 ：如同鍵入「　」1個全形空格，且各個空格間有分段符

Ctrl + 9 ：如同鍵入「　　」2個全形空格，且各個空格間有分段符

Ctrl + 0 ：如同鍵入「　　　　」4個全形空格，且各個空格間有分段符

Alt + 1 : 鍵入本站制式留空空格標記「􏿽」：若有選取則取代全形空格「　」為「􏿽」

Alt + Shift + 1 如宋詞中的換片空格，只將文中的空格轉成空白，其他如首綴前罝以明段落或標題者不轉換

Alt + 2 : 鍵入全形空格「　」

Alt + Shift + 2 : 將選取區內的「<p>」取代為「|」 ，而「　」取代為「􏿽」並清除「*」且將無「|」前綴的分行符號加上「|」（詩偈排版格式用）

Alt + 3 : 鍵入全形空格「〇」

Alt + 6 : 鍵入 「"}}"+ newline +"{{"」

Alt + 7 : 鍵入 「"}}"+ newline +"{{"」

Alt + 8 : 鍵入 「　　*」

Alt + 9 : 鍵入 「 

Alt + 0 : 鍵入 『

Alt + i : 鍵入 》（如 MS Word 自動校正，會依前面的符號作結尾號（close），如前是「〈」，則轉為「〉」……）

Alt + j : 鍵入換行分段符號（newline）（同 Ctrl + j 的系統預設）

Alt + p 或 Alt + ` : 鍵入 "<p>" + newline（分行分段符號）；若置於行/段之首，則會自動移至前一段末再執行

Alt + s 或 Alt + Shift + 6 ：小注文不換行 ： notes_a_line()

Alt + Shift + s :  所有小注文都不換行

Alt + u : 鍵入 《

Alt + y : 鍵入 〈

Alt + . : 鍵入 ·

Alt + Del : 刪除插入點後第一個分行分段

Alt + Insert ：將剪貼簿的文字內容讀入textBox1中

F2 : 全選/取消全選框裡文字。若原有選取文字則取消選取至其尾端

F3 ： 在textBox1 從插入點（游標所在處）開始尋找下一個符合所選取的字串；如果沒有選取，則以 textBox2 的字串為據

Shift + F3 ： 從插入點（游標所在處）開始在textBox1 尋找上一個符合所選取的字串；如果沒有選取，則以 textBox2 的字串為據

F4 ： 重複輸入最後一個字

Shift + F5 ： 在textBox1 回到上1次插入點（游標）所在處（且與最近「charIndexListSize」次瀏覽處作切換，如 MS Word）。charIndexListSize 目前=  3。

F6 : 標題降階（增加標題前之星號）

Alt + F6 或 Alt + F8 或 選取標題文字前之空格再按下 Alt + ` : run autoMarkTitles 自動標識標題（篇名）

F7 ： 每行縮排，即每行/段前空一格；全部縮排的機會少，若要全部，則請將插入點放在全文前端或末尾

Shfit + F7 : 每行凸排: deleteSpacePreParagraphs_ConvexRow()；全部凸排的機會少，若要全部，則請將插入點放在全文前端或末尾

Alt + F7 : 每行縮排一格後將其末誤標之<p>清除:keysSpacePreParagraphs_indent_ClearEnd＿P_Mark；全部縮排的機會少，若要全部，則請將插入點放在全文前端或末尾

F8 或 Alt + ` ： 加上篇名格式代碼

F11 : run replaceXdirrectly() 維基文庫等欲直接抽換之字

Ctrl + h ：移至 textBox4 準備取代文本文字

Ctrl + F12 ：就 textBox1 所選之字串，執行「[查詢國語辭典.exe](https://github.com/oscarsun72/lookupChineseWords.git)」以查詢網路詞典

Alt + G ：就 textBox1 所選之字串，執行「[網路搜尋_元搜尋-同時搜多個引擎.exe](https://github.com/oscarsun72/SearchEnginesConsole.git)」以查詢 Google 等網站

Ctrl + + （加號，含函數字鍵盤） 或 Ctrl + -（數字鍵盤）  或 Ctrl + 5 (數字鍵盤） 或 Alt + + 或 Alt + a ：將插入點或選取文字（含）之前的文本剪下貼到 ctext 的[簡單修改模式]框中，並按下「保存編輯」鈕，且在[簡單修改模式]下于瀏覽器新頁籤開啟下一頁準備編輯文本，並回到前一頁籤以供檢視所貼上之文本是否無誤。

Ctrl + + Shift ：同前，只是按下"Shift"表示不要自動翻到下一頁。

Ctrl + [：從插入點開始向前移至{{前

Ctrl + ]：從插入點開始向後移至}}後

Ctrl + ↑ ：從插入點開始向前移至上一段尾

Ctrl + ↓ 或 Alr + ↓：從插入點開始向後移至這一段末（無分段則不移動）

Ctrl + →：：插入點若在漢字中,從插入點開始向後移至任何非漢字前(即漢字後);反之亦然

Ctrl + ←：：插入點若在漢字中,從插入點開始向後移至任何非漢字後(即漢字前);反之亦然

以上2者若再按下 Shift 鍵則會選取範圍並將其中的「　」取代為「􏿽」。

Ctrl+Shift+↑：從插入點開始向前選取整段

Ctrl+Shift+↓：從插入點開始向後選取整段

Ctrl + <：到下一個<頭頂(原擬作縮小字型1點然，此功能不常用，擬改用滑鼠方式)

Ctrl + >：到下一個>尾端

Ctrl + Shift + Delete ： 將選取文字於文本中全部清除(Ctrl + z 還原功能支援)


Alt + 滑鼠左鍵 ： 更新最後的備份頁文本

Ctrl+ 滑鼠左鍵：在插入點後分行分段（原為切換RichTextBox用）

Ctrl+ 滑鼠右鍵：切換RichTextBox用

Ctrl+ Alt + 滑鼠左鍵：將插入點後的分行分段清除


滑鼠點二下，執行 Ctrl + + , 將插入點所在之前的文本貼到 Ctext 網頁 [簡單修改模式] 文字方塊中

滑鼠上一頁鍵： 同Ctrl + PageUp

滑鼠下一頁鍵： 同Ctrl + PageDown


### 在 textBox2、4 中按下以下鍵：

F2 : 全選/取消全選框裡文字。若原有選取文字則取消選取至其尾端

Ctrl+ 滑鼠左鍵：清除框中所有文字

### 在 textBox2 ：
- 按下 F1 鍵：以找到的字串位置**前**分行分段
- 按下 Pause Break 鍵：以找到的字串位置**後**分行分段
- Ctrl + + （加號，含函數字鍵盤） 或 Ctrl + -（數字鍵盤）  或 Ctrl + 5 (數字鍵盤） ：同 textBox1
- 輸入末綴為「00」的數字可以設定開啟Chrome頁面的等待毫秒時間

### 在 textBox3 ：
- 拖曳網址在 textBox3 或 textBox1 上放開，則會讀入所拖曳的網址值給 textBox3

### textBox4 取代文字用方塊框
- (Ctrl + z 還原功能支援)
- 若有指定要取代的文字，進入後會自動填入之前用以取代過的文字以便輸入（即自動填入對應的預設值）
- 指定要被取代的文字方式：1. 在textBox1中選取文字；2. 若按下 button2 切換成「選取文」（背景紅色）狀態，則將以 textBox2 內的文字為被取代的字串。
- Alt + 1 ：輸入「·」。

### 在表單：
- 點二下滑鼠左鍵，則將剪貼簿的文字內容讀入textBox1中

## 參照資源
（大於 25MB 無法在此上傳的檔案則表列於此。若有疏漏，尚祈提醒末學。感恩感恩　南無阿彌陀佛）
- [查字.mdb](https://www.dropbox.com/s/6vn9hi7i95cbhy4/%E6%9F%A5%E5%AD%97.mdb?dl=0)
- [《重編國語辭典修訂本》資料庫.mdb](https://www.dropbox.com/s/a6t3yhou4smpdv7/%E3%80%8A%E9%87%8D%E7%B7%A8%E5%9C%8B%E8%AA%9E%E8%BE%AD%E5%85%B8%E4%BF%AE%E8%A8%82%E6%9C%AC%E3%80%8B%E8%B3%87%E6%96%99%E5%BA%AB.mdb?dl=0)


