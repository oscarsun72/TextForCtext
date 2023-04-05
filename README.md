# TextForCtext

為了[《中國哲學書電子化計劃》](https://ctext.org/)[（Chinese Text Project,簡稱ctext）](https://zh.wikipedia.org/wiki/%E4%B8%AD%E5%9C%8B%E5%93%B2%E5%AD%B8%E6%9B%B8%E9%9B%BB%E5%AD%90%E5%8C%96%E8%A8%88%E5%8A%83)輸入用

尤其由[中研院史語所《漢籍電子文獻資料庫》](http://hanchi.ihp.sinica.edu.tw/ihp/hanji.htm)輸入《十三經注疏》、[《維基文庫》](https://zh.wikisource.org/zh-hant/)輸入《四部叢刊》本、[《國學大師》](http://www.guoxuedashi.net/)輸入《四庫全書》本諸書圖文對照時，輔助加速，避免人工之失誤。感恩感恩　南無阿彌陀佛

昨天邊寫程式、測試，邊完成了[《四部叢刊》《南華真經》（《莊子》）](https://ctext.org/wiki.pl?if=gb&chapter=941297#lib77891.114)[第一份文件](https://ctext.org/wiki.pl?if=gb&chapter=941297&action=history)輸入的工作；真是感覺像飛了起來，和之前用手、眼合作判斷分行切割的速度，懸若天壤、判若兩人。感恩感恩　讚歎讚歎　南無阿彌陀佛
20211217在不斷修改增潤的過程中，也將把[此部《莊子》](https://ctext.org/library.pl?if=gb&res=77451)書[維基文本](https://ctext.org/wiki.pl?if=gb&res=393223)建置完畢了。感恩感恩　讚歎讚歎　南無阿彌陀佛 20211218：1951 [建置完畢](https://ctext.org/wiki.pl?if=gb&res=393223) 感恩感恩　南無阿彌陀佛

其他最新進度，詳鄙人此帖： [transferkit IPFS 永遠保存的電子文獻-藏富天下 暨《中國哲學書電子化計劃》愚所輸入完竣之諸本-任真的網路書房-千慮一得齋OnLine-觀死書齋原著及電子化文獻(不屑智慧財產權)歡迎多利用共玉于成](https://oscarsun72.blogspot.com/2022/02/transferkit-ipfs.html) 

- 本軟件架構為以下三種操作模式：
  - 在textbox2輸入「ap,」「sl,」「sg,」，可切換瀏覽操作模式設定：
    - ap,=appActivateByName
    - sl,=seleniumNew
    - sg,=seleniumGet
      - 第一種為預設模式，即在現前開啟的Chrome瀏覽器即可操作。（去年（2022）大致完成了）
      - 第二種操作模式是由selenium自動開啟另一個新的Chrome瀏覽器執行體來加以操作。（大致完成了 20230113）
       > 用 [TextForCtextPortable.zip](https://github.com/oscarsun72/TextForCtext/blob/master/TextForCtextPortable.zip) 者 請記得下載與您的Chrome瀏覽器對應的[chromedriver.exe](https://chromedriver.chromium.org/downloads)版本，並和本軟件 TextForCtext.exe 放在同一個目錄/路徑下即可。感恩感恩　南無阿彌陀佛
       
       > ★在全自動連續輸入模式下可配合 Windows 內建的語音辨識軟體 *Windows Speech Recognition* 完全不動手即可操作。快速鍵**Ctrl + F2**可切換此操作，並自動啟動軟體與結束）20230121 23:50壬寅年除夕夜
      - 第三種模式則是混搭前兩種， 或由selenium 取得現用的瀏覽器。來操作。。尚未實作。
  - 要切換三種模式。可在textbox2輸入以上指令。
- [免安裝可執行檔TextForCtextPortable下載，解壓後點擊 TextForCtext.exe 檔案即可](https://github.com/oscarsun72/TextForCtext/blob/master/TextForCtextPortable.zip)：202301052034（2023/1/5 20:34)
  - 以下非 appActivateByName 模式乃適用：
    - 無寫入權限的電腦(如無法安裝Chrome)，請將[GoogleChromePortable](https://portableapps.com/apps/internet/google_chrome_portable)複製到我的文件，並將壓縮檔內的chromedriver.exe移到:
      > C:\Users\(這是使用者登入作業系統的帳號名稱)\Documents\GoogleChromePortable\App\Chrome-bin 目錄下，與「chrome.exe」並置同一資料夾內
    - 末學目前無它電腦可試，以 Selenium 操控 Chrome瀏覽器或許需要其他權限，然而在母校華岡學習雲的公用電腦也可以成功動啟了，若無法開啟，請將您之前打開的Chrome瀏覽器給關閉再啟動本軟件。。若還有問題，請多反饋，仝玉于成。感恩感恩　南無阿彌陀佛

## 介面簡介：
![操作介面](https://github.com/oscarsun72/TextForCtext/blob/master/TextforCtext%E4%BB%8B%E9%9D%A2%E7%B0%A1%E4%BB%8B.png)

textBox1:文本框

textBox2:尋找文本用、與設定配置指令用

textBox3: URL瀏覽參照用（點一下，貼上要瀏覽的網址；在上滑駐滑鼠游標，則顯示提示文字，「現在在第x頁」，以供稽核）

textBox4:文本取代用

button1 「分行分段」或「送出貼上」按鈕: 

 - 直接按下： 預設執行「分行分段」功能。
 
   然切換到自動連續輸入模式（按下 Ctrl + / （數字鍵盤上的） 詳後 ）時，會轉成「送出貼上」 [簡單修改模式]（quick edit文字框）的功能。
   
   若切換到手動輸入模式（按下 Ctrl + Shift + * （數字鍵盤上的）時，會再切換回「分行分段」的功能。因為一般在手動輸入時才有分行分段的必要 20230107
   > 若為鄰近連動編輯模式*（check_the_adjacent_pages=true）*，則顯示為較淺之青色*Aquamarine*，否則為深青色 *DarkCyan*。

 - 若有按下Ctrl才按此鈕則執行[圖文脫鉤 Word VBA](https://github.com/oscarsun72/TextForCtext/blob/4e140975f3881bae8e7f5acc5899c7d61a4794d3/WordVBA/%E4%B8%AD%E5%9C%8B%E5%93%B2%E5%AD%B8%E6%9B%B8%E9%9B%BB%E5%AD%90%E5%8C%96%E8%A8%88%E5%8A%83.bas#L346) ；或在本應用程式介面視窗從作業系統得到焦點、成為作用中的視窗時，只要有按下 Ctrl 或 Shift ，亦會由剪貼簿現存的內容來判斷，是否要執行同一個圖文脫鉤的VBA程序；沒按下的話，預設是執行「中國哲學書電子化計劃.清除頁前的分段符號」程序，如果剪貼簿裡的文字本含有完整「編輯」模式下的文本特徵時（詳程式碼內原理，在[中國哲學書電子化計劃.bas](https://github.com/oscarsun72/TextForCtext/blob/4e140975f3881bae8e7f5acc5899c7d61a4794d3/WordVBA/%E4%B8%AD%E5%9C%8B%E5%93%B2%E5%AD%B8%E6%9B%B8%E9%9B%BB%E5%AD%90%E5%8C%96%E8%A8%88%E5%8A%83.bas#L136)模組檔案裡）。

## 基本功能
（**一切為加速 ctext 網站圖文對照文本編輯而設。目前不免以本人主觀習慣為主**）
- 操作介面之表單視窗預設為最上層顯示，當表單視窗不在作用中時，只要焦點/插入點不在 textBox2 中，即非最上層顯示；若恢復作用中時（取得焦點時），則最上層顯示。（或：即自動隱藏至系統右下方之系統列/任務列中，當滑鼠滑過任務列中的縮圖ico時，即還原/恢復視窗窗體－－此隱藏功能先棄置）
- 當 textBox1、textBox3 內容為空白時，滑鼠左鍵點一下即讀進剪貼簿內容。
- 當離開 textBox2 或 textBox4 文字方塊時，即自執行在 textBox1 尋找或取代文字功能（增罝透過 button2 的切換，決定是否只在選取區中執行尋找、取代）
　
　若沒有找到 textBox2 內的指定字串時，該方塊會顯示紅色半秒，並且插入點還是保留在該方塊中，待繼續輸入其他尋找條件。若要離開，須清空內容
 
  若符合尋找的字串並非獨一無二，則 textBox2 會顯示黃色。
  
  若只有一個符合尋找字串，則 textBox2 會顯示黃綠色，並發出提示音；若再配合按下 F1 鍵（以找到的字串位置前分行分段）、按下 Pause Break 鍵（以找到的字串位置後分行分段）該可加速分行分段
  //暫時取消，釋放 F1、 Pause 鍵給 Alt + Shift + 2 用
  
- 自動依據第一分段字數將 textBox1 插入點其後的文本分段。按下左上方的按鈕或按下 Ctrl + q （參見下文）
- 清除 textBox1 插入點後的分段。按下 Ctrl + \ 
- 當 textBox4 取得焦點（插入點）時自行調整其大小，焦點離開時恢復
- 儲存textbox1文本，（快速鍵 Ctrl + s ： 儲存路徑在 DropBox 的預設安裝路徑的根目錄（C:\Users\ **使用者名** \Dropbox\）中，名「cText.txt」）
- 將 textBox1 插入點前或含選取文字前的文本貼入 ctext [簡單修改模式]框中，並自動按下「保存編輯」鈕，且在 Chrome 瀏覽器新分頁[簡單修改模式]下開啟下一頁準備編輯文本。（參見下文，快速鍵「Ctrl + + 」處。執行此項時，自動在背後進行該次頁面文本的備份，儲存路徑在 DropBox 的預設安裝路徑的根目錄中，名「cTextBK.txt」，是以追加的方式備份。）
　貼至[簡單修改模式]框有 auto連續模式與單一模式，也有移至下一頁或停留在本頁兩種選擇（停留在本頁則加按 Shift 鍵即可），可以 Ctrl + / (數字鍵）切換。我現在常用的是再以 Ctrl + Shift + / （數字鍵） 切換的，Ctext 的鄰頁編輯模式，以確保後來貼上的不會蓋過前頁的。詳各指定鍵（快速鍵）下的說明。
- 自動文本備份及更正備份功能（在複製到剪貼簿時），蓋剛才辛苦做的《四部叢刊》本《南華真經》（《莊子》）第四冊文本，竟然莫名其妙地遺失了，只殘留最後幾頁，整個半天乃至一天的勞動，化為烏有。20211217
- 預設為最上層顯示，則按下Esc鍵或滑鼠中鍵會隱藏到任務列（系統列）中；滑鼠在其 ico 圖示上滑過即恢復
- 要清除所選文字，則選取其字，然後在 textBox4 輸入兩個英文半形雙引號 「""」（即表空字串），則不會取代成「""」，而是清除之。⊙或按下 Ctrl + Shift + Delete 組合鍵即可。
- Ctrl + z 還原 textBox1 文本功能。支援打字與取代文字後的還原。還原上限為50次。

- isShortLine() 配合「每行字數判斷用」資料表作為每行字數判斷參考 


## 快速鍵一覽：

### 在表單（操作介面視窗）任何位置按下：

F5 ：重新載入所儲存的文本

F9 ：重啟小小輸入法

F10 ： 執行 Word VBA Sub 巨集指令「中國哲學書電子化計劃_只保留正文注文_且注文前後加括弧_貼到古籍酷自動標點」

F12  ： 更新最後的備份頁文本

Esc 則按下Esc鍵會隱藏到任務列（系統列）中；滑鼠在其 ico 圖示上滑過即恢復

Ctrl + 1 ：執行 Word VBA Sub 巨集指令「漢籍電子文獻資料庫文本整理_以轉貼到中國哲學書電子化計劃」【 附件即有 [Word VBA](https://github.com/oscarsun72/TextForCtext/tree/master/WordVBA) 相關模組 】

Ctrl + 3 ：執行 Word VBA Sub 巨集指令「漢籍電子文獻資料庫文本整理_十三經注疏」

Ctrl + 4 ：執行 Word VBA Sub 巨集指令「維基文庫四部叢刊本轉來」

Ctrl + f ：移至 textBox2 準備尋找文本

Alt + f ：切換 Fast Mode 不待網頁回應即進行下一頁的貼入動作（即在不須檢覈貼上之文本正確與否，肯定、八成是無誤的，就可以執行此項以加快輸入文本的動作）當是 fast mode 模式時「送出貼上」按鈕會呈現紅綠燈的綠色表示一路直行通行順暢 20230130癸卯年初九第一上班日週一

Ctrl + n ：開新預設瀏覽器視窗 //原：在新頁籤開啟 google 網頁，以備用（在預設瀏覽器為 Chrome 時）

Ctrl + Shift + n 或 Shift + F1 : 開新Form1 實例

Ctrl + r ：刷新目前 Chrome瀏覽器 或 預設瀏覽器 網頁（同於網頁上按下F5鍵）；當瀏覽器網頁未能完整開啟必須重載時可用。

Ctrl + s 或 Shift + F12：儲存文本至 DropBox 根目錄下的「cText.txt」檔案

Ctrl + w 關閉 Chrome 網頁頁籤

Ctrl + F2 切換語音操作（預設為非 Windows 語音辨識 *Windows Speech Recognition* 操作）

Ctrl + Shift + w 關閉 Chrome 網頁視窗

Ctrl + Shift + \ 切換抬頭平抬格式設定（bool TopLine）

Ctrl + PageUp 或 Alt + 滑鼠滾輪向上：根據 textBox3所載的網址，瀏覽ctext書影的上一頁

Ctrl + PageDown 或 Alt + 滑鼠滾輪向下 ： 根據 textBox3所載的網址，瀏覽ctext書影的下一頁

Ctrl + + （加號，含函數字鍵盤） 或 Ctrl + -（數字鍵盤）  或 Ctrl + 5 (數字鍵盤） 或 Alt + + ：將插入點或選取文字（含）之前的文本剪下貼到 ctext 的[簡單修改模式]框中，並按下「保存編輯」鈕，
        在非自動連續輸入時于瀏覽器新頁籤（預設值，Selenium架構時不會）開啟下一頁準備編輯文本，並回到前一頁籤以供檢視所貼上之文本是否無誤。
        
 - Ctrl + Shift + + （即若按住 Shift ）即留在原頁，不會移至下一頁（若textBox3有指定之網址時）；且同時抓到所在處理頁面，並將Selenium 創建的 driver 的 URL屬性指定到同一頁面或同一網址頁面（如有多分頁網址相同時。機制目前是以網址作判斷故；即類似人工手動指定 driver.Url 的值。） 感恩感恩　讚歎讚歎　南無阿彌陀佛 20230304

Ctrl + Shift + c ：將textBox1的文本複製到剪貼簿 以備用

Ctrl + Alt + o :下載書頁圖片（簡稱「書圖」），交給[Google Keep](https://keep.new/) OCR : 

  複製圖片位址或由textBox3指定要下載的網頁網址即可下載，下載完成後將會自動開啟檔案總管並將該檔案選取。為尊重版權，防止濫用，僅設計一次一頁書圖，以便用即可。知足常樂。感恩感恩　南無阿彌陀佛
  - 下載路徑預設為 Dropbox 根目錄，檔名為 Ctext_Page_Image.png
   - 若不安裝Dropbox者可以自行在其裝路徑裡新增資料夾備置，本軟件許多功能仰賴於此。如備份已輸入之文本及暫存將貼入的文本等等。如我登入Windows作業系統的帳戶名稱為「oscar」，其默認安裝路徑即為：
        C:\Users\oscar\Dropbox。
     - 在非appActivateByName模式下：
       - Google Keep OCR ，模擬使用者手動操作的功能完成。利用其「擷取圖片文字」功能實作。完成後將結果文本複製到剪貼簿，以利貼上。配合快捷鍵 Alt + Insert（將剪貼簿的文字內容讀入textBox1中；在手動輸入鍵入模式下，會自動標出書名號、篇名號）則可直接載入到textBox1中。感恩感恩　讚歎讚歎　南無阿彌陀佛 20230309
       - Alt + Shift + o ：下載書圖並交給[《古籍酷》](https://ocr.gj.cool/)[OCR](https://ocr.gj.cool/try_ocr) 。模擬使用者手動操作，與交給 Google Keep 功能均已全自動化了。感恩感恩　讚歎讚歎　南無阿彌陀佛 20230311

Alt + ←：視窗向左移動30dpi（+ Ctrl：徵調；插入點在textBox1時例外）

Alt + →：視窗向右移動30dpi（+ Ctrl：徵調；插入點在textBox1時例外）

Alt + Shift + F1 ：切換 textbox1 之字型： 切換支援 CJK - Ext 擴充字集的大字集字型

滑鼠上一頁鍵： 同Ctrl + PageUp

滑鼠下一頁鍵： 同Ctrl + PageDown

按下 Ctrl 、Alt 或 Shift 任一鍵再啟用表單成為現用的（activate form1)則會啟動自動輸入( auto paste to Quick Edit textbox in Ctext).

按下 Ctrl + * （Multiply）設定為將《四部叢刊》資料庫所複製的文本在表單得到焦點時直接貼到 textBox1 的末尾,或反設定

按下 ctrl + shift + * （數字鍵盤上的「*」）  切換手動鍵入模式

按下 Ctrl + / （Divide，數字鍵盤上的） 切換自動連續輸入功能

按下 Ctrl + Shift + / （Divide）  切換 check_the_adjacent_pages 值

Ctrl + Shift + t 同Chrome瀏覽器 --還原最近關閉的頁籤


### 在 textBox1 中按下以下組合鍵：

Insert : 如 MS Word ，切換插入/取代文字模式（hit! 還原機制亦大致成功了。還原上限為50次。Ctrl + z ： 支援打字輸入時及取代文字時的還原。）
    > 新增標點符號不取代功能，以便點校句讀也。至於諸如《》〈〉·「」『』等符號則另有快速鍵方便輸入，也不會取代原有漢字，詳各組合鍵下說明，可多利用。20230118

` 或 Ctrl + ` ： 於插入點處起至「　」或「􏿽」或「|」或「{」或「<」或分段符號前止之文字加上黑括號【】；若插入點位置前不是「　􏿽」等，則移至該處。如果非插入點，則將選取區前後加上黑括號//Print/SysRq 為OS鎖定不能用

Ctrl + Backspace : 清除插入點之前的所有「　」或「􏿽」，若插入點前為「\<p\>」則一併清除

Ctrl + insert ：無選取時則複製插入點後一CJK字長

Ctrl + q 或 Alt + q：據游標（插入點）所在前1段的長度來將textBox1中的文本分行分段

Alt + Shift + q : 據選取區的CJK字長以作分段（末後植入\<p\>，分行則以版式常態值劃分），為非《維基文庫》版式之電子文本，如《寒山子詩集》組詩

Ctrl + \ （反斜線） 或 Alt + \ ： 清除textBox1文本插入點後的分段

按下 F1 鍵：以找到的字串位置**前**分行分段（在文字選取內容接的不是newline時；若是，且選取長度等於常數「predictEndofPageSelectedTextLen」則進行自動貼入 Ctext 的 quit edit 方塊中）
//暫時取消，釋放 F1、 Pause 鍵給 Alt + Shift + 2 用

按下 Pause Break 鍵：以找到的字串位置**後**分行分段//暫時取消，釋放  Pause 鍵給 Alt + F7 （原Alt + Shift + 2）

按下 Scroll Lock 將字數較少的行/段落尾末標上「\<p\>」符號

Ctrl + 6 ：鍵入「{{」

Ctrl + Shift + 6 ：鍵入「}}」(在前面有「{」時，按下 Alt + i 也可以鍵入此值---此似未實作)

Ctrl + F1 ：選取範圍前後加上{{}}

Alt + Shift + 6 或 Alt + s：小注文不換行：notes_a_line()

Alt + Shift + s : 小注文不換行：notes_a_line_all 

Alt + Shift + Ctrl + s : 小注文不換行(短於指定漢字長者 由變數 noteinLineLenLimit 限定）：notes_a_line_all 

Ctrl + Alt + s : 標題下之小注文才不換行( 會與小小輸入法預設的繁簡轉換鍵衝突，使用時請先關閉輸入法。其他快捷鍵若無作用，也多係因有較其優先之如此系統快速鍵已指定的緣故) 20230108

Ctrl + 7 ：如同鍵入「。}}」，於《周易正義》〈彖、象〉辭時適用

Ctrl + 8 ：如同鍵入「　」1個全形空格，且各個空格間有分段符

Ctrl + 9 ：如同鍵入「　　」2個全形空格，且各個空格間有分段符

Ctrl + 0 ：如同鍵入「　　　　」4個全形空格，且各個空格間有分段符

Alt + 1 : 鍵入本站制式留空空格標記「􏿽」：若有選取則取代全形空格「　」為「􏿽」；若已選取「{{」或「}}」則逕以「􏿽」取代

Alt + Shift + 1 如宋詞中的換片空格，只將文中的空格轉成空白，其他如首綴前罝以明段落或標題者不轉換

Alt + 2 : 鍵入全形空格「　」

Alt + Shift + 2 : 將選取區內的「\<p\>」取代為「|」 ，而「　」取代為「􏿽」並清除「*」且將無「|」前綴的分行符號加上「|」（詩偈排版格式用）

Alt + 3 : 鍵入「〇」（原文有大圈界格者）

Alt + 4 : 新增【四部叢刊造字對照表】資料並取代其造字,若無選取文字以指定文字，則加以取代

Alt + 6 : 鍵入 「"}}"+ newline +"{{"」

Alt + 7 : 鍵入 「"}}"+ newline +"{{"」

Alt + 8 : 鍵入 「　　*」

Alt + 9 : 鍵入 「 

Alt + 0 : 鍵入 『

Alt + i : 鍵入 》（如 MS Word 自動校正，會依前面的符號作結尾號（close），如前是「〈」，則轉為「〉」……）

Alt + j : 鍵入換行分段符號（newline）（同 Ctrl + j 的系統預設）

Alt + l : 檢查/輸入抬頭平抬時的條件：執行topLineFactorIuput04condition()

    > 目前只支援新增 condition=0與4 的情形，故名為 04condition，即當後綴是什麼時，此行文字雖短，不是分段，乃是平抬 
    >> 0=後綴；1=前綴；2=前後之前；3前後之後；4是前+後之詞彙；5非前+後之詞彙；6非後綴之詞彙；7非前綴之詞彙

Alt + p 或 Alt + ` : 鍵入 "\<p\>" + newline（分行分段符號）；若置於行/段之首，則會自動移至前一段末再執行

Alt + s 或 Alt + Shift + 6 ：小注文不換行 ： notes_a_line()

Alt + Shift + s :  所有小注文都不換行

Alt + t ：預測游標所在行是否為標題（在前無空格縮排時） 執行 detectTitleYetWithoutPreSpace()

Alt + u : 鍵入 《

Alt + v： 檢查[查字.mdb].[異體字反正]資料表中是否已有該字記錄；如果已有資料對應，則閃示橘紅色（表單顏色=Color.Tomato）0.02秒以示警

Alt + y : 鍵入 〈

Alt + . : 鍵入 ·

Alt + -（字母區與數字鍵盤的減號） : 如果被選取的是「􏿽」則與下一個「{{」對調；若是「}}」則與「􏿽」對調。（若無選取文字，則自動從插入點往後找「􏿽」或「}}」，直到該行/段末為止。針對《國學大師》《四庫全書》文本小注文誤標而開發）

    > 每頁書圖只檢查一次，只要有嫌疑即暫停，餘請自行檢查 （寫在函式：detectIncorrectBlankAndCurlybrackets_Suspected_aPageaTime()）

Alt + Del : 刪除插入點後第一個分行分段

Alt + Insert ：將剪貼簿的文字內容讀入textBox1中；在手動輸入鍵入模式下，會自動加上書名號、篇名號。

Alt + F1 : 輸入■；若其後為「　」或「􏿽」或「\<p\>」或「*」則清除之。若有選取，則置換選取區中的「　」或「􏿽」或「\<p\>」

Alt + F2 : 輸入□；若其後為「　」或「􏿽」或「\<p\>」或「*」則清除之。若有選取，則置換選取區中的「　」或「􏿽」或「\<p\>」

F2 : 全選/取消全選框裡文字。若原有選取文字則取消選取至其尾端

F3 ： 在textBox1 從插入點（游標所在處）開始尋找下一個符合所選取的字串；如果沒有選取，則以 textBox2 的字串為據

Shift + F3 ： 從插入點（游標所在處）開始在textBox1 尋找上一個符合所選取的字串；如果沒有選取，則以 textBox2 的字串為據

F4 ： 重複輸入最後一個字

Shift + F5 ： 在textBox1 回到上1次插入點（游標）所在處（且與最近「charIndexListSize」次瀏覽處作切換，如 MS Word）。charIndexListSize 目前=  3。

F6 : 標題降階（增加標題前之星號）[keysAsteriskPreTitle()]

Alt + F6 或 Alt + F8 或 選取標題文字前之空格再按下 Alt + ` : run autoMarkTitles 自動標識標題（篇名）[autoMarkTitles()]

F7 ： 每行縮排，即每行/段前空一格；全部縮排的機會少，若要全部，則請將插入點放在全文前端或末尾

Shift + F7 : 每行凸排: deleteSpacePreParagraphs_ConvexRow()；全部凸排的機會少，若要全部，則請將插入點放在全文前端或末尾

Alt + F7 : 每行縮排一格後將其末誤標之\<p\>清除:keysSpacePreParagraphs_indent_ClearEnd＿P_Mark；全部縮排的機會少，若要全部，則請將插入點放在全文前端或末尾

F8 或 Alt + ` ： 加上篇名格式代碼

F11 : run replaceXdirrectly() 維基文庫等欲直接抽換之字

Ctrl + h ：移至 textBox4 準備取代文本文字

Ctrl + F12 ：就 textBox1 所選之字串，執行「[查詢國語辭典.exe](https://github.com/oscarsun72/lookupChineseWords.git)」以查詢網路詞典

Alt + G ：就 textBox1 所選之字串，執行「[網路搜尋_元搜尋-同時搜多個引擎.exe](https://github.com/oscarsun72/SearchEnginesConsole.git)」以查詢 Google 等網站

Ctrl + + （加號，含函數字鍵盤） 或 Ctrl + -（數字鍵盤）  或 Ctrl + 5 (數字鍵盤） 或 Alt + + 或 Alt + a ：將插入點或選取文字（含）之前的文本剪下貼到 ctext 的[簡單修改模式]框中，並按下「保存編輯」鈕，且在[簡單修改模式]下于瀏覽器新頁籤開啟下一頁準備編輯文本，並回到前一頁籤以供檢視所貼上之文本是否無誤。

Ctrl + -（數字鍵盤） 會重設以插入點位置為頁面結束位國,如以滑鼠左鍵點二下

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

滑鼠點二下，執行 Ctrl + + , 將插入點所在之前的文本貼到 Ctext 網頁 [簡單修改模式] 文字方塊中，並會重設以插入點位置為頁面結束位國（同Ctrl + -（數字鍵盤））

滑鼠上一頁鍵： 同Ctrl + PageUp

滑鼠下一頁鍵： 同Ctrl + PageDown

按住 Ctrl 再滑鼠滾輪向上為增大字型，向下滾為縮小字型

按住 Alt  再滑鼠滾輪向上為上一頁（前一頁），向下滾為下一頁（後一頁）

### 在 textBox2、4 中按下以下鍵：

F2 : 全選/取消全選框裡文字。若原有選取文字則取消選取至其尾端

Ctrl+ 滑鼠左鍵：清除框中所有文字

### 在 textBox2 尋找文本及設置指令方塊框：
- 按下 F1 鍵：以找到的字串位置**前**分行分段
- 按下 Pause Break 鍵：以找到的字串位置**後**分行分段
- Ctrl + + （加號，含函數字鍵盤） 或 Ctrl + -（數字鍵盤）  或 Ctrl + 5 (數字鍵盤） ：同 textBox1
- 輸入末綴為「0」的數字可以設定開啟Chrome頁面的等待毫秒時間
- 輸入前綴關鍵字「note:」，可以後綴之數字設定小注不換行的長度限制（byte : 0~255）
- 輸入「msedge」「chrome」「brave」「vivaldi」，可以設定預設瀏覽器名稱
- 輸入「ap,」（或 aa）「sl,」(或 br、bb、ss )「sg,」，可以切換瀏覽操作模式設定：

        ap,=appActivateByName

        sl,=seleniumNew

        sg,=seleniumGet

- 輸入「tS」前綴，設定 Selenium 操控的 Chrome瀏覽器伺服器（ChromeDriverService）的等待秒數（即「new ChromeDriver()」的「TimeSpan」引數值）。預設為 20.5。因昨大年夜 Ctext.org 網頁載入速慢又不穩，因此設置，以防萬一 20230122癸卯年初一 感恩感恩　讚歎讚歎　南無阿彌陀佛(今改為30.5，《古籍酷》OCR頁面所需)
- 輸入「tE」前綴，設定 Selenium 操控的 Chrome瀏覽器中網頁元件的的等待秒數（WebDriverWait。即「new WebDriverWait()」的「TimeSpan」引數值）。預設為 3。
    > 如「tS10」即設定伺服器等候上限是10秒鐘，「tE8」則是設定網頁元件出現的逾時點是8秒鐘
- 輸入「nb,」可以切換 GXDS.SKQSnoteBlank 值以指定是否要檢查注文中因空白而誤標的情形
- 輸入資料夾路徑可指定有效的Chrome瀏覽器的下載位置

### 在 textBox3 網址資訊專用方塊框：
- 拖曳網址在 textBox3 或 textBox1 上放開，則會讀入所拖曳的網址值給 textBox3
- 若已複製網址在剪貼簿，則滑鼠點擊即會讀入所複製的網址值給 textBox3；在軟件介面縮小至任務列時，滑過軟件圖示，也會啟動此功能
- 在非預設模式（ BrowserOPMode.appActivateByName ）模式下，即使用 Selenium 操控 Chrome瀏覽器時，則會自動前往該網址所指向的網頁。


### textBox4 取代文字用方塊框
- (Ctrl + z 還原功能支援)
- 若有指定要取代的文字，進入後會自動填入之前用以取代過的文字以便輸入（即自動填入對應的預設值）
- 指定要被取代的文字方式：1. 在textBox1中選取文字；2. 若按下 button2 切換成「選取文」（背景紅色）狀態，則將以 textBox2 內的文字為被取代的字串。
- Alt + 1 ：輸入「·」。

### 在表單：
- 點二下滑鼠左鍵，則將剪貼簿的文字內容讀入textBox1中（此暫取消，以免和 textBox1 的點二下衝突 ）
    - 此功能現在好像改成了與 「 Ctrl + - 」相同，即在連續自動輸入模式時重新指定頁尾位置

## 參照資源：
（大於 25MB 無法在此上傳的檔案則表列於此，且均是末學本機Dropbox上自己正在使用的最新檔。若有疏漏，尚祈提醒末學。感恩感恩　南無阿彌陀佛）
- [查字.mdb](https://www.dropbox.com/s/nbbm2hbneq5g3vx/%E6%9F%A5%E5%AD%97.mdb?dl=0)：此為資料檔；資料結構可參看此檔
- [查字forInput.mdb](https://www.dropbox.com/s/x4mgb5wd2s5m2gs/%E6%9F%A5%E5%AD%97forInput.mdb?dl=0)：此為使用者介面及程式檔，末學均用此作為輸入的前端介面，將資料回存上一「查字.mdb」檔；即程式碼可參考此檔 202301051619（2023/1/5 16:19）
- [《重編國語辭典修訂本》資料庫.mdb](https://www.dropbox.com/s/dxumn4awnx4e0o9/%E3%80%8A%E9%87%8D%E7%B7%A8%E5%9C%8B%E8%AA%9E%E8%BE%AD%E5%85%B8%E4%BF%AE%E8%A8%82%E6%9C%AC%E3%80%8B%E8%B3%87%E6%96%99%E5%BA%AB.mdb?dl=0)

## 操作演示：
- [我讀《墨子閒詁》文本整理圖文對照程式設計實境秀-自製 TextForCtext 小工具輔助由《漢籍電子文獻資料庫》輸入至《中國哲學書電子化計劃》](https://youtu.be/hnMFTpNfAWg)
- [《法苑珠林》圖文對照錄入實境秀-在《中國哲學書電子化計劃》網站](https://youtu.be/9D9pJKhKx7E)
- [《臨川先生文集》圖文對照錄入實境秀-在《中國哲學書電子化計劃》網站](https://youtu.be/E7iNSZplEC8)
- [《臨川先生文集》圖文對照錄入實境秀續-把她完成-在《中國哲學書電子化計劃》網站](https://youtu.be/Fdb2NUuHCuA)
- [《臨川先生文集》圖文對照錄入實境秀再續-把她完成-在《中國哲學書電子化計劃》網站](https://youtu.be/mO5TUsovwec)
- [《臨川先生文集》圖文對照錄入實境秀-程式設計增益功能-在《中國哲學書電子化計劃》網站](https://youtu.be/THOe_56bknQ)
- [TextForCtext 聲控全自動連續輸入至《中國哲學書電子化計劃》Quick edit Box演示](https://youtu.be/bgsOwh2rEkc)
- [TextForCtext 輸入《中國哲學書電子化計劃》《四庫全書》本《玉海》實境秀](https://youtube.com/live/kwhOSXiNJVs)
- [TextForCtext 輸入《中國哲學書電子化計劃》：使用者手動鍵入演示及實境秀](https://youtube.com/live/f4JlRogZorw)
