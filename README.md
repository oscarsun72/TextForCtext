# TextForCtext

Text for Ctext 是為了有效加速[《中國哲學書電子化計劃》](https://ctext.org/)[*（Chinese Text Project, 簡稱 CTP 或 ctext）*](https://zh.wikipedia.org/wiki/%E4%B8%AD%E5%9C%8B%E5%93%B2%E5%AD%B8%E6%9B%B8%E9%9B%BB%E5%AD%90%E5%8C%96%E8%A8%88%E5%8A%83) [Wiki（維基）](https://ctext.org/wiki.pl)文本的輸入─尤其是圖文對照頁面─量身訂做的 Windows 應用程式。主體以 C# 寫成，輔以 Word VBA （主要是應付視覺格式化文本）等諸功能。*末學邊大量參與編輯維基區文本邊改寫、增益其功能，自信當是有在參與編輯者，不可或缺的利器。工欲善其事必先利其器，但願多加利用，把吾生也有涯的有限精力用在電腦科技還辦不到的精校解讀詮釋上面* 其中某些功能還可應用在 CTP 外的環境。*如文字編排、取代、自動標點及檢索[《字統網》](https://zi.tools/)(內含[《漢語大字典》](https://homeinmists.ilotus.org/hd/hydzd.php)《異體字字典》[《漢語多功能字庫》](https://humanum.arts.cuhk.edu.hk/Lexis/lexi-mf/)[《全字庫》](https://www.cns11643.gov.tw/index.jsp)《康熙字典》等連結)[《異體字字典》](http://dict.variants.moe.edu.tw/)[《國語辭典》](https://dict.revised.moe.edu.tw/index.jsp)[《漢語大詞典》](https://ivantsoi.myds.me/web/hydcd/search.html)[《康熙字典網上版》](https://www.kangxizidian.com/)、以《易》學關鍵檢索[《漢籍全文資料庫》](https://hanchi.ihp.sinica.edu.tw/ihp/hanji.htm)（可改寫[檢索關鍵字之清單值](https://github.com/oscarsun72/TextForCtext/blob/811364aed5fa4d3e88a0eb96d0ab12660fbb1672/WindowsFormsApp1/Browser.cs#L7898)以滿足特定需求）……等等*。

*以下用DeepL翻譯再略加修訂：*

***Text for Ctext*** *is a Windows application tailored to speed up [Chinese Text Project](https://ctext.org/) [(CTP or ctext)](https://en.wikipedia.org/wiki/Chinese_Text_Project)  [Wiki](https://ctext.org/wiki.pl) text input - especially on the image contrast page. The main body is written in C#, supplemented by Word VBA (mainly for visually formatted text) and other functions. I am confident that it is an indispensable tool for those editing Wiki texts, as I am heavily involved in editing Wiki texts and rewriting this app to improve its functionality. To do a good job, we must first sharpen our tools, but I would like to make more use of our limited energy in computer technology, which is not yet able to do a fine proofreading interpretation of the above. Some of these functions can also be applied to the environment outside the CTP. Such as text arrangement, replacement, automatic punctuation, and retrieval of [“字統網”](https://zi.tools/)(including links to the [“漢語大字典”](https://homeinmists.ilotus.org/hd/hydzd.php), “異體字字典”, [“漢語多功能字庫”](https://humanum.arts.cuhk.edu.hk/Lexis/lexi-mf/), [“全字庫”](https://www.cns11643.gov.tw/index.jsp), “康熙字典” etc.),[ “異體字字典”](http://dict.variants.moe.edu.tw/), [“國語辭典”](https://dict.revised.moe.edu.tw/index.jsp), [“漢語大詞典”](https://ivantsoi.myds.me/web/hydcd/search.html), [“康熙字典網上版”](https://www.kangxizidian.com/), and Searching the [Scripta Sinica database](https://hanchi.ihp.sinica.edu.tw/ihp/hanji.htm) with the Keywords of Yi (the [list of keywords](https://github.com/oscarsun72/TextForCtext/blob/811364aed5fa4d3e88a0eb96d0ab12660fbb1672/WindowsFormsApp1/Browser.cs#L7898) can be rewritten to meet specific needs)...... and so on.──edited from the Translation of DeepL.com (free version)*

> 尤其由[中研院史語所《漢籍電子文獻資料庫》](http://hanchi.ihp.sinica.edu.tw/ihp/hanji.htm)輸入《十三經注疏》、[《維基文庫》](https://zh.wikisource.org/zh-hant/)輸入《四部叢刊》本、[《國學大師》](http://www.guoxuedashi.net/)輸入《四庫全書》本諸書圖文對照時，輔助加速，避免人工之失誤。感恩感恩　南無阿彌陀佛

>> 昨天邊寫程式、測試，邊完成了[《四部叢刊》《南華真經》（《莊子》）](https://ctext.org/wiki.pl?if=gb&chapter=941297#lib77891.114)[第一份文件](https://ctext.org/wiki.pl?if=gb&chapter=941297&action=history)輸入的工作；真是感覺像飛了起來，和之前用手、眼合作判斷分行切割的速度，懸若天壤、判若兩人。感恩感恩　讚歎讚歎　南無阿彌陀佛
20211217在不斷修改增潤的過程中，也將把[此部《莊子》](https://ctext.org/library.pl?if=gb&res=77451)書[維基文本](https://ctext.org/wiki.pl?if=gb&res=393223)建置完畢了。感恩感恩　讚歎讚歎　南無阿彌陀佛 20211218：1951 [建置完畢](https://ctext.org/wiki.pl?if=gb&res=393223) 感恩感恩　南無阿彌陀佛

>> 其他最新進度，詳鄙人此帖： [transferkit IPFS 永遠保存的電子文獻-藏富天下 暨《中國哲學書電子化計劃》愚所輸入完竣之諸本-任真的網路書房-千慮一得齋OnLine-觀死書齋原著及電子化文獻(不屑智慧財產權)歡迎多利用共玉于成](https://oscarsun72.blogspot.com/2022/02/transferkit-ipfs.html) 

> 20240725：配合運用賢超法師[《古籍酷》AI](https://gj.cool/)或[《看典古籍》](https://kandianguji.com/)OCR輸入，將事半功倍也。感恩感恩　讚歎讚歎　南無阿彌陀佛。目前鄙人主要以sl（詳下） 模式在操作，技術已趨成熟穩定可用。阿彌陀佛
 >> 因末學個人使用需要，故《古籍酷》OCR預設為批量授權帳號處理，若無批量授權，請在textBox2中輸入「bF」以關閉之，程式就會改用一般帳號來處理OCR程序（即每日贈予之1000點，約6次OCR額度者）。

\*作業環境、系統需求：**Windows**、.NET 4.7…、Chrome 瀏覽器（Selenium 模式才必要：[chromedriver](https://developer.chrome.com/docs/chromedriver/downloads?hl=zh-tw)） 
 > **不保留任何權利**，歡迎改寫應用到麥金塔(Mac)或 Linux 等作業系統環境中運行

- 本軟件架構為以下三種操作模式（目前本人主要以sl模式在操作）：
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
- [免安裝可執行檔TextForCtextPortable下載，解壓後點擊 TextForCtext.exe 檔案即可](https://github.com/oscarsun72/TextForCtext/blob/master/TextForCtextPortable.zip)：202301052034（2023/1/5 20:34)[直接下載](https://github.com/oscarsun72/TextForCtext/raw/master/TextForCtextPortable.zip)(20240817)
 > [chromedriver下載](https://googlechromelabs.github.io/chrome-for-testing/)（請選擇Windows版：win64或win32看您使用的Chrome瀏覽器是64位元版還是32位元版的，且和您所使用的Chrome瀏覽器版本號相同的版本下載）
 
 > 只要將其中的 chromedriver.exe 放於免安裝版的解壓目錄中（和TextForCtext.exe同一路徑）即可。
  - 以下非 appActivateByName 模式乃適用：
    - 無寫入權限的電腦(如無法安裝Chrome)，請將[GoogleChromePortable](https://portableapps.com/apps/internet/google_chrome_portable)複製到我的文件，並將壓縮檔內的chromedriver.exe移到:
      > C:\Users\(這是使用者登入作業系統的帳號名稱)\Documents\GoogleChromePortable\App\Chrome-bin 目錄下，與「chrome.exe」並置同一資料夾內
    - 末學目前無它電腦可試，以 Selenium 操控 Chrome瀏覽器或許需要其他權限，然而在母校華岡學習雲的公用電腦也可以成功動啟了，若無法開啟，請將您之前打開的Chrome瀏覽器給關閉再啟動本軟件。。若還有問題，請多反饋，仝玉于成。感恩感恩　南無阿彌陀佛

## 介面簡介：
![操作介面](https://github.com/oscarsun72/TextForCtext/blob/master/TextforCtext%E4%BB%8B%E9%9D%A2%E7%B0%A1%E4%BB%8B.png)

textBox1:文本框

textBox2:尋找文本用、與設定配置指令用。
> 若成功下達指令，所輸入之指令字符即會即刻消失。

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
- 操作介面之表單視窗預設為最上層顯示，當表單視窗不在作用中時，只要焦點/插入點不在 textBox2 中，即非最上層顯示；若恢復作用中時（取得焦點時），則最上層顯示。（或：即自動隱藏至系統右下方之系統列/任務列中，當滑鼠滑過任務列中的縮圖ico時，即還原/恢復視窗窗體）
- > 縮小至系統任務列後只要是在圖文對照頁面瀏覽網站，且不是在書圖的第1頁，則會自動開啟簡單修改模式（Quick edit），並將其內容讀入 textBox1 以供備用。
    > 在 SeleniumNew + 手動輸入模式下，若此時按下 Shift 則不會取得文本而是逕行送去《古籍酷》OCR取回文本至textBox1以備用
- > 上述步驟當按下 Ctrl 鍵再執行時雷同，唯此時似乎是要先複製簡單修改模式的網址才行（詳程式碼。可 fork 去改作）。
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
- 預設為最上層顯示，則按下Esc鍵或滑鼠中鍵會隱藏到任務列（系統列）中；滑鼠在其 ico 圖示上滑過即恢復（若在SeleniumNew+手動輸入模式下，瀏覽圖文對照的頁面時，即會開啟該頁簡單修改模式頁面，並將其內容讀入textBox1中。若按住Shift再滿滑過（或未按下時，已切換到OCR輸入模式（ocrTextMode=true） ，則會直接送交賢超法師《古籍酷AI》OCR，若識讀成功，則直接取回文本並加上[查字.mdb](https://www.dropbox.com/s/nbbm2hbneq5g3vx/%E6%9F%A5%E5%AD%97.mdb?dl=0)資料庫已存在的書名篇名號等標點符號。）
- 要清除所選文字，則選取其字，然後在 textBox4 輸入兩個英文半形雙引號 「""」（即表空字串），則不會取代成「""」，而是清除之。⊙或按下 Ctrl + Shift + Delete 組合鍵即可。
- Ctrl + z 還原 textBox1 文本功能。支援打字與取代文字後的還原。還原上限為50次。

- isShortLine() 配合「每行字數判斷用」資料表作為每行字數判斷參考 


## 快速鍵一覽：

### 在表單（操作介面視窗）任何位置按下：

F5 ：重新載入所儲存的文本

Shift + F9 ：重啟小小輸入法

Alt + F9 或 Alt + , : 在《漢籍全文資料庫》或《中國哲學書電子化計劃》中**檢索《易》學關鍵字**

Shift + F10 ： 執行 Word VBA Sub 巨集指令「中國哲學書電子化計劃_只保留正文注文_且注文前後加括弧_貼到古籍酷自動標點」

Alt + F10 、 Alt + F11 ： 將textBox1中選取的文字送去《古籍酷》自動標點。若無選取則將整個textBox1的內容送去。（小於20字元不處理）20240808（臺灣父親節）

Ctrl + F10、 Ctrl + F11： 將textBox1中選取的文字送去《古籍酷》舊版自動標點。若無選取則將整個textBox1的內容送去。（小於20字元不處理）20240808（臺灣父親節）

F12 ： 同 F8 或 Ctrl + Shift + Alt + + 或在非自動且手動輸入模式下，在textBox1 單獨按下數字鍵盤的「+」

Alt + shift + F12  ： 更新最後的備份頁文本

Esc 則按下Esc鍵會隱藏到任務列（系統列）中；滑鼠在其 ico 圖示上滑過即恢復

Ctrl + 1 ：執行 Word VBA Sub 巨集指令「漢籍電子文獻資料庫文本整理_以轉貼到中國哲學書電子化計劃」【 附件即有 [Word VBA](https://github.com/oscarsun72/TextForCtext/tree/master/WordVBA) 相關模組 】

Ctrl + 3 ：執行 Word VBA Sub 巨集指令「漢籍電子文獻資料庫文本整理_十三經注疏」

Ctrl + 4 ：執行 Word VBA Sub 巨集指令「維基文庫四部叢刊本轉來」

Ctrl + f ：移至 textBox2 準備尋找文本

Alt + f ：切換 Fast Mode 不待網頁回應即進行下一頁的貼入動作（即在不須檢覈貼上之文本正確與否，肯定、八成是無誤的，就可以執行此項以加快輸入文本的動作）當是 fast mode 模式時「送出貼上」按鈕會呈現紅綠燈的綠色表示一路直行通行順暢 20230130癸卯年初九第一上班日週一

Alt + r ：在Selenium模式+手動輸入模式下、關閉所在Chrome瀏覽器右側之分頁。（因應《古籍酷》連線不暢所衍生之措施）20231026

Ctrl + n ：開新預設瀏覽器視窗 //原：在新頁籤開啟 google 網頁，以備用（在預設瀏覽器為 Chrome 時）

Ctrl + Shift + n 或 Shift + F1 : 開新Form1 實例

Ctrl + r ：刷新目前 Chrome瀏覽器 或 預設瀏覽器 網頁（同於網頁上按下F5鍵）；當瀏覽器網頁未能完整開啟必須重載時可用。

Ctrl + s 或 Shift + F12：儲存文本至 DropBox 根目錄下的「cText.txt」檔案

Ctrl + w 關閉 Chrome 網頁頁籤

Ctrl + F2 切換語音操作（預設為非 Windows 語音辨識 *Windows Speech Recognition* 操作）

Ctrl + Shift + ` 切換OBS開始串流和停止串流時可處理的程序（這是我於OBS所設定的快捷鍵，可同時觸發）
> 目前是執行 YAKCSwitchr()； YAKC 鍵盤、滑鼠點擊顯示器開關功能。開始串流時即開，關閉時即關閉

Ctrl + Alt + i 顯示IP現狀訊息方塊

Ctrl + Shift + o 執行《看典古籍》OCR API ，執行 GetOCRResult 方法。（須將token存成「OCRAPItoken.txt」檔置於「我的文件\\CtextTempFiles」下，並在程式碼中覆寫本人帳號/郵箱。）

Ctrl + Shift + w 關閉 Chrome 網頁視窗

Ctrl + Shift + \ 切換抬頭平抬格式設定（bool TopLine）

Ctrl + PageUp 或 Alt + 滑鼠滾輪向上：根據 textBox3所載的網址，瀏覽ctext書影的上一頁

Ctrl + PageDown 或 Alt + 滑鼠滾輪向下 ： 根據 textBox3所載的網址，瀏覽ctext書影的下一頁

Ctrl + + （加號，含函數字鍵盤） 或 Ctrl + -（數字鍵盤）  或 Ctrl + 5 (數字鍵盤） 或 Alt + + ：將插入點或選取文字（含）之前的文本剪下貼到 ctext 的[簡單修改模式]框中，並按下「保存編輯」鈕，
        在非自動連續輸入時于瀏覽器新頁籤（預設值，Selenium架構時不會）開啟下一頁準備編輯文本，並回到前一頁籤以供檢視所貼上之文本是否無誤。

 - 在seleniumNew手動輸入模式下，若在貼到簡單修改模式框並翻到下一頁時按住Shift，則會直接將下一頁送給賢超法師《古籍酷AI》OCR 感恩感恩　讚歎讚歎　賢超法師　南無阿彌陀佛        

 - Ctrl + Shift + + （即若按住 Shift ）即留在原頁，不會移至下一頁（若textBox3有指定之網址時）；且同時抓到所在處理頁面，並將Selenium 創建的 driver 的 URL屬性指定到同一頁面或同一網址頁面（如有多分頁網址相同時。機制目前是以網址作判斷故；即類似人工手動指定 driver.Url 的值。） 感恩感恩　讚歎讚歎　南無阿彌陀佛 20230304

Ctrl + Shift + c ：將textBox1的文本複製到剪貼簿 以備用

Ctrl + Alt + o :下載書頁圖片（簡稱「書圖」），交給[Google Keep](https://keep.new/) OCR : 

  複製圖片位址或由textBox3指定要下載的網頁網址即可下載，下載完成後將會自動開啟檔案總管並將該檔案選取。為尊重版權，防止濫用，僅設計一次一頁書圖，以便用即可。知足常樂。感恩感恩　南無阿彌陀佛
  - 下載路徑預設為 Dropbox 根目錄，檔名為 Ctext_Page_Image.png（下載路徑改為 我的文件\CtextTempFiles - 暫時沒有Dropbox同步的需求了，以免頻繁操作OCR時，系統多餘的負擔。）
   - 若不安裝Dropbox者可以自行在其裝路徑裡新增資料夾備置，本軟件許多功能仰賴於此。如備份已輸入之文本及暫存將貼入的文本等等。如我登入Windows作業系統的帳戶名稱為「oscar」，其默認安裝路徑即為：
        C:\Users\oscar\Dropbox。
     - 在非appActivateByName模式下：
       - Google Keep OCR ，模擬使用者手動操作的功能完成。利用其「擷取圖片文字」功能實作。完成後將結果文本複製到剪貼簿，以利貼上。配合快捷鍵 Alt + Insert（將剪貼簿的文字內容讀入textBox1中；在手動輸入鍵入模式下，會自動標出書名號、篇名號）則可直接載入到textBox1中。感恩感恩　讚歎讚歎　南無阿彌陀佛 20230309
       - Alt + Shift + o ：下載書圖並交給[《古籍酷》](https://ocr.gj.cool/)[OCR](https://ocr.gj.cool/try_ocr) 。模擬使用者手動操作，與交給 Google Keep 功能均已全自動化了。感恩感恩　讚歎讚歎　南無阿彌陀佛 20230311
       - Alt + Shift + k ：下載書圖並交給[《看典古籍》](https://kandianguji.com/)[OCR](https://kandianguji.com/ocr) 。模擬使用者手動操作。感恩感恩　讚歎讚歎　南無阿彌陀佛 20240623

Ctrl + Alt + r ：將如《趙城金藏》3欄式的版面書圖《古籍酷》AI服務OCR結果重新排列20240405清明後一日 調用 Rearrangement3ColumnLayout 方法

Alt + ←：視窗向左移動30dpi（+ Ctrl：徵調；插入點在textBox1時例外）//目前在textBox1時照樣

Alt + →：視窗向右移動30dpi（+ Ctrl：徵調；插入點在textBox1時例外）//目前在textBox1時照樣

Alt + ↑：視窗向上移動30dpi（+ Ctrl：徵調；插入點在textBox1時例外）//目前在textBox1時照樣

Alt + ↓：視窗向下移動30dpi（+ Ctrl：徵調；插入點在textBox1時例外）//目前在textBox1時照樣

Alt + 5 （數字鍵盤）：清除標題符碼標記（執行clearTitleMarkCode）

Alt + Shift + d : 下載當前頁面書圖

Alt + Shift + F1 ：切換 textbox1 之字型： 切換支援 CJK - Ext 擴充字集的大字集字型

滑鼠上一頁鍵： 同Ctrl + PageUp

滑鼠下一頁鍵： 同Ctrl + PageDown 
> 如果書圖處有拉出截圖區域，則會自動執行如下輸入截圖模式（滑鼠下一頁鍵 + Ctrl 鍵）

滑鼠下一頁鍵 + Ctrl 鍵： 在需要連續輸入截圖時 ，須先畫出之截圖區域，然後按下Ctrl並按下滑鼠下一頁鍵時，會自動按下頁面中的[Input picture]連結並再按下 Replace page with this data 按鈕

按下 Ctrl 、Alt 或 Shift 任一鍵再啟用表單成為現用的（activate form1)則會啟動自動輸入( auto paste to Quick Edit textbox in Ctext).

按下 Ctrl + * （Multiply）設定為將《四部叢刊》資料庫所複製的文本在表單得到焦點時直接貼到 textBox1 的末尾,或反設定

按下 Ctrl + shift + * （數字鍵盤上的「*」）  切換手動鍵入模式

按下 Ctrl + Shift + - ： 切換OCR輸入模式（開啟/關閉OCR模式；關閉時，若批量處理模式已開啟，亦關閉）
> 若在「是否要自動標識標題，在OCR識讀匯入後」訊息方塊按下確定，則會在讀入OCR識讀結果後自動依 Shift + F8 所指定（或預設）的標題格式（如標題前要空幾格之指定），標上標題語法記號
>> 若不是想要的格式，或自動標識有誤，可以在結果輸出後按下 Ctrl + z 予以還原到原來OCR的結果

按下 Ctrl + / （Divide，數字鍵盤上的） 切換自動連續輸入功能

按下 Ctrl + Shift + / （Divide）  切換 check_the_adjacent_pages 值

Ctrl + Shift + t 同Chrome瀏覽器 --還原最近關閉的頁籤

Ctrl + Shift + p ： 自動翻頁，逐頁瀏覽圖文對照的書圖頁面。要中止，則對本軟件介面或Chrome瀏覽器按住 Ctrl 同時按下滑鼠左鍵。
>> 逐頁瀏覽肉眼檢查是否有空白頁，以免白跑OCR 20240727 執行 CheckBlankPagesBeforeOCR


### 在 textBox1 中按下以下組合鍵：

Insert : 如 MS Word ，切換插入/取代文字模式（hit! 還原機制亦大致成功了。還原上限為50次。Ctrl + z ： 支援打字輸入時及取代文字時的還原。）
    > 新增標點符號不取代功能，以便點校句讀也。至於諸如《》〈〉·「」『』等符號則另有快速鍵方便輸入，也不會取代原有漢字，詳各組合鍵下說明，可多利用。20230118

\` 或 Ctrl + \` ： 於插入點處起至「　」或「􏿽」或「|」或「{」或「<」或分段符號前止之文字加上黑括號【】；若插入點位置前不是「　􏿽」等，則移至該處。如果非插入點，則將選取區前後加上黑括號 (以下不知是什麼，疑是誤貼的文字，待無誤後可刪除： //Print/SysRq 為OS鎖定不能用)

Alt + [ ： 於插入點處起至「　」或「􏿽」或「|」或「{」或「<」或分段符號前止之文字加上**中空黑括號〖〗**；若插入點位置前不是「　􏿽」等，則移至該處。如果非插入點，則將選取區前後加上中空黑括號

Ctrl + Backspace : 清除插入點之前的所有「　」或「􏿽」，若插入點前為「\<p\>」則一併清除

Ctrl + insert ：無選取時則複製插入點後一CJK字長

Ctrl + q 或 Alt + q：據游標（插入點）所在前1段的長度來將textBox1中的文本分行分段

Alt + Shift + q : 據選取區的CJK字長以作分段（末後植入\<p\>，分行則以版式常態值劃分），為非《維基文庫》版式之電子文本，如《寒山子詩集》組詩

Ctrl + \ （反斜線） 或 Alt + \ ： 清除textBox1文本插入點後的分段

按下 F1 鍵：以找到的字串位置**前**分行分段（在文字選取內容接的不是newline時；若是，且選取長度等於常數「predictEndofPageSelectedTextLen」則進行自動貼入 Ctext 的 quit edit 方塊中）
//暫時取消，釋放 F1、 Pause 鍵給 Alt + Shift + 2 用
> 目前按下F1時，若無選取，則複製textBox1的內容，若有選取，則執行 Alt + Shift + 2 功能（poetryFormat函式）

按下 Pause Break 鍵：以找到的字串位置**後**分行分段//暫時取消，釋放  Pause 鍵給 Alt + F7 （原Alt + Shift + 2）

按下 Scroll Lock 將字數較少的行/段落尾末標上分行/段符號（「\<p\>」或「\。<p\>」
> -： 在非自動且手動輸入模式下，在 textBox1 單獨按下數字鍵盤的「-」，執行與按下 Scroll Lock 一樣的功能

F10 : 同上

Ctrl + 6 ：鍵入「{{」

Ctrl + Shift + 6 ：鍵入「}}」(在前面有「{」時，按下 Alt + i 也可以鍵入此值---此似未實作)

Ctrl + F1 ：選取範圍前後加上{{}}

Ctrl + Shift + F1：選取範圍前後加上{{}}並清除分行/段符號

Alt + Shift + 6 或 Alt + s：小注文不換行：notes_a_line()

Alt + Shift + s : 小注文不換行：notes_a_line_all 

Alt + Shift + Ctrl + s : 小注文不換行(短於指定漢字長者 由變數 noteinLineLenLimit 限定）：notes_a_line_all 

Ctrl + Alt + s : 標題下之小注文才不換行( 會與小小輸入法預設的繁簡轉換鍵衝突，使用時請先關閉輸入法。其他快捷鍵若無作用，也多係因有較其優先之如此系統快速鍵已指定的緣故) 20230108

Ctrl + Alt + k 或 Alt + e： 在完整編輯頁面中直接取代文字。請將被取代+取代成之二字前後並置，並將其選取後（或在被取代之文字前放置插入點）再按下此組合鍵以執行直接取代 20240718

Ctrl + 7 ：如同鍵入「。}}」。
> 如於《周易正義》輸入〈彖、象〉辭時適用

Ctrl + 8 ：如同鍵入「　」1個全形空格，且各個空格間有分段符

Ctrl + 9 ：如同鍵入「　　」2個全形空格，且各個空格間有分段符

Ctrl + 0 ：如同鍵入「　　　　」4個全形空格，且各個空格間有分段符

Alt + 1 : 鍵入本站制式留空空格標記「􏿽」：若有選取則取代全形空格「　」為「􏿽」；若已選取「{{」或「}}」則逕以「􏿽」取代

Alt + Shift + 1 如宋詞中的換片空格，只將文中的空格轉成空白，其他如首綴前罝以明段落或標題者不轉換

Alt + 2 : 鍵入全形空格「　」

Alt + Shift + 2 : 將選取區內的「\<p\>」取代為「|」 ，而「　」取代為「􏿽」並清除「*」且將無「|」前綴的分行符號加上「|」（詩偈排版格式用）

Alt + 3 : 鍵入「◯」（原文有大圈界格者，原作〇）

Alt + 4 : 新增【四部叢刊造字對照表】資料並取代其造字,若無選取文字以指定文字，則加以取代

Alt + 6 : 鍵入 「"}}"+ newline +"{{"」

Alt + 7 : 鍵入 「"}}"+ newline +"{{"」

Alt + 8 : 鍵入 「　　*」

Alt + 9 : 鍵入 「 

Alt + 0 : 鍵入 『

Alt + i : 鍵入 》（如 MS Word 自動校正，會依前面的符號作結尾號（close），如前是「〈」，則轉為「〉」……）

Alt + j : 鍵入換行分段符號（newline）（同 Ctrl + j 的系統預設）

Alt + k : 將選取的字詞句及其網址位址送到以下檔案的末後
> C:\Users\oscar\Dropbox\《古籍酷》AI%20OCR%20待改進者隨記%20感恩感恩　讚歎讚歎　南無阿彌陀佛.docx

Alt + n : 將選取的字詞句及其網址位址送到以下檔案的末後
> C:\Users\oscar\Dropbox\《看典古籍》OCR 待改進者隨記 感恩感恩　讚歎讚歎　南無阿彌陀佛

Alt + l : 檢查/輸入抬頭平抬時的條件：執行topLineFactorIuput04condition()

    > 目前只支援新增 condition=0與4 的情形，故名為 04condition，即當後綴是什麼時，此行文字雖短，不是分段，乃是平抬 
    >> 0=後綴；1=前綴；2=前後之前；3前後之後；4是前+後之詞彙；5非前+後之詞彙；6非後綴之詞彙；7非前綴之詞彙

Alt + p 或 Alt + ` : 鍵入 "\<p\>" + newline（分行分段符號）；若置於行/段之首，則會自動移至前一段末再執行

Alt + Shift + p : 鍵入 "。\<p\>" + newline（句號+分行分段符號）；若置於行/段之首，則會自動移至前一段末再執行

Alt + s 或 Alt + Shift + 6 ：小注文不換行 ： notes_a_line()

Alt + Shift + s :  所有小注文都不換行

Alt + t ：預測游標所在行是否為標題（在前無空格縮排時） 執行 detectTitleYetWithoutPreSpace()

Alt + u : 鍵入 《

Alt + v： 檢查[查字.mdb].[異體字反正]資料表中是否已有該字記錄；如果已有資料對應，則閃示橘紅色（表單顏色=Color.Tomato）0.02秒以示警

Alt + y : 鍵入 〈

Alt + . : 鍵入 ·   插入書名、篇名號中間符號

Alt + -（字母區與數字鍵盤的減號） : 如果被選取的是「􏿽」則與下一個「{{」對調；若是「}}」則與「􏿽」對調。（若無選取文字，則自動從插入點往後找「􏿽」或「}}」，直到該行/段末為止。針對《國學大師》《四庫全書》文本小注文誤標而開發）

    > 每頁書圖只檢查一次，只要有嫌疑即暫停，餘請自行檢查 （寫在函式：detectIncorrectBlankAndCurlybrackets_Suspected_aPageaTime()）

Alt + Del : 刪除插入點後第一個分行分段

Alt + Insert ：將剪貼簿的文字內容讀入textBox1中；在手動輸入鍵入模式下，會自動加上書名號、篇名號。

Alt + F1 : 輸入■；若其後為「　」或「􏿽」或「\<p\>」或「*」則清除之。若有選取，則置換選取區中的「　」或「􏿽」或「\<p\>」

Alt + F2 : 輸入□；若其後為「　」或「􏿽」或「\<p\>」或「*」則清除之。若有選取，則置換選取區中的「　」或「􏿽」或「\<p\>」

F1 : 複製textBox1的內容到剪貼簿

F2 : 全選/取消全選框裡文字。若原有選取文字則取消選取至其尾端。20240225元宵後一日：並複製textBox1的內容到剪貼簿

F3 ： 在textBox1 從插入點（游標所在處）開始尋找下一個符合所選取的字串；如果沒有選取，則以 textBox2 的字串為據

Shift + F3 ： 從插入點（游標所在處）開始在textBox1 尋找上一個符合所選取的字串；如果沒有選取，則以 textBox2 的字串為據

F4 ： 重複輸入最後一個輸入的字（字碼）
> 字元、字符，包括特殊字及指令

Shift + F5 ： 在textBox1 回到上1次插入點（游標）所在處（且與最近「charIndexListSize」次瀏覽處作切換，如 MS Word）。charIndexListSize 目前=  3。

F6 : 標題降階（增加標題前之星號）[keysAsteriskPreTitle()]

Alt + F6 或 Alt + F8 或 選取標題文字前之空格再按下 Alt + ` : run autoMarkTitles 自動標識標題（篇名）[autoMarkTitles()]

F7 ： 每行縮排，即每行/段前空一格；全部縮排的機會少，若要全部，則請將插入點放在全文前端或末尾

Shift + F7 : 每行凸排: deleteSpacePreParagraphs_ConvexRow()；全部凸排的機會少，若要全部，則請將插入點放在全文前端或末尾

Alt + F7 : 每行縮排一格後將其末誤標之\<p\>清除:keysSpacePreParagraphs_indent_ClearEnd＿P_Mark；全部縮排的機會少，若要全部，則請將插入點放在全文前端或末尾

Alt + ` ： 加上篇名格式代碼

F8 或 F9 或 F12 或 Ctrl + Alt + + 或數字鍵盤「+」： 整頁貼上Quick edit [簡單修改模式]  並將下一頁直接送交《古籍酷》OCR// 原為加上篇名格式代碼
> 在OCR模式時才會直接送交《古籍酷》OCR。非OCR模式時是送出資料到 Quick edit 並翻到下一頁（已OCR之文本將重新加書名號篇名號等標點。）

Shift + F8 或 Alt + Shift + Pause ： 加上篇名格式代碼並前置N個全形空格.N，預設為2.且可在執行此項時，選取空格數以重設篇名前要空的格數

Alt + Pause 或 當表單在Num Lock關閉時按下數字鍵盤的「5」 ： 自動判斷標題行，加上篇名格式代碼並前置N個全形空格.N，預設為2.且可在執行此項時，選取空格數以重設篇名前要空的格數
    > 此法可與 Alt + t detectTitleYetWithoutPreSpace() 參互應用

F11 : run replaceXdirrectly() 維基文庫等欲直接抽換之字

Ctrl + c ：若無選取，則複製textBox1內的內容

Ctrl + h ：移至 textBox4 準備取代文本文字（若已有取代成的預設值，可以前綴「7」來指定新的取代字串）

Ctrl + K : 依選取文字取得目前URL加該選取字為該頁之關鍵字的連結。如欲在此頁中標出「𢔶」字，即為：
> https://ctext.org/library.pl?if=gb&file=36575&page=53#𢔶

Ctrl + y ： 重做（即復原還原的動作），目前上限為50個記錄

Ctrl + z ： 還原文本，目前上限為50個記錄

Ctrl + F12 ：就 textBox1 所選之字串，查詢《教育部重編國語辭典修訂本》網路版 https://dict.revised.moe.edu.tw/
 > 之前是執行「[查詢國語辭典.exe](https://github.com/oscarsun72/lookupChineseWords.git)」以查詢網路詞典

Alt + F12 查找《異體字字典》。20240817

  > 若在非 appActivateByName 模式下，則但開啟一個分頁以查詢國語辭典耳。唯有在 drive 是 null 時才會執行上述之網路辭典查詢

Alt + g ：就 textBox1 所選之字串，執行「[網路搜尋_元搜尋-同時搜多個引擎.exe](https://github.com/oscarsun72/SearchEnginesConsole.git)」以查詢 Google 等網站
  > 若在非 appActivateByName 模式下，則但開啟一個分頁以檢索Google大神耳。唯有在 drive 是 null 時才會執行上述之搜尋。

Alt + z ：以所選之字（或插入點後之一字）檢索《字統網》 https://zi.tools/
  > 在 appActivateByName 模式下是執行【速檢網路字辭典.exe】

Alt + c ：以所選之詞（不能少於2字）檢索《漢語大詞典》 https://ivantsoi.myds.me/web/hydcd/search.html

Alt + x ：以所選之字（不能不等於1字）檢索《康熙字典網上版 》 https://www.kangxizidian.com/

Ctrl + + （加號，含函數字鍵盤） 或 Ctrl + -（數字鍵盤）  或 Ctrl + 5 (數字鍵盤） 或 Alt + + 或 Alt + a ：將插入點或選取文字（含）之前的文本剪下貼到 ctext 的[簡單修改模式]框中，並按下「保存編輯」鈕，且在[簡單修改模式]下于瀏覽器新頁籤開啟下一頁準備編輯文本，並回到前一頁籤以供檢視所貼上之文本是否無誤。

Ctrl + Alt + + （數字鍵盤加號） ： 同上，唯先將textBox1全選後再執行貼入；即按下此組合鍵則會並不會受插入點所在位置處影響。

Ctrl + Shift + Alt + + 或 Ctrl + Alt + Shift + + （數字鍵盤加號）或只按下「+」鍵（數字鍵盤加號） ： 同上，唯先將textBox1全選後再執行貼入；即按下此組合鍵則會並不會受插入點所在位置處影響。並翻到下一頁直接將它送去《古籍酷》OCR（//欲中止，請按下Ctrl鍵）

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

Ctrl + < 或 Ctrl + , ：到下一個<頭頂(原擬作縮小字型1點然，此功能不常用，擬改用滑鼠方式)

Ctrl + > 或 Ctrl + . ：到下一個>尾端


Ctrl + Shift + Delete ： 將選取文字於文本中全部清除(Ctrl + z 還原功能支援)
> 若是選取《·》〈〉{{}}以執行，則會清除相對應的符號，以便書名號篇名號及注文語法標記之增修。
> 若是選取「\*」或「。\<p\>」則清除「*」或「。\<p\>」（即清除OCR模式下自動標識的標題暨段落符碼
> 若無選取，則清除所有標點符號等（即據以判斷是否已經人為手動編號的條件。）

Ctrl + Delete ： 將插入點所在位置之後的文字一律清除(Ctrl + z 還原功能支援)
> 如果插入點後是空格（space）或空白（􏿽）則清除到非空格空白，否則就一律清除

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
- 輸入「fc」可執行「formatCategory2Columns」函式：以選取範圍為格式化依據，將上下兩欄的目錄/目次內容，從插入點所在位置開始向後格式化（取format,Category二字首，故為fc）執行時若無選取，則以之前的設定為準。若第一次，請務必要選取以供指定。
- 輸入「ws」（wait second）以指定延長等待開啟舊檔對話方塊出現的時間（毫秒數），如「ws1000」即延長1秒；若要縮減時間，請指定負數，如「ws-200」則等待時間再減200毫秒
- 輸入「wO」（wait OCR）以指定等待OCR諸過程最久的時間（以秒數），如「wO60」即最久等到60秒（1分鐘）
   > 由變數 OCR_wait_time_Top_Limit＿second 掌握）
- 輸入「oT」（ocr first ture）設定直接貼入OCR結果先不管版面行款排版模式 PasteOcrResultFisrtMode=true
- 輸入「oF」（ocr first false ）設定直接貼入OCR結果先不管版面行款排版模式 PasteOcrResultFisrtMode=false
- 輸入「bT」（batch processing true ）《古籍酷》OCR批量處理。輸入bT以啟用，輸入bF以停用 BatchProcessingGJcoolOCR=true
- 輸入「bF」（batch processing false ）《古籍酷》OCR批量處理。輸入bT以啟用，輸入bF以停用 BatchProcessingGJcoolOCR=false
- 輸入「mt」（Mute in Processing）則在操作過程中靜音-不撥放音效。MuteProcessing=true。20240315
 > 今天改為切換式的、開關式的。 「mf」依然有效。20240821
- 輸入「mf」（Mute in Processing=false）則在操作過程中撥放音效。MuteProcessing=false。20240315
- 輸入「fm」（form move）切換設定-自動移動表單位置以迴避圖文對照頁面的文本區，以便檢校是否已經編輯過 autoTestPositionAvoidance=true 20240501
- 
- 輸入「x,y」（x、y 為整數以半形逗號間隔，如「835,711」；請打好後用複製貼上的方式來輸入），指定《古籍酷》首頁快速體驗OCR的複製按鈕位置 Copybutton_GjcoolFastExperience_Location的 X 與 Y值

- 輸入「lx」重設《漢籍全文資料庫》或《中國哲學書電子化計劃》**檢索易學關鍵字**清單之索引值為0 即 ListIndex_Hanchi_SearchingKeywordsYijing=0。 

- 在textBox2中輸入開關切換要整頁貼上Quick edit [簡單修改模式]  並將下一頁直接送交去OCR的網站
   - kd：《看典古籍》 （kandianguji）網頁
   - kapi：《看典古籍》api
   - df ：default 古籍酷



### 在 textBox3 網址資訊專用方塊框：
- 拖曳網址在 textBox3 或 textBox1 上放開，則會讀入所拖曳的網址值給 textBox3
- 若已複製網址在剪貼簿，則滑鼠點擊即會讀入所複製的網址值給 textBox3；在軟件介面縮小至任務列時，滑過軟件圖示，也會啟動此功能
- 在非預設模式（ BrowserOPMode.appActivateByName ）模式下，即使用 Selenium 操控 Chrome瀏覽器時，則會自動前往該網址所指向的網頁。


### textBox4 取代文字用方塊框
- (Ctrl + z 還原功能支援)
- 若有指定要取代的文字，進入後會自動填入之前用以取代過的文字以便輸入（即自動填入對應的預設值）
- 指定要被取代的文字方式：1. 在textBox1中選取文字；2. 若按下 button2 切換成「選取文」（背景紅色）狀態，則將以 textBox2 內的文字為被取代的字串。
- Alt + 1 ：輸入「·」。
- 如果在此框輸入的字串前綴半形「@」符號，則會將被取代的字串其對應的用以取代之字串改成目前指定的這個（即在「@」後的字串）20230903蘇拉Saola颱風大菩薩往生後海葵Haikui颱風大菩薩光臨臺灣本島日。感恩感恩　讚歎讚歎　南無阿彌陀佛

### 在表單：
- 點二下滑鼠左鍵，則將剪貼簿的文字內容讀入textBox1中（此暫取消，以免和 textBox1 的點二下衝突 ）
    - 此功能現在好像改成了與 「 Ctrl + - 」相同，即在連續自動輸入模式時重新指定頁尾位置

## 參照資源：
（大於 25MB 無法在此上傳的檔案則表列於此，且均是末學本機Dropbox上自己正在使用的最新檔。若有疏漏，尚祈提醒末學。感恩感恩　南無阿彌陀佛）
- [查字.mdb](https://www.dropbox.com/s/nbbm2hbneq5g3vx/%E6%9F%A5%E5%AD%97.mdb?dl=0)：此為資料檔；資料結構可參看此檔
- [查字forInput.mdb](https://www.dropbox.com/scl/fi/meazmnt9o5pim0xssw5s0/forinput.mdb?rlkey=zdqcgvrvkc3xdz7wnfrjtd3sf&dl=0)：此為使用者介面及程式檔，末學均用此作為輸入的前端介面，將資料回存上一「查字.mdb」檔；即程式碼可參考此檔 202301051619（2023/1/5 16:19）
- [《重編國語辭典修訂本》資料庫.mdb](https://www.dropbox.com/s/dxumn4awnx4e0o9/%E3%80%8A%E9%87%8D%E7%B7%A8%E5%9C%8B%E8%AA%9E%E8%BE%AD%E5%85%B8%E4%BF%AE%E8%A8%82%E6%9C%AC%E3%80%8B%E8%B3%87%E6%96%99%E5%BA%AB.mdb?dl=0)
> 請將以上3檔均複製到 Dropbox 安裝目錄根目錄中，許多功能才能正常執行（若無安裝 Dropbox ，請自己在建立相關路徑，如末學登入Windows的帳號是「**oscar**」，路徑就是：C:\Users\**oscar**\Dropbox 。將以上路徑中的帳號換成您的應該就可以了）。
### Word VBA執行環境配置：
- 安裝 MS Word 32位元
- 在控制台→時鐘、語言和區域→「地區」方塊→「系統管理」頁籤下，「非unicode程式的語言」要「變更系統地區設定」為「中文（繁體，台灣）」
- 將 [TextForCtextPortable.zip](https://github.com/oscarsun72/TextForCtext/raw/master/TextForCtextPortable.zip) 解壓目錄下的 WordVBASeleniumTLB 資料夾中的 TextForCtextWordVBA.dotm 檔案，加入MS Word安裝路逕"%appdata%\Microsoft\Word\STARTUP"即可。
> 相關用得上的 Word VBA 均配置好在這個範本檔案裡，其他設定，請看操作演示。

> 在檔案總管的網址列輸入「%appdata%\Microsoft\Word\STARTUP」再按Enter鍵即可到達此路徑

- 複製一份和本軟件所需相同的「chromedriver.exe」到「chrome.exe」的同一目錄（路徑）下

## 操作演示：
- [TextForCtext簡介展示：以TextForCtext 善用《古籍酷》《看典古籍》OCR暨自動標點《字統網》《異體字字典》《國語辭典》《漢語大詞典》等工具加速輸入《中國哲學書電子化計劃》](https://youtube.com/live/IUzAI5kXkuY?feature=share)
- [我讀《墨子閒詁》文本整理圖文對照程式設計實境秀-自製 TextForCtext 小工具輔助由《漢籍電子文獻資料庫》輸入至《中國哲學書電子化計劃》](https://youtu.be/hnMFTpNfAWg)
- [《法苑珠林》圖文對照錄入實境秀-在《中國哲學書電子化計劃》網站](https://youtu.be/9D9pJKhKx7E)
- [《臨川先生文集》圖文對照錄入實境秀-在《中國哲學書電子化計劃》網站](https://youtu.be/E7iNSZplEC8)
- [《臨川先生文集》圖文對照錄入實境秀續-把她完成-在《中國哲學書電子化計劃》網站](https://youtu.be/Fdb2NUuHCuA)
- [《臨川先生文集》圖文對照錄入實境秀再續-把她完成-在《中國哲學書電子化計劃》網站](https://youtu.be/mO5TUsovwec)
- [《臨川先生文集》圖文對照錄入實境秀-程式設計增益功能-在《中國哲學書電子化計劃》網站](https://youtu.be/THOe_56bknQ)
- [TextForCtext 聲控全自動連續輸入至《中國哲學書電子化計劃》Quick edit Box演示](https://youtu.be/bgsOwh2rEkc)
- [TextForCtext 輸入《中國哲學書電子化計劃》《四庫全書》本《玉海》實境秀](https://youtube.com/live/kwhOSXiNJVs)
- [TextForCtext 輸入《中國哲學書電子化計劃》：使用者手動鍵入演示及實境秀](https://youtube.com/live/f4JlRogZorw)
- [TextForCtext 輸入《中國哲學書電子化計劃》：利用《古籍酷》OCR功能，由使用者手動鍵入之演示及實境秀](https://youtube.com/live/iLxgIiIdXuY?feature=share)
- [TextForCtext 輸入《中國哲學書電子化計劃》：善用賢超法師《古籍酷AI》OCR功能，半自動操作演示曁實境秀：讀錄陳澧《東塾讀書記》褚人穫《堅瓠集》王念孫《讀書雜志》王引之《經義述聞》等](https://www.youtube.com/live/hF-vsdS9kb4?si=Z1YkZT_kVqr_9Y0l)
- [TextForCtext 輸入《中國哲學書電子化計劃》：善用賢超法師《古籍酷AI》OCR 讀錄褚人穫《堅瓠集》實境秀](https://www.youtube.com/live/DSY_jkrUyKc?si=ImWaAAbTcRGIupd2)
- [TextForCtext 輸入《中國哲學書電子化計劃》：善用賢超法師《古籍酷AI》OCR + ProtonVPN 讀錄 錢大昕《廿二史考異》褚人穫《堅瓠集》實境秀](https://youtube.com/live/E1i6cdej_Lk?feature=share)
- [TextForCtext 輸入《中國哲學書電子化計劃》：善用賢超法師《古籍酷AI》OCR + TouchVPN、VPN by Google One 讀錄褚人穫《堅瓠集》完竣及陳澧《東塾讀書記》 實境秀](https://youtube.com/live/IEEX9jbpMIw?feature=share)
- [TextForCtext 輸入《中國哲學書電子化計劃》：善用賢超法師《古籍酷》AI服務「數字萬舟」計劃個人授權帳戶OCR批量處理清儒俞樾《春在堂全書·曲園雜纂》實境秀](https://youtube.com/live/nxKoY8y0OuQ?feature=share)
- [TextForCtext 輸入《中國哲學書電子化計劃》實境秀：善用賢超法師《古籍酷》AI服務OCR及自動標點功能，以清儒文廷式《純常子枝語》示範（操作環境配置）](https://youtube.com/live/wdr8JvSkkhI?feature=share)
- [TextForCtext 輸入《中國哲學書電子化計劃》實境秀：善用賢超法師《古籍酷》AI服務OCR及自動標點功能，以清儒潘平格《潘子求仁錄輯要》示範](https://youtube.com/live/XKcWIo0vHfU?feature=share)
- [TextForCtext 輸入《中國哲學書電子化計劃》實境秀：檢索《易》學關鍵字，善用賢超法師《古籍酷》自動標點，以清儒潘平格《潘子求仁錄輯要》示範](https://www.youtube.com/live/RLzG4AlPe8Q?si=QJVHlcObxOl0OgY2)
- [TextForCtext 輸入《中國哲學書電子化計劃》重點實境秀：蒐集《易》學資料，檢索《易》學關鍵字，以本軟件作為中介工具、善用賢超法師《古籍酷》自動標點，以清儒潘平格《潘子求仁錄輯要》示範](https://youtube.com/live/TyiPkvdUzhg)
- [TextForCtext 輸入《中國哲學書電子化計劃》實境秀：善用賢超法師《古籍酷AI》OCR與自動標點，以清儒文廷式《純常子枝語》示範](https://youtube.com/live/I2Djbck5R6Q?feature=share)
- [以 TextForCtext 軟件善用賢超法師《古籍酷AI》自動標點功能簡要示範（以 kanripo.org 中資料為例）【Word VBA 運行環境的配置】](https://youtube.com/live/2dE0k3_nWi8?feature=share)
- [以 TextForCtext 軟件善用賢超法師《古籍酷AI》自動標點功能簡要示範（以《維基文庫》所收《四庫全書》 為例）](https://youtube.com/live/BEsdf7HXREY?feature=share)
- [TextForCtext 輸入《中國哲學書電子化計劃》：善用《看典古籍》OCR網頁版與API輸入，以清儒文廷式《純常子枝語》簡要示範](https://youtube.com/live/1xQDbnkxA1k?feature=share)
- [以TextForCtext 輸入《中國哲學書電子化計劃》：整理《看典古籍》OCR的結果，以清儒文廷式《純常子枝語》簡要示範](https://youtube.com/live/IK8Wns3UhJk?feature=share)
- [TextForCtext 輸入《中國哲學書電子化計劃》：蒐集《易》學資料，檢索《易》學關鍵字，以本軟件作為介面、善用賢超法師《古籍酷》自動標點及Word VBA，以清儒潘平格《潘子求仁錄輯要》簡要示範](https://youtube.com/live/b1QtGT8bDD0?feature=share)
- [TextForCtext Word VBA環境配置及測試：以清儒潘平格《潘子求仁錄輯要》蒐集其中《易》學資料為簡要示範。《中國哲學書電子化計劃》、賢超法師《古籍酷》自動標點](https://youtube.com/live/zzPAmmw253E)
- [TextForCtext 輸入《中國哲學書電子化計劃》實境秀：善用賢超法師《古籍酷》AI服務OCR（圖像數字化：標注平台、批量處理），以明儒胡震亨《讀書雜錄》示範。任真吟任真曲](https://youtube.com/live/SJmZmtqnSs0)
- [以 TextForCtext 檢索《異體字字典》《國語辭典》及《字統網》。整理《古籍酷》OCR的結果，以輸入《中國哲學書電子化計劃》所收明儒胡震亨《讀書雜錄》重點演示](https://youtube.com/live/Kn51sJ2EhSQ?feature=share)
- [TextForCtext 輸入《中國哲學書電子化計劃》：以賢超法師《古籍酷》OCR 標注平台輸入及整理，以清儒文廷式《純常子枝語》簡要示範](https://youtube.com/live/GM44KOQ54cE?feature=share)
- [TextForCtext 輸入《中國哲學書電子化計劃》：以 VirtualBox 安裝 Windows10 虛擬電腦。以賢超法師《古籍酷》圖像數字化批量處理及《看典古籍》OCR《皇清經解》實境秀](https://youtube.com/live/vUKTJynVES4)
- [TextForCtext 輸入《中國哲學書電子化計劃》：VirtualBox 安裝 Windows10 虛擬電腦、配置運行環境重點演示](https://youtube.com/live/S8IVAc7_VK8?feature=share)
- [TextForCtext 輸入《中國哲學書電子化計劃》：VirtualBox Windows10 虛擬電腦配置 32位元 Word VBA 運行環境重點演示。TextForCtext在新機運行成功](https://youtube.com/live/GI1hfa9CPkc?feature=share)
- [檢查文本中是否闌入版心資訊，以便刪除：文本相似度比對，感恩Copilot大菩薩協作，《字統網》檢索演示，標書名號，音效調適─以TextForCtext 善用《古籍酷》OCR輸入《中國哲學書電子化計劃》](https://youtube.com/live/b0AnDAKYQzA?feature=share)
- []()
- **餘詳此播放清單：** [TextForCtext加速輸入《中國哲學書電子化計劃》(Chinese Text Project) ](https://youtube.com/playlist?list=PLxcUMvfqARSJtLJtRn76Cq4c6-A7WQXta&si=4ezG5WV3yR45UrL1)
- ### 開發秀：
- [TextForCtext 開發檢索《康熙字典網上版》功能實境秀，應用之前的檢索《字統網》《異體字字典》《國語辭典》《漢語大詞典》函式方法](https://youtube.com/live/aFNWSxJi7cs?feature=share)
- []()

