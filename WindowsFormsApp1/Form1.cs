// 為WordVBA做API等事宜  https://sl.bing.net/fljln2wR7mK 感恩感恩　Copilot大菩薩 南無阿彌陀佛 20240914
using Microsoft.Win32;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Media;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Policy;
using System.Text;
//using System.Text.Encodings.Web;
using System.Text.RegularExpressions;
using System.Threading;
//using Task = System.Threading.Tasks.Task;
using System.Threading.Tasks;
using System.Web;

//using System.Windows;
using System.Windows.Forms;
//using System.Windows.Input;
using System.Windows.Media.TextFormatting;
using TextForCtext;
using WebSocketSharp;
using static System.Net.Mime.MediaTypeNames;
using static TextForCtext.Browser;
//using static System.Windows.Forms.VisualStyles.VisualStyleElement;
//引用adodb 要將其「內嵌 Interop 類型」（Embed Interop Type）屬性設為false（預設是true）才不會出現以下錯誤：  HResult=0x80131522  Message=無法從組件 載入類型 'ADODB.FieldsToInternalFieldsMarshaler'。
//https://stackoverflow.com/questions/5666265/adodbcould-not-load-type-adodb-fieldstointernalfieldsmarshaler-from-assembly  https://blog.csdn.net/m15188153014/article/details/119895082
using ado = ADODB;//https://docs.microsoft.com/zh-tw/dotnet/csharp/language-reference/keywords/using-directive
using Application = System.Windows.Forms.Application;
using br = TextForCtext.Browser;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;

//using System.Windows.Input;
//using Microsoft.Office.Interop.Word;

namespace WindowsFormsApp1
{

    public partial class Form1 : Form
    {

        /// <summary>
        /// 20240729 Copilot大菩薩：使用《看典古籍》OCR API的C#示例:
        /// 這行程式碼 `private readonly OCRClient _ocrClient = new OCRClient();` 是在類別的層級（class level）上定義的。這意味著 `_ocrClient` 是 `Form1` 類別的一個實例變數（instance variable），它的生命週期與 `Form1` 的實例相同。
        /// 將 `_ocrClient` 定義為實例變數有以下幾個好處：
        /// 1. **可重用性**：在 `Form1` 類別的任何方法中，都可以使用 `_ocrClient`。如果我們在方法內部創建 `OCRClient` 的實例，那麼只能在該方法中使用它。
        ///2. **效能**：由於 `_ocrClient` 在 `Form1` 的生命週期內只被創建一次，所以可以節省創建新實例的開銷。如果我們在每次需要時都在方法內部創建新的 `OCRClient` 實例，可能會浪費資源。
        ///3. **一致性**：如果 `OCRClient` 有任何狀態（state），那麼在 `Form1` 的生命週期內，這些狀態將保持一致。如果我們在每次需要時都在方法內部創建新的 `OCRClient` 實例，那麼每個實例都將有自己的狀態，可能會導致不一致。
        ///希望這個解釋能幫助您理解！如果您有任何其他問題，請隨時向我提問。祝您一切順利！南無阿彌陀佛。
        /// </summary>
        //private readonly OCRClient _ocrClient = new OCRClient();
        private OCRClient _ocrClient = null;
        /// <summary>
        /// 操作過程中靜音
        /// </summary>
        internal static bool MuteProcessing = false;
        /// <summary>
        /// Dropbox路徑（含反斜線）
        /// </summary>
        internal string dropBoxPathIncldBackSlash;
        internal string MydocumentsPathIncldBackSlash;
        readonly System.Drawing.Point textBox4Location; readonly Size textBox4Size;
        readonly System.Drawing.Font textBox4FontDefault;
        private readonly Color textBox2BackColorDefault;
        private readonly Color FormBackColorDefault;
        readonly Size textBox1SizeToForm;
        /// <summary>
        /// 記下主表單Form1的位置
        /// </summary>
        internal Point Form1Pos = new Point();
        /// <summary>
        /// 《古籍酷》OCR批量處理。在textBox2中輸入bT以啟用，輸入bF以停用
        /// </summary>
        internal static bool BatchProcessingGJcoolOCR = false;

        /// <summary>
        /// CJK大字集字型集合（陣列。含CJK 擴充字集者）
        /// </summary>
        //string[] CJKBiggestSet = new string[]{ "HanaMinB", "KaiXinSongB", "TH-Tshyn-P1" };
        readonly string[] CJKBiggestSet = { "全宋體(等寬)", "新細明體-ExtB", "HanaMinB", "KaiXinSongB", "TH-Tshyn-P1", "HanaMinA", "Plangothic P1", "Plangothic P2" };
        readonly Color button2BackColorDefault;

        /// <summary>
        /// 在 Selenium連續輸入時是否為快捷模式，即不檢視貼上結果即進行至下一頁的動作
        /// Alt + f ：切換 Fast Mode 不待網頁回應即進行下一頁的貼入動作（即在不須檢覈貼上之文本正確與否，肯定、八成是無誤的，就可以執行此項以加快輸入文本的動作）當是 fast mode 模式時「送出貼上」按鈕會呈現紅綠燈的綠色表示一路直行通行順暢 20230130癸卯年初九第一上班日週一
        /// </summary>
        private bool fastMode = false;
        /// <summary>
        /// 記下切換 FastModd 前的顏色
        /// </summary>
        Color notFastModeColor;

        /// <summary>
        /// 記下當前頁數頁碼
        /// </summary>
        string _currentPageNum = "";

        /// <summary>
        /// 插入鍵入或取代鍵入模式；取代模式（overwrite mode, overtype mode）時則為false
        /// </summary>
        bool insertMode = true;

        /// <summary>
        /// 鄰近頁連動編輯模式
        /// </summary>
        bool check_the_adjacent_pages = false;

        /// <summary>
        /// 手動輸入模式時為true
        /// </summary>
        bool keyinTextMode = false;
        /// <summary>
        /// 是否進行非正常行長（CheckAbnormalLinePara 方法）的檢查
        /// 由textBox2輸入 "alp" 來切換設定
        /// 預設為 true
        /// </summary>
        bool abnormalLineParaChecking = true;

        /// <summary>
        /// OCR輸入模式時為true（直接連續輸入OCR結果，先不校讀）
        /// </summary>
        bool ocrTextMode = false;
        /// <summary>
        /// 自動移動表單位置以迴避圖文對照頁面的文本區，以便檢校是否已經編輯過 20240501
        /// 在textBox2中以「fm」（form moving）切換設定
        /// </summary>
        bool autoTestPositionAvoidance = false;
        /// <summary>
        /// 直接貼入OCR結果，先不管版面行款排版、及是否有編輯標記
        /// </summary>
        internal bool PasteOcrResultFisrtMode = false;
        /// <summary>
        /// 在textBox2中輸入開關切換要整頁貼上Quick edit [簡單修改模式]  並將下一頁直接送交去OCR的網站
        /// kd：《看典古籍》 （kandianguji)
        /// kapi：《看典古籍》api
        /// df ：default 古籍酷
        /// </summary>
        internal br.OCRSiteTitle PagePast2OCRsite = br.OCRSiteTitle.GJcool;
        /// <summary>
        /// 指定是否要在OCR讀入後自動標識標題語法標記
        /// </summary>
        bool autoTitleMark_OCRTextMode = false;

        /// <summary>
        /// 原文有抬頭平抬格式
        /// </summary>
        bool TopLine = false;

        /// <summary>
        /// 現行行是否屬縮排；或書內是否含有縮排格式
        /// </summary>
        bool Indents = false;//原來不知道怎麼預設為true,待觀察,若無誤，即保留新設定之預設值 20231102
        //bool Indents = true;
        internal string TextBox1_Text
        {
            get { return textBox1.Text; }
            set { textBox1.Text = value; }
        }
        internal int TextBox1_SelectionStart
        {
            get { return textBox1.SelectionStart; }
            set { textBox1.SelectionStart = value; }
        }
        internal int TextBox1_SelectionLength
        {
            get { return textBox1.SelectionLength; }
            set { textBox1.SelectionLength = value; }
        }
        internal string textBox3Text
        {
            get { return textBox3.Text; }
            set { textBox3.Text = value; }
        }
        internal string textBox4Text
        {
            get { return textBox4.Text; }
            set { textBox4.Text = value; }
        }
        internal Font textBox4Font
        {
            get { return textBox4.Font; }
            set { textBox4.Font = value; }
        }
        //取得輸入模式：手動或自動
        internal bool KeyinTextMode { get { return keyinTextMode; } }

        internal string CurrentPageNum { get { return _currentPageNum; } }
        //static internal string mainFromTextBox3Text;

        /// <summary>
        /// 軟體進行時的架構
        /// </summary>
        /*browser operation mode:
             appActivateByName:本來原始的；網路學來的
                selenium 純selenium模式，啟動新的 chrome 執行個體，且須登入；chatGPT 教的
                seleniumGet 混合模式，且夫不啟動 chrome，而是取得已經運動的chrome的執行個體的； chatGPT 教的+之前網路學的
        */
        public enum BrowserOPMode { appActivateByName, seleniumNew, seleniumGet };

        /// <summary>
        /// 現行軟體運行之架構是哪種（appActivateByName, seleniumNew, seleniumGet……）
        /// 每表獨立的表單可指定不同的模式，故不能是 static；有空再來改
        /// </summary>
        internal static BrowserOPMode browsrOPMode = BrowserOPMode.appActivateByName;


        /// <summary>
        /// 隱藏到系統列用物件
        /// </summary>
        readonly System.Windows.Forms.NotifyIcon ntfyICo;

        int thisHeight, thisWidth, thisLeft, thisTop;
        [DllImport("user32.dll")]
        static extern bool CreateCaret(IntPtr hWnd, IntPtr hBitmap, int nWidth, int nHeight);
        // 引入 user32.dll 來使用 GetKeyState 方法
        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern short GetKeyState(int keyCode);

        private const int VK_CAPITAL = 0x14;
        [DllImport("user32.dll")]
        static extern bool ShowCaret(IntPtr hWnd);

        public Document _document;
        /// <summary>
        /// 作為主表單Form1的代理實例。20241003 Copilot大菩薩：C# Windows.Forms 屬性讀取：https://sl.bing.net/hPdFVUz4788 依賴注入（Dependency Injection, DI），使用依賴注入來管理實例
        /// </summary>
        public static Form1 InstanceForm1 { get; private set; }

        public Form1()
        {

            InitializeComponent();

            //設定屬性
            textBox1FontDefaultSize = textBox1.Font.Size;
            textBox4Location = textBox4.Location;
            textBox4Size = textBox4.Size;
            textBox4FontDefault = textBox4.Font;
            textBox1SizeToForm = new Size(this.Width - textBox1.Width, this.Height - textBox1.Height);
            MydocumentsPathIncldBackSlash = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\";
            dropBoxPathIncldBackSlash = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Dropbox\";
            dropBoxPathIncldBackSlash = Directory.Exists(dropBoxPathIncldBackSlash) ? dropBoxPathIncldBackSlash : dropBoxPathIncldBackSlash.Replace(@"C:\", @"A:\");
            FormBackColorDefault = this.BackColor;
            button2BackColorDefault = button2.BackColor;
            textBox2BackColorDefault = textBox2.BackColor;
            defaultBrowserName = GetWebBrowserName();
            var cjk = getCJKExtFontInstalled(CJKBiggestSet[FontFamilyNowIndex]);
            if (cjk != null)
            {
                if (cjk.Name == "KaiXinSongB")
                {
                    textBox1.Font = new Font(cjk, (float)17);
                }
                else
                {
                    textBox1.Font = new Font(cjk, textBox1FontDefaultSize);
                }
                textBox2.Font = new Font(cjk, textBox2.Font.Size);
                textBox4.Font = new Font(cjk, textBox4.Font.Size);
            }
            thisHeight = this.Height; thisWidth = this.Width; thisLeft = this.Left; thisTop = this.Top;
            //加入事件處理常式
            this.ntfyICo = new NotifyIcon
            {
                Icon = this.Icon
            };
            this.ntfyICo.MouseClick += new System.Windows.Forms.MouseEventHandler(nICo_MouseClick);
            //this.ntfyICo.MouseClick += new System.Windows.Forms.MouseEventHandler(nICo_MouseMove);
            this.ntfyICo.MouseMove += new System.Windows.Forms.MouseEventHandler(nICo_MouseMove);
            //this.Shown += Form1_Shown;//https://stackoverflow.com/questions/32720207/change-caret-cursor-in-textbox-in-c-sharp

            this.FormClosing += Form1_FormClosing;//202301050101 creedit
            textBox3.MouseMove += textBox3_MouseMove;
            textBox1.MouseWheel += new MouseEventHandler(textBox1_MouseWheel);

            if (Application.OpenForms.Count == 0)
                InstanceForm1 = this;

            _document = new Document(ref textBox1);
        }

        /// <summary>
        /// 調整textBox1的字形大小、切換上一頁下一頁等功能
        /// 按住 Ctrl 再滑鼠滾輪向上為增大字型，向下滾為縮小字型
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox1_MouseWheel(object sender, MouseEventArgs e)
        {
            bool bold = textBox1.Font.Bold;
            switch (ModifierKeys)
            {
                case Keys.Control:
                    if (e.Delta > 0)
                    {
                        // 滾輪向上，增大字型
                        textBox1.Font = new Font(textBox1.Font.FontFamily, textBox1.Font.Size + 1);
                    }
                    else
                    {
                        // 滾輪向下，縮小字型
                        textBox1.Font = new Font(textBox1.Font.FontFamily, textBox1.Font.Size - 1);
                    }
                    break;
                case Keys.Alt:
                    if (e.Delta > 0)
                    {
                        // 滾輪向上，上一頁
                        nextPages(Keys.PageUp, true);
                        if (autoPaste2QuickEdit || keyinTextMode) AvailableInUseBothKeysMouse();
                    }
                    else
                    {
                        // 滾輪向下，下一頁
                        nextPages(Keys.PageDown, true);
                        if (autoPaste2QuickEdit || keyinTextMode) AvailableInUseBothKeysMouse();
                    }
                    break;
            }
            if (bold)
                textBox1.Font = new Font(textBox1.Font.FontFamily, textBox1.Font.Size, FontStyle.Bold);
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {

            //終止由 chromedriver.exe 程序開啟的Chrome瀏覽器,釋放系統記憶體
            //new Task(Action ).Wait(4500);
            if (Name == "Form1" && (br.driver != null || browsrOPMode != BrowserOPMode.appActivateByName))
            {
                //if (MessageBox.Show("本軟件即將關閉，也會同時關閉由其開啟的Chrome瀏覽器，若有沒儲存的資訊，請先儲存再按「確定」鈕繼續；否則請按「取消」", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.Cancel) { e.Cancel = true; return; }
                if (MessageBoxShowOKCancelExclamationDefaultDesktopOnly("本軟件即將關閉，也會同時關閉由其開啟的Chrome瀏覽器，若有沒儲存的資訊，請先儲存再按「確定」鈕繼續；否則請按「取消」") == DialogResult.Cancel) { e.Cancel = true; return; }
                else
                {
                    try
                    {
                        if (br.ImproveGJcoolOCRMemoDoc != null)
                        {
                            //ImproveGJcoolOCRMemoDoc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);
                            ImproveGJcoolOCRMemoDoc.Application.Quit();
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.HResult + ex.Message);
                        //Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                    }

                    if (MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否關閉Chrome瀏覽器？") == DialogResult.OK)
                    {
                        Task.Run(() =>
                    {
                        if (br.driver != null)
                        {
                            try
                            {

                                //Task.WaitAny();
                                br.driver.Quit();
                                //chatGPT：不過，建議使用 Quit() 方法來關閉 WebDriver 實例並釋放所有資源，因為它會同時處理 Close() 和 Dispose() 方法20230108
                                //br.driver.Close();
                                //br.driver.Dispose();
                            }
                            catch (Exception)
                            {
                                br.driver = null;
                                //throw;
                            }

                        }
                    });
                        Task.WaitAll();
                    }
                    else
                        br.driver = null;
                    //終止 chromedriver.exe 程序,釋放系統記憶體
                    //Process[] processes = Process.GetProcessesByName("chromedriver");
                    //foreach (Process process in processes)
                    //{
                    //    process.Kill();
                    //}
                    //br.killProcesses(new string[] { "chromedriver" });
                    br.killchromedriverFromHere();
                    Thread.Sleep(850);
                    if (br.getChromedrivers().Length > 0)
                        if (MessageBox.Show("還有chromedriver.exe程序在運行，是否全部清除？", "chromedrivers still there"
                            , MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                            == DialogResult.OK)
                            br.killProcesses(new string[] { "chromedriver" });
                }
            }

            //有上式便不用以下了
            //try
            //{
            //    //釋放應用程式佔用的記憶體
            //    br.driver.Close();//202301051447(2013/1/5 14:47) creedit

            //}
            //catch (OpenQA.Selenium.WebDriverException ex)
            //{
            //    bool v = ex.HResult == -2146233088;//先手動關了 chromedriver.exe 時
            //    if (v) { }
            //    else
            //    {
            //        v = ex.HResult == -2146233036;
            //        if (v) { }
            //        else throw;
            //    }

            //}
        }

        /// <summary>
        /// 主表單是否在作用中（有系統焦點）
        /// </summary>
        internal bool Active
        {//20230114 creedit chatGPT：Windows Forms Active Detection
            get
            {
                if (this.WindowState == FormWindowState.Minimized || !this.Visible)
                    return false;
                if (this.Visible && this.WindowState != FormWindowState.Minimized)
                {
                    bool focused = false;
                    foreach (Control item in this.Controls)
                    {
                        if (item.Focused) { focused = true; break; }
                    }

                    if (focused) return focused & this.Focused;
                    else return this.Focused;//textBox1.Focused=true時這個會是false
                }
                else
                    return this.Focused;
            }
        }

        /// <summary>
        /// 游標寬度恢復（在插入輸入模式時）
        /// </summary>
        /// <param name="ctl"></param>
        void Caret_Shown(Control ctl)
        {
            CreateCaret(ctl.Handle, IntPtr.Zero, 4, Convert.ToInt32(ctl.Font.Size * 1.5));
            ShowCaret(ctl.Handle);
        }
        /// <summary>
        /// 游標寬度加寬（在取代輸入模式時）
        /// </summary>
        /// <param name="ctl"></param>
        /* 20230326 Bing大菩薩：根據你提供的代碼，你正在使用 `CreateCaret` 和 `ShowCaret` 函數來更改游標的寬度。這些函數是 Windows API 函數，用於創建和顯示游標。
            有幾個原因可能會導致你的代碼無法正確改變游標寬度。例如，如果 `CreateCaret` 函數返回非零值，則表示創建游標失敗。此外，如果 `ShowCaret` 函數返回非零值，則表示顯示游標失敗。
            你可以嘗試檢查這些函數的返回值，以確定是否存在錯誤。此外，你也可以嘗試使用 `GetLastError` 函數來獲取更多關於錯誤的信息。
            希望這對你有幫助！
         */
        void Caret_Shown_OverTypeMode(Control ctl)
        {
            CreateCaret(ctl.Handle, IntPtr.Zero, Convert.ToInt32(ctl.Font.Size * 1.5), Convert.ToInt32(ctl.Font.Size * 1.5));
            ShowCaret(ctl.Handle);
        }
        //void Form1_Shown(object sender, EventArgs e)
        //{//插入點游標寬廣設定.文字框中更改插入符號(游標) https://www.796t.com/post/OXZ0cjQ=.html 如何改變鼠標的樣式: https://blog.xuite.net/akira32/home/109034425-Visual+C%23,%E5%A6%82%E4%BD%95%E6%94%B9%E8%AE%8A%E9%BC%A0%E6%A8%99%E7%9A%84%E6%A8%A3%E5%BC%8F+(change+cursor+to+system+type+or+customed+type) https://www.google.com/search?q=%E6%88%91%E7%94%A8%E4%BB%A5%E4%B8%8A%E7%9A%84%E4%BD%9C%E6%B3%95%E9%83%BD%E7%84%A1%E6%B3%95%E6%94%B9%E8%AE%8A%E9%BC%A0%E6%A8%99%E7%9A%84%E6%A8%A3%E5%BC%8F%2C%E8%AB%8B%E5%95%8F%E4%B8%80%E4%B8%8B%E8%A6%81%E5%A6%82%E4%BD%95%E9%81%94%E6%88%90%E5%91%A2%3F&rlz=1C1GCEU_zh-TWTW823TW823&sourceid=chrome&ie=UTF-8
        //    Caret_Shown(textBox1);
        //    //CreateCaret(textBox1.Handle, IntPtr.Zero, 5, int.Parse(textBox1.Font.SizeInPoints.ToString()));
        //    //ShowCaret(textBox1.Handle);            
        //}
        /* 20230326 那請問在上述我給您看的程式碼中，如何避免錯誤出現呢？錯誤又在怎樣的操作實況下會造成？Bing大菩薩：
         在你提供的代碼中，你使用了 `CreateCaret` 和 `ShowCaret` 函數來創建和顯示游標。這些函數都有可能會失敗，導致錯誤發生。
            為了避免錯誤發生，你可以在調用這些函數後檢查它們的返回值。如果返回值為非零，則表示函數調用失敗。此時，你可以使用 `GetLastError` 函數來獲取更多關於錯誤的信息。
            下面是一個簡單的示例代碼，演示如何檢查 `CreateCaret` 和 `ShowCaret` 函數的返回值，並在發生錯誤時使用 `GetLastError` 函數獲取錯誤信息：
            ```csharp
            [DllImport("kernel32.dll")]
            static extern uint GetLastError();

            void Caret_Shown(Control ctl)
            {
                if (!CreateCaret(ctl.Handle, IntPtr.Zero, 4, Convert.ToInt32(ctl.Font.Size * 1.5)))
                {
                    uint errorCode = GetLastError();
                    // 處理錯誤
                }
                if (!ShowCaret(ctl.Handle))
                {
                    uint errorCode = GetLastError();
                    // 處理錯誤
                }
            }
            ```
            在上面的代碼中，我們在調用 `CreateCaret` 和 `ShowCaret` 函數後檢查了它們的返回值。如果返回值為非零，則表示函數調用失敗。此時，我們使用 `GetLastError` 函數來獲取錯誤碼，並進行相應的錯誤處理。
            至於錯誤會在怎樣的操作實況下造成，這取決於許多因素。例如，如果系統資源不足，則可能會導致 `CreateCaret` 函數失敗。此外，如果控件的句柄無效，則也可能會導致 `CreateCaret` 和 `ShowCaret` 函數失敗。
            希望這對你有幫助！
         */


        /* 20230326 C# System.Windows.Forms textBox 如何更改游標顏色？ Bing大菩薩：
         在 `System.Windows.Forms` 中，`TextBox` 控件並沒有提供直接更改游標顏色的方法。不過，你可以通過創建自定義的 `TextBox` 控件來實現這個功能。
            你可以重寫 `TextBox` 控件的 `WndProc` 方法，並在其中攔截 `WM_SETCURSOR` 消息。然後，你可以使用 `CreateCaret` 函數創建一個新的游標，並使用 `SetCaretBlinkTime` 函數設置游標的閃爍時間。最後，你可以使用 `ShowCaret` 函數顯示新創建的游標。
            下面是一個簡單的示例代碼，演示如何創建一個自定義的 `TextBox` 控件，並在其中更改游標顏色：
            ```csharp
            using System;
            using System.Drawing;
            using System.Runtime.InteropServices;
            using System.Windows.Forms;

            public class MyTextBox : TextBox
            {
                private const int WM_SETCURSOR = 0x20;

                [DllImport("user32.dll")]
                private static extern IntPtr CreateCaret(IntPtr hWnd, IntPtr hBitmap, int nWidth, int nHeight);

                [DllImport("user32.dll")]
                private static extern bool ShowCaret(IntPtr hWnd);

                [DllImport("user32.dll")]
                private static extern bool SetCaretBlinkTime(uint uMSeconds);

                protected override void WndProc(ref Message m)
                {
                    if (m.Msg == WM_SETCURSOR)
                    {
                        CreateCaret(this.Handle, IntPtr.Zero, 4, this.Height);
                        SetCaretBlinkTime(500);
                        ShowCaret(this.Handle);
                        return;
                    }

                    base.WndProc(ref m);
                }
            }
            ```
            在上面的代碼中，我們創建了一個名為 `MyTextBox` 的自定義控件，它繼承自 `TextBox` 控件。在 `WndProc` 方法中，我們攔截了 `WM_SETCURSOR` 消息，並使用 `CreateCaret`、`SetCaretBlinkTime` 和 `ShowCaret` 函數來創建和顯示新的游標。
            希望這對你有幫助！
         */

        /// <summary>
        /// 還原已隱藏到系統列的物件
        /// </summary>
        void show_nICo(Keys modifierKeys)
        {
            try
            {

                this.Show();
                if (ntfyICo.Visible) ntfyICo.Visible = false;
                this.WindowState = FormWindowState.Normal;
                this.Height = thisHeight;
                this.Width = thisWidth;
                this.Left = thisLeft;
                this.Top = thisTop;
                //手動編輯模式時：
                //if (!autoPaste2QuickEdit && keyinTextMode)//今連續輸入模式也需要 20250301
                if (autoPaste2QuickEdit && keyinTextMode)
                {
                    string xClp = string.Empty;
                    try
                    {
                        xClp = Clipboard.GetText();
                    }
                    catch (Exception)
                    {
                    }
                    if (modifierKeys != Keys.Control && xClp.StartsWith("http") &&//xClp != "" &&
                        xClp.Length > "https://ctext.org/".Length + "#editor".Length
                        && xClp.Substring(0, "https://ctext.org/".Length) == "https://ctext.org/"
                        && !ocrTextMode)
                    {
                        string url = xClp;
                        textBox3_Click(new object(), new MouseEventArgs(MouseButtons.Left, 0, 0, 0, 0));//textBox3_MouseMove(new object(), new MouseEventArgs(MouseButtons.Left, 0, 0, 0, 0));
                        if (browsrOPMode != BrowserOPMode.appActivateByName)
                        {
                            br.driver = br.driver ?? br.DriverNew();
                            try
                            {
                                br.GoToUrlandActivate(url, keyinTextMode);
                            }
                            catch (Exception ex)
                            {
                                switch (ex.HResult)
                                {
                                    default:
                                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                                        break;
                                }
                            }
                            if (xClp.IndexOf("edit") > -1 && xClp.IndexOf("&page") > -1)//xClp.Substring(xClp.LastIndexOf("#editor")) == "#editor")
                                                                                        //url此頁的Quick edit值傳到textBox1,並存入剪貼簿以備用
                            {
                                //若此時按下 Shift 則不會取得文本而是逕行送去《古籍酷》OCR取回文本至textBox1以備用
                                if (modifierKeys == Keys.Shift && !PagePaste2GjcoolOCR_ing)
                                    toOCR(br.OCRSiteTitle.GJcool);
                                else
                                {
                                    //if (keyinTextMode)
                                    //{
                                    //全選文字方塊內容以備貼入
                                    //ie.SendKeys(OpenQA.Selenium.Keys.Control + "a");
                                    //br.SelectAllQuickedit_data_textboxContent();//20240913作廢（現已可以由value屬性設定了，儘量不用剪貼簿！）
                                    //}
                                    //string text = br.CopyQuickedit_data_textboxText();
                                    string text = br.Quickedit_data_textboxTxt;
                                    CnText.BooksPunctuation(ref text, false);
                                    undoRecord();
                                    textBox1.Text = text;
                                    //if (!OcrTextMode)
                                    //clearBracketsInsidePairsBrackets();


                                    //20240913先作廢看看
                                    //if (Clipboard.GetText() != text && text != "")//CopyQuickedit_data_textboxText已用到等價 SetText 的方法了
                                    //    Clipboard.SetText(text);
                                }

                                //改到後面呼叫
                                //if (!Active)
                                //{
                                //    AvailableInUseBothKeysMouse();
                                //}

                                //避免剪貼簿內還殘留上一次用過的網址
                                xClp = Clipboard.GetText();
                                if (xClp.IndexOf("edit") > -1 && xClp.IndexOf("&page") > -1) Clipboard.Clear();
                            }

                        }
                        else//if (browsrOPMode == BrowserOPMode.appActivateByName)
                        { Process.Start(url); appActivateByName(); }
                        //Clipboard.Clear();

                        //bringBackMousePosFrmCenter();
                        //if (!Active)
                        //{
                        AvailableInUseBothKeysMouse();
                        //}
                    }
                    //若此時按下 Shift 則不會取得文本而是逕行送去《古籍酷》OCR取回文本至textBox1以備用
                    else if ((modifierKeys == Keys.Shift || (ocrTextMode && modifierKeys != Keys.Control))
                        && !PagePaste2GjcoolOCR_ing && browsrOPMode != BrowserOPMode.appActivateByName)
                    {
                        if (modifierKeys == Keys.Shift) ocrTextMode = true;
                        br.GoToCurrentUserActivateTab();
                        string brUrl = br.GetDriverUrl;//.driver.Url;
                        if (brUrl.IndexOf("ctext.org") > -1 && brUrl.IndexOf("&file=") > -1 && brUrl.IndexOf("&page=") > -1)
                        {
                            if (brUrl.IndexOf("edit") == -1)
                            {
                                brUrl = br.GetQuickeditUrl();
                                if (brUrl != string.Empty)
                                    br.driver.Navigate().GoToUrl(brUrl);
                            }
                            if (brUrl != string.Empty)
                            {
                                if (textBox3.Text != brUrl) textBox3.Text = brUrl;
                                Form1.playSound(Form1.soundLike.press);
                                toOCR(br.OCRSiteTitle.GJcool);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //Debugger.Break();
                Console.WriteLine(ex.HResult + ex.Message);
            }
        }
        /// <summary>
        /// 在已隱藏到系統列的物件圖示上點一下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void nICo_MouseClick(object sender, MouseEventArgs e)
        {
            show_nICo(ModifierKeys);
        }
        /// <summary>
        /// 在已隱藏到系統列的物件圖示上滑過滑鼠（nICo= notifyIcon）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void nICo_MouseMove(object sender, MouseEventArgs e)
        {
            if (Visible || !HiddenIcon) return;
            if (!EventsEnabled) return;
            PauseEvents();

            #region 記錄相關變數
            Keys modifierKeys = ModifierKeys;
            int x = Cursor.Position.X, y = Cursor.Position.Y;
            #endregion

            #region 縮至系統工具列在右方時
            //if (Cursor.Position.Y > this.Top + this.Height ||
            //    Cursor.Position.X > this.Left + this.Width) show_nICo();
            #endregion

            #region 縮至系統工具列在左方時
            //if (Cursor.Position.Y > this.Top + this.Height &&
            //    Cursor.Position.X < 420) show_nICo();//this.Left + this.Width) show_nICo();
            if (y > Screen.PrimaryScreen.Bounds.Height - 230 &&
                x < 80)
            {
                ntfyICo.Visible = false; if (modifierKeys == Keys.Shift) Form1.playSound(Form1.soundLike.press);
                //按下Ctrl時，自動將Quick edit的連結複製到剪貼簿
                if (keyinTextMode) copyQuickeditLinkWhenKeyinMode(modifierKeys);
                //Form1.playSound(Form1.soundLike.exam);
                show_nICo(modifierKeys);//this.Left + this.Width) show_nICo();                
            }
            #endregion

            #region 縮至系統工具列在下方時
            else if (y > Screen.PrimaryScreen.Bounds.Height - 50 &&
                x > Screen.PrimaryScreen.Bounds.Width - 270)
            {
                ntfyICo.Visible = false; if (modifierKeys == Keys.Shift) Form1.playSound(Form1.soundLike.press);
                if (keyinTextMode) copyQuickeditLinkWhenKeyinMode(modifierKeys);
                Form1.playSound(Form1.soundLike.over);
                show_nICo(modifierKeys);//this.Left + this.Width) show_nICo();
            }
            #endregion
            ////if (this.Top <0 && this.Left<0) show_nICo();
            ///
            #region 20230207 creedit with chatGPT大菩薩：失敗
            ////20230207 creedit with chatGPT大菩薩：

            //Point iconLocation = new Point(ntfyICo.Bounds.X, ntfyICo.Bounds.Y);
            //Point iconLocationOnScreen = ntfyICo.Parent.PointToScreen(iconLocation);
            //int iconX = iconLocationOnScreen.X;
            //int iconY = iconLocationOnScreen.Y;

            //Control ni = (Control)sender;
            //Point pnt = ni.PointToScreen(new Point(ni.Left, ni.Top));
            //int iconX = pnt.X, iconY = pnt.Y;
            //// 計算滑鼠位置是否在表單圖示的範圍內
            ////if (e.X >= iconX && e.X <= iconX + iconWidth && e.Y >= iconY && e.Y <= iconY + iconHeight)
            //if (e.X >= iconX && e.X <= iconX + ni.Width && e.Y >= iconY && e.Y <= iconY + ni.Height)
            //{

            //    // 顯示表單
            //    show_nICo();
            //} 
            #endregion
            ResumeEvents();
        }


        /// <summary>
        /// 自動將Quick edit的連結複製到剪貼簿（不一定是在手動模式下才執行）
        /// 按下的控制鍵是Ctrl時才執行
        /// </summary>
        /// <param name="modifierKeys">按下的控制鍵是Ctrl時才執行</param>
        void copyQuickeditLinkWhenKeyinMode(Keys modifierKeys)
        {
            string url = null;
            try
            {
                //如果是完整編輯頁面則逕予返回20240810                
                url = br.GetChromeActiveUrl;
                if (url.IsNullOrEmpty())
                    //url = br.GetActiveChromeTabUrl();//這個太慢了，沒效率
                    if (DialogResult.OK == MessageBoxShowOKCancelExclamationDefaultDesktopOnly("無法取得現用的分頁，是否抓取最後開的一頁？"))
                    {
                        driver.SwitchTo().Window(driver.WindowHandles.Last());
                        url = driver.Url;
                    }
                    else
                    {
                        if (DialogResult.OK == MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否要巡覽到目前現用的分頁？"))
                        {
                            for (int i = driver.WindowHandles.Count - 1; i > -1; i--)
                            {
                                driver.SwitchTo().Window(driver.WindowHandles[i]);
                                if (DialogResult.OK == MessageBoxShowOKCancelExclamationDefaultDesktopOnly("剛才現用的，是這個分頁嗎？"))
                                {
                                    url = driver.Url;
                                    break;
                                }

                            }
                        }
                    }
                if (url.Contains("action=editchapter")) return;
                //if (Clipboard.GetText().IndexOf("<scanbegin ") > -1) return;
                if (Clipboard.GetText().IndexOf("<scanbegin ") > -1)
                {
                    if (keyinTextMode && !ocrTextMode)
                    {
                        //應該是在[新增單位]之後：
                        Clipboard.Clear();
                        if (IsValidUrl＿ImageTextComparisonPage(br.driver.Url))
                        {
                            if (null == br.Page_textbox)//br.driver.Navigate().Refresh();
                            {
                                MessageBoxShowOKExclamationDefaultDesktopOnly("請檢查textBox3中的網址值");
                                return;
                            }
                            else
                                br.Page_textbox.SendKeys(OpenQA.Selenium.Keys.Enter);
                        }
                    }
                }
            }
            catch (Exception)
            {
                return;
            }
            //在規範編輯/修改模式中的文字時不處理
            switch (modifierKeys)
            {
                case Keys.Control:
                    try
                    {
                        br.driver = br.driver ?? Browser.DriverNew();//creedit with chatGPT大菩薩
                        //OpenQA.Selenium.IWebElement quickEditLink = br.driver.FindElement(OpenQA.Selenium.By.XPath("//a[@title='Quick edit']"));
                        OpenQA.Selenium.IWebElement quickEditLink = br.driver.FindElement(OpenQA.Selenium.By.XPath("//*[@id=\"quickedit\"]/a"));
                        if (quickEditLink != null)
                        {
                            string quickEditLinkUrl = quickEditLink.GetAttribute("href");
                            Clipboard.SetText(quickEditLinkUrl);
                        }
                    }
                    catch (Exception ex)
                    {
                        switch (ex.HResult)
                        {
                            case -2146233088:
                                //"stale element reference: element is not attached to the page document\n  (Session info: chrome=110.0.5481.100)"
                                //"no such window: target window already closed\nfrom unknown error: web view not found\n  (Session info: chrome=110.0.5481.100)"
                                if (ex.Message.StartsWith(@"The HTTP request to the remote WebDriver server for URL http://localhost:") && ex.Message.IndexOf("timed out after") > -1)
                                {
                                    MessageBoxShowOKExclamationDefaultDesktopOnly("Chrome瀏覽器分頁頁籤進入休眠省電模式，請藉由激活（Activate）它將之喚醒。");
                                    br.LastValidWindow = br.driver.CurrentWindowHandle;
                                    for (int i = br.driver.WindowHandles.Count - 1; i > -1; i--)
                                    {
                                        br.driver.SwitchTo().Window(br.driver.WindowHandles[i]);
                                    }
                                    br.driver.SwitchTo().Window(br.LastValidWindow);
                                }
                                break;
                            default:
                                //MessageBox.Show(ex.HResult + ex.Message);
                                MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                                //throw;
                                break;
                        }
                    }
                    break;
                //自動擷取「簡單修改模式」（selector: # quickedit > a的連結)準備到《古籍酷》OCR
                case Keys.Shift:
                    toOCR(br.OCRSiteTitle.GJcool);
                    //copyQuickeditLinkWhenKeyinModeSub();
                    //ResetLastValidWindow();
                    break;
                //自動擷取「簡單修改模式」（selector: # quickedit > a的連結)
                case Keys.None:
                    if (!url.IsNullOrEmpty()) Clipboard.SetText(url);//●●●●●●●●●●●●●●●●●●●●●●●●●
                    copyQuickeditLinkWhenKeyinModeSub(url);
                    Clipboard.Clear();
                    ResetLastValidWindow();
                    break;
            }
        }

        internal static void ResetLastValidWindow()
        {
            string wh;
        retry:
            try
            {
                if (br.driver == null)
                {
                    browsrOPMode = BrowserOPMode.seleniumNew;
                    DriverNew();
                }
                if (!br.IsDriverInvalid())
                    wh = br.driver.CurrentWindowHandle;
                else
                {
                    try
                    {
                        br.driver.SwitchTo().Window(driver.WindowHandles.Last());
                        wh = br.driver.CurrentWindowHandle;
                    }
                    catch (Exception)
                    {
                        RestartDriver();//●●●●●●●●●●●●●●●●●●●
                        br.driver.SwitchTo().Window(driver.WindowHandles.Last());
                        wh = br.driver.CurrentWindowHandle;
                    }
                }
            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case -2146233088:
                        if (ex.Message.StartsWith("An unknown exception was encountered sending an HTTP request to the remote WebDriver server for URL "))//http://localhost:5395/session/f4e2d1e40ba5f7dbb4254af16a9ed493/window/handles. The exception message was: 傳送要求時發生錯誤。"))
                        {
                            RestartChromedriver();
                            goto retry;
                        }
                        else
                            goto default;
                    default:
                        Console.WriteLine(ex.HResult + ex.Message);
                        wh = string.Empty;
                        break;
                }
            }
            if (wh != string.Empty)
                br.LastValidWindow = wh;
            else
            {
                try
                {
                    wh = br.LastValidWindow;
                    if (br.driver.WindowHandles.Count > 0 && !br.driver.WindowHandles.Contains(wh))
                        br.LastValidWindow = br.driver.WindowHandles.Last();
                    br.driver.SwitchTo().Window(wh);

                }
                catch (Exception ex)
                {
                    switch (ex.HResult)
                    {
                        case -2146233088:
                            if (ex.Message.StartsWith("disconnected: not connected to DevTools"))//disconnected: not connected to DevTools
                                                                                                 //(failed to check if window was closed: disconnected: not connected to DevTools)
                                                                                                 //  (Session info: chrome = 129.0.6668.59)
                            {
                                //Debugger.Break();
                                MessageBoxShowOKExclamationDefaultDesktopOnly("請關閉Chrome瀏覽器後再按下「確定」以繼續！！感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主");
                                RestartChromedriver();
                                //br.driver = null;
                                //killchromedriverFromHere();
                                //br.DriverNew();
                                goto retry;
                            }
                            else if (ex.Message.StartsWith("An unknown exception was encountered sending an HTTP request to the remote WebDriver server for URL"))//An unknown exception was encountered sending an HTTP request to the remote WebDriver server for URL http://localhost:14698/session/0f432de43d64b3c61bb847ce517358a3/window/handles. The exception message was: 傳送要求時發生錯誤。
                            {
                                if (ChromedriverLose(ex))
                                    goto retry;
                                else
                                    goto default;
                            }
                            else
                                goto default;
                        default:
                            Console.WriteLine(ex.HResult + ex.Message);
                            if (DialogResult.OK ==
                                    Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly(ex.HResult + ex.Message))
                                RestartChromedriver();
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// 自動擷取「簡單修改模式」（selector: # quickedit > a的連結)到剪貼簿
        /// </summary>
        /// <returns>執行成功傳回true</returns>
        bool copyQuickeditLinkWhenKeyinModeSub(string currUrl = "")
        {
            if (Clipboard.GetText().IndexOf("#editor") == -1)
            {
                try
                {
                    br.driver = br.driver ?? Browser.DriverNew();
                    if (currUrl.IsNullOrEmpty()) currUrl = br.GoToCurrentUserActivateTab();
                    if (currUrl == "") return false;
                    if (br.IsDriverInvalid())
                    {
                        for (int i = driver.WindowHandles.Count - 1; i > -1; i--)
                        {
                            br.driver.SwitchTo().Window(driver.WindowHandles[i]);
                            if (br.driver.Url == currUrl) break;
                        }
                    }
                    else
                    {
                        int i = driver.WindowHandles.Count;
                        string url = driver.Url; currUrl = currUrl.StartsWith("http") ? currUrl : "https://" + currUrl;
                        while (i > -1 && url != currUrl)
                        {
                            i--;
                            driver.SwitchTo().Window(driver.WindowHandles[i]);
                            url = driver.Url;
                            if (url.Contains("%"))
                            {
                                //string decodeStr = url.Substring(url.IndexOf("#")+1);
                                //HttpUtility.deco
                                //url = br.DecodedStringURL(url);
                                url = HttpUtility.UrlDecode(url);

                            }

                        }
                    }
                    string quickEditLinkUrl = "";
                    try
                    {
                        quickEditLinkUrl = br.driver.Url;
                    }
                    catch (Exception ex)
                    {
                        switch (ex.HResult)
                        {
                            case -2146233088:
                                //"stale element reference: element is not attached to the page document\n  (Session info: chrome=110.0.5481.100)"
                                //"no such window: target window already closed\nfrom unknown error: web view not found\n  (Session info: chrome=110.0.5481.100)"
                                //br.driver.SwitchTo().Window(br.driver.WindowHandles[br.driver.WindowHandles.Count - 1]);
                                for (int i = driver.WindowHandles.Count - 1; i > -1; i--)
                                {
                                    br.driver.SwitchTo().Window(driver.WindowHandles[i]);
                                    if (br.driver.Url == currUrl) break;
                                }
                                quickEditLinkUrl = br.driver.Url;
                                break;
                            default:
                                MessageBox.Show(ex.HResult + ex.Message);
                                //throw;
                                break;
                        }
                    }

                    if (quickEditLinkUrl.IndexOf("&page=") == -1 ||
                        (quickEditLinkUrl.IndexOf("#editor") > -1 && quickEditLinkUrl.IndexOf("&page=1") > -1))
                    {
                        string foundUrl = string.Empty;
                        for (int i = br.driver.WindowHandles.Count - 1; i > -1; i--)
                        {//找到分頁是書圖圖文對照瀏覽頁面且非第1頁者：                            
                            try
                            {
                                foundUrl = br.driver.SwitchTo().Window(br.driver.WindowHandles[i]).Url;
                            }
                            catch (Exception)
                            {
                                continue;
                            }
                            if (foundUrl.IndexOf("&page=") > -1 && foundUrl.IndexOf("&page=1&") == -1) break;
                        }
                        quickEditLinkUrl = br.driver.Url;
                    }
                    if (quickEditLinkUrl.IndexOf("#editor") == -1 && quickEditLinkUrl.IndexOf("&page=") > -1)
                    {
                        //OpenQA.Selenium.IWebElement quickEditLink = br.
                        //    waitFindWebElementBySelector_ToBeClickable("#quickedit > a");
                        //if (quickEditLink != null)
                        //{
                        //    quickEditLinkUrl = quickEditLink.GetAttribute("href");
                        //}
                        quickEditLinkUrl = br.GetQuickeditUrl();
                        //if (keyinTextMode && br.driver.Url != quickEditLinkUrl)
                        if (br.driver.Url != quickEditLinkUrl)
                            //br.driver.Url = quickEditLinkUrl;
                            br.GoToUrlandActivate(quickEditLinkUrl, true);
                        if (textBox3.Text != quickEditLinkUrl)
                        {
                            bool eventable = _eventsEnabled;
                            _eventsEnabled = true;
                            textBox3.Text = quickEditLinkUrl;
                            WindowHandles["currentPageNum"] = _currentPageNum;
                            _eventsEnabled = eventable;
                        }
                    }
                    if (quickEditLinkUrl.IndexOf("#editor") > -1)
                    {
                        //即使書ID、PageID一致，但若章節變了，對應的圖文對照網址也會改變：20231115
                        int editwikiID = GetEditwikiID_fromUrl(quickEditLinkUrl);
                        if (editwikiID > 0 && editwikiID != GetEditwikiID_fromUrl(textBox3.Text))
                        {
                            bool eventEnable = _eventsEnabled;
                            ResumeEvents();
                            playSound(soundLike.error, true);
                            textBox3.Text = quickEditLinkUrl;
                            _eventsEnabled = eventEnable;
                        }
                        //在手動輸入模式下，讀入Quick edit 方塊中的內容
                        string quicteditX = br.Quickedit_data_textboxTxt;
                        //if (keyinTextMode && !quicteditX.IsNullOrEmpty() && textBox1.Text != quicteditX)
                        if (!quicteditX.IsNullOrEmpty() && textBox1.Text != quicteditX)
                        {
                            playSound(soundLike.over, true);
                            textBox1.Text = quicteditX;
                        }
                        Clipboard.SetText(quickEditLinkUrl);
                    }
                }
                catch (Exception ex)
                {
                    switch (ex.HResult)
                    {
                        case -2146233088:
                            //"stale element reference: element is not attached to the page document\n  (Session info: chrome=110.0.5481.100)"
                            //"no such window: target window already closed\nfrom unknown error: web view not found\n  (Session info: chrome=110.0.5481.100)"
                            Console.WriteLine(ex.HResult + ex.Message);
                            return false;
                        default:
                            MessageBox.Show(ex.HResult + ex.Message);
                            return false;
                    }

                }
                return true;
            }
            else
                return false;
        }

        /// <summary>
        /// 曝露給其他類別操作textBox1的方法。20230126癸卯年初五 creedit with chatGPT大菩薩："C# 繼承 Form1 控制項"：在不更動 textBox1的修飾符（modifier）private 下的權宜方法
        /// </summary>
        /// <param name="operation"></param>
        internal void PerformTextBoxOperation(Action<TextBox> operation)
        {//chatGPT大菩薩：在Form1類別中的PerformTextBoxOperation方法可以接受一個Action<TextBox>類型的參數，該參數是一個委派，將會在呼叫時執行。您可以使用SwapText類別中的Swap方法來實例化該委派並將其傳入PerformTextBoxOperation方法中。如此一來，在執行PerformTextBoxOperation方法時，就會執行SwapText類別中的Swap方法，對Form1上的textBox1進行操作。
            operation(textBox1);
        }

        int FontFamilyNowIndex = 0;
        FontFamily getCJKExtFontInstalled(string fontName)
        { //https://www.cnblogs.com/arxive/p/7795232.html            
            InstalledFontCollection MyFont = new InstalledFontCollection();
            FontFamily[] fontFamilys = MyFont.Families;
            if (fontFamilys == null || fontFamilys.Length < 1)
            {
                return null;
            }
            foreach (FontFamily item in fontFamilys)
            {
                if (item.Name == fontName) return item;
            }
            return null;
        }

        //private void button1_Click(object sender, EventArgs e)
        //{
        //    splitLineByFristLen();
        //    textBox1.Focus();
        //}

        private void splitLineByFristLen()
        {
            //據第一行長度來分行分段//只要插入點（游標）所在位置前有分段，則依其前一段長度來分行分段
            bool noteFlg = false;
            string x = textBox1.Text;
            int selStart = textBox1.SelectionStart; int s;
            //if (x.Substring(selStart).IndexOf(Environment.NewLine) == -1)
            //{// 結果插入點所在處無分段符號，則取其前一段
            s = x.Substring(0, selStart).LastIndexOf(Environment.NewLine);
            if (s > -1)
            {
                selStart = x.Substring(0, s).LastIndexOf(Environment.NewLine) + Environment.NewLine.Length;
                if (selStart == -1) selStart = 0;
            }
            //}
            s = selStart;
            if (selStart == textBox1.Text.Length) selStart = 0;
            if (selStart != 0)
            {
                selToNewline(ref selStart, ref selStart, textBox1.Text, false, textBox1);
            }
            string xPre = textBox1.Text.Substring(0, selStart);
            x = textBox1.Text.Substring(selStart);
            if (x == "") Clipboard.GetText();
            int wordCntr = 0; int noteCtr = 0;
            StringInfo mystrInof = new StringInfo(x);
            if (x.IndexOf(Environment.NewLine) == -1)
            {
                MessageBox.Show("請先於第一行分段");
                return;
            }
            if ((x.IndexOf("{") == -1 && x.IndexOf("}") > -1) || x.IndexOf("}") < x.IndexOf("{"))
            {
                noteFlg = true;
            }
            string resltTxt = x.Substring(0, x.IndexOf("\r\n"));
            x = x.Replace("\r\n", "");
            TextElementEnumerator mystrEnum = StringInfo.GetTextElementEnumerator(resltTxt);
            while (mystrEnum.MoveNext())
            {
                string mystr = mystrEnum.Current.ToString();
                if (mystr == "{") noteFlg = true;
                if (mystr == "}")
                {
                    noteFlg = false;
                    if (noteCtr % 2 == 1) noteCtr++;
                }
                if (noteFlg)
                {
                    if (omitStr.IndexOf(mystr) == -1)
                    {
                        noteCtr++;
                    }
                }
                else
                {
                    if (omitStr.IndexOf(mystr) == -1)
                    {
                        wordCntr++;

                    }

                }
            }
            int lineLen;
            if (wordCntr == 0 && noteFlg)//純注文
                lineLen = noteCtr;
            else
                lineLen = wordCntr + noteCtr / 2;//wordCntr+((int)Math.Round(noteCtr/2.0));
            NormalLineParaLength = lineLen;
            resltTxt = ""; wordCntr = 0; noteCtr = 0; int noteBrk = 0; int noteBrkCtr = 0; noteFlg = false;
            if ((x.IndexOf("{") == -1 && x.IndexOf("}") > -1) || x.IndexOf("}") < x.IndexOf("{"))
            {
                noteFlg = true;
            }
            mystrEnum = StringInfo.GetTextElementEnumerator(x);
            while (mystrEnum.MoveNext())
            {
                string mystr = mystrEnum.Current.ToString();
                //if (mystr == "《" || mystr == "〈")
                //{
                //    break;
                //}
                if (mystr == "{")
                {
                    noteFlg = true;
                }
                if (mystr == "}")
                {
                    noteFlg = false;
                    if (noteCtr % 2 == 1) noteCtr++;
                }
                if (noteFlg)
                {//如果是注文                    
                    if (omitStr.IndexOf(mystr) == -1)
                    {
                        noteCtr++;
                    }

                }
                else
                {//正文                    
                    if (omitStr.IndexOf(mystr) == -1)
                        wordCntr++;
                }
                resltTxt += mystr;

                //if (wordCntr + Math.Ceiling(noteCtr / 2.0) == lineLen)
                //if (wordCntr + Math.Round(noteCtr / 2.0) == lineLen)                
                if (wordCntr + noteCtr / 2 == lineLen)
                {
                    if (wordCntr == 0)
                    {//純注文
                        StringInfo resltxtinof = new StringInfo(resltTxt);
                        for (int i = resltxtinof.LengthInTextElements; i > 0; i--)//-1; i--)
                        {

                            if (omitStr.IndexOf(resltxtinof.SubstringByTextElements(i - 1, 1)) == -1)
                            { noteBrkCtr++; }
                            if (noteBrkCtr == lineLen)
                            {
                                noteBrk = i - 1;
                                noteBrkCtr = 0;
                                break;
                            }
                        }
                        resltTxt = resltxtinof.SubstringByTextElements(0, noteBrk)
                            + System.Environment.NewLine + resltxtinof.SubstringByTextElements(noteBrk);
                    }
                    resltTxt += "\r\n";
                    wordCntr = 0;
                    noteBrk = 0;
                    noteCtr = 0;
                }


            }
            undoRecord();
            textBox1.Text = xPre + resltTxt.Replace("}" + Environment.NewLine + "}", "}}" + Environment.NewLine)
                .Replace(Environment.NewLine + "》", "》" + Environment.NewLine)
                .Replace(Environment.NewLine + "〉", "〉" + Environment.NewLine)
                .Replace(Environment.NewLine + "}}", "}}" + Environment.NewLine)
                .Replace("{{" + Environment.NewLine, Environment.NewLine + "{{");
            //textBox1.Focus();
            //textBox1.SelectionStart = s + resltTxt.IndexOf(Environment.NewLine) + Environment.NewLine.Length;//selStart;
            //textBox1.SelectionLength = 0;
            //textBox1.ScrollToCaret();
            s = s + resltTxt.IndexOf(Environment.NewLine) + Environment.NewLine.Length;
            restoreCaretPosition(textBox1, s, 0);
            ////Clipboard.SetText(resltTxt);
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" && ModifierKeys == Keys.None)
            {
                if (isClipBoardAvailable_Text())
                {
                    try
                    {
                        textBox1.Text = Clipboard.GetText();
                    }
                    catch (Exception)
                    {
                        Thread.Sleep(2000);
                    }
                    textBox1.Select(0, 0);
                    textBox1.ScrollToCaret();
                }
            }
        }

        string lastFindStr = "";
        private void textBox2_Leave(object sender, EventArgs e)
        {
            if (button2.Text == "選取文") return;
            string s = textBox2.Text;
            if (s == "" || s == textBox1.SelectedText)
            { textBox2.BackColor = textBox2BackColorDefault; return; }
            //如何判斷字串是否代表數值 (c # 程式設計手冊):https://docs.microsoft.com/zh-tw/dotnet/csharp/programming-guide/strings/how-to-determine-whether-a-string-represents-a-numeric-value
            bool result = int.TryParse(s, out int i); //i now = textBox2.Text
            if (result && (processID == null || processID == ""))
            {
                processID = s;
            }
            string x = textBox1.Text; int xStart = x.IndexOf(s), nextStart = x.IndexOf(s, xStart + 1);
            Color C = textBox2.BackColor;
            if (s != "") lastFindStr = s;
            if (xStart > -1)
            {//若有找到
                textBox1.Focus();
                textBox1.Select(xStart, textBox2.Text.Length);
                textBox1.ScrollToCaret();
                x = x.Substring(0, xStart + textBox2.Text.Length);
                Clipboard.SetText(x);
                //textBox1.Text = textBox1.Text.Substring(xStart + 2);
                if (nextStart > -1) textBox2.BackColor = Color.Yellow;//若符合尋找的字串並非獨一無二，則 textBox2 會顯示黃色
                else textBox2.BackColor = textBox2BackColorDefault;
            }
            else
            {//若沒找到
                textBox2.BackColor = Color.Red;
                Task.Delay(500).Wait();
                //C# Leave envent cancel
                textBox2.BackColor = Color.GreenYellow;//https://docs.microsoft.com/zh-tw/dotnet/api/system.windows.forms.control.leave?view=windowsdesktop-6.0
                textBox2.Focus();

                //TextBox tb = (TextBox)sender;
                //此法成了移除了：
                //tb.Leave -= textBox2_Leave;//https://stackoverflow.com/questions/2664639/cancel-leave-event-when-closing
                /*
                private void tabPage1_Validating(object sender,System.ComponentModel.CancelEventArgs e)

                {//https://social.msdn.microsoft.com/Forums/en-US/0a6251c5-b4bd-42a7-bbd3-cdac893df04f/how-to-cancel-or-abort-occurs-leave-event-of-tabcontrol?forum=csharplanguage

                    if (!checkValidated.Checked)

                        e.Cancel = true;

                }*/
            }
            //textBox2.BackColor = textBox2BackColorDefault;
        }
        /// <summary>
        /// 標上小注標記{{……}} 20240905修訂
        /// </summary>
        void noteMark()//Ctrl + F1 ：選取範圍前後加上{{}}
        {
            if (insertMode && textBox1.SelectionLength > 0 || !insertMode)
            {
                undoRecord();
                if (textBox1.SelectionLength == 0)
                    overtypeModeSelectedTextSetting(ref textBox1);
                string x = textBox1.SelectedText;//如果其中有分行/段符號，則必均為獨立行小注（非與正文夾注者），故直接於其後前分別標識{{、}}，蓋}}{{可Ctext有自行取消成""之機制已。感恩感恩　讚歎讚歎　南無阿彌陀佛 20231031
                stopUndoRec = true; PauseEvents();
                //textBox1.SelectedText = ("{{" + x + "}}").Replace(Environment.NewLine, "}}" + Environment.NewLine + "{{")
                //    .Replace("{{{{", "{{").Replace("}}}}", "}}");
                x = ("{{" + x + "}}").Replace(Environment.NewLine, "}}" + Environment.NewLine + "{{");

                CnText.CurlybracesFormalizer(ref x);

                #region 全注文標記之處理
                int s = textBox1.SelectionStart,
                    openBrace = textBox1.Text.LastIndexOf("{{", s), closeBrace = textBox1.Text.LastIndexOf("}}", s),
                    paraMark = textBox1.Text.LastIndexOf("<p>", s);
                //檢查前面是否下大括號在上大括弧前或沒有下大括號存在,有則清除前端的「{{」
                if (s > "{{".Length
                    && ((closeBrace == -1 && openBrace != -1) || (openBrace == -1 && closeBrace != -1))
                    || (closeBrace < openBrace
                        && (paraMark == -1 || paraMark < openBrace)))
                    x = x.Substring("{{".Length);
                //若前一行段末為下大括弧，則予清理
                if (s > "}}".Length + Environment.NewLine.Length
                    && textBox1.Text.Substring(s - 4, 4) == "}}" + Environment.NewLine)
                {
                    textBox1.Select(textBox1.SelectionStart - 4, textBox1.SelectionLength + 4);
                    x = string.Empty + Environment.NewLine + x.Substring("{{".Length);
                }
                #endregion

                textBox1.SelectedText = x;

                if (!Active) bringBackMousePosFrmCenter();
                stopUndoRec = false; ResumeEvents();
            }
            else if (keyinTextMode && textBox1.SelectionLength == 0)
            {
                new Document(ref textBox1).MergeParagraphsAtCaret();
            }
        }

        /// <summary>
        /// 取得游標/插入點所在行/段的文字（含標點標誌tag(*|<p>)）
        /// </summary>
        /// <param name="x">要處理的文本</param>
        /// <param name="s">插入點位置</param>
        /// <returns>失敗則傳回空字串</returns>
        internal static string GetLineText(string x, int s)
        {
            if (s < 0 || string.IsNullOrEmpty(x)) return string.Empty;
            int preP = x.LastIndexOf(Environment.NewLine, s), p = x.IndexOf(Environment.NewLine, s);
            //if (p == 0) return "";///////////watching  if ok  then the comment can be remove 20230617  
            int lineS = preP == -1 ? 0 : preP + (preP == -1 ? 0 : Environment.NewLine.Length);
            int lineL = p == -1 ? x.Length - lineS : preP == -1 ? p : p - Environment.NewLine.Length - preP;
            if (lineL < 0) return string.Empty;
            return x.Substring(lineS, lineL);
        }
        /// <summary>
        /// 取得游標/插入點所在行/段的文字+行的起始位置與長度（含標點與標誌tag(*|<p>)）
        /// </summary>
        /// <param name="x">要處理的文本</param>
        /// <param name="s">插入點位置</param>
        /// <param name="lineS">本行的起始位置</param>
        /// <param name="lineL">本行的長度</param>
        /// <returns></returns>
        internal static string GetLineText(string x, int s, out int lineS, out int lineL)
        {
            if (s < 0 || string.IsNullOrEmpty(x)) { lineS = 0; lineL = 0; return ""; }
            int preP = x.LastIndexOf(Environment.NewLine, s), p = x.IndexOf(Environment.NewLine, s);
            lineS = preP == -1 ? 0 : preP + Environment.NewLine.Length;
            //lineL = p == -1 ? x.Length - lineS : p - Environment.NewLine.Length - preP;
            lineL = p == -1 ? x.Length - lineS : p - lineS;
            lineL = lineL < 0 ? 0 : lineL;
            return x.Substring(lineS, lineL);
        }

        /// <summary>
        /// 取得游標/插入點所在行文字（不含標點）
        /// </summary>
        /// <param name="x">要處理的文本</param>
        /// <param name="s">插入點位置</param>
        /// <returns></returns>
        internal static string GetLineText_WithoutPunctuation(string x, int s)
        {
            if (s < 0 || string.IsNullOrEmpty(x)) return "";
            string returnTxt = GetLineText(x, s);
            //https://useadrenaline.com/playground
            //20230115 adrenaline 大菩薩：
            for (int i = 0; i < PunctuationsNum.Length; i++)
            {
                //returnTxt = returnTxt.Replace(punctuationsNum[i].ToString(), " ".ToCharArray()[0].ToString());
                returnTxt = returnTxt.Replace(PunctuationsNum[i].ToString(), string.Empty);
            }
            return returnTxt.Replace("   ", " ").Replace("  ", " ");

            //for (int i = 0; i < punctuations.Length; i++)
            //{
            //    returnTxt.Replace(punctuations[i], "".ToCharArray()[0]);
            //}
            //return returnTxt;
        }
        /// <summary>
        /// 取得插入點下一行/段的文字
        /// </summary>
        /// <param name="x">要找的全域文字</param>
        /// <param name="s">插入點所在位置</param>
        /// <returns></returns>
        internal static string GetNextLineText_IncludingMarkers(string x, int s)
        {
            if (s < 0 || string.IsNullOrEmpty(x)) return "";
            int p = x.IndexOf(Environment.NewLine, s);
            if (p == -1) return string.Empty;
            p += Environment.NewLine.Length;
            int nextP = x.IndexOf(Environment.NewLine, p);
            if (nextP == -1) nextP = x.Length;
            return x.Substring(p, nextP - p);
        }

        bool chkPTitleNotEnd = false;//為檢查不當分段設置的，以判斷前一頁末是否不含標明尾，其尾卻在此頁前，以便略過檢查

        /// <summary>
        /// 貼去Ctext 後（包括對要貼去的內容作最後的檢查）設定新的 textBox1的內容。若執行成功則傳回true
        /// </summary>
        /// <param name="s">若檢查不通過，傳回有問題的地方起始點start</param>
        /// <param name="l">若檢查不通過，傳回有問題的地方之長度length</param>
        /// <returns>若通過檢查，執行成功則傳回true；否則為false</returns>
        private bool newTextBox1(out int s, out int l, string x)
        {
            s = textBox1.SelectionStart; l = textBox1.SelectionLength;
            //string x = textBox1.Text;
            if (x == string.Empty) x = textBox1.Text;
            if (x == "") return false;
            textBox1.Select(0, s + l); string xHandle = textBox1.SelectedText;

        retry:
            try
            {
                saveText();//備份以防萬一
            }
            catch (Exception)
            {
                Thread.Sleep(500);
                goto retry;
            }
            //if (textBox1.SelectedText != "")
            //{
            if (textBox2.Text != "＠" && textBox2.Text != "") textBox2.Text = "";


            #region 清除冗碼（清除冗餘要留意會動到 l 的值！！）
            if (x.IndexOf("<p>|") > -1 || x.IndexOf("|<p>") > -1 || x.IndexOf("||") > -1)
            {//先除掉全部的，再清除目前要貼上的
                x = x.Replace("<p>|", "<p>").Replace("|<p>", "<p>").Replace("||", "|");
                if (xHandle.IndexOf("<p>|") > -1 || xHandle.IndexOf("|<p>") > -1
                    || xHandle.IndexOf("||") > -1)
                {
                    textBox1.SelectedText = xHandle.Replace("<p>|", "<p>").Replace("|<p>", "<p>")
                        .Replace("||", "|");
                    s = textBox1.SelectionStart; l = textBox1.SelectionLength;
                }
            }
            #endregion

            if (pageTextEndPosition - 10 < 0 || pageTextEndPosition > x.Length) pageTextEndPosition = s;
            if (pageTextEndPosition != 0 && s + l < pageTextEndPosition)
            {
                if (pageEndText10 != x.Substring(pageTextEndPosition - 10, 10))
                {
                    int es = x.IndexOf(pageEndText10);
                    if (es > -1)
                    {
                        es += 10;
                        int 孫守真按 = 20;
                        if (x.Substring(0, pageTextEndPosition).IndexOf("{{{孫守真按：") > -1) 孫守真按 = 40;
                        if (Math.Abs(es - pageTextEndPosition) < 孫守真按)
                        {
                            pageTextEndPosition = es;
                        }
                        else
                        {
                            string PredictCopyX = CnText.GetSelectionTextByLineParaCount(ref textBox1, linesParasPerPage / 2);
                            if (DialogResult.OK == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("請重新指定頁面結束位置" + Environment.NewLine + Environment.NewLine +
                                   "是不是這些內容？" +
                                   Environment.NewLine + Environment.NewLine +
                                   PredictCopyX))
                            {
                                l = 0; s = PredictCopyX.Length; pageTextEndPosition = s; pageEndText10 = PredictCopyX.Substring(s - 10);
                            }
                            else
                            {
                                //MessageBox.Show("請指定頁尾處位置");
                                //MessageBoxShowOKExclamationDefaultDesktopOnly("請指定頁尾處位置");
                                textBox1.Select(pageTextEndPosition, 0); pageTextEndPosition = 0;
                                pageEndText10 = ""; return false;
                            }
                        }

                    }
                    else
                    {
                        //MessageBox.Show("請指定頁尾處位置"); 
                        MessageBoxShowOKExclamationDefaultDesktopOnly("請指定頁尾處位置");
                        textBox1.Select(pageTextEndPosition, 0); pageTextEndPosition = 0;
                        pageEndText10 = ""; return false;
                    }
                }
                else
                {
                    int es = x.IndexOf(pageEndText10);
                    if (es > 2000) pageTextEndPosition = textBox1.SelectionStart;
                    else
                    {
                        es += 10;
                        if (es > -1 && es > pageTextEndPosition)
                        {
                            pageTextEndPosition = es;
                        }
                    }

                }
                if (s != pageTextEndPosition) s = pageTextEndPosition;
            }
            if (s == x.Length) l = 0;
            if (s + l <= x.Length)
            {
                if (x.Substring(0, s + l) == "") return false;
            }
            else
            { s = textBox1.SelectionStart; l = textBox1.SelectionLength; }
            string xCopy = x.Substring(0, s + l > x.Length ? x.Length : s + l);


            #region //規範化文本，如置換為全形符號、及清除冗餘（清除冗餘要留意會動到 l 的值！！）
            //此方法目前沒有別處參考，希望與checkAbnormalLinePara等合併或交互參照來做，但目前仍是取還textBox1的值，且s、l牽連影響甚大，俟後 20230806
            //CnText.FormalizeText(ref xCopy);//在呼叫端執行過，暫略看看 20240426

            //string[] replaceDChar = { "'", ",", ";", ":", "．", "?", "：：", "《《", "》》", "〈〈", "〉〉", "。}}。}}", "。。", "，，", "@" };
            //string[] replaceChar = { "、", "，", "；", "：", "·", "？", "：", "《《", "》", "〈", "〉", "。}}", "。", "，", "●" };
            //foreach (var item in replaceDChar)
            //{
            //    if (xCopy.IndexOf(item) > -1)
            //    {
            //        //if (MessageBox.Show("含半形標點，是否取代為全形？", "", MessageBoxButtons.OKCancel,
            //        //    MessageBoxIcon.Error) == DialogResult.OK)
            //        //{//直接將半形標點符號轉成全形
            //        for (int i = 0; i < replaceChar.Length; i++)
            //        {
            //            xCopy = xCopy.Replace(replaceDChar[i], replaceChar[i]);
            //        }
            //        //}
            //        break;
            //    }
            //}
            ////置換中文文本中的英文句號（小數點）
            //CnText.PeriodsReplace_ChinesePunctuationMarks(ref xCopy);
            #endregion
            #region 清空末尾空行段落
            int blankParagraphPosition = xCopy.LastIndexOf(Environment.NewLine);
            while ((xCopy.Length - 2) > -1 && xCopy.Length == blankParagraphPosition + 2)
            {
                xCopy = xCopy.Substring(0, xCopy.Length - 2);
                blankParagraphPosition = xCopy.LastIndexOf(Environment.NewLine);
            }
            #endregion
            #region 將連空行段落前綴|字符
            blankParagraphPosition = xCopy.IndexOf(Environment.NewLine);
            while (blankParagraphPosition > -1)
            {
                //if (blankParagraphPosition + 4 >= xCopy.Length) break;
                if (xCopy.Substring(0, 2) == Environment.NewLine)
                {
                    xCopy = "|" + xCopy;
                }
                else if (xCopy.Length >= blankParagraphPosition + 2 + 2 && xCopy.Substring(blankParagraphPosition + 2, 2) == Environment.NewLine)
                {
                    xCopy = xCopy.Substring(0, blankParagraphPosition + 2) + "|" + xCopy.Substring(blankParagraphPosition + 2);
                }
                blankParagraphPosition = xCopy.IndexOf(Environment.NewLine, blankParagraphPosition + 1);
                if (blankParagraphPosition + 4 >= xCopy.Length) break;
            }
            #endregion

            #region 檢查無效的漢字字元
            int missWordPositon = xCopy.IndexOf(" ");
        checkOthers:
            if (missWordPositon == -1)
            {//如果沒有半形空格
                missWordPositon = xCopy.IndexOfAny("�".ToCharArray());
                if (missWordPositon == -1) missWordPositon = xCopy.IndexOf("□");
                if (missWordPositon == -1) missWordPositon = xCopy.IndexOf("◊");
                if (missWordPositon == -1) missWordPositon = xCopy.IndexOf("▫");
                if (missWordPositon == -1) missWordPositon = xCopy.IndexOf("စ");
                if (missWordPositon == -1) missWordPositon = xCopy.IndexOf("ခ");
                if (missWordPositon == -1) missWordPositon = xCopy.IndexOf("င");
                if (missWordPositon == -1) missWordPositon = xCopy.IndexOf("ဇ");
                if (missWordPositon == -1) missWordPositon = xCopy.IndexOf("ဌ");
                if (missWordPositon == -1) missWordPositon = xCopy.IndexOf("◍");
                if (missWordPositon == -1) missWordPositon = xCopy.IndexOf("ᗍ");
                if (missWordPositon == -1) missWordPositon = xCopy.IndexOf("Ⲳ");
                if (missWordPositon == -1) missWordPositon = xCopy.IndexOf("⛋");
                if (missWordPositon == -1) missWordPositon = xCopy.IndexOf("ဂ");
                if (missWordPositon == -1) missWordPositon = xCopy.IndexOf("ဃ");
                if (missWordPositon == -1) missWordPositon = xCopy.IndexOf("ဆ");
                if (missWordPositon == -1) missWordPositon = xCopy.IndexOf("ဈ");
                if (missWordPositon == -1) missWordPositon = xCopy.IndexOf("ဉ");
                if (missWordPositon == -1) missWordPositon = xCopy.IndexOf("▱");
                if (missWordPositon == -1) missWordPositon = xCopy.IndexOf("ꗍ");
                //以下《國學大師》本《四庫全書》發生者
                if (missWordPositon == -1) missWordPositon = xCopy.IndexOf("B");
                if (missWordPositon == -1) missWordPositon = xCopy.IndexOf("/");
                //檢查「/」
                if (xCopy.IndexOf("/") > -1 && missWordPositon != -1)
                {//20240803 Copilot大菩薩：C# Windows.Forms 應用程式中檢測半形空格 ：因為在檢查斜線「/」後面不是接著下角括號「>」的情況下，我們需要確保每個匹配的結果都符合這個條件，所以使用了 matches 物件來進一步檢查。
                 //而在檢查半形空格的情況下，正則表達式(?< ! [\w""'\\])\s(?![\w""'\\]) 已經包含了前後字符的條件，因此不需要進一步使用 matches 物件來檢查。
                    string pattern = @"(?<!<)/(?!>)";//20240804 Copilot大菩薩：修正正則表達式問題：您可以使用以下的正則表達式來同時檢測「/」前面不是「<」且後面不是「>」的情況：https://sl.bing.net/hBPULfGmh0m
                    MatchCollection matches = Regex.Matches(xCopy, pattern);//這裡的 (?<!<) 是一個負向前瞻，用來確保「/」前面不是「<」，而 (?!>) 是一個負向後瞻，用來確保「/」後面不是「>」。……在後瞻的語法中，所謂的「額外符號」就是指「<」這個符號。這個符號用來指明我們要檢查的字符是在目標字符的前面。
                    if (matches.Count == 0)//如 <entity entityid="808941" type="work">括異志</entity>  https://ctext.org/library.pl?if=gb&file=234676&page=117&editwiki=186218#editor
                        missWordPositon = -1;
                    else
                        missWordPositon = matches[0].Index;
                    #region 另一種方式：
                    /* 
                     
                    bool isValid = true;

                    // 檢查後一個字符
                    if (missWordPositon < xCopy.Length - 1)
                    {
                        char nextChar = xCopy[missWordPositon + 1];
                        if (nextChar == '>')
                        {
                            isValid = false;
                        }
                    }

                    if (isValid)
                    {
                        break;
                    }

                    missWordPositon = xCopy.IndexOf("/", missWordPositon + 1);
                    */
                    #endregion 

                }
            }
            else//檢查半形空格" "
            {//20240803 Copilot大菩薩：C# Windows.Forms 應用程式中檢測半形空格
                //missWordPositon = xCopy.IndexOf(" ");
                while (missWordPositon != -1)
                {
                    bool isValid = true;

                    // 檢查前一個字符
                    if (missWordPositon > 0)
                    {
                        char prevChar = xCopy[missWordPositon - 1];
                        if ((prevChar >= '0' && prevChar <= '9') ||
                            (prevChar >= 'A' && prevChar <= 'Z') ||
                            (prevChar >= 'a' && prevChar <= 'z') ||
                            prevChar == '"' || prevChar == '/')
                        {
                            isValid = false;
                        }
                    }

                    // 檢查後一個字符
                    if (missWordPositon < xCopy.Length - 1)
                    {
                        char nextChar = xCopy[missWordPositon + 1];
                        if ((nextChar >= '0' && nextChar <= '9') ||
                            (nextChar >= 'A' && nextChar <= 'Z') ||
                            (nextChar >= 'a' && nextChar <= 'z') ||
                            nextChar == '"' || nextChar == '/')
                        {
                            isValid = false;
                        }
                    }

                    if (isValid)
                    {
                        break;
                    }

                    missWordPositon = xCopy.IndexOf(" ", missWordPositon + 1);
                }

                if (missWordPositon == -1)
                {
                    goto checkOthers;
                }


                int lastspacePosition = xCopy.LastIndexOf(" ");
                // 空白錯亂之注文要先處理最末一行（因為送出前會清除每頁最末的注腳標記「}}」會影響 Alt + b 的判斷式，如是才能用 Ctrl + z 來還原之） 20250618●●●●●●●●●●●●●●●
                if (xCopy.Substring(lastspacePosition).IndexOf(Environment.NewLine) == -1
                        && xCopy.Substring(lastspacePosition).IndexOf("}}") == -1)
                {
                    playSound(soundLike.notify, true);
                    Thread.Sleep(300);
                    if (lastspacePosition != missWordPositon) missWordPositon = lastspacePosition;
                }

            }
            #endregion

            #region 檢查不當分段

            int chkP = xCopy.IndexOf("<p>") + ("<p>".Length + Environment.NewLine.Length);//檢查不當分段（目前僅找最前面的一個，餘於停下手動檢索時人工目測
            if (keyinTextMode) { chkP = -1; goto chksum; }//手動輸入、非半自動連續輸入時，略過不處理
            if (xCopy.IndexOf("<p>") > -1 && chkP + 1 <= x.Length)//&& ("　􏿽|" + Environment.NewLine).IndexOf(x.Substring(chkP, 1)) == -1)
            {
                int asteriskPos = xCopy.IndexOf("*");

                if ("{}".IndexOf(x.Substring(chkP, 1)) > -1)
                {
                    if (asteriskPos > -1 && asteriskPos < chkP) chkP = -1;
                }
                else
                {
                    if (asteriskPos > -1 && asteriskPos > chkP)
                    {
                        //須檢查的<p>和* 之間不能再有<p>，因為標題*號前必有<p>
                        if (xCopy.Substring(chkP, asteriskPos - chkP).IndexOf("<p>") == -1) chkP = -1;
                    }
                }
                //如果是標題
                int prePPos = chkP - ("<p>".Length + Environment.NewLine.Length) > -1 ?
                    xCopy.LastIndexOf(Environment.NewLine, chkP - ("<p>".Length + Environment.NewLine.Length)) : -1;
                if (asteriskPos > -1)
                {//標題*與<p>在同一行，長度沒跨行
                    int pre = prePPos > -1 ? prePPos + 2 : 0;
                    if (pre < asteriskPos && chkP - "<p>".Length + Environment.NewLine.Length > asteriskPos &&
                        GetLineText(xCopy, chkP - ("<p>".Length + Environment.NewLine.Length)).IndexOf(Environment.NewLine) == -1)
                        //xCopy.Substring(pre, chkP - ("<p>".Length + Environment.NewLine.Length) - pre).IndexOf(Environment.NewLine) == -1)

                        chkP = -1;
                }
                if (prePPos > -1 && chkP > -1)
                {
                    if (Math.Abs(
                        new StringInfo(xCopy.Substring(prePPos + 2,
                        chkP - ("<p>".Length + Environment.NewLine.Length) - (prePPos + 2))).LengthInTextElements
                         - NormalLineParaLength) > 2)//檢查<p>位置不當者（不當分段）                    
                    //if (Math.Abs(
                    //        CountWordsLenPerLinePara(GetLineText(xCopy, prePPos + 2))
                    //        - NormalLineParaLength) > 2)//檢查<p>位置不當者（不當分段）
                    {
                        chkP = -1;
                    }
                    if (chkP > -1)
                    {
                        //如果前文不是縮排,後面不再縮排
                        if ("　􏿽".IndexOf(xCopy.Substring(prePPos + 2, 1)) > -1 &&
                            x.Substring(prePPos + 2, chkP - (prePPos + 2)).IndexOf("*") == -1 &&//前一行不是標題
                            "　􏿽".IndexOf(xCopy.Substring(x.LastIndexOf(Environment.NewLine, prePPos) + 2, 1)) > -1 &&
                            GetLineText(xCopy, prePPos).IndexOf("*") == -1 &&//前二行不是標題
                            (x.IndexOf(Environment.NewLine, chkP) + 2 + 1) <= x.Length &&
                            "　􏿽".IndexOf(x.Substring(x.IndexOf(Environment.NewLine, chkP) + 2, 1)) == -1)
                        {
                            chkP = -1;
                        }
                    }
                    if (chkP > -1)
                    {//後面有縮排一行如標題者：
                        if ((x.IndexOf(Environment.NewLine, chkP) + 2 + 1) <= x.Length &&
                            "　􏿽".IndexOf(x.Substring(x.IndexOf(Environment.NewLine, chkP) + 2, 1)) > -1 &&
                            "　􏿽".IndexOf(xCopy.Substring(x.LastIndexOf(Environment.NewLine, prePPos) + 2, 1)) == -1)
                            chkP = -1;
                    }
                }
                if (chkP > -1)
                {//如果是單行標題
                    if (GetLineText(xCopy, chkP - ("<p>".Length + Environment.NewLine.Length)).IndexOf("*") > -1) chkP = -1;
                }
                if (chkP > -1)
                {//過短的行略過不檢查
                    if (Math.Abs(CountWordsLenPerLinePara(GetLineText(xCopy, chkP - ("<p>".Length + Environment.NewLine.Length)).Replace("<p", ""))
                        - NormalLineParaLength) > 2)//3)
                    {
                        chkP = -1;
                    }
                }
                if (chkP > -1)
                {//後面是縮排
                    if (("　􏿽" + Environment.NewLine).IndexOf(x.Substring(chkP, 1)) > -1) chkP = -1;
                    //後面是凸排
                    else
                    {
                        int nextPPos = x.IndexOf(Environment.NewLine, chkP);
                        if (nextPPos > -1)
                        {
                            if (nextPPos + 2 + 1 <= x.Length &&
                                    ("　􏿽" + Environment.NewLine).IndexOf(x.Substring(
                                nextPPos + 2, 1)) > -1)
                                chkP = -1;
                        }
                    }
                }
                //如果正文中真有分段，則設定chkP=0作為提示音就好,一如含有「□」之文本處理方式
                //(檢查過沒問題者於<p>前加
                //以識別(用「+」乃不可能存在的字符故）
                if (chkP > -1 && chkPTitleNotEnd == false)
                {
                    if (chkP - ("<p>".Length + Environment.NewLine.Length) - 1 > -1 &&
                        xCopy.Substring(chkP - ("<p>".Length + Environment.NewLine.Length) - 1, 1) == "+")
                    {
                        chkP = 0;
                        xCopy = xCopy.Replace("+<p>", "<p>");
                    }
                }
                if (chkPTitleNotEnd && chkP > -1)
                {
                    chkP = -1; chkPTitleNotEnd = false;
                }
                if (xCopy.LastIndexOf("*") > -1 && xCopy.LastIndexOf("*") > xCopy.LastIndexOf("<p>"))
                    chkPTitleNotEnd = true;
                else
                if (chkPTitleNotEnd) chkPTitleNotEnd = false;
                if (chkP > -1 && x.Substring(chkP, 1) == "*")
                    chkP = -1;

            }
            else chkP = -1;
            #endregion
            chksum:
            if (missWordPositon > -1 || chkP > -1)
            //if (xCopy.IndexOf(" ") > -1 || xCopy.IndexOfAny("�".ToCharArray()) > -1 ||
            //xCopy.IndexOf("□") > -1)//□為《維基文庫》《四庫全書》的缺字符，" "則是《四部叢刊》的，"�"則是《四部叢刊》的造字符。
            {//  「�」甚特別，indexof會失效，明明沒有，而傳回 0 //https://docs.microsoft.com/zh-tw/dotnet/csharp/how-to/compare-strings
             //  //https://docs.microsoft.com/zh-tw/dotnet/api/system.string.compare?view=net-6.0
             //SystemSounds.Hand.Play();//文本有缺字警告
             //if (File.Exists(soundWarningLocation)) new SoundPlayer(soundWarningLocation).Play();
                Color c = this.BackColor;
                new SoundPlayer(@"C:\Windows\Media\Windows Notify Email.wav").Play();
                this.BackColor = Color.Yellow;
                Task.Delay(400).Wait();
                this.BackColor = c;

                //若要忽略則為true
                bool omit = false;
                if (chkP > 0)
                {
                    string xTbx1 = textBox1.Text; int double_chkP = chkP - (+"<p>".Length + Environment.NewLine.Length);
                    string[] stringPre = { "銘曰", "辭曰" };
                    string[] stringNext = { "論曰", "舊史氏曰" };
                    string[][] strings = { stringPre, stringNext };
                    for (int i = 0; i < strings.Length; i++)
                    {
                        foreach (var item in strings[i])
                        {
                            int sItem = 0;
                            switch (i)
                            {
                                case 0://pre
                                    sItem = double_chkP - item.Length;
                                    if (sItem >= 0)
                                        if (xTbx1.Substring(sItem, item.Length) == item) { omit = true; break; }
                                    break;
                                case 1://next
                                    sItem = double_chkP + "<p>".Length + Environment.NewLine.Length;
                                    if (xTbx1.Length >= sItem + item.Length)
                                        if (xTbx1.Substring(sItem, item.Length) == item) { omit = true; break; }
                                    break;
                                default:
                                    break;
                            }
                            if (omit) break;
                        }
                        if (omit) break;
                    }
                    if (!omit)
                    {
                        textBox1.Select(chkP - (+"<p>".Length + Environment.NewLine.Length), 3);
                        textBox1.ScrollToCaret();
                        Clipboard.SetText("　");//準備空格以填補缺額 20240913 擬作廢，蓋已有 Alt + 2 方便輸入空格了
                        return false;
                    }
                }
                //if (xCopy.IndexOf("□") > -1 && xCopy.IndexOfAny("�".ToCharArray()) == -1 && xCopy.IndexOf(" ") == -1
                //    || chkP == 0 || omit)
                if ((xCopy.IndexOf("□") > -1 && !keyinTextMode) && xCopy.IndexOfAny("�".ToCharArray()) == -1 && xCopy.IndexOf(" ") == -1
                || chkP == 0 || omit)
                {
                    //if (DialogResult.OK==MessageBox.Show("有造字，是否先予訂補上？", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1) )
                    //{
                    //●●●●●●●●●●●●●●● 20250227
                    if (missWordPositon > -1 && xCopy.IndexOf("+<p>") == -1)
                    {
                        textBox1.Select(missWordPositon, 1);
                        textBox1.ScrollToCaret();
                        return false;
                    }
                }
                else
                {
                    textBox1.Select(missWordPositon, insertMode ? 1 : 0);
                    textBox1.ScrollToCaret();
                    return false;
                }
                //string[] rTxt = { " ", "�" };//, "□" };
                //foreach (string rs in rTxt)
                //{
                //    xCopy = xCopy.Replace(rs, "●");//「●」為《中國哲學書電子化計劃》的缺字符，詳：https://ctext.org/instructions/wiki-formatting/zh
                //}

                //omit = false;//20250227●●●●●●●●●●●●●●●●●●●●●
            }


            string[] clearedStr = { "" };
            foreach (var item in clearedStr)
            {
                xCopy = xCopy.Replace(item, "");
            }

            if (xCopy == "") return false;

            #region 檢查如《國學大師》《四庫全書》文本小注標識錯位處--每頁只檢查第一個可疑者，其他請自行注意 癸卯元宵前2日
            if (GXDS.SKQSnoteBlank && !keyinTextMode)
            {
                using (GXDS gxds = new GXDS(this))
                {
                    int sGxds = 0; int lGxds = 0;//之後還要參考 s、l 不能於此更動
                    string[] lines_xCopy = xCopy.Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                    foreach (string line_xCopy in lines_xCopy)
                    {
                        if (gxds.detectIncorrectBlankAndCurlybrackets_Suspected_aPageaTime(line_xCopy, out sGxds, out lGxds))
                        {// out 出來的是每行(item）位置，不是原來textBox1裡要給剪貼簿的文字(xCopy)內的位置
                            playSound(soundLike.warn);
                            textBox1.Select(xCopy.IndexOf(line_xCopy), line_xCopy.Length);
                            textBox1.ScrollToCaret();
                            return false;
                        }
                    }
                }
            }
            #endregion

            /* 20240913作廢
            DateTime dt = DateTime.Now;
            while (!isClipBoardAvailable_Text()) { if (DateTime.Now.Subtract(dt).TotalSeconds > 2) break; }
            try
            {
                Clipboard.SetText(xCopy);
            }
            catch (Exception)
            {
                while (!isClipBoardAvailable_Text()) { if (DateTime.Now.Subtract(dt).TotalSeconds > 2) break; }
                //playSound(soundLike.error);
                //Clipboard.SetText(xCopy);
            }
            */

            br.TextPatst2Quick_editBox = xCopy;

            BackupLastPageText(xCopy, false, false);


            if (s + l + 2 < textBox1.Text.Length)
            {
                //if (x.Substring(s + l, 1) == Environment.NewLine)//原式 20230824以前
                if (s + l + 2 <= x.Length && x.Substring(s + l, 2) == Environment.NewLine)
                {
                    x = x.Substring(s + l + 2);
                    //用新的要貼上Quit Edit的文字作為還原textBox1.Text時的備份 20230824
                    undoRecord(xCopy + Environment.NewLine + x);
                }
                else
                {
                    if (s + l >= x.Length) return false;//在自動貼入模式和手動輸入模式錯亂時防呆 20240520
                    x = x.Substring(s + l);
                    undoRecord(xCopy + x);
                }
            }



            if (s + l >= textBox1.Text.Length)
            {
                x = "";
            }
            //textBox1.Text = x;
            //}
            if (x.Length > 1)
            {
                //清除不需要的部分
                if (x.Substring(0, 2) == Environment.NewLine) x = x.Substring(2);
                string[] rTxt = { "。}}<p>" };
                int lng = rTxt[0].Length;
                if (x.Length >= lng)
                {
                    string xClr = x.Substring(0, lng);
                    int clr = xClr.IndexOf(rTxt[0]);
                    while (clr > -1 && x != "" && x.Length >= lng && x.Substring(clr + lng) != "")
                    {
                        x = x.Substring(clr + lng);
                        if (x.Length < lng) break;
                        xClr = x.Substring(0, lng);
                        clr = xClr.IndexOf(rTxt[0]);
                    }
                }
            }
            if (x.Length > 2)
            {
                #region 清除空行//前面處理跨頁小注時須有newline 判斷，故不可寫在其前而執行清除
                while (x.Substring(0, 2) == Environment.NewLine)
                {
                    x = x.Substring(2);
                    if (x.Length < 2) break;
                }
            }
            #endregion//清除空行

            if (Name == "Form1")
            {
                PauseEvents();//20240919所增，避免誤觸 textBox1_TextChanged事件程序中的 hideToNICo();方法            
                textBox1.Text = x;
                ResumeEvents();

                textBox1.SelectionStart = 0; textBox1.SelectionLength = 0;
                textBox1.ScrollToCaret();
            }
            return true;
        }


        //const string soundWarningLocation = @"c:\windows\media\Windows Foreground.wav";

        string textBox1OriginalText = "";

        List<int> charIndexList = new List<int>();
        const int charIndexListSize = 3;

        void caretPositionRecord()
        {//C# caret position record
            int charIndexToken = textBox1.SelectionStart;
            if (charIndexList.Count == 0)
            {
                charIndexList.Add(charIndexToken); return;
            }
            int sLast = charIndexList[charIndexList.Count - 1];
            if (charIndexToken != sLast)
                charIndexList.Add(charIndexToken);
            if (charIndexList.Count > charIndexListSize) charIndexList.RemoveAt(0);
        }

        int charIndexRecallTimes = charIndexListSize - 1;
        void caretPositionRecall()
        {
            if (charIndexList.Count == 0) return;
            if (charIndexRecallTimes - 1 < 0) { charIndexRecallTimes = charIndexListSize - 1; return; }
            TextBox tb = textBox1;
            int s = tb.SelectionStart;
            if (charIndexList.Count - 1 < 0 || charIndexRecallTimes - 1 < 0 ||
                charIndexRecallTimes > charIndexList.Count - 1) return;
            int sLast = charIndexList[charIndexRecallTimes - 1 > charIndexList.Count - 1 ?
                                        charIndexList.Count - 1 :
                                            charIndexRecallTimes--];
            while (sLast == s)
            {
                sLast = charIndexList[charIndexRecallTimes - 1 > charIndexList.Count - 1 ?
                charIndexList.Count - 1 :
                charIndexRecallTimes--];
            }
            tb.Select(sLast, 0);
            restoreCaretPosition(tb, sLast, 0);
            if (charIndexRecallTimes < 0) charIndexRecallTimes = charIndexListSize - 1;
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            var m = ModifierKeys; keycodeNow = e.KeyCode;
            //if (e.KeyCode != Keys.F5 && (m & Keys.Shift) != Keys.Shift) caretPositionRecord();
            if (e.KeyCode == Keys.Insert && (m & Keys.Shift) == Keys.Shift) caretPositionRecord();
            else
            if (e.KeyCode != Keys.F5) caretPositionRecord();
            if (keycodeNow == Keys.Delete) undoRecord();

            //if ((m & Keys.None) == Keys.None && e.KeyCode == Keys.Delete) undoRecord();
            //if ((m & Keys.Control) == Keys.Control && (m & Keys.Alt) == Keys.Alt && e.KeyCode == Keys.G)
            //if((int)Control.ModifierKeys ==
            //    (int)Keys.Control + (int)Keys.Alt && e.KeyCode == Keys.G)
            if ((m & Keys.Shift) == Keys.Shift && e.KeyCode == Keys.Insert && !keyinTextMode) { pasteAllOverWrite = true; dragDrop = false; }
            else pasteAllOverWrite = false;

            #region 同時按下 Ctrl + Shift + Alt 
            if (e.Control && e.Shift && e.Alt)
            {
                string x;
                switch (e.KeyCode)
                {
                    case Keys.G://Ctrl + Shift + Alt + g ： 以選取文字加上雙引號`""`檢索Google
                        x = overtypeModeSelectedTextSetting(ref textBox1);//CnText.ChangeSeltextWhenOvertypeMode(insertMode, textBox1);
                        GoogleSearch(x, true);
                        break;
                    case Keys.Y:
                        //- Ctrl + Shift + Alt + y 查[韻典網](https://ytenx.org/) y=yun（韻）的y
                        overtypeModeSelectedTextSetting(ref textBox1);
                        if (textBox1.SelectionLength == 0) return;
                        x = textBox1.SelectedText;
                        this.Invoke((MethodInvoker)delegate { Clipboard.SetText(x); });
                        Task.Run(() =>
                        {
                            LookupYTenx(x);
                        });
                        return;
                    default:
                        break;
                }
            }
            if ((m & Keys.Alt) == Keys.Alt
                && (m & Keys.Shift) == Keys.Shift && (m & Keys.Control) == Keys.Control
                && e.KeyCode == Keys.S)
            {//Alt + Shift + Ctrl + s : 小注文不換行(短於指定漢字長者)：notes_a_line_all 
                e.Handled = true; notes_a_line_all(true); return;
            }

            if (e.Control && e.Shift && e.Alt && e.KeyCode == Keys.Add)
            {//Ctrl + Shift + Alt + + 或 Ctrl + Alt + Shift + + （數字鍵盤加號） ： 同上，唯先將textBox1全選後再執行貼入；即按下此組合鍵則會並不會受插入點所在位置處影響。並翻到下一頁直接將它送去《古籍酷》OCR
                e.Handled = true;
                pagePaste2GjcoolOCR();//Ctrl + Shift + Alt + + 或 Ctrl + Alt + Shift + + （數字鍵盤加號） 
                return;
            }
            #endregion

            #region 同時按下Ctrl+Shift

            if ((m & Keys.Control) == Keys.Control
                && (m & Keys.Shift) == Keys.Shift
                && e.KeyCode == Keys.Delete)
            {//Ctrl + Shift + Delete ： 將選取文字於文本中全部清除
             //int s = textBox1.SelectionStart;
             //if (textBox1.SelectionLength > 0)
             //{
                e.Handled = true;
                clearSeltxt();
                return;
                //}
            }
            if ((m & Keys.Control) == Keys.Control
                    && (m & Keys.Shift) == Keys.Shift
                    && e.KeyCode == Keys.Up)
            {//Ctrl + Shift + ↑
                e.Handled = true;
                int s = textBox1.SelectionStart, ed = s;
                selToNewline(ref s, ref ed, textBox1.Text, false, textBox1); return;
            }
            if ((m & Keys.Control) == Keys.Control
                && (m & Keys.Shift) == Keys.Shift
                && e.KeyCode == Keys.Down)
            {//Ctrl + Shift + ↓
                e.Handled = true;
                int s = textBox1.SelectionStart, ed = s;
                selToNewline(ref s, ref ed, textBox1.Text, true, textBox1); return;
            }

            //同時按下Ctrl+Shift
            if ((m & Keys.Control) == Keys.Control && (m & Keys.Shift) == Keys.Shift)
            {
                if (e.KeyCode == Keys.Add || e.KeyCode == Keys.Oemplus || e.KeyCode == Keys.Subtract || e.KeyCode == Keys.NumPad5)
                {// Ctrl + Shift + + 
                    e.Handled = true; bool thisTopMost = this.TopMost;
                    this.TopMost = false;
                    if (!ocrTextMode) br.BringToFront("chrome");
                    string urlActive = br.ActiveTabURL_Ctext_Edit;
                    if (textBox3.Text == "" && IsValidUrl＿keyDownCtrlAdd(urlActive))
                    {
                        playSound(soundLike.done);
                        textBox3.Text = urlActive;
                    }
                    if (browsrOPMode != BrowserOPMode.appActivateByName && br.driver != null
                        && textBox3.Text != urlActive)
                    {
                        br.SwitchToCurrentForeActivateTab(ref textBox3, urlActive);
                    }
                    this.TopMost = false;
                    if (!ocrTextMode) br.BringToFront("chrome");
                    if (keyDownCtrlAdd(true))
                    {
                        //非最上層顯示以便檢視
                        //this.TopMost = false;
                        //this.WindowState = FormWindowState.Minimized;
                        //if (textBox1.Text != "" && keyinTextMode)
                        //{
                        //    pauseEvents(); textBox1.Text = ""; resumeEvents();
                        //}
                        Visible = false;
                        hideToNICo();
                        //隱藏表單後使瀏覽器取得焦點；以下皆未妥當
                        //if (browsrOPMode != BrowserOPMode.appActivateByName && br.driver != null)
                        //br.driver.Navigate().Refresh();
                        //br.driver.Navigate().GoToUrl(br.driver.Url);
                        //    //br.SwitchToCurrentForeActivateTab(ref textBox3);
                        //br.driver.SwitchTo().Window(br.driver.CurrentWindowHandle);
                    }
                    this.TopMost = thisTopMost;
                    return;
                }
            }
            //以上 //同時按下Ctrl+Shift
            #endregion


            #region 同時按下 Ctrl + Alt

            if (e.Control && e.Alt)
            {
                if (e.KeyCode == Keys.Oemplus)
                {
                    //Ctrl + Alt + = : 以選取文字檢索CTP中阮元刻《十三經注疏》本《周易正義》。便於擷取《易》學資料用。20240920
                    //if (textBox1.SelectionLength == 0) return;
                    e.Handled = true;
                    overtypeModeSelectedTextSetting(ref textBox1);
                    string url = "https://ctext.org/wiki.pl?if=gb&res=315747&searchu=" + textBox1.SelectedText;
                    textBox1.Copy();
                    switch (browsrOPMode)
                    {
                        case BrowserOPMode.appActivateByName:
                            Process.Start(url);
                            break;
                        case BrowserOPMode.seleniumNew:
                            br.LastValidWindow = br.GetCurrentWindowHandle(driver);
                            br.openNewTabWindow();
                            br.GoToUrlandActivate(url);
                            break;
                        case BrowserOPMode.seleniumGet:
                            break;
                        default:
                            break;
                    }
                    return;
                }
                if (e.KeyCode == Keys.A)
                {//Ctrl + Alt + a ： [AI太炎](https://t.shenshen.wiki/)標點 20241105
                    e.Handled = true;
                    //避免事件被終止
                    if (!_eventsEnabled) _eventsEnabled = true;
                    //避免還原記錄被終止
                    if (stopUndoRec) stopUndoRec = false;

                    string x = textBox1.Text;
                    if (x.IsNullOrEmpty()) return;
                    if (textBox1.SelectedText.IsNullOrEmpty())
                    {
                        if (textBox1.Text.IndexOf("<") == -1)
                            textBox1.SelectAll();
                        else
                        {//選取一個段落長 即 <p> 之區間 20241225 聖誕節
                            int s = textBox1.SelectionStart;
                            if (s + 1 <= x.Length)
                            {
                                if (x.Substring(s, 1) == "p")
                                    textBox1.SelectionStart--;
                                else if (x.Substring(s, 1) == ">")
                                    textBox1.SelectionStart -= 2;

                            }
                            else
                            {
                                if (s > 2 && x.Substring(s - 3, 3) == "<p>")
                                    textBox1.SelectionStart -= "<p>".Length;
                            }
                            s = x.LastIndexOf(">", textBox1.SelectionStart); int l = x.IndexOf("<", textBox1.SelectionStart);
                            s = s == -1 ? 0 : s + 1; l = l == -1 ? x.Length - s : l - s;
                            //textBox1.Select(s == 0 ? s : s + Environment.NewLine.Length, l - Environment.NewLine.Length);//+2 分段符號的長度
                            textBox1.Select(s == 0 ? s : s + Environment.NewLine.Length, l);

                        }
                    }
                    else if (textBox1.SelectedText.Length < 5)
                        textBox1.SelectAll();
                    else
                        overtypeModeSelectedTextSetting(ref textBox1);
                    while (("<p>" + Environment.NewLine).IndexOf(textBox1.SelectedText.Substring(textBox1.SelectedText.Length - 1, 1)) > -1)
                    {
                        textBox1.SelectionLength--;
                    }
                    //避開末尾的「}}」：
                    while (textBox1.SelectedText.Length > 0 && textBox1.SelectedText.Substring(textBox1.SelectedText.Length - 1, 1) == "}")
                        textBox1.SelectionLength--;
                    x = textBox1.SelectedText; string original = x, preSpaces = string.Empty; int iSpace = 0;
                    while (x.Substring(iSpace, 1) == "　")//傳回的值會缺第一行/段的的縮排空格故 20241201 感恩感恩　讚歎讚歎　南無阿彌陀佛
                    {
                        iSpace++;
                        preSpaces += "　";
                    }
                    //送出《AI太炎》標點：
                    if (AITShenShenWikiPunct(ref x))
                    {
                        x = x.Replace("“", "「").Replace("”", "」").Replace("‘", "『").Replace("’", "』");

                        //回到之前的分頁頁籤
                        if (IsDriverInvalid()) driver.SwitchTo().Window(driver.WindowHandles.Last());
                        if (driver.Url != textBox3Text)
                        {
                            for (int i = driver.WindowHandles.Count - 1; i > -1; i--)
                            {
                                driver.SwitchTo().Window(driver.WindowHandles[i]);
                                if (driver.Url == textBox3Text) break;
                            }
                        }

                        //檢查原文是否遭篡改！20241126
                        if (IsTextModified(x, original)) AvailableInUseBothKeysMouse();

                        CnText.RestoreParagraphs(original, ref x);
                        //undoRecord();
                        textBox1.SelectedText = preSpaces + CnText.BooksPunctuation(ref x, true);
                        //避免事件被終止
                        if (!_eventsEnabled) _eventsEnabled = true;
                        undoRecord();
                    }
                    AvailableInUseBothKeysMouse();
                    //避免事件被終止
                    if (!_eventsEnabled) _eventsEnabled = true;
                    return;
                }
                if (e.KeyCode == Keys.J)
                {//Ctrl + Alt + j ：以選取文字進行[《看典古籍·古籍全文檢索》](https://kandianguji.com/search) (d=dian 典；j=籍 ji) 20241018
                    if (driver == null) return;
                    e.Handled = true;
                    if (!br.IsDriverInvalid()) br.LastValidWindow = br.driver.CurrentWindowHandle;
                    TopMost = false;
                    overtypeModeSelectedTextSetting(ref textBox1);
                    string str = textBox1.SelectedText;
                    Clipboard.SetText(str);
                    Task.Run(() => { br.KanDianGuJiSearchAll(str); });
                    return;
                }
                #region Ctrl + Alt + pageup Ctrl + Alt + pagedown                
                if (e.KeyCode == Keys.PageUp || e.KeyCode == Keys.PageDown)
                {//Ctrl + Alt + pageup : 在新的分頁開啟CTP圖文對照前一頁以供檢視 20240920
                    //Ctrl + Alt + pagedown : 在新的分頁開啟CTP圖文對照下一頁以供檢視
                    if (browsrOPMode != BrowserOPMode.seleniumNew) return;
                    string url = null;
                    try
                    {
                        if (!IsDriverInvalid())
                            url = br.driver.Url;
                        else
                        {
                            if (IsValidUrl＿ImageTextComparisonPage(textBox3.Text))
                            {
                                bool found = false;
                                for (int i = driver.WindowHandles.Count - 1; i > -1; i--)
                                {
                                    driver.SwitchTo().Window(driver.WindowHandles[i]);
                                    if (ReplaceUrl_Box2Editor(driver.Url) == textBox3.Text)
                                    {
                                        found = true;
                                        break;
                                    }
                                }
                                if (!found) br.driver.SwitchTo().Window(driver.WindowHandles.Last());
                            }
                            else
                                br.driver.SwitchTo().Window(driver.WindowHandles.Last());
                            url = br.driver.Url;
                        }
                        LastValidWindow = br.driver.CurrentWindowHandle;
                    }
                    catch (Exception ex)
                    {
                        switch (ex.HResult)
                        {
                            case -2146233088:
                                if (ex.Message.StartsWith("no such window: target window already closed"))//no such window: target window already closed
                                                                                                          //from unknown error: web view not found
                                                                                                          //  (Session info: chrome = 129.0.6668.59)
                                    try
                                    {
                                        br.driver.SwitchTo().Window(br.LastValidWindow);
                                        url = br.driver.Url;
                                    }
                                    catch (Exception exx)
                                    {
                                        Console.WriteLine(exx.HResult + exx.Message);
                                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(exx.HResult + exx.Message);
                                        return;
                                    }
                                break;
                            default:
                                Console.WriteLine(ex.HResult + ex.Message);
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                                return;
                        }

                    }
                    if (url == null) return;
                    if (!IsValidUrl＿ImageTextComparisonPage(url)) return;

                    e.Handled = true;
                    int page = GetPageNumFromUrl(url);
                    if (e.KeyCode == Keys.PageUp)
                        page--;
                    else
                        page++;
                    url = br.ChangePageParameter(url, page);
                    br.LastValidWindow = br.driver.CurrentWindowHandle;
                    br.openNewTabWindow();
                    br.GoToUrlandActivate(url, true);
                    return;

                }//以上 Ctrl + Alt + pageup : 在新的分頁開啟CTP圖文對照前一頁以供檢視 20240920
                //Ctrl + Alt + pagedown : 在新的分頁開啟CTP圖文對照下一頁以供檢視
                #endregion
                if (e.KeyCode == Keys.P)
                {//Ctrl + Alt + p ： 清除textBox1選取內的`<p>` 20250215
                    e.Handled = true;
                    clearKeysParagraphSymbol();
                    return;
                }
            }

            if ((m & Keys.Control) == Keys.Control
                && (m & Keys.Alt) == Keys.Alt)//https://zhidao.baidu.com/question/628222381668604284.html
            {//https://bbs.csdn.net/topics/350010591                
                if (e.KeyCode == Keys.G || e.KeyCode == Keys.Packet)
                { e.Handled = true; return; }

            }

            //alt + Shift + f ： 將章節單位的頁面樹狀目錄收起或展開
            if (e.Shift && e.Alt && e.KeyCode == Keys.F)
            {
                OutlineTitlesCloseOpenFoldExpandSwitcher();
            }

            //Alt + Shift + s :  所有小注文都不換行。這個和我所使用的小小輸入法繁簡轉換快捷鍵有衝突，故須先停用小小輸入法才有作用。感恩感恩　南無阿彌陀佛
            /*
            if ((m & Keys.Alt) == Keys.Alt && (m & Keys.Control) == Keys.Control && (m & Keys.Shift) != Keys.Shift
                && e.KeyCode == Keys.S)
            { e.Handled = true; notes_a_line_all(false, true); return; }
            if (KeyboardInfo.getKeyStateDown(System.Windows.Input.Key.LeftAlt)  &&
                KeyboardInfo.getKeyStateDown(System.Windows.Input.Key.LeftCtrl) &&
                !KeyboardInfo.getKeyStateDown(System.Windows.Input.Key.LeftShift) &&
                e.KeyCode == Keys.S)
            { e.Handled = true; notes_a_line_all(false, true); return; }
            */
            if (e.Control && e.Alt && e.KeyCode == Keys.S)//chatGPT 202230107
                                                          //ctrl + alt + s 標題下之小注文才不換行( 會與小小輸入法預設的繁簡轉換鍵衝突，使用時請先關閉輸入法。其他快捷鍵若無作用，也多係因有較其優先之如此系統快速鍵已指定的緣故) 20230108
            { e.Handled = true; notes_a_line_all(false, true); return; }
            //以上三種皆可Alt + Shift + s :  所有小注文都不換行。

            //Ctrl + Alt + + （數字鍵盤加號） ： 同上，唯先將textBox1全選後再執行貼入；即按下此組合鍵則會並不會受插入點所在位置處影響。
            if (e.Control && e.Alt && e.KeyCode == Keys.Add)
            {
                e.Handled = true;
                ////PauseEvents();
                PressAddKeyMethodPaste2QuickEditBox();
                //textBox1.SelectAll();
                //string x = textBox1.Text;
                //if (keyDownCtrlAdd(false))
                //{
                //    if (x != br.Quickedit_data_textboxTxt)
                //    {
                //        playSound(soundLike.exam);
                //        x = br.Quickedit_data_textboxTxt;
                //    }
                //    //非同步整理OCR文本時，這行就很需要：
                //    textBox1.Text = CnText.RemarkBooksPunctuation(ref x);
                //}
                //bringBackMousePosFrmCenter();
                ////ResumeEvents();
                return;
            }

            #endregion


            #region 同時按下Alt+Shift
            //同時按下Alt+Shift
            if ((m & Keys.Alt) == Keys.Alt
                && (m & Keys.Shift) == Keys.Shift
                && e.KeyCode == Keys.S)
            {//Alt + Shift + s :  所有小注文都不換行
                e.Handled = true; notes_a_line_all(false); return;
            }

            if ((m & Keys.Alt) == Keys.Alt && (m & Keys.Shift) == Keys.Shift)
            {
                if (e.KeyCode == Keys.D1)
                {//Alt + Shift + 1 如宋詞中的換片空格，只將文中的空格轉成空白，其他如首綴前罝以明段落或標題者不轉換
                    e.Handled = true; TopMost = false;
                    SpacesBlanksInContext();
                    TopMost = true;
                    AvailableInUseBothKeysMouse();
                    return;
                }
                if (e.KeyCode == Keys.D2)
                {//Alt + Shift + 2 : 將選取區內的「<p>」取代為「|」 ，而「　」取代為「􏿽」並清除「*」且將無「|」前綴的分行符號加上「|」
                    e.Handled = true;
                    poetryFormat();
                    return;
                }
                if (e.KeyCode == Keys.D6)
                {//Alt + Shift + 6 小注文不換行
                    e.Handled = true; notes_a_line(); return;
                }
                if (e.KeyCode == Keys.Q)
                {//Alt + Shift + q : 據選取區的CJK字長以作分段（末後植入 < p >，分行則以版式常態值劃分），為非《維基文庫》版式之電子文本，如《寒山子詩集》組詩
                    e.Handled = true; markParagraphwithSelectionLen(); return;
                }
                if (e.KeyCode == Keys.R)
                {//Alt + Shift + r ：將夾注文本倒置者重整，如「{{雪電}}　{{雨霧}}　{{霽虹}}　{{雷}}」這樣的文本，改成「{{雪􏿽雨􏿽霽􏿽雷、電􏿽霧􏿽虹　　}}」
                    e.Handled = true;
                    undoRecord();
                    //當無選取時自動選取適當範圍
                    if (textBox1.SelectionLength == 0)
                    {
                        int s = textBox1.SelectionStart; Paragraph p = _document.Range(s, s).GetCurrentParagraph();
                        string currParaTxt = p.Text;
                        s = currParaTxt.IndexOf("{{");
                        textBox1.Select(s + p.Start, currParaTxt.LastIndexOf("}}") - s + 2);
                    }

                    textBox1.SelectedText = CnText.TransformText(textBox1.SelectedText);
                    undoRecord();
                    return;

                }
                if (e.KeyCode == Keys.T)
                {//Alt + Shift + t : 查中國哲學書電子化計劃網域 (以Google檢索《中國哲學書電子化計劃》) 20241024
                    string x = overtypeModeSelectedTextSetting(ref textBox1);//CnText.ChangeSeltextWhenOvertypeMode(insertMode, textBox1);
                    if (x.IsNullOrEmpty()) return;
                    e.Handled = true;
                    string url = "site: https://ctext.org/";
                    x = x.EndsWith("》") ? x.Substring(0, x.Length - 1) : x;
                    x = x.EndsWith(Environment.NewLine) ? x.Substring(0, x.Length - 2) : x;
                    x = x.EndsWith("\n") ? x.Substring(0, x.Length - 1) : x;
                    url = x + " " + url;//url置後方便按下 Ctrl + Delete 清除，以改用Google全球/全域搜尋
                    Clipboard.SetText(x);
                    if (br.driver != null)
                    {
                        br.openNewTabWindow(OpenQA.Selenium.WindowType.Tab);
                        br.driver.Navigate().GoToUrl("https://www.google.com/search?q=" + url);
                    }
                    else
                        Process.Start("https://www.google.com/search?q=" + url);
                    return;
                }
            }
            #endregion


            #region 按下Ctrl鍵
            if ((m & Keys.Control) == Keys.Control)
            //if (e.Control && !e.Shift && !e.Alt) //20240814
            /* 這裡要參照，故不能這寫↑↑↑↑ 20240815
             * if (e.KeyCode == Keys.D6)
                    {
                        if ((int)m == (int)Keys.Shift + (int)Keys.Control)
                        {
                            insX = "}}";
                        }
             */
            {//按下Ctrl鍵
             //Ctrl + v
                if (e.KeyCode == Keys.V) pasteAllOverWrite = false;
                else pasteAllOverWrite = false;

                if (e.KeyCode == Keys.F1)
                {
                    e.Handled = true;
                    noteMark(); return;
                }
                if (e.KeyCode == Keys.F12)
                {//Ctrl + F12 查詢國語辭典
                    //overtypeModeSelectedTextSetting(ref textBox1);
                    //string x = textBox1.SelectedText;
                    string x;
                    if (textBox1.SelectionLength == 0)
                        x = SelectSingleCharacter();
                    else
                        x = overtypeModeSelectedTextSetting(ref textBox1);
                    e.Handled = true;
                    if (x != "")
                    {
                        Clipboard.SetText(x);

                        if (browsrOPMode != BrowserOPMode.appActivateByName)
                        {
                            if (br.driver != null)
                            {
                                //Task.Run(() =>
                                //{
                                //if (LookupDictRevised(SelectSingleCharacter()).urlSearch == null)
                                //if (LookupDictRevised(x).urlSearch == null)
                                if (LookupDictRevised(x).Item1 == null)
                                    MessageBoxShowOKExclamationDefaultDesktopOnly("發生錯誤，請重新查詢");
                                //br.openNewTabWindow(OpenQA.Selenium.WindowType.Tab);
                                //br.driver.Navigate().GoToUrl("https://dict.revised.moe.edu.tw/search.jsp?md=1&word=" + x + "&qMd=0&qCol=1");
                                //});
                            }
                            else
                                if (MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否要執行【查詢網路辭典】？") == DialogResult.OK)
                                Process.Start(dropBoxPathIncldBackSlash + @"VS\VB\查詢國語辭典\查詢國語辭典\bin\Debug\查詢國語辭典.exe");
                        }
                        else
                            Process.Start(dropBoxPathIncldBackSlash + @"VS\VB\查詢國語辭典\查詢國語辭典\bin\Debug\查詢國語辭典.exe");
                    }
                    return;
                }
                if (e.KeyCode == Keys.Back)
                {//Ctrl + Backspace : 清除插入點之前的所有「　」或「􏿽」
                    e.Handled = true; 清除插入點之前的所有空格(); return;
                }
                if (e.KeyCode == Keys.Insert)
                {//Ctrl + Insert ：無選取時則複製插入點後一CJK字長
                    e.Handled = true;
                    if (textBox1.SelectionStart == textBox1.TextLength) return;
                    if (insertMode)
                    {
                        string x = textBox1.Text; int s = textBox1.SelectionStart;
                        if (textBox1.SelectionLength == 0)
                            Clipboard.SetText(
                            char.IsHighSurrogate(x.Substring(s, 1).ToCharArray()[0]) ? x.Substring(s, 2) : x.Substring(s, 1)
                                );
                        else
                            Clipboard.SetText(textBox1.SelectedText);
                    }
                    else
                    {
                        overtypeModeSelectedTextSetting(ref textBox1);
                        if (textBox1.SelectedText != string.Empty)
                            Clipboard.SetText(new StringInfo(textBox1.SelectedText).String);
                    }
                    return;
                }
                if (e.KeyCode == Keys.Oem3 && !e.Shift)
                {//` 或 Ctrl + ` ： 於插入點處起至「　」或「􏿽」或「|」或「<」或分段符號前止之文字加上黑括號【】//Print/SysRq 為OS鎖定不能用
                    e.Handled = true; preceded_followed_specify_symbols("【】"); return;
                }

                if (e.KeyCode == Keys.Add || e.KeyCode == Keys.Oemplus || e.KeyCode == Keys.Subtract || e.KeyCode == Keys.NumPad5)
                {//Ctrl + + （Ctrl + 數字鍵盤 +） Ctrl + -
                    e.Handled = true;
                    if (fastMode) BeginUpdate();
                    //還原放大的書圖
                    RestoreImageSize();
                    undoRecord();
                    //autoPastetoCtextQuitEditTextboxCancel = false;
                    TopMost = false;

                    betweenAddParaMarkMoveEnd();

                    if (lines_perPage == 0)
                    {
                        lines_perPage = countLinesPerPage(textBox1.Text.Substring(0, textBox1.SelectionStart + textBox1.SelectionLength));
                    }

                    if (!ocrTextMode) br.BringToFront("chrome");
                    if (e.KeyCode == Keys.Subtract)
                    {// Ctrl + -（數字鍵盤） 會重設以插入點位置為頁面結束位國


                        resetPageTextEndPositionPasteToCText();
                        if (!autoPaste2QuickEdit && Name == "Form1") AvailableInUseBothKeysMouse();
                        EndUpdate();
                        return;//因為函式中已有keyDownCtrlAdd
                    }

                    ////if (keyDownCtrlAdd(false))  if (textBox1.Text != "") { pauseEvents(); textBox1.Text = ""; resumeEvents(); }
                    //keyDownCtrlAdd(false);// if (textBox1.Text != "") { pauseEvents(); textBox1.Text = ""; resumeEvents(); }

                    if (autoPaste2QuickEdit)
                    {
                        playSound(soundLike.press, true); altA_predictEndofPageRange();
                    }
                    keyDownCtrlAdd(false);

                    EndUpdate();
                    TopMost = true;
                    return;
                }

                //Ctrl + 0, Ctrl + 9, Ctrl + 8, Ctrl + 7
                if (e.KeyCode == Keys.D0 || e.KeyCode == Keys.D9 || e.KeyCode == Keys.D8 || e.KeyCode == Keys.D7 || e.KeyCode == Keys.D6)
                {
                    e.Handled = true;
                    int s = textBox1.SelectionStart, l = textBox1.SelectionLength;
                    string insX = "", x = textBox1.Text;
                    //if (textBox1.SelectedText != "")
                    //    x = x.Substring(0, s) + x.Substring(s + l);
                    undoRecord();
                    #region 《十三經注疏》四合一版區塊識識別用                    
                    if (e.KeyCode == Keys.D0)
                    {
                        //insX = Environment.NewLine + "　" + Environment.NewLine +
                        //    "　" + Environment.NewLine +
                        //    "　" + Environment.NewLine +
                        //    "　" + Environment.NewLine;
                        insX = "　" + Environment.NewLine +
                            "　" + Environment.NewLine +
                            "　" + Environment.NewLine + "　";
                    }
                    if (e.KeyCode == Keys.D9)
                    {
                        //insX = Environment.NewLine + "　" + Environment.NewLine +
                        //    "　" + Environment.NewLine;
                        insX = "　" + Environment.NewLine + "　";
                    }
                    if (e.KeyCode == Keys.D8)
                    {
                        //insX = Environment.NewLine + "　" + Environment.NewLine;
                        insX = "　";
                    }

                    if (s + 2 <= x.Length)
                    {
                        if (x.Substring(s, 2) == Environment.NewLine)
                        {
                            insX = Environment.NewLine + insX;
                        }
                    }
                    if (s - 2 >= 0)
                    {
                        if (x.Substring(s - 2, 2) == Environment.NewLine)
                        {
                            insX += Environment.NewLine;
                        }
                    }
                    if (s + 2 <= x.Length && s - 2 >= 0)
                    {
                        if (x.Substring(s, 2) != Environment.NewLine && x.Substring(s - 2, 2) != Environment.NewLine)
                        {
                            insX = Environment.NewLine + insX + Environment.NewLine;
                        }

                    }
                    //以上需要間隔，蓋為《十三經注疏》四合一版間隔區別四區塊專用耳。
                    #endregion
                    // Ctrl + 6 Ctrl + Shift+ 6 Ctrl + 7 並不需要間隔
                    if (e.KeyCode == Keys.D7)
                    {
                        insX = "。}}";
                    }
                    if (e.KeyCode == Keys.D6)
                    {
                        if ((int)m == (int)Keys.Shift + (int)Keys.Control)
                        {
                            insX = "}}";
                        }
                        if (m == Keys.Control)
                        {
                            insX = "{{";
                        }
                    }
                    stopUndoRec = true;
                    //insertWords(insX, textBox1, x);
                    insertWords(insX, textBox1);
                    //x = x.Substring(0, s) + insX + x.Substring(s);
                    //textBox1.Text = x;
                    //textBox1.SelectionStart = s + insX.Length;
                    //textBox1.ScrollToCaret();
                    stopUndoRec = false;
                    return;
                }

                //Ctrl + c ：若無選取，則複製textBox1內的內容
                if (e.KeyCode == Keys.C)
                {
                    if (textBox1.Text == string.Empty) return;
                    e.Handled = true;
                    Clipboard.SetText(textBox1.Text);
                    return;
                }


                //Ctrl + y
                if (e.KeyCode == Keys.Y)
                {//還原功能
                    e.Handled = true;
                    redoTextBox(textBox1);
                    return;
                }

                //Ctrl + z
                if (e.KeyCode == Keys.Z)
                {//還原功能
                    e.Handled = true;
                    saveText();//備份以便按下F5鍵還原操作此還原方法前的文本
                    undoTextBox(textBox1);
                    //if (undoTextBox1Text.Last<string>() != textBox1.Text)
                    //    undoRecord();
                    return;
                }


                //Ctrl + h
                if (e.KeyCode == Keys.H)
                //if ((m & Keys.Control) == Keys.Control && e.KeyCode == Keys.H)
                {
                    //不知為何，就是會將插入點前一個字元給刪除,即使有以下此行也無效
                    e.Handled = true;
                    undoRecord();
                    overtypeModeSelectedTextSetting(ref textBox1);
                    textBox1OriginalText = textBox1.Text; selStart = textBox1.SelectionStart; selLength = textBox1.SelectionLength;
                    if (textBox1.SelectedText != string.Empty)
                        try
                        {
                            Clipboard.SetText(textBox1.SelectedText);
                        }
                        catch (Exception)
                        {
                        }
                    ////插件/取代模式不同處理
                    //char nextChar;
                    //if (selStart + selLength + 1 <= textBox1.TextLength)
                    //{
                    //    nextChar = textBox1.Text.Substring(selStart, selLength + 1).ToArray()[0];
                    //    selLength = insertMode ? selLength :
                    //        (selLength + 1 > textBox1.TextLength) ? selLength :
                    //        char.IsHighSurrogate(nextChar) || nextChar == Environment.NewLine.ToArray()[0] ?
                    //        textBox1.SelectionLength += 2 : ++textBox1.SelectionLength;
                    //}
                    textBox4.Focus();
                    return;
                }

                if (e.KeyCode == Keys.K)
                {// 依選取文字取得目前URL加該選取字為該頁之關鍵字的連結。如欲在此頁中標出「𢔶」字，即為：
                    /// https://ctext.org/library.pl?if=gb&file=36575&page=53#𢔶
                    /// Ctrl + k
                    e.Handled = true;
                    //CnText.ChangeSeltextWhenOvertypeMode(insertMode, textBox1);//這個只是取得文字，不會改變選取行為
                    overtypeModeSelectedTextSetting(ref textBox1);//暫時恢復，再觀察。→因為之前常會先複製該關鍵字詞，就已會有選取範圍
                    if (textBox1.SelectionLength > 0)
                    {
                        string w = textBox1.SelectedText, url = textBox3.Text;
                        string x = br.GetPageUrlKeywordLink(w, url, true);
                        if (x != string.Empty) Clipboard.SetText(x);
                        else MessageBoxShowOKExclamationDefaultDesktopOnly("無法取得具關鍵字的連結字串，請檢查！");
                    }
                    return;
                }

                //Ctrl + q
                if (e.KeyCode == Keys.Q)
                {
                    e.Handled = true;
                    splitLineByFristLen(); return;
                }

                //Ctrl +\
                if (e.KeyCode == Keys.OemBackslash || e.KeyCode == Keys.Oem5)
                {
                    e.Handled = true;
                    clearNewLinesAfterCaret();
                    return;
                }

                //Ctrl + ↑ Ctrl + ↓
                if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down)
                {/*Ctrl + ↑：從插入點開始向前移至上一段尾
                  * Ctrl + ↓：從插入點開始向後移至這一段末（無分段則不移動）*/
                    keyDownCtrlAltUpDown(e);
                    return;
                }

                //Ctrl + [,Ctrl + ]
                if (e.KeyCode == Keys.OemCloseBrackets || e.KeyCode == Keys.OemOpenBrackets)
                {/*Ctrl + [：從插入點開始向前移至{{前
                    Ctrl + ]：從插入點開始向後移至}}後*/
                    e.Handled = true;
                    string x = textBox1.Text;
                    if (x.IndexOf("{{") == -1 && x.IndexOf("}}") == -1)
                    {
                        MessageBox.Show("not found {{ or  }} ");
                        return;
                    }
                    int s = textBox1.SelectionStart;
                    if (e.KeyCode == Keys.OemCloseBrackets)
                        s = x.IndexOf("}}", s + 1) + 2;
                    else
                    {
                        if (s > 0)
                        {
                            s = x.LastIndexOf("{{", s - 1);
                        }
                    }
                    if (s > -1)
                        textBox1.SelectionStart = s;
                    else
                        MessageBox.Show("not found!");
                    textBox1.ScrollToCaret();
                    return;
                }

                //Ctrl + ← Ctrl + →
                if (e.KeyCode == Keys.Left || e.KeyCode == Keys.Right)
                {/*Ctrl + →：插入點若在漢字中,從插入點開始向後移至任何非漢字前(即漢字後) 反之亦然
                  * Ctrl + ←：：插入點若在漢字中,從插入點開始向後移至任何非漢字後(即漢字前) 反之亦然*/
                    string x = textBox1.Text;
                    int s = textBox1.SelectionStart, ss = s;
                    int l; bool isIPCharHanzi;
                    string endStr = "}>" + Environment.NewLine;
                    e.Handled = true;
                    if (e.KeyCode == Keys.Left)
                    {//Ctrl  + ← Ctrl + 向左鍵
                        if (s - 1 > 0)
                        {
                            if (endStr.IndexOf(textBox1.Text.Substring(s - 1, 1)) > -1) s--;
                            isIPCharHanzi = isChineseChar(x.Substring(s - 1, 1), true) == 0 ? false : true;
                            if (isIPCharHanzi) l = findNotChineseCharFarLength(x.Substring(0, s), false);
                            else l = findChineseCharFarLength(x.Substring(0, s), false);
                            if (l != -1)
                            {
                                s = s - l + 1;

                                if (endStr.IndexOf(textBox1.Text.Substring(s, 1)) > -1)
                                {
                                    if (textBox1.Text.Substring(s, 2) == Environment.NewLine)
                                        s += 2;
                                    //else
                                    //s++;
                                }
                                ////if (textBox1.Text.Substring(s - 1, 2)!= "􏿽")
                                ////{
                                //s--;////////////////////新增以除錯的。還原「s = s - l + 1;」多加之1
                                ////}
                                textBox1.Select(s, 0);
                                restoreCaretPosition(textBox1, s, 0);//textBox1.ScrollToCaret();
                                e.Handled = true;
                                //return;
                            }
                        }
                    }
                    else
                    {// Ctrl + →  Ctrl + 向右鍵
                        if (s + 1 <= x.Length)
                        {
                            int si = 0;
                            if (char.IsLowSurrogate(x.Substring(s, 1).ToCharArray()[0])) si++;
                            isIPCharHanzi = isChineseChar(x.Substring(s, ++si), true) == 0 ? false : true;
                            //s += si;

                            ////s++;
                            //if (char.IsLowSurrogate(x.Substring(s, 1).ToCharArray()[0])) s++;
                            //isIPCharHanzi = isChineseChar(x.Substring(s, 1), true) == 0 ? false : true;
                        }
                        else
                            isIPCharHanzi = false;
                        if (isIPCharHanzi) l = findNotChineseCharFarLength(x.Substring(s), true);
                        else l = findChineseCharFarLength(x.Substring(s), true);
                        if (l != -1)
                        {
                            s = s + l - 1;
                            //if ("。，、；：？！「」『』《》〈〉".IndexOf(textBox1.Text.Substring(s, 1)) > -1) s++;
                            if (x.Substring(s, 1) == "}") s += 2;
                            if (s + 3 <= x.Length)
                            { if (x.Substring(s, 3) == "<p>") s += 3; }
                            else
                                s = x.Length;

                            ////if (textBox1.Text.Substring(s + 1, 2) == "􏿽")
                            ////{
                            //s++;////////////////////新增以除錯的。還原「s = s + l - 1;」多減之1
                            ////}
                            textBox1.Select(s, 0);
                            restoreCaretPosition(textBox1, s, 0);//textBox1.ScrollToCaret();
                            e.Handled = true;
                            //return;
                        }
                    }
                    if ((m & Keys.Control) == Keys.Control && (m & Keys.Shift) == Keys.Shift)
                    {// Ctrl+ Shift + ←  Ctrl+ Shift + → 選取文字 ，並將空格「　」與空白「􏿽」對轉
                        textBox1.Select(ss, s - ss);
                        //if (textBox1.SelectedText.Replace("　", "") == "")
                        {
                            //將空格改成空白
                            undoRecord();
                            stopUndoRec = true;
                            s = textBox1.SelectionStart; l = textBox1.SelectionLength; x = textBox1.Text;
                            switch (e.KeyCode)
                            {
                                case Keys.Left:
                                    while (s > 1)
                                    {
                                        if (x.Substring(s - 2, 2) == "􏿽")
                                        { s -= 2; l += 2; }
                                        else if (x.Substring(s - 1, 1) == "　")
                                        { s--; l++; }
                                        else
                                            break;
                                    }
                                    break;
                                case Keys.Right:
                                    while (s + l + 2 <= x.Length)
                                    {
                                        if (x.Substring(s + l, 2) == "􏿽")
                                            l += 2;
                                        else if (x.Substring(s + l, 1) == "　")
                                            l++;
                                        else
                                            break;
                                    }
                                    break;
                                    //default:
                            }
                            textBox1.Select(s, l);
                            if (textBox1.SelectedText.IndexOf("　") > -1)
                                textBox1.SelectedText = textBox1.SelectedText.Replace("　", "􏿽");
                            else if (textBox1.SelectedText.IndexOf("􏿽") > -1)
                                textBox1.SelectedText = textBox1.SelectedText.Replace("􏿽", "　");
                            stopUndoRec = false;
                            if (e.KeyCode == Keys.Left)
                            {
                                if (textBox1.Text.Substring(s, 2) == "}}") s += 2;
                                textBox1.Select(s, 0);
                            }
                        }
                    }
                    return;
                }

                //Ctrl + . // Ctrl + ,
                if (e.KeyCode == Keys.OemPeriod || e.KeyCode == Keys.Oemcomma)
                {
                    e.Handled = true;
                    int s = textBox1.SelectionStart; string x = textBox1.Text;
                    string findwhat;
                    if (e.KeyCode == Keys.OemPeriod)
                        findwhat = ">";
                    else
                        findwhat = "<";
                    int p = x.IndexOf(findwhat, s + 1);
                    if (p > -1)
                    {
                        int l = 0;
                        if (findwhat == ">")
                            l = 1;
                        textBox1.Select(p + l, 0);
                        textBox1.ScrollToCaret();
                    }
                    else
                        MessageBox.Show("not found!");
                    return;
                }

                if (e.KeyCode == Keys.Delete && !e.Shift && !e.Alt)
                {//Ctrl + Delete ： 將插入點所在位置之後的文字一律清除(Ctrl + z 還原功能支援)
                    //> 如果插入點後是空格（space）、空白（􏿽）、段落標記<p>或*，則清除到非此諸標記符號，否則就一律清除
                    e.Handled = true;
                    int s = textBox1.SelectionStart, l = textBox1.SelectionLength; string x = textBox1.Text;
                    if (s < textBox1.TextLength && l == 0)
                    {
                        if (x[s] == '*')//if (x.Substring(s, 1) == "*")
                        {
                            while (x[s + (++l)] == '*') { }
                            textBox1.Select(s, l);
                        }
                        else if (x.Substring(s, 1) == "　")
                        {
                            while (s + l + 1 <= textBox1.TextLength && textBox1.Text.Substring(s + l, 1) == "　")
                            {
                                textBox1.Select(s + l++, 1);
                                //textBox1.SelectedText = string.Empty;
                            }
                            textBox1.Select(s, l);
                        }
                        else if (s < x.Length - 1 && x.Substring(s, 2) == "􏿽")//2="􏿽".Length
                        {
                            while (textBox1.TextLength >= s + l + 2 && textBox1.Text.Substring(s + l, 2) == "􏿽")
                            {
                                textBox1.Select(s + l, 2); l += 2;
                                //textBox1.SelectedText = string.Empty;
                            }
                            textBox1.Select(s, l);
                        }
                        //清除<p> 及成對 < > 的一切語法
                        else if (s < x.Length - 1 && x.Substring(s, 1) == "<")//"<".Length
                        {
                            while (textBox1.TextLength >= s + l + 1 && textBox1.Text.Substring(s + l, 1) != ">")//1=">".Length
                            {
                                textBox1.Select(s + l, 1); l += 1;
                                //textBox1.SelectedText = string.Empty;
                            }
                            textBox1.Select(s, l + ">".Length);
                        }
                        else if (s + 1 < x.Length && x.Substring(s, 1) == "〖")
                        {//清除〖〗的全部內容
                            clearWhiteSquareBracketsRange(s, x, false);
                        }
                        else
                        {
                            textBox1.Select(textBox1.SelectionStart, textBox1.TextLength - textBox1.SelectionStart);
                        }
                        undoRecord(); stopUndoRec = true; PauseEvents();
                        textBox1.SelectedText = string.Empty;
                        stopUndoRec = false; undoRecord(); ResumeEvents();
                    }
                    return;
                }

            }//以上 Ctrl

            #endregion

            #region 按下Shift鍵

            //按下Shift鍵
            if ((m & Keys.Shift) == Keys.Shift)
            {
                if (e.KeyCode == Keys.Enter)
                {//Shift + Enter（ 與 Ctrl + 滑鼠左鍵 同） ： 新行、折行、分行，並可依插入點所在行之縮排情形，自動補上前綴全形空格
                    e.Handled = true;
                    newLine();
                    SendKeys.Send("{backspace}");
                    return;
                }
                if (e.KeyCode == Keys.F3)
                {//Shift + F3
                    e.Handled = true;
                    int foundwhere;
                    if (textBox1.SelectionLength == 0) overtypeModeSelectedTextSetting(ref textBox1);
                    string findword = textBox1.SelectionLength == 0 ? lastFindStr : textBox1.SelectedText;
                    if (findword == "") findword = textBox2.Text;
                    if (findword != "")
                    {
                        if (textBox1.SelectionStart > 0)
                        {
                            int start = textBox1.SelectionStart - 1; string x = textBox1.Text;
                            foundwhere = x.LastIndexOf(findword, start, StringComparison.Ordinal);
                            if (foundwhere == -1)
                            {
                                MessageBox.Show("not found next!"); return;
                            }
                            textBox1.SelectionStart = foundwhere;
                            textBox1.SelectionLength = findword.Length;
                            textBox1.ScrollToCaret();
                        }
                    }
                    if (findword != "") lastFindStr = findword;
                    return;
                }//以上 Shift + F3

                if (e.KeyCode == Keys.F5)
                {//Shift + F5 ： 在textBox1 回到上1次插入點（游標）所在處（且與最近「caretPositionListSize」次瀏覽處作切換，如 MS Word）
                    e.Handled = true; caretPositionRecall(); return;
                }//以上 Shift + F5

                if (e.KeyCode == Keys.F7)
                {//Shift + F7 每行凸排
                    e.Handled = true;
                    #region 前端介面來做的
                    #region 如果插入點剛好在行/段前 20250124
                    if (textBox1.SelectionLength == 0 && textBox1.SelectionStart > 0
                        && textBox1.SelectionStart < textBox1.TextLength - Environment.NewLine.Length &&
                        textBox1.Text.Substring(textBox1.SelectionStart, Environment.NewLine.Length) == Environment.NewLine) textBox1.SelectionStart--;
                    #endregion
                    int s = textBox1.SelectionStart, l = textBox1.SelectionLength; string x = textBox1.Text;
                    expandSelectedTextRangeToWholeLinePara(s, l, x);
                    s = textBox1.SelectionStart;
                    l = textBox1.SelectionLength;
                    x = textBox1.SelectedText;
                    #endregion 前端介面來做的                    
                    #region 後端伺服                    
                    //outdent_ConvexRow();//原式，純操作介面操作者
                    //Range range = new Document(ref x).Range();
                    Range range = new Document(x).Content;
                    outdent_ConvexRow(ref range);
                    #endregion 後端伺服

                    #region 前端介面
                    textBox1.SelectedText = range.Text;
                    undoRecord();
                    expandSelectedTextRangeToWholeLinePara(s, l - (l - range.Text.Length) - Environment.NewLine.Length, textBox1.Text);
                    if (!Active)
                        bringBackMousePosFrmCenter();
                    return;
                    #endregion 前端介面
                }//以上 Shift + F7

                if (e.KeyCode == Keys.F8)//20230929實歲五十一之生日
                {//Shift + F8
                    e.Handled = true;
                    string x = textBox1.Text; int s = textBox1.SelectionStart, p = x.IndexOf(Environment.NewLine, s) == -1 ? x.Length : x.IndexOf(Environment.NewLine, s),
                        preP = x.LastIndexOf(Environment.NewLine, s) == -1 ? 0 : x.LastIndexOf(Environment.NewLine, s);
                    undoRecord(); PauseEvents(); stopUndoRec = true;
                    if (preP < p)
                    {
                        p = x.IndexOf("。<p>", preP, p - preP);
                        if (p > -1)
                        {//清除「。<p>」中的句號 20231119
                            textBox1.Text = x.Substring(0, p) + x.Substring(p + "。".Length);
                            textBox1.Select(s, 0); textBox1.ScrollToCaret();
                        }
                    }
                    keysTitleCodeAndPreWideSpace();
                    ResumeEvents(); stopUndoRec = false;
                    Clipboard.SetText(textBox1.Text);//通常標識後是要再重標點，如書名等 20240306
                    return;
                }//以上 Shift + F8



            }//以上 Shift
            #endregion




            #region 按下Alt鍵            
            //按下Alt鍵
            if ((m & Keys.Alt) == Keys.Alt)//⇌ if (Control.ModifierKeys == Keys.Alt)
            {
                if (e.KeyCode == Keys.Enter)//Alt + Enter（ 與 Ctrl + 滑鼠左鍵 同） ： 新行、折行、分行，並可依插入點所在行之縮排情形，自動補上前綴全形空格
                {
                    e.Handled = true;
                    newLine();
                    return;
                }
                if (e.KeyCode == Keys.F1)// Alt + F1
                {
                    e.Handled = true;
                    keySymbols("■");
                    return;
                }
                if (e.KeyCode == Keys.F2)// Alt + F2
                {
                    e.Handled = true;
                    keySymbols("□");
                    return;
                }
                if (e.KeyCode == Keys.F12)
                {//Alt + F12 查詢《異體字字典》
                    e.Handled = true;
                    var urls = LookupDictionary_of_ChineseCharacterVariants(SelectSingleCharacter());
                    //Console.WriteLine(urls.urlSearch + Environment.NewLine + urls.urlResult);
                    return;
                }
                if (e.KeyCode == Keys.Multiply)// Alt + *
                {
                    e.Handled = true; 歐陽文忠公集_集古錄跋尾校語專用(); return;
                }

                if (e.KeyCode == Keys.OemPeriod)
                {// Alt + . //插入書名、篇名號中間符號
                    //insertWords("·", textBox1, textBox1.Text);
                    insertWords("·", textBox1);
                    e.Handled = true;
                    return;
                }

                if (e.KeyCode == Keys.D1)//D1=Menu?
                {//Alt + 1 : 鍵入本站制式留空空格標記「􏿽」：若有選取則取代全形空格「　」為「􏿽」
                    e.Handled = true;
                    keysSpacesBlank();
                    if (!Active) AvailableInUseBothKeysMouse();
                    return;
                }

                if (e.KeyCode == Keys.D2)
                {//Alt + 2 : 鍵入全形空格「　」
                    e.Handled = true;
                    keysSpaces();
                    return;
                }
                if (e.KeyCode == Keys.D3)
                {//Alt + 3 : 鍵入全形空格「◯」
                    e.Handled = true;
                    insertWords("◯", textBox1);
                    return;
                }
                if (e.KeyCode == Keys.D4)
                {//Alt + 4 : 新增【四部叢刊造字對照表】資料並取代其造字
                    e.Handled = true;
                    addData四部叢刊造字對照表andReplace();
                    return;
                }
                if (e.KeyCode == Keys.D6 || e.KeyCode == Keys.D7)
                {//Alt + 6 、 Alt + 7 : 鍵入 「"}}"+ newline +"{{"」
                    e.Handled = true;
                    //自動清除插入點後的分行/段符號
                    int s1 = textBox1.SelectionStart + 1;
                    if (s1 <= textBox1.TextLength && textBox1.Text.Substring(s1 - 1, 1) == Environment.NewLine)
                    {
                        textBox1.Select(s1 - 1, 1);
                        textBox1.SelectedText = string.Empty;
                    }
                    //insertWords("}}" + Environment.NewLine + "{{", textBox1, textBox1.Text);
                    insertWords("}}" + Environment.NewLine + "{{", textBox1);
                    return;
                }
                if (e.KeyCode == Keys.D8)
                {//Alt + 8 : 鍵入 「　　*」

                    e.Handled = true;
                    //insertWords("　　*", textBox1, textBox1.Text);
                    insertWords("　　*", textBox1);
                    return;
                }

                #region alt + 9 、alt + 0、alt + u、alt + y、alt + i
                //20240810 creedit with Copilot大菩薩：C# Windows.Forms 程式碼中的取代模式處理 https://sl.bing.net/jOCLQeh6cyi
                if (e.Alt && (e.KeyCode == Keys.D9 || e.KeyCode == Keys.D0 || e.KeyCode == Keys.U || e.KeyCode == Keys.Y || e.KeyCode == Keys.I))
                {
                    e.Handled = true;
                    string insX = GetInsertSymbol(e.KeyCode);
                    InsertSymbol(insX);
                    return;
                }

                string GetInsertSymbol(Keys keyCode)
                {
                    switch (keyCode)
                    {
                        case Keys.D9: return "「";
                        case Keys.D0: return "『";
                        case Keys.U: return "《";
                        case Keys.Y: return "〈";
                        case Keys.I: return GetClosingSymbol();
                        default: return "》";
                    }
                }

                string GetClosingSymbol()
                {
                    int s = textBox1.SelectionStart;
                    if (s > 0)
                    {
                        string xPrevious = textBox1.Text.Substring(0, s);
                        const string symbol = "{（〈《「『』」》〉）";
                        string whatSymbolPrefix = "";
                        string xChk = ""; bool chk = false; //bool closeFlag = false;
                        for (int i = xPrevious.Length - 1; i > -1; i--)
                        {
                            whatSymbolPrefix = xPrevious.Substring(i, 1);
                            if (symbol.IndexOf(whatSymbolPrefix) > -1)
                            {
                                xChk = xPrevious.Substring(0, i + 1); chk = true;
                                break;
                            }
                        }
                        if (chk)
                        {
                            const string symbolPairChk = "（〈《「『）〉》」』";
                            const string symbolPairChkClose = "）〉》」』";
                            int sFirst = -1;
                            List<string> sPairOpenFirst = new List<string>();
                            for (int i = xChk.Length - 1; i > -1; i--)
                            {
                                sFirst = symbolPairChk.IndexOf(xChk[i]);
                                bool sPairOpenFirstContained = sPairOpenFirst.Contains(xChk[i].ToString());
                                if (sFirst > -1 && !sPairOpenFirstContained)
                                {
                                    string insX = symbolPairChk[sFirst].ToString();
                                    if (symbolPairChkClose.IndexOf(xChk[i]) == -1)
                                    {
                                        if (sPairOpenFirst.Count == 0 || !sPairOpenFirst.Contains(xChk[i].ToString()))
                                        {
                                            return symbolPairChkClose[sFirst].ToString();
                                        }
                                    }
                                    else
                                    {
                                        string sPOF = symbolPairChk[symbolPairChkClose.IndexOf(insX)].ToString();
                                        if (sPairOpenFirst.Count == 0 || sPairOpenFirst.Contains(sPOF))
                                        {
                                            sPairOpenFirst.Add(sPOF);
                                        }
                                        continue;
                                    }
                                }
                            }
                            return "》";
                        }
                        else
                        {
                            return "》";
                        }
                    }
                    else
                    {
                        return "》";
                    }
                }

                void InsertSymbol(string insX)
                {
                    string x = textBox1.Text;
                    //insertWords(insX, textBox1, x);
                    insertWords(insX, textBox1);
                    //如果是」、「且後面是接著英文的雙引號 " 就直接蓋過去（直接取代掉） 20240810
                    //if ("「」".IndexOf(insX) > -1 && textBox1.SelectionStart + 1 <= textBox1.TextLength && textBox1.Text.Substring(textBox1.SelectionStart, 1) == "\"")
                    if (textBox1.SelectionStart + 1 <= textBox1.TextLength && "\"\'".IndexOf(textBox1.Text.Substring(textBox1.SelectionStart, 1)) > -1)
                    {
                        textBox1.Select(textBox1.SelectionStart, 1);
                        textBox1.SelectedText = string.Empty;
                    }
                }

                //if (e.KeyCode == Keys.D9 || e.KeyCode == Keys.D0 || e.KeyCode == Keys.U || e.KeyCode == Keys.Y || e.KeyCode == Keys.I)
                //{/* Alt + 9 : 鍵入 「 
                //  * Alt + 0 : 鍵入 『 
                //  * Alt + u : 鍵入 《 
                //  * Alt + y : 鍵入 〈 
                //  * Alt + i : 鍵入 》（如 MS Word 自動校正(如在「選項>印刷樣式」中的設定值)，會依前面的符號作結尾號（close），如前是「〈」，則轉為「〉」……）*/
                //    e.Handled = true;
                //    string insX = "", x = textBox1.Text;
                //    if (e.KeyCode == Keys.D9) { insX = "「"; goto insert; }
                //    if (e.KeyCode == Keys.D0) { insX = "『"; goto insert; }
                //    if (e.KeyCode == Keys.U) { insX = "《"; goto insert; }
                //    if (e.KeyCode == Keys.Y) { insX = "〈"; goto insert; }
                //    if (e.KeyCode == Keys.I)
                //    {
                //        int s = textBox1.SelectionStart;
                //        if (s > 0)
                //        {
                //            string xPrevious = x.Substring(0, s);
                //            const string symbol = "{（〈《「『』」》〉）";
                //            string whatSymbolPrefix = "";
                //            string xChk = ""; bool chk = false; bool closeFlag = false;
                //            for (int i = xPrevious.Length - 1; i > -1; i--)
                //            {
                //                whatSymbolPrefix = xPrevious.Substring(i, 1);
                //                if (symbol.IndexOf(whatSymbolPrefix) > -1)
                //                {
                //                    xChk = xPrevious.Substring(0, i + 1); chk = true;
                //                    break;
                //                }
                //            }
                //            if (chk)//需要檢查誰沒配對
                //            {
                //                const string symbolPairChk = "（〈《「『）〉》」』";
                //                const string symbolPairChkClose = "）〉》」』";
                //                int sFirst = -1;
                //                List<string> sPairOpenFirst = new List<string>();
                //                for (int i = xChk.Length - 1; i > -1; i--)
                //                {
                //                    sFirst = symbolPairChk.IndexOf(xChk[i]);
                //                    bool sPairOpenFirstContained = sPairOpenFirst.Contains(xChk[i].ToString());
                //                    if (sFirst > -1 && !sPairOpenFirstContained)
                //                    {
                //                        insX = symbolPairChk[sFirst].ToString();
                //                        if (symbolPairChkClose.IndexOf(xChk[i]) == -1)
                //                        {//如果是open 
                //                            if (sPairOpenFirst.Count == 0 ||
                //                                !sPairOpenFirst.Contains(xChk[i].ToString()))
                //                            {
                //                                insX = symbolPairChkClose[sFirst].ToString();
                //                                closeFlag = true;
                //                                break;
                //                            }
                //                        }
                //                        else
                //                        {//如果是close,取得其配對的 open
                //                            string sPOF = symbolPairChk[
                //                                symbolPairChkClose.IndexOf(insX)].ToString();
                //                            if (sPairOpenFirst.Count == 0 || !sPairOpenFirst.Contains(sPOF))
                //                            {
                //                                sPairOpenFirst.Add(sPOF);
                //                            }
                //                            continue;
                //                        }

                //                    }

                //                }//end of for loop 
                //                if (!closeFlag)
                //                {
                //                    insX = "》";
                //                }
                //            }
                //            else
                //            {
                //                insX = "》";

                //            }

                //        }
                //        else
                //        {//pick up the close symbol according to the open one
                //            switch (insX)
                //            {
                //                case "{":
                //                    insX = "}}";
                //                    break;
                //                case "（":
                //                    insX = "）";
                //                    break;
                //                case "〈":
                //                    insX = "〉";
                //                    break;
                //                case "《":
                //                    insX = "》";
                //                    break;
                //                case "「":
                //                    insX = "」";
                //                    break;
                //                case "『":
                //                    insX = "』";
                //                    break;
                //                default:
                //                    insX = "》";
                //                    break;
                //            }


                //        }
                //    }
                //    else
                //    {
                //        insX = "》";
                //    }
                //insert:
                //    insertWords(insX, textBox1, x);
                //    //若是「、」則取代"
                //    if ("「」".IndexOf(insX) > -1 && textBox1.SelectionStart + 1 <= textBox1.TextLength && textBox1.Text.Substring(textBox1.SelectionStart + 1, 1) == "\"")
                //    { textBox1.Select(textBox1.SelectionStart + 1, 1); textBox1.SelectedText = string.Empty; }

                //    return;
                //}
                #endregion

                if (e.KeyCode == Keys.A)
                {//Alt + a : 通常是用在自動輸入模式時根據上一次判斷的頁尾來自動貼入本頁內容
                    e.Handled = true;
                    TopMost = false; bool foundTab = false;
                    if (IsDriverInvalid())
                    {
                        for (int i = driver.WindowHandles.Count - 1; i > -1; i--)
                        {
                            driver.SwitchTo().Window(driver.WindowHandles[i]);
                            if (driver.Url == textBox3.Text)
                            {
                                foundTab = true; break;
                            }
                        }
                        if (!foundTab) br.driver.SwitchTo().Window(driver.WindowHandles.Last());
                    }
                    else
                        br.driver.SwitchTo().Window(driver.CurrentWindowHandle);
                    if (fastMode) BeginUpdate();
                    //還原放大的書圖
                    //if (autoPaste2QuickEdit)
                    RestoreImageSize();
                    if (!keyinTextMode)
                    { playSound(soundLike.press, true); altA_predictEndofPageRange(); }
                    keyDownCtrlAdd(false);
                    RestoreImageSize();
                    if (!autoPaste2QuickEdit && !isSKQSFrontPage(textBox1.Text)) AvailableInUseBothKeysMouse();
                    EndUpdate();
                    return;
                }

                if (e.KeyCode == Keys.B)
                {//Alt + b： 訂正註文中空白錯亂的文本。如「{{帝和霍王以下 句亡}}」訂正為「{{帝 霍王以下和句亡}}」，將半形空格與其前半對應的漢字對調。 20250219
                    e.Handled = true;
                    runCorrectNoteBlankContent();
                }
                if (e.KeyCode == Keys.C)
                {//Alt + c ：以所選之詞（不能少於2字）檢索《漢語大詞典》 https://ivantsoi.myds.me/web/hydcd/search.html
                    e.Handled = true;
                    overtypeModeSelectedTextSetting(ref textBox1);
                    LookupHYDCD(textBox1.SelectedText);
                    return;
                }
                if (e.KeyCode == Keys.D)
                {//Alt + d ：以選取文字進行[《看典古籍·古籍全文檢索》](https://kandianguji.com/search) (d=dian 典) 20241008
                    e.Handled = true;
                    if (!br.IsDriverInvalid())
                        br.LastValidWindow = br.driver.CurrentWindowHandle;
                    else
                        ResetLastValidWindow();

                    TopMost = false;
                    overtypeModeSelectedTextSetting(ref textBox1);
                    string str = textBox1.SelectedText;
                    if (str != string.Empty)
                    {
                        // 在主執行緒中設置剪貼簿 20241030 Copilot大菩薩
                        this.Invoke((MethodInvoker)delegate
                    {
                        Clipboard.SetText(str);
                    });
                    }
                    // 在新的執行緒中進行網頁檢索
                    Task.Run(() => { br.KanDianGuJiSearchAll(str); });
                    return;
                }

                if (e.KeyCode == Keys.E)
                {// Alt + e ：在完整編輯頁面中直接取代文字。請將被取代+取代成之二字前後並置，並將其選取後（或在被取代之文字前放置插入點）再按下此組合鍵以執行直接取代 20240718
                    e.Handled = true;
                    bool overTypeMode = !insertMode;
                    playSound(soundLike.press, true);
                    StringInfo character = new StringInfo(string.Empty); int s = textBox1.SelectionStart, l = 1;
                    if (textBox1.SelectionLength == 0)
                    {
                        while (character.LengthInTextElements < 2)
                        {
                            if (s + l <= textBox1.Text.Length && char.IsHighSurrogate(textBox1.Text.Substring(s + l - 1, 1).ToCharArray()[0]))
                                l++;
                            character = new StringInfo(textBox1.Text.Substring(s, l));
                            l++;
                        }
                        textBox1.Select(s, l - 1);

                    }
                    else
                        character = new StringInfo(textBox1.SelectedText);

                    if (!IsDriverInvalid())
                        //br.driver.SwitchTo().Window(br.driver.CurrentWindowHandle);                    
                        LastValidWindow = br.driver.CurrentWindowHandle;
                    else
                    {
                        ResetLastValidWindow();
                        LastValidWindow = br.driver.CurrentWindowHandle;
                    }
                    if (br.DirectlyReplacingCharacters(character))
                    {
                        if (IsValidUrl＿keyDownCtrlAdd(textBox3.Text))
                        {
                            if (driver.Url != textBox3.Text)
                            {
                                for (int i = driver.WindowHandles.Count - 1; i > -1; i--)
                                {
                                    driver.SwitchTo().Window(driver.WindowHandles[i]);
                                    string driverUrl = ReplaceUrl_Box2Editor(driver.Url);
                                    if (driverUrl.IndexOf(textBox3.Text) > -1 || textBox3.Text.Contains(driverUrl))
                                    {
                                        if (br.driver.Url.IndexOf("#box") > -1) br.driver.Url = textBox3.Text;//driverUrl;
                                        break;
                                    }
                                }

                            }
                        }
                        AvailableInUseBothKeysMouse();
                        //清除前一個要被取代的單字
                        undoRecord(); PauseEvents();
                        textBox1.SelectedText = character.SubstringByTextElements(0, 1);
                        textBox1.Select(textBox1.SelectionStart - character.SubstringByTextElements(0, 1).Length, character.SubstringByTextElements(0, 1).Length);
                        //textBox1.Text = textBox1.Text.Replace(character.SubstringByTextElements(0, 1), character.SubstringByTextElements(1, 1));
                        replaceWord(character.SubstringByTextElements(0, 1), character.SubstringByTextElements(1, 1));
                        ResumeEvents(); undoRecord();
                        textBox1.Select(textBox1.SelectionStart, 0);
                        if (overTypeMode) insertMode = !overTypeMode;
                    }
                    return;
                }

                if (e.KeyCode == Keys.G)
                {//Alt + g
                    e.Handled = true;
                    string x = overtypeModeSelectedTextSetting(ref textBox1);//CnText.ChangeSeltextWhenOvertypeMode(insertMode, textBox1);
                    GoogleSearch(x);
                    return;
                }

                if (e.KeyCode == Keys.H)
                {//Alt + h ：以選取文字檢索[《漢籍全文資料庫》](https://hanchi.ihp.sinica.edu.tw/) (h=han 漢) 20241008
                    e.Handled = true;
                    overtypeModeSelectedTextSetting(ref textBox1);
                    if (textBox1.SelectedText.IsNullOrEmpty()) return;
                    string x = textBox1.SelectedText;
                    if (!br.IsDriverInvalid())
                        br.LastValidWindow = br.driver.CurrentWindowHandle;
                    else
                        ResetLastValidWindow();
                    TopMost = false;
                    // 在主執行緒中設置剪貼簿 20241030 Copilot大菩薩
                    this.Invoke((MethodInvoker)delegate
                    {
                        Clipboard.SetText(x);
                    });

                    // 在新的執行緒中進行網頁檢索
                    Task.Run(() =>
                    {
                        br.HanchiSearch(x);
                    });
                    return;
                }

                if (e.KeyCode == Keys.J)
                {//Alt + j : 鍵入換行分段符號（newline）（同 Ctrl + j 的系統預設）
                    e.Handled = true;
                    //insertWords(Environment.NewLine, textBox1, textBox1.Text);
                    insertWords(Environment.NewLine, textBox1);
                    return;
                }
                if (e.KeyCode == Keys.K)
                {//Alt + k : 將選取的字詞句及其網址位址送到以下檔案的末後
                 // C:\Users\oscar\Dropbox\《古籍酷》AI%20OCR%20待改進者隨記%20感恩感恩 讚歎讚歎 南無阿彌陀佛.docx
                    e.Handled = true;
                    if (insertMode && textBox1.SelectionLength == 0)
                    {
                        SelectOneCharacter();
                    }
                    else
                        overtypeModeSelectedTextSetting(ref textBox1);
                    if (textBox1.SelectionLength > 0)
                    {
                        string txtbox1SelText = textBox1.SelectedText;
                        //if (Math.Abs(isChineseChar(txtbox1SelText, false)) != 1 && "■□◯".IndexOf(txtbox1SelText) == -1) return;
                        if (!IsChineseString(txtbox1SelText) && "■□◯".IndexOf(txtbox1SelText) == -1) return;
                        //playSound(soundLike.press, true);
                        Color clr = BackColor;
                        BackColor = Color.Aqua;
                        Refresh();//沒有這行還不行！20240709
                        Thread.Sleep(9);
                        BackColor = clr;
                        string url = textBox3.Text;
                        //Task.Run(() => { br.ImproveGJcoolOCRMemo(); });//因為即使開新執行緒，但仍是用同一個表單！                        
                        Task.Run(() =>
                        {
                            //if (tkImproveGJcoolKandiangujiOCRMemo != null) tkImproveGJcoolKandiangujiOCRMemo.Wait(160);                                
                            br.ImproveGJcoolKandiangujiOCRMemo(txtbox1SelText, url);
                            //tkImproveGJcoolKandiangujiOCRMemo = null;
                        });
                        //try
                        //{
                        //    Clipboard.SetText(textBox1.Text);//通常改正後是要再重標點，如書名等 20240306
                        //}
                        //catch (Exception)
                        //{
                        //    playSound(soundLike.error);
                        //}
                        AvailableInUseBothKeysMouse();
                    }
                    return;
                }
                if (e.KeyCode == Keys.M)
                {//Alt + m ： 以選取文字 search CTP的《史記三家注》 （m=ma(司馬遷的馬） 20241030
                    e.Handled = true;
                    overtypeModeSelectedTextSetting(ref textBox1);
                    if (textBox1.SelectedText.IsNullOrEmpty()) return;
                    string x = textBox1.SelectedText, url = "https://ctext.org/wiki.pl?if=gb&res=384378&searchu=" + System.Net.WebUtility.UrlEncode(x);
                research:
                    if (!IsDriverInvalid())
                    {
                        try
                        {
                            br.openNewTabWindow();
                            br.BringToFront("chrome");
                            driver.Navigate().GoToUrl(url);
                        }
                        catch (Exception)
                        {
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請重試！");
                        }
                    }
                    else
                    {
                        if (browsrOPMode == BrowserOPMode.appActivateByName)
                            Process.Start(url);
                        else
                        {
                            if (driver == null)
                                RestartChromedriver();
                            else
                            {
                                try
                                {
                                    driver.SwitchTo().Window(LastValidWindow);

                                }
                                catch (Exception)
                                {
                                    driver.SwitchTo().Window(driver.WindowHandles.Last());
                                    ResetLastValidWindow();
                                }
                            }

                            goto research;
                        }
                    }
                    return;
                }
                if (e.KeyCode == Keys.N)
                {//Alt + n : 將選取的字詞句及其網址位址送到以下檔案的末後
                 //> C:\Users\oscar\Dropbox\《看典古籍》OCR 待改進者隨記 感恩感恩 讚歎讚歎 南無阿彌陀佛                    
                    e.Handled = true;
                    if (insertMode && textBox1.SelectionLength == 0)
                    {
                        SelectOneCharacter();
                    }
                    else
                        overtypeModeSelectedTextSetting(ref textBox1);
                    if (textBox1.SelectionLength > 0)
                    {
                        string txtbox1SelText = textBox1.SelectedText;
                        //if (Math.Abs(isChineseChar(txtbox1SelText, false)) != 1 && "■□◯".IndexOf(txtbox1SelText) == -1) return;
                        if (!IsChineseString(txtbox1SelText) && "■□◯".IndexOf(txtbox1SelText) == -1) return;
                        //playSound(soundLike.press, true);
                        Color clr = BackColor;
                        BackColor = Color.Aqua;
                        Refresh();//沒有這行還不行！20240709
                        Thread.Sleep(9);
                        BackColor = clr;
                        string url = textBox3.Text;

                        //Task.Run(() => { br.ImproveGJcoolOCRMemo(); });//因為即使開新執行緒，但仍是用同一個表單！
                        Task.Run(() => { br.ImproveGJcoolKandiangujiOCRMemo(txtbox1SelText, url, "《看典古籍》"); });
                        //try
                        //{
                        //    Clipboard.SetText(textBox1.Text);//通常改正後是要再重標點，如書名等 20240306 20150119 現在標點機制調整，故可略去矣。感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主
                        //}
                        //catch (Exception)
                        //{
                        //    playSound(soundLike.error);
                        //}
                        AvailableInUseBothKeysMouse();
                    }
                    return;
                }
                /* Alt + l : 檢查/輸入抬頭平抬時的條件：執行topLineFactorIuput0condition()
                 *     > 目前只支援新增 condition=0 的情形，故名為 0condition，即當後綴是什麼時，此行文字雖短，不是分段，乃是平抬 
                 *     >> 0=後綴；1=前綴；2=前後之前；3前後之後；4是前+後之詞彙；5非前+後之詞彙；6非後綴之詞彙；7非前綴之詞彙*/
                if (e.KeyCode == Keys.L)
                {
                    e.Handled = true;
                    if (examSeledWord(out string termtochk))
                        Mdb.TopLineFactorIuput04condition(termtochk);
                    return;
                }

                if (e.KeyCode == Keys.P || e.KeyCode == Keys.Oem3)
                {//Alt + p 或 Alt + ` : 鍵入 "<p>" + newline（分行分段符號）
                    e.Handled = true;
                    if (e.KeyCode == Keys.P) { keysParagraphSymbol(false); return; }

                    if (textBox1.SelectedText != "" && textBox1.SelectedText.Replace("　", "") == "")
                    {//如果選取的是空格，則進行自動標記（mark）
                        if (wordsPerLinePara == -1)
                        {
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請先用Scroll Lock鍵讓程式取得每行/段的正常長度（字數）以供判斷。"); return;
                        }
                        autoMarkTitles(); return;
                        //string sps = textBox1.SelectedText;
                        //if (sps.Replace("　", string.Empty) != string.Empty) return;
                        //new Document(textBox1).AutoMarkTitles(sps.Length);
                    }

                    int s = textBox1.SelectionStart, l = textBox1.SelectionLength; string x = textBox1.Text;

                    //如果是沒有選取，且插入點在分行符號後或前，則加上段落標記<p>
                    if (x.Length == s ||
                        (s + 2 <= x.Length && (x.Substring(s, 2) == Environment.NewLine || x.Substring(s < 2 ? s : s - 2, 2)
                            == Environment.NewLine) && textBox1.SelectionLength == 0))//||
                                                                                      //(x.Substring(s < 2 ? s : s - 2, 2)== Environment.NewLine &&
                                                                                      //  x.Substring(s+1>x.Length?x.Length:s,1)!="　") // 有時標題是頂行的                       
                    {
                        keysParagraphSymbol();
                        undoRecord(); PauseEvents();
                        //將其後的空格也改成空白20241123
                        while (textBox1.SelectionStart + 1 < textBox1.TextLength && textBox1.Text.Substring(textBox1.SelectionStart, 1) == "　")
                        {

                            textBox1.Select(textBox1.SelectionStart, 1);
                            textBox1.SelectedText = "􏿽";
                            textBox1.Select(textBox1.SelectionStart + textBox1.SelectionLength, 0);
                        }
                        undoRecord(); ResumeEvents();
                        return;
                    }


                    //如果沒有選取，則自動選取插入點所在行/段
                    if (l == 0)
                    {
                        //如果插入點在0
                        if (s == 0)
                            l = textBox1.Text.IndexOf(Environment.NewLine);
                        //如果插入點前為分行符號
                        else if (s - 2 > 0 && x.Substring(s - 2, 2) == Environment.NewLine)
                        {
                            if (x.IndexOf(Environment.NewLine) > -1)
                                l = x.IndexOf(Environment.NewLine, s) - s;
                            else
                                l = x.Length - s;
                        }
                        else if (s - 1 > 0)// && s + 1 <= x.Length)
                        {
                            //如果插入點前不是全形空格
                            if (x.Substring(s - 1, 1) != "　")// && x.Substring(s + 1, 1) != "　")
                            {
                                if (x.LastIndexOf(Environment.NewLine, s) > -1)
                                {
                                    s = x.LastIndexOf(Environment.NewLine, s) + 2;
                                    while (s + 1 <= x.Length && x.Substring(s, 1) == "　")
                                        s++;//移動到前首空格之末
                                }
                                else
                                    s = 0;
                                if (x.IndexOf(Environment.NewLine, s) > -1)
                                    l = x.IndexOf(Environment.NewLine, s) - s;
                                else
                                    l = x.Length - s;
                            }
                            else//如果插入點前是全形空格
                            {
                                string xLine = GetLineText(x, s, out int lineS, out int lineL);
                                if (x.Substring(lineS, s - lineS).Replace("　", string.Empty) == string.Empty
                                    && x.Substring(s, 1) != "*")
                                {//如果插入點前全是全形空格
                                    while (s + 1 <= x.Length && x.Substring(s, 1) == "　")//移動到第一個非全形空格字的前面
                                        s++;
                                }
                                else//如果插入點前還有非全形空格
                                {
                                    s = lineS;
                                    while (s + 1 <= x.Length && x.Substring(s, 1) == "　")
                                        s++;
                                }

                                l = _document.GetCurrentParagraph().End - s;
                            }
                        }
                        else if (s - 1 == 0)
                            //if (x.Substring(s , 1) == "*")
                            //    s++;
                            //else
                            s--;

                        textBox1.Select(s, l);
                    }//如果沒有選取，則自動選取插件點所在行/段

                    BeginUpdate(false);

                    undoRecord(); stopUndoRec = true; PauseEvents();

                    titleMarkCode();

                    #region 標題標記完後若其行/段後面還有文字，則予以格式化（全形空格轉空白）
                    if (CountWordsinDomain("<p>", GetLineText(textBox1.Text, textBox1.SelectionStart)) > 1)
                    {
                        Paragraph p = new Range(_document, textBox1.SelectionStart, textBox1.SelectionStart, _document.Content).GetCurrentParagraph();
                        textBox1.Select(p.Start, p.End - p.Start);

                        textBox1.SelectedText = CnText.ReplaceFullWidthSpace_withBlank(textBox1.SelectedText);

                    }
                    #endregion 標題標記完後若其行/段後面還有文字，則予以格式化（全形空格轉空白）\
                    #region 訂正諸如作者中間全形空格誤轉空白的問題
                    x = textBox1.Text;
                    textBox1.Select(s, x.IndexOf("<p>", s) - s);
                    x = textBox1.SelectedText;
                    stopUndoRec = true;
                    replaceXdirectly(ref x, string.Empty, true);
                    stopUndoRec = false;
                    #endregion


                    undoRecord(); stopUndoRec = false; ResumeEvents();
                    EndUpdate();

                    #region 方便要按2次以上以降階標題                    
                    if (s + 1 <= textBox1.TextLength && textBox1.Text.Substring(s + 1) != "<"
                                    && Environment.NewLine.IndexOf(textBox1.Text.Substring(s + 1)) == -1)
                        if (s + 3 <= textBox1.TextLength && textBox1.Text.Substring(s + 1, 2) == Environment.NewLine)
                            s += 3;
                        else
                        {
                            while (s < textBox1.TextLength && textBox1.Text.Substring(s, 1) != "*")
                                ++s;
                            textBox1.SelectionStart = s;
                        }
                    #endregion

                    AvailableInUseBothKeysMouse();

                    return;
                }
                if (e.KeyCode == Keys.Q)
                {
                    e.Handled = true; splitLineByFristLen(); return;
                }
                if (e.KeyCode == Keys.S)
                {//Alt + s 小注文不換行
                    e.Handled = true; int s = textBox1.SelectionStart;
                    if (textBox1.SelectionLength == 0 && s < textBox1.TextLength)
                        textBox1.SelectionStart++;
                    notes_a_line();
                    textBox1.SelectionStart = s;
                    return;
                }

                if (e.KeyCode == Keys.T)
                {//Alt + t ：預測游標所在行是否為標題
                    e.Handled = true; detectTitleYetWithoutPreSpace(); return;
                }

                if (e.KeyCode == Keys.V)// Alt + v
                {
                    e.Handled = true;
                    if (examSeledWord(out string wordtoChk))
                        if (Mdb.VariantsExist(wordtoChk)) //SystemSounds.Hand.Play();
                                                          //如果已有資料對應，則閃示橘紅色（表單顏色）示警
                                                          //MessageBox.Show("existed!!");
                        {
                            Form1.playSound(soundLike.info);
                            this.BackColor = Color.Tomato;
                            this.Refresh();
                            Thread.Sleep(35);
                            this.BackColor = this.FormBackColorDefault;
                        }

                    return;
                }

                if (e.KeyCode == Keys.W)
                {//Alt + w ： 將插入點後的2行/段內容，改成夾注語法並接在插入點本行後（類似按下Ctrl + Shift + F1） 20250131大年初三
                    if (!e.Control & !e.Shift)
                    {
                        e.Handled = true;
                        Document document = new Document(ref textBox1);
                        undoRecord(); PauseEvents();
                        try
                        {
                            document.MergeParagraphsAtCaret();
                            //SendKeys.Send("{del}");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.HResult + ex.Message);
                            MessageBox.Show(ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        undoRecord(); ResumeEvents();
                        return;
                    }
                    else if (e.Shift || e.Control)
                    {//Alt + Shift + w 或 Ctrl + Alt + w ： 將執行 MergeParagraphsAtCaretWithShift 方法，插入點所在行將視為夾注的第1行，並將其後的1行/段合併上來，前後分別加上「{{」和「}}」，然後再將它們後面的那1行也合併上來。最後，插入點將停留在最後合併上來的那行/段文字的起始處。
                        //> 即插入點所在視為後面有正文的夾注第一行（則加按Shift鍵）
                        e.Handled = true;
                        Document document = new Document(ref textBox1);
                        undoRecord(); PauseEvents();
                        try
                        {
                            document.MergeParagraphsAtCaretWithShift();
                            //SendKeys.Send("{del}");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.HResult + ex.Message);
                            MessageBox.Show(ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        undoRecord(); ResumeEvents();
                        return;

                    }
                }

                if (e.KeyCode == Keys.X)
                {//Alt + x ：以所選之字（不能不等於1字）檢索《康熙字典網上版 》 https://www.kangxizidian.com/
                    e.Handled = true;
                    string x = SelectSingleCharacter();
                    if (browsrOPMode != BrowserOPMode.appActivateByName)
                    {
                        Task.Run(() =>
                        {
                            br.LookupKangxizidian(x);
                        });
                    }
                    return;
                }

                if (e.KeyCode == Keys.Z)
                {// Alt + z ：以所選之字（或插入點後之一字）檢索《字統網》等（或 執行【速檢網路字辭典.exe】）
                    e.Handled = true;
                    string x = SelectSingleCharacter();

                    if (browsrOPMode != BrowserOPMode.appActivateByName)
                    {
                        if (br.driver != null)
                        {
                            TopMost = false;
                            if (new StringInfo(x).LengthInTextElements > 1)
                                Clipboard.SetText(x);
                            //Task.Run(() =>
                            //{
                            if (!LookupZitools(x))
                                MessageBoxShowOKExclamationDefaultDesktopOnly("查找《字統網》發生錯誤，請重來一遍。感恩感恩　南無阿彌陀佛");
                            //});
                        }
                        else
                        {
                            if (MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否要執行【速檢網路字辭典】") == DialogResult.OK)
                                Process.Start(dropBoxPathIncldBackSlash + @"VS\VB\速檢網路字辭典\速檢網路字辭典\bin\Debug\速檢網路字辭典.exe");
                        }

                    }
                    else
                        Process.Start(dropBoxPathIncldBackSlash + @"VS\VB\速檢網路字辭典\速檢網路字辭典\bin\Debug\速檢網路字辭典.exe");

                    return;
                }

                if (e.KeyCode == Keys.F7)
                {// Alt + F7 : 每行縮排一格後將其末誤標之<p>
                    e.Handled = true; indent_ClearEnd＿P_Mark(); return;
                }

                if (e.KeyCode == Keys.Add || e.KeyCode == Keys.Oemplus)//|| e.KeyCode == Keys.Subtract || e.KeyCode == Keys.NumPad5)
                {// Alt + +
                    if (e.KeyCode == Keys.Oemplus && autoPaste2QuickEdit) return;//防止在連續輸入時誤按
                    e.Handled = true;
                    //還原放大的書圖
                    RestoreImageSize();
                    //if (keyDownCtrlAdd(false)) if (textBox1.Text != "") { pauseEvents(); textBox1.Text = ""; resumeEvents(); }
                    keyDownCtrlAdd(false);
                    return;
                }

                if (e.KeyCode == Keys.OemMinus || e.KeyCode == Keys.Subtract)
                {// Alt + -（字母區與數字鍵盤的減號） : 如果被選取的是「􏿽」則與下一個「{{」對調；若是「}}」則與「􏿽」對調。。（針對《國學大師》《四庫全書》文本小注文誤標而開發）
                    e.Handled = true;
                    undoRecord(); stopUndoRec = true;
                    using (GXDS gxds = new GXDS(this)) { gxds.correctBlankAndUppercurlybrackets(ref textBox1); }
                    stopUndoRec = false;
                    return;
                }

                if (e.KeyCode == Keys.OemBackslash || e.KeyCode == Keys.Oem5)
                {// Alt + \ 
                    e.Handled = true; clearNewLinesAfterCaret(); return;
                }
                if (e.KeyCode == Keys.Down)
                //Ctrl + ↓ 或 Alr + ↓：從插入點開始向後移至這一段末（無分段則不移動）
                {
                    e.Handled = true; keyDownCtrlAltUpDown(e); return;
                }

                if (e.KeyCode == Keys.Delete)
                {//Alt + Del : 刪除插入點後第一個分行分段
                    e.Handled = true;
                    string x = textBox1.Text;
                    int s = textBox1.SelectionStart, p = x.IndexOf(Environment.NewLine, s), l = textBox1.SelectionLength;
                    if (p == -1) return;
                    x = x.Substring(0, p) + x.Substring(p + Environment.NewLine.Length);
                    textBox1.Text = x;
                    restoreCaretPosition(textBox1, s, l);
                    return;
                }

                if (e.KeyCode == Keys.Insert)
                {//Alt + Insert ：將剪貼簿的文字內容讀入textBox1中;若在手動鍵入輸入模式下則自動加上書名號篇名號
                    e.Handled = true;
                    int s = textBox1.SelectionStart, l = textBox1.SelectionLength; string x = textBox1.Text;
                    caretPositionRecord();
                    //string clpTxt = Clipboard.GetText();
                    //if (clpTxt.StartsWith("http"))
                    //{
                    //    playSound(soundLike.warn);
                    //    clpTxt = textBox1.Text;
                    //}
                    //if (keyinTextMode && clpTxt != ClpTxtBefore)// &&clpTxt.IndexOf("《") == -1 && clpTxt.IndexOf("〈") == -1 && clpTxt.IndexOf("·") == -1)//之前是沒有優化 booksPunctuation 才需要避免已經標點過的又標，現在有正則表達式把關，就沒有這問題了。感恩感恩　讚歎讚歎　chatGPT大菩薩+Bing大菩薩 南無阿彌陀佛
                    //{
                    //bool gjcoolocrResultManual = clpTxt.IndexOf(Environment.NewLine + Environment.NewLine) > -1;
                    //if (gjcoolocrResultManual) clpTxt = clpTxt.Replace(Environment.NewLine + Environment.NewLine, Environment.NewLine);
                    ////if (!ocrTextMode)
                    ////{
                    //textBox1.Text = clpTxt;
                    ////clearBracketsInsidePairsBrackets();
                    //clpTxt = textBox1.Text;
                    //textBox1.Text = CnText.BooksPunctuation(ref clpTxt, true);
                    textBox1.Text = CnText.BooksPunctuation(ref x, true);
                    //}
                    //if (gjcoolocrResultManual)
                    //    {
                    //        if (br.driver != null)
                    //        {
                    //            try
                    //            {
                    //                br.driver.SwitchTo().Window(br.driver.CurrentWindowHandle);
                    //                SendKeys.Send("%r");
                    //                Thread.Sleep(550);
                    //                //br.driver.SwitchTo().Alert().SendKeys(OpenQA.Selenium.Keys.Space);
                    //                br.driver.SwitchTo().Alert().Accept();
                    //                //SendKeys.Send(" ");

                    //                playSound(soundLike.exam);
                    //                //Activate();
                    //                bringBackMousePosFrmCenter();
                    //            }
                    //            catch (Exception)
                    //            {
                    //                //throw;
                    //            }
                    //        }
                    //    }
                    //}
                    //else textBox1.Text = clpTxt;
                    dragDrop = false;
                    AvailableInUseBothKeysMouse();
                    //if (s > 0) restoreCaretPosition(textBox1, s, 0);
                    if (s > 0) restoreCaretPosition(textBox1, s, l);//20250122
                    if (textBox1.SelectionStart == 0 && s > 0)
                    {
                        textBox1.SelectionStart = s;
                        textBox1.SelectionLength = l;
                    }

                    //caretPositionRecord();
                    return;
                }

                // Alt + Pause
                if (e.KeyCode == Keys.Pause)//20231005
                {//Shift + F8 或 Alt + Shift + Pause ： 加上篇名格式代碼並前置N個全形空格.N，預設為2.且可在執行此項時，選取空格數以重設篇名前要空的格數
                    e.Handled = true;
                    undoRecord(); stopUndoRec = true; PauseEvents();
                    if (e.Shift)
                        keysTitleCodeAndPreWideSpace();
                    else//alt + pause 自動判斷標題行，加上篇名格式代碼並前置N個全形空格.N，預設為2.且可在執行此項時，選取空格數以重設篇名前要空的格數
                        autoKeysTitleCodeAndPreWideSpace();
                    ResumeEvents(); stopUndoRec = false;
                    Clipboard.SetText(textBox1.Text);//通常標識後是要再重標點，如書名等 20240306
                    return;
                }

                /*20230723Bing大菩薩：
                                 您可以在 KeyDown 事件中檢查是否按下了「Alt + [」組合鍵。在 KeyEventArgs 中，您可以使用 e.Modifiers 屬性來檢查是否按下了修飾鍵（例如 Alt），並使用 e.KeyCode 屬性來檢查是否按下了其他鍵（例如 [）。以下是一個示例代碼：
                        請注意，您需要將 Form 的 KeyPreview 屬性設置為 true，才能使此代碼正常工作。¹

                        來源: 與 Bing 的交談， 2023/7/23
                        (1) 檢查按下哪一個修飾詞按鍵 - Windows Forms .NET | Microsoft Learn. https://learn.microsoft.com/zh-tw/dotnet/desktop/winforms/input-keyboard/how-to-check-modifier-key?view=netdesktop-7.0.
                        (2) 鍵盤輸入概觀 - Windows Forms .NET | Microsoft Learn. https://learn.microsoft.com/zh-tw/dotnet/desktop/winforms/input-keyboard/overview?view=netdesktop-7.0.
                        (3) 作法：判斷按下的輔助按鍵 - Windows Forms .NET Framework. https://learn.microsoft.com/zh-tw/dotnet/desktop/winforms/how-to-determine-which-modifier-key-was-pressed?view=netframeworkdesktop-4.8.
                 */
                if (e.Modifiers == Keys.Alt && e.KeyCode == Keys.Oem4)
                {
                    e.Handled = true;
                    preceded_followed_specify_symbols("〖〗");
                    return;
                }
                /*
                 `Keys.OemOpenBrackets` 是 `System.Windows.Forms.Keys` 枚舉中的一個值，表示 OEM 左方括號鍵。OEM 鍵是指隨地區鍵盤而變化的鍵。例如，美國鍵盤上有方括號和大括號，而德國鍵盤上則有變音符號。它們被稱為「OEM」，因為鍵盤的原始設備製造商負責定義它們的功能¹。
                    來源: 與 Bing 的交談， 2023/7/23
                    (1) .net - What are the "OEM" keys in the System.Windows.Forms.Keys enumeration? - Stack Overflow. https://stackoverflow.com/questions/582403/what-are-the-oem-keys-in-the-system-windows-forms-keys-enumeration.
                    (2) Key Enum (System.Windows.Input) | Microsoft Learn. https://learn.microsoft.com/en-us/dotnet/api/system.windows.input.key?view=windowsdesktop-7.0.
                    (3) Keys Enumeration. http://docs.go-mono.com/monodoc.ashx?link=T%3ASystem.Windows.Forms.Keys.
                 */

            }//以上 Alt
            #endregion

            #region 按下單一鍵            
            if (ModifierKeys == Keys.None)
            {//按下單一鍵
                if (e.KeyCode == Keys.Scroll)
                {//按下 Scroll Lock 將字數較少的行/段落尾末標上「<p>」符號
                    e.Handled = true; paragraphMarkingAccordingFirstOneLineLength();
                    //if (!textBox1.Text.IsNullOrEmpty())
                    //    Clipboard.SetText(textBox1.Text);
                    return;
                }

                if (e.KeyCode == Keys.Insert)
                {
                    e.Handled = true;
                    InsertModeSwitcher();
                    return;
                }
                if (e.KeyCode == Keys.F1 || e.KeyCode == Keys.Pause)//暫時取消，釋放 F1、 Pause 鍵給 Alt + Shift + 2 、Alt + F7用
                {//- 按下 F1 鍵：以找到的字串位置**前**分行分段
                 // -按下 Pause Break 鍵：以找到的字串位置** 後**分行分段
                    e.Handled = true;
                    if (e.KeyCode == Keys.Pause)
                    {
                        indent_ClearEnd＿P_Mark(); return;
                    }
                    if (textBox1.SelectionLength == 0)
                        Clipboard.SetText(textBox1.Text);
                    else
                        poetryFormat();
                    //else
                    //splitLineParabySeltext(e.KeyCode);
                    return;
                }


                if (e.KeyCode == Keys.F2)
                {//按下 F2 鍵,並複製textBox1的內容到剪貼簿
                    keyDownF2(textBox1);
                    try
                    {
                        Clipboard.SetText(textBox1.Text);
                    }
                    catch (Exception)
                    {
                    }
                    return;
                }

                if (e.KeyCode == Keys.F3)
                {//按下 F3 鍵
                    e.Handled = true;
                    int foundwhere;
                    if (textBox1.SelectionLength == 0) overtypeModeSelectedTextSetting(ref textBox1);
                    string findword = textBox1.SelectionLength == 0 ? lastFindStr : textBox1.SelectedText;
                    if (findword == "") findword = textBox2.Text;
                    if (findword != "")
                    {
                        int start = textBox1.SelectionStart + 1; string x = textBox1.Text;
                        if (start >= textBox1.Text.Length) return;
                        foundwhere = x.IndexOf(findword, start, StringComparison.Ordinal);
                        if (foundwhere == -1)
                        {
                            MessageBox.Show("not found next!"); return;
                        }
                        textBox1.SelectionStart = foundwhere;
                        //if ()//標題搜尋時不選取，以利keysTitleCode()執行
                        //{

                        //}
                        textBox1.SelectionLength = findword.Length;
                        textBox1.ScrollToCaret();
                    }
                    if (findword != "") lastFindStr = findword;
                    return;
                }

                if (e.KeyCode == Keys.F4)
                {//按下 F4 鍵： 重複做最後一次的輸入1次
                    int c = lastKeyPress.Count - 1;
                    if (c < 0) return;
                    string lk = lastKeyPress[c];
                    if (lk != "")
                    {
                        //undoRecord();
                        //stopUndoRec = true;
                        //textBox1.SelectedText = lk;
                        if ("{}".IndexOf(lk) > -1) lk = "{" + lk + "}";
                        SendKeys.Send(lk);
                        //stopUndoRec = false;
                    }
                    return;
                }
                if (e.KeyCode == Keys.F6)
                {//F6 : 標題降階（增加標題前之星號）
                    e.Handled = true;
                    keysAsteriskPreTitle();
                    return;
                }
                if (e.KeyCode == Keys.F7)
                {//F7 ： 每行/段前空一格
                    e.Handled = true;
                    int st = textBox1.SelectionStart, ed = textBox1.SelectionStart + textBox1.SelectionLength;
                    expandSelectedTextRangeToWholeLinePara(st, ed - st, textBox1.Text);
                    //string x = textBox1.Text;
                    st = textBox1.SelectionStart; ed = textBox1.SelectionStart + textBox1.SelectionLength;
                    //Range range = new Range(new Document(ref x), st, ed);
                    string x = textBox1.SelectedText;
                    Document document = new Document(x);
                    Range range = new Range(document, 0, x.Length, document.Content);
                    undoRecord(); stopUndoRec = true; PauseEvents();
                    indent(ref range);
                    textBox1.SelectedText = range.Text;
                    undoRecord(); stopUndoRec = false; ResumeEvents();
                    expandSelectedTextRangeToWholeLinePara(st, ed - st, textBox1.Text);
                    if (!Active)
                        bringBackMousePosFrmCenter();
                    //textBox1.Select(st > 0 ? st - 1 : st, textBox1.Text.IndexOf(Environment.NewLine, ed) == -1 ? textBox1.TextLength : textBox1.Text.IndexOf(Environment.NewLine, ed));
                    return;
                }
                if (e.KeyCode == Keys.F11)
                {
                    //F11 : run replaceXdirrectly() 維基文庫等欲直接抽換之字
                    e.Handled = true;
                    string x = textBox1.Text;
                    undoRecord();
                    replaceXdirectly(ref x, string.Empty, false);
                    undoRecord();
                    return;
                }
                if (e.KeyCode == Keys.Add)
                {//在非自動且手動輸入模式下單獨按下數字鍵盤的「+」("+") →方便檢索到這塊程式碼
                 //整頁貼上Quick edit [簡單修改模式]  並將下一頁直接送交《古籍酷》OCR// 原為加上篇名格式代碼

                    //還原放大的書圖
                    RestoreImageSize();

                    #region 全自動貼上模式（自動連續輸入模式）不適用，因為要輸入「+」成「+<p>」以取消「<p>」的檢查判斷
                    if (autoPaste2QuickEdit)
                    {
                        return;
                    }
                    #endregion

                    //防止誤按
                    if (br.Quickedit_data_textboxTxt == "+" &&
                        textBox1.Text.Replace("+", string.Empty) == string.Empty)//textBox1.Text == string.Empty 已包含
                        return;
                    if (keyinTextMode && OcrTextMode)
                    {
                        e.Handled = true;
                        //if (textBox1.Text != string.Empty)
                        //{ undoRecord(); pauseEvents(); textBox1.Text = string.Empty; resumeEvents(); }
                        //OpenQA.Selenium.IWebElement iw = br.waitFindWebElementBySelector_ToBeClickable("#canvas > svg");
                        TopMost = false;
                        OpenQA.Selenium.IWebElement iw = br.WaitFindWebElementBySelector_ToBeClickable("#content");
                        if (iw != null) // clickCopybutton_GjcoolFastExperience(iw.Location); 
                            Cursor.Position = (Point)iw.Location;

                        rep://OCR連續輸入
                        if (pagePaste2GjcoolOCR() && PasteOcrResultFisrtMode && ModifierKeys != Keys.Control && !br.confirm_that_you_are_human)
                            goto rep;

                        //if (!pagePaste2GjcoolOCR())//因為失敗的結果非唯一，故改寫在方法之中
                        //{ //數字鍵盤的「+」
                        //    PauseEvents();//為了等待時可以切到別的視窗看看，在執行成功且完成後，再把這2個關鍵的視窗置前
                        //    try
                        //    {
                        //        br.driver.SwitchTo().Window(br.LastValidWindow);
                        //        Thread.Sleep(300);
                        //    }
                        //    catch (Exception)
                        //    {
                        //        br.LastValidWindow = null;
                        //    }
                        //    ResumeEvents();
                        //    //SendKeys.Send("^z");//因為會輸入「+」取代選取區文字//這應該是在下一個程序textBox1_KeyPress才會輸入，而不是在這時
                        //}
                        //if (!Visible) Visible = true;
                        //bringBackMousePosFrmCenter();
                        if (PagePaste2GjcoolOCR_ing && (BatchProcessingGJcoolOCR || PasteOcrResultFisrtMode)) PagePaste2GjcoolOCR_ing = false;
                        return;
                    }
                    else
                    {
                        e.Handled = true;
                        if (keyinTextMode)
                            SelectAll2Quickedit();
                        else// if(!autoPaste2QuickEdit)//前已有 return了
                        {
                            { playSound(soundLike.press, true); altA_predictEndofPageRange(); }
                            keyDownCtrlAdd(false);
                            RestoreImageSize();
                            if (!autoPaste2QuickEdit && !isSKQSFrontPage(textBox1.Text)) AvailableInUseBothKeysMouse();
                            EndUpdate();
                            PagePaste2GjcoolOCR_ing = true;//借用這個變數，以免誤輸入「+」
                        }
                        return;
                    }

                }
                if (e.KeyCode == Keys.Subtract)
                {//"-"： 在非自動且手動輸入模式下單獨按下數字鍵盤的「-」，執行與按下 Scroll Lock 一樣的功能
                 //將字數較少的行 / 段落尾末標上分行 / 段符號（「\< p\>」或「\。< p\>」
                    if (keyinTextMode && !autoPaste2QuickEdit)
                    {
                        e.Handled = true;
                        paragraphMarkingAccordingFirstOneLineLength();
                        if (textBox1.Text.IsNullOrEmpty()) return;
                        if (isClipBoardAvailable_Text())
                            try
                            {
                                Clipboard.SetText(textBox1.Text);
                            }
                            catch (Exception)
                            {
                                //playSound(soundLike.error, true);
                            }
                        return;
                    }
                }
                if (e.KeyCode == Keys.F10)
                {//F10 同上
                    e.Handled = true;
                    paragraphMarkingAccordingFirstOneLineLength();
                    if (textBox1.Text != string.Empty)
                    {
                        try
                        {
                            Clipboard.SetText(textBox1.Text);
                        }
                        catch (Exception)
                        {
                        }
                    }
                    return;

                }

                //以上按下單一鍵
                #endregion
            }
        }

        /// <summary>
        /// 將textBox1的內容送去簡單修改模式的方塊中
        /// </summary>
        internal void SelectAll2Quickedit()
        {
            PagePaste2GjcoolOCR_ing = true;
            PressAddKeyMethodPaste2QuickEditBox(); WindowsScrolltoTop();
            bringBackMousePosFrmCenter();
            if (ModifierKeys == Keys.Shift)
            {
                playSound(soundLike.press, true);
                //toOCR(br.OCRSiteTitle.GJcool);
                toOCR(PagePast2OCRsite);
            }

            //避免事件被終止
            if (!_eventsEnabled) _eventsEnabled = true;

            //TopMost = true;
        }

        /// <summary>
        /// 執行注文空白錯亂文本的校正 20250225
        /// </summary>
        private void runCorrectNoteBlankContent()
        {
            string x, openMark = string.Empty; ;
            if (textBox1.SelectedText == " ") textBox1.SelectionLength = 0;
            if (textBox1.SelectionLength == 0)
            {
                #region 前端操作用則用   _document
                //Paragraph p = _document.Range(textBox1.SelectionStart, textBox1.SelectionStart).GetCurrentParagraph();
                #endregion
                #region 後端計算則用 new Document(ref x) 效能差太多！！
                x = textBox1.Text;
                Paragraph p = new Document(x).Range(textBox1.SelectionStart, textBox1.SelectionStart).GetCurrentParagraph();
                #endregion
                textBox1.Select(p.Start, p.End - p.Start);
                if (textBox1.SelectedText.IndexOf("{{") == -1)
                    textBox1.Select(p.Range.GetPreviousParagraph().Start, p.End - p.Range.GetPreviousParagraph().Start);
                if (textBox1.SelectedText.IndexOf("}}") == -1)
                    textBox1.Select(p.Start, p.Range.GetNextParagraph().End - p.Start);
            }
            x = textBox1.SelectedText;//以一行/段為處理的單位
            if (x.IndexOf("{{") == -1 && textBox1.SelectionStart == 0)
            {
                openMark = "{{"; int end = textBox1.SelectionLength;
                x = openMark + x;
                textBox1.SelectedText = x;
                textBox1.Select(0, end + 2);
            }
            if (x.IndexOf("{{") == -1 || x.IndexOf("}}") == -1
                || x.IndexOf(" ") == -1)
                return;

            //以上防呆

            //x = textBox1.SelectedText;//以一行/段為處理的單位
            x = CnText.CorrectNoteBlankContent(x, out int spaceIndex);
            if (!x.IsNullOrEmpty() && textBox1.SelectedText != x)
            {
                undoRecord();
                PauseEvents();
                if (openMark != string.Empty)
                    x = x.Substring(2);
                textBox1.SelectedText = x;
                ResumeEvents();
                undoRecord();
                //標識被置換位置後的「􏿽」以便檢閱 20250618 感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主
                textBox1.Select(textBox1.Text.LastIndexOf("􏿽", textBox1.SelectionStart), insertMode ? "􏿽".Length : 0);
            }
            else if (spaceIndex > -1)
            {
                textBox1.Select(spaceIndex + textBox1.SelectionStart, 1);
                Refresh();
                textBox1.ScrollToCaret();
                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("1次僅處理1個半形空格！"
                    + Environment.NewLine + Environment.NewLine + "請先整理過文本再執行校正！！");
            }
            return;
        }

        /// <summary>
        /// 如果插入點在「+」、「<p>」之間則移至「<p>」之後 20250215
        /// </summary>
        private void betweenAddParaMarkMoveEnd()
        {
            int s = textBox1.SelectionStart; string x = textBox1.Text;
            if (textBox1.TextLength >= s + 3 && s - 1 >= 0
                && x.Substring(s, 3) == "<p>"
                && x.Substring(s - 1, 1) == "+")
                textBox1.SelectionStart += 3;
        }
        /// <summary>
        /// 在自動連續輸入時自動預測送出的文本範圍
        /// 供 Alt + a 供及自動連續輸入時數字鍵盤「Ctrl + 數字鍵盤 +」使用
        /// </summary>
        /// <returns>失敗則傳回false</returns>
        private bool altA_predictEndofPageRange()
        {
            if (textBox1.TextLength == 0) return false;

            undoRecord();
            //if (keyDownCtrlAdd(false)) // if (textBox1.Text != "") { pauseEvents(); textBox1.Text = ""; resumeEvents(); }
            //autoPastetoCtextQuitEditTextboxCancel = false;
            betweenAddParaMarkMoveEnd();

            if (autoPaste2QuickEdit && textBox1.SelectionLength > 0)
            {
                textBox1.SelectionStart += textBox1.SelectionLength;
                textBox1.DeselectAll();//若有選取，會影響自動判別各頁尾端，但這只是將 SelectionLength 屬性歸零而已
            }
            int end;
            //20250213元宵後一日
            if (pageEndText10 != string.Empty)
            {
                string pageEndText10_original = pageEndText10;
                pageEndText10 = char.IsLowSurrogate(pageEndText10.Substring(0, 1).ToCharArray()[0]) ? pageEndText10.Substring(1) : pageEndText10;
                int discrepancy = pageEndText10.Length - pageEndText10_original.Length;
                //end = textBox1.Text.IndexOf(pageEndText10);//當surrogate被截斷時，用此無效 20250216
                end = textBox1.Text.IndexOf(pageEndText10) + discrepancy;

                // 使用自定義函數查找部分匹配的字符串位置                        
                //bool found = CnText.PartialMatch(textBox1.Text, pageEndText10, out end);

                //if (found)
                if (end > -1)
                {
                    //end += (pageEndText10.Length - discrepancy);
                    end += 10;
                    if (end > pageTextEndPosition)
                        pageTextEndPosition = end;
                    //else if (textBox1.Text.IndexOf(pageEndText10, end + 1) == -1 && pageTextEndPosition != end)
                    else if (CnText.IsPageEndTextUnique(pageEndText10, textBox1) && pageTextEndPosition != end)
                        pageTextEndPosition = end;

                    textBox1.Select(end, 0);

                }
                //if (end > pageTextEndPosition) pageTextEndPosition = end;

            }

            //lines_perPage = 16;//●●●●●●●●●●●●●●●●●●
            if (lines_perPage == 0)//&& autoPaste2QuickEdit)
            {
                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請先設定每頁正常的行/段數！");
                return false;
            }

            string x = textBox1.Text;
            int st = (pageTextEndPosition == 0 || pageTextEndPosition < textBox1.SelectionStart) ? textBox1.SelectionStart : pageTextEndPosition;
            Paragraph p;
            if (st > _document.Text.Length) st = _document.Text.Length;
            //if (_document.Range(0, st).Paragraphs.Count > lines_perPage / 2)
            //string rngTxt = _document.Text.Substring(0, _document.Range(st, st).GetCurrentParagraph().End);
            string rngTxt = textBox1.Text.Substring(0, textBox1.Text.IndexOf(Environment.NewLine, st) > -1 ?
                                                                textBox1.Text.IndexOf(Environment.NewLine, st) :
                                                                textBox1.TextLength);


            int paraCount = countLinesPerPage(rngTxt);


            if (paraCount > lines_perPage)
            {
                //List<Paragraph> paragraphs = _document.GetParagraphs();
                List<Paragraph> paragraphs = new Document(x).GetParagraphs();
                if (paragraphs.Count > 0 && paragraphs.Count >= paraCount)
                {
                    if (paragraphs[paraCount - 1].Text.StartsWith("{{"))
                        p = paragraphs[paraCount - 2 < 0 ? 0 : paraCount - 2];//小注文則提前2行/段                    
                    else
                        p = paragraphs[paraCount - 1 < 0 ? 0 : paraCount - 1];
                }
                else
                    p = paragraphs[paragraphs.Count - 1];
                end = p.End;
            }
            else
            {
                //p = _document.Range(st, st).GetCurrentParagraph();
                p = new Document(x).Range(st, st).GetCurrentParagraph();
                end = p.End;
            }
            //如果所抓到或預測的送出範圍少於正常每頁的行/段數的話： 20250216
            if (paraCount < lines_perPage)
            {
                if (!fastMode && textBox1.Text.Contains(pageEndText10) && pageTextEndPosition > 0) return true;//按下 Ctrl + - 指定了送出位置，則不再以每頁行數來交由程式自動判斷了
                predictRange(ref end);
            }


            if (end != textBox1.SelectionStart && textBox1.SelectionLength != 0)
            {
                if (DialogResult.Cancel == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly
                    ("確定是這個位置之前的內容要送出嗎？" +
                    Environment.NewLine + Environment.NewLine +
                    p.Text))
                    return false;
                else//確定是這個位置之前的內容要送出,則須重設
                {

                    //pageTextEndPosition = 0; pageEndText10 = string.Empty;
                    //pageTextEndPosition = end;
                    //if (pageTextEndPosition - 10 > -1)
                    //    pageEndText10 = textBox1.Text.Substring(pageTextEndPosition - 10, 10);
                    ////textBox1.SelectionStart = end;textBox1.SelectionLength = 0;//參考 8604行
                    textBox1.Select(end, 0);
                    pageTextEndPosition = end;
                    pageEndText10 = textBox1.Text.Substring(end - 10, 10);
                }
            }
            else
            {
                textBox1.Select(end, 0);
                pageTextEndPosition = end;
                if (textBox1.TextLength > 10)
                    pageEndText10 = textBox1.Text.Substring(end - 10, 10);
                else
                    pageEndText10 = string.Empty;
            }
            //}

            //防呆 20250217
            //int paraCount = countLinesPerPage(textBox1.Text.Substring(0, textBox1.SelectionStart));
            paraCount = countLinesPerPage(textBox1.Text.Substring(0, textBox1.SelectionStart));


            if (paraCount != lines_perPage)
            {
                if (!predictRange(ref end)) return false;
            }

            ////送去這個方法以送出至CTP前要先把插入點放在適當的位置
            //if (!KeyboardInfo.getKeyStateDown(System.Windows.Input.Key.LeftShift))
            //keyDownCtrlAdd(false);
            return true;
        }
        /// <summary>
        /// 預測、評估要送出去的頁面範圍
        /// </summary>
        /// <returns>失敗時傳回false</returns>
        bool predictRange(ref int end)
        {
            if (textBox1.TextLength == 0) return false;
            if (lines_perPage == 0)
            {
                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請先設定每頁正常的行/段數！");
                return false;
            }
            //if (x != textBox1.Text) x = textBox1.Text;
            string x = textBox1.Text;
            end = 0;
            int paraCount = countLinesPerPage(x.Substring(0, end));
            while (paraCount < lines_perPage)
            {
                end = x.IndexOf(Environment.NewLine, end + 1);// + 2;
                if (end > -1)
                {
                    if (end > x.Length)
                    {
                        end = x.Length;
                        break;
                    }
                }
                else
                {
                    end = x.Length;
                    break;
                }
                string countRange = x.Substring(0, end).Replace("+<p>", "<p>");
                paraCount = countLinesPerPage(countRange);
            }
            //p = _document.Range(0, end).Paragraphs[_document.Range(0, end).Paragraphs.Count - 1];
            //p = new Document(ref x).Range(0, end).Paragraphs[_document.Range(0, end).Paragraphs.Count - 1];
            //end = p.End;
            textBox1.Select(end, 0);
            pageTextEndPosition = end;
            if (textBox1.TextLength > 10)
                pageEndText10 = textBox1.Text.Substring(end - 10, 10);
            else
                pageEndText10 = string.Empty;
            return true;
        }
        /// <summary>
        /// 檢查原文是否已遭篡改
        /// </summary>
        /// <param name="x">要檢查的文本</param>
        /// <param name="original">原文</param>
        /// <returns>遭篡改則傳回true</returns>
        internal static bool IsTextModified(string x, string original)
        {
            int punctsCntr = 0;
            string originalPure = original;
            foreach (var item in PunctuationsNum + "{}　")//先略過注文標記及空格
            {
                //if (item.ToString() == "：") Debugger.Break();
                originalPure = originalPure.Replace(item.ToString(), string.Empty);
            }
            originalPure = originalPure.Replace(Environment.NewLine, string.Empty);
            bool result = false;
            StringInfo xSI = new StringInfo(x);
            StringInfo originalPureSI = new StringInfo(originalPure);

            for (int i = 0; i < xSI.LengthInTextElements; i++)
            {
                if (xSI.SubstringByTextElements(i) == "□" ||
                    xSI.SubstringByTextElements(i) == "􏿽") Debugger.Break();
                if (Form1.PunctuationsNum.IndexOf(xSI.SubstringByTextElements(i, 1)) == -1)
                {
                    string originalTxt = originalPureSI.SubstringByTextElements(i - punctsCntr, 1),
                            modifiedTxt = xSI.SubstringByTextElements(i, 1);
                    if (modifiedTxt != originalTxt)
                    {
                        Clipboard.SetText("「" + originalTxt + "」被篡改成「" + modifiedTxt.Replace("􏿽", "□") + "」！！！阿彌陀佛");
                        result = true;
                        if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly(

                            "原文已遭篡改，請檢查！" + Environment.NewLine + Environment.NewLine
                                + "被改成：  " + modifiedTxt
                                + Environment.NewLine
                                + "原來是：  " +
                                originalTxt
                                ) == DialogResult.Cancel)
                            break;
                    }
                }
                else
                    punctsCntr++;
            }

            //for (int i = 0; i < x.Length; i++)
            //{
            //    if (Form1.PunctuationsNum.IndexOf(x.Substring(i, 1).ToString()) == -1)
            //    {
            //        if (x.Substring(i, 1) != originalPure.Substring(i - punctsCntr, 1))
            //        {
            //            //Clipboard.SetText((char.IsHighSurrogate(x.Substring(i, 1).ToCharArray()[0]) ? x.Substring(i, 2) : x.Substring(i, 1)).ToString());
            //            string originalTxt = char.IsHighSurrogate(originalPure.Substring(i - punctsCntr, 1).ToCharArray()[0]) ? originalPure.Substring(i - punctsCntr, 2) : originalPure.Substring(i - punctsCntr, 1),
            //                modifiedTxt = x.Substring(i, 1);
            //            //(char.IsHighSurrogate(x.Substring(i, 1).ToCharArray()[0]) ? x.Substring(i, 2) : x.Substring(i, 1)).ToString();
            //            //char.IsHighSurrogate(x.Substring(i, 1).ToString().ToCharArray()[0]) ? x.Substring(i, 2) : x.Substring(i, 1);

            //            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(

            //                "原文已遭篡改，請檢查！" + Environment.NewLine + Environment.NewLine
            //                    + "被改成：  " + modifiedTxt
            //                    + Environment.NewLine
            //                    + "原來是：  " +
            //                    originalTxt
            //                    );
            //            Clipboard.SetText("「" + originalTxt + "」被篡改成「" + modifiedTxt + "」！！！阿彌陀佛");
            //            result= true;
            //        }
            //    }
            //    else
            //        punctsCntr++;
            //}


            return result;
        }

        /// <summary>
        /// 在textBox1選取1個字
        /// </summary>
        internal void SelectOneCharacter()
        {
            if (char.IsHighSurrogate(textBox1.Text.Substring(textBox1.SelectionStart, 1).ToCharArray()[0]))
                textBox1.SelectionLength = 2;
            else
                textBox1.SelectionLength = 1;
        }

        /// <summary>
        /// 在textBox1中選取單字
        /// </summary>
        /// <returns>傳回被選取的單字</returns>
        internal string SelectSingleCharacter()
        {
            string x;
            //選取單字
            if (textBox1.SelectionLength == 0 && textBox1.SelectionStart < textBox1.TextLength)
            {
                if (!insertMode)
                {
                    overtypeModeSelectedTextSetting(ref textBox1);
                }
                else
                {
                    x = textBox1.Text.Substring(textBox1.SelectionStart, 1);
                    textBox1.SelectionLength = char.IsHighSurrogate(x.ToCharArray()[0]) ? 2 : 1;
                }
            }
            x = textBox1.SelectedText;
            x = x.EndsWith(Environment.NewLine) ? x.Substring(0, x.Length - 2) : x;
            x = x.EndsWith("\n") ? x.Substring(0, x.Length - 1) : x;
            if (x != string.Empty) Clipboard.SetText(x);
            return x;
        }


        /// <summary>
        /// 在已經《古籍酷》OCR文本化後的資料，加以編輯再送出
        /// 即整理完《古籍酷》OCR的文本後送出到簡易編輯【Quick edit】的機制
        /// </summary>
        internal void PressAddKeyMethodPaste2QuickEditBox()
        {
            #region 與 Ctrl + Alt + + 同
            //PauseEvents();
            pageTextEndPosition = 0; pageEndText10 = "";
        reSelect:
            textBox1.SelectAll();
            //textBox1.Select(textBox1.TextLength, 0);

            TopMost = false;
            if (!ocrTextMode) br.BringToFront("chrome");
            if (textBox1.SelectionLength < textBox1.TextLength)
                goto reSelect;
            if (keyDownCtrlAdd(false, "", true))
            {
                //if (x != br.Quickedit_data_textboxTxt)
                //{
                //    playSound(soundLike.exam);
                //    x = br.Quickedit_data_textboxTxt;
                //}
                string x = textBox1.Text;
                //非同步整理OCR文本時，這行就很需要：（因為查字資料庫的標書名號資料表可能不同步。）
                if (x.IndexOf("，") == -1 && x.IndexOf("。") == -1
                    && (x.IndexOf("《") > -1 || x.IndexOf("〈") > -1 || x.IndexOf("：") > -1))
                    textBox1.Text = CnText.RemarkBooksPunctuation(ref x);
                else
                {
                    if (ModifierKeys != Keys.Shift)//按下Shift鍵即欲逕送去OCR，故不必再標書名號、篇名號了
                        textBox1.Text = CnText.BooksPunctuation(ref x);
                    else
                        textBox1.Text = x;
                }
                try
                {
                    //將頁面移至頂端，以便校對輸入時檢視
                    if (br.driver.Url != textBox3.Text)
                        br.GoToUrlandActivate(br.driver.Url, true);
                }
                catch (Exception)
                {
                }
            }
            bringBackMousePosFrmCenter();
            //ResumeEvents();
            #endregion
        }

        /// <summary>
        /// 記錄程式執行是否在 pagePaste2GjcoolOCR 方法套用的堆疊（stack）裡
        /// </summary>
        internal bool PagePaste2GjcoolOCR_ing = false;
        /// <summary>
        /// Ctrl + Shift + Alt + + 或 Ctrl + Alt + Shift + + （數字鍵盤加號） ： 同上，唯先將textBox1全選後再執行貼入；即按下此組合鍵則會並不會受插入點所在位置處影響。並翻到下一頁直接將它送去《古籍酷》OCR
        /// 或只按下F8
        /// 整頁貼上Quick edit [簡單修改模式]  並將下一頁直接送交《古籍酷》OCR
        /// 若欲中斷、不交去《古籍酷》OCR則須按下Ctrl
        /// 20240730 新增《看典古籍》OCR API 功能
        /// </summary>
        /// <returns>執行失敗傳回false</returns>
        private bool pagePaste2GjcoolOCR()
        {
            bool returnValue = false; PagePaste2GjcoolOCR_ing = true;
            //playSound(soundLike.press);
            //textBox1.SelectAll();//此方法必須在表單有焦點時才行
            TopMost = false;//將焦點交給Chrome瀏覽器

            bool eventsEnable = _eventsEnabled;
            if (eventsEnable) PauseEvents();
            string _lastValidWindow = br.LastValidWindow;
            if (!string.IsNullOrEmpty(_lastValidWindow)) br.driver.SwitchTo().Window(_lastValidWindow);
            //ResumeEvents();//見return前
            textBox1.SelectionStart = textBox1.TextLength; textBox1.SelectionLength = 0;
            pageTextEndPosition = 0; pageEndText10 = string.Empty;
            playSound(soundLike.waiting);//請靜待OCR完成

            if (keyDownCtrlAdd(false, "", true, true))
            {
                if (textBox1.Text != string.Empty)
                { undoRecord(); PauseEvents(); textBox1.Text = string.Empty; ResumeEvents(); }
                //欲中止，請按下Ctrl鍵
                //Console.WriteLine(ModifierKeys.ToString());//just for test
                if (ModifierKeys != Keys.Control)
                {
                    playSound(soundLike.press);
                    //if (toOCR(br.OCRSiteTitle.GJcool))
                    if (toOCR(PagePast2OCRsite))
                        returnValue = true;
                    else
                        if (!Visible) Visible = true;
                }
                else
                {
                    playSound(soundLike.stop);
                    if (!Visible) Visible = true;
                    bringBackMousePosFrmCenter();
                    _eventsEnabled = eventsEnable;
                    return false;
                }
            }
            else
            {//keyDownCtrlAdd 操作中斷，直接退出
                if (!Visible) Visible = true;
                bringBackMousePosFrmCenter();
                _eventsEnabled = eventsEnable;
                return false;
            }
            if (returnValue)
            {
                if (!Visible) Visible = true;
                bringBackMousePosFrmCenter();
            }
            else
            {
                //前後已有PauseEvents，故略去
                //PauseEvents();//為了等待時可以切到別的視窗看看，在執行成功且完成後，再把這2個關鍵的視窗置前
                try
                {
                    //br.driver.SwitchTo().Window(br.LastValidWindow);
                    br.driver?.SwitchTo().Window(br.driver.CurrentWindowHandle);
                    Thread.Sleep(300);
                }
                catch (Exception)
                {
                    //br.LastValidWindow = null;
                }
                //ResumeEvents();

            }
            //ResumeEvents();//交由呼叫端處理
            _eventsEnabled = eventsEnable;
            return returnValue;
        }

        /// <summary>
        /// Shift + F8 或 Alt + Shift + Pause ： 加上篇名格式代碼並前置2個全形空格
        /// 加上篇名格式代碼並前置N個全形空格.N，預設為2.且可在執行此項時，選取空格數以重設篇名前要空的格數
        /// </summary>
        private void keysTitleCodeAndPreWideSpace()
        {
            //若有選取全形空格數，則據以前置其數量，否則預設為2
            int spaceCnt = 0;
            //重設篇名前的全形空格（spaceStrBreforeTitle）之值
            string spaceStrBreforeTitleString = textBox1.SelectedText;
            if (spaceStrBreforeTitleString.IndexOf("　") > -1 && spaceStrBreforeTitleString.Replace("　", "") == string.Empty)
            {
                spaceCnt = textBox1.SelectedText.Length;
                spaceStrBreforeTitle = "";
                for (int i = 0; i < spaceCnt; i++)
                {
                    spaceStrBreforeTitle += "　";
                }
            }
            if (spaceStrBreforeTitleString == string.Empty)
            {
                string lineX = GetLineText(textBox1.Text, textBox1.SelectionStart);
                if (lineX.Contains("　}}"))
                {
                    _document.GetCurrentParagraph().Range.Select();
                }
            }

            if (!stopUndoRec) stopUndoRec = true; if (EventsEnabled) undoRecord();
            //undoRecord();stopUndoRec = true;            
            titleMarkCode();

            if (!GetLineText(textBox1.Text, textBox1.SelectionStart).Contains("*") &&
                textBox1.SelectionStart > 1 && textBox1.Text.Substring(textBox1.SelectionStart - 2, 2) == Environment.NewLine)
                textBox1.Select(textBox1.SelectionStart - 2, 0);
            int sPre = textBox1.Text.LastIndexOf(Environment.NewLine, textBox1.SelectionStart);
            sPre = sPre == -1 ? 0 : sPre + 2;
            textBox1.Select(sPre, textBox1.SelectionStart - sPre);
            if (!textBox1.SelectedText.StartsWith("　"))
                textBox1.SelectedText = spaceStrBreforeTitle + textBox1.SelectedText;
            stopUndoRec = false;//undoRecord();
            if (!Active)
                bringBackMousePosFrmCenter();
        }
        /// <summary>
        /// Alt + Pause 或 數字鍵盤 5（數字鍵盤「5」）： 自動判斷標題行（目前為少於12字），加上篇名格式代碼並前置2個全形空格
        /// 加上篇名格式代碼並前置N個全形空格.N，預設為2.且可在執行此項時，選取空格數以重設篇名前要空的格數
        /// 此法可與 Alt + t detectTitleYetWithoutPreSpace() 參互應用 20231018
        /// </summary>
        private void autoKeysTitleCodeAndPreWideSpace()
        {
            int wordCountLimit = 12;//少於12字才視標題
                                    //int wordCountLimit = 17;//少於17字才視標題
            if (wordCountLimit + 2 >= wordsPerLinePara) wordCountLimit = wordsPerLinePara - 2;//一般題目都是空二格故
            string x = textBox1.Text;
            int sOriginal = textBox1.SelectionStart, lOriginal = textBox1.SelectionLength, lenOriginal = x.Length;
            //清除最後末的分行/段符號
            if (x.Length > 2 && x.Substring(x.Length - Environment.NewLine.Length, Environment.NewLine.Length) == Environment.NewLine)
            {
                //pauseEvents(); stopUndoRec = true;//交由呼叫端
                textBox1.Text = textBox1.Text.Substring(0, textBox1.TextLength - Environment.NewLine.Length);
                x = textBox1.Text;
                //resumeEvents(); stopUndoRec = false;//交由呼叫端
            }
            int s = x.IndexOf(Environment.NewLine), lenLine = 0, lenLineNext = 0;
            //取得正常行/段字數
            if (wordsPerLinePara < 1) wordsPerLinePara = CountWordsLenPerLinePara(GetLineText_WithoutPunctuation(x, s));
            while (s > -1)
            {
                //StringInfo si = new StringInfo(getLineTxtWithoutPunctuation(x, s));                

                //if (si.String == "岳武穆遺詩") Debugger.Break();//just for check

                //lenLine = si.LengthInTextElements;
                lenLine = CountWordsLenPerLinePara(GetLineText_WithoutPunctuation(x, s));
                int sNext = x.IndexOf(Environment.NewLine, s + 1), previousLineLen = 0;
                if (sNext > -1)
                {
                    //StringInfo siNext = new StringInfo(getLineTxtWithoutPunctuation(x, sNext));
                    //lenLineNext = siNext.LengthInTextElements;
                    lenLineNext = CountWordsLenPerLinePara(GetLineText_WithoutPunctuation(x, sNext));
                }
                else
                {
                    if (s + 2 <= x.Length && x.Length - (s + 2) >= 0)
                        //lenLineNext = new StringInfo(x.Substring(s + 2, x.Length - (s + 2))).LengthInTextElements;
                        lenLineNext = CountWordsLenPerLinePara(x.Substring(s + 2, x.Length - (s + 2)));
                    else
                        Debugger.Break();
                }

                //if (si.String == "我今弔死三清殿知道來年荒不荒至今") Debugger.Break();//just for check

                //所在段落小於正常行長，且後面的行長須等於或大於正常行長、或是不存在後面的行/段
                if (lenLine > 0 && lenLine < wordCountLimit && !GetLineText(x, s).EndsWith("<p>"))
                //if (lenLine > 0)
                {
                    //if (lenLine < wordCountLimit)
                    //{
                    textBox1.Select(s, 0);
                    if (lenLineNext >= wordsPerLinePara)
                        keysTitleCodeAndPreWideSpace();
                    else if (lenLineNext == 0)
                    {//如果是最後一行，且前一行短於正常行長（等於者手動鍵入）
                        if (previousLineLen < wordsPerLinePara) keysTitleCodeAndPreWideSpace();
                    }
                    else
                        keysParagraphSymbol(true);
                    //}
                }
                else
                {
                    //if (lenLine < wordsPerLinePara) keysParagraphSymbol(true);
                    //else
                    previousLineLen = lenLine;
                }
                s += textBox1.TextLength - x.Length;
                x = textBox1.Text;
                s = x.IndexOf(Environment.NewLine, s + 1);
            }

            //if (lenLineNext > 0 && lenLineNext < wordCountLimit)
            if (lenLineNext > 0)// && !getLineTxt(x, s).EndsWith("<p>"))
            {
                if (lenLineNext < wordCountLimit)
                {
                    textBox1.Select(x.Length, 0);
                    if (lenLine >= wordsPerLinePara)
                        //其後一行若短於正常行長，且其本身不短於正常行長，                    
                        keysParagraphSymbol(true);//則視為標題
                    else//否則只當作段落
                        keysTitleCodeAndPreWideSpace();
                }
                else
                {
                    if (lenLineNext < wordsPerLinePara)
                    {
                        textBox1.Select(x.Length, 0);
                        keysParagraphSymbol(true);
                    }
                }
            }

            #region 最後檢查            
            clearParagraphMarkersBetweenPairsBrackets();
            x = textBox1.Text;//使用不再使用的變數
            CnText.FormalizeText(ref x);
            textBox1.Text = x;
            clearParagraphMarkersInsidePairsBrackets();
            movePeriodsToFrontofBlank();
            #endregion

            //處理完畢，回到原來/開始執行時的位置（有時末尾有些不要的贅文，可從此處在送出時截斷）
            //再研究
            //s = textBox1.se
            //textBox1.Select(sOriginal + (s - (lenOriginal-(s-text(lenOriginal-sOriginal)))), lOriginal);            
            //textBox1.Select(sOriginal + (textBox1.TextLength - lenOriginal), lOriginal);
            //先假設必然在原來插入點之前會增加符號文字，再找接在其後、最近的分行/段符號位置，以此位置權作新的定位，至少有某個程度的可靠，省卻某一部分冗餘的操作20231019
            if (sOriginal > textBox1.TextLength)
                s = -1;
            else
                s = textBox1.Text.IndexOf(Environment.NewLine, sOriginal);
            sOriginal = s == -1 ? textBox1.TextLength : s;
            textBox1.Select(sOriginal, lOriginal);
        }

        /// <summary>
        /// 作為以選取範圍為格式化依據，將上下兩欄的目次內容格式化 20231018
        /// formatCategory2Columns函式的參照
        /// 第2次以後執行時的前參考準據
        /// </summary>
        int wordCntBeforeNextColume = 0;

        /// <summary>
        /// 以選取範圍為格式化依據，將上下兩欄的目次內容格式化 20231018
        /// 在textBox2中輸入「fc」以執行（取format,Category二字首，故為fc）
        /// 執行時若無選取，則以之前的設定為準。若第一次，請務必要選取以供指定
        /// 如果行末是<p>（不含分行/段符號）則停止處理
        /// </summary>
        private void formatCategory2Columns_GjcoolOCRResult()
        {
            //以選取範圍為格式化依據
            string xSel = textBox1.SelectedText;
            if (xSel == string.Empty && wordCntBeforeNextColume == 0)
            {
                MessageBoxShowOKExclamationDefaultDesktopOnly("請先選取要作為依據的行/段"); return;
            }
            string x = textBox1.Text; int e;
            if (xSel != string.Empty)
            {
                e = xSel.LastIndexOf("􏿽");
                e += "􏿽".Length;
                //取得下一欄之前的字數
                //if (wordCntBeforeNextColume == 0 && xSel != )
                wordCntBeforeNextColume = new StringInfo(GetLineText_WithoutPunctuation(xSel.Substring(0, e), e)).LengthInTextElements;
                if (e == -1) return;
            }
            int s = xSel == string.Empty ? 0 : x.IndexOf(Environment.NewLine, xSel == string.Empty ? 0 : textBox1.SelectionStart)
                , cntr = xSel == string.Empty ? -1 : 0;

            undoRecord();
            PauseEvents(); stopUndoRec = true;



            //自動補上末尾的段落符號，以利while迴圈判斷
            if (x.Length > 2 && x.Substring(x.Length - 2, 2) != Environment.NewLine) x += Environment.NewLine;
            textBox1.Text = x;

            while (s > -1)
            {

                cntr++;
                //if (cntr % 2 == (xS/*e*/l == string.Empty ? 0 : 1)) continue;
                if (cntr % 2 == 1) continue;

                if (cntr > 0)
                    s += Environment.NewLine.Length;

                //清除段落（下一個行/段併到上一個）
                if (x.IndexOf(Environment.NewLine, s) < 0) break;
                //如果行末是<p>（不含分行/段符號）則停止處理
                if (GetLineText(x, s).EndsWith("<p>")) break;
                textBox1.Select(x.IndexOf(Environment.NewLine, s), Environment.NewLine.Length);
                textBox1.SelectedText = string.Empty;
                x = textBox1.Text;
                string xLine = GetLineText(x, s);
                int xLineLen = xLine.Length; e = xLine.LastIndexOf("􏿽") + "􏿽".Length;
                while (new StringInfo(GetLineText_WithoutPunctuation(xLine.Substring(0, e), e)).LengthInTextElements < wordCntBeforeNextColume)
                {
                    xLine = xLine.Substring(0, e) + "􏿽" + xLine.Substring(e);
                    e = xLine.LastIndexOf("􏿽") + "􏿽".Length;
                }

                textBox1.Select(s, xLineLen);
                textBox1.SelectedText = xLine; x = textBox1.Text;
                s = x.IndexOf(Environment.NewLine, s + 1);
            }

            x = textBox1.Text;
            textBox1.Text = x.Replace(Environment.NewLine, "|" + Environment.NewLine);
            //清除自動補上末尾的段落標記<p>
            if (textBox1.TextLength > 2)
            {
                x = textBox1.Text;
                if (x.Substring(x.Length - Environment.NewLine.Length, Environment.NewLine.Length) == Environment.NewLine)
                    textBox1.Text = x.Substring(0, x.Length - Environment.NewLine.Length);
            }

            #region 清除最後落單的欄後面的冗餘􏿽􏿽
            x = textBox1.Text;
            if (x.Length > 2 && x.Substring(x.Length - 3, 3) == "􏿽|")
            {
                //清除末尾的「|」
                x = x.Substring(0, x.Length - 1);
                while (x.Length > -1 && x.Substring(x.Length - "􏿽".Length, "􏿽".Length) == "􏿽")
                {
                    x = x.Substring(0, x.Length - "􏿽".Length);
                }
                //末尾改成用「<p>」
                x += "<p>";
                textBox1.Text = x;
            }
            #endregion

            textBox1.Select(textBox1.TextLength, 0);

            ResumeEvents(); stopUndoRec = false;
        }

        /// <summary>
        /// 切換「插入」輸入模式與「取代」輸入模式
        /// </summary>
        internal void InsertModeSwitcher()
        {
            if (insertMode)
            {
                insertMode = false;
                textBox1.Font = new Font(textBox1.Font.FontFamily, textBox1.Font.Size, FontStyle.Bold);
                Caret_Shown_OverTypeMode(textBox1);
            }
            else
            {
                insertMode = true;
                textBox1.Font = new Font(textBox1.Font.FontFamily, textBox1.Font.Size, FontStyle.Regular);
                Caret_Shown(textBox1);
            }
        }

        /// <summary>
        /// 取代模式時的選字行為
        /// 此法可與CnText類別中的 ChangeSeltextWhenOverwriteMode 互用。該法「不會」改變文字方塊的選取範圍，只會取得其「改變」後的實際值
        /// </summary>
        /// <param name="tBox">要操作的文字方塊對象。傳址（pass by reference）</param>        
        /// <returns>回傳文字方塊中，被重新選取的字</returns>
        string overtypeModeSelectedTextSetting(ref TextBox tBox)
        {
            if (!insertMode)
            {
                int s = tBox.SelectionStart, l = tBox.SelectionLength;
                if (s + l < tBox.TextLength)
                {
                    string nextchar = tBox.Text.Substring(s + l, 1);
                    l += char.IsHighSurrogate(nextchar.ToCharArray()[0]) || nextchar == Environment.NewLine.Substring(0, 1) ? 2 : 1;
                }
                l = s + l > tBox.TextLength ? l - 1 : l;
                //改變選取範圍
                tBox.Select(s, l);
            }
            return tBox.SelectedText;
        }

        internal bool examSeledWord(out string wordtoChk)
        {
            int s = textBox1.SelectionStart;
            string x = textBox1.Text;
            if (textBox1.SelectedText != "") wordtoChk = textBox1.SelectedText;
            else if (s + 2 <= textBox1.TextLength)
                wordtoChk = x.Substring(s, char.IsHighSurrogate(x.Substring(s, 1).ToCharArray()[0]) ? 2 : 1);
            else
            {
                if (s + 1 > x.Length) { wordtoChk = ""; return false; }
                wordtoChk = x.Substring(s, 1);
            }

            return isChineseChar(wordtoChk, false) != 0;
        }

        private void keySymbols(string symbol)
        {
            undoRecord();
            if (textBox1.SelectionLength > 0)
            {
                textBox1.SelectedText = textBox1.SelectedText.Replace("　", symbol).Replace("􏿽", symbol).Replace("<p>", symbol).Replace("*", "");
            }
            else
            {
                textBox1.SelectedText = symbol; int s = textBox1.SelectionStart;
                if ((s + 2) <= textBox1.TextLength && "　􏿽*".IndexOf(textBox1.Text.Substring(s, 1), 0, StringComparison.Ordinal) > -1)
                {
                    textBox1.Select(s, char.IsHighSurrogate(textBox1.Text.Substring(s, 1).ToCharArray()[0]) ? 2 : 1); textBox1.SelectedText = "";
                }
                else if ((s + 3) <= textBox1.TextLength && "<p>" == textBox1.Text.Substring(s, 3))
                {
                    textBox1.Select(s, 3); textBox1.SelectedText = "";
                }
            }
        }

        private void addData四部叢刊造字對照表andReplace()
        {
            //Alt + 4 : 新增【四部叢刊造字對照表】資料並取代其造字,若無選取文字以指定文字，則加以取代
            //throw new NotImplementedException();
            //選取文字第一個是造字，第2個是系統字（CJK）
            string x = textBox1.SelectedText;
            ado.Connection cnt = new ado.Connection();
            ado.Recordset rst = new ado.Recordset();
            Mdb.openDatabase("查字.mdb", ref cnt);
            if (x != "")
            {
                StringInfo xInfo = new StringInfo(x);
                if (xInfo.LengthInTextElements == 2)
                {

                    string w = xInfo.SubstringByTextElements(0, 1);
                    rst.Open("select 造字,字 from 四部叢刊造字對照表 where strcomp(造字,\"" + w + "\")=0",
                        cnt, ado.CursorTypeEnum.adOpenKeyset, ado.LockTypeEnum.adLockOptimistic);
                    if (rst.RecordCount == 0)
                    {
                        rst.AddNew();
                        rst.Fields[0].Value = w;
                        rst.Fields[1].Value = xInfo.SubstringByTextElements(1, 1);
                        rst.Update();
                    }
                    textBox1.SelectedText = xInfo.SubstringByTextElements(0, 1);
                    rst.Close();
                }
            }
            rst.Open("select 造字,字 from 四部叢刊造字對照表", cnt, ado.CursorTypeEnum.adOpenForwardOnly
                , ado.LockTypeEnum.adLockReadOnly);
            x = textBox1.Text;
            while (!rst.EOF)
            {
                if (x.IndexOf(rst.Fields[0].Value.ToString(), StringComparison.Ordinal) > -1)
                    x = x.Replace(rst.Fields[0].Value.ToString(), rst.Fields[1].Value.ToString());
                rst.MoveNext();
            }
            undoRecord();
            stopUndoRec = true;
            caretPositionRecord();
            textBox1.Text = x;
            stopUndoRec = false;
            restoreCaretPosition(textBox1, selStart, selLength);//textBox1.SelectionStart, textBox1.SelectionLength);
            caretPositionRecall();
            if (textBox1.SelectionLength > 0)
            {
                if (char.IsLowSurrogate(textBox1.Text.Substring(textBox1.SelectionStart + textBox1.SelectionLength, 1).ToCharArray()[0]))
                {
                    textBox1.Select(textBox1.SelectionStart, ++textBox1.SelectionLength);
                }
            }
            else
            {
                if (textBox1.TextLength < textBox1.SelectionStart + 1)
                    textBox1.Select(textBox1.SelectionStart, 1);
                else
                {
                    if (char.IsLowSurrogate(textBox1.Text.Substring(textBox1.SelectionStart, 1).ToCharArray()[0]))
                    {
                        textBox1.Select(++textBox1.SelectionStart, 0);
                    }
                }
            }

            rst.Close(); cnt.Close();
        }

        private void markParagraphwithSelectionLen()
        {//Alt + Shift + q : 據選取區的CJK字長以作分段（末後植入 < p >，分行則以版式常態值劃分），為非《維基文庫》版式之電子文本，如《寒山子詩集》組詩
         //throw new NotImplementedException();
            string x = textBox1.Text; int p = x.IndexOf("<p>");
            if (p == -1) return;
            if (textBox1.SelectionLength == 0)
            {
                textBox1.Select(0, p);//第一次須是指定段落長度者，且其後文字須無分段/行
            }
            if (NormalLineParaLength == 0) NormalLineParaLength = x.IndexOf(Environment.NewLine);
            string xSl = textBox1.SelectedText; x = x.Substring(textBox1.SelectionStart + textBox1.SelectionLength + Environment.NewLine.Length + "<p>".Length);
            StringInfo xSlInfo = new StringInfo(xSl);//, xInfo = new StringInfo(x);
            int l = xSlInfo.LengthInTextElements, s = 0;//取得段落長
                                                        // lInfo = xInfo.LengthInTextElements
                                                        //TextElementEnumerator xEl = StringInfo.GetTextElementEnumerator(x);
            while (s + l < x.Length)
            {
                int iPara = 0, iLine = 0, sLine = s;
                while (s + l < x.Length && new StringInfo(x.Substring(s, ++iPara)).LengthInTextElements < l)
                {//取得段長位置
                    if (new StringInfo(x.Substring(sLine, ++iLine)).LengthInTextElements == NormalLineParaLength)
                    {

                        if (char.IsLowSurrogate(x.Substring(sLine + iLine, 1).ToCharArray()[0])) iLine++;
                        x = x.Substring(0, sLine + iLine) + Environment.NewLine + x.Substring(sLine + iLine);
                        sLine += iLine; iLine = 0;
                        sLine += 2;  //Environment.NewLine.Length;
                    }
                }
                s += iPara; //iPara = 0;
                if (char.IsLowSurrogate(x.Substring(s, 1).ToCharArray()[0])) s++;
                x = x.Substring(0, s) + "<p>" + Environment.NewLine + x.Substring(s);
                s += 5;//"<p>" + Environment.NewLine
                       //if (s > 5000) break;
            }
            undoRecord();
            stopUndoRec = true;
            textBox1.Text = xSl + "<p>" + Environment.NewLine + x;
            stopUndoRec = false;
        }
        /// <summary>
        /// 清除插入點之前的所有空格「　」、空白「􏿽」、「<p>」、「}}」、「*」
        /// </summary>
        private void 清除插入點之前的所有空格()
        {//Ctrl + Backspace,若插入點前為「<p>」則一併清除
         //throw new NotImplementedException();
            int s = textBox1.SelectionStart, e = s; string x = textBox1.Text;
            if (s > 0 && x[s - 1] == '*')
            {
                while (s > 0 && x[--s] == '*') { }
                textBox1.Select(s + 1, e - s - 1);
            }
            else if (s > 3 && x.Substring(s - 3, 3) == "<p>")
            {
                textBox1.Select(s - 3, 3);
            }
            else if (s > 2 && x.Substring(s - 2, 2) == "}}")
            {
                textBox1.Select(s - 2, 2);//2="}}".Length
            }
            else if (x.Contains("〖"))
            {
                clearWhiteSquareBracketsRange(s, x, true);
            }
            else
            {
                while (s > 1 && "　􏿽".IndexOf(x.Substring(--s, 1)) > -1) { }
                textBox1.Select(s + 1, e - s - 1);
            }
            undoRecord();
            stopUndoRec = true; PauseEvents();
            textBox1.SelectedText = string.Empty;
            stopUndoRec = false; ResumeEvents();
        }
        /// <summary>
        /// 清除 的全部內容，如印鑑或框圍之全部內容 20250325
        /// </summary>
        /// <param name="s">插入點位置，如textBox1的</param>
        /// <param name="x">全文本內容，如textBox1的Text屬性值</param>
        /// <param name="back">如果是往後（前）查找，如按下Ctrl + BackSpace 鍵</param>
        /// <returns>有執行成功則傳回true</returns>
        private bool clearWhiteSquareBracketsRange(int s, string x, bool back = false)
        {//White Square Brackets（空心方括號 〖〗） ，乃詢問Gemini大菩薩的：https://g.co/gemini/share/0718f11f8a9c 
            int pos = x.IndexOf("〗");
            if (pos == -1) return false;
            if (back)
            {
                if (pos == s - 1)
                {
                    while (pos - 1 > -1 && x.Substring(--pos, 1) != "〖") { }
                }
            }
            else
            {//清除〖〗的全部內容
                pos = x.IndexOf("〖"); if (pos == -1) return false;
                if (pos == s)
                {
                    while (pos + 1 <= x.Length && x.Substring(++pos - 1, 1) != "〗") { }
                }
            }
            textBox1.Select(pos, s - pos);
            undoRecord();
            textBox1.SelectedText = string.Empty;
            undoRecord();
            return true;
        }


        /// <summary>
        /// 前後加上指定符號－－ 由「加上黑括號」改編擴展而來
        /// </summary>
        /// <param name="whatSymbol">成對的符號，如果前後一致，如若想前後加上「●」，則要傳入「●●」，或只傳一個「●」</param>
        private void preceded_followed_specify_symbols(string whatSymbol)
        {//`： 於插入點處起至「　」或「􏿽」前止之文字加上黑括號【】//Print/SysRq 為OS鎖定不能用
         //throw new NotImplementedException();
            int s = textBox1.SelectionStart; string x = textBox1.Text;
            if (textBox1.SelectionLength == 0)
            {
                //；若插入點位置前不是「　􏿽」等，則移至該處
                while (s > 0 && (Environment.NewLine + "　|>}" + "􏿽".Substring(1, 1)).IndexOf(x.Substring(--s, 1)) == -1)
                {

                }
                int so = s;//記下起始處
                while (s + 1 < x.Length && (Environment.NewLine + "　|<{" + "􏿽".Substring(0, 1)).IndexOf(x.Substring(++s, 1)) == -1)
                //while (s + 1 <= x.Length && (Environment.NewLine + "　|<{" + "􏿽".Substring(0, 1)).IndexOf(x.Substring(++s, 1)) == -1)
                {

                }
                textBox1.Select(so == 0 ? so : ++so, s - so);
            }
            else//如果非插入點，則將選取區前後加上黑括號
            { }
            undoRecord(); stopUndoRec = true;
            //成對的符號，如果前後一致，如若想前後加上「●」，則要傳入「●●」，或只傳一個「●」
            StringInfo si = new StringInfo(whatSymbol);
            //textBox1.SelectedText = "【" + textBox1.SelectedText + "】";
            if (si.LengthInTextElements > 1)
                textBox1.SelectedText = si.SubstringByTextElements(0, 1) + textBox1.SelectedText + si.SubstringByTextElements(1);
            else
                textBox1.SelectedText = si.SubstringByTextElements(0) + textBox1.SelectedText + si.SubstringByTextElements(0);
            stopUndoRec = false;
        }

        private void poetryFormat()
        {
            if (textBox1.SelectionLength == 0) return;
            undoRecord(); stopUndoRec = true;
            int s = textBox1.SelectionStart;
            string xSel = textBox1.SelectedText;
            if (xSel.Length > 2 && xSel.Substring(xSel.Length - 3) == "<p>") { textBox1.Select(s, xSel.Length - 3); xSel = textBox1.SelectedText; }//最後一個<p>不處理
            if (xSel.Length > 4 && xSel.Substring(xSel.Length - 5) == "<p>" + Environment.NewLine) { textBox1.Select(s, xSel.Length - 5); xSel = textBox1.SelectedText; }//最後一個<p>不處理
            //xSel = xSel.Replace("<p>", "|").Replace("　", "􏿽");
            xSel = xSel.Replace("<p>", "|");//.Replace("　", "􏿽");
            //xSel = Regex.Replace(xSel, @"　?(!}})", "􏿽");
            //xSel = Regex.Replace(xSel, @"　*(!}})", "􏿽");// Copilot大菩薩：// 使用負向環視 (?<! ) 確保不替換 "}}" 前的空格 20250306 https://copilot.microsoft.com/shares/5vKyp5oM72zFwJmbboSMw
            xSel = ReplaceSpaces(xSel);
            string ReplaceSpaces(string input)//20250306 Gemini大菩薩：https://g.co/gemini/share/d5b66edbaea1
            {
                // 修正後的正規表達式模式：
                // (?!(\\s*\\}\\})|(\\{\\{.*?\\}\\})) - 負向先行斷言，確保不匹配 }} 前面的連續空白，以及 {{}} 標籤內的空白
                // 　                                  - 匹配全形空白
                string pattern = "(?!((\\s*\\}\\})|(\\{\\{.*?\\}\\})))　";
                return Regex.Replace(input, pattern, "􏿽");
            }

            if (xSel.IndexOf("*") > -1)
            {
                xSel = xSel.Replace("*", "");
                xSel = xSel.Replace(Environment.NewLine, "|" + Environment.NewLine).Replace("||", "|");
            }
            xSel = xSel.Replace("。", string.Empty);
            textBox1.SelectedText = xSel;
            //string tx = textBox1.Text;
            //replaceXdirectly(ref tx);
            textBox1.Select(s, xSel.Length);
            stopUndoRec = false;
        }

        void notes_a_line_all(bool ctrl, bool onlyUnderTitle = false)
        {//Alt + Shift + s :  所有小注文都不換行//Alt + Shift + Ctrl + s : 小注文不換行(短於指定漢字長者)
            int s = textBox1.SelectionStart, i = textBox1.Text.IndexOf("}}"), space = 0;
            //'if (textBox1.SelectedText == "") textBox1.SelectAll();
            undoRecord();
            stopUndoRec = true;
            while (i > -1)
            {
                if ((textBox1.Text.LastIndexOf(Environment.NewLine, i) == -1 && textBox1.Text.LastIndexOf("{{", i) > -1)
                    || (textBox1.Text.LastIndexOf(Environment.NewLine, i) < textBox1.Text.LastIndexOf("{{", i)))
                {
                    if (onlyUnderTitle)
                    { if (GetLineText(textBox1.Text, i).IndexOf("*") == -1) goto omit; }

                    textBox1.Select(i, 0);
                    space = notes_a_line(false, ctrl);
                omit:
                    if (textBox1.TextLength >= i + space + 1)
                        i = textBox1.Text.IndexOf("}}", i + space + 1);
                    else
                        i = -1;
                }
                else
                    i = textBox1.Text.IndexOf("}}", i + 1);
            }
            stopUndoRec = false;
            textBox1.Select(s, 0); textBox1.ScrollToCaret();
        }

        /// <summary>
        /// 自動執行「小注文不換行」的漢字數限定
        /// </summary>
        byte noteinLineLenLimit = 3;
        /// <summary>
        /// 小注文不換行 Alt + Shift + 6 或 Alt + s 
        /// </summary>
        /// <param name="undoRe"></param>
        /// <param name="ctrl"></param>
        /// <returns></returns>
        private int notes_a_line(bool undoRe = true, bool ctrl = false)
        {
            textBox1.DeselectAll();
            string xSel = textBox1.SelectedText, x = textBox1.Text; int s = textBox1.SelectionStart; bool flg = false;
            //如果插入點在最末端
            if (s == x.Length)
            {
                //若最末之字元為「}」則改定插入點位置，照常執行
                if (s - 1 > 0 && x.Substring(s - 1, 1) == "}")
                    textBox1.Select(--s, 0);
                else
                    return 0;
            }
            if (undoRe)
            {
                undoRecord(); stopUndoRec = true;
            }
            if (x.Substring(0, s).IndexOf("{{") == -1)
            {
                textBox1.Text = "{{" + x;//暫時補的「{{」
                x = textBox1.Text;
                s += 2; flg = true;
            }
            #region expand the notes text to get it
            if (xSel == "")
            {

                while (s - 1 > 0)
                {
                    if (x.Substring(s, 1) == "{")
                    {
                        break;
                    }
                    s--;
                }

            }
            #endregion
            s++;
            int e = x.IndexOf("}", s), spaceCntr = 0;
            if (e < 0)
            {
                MessageBox.Show("請在注文末端加入「}}」再繼續！");
                return 0;
            }
            xSel = x.Substring(s, e - s); int i = 0;
            if (xSel != string.Empty)
            {
                #region 如果末已綴有空格
                if (xSel.Length > 0 && xSel.Replace("　", "") != string.Empty)//如果剛好是造字圖，如「㽦」从「由」不从「田」者 20240303整理《四庫》本《朱子語類》時（朱子語類卷一百七~卷一百十 https://ctext.org/library.pl?if=en&file=2300&page=147#%E3%BD%A6 ）
                {
                    while (xSel.Substring(xSel.Length - ++spaceCntr, 1) == "　") { }
                }
                spaceCntr--;
                #endregion //如果末已綴有空格
                StringInfo xSelInfo = new StringInfo(xSel.Substring(0, xSel.Length - spaceCntr).Replace(Environment.NewLine, ""));
                if (ctrl)
                {
                    if (xSelInfo.LengthInTextElements >= noteinLineLenLimit || xSelInfo.LengthInTextElements == 1)//Alt + Shift + Ctrl + s : 小注文不換行(短於指定漢字長者)
                    {
                        return 0;
                    }
                    else
                    {
                        if ((s - 6 >= 0 && x.Substring(s - 6, 2) == "}}" && x.Substring(s - 4, 2) == Environment.NewLine)
                            || (e + 4 + 2 <= x.Length && x.Substring(e + 4, 2) == "{{" && x.Substring(e + 2, 2) == Environment.NewLine))
                        {//如果注文換行
                            return 0;
                        }
                        else if ((s - 7 >= 0 && x.Substring(s - 7, 2) == "}}" && x.Substring(s - 5, 2) == Environment.NewLine && x.Substring(s - 3, 1) == "　")
                            || (e + 5 + 2 <= x.Length && x.Substring(e + 5, 2) == "{{") && x.Substring(e + 2, 2) == Environment.NewLine && x.Substring(e + 4, 1) == "　")
                        {//縮排一格（字）；先只作縮排一格的，若縮排2字以上可另寫類推。
                            return 0;
                        }
                    }
                }
                for (i = 0; i + spaceCntr < xSelInfo.LengthInTextElements; i++)
                {
                    if (PunctuationsNum.IndexOf(xSelInfo.SubstringByTextElements(i, 1)) == -1)
                        xSel += "　";
                }
                textBox1.Select(s, e - s);
                textBox1.SelectedText = xSel;
            }

            if (flg)
            {
                textBox1.Text = textBox1.Text.Substring(2);//還原暫時補的「{{」
            }

            if (undoRe) stopUndoRec = false;
            return i;
            //throw new NotImplementedException();
        }

        /// <summary>
        /// Alt + 2 : 鍵入全形空格「　」space（非空白「􏿽」blank）
        /// </summary>
        private void keysSpaces()
        {
            string selX = textBox1.SelectedText;
            int s = textBox1.SelectionStart, l = textBox1.SelectionLength; string x = textBox1.Text;
            if (s + l + 2 <= x.Length && s - 2 >= 1)
            {
                if (x.Substring(s + l, 2) == "􏿽" || x.Substring(s - 2, 2) == "􏿽")
                {//自動把插入點所在處前後「􏿽」置換成「　」
                    undoRecord();
                    stopUndoRec = true; s--;
                    while (s >= 0 && (x.Substring(s, 1) == "\udbff" || x.Substring(s, 1) == "\udffd"))
                    {
                        s--;
                    }
                    l = 0; s++;
                    while (s + l + 1 <= x.Length && (x.Substring(s + l, 1) == "\udffd" || x.Substring(s + l, 1) == "\udbff")) { l++; }
                    textBox1.Select(s, l);
                    textBox1.SelectedText = textBox1.SelectedText.Replace("􏿽", "　");
                    stopUndoRec = false;
                    return;
                }
            }
            else if (s <= 2 && x.Length > 1 && x.Substring(s, 2) == "􏿽")
            {//插入點在textBox1最前端時的處理
                if (selX == "")
                {
                    int sLen = 2;
                    while (s + 2 < x.Length && x.Substring(s, sLen) == "􏿽")
                    {
                        sLen += 2;
                    }
                    textBox1.Select(s, sLen -= 2); selX = textBox1.SelectedText;
                }
            }
            undoRecord();
            stopUndoRec = true;
            if (selX != "")
                if (selX == " ")
                    textBox1.SelectedText = selX.Replace(" ", "　");
                else
                    textBox1.SelectedText = selX.Replace("􏿽", "　");
            else
            {
                while (textBox1.SelectionStart < textBox1.TextLength && textBox1.Text.Substring(textBox1.SelectionStart, 1) == " ")
                {//清除插入點末的半形空格
                    textBox1.Select(textBox1.SelectionStart, 1);
                    textBox1.SelectedText = string.Empty;
                }
                //insertWords("　", textBox1, textBox1.Text);
                insertWords("　", textBox1);

            }
            stopUndoRec = false;
        }

        int expandSelRange(string what, string domain)
        {
            int i = 0, s = textBox1.SelectionStart, sn = s - i; string p = domain.Substring(sn, 3);
            while (p != what && i < 4)
            {
                sn = s - i++;
                p = domain.Substring(sn, 3);
            }
            if (p == what)
            {
                //textBox1.Select(sn, 3);
                return sn;
            }
            else
            {
                //MessageBox.Show("not found", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1
                //      , MessageBoxOptions.ServiceNotification);
                //textBox1.DeselectAll();
                return 0;
            }
            //return textBox1.SelectedText;

        }
        /// <summary>
        /// 在指定文本中計算某單字出現的次數
        /// 若是要找分段符號 Environment.NewLine，則只需輸入 \r 或 \n 來查找計算即可
        /// </summary>
        /// <param name="whatWord">任何字串（不必是單個漢字）</param>
        /// <param name="domain"></param>
        /// <returns></returns>
        public static int CountWordsinDomain(string whatWord, string domain)
        {
            StringInfo dw = new StringInfo(domain); int cntr = 0;
            for (int i = 0; i < dw.LengthInTextElements; i++)
            {
                // SubstringByTextElements方法於分段符號或切成 \r \n 兩個部分
                //if (dw.SubstringByTextElements(i, 1) == whatWord)
                int siLength = new StringInfo(whatWord).LengthInTextElements;
                if (i + siLength > dw.LengthInTextElements) break;
                if (dw.SubstringByTextElements(i, siLength) == whatWord)
                {
                    cntr++;
                }
            }
            return cntr;
        }
        /// <summary>
        /// Alt + Shift + 1 如宋詞中的換片空格，只將文中的空格轉成空白，其他如首綴前罝以明段落或標題者不轉換
        /// </summary>
        private void SpacesBlanksInContext()
        {
            bool notTitleIndent = true; int s = textBox1.SelectionStart, offset = 0;//記下位移數（因為「　」與「􏿽」的Length不同
            //string x = textBox1.SelectedText;
            //if (x == "") x = textBox1.Text;
            string x = textBox1.Text;
            undoRecord(); stopUndoRec = true;
            for (int i = 0; i < x.Length; i++)//用每個字去算
            {
                if (x.Substring(i, 1) == "　")//逐如果字比對，如果是「　」（space空格）
                {
                    //if (GetLineText(x, i).Contains("真珠簾下")) Debugger.Break();

                    #region }}前的全形空格不處理 20250304
                    int spsCount = 1;
                    while (i + 1 <= x.Length && x.Substring(i + spsCount, 1) == "　")
                    {
                        spsCount++;
                    }
                    //if (spsCount > 1) spsCount--;
                    if (i + spsCount + 2 <= x.Length && (x.Substring(i + spsCount, 2) == "}}" || x.Substring(i + spsCount, 1) == "*"))
                    {
                        i += spsCount;
                        continue;
                    }
                    #endregion                    
                    #region 作為縮排的空格也不處理(正文者詳後，此處理注文的縮排）
                    string currentLine = GetLineText(x, i, out int lineS, out int lineL);

                    if (currentLine.Contains("不忍移災於卿佐曰移腹心之疾寘諸")) Debugger.Break();
                    //if (currentLine.Contains("杜預曰：慆疑")) Debugger.Break();
                    //if (x.Substring(i-1,1)== "述") Debugger.Break();

                    if ((currentLine.StartsWith("{{　") || (currentLine.StartsWith("　") && currentLine.Contains("{{") == false))//獨立注文的縮排
                        && ((currentLine.EndsWith("}}") || currentLine.EndsWith("}}<p>") || currentLine.EndsWith("}}。<p>")
                                            && currentLine.LastIndexOf("}}", currentLine.LastIndexOf("}}")) == -1)
                        || (currentLine.StartsWith("{{　") && currentLine.IndexOf("}}") == -1))
                        /*如：􏿽臣光曰：晏嬰有言天命不慆不貳其命{{晏子對齊侯禳彗之辭也}}
                            {{　杜預曰：慆疑　也音他刀翻}}禍福之至安可移乎昔楚昭王宋景公……
                        */
                        || (currentLine.StartsWith("{{　") && currentLine.IndexOf("}}", 3) > currentLine.IndexOf("　", 3)
                            && CountWordsinDomain("　", x.Substring(lineS, i - lineS)) < 2)//3="{{　".Length 
                        )
                    {

                        spsCount = 1;
                        while (i + 1 <= x.Length && x.Substring(i + spsCount, 1) == "　")
                        {
                            spsCount++;

                        }
                        //if (spsCount > 1) spsCount--;
                        i += spsCount;
                        continue;
                    }

                    #endregion
                    if (i > 1 && x.Substring(i - 2, 2) != Environment.NewLine && notTitleIndent)
                    {
                        if (i - 3 > -1 && textBox1.Text.Substring(i - 3, 3) != "<p>")//<p>後面的全形空格（　）也不取代為空白（􏿽）
                        {
                            textBox1.Select(i + offset, 1);
                            if (textBox1.SelectedText == "　")
                            {
                                textBox1.SelectedText = "􏿽";
                                offset++;
                            }
                        }
                    }
                    else
                    {
                        //if (x.Substring(i - 5, 3) == "<p>")
                        int p = x.IndexOf(Environment.NewLine, i);
                        if (x.Substring(i, (p > -1 ? p : x.Length) - i).IndexOf("*") > -1)
                        {
                            notTitleIndent = false;
                        }

                        //if (currentLine.Contains("侍御㟧渚山春")) Debugger.Break();


                        //前、後一行（段）開頭都不是空格space（縮排），且不是標題時
                        if (x.Length > p + Environment.NewLine.Length + 1 && x.Substring(p + Environment.NewLine.Length, 1) != "　"
                            && x.LastIndexOf(Environment.NewLine, i) > -1
                            && x.LastIndexOf(Environment.NewLine, x.LastIndexOf(Environment.NewLine, i)) > 0
                            && x.Substring(x.LastIndexOf(Environment.NewLine, x.LastIndexOf(Environment.NewLine, i)) + Environment.NewLine.Length, 1) != "　")
                        {//縮排1個空格者
                            string xLine = GetLineText(x, i);
                            #region 偵錯檢查用
                            //if (xLine.IndexOf("明後世猶或")>-1)
                            //{
                            //    xLine = xLine;
                            //}
                            #endregion

                            int preParaEnd = x.LastIndexOf(Environment.NewLine, i);
                            int preParaStart = x.LastIndexOf(Environment.NewLine, preParaEnd - 1) + 2; string prePara = string.Empty;
                            if (preParaStart > -1)
                                prePara = x.Substring(preParaStart, preParaEnd - preParaStart);
                            if (xLine.Substring(1, 1) != "　" && xLine.IndexOf("*") == -1
                                && (!prePara.StartsWith("{{　") && !prePara.StartsWith("　"))
                                && (!(prePara.StartsWith("􏿽") && CountWordsLenPerLinePara(prePara) == wordsPerLinePara))
                                //不是「本行短於正常行長，而前一行等於正常行常」
                                && !(CountWordsLenPerLinePara(xLine) < wordsPerLinePara && CountWordsLenPerLinePara(prePara) == wordsPerLinePara)
                                && (CountWordsLenPerLinePara(xLine) == wordsPerLinePara && CountWordsLenPerLinePara(prePara) == wordsPerLinePara
                                        && !xLine.EndsWith("<p>"))
                                )
                            {
                                textBox1.Select(i + offset, 1);
                                if (textBox1.SelectedText == "　")
                                {
                                    textBox1.SelectedText = "􏿽";
                                    offset++;
                                }
                                notTitleIndent = true;
                            }
                        }
                        else//縮排2個空格以上者
                        {
                            #region 前、後一行/段若是縮排                            
                            int indentSpsCount = 0, adjacentLineS = adjacentLineS = textBox1.Text.LastIndexOf(Environment.NewLine, i);
                            string adjacentLineText;
                            if (adjacentLineS > -1)
                            {
                                adjacentLineS = textBox1.Text.LastIndexOf(Environment.NewLine, adjacentLineS);
                                if (adjacentLineS > -1)
                                {
                                    adjacentLineText = GetLineText(textBox1.Text, adjacentLineS);
                                    if (adjacentLineText.Contains("*") == false && _leadingSpacesRegex.Match(adjacentLineText).Value.Length > 0)
                                        indentSpsCount = _leadingSpacesRegex.Match(adjacentLineText).Value.Length;
                                }
                            }
                            if (indentSpsCount == 0)
                            {
                                //找後面一行/段：
                                adjacentLineS = textBox1.Text.IndexOf(Environment.NewLine, i) + Environment.NewLine.Length;
                                if (adjacentLineS > -1)
                                {
                                    adjacentLineText = GetLineText(textBox1.Text, adjacentLineS);
                                    if (adjacentLineText.Contains("*") == false && _leadingSpacesRegex.Match(adjacentLineText).Value.Length > 0)
                                        indentSpsCount = _leadingSpacesRegex.Match(adjacentLineText).Value.Length;
                                }

                            }
                            #endregion

                            spsCount = 1;
                            while (i + 1 <= x.Length && x.Substring(i + spsCount, 1) == "　")
                            {
                                if (indentSpsCount > 0 && spsCount >= indentSpsCount) break;
                                spsCount++;
                            }
                            i = i + spsCount;
                            continue;

                        }
                        //    notTitleIndent = false;
                    }

                }
                else notTitleIndent = true;
                //x = textBox1.Text;//不能更新，因為用了 offset
            }

            CnText.ReplaceBlanksWithSpaces(textBox1);

            restoreCaretPosition(textBox1, s, 0);
            stopUndoRec = false;
            if (textBox1.TextLength > 1100) playSound(soundLike.over, true);

        }
        /// <summary>
        /// Alt + 1 : 鍵入本站制式留空空格標記「􏿽」：若有選取則取代全形空格「　」為「􏿽」
        /// 若被選取的是{{或}}則逕以「􏿽」取代（《國學大師》的《四庫全書》本常見
        /// </summary>
        private void keysSpacesBlank()
        {
            string x = textBox1.Text;
            int s = textBox1.SelectionStart;//, l = textBox1.SelectionLength;
            string sTxt = textBox1.SelectedText;
            dontHide = true;
            if (sTxt != "")
            {//有選取範圍
             //如果已選取「{{」或「}}」或「 」則逕以「􏿽」取代（《國學大師》的《四庫全書》本常見
                if ("{{}} ".IndexOf(sTxt) > -1)
                {
                    undoRecord();
                    stopUndoRec = true;
                    textBox1.SelectedText = "􏿽";
                    dontHide = false;
                    stopUndoRec = false;
                    return;
                }
                if (sTxt == "<p>")
                { undoRecord(); stopUndoRec = true; textBox1.SelectedText = "􏿽"; stopUndoRec = false; }
                else
                {
                    if (sTxt.IndexOf("　") == -1) { stopUndoRec = false; dontHide = false; return; }

                    undoRecord();
                    stopUndoRec = true;
                    //textBox1.SelectedText = sTxtChk;
                    textBox1.SelectedText = CnText.ReplaceFullWidthSpace_withBlank(sTxt);//進階版 string sTxtChk = sTxt.Replace("　", "􏿽");
                    dontHide = false;
                    stopUndoRec = false;
                    /*
                    textBox1.Text = x.Substring(0, s) + sTxtChk + x.Substring(s + sTxt.Length);
                    //string sTxtChk = sTxt.Replace("　", "");
                    //if (sTxtChk != "") return;
                    //for (int i = 0; i < sTxt.Length; i++)
                    //{
                    //    sTxtChk += "􏿽";
                    //}
                    //x = x.Substring(0, s) + sTxtChk + x.Substring(s + l);
                    //textBox1.Text = x;
                    if (s + l + countWordsinDomain("　", x.Substring(s)) == textBox1.TextLength)
                        textBox1.SelectionStart = s;
                    else
                        textBox1.SelectionStart = s + sTxtChk.Length;
                    */

                }
                //caretPositionRecall();
                //textBox1.ScrollToCaret();
            }
            else
            {//無選取範圍
                #region 將行/段尾在插入點附近的<p>置換成「􏿽」----配合 paragraphMarkAccordingFirstOne()使用
                if (s + 3 <= x.Length && x.Length > 3 && s >= 3)
                {
                    int sn = expandSelRange("<p>", x);
                    if (sn + 3 + 2 <= x.Length && x.Substring(sn + 3 + 2).Replace(Environment.NewLine, "") != "")
                    {
                        if (sn > 0 && x.Substring(sn + 3, 2) == Environment.NewLine)
                        {
                            textBox1.Select(sn, 3);
                            undoRecord(); stopUndoRec = true; textBox1.SelectedText = "􏿽"; stopUndoRec = false;
                            stopUndoRec = false; dontHide = false;
                            return;
                        }
                    }

                }
                #endregion
                //else
                //{
                //if (s + 1 <= x.Length && x.Substring(s, 1) == "　")
                if (s + 1 <= x.Length && (x.Substring(s, 1) == "　" || x.Substring(s, 1) == " "))
                    //x = x.Substring(0, s) + "􏿽" + x.Substring(s + 1);// 自動清除後面的「　」與「　」字元                
                    textBox1.Select(s, 1);
                else
                    //x = x.Substring(0, s) + "􏿽" + x.Substring(s);
                    textBox1.Select(s, 0);
                //if (textBox1.Text != x)
                //{
                undoRecord();
                stopUndoRec = true;
                //textBox1.SelectedText = "􏿽";
                //20240126 Bing大菩薩：輸入特殊字符
                textBox1.SelectedText = "\uDBFF\uDFFD";
                //textBox1.Text = x;
                //textBox1.SelectionStart = s + "􏿽".Length;
                stopUndoRec = false;
                //}
                //}

            }
            dontHide = false;
        }

        /// <summary>
        /// F6：標題降階（增加標題前之星號）
        /// </summary>
        private void keysAsteriskPreTitle()
        {
            if (textBox1.Text.IndexOf("*") == -1) return;
            string x = textBox1.SelectedText; int s = textBox1.SelectionStart, originalS = s;
            if (keyinTextMode && x != string.Empty && x.IndexOf("*") == -1) { textBox1.DeselectAll(); x = string.Empty; }
            //if (textBox1.SelectedText != "")                
            if (x != "")
                expandSelectedTextRangeToWholeLinePara(s, textBox1.SelectionLength, textBox1.Text);
            else
            {
                x = textBox1.Text;
                s = 0;
            }

            caretPositionRecord();
            undoRecord();
            stopUndoRec = true;

            int i = x.IndexOf("*"), j = 0;

            while (i > -1 && i <= x.Length)
            {
                textBox1.Select(i + s + j, 1);
                textBox1.SelectedText += "*";
                //x = textBox1.Text;
                //i = x.IndexOf("*", i + 1);
                if (x.IndexOf(Environment.NewLine, i) == -1) break;
                i = x.IndexOf("*", x.IndexOf(Environment.NewLine, i));
                //if (i > -1) i += j;
                j++;
            }
            //回到執行前之位置
            if (j > 0) textBox1.Select(originalS + j, 0);
            else caretPositionRecall();
            stopUndoRec = false;
        }

        /// <summary>
        /// 標題（篇名）判斷，在無前置空格時（無縮排時） 20230114
        /// 此法可與 Alt + Pause 參用
        /// </summary>
        void detectTitleYetWithoutPreSpace()
        {//Alt + t ：預測游標所在行是否為標題
            if (NormalLineParaLength == 0)
                NormalLineParaLength = wordsPerLinePara;
            if (NormalLineParaLength == 0)
            {
                MessageBox.Show("請先執行Word VBA 「轉成黑豆以作行字數長度判斷用」程序以取得正文每行正常長度再執行。感恩感恩　南無阿彌陀佛", "", MessageBoxButtons.OK, MessageBoxIcon.Stop); return;//碼詳：https://github.com/oscarsun72/TextForCtext/blob/f75b5da5a5e6eca69baaae0b98ed2d6c286a3aab/WordVBA/%E4%B8%AD%E5%9C%8B%E5%93%B2%E5%AD%B8%E6%9B%B8%E9%9B%BB%E5%AD%90%E5%8C%96%E8%A8%88%E5%8A%83.bas#L316
            }

            int s = textBox1.SelectionStart, p;
            string x = textBox1.Text, newLineTag = x.IndexOf("<p>") == -1 ? Environment.NewLine : "<p>" + Environment.NewLine;
            do
            {
                if (textBox1.Text.Substring(s).Length < 100)
                {
                    textBox1.Select(s, 1); break;
                }
                string currentLineTxt = GetLineText_WithoutPunctuation(textBox1.Text, s);
                int lenCurrentLine = new StringInfo(currentLineTxt).LengthInTextElements;//行長度
                p = textBox1.Text.IndexOf(Environment.NewLine, s);
                string nextLineTxt = GetLineText_WithoutPunctuation(textBox1.Text, p + newLineTag.Length);
                int nextLineTxtLength = new StringInfo(nextLineTxt).LengthInTextElements;
                if (currentLineTxt.IndexOf("*") > -1 || currentLineTxt == "")
                {
                    s = p + newLineTag.Length + 1;
                    continue;
                }
                if (NormalLineParaLength > lenCurrentLine)
                {
                    MessageBoxDefaultButton dbtn = MessageBoxDefaultButton.Button2;
                    //如果行長度相差太多（目前設為4）且下一行又是一般正文的行長度時，則很可能是標題，故預設按鈕為 Yes
                    if (//下一段（行）長等於正文行長度
                        nextLineTxtLength >= NormalLineParaLength)
                    {
                        //檢查標題關鍵字
                        var keywordPostion = chkTitleKeyWords(currentLineTxt);
                        if ((int)keywordPostion < 2)
                        {
                            if (currentLineTxt.IndexOf("{") > -1)
                            {
                                //在有{{}}且末綴<p>的情況下，keysTitleCode();會出錯，但清掉<p>尾綴即可
                                GetLineText(textBox1.Text, s, out int linStart, out int lineLength);
                                textBox1.Select(linStart, lineLength);
                                textBox1.SelectedText = textBox1.SelectedText.Replace("<p>", "");
                                textBox1.Select(linStart + 1, 0);
                            }
                            else
                                textBox1.Select(s + 1, 0);
                            SystemSounds.Beep.Play();
                            titleMarkCode();
                            s = p + newLineTag.Length + 1;
                            continue;
                        }
                        else if (keywordPostion == keyWordPos.no)
                            dbtn = MessageBoxDefaultButton.Button2;
                        if (dbtn == MessageBoxDefaultButton.Button2 &&
                            NormalLineParaLength - lenCurrentLine > 4 &&//尾綴沒有「|」（平抬）時
                            currentLineTxt.IndexOf("|") == -1)
                            dbtn = MessageBoxDefaultButton.Button1;

                        //末尾只要是"銘曰", "辭曰"之類的都略去
                        string[] items = { "銘曰", "辭曰" }; bool continu = false;
                        foreach (var item in items)
                        {
                            if (currentLineTxt.IndexOf(item) > -1)
                            {
                                string chkTextEndStr = currentLineTxt.Substring(currentLineTxt.LastIndexOf(item));
                                if (chkTextEndStr == item + (currentLineTxt.IndexOf("<p>") > -1 ? "<p>" : "") ||
                                    chkTextEndStr == item + (currentLineTxt.IndexOf("|") > -1 ? "|" : ""))//「|<p>」在newTextBox1()會取代掉，故可於此忽略不管，而「<p><p>」送出網頁後自動會清除多餘的<p>，也不必管
                                {
                                    //textBox1.Select(s + 1, 0);//此2行debug用而已
                                    //textBox1.ScrollToCaret();
                                    s = p + newLineTag.Length + 1;
                                    continu = true; break;
                                }
                            }
                        }
                        if (continu) continue;
                    }//以上下一段等於或大於正常長，以下則否
                    else
                    {
                        string[] items = { "{{佚}}", "{{并序}}" };
                        foreach (var item in items)
                        {
                            if (currentLineTxt.LastIndexOf(item) > -1)
                            {
                                string chkTextEndStr = currentLineTxt.Substring(currentLineTxt.LastIndexOf(item));
                                if (chkTextEndStr == item + (currentLineTxt.IndexOf("<p>") > -1 ? "<p>" : "") ||
                                    chkTextEndStr == item + (currentLineTxt.IndexOf("|") > -1 ? "|" : ""))//「|<p>」在newTextBox1()會取代掉，故可於此忽略不管，而「<p><p>」送出網頁後自動會清除多餘的<p>，也不必管
                                                                                                          //,"{{佚}}" 後面不會有內容的，所以不太可能長度等於正文正常長度，應該是接下一個標題
                                {
                                    //在有{{}}且末綴<p>的情況下，keysTitleCode();會出錯，但清掉<p>尾綴即可
                                    GetLineText(textBox1.Text, s, out int linStart, out int lineLength);
                                    textBox1.Select(linStart, lineLength);
                                    textBox1.SelectedText = textBox1.SelectedText.Replace("<p>", "");
                                    textBox1.Select(linStart + 1, 0);
                                    SystemSounds.Beep.Play();
                                    titleMarkCode();
                                    continue;
                                    //dbtn = MessageBoxDefaultButton.Button1;
                                }
                            }
                        }
                    }

                    textBox1.Select(s, 0); textBox1.ScrollToCaret();
                    if (dbtn == MessageBoxDefaultButton.Button2) SystemSounds.Asterisk.Play();
                    DialogResult response =
                    MessageBox.Show(currentLineTxt + Environment.NewLine + Environment.NewLine +
                    "這行是標題（篇名）嗎？", "篇名判斷式", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question,
                        dbtn, MessageBoxOptions.DefaultDesktopOnly);
                    switch (response)
                    {
                        case DialogResult.Yes:
                            textBox1.Select(s + 1, 0);
                            titleMarkCode();
                            break;
                        case DialogResult.No:
                            break;//繼續do while 迴圈
                        default://Cancel,停止操作
                                //訊息方塊就是一個表單，顯示時會讓此表單失去焦點。
                            if (!Active) Activate();
                            return;
                    }
                }
                s = p + newLineTag.Length + 1;
            } while (p > -1);
            playSound(soundLike.over);
            bringBackMousePosFrmCenter();
        }

        enum keyWordPos { pre, end, yes, no }
        //檢查標題關鍵字
        keyWordPos chkTitleKeyWords(string chkText)
        {
            #region 檢查標題關鍵字宣告
            string[] titleKeywordEnd = { "文","序", "又", "詩", "韻", "韵","歌","咏","詠","篇","章","聯句",
                "解","疏","章奏","讚","贊",
                "議","論","策","䇿","詔","詔文","旨","㫖","令","說","說二","說上","說下",
                "傳", "記","逸事","述","賦", "碑","𥓓", "銘","詺", "碣","表", "誌", "權厝志","書","書後","一","二","三","四","五",
                "序","敘","敍","叙","引","跋","䟦","箋","牋","略","狀","道","箴","頌","辭"
                ,"帖","事","實録","實錄","起居注","起居註","政紀","録","錄","注","註",
                "{{并序}}","{{代}}" };//,"{{佚}}" 後面不會有內容的，所以不太可能長度等於正文正常長度，應該是接下一個標題
                                   //,"墓表", "墓誌", "墓誌銘", "墓志銘"};//後綴
            string[] titleKeywordPre = { "上", "答", "又", "再", "荅", "復", "覆", "與", "題", "祭", "讀", "說"
                    , "釋", "記", "書" ,"辯","論","送","擬","最錄"}; //前綴
            string[] titleKeyword = { "", "" };
            //20230114 chatGPT菩薩：C# 2D Array Conversion:
            string[][] titleKeywords = { titleKeywordEnd, titleKeywordPre, titleKeyword };
            #endregion
            for (int keysMember = 0; keysMember < titleKeywords.Length; keysMember++)
            {
                string[] Items = titleKeywords[keysMember];
                foreach (string item in Items)
                {
                    if (item != "" && chkText.IndexOf(item) > -1)
                    {
                        switch (keysMember)
                        {
                            case 0://後綴
                                string chkTextEndStr = chkText.Substring(chkText.LastIndexOf(item));
                                if (chkTextEndStr == item + (chkText.IndexOf("<p>") > -1 ? "<p>" : "") ||
                                    chkTextEndStr == item + (chkText.IndexOf("|") > -1 ? "|" : ""))//「|<p>」在newTextBox1()會取代掉，故可於此忽略不管，而「<p><p>」送出網頁後自動會清除多餘的<p>，也不必管
                                    return (keyWordPos)keysMember;//keyWordPos.end;
                                break;
                            case 1://前綴
                                if (chkText.IndexOf(item) == 0)
                                    return (keyWordPos)keysMember;//keyWordPos.pre;
                                break;
                            case 2:
                                //前已有判斷了 if (…… chkText.IndexOf(item) > -1)
                                return (keyWordPos)keysMember;//keyWordPos.yes;
                            default:
                                break;
                        }
                    }
                }
            }
            return keyWordPos.no;
        }
        void autoMarkTitles()
        {
            undoRecord();
            stopUndoRec = true;
            //select the spaces front of the title
            string sps = textBox1.SelectedText;
            if (sps == "" && !TitleLeadingSpacesCount.Contains(0)) return;//也有沒空格的標題，如《經義考》等 https://ctext.org/library.pl?if=gb&file=76738&page=43
            if (sps.Replace("　", "") != "") return;

            PauseEvents();

            #region 設定各階標題的前置空格數欄位參考20250222
            titleLeadingSpaceCount = sps.Length;
            //if (!TitleLeadingSpacesCount.Contains(sps.Length))//一定都是從第1階的標題開始mark加上標題標記
            //    TitleLeadingSpacesCount.Add(sps.Length);
            if (!TitleLeadingSpacesCount.Contains(titleLeadingSpaceCount)) TitleLeadingSpacesCount.Add(titleLeadingSpaceCount);
            #endregion

            int s = textBox1.Text.IndexOf(Environment.NewLine), ss = textBox1.SelectionStart, sPre = 0;
            string x = textBox1.Text;
            if (x.Substring(0, sps.Length).Replace("　", "") == "")
            {
                textBox1.Select(0, 0);//從頭標題起
                stopUndoRec = true;
                titleMarkCode();
                s = textBox1.SelectionStart;
            }
            while (s > -1)//s=分段符號位置：s = textBox1.Text.IndexOf(Environment.NewLine)
            {
                s += 2;
                if (s + sps.Length >= textBox1.TextLength) break;
                if (textBox1.Text.Substring(s, sps.Length) == sps)
                {
                    textBox1.Select(s + sps.Length, 0);

                    x = textBox1.Text;
                    if (s + 2 <= x.Length && x.IndexOf(Environment.NewLine, s + 2) == -1)
                    {
                        stopUndoRec = true;
                        if (!titleMarkCode()) break;
                        //s = textBox1.SelectionStart; 
                        break;
                    }

                    string xp = x.Substring(s + 2, x.IndexOf(Environment.NewLine, s + 2) - (s + 2));
                    if ((!(xp.IndexOf("}}") > -1 && xp.IndexOf("{{") == -1) &&
                        textBox1.Text.Substring(sPre, sps.Length) != sps)
                        ||
                        (_leadingSpacesRegex.Match(GetLineText(textBox1.Text, s)).Value.Length == sps.Length
                            && _leadingSpacesRegex.Match(GetLineText(textBox1.Text, sPre)).Value.Length > sps.Length//●●●●●●●●●●●●●●●●20250308
                            && !((xp.IndexOf("}}") > -1 && xp.IndexOf("{{") == -1) ||//頂行不是縮排的注文、行首不能是注文縮排
                            (xp.IndexOf("}}") < xp.IndexOf("{{"))))
                        )
                    {
                        stopUndoRec = true;
                        if (sps.Length == 0 && !xp.StartsWith("{{　"))//如果無空格的是注文縮排則略過不做
                            textBox1.Select(s, x.IndexOf(Environment.NewLine, s) - s);//如果標題前端無空格，則僅作一行/段的標題●●●●●●●●●●●●
                        if (!titleMarkCode()) break;
                        s = textBox1.SelectionStart;

                    }
                }
                if (s + 1 >= textBox1.TextLength || textBox1.SelectionStart + 1 >= textBox1.TextLength) break;
                sPre = s;
                s = textBox1.Text.IndexOf(Environment.NewLine, s++);
            }
            //textBox1.Text = textBox1.Text.Replace("<p>" + Environment.NewLine + sps + "*", Environment.NewLine + sps);
            stopUndoRec = false; undoRecord();
            titleLeadingSpaceCount = 0;//歸零
            textBox1.Select(ss, 0); textBox1.ScrollToCaret();
            ResumeEvents();
        }
        /// <summary>
        /// 篇名前的全形空格字串，預設為0個全形空格(如《人境廬詩草》即是）
        /// </summary>
        string spaceStrBreforeTitle = "";
        /// <summary>
        /// 篇名標題標注
        /// 加上篇名格式代碼
        /// </summary>
        /// <returns>出錯或失敗則傳回false。不標或已標而不再重標亦傳回true</returns>
        private bool titleMarkCode()
        {
            //移至呼叫端
            //if (textBox1.SelectionLength==0 && textBox1.Text  wordsPerLinePara == -1)
            //{
            //    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請先用Scroll Lock鍵讓程式取得每行/段的正常長度（字數）以供判斷。"); return false;
            //}

            int s = textBox1.SelectionStart, i = s;
            bool groupTitle = false;
            string x = textBox1.Text, asterisks = "*", asteriskAppend = string.Empty;
            for (int t = 0; t < TitleLeadingSpacesCount.Count; t++)
            {
                if (TitleLeadingSpacesCount[t] == titleLeadingSpaceCount)
                {
                    asteriskAppend = new string('*', t);
                    break;
                }

            }

            asterisks += asteriskAppend;
            //下行僅debug時用，因為還有2星以上的階層標題
            //if (getLineTxt(x,s).IndexOf("*") > -1) return;
            if (!stopUndoRec)
            {
                undoRecord();
                stopUndoRec = true;
            }
            if (textBox1.SelectedText != string.Empty)
            {//目前好像用不到選取指定標題，故暫去掉，以便配合 按 F3鍵找標題處加標題格式
                if (textBox1.SelectedText.Replace("　", "") == "")
                {
                    textBox1.DeselectAll();
                }
            }
            string context = textBox1.Text;
            Document document = new Document(context);
            if (textBox1.SelectedText == string.Empty)//目前好像用不到選取指定標題，故暫去掉
            {
                if (x.Length >= s + 1 && x.Substring(s, 1) == "　")
                {
                    int l = x.Length;
                    while (x.Substring(i++, 1) == "　")
                    {
                        if (i == l) break;
                    }
                    s = i;
                }

                if (!(s > 1 && x.Substring(s - 2, 2) == Environment.NewLine))
                {
                    //看目前插入點位置前是否是空格，若是，則停駐不再往前移位
                    //string titieBeginChar = x.Substring(i == 0 ? i : --i, 1);//若寫成「i--」，則在 i==x.Length時會出現錯誤，因為為--i是先減再用，而i--則是先用再減，先用，則第2個引數就會超出x的長度 20230930
                    string titieBeginChar = x.Substring(i == 0 ? i : i - 1, 1);//「寫成」--i，會動到i值！ 20250204 
                    while (titieBeginChar != "　" &&
                        titieBeginChar != Environment.NewLine.Substring(Environment.NewLine.Length - 1, 1))
                    {
                        if (i == 0) break;
                        titieBeginChar = x.Substring(i == 0 ? i : --i, 1);
                    }
                    if (i > 0 && i < s) i++;//●●●●●●●●●●●●●●●●
                    s = i;//●●●●●●●●●●●●OK 20250204
                          //if (i != 0)
                          //    s = i;//+ 2;//20240520 待觀察/////////////////////////// 全自動才用此（待測試）
                          //          //s = i + 1;//20240521 自動標題會被影響！
                          //else s = i;
                }

                //以上s為標題內容起始位置，即空格末後位置

                string titleFirstParaText;// = pTitleFirstPara.Text;//標題首行/段的string
                titleFirstParaText = GetLineText(x, s);
                if (titleFirstParaText == string.Empty) return true;

                //若前置空格數不相符則不作標記（若原始文本排版錯誤則會漏標，只能人工校正了）20250223
                if (titleLeadingSpaceCount > 0 && _leadingSpacesRegex.Match(titleFirstParaText).Value.Length != titleLeadingSpaceCount) return true;

                //if (titleFirstParaText.IndexOf("中蝦墓隨其生") > -1) Debugger.Break();
                ////小注縮排起頭者亦不加標
                //if (titleLeadingSpaceCount == 0 && titleFirstParaText.StartsWith("{{　")) return true;


                //if (titleFirstParaText.Contains("歌")) Debugger.Break();

                #region 星號檢查
                if (!keyinTextMode && s > 0 && titleFirstParaText.IndexOf(asterisks) > -1
                    && s > x.IndexOf(Environment.NewLine)) return true;//如果已有星號標記則斥回（第1行/段例外）●●●●●●●●●●●●●應該是區分手動與自動時，俟考！
                                                                       //檢查前一行/段的內容

                int preLineEnd = x.LastIndexOf(Environment.NewLine, s);
                string preLineText, nextLineText;
                if (preLineEnd > -1)
                {
                    preLineText = GetLineText(x, preLineEnd);
                    if ((preLineText.IndexOf("*") > -1 && CountWordsLenPerLinePara(preLineText) >= wordsPerLinePara)
                        //如果前一段已有星號，且其長度又與正常長度相同或更長（則此行當是標題太長而折行）則駁回
                        ||//或者是第一行小於正常長長而本行又同於正常行長，則本行當屬縮排的序文等項
                        (preLineText.IndexOf("*") > -1 && CountWordsLenPerLinePara(preLineText) < wordsPerLinePara
                            && CountWordsLenPerLinePara(titleFirstParaText) == wordsPerLinePara)
                        ||//或者不是縮排中的文字
                        (preLineText.IndexOf("*") == -1 && titleLeadingSpaceCount > 0 && _leadingSpacesRegex.Match(preLineText).Value.Length == titleLeadingSpaceCount
                            && CountWordsLenPerLinePara(preLineText) >= CountWordsLenPerLinePara(titleFirstParaText))
                        )
                        return true;

                }
                #endregion

                int nextLineStart = x.IndexOf(Environment.NewLine, s);
                if (!keyinTextMode && nextLineStart == -1) return true;//沒有下一段當然不會是標題(手動輸入時通常是一頁頁，不是一卷卷故）
                nextLineStart += Environment.NewLine.Length;
                nextLineText = GetLineText(x, nextLineStart);
                #region 空一格之行/段檢查
                if (_leadingSpacesRegex.Match(titleFirstParaText).Value.Length == 1)
                {//若是空一格之行/段，且其長不短於正常行長，而其後不是縮排者，則駁回：
                    if (CountWordsLenPerLinePara(titleFirstParaText) >= lines_perPage)
                        if (nextLineText != string.Empty
                            && nextLineText.StartsWith("{{") == false
                            && _leadingSpacesRegex.Match(nextLineText).Value.Length == 0
                            && s > x.IndexOf(Environment.NewLine))//第1行例外●●●●●●●●●●●●
                            return true;
                }
                #endregion


                #region 取得標題範圍（多行標題下的範圍）

                //標題首行/段－－此時尚未標上星號（*）（是在第一階標題才未，若2階以後則標上了）　20250204 20250221增修
                //Paragraph pTitleFirstPara = _document.Range(s, s).GetCurrentParagraph();
                //Paragraph pNextTitleFirstPara = pTitleFirstPara.Range.GetNextParagraph();//標題首行/段的下一行/段
                int pTitleFirstParaEnd;//= pTitleFirstPara.End;
                string pNextTitleFirstParaText;// = pNextTitleFirstPara.Text;
                titleFirstParaText = GetLineText(x, s, out int lineS, out int lineL);
                pTitleFirstParaEnd = lineS + lineL;
                if (pTitleFirstParaEnd + Environment.NewLine.Length > x.Length)
                    pNextTitleFirstParaText = string.Empty;
                else
                    pNextTitleFirstParaText = GetLineText(x, pTitleFirstParaEnd + Environment.NewLine.Length);
                context = textBox1.Text;
                document = new Document(context);
                //借用x變數，取得插入點後的文字
                x = x.Substring(s);
                for (int j = 0; j + 2 <= x.Length; j++)//j為偏移的個數
                {
                    string nx = x.Substring(j, 2);

                    //if (GetLineTxt(textBox1.Text, s).Contains("送大理封主簿五郎親事不合却赴通州主簿")) Debugger.Break();

                    if (nx == Environment.NewLine || nx == "{{" || nx == "<p")
                    {

                    longTitle:
                        if (nx == Environment.NewLine)
                        {

                            //if (titleFirstParaText.Contains("送大理")) Debugger.Break();//just for test
                            //string titleFistLineStr = pFirst.Text;//取得標題首行字串
                            //標題首行長度若小於常規就不再繼續找下去

                            //singleLineTitle:
                            if (CountWordsLenPerLinePara(titleFirstParaText) < wordsPerLinePara)
                            {
                                textBox1.Select(s, pTitleFirstParaEnd - s);//選取標題文字內容,準備將標題格式，置換成標題語法格式


                                //if (textBox1.SelectedText.Contains("涼雨一章")) Debugger.Break();//just for test


                                //下一行/段也是短於正常行長者，則當屬組詩之類的標題了 20250221
                                //且其前縮排之空格當較正題為多 20250222
                                if (CountWordsLenPerLinePara(pNextTitleFirstParaText) < wordsPerLinePara &&
                                    _leadingSpacesRegex.Match(titleFirstParaText).Value.Length <= _leadingSpacesRegex.Match(pNextTitleFirstParaText).Value.Length)
                                {

                                    if (!groupTitle)
                                    {
                                        playSound(soundLike.warn, true);
                                        groupTitle = true;
                                    }
                                }
                                break;//單行/段標題，則後面的過長時的狀況自然就不再追查了 20250222
                            }

                            groupTitle = false;


                            //標題（篇名）過長時之處理：
                            if (j + 2 + 1 <= x.Length)
                            {
                                if (x.Substring(j + 2, 1) == "　")
                                {//下一行為縮排空格起者

                                    if (x.IndexOf(Environment.NewLine, j + 2) > -1
                                        && CountWordsLenPerLinePara(x.Substring(j + 2, x.IndexOf(Environment.NewLine, j + 2) - (j + 2))) == wordsPerLinePara)
                                    {//這一行的字數和正常行/段長同者（表示標題還未結束）
                                        continue;
                                    }
                                    ////●●●●●●●●●●●●●●●●●●
                                    //string titleParaText = pFirst.Text;
                                    //標題首行前面空格數（縮排空格數）
                                    Match matchTitleLeadingSpace = _leadingSpacesRegex.Match(titleFirstParaText);
                                    //現在處理的位置 s + j + 2 + 1 前面空格數（縮排空格數,j 只是偏移、位移 offset的單位長）
                                    //Match matchSelectionLeadingSpace = _leadingSpacesRegex.Match(new Range(ref new Document(ref textBox1), s + j + 2 + 1, s + j + 2 + 1).GetCurrentParagraph().Text);
                                    Match matchSelectionLeadingSpace = _leadingSpacesRegex.Match(document.Range(s + j + 2 + 1, s + j + 2 + 1).GetCurrentParagraph().Text);
                                    if ((matchSelectionLeadingSpace.Success ? matchSelectionLeadingSpace.Value.Length : 0) >=//如果是標題本身換行之縮排，通常
                                    (matchTitleLeadingSpace.Success ? matchTitleLeadingSpace.Value.Length : 0))
                                        continue;
                                }
                                else if (j + 2 + 3 < x.Length && x.Substring(j + 2, 3) == "{{　")
                                {
                                    if (x.IndexOf(Environment.NewLine, j + 2) > -1)
                                    {
                                        string currentLineText = x.Substring(j + 2, x.IndexOf(Environment.NewLine, j + 2) - (j + 2));
                                        if (CountWordsLenPerLinePara(x.Substring(j + 2, x.IndexOf(Environment.NewLine, j + 2) - (j + 2))) == wordsPerLinePara)
                                        {//這一行的字數和正常行/段長同者，且起頭是注文標記又有縮排（表示標題還未結束）
                                            Match matchTitleLeadingSpace = _leadingSpacesRegex.Match(titleFirstParaText);
                                            //現在處理的位置 s + j + 2 + 1 前面空格數（縮排空格數,j 只是偏移、位移 offset的單位長）
                                            //Match matchSelectionLeadingSpace = _leadingSpacesRegex.Match(new Range(ref new Document(ref textBox1), s + j + 2 + 1, s + j + 2 + 1).GetCurrentParagraph().Text);
                                            //Match matchSelectionLeadingSpace = _leadingSpacesRegex.Match(document.Range(s + j + 2 + 1, s + j + 2 + 1).GetCurrentParagraph().Text);
                                            Match matchSelectionLeadingSpace = _leadingSpacesRegex.Match(currentLineText.Substring("{{".Length));
                                            if ((matchSelectionLeadingSpace.Success ? matchSelectionLeadingSpace.Value.Length : 0) >=//如果是標題本身換行之縮排，通常
                                            (matchTitleLeadingSpace.Success ? matchTitleLeadingSpace.Value.Length : 0))
                                                continue;
                                            continue;
                                        }
                                    }
                                }

                            }
                        }
                        //如果篇名標題有小注，則在其結尾處加上分段符號<p>
                        if (nx == "{{")
                        {
                            if (x.IndexOf(Environment.NewLine, j) > -1 &&
                                CountWordsLenPerLinePara(x.Substring(j, x.IndexOf(Environment.NewLine, j) - j)) == wordsPerLinePara)
                                continue;//這一行的字數和正常行/段長同者（表示標題還未結束）

                            //Paragraph pNext = _document.Range(s, s).GetNextParagraph();
                            //string pNextTitleFirstParaText = pNext.Text;
                            //if (pNext.Text.IndexOf("忽止曰嘶") > -1) Debugger.Break();
                            //如果下一行/段的起始為空格、縮排，則須再往後找。
                            else if (pNextTitleFirstParaText.StartsWith("　"))
                            {
                                if (pNextTitleFirstParaText.Replace("　", string.Empty).StartsWith("{{"))
                                    continue;
                                else
                                {
                                    //●●●●●●●●●●●●●標題行/段－－此時尚未標上星號（*）　20250204
                                    //pFirst = _document.Range(s, s).GetCurrentParagraph();

                                    //string titleParaText = pFirst.Text;


                                    //標題首行前面空格數（縮排空格數）
                                    Match matchTitleLeadingSpace = _leadingSpacesRegex.Match(titleFirstParaText);
                                    //現在處理的位置 s + j + 2 + 1 前面空格數（縮排空格數,j 只是偏移、位移 offset的單位長）
                                    //Match matchSelectionLeadingSpace = _leadingSpacesRegex.Match(new Range(ref new Document(ref textBox1), s + j + 2 + 1, s + j + 2 + 1).GetCurrentParagraph().Text);
                                    Match matchSelectionLeadingSpace = _leadingSpacesRegex.Match(document.Range(s + j + 2 + 1, s + j + 2 + 1).GetCurrentParagraph().Text);
                                    if ((matchSelectionLeadingSpace.Success ? matchSelectionLeadingSpace.Value.Length : 0) >=//如果是標題本身換行之縮排，通常
                                    (matchTitleLeadingSpace.Success ? matchTitleLeadingSpace.Value.Length : 0))
                                        continue;
                                }
                            }

                            #region 標題中有小注 bugs still
                            int sCloseCurlyBrackets = x.IndexOf("}}", j), sNewLine = x.IndexOf(Environment.NewLine, j);
                            if (sCloseCurlyBrackets > -1)
                            {
                                if (sCloseCurlyBrackets < sNewLine - "}}".Length &&
                                    x.Substring(sCloseCurlyBrackets + 2, 3) != "<p>")
                                {
                                    //nx = Environment.NewLine;
                                    j = sNewLine; nx = x.Substring(j, 2);
                                    goto longTitle;
                                }

                                if (sCloseCurlyBrackets < sNewLine - "}}".Length &&
                                    x.Substring(sCloseCurlyBrackets + 2, 3) == "<p>")
                                {
                                    textBox1.Select(s, sCloseCurlyBrackets + 2 + 3);
                                    break;
                                }
                            }
                            #endregion
                            if (j + 2 + 1 <= x.Length)
                            {
                                int k = x.IndexOf(Environment.NewLine, j + 2 + 1);
                                //if (k > -1 || k == x.Length - 2)
                                while (k + 1 <= x.Length || k > -1 || k == x.Length - 2)
                                {
                                    if (k - 2 < 0)
                                    {

                                        Debugger.Break();//●●●●●●●●●●●●●●●●●●
                                        //stopUndoRec = false; return true;
                                        break;
                                    }
                                    if (x.Substring(k - 2, 2) == "}}" && x.Substring(k + 2, 2) != "{{")
                                    {
                                        textBox1.Select(s + k, 0); textBox1.SelectedText = "<p>"; break;
                                    }
                                    k = x.IndexOf(Environment.NewLine, k + 1);
                                }
                            }
                        }
                        textBox1.Select(s, j);//選取標題文字內容,準備將標題格式，置換成標題語法格式
                                              //textBox1.Select(s - 1 > 0 ? s - 1 : 0, j + 1);//20250204●●●●●●●●●●●●
                                              //textBox1.Select(s - 1 > 0 ? s -- : 0, j );//20250204●●●●●●●●●●●●
                                              //while(textBox1.SelectedText.StartsWith("　"))
                                              //    textBox1.Select(s - 1 > 0 ? s - 1 : 0, j + 1);//20250204●●●●●●●●●●●●


                        break;
                    }

                    //if (j > 15) Debugger.Break();//just for test ●●●●●●●●●●●●●● 20250222
                }
                #endregion
            }

            #region 20221019補訂，為第一行為標題者
            if (s == 0)
            {
                int sl = textBox1.SelectionLength;
                while (textBox1.SelectedText != "" && textBox1.SelectedText.Substring(0, 1) == "　")
                {
                    textBox1.Select(++s, --sl);
                }
            }
            #endregion

            #region 20230723補訂，為最後一行為標題者
            if (textBox1.TextLength == s + x.Length
                && x.IndexOf(Environment.NewLine) == -1)
                textBox1.Select(s, textBox1.TextLength);
            #endregion

            x = textBox1.Text; string endCode = "<p>";
            if (s + textBox1.SelectionLength - 3 < 0)
            {
                if (textBox1.SelectedText != "" && textBox1.SelectionStart == 0)
                {
                    if (textBox1.SelectionStart + textBox1.SelectionLength + "<p>".Length <= textBox1.TextLength
                        && textBox1.Text.Substring(textBox1.SelectionStart + textBox1.SelectionLength, "<p>".Length) == "<p>")
                    {
                        //textBox1.SelectedText = "*" + textBox1.SelectedText;
                        textBox1.SelectedText = asterisks + textBox1.SelectedText;
                    }
                    else
                        //textBox1.SelectedText = "*" + textBox1.SelectedText + "<p>";
                        textBox1.SelectedText = asterisks + textBox1.SelectedText + "<p>";
                }
                stopUndoRec = false; return true;
            }
            if (s + textBox1.SelectionLength + 3 <= x.Length
                && (x.Substring(s + textBox1.SelectionLength - 3, 3) == "<p>" ||
                x.Substring(s + textBox1.SelectionLength, 3) == "<p>")) endCode = "";

            #region 設定標題格式（完成標題語法設置）
            if (IsNoChineseContent(textBox1.SelectedText))
            {
                return true;
            }
            //設定標題格式（完成標題語法設置）
            //string title = ("*" + textBox1.SelectedText + endCode)
            string title = (asterisks + textBox1.SelectedText + endCode)
                    .Replace("《", "").Replace("》", "").Replace("〈", "").Replace("〉", "").Replace("·", "")
                    .Replace("|<p>", "<p>").Replace("。<p>", "<p>");

            //if (title.IndexOf("人乃傳秋露何也謝少谿侍郎者佳") > -1) Debugger.Break();//just for debug


            int st_title = textBox1.SelectionStart;//, ed_title =st_title+title.Length;
            PauseEvents();
            textBox1.SelectedText = title;//完成初步的標題語法設置            
            ResumeEvents();
            //設定值完成後，textBox1.SelectionLength 值會歸0
            //故要還原選取區
            textBox1.Select(st_title, title.Length);//標題格式此時已加上了： textBox1.SelectedText= *標題內容<p>
                                                    //context = textBox1.SelectedText;
                                                    //document = new Document(ref context);
                                                    //進一步作檢查其他的標題格式：
            int linesCount_title = linesCounter(title);

            //標題只有一行/段者：
            if (linesCount_title == 1)
            {
                #region 先處理特殊的

                if (title.IndexOf("　}}　") > -1)
                {
                    //creedit_with_Copilot大菩薩：20250212元宵節
                    //string title = "　　　*念奴嬌{{錄一　}}　　　　　{{宋}}蘇　軾<p>";

                    // 匹配 * 和 <p> 之间的内容
                    string pattern = @"(?<=\*)(.*?)(?=<p>)";

                    // 使用正则替换
                    string result = Regex.Replace(title, pattern, m =>
                    {
                        // 在匹配的内容中，替换不在 {{}} 之间的全形空格
                        string replaced = Regex.Replace(m.Value, @"{{.*?}}", match =>
                        {
                            // 对于 {{}} 内的内容，不进行替换
                            return match.Value.Replace("　", "\0"); // 暂时替换为不可见字符
                        });

                        // 替换全形空格为 '􏿽'
                        replaced = replaced.Replace("　", "􏿽");

                        // 恢复 {{}} 内的全形空格
                        replaced = replaced.Replace("\0", "　");

                        return replaced;

                    });
                    title = result;
                    PauseEvents();
                    textBox1.SelectedText = title;
                    ResumeEvents();
                    //設定值完成後，textBox1.SelectionLength 值會歸0
                    //故要還原選取區
                    textBox1.Select(st_title, title.Length);
                }
                #endregion

                #region 單行標題中的全形空格格式化、標準化
                //   title = title.Replace("　", "􏿽");//原式，不加判斷，一律取代為空白
                //20250204大年初六立春~初七子夜 感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主
                int spsStart = title.IndexOf("　");
                if (spsStart > 1)
                {
                    if (!title.Contains("{{"))//如果標題不含注文
                        //title = title.Replace("　", "􏿽");
                        title = CnText.ReplaceFullWidthSpace_withBlank(title);
                    else
                    {//如果標題中有注文
                        title = title.Replace("}}　{{", "}}􏿽{{");
                        if (title.Substring(spsStart + 1, 1) == "　" || spsStart < title.Length - (3 + 3)) //至少有3個中文（+<p>.Length）在末後才轉此前之空格為空白
                        {
                            int isps = 0; StringBuilder sb = new StringBuilder();
                            while (title.Substring(spsStart + isps, 1) == "　")
                            {
                                isps++;
                                sb.Append("􏿽");
                            }
                            if (title.Substring(spsStart + isps, 1) == "}") sb.Replace("􏿽", "　");//若是小注不換行的補空格則不能轉為空白
                            title = title.Substring(0, spsStart) + sb.ToString() + title.Substring(spsStart + isps);
                        }
                    }
                }
                if (textBox1.SelectedText != title) textBox1.SelectedText = title;
                #endregion 單行標題中的全形空格格式化、標準化
            }
            else
            {//兩行以上的標題格式處理：20250204

                //if (title.Contains("孫守真")) Debugger.Break();
                //if (title.Contains("送大理封主簿五郎親事不合却赴通州主簿")) Debugger.Break();

                //for (int i_title = 0; i < linesCount_title-1; i++)
                //{

                //}

                //Range range = new Range(_document, textBox1.SelectionStart, textBox1.SelectionStart + textBox1.SelectionLength);
                //Range range = new Range(ref _document, st_title, st_title + title.Length);
                Range range = _document.Range(st_title, st_title + title.Length);
                //Range range = document.Range(st_title, st_title + title.Length);
                //Range range = document.Content;
                List<Paragraph> paragraphs = range.Paragraphs;
                int st_range = range.Start;//, ed_range = range.End;
                int leadspaceCount_firstLine = 0; //記下首行/段長
                int spsCountIndent = 0;
                for (int iPara = 0; iPara < paragraphs.Count; iPara++)
                {
                    Paragraph item = range.Paragraphs[iPara];
                    int spsStart = item.Text.IndexOf("　");//空格起始處
                    switch (spsStart)
                    {
                        case -1:

                            break;
                        case 0://一開始為空格之縮排
                            if (spsStart + 2 < item.Text.Length)
                            {//計算縮排空格之數量
                                while (item.Text.Substring(++spsStart, 1) == "　")
                                {
                                    if (spsStart + 2 > item.Text.Length) break;
                                }
                            }
                            spsCountIndent = spsStart;
                            spsStart = item.Text.IndexOf("　", spsStart);
                            switch (iPara)
                            {
                                case 0://標題內的第一行/段
                                    leadspaceCount_firstLine = spsCountIndent;
                                    break;
                                default://標題內的第2行/段以後
                                    if (spsStart > -1)
                                    {
                                        int noteMarkClose = item.Text.IndexOf("}}", spsStart);
                                        if (noteMarkClose > -1 && item.Text.IndexOf("{{") == -1)//如果空格在夾注內只是縮排用的空格
                                        {
                                            //只處理夾注後正文前的空格
                                            if (noteMarkClose < item.Text.IndexOf("　", noteMarkClose))
                                            {
                                                noteMarkClose = item.Text.IndexOf("　", noteMarkClose);
                                                if (noteMarkClose > -1)
                                                    spsStart = noteMarkClose;
                                            }

                                            //如果標題第2行起始為空格、縮排，則須置換其文中的空格為空白，唯作者間的的空格不置換
                                            StringBuilder sb = new StringBuilder();
                                            int sbRootCount = 0;
                                            while (item.Text.Substring(spsStart + sbRootCount, 1) == "　")
                                            {//一直到非空格處即止 20250205
                                                sbRootCount++;
                                                sb.Append("􏿽");
                                            }
                                            Debugger.Break();
                                            //item.Range.Text = item.Range.Text.Substring(0, spsStart) + sb.ToString() + item.Range.Text.Substring(spsStart + sbRootCount);
                                            item.Range.Text = item.Text.Substring(0, spsStart) + sb.ToString() + item.Text.Substring(spsStart + sbRootCount);
                                            range.End += sbRootCount;
                                        }
                                        else if (noteMarkClose > -1 && item.Text.IndexOf("{{") > spsStart)//如果空格在夾注前
                                        {/* 要轉為空白的空格在夾注前如：
                                          　　　*余舊逰魏郡日有雙魚之饌今絶不可得問之
 　　　　                                   　　　　則河徙故也戲成一章　　　{{明}}王世貞<p> http://skqs.guoxuedashi.net/wen_2885i/175137.html
                                          */
                                            StringBuilder sb = new StringBuilder(); int ispscount = 0;
                                            for (int isps = spsStart; isps < item.Text.IndexOf("{{"); isps++)
                                            {
                                                sb.Append("􏿽");
                                                ispscount++;
                                            }
                                            Debugger.Break();
                                            item.Range.Text = item.Text.Substring(0, spsStart) + sb.ToString() + item.Text.Substring(item.Text.IndexOf("{{"));
                                            //range.End += ispscount;//已與Paragraph.Range.Text同步

                                        }
                                    }

                                    else if (spsCountIndent > 0)//●●●●●●●●●●●● 20250207
                                    {//可能是末尾單純的作者格式（第1行是篇名的開頭），則將其前的空格轉成空白
                                        Match match = _leadingSpacesRegex.Match(item.Text);
                                        //如果此行/段的縮排空格較前一行/段的縮排空格較首行/段長不過2個字元，則可能只是標題的第2行
                                        if ((match.Success ? match.Value.Length : 0) - leadspaceCount_firstLine > 2)//通常第2行後縮排不會縮超過2個字的
                                        {
                                            int isps = 0; StringBuilder sb = new StringBuilder();
                                            while (item.Text.Substring(0 + isps, 1) == "　")
                                            {
                                                isps++;
                                                sb.Append("􏿽");
                                            }
                                            if (sb.ToString() != string.Empty)
                                            {
                                                //Debugger.Break();
                                                playSound(soundLike.exam, true);
                                                item.Range.Text = sb.ToString() + item.Text.Substring(isps);
                                                //range.End = ed_range + isps;//"􏿽".Length=2 - "　".Length=1 ;
                                                //range.End += isps;//"􏿽".Length=2 - "　".Length=1 ;//已與Paragraph.Range.Text同步
                                            }
                                        }

                                        //StringBuilder sb = new StringBuilder();
                                        //for (int isps = 0; isps < spsCount; isps++)
                                        //{
                                        //    sb.Append("􏿽");
                                        //}
                                        //item.Text = sb.ToString() + item.Text.Substring(spsCount);
                                        ////range.End = ed_range + spsCount;//"􏿽".Length=2 - "　".Length=1 ;
                                        //range.End += spsCount;//"􏿽".Length=2 - "　".Length=1 ;
                                    }
                                    break;

                            }
                            goto default;
                        //break;
                        default:
                            if (spsStart > 1)
                            {
                                Match match = _leadingSpacesRegex.Match(item.Text);
                                //如果此行/段的縮排空格較前一行/段的縮排空格較首行/段長不過2個字元，則可能只是標題的第2行
                                if ((match.Success ? match.Value.Length : 0) - leadspaceCount_firstLine < 3)
                                {
                                    #region 將非縮排亦非在注文內的空格（　）轉成空白（􏿽）creedit_with_Copilot大菩薩：20250212元宵節
                                    string text = item.Text.Substring(spsCountIndent);
                                    // 匹配 * 和 <p> 之间的内容
                                    //string pattern = @"(?<=\*)(.*?)(?=<p>)";
                                    //string pattern = @"(?<=　)(.*?)";
                                    //// 使用正则替换
                                    //string result = Regex.Replace(text, pattern, m =>
                                    //{
                                    string result1 = Regex.Replace(text, @"(?<=\{\{)(.*?)(?=}})", m1 =>
                                    {
                                        string content = m1.Value;
                                        if (content.Length > 0)
                                        {
                                            // 保留最后一个全形空格
                                            int lastIndex = content.Length - 1;
                                            if (content[lastIndex] == '　')
                                            {
                                                string replaced = content.Substring(0, lastIndex).Replace("　", "􏿽") + '　';
                                                return replaced;
                                            }
                                            else
                                            {
                                                return content.Replace("　", "􏿽");
                                            }
                                        }
                                        return content;
                                    });

                                    //    return result1;
                                    //});

                                    if (result1 != text)
                                    {
                                        StringBuilder sbb = new StringBuilder();
                                        for (int ispace = 0; ispace < spsCountIndent; ispace++)
                                        {
                                            sbb.Append("　");
                                        }
                                        //Debugger.Break();
                                        playSound(soundLike.exam, true);
                                        item.Range.Text = sbb.ToString() + result1;
                                        //range.End += (result1.Length - text.Length);//已於Range.Text中調用UpdateParentRange()方法解決同步更新連動的問題了！
                                        //20250306監測成功！●●●●●●●●●●●●●●
                                    }
                                    else
                                    {//creedit_with_Copilot大菩薩：
                                        /*
                                         * 　　　*自晨至午起居飲食皆以牆外人物之聲為節
            　　　　                                            戲書四絶{{錄一　}}􏿽􏿽􏿽􏿽􏿽􏿽范成大<p> http://skqs.guoxuedashi.net/wen_2885i/175135.html
                                         */

                                        string pattern = @"(?<=\}\})([　]+)(?=[^\s　])";
                                        string result = Regex.Replace(text, pattern, m => m.Value.Replace("　", "􏿽"));
                                        if (result != text)
                                        {
                                            StringBuilder ssb = new StringBuilder();
                                            for (int iSps = 0; iSps < spsCountIndent; iSps++)
                                            {
                                                ssb.Append("　");
                                            }
                                            Debugger.Break();
                                            item.Range.Text = ssb.ToString() + result;
                                            range.End += (result.Length - text.Length);
                                        }
                                    }

                                    //string result = Regex.Replace(text, pattern, m =>
                                    //{
                                    //    // 在匹配的内容中，替换不在 {{}} 之间的全形空格
                                    //    //string replaced = Regex.Replace(m.Value, @"{{.*?}}", matchNote =>//其實應該是小注不換行的空格才不轉為空白，不能是在小注夾注中的空格都不轉為空白，所以要將「 @"{{.*?}}"」改為「@"\S.*?}}"」
                                    //    string replaced = Regex.Replace(m.Value, @"\S.*?}}", matchNote =>
                                    //    {
                                    //        // 对于 {{}} 内的内容，不进行替换
                                    //        return matchNote.Value.Replace("　", "\0"); // 暂时替换为不可见字符
                                    //    });

                                    //    // 替换全形空格为 '􏿽'
                                    //    replaced = replaced.Replace("　", "􏿽");

                                    //    // 恢复 {{}} 内的全形空格
                                    //    replaced = replaced.Replace("\0", "　");

                                    //    return replaced;
                                    //});

                                    //if (result != text)
                                    //{
                                    //    item.Text = result;
                                    //    range.End += (result.Length - text.Length);
                                    //}
                                    #region 以下為原式
                                    /*如果全是正文無夾注者，如
                                     * 　　　*西湖秋涸東池魚窘甚因㑹客呼網師遷之西
            　　　　                                 　　　　池為一笑之樂夜歸被酒不能寐戲作放魚
            　　　　                                 　　　　一首　　　　　　　　　　蘇　軾<p> http://skqs.guoxuedashi.net/wen_2885i/175137.html
                                     * */
                                    if (item.Text.IndexOf("{{") == -1 && item.Text.IndexOf("}}") == -1)
                                    {
                                        int isps = 0; StringBuilder sb = new StringBuilder();
                                        while (spsStart + isps + 1 <= item.Text.Length && item.Text.Substring(spsStart + isps, 1) == "　")
                                        {
                                            isps++;
                                            sb.Append("􏿽");
                                        }
                                        if (sb.ToString() != string.Empty)
                                        {
                                            //Debugger.Break();
                                            item.Range.Text = item.Text.Substring(0, spsStart) + sb.ToString() + item.Text.Substring(spsStart + isps);
                                            ////range.End = ed_range + isps;//"􏿽".Length=2 - "　".Length=1 ;
                                            //range.End += isps;//"􏿽".Length=2 - "　".Length=1 ;//已於Range.Text中以UpdateParentRange()更新 20250306●●●●●●●●●
                                        }
                                    }
                                    #endregion
                                    #endregion
                                }
                                else
                                {//則第2行可能是單純作者的格式（第1行是篇名），則將其前的空格轉成空白
                                    StringBuilder sb = new StringBuilder();
                                    for (int isps = 0; isps < spsCountIndent; isps++)
                                    {
                                        sb.Append("􏿽");
                                    }
                                    Debugger.Break();
                                    item.Range.Text = sb.ToString() + item.Text.Substring(spsCountIndent);
                                    //range.End = ed_range + spsCount;//"􏿽".Length=2 - "　".Length=1 ;
                                    range.End += spsCountIndent;//"􏿽".Length=2 - "　".Length=1 ;
                                }


                            }
                            break;
                    }
                }
                //if(!range.Equals(document.Content))
                ////if (range != document.Content)
                //    range = document.Content;
                //textBox1.SelectedText = range.Text;
                textBox1.SelectionStart = range.End;
            }//if (linesCount_title == 1)
            #endregion

            #region 標題篇名前段補上分段符號
            int endPostion = textBox1.SelectionStart;
            //標題篇名前段補上分段符號
            i = x.LastIndexOf(Environment.NewLine, s);
            if (i > -1)
            {
                if (x.Substring(i > 3 ? i - 3 : i, 5).IndexOf("<p>") == -1)
                {
                    endCode = "<p>" + Environment.NewLine;
                    if (i + 2 + 2 <= x.Length && x.Substring(i + 2, 2) == Environment.NewLine)
                        endCode = "<p>";
                    #region 清除<p>前的空白􏿽（CTP網頁程式自會代勞，不必清除）
                    //int blankPosEnd = i;//清除<p>前的空白􏿽
                    //while (i - 2 > -2 && textBox1.Text.Substring(i - 2, 2) == "􏿽")
                    //    i -= 2;
                    //if (i < blankPosEnd)
                    //    textBox1.Text = textBox1.Text.Substring(0, i) + textBox1.Text.Substring(blankPosEnd);
                    #endregion
                    textBox1.Select(i, 2);
                    textBox1.SelectedText = endCode; endPostion += endCode.Length;
                }
            }
            #endregion //標題篇名前段補上分段符號

            if (!groupTitle)
                textBox1.Select(endPostion, 0);//將插入點置於標題尾端以便接著貼入Quit Edit中
            else
            {
                if (endPostion > Environment.NewLine.Length)
                    textBox1.Select(endPostion - Environment.NewLine.Length, 0);//是組詩標題的話，則將插入點移回標題行/段末，以供呼叫端定位用
            }

            keysTitleCode＿WithPrefaceNote();//處理「并序」
            stopUndoRec = false;
            return true;
        }
        /// <summary>
        /// 計算x中的行/段有幾個(不顧慮正文注文的差別，純粹只是計算分行/段符號的數量）
        /// </summary>
        /// <param name="x"></param>
        /// <returns>x中的總行/段數</returns>
        int linesCounter(string x)
        {
            int lineCnt = 1, i = x.IndexOf(Environment.NewLine);
            while (i > -1)
            {
                lineCnt++;
                i = x.IndexOf(Environment.NewLine, i + 1);
            }
            return lineCnt;
        }
        void keysTitleCode＿WithPrefaceNote()
        {//由 keysTitleCode 調用，keysTitleCode完成時是停在「并序」字前
            int s = textBox1.SelectionStart; bool replaceIt = false;
            string x = textBox1.Text; const string n = "<p>{{";
            if (x.Length < 12 || s < n.Length || s + 2 > x.Length) return;
            if (x.Substring(s, 2) == "{{") textBox1.SelectionStart = s += 2;
            if (s + 2 + 2 <= x.Length)
            {
                if (x.Substring(s + 2, 2) != "}}") return;
            }
            string px = x.Substring(s, 2);
            switch (px)
            {
                case "并叙":
                    replaceIt = true;
                    break;
                case "并序":
                    replaceIt = true;
                    break;
                case "幷序":
                    replaceIt = true;
                    break;
                case "並序":
                    replaceIt = true;
                    break;
                case "并敘":
                    replaceIt = true;
                    break;
                case "有序":
                    replaceIt = true;
                    break;
                case "并引":
                    replaceIt = true;
                    break;
                case "幷引":
                    replaceIt = true;
                    break;
                //case "*":
                //    break;

                default:
                    break;
            }
            if (replaceIt)
            {
                int ns = s - n.Length;
                textBox1.Select(ns, n.Length + px.Length);
                textBox1.SelectedText = "{{" + px + "　　";
                textBox1.SelectionStart += n.Length;
            }
        }
        /// <summary>
        /// 延展選取範圍至整個行/段
        /// </summary>
        /// <param name="s">選取開始</param>
        /// <param name="l">選取長度</param>
        /// <param name="x">整個文本</param>
        void expandSelectedTextRangeToWholeLinePara(int s, int l, string x)
        {
            int so = s;
            while (s >= 0)
            {
                if (s > 0)
                {
                    if (x.Substring(s - 1, 1) == Environment.NewLine.Substring(1, 1))
                    {
                        //s += 2;
                        break;
                    }
                    s--;
                }
                else break;
            }
            l += (so - s);
            while (s + l <= x.Length)
            {
                if (s + l + 1 < x.Length)
                {
                    if (x.Substring(s + (++l), 1) == Environment.NewLine.Substring(1, 1))
                    {
                        l--;
                        break;
                    }

                }
                else
                {
                    l++;
                    break;
                }
            }
            textBox1.Select(s, l);
        }
        /// <summary>
        /// Shift + F7 每行凸排
        /// 以textBox1操作
        /// 原名DeleteSpacePreParagraphs_ConvexRow
        /// </summary>
        private void outdent_ConvexRow()
        {
            List<Paragraph> paragraphs;
            int s = textBox1.SelectionStart, l = textBox1.SelectionLength, offset = 0; string x = textBox1.Text;

            //如果沒選取且在textBox的頭尾
            if (textBox1.SelectionLength == 0 && (s == 0 || s == textBox1.TextLength))
                paragraphs = _document.GetParagraphs();
            else
            {
                if (textBox1.Text.Substring(s + l - 2, 2) == Environment.NewLine)
                {
                    textBox1.SelectionLength++; l++;
                }
                paragraphs = _document.Range(textBox1.SelectionStart, textBox1.SelectionStart + textBox1.SelectionLength).Paragraphs;
            }
            PauseEvents();
            undoRecord();
            foreach (var paragraph in paragraphs)
            {
                string text = paragraph.Text;
                if (text.StartsWith("{{"))
                {
                    int endIndex = text.IndexOf("}}");
                    if (endIndex != -1)
                    {
                        noteOutdent(text, endIndex, paragraph, 0);
                    }
                    else if (text.Length > 2 && text[2] == '　')
                    {
                        paragraph.Text = text.Remove(2, 1); offset--;
                    }
                }
                else
                {
                    int endIndex = text.IndexOf("}}");
                    int openIndex = text.IndexOf("{{");
                    if (endIndex != -1 &&
                        (openIndex == -1 || (openIndex > -1 && openIndex > endIndex))
                        && (text.IndexOf("}}", endIndex + 1) == -1 && endIndex != text.Length - "}}".Length))//「}}」不會是該行、該段唯一的「}}」且在末尾。
                    {//開頭是注文，但沒有「{{」
                        noteOutdent(text, endIndex, paragraph, openIndex);
                    }
                    else//開頭不是注文
                    {
                        if (text.StartsWith("　"))
                        {
                            paragraph.Text = text.Remove(0, 1); offset--;
                        }
                    }

                }
            }
            if (paragraphs.Count == 0)
            {
                Debugger.Break();
                return;
            }
            paragraphs = _document.Range(paragraphs[0].Start, paragraphs[paragraphs.Count - 1].End).Paragraphs;
            //還原選取範圍或插入點所在處
            if (s == 0 && l == 0)
                textBox1.Select(s, l);
            else if (s == x.Length)
                textBox1.Select(textBox1.TextLength, 0);
            //else if (l == 0)//如果原無選取,只是插入點
            //    textBox1.Select(s + offset, 0);
            else
            {
                s = paragraphs[0].Start;
                l = paragraphs[paragraphs.Count - 1].End - s;
                textBox1.Select(s, l);
            }

            //finish:
            undoRecord();
            ResumeEvents();
            return;

            void noteOutdent(string text, int endIndex, Paragraph paragraph, int openIndex)
            {
                string before = text.Substring(0, endIndex + 2);
                StringInfo si;
                if (openIndex > -1)
                    si = new StringInfo(before);
                else
                    si = new StringInfo(before.Replace("{", string.Empty).Replace("}", string.Empty));
                int siLength = si.LengthInTextElements;
                int midBefore = (siLength % 2 == 0 ? siLength : siLength + 1) / 2;
                midBefore = si.SubstringByTextElements(0, midBefore).ToString().Length - 1;//因前半已移除1個全形空格，故須減1
                if (openIndex > -1)
                {
                    //if (before.IndexOf("《老子") > -1) Debugger.Break();

                    if (before.Length > 2 && before[2] == '　')
                    {
                        before = before.Remove(2, 1);
                    }

                    checkPunctuations();

                    if (before.Length > midBefore && before[midBefore] == '　')
                    {
                        before = before.Remove(midBefore, 1);
                    }
                }
                else
                {
                    if (before.Length > 0 && before[0] == '　')
                    {
                        before = before.Remove(0, 1);
                    }

                    checkPunctuations();

                    if (before.Length > midBefore && before[midBefore] == '　')
                    {
                        before = before.Remove(midBefore, 1);
                    }
                }
                offset -= 2;

                string after = text.Substring(endIndex + 2);
                paragraph.Text = before + after;

                return;

                void checkPunctuations()
                {
                    char before_midBefore = before[midBefore];
                    if (before_midBefore != '　' && PunctuationsNum.IndexOf(before_midBefore) > -1)
                        //如果不是空格而是標點符號，則向前移一格再檢查
                        if (midBefore > 0) midBefore--;
                }
            }


        }
        /// <summary>
        /// 凸排 Shift + F7 每行凸排
        /// 以 Range range 操作        
        /// </summary>
        /// <param name="document">要操作的文件（文本）</param>
        private void outdent_ConvexRow(ref Range range)
        {
            if (range.Text.Length == 0) return;

            List<Paragraph> paragraphs = range.Paragraphs;
            if (paragraphs.Count == 0) return;

            foreach (var paragraph in paragraphs)
            {
                string text = paragraph.Text;

                //if (text.Contains("杜甫")) Debugger.Break();

                if (text.StartsWith("{{"))
                {
                    int endIndex = text.IndexOf("}}");
                    if (endIndex != -1)
                    {
                        noteOutdent(text, endIndex, paragraph, 0);
                    }
                    else if (text.Length > 2 && text[2] == '　')
                    {
                        paragraph.Text = text.Remove(2, 1);
                    }
                    else if (text.Length > 3 && text[2] == "􏿽".ToArray()[0] && text[3] == "􏿽".ToArray()[1])
                    {
                        paragraph.Text = text.Remove(2, 2);
                    }
                }
                else//if (!text.StartsWith("{{"))
                {
                    int endIndex = text.IndexOf("}}");
                    int openIndex = text.IndexOf("{{");
                    if (endIndex != -1 &&
                        (openIndex == -1 || (openIndex > -1 && openIndex > endIndex))
                        && (text.IndexOf("}}", endIndex + 1) == -1 && endIndex != text.Length - "}}".Length))//「}}」不會是該行、該段唯一的「}}」且在末尾。
                    {//開頭是注文，但沒有「{{」
                        noteOutdent(text, endIndex, paragraph, openIndex);
                    }
                    else//開頭不是注文
                    {
                        if (text.StartsWith("　"))
                        {
                            paragraph.Text = text.Remove(0, 1);
                        }
                        else if (text.StartsWith("􏿽"))
                        {
                            paragraph.Text = text.Remove(0, 2);
                        }
                    }
                }
            }

            ////更新range//●●●●●●●●●●●●●●●//已改寫在 Paragraph中的 UpdateParentRange方法
            //if (range != paragraphs[0].Range.Document.Content)
            //    range = paragraphs[0].Range.Document.Content;

            //finish:            
            return;
            void noteOutdent(string text, int endIndex, Paragraph paragraph, int openIndex)
            {
                string before = text.Substring(0, endIndex + 2);
                StringInfo si;
                if (openIndex > -1)
                    si = new StringInfo(before);
                else
                    si = new StringInfo(before.Replace("{", string.Empty).Replace("}", string.Empty));
                int siLength = si.LengthInTextElements;
                int midBefore = (siLength % 2 == 0 ? siLength : siLength + 1) / 2;
                midBefore = si.SubstringByTextElements(0, midBefore).ToString().Length - 1;//因前半已移除1個全形空格，故須減1
                if (openIndex > -1)
                {
                    //if (before.IndexOf("《老子") > -1) Debugger.Break();

                    if (before.Length > 2 && before[2] == '　')
                    {
                        before = before.Remove(2, 1);
                    }
                    else if (before.Length > 3 && before[2] == "􏿽".ToArray()[0])
                    {
                        before = before.Remove(2, 2);
                    }

                    checkPunctuations();

                    if (before.Length > midBefore && before[midBefore] == '　')
                    {
                        before = before.Remove(midBefore, 1);
                    }
                    else if (before.Length > midBefore && before[midBefore] == "􏿽".ToArray()[0])
                    {
                        before = before.Remove(midBefore, 2);
                    }
                }
                else
                {
                    if (before.Length > 0 && before[0] == '　')
                    {
                        before = before.Remove(0, 1);
                    }
                    else if (before.Length > 0 && before[0] == "􏿽".ToArray()[0])
                    {
                        before = before.Remove(0, 2);
                    }

                    checkPunctuations();

                    if (before.Length > midBefore && before[midBefore] == '　')
                    {
                        before = before.Remove(midBefore, 1);
                    }
                    else if (before.Length > midBefore && before[midBefore] == "􏿽".ToArray()[0])
                    {
                        before = before.Remove(midBefore, 2);
                    }
                }

                string after = text.Substring(endIndex + 2);
                paragraph.Text = before + after;

                return;

                void checkPunctuations()
                {
                    char before_midBefore = before[midBefore];
                    if (before_midBefore != '　' && PunctuationsNum.IndexOf(before_midBefore) > -1)
                        //如果不是空格而是標點符號，則向前移一格再檢查
                        if (midBefore > 0) midBefore--;
                }
            }


        }

        #region 凸排 deleteSpacePreParagraphs_ConvexRow_Outdent舊式
        //private void deleteSpacePreParagraphs_ConvexRow()
        //{
        //    int s = textBox1.SelectionStart, l = textBox1.SelectionLength, cntr = 0;//, i;
        //    dontHide = true; string x = textBox1.Text, selTxt;
        //    if (l == 0)
        //    {
        //        if (s == 0 || s == textBox1.TextLength)
        //        {//全部凸排的機會少，若要全部，則請將插入點放在全文前端或末尾
        //            textBox1.SelectAll();
        //            l = textBox1.TextLength;
        //        }
        //        else { textBox1.Select(s, 1); l = 1; }
        //    }
        //    undoRecord(); stopUndoRec = true;
        //    //while (s - 1 > -1 && textBox1.Text.Substring(s--, 2) != Environment.NewLine)
        //    //{
        //    //    l++;
        //    //}
        //    ////while (e < textBox1.TextLength && textBox1.Text.Substring(e++, 2) != Environment.NewLine)
        //    ////{

        //    ////}
        //    //textBox1.Select(s, l + (so - s));
        //    //s = textBox1.SelectionStart; l = textBox1.SelectionLength;
        //    expandSelectedTextRangeToWholeLinePara(s, l, x);
        //    s = textBox1.SelectionStart; l = textBox1.SelectionLength;
        //    selTxt = textBox1.SelectedText;

        //    #region 凸排處理 20250204大年初六立秋~初七子夜
        //    string first = selTxt.Substring(0, 1);
        //    if (first == "　")
        //    {
        //        selTxt = selTxt.Substring(1);
        //        cntr++;
        //    }
        //    else if (first == "{")//下一個字元是「　」或「􏿽」
        //    {
        //        int sSpaceBalank = 2;//2="{{".Length
        //        if (selTxt.Substring(sSpaceBalank, 1) == "　")
        //        {
        //            selTxt = selTxt.Substring(0, sSpaceBalank) + selTxt.Substring(sSpaceBalank + 1);
        //            cntr++;
        //        }
        //        if (selTxt.Substring(sSpaceBalank, 1) == "􏿽")
        //        {
        //            selTxt = selTxt.Substring(0, sSpaceBalank) + selTxt.Substring(sSpaceBalank + 2);
        //            cntr++;
        //        }
        //    }
        //    else
        //    {
        //        first = selTxt.Substring(0, 2);
        //        if (first == "􏿽")
        //        {
        //            selTxt = selTxt.Substring(2);
        //            cntr += 2;
        //        }
        //    }//以上第1行/段凸排完成，以下到底凸排完成 
        //    int newline = selTxt.IndexOf(Environment.NewLine);
        //    while (newline > -1)
        //    {
        //        first = selTxt.Substring(newline + Environment.NewLine.Length, 1);
        //        if (first == "　")
        //        {
        //            selTxt = selTxt.Substring(0, newline + Environment.NewLine.Length)
        //                + selTxt.Substring(newline + Environment.NewLine.Length + 1);
        //            cntr++;
        //        }
        //        else if (first == "{")//下一個字元是「　」或「􏿽」
        //        {
        //            int sSpaceBalank = newline + Environment.NewLine.Length + 2;//2="{{".Length
        //            if (selTxt.Substring(sSpaceBalank, 1) == "　")
        //            {
        //                selTxt = selTxt.Substring(0, sSpaceBalank) + selTxt.Substring(sSpaceBalank + 1);
        //                cntr++;
        //            }
        //            if (selTxt.Substring(sSpaceBalank, 1) == "􏿽")
        //            {
        //                selTxt = selTxt.Substring(0, sSpaceBalank) + selTxt.Substring(sSpaceBalank + 2);
        //                cntr++;
        //            }
        //        }

        //        else
        //        {
        //            if (newline + Environment.NewLine.Length + 2 > selTxt.Length) break;
        //            first = selTxt.Substring(newline + Environment.NewLine.Length, 2);
        //            if (first == "􏿽")
        //            {
        //                selTxt = selTxt.Substring(0, newline + Environment.NewLine.Length)
        //                    + selTxt.Substring(newline + Environment.NewLine.Length + 2);
        //                cntr += 2;
        //            }
        //        }
        //        newline = selTxt.IndexOf(Environment.NewLine, newline + 1);
        //    }
        //    textBox1.SelectedText = selTxt;

        //    #endregion

        //    #region 以下舊式 20250204 大年初七
        //    //#region 將第一行/段是空白「􏿽」而後面是空格「　」的縮排改成都是空格「　」 20250124
        //    //if (selTxt.Substring(0, 2) == "􏿽" && selTxt.IndexOf(Environment.NewLine) > -1
        //    //        && selTxt.Substring(selTxt.IndexOf(Environment.NewLine) + Environment.NewLine.Length, 1) == "　")
        //    //{
        //    //    textBox1.SelectedText = "　" + textBox1.SelectedText.Substring(2);//這樣會取消選取
        //    //    l--;
        //    //    textBox1.Select(s, l);
        //    //    selTxt = textBox1.SelectedText;
        //    //}
        //    //#endregion

        //    //if (selTxt.Length > 1 && selTxt.Substring(0, 2) == "􏿽")//(textBox1.SelectedText.IndexOf("􏿽") > -1)
        //    //{
        //    //    i = selTxt.IndexOf(Environment.NewLine + "􏿽");
        //    //    while (i > -1)
        //    //    {
        //    //        cntr++;
        //    //        i = selTxt.IndexOf(Environment.NewLine + "􏿽", i + 1);
        //    //    }
        //    //    if (textBox1.SelectedText.Substring(0, 2) == "􏿽") textBox1.SelectedText = textBox1.SelectedText.Substring(2);
        //    //    l -= "􏿽".Length;
        //    //    textBox1.Select(s, l);
        //    //    textBox1.SelectedText = textBox1.SelectedText.Replace(Environment.NewLine + "􏿽", Environment.NewLine);
        //    //    cntr *= 2;
        //    //}
        //    //else
        //    //{
        //    //    i = selTxt.IndexOf(Environment.NewLine + "　");
        //    //    while (i > -1)
        //    //    {
        //    //        cntr++;
        //    //        i = selTxt.IndexOf(Environment.NewLine + "　", i + 1);
        //    //    }
        //    //    if (textBox1.SelectedText.Substring(0, 1) == "　") textBox1.SelectedText = textBox1.SelectedText.Substring(1);
        //    //    l -= "　".Length;
        //    //    textBox1.Select(s, l);
        //    //    textBox1.SelectedText = textBox1.SelectedText.Replace(Environment.NewLine + "　", Environment.NewLine);
        //    //}

        //    //////自取消再觀察！ 20240927
        //    ////if (s == 0)
        //    ////{
        //    ////    if ("　".IndexOf(textBox1.Text.Substring(0, 1)) > -1)
        //    ////        textBox1.Text = textBox1.Text.Substring(1);
        //    ////    else if ("􏿽".IndexOf(textBox1.Text.Substring(0, "􏿽".Length)) > -1)
        //    ////        textBox1.Text = textBox1.Text.Substring("􏿽".Length);
        //    ////}//以上自取消再觀察！ 20240927
        //    #endregion

        //    textBox1.Select(s, l - cntr);
        //    stopUndoRec = false;
        //    dontHide = false;
        //}
        #endregion


        int indentRow()
        {//每行縮排 //此函式執行完時會將執行結果的範圍選取，以便後續處理。傳回值為處理了幾行/段
            int s = textBox1.SelectionStart; int l = textBox1.SelectionLength; String xn, x = textBox1.Text;
            bool stopUndoRecFlag = false;
            if (textBox1.SelectedText == "")//全部縮排的機會少，若要全部，則請將插入點放在全文前端或末尾
            {
                if (s == 0 || s == textBox1.TextLength || (s == textBox1.TextLength - 2 && textBox1.Text.Substring(s, 2) == Environment.NewLine))
                {
                    if (!keyinTextMode) pasteAllOverWrite = true;
                    if (s == textBox1.TextLength - 2 && textBox1.Text.Substring(s, 2) == Environment.NewLine)
                    {
                        textBox1.Text = textBox1.Text.Substring(0, s);
                    }
                    textBox1.SelectAll(); s = 0; l = textBox1.TextLength;
                }
                //已有「expandSelectedTextRangeToWholeLinePara」此行當不必下行
                //else { textBox1.Select(s, 1); l = 1; }

            }
            //else //先延展選取範圍至整個行/段
            expandSelectedTextRangeToWholeLinePara(s, l, x);
            String slTxt = textBox1.SelectedText; int i = slTxt.IndexOf(Environment.NewLine), cntr = 0;
            while (i > -1)
            {
                cntr++;//計下處理了幾行/段
                i = slTxt.IndexOf(Environment.NewLine, i + 1);
            }
            if (!stopUndoRec) { undoRecord(); caretPositionRecord(); stopUndoRec = true; stopUndoRecFlag = true; }
            //s = textBox1.SelectionStart;
            if (s == 0 || s > 2 && textBox1.Text.Substring(s - 2, 2) == Environment.NewLine)
            {
                xn = textBox1.SelectedText.Replace(Environment.NewLine, Environment.NewLine + "　");
                //if (s > 0) s = s - "　".Length;
                l = ("　" + xn).Length;
                textBox1.SelectedText = ("　" + xn).Replace(Environment.NewLine + "　{{", Environment.NewLine + "{{　");
            }
            else
            {
                //int f = textBox1.Text.LastIndexOf(Environment.NewLine, s);
                xn = textBox1.SelectedText.Replace(Environment.NewLine, Environment.NewLine + "　");
                textBox1.SelectedText = ("　" + xn).Replace(Environment.NewLine + "　{{", Environment.NewLine + "{{　"); s -= "　".Length;
                //if (textBox1.SelectedText == xn)
                //textBox1.SelectedText = "　" + textBox1.SelectedText;
                //else
                //textBox1.SelectedText = xn;

                //textBox1.Select(f == -1 ? 0 : f + 2, s - f);//只讀取了第一行前端
                //s = textBox1.SelectionStart - "　".Length; if (s < 0) s = 0;
                l = ("　" + xn + textBox1.SelectionLength).Length;
                //textBox1.SelectedText = "　" + textBox1.SelectedText;
            }
            textBox1.Select(s, l);//將執行結果的範圍選取，以便後續處理。
            pasteAllOverWrite = false;
            if (stopUndoRecFlag) stopUndoRec = false;
            return cntr;
        }
        /// <summary>
        /// 每行/段縮排（按下F7改指定於以Range操作者）
        /// 以textBox1操作
        /// </summary>
        private void indent()
        {
            List<Paragraph> paragraphs;
            int s = textBox1.SelectionStart, l = textBox1.SelectionLength, offset = 0; string x = textBox1.Text;

            //如果沒選取且在textBox的頭尾
            if (textBox1.SelectionLength == 0 && (s == 0 || s == textBox1.TextLength))
                paragraphs = _document.GetParagraphs();
            else
            {
                if (l > 1 && textBox1.Text.Substring(s + l - 2, 2) == Environment.NewLine)
                {
                    textBox1.SelectionLength++; l++;
                }
                paragraphs = _document.Range(textBox1.SelectionStart, textBox1.SelectionStart + textBox1.SelectionLength).Paragraphs;
            }
            PauseEvents();
            undoRecord();
            foreach (var paragraph in paragraphs)
            {
                string text = paragraph.Text;
                if (text.StartsWith("{{"))
                {
                    int endIndex = text.IndexOf("}}");
                    if (endIndex != -1
                        //須不是獨立注文才執行
                        && ((text.IndexOf("<p>") == -1 && text.IndexOf("}}", endIndex + 1) == -1 && endIndex != text.Length - "}}".Length)
                        || (text.IndexOf("}}<p>") > -1 && text.IndexOf("}}", endIndex + 1) == -1 && endIndex != text.Length - "}}<p>".Length)
                        || (text.IndexOf("}}。<p>") > -1 && text.IndexOf("}}", endIndex + 1) == -1 && endIndex != text.Length - "}}。<p>".Length)
                        || (text.IndexOf("}}", endIndex + 1) > -1 || text.IndexOf("{{", endIndex + 1) > -1)))

                    {
                        noteIndent(text, endIndex, paragraph, 0);
                    }
                    else
                    {
                        paragraph.Text = text.Insert(2, "　"); offset++;
                    }
                }
                else//if (!text.StartsWith("{{"))
                {
                    int endIndex = text.IndexOf("}}");
                    int openIndex = text.IndexOf("{{");
                    if (endIndex != -1
                        && (openIndex > -1 && endIndex < openIndex)
                        && (openIndex == -1 || (openIndex > -1 && openIndex > endIndex))
                        //須不是獨立注文才執行
                        //「}}」不會是該行、該段唯一的「}}」且在末尾。
                        && ((text.IndexOf("<p>") == -1 && text.IndexOf("}}", endIndex + 1) == -1 && endIndex != text.Length - "}}".Length)
                            || (text.IndexOf("}}<p>") > -1 && text.IndexOf("}}", endIndex + 1) == -1 && endIndex != text.Length - "}}<p>".Length)
                            || (text.IndexOf("}}。<p>") > -1 && text.IndexOf("}}", endIndex + 1) == -1 && endIndex != text.Length - "}}。<p>".Length))
                        || ((openIndex > -1 && endIndex < openIndex) && (text.IndexOf("}}", endIndex + 1) > -1 || text.IndexOf("{{", endIndex + 1) > -1)))
                    {//開頭是注文，但沒有「{{」
                        noteIndent(text, endIndex, paragraph, openIndex);
                    }
                    else//開頭不是注文
                    {
                        paragraph.Text = "　" + text; offset++;
                    }

                }
            }
            if (paragraphs.Count == 0)
            {
                Debugger.Break();
                return;
            }
            paragraphs = _document.Range(paragraphs[0].Start, paragraphs[paragraphs.Count - 1].End).Paragraphs;
            //還原選取範圍或插入點所在處
            if (s == 0 && l == 0)
                textBox1.Select(s, l);
            else if (s == x.Length)
                textBox1.Select(textBox1.TextLength, 0);
            //else if (l == 0)//如果原無選取,只是插入點
            //    textBox1.Select(s + offset, 0);
            else
            {
                s = paragraphs[0].Start;
                l = paragraphs[paragraphs.Count - 1].End - s;
                textBox1.Select(s, l);
            }

            //finish:
            undoRecord();
            ResumeEvents();
            return;

            void noteIndent(string text, int endIndex, Paragraph paragraph, int openIndex)
            {
                string before = text.Substring(0, endIndex + 2);
                StringInfo si;
                if (openIndex > -1)
                    si = new StringInfo(before);
                else
                    si = new StringInfo(before.Replace("{", string.Empty).Replace("}", string.Empty));
                int siLength = si.LengthInTextElements;

                int midBefore = (siLength % 2 == 0 ? siLength : siLength + 1) / 2;
                if (midBefore > 0)
                    midBefore = si.SubstringByTextElements(0, midBefore).ToString().Length + 1;//因前半已插入1個全形空格，故須加1
                string midBeforeInsertString = new StringInfo(before.Replace("{{", string.Empty).Replace("}}", string.Empty)).LengthInTextElements == 1 ? string.Empty : "　";
                if (openIndex > -1)
                {
                    before = before.Insert(2, "　").Insert(midBefore, midBeforeInsertString);
                }
                else
                {
                    before = before.Insert(0, "　").Insert(midBefore, midBeforeInsertString);
                }
                offset += 2;

                string after = text.Substring(endIndex + 2);
                paragraph.Text = before + after;
            }

        }
        /// <summary>
        /// 按下F7：每行/段縮排
        /// </summary>
        /// <param name="range">要縮排的文本範圍</param>
        private void indent(ref Range range)
        {
            if (range.Text.Length == 0) return;
            List<Paragraph> paragraphs = range.Paragraphs;
            //List<Paragraph> paragraphs = range.Document.GetParagraphs(ref range);
            if (paragraphs.Count == 0) return;

            foreach (var paragraph in paragraphs)
            {
                string text = paragraph.Text;
                if (text.StartsWith("{{"))
                {
                    int endIndex = text.IndexOf("}}");
                    if (endIndex != -1
                        //須不是獨立注文才執行
                        && ((text.IndexOf("<p>") == -1 && text.IndexOf("}}", endIndex + 1) == -1 && endIndex != text.Length - "}}".Length)
                        || (text.IndexOf("}}<p>") > -1 && text.IndexOf("}}", endIndex + 1) == -1 && endIndex != text.Length - "}}<p>".Length)
                        || (text.IndexOf("}}。<p>") > -1 && text.IndexOf("}}", endIndex + 1) == -1 && endIndex != text.Length - "}}。<p>".Length)
                        || (text.IndexOf("}}", endIndex + 1) > -1 || text.IndexOf("{{", endIndex + 1) > -1))
                        || (text.LastIndexOf("}}") > -1 && text.LastIndexOf("}}") < text.Length - "{{".Length)//行尾端是正文
                        )

                    {
                        noteIndent(text, endIndex, paragraph, 0);
                    }
                    else
                    {
                        paragraph.Text = text.Insert(2, "　");
                    }
                }
                else//if (!text.StartsWith("{{"))
                {
                    int endIndex = text.IndexOf("}}");
                    int openIndex = text.IndexOf("{{");
                    if (endIndex != -1
                        && (openIndex > -1 && endIndex < openIndex)
                        && (openIndex == -1 || (openIndex > -1 && openIndex > endIndex))
                        //須不是獨立注文才執行
                        //「}}」不會是該行、該段唯一的「}}」且在末尾。
                        && ((text.IndexOf("<p>") == -1 && text.IndexOf("}}", endIndex + 1) == -1 && endIndex != text.Length - "}}".Length)
                            || (text.IndexOf("}}<p>") > -1 && text.IndexOf("}}", endIndex + 1) == -1 && endIndex != text.Length - "}}<p>".Length)
                            || (text.IndexOf("}}。<p>") > -1 && text.IndexOf("}}", endIndex + 1) == -1 && endIndex != text.Length - "}}。<p>".Length))
                        || (endIndex != -1 && (openIndex > -1 && endIndex < openIndex) && (text.IndexOf("}}", endIndex + 1) > -1 || text.IndexOf("{{", endIndex + 1) > -1))
                        //|| (text.LastIndexOf("}}") > -1 && text.LastIndexOf("}}") < text.Length - "{{".Length)
                        )
                    {//開頭是注文，但沒有「{{」
                        noteIndent(text, endIndex, paragraph, openIndex);
                    }
                    else//開頭不是注文
                    {
                        paragraph.Text = "　" + text;
                    }
                    //paragraphs = range.Document.GetParagraphs(true);
                }
            }
            //finish:
            ////更新range 已在Paragraph的Text中以UpdateParentRange更新了 20250227
            //if (range.Start != paragraphs[0].Range.Document.Range().Start ||
            //    range.End != paragraphs[0].Range.Document.Range().End)
            //    range = paragraphs[0].Range.Document.Range();

            return;

            void noteIndent(string text, int endIndex, Paragraph paragraph, int openIndex)
            {
                string before = text.Substring(0, endIndex + 2);
                StringInfo si;
                if (openIndex > -1)
                    si = new StringInfo(before);
                else
                    si = new StringInfo(before.Replace("{", string.Empty).Replace("}", string.Empty));
                int siLength = si.LengthInTextElements;

                int midBefore = (siLength % 2 == 0 ? siLength : siLength + 1) / 2;
                if (midBefore > 0)
                    midBefore = si.SubstringByTextElements(0, midBefore).ToString().Length + 1;//因前半已插入1個全形空格，故須加1
                                                                                               //string midBeforeInsertString =  "　";
                                                                                               //string midBeforeInsertString = new StringInfo(before.Replace("{{", string.Empty).Replace("}}", string.Empty)).LengthInTextElements == 1 ? string.Empty : "　";
                                                                                               //if(text.EndsWith("<p>") && endIndex< text.IndexOf("<p>")- text.Length-"}}".Length ||
                                                                                               //text.EndsWith("。<p>") && endIndex < text.IndexOf("。<p>") - text.Length - "}}".Length )
                                                                                               //midBeforeInsertString = string.Empty;
                string midBeforeInsertString = (text.EndsWith("<p>") && endIndex < text.IndexOf("<p>") - text.Length - "}}".Length ||
                                                text.EndsWith("。<p>") && endIndex < text.IndexOf("。<p>") - text.Length - "}}".Length)
                                                ? string.Empty : "　";
                if (openIndex > -1)
                {
                    before = before.Insert(2, "　").Insert(midBefore, midBeforeInsertString);
                }
                else
                {
                    before = before.Insert(0, "　").Insert(midBefore, midBeforeInsertString);
                }

                string after = text.Substring(endIndex + 2);
                paragraph.Text = before + after;
            }

        }
        #region keysSpacePreParagraphs_indent舊式

        //int l = textBox1.SelectionLength; int s = textBox1.SelectionStart; dontHide = true;
        //bool allIndent = s == textBox1.TextLength || s == 0 ? true : false;
        //if (l == textBox1.TextLength)
        //{
        //    l = 0;
        //}
        //undoRecord(); stopUndoRec = true; PauseEvents();
        //int cntr = indentRow();//此函式執行完時會將執行結果的範圍選取，以便後續處理。傳回值為處理了幾行/段
        //                       //if (l != 0)
        //                       //{
        //                       //textBox1.Select(s, l + 1 + cntr);                
        //                       //}
        //                       //textBox1.Select(s + 1 + cntr, l);
        //undoRecord(); stopUndoRec = false; ResumeEvents();
        //if (!allIndent)
        //    textBox1.Select(s + 1, l + cntr);
        //else
        //    textBox1.Select(0, 0);
        //dontHide = false;

        #endregion

        #region 原式
        /*
        int s = textBox1.SelectionStart;
        if (textBox1.SelectedText == "")
        {
            pasteAllOverWrite = true;
            textBox1.SelectAll();
        }
        int l = textBox1.SelectionLength;
        if (l == textBox1.TextLength)
        {
            l = 0;
        }
        String slTxt = textBox1.SelectedText; int i = slTxt.IndexOf(Environment.NewLine), cntr = 0;
        while (i > -1)
        {
            cntr++;
            i = slTxt.IndexOf(Environment.NewLine, i + 1);
        }
        undoRecord(); caretPositionRecord(); stopUndoRec = true;
        if (s == 0 || s > 2 && textBox1.Text.Substring(s - 2, 2) == Environment.NewLine)
        {

            textBox1.SelectedText = "　" + textBox1.SelectedText.Replace(Environment.NewLine, Environment.NewLine + "　");
        }
        else
        {
            int f = textBox1.Text.LastIndexOf(Environment.NewLine, s);
            textBox1.SelectedText = textBox1.SelectedText.Replace(Environment.NewLine, Environment.NewLine + "　");
            textBox1.Select(f == -1 ? 0 : f + 2, s - f);
            textBox1.SelectedText = "　" + textBox1.SelectedText;
        }
        if (l != 0)
        {
            textBox1.Select(s, l + 1 + cntr);
        }
        pasteAllOverWrite = false;
        stopUndoRec = false;
        */
        #endregion


        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, Int32 wMsg, bool wParam, Int32 lParam);
        private const int WM_SETREDRAW = 11;

        /// <summary>
        /// Alt + F7 (先改 Pause/Break）: 每行縮排一格後、清除其末誤標之<p>
        /// </summary>
        private void indent_ClearEnd＿P_Mark()
        {
            int l = textBox1.SelectionLength; int s = textBox1.SelectionStart; dontHide = true;
            if (l == 0)
            {
                if (s == 0 || s == textBox1.TextLength)
                {//全部縮排的機會少，若要全部，則請將插入點放在全文前端或末尾
                    textBox1.SelectAll();
                    l = textBox1.TextLength;
                }
                //else { textBox1.Select(s, 1); l = 1; }
            }

            if (l == textBox1.TextLength)
            {
                l = 0;
            }
            undoRecord(); stopUndoRec = true;
            expandSelectedTextRangeToWholeLinePara(s, l, textBox1.Text);
            int cntr = indentRow();//此函式執行完時會將執行結果的範圍選取，以便後續處理。傳回值為處理了幾行/段
            expandSelectedTextRangeToWholeLinePara(s, l, textBox1.Text);
            //if (l != 0)
            //{
            //    textBox1.Select(s, l + 1 + cntr - cntr * "<p>".Length);
            //}

            //http://stackoverflow.com/questions/487661/how-do-i-suspend-painting-for-a-control-and-its-children
            //https://stackoverflow.com/questions/126876/how-do-i-disable-updating-a-form-in-windows-forms
            SendMessage(this.Handle, WM_SETREDRAW, false, 0);

            while (textBox1.SelectionStart + textBox1.SelectionLength + 2 <= textBox1.TextLength
                    && textBox1.Text.Substring(textBox1.SelectionStart + textBox1.SelectionLength, 2) != Environment.NewLine
                    && textBox1.SelectedText.Substring(textBox1.SelectedText.Length - 3) != "<p>")
            {//找到處理範圍裡最後一個<p>，若碰到換行而無<p>者，即停止
                textBox1.Select(textBox1.SelectionStart, textBox1.SelectionLength++);
                if (textBox1.Text.IndexOf(Environment.NewLine) == -1 || textBox1.Text.IndexOf("<p>") == -1) break;
            }
            if (textBox1.SelectedText != string.Empty && "<p>".IndexOf(textBox1.SelectedText.Substring(0, 1)) > -1)
            {
                while (textBox1.SelectionStart >= 1 && textBox1.SelectedText.Substring(0, 3) != "<p>")
                {
                    textBox1.SelectionStart--; textBox1.SelectionLength++;
                    if (textBox1.SelectionStart == 0) break;
                }
            }
            if (textBox1.SelectionStart + textBox1.SelectionLength +
                +Environment.NewLine.Length + "　".Length <= textBox1.TextLength)
            {
                if (textBox1.Text.Substring(textBox1.SelectionStart + textBox1.SelectionLength
                    + Environment.NewLine.Length, "　".Length) != "　" &&
                    textBox1.SelectionLength - 3 > -1 &&
                    textBox1.SelectedText.Substring(textBox1.SelectionLength - 3) == "<p>")
                {//如果接下來是頂行，則不取代最後的<p>
                    textBox1.Select(textBox1.SelectionStart, textBox1.SelectionLength - "<p>".Length);
                }
            }
            s = textBox1.SelectionStart; l = textBox1.SelectionLength;
            int i = textBox1.SelectedText.IndexOf("􏿽");
            while (i != -1)
            {
                if (textBox1.SelectedText.Substring(i - 1, 1) == "　")
                {
                    textBox1.Select(s + i, "􏿽".Length);
                    if (textBox1.SelectedText == "􏿽") textBox1.SelectedText = "";//textBox1.SelectedText = "　";
                    l -= ("􏿽".Length - "　".Length);
                    textBox1.Select(s, l);
                }
                i = textBox1.SelectedText.IndexOf("􏿽", i + 1);
            }
            //textBox1.SelectedText = textBox1.SelectedText.Replace("<p>" + Environment.NewLine, Environment.NewLine);            
            //textBox1.Select(s, l);
            textBox1.SelectedText = textBox1.SelectedText.Replace("<p>", "");
            // Do your thingies here

            SendMessage(this.Handle, WM_SETREDRAW, true, 0);
            this.Refresh();
            stopUndoRec = false;
            dontHide = false;
        }
        /// <summary>
        /// Ctrl + Alt + p ： 清除textBox1選取內的`<p>`
        /// </summary>
        private void clearKeysParagraphSymbol()//(bool period = false)
        {
            if (textBox1.SelectedText.IndexOf("<p>") == -1) return;
            textBox1.SelectedText = textBox1.SelectedText.Replace("<p>", string.Empty);
        }


        /// <summary>
        /// 插入標識分行符號/分段符號 Alt + p  
        /// Alt + Shift + p
        /// </summary>
        private void keysParagraphSymbol(bool period = false)
        {
            if (textBox1.TextLength < 2) return; int s = textBox1.SelectionStart;
            string x = textBox1.Text, stxtPre = x.Substring(s < 2 ? s : s - 2, 2);
            if (EventsEnabled) undoRecord();
            stopUndoRec = true;
            string insertX = period ? "。<p>" : "<p>";
            if (stxtPre == Environment.NewLine && s > 1)
                textBox1.SelectionStart = s - 2 > 0 ? s - 2 : 0;
            else if (stxtPre.IndexOf("|", 1) > -1)
            {
                textBox1.Select(s - 1, 1);
                textBox1.SelectedText = "";
            }
            if (s + 2 >= x.Length || x.Substring(s, 2) == Environment.NewLine ||
                        x.Substring(s - 2 < 0 ? 0 : s - 2, 2) == Environment.NewLine)
                //insertWords("<p>", textBox1, textBox1.Text);
                //insertWords(insertX, textBox1, textBox1.Text);
                insertWords(insertX, textBox1);
            else
                //insertWords("<p>" + Environment.NewLine, textBox1, textBox1.Text);
                //insertWords(insertX + Environment.NewLine, textBox1, textBox1.Text);
                insertWords(insertX + Environment.NewLine, textBox1);
            if (x.Substring(s - 2 < 0 ? 0 : s - 2, 2) == Environment.NewLine)
            {
                if (s - 3 > -1 && x.Substring(s - 3, 1) == "|")
                {//清除「|」
                    textBox1.Select(s - 3, 1);
                    textBox1.SelectedText = string.Empty;
                    s--;
                }
                //textBox1.SelectionStart = s + "<p>".Length; textBox1.ScrollToCaret();
                textBox1.SelectionStart = s + insertX.Length; textBox1.ScrollToCaret();
            }

            stopUndoRec = false;
        }

        bool stopUndoRec = false;

        /// <summary>
        /// 清除標題符碼標記，並讀入剪貼簿中
        /// </summary>
        void clearTitleMarkCode()
        {
            string x = textBox1.Text;
            Regex rx = new Regex("[　*。<p>]");
            x = rx.Replace(x, string.Empty);
            if (!x.IsNullOrEmpty())
                try
                {
                    Clipboard.SetText(textBox1.Text = x);
                }
                catch (Exception)
                {
                    playSound(soundLike.error, true);
                }
            else
                textBox1.Text = x;
        }
        /// <summary>
        /// Ctrl + Shift + Delete ： 將選取文字於文本中全部清除(Ctrl + z 還原功能支援)
        /// 若是選取《·》〈〉{{}}以執行，則會清除相對應的符號，以便書名號篇名號及注文語法標記之增修。
        /// 若是選取「*」或「。<p>」則清除「*」或「。<p>」（即清除OCR模式下自動標識的標題暨段落符碼
        /// </summary>
        private void clearSeltxt()
        {
            caretPositionRecord();
            string xClear = textBox1.SelectedText, x = textBox1.Text;
            int s = textBox1.SelectionStart;
            undoRecord();
            if (xClear == "")
            {
                if (CnText.ClearHasEditedWithPunctuationMarks(ref x))
                {
                    textBox1.Text = x;
                    Clipboard.SetText(x);//寫入剪貼簿以備用，如重新標點符號。
                }
            }
            else
            {
                int xLen = x.Length, index = x.Substring(0, (s == 0 ? s : s - 1)).IndexOf(xClear);
                if ("{{}}".IndexOf(xClear) > -1)//自行將所有大括弧清除
                                                //textBox1.Text = textBox1.Text.Replace("{", "").Replace("}", "");
                    textBox1.Text = x.Replace("{", "").Replace("}", "");
                else if ("《·》〈〉".IndexOf(xClear) > -1)
                {//若是選取《·》〈〉{{}}以執行，則會清除相對應的符號，以便書名號篇名號及注文語法標記之增修。
                    Regex rx = new Regex("[《·》〈〉]");
                    textBox1.Text = rx.Replace(x, string.Empty);
                    Clipboard.SetText(textBox1.Text);//以便按下 Alt + Insert 檢視書名號篇名號增修之結果。20231124
                }
                else if ("*。<p>".IndexOf(xClear) > -1
                    && textBox1.Text.IndexOf("*") > -1)
                {//若是選取「*」或「。<p>」則清除「*」或「。<p>」（即清除OCR模式下自動標識的標題暨段落符碼                    
                 //Regex rx = new Regex("[*。<p>]");
                 //textBox1.Text = rx.Replace(x, string.Empty);
                    clearTitleMarkCode();
                }
                else
                    textBox1.Text = x.Replace(xClear, "");
                if (index > -1) s = -(xLen - textBox1.TextLength);
            }
            undoRecord();
            caretPositionRecall();
            if (s > 0) restoreCaretPosition(textBox1, s, 0);
            if (textBox1.SelectionStart == 0 && s > 0)
                textBox1.SelectionStart = s;
            //textBox1.SelectionStart = selStart;
            //textBox1.ScrollToCaret();
        }

        private void keyDownCtrlAltUpDown(KeyEventArgs e)
        {
            e.Handled = true;
            int s = textBox1.SelectionStart; string x = textBox1.Text;
            if (e.KeyCode == Keys.Down)
            {
                if (s == x.Length) goto notFound;
                if (s + Environment.NewLine.Length > x.Length) goto notFound;
                s = x.IndexOf(Environment.NewLine, s + Environment.NewLine.Length);
                if (s > x.Length) goto notFound;
            }
            else
            {
                if (s == 0) goto notFound;
                if (s - Environment.NewLine.Length < 0) goto notFound;
                s = x.LastIndexOf(Environment.NewLine, s - Environment.NewLine.Length) + Environment.NewLine.Length;
                if (s < 0) goto notFound;
            }
            if (s > -1)
                textBox1.SelectionStart = s;
            else
                goto notFound;
            textBox1.ScrollToCaret();
            return;
        notFound:
            MessageBox.Show("not found!");

        }

        private void clearNewLinesAfterCaret()
        {
            //clear the newline after the caret
            string x = textBox1.Text;
            int s = textBox1.SelectionStart;
            string xNext = x.Substring(s);
            x = x.Substring(0, textBox1.SelectionStart);
            xNext = xNext.Replace(Environment.NewLine, "");
            NormalLineParaLength = 0;
            x += xNext;
            textBox1.Text = x;
            textBox1.SelectionStart = s;// textBox1.SelectionLength = 1;
            restoreCaretPosition(textBox1, s, 0);//textBox1.ScrollToCaret();
        }

        /// <summary>
        /// 記下還原了幾次
        /// </summary>
        private int undoTimes;
        /// <summary>
        /// 設定不要進行還原記錄
        /// </summary>
        bool undoTextBoxing = false;

        /// <summary>
        /// Ctrl + z 還原機制，目前上限為50個記錄
        /// </summary>
        /// <param name="textBox1"></param>
        private void undoTextBox(TextBox textBox1)
        {
            int s = textBox1.SelectionStart, l = textBox1.SelectionLength;
            if (selStart != s && selStart != 0)
            {
                s = selStart; l = selLength;
            }
            if (undoTextBox1Text.Count - undoTimes - 1 > -1)
            {
                //string x = undoTextBox1Text[undoTextBox1Text.Count - ++undoTimes];
                //string x = undoTextBox1Text[undoTextBox1Text.Count - 1 - ++undoTimes];//20241001
                string x = undoTextBox1Text[undoTextBox1Text.Count - ++undoTimes];//20241001
                if (textBox1.Text == x && undoTextBox1Text.Count - (undoTimes + 1) > 0)//●●●●●●●●●●●●●●●20250304
                    x = undoTextBox1Text[undoTextBox1Text.Count - ++undoTimes];
                else
                    x = undoTextBox1Text[undoTextBox1Text.Count - undoTimes];
                while (x == "")
                {
                    if (undoTextBox1Text.Count - undoTimes - 1 < 0) break;
                    x = undoTextBox1Text[undoTextBox1Text.Count - ++undoTimes];

                }
                if (x != "")
                {
                    undoTextBoxing = true;
                    textBox1.Text = x;
                    restoreCaretPosition(textBox1, s, l);
                    undoTextBoxing = false;
                }
            }
            else
                MessageBox.Show("no more to undo!");

        }
        /// <summary>
        /// Ctrl + y 重做（即復原還原的動作），目前上限為300個記錄，詳此行操作： if (undoTextBox1Text.Count > 300)
        /// </summary>
        /// <param name="textBox1"></param>
        private void redoTextBox(TextBox textBox1)
        {
            if (undoTimes == 0) { MessageBox.Show("no more to redo!"); return; }
            int s = textBox1.SelectionStart, l = textBox1.SelectionLength;
            if (selStart != s && selStart != 0)
            {
                s = selStart; l = selLength;
            }
            if (undoTextBox1Text.Count - undoTimes - 1 > -1)
            {
                string x = undoTextBox1Text[undoTextBox1Text.Count - --undoTimes - 1];
                while (x == "")
                {
                    if (undoTextBox1Text.Count - --undoTimes - 1 < 0) break;
                    if (undoTextBox1Text.Count - --undoTimes - 1 > undoTextBox1Text.Count - 1) break;
                    x = undoTextBox1Text[undoTextBox1Text.Count - --undoTimes - 1];
                }
                if (x != "")
                {
                    if (x == textBox1.Text && undoTimes > 0)//●●●●●●●●●●●●●●●●●
                        x = undoTextBox1Text[undoTextBox1Text.Count - --undoTimes - 1];

                    undoTextBoxing = true;
                    //textBox1.Text = undoTextBox1Text[undoTextBox1Text.Count - ++undoTimes];
                    textBox1.Text = x;
                    restoreCaretPosition(textBox1, s, l);
                    undoTextBoxing = false;
                }
            }
            else
                MessageBox.Show("no more to redo!");

        }

        private static void restoreCaretPosition(TextBox textBox1, int s, int l)
        {
            if (s < 0 || l < 0) return;
            textBox1.Select(s, l);
            Point caretPosition = textBox1.GetPositionFromCharIndex(s);//c# caret position: https://stackoverflow.com/questions/37782986/how-to-find-the-caret-position-in-a-textbox-using-c
            if (caretPosition.Y > textBox1.Height - textBox1.Top || caretPosition.Y < textBox1.Top)
            {
                textBox1.ScrollToCaret();
                textBox1.AutoScrollOffset = caretPosition;//還不行！再研究 20220723
                                                          //ScrollableControl scrl = new ScrollableControl();
                                                          //scrl.ScrollControlIntoView(textBox1);
            }
        }
        /// <summary>
        /// 計算xPage中的實際行/段數（正文算1，注文算2，故8行正文之文本傳回的值即16）
        /// </summary>
        /// <param name="xPage">要計算的文本</param>
        /// <returns>xPage中的實際行/段數</returns>
        int countLinesPerPage(string xPage)
        {
            #region 規範文本

            while (xPage.Contains("􏿽<p>"))
            {
                xPage = xPage.Replace("􏿽<p>", string.Empty);
            }

            #endregion
            int i = 0, openBracketS, closeBracketS, e = xPage.IndexOf(Environment.NewLine); bool openNote = false;
            string[] linesParasPage = xPage.Split(Environment.NewLine.ToArray(), StringSplitOptions.RemoveEmptyEntries);
            string item;
            if (linesParasPage.Length == 1)
            {
                item = linesParasPage[0];
                if (item == "|") return 2;
                else
                {
                    if (item.IndexOf("{{") == -1 && item.IndexOf("}}") == -1)
                        return 2;
                    else
                    {
                        if ((item.StartsWith("{{") && item.IndexOf("}}") == -1)
                            || (item.EndsWith("}}") && item.IndexOf("{{") == -1)
                            || (item.StartsWith("{{") && item.EndsWith("}}")))
                            return 1;
                        else
                            return 2;

                    }
                }
            }
            for (int index = 0; index < linesParasPage.Length; index++)
            //foreach (string item in linesParasPage)
            {
                #region 舊式
                item = linesParasPage[index];
                //if (item == "") return;
                openBracketS = item.IndexOf("{{"); closeBracketS = item.IndexOf("}}");

                if (item == "}}<p>" || (closeBracketS == -1 && openBracketS == 0 && item.Length < 5))//《維基文庫》純注文空及其前一行
                {
                    i++;
                    if (item == "}}<p>") openNote = false; else openNote = true;
                }

                else if (i == 0 && xPage.IndexOf("}}") > -1 && xPage.IndexOf("}}") < (xPage.IndexOf("{{") == -1 ? xPage.Length : xPage.IndexOf("{{")) && xPage.IndexOf("}}") > e)
                { i++; openNote = true; }//第一段/行是純注文        
                else if (i == 0 && item.IndexOf("{{") == -1 && item.IndexOf("}}") == -1)
                {
                    string xx = linesParasPage[i + 1];
                    if (xx.IndexOf("}}") > -1 && xx.IndexOf("{{") == -1)//&& x.IndexOf("}}") > e)
                    { i++; openNote = true; }//第一段/行是純注文
                    else { i += 2; openNote = false; }//第一段/行是純正文
                }

                else if (i == 0 && ((closeBracketS > -1 && openBracketS > closeBracketS) ||
                    (openBracketS == -1 && closeBracketS > -1 && closeBracketS < item.Length - 2))) //第一行正、注夾雜
                {
                    if (openBracketS > 2)
                    {
                        i += 2;
                    }
                    else
                    {
                        if (openBracketS == -1) i += 2;
                        else if (openBracketS == 1)
                        {//目前分行分段於有標點者切割有誤差，權以此暫補丁
                            if (omitStr.IndexOf(item.Substring(0, 1)) == -1)
                            {
                                i += 2;
                            }
                            else i++;
                        }
                        else if (openBracketS == 2)
                        {
                            if (omitStr.IndexOf(item.Substring(0, 1)) == -1) i += 2;
                            else if (omitStr.IndexOf(item.Substring(1, 1)) == -1)
                            {//目前分行分段於有標點者切割有誤差，權以此暫補丁
                                i += 2;
                            }
                            else
                            {
                                i++;
                            }
                        }
                    }

                    if (item.LastIndexOf("}}") > item.LastIndexOf("{{"))
                        openNote = false;
                    else
                        openNote = true;
                }

                else if (openBracketS == 0 && closeBracketS == -1)//獨立注文（純注文）（開始）
                { i++; openNote = true; }
                else if (openBracketS == -1 && openNote)
                {//純注文（末截）
                    if (closeBracketS == item.Length - 2)//第2行/段純注文、獨立注文
                    { i++; openNote = false; }
                    else if (item.Length > 4)
                    {
                        if (item.Substring(item.Length - 5) == "}}<p>") { i++; openNote = false; }
                        else
                        {
                            if (closeBracketS == -1)
                            {
                                if (openNote)
                                    i++;
                                else
                                    i += 2;
                            }
                            else
                            {
                                i += 2;
                                openNote = false;
                            }

                        }
                    }
                }
                else if (openBracketS == -1 && closeBracketS > -1 && closeBracketS < item.Length - 2)
                {//正注夾雜注文結束
                    { i += 2; openNote = false; }
                }
                else if (openBracketS > -1 && item.IndexOf("{{", openBracketS + 2) > -1
                    && item.IndexOf("{{{", openBracketS + 2) != item.IndexOf("{{", openBracketS + 2))//正注夾雜
                {
                    i += 2;
                    if (item.LastIndexOf("}}") < item.LastIndexOf("{{")) openNote = true;
                    else openNote = false;
                }
                else if (openBracketS > -1 && closeBracketS > -1 && closeBracketS < item.Length - 2)
                {
                    if (openBracketS == 0 && (item.EndsWith("}}") || item.EndsWith("}}<p>") || item.EndsWith("}}。<p>"))
                        && ((item.IndexOf("{{", openBracketS + 2) == -1 && item.LastIndexOf("}}", closeBracketS) == -1)
                        || (item.IndexOf("{{", openBracketS + 2) == item.IndexOf("{{{", openBracketS + 2) &&
                                item.LastIndexOf("}}", closeBracketS) == item.LastIndexOf("}}}", closeBracketS))
                        ))//獨立注文
                    {
                        string preItem, preItem1;
                        if (index == 0) i += 2;
                        else if (index > 0)
                        {
                            preItem = linesParasPage[index - 1];
                            if (index > 1)
                                preItem1 = linesParasPage[index - 2];
                            else
                                preItem1 = null;
                            if (((preItem.IndexOf("{{") == -1 || (preItem.StartsWith("{{") && preItem.IndexOf("{{", 2) == -1))
                                        || (preItem.IndexOf("{{") > 0 && preItem.IndexOf("}}") > (preItem.IndexOf("{{"))))
                                && (preItem.EndsWith("}}") || preItem.EndsWith("}}<p>") || preItem.EndsWith("}}。<p>"))
                                && preItem.LastIndexOf("}}", preItem.LastIndexOf("}}")) == -1
                                //前一行是獨立注文
                                && ((item.EndsWith("}}<p>") || item.EndsWith("}}。<p>")) &&
                                    ((preItem.StartsWith("{{") && preItem.EndsWith("}}") && preItem.IndexOf("{{", 2) == -1)
                                    || preItem.Contains("{{") == false && preItem.EndsWith("}}")))
                                && (//本段是末有<p>的獨立注文，而前2段不是獨立注文
                                    !(((item.EndsWith("}}<p>") || item.EndsWith("}}。<p>")) && item.LastIndexOf("}}", item.LastIndexOf("}}")) == -1)
                                        && ((preItem1 != null && (preItem1.StartsWith("{{") && preItem1.IndexOf("{{", 2) == -1)
                                        && preItem.EndsWith("}}") && preItem.LastIndexOf("}}", preItem.Length - 2) == -1)
                                        ||
                                        (preItem1 != null && (!preItem1.Contains("{{") && !preItem1.Contains("}}"))
                                            && !preItem.Contains("{{") && preItem.EndsWith("}}") && preItem.LastIndexOf("}}", preItem.Length - 2) == -1)
                                        ))
                                    //或者本/行段與前行/段字數差不超過1
                                    || Math.Abs(CountWordsLenPerLinePara(item) - CountWordsLenPerLinePara(preItem)) <= 1)
                                )
                                i++;
                            else
                                i += 2;
                        }
                        openNote = false;
                    }
                    else//●●●●●●●●●●●●●●●
                    {//正注夾雜
                        i += 2;
                        if (item.LastIndexOf("}}") < item.LastIndexOf("{{")) openNote = true;
                        else openNote = false;
                    }

                }

                //無{{}}標記：
                else if (openBracketS == -1 && closeBracketS == -1)
                {
                    if (openNote == false)//《維基文庫》純正文
                        i += 2;
                    else //《維基文庫》純注文
                        i++;
                }

                //《維基文庫》正注文夾雜
                else if (openBracketS > 0)//正注夾雜
                {
                    if (openBracketS > 2)
                    {
                        i += 2;
                    }
                    else
                    {
                        if (openBracketS == 1)
                        {//目前分行分段於有標點者切割有誤差，權以此暫補丁
                            if (omitStr.IndexOf(item.Substring(0, 1)) == -1)
                            {
                                i += 2;
                            }
                            else i++;
                        }
                        if (openBracketS == 2)
                        {
                            if (omitStr.IndexOf(item.Substring(0, 1)) == -1) i += 2;
                            else if (omitStr.IndexOf(item.Substring(1, 1)) == -1)
                            {//目前分行分段於有標點者切割有誤差，權以此暫補丁
                                i += 2;
                            }
                            else
                            {
                                i++;
                            }
                        }
                    }
                    if (closeBracketS == -1) openNote = true;
                    else
                    {
                        if (item.LastIndexOf("}}") > item.LastIndexOf("{{"))
                            openNote = false;
                        else
                            openNote = true;
                    }
                }
                //else if (openBracketS > 0 && closeBracketS == -1) { i += 2; openNote = true; }
                else if (openBracketS == -1 && closeBracketS > -1 && closeBracketS < item.Length - 2) { i += 2; openNote = false; }
                /*
                if ((item.IndexOf("{{") == -1 && item.IndexOf("}}") == -1)//純正文
                    || item.IndexOf("{{") > 0 || item.IndexOf("}}") + 2 < item.Length)////正注文夾雜
                {
                    count += 2;
                }
                else if ((item.IndexOf("{{") > -1 && item.IndexOf("}}") > -1)//純注文
                    || (item.IndexOf("{{") > -1 && item.IndexOf("}}") == -1)
                    || (item.IndexOf("{{") == -1 && item.IndexOf("}}") > -1))
                {
                    count++;
                }
                */

                //else if 一行裡面不能有2個「{{」或「}}」，請完成這樣的判斷式 20250304 GitHub Copilot大菩薩：
                else if (item.Count(c => c == '{') > 2 || item.Count(c => c == '}') > 2)//只要不止1個大括號，不論是上或下大括弧，就必非獨立注文
                {
                    // 處理包含兩個或更多「{{」或「}}」的行
                    Debugger.Break();
                    i += 2;//●●●●●●●●●●●●●●●●●●●●●
                }
                else if (openBracketS == 0 && closeBracketS == item.Length - 2
                    && item.IndexOf("{{", openBracketS + 2) == -1 && item.LastIndexOf("}}", closeBracketS) == -1) { i++; openNote = false; }
                //●●●●●●●●●●●●●●●●●●●●●
                else if (openBracketS == -1 && closeBracketS == item.Length - 2
                    && item.IndexOf("{{", openBracketS + 2) == -1 && item.LastIndexOf("}}", closeBracketS) == -1) { i++; openNote = false; }
            }
            return i;
            #endregion

        }
        /// <summary>
        /// 作為 countLinesPerPage 之改良，但未成功！202050301
        /// </summary>
        /// <param name="xPage"></param>
        /// <returns></returns>
        int countLinesPerPage_NewTry(string xPage)
        {

            string[] linesParasPage = xPage.Split(Environment.NewLine.ToArray(), StringSplitOptions.RemoveEmptyEntries);

            int count = 0, index = -1;
            foreach (string item in linesParasPage)
            {
                index++;
                if (item == "|") count += 2;
                else
                {
                    if (item.IndexOf("{{") == -1 && item.IndexOf("}}") == -1)
                        count += 2;
                    else if (item.Contains("*"))
                        count += 2;
                    else if (item.IndexOf("{{") > 0)
                        count += 2;
                    else if (item.IndexOf("<p>") == -1 && item.LastIndexOf("}}") > -1 && item.LastIndexOf("}}") < item.Length - "}}".Length)
                        count += 2;
                    else if (item.LastIndexOf("}}") + 2 < item.IndexOf("<p>") || item.LastIndexOf("}}") + 2 < item.IndexOf("。<p>"))
                        count += 2;
                    else if (item.StartsWith("{{") && item.LastIndexOf("}}") == -1)
                        count++;
                    else
                    {
                        //落單的獨立注文，或夾注尾字數不足者
                        if ((item.StartsWith("{{") && item.EndsWith("}}"))
                            || (item.StartsWith("{{") && item.EndsWith("}}<p>"))
                            || (item.StartsWith("{{") && item.EndsWith("}}。<p>")))
                        {
                            string itemShort = item.Substring(2, item.IndexOf("}}") - 2);
                            if (!itemShort.Contains("{{") && !itemShort.Contains("}}"))
                            {
                                bool singleNoteLine = false;
                                if (index > 0)
                                {
                                    string preLine = linesParasPage[index - 1];
                                    string nextLine = string.Empty;

                                    if (index + 1 < linesParasPage.Length)
                                        nextLine = linesParasPage[index + 1];
                                    if (!preLine.Contains("{{") && !preLine.Contains("}}"))
                                        singleNoteLine = true;
                                    else if (preLine.EndsWith("}}"))
                                    {
                                        if (nextLine != string.Empty && !nextLine.StartsWith("{{"))
                                        {
                                            string preLineShort = preLine.Substring(2, preLine.IndexOf("}}") - 2);
                                            if (!(preLine.StartsWith("{{") && preLine.EndsWith("}}")
                                                && preLineShort.IndexOf("{{") == -1 && preLineShort.IndexOf("}}") == -1))
                                                singleNoteLine = true;
                                        }
                                        else if (new StringInfo(itemShort).LengthInTextElements == 1)
                                            singleNoteLine = true;
                                    }

                                    if (singleNoteLine)
                                        count += 2;
                                    else
                                        count++;
                                }
                                else
                                    count += 2;
                            }//以上落單獨立注文的處理
                             //獨立注文
                            else if ((item.StartsWith("{{") && item.IndexOf("}}") == -1)
                                || (item.EndsWith("}}") && item.IndexOf("{{") == -1)
                                || (item.EndsWith("}}<p>") && item.IndexOf("{{") == -1)
                                || (item.EndsWith("}}。<p>") && item.IndexOf("{{") == -1))
                                count++;
                            else
                                count += 2;

                        }
                    }
                }
            }
            return count;
        }

        /// <summary>
        /// 每頁行/段數。初始化（歸零）值為-1
        /// </summary>
        int linesParasPerPage = -1;
        /// <summary>
        /// 每行/段字數。初始化（歸零）值為-1
        /// </summary>
        int wordsPerLinePara = -1;
        internal static int CountNoteLen(string notePure)
        {//同時取商數與餘數 https://dotblogs.com.tw/abbee/2010/09/28/17943
            int l = new StringInfo(notePure).LengthInTextElements;
            int x = l / 2; ; //商數
            int y = l - (x * 2);//餘數
                                //return (((l + 1) % 2) == 1) ? ++l / 2 : l / 2;
            return y == 0 ? x : ++x;
        }

        /* https://ctext.org/library.pl?if=en&file=36096&page=99&editwiki=644293#editor (20250224 已測試） */
        /// <summary>
        /// 計算單行/段的字數
        /// 含空格空白；而標點符號、校按、標題標記、段落標記、分行標記均不計（即只計算原書內文之長度，以供比較行/段長度）
        /// 待除錯，可以此頁內容檢測（當已標記時再執行此函式會誤改動 20240905 
        /// </summary>
        /// <param name="xLinePara">要計算的行/段的文字字串</param>
        /// <returns></returns>
        internal static int CountWordsLenPerLinePara(string xLinePara)
        {
            //if (xLinePara.IndexOf("是歲復置函谷關") > -1)//just for debugging
            //    Debugger.Break();

            #region  清除{{{}}}內容不算入字數∵圖文對照頁面並不會顯示出來
            /* 20231102 Bing大菩薩：C#正則表達式：
             * …在C#中，您可以使用正則表達式來滿足您的需求。以下是一個範例程式碼，它將會找到「{{{」和「}}}」之間的所有文字並將其移除：…
             * …在這個程式碼中，我們使用了 Regex.Replace 方法來替換匹配到的部分。正則表達式 {{{.*?}}} 會匹配到「{{{」和「}}}」之間的所有文字（包含「{{{」和「}}}」）。請注意，我們在 .*? 中使用了 ? 來實現非貪婪匹配，這樣可以確保當有多組「{{{」和「}}}」時能夠正確地匹配。…
             */
            string pattern;
            if (xLinePara.IndexOf("{{{") > -1 || xLinePara.IndexOf("}}}") > -1)
            {
                pattern = "{{{.*?}}}";//Copilot大菩薩： 如果要使用貪婪匹配，只需要把正則表達式中的「?」去掉即可。原本的模式 ＝.*?＝ 中，「?」使得匹配變成非貪婪模式，而去掉它後，模式 ＝.*＝ 就變成了貪婪匹配。 20250224
                xLinePara = Regex.Replace(xLinePara, pattern, string.Empty);
            }
            #endregion
            #region 構字表達式「●＝＝」只保留「●」
            pattern = "＝.*?＝";//Copilot大菩薩： 非貪婪匹配（Lazy Matching）和貪婪匹配（Greedy Matching）是正則表達式中的兩種匹配模式。貪婪匹配會盡可能多地匹配字符，而非貪婪匹配則會盡可能少地匹配字符。
                              //例如，在字串 ＝abc＝def＝ 中，正則表達式 ＝.*＝（貪婪匹配）會匹配整個字串 ＝abc＝def＝，而 ＝.*?＝（非貪婪匹配）則會分別匹配 ＝abc＝ 和 ＝def＝。
            xLinePara = Regex.Replace(xLinePara, pattern, string.Empty);
            #endregion

            #region 語法標記不記（如插圖、專名標誌等語法）
            //待實作,碰到再說
            #endregion

            #region 標題標記、段落標記、分行標記均不計（即只計算原書內文之長度）

            #endregion
            pattern = "[*<p>|]";
            xLinePara = Regex.Replace(xLinePara, pattern, string.Empty);
            #region 標點符號不計
            //StringInfo seInfo = new StringInfo(se);
            foreach (var item in PunctuationsNum)
            {
                xLinePara = xLinePara.Replace(item.ToString(), "");
            }
            #endregion



            int openCurlybracketsPostion = xLinePara.IndexOf("{{"), closeCurlybracketsPostion = xLinePara.IndexOf("}}"),
                s = 0, countResult = 0;//, e = 0

            //if(openCurlybracketsPostion==-1 && closeCurlybracketsPostion>-1)
            //{
            //    xLinePara = "{{" + xLinePara;
            //    openCurlybracketsPostion = xLinePara.IndexOf("{{");
            //}

            string txt = "", note = "";//se = ""

            if (openCurlybracketsPostion == -1 && closeCurlybracketsPostion == -1)//純正文、純注文
                return new StringInfo(xLinePara).LengthInTextElements;
            //else if (openCurlybracketsPostion > -1 && closeCurlybracketsPostion > -1)
            else if (!(openCurlybracketsPostion == 0 && closeCurlybracketsPostion == xLinePara.Length - 2)
                && !(openCurlybracketsPostion == 0 && closeCurlybracketsPostion == -1)
                && !(openCurlybracketsPostion == -1 && closeCurlybracketsPostion == xLinePara.Length - 2)) //20250220
            {//兼具 {{、}} 正文、注文夾雜者
                while (true)
                {
                    //if (openCurlybracketsPostion == 0 && closeCurlybracketsPostion > openCurlybracketsPostion &&
                    //        xLinePara.IndexOf("{{", closeCurlybracketsPostion) == -1 &&
                    //        xLinePara.IndexOf("}}", closeCurlybracketsPostion + 2) == -1)
                    //{// like this :     {{……}}……
                    //    return new StringInfo(xLinePara.Substring(closeCurlybracketsPostion + 2)).LengthInTextElements +
                    //            countNoteLen(xLinePara.Substring(openCurlybracketsPostion + 2, closeCurlybracketsPostion - 2));
                    //}
                    //else 
                    if (closeCurlybracketsPostion > -1 &&
                        (openCurlybracketsPostion > closeCurlybracketsPostion) || openCurlybracketsPostion == -1)
                    {//先出現 }} 的話
                     //s = closeCurlybracketsPostion + 2;
                     //   countResult += new StringInfo(xLinePara.Substring(0, closeCurlybracketsPostion)).LengthInTextElements;
                        countResult += CountNoteLen(xLinePara.Substring(0, closeCurlybracketsPostion));
                        //closeCurlybracketsPostion = xLinePara.IndexOf("}}", closeCurlybracketsPostion + 2);
                    }
                    else if (closeCurlybracketsPostion > -1)//&& openCurlybracketsPostion>-1
                    {
                        txt = xLinePara.Substring(s, openCurlybracketsPostion - s);
                        countResult += new StringInfo(txt).LengthInTextElements;
                        note = xLinePara.Substring(openCurlybracketsPostion + 2, closeCurlybracketsPostion - (openCurlybracketsPostion + 2));
                        countResult += CountNoteLen(note);
                    }
                    else if (closeCurlybracketsPostion == -1 && openCurlybracketsPostion > -1)
                    {
                        txt = xLinePara.Substring(s, openCurlybracketsPostion);
                        countResult += new StringInfo(txt).LengthInTextElements;
                        note = xLinePara.Substring(openCurlybracketsPostion + 2);
                        countResult += CountNoteLen(note);
                        break;
                    }
                    else {; }
                    s = closeCurlybracketsPostion + 2;
                    openCurlybracketsPostion = xLinePara.IndexOf("{{", closeCurlybracketsPostion);
                    if (openCurlybracketsPostion == -1)
                    {
                        if (closeCurlybracketsPostion + 2 > xLinePara.Length) txt = "";
                        else txt = xLinePara.Substring(closeCurlybracketsPostion + 2);
                        return countResult += new StringInfo(txt).LengthInTextElements;
                    }
                    closeCurlybracketsPostion = xLinePara.IndexOf("}}", closeCurlybracketsPostion + 2);
                    if (closeCurlybracketsPostion == -1)
                    {
                        txt = xLinePara.Substring(s, openCurlybracketsPostion - s);
                        countResult += new StringInfo(txt).LengthInTextElements;
                        note = xLinePara.Substring(openCurlybracketsPostion + 2);
                        return countResult += CountNoteLen(note);
                    }
                    //if (!(openCurlybracketsPostion > -1)) break;
                    if (openCurlybracketsPostion < 0) break;
                }

                return countResult;

            }
            else if (openCurlybracketsPostion > 0 && closeCurlybracketsPostion == -1)
            {//只有 {{ 雜正文                
                return new StringInfo(xLinePara.Substring(0, openCurlybracketsPostion)).LengthInTextElements +
                        CountNoteLen(xLinePara.Substring(openCurlybracketsPostion + 2));
            }
            else if (closeCurlybracketsPostion < xLinePara.Length - 2 && openCurlybracketsPostion == -1)
            {//只有 }} 雜正文
                return CountNoteLen(xLinePara.Substring(0, closeCurlybracketsPostion)) +
                    new StringInfo(xLinePara.Substring(closeCurlybracketsPostion + 2)).LengthInTextElements;
            }
            else//純注文（獨立注文，頭為「{{」，尾為「}}」）20250220
                return new StringInfo(xLinePara.Replace("{{", string.Empty).Replace("}}", string.Empty)).LengthInTextElements;

        }

        /// <summary>
        /// 快速取得字串前面連續空格數的方法        
        /// Gemini大菩薩： 更有效率的做法： 如果您需要多次執行此操作，可以預先編譯正則表達式以提高效率：https://g.co/gemini/share/9ff15700e2b1 20250204
        /// </summary>
        internal static readonly Regex _leadingSpacesRegex = new Regex(@"^\s+");
        /* 然後在需要時使用：
            Match match = _leadingSpacesRegex.Match(str);
            這樣可以避免每次都重新編譯正則表達式。
         */

        /// <summary>
        /// 若沒有用●的長度來指定每行字數，則根據第一段長來自動將故短的行尾，標上段落標記<p>
        /// 按下 Scroll Lock 將字數較少的行/段落尾末標上「<p>」符號
        /// </summary>
        void paragraphMarkingAccordingFirstOneLineLength()
        {
            bool topmost = TopMost, eventEnabled = _eventsEnabled;
            TopMost = false;
            if (textBox1.TextLength > 2 && textBox1.Text.Substring(textBox1.TextLength - 2, 2) != Environment.NewLine)
            {
                if (eventEnabled)
                    PauseEvents();
                textBox1.Text += Environment.NewLine;
                if (eventEnabled)
                    ResumeEvents();
            }
            undoRecord();

            string x = textBox1.Text;
            #region 直接取代
            replaceXdirectly(ref x, string.Empty, false);
            #endregion

            //l 在後面作為正常的行/段長度（含幾個漢字中文字）的儲存變量
            int s = 0, l, e = textBox1.Text.IndexOf(Environment.NewLine); if (e < 0) return;
            //PauseEvents();
            int rs = textBox1.SelectionStart, rl = textBox1.SelectionLength;
            string se = textBox1.Text.Substring(s, e - s);//取得s(start)和e(end)之間的字串，就叫se
                                                          //int l = new StringInfo(se).LengthInTextElements;
            if (wordsPerLinePara == -1 && new StringInfo(se).LengthInTextElements < 10)
            {
                MessageBoxShowOKExclamationDefaultDesktopOnly("請先在第1行指定正常行長再繼續！");
                return;
            }
            if (se.Contains("●") && se.Replace("●", "") == "")
            {
                l = se.Length;
                wordsPerLinePara = l;
                NormalLineParaLength = wordsPerLinePara;
            }
            else
            {
                l = wordsPerLinePara != -1 ? wordsPerLinePara : CountWordsLenPerLinePara(se);
                if (wordsPerLinePara == 0)
                {
                    Debugger.Break();
                    wordsPerLinePara = CountWordsLenPerLinePara(se);
                }
                //if (NormalLineParaLength == 0) NormalLineParaLength = wordsPerLinePara;
                if (normalLineParaLength != wordsPerLinePara)
                    NormalLineParaLength = wordsPerLinePara;
            }

            bool ev = _eventsEnabled;
            if (se.Replace("●", "") == "")
            {
                if (ev) _eventsEnabled = false;
                textBox1.Text = textBox1.Text.Substring(e + 2);//●●●●●●●●乃作為權訂每行字數之參考，故可刪去
                _eventsEnabled = ev;                                                                   //if (CountWordsLenPerLinePara(se) == wordsPerLinePara && se.Replace("●", "") == "") textBox1.Text = textBox1.Text.Substring(e + 2);
            }
            else
            {//如果據以判斷的第一行不是用●●●●●●●●●●長度來判斷行/段長的話，亦清除此第1行 20250109
                if (x != textBox1.Text) x = textBox1.Text;
                if (x.IndexOf(Environment.NewLine) > -1 && autoPaste2QuickEdit && x.Length > 1100)
                {
                    if (x.Substring(x.IndexOf(Environment.NewLine) + Environment.NewLine.Length, 1) == "*" && x.Substring(0, x.IndexOf(Environment.NewLine)).IndexOf("*") == -1)
                    {
                        if (ev) _eventsEnabled = false;
                        textBox1.Text = x.Substring(x.IndexOf(Environment.NewLine) + Environment.NewLine.Length);
                        _eventsEnabled = ev;
                    }
                }
            }

            // undoRecord(); //前已有
            stopUndoRec = true; PauseEvents();

            //string p = "<p>";
            string p = keyinTextMode && !ocrTextMode ? "。<p>" : "<p>";

            if (wordsPerLinePara == -1)
            {
                wordsPerLinePara = l;//l作為正常的行/段長度（含幾個漢字中文字）的儲存變量
                NormalLineParaLength = wordsPerLinePara;
            }
            else
            {
                if (se.IndexOf("<p>") == -1 && se.IndexOf("*") == -1 && CountWordsLenPerLinePara(se) < wordsPerLinePara)
                {
                    textBox1.Text = textBox1.Text.Substring(0, e) + p//"<p>"
                        + textBox1.Text.Substring(e);
                    //e += 3;//"<p>".length
                    e += p.Length;
                }
            }
            bool topLine = TopLine;//抬頭？
            ado.Connection cnt = new ado.Connection(); ado.Recordset rst = new ado.Recordset();
            if (topLine)
            {
                Mdb.openDatabase("查字.mdb", ref cnt);
                rst.Open("select * from 每行字數判斷用 where condition=0", cnt, ado.CursorTypeEnum.adOpenKeyset, ado.LockTypeEnum.adLockReadOnly);
            }

            ////第一行/段
            if (x != textBox1.Text) x = textBox1.Text;
            //string currentLineText = GetLineTxt(x, s);
            //string currentLineTextMissCharactersConverted=CnText.ConvertSpace2BlankAfterMissingCharacter(currentLineText);
            //if(currentLineText!=currentLineTextMissCharactersConverted)
            //{
            //    textBox1.Select(s, e);
            //    textBox1.SelectedText = currentLineTextMissCharactersConverted;
            //    e += currentLineTextMissCharactersConverted.Length - currentLineText.Length;
            //}

            //開始定位
            while (e > -1)
            {


                s = e + 2;
                e = textBox1.Text.IndexOf(Environment.NewLine, s);
                if (e == -1) break;
                se = textBox1.Text.Substring(s, e - s);//本行/段文字==currentLineText
                                                       //foreach (var item in punctuations)
                                                       //{
                                                       //    se = se.Replace(item.ToString(), "");
                                                       //}

                string currentLineText = se;//GetLineTxt(x, s);
                string currentLineTextMissCharactersConverted = CnText.ConvertSpace2BlankAfterMissingCharacter(currentLineText);
                if (currentLineText != currentLineTextMissCharactersConverted)
                {
                    textBox1.Select(s, e - s);
                    textBox1.SelectedText = currentLineTextMissCharactersConverted;
                    e += currentLineTextMissCharactersConverted.Length - currentLineText.Length;
                }

                //if (se.IndexOf("跡驗父故邪夏侯方") > -1)//just for test
                //    Debugger.Break();

                //if (se.IndexOf("吾山}}") > -1) Debugger.Break();

                if (se != string.Empty)//se 正要處理的這一行/段的內容string
                {
                    string textBox1Text = textBox1.Text;//l作為正常的行/段長度（含幾個漢字中文字）的儲存變量
                    if (CountWordsLenPerLinePara(se) < l//長度小於常規
                        && (!(se.StartsWith("{{") && se.EndsWith("}}") == false)//而且：不是有「{{」而無「}}」的獨立小注
                            || (se.IndexOf("{{") > 0 || se.IndexOf("}}") < se.Length - "}}".Length)))//或者是正注夾雜
                    {
                        //if (((se.IndexOf("{{") == -1 && se.IndexOf("}}") == -1)
                        //    || (se.IndexOf("{{") == -1 && se.IndexOf("}}") > -1)
                        //    || (se.IndexOf("{{") > 0 && se.IndexOf("}}") > -1)) //「{{」不能是開頭
                        //    && se.IndexOf("<p>") == -1)
                        if (se.IndexOf("*") == -1 && se.IndexOf("<p>") == -1 && se.IndexOf("|") == -1
                            && !(se.IndexOf("{{") == 0 && se.IndexOf("}}") == -1))
                        //if (se.Substring(se.Length - 3, 3)!="<p>")
                        {
                            //string tx = textBox1.Text;
                            if (textBox1Text.IndexOf(Environment.NewLine, e + 2) > -1)//後面尚有行/段
                            {
                                //定位要插入段落標記之處
                                textBox1.Select(e, 0);

                                //是否有抬頭格式？
                                if (topLine)
                                {
                                    if (isShortLine(textBox1Text.Substring(e + 2, textBox1Text.IndexOf(Environment.NewLine, e + 2) - e - 2),
                                        textBox1Text.Substring(textBox1Text.LastIndexOf(Environment.NewLine, e) + 2, e - textBox1Text.LastIndexOf(Environment.NewLine, e) - 2)
                                        , cnt, rst))
                                    {
                                        textBox1.SelectedText = p;//"<p>";
                                        e += p.Length;
                                        if ((int)rst.AbsolutePosition > 1) rst.MoveFirst();
                                    }
                                    else
                                    {
                                        textBox1.SelectedText = "|";
                                        e++;
                                        if ((int)rst.AbsolutePosition > 1) rst.MoveFirst();
                                    }
                                }
                                else
                                {
                                    //取得插入點標題範圍
                                    //Paragraph para = _document.GetCurrentParagraph();
                                    //int asteriskPos = _document.Text.LastIndexOf("*", para.Start), pMarkPos = _document.Text.IndexOf("<p>", para.End);
                                    int asteriskPos = textBox1.Text.LastIndexOf("*", textBox1.SelectionStart), titleCloseMark = textBox1.Text.LastIndexOf("<p>", textBox1.SelectionStart), pMarkPos = textBox1.Text.IndexOf("<p>", textBox1.SelectionStart);
                                    if (asteriskPos > -1 && pMarkPos > -1)//如果在標題範圍內
                                    {
                                        if (titleCloseMark > -1)
                                        {
                                            if (titleCloseMark > asteriskPos && titleCloseMark < textBox1.SelectionStart)
                                            {
                                                #region 取得下一行/段來判斷20250221修訂
                                                //se = textBox1.Text.Substring(s, e - s);//本行/段文字
                                                string seNext = textBox1.Text.Substring(e + Environment.NewLine.Length,
                                                    textBox1Text.IndexOf(Environment.NewLine, e + 2) - (e + Environment.NewLine.Length));

                                                #endregion
                                                if (CountWordsLenPerLinePara(se) < l//l作為正常的行/段長度（含幾個漢字中文字）的儲存變量
                                                    && (se.StartsWith("{{") && se.EndsWith("}}")
                                                    && CountWordsLenPerLinePara(seNext) < l
                                                    && (seNext.StartsWith("{{")
                                                        && (seNext.EndsWith("}}") || seNext.EndsWith("}}" + p)))))//長度小於常規
                                                {
                                                    /*如果獨立注文較常行/段長短，且後面也是較常行/段短的獨立注文的話
                                                     就在此行/段末不加<p> 20250221修訂 感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主*/
                                                }
                                                else
                                                {
                                                    #region 原式 20250221修訂
                                                    textBox1.SelectedText = p;//"<p>";
                                                    e += p.Length;
                                                    #endregion
                                                }
                                            }
                                        }
                                        else
                                        {
                                            //Range rangeTitle = _document.Range(asteriskPos, pMarkPos);//取得插入點標題範圍
                                            //if (!(asteriskPos < para.Start && pMarkPos > para.End))
                                            if (!(asteriskPos < textBox1.SelectionStart && pMarkPos > textBox1.SelectionStart))
                                            {//所在處不是在標題內才執行段落標記，標題只能有其結尾1個段落標記<p>，以*開頭Open，以<p>結尾Close 20250206
                                                textBox1.SelectedText = p;//"<p>";
                                                e += p.Length;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        textBox1.SelectedText = p;//"<p>";
                                        e += p.Length;
                                    }
                                }
                            }
                            else
                            {
                                textBox1.Select(e, 0);
                                textBox1.SelectedText = p;//"<p>";
                                e += p.Length;
                                if (topLine)
                                {
                                    if ((int)rst.AbsolutePosition > 1) rst.MoveFirst();
                                }
                            }

                        }
                        else
                        {//如果在標題中則要將空格轉為空白
                            if (se.IndexOf("*") > -1)
                            {
                                //如果在標題中，且下一行/段是空格開頭且又多於標題題首空格2格的話 20250204
                                string next_se = textBox1.Text.IndexOf(Environment.NewLine, e + 1) < 0 ? string.Empty :
                                    textBox1.Text.Substring(e + Environment.NewLine.Length,
                                    (e + Environment.NewLine.Length + textBox1.Text.IndexOf(Environment.NewLine, e + 1) - e + Environment.NewLine.Length) > textBox1.TextLength ?
                                    (textBox1.TextLength - (e + Environment.NewLine.Length)) :
                                    textBox1.Text.IndexOf(Environment.NewLine, e + 1) - e + Environment.NewLine.Length);
                                //if (nextse.IndexOf("{{")==-1 &&)
                                //20250204 creedit with Gemini大菩薩 https://g.co/gemini/share/eb6c0ed286c6
                                // 使用正則表達式匹配開頭的連續空格
                                // 然後在需要時使用：
                                //Match match = _leadingSpacesRegex.Match(str);
                                //Match match = Regex.Match(next_se, @"^\s+"), match_se = Regex.Match(se, @"^\s+");
                                Match match = Form1._leadingSpacesRegex.Match(next_se), match_se = _leadingSpacesRegex.Match(se);
                                int spaceCount_nextse = match.Success ? match.Value.Length : 0,
                                    spaceCount_se = match_se.Success ? match_se.Value.Length : 0;
                                if (spaceCount_nextse - spaceCount_se > 3)//通常標題顯首前之空格與其後行/段縮排之空格應該不會超過3個
                                {
                                    textBox1.Select(e + Environment.NewLine.Length, spaceCount_nextse);
                                    StringBuilder sb = new StringBuilder();
                                    for (int iblk = 0; iblk < spaceCount_nextse; iblk++)
                                    {
                                        sb.Append("􏿽");
                                    }
                                    textBox1.SelectedText = sb.ToString();
                                }

                            }
                        }

                    }
                    else//長度不小於常規 l
                    {
                        if (Indents)//書內含縮排格式
                        {
                            //如果本行沒有段末標記
                            if (se.IndexOf("<p>") == -1 && se.IndexOf("|") == -1 && se.IndexOf("*") == -1)
                            //&& !(se.IndexOf("{{") == 0 && se.IndexOf("}}") == -1))
                            {
                                //如果本行有縮排
                                if ("　􏿽".IndexOf(se.Substring(0, 1)) > -1 ||
                                    (se.Substring(0, 2) == "{{" && "　􏿽".IndexOf(se.Substring(2, 1)) > -1))
                                {
                                    int en = textBox1Text.IndexOf(Environment.NewLine, e + 2); int spaceCnt, isp = 0;
                                    while ("　􏿽".IndexOf(se.Substring(++isp, 1), StringComparison.Ordinal) > -1)
                                    {

                                    }
                                    //isp--;
                                    spaceCnt = new StringInfo(se.Substring(0, isp)).LengthInTextElements;
                                    if (en > -1)
                                    {
                                        if ("　􏿽".IndexOf(textBox1Text.Substring(e + 2, 1)) == -1 ||//如果下一行/段不是縮排而是頂格、頂行
                                            (textBox1Text.Substring(e + 2, 2) == "{{" && "　􏿽".IndexOf(se.Substring(2, 1)) == -1))
                                        {
                                            textBox1.Select(e, 0);
                                            textBox1.SelectedText = p;//"<p>";
                                            e += 3;
                                        }
                                        else
                                        {//如果下一行/段再縮排（且不是注文）
                                            if ("　􏿽".IndexOf(textBox1Text.Substring(e + 2, 1)) > -1)//&& se.IndexOf("*") == -1)//&& tx.Substring(s, e).IndexOf("*") == -1)
                                            {
                                                isp = 0;
                                                while ("　􏿽".IndexOf(textBox1Text.Substring(e + 2 + (++isp), 1), StringComparison.Ordinal) > -1)//有「�」時會影響判斷
                                                {
                                                    //取得縮排數
                                                }
                                                //if (new StringInfo(tx.Substring(e + 2, --isp)).LengthInTextElements > spaceCnt)
                                                if (new StringInfo(textBox1Text.Substring(e + 2, isp)).LengthInTextElements > spaceCnt)
                                                {
                                                    textBox1.Select(e, 0);
                                                    textBox1.SelectedText = p;//"<p>";
                                                    e += p.Length;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            //最後一行
            string lastLineText = GetLineText_WithoutPunctuation(textBox1.Text, s);
            if (new StringInfo(lastLineText).LengthInTextElements < wordsPerLinePara && lastLineText.IndexOf("<p>") == -1)
                textBox1.Text += p;
            stopUndoRec = false; ResumeEvents();
            replaceBlank_ifNOTTitleAndAfterparagraphMark();
            fillSpace_to_PinchNote_in_LineStart();
            if (EventsEnabled) PauseEvents();
            stopUndoRec = true; PauseEvents();


            #region 最後一行處理

            if (textBox1.TextLength > 1
                && textBox1.Text.Substring(textBox1.TextLength - Environment.NewLine.Length, Environment.NewLine.Length) != Environment.NewLine)
            {
                se = GetLineText_WithoutPunctuation(textBox1.Text, textBox1.Text.LastIndexOf(Environment.NewLine)
                    + Environment.NewLine.Length);

                //if (se.IndexOf("跡驗父故邪夏侯方") > -1)//just for test
                //    Debugger.Break();
                e = CountWordsLenPerLinePara(se);
                if (e < wordsPerLinePara)
                    if (se.StartsWith("{{") && se.EndsWith("}}"))
                    {
                        if (e < wordsPerLinePara / 2)
                            if (textBox1.Text.Substring(textBox1.TextLength - 1, 1) != ">") textBox1.Text += "<p>";
                    }
                    else
                        if (textBox1.Text.Substring(textBox1.TextLength - 1, 1) != ">") textBox1.Text += "。<p>";
            }
            #endregion

            if (eventEnabled) PauseEvents();
            textBox1.Text = textBox1.Text.Replace("\r\n　<p>", "\r\n|");
            if (textBox1.TextLength > Environment.NewLine.Length + p.Length &&
                textBox1.Text.Substring(textBox1.TextLength - (Environment.NewLine.Length + p.Length), Environment.NewLine.Length + p.Length) == Environment.NewLine + p)
                textBox1.Text = textBox1.Text.Substring(0, textBox1.TextLength - (Environment.NewLine.Length + p.Length));
            if (textBox1.TextLength > p.Length
                && textBox1.Text.Substring(textBox1.TextLength - p.Length, p.Length) == "。<p>"
                && textBox1.Text.Substring(textBox1.Text.LastIndexOf(Environment.NewLine) + Environment.NewLine.Length
                    , textBox1.TextLength - (textBox1.Text.LastIndexOf(Environment.NewLine) + Environment.NewLine.Length)).Replace("　", "").Replace("􏿽", "") == p)
                textBox1.Text = textBox1.Text.Substring(0, textBox1.Text.LastIndexOf(Environment.NewLine));
            if (eventEnabled) ResumeEvents();

            #region 如果是《欽定四庫全書》的文本
            if (textBox1.Text.IndexOf(Environment.NewLine) > -1 && textBox1.Text.Substring(0, textBox1.Text.IndexOf(Environment.NewLine)).Contains("欽定四庫全書"))
            {

                if (textBox1.Text.Substring(textBox1.TextLength - 3, 3) != "<p>")
                    textBox1.Text += "<p>";

                #region 第2行/段標題的自動格式化
                string secondLineText = GetLineText(textBox1.Text, textBox1.Text.IndexOf(Environment.NewLine) + 3, out int lineS, out int lineL);
                if (!secondLineText.Contains("*"))
                {
                    secondLineText = secondLineText.Replace("􏿽", "　");
                    int leadingSpacesLength = _leadingSpacesRegex.Match(secondLineText).Value.Length;
                    //if (!secondLineText.Substring(_leadingSpacesRegex.Match(secondLineText).Value.Length).Contains("　"))
                    if (!secondLineText.Substring(leadingSpacesLength).StartsWith("　"))
                    {
                        //將第2行/段的文本做標題標記（如果是《欽定四庫全書》的文本）
                        int start = textBox1.SelectionStart, len = textBox1.SelectionLength;
                        textBox1.Select(textBox1.Text.IndexOf(Environment.NewLine) + 2, leadingSpacesLength * 2);
                        string blank2Space = textBox1.SelectedText.Replace("􏿽", "　");
                        textBox1.SelectedText = blank2Space;
                        secondLineText = GetLineText(textBox1.Text, lineS, out lineS, out lineL);
                        int spsPos = secondLineText.IndexOf("　", leadingSpacesLength);//如果標題前置空格後面還有空格

                        if (spsPos > -1)
                            //textBox1.Select(lineS + leadingSpacesLength, spsPos - (lineS + leadingSpacesLength));
                            //textBox1.Select(lineS + leadingSpacesLength + leadingSpacesLength, spsPos - (lineS + leadingSpacesLength) - leadingSpacesLength);
                            textBox1.Select(lineS + leadingSpacesLength, spsPos - leadingSpacesLength);
                        else
                            //textBox1.Select(lineS, lineL);
                            textBox1.Select(lineS + leadingSpacesLength, lineL - leadingSpacesLength);
                        titleMarkCode();

                        #region 標題標記完後若其行/段後面還有文字，則予以格式化（全形空格轉空白）
                        if (CountWordsinDomain("<p>", GetLineText(textBox1.Text, textBox1.SelectionStart)) > 1)
                        {
                            Paragraph paragraph = new Range(_document, textBox1.SelectionStart, textBox1.SelectionStart, _document.Content).GetCurrentParagraph();
                            textBox1.Select(paragraph.Start, paragraph.End - paragraph.Start);
                            undoRecord(); stopUndoRec = true; PauseEvents();
                            textBox1.SelectedText = CnText.ReplaceFullWidthSpace_withBlank(textBox1.SelectedText);
                            undoRecord(); stopUndoRec = false; ResumeEvents();
                        }
                        #endregion 標題標記完後若其行/段後面還有文字，則予以格式化（全形空格轉空白）

                        textBox1.Select(start, len);
                    }
                }
                #endregion 第2行/段標題的自動格式化

                //if (_document.Text.IndexOf("􏿽<p>") > -1)
                //{
                x = textBox1.Text;
                replaceXdirectly(ref x, string.Empty, false);
                textBox1.Text = x;
                //}

            }
            #endregion

            playSound(soundLike.over);
            if (topLine) { rst.Close(); cnt.Close(); rst = null; cnt = null; }
            //if (keyinTextMode)
            if (!ocrTextMode)
            {
                TopMost = true;//方便檢視要清除的誤標「<p>」
                clearParagraphMarkersBetweenPairsBrackets();
                clearParagraphMarkersInsidePairsBrackets();
                TopMost = topmost;
                lastLineText = textBox1.Text;//借用此變數
                CnText.FormalizeText(ref lastLineText);
                textBox1.Text = lastLineText;
                movePeriodsToFrontofBlank();
                textBox1.Select(textBox1.TextLength, 0);

            }
            else
                textBox1.Select(rs, rl);


            #region 檢查縮排而等行長的分段處 20250424 如 https://ctext.org/library.pl?if=en&res=5677
            x = textBox1.Text; s = 0;
            s = x.IndexOf(Environment.NewLine, s);
            while (s > -1)
            {
                string currLineX = GetLineText(x, s);
                //本行與正常行等長
                if (CountWordsLenPerLinePara(currLineX) == wordsPerLinePara)
                {//本行有縮排
                    if (_leadingSpacesRegex.Match(currLineX).Value.Length > 0)
                    {
                        int sCurr = s;
                        s = x.IndexOf(Environment.NewLine, s + 1);
                        if (s > -1)
                        {
                            string nextLineX = GetLineText(x, s);
                            //下行與正常行等長且無縮排 或者下行文字長度為0（如為「|」）
                            if (((CountWordsLenPerLinePara(nextLineX) == wordsPerLinePara &&
                                    _leadingSpacesRegex.Match(nextLineX).Value.Length == 0)
                                || (CountWordsLenPerLinePara(nextLineX) == 0))
                                && (!nextLineX.StartsWith("{{") && !currLineX.EndsWith("}}"))
                                )
                            {
                                if (sCurr - 3 > -1 && x.Substring(sCurr - 3, 3) != "<p>")
                                {
                                    x = x.Substring(0, sCurr) + "<p>" + x.Substring(sCurr);
                                    s += "<p>".Length;
                                }
                            }
                        }
                    }
                    else
                    {//本行沒有縮排
                        int sCurr = s;
                        s = x.IndexOf(Environment.NewLine, s + 1);
                        if (s > -1)
                        {
                            string nextLineX = GetLineText(x, s);
                            //下行與正常行等長且無縮排 或者下行文字長度為0（如為「|」）
                            if (CountWordsLenPerLinePara(nextLineX) == 0
                                && (!nextLineX.StartsWith("{{") && !currLineX.EndsWith("}}")))
                            {
                                if (sCurr - 3 > -1 && x.Substring(sCurr - 3, 3) != "<p>")
                                {
                                    x = x.Substring(0, sCurr) + "<p>" + x.Substring(sCurr);
                                    s += "<p>".Length;
                                }
                            }
                        }
                    }
                }
                else
                    s = x.IndexOf(Environment.NewLine, s + 1);


            }
            if (x != textBox1.Text)
            {
                s = textBox1.SelectionStart;
                textBox1.Text = x;
                textBox1.SelectionStart = s;
            }
            #endregion 檢查縮排而等行長的分段處 20250424 如 https://ctext.org/library.pl?if=en&res=5677

            textBox1.ScrollToCaret();
            TopMost = topmost; stopUndoRec = false; ResumeEvents();
        }
        /// <summary>
        /// 清除成對的大括號間的分段標記
        /// 小注間的分段標記檢查與清除
        /// </summary>
        private void clearParagraphMarkersBetweenPairsBrackets()
        {
            stopUndoRec = true; PauseEvents();
            string x = textBox1.Text; int ip = -1;
            //要檢查的字串
            string[] paragraphMarkIn = { "}}。<p>\r\n{{", "}}<p>\r\n{{" };
            foreach (var item in paragraphMarkIn)
            {
                int s = ip + 1; ip = textBox1.Text.IndexOf(item, s); if (ip == -1) continue;
                string nextLineTxt = GetNextLineText_IncludingMarkers(x, ip);
                //textBox1.Text = textBox1.Text.Replace("。<p>\r\n{{", "\r\n{{");//此不宜逕行取代，參見《札迻》版式，故今以下式取代，半自動手動校勘 20231114 感恩感恩　讚歎讚歎　南無阿彌陀佛

                //參閱 paragraphMarkAccordingFirstOne 的這行 ：
                //string p = keyinTextMode && !ocrTextMode ? "。<p>" : "<p>"; 
                //兩邊當同步
                string punctuation = (keyinTextMode && !ocrTextMode) ? PunctuationsNum.Replace("。", string.Empty) : PunctuationsNum;

                while (ip > -1)
                {
                    //下一行若是如「{{九百五十三}}湯浚對鬯酒」則不清除
                    int noteMarkClosePos = nextLineTxt.IndexOf("}}");//小注結束標記的位置
                                                                     //Debugger.Break();
                    if (noteMarkClosePos > -1 &&//如下
                        ((noteMarkClosePos == nextLineTxt.Length - 2)// 2 = "}}".Length

                        ||//「}}」後不是中文、不是Surrogate字符、且不是標點符號

                        (noteMarkClosePos + 2 < nextLineTxt.Length//下一行下大括號位置如果不是行末
                            &&
                            /*1.：下一段全是注文，但尾端有「<p>」標記，如：       ……}}<p>
                                                                        {{卷上}}<p> 
                                                                即下一行的尾端一定是「}}<p>」*/
                            (
                                (nextLineTxt.Length > 5 && nextLineTxt.Substring(nextLineTxt.Length - 5) == "}}<p>")// 5 = ""}}<p>"".Length

                            ||/* 2:             }}<p>
                                            {{卷上}}君臣取象…… 
                                            通常正常長度時自動標上段落標記的 paragraphMarkAccordingFirstOne 方法是不會在「}}」後標上「<p>」 
                                            然非正常長度、每行長度不一或常變化時，則例外，須另想辦法或人工處理 20240906 */

                                (!IsCJKorSurrogate(nextLineTxt.Substring(noteMarkClosePos + 2, 1))//下大括號後的第1個字不是中文或 surrogate
                                                                                                  //而且下大括弧後的第1個字如果是中文，且長度要小於4（個字）
                                    || (IsCJKorSurrogate(nextLineTxt.Substring(noteMarkClosePos + 2, 1)) && new StringInfo(CnText.RemovePunctuationsNum(nextLineTxt.Substring(noteMarkClosePos + 2)).Replace("<p>", string.Empty).Replace("|", string.Empty)).LengthInTextElements < 4)
                                )
                            )

                            && punctuation.IndexOf(nextLineTxt.Substring(noteMarkClosePos + 2, 1)) == -1


                         )
                         )
                        )
                    //如果需要略過}}後是句號的判斷，如 「……}}。<p>」，則可用以下此行替換上行
                    //&& PunctuationsNum.Replace("。", string.Empty).IndexOf(nextLineTxt.Substring(noteMarkClosePos + 2, 1)) == -1)))))

                    //(char.IsHighSurrogate(nextLineTxt.Substring(noteMarkClosePos + 2, 1).ToCharArray()[0])||
                    //IsChineseString(nextLineTxt.Substring(noteMarkClosePos+2,1))))))
                    {
                        textBox1.Select(ip, item.Length);
                        DialogResult dr = DialogResult.No;
                        //if (Math.Abs(CountWordsLenPerLinePara(GetLineText(textBox1.Text, ip + item.Length)) - CountWordsLenPerLinePara(GetLineText(textBox1.Text, ip))) <= 1)
                        if (Math.Abs(CountWordsLenPerLinePara(GetNextLineText_IncludingMarkers(textBox1.Text, ip)) - CountWordsLenPerLinePara(GetLineText(textBox1.Text, ip))) <= 1)
                            dr = DialogResult.Yes;
                        if (dr != DialogResult.Yes)
                        {
                            //Refresh();
                            ScrollControlIntoView(textBox1);
                            textBox1.ScrollToCaret();
                            //if (MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否要清除注文間的段落標記<p>？"
                            //    + Environment.NewLine + Environment.NewLine
                            //    + textBox1.Text.Substring(textBox1.SelectionStart - 10 > -1 ? textBox1.SelectionStart - 10 : 0, 10)
                            //    + textBox1.SelectedText
                            //    + textBox1.Text.Substring(textBox1.SelectionStart + textBox1.SelectionLength <= textBox1.TextLength ? textBox1.SelectionStart + textBox1.SelectionLength : textBox1.TextLength
                            //                                , textBox1.SelectionStart + textBox1.SelectionLength + 10 <= textBox1.TextLength ? 10 : textBox1.TextLength - (textBox1.SelectionStart + textBox1.SelectionLength))) == DialogResult.OK)

                            dr = MessageBox.Show("是否要清除注文間的段落標記<p>？"
                                + Environment.NewLine + Environment.NewLine
                                + textBox1.Text.Substring(textBox1.SelectionStart - 10 > -1 ? textBox1.SelectionStart - 10 : 0, 10)
                                + textBox1.SelectedText
                                + textBox1.Text.Substring(textBox1.SelectionStart + textBox1.SelectionLength <= textBox1.TextLength ? textBox1.SelectionStart + textBox1.SelectionLength : textBox1.TextLength
                                                            , textBox1.SelectionStart + textBox1.SelectionLength + 10 <= textBox1.TextLength ? 10 : textBox1.TextLength - (textBox1.SelectionStart + textBox1.SelectionLength)), string.Empty, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                        }
                        if (dr == DialogResult.Yes)
                            textBox1.SelectedText = "}}\r\n{{";
                        //textBox1.Text = textBox1.Text.Replace("}}。<p>\r\n{{", "}}\r\n{{");
                        else if (dr == DialogResult.Cancel)//●●●●●●●●●●●●●●●●20250308
                            break;
                    }
                    ip = textBox1.Text.IndexOf(item, ip + 1);
                }

            }
            stopUndoRec = false; ResumeEvents();
        }
        /// <summary>
        /// 清除注腳間「{{」與「}}」間的分段標記<p>
        /// 小注間的分段標記檢查與清除
        /// 20240904 creedit_with_Copilot大菩薩：C# Windows.Forms 中檢查並清除大括號內的 `<p>` 標籤：https://sl.bing.net/dc0DoIfOFSm
        /// </summary>
        private void clearParagraphMarkersInsidePairsBrackets()
        {
            string x = textBox1.Text;
            while (isParagraphMarkersInsidePairsBrackets())
            {
                if (x == textBox1.Text) break;
                x = textBox1.Text;
            }
        }
        /// <summary>
        /// 判斷注腳間「{{」與「}}」間是否有分段標記<p>        
        /// 20240904 creedit_with_Copilot大菩薩：C# Windows.Forms 中檢查並清除大括號內的 `<p>` 標籤：https://sl.bing.net/dc0DoIfOFSm
        /// </summary>
        private bool isParagraphMarkersInsidePairsBrackets()
        {//這段程式碼會在textBox1中檢查大括號內是否有<p>標籤，並在找到後選取該部分，提示使用者是否要清除。如果使用者選擇清除，則會移除<p>標籤。


            string input = textBox1.Text;
            string[] checkStr = { @"\{[^{}]*。<p>[^{}]*\}", @"\{[^{}]*<p>[^{}]*\}" };
            //string pattern = @"\{[^{}]*<p>[^{}]*\}";
            foreach (var item in checkStr)
            {
                string pattern = item;
                Match match = Regex.Match(input, pattern);

                if (match.Success)
                {
                    string marker = item.Substring(item.IndexOf("*") + 1, item.IndexOf(">") - item.IndexOf("*"));
                    //textBox1.Select(match.Index, match.Length);
                    textBox1.Select(textBox1.Text.IndexOf(marker, match.Index), marker.Length);
                    DialogResult result;
                    if (!ocrTextMode)
                        result = MessageBox.Show("發現小注間有<p>標籤，是否清除？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    else
                        result = DialogResult.Yes;
                    AvailableInUseBothKeysMouse();
                    if (result == DialogResult.Yes)
                    {
                        stopUndoRec = true; PauseEvents();

                        //string textBox1SelectedText = textBox1.SelectedText;
                        //textBox1.SelectedText = string.Empty;
                        //textBox1.Text = Regex.Replace(input, pattern, m => m.Value.Replace("<p>", ""));
                        //textBox1.SelectedText = Regex.Replace(textBox1SelectedText, pattern, m => m.Value.Replace("<p>", ""));
                        textBox1.SelectedText = string.Empty;//清除 <p>
                                                             //將接續的空白「􏿽」轉成空格「　」
                        if (textBox1.SelectionStart < textBox1.TextLength - 4)
                        {
                            //如果插入點後是分段/行符號
                            if (textBox1.Text.Substring(textBox1.SelectionStart, Environment.NewLine.Length) == Environment.NewLine)
                            {
                                textBox1.SelectionStart += Environment.NewLine.Length;
                                while (textBox1.Text.Substring(textBox1.SelectionStart, 2) == "􏿽")
                                {
                                    textBox1.Select(textBox1.SelectionStart, 2);
                                    textBox1.SelectedText = "　";
                                }
                            }
                        }

                        stopUndoRec = false; ResumeEvents();
                    }
                    return true;
                }
                //else
                //{
                //    //MessageBox.Show("未發現<p>標籤。");
                //    return false;
                //}
            }
            return false;
        }
        /// <summary>
        /// 清除大括號間的下大括號
        /// </summary>
        void clearCloseBracketsInsidePairsBrackets()
        {
            TextBox tb = textBox1;
            if (tb.Text.IndexOf("{") == -1 || tb.Text.IndexOf("}") == -1) return;
            int openP = tb.Text.IndexOf("{{");
            int p = tb.Text.IndexOf("}}");
            int nextopenP = tb.Text.IndexOf("{{", openP + 2);//2="{{".Length
            if (nextopenP == -1) return;
            int closeP = tb.Text.LastIndexOf("}}", nextopenP);
            while (closeP > -1)
            {

                if (p < closeP)
                {
                    stopUndoRec = true; PauseEvents();
                    tb.Select(p, 2);//2="}}".Length
                    tb.SelectedText = string.Empty;
                    stopUndoRec = false; ResumeEvents();
                }
                if (nextopenP == -1) break;
                openP = nextopenP;
                p = tb.Text.IndexOf("}}", openP);
                if (p == -1) break;
                nextopenP = tb.Text.IndexOf("{{", openP + 2);//2="{{".Length
                if (nextopenP == -1) //break;
                    closeP = tb.Text.LastIndexOf("}}", tb.TextLength);
                else
                    closeP = tb.Text.LastIndexOf("}}", nextopenP);
                //Console.WriteLine(tb.Text.Substring(openP, nextopenP - openP));
                //Console.WriteLine(tb.Text.Substring(openP, p - openP));
                //Console.WriteLine(tb.Text.Substring(openP, closeP - openP));
            }
        }
        /// <summary>
        /// 清除大括號間的下大括號
        /// 此法已臻完善，唯讀入《古籍酷》OCR批量處理文本之方法已得改善，當不會再出現如此情況，故今先束之高閣，以待有需時再調用 20240913
        /// </summary>
        void clearBracketsInsidePairsBrackets()
        {
            TextBox tb = textBox1;
            if (tb.Text.IndexOf("{{") == -1 || tb.Text.IndexOf("}}") == -1) return;
            int s = 0;
            //先看第一個出現的是不是上大括號            
            if (tb.Text.IndexOf("}}") > -1 && tb.Text.IndexOf("{{") > tb.Text.IndexOf("}}"))
                s = tb.Text.IndexOf("}}");
            int openP = tb.Text.IndexOf("{{", s);
            int closeP = tb.Text.IndexOf("}}", s + 2);
            int nextopenP = tb.Text.IndexOf("{{", openP + 2);//2="{{".Length
            int nextcloseP = 0;
            while (nextopenP > -1 && nextopenP > -1)
            {
                if (closeP < openP)
                {
                    tb.Select(closeP, openP);
                    MessageBoxShowOKExclamationDefaultDesktopOnly("大括號錯亂處請先清理！");
                    break;
                }
                string stringBrackets =
                    tb.Text.Substring(openP + 2, closeP - openP - 2);
                if (stringBrackets.StartsWith("{"))
                {
                    openP = tb.Text.IndexOf("{{", openP + 2);
                    closeP = tb.Text.IndexOf("}}", openP);
                    nextopenP = tb.Text.IndexOf("{{", openP + 2);
                    nextcloseP = tb.Text.IndexOf("}}", closeP + 2);
                    continue;//{{{……}}} 不處理
                }

                if (stringBrackets.IndexOf("{{") > -1)
                {
                    tb.Select(nextopenP, 2);
                    if (tb.Text.Substring(tb.SelectionStart + tb.SelectionLength, 1) != "{")//{{{……}}} 不處理
                    {
                        stopUndoRec = true; PauseEvents();
                        tb.SelectedText = string.Empty;
                        stopUndoRec = false; ResumeEvents();
                    }
                }
                nextopenP = tb.Text.IndexOf("{{", openP + 2);
                if (nextopenP == -1)
                {
                    nextcloseP = tb.Text.IndexOf("}}", closeP + 2);
                    while (nextcloseP > -1)
                    {
                        tb.Select(nextcloseP, 2);
                        if (tb.Text.Substring(tb.SelectionStart + tb.SelectionLength, 1) != "}")
                        {
                            stopUndoRec = true; PauseEvents();
                            tb.SelectedText = string.Empty;
                            stopUndoRec = false; ResumeEvents();
                            nextcloseP = tb.Text.IndexOf("}}", closeP + 2);
                        }
                        else
                            nextcloseP = tb.Text.IndexOf("}}", nextcloseP + 2);
                    }
                    break;
                }
                closeP = tb.Text.IndexOf("}}", openP);
                nextcloseP = tb.Text.IndexOf("}}", closeP + 2);
                if (nextcloseP == -1)
                {
                    nextopenP = tb.Text.IndexOf("{{", nextopenP + 2);
                    while (nextopenP > -1)
                    {
                        tb.Select(nextopenP, 2);
                        if (tb.Text.Substring(tb.SelectionStart + tb.SelectionLength, 1) != "{")
                        {
                            stopUndoRec = true; PauseEvents();
                            tb.SelectedText = string.Empty;
                            stopUndoRec = false; ResumeEvents();
                            nextopenP = tb.Text.IndexOf("{{", closeP + 2);
                        }
                        else
                            nextopenP = tb.Text.IndexOf("{{", nextopenP + 2);
                    }
                    break;
                }
                if (nextcloseP < nextopenP)
                {
                    stringBrackets = tb.Text.Substring(openP + 2, nextcloseP - openP - 2);
                    if (stringBrackets.IndexOf("}}") > -1)
                    {
                        tb.Select(closeP, 2);
                        if (tb.Text.Substring(tb.SelectionStart + tb.SelectionLength, 1) != "}")
                        {
                            stopUndoRec = true; PauseEvents();
                            tb.SelectedText = string.Empty;
                            stopUndoRec = false; ResumeEvents();
                        }
                    }
                }
                if (openP + 2 > tb.TextLength) break;
                openP = tb.Text.IndexOf("{{", openP + 2);
                closeP = tb.Text.IndexOf("}}", openP);
                nextopenP = tb.Text.IndexOf("{{", openP + 2);
                nextcloseP = tb.Text.IndexOf("}}", closeP + 2);
                stringBrackets =
                    tb.Text.Substring(openP + 2, closeP - openP - 2);
                if (stringBrackets.StartsWith("{")) continue;//{{{……}}} 不處理
                while (nextcloseP > -1 && nextopenP == -1)
                {//先清掉最後多餘的下大括弧，而不是清掉其中間的下大括號 20240904
                    tb.Select(nextcloseP, 2);//2 = "}}".Length
                    if (nextcloseP + 2 == tb.TextLength || tb.Text.Substring(tb.SelectionStart + tb.SelectionLength, 1) != "}")
                    {
                        stopUndoRec = true; PauseEvents();
                        tb.SelectedText = string.Empty;
                        stopUndoRec = false; ResumeEvents();
                        nextcloseP = tb.Text.IndexOf("}}", closeP + 2);
                    }
                    else
                        nextcloseP = tb.Text.IndexOf("}}", nextcloseP + 2);
                }
            }
        }

        #region 將空格後的句號（如「　。」）置於所有空格前
        /// <summary>
        /// 將全形空格後的句號（如「　。」）置於所有空格前 20240904
        /// </summary>
        void movePeriodsToFrontofBlank()
        {
            string x = textBox1.Text;
            int p = x.IndexOf("　。");
            if (p == -1) return;
            int s = 1;
            while (p > -1)
            {
                //移動到非全形空格時
                while (p - 1 > -1 && textBox1.Text.Substring(p - s, 1) == "　")
                { s++; }

                stopUndoRec = true; PauseEvents();

                textBox1.Select(p - s + 1, p + 2 - (p - s + 1));//2="　。".Length
                textBox1.SelectedText = "。" + textBox1.SelectedText.Substring(1, textBox1.SelectionLength - 2) + "　";

                stopUndoRec = false; ResumeEvents();

                p = textBox1.Text.IndexOf("　。");
            }
        }
        #endregion


        /// <summary>
        /// 軟體操作時提醒之系統音效參照
        /// </summary>
        public enum soundLike
        {
            none, over, done, stop, info, error, warn, exam, processing, press, waiting, finish,
            notify
        }
        /// <summary>
        /// 播放指定音效
        /// </summary>
        /// <param name="sndlike">音效的名稱（作用、含義）</param>
        public static void playSound(soundLike sndlike, bool soundAnyway = false)
        {
            if (MuteProcessing && !soundAnyway) return;
            string mediaPathWithBackslash = Environment.GetFolderPath(Environment.SpecialFolder.Windows) + "\\Media\\";
            string wav = "";
            switch (sndlike)
            {
                case soundLike.over:
                    wav = "windows logoff sound";
                    break;
                case soundLike.done:
                    wav = "Windows Notify Messaging";
                    break;
                case soundLike.stop:
                    wav = "Windows Exclamation";
                    break;
                case soundLike.info:
                    wav = "tada";
                    break;
                case soundLike.error:
                    wav = "Windows Notify";
                    break;
                case soundLike.warn:
                    wav = "Windows Proximity Notification";
                    break;
                case soundLike.exam:
                    wav = "Windows Notify Email";
                    break;
                case soundLike.processing:
                    wav = "Chimes";
                    break;
                case soundLike.press:
                    wav = "Windows Pop-up Blocked";
                    break;
                case soundLike.waiting:
                    wav = "Ring10";
                    break;
                case soundLike.finish:
                    wav = "Ring03";
                    break;
                case soundLike.notify:
                    //wav = "Windows Unlock";
                    wav = "Windows Shutdown";//Windows 關機
                                             //wav = "Windows Information Bar";//Windows 資訊列
                    break;
                default:
                    break;
            }
            mediaPathWithBackslash += (wav + ".wav");
            if (File.Exists(mediaPathWithBackslash))
                new SoundPlayer(mediaPathWithBackslash).Play();

        }
        /// <summary>
        /// 自動批量連續OCR操作被中止時發出的警示聲
        /// </summary>
        internal static void OCRBreakSoundNotification()
        {
            Task.Run(() =>
            {
                using (SoundPlayer sp = new SoundPlayer("C:\\Windows\\Media\\ring05.wav"))
                {
                    sp.Play();
                    Keys k = ModifierKeys; DateTime dt = DateTime.Now;
                    Task.Run(() => { while (DateTime.Now.Subtract(dt).TotalSeconds < 8 && k == Keys.None) k = ModifierKeys; });
                    if (k != Keys.Control) Thread.Sleep(12000);
                    if (File.Exists("C:\\Windows\\Media\\ring04.wav") && ModifierKeys != Keys.Control && k != Keys.Control)//若需中止，按下Ctrl鍵
                    {
                        sp.SoundLocation = "C:\\Windows\\Media\\ring04.wav";
                        sp.Play();
                        Thread.Sleep(3000);
                    }
                }
            });
        }

        /// <summary>
        /// 將<p>後的空格「　」取代為「􏿽」，只要該行不是篇名
        /// </summary>
        void replaceBlank_ifNOTTitleAndAfterparagraphMark()
        {
            int e = textBox1.Text.IndexOf(Environment.NewLine), s = 0; string px, x = textBox1.Text;
            while (e > -1 && s < x.Length)
            {
                int j = 0;
                px = x.Substring(s, e - s);//取得這一行的文字
                if (s > 5 && x.Substring(s - 5, 5) == "<p>" + Environment.NewLine
                    && px.IndexOf("*") == -1)
                {
                    int i = 0;
                    while (px != "" && i + 1 <= px.Length && px.Substring(i, 1) == "　")
                    {
                        x = x.Substring(0, s + i + j) + "􏿽" + x.Substring(s + i + 1 + j); i++; j++;
                    }
                }
                s = e + Environment.NewLine.Length + j;
                e = x.IndexOf(Environment.NewLine, s);

            }
            stopUndoRec = true; undoRecord(); PauseEvents();
            textBox1.Text = x;
            stopUndoRec = false; ResumeEvents();
        }

        void fillSpace_to_PinchNote_in_LineStart()
        {
            //若行首為有縮排的夾注文，則補上空格
            int e = textBox1.Text.IndexOf(Environment.NewLine), s = 0; string px, x = textBox1.Text;
            while (e > -1 && s < x.Length)
            {

                px = x.Substring(s, e - s);//取得這一行的文字
                int ne = px.IndexOf("}}"), ns = px.IndexOf("{{"), i = 0;//記下注文起訖位置
                if (ns > -1 && px.Substring(0, 1) == "　" && ne < px.Length - 2)
                {
                    string nx = px.Substring(0, ns);
                    if (nx.Replace("　", "") == "")
                    {
                        //若是在標題內則不處理,不在標題內才處理 20250204
                        if (x.LastIndexOf("*", s) < x.LastIndexOf("<p>", s))
                        {
                            if (ns > -1 && ns < ne)
                            {
                                nx = px.Substring(ns + 2, ne - ns - 2);//取得要處理的夾注文
                                StringInfo nxInfo = new StringInfo(nx);
                                int nxCnt = nxInfo.LengthInTextElements;
                                if (nxCnt % 2 == 1)
                                {
                                    nxCnt++;
                                }

                                string space = "　";
                                while (px.Substring(++i, 1) == "　")
                                {
                                    space += "　";
                                }
                                //把{{移到行首
                                x = x.Substring(0, s) + "{{" + space + x.Substring(s + 2 + space.Length);
                                //插入空格
                                x = x.Substring(0, s + ns + 2 +
                                    (int)(nxCnt / 2)) + space +
                                    x.Substring(s + ns + 2 + (int)(nxCnt / 2));
                            }
                        }
                        else //在標題內則將空格轉成空白（當是標題內換行/段的空白）
                        {
                            StringBuilder sb = new StringBuilder();
                            for (int iblk = 0; iblk < nx.Length; iblk++)
                            {
                                sb.Append("􏿽");
                            }
                            x = x.Substring(0, s) + sb.ToString() + x.Substring(s + ns);
                        }
                    }
                }

                s = e + Environment.NewLine.Length + i;
                e = x.IndexOf(Environment.NewLine, s);

            }
            stopUndoRec = true; undoRecord(); PauseEvents();
            textBox1.Text = x;
            stopUndoRec = false; ResumeEvents();

        }

        internal void insertWords(string insX, TextBox tBox)
        {
            undoRecord();
            stopUndoRec = true;
            //textBox1.SelectedText = insX;
            tBox.SelectedText = insX;
            stopUndoRec = false;
            //int s = textBox1.SelectionStart, l = textBox1.SelectionLength;
            //if (l == 0)                
            //    //x = x.Substring(0, s) + insX + x.Substring(s);
            //else
            //x = x.Substring(0, s) + insX + x.Substring(s + l);
            //textBox1.Text = x;
            //s += insX.Length;
            ////textBox1.SelectionStart = s + insX.Length;
            ////textBox1.ScrollToCaret();
            //restoreCaretPosition(textBox1, s, 0);

        }

        List<string> lastKeyPress = new List<string>();
        internal static int findNotChineseCharFarLength(string x, bool forward)
        {
            int isC = 0, l = 0;
            StringInfo xInfo = new StringInfo(x);
            if (forward)
            {
                for (int i = 0; i < xInfo.LengthInTextElements; i++)
                {
                    isC = isChineseChar(xInfo.SubstringByTextElements(i, 1), true);
                    if (isC == 1) l++;
                    //if (isC == 0 && xInfo.SubstringByTextElements(i, 1) == "􏿽") l++;
                    if (isC == 0) return i + 1 + l;//https://www.jb51.net/article/45556.htm
                }
            }
            else
            {
                for (int i = xInfo.LengthInTextElements - 1; i > -1; i--)
                {
                    isC = isChineseChar(xInfo.SubstringByTextElements(i, 1), true);
                    if (isC == 1) l++;
                    //if (isC == 0 && xInfo.SubstringByTextElements(i, 1) == "􏿽") l++;
                    if (isC == 0) return xInfo.LengthInTextElements - i + l;
                }

            }
            return -1;
        }

        int findChineseCharFarLength(string x, bool forward)
        {
            StringInfo xInfo = new StringInfo(x);
            int isC = 0, l = 0;
            if (forward)
            {
                for (int i = 0; i < xInfo.LengthInTextElements; i++)
                {
                    isC = isChineseChar(xInfo.SubstringByTextElements(i, 1), true);
                    if (isC == 1) l++;
                    if (isC == 0 && xInfo.SubstringByTextElements(i, 1) == "􏿽") l++;
                    if (isC != 0) return i + 1 + l;//https://www.jb51.net/article/45556.htm
                }
            }
            else
            {
                for (int i = xInfo.LengthInTextElements - 1; i > -1; i--)
                {
                    isC = isChineseChar(xInfo.SubstringByTextElements(i, 1), true);
                    if (isC == 1) l++;
                    if (isC == 0 && xInfo.SubstringByTextElements(i, 1) == "􏿽") l++;
                    if (isC != 0) return xInfo.LengthInTextElements - i + l;
                }

            }
            return -1;
        }
        /// <summary>
        /// 標點符號和數字
        /// </summary>
        public static readonly string PunctuationsNum = "－.,;?@'\"。，；！？、—…:：《·》〈‧〉「」『』〖〗【】（）()[]〔〕［］“”‘’0123456789-";
        //public static readonly string punctuationsNum = ".,;?@'\"。，；！？、－-—…:：《·》〈‧〉「」『』〖〗【】（）()[]〔〕［］0123456789";
        /// <summary>
        /// 判斷中文字
        /// </summary>
        /// <param name="x">要檢測的字元字串</param>
        /// <param name="skipPunctuation">是否忽略標點符號</param>
        /// <returns>如果是中文傳回絕對值1（非surrogate=-1，surrogate=1；High or Low Surrogate=-1）</returns>
        internal static int isChineseChar(string x, bool skipPunctuation)
        {
            //if (skipPunctuation) if (punctuationsNum.IndexOf(x, StringComparison.Ordinal) > -1) return -1;
            if (skipPunctuation) if (PunctuationsNum.Replace("《", "").IndexOf(x, StringComparison.Ordinal) > -1) return -1;//先拿掉「《」不計 20240315
            const string cha = "�□▫စခငဇဌ◍ᗍⲲ⛋ဂဃဆဈဉ";
            string notChineseCharPriority = cha + "〇◯　 \r\n<>{}.,;?@●'\"。，；！？、－-《》〈〉「」『』〖〗【】（）()[]〔〕［］0123456789";

            if (notChineseCharPriority.IndexOf(x, StringComparison.Ordinal) > -1) return 0;

            if (Regex.IsMatch(x, @"[a-zA-z]")) return 0;
            //if (x == "\udffd") return 0;

            //https://www.jb51.net/article/45556.htm
            //https://zh.wikipedia.org/wiki/%E4%B8%AD%E6%97%A5%E9%9F%93%E7%B5%B1%E4%B8%80%E8%A1%A8%E6%84%8F%E6%96%87%E5%AD%97
            if (Regex.IsMatch(x, @"[\u4e00-\u9fbb]")) return -1;
            if (Regex.IsMatch(x, @"[\u3400-\u4dbf]")) return -1;//擴充A區包含有6,592個漢字，位置在U+3400—U+4DBF
                                                                //以下長度不同,恐怕失效,目前知C即不行,有空再測試            
                                                                //c# 中文 轉 unicode:
                                                                //https://www.google.com/search?q=c%23+%E4%B8%AD%E6%96%87+%E8%BD%89+unicode&rlz=1C1GCEU_zh-TWTW823TW823&sxsrf=AOaemvJI_o6pHrTEJVPCsVy0iyVsclLtjQ%3A1640527095825&ei=93TIYbnqMYOmoATnx4rwBg&oq=c%23++%5Cu%E4%B8%AD%E6%96%87%E5%AD%97%E7%A2%BC&gs_lcp=Cgdnd3Mtd2l6EAMYATIFCAAQzQIyBQgAEM0COggIABCwAxDNAjoECCMQJ0oECEEYAUoECEYYAFCzWFjiY2DfcGgCcAB4AIABVYgB1gGSAQEzmAEAoAEByAECwAEB&sclient=gws-wiz
                                                                //https://www.itread01.com/p/1418585.html


            if (Regex.IsMatch(x, @"[\u20000-\u2A6DD]")) return -1;//擴充B區包含有42,717個漢字，位置在U+20000—U+2A6DD
            if (Regex.IsMatch(x, @"[\u2A700-\u2B734]")) return -1;//C:位置在U+2A700—U+2B734
            if (Regex.IsMatch(x, @"[\u2B740-\u2B81F]")) return -1;//D:範圍為U+2B740–U+2B81F（實際有字元為U+2B740–U+2B81D）
            if (Regex.IsMatch(x, @"[\u2B820-\u2CEAF]")) return -1;//true;//E:編碼範圍U+2B820–U+2CEAF
            if (Regex.IsMatch(x, @"[\u2CEB0-\u2EBEF]")) return -1;//true;//F:U+2CEB0–U+2EBEF
            if (Regex.IsMatch(x, @"[\u30000-\u3134A]")) return -1;//true;//G:U+30000–U+3134A
                                                                  //if (Regex.IsMatch(x, @"[\u-\u]")) return true;//

            //if (char.IsSurrogate(x.ToCharArray()[0]) || char.IsSurrogatePair(x, 0)) return true;
            if (char.IsSurrogatePair(x, 0))
            {
                if (x == "􏿽")
                {
                    return 0;
                }
                return 1;
            }
            if (char.IsLowSurrogate(x, 0))
            {
                if (x == "􏿽".Substring(1, 1))
                {
                    return 0;
                }
                return -1;
            }
            if (char.IsHighSurrogate(x, 0))
            {
                if (x == "􏿽".Substring(0, 1))
                {
                    return 0;
                }
                return -1;
            }
            /*
            //https://www.itread01.com/p/1418585.html
            //C#中文字轉換Unicode(\u ) : http://trufflepenne.blogspot.com/2013/03/cunicode.html
            string outStr = "";
            if (!string.IsNullOrEmpty(x))
            {
                for (int i = 0; i < x.Length; i++)
                {
                    outStr += "/u" + ((int)x[i]).ToString("x");
                }
            }
            x = outStr;
            */
            return 0;
        }

        /// <summary>
        ///判断CJK字符集// 判断一个字符是否是CJK或CJK扩展字符集中的汉字
        /// 202302015 chatGPT大菩薩
        /// 在C#中，可以使用Unicode字符编码值的范围来判断一个字符是否是CJK或CJK扩展字符集中的汉字。……
        /// 这段代码中的 IsChineseCharacter 方法用于判断单个字符是否是CJK或CJK扩展字符集中的汉字，而 IsChineseString 方法则用于判断一个字符串是否全部由CJK或CJK扩展字符集中的汉字组成。
        /// 在判断一个字符是否是CJK或CJK扩展字符集中的汉字时，我们使用Unicode字符编码值的范围来进行判断。CJK字符集范围是从0x4e00到0x9fff，而CJK扩展字符集范围是从0x20000到0x2a6df。因此，如果一个字符的Unicode编码值在这个范围内，那么就可以判断它是CJK或CJK扩展字符集中的汉字。
        /// </summary>
        /// <param name="c">要檢查的字元</param>
        /// <returns></returns>
        public static bool IsChineseCharacter(char c)
        {
            // Unicode范围: CJK字符集范围：4E00–9FFF，CJK扩展字符集范围：20000–2A6DF
            return (c >= 0x4e00 && c <= 0x9fff) || (c >= 0x20000 && c <= 0x2a6df);
        }

        /// <summary>
        /// 判断一个字符串是否全部由CJK或CJK扩展字符集中的汉字组成
        /// 202302015 chatGPT大菩薩
        /// </summary>
        /// <param name="s">要檢查的文本</param>
        /// <returns></returns>        
        public static bool IsChineseString(string s)
        {
            foreach (char c in s)
            {
                if (!IsChineseCharacter(c))
                {
                    //return false;
                    //return (Math.Abs(isChineseChar(s, false)) == 1) ? true : false;
                    return (Math.Abs(isChineseChar(c.ToString(), false)) == 1);
                }
            }
            return true;
        }
        /// <summary>
        /// 判断內容字串是否全部由CJK或CJK擴充字集中的漢字组成（含中文標點符號也算中文。即以CTP除了語法標記外的內文為主）
        /// 202302015 chatGPT大菩薩
        /// </summary>
        /// <param name="s">要檢查的文本</param>
        /// <returns></returns>        
        public static bool IsChineseContent(string s)
        {
            foreach (char c in s)
            {
                if (!IsChineseCharacter(c) && "《》〈〉·。，！；、：「」『』？".IndexOf(c) == -1)
                {
                    //return false;
                    //return (Math.Abs(isChineseChar(s, false)) == 1) ? true : false;
                    return (Math.Abs(isChineseChar(c.ToString(), false)) == 1);
                }
            }
            return true;
        }
        /// <summary>
        /// 判断內容字串是否全無由CJK或CJK擴充字集中的漢字组成（含中文標點符號也算中文。即以CTP除了語法標記外的內文為主）
        /// 202302015 chatGPT大菩薩
        /// </summary>
        /// <param name="s">要檢查的文本</param>
        /// <returns></returns>        
        public static bool IsNoChineseContent(string s)
        {
            foreach (char c in s)
            {
                if (IsChineseCharacter(c) && "《》〈〉·。，！；、：「」『』？".IndexOf(c) == -1)
                    return false;
                else
                {
                    if (Math.Abs(isChineseChar(c.ToString(), false)) == 1)
                    {
                        return false;
                    }
                }
            }
            return true;
        }
        /// <summary>
        /// 是否是中文或surrogate
        /// </summary>
        /// <param name="xCheck"></param>
        /// <returns></returns>
        public static bool IsCJKorSurrogate(string xCheck)
        {
            return char.IsSurrogate(xCheck.ToCharArray()[0]) ||
                        IsChineseString(xCheck);
        }

        /// <summary>
        /// C#中文字轉換Unicode(\u ):http://trufflepenne.blogspot.com/2013/03/cunicode.html
        /// </summary>
        /// <param name="srcText"></param>
        /// <returns></returns>        
        public static string StringToUnicode(string srcText)
        {
            string dst = "";
            char[] src = srcText.ToCharArray();
            for (int i = 0; i < src.Length; i++)
            {
                byte[] bytes = Encoding.Unicode.GetBytes(src[i].ToString());
                string str = @"\u" + bytes[1].ToString("X2") + bytes[0].ToString("X2");
                dst += str;
            }
            return dst;
        }

        public static string UnicodeToString(string srcText)
        {
            string dst = "";
            string src = srcText;
            int len = srcText.Length / 6;

            for (int i = 0; i <= len - 1; i++)
            {
                string str = "";
                str = src.Substring(0, 6).Substring(2);
                src = src.Substring(6);
                byte[] bytes = new byte[2];
                bytes[1] = byte.Parse(int.Parse(str.Substring(0, 2), System.Globalization.NumberStyles.HexNumber).ToString());
                bytes[0] = byte.Parse(int.Parse(str.Substring(2, 2), System.Globalization.NumberStyles.HexNumber).ToString());
                dst += Encoding.Unicode.GetString(bytes);
            }
            return dst;
        }


        /* 20230408 Bing大菩薩 ： 您可以使用正則表達式來簡化您的 if 判斷句。例如，您可以將條件提取到一個單獨的函數中，並使用正則表達式來檢查 url 是否包含特定字符串：
         */
        /// <summary>
        /// 檢查要輸入簡單修改模式頁面的指定網址是否合法
        /// </summary>
        /// <param name="url">要檢查的網址字串值</param>
        /// <returns>回傳網址是否合法</returns>
        internal static bool IsValidUrl＿keyDownCtrlAdd(string url)
        {
            url = br.ClearUrl_BoxEtc(url);
            //return Regex.IsMatch(url, @"(#editor|&page=|ctext\.org)");
            //return Regex.IsMatch(url, @"ctext\.org.*&file.*&page=.*#editor");
            //也有可能是這種網址：https://ctext.org/library.pl?if=gb&file=34195&page=142&editwiki=826120#box(140,120,2,0)
            //return Regex.IsMatch(url, @"ctext\.org.*&file.*&page=.*&edit");
            return Regex.IsMatch(url, @"ctext\.org.*&file.*&page=.*#edit") || Regex.IsMatch(url, @"ctext\.org.*&file.*&page=.*&editwiki=.*");//20250126
            /*
             * Bing大菩薩：是的，在正則表達式中，小數點「.」是一個特殊字符，它匹配任何單個字符（除了換行符）。如果您想在正則表達式中匹配字面上的小數點，則需要在前面加上反斜杠「\」來對其進行轉義。
             * 在 C# 中，由於反斜杠「\」本身也是一個轉義字符，所以您需要使用兩個反斜杠「\\」來表示一個字面上的反斜杠。因此，在 C# 中的正則表達式中，要匹配字面上的小數點，您需要寫成「\\.」。
                希望這對您有所幫助！*/
        }
        /// <summary>
        /// 檢查是否是瀏覽圖文對照之頁面
        /// 可與 isQuickEditUrl 方法互參用
        /// </summary>
        /// <param name="url">要檢查的網址字串值</param>
        /// <returns></returns>
        internal static bool IsValidUrl＿ImageTextComparisonPage(string url)
        {
            return Regex.IsMatch(url, @"ctext\.org.*&file.*&page=");
        }
        /// <summary>
        /// 將圖文對照網址修整、規範之
        /// 20240813 creedit with Copilot大菩薩：改進C#程式碼：圖文對照網址修整：https://sl.bing.net/f2S0RcHJLyK
        /// 與 ReplaceUrl_Box2Editor 可互參考
        /// </summary>
        /// <param name="url">要被修整、規範化的圖文對照網址</param>
        /// <param name="editor">是否要在末尾改綴上"#editor"字串</param>
        /// <param name="driverGoToUrl">是否要移至這個網址</param>
        /// <returns>回傳修整過、規範的圖文對照網址</returns>
        internal static string FixUrl＿ImageTextComparisonPage(string url, bool editor = false, bool driverGoToUrl = false)
        {
            #region 防呆
            if (!IsValidUrl＿ImageTextComparisonPage(url) || browsrOPMode == BrowserOPMode.appActivateByName || br.driver == null) return null;
            #endregion

            // 使用正則表達式檢查和替換網址中的特定字串
            url = System.Text.RegularExpressions.Regex.Replace(url, "#box\\(.*?\\)", editor ? "#editor" : string.Empty);

            try
            {
                if (driverGoToUrl) br.driver.Navigate().GoToUrl(url);
            }
            catch (Exception ex)
            {
                // 記錄詳細的錯誤訊息
                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly($"Error: {ex.HResult} - {ex.Message}");
            }

            return url;
        }
        /// <summary>
        /// 是否是四庫全書的扉頁
        /// </summary>
        /// <param name="x">要作為判斷的文本</param>
        /// <returns></returns>
        bool isSKQSFrontPage(string x)
        {
            if (Name != "Form1")
            {

                if (x.Contains("欽定四庫全書") && x.Contains("部") && x.Contains("校官") && x.Contains("{{臣}}"))
                {
                    return true;
                }
            }
            return false;
        }

        /* 使用正則表達式：可以使用正則表達式來檢查和替換網址中的特定字串，這樣會更靈活和高效。
            簡化條件檢查：將防呆區塊的條件檢查合併成一行，讓程式碼更簡潔。
            改進例外處理：在例外處理區塊中，記錄詳細的錯誤訊息，方便日後除錯。這樣的改進應該能讓程式碼更簡潔、更高效。……
         */

        //internal static string FixUrl＿ImageTextComparisonPage(string url, bool editor = false, bool driverGoToUrl = false)
        //{
        //    #region 防呆
        //    if (!IsValidUrl＿ImageTextComparisonPage(url)) return null;
        //    if (browsrOPMode == BrowserOPMode.appActivateByName) return null;
        //    if (br.driver == null) return null;
        //    #endregion

        //    int boxTag = url.IndexOf("#box(");
        //    if (boxTag > -1)
        //    {
        //        playSound(soundLike.exam);
        //        url = url.Substring(0, boxTag) + (editor ? "#editor" : string.Empty);
        //        try
        //        {
        //            if (driverGoToUrl) br.driver.Navigate().GoToUrl(url);

        //        catch (Exception ex)
        //        {
        //            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
        //        }
        //    }

        //    return url;
        //}

        /// <summary>
        /// Ctrl + + （加號，含函數字鍵盤） 或 Ctrl + -（數字鍵盤）  或 Ctrl + 5 (數字鍵盤） 或 Alt + + ：
        /// 將插入點或選取文字（含）之前的文本剪下貼到 ctext 的[簡單修改模式]框中，並按下「保存編輯」鈕，且
        /// 在非自動連續輸入時于瀏覽器新頁籤（預設值，Selenium架構時不會）開啟下一頁準備編輯文本，並回到前一頁籤以供檢視所貼上之文本是否無誤。
        /// </summary>
        /// <param name="shiftKeyDownYet">按下Shift則留下本頁不自動翻至下一頁</param>
        /// <param name="clear">選擇性參數：若指定chkClearQuickedit_data_textboxTxtStr則會清除當前文字框內容而非輸入新內容</param>        
        /// <returns>執行不成功則回傳false</returns>
        private bool keyDownCtrlAdd(bool shiftKeyDownYet = false, string clear = "", bool notBooksPunctuation = false, bool pagePaste2GjcoolOCR = false)
        {
            if (textBox1.TextLength == 0) return false;

            undoRecord();
            stopUndoRec = true; PauseEvents();//●●●●●●●●●●●●●
            int s = textBox1.SelectionStart, l = textBox1.SelectionLength; string x = textBox1.Text; //今定義再置前
            int chkLoaction;// = 0;//檢查文本定位用
            bool _eventabled = _eventsEnabled;
            TopMost = false;

        reGetURL:
            string url = textBox3.Text, urlShort = url.Substring(0, url.IndexOf("#editor") == -1 ? url.Length : url.IndexOf("#editor")), urlDriver = string.Empty;
            try
            {
                urlDriver = br.driver.Url;
                if (!IsValidUrl＿keyDownCtrlAdd(url) && !IsValidUrl＿keyDownCtrlAdd(urlDriver)) { MessageBoxShowOKExclamationDefaultDesktopOnly("請檢查網址再重試！" + Environment.NewLine + "driver.Url= " + urlDriver + Environment.NewLine + "textBox3.Text= " + url); return false; }
                if (urlDriver.StartsWith("https://ctext.org/library.pl?if=gb&file=") && br.isQuickEditUrl(urlDriver) == false)
                {
                    int boxTag = urlDriver.IndexOf("#box(");
                    if (boxTag > -1)
                    {
                        //playSound(soundLike.exam);
                        //urlDriver = urlDriver.Substring(0, boxTag) + "#editor";
                        urlDriver = FixUrl＿ImageTextComparisonPage(urlDriver, true, true);
                        //br.driver.Navigate().GoToUrl(urlDriver);
                        goto reGetURL;
                    }
                    if (DialogResult.OK == MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否要開啟[簡單修改模式](quick edit)？" + Environment.NewLine + Environment.NewLine + "driver.Url= " + urlDriver))
                    { br.QuickeditLinkIWebElement?.Click(); textBox3.Text = br.driver.Url; goto reGetURL; }
                }
            }
            catch (Exception)
            {
                try
                {
                    br.driver.SwitchTo().Window(br.LastValidWindow);
                    urlDriver = br.driver.Url;
                }
                catch (Exception)
                {
                    ResetLastValidWindow();
                    try
                    {
                        br.driver.SwitchTo().Window(br.LastValidWindow);
                        urlDriver = br.driver.Url;
                    }
                    catch (Exception)
                    {
                        RestartChromedriver();
                        return false;
                    }
                }
            }
        reCheckUrl:
            if (!urlDriver.StartsWith(urlShort))
            {//真的都是擴充功能Google翻譯在作怪！難怪害得 driver.Url 和 WindowHandles.Count 都無法取得正確值！
                if (urlDriver == "chrome-extension://aapbdbdomjkkjkaonfhkkikfgjllcleb/offscreen.html") playSound(soundLike.exam, true);//Debugger.Break();
                else if (urlDriver.StartsWith("chrome-extension://")) playSound(soundLike.exam, true);//Debugger.Break();
                if (urlDriver.StartsWith("https://ctext.org/library.pl?if=gb&file=") && url.StartsWith("https://ctext.org/library.pl?if=gb&file=")
                    && urlDriver.EndsWith("#editor") && url.EndsWith("#editor") && url == textBox3.Text)//string url = textBox3.Text 見前 20240803
                {
                    playSound(soundLike.exam, true);
                    textBox3.Text = urlDriver;
                    url = urlDriver;
                }
                else
                {
                    bool found = false;
                    //foreach (var item in br.driver.WindowHandles)
                    for (int i = br.driver.WindowHandles.Count - 1; i > -1; i--)
                    {
                        br.driver.SwitchTo().Window(br.driver.WindowHandles[i]);
                        if (br.driver.Url.StartsWith(urlShort))
                        {
                            found = true; break;
                        }
                    }
                    if (!found)
                    {
                        if (!string.IsNullOrEmpty(br.LastValidWindow))
                        {
                            br.driver.SwitchTo().Window(br.LastValidWindow);
                            if (br.driver.Url != textBox3.Text)
                                textBox3.Text = br.driver.Url;
                        }
                    }
                    else
                    {
                        br.LastValidWindow = br.driver.CurrentWindowHandle;
                        urlDriver = br.driver.Url;
                        if (!urlDriver.StartsWith(urlShort)) goto reCheckUrl;
                    }
                }
            }

            #region 若是《四庫全書》原書的扉頁直接加入〖文淵|閣寶〗
            if (isSKQSFrontPage(x))
            {
                TopMost = false;
                br.driver.SwitchTo().Window(driver.CurrentWindowHandle);
                br.BringToFront("chrome");
                br.ChromeSetFocus();
                //WindowState = FormWindowState.Minimized;
            }
            #endregion 若是《四庫全書》直接加入〖文淵|閣寶〗

            #region 在手動編輯模式下（尤其是需要OCR時）的前置檢查
            if (keyinTextMode)
            {//如果是在手動輸入模式下：
                if (s == 0 && l == 0) s = textBox1.TextLength;
                else if (s + l < textBox1.TextLength &&//空格與分行/段符號網頁會自動忽略
                                                       //textBox1.Text.Substring(s + l, textBox1.TextLength - s - l).Replace("　", "").Replace(Environment.NewLine, "") != "")
                    x.Substring(s + l, textBox1.TextLength - s - l).Replace("　", "").Replace(Environment.NewLine, "") != "")
                {
                    if (DialogResult.Cancel == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly(
                        "插入點位置似有誤，請「確定」從此處之前的才貼上？\n\r\n\r" +
                        "忽略此訊息，改為【整面貼上】請按「取消」感恩感恩　南無阿彌陀佛", string.Empty, false))
                    {
                        textBox1.DeselectAll();
                        textBox1.SelectionStart = textBox1.TextLength;
                        s = textBox1.SelectionStart; l = 0; pageTextEndPosition = s + l;
                    }
                }

                if (keyinTextMode && !OcrTextMode)
                {
                    //檢查版心內容是否闌入？
                    chkLoaction = CnText.HasPlatecenterTextIncluded(x);
                    if (chkLoaction > -1)
                    {
                        int selStart = 0;
                        if (chkLoaction == 0)
                            selStart = chkLoaction;
                        else
                        {//if (chkLoaction > 0)
                            selStart = x.LastIndexOf(Environment.NewLine, chkLoaction);
                            if (selStart == -1)
                                selStart = 0;
                            //else //分段符號也可以刪掉，故也一併選取，不用避開
                            //selStart += Environment.NewLine.Length;
                        }
                        int selEnd = selStart == 0 ? x.IndexOf(Environment.NewLine, chkLoaction) : x.Length;
                        textBox1.Select(selStart, selStart == 0 ?
                            (selEnd == -1 ? (x.Length - selStart) : (selEnd - selStart)) + Environment.NewLine.Length :
                            (selEnd == -1 ? (x.Length - selStart) : (selEnd - selStart)));
                        textBox1.ScrollToCaret();
                        if (MessageBoxShowOKCancelExclamationDefaultDesktopOnly("【版心】內容似還殘留，確定送出？", "阿彌陀佛", true, MessageBoxDefaultButton.Button2) == DialogResult.Cancel)
                        {
                            AvailableInUseBothKeysMouse();
                            TopMost = true;
                            //選取疑似版心內容段落（行）以供檢查或逕予刪除
                            return false;
                        }
                    }
                }

                if (!PasteOcrResultFisrtMode)
                {//檢查查是否有編輯標記
                    CnText.FormalizeText(ref x);
                    if (!CnText.HasEditedWithPunctuationMarks(ref x))
                    {
                        //playSound(soundLike.warn);
                        if (MessageBoxShowOKCancelExclamationDefaultDesktopOnly("尚未有以供程式判斷之編輯標記（標點符號及符號格式化字元等），是否確定送出？", string.Empty, true, MessageBoxDefaultButton.Button2) == DialogResult.Cancel) return false;
                    }
                }

                //TopMost = false;//前面已有： if (TopMost) TopMost = false; //20240211
                //將焦點交給Chrome瀏覽器
                if (browsrOPMode != BrowserOPMode.appActivateByName && br.driver != null)
                {
                    PauseEvents();
                    try
                    {
                        br.driver.SwitchTo().Window(br.LastValidWindow);
                    }
                    catch (Exception ex)
                    {
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message + Environment.NewLine + "◆請按「確定」繼續……◆");
                    }
                    ResumeEvents();
                }
                else
                    appActivateByName();
                //Point formPos = new Point(this.Location.X - 10, this.Location.Y);
                //Cursor.Position = formPos;
                //Thread.Sleep(501);// regedit : HKEY_CURRENT_USER\Control Panel\Desktop\ActiveWndTrkTimeout = 500（十進位）
            }
            //如果不是在手動輸入模式
            else { if (s == 0 && l == 0) { Activate(); return false; } }
            #endregion

            x = textBox1.Text; //今定義置前


            //if (pageTextEndPosition == 0) pageTextEndPosition = s;
            if (textBox1.SelectedText == "＠")
            {
                textBox1.Text = x.Substring(0, s) + x.Substring(s + 1);
                l = 0;
                textBox1.Select(s, l);
            }


            #region 小注跨頁處理            
            if (s > 2 && s + 2 <= x.Length && s + l + Environment.NewLine.Length + 2 <= x.Length)
            {
                const string curlyBracketsOpen = "{{", curlyBracketsClose = "}}";
                if (l >= 2)//有選取
                {
                    if (textBox1.SelectedText.Substring(l - 2, 2) == curlyBracketsClose)
                    {
                        if (x.Substring(s + l, 2) == curlyBracketsOpen)
                        {
                            textBox1.Select(s + l - 2, 2 + 2); textBox1.SelectedText = Environment.NewLine;
                        }
                        else if (x.Substring(s + l, Environment.NewLine.Length + 2) == Environment.NewLine + curlyBracketsOpen)
                        {
                            textBox1.Select(s + l - 2, 2 + Environment.NewLine.Length + 2); textBox1.SelectedText = Environment.NewLine;
                        }

                    }

                }
                else if (l == 0)//無選取時
                {
                    if (x.Substring(s - 2, 2) == curlyBracketsClose)
                    {
                        if (x.Substring(s, 2) == curlyBracketsOpen)
                        {
                            textBox1.Select(s - 2, 2 + 2); textBox1.SelectedText = Environment.NewLine;
                        }
                        else if (x.Substring(s, Environment.NewLine.Length + 2) == Environment.NewLine + curlyBracketsOpen)
                        {
                            textBox1.Select(s - 2, 2 + Environment.NewLine.Length + 2);
                            textBox1.SelectedText = Environment.NewLine;
                        }

                    }
                    else if (x.Substring(s - 2, 2) == Environment.NewLine)
                    {
                        if (s - Environment.NewLine.Length - curlyBracketsOpen.Length >= 0 && s + 2 <= x.Length)
                        {
                            if (x.Substring(s, 2) == curlyBracketsOpen &&
                                x.Substring(s - Environment.NewLine.Length - curlyBracketsOpen.Length, 2) == curlyBracketsClose)
                            {
                                textBox1.Select(s - Environment.NewLine.Length - curlyBracketsOpen.Length, 2 * 3); textBox1.SelectedText = Environment.NewLine;
                            }
                        }
                    }
                }
                x = textBox1.Text;
                s = textBox1.SelectionStart; l = textBox1.SelectionLength;
            }
            stopUndoRec = false;//●●●●●●●●●●●●●●●●●●●●●
            ResumeEvents();
            #endregion 跨頁小注處理


            /*
            #region 清除空行//前面處理跨頁小注時須有newline 判斷，故不可寫在其前而執行清除
            if (pageTextEndPosition != 0) s = pageTextEndPosition;
            while (s != textBox1.TextLength && s + 2 < textBox1.TextLength && textBox1.Text.Substring(s, 2) == Environment.NewLine)
            {
                if (textBox1.Text.Substring(s, 2) == Environment.NewLine)
                {
                    textBox1.Select(s, 2); textBox1.SelectedText = "";
                    s = textBox1.SelectionStart;
                }
            }
            #endregion//清除空行
            */
            //tryAgain:
            if (pageTextEndPosition == 0) pageTextEndPosition = s + l;
            else { s = pageTextEndPosition; l = 0; }
            if (s < 0 || s + l > x.Length) s = textBox1.SelectionStart;

            //string xCopy = x.Substring(0, s + l);
            string xCopy = x.Substring(0, (s + l) <= x.Length ? s + l : x.Length);
            ////////////string xCopy = x.Substring(0, s + textBox1.SelectionLength);//前有處理過文本∴不能用textBox1.SelectionLength！！！20230102

            //if (pageEndText10.Length > 20)//此bug已在autoPastetoCtextQuitEditTextbox()內抓到了20230117
            //    pageEndText10 = "";//20230102

            if (pageEndText10 == "") pageEndText10 = xCopy.Substring(xCopy.Length - 10 >= 0 ? xCopy.Length - 10 : xCopy.Length);
            else
            {
                if (xCopy.Length - 10 >= 0 && pageEndText10 != xCopy.Substring(xCopy.Length - 10)
                        && pageEndText10.Replace("}}", Environment.NewLine) != xCopy.Substring(xCopy.Length - 10)
                        && pageEndText10.Replace("}}", Environment.NewLine) != xCopy.Substring(xCopy.Length - 10).Replace("}}", Environment.NewLine))
                {
                    int sNew = x.IndexOf(pageEndText10);
                    if (sNew > -1)
                    {
                        //textBox1.Select(sNew + pageEndText10.Length - predictEndofPageSelectedTextLen, predictEndofPageSelectedTextLen);
                        //textBox1.Select(sNew, pageEndText10.Length);
                        //s = textBox1.SelectionStart; l = textBox1.SelectionLength;
                        s = sNew; l = pageEndText10.Length;
                        if (s + l > x.Length)
                            xCopy = x.Substring(0, x.Length);
                        else
                            xCopy = x.Substring(0, s + l);
                    }
                    else
                    {
                        //如果在手動輸入模式下且所指定的範圍為整個textBox1文字方塊，就不再作頁面末10字（pageEndText10）之檢查
                        if (keyinTextMode)
                        {
                            if (textBox1.SelectionStart < textBox1.TextLength)
                            {
                                //MessageBox.Show("請重新指定頁面結束位置", "", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請重新指定頁面結束位置");
                                pageTextEndPosition = 0; pageEndText10 = "";
                                Activate();
                                return false;
                            }
                            else
                            {
                                s = textBox1.SelectionStart; l = 0;
                                xCopy = textBox1.Text.Substring(0, s + l); pageEndText10 = string.Empty;
                            }
                        }
                        else
                        {//如果在自動輸入模式，且已取得每頁正常行/段數時
                            //if (autoPaste2QuickEdit && (linesParasPerPage != -1 || lines_perPage != 0))
                            if (!keyinTextMode && (linesParasPerPage != -1 || lines_perPage != 0))
                            {
                                if (lines_perPage > 0 && countLinesPerPage(textBox1.Text.Substring(0, textBox1.SelectionStart)) != lines_perPage)
                                {
                                    string PredictCopyX = CnText.GetSelectionTextByLineParaCount(ref textBox1, lines_perPage / 2);
                                    if (DialogResult.OK == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("請重新指定頁面結束位置" + Environment.NewLine + Environment.NewLine +
                                           "是不是這些內容？" +
                                           Environment.NewLine + Environment.NewLine +
                                           PredictCopyX))
                                    {
                                        s = 0; l = PredictCopyX.Length; xCopy = PredictCopyX; pageEndText10 = xCopy.Substring(l - 10);
                                    }
                                    else
                                    {
                                        MessageBox.Show("請重新指定頁面結束位置", "", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                                        pageTextEndPosition = 0; pageEndText10 = "";
                                        Activate(); return false;
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("請重新指定頁面結束位置", "", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                                pageTextEndPosition = 0; pageEndText10 = "";
                                //linesParasPerPage = -1;//●●●●●●●●●●●●20250213元宵後一日
                                //lines_perPage = 0;
                                Activate(); return false;
                            }
                        }
                    }
                }
            }


            //規範化文本，如半形標點符號轉全形：//在 下面 newTextBox1 會執行，此略（須加第3引數×才行，否則原本是根據textBox1.Text來執行的 20240427）
            CnText.FormalizeText(ref xCopy);
            if (!PasteOcrResultFisrtMode && (autoPaste2QuickEdit && lines_perPage == 0))//自動輸入時 lines_perPage 要由 checkAbnormalLinePara 取得
            {
                #region checkAbnormalLinePara method test unit
                try
                {
                    int[] chk;
                    if (keyinTextMode)
                        chk = checkAbnormalLinePara(xCopy.Replace
                           ("<p>" + Environment.NewLine, "★★★").Replace("<p>", string.Empty).Replace("★★★", "<p>" + Environment.NewLine)
                           + xCopy.Substring(xCopy.Length - "<p>".Length, "<p>".Length) == "<p>" ? "<p>" : string.Empty);
                    else
                        chk = checkAbnormalLinePara(xCopy);
                    if (chk.Length > 0)
                    {
                        bringBackMousePosFrmCenter();
                        if (MessageBox.Show("there is abnormal LinePara Length , check it now?" +
                            Environment.NewLine + Environment.NewLine +
                            "normal= " + chk[2] + "\t【abnormal】= " + chk[3], "",
                            MessageBoxButtons.OKCancel, MessageBoxIcon.Warning,
                            MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly) == DialogResult.OK)//檢查行/段落長
                        {
                            textBox1.Select(chk[0], chk[1]);
                            textBox1.ScrollToCaret();
                            if (s > pageTextEndPosition)
                            {
                                pageTextEndPosition = 0;
                            }
                            else
                            {
                                //if (MessageBox.Show("reset the page end ? ", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                //    MessageBoxOptions.ServiceNotification) == DialogResult.OK)
                                pageTextEndPosition = s;
                            }

                            AvailableInUseBothKeysMouse();//BringToFront(); 
                            TopMost = true;
                            return false;
                        }
                        else//按下cancel按鈕,忽略非常的行長度
                        {
                            //wordsPerLinePara為許多判斷行字數函式的重要參考，暫時不在此作調整！20231101
                            NormalLineParaLength = 0; //wordsPerLinePara = chk[chk.Length - 1];
                            TopMost = false;
                            br.driver?.SwitchTo().Window(br.driver.CurrentWindowHandle);

                        }// 目前 chk[chk.Length-1]=3
                    }
                }
                catch (Exception ex)
                {
                    MessageBoxShowOKExclamationDefaultDesktopOnly("  checkAbnormalLinePara函式有誤，請留意！！\n\r" + ex.HResult + ex.Message);
                    AvailableInUseBothKeysMouse();//BringToFront();
                }
                #endregion

            }

            //貼到 Ctext Quick edit 前的文本檢查
            if (abnormalLineParaChecking && fastMode)
            {//在快速模式下，有時《國學大師》或《Kanripo漢籍リポジトリ》等的文本分行不當，2行成1行，故添此判斷
                if (wordsPerLinePara < 1)
                {
                    MessageBoxShowOKExclamationDefaultDesktopOnly("請先設定每行正常長度！感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主"); return false;
                }
                //int[] chk = checkAbnormalLinePara(xCopy);// int st = 0, ln = 0;
                int[] chk = CheckAbnormalLinePara(xCopy, wordsPerLinePara);
                if (chk.Length > 0)
                {
                    br.driver.SwitchTo().Window(br.driver.CurrentWindowHandle);
                    //if (chk[1] - chk[2] > 4)
                    //{//)chk[2] * 0.5)
                    int st = textBox1.SelectionStart;
                    int ln = textBox1.SelectionLength;
                    textBox1.Select(chk[0], chk[1]);
                    //Refresh();
                    if (DialogResult.OK == MessageBoxShowOKCancelExclamationDefaultDesktopOnly("發現不正常的分行，是否終止？"))
                    {
                        //textBox1.Select(chk[0] + chk[2] + (chk[0] > 0 ? 2 : 0), 0);//定位以便Enter切行
                        int i = wordsPerLinePara;
                        //while (new StringInfo(CnText.RemovePunctuationsNum(textBox1.Text.Substring(chk[0], i))).LengthInTextElements < wordsPerLinePara)
                        while (CountWordsLenPerLinePara(textBox1.Text.Substring(chk[0], i)) < wordsPerLinePara)
                            i++;
                        if (char.IsLowSurrogate(textBox1.Text.Substring(chk[0] + i, 1).ToCharArray()[0]))
                            textBox1.Select(chk[0] + i + 1, 0);//定位以便Enter切行                        
                        else
                            textBox1.Select(chk[0] + i, 0);//定位以便Enter切行                        
                        return false;
                    }
                    //br.ChromeSetFocus();
                    br.driver.SwitchTo().Window(br.driver.CurrentWindowHandle);
                    textBox1.Select(st, ln);
                    //}
                }
            }
            //if (!newTextBox1(out s, out l, autoPastetoQuickEdit ? x : xCopy))
            //if (!newTextBox1(out s, out l,  x))// 只用 xCopy的話 如《漢籍全文資料庫》的《十三經注疏》輸入就會傳回false
            if (!newTextBox1(out s, out l, autoPaste2QuickEdit ? x : (keyinTextMode ? xCopy : x)))
            {
                if (s != 0 && l != 0 && textBox1.SelectionLength == 0)
                {//若無選取，則將有問題的部分選取以供檢視
                    textBox1.Select(s, l); textBox1.ScrollToCaret();
                }
                AvailableInUseBothKeysMouse(); return false;
            }//在 newTextBox1函式中可能會更動 s、l 二值，故得如此處置，以免s、l值跑掉


            #region 貼到/送至 Ctext Quick edit 
            //根據不同輸入模式需求操作
            switch (browsrOPMode)
            {
                case BrowserOPMode.appActivateByName://預設模式（1）
                    pasteToCtext();
                    break;
                case BrowserOPMode.seleniumNew://純Selenium模式（2）
                                               //終於找到bug了 nextPage()裡的textBox3.Text=url 設定太晚
                    if (url != textBox3.Text) url = textBox3.Text;
                    //if (url.IndexOf("#editor") == -1 && url.IndexOf("&page=") == -1 && url.IndexOf("ctext.org") == -1)
                    string driverUrl = "";
                    try
                    {
                        driverUrl = br.driver.Url;//竟然是Google翻譯的擴充功能在作怪！難怪用selenium一直出錯 ：chrome-extension://aapbdbdomjkkjkaonfhkkikfgjllcleb/offscreen.html   https://www.facebook.com/oscarsun72/posts/pfbid02t9mQG6j2b5GFsbzoMiuUVWASczcPMLnvrGWPHnWQGeoaGo3x234PkymfPdMYSLj4l
                                                  //https://sl.bing.net/cZ2i6BiKAmW
                        if (driverUrl != urlDriver && new Uri(urlDriver).Authority == "ctext.org")
                        {
                            playSound(soundLike.exam, true);
                            for (int i = br.driver.WindowHandles.Count - 1; i > -1; i--)
                            {
                                br.driver.SwitchTo().Window(br.driver.WindowHandles[i]);
                                if (br.driver.Url == urlDriver) { driverUrl = br.driver.Url; break; }
                            }
                        }
                        else if (driverUrl != urlDriver && new Uri(driverUrl).Authority == "ctext.org")
                        {
                            playSound(soundLike.exam, true);
                            urlDriver = driverUrl;
                        }

                        if (driverUrl != urlDriver)
                        {
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("現行分頁似乎有問題！網址應該是：" + urlDriver
                                + Environment.NewLine + "可按【確定】繼續。感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主");
                            Debugger.Break();
                        }
                    }
                    catch (Exception ex)
                    {
                        switch (ex.HResult)
                        {
                            case -2146233088:
                                //"no such window: target window already closed\nfrom unknown error: web view not found\n  (Session info: chrome=111.0.5563.147)"
                                if (ex.Message.IndexOf("no such window: target window already closed") > -1)
                                {
                                    br.GoToUrlandActivate(url);
                                }
                                break;
                            default:
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.Message);
                                Debugger.Break();
                                break;
                        }
                        //throw;
                    }
                    if (!IsValidUrl＿keyDownCtrlAdd(driverUrl))
                    {
                        string brdriverUrl = br.GoToCurrentUserActivateTab();
                        //if (!IsValidUrl＿keyDownCtrlAdd(br.driver.Url))
                        if (!IsValidUrl＿keyDownCtrlAdd(brdriverUrl))
                        //if (br.driver.Url.IndexOf("#editor") == -1 && br.driver.Url.IndexOf("&page=") == -1 && br.driver.Url.IndexOf("ctext.org") == -1)
                        {
                            //Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請檢查 textBox3 中是否是有效的簡單修改模式的網址");
                            if (IsValidUrl＿ImageTextComparisonPage(brdriverUrl))
                            {
                                if (DialogResult.OK == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否要切換到「簡單修改模式」頁面以供輸入？"))
                                {
                                    brdriverUrl = br.GetQuickeditUrl();
                                    br.QuickeditLinkIWebElement.Click();
                                    if (!input2Ctext(brdriverUrl, shiftKeyDownYet)) return false;
                                }
                                else
                                    return false;
                            }
                            else
                                return false;
                        }
                        else
                            //pasteToCtext(br.driver.Url, shiftKeyDownYet);
                            if (!input2Ctext(brdriverUrl, shiftKeyDownYet)) return false;
                    }
                    else
                        if (!input2Ctext(textBox3.Text, shiftKeyDownYet)) return false;//string currentUrl = br.driver.Url;
                                                                                       //pasteToCtext(currentUrl);//故改用 br.……
                    break;
                case BrowserOPMode.seleniumGet://Selenium配合Windows API模式（1+2）或純不用Selenium模式
                                               //還未實作
                    break;
                default:
                    break;
            }
            #endregion

            if (autoPaste2QuickEdit && textBox1.Text == br.TextPatst2Quick_editBox + Environment.NewLine)
            {
                playSound(soundLike.stop, true);
                undoRecord(); PauseEvents();
                textBox1.Clear(); ResumeEvents();
            }

            #region 決定是否要到下一頁
            //if (!shiftKeyDownYet ) nextPages(Keys.PageDown, false);
            if (!shiftKeyDownYet && !check_the_adjacent_pages)
                if (!nextPages(Keys.PageDown, false, notBooksPunctuation, pagePaste2GjcoolOCR))
                    return false;
            #endregion

            #region 預測下一頁頁末尾端在哪裡               
            if (!pagePaste2GjcoolOCR && autoPaste2QuickEdit && !autoPastetoCtextQuitEditTextboxCancel)
            {
                //if (pageTextEndPosition == 0 && pageEndText10 == "" && !keyinText && autoPastetoQuickEdit)
                //{
                //    pageTextEndPosition = textBox1.SelectionStart + textBox1.SelectionLength;
                //    pageEndText10 = textBox1.Text.Substring(pageTextEndPosition, 10);
                //}
                //現在不用剪貼簿了，所以要傳引數以供參考 20241109
                //predictEndofPage(xCopy);                i
                if (!altA_predictEndofPageRange())
                {
                    //pageTextEndPosition = 0; pageEndText10 = "";
                    return false;
                }
                if (fastMode && rePaint) BeginUpdate(true);
            }
            //重設自動判斷頁尾之值(有翻頁就得重設！）
            //pageTextEndPosition = 0; pageEndText10 = "";
            #endregion

            //DialogResult dialogresult = new DialogResult(); 原來在這裡！！！ 20231022
            if (browsrOPMode != BrowserOPMode.appActivateByName && !pagePaste2GjcoolOCR)
            {//使用selenium模式時（非預設模式時）
             //= new DialogResult();
                if (autoPaste2QuickEdit && !keyinTextMode && !autoPastetoCtextQuitEditTextboxCancel)
                {//全自動輸入模式時
                    autoPaste2CtextQuitEditTextbox(out DialogResult dialogresult);//在此中雖有判斷autoPastetoQuickEdit時，然呼叫它會造成無限遞迴（recursion）                    
                    //Debugger.Break();
                    //●●●●●●●●●●●●●20250213

                    //if (dialogresult == DialogResult.Cancel)
                    //    autoPastetoCtextQuitEditTextboxCancel = true;
                    //else
                    //    autoPastetoCtextQuitEditTextboxCancel = false;


                }
                //鍵入輸入模式或非全自動輸入時（如欲瀏覽、或順便編輯時）還原被隱藏的主表單以利後續操作，若不欲，則按Esc鍵即可再度隱藏：20230119壬寅大寒小年夜前一日
                //else if (keyinText || !autoPastetoQuickEdit)
                else if (keyinTextMode && (!autoPaste2QuickEdit || fastMode))// && !pagePaste2GjcoolOCR)
                {
                    if (HiddenIcon) show_nICo(ModifierKeys);
                    //if (!pagePaste2GjcoolOCR) availableInUseBothKeysMouse();
                    AvailableInUseBothKeysMouse();

                    if (ModifierKeys == Keys.Shift)
                    {//自動送交賢超法師《古籍酷AI》OCR
                     //已改寫在 nextPage 裡
                     //Form1.playSound(Form1.soundLike.press);
                     //toOCR(br.OCRSiteTitle.GJcool);
                    }
                    else
                    {
                        //if (!pagePaste2GjcoolOCR)
                        //{
                        //將插入點置於頁首，以備編輯
                        textBox1.Select(0, 0);
                        textBox1.ScrollToCaret();
                        //br.WindowsScrolltoTop();
                        //}
                    }
                }
            }

            #region 如果有「check the adjacent pages」連結控制項則將其點開
            if (keyinTextMode)// && br.CheckAdjacentPages_Linkbox != null)//&& CheckAdjacentPages_DataNext == null)
                br.CheckAdjacentPages_Linkbox?.Click();
            #endregion

            #region 此程序執行完畢，表單顏色閃爍顯示，以供提示 20231029
            ////其實在執行時多數時看不到表單的，也會被其他的樂音遮蔽，故不作！
            //Form1.playSound(soundLike.info);
            //if (!Visible) Visible = true;
            //BringToFront();
            //this.BackColor = Color.Tan;
            //this.Refresh();
            //Thread.Sleep(20);
            //this.BackColor = this.FormBackColorDefault;
            #endregion
            //若停止重畫，則恢復之
            if (!rePaint)
                EndUpdate();

            #region 若是《四庫全書》原書的扉頁直接加入〖文淵|閣寶〗
            if (isSKQSFrontPage(x))
            {
                //int st = Form1.InstanceForm1.textBox1.SelectionStart, len = Form1.InstanceForm1.textBox1.SelectionLength;
                string tx1 = Form1.InstanceForm1.textBox1.Text;
                int wygbStart = tx1.IndexOf("欽定四庫全書<p>");
                if (wygbStart > -1 && !tx1.Contains("〖文淵|閣寶〗"))
                {
                    wygbStart += "欽定四庫全書<p>".Length;
                    //tx1 = tx1.Substring(0, wygbStart)+"〖文淵|閣寶〗<p>"+tx1.Substring(wygbStart);
                    PauseEvents(); stopUndoRec = true;
                    Form1.InstanceForm1.textBox1.Select(wygbStart, 0);
                    Form1.InstanceForm1.textBox1.SelectedText = "〖文淵|閣寶〗<p>";
                    if (int.Parse(br.Svg_image_PageImageFrame.GetAttribute("width")) > 500)
                        br.Svg_image_PageImageFrame.Click();
                    ResumeEvents(); stopUndoRec = false;
                }
                Form1.InstanceForm1.textBox3.Text = textBox3.Text;//主表單網址與本表單連動、同步
                //TopMost = false;
                //br.driver.SwitchTo().Window(driver.CurrentWindowHandle);
                ////WindowState = FormWindowState.Minimized;
                Form1.InstanceForm1.AvailableInUseBothKeysMouse();
                br.WindowHandles["currentPageNum"] = (int.Parse(_currentPageNum) - 1).ToString();
                //Form1.InstanceForm1.textBox1.Select(s, l);
            }
            #endregion

            if (autoPaste2QuickEdit) { pageEndText10 = string.Empty; pageTextEndPosition = 0; }
            return true;
        }
        /// <summary>
        ///還原放大的書圖
        /// </summary>
        internal static void RestoreImageSize()
        {
            OpenQA.Selenium.IWebElement iwe = br.Svg_image_PageImageFrame;
            DateTime dt = DateTime.Now;
            while (iwe == null)
            {
                iwe = br.Svg_image_PageImageFrame;
                if (DateTime.Now.Subtract(dt).TotalSeconds > 2) break;
            }
            string width = iwe?.GetAttribute("width");
            if (width != null)
                if (int.Parse(width) > 500)
                    br.Svg_image_PageImageFrame?.Click();

        }

        /// <summary>
        /// 把作業系統的焦點與游標拉回主表單中
        /// </summary>
        internal void AvailableInUseBothKeysMouse()
        {
            //if (!Active) Activate();//下面bringBackMousePosFrmCenter方法中已調用此方法
            TopMost = true;
            bringBackMousePosFrmCenter();
        }

        const string omitStr = "．‧.…【】〖〗＝{}<p>（）《》〈〉：；、，。「」『』？！0123456789-‧·\r\n";//"　"
        string clearOmitChar(string x)
        {
            foreach (var item in omitStr)
            {
                x = x.Replace(item.ToString(), "");
            }
            return x;
        }

        /// <summary>
        /// 是否是自動連續輸入模式
        /// </summary>
        internal bool AutoPasteToCtext { get { return autoPaste2QuickEdit; } }

        /// <summary>
        /// 是否是自動連續輸入模式
        /// </summary>
        bool autoPaste2QuickEdit = false;
        /// <summary>
        /// 前一本所處理的書籍ID（網址中「&file=」的引數值）以供與現在要處理的作比較，看是不是同一本書（可決定版面特徵是否當予更改）
        /// </summary>
        int previousBookID = 0;
        /// <summary>
        /// 前一部所處理的書籍ID（即URN: ctp:wb728745中的數值）以供與現在要處理的作比較，看是不是同一本書（可決定版面特徵是否當予更改）
        /// 或如 https://ctext.org/wiki.pl?if=en&res=728745 網址中的 res=後面的數值       
        /// </summary>
        int previousResID = 0;
        /// <summary>
        /// 前一個章節(wiki chapter）書籍 chapter ID（即URN: ctp:ws829144中的數值）以供與現在要處理的作比較，看是不是同一本書的同個章節（可決定版面特徵是否當予更改）
        /// 或如 https://ctext.org/wiki.pl?if=en&chapter=829144 網址中的 chapter=後面的數值，
        /// 或 https://ctext.org/library.pl?if=en&file=36583&page=27&editwiki=829144#editor 網址中的 editwiki= 後面的數值
        /// </summary>
        int previousEditwikiID = 0;
        /// <summary>
        /// 記下之前頁數頁碼
        /// </summary>
        string _previousPageNum = string.Empty;


        /// <summary>
        /// 預測準備處理的那一頁，其頁末的位置及其文字
        /// 20250221擬以 altA_predictEndofPageRange 取代
        /// 然此方法純以text屬性操作，不假類別，故可以更高效
        /// </summary>
        void predictEndofPage(string xCopy)
        {
            //if (lines_perPage == 0)
            if (lines_perPage < 6)
            {
                if (AutoPasteToCtext)
                    //現在不用剪貼簿了，所以以引數來判斷 20241109
                    //lines_perPage = (linesParasPerPage != -1 && linesParasPerPage != 0) ? linesParasPerPage : countLinesPerPage(Clipboard.GetText());
                    lines_perPage = (linesParasPerPage != -1 && linesParasPerPage != 0) ? linesParasPerPage : countLinesPerPage(xCopy);
                if (lines_perPage == 0) return;
            }
            string x = textBox1.Text;
            if (x.Length < 30) return;
            string[] xPredict = x.Split(Environment.NewLine.ToArray(), StringSplitOptions.RemoveEmptyEntries);
            //if (xPredict.Length < predictEndofPageSelectedTextLen) return;
            if (x.Replace(Environment.NewLine, "").Replace("　", "") == "") return;
            int s = 0, e = 0, i = 0, predictEndofPagePosition = 0, openBracketS = 0, closeBracketS = 0;
            string item;
            bool openNote = false;//, closeNote=false;
            while (e > -1)
            {
                e = x.IndexOf(Environment.NewLine, s);
                if (e - s < 0 || s < 0) break;
                item = x.Substring(s, e - s); if (item == "") return;
                openBracketS = item.IndexOf("{{"); closeBracketS = item.IndexOf("}}");
                string[] linesParasPage = x.Substring(0, e * 3 <= x.Length ? e * 3 : x.Length).Split(Environment.NewLine.ToArray(), StringSplitOptions.RemoveEmptyEntries);

                if (item == "}}<p>" || (closeBracketS == -1 && openBracketS == 0 && item.Length < 5))//《維基文庫》純注文空及其前一行
                {
                    i++;
                    if (item == "}}<p>") openNote = false; else openNote = true;
                }

                else if (i == 0 && x.IndexOf("}}") > -1 && x.IndexOf("}}") < (x.IndexOf("{{") == -1 ? x.Length : x.IndexOf("{{")) && x.IndexOf("}}") > e)
                { i++; openNote = true; }//第一段/行是純注文
                else if (i == 0 && item.IndexOf("{{") == -1 && item.IndexOf("}}") == -1 && linesParasPage.Length > 1)
                {
                    string xx = linesParasPage[i + 1];
                    if (xx.IndexOf("}}") > -1 && xx.IndexOf("{{") == -1)//&& x.IndexOf("}}") > e)
                    { i++; openNote = true; }//第一段/行是純注文
                    else { i += 2; openNote = false; }//第一段/行是純正文
                }

                else if (i == 0 && ((openBracketS > closeBracketS) ||
                                (openBracketS == -1 && closeBracketS > -1 && closeBracketS < e - 2))) //第一行正、注夾雜
                {
                    if (openBracketS > 2)
                    {
                        i += 2;
                    }
                    else
                    {
                        if (openBracketS == -1) i += 2;
                        else if (openBracketS == 1)
                        {//目前分行分段於有標點者切割有誤差，權以此暫補丁
                            if (omitStr.IndexOf(item.Substring(0, 1)) == -1)
                            {
                                i += 2;
                            }
                            else i++;
                        }
                        else if (openBracketS == 2)
                        {
                            if (omitStr.IndexOf(item.Substring(0, 1)) == -1) i += 2;
                            else if (omitStr.IndexOf(item.Substring(1, 1)) == -1)
                            {//目前分行分段於有標點者切割有誤差，權以此暫補丁
                                i += 2;
                            }
                            else
                            {
                                i++;
                            }
                        }
                    }

                    if (item.LastIndexOf("}}") > item.LastIndexOf("{{"))
                        openNote = false;
                    else
                        openNote = true;

                }

                else if (openBracketS == 0 && closeBracketS == -1)//注文（開始）
                { i++; openNote = true; }
                else if (openBracketS == -1 && openNote)
                {//純注文（末截）
                    if (closeBracketS == item.Length - 2)
                    { i++; openNote = false; }
                    else if (item.Length > 4)
                    {
                        if (item.Substring(item.Length - 5) == "}}<p>") { i++; openNote = false; }
                        else
                        {
                            if (closeBracketS == -1)
                            {
                                if (openNote)
                                    i++;
                                else
                                    i += 2;
                            }
                            else
                            {
                                i += 2;
                                openNote = false;
                            }

                        }
                    }
                }
                else if (openBracketS == -1 && closeBracketS > -1 && closeBracketS < item.Length - 2)
                {//正注夾雜注文結束
                    { i += 2; openNote = false; }
                }
                else if (openBracketS > -1 && item.IndexOf("{{", openBracketS + 2) > -1)//正注夾雜
                {
                    i += 2;
                    if (item.LastIndexOf("}}") < item.LastIndexOf("{{")) openNote = true;
                    else openNote = false;
                }
                else if (openBracketS > -1 && closeBracketS > -1 && closeBracketS < item.Length - 2)//正注夾雜
                {
                    i += 2;
                    if (item.LastIndexOf("}}") < item.LastIndexOf("{{")) openNote = true;
                    else openNote = false;
                }

                //無{{}}標記：
                else if (openBracketS == -1 && closeBracketS == -1)
                {
                    if (openNote == false)//《維基文庫》純正文
                        i += 2;
                    else //《維基文庫》純注文
                        i++;
                }

                //《維基文庫》正注文夾雜
                else if (openBracketS > 0)//正注夾雜
                {
                    if (openBracketS > 2)
                    {
                        i += 2;
                    }
                    else
                    {
                        if (openBracketS == 1)
                        {//目前分行分段於有標點者切割有誤差，權以此暫補丁
                            if (omitStr.IndexOf(item.Substring(0, 1)) == -1)
                            {
                                i += 2;
                            }
                            else i++;
                        }
                        if (openBracketS == 2)
                        {
                            if (omitStr.IndexOf(item.Substring(0, 1)) == -1) i += 2;
                            else if (omitStr.IndexOf(item.Substring(1, 1)) == -1)
                            {//目前分行分段於有標點者切割有誤差，權以此暫補丁
                                i += 2;
                            }
                            else
                            {
                                i++;
                            }
                        }
                    }
                    if (closeBracketS == -1) openNote = true;
                    else
                    {
                        if (item.LastIndexOf("}}") > item.LastIndexOf("{{"))
                            openNote = false;
                        else
                            openNote = true;
                    }

                }
                //else if (openBracketS > 0 && closeBracketS == -1) { i += 2; openNote = true; }
                else if (openBracketS == -1 && closeBracketS > -1 && closeBracketS < item.Length - 2) { i += 2; openNote = false; }
                //else if (item.IndexOf("{{") > 0 || 
                //    (item.IndexOf("}}") > -1 &&
                //    ((item.IndexOf("<p>") == -1 && item.IndexOf("}}") < item.Length - 2) ||
                //        (item.IndexOf("<p>") > -1 && item.IndexOf("}}") < item.Length - 5))
                //    ))
                //    i += 2;
                //else if ((item.IndexOf("{{") == 0 && item.IndexOf("}}") == -1)
                //    || (item.IndexOf("{{") == -1 && (item.Substring(item.Length - 5) == "}}<p>" ||
                //                                item.Substring(item.Length - 2) == "}}")))
                //    //純注文
                //    i++;
                /*
                else if (item.Length > 4 && item.Substring(0, 2) == "{{" && item.Substring(item.Length - 2, 2) == "}}"
                        && item.Substring(2, item.Length - 4).IndexOf("{{") == -1 && item.Substring(2, item.Length - 4).IndexOf("}}") == -1)
                    i++;
                else if (item.Length > 2 
                    && (item.Substring(0, 2) == "{{" && item.IndexOf("}}") == -1 
                        || item.Substring(item.Length - 2, 2) == "}}" && item.IndexOf("{{") == -1))
                    i++;
                */
                else//純正文及注文夾雜者
                    i += 2;
                s = e + 2;
                if (i >= lines_perPage)
                {
                    predictEndofPagePosition = s - 2;
                    break;
                }
            }
            if (predictEndofPagePosition != 0 && predictEndofPagePosition - predictEndofPageSelectedTextLen >= 0)
            {
                textBox1.Select(predictEndofPagePosition - predictEndofPageSelectedTextLen, predictEndofPageSelectedTextLen);
                textBox1.ScrollToCaret();
            }
        }

        /// <summary>
        /// 正常的每頁行數（正文算1，注文算2，故8行正文之文本傳回的值即16）
        /// 起始值（初始化）為0
        /// </summary>
        int lines_perPage = 0;
        /// <summary>
        /// 正常的行/段長度（漢字數）
        /// 起始值（初始化）為0
        /// </summary>
        int normalLineParaLength = 0;

        /// <summary>
        /// 檢查非常長度的行（段）
        /// </summary>
        /// <param name="xChk">這引數是指定要傳入檢查的文本</param>
        /// <returns>若發現非常長度的行，則傳回一個數組（陣列）以表示非常行諸特徵：
        /// { lineSeprtStart（起點）, lineSeprtEnd - lineSeprtStart（非常長度） ,
        ///     normalLineParaLength（通常長度）};
        /// </returns>
        public static int[] CheckAbnormalLinePara(string xChk, int normalLenght)
        {
            string[] xLineParas = xChk.Split(
                    new[] { Environment.NewLine },
                        StringSplitOptions.RemoveEmptyEntries);
            //string[] xLineParas = Regex.Split(xChk, @"\r?\n");//  20250608 GitHub　Copilot大菩薩：	若你要支援不同平台的換行（如 \n、\r），可以用 Regex.Split(xChk, @"\r?\n")。
            foreach (string line in xLineParas)
            {
                int l = Form1.CountWordsLenPerLinePara(line);
                if (l - normalLenght > 4)
                {
                    return new int[] { xChk.IndexOf(line), l, normalLenght };
                }

            }
            return new int[0];

        }

        //20230117 creedit chatGPT大菩薩：C# Visual Studio 註解顯示:/// 是用於多行註解，用於註釋程式碼的多行。……在 C# 中，使用三個斜線 (///) 來撰寫註解文字，並將它放在該函式的宣告之前，就可以在 Visual Studio 中在自訂函式上停駐滑鼠游標時顯示該函式的提示文字。……這樣可以顯示註解文字，且註解文字可以在 Intellisense 中顯示。
        /// <summary>
        /// 檢查非常長度的行（段）
        /// </summary>
        /// <param name="xChk">這引數是指定要傳入檢查的文本</param>
        /// <returns>若發現非常長度的行，則傳回一個數組（陣列）以表示非常行諸特徵：
        /// { lineSeprtStart（起點）, lineSeprtEnd - lineSeprtStart（非常長度） ,
        ///     normalLineParaLength（通常長度）,len（長度）};
        /// </returns>
        private int[] checkAbnormalLinePara(string xChk)
        {

            if (!fastMode) saveText();//備份以防萬一
            string[] xLineParas = xChk.Split(
                Environment.NewLine.ToArray(),
                StringSplitOptions.RemoveEmptyEntries);

            #region get lines_perPage//取得該頁的每行（段）文字
            //lines_perPage = 0;
            if (keyinTextMode)
            {//如果是手動輸入模式：
                lines_perPage = countLinesPerPage(xChk);
                linesParasPerPage = lines_perPage;
                //略過只含有「　」的元素
                xLineParas = xLineParas.Where(x => x.Trim('　') != "").ToArray();
                /* 20230907 Bing大菩薩：
                 太好了，我很高興能幫助您！LINQ 是 Language Integrated Query 的簡稱，它是一種用於查詢和操作數據的語言。在 C# 中，LINQ 可以用來查詢各種數據源，包括數組、列表、XML 文件等。
                    在上面的例子中，我們使用了 LINQ 的 `Where` 方法來過濾出 `xLineParas` 中不符合條件的元素。`Where` 方法接受一個委託作為參數，該委託定義了過濾條件。在這個例子中，我們定義了一個匿名函數 `x => x.Trim('　') != ""` 作為過濾條件。這個函數會對每個元素進行判斷，如果元素去掉前後的「　」後不為空字符串，則返回 `true`，否則返回 `false`。
                    `Where` 方法會根據過濾條件返回一個新的序列，其中只包含符合條件的元素。最後，我們使用 `ToArray` 方法將序列轉換為數組。
                    希望這對您有所幫助！如果您還有其他問題，請隨時告訴我。感恩感恩　讚歎讚歎　南無阿彌陀佛。
                */
            }
            else//不是手動輸入模式：
                lines_perPage = (linesParasPerPage != -1 && linesParasPerPage != 0) ? linesParasPerPage : countLinesPerPage(xChk);
            if (linesParasPerPage == -1) linesParasPerPage = lines_perPage;
            //lines_perPage = xLineParas.Length;
            /*
            foreach (string item in xLineParas)
            {
                if (lines_perPage == 0 & xChk.IndexOf("}}") < xChk.IndexOf("{{") && xChk.IndexOf("}}") > item.Length)
                    lines_perPage++;
                else if (item.Length > 4 && item.Substring(0, 2) == "{{" && item.Substring(item.Length - 2, 2) == "}}"
                        && item.Substring(2, item.Length - 4).IndexOf("{{") == -1 && item.Substring(2, item.Length - 4).IndexOf("}}") == -1)
                    lines_perPage++;
                else if (item.Length > 2 && (item.Substring(0, 2) == "{{" && item.IndexOf("}}") == -1 || item.Substring(item.Length - 2, 2) == "}}" && item.IndexOf("{{") == -1))
                    lines_perPage++;
                else
                    lines_perPage += 2;
            }
            */
            #endregion
            if (NormalLineParaLength == 0)
            {
                if (wordsPerLinePara != -1) NormalLineParaLength = wordsPerLinePara;
                else
                {
                    if (xLineParas.Length > 0)//通常第一行會有卷首篇題，故不準；最末行又可能收尾，故取其次末行
                        NormalLineParaLength = CountWordsLenPerLinePara(xLineParas[xLineParas.Length - 1]);// new StringInfo(xLineParas[0]).LengthInTextElements;
                }
            }

            /////暫時取消此條件，7改成4（即每行3字內，自行目測檢查。）20230822
            //if (normalLineParaLength < 7)
            if (NormalLineParaLength < 4)
            {//如果正常漢字數小於7則不執行
             //normalLineParaLength歸零、wordsPerLinePara歸零
                if (keyinTextMode) { NormalLineParaLength = 0; wordsPerLinePara = -1; }
                return new int[0];
            }

            int i = -1, len = 0;
            foreach (string lineParaText in xLineParas)
            {
                //if (lineParaText.IndexOf("竝當與") > -1) //just for check 
                //    Debugger.Break();


                i++;
                if (lineParaText.IndexOf("{{{") > -1 || lineParaText.IndexOf("孫守真") > -1 || lineParaText.IndexOf("＝") > -1)//{{{孫守真按：}}}、缺字說明等略去，以人工校對
                {
                    continue;
                }
                int noteTextBlendStart = lineParaText.IndexOf("{"),
                    noteTextBlendEnd = lineParaText.IndexOf("}");
                int gap;
                if (noteTextBlendStart != -1 || noteTextBlendEnd != -1)
                {//blend text and note                     
                    string text = "", note = "";
                    if (noteTextBlendStart != -1 && noteTextBlendEnd == -1)
                    {// {{ only
                        text = clearOmitChar(lineParaText.Substring(0, noteTextBlendStart));
                        note = clearOmitChar(lineParaText.Substring(noteTextBlendStart + 2));
                        if (text == "")
                            len = new StringInfo(note).LengthInTextElements;
                        else
                            len = new StringInfo(text).LengthInTextElements +
                                (int)Math.Ceiling((decimal)new StringInfo(note).LengthInTextElements / 2);
                    }
                    if (noteTextBlendStart == -1 && noteTextBlendEnd != -1)
                    {// }} only
                        note = clearOmitChar(lineParaText.Substring(0, noteTextBlendEnd));
                        text = clearOmitChar(lineParaText.Substring(noteTextBlendEnd + 2));
                        if (text == "")
                            len = new StringInfo(note).LengthInTextElements;
                        else
                            len = new StringInfo(text).LengthInTextElements +
                                (int)Math.Ceiling((decimal)new StringInfo(note).LengthInTextElements / 2);
                    }
                    if (noteTextBlendStart != -1 && noteTextBlendEnd != -1)
                    {// {{ and }} both
                        if (noteTextBlendStart == 0 && noteTextBlendEnd + "}}".Length == lineParaText.Length)
                        {
                            note += lineParaText.Substring
                                (noteTextBlendStart + 2,
                                noteTextBlendEnd == -1 ?
                                lineParaText.Length - (noteTextBlendStart + 2)
                                : noteTextBlendEnd - (noteTextBlendStart + 2));
                            note += new String('　', CountWordsLenPerLinePara(note));//單行注文則補上空格以方便計算字數
                            len = CountWordsLenPerLinePara(note) / 2;
                        }
                        else if (noteTextBlendStart < noteTextBlendEnd)
                        {
                            int st = 0, lText = noteTextBlendStart;
                            while (noteTextBlendStart != -1)
                            {
                                text += lineParaText.Substring(st,
                                    lText);
                                note += lineParaText.Substring
                                    (noteTextBlendStart + 2,
                                    noteTextBlendEnd == -1 ?
                                    lineParaText.Length - (noteTextBlendStart + 2)
                                    : noteTextBlendEnd - (noteTextBlendStart + 2));
                                if (CountWordsLenPerLinePara(note) % 2 == 1)
                                    note += "　";//如果一行中有兩處注文以上，可能剛好都缺1字（即均為單數長，又剛好有2的倍數量），造成字數統計上的失誤，如此例：　斗字作斤{{詳前《急就篇》}}與什形近{{《説文·敘》云：人持十為斗。}}此什卽斗字趙
                                                //故補上空白以供計算
                                noteTextBlendStart = lineParaText.IndexOf
                                    ("{", noteTextBlendStart + 2);
                                if (noteTextBlendStart == -1)
                                {
                                    if (noteTextBlendEnd != -1)
                                        text += lineParaText.Substring
                                            (noteTextBlendEnd + 2);
                                    break;
                                }
                                st = noteTextBlendEnd + 2;
                                lText = noteTextBlendStart - st;//(noteTextBlendEnd + 2);
                                if (lText < 0)
                                {
                                    MessageBox.Show("somethins must be wrong,plx check it out !", "", MessageBoxButtons.OK, MessageBoxIcon.Error
                                        , MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                                    return new int[0];
                                }
                                text += lineParaText.Substring
                                    (st, lText);
                                lText = noteTextBlendEnd;
                                noteTextBlendEnd = lineParaText.IndexOf("}",
                                   noteTextBlendStart);
                                if (noteTextBlendEnd == -1)
                                {
                                    note += lineParaText.Substring(
                                        noteTextBlendStart,
                                        lineParaText.Length - noteTextBlendStart);
                                    break;
                                }
                                st = noteTextBlendEnd + 2;
                                lText = noteTextBlendStart;
                                noteTextBlendStart = lineParaText.IndexOf("{", st);
                                if (noteTextBlendStart == -1)
                                {
                                    note += lineParaText.Substring(lText + 2,
                                        noteTextBlendEnd - (lText + 2));
                                    lText = lineParaText.Length - st;
                                    text += lineParaText.Substring(st,
                                        lineParaText.Length - st);
                                    break;
                                }
                                note += lineParaText.Substring(lText + 2, noteTextBlendEnd - (lText + 2));
                                noteTextBlendEnd = lineParaText.IndexOf("}", noteTextBlendStart);
                                lText = noteTextBlendStart - st;
                            }
                            text = clearOmitChar(text); note = clearOmitChar(note);
                            //len = new StringInfo(text).LengthInTextElements + (int)Math.Ceiling((decimal)(new StringInfo(note).LengthInTextElements
                            //    / ((lineParaText.Length - note.Length == 4 && lineParaText.StartsWith("{{") && lineParaText.EndsWith("}}") &&
                            //    new StringInfo(note).LengthInTextElements == normalLineParaLength) ? 1 : 2)));
                            len = new StringInfo(text).LengthInTextElements +
                                (int)Math.Ceiling((decimal)new StringInfo(note).LengthInTextElements
                                / ((lineParaText.StartsWith("{{") && lineParaText.EndsWith("}}") &&
                                new StringInfo(lineParaText.Substring(2, lineParaText.Length - "{{}}".Length)).LengthInTextElements == NormalLineParaLength) ? 1 : 2));
                            //單行小注而字數與正文大字同時，則不折半
                        }
                        else
                        {// noteTextBlendEnd < noteTextBlendStart  
                            int stNote = 0, lNote = noteTextBlendEnd;
                            while (noteTextBlendStart != -1)
                            {
                                note += lineParaText.Substring(stNote, lNote);
                                text += lineParaText.Substring(noteTextBlendEnd + 2,
                                    noteTextBlendStart - (noteTextBlendEnd + 2));
                                noteTextBlendEnd = lineParaText.IndexOf("}",
                                    noteTextBlendStart + 2);
                                stNote = noteTextBlendStart + 2;
                                if (noteTextBlendEnd == -1)
                                {
                                    note += lineParaText.Substring(stNote); break;
                                }
                                else
                                {
                                    note += lineParaText.Substring(stNote, noteTextBlendEnd -
                                        (noteTextBlendStart + 2));
                                    stNote = noteTextBlendStart;//暫記下備用
                                    noteTextBlendStart = lineParaText.IndexOf("{",
                                        noteTextBlendStart + 2);
                                    if (noteTextBlendStart == -1)
                                    {
                                        text += lineParaText.Substring(noteTextBlendEnd + 2);
                                        break;
                                    }
                                    text += lineParaText.Substring(noteTextBlendEnd + 2,
                                        noteTextBlendStart - (noteTextBlendEnd + 2));
                                    stNote = noteTextBlendStart + 2;
                                    lNote = noteTextBlendEnd;
                                    noteTextBlendEnd = lineParaText.IndexOf("}", stNote);
                                    if (noteTextBlendEnd == -1)
                                    {
                                        text += lineParaText.Substring(lNote + 2,
                                            noteTextBlendStart - (lNote + 2));
                                        lNote = lineParaText.Length - stNote;
                                        note += lineParaText.Substring(stNote, lNote);
                                        break;
                                    }
                                    lNote = noteTextBlendEnd - stNote;
                                    noteTextBlendStart = lineParaText.IndexOf("{",
                                        noteTextBlendEnd);
                                    if (noteTextBlendStart == -1)
                                    {
                                        text += lineParaText.Substring(
                                            noteTextBlendEnd + 2);
                                    }
                                }
                            }
                            text = clearOmitChar(text); note = clearOmitChar(note);
                            len = new StringInfo(text).LengthInTextElements +
                                (int)Math.Ceiling((decimal)
                                new StringInfo(note).LengthInTextElements / 2);
                        }
                    }
                    gap = Math.Abs(len - NormalLineParaLength);
                }
                else//only text or note
                {
                    len = CountWordsLenPerLinePara(lineParaText.EndsWith("<p>") ? lineParaText.Substring(0, lineParaText.Length - "<p>".Length) : lineParaText);
                    //len = new StringInfo(clearOmitChar(lineParaText)).LengthInTextElements;
                    if ((xChk.IndexOf(lineParaText) + lineParaText.Length + lineParaText.Length <= xChk.Length
                        && xChk.Substring(xChk.IndexOf(lineParaText) + lineParaText.Length, "<p>".Length) == "<p>") ||
                        lineParaText.EndsWith("<p>"))
                        gap = 0;
                    else
                        gap = Math.Abs(len - NormalLineParaLength);
                }

                //誤差容錯值
                const int gapRef = 0;//3;//9;

                //the normal rule
                if (gap > gapRef && !(len < NormalLineParaLength
                    && lineParaText.IndexOf("<p>") > -1)
                    && lineParaText != "　" && lineParaText.IndexOf("*") == -1 &&
                        lineParaText.IndexOf("|") == -1) //&& gap < 8)
                {//select the abnormal one
                    bool alarm = true;
                    if (i + 1 < xLineParas.Length)
                    {
                        if (gap > gapRef && len < NormalLineParaLength
                            && xLineParas[i + 1].IndexOf("}}") > -1
                            && CountWordsLenPerLinePara(xLineParas[i + 1]) < NormalLineParaLength)
                        //&& xChk.IndexOf(lineParaText) + lineParaText.Length - 1 > 0
                        //&& xChk.Substring(xChk.IndexOf(lineParaText) + lineParaText.Length , "<p>".Length) == "<p>")
                        {
                            alarm = false;
                        }
                    }
                    if (alarm)
                    {
                        string x = textBox1.Text;
                        int j = -1, lineSeprtEnd = 0, lineSeprtStart = lineSeprtEnd;//Seprt=Separate
                        lineSeprtEnd = x.IndexOf(Environment.NewLine, lineSeprtEnd);
                        while (lineSeprtEnd > -1)
                        {

                            if (++j == i) break;
                            lineSeprtStart = lineSeprtEnd;
                            lineSeprtEnd = x.IndexOf(Environment.NewLine, ++lineSeprtEnd);
                        }
                        if (gap > 10)
                        {
                            SystemSounds.Hand.Play();
                        }
                        return new int[] { lineSeprtStart, (lineSeprtEnd==-1?x.Length:lineSeprtEnd) - lineSeprtStart ,
                        NormalLineParaLength,len};

                    }
                }
            }
            return new int[0];
            //throw new NotImplementedException();
        }


        /*
         * https://copilot.microsoft.com/shares/uUY14uhuKY6qdvTXpAM3o
         * 20250301 Copilot大菩薩：要在C# Windows.Forms中實現選取textBox1中的文字而不觸發重畫，可以嘗試使用TextBox的BeginUpdate和EndUpdate方法來凍結和解凍控制項。雖然TextBox本身沒有這樣的方法，但你可以通過封裝這些方法來減少重畫。
         */
        bool rePaint = true;
        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int wMsg, IntPtr wParam, IntPtr lParam);
        //private const int WM_SETREDRAW = 0x0b;
        /// <summary>
        /// 停止重畫並隱藏表單
        /// </summary>
        /// <param name="hideForm">指定是否要隱藏表單，預設為隱藏</param>
        public void BeginUpdate(bool hideForm = true)
        {
            if (hideForm && Visible)
            {
                Visible = false;
                if (!IsDriverInvalid()) driver.SwitchTo().Window(driver.CurrentWindowHandle);
                br.BringToFront("chrome");
                br.ChromeSetFocus();
            }
            rePaint = false;
            SendMessage(this.Handle, WM_SETREDRAW, (IntPtr)0, IntPtr.Zero);//通過使用SendMessage來控制WM_SETREDRAW消息，可以暫時禁用或啟用重畫。
        }
        /// <summary>
        /// 恢復重畫並顯示表單
        /// </summary>
        public void EndUpdate()
        {
            if (!_eventsEnabled) _eventsEnabled = true;
            if (!Visible)
            {
                Visible = true;
                this.Show();
                Activate();
                //show_nICo(Keys.None);//會影響表單大小及位置
            }
            rePaint = true;
            SendMessage(this.Handle, WM_SETREDRAW, (IntPtr)1, IntPtr.Zero);
            this.Refresh();
            if (browsrOPMode != BrowserOPMode.appActivateByName && !br.IsDriverInvalid())
                br.WindowsScrolltoTop();
        }

        /// <summary>
        /// 記下遞迴的ws對話方塊傳回的參數 20250213
        /// 停止遞迴（recursion），則設為true 20250222
        /// 預設為true。
        /// </summary>
        bool autoPastetoCtextQuitEditTextboxCancel = true;
        /// <summary>
        /// Ctrl + F2 切換語音操作（預設為非 Windows 語音辨識操作）識別用
        /// </summary>
        bool speechRecognitionOPmode = false;
        /// <summary>
        /// 直接貼到CTP的Quick edit的文字方塊中
        /// 進行自動連續輸入實作的主要函式方法
        /// </summary>
        /// <param name="dialogResult"></param>
        /// <returns>結束或失敗、沒操作則傳回false</returns>
        bool autoPaste2CtextQuitEditTextbox(out DialogResult dialogResult)
        {
            if (autoPastetoCtextQuitEditTextboxCancel || textBox1.TextLength == 0)
            {
                //autoPastetoQuickEdit = false;
                dialogResult = DialogResult.Cancel;
                ////autoPastetoCtextQuitEditTextboxCancel = false;
                return false;
            }


            ////if (new StringInfo(textBox1.SelectedText).LengthInTextElements == predictEndofPageSelectedTextLen &&
            ////        textBox1.Text.Substring(textBox1.SelectionStart + textBox1.SelectionLength, 2) == Environment.NewLine)
            dialogResult = DialogResult.Cancel;

            //感覺ModifierKeys屬性似會不靈光，按下Ctrl鍵後似乎會黏住而彈不起來，故行此：20250215
            if (ModifierKeys == Keys.Control) SendKeys.Send("^");

            bool keyDownCtrlAdd_ReturnVale = false;
            if (!autoPastetoCtextQuitEditTextboxCancel
                && textBox1.SelectionStart + textBox1.SelectionLength + 2 <= textBox1.TextLength
                && textBox1.Text.Substring(textBox1.SelectionStart + textBox1.SelectionLength, 2) == Environment.NewLine)
            //這是舊的、原來的 predictEndofPage方法要參照的： && textBox1.SelectionLength == predictEndofPageSelectedTextLen )
            {
                if (autoPaste2QuickEdit)
                {
                    //        //if (MessageBox.Show("auto paste to Ctext Quit Edit textBox?" + Environment.NewLine + Environment.NewLine
                    //        //    + "……" + textBox1.SelectedText, "", MessageBoxButtons.OKCancel,MessageBoxIcon.Question,
                    //        //    MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly) == DialogResult.OK)
                    //        if (MessageBox.Show("auto paste to Ctext Quit Edit textBox?" + Environment.NewLine + Environment.NewLine
                    //            + "……" + textBox1.SelectedText, "", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                    //            //獨占式訊息
                    //            MessageBoxOptions.DefaultDesktopOnly)
                    //                == DialogResult.OK)
                    //        {
                    //            if (autoPastetoQuickEdit && (ModifierKeys == Keys.Control || check_the_adjacent_pages))
                    //            {
                    //                appActivateByName();
                    //                //Browser.driver == null)
                    //                if (browsrOPMode == BrowserOPMode.appActivateByName)
                    //                    //當啟用預估頁尾後，按下 Ctrl 或 Shift Alt 可以自動貼入 Quick Edit ，唯此處僅用 Ctrl 及 Shift 控制關閉前一頁所瀏覽之 Ctext 網頁                
                    //                    SendKeys.Send("^{F4}");//關閉前一頁
                    //                if (check_the_adjacent_pages) nextPages(Keys.PageDown, false);
                    //            }
                    //            keyDownCtrlAdd(false);
                    //        }
                    //20230113 creedit with chatGPT：設定訊息方塊獨占性：
                    bool _autoPastetoQuickEdit = autoPaste2QuickEdit;
                    bool _check_the_adjacent_pages = check_the_adjacent_pages;

                    bool doNotShowMsgBox = false;
                    //按著Ctrl鍵則直接ok 20250109（並啟動快捷模式）
                    //if (ModifierKeys == Keys.Control && dialogResult != DialogResult.Abort) { dialogResult = DialogResult.OK; goto ok; }

                    if ((ModifierKeys == Keys.Control && keycodeNow != Keys.Subtract
                                && !KeyboardInfo.getKeyStateToggled(System.Windows.Input.Key.Add)
                                && !KeyboardInfo.getKeyStateToggled(System.Windows.Input.Key.Subtract)
                                && !KeyboardInfo.getKeyStateDown(System.Windows.Input.Key.LeftCtrl)
                                && !KeyboardInfo.getKeyStateDown(System.Windows.Input.Key.RightCtrl))
                        || Control.IsKeyLocked(Keys.CapsLock)
                        // 檢查 Caps Lock 狀態
                        || KeyboardInfo.getKeyStateToggled(System.Windows.Input.Key.CapsLock)
                        || KeyboardInfo.getKeyStateDown(System.Windows.Input.Key.CapsLock)
                        //|| KeyboardInfo.getKeyStateDown(System.Windows.Input.Key.LeftShift)//因為多工時易遭誤按，故改用Caps鍵就好
                        || Control.IsKeyLocked(Keys.CapsLock)
                        || IsCapsLockOn())
                    {
                        if (KeyboardInfo.getKeyStateToggled(System.Windows.Input.Key.CapsLock)
                            || KeyboardInfo.getKeyStateDown(System.Windows.Input.Key.CapsLock)
                            //|| (KeyboardInfo.getKeyStateDown(System.Windows.Input.Key.LeftShift)
                            //    && !KeyboardInfo.getKeyStateDown(System.Windows.Input.Key.Insert)//非在複製貼上
                            //    && !KeyboardInfo.getKeyStateDown(System.Windows.Input.Key.Delete)
                            //    && !KeyboardInfo.getKeyStateDown(System.Windows.Input.Key.Left)//非在做選取時
                            //    && !KeyboardInfo.getKeyStateDown(System.Windows.Input.Key.Right)
                            //    && !KeyboardInfo.getKeyStateDown(System.Windows.Input.Key.Up)
                            //    && !KeyboardInfo.getKeyStateDown(System.Windows.Input.Key.Down))//因為多工時易遭誤按，故改用Caps鍵就好
                            || Control.IsKeyLocked(Keys.CapsLock)
                            || IsCapsLockOn())//GitHub　Copilot大菩薩：要實現根據 Caps Lock 燈的狀態來執行 FastModeSwitcher 方法，我們可以使用 Control.IsKeyLocked 方法來檢查 Caps Lock 燈的狀態。這個方法可以直接檢查 Caps Lock 燈是否亮著。……這樣可以確保在 Caps Lock 燈亮時觸發 FastModeSwitcher 方法，而不需要按住 Caps Lock 鍵。
                        {
                            if (fastMode)
                                FastModeSwitcher();
                        }
                        else
                            if (keycodeNow != Keys.Subtract &&
                                !KeyboardInfo.getKeyStateToggled(System.Windows.Input.Key.Add) &&
                                !KeyboardInfo.getKeyStateDown(System.Windows.Input.Key.Add) &&
                                KeyboardInfo.getKeyStateNone(System.Windows.Input.Key.Add) &&
                                KeyboardInfo.getKeyStateNone(System.Windows.Input.Key.Subtract) &&
                                !KeyboardInfo.getKeyStateDown(System.Windows.Input.Key.Subtract) &&
                                !KeyboardInfo.getKeyStateToggled(System.Windows.Input.Key.Subtract))
                            FastModeSwitcher();
                        //if (!FastMode) autoPastetoCtextQuitEditTextboxCancel = true;
                    }

                    if (fastMode)
                    { dialogResult = DialogResult.OK; doNotShowMsgBox = true; goto ok; }
                    else if (ModifierKeys == Keys.Control)
                    { dialogResult = DialogResult.OK; doNotShowMsgBox = true; goto ok; }


                    #region 取得最後不含 <p> 與 。<p> 的5個字來顯示於訊息方塊中20250118
                    string textbox1Text = textBox1.Text;
                    StringInfo si = new StringInfo(new Document(textbox1Text).Range(0, textBox1.SelectionStart).Text.Replace("*", string.Empty).Replace("<p>", string.Empty));
                    string last5Characters = si.SubstringByTextElements(si.LengthInTextElements - 5);
                    //string last5Characters = textBox1.SelectedText; int eLast5Characters = last5Characters.IndexOf("<p>");
                    int eLast5Characters = last5Characters.IndexOf("<p>");
                    if (eLast5Characters > -1)
                    {
                        if (eLast5Characters - 1 > -1 && last5Characters.Substring(eLast5Characters - 1, 1) == "。")//如果是「。<p>」
                            eLast5Characters--;
                        StringInfo siLast5Characters = new StringInfo(last5Characters.Replace("。", string.Empty).Replace("<p>", string.Empty));
                        int sLast5Characters = textBox1.Text.IndexOf(last5Characters); eLast5Characters = sLast5Characters + last5Characters.Length;
                        while (sLast5Characters > -1 && eLast5Characters - sLast5Characters > -1
                            && siLast5Characters.LengthInTextElements < 5)
                        {
                            if (sLast5Characters - 1 < 0) break;
                            siLast5Characters = new StringInfo(textBox1.Text.Substring(--sLast5Characters, eLast5Characters - sLast5Characters).Replace("。", string.Empty).Replace("<p>", string.Empty));
                        }
                        last5Characters = siLast5Characters.String;
                    }
                    #endregion

                    if (!fastMode) Refresh();

                    if (int.Parse(Form1.InstanceForm1.CurrentPageNum) != Form1.InstanceForm1.GetPageNumFromUrl(driver.Url))//!CheckPageNumBeforeSubmitSaveChanges(driver))
                    {
                        MessageBoxShowOKCancelExclamationDefaultDesktopOnly("頁碼不同！請檢查！！");
                        autoPastetoCtextQuitEditTextboxCancel = true;
                        dialogResult = DialogResult.Cancel;
                        return false;
                    }

                    //if (!autoPastetoCtextQuitEditTextboxCancel && ModifierKeys == Keys.None && (!speechRecognitionOPmode || !FastMode))
                    if (ModifierKeys == Keys.None && (!speechRecognitionOPmode || !fastMode))
                    {
                        dialogResult = MessageBox.Show("auto paste to Ctext Quit Edit textBox?" + Environment.NewLine + Environment.NewLine
                                                + "……" + Environment.NewLine + Environment.NewLine +
                                                last5Characters, "現在處理第" + (
                                                _check_the_adjacent_pages ? (int.Parse(_currentPageNum) + 1).ToString() : CurrentPageNum)
                                                 + "頁", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                                 MessageBoxOptions.DefaultDesktopOnly);
                    }
                    else
                    {//非語音操作時
                     //if (!autoPastetoCtextQuitEditTextboxCancel && ModifierKeys == Keys.None && !FastMode)
                        if (ModifierKeys == Keys.None && !fastMode)
                            dialogResult = MessageBox.Show("auto paste to Ctext Quit Edit textBox?" + Environment.NewLine + Environment.NewLine
                                                    + "……" + Environment.NewLine + Environment.NewLine +
                                                    last5Characters, "現在處理第" + (
                                                    _check_the_adjacent_pages ? (int.Parse(_currentPageNum) + 1).ToString() : CurrentPageNum)
                                                    + "頁", MessageBoxButtons.OKCancel, MessageBoxIcon.Question
                                                     );
                    }

                ok:
                    if (dialogResult == DialogResult.OK)
                    {
                        //如按下確定時同時也按下Ctrl鍵，也可以啟動快捷模式
                        if ((KeyboardInfo.getKeyStateDown(System.Windows.Input.Key.LeftCtrl) || ModifierKeys == Keys.Control)
                            && (!KeyboardInfo.getKeyStateDown(System.Windows.Input.Key.Add) && !KeyboardInfo.getKeyStateDown(System.Windows.Input.Key.Subtract))
                            && !fastMode) FastModeSwitcher();
                        //if (!doNotShowMsgBox && (KeyboardInfo.getKeyStateDown(System.Windows.Input.Key.LeftCtrl) || ModifierKeys == Keys.Control) && !fastMode) FastModeSwitcher();

                        if (browsrOPMode == BrowserOPMode.appActivateByName) textBox1.Enabled = false;//避免誤觸
                        if (_autoPastetoQuickEdit && (ModifierKeys == Keys.Control || _check_the_adjacent_pages))
                        {
                            if (browsrOPMode == BrowserOPMode.appActivateByName)
                            {
                                appActivateByName();
                                //關閉瀏覽器分頁
                                SendKeys.Send("^{F4}");
                            }
                            if (_check_the_adjacent_pages) nextPages(Keys.PageDown, false);

                        }
                        if (fastMode && rePaint) BeginUpdate();
                        keyDownCtrlAdd_ReturnVale = keyDownCtrlAdd(false);
                        //if (fastMode) EndUpdate();
                        if (!keyDownCtrlAdd_ReturnVale && doNotShowMsgBox == true) { dialogResult = DialogResult.Cancel; WindowsScrolltoTop(); goto ok; }//dialogRes-ult = DialogResult.Abort;
                        doNotShowMsgBox = false;
                        //if (browsrOPMode != BrowserOPMode.appActivateByName && !autoPastetoCtextQuitEditTextboxCancel)
                        if (browsrOPMode != BrowserOPMode.appActivateByName)
                        {//if (autoPastetoQuickEdit) 會在autoPastetoCtextQuitEditTextbox()中判斷
                         //預估下一頁尾位置
                         //predictEndofPage();//在前面keyDownCtrlAdd(false);已做一次，這次做是給遞迴（recursion）用的「if (textBox1.SelectionLength == predictEndofPageSelectedTextLen &&……」這行要判斷                            
                            autoPaste2CtextQuitEditTextbox(out DialogResult dialogresult);//遞迴（recursion） 20230113
                                                                                          //if (dialogresult == DialogResult.Cancel)
                                                                                          //{
                                                                                          //    dialogResult = dialogresult;//●●●●●●●●●●●●●●●●●●●元宵後一日
                                                                                          //    goto ok;
                                                                                          //}
                        }
                    }

                    //取消自動輸入時
                    else
                    {
                        autoPastetoCtextQuitEditTextboxCancel = true;
                        keyDownCtrlAdd_ReturnVale = false;
                        if (fastMode) EndUpdate();

                        //如果是鄰近頁牽連編輯，則自動翻到下一頁書圖以備檢覆
                        if (_check_the_adjacent_pages)
                        {
                            nextPages(Keys.PageDown, false);
                        }
                        else//如果不是鄰近頁牽連編輯
                        {
                            if (browsrOPMode != BrowserOPMode.appActivateByName)
                            {//決定是否清除原在quickedit_data_textbox裡的文字
                                switch (quickedit_data_textboxtxt)
                                {
                                    //如果是Word VBA 新頁面所產生的 tab鍵 \t（讀取值是「 」，非"\t"）
                                    case " ":// br.chkClearQuickedit_data_textboxTxtStr:
                                        BringToFront();
                                        dialogResult = MessageBox.Show("是否清除當前頁面中的空白內容？（其實是有由tab鍵所按下的值）", "",
                                        MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                        MessageBoxOptions.ServiceNotification);
                                        //if (DialogResult.OK == dialogResult && !autoPastetoCtextQuitEditTextboxCancel)
                                        if (DialogResult.OK == dialogResult)
                                        {

                                            #region 以下是據方法函式「keyDownCtrlAdd(bool shiftKeyDownYet = false)」而來
                                            input2Ctext(textBox3.Text, false, br.chkClearQuickedit_data_textboxTxtStr);
                                            //if (!textBox1.Enabled) { textBox1.Enabled = true; textBox1.Focus(); }
                                            //Task.WaitAll(); Thread.Sleep(500);
                                            nextPages(Keys.PageDown, false);
                                            ////預測下一頁頁末尾端在哪裡
                                            //predictEndofPage();
                                            ////重設自動判斷頁尾之值
                                            //pageTextEndPosition = 0; pageEndText10 = "";
                                            #endregion
                                            autoPaste2CtextQuitEditTextbox(out DialogResult dialogresult);//在此中自會判斷autoPastetoQuickEdit值
                                        }
                                        break;
                                    case ""://如果文字框裡沒內容（即空白頁）
                                        BringToFront();
                                        if (br.pageUBound > int.Parse(br.Page_textbox.GetAttribute("value")))
                                        {
                                            dialogResult = MessageBox.Show("是否移到下一頁？", "",
                                                MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                                MessageBoxOptions.DefaultDesktopOnly);
                                            //if (DialogResult.OK == dialogResult && !autoPastetoCtextQuitEditTextboxCancel)
                                            if (DialogResult.OK == dialogResult)
                                            {
                                                nextPages(Keys.PageDown, false);
                                                autoPaste2CtextQuitEditTextbox(out DialogResult dialogresult);
                                            }
                                        }
                                        else
                                        {
                                            dialogResult = DialogResult.Cancel;
                                            if (File.Exists("C:\\Windows\\Media\\ring07.wav"))
                                                using (SoundPlayer sp = new SoundPlayer("C:\\Windows\\Media\\ring07.wav")) { sp.Play(); }
                                            else
                                                Form1.playSound(Form1.soundLike.waiting, true);
                                        }
                                        break;
                                    default:
                                        break;
                                }
                            }
                        }
                        //避免誤觸
                        if (browsrOPMode != BrowserOPMode.appActivateByName) textBox1.Enabled = false;
                        if (keyDownCtrlAdd_ReturnVale && dialogResult == DialogResult.OK)
                            pageTextEndPosition = textBox1.SelectionStart + predictEndofPageSelectedTextLen;
                        if (textBox1.TextLength > 0 && textBox1.TextLength - pageTextEndPosition > 0)
                        {
                            pageEndText10 = textBox1.Text.Substring(pageTextEndPosition > 9 ? pageTextEndPosition - 10 : pageTextEndPosition,
                                                                textBox1.TextLength - pageTextEndPosition >= 10 ? 10
                                                                    : textBox1.TextLength - pageTextEndPosition);//終於抓到這個bug了，忘了加第2個參數
                        }
                        #region 連續輸入終止時的插入點位置
                        if (textBox1.TextLength > 0 && !keyDownCtrlAdd_ReturnVale
                            && textBox1.TextLength >= pageTextEndPosition && textBox1.Text.Substring(0, pageTextEndPosition).IndexOf("/") > -1)
                            textBox1.Select(textBox1.Text.Substring(0, pageTextEndPosition).IndexOf("/"), insertMode ? 1 : 0);
                        //else
                        //    textBox1.Select(pageTextEndPosition, 0);
                        #endregion
                        driver.SwitchTo().Window(driver.CurrentWindowHandle);
                        //焦點交回表單
                        #region 解除 textbox防觸鎖定，並準備檢視編輯；如果訊息方塊不是在取得 Cancel回應值時即關閉，則此下程式恐怕要移出這個 if else區塊才行
                        //Activate();已會觸發Form1_Activated(new object(), new EventArgs());事件
                        //訊息方塊就是一個表單，顯示時會讓此表單失去焦點。
                        if (!Active)
                        {
                            bringBackMousePosFrmCenter();
                        }
                        //解除 textbox防觸鎖定，並準備檢視編輯，交給上一行Activate();處理
                        if (!textBox1.Enabled) { textBox1.Enabled = true; textBox1.Focus(); }
                        #endregion

                        return false; //20250221

                    }//取消自動輸入時
                }
                else
                {
                    dialogResult = DialogResult.Cancel;
                    keyDownCtrlAdd_ReturnVale = keyDownCtrlAdd(false);
                    return false;//注意會不會造成無窮遞迴
                }
            }
            else
            {//不符合自動連續輸入執行的條件時：
                if (!autoPastetoCtextQuitEditTextboxCancel)
                {
                    autoPastetoCtextQuitEditTextboxCancel = true;
                    br.driver.SwitchTo().Window(br.driver.CurrentWindowHandle);
                    br.ChromeSetFocus();
                    AvailableInUseBothKeysMouse();
                    if (fastMode)
                    {
                        if (br.pageUBound == int.Parse(br.Page_textbox.GetAttribute("value")))
                        {
                            if (File.Exists("C:\\Windows\\Media\\ring07.wav"))
                                using (SoundPlayer sp = new SoundPlayer("C:\\Windows\\Media\\ring07.wav")) { sp.Play(); }
                            else
                                Form1.playSound(Form1.soundLike.waiting, true);
                        }
                        else
                            Form1.playSound(Form1.soundLike.waiting, true);
                    }

                    return false;//20250301●●●●●●●●●●●●●●●●●
                }
                else//20250301●●●●●●●●●●●●●●●●●
                {//會影響自動讀如下一卷文本的機制，故須在有內容時（尚有內容待輸入時）才觸動
                    br.driver.SwitchTo().Window(br.driver.CurrentWindowHandle);
                    if (!Active && textBox1.Text != string.Empty) AvailableInUseBothKeysMouse();
                    //autoPastetoCtextQuitEditTextboxCancel = false;
                    if (fastMode)
                    {
                        if (br.pageUBound == int.Parse(br.Page_textbox.GetAttribute("value")))
                        {
                            if (File.Exists("C:\\Windows\\Media\\ring07.wav"))
                                using (SoundPlayer sp = new SoundPlayer("C:\\Windows\\Media\\ring07.wav")) { sp.Play(); }
                            else
                                Form1.playSound(Form1.soundLike.waiting, true);
                        }
                        else
                            Form1.playSound(Form1.soundLike.waiting, true);
                    }

                    return false;
                }
            }

            //20250221
            return true;
        }

        /// <summary>
        /// 將滑鼠位置帶回主表單中心
        /// </summary>
        private void bringBackMousePosFrmCenter()
        {
            if (browsrOPMode != BrowserOPMode.appActivateByName && driver != null
                && !IsDriverInvalid() && br.IsConfirmHumanPage()) return;
            if (this.InvokeRequired)
            {
                this.Invoke((MethodInvoker)delegate
                {
                    // 你的程式碼
                    if (!Visible) Visible = true;
                    if (!rePaint) EndUpdate();
                    BringToFront();
                    Activate(); //Application.DoEvents();
                                //if (!this.TopMost) this.TopMost = true;

                    // 判斷滑鼠游標是否在表單內 20230824Bing大菩薩
                    if (this.Bounds.Contains(Cursor.Position)) return;
                    //20230115 chatGPT大菩薩：Cursor back to form：
                    Point formPos = new Point(this.Location.X + this.Size.Width / 2, this.Location.Y + this.Size.Height / 2);
                    Cursor.Position = formPos;
                    //上面這段程式碼將滑鼠游標的位置設置為表單的中心點位置。
                    //Point cursorPos = Cursor.Position;
                    //this.Location = cursorPos;
                });
            }
            else
            {
                if (!Visible) Visible = true;
                if (!rePaint) EndUpdate();
                BringToFront();
                Activate(); //Application.DoEvents();
                            //if (!this.TopMost) this.TopMost = true;

                // 判斷滑鼠游標是否在表單內 20230824Bing大菩薩
                if (this.Bounds.Contains(Cursor.Position)) return;
                //20230115 chatGPT大菩薩：Cursor back to form：
                Point formPos = new Point(this.Location.X + this.Size.Width / 2, this.Location.Y + this.Size.Height / 2);
                Cursor.Position = formPos;
                //上面這段程式碼將滑鼠游標的位置設置為表單的中心點位置。
                //Point cursorPos = Cursor.Position;
                //this.Location = cursorPos;
            }
        }

        bool autoPasteFromSBCKwhether = false;
        void autoPasteFromSBCK(bool autoPasteFromSBCKwhether)
        {
            string x = textBox1.Text, xClipboard = Clipboard.GetText();
            if (!autoPasteFromSBCKwhether) return;
            if (x.IndexOf(xClipboard) > -1) return;
            if (!TopMost) TopMost = true;
            textBox1.Text += xClipboard;
            textBox1.Select(textBox1.TextLength, 0);
            textBox1.ScrollToCaret();
            //每分鐘自動備份
            TimeSpan timeSpan = new TimeSpan();
            timeSpan = DateTime.Now.Subtract(new FileInfo(FName_to_Save_Txt_fullname).LastWriteTime);
            if (timeSpan.TotalMinutes > 1) saveText();
            new SoundPlayer(@"C:\Windows\Media\windows default.wav").Play();
        }

        const int predictEndofPageSelectedTextLen = 5;
        void splitLineParabySeltext(Keys kys)
        {
            if (!(kys == Keys.F1 || kys == Keys.Pause) || ModifierKeys != Keys.None) return;
            if (kys == Keys.F1)
            {
                autoPaste2CtextQuitEditTextbox(out DialogResult dialogResult);
                return;
            }
            string x = textBox1.SelectedText;
            if (x == "") return;
            x = textBox1.Text;
            int s = textBox1.SelectionStart, l = textBox1.SelectionLength;
            if (kys == Keys.Pause)
                // -按下 Pause Break 鍵：以找到的字串位置** 後**分行分段
                x = x.Substring(0, s + l) + Environment.NewLine + x.Substring(s + l);
            if (kys == Keys.F1)
                //- 按下 F1 鍵：以找到的字串位置**前**分行分段
                x = x.Substring(0, s) + Environment.NewLine + x.Substring(s);
            if (!textBox1.Focused) textBox1.Focus();
            undoRecord();
            stopUndoRec = true;
            textBox1.Text = x;
            stopUndoRec = false;
            textBox1.SelectionStart = s; textBox1.SelectionLength = l;
            textBox1.ScrollToCaret();
        }
        private void selToNewline(ref int s, ref int ed, string x, bool forward, TextBox tBox)
        {
            if (forward)
            {
                for (int i = s + 1; i + 1 < x.Length; i++)
                {
                    if (x.Substring(i, 2) == Environment.NewLine)
                    {
                        ed = i + 2;
                        break;
                    }
                }
            }
            else
            {
                for (int i = s - 1; i > -1; i--)
                {
                    if (x.Length < i + 2) break;
                    if (x.Substring(i, 2) == Environment.NewLine)
                    {
                        s = i + 2;
                        break;
                    }
                    if (i == 0)
                    {
                        s = i;
                    }
                }
            }
            if (s > -1 && ed - s > 0)
            {
                tBox.SelectionStart = s; tBox.SelectionLength = ed - s;
                tBox.ScrollToCaret();
            }


        }


        private void textBox3_Click(object sender, EventArgs e)
        {
            string x = "";
            Task.WaitAny();
            Application.DoEvents();
            try
            {
                if (Clipboard.ContainsText())
                {
                    x = Clipboard.GetText();
                }
            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    //"要求的剪貼簿作業失敗。"
                    case -2147221040://chatGPT 20230108                        
                        x = Task.Run(() => Clipboard.GetText()).Result;
                        break;
                    default:
                        throw;
                }
            }
            if (x == "" || x.Length < 4 || x == textBox3.Text)
            {
                inputSKQSFrontPage(); return;
            }
            if (x.Substring(0, 4) == "http")
                if (x.IndexOf("ctext.org") > -1)
                {
                    if (textBox3.Text != x)
                        textBox3.Text = x;
                    //SystemSounds.Beep.Play();
                    //Form1.playSound(Form1.soundLike.processing);
                    Form1.playSound(Form1.soundLike.done); if (TopMost) TopMost = false;
                    //if (browsrOPMode != BrowserOPMode.appActivateByName) br.GoToUrlandActivate(textBox3.Text);
                    textBox1.Focus();
                }

            inputSKQSFrontPage();


        }


        #region 如果是在自動連續輸入下
        string inputTextFrontPage = "{{{封面}}}<p>";

        /// <summary>
        /// 如果是在自動連續輸入下，自動輸入《四庫全書》扉頁資訊到Quick edit中 20250221
        /// </summary>
        void inputSKQSFrontPage()
        {
            if (autoPaste2QuickEdit)
            {
                if (!IsDriverInvalid() && driver.Url.Contains("&chapter="))
                {
                    GraphicMatchingPagesLink?.Click();//圖文對照圖標按鈕元件
                    textBox3.Text = driver.Url;
                }
                //如果是在第1頁
                if (CurrentPageNum != string.Empty && int.Parse(CurrentPageNum) != 1 && (br.driver.Url.EndsWith("&page=1") || br.driver.Url.Contains("&page=1&")))
                {
                    textBox3.Text = driver.Url;
                    string url = textBox3.Text;
                    if (url.IndexOf("&page=") > -1)
                    {
                        int s = url.IndexOf("&page=") + "&page=".Length;
                        _currentPageNum = url.Substring(s, url.IndexOf("&", s) > -1 ? url.IndexOf("&", s) - s : url.Length - s);
                    }
                }
                //if (textBox3.Text.IndexOf("&page=1") > -1)// && textBox3.Text.EndsWith("#editor"))
                if (CurrentPageNum == "1")// && textBox3.Text.EndsWith("#editor"))
                    if (!IsDriverInvalid())
                    {
                        br.driver.Navigate().GoToUrl(textBox3.Text);
                        if (br.QuickeditLinkIWebElement != null)
                        {
                            //const string inputText = "《四庫全書》􏿽{{經部　}}<p>";
                            //const string inputText = "《四庫全書》􏿽{{史部　}}<p>";
                            //const string inputText = "《四庫全書》􏿽{{子部　}}<p>";
                            //const string inputText = "《四庫全書》􏿽{{集部　}}<p>";
                            //const string inputText = "《小　倦　遊　閣　集》<p>";
                            //const string inputText = "《三　才　圖　會》<p>";
                            //const string inputTextFrontPage = "《帶　經　堂　詩　話》<p>";                            
                            br.QuickeditLinkIWebElement.Click();
                            PauseEvents();
                            textBox3.Text = driver.Url;
                            ResumeEvents();
                            br.driver.SwitchTo().Window(driver.CurrentWindowHandle);
                            //設定Quick edit textBox
                            DateTime dt = DateTime.Now;
                            while (br.Quickedit_data_textbox == null) { if (DateTime.Now.Subtract(dt).TotalSeconds > 5) break; }
                            string text = br.Quickedit_data_textboxTxt;
                            if (text == string.Empty || text == "●" || text == "●\t" || text == "●<p>")
                            {
                                br.SetIWebElementValueProperty(br.Quickedit_data_textbox, inputTextFrontPage);
                                br.SavechangesButton?.Click();//送出
                                nextPages(Keys.PageDown, false);//翻到下一頁
                                string clpTxt = Clipboard.GetText();
                                if (textBox1.TextLength == 0 && clpTxt.Length > 0
                                    &&
                                    (clpTxt.IndexOf("a]") > -1 || clpTxt.IndexOf("a] ") > -1 ||
                                        (!clpTxt.Contains("感恩感恩　讚歎讚歎　南無阿彌陀佛") && clpTxt.IndexOf("P") > -1
                                        && int.TryParse(clpTxt.Substring(clpTxt.IndexOf("P") + 1, 1), out _)))//P 乃「北京元引科技有限公司《元引科技引得數字人文資源平臺·中國歷代文獻》」的文本特徵
                                    )
                                {
                                    //autoExecuteSKQSContextMark();
                                    runWordMacro("中國哲學書電子化計劃.國學大師_Kanripo_四庫全書本轉來");
                                    AutoMarkTitleParagraph();
                                }
                            }
                        }
                    }
            }
        }
        #endregion

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            //if ((int)ModifierKeys == (int)Keys.Control+ (int)Keys.Shift&&e.KeyCode==Keys.C)
            //https://bbs.csdn.net/topics/350010591
            //https://zhidao.baidu.com/question/628222381668604284.html
            var m = ModifierKeys;

            #region 同時按下 Ctrl Shift Alt

            if (e.Control && e.Shift && e.Alt && e.KeyCode == Keys.T)// 20240718
                                                                     //{//ctrl + shift + alt + t ： 測試專用
                                                                     //{//ctrl + alt + shift + t ： 測試專用
            {
                e.Handled = true;
                playSound(soundLike.exam, true);
                br.HanchiTextReadinginPagebyPage();

                //string x = textBox1.Text;
                //CnText.ClearFirstParaLeadingSpace(ref x);
                //textBox1.Text = x;

                //MessageBoxShowOKExclamationDefaultDesktopOnly(CountWordsLenPerLinePara(GetLineText(textBox1.Text, textBox1.SelectionStart)).ToString() + " 字長");


                //AutoMarkTitleParagraph();

                //if (!IsDriverInvalid())
                //br.CopySKQSNextVolume();

                //expandSelectedTextRangeToWholeLinePara(textBox1.SelectionStart,textBox1.SelectionLength,textBox1.Text);
                //Range range = new Document(ref textBox1).Range(textBox1.SelectionStart, textBox1.SelectionStart + textBox1.SelectionLength);
                //outdent_ConvexRow(ref range);
                //undoRecord();
                //PauseEvents();                
                //textBox1.SelectedText = range.Text;
                //ResumeEvents();

                return;
            }
            #endregion

            #region 同時按下 Ctrl Shift
            if ((m & Keys.Control) == Keys.Control
                && (m & Keys.Shift) == Keys.Shift
                && e.KeyCode == Keys.W)
            {
                e.Handled = true;
                closeChromeWindow();//Ctrl + Shift + w 關閉 Chrome 網頁視窗
                return;
            }
            if ((m & Keys.Control) == Keys.Control
                && (m & Keys.Shift) == Keys.Shift
                && e.KeyCode == Keys.C)
            {
                e.Handled = true;
                Clipboard.SetText(textBox1.Text);
                return;
            }
            if ((m & Keys.Control) == Keys.Control
                && (m & Keys.Shift) == Keys.Shift && e.KeyCode == Keys.Divide)
            {//按下 Ctrl + Shift + / （Divide）  切換 check_the_adjacent_pages 值
                e.Handled = true;
                toggleCheck_the_adjacent_pages();
                return;
            }
            if ((m & Keys.Control) == Keys.Control
                && (m & Keys.Shift) == Keys.Shift)
            {

                if (e.KeyCode == Keys.D4)
                {
                    e.Handled = true;
                    //Ctrl + Shift + 4 ：翻到[Kanripo](https://www.kanripo.org/)下一卷在複製其文本後即執行 Word VBA Sub 巨集指令「國學大師_Kanripo_四庫全書本轉來」
                    //runWordMacro("國學大師_Kanripo_四庫全書本轉來");
                    textBox1.Clear();
                    br.SikuQuanshu_SKQSContextCopyReading();
                    return;
                }
                //按下 ctrl + shift + * （數字鍵盤的星號）  toggle keyinTextmode 切換手動鍵入模式
                if (e.KeyCode == Keys.Multiply)
                {
                    e.Handled = true;
                    KeyinTextmodeSwitcher(true);
                    return;
                }
                if (e.KeyCode == Keys.Subtract)
                {//按下 Ctrl + Shift + - (數字鍵盤） ： 切換OCR輸入模式（直接連續輸入）
                    e.Handled = true;

                    if (!_eventsEnabled) _eventsEnabled = true;
                    if (confirm_that_you_are_human) confirm_that_you_are_human = false;
                    if (ocrTextMode)
                    {
                        new SoundPlayer(@"C:\Windows\Media\Speech Off.wav").Play();
                        autoTitleMark_OCRTextMode = false; PagePaste2GjcoolOCR_ing = false;
                        ocrTextMode = false; return;
                        //if (BatchProcessingGJcoolOCR) BatchProcessingGJcoolOCR = false; return;
                    }
                    new SoundPlayer(@"C:\Windows\Media\Speech On.wav").Play();
                    //設定成手動OCR輸入模式，自動及全部覆蓋之貼上則設成false
                    ocrTextMode = true; PasteOcrResultFisrtMode = true; keyinTextMode = true; pasteAllOverWrite = false; autoPaste2QuickEdit = false;
                    PagePaste2GjcoolOCR_ing = false;

                    string pagePast2OCRsiteName = string.Empty;
                    switch (PagePast2OCRsite)
                    {
                        case br.OCRSiteTitle.GoogleKeep:
                            pagePast2OCRsiteName = "Google Keep";
                            break;
                        case br.OCRSiteTitle.GJcool:
                            pagePast2OCRsiteName = "《古籍酷》";
                            break;
                        case br.OCRSiteTitle.KanDianGuJi:
                            pagePast2OCRsiteName = "《看典古籍》OCR網頁版";
                            break;
                        case br.OCRSiteTitle.KanDianGuJiAPI:
                            pagePast2OCRsiteName = "《看典古籍》OCR API";
                            break;
                        default:
                            break;
                    }
                    if (MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否要自動標識標題，在OCR識讀匯入後", pagePast2OCRsiteName + "──要送去OCR的網站") == DialogResult.OK)
                    {
                        autoTitleMark_OCRTextMode = true;
                        linesParasPerPage = countLinesPerPage(textBox1.Text);
                    }
                    else autoTitleMark_OCRTextMode = false;
                    if (!BatchProcessingGJcoolOCR)
                    {
                        if (MessageBoxShowOKCancelExclamationDefaultDesktopOnly("要《古籍酷》批量處理嗎？") == DialogResult.OK)
                            BatchProcessingGJcoolOCR = true;
                    }

                    button1.Text = "分行分段";
                    button1.ForeColor = new System.Drawing.Color();//預設色彩 預設顏色 https://stackoverflow.com/questions/10441000/how-to-programmatically-set-the-forecolor-of-a-label-to-its-default
                    return;
                }
                if (e.KeyCode == Keys.Oem5)
                {//Ctrl + Shift + \ 切換抬頭平抬格式設定（bool TopLine）
                    e.Handled = true;
                    if (!TopLine) SystemSounds.Hand.Play();
                    else SystemSounds.Exclamation.Play();
                    TopLine = !TopLine;
                    return;
                }
                if (e.KeyCode == Keys.Oem3)
                {//Ctrl + Shift + ` 切換OBS開始串流和停止串流時（這是我於OBS所設定的快捷鍵，可同時觸發）
                    e.Handled = true;
                    YAKCSwitchr();
                    return;
                }
            }

            //Ctrl + Shift + t 同Chrome瀏覽器 --還原最近關閉的頁籤
            if ((m & Keys.Control) == Keys.Control && (m & Keys.Shift) == Keys.Shift && e.KeyCode == Keys.T)
            {
                e.Handled = true;
                appActivateByName();
                SendKeys.Send("^+t");
                return;

            }
            //Ctrl + Shift + o 執行《看典古籍》OCR API
            if (e.Control && e.Shift && e.KeyCode == Keys.O)
            {
                //await PerformOCR();
                e.Handled = true;
                playSound(soundLike.press, true);
                toOCR(br.OCRSiteTitle.KanDianGuJiAPI);
                AvailableInUseBothKeysMouse();
                return;
            }
            //Ctrl + Shift + p ： 逐頁瀏覽肉眼檢查空白頁，以免白跑OCR 20240727 執行 CheckBlankPagesBeforeOCR
            if (e.Control && e.Shift && e.KeyCode == Keys.P)
            {
                e.Handled = true;
                string url = string.Empty;
                if (br.driver == null) return;
                try
                {
                    url = br.driver.Url;
                }
                catch (Exception)
                {
                    //Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                    //return;
                    if (br.IsWindowHandleValid(br.driver, br.LastValidWindow))
                        br.driver.SwitchTo().Window(br.LastValidWindow);
                }
                //檢查是否是可操作的頁面（分頁）
                if (!IsValidUrl＿ImageTextComparisonPage(url))
                {
                    foreach (var item in br.driver.WindowHandles)
                    {
                        br.driver.SwitchTo().Window(item);
                        url = br.driver.Url;
                        if (IsValidUrl＿ImageTextComparisonPage(url))
                            if (MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否是這個頁面？") == DialogResult.OK) break;
                    }
                    if (!IsValidUrl＿ImageTextComparisonPage(url))
                    {
                        MessageBoxShowOKExclamationDefaultDesktopOnly("請開啟要瀏覽檢查的頁面？"); return;
                    }
                }
                if (url != textBox3.Text) textBox3.Text = url;

                int stopPageNum = br.pageUBound;

                CheckBlankPagesBeforeOCR_NextPage(url, int.Parse(br.Page_textbox.GetAttribute("defaultValue")), stopPageNum);
                return;
            }

            //Ctrl + Shift + F1：選取範圍前後加上{{}}並清除分行/段符號
            if (e.Control && e.Shift && e.KeyCode == Keys.F1)
            {
                e.Handled = true;
                if (insertMode && textBox1.SelectionLength > 0 || !insertMode)
                {
                    undoRecord();
                    stopUndoRec = true;
                    if (textBox1.SelectionLength == 0)
                        overtypeModeSelectedTextSetting(ref textBox1);
                    string x = textBox1.SelectedText;
                    //x = "{{" + x.Replace(Environment.NewLine, "") + "}}".Replace("{{{{","{{").Replace("}}}}","}}");
                    x = "{{" + x.Replace(Environment.NewLine, "") + "}}";
                    CnText.CurlybracesFormalizer(ref x);
                    textBox1.SelectedText = x;
                    //清除後續的分行/段符號及1個全形空格
                    int s = textBox1.SelectionStart, l = textBox1.SelectionLength;
                    if (textBox1.TextLength > s + l + 2)
                    {
                        string nextStr = textBox1.Text.Substring(s + l, 2);
                        if (nextStr == Environment.NewLine)
                        {

                            stopUndoRec = true; PauseEvents();
                            textBox1.Select(s + l, 2); textBox1.SelectedText = string.Empty;
                            s = textBox1.SelectionStart;
                            if (s + 2 < textBox1.TextLength && textBox1.Text.Substring(s, 1) == "　")
                            {
                                if (textBox1.Text.Substring(s + 1, 1) != "　")
                                    textBox1.Select(s + l, 1); textBox1.SelectedText = string.Empty;
                            }
                            stopUndoRec = false; ResumeEvents();
                        }
                        else if (nextStr.StartsWith("　"))
                        {
                            if (nextStr.Substring(1, 1) != "　")
                            {
                                stopUndoRec = true; PauseEvents();
                                textBox1.Select(s + l, 1);
                                textBox1.SelectedText = string.Empty;
                                stopUndoRec = false; ResumeEvents();
                            }

                        }
                    }
                    stopUndoRec = false;
                    undoRecord();
                }
                else if (keyinTextMode && textBox1.SelectionLength == 0)
                {//>若無選取範圍時就執行 Alt + Shift + w 之功能 20250131大年初三 感恩感恩　讚歎讚歎　GitHub Copilot大菩薩　南無阿彌陀佛　讚美主
                    new Document(ref textBox1).MergeParagraphsAtCaretWithShift();
                }
                return;
            }

            //Ctrl + Shift + n 或 Shift + F1 : 開新Form1 實例
            if (((m & Keys.Control) == Keys.Control && (m & Keys.Shift) == Keys.Shift && e.KeyCode == Keys.N)
                //|| ((m & Keys.Shift) == Keys.Shift && e.KeyCode == Keys.F1))
                || (e.Shift && !e.Alt & e.Control && e.KeyCode == Keys.F1))
            {
                e.Handled = true;
                newForm1();
                return;
            }
            //以上 Ctrl + Shift
            #endregion

            #region Ctrl + Alt
            if ((m & Keys.Control) == Keys.Control && (m & Keys.Alt) == Keys.Alt)
            {
                //Ctrl + Alt + i 檢查IP現狀
                if (e.KeyCode == Keys.I)
                {
                    e.Handled = true;
                    SystemSounds.Exclamation.Play();
                    br.IPStatusMessageShow(out Tuple<bool, bool, bool, bool, DateTime> ipStatus, string.Empty, false, true);
                    if (Clipboard.GetText() != br.CurrentIP) Clipboard.SetText(br.CurrentIP);
                    bringBackMousePosFrmCenter();
                    return;
                }
                if (e.KeyCode == Keys.O)
                {//Ctrl + Alt + o :下載圖片，交給Google Keep OCR
                    if (browsrOPMode == BrowserOPMode.appActivateByName) return;
                    e.Handled = true; Form1.playSound(Form1.soundLike.press);
                    TopMost = false;
                    OpenQA.Selenium.IWebElement iw = br.WaitFindWebElementBySelector_ToBeClickable("#content");
                    if (iw != null) // clickCopybutton_GjcoolFastExperience(iw.Location); 
                        Cursor.Position = (Point)iw.Location;
                    toOCR(br.OCRSiteTitle.GoogleKeep);
                    return;
                }
                if (e.KeyCode == Keys.R)
                {//Ctrl + Alt + r :將如《趙城金藏》3欄式的版面書圖《古籍酷》AI服務OCR結果重新排列 
                 //if (browsrOPMode == BrowserOPMode.appActivateByName) return;
                    e.Handled = true; Form1.playSound(Form1.soundLike.press);
                    string x = textBox1.Text;
                    undoRecord();
                    CnText.Rearrangement3ColumnLayout(ref x);
                    textBox1.Text = x;
                    return;
                }
                if (e.KeyCode == Keys.F10)
                {//Ctrl + Alt + f10： 將textBox1中選取的文字送去《古籍酷》自動標點。若無選取則將整個textBox1的內容送去。（略去其他檢查，唯小於20字元不處理）20240910
                    if (browsrOPMode != BrowserOPMode.seleniumNew)
                    {
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請先以Selenium模式啟動Chrome瀏覽器（在textBox2中輸入「br」或「bb」就可以了），再繼續。");
                        return;
                    }
                    e.Handled = true; Form1.playSound(Form1.soundLike.press);
                    undoRecord(); PauseEvents();
                    if (!textBox1.Focused) textBox1.Focus();
                    this.toGjcoolPunct("https://gj.cool/punct", true, true);
                    ResumeEvents();
                    return;
                }

            }

            #endregion

            #region 按下 Alt+ Shift
            if ((m & Keys.Alt) == Keys.Alt && (m & Keys.Shift) == Keys.Shift)
            {
                if (e.KeyCode == Keys.F1)
                {//Alt + Shift + F1 ：切換 textbox1 之字型： 切換支援 CJK - Ext 擴充字集的大字集字型
                    var cjk = getCJKExtFontInstalled(CJKBiggestSet[++FontFamilyNowIndex]);
                    if (FontFamilyNowIndex == CJKBiggestSet.Length - 1) FontFamilyNowIndex = -1;
                    if (cjk != null)
                    {
                        MessageBox.Show(cjk.Name);
                        if (cjk.Name == "KaiXinSongB")
                        {
                            textBox1.Font = new Font(cjk, (float)17);
                        }
                        else
                        {
                            textBox1.Font = new Font(cjk, textBox1FontDefaultSize);
                        }
                    }
                    e.Handled = true; return;
                }

                if (e.KeyCode == Keys.D)
                {//Alt + Shift + d : 下載當前頁面書圖
                    e.Handled = true;
                    Form1.playSound(Form1.soundLike.press, true);
                    string f = MydocumentsPathIncldBackSlash + "CtextTempFiles\\Ctext_Page_Image.png", mspaint = "C:\\WINDOWS\\system32\\mspaint.exe";
                    if (File.Exists(f))
                        File.Delete(f);
                    toOCR(br.OCRSiteTitle.KanDianGuJi, true);

                    if (File.Exists(mspaint))
                    {
                        if (MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否要以小畫家開啟編輯？") == DialogResult.OK)
                        {
                            Process.Start(mspaint);
                            Thread.Sleep(2000);
                            SendKeys.SendWait("^o");
                            SendKeys.SendWait(f + "~");
                        }
                    }
                    return;
                }


                if (e.KeyCode == Keys.K)
                {// Alt + Shift + k ：下載書圖並交給《看典古籍》OCR（網頁版）
                    e.Handled = true; Form1.playSound(Form1.soundLike.press, true);
                    TopMost = false;
                    OpenQA.Selenium.IWebElement iw = br.WaitFindWebElementBySelector_ToBeClickable("#content");
                    if (iw != null) // clickCopybutton_GjcoolFastExperience(iw.Location); 
                        Cursor.Position = (Point)iw.Location;
                    toOCR(br.OCRSiteTitle.KanDianGuJi);
                    return;
                }
                if (e.KeyCode == Keys.O)
                {//Alt + Shift + o ：交給《古籍酷》 OCR ，模擬使用者手動操作的功能（測試成功！！！！）
                    if (PagePaste2GjcoolOCR_ing) return;
                    if (browsrOPMode == BrowserOPMode.appActivateByName) return;
                    if (!IsValidUrl＿ImageTextComparisonPage(textBox3.Text)) return;
                    e.Handled = true; Form1.playSound(Form1.soundLike.press, true);
                    TopMost = false;
                    OpenQA.Selenium.IWebElement iw = br.WaitFindWebElementBySelector_ToBeClickable("#content");
                    if (iw != null) // clickCopybutton_GjcoolFastExperience(iw.Location); 
                        Cursor.Position = (Point)iw.Location;
                    toOCR(br.OCRSiteTitle.GJcool);
                    AvailableInUseBothKeysMouse();
                    //if (browsrOPMode == BrowserOPMode.appActivateByName) return;
                    //e.Handled = true;
                    //string imgUrl = Clipboard.GetText(), downloadImgFullName;
                    //if (imgUrl.Length > 4
                    //    && imgUrl.Substring(0, 4) == "http"
                    //    && imgUrl.Substring(imgUrl.Length - 4, 4) == ".png")
                    //    downloadImage(imgUrl, out downloadImgFullName);
                    //else
                    //{
                    //    imgUrl = br.GetImageUrl();
                    //    downloadImage(imgUrl, out downloadImgFullName);
                    //}

                    //#region toOCR
                    //bool ocrResult = br.OCR_GJcool_AutoRecognizeVertical(downloadImgFullName);
                    //if (!ocrResult) MessageBox.Show("請重來一次；重新執行一次。感恩感恩　南無阿彌陀佛", "", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    //if (!Active) bringBackMousePosFrmCenter();
                    //#region 如果是手動鍵入輸入模式且OCR程序無誤則直接貼上結果並自動標上書名號篇名號，20230309 creedit with chatGPT大菩薩：
                    //if (ocrResult && keyinTextMode)
                    //{ //Form1_KeyDown(sender, new KeyEventArgs(Keys.Alt & Keys.Insert));
                    //  // 建立 Keys.Alt + Keys.Insert 的組合鍵
                    //    Keys comboKey = Keys.Alt | Keys.Insert;//在 C# 中，要表示兩個按鍵的組合鍵，需要使用 "|" 運算子進行位元運算，而不是 "&" 或 "+" 運算子。 "|" 運算子可以將兩個按鍵的 KeyCode 合併成一個整數，表示按下這兩個按鍵的組合鍵。
                    //                                           // 使用 SendKeys 方法觸發按下組合鍵
                    //    SendKeys.Send("{" + comboKey + "}");
                    //}
                    //#endregion
                    ////const string gjcool = "https://gj.cool/try_ocr";
                    ////Process.Start(gjcool);
                    //#endregion
                    return;
                }

                if (e.KeyCode == Keys.P)
                {//Alt + Shift + p
                    e.Handled = true;
                    keysParagraphSymbol(true); return;
                }


                if (e.KeyCode == Keys.V)
                {//Alt + Shift + v ：新增一個直書的文字方塊
                 //20230824 chatGPT大菩薩：中文文字方塊直書示例:根本就是失敗的。沒有用。沒變成直書，且亂排了一通。如果能截圖給您看，我就截給您看了。感恩感恩　南無阿彌陀佛
                    e.Handled = true; Form1.playSound(Form1.soundLike.press);
                    //AddVerticalTextBox();
                    return;
                }

                if (e.KeyCode == Keys.F12)
                {//Alt + Shift + F12 ：
                    e.Handled = true;
                    BackupLastPageText(Clipboard.GetText(), true, true);// 更新最後的備份頁文本
                    return;
                }

            }
            #endregion

            #region 按下Ctrl鍵
            if (Control.ModifierKeys == Keys.Control)
            {//按下Ctrl鍵

                // Ctrl + F2 切換語音操作（預設為非 Windows 語音辨識操作）識別用
                if (e.KeyCode == Keys.F2)
                {
                    e.Handled = true;
                    speechRecognitionOPSwitchr();
                    return;
                }

                if (e.KeyCode == Keys.F10)
                {//Ctrl + F10： 將textBox1中選取的文字送去《古籍酷》舊版自動標點。若無選取則將整個textBox1的內容送去。（小於20字元不處理）20240808（臺灣父親節）
                    if (browsrOPMode != BrowserOPMode.seleniumNew)
                    {
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請先以Selenium模式啟動Chrome瀏覽器（在textBox2中輸入「br」或「bb」就可以了），再繼續。");
                        return;
                    }
                    e.Handled = true;
                    if (textBox1.SelectedText.Length < 5) textBox1.DeselectAll();
                    toGjcoolPunct("https://old.gj.cool/gjcool/index");
                    return;
                }
                if (e.KeyCode == Keys.F11)
                {//Ctrl + F11： 將textBox1中選取的文字送去《古籍酷》舊版自動標點。若無選取則將整個textBox1的內容送去。（小於20字元不處理）20240808（臺灣父親節）
                    if (browsrOPMode != BrowserOPMode.seleniumNew)
                    {
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請先以Selenium模式啟動Chrome瀏覽器（在textBox2中輸入「br」或「bb」就可以了），再繼續。");
                        return;
                    }
                    e.Handled = true;
                    if (textBox1.SelectedText.Length < 5) textBox1.DeselectAll();
                    toGjcoolPunct("https://old.gj.cool/gjcool/index");
                    return;
                }
                if (e.KeyCode == Keys.F)
                {
                    e.Handled = true;
                    textBox2.Focus();
                    textBox2.SelectionStart = 0; textBox2.SelectionLength = textBox2.Text.Length;
                    return;
                }

                if (e.KeyCode == Keys.R)
                {//Ctrl + r ：刷新目前 Chrome瀏覽器 或 預設瀏覽器 網頁（同於網頁上按下F5鍵）；因為有時瀏覽器網頁未完整開啟必須重載
                    e.Handled = true;
                    appActivateByName();
                    SendKeys.Send("{F5}");
                    //textBox1.Focused; ⇒此為取得在桌面的焦面，不是表單中的焦點，故當表單並非桌面中作用中的表單時，此值也會是fasle，因為TextBox是在表單Form容器中的子元件故 20230125癸卯年初四
                    if (ActiveControl.Name == "textBox1" && !Active) { AvailableInUseBothKeysMouse(); }
                    return;
                    /* 20230125 孫守真 × chatGPT大菩薩：
                     * 老師菩薩新年吉祥，煩請老師菩薩慈悲指導，我這想法是否正確，如下程式碼，我想要在喚起預設瀏覽器後判斷若textBox1是表單中有焦點的控制項，則我想把桌面的焦點再拉回我的表單中，其中的程式碼註解就是我的想法，焦點的取得與判斷，是否是像我想的這樣呢？……
                     * 您的想法是正確的。程式碼中，當使用者按下Ctrl + r鍵時會啟動預設瀏覽器並刷新網頁，然後判斷textBox1是否為表單中有焦點的控制項。如果是，則調用availableInUseBothKeysMouse()方法將桌面的焦點拉回表單中。
                     * ActiveControl.Name== "textBox1" 可以取得目前表單中有焦點的控制項
                     * Active 屬性判斷目前表單是否是作用中的表單。
                     * 如果您的目的是在啟動預設瀏覽器後，將桌面的焦點拉回表單中，那麼程式碼應該是能夠達到您的需求。
                     * 孫守真
                     * 確實是執行正常，達到需求，也謝謝老師菩薩幫我印證我的想法是正確的。感恩感恩　南無阿彌陀佛
                     * 我很高興能夠幫助您。如果您有任何其他問題，歡迎隨時問我。南無阿彌陀佛，祝您新年吉祥。
                     */
                }

                if (e.KeyCode == Keys.PageDown || e.KeyCode == Keys.PageUp)
                {//Ctrl + PageDown Ctrl + PageUP
                    e.Handled = true;//取得或設定值，指出是否處理事件。https://docs.microsoft.com/zh-tw/dotnet/api/system.windows.forms.keyeventargs.handled?view=netframework-4.7.2&f1url=%3FappId%3DDev16IDEF1%26l%3DZH-TW%26k%3Dk(System.Windows.Forms.KeyEventArgs.Handled);k(TargetFrameworkMoniker-.NETFramework,Version%253Dv4.7.2);k(DevLang-csharp)%26rd%3Dtrue
                    nextPages(e.KeyCode, true);
                    if (browsrOPMode != BrowserOPMode.appActivateByName) driver.SwitchTo().Window(driver.CurrentWindowHandle);
                    //if (autoPastetoQuickEdit) AvailableInUseBothKeysMouse();
                    AvailableInUseBothKeysMouse();
                    return;
                }
                if (e.KeyCode == Keys.Left || e.KeyCode == Keys.Right)
                {//Ctrl+左右鍵：徵調
                    if (textBox1.Focused) return;
                    e.Handled = true;
                    Point mouseIP = Cursor.Position;//https://jjnnykimo.pixnet.net/blog/post/27155696
                    if (e.KeyCode == Keys.Left) { this.Left--; mouseIP.X--; }//-= w; }
                    if (e.KeyCode == Keys.Right) { this.Left++; mouseIP.X++; } //+= w; }
                    Cursor.Position = mouseIP;
                    //SetCursorPos(mouseIP.X,mouseIP.Y);
                    return;
                }

                if (e.KeyCode == Keys.D1)
                {//Ctrl + 1 ：執行 Word VBA Sub 巨集指令「漢籍電子文獻資料庫文本整理_以轉貼到中國哲學書電子化計劃」【 附件即有 [Word VBA](https://github.com/oscarsun72/TextForCtext/tree/master/WordVBA) 相關模組 】
                 //現在少用，故以此機制防制：
                    if (DialogResult.OK == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("執行「漢籍電子文獻資料庫文本整理_以轉貼到中國哲學書電子化計劃」？"))
                        runWordMacro("漢籍電子文獻資料庫文本整理_以轉貼到中國哲學書電子化計劃");
                    e.Handled = true; return;
                }
                if (e.KeyCode == Keys.D3)
                {//Ctrl + 3 ：執行 Word VBA Sub 巨集指令「漢籍電子文獻資料庫文本整理_十三經注疏」
                    runWordMacro("漢籍電子文獻資料庫文本整理_十三經注疏");
                    e.Handled = true; return;
                }
                if (e.KeyCode == Keys.D4)
                {//Ctrl + 4 ：執行 Word VBA Sub 巨集指令「維基文庫四部叢刊本轉來」                 
                    e.Handled = true;
                    runWordMacro("維基文庫四部叢刊本轉來");
                    return;
                }

                if (e.KeyCode == Keys.N)
                {// Ctrl + n ：開新預設瀏覽器視窗 //原：在新頁籤開啟 google 網頁，以備用（在預設瀏覽器為 Chrome 時）
                    e.Handled = true;
                    //Process.Start("https://www.google.com.tw/?hl=zh_TW");
                    appActivateByName();
                    SendKeys.Send("^n");
                    this.Activate();
                    appActivateByName();
                    e.Handled = true; return;
                }

                if (e.KeyCode == Keys.S) { e.Handled = true; saveText(); return; }

                if (e.KeyCode == Keys.W)
                {//Ctrl + w 關閉 Chrome 網頁頁籤
                    e.Handled = true;
                    //主表單才執行 20241224
                    if (Name == "Form1")
                        closeChromeTab();
                    else
                        Close();
                    return;
                }

                if (e.KeyCode == Keys.Multiply)
                {//按下 Ctrl + * 設定為將《四部叢刊》資料庫所複製的文本在表單得到焦點時直接貼到 textBox1 的末尾,或反設定
                    e.Handled = true;
                    //避免誤按
                    if (!autoPaste2QuickEdit && !fastMode)
                    {
                        toggleAutoPasteFromSBCKwhether();
                    }
                    return;
                }

                if (e.KeyCode == Keys.Divide)
                {//按下 Ctrl + / （Divide） 切換自動連續輸入功能
                    e.Handled = true;
                    toggleAutoPastetoQuickEdit();
                    return;
                }


            }//按下 Ctrl鍵 終
            #endregion


            //if (((m & Keys.Control) == Keys.Control && (m & Keys.Alt) == Keys.Alt) && 
            //    (e.KeyCode == Keys.Left || e.KeyCode == Keys.Right||e.KeyCode==Keys.Menu))
            //{
            //    const int w= 10;
            //    if (e.KeyCode == Keys.Left) this.Left -= w;
            //    if (e.KeyCode == Keys.Right) this.Left += w;
            //    return;
            //}

            #region 按下Shift鍵
            if (Control.ModifierKeys == Keys.Shift)
            {//按下Shift鍵                
                if (e.KeyCode == Keys.F9)
                {//Shift + F9 ：重啟小小輸入法
                    e.Handled = true;
                    Process.Start(dropBoxPathIncldBackSlash + @"VS\bat\重啟小小輸入法.bat");
                    AvailableInUseBothKeysMouse();
                    return;
                }
                if (e.KeyCode == Keys.F10)
                {//Shift + F10 ： 執行 Word VBA Sub 巨集指令「中國哲學書電子化計劃_只保留正文注文_且注文前後加括弧_貼到古籍酷自動標點」
                    e.Handled = true;
                    //SystemSounds.Question.Play();
                    string x = Clipboard.GetText();
                    if (x == "") { x = textBox1.Text; textBox1.Clear(); }
                    if (MessageBox.Show("clear the Environment.NewLine?", "", MessageBoxButtons.OKCancel) == DialogResult.OK)
                    {
                        x = x.Replace(Environment.NewLine, "");
                        Clipboard.SetText(x);
                    }
                    runWordMacro("Docs.中國哲學書電子化計劃_只保留正文注文_且注文前後加括弧_貼到古籍酷自動標點");
                    return;
                }
                if (e.KeyCode == Keys.F12)
                {// Shift + F12
                    e.Handled = true;
                    saveText();
                    return;
                }
            }//按下Shift鍵 終
            #endregion

            #region 按下Alt鍵
            if (Control.ModifierKeys == Keys.Alt)
            {//按下Alt鍵

                if (e.KeyCode == Keys.PageUp || e.KeyCode == Keys.PageDown)
                {
                    //if (browsrOPMode != BrowserOPMode.seleniumNew||driver==null) return;
                    e.Handled = true;
                    switch (e.KeyCode)
                    {
                        case Keys.PageUp:
                            chromeSendkeys("^{PGUP}");
                            break;
                        case Keys.PageDown:
                            chromeSendkeys("^{PGDN}");
                            break;
                        default:
                            break;
                    }
                    return;
                }

                if (e.KeyCode == Keys.F)
                {//Alt + f ：切換 Fast Mode 不待網頁回應即進行下一頁的貼入動作（即在不須檢覈貼上之文本正確與否，肯定、八成是無誤的，就可以執行此項以加快輸入文本的動作）當是 fast mode 模式時「送出貼上」按鈕會呈現紅綠燈的綠色表示一路直行通行順暢 20230130癸卯年初九第一上班日週一
                    e.Handled = true;
                    FastModeSwitcher();
                    return;
                }
                if (e.KeyCode == Keys.R)
                {//Alt + r ：在Selenium模式+手動輸入模式下、關閉所在Chrome瀏覽器右側之分頁。（因應《古籍酷》連線不暢所衍生之措施）20231026
                 //有時--尤其在傳回OCR結果時，等待過久，可以多開幾個《古籍酷》的頁面以刺激之。因為取得OCR結果後會切回目前交付OCR的頁面，故將其右方的分頁悉數關閉即可。
                    e.Handled = true;
                    if (browsrOPMode != BrowserOPMode.appActivateByName && keyinTextMode)
                    {
                        try
                        {
                            br.driver = br.driver ?? br.DriverNew();
                            br.driver.SwitchTo().Window(br.driver.CurrentWindowHandle);

                            OpenQA.Selenium.IWebElement iw = br.driver.FindElement(OpenQA.Selenium.By.XPath("/html/body/div[2]"));
                            iw.Click();
                            //Thread.Sleep(800);
                            //Point copyBtnPos = new Point(100, 1050);
                            //Cursor.Position = copyBtnPos;
                            //MouseOperations.MouseEventMousePos(MouseOperations.MouseEventFlags.LeftDown, copyBtnPos);
                            //MouseOperations.MouseEventMousePos(MouseOperations.MouseEventFlags.LeftUp, copyBtnPos);

                            SendKeys.Send("%r");
                            Thread.Sleep(350);
                            //Activate();
                            bringBackMousePosFrmCenter();
                        }
                        catch (Exception)
                        {

                            //throw;
                        }
                    }
                    return;
                }


                if (e.KeyCode == Keys.Left || e.KeyCode == Keys.Right)
                {/*Alt + ←：視窗向左移動30dpi（+ Ctrl：徵調）
                  * Alt + →：視窗向右移動30dpi（+ Ctrl：徵調）*/
                    e.Handled = true;//目前在textBox1時照樣
                    const int w = 30;
                    //int w = this.Width / 2;
                    if (e.KeyCode == Keys.Left) this.Left -= w;
                    if (e.KeyCode == Keys.Right) this.Left += w;
                    mouseMovein();
                    return;
                }

                if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down)
                {/*Alt + ↑：視窗向上移動30dpi（+ Ctrl：徵調；插入點在textBox1時例外）
                  *Alt + ↓：視窗向下移動30dpi（+ Ctrl：徵調；插入點在textBox1時例外）*/
                    e.Handled = true;
                    const int h = 30;//目前在textBox1時照樣
                    if (e.KeyCode == Keys.Up) this.Top -= h;
                    if (e.KeyCode == Keys.Down) this.Top += h;
                    mouseMovein();
                    return;
                }

                if (e.KeyCode == Keys.F5)
                {//Alt + F5 : 在《漢籍全文資料庫》或 CTP 檢索易學關鍵字
                    e.Handled = true;
                    doHanchi_SearchingKeywordsYijing();
                    return;
                }

                if (e.KeyCode == Keys.F6 || e.KeyCode == Keys.F8)
                {//Alt + F6、Alt + F8 : run autoMarkTitles 自動標識標題（篇名）
                    e.Handled = true;
                    autoMarkTitles(); return;
                }

                if (e.KeyCode == Keys.F9)
                {//Alt + F9 : 在《漢籍全文資料庫》或 CTP 檢索易學關鍵字
                    e.Handled = true;
                    doHanchi_SearchingKeywordsYijing();
                    return;
                }

                if (e.KeyCode == Keys.Oemcomma)
                {//Alt + , : 在《漢籍全文資料庫》檢索易學關鍵字
                    if (browsrOPMode != BrowserOPMode.seleniumNew)
                    {
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請先以Selenium模式啟動Chrome瀏覽器（在textBox2中輸入「br」或「bb」就可以了），再繼續。");
                        return;
                    }
                    e.Handled = true;
                    doHanchi_SearchingKeywordsYijing();
                    return;

                }

                if (e.KeyCode == Keys.F10)
                {//Alt + F10 ： 將textBox1中選取的文字送去《古籍酷》自動標點。若無選取則將整個textBox1的內容送去。（小於20字元不處理）20240808（臺灣父親節）
                    if (browsrOPMode != BrowserOPMode.seleniumNew)
                    {
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請先以Selenium模式啟動Chrome瀏覽器（在textBox2中輸入「br」或「bb」就可以了），再繼續。");
                        return;
                    }
                    e.Handled = true;
                    if (textBox1.SelectedText.Length < 5) textBox1.DeselectAll();
                    toGjcoolPunct("https://gj.cool/punct");
                    return;
                }
                if (e.KeyCode == Keys.F11)
                {//Alt + F11 ： 將textBox1中選取的文字送去《古籍酷》自動標點。若無選取則將整個textBox1的內容送去。（小於20字元不處理）20240808（臺灣父親節）
                    if (browsrOPMode != BrowserOPMode.seleniumNew)
                    {
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請先以Selenium模式啟動Chrome瀏覽器（在textBox2中輸入「br」或「bb」就可以了），再繼續。");
                        return;
                    }
                    e.Handled = true;
                    if (textBox1.SelectedText.Length < 5) textBox1.DeselectAll();
                    toGjcoolPunct("https://gj.cool/punct");
                    return;
                }

                if (e.KeyCode == Keys.Clear)
                {// Alt + 5 （數字鍵盤 5）清除標題符碼標記
                    e.Handled = true;
                    clearTitleMarkCode();
                    return;
                }
            }//以上 按下Alt鍵
            #endregion

            #region 多鍵
            //if(e.KeyCode==Keys.Down&&e.KeyCode==Keys.Right|| e.KeyCode == Keys.Down && e.KeyCode == Keys.Left)
            #endregion

            #region 單一鍵
            if (ModifierKeys == Keys.None)
            {//按下單一鍵
                if (e.KeyCode == Keys.F5)
                {
                    e.Handled = true;
                    selLength = textBox1.SelectionLength; selStart = textBox1.SelectionStart;
                    undoRecord();
                    loadText();
                    restoreCaretPosition(textBox1, selStart, selLength);
                    return;
                }
                if (e.KeyCode == Keys.F8)
                {
                    //F8 ：整頁貼上Quick edit [簡單修改模式]  並轉到下一頁
                    e.Handled = true;
                    //還原放大的書圖
                    RestoreImageSize();
                    if (!OcrTextMode)
                    {
                        if (keyinTextMode)
                            SelectAll2Quickedit();
                        else if (autoPaste2QuickEdit)
                        {
                            playSound(soundLike.press, true);
                            altA_predictEndofPageRange();
                            keyDownCtrlAdd(false);
                        }
                    }
                    else
                        pagePaste2GjcoolOCR();//F8 :原為 keysTitleCode();
                    return;
                }
                if (e.KeyCode == Keys.F9)
                {//F9 ：同數字鍵盤「+」 F8 20231213
                    e.Handled = true;
                    //還原放大的書圖
                    RestoreImageSize();
                    //Process.Start(dropBoxPathIncldBackSlash + @"VS\bat\重啟小小輸入法.bat");
                    if (ocrTextMode)
                        pagePaste2GjcoolOCR();
                    else
                    {
                        if (keyinTextMode)
                            SelectAll2Quickedit();
                        else if (autoPaste2QuickEdit)
                        {
                            playSound(soundLike.press, true);
                            altA_predictEndofPageRange();
                            keyDownCtrlAdd(false);
                        }
                    }

                    return;
                }
                if (e.KeyCode == Keys.F12)
                {//按下F12
                    e.Handled = true;
                    //還原放大的書圖
                    RestoreImageSize();
                    if (OcrTextMode)
                        pagePaste2GjcoolOCR();//F12
                    else
                    {
                        if (keyinTextMode)
                            SelectAll2Quickedit();
                        else if (autoPaste2QuickEdit)
                        {
                            playSound(soundLike.press, true);
                            altA_predictEndofPageRange();
                            keyDownCtrlAdd(false);
                        }
                    }
                    return;
                }
                if (e.KeyCode == Keys.Escape)
                {//按下 Esc鍵
                    e.Handled = true;
                    if (!textBox4.Focused && !textBox2.Focused)
                    {
                        if (!textBox4.Text.IsNullOrEmpty() && int.TryParse(textBox4.Text, out int i))
                        {//如果在檢索《易》學關鍵字 20241212
                            this.WindowState = FormWindowState.Minimized;
                        }
                        else
                        {
                            if (MessageBoxShowOKCancelExclamationDefaultDesktopOnly("將表單隱藏到系統任務列中？") == DialogResult.OK)
                                hideToNICo();
                            else
                                AvailableInUseBothKeysMouse();
                        }
                    }
                    return;
                    //if (textBox1.Text == "")
                    ////預設為最上層顯示，若textBox1值為空，則按下Esc鍵會隱藏到任務列中；點一下即恢復
                    //{
                    //    hideToNICo();
                    //}
                }

                if (e.KeyCode == Keys.Clear)
                {//與 textBox1 按下 「Alt + Pause」同
                 //當表單在Num Lock關閉時按下數字鍵盤的「5」 ：在表單按下數字鍵盤的「5」 ： 自動判斷標題行，加上篇名格式代碼並前置N個全形空格.N，預設為2.且可在執行此項時，選取空格數以重設篇名前要空的格數
                 //> 此法可與 Alt + t detectTitleYetWithoutPreSpace() 參互應用
                    e.Handled = true;
                    undoRecord(); stopUndoRec = true; PauseEvents();
                    autoKeysTitleCodeAndPreWideSpace();
                    ResumeEvents(); stopUndoRec = false;
                    undoRecord();
                    //if (!textBox1.Text.IsNullOrEmpty())
                    //    try
                    //    {
                    //        Clipboard.SetText(textBox1.Text);
                    //    }
                    //    catch (Exception)
                    //    {
                    //        playSound(soundLike.error, true);
                    //    }
                    return;

                }


            }//以上 按下單一鍵
            #endregion
        }

        internal void FastModeSwitcher()
        {
            fastMode = !fastMode;
            if (!autoPaste2QuickEdit && textBox1.TextLength > 1000) turnOn_autoPastetoQuickEdit();//在內容量大時自動切到自動連續輸入模式
            if (fastMode)
            {
                playSound(soundLike.over, true);//播放音效，啟動快捷模式之通知 20250206 蛇年初九
                                                //YouChat菩薩：在C#中，紅綠燈的綠色值為Color.FromArgb(0, 255, 0)。它可以指定RGB顏色，其中紅色的值為0，綠色的值為255，藍色的值為0
                notFastModeColor = button1.ForeColor;
                button1.ForeColor = Color.FromArgb(0, 255, 0);//https://www.google.com/search?q=Color.FromArgb(0%2C+255%2C+0)%3B&oq=Color.FromArgb(0%2C+255%2C+0)%3B&gs_lcrp=EgZjaHJvbWUyBggAEEUYOTIICAEQABgIGB4yCggCEAAYgAQYogQyBwgDEAAY7wUyBwgEEAAY7wXSAQczMDhqMGo3qAIAsAIA&sourceid=chrome&ie=UTF-8
            }
            else
            {
                playSound(soundLike.warn, true);//播放音效，關閉快捷模式之通知 20250207 蛇年初十
                if (notFastModeColor != null) button1.ForeColor = notFastModeColor;
            }
        }

        /// <summary>
        /// YAKC - Key-Mouse Click Visualizer 
        /// OBS運行時才處理
        /// </summary>
        private void YAKCSwitchr()
        {/* 20240816 
          * https://obsproject.com/forum/resources/yakc-key-mouse-click-visualizer.1828/
          * https://github.com/iammodev/YAKC?tab=readme-ov-file#usage
                  */
            Process[] p = Process.GetProcessesByName("obs64");
            if (p.Count() == 0) return;
            //string fn = "YAKC", f = "X:\\YAKC-win32-x64\\" + fn+ ".exe";
            string fn = "YAKC", f = Path.Combine("X:\\YAKC-win32-x64\\", fn) + ".exe";
            if (!File.Exists(f)) return;
            p = Process.GetProcessesByName(fn);
            if (p.Count() == 0)
            {
                Process.Start(f);
            }
            else
            {
                //終止程序,釋放系統記憶體                        
                br.killProcesses(new string[] { fn });
            }
        }
        /// <summary>
        /// 切換語音辨識功能
        /// Ctrl + F2 切換語音操作（預設為非 Windows 語音辨識操作）識別用
        /// </summary>
        private void speechRecognitionOPSwitchr()
        {/*20230122癸卯年初一 chatGPT大菩薩：
                  * 您的程式碼看起來很好，可以正常啟動 Windows 內建的語音辨識軟體。您使用了 Process.Start() 方法啟動 sapisvr.exe 程序，並使用了 -SpeechUX 參數來啟動語音辨識功能。
                    在終止語音辨識軟體程序時，您使用了 killProcesses() 方法來釋放系統記憶體。這是正確的做法，因為當不再使用語音辨識時，終止程序可以釋放系統資源。
                    還有一點要注意的是, Windows內建的語音辨識軟體路徑可能會因為不同版本或系統有所不同,因此您可能需要檢查並確保路徑是正確的。
                  */
            speechRecognitionOPmode = !speechRecognitionOPmode;
            string f = Environment.GetFolderPath(Environment.SpecialFolder.Windows) + "\\Speech\\Common\\sapisvr.exe";
            if (!File.Exists(f)) return;
            if (speechRecognitionOPmode)
            {
                new SoundPlayer(@"C:\Windows\Media\Speech On.wav").Play();
                //%windir%\\Speech\\Common\\sapisvr.exe -SpeechUX
                Process.Start(f, "-SpeechUX");
            }
            else
            {
                new SoundPlayer(@"C:\Windows\Media\Speech Off.wav").Play();
                //終止 語音辨識軟體 程序,釋放系統記憶體                        
                br.killProcesses(new string[] { "sapisvr" });
            }
        }

        /// <summary>
        /// 送去《古籍酷》自動標點
        /// </summary>
        /// <param name="url">要送去的網站網址</param>
        /// <param name="copyResult">是否要複製自動標點的結果</param>
        /// <param name="omitExam">是否要略過除了短於20字的檢查、防呆</param>
        /// <returns>若失敗則回傳false</returns>
        private bool toGjcoolPunct(string url, bool copyResult = false, bool omitExam = false)
        {
            #region 防呆
            if (!omitExam)
            {
                try
                {
                    if (!IsValidUrl＿keyDownCtrlAdd(br.GetDriverUrl) || !IsValidUrl＿keyDownCtrlAdd(textBox3.Text))
                    {
                        bool found = false;
                        if (!IsValidUrl＿keyDownCtrlAdd(br.GetDriverUrl) && IsValidUrl＿keyDownCtrlAdd(textBox3.Text))
                        {
                            for (int i = br.driver.WindowHandles.Count - 1; i > -1; i--)
                            {
                                br.driver.SwitchTo().Window(br.driver.WindowHandles[i]);
                                if (br.driver.Url == textBox3.Text)
                                { found = true; break; }
                            }
                            if (!found)
                            {
                                string preUrl = textBox3.Text.Substring(0, textBox3.Text.IndexOf("&editwiki="));
                                for (int i = br.driver.WindowHandles.Count - 1; i > -1; i--)
                                {
                                    br.driver.SwitchTo().Window(br.driver.WindowHandles[i]);
                                    if (br.driver.Url.StartsWith(preUrl))
                                    {
                                        found = true;
                                        if (MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否是這個頁面？" + Environment.NewLine + Environment.NewLine + "textBox3.Text=" + textBox3.Text) == DialogResult.OK)
                                        {
                                            if (IsValidUrl＿keyDownCtrlAdd(textBox3Text))
                                                br.driver.Url = textBox3Text;
                                            else
                                                br.driver.Url = ReplaceUrl_Box2Editor(br.driver.Url);
                                            br.QuickeditLinkIWebElement.Click();
                                        }
                                        break;
                                    }
                                }
                            }
                        }
                        else if (IsValidUrl＿keyDownCtrlAdd(br.GetDriverUrl) || !IsValidUrl＿keyDownCtrlAdd(textBox3.Text))
                            textBox3.Text = br.GetDriverUrl;
                        if (!found)
                            if (DialogResult.Cancel == MessageBoxShowOKCancelExclamationDefaultDesktopOnly("當前頁面似乎沒有自動標點的必要性，確定要送出？", "送交《古籍酷》自動標點", true, MessageBoxDefaultButton.Button2))
                                return false;
                    }
                }
                catch (Exception ex)
                {
                    switch (ex.HResult)
                    {
                        case -2146233088:
                            if (ex.Message.StartsWith("An unknown exception was encountered sending an HTTP request to the remote WebDriver server for URL "))//An unknown exception was encountered sending an HTTP request to the remote WebDriver server for URL http://localhost:6439/session/91263fbb95d208679da86ed250a23ed8/window. The exception message was: 傳送要求時發生錯誤。
                                return false;
                            else
                            {
                                Console.WriteLine(ex.HResult + ex.Message);
                                MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                            }
                            break;
                        default:
                            break;
                    }
                    MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                }
            }
            #endregion//以上防呆

            #region 設定或取得選取文字（要處理的部分）--原理是處理選取的部分，若無選取，則處理整個textBox1內容 20240914
            string x = textBox1.SelectedText == string.Empty ? textBox1.Text : textBox1.SelectedText; bool selAll = false;
            if (x == textBox1.Text)
            {
                //textBox1.Select(0, textBox1.TextLength); 
                //textBox1.Clear();
                textBox1.SelectAll();
                selAll = true;
            }//textBox1.SelectionLength = textBox1.TextLength;//textBox1.SelectAll();//這個方法好像會失靈（應該不會，是自己在設定 SelectionSart SelectionLength 時有問題，或VisualStudio當掉了，重啟即可。）
            else
            {
                overtypeModeSelectedTextSetting(ref textBox1);
                //最後不要選到分段符號及Xml標記
                if (textBox1.SelectionLength > 2)
                {
                    while (("<" + Environment.NewLine).Contains(textBox1.SelectedText.Substring(textBox1.SelectionLength - 1, 1)))
                    {
                        textBox1.SelectionLength--;
                        if (textBox1.SelectionLength - 1 < 0) break;
                    }
                }
                x = textBox1.SelectedText;
            }
            #endregion

            //小於20字元不處理
            if (new StringInfo(CnText.RemovePunctuationsNum(x)).LengthInTextElements < 20)
                if (DialogResult.Cancel == MessageBoxShowOKCancelExclamationDefaultDesktopOnly("字數太少！碇定要送去《古籍酷》自動標點？", "《古籍酷》自動標點", true, MessageBoxDefaultButton.Button2))
                    return false;
            TopMost = false; int s = textBox1.SelectionStart;//,l=textBox1.SelectionLength;
                                                             //舊版會破壞"<p>"記號，故先予清除，之後可用軟件中標識<p>的功能補諸20240809(或有空時再學昨天恢復分段符號的方法 RestoreParagraphs ，只是這次不是分段符號，而是分段記號（<p>），或將之擴展為傳入指定字符作為引數）。
            bool reMarkFlag = false;
            saveText();
            if (url == "https://old.gj.cool/gjcool/index")
            {
                //記下最後是否是<p>，是的話最後補上
                if (x.Length > 3 && x.Substring(x.Length - "<p>".Length) == "<p>") reMarkFlag = true;
                x = x.Replace("<p>", string.Empty);
            }
            CnText.FormalizeText(ref x);
            //CnText.RemoveBooksPunctuation(ref x);//有些是手動添加的書名號或篇名號，不宜逕削去 20240813 然《古籍酷》自動標點仍會先清除書名號，但篇名號不管。20240911 今仍清除篇名號，以免橫生枝節
            x = x.Replace("《", "［").Replace("》", "］");//書名號亦會被自動標點清除故,以備還原 20241001
            x = x.Replace("「", "〔").Replace("」", "〕");//引號亦會被自動標點清除故,以備還原 20241001

            const string symbolToReplaceWithFullWidthSpace = "◎";
            if (!omitExam)
            {
                CnText.ReplaceFullWidthSpace(ref x, symbolToReplaceWithFullWidthSpace);
            }
            else
                x = x.Replace(Environment.NewLine, Environment.NewLine + Environment.NewLine + Environment.NewLine + Environment.NewLine); //WordVBA處理《漢籍全文資料庫》送交《古籍酷》標點單個分段符號會被清除故 20241001

            //清除縮排即凸排格式標記，即將分段符號前後的空格「　」均予清除
            //x = Regex.Replace(x, $@"\s*{Environment.NewLine}+\s*", Environment.NewLine);//發現問題出在使用了 .Text屬性值作判斷故自動標點之方法過早結束迴圈，故今先還原，再觀察 20240918

            string originalText = x;// 
                                    //x = x.Replace(Environment.NewLine, string.Empty).Replace("·", string.Empty);//OCR回來後我這裡自動標點如「嗚呼」仍會標上驚嘆號，故交由 FormalizeText 來處理
            x = x.Replace(Environment.NewLine, string.Empty);//.Replace("·", string.Empty);音節號已於 RemoveBooksPunctuation 中清除
                                                             //x = x.Replace(Environment.NewLine, string.Empty).Replace("·", string.Empty).Replace("！",string.Empty);
            switch (url)
            {
                case "https://old.gj.cool/gjcool/index":
                    if (!br.GjcoolPunctOld(ref x)) return false;//舊版不會去除分段符號，但會在每段前誤加句號，故還是先清除分段符號再送去
                    break;
                case "https://gj.cool/punct":
                    if (!br.GjcoolPunct(ref x)) return false;//新版會去除分段符號（但感覺有時會干擾，還不如先幫它清除分段符號試試。20240813
                    break;
                default:
                    break;
            }

            //textBox1.SelectedText = x;//先作個備份（還原）記錄，以防萬一 20240914 作為下面 RestoreParagraphs 方法除錯用，因其中已有 Debugger.Break(); 故今省略
            //textBox1.Select(s, l);

            if (originalText != x //如果傳回的值與原來已有不同（即當標點過了）
                && !originalText.Contains(x)
                && !originalText.Replace(Environment.NewLine, string.Empty).Contains(x))
            //恢復段落符號
            //if (!omitExam) x = CnText.RestoreParagraphs(ref originalText, ref x);
            {
                x = CnText.RestoreParagraphs(originalText, ref x);//有時《漢籍全文資料庫》亦會不止有一段文字而已，故須保留/還原分段符號
                if (!omitExam)
                {
                    x = x.Replace(symbolToReplaceWithFullWidthSpace, "　");//將 ReplaceFullWidthSpace()所置換的全形空格還原
                }
                x = x.Replace("［", "《").Replace("］", "》");//書名號亦會被自動標點清除故
                x = x.Replace("〔", "「").Replace("〕", "」");//引號亦會被自動標點清除故,以還原 20241001
                try
                {
                    br.driver.SwitchTo().Window(br.LastValidWindow);
                    if (br.driver.Url != textBox3.Text)
                    {
                        foreach (var item in br.driver.WindowHandles)
                        {
                            br.driver.SwitchTo().Window(item);
                            if (br.driver.Url == textBox3.Text) break;
                        }
                    }
                    if (br.driver.Url != textBox3.Text)
                    {
                        //Debugger.Break();
                        if (IsValidUrl＿keyDownCtrlAdd(br.driver.Url))
                        {
                            FixUrl＿ImageTextComparisonPage(br.driver.Url, true, true);
                            //int boxPosition = br.driver.Url.IndexOf("#box(");
                            //if (boxPosition > -1) br.driver.Navigate().GoToUrl(br.driver.Url.Substring(0, boxPosition) + "#editor");
                            if (br.driver.Url != textBox3.Text)
                                textBox3.Text = br.driver.Url;
                        }
                    }

                }
                catch (Exception ex)
                {
                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                }
                //OCR結果文本規範化
                CnText.FormalizeText(ref x);
                bool p = pasteAllOverWrite;
                pasteAllOverWrite = true;//防止隱藏到系統任務列去
                if (reMarkFlag) x += "<p>";
                undoRecord(); stopUndoRec = true; PauseEvents();
                if (selAll && textBox1.SelectedText != textBox1.Text)
                    textBox1.SelectAll();
                textBox1.SelectedText = CnText.BooksPunctuation(ref x, true);

                pasteAllOverWrite = p;
                //textBox1.SelectedText = x;
                AvailableInUseBothKeysMouse();
                if (!selAll) textBox1.Select(s, x.Length);
                else textBox1.Select(0, 0);
                textBox1.ScrollToCaret();
                stopUndoRec = false; ResumeEvents();
                undoRecord();
                if (copyResult)
                {
                    if (textBox1.SelectionLength == 0)
                        try
                        {
                            Clipboard.SetText(textBox1.Text);
                        }
                        catch (Exception)
                        {
                        }
                    else
                        textBox1.Copy();
                }
                return true;
            }
            else
            {
                MessageBoxShowOKExclamationDefaultDesktopOnly("沒有標點，請檢查後重試，或送去舊站看看。感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主");
                return false;
            }
        }

        private void doHanchi_SearchingKeywordsYijing()
        {
            //if(keyinTextMode)keyinTextMode = false;
            //if(autoPasteFromSBCKwhether)autoPasteFromSBCKwhether    = false;
            //if (autoPastetoQuickEdit) autoPastetoQuickEdit = false;
            if (keyinTextMode) KeyinTextmodeSwitcher(false);
            playSound(soundLike.press, true);
            WindowHandles.TryGetValue("Hanchi_CTP_SearchingKeywordsYijing", out string windowHandle_Hanchi_CTP_SearchingKeywordsYijing);
            if (br.driver != null && br.IsDriverInvalid())
            {
                try
                {
                    if (windowHandle_Hanchi_CTP_SearchingKeywordsYijing != string.Empty)
                    {
                        if (driver.WindowHandles.Contains(windowHandle_Hanchi_CTP_SearchingKeywordsYijing))
                            br.driver.SwitchTo().Window(windowHandle_Hanchi_CTP_SearchingKeywordsYijing);
                        else
                        {
                            WindowHandles.Remove("Hanchi_CTP_SearchingKeywordsYijing");
                            windowHandle_Hanchi_CTP_SearchingKeywordsYijing = string.Empty;
                        }
                    }
                    else
                        br.driver.SwitchTo().Window(br.LastValidWindow);

                }
                catch (Exception ex)
                {
                    switch (ex.HResult)
                    {
                        case -2146233088://invalid argument: 'handle' must be a string
                                         //(Session info: chrome = 130.0.6723.70)
                            if (ex.Message.StartsWith("invalid argument: 'handle' must be a string"))
                                RestartChromedriver();
                            else if (ex.Message.StartsWith("An unknown exception was encountered sending an HTTP request to the remote WebDriver server for URL"))//An unknown exception was encountered sending an HTTP request to the remote WebDriver server for URL http://localhost:1116/session/faf3a898f90c7f255a3a0d3264405372/window. The exception message was: 傳送要求時發生錯誤。
                                RestartChromedriver();
                            else
                                goto default;
                            break;
                        default:
                            Console.WriteLine(ex.HResult + ex.Message);
                            //MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                            Debugger.Break();
                            ////chromedriver被誤關時 20241008
                            //br.driver = null;
                            //br.DriverNew();
                            RestartChromedriver();
                            break;
                    }
                }
            }
            #region 關閉《漢籍全文資料庫》開啟的頁面20240926
            if (windowHandle_Hanchi_CTP_SearchingKeywordsYijing != string.Empty)
            {
                for (int i = br.driver.WindowHandles.Count - 1; i > -1; i--)
                {
                    if (driver.CurrentWindowHandle == windowHandle_Hanchi_CTP_SearchingKeywordsYijing) break;
                    br.driver.SwitchTo().Window(driver.WindowHandles[i]);
                    //如果「回查詢結果」元件存在的話//文本閱讀內的檢索（《漢籍全文資料庫》），如果有開啟「回查詢結果」的頁面，則關閉20240926
                    if (br.WaitFindWebElementBySelector_ToBeClickable("body > form > table > tbody > tr:nth-child(2) > td > table > tbody > tr:nth-child(1) > td > table > tbody > tr > td.btn62 > a")?.GetAttribute("text") == "回查詢結果")
                    {//如果不能返回上一頁，即開啟新分頁者，即予關閉。
                        if (!br.CanGoBack())
                        {
                            br.driver.Close();
                            br.driver.SwitchTo().Window(br.driver.WindowHandles.Last());
                        }
                        //else
                        //    br.driver.Navigate().Back();//CanGoBack()裡頭已有！ 20240926
                        if (br.driver.Title.Contains("漢籍全文文本閱讀"))
                        {
                            while (br.WaitFindWebElementBySelector_ToBeClickable("body > form > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td > table > tbody > tr:nth-child(1) > td.btn62 > a")?.GetAttribute("text") == "回瀏覽")
                            { br.driver.Navigate().Back(); }
                            break;
                        }
                        //break;//如果有開啟多個
                    }
                    //《漢籍全文資料庫》【檢索報表】標籤控制項(關閉開啟的分頁）
                    else if (WaitFindWebElementBySelector_ToBeClickable("body > form > table > tbody > tr:nth-child(2) > td:nth-child(1) > font > b > nobr")?.GetAttribute("textContent") == "【檢索報表】")
                    //while (null != waitFindWebElementBySelector_ToBeClickable("body > form > table > tbody > tr:nth-child(2) > td:nth-child(1) > font > b > nobr"))
                    {
                        driver.Close(); driver.SwitchTo().Window(driver.WindowHandles.Last());
                        if (br.driver.Title.Contains("漢籍全文資料庫"))
                        {
                            //搜尋按鈕
                            if (br.WaitFindWebElementBySelector_ToBeClickable("#frmTitle > table > tbody > tr:nth-child(2) > td > table > tbody > tr:nth-child(8) > td > input[type=IMAGE]:nth-child(2)")?.GetAttribute("title") == "搜尋")
                                break;
                        }
                    }
                }
            }
            #endregion// 關閉《漢籍全文資料庫》開啟的頁面20240926

            while (true)
            {
                TopMost = false;
                if (br.Hanchi_CTP_SearchingKeywordsYijing()) break;
            }
        }

        /// <summary>
        /// 手動輸入模式切換用 20240719
        /// </summary>
        internal void KeyinTextmodeSwitcher(bool soundplay = true)
        {
            //避免事件被終止
            if (!_eventsEnabled) _eventsEnabled = true;

            //避免還原記錄被終止
            if (stopUndoRec) stopUndoRec = false;


            //重設欄位變量，以免OCR快速鍵失效
            PagePaste2GjcoolOCR_ing = false;
            if (keyinTextMode)
            {
                if (soundplay)
                    new SoundPlayer(@"C:\Windows\Media\Speech Off.wav").Play();
                keyinTextMode = false; return;
            }
            else
                if (fastMode) FastModeSwitcher();
            if (soundplay)
                new SoundPlayer(@"C:\Windows\Media\Speech On.wav").Play();
            //設定成手動，自動及全部覆蓋之貼上則設成false
            keyinTextMode = true; pasteAllOverWrite = false; autoPaste2QuickEdit = false;
            button1.Text = "分行分段";
            button1.ForeColor = new System.Drawing.Color();//預設色彩 預設顏色 https://stackoverflow.com/questions/10441000/how-to-programmatically-set-the-forecolor-of-a-label-to-its-default
        }

        /// <summary>
        /// 開啟/新增一個Form1（主表單）
        /// </summary>
        /// <returns></returns>
        private static Form1 newForm1()
        {
            Form1 formNew = new Form1();
            formNew.Show();
            formNew.Name = "Form" + Application.OpenForms.Count;
            formNew.Text = formNew.Name;
            //formNew.textBox1.Text = formNew.Name;

            // 調整背景顏色
            lightenTextBoxBackground(formNew);

            return formNew;

            void lightenTextBoxBackground(Form1 newForm)
            {// 20250308 Copilot大菩薩
                // 確保有其他表單已經開啟
                if (Application.OpenForms.Count > 1)
                {
                    // 使用目前已開啟最後一個表單的背景顏色
                    Form1 lastOpenedForm = (Form1)Application.OpenForms[Application.OpenForms.Count - 2];
                    Color currentColor = lastOpenedForm.textBox1.BackColor;

                    // 計算變淡後的顏色
                    int lightenAmount = 10; // 調整這個值來控制變淡程度
                    int red = Math.Min(currentColor.R + lightenAmount, 255);
                    int green = Math.Min(currentColor.G + lightenAmount, 255);
                    int blue = Math.Min(currentColor.B + lightenAmount, 255);

                    // 設定新表單 textBox1 的背景顏色
                    newForm.textBox1.BackColor = Color.FromArgb(currentColor.A, red, green, blue);
                }
                else
                {
                    // 如果是第一個表單，使用默認背景顏色或指定顏色
                    //newForm.textBox1.BackColor = Color.LightGray;
                }
            }
        }

        //private static Form1 newForm1()
        //{
        //    Form1 formNew = new Form1();
        //    formNew.Show();
        //    formNew.Name = "Form" + Application.OpenForms.Count;
        //    formNew.Text = formNew.Name;
        //    formNew.textBox1.Text = formNew.Name;
        //    lightenTextBoxBackground(formNew);
        //    return formNew;

        //    //Copilot大菩薩：20250308 https://copilot.microsoft.com/shares/SZjL4PEMMyz3KUNszbUeB
        //    void lightenTextBoxBackground(Form1 newForm)
        //    {
        //        // 取得目前的背景顏色
        //        Color currentColor = ((Form1)Application.OpenForms[Application.OpenForms.Count-1]).textBox1.BackColor;

        //        // 計算變淡後的顏色
        //        int lightenAmount = 20; // 調整這個值來控制變淡程度
        //        int red = Math.Min(currentColor.R + lightenAmount, 255);
        //        int green = Math.Min(currentColor.G + lightenAmount, 255);
        //        int blue = Math.Min(currentColor.B + lightenAmount, 255);

        //        // 設定 textBox1 的背景顏色為變淡後的顏色
        //        newForm.textBox1.BackColor = Color.FromArgb(currentColor.A, red, green, blue);
        //    }
        //}


        /// <summary>
        /// 執行OCR主程式
        /// </summary>
        /// <param name="ocrSiteTitle">指OCR網站（Google Keep或《古籍酷》）</param>
        /// <returns>成功執行傳回true</returns>
        private bool toOCR(br.OCRSiteTitle ocrSiteTitle, bool justDownloadImage = false)
        {
            //Form1.playSound(Form1.soundLike.press);

            TopMost = false;

            //br.ActiveForm1 = this;

            try
            {

                try
                {
                    if (!br.driver.WindowHandles.Contains(br.driver.CurrentWindowHandle))
                        br.driver.SwitchTo().Window(br.LastValidWindow);
                }
                catch (Exception)
                {
                    try
                    {
                        br.driver.SwitchTo().Window(br.LastValidWindow);
                        playSound(soundLike.exam);
                    }
                    catch (Exception)
                    {
                        if (br.driver.WindowHandles.Count > 0)
                        {
                            br.driver.SwitchTo().Window(br.driver.WindowHandles[0]);
                            br.LastValidWindow = br.driver.WindowHandles[0];
                        }

                    }
                }
                br.LastValidWindow = br.driver.CurrentWindowHandle;
            }
            catch (Exception)
            {
                return false;
            }

            if (!justDownloadImage)
            {
                #region 檢查是否必要 20230804Bard大菩薩：https://g.co/bard/share/9130d688e253            
                string quickedit_data_textboxTxt = br.Quickedit_data_textbox?.Text;//br.Quickedit_data_textboxTxt;
                                                                                   //bool chk = false;
                                                                                   //if (quickedit_data_textboxTxt.Length > 1000)
                                                                                   //{
                                                                                   //    Regex regex = new Regex(@"\，|\。");
                                                                                   //    Match match = regex.Match(quickedit_data_textboxTxt);
                                                                                   //    chk = match.Success;
                                                                                   //}
                                                                                   //else
                                                                                   //{
                                                                                   //    chk = quickedit_data_textboxTxt.Contains("，") || quickedit_data_textboxTxt.Contains("。");
                                                                                   //}
                if (quickedit_data_textboxTxt == null) return false;
                if ((CnText.HasEditedWithPunctuationMarks(ref quickedit_data_textboxTxt) && (textBox1.Text != "●" || quickedit_data_textboxTxt != "●")) ||
                    quickedit_data_textboxTxt.Contains("picture") || quickedit_data_textboxTxt.Contains("entity"))
                {
                    OCRBreakSoundNotification();
                    if (DialogResult.Cancel == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("目前頁面似乎已經整理過了，確定還要繼續嗎？" +
                          Environment.NewLine + Environment.NewLine + "================" + Environment.NewLine +
                        quickedit_data_textboxTxt))
                    {
                        undoRecord();
                        textBox1.Text = br.CopyQuickedit_data_textboxText();//quickedit_data_textboxTxt;
                        if (!Active) AvailableInUseBothKeysMouse();
                        //br.WindowsScrolltoTop();
                        return false;
                    }
                }
                //else if ((br.Quickedit_data_textbox == null ? 0 : (new StringInfo(br.Quickedit_data_textbox?.Text)?.LengthInTextElements)) < (normalLineParaLength == 0 ? 20 : normalLineParaLength)
                //    && quickedit_data_textboxTxt != "\t")// 「	」"\t"是新建的維基文本故 20240405
                else if ((quickedit_data_textboxTxt == null ? 0 : (new StringInfo(quickedit_data_textboxTxt)?.LengthInTextElements)) < (NormalLineParaLength == 0 ? 20 : NormalLineParaLength)
                    && quickedit_data_textboxTxt != "\t" && quickedit_data_textboxTxt != " "
                    && quickedit_data_textboxTxt != "●")// 「	」"\t"是新建的維基文本故 20240405 "\t"會被Text屬性轉換成 " "
                {
                    OCRBreakSoundNotification();
                    if (DialogResult.Cancel == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("目前頁面內容似乎太短了，確定還要交給OCR嗎？" +
                            Environment.NewLine + Environment.NewLine + "================" + Environment.NewLine +
                            quickedit_data_textboxTxt))
                    {
                        undoRecord();
                        textBox1.Text = br.CopyQuickedit_data_textboxText();//quickedit_data_textboxTxt;
                        if (!Active) AvailableInUseBothKeysMouse();
                        //br.WindowsScrolltoTop();
                        return false;
                    }
                }
                #endregion
            }

            #region 記下送去OCR前的分頁
            string currentWindowHndl = string.Empty;
            try
            {
                currentWindowHndl = br.driver.CurrentWindowHandle;
            }
            catch (Exception)
            {
                currentWindowHndl = br.driver.WindowHandles.Last();
            }
            br.LastValidWindow = currentWindowHndl;
            #endregion

            //下載書圖
            string imgUrl = Clipboard.GetText(), downloadImgFullName; bool ocrResult = false;
            if (imgUrl.Length > 4
            && imgUrl.Substring(0, 4) == "http"
            && imgUrl.Substring(imgUrl.Length - 4, 4) == ".png")
                ocrResult = DownloadImage(imgUrl, out downloadImgFullName);
            else
            {
                imgUrl = br.GetImageUrl();
                if (imgUrl == "")
                {
                    //br.WindowsScrolltoTop();
                    return false;
                }
                ocrResult = DownloadImage(imgUrl, out downloadImgFullName);
                if (downloadImgFullName == "")
                {
                    //br.WindowsScrolltoTop();
                    return false;
                }
            }

            if (!ocrResult)
            {
                //br.WindowsScrolltoTop();
                return false;
            }

            if (justDownloadImage) return true;

            ocrResult = false; TopMost = false;// Visible = false;//WindowState = FormWindowState.Minimized;

            #region toOCR
            int windowsCount = driver.WindowHandles.Count;//作為判斷在執行OCR程序時有沒有新開的分頁視窗
            br.StopOCR = false;
            ////string currentWindowHndl = br.driver.CurrentWindowHandle;
            //br.LastValidWindow = currentWindowHndl;//br.driver.CurrentWindowHandle;//改到前面
            switch (ocrSiteTitle)
            {
                //Google Keep
                case br.OCRSiteTitle.GoogleKeep:
                    ocrResult = br.OCR_GoogleKeep(downloadImgFullName);
                    break;
                //《古籍酷》
                case br.OCRSiteTitle.GJcool:
                    //br.ActiveForm1 = this;
                    //br.ActiveForm1.TopMost = false;
                    //br.ActiveForm1 = this;
                    TopMost = false;
                    //try
                    //{
                    br.driver.SwitchTo().Window(currentWindowHndl);
                    if (BatchProcessingGJcoolOCR)
                        //ocrResult = br.OCR_GJcool_BatchProcessing(downloadImgFullName);
                        ocrResult = br.OCR_GJcool_BatchProcessing_new(downloadImgFullName);
                    else
                        ocrResult = br.OCR_GJcool_AutoRecognizeVertical(downloadImgFullName);
                    //}
                    //catch (Exception ex)
                    //{
                    //    switch (ex.HResult)
                    //    {
                    //        case -2146233088://"no such window\n  (Session info: chrome=112.0.5615.138)"
                    //            if (ex.Message.IndexOf("no such window") > -1)
                    //                break;//手動關閉或誤關視窗時忽略不計。
                    //            break;
                    //        default:
                    //            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                    //            break;
                    //    }
                    //}
                    break;
                //《看典古籍》網頁版
                case br.OCRSiteTitle.KanDianGuJi:
                    try
                    {
                        ocrResult = br.OCR_KanDianGuJi(downloadImgFullName);
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                    break;
                //《看典古籍》OCR API
                case br.OCRSiteTitle.KanDianGuJiAPI:
                    //20240730 Copilot大菩薩：如果您不希望在呼叫端使用 await，您可以使用 Task.Result 或 Task.GetAwaiter().GetResult() 來獲取 Task 的結果。這兩種方法都會阻塞當前線程，直到 Task 完成。以下是一個範例：
                    //bool ocrResult = PerformOCR().GetAwaiter().GetResult();
                    //或者
                    //因為沒有介面，只好用這樣來視覺化操作效果：
                    if (!ocrTextMode)
                    {
                        PauseEvents();
                        textBox1.Clear();
                        ResumeEvents();
                    }
                    ocrResult = PerformOCR();
                    if (!ocrTextMode)
                        Form1.playSound(Form1.soundLike.info, true);
                    //請注意，這種方法會阻塞當前線程，直到 PerformOCR 方法完成。如果 PerformOCR 方法需要花費很長時間，這可能會導致您的應用程式暫時無響應。因此，雖然這種方法可以避免將呼叫端方法變為異步，但它可能會降低您的應用程式的響應性。
                    break;
                default:
                    break;
            }
            try
            {
                br.driver?.SwitchTo().Window(currentWindowHndl);
            }
            catch (Exception)
            {
                try
                {
                    if (br.driver.WindowHandles.Contains(LastValidWindow))
                        br.driver?.SwitchTo().Window(LastValidWindow);
                    else
                        br.driver?.SwitchTo().Window(br.driver.WindowHandles.Last());
                }
                catch (Exception)
                {
                    try
                    {
                        if (!br.IsDriverInvalid())
                        {
                            br.driver.SwitchTo().Window(br.driver.WindowHandles.Last());
                            StopOCR = true; return false;
                        }
                        else
                            StopOCR = true; return false;
                    }
                    catch (Exception)
                    {
                        StopOCR = true; return false;
                    }
                }
                StopOCR = true; return false;
            }

            if (!ocrResult)
            {
                MessageBox.Show("請重來一次；重新執行一次。感恩感恩　南無阿彌陀佛", "發生錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                //try
                //{
                //    bool eventenable = _eventsEnabled;
                //    if (EventsEnabled) PauseEvents();
                //    //br.driver?.SwitchTo().Window(currentWindowHndl);
                //    if (Clipboard.GetText() != string.Empty)
                //    {
                //        playSound(soundLike.waiting);
                //        AvailableInUseBothKeysMouse();
                //        SendKeys.SendWait("%{ins}");//Alt + Insert
                //        textBox1.Select(0, 0);
                //    }
                //    else
                //    {
                //        SendKeys.Send("%r");//Alt + r 關閉Chrome瀏覽器右邊所有分頁
                //    }

                //    _eventsEnabled = eventenable;
                //}
                //catch (Exception)
                //{
                //    //br.WindowsScrolltoTop();
                //    br.StopOCR = true;
                //    return false;                    
                //}
            }

            //WindowState = FormWindowState.Normal;
            //Visible = true; TopMost = true;

            br.StopOCR = true;

            #region 如果是手動鍵入輸入模式且OCR程序無誤則直接貼上結果並自動標上書名號篇名號，20230309 creedit with chatGPT大菩薩：
            if (ocrResult && keyinTextMode)
            {
                //KeyEventArgs e = new KeyEventArgs(new Keys().);
                //e.KeyCode = (Keys.Alt & Keys.Insert);
                //Form1_KeyDown(sender, e);
                // 建立 Keys.Alt + Keys.Insert 的組合鍵
                //Keys comboKey = Keys.Alt & Keys.Insert;//在 C# 中，要表示兩個按鍵的組合鍵，需要使用 "|" 運算子進行位元運算，而不是 "&" 或 "+" 運算子。 "|" 運算子可以將兩個按鍵的 KeyCode 合併成一個整數，表示按下這兩個按鍵的組合鍵。
                //                                       // 使用 SendKeys 方法觸發按下組合鍵                
                //if (!PasteOcrResultFisrtMode || ocrSiteTitle == br.OCRSiteTitle.KanDianGuJi)
                if (!ocrTextMode || ocrSiteTitle == br.OCRSiteTitle.KanDianGuJi)
                {
                    AvailableInUseBothKeysMouse();//Activate();
                    if (!textBox1.Focused) textBox1.Focus();
                }
                //取得OCR結果
                string x = Clipboard.GetText();

                //清除末綴的分行/段符號
                while (x.Length > 1 && x.Substring(x.Length - Environment.NewLine.Length, Environment.NewLine.Length) == Environment.NewLine)
                    x = x.Substring(0, x.Length - Environment.NewLine.Length);

                //textBox1.Text = x;
                //SendKeys.Send("{" + comboKey + "}");
                //SendKeys.Send("%{insert}");
                ////儲存結果備份，以備還原（若原文即含英數字者）(剪貼簿還保留原樣，姑不用。感恩感恩　南無阿彌陀佛）
                //saveText();
                //清除英數字（OCR辨識誤讀者）                //加上書名號篇名號
                undoRecord();//以便還原

                //OCR自動校正
                if (!autoPaste2QuickEdit) replaceXdirectly(ref x, "OCR", false);

                if (OcrTextMode)//不是OCR直接連續輸入時（即只是單一輸入時）才加上自動加上書名號等標點，畢竟批量輸入的特徵是容易識別的，
                                //如此也可以方便日後後續再手動加工批量OCR讀入的結果時，判斷哪些頁面是已經人工處理/校讀過的，才不會在重新自動標點時，泯滅前人的辛勞，重複做白工 20240116
                                //textBox1.Text = CnText.BooksPunctuation(ref CnText.ClearOthers_ExceptUnicodeCharacters(ref x), false);
                    textBox1.Text = CnText.ClearOthers_ExceptUnicodeCharacters(ref x);
                else
                    //textBox1.Text = CnText.BooksPunctuation(ref CnText.ClearOthers_ExceptUnicodeCharacters(ref x), true);
                    textBox1.Text = CnText.BooksPunctuation(ref CnText.ClearOthers_ExceptUnicodeCharacters(ref x), !OcrTextMode);
                //textBox1.Text = CnText.BooksPunctuation(ref CnText.ClearLettersAndDigits_UseUnicodeCategory(ref x));//清不掉「-」
                //textBox1.Text = CnText.BooksPunctuation(ref CnText.ClearLettersAndDigits(ref x));

                #region OCR成功後則刪除下載的書圖,備份OCR結果; 因為 https://gj.cool/try_ocr 頁面時常傳回假資料（之前曾識別的文本），故今改寫在 textBox3.TextChanged事件中
                //OCR成功後則刪除下載的書圖,備份OCR結果
                //if (File.Exists(downloadImgFullName))
                //{
                //    try
                //    {
                //        File.Delete(downloadImgFullName);
                //    }
                //    catch (Exception ex1)
                //    {
                //        switch (ex1.HResult)
                //        {
                //            case -2147024864:
                //                Task.Run(() =>
                //                {
                //                    Thread.Sleep(600);//"由於另一個處理序正在使用檔案 'X:\\Ctext_Page_Image.txt'，所以無法存取該檔案。"
                //                    File.Delete(downloadImgFullName);
                //                });
                //                break;
                //            default:
                //                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex1.HResult + ex1.Message);
                //                return false;
                //        }
                //    }

                //}
                #endregion

                saveText();
                #region 如果在右邊有新開啟的分頁且網域均為《古籍酷》者等，即予關閉（按下Ctrl鍵略過；有時只是開啟了上一頁欲修訂，就不希望被關掉）
                bool chk = false;
                try
                {
                    chk = ModifierKeys != Keys.Control
                        && br.driver.WindowHandles[br.driver.WindowHandles.Count - 1] != currentWindowHndl
                        && windowsCount < driver.WindowHandles.Count;
                }
                catch (Exception)
                {
                }
                if (chk)
                {
                    //按下Chrome瀏覽器快速鍵關閉右側分頁
                    SendKeys.Send("%r");//這是利用擴充功能設定的快速鍵：https://chrome.google.com/webstore/detail/shortkeys-custom-keyboard/logpjaacgmcbpdkdchjiaagddngobkck
                    Thread.Sleep(250);
                }
                #endregion
            }
            #endregion
            //const string gjcool = "https://gj.cool/try_ocr";

            //Process.Start(gjcool);
            //Process.Start(keep);
            //if (!Active) AvailableInUseBothKeysMouse();//前面已有

            //在連續輸入OCR結果時，提供一次（一頁）操作完成的提示音，以提醒繼續下一頁 20231128
            if (!MuteProcessing && ocrTextMode)// && PasteOcrResultFisrtMode)
                if (ocrResult && PasteOcrResultFisrtMode && File.Exists("C:\\Windows\\Media\\ring07.wav"))
                    using (SoundPlayer sp = new SoundPlayer("C:\\Windows\\Media\\ring07.wav")) { sp.Play(); }

            //if (keyinTextMode)
            //{//如果在手動輸入模式下則自動選取[Quick edit]的內容，方便有時候須用剪下貼上者
            // //其實不用，只要按下1個Tab鍵再全選即可。
            // //br.SelectAllQuickedit_data_textboxContent();

            //    //#region 如果在右邊有新開啟的分頁且網域均為《古籍酷》者等，即予關閉
            //    //for (int i = br.driver.WindowHandles.Count - 1; i > -1; i--)
            //    //{
            //    //    if (br.driver.WindowHandles[i] == currentWindowHndl) break;
            //    //    //這樣的話，還要等網頁開好才能執行
            //    //    br.driver.SwitchTo().Window(br.driver.WindowHandles[i]);
            //    //    string urls = br.driver.Url;
            //    //    if (urls.StartsWith("https://gj.cool/")
            //    //        || urls.StartsWith("https://ocr.gj.cool/")
            //    //        || urls == "https://iplocation.com/"
            //    //        || urls.StartsWith("https://stackoverflow.com/"))
            //    //    {
            //    //        br.driver.Close();
            //    //    }
            //    //}
            //    //#endregion
            //}

            if (!textBox1.Focused) textBox1.Focus();
            //br.WindowsScrolltoTop();
            if (autoTitleMark_OCRTextMode)
            {
                //or lines_perPage 
                if (linesParasPerPage >= countLinesPerPage(textBox1.Text))//行數小於或等於正常行數時才執行，蓋《古籍酷》OCR會將小注分行輸出 20240126
                {
                    undoRecord(); stopUndoRec = true;
                    bool ee = _eventsEnabled;
                    PauseEvents();
                    autoKeysTitleCodeAndPreWideSpace();
                    _eventsEnabled = ee; stopUndoRec = false;
                }
            }
            return ocrResult;// true;
            #endregion


        }

        private void toggleCheck_the_adjacent_pages()
        {
            if (!check_the_adjacent_pages)
            {
                check_the_adjacent_pages = true; new SoundPlayer(@"C:\Windows\Media\Speech On.wav").Play();
                button1.ForeColor = Color.Aquamarine;//如果是鄰近頁連動編輯模式，則顯示為較亮青色Aquamarine，否則為深青色 Color.DarkCyan
                if (MessageBox.Show("是否先檢查文本先前是否曾編輯過？" + Environment.NewLine +
                    "要檢查的話，請先複製其文本，再按確定（ok）按鈕", "", MessageBoxButtons.OKCancel
                    , MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly) == DialogResult.OK)
                {
                    runWordMacro("checkEditingOfPreviousVersion");
                }
            }
            else
            {
                check_the_adjacent_pages = false; new SoundPlayer(@"C:\Windows\Media\Speech Off.wav").Play();
                button1.ForeColor = Color.DarkCyan;//https://learn2android.blogspot.com/2013/04/c.html?lr=1                
            }
            autoPasteFromSBCKwhether = false;
        }

        /// <summary>
        /// 設定是否要以《四部叢刊資料庫》中複製的文本貼上
        /// </summary>
        private void toggleAutoPasteFromSBCKwhether()
        {
            if (!autoPasteFromSBCKwhether)
            {
                new SoundPlayer(@"C:\Windows\Media\Speech On.wav").Play(); autoPasteFromSBCKwhether = true;
            }
            else
            {
                autoPasteFromSBCKwhether = false; new SoundPlayer(@"C:\Windows\Media\Speech Off.wav").Play();
            }
            return;
        }
        /// <summary>
        /// 設定是否是自動連續輸入模式
        /// </summary>
        private void toggleAutoPastetoQuickEdit()
        {
            if (!autoPaste2QuickEdit)
            {
                turnOn_autoPastetoQuickEdit();
            }
            else
            {
                if (fastMode) FastModeSwitcher();
                new SoundPlayer(@"C:\Windows\Media\Speech Off.wav").Play();
                autoPaste2QuickEdit = false;
                //button1.Text = "分行分段";
                button1.ForeColor = Color.FromArgb(255, 0, 0, 0);
                ////如果是鄰近頁連動編輯模式，則顯示為較亮青色 Aquamarine，否則為深青色 Color.DarkCyan。
                //if (check_the_adjacent_pages)
                //{
                //    button1.ForeColor = Color.LightCyan;
                //}
                //else
                //    button1.ForeColor = Color.DarkCyan;//
            }
        }
        /// <summary>
        /// 設定自動連續輸入的實作處理程式
        /// >> 如果是鄰近頁連動編輯模式，則顯示為較淺之青色 LightCyan，否則為深青色 Color.DarkCyan。
        /// </summary>
        private void turnOn_autoPastetoQuickEdit()
        {//set autoPastetoQuickEdit = true//禁遏《四部叢刊資料庫》貼上機制，手動鍵入亦設成false
            autoPaste2QuickEdit = true; keyinTextMode = false; autoPasteFromSBCKwhether = false;
            new SoundPlayer(@"C:\Windows\Media\Speech On.wav").Play();
            button1.Text = "送出貼上";
            //如果是鄰近頁連動編輯模式，則顯示為較亮青色 Aquamarine，否則為深青色 Color.DarkCyan。
            if (check_the_adjacent_pages)
            {
                button1.ForeColor = Color.LightCyan;
            }
            else
                button1.ForeColor = Color.DarkCyan;//https://learn2android.blogspot.com/2013/04/c.html?lr=1                
        }

        [DllImport("user32")]
        static extern bool SetCursorPos(int X, int Y);
        /// <summary>
        /// 讓滑鼠游標光標Cursor拉回到表單範圍
        /// </summary>
        private void mouseMovein()
        {//https://lolikitty.pixnet.net/blog/post/164569578
            SetCursorPos(this.Left + 30, this.Top + 100);
        }

        /// <summary>
        /// 指示是否要隱藏主表單到系統列中：=true則不隱藏
        /// </summary>
        bool dontHide = false;

        /// <summary>
        /// 指示現在主表單是否已隱藏到系統列中
        /// </summary>
        internal bool HiddenIcon { get { return ntfyICo.Visible; } }
        /// <summary>
        /// 隱藏到系統列中
        /// </summary>
        void hideToNICo()
        {
            if (dontHide) return;
            //https://dotblogs.com.tw/jimmyyu/2009/09/21/10733
            //https://dotblogs.com.tw/chou/2009/02/25/7284 https://yl9111524.pixnet.net/blog/post/49024854
            if (this.WindowState != FormWindowState.Minimized)
            {//記下隱藏前的位置與大小
                thisHeight = this.Height; thisWidth = this.Width; thisLeft = this.Left; thisTop = this.Top;
            }
            //this.WindowState = FormWindowState.Minimized;
            this.TopMost = false;
            this.Hide();
            this.ntfyICo.Visible = true;
        }
        /// <summary>
        /// 封裝hideToNICo()方法以供外用
        /// </summary>
        internal void HideToNICo()
        {
            hideToNICo();
        }
        /// <summary>
        /// 備份已貼上之文本的檔名+副檔名（不含路徑）
        /// </summary>
        const string fName_to_Backup_Txt = "cTextBK.txt";
        /// <summary>
        /// 備份已貼上之文本到指定的檔案（以追加方式）
        /// </summary>
        /// <param name="x">要追加備份的內容</param>
        /// <param name="updateLastBackup"></param>
        /// <param name="showColorSignal">是否以顏色指示操作中</param>
        void BackupLastPageText(string x, bool updateLastBackup, bool showColorSignal)
        {
            Color C = this.BackColor; string fBackup = (ocrTextMode & PasteOcrResultFisrtMode) ? "cTextBKOCR.txt" : fName_to_Backup_Txt;
            if (showColorSignal) { this.BackColor = Color.Red; Task.Delay(800).Wait(); }
            //C# 對文字檔案的幾種讀寫方法總結:https://codertw.com/%E7%A8%8B%E5%BC%8F%E8%AA%9E%E8%A8%80/542361/
            string lastPageText = x + Environment.NewLine + "＠"; //"＠" 作為每頁的界號
            if (File.Exists(dropBoxPathIncldBackSlash + fBackup))
            {
                if (updateLastBackup)
                {
                    string bk = File.ReadAllText(dropBoxPathIncldBackSlash + fBackup);
                    int bkLastEnd = bk.LastIndexOf("＠"), bkLastStart = bk.LastIndexOf("＠", bkLastEnd - 1) + 1;
                    //if (bkLastStart == -1) bkLastStart = 0;
                    bk = bk.Substring(0, bkLastStart) + lastPageText;
                    File.WriteAllText(dropBoxPathIncldBackSlash + fBackup, bk, Encoding.UTF8);
                    if (showColorSignal) this.BackColor = C;
                    return;
                }
            }
            File.AppendAllText(dropBoxPathIncldBackSlash + fBackup, lastPageText, Encoding.UTF8);
            if (showColorSignal) this.BackColor = C;
        }

        int waitTimeforappActivateByName = 680;//1100;                                               

        private string quickedit_data_textboxtxt = "";



        /// <summary>
        /// 為了OCR進程，事前檢查空白頁。自動翻頁
        /// Ctrl + Shift + p ： 逐頁瀏覽肉眼檢查空白頁，以免白跑OCR 20240727
        /// 按下 Ctrl 鍵中止（有時要配合對Chrome瀏覽器或對表單按下滑鼠左鍵一下
        /// 20240727 Copilot大菩薩：使用 Selenium 自動瀏覽並檢查空白頁面
        /// </summary>
        /// <param name="url">要開啟瀏覽的頁面網址</param>
        /// <param name="startPageNum">啟始頁碼</param>
        /// <param name="stopPageNum">結束頁碼</param>
        /// <returns></returns>
        internal static bool CheckBlankPagesBeforeOCR_NextPage(string url, int startPageNum, int stopPageNum)
        {
            if (!Form1.IsValidUrl＿ImageTextComparisonPage(url)) return false;
            if (br.driver == null) return false;
            string baseUrl = url.Substring(0, url.IndexOf("&page="));
            ChromeSetFocus();
            driver.SwitchTo().Window(driver.CurrentWindowHandle);
            for (int i = startPageNum + 1; i <= stopPageNum; i++)
            {
                //https://ctext.org/library.pl?if=gb&file=150025&page=59
                string pageUrl = $"{baseUrl}&page={i}";
                br.driver.Navigate().GoToUrl(pageUrl);
                //if (MessageBoxShowOKCancelExclamationDefaultDesktopOnly("繼續下一頁？") == DialogResult.Cancel) break;
                if (ModifierKeys == Keys.Control) break;
                //Thread.Sleep(300);
                //while (br.waitFindWebElementBySelector_ToBeClickable("#canvas > svg") == null) { }
                while (br.Svg_image_PageImageFrame == null) { }
                if (br.Div_generic_TextBoxFrame == null)
                    if (MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否終止/中斷？") == DialogResult.OK)
                        break;
                    else
                    { ChromeSetFocus(); }


            }
            return true;
        }
        /// <summary>
        /// 到下一頁
        /// </summary>
        /// <param name="eKeyCode">按下什麼鍵</param>
        /// <param name="stayInHere">留在本頁而不到下一頁則為true</param>
        /// <param name="notBooksPunctuation">不作書名號等標點時為true</param>
        /// <param name="pagePaste2GjcoolOCR"></param>
        /// <returns>失敗則傳回false</returns>
        private bool nextPages(Keys eKeyCode, bool stayInHere, bool notBooksPunctuation = false, bool pagePaste2GjcoolOCR = false)
        {
            //記下完成的頁面頁碼
            WindowHandles["currentPageNum"] = _currentPageNum;
            string url = textBox3.Text;
            if (url == "") return false;
            if (url.IndexOf("&page=") == -1) return false;
            int edit = url.IndexOf("&editwiki");
            int page = 0; string urlSub = url;
            if (edit > -1)
            {
                urlSub = url.Substring(0, url.IndexOf("&page=") + "&page=".Length);
                page = Int32.Parse(
                    url.Substring(url.IndexOf("&page=") + "&page=".Length,
                    url.IndexOf("&editwiki=") - (url.IndexOf("&page=") + "&page=".Length)));
                if (eKeyCode == Keys.PageDown)
                    url = urlSub + (page + 1).ToString() + url.Substring(url.IndexOf("&editwiki="));
                if (eKeyCode == Keys.PageUp)
                    url = urlSub + (page - 1).ToString() + url.Substring(url.IndexOf("&editwiki="));
                //newTextBox1();
            }
            else
            {
                urlSub = url.Substring(0, url.IndexOf("&page=") + "&page=".Length);
                int ed = url.IndexOf("#");
                if (ed > -1)
                    page = Int32.Parse(url.Substring(url.IndexOf("&page=") + "&page=".Length,
                        url.IndexOf("#") - (url.IndexOf("&page=") + "&page=".Length)));
                else
                    page = Int32.Parse(url.Substring(url.IndexOf("&page=") + "&page=".Length));
                if (eKeyCode == Keys.PageDown)
                    url = urlSub + (page + 1).ToString();
                if (eKeyCode == Keys.PageUp)
                    url = urlSub + (page - 1).ToString();
            }
            bool eventAble = _eventsEnabled;
            if (!_eventsEnabled) _eventsEnabled = true;
            textBox3.Text = url;//此會觸發textchanged事件程序
            _eventsEnabled = eventAble;
            #region 僅瀏覽而不編輯
            switch (browsrOPMode)
            {
                case BrowserOPMode.appActivateByName:
                    Process.Start(url);
                    appActivateByName();
                    break;
                case BrowserOPMode.seleniumNew:
                    if (br.driver == null) br.DriverNew();
                    //br.GoToUrlandActivate(url);
                    //Task wait = Task.Run(() =>//此間操作，因為沒有要操作的元件，所以可以跑線程。20230111
                    //{以別處還要參考，故取消Task
                    //br.GoToUrlandActivate(url, keyinTextMode);
                    try
                    {
                        if (br.driver.Url != url)
                            if (!fastMode)
                                br.GoToUrlandActivate(url, true);
                            else
                                br.driver.Url = url;
                        //br.driver.Navigate().GoToUrl(url);
                    }
                    catch (Exception)
                    {
                        br.GoToUrlandActivate(url, true);
                    }
                    //});
                    //Task.WaitAll();
                    //wait.Wait();

                    //if (!keyinTextMode && autoPaste2QuickEdit) Activate();//●●●●●●●●●●●●●●●20250301省略
                    if (!keyinTextMode && autoPaste2QuickEdit && !fastMode) Activate();//●●●●●●●●●●●●●●●20250301連續輸入時因為要看訊息方法以核對內容，故需要，但快捷模式時，因已略過方塊，故不要

                    break;
                case BrowserOPMode.seleniumGet:
                    //後面的textBox3.Text = url;會觸發private void textBox3_TextChanged 事件程序，於彼處執行瀏覽即可
                    //尚未實作完成
                    break;
                default:
                    return false;
                    //break;
            }
            #endregion

            #region 另一章節文本時 20241226
            try
            {
                if (br.Div_generic_TextBoxFrame != null && br.Div_generic_TextBoxFrame.GetAttribute("textContent") != ""
                    && br.Quickedit_data_textboxTxt == string.Empty)
                {
                    br.driver.Navigate().Refresh();
                    url = br.QuickeditLinkIWebElement.GetAttribute("href");
                    textBox3.Text = url;//此會觸發textchanged事件程序
                    br.QuickeditLinkIWebElement.Click();
                }
            }
            catch (Exception)
            {
            }
            #endregion

            #region 要編輯時
            if (edit > -1)
            {//編輯才執行，瀏覽則省略

                switch (browsrOPMode)
                {
                    //預設瀏覽模式（即開發Selenium架構前）：
                    case BrowserOPMode.appActivateByName:
                        //Task.Delay(500).Wait(); //2200
                        //Task.Delay(1900).Wait(); //2200
                        //Task.Delay(650).Wait(); //目前疾速是650，而穩定是700，乃至1100、1900、2200，看網速
                        Task.Delay(waitTimeforappActivateByName).Wait();
                        //SendKeys.Send("{Tab 24}");
                        SendKeys.Send("{Tab}"); //("{Tab 24}");
                                                //網址尾綴為「#editor」的才能只按一個tab就入文字框中，如 ： https://ctext.org/library.pl?if=gb&file=77367&page=59&editwiki=415472#editor
                        Task.Delay(200).Wait();//200
                        Task.WaitAll();
                        SendKeys.Send("^a");
                        break;
                    //Selenium New Chrome瀏覽器實例模式：
                    case BrowserOPMode.seleniumNew:
                        try
                        {
                            quickedit_data_textboxtxt = br.Quickedit_data_textboxTxt;
                        }
                        catch (Exception ex)
                        {
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                        }
                        //if(!keyinText&& autoPastetoQuickEdit) Activate();
                        break;
                    case BrowserOPMode.seleniumGet:
                        break;
                    default:
                        break;
                }

                //如果是手動鍵入模式：
                if (keyinTextMode)
                {
                    Keys modifierKeys = ModifierKeys;
                    if (modifierKeys == Keys.Shift && !stayInHere)
                        Form1.playSound(Form1.soundLike.press);

                    //OpenQA.Selenium.IWebElement quick_edit_box;
                    switch (browsrOPMode)
                    {
                        case BrowserOPMode.appActivateByName:
                            Task.Delay(290).Wait();
                            Task.WaitAll();
                            SendKeys.Send("^x");//剪下一頁以便輸入備用
                            break;
                        case BrowserOPMode.seleniumNew:
                            if (modifierKeys == Keys.None && !pagePaste2GjcoolOCR)
                            {
                                // 20240913 今因發現可由元件的「value、textContent……」等 Property 取得正確的文本內容，以下就先作廢了

                                //    int retrytimes = 0;
                                //retry:
                                if (br.driver == null) Debugger.Break();
                                //br.driver = br.driver ?? br.DriverNew();

                                //    try
                                //    {//這裡需要參照元件來操作就不宜跑線程了！故此區塊最後的剪貼簿，要求須是單線程者，蓋因剪貼簿須獨占式使用故也20230111                                
                                //     //quick_edit_box = br.waitFindWebElementByName_ToBeClickable("data", br.WebDriverWaitTimeSpan);//br.driver.FindElement(OpenQA.Selenium.By.Name("data"));
                                //     //                                                                                             ////chatGPT：
                                //     //                                                                                             //// 等待網頁元素出現，最多等待 2 秒
                                //     //                                                                                             //OpenQA.Selenium.Support.UI.WebDriverWait wait =
                                //     //                                                                                             //    new OpenQA.Selenium.Support.UI.WebDriverWait
                                //     //                                                                                             //    (br.driver, TimeSpan.FromSeconds(2));
                                //     //                                                                                             ////安裝了 Selenium.WebDriver 套件，才說沒有「ExpectedConditions」，然後照Visual Studio 2022的改正建議又用NuGet 安裝了 Selenium.Suport 套件，也自動「 using OpenQA.Selenium.Support.UI;」了，末學自己還用物件瀏覽器找過了 「OpenQA.Selenium.Support.UI」，可就是沒有「ExpectedConditions」靜態類別可用，即使官方文件也說有 ： https://www.selenium.dev/selenium/docs/api/dotnet/html/T_OpenQA_Selenium_Support_UI_ExpectedConditions.htm 20230109 未知何故 阿彌陀佛
                                //     //                                                                                             //wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(quick_edit_box));

                                //        ////// 在網頁元素載入完畢後才能讀取其.Text屬性值，存入剪貼簿,前置空格會被削去，當是Selenium實作時的問題。
                                //        ////string xq = quick_edit_box.Text;
                                //        ////Clipboard.SetText(xq);
                                //        //if (quick_edit_box != null)
                                //        //{
                                //        //    //用Text屬性（quick_edit_box.Text）取得的值若前有全形空格會被清除
                                //        //    quick_edit_box.SendKeys(OpenQA.Selenium.Keys.LeftControl + "a");
                                //        //    quick_edit_box.SendKeys(OpenQA.Selenium.Keys.LeftControl + "c");
                                //        //}
                                //        //////Task.Delay(-1);
                                //        ////Clipboard.SetText(quick_edit_box.Text);
                                //        br.CopyQuickedit_data_textboxText();
                                //    }
                                //    catch (Exception)
                                //    {
                                //        if (retrytimes < 2)
                                //        {
                                //            Task.Delay(1200); retrytimes++; goto retry;
                                //        }
                                //        //throw;
                                //    }
                            }
                            else if (modifierKeys == Keys.Shift && !pagePaste2GjcoolOCR)//&& !PagePaste2GjcoolOCR_ing)
                            {
                                //toOCR(br.OCRSiteTitle.GJcool);
                                Form1.playSound(Form1.soundLike.press, true);
                                toOCR(PagePast2OCRsite);
                            }
                            break;
                        case BrowserOPMode.seleniumGet:

                            break;
                        default:
                            break;
                    }
                    if (modifierKeys != Keys.Shift && !pagePaste2GjcoolOCR)
                    {
                        //備份textbox1的內容
                        undoRecord();
                        Task.WaitAll();
                        Application.DoEvents();
                        //設定textbox1的內容以備編輯
                        //string nextpagetextBox1Text_Default = Clipboard.GetText();
                        string nextpagetextBox1Text_Default = br.Quickedit_data_textboxTxt;
                        //textBox1.Text = CnText.BooksPunctuation(ref nextpagetextBox1Text_Default, false);// + Environment.NewLine + Environment.NewLine + Environment.NewLine + textBox1.Text;                    

                        string chkX = string.Empty;
                        if (nextpagetextBox1Text_Default != string.Empty)
                        {
                            if (!ocrTextMode)
                            {
                                textBox1.Text = nextpagetextBox1Text_Default;
                                //clearBracketsInsidePairsBrackets();
                                //nextpagetextBox1Text_Default = textBox1.Text;
                            }
                            if (!notBooksPunctuation)
                                if (modifierKeys != Keys.Shift && ModifierKeys != Keys.Shift)//按下Shift鍵即欲逕送去OCR，故不必再標書名號、篇名號了
                                    chkX = CnText.BooksPunctuation(ref nextpagetextBox1Text_Default, false);
                                else
                                    chkX = nextpagetextBox1Text_Default;
                            textBox1.Text = notBooksPunctuation ? nextpagetextBox1Text_Default : chkX;

                        }
                        #region 如果已經編輯，則將Form1移至旁邊
                        if (!ocrTextMode && keyinTextMode && autoTestPositionAvoidance)
                        {
                            //int lf = Left, tp = Top;
                            if (Form1Pos.X != Left && Form1Pos.Y != Top && (Left < 1235 || Top < 687)) { Form1Pos.X = Left; Form1Pos.Y = Top; }
                            if (chkX != string.Empty && chkX.IndexOf("<p>") > -1 && chkX == nextpagetextBox1Text_Default && (Form1Pos.X < ((double)1235 / SystemInformation.PrimaryMonitorSize.Width) * 1235 || Form1Pos.Y < ((double)687 / SystemInformation.PrimaryMonitorSize.Height) * 687))
                            //if (chkX == nextpagetextBox1Text_Default && (Left < SystemInformation.PrimaryMonitorSize.Width / 587 * 587 || Top < SystemInformation.PrimaryMonitorSize.Height / 1035 * 1035))
                            {
                                Top = 687; Left = 1235;//移動表單以供檢視已編輯的情形
                                if (MessageBoxShowOKCancelExclamationDefaultDesktopOnly("操作介面是否要回原位？") == DialogResult.OK)
                                {
                                    Left = Form1Pos.X; Top = Form1Pos.Y;
                                }
                            }
                            else
                            {
                                if (chkX != string.Empty && chkX.IndexOf("<p>") == -1 && (chkX.IndexOf("{") == -1 || chkX.IndexOf("}") == -1)
                                    && (Top >= 687 || Left >= 1235))
                                //{ Top = 200; Left = 800; }
                                { Top = Form1Pos.Y; Left = Form1Pos.X; }
                            }
                        }
                        #endregion
                    }
                }//end if (keyinTextMode)

                //如果該書沒有「OCR_MATCH」tag的話，即不是上下臨近頁牽連編輯模式者：
                if (!check_the_adjacent_pages)
                {
                    switch (browsrOPMode)
                    {
                        case BrowserOPMode.appActivateByName:
                            Task.Delay(500).Wait();
                            SendKeys.Send("^{PGUP}");//回瀏覽器上一頁籤檢查文本是否如願貼好
                            break;
                        //Selenium模式不須回上一頁，但若是自己keyin的失誤，漏字多字掉字……還是自行回去看看比較好，目前是原頁顯示頁面後才會再翻去下一頁書圖
                        case BrowserOPMode.seleniumNew:
                            break;
                        case BrowserOPMode.seleniumGet:
                            break;
                        default:
                            break;
                    }
                }
            }
            #endregion
            #region 檢查是否真有到下一頁或上一頁
            if (autoPaste2QuickEdit && int.Parse(_currentPageNum) > 2 &&
                Math.Abs(int.Parse(_currentPageNum) - int.Parse(br.WindowHandles["currentPageNum"])) != 1)
            {

                MessageBoxShowOKExclamationDefaultDesktopOnly("★★★★★★翻頁失敗，請注意！！！！！！", "●翻頁失敗！" +
                    Environment.NewLine + Environment.NewLine +
                    "現在textBox3的頁碼是：" + _currentPageNum +
                    "存放機制的頁碼是：" + br.WindowHandles["currentPageNum"]);
                if (fastMode) FastModeSwitcher();
                return false;
            }
            //else
            //    WindowHandles["currentPageNum"] = _currentPageNum;

            #endregion
            #region 在自動連續輸入模式下若是欽定四庫全書則自動輸入Kanripo或《國學大師》的文本,如同按下Ctrl + Shift + 4
            if (autoPaste2QuickEdit && eKeyCode == Keys.PageDown)
            {

                //if (textBox1.TextLength < 100 && (textBox1.Text.IndexOf("|") == -1 || textBox1.Text.IndexOf("<p>") == -1))
                //&& br.Div_generic_TextBoxFrame.Text==string.Empty)
                //&& br.pageUBound < int.Parse(_currentPageNum))
                if (textBox1.TextLength == 0 || (textBox1.Text.IndexOf("|") == -1 && textBox1.Text.IndexOf("<p>") == -1))
                {
                    undoRecord(); stopUndoRec = true; PauseEvents();

                    if (autoNextVolumnContextMark)
                        if (autoExecuteSKQSContextMark())
                            AvailableInUseBothKeysMouse();
                    //if (textBox1.TextLength == 0 && br.pageUBound >= int.Parse(_currentPageNum))
                    if (autoNextVolumnContextMark && br.pageUBound >= int.Parse(_currentPageNum))
                    {
                        if (Clipboard.GetText().Contains("-1a]"))
                        {
                            runWordMacro("中國哲學書電子化計劃.國學大師_Kanripo_四庫全書本轉來");
                            AutoMarkTitleParagraph();
                            AvailableInUseBothKeysMouse();
                        }
                        if (Clipboard.GetText().Contains(":KR"))
                        {
                            runWordMacro("中國哲學書電子化計劃.Kanripo_GitHub轉來");
                            AutoMarkTitleParagraph();
                            AvailableInUseBothKeysMouse();
                        }
                    }
                    else
                        if (br.pageUBound < int.Parse(_currentPageNum))
                        if (!gotoNextChapter_FormatContentInput_SKQS())
                        {
                            //Debugger.Break();
                            if (!_eventsEnabled) _eventsEnabled = true;
                            //if (!Visible) EndUpdate();//show_nICo();

                            return false;
                        }

                    undoRecord(); stopUndoRec = false; ResumeEvents();
                }
            }
            #endregion

            /////////////這先取消，交由呼叫端處理 20240904
            //if (stayInHere && !pagePaste2GjcoolOCR) AvailableInUseBothKeysMouse();//this.Activate();
            return true;

        }
        /// <summary>
        /// 自動讀入下一章節chapter（冊）的文本
        /// 如各書首頁資訊顯示「拆分成n冊」的「冊」
        /// </summary>
        /// <returns>失敗則傳回false</returns>
        private bool gotoNextChapter_FormatContentInput_SKQS()
        {
            //從每冊Chapter的最後一頁或其後開始

            if (IsDriverInvalid()) return false;

            //點擊書名連結控制項
            Title_Linkbox?.Click();
            //點擊下一個chapter的連結
            string nextChapterSelector = NextChapterSelector;
            if (nextChapterSelector.IsNullOrEmpty()) return false;
            //點擊本書首頁的冊連結
            OpenQA.Selenium.IWebElement iwe = WaitFindWebElementBySelector_ToBeClickable(nextChapterSelector, 5);
            if (iwe == null) return false;
            iwe.Click();
            //翻到本書該冊的第1頁
            string Div_generic_TextBoxFrameGetTextContent = Div_generic_TextBoxFrame?.GetAttribute("textContent");
            if (Div_generic_TextBoxFrameGetTextContent == "●\t"
                || Div_generic_TextBoxFrameGetTextContent == "●<p>"
                || Div_generic_TextBoxFrameGetTextContent == "●"
                )//如果是WordVBA.中國哲學書電子化計劃.新頁面() 產生的新頁面
                inputSKQSFrontPage();
            else
            {
                while (Div_generic_TextBoxFrame == null || Div_generic_TextBoxFrame?.GetAttribute("textContent") == string.Empty)//當冊首頁圖文對照之文字框沒有內容時才轉到下一頁
                {
                    if (GetPageNumFromUrl(driver.Url) < 4)
                        NextPage?.Click();
                    else
                        return false;
                }
                formatContentInput_SKQS(Edit_Linkbox?.GetAttribute("href"));
            }
            return true;
        }

        /// <summary>
        /// 設定是否要自動複製下一卷/單位文本
        /// </summary>
        bool autoNextVolumnContextMark = true;

        #region 在自動連續輸入模式下若是欽定四庫全書則自動輸入Kanripo或《國學大師》的文本,如同按下Ctrl + Shift + 4
        /// <summary>
        /// 在自動連續輸入模式下若是欽定四庫全書則自動輸入Kanripo或《國學大師》的文本,如同按下Ctrl + Shift + 4
        /// </summary>
        /// <returns>失敗或未執行則傳回false</returns>
        private bool autoExecuteSKQSContextMark()
        {
            //if (br.Version_LinkBox != null && textBox1.TextLength < 100
            //                && (textBox1.Text.IndexOf("|") == -1 || textBox1.Text.IndexOf("<p>") == -1)
            //                && br.pageUBound >= int.Parse(_currentPageNum))
            //已非四庫本專用了！20250418
            if (textBox1.TextLength < 100
                        && (textBox1.Text.IndexOf("|") == -1 || textBox1.Text.IndexOf("<p>") == -1)
                        && br.pageUBound >= int.Parse(_currentPageNum))
            //&& br.pageUBound >= GetPageNumFromUrl(textBox3Text))
            {
                autoPastetoCtextQuitEditTextboxCancel = true;
                //if (fastMode) FastModeSwitcher();
                textBox1.Clear();
                //Task.Run(() =>
                //{
                //讀入四庫全書文本
                br.SikuQuanshu_SKQSContextCopyReading();
                AutoMarkTitleParagraph();
                //});
                return true;
            }
            return false;
        }
        #endregion

        /// <summary>
        /// 自動標題、段落標記
        /// </summary>
        /// <returns>成功傳回true</returns>
        internal bool AutoMarkTitleParagraph()
        {
            if (textBox1.TextLength > 100)
            {

                if (fastMode && (GetLineText(textBox1.Text, 0).Contains("目錄")
                        || GetLineText(textBox1.Text, 0).Contains("目録")))
                    return true;
                //br.driver.SwitchTo().Window(br.driver.CurrentWindowHandle);
                //bringBackMousePosFrmCenter();
                //textBox1.Refresh();
                //if (DialogResult.Cancel == MessageBoxShowOKCancelExclamationDefaultDesktopOnly("要自動標題標記、段落標記否？")) return true;


                #region 自動標題標記

                foreach (var item in TitleLeadingSpacesCount)
                {
                    //取得該階標題的前置空格值後，找到textBox1中第1個符合此空格的行/段
                    List<Paragraph> paragraphs = _document.GetParagraphs();
                    for (int i = 0; i < paragraphs.Count - 1; i++)
                    {
                        int spscount = _leadingSpacesRegex.Match(paragraphs[i].Text).Value.Length;
                        if (spscount == item)//找到textBox1中第1個符合此空格的行/段（從頭標題起）
                        {
                            textBox1.Select(paragraphs[i].Start, spscount);
                            autoMarkTitles();
                            break;
                        }
                    }

                }
                #endregion
                #region 自動段落標記
                playSound(soundLike.over, true);
                if (TitleLeadingSpacesCount.Count > 0)
                    paragraphMarkingAccordingFirstOneLineLength();
                #endregion

                return true;
            }
            return false;
        }


        internal void runWordMacro(string runName)
        {
            if (!isClipBoardAvailable_Text()) return;
            string xClpBd = Clipboard.GetText();
            if (autoPaste2QuickEdit && xClpBd.Length < 100 && xClpBd.IndexOf("Bot", StringComparison.Ordinal) == -1) return;
            //&& runName!="Docs.中國哲學書電子化計劃_只保留正文注文_且注文前後加括弧_貼到古籍酷自動標點"
            Color C = this.BackColor; this.BackColor = Color.Green;
            SystemSounds.Hand.Play();
            hideToNICo();
            if (this.Visible)
            {
                this.WindowState = FormWindowState.Minimized;
                this.Hide();
            }
            Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office
                                    .Interop.Word.Application();
            try
            {
                //Task.Run(() =>
                //{
                //    switch (runName)
                //    {
                //        case "中國哲學書電子化計劃.國學大師_Kanripo_四庫全書本轉來":
                //            //自動翻至Kanripo下一卷並複製其文字


                //            break;
                //        default:
                //            break;
                //    }
                //});
                appWord.Run(runName);
                Application.DoEvents();

            }
            catch (Exception e)
            {
                if (e.HResult.ToString() != "0x80010105")
                {/* 
                  System.Runtime.InteropServices.COMException
                HResult=0x80010105
                Message=伺服器丟出一個例外。 (發生例外狀況於 HRESULT: 0x80010105 (RPC_E_SERVERFAULT))
                Source=Microsoft.Office.Interop.Word
                    */
                    MessageBoxShowOKExclamationDefaultDesktopOnly(e.HResult + e.Message);
                    goto finish;
                    //throw;
                }
            }

            //自剪貼簿擷取Word VBA結果
            while (!isClipBoardAvailable_Text()) { }
            xClpBd = Clipboard.GetText();
            switch (runName)
            {
                case "中國哲學書電子化計劃.清除頁前的分段符號":
                    break;
                case "中國哲學書電子化計劃.撤掉與書圖的對應_脫鉤":
                    break;

                default:
                    //清除多餘的空行,排除卷末的空行                                        
                    if (xClpBd.Length > 100)
                    {
                        textBox1.Text = xClpBd.Substring(0, xClpBd.Length - 100).Replace(Environment.NewLine + Environment.NewLine, Environment.NewLine)
                            + xClpBd.Substring(xClpBd.Length - 100);
                    }
                    //else
                    //    textBox1.Text = xClpBd;

                    switch (runName)
                    {
                        case "漢籍電子文獻資料庫文本整理_以轉貼到中國哲學書電子化計劃":

                            textBox1.Text = xClpBd;
                            saveText();
                            break;
                        case "中國哲學書電子化計劃.國學大師_Kanripo_四庫全書本轉來":
                            using (GXDS gxds = new GXDS(this)) { gxds.StandardizeSKQSContext(ref xClpBd); }

                            #region 若是《四庫全書》本則將頭尾縮排的文本行/段還原

                            standardizeSKQSContext(ref xClpBd);
                            textBox1.Text = xClpBd;
                            //standardizeSKQSContext();
                            //xClpBd = textBox1.Text;
                            #endregion
                            saveText();
                            Clipboard.Clear();
                            //xClpBd = xClpBd.Replace(" /\v\v", Environment.NewLine).Replace("\v", Environment.NewLine)                                    
                            //        .Replace(" /", "");
                            //        //這要做標題判斷，不能取代掉.Replace(Environment.NewLine + Environment.NewLine, Environment.NewLine)
                            //xClpBd = "*欽定四庫全書<p>" + xClpBd.Substring(xClpBd.IndexOf("欽定《四庫全書》") + "欽定《四庫全書》".Length);
                            //bringBackMousePosFrmCenter();//交由呼叫端處理
                            ////AvailableInUseBothKeysMouse();
                            break;
                        case "中國哲學書電子化計劃.Kanripo_GitHub轉來":
                            using (GXDS gxds = new GXDS(this)) { gxds.StandardizeSKQSContext(ref xClpBd); }

                            #region 若是《四庫全書》本則將頭尾縮排的文本行/段還原

                            standardizeSKQSContext(ref xClpBd);
                            textBox1.Text = xClpBd;
                            //standardizeSKQSContext();
                            //xClpBd = textBox1.Text;
                            #endregion
                            saveText();
                            Clipboard.Clear();
                            //xClpBd = xClpBd.Replace(" /\v\v", Environment.NewLine).Replace("\v", Environment.NewLine)                                    
                            //        .Replace(" /", "");
                            //        //這要做標題判斷，不能取代掉.Replace(Environment.NewLine + Environment.NewLine, Environment.NewLine)
                            //xClpBd = "*欽定四庫全書<p>" + xClpBd.Substring(xClpBd.IndexOf("欽定《四庫全書》") + "欽定《四庫全書》".Length);
                            //bringBackMousePosFrmCenter();//交由呼叫端處理
                            ////AvailableInUseBothKeysMouse();
                            break;
                        default:
                            break;
                    }

                    // 再進行賦值
                    if (textBox1.Text != xClpBd) textBox1.Text = xClpBd;
                    textBox1.Select(0, 0);
                    textBox1.ScrollToCaret();
                    break;
            }
            try
            {
                if (runName != "checkEditingOfPreviousVersion")
                    appWord.Quit(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);
            }
            catch (Exception)
            {
                //word 已被關閉
                //throw;
            }

        finish:
            this.BackColor = C;
            //NormalLineParaLength = 0;

            if (runName == "中國哲學書電子化計劃.清除頁前的分段符號")
            {
                //TopMost = false;
                br.ChromeSetFocus();
                br.BringToFront("chrome");
            }
            else
                show_nICo(ModifierKeys);
        }

        /// <summary>
        /// 若是《四庫全書》本則將頭尾縮排的文本行/段還原 20250220
        /// 以利自動標題標記。在手動編輯時再還原其因超平抬而縮排的文本行 20250315
        /// 以字串string xClpBd操作，不藉由textBox1
        /// </summary>
        /// <param name="xClpBd">要操作的字串</param>
        private void standardizeSKQSContext(ref string xClpBd)
        {
            if (xClpBd == string.Empty) return;
            //if (lines_perPage == 0) lines_perPage = 16;//return;

            if (xClpBd.IndexOf("欽定四庫全書") > -1)
            {
                //取得縮排的行/段數（通常是2頁）
                Document document = new Document(xClpBd);
                List<Paragraph> paragraphs = document.GetParagraphs();
                Range range = null;
                int lastParaIndex = 2, leadingSpaceCount = _leadingSpacesRegex.Match(paragraphs[0].Text).Value.Length;//首行「欽定四庫全書」前的縮排空格數
                int lineParaCount;
                if (leadingSpaceCount == 0) return;
                for (int i = 2; i < paragraphs.Count; i++)//通常第2行/段是書名，會超平抬頭，故從第3行/段起算
                {
                    range = new Range(document, paragraphs[2].Start, paragraphs[i].End, document.Content);
                    lineParaCount = countLinesPerPage(range.Text);
                    lineParaCount = lineParaCount % 2 == 0 ? lineParaCount : lineParaCount + 1;//若獨立小注之末可能只有單行，不會成雙
                    if (lines_perPage * 2 - 2 * 2 <= lineParaCount)
                    {
                        lastParaIndex = i;
                        break;
                    }
                }
                if (range != null)
                {
                    for (int i = 0; i < leadingSpaceCount; i++)
                    {
                        outdent_ConvexRow(ref range);
                        range.End = document.GetParagraphs()[lastParaIndex].End;//●●●●●●●●●●●●●●●
                    }
                }
                //以上卷首縮排的處理
                //playSound(soundLike.processing, true);
                //以下卷尾縮排的處理

                paragraphs = document.GetParagraphs();//更新paragraphs，因為_document（也就是textBox1內容已經因為前面的凸排而更動了）                
                range = null;

                for (int i = paragraphs.Count - 3; i > -1; i--)//通常末1行/段是書名，會超平抬頭，故末第2行/段起算
                {
                    if (lastParaIndex >= i) break;//若單位文字長度太短，前面卷首已經處理過了，就不要再處理了，否則成了凸排2次！感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主20250224
                    range = new Range(document, paragraphs[i].Start, paragraphs[paragraphs.Count - 2].End, document.Content);
                    lineParaCount = countLinesPerPage(range.Text);//若獨立小注之尾可能只有單行，不會成雙
                    lineParaCount = lineParaCount % 2 == 0 ? lineParaCount : lineParaCount + 1;
                    if (lines_perPage * 2 - 1 * 2 <= lineParaCount)
                    {
                        lastParaIndex = i;
                        break;
                    }
                }
                if (range != null)
                {
                    for (int i = 0; i < leadingSpaceCount; i++)
                    {

                        outdent_ConvexRow(ref range);
                        paragraphs = document.GetParagraphs();
                    }
                }
                xClpBd = document.Text;
            }
        }
        /// <summary>
        /// 若是《四庫全書》本則將頭尾縮排的文本行/段還原 20250220
        /// 以textBox1操作
        /// </summary>        
        private void standardizeSKQSContext()
        {
            //if (lines_perPage == 0) lines_perPage = 16;//return;

            if (textBox1.Text.IndexOf("欽定四庫全書") > -1)
            {
                //取得縮排的行/段數（通常是2頁）
                int leadingSpaceCount = _leadingSpacesRegex.Match(GetLineText(_document.Text, _document.Text.IndexOf(Environment.NewLine))).Value.Length;//首行「欽定四庫全書」前的縮排空格數
                Document document = _document;
                Range range = null;
                List<Paragraph> paragraphs = document.GetParagraphs();
                int lineParaCount, lastParaIndex = 2; ;
                if (leadingSpaceCount == 0) return;
                for (int i = 2; i < paragraphs.Count; i++)//通常第2行/段是書名，會超平抬頭，故從第3行/段起算
                {
                    range = new Range(document, paragraphs[2].Start, paragraphs[i].End, document.Content);
                    lineParaCount = countLinesPerPage(range.Text);
                    lineParaCount = lineParaCount % 2 == 0 ? lineParaCount : lineParaCount + 1;//若獨立小注之末可能只有單行，不會成雙
                    if (lines_perPage * 2 - 2 * 2 <= lineParaCount)
                    {
                        lastParaIndex = i;
                        break;
                    }
                }
                if (range != null)
                {
                    for (int i = 0; i < leadingSpaceCount; i++)
                    {
                        textBox1.Select(paragraphs[2].Start, range.End - paragraphs[2].Start);
                        //textBox1.Select(range.Start, range.End- range.Start);https://ctext.org/wiki.pl?if=en&chapter=443808&action=editchapter#42
                        outdent_ConvexRow();
                        range.End = document.GetParagraphs()[lastParaIndex].End;
                    }
                }
                //以上卷首縮排的處理
                playSound(soundLike.processing, true);
                //以下卷尾縮排的處理
                range = null;
                paragraphs = document.GetParagraphs();//更新paragraphs，因為_document（也就是textBox1內容已經因為前面的凸排而更動了）                
                for (int i = paragraphs.Count - 3; i > -1; i--)//通常末1行/段是書名，會超平抬頭，故末第2行/段起算
                {
                    if (lastParaIndex >= i) break;//若單位文字長度太短，前面卷首已經處理過了，就不要再處理了，否則成了凸排2次！感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主20250224
                    range = new Range(document, paragraphs[i].Start, paragraphs[paragraphs.Count - 2].End, document.Content);
                    lineParaCount = countLinesPerPage(range.Text);//若獨立小注之尾可能只有單行，不會成雙
                    lineParaCount = lineParaCount % 2 == 0 ? lineParaCount : lineParaCount + 1;
                    if (lines_perPage * 2 - 1 * 2 <= lineParaCount)
                    {
                        lastParaIndex = i;
                        break;
                    }
                }
                if (range != null)
                {
                    for (int i = 0; i < leadingSpaceCount; i++)
                    {
                        textBox1.Select(paragraphs[lastParaIndex].Start, range.End - paragraphs[lastParaIndex].Start);
                        outdent_ConvexRow();
                        paragraphs = document.GetParagraphs();
                    }
                }

            }
        }

        /// <summary>
        /// 等待剪貼簿內的文字可用時
        /// </summary>
        /// <returns></returns>
        public static bool isClipBoardAvailable_Text(int waitMilliSecond = 1000)
        {// creedit with chatGPT：Clipboard Availability in C#：https://www.facebook.com/oscarsun72/posts/pfbid0dhv46wssuupa5PfH6RTNSZF58wUVbE6jehnQuYF9HtE9kozDBzCvjsowDkZTxkmcl
            /*在 C# 的 System.Windows.Forms 中，可以使用 Clipboard.ContainsData 或 Clipboard.ContainsText 方法來確定剪貼簿是否可用。*/
            DateTime dt = DateTime.Now;
        retry:
            if (DateTime.Now.Subtract(dt).TotalSeconds > 2)
            {
                //if (MessageBox.Show("剪貼簿檢查已逾3秒，是否繼續？", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly) == DialogResult.Cancel)
                return true;
            }
            try
            {
                if (!Clipboard.ContainsText())
                {
                    Thread.Sleep(waitMilliSecond);
                    Task.Delay(waitMilliSecond);
                }

            }
            catch (Exception)
            {
                Thread.Sleep(waitMilliSecond);
                Task.Delay(waitMilliSecond); goto retry;
                throw;
            }
            return Clipboard.ContainsText();
        }


        /// <summary>
        /// 儲存當前文本的標名+副檔名
        /// </summary>
        const string fName_to_Save_Txt = "cText.txt";

        /// <summary>
        /// 儲存當前文本的全檔名；預設路徑在Dropbox預設安裝路徑
        /// 20230313 creedit with Adrenaline
        /// </summary>
        internal string FName_to_Save_Txt_fullname { get => Path.Combine(dropBoxPathIncldBackSlash + fName_to_Save_Txt); }

        //internal string FName_to_Save_Txt_fullname { get { return dropBoxPathIncldBackSlash + fName_to_Save_Txt; } }
        /*菩薩慈悲： 請這兩個程式碼是否與等價？ 
          * Yes, these two code snippets are equivalent.
            They both define an internal property called
            `FName_to_Save_Txt_fullname` that returns a string value made up of a file path.
            In the first code snippet, an expression-bodied member syntax is used to define the property, and it concatenates the
            `dropBoxPathIncldBackSlash` and
            `fName_to_Save_Txt` values using the
            `Path.Combine` method.
            In the second code snippet, a property with a getter is used instead, and it returns a string produced by concatenating
            `dropBoxPathIncldBackSlash` and
            `fName_to_Save_Txt`.
         */

        /// <summary>
        /// 儲存當前文本以備份；預設路徑在Dropbox預設安裝路徑
        /// 按F5以載入所備份（刷新textBox1的內容）
        /// </summary>
        internal void saveText()
        {
            //C# 對文字檔案的幾種讀寫方法總結:https://codertw.com/%E7%A8%8B%E5%BC%8F%E8%AA%9E%E8%A8%80/542361/
            string str1 = textBox1.Text, f = dropBoxPathIncldBackSlash + ((ocrTextMode & PasteOcrResultFisrtMode) ? "cTextOCR.txt" : fName_to_Save_Txt);
            try
            {
                using (var streamWriter = new StreamWriter(f)) { streamWriter.Write(str1); }
                //File.WriteAllText(f, str1, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                //Task.WaitAny();
                //Thread.Sleep(TimeSpan.FromSeconds(2));
                //File.WriteAllText(f, str1, Encoding.UTF8);
                //throw;
                Console.WriteLine(ex.HResult + ":" + ex.Message);
                switch (ex.HResult)
                {
                    case -2147024809:
                        if (ex.Message.StartsWith("無法將索引") && ex.Message.EndsWith("轉譯為指定的字碼頁。"))//20240921 應該是由於在分段/行或還原時時切到了surrogate的字
                            return;
                        break;
                    default:
                        Debugger.Break();
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                        break;
                }

                return;
            }
            // 也可以指定編碼方式 File.WriteAllText(@”c:\temp\test\ascii-2.txt”, str1, Encoding.ASCII);

            if (keyinTextMode && !autoPaste2QuickEdit)
            {
                //取得OCR所匯出的檔案路徑
                #region 再檢查瀏覽器下載目錄並取得 ：
                Task.Run(() =>
                {
                    string downloadDirectory = br.DownloadDirectory_Chrome;
                    //string downloadImgFullName = dropBoxPathIncldBackSlash + "Ctext_Page_Image.png";
                    string downloadImgFullName = MydocumentsPathIncldBackSlash + "CtextTempFiles\\Ctext_Page_Image.png";
                    if (br.ChkDownloadDirectory_Chrome(downloadImgFullName, downloadDirectory))
                    {
                        #endregion
                        string filePath = Path.Combine(downloadDirectory,
                            Path.GetFileNameWithoutExtension(downloadImgFullName) + ".txt");//@"X:\Ctext_Page_Image.txt";
                                                                                            //刪除之前的檔案，以免因檔案存在而被下載端重新命名

                        if (File.Exists(filePath)) File.Delete(filePath);
                    }
                });
            }
        }

        /// <summary>
        /// 由所儲存備份的文本重新載入到 textBox1
        /// </summary>
        private void loadText()
        {
            if (!File.Exists(dropBoxPathIncldBackSlash + fName_to_Save_Txt))
            {
                MessageBoxShowOKExclamationDefaultDesktopOnly("找不到檔案： " + dropBoxPathIncldBackSlash + fName_to_Save_Txt);
                return;
            }
            //C# 對文字檔案的幾種讀寫方法總結:https://codertw.com/%E7%A8%8B%E5%BC%8F%E8%AA%9E%E8%A8%80/542361/
            textBox1.Text = File.ReadAllText(dropBoxPathIncldBackSlash + fName_to_Save_Txt);
        }

        #region browsers 

        /// <summary>
        /// 取得預設瀏覽器名稱
        /// </summary>
        /// <returns></returns>
        static public string GetWebBrowserName()
        {

            //https://stackoverflow.com/questions/13621467/how-to-find-default-web-browser-using-c
            //Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoic
            //電腦\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice            
            RegistryKey userChoiceKey = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice");
            object progIdValue = userChoiceKey.GetValue("Progid");
            if (progIdValue == null)
            {
                MessageBox.Show("BrowserApplication.Unknown;");
                return "chrome";
                //browser = BrowserApplication.Unknown;
                //break;
            }
            string progId = progIdValue.ToString();
            progId = progId.Substring(0, progId.IndexOf(".") == -1 ? progId.Length : progId.IndexOf("."));
            switch (progId)
            {
                case "IE.HTTP":
                    return "iexplore";
                case "FirefoxURL":
                    return "firefox";
                case "ChromeHTML":
                    return "chrome";
                case "BraveHTML":
                    return "brave";
                case "OperaStable":
                    return "Opera";
                case "SafariHTML":
                    return "Safari";
                case "AppXq0fevzme2pys62n3e0fbqa7peapykr8v":
                    //browser = BrowserApplication.Edge;
                    return "msedge";
                default:
                    return "chrome";
            }

            /*
            //https://cybarlab.com/web-browser-name-in-c-sharp            
            string WebBrowserName = string.Empty;
            WebBrowserName = GetDefaultWebBrowserFilePath();
            return WebBrowserName.Substring(WebBrowserName.LastIndexOf("\\") + 1, WebBrowserName.LastIndexOf(".exe") - WebBrowserName.LastIndexOf("\\") - 1);
            */

            /*
            try
            {
                HttpContext context = HttpContext.Current;//https://docs.microsoft.com/zh-tw/dotnet/api/system.web.httpcontext.current?view=netframework-4.8&f1url=%3FappId%3DDev16IDEF1%26l%3DZH-TW%26k%3Dk(System.Web.HttpContext.Current)%3Bk(TargetFrameworkMoniker-.NETFramework%2CVersion%253Dv4.8)%3Bk(DevLang-csharp)%26rd%3Dtrue
                WebBrowserName = HttpContext.Current.Request.Browser.Browser + " " + HttpContext.Current.Request.Browser.Version;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            */

            //https://stackoverflow.com/questions/13621467/how-to-find-default-web-browser-using-c
        }

        /// <summary>
        /// 取得預設瀏覽器執行檔全檔名
        /// </summary>
        /// <returns></returns>
        internal static string getDefaultBrowserEXE()
        {
            string userProfilePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);//https://stackoverflow.com/questions/38252474/c-sharp-service-how-to-get-user-profile-folder-path
            switch (defaultBrowserName)
            {

                case "iexplore":
                    return @"C:\Program Files\Internet Explorer\iexplore.exe";
                case "firefox":
                    if (!File.Exists(@"W:\PortableApps\PortableApps\FirefoxPortable\App\Firefox64\firefox.exe"))
                        return @"C:\Program Files\Mozilla Firefox\firefox.exe";
                    else
                        return @"W:\PortableApps\PortableApps\FirefoxPortable\App\Firefox64\firefox.exe";
                case "brave":
                    if (!File.Exists(userProfilePath + @"\AppData\Local\BraveSoftware\Brave-Browser\Application\brave.exe"))
                        return @"C:\Program Files (x86)\BraveSoftware\Brave-Browser\Application\brave.exe";
                    else
                        return userProfilePath + @"\AppData\Local\BraveSoftware\Brave-Browser\Application\brave.exe";
                case "vivaldi":
                    return userProfilePath + @"\AppData\Local\Vivaldi\Application\vivaldi.exe";
                case "Opera":
                    return "";
                case "Safari":
                    return "";
                case "edge":
                    return @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe";//"msedge"

                case "chrome":// "ChromeHTML"://, "google chrome": '"chrome"
                    if (File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\GoogleChromePortable\App\Chrome-bin\chrome.exe"))
                        return "C:\\Users\\ssz3\\Documents\\GoogleChromePortable\\App\\Chrome-bin\\chrome.exe";
                    else if (File.Exists(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"))
                        return @"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe";
                    //return @"W:\PortableApps\PortableApps\GoogleChromePortable\GoogleChromePortable.exe";
                    else if (File.Exists(@"C:\Program Files\Google\Chrome\Application\chrome.exe"))
                        return @"C:\Program Files\Google\Chrome\Application\chrome.exe";
                    else if (File.Exists(@"W:\PortableApps\PortableApps\GoogleChromePortable\App\Chrome-bin\chrome.exe"))
                        return @"W:\PortableApps\PortableApps\GoogleChromePortable\App\Chrome-bin\chrome.exe";
                    else
                        return @"W:\PortableApps\PortableApps\GoogleChromePortable64\App\Chrome-bin\chrome.exe";

                default:
                    return GetDefaultWebBrowserFilePath();
            }
        }

        /// <summary>
        /// 取得預設瀏覽器執行檔路徑
        /// </summary>
        /// <returns></returns>
        static private string GetDefaultWebBrowserFilePath()//chrome-extension://lcghoajegeldpfkfaejegfobkapnemjl/sandbox.html?src=https%3A%2F%2Fwww.796t.com%2Fcontent%2F1546728863.html
        {
            //舊的如此，抓不準！廢罝不用！
            //從登錄檔中讀取預設瀏覽器可執行檔案路徑
            RegistryKey key = Registry.ClassesRoot.OpenSubKey(@"http\shell\open\command\");
            string s = key.GetValue("").ToString();
            return s.Substring(s.IndexOf(@":\") - 1, s.LastIndexOf(".exe") + 4 - 1);


            //s就是你的預設瀏覽器，不過後面帶了引數，把它截去，不過需要注意的是：不同的瀏覽器後面的引數不一樣！
            //"D:\Program Files (x86)\Google\Chrome\Application\chrome.exe" -- "%1"
        }

        /*
        //https://www.programmerall.com/article/8861915667/
        /// <summary>
        /// Get the path of the default browser
        /// </summary>
        /// <returns></returns>
        public String GetDefaultWebBrowserFilePath()
        {
            string _BrowserKey1 = @"Software\Clients\StartmenuInternet\";
            string _BrowserKey2 = @"\shell\open\command";

            RegistryKey _RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(_BrowserKey1, false);
            if (_RegistryKey == null)
                _RegistryKey = Registry.LocalMachine.OpenSubKey(_BrowserKey1, false);
            String _Result = _RegistryKey.GetValue("").ToString();
            _RegistryKey.Close();

            _RegistryKey = Registry.LocalMachine.OpenSubKey(_BrowserKey1 + _Result + _BrowserKey2);
            _Result = _RegistryKey.GetValue("").ToString();
            _RegistryKey.Close();

            if (_Result.Contains("\""))
            {
                _Result = _Result.TrimStart('"');
                _Result = _Result.Substring(0, _Result.IndexOf('"'));
            }
            return _Result;
        }
        */

        #endregion

        string processID;
        //https://stackoverflow.com/questions/58302052/c-microsoft-visualbasic-interaction-appactivate-no-effect
        [DllImport("user32.dll", SetLastError = true)]
        static extern void SwitchToThisWindow(IntPtr hWnd, bool turnOn);

        internal static string defaultBrowserName = string.Empty;//https://cybarlab.com/web-browser-name-in-c-sharp
        internal void appActivateByName()
        {
        //Process[] procsBrowser = Process.GetProcessesByName("chrome");
        tryagain:
            Process[] procsBrowser = Process.GetProcessesByName(defaultBrowserName);
            if (procsBrowser.Length <= 0)
            {
                //MessageBox.Show("Chrome is not running");
                if (defaultBrowserName != "chrome")
                {
                    defaultBrowserName = "chrome";
                    goto tryagain;
                }
                else if (defaultBrowserName != "brave")
                {
                    defaultBrowserName = "brave";
                    goto tryagain;
                }
                MessageBox.Show(defaultBrowserName + " is not running");
            }
            foreach (Process proc in procsBrowser)
            {
                if (proc.MainWindowHandle != IntPtr.Zero)
                    SwitchToThisWindow(proc.MainWindowHandle, true);
            }

        }

        void appActivateByID()
        { //https://docs.microsoft.com/zh-tw/dotnet/csharp/programming-guide/strings/how-to-determine-whether-a-string-represents-a-numeric-value
            int i = 0;
            if (processID == null || processID == "")
            {
                processID = textBox2.Text;
                bool result = int.TryParse(processID, out i); //i now = 108  
                if (!result)
                {
                    MessageBox.Show("plz input the id of process in the textbox2 ,then go on……");
                    return;
                }
            }
            var process = Process.GetProcessById(Int32.Parse(processID));
            if (process == null) { MessageBox.Show("plz input the id of process in the textbox2 ,then go on……"); return; }
            SwitchToThisWindow(process.MainWindowHandle, true);
            //ShowWindow(process.MainWindowHandle, SW_RESTORE);
            //SetForegroundWindow(process.MainWindowHandle);
        }

        /// <summary>
        /// 要比較兩個 System.Collections.ObjectModel.ReadOnlyCollection<string> 物件是否相同（包括順序和內容），可以使用 LINQ 提供的 SequenceEqual 方法。這個方法會逐一比較兩個集合中的元素，確保它們的順序和內容完全一致。
        /// 20240731 Copilot大菩薩：比較 ReadOnlyCollection<string> 物件
        /// </summary>
        /// <param name="ReadOnlyCollection1"></param>
        /// <param name="ReadOnlyCollection2"></param>
        /// <returns>若兩個集合完全相同（次序+內容）則傳回true，否則為false</returns>
        public static bool CompareReadOnlyCollection<T>(ReadOnlyCollection<T> ReadOnlyCollection1, ReadOnlyCollection<T> ReadOnlyCollection2)
        {
            return ReadOnlyCollection1.SequenceEqual(ReadOnlyCollection2);
        }
        /// <summary>
        /// 確定ReadOnlyCollection集合A是否包含B而B不包含A
        /// 20240731 Copilot大菩薩：比較 ReadOnlyCollection<string> 物件
        /// </summary>
        /// <param name="ReadOnlyCollectionA">疑似較大的集合</param>
        /// <param name="ReadOnlyCollectionB">疑似較小的集合</param>
        /// <returns>如果A包含B而B不包含A則傳回true，否則為false</returns>
        public static bool IsAContainBandBnotContainA_ReadOnlyCollection<T>(ReadOnlyCollection<T> ReadOnlyCollectionA, ReadOnlyCollection<T> ReadOnlyCollectionB)
        {
            // 你想要快速確定是否存在 tabWindowHandlesValid 中沒有但 tabWindowHandles 中卻有的元素，並回傳一個布林值（true 表示確實存在這樣的元素）。
            // 你可以使用 LINQ 的 Any 方法來實現這個需求。以下是範例程式碼：
            return ReadOnlyCollectionA.Any(handle => !ReadOnlyCollectionB.Contains(handle));
        }
        /// <summary>
        /// 找出ReadOnlyCollection物件A有B沒有與B有A沒有的元素出來
        /// </summary>
        /// <param name="ReadOnlyCollectionA"></param>
        /// <param name="ReadOnlyCollectionB"></param>
        /// <returns></returns>
        public static List<List<T>> FindNotInAorB_ReadOnlyCollection<T>(ReadOnlyCollection<T> ReadOnlyCollectionA, ReadOnlyCollection<T> ReadOnlyCollectionB)
        {
            //List<List<T>> list=new List<List<T>>();
            // 找出 ReadOnlyCollectionA 中有但 ReadOnlyCollectionB 中沒有的元素
            //var onlyInTabWindowHandles = ReadOnlyCollectionA.Except(ReadOnlyCollectionB).ToList();
            // 找出 ReadOnlyCollectionB 中有但 ReadOnlyCollectionA 中沒有的元素
            //var onlyInTabWindowHandlesValid = ReadOnlyCollectionB.Except(ReadOnlyCollectionA).ToList();            
            return new List<List<T>> { ReadOnlyCollectionA.Except(ReadOnlyCollectionB).ToList(), ReadOnlyCollectionB.Except(ReadOnlyCollectionA).ToList() };
        }

        /// <summary>
        /// for .BrowserOPMode.Selenium……    browsrOPMode!=BrowserOPMode.appActivateByName
        /// </summary>
        /// <param name="url">url to paste to</param>
        /// <param name="clear">whether clear the texts in quick edit box ;optional. if yes then set this param value to 「chkClearQuickedit_data_textboxTxtStr」 </param>
        /// <returns>執行不成功則傳回false</returns>
        private bool input2Ctext(string url, bool statyhere = false, string clear = "")
        {
            //if (!(url.IndexOf("&file=") > -1 && url.IndexOf("&page=") > -1 && url.IndexOf("&editwiki=") > -1 && url.EndsWith("#editor"))) return false;
            //也有可能是這種網址：https://ctext.org/library.pl?if=gb&file=34195&page=142&editwiki=826120#box(140,120,2,0)
            //if (!(url.IndexOf("&file=") > -1 && url.IndexOf("&page=") > -1 && url.IndexOf("&editwiki=") > -1 && url.IndexOf("#edit") == -1)) return false;
            if (!IsValidUrl＿keyDownCtrlAdd(url)) return false;

            br.driver = br.driver ?? br.DriverNew();
            //br.HideBrowserWindow(br.driver);
            //取得所有現行窗體（分頁頁籤）
            //System.Collections.ObjectModel.ReadOnlyCollection<string> tabWindowHandles = new ReadOnlyCollection<string>(new List<string>());
            //System.Collections.ObjectModel.ReadOnlyCollection<string> tabWindowHandles = br.driver.WindowHandles;
            //br.ConvertToReadOnlyCollection(br.GetValiOrdereddWindowHandles(br.driver));
            //br.ConvertToReadOnlyCollection(br.GetValidWindowHandles(br.driver));
            //br.ShowBrowserWindow(br.driver);
            //br.driver.SwitchTo().Window(br.GetCurrentWindowHandle(br.driver));
            //br.driver.SwitchTo().Window(br.driver.CurrentWindowHandle);
            //br.driver.Navigate().Refresh();

            //try
            //{
            //    if (!CompareReadOnlyCollection(tabWindowHandles, br.driver.WindowHandles))
            //        Debugger.Break();
            //}
            //catch (Exception ex)
            //{
            //    MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
            //    Debugger.Break();
            //}
            //檢查兩個視窗句柄集合是否相同

            //if (tabWindowHandles.Count == 0) return false;
            string currentWin = ""; bool Edited = false;
            try
            {
                //tabWindowHandles = br.driver.WindowHandles;
                //currentWin = br.driver.CurrentWindowHandle;
                string getCurrentWindowHandle = br.GetCurrentWindowHandle(br.driver);
                if (getCurrentWindowHandle != null) currentWin = getCurrentWindowHandle;
                else currentWin = br.GetValidWindowHandles(br.driver).Last();
            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case -2146233088: //"An unknown exception was encountered sending an HTTP request to the remote WebDriver server for URL http://localhost:6763/session/b17084f4c8e209d232d5a9eb18cf181a/window/handles. The exception message was: 傳送要求時發生錯誤。"
                        br.driver.Quit();
                        br.driver = null; br.driver = br.DriverNew();
                        //tabWindowHandles = br.driver.WindowHandles;
                        //tabWindowHandles = br.ConvertToReadOnlyCollection(br.GetValiOrdereddWindowHandles(br.driver));
                        break;
                    default:
                        throw;
                }
            }
            //手動輸入模式時。20241119：新增自動連續輸入時也可以
            string waitTabWindowHandle = string.Empty; DateTime dt;
            if (autoPastetoCtextQuitEditTextboxCancel && (keyinTextMode || autoPaste2QuickEdit))
            {
                //Task task = Task.Run(() =>
                //{
                //url = br.driver.Url;
                //string activeTabTitlel = br.driver.Title;
                //br.driver.SwitchTo();//br.driver.SwitchTo().Window(br.driver.CurrentWindowHandle).Url;
                //br.driver.SwitchTo().Window(br.getOriginalWindow); 
                //br.driver.SwitchTo();
                //url = br.driver.Url;
                //br.driver.SwitchTo().Window(br.driver.WindowHandles[0]);
                //br.driver.SwitchTo().Window(br.driver.WindowHandles.First());
                ////url = br.driver.Url;
                //20230103 目前無法抓到正在作用中的視窗，只能以使用者習慣，通常在用的都是最後一個視窗，先試試 感恩感恩　南無阿彌陀佛
                //br.driver.SwitchTo().Window(br.driver.WindowHandles.Last());//br.driver.SwitchTo().Window(br.driver.CurrentWindowHandle).Url;
                //20230103 目前所知也只能用以下的笨方法了，雖然土，但管用。∴就不必上一行多此一舉了。感恩感恩　南無阿彌陀佛
                //檢查driver物件是否為空

                #region 先檢查是否有已開啟的「編輯」頁尚未送出儲存(因為許多異體字須一次取代，往往會打開一個chapter單位來edit) 其網址有「&action=editchapter」關鍵字，如：https://ctext.org/wiki.pl?if=en&chapter=687756&action=editchapter#12450
                //mark:在版本netframework-4.8 之前的環境，好像無效（在母校華岡學習雲測試後的結果，似並不會執行這個檢查，該機唯有4.6.1版）
                bool waitUpdate = false;
                //Task wait = Task.Run(async () =>///●●●●●●●●●●●●●●●●●●●●20250215因為發現會較送出單頁新資料的慢送出，故取消多執行緒
                //{
                string tabWin;
                /*20230322 菩薩慈悲：請問C#中 for each 陳述句可以反向遍歷麼？ 感恩感恩　南無阿彌陀佛:
                 * Bing大菩薩：在 C# 中，`foreach` 语句是用来遍历集合中的每一个元素。它通常用于按顺序遍历集合，而不是反向遍历。如果您想要反向遍历一个集合，可以使用 `for` 循环⁵。您可以使用 `for` 循环的索引来直接访问列表中的元素，并以相反的顺序进行迭代。
                 * 來源: 與 Bing 的交談， 2023/3/22(1) c# - Possible to iterate backwards through a foreach? - Stack Overflow. https://stackoverflow.com/questions/1211608/possible-to-iterate-backwards-through-a-foreach 已存取 2023/3/22.
                 * (2) C# 用于每个逆序, C# for 循环逆序, 反向列表 C#, C# 反向迭代栈, C# for 循环倒退, C# 列表反向不起作用, C 锐利 for .... https://zditect.com/article/90996.html 已存取 2023/3/22.
                 * (3) Foreach如何实现反向遍历啊-CSDN社区. https://bbs.csdn.net/topics/390373229 已存取 2023/3/22.
                 * (4) C#-使用迭代器实现倒序遍历_c# foreach 倒序_dxm809的博客-CSDN博客. https://blog.csdn.net/dxm809/article/details/90552743 已存取 2023/3/22.
                 * (5) C# foreach循环. http://c.biancheng.net/csharp/foreach.html 已存取 2023/3/22.
                 * (6) Iteration statements -for, foreach, do, and while | Microsoft Learn. https://learn.microsoft.com/en-us/dotnet/csharp/language-reference/statements/iteration-statements 已存取 2023/3/22.
                 * (7) C# Foreach Reverse Loop. https://thedeveloperblog.com/foreach-reverse 已存取 2023/3/22.
                 */
                //因為多數情況皆是使用者在作用/使用中的的分頁為最後開啟的，故反向巡覽遍歷檢查，以省去不必要的查找（尤其在分頁很多的時候）
                //且多在目前簡單編輯（Quick edit）分頁後，故只找到目前簡單編輯（Quick edit）分頁為止，以加速效率
                //for (int i = tabWindowHandles.Count - 1; i > -1; i--)
                for (int i = br.driver.WindowHandles.Count - 1; i > -1; i--)
                //for (int i = 0; i < tabWindowHandles.Count; i++)
                {
                    //int retry = 0;
                    //reLoadWindowHandles:
                    //tabWin = tabWindowHandles[i];
                    tabWin = br.driver.WindowHandles[i];
                    //try
                    //{
                    //    currentWin = br.driver.CurrentWindowHandle;
                    //}
                    //catch (Exception ex)
                    //{
                    //    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.Message);
                    //    break;
                    //}
                    if (tabWin == currentWin)
                    {
                        ////重新開啟分頁，以取得在分頁集合中最後一個位置（如果可以的話）20240731
                        ////if (tabWindowHandles.Count > 1 && i < tabWindowHandles.Count - 1)
                        //reLoadWindowHandles:
                        //    int tabCount = br.driver.WindowHandles.Count;// DateTime dt = DateTime.Now;
                        //    //while (tabCount > 1 && i < tabCount - 1 && tabCount == br.driver.WindowHandles.Count) { if (DateTime.Now.Subtract(dt).TotalSeconds > 20) { playSound(soundLike.over, true); break; } }
                        //    //tabCount = br.driver.WindowHandles.Count;
                        //    if (tabCount > 1 && i < tabCount - 1)
                        //    {
                        //        ////寫在這裡就太快，必須寫在後面
                        //        //if (br.GetValidWindowHandles(br.driver).Count > 1 && waitUpdate == false && waitTabWindowHandle == ""
                        //        //        && ocrTextMode == false)
                        //        //{
                        //        //    retry++;//重新取得 tabWin = br.driver.WindowHandles[i]; 才比較會載入新的WindowHandles數量
                        //        //    if (retry == 1) { playSound(soundLike.over, true); goto reLoadWindowHandles; }
                        //        //    retry = 0;
                        //        //}

                        //        if (waitUpdate && waitTabWindowHandle != "") br.driver.SwitchTo().Window(waitTabWindowHandle); br.driver.Close();

                        //        //playSound(soundLike.warn, true);
                        //        br.driver.SwitchTo().Window(tabWin);
                        //        string url_Driver = br.driver.Url;
                        //        //if (br.GetValidWindowHandles(br.driver).Count > 2&& !(waitUpdate && waitTabWindowHandle != ""))
                        //        //莫名其妙，明明1個分類，都會平白成了2個！               
                        //        //if (br.GetValidWindowHandles(br.driver).Count > 1 && !(waitUpdate && waitTabWindowHandle != ""))
                        //        //這個 GetValidWindowHandles 方法完全沒有用，明明無中生有的分頁句柄也能切換過去而不出錯 20240802//恢復之前舊式的就好了！！
                        //        if (br.GetValidWindowHandles(br.driver).Count > 1 && waitUpdate == false && waitTabWindowHandle == ""
                        //                && ocrTextMode == false)
                        //        {
                        //            retry++;//重新取得 tabWin = br.driver.WindowHandles[i]; 才比較會載入新的WindowHandles數量（寫在這才能抓到新的視窗句柄集合）
                        //            if (retry == 1) { playSound(soundLike.over, true); goto reLoadWindowHandles; }
                        //            retry = 0;

                        //            br.driver.Close();
                        //            br.openNewTabWindow().Navigate().GoToUrl(url_Driver);
                        //            currentWin = br.GetCurrentWindowHandle(br.driver);
                        //        }
                        //        else
                        //            br.driver.SwitchTo().Window(br.GetCurrentWindowHandle(br.driver));
                        //}
                        break;
                    }
                    //}
                    //foreach (string tabWin in tabWindowHandles)
                    //{
                    int chkWindows = 0;
                    try
                    {
                        chkWindows = br.driver.SwitchTo().Window(tabWin).Url.IndexOf("&action=editchapter");
                    }
                    catch (Exception)
                    {
                        continue;
                    }

                    //if (br.driver.SwitchTo().Window(tabWin).Url.IndexOf("&action=editchapter") > -1)
                    if (chkWindows > -1)
                    {
                        waitUpdate = true; waitTabWindowHandle = tabWin;
                        OpenQA.Selenium.IWebElement commit = br.waitFindWebElementByName_ToBeClickable("commit", br.WebDriverWaitTimeSpan); //br.driver.FindElement(OpenQA.Selenium.By.Name("commit"));
                                                                                                                                            //OpenQA.Selenium.Support.UI.WebDriverWait waitcommit = new OpenQA.Selenium.Support.UI.WebDriverWait(br.driver, TimeSpan.FromSeconds(2));
                                                                                                                                            //waitcommit.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(commit));
                        string xInput = br.Textarea_data_Edit_textbox.GetAttribute("value"), urlEdit = driver.Url;
                    //await Task.Run(() =>
                    //{ //送出後也不必等待，也沒有其他須用到的元件，故可交給作業系統開個新線程去跑就好，但因為editchapter上傳儲存時常較Quit edit費時，故保險起見，還是在後加個Task.delay一下比較好
                    reCommit:
                        try
                        {
                            commit.Click();
                        }
                        catch (Exception ex)
                        {
                            switch (ex.HResult)
                            {
                                case -2146233088://原有窗體被關閉時："stale element reference: element is not attached to the page document
                                                 ////string urlOld = br.driver.SwitchTo().Window(tabWin).Url;
                                                 ////br.driver.Navigate().GoToUrl(br.driver.SwitchTo().Window(tabWin).Url);
                                                 //MessageBox.Show("原有窗體已被關閉，請自行先送出儲存，再回來按OK確定！","",
                                                 //    MessageBoxButtons.OK,MessageBoxIcon.Warning,MessageBoxDefaultButton.Button1,
                                                 //    MessageBoxOptions.DefaultDesktopOnly);
                                    break;
                                default:
                                    throw;
                            }
                        }

                        #region 送出後檢查是否是「Please confirm that you are human! 敬請輸入認證圖案」頁面 網址列：https://ctext.org/wiki.pl
                        if (br.IsConfirmHumanPage())
                        {
                            try
                            {
                                Clipboard.SetText(xInput);//複製到剪貼簿備用
                            }
                            catch (Exception)
                            {
                            }

                            //點選輸入框
                            OpenQA.Selenium.IWebElement iweConfirm = WaitFindWebElementBySelector_ToBeClickable("#content3 > form > table > tbody > tr:nth-child(2) > td:nth-child(2) > input[type=text]");
                            if (iweConfirm == null) driver.Navigate().Back();//因非同步，若已翻到下一頁
                            iweConfirm = WaitFindWebElementBySelector_ToBeClickable("#content3 > form > table > tbody > tr:nth-child(2) > td:nth-child(2) > input[type=text]");
                            if (iweConfirm == null)
                                Debugger.Break();
                            else
                                iweConfirm.Click();
                            if (DialogResult.Cancel ==
                                Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("Please confirm that you are human! 請輸入認證圖案"
                                + Environment.NewLine + Environment.NewLine + "請輸入完畢後再按「確定」！", string.Empty, false))
                            {
                                Debugger.Break();
                                TopMost = false;
                                driver.SwitchTo().Window(driver.CurrentWindowHandle);
                                Form1.InstanceForm1.EndUpdate();
                                return false;
                            }
                            driver.Navigate().Back();
                            while (driver.Url == "https://ctext.org/wiki.pl")
                            {
                                driver.Navigate().Back();
                            }
                            if (driver.Url != urlEdit)
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("網址並非 " + urlEdit + " 請檢查後再按下確定");
                            if (driver.Url == url)
                            {
                                br.SetIWebElementValueProperty(br.Textarea_data_Edit_textbox, xInput);
                                goto reCommit;//commit.Click();
                            }
                            else Debugger.Break();
                        }
                        #endregion //送出後檢查是否是「Please confirm that you are human! 敬請輸入認證圖案」頁面 網址列：https://ctext.org/wiki.pl

                        //});//只要有找到，都按下送出，反正若沒修改，也沒有任何影響202301112128

                    }
                }
                //});
                //確保所有editchapter都已上傳完畢
                //https://learn.microsoft.com/zh-tw/dotnet/api/system.threading.tasks.task.delay?view=netframework-4.8&f1url=%3FappId%3DDev16IDEF1%26l%3DZH-TW%26k%3Dk(System.Threading.Tasks.Task.Delay)%3Bk(TargetFrameworkMoniker-.NETFramework%2CVersion%253Dv4.8)%3Bk(DevLang-csharp)%26rd%3Dtrue
                //if (waitUpdate)
                //{
                //    //Task.Delay(4000).Wait(); //20230117 chatGPT大菩薩：Task 類別是用來創建新的執行緒來執行非同步作業的，而 Thread 類別則是用來管理當前執行緒的。Task 類別提供了許多用於創建和管理多個執行緒的方法，而 Thread 類別則提供了許多用於管理當前執行緒的方法，例如 Sleep() 方法和 Start() 方法。
                //    //Thread.Sleep(1200);
                //}
                //wait.Wait();//要有這行和async await 配合才行


                /* 20230118 creedit 與chatGPT菩薩討論：
                 * 是的 我想也應該是這樣的 我的程式改成如下 就成功了。 在其中 async 、 await  、 .wait() 三者 缺一不可 您看是嗎？（我已試著省略 wait.wait() 這行，則即使已有了 async await ，也不會等待而會接著下面的程式去做。只有加了 wait.wait() 這行 async await的標記才有作用
                 總結來說 它的邏輯應該是這樣的：
                  await 是在宣告 async 的 Task.Run 裡 等待這個Run 方法裡的另一個Task.Run()方法完成 故此第二個Task.Run() 前面會冠上  await ；而 第一個Task.Run方法回傳的名為 wait 的Task型別變數，使用它的 .Wait() 方法來等待第一個（即最外層的） Task.Run()完成 這樣 程式在執行時才能確實等待最外圈的 Task.完成 而最外圈的 Task 也確實等到了 內圈有加 await 關鍵字的 Task 都完成了，才算完成 是這樣吧
                 */
                //如果有編輯送出，待完成後關閉該分頁視窗
                if (waitUpdate && waitTabWindowHandle != string.Empty)
                {
                    dt = DateTime.Now;
                    if (autoPaste2QuickEdit)
                    {
                        //while (!wait.IsCompleted)
                        //{
                        //    if (DateTime.Now.Subtract(dt).TotalSeconds > 5)
                        //        if (DialogResult.Cancel == MessageBoxShowOKCancelExclamationDefaultDesktopOnly("逾時，是否繼續等候？"))
                        //            break;
                        //} //20250213●●●●●●●●●●●●●
                    }
                    Edited = true;
                    br.driver.SwitchTo().Window(waitTabWindowHandle); br.driver.Close();//關掉開啟的Edit頁面
                    if (br.DirectlyReplacingCharactersPageWindowHandle != string.Empty)
                        br.DirectlyReplacingCharactersPageWindowHandle = string.Empty;//重設；
                    br.driver.SwitchTo().Window(currentWin);
                    br.LastValidWindow = currentWin;
                }

                #endregion

                //檢查textbox3的值與現用網頁相同否
                //url = chkUrlIsTextBox3Text(tabWindowHandles);

                ////手動輸入時一般當不必自動清除框中文字
                //br.在Chrome瀏覽器的Quick_edit文字框中輸入文字(br.driver, Clipboard.GetText(), url);                
                //});
            }
            //else//不是在手動鍵入時
            //{//檢查textbox3的值與現用網頁相同否
            //string newCurrentWin = br.GetCurrentWindowHandle(br.driver);
            //if (newCurrentWin != null && br.IsWindowHandleValid(br.driver, newCurrentWin))
            //{

            //    if (currentWin != newCurrentWin)
            //    {
            //        if (br.GetValidWindowHandles(br.driver).Contains(currentWin))
            //            br.driver.SwitchTo().Window(currentWin);
            //        else
            //            br.driver.SwitchTo().Window(newCurrentWin);
            //    }
            //}
            //else
            //{
            // 當前視窗句柄無效
            //string currents = driver.CurrentWindowHandle;
            if (!br.IsDriverInvalid() && driver.CurrentWindowHandle != currentWin && !fastMode)
            {
                br.driver.SwitchTo().Window(currentWin);//切回圖文對照的頁面分頁視窗
                LastValidWindow = currentWin;

            }


            string beforeUpdated = null;
            if (waitTabWindowHandle != string.Empty || Edited)
                beforeUpdated = br.Div_generic_TextBoxFrame == null ? string.Empty : br.Div_generic_TextBoxFrame.GetAttribute("textContent");

            //如果存在「參考上下頁」控制項，則須刷新，否則會被前後頁的舊資料所干擾
            if (br.CheckAdjacentPages_Linkbox != null && Edited)
            { //要先記下可能有所編輯的前、後頁，否則一刷新就沒有了：
                if (br.CheckAdjacentPages_DataPrev != null)
                { //br.CheckAdjacentPages_Linkbox.Click();
                    string prePageText = string.Empty, nextPageText = string.Empty; //Clipboard.Clear(); 20240913作廢
                    try
                    {
                        if (br.CheckAdjacentPages_DataPrev.Text != string.Empty)
                        {
                            //br.CheckAdjacentPages_DataPrev.SendKeys(OpenQA.Selenium.Keys.Control + "ac");20240913作廢
                            //prePageText = Clipboard.GetText();
                            prePageText = br.CheckAdjacentPages_DataPrev.GetAttribute("value");
                        }
                        if (br.CheckAdjacentPages_DataNext.Text != string.Empty)
                        {
                            //br.CheckAdjacentPages_DataNext.SendKeys(OpenQA.Selenium.Keys.Control + "ac");20240913作廢
                            //nextPageText = Clipboard.GetText();
                            nextPageText = br.CheckAdjacentPages_DataNext.GetAttribute("value");
                        }
                    }
                    catch (Exception ex)
                    {
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                    }
                    br.driver.Navigate().Refresh();
                    br.CheckAdjacentPages_Linkbox.Click();
                    try
                    {
                        if (prePageText != string.Empty)
                        {
                            //Clipboard.SetText(prePageText);
                            //br.CheckAdjacentPages_DataPrev.SendKeys(OpenQA.Selenium.Keys.Control + "av");//20240913作廢
                            br.SetIWebElementValueProperty(br.CheckAdjacentPages_DataPrev, prePageText);
                        }
                        if (nextPageText != string.Empty)
                        {
                            //Clipboard.SetText(nextPageText);
                            //br.CheckAdjacentPages_DataNext.SendKeys(OpenQA.Selenium.Keys.Control + "av");//20240913作廢
                            br.SetIWebElementValueProperty(br.CheckAdjacentPages_DataNext, nextPageText);
                        }
                    }
                    catch (Exception ex)
                    {
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                    }
                }
                else//br.CheckAdjacentPages_DataPrev == null
                    br.driver.Navigate().Refresh();
                playSound(soundLike.over, true);
            }
            //無效，改寫到後面！
            //else
            //{
            //if (waitTabWindowHandle != string.Empty || Edited)
            //{
            //    driver.Navigate().Refresh();//如果有開啟Edit頁面分頁，則當切回圖文對照分頁時都一律刷新當前頁面，以免Edit頁面的舊資料覆蓋過圖文對照頁面的新資料 20250218

            //    playSound(soundLike.exam, true);//just for test
            //}
            //}
            //Task wait1 = Task.Run(() =>
            //{
            //    //chkUrlIsTextBox3Text(tabWindowHandles, textBox3.Text);
            chkUrlIsTextBox3Text(br.driver.WindowHandles, textBox3.Text);
            //});
            //wait1.Wait();
            ////}
            ////Task.WaitAny();//如上所設「wait.Wait();」「wait1.Wait();」，即不必此行了
            ////在連續輸入時能清除框中文字；手動輸入時一般當不必自動清除框中文字
            ////br.在Chrome瀏覽器的Quick_edit文字框中輸入文字(br.driver, clear == " " ? clear : Clipboard.GetText(), url);
            ////br.在Chrome瀏覽器的Quick_edit文字框中輸入文字(br.driver, clear == br.chkClearQuickedit_data_textboxTxtStr ? clear : Clipboard.GetText(), url);
            string formalX = clear == br.chkClearQuickedit_data_textboxTxtStr ? clear : br.TextPatst2Quick_editBox;
            CnText.FormalizeText(ref formalX);

            dt = DateTime.Now;
            if (waitTabWindowHandle != string.Empty)
            {
                while (driver.WindowHandles.Contains(waitTabWindowHandle))
                {
                    if (DateTime.Now.Subtract(dt).TotalSeconds > 5)
                        if (DialogResult.Cancel == MessageBoxShowOKCancelExclamationDefaultDesktopOnly("逾時，是否繼續等候？"))
                            break;
                }
            }


            if (autoPaste2QuickEdit && autoPastetoCtextQuitEditTextboxCancel) autoPastetoCtextQuitEditTextboxCancel = false;

            if (br.在Chrome瀏覽器的Quick_edit文字框中輸入文字(br.driver,
                formalX
                , url))
            {
                if (beforeUpdated != null)
                {
                    //檢查是否真有成功送出
                    if (br.Div_generic_TextBoxFrame?.GetAttribute("textContent") == beforeUpdated)
                    //|| driver.Url.Contains("&chapter="))
                    {
                        //br.CheckAdjacentPages_Linkbox != null && Edited

                        //Debugger.Break();
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("★★★★資料並未送出，請檢查!!!!!!!!");
                        if (fastMode) FastModeSwitcher();
                        return false;
                    }
                }
                //br.LastValidWindow = br.driver.CurrentWindowHandle;//在br.在Chrome瀏覽器的Quick_edit文字框中輸入文字()方法中已有
                return true;
            }
            else
            {
                if (beforeUpdated != null || AddTranscription_Linkbox != null)
                {
                    if (AddTranscription_Linkbox != null
                        || beforeUpdated == string.Empty || beforeUpdated == "\t"
                        || br.Div_generic_TextBoxFrame?.GetAttribute("textContent") == beforeUpdated
                        || driver.Url.Contains("&chapter="))
                    {
                        //Debugger.Break();
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("★★★★資料並未送出，請檢查!!!!!!!!");
                        if (fastMode) FastModeSwitcher();
                        return false;
                    }
                }
                return false;
            }

        }

        /// <summary>
        /// 檢查textbox3的Text值與現用網頁是否相同
        /// </summary>
        /// <param name="tabWindowHandles"></param>
        /// <param name="url"></param>
        /// <returns></returns>
        private string chkUrlIsTextBox3Text(ReadOnlyCollection<string> tabWindowHandles, string url)
        {
            if (url == "") return url;
            //再回到正在編輯的本頁，準備貼入

            if (br.GetDriverUrl != textBox3.Text)
            {
                bool found = false;
                if (tabWindowHandles.Count < br.driver.WindowHandles.Count) tabWindowHandles = br.driver.WindowHandles;//避免分頁視窗被關閉了。
                for (int i = tabWindowHandles.Count - 1; i > -1; i--)
                {
                    string tabWindowHandle = tabWindowHandles[i]; string taburl = string.Empty;
                    try
                    {
                        taburl = br.driver.SwitchTo().Window(tabWindowHandle).Url;
                    }
                    catch (Exception ex)
                    {
                        switch (ex.HResult)
                        {
                            case -2146233088://"no such window: target window already closed
                                continue;
                            default:
                                break;
                        }
                    }
                    //if (taburl == textBox3.Text || taburl.IndexOf(textBox3.Text.Replace("editor", "box")) > -1) { found = true; break; }
                    if (taburl == textBox3.Text || taburl.IndexOf(url.Replace("editor", "box")) > -1) { found = true; break; }
                }
                if (!found)
                    br.driver = br.openNewTabWindow();//網址由下面「在Chrome瀏覽器的Quick_edit文字框中輸入文字」那行給
            }
            //url = textBox3.Text;
            //textBox3.Text = url;
            //string urlLast= url = br.driver.Url;
            return url;
        }

        private void pasteToCtext()
        {//for .BrowserOPMode.appActivateByName
            appActivateByName();

            //if (ModifierKeys == Keys.None)//在textBox1_TextChanged事件中已處理按著Shift時的行為
            //{
            //    string currentForeTabUrl = br.ActiveTabURL_Ctext_Edit;
            //    appActivateByName();
            //    //如果視窗改變，非原所見之分頁，則回到按下 Ctrl + Shift + + 組合鍵時的分頁視窗
            //    if (currentForeTabUrl != br.ActiveTabURL_Ctext_Edit)
            //    {
            //        br.driver = br.driver ?? br.driverNew();
            //        br.GoToUrlandActivate(currentForeTabUrl);
            //    }
            //}

            //if (keyinText)
            //{
            //    hideToNICo();
            //}
            if (ModifierKeys == Keys.Shift && autoPaste2QuickEdit)//|| (autoPastetoQuickEdit && ModifierKeys == Keys.Control)) //|| ModifierKeys == Keys.Control
                                                                  //||autoPastetoQuickEdit)//
                                                                  //&& (textBox1.SelectionLength == predictEndofPageSelectedTextLen
                                                                  //&& textBox1.Text.Substring(textBox1.SelectionStart + textBox1.SelectionLength, 2) == Environment.NewLine))
            {//當啟用預估頁尾後，按下 Ctrl 或 Shift Alt 可以自動貼入 Quick Edit ，唯此處僅用 Ctrl 及 Shift 控制關閉前一頁所瀏覽之 Ctext 網頁                
                SendKeys.Send("^{F4}");//關閉前一頁                
            }
            Task.Delay(100).Wait();
            Task.WaitAll();
            if (Clipboard.GetText() != br.TextPatst2Quick_editBox) Clipboard.SetText(br.TextPatst2Quick_editBox);
            SendKeys.Send("^v{tab}~");
            //this.WindowState = FormWindowState.Minimized;
            //throw new NotImplementedException();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            if (button2.Text == "全部文")
            {
                button2.Text = "選取文";
                button2.BackColor = Color.Red;
            }
            else
            {
                button2.Text = "全部文";
                button2.BackColor = button2BackColorDefault;
            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            textBox1ReSize();
            //if (this.WindowState == FormWindowState.Minimized) hideToNICo();
        }

        void textBox1ReSize()
        {
            textBox1.Height = this.Height - textBox1SizeToForm.Height;// this.Height - textBox2.Height * 3 - textBox2.Top;
            textBox1.Width = this.Width - textBox1SizeToForm.Width;
        }

        const int CJK_Crtr_Len_Max = 2;//因為目前CJK最長為2字元

        /// <summary>
        /// 取代文字
        /// 用「7」為前綴以更新要取代成的字串
        /// </summary>
        /// <param name="replacedword">要被取代的字串</param>
        /// <param name="rplsword">用以取代的字串</param>
        private void replaceWord(string replacedword, string rplsword)
        {
            if (rplsword == "") return;
            bool editListMode = false;
            if (rplsword.StartsWith("7"))
            {//如果在此框輸入的字串前綴半形「@」符號，則會將被取代的字串其對應的用以取代之字串改成目前指定的這個（即在「@」後的字串）20230903蘇拉Saola颱風大菩薩往生後海葵Haikui颱風大菩薩光臨臺灣本島日。感恩感恩　讚歎讚歎　南無阿彌陀佛
                editListMode = true;
                rplsword = rplsword.Substring("7".Length);
            }
            if (textBox1.SelectionStart == textBox1.Text.Length) return;
            StringInfo selWord = new StringInfo(rplsword);
            string x = textBox1.Text;
            //string replacedword = textBox1.SelectedText;
            if (replacedword == "")//(!(selWord.LengthInTextElements > 1 || textBox1.SelectionLength == 0))
            {//無選取文字則以插入點後一字為被取代字                
                StringInfo replacedWord;
                if (textBox1.SelectionStart + CJK_Crtr_Len_Max > textBox1.Text.Length)
                    replacedWord = new StringInfo(
                            x.Substring(textBox1.SelectionStart, 1));
                else
                    replacedWord = new StringInfo(
                            x.Substring(textBox1.SelectionStart, CJK_Crtr_Len_Max));
                replacedword = replacedWord.SubstringByTextElements(0, 1);//取CJK一個單位字
            }
            if (replacedword == rplsword) return;
            int s = textBox1.SelectionStart; int l = 0;

            undoRecord();
            stopUndoRec = true; int cntr = 0, beforeScntr = 0;
            if (button2.Text == "選取文")
            {
                replacedword = textBox2.Text;
                if (replacedword == "") { stopUndoRec = false; return; }
                l = textBox1.SelectionLength;
                string xBefore = x.Substring(0, s), xAfter = x.Substring(s + l);
                x = textBox1.SelectedText;
                if (rplsword == "\"\"") rplsword = "";//要清除所選文字，則選取其字，然後在 textBox4 輸入兩個英文半形雙引號 「""」（即表空字串），則不會取代成「""」，而是清除之。
                                                      //textBox1.Text = xBefore + x.Replace(replacedword, rplsword) + xAfter;
                cntr = ReplaceCntr(ref x, replacedword, rplsword, s, ref beforeScntr);
                textBox1.Text = xBefore + x + xAfter;
            }
            else
            {
                l = selWord.LengthInTextElements;
                if (rplsword == "\"\"") rplsword = "";
                //textBox1.Text = x.Replace(replacedword, rplsword);
                cntr = ReplaceCntr(ref x, replacedword, rplsword, s, ref beforeScntr);
                textBox1.Text = x;

            }
            if (replacedword.Length != rplsword.Length)
            {
                s += beforeScntr * (rplsword.Length - replacedword.Length);
            }
            addReplaceWordDefault(replacedword, rplsword, editListMode);
            #region 自動將圓括弧置換成{{}}
            if (replacedword == "（" && rplsword == "{{") textBox1.Text = textBox1.Text.Replace("）", "}}");
            if (replacedword == "）" && rplsword == "}}") textBox1.Text = textBox1.Text.Replace("（", "{{");
            #endregion
            textBox1.SelectionStart = s; textBox1.SelectionLength = l;
            //restoreCaretPosition(textBox1, s, l == 0 ? 1 : l);//textBox1.ScrollToCaret();
            restoreCaretPosition(textBox1, s, rplsword.Length);//textBox1.ScrollToCaret();
                                                               //if (l != 0)
                                                               //{
                                                               //    if (new StringInfo(replacedword).LengthInTextElements == 1)
                                                               //    {
                                                               //        l = rplsword.Length;
                                                               //    }
                                                               //}
                                                               ////restoreCaretPosition(textBox1, s,  l);
            textBox1.Focus();

            if (!insertMode)
            {
                if (new StringInfo(textBox1.SelectedText).LengthInTextElements == 1)
                    textBox1.SelectionLength = 0;
            }

            undoRecord();

            stopUndoRec = false;
            playSound(soundLike.notify, true);//若有取代則播音效，否則有些字筆畫太似，不放大不知道有沒有一致，取代了沒 20240913
        }

        List<string> replaceWordList = new List<string>();
        List<string> replacedWordList = new List<string>();

        string getReplaceWordDefault(string replacedWord)
        {
            if (replacedWordList.Count == 0) return "";
            string replsWord = "";
            //for (int i = 0; i < replacedWordList.Count; i++)
            for (int i = replacedWordList.Count - 1; i > -1; i--)
            {
                if (replacedWord == replacedWordList[i])
                {
                    replsWord += replaceWordList[i];
                }
            }
            return replsWord;

        }
        /// <summary>
        /// 加入取代字串之資料
        /// </summary>
        /// <param name="replacedWord"></param>
        /// <param name="replaceWord"></param>
        /// <param name="editMode">如果是改訂已存在的資料，就是true</param>
        void addReplaceWordDefault(string replacedWord, string replaceWord, bool editMode = false)
        {
            if (replacedWordList.Contains(replacedWord))
            {
                int i = 0, count = replacedWordList.Count;
                if (editMode)
                {
                    if (replaceWordList[replacedWordList.IndexOf(replacedWord)] != replaceWord)
                        replaceWordList[replacedWordList.IndexOf(replacedWord)] = replaceWord;
                    return;
                }
                else
                {
                    while (i < count)
                    {

                        if (replacedWordList.IndexOf(replacedWord, i) == replaceWordList.IndexOf(replaceWord, i))
                        {
                            return;
                        }
                        i++;
                    }
                }

            }
            replacedWordList.Add(replacedWord);
            replaceWordList.Add(replaceWord);

        }

        int ReplaceCntr(ref string xDomain, string replacedword,
            string rplsword, int s, ref int beforeScntr)
        {
            //if (xDomain.IndexOf("�������") > -1)
            //    xDomain = xDomain.Replace("�������", "●●●●●●●");
            //if (replacedword.IndexOf("�������") > -1)
            //    replacedword = replacedword.Replace("�������", "●●●●●●●");
            int i = xDomain.IndexOf(replacedword, StringComparison.Ordinal), cntr = 0;
            while (i < xDomain.Length && i > -1)
            {
                xDomain = xDomain.Substring(0, i) + rplsword +
                    //xDomain.Substring(i + xDomain.Length);
                    xDomain.Substring(i + replacedword.Length);
                cntr++;
                if (i < s)
                {
                    beforeScntr++;
                    s += rplsword.Length - replacedword.Length;
                }
                i = xDomain.IndexOf(replacedword, i + rplsword.Length, StringComparison.Ordinal);
            }
            return cntr;
        }
        /// <summary>
        /// 取代字串
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox4_Leave(object sender, EventArgs e)
        {
            if (textBox4.Text == "")
            {
                textBox4Resize(); textBox1.Focus(); return;
            }
            if (textBox1.SelectionLength == 0 && insertMode == true)
            {
                textBox4Resize(); textBox1.Focus(); return;
            }
            //else if (textBox1.SelectionLength == 0 && insertMode == false)
            //{
            int s = textBox1.SelectionStart; string x = textBox4.Text;
            //    textBox1.Select(s
            //        , char.IsHighSurrogate(textBox1.Text.Substring(s, 1), 0) ? 2 : 1);
            //}
            if (textBox1.SelectedText != "")
            {
                if (char.IsHighSurrogate(textBox1.SelectedText, 0) && textBox1.SelectedText.Length < 3)
                    textBox1.Select(s, 2);
            }
            //取代前備份
            saveText();

            //實際執行取代文字
            if (ModifierKeys != Keys.Shift)
            {
                replaceWord(textBox1.SelectedText, x);// textBox4.Text);
                if (textBox4.Text != "")
                {
                    try
                    {
                        Clipboard.SetText(x);// textBox4.Text);
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            textBox4Resize();
            PauseEvents();
            textBox4.Text = "";
            ResumeEvents();
            textBox1.Focus();
            undoRecord();
            if (ModifierKeys == Keys.Shift)
            {
                textBox1.Select(textBox1.SelectionStart + textBox1.SelectionLength, 0);
                textBox1.SelectedText = x;
                StringInfo si = new StringInfo(textBox1.Text.Substring(s, textBox1.SelectionStart - s));
                textBox1.Select(s, 0);
                SendKeys.Send("%e");
            }
        }

        private void textBox4Resize()
        {
            textBox4.Location = textBox4Location;
            textBox4.Size = textBox4Size;
            textBox4.ScrollBars = ScrollBars.None;
        }

        /* 20250215 GitHub　Copilot大菩薩：是的，可以使用 KeyboardInfo 類別來實作檢查 Caps Lock 狀態並執行 FastModeSwitcher 方法的功能。以下是如何在 Form1 類中使用 KeyboardInfo 類別來實現這個功能的步驟：
            1.	在 Form1 類中新增一個方法來檢+查 Caps Lock 狀態並執行 FastModeSwitcher 方法。
            2.	在 Form1_Load 事件中註冊鍵盤掛鉤。
            以下是具體的程式碼：
        */
        /// <summary>
        /// 這段程式碼會在 Form1 載入時註冊一個全域鍵盤掛鉤，並在每次按下鍵盤按鍵時檢查 Caps Lock 狀態。如果 Caps Lock 燈亮，則執行 FastModeSwitcher 方法。
        ///這樣可以確保在 Caps Lock 燈亮時觸發 FastModeSwitcher 方法，而不需要按住 Caps Lock 鍵。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void HookManager_KeyDown(object sender, KeyEventArgs e)
        {
            //    // 檢查 Caps Lock 狀態
            //    if (KeyboardInfo.getKeyStateToggled(System.Windows.Input.Key.CapsLock))
            //    {
            //        if (ModifierKeys == Keys.Control)
            //            FastModeSwitcher();
            //    }
        }

        private bool IsCapsLockOn()
        {
            return (GetKeyState(VK_CAPITAL) & 0x0001) != 0;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            //20250215 creedit_with_GitHub　Copilot大菩薩：
            // 註冊鍵盤掛鉤
            HookManager.KeyDown += HookManager_KeyDown;

            //C# 如何取得使用者的螢幕解析度:https://blog.xuite.net/q10814/blog/48070595 https://www.delftstack.com/zh-tw/howto/csharp/screen-size-in-csharp/
            Size Size = SystemInformation.PrimaryMonitorSize;
            int Width = SystemInformation.PrimaryMonitorSize.Width;
            int Height = SystemInformation.PrimaryMonitorSize.Height;
            //MessageBox.Show("你的螢幕解析度是" + Size + "\n Width = " + Width + "\n Height = " + Height);
            //FormStartPosition 列舉:https://docs.microsoft.com/zh-tw/dotnet/api/system.windows.forms.formstartposition?view=netframework-4.7.2
            this.Location = new Point
                //(Width - this.Width, Height - this.Height - (int)(textBox1.Height / 3));
                (Width - this.Width, Height / 9);
            textBox1ReSize();
            //this.PointToScreen();
        }

        private void textBox4_Enter(object sender, EventArgs e)
        {
            //if (textBox1.Text == "") return;
            if (textBox4.Size == textBox4Size)
                textBox4SizeLarger();
            if (new StringInfo(textBox1.SelectedText).LengthInTextElements > 1)
            {
                try
                {
                    Clipboard.SetText(textBox1.SelectedText);
                }
                catch (Exception)
                {
                    playSound(soundLike.error, true);
                }
                textBox4.Text = textBox1.SelectedText; textBox4.DeselectAll();
            }
            string rplsdWord = textBox1.SelectedText, x = textBox1.Text;
            int s = textBox1.SelectionStart, l = s < textBox1.TextLength ? (char.IsHighSurrogate(x.Substring(s, 1), 0) ? 2 : 1) : 0;
            if (rplsdWord == "") //&& insertMode == false)
            {
                rplsdWord = x.Substring(s, l);
            }
            #region 預設取代字串
            if (rplsdWord != "")
            {
                string rplsWord = getReplaceWordDefault(rplsdWord);
                if (rplsWord != "")
                {
                    textBox4.Text = rplsWord;
                    if (rplsWord.IndexOf(Environment.NewLine) > -1) textBox4.Height = textBox4Size.Height * 3;
                }
            }
            #endregion 預設取代字串
            restoreCaretPosition(textBox1, selStart, selLength == 0 ? l : selLength);
        }

        private void textBox4SizeLarger()
        {
            textBox4.Location = new Point(button1.Location.X, textBox4Location.Y);
            int width = textBox2.Size.Width + textBox2.Size.Width + textBox3.Width + textBox4Size.Width;
            textBox4.Size = new Size(
                (width < (Width - width - 50)) ? (Width - 50) : width
                    , textBox4Size.Height);
            textBox4.ScrollBars = ScrollBars.Horizontal;
            if (textBox4.Font != textBox4FontDefault) textBox4.Font = textBox4FontDefault;
        }

        private void textBox4_KeyDown(object sender, KeyEventArgs e)
        {
            Keys m = ModifierKeys;
            if (e.KeyCode == Keys.F2)
            {
                e.Handled = true;
                keyDownF2(textBox4);
                return;
            }
            if (m == Keys.Alt && e.KeyCode == Keys.D1)
            {
                insertWords("·", textBox4);
            }

        }
        bool doNotLeaveTextBox2 = false;
        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.NumPad5 || e.KeyCode == Keys.Oemplus || e.KeyCode == Keys.Add || e.KeyCode == Keys.Subtract)
            {
                textBox1_KeyDown(sender, e);
                return;
            }

            #region 只按下單一鍵            
            if (ModifierKeys == Keys.None)
            {//只按下單一鍵
                if (e.KeyCode == Keys.F1 || e.KeyCode == Keys.Pause)
                {//按下 F1 鍵 或 Pause 鍵
                    e.Handled = true;
                    splitLineParabySeltext(e.KeyCode);
                    if (doNotLeaveTextBox2) textBox2.Focus();//方便快速分行分段
                    return;
                }
                if (e.KeyCode == Keys.F2)
                {//按下 F2 鍵
                    keyDownF2(textBox2);
                    return;
                }
                if (e.KeyCode == Keys.F3)
                {//按下 F3 鍵
                    KeyEventArgs ekey = new KeyEventArgs(Keys.F3);
                    textBox1_KeyDown(textBox1, ekey);
                }
            }//以上 只按下單一鍵
            #endregion
        }

        private void keyDownF2(TextBox textBox)
        {
            if (textBox.Text != "")
            {//F2 : 全選/取消全選框裡文字。若原有選取文字則取消選取至其尾端
                string x = textBox.SelectedText;
                if (x != "")
                    textBox.Select(textBox.SelectionStart + textBox.SelectionLength, 0);
                if (x == textBox.Text)
                    textBox.SelectionStart = textBox.Text.Length;
                if (x == "")
                    textBox.Select(0, textBox.TextLength);
                textBox.ScrollToCaret();
            }
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            //textBox2.Text = "";
        }

        private void textBox1_MouseDown(object sender, MouseEventArgs e)
        {
            var m = ModifierKeys;
            //mouseMiddleBtnDown(textBox1, e);
            mouseBtnDown(sender, e);
            //if ((m & Keys.Control) == Keys.Control && (m & Keys.Shift) == Keys.Shift)
            //{
            //    runWord("漢籍電子文獻資料庫文本整理_十三經注疏");
            //}
            //if ((m & Keys.Alt) == Keys.Alt)
            //{
            //    runWord("維基文庫四部叢刊本轉來");
            //}

            if (e.Button == MouseButtons.Left && (m & Keys.Alt) == Keys.Alt && (m & Keys.Control) == Keys.Control)
            {//Ctrl+ Alt + 滑鼠左鍵：將插入點後的分行分段清除
                int s = textBox1.SelectionStart; string xSl = textBox1.Text.Substring(s, 2);
                if (xSl != Environment.NewLine) s = textBox1.Text.IndexOf(Environment.NewLine, s);
                textBox1.Select(s, 2);
                undoRecord();
                textBox1.SelectedText = "";
                return;

            }

            //點一下加新分行
            if (ModifierKeys == Keys.Control && e.Button == MouseButtons.Left)
            {
                newLine();
                return;
            }

            if (ModifierKeys == Keys.Control && e.Button == MouseButtons.Right)
            { switchRichTextBox1(); return; }
            if (m == Keys.Alt && e.Button == MouseButtons.Left)
            {
                if (textBox1.SelectionLength == predictEndofPageSelectedTextLen
                    && textBox1.Text.Substring(textBox1.SelectionStart + predictEndofPageSelectedTextLen, 2) == Environment.NewLine)
                    //if (keyDownCtrlAdd(false)) if (textBox1.Text != "") { pauseEvents(); textBox1.Text = ""; resumeEvents(); }
                    keyDownCtrlAdd(false);
                else
                    BackupLastPageText(Clipboard.GetText(), true, true);
                return;
            }

        }

        /// <summary>
        /// 折行、分行、新行
        /// </summary>
        private void newLine()
        {
            if (textBox1.SelectionLength == 0)
            {
                //Point p = e.Location;
                //int s = textBox1.GetCharIndexFromPosition(p);
                //string x = textBox1.Text;
                undoRecord();
                caretPositionRecord();
                string clipChkSpace = Clipboard.GetText();  //如果剪貼簿裡是分段符號+全形空格者，蓋為縮排之文本，將自動再按下其空格數
                StringBuilder sbInsertBreak = new StringBuilder(); int i = 0;
                sbInsertBreak.Append(Environment.NewLine);
                if (clipChkSpace.StartsWith(Environment.NewLine) && clipChkSpace.EndsWith("　") && clipChkSpace.Replace("　", string.Empty) == Environment.NewLine)
                {
                    while (Environment.NewLine.Length + i + 1 <= clipChkSpace.Length && clipChkSpace.Substring(Environment.NewLine.Length + i, 1) == "　")
                    {
                        i++;
                        sbInsertBreak.Append("　");
                    }
                }
                else//如果剪貼簿裡不是分段符號+全形空格者，則以插入點所在行/段的縮排格數為準
                {
                    while (i < _leadingSpacesRegex.Match(GetLineText(textBox1.Text, textBox1.SelectionStart).Replace("􏿽", "　")).Value.Length)
                    {
                        i++;
                        sbInsertBreak.Append("　");
                    }
                }
                textBox1.SelectedText = sbInsertBreak.ToString();
                textBox1.Select(textBox1.SelectionStart + textBox1.SelectionLength, 0);
                //textBox1.Text = x.Substring(0, s) + Environment.NewLine + x.Substring(s, x.Length - s);
                //resumeLocationView(p, s);
            }
            //switchRichTextBox1();
        }

        private void resumeLocationView(Point p, int s)
        {//回到原來的插入點位置、定位、視界 20220205年初五立春後1日
         //caretPositionRecall();
         //restoreCaretPosition(textBox1, s, 0);
            p.Y += textBox1.GetPositionFromCharIndex(s).Y;
            textBox1.Select(textBox1.GetCharIndexFromPosition(p), 0);
            textBox1.ScrollToCaret();
            //s = textBox1.GetCharIndexFromPosition(p);
            textBox1.Select(s, 0);
            textBox1.ScrollToCaret();
            //textBox1.AutoScrollOffset = p;                    
            //ScrollableControl scrollableControl = new ScrollableControl
            //{
            //    AutoScroll = true
            //};
            //scrollableControl.ScrollControlIntoView(textBox1 as TextBox);   
            //textBox1.AutoScrollOffset = textBox1.GetPositionFromCharIndex(textBox1.SelectionStart);
        }

        private void switchRichTextBox1()
        {
            richTextBox1.Size = textBox1.Size;
            richTextBox1.Location = textBox1.Location;
            richTextBox1.Show();
        }

        private void richTextBox1_MouseDown(object sender, MouseEventArgs e)
        {
            if (ModifierKeys == Keys.Control && e.Button == MouseButtons.Right)
            {
                richTextBox1.Size = textBox1.Size;
                richTextBox1.Location = textBox1.Location;
                richTextBox1.Visible = false;
            }
        }

        private void textBox2_MouseDown(object sender, MouseEventArgs e)
        {
            Keys m = ModifierKeys;
            if ((m & Keys.Control) == Keys.Control)
            {
                if (e.Button == MouseButtons.Left)
                {
                    textBox2.Text = "";
                }
            }
        }

        private void textBox4_MouseDown(object sender, MouseEventArgs e)
        {
            Keys m = ModifierKeys;
            if ((m & Keys.Control) == Keys.Control)
            {
                if (e.Button == MouseButtons.Left)
                {
                    textBox4.Text = "";
                }
            }

        }

        /// <summary>
        /// textBox1內容選取範圍的參數 20250221 補充說明
        /// </summary>
        int selStart = 0; int selLength = 0;
        /// <summary>
        /// 自動連續輸入時，判斷頁尾所在位置之值(有翻頁就得重設！）
        /// 與 pageEndText10 參照互用
        /// </summary>
        int pageTextEndPosition = 0;

        /// <summary>
        /// 記下每頁最後10字元長的字以作判斷用，與 pageTextEndPosition 參照互用
        /// </summary>
        string pageEndText10 = "";
        /// <summary>
        /// 記下現在按下的按鍵
        /// </summary>
        Keys keycodeNow = new Keys();

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (!_eventsEnabled) return;
            Keys mk = ModifierKeys;
            if (textBox1.Text.IndexOf("") > -1)
            {//Ctrl+Shift+6會插入這個""符號
                int s = textBox1.SelectionStart, l = textBox1.SelectionLength;
                textBox1.Text = textBox1.Text.Replace("", "");
                restoreCaretPosition(textBox1, s - 1, l);
            }
            //if (!undoTextBoxing && (ModifierKeys != Keys.Control && keycodeNow != Keys.Z))
            if (!undoTextBoxing && !(ModifierKeys == Keys.Control && keycodeNow == Keys.Z)
                    && !(ModifierKeys == Keys.Control && keycodeNow == Keys.H))//GitHub　Copilot大菩薩幫我改的！感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主20250325
                undoRecord();
            undoTextValueChanged(selStart, selLength);
            if (textBox1.Text == "" && !pasteAllOverWrite)
            {
                /*if (!keyinTextMode)*///非手動輸入時
                                       //hideToNICo();//20240926取消此功能，以罕用故（要隱藏到系統列任務列可以用Esc鍵或滑鼠中鍵）
                                       //else
                if (keyinTextMode)
                {//在手動輸入模式下
                    if (mk != Keys.None)
                    {//可能按下Shift+Delete 剪下textBox1的內容時
                        if (mk == Keys.ShiftKey)//20240920 Copilot大菩薩：全域鍵盤掛鉤在 C# Windows.Forms 中的使用：
                            hideToNICo();//https://sl.bing.net/f5zjhC1h9DU
                                         //,通常是要準備貼上的，所以就要將目前在用的瀏覽器置前，確保它取得焦點，否則有時系統焦點會或交給工作列                        
                        if (browsrOPMode == BrowserOPMode.appActivateByName)
                        {
                            //string currentForeTabUrl = br.ActiveTabURL_Ctext_Edit;
                            appActivateByName();
                            //if (currentForeTabUrl != br.ActiveTabURL_Ctext_Edit)
                            //    br.GoToUrlandActivate(currentForeTabUrl);
                        }
                        else
                        {
                        retry:
                            try
                            {
                                br.driver = br.driver ?? br.DriverNew();
                                //chatGPT：在 C# 中使用 Selenium 控制 Chrome 瀏覽器時，可以使用以下方法切換到 Chrome 瀏覽器視窗：
                                if (br.driver == null) { Form1.browsrOPMode = BrowserOPMode.seleniumNew; br.DriverNew(); }

                                br.driver.SwitchTo().Window(br.driver.CurrentWindowHandle);

                                //以下按鍵判斷若仍出錯，則改用新增一個欄位作參考，記錄下在非按下 Ctrl + Shift + + 等鍵時造成的text改變
                                if (textBox3.Text.IndexOf("edit") > -1 &&
                                    (!KeyboardInfo.getKeyStateDown(System.Windows.Input.Key.LeftCtrl) && KeyboardInfo.getKeyStateNone(System.Windows.Input.Key.Add)) &&
                                    KeyboardInfo.getKeyStateToggled(System.Windows.Input.Key.Delete))//判斷Delete鍵是否被按下彈起
                                {//手動輸入時，當按下 Shift+Delete 當即時要準備貼上該頁，故如此操作，以備確定無誤後手動按下 submit 按鈕
                                    br.GoToCurrentUserActivateTab();//if (browsrOPMode != BrowserOPMode.appActivateByName) 前已判斷
                                    OpenQA.Selenium.IWebElement quick_edit_box = br.waitFindWebElementByName_ToBeClickable("data", br.WebDriverWaitTimeSpan);//br.driver.FindElement(OpenQA.Selenium.By.Name("data"));
                                                                                                                                                             //OpenQA.Selenium.Support.UI.WebDriverWait wait = new OpenQA.Selenium.Support.UI.WebDriverWait(br.driver, TimeSpan.FromSeconds(2));
                                                                                                                                                             //wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(quick_edit_box));
                                    quick_edit_box.Clear();
                                    quick_edit_box.Click();
                                    quick_edit_box.SendKeys(OpenQA.Selenium.Keys.LeftShift + OpenQA.Selenium.Keys.Insert);
                                    OpenQA.Selenium.IWebElement submit = br.waitFindWebElementById_ToBeClickable("savechangesbutton", br.WebDriverWaitTimeSpan);//br.driver.FindElement(OpenQA.Selenium.By.Id("savechangesbutton"));
                                                                                                                                                                //wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(submit));
                                                                                                                                                                //放在一個 Task 中去執行，並立即返回。                                     
                                    Task.Run(() =>
                                    {//送出按鈕按下後可以跑線程，其他要取得元件操作者，就不移另跑線程。20230111 現在終於破解bugs找到癥結所在了。感恩感恩　讚歎讚歎　南無阿彌陀佛 19:09
                                        submit.Click();
                                    });
                                }

                                else if (mk == Keys.Shift && mk != (Keys.Shift | Keys.Control))//(mk == (Keys.Shift|Keys.Delete))
                                {//如果是準備剪下貼上：           //20241008先作廢
                                 //playSound(soundLike.press);
                                 //br.SelectAllQuickedit_data_textboxContent();
                                }
                            }
                            catch (Exception ex1)
                            {
                                switch (ex1.HResult)
                                {
                                    case -2146233088:
                                        if (ex1.Message.IndexOf("no such window: target window already closed") > -1)
                                        {
                                            //br.GoToCurrentUserActivateTab(); 
                                            ResetLastValidWindow();
                                            goto retry;
                                        }
                                        break;
                                    default:
                                        try
                                        {
                                            br.driver = null;
                                            br.driver = br.DriverNew(); goto retry;
                                        }
                                        catch (Exception ex)
                                        {
                                            switch (ex.HResult)
                                            {
                                                case -2146233088://"no such window: target window already closed
                                                    if (ex.Message.IndexOf("no such window") > -1)
                                                        break;
                                                    break;
                                                default:
                                                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                                                    Debugger.Break();
                                                    break;
                                            }
                                        }
                                        break;
                                }

                            }
                        }
                    }
                }

            }
        }

        /// <summary>
        /// textBox1還原記錄儲存器。目前還原上限定為300個（原為50個），詳此行操作： if (undoTextBox1Text.Count > 300)
        /// </summary>
        List<string> undoTextBox1Text = new List<string>();
        private void undoTextValueChanged(int s, int l)
        {
            if (textBox1OriginalText != "" &&
                                textBox1.Text != textBox1OriginalText)
            {
                textBox1.Text = textBox1OriginalText;
                textBox1OriginalText = "";
                //textBox1.SelectionStart = s; textBox1.SelectionLength = l;
                restoreCaretPosition(textBox1, s, l == 0 ? 1 : l);
            }
        }

        string ClpTxtBefore = "";
        private void Form1_Activated(object sender, EventArgs e)

        {//此中斷點專為偵錯測試用 感恩感恩　南無阿彌陀佛 20230314

            //if (_currentPageNum == "115") Debugger.Break();

            //alt + Shift + f ： 將章節單位的頁面樹狀目錄收起或展開
            //OutlineTitlesCloseOpenFoldExpandSwitcher();

            //20250113
            if (this.Name != "Form1")
            {
                if (Name == "Form2" && textBox1.Text.Contains("校官") && textBox1.Text.StartsWith("*欽定四庫全書")
                        && !Form1.InstanceForm1.textBox1.Text.Contains("〖文淵|閣寶〗")
                        && br.Svg_image_PageImageFrame != null)
                {//如果是《四庫全書》原扉頁，則將其書圖點大，以便檢視輸入
                    if (int.Parse(br.Svg_image_PageImageFrame.GetAttribute("width")) <= 500)
                        br.Svg_image_PageImageFrame.Click();
                }
                if (Application.OpenForms[0].Controls["textBox3"].Text != string.Empty && textBox3.Text != Application.OpenForms[0].Controls["textBox3"].Text)
                {
                    //PauseEvents();
                    textBox3.Text = Application.OpenForms[0].Controls["textBox3"].Text;
                    //setCurrentPageNum(textBox3.Text);
                    //ResumeEvents();
                }
            }
            //else
            //    if (fastMode) ChromeSetFocus();

            #region forDebugTest權作測試偵錯用20230310            
            //br.SetQuickedit_data_textboxTxt(textBox1.Text);
            //string x = Clipboard.GetText();
            //x = CnText.RemoveNestedBrackets(x);
            //Console.WriteLine(x);//在「即時運算視窗」寫出訊息

            //CnText.ClearLettersAndDigits(ref x);
            //CnText.ClearLettersAndDigits_UseUnicodeCategory(ref x);
            //CnText.ClearOthers_ExceptUnicodeCharacters(ref x);
            //keyinNotepadPlusplus("","南無阿彌陀佛");
            #endregion

            if (!_eventsEnabled) return;

            //Keys modifierKey = ModifierKeys;
            ////直接針對目前的分頁開啟古籍酷OCR//20240328暫取消
            //if (modifierKey == Keys.Shift && keyinTextMode && !HiddenIcon && !PagePaste2GjcoolOCR_ing)
            //{
            //    copyQuickeditLinkWhenKeyinMode(modifierKey);
            //    return;
            //}

            //最上層顯示
            if (!this.TopMost && keyinTextMode && !ocrTextMode && !PagePaste2GjcoolOCR_ing) this.TopMost = true;
            //if (!this.TopMost && !PagePaste2GjcoolOCR_ing || ModifierKeys != Keys.Control) this.TopMost = true;

            //不全部貼上取代原文字
            if (keyinTextMode && !pasteAllOverWrite) pasteAllOverWrite = false;

            //當自動由《四部叢刊資料庫》貼入，不做以下處置
            if (autoPasteFromSBCKwhether) { autoPasteFromSBCK(autoPasteFromSBCKwhether); return; }

            //汲取剪貼簿內資料
            //Application.DoEvents(); 
            string clpTxt = "";//記錄剪貼簿內文字資料
            try
            {
                clpTxt = Clipboard.GetText();
            }
            catch (Exception)
            {//等候剪貼簿可用

                //Task.Delay(900).Wait();
                Task.WaitAll();
                Application.DoEvents();
                while (!isClipBoardAvailable_Text()) { }
                try
                {
                    clpTxt = Clipboard.GetText();
                }
                catch (Exception)
                {
                }
                //throw;
            }

            #region 鍵入模式（手動輸入）時的處置
            if (keyinTextMode || autoPaste2QuickEdit)//20250117修訂
            {
                #region 如果剪貼簿裡的內容是尾綴含「#editor」的網址內容的話 20250117修訂
                if (ClpTxtBefore != clpTxt && clpTxt.StartsWith("http") && clpTxt.EndsWith("#editor"))
                {
                    //new SoundPlayer(@"C:\Windows\Media\Windows Balloon.wav").Play();
                    System.Media.SystemSounds.Asterisk.Play();

                    //更新網址
                    textBox3.Text = clpTxt;
                    WindowHandles["currentPageNum"] = _currentPageNum;

                    //記下次的網址，作為與下次的比對
                    ClpTxtBefore = clpTxt;

                    //據程式進行模式架構分別處置
                    switch (browsrOPMode)
                    {
                        //於預設模式時鍵入
                        case BrowserOPMode.appActivateByName:
                            appActivateByName();//取得網址時順便貼上簡單修改模式下的文字
                            Task.WaitAll();
                            SendKeys.Send("{F6 3}");
                            Task.WaitAll();
                            SendKeys.Send("^a");
                            SendKeys.Send("^x");
                            break;
                        //在Selenium操控Chrome瀏覽器時鍵入
                        case BrowserOPMode.seleniumNew:
                            if (br.driver == null) br.driver = br.DriverNew();

                            //自動取得網址 textBox3.Text = clpTxt 網頁內 quick_edit_box 框內的文字內容
                            try
                            {
                                //網頁就定位
                                if (br.driver.Url != clpTxt)
                                {
                                    //br.driver.ExecuteScript("window.open();");
                                    //br.driver.SwitchTo().NewWindow(OpenQA.Selenium.WindowType.Tab);//取得網址時順便貼上簡單修改模式下的文字
                                    br.openNewTabWindow();
                                    br.GoToUrlandActivate(clpTxt, keyinTextMode);
                                }

                                //如果是要編輯而不瀏覽，使擷取其中 quick_edit_box 框內的文字內容，複製到剪貼簿
                                if (br.driver.Url.IndexOf("edit") > -1 || clpTxt.Contains("edit"))
                                {
                                    //OpenQA.Selenium.IWebElement quick_edit_box = br.driver.FindElement(OpenQA.Selenium.By.Name("data"));
                                    //Clipboard.SetText(quick_edit_box.Text);//本事件函式最後會清除剪貼簿，故無用！●●●●●●●●●●●●●●● 20250410
                                    string quickeditX = br.Quickedit_data_textboxTxt;
                                    if (textBox1.Text != quickeditX) textBox1.Text = quickeditX;
                                }
                            }
                            catch (Exception ex)
                            {
                                switch (ex.HResult)
                                {
                                    //重新定位網頁
                                    case -2146233088://"no such window: target window already closed\nfrom unknown error: web view not found\n  (Session info: chrome=109.0.5414.75)"
                                                     //br.driver.SwitchTo().Window(br.driver.WindowHandles.Last());
                                                     //br.driver.Navigate().GoToUrl(clpTxt);
                                        br.GoToUrlandActivate(clpTxt, keyinTextMode);
                                        break;
                                    default:
                                        Console.WriteLine(ex.HResult + ex.Message);
                                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                                        //throw;
                                        break;
                                }
                            }
                            break;
                        //尚未實作
                        case BrowserOPMode.seleniumGet:
                            //Task.Run(() => { if (br.driver == null) br.driver = br.driverNew(); });
                            break;
                        default:
                            break;
                    }

                    //確定剪貼簿是可用的
                    while (!Form1.isClipBoardAvailable_Text()) { }


                    #region 讀取剪貼簿裡的內容（即擷取自 quick_edit_box 框內的文字）__先取消此功能，改交由別處進行！！！20240211大年初二
                    //string nowClpTxt = Clipboard.GetText();
                    ////確認資料
                    //if (nowClpTxt != "" && nowClpTxt != ClpTxtBefore && nowClpTxt.IndexOf("http") == -1)
                    //{
                    //    undoRecord();
                    //    //設定內容
                    //    textBox1.Text = nowClpTxt;
                    //    ClpTxtBefore = nowClpTxt;//clpTxt;//記下這次內容以供下次比對
                    //                             //自動加上書名號
                    //                             ////只要剪貼簿裡的內容合於以下條件
                    //                             //if (ClpTxtBefore != clpTxt && textBox1.Text == "" && clpTxt.IndexOf("http") == -1 && clpTxt.IndexOf("<scanb") == -1)
                    //                             //{
                    //    textBox1.Text = CnText.BooksPunctuation(ref nowClpTxt, false);
                    //    //return;
                    //    //}
                    //    //插入點游標置於頁首
                    //    //if(keyinText)//已於巢外的if判定了
                    //    textBox1.Select(0, 0);
                    //}
                    #endregion

                    #region 表單最上層顯示且滑鼠鍵盤可用－－ 先取消這個，改由別處控制 20240902
                    //if (!Active && !PagePaste2GjcoolOCR_ing)
                    //////if (!Active && !PagePaste2GjcoolOCR_ing&& ModifierKeys!=Keys.Control)
                    //{
                    //    Debugger.Break();
                    //    //    PauseEvents();
                    //    //    AvailableInUseBothKeysMouse();
                    //    //    //表單最上層顯示
                    //    //    if (!this.TopMost) this.TopMost = true;
                    //    //    ResumeEvents();
                    //}
                    #endregion

                    Clipboard.Clear();
                    return;
                }
                #endregion//如果剪貼簿裡是網址內容的話

            }//以上處置鍵入模式（keyinText=true）
            #endregion

            #region 自動連續輸入模式的處置
            if (autoPaste2QuickEdit && textBox1.Enabled == false)
            {
                textBox1.Enabled = true;
                textBox1.Focus(); textBox1.Refresh();
            }
            if (textBox1.Focused)
            {
                //設置插入點游標
                if (insertMode) Caret_Shown(textBox1); else Caret_Shown_OverTypeMode(textBox1);

                //if (textBox1.TextLength > 0 && textBox1.SelectionLength == textBox1.TextLength && selLength < textBox1.SelectionLength && selLength < 30)
                if (textBox1.TextLength > 0 && textBox1.SelectionLength == textBox1.TextLength && selLength > 0 && selLength < textBox1.SelectionLength && selLength < 30)
                {
                    textBox1.Select(selStart, selLength);
                }

                //如果是在全自動模式下，且無按下控制鍵 Ctrl 等
                if (!keyinTextMode && (autoPaste2QuickEdit || (autoPaste2QuickEdit && ModifierKeys != Keys.None)))
                {
                    //20230115 非Selenium模式才執行，因為 Selenium模式 已在函式方法裡啟用遞迴（recursion），不必靠表單此Activated事件才能再次啟動了貼上機制了，真正達到全自動化的境地
                    //if (browsrOPMode == BrowserOPMode.appActivateByName && !autoPastetoCtextQuitEditTextboxCancel)
                    if (browsrOPMode == BrowserOPMode.appActivateByName)
                        autoPaste2CtextQuitEditTextbox(out DialogResult dialogResult);
                    if (textBox1.TextLength >= 100)//配合下面「if (textBox1.TextLength < 100)」還要執行
                        return;//20230113
                }
            }
            #endregion

            //對textBox2的設置（若在textBox1找不到內容時）
            if (textBox2.BackColor == Color.GreenYellow &&
                doNotLeaveTextBox2 && textBox2.Focused) textBox2.SelectAll();

            //函式內會作判斷要不要自動執行Word VBA相關的程序
            autoRunWordVBAMacro();

            //bool autoPasteFromSBCKwhether = false; this.autoPasteFromSBCKwhether = autoPasteFromSBCKwhether;            

            #region 在textBox1內容文字少於100時的檢查，以自行決定其他的操作，如《中國哲學書電子化計劃》清除頁前的分段符號、撤掉與書圖的對應_脫鉤,《國學大師》的《四庫全書》本文等
            if (textBox1.TextLength < 100)
            {
                //如果剪貼簿裡的文字內容長於99個字元，則執行相關的 Word VBA
                if (clpTxt.Length > 99)
                {
                    //根據剪貼簿裡的文本特徵來作動作
                    if (clpTxt.IndexOf("<scanbegin file=") > -1 && clpTxt.IndexOf(" page=") > -1)
                    {
                        ocrTextMode = false;
                        //若有按下Ctrl 或 Shift 則執行圖文脫鉤 Word VBA
                        if ((ModifierKeys & Keys.Control) == Keys.Control && (ModifierKeys & Keys.Alt) == Keys.Alt)
                        {
                            runWordMacro("中國哲學書電子化計劃.撤掉與書圖的對應_脫鉤");
                            return;
                        }
                        //if (ModifierKeys == (Keys.Control | Keys.Shift))
                        //{
                        //    runWordMacro("中國哲學書電子化計劃.撤掉與書圖的對應_脫鉤");
                        //    return;
                        //}
                        //若沒有按下Ctrl 或 Shift 則執行 Word VBA
                        runWordMacro("中國哲學書電子化計劃.清除頁前的分段符號");


                        //Application.DoEvents();//●●●●●●●●●●●●●●●●●●20250214
                        //Task.WaitAny();

                        ////Task.Delay(waitTimeforappActivateByName).Wait();
                        //Task.Delay(550).Wait();

                        //清除剪貼簿
                        try
                        {
                            Clipboard.Clear();
                        }
                        catch (Exception)
                        {
                            //Application.DoEvents();
                            //Task.WhenAny();
                            //確定剪貼簿是可用的
                            isClipBoardAvailable_Text();
                            Clipboard.Clear();
                            //throw;
                        }
                        return;

                    }

                    //對複製自《國學大師》的《四庫全書》文本的處置
                    else if (clpTxt.IndexOf("a]") > -1 || clpTxt.IndexOf("a] ") > -1 ||
                        (!clpTxt.Contains("感恩感恩　讚歎讚歎　南無阿彌陀佛") && clpTxt.IndexOf("P") > -1
                            && int.TryParse(clpTxt.Substring(clpTxt.IndexOf("P") + 1, 1), out _)))//P 乃「北京元引科技有限公司《元引科技引得數字人文資源平臺·中國歷代文獻》」的文本特徵
                    {
                        ocrTextMode = false;
                        runWordMacro("中國哲學書電子化計劃.國學大師_Kanripo_四庫全書本轉來");
                        undoRecord();
                        if (fastMode && !textBox1.Text.Contains("　*") && TitleLeadingSpacesCount.Count > 0)
                            //if (fastMode && TitleLeadingSpacesCount.Count > 0)
                            AutoMarkTitleParagraph();
                        undoRecord();
                        return;
                    }
                    //對複製自《Kanripo漢籍リポジトリ》GitHub存放庫（repository）《Kanseki Repository 漢籍リポジトリ》的文本的處置
                    else if (Clipboard.GetText().Contains(":KR"))
                    {
                        ocrTextMode = false;
                        runWordMacro("中國哲學書電子化計劃.Kanripo_GitHub轉來");
                        undoRecord();
                        if (fastMode && !textBox1.Text.Contains("　*") && TitleLeadingSpacesCount.Count > 0)
                            //if (fastMode && TitleLeadingSpacesCount.Count > 0)
                            AutoMarkTitleParagraph();
                        undoRecord();
                        return;
                    }
                }
                else if (clpTxt.Contains("action=editchapter"))
                {
                    formatContentInput_SKQS();
                }
            }
            #endregion

        }//完成 From1的 Activated事件處理程序

        /// <summary>
        /// 執行讀入文本的格式化整理與輸入首頁的動作
        /// </summary>
        private void formatContentInput_SKQS(string clpTxt = "")
        {
            if (clpTxt.IsNullOrEmpty()) clpTxt = Clipboard.GetText();
            bool modifyParaMark = prepareEditTextReForm(clpTxt);
            clpTxt = Clipboard.GetText();
            if (clpTxt.IndexOf("<scanbegin file=") > -1 && clpTxt.IndexOf(" page=") > -1 &&
                clpTxt.Contains(" page=\"1\" />●\t<") == false && modifyParaMark)// page="1" />●	<   為 WordVBA 中國哲學書電子化計劃.新頁面 巨集程序所產生的新頁面語法標記，含「\t」為其特徵
                runWordMacro("中國哲學書電子化計劃.清除頁前的分段符號");
            if (autoPaste2QuickEdit || fastMode)
                inputSKQSFrontPage();

        }

        /// <summary>
        /// 執行WordVBA
        /// </summary>
        private void autoRunWordVBAMacro()
        {
            string xClip = "";
            try
            {
                xClip = Clipboard.GetText() ?? "";
            }
            catch (Exception)
            {
                Task.WaitAll();
                try
                {
                    xClip = Clipboard.GetText() ?? "";
                }
                catch (Exception)
                {
                }
            }
            if ((xClip.IndexOf("MidleadingBot") > 0 || xClip.IndexOf("此頁面可能存在如下一些問題：") > -1 || xClip.IndexOf("Wmr-bot") > -1)
                    && textBox1.TextLength < 100)//xClip.Length > 500 )                
            {
                bool nextPageAuto = false;
                if (ModifierKeys == Keys.Control)//如果按下Ctrl則自動翻到下一頁
                    nextPageAuto = true;
                //處理《維基文庫》的每卷文本準備貼入
                runWordMacro("維基文庫四部叢刊本轉來");

                if (textBox1.Text.IndexOf("MidleadingBot") > -1)
                {
                    xClip = textBox1.Text;
                    xClip = xClip.Substring(0, xClip.IndexOf("MidleadingBot"));
                    while (xClip.Substring(xClip.Length - 2, 2) == Environment.NewLine)
                        xClip = xClip.Substring(0, xClip.Length - 2);
                    textBox1.Text = xClip;
                    try
                    {
                        Clipboard.Clear();
                    }
                    catch (Exception)
                    {
                    }
                }
                if (nextPageAuto && browsrOPMode == BrowserOPMode.appActivateByName)
                {//自動模式通常在最後一頁會停住，故自行翻下一頁（下一卷首）備用
                    Task.WaitAll();
                    nextPages(Keys.PageDown, false);
                    Task.WaitAll();
                }
            }

        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            if (insertMode) Caret_Shown(textBox1);
            else Caret_Shown_OverTypeMode(textBox1);

        }

        private void textBox2_Enter(object sender, EventArgs e)
        {
            //避免事件被終止
            if (!_eventsEnabled) _eventsEnabled = true;

            //避免還原記錄被終止
            if (stopUndoRec) stopUndoRec = false;


            //重設欄位變量，以免OCR快速鍵失效
            PagePaste2GjcoolOCR_ing = false;

            textBox2.BackColor = Color.GreenYellow;
        }

        int surrogate = 0;

        bool isKeyDownSurrogate(string x)
        {/*解決輸入CJK字元長度為2的字串問題 https://docs.microsoft.com/en-us/previous-versions/windows/desktop/indexsrv/surrogate-pairs
          * 
        https://stackoverflow.com/questions/50180815/is-string-replacestring-string-unicode-safe-in-regards-to-surrogate-pairs */
            //UnicodeCategory category = UnicodeCategory.Surrogate;//https://docs.microsoft.com/zh-tw/dotnet/api/system.globalization.unicodecategory?view=net-6.0
            char[] xChar = x.ToArray();
            foreach (char item in xChar)
            {
                //if (CharUnicodeInfo.GetUnicodeCategory(item) == category)
                if (Char.IsSurrogate(item))
                {
                    surrogate++;
                }
            }
            if (surrogate % 2 != 0) { surrogate = 0; return true; }
            else { surrogate = 0; return false; }
        }

        /// <summary>
        /// 允許觸發事件程序
        /// 20230321 creedit with chatGPT大菩薩： C#中有類似vba的 stop 述句嗎 感恩感恩　南無阿彌陀佛
        /// Unfortunately, C# does not have an EnableEvent method that can be used to temporarily pause events from being triggered. However, there are a few workarounds that can be used to achieve something similar:
        /// Use a boolean variable to indicate whether the event should be triggered or not.
        /// Remove and re-add the event handlers as needed.
        /// Make use of the Application.DoEvents() method to allow the event to be processed before continuing with the remainder of the code.
        /// Here is an example of using a boolean variable to temporarily disable an event handler:
        /// private bool _eventsEnabled = true;
        ///private void MyEventHandler(object sender, EventArgs e)
        ///{
        ///    if (!_eventsEnabled)
        ///        return;
        ///    // Handle event normally
        ///}
        ///By setting _eventsEnabled to false, the event handler will not be executed until it is resumed by setting _eventsEnabled back to true.
        ///Keep in mind that if you have multiple event handlers attached to the same event, this approach will disable all of them, since it is based on a boolean variable check.
        /// </summary>
        private bool _eventsEnabled = true;
        /// <summary>
        /// 取得與設定允許事件處理程序與否
        /// </summary>
        public bool EventsEnabled { get => _eventsEnabled; set => _eventsEnabled = value; }
        public int PreviousEditwikiID { get => previousEditwikiID; set => previousEditwikiID = value; }
        /// <summary>
        /// 是否為OCR連續自動輸入模式
        /// </summary>
        public bool OcrTextMode { get => ocrTextMode; set => ocrTextMode = value; }
        /// <summary>
        /// 取得每行的長度（正常的行/段長度（含幾個漢字中文字））
        /// </summary>
        public int NormalLineParaLength { get => normalLineParaLength; set => normalLineParaLength = value; }
        public bool FastMode { get => fastMode; }//set => fastMode = value; }
        /// <summary>
        /// 暫停事件處理程序
        /// </summary>
        internal void PauseEvents()
        {
            EventsEnabled = false;
        }
        /// <summary>
        /// 恢復事件處理程序
        /// </summary>
        internal void ResumeEvents()
        {
            EventsEnabled = true;
        }

        /// <summary>
        /// 20230321 chatGPT大菩薩慈悲：C#中如何判斷磁碟機存在不存在，不能用 Directory.Exists 方法判斷，是吧？感恩感恩　南無阿彌陀佛
        /// 您好！ 要判斷C#中磁碟機是否存在，可以使用DriveInfo類別。以下是使用DriveInfo類別的簡單範例：
        /// DriveInfo driveInfo = new DriveInfo("C");
        /// bool isDriveExist = (driveInfo != null && driveInfo.IsReady);
        /// 在上面的例子中，DriveInfo建構式的參數是磁碟機的名稱，例如 C。然後，使用IsReady屬性檢查磁碟是否已準備好。如果磁碟機不存在或未就緒，則DriveInfo將為null，並且IsReady屬性將返回false。 希望這可以幫助您！
        /// </summary>        
        /// <summary>
        /// 判斷磁碟機存在不存在
        /// </summary>
        internal static bool DriveExist(string driveName)
        {
            if (driveName.IsNullOrEmpty() || driveName.Length > 3) return false;
            DriveInfo driveInfo = new DriveInfo(driveName.Substring(0, 1));//("C"); 
            return driveInfo != null && driveInfo.IsReady;
        }



        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (!EventsEnabled) return;
            if (!textBox1.Enabled) textBox1.Enabled = true;
            string x = textBox2.Text;
            if (x == "") return;

            #region 設置章節單元chapter所在
            /* 在textBox2中輸入「tr:nth-child(2)」，以指定或重設目前的文字版(View)單元（章節 chapter） 20250305
                > 如 tr:nth-child(3)，即 tr:nth-child(2) 的下一章（節、單元 chapter） */
            if (x.Contains("tr:nth-child(") || x.StartsWith("cpt"))
            {
                if (x.StartsWith("tr:nth-child("))
                {
                    PauseEvents();
                    ChapterSelector = "#content > div:nth-child(6) > table > tbody > " + x + " > td:nth-child(1) > a";
                    textBox2.Text = ""; ResumeEvents(); return;
                }
                else if (x.StartsWith("#content > div:nth-child(6) > table > tbody > "))
                {
                    PauseEvents();
                    ChapterSelector = x;
                    textBox2.Text = ""; ResumeEvents(); return;
                }
                else if (x.StartsWith("cpt"))
                {
                    x = x.Substring("cpt".Length);
                    if (int.TryParse(x, out int chapterNum))
                    {
                        PauseEvents();
                        x = (++chapterNum).ToString();//第一列欄名也要算在內，故+1 （詳書首頁，如：https://ctext.org/library.pl?if=gb&res=1069）
                        ChapterSelector = "#content > div:nth-child(6) > table > tbody > tr:nth-child(" + x + ") > td:nth-child(1) > a";
                        textBox2.Text = ""; ResumeEvents(); return;
                    }
                    else
                        x = textBox2.Text;
                }
            }
            #endregion

            // 是否進行非正常行長（CheckAbnormalLinePara 方法）的檢查
            if (x == "alp")// 由textBox2輸入 "alp" 來切換設定
            {
                PauseEvents(); textBox2.Text = "";
                abnormalLineParaChecking = !abnormalLineParaChecking;
                ResumeEvents(); return;
            }

            //- 輸入「ifp」 設定各冊冊首/書首內容值（inputTextFrontPage）值
            if (x.StartsWith("ifp"))// 由textBox2輸入 "ifp" 來切換設定:如「ifp冊府元龜」即設定inputTextFrontPage值為"冊府元龜"，預設值為"{{{封面}}}\<p>"
            {
                PauseEvents(); textBox2.Text = "";
                if (x == "ifp")// 若只輸入「ifp」則還原為預設值
                    inputTextFrontPage = "{{{封面}}}<p>";
                else
                    inputTextFrontPage = x.Substring(3);
                ResumeEvents(); return;
            }

            //輸入「anv」 設定是否要自動複製下一卷/單位文本
            // autoNextVolumnContextMark值的切換。預設為true
            if (x.StartsWith("anv"))// 由textBox2輸入 "anv" 來切換設定
            {
                PauseEvents(); textBox2.Text = "";
                autoNextVolumnContextMark = !autoNextVolumnContextMark;
                ResumeEvents(); return;
            }

            #region 設置標題空格參數-即操作「TitleLeadingSpacesCount」欄位
            /* - 輸入「tlsc」(TitleLeadingSpacesCount)後可以在textBox1前端列出目前的標題階級及其空格數
                - 輸入「tlsc.rmv1」(rmv=Remove)，可以移除第1個項目，「tlsc.clr」(clr=Clear)可以清除全部；清除textBox3的內容亦可以清除所有項目，及重設所有書面特徵參數（如每頁幾行、每行幾字等）*/
            if (x == "tlsc")
            {

                // 列舉 TitleLeadingSpacesCount 的內容
                StringBuilder sb = new StringBuilder();
                foreach (int count in TitleLeadingSpacesCount)
                    sb.Append(count + Environment.NewLine); // 將項目加入
                PauseEvents(); textBox2.Text = "";
                textBox1.Text = sb.ToString() + Environment.NewLine + Environment.NewLine + textBox1.Text;
                ResumeEvents(); return;
            }
            else if (x == "tlsc.clr")
            {
                PauseEvents(); textBox2.Text = "";
                TitleLeadingSpacesCount.Clear();
                ResumeEvents(); return;
            }
            else if (x.StartsWith("tlsc.rmv"))
            {
                if (int.TryParse(x.Substring("tlsc.rmv".Length), out int i))
                {//在textBox2中輸入「tlsc.rmv1」(rmv=Remove)，可以移除第1個項目
                    if (i > 0)
                    {
                        i--;
                        PauseEvents(); textBox2.Text = "";
                        TitleLeadingSpacesCount.RemoveAt(i);
                        ResumeEvents(); return;
                    }
                }
            }
            else if (x == "tlsc.add")//- 輸入「tlsc.add」可以依textBox1被選取的行段資料來新增標題空格之項目。一行一筆，元素值為整數。
            {
                if (textBox1.SelectedText != string.Empty)
                {
                    PauseEvents(); textBox2.Text = "";
                    string[] elements = textBox1.SelectedText.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                    foreach (var item in elements)
                    {
                        if (int.TryParse(item, out int element))
                            TitleLeadingSpacesCount.Add(element);
                    }
                    ResumeEvents(); return;
                }
            }


            #endregion
            #region 輸入資料夾路徑可指定有效的Chrome瀏覽器的下載位置
            if (Directory.Exists(x))
            {
                PauseEvents();
                br.DownloadDirectory_Chrome = x; textBox2.Text = "";
                ResumeEvents(); return;
            }
            #endregion

            #region 輸入「x,y」（x、y 為整數以半形逗號間隔，如「835,711」），指定《古籍酷》首頁快速體驗OCR的複製按鈕位置 Copybutton_GjcoolFastExperience_Location的 X 與 Y值
            //20231128Bing大菩薩：檢查文字方塊中的內容是否符合「x,y」的格式
            string pattern = @"^\d+,\d+$";
            Regex rgx = new Regex(pattern);
            if (rgx.IsMatch(x))
            {
                string[] numbers = x.Split(',');
                br.Copybutton_GjcoolFastExperience_Location.X = Int32.Parse(numbers[0]);
                br.Copybutton_GjcoolFastExperience_Location.Y = Int32.Parse(numbers[1]);
                PauseEvents(); textBox2.Text = "";
                ResumeEvents(); return;

            }
            #endregion

            #region 輸入「nb,」可以切換 GXDS.SKQSnoteBlank 值以指定是否要檢查注文中因空白而誤標的情形
            if (x == "nb,")
            {
                GXDS.SKQSnoteBlank = !GXDS.SKQSnoteBlank;
                PauseEvents();
                textBox2.Text = "";
                ResumeEvents(); return;
            }
            #endregion

            #region 重設《易》學清單起始索引值
            if (x.StartsWith("lx") && x.Length > 2)
            {//輸入「lx9」，即重設《漢籍全文資料庫》檢索易學關鍵字清單之起始索引值為9 即 ListIndex_Hanchi_SearchingKeywordsYijing=9。 

                if (Int32.TryParse(x.Substring("lx".Length), out br.ListIndex_Hanchi_SearchingKeywordsYijing))
                    PauseEvents();
                textBox2.Text = "";
                ResumeEvents(); return;

            }
            #endregion

            switch (x)
            {
                #region 輸入「oT」（ocr first ture）設定直接貼入OCR結果先不管版面行款排版模式 輸入「oF」（ocr first false ）設定直接貼入OCR結果先不管版面行款排版模式 PasteOcrResultFisrtMode = false

                case "oT":
                    BatchProcessingGJcoolOCR = false; PasteOcrResultFisrtMode = true; ocrTextMode = true; PagePaste2GjcoolOCR_ing = false; _eventsEnabled = true;
                    br.OCR_wait_time_Top_Limit＿second = 60;
                    PauseEvents();
                    textBox2.Text = "";
                    ResumeEvents(); return;
                case "oF":
                    BatchProcessingGJcoolOCR = true; PasteOcrResultFisrtMode = false; ocrTextMode = false; PagePaste2GjcoolOCR_ing = false; _eventsEnabled = true;
                    br.OCR_wait_time_Top_Limit＿second = 15;
                    PauseEvents();
                    textBox2.Text = "";
                    ResumeEvents(); return;
                #endregion
                #region 《古籍酷》OCR批量處理。在textBox2中輸入bT以啟用，輸入bF以停用
                case "bT":
                    //BatchProcessingGJcoolOCR = true; PasteOcrResultFisrtMode = true; ocrTextMode = true; PagePaste2GjcoolOCR_ing = false; _eventsEnabled = true;
                    BatchProcessingGJcoolOCR = true; PagePaste2GjcoolOCR_ing = false; _eventsEnabled = true;
                    br.OCR_wait_time_Top_Limit＿second = 60;
                    PauseEvents();
                    textBox2.Text = "";
                    ResumeEvents(); return;
                case "bF":
                    //BatchProcessingGJcoolOCR = false; PasteOcrResultFisrtMode = false; ocrTextMode = false; PagePaste2GjcoolOCR_ing = false; _eventsEnabled = true;
                    BatchProcessingGJcoolOCR = false; PagePaste2GjcoolOCR_ing = false; _eventsEnabled = true;
                    br.OCR_wait_time_Top_Limit＿second = 15;
                    PauseEvents();
                    textBox2.Text = "";
                    ResumeEvents(); return;
                #endregion
                case "mt":
                    Form1.MuteProcessing = !Form1.MuteProcessing;
                    PauseEvents();
                    textBox2.Text = "";
                    ResumeEvents(); return;
                case "mf":
                    Form1.MuteProcessing = false;
                    PauseEvents();
                    textBox2.Text = "";
                    ResumeEvents(); return;
                case "fm"://輸入「fm」（form move）切換設定-自動移動表單位置以迴避圖文對照頁面的文本區，以便檢校是否已經編輯過 autoTestPositionAvoidance=true 20240501
                    if (autoTestPositionAvoidance) autoTestPositionAvoidance = false;
                    else autoTestPositionAvoidance = true;
                    PauseEvents();
                    textBox2.Text = "";
                    ResumeEvents(); return;
                case "lx"://輸入「lx」重設《漢籍全文資料庫》檢索易學關鍵字清單之索引值為0 即 ListIndex_Hanchi_SearchingKeywordsYijing=0。 
                    br.ListIndex_Hanchi_SearchingKeywordsYijing = 0;
                    PauseEvents();
                    textBox2.Text = "";
                    ResumeEvents(); return;
                /// 在textBox2中輸入開關切換要整頁貼上Quick edit [簡單修改模式]  並將下一頁直接送交去OCR的網站
                /// kd：《看典古籍》 （kandianguji)
                /// kdapi：《看典古籍》api
                /// df ：default 古籍酷
                case "kd"://《看典古籍》OCR網頁
                    PagePast2OCRsite = br.OCRSiteTitle.KanDianGuJi;
                    PauseEvents();
                    textBox2.Text = "";
                    ResumeEvents(); return;
                case "kapi"://《看典古籍》api
                    PagePast2OCRsite = br.OCRSiteTitle.KanDianGuJiAPI;
                    PauseEvents();
                    textBox2.Text = "";
                    ResumeEvents(); return;
                case "df"://default 古籍酷
                    PagePast2OCRsite = br.OCRSiteTitle.GJcool;
                    PauseEvents();
                    textBox2.Text = "";
                    ResumeEvents(); return;
                default:
                    break;
            }





            #region 輸入末綴為「0」的數字可以設定開啟Chrome頁面的等待毫秒時間
            if (x != "" && x.Length > 2 && int.TryParse(x, out int c))
            {
                if (x.Substring(x.Length - 1) == "0")
                {
                    if (Int32.Parse(x) % 10 == 0 && x.Length > 2)
                    {
                        if (Int32.TryParse(x, out int w))
                        {
                            waitTimeforappActivateByName = w;
                            PauseEvents();
                            textBox2.Text = "";
                            ResumeEvents();
                            return;
                        }
                    }
                }
            }
            #endregion

            #region 預設瀏覽器名稱設定
            //輸入「msedge」「chrome」「brave」「vivaldi」，可以設定預設瀏覽器名稱
            if (x == "msedge" || x == "chrome" || x == "brave" || x == "vivaldi")
            {
                defaultBrowserName = x;
                PauseEvents(); textBox2.Text = ""; ResumeEvents();
                return;
            }
            #endregion

            #region 軟件架構-瀏覽操作模式設定
            switch (textBox2.Text)
            {
                case "ap,":
                ap: browsrOPMode = BrowserOPMode.appActivateByName;
                    PauseEvents();
                    textBox2.Text = "";
                    ResumeEvents();
                    return;
                case "aa":
                    goto ap;
                case "sl,":
                sl: browsrOPMode = BrowserOPMode.seleniumNew;
                    PauseEvents(); textBox2.Text = string.Empty; ResumeEvents();
                    //第一次開啟Chrome瀏覽器，或前有未關閉的瀏覽器時
                    if (br.driver == null)
                        br.driver = br.DriverNew();//不用Task.Run()包裹也成了
                    else
                    {//如果Chrome瀏覽器都沒有開啟或被誤關的話20230109
                     //因為 br.driver != null 先清除chromedriver：
                        Process[] chromeInstances = Process.GetProcessesByName("chrome");
                        if (chromeInstances.Length == 0)
                        {
                            chromeInstances = Process.GetProcessesByName("chromedriver");
                            foreach (var chromeInstance in chromeInstances)
                            {
                                chromeInstance.Kill();
                            }
                            Task.WaitAll();
                            //清除完後創建新的執行個體實例
                            br.driver = null; br.DriverNew();
                        }
                    }
                    try
                    {
                        if (br.driver != null && br.driver.Url != textBox3.Text) br.GoToUrlandActivate(textBox3.Text, keyinTextMode);
                    }
                    catch (Exception ex)
                    {
                        switch (ex.HResult)
                        {
                            case -2146233088:
                                if (ex.Message.StartsWith("no such window: target window already closed"))//"no such window: target window already closed\nfrom unknown error: web view not found\n  (Session info: chrome=109.0.5414.75)"
                                    br.GoToUrlandActivate(textBox3.Text, keyinTextMode);
                                else
                                    //chromedriver被誤關了
                                    if (!br.ChromedriverLose(ex))
                                {
                                    Debugger.Break();
                                    Console.WriteLine(ex.HResult + ex.Message);
                                    MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                                }
                                break;
                            default:
                                Console.WriteLine(ex.HResult + ex.Message);
                                MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                                Debugger.Break();
                                break;
                        }

                    }
                    //PauseEvents();
                    //textBox2.Text = "";
                    //ResumeEvents();
                    return;
                case "br":
                    goto sl;
                case "bb":
                    goto sl;
                case "ss":
                    goto sl;
                case "sg,":
                    //還未實作
                    //browsrOPMode = BrowserOPMode.seleniumGet;
                    //if (br.driver == null)
                    //{
                    //    Task.Run(() =>
                    //    {
                    //        br.driver = br.driverNew();
                    //    });
                    //}
                    //textBox2.Text = "";
                    return;

                default:
                    break;
            }
            #endregion 軟件架構-瀏覽操作模式設定

            #region 設定小注不換行的長度限制

            if (x.Length > 5 && x.Substring(0, 5) == "note:")
            {
                if (Int32.TryParse(x.Substring(5), out int n))
                {
                    noteinLineLenLimit = (byte)(n > 255 ? 255 : n);
                    PauseEvents();
                    textBox2.Text = string.Empty;
                    ResumeEvents();
                    return;
                }
            }
            #endregion


            #region 輸入前2字元為字母之指令。如輸入tS或tE分別設定等待伺服器或網頁元件的時間上限（秒鐘）
            //- 輸入「tS」前綴，如「tS10」即10秒設定 Selenium 操控的 Chrome瀏覽器伺服器（ChromeDriverService）的等待秒數（即「new ChromeDriver()」的「TimeSpan」引數值）。預設為 8.5。因昨大年夜 Ctext.org 網頁載入速慢又不穩，因此設置，以防萬一 20230122癸卯年初一 感恩感恩　讚歎讚歎　南無阿彌陀佛
            //- 輸入「tE」前綴，如「tE5」即5秒，設定 Selenium 操控的 Chrome瀏覽器中網頁元件的的等待秒數（WebDriverWait。即「new WebDriverWait()」的「TimeSpan」引數值）。預設為 3。
            if (x.Length > 2)
            {
                if (double.TryParse(x.Substring(2), out double t))
                {
                    switch (x.Substring(0, 2))
                    {
                        case "tS":
                            br.ChromeDriverServiceTimeSpan = t;
                            PauseEvents();
                            textBox2.Clear();
                            ResumeEvents();
                            return;
                        case "tE":
                            br.WebDriverWaitTimeSpan = t;
                            PauseEvents();
                            textBox2.Clear();
                            ResumeEvents();
                            return;
                        case "ws"://輸入「ws」（wait second）以指定延長等待開啟舊檔對話方塊出現的時間（毫秒數），如「ws1000」即延長1秒
                            br.Extend_the_wait_time_for_the_Open_Old_File_dialog_box_to_appear_Millisecond = (int)t;
                            PauseEvents();
                            textBox2.Clear();
                            ResumeEvents();
                            return;
                        case "wO"://輸入「wO」（wait OCR）以指定等待OCR諸過程最久的時間（以秒數），如「wO60」即最久等到60秒（1分鐘）
                            br.OCR_wait_time_Top_Limit＿second = (int)t;
                            PauseEvents();
                            textBox2.Clear();
                            ResumeEvents();
                            return;

                    }
                }
            }


            #endregion

            if (x.Length >= 2)
            {
                #region 切換《古籍酷》帳號
                if (x == "gjk" || x == "gg" || x == "jj" || x == "kk" || x == "jk")
                {
                    PauseEvents();
                    textBox2.Text = ""; ResumeEvents(); bool topmost = TopMost; TopMost = false;
                    if (x == "gjk")//只手動告知系統《古籍酷》帳號已切換                        
                    {
                        br.OCR_GJcool_AccountChanged = true;
                        if (br.waitGJcoolPoint) br.waitGJcoolPoint = false;
                    }
                    else if (x == "kk")//只切換IP，不切換《古籍酷》帳戶
                    {
                        TopMost = false;
                        Task ts = Task.Run(() =>
                        {
                            br.IPSwitchOnly();


                        });
                        //br.IPStatusMessageShow();//在上一行 IPSwitchOnly 內已有
                        ts.Wait(7000);
                        //SystemSounds.Exclamation.Play();
                        AvailableInUseBothKeysMouse();
                        //if (Mdb.IPStatus(br.CurrentIP??br.GetPublicIpAddress("")).Item4) textBox2.Text = "kk";
                    }
                    else if (x == "jk")//不切換IP，不切換《古籍酷》帳戶，欲直接進入首頁快速體驗者
                    {
                        if (!br.waitGJcoolPoint) br.waitGJcoolPoint = true;
                        br.OCR_GJcool_AccountChanged = false;
                        //if (!Active) BringToFront(); 
                        //availableInUseBothKeysMouse();
                    }
                    else if (x == "jj")//只切換《古籍酷》帳號，不換IP
                        br.OCR_GJcool_AccountChanged_Switcher(false, true);
                    else
                    {//"gg" : 切換《古籍酷》帳號，且也換IP
                        if (TopMost) TopMost = false;
                        br.OCR_GJcool_AccountChanged_Switcher();
                    }
                    TopMost = topmost;
                    return;
                }
                #endregion
                if (x == "fc")
                {
                    PauseEvents(); textBox2.Text = ""; ResumeEvents();
                    formatCategory2Columns_GjcoolOCRResult();
                    textBox1.Focus(); bringBackMousePosFrmCenter();
                    return;
                }
            }

            if (button2.Text == "選取文") return;
            string x1 = textBox1.Text;
            if (x == "" || x1 == "") return;
            if (isKeyDownSurrogate(x)) return;//surrogate字在文字方塊輸入時會引發2次keyDown事件            
            var sa = findWord(x, x1);
            if (sa == null) return;
            int s = sa[0]; int nextS = sa[1];
            if (s > -1)
            {
                textBox1.Select(s, x.Length);
                textBox1.ScrollToCaret();
                if (nextS > -1) { textBox2.BackColor = Color.Yellow; doNotLeaveTextBox2 = false; return; }
            }
            else
            {
                textBox2.BackColor = Color.Red;
                doNotLeaveTextBox2 = false; return;
            }
            textBox2.BackColor = Color.GreenYellow;
            SystemSounds.Hand.Play();//文本唯一提示
            doNotLeaveTextBox2 = true;
            textBox2.SelectAll();

        }


        bool checkSurrogatePairsOK(char cr)
        {
            if (char.IsSurrogate(cr))
            {
                surrogate++;
                if (surrogate % 2 != 0)
                {
                    return false;
                }
                surrogate = 0;
                return true;
            }
            else return true;
        }
        string lastKeyPressElement = "";

        bool pasteAllOverWrite = false;
        private readonly float textBox1FontDefaultSize;

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //按下BackSpace鍵
            if (e.KeyChar == 8) return;
            //按下數字鍵盤的「+」執行 pagePaste2GjcoolOCR 方法時
            if ((PagePaste2GjcoolOCR_ing && e.KeyChar == 43)) { e.Handled = true; PagePaste2GjcoolOCR_ing = false; return; }

            //按下 Scroll Lock 將字數較少的行/段落尾末標上分行/段符號（「\<p\>」或「\。<p\>」
            //> -： 在非自動且手動輸入模式下，在 textBox1 單獨按下數字鍵盤的「-」，執行與按下 Scroll Lock 一樣的功能
            if (keyinTextMode && !autoPaste2QuickEdit && e.KeyChar == 45) { e.Handled = true; return; }


            //if (e.KeyChar==30)
            //{
            //    return;
            //}
            //if (ModifierKeys!=Keys.None)
            //{
            //    return;
            //}
            //https://social.msdn.microsoft.com/Forums/vstudio/en-US/5d021d76-36cd-43e6-b858-5a905c2e86d4/how-to-determine-if-in-insert-mode-or-overwrite-mode?forum=wpf
            //https://stackoverflow.com/questions/1428047/how-to-set-winforms-textbox-to-overwrite-mode/17962132#17962132
            //How can I place a TextBox in overwrite mode instead of insert mode:https://www.syncfusion.com/faq/windowsforms/textbox/how-can-i-place-a-textbox-in-overwrite-mode-instead-of-insert-mode
            if (e.KeyChar == Environment.NewLine.ToCharArray()[0])
            {
                return;
            }
            if (!checkSurrogatePairsOK(e.KeyChar))
            {
                lastKeyPressElement += e.KeyChar;
                return;
            }
            if (e.KeyChar == 96)
            {//`： 於插入點處起至「　」或「􏿽」前止之文字加上黑括號【】//Print/SysRq 為OS鎖定不能用
             //e.Handled = true; 加上黑括號(); return;
                e.Handled = true; preceded_followed_specify_symbols("【】"); return;
            }
            if (e.KeyChar == 127)
            {//`： 於插入點處起至「　」或「􏿽」前止之文字加上黑括號【】//Print/SysRq 為OS鎖定不能用
                e.Handled = true; ; return;//清除插入點之前的所有空格()
            }
            else
            {
                //https://zh.wikipedia.org/wiki/UTF-16
                if (char.IsLowSurrogate(e.KeyChar))
                {
                    lastKeyPressElement += e.KeyChar;
                    lastKeyPress.Add(lastKeyPressElement);
                    lastKeyPressElement = "";
                    if (lastKeyPress.Count > 2)
                    {
                        lastKeyPress.RemoveAt(0);
                    }
                }
                else
                {
                    lastKeyPressElement += e.KeyChar;
                    lastKeyPress.Add(lastKeyPressElement);
                    lastKeyPressElement = "";
                    if (lastKeyPress.Count > 2)
                    {
                        lastKeyPress.RemoveAt(0);
                    }
                }
            }


            #region 如果是取代輸入模式：標點符號不取代漢字，但可被取代；分段處不取代
            keyPressOverTyping(e);
            //以上取代模式處理
            #endregion
        }


        /* 20230405 癸卯清明 creedit with Bing大菩薩：
            Word 中的文字輸入有取代與插入兩種模式，請問「取代」的英文原文是？
            Word 中的文字輸入有取代與插入兩種模式，請問「取代」的英文原文是？
            在 Word 中，「取代」模式的英文原文是 “Overtype”。            
            在繁中版的 MS中是「取代」，在英文原版的英文是？            
            在英文版的 Word 中，「取代」模式的英文原文是 “Overtype”。
            相信您，您這次不會錯吧？（我是要寫C#程式函式命名用的，將其中輸入文字方塊時的取代模式獨立出來，程式碼如後所附） 非常感謝 感恩感恩　讚歎讚歎　inBing大菩薩 佛南無阿彌陀佛
            是的，我確定在英文版的 Word 中，「取代」模式的英文原文是 “Overtype”。希望這對您的程式編寫有所幫助。            
            那繁中版Word中的「插入」的英文版是？
            在繁中版的 Word 中，「插入」模式的英文原文是 “Insert”。            
            嗯 那我就用對了 程式碼裡有 Insertmode 就是了 感恩感恩　讚歎讚歎　南無阿彌陀佛            
            太好了！很高興能夠幫助您。如果您還有其他問題，請隨時告訴我。
            今年是癸卯年嗎
            是的，根據天干地支排列，2023年是癸卯年¹。
            來源: 與 Bing 的交談， 2023/4/5(1) 2023年：癸卯兔年不一般，双春闰月黑兔年，明年都有啥说法？. https://new.qq.com/rain/a/20221202A00FCZ00 已存取 2023/4/5.
            (2) 2023年农历表，二〇二三年天干地支日历表，农历日历表2023癸卯年兔年阳历农历通书. https://www.xurili.com/wannianli/2023/2023nongli.php 已存取 2023/4/5.
            (3) 2023年天干地支对照表 干支日历表 - 第一星座网. https://www.d1xz.net/fsml/bgml/art552786.aspx 已存取 2023/4/5.
            (4) 2023年の干支「癸卯（みずのとう）」にはどんな意味がある？. https://print.shop.post.japanpost.jp/nenga/feature/article_67 已存取 2023/4/5.
            (5) 2023年の干支は「癸•卯(みずのと•う)」干支の意味や由来、おすすめ干支グッズを紹介 – 株式会社アスタス. https://www.astas.co.jp/eto/ 已存取 2023/4/5.
            程式碼改好了（詳後） 。真的非常感謝您　感恩感恩　讚歎讚歎　Bing大菩薩　南無阿彌陀佛
            太好了！很高興能夠幫助您。如果您還有其他問題，請隨時告訴我。

            喔 這裡的發問是有字數限制的？ 我發現剛才貼給您的程式碼並不完整，後面被截掉，不見了。            
            是的，這裡的發問和回答都有字數限制。如果您的程式碼太長，可以考慮將其分成幾個部分發送。
            那如果日後要給您看很長的程式碼和長文章摘要話 怎麼辦呢？ 只要在分幾個部分時發送時，暫時停止您的回應就好了嗎？
            如果您想要給我看很長的程式碼或文章摘要，您可以將其分成幾個部分發送。我會在收到每一部分後回應您。如果您希望我暫時不回應，可以在發送信息時註明，我會等待您發送完整個信息後再回應您。
            喔 真的嘛 「如果您希望我暫時不回應，可以在發送信息時註明」 是在一開始告訴您還是每段發送時都要告訴您，請您暫不回應
            您可以在發送第一部分信息時註明希望我暫時不回應，直到您發送完整個信息。這樣我就會等待您發送完整個信息後再回應您。
            好的 也就是說在第一則信息請您暫別回應，在最後一則發完後再請您對以上各部分作總結整體的回應 ？
            是的，您可以在發送第一部分信息時註明希望我暫時不回應，直到您發送完整個信息。當您發送完整個信息後，我會對您發送的所有信息進行總結並回應您。
            好的 謝謝您 Bing大菩薩 感恩感恩　讚歎讚歎　Bing大菩薩　南無阿彌陀佛
            不客氣，很高興能夠幫助您。如果您還有其他問題，請隨時告訴我。
         */
        /// <summary>
        /// KeyPress事件中的取代模式處理。
        /// 1.如果是取代輸入模式：標點符號不取代漢字，但可被取代；
        /// 2.英文、上書名號、前篇名號（《〈）、缺字構字式符號（●＝＝）、小注文標識號（{}）、及分段處不取代。        
        /// </summary>
        /// <param name="e"></param>
        private void keyPressOverTyping(KeyPressEventArgs e)
        {

            /* 20230309 chatGPT大菩薩晚安吉祥：
             * public static string punctuationsNum = ".,;?@'\"。，；！？、－-—…:：《·》〈‧〉「」『』〖〗【】（）()[]〔〕［］0123456789";
             * 在 C# 我想剔除以上字中的 「《」和「〈」要怎麼寫才最簡潔有效呢？正則表達式的寫法會是最好的選擇嗎？ 感恩感恩　南無阿彌陀佛
             * 感恩合十，南无阿弥陀佛。
             * 在 C# 中，您可以使用正则表达式来删除给定字符串中的特定字符。以下是删除 punctuationsNum 字符串中的 "《" 和 "〈" 字符的示例代码：……
             * 在这里，我们使用 Regex.Replace 方法将匹配正则表达式模式 [《〈] 的所有字符替换为空字符串。此模式匹配任何包含 "《" 或 "〈" 的字符。
             * */
            //if (e.KeyChar == " ".ToCharArray()[0]) return;//半形空格可被輸入、被取代，而不能取代別人
            //if (e.KeyChar == "/".ToCharArray()[0]) return;//半形空格可被輸入、被取代，而不能取代別人
            //20240731 Copilot大菩薩 ： 簡化程式碼：您可以簡化這兩行檢查的程式碼如下：這樣可以達到相同的效果，並使程式碼更加簡潔明瞭。
            if (e.KeyChar == ' ' || e.KeyChar == '/') return;

            string regexPattern = "[《〈」】〗]", omitSymbols = @"＝{}□■<>*〇◯○⿰⿱」』|\\" + Environment.NewLine;//輸入缺字構字式●＝＝、及注文標記符{{}}、及標題星號*時不取代
            checkkeyPressOverTyping_oscarsun72note_Inserting_switch2insertMode(e.KeyChar, regexPattern + omitSymbols);
            string w;//, punctuationsNumWithout前書名號與前篇名號 = Regex.Replace(Form1.punctuationsNum, regexPattern, ""); 
            if (!insertMode
                && textBox1.SelectionStart < textBox1.TextLength
                //現在鍵入位置的後一個字不能是
                && ((regexPattern + omitSymbols).IndexOf(textBox1.Text.Substring(textBox1.SelectionStart, 1)) == -1 ||
                    textBox1.Text.Substring(textBox1.SelectionStart, 1) == "�")
                //&& (regexPattern.Replace("[",string.Empty).Replace("]",string.Empty) + omitSymbols).IndexOf(textBox1.Text.Substring(textBox1.SelectionStart, 1)) == -1
                //&& omitSymbols.IndexOf(e.KeyChar.ToString()) == -1
                && Regex.IsMatch(e.KeyChar.ToString(), "[^a-zA-Z" + omitSymbols + "]"))//YouChat菩薩
            {//https://stackoverflow.com/questions/1428047/how-to-set-winforms-textbox-to-overwrite-mode/70502655#70502655
                if (textBox1.Text.Length != textBox1.MaxLength && textBox1.SelectedText == ""
                    && textBox1.Text != "" && textBox1.SelectionStart != textBox1.Text.Length)
                {
                    //string x = textBox1.Text; int s = textBox1.SelectionStart;
                    //    string xNext = x.Substring(s);
                    //    StringInfo xInfo = new StringInfo(xNext);                    
                    textBox1.SelectionLength = 1;//對於已經輸入完成的 surrogate C#應該會正確判斷其字長度；實際測試非然也
                                                 //對標點符號punctuations所佔字位不取代
                    w = textBox1.SelectedText;
                    //標點符號不取代漢字，但可被取代
                    if ((PunctuationsNum + "●�\"").IndexOf(e.KeyChar) > -1 &&
                        (PunctuationsNum + "●�\"").IndexOf(textBox1.Text.Substring(textBox1.SelectionStart, 1)) == -1)
                        textBox1.SelectionLength = 0;
                    else if (char.IsSurrogate(w.ToCharArray()[0])) textBox1.SelectionLength = 2;
                }
            }

            //if (ModifierKeys == Keys.None)
            //{
            //    undoRecord();
            //}
            if (keyinTextMode && textBox1.TextLength > 0 && textBox1.SelectionLength == textBox1.TextLength)
            {
                if (!insertMode
                    && textBox1.SelectionStart < textBox1.TextLength && selStart < textBox1.TextLength
                    //現在鍵入位置的後一個字不能是
                    //&& (regexPattern + omitSymbols + Environment.NewLine).IndexOf(textBox1.Text.Substring(textBox1.SelectionStart, 1)) == -1
                    && (regexPattern + omitSymbols).IndexOf(textBox1.Text.Substring(textBox1.SelectionStart, 1)) == -1
                    //&& omitSymbols.IndexOf(e.KeyChar.ToString()) == -1
                    && Regex.IsMatch(e.KeyChar.ToString(), "[^a-zA-Z" + omitSymbols + "]"))//YouChat菩薩
                {
                    w = textBox1.Text.Substring(selStart, 1);//對標點符號punctuations所佔字位不取代
                    if (selStart + 1 > textBox1.TextLength ||
                        ((PunctuationsNum + "●�\"").IndexOf(e.KeyChar) > -1 &&
                        //標點符號不取代漢字，但可被取代
                        (PunctuationsNum + "●�\"").IndexOf(textBox1.Text.Substring(textBox1.SelectionStart, 1)) == -1))
                        textBox1.Select(selStart, 0);
                    else
                    {
                        textBox1.Select(selStart, char.IsHighSurrogate(w.ToCharArray()[0]) ? 2 : 1);
                    }
                }
                else
                    textBox1.Select(selStart, selLength);
                //textBox1.DeselectAll();
            }
        }

        /// <summary>
        /// 檢查是否正在手動輸入{{{孫守真按：}}}{{{佛弟子文獻學者孫守真任真甫按：}}}之類的文字。如果是，則改為插入輸入模式而非取代輸入模式。20230902
        /// 取代模式時才處理
        /// {{{ }}}	Specifies that the characters between the {{{ and }}} are marginal notes or textual remarks not occurring in the main body of the text. https://ctext.org/instructions/wiki-formatting
        /// </summary>
        /// <param name="c">正在輸入的字元</param>
        void checkkeyPressOverTyping_oscarsun72note_Inserting_switch2insertMode(char c, string omitSymbols = "")
        {

            int s = textBox1.SelectionStart; string x = textBox1.Text;
            if (insertMode || c != '{' || s < 2 || s == x.Length) return;
            //如果插入點在分段符號前（輸入）亦略過
            //if (s + 1 < x.Length) { if (x.Substring(s, 1) == Environment.NewLine.Substring(0, 1)) return; }
            //只要後面接著omitSymbols所包含的就略過。
            if (s + 1 < x.Length) { if (omitSymbols.Contains(x.Substring(s, 1))) return; }

            if (x.Substring(s - 2, 2) == "{{" && c == '{')
            //if (textBox1.Text.Substring(textBox1.SelectionStart - 2, 2) == "{{" && c == '{')
            {//if (x.Substring(s - 2) == "{{" && c == "{".ToCharArray()[0])
             //insertMode = true;
                InsertModeSwitcher();
                playSound(soundLike.exam);
            }
        }

        /// <summary>
        /// 記下更動前的文本以利還原
        /// 若textBox1是空字串、無內容則不記錄。
        /// <param name="undoText">若另有指定要備份的還原內容，則指定此引數。</param>
        /// </summary>
        private void undoRecord(string undoText = "")
        {
            if (stopUndoRec) return; if (textBox1.TextLength == 0) return;
            selStart = textBox1.SelectionStart; selLength = textBox1.SelectionLength;

            if (undoText == string.Empty)
            {
                if (undoTextBox1Text.Count == 0 || textBox1.Text != undoTextBox1Text.Last())
                    undoTextBox1Text.Add(textBox1.Text);
            }

            else
            {
                if (undoTextBox1Text.Count == 0 || undoText != undoTextBox1Text.Last())
                    undoTextBox1Text.Add(undoText);
            }

            if (undoTimes != 0) undoTimes = 0;
            if (undoTextBox1Text.Count > 300)//還原上限定為300個(原為50個）
            {
                undoTextBox1Text.RemoveAt(0);
            }
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            selStart = textBox1.SelectionStart; selLength = textBox1.SelectionLength;
        }

        private void textBox1_MouseUp(object sender, MouseEventArgs e)
        {
            caretPositionRecord();
        }

        private void Form1_MouseDown(object sender, MouseEventArgs e)
        {

            mouseBtnDown(sender, e);
        }

        DateTime nextPageStartTime;
        void mouseBtnDown(object sender, MouseEventArgs e)
        {
            #region ModifierKeys == Keys.None
            if (ModifierKeys == Keys.None)
            {
                TimeSpan timeDifference;
                switch (e.Button)
                {
                    case MouseButtons.Left:
                        //if (sender.GetType().Name == "TextBox")
                        //{//拖曳拖放移動文字
                        //    TextBox tb = sender as TextBox;
                        //    if (tb.SelectedText != "")
                        //    {//https://docs.microsoft.com/zh-tw/dotnet/desktop/winforms/advanced/walkthrough-performing-a-drag-and-drop-operation-in-windows-forms?view=netframeworkdesktop-4.8
                        //        tb.DoDragDrop(tb.SelectedText, DragDropEffects.Copy |
                        //        DragDropEffects.Move);
                        //    }
                        //}
                        break;
                    case MouseButtons.None:
                        break;
                    case MouseButtons.Right:
                        break;
                    //隱藏到任務列（系統列）中
                    case MouseButtons.Middle:
                        if (textBox1.SelectionLength == predictEndofPageSelectedTextLen
                                && textBox1.SelectionStart + predictEndofPageSelectedTextLen + 2 <= textBox1.TextLength
                                && textBox1.Text.Substring(textBox1.SelectionStart + predictEndofPageSelectedTextLen, 2) == Environment.NewLine)
                            //if (keyDownCtrlAdd(false)) if (textBox1.Text != "") { pauseEvents(); textBox1.Text = ""; resumeEvents(); }
                            keyDownCtrlAdd(false);
                        else
                            //預設為最上層顯示，則按下Esc鍵或滑鼠中鍵會隱藏到任務列（系統列）中；滑鼠在其 ico 圖示上滑過即恢復
                            hideToNICo();
                        break;

                    //須避免textbox1 與Form的同一事件衝突（呼叫 nextPage方法太頻繁）變成連翻上下頁，以致 chromedriver不及反應而出錯當掉20230111
                    case MouseButtons.XButton1:
                        if (browsrOPMode != BrowserOPMode.appActivateByName)
                        {//過於頻繁會造成chromedriver反應不及而當掉。終於抓到 bugs了！202301110632
                            timeDifference = DateTime.Now.Subtract(nextPageStartTime);
                            if (timeDifference.TotalSeconds < 0.3)
                                return;
                            nextPageStartTime = DateTime.Now;
                        }
                        nextPages(Keys.PageUp, false);
                        br.driver.SwitchTo().Window(driver.CurrentWindowHandle);
                        //上一頁
                        //if (autoPaste2QuickEdit || keyinTextMode) AvailableInUseBothKeysMouse();
                        AvailableInUseBothKeysMouse();
                        #region 檢查textBox1的Text值                        
                        if (keyinTextMode && browsrOPMode == BrowserOPMode.seleniumNew)
                        {
                            int cntr = 0;
                            while (textBox1.Text != br.Quickedit_data_textboxTxt)
                            {
                                playSound(soundLike.info);
                                textBox1.Text = br.Quickedit_data_textboxTxt;
                                if (cntr > 2) Debugger.Break();
                                cntr++;
                            }
                        }
                        #endregion
                        //keyDownCtrlAdd(false);
                        break;
                    case MouseButtons.XButton2:
                        bool rect = false;
                        if (browsrOPMode != BrowserOPMode.appActivateByName)
                        {//過於頻繁會造成chromedriver反應不及而當掉
                            timeDifference = DateTime.Now.Subtract(nextPageStartTime);
                            if (timeDifference.TotalSeconds < 0.3)
                                return;
                            nextPageStartTime = DateTime.Now;
                            if (br.WaitFindWebElementBySelector_ToBeClickable("#canvas > svg > rect") != null)
                            {
                                rect = true;
                                br.Input_picture(); //圖像的輔助輸入
                            }
                        }
                        //keyDownCtrlAdd(true);
                        //下一頁
                        nextPages(Keys.PageDown, false);
                        br.driver.SwitchTo().Window(driver.CurrentWindowHandle);
                        //if (autoPaste2QuickEdit || keyinTextMode) AvailableInUseBothKeysMouse();
                        if (!rect) AvailableInUseBothKeysMouse();
                        break;
                    default:
                        break;
                }
                //if (e.Button == MouseButtons.Middle)
                //{//預設為最上層顯示，則按下Esc鍵或滑鼠中鍵會隱藏到任務列（系統列）中；滑鼠在其 ico 圖示上滑過即恢復
                //    hideToNICo();
                //}

            }
            #endregion
            else
            {
                switch (e.Button)
                {
                    case MouseButtons.Left:
                        break;
                    case MouseButtons.None:
                        break;
                    case MouseButtons.Right:
                        break;
                    case MouseButtons.Middle:
                        //if (new StringInfo(textBox1.SelectedText).LengthInTextElements == predictEndofPageSelectedTextLen)
                        if (textBox1.SelectionLength == predictEndofPageSelectedTextLen
                                && textBox1.Text.Substring(textBox1.SelectionStart + predictEndofPageSelectedTextLen, 2) == Environment.NewLine)
                            //if (keyDownCtrlAdd(false)) if (textBox1.Text != "") { pauseEvents(); textBox1.Text = ""; resumeEvents(); }
                            keyDownCtrlAdd(false);
                        break;
                    case MouseButtons.XButton1://上一頁按鈕
                        break;
                    case MouseButtons.XButton2://下一頁按鈕
                        if (browsrOPMode != BrowserOPMode.appActivateByName)
                        {//過於頻繁會造成chromedriver反應不及而當掉
                         //timeDifference = DateTime.Now.Subtract(nextPageStartTime);
                         //if (timeDifference.TotalSeconds < 0.3)
                         //return;
                            if (ModifierKeys == Keys.Control)
                            {//按住Ctrl再按五鍵滑鼠的下一頁按鈕，則可以以預設的書頁圖大小來設定紅框以供輸入。可以網址來產生紅框如下： 20250202大年初五 感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主
                                //br.driver.Navigate().GoToUrl(br.driver.Url.Replace("#editor", "#box(2,14,792,1146)"));
                                br.driver.Navigate().GoToUrl(br.driver.Url.Replace("#editor", "#box(0,2,798,1288)"));
                                br.driver.Navigate().Refresh();
                                br.Input_picture();

                            }
                            //nextPageStartTime = DateTime.Now;
                            nextPages(Keys.PageDown, true);
                            //br.driver.SwitchTo().Window(driver.CurrentWindowHandle);
                        }
                        break;
                    default:
                        break;
                }
            }

        }

        private void textBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (ModifierKeys == Keys.None)
            {//滑鼠左鍵點二下，同 Ctrl + -（數字鍵盤） 會重設以插入點位置為頁面結束位國
             //if (autoPaste2QuickEdit && lines_perPage == 0)
             //{
                lines_perPage = countLinesPerPage(textBox1.Text.Substring(0, textBox1.SelectionStart));
                //}
                resetPageTextEndPositionPasteToCText();
                if (!autoPaste2QuickEdit) AvailableInUseBothKeysMouse();
            }
        }

        private bool resetPageTextEndPositionPasteToCText()
        {
            int s = textBox1.SelectionStart;
            if (s > 2 && textBox1.Text.Substring(s - 2, 2) == Environment.NewLine) s -= 2;
            pageTextEndPosition = s + textBox1.SelectionLength;//重設 pageTextEndPosition 值
            pageEndText10 = string.Empty;
            //if (keyDownCtrlAdd(false)) if (textBox1.Text != "") { pauseEvents(); textBox1.Text = ""; resumeEvents(); }
            return keyDownCtrlAdd(false);
        }

        private void Form1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //textBox1.Text = Clipboard.GetText();
        }

        private void textBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.Text) && textBox1.TextLength < 100)
            {
                string clpTxt = (string)e.Data.GetData(DataFormats.UnicodeText);
                //如果剪貼簿裡的文字內容長於99個字元，則執行相關的 Word VBA
                if (clpTxt.Length > 99)
                {
                    //根據剪貼簿裡的文本特徵來作動作
                    if (clpTxt.IndexOf("<scanbegin file=") > -1 && clpTxt.IndexOf(" page=") > -1)
                    {
                        Clipboard.SetText(clpTxt);
                        //以供 runWordMacro("中國哲學書電子化計劃.清除頁前的分段符號");

                        return;
                    }
                    else
                        dragEnterTxt(e);
                }
                else if (clpTxt.Contains("action=editchapter"))
                    Clipboard.SetText(clpTxt);
                //以供 runWordMacro("中國哲學書電子化計劃.清除頁前的分段符號");
                else
                    dragEnterTxt(e);
            }
            else
                dragEnterTxt(e);
        }

        private static void dragEnterTxt(DragEventArgs e)
        {
            //https://docs.microsoft.com/zh-tw/dotnet/desktop/winforms/advanced/walkthrough-performing-a-drag-and-drop-operation-in-windows-forms?view=netframeworkdesktop-4.8
            //https://docs.microsoft.com/zh-tw/dotnet/desktop/winforms/controls/enable-drag-and-drop-operations-with-wf-richtextbox-control?view=netframeworkdesktop-4.8
            if (e.Data.GetDataPresent(DataFormats.Text))
                //e.Effect = DragDropEffects.Copy;
                e.Effect = DragDropEffects.All;
            else
                e.Effect = DragDropEffects.None;
        }

        bool dragDrop = false;
        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {
            dragDrop = false;//●●●●●●●●●●●●●●●●●●● 20250301

            if (ModifierKeys == Keys.None)
            {
                //https://docs.microsoft.com/zh-tw/dotnet/desktop/winforms/controls/enable-drag-and-drop-operations-with-wf-richtextbox-control?view=netframeworkdesktop-4.8
                int i;
                String s;

                // Get start position to drop the text.  
                i = textBox1.SelectionStart;
                s = textBox1.Text.Substring(i);
                //textBox1.Text = textBox1.Text.Substring(0, i);

                // Drop the text on to the RichTextBox.  
                //textBox1.Text = textBox1.Text +
                //e.Data.GetData(DataFormats.Text).ToString();
                //e.Data.GetData(DataFormats.UnicodeText).ToString();
                //textBox1.Text = textBox1.Text + s;

                ////textBox1.Text = e.Data.GetData(DataFormats.Text).ToString();
                string dropStr = e.Data.GetData(DataFormats.UnicodeText).ToString();
                if (dropStr.IndexOf("ctext.org/library.pl?") > -1 && dropStr.Length < 80)
                {
                    //if (dropStr.IndexOf("https://") == -1) dropStr = "https://" + dropStr;
                    //textBox3.Text = dropStr;
                    //new SoundPlayer(@"C:\Windows\Media\recycle.wav").Play();
                    //dragDropUrqopd l = true;
                    textBox3_DragDrop(sender, e);
                    //dragDropUrl = false;

                    #region 如果是在自動連續輸入下

                    if (autoPaste2QuickEdit)
                    {
                        //如果是在第1頁
                        if (textBox3.Text.IndexOf("&page=1&") > -1)
                            textBox1.Text = dropStr;

                    }
                    #endregion

                }
                else
                {
                    textBox1.Text = dropStr;//e.Data.GetData(DataFormats.UnicodeText).ToString();
                                            //textBox1.Text = e.Data.GetData(DataFormats.UnicodeText).ToString();
                    dragDrop = true;
                }
            }
            else if (ModifierKeys == Keys.Control)
            {
                //Clipboard.SetData(DataFormats.UnicodeText, e.Data.GetData(DataFormats.UnicodeText));//由 vs studio 的 intel 幫我填入的（按兩下tab鍵後）
                Clipboard.SetDataObject(e.Data);//由以上一行自己摸索、測試的
                textBox1.Text = "autoRunWordMacro=true……";
                runWordMacro("維基文庫四部叢刊本轉來");//不想被鎖死的話得考慮用多工、多執行緒了
                return;
            }
        }

        int[] findWord(string x, string x1)
        {
            if (x == "" || x1 == "") return null;
            if (x.Length > x1.Length) return null;
            int s, nextS;
            s = x1.IndexOf(x);
            nextS = x1.IndexOf(x, s + x.Length);
            return new int[] { s, nextS };
        }

        ToolTip toolTip = new ToolTip();
        private void tooltipConstructor(object sender, string tooltipText)
        {
            if (toolTip.GetToolTip((Control)sender) != tooltipText)
                toolTip.SetToolTip((Control)sender, tooltipText);
        }
        //bool dragDropUrl = false;
        private void textBox3_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_currentPageNum.IsNullOrEmpty())
            {
                //校正_currentPageNum的值
                string pageNum = GetPageNumFromUrl(textBox3.Text).ToString();
                if (_currentPageNum != pageNum)
                {
                    //Debugger.Break();
                    _currentPageNum = pageNum;
                    WindowHandles["currentPageNum"] = _currentPageNum;

                }
                tooltipConstructor(sender, "現在在第" + _currentPageNum + "頁"
                    + Environment.NewLine + Environment.NewLine + "file= " + GetBookID_fromUrl(textBox3.Text)
                    + Environment.NewLine + Environment.NewLine + "textBox3.Text = " + textBox3Text);
            }
        }


        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (!_eventsEnabled) return;
            string url = textBox3.Text;

            #region 取代#box 為 #editor，如 https://ctext.org/library.pl?if=gb&file=185615&page=200&editwiki=2330034#box(428,674,2,4)

            //if (url.IndexOf("#box(") > -1 && (url.IndexOf("&editwiki=") > -1)
            if (url.IndexOf("#box(") > -1)
            {
                playSound(soundLike.waiting, true);
                if (MessageBoxShowOKCancelExclamationDefaultDesktopOnly("網址內含「#box(……)」是否要偵錯？", string.Empty, true, MessageBoxDefaultButton.Button2) == DialogResult.OK)
                    Debugger.Break();
                else
                {
                    url = br.ReplaceUrl_Box2Editor(url);
                    bool events = _eventsEnabled;
                    PauseEvents();
                    textBox3.Text = url;
                    _eventsEnabled = events;
                    if (browsrOPMode != BrowserOPMode.appActivateByName)
                    {
                        try
                        {
                            string ur = br.GetDriverUrl;
                            if (ur.IndexOf("#box(") > -1)
                                //    br.driver.Navigate().GoToUrl(ur.Substring(0, ur.IndexOf("#box(")));
                                br.driver.Url = Form1.FixUrl＿ImageTextComparisonPage(ur, false, true);
                        }
                        catch (Exception ex)
                        {
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                        }
                    }
                }
            }


            if (url.IndexOf("#box(") > url.IndexOf("&editwiki="))
            {
                PauseEvents();
                //url = url.Substring(0, url.IndexOf("#box(")) + "#editor";
                url = FixUrl＿ImageTextComparisonPage(url, true, false);
                textBox3.Text = url;
                try
                {
                    if (browsrOPMode != BrowserOPMode.appActivateByName)
                        if (br.driver?.Url.StartsWith(url.Substring(0, url.IndexOf("#editor"))) == true) br.driver.Navigate().GoToUrl(url);
                }
                catch (Exception ex)
                {
                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                }
                ResumeEvents();
            }

            #endregion

            #region 取得現前頁碼
            setCurrentPageNum(url);
            #endregion

            //mainFromTextBox3Text = textBox3.Text;
            string oldValue = (string)textBox3.Tag;//chatGPT 20230108

            //if (!autoPastetoQuickEdit && mainFromTextBox3Text != "" && textBox3.Text.IndexOf("http") == 0 && browsrOPMode != BrowserOPMode.appActivateByName && oldValue != mainFromTextBox3Text)
            //{
            //    if (dragDropUrl)
            //    {
            //        Task.Run(() =>/*須使用多執行緒才不會出現以下錯誤:
            //                   The HTTP request to the remote WebDriver server for URL http://localhost:6164/session/c574811b5a05d8f951364b5156b15ff8/window/handles timed out after 4.5 seconds.
            //                   因為textBox3_DragDrop(sender, e)要用到這個函式（DragDrop事件執行時會吃掉系統焦點，則br.driver無法正常運行）20230108                               */
            //        {
            //            if (br.driver == null) br.driver = br.driverNew();
            //            //if (Clipboard.GetText().IndexOf("http") == 0) Clipboard.Clear();
            //            string url = br.driver.Url;//交給區域變數，才好監看
            //            if (oldValue != mainFromTextBox3Text && url != mainFromTextBox3Text) br.GoToUrlandActivate(mainFromTextBox3Text);
            //        });
            //        dragDropUrl = false;
            //    }
            //    else
            //    {
            //        if (br.driver == null) br.driver = br.driverNew();
            //        //if (Clipboard.GetText().IndexOf("http") == 0) Clipboard.Clear();                    
            //        if (oldValue != mainFromTextBox3Text) br.GoToUrlandActivate(mainFromTextBox3Text);
            //    }
            //}
            //Task.WaitAll();

            #region 重設判斷不正常行長度的變數。
            int bookID = GetBookID_fromUrl(textBox3Text);//連續的冊數間的bookID其實是不連續的
            int resID, editwikiID;
            if (string.IsNullOrEmpty(url))
            {
                resID = 0;
                editwikiID = 0;
            }
            else
            {
                editwikiID = GetEditwikiID_fromUrl(url);
            //OpenQA.Selenium.IWebElement ie = br.Full_text_search_textbox_searchressingle;
            retry:
                if (br.IsDriverInvalid())
                {
                    try
                    {
                        openNewTabWindow();
                        driver.Url = url;
                    }
                    catch (Exception)
                    {
                        RestartChromedriver();
                        goto retry;
                    }
                }
                OpenQA.Selenium.IWebElement ie = br.Title_Linkbox_Link;
                try
                {
                    //resID = ie == null ? 0 : int.Parse(ie.GetAttribute("value").Substring("wiki:".Length));
                    resID = ie == null ? 0 : int.Parse(ie.GetAttribute("href").Substring("https://ctext.org/library.pl?if=en&res=".Length));

                }
                catch (Exception)
                {
                    resID = 0;
                    //throw;
                }
            }
            if (previousBookID != bookID) previousBookID = bookID;
            //if (Math.Abs(previousBookID - bookID) > 1 || url == string.Empty)
            //if (autoPaste2QuickEdit && (previousResID == 0 || (previousResID != resID && resID > 0)))
            if (previousResID == 0 || (previousResID != resID && resID > 0))
            { //normalLineParaLenggth = 0;

                //if (url != string.Empty) Debugger.Break(); //just for test 
                playSound(soundLike.done, true);
                if (previousResID == 0 || (previousResID != resID && resID > 0))
                    //if (DialogResult.OK == MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否要重設書籍版面資訊？"))
                    resetBooksPagesFeatures();
                previousResID = resID;
                if (editwikiID > 0 && editwikiID != previousEditwikiID) previousEditwikiID = editwikiID;
                if (textBox3.Text.StartsWith("http"))
                    playSound(soundLike.warn);
            }
            //else
            //{
            //    if (editwikiID > 0 && editwikiID != previousEditwikiID)
            //    {
            //        resetBooksPagesFeatures();
            //        previousResID = resID;
            //        previousEditwikiID = editwikiID;
            //        playSound(soundLike.warn);
            //    }
            //}

            #endregion


            #region OCR成功後則刪除下載的書圖,備份OCR結果; 因為 https://gj.cool/try_ocr 頁面時常傳回假資料（之前曾識別的文本），故今改寫在 textBox3.TextChanged事件中
            //OCR成功後則刪除下載的書圖,備份OCR結果
            string downloadImgFullName = string.Empty;//, imgUrl = br.GetImageUrl(); //bool imgResult = false;
                                                      //if (imgUrl != "")
            if (_previousPageNum != _currentPageNum ||
                (_previousPageNum == _currentPageNum && previousBookID != GetBookID_fromUrl(textBox3Text))
                || (_previousPageNum == _currentPageNum && previousBookID == GetBookID_fromUrl(textBox3Text)//頁碼一樣但章節不一樣時
                    && editwikiID != previousEditwikiID))
            {
                previousEditwikiID = editwikiID;//更新previousEditwikiID值
                                                //只要是換頁了就檢查
                                                //imgResult = downloadImage(imgUrl, out downloadImgFullName);
                                                //if (downloadImgFullName != "")
                                                //{
                downloadImgFullName = MydocumentsPathIncldBackSlash + "CtextTempFiles\\Ctext_Page_Image.png";
                if (File.Exists(downloadImgFullName))
                {
                    try
                    {
                        //Color cl = ForeColor;//表單若不出，如此設定沒意義。
                        //ForeColor = Color.AliceBlue;
                        File.Delete(downloadImgFullName);
                        //Thread.Sleep(350);
                        //ForeColor = cl;
                        //Task.Run(() =>
                        //{
                        //    using (SoundPlayer sp= new SoundPlayer("C:\\Windows\\Media\\recycle.wav")) { sp.Play(); }

                        //    //playSound(soundLike.error);
                        //});
                    }
                    catch (Exception ex1)
                    {
                        switch (ex1.HResult)
                        {
                            case -2147024864:
                                Task.Run(() =>
                                {
                                    Thread.Sleep(600);//"由於另一個處理序正在使用檔案 'X:\\Ctext_Page_Image.txt'，所以無法存取該檔案。"
                                    File.Delete(downloadImgFullName);
                                });
                                break;
                            default:
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex1.HResult + ex1.Message);
                                break;
                        }
                    }

                }
                //}
            }
            #endregion


            #region 記下這次的相關資訊以供下次參考
            textBox3.Tag = textBox3Text;
            _previousPageNum = _currentPageNum;
            #endregion

            //if (keyinTextMode) return;


            if (url == "")
            {
                resetBooksPagesFeatures(); previousBookID = 0; previousResID = 0;
                return;
            }
            if (url.IndexOf("ctext.org") > -1) if (url.IndexOf("https://") == -1) textBox3.Text = "https://" + url;
            if (oldValue == "" || oldValue == null) autoPastetoOrNot();



        }
        /// <summary>
        /// 設定目前頁碼欄位的值
        /// </summary>
        /// <param name="url"></param>
        private void setCurrentPageNum(string url)
        {
            if (url.IndexOf("&page=") > -1)
            {
                int s = url.IndexOf("&page=") + "&page=".Length;
                _currentPageNum = url.Substring(s, url.IndexOf("&", s) > -1 ? url.IndexOf("&", s) - s : url.Length - s);
            }
            else _currentPageNum = "";
        }

        /// <summary>
        /// 由指定的url中擷取出 book ID
        /// </summary>
        /// <returns>傳回ID值，找不到就傳回0</returns>
        internal int GetBookID_fromUrl(string url)
        {
            const string f = "file=";
            if (url == "" || url.IndexOf(f) == -1 || url.IndexOf("&") == -1) return 0;
            //if(url==string.Empty) url = textBox3Text; 
            int s = url.IndexOf(f);
            return int.Parse(url.Substring(s + f.Length, url.IndexOf("&", s + 1) - s - f.Length));
        }
        /// <summary>
        /// 由指定的url中擷取出 editwiki ID 的數值
        /// 如： https://ctext.org/library.pl?if=en&file=36583&page=27&editwiki=829144#editor 中的 829144
        /// </summary>
        /// <returns>傳回ID值，找不到就傳回0</returns>
        internal int GetEditwikiID_fromUrl(string url)
        {
            const string f = "editwiki=";
            if (url == "" || url.IndexOf(f) == -1 || url.IndexOf("&") == -1 || url.IndexOf("#") == -1) return 0;
            //if(url==string.Empty) url = textBox3Text; 
            int s = url.IndexOf(f);
            return int.Parse(url.Substring(s + f.Length, url.IndexOf("#", s + 1) - s - f.Length));
        }
        /// <summary>
        /// 由指定的url中擷取出 頁數
        /// </summary>
        /// <returns>傳回頁碼值</returns>
        internal int GetPageNumFromUrl(string url)
        {
            const string p = "page=";
            if (url == "" || url.IndexOf(p) == -1 || url.IndexOf("&") == -1) return 0;
            int s = url.LastIndexOf("#");
            if (s > -1) url = url.Substring(0, s);
            s = url.IndexOf(p);
            return int.Parse(url.Substring(s + p.Length, url.IndexOf("&", s + 1) == -1 ? url.Length - (s + p.Length) : url.IndexOf("&", s + 1) - s - p.Length));
        }

        private void autoPastetoOrNot()
        {
            string x = textBox3.Text;
            if (x == "") return;
            if (x.IndexOf("https://ctext.org/") == -1 || x.IndexOf("edit") == -1) return;
            //const string f = "file="; int s = x.IndexOf(f);
            int bookID = GetBookID_fromUrl(textBox3Text);// int.Parse(x.Substring(s + f.Length, x.IndexOf("&", s + 1) - s - f.Length));

            if (autoPaste2QuickEdit && (bookID != previousBookID || previousBookID == 0))
            {
                new SoundPlayer(@"C:\Windows\Media\Windows Notify Messaging.wav").Play();
                if (Math.Abs(bookID - previousBookID) > 1) if (MessageBox.Show("是否更新頁面每行字數及每頁行數等資訊？", "", MessageBoxButtons.OKCancel
                       , MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly) == DialogResult.OK)
                        resetBooksPagesFeatures();

                if (autoPaste2QuickEdit == false)
                {
                    new SoundPlayer(@"C:\Windows\Media\Windows Exclamation.wav").Play();
                    //https://www.facebook.com/oscarsun72/posts/4780524142058682
                    //messagebox topmost
                    if (MessageBox.Show("AUTO paste to Ctext Quick Edit textBox ?", "", MessageBoxButtons.OKCancel
                            , MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly) == DialogResult.OK)
                    {
                        //autoPastetoQuickEdit = true; autoPasteFromSBCKwhether = false;
                        turnOn_autoPastetoQuickEdit();
                        //if (MessageBox.Show("是否先檢查文本先前是否曾編輯過？" + Environment.NewLine +
                        //    "要檢查的話，請先複製其文本，再按確定（ok）按鈕", "", MessageBoxButtons.OKCancel
                        //    , MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly) == DialogResult.OK)
                        //{
                        //    runWordMacro("checkEditingOfPreviousVersion");
                        //}
                    }
                    else
                        autoPaste2QuickEdit = false;
                }
            }
            previousBookID = bookID;
        }
        /// <summary>
        /// 各階標題的空格有幾個,第1階為第1元素,元素的值即該階標題的前置空格數
        /// </summary>
        internal List<int> TitleLeadingSpacesCount = new List<int>();

        /// <summary>
        /// 重設書本的頁面資訊（一頁幾行，一行幾字，各階標題的空格有幾個……）。
        /// 可藉由textBox3.Text值的改變（不同的書ID值）即會自動執行此項
        /// </summary>
        private void resetBooksPagesFeatures()
        {

            linesParasPerPage = -1;//每頁行/段數
            wordsPerLinePara = -1;//每行/段字數 reset
            pageTextEndPosition = 0; pageEndText10 = "";
            lines_perPage = 0;
            //normalLineParaLength = 0;
            NormalLineParaLength = 0; //wordsPerLinePara = -1;
            abnormalLineParaChecking = true;
            //resetPageTextEndPositionPasteToCText();//不知何時誤貼的，到無問題時，即可刪去
            //TopLine = false; Indents = true;
            TopLine = false; Indents = false;
            TitleLeadingSpacesCount.Clear();
            inputTextFrontPage = "{{{封面}}}<p>";
        }

        private void textBox3_DragDrop(object sender, DragEventArgs e)
        {
            string data = e.Data.GetData(DataFormats.UnicodeText).ToString();
            //textBox3.DoDragDrop(e.Data, DragDropEffects.Copy);            
            if (textBox3.Text == data) return;

            //if (browsrOPMode != BrowserOPMode.seleniumNew ||
            //    br.driver == null || chromedriversPID.Count == 0)
            //{
            PauseEvents();
            textBox3.Text = data;
            ResumeEvents();
            //}
            //else
            //    textBox3.Text = data;

            textBox1.Select(0, 0); textBox1.ScrollToCaret();
            new SoundPlayer(@"C:\Windows\Media\recycle.wav").Play();
        }

        private void textBox3_DragEnter(object sender, DragEventArgs e)
        {
            dragEnterTxt(e);
        }

        private void Form1_Deactivate(object sender, EventArgs e)
        {//預設表單視窗為最上層顯示，當表單視窗不在作用中時，自動隱藏至系統右下方之系統列/任務列中，當滑鼠滑過任務列中的縮圖ico時，即還原/恢復視窗窗體
            if (!EventsEnabled) return;
            //if (!textBox2.Focused && textBox1.Text != "" && !dragDrop &&
            //    !autoPasteFromSBCKwhether) this.TopMost = false;//hideToNICo();
            if (!textBox2.Focused && !dragDrop &&
                            !autoPasteFromSBCKwhether) this.TopMost = false;//hideToNICo();
            if (dragDrop) dragDrop = false;
            selStart = textBox1.SelectionStart; selLength = textBox1.SelectionLength;
            //if (this.WindowState==FormWindowState.Minimized)
            //{
            //    hideToNICo();
            //}
        }

        /// <summary>
        /// Ctrl + w 關閉 Chrome 網頁頁籤
        /// </summary>
        void closeChromeTab()
        {
            chromeSendkeys("^{F4}");

        }
        /// <summary>
        /// 對Chrome瀏覽器送出按鍵（由closeChromeTab()抽離出來）20241022
        /// </summary>
        /// <param name="keys"></param>
        private void chromeSendkeys(string keys)
        {
            appActivateByName();
            Thread.Sleep(115);
            SendKeys.Send(keys);//關閉頁籤
                                //switch (browsrOPMode)
                                //{
                                //    case BrowserOPMode.appActivateByName:
                                //        appActivateByName();
                                //        SendKeys.Send("^{F4}");//關閉頁籤
                                //        break;
                                //    case BrowserOPMode.seleniumNew:
                                //        if (br.driver != null && Active)//表單不在最前面時也會觸發
                                //        {
                                //            //br.GoToCurrentUserActivateTab();
                                //            //br.driver.Navigate().Refresh();
                                //            br.driver.Close();
                                //        }
                                //        break;
                                //    case BrowserOPMode.seleniumGet:
                                //        break;
                                //}
            Thread.Sleep(115);
            bool autoPastetoQuickEditMemo = autoPaste2QuickEdit;
            autoPaste2QuickEdit = false;
            //this.Activate();
            AvailableInUseBothKeysMouse();
            autoPaste2QuickEdit = autoPastetoQuickEditMemo;
        }

        void closeChromeWindow()
        {//Ctrl + Shift + w 關閉 Chrome 網頁視窗
            appActivateByName();
            SendKeys.Send("%{F4}");//關閉頁籤
            bool autoPastetoQuickEditMemo = autoPaste2QuickEdit;
            autoPaste2QuickEdit = false;
            this.Activate();
            autoPaste2QuickEdit = autoPastetoQuickEditMemo;
        }

        bool isShortLine(string nextLine, string currentLine = "", ado.Connection cnt = null, ado.Recordset rst = null)
        {//傳回 fasle 則此行末不加<p>，因其非短行，乃抬頭等故耳；傳回true才加<p>
         //ado.Connection cnt = new ado.Connection();
         //ado.Recordset rst = new ado.Recordset();
            bool cntClose = false, rstClose = false, flg = true;//const string tableName = "每行字數判斷用";
            if (cnt == null)
            {
                Mdb.openDatabase("查字.mdb", ref cnt);
                cntClose = true;
            }
            //SELECT 每行字數判斷用.term, 每行字數判斷用.condition FROM 每行字數判斷用 WHERE(((每行字數判斷用.condition) = 0)) ORDER BY 每行字數判斷用.term DESC;
            if (rst == null) { rst.Open("select * from 每行字數判斷用 where condition=0 ORDER BY term DESC;", cnt, ado.CursorTypeEnum.adOpenKeyset, ado.LockTypeEnum.adLockReadOnly); rstClose = true; }
            nextLine = nextLine.TrimStart("　".ToCharArray());//頂端的空格縮排不計
            while (!rst.EOF)
            {
                string trm = rst.Fields["term"].Value.ToString();
                if (nextLine.IndexOf(trm) == 0)
                {
                    flg = false;
                    ado.Recordset rstDoubleCheck = new ado.Recordset();
                    string trmDoubleCheck;
                    //檢查後綴不能是什麼詞彙；condition欄位=6
                    rstDoubleCheck.Open("select * from 每行字數判斷用 where condition=6 and instr(term,\"" + trm + "\")>0", cnt, ado.CursorTypeEnum.adOpenForwardOnly, ado.LockTypeEnum.adLockReadOnly);

                    while (!rstDoubleCheck.EOF)
                    {
                        trmDoubleCheck = rstDoubleCheck.Fields["term"].Value.ToString();
                        if (nextLine.IndexOf(trmDoubleCheck) == 0)
                        {
                            rstDoubleCheck.Close(); rstDoubleCheck = null;
                            if (rstClose) { rst.Close(); rst = null; } else rst.MoveFirst(); if (cntClose) { cnt.Close(); cnt = null; }
                            return true;
                        }
                        rstDoubleCheck.MoveNext();
                    }

                    //檢前後若是什麼詞彙；condition欄位=4
                    if (rstDoubleCheck.State != 0)//ado.ObjectStateEnum.adStateClosed==0
                    {
                        rstDoubleCheck.Close();
                    }
                    rstDoubleCheck.Open("select * from 每行字數判斷用 where condition=4 and instr(term,\"" + trm + "\")>0", cnt, ado.CursorTypeEnum.adOpenForwardOnly, ado.LockTypeEnum.adLockReadOnly);

                    while (!rstDoubleCheck.EOF)
                    {
                        trmDoubleCheck = rstDoubleCheck.Fields["term"].Value.ToString();
                        string trmPrprefix = trmDoubleCheck.Substring(0, trmDoubleCheck.IndexOf(trm));
                        if (currentLine.LastIndexOf(trmPrprefix) + trmPrprefix.Length == currentLine.Length)
                        {
                            rstDoubleCheck.Close(); rstDoubleCheck = null;
                            if (rstClose) { rst.Close(); rst = null; } else rst.MoveFirst(); if (cntClose) { cnt.Close(); cnt = null; }
                            return false;
                        }
                        if (!flg) flg = true;
                        rstDoubleCheck.MoveNext();
                    }
                    rstDoubleCheck.Close(); rstDoubleCheck = null;
                    if (!flg)
                    {
                        if (rstClose) { rst.Close(); rst = null; } else rst.MoveFirst(); if (cntClose) { cnt.Close(); cnt = null; }
                        return flg;
                    }
                }
                rst.MoveNext(); if (!flg) flg = true;
            }
            if (rstClose) { rst.Close(); rst = null; } else rst.MoveFirst(); if (cntClose) { cnt.Close(); cnt = null; }
            return flg;

        }

        private void button1_MouseDown(object sender, MouseEventArgs e)
        {//預設執行分行分段。然切換到自動連續輸入模式時，會轉成送出貼上 [簡單修改模式]（quick edit文字框）的功能。若切換到手動輸入模式時，才有必要分行分段的功能 20230107
            if (e.Button == MouseButtons.Left)
            {
                if (ModifierKeys == Keys.Control)
                {
                    //分行分段按鈕：若有按下Ctrl才按此鈕則執行圖文脫鉤 Word VBA
                    runWordMacro("中國哲學書電子化計劃.撤掉與書圖的對應_脫鉤");
                    textBox1.Focus();
                    return;
                }
                if (button1.Text == "分行分段")
                {
                    splitLineByFristLen();
                    textBox1.Focus();
                }
                else if (button1.Text == "送出貼上")
                {
                    //if (keyDownCtrlAdd(ModifierKeys == Keys.Shift)) if (textBox1.Text != "") { pauseEvents(); textBox1.Text = ""; resumeEvents(); }
                    keyDownCtrlAdd(ModifierKeys == Keys.Shift);
                }
            }

        }
        /// <summary>
        /// 下載書圖以供OCR用
        /// </summary>
        /// <param name="imageUrl">書圖網址</param>
        /// <param name="downloadImgFullName">下載路徑全檔名</param>
        /// <param name="selectedInExplorer">是否在載後於檔案總管開啟、並將所下載之檔案選取</param>
        /// <returns>若下載成功則傳回true</returns>
        internal bool DownloadImage(string imageUrl, out string downloadImgFullName, bool selectedInExplorer = false)
        {
            if (imageUrl == "")
            {
                downloadImgFullName = ""; return false;
            }
            downloadImgFullName = MydocumentsPathIncldBackSlash + "CtextTempFiles\\Ctext_Page_Image.png";
            ////若圖已存在則不復下載，因OCR成功後會刪除此圖故//避免隔太久又忘了刪除圖檔，還是改以下判斷
            if (File.Exists(downloadImgFullName)) return true;
            //////若圖檔已存在，且是2.5分鐘前存檔的，則不復下載，以免重複，又免誤按。20230404，改 google keep的快捷鍵以免誤按
            //if (File.Exists(downloadImgFullName) &&
            //    (DateTime.Now.Subtract(File.GetLastWriteTime(downloadImgFullName)).TotalMinutes < 1.5
            //    && !(DateTime.Now.Subtract(File.GetLastWriteTime(downloadImgFullName)).TotalDays >= 1))) return;

            /*20230103 creedit,chatGPT：
          你可以使用 Selenium 來下載網絡圖片。
            首先，你需要獲取圖片的 URL。然後，使用 WebClient 的 DownloadData 方法下載圖片的二進制數據。
            最後，使用 FileStream 將二進制數據寫入文件即可。  
          */
            // 獲取圖片的 URL。
            //imageUrl = "https://example.com/image.png";

            TopMost = false;

            bool returnVal = true;
            try
            {



                #region //-2146233079：遠端伺服器傳回一個錯誤: (404) 找不到。

                // 使用 WebClient 下載圖片的二進制數據。
                System.Net.WebClient webClient = new System.Net.WebClient();
                //webClient.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537");// Copilot大菩薩 20240430：這段程式碼將 WebClient 的 User-Agent 設定為一個常見的瀏覽器（在這個例子中是 Chrome），然後嘗試下載圖片。如果這個方法仍然無法解決問題，那麼可能需要更深入地研究該網站的防爬蟲機制，或者聯繫網站管理員以獲取更多信息。
                byte[] imageBytes = webClient.DownloadData(imageUrl);//-2146233079：遠端伺服器傳回一個錯誤: (404) 找不到。

                // 將二進制數據寫入文件。            
                using (FileStream fileStream = new FileStream(downloadImgFullName, FileMode.Create))
                {
                    fileStream.Write(imageBytes, 0, imageBytes.Length);
                    //Console.WriteLine("圖片已成功下載。");//在「即時運算視窗」寫出訊息
                }
                if (File.Exists(downloadImgFullName))
                    //取得 downloadImgFullName 的檔案大小，20250408 GitHub　Copilot大菩薩：
                    if (new FileInfo(downloadImgFullName).Length <= 7620)
                    {
                        File.Delete(downloadImgFullName);
                        returnVal = br.DownloadImage(imageUrl, downloadImgFullName);
                    }
                #endregion
            }
            catch (Exception ex)
            {
                if (ex.HResult == -2146233079 && (
                    ex.Message.StartsWith("遠端伺服器傳回一個錯誤: (404) 找不到。")
                    || ex.Message.StartsWith("遠端伺服器傳回一個錯誤: (403) 禁止。")
                    || ex.Message.StartsWith("要求已經中止: 無法建立 SSL/TLS 的安全通道。")
                    || ex.Message.StartsWith("無法連接至遠端伺服器")))
                    returnVal = br.DownloadImage(imageUrl, downloadImgFullName);
                //20240430 Copilot大菩薩：如果 WebClient 的 DownloadData 方法無法滿足需求，那麼您可能需要考慮使用其他的方法來下載圖片。我之前提到的兩種方法是：
                //使用 Selenium 模擬瀏覽器操作：這種方法可以模擬「另存圖片」的操作，但可能需要一些複雜的程式碼，並且可能需要安裝特定的瀏覽器擴充功能。
                //使用 HttpClient 或其他第三方函式庫：這些函式庫通常提供了更靈活和強大的功能，可以處理更複雜的網路操作，例如處理 cookies、session、referer 等等。
                else
                {
                    MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message, "DownloadImage Fail！！");
                    return false;
                }
            }
            #region 在下載完後在檔案總管中將其選取MyRegion
            if (selectedInExplorer)
            {
                //https://www.ruyut.com/2022/05/csharp-open-in-file-explorer.html
                //把之前開過的關閉
                if (prcssDownloadImgFullName != null && prcssDownloadImgFullName.HasExited)
                {
                    prcssDownloadImgFullName.WaitForExit();
                    //prcssDownloadImgFullName.CloseMainWindow();
                    prcssDownloadImgFullName.Close();
                    ////prcssDownloadImgFullName.WaitForExit();
                    //prcssDownloadImgFullName.Kill();
                }
                prcssDownloadImgFullName = System.Diagnostics.Process.Start("Explorer.exe", $"/e, /select ,{downloadImgFullName}");

                //以下chatGPT的 無效,上面才有效
                //ProcessStartInfo startInfo = new ProcessStartInfo
                //{
                //    FileName = downloadImgFullName,
                //    UseShellExecute = true,
                //    Verb = "select"//選取
                //    //Verb = "open" //打開
                //};
                //Process.Start(startInfo);
                ////並將之打開
                //Process.Start(downloadImgFullName);
            }
            #endregion
            return returnVal;//true;
        }
        Process prcssDownloadImgFullName;
        /// <summary>
        /// 標題前的全形空格數 20250223
        /// </summary>
        private int titleLeadingSpaceCount;

        internal static void MessageBoxShowOKExclamationDefaultDesktopOnly(string text, string caption = "", bool formActivated = true)
        {
            Form1 form1 = Application.OpenForms[0] as Form1;
            //20221021Bing大菩薩：C# 跨執行緒作業無效：
            //form1.bringBackMousePosFrmCenter();
            if (br.ActiveForm1.InvokeRequired)
            {
                br.ActiveForm1.Invoke((MethodInvoker)delegate
                {
                    // 你的程式碼
                });
            }
            else
            {
                // 你的程式碼
                if (formActivated)
                {
                    form1.bringBackMousePosFrmCenter();
                }
            }
            MessageBox.Show(text, caption, MessageBoxButtons.OK
                , MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);

            //20221021Bing大菩薩：C# 跨執行緒作業無效：
            if (br.ActiveForm1.InvokeRequired)
            {
                br.ActiveForm1.Invoke((MethodInvoker)delegate
                {
                    // 你的程式碼
                });
            }
            else
            {
                // 你的程式碼
                if (formActivated)
                {
                    form1.BringToFront(); form1.AvailableInUseBothKeysMouse();
                }
            }

        }
        internal static DialogResult MessageBoxShowOKCancelExclamationDefaultDesktopOnly(string text, string caption = "", bool formActivated = true, MessageBoxDefaultButton defaultButton = MessageBoxDefaultButton.Button1)
        {
            Form1 form1 = Application.OpenForms[0] as Form1;
            form1.bringBackMousePosFrmCenter();
            DialogResult dr = MessageBox.Show(text, caption, MessageBoxButtons.OKCancel
                , MessageBoxIcon.Exclamation, defaultButton, MessageBoxOptions.DefaultDesktopOnly);
            if (formActivated)
            {
                if (br.ActiveForm1.InvokeRequired)
                {
                    br.ActiveForm1.Invoke((MethodInvoker)delegate
                    {
                        // 你的程式碼
                        form1.BringToFront(); form1.AvailableInUseBothKeysMouse();
                    });
                }
                else
                {
                    // 你的程式碼
                    if (formActivated)
                    {
                        form1.BringToFront(); form1.AvailableInUseBothKeysMouse();
                    }
                }
            }
            return dr;

        }

        #region 取得Windows作業系統現行的程式視窗。此乃為自己練習&測試用爾 https://ithelp.ithome.com.tw/questions/10212282#answer-388757        

        // 定義Windows API函數
        [DllImport("user32.dll")]
        static extern IntPtr GetDlgItem(IntPtr hDlg, int nIDDlgItem);

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, string lParam);

        // 定義WM_SETTEXT訊息常數
        const int WM_SETTEXT = 0x000C;

        internal static void keyinNotepadPlusplus(string getProcessesByName, string keyinText)
        {
            // 取得目標程式的視窗控制代碼（handle）
            //IntPtr targetWindow = Process.GetProcessesByName("notepad")[0].MainWindowHandle;
            IntPtr targetWindow = Process.GetProcessesByName("notepad++")[0].MainWindowHandle;

            // 取得目標程式的textbox輸入框的控制代碼（handle）
            IntPtr targetTextBox = GetDlgItem(targetWindow, 15);

            // 傳送WM_SETTEXT訊息和文字到textbox輸入框
            SendMessage(targetTextBox, WM_SETTEXT, 0, keyinText);
        }
        #endregion


        private void richTextBox1_Enter(object sender, EventArgs e)
        {//20230111 creedit YouChat：如果您想要在指定的控制項之前捕獲鼠標按下事件，您可以將控件的TabStop屬性設置為false，這樣就可以確保該控制項的Mousedown事件會先被捕獲。你可以使用以下示例代碼來實現：
            textBox1.TabStop = false;
        }


        /// <summary>
        /// 可以或必須直接取代之文字（以doit欄位控制是否執行）20250131大年初三增修
        /// </summary>
        /// <param name="tx">要取代的字串</param>
        /// <param name="whereNoteInstr">備註欄位要篩選的條件值</param>
        /// <param name="selection">只作選取區的取代則為true</param>
        void replaceXdirectly(ref string tx, string whereNoteInstr = "", bool selection = false)
        {// F11
         //string tx = textBox1.Text, rx;
            string rx;
            ado.Connection cnt = new ado.Connection();
            Mdb.openDatabase("查字.mdb", ref cnt);
            ado.Recordset rst = new ado.Recordset();
            if (whereNoteInstr == string.Empty)
                rst.Open("select * from 維基文庫等欲直接抽換之字 where doit=true order by len(replaced) desc", cnt, ado.CursorTypeEnum.adOpenForwardOnly, ado.LockTypeEnum.adLockReadOnly);
            else
                rst.Open("select * from 維基文庫等欲直接抽換之字 where (doit=true and instr(備註,\"" + whereNoteInstr + "\")>0) order by len(replaced) desc", cnt, ado.CursorTypeEnum.adOpenKeyset, ado.LockTypeEnum.adLockReadOnly);


            while (!rst.EOF)
            {
                rx = rst.Fields[0].Value.ToString();
                if (tx.IndexOf(rx) > -1)
                {
                    playSound(soundLike.notify, true);
                    tx = tx.Replace(rx, rst.Fields[1].Value.ToString());
                }
                rst.MoveNext();
            }
            rst.Close(); cnt.Close(); rst = null; cnt = null;//當您透過開啟 的 Recordset 物件結束作業時，請使用 Close 方法來釋放任何相關聯的系統資源。 關閉物件並不會從記憶體中移除它;您可以變更其屬性設定，並使用 Open 方法來稍後再次開啟它。 若要完全排除記憶體中的物件，請將物件變數設定為 Nothing。 https://docs.microsoft.com/zh-tw/sql/ado/reference/ado-api/open-method-ado-recordset?view=sql-server-ver16
            undoRecord();
            if (selection)
                textBox1.SelectedText = tx;
            else
                textBox1.Text = tx;
            //replaceBlank_ifNOTTitleAndAfterparagraphMark();

            //fixFormatErrorlike王文成公全書();//●●●●●●●●●●●●●●●●20250306先取消 因為會造成如《詞綜》排版的錯誤 https://ctext.org/library.pl?if=gb&file=63722&page=2 http://skqs.guoxuedashi.net/wen_3024p/178527.html

            caretPositionRecall();
        }
        /// <summary>
        /// 為送出WordVBA執行"中國哲學書電子化計劃.清除頁前的分段符號"巨集作準備
        /// 將文字版Edit編輯頁面中的文字內容複製至剪貼簿中
        /// </summary>
        /// <param name="clpTxt"></param>
        /// <returns>失敗或不執行則傳回false</returns>
        bool prepareEditTextReForm(string clpTxt = "")
        {
            if (driver == null) return false;
            if (clpTxt.IsNullOrEmpty()) clpTxt = Clipboard.GetText();
            if (clpTxt.Contains("action=editchapter"))
            {
                if (br.IsDriverInvalid())
                    br.driver.SwitchTo().Window(driver.WindowHandles.Last());
                else
                    br.driver.SwitchTo().Window(driver.CurrentWindowHandle);
                driver.Url = clpTxt;
                OpenQA.Selenium.IWebElement iwe = br.Textarea_data_Edit_textbox;
                if (iwe != null)
                {
                    if (clpTxt.Contains(" page=\"1\" />●\t<")) return false;
                    Clipboard.SetText(iwe.GetAttribute("value"));
                    //br.SetIWebElementValueProperty(iwe, string.Empty);
                    iwe.SendKeys(OpenQA.Selenium.Keys.Control + "a");
                    iwe.SendKeys(OpenQA.Selenium.Keys.Delete);
                    iwe.Click();
                    return true;
                }
            }
            return false;
        }

        void 歐陽文忠公集_集古錄跋尾校語專用()
        {//Alt + * //清除多餘的【】
            if (textBox1.SelectedText == "") return;
            int s = textBox1.SelectionStart;
            string x = textBox1.SelectedText;
            x = x.Replace("【", "").Replace("】", "");
            undoRecord();
            stopUndoRec = true;
            textBox1.SelectedText = x;
            stopUndoRec = false;
        }

        void 歐陽文忠公集_集古錄跋尾校語專用_()
        {//Alt + * //未完成
            int s = textBox1.SelectionStart;
            if (textBox1.SelectedText == "") textBox1.SelectAll();
            string x = textBox1.Text, xSel = textBox1.SelectedText;
            string[] rp = { "<p>" + Environment.NewLine, "{{", "}}" };
            string[] rpw = { "<p>" + Environment.NewLine + "【", "】{{【", "】}}" };
            undoRecord();
            stopUndoRec = true;

            #region get the range to handle
            int sp = xSel.IndexOf("<p>"), sn = 0;
            while (sp > -1)
            {
                String xSelP = xSel.Substring(sn, sp);
                string xSelN = xSelP.Substring(sn, xSelP.IndexOf("}}") + 2);
                for (int i = 0; i < rp.Length; i++)

                {
                    xSelN = xSelN.Replace(rp[i], rpw[i]);
                }

                sn = sp + 5;
                sp = xSel.IndexOf("<p>", sp + 1);

            }
            #endregion
            textBox1.SelectedText = "【" + xSel;
            stopUndoRec = false;

        }

        //重複字串、串接字串用，如「　」→「　　　　」……
        string concatenationStr(int count, string concatenationWhatStr)
        {
            string sp = "";
            while (count > 0)
            {
                count--;
                sp += concatenationWhatStr;
            }
            return sp;
        }
        //如首行凸排而其他縮進一字者，參差不齊，加以修齊；如《王文成公全書》中之《傳習錄》格式 https://ctext.org/library.pl?if=en&file=112524&page=32
        void fixFormatErrorlike王文成公全書()
        {
            string x = textBox1.Text; int s = x.IndexOf(Environment.NewLine), e = x.IndexOf(Environment.NewLine, s + 1);
            while (e > -1 && s + e <= x.Length)
            {
                string xp = x.Substring(s, e - s);
                if (0 <= xp.Length - "<p>".Length
                    && xp.Substring(xp.Length - "<p>".Length) == "<p>"
                    && xp.IndexOf("*") == -1)
                {

                    if (e + Environment.NewLine.Length + "􏿽".Length > x.Length) break;
                    if (x.Substring(e + Environment.NewLine.Length, "􏿽".Length) == "􏿽")
                    {
                        int eo = e;
                        while (e + 2 <= x.Length && x.Substring(e + 2, 2) == "􏿽")
                        {
                            e += 2;
                        }
                        int space_blank_Count = (int)(e - eo) / 2, en = x.IndexOf(Environment.NewLine, e + 2);
                        if (CountWordsLenPerLinePara(x.Substring(s, e - s - "<p>".Length - Environment.NewLine.Length))
                            + space_blank_Count == wordsPerLinePara
                            && e + 2 < x.Length &&
                            x.Substring(e + 2, en - (e + 2)).IndexOf("*") == -1)
                        {//後面行/段不能是標題（篇名），前面行/段不能太短-要剛好正常行/段長才行
                            string space_or_blanks = concatenationStr(space_blank_Count, "　");
                            x = x.Substring(0, s) + space_or_blanks +
                                x.Substring(s, eo - s - 3) + Environment.NewLine
                                + space_or_blanks + x.Substring(e + 2);
                        }
                    }

                }
                s = e + Environment.NewLine.Length;
                e = x.IndexOf(Environment.NewLine, e + 1);
            }
            if (!stopUndoRec)
            {
                undoRecord();
            }
            if (textBox1.Text != x) textBox1.Text = x;
        }

        #region 資料庫匯出
        void mdbExport()
        {//未完成
         //string bookName = "原抄本日知錄";
         //string rstStr = "SELECT 書.書名, 篇.卷, 篇.頁, 篇.末頁, 篇.篇名, 札.札記, 札.頁, 札.札ID, 札.類ID, 類別主題.類別主題" +
         //        "FROM 類別主題 INNER JOIN((書 INNER JOIN 篇 ON 書.書ID = 篇.書ID) INNER JOIN 札 ON 篇.篇ID = 札.篇ID) ON 類別主題.類ID = 札.類ID " +
         //        "WHERE(((書.書名) = \"" + bookName + "\") AND((類別主題.類別主題)Not Like \" * 真按 * \" Or(類別主題.類別主題) Is Null)) ORDER BY 篇.卷, 篇.頁, 篇.末頁, 札.頁, 札.札ID;";

            //const string cntStr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\千慮一得齋\書籍資料\開發_千慮一得齋.mdb";
            runWordMacro("中國哲學書電子化計劃.mdb開發_千慮一得齋Export");
        }
        #endregion

        #region 直書文字方塊 20230824 chatGPT大菩薩 中文文字方塊直書示例:根本就是失敗的。沒有用。沒變成直書，且亂排了一通。

        internal void AddVerticalTextBox()
        {
            InitializeVerticalTextBox();
        }

        private void Form1_Click(object sender, EventArgs e)
        {
            if (autoPaste2QuickEdit)
            {
                if (IsDriverInvalid())
                {
                    driver.SwitchTo().Window(driver.WindowHandles.Last());
                    LastValidWindow = driver.CurrentWindowHandle;
                }
                if (textBox1.TextLength == 0 && br.pageUBound < int.Parse(_currentPageNum) && IsValidUrl＿keyDownCtrlAdd(driver.Url))
                {
                    undoRecord(); stopUndoRec = true; PauseEvents();
                    if (!gotoNextChapter_FormatContentInput_SKQS())
                        Debugger.Break();
                    undoRecord(); stopUndoRec = false; ResumeEvents();
                }
                else
                    inputSKQSFrontPage();
            }
        }

        private void InitializeVerticalTextBox()
        {
            TextBox verticalTextBox = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Width = 100,
                Height = 200,
                Location = new Point(50, 50)
            };
            verticalTextBox.Paint += VerticalTextBox_Paint;
            // 設定文字方向為垂直
            verticalTextBox.Text = "感恩感恩\n南無阿彌陀佛";
            verticalTextBox.TextAlign = HorizontalAlignment.Center;
            verticalTextBox.RightToLeft = RightToLeft.Yes;
            Form1 form1 = newForm1();
            form1.Controls.Add(verticalTextBox);
            //this.Controls[this.Controls.Count - 1].Visible = true;
            form1.Controls["textBox1"].Visible = false;

        }

        private void VerticalTextBox_Paint(object sender, PaintEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            string text = textBox.Text;

            Font font = new Font("Arial", 12);
            Brush brush = Brushes.Black;
            float x = 0;
            float y = 0;
            float lineHeight = font.GetHeight();

            foreach (char c in text)
            {
                e.Graphics.DrawString(c.ToString(), font, brush, x, y);
                y += lineHeight;
            }
        }

        #endregion

        #region 《看典古籍》OCR API
        private bool PerformOCR()
        {
            //string imageUrl = br.GetImageUrl();
            //string result = await _ocrClient.GetOCRResult(imageUrl);
            string imagePath = MydocumentsPathIncldBackSlash + "CtextTempFiles\\Ctext_Page_Image.png", result = string.Empty;
            //if (DownloadImage(br.GetImageUrl(), out imagePath))
            DateTime dt = DateTime.Now;
            TopMost = false;
            br.BringToFront("chrome");
            br.driver.SwitchTo().Window(br.GetCurrentWindowHandle(br.driver));
            while (!File.Exists(imagePath))
            {
                //可按下Ctrl鍵中斷！！20241213
                if (ModifierKeys == Keys.Control) return false;
                if (DateTime.Now.Subtract(dt).TotalSeconds > 20)
                    if (MessageBoxShowOKCancelExclamationDefaultDesktopOnly("書圖下載尚未完成，是否繼續？") == DialogResult.Cancel)
                        return false;
            }
            if (_ocrClient == null) _ocrClient = new OCRClient();
            result = _ocrClient.GetOCRResult(imagePath);

            //Clipboard.Clear();
            if (!result.IsNullOrEmpty())
                // 在這裡處理OCR結果
                //Console.WriteLine(result);
                Clipboard.SetText(result);
            else
                return false;
            return true;
        }


        #endregion


    }
}
