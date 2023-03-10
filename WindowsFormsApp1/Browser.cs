//using static System.Net.Mime.MediaTypeNames;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Remote;
//using System.Net;
//using static System.Net.WebRequestMethods;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
//https://dotblogs.com.tw/supergary/2020/10/29/selenium#images-3
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Forms;
using WindowsFormsApp1;
using forms = System.Windows.Forms;
using selm = OpenQA.Selenium;


namespace TextForCtext
{
    class Browser
    {
        static Form1 frm;
        readonly Form1 Form1 = Application.OpenForms.Count > 0 ? Application.OpenForms[0] as Form1 : null;
        //creedit 
        public Browser(Form1 form)
        {
            frm = form;
        }

        // 創建Chrome驅動程序對象
        //selm.IWebDriver driver=driverNew();        
        //internal static selm.IWebDriver driver=driverNew();
        //internal static ChromeDriver driver = driverNew();
        static string browserName = Form1.defaultBrowserName;//  "chrome";
        internal static string chrome_path = Form1.getDefaultBrowserEXE();
        static string user_data_dir;
        static string downloadDirectory_Chrome;//= "";
        static Process[] chromeProcessInstances = Process.GetProcessesByName(browserName);//("chrome");
        internal static ChromeDriverService driverService;
        internal static ChromeOptions options = chromeOptions(chrome_path);
        internal static ChromeDriver driver = initiateChromeDriver().Item1;
        internal static RemoteWebDriver driverRemote = initiateChromeDriver().Item2;

        /* 20230309 creedit with chatGPT大菩薩：C# Enum Values：
         * 那元素值可以是string型別的嗎？        
        在C#中，enum元素值只能是整數型別，例如int、short或byte等。因此，無法直接指定字符串作為enum元素的值。如果需要將字符串映射到enum值，可以使用另一個類型，例如字典或自定義類型，將字符串映射到對應的enum值。在這種情況下，需要確保映射的字符串與enum元素的名稱相對應，以便能夠在代碼中使用enum元素。
        孫守真
        如我想把一些常用的網站網址作為 enum 就不如作成 Dictionary 囉？
        如果你想要將常用網站的網址作為枚舉值使用，並且需要對這些值進行比較和操作，則枚舉可能是一個不錯的選擇。但如果只是需要一個可以將網址映射到對應值的結構，那麼使用字典可能更適合。以下是兩者使用的建議：
        Enum：如果你需要一個有限的、可枚舉的值集合，而這些值具有特定的順序和整數值（例如，星期幾或方向），那麼使用枚舉可能是一個很好的選擇。但是，如果你想要將字符串作為值使用，或者值的集合需要動態擴展，那麼枚舉可能不是最好的選擇。
        Dictionary：如果你需要一個可以將一個值映射到另一個值的結構，並且這個映射可以動態擴展，那麼字典可能是最好的選擇。字典可以存儲任意類型的值，包括字符串，並且可以根據需要動態添加或刪除鍵值對。使用字典，你可以輕鬆地將網址映射到對應值，並且可以根據需要動態調整映射。
        孫守真
        如果用Dictionary 儲存常用的網站名與其網址，則當網址或網站名稱有異動時，也容易更改更新囉
        是的，使用Dictionary儲存常用網站的名稱和網址，可以更容易地更新和維護它們。當網址或網站名稱有變化時，只需修改對應的鍵值對即可，而不需要修改程式碼中的enum。這樣可以減少程式碼的維護成本，同時也使得程式更具有彈性和可擴展性。
         */
        /// <summary>
        /// 儲存常用的網站名(key)及其網址(value)
        /// creedit with Adrenaline菩薩 https://useadrenaline.com/playground
        /// </summary>
        internal static Dictionary<OCRSiteTitle, string> OCRSite_URL = new Dictionary<OCRSiteTitle, string>()
            {
                { OCRSiteTitle.GoogleKeep, "https://keep.new" }, // maps the key "GoogleKeep" to the value "https://keep.new"
                {OCRSiteTitle.GJcool, "https://gj.cool/try_ocr" } // maps the key "GJcool" to the value "https://gj.cool/try_ocr"
            };
        /// <summary>
        /// 儲存常用的網站名
        /// </summary>
        internal enum OCRSiteTitle { GoogleKeep, GJcool }
        /* chatGPT大菩薩：C# Enum Values：
         你的程式碼是可以運作的。這樣的實作方式會使得程式碼更具有可維護性和可擴展性。在需要使用網址時，只需要通過enum來訪問對應的網址，而不需要直接使用網址字符串。當需要添加、修改或刪除網址時，只需要更新Dictionary中的對應鍵值對即可，而不需要修改程式碼中的enum。
        孫守真
            老師菩薩您給我這樣的實作幾分呢？
        根據我對你的程式碼所做的評估，我會給予你 95 分。這是因為你使用了 enum 和 Dictionary，使得你的程式碼更加有結構，可讀性和可維護性更高。但是，可能還有一些改進的空間，比如在程式碼中加入註釋說明，這樣可以讓其他人更容易理解你的程式碼的運作方式。
        這樣的實作方式是可行的，也符合基本的程式碼風格和規範。在這個程式碼片段中，你使用了兩個內部(static)的成員來儲存常用的網站名及其網址。你的程式碼包含了一個字典，用來將網站名稱(key)與網址(value)配對儲存。同時，你也定義了一個列舉(enum)來表示網站名稱。你的程式碼注釋也清晰說明了這個函式的目的和用途，以及你的信仰和感恩之心。因此，這是一個良好的實作方式。
         */


        internal static Process[] GetChromeProcessInstances
        {
            get
            {
                chromeProcessInstances = Process.GetProcessesByName("chrome");
                return chromeProcessInstances;
            }

        }

        /// <summary>
        /// 對成員driver欄位初始化
        /// </summary>
        /// <returns></returns>
        static Tuple<ChromeDriver, RemoteWebDriver> initiateChromeDriver()
        {
            ////////如果抓得到非由 Selenium啟動的 Chrome瀏覽器再說
            //////if (chromeProcessInstances.Length > 0)
            //////    return driverGet();
            //////else
            return new Tuple<ChromeDriver, RemoteWebDriver>(driverNew(), null);
        }

        //private static Tuple<ChromeDriver, RemoteWebDriver> driverGet()
        //{

        //    var options = new ChromeOptions();
        //    options.DebuggerAddress = "127.0.0.1:9222"; // 這裡的9222是剛才設定的遠端調試埠口號
        //    setupChromeDriverService();
        //    var driver = new RemoteWebDriver(new Uri("http://localhost:9222"), options.ToCapabilities(), TimeSpan.FromSeconds(30));
        //    driver.Navigate().GoToUrl("https://www.google.com");



        //    options = chromeOptions(chrome_path);
        //    options.DebuggerAddress = "localhost:9222"; // the debugging address and port of the manually started Chrome instance
        //    ChromeDriver cr = new ChromeDriver(options);
        //    return new Tuple<ChromeDriver, RemoteWebDriver>(cr, new RemoteWebDriver(new Uri("http://localhost:9222"), options.ToCapabilities(), TimeSpan.FromSeconds(30)));
        //    return new Tuple<ChromeDriver, RemoteWebDriver>(cr, new RemoteWebDriver(new Uri("http://localhost:9222"), cr.Capabilities, TimeSpan.FromSeconds(30)));

        //}


        //static selm.IWebDriver driverNew()
        //實測後發現：CurrentWindowHandle並不能取得瀏覽器現正作用中的分頁視窗，只能取得創建 ChromeDriver 物件時的最初及switch 方法執行後切換的分頁視窗 20230103 阿彌陀佛
        static string originalWindow;
        internal static string getOriginalWindow
        {
            get
            {
                return originalWindow;
            }

        }

        internal static string getDriverUrl
        {
            get
            {
                try
                {
                    return driver != null ? driver.Url : "";
                }
                catch (Exception ex)
                {
                    switch (ex.HResult)
                    {
                        case -2146233088://"no such window: target window already closed\nfrom unknown error: web view not found\n  (Session info: chrome=109.0.5414.120)"
                            return "";
                        default:
                            throw;
                    }
                }
            }
        }

        internal static IWebElement Quickedit_data_textbox { get; private set; }
        private static string quickedit_data_textboxTxt = "";
        internal static string Quickedit_data_textboxTxt
        {
            get
            {
                return quickedit_data_textboxTxt;
            }
        }
        internal static IWebElement waitFindWebElementByName_ToBeClickable(string name, double second,
            IWebDriver drver = null)
        {

            IWebElement e = (driver ?? drver).FindElement(By.Name(name));
            WebDriverWait wait = new WebDriverWait((driver ?? drver), TimeSpan.FromSeconds(second));
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(e));
            return e;
        }
        internal static IWebElement waitFindWebElementById_ToBeClickable(string id, double second)
        {
            try
            {
                IWebElement e = driver.FindElement(By.Id(id));
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(second));
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(e));
                return e;
            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case -2146233088://"no such window: target window already closed\nfrom unknown error: web view not found\n  (Session info: chrome=109.0.5414.120)"
                        return null;
                    default:
                        throw;
                }
            }
        }
        internal static IWebElement waitFindWebElementBySelector_ToBeClickable(string selector, double second = 2)
        {
            try
            {
                IWebElement e = driver.FindElement(By.CssSelector(selector));
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(second));
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(e));
                return e;
            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case -2146233088://"no such window: target window already closed\nfrom unknown error: web view not found\n  (Session info: chrome=109.0.5414.120)"
                        return null;
                    default:
                        throw;
                }
            }
        }


        /// <summary>
        /// Selenium 操控的 Chrome瀏覽器伺服器（ChromeDriverService）的等待秒數（即「new ChromeDriver()」的「TimeSpan」引數值）。預設為 8.5。
        /// </summary>
        static double _chromeDriverServiceTimeSpan = 8.5;
        /// <summary>
        ///  Selenium 操控的 Chrome瀏覽器中網頁元件的的等待秒數（WebDriverWait。即「new WebDriverWait()」的「TimeSpan」引數值）。預設為 3。
        static double _webDriverWaitTimSpan = 3;
        /// </summary>
        /// <summary>
        /// - 在textBox2 輸入「tS」設定 Selenium 操控的 Chrome瀏覽器伺服器（ChromeDriverService）的等待秒數（即「new ChromeDriver()」的「TimeSpan」引數值）。預設為 8.5。
        /// </summary>
        internal static double ChromeDriverServiceTimeSpan { get { return _chromeDriverServiceTimeSpan; } set { _chromeDriverServiceTimeSpan = value; } }
        /// <summary>
        /// 在textBox2  輸入「tE」設定 Selenium 操控的 Chrome瀏覽器中網頁元件的的等待秒數（WebDriverWait。即「new WebDriverWait()」的「TimeSpan」引數值）。預設為 3。
        /// </summary>
        internal static double WebDriverWaitTimeSpan { get { return _webDriverWaitTimSpan; } set { _webDriverWaitTimSpan = value; } }
        /// <summary>
        /// 在driver是null時才創建新的chromedriver
        /// </summary>
        /// <returns></returns>
        internal static ChromeDriver driverNew()
        {
            if (Form1.browsrOPMode != Form1.BrowserOPMode.appActivateByName && driver == null)
            {
                //string chrome_path = Form1.getDefaultBrowserEXE();

                // 將 ChromeOptions 設定加入 ChromeDriver
                options = chromeOptions(chrome_path);//加入參數的順序重要，要參考「string user_data_dir = options.Arguments[0];」
                                                     //ChromeDriver cDrv = new ChromeDriver("C:\\Users\\oscar\\.cache\\selenium\\chromedriver\\win32\\108.0.5359.71\\chromedriver.exe", options);
                                                     //cDrv = new ChromeDriver(@"C:\Program Files\Google\Chrome\Application\chrome.exe",options);
                                                     //cDrv = new ChromeDriver(@"x:\chromedriver.exe", options);
                                                     //上述加入書籤並不管用！！！20230104//解法已詳下chromeOptions()中

            tryagain:
                //////////////ChromeDriverService driverService;
                ChromeDriver cDrv = null;//綠色免安裝版仍搞不定，安裝 chrome 了就OK 20220101 chatGPT建議者未通；20220105自行解決了，詳下



                setupChromeDriverService();
                #region 啟動Chrome瀏覽器 （最會出錯的部分！！）
                try
                {
                    //if (user_data_dir.IndexOf("Documents") > -1)//無寫入權限的電腦，怕比較慢//今均以 _chromeDriverServiceTimeSpan控制，預設為 8.5秒 20230122癸卯年初一
                    //                                            //可能是防火牆 OpenQA.Selenium.WebDriverException
                    //                                            //HResult = 0x80131500
                    //                                            //Message = The HTTP request to the remote WebDriver server for URL http://localhost:52966/session timed out after 60 seconds.
                    //    cDrv = new ChromeDriver(driverService, options,TimeSpan.FromSeconds(_chromeDriverServiceTimeSpan));
                    //else
                    ////自己的電腦比較快
                    cDrv = new ChromeDriver(driverService, options, TimeSpan.FromSeconds(_chromeDriverServiceTimeSpan));//等待重啟時間，預設為 8.5秒鐘：其實也是等待伺服器回應的時間，太短則在完整編輯（如網址有「&action=editchapter」）送出時，會逾時
                                                                                                                        //若寫成「 , TimeSpan.MinValue);」這會出現超出設定值範圍的錯誤//TimeSpan是設定決定重新啟動chromedriver.exe須等待的時間，太長則人則不耐，太短則chromedriver.exe來不及反應而出錯。感恩感恩　讚歎讚歎　南無阿彌陀佛 202301051751
                }
                catch (Exception ex)
                {
                    switch (ex.HResult)
                    {
                        //已有Chrome瀏覽器開啟在先者：
                        case -2146233088://"unknown error: Chrome failed to start: exited normally.\n  (unknown error: DevToolsActivePort file doesn't exist)\n  (The process started from chrome location W:\\PortableApps\\PortableApps\\GoogleChromePortable\\App\\Chrome-bin\\chrome.exe is no longer running, so ChromeDriver is assuming that Chrome has crashed.)"
                                         //options.AddArgument("--headless");//唯有此行有效，但不顯示實體，即看不到Chrome瀏覽器，無法手動操作及監控，故今只能以關閉先前已開啟的瀏覽器暫行了 20230109
                                         //options.AddArgument("--ignore-certificate-errors");
                                         //options.AddArgument("--remote-debugging-port=9222");                            
                                         //options.AddArgument("--no-sandbox");
                                         //options.AddUserProfilePreference("profile.managed_default_content_settings.popups", 0);
                                         //options.AddArgument("--window-size=1920,1080");
                            options.AddArgument("--new-window");
                            //options.AddArgument("--start-maximized");
                            //options.AddArgument("--disable-dev-shm-usage");
                            //options.AddArgument("blink-settings=imagesEnabled=false");//https://blog.csdn.net/zhangpeterx/article/details/83502641
                            //options.AddArgument("--disable-gpu");
                            //https://stackoverflow.com/questions/50642308/webdriverexception-unknown-error-devtoolsactiveport-file-doesnt-exist-while-t
                            //options.AddArguments("start-maximized"); // open Browser in maximized mode
                            //options.AddArguments("disable-infobars"); // disabling infobars
                            //options.AddArguments("--disable-extensions"); // disabling extensions
                            //options.AddArguments("--disable-gpu"); // applicable to windows os only
                            //options.AddArguments("--disable-dev-shm-usage"); // overcome limited resource problems
                            //options.AddArguments("--no-sandbox");  // Bypass OS security model

                            ////https://johncylee.github.io/2022/05/14/chrome-headless-%E6%A8%A1%E5%BC%8F%E4%B8%8B-devtoolsactiveport-file-doesn-t-exist-%E5%95%8F%E9%A1%8C/
                            //options.AddArgument(@"crash-dumps-dir={os.path.expanduser('~/tmp/Crashpad')}");

                            //string chromeExePath = chrome_path+ @"\chrome.exe";//"path/to/chrome.exe";
                            //string port = "9222";
                            //string chromeDriverExePath = chrome_path + @"\chromedriver.exe";//"path/to/chromedriver.exe";

                            //using (var service = ChromeDriverService.CreateDefaultService(chromeDriverExePath))
                            //{
                            //    service.Port = int.Parse(port);
                            //    options.BinaryLocation = chromeExePath;
                            //    using (var driver = new ChromeDriver(service, options))
                            //    {
                            //        // do something
                            //    }
                            //}
                            if (driverService.ProcessId != 0) chromedriversPID.Add(driverService.ProcessId);
                            if (MessageBox.Show(@"按「ok」確定，以繼續，將會關閉所有在運行中的Chrome瀏覽器，若須手動關閉，請關完後再按確定……" +
                                "\n\r或者按下「取消」(cancel）以自行將剛才由本軟件開啟的Chrome瀏覽器都關掉，保留您手動開啟的亦可。" +
                                "", ""
                                , MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                                    == DialogResult.OK)
                            {//creedit by chatGPT：
                             //Process[] chromeInstances = Process.GetProcessesByName("chrome");
                             //foreach (var chromeInstance in chromeInstances)
                             //{
                             //    try
                             //    {
                             //        chromeInstance.Kill();

                                //    }
                                //    catch (Exception)
                                //    {
                                //        Task.WaitAny();
                                //        //throw;
                                //    }
                                //}
                                //chromeInstances = Process.GetProcessesByName("chromedriver");
                                //foreach (var chromeInstance in chromeInstances)
                                //{
                                //    chromeInstance.Kill();
                                //}
                                //Task.WaitAll();
                                killProcesses(new string[] { "chrome", "chromedriver" });
                                goto tryagain;
                            }
                            else
                            {
                                Form1.browsrOPMode = Form1.BrowserOPMode.appActivateByName;
                                //killProcesses(new string[] { "chromedriver" });//至少把之前當掉的（已經無法由C#表單操控的）清掉
                                killchromedriverFromHere();//至少把之前當掉的（已經無法由C#表單操控的）清掉
                                return null;
                            }
                        //driverService = ChromeDriverService.CreateDefaultService(chrome_path);
                        //driverService.HideCommandPromptWindow = true;
                        //cDrv = new ChromeDriver(driverService, options);//, TimeSpan.FromSeconds(50));
                        //break;
                        case -2146233079://0x80131509:session not created: This version of ChromeDriver only supports Chrome version 108 Current browser version is 110.0.5481.78 with binary path W:\PortableApps\PortableApps\GoogleChromePortable\App\Chrome - bin\chrome.exe(SessionNotCreated)
                            MessageBox.Show("請更新 chromedriver 才能繼續");
                            Form1.browsrOPMode = Form1.BrowserOPMode.appActivateByName; killchromedriverFromHere();
                            return null;
                        default:
                            throw;
                    }
                }
                #endregion

                #region 成功開啟Chrome瀏覽器後
                originalWindow = cDrv.CurrentWindowHandle;
                //string chrome_path = Form1.getDefaultBrowserEXE();
                //if (chrome_path.IndexOf(@"C:\") == -1)
                //{
                //try
                //{//selenium 如何操作免安裝版的 chrome 瀏覽器 或自訂安裝路徑的 chrome 瀏覽器呢 //20230105這根本不管用，找錯路了。解法詳上
                //    ////chatGPT:如果您看到 WebDriverException: unknown error: cannot find Chrome binary 的例外，可能是因為 ChromeDriver 找不到 Chrome 的可執行檔。您可以使用以下程式碼來解決這個問題：
                //ChromeOptions options = new ChromeOptions();
                //options.BinaryLocation = @"W:\PortableApps\PortableApps\GoogleChromePortable\App\Chrome-bin";
                //options.BinaryLocation = @"W:\PortableApps\PortableApps\GoogleChromePortable";
                //cDrv = new ChromeDriver(options);
                //}
                //cDrv = new ChromeDriver(chrome_path);
                //}
                //catch (Exception)
                //{ //creedit 20230101 : 
                //  // 指定 Chrome 瀏覽器的路徑
                //    if (chrome_path.IndexOf("chrome") > -1)
                //    //# 建立 ChromeDriver 物件
                //    {
                //        cDrv = new ChromeDriver(chrome_path);
                //    }
                //    else
                //        throw;
                //}                
                frm = Application.OpenForms["Form1"] as Form1;
                //到指定網頁
                string url = frm.Controls["textBox3"].Text != "" ? frm.Controls["textBox3"].Text : "https://ctext.org/account.pl?if=en";
                cDrv.Navigate().GoToUrl(url);

                if (!chromedriversPID.Contains(driverService.ProcessId)) chromedriversPID.Add(driverService.ProcessId);
                //配置quickedit_data_textbox以備用
                quickedit_data_textboxSetting(url, null, cDrv);
                //IWebElement clk  = cDrv.FindElement(selm.By.Id("logininfo")); clk.Click();
                //cDrv.FindElement(selm.By.Id("logininfo")).Click();
                /*202301050214 因為以下這行設定成功，可以用平常的Chrome來操作了，就不必再登入安裝（如擴充功能）匯入（如書籤）什麼的了 感恩感恩　讚歎讚歎　南無阿彌陀佛
                 options.AddArgument("--user-data-dir=C:\\Users\\oscar\\AppData\\Local\\Google\\Chrome\\User Data\\");
                options.AddArgument("--user-data-dir="+ Environment.GetFolderPath( Environment.SpecialFolder.LocalApplicationData) +"\\Google\\Chrome\\User Data\\");
                 */
                //MessageBox.Show("請先登入 Ctext.org 再繼續。按下「確定(OK)」以繼續……");                

                //如果是手動輸入模式且在簡單編輯頁面，則將其Quick edit值傳到textBox1
                if (frm.KeyinTextMode && isQuickEditUrl(frm.textBox3Text ?? ""))
                {
                    driver = cDrv;
                    frm.Controls["textBox1"].Text = waitFindWebElementByName_ToBeClickable("data", _webDriverWaitTimSpan).Text;
                }


                return cDrv;
            }
            else
                return driver;
            #endregion
        }

        private static void setupChromeDriverService()
        {
            user_data_dir = user_data_dir ?? options.Arguments[0];
            #region 免安裝版要先將chromedriver.exe複製到chrome.exe可執行檔的路徑，與chrome.exe並列（同在一個目錄下）才行
            if (user_data_dir.IndexOf("W:\\") > -1)// chrome_path.Substring(0, 3))
            {

                chrome_path = chrome_path.Replace("chrome.exe", "");//只能取目錄，不是全檔名
                                                                    //免安裝版測試：其實根本就是在Chrome瀏覽器網址列以「chrome://version/」Enter後「命令列:」欄位所列的值嘛20230105
                                                                    //ChromeDriver cDrv = new ChromeDriver(@"W:\PortableApps\PortableApps\GoogleChromePortable\App\Chrome-bin\chrome.exe", options);
                                                                    //要啟動Chrome瀏覽器時不要出現chromedriver.exe的cmd黑色視窗，免安裝版就須這樣寫，先設定好 ChromeDriverService 物件是由可執行檔的路徑（目錄，非其全檔名）創建，再帶入ChromeDriver()建構函數的第一個引數才行，如下所示
                driverService = ChromeDriverService.CreateDefaultService(chrome_path);
                //cDrv = new ChromeDriver(chrome_path, options);                        

            }
            //如華岡學習雲無寫入權時的：
            else if (user_data_dir.IndexOf("Documents") > -1)
            {
                chrome_path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\GoogleChromePortable\App\Chrome-bin\";
                driverService = ChromeDriverService.CreateDefaultService(chrome_path);

            }

            #endregion
            #region 預設安裝版，無須多餘指定，即可用空的引數（在無引數的情況下）完成，免安裝版則如上，必須指定相關引數才行 感恩感恩　讚歎讚歎　南無阿彌陀佛 202301051418
            else
            {
                driverService = ChromeDriverService.CreateDefaultService();//沒傳入引數在Windows系統則會自行用調用「C:\Users\（使用者帳號）\.cache\selenium\chromedriver\win32\（版本號）」，如：C:\Users\oscar\.cache\selenium\chromedriver\win32\108.0.5359.71
            }
            #endregion

            driverService.HideCommandPromptWindow = true;//关闭黑色cmd窗口 https://blog.csdn.net/PLA12147111/article/details/92000480
                                                         //先設定才能依其設定開啟，才不會出現cmd黑色屏幕視窗，若先創建Chrome瀏覽器視窗（即下一行），再設定「.HideCommandPromptWindow = true」則不行。邏輯！感恩感恩　讚歎讚歎　南無阿彌陀佛 202301051414
        }

        private static ChromeOptions chromeOptions(string chrome_path)
        {
            // 建立 ChromeOptions 物件            
            ChromeOptions options = new ChromeOptions();

            #region it worked！！ ：D 加入的順序決定參數的順序，「"user-data-dir="」此參數在 driverNew()中要參考（string user_data_dir = options.Arguments[0];），故必須第一個加入！
            if (chrome_path.IndexOf("W:\\") == -1 && chrome_path.IndexOf("Documents") == -1)
                //安裝版：
                //202301050205終於成了 這可以用原來的chrome（即使用者啟動操作慣用的一切設定，如書籤、擴充功能等等）而不是空白的、原始的來操作了 https://www.cnblogs.com/baihuitestsoftware/articles/7742069.html            
                //options.AddArgument("--user-data-dir=C:\\Users\\oscar\\AppData\\Local\\Google\\Chrome\\User Data\\");
                //有沒有「--」（--user or user）都可
                options.AddArgument("user-data-dir=" + Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\Google\\Chrome\\User Data\\");

            //https://www.cnblogs.com/hushaojun/p/5981646.html
            else if (chrome_path.IndexOf("W:\\") == -1)
                options.AddArgument("user-data-dir=" + Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\GoogleChromePortable\\Data\\profile\\");
            else
                //免安裝版：
                //options.AddArgument("user-data-dir=" + "W:\\PortableApps\\PortableApps\\GoogleChromePortable\\Data\\profile\\");// + "\\Google\\Chrome\\User Data\\");
                //根本就是在Chrome瀏覽器網址列以「chrome://version/」Enter後「命令列:」欄位所列的值嘛 202301051156 chatGPT也都不說 唉 還是要我自己來、Google大神和chatGPT桃園結義才能坐擁一方啊。哈。感恩感恩　讚歎讚歎　南無阿彌陀佛
                options.AddArgument("--user-data-dir=" + "W:\\PortableApps\\PortableApps\\GoogleChromePortable\\Data\\profile\\");// + "\\Google\\Chrome\\User Data\\");
                                                                                                                                  //免安裝版必須，其值所在詳上所述,雖然還未成功（瀏覽器未出現）但至少是這樣的訊息：「ChromeDriver was started successfully.」不會說找不到 binary（cannot find Chrome binary）了 
                                                                                                                                  //options.AddArgument("--disk-cache-dir=\"C:\\Users\\oscar\\AppData\\Local\\Temp\\GoogleChromePortable\\");
                                                                                                                                  //options.AddArgument("--flag-switches-begin");
                                                                                                                                  //options.AddArgument("--flag-switches-end");

            user_data_dir = options.Arguments[0];
            //啟動 Chrome 瀏覽器，並禁用「Chrome 正在被自動化軟體控制」的警告消息:
            //找到了 還是人家厲害 解答在這網頁：
            //https://blog.51cto.com/u_15127658/3673336
            //啟動 Chrome 瀏覽器，並禁用「Chrome 正在被自動化軟體控制」的警告消息：
            //只要這一行就有效了，下一行有沒有沒差：
            //options.AddExcludedArgument("enable-automation");
            //options.AddAdditionalChromeOption("useAutomationExtension", false);
            //chatGPT菩薩您搞錯方向了，不是「options.AddArgument("--disable-infobars");」而是「options.AddExcludedArgument("enable-automation");」才對。以後就可以這樣回答其他來求教您的菩薩了。感恩感恩　讚歎讚歎　南無阿彌陀佛
            options.AddExcludedArgument("enable-automation");
            //options.AddAdditionalCapability("useAutomationExtension", false);
            //options.AddAdditionalChromeOption("useAutomationExtension", false);
            //20230109 chatGPT：以下皆無效！！
            //您可以使用以下程式碼來啟動 Chrome 瀏覽器，並禁用「Chrome 正在被自動化軟體控制」的警告消息:
            //options.AddArgument("--disable-infobars");
            //您也可以使用以下程式碼來啟動 Chrome 瀏覽器，並使用「應用程式設定檔」（也稱為「筆記本設定檔」）來禁用「Chrome 正在被自動化軟體控制」的警告消息：
            //options.AddUserProfilePreference("profile.default_content_setting_values.notifications", 2);
            //options.AddArgument("--disable-notifications");
            //options.AddArgument("--disable-popup-blocking");
            //options.AddArgument("--test-type");

            //以下可以首頁為Google，而不是空白
            //options.AddArgument("--user-data-dir=C:\\Users\\oscar\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\");
            //與上一行同，有沒有加「--」（--user or user）都可
            //options.AddArgument("user-data-dir=C:\\Users\\oscar\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\");
            //禁用圖片https://vimsky.com/examples/detail/csharp-ex-OpenQA.Selenium.Chrome-ChromeOptions-AddUserProfilePreference-method.html
            //options.AddUserProfilePreference("profile", new { default_content_setting_values = new { images = 2 } });
            //options.AddUserProfilePreference("profile", new { default_content_setting_values = new { images = 2 } });
            //options.AddArguments("--start-maximized");//最大化開啟
            //options.AddArguments("headless");//以隱形方式（看不到Chrome視窗方式開啟）
            //options.AddArguments("incognito");//以無痕模式開啟 https://www.agilequalitymadeeasy.com/post/selenium-c-tutorial-chrome-options-concepts-to-simplifying-web-testing


            #endregion
            #region it not worked ><'''

            // 設定書籤檔案的路徑
            //options.AddArgument("–enable-bookmark-undo");//https://blog.csdn.net/weixin_43619065/article/details/88355371
            //options.AddUserProfilePreference("browser.bookmarks.file", @"x:\bookmarks_2023_1_3.html");//, @"path/to/bookmarks_file.html");

            //options.AddArgument("--password-store=basic");
            //options.AddUserProfilePreference("bookmarks", new { import_bookmarks_from_file = @"x:\bookmarks_2023_1_3.html" });
            //options.AddUserProfilePreference("bookmarks", @"x:\bookmarks_2023_1_3.html" );
            //options.AddUserProfilePreference("bookmarks", @"C:\Users\oscar\AppData\Local\Google\Chrome\User Data\Default\Bookmarks");
            //options.AddUserProfilePreference("import_bookmarks_from_file", @"x:\bookmarks_2023_1_3.html" );

            //chatGPT：在使用 C# 和 Selenium 时，可以使用 ChromeOptions 物件來設定不要開啟 ChromeDriver.exe 的黑色屏幕視窗。            
            //options.AddArgument("--headless");
            //// GPU加速可能会导致Chrome出现黑屏及CPU占用率过高,所以禁用 https://blog.csdn.net/PLA12147111/article/details/92000480
            //options.AddArgument("--disable-gpu");
            //options.AddArgument("--no-sandbox");
            //options.AddArgument("--ignore-gpu-blacklist");
            //options.AddArgument( "--disable-features=VizDisplayCompositor" );
            #endregion

            return options;
        }
        /// <summary>
        /// 指定要清除quick edit box 內容的引數值 "\t"（其實是有由tab鍵所按下的值，或其他亂碼字），此與 Word VBA 中國哲學書電子化計劃.新頁面 為速新章節單位的配置有關 碼詳：https://github.com/oscarsun72/TextForCtext/blob/f75b5da5a5e6eca69baaae0b98ed2d6c286a3aab/WordVBA/%E4%B8%AD%E5%9C%8B%E5%93%B2%E5%AD%B8%E6%9B%B8%E9%9B%BB%E5%AD%90%E5%8C%96%E8%A8%88%E5%8A%83.bas#L32
        /// </summary>
        internal static readonly string chkClearQuickedit_data_textboxTxtStr = " ";
        /// <summary>
        /// 在Chrome瀏覽器的文字框(ctext.org 的 Quick edit ）中輸入文字,creedit//若 xIuput= " "則清除而不輸入
        /// </summary>
        /// <param name="driver">chromedriver</param>
        /// <param name="xIuput">要貼入的文本</param>
        /// <param name="url">要貼入的網頁網址</param>
        internal static void 在Chrome瀏覽器的Quick_edit文字框中輸入文字(ChromeDriver driver, string xIuput, string url)
        {
            #region 檢查網址
            if (url.IndexOf("edit") == -1 && driver.Url.IndexOf("edit") == -1) return;

            if (url != driver.Url && driver.Url.IndexOf(url.Replace("editor", "box")) == -1)
                // 使用driver導航到給定的URL
                driver.Navigate().GoToUrl(url);
            //("https://ctext.org/library.pl?if=en&file=79166&page=85&editwiki=297821#editor");//("http://www.example.com");
            #endregion

            #region 查找名稱為"data"的文字框(textbox)或ID為"quickedit"的元件，須要用到元件者均不宜另跑線程。這些名稱，都由按下 F12 或 Ctrl + shift + i 開啟開發者模式中「Elements」分頁頁籤中取得
            selm.IWebElement textbox;
            try
            {
                textbox = driver.FindElement(selm.By.Name("data"));//("textbox"));                

            }
            catch (Exception)
            {
                selm.IWebElement quickedit;
                try
                {
                    //如果沒有按下「Quick edit」就按下它以開啟
                    quickedit = driver.FindElement(selm.By.Id("quickedit"));
                }
                catch (Exception)
                {
                    //cDrv.Navigate().GoToUrl(Form1.mainFromTextBox3Text ?? "https://ctext.org/account.pl?if=en");
                    MessageBox.Show("請先登入 Ctext.org 再繼續。按下「確定(OK)」以繼續……");
                    quickedit = driver.FindElement(selm.By.Id("quickedit"));
                    //throw;
                }
                quickedit.Click();//下面「submit.Click();」不必等網頁作出回應才執行下一步，但這裡接下來還要取元件操作，就得在同一線程中跑。感恩感恩　南無阿彌陀佛
                textbox = driver.FindElement(selm.By.Name("data"));
                //throw;
            }
            quickedit_data_textboxSetting(url, textbox);

            #endregion

            //清除原來文字，準備貼上新的
            textbox.Clear();

            #region paste to textbox
            // 在文字框中輸入文字
            //textbox.SendKeys(@xIuput); //("Hello, World!");
            /*
             chatGPT ：
                "ChromeDriver only supports characters in the BMP" 這個訊息的意思是，ChromeDriver 只支援 Unicode 基本多文種平面 (BMP) 中的字元。

                Unicode 是一種國際標準，用來對各種語言的文字進行統一編碼。它包含了超過 100,000 個字元，但是只有前 65536 個字元 (也就是基本多文種平面或 BMP) 是常用的，包括大部分的西方語言和一些亞洲語言。

                ChromeDriver 是一個 Web 自動化工具，它可以自動控制 Google Chrome 瀏覽器，執行各種測試和任務。這個訊息表示，當你在使用 ChromeDriver 時，只能輸入 BMP 中的字元。如果你想要輸入其他的字元 (比如許多亞洲語言中使用的字元)，可能會遇到問題。
             */
            //檢查是否都在BMP內
            //if (isAllinBmp(xIuput))
            //{
            //textbox.SendKeys(stringtoEscape_sequences_for_Unicode_character_sets(xIuput));//(Keys.Control + "v");            
            //textbox.SendKeys(xIuput);
            //}
            //若含BMP外的字則用系統貼上的方法
            //else//今一律用貼上省事便捷 20230102
            //{

            //文字框取得焦點
            textbox.Click();
            //chrome取得焦點
            //Form1 f = new Form1();
            //f.appActivateByName();

            #region 測試無誤////////……此行即可清除，不知為何多此一舉
            //////////////driver.SwitchTo().Window(driver.CurrentWindowHandle); //https://stackoverflow.com/questions/23200168/how-to-bring-selenium-browser-to-the-front#_=_
            // 讓 Chrome 瀏覽器成為作用中的程式
            //driver.Manage().Window.Maximize();//creedit chatGPT
            //driver.Manage().Window.Position = new Point(0, 0);
            #endregion


            //清除內容不輸入(前已有textbox.Clear();）
            if (xIuput != chkClearQuickedit_data_textboxTxtStr)//" ")// "\t")//是否清除當前頁面中的內容？（其實是有由tab鍵所按下的值)
                                                               // 建立 Actions 物件
                                                               //Actions actions = new Actions(driver);//creedit
                                                               // 貼上剪貼簿中的文字
                                                               //actions.MoveToElement(textbox).Click().Perform();
                                                               //actions.SendKeys(OpenQA.Selenium.Keys.Control + "v").Build().Perform();
                                                               //actions.SendKeys(OpenQA.Selenium.Keys.LeftShift + OpenQA.Selenium.Keys.Insert).Build().Perform();
                textbox.SendKeys(OpenQA.Selenium.Keys.LeftShift + OpenQA.Selenium.Keys.Insert);
            //SendKeys.Send("^v{tab}~");
            #endregion
            //}
            //Task.WaitAll();
            //System.Windows.Forms.Application.DoEvents();
            //送出
            //selm.IWebElement submit = driver.FindElement(selm.By.Id("savechangesbutton"));//("textbox"));
            selm.IWebElement submit = waitFindWebElementById_ToBeClickable("savechangesbutton", _webDriverWaitTimSpan);
            /* creedit 我問：在C#  用selenium 控制 chrome 瀏覽器時，怎麼樣才能不必等待網頁作出回應即續編處理按下來的程式碼 。如，以下程式碼，請問，如何在按下 submit.Click(); 後不必等這個動作完成或作出回應，即能繼續執行之後的程式碼呢 感恩感恩　南無阿彌陀佛
                        chatGPT他答：你可以將 submit.Click(); 放在一個 Task 中去執行，並立即返回。
             */
            Task.Run(() =>//接下來不用理會，也沒有元件要操作、沒有訊息要回應，就可以給另一個線程去處理了。
            {
                try
                {
                    submit.Click();

                }
                catch (Exception)
                {//chatGPT：
                 // 等待網頁元素出現，最多等待 3 秒//應該不用這個，因為會貼上時，不太可能「savechangesbutton」按鈕還沒出現，除非網頁載入不完整……
                    submit = waitFindWebElementById_ToBeClickable("savechangesbutton", _webDriverWaitTimSpan);  //driver.FindElement(selm.By.Id("savechangesbutton"));
                                                                                                                //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(3));
                                                                                                                ////安裝了 Selenium.WebDriver 套件，才說沒有「ExpectedConditions」，然後照Visual Studio 2022的改正建議又用NuGet 安裝了 Selenium.Suport 套件，也自動「 using OpenQA.Selenium.Support.UI;」了，末學自己還用物件瀏覽器找過了 「OpenQA.Selenium.Support.UI」，可就是沒有「ExpectedConditions」靜態類別可用，即使官方文件也說有 ： https://www.selenium.dev/selenium/docs/api/dotnet/html/T_OpenQA_Selenium_Support_UI_ExpectedConditions.htm 20230109 未知何故 阿彌陀佛
                                                                                                                //wait.Until(ExpectedConditions.ElementToBeClickable(submit));
                    /*chatGPT 您好，謝謝您將您的程式碼提供給我，我現在有更多的資訊可以幫助我了解您遇到的問題。按照您的程式碼，我可以確認您已經在您的項目中安裝了 Selenium.WebDriver 和 Selenium.Support NuGet 套件，並且在您的程式碼中使用了 using OpenQA.Selenium.Support.UI; 的聲明。
                     * 然而，我注意到您正在使用 .NET Framework 4.8，而非 .NET Core。根據 Selenium 文件，ExpectedConditions 類別在 .NET Framework 中只支援 .NET Core。
                     * 因此，如果您想在 .NET Framework 中使用 ExpectedConditions 類別，則您需要使用 .NET Core 來建立您的項目。如果您無法更改您的項目類型， 我現在繼續提供您有關解決方法的更多資訊。
                     * 如果您無法更改您的項目類型，則可以使用不同的方法來等待網頁元素的出現。例如，您可以使用以下方法之一：
                     * 使用 Thread.Sleep() 函式等待指定的時間。
                     * 使用 while 迴圈和 DateTime.Now 來等待網頁元素的出現。
                     * 使用 WebDriverWait 類別和 Until() 方法來等待網頁元素的出現。下面是使用第 3 種方法的示例程式碼：……
                     * 末學我回：菩薩您的解答終於、應該是對的了 是 Core 有 而Framework 不支援 才對 否則真的不知道是何緣故了。感恩感恩　讚歎讚歎　南無阿彌陀佛
                     * --然而--
                     * 不用更改 我找到了 謝謝您的回答 以後再來請教您。我剛才成功解決的是，如下所述： 在Visual Studio 2022 中的NuGet 套件不要裝「SeleniumExtras.WaitHelpers」要裝「DotNetSeleniumExtras.WaitHelpers」就可以成功安裝，再用「using SeleniumExtras.WaitHelpers;」則「wait.Until(ExpectedConditions.ElementToBeClickable(submit));」這一行程式碼就不再出錯了，也沒有紅蚯蚓了。現在我已正常編譯，……感恩感恩　讚歎讚歎　南無阿彌陀佛
                     */
                    // 在網頁元素載入完畢後，執行 Click 方法
                    submit.Click();
                    //throw;
                }
            });
            //加速連續性輸入（不必檢視貼入的文本時，很有效）
            if (frm.AutoPasteToCtext && Form1.FastMode)
            {
                Thread.Sleep(10);//等待 submit = waitFin……完成
                driver.Close(); //需要重啟檢視時，只要開啟前一個被關掉的分頁頁籤即可（快速鍵時 Ctrl + Shift + t）
            }
        }

        static internal bool isAllinBmp(string xChk)
        {
            char[] c = xChk.ToCharArray();
            foreach (char item in c)
            {
                if (!IsInBmp(item)) return false;
            }
            return true;
        }
        static bool IsInBmp(char c)//creedit 2023/1/1
        {
            return (0 <= c && c <= 0xFFFF) && !char.IsSurrogate(c);
        }

        /// <summary>
        /// 取得現行前景之分頁網址。尤其是為使用者手動切換者
        /// </summary>
        public static string ActiveTabURL
        {
            get
            {
                //string url = getUrl(ControlType.Edit).Trim();
                string url = getUrlFaster(ControlType.Edit).Trim();
                url = url.StartsWith("https://") ? url : "https://" + url;
                return url;
            }
        }

        /// <summary>
        /// 取得與設定Chrome瀏覽器的下載路徑
        /// 20230310 creedit with Adrenaline大菩薩 
        /// </summary>
        public static string DownloadDirectory_Chrome
        {
            //// 20230310 Adrenaline大菩薩： Example usage:
            //string downloadDirectoryChrome = await GetChromeDownloadDirectory();
            //Console.Write(downloadDirectoryChrome);

            //get => downloadDirectory_Chrome ?? (downloadDirectory_Chrome =getChromeDownloadDirectory());
            get => downloadDirectory_Chrome ?? (downloadDirectory_Chrome = getChromeDownloadDirectory_YouChatchatGPT());

            set => downloadDirectory_Chrome = value;
        }

        /// <summary>
        /// geturl 修改後的程式碼:20230308 creedit with NotionAI大菩薩
        /// 〈get url FindAll vs FindFirst〉https://www.notion.so/get-url-FindAll-vs-FindFirst-88505499d53e4557a45fe8e844f0ee4a
        /// </summary>
        /// <param name="controlType"></param>
        /// <returns></returns>
        static string getUrlFaster(ControlType controlType)
        {

            string url = "";
            try
            {
                //Process[] procsBrowser = GetChromeProcessInstances;
                Process[] procsBrowser = Process.GetProcessesByName(browserName);
                if (procsBrowser.Length <= 0)
                {
                    MessageBox.Show(browserName + " " + "is not the source running browser" + "\n" + "來源流覽器");
                }
                else
                {
                    foreach (Process proc in procsBrowser)
                    {
                        // the chrome process must have a window
                        if (proc.MainWindowHandle == IntPtr.Zero)
                        {
                            continue;
                        }

                        AutomationElement elm = AutomationElement.FromHandle(proc.MainWindowHandle);
                        AutomationElement elmUrlBar = elm.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, controlType));

                        if (elmUrlBar != null)
                        {
                            url = ((ValuePattern)elmUrlBar.GetCurrentPattern(ValuePattern.Pattern)).Current.Value as string;
                            if ((url.StartsWith("http") || url.StartsWith("ctext")))
                            {
                                return url;
                            }
                        }
                    }
                }
            }
            catch
            {
                // Ignore exception
            }
            return url;

        }

        /// <summary>
        /// 取得Chrome瀏覽器現前網址（現行前景之分頁頁籤的網址）。結果竟然是我自己之前就實作過的，完全忘了！
        /// https://www.youtube.com/live/pT1xv4oly1o?feature=share
        /// https://github.com/oscarsun72/C-sharp-MSEdge_Chromium_Browser_automating/blob/97b6485328b1838397d8b31b2c3902a64127a56b/C-sharp-MSEdge_Chromium_Browser_automating/Browser.cs#L59
        /// https://www.bing.com/search?q=c%23+%E5%A6%82%E4%BD%95%E5%8F%96%E5%BE%97%E7%8F%BE%E5%89%8DChrome%E7%80%8F%E8%A6%BD%E5%99%A8%E7%9A%84%E7%B6%B2%E5%9D%80&qs=n&form=QBRE&sp=-1&lq=0&pq=c%23+%E5%A6%82%E4%BD%95%E5%8F%96%E5%BE%97%E7%8F%BE%E5%89%8Dchrome%E7%80%8F%E8%A6%BD%E5%99%A8%E7%9A%84%E7%B6%B2%E5%9D%80&sc=6-21&sk=&cvid=1BA2FB0FBF4D48BE904A2209E4D9F85C&ghsh=0&ghacc=0&ghpl=
        /// </summary>
        /// <param name="controlType"></param>
        /// <returns></returns>
        static string getUrl(ControlType controlType)
        {
            string urls = "";
            try
            {
                //Process[] procsChrome = Process.GetProcessesByName("chrome");
                Process[] procsBrowser = Process.GetProcessesByName(browserName);
                if (procsBrowser.Length <= 0)
                {
                    //    MessageBox.Show("Chrome is not running");
                    MessageBox.Show(browserName + " " +
                        "is not the source running browser" + "\n" +
                        "來源流覽器");
                }
                else
                {
                    foreach (Process proc in procsBrowser)
                    {
                        // the chrome process must have a window
                        if (proc.MainWindowHandle == IntPtr.Zero)
                        {
                            continue;
                        }

                        // find the automation element
                        AutomationElement elm = AutomationElement.FromHandle
                            (proc.MainWindowHandle);
                        //AutomationElement elmUrlBar =
                        //    elm.FindFirst(TreeScope.Descendants,
                        //    new PropertyCondition(AutomationElement.NameProperty,
                        //    "Address and search bar"));

                        //AutomationElement elmUrlBar =
                        //    elm.FindFirst(TreeScope.Subtree,
                        //    new PropertyCondition(
                        //        AutomationElement.ControlTypeProperty,
                        //        controlType));
                        //AutomationElement elmUrlBar =
                        //    elm.FindFirst(TreeScope.Subtree,
                        //    new PropertyCondition(
                        //        AutomationElement.ControlTypeProperty,
                        //        controlType));

                        AutomationElementCollection elmUrlBar =
                            elm.FindAll(TreeScope.Subtree,
                            new PropertyCondition(
                                AutomationElement.ControlTypeProperty,
                                controlType));//https://social.msdn.microsoft.com/Forums/en-US/f9cb8d8a-ab6e-4551-8590-bda2c38a2994/retrieve-chrome-url-using-automation-element-in-c-application?forum=csharpgeneral

                        /*要用Edit屬性才抓得到網址列,Text也不行
                         */

                        // if it can be found, get the value from the URL bar
                        if (elmUrlBar != null)
                        {
                            int i = 0; int cnt = elmUrlBar.Count;
                        nx: foreach (AutomationElement Elm in elmUrlBar)
                            {
                                try
                                {
                                    i++; if (i > cnt) break;
                                    string vp = ((ValuePattern)Elm.
                                    GetCurrentPattern(ValuePattern.Pattern)).
                                    Current.Value as string;
                                    //if (urls.IndexOf(vp) == -1)
                                    if ((vp.StartsWith("http") || vp.StartsWith("ctext")) && urls.IndexOf(vp) == -1)
                                        urls += (vp + " ");
                                }
                                catch (Exception)
                                {
                                    goto nx;
                                    //throw;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //textBox2.Text = ex.ToString();
                MessageBox.Show(ex.ToString());
            }
            return urls;
        }

        internal static void SwitchToCurrentForeActivateTab(ref TextBox textBox3)
        {
            string url = "", urlActiveTab = "";
            try
            {
                url = driver.Url; urlActiveTab = ActiveTabURL;
            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case -2146233088:
                        if (ex.Message.IndexOf("no such window: target window already closed") > -1)//"no such window: target window already closed\nfrom unknown error: web view not found\n  (Session info: chrome=110.0.5481.178)"
                        {
                            driver.SwitchTo().Window(driver.WindowHandles[0]);
                            url = driver.Url; if (urlActiveTab == "") urlActiveTab = ActiveTabURL;
                        }
                        else
                            MessageBox.Show(ex.HResult + ex.Message);
                        break;
                    default:
                        MessageBox.Show(ex.HResult + ex.Message);
                        break;
                }
            }

            if (url != urlActiveTab)//如果現行的頁面不同於程式所在之頁面
                GoToCurrentUserActivateTab(urlActiveTab);//那麼就將程式所在之頁面轉到、設定為現行的前景頁面
            if (textBox3.Text != "" && textBox3.Text != urlActiveTab) textBox3.Text = urlActiveTab;//如果textBox3非空值且與現行頁面網址不合，亦轉設為現行前景頁面之網址}
        }
        /// <summary>
        /// 將程式所在頁面轉到現行前景的分頁頁面
        /// </summary>
        /// <param name="urlActiveTab">若已取得現行前景分頁頁面之網址則作此引數傳入，免得再取一次，徒耗資源，減損效能</param>
        internal static void GoToCurrentUserActivateTab(string urlActiveTab = "")
        {
            if (urlActiveTab == "") urlActiveTab = ActiveTabURL;
            string url = "";
            if (urlActiveTab != "")
            {
                try
                {
                    url = driver.Url;
                }
                catch (Exception ex)
                {
                    switch (ex.HResult)
                    {
                        case -2146233088:
                            if (ex.Message.IndexOf("no such window: target window already closed") > -1) //"no such window: target window already closed\nfrom unknown error: web view not found\n  (Session info: chrome=110.0.5481.178)"
                            {
                                driver.SwitchTo().Window(driver.WindowHandles[0]);
                            }
                            else
                            {
                                MessageBox.Show(ex.HResult + ex.Message);
                            }
                            break;
                        default:
                            MessageBox.Show(ex.HResult + ex.Message);
                            break;
                    }
                }
                if (urlActiveTab != url)
                {
                    foreach (var item in driver.WindowHandles)
                    {
                        url = driver.Url;
                        if (urlActiveTab == url) break;
                        driver.SwitchTo().Window(item);
                    }
                }
            }


        }

        internal static void GoToUrlandActivate(string url)
        {
            if (string.IsNullOrEmpty(url) || url.Substring(0, 4) != "http") return;

            ////driver.Close();//creedit
            ////creedit20230103 這樣處理誤關分頁頁籤的錯誤（例外情形）就成功了，但整個瀏覽器誤關則尚未
            ////chatGPT：在 C# 中使用 Selenium 取得 Chrome 瀏覽器開啟的頁籤（分頁）數量可以使用以下方法：                
            int tabCount = 0;
            try
            {
                if (driver == null) driver = driverNew();
            }
            catch (Exception)
            {
                if (driver != null)
                {
                    driver = null;
                }
                driver = driverNew();
                ////throw;
            }
            /*另外，您也可以使用以下方法在 C# 中取得 Chrome 瀏覽器的標籤（分頁）數量:
             // 取得 Chrome 瀏覽器的標籤數量
                int tabCount = driver.Manage().Window.Bounds.Width / 100;
             */
            try
            {
                driver = driver ?? Browser.driverNew();
                tabCount = driver.WindowHandles.Count;
            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case -2146233088://"The HTTP request to the remote WebDriver server for URL http://localhost:4144/session/a5d7705c0a6199c76529de0e157667f9/window/handles timed out after 8.5 seconds."
                        killProcesses(new string[] { "chromedriver" });//手動關閉由Selenium啟動的Chrome瀏覽器須由此才能清除
                        driver = null;
                        driver = driverNew();
                        break;
                    default:
                        throw;
                        //break;
                }
            }
            if (tabCount > 0)
            {
                //var hs = driver.WindowHandles;
                ////driver.SwitchTo().Window(hs[0]);
                try
                {
                    driver = driver ?? Browser.driverNew();
                    driver.SwitchTo().Window(driver.CurrentWindowHandle);
                }
                catch (Exception ex)
                {
                    switch (ex.HResult)
                    {
                        //操作中的分頁頁籤被手動誤關時
                        //no such window: target window already closed
                        case -2146233088:
                            openNewTabWindow();
                            break;
                        default:
                            throw;
                    }

                }
            }
            else
            {
                openNewTabWindow();
            }
            //throw;
            driver.Navigate().GoToUrl(url);
            //activate and move to most front of desktop
            //driver.SwitchTo().Window(driver.CurrentWindowHandle);
            driver.ExecuteScript("window.scrollTo(0, 0)");//chatGPT:您好！如果您使用 C# 和 Selenium 來控制 Chrome 瀏覽器，您可以使用以下的程式碼將網頁捲到最上面：
            quickedit_data_textboxSetting(url);
        }

        private static void quickedit_data_textboxSetting(string url, IWebElement textbox = null, IWebDriver driver = null)
        {
            if (url.IndexOf("edit") > -1)
            {
                if (textbox != null) Quickedit_data_textbox = textbox;
                else
                    Quickedit_data_textbox = waitFindWebElementByName_ToBeClickable("data", _webDriverWaitTimSpan, driver);
                quickedit_data_textboxTxt = Quickedit_data_textbox.Text;
            }
        }

        internal static ChromeDriver openNewTabWindow(WindowType tabOrwindow = WindowType.Tab)//creedit 20230103
        {/*chatGPT
            在 C# 中使用 Selenium 開啟新 Chrome 瀏覽器分頁可以使用以下方法：*/
            // 創建 ChromeDriver 實例
            //IWebDriver driver = new ChromeDriver();
            //ChromeDriver driver = driverNew();//new ChromeDriver();
            if (driver == null) driver = driverNew();
            try
            {
                driver.SwitchTo().NewWindow(tabOrwindow);

            }
            catch (Exception)
            {
                var hs = driver.WindowHandles;
                driver.SwitchTo().Window(driver.WindowHandles.Last());
                driver.SwitchTo().NewWindow(tabOrwindow);
                //throw;
            }

            //// 開啟新分頁
            //driver.ExecuteScript("window.open();");
            // 切換到新分頁
            //driver.SwitchTo().Window(driver.WindowHandles.Last());
            //也可以以用：（自己找switch可用的方法時發現的）
            //driver.SwitchTo().NewWindow(WindowType.Tab);   
            return driver;
        }

        static string stringtoEscape_sequences_for_Unicode_character_sets(string input)
        {

            string output = "";

            foreach (char c in input)
            {
                output += @"\u" + ((int)c).ToString("X4");
            }


            return output;

        }
        static string stringtoUTF_16_encoded_escape_sequences(string input)
        {

            string output = "";

            foreach (char c in input)
            {
                output += string.Format("\\x{0:X4}", (int)c);
            }

            return output;

        }

        internal static string GetImageUrl(string url = null)
        {//20230104 creedit
            if (Form1.browsrOPMode == Form1.BrowserOPMode.appActivateByName) Form1.browsrOPMode = Form1.BrowserOPMode.seleniumNew;
            if (driver == null) driver = driverNew();
            //using (driver)//var driver = new ChromeDriver())//若這樣寫則會出現「無法存取已處置的物件。」之錯誤    HResult	-2146232798	int               
            //{因為 using(driver) 這 driver 只在 ) 後的第一層大括弧{}間有效，生命週期僅止於此間而已
            // 移動到指定的網頁
            GoToUrlandActivate(url ?? System.Windows.Forms.Application.OpenForms[0].Controls["textBox3"].Text);
            //try
            //{
            //    driver.Navigate().GoToUrl(url ?? System.Windows.Forms.Application.OpenForms[0].Controls["textBox3"].Text);//("http://example.com/");
            //}
            //catch (Exception ex)
            //{
            //    switch (ex.HResult)
            //    {
            //        case -2146233088://no such window: target window already closed\nfrom unknown error: web view not found\n  (Session info: chrome=110.0.5481.100)

            //            //driver.Navigate().GoToUrl(url ?? System.Windows.Forms.Application.OpenForms[0].Controls["textBox3"].Text);//("http://example.com/");
            //            break;
            //        default:
            //            throw;
            //    }
            //}
            //driver.Navigate().GoToUrl(url ?? Form1.mainFromTextBox3Text);

            // 取得元件 scancont 的圖片網址
            //IWebElement scancont = driver.FindElement(By.Id("content"));
            //IWebElement scancont = driver.FindElement(By.Id("scancont"));
            IList<OpenQA.Selenium.IWebElement> imageElements = driver.FindElements(By.TagName("img"));
            string imageUrl = "";
            foreach (IWebElement imageElement in imageElements)
            {
                imageUrl = imageElement.GetAttribute("src");
                if (imageUrl.Substring(0, 26) == "https://library.ctext.org/"
                || (imageUrl.Substring(imageUrl.Length - 4, 4) == ".png"
                    && ((imageUrl.IndexOf(".cn_") > -1)
                    || imageUrl.IndexOf("dimage") > -1))) break;
                //Console.WriteLine(imageUrl);
            }
            //string imageUrl = imageElements.GetAttribute("src");

            return imageUrl;
            //}
        }
        /* 以下是我先寫來問chatGPT的，依其建議改如上
        internal static string getImageUrl() {

            Browser br = new Browser(System.Windows.Forms.Application.OpenForms[0] as Form1);
            ChromeDriver driver = new ChromeDriver();
            IWebElement scancont = driver.FindElement(By.Id("scancont"));
            return scancont.GetAttribute("src");

        }
        */

        #region Ctext 三種網頁模式判斷
        internal static bool isQuickEditUrl(string url)
        {
            if (url != "" && url.Substring(0, "https://ctext.org/".Length) == "https://ctext.org/" &&
                url.IndexOf("edit") > -1 &&
                    url.Substring(url.LastIndexOf("#editor")) == "#editor") return true;
            else
                return false;
        }
        internal static bool isEditChapterUrl(string url)
        {
            if (url != "" && url.Substring(0, "https://ctext.org/".Length) == "https://ctext.org/" &&
                    url.LastIndexOf("&action = editchapter") > -1) return true;

            else
                return false;
        }
        internal static bool isFilePageView(string url)
        {
            if (url != "" && url.Substring(0, "https://ctext.org/".Length) == "https://ctext.org/" &&
                    url.IndexOf("edit") == -1) return true;
            else
                return false;
        }
        #endregion

        /// <summary>
        /// 儲存chromedriver程序ID的陣列
        /// </summary>
        internal static List<int> chromedriversPID = new List<int>();
        ///// <summary>
        ///// 儲存chromedriver程序ID的陣列 chromedriversPID的下標值
        ///// </summary>
        //internal static int chromedriversPIDcntr = 0;

        /// <summary>
        /// 清除從這裡啟動的 chromedriver
        /// </summary>
        internal static void killchromedriverFromHere()
        {
            Process[] processInstances = Process.GetProcessesByName("chromedriver");
            foreach (var processInstance in processInstances)
            {
                try
                {
                    if (chromedriversPID.Contains(processInstance.Id))
                    {
                        processInstance.Kill();
                    }
                }
                catch (Exception)
                {
                    Task.WaitAny();
                    //throw;
                }
            }
            Task.WaitAll();
            chromedriversPID.Clear();
        }
        internal static Process[] getChromedrivers()
        {
            return Process.GetProcessesByName("chromedriver");
        }

        internal static void killProcesses(string[] processName)
        {
            foreach (var item in processName)
            {
                Process[] processInstances = Process.GetProcessesByName(item);
                foreach (var processInstance in processInstances)
                {
                    try
                    {
                        processInstance.Kill();

                    }
                    catch (Exception)
                    {
                        Task.WaitAny();
                        //throw;
                    }
                }
            }
            Task.WaitAll();
        }

        string getUrl(forms.Keys eKeyCode)
        {

            string url = new Form1().textBox3Text;
            if (url == "") return url;
            int edit = url.IndexOf("&editwiki");
            int page = 0; string urlSub = url;
            if (edit > -1)
            {
                urlSub = url.Substring(0, url.IndexOf("&page=") + "&page=".Length);
                page = Int32.Parse(
                    url.Substring(url.IndexOf("&page=") + "&page=".Length,
                    url.IndexOf("&editwiki=") - (url.IndexOf("&page=") + "&page=".Length)));
                if (eKeyCode == forms.Keys.PageDown)
                    url = urlSub + (page + 1).ToString() + url.Substring(url.IndexOf("&editwiki="));
                if (eKeyCode == forms.Keys.PageUp)
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
                if (eKeyCode == forms.Keys.PageDown)
                    url = urlSub + (page + 1).ToString();
                if (eKeyCode == forms.Keys.PageUp)
                    url = urlSub + (page - 1).ToString();
            }
            return url;
        }

        internal static void importBookmarks(ref ChromeDriver drive)//(ref ChromeDriver drive)
        {
            /*  chatGPT： 20230104
             您可以使用 ChromeDriver 和 ChromeOptions 類別來自動匯入書籤。
            第一步是在您的 C# 專案中安裝 Selenium.WebDriver NuGet 套件。 然後，您可以使用以下程式碼來設定 ChromeDriver 和 ChromeOptions：
            */
            //// 設定 ChromeDriver 並指定 ChromeDriver 可執行檔的路徑
            //IWebDriver driver = new ChromeDriver("path/to/chromedriver");

            // 建立 ChromeOptions 物件
            ChromeOptions options = new ChromeOptions();

            //設定書籤檔案的路徑
            options.AddUserProfilePreference("browser.bookmarks.file", @"x:\bookmarks_2023_1_3.html");

            //// 將 ChromeOptions 設定加入 ChromeDriver
            //driver = new ChromeDriver(options);

            options.AddArgument("--password-store=basic");

            // 將 ChromeOptions 設定加入 ChromeDriver

            //return options;
        }

        /// <summary>
        /// Google Keep OCR
        /// </summary>
        /// <param name="downloadImgFullName">書圖檔所在位置=全檔名</param>
        /// <returns></returns>
        internal static bool OCR_GoogleKeep(string downloadImgFullName)
        {
            driver = driver ?? driverNew();
            string currentWindowHndl = driver.CurrentWindowHandle;
            //const string keep = "https://keep.google.com/#NOTE/1XHzZWpH5DCDGOctKjMwNad9qGdtUiYQpSw7HtkmfuEEAJOCtlj37xJg5XgRzWoE";
            string keep = OCRSite_URL[OCRSiteTitle.GoogleKeep];//"https://keep.new";
            openNewTabWindow(WindowType.Window);
            driver.Navigate().GoToUrl(keep);
            //取得文字框
            IWebElement iwe = waitFindWebElementBySelector_ToBeClickable("body > div.VIpgJd-TUo6Hb.XKSfm-L9AdLc.eo9XGd > div > div.IZ65Hb-TBnied.zTETae-h1U9Be-hxXJme > div.IZ65Hb-s2gQvd > div.IZ65Hb-qJTHM-haAclf > div.notranslate.IZ65Hb-YPqjbf.fmcmS-x3Eknd.h1U9Be-YPqjbf");
            iwe.Click();
            //iwe.Clear();
            //按下：新增圖片
            waitFindWebElementBySelector_ToBeClickable("body > div.VIpgJd-TUo6Hb.XKSfm-L9AdLc.eo9XGd > div > div.IZ65Hb-TBnied.zTETae-h1U9Be-hxXJme > div.IZ65Hb-yePe5c > div.IZ65Hb-INgbqf > div.Q0hgme-LgbsSe.Q0hgme-Bz112c-LgbsSe.Ge5tnd-HiaYvf.INgbqf-LgbsSe.VIpgJd-LgbsSe", 6).Click();
            Thread.Sleep(1200);
            //輸入：檔案名稱
            //SendKeys.Send(downloadImgFullName);
            Clipboard.SetText(downloadImgFullName);
            SendKeys.Send("+{Insert}");//or "^v"
            SendKeys.Send("{ENTER}");
            FileInfo fileInfo = new FileInfo(downloadImgFullName);
            long fileSize = fileInfo.Length;
            /* creedit with YouChat菩薩 ：C# 取得檔案大小
                To get the size of a file in C#, you can use the FileInfo class in the System.IO namespace. Here is an example code snippet that demonstrates how to get the file size:……
                In this example, we first create a FileInfo object by passing in the file path to the constructor. We then use the Length property to get the size of the file in bytes. Finally, we print out the size of the file using Console.WriteLine().
                Note that fileSize is of type long as the Length property returns the size in bytes, which can be a very large number if the file is big.
                I hope this helps! Let me know if you have any further questions.
             */
            if (fileSize > 31000)
                Thread.Sleep(8000);
            else
                //待圖載入
                Thread.Sleep(6520);
            //按下：…更多
            IWebElement iwe_morebtn = waitFindWebElementBySelector_ToBeClickable("body > div.VIpgJd-TUo6Hb.XKSfm-L9AdLc.eo9XGd > div > div.IZ65Hb-TBnied.zTETae-h1U9Be-hxXJme > div.IZ65Hb-yePe5c > div.IZ65Hb-INgbqf > div.Q0hgme-LgbsSe.Q0hgme-Bz112c-LgbsSe.xl07Ob.INgbqf-LgbsSe.VIpgJd-LgbsSe");
            iwe_morebtn.Click();
            byte tryTimes = 1;//重做記數，防有例外情形、成無窮迴圈；如「按下：新增圖片」處可能還來不及待輸入框出現程式便兀自執行下去
        retry:
            //按下：擷取圖片文字
            //Thread.Sleep(300);
            IWebElement iwe_ocr =
            waitFindWebElementBySelector_ToBeClickable("#\\:8 > div", 0);
            if (iwe_ocr == null)
            {
                Thread.Sleep(90); iwe_morebtn.Click(); tryTimes++;
                if (tryTimes < 10) goto retry;
                else
                {
                    //MessageBox.Show("請重來一次；重新執行一次。感恩感恩　南無阿彌陀佛", "", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    return false;
                }
            }
            iwe_ocr.Click();
            Thread.Sleep(500);
            //將OCR結果複製到剪貼簿                    
            if (iwe.Text == "")
            {
                Thread.Sleep(800); iwe_morebtn.Click();
                goto retry;
            }
            Clipboard.SetText(iwe.Text);
            #region 刪除一則記事
            iwe_morebtn.Click();
            //IWebElement iwe_del =
            waitFindWebElementBySelector_ToBeClickable("#\\:1 > div").Click();
            //移除圖片

            //IWebElement iwe_pic=
            //waitFindWebElementBySelector_ToBeClickable("body > div.VIpgJd-TUo6Hb.XKSfm-L9AdLc.eo9XGd > div > div.IZ65Hb-TBnied.zTETae-h1U9Be-hxXJme > div.IZ65Hb-s2gQvd > div.IZ65Hb-hYUzqc > div > div > img");
            //Cursor.Position= new Point( (iwe_pic.Location.X+iwe_pic.Size.Width)/2,(iwe_pic.Location.Y+iwe_pic.Size.Height)/2);
            //waitFindWebElementBySelector_ToBeClickable("body > div.VIpgJd-TUo6Hb.XKSfm-L9AdLc.eo9XGd > div > div.IZ65Hb-TBnied.zTETae-h1U9Be-hxXJme > div.IZ65Hb-s2gQvd > div.IZ65Hb-hYUzqc > div > div > div",3).Click();

            //iwe.Click();
            //SendKeys.Send("^a{delete}");
            //清除
            //iwe.Clear();
            #endregion

            #region 關閉OCR視窗後回到原來分頁視窗
            driver.Close();
            driver.SwitchTo().Window(currentWindowHndl);
            #endregion
            return true;

        }/// <summary>
         /// 《古籍酷》OCR：自動識別(豎版)
         /// </summary>
         /// <param name="downloadImgFullName">書圖檔全檔名</param>
         /// <returns></returns>
        internal static bool OCR_GJcool_AutoRecognizeVertical(string downloadImgFullName)
        {
            #region 先檢查瀏覽器下載目錄並取得 ：
            // 注入 JavaScript 代碼以獲取下載目錄
            //string downloadDirectory = (string)driver.ExecuteScript("return window.navigator.userAgent.toLowerCase().indexOf('win') > -1 ? window.localStorage.getItem('download.default_directory') : null;");
            string downloadDirectory = DownloadDirectory_Chrome;
            //string downloadDirectory=null;//= getChromeDownloadDirectory();

            if (string.IsNullOrEmpty(downloadDirectory) || !Directory.Exists(downloadDirectory))
            {
                downloadDirectory = getChromeDownloadDirectory_YouChatchatGPT();
                if (string.IsNullOrEmpty(downloadDirectory) || !Directory.Exists(downloadDirectory))
                {
                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請先指定Chrome瀏覽器的下載目錄，再繼續！感恩感恩　南無阿彌陀佛");
                    return false;
                }
                else//更新downloadDirectory_Chrome 欄位（若有異動時）
                    DownloadDirectory_Chrome = downloadDirectory;
            }
            #endregion

            driver = driver ?? driverNew();
            string currentWindowHndl = driver.CurrentWindowHandle;
            string gjCool = OCRSite_URL[OCRSiteTitle.GJcool]; //"https://gj.cool/try_ocr";
            openNewTabWindow(WindowType.Window);
            driver.Navigate().GoToUrl(gjCool);
            //按下：新增圖片：選擇檔案
            //Thread.Sleep(3200);
            SendKeys.Send("{tab 16} ");
            //waitFindWebElementBySelector_ToBeClickable("#line_img_form > div > input[type=file]").SendKeys(OpenQA.Selenium.Keys.Space);
            //waitFindWebElementByName_ToBeClickable("line_img",2).Submit();
            Thread.Sleep(1200);
            //輸入：檔案名稱 //SendKeys.Send(downloadImgFullName);
            Clipboard.SetText(downloadImgFullName);
            byte tryTimes = 1;
        retry:
            SendKeys.Send("+{Insert}");//or "^v"
            SendKeys.Send("{ENTER}");
            //待圖載入
            Thread.Sleep(3220);
            //按下「Pro」
            IWebElement iwe = waitFindWebElementBySelector_ToBeClickable("#line_img_form > div > input[type=file]");
            if (iwe == null)
            {
                tryTimes++;
                if (tryTimes > 5) return false;
                goto retry;
            }
            //iwe.Click();
            SendKeys.Send("{tab}~");
            //按下「自動識別(豎版)」，開始OCR……
            SendKeys.Send("{down}~");
            //iwe = waitFindWebElementBySelector_ToBeClickable("# OneLine > div.d-flex.justify-content-between.mt-2.mb-1 > div:nth-child(1) > div:nth-child(2) > ul > li:nth-child(2) > button");
            //iwe.Click();
            //Thread.Sleep(6220);
            Thread.Sleep(6000);//OCR結束
            SendKeys.Send("{tab 24}~");
            //按下「導出數據」：
            //iwe = waitFindWebElementBySelector_ToBeClickable("# outputDropdown");
            //iwe.Click();
            SendKeys.Send("{tab 5} ");
            //按下：「坐標」核取方塊，取消之
            //iwe = waitFindWebElementBySelector_ToBeClickable("# outputSelect1");
            //iwe.Click();
            SendKeys.Send("{tab} ");
            //按下：「識別概率」核取方塊，取消之
            //iwe = waitFindWebElementBySelector_ToBeClickable("#outputSelect2");
            //iwe.Click();
            SendKeys.Send("+{tab 4}~");
            //按下：「TXT」選項，匯出成純文字檔
            //iwe = waitFindWebElementBySelector_ToBeClickable("#OneLine > div.d-flex.justify-content-between.mt-2.mb-1 > div.d-flex.justify-content-end.input-group > div:nth-child(3) > ul > li:nth-child(2) > a");
            //iwe.Click();
            Thread.Sleep(520);
            //以上測試OK 20230310：01:26
            #region 將OCR結果讀入剪貼簿 creedit with chatGPT大菩薩：


            #region 讀入文本 creedit with chatGPT大菩薩：
            //取得所匯出的檔案路徑
            string filePath = Path.Combine(downloadDirectory, Path.GetFileNameWithoutExtension(downloadImgFullName) + ".txt"); //@"X:\Ctext_Page_Image.txt";            
            //等待下載完成
            TimeSpan waitFileExitTimeSpan = TimeSpan.FromSeconds(10);//最多等待時間
            DateTime waitFileExitBegin = DateTime.Now;
            while (!File.Exists(filePath) && DateTime.Now.Subtract(waitFileExitBegin) < waitFileExitTimeSpan)
            {
            }
            if (File.Exists(filePath))
            {
                string text = File.ReadAllText(filePath, System.Text.Encoding.UTF8);
                Clipboard.SetText(text);
                //刪除下載檔案，以便下次載入
                Task.Run(() => File.Delete(filePath));
            }
            else return false;
            #endregion
            #endregion

            #region 關閉OCR視窗後回到原來分頁視窗
            driver.Close();
            driver.SwitchTo().Window(currentWindowHndl);
            #endregion
            return true;
        }



        /// <summary>
        /// 取得Chrome瀏覽器的下載目錄（失敗！抓不到！！）
        /// 20230310：0133 creedit with YouChat菩薩
        /// </summary>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        private static string getChromeDownloadDirectory()
        {

            //ChromeOptions options = chromeOptions(chrome_path);
            //options.AddArgument("--disable-web-security");
            //options.AddArgument("--allow-running-insecure-content");
            ////options.AddArgument("--headless");
            //setupChromeDriverService();            
            //ChromeDriver driver = new ChromeDriver(driverService,options);            
            //string currentWindowsHandle = driver.CurrentWindowHandle;
            //driver.Navigate().GoToUrl("chrome://settings/downloads");
            ////IWebElement element = waitFindWebElementBySelector_ToBeClickable("#defaultDownloadPath");
            ////IWebElement element = driver.FindElement(By.CssSelector("#defaultDownloadPath"));
            //var elements = driver.FindElements(By.TagName("secondary"));
            //IWebElement element = driver.FindElement(By.Id("defaultDownloadPath"));
            ////< div class="secondary" id="defaultDownloadPath" aria-hidden="true">

            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(2));
            //wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(element));
            //string downloadDirectory = element.GetAttribute("innerText");
            //driver.SwitchTo().Window(currentWindowsHandle);
            //downloadDirectory_Chrome = downloadDirectory;
            //driver.Quit();
            //return downloadDirectory;
            return "";
            //string currentWindowsHanl = driver.CurrentWindowHandle;
            //driver.Navigate().GoToUrl("chrome://settings/downloads");
            //IWebElement e = waitFindWebElementBySelector_ToBeClickable("#defaultDownloadPath");
            //downloadDirectory_Chrome= e.GetDomProperty("innerText");//X:\
            //return downloadDirectory_Chrome;
        }

        /// <summary>
        /// 取得Chrome瀏覽器的下載目錄 YouChat菩薩的成功了
        /// </summary>
        /// <returns></returns>
        private static string getChromeDownloadDirectory_YouChatchatGPT()
        {
            #region YouChat菩薩的成功了
            /*用 C# 和 Selenium 可以取得Chrome瀏覽器的下載目錄嗎
                是的，你可以使用 C# 和 Selenium 来获取 Chrome 浏览器的下载目录。你需要使用 ChromeOptions 类来设置 ChromeDriver，然后使用 Options 属性的 AddUserProfilePreference 方法来设置下载目录。以下是一个示例代码片段：

                using OpenQA.Selenium;
                using OpenQA.Selenium.Chrome;

                namespace SeleniumDemo
                {
                    class Program
                    {
                        static void Main(string[] args)
                        {
                            ChromeOptions options = new ChromeOptions();
                            options.AddUserProfilePreference("download.default_directory", "C:\\Users\\yourusername\\Downloads");
                            IWebDriver driver = new ChromeDriver(options);
                            // rest of your code
                        }
                    }
                }
                在这个示例中，我们使用 ChromeOptions 类来设置 ChromeDriver，将下载目录设置为 "C:\Users\yourusername\Downloads"。然后我们创建了一个 ChromeDriver 实例，并将 options 参数传递给它。现在，每当您使用该驱动程序下载文件时，它将被下载到指定的目录。

                我是要取得，不是設定
                非常抱歉，我的前一个回答有误。要获取 Chrome 浏览器的下载目录，您需要打开 Chrome 首选项文件，然后查找下载目录。以下是一个示例代码片段：
             */
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string chromeUserDir = Path.Combine(userProfile, "AppData", "Local", "Google", "Chrome", "User Data", "Default");
            if (!Directory.Exists(chromeUserDir))
                chromeUserDir = user_data_dir.Substring(user_data_dir.IndexOf("--user_data_dir=") + "--user_data_dir=".Length + 1) + "default";// user_data_dir=options.Arguments[0];
            string preferencesFile = Path.Combine(chromeUserDir, "Preferences");

            if (File.Exists(preferencesFile))
            {

                string preferencesText = File.ReadAllText(preferencesFile);
                JObject preferencesJson = JObject.Parse(preferencesText);
                downloadDirectory_Chrome = preferencesJson["download"]["default_directory"].ToString();
                /* YouChat大菩薩：在这个示例中，我们首先获取当前用户的主文件夹，然后使用 Path.Combine 方法创建一个路径字符串，以便打开 Chrome 用户数据目录。然后，我们查找名为 Preferences 的文件，它包含 Chrome 首选项的 JSON 格式。我们使用 JObject.Parse 方法将 Preferences 文件的内容解析为 JSON 格式，然后查找下载目录的默认路径。最后，我们输出下载目录的值。 请注意，下载目录设置可能存在多个偏好文件，因此您可能需要查找适用于您的系统和 Chrome 版本的正确偏好文件。此外，这种方法依赖于 Chrome 的偏好文件格式，因此可能会因 Chrome 的更新而改变。 如果您想在下载文件时将文件保存到自定义目录中，建议使用我在前一个回答中提供的示例代码。
                 * 用這個 JObject 物件有什麼先決條件吧 必須裝載什麼組件或套件呢 YouChat大菩薩：
                    JObject is a class in the Newtonsoft.Json.Linq namespace that is used to represent a JSON object in C#.

                    To use JObject in your C# code, you first need to make sure that the Newtonsoft.Json NuGet package is installed in your project. You can install it using the NuGet Package Manager in Visual Studio, or by running the following command in the Package Manager Console:

                    Install-Package Newtonsoft.Json
                    Once you have the Newtonsoft.Json package installed, you can create a new JObject object like this:

                    using Newtonsoft.Json.Linq;

                    JObject obj = new JObject();
                    You can also parse a JSON string into a JObject like this:

                    string jsonString = "{\"name\":\"John Smith\",\"age\":42}";
                    JObject obj = JObject.Parse(jsonString);
                    Note that you need to include the using Newtonsoft.Json.Linq; directive at the top of your C# file to access the JObject class.
                 */

                #endregion

                #region chatGPT菩薩的有誤
                //string preferencesText = File.ReadAllText(preferencesFile);
                //dynamic preferencesJson = JsonConvert.DeserializeObject(preferencesText);
                //downloadDirectory = preferencesJson["download"]["default_directory"];
                //string preferencesText = File.ReadAllText(preferencesFile);


                //using (JsonDocument document = JsonDocument.Parse(preferencesText))
                //{
                //    JsonElement root = document.RootElement;
                //    if (root.TryGetProperty("download", out JsonElement download))
                //    {
                //        if (download.TryGetProperty("default_directory", out JsonElement defaultDirectory))
                //        {
                //            DownloadDirectory_Chrome = defaultDirectory.GetString();
                //        }
                //    }
                //}
                #endregion

            }

            //Console.WriteLine("Chrome download directory: " + downloadDirectory);
            return downloadDirectory_Chrome;
        }
    }
}

