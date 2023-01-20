using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using selm = OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using forms = System.Windows.Forms;
using WindowsFormsApp1;
using System.Windows.Forms;
using System.Drawing.Imaging;
//using static System.Net.Mime.MediaTypeNames;
using System.Security.Policy;
using System.Drawing;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.DevTools.V85.ApplicationCache;
using System.ComponentModel;
using System.Runtime.CompilerServices;
//https://dotblogs.com.tw/supergary/2020/10/29/selenium#images-3
using System.IO;
//using System.Net;
//using static System.Net.WebRequestMethods;
using System.Runtime.InteropServices.ComTypes;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System.Diagnostics;

namespace TextForCtext
{
    class Browser
    {
        static Form1 frm;

        //creedit 
        public Browser(Form1 form)
        {
            frm = form;
        }

        // 創建Chrome驅動程序對象
        //selm.IWebDriver driver=driverNew();        
        //internal static selm.IWebDriver driver=driverNew();
        internal static ChromeDriver driver = driverNew();
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
                return driver != null ? driver.Url : "";
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
        internal static IWebElement waitFindWebElementByNameToBeClickable(string name, float second,
            IWebDriver drver = null)
        {

            IWebElement e = (driver ?? drver).FindElement(By.Name(name));
            WebDriverWait wait = new WebDriverWait((driver ?? drver), TimeSpan.FromSeconds(second));
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(e));
            return e;
        }
        internal static IWebElement waitFindWebElementByIdToBeClickable(string id, float second)
        {
            IWebElement e = driver.FindElement(By.Id(id));
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(second));
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(e));
            return e;
        }

        internal static ChromeDriver driverNew()
        {
            if (Form1.browsrOPMode != Form1.BrowserOPMode.appActivateByName && driver == null)
            {
                string chrome_path = Form1.getDefaultBrowserEXE();

                // 將 ChromeOptions 設定加入 ChromeDriver
                ChromeOptions options = chromeOptions(chrome_path);//加入參數的順序重要，要參考「string user_data_dir = options.Arguments[0];」
                                                                   //ChromeDriver cDrv = new ChromeDriver("C:\\Users\\oscar\\.cache\\selenium\\chromedriver\\win32\\108.0.5359.71\\chromedriver.exe", options);
                                                                   //cDrv = new ChromeDriver(@"C:\Program Files\Google\Chrome\Application\chrome.exe",options);
                                                                   //cDrv = new ChromeDriver(@"x:\chromedriver.exe", options);
                                                                   //上述加入書籤並不管用！！！20230104//解法已詳下chromeOptions()中

            tryagain:
                ChromeDriverService driverService;
                ChromeDriver cDrv;//綠色免安裝版仍搞不定，安裝 chrome 了就OK 20220101 chatGPT建議者未通；20220105自行解決了，詳下



                string user_data_dir = options.Arguments[0];
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
                #region 啟動Chrome瀏覽器 （最會出錯的部分！！）
                try
                {
                    if (user_data_dir.IndexOf("Documents") > -1)//無寫入權限的電腦，怕比較慢
                                                                //可能是防火牆 OpenQA.Selenium.WebDriverException
                                                                //HResult = 0x80131500
                                                                //Message = The HTTP request to the remote WebDriver server for URL http://localhost:52966/session timed out after 60 seconds.
                        cDrv = new ChromeDriver(driverService, options);
                    else
                        //自己的電腦比較快
                        cDrv = new ChromeDriver(driverService, options, TimeSpan.FromSeconds(8.5));//等待重啟時間=8.5秒鐘：其實也是等待伺服器回應的時間，太短則在完整編輯（如網址有「&action=editchapter」）送出時，會逾時
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
                    frm.Controls["textBox1"].Text = waitFindWebElementByNameToBeClickable("data", 3).Text;
                }


                return cDrv;
            }
            else
                return driver;
            #endregion
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

        internal static readonly string chkClearQuickedit_data_textboxTxtStr = " ";// "\t"（其實是有由tab鍵所按下的值，或其他亂碼字），此與 Word VBA 中國哲學書電子化計劃.新頁面 為速新章節單位的配置有關 碼詳：https://github.com/oscarsun72/TextForCtext/blob/f75b5da5a5e6eca69baaae0b98ed2d6c286a3aab/WordVBA/%E4%B8%AD%E5%9C%8B%E5%93%B2%E5%AD%B8%E6%9B%B8%E9%9B%BB%E5%AD%90%E5%8C%96%E8%A8%88%E5%8A%83.bas#L32
        //在Chrome瀏覽器的文字框(ctext.org 的 Quick edit ）中輸入文字,creedit//若 xIuput= " "則清除而不輸入
        internal static void 在Chrome瀏覽器的Quick_edit文字框中輸入文字(ChromeDriver driver, string xIuput, string url)
        {
            #region 檢查網址
            if (url.IndexOf("edit") == -1) return;

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
            driver.SwitchTo().Window(driver.CurrentWindowHandle); //https://stackoverflow.com/questions/23200168/how-to-bring-selenium-browser-to-the-front#_=_
                                                                  // 讓 Chrome 瀏覽器成為作用中的程式
                                                                  //driver.Manage().Window.Maximize();//creedit chatGPT
                                                                  //driver.Manage().Window.Position = new Point(0, 0);

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
            selm.IWebElement submit = waitFindWebElementByIdToBeClickable("savechangesbutton", 3);
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
                    submit = waitFindWebElementByIdToBeClickable("savechangesbutton", 3);  //driver.FindElement(selm.By.Id("savechangesbutton"));
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
                            openNewTab();
                            break;
                        default:
                            throw;
                    }

                }
            }
            else
            {
                openNewTab();
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
                    Quickedit_data_textbox = waitFindWebElementByNameToBeClickable("data", 2, driver);
                quickedit_data_textboxTxt = Quickedit_data_textbox.Text;
            }
        }

        internal static ChromeDriver openNewTab()//creedit 20230103
        {/*chatGPT
            在 C# 中使用 Selenium 開啟新 Chrome 瀏覽器分頁可以使用以下方法：*/
            // 創建 ChromeDriver 實例
            //IWebDriver driver = new ChromeDriver();
            //ChromeDriver driver = driverNew();//new ChromeDriver();
            if (driver == null) driver = driverNew();
            try
            {
                driver.SwitchTo().NewWindow(WindowType.Tab);

            }
            catch (Exception)
            {
                var hs = driver.WindowHandles;
                driver.SwitchTo().Window(driver.WindowHandles.Last());
                driver.SwitchTo().NewWindow(WindowType.Tab);
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
            driver.Navigate().GoToUrl(url ?? System.Windows.Forms.Application.OpenForms[0].Controls["textBox3"].Text);//("http://example.com/");
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

    }
}

