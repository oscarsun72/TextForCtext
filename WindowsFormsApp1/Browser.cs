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
using static System.Net.Mime.MediaTypeNames;
using System.Security.Policy;
using System.Drawing;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.DevTools.V85.ApplicationCache;
using System.ComponentModel;
using System.Runtime.CompilerServices;
//https://dotblogs.com.tw/supergary/2020/10/29/selenium#images-3
using System.IO;
using System.Net;
using static System.Net.WebRequestMethods;
using System.Runtime.InteropServices.ComTypes;

namespace TextForCtext
{
    class Browser
    {
        static ChromeDriverService driverService = ChromeDriverService.CreateDefaultService();        

        // 創建Chrome驅動程序對象
        //selm.IWebDriver driver=driverNew();        
        //internal static selm.IWebDriver driver=driverNew();
        internal static ChromeDriver driver = driverNew();
        //static selm.IWebDriver driverNew()
        //實測後發現：CurrentWindowHandle並不能取得瀏覽器現正作用中的分頁視窗，只能取得創建 ChromeDriver 物件時的最初及switch 方法執行後切換的分頁視窗 20230103 阿彌陀佛
        static string originalWindow = driver.CurrentWindowHandle;

        internal static string getOriginalWindow
        {
            get
            {
                return originalWindow;
            }
        }

        static Form1 frm;//creedit 
        public Browser(Form1 form)
        {
            frm = form;
        }



        static ChromeDriver driverNew()
        {
            if (driver == null)
            {
                driverService.HideCommandPromptWindow = true;//关闭黑色cmd窗口
                /*ChromeDriver cDrv;*///綠色免安裝版仍搞不定，安裝 chrome 了就OK 20220101 chatGPT建議者未通



                ChromeOptions options = chromeOptions();
                // 將 ChromeOptions 設定加入 ChromeDriver
                ChromeDriver cDrv = new ChromeDriver(driverService, options);
                //ChromeDriver cDrv = new ChromeDriver("C:\\Users\\oscar\\.cache\\selenium\\chromedriver\\win32\\108.0.5359.71\\chromedriver.exe", options);
                //cDrv = new ChromeDriver(@"C:\Program Files\Google\Chrome\Application\chrome.exe",options);
                //cDrv = new ChromeDriver(@"x:\chromedriver.exe", options);
                //上述加入書籤並不管用！！！20230104

                //string chrome_path = Form1.getDefaultBrowserEXE();
                //if (chrome_path.IndexOf(@"C:\") == -1)
                //{
                //try
                //{//selenium 如何操作免安裝版的 chrome 瀏覽器 或自訂安裝路徑的 chrome 瀏覽器呢
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
                cDrv.Navigate().GoToUrl(Form1.mainFromTextBox3Text ?? "https://ctext.org/account.pl?if=en");
                //IWebElement clk = cDrv.FindElement(selm.By.Id("logininfo")); clk.Click();
                //cDrv.FindElement(selm.By.Id("logininfo")).Click();
                /*202301050214 因為以下這行設定成功，可以用平常的Chrome來操作了，就不必再登入安裝（如擴充功能）匯入（如書籤）什麼的了 感恩感恩　讚歎讚歎　南無阿彌陀佛
                 options.AddArgument("--user-data-dir=C:\\Users\\oscar\\AppData\\Local\\Google\\Chrome\\User Data\\");
                options.AddArgument("--user-data-dir="+ Environment.GetFolderPath( Environment.SpecialFolder.LocalApplicationData) +"\\Google\\Chrome\\User Data\\");
                 */
                //MessageBox.Show("請先登入 Ctext.org 再繼續。按下「確定(OK)」以繼續……");

                return cDrv;
            }
            else
                return driver;
        }

        private static ChromeOptions chromeOptions()
        {
            // 建立 ChromeOptions 物件            
            ChromeOptions options = new ChromeOptions();

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
            #region it worked！！ ：D
            //202301050205終於成了 這可以用原來的chrome（即使用者啟動操作慣用的一切設定，如書籤、擴充功能等等）而不是空白的、原始的來操作了 https://www.cnblogs.com/baihuitestsoftware/articles/7742069.html            
            //options.AddArgument("--user-data-dir=C:\\Users\\oscar\\AppData\\Local\\Google\\Chrome\\User Data\\");
            //有沒有「--」（--user or user）都可
            options.AddArgument("user-data-dir=" + Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\Google\\Chrome\\User Data\\");
            //https://www.cnblogs.com/hushaojun/p/5981646.html

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
            return options;
        }

        internal static void 在Chrome瀏覽器的Quick_edit文字框中輸入文字(ChromeDriver driver, string xIuput, string url)//在Chrome瀏覽器的文字框中輸入文字,creedit
        {

            if (url.IndexOf("edit") == -1) return;
            //selm.IWebDriver driver = new ChromeDriver();

            if (url != driver.Url && driver.Url.IndexOf(url.Replace("editor", "box")) == -1)
                // 使用driver導航到給定的URL
                driver.Navigate().GoToUrl(url);
            //("https://ctext.org/library.pl?if=en&file=79166&page=85&editwiki=297821#editor");//("http://www.example.com");

            // 查找名稱為"textbox"的文字框元素
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
                quickedit.Click();//預設當如下面「submit.Click();」會等網頁作出回應才執行下一步。感恩感恩　讚歎讚歎　南無阿彌陀佛
                textbox = driver.FindElement(selm.By.Name("data"));
                //throw;
            }

            textbox.Clear();
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
            #region paste to textbox
            //文字框取得焦點
            textbox.Click();
            //chrome取得焦點
            //Form1 f = new Form1();
            //f.appActivateByName();
            driver.SwitchTo().Window(driver.CurrentWindowHandle); //https://stackoverflow.com/questions/23200168/how-to-bring-selenium-browser-to-the-front#_=_
                                                                  // 讓 Chrome 瀏覽器成為作用中的程式
                                                                  //driver.Manage().Window.Maximize();//creedit chatGPT
                                                                  //driver.Manage().Window.Position = new Point(0, 0);

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
            selm.IWebElement submit = driver.FindElement(selm.By.Id("savechangesbutton"));//("textbox"));
            /* creedit 我問：在C#  用selenium 控制 chrome 瀏覽器時，怎麼樣才能不必等待網頁作出回應即續編處理按下來的程式碼 。如，以下程式碼，請問，如何在按下 submit.Click(); 後不必等這個動作完成或作出回應，即能繼續執行之後的程式碼呢 感恩感恩　南無阿彌陀佛
                        chatGPT他答：你可以將 submit.Click(); 放在一個 Task 中去執行，並立即返回。
             */
            Task.Run(() =>
            {
                submit.Click();
            });
            //f = null;


            //// 關閉瀏覽器
            //driver.Close();
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
            try
            {
                driver.SwitchTo().Window(driver.CurrentWindowHandle);

            }
            catch (Exception)
            {
                //driver.Close();//creedit
                //creedit20230103 這樣處理誤關分頁頁籤的錯誤（例外情形）就成功了，但整個瀏覽器誤關則尚未
                //chatGPT：在 C# 中使用 Selenium 取得 Chrome 瀏覽器開啟的頁籤（分頁）數量可以使用以下方法：
                int tabCount = driver.WindowHandles.Count;
                /*另外，您也可以使用以下方法在 C# 中取得 Chrome 瀏覽器的標籤（分頁）數量:
                 // 取得 Chrome 瀏覽器的標籤數量
                    int tabCount = driver.Manage().Window.Bounds.Width / 100;
                 */
                if (tabCount > 0)
                {
                    var hs = driver.WindowHandles;
                    //driver.SwitchTo().Window(hs[0]);
                    driver.SwitchTo().Window(driver.CurrentWindowHandle);
                }
                else
                {
                    Browser.openNewTab();
                }
                //throw;
            }
            driver.Navigate().GoToUrl(url);
            //activate and move to most front of desktop
            driver.SwitchTo().Window(driver.CurrentWindowHandle);
            driver.ExecuteScript("window.scrollTo(0, 0)");//chatGPT:您好！如果您使用 C# 和 Selenium 來控制 Chrome 瀏覽器，您可以使用以下的程式碼將網頁捲到最上面：
        }


        internal static ChromeDriver openNewTab()//creedit 20230103
        {/*chatGPT
            在 C# 中使用 Selenium 開啟新 Chrome 瀏覽器分頁可以使用以下方法：*/
            // 創建 ChromeDriver 實例
            //IWebDriver driver = new ChromeDriver();
            ChromeDriver driver = new ChromeDriver();

            // 開啟新分頁
            driver.ExecuteScript("window.open();");
            // 切換到新分頁
            driver.SwitchTo().Window(driver.WindowHandles.Last());
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
            using (var driver = new ChromeDriver())
            {
                // 移動到指定的網頁
                driver.Navigate().GoToUrl(url ?? System.Windows.Forms.Application.OpenForms[0].Controls["textBox3"].Text);//("http://example.com/");

                // 取得元件 scancont 的圖片網址
                //IWebElement scancont = driver.FindElement(By.Id("content"));
                //IWebElement scancont = driver.FindElement(By.Id("scancont"));
                IList<OpenQA.Selenium.IWebElement> imageElements = driver.FindElements(By.TagName("img"));
                string imageUrl = "";
                foreach (IWebElement imageElement in imageElements)
                {
                    imageUrl = imageElement.GetAttribute("src");
                    if (imageUrl.Substring(imageUrl.Length - 4, 4) == ".png"
                        && ((imageUrl.IndexOf(".cn_") > -1)
                        || imageUrl.IndexOf("dimage") > -1)) break;
                    //Console.WriteLine(imageUrl);
                }
                //string imageUrl = imageElements.GetAttribute("src");

                return imageUrl;
            }
        }
        /* 以下是我先寫來問chatGPT的，依其建議改如上
        internal static string getImageUrl() {

            Browser br = new Browser(System.Windows.Forms.Application.OpenForms[0] as Form1);
            ChromeDriver driver = new ChromeDriver();
            IWebElement scancont = driver.FindElement(By.Id("scancont"));
            return scancont.GetAttribute("src");

        }
        */

        internal static void downloadImage(string imageUrl)
        {/*20230103 creedit,chatGPT：
          你可以使用 Selenium 來下載網絡圖片。
            首先，你需要獲取圖片的 URL。然後，使用 WebClient 的 DownloadData 方法下載圖片的二進制數據。
            最後，使用 FileStream 將二進制數據寫入文件即可。  
          */
            // 獲取圖片的 URL。
            //imageUrl = "https://example.com/image.png";

            // 使用 WebClient 下載圖片的二進制數據。
            WebClient webClient = new WebClient();
            byte[] imageBytes = webClient.DownloadData(imageUrl);

            // 將二進制數據寫入文件。
            using (FileStream fileStream = new FileStream(@"x:\Ctext_Page_Image.png", FileMode.Create))
            {
                fileStream.Write(imageBytes, 0, imageBytes.Length);
                Console.WriteLine("圖片已成功下載。");//在「即時運算視窗」寫出訊息
            }
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

