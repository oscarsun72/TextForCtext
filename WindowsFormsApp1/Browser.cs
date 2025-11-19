//using static System.Net.Mime.MediaTypeNames;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
//https://dotblogs.com.tw/supergary/2020/10/29/selenium#images-3
using System.IO;
using System.Linq;
using System.Media;
using System.Resources;


//using System.Net;
//using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Automation;
using System.Windows.Forms;
using WebSocketSharp;
using WindowsFormsApp1;



//using static System.Net.Mime.MediaTypeNames;
using forms = System.Windows.Forms;
using selm = OpenQA.Selenium;


namespace TextForCtext
{
    class Browser
    {
        /// <summary>
        /// 作為模擬「參數」供給此類別操作時參考
        /// </summary>
        private static Form1 frm;

        /// <summary>
        /// 指定延長等待開啟舊檔對話方塊出現的時間（毫秒數）
        /// 在 textBox2 中輸入「ws」（wait second）以設定此值
        /// 若要縮減時間，請指定負數
        /// </summary>
        internal static int Extend_the_wait_time_for_the_Open_Old_File_dialog_box_to_appear_Millisecond = 0;
        /// <summary>
        /// 指定等待OCR諸過程最久的時間（以秒數）
        /// 在 textBox2 中輸入「wO」（wait second）以設定此值
        /// </summary>
        internal static int OCR_wait_time_Top_Limit＿second = 15;

        /// <summary>
        /// 記錄當前對外IP位置
        /// </summary>
        internal static string CurrentIP;
        /// <summary>
        /// 記錄切換VPN的次數
        /// </summary>
        internal static int VPNSwitchedTimer = 0;
        /// <summary>
        /// 因為剪貼簿的操作頻繁，故以此記下要貼去簡單修改模式方塊中的文字內容，以免被干擾
        /// </summary>
        internal static string TextPatst2Quick_editBox = "";


        //readonly Form1 Form1 = Application.OpenForms.Count > 0 ? Application.OpenForms[0] as Form1 : null;
        /*如何取得作用中的表單？ 20230312 chatGPT大菩薩：
        您好！如果您使用的是 Windows.Forms，您可以使用以下方法来获取当前处于活动状态的表单：
        Form activeForm = Form.ActiveForm;
        这将返回当前活动的表单对象，您可以使用该对象的属性和方法来访问和操作该表单。
        请注意，此方法仅适用于您的应用程序具有焦点，并且该方法将返回具有焦点的表单。如果没有焦点的表单，该方法将返回 null。
        如果您想要获取应用程序中的其他表单而不是当前处于活动状态的表单，您可以使用以下方法之一：
        使用表单的名称或类型来查找表单：
        Form myForm = Application.OpenForms["MyFormName"];
        或
        Form myForm = Application.OpenForms.OfType<MyFormType>().FirstOrDefault();
        在您的代码中保留对表单对象的引用，并在需要时使用该引用来访问该表单。
        */

        //creedit 
        public Browser(Form1 form)
        {
            ActiveForm1 = form;
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
                { OCRSiteTitle.GJcool, "https://gj.cool/try_ocr" }, // maps the key "GJcool" to the value "https://gj.cool/try_ocr"
                //{OCRSiteTitle.GJcool, "https://ocr.gj.cool/try_ocr" } // maps the key "GJcool" to the value "https://gj.cool/try_ocr" 
                //"https://ocr.gj.cool/try_ocr" 這個 oscarsun72 此帳戶可以登錄，而 "https://gj.cool/try_ocr" 似不行 20240208
                { OCRSiteTitle.KanDianGuJi, "https://kandianguji.com/ocr" } // maps the key "KanDianGuJi" to the value "https://kandianguji.com/ocr"
            };
        /// <summary>
        /// 儲存常用的網站名
        /// </summary>
        internal enum OCRSiteTitle { GoogleKeep, GJcool, KanDianGuJi, KanDianGuJiAPI }
        /* chatGPT大菩薩：C# Enum Values：
         你的程式碼是可以運作的。這樣的實作方式會使得程式碼更具有可維護性和可擴展性。在需要使用網址時，只需要通過enum來訪問對應的網址，而不需要直接使用網址字符串。當需要添加、修改或刪除網址時，只需要更新Dictionary中的對應鍵值對即可，而不需要修改程式碼中的enum。
        孫守真
            老師菩薩您給我這樣的實作幾分呢？
        根據我對你的程式碼所做的評估，我會給予你 95 分。這是因為你使用了 enum 和 Dictionary，使得你的程式碼更加有結構，可讀性和可維護性更高。但是，可能還有一些改進的空間，比如在程式碼中加入註釋說明，這樣可以讓其他人更容易理解你的程式碼的運作方式。
        這樣的實作方式是可行的，也符合基本的程式碼風格和規範。在這個程式碼片段中，你使用了兩個內部(static)的成員來儲存常用的網站名及其網址。你的程式碼包含了一個字典，用來將網站名稱(key)與網址(value)配對儲存。同時，你也定義了一個列舉(enum)來表示網站名稱。你的程式碼注釋也清晰說明了這個函式的目的和用途，以及你的信仰和感恩之心。因此，這是一個良好的實作方式。
         */

        /// <summary>
        /// 取得Chrome瀏覽器的Processes陣列
        /// </summary>
        internal static Process[] GetChromeProcessInstances
        {
            get
            {
                chromeProcessInstances = Process.GetProcessesByName("chrome");
                return chromeProcessInstances;
            }

        }
        /// <summary>
        /// 判斷Chrome瀏覽器是否已開啟、在運行
        /// 20241002 creedit_with_Copilot大菩薩：Word VBA 中的 Selenium 操作：https://sl.bing.net/juGuOGNZlee
        /// </summary>
        internal static bool IsChromeRunning
        {
            get
            {
                return GetChromeProcessInstances.Length > 0;
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
            return new Tuple<ChromeDriver, RemoteWebDriver>(DriverNew(), null);
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


        ////static selm.IWebDriver driverNew()
        ////實測後發現：CurrentWindowHandle並不能取得瀏覽器現正作用中的分頁視窗，只能取得創建 ChromeDriver 物件時的最初及switch 方法執行後切換的分頁視窗 20230103 阿彌陀佛
        ///// <summary>
        ///// 記下啟動Selenium chromedriver時的第一個WindowHandle
        ///// </summary>
        //static string originalWindow;//20240911作廢
        //internal static string getOriginalWindow
        //{
        //    get
        //    {
        //        return originalWindow;
        //    }

        //}

        /// <summary>
        /// 記下操作時最後一個有效的WindowHandle        
        /// </summary>
        static string _lastValidWindowHandle;
        /// <summary>
        /// 取得/設定最近一個有效的WindowHandle
        /// 以便取得視窗或頁籤關閉後所發生的錯誤，在此之前可用有效的視窗或分頁。
        /// 若有錯誤，則傳回null值；若沒有則傳回空字串""（string.Empty)
        /// </summary>
        public static string LastValidWindow
        {
            get
            {
                if (driver == null) return null; ReadOnlyCollection<string> whs;
                try
                {
                    whs = driver.WindowHandles;
                }
                catch (Exception ex)
                {
                    switch (ex.HResult)
                    {
                        case -2146233088:
                            if (ex.Message.IndexOf("An unknown exception was encountered sending an HTTP request to the remote WebDriver server for URL") > -1)
                            {
                                Form1.playSound(Form1.soundLike.error);
                                try
                                {
                                    _lastValidWindowHandle = driver.WindowHandles[driver.WindowHandles.Count - 1];
                                }
                                catch (Exception)
                                {
                                    try
                                    {
                                        return driver.CurrentWindowHandle;

                                    }
                                    catch (Exception)
                                    {
                                        return null;
                                    }
                                }
                                return _lastValidWindowHandle;
                            }
                            else
                            {
                                if (DialogResult.OK == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly(ex.HResult + ex.Message + Environment.NewLine +
                                    "是否要重啟chromedriver？"))
                                {
                                    RestartChromedriver();
                                    if (driver.WindowHandles.Contains(_lastValidWindowHandle))
                                        return _lastValidWindowHandle;
                                    else
                                        return driver.WindowHandles.Last();
                                }
                                else
                                    return null;
                            }
                        default:
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                            return null;
                    }
                }

                whs = driver.WindowHandles;
                if (whs == null || whs.Count == 0) return null;// string.Empty;                
                                                               //_lastValidWindowHandle = _lastValidWindowHandle ?? (whs.Count > 0 ? whs[whs.Count - 1] : null);
                                                               //if (driver.WindowHandles.IndexOf(_lastValidWindowHandle)>-1)
                if (!driver.WindowHandles.Contains(_lastValidWindowHandle))
                {
                    for (int i = driver.WindowHandles.Count - 1; i > -1; i--)
                    {
                        driver.SwitchTo().Window(driver.WindowHandles[i]);
                        //if (DialogResult.OK == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是這個分頁嗎？"))
                        if (Form1.IsValidUrl＿ImageTextComparisonPage(driver.Url))
                        {
                            _lastValidWindowHandle = driver.WindowHandles[i];
                            return _lastValidWindowHandle;
                        }
                    }
                    return driver.WindowHandles.Last();
                }
                else
                    return _lastValidWindowHandle;
                /* 20230822 Bing大菩薩： https://sl.bing.net/fo9YWhdvMWG
                 沒關係，我知道您想問的是 `vbNullString`。`vbNullString` 是 VBA 中的一個常量，它表示空字串。它與 `""` 不同，因為 `vbNullString` 是一個空指針，而 `""` 是一個長度為零的實際字串¹²。在 VBA 中，使用 `vbNullString` 比使用 `""` 更快且更省內存²。但是，在 C# 中，您可以使用 `string.Empty` 或 `""` 來表示空字串，它們的效能差異非常小，幾乎可以忽略不計¹³⁴。希望我的回答能夠幫助您！
                    來源: 與 Bing 的交談， 2023/8/22
                    (1) vba - What is vbNullString, How it use? - Stack Overflow. https://stackoverflow.com/questions/46006679/what-is-vbnullstring-how-it-use.
                    (2) Is there any difference between vbNullString and - Stack Overflow. https://stackoverflow.com/questions/32435320/is-there-any-difference-between-vbnullstring-and.
                    (3) VBAのEmptyやNothingやNull等の違いを整理 ｜ Excel作業をVBAで効率化. https://vbabeginner.net/organize-differences-empty-nothing-null/.
                    (4) VB中判断空的几种方法,Null, Missing, Empty, Nothing, vbNullString区别 - *（00）* - 博客园. https://www.cnblogs.com/zouhao/p/3664651.html.
                    (5) undefined. http://www.aivosto.com/vbtips/stringopt.html.
                 */
            }

            set => _lastValidWindowHandle = value;
        }

        /// <summary>
        /// chromedriver被誤關時的處理 20241005
        /// 發生"An unknown exception was encountered sending an HTTP request to the remote WebDriver server for URL …"例外情形時
        /// </summary>
        /// <param name="ex"></param>
        /// <returns>是這個錯誤則傳回true </returns>
        internal static bool ChromedriverLose(Exception ex)
        {
            if (ex.Message.StartsWith("An unknown exception was encountered sending an HTTP request to the remote WebDriver server for URL ") ||
                ex.Message.StartsWith("disconnected: not connected to DevTools") ||
                ex.Message.StartsWith("invalid session id"))//An unknown exception was encountered sending an HTTP request to the remote WebDriver server for URL http://localhost:13451/session/6f6c77cfb73c5c388c6cdfd40a06b806/url. The exception message was: 傳送要求時發生錯誤。
            {
                Form1.playSound(Form1.soundLike.over);
                killchromedriverFromHere();
                driver = null;
                DriverNew();
                return true;
            }
            else
            {
                Console.WriteLine(ex.HResult + ex.Message);
                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                Debugger.Break();
                return false;
            }

        }
        /// <summary>
        /// 20240731 Copilot大菩薩：Selenium WebDriver 瀏覽器畫面閃爍問題解決方法
        /// Unfortunately, Selenium WebDriver does not have a built-in feature like MS Word VBA’s ScreenUpdating = false to prevent the browser from updating the screen when switching windows. However, you can try a workaround by using JavaScript to minimize the visual impact. Here’s an example of how you might achieve this:
        /// Hide the browser window before switching:        
        /// </summary>
        /// <param name="driver"></param>
        public static void HideBrowserWindow(IWebDriver driver)
        {
            ((IJavaScriptExecutor)driver).ExecuteScript("document.body.style.visibility='hidden'");
        }
        /// <summary>
        /// 20240731 Copilot大菩薩：Selenium WebDriver 瀏覽器畫面閃爍問題解決方法
        /// Unfortunately, Selenium WebDriver does not have a built-in feature like MS Word VBA’s ScreenUpdating = false to prevent the browser from updating the screen when switching windows. However, you can try a workaround by using JavaScript to minimize the visual impact. Here’s an example of how you might achieve this:
        /// Hide the browser window before switching:        
        /// </summary>
        /// <param name="driver"></param>
        public static void ShowBrowserWindow(IWebDriver driver)
        {
            ((IJavaScriptExecutor)driver).ExecuteScript("document.body.style.visibility='visible'");
        }

        /// <summary>
        /// 檢測視窗是否有效的方法
        /// 20240730 Copilot大菩薩：C# Selenium 瀏覽器分頁次序
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="handle"></param>
        /// <returns></returns>
        public static bool IsWindowHandleValid(IWebDriver driver, string handle)
        {
            try
            {//These methods should help reduce the flickering effect when switching windows. Please note that hiding or minimizing the browser window might affect the user experience, so use these methods with caution.
                driver.SwitchTo().Window(handle);
                return true;
            }
            catch (NoSuchWindowException)
            {
                return false;
            }
            catch (Exception ex)
            {
                //Console.WriteLine("發生錯誤: " + ex.Message);
                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + "發生錯誤: " + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// 定期更新視窗句柄集合的方法
        /// 20240730 Copilot大菩薩：C# Selenium 瀏覽器分頁次序
        /// </summary>
        /// <param name="driver"></param>
        /// <returns>有效的視窗句柄清單</returns>
        public static List<string> GetValidWindowHandles(IWebDriver driver)
        {
            string currentWindowsHandle = GetCurrentWindowHandle(driver);
            var windowHandles = driver.WindowHandles;
            var validWindowHandles = new List<string>();

            foreach (var handle in windowHandles)
            {
                if (IsWindowHandleValid(driver, handle))
                {
                    validWindowHandles.Add(handle);
                }
            }
            driver.SwitchTo().Window(currentWindowsHandle);
            return validWindowHandles;
        }
        /// <summary>
        /// 取得現行各分頁與視窗的有序的集合清單（失敗！！！！！）
        /// 20240730 Copilot大菩薩：C# Selenium 瀏覽器分頁次序
        /// </summary>
        /// <param name="driver"></param>
        /// <returns>現行各分頁與視窗的有序的集合清單</returns>
        public static List<string> GetValiOrdereddWindowHandles(IWebDriver driver)
        {
            // 獲取所有有效的視窗句柄
            var validWindowHandles = GetValidWindowHandles(driver);

            // 使用 JavaScript 獲取每個分頁的 URL
            var urls = new Dictionary<string, string>();
            foreach (var handle in validWindowHandles)
            {
                driver.SwitchTo().Window(handle);
                string url = driver.Url;
                urls[handle] = url;
            }

            // 使用 JavaScript 獲取分頁順序
            // 使用更穩健的 JavaScript 脚本来获取分頁順序
            var orderedUrls = ((IJavaScriptExecutor)driver).ExecuteScript(@"
        var iframes = document.querySelectorAll('iframe');
        var urls = [];
        for (var i = 0; i < iframes.length; i++) {
            try {
                urls.push(iframes[i].contentWindow.location.href);
            } catch (error) {
                console.error('Error getting iframe URL:', error);
            }
        }
        return urls;
    ") as List<object>;

            // 檢查 orderedUrls 是否為 null
            var orderedHandles = new List<string>();

            if (orderedUrls == null)
            {
                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("JavaScript 腳本未能正確返回分頁順序。");
                return orderedHandles;
            }

            // 根據 URL 比對視窗句柄順序
            foreach (var url in orderedUrls)
            {
                foreach (var kvp in urls)
                {
                    if (kvp.Value == url.ToString())
                    {
                        orderedHandles.Add(kvp.Key);
                        break;
                    }
                }
            }

            return orderedHandles;
        }




        /// 將取得的 List<string>清單轉換成ReadOnlyCollection<string>集合
        /// 20240730 Copilot大菩薩：Convert List of Window Handles to ReadOnlyCollection
        /// </summary>
        /// <param name="orderedHandles">要被轉換的清單變數</param>
        /// <returns></returns>
        public static ReadOnlyCollection<string> ConvertToReadOnlyCollection(List<string> orderedHandles)
        {
            return new ReadOnlyCollection<string>(orderedHandles);
        }
        /// <summary>
        /// 取得指定視窗句柄在Chrome瀏覽器中的的序號
        /// 20240730 Copilot大菩薩：C# Selenium 瀏覽器分頁次序
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="orderedHandles">現行有效依序的分頁或視窗句柄之集合清單</param>
        /// <returns></returns>
        public static int GetIndexofSpecificValidWindowHandle(List<string> orderedHandles)
        {
            // 查找指定句柄在清單中的位置
            string targetHandle = orderedHandles[1]; // 假設我們要查找第二個分頁的句柄
            int index = orderedHandles.IndexOf(targetHandle);
            //Console.WriteLine("指定句柄的位置: " + index);
            return index;
        }
        /// <summary>
        /// 獲取當前有效的視窗句柄之方法
        /// 20240730 Copilot大菩薩：C# Selenium 瀏覽器分頁次序
        /// </summary>
        /// <param name="driver"></param>
        /// <returns></returns>
        public static string GetCurrentWindowHandle(IWebDriver driver)
        {
            try
            {
                if (IsDriverInvalid())
                    driver.SwitchTo().Window(driver.WindowHandles.LastOrDefault());
                string currentHandle = driver.CurrentWindowHandle;
                if (driver.WindowHandles.Contains(currentHandle))
                {
                    return currentHandle;
                }
                else
                {
                    //throw new NoSuchWindowException("當前視窗句柄無效");
                    //Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("當前視窗句柄無效");
                    // 返回 null 而不是拋出異常
                    return null;
                }
            }
            catch (NoSuchWindowException)
            {
                try
                {
                    if (IsWindowHandleValid(driver, LastValidWindow))
                    {
                        driver.SwitchTo().Window(LastValidWindow);
                        return LastValidWindow;
                    }
                }
                catch (Exception ex)
                {
                    try
                    {
                        if (IsWindowHandleValid(driver, driver.WindowHandles.Last()))
                        {
                            driver.SwitchTo().Window(driver.WindowHandles.Last());
                            return driver.WindowHandles.Last();
                        }
                    }
                    catch (Exception)
                    {
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                    }
                }
                //Console.WriteLine("當前視窗句柄無效");
                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("當前視窗句柄無效");
                return null;
            }
            catch (Exception ex)
            {
                //Console.WriteLine("發生錯誤: " + ex.Message);
                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HelpLink + "發生錯誤: " + ex.Message);
                return null;
            }
        }

        /// <summary>
        /// 封裝視窗操作的方法，包含顯式等待
        /// 20240730 Copilot大菩薩：C# Selenium 瀏覽器分頁次序
        /// 定義了一個委派方法 PerformWindowOperation，這個方法接受三個參數：
        /// </summary>
        /// <param name="driver">WebDriver 物件</param>
        /// <param name="handle">視窗句柄。</param>
        /// <param name="operation">一個接受 IWebDriver 參數的操作（匿名方法）。</param>
        public static void PerformWindowOperation(IWebDriver driver, string handle, Action<IWebDriver> operation)
        {
            if (IsWindowHandleValid(driver, handle))
            {
                try
                {
                    // 顯式等待，確保視窗已經完全加載
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                    wait.Until(drv => drv.SwitchTo().Window(handle) != null);

                    operation(driver);
                }
                catch (WebDriverTimeoutException)
                {
                    //Console.WriteLine("視窗加載超時: " + handle);
                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("視窗加載超時: " + handle);
                }
                catch (Exception ex)
                {
                    //Console.WriteLine("操作失敗: " + ex.Message);
                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + "操作失敗: " + ex.Message);
                }
            }
            else
            {
                //Console.WriteLine("視窗無效: " + handle);
                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("視窗無效: " + handle);
            }
        }

        /// <summary>
        /// 取得 driver 的 URL
        /// </summary>
        internal static string GetDriverUrl
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
                            return driver.SwitchTo().Window(GetCurrentWindowHandle(driver)).Url;//"";
                        default:
                            throw;
                    }
                }
            }
        }

        /// <summary>
        /// 取得[簡單修改模式](quick edit)超連結控制項（元件）
        /// </summary>
        /// <returns>傳回[簡單修改模式](quick edit)控制項</returns>
        internal static IWebElement QuickeditLinkIWebElement
        {
            get
            {
                if (driver == null) driver = DriverNew();
                IWebElement iwe = WaitFindWebElementBySelector_ToBeClickable("#quickedit > a", 5);
                if (iwe != null)
                {
                    string iweText = iwe.GetAttribute("text");
                    if (iweText != "簡單修改模式" && iweText != "Quick edit")
                    {
                        //#quickedit > a:nth-child(1)                        
                        //# quickedit > a:nth-child(2)
                        if (DialogResult.OK == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否是這個超連結控制項？"
                            + Environment.NewLine + Environment.NewLine + iweText))
                            return iwe;
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("沒有找到正確的「簡單修改模式Quick edit」超連結控制項，請檢查！");
                    }
                }
                return iwe;
            }
        }

        /// <summary>
        /// 取得[簡單修改模式](quick edit)下的Save changes按鈕（元件）
        /// </summary>
        /// <returns>傳回[簡單修改模式](quick edit)下的Save changes按鈕控制項</returns>
        internal static IWebElement SavechangesButton
        {
            get
            {
                IWebElement iwe = null;
                //if (driver == null) driver = DriverNew();
                if (!IsDriverInvalid())
                {
                    iwe = WaitFindWebElementBySelector_ToBeClickable("#savechangesbutton");
                }
                return iwe;
            }
        }


        /// <summary>
        /// 取得CTP網頁中的「書名」（title）超連結控制項，含 href 屬性者
        /// </summary>
        internal static IWebElement Title_Linkbox_Link
        {
            get
            {
                const string selector = "#content > div:nth-child(3) > span:nth-child(2) > a";//32位元免安裝版Chrome瀏覽器               
                IWebElement iwe;
                //if (Form1.IsValidUrl＿keyDownCtrlAdd(ActiveForm1.textBox3Text))
                //{
                iwe = WaitFindWebElementBySelector_ToBeClickable(selector);
                return iwe;
                //}
                //else
                //    return null;                
            }
        }
        /// <summary>
        /// 取得圖文對照頁面中的「Add transcription」控制項（元件）
        /// </summary>
        internal static IWebElement AddTranscription_Linkbox
        {
            get
            {
                IWebElement iwe = WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > a");
                if (iwe == null) return null;
                if (iwe.GetAttribute("text") == "Add transcription")
                    return iwe;
                else
                    return null;
            }
        }
        /// <summary>
        /// 取得CTP網頁中的「書名」（title）控制項
        /// <span itemprop="title">純常子枝語</span>
        /// </summary>
        internal static IWebElement Title_Linkbox
        {
            get
            {
                const string selector = "#content > div:nth-child(3) > span:nth-child(2) > a > span";//32位元免安裝版Chrome瀏覽器
                const string selector1 = "#content > div:nth-child(5) > span:nth-child(2) > a > span"; //64位元安裝版Chrome瀏覽器
                IWebElement iwe;
                if (Form1.IsValidUrl＿keyDownCtrlAdd(ActiveForm1.textBox3Text))
                {
                    iwe = WaitFindWebElementBySelector_ToBeClickable(selector);
                reCheck:
                    if (iwe != null)
                    {
                        string tx = iwe.GetAttribute("outerHTML");
                        if (!tx.StartsWith("<span itemprop=\"title\">"))
                        {
                            iwe = WaitFindWebElementBySelector_ToBeClickable(selector1);//64位元安裝版Chrome瀏覽器
                            if (iwe != null)
                            {
                                tx = iwe.GetAttribute("outerHTML");
                                if (!tx.StartsWith("<span itemprop=\"title\">"))
                                {
                                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("未能找到正確的「書名（title）」超連結控制項，請檢查！", "Title_Linkbox div:nth-child(5) !tx.StartsWith(\"<span itemprop=\\\"title\\\">\"))");
                                    return null;
                                }
                            }
                            else
                            {
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("未能找到正確的「書名（title）」超連結控制項，請檢查！", "Title_Linkbox div:nth-child(5)=null");
                                return null;
                            }
                        }
                        else
                            return iwe;
                    }
                    else
                    {
                        iwe = WaitFindWebElementBySelector_ToBeClickable(selector1);
                        if (iwe != null)
                            goto reCheck;
                        else
                        {
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("未能找到正確的「書名（title）」超連結控制項，請檢查！", "Title_Linkbox");
                            return null;
                        }
                    }
                }
                else
                    return null;
                return iwe;
            }
        }
        /// <summary>
        /// 取得「Please confirm that you are human! 敬請輸入認證圖案」元件
        /// </summary>
        internal static IWebElement Please_confirm_that_you_are_human_Page
        {
            get
            {
                IWebElement iwe = WaitFindWebElementBySelector_ToBeClickable("#content > font");
                if (iwe == null) return null;
                if (iwe.GetAttribute("textContent") == "Please confirm that you are human! 敬請輸入認證圖案")
                    return iwe;
                else
                    return null;
            }
        }
        internal static IWebElement GraphicMatchingPagesLink
        {
            get => WaitFindWebElementBySelector_ToBeClickable("#p2 > td:nth-child(1) > div > a.sprite-photo > div", 3);
        }

        /// <summary>
        /// 取得CTP網頁中的「編輯」（Edit）控制項
        /// </summary>
        internal static IWebElement Edit_Linkbox
        {
            get
            {
                IWebElement iwe;
                if (Form1.IsValidUrl＿keyDownCtrlAdd(ActiveForm1.textBox3Text))
                {
                    //會因位置而移動，如：Add to 學海蠡測 Add to 思舊錄 [文字版] [編輯] [簡單修改模式] [編輯指南] https://ctext.org/library.pl?if=gb&file=194081&page=75&editwiki=5083072#editor
                    //故得逐一比對，目前應該只會有2種情形，當然也可能會不止如此
                    iwe = WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > div:nth-child(2) > a:nth-child(2)");
                reCheck:
                    if (iwe != null)
                    {
                        string tx = iwe.GetAttribute("text");
                        if (tx != "編輯" && tx != "Edit")
                        {
                            iwe = WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > div:nth-child(2) > a:nth-child(4)");
                            if (iwe != null)
                                tx = iwe.GetAttribute("text");
                            if (tx != "編輯" && tx != "Edit")
                            {
                                iwe = WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > div:nth-child(2) > a:nth-child(3)");
                                if (iwe != null)
                                    tx = iwe.GetAttribute("text");
                                if (tx != "編輯" && tx != "Edit")
                                {
                                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("未能找到正確的「編輯（Edit）」超連結控制項，請檢查！");
                                    return null;
                                }
                                else
                                    return iwe;
                            }
                            else
                                return iwe;
                        }
                        //Edit_Linkbox = waitFindWebElementByName_ToBeClickable("#content > div:nth-child(7) > div:nth-child(2) > a:nth-child(2)", WebDriverWaitTimeSpan);
                    }
                    else
                    {
                        iwe = WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(9) > div:nth-child(2) > a:nth-child(2)");
                        if (iwe != null) goto reCheck;
                    }
                }
                else
                    return null;
                return iwe;
            }
        }
        /// <summary>
        /// 取得CTP網頁中的「參考上下頁」（ check the adjacent pages）控制項
        /// </summary>
        internal static IWebElement CheckAdjacentPages_Linkbox
        {
            //get { return quickedit_data_textbox == null ? waitFindWebElementByName_ToBeClickable("data", WebDriverWaitTimeSpan) : quickedit_data_textbox; }
            get
            {
                IWebElement iwe;
                if (Form1.IsValidUrl＿keyDownCtrlAdd(ActiveForm1.textBox3Text))
                {
                    iwe = WaitFindWebElementBySelector_ToBeClickable("#editor > a:nth-child(13)");
                }
                else
                    return null;
                return iwe;
            }
        }
        /// <summary>
        /// 取得CTP網頁中的「上一頁」的編輯文字方塊
        /// </summary>
        internal static IWebElement CheckAdjacentPages_DataPrev
        {
            //get { return quickedit_data_textbox == null ? waitFindWebElementByName_ToBeClickable("data", WebDriverWaitTimeSpan) : quickedit_data_textbox; }
            get
            {
                IWebElement iwe;
                if (Form1.IsValidUrl＿keyDownCtrlAdd(ActiveForm1.textBox3Text))
                {
                    iwe = WaitFindWebElementBySelector_ToBeClickable("#dataprev");
                }
                else
                    return null;
                return iwe;
            }
        }
        /// <summary>
        /// 取得CTP網頁中的「下一頁」(Next page:)的編輯文字方塊
        /// </summary>
        internal static IWebElement CheckAdjacentPages_DataNext
        {
            //get { return quickedit_data_textbox == null ? waitFindWebElementByName_ToBeClickable("data", WebDriverWaitTimeSpan) : quickedit_data_textbox; }
            get
            {
                IWebElement iwe;
                if (Form1.IsValidUrl＿keyDownCtrlAdd(ActiveForm1.textBox3Text))
                {
                    iwe = WaitFindWebElementBySelector_ToBeClickable("#datanext");
                }
                else
                    return null;
                return iwe;
            }
        }
        /// <summary>
        /// 下一頁控制項（元件）
        /// </summary>
        internal static IWebElement NextPage
        {
            get => WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(3) > div:nth-child(5) > a > div", 5);
        }
        /// <summary>
        /// 取得CTP網頁中的「顯示頁碼，可輸入頁碼的」（page）控制項
        /// </summary>
        internal static IWebElement Page_textbox
        {
            //get { return quickedit_data_textbox == null ? waitFindWebElementByName_ToBeClickable("data", WebDriverWaitTimeSpan) : quickedit_data_textbox; }
            get
            {
                IWebElement iwe;
                bool checkNamePorp()
                {
                    return iwe?.GetAttribute("name") == "page";
                }
                if (Form1.IsValidUrl＿ImageTextComparisonPage(ActiveForm1.textBox3Text))
                {
                    iwe = WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(3) > form > input[type=text]:nth-child(3)");
                    if (iwe == null)
                    {
                        iwe = WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(5) > form");
                    }
                    if (!checkNamePorp()) return null;
                }
                else
                    return null;
                return iwe;
            }
        }
        /// <summary>
        /// 取得CTP網頁中的「顯示頁碼資訊的條幅」（page）控制項（以取得該書的末頁）
        /// </summary>
        internal static IWebElement Div_generic_IncludePathAndEndPageNum
        {
            //get { return quickedit_data_textbox == null ? waitFindWebElementByName_ToBeClickable("data", WebDriverWaitTimeSpan) : quickedit_data_textbox; }
            get
            {
                IWebElement iwe;
                if (Form1.IsValidUrl＿ImageTextComparisonPage(ActiveForm1.textBox3Text))
                {
                    iwe = WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(3)");
                }
                else
                    return null;
                return iwe;
            }
        }
        /// <summary>
        /// 取得某書書頁碼的上限值
        /// 若出錯則傳回0
        /// </summary>
        /// <returns></returns>
        internal static int pageUBound
        {
            get
            {
                IWebElement iwe = Div_generic_IncludePathAndEndPageNum;
                if (iwe == null) return 0;
                string input = iwe.GetAttribute("textContent");//"線上圖書館 -> 松煙小錄 -> 松煙小錄三  /117 ";
                return CnText.ExtractNumberAfterSlash(input);
            }
        }
        /// <summary>
        /// 取得CTP網頁中的「文本框」（文字框）（圖文對照的文框）控制項
        /// </summary>
        internal static IWebElement Div_generic_TextBoxFrame
        {
            //get { return quickedit_data_textbox == null ? waitFindWebElementByName_ToBeClickable("data", WebDriverWaitTimeSpan) : quickedit_data_textbox; }
            get
            {
                IWebElement iwe;
                if (Form1.IsValidUrl＿ImageTextComparisonPage(ActiveForm1.textBox3Text))
                {
                    iwe = WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > div:nth-child(1)");
                }
                else
                    return null;
                return iwe;
            }
        }
        /// <summary>
        /// 取得CTP網頁中的「書圖框」（圖文對照的圖框.svg）控制項
        /// </summary>
        internal static IWebElement Svg_image_PageImageFrame
        {
            //get { return quickedit_data_textbox == null ? waitFindWebElementByName_ToBeClickable("data", WebDriverWaitTimeSpan) : quickedit_data_textbox; }
            get
            {
                IWebElement iwe;
                if (Form1.IsValidUrl＿ImageTextComparisonPage(ActiveForm1.textBox3Text))
                {
                    iwe = WaitFindWebElementBySelector_ToBeClickable("#canvas > svg");
                }
                else
                    return null;
                if (iwe == null)
                    iwe = WaitFindWebElementBySelector_ToBeClickable("#previmg");
                if (iwe == null)
                {
                    if (!IsDriverInvalid())
                        iwe = WaitFindWebElementBySelector_ToBeClickable("#previmg");
                    //iwe = driver.FindElement(By.XPath("/html/body/div[2]/div[3]/img"));
                    else
                    {
                        if (!driver.WindowHandles.Contains(LastValidWindow))
                            LastValidWindow = driver.WindowHandles.Last();
                        driver.SwitchTo().Window(LastValidWindow);

                        iwe = WaitFindWebElementBySelector_ToBeClickable("#previmg");
                    }
                }
                return iwe;
            }
        }
        /// <summary>
        /// 自動全選[Quick edit]的內容，方便有時候須用剪下貼上者
        /// </summary>
        /// <returns>成功則傳回true</returns>
        internal static bool SelectAllQuickedit_data_textboxContent()
        {
            OpenQA.Selenium.IWebElement ie = Quickedit_data_textbox;//br.QuickeditIWebElement;
            if (ie != null)
            {
                ie.SendKeys(OpenQA.Selenium.Keys.Control + "a");
                return true;
            }
            return false;
        }


        /// <summary>
        /// 取得如欽定四庫全書的版本連結元件；若失敗則回傳null
        /// </summary>
        internal static IWebElement Version_LinkBox
        {
            get
            {
                IWebElement version_LinkBox;
                if (Form1.IsValidUrl＿keyDownCtrlAdd(ActiveForm1.textBox3Text))
                {
                    //  /html/body/div[2]/div[5]/div[3]/a[1]
                    version_LinkBox = WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(8) > div:nth-child(3) > a:nth-child(1)");
                    if (version_LinkBox == null)
                        // /html/body/div[2]/div[5]/div/a[1] /html/body/div[2]/div[5]/div/a[1]
                        version_LinkBox = WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(8) > div > a:nth-child(1)");

                }
                else
                    version_LinkBox = null;
                return version_LinkBox;
            }
        }
        /// <summary>
        /// 取得[簡單修改模式]的文字方塊（編輯區的文字方塊）；若失敗則回傳null        
        /// Get the textbox of [Quick edit] 
        /// </summary>
        internal static IWebElement Quickedit_data_textbox
        {
            //get { return quickedit_data_textbox == null ? waitFindWebElementByName_ToBeClickable("data", WebDriverWaitTimeSpan) : quickedit_data_textbox; }
            get
            {
                if (Form1.IsValidUrl＿keyDownCtrlAdd(ActiveForm1.textBox3Text))
                {
                    quickedit_data_textbox = waitFindWebElementByName_ToBeClickable("data", WebDriverWaitTimeSpan);
                }
                else
                    quickedit_data_textbox = null;
                return quickedit_data_textbox;
            }
            private set { quickedit_data_textbox = value; }
        }
        /// <summary>
        /// 取得[編輯]的文字方塊（編輯區的文字方塊）；若失敗則回傳null 20240929 于52生日
        /// Get the textbox of [edit] 
        /// </summary>
        internal static IWebElement Textarea_data_Edit_textbox
        {
            //get { return quickedit_data_textbox == null ? waitFindWebElementByName_ToBeClickable("data", WebDriverWaitTimeSpan) : quickedit_data_textbox; }
            get
            {

                if (driver.Url.IndexOf("&action=editchapter#") > -1)
                {
                    return WaitFindWebElementBySelector_ToBeClickable("#data", WebDriverWaitTimeSpan);
                }
                else
                    return null;
            }
            //private set { Textarea_data_Edit_textbox = value; }
        }


        /// <summary>
        /// 儲存[簡單修改模式]的文字方塊
        /// </summary>
        private static IWebElement quickedit_data_textbox = null;

        private static string quickedit_data_textboxTxt = "";

        /// <summary>
        /// 取得[簡單修改模式]的文字；若失敗則回傳空字串
        /// 原來取該元件的「value」Property就可以了20240913
        /// </summary>
        internal static string Quickedit_data_textboxTxt
        {
            get
            {
                if (!Form1.IsValidUrl＿keyDownCtrlAdd(ActiveForm1.textBox3Text)) return string.Empty;
                IWebElement ie = Quickedit_data_textbox;
                if (ie != null)
                {
                    //20240913 原來取該元件的「value、textContent……」等 Property 就可以了！
                    //.Text屬性會清除起首的全形空格！！20240313
                    //if (quickedit_data_textboxTxt != Quickedit_data_textbox.Text) quickedit_data_textboxTxt = quickedit_data_textbox.Text;                    
                    //string quickedit_data_textbox_Txt = CopyQuickedit_data_textboxText();                    
                    //if (quickedit_data_textboxTxt != quickedit_data_textbox_Txt) quickedit_data_textboxTxt = quickedit_data_textbox_Txt;
                    //return quickedit_data_textboxTxt;
                    return ie.GetAttribute("value");
                }
                else
                    return string.Empty;
            }
        }
        /// <summary>
        /// 設定Quickedit_data_textbox的value屬性值  20240913
        /// creedit_with_Copilot大菩薩：C# Selenium 屬性設定方法： https://sl.bing.net/jv1AQReen36
        /// </summary>
        /// <param name="txt">要設定的值</param>
        /// <returns>若失敗則傳回false</returns>
        internal static bool SetQuickedit_data_textboxTxt(string txt)
        {
            if (!Form1.IsValidUrl＿keyDownCtrlAdd(ActiveForm1.textBox3Text)) return false;
            IWebElement ie = Quickedit_data_textbox;
            if (ie != null)
            {
                if (SetIWebElementValueProperty(ie, txt))
                    return true;
                else
                    return false;
            }
            else
                return false;

        }
        /// <summary>
        /// 設定元件 IWebElement的value屬性值  20240913
        /// creedit_with_Copilot大菩薩：C# Selenium 屬性設定方法： https://sl.bing.net/jv1AQReen36
        /// </summary>
        /// <param name="txt">要設定的值</param>
        /// <returns>若失敗則傳回false</returns>
        internal static bool SetIWebElementValueProperty(IWebElement iwe, string txt)
        {/* 其餘諸如 屬性 defaultValue、innerHTML、textContent 
          * ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].defaultValue = arguments[1];", iwe, txt);
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].innerHTML = arguments[1];", iwe, txt);
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].textContent = arguments[1];", iwe, txt);
            https://sl.bing.net/ggwydu064om
          */
            if (iwe != null)
            {
                //// 假設您已經初始化了 WebDriver
                //IWebDriver driver = new ChromeDriver();

                //// 找到您想要設定屬性的元素
                //IWebElement element = driver.FindElement(By.Id("elementId"));

                // 使用 ExecuteScript 方法來設定屬性值
                //((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].setAttribute('attributeName', 'attributeValue')", element);
                //((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].setAttribute('value', '" + txt + "')", ie);
                //((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].setAttribute('defaultValue', 'test')", ie);
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].value = arguments[1];", iwe, txt);
                /*這是因為 value 屬性在 JavaScript 中是動態的，當您使用 arguments[0].value = arguments[1]; 設定值時，它會改變元素的顯示值，但不會改變元素的 HTML 屬性值。
                    如果您希望在開發人員工具中也看到 value 屬性值的改變，可以使用 setAttribute 方法來同步更新：*/
                //((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].setAttribute('value', arguments[1]);", iwe, txt);//現在不需要在開發窗格中顯示，故省略
                // 這樣，您就可以在開發人員工具中看到 value 屬性值的變化了。

                return true;
            }
            else
                return false;
        }
        /// <summary>
        /// 設定元件 IWebElement的textContent屬性值  20240914
        /// creedit_with_Copilot大菩薩：C# Selenium 屬性設定方法： https://sl.bing.net/jv1AQReen36
        /// </summary>
        /// <param name="txt">要設定的值</param>
        /// <returns>若失敗則傳回false</returns>
        internal static bool SetIWebElement_textContent_Property(IWebElement iwe, string txt)
        {/* 其餘諸如 屬性 defaultValue、innerHTML、textContent 
          * ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].defaultValue = arguments[1];", iwe, txt);
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].innerHTML = arguments[1];", iwe, txt);
            https://sl.bing.net/ggwydu064om
          */
            if (iwe != null)
            {
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].textContent = arguments[1];", iwe, txt);
                /*這是因為 value 屬性在 JavaScript 中是動態的，當您使用 arguments[0].value = arguments[1]; 設定值時，它會改變元素的顯示值，但不會改變元素的 HTML 屬性值。
                    如果您希望在開發人員工具中也看到 value 屬性值的改變，可以使用 setAttribute 方法來同步更新：*/
                //((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].setAttribute('value', arguments[1]);", iwe, txt);//現在不需要在開發窗格中顯示，故省略
                // 這樣，您就可以在開發人員工具中看到 value 屬性值的變化了。

                return true;
            }
            else
                return false;
        }
        /// <summary>
        /// 當Quickedit_data_textbox的內容是以全形空格開頭的會被清除，類似Trim的功能，故須用複製文本的方式取得正確的值
        /// 解決Selenium在[簡單修改模式]文字方塊內容若以全形空格為開頭的，會被截去的方案 20230829        /// 
        /// 原來取該元件的「value」Property就可以了 20240913
        /// </summary>
        /// <returns>回傳所複製的Quickedit_data_textbox文本</returns>
        internal static string CopyQuickedit_data_textboxText()
        {
            IWebElement ie = Quickedit_data_textbox;
            if (ie != null)
            {
                //[簡單修改模式]方塊若不存在
                if (WaitFindWebElementBySelector_ToBeClickable("#data") == null)
                {
                    //[簡單修改模式]超連結
                    if (WaitFindWebElementBySelector_ToBeClickable("#quickedit > a") != null)
                    {
                        //按下[簡單修改模式]超連結
                        WaitFindWebElementBySelector_ToBeClickable("#quickedit > a").Click();
                    }
                    else
                        return string.Empty;
                    quickedit_data_textbox = WaitFindWebElementBySelector_ToBeClickable("#data");
                    ie = Quickedit_data_textbox;
                }
                if (ie.Text != string.Empty)
                {
                    //ie.SendKeys(OpenQA.Selenium.Keys.Control + "a");//會移動視窗焦點到文字方塊 ie（Quickedit_data_textbox）中
                    SelectAllQuickedit_data_textboxContent();
                    ie.SendKeys(OpenQA.Selenium.Keys.Control + "c");
                    WindowsScrolltoTop();
                    //Clipboard.SetText(ie.Text);//.Text屬性會清除前首的全形空格，不適用！！20240313
                    DateTime dt = DateTime.Now;
                    while (!Form1.isClipBoardAvailable_Text())
                        if (DateTime.Now.Subtract(dt).TotalSeconds > 2) break;
                }
                else
                    Clipboard.Clear();
                return Clipboard.GetText();
            }
            else
            {
                Clipboard.Clear();
                return string.Empty;
            }
        }

        internal static IWebElement Full_text_search_textbox_searchressingle
        {
            get
            {
                if (driver == null) return null;
                IWebElement full_text_search_textbox_searchressingle = null;
                try
                {
                    if (Form1.IsValidUrl＿keyDownCtrlAdd(ActiveForm1.textBox3Text))
                    {
                        //< input type = "hidden" name = "searchressingle" id = "searchressingle" value = "wiki:728745" style = "width: 80px;" >
                        //full_text_search_textbox_searchressingle = driver.FindElement(By.Name("searchressingle"));
                        full_text_search_textbox_searchressingle = driver.FindElement(By.CssSelector("#searchressingle"));
                        if (full_text_search_textbox_searchressingle == null)
                        {
                            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(2));
                            full_text_search_textbox_searchressingle =
                                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.Name("searchressingle")));
                        }
                    }
                    return full_text_search_textbox_searchressingle;
                }
                catch (Exception)
                {
                    return null;
                }

            }
        }
        internal static IWebElement waitFindWebElementByName_ToBeClickable(string name, double second,
            IWebDriver drver = null)
        {
            IWebElement e = null;
            try
            {
                if (driver == null && drver == null)
                    RestartDriver();
                else
                    e = (driver ?? drver).FindElement(By.Name(name));
            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case -2146233088:
                        if (ex.Message.IndexOf("no such window: target window already closed") > -1)
                            if (GoToCurrentUserActivateTab() == string.Empty) return null;
                            else if (ex.Message.IndexOf("no such element: Unable to locate element") > -1)
                                GoToUrlandActivate(GetQuickeditUrl());
                        try
                        {
                            e = (driver ?? drver).FindElement(By.Name(name));
                        }
                        catch (Exception exex)
                        {
                            if (exex.Message.IndexOf("no such element: Unable to locate element") == -1)
                            {
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(exex.HResult + exex.Message);
                                return null;
                            }
                        }
                        break;
                    case -2147467261:
                        if (ex.Message.IndexOf("並未將物件參考設定為物件的執行個體。") > -1)
                            return null;
                        else
                        {
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                            return null;
                        }
                    default:
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                        return null;
                }
            }
            if (e != null)
            {
                try
                {
                    WebDriverWait wait = new WebDriverWait((driver ?? drver), TimeSpan.FromSeconds(second));
                    wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(e));
                }
                catch (Exception ex)
                {
                    switch (ex.HResult)
                    {
                        case -2147467261:
                            if (ex.Message.IndexOf("並未將物件參考設定為物件的執行個體。") > -1)
                            {
                                return null;
                            }
                            break;
                        default:
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                            break;
                    }
                }
            }
            return e;
        }
        internal static IWebElement waitFindWebElementById_ToBeClickable(string id, double second)
        {
            try
            {
                //driver.SwitchTo().Window(driver.CurrentWindowHandle);
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

        /// <summary>
        /// 取出如以下這個字串中的「tr:nth-child(2)」這個個部分的「2」這個數值以供計算，如加1後變成3，而轉置回這個Selector的字串中
        /// #content > div:nth-child(6) > table > tbody > tr:nth-child(2) > td:nth-child(1) > a
        /// 20250305 GitHub　Copilot大菩薩
        /// </summary>
        /// <param name="selector"></param>
        /// <returns></returns>
        internal static string IncrementNthChild(string selector)
        {
            var match = Regex.Match(selector, @"tr:nth-child\((\d+)\)");
            if (match.Success)
            {
                int number = int.Parse(match.Groups[1].Value);
                number++;
                return Regex.Replace(selector, @"tr:nth-child\(\d+\)", $"tr:nth-child({number})");
            }
            return selector;
        }
        /// <summary>
        /// 取得目前chapter（冊）的Selector值，不存在則傳回null
        /// </summary>
        internal static string ChapterSelector
        {
            set
            {
                if (!WindowHandles.TryGetValue("ChapterSelector", out _))
                    WindowHandles.Add("ChapterSelector", value);
                else
                    WindowHandles["ChapterSelector"] = value;
            }
            get
            {
                if (!WindowHandles.TryGetValue("ChapterSelector", out string chapterSelector))
                    return null;
                else
                    return chapterSelector;
            }
        }
        /// <summary>
        /// 取得下一個chapter（冊）的Selector值，不存在則傳回null
        /// </summary>
        internal static string NextChapterSelector
        {
            get
            {
                if (ChapterSelector == null)
                    return null;

                string selector = ChapterSelector;//"#content > div:nth-child(6) > table > tbody > tr:nth-child(2) > td:nth-child(1) > a";
                                                  //if (!WindowHandles.TryGetValue("ChapterSelecto", out string chapterSelector))
                                                  //    WindowHandles.Add("ChapterSelector ",);
                                                  //else
                                                  //{
                string newSelector = IncrementNthChild(selector);
                //Console.WriteLine(newSelector); // 輸出: #content > div:nth-child(6) > table > tbody > tr:nth-child(3) > td:nth-child(1) > a
                ChapterSelector = newSelector;
                return newSelector;
                //}
            }
        }
        /// <summary>
        /// 取得目前章節chapter（冊）的序號，以供Selector字串參照使用
        /// </summary>
        internal static string CurrentChapterNum_Selector
        {
            get
            {
                string selector = ChapterSelector;//"#content > div:nth-child(6) > table > tbody > tr:nth-child(2) > td:nth-child(1) > a";
                var match = Regex.Match(selector, @"tr:nth-child\((\d+)\)");
                //if (match.Success)
                return match.Groups[1].Value;


                //string pattern = @"tr:nth-child\((\d+)\)";//@"nth-child\((\d+)\)";
                //MatchCollection matches = Regex.Matches(selector, pattern);
                //foreach (Match match in matches)
                //{
                //    // 提取括號中的數值
                //    int value = int.Parse(match.Groups[1].Value);
                //    //Console.WriteLine($"nth-child 的值: {value}");
                //    // 在這裡進行你的後續計算
                //    retun
                //}
            }
        }
        /// <summary>
        /// 以 Selector 為條件尋找網頁元件
        /// </summary>
        /// <param name="selector">所根據的 Selector 字串值</param>
        /// <param name="second">等待元素出現的秒數上限</param>
        /// <returns>成功找到則回傳該元件物件，出錯則回傳 null</returns>
        internal static IWebElement WaitFindWebElementBySelector_ToBeClickable(string selector, double second = 2)
        {
            try
            {
                if (driver == null)
                {
                    Form1.browsrOPMode = Form1.BrowserOPMode.seleniumNew;
                    driver = DriverNew();
                }
                IWebElement e = driver.FindElement(By.CssSelector(selector));
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(second));
                try
                {
                    wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(e));
                }
                catch (Exception)
                {
                    return null;
                    //throw;
                }
                return e;
            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case -2146233088://"no such window: target window already closed\nfrom unknown error: web view not found\n  (Session info: chrome=109.0.5414.120)"
                        return null;
                    case -2147467261:
                        if (ex.Message.IndexOf("並未將物件參考設定為物件的執行個體。") > -1)
                            return null;
                        else
                        {
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                            return null;
                        }
                    default:
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                        return null;
                }
            }
        }

        internal static IWebElement waitFindWebElementByTagName_ToBeClickable(string TagName, double second = 2)
        {
            try
            {
                IWebElement e = driver.FindElement(By.TagName(TagName));
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
        /// Selenium 操控的 Chrome瀏覽器伺服器（ChromeDriverService）的等待秒數（即「new ChromeDriver()」的「TimeSpan」引數值）。預設為 30.5。
        /// </summary>
        static double _chromeDriverServiceTimeSpan = 30.5;//《古籍酷》OCR所需
        /// <summary>
        ///  Selenium 操控的 Chrome瀏覽器中網頁元件的的等待秒數（WebDriverWait。即「new WebDriverWait()」的「TimeSpan」引數值）。預設為 3。
        static double _webDriverWaitTimSpan = 5;
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
        /// </summary>gjc
        /// <returns></returns>
        internal static ChromeDriver DriverNew()
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

            //bool isChromeRunning;= IsChromeRunning;//作為是否在前已開啟Chrome瀏覽器時，關閉新多開的 20241002

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
                    if (_chromeDriverServiceTimeSpan < 30.5) _chromeDriverServiceTimeSpan = 30.5;//《古籍酷》OCR所需 
                    if (chromedriversPID == null) chromedriversPID = new List<int>();
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
                                         //options.AddArgument("--remote-debugging-port=9222");//chrome://inspect/#devices 20241002 Copilot大菩薩：Word VBA 中的 Selenium 操作： https://sl.bing.net/BkGRjHWbCe
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
                            if (ex.Message.IndexOf("This version of ChromeDriver only supports Chrome") > -1)
                            {
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請更新 chromedriver 才能繼續");
                                Form1.browsrOPMode = Form1.BrowserOPMode.appActivateByName; killchromedriverFromHere();
                                Process.Start("https://googlechromelabs.github.io/chrome-for-testing/#stable");
                                Process.Start(chrome_path);

                                return null;
                            }
                            else if (ex.Message.StartsWith("session not created: Chrome failed to start: exited normally."))//"session not created: Chrome failed to start: exited normally.\n  (session not created: DevToolsActivePort file doesn't exist)\n  (The process started from chrome location W:\\PortableApps\\PortableApps\\GoogleChromePortable\\App\\Chrome-bin\\chrome.exe is no longer running, so ChromeDriver is assuming that Chrome has crashed.) (SessionNotCreated)"
                            {
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(@"請手動關閉Chrome瀏覽器，再按「ok」確定，以繼續");
                                killchromedriverFromHere();
                                goto tryagain;
                            }
                            else if (ex.Message.StartsWith("session not created\nfrom disconnected: unable to connect to renderer (SessionNotCreated)"))//-2146233079session not created
                                                                                                                                                        //from disconnected: unable to connect to renderer(SessionNotCreated)
                            {
                                //Debugger.Break();
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(@"請手動關閉Chrome瀏覽器，再按「ok」確定，以繼續");
                                killchromedriverFromHere();
                                //driver = null;
                                RestartChromedriver();
                                return cDrv;
                                //goto tryagain;
                            }
                            else
                            {
                                Console.WriteLine(ex.HResult + ex.Message);
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                                return null;
                                //-2146233079session not created（重新安裝Chrome瀏覽器（免安裝版）即可解決）
                                //from unknown error: cannot parse internal JSON template: Line: 1, column: 1, Unexpected token. (SessionNotCreated)
                            }
                        default:
                            Console.WriteLine(ex.HResult + ex.Message);
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                            return null;
                    }
                }
                #endregion

                #region 免安裝版Chrome瀏覽器適用
                if (cDrv.WindowHandles.Count > 1)
                {
                    foreach (var item in cDrv.WindowHandles)
                    {
                        cDrv.SwitchTo().Window(item);
                        if (cDrv.Title == "新分頁" || cDrv.Title == string.Empty || cDrv.Url == "chrome://new-tab-page/")
                        {
                            cDrv.Close();
                            Form1.playSound(Form1.soundLike.over, true);

                            //break;
                        }

                    }


                    //20250212creedit with Gemini大菩薩：C# Windows.Forms 我想把視窗標題為「新分頁 - Google Chrome」的視窗關掉請問該怎麼寫？
                    Process[] chromeProcesses = Process.GetProcessesByName("chrome");

                    foreach (Process process in chromeProcesses)
                    {
                        // 檢查視窗標題是否為「新分頁 - Google Chrome」
                        if (process.MainWindowTitle == "新分頁 - Google Chrome")
                        {
                            // 關閉視窗
                            process.CloseMainWindow();
                            // 如果需要強制關閉，可以使用 Kill() 方法
                            // process.Kill();
                            break; // 找到目標視窗後，跳出迴圈
                        }
                    }
                    cDrv.SwitchTo().Window(cDrv.WindowHandles.Last());
                    try
                    {
                        cDrv.SwitchTo().NewWindow(WindowType.Tab);
                    }
                    catch (Exception)
                    {
                        cDrv.SwitchTo().Window(cDrv.WindowHandles.Last());
                        openNewTabWindow();
                        //cDrv.SwitchTo().NewWindow(WindowType.Tab);
                    }

                }
                #endregion//#region 免安裝版Chrome瀏覽器適用

                #region 成功開啟Chrome瀏覽器後
                //originalWindow = cDrv.CurrentWindowHandle; LastValidWindow = originalWindow;
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
                ActiveForm1 = Application.OpenForms["Form1"] as Form1;
                //到指定網頁
                string url = ActiveForm1.Controls["textBox3"].Text != "" ? ActiveForm1.Controls["textBox3"].Text : "https://ctext.org/account.pl?if=en";
                if (url.StartsWith("http"))
                    try
                    {
                        cDrv.Navigate().GoToUrl(url);
                        cDrv.SwitchTo().Window(cDrv.CurrentWindowHandle);
                    }
                    catch (Exception ex)
                    {
                        switch (ex.HResult)
                        {
                            case -2146233088:
                                if (ex.Message.IndexOf("no such window: target window already closed") > -1)//"no such window: target window already closed\nfrom unknown error: web view not found\n  (Session info: chrome=110.0.5481.178)"
                                {
                                    NoSuchWindowErrHandler();
                                }
                                else if (ex.Message.StartsWith("disconnected: not connected to DevTools"))//Message=disconnected: not connected to DevTools
                                                                                                          //(failed to check if window was closed: disconnected: not connected to DevTools)
                                    RestartChromedriver();
                                else
                                    goto default;
                                break;
                            default:
                                Console.WriteLine(ex.HResult + ex.Message);
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                                break;
                        }
                        //throw;
                    }

                if (!chromedriversPID.Contains(driverService.ProcessId)) chromedriversPID.Add(driverService.ProcessId);
                //配置quickedit_data_textbox以備用

                driver = cDrv;// quickedit_data_textboxSetting 方法堆疊（stack）中要用到driver參考

                if (Form1.IsValidUrl＿ImageTextComparisonPage(url))
                    quickedit_data_textboxSetting(url, null, cDrv);
                //IWebElement clk  = cDrv.FindElement(selm.By.Id("logininfo")); clk.Click();
                //cDrv.FindElement(selm.By.Id("logininfo")).Click();
                /*202301050214 因為以下這行設定成功，可以用平常的Chrome來操作了，就不必再登入安裝（如擴充功能）匯入（如書籤）什麼的了 感恩感恩　讚歎讚歎　南無阿彌陀佛
                 options.AddArgument("--user-data-dir=C:\\Users\\oscar\\AppData\\Local\\Google\\Chrome\\User Data\\");
                options.AddArgument("--user-data-dir="+ Environment.GetFolderPath( Environment.SpecialFolder.LocalApplicationData) +"\\Google\\Chrome\\User Data\\");
                 */
                //MessageBox.Show("請先登入 Ctext.org 再繼續。按下「確定(OK)」以繼續……");                

                //如果是手動輸入模式且在簡單編輯頁面，則將其Quick edit值傳到textBox1
                try
                {
                    if (ActiveForm1.KeyinTextMode && isQuickEditUrl(ActiveForm1.textBox3Text ?? ""))
                    {
                        try
                        {
                            driver = driver ?? cDrv;
                            try
                            {
                                ActiveForm1.Controls["textBox1"].Text = quickedit_data_textboxTxt;//waitFindWebElementByName_ToBeClickable("data", _webDriverWaitTimSpan)?.Text;//.Text屬性會Trim前空格
                            }
                            catch (Exception)
                            {
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
                catch (Exception ex)
                {
                    switch (ex.HResult)
                    {
                        case -2146233079:
                            if (ex.Message == "跨執行緒作業無效: 存取控制項 'textBox1' 時所使用的執行緒與建立控制項的執行緒不同。")
                                break;
                            else
                                goto default;
                        default:
                            Console.WriteLine(ex.HResult + ex.Message);
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                            break;
                    }
                }

                return cDrv;
            }
            else
            {
                //測試driver是否有效
                string urlChk = null;
                try
                {
                    urlChk = driver?.Url;
                }
                catch (Exception ex)
                {
                    switch (ex.HResult)
                    {

                        default:
                            Debugger.Break();
                            DriverNew();
                            break;
                    }
                }
                if (urlChk != null)
                    return driver;
                else
                {
                    if (Form1.browsrOPMode == Form1.BrowserOPMode.seleniumNew)
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請重新在textBox2下「br」或「bb」的指令以啟動chromedriver。感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主");
                    //Debugger.Break();
                    return null;
                }
            }
            #endregion
        }
        /// <summary>
        /// 重啟chromedriver
        /// </summary>
        /// <returns>重啟成功則傳回true</returns>
        internal static bool RestartDriver()
        {
            driver = null;
            killchromedriverFromHere();
            if (Form1.browsrOPMode != Form1.BrowserOPMode.seleniumNew)
                Form1.browsrOPMode = Form1.BrowserOPMode.seleniumNew;
            return DriverNew() != null;
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

            #region 遠程調試端口的設定 20241002 chrome://version/
            options.AddArgument("--remote-debugging-port=9222");//chrome://inspect/#devices 20241002 Copilot大菩薩：Word VBA 中的 Selenium 操作： https://sl.bing.net/BkGRjHWbCe
            #endregion
            return options;
        }
        /// <summary>
        /// 指定要清除quick edit box 內容的引數值 "\t"（其實是有由tab鍵所按下的值，或其他亂碼字），此與 Word VBA 中國哲學書電子化計劃.新頁面 為速新章節單位的配置有關 碼詳：https://github.com/oscarsun72/TextForCtext/blob/f75b5da5a5e6eca69baaae0b98ed2d6c286a3aab/WordVBA/%E4%B8%AD%E5%9C%8B%E5%93%B2%E5%AD%B8%E6%9B%B8%E9%9B%BB%E5%AD%90%E5%8C%96%E8%A8%88%E5%8A%83.bas#L32
        /// </summary>
        internal static readonly string chkClearQuickedit_data_textboxTxtStr = " ";
        internal static bool confirm_that_you_are_human = false;
        /// <summary>
        /// 在Chrome瀏覽器的文字框(ctext.org 的 Quick edit ）中輸入文字,creedit//若 xIuput= " "則清除而不輸入
        /// </summary>
        /// <param name="driver">chromedriver</param>
        /// <param name="xInput">要貼入的文本</param>
        /// <param name="url">要貼入的網頁網址</param>
        /// <returns>執行成功則回傳true</returns>
        internal static bool 在Chrome瀏覽器的Quick_edit文字框中輸入文字(ChromeDriver driver, string xInput, string url)
        {
            #region 檢查網址
            Uri uri = new Uri(url);
            if (uri.Authority != "ctext.org") { Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("想要輸入的網址並不是CTP網址"); return false; }
            if (driver.Url == "about:blank")
            {
                driver.Close();
                bool found = false; string urlDriver;
                driver.SwitchTo().Window(driver.WindowHandles.Last());
                for (int i = driver.WindowHandles.Count - 1; i > -1; i--)
                {
                    urlDriver = ReplaceUrl_Box2Editor(driver.Url);
                    if (urlDriver == url || url.Contains(urlDriver))
                    {
                        driver.Url = url;
                        found = true; break;
                    }
                }
                if (!found) driver.SwitchTo().Window(driver.WindowHandles.Last());
            }

            uri = new Uri(driver.Url);
            if (uri.Authority != "ctext.org") { Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("目前 driver的網址並不是CTP網址"); return false; }

            if (url.IndexOf("edit") == -1 && driver.Url.IndexOf("edit") == -1)
            {
                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("網址中不包含「edit」");
                return false;
            }


            if (url != driver.Url)
            {
                if (driver.Url.IndexOf(url.Replace("editor", "box")) == -1)
                    //if (url != driver.Url && driver.Url.IndexOf(url.Replace("editor", "box")) == -1)
                    // 使用driver導航到給定的URL
                    driver.Navigate().GoToUrl(url);
                //("https://ctext.org/library.pl?if=en&file=79166&page=85&editwiki=297821#editor");//("http://www.example.com");

                //Uri uri=new  Uri(url);

                string urlShort = url.EndsWith("#editor") ? url.Substring(0, url.IndexOf("#editor")) : url;
                if (Form1.IsValidUrl＿keyDownCtrlAdd(url) && Form1.IsValidUrl＿keyDownCtrlAdd(driver.Url) == false)
                {
                    bool found = false;
                    foreach (var item in driver.WindowHandles)
                    {
                        driver.SwitchTo().Window(item);
                        if (driver.Url.StartsWith(urlShort))
                        {
                            if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否是這個頁面要進行輸入？") == DialogResult.OK) { found = true; break; }

                        }
                    }
                    if (!found)
                    {
                        Form1.playSound(Form1.soundLike.error, true);
                        return false;
                    }
                }
                else if (Form1.IsValidUrl＿keyDownCtrlAdd(url) && Form1.IsValidUrl＿keyDownCtrlAdd(driver.Url))
                {
                    if (!driver.Url.StartsWith(urlShort))
                    {
                        bool found = false;
                        foreach (var item in driver.WindowHandles)
                        {
                            driver.SwitchTo().Window(item);
                            if (driver.Url.StartsWith(urlShort))
                            {
                                if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否是這個頁面要進行輸入？") == DialogResult.OK) { found = true; break; }

                            }
                        }
                        if (!found)
                        {
                            Form1.playSound(Form1.soundLike.error, true);
                            return false;
                        }
                    }
                }
                else
                    Debugger.Break();
            }

            #endregion

            #region 查找名稱為"data"的文字框(textbox)或ID為"quickedit"的元件，須要用到元件者均不宜另跑線程。這些名稱，都由按下 F12 或 Ctrl + shift + i 開啟開發者模式中「Elements」分頁頁籤中取得
            selm.IWebElement textbox;
            try
            {
                textbox = driver.FindElement(selm.By.Name("data"));//("textbox"));                

            }
            catch (Exception)
            {
                selm.IWebElement quickedit = null;
                try
                {
                    //如果沒有按下「Quick edit」就按下它以開啟
                    quickedit = driver.FindElement(selm.By.Id("quickedit"));
                }
                catch (Exception ex)
                {
                    switch (ex.HResult)
                    {
                        case -2146233088:
                            if (ex.Message.IndexOf("no such window: target window already closed") > -1)//"no such window: target window already closed\nfrom unknown error: web view not found\n  (Session info: chrome=110.0.5481.178)"
                            {
                                if (!url.EndsWith("#editor")) url = ActiveTabURL_Ctext_Edit_includingEditorStr;
                                GoToUrlandActivate(url);
                                return false;
                            }
                            //"no such element: Unable to locate element: {\"method\":\"css selector\",\"selector\":\"#quickedit\"}\n  (Session info: chrome=111.0.5563.147)"
                            else if (ex.Message.IndexOf("no such element: Unable to locate elementno") > -1)
                            {
                                GoToCurrentUserActivateTab();
                                quickedit = driver.FindElement(selm.By.Id("quickedit"));
                            }
                            else
                            {
                                Console.WriteLine(ex.HResult + ex.Message);
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                                Debugger.Break();
                            }
                            break;
                        default:
                            //cDrv.Navigate().GoToUrl(Form1.mainFromTextBox3Text ?? "https://ctext.org/account.pl?if=en");                    
                            //MessageBox.Show("請先登入 Ctext.org 再繼續。按下「確定(OK)」以繼續……");
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請先登入 Ctext.org 再繼續。按下「確定(OK)」以繼續……");
                            quickedit = driver.FindElement(selm.By.Id("quickedit"));
                            //throw;
                            break;
                    }
                }
                quickedit.Click();//下面「submit.Click();」不必等網頁作出回應才執行下一步，但這裡接下來還要取元件操作，就得在同一線程中跑。感恩感恩　南無阿彌陀佛
                textbox = driver.FindElement(selm.By.Name("data"));
                //throw;
            }
            quickedit_data_textboxSetting(url, textbox);

            #endregion

            ////清除原來文字，準備貼上新的
            //textbox.Clear();//20240913作廢

            #region input to textbox（old : paste to textbox）
            // 在文字框中輸入文字
            //textbox.SendKeys(@xInput); //("Hello, World!");
            /*
             chatGPT ：
                "ChromeDriver only supports characters in the BMP" 這個訊息的意思是，ChromeDriver 只支援 Unicode 基本多文種平面 (BMP) 中的字元。

                Unicode 是一種國際標準，用來對各種語言的文字進行統一編碼。它包含了超過 100,000 個字元，但是只有前 65536 個字元 (也就是基本多文種平面或 BMP) 是常用的，包括大部分的西方語言和一些亞洲語言。

                ChromeDriver 是一個 Web 自動化工具，它可以自動控制 Google Chrome 瀏覽器，執行各種測試和任務。這個訊息表示，當你在使用 ChromeDriver 時，只能輸入 BMP 中的字元。如果你想要輸入其他的字元 (比如許多亞洲語言中使用的字元)，可能會遇到問題。
             */
            //檢查是否都在BMP內
            //if (isAllinBmp(xInput))
            //{
            //textbox.SendKeys(stringtoEscape_sequences_for_Unicode_character_sets(xInput));//(Keys.Control + "v");            
            //textbox.SendKeys(xInput);
            //}
            //若含BMP外的字則用系統貼上的方法
            //else//今一律用貼上省事便捷 20230102
            //{

            ////文字框取得焦點
            //textbox.Click(); //20240913取消


            //chrome取得焦點
            //Form1 f = new Form1();
            //f.appActivateByName();

            #region 測試無誤////////……此行即可清除，不知為何多此一舉
            //////////////driver.SwitchTo().Window(driver.CurrentWindowHandle); //https://stackoverflow.com/questions/23200168/how-to-bring-selenium-browser-to-the-front#_=_
            // 讓 Chrome 瀏覽器成為作用中的程式
            //driver.Manage().Window.Maximize();//creedit chatGPT
            //driver.Manage().Window.Position = new Point(0, 0);
            #endregion

            //確定要送出文本時為true
            bool submitting = false;
            //清除內容不輸入(前已有textbox.Clear();）
            if (xInput != chkClearQuickedit_data_textboxTxtStr)//" ")// "\t")//是否清除當前頁面中的內容？（其實是有由tab鍵所按下的值)
                                                               // 建立 Actions 物件
                                                               //Actions actions = new Actions(driver);//creedit
                                                               // 貼上剪貼簿中的文字
                                                               //actions.MoveToElement(textbox).Click().Perform();
                                                               //actions.SendKeys(OpenQA.Selenium.Keys.Control + "v").Build().Perform();
                                                               //actions.SendKeys(OpenQA.Selenium.Keys.LeftShift + OpenQA.Selenium.Keys.Insert).Build().Perform();
            {
                if (quickedit_data_textboxTxt != xInput)
                    if (!SetQuickedit_data_textboxTxt(xInput))
                    {
                        ActiveForm1.textBox3Text = driver.Url;
                        if (!SetQuickedit_data_textboxTxt(xInput))
                            Debugger.Break();
                        if (Quickedit_data_textboxTxt != xInput)
                            Debugger.Break();
                        else
                            submitting = true;
                        //waitFindWebElementBySelector_ToBeClickable("#savechangesbutton")?.Click();
                    }
                    else
                        submitting = true;
                //20240913 改寫：以下作廢
                /*
                //Sendkeys(textbox, xInput);
                //while (!Form1.isClipBoardAvailable_Text()) { }
                try
                {
                    Clipboard.SetText(xInput);
                }
                catch (Exception)
                {
                    //Thread.Sleep(1500);
                    //Clipboard.Clear();
                    //Clipboard.SetText("x");
                    //Form1.playSound(Form1.soundLike.error, true);
                    //Clipboard.SetText(xInput);
                }
                //textbox.SendKeys(OpenQA.Selenium.Keys.LeftShift + OpenQA.Selenium.Keys.Insert);
                textbox.SendKeys(OpenQA.Selenium.Keys.Shift + OpenQA.Selenium.Keys.Insert);*/
            }

            //SendKeys.Send("^v{tab}~");
            #endregion
            //}
            //Task.WaitAll();
            //System.Windows.Forms.Application.DoEvents();

            //內容經過編輯才送出，否則直接翻到下一頁或停留在此頁
            if (submitting)
            {
                #region 送出


                //selm.IWebElement submit = driver.FindElement(selm.By.Id("savechangesbutton"));//("textbox"));
                selm.IWebElement submit = waitFindWebElementById_ToBeClickable("savechangesbutton", _webDriverWaitTimSpan);
                /* creedit 我問：在C#  用selenium 控制 chrome 瀏覽器時，怎麼樣才能不必等待網頁作出回應即續編處理按下來的程式碼 。如，以下程式碼，請問，如何在按下 submit.Click(); 後不必等這個動作完成或作出回應，即能繼續執行之後的程式碼呢 感恩感恩　南無阿彌陀佛
                            chatGPT他答：你可以將 submit.Click(); 放在一個 Task 中去執行，並立即返回。
                 */
                if (submit == null)
                {
                    submit = WaitFindWebElementBySelector_ToBeClickable("#savechangesbutton");
                    if (submit == null)
                    {
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請檢查頁面中的 Quict edit 是否可用，再按下確定繼續！");
                        //submit = waitFindWebElementById_ToBeClickable("savechangesbutton", _webDriverWaitTimSpan);
                        submit = driver.FindElement(By.XPath("/html/body/div[2]/div[4]/form/div/input"));
                    }
                }
                LastValidWindow = driver.CurrentWindowHandle;
            //20250218取消多線程（多執行緒操作）
            //Task.Run(() =>//接下來不用理會，也沒有元件要操作、沒有訊息要回應，就可以給另一個線程去處理了。
            //{
            reSubmit:
                try
                {
                    if (submit == null)
                        //Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請檢查頁面中的 Quict edit 是否可用，再按下確定繼續！");
                        //submit = waitFindWebElementById_ToBeClickable("savechangesbutton", _webDriverWaitTimSpan);
                        //submit = waitFindWebElementBySelector_ToBeClickable("#savechangesbutton");
                        submit = driver.FindElement(By.XPath("/html/body/div[2]/div[4]/form/div/input"));
                    if (submit == null)
                    {
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請檢查頁面中的 Quict edit 是否可用!!!!!！");
                        if (Form1.InstanceForm1.FastMode)
                            Form1.InstanceForm1.FastModeSwitcher();
                        return false;
                    }
                    if (ActiveForm1.KeyinTextMode || int.Parse(ActiveForm1.CurrentPageNum) < 3)
                        submit.Click();
                    else
                    {
                        if (!CheckPageNumBeforeSubmitSaveChanges(driver, submit))
                            return false;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.HelpLink + ex.Message);
                    //chatGPT：
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
                    if (submit != null)
                        try
                        {
                            return CheckPageNumBeforeSubmitSaveChanges(driver, submit);
                        }
                        catch (Exception)
                        {
                        }
                    else
                    {
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請手動檢查資料是否有正確送出。");
                        driver.Manage().Timeouts().PageLoad += new TimeSpan(0, 0, 3);
                        //LastValidWindow = driver.CurrentWindowHandle;
                        //openNewTabWindow();
                        try
                        {
                            driver.Navigate().GoToUrl(url);
                        }
                        catch (Exception)
                        {
                        }
                    }
                    //throw;
                }
                #region 送出後檢查是否是「Please confirm that you are human! 敬請輸入認證圖案」頁面 網址列：https://ctext.org/wiki.pl
                if (IsConfirmHumanPage())
                {
                    //Debugger.Break();
                    Form1.playSound(Form1.soundLike.waiting, true);
                    //if (ActiveForm1.FastMode) ActiveForm1.FastModeSwitcher();
                    try
                    {
                        Clipboard.SetText(xInput);//複製到剪貼簿備用
                    }
                    catch (Exception)
                    {
                    }

                    //點選輸入框
                    //waitFindWebElementBySelector_ToBeClickable("#content3 > form > table > tbody > tr:nth-child(2) > td:nth-child(2) > input[type=text]")?.Click();
                    IWebElement iweConfirm = WaitFindWebElementBySelector_ToBeClickable("#content3 > form > table > tbody > tr:nth-child(2) > td:nth-child(2) > input[type=text]");
                    if (iweConfirm == null) driver.Navigate().Back();//因非同步，若已翻到下一頁
                    iweConfirm = WaitFindWebElementBySelector_ToBeClickable("#content3 > form > table > tbody > tr:nth-child(2) > td:nth-child(2) > input[type=text]");
                    if (iweConfirm == null)
                    {
                        Debugger.Break();
                        ActiveForm1.TopMost = false;
                        driver.SwitchTo().Window(driver.CurrentWindowHandle);
                        Form1.InstanceForm1.EndUpdate();
                        return false;
                    }
                    else
                        iweConfirm.Click();
                    if (DialogResult.Cancel ==
                        Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("Please confirm that you are human! 請輸入認證圖案"
                        + Environment.NewLine + Environment.NewLine + "請輸入完畢後再按「確定」！程式會幫忙按下「OK」送出"
                        + Environment.NewLine + Environment.NewLine + "★★！最好按下「取消」以回到前數頁檢查是否有正確送出，以免白做！！", string.Empty, false))
                    {
                        //Debugger.Break();
                        ActiveForm1.TopMost = false;
                        driver.SwitchTo().Window(driver.CurrentWindowHandle);
                        //BringToFront("chrome");
                        Form1.InstanceForm1.EndUpdate();
                        return false;
                    }
                    while (true)
                    {
                        WaitFindWebElementBySelector_ToBeClickable("#content3 > form > table > tbody > tr:nth-child(3) > td:nth-child(2) > input[type=submit]")?.Click();
                        if (DialogResult.Cancel == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否重試？")) break;
                    }
                    driver.Navigate().Back();
                    while (driver.Url == "https://ctext.org/wiki.pl" || driver.Url == "https://ctext.org/wiki.pl?if=en")
                    {
                        driver.Navigate().Back();
                    }
                    if (driver.Url != url)
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("網址並非 " + url + " 請檢查後再按下確定");
                    if (driver.Url == url)
                    {
                        SetQuickedit_data_textboxTxt(xInput);
                        goto reSubmit;
                    }

                    else Debugger.Break();
                }
                #endregion
                //});


                //加速連續性輸入（不必檢視貼入的文本時，很有效）
                //if (ActiveForm1.AutoPasteToCtext && Form1.FastMode)
                //if (ActiveForm1.AutoPasteToCtext && Form1.fastMode && Form1.browsrOPMode == Form1.BrowserOPMode.appActivateByName)
                if (ActiveForm1.AutoPasteToCtext && Form1.InstanceForm1.FastMode && Form1.browsrOPMode == Form1.BrowserOPMode.appActivateByName)
                {
                    Thread.Sleep(10);//等待 submit = waitFin……完成
                    driver.Close(); //需要重啟檢視時，只要開啟前一個被關掉的分頁頁籤即可（快速鍵時 Ctrl + Shift + t）
                }
                #endregion
            }
            else//若文本沒有改變，不用送出，則播放音效
                Form1.playSound(Form1.soundLike.notify, true);
            return true;
        }
        /// <summary>
        /// 在按下
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="submit_saveChanges"></param>
        /// <returns></returns>
        internal static bool CheckPageNumBeforeSubmitSaveChanges(ChromeDriver driver, IWebElement submit_saveChanges = null)
        {
            if (!IsDriverInvalid() && int.Parse(ActiveForm1.CurrentPageNum) > 2)
            {
                int currentPageNum = int.Parse(Form1.InstanceForm1.CurrentPageNum);
                if (ActiveForm1.AutoPasteToCtext && currentPageNum != Form1.InstanceForm1.GetPageNumFromUrl(driver.Url) ||
                    Math.Abs(int.Parse(ActiveForm1.CurrentPageNum) - int.Parse(WindowHandles["currentPageNum"])) != 1)
                {
                    if (DialogResult.OK == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("頁碼不同！請轉至頁面" +
                        "頁再按下「確定」以供輸入"))
                    {
                        submit_saveChanges?.Click();//按下 Save changes button（「保存編輯」按鈕）
                        return true;
                    }
                    else
                        return false;
                }
                else
                {
                    Form1.InstanceForm1.PauseEvents();
                    submit_saveChanges?.Click();//按下 Save changes button（「保存編輯」按鈕）
                    Form1.InstanceForm1.ResumeEvents();
                    return true;
                }
            }
            else
                return false;
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
        /// 取得現行Ctext 編輯時前景之分頁網址。尤其是為使用者手動切換者；若找不到則傳回""（空字串）
        /// </summary>
        public static string ActiveTabURL_Ctext_Edit
        {
            get
            {
                //string url = getUrl(ControlType.Edit).Trim();
                string url = getUrlFirst_Ctext_Edit(ControlType.Edit).Trim();
                if (url == "")
                {
                    try
                    {
                        string urlDriver = driver.Url;
                    }
                    catch (Exception)
                    {
                        if (Form1.IsValidUrl＿keyDownCtrlAdd(ActiveForm1.textBox3Text))
                        {
                            //如：https://ctext.org/library.pl?if=en&file=38675&page=1&editwiki=573099#editor
                            string shortUrl = ActiveForm1.textBox3Text.Substring(0, ActiveForm1.textBox3Text.IndexOf("#editor") == -1 ? ActiveForm1.textBox3Text.Length : ActiveForm1.textBox3Text.IndexOf("#editor"));
                            for (int i = driver.WindowHandles.Count - 1; i > -1; i--)
                            {
                                driver.SwitchTo().Window(driver.WindowHandles[i]);
                                if (driver.Url.StartsWith(shortUrl)) break;
                            }
                        }
                    }

                    if (!Form1.IsValidUrl＿ImageTextComparisonPage(driver.Url))
                    {
                        for (int i = driver.WindowHandles.Count - 1; i > -1; i--)
                        {
                            driver.SwitchTo().Window(driver.WindowHandles[i]);
                            if (Form1.IsValidUrl＿ImageTextComparisonPage(driver.Url)) break;
                        }
                        if (!Form1.IsValidUrl＿ImageTextComparisonPage(driver.Url))
                        {
                            int windowsCount;// = 0;
                            try
                            {
                                windowsCount = driver.WindowHandles.Count;
                            }
                            catch (Exception)
                            {
                                windowsCount = GetValidWindowHandles(driver).Count;
                            }
                            if (windowsCount > 1)
                            {
                                if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("目前作用中的分頁並非有效的圖文對照頁面，是否要讓程式繼續比對？"
                                        , "ActiveTabURL_Ctext_Edit\n\r\n\rgetUrlFirst_Ctext_Edit=\"\"") == DialogResult.OK)
                                    url = getUrl(ControlType.Edit).Trim();
                            }
                        }
                        else
                        {
                            url = driver.Url;
                            ActiveForm1.textBox3Text = url;
                        }
                    }
                    else
                    {
                        url = driver.Url;
                        ActiveForm1.textBox3Text = url;
                    }
                }
                if (url != "") url = url.StartsWith("https://") ? url : "https://" + url;
                return url;
            }
        }
        /// <summary>
        /// 取得現行Ctext 編輯時前景之分頁網址（須含有"#editor"尾綴）。尤其是為使用者手動切換者；若找不到則傳回""（空字串）
        /// </summary>
        public static string ActiveTabURL_Ctext_Edit_includingEditorStr
        {
            get
            {
                string url = getUrlFirst_Ctext_Edit(ControlType.Edit, true).Trim();
                if (url == "")
                {
                    if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("目前作用中的分頁並非有效的圖文對照頁面，是否要讓程式繼續比對？") == DialogResult.OK)
                    {

                        url = getUrl(ControlType.Edit).Trim();
                    }
                }
                if (url != "") url = url.StartsWith("https://") ? url : "https://" + url;
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
        internal static Form1 ActiveForm1 { get => frm ?? Application.OpenForms[0] as Form1; set => frm = value; }

        /// <summary>
        /// 取得「簡單修改模式」的網址
        /// </summary>
        /// <returns>傳回「簡單修改模式」的網址</returns>
        internal static string GetQuickeditUrl()
        {
            string url = "";
            if (driver == null) driver = DriverNew();
            IWebElement ie = QuickeditLinkIWebElement;
            if (ie != null) url = ie.GetAttribute("href");
            return url;
            /*
             OpenQA.Selenium.IWebElement quickEditLink = br.
                 waitFindWebElementBySelector_ToBeClickable("#quickedit > a");
                    if (quickEditLink != null)
                    {
                        quickEditLinkUrl = quickEditLink.GetAttribute("href");
                    }
             */
        }
        /// <summary>
        /// 傳回Chrome瀏覽器作用中分頁視窗頁籤url 
        /// 20250410 GitHub　Copilot大菩薩
        /// </summary>
        /// <returns></returns>
        public static string GetActiveChromeTabUrl()
        {
            // 找到 Chrome 主視窗
            IntPtr chromeHandle = FindWindow("Chrome_WidgetWin_1", null);
            if (chromeHandle == IntPtr.Zero)
            {
                Console.WriteLine("未找到 Chrome 瀏覽器！");
                return null;
            }

            try
            {
                AutomationElement rootElement = AutomationElement.FromHandle(chromeHandle);
                if (rootElement == null)
                {
                    Console.WriteLine("無法取得 Chrome 的 AutomationElement！");
                    return null;
                }

                // 列舉所有子元素
                AutomationElementCollection children = rootElement.FindAll(TreeScope.Descendants, Condition.TrueCondition);
                foreach (AutomationElement child in children)
                {
                    Console.WriteLine($"Name: {child.Current.Name}, ControlType: {child.Current.ControlType.ProgrammaticName}");
                }

                // 嘗試找到地址列
                AutomationElement addressBar = rootElement.FindFirst(TreeScope.Descendants,
                      new PropertyCondition(AutomationElement.NameProperty, "地址和搜尋列")); // 根據實際名稱調整
                                                                                        //new PropertyCondition(AutomationElement.NameProperty, "Address and search bar")); // 根據實際名稱調整
                if (addressBar != null)
                {
                    ValuePattern valuePattern = addressBar.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                    return valuePattern?.Current.Value;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"發生錯誤: {ex.Message}");
            }

            return null;
        }
        /// <summary>
        /// 用 UI Automation 獲取 Chrome 瀏覽器中所有分頁的 URL
        /// 20250410 GitHub　Copilot大菩薩
        /// </summary>
        public static void GetAllChromeTabUrls()
        {
            // 找到 Chrome 主視窗
            IntPtr chromeHandle = FindWindow("Chrome_WidgetWin_1", null);
            if (chromeHandle == IntPtr.Zero)
            {
                Console.WriteLine("未找到 Chrome 瀏覽器！");
                return;
            }

            try
            {
                // 獲取 Chrome 的 AutomationElement
                AutomationElement rootElement = AutomationElement.FromHandle(chromeHandle);
                if (rootElement == null)
                {
                    Console.WriteLine("無法取得 Chrome 的 AutomationElement！");
                    return;
                }

                // 列舉所有子元素
                AutomationElementCollection children = rootElement.FindAll(TreeScope.Descendants, Condition.TrueCondition);
                foreach (AutomationElement child in children)
                {
                    Console.WriteLine($"ControlType: {child.Current.ControlType.ProgrammaticName}");
                    Console.WriteLine($"Name: {child.Current.Name}");
                    Console.WriteLine($"AutomationId: {child.Current.AutomationId}");
                    Console.WriteLine($"LocalizedControlType: {child.Current.LocalizedControlType}");

                    // 列舉所有屬性
                    foreach (AutomationProperty property in child.GetSupportedProperties())
                    {
                        object value = child.GetCurrentPropertyValue(property);
                        Console.WriteLine($"Property: {property.ProgrammaticName}, Value: {value}");
                    }

                    Console.WriteLine("--------------------------------------------------");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"發生錯誤: {ex.Message}");
            }
        }

        /// <summary>
        /// geturl 修改後的程式碼:20230308 creedit with NotionAI大菩薩
        /// 〈get url FindAll vs FindFirst〉https://www.notion.so/get-url-FindAll-vs-FindFirst-88505499d53e4557a45fe8e844f0ee4a
        /// </summary>
        /// <param name="controlType"></param>
        /// <param name="endwithEditorStr">是否要取得末綴為「#editor」的網址</param>
        /// <returns></returns>
        static string getUrlFirst_Ctext_Edit(ControlType controlType, bool endwithEditorStr = false)
        {
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
                        /* 20230313 AutomationElement 可以用using (){} 來寫嗎：Bing大菩薩：
                         * `AutomationElement` 不實現 `IDisposable`，因為它們沒有非托管資源。只要刪除了對它的所有引用，內存就會被釋放 - 例如將每個指針設置為空或變量超出範圍⁷。
                         * 您所謂的「指針」就是指參考型別的變數，而「變量」是指實質型別的變數吧？設置為空，就是「=null」吧 是嗎？                         * 
                         * 是的，您說得對。在 C# 中，「指針」通常指的是參考型別的變數，而「變量」可以指實質型別或參考型別的變數。將一個參考型別的變數設置為空，就是將它賦值為 null。

                            來源: 與 Bing 的交談， 2023/3/13(1) When does AutomationElement get disposed?. https://social.msdn.microsoft.com/Forums/windowsdesktop/en-US/105794db-fa1a-41d0-827a-6993973abde9/when-does-automationelement-get-disposed?forum=windowsaccessibilityandautomation 已存取 2023/3/13.
                            (2) Invoke a Control Using UI Automation - .NET Framework. https://learn.microsoft.com/en-us/dotnet/framework/ui-automation/invoke-a-control-using-ui-automation 已存取 2023/3/13.
                            (3) Obtaining UI Automation Elements - .NET Framework. https://learn.microsoft.com/en-us/dotnet/framework/ui-automation/obtaining-ui-automation-elements 已存取 2023/3/13.
                            (4) How can I send a right-click event to an AutomationElement using WPF's UI automation? - Stack Overflow. https://stackoverflow.com/questions/6554494/how-can-i-send-a-right-click-event-to-an-automationelement-using-wpfs-ui-automa 已存取 2023/3/13.
                            (5) Invoke a Control Using UI Automation - .NET Framework. https://learn.microsoft.com/en-us/dotnet/framework/ui-automation/invoke-a-control-using-ui-automation 已存取 2023/3/13.
                            (6) Using objects that implement IDisposable | Microsoft Learn. https://learn.microsoft.com/en-us/dotnet/standard/garbage-collection/using-objects 已存取 2023/3/13.
                            (7) Obtaining UI Automation Elements - .NET Framework. https://learn.microsoft.com/en-us/dotnet/framework/ui-automation/obtaining-ui-automation-elements 已存取 2023/3/13.
                            (8) AutomationElement Class (System.Windows.Automation). https://learn.microsoft.com/en-us/dotnet/api/system.windows.automation.automationelement?view=windowsdesktop-8.0 已存取 2023/3/13.
                            (9) Obtaining UI Automation Elements - .NET Framework. https://learn.microsoft.com/en-us/dotnet/framework/ui-automation/obtaining-ui-automation-elements 已存取 2023/3/13.
                            (10) c# - selecting combobox item using ui automation - Stack Overflow. https://stackoverflow.com/questions/5814779/selecting-combobox-item-using-ui-automation 已存取 2023/3/13.
                         */
                        AutomationElement elm = AutomationElement.FromHandle(proc.MainWindowHandle);
                        AutomationElement elmUrlBar = elm.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, controlType));

                        if (elmUrlBar != null)
                        {
                            string url = ((ValuePattern)elmUrlBar.GetCurrentPattern(ValuePattern.Pattern)).Current.Value as string;
                            //if ((url.StartsWith("http") || url.StartsWith("ctext")))
                            if (endwithEditorStr)
                            {
                                if ((url.StartsWith("ctext.org/") || url.StartsWith("https://ctext.org/")) && url.IndexOf("&file=") > -1 && url.IndexOf("&page=") > -1 && url.EndsWith("#editor"))
                                {
                                    return url;
                                }
                            }
                            else
                            {
                                if ((url.StartsWith("ctext.org/") || url.StartsWith("https://ctext.org/")) && url.IndexOf("&file=") > -1 && url.IndexOf("&page=") > -1)//&& url.EndsWith("#editor"))
                                {
                                    return url;
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                // Ignore exception
            }
            return "";//url;

        }
        /// <summary>
        /// 取得Chrome瀏覽器現前作用中的《中國哲學書電子化計劃》頁面網址
        /// </summary>
        public static string GetChromeActiveUrl
        {
            get { return getUrlFirst_Ctext_Edit(ControlType.Edit).Trim(); }
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
            string urls = ""; Form1.playSound(Form1.soundLike.over);
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

        /// <summary>
        /// 切換到目前現在正在作用中的（程式抓到最前端的）分頁
        /// </summary>
        /// <param name="textBox3"></param>
        /// <param name="urlActiveTab"></param>
        /// <returns>若切換失敗，即回傳false</returns>
        internal static bool SwitchToCurrentForeActivateTab(ref TextBox textBox3, string urlActiveTab = "")
        {
            string url = "";
            try
            {
                url = driver.Url; urlActiveTab = urlActiveTab == "" ? ActiveTabURL_Ctext_Edit : urlActiveTab;
            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case -2146233088:
                        if (ex.Message.IndexOf("no such window: target window already closed") > -1)//"no such window: target window already closed\nfrom unknown error: web view not found\n  (Session info: chrome=110.0.5481.178)"
                        {
                            driver.SwitchTo().Window(driver.WindowHandles[0]);
                            url = driver.Url; if (urlActiveTab == "") urlActiveTab = ActiveTabURL_Ctext_Edit;
                        }
                        else
                            MessageBox.Show(ex.HResult + ex.Message);
                        break;
                    default:
                        MessageBox.Show(ex.HResult + ex.Message);
                        break;
                }
            }

            if (url != urlActiveTab//如果現行的頁面不同於程式所在之頁面
                && url.Length > 8 && url.Substring(8) != urlActiveTab)//會有省略「https://」者
            {
                LastValidWindow = driver.CurrentWindowHandle;
                if (GoToCurrentUserActivateTab(urlActiveTab) == string.Empty)//那麼就將程式所在之頁面轉到、設定為現行的前景頁面
                {
                    driver.SwitchTo().Window(LastValidWindow);
                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("程式找到的作用中的分頁並不在chromedriver中，請手動檢查，是否是另外開了一個Chrome瀏覽器實例。");
                    return false;
                }
            }
            if (textBox3.Text != "" && textBox3.Text != urlActiveTab && urlActiveTab.StartsWith("http")) textBox3.Text = urlActiveTab;//如果textBox3非空值且與現行頁面網址不合，亦轉設為現行前景頁面之網址}
            return true;
        }

        /// <summary>
        /// 將程式所在頁面轉到現行前景的分頁頁面
        /// </summary>
        /// <param name="urlActiveTab">若已取得現行前景分頁頁面之網址則作此引數傳入，免得再取一次，徒耗資源，減損效能</param>
        /// <returns>傳回目前作用中的分頁頁籤網址字串值；如果無匹配者，即網址不存在於目前的chromedriver中則傳回空字串""</returns>
        internal static string GoToCurrentUserActivateTab(string urlActiveTab = "")
        {
            if (urlActiveTab == "") urlActiveTab = ActiveTabURL_Ctext_Edit;
            return urlActiveTab;
            #region 現在有設定Chrome瀏覽器的啟動參數 --remote-debugging-port=9222 了，這個可以省了!!●●●●●●●●●真的可以省了

            string url = "";
            if (urlActiveTab != "")
            {
                try
                {
                    //現在有設定Chrome瀏覽器的啟動參數 --remote-debugging-port=9222 了，這個可以省了!!●●●●●●●●●真的可以省了
                    if (driver == null && Form1.browsrOPMode == Form1.BrowserOPMode.appActivateByName)
                    {
                        if (DialogResult.OK == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("請先在textBox2執行「br」指令，切換為SeleniumNew模式再繼續。" +
                            Environment.NewLine + Environment.NewLine + "---若已安裝配置器好chromedriver.exe，請按下【取消】，則程式會自動啟動，無須手動下指令--- 20241218感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主"))
                            return "";
                        else
                        {
                            if (driver == null)
                            {
                                Form1.browsrOPMode = Form1.BrowserOPMode.seleniumNew;
                                DriverNew();
                                driver.Close();//關閉新開啟的「首頁」分類頁籤，網址是： https://ctext.org/account.pl?if=en
                            }
                            else
                                RestartChromedriver();
                        }
                    }
                    try
                    {
                        url = GetDriverUrl;
                        //url = driver.Url;
                    }
                    catch (Exception ex)
                    {
                        switch (ex.HResult)
                        {
                            case -2146233088:
                                if (ex.Message.StartsWith("no such window: target window already closed"))
                                {
                                    Form1.playSound(Form1.soundLike.error);
                                    driver.SwitchTo().Window(driver.WindowHandles[driver.WindowHandles.Count - 1]);
                                    url = driver.Url;
                                }
                                else if (ex.Message.StartsWith(@"An unknown exception was encountered sending an HTTP request to the remote WebDriver server for URL http:/"))
                                {
                                    return string.Empty;
                                }
                                else
                                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                                break;
                            default:
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                                break;
                        }

                    }
                }
                catch (Exception ex)
                {
                    switch (ex.HResult)
                    {
                        case -2146233088:
                            if (ex.Message.IndexOf("no such window: target window already closed") > -1) //"no such window: target window already closed\nfrom unknown error: web view not found\n  (Session info: chrome=110.0.5481.178)"
                            {
                                openNewTabWindow();
                                url = urlActiveTab;
                                driver.Navigate().GoToUrl(url);
                                //driver.SwitchTo().Window(driver.WindowHandles[0]);
                            }
                            else if (ex.Message.StartsWith("not connected to DevTools"))//-2146233088disconnected: not connected to DevTools
                                                                                        //(failed to check if window was closed: disconnected: not connected to DevTools)
                                RestartChromedriver();
                            else if (ex.Message.StartsWith("disconnected: not connected to DevTools"))//disconnected: not connected to DevTools\n  (failed to check if window was closed: disconnected: not connected to DevTools)\n  (Session info: chrome=130.0.6723.117)
                            {
                                RestartChromedriver();
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

                //會有這種情形： url= "https://ctext.org/library.pl?if=gb&file=34096&page=29#%E6%98%93"
                //urlActiveTab= "https://ctext.org/library.pl?if=gb&file=34096&page=29#易"
                //即一用中文，一為編碼
                urlActiveTab = urlActiveTab.IndexOf("#") == -1 ? urlActiveTab : urlActiveTab.Substring(0, urlActiveTab.IndexOf("#"));
                url = url.IndexOf("#") == -1 ? url : url.Substring(0, url.IndexOf("#"));
                if (urlActiveTab != url)
                {
                    bool found = false;
                    //foreach (var item in driver.WindowHandles)
                    //{
                    for (int i = driver.WindowHandles.Count - 1; i > -1; i--)
                    {
                        driver.SwitchTo().Window(driver.WindowHandles[i]);
                        url = driver.Url; url = url.IndexOf("#") == -1 ? url : url.Substring(0, url.IndexOf("#"));
                        if (urlActiveTab == url) { found = true; break; }
                    }
                    if (!found) return string.Empty;
                }
            }
            if (url == "" && urlActiveTab == "")
                return string.Empty;//driver.Url;
            else
                return url == "" ? urlActiveTab : url;

            #endregion
        }
        /// <summary>
        /// Selenium 瀏覽所指定的網址所在的網頁
        /// </summary>
        /// <param name="url">要瀏覽的網址</param>
        /// <param name="frmKeyinTextModeTopWindow">是否將視窗內容可見位置調到最頂端</param>
        /// <returns>失敗則傳回false</returns>
        internal static bool GoToUrlandActivate(string url, bool frmKeyinTextModeTopWindow = false)
        {
            if (string.IsNullOrEmpty(url) || url.Length < 7 || url.Substring(0, 4) != "http") return false;

            ////driver.Close();//creedit
            ////creedit20230103 這樣處理誤關分頁頁籤的錯誤（例外情形）就成功了，但整個瀏覽器誤關則尚未
            ////chatGPT：在 C# 中使用 Selenium 取得 Chrome 瀏覽器開啟的頁籤（分頁）數量可以使用以下方法：                
            int tabCount = 0;
            try
            {
                if (driver == null) driver = DriverNew();
            }
            catch (Exception)
            {
                if (driver != null)
                {
                    driver = null;
                }
                driver = DriverNew();
                ////throw;
            }
        /*另外，您也可以使用以下方法在 C# 中取得 Chrome 瀏覽器的標籤（分頁）數量:
         // 取得 Chrome 瀏覽器的標籤數量
            int tabCount = driver.Manage().Window.Bounds.Width / 100;
         */
        retry:
            try
            {
                driver = driver ?? Browser.DriverNew();
                if (driver != null)
                    tabCount = driver.WindowHandles.Count;
                else
                    return false;
            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case -2146233088:
                        if (ex.Message.StartsWith("The HTTP request to the remote WebDriver server for URL")
                            //"The HTTP request to the remote WebDriver server for URL http://localhost:4144/session/a5d7705c0a6199c76529de0e157667f9/window/handles timed out after 8.5 seconds."
                            || ex.Message.StartsWith("disconnected: not connected to DevTools"))//"disconnected: not connected to DevTools\n  (failed to check if window was closed: disconnected: not connected to DevTools)\n  (Session info: chrome=128.0.6613.138)"

                        {
                            killProcesses(new string[] { "chromedriver" });//手動關閉由Selenium啟動的Chrome瀏覽器須由此才能清除
                            driver = null;
                            driver = DriverNew();
                            if (driver != null) tabCount = driver.WindowHandles.Count;
                            else return false;

                        }
                        else if (ex.Message.StartsWith("An unknown exception was encountered sending an HTTP request to the remote WebDriver server for URL"))
                        {
                            killchromedriverFromHere();
                            Form1.playSound(Form1.soundLike.error, true);
                            //Debugger.Break();
                            driver = null; Form1.browsrOPMode = Form1.BrowserOPMode.seleniumNew;
                            DriverNew();
                            goto retry;
                        }
                        else if (ex.Message.StartsWith("invalid session id"))
                        {
                            Form1.playSound(Form1.soundLike.error, true);
                            RestartChromedriver();
                            goto retry;
                        }

                        else
                        {
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                            Console.WriteLine(ex.HResult + ex.Message);
                            Debugger.Break();
                        }
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
                    //activate Chrome瀏覽器
                    driver = driver ?? Browser.DriverNew();
                    driver.SwitchTo().Window(driver.CurrentWindowHandle);

                }
                catch (Exception ex)
                {
                    switch (ex.HResult)
                    {
                        //操作中的分頁頁籤被手動誤關時
                        //no such window: target window already closed
                        case -2146233088:
                            if (ex.Message.StartsWith("invalid session id"))
                                RestartChromedriver();
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
            try
            {
                if (driver.Url != url)
                    driver.Navigate().GoToUrl(url);
                //LastValidWindow = driver.CurrentWindowHandle;//●●●●●●●●●●●●●●●●●●●● 20250408取消
                //activate and move to most front of desktop
                //driver.SwitchTo().Window(driver.CurrentWindowHandle;
                if (frmKeyinTextModeTopWindow) WindowsScrolltoTop();//將分頁視窗頁面捲到頂端
                quickedit_data_textboxSetting(url);
            }
            catch (Exception ex)
            {
                //忽略錯誤不處理
                if (ex.Message.StartsWith("invalid session id"))
                    RestartChromedriver();
            }
            return true;
        }

        /// <summary>
        /// 將分頁視窗頁面捲到頂端
        /// activate and move to most front of desktop
        /// </summary>
        internal static void WindowsScrolltoTop()
        {
            if (ActiveForm1.KeyinTextMode)
            {
                //if (Form1.ModifierKeys != forms.Keys.LControlKey) return;

                if (ActiveForm1.InvokeRequired)
                {
                    ActiveForm1.Invoke((MethodInvoker)delegate
                    {
                        // 你的程式碼

                        //此法有效！！20231019
                        bool _events = ActiveForm1.EventsEnabled;
                        ActiveForm1.PauseEvents();
                        driver.ExecuteScript("window.scrollTo(0, 0)");//chatGPT:您好！如果您使用 C# 和 Selenium 來控制 Chrome 瀏覽器，您可以使用以下的程式碼將網頁捲到最上面：
                        ActiveForm1.EventsEnabled = _events;
                    });
                }
                else
                {//此法有效！！20231019
                    bool _events = ActiveForm1.EventsEnabled;
                    ActiveForm1.PauseEvents();
                    driver.ExecuteScript("window.scrollTo(0, 0)");//chatGPT:您好！如果您使用 C# 和 Selenium 來控制 Chrome 瀏覽器，您可以使用以下的程式碼將網頁捲到最上面：
                    ActiveForm1.EventsEnabled = _events;
                }
                //driver.SwitchTo().Window(driver.CurrentWindowHandle).SwitchTo().DefaultContent();//20231019

                ///*20220312 chatGPT大菩薩：您好！要將分頁視窗的瀏覽位置調整到最上方，可以使用 Selenium 的 JavaScriptExecutor 物件，透過執行 JavaScript 的方式來操作瀏覽器。
                //    以下是一個使用 JavaScriptExecutor 調整瀏覽位置的範例：*/
                //// 取得目前的分頁視窗物件
                //IWebElement body = driver.FindElement(By.TagName("body"));
                //var data = waitFindWebElementByName_ToBeClickable("data", 10);
                //if (data != null)
                //{
                //    // 使用 WebDriverWait 等待目標元素出現
                //    //var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                //    //var body = wait.Until(ExpectedConditions.ElementExists(By.TagName("body")));

                //    //IWebElement body = waitFindWebElementByTagName_ToBeClickable("body",10);

                //    // 使用 JavaScriptExecutor 將分頁視窗的瀏覽位置調整到最上方
                //    IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                //    js.ExecuteScript("arguments[0].scrollTop = 0;", body);
                //    /*程式碼會使用 FindElement 方法取得分頁視窗的 body 元素，
                //     * 然後再透過 JavaScriptExecutor 的 ExecuteScript 方法來執行 JavaScript，將分頁視窗的 scrollTop 屬性設定為 0，即將瀏覽位置調整到最上方。
                //     */

                //    //這兩種方法都可以將瀏覽位置調整到最上方，但有時候可能會因為某些原因無法正常工作。
                //    //建議您可以嘗試使用以下程式碼來調整瀏覽位置，這個方法使用了 window.scrollBy(0, -document.body.scrollHeight)，可以將瀏覽位置調整到最上方。
                //    js.ExecuteScript("window.scrollBy(0, -document.body.scrollHeight)");
                //    //如果這個方法仍然無法正常工作，建議您再檢查一下網頁上是否有其他的元素會影響瀏覽位置，例如 iframe 或 div。如果有的話，您需要先切換到相應的 iframe 或找到影響瀏覽位置的 div，然後再使用 JavaScriptExecutor 調整瀏覽位置。
                //}
            }
            else
            {
                //driver.SwitchTo().Window(driver.CurrentWindowHandle).SwitchTo().DefaultContent();//20231019
                //driver.Navigate().GoToUrl(driver.Url);
                ////以上都不行，只好來這個殺手鐧了
                //driver.Navigate().Refresh();//https://www.guru99.com/selenium-refresh-page.html
                //Pressing F5 key on any textbox using sendKeys command
                //driver.findElement(By textboxLocator).sendKeys(Keys.F5);
                //data.SendKeys(OpenQA.Selenium.Keys.F5);
                //https://artoftesting.com/refresh-a-page-in-selenium-webdriver-java
                driver.ExecuteScript("window.scrollTo(0, 0)");//chatGPT:您好！如果您使用 C# 和 Selenium 來控制 Chrome 瀏覽器，您可以使用以下的程式碼將網頁捲到最上面：
            }
        }

        private static void quickedit_data_textboxSetting(string url, IWebElement textbox = null, IWebDriver driver = null)
        {
            if (url.IndexOf("edit") > -1)
            {
                if (textbox != null) Quickedit_data_textbox = textbox;
                else
                    try
                    {
                        Quickedit_data_textbox = waitFindWebElementByName_ToBeClickable("data", _webDriverWaitTimSpan, driver);
                    }
                    catch (Exception ex)
                    {
                        //分頁視窗若關閉則忽略、繼續
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.Message);
                        return;
                        //throw;
                    }
                //Quickedit_data_textbox = waitFindWebElementByName_ToBeClickable("data", _webDriverWaitTimSpan, driver);
                try
                {
                    //.Text屬性會清除起首的全形空格！！20240313
                    //quickedit_data_textboxTxt = Quickedit_data_textbox == null ? "" : Quickedit_data_textbox.Text;
                    if (Quickedit_data_textbox == null)
                        quickedit_data_textboxTxt = "";
                    else
                    {
                        quickedit_data_textboxTxt = Quickedit_data_textboxTxt;
                    }

                }
                catch (Exception)
                {
                    quickedit_data_textboxTxt = string.Empty;
                }
            }
        }

        internal static ChromeDriver openNewTabWindow(WindowType tabOrwindow = WindowType.Tab)//creedit 20230103
        {/*chatGPT
            在 C# 中使用 Selenium 開啟新 Chrome 瀏覽器分頁可以使用以下方法：*/
            // 創建 ChromeDriver 實例
            //IWebDriver driver = new ChromeDriver();
            //ChromeDriver driver = driverNew();//new ChromeDriver();
            if (driver == null)
            {
                if (Form1.browsrOPMode == Form1.BrowserOPMode.appActivateByName)
                    Form1.browsrOPMode = Form1.BrowserOPMode.seleniumNew;
                driver = DriverNew();
            }
            try
            {
                //LastValidWindow = driver.CurrentWindowHandle;//●●●●●●●●●●●●●●●●●●●●● 20250408取消
                //Form1.ResetLastValidWindow();
                driver = (ChromeDriver)driver.SwitchTo().NewWindow(tabOrwindow);

            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case -2146233088:
                        if (ex.Message.StartsWith(@"An unknown exception was encountered sending an HTTP request to the remote WebDriver server for URL http://localhost:"))
                        {
                            RestartDriver();
                        }
                        else
                            goto default;
                        break;
                    default:
                        ////Console.WriteLine(ex.HResult + ex.Message);
                        try
                        {
                            var hs = driver.WindowHandles;
                            //driver.SwitchTo().Window(driver.WindowHandles.Last());
                            driver.SwitchTo().Window(LastValidWindow);
                            driver.SwitchTo().NewWindow(tabOrwindow);
                        }
                        catch (Exception ex1)
                        {
                            switch (ex1.HResult)
                            {
                                case -2146233079://"序列未包含項目"
                                                 //誤關Chrome瀏覽器的時候
                                                 //openNewTabWindow(WindowType.Window);
                                    RestartDriver();
                                    break;
                                case -2146233088:
                                    if (ex1.Message.StartsWith(@"An unknown exception was encountered sending an HTTP request to the remote WebDriver server for URL http://localhost:"))
                                    {
                                        RestartDriver();
                                    }
                                    else if (ex1.Message.StartsWith("no such window"))
                                    //可能是按下擴充功能故 20231228
                                    {
                                        BringToFront("chrome");
                                        if (!ActiveForm1.Active)
                                            SendKeys.SendWait("{esc}");
                                        driver.SwitchTo().Window(driver.WindowHandles.Last());
                                        LastValidWindow = driver.WindowHandles.Last();
                                        try
                                        {
                                            driver.SwitchTo().NewWindow(tabOrwindow);
                                        }
                                        catch (Exception)
                                        {
                                            Form1.playSound(Form1.soundLike.error, true);
                                            Form1.playSound(Form1.soundLike.warn, true);
                                            //Copilot大菩薩 20240727：
                                            ((IJavaScriptExecutor)driver).ExecuteScript("window.open();");
                                            driver.SwitchTo().Window(driver.WindowHandles.Last());
                                        }
                                    }
                                    else if (ex1.Message.StartsWith("invalid session id"))
                                        RestartChromedriver();
                                    else
                                        goto default;
                                    break;
                                default:
                                    Console.Write(ex1.HResult + ex1.Message);
                                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex1.HResult + ex1.Message);
                                    break;
                            }
                        }
                        break;
                }
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
        retry:
            if (Form1.browsrOPMode == Form1.BrowserOPMode.appActivateByName) Form1.browsrOPMode = Form1.BrowserOPMode.seleniumNew;
            if (driver == null) driver = DriverNew();
            //using (driver)//var driver = new ChromeDriver())//若這樣寫則會出現「無法存取已處置的物件。」之錯誤    HResult	-2146232798	int               
            //{因為 using(driver) 這 driver 只在 ) 後的第一層大括弧{}間有效，生命週期僅止於此間而已
            // 移動到指定的網頁
            try
            {
                //url = url ?? System.Windows.Forms.Application.OpenForms[0].Controls["textBox3"].Text;
                url = url ?? ActiveForm1.textBox3Text;

                //if (IsSameBookPageWithDrive(url) || ActiveForm1.GetEditwikiID_fromUrl(url) != ActiveForm1.PreviousEditwikiID) GoToUrlandActivate(url, true);
                if (!IsSameBookPageWithDrive(url)) GoToUrlandActivate(url, true);
            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case -2146233088:
                        if (ex.Message.IndexOf("timed out after ") > -1) return "";
                        else if (ex.Message.StartsWith("no such window: target window already closed"))
                        {
                            ReadOnlyCollection<string> driverWindowHandles = null;
                            try
                            {
                                driverWindowHandles = driver.WindowHandles;
                            }
                            catch (Exception)
                            {
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                            }
                            if (driverWindowHandles != null)
                            {
                                foreach (var item in driverWindowHandles)
                                {
                                    if (driver.SwitchTo().Window(item).Url == url)
                                    {
                                        if (!driverWindowHandles.Contains(LastValidWindow))
                                            LastValidWindow = driver.CurrentWindowHandle;
                                        if (!IsSameBookPageWithDrive(url)) GoToUrlandActivate(url, true);
                                    }
                                    break;
                                }
                            }
                            break;
                        }
                        else
                        {
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                            break;

                        }
                    default:
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                        break;
                }
                //throw;
            }
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
            string imageUrl = ""; IList<IWebElement> imageElements;
            try
            {
                imageElements = driver.FindElements(By.TagName("img"));
            }
            catch (Exception)
            {
                return imageUrl;
                //throw;
            }
            foreach (IWebElement imageElement in imageElements)
            {
                try
                {
                    imageUrl = imageElement.GetAttribute("src");
                }
                catch (Exception ex)
                {
                    switch (ex.HResult)
                    {
                        case -2146233088:
                            if (ex.Message.IndexOf("stale element reference: stale element not found") > -1)
                                goto retry;
                            else
                                throw;
                        default:
                            throw;
                    }
                }

                if (imageUrl.IsNullOrEmpty()) continue;
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
        /// <summary>
        /// 判斷是否與目前的drive在同一本書的同一頁
        /// </summary>
        /// <param name="url">要比對的網址</param>
        /// <returns></returns>
        internal static bool IsSameBookPageWithDrive(string url)
        {
            int bookidDrive = ActiveForm1.GetBookID_fromUrl(driver?.Url ?? string.Empty), pageNumDrive = ActiveForm1.GetPageNumFromUrl(driver.Url), bookid = ActiveForm1.GetBookID_fromUrl(url), pageNum = ActiveForm1.GetPageNumFromUrl(url);
            //if (bookidDrive != bookid && pageNumDrive != pageNum)
            if (bookidDrive == bookid && pageNumDrive == pageNum)
                return true;
            else return false;
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
        /// <summary>
        /// 由Url判斷是否是[簡單修改模式][Quick edit] 
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        internal static bool isQuickEditUrl(string url)
        {
            return url != "" && url.Length >= "https://ctext.org/".Length
                && url.Substring(0, "https://ctext.org/".Length) == "https://ctext.org/" && url.IndexOf("edit") > -1
                    && url.LastIndexOf("#editor") > -1 && url.Substring(url.LastIndexOf("#editor")) == "#editor";
            //if (url != "" && url.Length >= "https://ctext.org/".Length
            //    && url.Substring(0, "https://ctext.org/".Length) == "https://ctext.org/" && url.IndexOf("edit") > -1
            //        && url.LastIndexOf("#editor") > -1 && url.Substring(url.LastIndexOf("#editor")) == "#editor") return true;
            //else
            //    return false;
        }
        /// <summary>
        /// 由Url判斷是否是[編輯]頁面（chapter=……&action=editchapter）
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        internal static bool isEditChapterUrl(string url)
        {
            return url != "" && url.Substring(0, "https://ctext.org/".Length) == "https://ctext.org/" &&
                    url.LastIndexOf("&action = editchapter") > -1;
            //if (url != "" && url.Substring(0, "https://ctext.org/".Length) == "https://ctext.org/" &&
            //        url.LastIndexOf("&action = editchapter") > -1) return true;

            //else
            //return false;
        }
        /// <summary>
        /// 由Url判斷是否是瀏覽圖文對照頁面，非[簡單修改模式][Quick edit] 
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        internal static bool isFilePageView(string url)
        {
            //if (url != "" && url.Substring(0, "https://ctext.org/".Length) == "https://ctext.org/" &&
            //url.IndexOf("edit") == -1) return true;
            return (url != "" && url.StartsWith("https://ctext.org/library.pl?") &&
                            url.IndexOf("&file=") > -1 && url.IndexOf("&page=") > -1 &&
                            url.IndexOf("edit") == -1);
            //    return true;
            //else
            //    return false;
        }
        #endregion

        /// <summary>
        /// 儲存chromedriver程序ID的陣列
        /// </summary>
        internal static List<int> chromedriversPID;// = new List<int>();
                                                   ///// <summary>
                                                   ///// 儲存chromedriver程序ID的陣列 chromedriversPID的下標值
                                                   ///// </summary>
                                                   //internal static int chromedriversPIDcntr = 0;

        /// <summary>
        /// 清除從這裡啟動的 chromedriver
        /// </summary>
        internal static void killchromedriverFromHere()
        {
            if (chromedriversPID == null || chromedriversPID.Count == 0) return;
            Process[] processInstances;//= null;
            try
            {
                processInstances = Process.GetProcessesByName("chromedriver");
                if (processInstances.Count() == 0)
                {
                    chromedriversPID.Clear();
                    return;
                }
            }
            catch (Exception)
            {
                return;
            }
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
        /// <summary>
        /// 取得工作管理員中的chromedriver
        /// </summary>
        /// <returns></returns>
        internal static Process[] getChromedrivers()
        {
            return Process.GetProcessesByName("chromedriver");
        }
        /// <summary>
        /// 依工作管理員中的名稱，中止、清除此名稱的所有程序
        /// </summary>
        /// <param name="processName">要找的名稱（不含副檔名）</param>
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
            int page;
            string urlSub;
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

        internal static void importBookmarks()//(ref ChromeDriver drive)
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
            driver = driver ?? DriverNew();
            //string currentWindowHndl = driver.CurrentWindowHandle;
            string currentWindowHndl = LastValidWindow;
            //const string keep = "https://keep.google.com/#NOTE/1XHzZWpH5DCDGOctKjMwNad9qGdtUiYQpSw7HtkmfuEEAJOCtlj37xJg5XgRzWoE";
            string keep = OCRSite_URL[OCRSiteTitle.GoogleKeep];//"https://keep.new";
            openNewTabWindow(WindowType.Window);
            ActiveForm1.TopMost = false;
            driver.Navigate().GoToUrl(keep);
            //取得文字框
            IWebElement iwe = WaitFindWebElementBySelector_ToBeClickable("body > div.VIpgJd-TUo6Hb.XKSfm-L9AdLc.eo9XGd > div > div.IZ65Hb-TBnied.zTETae-h1U9Be-hxXJme > div.IZ65Hb-s2gQvd > div.IZ65Hb-qJTHM-haAclf > div.notranslate.IZ65Hb-YPqjbf.fmcmS-x3Eknd.h1U9Be-YPqjbf");
            iwe.Click();
            //iwe.Clear();
            //按下：新增圖片
            WaitFindWebElementBySelector_ToBeClickable("body > div.VIpgJd-TUo6Hb.XKSfm-L9AdLc.eo9XGd > div > div.IZ65Hb-TBnied.zTETae-h1U9Be-hxXJme > div.IZ65Hb-yePe5c > div.IZ65Hb-INgbqf > div.Q0hgme-LgbsSe.Q0hgme-Bz112c-LgbsSe.Ge5tnd-HiaYvf.INgbqf-LgbsSe.VIpgJd-LgbsSe", 6).Click();
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
            IWebElement iwe_morebtn = WaitFindWebElementBySelector_ToBeClickable("body > div.VIpgJd-TUo6Hb.XKSfm-L9AdLc.eo9XGd > div > div.IZ65Hb-TBnied.zTETae-h1U9Be-hxXJme > div.IZ65Hb-yePe5c > div.IZ65Hb-INgbqf > div.Q0hgme-LgbsSe.Q0hgme-Bz112c-LgbsSe.xl07Ob.INgbqf-LgbsSe.VIpgJd-LgbsSe");
            if (iwe_morebtn == null) return false;
            iwe_morebtn.Click();
            byte tryTimes = 1;//重做記數，防有例外情形、成無窮迴圈；如「按下：新增圖片」處可能還來不及待輸入框出現程式便兀自執行下去
        retry:
            //按下：擷取圖片文字
            //Thread.Sleep(300);

            IWebElement iwe_ocr = WaitFindWebElementBySelector_ToBeClickable("#\\:8 > div", 0); int waitTime = 900;
            while (iwe_ocr == null)
            {   //愈等愈短時間    
                Thread.Sleep(waitTime); iwe_morebtn.Click(); tryTimes++;
                waitTime -= 90;
                if (waitTime < 0 || tryTimes == 10)
                {
                    //MessageBox.Show("請重來一次；重新執行一次。感恩感恩　南無阿彌陀佛", "", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    return false;
                }
                else
                {
                    goto retry;

                }
            }
            Thread.Sleep(3000);
            iwe_ocr.Click();
            //if (iwe_ocr.Location.X != 0)
            //    clickCopybutton_GjcoolFastExperience(new Point(iwe_ocr.Location.X + iwe_ocr.Size.Width / 2, iwe_ocr.Location.Y + iwe_ocr.Size.Height / 2));
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
            WaitFindWebElementBySelector_ToBeClickable("#\\:1 > div").Click();
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

        }
        /// <summary>
        /// 執行《看典古籍》OCR（網頁版）
        /// Alt + Shift + k
        /// </summary>
        /// <param name="downloadImgFullName">書圖下載全檔名</param>
        /// <returns></returns>
        internal static bool OCR_KanDianGuJi(string downloadImgFullName)
        {
            ActiveForm1.TopMost = false;
            //LastValidWindow = driver.CurrentWindowHandle;
            //Form1.ResetLastValidWindow();
            openNewTabWindow();
            //GoToUrlandActivate("https://kandianguji.com/ocr");//https://kandianguji.com/shuzihua?page_mode=img_file
            GoToUrlandActivate("https://kandianguji.com/shuzihua?page_mode=img_file");
            //driver.Navigate().GoToUrl("https://kandianguji.com/ocr");
            Browser.BringToFront("chrome");

            //按下「選擇檔案」按鈕
            //IWebElement iwe = waitFindWebElementBySelector_ToBeClickable("#image-input");
            //IWebElement iwe = waitFindWebElementBySelector_ToBeClickable("#convert-form > label.drop-container", 3);
            IWebElement iwe = WaitFindWebElementBySelector_ToBeClickable("#app > div:nth-child(1) > div > div.col-md-3 > div > div.col-md-10 > label", 3);
            DateTime dt = DateTime.Now;
            while (iwe == null)
            {
                if (DateTime.Now.Subtract(dt).TotalSeconds > 10)
                    if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("等候「選擇檔案」按鈕逾時，是否繼續？") == DialogResult.Cancel)
                    { StopOCR = true; return false; }
                iwe = WaitFindWebElementBySelector_ToBeClickable("#app > div:nth-child(1) > div > div.col-md-3 > div > div.col-md-10 > label");
            }
            //ActiveForm1.TopMost = false;//前已有
            try
            {
                driver.SwitchTo().Window(driver.CurrentWindowHandle);
                iwe.Click();
            }
            catch (Exception)
            {
                StopOCR = true; return false;
            }

            //等待書圖檔下載完成
            dt = DateTime.Now;
            while (!File.Exists(downloadImgFullName))
            {
                if (DateTime.Now.Subtract(dt).TotalSeconds > 38) { StopOCR = true; return false; }
            }

            try
            {
                Clipboard.SetText(downloadImgFullName);
            }
            catch (Exception)
            {
            }
            try
            {
                driver.SwitchTo().Window(driver.CurrentWindowHandle);
            }
            catch (Exception)
            {
                StopOCR = true; return false;
            }
            //等待選取檔案對話框開啟
            //Thread.Sleep(800 + (
            Thread.Sleep(1600 + (
                800 + Extend_the_wait_time_for_the_Open_Old_File_dialog_box_to_appear_Millisecond < 0 ? 0 : Extend_the_wait_time_for_the_Open_Old_File_dialog_box_to_appear_Millisecond));//最小值（須在重開機後或系統最小負載時）（連「開啟」舊檔之視窗也看不見，即可完成）

            while (!File.Exists(downloadImgFullName)) { }

            //輸入：檔案名稱 //SendKeys.Send(downloadImgFullName);
            SendKeys.SendWait("+{Insert}~");//or "^v"
                                            //SendKeys.SendWait("{ENTER}");
                                            //Clipboard.Clear();

            //圖像載入訊息框
            iwe = WaitFindWebElementBySelector_ToBeClickable("#img_create_message");
            while (iwe == null)
            {
                iwe = WaitFindWebElementBySelector_ToBeClickable("#img_create_message");
                if (DateTime.Now.Subtract(dt).TotalSeconds > 38) { StopOCR = true; return false; }
            }
            while (!iwe.GetAttribute("textContent").Contains("Ctext_Page_Image.png")) { if (DateTime.Now.Subtract(dt).TotalSeconds > 38) { StopOCR = true; return false; } }
            //Thread.Sleep(300);

            ////「文本排版方向」點選「豎排」：
            //iwe = waitFindWebElementBySelector_ToBeClickable("#img_det_mode");
            //iwe.Click();

            //點選「 语序优化beta版」核取方塊：（對於正文、夾注之次予至關重要）20240803
            //iwe = waitFindWebElementBySelector_ToBeClickable("#version_2");
            iwe = WaitFindWebElementBySelector_ToBeClickable("#img_rec_version");
            iwe.Click();
            //SetIWebElementValueProperty(iwe, "beta");//這樣設定在按下「開始識別」按鈕：時會被還原
            //點選<option value="beta">语序优化beta版</option>
            iwe = WaitFindWebElementBySelector_ToBeClickable("#img_rec_version > option:nth-child(2)");
            iwe.Click();



            dt = DateTime.Now;
            //按下「開始識別」按鈕：
            iwe = WaitFindWebElementBySelector_ToBeClickable("#img_create_task_button");
            while (iwe == null)
            {
                if (DateTime.Now.Subtract(dt).TotalSeconds > 5)
                {
                    if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("等候「開始識別」按鈕逾時，是否繼續？") == DialogResult.Cancel) { StopOCR = true; return false; }
                }
                iwe = WaitFindWebElementBySelector_ToBeClickable("#img_create_task_button");
            }
            iwe.Click();

            //#result_image
            //currentSrc=https://kandianguji.com/static/loading.gif
            //檢查結果出來沒：            
            dt = DateTime.Now;
            iwe = WaitFindWebElementBySelector_ToBeClickable("#result_text");
            while (iwe == null)
            {
                if (DateTime.Now.Subtract(dt).TotalSeconds > 10)
                    if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("等候OCR結果逾時，是否繼續？") == DialogResult.Cancel) { StopOCR = true; return false; }
                iwe = WaitFindWebElementBySelector_ToBeClickable("#result_text");
            }

            dt = DateTime.Now;
            try
            {
                //while (iwe.GetAttribute("textContent") == "识别结果")
                while (iwe.GetAttribute("value") == "")
                {
                    if (DateTime.Now.Subtract(dt).TotalSeconds > 20)
                        if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("等候OCR結果逾時，是否繼續？") == DialogResult.Cancel) { StopOCR = true; return false; }
                }
            }
            catch (Exception ex)
            {
                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                StopOCR = true; return false;
            }

            //選取OCR結果
            string ocrResult;// = string.Empty;
            try
            {
                ocrResult = iwe.GetAttribute("value");
            }
            catch (Exception ex)
            {
                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                StopOCR = true; return false;
            }

            if (ocrResult.IsNullOrEmpty()) { StopOCR = true; return false; }
            ocrResult = ocrResult.Replace("  ", Environment.NewLine);

            //複製OCR結果

            Clipboard.SetText(ocrResult);

            driver.Close();
            driver.SwitchTo().Window(LastValidWindow);
            return true;
        }

        /// <summary>
        /// 當《古籍酷》點數（算力值、算力配额）小於150時=true
        /// </summary>
        internal static bool waitGJcoolPoint = false;
        /// <summary>
        /// 當《古籍酷》點數（算力值、算力配额）小於150時，須待3小時以上乃可再執行故，以此記下其時間
        /// </summary>
        internal static DateTime gjCoolPointLess150When = new DateTime();
        /// <summary>
        /// 150點算力值（算力配额）須3個多鐘頭才能補足
        /// </summary>
        internal static TimeSpan gjCoolPointEnoughTimespan = new TimeSpan(3, 20, 0);

        /// <summary>
        /// 參照賢超法師《古籍酷AI》OCR視窗是否已關閉用
        /// </summary>
        static bool _OCR_GJcool_WindowClosed = true;
        /// <summary>
        ///若換切換《古籍酷》帳號則為true，以重設算力值的時間區段（點數；算力值、算力配额） 
        /// </summary>
        static bool _OCR_GJcool_AccountChanged = false;
        /// <summary>
        /// 設定 _OCR_GJcool_AccountChanged值；若換切換《古籍酷》帳號則請設為true，以重設算力值的時間區段（點數；算力值、算力配额） 
        /// </summary>
        public static bool OCR_GJcool_AccountChanged { get => _OCR_GJcool_AccountChanged; set => _OCR_GJcool_AccountChanged = value; }

        /// <summary>
        /// 計算已經使用的帳戶
        /// </summary>
        static int gjcoolAccountCounter = 0;
        /// <summary>
        /// 儲存帳的清單：每個元素由帳戶名稱與登入時間組成。登入時間預設為今天（開始設計此機制）的日期20230929
        /// </summary>
        internal static List<Tuple<string, DateTime>> GJcoolAccounts = new List<Tuple<string, DateTime>>()
            { new Tuple<string,DateTime>("oscarsun", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("ssz3ulive", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("ssz3google", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("osyahoo", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("osfoxmail", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("osqqmail", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("997013585", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("9401121", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("oscarsun72ap1pccu", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("g8910512", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("g9206508", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("92065082", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("ossina", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("997013587", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("997013588", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("997013589", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("997013584", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("997013583", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("997013582", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("997013580", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("997013581", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("osprotonme", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("ssz3Proton", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("ssz2Proton", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("ssz1Proton", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("sszProton", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("osproton", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("osprotonssz", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("protonmeos", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("protonos", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("ossszproton", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("sszprotonos", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("28324674", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("protonsun", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("0936158941", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("philipssun", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("sunchouchen", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("sunshouzhen", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("renzhen", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("Buddhistdisciple", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("chiefsun", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("ambilight", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("ambilighttv", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("oscarsun72maggie", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("amituofo", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("Amitabha", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("AmidaBuddha", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("AmitabhaBuddha", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("simplesetupapp", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("gjcool", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("ilovegjcool", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("gjcoolocr", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("TextForCtext", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("touchvpn", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("ivacyvpn", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("thanksxianchaofashi", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("xianchaofashi", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("protonpasstry", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("gjprotonpasstry", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("gjsphere716", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("gjrecast113", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("ivacylithuania", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("ivacyhongkong", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("ivacyaustralia", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("ivacynorway", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("italyfastexperienceivacy", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("ivacyspain", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("ivacyportugal", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("slovakiafastexperienceivacy", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("britishvirginislandsivacy", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("bulgariasivacy", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("gjfrisk195", DateTime.Parse("2023/9/29")),
                new Tuple<string,DateTime>("estoniafastexperienceivacy", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("latviafastexperienceivacy", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("belgiumfastexperienceivacy", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("finlandfastexperienceivacy", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("albaniaivacy", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("czechiaivacy", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("denmarkivacy", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("ivacyrussia", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("freevpnforchrome", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("vpnproxyveepn", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("canadaadblockvpn", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("hongkongadblockvpn", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("irelandadblockvpn", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("mexicoadblockvpnfastexperience", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("austriaadblockvpnfastexperienc", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("franceadblockvpn", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("braziladblockvpnfastexperience", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("czechiaadblockvpnfastexperienc", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("romaniaadblockvpn", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("switzerlandadblockvpnfastexper", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("uslasvegasadblockvpn", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("ussanfranciscoadblockvpn", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("usseattleadblockvpn", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("usmiamiadblockvpn", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("usphoenixadblockvpn", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("finlandadblockvpn", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("netherlandsadblockvpn", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("spainadblockvpn", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("ctextorg", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("greeceivacy", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("egyptivacy", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("bahrainivacy", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("luxembourgivacy", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("afghanistanivacy", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("türkiyeivacy", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("omanivacy", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("barbadosivacy", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("christmaseve", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("christmas", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("jesuschrist", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("germanyzenmate", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("886936158941", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("88628324674", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("111018", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("chilenord", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("czechianord", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("chileproton", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("cyprusproton", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("georgiaproton", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("pakistanproton", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("myanmarproton", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("peruproton", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("estoniaproton", DateTime.Parse("2023/9/29")) ,
                new Tuple<string,DateTime>("unitedstatesivacy", DateTime.Parse("2023/9/29")) ,


                new Tuple<string,DateTime>("vpnbygoogleone", DateTime.Parse("2023/9/29")) };
        /// <summary>
        /// 切換《古籍酷》帳戶時用
        /// </summary>
        /// <param name="justIPSwitch">只切換IP，不換《古籍酷》帳號，則為true</param>
        /// <param name="justSwitchAccount">只切換《古籍酷》帳號，不換IP，則為true</param>
        /// <returns>失敗則傳回false</returns>
        public static bool OCR_GJcool_AccountChanged_Switcher(bool justIPSwitch = false, bool justSwitchAccount = false)
        {

            //waitFindWebElementBySelector_ToBeClickable("#navbarNav > ul:nth-child(2) > li:nth-child(2) > a > p.mb-0.fs-6.fst-italic").Click();
            //Thread.Sleep(150);
            //bool ProtonVPNOn = ProtonVPNSwitcher();

            //ActiveForm1.TopMost = false;//改寫在呼叫端，以免多執行緒時出錯
            try
            {
                if (driver.WindowHandles.Contains(driver.CurrentWindowHandle ?? string.Empty))
                    LastValidWindow = driver.CurrentWindowHandle;
            }
            catch (Exception)
            {
            }
            if (!justSwitchAccount)
            {
                bool switherOk = false;
                //if (!GoogleOneVPNSwitcher()) if (!ProtonVPNSwitcher()) if (!IvacyVPNSwitcher()) TouchVPN_IvacyVPN_VeePN_ExtensionSwitcher();
                switherOk = GoogleOneVPNSwitcher();
                if (!switherOk)
                    switherOk = ProtonVPNSwitcher();
                if (!switherOk)
                    switherOk = IvacyVPNSwitcher();
                if (!switherOk) switherOk = TouchVPN_IvacyVPN_VeePN_ExtensionSwitcher();

                if (switherOk)
                {//檢查切換到的新IP狀態：
                    Thread.Sleep(800);
                    bool showBox = true;
                    //Tuple<bool, bool, bool, bool, DateTime> ipStatus = Mdb.IPStatus(CurrentIP == null ? GetPublicIpAddress("") : CurrentIP);
                    Tuple<bool, bool, bool, bool, DateTime> ipStatus = Mdb.IPStatus(CurrentIP ?? GetPublicIpAddress(""));
                    if (ipStatus != null) showBox = ipStatus.Item4 ? false : true;
                    if (!IPStatusMessageShow(out ipStatus, string.Empty, false, showBox))
                    {

                        //20231228  Bing大菩薩：C# 多執行緒存取剪貼簿：
                        Thread thread = new Thread(() =>
                        {
                            VPNSwitchedCheckOut(CurrentIP);
                        });
                        thread.SetApartmentState(ApartmentState.STA); // 設定執行緒為 STA 模式
                        thread.Start();
                        //thread.Join(); // 等待該執行緒完成
                        if (thread.IsAlive)
                        {
                            // The thread is still running.
                            thread.Join(1000);
                        }
                        else
                        {
                            // The thread has completed.
                        }

                    }

                }
            }



            if (justIPSwitch) { } //{ if (!ActiveForm1.Active) { ActiveForm1.BringToFront(); } }//改寫在呼叫端，以免多執行緒時出錯
            else
            {
                _OCR_GJcool_AccountChanged = true; bool connectedOK = true;
                if (!justSwitchAccount) Thread.Sleep(5950);
                //Thread.Sleep(6950);

                Task ts = null;
                CancellationTokenSource cts = new CancellationTokenSource();
                try
                {
                    //20231021 Bing大菩薩：等待一分鐘
                    ts = Task.Run(() =>
                    {
                        while (!cts.Token.IsCancellationRequested)
                        {
                            int retryCntr = 0;
                        retry:
                            openNewTabWindow(WindowType.Tab);
                            try
                            {
                                //driver.Navigate().GoToUrl("https://gj.cool/account");
                                //driver.Navigate().GoToUrl("https://gj.cool/login");
                                driver.Navigate().GoToUrl("https://gj.cool/login?next=%2Ftry_ocr");
                                //driver.Navigate().GoToUrl("https://ocr.gj.cool/login?next=%2Faccount");
                                //https://ocr.gj.cool/login?next=%2Ftry_ocr
                                break;
                            }
                            catch (Exception ex)
                            {
                                switch (ex.HResult)
                                {
                                    case -2146233088:
                                        if (ex.Message.EndsWith("seconds."))//("timed out after 30.5 seconds."))
                                        {
                                            /* 20231020Bing大菩薩：
                                             * 在C#中，如果你想要確認目前的程式碼是否在主執行緒上執行，你可以使用 `MainThread.IsMainThread` 屬性。如果呼叫此屬性的程式碼是在主執行緒上執行，則 `MainThread.IsMainThread` 屬性會傳回 `true`⁴。

                                                至於你遇到的例外情形，這是因為你嘗試從一個執行緒去存取另一個執行緒所建立的控制項。在Windows Forms中，只有建立控制項的執行緒可以直接操作該控制項。如果你需要從其他執行緒操作控制項，你需要使用 `Control.Invoke` 或 `Control.BeginInvoke` 方法來將操作封送回建立控制項的執行緒。

                                                以下是一個範例：

                                                ```csharp
                                                this.Invoke((MethodInvoker)delegate {
                                                    // 你的程式碼
                                                });
                                                ```

                                                這段程式碼會將你的操作封送回UI執行緒，這樣就可以避免跨執行緒作業的問題了。¹²⁴⁵

                                                來源: 與 Bing 的交談， 2023/10/20
                                                (1) 在主要 UI 執行緒上執行程式碼 - .NET MAUI | Microsoft Learn. https://learn.microsoft.com/zh-tw/dotnet/maui/platform-integration/appmodel/main-thread.
                                                (2) 玩轉C#之【執行序-實際實作】 - iT 邦幫忙::一起幫忙解決難題，拯救 IT 人的一天. https://ithelp.ithome.com.tw/articles/10292658.
                                                (3) 對多執行緒應用程式進行偵錯 - Visual Studio (Windows) | Microsoft Learn. https://learn.microsoft.com/zh-tw/visualstudio/debugger/how-to-use-the-threads-window?view=vs-2022.
                                                (4) 玩轉C#之【執行序-執行緒安全】 - iT 邦幫忙::一起幫忙解決難題，拯救 IT 人的一天. https://ithelp.ithome.com.tw/articles/10292661.
                                                (5) 在主要 UI 執行緒上執行程式碼 - .NET MAUI | Microsoft Learn. https://bing.com/search?q=C%23+%e5%88%a4%e6%96%b7%e6%98%af%e5%90%a6%e5%9c%a8%e5%90%8c%e4%b8%80%e5%9f%b7%e8%a1%8c%e7%b7%92.
                                                (6) undefined. https://bing.com/search?q=.
                                             */
                                            ActiveForm1.Invoke((MethodInvoker)delegate
                                            {
                                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("連線超時！請試著改用其他IP。");
                                            });
                                            connectedOK = false;//return false;
                                        }
                                        else if (ex.Message.StartsWith("unknown error: net::ERR_CONNECTION_RESET"))
                                        {
                                            if (retryCntr++ < 3)
                                            {
                                                Thread.Sleep(2000);
                                                goto retry;
                                            }
                                            else
                                            {
                                                cts.Cancel();
                                                ActiveForm1.Invoke((MethodInvoker)delegate
                                                {
                                                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("連線超時！請試著改用其他IP。");
                                                });
                                                connectedOK = false;//return false;
                                            }

                                        }
                                        else
                                        {
                                            Console.WriteLine(ex.HResult + ex.Message);
                                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                                        }
                                        break;
                                    default:
                                        Console.WriteLine(ex.HResult + ex.Message);
                                        ActiveForm1.Invoke((MethodInvoker)delegate
                                        {
                                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                                        });
                                        break;
                                }
                            }
                        }
                        //});
                    }, cts.Token);
                }
                catch (Exception ex)
                {
                    switch (ex.HResult)
                    {
                        case -2146233088:
                            if (ex.Message.EndsWith("seconds."))//("timed out after 30.5 seconds."))
                            {
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("連線超時！請試著改用其他IP。");
                                return false;
                            }
                            else
                            {
                                Console.WriteLine(ex.HResult + ex.Message);
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                            }
                            break;
                        default:
                            Console.WriteLine(ex.HResult + ex.Message);
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                            break;
                    }
                }
                if (!connectedOK) return false;




                ActiveForm1.HideToNICo();//if (ActiveForm1.TopMost) ActiveForm1.TopMost = false;
                                         //隱藏主表單，以便在切換帳號後，以【按下Shift鍵+滑鼠滑過任務列的表單圖示】，來直接送交《古籍酷》OCR
                                         //if (ProtonVPNOn)
                                         //{

                ActiveForm1.OcrTextMode = true;

                //我想在0~20個數字中隨機取1個 chatGPT大菩薩、Bing大菩薩 20230929

                Random rand = new Random(); int gjcoolaccountsCount = GJcoolAccounts.Count; int i = 0;
                if (gjcoolAccountCounter == gjcoolaccountsCount)
                {
                    for (int j = 0; j < GJcoolAccounts.Count; j++)
                    {
                        GJcoolAccounts[j] = new Tuple<string, DateTime>(GJcoolAccounts[j].Item1, DateTime.Parse("2023/9/29"));
                    }
                    gjcoolAccountCounter = 0;
                    Form1.playSound(Form1.soundLike.warn);
                }
                else
                    gjcoolAccountCounter++;
                while (GJcoolAccounts[i].Item2 != DateTime.Parse("2023/9/29"))
                {
                    //if (DateTime.Now.Subtract(GJcoolAccounts[i].Item2).Days > 0)
                    //if (DateTime.Now.Subtract(GJcoolAccounts[i].Item2).Hours > 22)
                    if (DateTime.Now.Subtract(GJcoolAccounts[i].Item2).TotalHours > 23.6)
                    {
                        gjcoolAccountCounter--;
                        Form1.playSound(Form1.soundLike.exam);
                        break;
                    }
                    i = rand.Next(0, gjcoolaccountsCount);
                }

                string currentAccount = GJcoolAccounts[i].Item1;
                Clipboard.SetText(currentAccount);


                //ts.Wait();
                if (!ts.Wait(TimeSpan.FromMinutes(1)))
                {
                    cts.Cancel();
                    if (!ActiveForm1.Visible && !ActiveForm1.HiddenIcon) ActiveForm1.Visible = true;
                }

                IWebElement ie = WaitFindWebElementBySelector_ToBeClickable("#username");
                if (ie != null)
                {
                    //Task tsk = Task.Run(() =>
                    //{
                    //ie.Clear();ie.SendKeys(OpenQA.Selenium.Keys.Control + "v");            
                    while (ie.Text != "" || ie.Text == null) ie.Clear();//取得的竟是""，故只能以下行執行了
                    ie.SendKeys(OpenQA.Selenium.Keys.Control + "a");
                    //Thread.Sleep(200);
                    //while(ie.Text != "") ie.Clear();
                    ie.SendKeys(currentAccount);
                    //將插入點置於輸入「驗證碼」處：
                    IWebElement ie1 = WaitFindWebElementBySelector_ToBeClickable("#captcha");
                    ie1.Click();//此方法無法將鍵入輸入之插入點實際切換到此方塊內，故須以下行執行
                    ie1.SendKeys(currentAccount);

                    //SendKeys.Send("{tab 14}");//driver.Navigate().GoToUrl("https://gj.cool/login");
                    //SendKeys.Send("{tab 15}");//driver.Navigate().GoToUrl("https://gj.cool/account");
                    //Form1.playSound(Form1.soundLike.done);                    
                    //Point pt = new Point(1002, 656);
                    Point pt = new Point(1002, 666);
                    //Cursor.Position = pt;
                    clickCopybutton_GjcoolFastExperience(pt, Form1.soundLike.done);

                    GJcoolAccounts[i] = new Tuple<string, DateTime>(GJcoolAccounts[i].Item1, DateTime.Now);

                    //});
                    //tsk.Wait();
                    if (ie.Text != currentAccount)
                    {
                        ie.Clear();
                        ie.SendKeys(currentAccount);
                        //SendKeys.Send("{tab 2}");
                        clickCopybutton_GjcoolFastExperience(pt, Form1.soundLike.none);
                    }

                }
            }
            if (!ActiveForm1.Visible && !ActiveForm1.HiddenIcon) ActiveForm1.Visible = true;
            return true;
        }

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        private const int SW_RESTORE = 9;
        private const int SW_MAXIMIZE = 3; // 使用SW_MAXIMIZE來最大化視窗:chatGPT大菩薩：如果你想將視窗最大化，你可以使用 SW_MAXIMIZE 作為 ShowWindow 函數的參數，而不是 SW_RESTORE。這樣可以確保視窗被最大化，而不僅僅是還原到正常大小。
        private const int SW_MINIMIZE = 6;

        /// <summary>
        /// 切換TouchVP、VeePN、IvacyVPN擴充功能。
        /// </summary>
        /// <returns>成功則傳回true</returns>
        internal static bool TouchVPN_IvacyVPN_VeePN_ExtensionSwitcher()
        {
            string targetProcessName = "Ivacy"; // 目標程序的名稱

            // 查找具有指定程式名稱的窗體
            IntPtr targetWindowHandle = FindWindow(null, targetProcessName);

            if (targetWindowHandle != IntPtr.Zero && !IsWindowVisible(targetWindowHandle))
                IvacyVPNExtensionSwithcher();
            else
            {

                //一個擴充功能按鈕長寬大約 35-39 點（35×35）
                Point copyBtnPos = new Point(1705, 55);//擴充功能顯示清單中右邊第2位置 	120.0.6099.110版以後
                                                       //Point copyBtnPos = new Point(1700, 55);//擴充功能顯示清單中最右邊的位置 118.0.5993.89版以後
                                                       //Point copyBtnPos = new Point(1739, 55);//擴充功能顯示清單中最右邊的位置
                Cursor.Position = copyBtnPos;
                //ClickLeftMouse(x, y);
                //Thread.Sleep(150);
                clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.press);
                Thread.Sleep(450);
                //copyBtnPos = new Point(1597, 295);//連接（Connect）按鈕位置
                //copyBtnPos = new Point(1595, 333);//連接（Connect）按鈕位置（此與TouchVPN的有交集）
                //copyBtnPos = new Point(1525, 333);//連接（Connect）按鈕位置（此與TouchVPN的有交集） 118.0.5993.89版以後
                copyBtnPos = new Point(1525, 320);//連接（Connect）按鈕位置（此與TouchVPN、VeePN的有交集） 118.0.5993.89版以後
                Cursor.Position = copyBtnPos;
                clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.over);
                Thread.Sleep(250);
                //copyBtnPos = new Point(1700, 160);//TouchVPN的Stop按鈕
                copyBtnPos = new Point(1630, 160);//TouchVPN的Stop按鈕  118.0.5993.89版以後
                Cursor.Position = copyBtnPos;
                clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.over);
                if (VPNSwitchedTimer == 0)
                    Thread.Sleep(2400);//TouchVPN比較久
                else
                    Thread.Sleep(3500);//TouchVPN比較久
                                       //copyBtnPos = new Point(1595, 333);
                                       //copyBtnPos = new Point(1525, 333);// TouchVPN 連接（Connect）按鈕位置 118.0.5993.89版以後
                copyBtnPos = new Point(1525, 320);// TouchVPN 連接（Connect）按鈕位置 118.0.5993.89版以後
                Cursor.Position = copyBtnPos;
                clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.done);
                Thread.Sleep(150);
                SendKeys.SendWait("{esc}");
                try
                {
                    LastValidWindow = driver.CurrentWindowHandle;

                }
                catch (Exception)
                {
                    //LastValidWindow = driver.WindowHandles[0];
                }
                Thread.Sleep(650);

                Task.Run(() => { Form1.playSound(Form1.soundLike.over); });
                if (VPNSwitchedTimer == 0)
                    Thread.Sleep(850);
                else
                    Thread.Sleep(1850);
            }
            return true;

        }
        /// <summary>
        /// 切換IvacyVPN擴充功能。
        /// </summary>
        /// <returns>成功則傳回true</returns>
        internal static bool IvacyVPNExtensionSwithcher()
        {

            Point copyBtnPos = new Point(1739, 55);//擴充功能顯示清單中最右邊的位置
            Cursor.Position = copyBtnPos;
            clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.press);
            Thread.Sleep(450);

            copyBtnPos = new Point(1606, 401);
            Cursor.Position = copyBtnPos;
            clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.press);
            Thread.Sleep(250);

            SendKeys.SendWait("{tab}");
            Thread.Sleep(250);

            List<string> list = new List<string>() { "United States", "Afghanistan", "Albania", "Australia", "Austria", "Bahrain", "Barbados", "Belgium", "Bolivia", "Brazil", "British Virgin Islands", "Brunei Darussalam", "Bulgaria", "Canada", "Chile", "China", "Czech Republic", "Denmark", "Egypt", "Estonia", "Finland", "France", "Germany", "Greece", "Hong Kong(SAR)", "Hungary", "India", "Ireland", "Italy", "Japan", "Korea, South", "Latvia", "Lithuania", "Luxembourg", "Moldova", "Monaco", "Netherlands", "Nigeria", "Norway", "Oman", "Panama", "Philippines", "Poland", "Portugal", "Romania", "Russia", "Serbia", "Singapore", "Slovakia", "South Africa", "Spain", "Sweden", "Switzerland", "Taiwan", "Turkey", "United Arab Emirates", "United Kingdom", "Vietnam" };
            //20231228 Bing大菩薩：使用C#隨機從清單中選取值：
            /*如果您將 Random 物件的實例化放在方法內部，並且在很短的時間內連續呼叫該方法，那麼可能會產生相同的隨機數字。這是因為 Random 物件的預設種子值是系統時鐘，而系統時鐘的解析度可能只有15毫秒。因此，如果在這個解析度內連續創建 Random 物件，那麼它們的種子值將會是相同的，從而產生相同的隨機數字。
             * 要解決這個問題，您可以將 Random 物件實例化一次，並將其作為類別的欄位或屬性，然後在需要的地方使用該物件。……是的，如果您相隔幾分鐘才調用這個方法一次，那麼 Random 物件的種子值（基於系統時鐘）將有足夠的時間變化，因此每次調用方法時，都能產生不同的隨機數字，從而選取清單中不同的元素。在這種情況下，您不需要將 Random 物件作為類別的欄位或屬性，可以直接在方法內部實例化 Random 物件。
             */
            Random rand = new Random();
            //int index = rand.Next(list.Count);
            //Console.WriteLine(list[index]);
            string country = list[rand.Next(list.Count)];
            Thread.Sleep(350);
            SendKeys.SendWait(country + "~");
            SendKeys.SendWait("{tab}");
            SendKeys.SendWait("{tab}~");

            try
            {
                LastValidWindow = driver.CurrentWindowHandle;

            }
            catch (Exception)
            {
                //LastValidWindow = driver.WindowHandles[0];
            }
            Thread.Sleep(650);
            //SystemSounds.Exclamation.Play();
            //SendKeys.SendWait("{ESC}");
            SendKeys.SendWait("{ESC}");
            Thread.Sleep(850);

            Task.Run(() => { Form1.playSound(Form1.soundLike.over); });

            return true;

        }
        /// <summary>
        /// 在VPN切換IP後的檢查
        /// </summary>
        /// <param name="ip"></param>
        /// <param name="already">若已經檢查，欲略過，則為true</param>
        internal static void VPNSwitchedCheckOut(string ip = "", bool already = false)
        {
            if (already) return;
            openNewTabWindow(WindowType.Tab);
            Thread.Sleep(850);
            if (ip == string.Empty)
            {
                GoToUrlandActivate("https://iplocation.com/");
                ip = WaitFindWebElementBySelector_ToBeClickable("body > div.top-container > div.bottom-container > div.rubber-container.result > div > table > tbody > tr:nth-child(1) > td > b")?.Text;
                Thread.Sleep(850);
            }
            Clipboard.SetText(ip);
            //openNewTabWindow(WindowType.Tab);
            GoToUrlandActivate("https://gj.cool/login?next=%2Ftry_ocr");//檢視IP是否被封鎖
                                                                        //Thread.Sleep(850);
            Thread.Sleep(1450);
            IWebElement iwe = WaitFindWebElementBySelector_ToBeClickable("body > div.alert.alert-danger.alert-dismissible");
            string info = string.Empty;
            if (iwe != null)
            {
                info = iwe.Text;
                if (info.Contains("IP已被封锁。 IP is blocked.") || info.Contains("系统太忙。 System is busy.") || info.StartsWith("system is busy"))
                {//開啟Google Keep記錄，以檢視封鎖情況：
                 //openNewTabWindow(WindowType.Tab);
                 //GoToUrlandActivate("https://keep.google.com/#NOTE/1-bHzJG4vtIyJMsT7SSRkgSZ2DSvcAabhLC88WERCnPxTd9MqsSXwgpHxYFU2");
                 //if (!ip.IsNullOrEmpty())
                 //{
                 //    SystemSounds.Exclamation.Play();
                 //    Thread.Sleep(2200);
                 //    SendKeys.SendWait("^f");
                 //    SendKeys.SendWait(ip + "~");
                 //}
                 //20231228  Bing大菩薩：C# 多執行緒存取剪貼簿：


                    Task.Run(() =>
                    {
                        if (File.Exists("C:\\Windows\\Media\\ring05.wav"))
                        {
                            using (SoundPlayer sp = new SoundPlayer("C:\\Windows\\Media\\ring05.wav"))
                            {
                                sp.Play();
                            }
                        }
                    });

                    var ipstatus = Mdb.IPStatus(ip);
                    if (ipstatus == null)
                    {
                        ADODB.Connection cnt = new ADODB.Connection();
                        Mdb.openDatabase("查字.mdb", ref cnt);
                        ADODB.Recordset rst = new ADODB.Recordset();
                        rst.Open("IP", cnt, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic);
                        rst.AddNew();
                        rst.Fields["IP"].Value = ip;
                        if (info.Contains("IP已被封锁。 IP is blocked."))
                        {
                            rst.Fields["IpAddressBanned"].Value = true;
                            rst.Fields["IPisblocked"].Value = false;
                            rst.Fields["ctext"].Value = false;
                            rst.Fields["Systemisbusy"].Value = false;
                        }
                        else if (info.Contains("系统太忙。 System is busy.") || info.StartsWith("system is busy"))
                        {
                            rst.Fields["IpAddressBanned"].Value = false;
                            rst.Fields["IPisblocked"].Value = false;
                            rst.Fields["ctext"].Value = false;
                            rst.Fields["Systemisbusy"].Value = true;
                        }

                        rst.Update();
                        rst.Close(); cnt.Close();
                    }

                    if (ActiveForm1.InvokeRequired)
                    {
                        ActiveForm1.Invoke((MethodInvoker)delegate
                        {
                            ActiveForm1.AvailableInUseBothKeysMouse();
                        });
                    }
                }

            }
            try
            {
                driver.Close();
                if (info.Contains("系统太忙。 System is busy.") || info.StartsWith("system is busy"))
                {
                    if (ActiveForm1.InvokeRequired)
                    {
                        if (VPNSwitchedTimer >= 0 && VPNSwitchedTimer % 20 == 0)
                        {
                            if (DialogResult.OK == MessageBox.Show("已經切換" + VPNSwitchedTimer + "次了，還要繼續嗎？", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly))
                            {
                                ActiveForm1.Invoke((MethodInvoker)delegate { ActiveForm1.Controls["textBox2"].Text = "kk"; });
                                VPNSwitchedTimer++;
                            }
                            else
                                VPNSwitchedTimer = 0;
                        }
                        else
                        {
                            ActiveForm1.Invoke((MethodInvoker)delegate { ActiveForm1.Controls["textBox2"].Text = "kk"; });
                            VPNSwitchedTimer++;
                        }

                    }
                }

            }
            catch (Exception)
            {
            }
        }

        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        /// <summary>
        /// 切換ProtonVPN。請將其視窗最大化
        /// </summary>
        /// <returns>成功則傳回true</returns>
        internal static bool ProtonVPNSwitcher()
        {
            string targetProcessName = "Proton VPN";//"ProtonVPN.exe"; // 目標程序的名稱

            // 查找具有指定程式名稱的窗體
            IntPtr targetWindowHandle = FindWindow(null, targetProcessName);

            if (targetWindowHandle != IntPtr.Zero)
            {

                if (IsWindowVisible(targetWindowHandle))
                {

                    // 將目標窗口切換到最前面
                    ShowWindow(targetWindowHandle, SW_MAXIMIZE);//SW_RESTORE);
                    SetForegroundWindow(targetWindowHandle);
                    /* chatGPT大菩薩：20230926
                     * 你正確，SetForegroundWindow 方法在視窗最小化時可能無法成功將其切換到最前面。為了解決這個問題，你可以嘗試使用 ShowWindow 函數來將視窗恢復到正常狀態，然後再調用 SetForegroundWindow。這樣可以確保視窗在最前面並且可見。
                     */
                    Thread.Sleep(150);
                    // 模擬滑鼠左鍵點擊指定座標（Random Connect按鈕）
                    //免費版：
                    //int x = 229;////338;
                    //int y = 210;//161//364;
                    //            //Change server 按鈕
                    //            //Disconnect 按鈕
                    //付費版：
                    int x = 338;
                    int y = 364;
                    Point copyBtnPos = new Point(x, y);
                    Cursor.Position = copyBtnPos;
                    //ClickLeftMouse(x, y);
                    Thread.Sleep(350);
                    clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.none);
                    Thread.Sleep(150);
                    //Task.Run(() => { Thread.Sleep(5800); ShowWindow(targetWindowHandle, SW_MINIMIZE); });
                    Task.Run(() => { Thread.Sleep(8800); ShowWindow(targetWindowHandle, SW_MINIMIZE); });
                }
                else
                    ChromeExtensionsProtonVPNSwitcher();

                return true;
            }
            return false;
        }


        internal static bool ChromeExtensionsProtonVPNSwitcher()
        {

            //一個擴充功能按鈕長寬大約 35-39 點（35×35）
            //Point copyBtnPos = new Point(1662, 55);//擴充功能顯示清單中右邊第3位置 120.0.6099.110版以後
            //Point copyBtnPos = new Point(1705, 55);//擴充功能顯示清單中右邊第2位置 	120.0.6099.110版以後
            //Point copyBtnPos = new Point(1700, 55);//擴充功能顯示清單中最右邊的位置 118.0.5993.89版以後
            Point copyBtnPos = new Point(1739, 55);//擴充功能顯示清單中最右邊的位置
            Cursor.Position = copyBtnPos;
            //ClickLeftMouse(x, y);
            //Thread.Sleep(150);
            clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.press);
            Thread.Sleep(450);

            List<string> list = new List<string>() { "土耳其", "丹麥", "厄瓜多", "巴西", "巴基斯坦", "日本", "比利時", "以色列", "加拿大", "北馬其頓", "台灣", "立陶宛", "冰島", "匈牙利", "印尼", "印度", "西班牙", "克羅埃西亞", "希臘", "奈及利亞", "拉脫維亞", "法國", "波多黎各", "波蘭", "芬蘭", "阿拉伯聯合大公國", "阿根廷", "俄羅斯", "保加利亞", "南非", "南韓", "柬埔寨", "美國", "英國", "香港", "哥倫比亞", "哥斯達黎加", "埃及", "挪威", "泰國", "烏克蘭", "秘魯", "紐西蘭", "馬來西亞", "馬爾他", "捷克", "荷蘭", "喬治亞", "斯洛伐克", "斯洛維尼亞", "智利", "菲律賓", "越南", "塞爾維亞", "奧地利", "愛沙尼亞", "愛爾蘭", "新加坡", "瑞士", "瑞典", "義大利", "葡萄牙", "德國", "摩洛哥", "摩爾多瓦", "緬甸", "墨西哥", "澳洲", "盧森堡", "賽普勒斯", "羅馬尼亞" };
            //20231228 Bing大菩薩：使用C#隨機從清單中選取值：
            /*如果您將 Random 物件的實例化放在方法內部，並且在很短的時間內連續呼叫該方法，那麼可能會產生相同的隨機數字。這是因為 Random 物件的預設種子值是系統時鐘，而系統時鐘的解析度可能只有15毫秒。因此，如果在這個解析度內連續創建 Random 物件，那麼它們的種子值將會是相同的，從而產生相同的隨機數字。
             * 要解決這個問題，您可以將 Random 物件實例化一次，並將其作為類別的欄位或屬性，然後在需要的地方使用該物件。……是的，如果您相隔幾分鐘才調用這個方法一次，那麼 Random 物件的種子值（基於系統時鐘）將有足夠的時間變化，因此每次調用方法時，都能產生不同的隨機數字，從而選取清單中不同的元素。在這種情況下，您不需要將 Random 物件作為類別的欄位或屬性，可以直接在方法內部實例化 Random 物件。
             */
            Random rand = new Random();
            //int index = rand.Next(list.Count);
            //Console.WriteLine(list[index]);
            string country = list[rand.Next(list.Count)];
            Thread.Sleep(1650);
            SendKeys.SendWait(country + "~");
            try
            {
                LastValidWindow = driver.CurrentWindowHandle;

            }
            catch (Exception)
            {
                //LastValidWindow = driver.WindowHandles[0];
            }
            Thread.Sleep(650);
            //SystemSounds.Exclamation.Play();
            SendKeys.SendWait("{ESC}");
            SendKeys.SendWait("{ESC}");
            Thread.Sleep(850);

            Task.Run(() => { Form1.playSound(Form1.soundLike.over); });

            return true;
        }

        /// <summary>
        /// 顯示IP現狀。快速鍵 Ctrl + Alt + i
        /// </summary>        
        /// <param name="currentip"></param>
        /// <param name="already">若已經檢查，則為true</param>
        /// <returns>若IP OK，未被封鎖，則傳回 false 且不顯示訊息方塊 ；若封鎖，則傳回 true 且顯示訊息方塊</returns>
        internal static bool IPStatusMessageShow(out Tuple<bool, bool, bool, bool, DateTime> ipStatus, string currentip = "", bool already = false, bool showMsgBox = true)
        {
            if (currentip == string.Empty) currentip = GetPublicIpAddress(string.Empty);
            else
            {
                if (currentip == GetPublicIpAddress(string.Empty))
                {
                    MessageBox.Show("IP未更動！", "IP Status", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    ipStatus = null;
                    return true;
                }
            }
            CurrentIP = currentip;
            ipStatus = Mdb.IPStatus(CurrentIP);
            if (already) return false;
            string msgText = "";
            if (ipStatus == null)
            {
                msgText = "目前IP " + CurrentIP + " OK!";
                if (File.Exists("C:\\Windows\\Media\\windows balloon.wav"))
                {
                    using (SoundPlayer sp = new SoundPlayer("C:\\Windows\\Media\\windows balloon.wav"))
                    {
                        sp.Play();
                    }
                }
            }
            else
            {
                if (showMsgBox)
                {
                    if (ipStatus.Item4)
                    {
                        msgText = "系统太忙。 System is busy.";
                        MessageBox.Show(msgText, CurrentIP + "IP Status", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    }
                    else if (ipStatus.Item1 && !ipStatus.Item2)
                    {
                        msgText = "只有快速体验 Fast Experience 被封鎖";
                        MessageBox.Show(msgText, CurrentIP + "IP Status", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    }
                    else if (!ipStatus.Item1 && ipStatus.Item2)
                    {
                        msgText = "只有 try_ocr 被封鎖";
                        MessageBox.Show(msgText, CurrentIP + "IP Status", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    }
                    else if (ipStatus.Item1 && ipStatus.Item2)
                    {
                        msgText = "《古籍酷》OCR皆被封鎖";
                        MessageBox.Show(msgText, CurrentIP + "IP Status", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    }
                    else if (ipStatus.Item3)
                    {
                        msgText = "只有《中國哲學書電子化計劃》被封鎖";
                        MessageBox.Show(msgText, CurrentIP + "IP Status", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    }
                }
                else
                {
                    if (ipStatus.Item4)
                    {
                        SystemSounds.Hand.Play();
                        if (ActiveForm1.InvokeRequired)
                        {
                            if (VPNSwitchedTimer > 0 && VPNSwitchedTimer % 20 == 0)
                            {
                                if (DialogResult.OK == MessageBox.Show("已經切換" + VPNSwitchedTimer + "次了，還要繼續嗎？", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly))
                                {
                                    ActiveForm1.Invoke((MethodInvoker)delegate { ActiveForm1.Controls["textBox2"].Text = "kk"; });
                                    VPNSwitchedTimer++;
                                }
                                else
                                    VPNSwitchedTimer = 0;
                            }
                            else
                            {
                                ActiveForm1.Invoke((MethodInvoker)delegate { ActiveForm1.Controls["textBox2"].Text = "kk"; });
                                VPNSwitchedTimer++;
                            }
                        }
                    }
                    else
                    {
                        SystemSounds.Asterisk.Play();
                    }
                }
                if (ActiveForm1.InvokeRequired)
                {
                    ActiveForm1.Invoke((MethodInvoker)delegate
                    {
                        ActiveForm1.AvailableInUseBothKeysMouse();
                    });
                }
            }
            return msgText != "目前IP " + CurrentIP + " OK!";
        }
        /// <summary>
        /// 切換IvacyVPN
        /// </summary>
        /// <returns>成功則傳回true</returns>
        internal static bool IvacyVPNSwitcher()
        {
            string targetProcessName = "Ivacy"; // 目標程序的名稱

            // 查找具有指定程式名稱的窗體
            IntPtr targetWindowHandle = FindWindow(null, targetProcessName);

            if (targetWindowHandle != IntPtr.Zero && IsWindowVisible(targetWindowHandle))
            {
                // 將目標窗口切換到最前面
                //ShowWindow(targetWindowHandle, SW_MAXIMIZE);//SW_RESTORE);
                ShowWindow(targetWindowHandle, SW_RESTORE);
                SetForegroundWindow(targetWindowHandle);
                /* chatGPT大菩薩：20230926
                 * 你正確，SetForegroundWindow 方法在視窗最小化時可能無法成功將其切換到最前面。為了解決這個問題，你可以嘗試使用 ShowWindow 函數來將視窗恢復到正常狀態，然後再調用 SetForegroundWindow。這樣可以確保視窗在最前面並且可見。
                 */
                Thread.Sleep(250);
                // 模擬滑鼠左鍵點擊指定座標（Random Connect按鈕）
                int x = 1018;
                int y = 652;
                Point copyBtnPos = new Point(x, y);
                Cursor.Position = copyBtnPos;
                //ClickLeftMouse(x, y);
                //Thread.Sleep(150);
                Thread.Sleep(450);
                clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.none);
                //Thread.Sleep(1150);//等待斷開 
                Thread.Sleep(1350);//等待斷開
                clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.none);
                Thread.Sleep(6500);//監看連線成功
                return true;
            }
            return false;
        }
        /// <summary>
        /// 切換Google One VPN
        /// 難怪用這個這麼超快急速的：因為我的回定IP就是設定這樣的：http://www.pcdiy.com.tw/detail/1412 〈Google Public DNS上網跑更快，用戶端趕快更換IPv4 DNS設定8.8.8.8與8.8.4.4〉感恩感恩　讚歎讚歎　Google大神大菩薩　南無阿彌陀佛
        /// https://briian.com/6667/ Google 更快更安全的 DNS 伺服器：8.8.8.8 與 8.8.4.4 （含 iPv6）
        /// https://www.google.com/search?q=DNS+8.8.4.48.8.8.8&oq=DNS+8.8.4.48.8.8.8&gs_lcrp=EgZjaHJvbWUyBggAEEUYOTIHCAEQABiiBDIHCAIQABiiBDIHCAMQABiiBDIHCAQQABiiBNIBCDEyMzZqMGo3qAIAsAIA&sourceid=chrome&ie=UTF-8
        /// IP 設定
        /// IP 指派: 手動
        /// IPv4 DNS 伺服器:	8.8.4.4
        /// 屬性
        /// 連結速度(接收/傳輸):100 / 100 (M bps)
        /// IPv4 位址:
        /// IPv4 DNS 伺服器:	8.8.4.4
        ///                 8.8.8.8
        /// 主要 DNS 尾碼:	dynamic.kbronet.com.tw
        /// 以上是用 Text-Grab 擷取畫面OCR再校訂的；以下是在Windows 設定中直接按下「複製」鈕的
        /// 連結速度(接收/傳輸) :   100/100 (Mbps)
        /// IPv4 位址:	180.176.40.208
        /// IPv4 DNS 伺服器:	8.8.4.4
        ///                 8.8.8.8
        /// 主要 DNS 尾碼:	dynamic.kbronet.com.tw
        /// 製造商:	Realtek        
        /// </summary>
        /// <returns>成功則傳回true</returns>
        internal static bool GoogleOneVPNSwitcher()
        {
            string targetProcessName = "VPN by Google One"; // 目標程序的名稱

            // 查找具有指定程式名稱的窗體
            IntPtr targetWindowHandle = FindWindow(null, targetProcessName);

            if (targetWindowHandle != IntPtr.Zero)
            {

                if (ActiveForm1.InvokeRequired)
                {
                    ActiveForm1.Invoke((MethodInvoker)delegate
                    {
                        ActiveForm1.TopMost = false;// 你的程式碼
                                                    //ActiveForm1.WindowState = FormWindowState.Minimized;
                    });
                }

                // 將目標窗口切換到最前面
                //ShowWindow(targetWindowHandle, SW_MAXIMIZE);//SW_RESTORE);
                ShowWindow(targetWindowHandle, SW_RESTORE);
                SetForegroundWindow(targetWindowHandle);
                /* chatGPT大菩薩：20230926
                 * 你正確，SetForegroundWindow 方法在視窗最小化時可能無法成功將其切換到最前面。為了解決這個問題，你可以嘗試使用 ShowWindow 函數來將視窗恢復到正常狀態，然後再調用 SetForegroundWindow。這樣可以確保視窗在最前面並且可見。
                 */
                Thread.Sleep(150);
                // 模擬滑鼠左鍵點擊指定座標（Random Connect按鈕）
                int x = 1066;
                int y = 528;
                Point copyBtnPos = new Point(x, y);
                Cursor.Position = copyBtnPos;
                //ClickLeftMouse(x, y);
                Thread.Sleep(150);
                clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.none);
                Thread.Sleep(850);//等待斷開
                clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.none);
                Thread.Sleep(900);//監看連線成功

                Point form1Location = ActiveForm1.Location; bool eventEnable = ActiveForm1.EventsEnabled;
                try
                {
                    if (ActiveForm1.InvokeRequired)
                    {
                        ActiveForm1.Invoke((MethodInvoker)delegate
                        {
                            ActiveForm1.PauseEvents();
                            ActiveForm1.Location = new Point(form1Location.X + 800, form1Location.Y);
                        });
                    }
                }
                catch (Exception)
                {
                    //throw;
                }
                Thread.Sleep(1500);//等候連線成功
                DateTime dt = DateTime.Now;
                //while (GetPCIpAddress("SUNS TOTOLINK") == GetPublicIpAddress())                
                while (GetPCIpAddress("乙太網路") == GetPublicIpAddress("乙太網路")
                || GetPCIpAddress("Wi-Fi NetGear") == GetPublicIpAddress("Wi-Fi NetGear"))
                {
                    Thread.Sleep(2200);
                    if (DateTime.Now.Subtract(dt).Seconds > _chromeDriverServiceTimeSpan)
                    {
                        //if (ActiveForm1.InvokeRequired)
                        //{
                        //    ActiveForm1.Invoke((MethodInvoker)delegate
                        //    {
                        //        ActiveForm1.TopMost = false;// 你的程式碼
                        //        //ActiveForm1.WindowState = FormWindowState.Normal;
                        //        ActiveForm1.Activate();
                        //    });
                        //}
                        return false;
                    }
                }
                if (IPExists1Day())
                {
                    if (ipChangedCounter < 10)//23)
                    {
                        ipChangedCounter++;
                        //Debugger.Break();
                        Form1.playSound(Form1.soundLike.over);
                        GoogleOneVPNSwitcher();
                    }
                    else
                    {
                        //if (DialogResult.OK == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("已嘗試10次了,是否繼續？"))
                        if (DialogResult.OK == MessageBox.Show("已嘗試10次了,是否繼續？\n\r\n\r已經用掉" + IPUsedList.Count + "個IP了", "google one vpn", MessageBoxButtons.OKCancel,
                            MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly))
                        {
                            //Console.WriteLine(IPUsedList.this[].ToString());
                            string ipList = string.Empty;
                            foreach (var item in IPUsedList)
                            {
                                ipList += item.Item1 + "\t" + item.Item2 + Environment.NewLine;
                                Console.WriteLine("IP: " + item.Item1 + ", Time: " + item.Item2);
                            }
                            //Debugger.Break();
                            MessageBox.Show("已經用掉" + IPUsedList.Count + "個IP:\n\r\n\r" + ipList, "google one vpn", MessageBoxButtons.OK,
                                MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                            ipChangedCounter = 0; GoogleOneVPNSwitcher();
                        }
                    }
                }
                else
                    ipChangedCounter = 0;//計數器歸零
                try
                {
                    if (ActiveForm1.InvokeRequired)
                    {
                        ActiveForm1.Invoke((MethodInvoker)delegate
                        {
                            ActiveForm1.EventsEnabled = eventEnable;
                            //ActiveForm1.ResumeEvents();
                            ActiveForm1.Location = form1Location;
                            ActiveForm1.AvailableInUseBothKeysMouse();
                        });
                    }
                }
                catch (Exception)
                {
                    throw;
                }
                Form1.playSound(Form1.soundLike.done);
                //if (ActiveForm1.InvokeRequired)
                //{
                //    ActiveForm1.Invoke((MethodInvoker)delegate
                //    {
                //        ActiveForm1.TopMost = false;// 你的程式碼
                //        //ActiveForm1.WindowState = FormWindowState.Normal;
                //        ActiveForm1.Activate();
                //    });
                //}
                return true;
            }
            Form1.playSound(Form1.soundLike.done);
            //if (ActiveForm1.InvokeRequired)
            //{
            //    ActiveForm1.Invoke((MethodInvoker)delegate
            //    {
            //        ActiveForm1.TopMost = false;// 你的程式碼
            //        //ActiveForm1.WindowState = FormWindowState.Normal;
            //        ActiveForm1.Activate();
            //    });
            //}
            return false;
        }


        /// <summary>
        /// 更換VPN IP
        /// </summary>
        /// <returns>成功則傳回true</returns>
        internal static bool IPSwitchOnly()
        {
            if (ActiveForm1.TopMost) ActiveForm1.TopMost = false; if (!waitGJcoolPoint) waitGJcoolPoint = true;
            Task tk = Task.Run(() => { OCR_GJcool_AccountChanged_Switcher(true, false); });
            if (tk.Wait(4500))
            {

                //ActiveForm1.Invoke((MethodInvoker)delegate
                //{
                //    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("連線超時！請試著改用其他IP。");
                //});
                ActiveForm1.Invoke((MethodInvoker)delegate
                {
                    //ActiveForm1.BringToFront();
                    ActiveForm1.AvailableInUseBothKeysMouse();
                    //ActiveForm1.Activate();
                });


                //Process.Start("https://gj.cool/login?next=%2Ftry_ocr https://iplocation.com/");
                //Microsoft.VisualBasic.Interaction.Shell("W:\\PortableApps\\PortableApps\\GoogleChromePortable\\App\\Chrome-bin\\chrome.exe https://gj.cool/login?next=%2Ftry_ocr https://iplocation.com/"
                //    , Microsoft.VisualBasic.AppWinStyle., bool, int)
                // 前面 OCR_GJcool_AccountChanged_Switcher 中已經檢查過了
                //if (!IPStatusMessageShow())
                //    VPNSwitchedCheckOut(CurrentIP);
            }
            return true;
        }

        /// <summary>
        /// 轉換IP、重試VPN連線的次數
        /// </summary>
        static int ipChangedCounter = 0;
        /// <summary>
        /// 用來存放已經使用的IP及開始使用的時間，以供後續參照比較
        /// </summary>
        internal static List<Tuple<string, DateTime>> IPUsedList = new List<Tuple<string, DateTime>>();
        /// <summary>
        /// 檢查目前IP是否在一天內已經用過了
        /// </summary>
        /// <returns>若在一天內已用過，則傳回true</returns>
        internal static bool IPExists1Day()
        {


            ////20231102 Bing大菩薩：取得本機電腦的IPv4地址
            //string hostName = Dns.GetHostName(); // 取得主機名稱
            //Console.WriteLine("主機名稱: " + hostName);

            //// 取得本機的IPv4地址
            //var ipAddresses = Dns.GetHostAddresses(hostName)
            //                      .Where(ip => ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
            //                      .Select(ip => ip.ToString());

            //if (ipAddresses.Count() > 2)
            //{

            //    Debugger.Break();//just for check out

            //    foreach (var ipAddress in ipAddresses)
            //    {
            //        Console.WriteLine("IPv4 地址: " + ipAddress);
            //    }
            //}

            string currentIP = GetPublicIpAddress("乙太網路"); //= ipAddresses.First();//.ElementAt(1);
            if (string.IsNullOrEmpty(currentIP)) currentIP = GetPublicIpAddress("Wi-Fi NetGear");
            //currentIP = GetVpnIpAddress("VPN by Google One");//("VPN by Google One 25");

            //20231102 Bing大菩薩：查找 List 中的元素
            //bool returnValue = IPUsedList.Exists(item => item.Item1 == currentIP && (DateTime.Now - item.Item2).TotalDays >= 1);
            bool returnValue = IPUsedList.Exists(item => item.Item1 == currentIP && (DateTime.Now - item.Item2).TotalHours >= 22);
            if (!returnValue)
            {
                //if(IPUsedList.Exists(item => item.Item1 == currentIP))
                //{
                //    IPUsedList.Add new to
                //}
                #region internal static void UpdateIP(string currentIP)
                //{
                int index = IPUsedList.FindIndex(item => item.Item1 == currentIP);
                if (index != -1)
                {
                    // 元素存在，更新日期時間
                    IPUsedList[index] = new Tuple<string, DateTime>(currentIP, DateTime.Now);
                }
                else
                {
                    // 元素不存在，添加新元素
                    IPUsedList.Add(new Tuple<string, DateTime>(currentIP, DateTime.Now));
                }
                #endregion    //}

            }
            return returnValue;
            //foreach (Tuple<string,DateTime> item in IPUsedList)
            //{
            //    if (item.Item1 == currentIP)
            //    {
            //        if
            //        break;
            //    }
            //}

        }
        /// <summary>
        /// 這個 GetPublicIpAddress 方法會返回您的公共IP地址
        /// 20231102 Bing大菩薩 ： 查找 List 中的元素
        /// </summary>
        /// <returns></returns>
        public static string GetPublicIpAddress(string name)
        {
            string publicIpAddress = string.Empty;

            using (var webClient = new System.Net.WebClient())
            {
                try
                {
                    publicIpAddress = webClient.DownloadString("https://api.ipify.org");
                }
                catch (Exception)
                {
                    Thread.Sleep(2800);
                    try
                    {
                        publicIpAddress = webClient.DownloadString("https://api.ipify.org");
                    }
                    catch (Exception ex)
                    {
                        if (!(ex.HResult == -2146233079 && ex.Message == "要求已經中止: 無法建立 SSL/TLS 的安全通道。")) //不顯示 20250119                        
                            MessageBox.Show(ex.HResult + ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    }
                }
                //string pcIpAddress = GetPCIpAddress("乙太網路");
                string pcIpAddress = GetPCIpAddress(name);
                if (publicIpAddress == pcIpAddress || pcIpAddress == null)
                {
                    //if (driver.Url == "https://api.ipify.org/") driver.Close();
                    ActiveForm1.PauseEvents();
                    string ipUrl = "https://api.ipify.org", selector = "body > pre";
                //string ipUrl = "https://www.whatismyip.com.tw/", selector = "body > b > span";
                retry:
                    //Thread.Sleep(5000);
                    openNewTabWindow();//要打開比較快更新
                    try
                    {
                        driver.Navigate().GoToUrl(ipUrl);
                    }
                    catch (Exception)
                    {
                        openNewTabWindow();//要打開比較快更新
                        try
                        {
                            driver.Navigate().GoToUrl(ipUrl);
                        }
                        catch (Exception)
                        {
                        }
                    }
                    DateTime dt = DateTime.Now;
                    IWebElement ie = WaitFindWebElementBySelector_ToBeClickable(selector);
                    while (ie == null)
                    {
                        ie = WaitFindWebElementBySelector_ToBeClickable(selector);
                        if (DateTime.Now.Subtract(dt).Seconds > 15)
                        {
                            if (ipUrl == "https://api.ipify.org")
                            {
                                selector = "body > b > span"; dt = DateTime.Now;
                                ipUrl = "https://www.whatismyip.com.tw/"; goto retry;
                            }
                            else if (ipUrl == "https://www.whatismyip.com.tw/")
                            {
                                selector = "body > pre"; dt = DateTime.Now;
                                ipUrl = "https://api.ipify.org"; ; goto retry;
                            }
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("找不到外部網路IP");
                            Debugger.Break();
                            return string.Empty;
                        }
                    }
                    publicIpAddress = ie.Text;
                    //ie = waitFindWebElementBySelector_ToBeClickable("body > i > span");
                    //if (ie != null)//顯示IP國家，ipUrl = "https://www.whatismyip.com.tw/" 才有 //Google One 目前也僅提供臺灣的（最多2個日本的），故略去 20231129
                    //    MessageBox.Show(ie.Text, "country:", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    try
                    {
                        driver.Close();
                        driver.SwitchTo().Window(LastValidWindow);
                    }
                    catch (Exception)
                    {
                    }
                    ActiveForm1.ResumeEvents();
                    if (publicIpAddress != pcIpAddress) return publicIpAddress;
                }
            }
            return publicIpAddress;
        }
        /// <summary>
        /// 在C#中，您可以使用 System.Net.NetworkInformation 命名空間下的 NetworkInterface 類來獲取網絡接口的信息，包括VPN的IP地址。以下是一個範例：
        /// 20231102Bing大菩薩：查找 List 中的元素
        ///public static string GetVpnIpAddress(string vpnName)
        /// </summary>
        /// <param name="vpnName"></param>
        /// <returns></returns>
        public static string GetPCIpAddress(string vpnName)
        {
            var vpn = System.Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces().FirstOrDefault(x => x.Name == vpnName);
            if (vpn != null)
            {
                var ipProperties = vpn.GetIPProperties();
                var ipv4Address = ipProperties.UnicastAddresses.FirstOrDefault(x => x.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork);
                if (ipv4Address != null)
                {
                    return ipv4Address.Address.ToString();
                }
            }

            return null;
        }

        internal static bool OCR_GJcool_BatchProcessing(string downloadImgFullName, bool _downloadResult = false)
        {

            try
            {
                driver = driver ?? DriverNew();
                string currentWindowHndl = driver.CurrentWindowHandle;
                openNewTabWindow(WindowType.Tab);
                GoToUrlandActivate(OCRSite_URL[OCRSiteTitle.GJcool]);
                //GoToUrlandActivate("https://gj.cool/try_ocr#");
                //GoToUrlandActivate("https://gj.cool/try_ocr#Batch");

                //按下「批量處理（面向授權用戶）」頁籤
                IWebElement iwe = WaitFindWebElementBySelector_ToBeClickable("#batch_Tab_A", 15);
                iwe.Click();

                //按下「OCR」
                iwe = WaitFindWebElementBySelector_ToBeClickable("#batchUploadSelect0");
                iwe.Click();

                driver.SwitchTo().Window(driver.CurrentWindowHandle);
                //按下「批量處理（面向授權用戶）」頁面下的「選擇檔案」
                //iwe = waitFindWebElementBySelector_ToBeClickable("#BatchFilesInput");
                iwe = WaitFindWebElementBySelector_ToBeClickable("#Batch > div.d-flex.justify-content-between.mt-2 > div.d-flex.ms-2 > div:nth-child(2) > button > i");
                iwe.Click();
                //clickCopybutton_GjcoolFastExperience(new Point(iwe.Location.X + 76 + (iwe.Size.Width) / 2, iwe.Location.Y + 120 + (iwe.Size.Height) / 2));//new Point(X, Y)=「選擇檔案」控制項之位置
                //iwe.Click();

                ////移動到「選擇檔案」按鈕
                //Task.Run(() => { SendKeys.SendWait("{tab 10}"); });

                Clipboard.SetText(downloadImgFullName);

                //等待選取檔案對話框開啟
                Thread.Sleep(800 + (
                    800 + Extend_the_wait_time_for_the_Open_Old_File_dialog_box_to_appear_Millisecond < 0 ? 0 : Extend_the_wait_time_for_the_Open_Old_File_dialog_box_to_appear_Millisecond));//最小值（須在重開機後或系統最小負載時）（連「開啟」舊檔之視窗也看不見，即可完成）
                                                                                                                                                                                              //Thread.Sleep(1200);
                                                                                                                                                                                              //Thread.Sleep(500);            


                //輸入：檔案名稱 //SendKeys.Send(downloadImgFullName);
                SendKeys.Send("+{Insert}~");//or "^v"
                                            //SendKeys.Send("{ENTER}");
                Clipboard.Clear();

            ////图片预览
            //iwe = waitFindWebElementBySelector_ToBeClickable("#batch_figure_0");
            //while (iwe == null)
            //{
            //    iwe = waitFindWebElementBySelector_ToBeClickable("#batch_figure_0");
            //    //提前結束用
            //    if (Clipboard.GetText() != string.Empty)// && !Clipboard.GetText().Contains("Ctext_Page_Image"))
            //    { StopOCR = true; return false; }
            //}

            ////按下「上傳」
            //iwe = waitFindWebElementBySelector_ToBeClickable("#batchUploadDropdown");
            //iwe.Click();

            ////按下「豎排自動識別」
            //iwe = waitFindWebElementBySelector_ToBeClickable("#Batch > div.d-flex.justify-content-between.mt-3 > div > div > div:nth-child(2) > ul > li.dropdown-item > div > label");
            //iwe.Click();


            retry:
                ////按下「上傳」
                //iwe = waitFindWebElementBySelector_ToBeClickable("#batchUploadDropdown");
                //iwe.Click();

                ////按下「圖片上傳」
                //iwe = waitFindWebElementBySelector_ToBeClickable("#Batch > div.d-flex.justify-content-between.mt-3 > div > div > div:nth-child(2) > ul > li:nth-child(1) > a");
                //iwe.Click();

                //Form1.playSound(Form1.soundLike.over);

                //Thread.Sleep(800);

                //按下「編輯」
                //iwe = waitFindWebElementBySelector_ToBeClickable("#result_edit_0", 30);
                iwe = WaitFindWebElementBySelector_ToBeClickable("#result_edit_0", 1);
                while (iwe == null)
                {
                    //訊息方塊：文件 1 识别失败。Ocr failed
                    IWebElement iw = WaitFindWebElementBySelector_ToBeClickable("#swal2-html-container");
                    //<div class="swal2-html-container" id="swal2-html-container" style="display: block;">文件 1 识别失败。Ocr failed</div>
                    //<div class="swal2-html-container" id="swal2-html-container" style="display: block;">出现错误。error.</div>
                    if (iw != null && (iw.Text.Contains("文件 1 识别失败。Ocr failed") || iw.Text.Contains("出现错误。error.")))
                    {
                        //按下ok:
                        iw = WaitFindWebElementBySelector_ToBeClickable("body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled");
                        iw.Click();
                        goto retry;
                    }

                    iwe = WaitFindWebElementBySelector_ToBeClickable("#result_edit_0");
                }
                iwe.Click();

                //#batchTable > tbody > tr > td.bs-checkbox > label > input[type=checkbox]

                //Thread.Sleep(1100);
                //Thread.Sleep(1000);
                Thread.Sleep(1050);


                //按下準備完畢OK
                iwe = WaitFindWebElementBySelector_ToBeClickable("body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled", 1);
                while (iwe == null)
                    iwe = WaitFindWebElementBySelector_ToBeClickable("body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled", 1);
                iwe.Click();

                //按下「文本行」
                //【文本行】按鈕
                iwe = WaitFindWebElementBySelector_ToBeClickable("#line_image_panel > div > div:nth-child(2) > div:nth-child(8) > button:nth-child(2)");

                while (iwe == null)
                    iwe = WaitFindWebElementBySelector_ToBeClickable("#line_image_panel > div > div:nth-child(2) > div:nth-child(8) > button:nth-child(2)");

                //if (iwe == null)
                //{
                //    SendKeys.SendWait("{esc}");
                //    iwe = waitFindWebElementBySelector_ToBeClickable("#OneLine > div.d-flex.justify-content-between.mt-2.mb-1 > div:nth-child(3) > div:nth-child(6) > button:nth-child(2)");
                //}

                if (iwe != null)
                {
                    try
                    {
                        iwe.Click();
                    }
                    catch (Exception)
                    {
                        Form1.playSound(Form1.soundLike.error, true);
                        driver.SwitchTo().Window(driver.CurrentWindowHandle);
                        SendKeys.SendWait("{esc}");
                        driver.SwitchTo().Window(driver.CurrentWindowHandle);
                        SendKeys.SendWait("%r");
                        iwe.Click();
                    }
                    //文本窗口
                    iwe = WaitFindWebElementBySelector_ToBeClickable("#TextArea");
                    if (iwe != null)
                    {
                        //將OCR結果讀入剪貼簿：
                        if (iwe.Text != string.Empty)
                            Clipboard.SetText(iwe.Text);
                        else
                        {
                            StopOCR = true;
                            return false;
                        }
                    }
                    else
                    {
                        StopOCR = true;
                        return false;
                    }
                }
                else
                {
                    StopOCR = true; return false;
                }

                //finished:
                #region 關閉OCR視窗後回到原來分頁視窗
                try
                {
                    driver.Close();

                }
                catch (Exception)
                {
                }
                _OCR_GJcool_WindowClosed = true;
                try
                {
                    driver.SwitchTo().Window(currentWindowHndl);
                    SendKeys.SendWait("%r");//關閉右方頁籤：Chrome瀏覽器擴充功能所設定的快速鍵
                }
                catch (Exception)
                {
                }
                #endregion

                StopOCR = true;
                return true;
            }
            catch (Exception)
            {
                StopOCR = true; return false;
            }
        }


        /// <summary>
        /// 20240420
        /// </summary>
        /// <param name="downloadImgFullName"></param>
        /// <param name="_downloadResult"></param>
        /// <returns></returns>
        internal static bool OCR_GJcool_BatchProcessing_new(string downloadImgFullName, bool _downloadResult = false)
        {

            driver = driver ?? DriverNew();
            string currentWindowHndl = driver.CurrentWindowHandle;
            IWebElement iwe;// = null;
            try
            {
                openNewTabWindow(WindowType.Tab);
                GoToUrlandActivate(OCRSite_URL[OCRSiteTitle.GJcool]);
                //GoToUrlandActivate("https://gj.cool/try_ocr#");
                //GoToUrlandActivate("https://gj.cool/try_ocr#Batch");
                //ActiveForm1.TopMost = false;
                //按下「批量處理(new)」頁籤
                //IWebElement iwe = waitFindWebElementBySelector_ToBeClickable("#batch_Tab_A", 15);
                iwe = WaitFindWebElementBySelector_ToBeClickable("#File_Tab", 15);
                if (iwe == null) { StopOCR = true; return false; }
                iwe.Click();
            }
            catch (Exception)
            {
                //Console.WriteLine(ex.HResult + ex.Message);
                StopOCR = true; return false;
            }


            #region 檢查並刪除文件
            iwe = WaitFindWebElementBySelector_ToBeClickable("#fileTable > tbody > tr:nth-child(1) > td.bs-checkbox > label > input[type=checkbox]", 1);
            if (iwe != null)
            {
                try
                {
                    //按下選取方塊，準備刪除
                    while (iwe == null)
                        iwe = WaitFindWebElementBySelector_ToBeClickable("body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled");
                    if (!iwe.Selected)
                        iwe.Click();
                }
                catch (Exception)
                {
                    //if (!waitFindWebElementBySelector_ToBeClickable("#fileTable > tbody > tr > td.bs-checkbox > label > input[type=checkbox]").Selected
                    //    && waitFindWebElementBySelector_ToBeClickable("body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled") != null)
                    //    goto reClickOCROK;

                    throw;
                }

                //刪除文件
                try
                {
                    iwe = WaitFindWebElementBySelector_ToBeClickable("#FileUploadDropdown");
                    //按下「上傳」
                    while (iwe == null)
                        iwe = WaitFindWebElementBySelector_ToBeClickable("#FileUploadDropdown");
                    iwe.Click();
                    //按下「刪除文件」                    
                    iwe = WaitFindWebElementBySelector_ToBeClickable("#File > div.d-flex.justify-content-between.mt-2 > div.d-flex.mx-3 > div:nth-child(1) > div > ul > li:nth-child(8) > a");
                    while (iwe == null)
                        iwe = WaitFindWebElementBySelector_ToBeClickable("#File > div.d-flex.justify-content-between.mt-2 > div.d-flex.mx-3 > div:nth-child(1) > div > ul > li:nth-child(8) > a");
                    iwe.Click();

                }
                catch (Exception)
                {
                    //if (!waitFindWebElementBySelector_ToBeClickable("#fileTable > tbody > tr > td.bs-checkbox > label > input[type=checkbox]").Selected
                    //    && waitFindWebElementBySelector_ToBeClickable("body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled") != null)
                    //    goto reClickOCROK;
                    throw;
                }
                Thread.Sleep(800);
            }
            #endregion //以上檢查並刪除文件

            //等待書圖檔下載完成
            DateTime dt = DateTime.Now;

        reUpload:
            //按下「上傳」
            iwe = WaitFindWebElementBySelector_ToBeClickable("#FileUploadDropdown");
            iwe.Click();

            while (!File.Exists(downloadImgFullName))
            {
                if (DateTime.Now.Subtract(dt).TotalSeconds > 38)
                {
                    if (DialogResult.Cancel == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("書圖下載完成了嗎？感恩感恩　南無阿彌陀佛　讚美主"))
                    { StopOCR = true; return false; }
                    else
                        goto reClickUploadOK;
                }
            }

            ChromeSetFocus();
            driver.SwitchTo().Window(driver.CurrentWindowHandle);
            //按下「上傳圖片」
            //iwe = waitFindWebElementBySelector_ToBeClickable("#File > div.d-flex.justify-content-between.mt-2 > div.d-flex.mx-3 > div:nth-child(1) > div > ul > li:nth-child(3) > a");

            iwe = WaitFindWebElementBySelector_ToBeClickable("#File > div.d-flex.justify-content-between.mt-2 > div.d-flex.mx-3 > div:nth-child(1) > div > ul > li:nth-child(4) > a");
            try
            {
                iwe.Click();
            }
            catch (Exception)
            {
                StopOCR = true;
                return false;
            }

            //clickCopybutton_GjcoolFastExperience(new Point(iwe.Location.X + 76 + (iwe.Size.Width) / 2, iwe.Location.Y + 120 + (iwe.Size.Height) / 2));//new Point(X, Y)=「選擇檔案」控制項之位置
            //iwe.Click();

            try
            {
                Clipboard.SetText(downloadImgFullName);
            }
            catch (Exception)
            {
            }

            //等待選取檔案對話框開啟
            //Thread.Sleep(800 + (
            Thread.Sleep(1600 + (
                800 + Extend_the_wait_time_for_the_Open_Old_File_dialog_box_to_appear_Millisecond < 0 ? 0 : Extend_the_wait_time_for_the_Open_Old_File_dialog_box_to_appear_Millisecond));//最小值（須在重開機後或系統最小負載時）（連「開啟」舊檔之視窗也看不見，即可完成）
                                                                                                                                                                                          //Thread.Sleep(1200);
                                                                                                                                                                                          //Thread.Sleep(500);            


            //輸入：檔案名稱 //SendKeys.Send(downloadImgFullName);
            SendKeys.SendWait("+{Insert}");//or "^v"
            Thread.Sleep(250);
            SendKeys.SendWait("{ENTER}");
            //Clipboard.Clear();


            //retry:
            ////按下「上傳」
            //iwe = waitFindWebElementBySelector_ToBeClickable("#batchUploadDropdown");
            //iwe.Click();

            ////按下「圖片上傳」
            //iwe = waitFindWebElementBySelector_ToBeClickable("#Batch > div.d-flex.justify-content-between.mt-3 > div > div > div:nth-child(2) > ul > li:nth-child(1) > a");
            //iwe.Click();

            //Form1.playSound(Form1.soundLike.over);

            Thread.Sleep(1400);
        reClickUploadOK:
            Byte reRunOCRTimer = 0;//避免虛耗額度，白白浪費 20240622
            try
            {
                //等待「上傳完成」訊息方塊出現
                iwe = WaitFindWebElementBySelector_ToBeClickable("#swal2-title", 0.1);
                dt = DateTime.Now;
                while (iwe == null)
                {
                    try
                    {
                        if (WaitFindWebElementBySelector_ToBeClickable("#fileTable > tbody > tr > td:nth-child(4)", 0.1)?.GetAttribute("textContent") != string.Empty
                            && WaitFindWebElementBySelector_ToBeClickable("#fileTable > tbody > tr > td:nth-child(7)", 0.1)?.GetAttribute("textContent") == " ") goto reRunOCR;
                    }
                    catch (Exception)
                    {
                    }
                    iwe = WaitFindWebElementBySelector_ToBeClickable("#swal2-title", 0.1);
                    if (DateTime.Now.Subtract(dt).TotalSeconds > 8)
                        if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("等待「上傳完成」訊息方塊出現、已逾時，要繼續嗎？") == DialogResult.Cancel)
                        { StopOCR = true; return false; }
                        else
                        {
                            ActiveForm1.TopMost = false;
                            driver.SwitchTo().Window(driver.CurrentWindowHandle);
                            BringToFront("chrome");
                        }
                }

                dt = DateTime.Now;
                while (iwe?.GetAttribute("textContent") != "上传完成")
                {
                    if (iwe?.GetAttribute("textContent") == "OCR完成") goto reClickOCROK;
                    if (DateTime.Now.Subtract(dt).TotalSeconds > 16)
                        //if (DateTime.Now.Subtract(dt).TotalSeconds > 68)
                        if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("等候上傳完成已逾時，要繼續嗎？") == DialogResult.Cancel)
                        { StopOCR = true; return false; }
                        else
                        {
                            ActiveForm1.TopMost = false;
                            driver.SwitchTo().Window(driver.CurrentWindowHandle);
                            BringToFront("chrome");
                        }
                }
                //按下「上傳完成」按鈕
                iwe = WaitFindWebElementBySelector_ToBeClickable("body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled", 0.1);
                while (iwe == null)
                {
                    try
                    {
                        if (WaitFindWebElementBySelector_ToBeClickable("#fileTable > tbody > tr > td:nth-child(4)", 0.1)?.GetAttribute("textContent") != string.Empty
                            && WaitFindWebElementBySelector_ToBeClickable("#fileTable > tbody > tr > td:nth-child(7)", 0.1)?.GetAttribute("textContent") == " "
                            && WaitFindWebElementBySelector_ToBeClickable("body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled", 0.1) == null) goto reRunOCR;
                    }
                    catch (Exception)
                    {
                    }
                    iwe = WaitFindWebElementBySelector_ToBeClickable("body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled", 0.1);
                }
                //{
                ////訊息方塊：文件 1 识别失败。Ocr failed
                //IWebElement iw = waitFindWebElementBySelector_ToBeClickable("body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled");

                ////<div class="swal2-html-container" id="swal2-html-container" style="display: block;">文件 1 识别失败。Ocr failed</div>
                ////<div class="swal2-html-container" id="swal2-html-container" style="display: block;">出现错误。error.</div>
                //if (iw != null && (iw.Text.Contains("文件 1 识别失败。Ocr failed") || iw.Text.Contains("出现错误。error.")))
                //{
                //    //按下ok:
                //    iw = waitFindWebElementBySelector_ToBeClickable("body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled");
                //    iw.Click();
                //    goto retry;
                //}

                //iwe = waitFindWebElementBySelector_ToBeClickable("#result_edit_0");
                //}

                iwe.Click();//按下「上傳完成」按鈕
            }
            catch (Exception)
            {
                goto reClickUploadOK;
                throw;
            }

            #region  檢查確實已上傳：
            DateTime dddt = DateTime.Now;
            try
            {
                iwe = driver.FindElement(By.XPath("/html/body/div[13]/div/div[2]/div[2]/div[1]/div[2]/div[2]/table/tbody/tr/td[6]/div/div"));
            }
            catch (Exception)
            {
                return false;
            }
            while (iwe == null)
            {
                iwe = driver.FindElement(By.XPath("/html/body/div[13]/div/div[2]/div[2]/div[1]/div[2]/div[2]/table/tbody/tr/td[6]/div/div"));
                if (DateTime.Now.Subtract(dddt).TotalSeconds > 14)
                    goto reUpload;
            }
            if (iwe.GetAttribute("textContent") != "已上传")
                //Debugger.Break();
                goto reUpload;
            #endregion

            Clipboard.Clear();

        reRunOCR:
            //Byte reRunOCRTimer = 0;//避免虛耗額度，白白浪費 20240622
            reRunOCRTimer++;
            if (reRunOCRTimer > 3) { Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("OCR結果有誤，請予檢查！"); StopOCR = true; return false; }

            //按下選取方塊，準備OCR
            try
            {
                string tx = WaitFindWebElementBySelector_ToBeClickable("#fileTable > tbody > tr > td:nth-child(7)", 0.1)?.GetAttribute("textContent");
                if (!tx.IsNullOrEmpty() && tx != " ")
                {
                    //System.Diagnostics.Debugger.Break();
                    goto copyResult;
                }
                while (WaitFindWebElementBySelector_ToBeClickable("#fileTable > tbody > tr > td:nth-child(4)", 0.1)?.GetAttribute("textContent") == string.Empty) ;
                iwe = WaitFindWebElementBySelector_ToBeClickable("#fileTable > tbody > tr:nth-child(1) > td.bs-checkbox > label > input[type=checkbox]", 0.1);
                while (iwe == null)
                    iwe = WaitFindWebElementBySelector_ToBeClickable("body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled", 0.1);
                dt = DateTime.Now;
                while (!iwe.Selected)
                {
                    if (DateTime.Now.Subtract(dt).TotalSeconds > 10)
                        if (DialogResult.Cancel == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("勾選檔案準備OCR作業逾時，是否繼續？")) { StopOCR = true; return false; }
                        else
                        {
                            ActiveForm1.TopMost = false;
                            driver.SwitchTo().Window(driver.CurrentWindowHandle);
                            BringToFront("chrome");
                        }
                    iwe.Click();
                }


            }
            catch (Exception)
            {
                goto reClickUploadOK;
                throw;
            }

            //按下「執行命令」按鈕
            try
            {
                iwe = WaitFindWebElementBySelector_ToBeClickable("#OcrFileDropdown", 0.1);

                iwe.Click();
            }
            catch (Exception)
            {
                goto reClickUploadOK;
                throw;
            }

            //按下「OCR」按鈕
            try
            {
                iwe = WaitFindWebElementBySelector_ToBeClickable("#File > div.d-flex.justify-content-between.mt-2 > div.d-flex.mx-3 > div.d-flex.mx-2.align-self-center > div:nth-child(1) > ul > li:nth-child(2) > a", 0.1);
                iwe.Click();

            }
            catch (Exception)
            {
                if (WaitFindWebElementBySelector_ToBeClickable("#fileTable > tbody > tr > td:nth-child(4)", 0.1).GetAttribute("textContent") != string.Empty)
                    goto reClickUploadOK;
                throw;
            }
            Thread.Sleep(2200);
        reClickOCROK:
            try
            {
                dt = DateTime.Now;
                //等待「OCR完成 OK」出現
                iwe = WaitFindWebElementBySelector_ToBeClickable("#swal2-title", 0.1);
                while (iwe == null)
                {
                    if (WaitFindWebElementBySelector_ToBeClickable("#fileTable > tbody > tr > td.bs-checkbox > label > input[type=checkbox]", 0.1)?.Selected == false
                        && WaitFindWebElementBySelector_ToBeClickable("#fileTable > tbody > tr > td:nth-child(7)", 0.1)?.GetAttribute("textContent") == " ") goto reRunOCR;
                    iwe = WaitFindWebElementBySelector_ToBeClickable("#swal2-title", 0.1);
                    if (DateTime.Now.Subtract(dt).TotalSeconds > 45)
                        if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("等待OCR完成已逾時，是否繼續？") == DialogResult.Cancel) { StopOCR = true; return false; }
                        else
                        {
                            ActiveForm1.TopMost = false;
                            driver.SwitchTo().Window(driver.CurrentWindowHandle);
                            BringToFront("chrome");
                        }
                }
                while (iwe.GetAttribute("textContent") != "OCR完成")
                {
                    if (DateTime.Now.Subtract(dt).TotalSeconds > 45)
                        if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("等待OCR完成已逾時，是否繼續？") == DialogResult.Cancel) { StopOCR = true; return false; }
                        else
                        {
                            ActiveForm1.TopMost = false;
                            driver.SwitchTo().Window(driver.CurrentWindowHandle);
                            BringToFront("chrome");
                        }
                }
                //按下「OCR完成 OK」按鈕
                iwe = WaitFindWebElementBySelector_ToBeClickable("body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled", 0.1);
                while (iwe == null)
                {
                    try
                    {
                        if (WaitFindWebElementBySelector_ToBeClickable("#fileTable > tbody > tr > td.bs-checkbox > label > input[type=checkbox]", 0.1)?.Selected == false
                            && WaitFindWebElementBySelector_ToBeClickable("#fileTable > tbody > tr > td:nth-child(7)", 0.1)?.GetAttribute("textContent") == " ") goto reRunOCR;

                    }
                    catch (Exception)
                    {
                    }
                    iwe = WaitFindWebElementBySelector_ToBeClickable("body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled", 0.1);
                    //提前結束用
                    if (Clipboard.GetText() != string.Empty)// && !Clipboard.GetText().Contains("Ctext_Page_Image"))
                    { StopOCR = true; return false; }
                }
                try
                {
                    //訊息方塊：成功: 0, 失败: 1
                    IWebElement iw = WaitFindWebElementBySelector_ToBeClickable("#swal2-html-container", 0.1);
                    bool stopProcess = false;

                    if (iw != null)
                    {
                        if (iw.Text.Contains("成功: 0, 失败: 1")) stopProcess = true;
                        if (iw.Text.Contains("成功: 1, 失败: 0")) iwe.Click();

                    }

                    if (stopProcess) { StopOCR = true; return false; }

                }
                catch (Exception)
                {

                    throw;
                }
            }
            catch (Exception)
            {
                IWebElement iewMsgBox = WaitFindWebElementBySelector_ToBeClickable("#swal2-title", 0.1);
                IWebElement iew = WaitFindWebElementBySelector_ToBeClickable("#fileTable > tbody > tr > td.bs-checkbox > label > input[type=checkbox]", 0.1);
                if (iew == null)
                {
                    if (iewMsgBox != null)
                    {
                        iewMsgBox = WaitFindWebElementBySelector_ToBeClickable("#swal2-title", 0.1);
                        if (iewMsgBox.GetAttribute("textContent").Contains("OCR完成")) goto reClickOCROK;
                    }
                    goto reUpload;
                }
                else if (WaitFindWebElementBySelector_ToBeClickable("body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled", 0.1) != null)
                    goto reClickOCROK;
                //有選取項目，且其「文本」欄位值非空
                //else if (iew?.Selected == false
                //    && waitFindWebElementBySelector_ToBeClickable("#fileTable > tbody > tr > td:nth-child(7)")?.GetAttribute("textContent") == " ")
                //    goto reRunOCR;
                else if (WaitFindWebElementBySelector_ToBeClickable("body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled", 0.1) != null)
                    goto reClickUploadOK;
                else
                    goto reClickOCROK;

                throw;
            }

        ////按下「上傳」
        //iwe = waitFindWebElementBySelector_ToBeClickable("#batchUploadDropdown");
        //iwe.Click();

        ////按下「豎排自動識別」
        //iwe = waitFindWebElementBySelector_ToBeClickable("#Batch > div.d-flex.justify-content-between.mt-3 > div > div > div:nth-child(2) > ul > li.dropdown-item > div > label");
        //iwe.Click();

        //Thread.Sleep(1100);
        //Thread.Sleep(1000);
        //Thread.Sleep(1050);


        #region 複製OCR結果
        copyResult:
            try
            {
                // 取得網頁元素的 title 屬性值 
                iwe = WaitFindWebElementBySelector_ToBeClickable("#fileTable > tbody > tr:nth-child(1) > td:nth-child(7) > div", 0.1);
                if (iwe == null)
                    goto reClickOCROK;
                //while (iwe == null)
                //    iwe = waitFindWebElementBySelector_ToBeClickable("#fileTable > tbody > tr:nth-child(1) > td:nth-child(7) > div");
                string ocrResult = iwe.GetAttribute("title");
                if (ocrResult.IsNullOrEmpty())
                {
                    DateTime dtt = DateTime.Now;
                    //檢查「OCR完成 OK」按鈕
                    while (WaitFindWebElementBySelector_ToBeClickable("body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled", 0.1) == null)
                    { if (DateTime.Now.Subtract(dtt).TotalSeconds > 45) break; }
                    iwe = WaitFindWebElementBySelector_ToBeClickable("body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled", 0.1);
                    if (iwe == null)
                    {
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("OCR結果是空字串，請檢查！");
                        if (WaitFindWebElementBySelector_ToBeClickable("#fileTable > tbody > tr:nth-child(1) > td:nth-child(7) > div", 0.1)?.GetAttribute("title").IsNullOrEmpty() == false) { goto copyResult; }
                    }
                    else
                    //按下「OCR完成 OK」按鈕                        
                    {
                        if (WaitFindWebElementBySelector_ToBeClickable("#fileTable > tbody > tr:nth-child(1) > td:nth-child(7) > div", 0.1)?.GetAttribute("title").IsNullOrEmpty() == false) { iwe.Click(); goto copyResult; }
                    }
                    StopOCR = true; return false;
                }


                if (ocrResult.IndexOf(" ") > -1)
                {
                    #region 檢查覆查追查用，記下原來OCR傳回的樣子 20240902
                    Console.WriteLine(ocrResult);
                    #endregion
                    //Debugger.Break();
                    CnText.Spaces2Braces(ref ocrResult);
                }
                if (ocrResult.IndexOf("【") > -1)
                {
                    #region 檢查覆查追查用，記下原來OCR傳回的樣子 20240902
                    Console.WriteLine(ocrResult);
                    #endregion

                    //ocrResult = ocrResult.Replace("】【", string.Empty).Replace("【", "{{").Replace("】", "}}");
                    //CnText.RemoveInnerBraces(ref ocrResult);
                    ocrResult = ocrResult.Replace("】【", string.Empty);
                    ocrResult = CnText.RemoveNestedBrackets(ocrResult);
                    ocrResult = ocrResult.Replace("【", "{{").Replace("】", "}}").Replace("{{}}", string.Empty);
                }

                //Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("copy to clipboard!");
                try
                {
                    Clipboard.SetText(ocrResult);
                }
                catch (Exception)
                {

                    //Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("無法複製OCR結果！");
                    //StopOCR = true; return false;
                }
                //.Replace("0","◯").Replace("〇", "◯"));

                //Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("copied to clipboard!");
            }
            catch (Exception ex)
            {
                if (ex.HResult == -2147467261 && ex.Message.StartsWith("值不能為 null。")) { Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("OCR結果是空字串，請檢查！"); StopOCR = true; return false; }
                goto reClickOCROK;
                throw;
            }
            #endregion

            //Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("ready to delete doc");

            #region 刪除文件
            try
            {

                //按下選取方塊，準備刪除
                //dt = DateTime.Now;
                //iwe = waitFindWebElementBySelector_ToBeClickable("#fileTable > tbody > tr:nth-child(1) > td.bs-checkbox > label > input[type=checkbox]", 1);
                iwe = WaitFindWebElementBySelector_ToBeClickable("#fileTable > thead > tr > th.bs-checkbox > div.th-inner > label > input[type=checkbox]", 0.1);
                while (iwe == null)
                {
                    //if (DateTime.Now.Subtract(dt).TotalSeconds > 30)
                    //    if (DialogResult.Cancel == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("「選取方塊」等候逾時，是否繼續？")) { StopOCR = true; return false; }
                    iwe = WaitFindWebElementBySelector_ToBeClickable("#fileTable > thead > tr > th.bs-checkbox > div.th-inner > label > input[type=checkbox]", 0.1);
                }

                if (!iwe.Selected)
                {
                    driver.SwitchTo().Window(driver.CurrentWindowHandle);
                    iwe.Click();
                }
            }
            catch (Exception)
            {
                if (!WaitFindWebElementBySelector_ToBeClickable("#fileTable > tbody > tr > td.bs-checkbox > label > input[type=checkbox]", 0.1).Selected
                    && WaitFindWebElementBySelector_ToBeClickable("body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled", 0.1) != null)
                    goto reClickOCROK;
                else
                {
                    driver.SwitchTo().Window(driver.CurrentWindowHandle);
                    iwe.Click();
                    //throw;
                }
            }

            //刪除文件
            try
            {
                //dt = DateTime.Now;
                iwe = WaitFindWebElementBySelector_ToBeClickable("#FileUploadDropdown", 0.1);
                //按下「上傳」
                while (iwe == null)
                {
                    //if (DateTime.Now.Subtract(dt).TotalSeconds > 30)
                    //    if (DialogResult.Cancel == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("「上傳」等候逾時，是否繼續？")) { StopOCR = true; return false; }
                    iwe = WaitFindWebElementBySelector_ToBeClickable("#FileUploadDropdown", 0.1);
                }
                driver.SwitchTo().Window(driver.CurrentWindowHandle);
                iwe.Click();
                //按下「刪除文件」
                //dt = DateTime.Now;
                iwe = WaitFindWebElementBySelector_ToBeClickable("#File > div.d-flex.justify-content-between.mt-2 > div.d-flex.mx-3 > div:nth-child(1) > div > ul > li:nth-child(8) > a", 0.1);
                while (iwe == null)
                {
                    //if (DateTime.Now.Subtract(dt).TotalSeconds > 30)
                    //    if (DialogResult.Cancel == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("「刪除文件」等候逾時，是否繼續？")) { StopOCR = true; return false; }
                    iwe = WaitFindWebElementBySelector_ToBeClickable("#File > div.d-flex.justify-content-between.mt-2 > div.d-flex.mx-3 > div:nth-child(1) > div > ul > li:nth-child(8) > a", 0.1);
                }
                driver.SwitchTo().Window(driver.CurrentWindowHandle);
                iwe.Click();

            }
            catch (Exception)
            {
                if (!WaitFindWebElementBySelector_ToBeClickable("#fileTable > tbody > tr > td.bs-checkbox > label > input[type=checkbox]", 0.1).Selected
                    && WaitFindWebElementBySelector_ToBeClickable("body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled", 0.1) != null)
                    goto reClickOCROK;
                throw;
            }
            //Thread.Sleep(300);
            dt = DateTime.Now;
            iwe = WaitFindWebElementBySelector_ToBeClickable("#swal2-title", 0.02);
            while (iwe == null || WaitFindWebElementBySelector_ToBeClickable("#swal2-html-container", 0.02) == null)
            {
                try
                {
                    if (iwe != null && iwe.GetAttribute("textContent") == "删除完成") break;
                }
                catch (Exception)
                {
                }
                if (DateTime.Now.Subtract(dt).TotalSeconds > 10)
                {
                    if (DialogResult.Cancel == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("等待刪除完成訊息逾時，是否繼續？")) { StopOCR = true; return false; }
                    else
                    {
                        ActiveForm1.TopMost = false;
                        driver.SwitchTo().Window(driver.CurrentWindowHandle);
                        BringToFront("chrome");
                    }
                }

                iwe = WaitFindWebElementBySelector_ToBeClickable("#swal2-title", 0.02);

            }

            //if (iwe != null)
            //{
            //    dt = DateTime.Now;
            //    while (iwe.GetAttribute("textContent") != "删除完成")
            //    {
            //        if (DateTime.Now.Subtract(dt).TotalSeconds > 10)
            //        {
            //            if (DialogResult.Cancel == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("等待刪除完成訊息逾時，是否繼續？")) { StopOCR = true; return false; }
            //        }
            //    }
            //}

            ////不必等待刪除完成按下OK按鈕了！！ 20240420
            ////按下「刪除完成 OK」按鈕
            //iwe = waitFindWebElementBySelector_ToBeClickable("body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled");
            //while (iwe == null)
            //    iwe = waitFindWebElementBySelector_ToBeClickable("body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled");
            ////if (iwe == null)
            ////{
            ////    SendKeys.SendWait("{esc}");
            ////    iwe = waitFindWebElementBySelector_ToBeClickable("#OneLine > div.d-flex.justify-content-between.mt-2.mb-1 > div:nth-child(3) > div:nth-child(6) > button:nth-child(2)");
            ////}

            //if (iwe != null)
            //{
            //    try
            //    {
            //        iwe.Click();
            //    }
            //    catch (Exception)
            //    {
            //        Form1.playSound(Form1.soundLike.error, true);
            //        driver.SwitchTo().Window(driver.CurrentWindowHandle);
            //        SendKeys.SendWait("{esc}");
            //        driver.SwitchTo().Window(driver.CurrentWindowHandle);
            //        SendKeys.SendWait("%r");
            //        iwe.Click();
            //    }
            //    ////文本窗口
            //    //iwe = waitFindWebElementBySelector_ToBeClickable("#TextArea");
            //    //if (iwe != null)
            //    //{
            //    //    //將OCR結果讀入剪貼簿：
            //    //    if (iwe.Text != string.Empty)
            //    //        Clipboard.SetText(iwe.Text);
            //    //    else
            //    //    {
            //    //        StopOCR = true;
            //    //        return false;
            //    //    }
            //    //}
            //    //else
            //    //{
            //    //    StopOCR = true;
            //    //    return false;
            //    //}
            //}
            //else
            //{
            //    StopOCR = true; return false;
            //}

            #endregion //以上刪除文件


            //finished:
            #region 關閉OCR視窗後回到原來分頁視窗
            try
            {
                driver.Close();

            }
            catch (Exception)
            {
            }
            _OCR_GJcool_WindowClosed = true;
            try
            {
                driver.SwitchTo().Window(currentWindowHndl);
                SendKeys.SendWait("%r");//關閉右方頁籤：Chrome瀏覽器擴充功能所設定的快速鍵
            }
            catch (Exception)
            {
            }
            #endregion

            StopOCR = true;
            return true;
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex.HResult + ex.Message);
            //    StopOCR = true; return false;
            //}

        }

        /// <summary>
        /// 《古籍酷》OCR：自動識別(豎版)。由原本程式碼改良而來
        /// 以直接點按【文本行】按鈕來將所開啟之文本窗口中的文本框，其內之文字內容選入剪貼簿中，以省卻下載、讀寫硬碟之煩，
        /// 且不必擔心視窗不在作用中（即可切換到其他視窗放心工作，順待OCR結果取回）感恩感恩　讚歎讚歎　南無阿彌陀佛 20231028
        /// </summary>
        /// <param name="downloadImgFullName">書圖檔全檔名</param>
        /// <param name="_downloadResult">若欲用原式（下載OCR結果）則為true。預設為改良式，並不下載，只讀入剪貼簿，故為false</param>
        /// <returns>順利完成則回傳true</returns>
        internal static bool OCR_GJcool_AutoRecognizeVertical(string downloadImgFullName, bool _downloadResult = false)
        {
            //Form1.playSound(Form1.soundLike.press);
            string gjCool = string.Empty; string currentWindowHndl = ""; WindowType windowType = WindowType.Tab;
            try
            {
                driver = driver ?? DriverNew();
                currentWindowHndl = driver.CurrentWindowHandle;
                //openNewTabWindow(WindowType.Window);
                openNewTabWindow(windowType);
                _OCR_GJcool_WindowClosed = false;

                #region 方便提早取消作業（藉由關閉OCR視窗）
                try
                {
                    if (currentWindowHndl != driver.CurrentWindowHandle) { }
                    ;
                }
                catch (Exception)
                {
                    StopOCR = true;
                    return false;
                }
                #endregion


                ActiveForm1.TopMost = false;
                //點數（算力值、算力配额）不足逕用「快速體驗」執行
                if (!OCR_GJcool_AccountChanged && waitGJcoolPoint || (!OCR_GJcool_AccountChanged && waitGJcoolPoint && DateTime.Now.Subtract(gjCoolPointLess150When) < gjCoolPointEnoughTimespan))
                {
                    //Form1.playSound(Form1.soundLike.processing);
                    bool fastXResulut = OCR_GJcool_FastExperience(downloadImgFullName);
                    //driver.Close();
                    driver?.Close();
                    _OCR_GJcool_WindowClosed = true;
                    try
                    {
                        driver?.SwitchTo().Window(currentWindowHndl);
                    }
                    catch (Exception)
                    {
                        StopOCR = true;
                        return false;
                    }
                    StopOCR = true;
                    return fastXResulut;
                }
                else
                    gjCool = OCRSite_URL[OCRSiteTitle.GJcool]; //"https://gj.cool/try_ocr";
                                                               //Form1.playSound(Form1.soundLike.processing);
                if (_OCR_GJcool_AccountChanged) { _OCR_GJcool_AccountChanged = !_OCR_GJcool_AccountChanged; gjCoolPointLess150When = DateTime.Now; }

                #region 方便提早取消作業（藉由關閉OCR視窗）
                try
                {
                    if (currentWindowHndl != driver.CurrentWindowHandle) { }
                    ;
                }
                catch (Exception)
                {
                    StopOCR = true;
                    return false;
                }
                #endregion

            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case -2146233088://"no such window: target window already closed\nfrom unknown error: web view not found\n  (Session info: chrome=109.0.5414.120)"                            
                        if (ex.Message.IndexOf("no such window: target window already closed") > -1)
                        {
                            #region OCR成功後則刪除下載的書圖,備份OCR結果; 因為 https://gj.cool/try_ocr 頁面時常傳回假資料（之前曾識別的文本），故今改寫在 textBox3.TextChanged事件中
                            //if (File.Exists(downloadImgFullName))
                            //{
                            //    /* -2147024864：
                            //     * "由於另一個處理序正在使用檔案 'C:\\Users\\oscar\\Dropbox\\Ctext_Page_Image.png'，所以無法存取該檔案。"
                            //     * 20231028 Bing大菩薩： 刪除檔案問題：
                            //     * 您好，這是Bing。您在刪除檔案時遇到的問題可能是因為該檔案正在被另一個程序使用。以下是一些可能的解決方案：
                            //     * 使用垃圾回收1：在刪除檔案之前，您可以嘗試調用 System.GC.Collect() 和 System.GC.WaitForPendingFinalizers() 來強制進行垃圾回收。
                            //     */
                            //    //System.GC.Collect();
                            //    //System.GC.WaitForPendingFinalizers();
                            //    //以上無效！！！
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
                            //                    Thread.Sleep(400);//"由於另一個處理序正在使用檔案 'X:\\Ctext_Page_Image.txt'，所以無法存取該檔案。"
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
                        }
                        else if (ex.Message.StartsWith("The HTTP request to the remote WebDriver server for URL"))
                        {
                            StopOCR = true; return false;
                        }
                        else
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                        break;
                    case -2147221040:
                        if (ex.Message.IndexOf("要求的剪貼簿作業失敗。") > -1)
                        {
                            DateTime dt = DateTime.Now;
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請手動複製OCR文本");
                            //Clipboard.Clear();//此時剪貼簿還不能用
                            while (!Form1.isClipBoardAvailable_Text(100)) { if (DateTime.Now.Subtract(dt).Seconds > 0.5) break; }
                            driver.Close();
                            _OCR_GJcool_WindowClosed = true;
                            driver.SwitchTo().Window(currentWindowHndl);
                            StopOCR = true;
                            return true;
                        }
                        else
                        {
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                            Debugger.Break();
                        }
                        break;

                    default:
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                        break;
                }

            }

            #region 方便提早取消作業（藉由關閉OCR視窗）
            try
            {
                if (driver == null)
                {
                    StopOCR = true;
                    return false;
                }
                if (currentWindowHndl != driver.CurrentWindowHandle) { }
                ;
            }
            catch (Exception)
            {
                StopOCR = true;
                return false;
            }
            #endregion



            if (gjCool == string.Empty)
            {
                StopOCR = true; return false;
            }
        reNavigate:
            try
            {
                driver.Navigate().GoToUrl(gjCool);
            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case -2146233088:
                        if (ex.Message.IndexOf("no such window: target window already closed") > -1
                            || ex.Message.EndsWith("seconds."))
                        {
                            try
                            {
                                driver.SwitchTo().Window(LastValidWindow);
                            }
                            catch (Exception)
                            {
                                //throw;
                            }
                            StopOCR = true;
                            return false;
                        }
                        else if (ex.Message.StartsWith("timeout: Timed out receiving message from renderer: "))//timeout: Timed out receiving message from renderer: 3.559
                                                                                                               //(Session info: chrome = 130.0.6723.70)
                        {
                            if (driver.Manage().Timeouts().PageLoad < new TimeSpan(0, 2, 0))
                            {
                                driver.Manage().Timeouts().PageLoad = driver.Manage().Timeouts().PageLoad.Add(new TimeSpan(0, 1, 0));
                                goto reNavigate;
                            }
                            else
                                goto default;
                        }
                        else
                            goto default;
                    default:
                        string msgText = ex.HResult.ToString() + ex.Message;
                        Console.WriteLine(msgText);
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(msgText);
                        StopOCR = true;
                        return false;
                }
            }

            #region 先檢查點數是否足夠
            const byte pointCoin = 120;// 150;//「自動識別(豎版)」所需點數120，載入圖檔要30//20250525： 現在改為上傳圖片20，OCR 100
            IWebElement iwe = null;
            try
            {//以備隨時被使用者關閉
             //IWebElement iwe = waitFindWebElementBySelector_ToBeClickable("#compute-value");
             //IWebElement iwe = driver.FindElement(By.CssSelector("#compute-value"));
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(_chromeDriverServiceTimeSpan));
                try
                {
                    if (wait != null)
                    {
                        while (true)
                        {
                            try
                            {
                                driver.SwitchTo().Window(driver.CurrentWindowHandle);
                            }
                            catch (Exception)
                            {
                                return false;
                            }
                            //iwe = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.CssSelector("#compute-value")));
                            iwe = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.CssSelector("#compute-value")));
                            if (iwe != null) break;
                        }
                    }

                }
                catch (Exception)
                {
                    StopOCR = true;
                    return false;
                }
                //Form1.playSound(Form1.soundLike.processing);
            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case -2147467261:
                        if (ex.Message.IndexOf("並未將物件參考設定為物件的執行個體。") > -1)
                        {
                            StopOCR = true;
                            return false;
                        }
                        else
                        {
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                            Debugger.Break();
                        }
                        break;
                    case -2146233088:
                        if (ex.Message.IndexOf("Timed out after") > -1)//"Timed out after 30.5 seconds"
                        {//"The HTTP request to the remote WebDriver server for URL http://localhost:5837/session/0e0cfa1c2cdcd0298a952b8267079906/element timed out after 30.5 seconds."
                            driver.SwitchTo().Window(LastValidWindow);
                            StopOCR = true;
                            return false;
                        }
                        else
                        {
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                            Console.WriteLine(ex.HResult + ex.Message);
                            Debugger.Break();
                            driver.SwitchTo().Window(LastValidWindow);
                        }
                        break;
                    default:
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                        Debugger.Break();
                        break;
                }
                //throw;
            }
            int points = 0;
            if (iwe != null)
            {
                //取得點數，如「 117 / 1000」格式
                string innerText = iwe.GetAttribute("innerText");
                if (innerText.IndexOf(" /") > -1 && " ".Length + innerText.IndexOf(" /") - " ".Length <= innerText.Length)
                    //int.TryParse(innerText.Substring(" ".Length, innerText.IndexOf(" /") - " ".Length), out points);
                    int.TryParse(innerText.Substring(0, innerText.IndexOf(" /")), out points);
                //當點數不足時：
                if (points < pointCoin)
                {
                    Form1.playSound(Form1.soundLike.processing);
                    waitGJcoolPoint = true;
                    gjCoolPointLess150When = DateTime.Now;
                    ////登出帳戶： 以下這會跳出訊息方塊，得處理，故改為傳引數的方式
                    //waitFindWebElementBySelector_ToBeClickable("#navbarNav > ul:nth-child(2) > li:nth-child(2) > a > p.mb-0.fs-6.fst-italic").Click();
                    //Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("點數（算力配额）不足！目前僅有"+ points + " 至少需要"+pointCoin);
                    //轉由首頁「快速體驗」執行
                    //bool fastXResulut = OCR_GJcool_FastExperience(downloadImgFullName, true);
                    bool fastXResulut = OCR_GJcool_FastExperience(downloadImgFullName, false);//申請「數字萬舟」計劃通過，取得批量處理授權帳戶後故
                    if (fastXResulut)
                    {
                        try
                        {
                            driver.Close();
                        }
                        catch (Exception)
                        {
                        }
                    }
                    _OCR_GJcool_WindowClosed = true;
                    try
                    {
                        driver?.SwitchTo().Window(currentWindowHndl);
                    }
                    catch (Exception ex)
                    {
                        switch (ex.HResult)
                        {
                            case -2146233088:
                                //"no such window: target window already closed
                                if (ex.Message.IndexOf("no such window") > -1)
                                {
                                    openNewTabWindow();
                                    GoToUrlandActivate(frm.textBox3Text);
                                }
                                else
                                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                                break;
                            default:
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                                break;
                        }
                    }
                    StopOCR = true;
                    return fastXResulut;

                }
                else
                {//點數足時：
                 //此帳戶下的最後一次
                    if (points - pointCoin < pointCoin)
                        //Form1.playSound(Form1.soundLike.stop);
                        using (SoundPlayer sp = new SoundPlayer("C:\\Windows\\Media\\chord.wav")) { sp.Play(); }
                    waitGJcoolPoint = false; innerText = null;
                }//釋放記憶體
            }
            else { StopOCR = true; return false; }
            #endregion

            #region 下載路徑取得
            //需要下載時才執行
            string filePath = string.Empty;
            if (_downloadResult)
            {
                #region 再檢查瀏覽器下載目錄並取得 ：
                string downloadDirectory = DownloadDirectory_Chrome;
                if (!ChkDownloadDirectory_Chrome(downloadImgFullName, downloadDirectory)) { StopOCR = true; return false; }
                #endregion

                //取得所匯出的檔案路徑
                filePath = Path.Combine(downloadDirectory, Path.GetFileNameWithoutExtension(downloadImgFullName) + ".txt");//@"X:\Ctext_Page_Image.txt";
                                                                                                                           //刪除之前的檔案，以免因檔案存在而被下載端重新命名
                Task.Run(() =>
                {
                    if (File.Exists(filePath)) File.Delete(filePath);
                });
            }
            #endregion


            #region 等「新增圖片」按鈕可按：選擇檔案
            //Thread.Sleep(3200);
            //等待「選擇檔案」控制項出現，最多等timeSpanSecs秒；
            //為免tab鍵數不同，而須手動操作，以免表單遮住畫面:
            ActiveForm1.TopMost = false;
            Clipboard.Clear();
            DateTime begin = DateTime.Now; const int timeSpanSecs = 30;
            TimeSpan timeSpan = new TimeSpan();
            iwe = WaitFindWebElementBySelector_ToBeClickable("#line_img_form > div > input[type=file]");
            while (iwe == null)
            {
                //iwe = waitFindWebElementBySelector_ToBeClickable("#line_img_form > div > input[type=file]");
                //iwe = WaitFindWebElementBySelector_ToBeClickable("#OneLine > div.d-flex.mt-2 > div:nth-child(1) > div.ps-1.pe-2.align-self-center > button > i");                
                iwe = WaitFindWebElementBySelector_ToBeClickable("#OneLine > div.d-flex.mt-2 > div:nth-child(1) > div:nth-child(2) > button");
                timeSpan = (DateTime.Now.Subtract(begin));
                if (timeSpan.TotalSeconds > timeSpanSecs ||
                    Clipboard.GetText() != string.Empty) { StopOCR = true; return false; }
            }

            //檢查使用者是否已關閉視窗，取消這次的操作（比如說才發現已經有OCR了、或弄錯頁了……等等，可逕接關閉《古籍酷》OCR視窗以終結之）
            try
            {
                if (driver.CurrentWindowHandle == currentWindowHndl) { }
                ;
            }
            catch (Exception)
            {
                StopOCR = true;
                return false;
            }

            //取得「選擇檔案」控制項
            //if (windowType == WindowType.Window)
            //{
            //    //SendKeys.Send("{tab 16} ");
            //    SendKeys.Send("{tab 16}");
            //    //如果按下tab鍵16次後「選擇檔案」控制項沒有被選中（不同環境下網頁元件數可能會有所不同！）
            //    //這種寫法應該不會成功，因為Selenium可應用的範圍是程式自動化操作而不是使用者手動manual操作者20230322 果然！ 13:47
            //    //if (!iwe.Selected) Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("菩薩慈悲：請手動執行OCR，OCR完成之後程式會接手執行。感恩感恩　南無阿彌陀佛"+ Environment.NewLine +"按下「確定（OK）」後繼續…… 阿彌陀佛");            
            //    //if (!iwe.Selected) System.Diagnostics.Debugger.Break();
            //}
            //else if (windowType == WindowType.Tab)
            //    SendKeys.Send("{tab 13}");
            //    //SendKeys.Send("{tab 13}");
            //else
            //    Debugger.Break();
            //    Debugger.Break();

            //if (iwe.Selected)// "line_img")
            //按下「選擇檔案」控制項
            //    SendKeys.Send(" ");
            //else
            //if(iwe.Selected)
            //iwe.Submit();
            //iwe.Click();//不行，會出錯

            //欲提早結束時：
            if (Clipboard.GetText().IndexOf(Environment.NewLine + Environment.NewLine) > -1) goto finished;

            try
            {
                driver.SwitchTo().Window(driver.CurrentWindowHandle);
            }
            catch (Exception)
            {
                return false;
            }


            //driver.SwitchTo().Window(driver.CurrentWindowHandle);//切換到目前Selenium操控的視窗，就不怕沒及時得到焦點而失誤了
            try
            {
                iwe.Click();//不行，會出錯;20240326新版不會出現錯誤了
            }
            catch (Exception exx)
            {
                Console.WriteLine(exx.HResult + exx.Message);
                throw;
            }
            //try
            //{
            //    //iwe = driver.FindElement(By.XPath("/html/body/div[13]/div/div[1]/div[1]/div[1]/form/div/input"));
            //    iwe = driver.FindElement(By.XPath("//*[@id=\"line_img_form\"]/div/input"));
            //    //wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            //    //  "//*[@id=\"line_img_form\"]/div/input"
            //    //iwe.Click();
            //    iwe.SendKeys(OpenQA.Selenium.Keys.End);
            //}
            //catch (Exception ex11)
            //{
            //    Console.WriteLine(ex11.HResult + ex11.Message);
            //    Debugger.Break();
            //    //throw;
            //}
            #endregion



            #region 點擊新增圖片按鈕並輸入書圖全檔名
            ////clickCopybutton_GjcoolFastExperience(new Point(137, 299), Form1.soundLike.press);//new Point(X, Y)=「選擇檔案」控制項之位置
            //clickCopybutton_GjcoolFastExperience(new Point(iwe.Location.X + 76 + (iwe.Size.Width) / 2, iwe.Location.Y + 120 + (iwe.Size.Height) / 2),
            //    points - pointCoin < pointCoin ? Form1.soundLike.none : Form1.soundLike.press);//new Point(X, Y)=「選擇檔案」控制項之位置
            //points = 0;//釋放記憶體
            ////76 系統工具列在左側時的寬度//120 Chrome瀏覽器頂遄到書籤列下端的長度

            ////waitFindWebElementBySelector_ToBeClickable("#line_img_form > div > input[type=file]").SendKeys(OpenQA.Selenium.Keys.Space);
            ////waitFindWebElementByName_ToBeClickable("line_img",2).Submit();
            byte tryTimes = 1;

            //等待書圖檔下載完成
            DateTime ddt = DateTime.Now;
            while (!File.Exists(downloadImgFullName))
            {
                if (DateTime.Now.Subtract(ddt).TotalSeconds > 38) { StopOCR = true; return false; }
            }
            try
            {
                Clipboard.SetText(downloadImgFullName);
            }
            catch (Exception)
            {
            }


            //等待選取檔案對話框開啟
            Thread.Sleep(1600 + (
                800 + Extend_the_wait_time_for_the_Open_Old_File_dialog_box_to_appear_Millisecond < 0 ? 0 : Extend_the_wait_time_for_the_Open_Old_File_dialog_box_to_appear_Millisecond));//最小值（須在重開機後或系統最小負載時）（連「開啟」舊檔之視窗也看不見，即可完成）
                                                                                                                                                                                          //Thread.Sleep(1200);
                                                                                                                                                                                          //Thread.Sleep(500);            

            //輸入：檔案名稱 //SendKeys.Send(downloadImgFullName);
            //retry:
            SendKeys.Send("+{Insert}~");//or "^v"
                                        //SendKeys.Send("{ENTER}");
                                        //Form1.playSound(Form1.soundLike.processing);

            if (_downloadResult)
                Form1.playSound(Form1.soundLike.processing);
            else
                //Form1.playSound(Form1.soundLike.waiting);
                using (SoundPlayer sp = new SoundPlayer("C:\\Windows\\Media\\ring08.wav")) { sp.Play(); }

            //待圖載入完畢：
            //Thread.Sleep(3220);
            //Thread.Sleep(1220);
            Thread.Sleep(920);
            //Thread.Sleep(1920);
            #endregion

            DateTime dtimr = DateTime.Now;
        redo:

            #region「上傳完畢」對話方塊的「OK」按鈕 20231103
            try
            {
                iwe = WaitFindWebElementBySelector_ToBeClickable
                //("body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled", 0.2);
                ("body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled", 3.3);

            }
            catch (Exception)
            {
            }
            //if (iwe == null) return false;
            //try
            //{
            //    iwe.Click();
            //}
            //catch (Exception)
            //{
            //    if (tryTimes > 5) return false;
            //    tryTimes++;
            //    goto redo;
            //    //throw;
            //}
            //tryTimes = 0;
            //if (iwe == null) return false;

            //由頁籤數判斷，是否在上傳完畢後，OCR進行時，按下Chrome瀏覽器快速鍵 Alt + r 以關閉右邊的頁籤。因為有時很卡、甚至連線中斷，須開啟多個《古籍酷》頁面以刺激其連上
            int tabCount = driver.WindowHandles.Count;

            //while (iwe == null && waitFindWebElementBySelector_ToBeClickable("#auto_ocr") == null)
            while (iwe == null)// && waitFindWebElementBySelector_ToBeClickable("#auto_ocr") == null)
            {
                try
                {
                    driver.SwitchTo().Window(driver.CurrentWindowHandle);
                }
                catch (Exception)
                {
                    return false;
                }

                if (Clipboard.GetText().IndexOf(Environment.NewLine + Environment.NewLine) > -1) goto finished;
                Thread.Sleep(150);
                try
                {
                    iwe = WaitFindWebElementBySelector_ToBeClickable
                                    ("body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled", 0.3);

                }
                catch (Exception)
                {
                }
                if (DateTime.Now.Subtract(dtimr).TotalSeconds > OCR_wait_time_Top_Limit＿second + 50) { StopOCR = true; return false; }
            }
            if (iwe != null)
            {
                try
                {
                    if (tryTimes < 51)
                        driver.SwitchTo().Window(driver.CurrentWindowHandle);//切換到目前Selenium操控的視窗，就不怕沒及時得到焦點而失誤了
                    iwe.Click();//點擊「上傳完畢」對話方塊的「OK」按鈕 
                }
                catch (Exception)
                {
                    if (tryTimes == 0) Form1.playSound(Form1.soundLike.error);
                    //if (tryTimes % 50 == 0)//> 50)
                    if (tryTimes % 50 == 0 && DateTime.Now.Subtract(dtimr).TotalSeconds > OCR_wait_time_Top_Limit＿second)//> 50)
                    {
                        if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("已超時，是否繼續等候？") == DialogResult.Cancel)
                        {
                            ActiveForm1.TopMost = false;
                            StopOCR = true;
                            _OCR_GJcool_WindowClosed = true;
                            driver.Close();
                            /* 20231114 Bing大菩薩： C# Selenium 警告訊息關閉                             
                             */
                            try
                            {
                                driver.SwitchTo().Alert().Accept();//.SendKeys(OpenQA.Selenium.Keys.Enter);
                            }
                            catch (Exception)
                            {
                            }
                            return false;
                        }

                    }
                    ActiveForm1.TopMost = false;
                    tryTimes++;
                    Thread.Sleep(100);
                    goto redo;
                    //throw;
                }
                tryTimes = 0;
            }
            #endregion
            else
            {
                StopOCR = true;
                return false;
            }

            #region 按下「Pro」
            //iwe = waitFindWebElementBySelector_ToBeClickable("#auto_ocr");
            iwe = WaitFindWebElementBySelector_ToBeClickable("#auto_ocr");
            //if (iwe == null)
            while (iwe == null)
            {
                //tryTimes++;
                //if (tryTimes > 5) return false;
                //goto retry;
                if (Clipboard.GetText().IndexOf(Environment.NewLine + Environment.NewLine) > -1) goto finished;

                Thread.Sleep(250);
                iwe = WaitFindWebElementBySelector_ToBeClickable("#auto_ocr");
                if (DateTime.Now.Subtract(dtimr).Seconds > 20)
                {
                    StopOCR = true; return false;
                }
            }
            driver.SwitchTo().Window(driver.CurrentWindowHandle);//切換到目前Selenium操控的視窗，就不怕沒及時得到焦點而失誤了
            try
            {
                iwe.Click();
                Form1.playSound(Form1.soundLike.processing);
            }
            catch (Exception)
            {
                goto redo;
            }
            //SendKeys.Send("{tab}~");
            //iwe.SendKeys(" ");//(OpenQA.Selenium.Keys.Enter);
            //clickCopybutton_GjcoolFastExperience(new Point(iwe.Location.X + iwe.Size.Width / 2, iwe.Location.Y + iwe.Size.Height / 2), Form1.soundLike.press);
            #endregion

            #region 按下「自動識別(豎版)」（今改為「古籍識別」20251109），開始OCR……
            //SendKeys.Send("{down}~");            
            //iwe = waitFindWebElementBySelector_ToBeClickable("#OneLine > div.d-flex.justify-content-between.mt-2.mb-1 > div:nth-child(1) > div:nth-child(2) > ul > li:nth-child(2) > button");
            iwe = WaitFindWebElementBySelector_ToBeClickable("#OneLine > div.d-flex.mt-2 > div:nth-child(1) > div:nth-child(3) > ul > li:nth-child(2) > button");
            driver.SwitchTo().Window(driver.CurrentWindowHandle);
            try
            {
                iwe.Click();
                //現在會跑出對話方塊
                Thread.Sleep(200);
                SendKeys.SendWait("~");
            }
            catch (Exception)
            {
                goto redo;
                //throw;
            }
            #endregion

            //若頁籤數較上傳時為多，則按下Chrome瀏覽器快速鍵 Alt + r 以關閉右側頁籤（當是輔助連線用所開啟者） 20231228
            if (driver.WindowHandles.Count > tabCount) SendKeys.SendWait("%r");

            //等待OCR，上限為30秒
            //iwe = waitFindWebElementBySelector_ToBeClickable("# OneLine > div.d-flex.justify-content-between.mt-2.mb-1 > div:nth-child(1) > div:nth-child(2) > ul > li:nth-child(2) > button");
            iwe = WaitFindWebElementBySelector_ToBeClickable("#line_list_table > tbody > tr:nth-child(1) > td:nth-child(2)");
            begin = DateTime.Now;
            while (iwe == null)
            {
                if (Clipboard.GetText().IndexOf(Environment.NewLine + Environment.NewLine) > -1) goto finished;

                iwe = WaitFindWebElementBySelector_ToBeClickable("#line_list_table > tbody > tr:nth-child(1) > td:nth-child(2)");
                //上限為30秒
                if (DateTime.Now.Subtract(begin).TotalSeconds > timeSpanSecs)
                {
                    StopOCR = true; return false;
                }
            }
            //iwe.Click();
            //Thread.Sleep(6220);
            //Thread.Sleep(6000);//OCR結束
            //Thread.Sleep(10500);//OCR結束
            Thread.Sleep(1100);//OCR結束

            if (_downloadResult)
            {

                //SendKeys.Send("{tab 24}~");
                SendKeys.Send("{tab 18}~");
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
                //等待下載完成
                TimeSpan waitFileExitTimeSpan = TimeSpan.FromSeconds(10);//最多等待時間
                DateTime waitFileExitBegin = DateTime.Now;
                while (!File.Exists(filePath) && DateTime.Now.Subtract(waitFileExitBegin) < waitFileExitTimeSpan)
                {
                }
                if (File.Exists(filePath))
                {

                    string text = null;
                retryReadFile:
                    try
                    {
                        text = File.ReadAllText(filePath, System.Text.Encoding.UTF8);
                    }
                    catch (Exception ex)
                    {
                        switch (ex.HResult)
                        {
                            case -2147024864:
                                Thread.Sleep(3);//"由於另一個處理序正在使用檔案 'X:\\Ctext_Page_Image.txt'，所以無法存取該檔案。"
                                text = File.ReadAllText(filePath, System.Text.Encoding.UTF8);
                                goto retryReadFile;
                            default:
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.Message);
                                StopOCR = true;
                                return false;
                        }

                    }

                    if (text == null)
                    {
                        if (DialogResult.Cancel == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly
                                ("文本未下成功，請自行下載，好了之後再按「確定」繼續完成讀入到textBox1的工作；      或按「取消」結束此次操作。\n\r感恩感恩　南無阿彌陀佛"))
                        {
                            StopOCR = true; return false;
                        }
                        Clipboard.SetText(text.Replace("\n", "\r\n"));
                        //刪除下載檔案，以便下次載入
                        Task.Run(() =>
                        {
                            //不必備份，似乎是《古籍酷》本身有bug，明明是一張圖，輸出的卻是不一樣的文本，或上次或之前哪次的文本 20230802
                            //File.Copy(filePath, Path.Combine(//備一份以備萬一
                            //    Path.GetDirectoryName(filePath) + Path.GetFileNameWithoutExtension(filePath) + "1" +
                            //        Path.GetExtension(filePath)), true);


                            File.Delete(filePath);
                        });
                    }
                    else
                    {
                        StopOCR = true; return false;
                    }
                    #endregion
                    #endregion
                }
            }

            #region 不下載時的處理：
            else
            {
                //【文本行】按鈕
                //iwe = waitFindWebElementBySelector_ToBeClickable("#OneLine > div.d-flex.justify-content-between.mt-2.mb-1 > div:nth-child(3) > div:nth-child(6) > button:nth-child(2) > i");
                //iwe = waitFindWebElementBySelector_ToBeClickable("#OneLine > div.d-flex.justify-content-between.mt-2.mb-1 > div:nth-child(3) > div:nth-child(6) > button:nth-child(2)");
                //iwe = waitFindWebElementBySelector_ToBeClickable("#line_image_panel > div > div:nth-child(2) > div:nth-child(8) > button:nth-child(2) > i");
                //iwe = WaitFindWebElementBySelector_ToBeClickable("#line_image_panel > div > div:nth-child(2) > div:nth-child(8) > button:nth-child(2)");
                //iwe = WaitFindWebElementBySelector_ToBeClickable("#line_image_panel > div > div:nth-child(2) > div.ps-3.align-self-center > button:nth-child(2)");
                iwe = WaitFindWebElementBySelector_ToBeClickable("#line_image_panel > div > div.d-flex.gap-2 > div:nth-child(6) > button:nth-child(2)");
                //if (iwe == null)
                //{
                //    SendKeys.SendWait("{esc}");
                //    iwe = waitFindWebElementBySelector_ToBeClickable("#OneLine > div.d-flex.justify-content-between.mt-2.mb-1 > div:nth-child(3) > div:nth-child(6) > button:nth-child(2)");
                //}
                if (iwe != null)
                {
                    try
                    {
                        iwe.Click();
                    }
                    catch (Exception)
                    {
                        Form1.playSound(Form1.soundLike.error, true);
                        driver.SwitchTo().Window(driver.CurrentWindowHandle);
                        SendKeys.SendWait("{esc}");
                        iwe.Click();
                    }
                    //文本窗口
                    iwe = WaitFindWebElementBySelector_ToBeClickable("#TextArea");
                    if (iwe != null)
                    {
                        //將OCR結果讀入剪貼簿：
                        if (iwe.Text != string.Empty)
                            Clipboard.SetText(iwe.Text);
                        else
                        {
                            StopOCR = true;
                            return false;
                        }
                    }
                    else
                    {
                        StopOCR = true;
                        return false;
                    }
                }
                else
                {
                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("「文本行」按鈕的Selector值已變，請重新取得！");
                    StopOCR = true; return false;
                }
            }


        #endregion


        finished:
            #region 關閉OCR視窗後回到原來分頁視窗
            try
            {
                driver.Close();
            }
            catch (Exception)
            {
            }
            _OCR_GJcool_WindowClosed = true;
            try
            {
                driver.SwitchTo().Window(currentWindowHndl);
            }
            catch (Exception)
            {
            }
            #endregion
            StopOCR = true;
            return true;
        }

        #region OCR_GJcool_AutoRecognizeVertical 原式
        /*
        /// <summary>
        /// 《古籍酷》OCR：自動識別(豎版)【此乃原式20231028】
        /// </summary>
        /// <param name="downloadImgFullName">書圖檔全檔名</param>
        /// <returns>順利完成則回傳true</returns>
        internal static bool OCR_GJcool_AutoRecognizeVertical_DownloadResult(string downloadImgFullName)
        {
            //Form1.playSound(Form1.soundLike.press);
            string gjCool = string.Empty; string currentWindowHndl = ""; WindowType windowType = WindowType.Tab;
            try
            {
                driver = driver ?? DriverNew();
                currentWindowHndl = driver.CurrentWindowHandle;
                //openNewTabWindow(WindowType.Window);
                openNewTabWindow(windowType);
                _OCR_GJcool_WindowClosed = false;
                if (ActiveForm1.TopMost) ActiveForm1.TopMost = false;
                //點數（算力值、算力配额）不足逕用「快速體驗」執行
                if (!OCR_GJcool_AccountChanged && waitGJcoolPoint || (!OCR_GJcool_AccountChanged && waitGJcoolPoint && DateTime.Now.Subtract(gjCoolPointLess150When) < gjCoolPointEnoughTimespan))
                {
                    //Form1.playSound(Form1.soundLike.processing);
                    bool fastXResulut = OCR_GJcool_FastExperience(downloadImgFullName);
                    driver.Close();
                    _OCR_GJcool_WindowClosed = true;
                    driver.SwitchTo().Window(currentWindowHndl);
                    return fastXResulut;
                }
                else
                    gjCool = OCRSite_URL[OCRSiteTitle.GJcool]; //"https://gj.cool/try_ocr";
                //Form1.playSound(Form1.soundLike.processing);
                if (_OCR_GJcool_AccountChanged) { _OCR_GJcool_AccountChanged = !_OCR_GJcool_AccountChanged; gjCoolPointLess150When = DateTime.Now; }
            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case -2146233088://"no such window: target window already closed\nfrom unknown error: web view not found\n  (Session info: chrome=109.0.5414.120)"                            
                        if (ex.Message.IndexOf("no such window: target window already closed") > -1) { if (File.Exists(downloadImgFullName)) File.Delete(downloadImgFullName); }
                        else
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                        break;
                    case -2147221040:
                        if (ex.Message.IndexOf("要求的剪貼簿作業失敗。") > -1)
                        {
                            DateTime dt = DateTime.Now;
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請手動複製OCR文本");
                            //Clipboard.Clear();//此時剪貼簿還不能用
                            while (!Form1.isClipBoardAvailable_Text(100)) { if (DateTime.Now.Subtract(dt).Seconds > 0.5) break; }
                            driver.Close();
                            _OCR_GJcool_WindowClosed = true;
                            driver.SwitchTo().Window(currentWindowHndl);
                            return true;
                        }
                        else
                        {
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                            Debugger.Break();
                        }
                        break;

                    default:
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                        break;
                }

            }

            try
            {
                driver.Navigate().GoToUrl(gjCool);
            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case -2146233088:
                        if (ex.Message.IndexOf("no such window: target window already closed") > -1
                            || ex.Message.EndsWith("seconds."))
                        {
                            driver.SwitchTo().Window(LastValidWindow);
                            return false;
                        }
                        else
                            goto default;
                    default:
                        string msgText = ex.HResult.ToString() + ex.Message;
                        Console.WriteLine(msgText);
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(msgText);
                        return false;
                }
            }

            #region 先檢查點數是否足夠
            const byte pointCoin = 150;//「自動識別(豎版)」所需點數120，載入圖檔要30
            IWebElement iwe = null;
            try
            {//以備隨時被使用者關閉
                //IWebElement iwe = waitFindWebElementBySelector_ToBeClickable("#compute-value");
                //IWebElement iwe = driver.FindElement(By.CssSelector("#compute-value"));
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(_chromeDriverServiceTimeSpan));
                iwe = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.CssSelector("#compute-value")));
                //Form1.playSound(Form1.soundLike.processing);
            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case -2147467261:
                        if (ex.Message.IndexOf("並未將物件參考設定為物件的執行個體。") > -1)
                            return false;
                        else
                        {
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                            Debugger.Break();
                        }
                        break;
                    case -2146233088:
                        if (ex.Message.IndexOf("Timed out after") > -1)//"Timed out after 30.5 seconds"
                        {//"The HTTP request to the remote WebDriver server for URL http://localhost:5837/session/0e0cfa1c2cdcd0298a952b8267079906/element timed out after 30.5 seconds."
                            driver.SwitchTo().Window(LastValidWindow);
                            return false;
                        }
                        else
                        {
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                            Console.WriteLine(ex.HResult + ex.Message);
                            Debugger.Break();
                            driver.SwitchTo().Window(LastValidWindow);
                        }
                        break;
                    default:
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                        Debugger.Break();
                        break;
                }
                //throw;
            }
            if (iwe != null)
            {
                Form1.playSound(Form1.soundLike.processing);
                //取得點數，如「 117 / 1000」格式
                string innerText = iwe.GetAttribute("innerText"); int points = 0;
                if (innerText.IndexOf(" /") > -1 && " ".Length + innerText.IndexOf(" /") - " ".Length <= innerText.Length)
                    int.TryParse(innerText.Substring(" ".Length, innerText.IndexOf(" /") - " ".Length), out points);
                if (points < pointCoin)
                {
                    waitGJcoolPoint = true;
                    gjCoolPointLess150When = DateTime.Now;
                    ////登出帳戶： 以下這會跳出訊息方塊，得處理，故改為傳引數的方式
                    //waitFindWebElementBySelector_ToBeClickable("#navbarNav > ul:nth-child(2) > li:nth-child(2) > a > p.mb-0.fs-6.fst-italic").Click();
                    //Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("點數（算力配额）不足！目前僅有"+ points + " 至少需要"+pointCoin);
                    //轉由首頁「快速體驗」執行
                    bool fastXResulut = OCR_GJcool_FastExperience(downloadImgFullName, true);
                    if (fastXResulut) driver.Close(); _OCR_GJcool_WindowClosed = true;
                    try
                    {
                        driver.SwitchTo().Window(currentWindowHndl);
                    }
                    catch (Exception ex)
                    {
                        switch (ex.HResult)
                        {
                            case -2146233088:
                                //"no such window: target window already closed
                                if (ex.Message.IndexOf("no such window") > -1)
                                {
                                    openNewTabWindow();
                                    GoToUrlandActivate(frm.textBox3Text);
                                }
                                else
                                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                                break;
                            default:
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                                break;
                        }
                    }
                    return fastXResulut;

                }
                else { waitGJcoolPoint = false; points = 0; innerText = null; }//釋放記憶體
            }
            else return false;
            #endregion

            #region 再檢查瀏覽器下載目錄並取得 ：
            string downloadDirectory = DownloadDirectory_Chrome;
            if (!ChkDownloadDirectory_Chrome(downloadImgFullName, downloadDirectory)) return false;
            #endregion

            //取得所匯出的檔案路徑
            string filePath = Path.Combine(downloadDirectory, Path.GetFileNameWithoutExtension(downloadImgFullName) + ".txt");//@"X:\Ctext_Page_Image.txt";
            //刪除之前的檔案，以免因檔案存在而被下載端重新命名
            Task.Run(() =>
            {
                if (File.Exists(filePath)) File.Delete(filePath);
            });

            //等「新增圖片」按鈕可按：選擇檔案
            //Thread.Sleep(3200);
            //等待「選擇檔案」控制項出現，最多等timeSpanSecs秒；
            //為免tab鍵數不同，而須手動操作，以免表單遮住畫面:
            if (ActiveForm1.TopMost) ActiveForm1.TopMost = false;
            DateTime begin = DateTime.Now; const int timeSpanSecs = 30;
            TimeSpan timeSpan = new TimeSpan();
            iwe = waitFindWebElementBySelector_ToBeClickable("#line_img_form > div > input[type=file]");
            while (iwe == null)
            {
                iwe = waitFindWebElementBySelector_ToBeClickable("#line_img_form > div > input[type=file]");
                timeSpan = (DateTime.Now.Subtract(begin));
                if (timeSpan.TotalSeconds > timeSpanSecs) return false;
            }

            //檢查使用者是否已關閉視窗，取消這次的操作（比如說才發現已經有OCR了、或弄錯頁了……等等，可逕接關閉《古籍酷》OCR視窗以終結之）
            try
            {
                if (driver.CurrentWindowHandle == currentWindowHndl) { };
            }
            catch (Exception)
            {
                return false;
            }

            //取得「選擇檔案」控制項
            //if (windowType == WindowType.Window)
            //{
            //    //SendKeys.Send("{tab 16} ");
            //    SendKeys.Send("{tab 16}");
            //    //如果按下tab鍵16次後「選擇檔案」控制項沒有被選中（不同環境下網頁元件數可能會有所不同！）
            //    //這種寫法應該不會成功，因為Selenium可應用的範圍是程式自動化操作而不是使用者手動manual操作者20230322 果然！ 13:47
            //    //if (!iwe.Selected) Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("菩薩慈悲：請手動執行OCR，OCR完成之後程式會接手執行。感恩感恩　南無阿彌陀佛"+ Environment.NewLine +"按下「確定（OK）」後繼續…… 阿彌陀佛");            
            //    //if (!iwe.Selected) System.Diagnostics.Debugger.Break();
            //}
            //else if (windowType == WindowType.Tab)
            //    SendKeys.Send("{tab 13}");
            //    //SendKeys.Send("{tab 13}");
            //else
            //    Debugger.Break();
            //    Debugger.Break();

            //if (iwe.Selected)// "line_img")
            //按下「選擇檔案」控制項
            //    SendKeys.Send(" ");
            //else
            //if(iwe.Selected)
            //iwe.Submit();
            //iwe.Click();//不行，會出錯

            //clickCopybutton_GjcoolFastExperience(new Point(137, 299), Form1.soundLike.press);//new Point(X, Y)=「選擇檔案」控制項之位置
            clickCopybutton_GjcoolFastExperience(new Point(iwe.Location.X + 76 + (iwe.Size.Width) / 2, iwe.Location.Y + 120 + (iwe.Size.Height) / 2), Form1.soundLike.press);//new Point(X, Y)=「選擇檔案」控制項之位置
                                                                                                                                                                             //76 系統工具列在左側時的寬度//120 Chrome瀏覽器頂遄到書籤列下端的長度

            //waitFindWebElementBySelector_ToBeClickable("#line_img_form > div > input[type=file]").SendKeys(OpenQA.Selenium.Keys.Space);
            //waitFindWebElementByName_ToBeClickable("line_img",2).Submit();
            //等待選取檔案對話框開啟
            Thread.Sleep(1200);
            //輸入：檔案名稱 //SendKeys.Send(downloadImgFullName);
            Clipboard.SetText(downloadImgFullName);
            byte tryTimes = 1;
        retry:
            SendKeys.Send("+{Insert}");//or "^v"
            SendKeys.Send("{ENTER}");
            Form1.playSound(Form1.soundLike.processing);
            //待圖載入完畢：
            //Thread.Sleep(3220);
            //Thread.Sleep(1220);
            Thread.Sleep(920);
            //按下「Pro」
            iwe = waitFindWebElementBySelector_ToBeClickable("#line_img_form > div > input[type=file]");
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
            Form1.playSound(Form1.soundLike.processing);
            //等待OCR，上限為30秒
            //iwe = waitFindWebElementBySelector_ToBeClickable("# OneLine > div.d-flex.justify-content-between.mt-2.mb-1 > div:nth-child(1) > div:nth-child(2) > ul > li:nth-child(2) > button");
            iwe = waitFindWebElementBySelector_ToBeClickable("#line_list_table > tbody > tr:nth-child(1) > td:nth-child(2)");
            begin = DateTime.Now;
            while (iwe == null)
            {
                iwe = waitFindWebElementBySelector_ToBeClickable("#line_list_table > tbody > tr:nth-child(1) > td:nth-child(2)");
                //上限為30秒
                if (DateTime.Now.Subtract(begin).TotalSeconds > timeSpanSecs) return false;
            }
            //iwe.Click();
            //Thread.Sleep(6220);
            //Thread.Sleep(6000);//OCR結束
            //Thread.Sleep(10500);//OCR結束
            Thread.Sleep(1100);//OCR結束

            //SendKeys.Send("{tab 24}~");
            SendKeys.Send("{tab 18}~");
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
            //等待下載完成
            TimeSpan waitFileExitTimeSpan = TimeSpan.FromSeconds(10);//最多等待時間
            DateTime waitFileExitBegin = DateTime.Now;
            while (!File.Exists(filePath) && DateTime.Now.Subtract(waitFileExitBegin) < waitFileExitTimeSpan)
            {
            }
            if (File.Exists(filePath))
            {

                string text = null;
            retryReadFile:
                try
                {
                    text = File.ReadAllText(filePath, System.Text.Encoding.UTF8);
                }
                catch (Exception ex)
                {
                    switch (ex.HResult)
                    {
                        case -2147024864:
                            Thread.Sleep(3);//"由於另一個處理序正在使用檔案 'X:\\Ctext_Page_Image.txt'，所以無法存取該檔案。"
                            text = File.ReadAllText(filePath, System.Text.Encoding.UTF8);
                            goto retryReadFile;
                        default:
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.Message);
                            break;
                    }

                }

                if (text == null)
                {
                    if (DialogResult.Cancel == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly
                            ("文本未下成功，請自行下載，好了之後再按「確定」繼續完成讀入到textBox1的工作；      或按「取消」結束此次操作。\n\r感恩感恩　南無阿彌陀佛"))
                        return false;
                }
                Clipboard.SetText(text.Replace("\n", "\r\n"));
                //刪除下載檔案，以便下次載入
                Task.Run(() =>
                {
                    //不必備份，似乎是《古籍酷》本身有bug，明明是一張圖，輸出的卻是不一樣的文本，或上次或之前哪次的文本 20230802
                    //File.Copy(filePath, Path.Combine(//備一份以備萬一
                    //    Path.GetDirectoryName(filePath) + Path.GetFileNameWithoutExtension(filePath) + "1" +
                    //        Path.GetExtension(filePath)), true);


                    File.Delete(filePath);
                });
            }
            else return false;
            #endregion
            #endregion

            #region 關閉OCR視窗後回到原來分頁視窗
            driver.Close(); _OCR_GJcool_WindowClosed = true;
            driver.SwitchTo().Window(currentWindowHndl);
            #endregion
            return true;
        }
        */
        #endregion

        /// <summary>
        /// 檢查Chrome瀏覽器的下載路徑。
        /// </summary>
        /// <param name="downloadImgFullName">準備下載的檔案之全檔名</param>
        /// <param name="downloadDirectory">Chrome瀏覽器的下載路徑</param>
        /// <returns>合格則傳回true</returns>
        internal static bool ChkDownloadDirectory_Chrome(string downloadImgFullName, string downloadDirectory)
        {
            // 注入 JavaScript 代碼以獲取下載目錄
            //string downloadDirectory = (string)driver.ExecuteScript("return window.navigator.userAgent.toLowerCase().indexOf('win') > -1 ? window.localStorage.getItem('download.default_directory') : null;");
            //string downloadDirectory = DownloadDirectory_Chrome;
            if (downloadDirectory.IsNullOrEmpty()) downloadDirectory = Path.GetFullPath(downloadImgFullName);
            //if (!Directory.Exists(downloadDirectory)&& !Form1.DriveExist(downloadDirectory)) { Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("菩薩慈悲："+Environment.NewLine+downloadDirectory +"並不存在！請在textBox2（尋找方塊）中輸入以指定正確的路徑。感恩感恩　南無阿彌陀佛"); return false; }
            if (!Directory.Exists(downloadDirectory)) { Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("菩薩慈悲：" + Environment.NewLine + downloadDirectory + "並不存在！請在textBox2（尋找方塊）中輸入以指定正確的路徑。感恩感恩　南無阿彌陀佛"); return false; }

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
            return true;
        }


        /// <summary>
        /// 作為中止OCR相關作業的指標
        /// 欲中止OCR作業則設定為true，預設為false
        /// </summary>
        internal static bool StopOCR { get; set; } = false;

        const int nChars = 256;
        //[DllImport("user32.dll")]
        //private static extern int GetForegroundWindow();//之前已有 [DllImport("user32.dll")]
        //public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        private static extern int GetWindowText(int hWnd, StringBuilder text, int count);
        /// <summary>
        /// Chrome瀏覽器是否為作用中的視窗（前景、最前端的應用程式視窗）
        /// 20250116 Copilot大菩薩
        /// </summary>
        /// <returns></returns>
        internal static bool IsChromeActive()
        {
            int handle = (int)GetForegroundWindow();
            StringBuilder buffer = new StringBuilder(nChars);
            if (GetWindowText(handle, buffer, nChars) > 0)
            {
                string windowTitle = buffer.ToString();
                if (windowTitle.Contains("Google Chrome"))
                {
                    Console.WriteLine("Currently focused window is Chrome");
                    return true;
                }
                else
                {
                    //Console.WriteLine("Currently focused window is not Chrome");
                    return false;
                }
            }
            return false;
        }


        /// <summary>
        /// 以《古籍酷》首頁快速體驗OCR。不計點數（算力配额）
        /// </summary>
        /// <param name="downloadImgFullName">由《中國哲學書電子化計劃》下載的書圖全檔名</param>
        /// <returns>順利完成則傳回true</returns>
        internal static bool OCR_GJcool_FastExperience(string downloadImgFullName, bool signOut = false)
        {
            #region 先檢查瀏覽器下載目錄並取得 ：
            //string downloadDirectory = DownloadDirectory_Chrome;
            //if (!ChkDownloadDirectory_Chrome(downloadImgFullName, downloadDirectory)) return false;
            #endregion
            bool returnFalse = false;
            driver = driver ?? DriverNew();
            string currentWindowHndl = driver.CurrentWindowHandle;
        reNavigate:
            const string gjCool = "https://gj.cool/";
            //openNewTabWindow(WindowType.Window);
            try
            {
                driver.Navigate().GoToUrl(gjCool);
                _OCR_GJcool_WindowClosed = false;
                if (signOut)
                {//登出帳號
                    Form1.playSound(Form1.soundLike.done);
                    WaitFindWebElementBySelector_ToBeClickable("#navbarNav > ul:nth-child(2) > li:nth-child(2) > a > p.mb-0.fs-6.fst-italic").Click();
                    //if (signOut)waitFindWebElementBySelector_ToBeClickable("#navbarNav > ul:nth-child(2) > li > a > p.mb-0.fs-6.fst-italic").Click();
                    Form1.playSound(Form1.soundLike.over);

                    //現在機制似乎改變了，故改成此式 20231228
                    returnFalse = true;
                    driver.Close();
                    Thread.Sleep(900);

                    string targetProcessName = "Proton VPN";//"ProtonVPN.exe"; // 目標程序的名稱
                                                            // 查找具有指定程式名稱的窗體
                    IntPtr targetWindowHandle = FindWindow(null, targetProcessName);
                    if (targetWindowHandle != IntPtr.Zero)
                    {
                        if (IsWindowVisible(targetWindowHandle))//使用Proton App
                            ActiveForm1.Controls["textBox2"].Text = "kk";
                        else//使用Proton 擴充功能
                            ActiveForm1.Controls["textBox2"].Text = "gg";
                    }
                    else
                        ActiveForm1.Controls["textBox2"].Text = "kk";
                    Thread.Sleep(900);
                    SendKeys.SendWait("{esc}");
                    goto finish;
                }
            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case -2146233088:
                        if (ex.Message.StartsWith("timeout: Timed out receiving message from renderer: "))//timeout: Timed out receiving message from renderer: 3.559
                                                                                                          //(Session info: chrome = 130.0.6723.70)
                        {
                            if (driver.Manage().Timeouts().PageLoad < new TimeSpan(0, 2, 0))
                            {
                                driver.Manage().Timeouts().PageLoad = driver.Manage().Timeouts().PageLoad.Add(new TimeSpan(0, 1, 0));
                                goto reNavigate;
                            }
                            else
                                goto default;
                        }
                        else
                            goto default;
                    default:
                        string msgText = ex.HResult.ToString() + ex.Message;
                        Console.WriteLine(msgText);
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(msgText);
                        //if (File.Exists(downloadImgFullName)) File.Delete(downloadImgFullName);
                        StopOCR = true;
                        return false;
                }
            }

            //C# 如何取得使用者的螢幕解析度:https://blog.xuite.net/q10814/blog/48070595 https://www.delftstack.com/zh-tw/howto/csharp/screen-size-in-csharp/
            Size Size = SystemInformation.PrimaryMonitorSize;
            int Width = SystemInformation.PrimaryMonitorSize.Width;
            int Height = SystemInformation.PrimaryMonitorSize.Height;
            //MessageBox.Show("你的螢幕解析度是" + Size + "\n Width = " + Width + "\n Height = " + Height);
            //複製按鈕的位置：20231106
            int copyBtnPosX = Width / 1920 * 835, copyBtnPosY = Height / 1200 * 730;//835,730
                                                                                    //int copyBtnPosX = (Width / 1920) * 835, copyBtnPosY = (Height / 1200) * 711;//835,711
            if (Copybutton_GjcoolFastExperience_Location.IsEmpty) Copybutton_GjcoolFastExperience_Location = new Point(copyBtnPosX, copyBtnPosY);
            //int copyBtnPosX = Width * (835 / 1920), copyBtnPosY = Height * (730 / 1200);


            //Form1.playSound(Form1.soundLike.processing);
            ActiveForm1.TopMost = false;

            //首頁「快速體驗」按鈕：
            //IWebElement iwe = WaitFindWebElementBySelector_ToBeClickable("body > div.container-fluid.bg-dark.px-1 > div > h2.text-center.my-2.py-4 > button > div", 10);
            IWebElement iwe = WaitFindWebElementBySelector_ToBeClickable("body > div.container-fluid.px-1 > div > h2.text-center.my-2.py-4 > button > div", 10);
            if (iwe == null) return false;
            Form1.playSound(Form1.soundLike.processing);
            try
            {
                driver.SwitchTo().Window(driver.CurrentWindowHandle);
                iwe.Click();//按下首頁「快速體驗」按鈕：
            }
            catch (Exception)
            {
                return false;
            }

            //「上傳 拍照」按鈕：
            iwe = WaitFindWebElementBySelector_ToBeClickable("#task-upload-btn");
            if (iwe == null) return false;
            driver.SwitchTo().Window(driver.CurrentWindowHandle);//切換到目前Selenium操控的視窗，就不怕沒及時得到焦點而失誤了
            Form1.playSound(Form1.soundLike.processing);
            Clipboard.SetText(downloadImgFullName);
            driver.SwitchTo().Window(driver.CurrentWindowHandle);
            iwe.Click();//按下「上傳 拍照」按鈕：

            //等待「開啟」檔案對話框開啟
            Thread.Sleep(1600 + (
                800 + Extend_the_wait_time_for_the_Open_Old_File_dialog_box_to_appear_Millisecond < 0 ? 0 :
                Extend_the_wait_time_for_the_Open_Old_File_dialog_box_to_appear_Millisecond));//最小值（須在重開機後或系統最小負載時）（連「開啟」舊檔之視窗也看不見，即可完成）
                                                                                              //Thread.Sleep(1200);
                                                                                              //Thread.Sleep(300);
                                                                                              //Thread.Sleep(800);//最大值（夠久了，當電腦順時會停頓一下）

            //MessageBox.Show((800 +
            //    800 + Extend_the_wait_time_for_the_Open_Old_File_dialog_box_to_appear_Millisecond < 0 ? 0 :
            //    Extend_the_wait_time_for_the_Open_Old_File_dialog_box_to_appear_Millisecond).ToString()); //jus for test

            //輸入：檔案名稱 //SendKeys.Send(downloadImgFullName);
            //貼上圖檔全名
            //byte tryTimes = 1;
            SendKeys.Send("+{Insert}~");//or "^v"
                                        //SendKeys.Send("{ENTER}");

            Form1.playSound(Form1.soundLike.processing);

            Thread.Sleep(200);
            Clipboard.Clear();

            //等待結果顯示：結果顯示元件：#text_b81d8450
            //也會變動  ！！              #text_1043686b9
            //iwe = waitFindWebElementBySelector_ToBeClickable("#text_b81d8450");

            //SendKeys.Send("{down}~");
            //等待OCR，上限為30秒
            //等待「複製」按鈕出現
            //「複製」按鈕的 BySelector 會變動！！
            //iwe = waitFindWebElementBySelector_ToBeClickable("#dialog_b81d8450 > div.col > div.d-flex.py-1 > button");
            //                                                  #dialog_483f217a > div.col > div.d-flex.py-1 > button

            //待OCR結束
            //Thread.Sleep(5200);//可多設時間以等待，若多餘，可手動按下複製按鈕即可。
            //Thread.Sleep(4300);
            //Thread.Sleep(3900);
            Thread.Sleep(1150);
            //Thread.Sleep(2950);
            #region 將OCR結果讀入剪貼簿：
            Point copyBtnPos = new Point(); DateTime begin = DateTime.Now;


            #region 複製OCR結果按鈕_嘗試用元件操作
            //20231103 Bing大菩薩：Selenium中的FindElement方法：
            //IWebElement e = driver.FindElement(By.CssSelector("div.col > div.d-flex.py-1 > button > i"));
            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(second));

            IWebElement e = null;
            WebDriverWait wait = null;
            //try
            //{
            //    e = driver.FindElement(By.XPath("//*[starts-with(@id, 'dialog_')] > div.col > div.d-flex.py-1 > button > i"));
            //    wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));
            //}
            //catch (Exception et)
            //{
            //    Console.WriteLine(et.HResult + et.Message);
            //    //throw;
            //}
            //if (e == null) Debugger.Break();


            //try
            //{
            //    e = driver.FindElement(By.XPath("//*[starts-with(@id, 'dialog_')]//div[contains(@class, 'col')]//div[contains(@class, 'd-flex py-1')]//button//i"));
            //    wait = new WebDriverWait(driver, TimeSpan.FromSeconds(4));
            //    wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(e));
            //}
            //catch (Exception)// et)
            //{
            //    //Console.WriteLine(et.HResult + et.Message);
            //    //Debugger.Break();
            //    //throw;
            //    Form1.playSound(Form1.soundLike.error);
            //}


            #region 方便提早取消作業（藉由關閉OCR視窗）
            try
            {
                if (currentWindowHndl != driver.CurrentWindowHandle) { }
                ;
            }
            catch (Exception)
            {
                StopOCR = true;
                return false;
            }
            #endregion

            DateTime dateTime = DateTime.Now; bool clicked = false, trafficLimit = false;
            //Thread.Sleep(950);

            if (Clipboard.GetText() != "") goto finish;


            //if (DateTime.Now.Subtract(dateTime).Seconds > 1 && !clicked && Clipboard.GetText() == string.Empty)
            //if (true)
            //{

            clicked = true;
            Task.Run(() =>
            {
                //按下「複製」圖示按鈕
                Task.Run(() =>
                {
                    //Form1.playSound(Form1.soundLike.info);

                    if (Copybutton_GjcoolFastExperience_Location.IsEmpty) Copybutton_GjcoolFastExperience_Location = new Point(copyBtnPosX, copyBtnPosY);//Copybutton_GjcoolFastExperience_Location = new Point(835, 711);
                                                                                                                                                         //copyBtnPos = new Point(838, 711);
                    copyBtnPos = Copybutton_GjcoolFastExperience_Location;//new Point(835, 730);
                                                                          //copyBtnPos = new Point(copyBtnPosX, copyBtnPosY);//複製按鈕的位置：20231106
                                                                          //while (Clipboard.GetText() == string.Empty && !StopOCR && !trafficLimit)
                                                                          //{//先取消20240223
                                                                          ////try
                                                                          ////{
                                                                          ////    driver.SwitchTo().Window(driver.CurrentWindowHandle);
                                                                          ////}
                                                                          ////catch (Exception)
                                                                          ////{
                                                                          ////    break;
                                                                          ////    //if (Clipboard.GetText() == string.Empty) return false;
                                                                          ////    //else goto finish;
                                                                          ////}
                                                                          ////以滑鼠座標按下複製按鈕
                                                                          //if (DateTime.Now.Subtract(dtMax).Seconds > 5 || Clipboard.GetText() != string.Empty || StopOCR) break;


                    Thread.Sleep(2950 - 1150 + 950);

                    ////bool frmActive = false;
                    ////ActiveForm1.Invoke((MethodInvoker)delegate { frmActive = ActiveForm1.Active; });
                    ////if (!frmActive)
                    ////{
                    //clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.none);
                    while (DateTime.Now.Subtract(dateTime).TotalSeconds < 6)
                    {
                        if (Clipboard.GetText() != string.Empty) break;
                        if (IsChromeActive() && IsBrowserMaximized(driver))//按下「複製」圖示按鈕
                            clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.press);
                        //Thread.Sleep(550);
                        Thread.Sleep(450);
                    }
                    ////}
                    ////else break;
                    //}
                    if (!trafficLimit) Form1.playSound(Form1.soundLike.info);
                    //Debugger.Break();
                    //if (Clipboard.GetText() != string.Empty) Application.OpenForms[0].Activate();
                });

                try
                {
                    e = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[1]/div[3]/div[2]/div[2]/button/i"));
                    //wait = new WebDriverWait(driver, TimeSpan.FromMilliseconds(150));
                    //wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(e));
                    while (e == null)
                    {
                        if (Clipboard.GetText() != string.Empty || DateTime.Now.Subtract(dateTime).TotalSeconds > 5) break;
                        e = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[1]/div[3]/div[2]/div[2]/button/i"));
                    }
                    if (e != null)
                    {
                        e.Click();
                        Thread.Sleep(255);
                    }

                }
                catch (Exception)
                {

                }

            });
            //}

            if (Clipboard.GetText() != string.Empty) goto finish;
            Form1.playSound(Form1.soundLike.info, true);//just for test
            string info;
        retry:
            //2024除夕
            try
            {
                //e = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[1]/div[3]/div[2]/div[2]/button/i"));
                e = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[1]/div[3]")).FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[1]/div[3]/div[2]/div[2]/button/i"));
                while (e == null)
                {
                    if (Clipboard.GetText() != string.Empty) break;
                    //e = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[1]/div[3]/div[2]/div[2]/button/i"));
                    e = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[1]/div[3]")).FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[1]/div[3]/div[2]/div[2]/button/i"));
                    if (DateTime.Now.Subtract(dateTime).TotalSeconds > 3 ||
                        //reach traffic limit. wait …… 訊息文字框
                        driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[1]/div[3]/div[2]/div")) != null)
                        break;
                    //Thread.Sleep(155);
                }
                if (Clipboard.GetText() != string.Empty) goto finish;
                if (e != null)
                {
                    e.Click();
                    //Form1.playSound(Form1.soundLike.press, true);
                    Form1.playSound(Form1.soundLike.info, true);//just for test
                    Thread.Sleep(455);
                }
            }
            catch (Exception)
            {
                if (DateTime.Now.Subtract(dateTime).TotalSeconds < 3)
                {
                    Thread.Sleep(200);
                    Form1.playSound(Form1.soundLike.press, true);
                    goto retry;
                }
            }
            if (Clipboard.GetText() != "") goto finish;
            try
            {
                //reach traffic limit. wait …… 訊息文字框
                e = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[1]/div[3]/div[2]/div"));
                if (e != null)
                {
                    info = e.GetAttribute("textContent");
                    if (info.StartsWith("reach traffic limit. wait ") || info.StartsWith("wait for "))
                    {
                        if (MessageBox.Show(info + Environment.NewLine + "是否要切換成批量處理模式？", "若按下【取消】，擬改用『標注平台』處理，請記得在textBox2下「gjk」指令以切換。感恩感恩　南無阿彌陀佛　讚美主",
                                                                MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly) == DialogResult.OK)
                            Form1.BatchProcessingGJcoolOCR = true;
                        else
                            Form1.BatchProcessingGJcoolOCR = false;
                        return false;
                    }
                    else if (info.StartsWith("system is busy"))
                    {//e = null; goto infos; 
                     //string ip = GetPublicIpAddress(string.Empty);
                     //Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("這個IP " + ip + " 不能用！");
                     //Clipboard.SetText(ip);
                     //ActiveForm1.Activate();
                     //ActiveForm1.BringToFront();
                        StopOCR = true; return false;
                    }
                }
            }
            catch (Exception)
            {
            }

            Form1.playSound(Form1.soundLike.over, true);

            if (Clipboard.GetText() != "") goto finish;
            else
            {
                //Form1.playSound(Form1.soundLike.processing);//just for test
                try
                {
                    if (Clipboard.GetText() != "") goto finish;
                    //e = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[1]/div[3]/div[2]/div[2]/button"));                    
                    e = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[1]/div[3]/div[2]/div[2]/button/i"));
                    //e = driver.FindElement(By.XPath("//*[starts-with(@id, 'dialog_')]//div[contains(@class, 'col')]//div[contains(@class, 'd-flex py-1')]//button//i"));
                    //wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));
                    wait = new WebDriverWait(driver, TimeSpan.FromSeconds(0.2));//20241228
                    wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(e));

                    //20241228 20241231
                    dateTime = DateTime.Now;
                    while (e != null)
                    {
                        e.Click();
                        Thread.Sleep(220);
                        e = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[1]/div[3]/div[2]/div[2]/button/i"));
                        //if (DateTime.Now.Subtract(dateTime).TotalSeconds > 13) break;
                        if (DateTime.Now.Subtract(dateTime).TotalSeconds > 13) break;
                    }

                }
                catch (Exception)
                { if (Clipboard.GetText() != "") goto finish; }
            }


            #region 測試可否快點，似無效，故省略
            //Thread.Sleep(1750);
            //try
            //{
            //    e = driver.FindElement(By.XPath("//*[starts-with(@id, 'dialog_')]//div[contains(@class, 'col')]//div[contains(@class, 'd-flex py-1')]//button//i"));
            //}
            //catch (Exception)
            //{
            //    try
            //    {
            //        e = driver.FindElement(By.XPath("//*[starts-with(@id, 'dialog_')]//div[contains(@class, 'col')]//div[contains(@class, 'd-flex py-1')]//button//i"));
            //    }
            //    catch (Exception)
            //    {

            //    }
            //}
            //try
            //{
            //    wait = new WebDriverWait(driver, TimeSpan.FromMilliseconds(1250));
            //    if (e != null && Clipboard.GetText() == string.Empty)
            //    {
            //        Form1.playSound(Form1.soundLike.exam);
            //        wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(e));
            //    }
            //}
            //catch (Exception)
            //{
            //}
            #endregion

            if (clicked && Clipboard.GetText() != string.Empty)
            {
                ActiveForm1.Activate();
                goto finish;
            }
            //infos:
            while (e == null)
            {

                if (clicked && Clipboard.GetText() != string.Empty)
                {
                    ActiveForm1.Activate();
                    goto finish;
                }

                #region 方便提早取消作業（藉由關閉OCR視窗）
                try
                {
                    if (currentWindowHndl != driver.CurrentWindowHandle) { }
                    ;
                }
                catch (Exception)
                {
                    StopOCR = true;
                    return false;
                }
                #endregion
                try
                {
                    //找出「複製」按鈕
                    //e = waitFindWebElementBySelector_ToBeClickable("#dialog_24bb81d92 > div.col > div.d-flex.py-1 > button");
                    //e = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[1]/div[3]/div[2]/div[2]/button"));
                    e = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[1]/div[3]/div[2]/div[2]/button/i"));
                    //e = driver.FindElement(By.XPath("//*[@id=\"dialog_df933239\"]/div[2]/div[2]/button"));
                    //e = driver.FindElement(By.XPath("//*[starts-with(@id, 'dialog_')]//div[contains(@class, 'col')]//div[contains(@class, 'd-flex py-1')]//button//i"));
                    wait = new WebDriverWait(driver, TimeSpan.FromMilliseconds(150));
                    wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(e));
                    e.Click();
                    Thread.Sleep(255);
                    //第 1 次好像會找不到，只好用手動了：
                    //Thread.Sleep(450);
                    //if (Clipboard.GetText() != "") goto finish;
                    //Form1.playSound(Form1.soundLike.processing);
                }
                catch (Exception)
                {
                    //throw;
                    if (Clipboard.GetText() != "") goto finish;
                    //else
                    //{

                    //    //Thread.Sleep(500);
                    //    try
                    //    {
                    //        Form1.playSound(Form1.soundLike.press);//just for test
                    //        e = driver.FindElement(By.XPath("//*[starts-with(@id, 'dialog_')]//div[contains(@class, 'col')]//div[contains(@class, 'd-flex py-1')]//button//i"));
                    //        wait = new WebDriverWait(driver, TimeSpan.FromSeconds(0.2));
                    //        wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(e));
                    //    }
                    //    catch (Exception)// et)
                    //    {
                    //        //Console.WriteLine(et.HResult + et.Message);
                    //        //Debugger.Break();
                    //        //throw;
                    //        //Form1.playSound(Form1.soundLike.error);


                    //        //if (Clipboard.GetText() != "") goto finish;
                    //        if (Clipboard.GetText().IndexOf(Environment.NewLine) > -1) goto finish;
                    //        //else return false;
                    //    }

                    //}
                }
                //Debugger.Break();
                //if (DateTime.Now.Subtract(dateTime).Seconds > 15)
                if (DateTime.Now.Subtract(dateTime).Seconds > OCR_wait_time_Top_Limit＿second)
                {
                    //if (DialogResult.Cancel == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("已超過15鈔，是否繼續？"))
                    if (DialogResult.Cancel == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("已超過" + OCR_wait_time_Top_Limit＿second + "鈔，是否繼續？"))
                    {
                        StopOCR = true; return false;
                    }
                    dateTime = DateTime.Now;
                    ActiveForm1.TopMost = false;
                    driver.SwitchTo().Window(driver.CurrentWindowHandle);
                }

                //20241217
                if (Clipboard.GetText() != string.Empty)
                {
                    ActiveForm1.Activate();
                    goto finish;
                }

                #region reach traffic limit
                IWebElement iwtext = null;
                try
                {
                    iwtext = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[1]/div[3]/div[2]/div"));
                    wait = new WebDriverWait(driver, TimeSpan.FromSeconds(0.2));
                    try
                    {
                        wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(iwtext));
                    }
                    catch (Exception)
                    {
                        //returnFalse = true;
                        //goto finish;
                    }
                }
                catch (Exception)
                {
                    try
                    {
                        iwtext = driver.FindElement(By.XPath("/html/body/div[2]/div/div/div[2]/div/div[1]/div[3]/div[2]/div"));
                        wait = new WebDriverWait(driver, TimeSpan.FromSeconds(0.2));
                        try
                        {
                            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(iwtext));
                        }
                        catch (Exception)
                        {
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
                if (Clipboard.GetText() != string.Empty)
                {
                    ActiveForm1.Activate();
                    goto finish;
                }
                if (iwtext != null)
                //textContent
                {
                    if (Clipboard.GetText() != string.Empty)
                    {
                        ActiveForm1.Activate();
                        goto finish;
                    }

                    try
                    {
                        info = iwtext.Text;
                        if (info == "reach limit"
                            || info.StartsWith("reach traffic limit.") || info.StartsWith("识别失败")
                            || info.StartsWith("ip address banned") || info.StartsWith("System is busy") || info.StartsWith("system is busy"))
                        {
                            trafficLimit = true; DialogResult ds = DialogResult.None;
                            StopOCR = true; ActiveForm1.PagePaste2GjcoolOCR_ing = false;
                            //Debugger.Break();
                            ActiveForm1.TopMost = false;
                            string targetProcessName = "VPN by Google One"; // 目標程序的名稱
                            IntPtr targetWindowHandle = FindWindow(null, targetProcessName);
                            Task tsRing = Task.Run(() =>
                            {
                                if (info.StartsWith("System is busy") || info.StartsWith("system is busy") || info.StartsWith("ip address banned"))
                                {
                                    Form1.OCRBreakSoundNotification();
                                    //if (File.Exists("C:\\Windows\\Media\\ring05.wav"))
                                    //{
                                    //    using (SoundPlayer sp = new SoundPlayer("C:\\Windows\\Media\\ring05.wav"))
                                    //    {
                                    //        sp.Play();
                                    //        Thread.Sleep(12000);
                                    //        if (File.Exists("C:\\Windows\\Media\\ring04.wav") && Control.ModifierKeys != forms.Keys.Control && ds == DialogResult.None)//若需中止，按下Ctrl鍵
                                    //        {
                                    //            sp.SoundLocation = "C:\\Windows\\Media\\ring04.wav";
                                    //            sp.Play();
                                    //            Thread.Sleep(3000);
                                    //        }
                                    //    }
                                    //}
                                }
                                if (File.Exists("C:\\Windows\\Media\\alarm05.wav") && Control.ModifierKeys != forms.Keys.Control && ds == DialogResult.None)
                                {
                                    using (SoundPlayer sp = new SoundPlayer("C:\\Windows\\Media\\alarm05.wav"))
                                    {
                                        sp.Play();
                                        Thread.Sleep(7000);
                                        if (File.Exists("C:\\Windows\\Media\\alarm10.wav") && Control.ModifierKeys != forms.Keys.Control && ds == DialogResult.None)//若需中止，按下Ctrl鍵
                                        {
                                            sp.SoundLocation = "C:\\Windows\\Media\\alarm10.wav";
                                            sp.Play();
                                            Thread.Sleep(3000);
                                        }
                                        if (File.Exists("C:\\Windows\\Media\\Windows Logon.wav") && Control.ModifierKeys != forms.Keys.Control && ds == DialogResult.None)
                                        {
                                            sp.SoundLocation = "C:\\Windows\\Media\\Windows Logon.wav";
                                            sp.Play();
                                            Thread.Sleep(4000);
                                        }
                                        if (File.Exists("C:\\Windows\\Media\\Windows Print complete.wav") && Control.ModifierKeys != forms.Keys.Control && ds == DialogResult.None)
                                        {
                                            sp.SoundLocation = "C:\\Windows\\Media\\Windows Print complete.wav";
                                            sp.Play();
                                            Thread.Sleep(2000);
                                        }
                                        if (File.Exists("C:\\Windows\\Media\\alarm01.wav") && Control.ModifierKeys != forms.Keys.Control && ds == DialogResult.None)
                                        {
                                            sp.SoundLocation = "C:\\Windows\\Media\\alarm01.wav";
                                            sp.Play();
                                            Thread.Sleep(5000);
                                        }

                                    }
                                }
                            });
                            //tsRing.Wait(6000);

                            try
                            {
                                if (info.StartsWith("System is busy") || info.StartsWith("system is busy") || info.StartsWith("ip address banned")) CurrentIP = GetPublicIpAddress(string.Empty);//CurrentIP = CurrentIP == string.Empty ? GetPublicIpAddress(string.Empty) : CurrentIP;
                            }
                            catch (Exception)
                            {
                                returnFalse = true;
                                goto finish;
                            }

                            string mark = info.StartsWith("识别失败") ? "●●●●●●●●●" :
                                (info.StartsWith("System is busy") || info.StartsWith("system is busy") || info.StartsWith("ip address banned")) ? "★★★★★★★★★★★" + CurrentIP + "★★★★" : "●";
                            if (info.StartsWith("ip address banned")) Clipboard.SetText(CurrentIP);
                            //現在有批量處理權限，此功能懸置，且VPN之IP多已遭《古籍酷》封鎖，故預設按鈕改為「取消」！20240831
                            ds = MessageBox.Show("是否讓程式自動更換IP？", "●切換IP？" + mark + "『" + info + "』" + mark
                                , MessageBoxButtons.OKCancel, MessageBoxIcon.None, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly); //Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否讓程式自動更換IP？", "●切換IP？")
                            if (DialogResult.OK == ds)
                            {//要自動切換IP時：
                             //driver.Close();//return以後也還會再執行一次哦！注意

                                Form1.playSound(Form1.soundLike.over); ActiveForm1.HideToNICo();// TopMost = false;

                                //Task ts = Task.Run(() =>
                                //{
                                //    IPSwitchOnly();//此方法在切換TouchVPN時會再開啟一分頁以檢視IP轉換情形
                                //    IPStatusMessageShow();
                                //});


                                //20231228  Bing大菩薩：C# 多執行緒存取剪貼簿：
                                Thread thread = new Thread(() =>
                                {
                                    IPSwitchOnly();
                                    //IPStatusMessageShow();
                                });
                                thread.SetApartmentState(ApartmentState.STA); // 設定執行緒為 STA 模式
                                thread.Start();
                                //thread.Join(); // 等待該執行緒完成
                                if (thread.IsAlive)
                                {
                                    // The thread is still running.
                                    thread.Join(1000);
                                }
                                else
                                {
                                    // The thread has completed.
                                }

                            }
                            else
                            {
                                CurrentIP = string.Empty;
                                if (targetWindowHandle != IntPtr.Zero)
                                {//如果有開啟 VPN by Google One                                
                                 //StopOCR = true;//前已有
                                    return TouchVPN_IvacyVPN_VeePN_ExtensionSwitcher();
                                }
                                else
                                {
                                    if (MessageBox.Show("是否要切換成批量處理模式？", "若按下【取消】，擬改用『標注平台』處理，請記得在textBox2下「gjk」指令以切換。感恩感恩　南無阿彌陀佛　讚美主",
                                        MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly) == DialogResult.OK)
                                        Form1.BatchProcessingGJcoolOCR = true;
                                    else
                                        Form1.BatchProcessingGJcoolOCR = false;
                                }
                            }
                            ActiveForm1.TopMost = false;
                            //不管要不要自動切換IP都執行
                            if (targetWindowHandle != IntPtr.Zero)
                            {//如果有開啟 VPN by Google One                                
                             //StopOCR = true;//前已有
                                if (ds == DialogResult.Cancel)
                                    Thread.Sleep(4500);//the seconds ref: internal static bool TouchVPN_IvacyVPN_ExtensionSwitcher()
                                return false;
                            }
                            else
                            {//如果沒開啟
                             //driver.Close();//呼叫端會關！
                                if (ds == DialogResult.OK)
                                    Thread.Sleep(4500);//the seconds ref: internal static bool TouchVPN_IvacyVPN_ExtensionSwitcher()
                                returnFalse = true;
                                goto finish;
                            }

                        }
                    }
                    catch (Exception)
                    {

                    }
                }
                #endregion

                if (clicked && Clipboard.GetText() != string.Empty)
                {
                    ActiveForm1.Activate();
                    goto finish;
                }

            }//以上待複製按鈕出現

            //找到複製按鈕以後

            //copyBtnPos = e.Location;

            //int x, y;
            //x = e.Location.X; y = e.Location.Y;
            ////Console.WriteLine(e.Location.ToString());
            ////Console.WriteLine(x);
            ////Console.WriteLine(y);

            //copyBtnPos.X = x + 76 + (e.Size.Width / 2);//76：Windows系統工具列在左邊時
            //copyBtnPos.Y = y + 120 + (e.Size.Height / 2);//120：最大化Chrome瀏覽器網頁頂端至瀏覽器頂端的高度

            ////Cursor.Position = copyBtnPos;

            driver.SwitchTo().Window(driver.CurrentWindowHandle);//切換到目前Selenium操控的視窗，就不怕沒及時得到焦點而失誤了

            //clickCopybutton_GjcoolFastExperience(copyBtnPos);

            if (Clipboard.GetText().IndexOf(Environment.NewLine) > -1) goto finish;

            try
            {
                e?.Click();
            }
            catch (Exception)
            {
                //throw;
            }
            Form1.playSound(Form1.soundLike.over, true);//找到複製按鈕按下的音效
            #endregion

            int timeSpanSecs = 25;//0;
            #region 原式
            /*
             * 上面 Bing大菩薩指導的 找到元件了，原來這個就不必要了，留作紀念。畢竟幫我們完成了許多書頁。感恩感恩　讚歎讚歎　南無阿彌陀佛

            //待手動成功複製，上限為 timeSpanSecs 秒
             Task task = null;
            try
            {

                //Clipboard.Clear();
                //按下複製按鈕複製到剪貼簿
                //SendKeys.Send("{tab 4}~");
                //SendKeys.Send("{tab 7}~");
                //SendKeys.Send("+{tab 10}");
                //Thread.Sleep(3000);
                //SendKeys.Send("+{tab 4}");
                //SendKeys.Send("~");

                //iwe.Click();
                //複製結果顯示到剪貼簿
                //Clipboard.SetText(iwe.Text);
                #endregion

                #region 關閉OCR視窗後回到原來分頁視窗-手動按下「複製」按鈕
                //！！！！此須手動按下「複製」按鈕了！！！！
                timeSpanSecs = 8;
                //滑鼠定位，以備手動按下「複製」按鈕（須視窗最大化）
                //copyBtnPos = new Point(838, 711);//用PRTSC鍵拍下全螢幕後，貼到小畫家以滑鼠取得坐標位置（即顯示在狀態列中）
                copyBtnPos = new Point(838, 721);//用PRTSC鍵拍下全螢幕後，貼到小畫家以滑鼠取得坐標位置（即顯示在狀態列中）

                Cursor.Position = copyBtnPos;
                //Thread.Sleep(800);//要等一下才行否則反應不過來
                //Form1.playSound(Form1.soundLike.info);
                if (ActiveForm1.TopMost) ActiveForm1.TopMost = false;
                Thread.Sleep(300);//要等一下才行否則反應不過來                
                / 20230401 Bing大菩薩：在C#中，您可以使用 `MouseOperations` 类来模拟鼠标点击。这个类中有一个名为 `MouseEvent` 的方法，它可以接受一个 `MouseEventFlags` 枚举值作为参数，用来指定要执行的鼠标操作¹。例如，要模拟鼠标左键点击，可以这样写：
                ```csharp
                    MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftDown);
                    MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftUp);
                    ```
                    來源: 與 Bing 的交談， 2023/4/1(1) .net - How do you simulate Mouse Click in C#? - Stack Overflow. https://stackoverflow.com/questions/2416748/how-do-you-simulate-mouse-click-in-c 已存取 2023/4/1.
                    (2) c# - Using SendMessage to simulate mouse clicks - Stack Overflow. https://stackoverflow.com/questions/14876345/using-sendmessage-to-simulate-mouse-clicks 已存取 2023/4/1.
                    (3) How to programatically trigger a mouse left click in C#?. https://stackoverflow.com/questions/2736965/how-to-programatically-trigger-a-mouse-left-click-in-c 已存取 2023/4/1.
                    (4) c# - I want to send mouse click with SendMessage but it's not working, What wrong with my code? - Stack Overflow. https://stackoverflow.com/questions/46306860/i-want-to-send-mouse-click-with-sendmessage-but-its-not-working-what-wrong-wit 已存取 2023/4/1.
                 /
                ////MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftDown);
                ////MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftUp);                
                //MouseOperations.MouseEventMousePos(MouseOperations.MouseEventFlags.LeftDown, copyBtnPos);
                ////Thread.Sleep(50);
                //MouseOperations.MouseEventMousePos(MouseOperations.MouseEventFlags.LeftUp, copyBtnPos);
                ////Form1.playSound(Form1.soundLike.info);
                //clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.none);
                task = Task.Run(() => { clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.press); });


                /Bing大菩薩：您好，`MouseOperations` 不是 C# 的内置类。它是一个自定义类，您可以在 Stack Overflow 上找到它的源代码。您可以将这些代码复制到您的项目中，然后使用它来模拟鼠标点击。
                 /

                //藉由手動關閉視窗以提早/強制中止程序
                try
                {
                    if (currentWindowHndl != driver.CurrentWindowHandle) { };
                }
                catch (Exception)
                {
                    return false;
                }

                task.Wait();
                Thread.Sleep(450);//待寫入剪貼簿

                if (Clipboard.GetText() != "") goto finish;


                task = Task.Run(async delegate
                {
                    await Task.Run(async () =>
                    {
                        if (Clipboard.GetText() == "")
                        {
                            //Thread.Sleep(850);
                            await Task.Run(async () => { await clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.none); });

                            Thread.Sleep(450);//等寫入剪貼簿

                            if (Clipboard.GetText() == "")
                            {
                                //Thread.Sleep(450);
                                await Task.Run(async () => { await clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.over); });
                                //Task tk2 = Task.Run(() => { clickCopybutton_GjcoolFastExperience(copyBtnPos); });
                                //tk2.Wait();
                                //Thread.Sleep(450);
                                Thread.Sleep(900);
                                if (Clipboard.GetText() == "")
                                {
                                    await Task.Run(async () => { await clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.none); });
                                    //Task tk3 = Task.Run(() => { clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.over); });
                                    //tk3.Wait();
                                    Thread.Sleep(450);
                                    if (Clipboard.GetText() == "")
                                    {
                                        Task tk4 = Task.Run(async () => { await clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.none); });
                                        //Task tk4 = Task.Run(() => { clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.done); });
                                        //tk4.Wait();
                                        Thread.Sleep(450);
                                    }
                                }
                            }
                        }

                    });
                });
            }
            catch (Exception ex)
            {
                if (ex.Message.IndexOf("剪貼簿") == -1)
                {
                    if (Clipboard.GetText().Length > 22)
                        Console.WriteLine(ex.HResult + ex.Message);
                }
                else
                    //剪貼簿失效忽略不計（可能是手動按下複製了；也就不必等待了）
                    switch (ex.HResult)
                    {
                        case -2147221040:
                            if (ex.Message.IndexOf("要求的剪貼簿作業失敗。") > -1)
                            {
                                //Clipboard.Clear();//Clipboard.Clear();//此時剪貼簿還不能用
                                //Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請手動操作複製按鈕！");
                                //Debugger.Break();
                                DateTime date = DateTime.Now;
                                while (!Form1.isClipBoardAvailable_Text(100)) { if (DateTime.Now.Subtract(date).Seconds > 0.5) break; }
                                //Thread.Sleep(500);
                            }
                            break;
                        default:
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message + "\n\r\n\r程式將忽略此錯誤，繼續進行。");
                            break;
                    }
                //throw;
            }

            #region 方便提早取消作業（藉由關閉OCR視窗）
            try
            {
                if (currentWindowHndl != driver.CurrentWindowHandle) { };
            }
            catch (Exception)
            {
                return false;
            }
            #endregion

            task.Wait();

            */

            #endregion

            #region 方便提早取消作業（藉由關閉OCR視窗）
            try
            {
                if (currentWindowHndl != driver.CurrentWindowHandle) { }
                ;
            }
            catch (Exception)
            {
                return false;
            }
            #endregion

            //如果 「Thread.Sleep(3900);」 太短，則再一次試試：
            if (Clipboard.GetText() == "")
            {
                Task ts = Task.Run(() =>
                {
                    //Thread.Sleep(1300);//要寫在這，讓_OCR_GJcool_WindowClosed能設定完成
                    Thread.Sleep(600);//要寫在這，讓_OCR_GJcool_WindowClosed能設定完成
                    if (Clipboard.GetText() == "" && !_OCR_GJcool_WindowClosed)
                        //if (!_OCR_GJcool_WindowClosed)
                        clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.press);

                    if (Clipboard.GetText() == "")
                    {
                        Task ts1 = Task.Run(() =>
                        {
                            Thread.Sleep(800);//要寫在這，讓_OCR_GJcool_WindowClosed能設定完成
                            if (Clipboard.GetText() == "" && !_OCR_GJcool_WindowClosed) clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.none);
                        });
                        ts1.Wait();
                        if (Clipboard.GetText() == "")
                        {
                            Task ts2 = Task.Run(() =>
                            {
                                //前已有ts1.Wait();或不再需要
                                Thread.Sleep(400);//要寫在這，讓_OCR_GJcool_WindowClosed能設定完成
                                if (Clipboard.GetText() == "" && !_OCR_GJcool_WindowClosed) clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.none);
                            });
                            ts2.Wait();
                            if (Clipboard.GetText() == "")
                            {
                                Task ts3 = Task.Run(() =>
                                {
                                    Thread.Sleep(800);//要寫在這，讓_OCR_GJcool_WindowClosed能設定完成
                                    if (Clipboard.GetText() == "" && !_OCR_GJcool_WindowClosed) clickCopybutton_GjcoolFastExperience(copyBtnPos, Form1.soundLike.done);
                                });
                                ts3.Wait();
                            }
                        }
                    }

                });
                ts.Wait();
            }
            else//剪貼簿已有資料
            {
                //_OCR_GJcool_WindowClosed = true;
                goto finish;
            }

            #region 方便提早取消作業（藉由關閉OCR視窗）
            try
            {
                if (currentWindowHndl != driver.CurrentWindowHandle) { }
                ;
            }
            catch (Exception)
            {
                return false;
            }
            #endregion

            if (Clipboard.GetText() != "") goto finish;


            #region 20250119 取消
            //while (!Form1.isClipBoardAvailable_Text(10))
            //{
            //    //Form1.playSound(Form1.soundLike.info);
            //    //if (timeSpanSecs > 0 && DateTime.Now.Subtract(begin).TotalSeconds > timeSpanSecs) return false;
            //    //藉由手動關閉視窗以提早/強制中止程序
            //    #region 方便提早取消作業（藉由關閉OCR視窗）
            //    try
            //    {
            //        if (currentWindowHndl != driver.CurrentWindowHandle) { };
            //    }
            //    catch (Exception)
            //    {
            //        return false;
            //    }
            //    #endregion
            //    //if (copyBtnPos.X > 0)//= Point(838, 711)
            //    //{
            //    //    MouseOperations.MouseEventMousePos(MouseOperations.MouseEventFlags.LeftDown, copyBtnPos);
            //    //    Thread.Sleep(50);
            //    //    MouseOperations.MouseEventMousePos(MouseOperations.MouseEventFlags.LeftUp, copyBtnPos);
            //    //    Thread.Sleep(100);
            //    //}
            //}
            #endregion

            //while (Clipboard.GetText().Length == 0)
            //{

            //    MouseOperations.MouseEventMousePos(MouseOperations.MouseEventFlags.LeftDown, copyBtnPos);
            //    MouseOperations.MouseEventMousePos(MouseOperations.MouseEventFlags.LeftUp, copyBtnPos);

            //    //藉由手動關閉視窗以提早/強制中止程序
            //    try
            //    {
            //        if (currentWindowHndl != driver.CurrentWindowHandle) { };
            //    }
            //    catch (Exception)
            //    {
            //        return false;
            //    }
            //    Thread.Sleep(450);
            //}

            if (Clipboard.GetText() != "")
            {
                _OCR_GJcool_WindowClosed = true;
                goto finish;
            }

            string txtchkClipboard = "";
            if (Clipboard.GetText() != "")
                txtchkClipboard = new StringInfo(Clipboard.GetText()).SubstringByTextElements(0);
            while ((txtchkClipboard == "" || txtchkClipboard.IndexOf("正在识别") > -1) && txtchkClipboard.Length < 22)
            {
                //每半秒按下滑鼠左鍵1次
                //Thread.Sleep(400);
                Thread.Sleep(200);
                MouseOperations.MouseEventMousePos(MouseOperations.MouseEventFlags.LeftDown, copyBtnPos);
                MouseOperations.MouseEventMousePos(MouseOperations.MouseEventFlags.LeftUp, copyBtnPos);
                Thread.Sleep(450);
                if (Clipboard.GetText() != "")
                    txtchkClipboard = new StringInfo(Clipboard.GetText()).SubstringByTextElements(0);
                if (txtchkClipboard != "")
                {
                    if (txtchkClipboard.Length > 22)
                    {
                        if (txtchkClipboard.IndexOf("正在识别") == -1)
                        {
                            Clipboard.SetText(txtchkClipboard);
                            break;
                        }
                    }
                }
                /*
                 20230330 Bing大菩薩：在C#中，與VBA中的Stop語句等效的是 `System.Diagnostics.Debugger.Break()`¹。這樣可以在程式執行到這一行時暫停並進入調試器，類似於設置斷點¹。
                    來源: 與 Bing 的交談， 2023 / 3 / 30(1) Can I do a Visual Basic(VB) Stop in C#?. https://social.msdn.microsoft.com/Forums/vstudio/en-US/db9dfe97-c98d-4f4b-bb8f-ba2edffee988/can-i-do-a-visual-basic-vb-stop-in-c?forum=csharpgeneral 已存取 2023/3/30.
                    (2) VBA) (Stop 語句 | Microsoft Learn.https://learn.microsoft.com/zh-tw/office/vba/language/reference/user-interface-help/stop-statement 已存取 2023/3/30.
                    (3) What is the equivalent of End(VB6/ VBA) in order to end in C# for Windows applications? - Stack Overflow. https://stackoverflow.com/questions/2033141/what-is-the-equivalent-of-end-vb6-vba-in-order-to-end-in-c-sharp-for-windows-a 已存取 2023/3/30.*/
                //System.Diagnostics.Debugger.Break();
                if (DateTime.Now.Subtract(begin).TotalSeconds > timeSpanSecs)
                {
                    if (DialogResult.Cancel == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("已過" + timeSpanSecs + "秒，是否再等5秒？"))
                        return false;
                    else
                    {
                        timeSpanSecs += 5;
                        driver.SwitchTo().Window(driver.CurrentWindowHandle);
                        driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div[1]/div[3]/div[2]/div[2]/button/i"))?.Click();
                        //BringToFront("chrome");
                        goto retry;
                    }
                }
                //if (Clipboard.GetText() == "")
                //    return false;
                //else
                //return true;
                //driver.Close();
                //driver.SwitchTo().Window(currentWindowHndl);

                //藉由手動關閉視窗以提早/強制中止程序
                try
                {
                    if (currentWindowHndl != driver.CurrentWindowHandle) { }
                    ;
                }
                catch (Exception)
                {
                    return false;
                }
            }
        //driver.Close();
        //driver.SwitchTo().Window(currentWindowHndl);
        #endregion

        #region 關閉OCR視窗後回到原來分頁視窗
        finish:
            StopOCR = true;
            if (!_OCR_GJcool_WindowClosed) _OCR_GJcool_WindowClosed = true;
            return !returnFalse;
            #endregion
        }
        /// <summary>
        /// 存放《古籍酷》快速體驗OCR完成時的複製按鈕位置資訊
        ///  835,711
        ///  835,730
        ///  706,840
        /// </summary>
        internal static Point Copybutton_GjcoolFastExperience_Location = new Point();

        /// <summary>
        /// 滑鼠左鍵點擊網頁上的複製按鈕並發出提示音
        /// </summary>
        /// <param name="copyBtnPos">要點擊的座標</param>
        /// <param name="soundlike">指定所發出的提示音</param>
        /// <returns>執行完畢即傳回一個Task物件以供await參考</returns>
        private static Task clickCopybutton_GjcoolFastExperience(Point copyBtnPos, Form1.soundLike soundlike = Form1.soundLike.info)
        {
            //Thread.Sleep(1300);
            if (Cursor.Position != copyBtnPos)
                Cursor.Position = copyBtnPos;
            //Form1.playSound(soundlike);
            //MouseOperations.MouseEventMousePos(MouseOperations.MouseEventFlags.LeftDown, copyBtnPos);
            //MouseOperations.MouseEventMousePos(MouseOperations.MouseEventFlags.LeftUp, copyBtnPos);

            return Task.Run(() =>
            {
                Form1.playSound(soundlike);
                MouseOperations.MouseEventMousePos(MouseOperations.MouseEventFlags.LeftDown, copyBtnPos);
                MouseOperations.MouseEventMousePos(MouseOperations.MouseEventFlags.LeftUp, copyBtnPos);
                //Thread.Sleep(150);//讓複製功能讀入剪貼簿。
                Task.WaitAll();
            });
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
        /// 20230311 YouChat菩薩，待測試！！！！
        /// </summary>
        /// <returns></returns>
        private static string getChromeDownloadDirectory_YouChat()
        {//用 C# 和 Selenium 可以取得Chrome瀏覽器的下載目錄嗎            
         //YouChat菩薩：可以使用 C# 和 Selenium 获取 Chrome 浏览器的下载目录。可以使用以下代码示例来实现：
            IWebElement downloadsFolderInput = driver.FindElement(By.Name("download.default_directory"));
            string downloadsFolder = downloadsFolderInput.GetAttribute("value");
            //在这个示例中，我们使用 OpenQA.Selenium 和 OpenQA.Selenium.Chrome 命名空间中提供的 ChromeDriver 类来创建一个 Chrome 浏览器实例。接下来，我们使用 driver.FindElement(By.Name("download.default_directory")) 方法查找 Chrome 浏览器下载目录的输入框，然后使用 downloadsFolderInput.GetAttribute("value") 获取输入框中的值，即下载目录的路径。
            //请注意，要在计算机上成功运行此代码，您必须安装 Selenium Webdriver 和 ChromeDriver，并将其与您的 C# 项目一起使用。
            return downloadsFolder;
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

                //20240911 Copilot大菩薩： 處理 Chrome 下載目錄的例外情形: https://sl.bing.net/kX4BWlGiCSi
                /*這個例外情形「並未將物件參考設定為物件的執行個體」通常是因為 preferencesJson["download"]["default_directory"] 的值為 null。這可能是因為 preferencesJson["download"] 或 preferencesJson["download"]["default_directory"] 不存在於 JSON 文件中。
                    你可以在存取這些值之前，先檢查它們是否存在。以下是修改後的程式碼：
                 */
                if (preferencesJson["download"] != null && preferencesJson["download"]["default_directory"] != null)
                {
                    downloadDirectory_Chrome = preferencesJson["download"]["default_directory"].ToString();
                }
                else
                {
                    // 處理找不到下載目錄的情況
                    //Console.WriteLine("無法找到下載目錄設定。");
                    if (MessageBox.Show("無法找到下載目錄設定。是否要設定為作業系統下載目錄？", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly) == DialogResult.OK)
                        downloadDirectory_Chrome = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads";
                    else
                        MessageBox.Show("請　賢友菩薩在textBox2中輸入下載目錄之路徑。設定正確，其文字即會消失。感恩感恩　讚歎讚歎　南無阿彌陀佛", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                }
                //downloadDirectory_Chrome = preferencesJson["download"]["default_directory"].ToString();
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

        /// <summary>
        /// 焦點必須在textBox1！！20240313
        /// 依選取文字取得目前URL加該選取字為該頁之關鍵字的連結。如欲在此頁中標出「𢔶」字，即為：
        /// https://ctext.org/library.pl?if=gb&file=36575&page=53#𢔶
        /// Ctrl + k
        /// </summary>
        /// <returns></returns>
        internal static string GetPageUrlKeywordLink(string w, string url, bool reMovePunctuations = false)
        {
            //if (!ActiveForm1.Controls["textBox1"].Focused) return string.Empty;
            //TextBox tb = ActiveForm1.Controls["textBox1"] as TextBox;
            //if (tb.SelectionLength == 0) return string.Empty;            
            if (url == null) return string.Empty;
            int i = url.IndexOf("&page=");
            if (i == -1) return string.Empty;

            i = url.IndexOf("&", i + "&page=".Length + 1);
            if (i > -1) //20240102 Bard大菩薩：C# 找到字串中「=53」的結束位置
                url = url.Substring(0, i);
            else
            {
                i = url.IndexOf("&page=") + "&page=".Length + 1;
                // 從起始位置開始，逐個字元比較，直到找到非數字或字串結束
                int end = i;
                while (end < url.Length && char.IsDigit(url[end]))
                {
                    end++;
                }
                url = url.Substring(0, end);
            }
            //Clipboard.SetText(w);
            //return url + "#" + HttpUtility.UrlEncode(w) ;//VBA中文編碼好像還是有問題，先用這個，並先複製一個字進剪貼簿，可以利用 Win + v 的方式檢視調用
            //以上VBA bug 已排除
            w = w.Replace(Environment.NewLine, string.Empty);
            return url + "#" + (reMovePunctuations ? CnText.RemovePunctuationsNum(w) : w);//到VBA再轉碼，以便複製此字、不必再key也。況昨晚才經Bing大菩薩、StackOverflow AI大菩薩的加持，得以成功建置此生第1個 dll檔案，供Word VBA調用。感恩感恩　讚歎讚歎　南無阿彌陀佛
        }

        /// <summary>
        /// Alt + F10 ： 將textBox1中選取的文字送去《古籍酷》自動標點。若無選取則將整個textBox1的內容送去。20240808（臺灣父親節）
        /// </summary>
        /// <param name="x">要送去《古籍酷》自動標點的文本內容</param>
        /// <returns>有誤則傳回false</returns>
        internal static bool GjcoolPunct(ref string x)
        {
            if (driver == null) return false;
            if (x == string.Empty) return false;
            //20240809 Copilot大菩薩：C# Windows.Forms 中判定文本是否只有英數字：這個正則表達式 ^[a-zA-Z0-9\s.,!?']+$ 會允許英文字母、數字、半形空格以及常見的標點符號（如句號、逗號、驚嘆號、問號和單引號）。
            if (Regex.IsMatch(x, @"^[a-zA-Z0-9\s.,!?']+$")) return false;

            //LastValidWindow = GetCurrentWindowHandle(driver)??(IsWindowHandleValid(driver,driver.WindowHandles.Last())? driver.WindowHandles.Last(): driver.WindowHandles.First());
            string lastWindowHandle = driver.WindowHandles.Last();
            if (!IsWindowHandleValid(driver, lastWindowHandle))
                Debugger.Break();
            LastValidWindow = GetCurrentWindowHandle(driver) ?? lastWindowHandle;
            openNewTabWindow();
            int retryCntr = 0;
        retry:
            try
            {
                driver.Navigate().GoToUrl("https://gj.cool/punct");
            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case -2146233088:
                        if (ex.Message.StartsWith("timeout: Timed out receiving message from renderer:"))//timeout: Timed out receiving message from renderer: -0.014
                                                                                                         //(Session info: chrome = 129.0.6668.59)
                        {
                            retryCntr++;
                            if (retryCntr > 5)
                            {
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("無法開啟《古籍酷》自動標點網頁，請檢查後再重試。感恩感恩　南無阿彌陀佛");
                                return false;
                            }
                            goto retry;
                        }
                        break;
                    default:
                        break;
                }
            }
            IWebElement iwe = WaitFindWebElementBySelector_ToBeClickable("#PunctArea");
            DateTime dt = DateTime.Now;
            while (iwe == null)
            {
                iwe = WaitFindWebElementBySelector_ToBeClickable("#PunctArea");
                if (DateTime.Now.Subtract(dt).TotalSeconds > 10) if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("等待頁面開啟已逾時，是否繼續？") == DialogResult.Cancel) return false;
            }
            SetIWebElement_textContent_Property(iwe, x);
            //rePaste:20240914作廢
            ////將要標點的文本寫入剪貼簿：
            //try
            //{
            //    Clipboard.SetText(x);
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    if (Clipboard.GetText() == string.Empty)
            //    {
            //        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("剪貼簿出錯！請重來"); return false;
            //    }
            //}
            //catch (Exception)
            //{
            //    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("剪貼簿出錯！請重來");
            //    return false;
            //}
            ////將要標點的文本貼到標點區：
            //iwe.SendKeys(OpenQA.Selenium.Keys.Shift + OpenQA.Selenium.Keys.Insert);
            //if (string.Empty == iwe.Text)
            //{
            //    if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("貼上失誤，是否重試一次？") == DialogResult.OK) goto rePaste;
            //}


            iwe = null;
            //按下「標點」按鈕：
            Thread.Sleep(640);//非得要等一會才能成功！
            dt = DateTime.Now;
            while (iwe == null)
            {
                iwe = WaitFindWebElementBySelector_ToBeClickable("#main > div > div.p-1.p-md-3.d-flex.justify-content-end > div:nth-child(6) > button > i");
                if (DateTime.Now.Subtract(dt).TotalSeconds > 8) if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("還沒找到「標點」按鈕，是否繼續？") == DialogResult.Cancel) return false;
            }
            iwe.Click();//Thread.Sleep(640);//非得要等一會才能成功！//自動標點會清除全形空格
                        //iwe.SendKeys(OpenQA.Selenium.Keys.Enter);//不能互動，會出現錯誤
            iwe = WaitFindWebElementBySelector_ToBeClickable("#PunctArea");
            dt = DateTime.Now; bool reClickFlag = false;
            //等待OCR結果
            //while (iwe.Text == x)//.Text屬性傳回的會是經過trim的，故若開頭是全形空格，則一下子就會誤判成已經標點過（文本經改過）的了
            while (iwe.GetAttribute("textContent") == x)//.Text屬性傳回的會是經過trim的，故若開頭是全形空格，則一下子就會誤判成已經標點過（文本經改過）的了
            {
                //檢查如果沒有按到「標點」按鈕，就再次按下 20240811 以出現等待圖示控制項為判斷
                if (WaitFindWebElementBySelector_ToBeClickable("#waitingSpinner") == null && reClickFlag == false && iwe.Text == x)
                {
                    WaitFindWebElementBySelector_ToBeClickable("#main > div > div.p-1.p-md-3.d-flex.justify-content-end > div:nth-child(6) > button > i").Click();
                    Form1.playSound(Form1.soundLike.over, true);
                    reClickFlag = true;
                }
                //reach traffic limit. wait 1.2 hours
                if (WaitFindWebElementBySelector_ToBeClickable("#main > div:nth-child(1)") != null)
                    if (WaitFindWebElementBySelector_ToBeClickable("#main > div:nth-child(1)").GetAttribute("textContent").StartsWith("reach traffic limit. wait "))
                    { Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("超過額度！"); return false; }

                if (DateTime.Now.Subtract(dt).TotalSeconds > 25) if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("標點逾時，是否繼續？") == DialogResult.Cancel) return false;
            }
            //x = iwe.Text;//.Text屬性傳回的會是經過trim的
            x = iwe.GetAttribute("textContent");
            //補上句號
            if (x.EndsWith("<p>") || x.EndsWith("<p>" + Environment.NewLine))
            {
                string endChar;
                if (x.EndsWith("<p>"))
                    endChar = x.Substring(x.Length - 4, 1);
                else
                    endChar = x.Substring(x.Length - 6, 1);
                if (char.IsLowSurrogate(endChar.ToCharArray()[0]) ||
                        Form1.IsChineseString(endChar))
                {
                    if (x.EndsWith("<p>"))
                        x = x.Substring(0, x.Length - 3) + "。" + "<p>";
                    else
                        x = x.Substring(0, x.Length - 5) + "。" + "<p>" + Environment.NewLine;
                }
            }
            //關閉https://gj.cool/punct頁面回到原來的頁面
            driver.Close();
            driver.SwitchTo().Window(LastValidWindow);
            return true;
        }
        /// <summary>
        /// Ctrl + F10： 將textBox1中選取的文字送去《古籍酷》舊版自動標點。若無選取則將整個textBox1的內容送去。（小於20字元不處理）20240808（臺灣父親節）
        /// </summary>
        /// <param name="x">要送去《古籍酷》舊版自動標點的文本內容</param>
        /// <returns>有誤則傳回false</returns>
        internal static bool GjcoolPunctOld(ref string x)
        {
            if (driver == null) return false;
            try
            {
                //LastValidWindow = GetCurrentWindowHandle(driver);
                LastValidWindow = driver.CurrentWindowHandle;
            }
            catch (Exception ex)
            {
                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message, "GjcoolPunctOld記下原來分頁視窗句柄");
            }
            openNewTabWindow();
            driver.Navigate().GoToUrl("https://old.gj.cool/gjcool/index");
            //文本輸入框
            IWebElement iwe = WaitFindWebElementBySelector_ToBeClickable("#origin000");
            DateTime dt = DateTime.Now;
            while (iwe == null)
            {
                iwe = WaitFindWebElementBySelector_ToBeClickable("#origin000");
                if (DateTime.Now.Subtract(dt).TotalSeconds > 10) if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("等待頁面開啟已逾時，是否繼續？") == DialogResult.Cancel) return false;
            }
            SetIWebElementValueProperty(iwe, x);
            //rePaste:
            ////將要標點的文本寫入剪貼簿：20240914作廢
            //try
            //{
            //    Clipboard.SetText(x);
            //}
            //catch (Exception)
            //{
            //}
            //try
            //{
            //    if (Clipboard.GetText() == string.Empty)
            //    {
            //        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("剪貼簿出錯！請重來"); return false;
            //    }
            //}
            //catch (Exception)
            //{
            //    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("剪貼簿出錯！請重來");
            //    return false;
            //}
            ////將要標點的文本貼到標點區：
            //iwe.SendKeys(OpenQA.Selenium.Keys.Shift + OpenQA.Selenium.Keys.Insert);
            //if (string.Empty == iwe.GetAttribute("value"))
            //{
            //    if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("貼上失誤，是否重試一次？") == DialogResult.OK) goto rePaste;
            //}


            iwe = null;
            //按下「標點」按鈕：
            Thread.Sleep(640);//非得要等一會才能成功！
            dt = DateTime.Now;
            while (iwe == null)
            {
                iwe = WaitFindWebElementBySelector_ToBeClickable("#processbtn");
                if (DateTime.Now.Subtract(dt).TotalSeconds > 8) if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("還沒找到「標點」按鈕，是否繼續？") == DialogResult.Cancel) return false;
            }
            iwe.Click();//Thread.Sleep(640);//非得要等一會才能成功！
                        //iwe.SendKeys(OpenQA.Selenium.Keys.Enter);//不能互動，會出現錯誤
                        //標點結果文本框
            while (null == WaitFindWebElementBySelector_ToBeClickable("#result001")) { }
            iwe = WaitFindWebElementBySelector_ToBeClickable("#result001");
            dt = DateTime.Now;
            while (iwe.GetAttribute("textContent") == x || iwe.GetAttribute("textContent") == string.Empty)
            {

                if (DateTime.Now.Subtract(dt).TotalSeconds > 25) if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("標點逾時，是否繼續？") == DialogResult.Cancel) return false;
            }
            x = iwe.GetAttribute("textContent");
            standardizedText(ref x);

            string standardizedText(ref string text)
            {
                //x = x.Substring(0, " ".Length) == " " ? x.Substring(" ".Length, x.Length - " ".Length) : x;
                text = text.Substring(0, " ".Length) == " " ? text.Substring(" ".Length) : text;
                text = text.Replace("　", string.Empty).Replace(" ", Environment.NewLine);
                return text;
            }

            //關閉https://gj.cool/punct頁面回到原來的頁面
            driver.Close();
            driver.SwitchTo().Window(LastValidWindow);
            return true;
        }
        /// <summary>
        /// 檢測是否可以回到上一頁
        /// 20240926 creedit_with_Copilot大菩薩：C# Selenium 檢查 Chrome 瀏覽器是否可以回到上一頁：https://sl.bing.net/ioeF5IoZq1I
        /// </summary>
        /// <returns>若可回到上一頁則為true</returns>
        internal static bool CanGoBack()
        {
            // 獲取當前 URL
            string currentUrl = driver.Url;

            // 嘗試回到上一頁
            driver.Navigate().Back();
            //System.Threading.Thread.Sleep(1000); // 等待頁面加載

            // 檢查 URL 是否改變
            if (driver.Url != currentUrl)
            {
                return true;
                //Console.WriteLine("可以回到上一頁");
            }
            else
            {
                //Console.WriteLine("無法回到上一頁");
                return false;
            }
        }
        /// <summary>
        /// 改變CTP圖文對照網址的 Page 參數以供翻頁
        /// 20240920 Copilot大菩薩：更改 URL 参数以翻页：https://sl.bing.net/jZV8afaj85Q
        /// </summary>
        /// <param name="url">要改變的網址</param>
        /// <param name="newPageNumber">Page參數要成的數值</param>
        /// <returns>傳回改動後的網址</returns>
        public static string ChangePageParameter(string url, int newPageNumber)
        {
            var uri = new Uri(url);
            var query = System.Web.HttpUtility.ParseQueryString(uri.Query);
            query.Set("page", newPageNumber.ToString());
            var uriBuilder = new UriBuilder(uri)
            {
                Query = query.ToString()
            };
            return uriBuilder.ToString();
        }
        /// <summary>
        /// 作為「《漢籍全文資料庫》檢索易學關鍵字」的欲檢索關鍵字瀏覽清單索引（已檢索瀏覽之位置）記錄用。在textBox2中輸入「lx」（list index clear(x=叉=清除））以歸零
        /// </summary>
        internal static int ListIndex_Hanchi_SearchingKeywordsYijing = 0;

        /// <summary>
        /// 在《漢籍全文資料庫》及《中國哲學書電子化計劃》檢索易學關鍵字
        /// 【進階檢索】中指定書名請自行輸入
        /// 在textBox2中輸入「lx」重設《漢籍全文資料庫》檢索易學關鍵字清單之索引值為0 即 ListIndex_Hanchi_SearchingKeywordsYijing=0。 
        /// textBox3.Text 會顯示關鍵字清單的索引值（從0開始）        
        /// <returns>檢索有結果、或關鍵字清單找完一遍了、或失敗則傳回true（以供呼叫端判斷是否停止繼續呼叫） 若檢索成功但沒結果，則傳回false </returns>
        /// </summary>
        internal static bool Hanchi_CTP_SearchingKeywordsYijing()
        {
            if (driver == null) return true;

            string title = null;

            if (WindowHandles.TryGetValue("Hanchi_CTP_SearchingKeywordsYijing", out string windowHandle_Hanchi_CTP_SearchingKeywordsYijing))
                if (driver.WindowHandles.Contains(windowHandle_Hanchi_CTP_SearchingKeywordsYijing))
                    if (!IsDriverInvalid())
                    {
                        if (driver.CurrentWindowHandle != windowHandle_Hanchi_CTP_SearchingKeywordsYijing)
                            driver.SwitchTo().Window(windowHandle_Hanchi_CTP_SearchingKeywordsYijing);
                    }

            #region 檢查title creedit_with_Copilot大菩薩： C# 比對字串中的特定子字串： https://sl.bing.net/dy0lfJJXE72
            //檢查title 如果不包含指定的條件，則傳回false
            bool checkTitle(string titleToCheck)
            {
                string pattern = "漢籍全文|中國哲學書電子化計劃|Chinese Text Project";
                return Regex.IsMatch(titleToCheck, pattern);
                //string[] keywords = { "漢籍全文", "中國哲學書電子化計劃", "Chinese Text Project" };
                //var foundKeywords = keywords.Where(keyword => title.Contains(keyword));
            }

            void forEachTabChectTitle()
            {
                //《漢籍全文資料庫》網頁介面
                foreach (var item in driver.WindowHandles)
                {
                    driver.SwitchTo().Window(item);
                    try
                    {
                        if (driver.Title.Contains("漢籍全文")) break;
                    }
                    catch (Exception exx)
                    {
                        switch (exx.HResult)
                        {
                            case -2146233088:
                                if (exx.Message.StartsWith("unknown error\nfrom no such execution context:"))//unknown error
                                                                                                             //from no such execution context: frame does not have execution context
                                                                                                             //  (Session info: chrome = 128.0.6613.138)
                                    continue;
                                break;
                            default:
                                Console.WriteLine(exx.HResult + exx.Message);
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(exx.HResult + exx.Message);
                                break;
                        }
                    }
                }
                if (!driver.Title.Contains("漢籍全文"))
                {
                    title = string.Empty;
                    foreach (var item in driver.WindowHandles)
                    {
                        try
                        {
                            if (driver.SwitchTo().Window(item).Url.StartsWith("https://ctext.org/wiki.pl?if="))//https://ctext.org/wiki.pl?if=gb&res=、https://ctext.org/wiki.pl?if=en&res=
                            { title = driver.Title; break; }
                        }
                        catch (Exception exx)
                        {
                            switch (exx.HResult)
                            {
                                case -2146233088:
                                    if (exx.Message.StartsWith("unknown error\nfrom no such execution context:"))//unknown error
                                                                                                                 //from no such execution context: frame does not have execution context
                                                                                                                 //  (Session info: chrome = 128.0.6613.138)
                                        continue;
                                    break;
                                default:
                                    Console.WriteLine(exx.HResult + exx.Message);
                                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(exx.HResult + exx.Message);
                                    break;
                            }
                        }

                    }
                }

                if (!WindowHandles.TryGetValue("Hanchi_CTP_SearchingKeywordsYijing", out windowHandle_Hanchi_CTP_SearchingKeywordsYijing))
                    WindowHandles.Add("Hanchi_CTP_SearchingKeywordsYijing", driver.CurrentWindowHandle);
                else
                    if (windowHandle_Hanchi_CTP_SearchingKeywordsYijing != driver.CurrentWindowHandle)
                    WindowHandles["Hanchi_CTP_SearchingKeywordsYijing"] = windowHandle_Hanchi_CTP_SearchingKeywordsYijing;
            }

            #endregion



            try
            {
                //不要切到Chrome瀏覽器！可以在等結果的時候做別的工作。20240915
                //driver.SwitchTo().Window(driver.CurrentWindowHandle);
                title = driver.Title;//避免誤關出錯
                if (!checkTitle(title)) forEachTabChectTitle();
                if (title.IsNullOrEmpty())
                {
                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請開啟《漢籍全文資料庫》網頁檢索介面，再開始操作");
                    return true;
                }
            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case -2146233088:
                        if (ex.Message.StartsWith("no such window: target window already closed"))
                            Browser.NoSuchWindowErrHandler();
                        else if (ex.Message.IndexOf("from no such execution context: frame does not have execution context") > -1)
                        {
                            if (IsWindowHandleValid(driver, LastValidWindow))
                                driver.SwitchTo().Window(LastValidWindow);
                            else
                                driver.SwitchTo().Window(driver.WindowHandles.Last());
                            goto default;
                        }
                        break;
                    default:
                        forEachTabChectTitle();
                        if (title.IsNullOrEmpty())
                        {
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請開啟《漢籍全文資料庫》網頁檢索介面，再開始操作");
                            return true;
                        }
                        break;
                }
                title = driver.Title;
            }

            List<string> keywords = new List<string> { "易", "五經", "六經","七經", "十三經","蓍", "爻", "繫詞", "繫辭", "文言", "乾坤","元亨","利貞", "咎",
                 "夬", "頤","巽","顨","坎","兌","小畜","大畜","歸妹","明夷","明𡗝","同人于宗","同人","大有","豫大","豫","蠱","噬嗑","臨卦","〈需〉","〈屯","蒙〉","𫎇〉","賁於外","外賁","內賁","賁","剝","大過","小過","遯世無悶","遯","大壯","睽","暌","蹇","姤","萃","艮其背", "乾知大始","乾以易知","坤作成物","坤以簡能" ,"艮","渙","中孚","既濟","未濟","乾卦","坤卦","訟卦","師卦","比卦","履卦","泰卦","否卦","謙卦","隨卦","觀卦","復卦","習坎","離卦","咸卦","恆卦","晉卦","家人卦","解卦","損卦","益卦","升卦","困卦","井卦","革卦","鼎卦","震卦","漸卦","豐卦","旅卦","節卦","卦",
                "咸恆","老陰", "老陽", "少陰", "少陽","十翼","四象","兩儀","大衍",
                "无妄", "彖", "象曰", "象傳", "象日", "象云","大象","小象","象文", "筮", // 不支援標點檢索，如「, "象："」
                "初九","九二","九三","九四","九五","上九","初六","六二","六三","六四","六五","上六","用九","用六", "繇辭","繇詞","往吝","往遴",
                "伏羲","庖羲","庖𦏁","宓𦏁","宓羲","宓犧","伏犧","庖犧","中正","見龍在田","括囊","噬膚","舊井","井谷","為麗","離麗","天行健",
                "隨時之義","來復","庖有魚","包有魚","精義入神","豶豕","童牛","承之羞","雷在天上","錫馬", "蕃庶","晝日","三接","懲忿","窒欲","窒慾","敬以直內","義以方外","迷後得主","利西南","品物咸章","天下大行","益動而", "日進無疆","頻巽","豚魚","頻復", "懲窒","閑邪","存誠","乾乾","悔吝","憧憧", "類萬物","柔順利貞","比之匪人","貞厲","履貞","履道坦坦","貞吉","貞凶","悔亡","時義","健順", "內健而外順", "內健外順", "外順而內健", "外順內健","敦復","直方","開物成務","窮神知化", "夕惕","惕若","研幾極深","極深研幾","研幾","硏幾","一陰一陽","允升","木上有水","勞民勸相","索而得","我有好爵","言有序","有聖人之道四","長子帥師","弟子輿尸","無悶","日用而不知", "日用不知","之道鮮","原始反終", "然不動", "感而遂通","朋從", "朋盍", "容民畜眾","有過則改","見善則遷","養正","養賢","知臨","臨大君", "默而成之","黙而成之","不言而信", "存乎德行","通天下之志","履正", "繼之者善", "仁者見之", "知者見之", "智者見之","屯其膏", "貞不字","翰音","鶾音", "善不積","立成器", "與地之","變化云為","神道設教",
                "象義","大貞","小貞", "帝出乎震","帝出於震","帝出于震","敬義", "日新","與時偕行","盈虛","山澤通氣","豐亨","居業","確乎", "碻乎","天在山中", "多識前言往行", "蹇蹇", "匪躬","洗心","龍德","慎言語","節飲食","艮其限","乃孚","幹父","裕父","係遯","甘臨","號咷", "風行水上",
                "終難","輝光","咸之九五","賁於丘園","賁于丘園","賁於邱園","立人之道","賁于邱園", "束帛","戔戔", "損下以益上", "其腓","菑畬","葘畬", "損下益上", "損下而益上", "貳用缶","納約自牖","利見大人", "何思何慮","同歸而殊塗","一致而百慮", "同歸殊塗","一致百慮","先天後天","改命吉","天下雷行","喪貝","羝羊","羝芉", "觸藩", "觸籓","事不密","艱貞","金矢","利有","攸往","包蒙", "童蒙", "蒙吉","不事王侯","高尚其事",
                "精氣為物","游魂為變","遊䰟為變","游䰟為變", "不家食","漣如","威如","焚如","知幾","禴祭", "東鄰","朋亡", "渙其群","有子考","甲三日","庚三日","不易乎世","不成乎名","天一地二","者其辭","升其高陵","天道虧盈","鞏用", "祗悔", "祇悔","秖悔","秪悔","履霜","蒞眾","理財", "正辭", "禁民為非","撝謙", "浚恒","浚恆", "立其誠","立誠","修辭立誠","開國承家",
                "寬以居之","確乎其不可拔","碻乎其不可拔"
                };

            //異體字處理（只用在《中國哲學書電子化計劃》，因為《漢籍全文資料庫》已俱。）
            if (title.EndsWith("中國哲學書電子化計劃") || title.EndsWith("Chinese Text Project"))
            {
                //新增List元素。因為《中國哲學書電子化計劃》異體字的支援機制沒有《漢籍全文資料庫》那麼好
                //20240719 Copilot大菩薩：C# Windows.Forms List 新增多個元素：您好，如果您想要在程式進行中對 List<string> 新增多個元素，可以使用 AddRange 方法。這是一個範例：
                //keywords.Add();
                List<string> additionalKeywords = new List<string> { "无𡚶", "𧰼", "系辭", "擊詞", "擊辭", "繫驟",
                    "乹","〈乾〉", "〈坤〉", "〈乾坤〉", "咸恒","剥","頥","㢲","旣濟","涣","兑","兊","大壮",
                    "〈泰〉","〈否〉","〈損〉","〈益〉","〈屯〉","〈豫〉","〈旡妄〉","〈復〉","〈震〉",
                    "少隂","太隂","𥘉九","𭃨九","𭃡九","𥘉六","𭃨六","𭃡六","徃吝","徃遴",
                    "索而𢔶","离麗","旧井","𦾔井","知𭙌","知㡬","天行徤","通乎晝夜","不事王矦","高尙其事","髙尙其事","髙尚其事",
                "悔亾","悔兦","无悶","遯世无悶","容民畜衆","盈虚","盈𮓡","敬𮊫","盈虗","匪躳","愼言語","賁於𠀉園", "賁于𠀉園","賁於𠀌園", "賁于𠀌園","賁於𨚑園", "賁于𨚑園", "𩔖萬物", "𩔗萬物","東隣殺牛","禴𥙊","禴𫞴","涣其群","渙其羣","涣其羣","攺命吉","撝謙","事不宻","脩辭立誠",
                "有子攷","来復","敬以直内","不易乎卋","不易乎丗","升其髙陵","蒞衆","莅眾","莅衆","大𧰼","䘮貝","𭈬貝","𠷔貝","丧貝","𠸶貝","𡂤貝","包𫎇", "童𫎇", "𫎇吉", "神道設敎",
                "寛以居之","宽以居之","𡩖以居之",
                "伏𦏁"};
                keywords.AddRange(additionalKeywords);
            }

            if (ListIndex_Hanchi_SearchingKeywordsYijing < 0) ListIndex_Hanchi_SearchingKeywordsYijing = 0;
            if (ListIndex_Hanchi_SearchingKeywordsYijing > keywords.Count - 1) ListIndex_Hanchi_SearchingKeywordsYijing = keywords.Count - 1;
            string keyword = keywords[ListIndex_Hanchi_SearchingKeywordsYijing]; //Clipboard.SetText(keyword);20240914作廢
            ActiveForm1.PauseEvents();
            string indexStr = ListIndex_Hanchi_SearchingKeywordsYijing.ToString();
            ActiveForm1.textBox4Text = indexStr;
            if (indexStr.Length > 2) if (ActiveForm1.textBox4Font.Size > 12)
                { ActiveForm1.textBox4Font = new Font(ActiveForm1.textBox4Font.FontFamily, 12); ActiveForm1.Refresh(); }
            ActiveForm1.Controls["textBox1"].Text = keyword;
            ActiveForm1.ResumeEvents();
            bool returnValue = false;

            //if (title.Contains("中國哲學書電子化計劃")) 即CTP
            if (title.EndsWith("中國哲學書電子化計劃") || title.EndsWith("Chinese Text Project"))
            {
            //檢索方塊
            researchCtext:
                //桌面版和手機版的尋找方塊不同
                //手機版
                IWebElement iwe = WaitFindWebElementBySelector_ToBeClickable("#deskimg");
                if (iwe != null)
                    iwe = WaitFindWebElementBySelector_ToBeClickable("#content > div.wikibox > table > tbody > tr.mobilesearch > td > form > input[type=text]:nth-child(3)");
                //桌面版
                else if (WaitFindWebElementBySelector_ToBeClickable("#mobimg") != null)
                    iwe = WaitFindWebElementBySelector_ToBeClickable("#searchform > input.searchbox");
                else
                    Debugger.Break();
                if (iwe != null)
                {
                    //輸入檢索條件
                    SetIWebElementValueProperty(iwe, keyword);
                    //以下作廢 20240913
                    //iwe.Clear();
                    //try
                    //{
                    //    Clipboard.SetText(keyword);
                    //}
                    //catch (Exception)
                    //{

                    //}
                    //輸入檢索條件

                    //try
                    //{
                    //    iwe.SendKeys(keyword);
                    //}
                    //catch (Exception ex)
                    //{
                    //    switch (ex.HResult)
                    //    {
                    //        case -2146233088:
                    //            if (ex.Message.StartsWith("unknown error: ChromeDriver only supports characters in the BMP"))
                    //            {
                    //                //iwe.SendKeys(OpenQA.Selenium.Keys.Shift + OpenQA.Selenium.Keys.Insert);
                    //                ChromeDriverOnlySupportsCharactersBMP(iwe, keyword);
                    //            }
                    //            else
                    //                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ":" + ex.Message);
                    //            break;
                    //        default:
                    //            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ":" + ex.Message);
                    //            break;
                    //    }
                    //}
                    iwe.SendKeys(OpenQA.Selenium.Keys.Enter);

                    //Total 0
                    iwe = WaitFindWebElementBySelector_ToBeClickable("#content > table.searchsummary > tbody > tr:nth-child(4) > th > b", 5);
                    if (iwe != null)
                    {
                        if (iwe.GetAttribute("textContent") == "Total 0")
                        {
                            driver.Navigate().Back();
                            returnValue = false;
                        }
                        else//檢索有結果-中國哲學書電子化計劃
                        {
                            returnValue = true; Clipboard.SetText(keyword); ActiveForm1.KeyinTextmodeSwitcher();
                            //有文本的文字框
                            //Thread.Sleep(800);
                            ActiveForm1.TopMost = false;
                            driver.SwitchTo().Window(driver.CurrentWindowHandle);
                            if (new StringInfo(keyword).LengthInTextElements > 1)
                                FindPageAndPaste2Find(driver, keyword);
                            else
                                HighlightKeywords(keyword);
                        }//ActiveForm1.HideToNICo(); }

                    }
                    else//檢索有結果-中國哲學書電子化計劃
                    {
                        returnValue = true; Clipboard.SetText(keyword); ActiveForm1.KeyinTextmodeSwitcher();
                        //有文本的文字框
                        //Thread.Sleep(800);
                        ActiveForm1.TopMost = false;
                        driver.SwitchTo().Window(driver.CurrentWindowHandle);
                        if (new StringInfo(keyword).LengthInTextElements > 1)
                            FindPageAndPaste2Find(driver, keyword);
                        else
                            HighlightKeywords(keyword);
                    }//ActiveForm1.HideToNICo(); }
                }
                else
                {
                    string url = driver.Url;
                    // 一直回到有「檢索方塊」的上一頁
                    while (null == WaitFindWebElementBySelector_ToBeClickable("#content > div.wikibox > table > tbody > tr.mobilesearch > td > form > input[type=text]:nth-child(3)"))
                    {
                        driver.Navigate().Back();
                        #region 送出後檢查是否是「Please confirm that you are human! 敬請輸入認證圖案」頁面 網址列：https://ctext.org/wiki.pl
                        if (IsConfirmHumanPage())
                        {
                            //點選輸入框
                            OpenQA.Selenium.IWebElement iweConfirm = WaitFindWebElementBySelector_ToBeClickable("#content3 > form > table > tbody > tr:nth-child(2) > td:nth-child(2) > input[type=text]");
                            if (iweConfirm == null) driver.Navigate().Back();//因非同步，若已翻到下一頁
                            iweConfirm = WaitFindWebElementBySelector_ToBeClickable("#content3 > form > table > tbody > tr:nth-child(2) > td:nth-child(2) > input[type=text]");
                            if (iweConfirm == null)
                                Debugger.Break();
                            else
                                iweConfirm.Click();
                            if (DialogResult.Cancel ==
                                Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("Please confirm that you are human! 敬請輸入認證圖案"
                                + Environment.NewLine + Environment.NewLine + "請輸入完畢後再按「確定」！"))
                            {
                                Debugger.Break();
                                Form1.InstanceForm1.EndUpdate();
                            }
                            driver.Navigate().Back();
                            while (driver.Url == "https://ctext.org/wiki.pl")
                            {
                                driver.Navigate().Back();
                            }
                        }
                        #endregion //送出後檢查是否是「Please confirm that you are human! 敬請輸入認證圖案」頁面 網址列：https://ctext.org/wiki.pl
                        if (url == driver.Url) //沒有上一頁了
                        {
                            returnValue = true;
                            break;
                        }
                        url = driver.Url;
                    }
                    if (null != WaitFindWebElementBySelector_ToBeClickable("#content > div.wikibox > table > tbody > tr.mobilesearch > td > form > input[type=text]:nth-child(3)"))
                        goto researchCtext;
                    ////Total 標籤
                    //if (waitFindWebElementBySelector_ToBeClickable("#content > table.searchsummary > tbody > tr:nth-child(4) > th > b", 5)?.GetAttribute("textContent").StartsWith("Total ") == true)
                    //{
                    //    driver.Navigate().Back();
                    //    goto researchCtext;
                    //}
                    //else
                    //{
                    //    driver.Navigate().Back();
                    //    returnValue = false;
                    //}
                }
            }
            else
            {//如果檢索《漢籍全文資料庫》

                //【檢索報表】標籤控制項(關閉開啟的分頁）
                while (null != WaitFindWebElementBySelector_ToBeClickable("body > form > table > tbody > tr:nth-child(2) > td:nth-child(1) > font > b > nobr"))
                { driver.Close(); driver.SwitchTo().Window(driver.WindowHandles.Last()); }


                string caption;//= string.Empty;// iwe1 == null ? "漢籍全文資料庫" : "漢籍全文文本閱讀";
                               //文本閱讀中的查詢輸入方塊 <input type="text" name="hanji/fld00.33.810" size="30" maxlength="200">
                IWebElement iwe1 = WaitFindWebElementBySelector_ToBeClickable("body > form > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td.leftbg > table > tbody > tr:nth-child(1) > td > table > tbody > tr:nth-child(2) > td > input[type=text]:nth-child(2)");
                if (iwe1 == null)
                {
                    //標題橫幅(中含有「回漢籍」）
                    iwe1 = WaitFindWebElementBySelector_ToBeClickable("body > form > table > tbody > tr:nth-child(1) > td");
                    if (iwe1 != null)
                    {
                        if (iwe1.GetAttribute("innerText").Contains("回漢籍"))
                        {
                            caption = "漢籍全文文本閱讀"; string url = string.Empty;
                            //文本閱讀中的查詢輸入方塊
                            while (null == WaitFindWebElementBySelector_ToBeClickable("body > form > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td.leftbg > table > tbody > tr:nth-child(1) > td > table > tbody > tr:nth-child(2) > td > input[type=text]:nth-child(2)"))
                            {
                                if (url == driver.Url) break;
                                url = driver.Url;
                                try
                                {
                                    driver.Navigate().Back();
                                }
                                catch (Exception)
                                {
                                }
                            }
                            //文本閱讀中的查詢輸入方塊
                            iwe1 = WaitFindWebElementBySelector_ToBeClickable("body > form > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td.leftbg > table > tbody > tr:nth-child(1) > td > table > tbody > tr:nth-child(2) > td > input[type=text]:nth-child(2)");
                        }
                        else
                            caption = "漢籍全文資料庫";
                    }
                    else
                        caption = "漢籍全文資料庫";

                }
                else
                    caption = "漢籍全文文本閱讀";

                if (!title.Contains(caption))
                {
                    //《漢籍全文資料庫》網頁介面
                    foreach (var item in driver.WindowHandles)
                    {
                        driver.SwitchTo().Window(item);
                        if (driver.Title.Contains(caption)) break;
                    }
                    if (!driver.Title.Contains(caption))
                    {
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請開啟《漢籍全文資料庫》網頁檢索介面，再開始操作");
                        return true;
                    }
                }

                if (caption == "漢籍全文資料庫")
                {

                    //輸入「任意詞」：KeywordInputBox
                    IWebElement iweKeywordInputBox = WaitFindWebElementBySelector_ToBeClickable("#frmTitle > table > tbody > tr:nth-child(2) > td > table > tbody > tr:nth-child(1) > td > table > tbody > tr > td > table > tbody > tr:nth-child(1) > td:nth-child(2) > nobr > input[type=TEXT]");
                    if (iweKeywordInputBox == null)
                    {
                        if (iwe1 == null)
                        {
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請開啟《漢籍全文資料庫》網頁檢索介面，再開始操作");
                            return true;
                        }

                    }

                    SetIWebElementValueProperty(iweKeywordInputBox, keyword);
                    //iweKeywordInputBox.Clear();
                    //try
                    //{
                    //    iweKeywordInputBox.SendKeys(keyword);
                    //}
                    //catch (Exception ex)
                    //{
                    //    switch (ex.HResult)
                    //    {
                    //        case -2146233088:
                    //            if (ex.Message.StartsWith("unknown error: ChromeDriver only supports characters in the BMP"))
                    //            {
                    //                ChromeDriverOnlySupportsCharactersBMP(iweKeywordInputBox, keyword);
                    //            }
                    //            else
                    //                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ":" + ex.Message);
                    //            break;
                    //        default:
                    //            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ":" + ex.Message);
                    //            break;
                    //    }
                    //}

                    //按下「搜尋」：
                    //IWebElement iwe = waitFindWebElementBySelector_ToBeClickable("#frmTitle > table > tbody > tr:nth-child(2) > td > table > tbody > tr:nth-child(8) > td > input[type=IMAGE]:nth-child(1)");
                    IWebElement iwe = waitFindWebElementByName_ToBeClickable("_IMG_搜尋", 3);
                    if (iwe != null && iweKeywordInputBox.GetAttribute("value") == keyword)
                    {
                        try
                        {
                            //iwe.Submit();
                            //iweKeywordInputBox.SendKeys(OpenQA.Selenium.Keys.Enter);
                            iwe.Click();
                        }
                        catch (Exception ex)
                        {
                            switch (ex.HResult)
                            {
                                case -2146233088:
                                    if (ex.Message.StartsWith("The HTTP request to the remote WebDriver server for URL") && ex.Message.EndsWith("seconds."))//The HTTP request to the remote WebDriver server for URL http://localhost:4778/session/50acc6066cba7f783fa177f82af56a91/element/f.B3F971C5B82A6FD17EB852AEB6512BFF.d.E5371FFD241AD1B605772DC930B35479.e.5894/click timed out after 30.5 seconds.
                                        if (DialogResult.Cancel == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("等候逾時，是否繼續等"))
                                            return true;
                                    break;
                                default:
                                    Console.WriteLine(ex.HResult + ex.Message);
                                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                                    break;
                            }
                            WaitFindWebElementBySelector_ToBeClickable("#frmTitle > table > tbody > tr:nth-child(2) > td > table > tbody > tr:nth-child(1) > td > table > tbody > tr > td > table > tbody > tr:nth-child(1) > td:nth-child(2) > nobr > input[type=TEXT]")?.SendKeys(OpenQA.Selenium.Keys.Enter);
                            //throw;
                        }

                        iwe = WaitFindWebElementBySelector_ToBeClickable("body > form > table > tbody > tr:nth-child(2) > td:nth-child(3) > center > table > tbody > tr:nth-child(2) > td > font");
                        if (iwe != null)
                        {
                            //if (iwe.GetAttribute("innerText") == "　抱歉，找不到您所查詢的資料")
                            //{
                            //    //Form1.playSound(Form1.soundLike.error, true);
                            //    //ActiveForm1.AvailableInUseBothKeysMouse();
                            //}
                        }
                        //找到了-《漢籍全文資料庫》
                        else
                        {//檢索有結果：《漢籍全文資料庫》
                            Form1.playSound(Form1.soundLike.info);
                            returnValue = true;
                            iwe = waitFindWebElementByName_ToBeClickable("_IMG_檢索報表", 2);
                            if (iwe != null)//   ?.Click();
                            {//20240710 Copilot大菩薩：要在 Selenium 中使用鍵盤的 Shift 鍵，您可以使用 Actions 類別來模擬鍵盤和滑鼠的操作。以下是一個範例程式碼：
                             // 建立一個新的 Actions 物件                                
                                Actions action = new Actions(driver);
                                // 按下 Shift 鍵，然後點擊元素，最後釋放 Shift 鍵
                                action.KeyDown(OpenQA.Selenium.Keys.Shift).Click(iwe).KeyUp(OpenQA.Selenium.Keys.Shift).Build().Perform();
                                //ActiveForm1.TopMost = false;//最後會有
                                //driver.SwitchTo().Window(driver.WindowHandles.Last());
                                //Browser.BringToFront("chrome");//最後會有

                                ActiveForm1.Activate(); ActiveForm1.TopMost = false;
                                //ActiveForm1.AvailableInUseBothKeysMouse();
                                //ActiveForm1.TopMost = false;

                                Thread.Sleep(300);//等待開新視窗開啟（實測290還不行，最快也要300微秒）20240726
                                driver.SwitchTo().Window(driver.WindowHandles.Last());
                                //「附註開啟」:「註」開展控制項
                                iwe = WaitFindWebElementBySelector_ToBeClickable("body > form > table > tbody > tr:nth-child(2) > td:nth-child(2) > input[type=IMAGE]:nth-child(1)", 5);
                                if (iwe != null)
                                {
                                    if (iwe.GetAttribute("title") == "附註開啟")
                                    //按下「附註開啟」以展開小註
                                    {
                                        iwe.Click();
                                        Form1.playSound(Form1.soundLike.done, true);
                                    }

                                }
                                FindPageAndPaste2Find(driver, keyword);

                                //else
                                //{//檢查「_IMG_附註關閉」是否已出現：
                                //    if (null == waitFindWebElementBySelector_ToBeClickable("body > form > table > tbody > tr:nth-child(2) > td:nth-child(2) > input[type=IMAGE]:nth-child(1)")) 
                                //        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請檢查附註是否已開啟！");
                                //}
                                //else
                                //{
                                //    try
                                //    {
                                //        //iwe = driver.FindElement(By.Name("_IMG_附註開啟"));
                                //        //iwe = driver.FindElement(By.Name("附註開啟"));
                                //        iwe = waitFindWebElementByName_ToBeClickable("_IMG_附註開啟", 2);
                                //    }
                                //    catch (Exception ex)
                                //    {
                                //        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                                //    }
                                //    if (iwe != null)
                                //    {
                                //        if (iwe.GetAttribute("title") == "附註開啟")
                                //        //按下「附註開啟」以展開小註
                                //        {
                                //            iwe.Click();
                                //            Form1.playSound(Form1.soundLike.done);
                                //        }
                                //    }
                                //    else
                                //    {
                                //        try
                                //        {
                                //            iwe = driver.FindElement(By.XPath("/html/body/form/table/tbody/tr[2]/td[2]/input[1]"));
                                //        }
                                //        catch (Exception ex)
                                //        {
                                //            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                                //        }
                                //        if (iwe != null)
                                //        {
                                //            if (iwe.GetAttribute("title") == "附註開啟")
                                //            //按下「附註開啟」以展開小註
                                //            {
                                //                iwe.Click();
                                //                Form1.playSound(Form1.soundLike.done);
                                //            }
                                //        }

                                //    }
                                //}

                            }
                        }

                    }
                }
                else//文本閱讀內的檢索（《漢籍全文資料庫》）
                {

                    //reClick:
                    //輸入查詢關鍵字
                    SetIWebElementValueProperty(iwe1, keyword);
                    iwe1.SendKeys(OpenQA.Selenium.Keys.Enter);
                    //iwe1.Clear();

                    //try
                    //{
                    //    iwe1.SendKeys(keyword);
                    //}
                    //catch (Exception ex)
                    //{
                    //    switch (ex.HResult)
                    //    {
                    //        case -2146233088:
                    //            if (ex.Message.StartsWith("unknown error: ChromeDriver only supports characters in the BMP"))
                    //            {
                    //                ChromeDriverOnlySupportsCharactersBMP(iwe1, keyword);
                    //            }
                    //            else
                    //                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ":" + ex.Message);
                    //            break;
                    //        default:
                    //            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ":" + ex.Message);
                    //            break;
                    //    }
                    //}


                    ////按下「查詢」按鈕
                    //iwe1 = waitFindWebElementBySelector_ToBeClickable("body > form > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td.leftbg > table > tbody > tr:nth-child(1) > td > table > tbody > tr:nth-child(2) > td > input.s_btn.hjblock");
                    //while (iwe1 == null)
                    //{
                    //    iwe1 = waitFindWebElementBySelector_ToBeClickable("body > form > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td.leftbg > table > tbody > tr:nth-child(1) > td > table > tbody > tr:nth-child(2) > td > input.s_btn.hjblock", 0.3);
                    //}
                    ////Task.Run(() => { iwe1.Click(); });
                    ////20240714 Copilot大菩薩：Selenium 網頁操作中的等待問題
                    //開始檢索
                    //try
                    //{
                    //    driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(2); // 設定頁面載入超時時間為10秒
                    //                                                                   //var element = driver.FindElement(By.CssSelector("your_css_selector"));
                    //                                                                   //element.Click();
                    //    try
                    //    {
                    //        iwe1.Click();
                    //    }
                    //    catch (Exception ex)
                    //    {
                    //        switch (ex.HResult)
                    //        {
                    //            case -2146233088:
                    //                if (ex.Message.StartsWith("element click intercepted: Element "))//element click intercepted: Element <input type="submit" name="_BTN_查詢^^^TQ,hanji,hanji_1st,hanji" value="查詢" alt="查詢" title="查詢" class="s_btn hjblock"> is not clickable at point (104, 183). Other element would receive the click: <div class="blockUI blockOverlay" style="z-index: 1000; border: none; margin: 0px; padding: 0px; width: 100%; height: 100%; top: 0px; left: 0px; background-color: rgb(0, 0, 0); opacity: 0.6; cursor: wait; position: fixed;"></div>
                    //                                                                                 //(Session info: chrome = 129.0.6668.59)
                    //                {
                    //                    if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("「查詢」按鈕現在無法按下，請將頁面調整好，或取消這次的操作。是否繼續？感恩感恩　南無阿彌陀佛") == DialogResult.Cancel)
                    //                        return true;
                    //                    else
                    //                        goto reClick;
                    //                }
                    //                else if (ex.Message.StartsWith("timeout: Timed out receiving message from renderer: "))//-2146233088timeout: Timed out receiving message from renderer: 2.000
                    //                                                                                                       //(Session info: chrome = 129.0.6668.59)
                    //                {
                    //                    if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("「查詢」按鈕現在無法按下，請將頁面調整好，或取消這次的操作。是否繼續？感恩感恩　南無阿彌陀佛") == DialogResult.Cancel)
                    //                        return true;
                    //                    else
                    //                        goto reClick;
                    //                }
                    //                else
                    //                {
                    //                    Console.WriteLine(ex.HResult + ex.Message);
                    //                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                    //                }
                    //                break;
                    //            default:
                    //                break;
                    //        }
                    //    }
                    //    //Form1.playSound(Form1.soundLike.info);
                    //    ActiveForm1.TopMost = false;
                    //}
                    //catch (WebDriverTimeoutException)
                    //{
                    //    // 處理頁面未在指定時間內載入完成的情況
                    //    // 在這裡不進行任何操作
                    //    //Form1.playSound(Form1.soundLike.error);
                    //}


                    // 繼續執行後續的程式碼


                    //查詢結果      -文本閱讀內的檢索（《漢籍全文資料庫》
                    //查詢結果編號「1」
                    //iwe1 = waitFindWebElementBySelector_ToBeClickable("body > form > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td > table > tbody > tr:nth-child(2) > td > table > tbody > tr:nth-child(1) > td.seqno > a");
                    iwe1 = WaitFindWebElementBySelector_ToBeClickable("body > form > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td.leftbg > table > tbody > tr:nth-child(1) > td > table > tbody > tr:nth-child(2) > td > div > span");
                    ////if (iwe1 != null)
                    ////Console.WriteLine(iwe1.GetAttribute("name"));
                    while (iwe1 == null)
                    {
                        if (WaitFindWebElementBySelector_ToBeClickable("body > form > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td > table > tbody > tr:nth-child(2) > td > table > tbody > tr:nth-child(1) > td.seqno", 0.3) != null)
                        {
                            iwe1 = WaitFindWebElementBySelector_ToBeClickable("body > form > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td > table > tbody > tr:nth-child(2) > td > table > tbody > tr:nth-child(1) > td.seqno");
                            returnValue = true;
                            Form1.playSound(Form1.soundLike.info, true);//靜音模式時仍播出
                            break;
                        }
                        iwe1 = WaitFindWebElementBySelector_ToBeClickable("body > form > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td.leftbg > table > tbody > tr:nth-child(1) > td > table > tbody > tr:nth-child(2) > td > input.s_btn.hjblock", 0.3);
                    }

                    //if (iwe1 != null && waitFindWebElementBySelector_ToBeClickable("body > form > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td > table > tbody > tr:nth-child(2) > td > table > tbody > tr:nth-child(1) > td.seqno", 0.3) == null)
                    //{
                    //    if (iwe1.GetAttribute("innerText") == "找不到您的檢索詞")
                    //    //while (waitFindWebElementBySelector_ToBeClickable("body > form > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td.leftbg > table > tbody > tr:nth-child(1) > td > table > tbody > tr:nth-child(2) > td > div > span")?.GetAttribute("innerText") == "找不到您的檢索詞"&& waitFindWebElementBySelector_ToBeClickable("body > form > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td > table > tbody > tr:nth-child(2) > td > table > tbody > tr:nth-child(1) > td.seqno", 0.3) == null)
                    //    {

                    //    }
                    //}

                    if (iwe1 == null)
                    {
                        //Form1.playSound(Form1.soundLike.error, true);
                        //ActiveForm1.AvailableInUseBothKeysMouse();

                    }
                    else
                    {
                        if (returnValue)//文本閱讀內的檢索有結果（《漢籍全文資料庫》）
                        {
                            ActiveForm1.TopMost = false;
                            driver.SwitchTo().Window(driver.CurrentWindowHandle);
                            Clipboard.SetText(keyword);
                            //BringToFront("chrome");
                            //FindPageAndPaste2Find(driver, keyword);//因為是一行行條件式的，所以比較不需要找頁面中關鍵字所在 20240914

                        }
                    }
                }
            }

            //因各頁面實際情況不同，故分開寫
            //if (returnValue)
            //{
            //    ActiveForm1.TopMost = false;
            //    driver.SwitchTo().Window(driver.CurrentWindowHandle);
            //    BringToFront("chrome");
            //}

            #region 單個關鍵字查詢結束
            ListIndex_Hanchi_SearchingKeywordsYijing++;
            if (ListIndex_Hanchi_SearchingKeywordsYijing > keywords.Count - 1)
            {
                ListIndex_Hanchi_SearchingKeywordsYijing = 0;
                Form1.playSound(Form1.soundLike.finish, true);//靜音模式時仍播出
                returnValue = true;
                if (ActiveForm1.textBox4Font.Size < 20.25)
                    ActiveForm1.textBox4Font = new Font(ActiveForm1.textBox4Font.FontFamily, (float)20.25);
                if (ActiveForm1.KeyinTextMode == false) ActiveForm1.KeyinTextmodeSwitcher(false);
            }
            #endregion//單個關鍵字查詢結束

            else if (returnValue && ActiveForm1.KeyinTextMode == false) ActiveForm1.KeyinTextmodeSwitcher();//20240912

            ////（還似不行！故還原）前已由Copilot大菩薩提供 Timeouts 方法及相關類別解決了。感恩感恩　讚歎讚歎　Copilot大菩薩　南無阿彌陀佛
            //if (!returnValue && caption == "漢籍全文文本閱讀")//因為網頁完全開啟會等很久
            //{
            //    driver.SwitchTo().Window(driver.CurrentWindowHandle);
            //    Browser.BringToFront("chrome");
            //    SendKeys.SendWait("{esc}");
            //    //// 建立一個新的 Actions 物件//這個還是要等前面的完成才會執行！
            //    //Actions action1 = new Actions(driver);
            //    //action1.KeyDown(OpenQA.Selenium.Keys.Escape).Perform();
            //    ////action.KeyDown(OpenQA.Selenium.Keys.Escape).Build().Perform();
            //    //SendKeys.Send("{esc}");
            //}

            //檢索結束
            return returnValue;
        }
        /// <summary>
        /// 在頁面按下Ctrl + f 後貼上剪貼簿內容來尋找
        /// </summary>
        /// <param name="driver"></param>
        /// <returns></returns>
        public static bool FindPageAndPaste2Find(WebDriver driver, string pasteTxt, bool paste2Find = true)
        {
            Form1.playSound(Form1.soundLike.over, true);
            ActiveForm1.TopMost = false;
            try
            {
                //driver.SwitchTo().Window(GetCurrentWindowHandle(driver));
                driver.SwitchTo().Window(driver.CurrentWindowHandle);
            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case 2146233088:
                        if (ex.Message.StartsWith("no such window: target window already closed"))
                            NoSuchWindowErrHandler();
                        else
                            goto default;
                        break;
                    default:
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                        break;
                }

            }
            if (ActiveForm1.Active) BringToFront("chrome");
            if (ActiveForm1.Active) Debugger.Break();
            SendKeys.SendWait("^f");//打開Chrome瀏覽器網頁上的「搜尋本頁內容」方塊
            if (paste2Find)
            {
                try
                {
                    Clipboard.SetText(pasteTxt);
                }
                catch (Exception)
                {
                }
                // 等待尋找方塊出現
                System.Threading.Thread.Sleep(200);
                // 貼上剪貼簿內容
                SendKeys.SendWait("+{insert}");
            }
            return true;
        }
        /// <summary>
        /// 20240731 Copilot大菩薩：C# Windows.Forms Selenium 螢光色標記：您可以使用C# Windows.Forms和Selenium在Chrome瀏覽器中將指定的關鍵字標上螢光色。以下是一個簡單的範例，展示如何使用JavaScript來實現這個功能：
        /// </summary>
        /// <param name="keyword"></param>
        /// <returns></returns>
        public static bool HighlightKeywords_(string keyword)
        {
            string script = $@"
                var keyword = '{keyword}';
                var regex = new RegExp(keyword, 'gi');
                document.body.innerHTML = document.body.innerHTML.replace(regex, function(matched) {{
                    return '<span style=""background-color: yellow;"">' + matched + '</span>';
                }});";//這段程式碼會在瀏覽器中將指定的關鍵字標上黃色的螢光色。您可以根據需要修改HighlightKeywords方法中的CSS樣式來改變螢光色的顏色。

            driver.ExecuteScript(script);
            return true;
        }
        /// <summary>
        /// 只能標識網頁中的一個中文（尤其是表格中的）標上螢光色
        /// 20240731 Copilot大菩薩：C# Windows.Forms Selenium 螢光色標記：您可以使用C# Windows.Forms和Selenium在Chrome瀏覽器中將指定的關鍵字標上螢光色。以下是一個簡單的範例，展示如何使用JavaScript來實現這個功能：
        /// </summary>
        /// <param name="keyword"></param>
        /// <returns></returns>
        public static bool HighlightKeywords(string keyword)
        {//這段程式碼已經考慮到處理超過100個關鍵字的情況。它會一次處理100個節點，但會在所有節點處理完之前繼續進行。也就是說，如果頁面上有超過100個關鍵字，它會分批次處理，直到所有關鍵字都被標記。
            string script = $@"
                var keyword = '{keyword}';
                var regex = new RegExp(keyword, 'gi');
                var walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
                var node;
                var nodes = [];
                while (node = walker.nextNode()) {{
                    nodes.push(node);
                }}
                function highlightNodes(start, end) {{
                    for (var i = start; i < end; i++) {{
                        var node = nodes[i];
                        var match = node.nodeValue.match(regex);
                        if (match) {{
                            var span = document.createElement('span');
                            span.style.backgroundColor = 'yellow';
                            span.textContent = match[0];
                            var after = node.splitText(match.index);
                            after.nodeValue = after.nodeValue.substring(match[0].length);
                            node.parentNode.insertBefore(span, after);
                        }}
                    }}
                }}
                var batchSize = 100;
                for (var i = 0; i < nodes.length; i += batchSize) {{
                    highlightNodes(i, Math.min(i + batchSize, nodes.length));
                }}
            ";

            driver.ExecuteScript(script);
            return true;
        }

        /// <summary>
        /// 20240731 Copilot大菩薩：C# Windows.Forms Selenium 螢光色標記
        /// </summary>
        /// <param name="keyword"></param>
        /// <returns></returns>
        public static bool HighlightKeywords_SingleCharacter(string keyword)
        {//如果您希望進一步優化，可以考慮使用setTimeout來分批次處理，這樣可以避免一次性處理過多節點導致的性能問題。以下是改進的版本：
         //這段程式碼使用setTimeout來分批次處理節點，每次處理100個節點，並在每批次之間等待50毫秒。這樣可以避免一次性處理過多節點導致的性能問題，同時確保所有關鍵字都能被標記。
            string script = $@"
        var keyword = '{keyword}';
        var regex = new RegExp(keyword, 'gi');
        var walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
        var node;
        var nodes = [];
        while (node = walker.nextNode()) {{
            nodes.push(node);
        }}
        function highlightNodes(start, end) {{
            for (var i = start; i < end; i++) {{
                var node = nodes[i];
                var match;
                while ((match = regex.exec(node.nodeValue)) !== null) {{
                    var span = document.createElement('span');
                    span.style.backgroundColor = 'yellow';
                    span.textContent = match[0];
                    var after = node.splitText(match.index);
                    after.nodeValue = after.nodeValue.substring(match[0].length);
                    node.parentNode.insertBefore(span, after);
                    node = after;
                }}
            }}
            if (end < nodes.length) {{
                setTimeout(function() {{
                    highlightNodes(end, Math.min(end + batchSize, nodes.length));
                }}, 50);
            }}
        }}
        var batchSize = 100;
        highlightNodes(0, batchSize);
    ";

            driver.ExecuteScript(script);
            return true;
        }
        public static bool HighlightKeywords_SingleCharacter1(string keyword)
        {
            string script = $@"
                var keyword = '{keyword}';
                var regex = new RegExp(keyword, 'g');
                var walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
                var node;
                var nodes = [];
                while (node = walker.nextNode()) {{
                    nodes.push(node);
                }}
                function highlightNodes(start, end) {{
                    for (var i = start; i < end; i++) {{
                        var node = nodes[i];
                        var match;
                        var matches = [];
                        while ((match = regex.exec(node.nodeValue)) !== null) {{
                            matches.push(match);
                        }}
                        if (matches.length > 0) {{
                            var parentNode = node.parentNode;
                            var newNode = document.createDocumentFragment();
                            var lastIndex = 0;
                            for (var j = 0; j < matches.length; j++) {{
                                match = matches[j];
                                newNode.appendChild(document.createTextNode(node.nodeValue.substring(lastIndex, match.index)));
                                var span = document.createElement('span');
                                span.style.backgroundColor = 'yellow';
                                span.textContent = match[0];
                                newNode.appendChild(span);
                                lastIndex = match.index + match[0].length;
                            }}
                            newNode.appendChild(document.createTextNode(node.nodeValue.substring(lastIndex)));
                            parentNode.replaceChild(newNode, node);
                        }}
                    }}
                    if (end < nodes.length) {{
                        setTimeout(function() {{
                            highlightNodes(end, Math.min(end + batchSize, nodes.length));
                        }}, 50);
                    }}
                }}
                var batchSize = 100;
                highlightNodes(0, batchSize);
                ";

            driver.ExecuteScript(script);
            return true;
        }

        /// <summary>
        /// 置換Url中的Box 為Editor 如 https://ctext.org/library.pl?if=gb&file=34873&page=78&editwiki=164323#editor
        /// https://ctext.org/library.pl?if=gb&file=34873&page=78&editwiki=164323#box(280,86,1,0) 20241101
        /// 與 FixUrl＿ImageTextComparisonPage 可互參考
        /// </summary>
        /// <param name="url"></param>
        /// <returns>傳回清除後的結果</returns>
        internal static string ReplaceUrl_Box2Editor(string url)
        {
            if (!url.StartsWith("http")) return url;
            int s = url.IndexOf("#box"); string xClear;// = null;
            if (s > -1)
            {
                xClear = url.Substring(s, url.IndexOf(")", s) - s + 1);
                url = url.Substring(0, s) + url.Substring(s + xClear.Length, url.Length - (s + xClear.Length))
                    + (url.IndexOf("#editor") == -1 ? "#editor" : string.Empty);
            }
            return url;
        }
        /// <summary>
        /// 清除Url中的雜項，如 #box(280,86,1,0)等（etc） 20241101 20250126
        /// </summary>
        /// <param name="url"></param>
        /// <returns>傳回清除後的結果</returns>
        internal static string ClearUrl_BoxEtc(string url)
        {
            if (!url.StartsWith("http")) return url;
            int s = url.IndexOf("#box");
            if (s > -1)
            {
                string xClear = url.Substring(s, url.IndexOf(")", s) - s + 1);
                url = url.Substring(0, s) + url.Substring(s + xClear.Length, url.Length - (s + xClear.Length));
            }
            return url;
        }

        /// <summary>
        /// 作為一些需要保留或比對驗證的視窗句柄集，鍵值是視窗ID（或唯一名稱）以供比較尋找，值為視窗句柄
        /// 也可作為其他必要的參考、參數，如鍵值為"currentPageNum"時，乃翻到下一頁前記下的前一頁的頁碼
        /// </summary>
        internal static Dictionary<string, string> WindowHandles = new Dictionary<string, string>();
        /// <summary>
        /// Alt + Shift + a ： [AI太炎](https://t.shenshen.wiki/)標點 20241105
        /// </summary>
        /// <param name="x">要標點的文本變數，標點結果亦儲存在此</param>
        /// <returns>成功傳回true</returns>
        public static bool AITShenShenWikiPunct(ref string x)
        {
            //限500字,大於10字
            StringInfo xInfo = new StringInfo(x); const string url = "https://t.shenshen.wiki/llm";
            if (xInfo.LengthInTextElements > 500 || xInfo.LengthInTextElements < 11)
            {
                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("限500字，或須多於10字");
                return false;
            }

            if (!IsDriverInvalid()) LastValidWindow = driver.CurrentWindowHandle;
            else LastValidWindow = driver.WindowHandles.Last();

            //string windowHandle= WindowHandles.FirstOrDefault(w => w.Value == "《AI太炎》").Key;            
            //if (WindowHandles.Count > 0)
            if (WindowHandles.TryGetValue("《AI太炎》", out string windowHandle))//20241224 Copilot大菩薩 ：是的，Dictionary 物件提供了一個非常快速的方法來根據鍵值 (Key) 取得對應的值 (Value)。在 C# 中，你可以使用字典的索引器來輕鬆地從字典中取得值。這裡有一個範例：……在這個範例中，TryGetValue 方法會嘗試根據指定的鍵來取得值。如果找到了對應的值，TryGetValue 方法會回傳 true 並將值存儲在 out 參數中；如果沒有找到，則回傳 false。 這種方法不僅簡潔，而且效能非常好，特別是對於大型字典。
            {
                if (driver.WindowHandles.Contains(windowHandle))//20241224 Copilot大菩薩：在 C# 中，Dictionary 物件沒有直接內建的方法可以根據值來查找鍵，但你可以使用 LINQ 來達到這個目的。以下是一個簡單的範例，展示如何使用 LINQ 根據值來查找鍵：
                /* namespace DictionaryValueSearch { class Program { static void Main(string[] args) { Dictionary<int, string> myDictionary = new Dictionary<int, string>() { { 1, "Apple" }, { 2, "Banana" }, { 3, "Cherry" } }; string valueToFind = "Banana"; var key = myDictionary.FirstOrDefault(x => x.Value == valueToFind).Key; if (!key.Equals(default(int))) { Console.WriteLine("Key: " + key); } else { Console.WriteLine("Value not found"); } } } }
                 * 在這個範例中，FirstOrDefault 方法會遍歷 Dictionary，並根據指定的值來尋找鍵。如果找到了匹配的值，它會回傳對應的鍵；如果沒有找到，則會回傳 default 值。在這裡，我們假設鍵是 int 型別。
                 */
                {
                    driver.SwitchTo().Window(windowHandle);
                    //if (driver.Url != "https://t.shenshen.wiki/")
                    if (driver.Url != url)
                    {
                        openNewTabWindow();
                        //driver.Navigate().GoToUrl("https://t.shenshen.wiki/");
                        driver.Navigate().GoToUrl(url);
                    }
                }
            }
            else
            {
                if (IsDriverInvalid())
                {
                    if (driver == null)
                        DriverNew();
                    else
                        driver.SwitchTo().Window(driver.WindowHandles.Last());
                    openNewTabWindow();
                    driver.Navigate().GoToUrl(url);
                }
                else
                {
                    driver.SwitchTo().Window(driver.WindowHandles.Last());
                    if (driver.Url != url)
                    {
                        openNewTabWindow();
                        driver.Navigate().GoToUrl(url);
                    }
                }
            }
            if (!WindowHandles.TryGetValue("《AI太炎》", out windowHandle))
                WindowHandles.Add("《AI太炎》", driver.CurrentWindowHandle);
            else
            {
                if (driver.WindowHandles.Contains(windowHandle))
                {
                    if (windowHandle != driver.CurrentWindowHandle)
                        WindowHandles["《AI太炎》"] = driver.CurrentWindowHandle;
                }
                else
                {
                    openNewTabWindow();
                    driver.Navigate().GoToUrl(url);
                    WindowHandles["《AI太炎》"] = driver.CurrentWindowHandle;
                }

            }


            //標點
            IWebElement iwe = WaitFindWebElementBySelector_ToBeClickable("#nav-biaodian-tab", 5);
            if (iwe == null) return false;
            iwe.Click();
            //輸入框
            iwe = WaitFindWebElementBySelector_ToBeClickable("#textarea-biaodian");
            if (iwe == null) return false;
            SetIWebElementValueProperty(iwe, x);
            //執行
            iwe = WaitFindWebElementBySelector_ToBeClickable("#button-submit");
            iwe.Click();
            DateTime dt = DateTime.Now;
            //结果怎么样？
            iwe = WaitFindWebElementBySelector_ToBeClickable("#feedback > div.feedback-button.feedback-tip");
            while (iwe == null)//while (iwe.Displayed==false)
            {
                if (DateTime.Now.Subtract(dt).TotalMinutes > 0.6) return false;
                iwe = WaitFindWebElementBySelector_ToBeClickable("#feedback > div.feedback-button.feedback-tip");
            }
            //結果
            iwe = WaitFindWebElementBySelector_ToBeClickable("#output-content");
            if (iwe == null) return false;
            x = iwe.GetAttribute("textContent").Replace("□", "􏿽");
            //driver.Close();//不關閉，以手動評量其標點良窳
            driver.SwitchTo().Window(LastValidWindow);
            return true;
        }
        /// <summary>
        /// 直接取代文字的編輯頁面
        /// </summary>
        internal static string DirectlyReplacingCharactersPageWindowHandle = string.Empty;
        /// <summary>
        /// 直接取代文字
        /// </summary>
        /// <param name="character">要直接被取代的單字（regexfrom）及取代成的單字（regexto）的字串陣列</param>
        /// <returns>成功則傳回true</returns>
        internal static bool DirectlyReplacingCharacters(StringInfo character)
        {
            #region 防呆

            if (character.LengthInTextElements != 2) { Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("指定的字元長度不對！請檢查"); return false; }
            if (character.SubstringByTextElements(0, 1) == character.SubstringByTextElements(1, 1)) { Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("所指定取代的字元相同，請重設！"); return false; }
            if (!Form1.IsChineseString(character.SubstringByTextElements(0, 1)) || !Form1.IsChineseString(character.SubstringByTextElements(1, 1))) { return false; }

            #endregion

            try
            {
                if (LastValidWindow != driver.CurrentWindowHandle)
                    driver.SwitchTo().Window(LastValidWindow);
                //else
                //    LastValidWindow = driver.CurrentWindowHandle;
            }
            catch (Exception)
            {
                driver.SwitchTo().Window(LastValidWindow);
            }

            string editUrl;// = string.Empty;
                           //找到「編輯」超連結
            IWebElement iwe = Edit_Linkbox;//waitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > div:nth-child(2) > a:nth-child(2)");
            if (iwe == null)
            {
                //iwe = driver.FindElement(By.XPath("//*[@id=\"content\"]/div[4]/div[2]/a[2]"));
                //iwe = driver.FindElement(By.XPath("/html/body/div[2]/div[4]/div[2]/a[2]"));

                driver.SwitchTo().Window(LastValidWindow);
                //找到「編輯」超連結
                iwe = Edit_Linkbox;//waitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > div:nth-child(2) > a:nth-child(2)");
                if (iwe == null)
                {
                    driver.SwitchTo().Window(driver.WindowHandles.Last());
                    iwe = Edit_Linkbox;//waitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > div:nth-child(2) > a:nth-child(2)");
                    if (iwe == null)
                    {
                        driver.SwitchTo().Window(driver.WindowHandles.LastOrDefault());
                        iwe = Edit_Linkbox;//waitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > div:nth-child(2) > a:nth-child(2)");
                        if (iwe == null)
                        {
                            string url = ActiveForm1.textBox3Text;
                            if (Form1.IsValidUrl＿ImageTextComparisonPage(url))
                            {
                                foreach (var item in driver.WindowHandles)
                                {
                                    if (ReplaceUrl_Box2Editor(driver.Url) == url)
                                    {
                                        iwe = Edit_Linkbox;//waitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > div:nth-child(2) > a:nth-child(2)");
                                        if (iwe == null)
                                        {
                                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請開啟有效的圖文對照頁面；若是新頁面，請先儲存，再執行此功能。");
                                            return false;
                                        }
                                        else
                                            break;
                                    }
                                }
                                iwe = Edit_Linkbox;//waitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > div:nth-child(2) > a:nth-child(2)");
                                if (iwe == null)
                                {
                                    bool found = false;
                                    if (Form1.IsValidUrl＿ImageTextComparisonPage(ActiveForm1.textBox3Text))
                                    {
                                        foreach (var item in driver.WindowHandles)
                                        {
                                            driver.SwitchTo().Window(item);
                                            if (ReplaceUrl_Box2Editor(driver.Url) == ActiveForm1.textBox3Text) { found = true; break; }
                                        }
                                    }
                                    if (!found)
                                    {
                                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請開啟有效的圖文對照頁面");
                                        return false;
                                    }
                                    iwe = Edit_Linkbox;//waitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > div:nth-child(2) > a:nth-child(2)");
                                    if (iwe == null)
                                    {
                                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請開啟有效的圖文對照頁面");
                                        return false;
                                    }

                                }

                            }
                            else
                            {
                                iwe = Edit_Linkbox;//waitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(7) > div:nth-child(2) > a:nth-child(2)");
                                if (iwe == null)
                                {
                                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請開啟有效的圖文對照頁面");
                                    return false;
                                }
                            }

                        }
                    }
                }
                editUrl = iwe.GetAttribute("href");
            }
            else
            {//取得「編輯」頁面的URL
                editUrl = iwe.GetAttribute("href");
            }

            if (DirectlyReplacingCharactersPageWindowHandle == string.Empty)
            {
                foreach (var item in driver.WindowHandles)
                {

                    string url;
                    try
                    {
                        url = ReplaceUrl_Box2Editor(driver.SwitchTo().Window(item).Url);
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                    //if (url.StartsWith("https://ctext.org/wiki.pl?") && url.Contains("&action=editchapter"))
                    if (url == editUrl)
                    {
                        DirectlyReplacingCharactersPageWindowHandle = driver.CurrentWindowHandle; break;
                    }
                }
            }
            else//如果 DirectlyReplacingCharactersPageWindowHandle 非空值
                if (!driver.WindowHandles.Contains(DirectlyReplacingCharactersPageWindowHandle))
                DirectlyReplacingCharactersPageWindowHandle = string.Empty; //goto reOpenEdittab; }
            reOpenEdittab:
            //如果分頁中沒有開啟「編輯」頁面
            if (DirectlyReplacingCharactersPageWindowHandle == string.Empty)
            {

                //開啟完整編輯頁面
                //openNewTabWindow();
                try
                {
                    //driver.SwitchTo().NewWindow(WindowType.Tab);
                    openNewTabWindow();
                }
                catch (Exception)
                {
                    driver.SwitchTo().Window(LastValidWindow);
                    try
                    {
                        driver.SwitchTo().NewWindow(WindowType.Tab);

                    }
                    catch (Exception)
                    {
                        driver.SwitchTo().Window(driver.WindowHandles.Last());
                        //LastValidWindow = driver.WindowHandles.Last();
                        driver.SwitchTo().NewWindow(WindowType.Tab);
                    }

                }
                try
                {
                    driver.Navigate().GoToUrl(editUrl);
                }
                catch (Exception ex)
                {
                    switch (ex.HResult)
                    {
                        case -2146233088:
                            if (ex.Message.StartsWith("The HTTP request to the remote WebDriver server for URL "))
                            {
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("連線超時，請再重試。感恩感恩　南無阿彌陀佛");
                            }
                            else
                            {
                                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                            }
                            return false;
                        default:
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                            return false;
                    }
                }

                //取代區中的「名稱」欄名
                //while (null == waitFindWebElementBySelector_ToBeClickable("#content > table.restable > tbody > tr > td > table > tbody > tr:nth-child(1) > th:nth-child(1)", 0.2)) { }
                iwe = WaitFindWebElementBySelector_ToBeClickable("#content > table.restable > tbody > tr > td > table > tbody > tr:nth-child(1) > th:nth-child(1)", 10);
                if (iwe != null)
                    DirectlyReplacingCharactersPageWindowHandle = driver.CurrentWindowHandle;
                else
                    return false;
            }
            else
            {//如果現成的分頁有找到「編輯」頁面則切換到該頁面
                try
                {
                    driver.SwitchTo().Window(DirectlyReplacingCharactersPageWindowHandle);
                }
                catch (Exception err)
                {
                    DirectlyReplacingCharactersPageWindowHandle = string.Empty;
                    if (editUrl != string.Empty) goto reOpenEdittab;
                    else
                    { Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(err.Message); return false; }
                }
            }

            //「內容:」欄位文字方塊控制項
            iwe = Textarea_data_Edit_textbox;//waitFindWebElementBySelector_ToBeClickable("#data");
            if (iwe == null) { DirectlyReplacingCharactersPageWindowHandle = string.Empty; goto reOpenEdittab; }

            #region 輸入取代後的值

            if (!SetIWebElementValueProperty(iwe, iwe.GetAttribute("value").Replace(character.SubstringByTextElements(0, 1), character.SubstringByTextElements(1, 1)))) Debugger.Break();
            /* 20240913 作廢
            //複製要編輯的文本
            iwe.SendKeys(OpenQA.Selenium.Keys.Control + "a");//直接用 iwe.Text讀取，若要取代多個便不行
            iwe.SendKeys(OpenQA.Selenium.Keys.Control + "c");
            if (iwe.Text.Length > 200000)
                Thread.Sleep(200);
            //借用變數，執行逕行取代
            editUrl = Clipboard.GetText().Replace(character.SubstringByTextElements(0, 1), character.SubstringByTextElements(1, 1));
            try
            {
                Clipboard.SetText(editUrl);
                if (iwe.Text.Length > 200000)
                    Thread.Sleep(200);
            }
            catch (Exception)
            {
            }
            //清除原文本
            iwe.Clear();
            //貼上已編輯的文本
            iwe.SendKeys(OpenQA.Selenium.Keys.Shift + OpenQA.Selenium.Keys.Insert);
            */
            #endregion



            //iwe = null;
            //ReadOnlyCollection<IWebElement> iwes = driver.FindElements(By.ClassName("resrow"));
            //foreach (var item in iwes)            
            //{
            //    string outerHTML = item.GetAttribute("outerHTML");
            //    if(outerHTML== "<a href=\"#\" onclick=\"document.getElementById('regexfrom').value='"+ character.SubstringByTextElements(0, 1) 
            //        + "'; document.getElementById('regexto').value='"+ character.SubstringByTextElements(1, 1) + 
            //        "'; document.getElementById('regexname').value='"+ character.SubstringByTextElements(1, 1) + "'; applyregex(); return false;\">執行</a>")

            //    //if (outerHTML.Length<220 && outerHTML.Contains(">執行<") && outerHTML.Contains("'regexfrom').value='" + character.SubstringByTextElements(0,1)) && outerHTML.Contains("'regexto').value='" + character.SubstringByTextElements(1,1)))
            //    {
            //        iwe = item;//取得要「執行」的元件
            //        break;
            //    }
            //}
            //if (iwe != null) iwe.Click();
            driver.SwitchTo().Window(LastValidWindow);
            return true;
        }

        /// <summary>
        /// C:\Users\oscar\Dropbox\《古籍酷》AI%20OCR%20待改進者隨記%20感恩感恩 讚歎讚歎 南無阿彌陀佛.docx
        /// 
        /// </summary>
        internal static Microsoft.Office.Interop.Word.Document ImproveGJcoolOCRMemoDoc;
        /// <summary>
        /// 作為Word 文件是否已經開啟的判鄉
        /// </summary>
        internal static Task taskOpeningDocument = null;
        /// <summary>
        /// 《古籍酷》與《看典古籍》OCR未善之隨記
        /// 必須焦點在textBox1才行！！20240313
        /// Alt + k : 將選取的字詞句及其網址位址送到以下檔案的末後
        /// C:\Users\oscar\Dropbox\《古籍酷》AI%20OCR%20待改進者隨記%20感恩感恩 讚歎讚歎 南無阿彌陀佛.docx
        /// 20240212大年初三
        /// </summary>
        /// <param name="imporvement">要改進的字詞句（textBox1中被選取的字串）</param>
        internal static void ImproveGJcoolKandiangujiOCRMemo(string imporvement, string url, string preName = "《古籍酷》AI ")
        {
            if (taskOpeningDocument != null && !taskOpeningDocument.IsCompleted)

            {

                taskOpeningDocument.Wait();

            }


            //TextBox tb = null;
            //if (ActiveForm1.InvokeRequired)
            //{
            //    ActiveForm1.Invoke((MethodInvoker)delegate
            //    {
            // 你的程式碼
            //if (!ActiveForm1.Controls["textBox1"].Focused) return;
            //tb = ActiveForm1.Controls["textBox1"] as TextBox;
            //if (tb.SelectionLength == 0) return;
            string f = Path.Combine(Mdb.DropBoxPathIncldBackSlash, preName + "OCR 待改進者隨記 感恩感恩　讚歎讚歎　南無阿彌陀佛Append.docx");

            //20241226 因《古籍酷》執事菩薩不處理，故今改用此檔為中介，再剪去與　他們共用的原檔，否則檔案太大、內容太多，執行效率太差。感恩感恩　南無阿彌陀佛
            //if (preName == "《古籍酷》AI ") f = f.Replace(".docx", "Append.docx");
            //20250102 今皆改以Append方式 感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主

            if (!File.Exists(f))
                if (!File.Exists(f = f.Replace("C:\\", "A:\\")))//A槽是我的虛擬機所設定者 20240822
                    return;

                retry:
            if (ImproveGJcoolOCRMemoDoc == null)
            {
                taskOpeningDocument = Task.Run(() =>
                {
                    Microsoft.Office.Interop.Word.Application wordapp = new Microsoft.Office.Interop.Word.Application
                    {
                        Visible = true
                    };
                    wordapp.Activate();
                    wordapp.WindowState = Microsoft.Office.Interop.Word.WdWindowState.wdWindowStateMinimize;
                    //ActiveForm1.AvailableInUseBothKeysMouse();//跨執行緒時會出錯（寫到呼叫端就好）
                    if (ActiveForm1.InvokeRequired)
                    {
                        ActiveForm1.Invoke((MethodInvoker)delegate
                        {
                            ActiveForm1.AvailableInUseBothKeysMouse();
                        });
                    }
                    //ImproveGJcoolOCRMemoDoc = wordapp.Documents.Open("C:\\Users\\oscar\\Dropbox\\《古籍酷》AI OCR 待改進者隨記 感恩感恩　讚歎讚歎　南無阿彌陀佛.docx");
                    ImproveGJcoolOCRMemoDoc = wordapp.Documents.Open(f);
                    //ImproveGJcoolOCRMemoDoc = wordapp.Documents.Open("C:\\Users\\oscar\\Dropbox\\《古籍酷》AI%20OCR%20待改進者隨記%20感恩感恩　讚歎讚歎　南無阿彌陀佛.docx");
                    //ImproveGJcoolOCRMemoDoc.ActiveWindow.Selection.EndKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);

                });
                taskOpeningDocument.Wait();
            }
            else
            {
                string fn = string.Empty;
                try
                {
                    fn = ImproveGJcoolOCRMemoDoc.Name;
                }
                catch (Exception)
                {
                    ImproveGJcoolOCRMemoDoc = null;
                    goto retry;
                }
                if (!fn.StartsWith(preName))
                {
                    Microsoft.Office.Interop.Word.Application wordapp = ImproveGJcoolOCRMemoDoc.Application;
                    ImproveGJcoolOCRMemoDoc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);
                    ImproveGJcoolOCRMemoDoc = wordapp.Documents.Open(f);
                }
            }

            while (ImproveGJcoolOCRMemoDoc == null)
            {
                taskOpeningDocument.Wait();
            }

        reView:
            try
            {
                //-2146823683：此方法或屬性無法使用，因為此命令無法在閱讀中使用。
                if (ImproveGJcoolOCRMemoDoc.ActiveWindow.View.ReadingLayout)//若是閱讀模式
                    ImproveGJcoolOCRMemoDoc.ActiveWindow.View.ReadingLayout = false;
                ImproveGJcoolOCRMemoDoc.ActiveWindow.Selection.EndKey(Microsoft.Office.Interop.Word.WdUnits.wdStory);

            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case -2147417846:
                        if (ex.Message.StartsWith("訊息篩選器顯示應用程式正在忙碌中。"))//訊息篩選器顯示應用程式正在忙碌中。 (發生例外狀況於 HRESULT: 0x8001010A (RPC_E_SERVERCALL_RETRYLATER))
                        {
                            ImproveGJcoolOCRMemoDoc.Application.Activate();
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("請先處理完Word的訊息再按「確定」繼續！");
                            Application.DoEvents();
                            goto reView;
                        }
                        else
                            goto default;
                    default:
                        Console.WriteLine(ex.HResult + ex.Message);
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                        break;
                }
            }
            try
            {
                string lnk = GetPageUrlKeywordLink(imporvement, url, true);
                if (lnk == string.Empty)
                {
                    //if (!tb.Focused) tb.Focus();
                    lnk = GetPageUrlKeywordLink(imporvement, url, true);
                }
                if (ImproveGJcoolOCRMemoDoc.Content.Text.IndexOf(lnk + Environment.NewLine.Substring(0, 1)) == -1)
                {
                    imporvement += ("\t" + lnk);
                    try
                    {
                        ImproveGJcoolOCRMemoDoc.Range().InsertAfter(imporvement + Environment.NewLine);
                    }
                    catch (Exception)
                    {
                        return;
                    }
                    try
                    {
                        ImproveGJcoolOCRMemoDoc.ActiveWindow.ScrollIntoView(ImproveGJcoolOCRMemoDoc.Range(), false);
                    }
                    catch (Exception)
                    {
                        return;
                    }
                    ImproveGJcoolOCRMemoDoc.Save();
                    ImproveGJcoolOCRMemoDoc.Activate();
                    //ImproveGJcoolOCRMemoDoc.Application.Activate();
                    //if (ImproveGJcoolOCRMemoDoc.Application.WindowState == Microsoft.Office.Interop.Word.WdWindowState.wdWindowStateMinimize)
                    //    ImproveGJcoolOCRMemoDoc.Application.WindowState = Microsoft.Office.Interop.Word.WdWindowState.wdWindowStateNormal;
                    //Thread.Sleep(1000);
                    //ImproveGJcoolOCRMemoDoc.Application.WindowState = Microsoft.Office.Interop.Word.WdWindowState.wdWindowStateMinimize;
                    Form1.playSound(Form1.soundLike.done, true);
                }
                else
                    Form1.playSound(Form1.soundLike.info, true);
            }
            catch (Exception)
            {
                ImproveGJcoolOCRMemoDoc = null;
                goto retry;
            }
            //    });
            //}
        }


        /// <summary>
        /// 作為Selenium發生"no such window: target window already closed"例外情形的處理函式
        /// </summary>
        internal static void NoSuchWindowErrHandler()
        {
            Form1.playSound(Form1.soundLike.error, true);
            if (driver == null) DriverNew();
            if (IsWindowHandleValid(driver, LastValidWindow))
                driver.SwitchTo().Window(LastValidWindow);
            else
                try
                {
                    Form1.ResetLastValidWindow();
                }
                catch (Exception)
                {
                    try
                    {
                        driver.SwitchTo().NewWindow(WindowType.Tab);
                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            Debugger.Break();
                            killchromedriverFromHere();
                            killProcesses(new string[] { "chrome" });
                            driver = DriverNew();

                        }
                        catch (Exception)
                        {
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                            throw;
                        }
                    }

                }
        }

        /// <summary>
        /// /// 作為Selenium發生"ChromeDriver only supports characters in the BMP"例外情形的處理函式
        /// -2146233088: ChromeDriver only supports characters in the BMP
        /// </summary>
        /// <param name="iwe">要操作的網頁元件</param>
        /// <param name="clipboardSetText">要貼上的文字（原本Sendkeys要送的按鍵）</param>
        internal static void ChromeDriverOnlySupportsCharactersBMP(IWebElement iwe, string clipboardSetText)
        {
            //Form1.playSound(Form1.soundLike.error, true);
            if (clipboardSetText == string.Empty) return;
            Clipboard.SetText(clipboardSetText);
            iwe.SendKeys(OpenQA.Selenium.Keys.Shift + OpenQA.Selenium.Keys.Insert);//改成貼上

        }
        /// <summary>
        /// 檢索《韻典網》
        /// </summary>
        /// <param name="x"></param>
        /// <returns></returns>
        public static bool LookupYTenx(string x)
        {
            if (!IsDriverInvalid())
            {
                LastValidWindow = driver.CurrentWindowHandle;
            }
            else
            {
                try
                {
                    LastValidWindow = driver.WindowHandles.Last();
                }
                catch (Exception)
                {
                    RestartChromedriver();
                }
            }
            openNewTabWindow();
            driver.Navigate().GoToUrl("https://ytenx.org/");
            //檢索框
            IWebElement iwe = WaitFindWebElementBySelector_ToBeClickable("#search-form > input.search-query.span3");
            DateTime dt = DateTime.Now;
            while (iwe == null)
            {
                iwe = WaitFindWebElementBySelector_ToBeClickable("#search-form > input.search-query.span3");
                if (DateTime.Now.Subtract(dt).TotalSeconds > 5)
                {
                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("逾時！");
                    return false;
                }
            }
            SetIWebElementValueProperty(iwe, x);
            iwe.SendKeys(OpenQA.Selenium.Keys.Enter);
            //檢索結果
            if (null != WaitFindWebElementBySelector_ToBeClickable("body > div.container.container-main > div > div > div.page-header > h1", 5))
                return true;
            else
                return false;
        }
        /// <summary>
        /// 查找《字統網》https://zi.tools/
        /// </summary>
        /// <param name="x">要查找的單字</param>
        /// <returns>執行無誤或毋需出現錯誤訊息則傳回true</returns>
        public static bool LookupZitools(string x)
        {
            if (x.IsNullOrEmpty()) return true;
            if (Math.Abs(Form1.isChineseChar(x, false)) != 1) return true;

            StringInfo si = new StringInfo(x);
            if (si.LengthInTextElements == 0) return false;

            //以上防呆

            TimeSpan ts = new TimeSpan();
        retry:
            try
            {
                if (!IsDriverInvalid())
                {
                    LastValidWindow = driver.CurrentWindowHandle;
                    ts = driver.Manage().Timeouts().PageLoad;
                }
                else
                {
                    LastValidWindow = driver.WindowHandles.Last();
                    Form1.playSound(Form1.soundLike.error, true);
                }
                openNewTabWindow(OpenQA.Selenium.WindowType.Tab);
                driver.Manage().Timeouts().PageLoad = new TimeSpan(0, 0, 4);
                string selector;
                if (si.LengthInTextElements == 1)
                {
                    driver.Navigate().GoToUrl("https://zi.tools/zi/" + x);
                    selector = "#mainContent > span > div.content > div > div.sidebar_navigation > div > div:nth-child(11)";
                }
                else
                {
                    driver.Navigate().GoToUrl("https://zi.tools/?secondary=search");
                    selector = "#search_input";//組字檢索框
                }

                //點擊"Relatives 相關字" .查詢《字統網》多是為找系統有無該異體字，故今改寫為查詢後在頁面尋找「異寫字」的功能，以利跳到該區塊 20240819
                DateTime dt = DateTime.Now;
                IWebElement iwe = null;
                while (true)
                {
                    iwe = WaitFindWebElementBySelector_ToBeClickable(selector);
                    if (iwe != null ||
                    DateTime.Now.Subtract(dt).TotalSeconds > 2 ||
                    (WaitFindWebElementBySelector_ToBeClickable("#content > div:nth-child(2) > div:nth-child(5) > div.mobile-row.ant-col-6 > div.left_row_column_1 > div.left_row_column_1_1 > div.pc-only > span > span > span > span") != null //大字圖
                    && WaitFindWebElementBySelector_ToBeClickable(selector) == null)) break;
                }

                //若沒有 "Relatives 相關字" 之元件，略過，不算錯誤 20250301
                if (iwe == null)
                {
                    //driver.SwitchTo().Window(LastValidWindow);
                    if (ts != new TimeSpan()) driver.Manage().Timeouts().PageLoad = ts;
                    //return true;
                }

                if (si.LengthInTextElements == 1)
                    iwe?.Click();//當DateTime.Now.Subtract(dt).TotalSeconds > 10) break; 時需要 iwe? 會是null值
                else
                {
                    //SetIWebElementValueProperty(iwe, x);//沒反應                    
                    iwe.SendKeys(selm.Keys.Shift + selm.Keys.Insert);
                    //iwe.Click();//將焦點移至組字檢索框
                    SendKeys.SendWait("^{F6}");
                    //SendKeys.SendWait("{F6}");//以下3行無用，不如上一行 20250529
                    //SendKeys.SendWait("{F6}");
                    //iwe.Click();
                }

            }
            catch (Exception ex)
            {
                if (ex.HResult == -2146233088)
                {
                    if (ex.Message.StartsWith("no such window: target window already closed"))
                    {
                        NoSuchWindowErrHandler();
                        goto retry;
                    }
                    else if (ex.Message.StartsWith("An unknown exception was encountered sending an HTTP request to the remote WebDriver server for URL "))
                    {
                        RestartDriver();
                        //MessageBox.Show("請關閉Chrome瀏覽器，並用本程式重新啟動 Chrome瀏覽器", "", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                        //if (ts != new TimeSpan()) driver.Manage().Timeouts().PageLoad = ts;
                        //return false;
                    }
                    else
                    {
                        Console.WriteLine(ex.HResult + ex.Message);
                        Form1.playSound(Form1.soundLike.error, true);
                        //MessageBox.Show(ex.HResult + ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    }
                }
                if (!IsDriverInvalid())
                    driver.SwitchTo().Window(driver.WindowHandles.Last());
                if (ts != new TimeSpan()) driver.Manage().Timeouts().PageLoad = ts;

                return false;
            }

            if (si.LengthInTextElements == 1)
            {
                if (!SetFocusOnWebPageBody())
                {
                    driver.SwitchTo().Window(LastValidWindow);
                    if (ts != new TimeSpan()) driver.Manage().Timeouts().PageLoad = ts;
                    return true;
                }
            }

            if (ts != new TimeSpan()) driver.Manage().Timeouts().PageLoad = ts;
            return true;
        }

        /// <summary>
        /// 將系統焦點移到網頁本體
        /// </summary>
        /// <return>失敗則傳回false</return>
        public static bool SetFocusOnWebPageBody()
        {
            try
            {
                //焦點移到瀏覽器（離開預設的網址列）
                //driver.SwitchTo().Window(driver.CurrentWindowHandle);

                ////// 使用 JavaScript 將焦點移到網頁本體 20240821 Copilot大菩薩：您說得對，JavaScript 無法控制到 Chrome 瀏覽器的外殼。可以嘗試使用 Actions 類來模擬按鍵操作，將焦點移到網頁本體。
                ////((IJavaScriptExecutor)driver).ExecuteScript("window.focus();");
                //// 使用 Actions 類將焦點移到網頁本體
                ////Actions actions = new Actions(driver);
                ////actions.SendKeys(OpenQA.Selenium.Keys.Escape).Perform();
                //// 使用 SendKeys 將焦點移到網頁本體
                ////SendKeys.SendWait("{esc}");
                ////// 使用 Actions 類模擬滑鼠點擊操作
                ////Actions actions = new Actions(driver);
                ////actions.MoveToElement(driver.FindElement(By.TagName("body"))).Click().Perform();
                //// 使用 Windows API 將焦點移到網頁本體 20240821:Selenium 網頁焦點問題解決方法:https://sl.bing.net/TU0iPVtD7k
                ////BringToFront("chrome");
                IntPtr hWnd = GetForegroundWindow();
                SetForegroundWindow(hWnd);
                ////SendKeys.SendWait("{esc}");//會誤送到Form1主表單裡
                SendKeys.SendWait("^{F6}");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.HResult + ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                return false;
            }
            return true;
        }

        /// <summary>
        /// 查找《漢語大詞典》。最後會將結果網址複製到剪貼簿備用。如果有其他複製項目，可開啟剪貼簿檢視器 Win + v 以選用        
        /// </summary>
        /// <param name="x">要查找的單字</param>
        /// <returns>傳回結果網址。找不到或執行有誤傳回null</returns>
        public static string LookupHYDCD(string x)
        {// 20240817 creedit with Gemini大菩薩：程式碼評析與改進建議 ： https://g.co/gemini/share/3f1f65fd36e0 (這個建議蠻好的，有空要再仔細看看。感恩感恩　讚歎讚歎　Gemini大菩薩　南無阿彌陀佛）
            StringInfo si = new StringInfo(x);
            if (si.LengthInTextElements < 2) return null;
            string url = "https://ivantsoi.myds.me/web/hydcd/search.html", urlResult;//= null;
            IWebElement iwe;
        retry:
            if (IsDriverInvalid()) driver.SwitchTo().Window(driver.WindowHandles.Last());
            try
            {
                LastValidWindow = driver.CurrentWindowHandle;
                openNewTabWindow(OpenQA.Selenium.WindowType.Tab);
                driver.Navigate().GoToUrl(url);
                Clipboard.SetText(x);
                //輸入「詞目」方塊（#SearchBox；name="T1"），再按下Enter鍵
                iwe = WaitFindWebElementBySelector_ToBeClickable("#SearchBox");
                if (iwe == null) return null;
                //Console.WriteLine(iwe.TagName + Environment.NewLine + iwe.Text + Environment.NewLine+ iwe.GetAttribute("value"));
                iwe.SendKeys(OpenQA.Selenium.Keys.Shift + OpenQA.Selenium.Keys.Insert);
                iwe.SendKeys(OpenQA.Selenium.Keys.Enter);
                //找不到時
                iwe = WaitFindWebElementBySelector_ToBeClickable("#SearchResult");
                if (iwe != null)
                {
                    Console.WriteLine(iwe.GetAttribute("textContent"));
                    if (iwe.GetAttribute("textContent").StartsWith("抱歉，無此詞語。"))
                        return null;
                    else
                        iwe.Click();
                }
                else
                {
                    iwe = WaitFindWebElementBySelector_ToBeClickable("#SearchResult > p > a > font");//ex:守真
                                                                                                     //iwe = waitFindWebElementBySelector_ToBeClickable("#SearchResult > p > a");//ex:總第 4709 頁，第三卷第 1303 頁
                    iwe?.Click();
                }
                driver.SwitchTo().Window(driver.WindowHandles.Last());
                urlResult = driver.Url;
            }
            catch (Exception ex)
            {
                if (ex.HResult == -2146233088)
                {
                    if (ex.Message.StartsWith("no such window: target window already closed"))
                    {
                        NoSuchWindowErrHandler();
                        //Form1.playSound(Form1.soundLike.error, true);
                        //if (IsWindowHandleValid(driver, LastValidWindow))
                        //    driver.SwitchTo().Window(LastValidWindow);
                        //else
                        //    Form1.ResetLastValidWindow();
                        goto retry;
                    }
                    else
                    {
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                    }
                }
                return null;
            }
            Clipboard.SetText(urlResult);
            return urlResult;
        }
        /// <summary>
        /// 查找《異體字字典》。最後會將結果網址複製到剪貼簿備用。如果有其他複製項目，可開啟剪貼簿檢視器 Win + v 以選用        
        /// </summary>
        /// <param name="x">要查找的單字</param>
        /// <returns>傳回查詢字串及結果網址。執行有誤則二者均傳回null</returns>
        //public static (string urlSearch, string urlResult) LookupDictionary_of_ChineseCharacterVariants(string x)
        public static Tuple<string, string> LookupDictionary_of_ChineseCharacterVariants(string x)
        {// 20240817 creedit with Gemini大菩薩：程式碼評析與改進建議 ： https://g.co/gemini/share/3f1f65fd36e0 (這個建議蠻好的，有空要再仔細看看。感恩感恩　讚歎讚歎　Gemini大菩薩　南無阿彌陀佛）
            StringInfo si = new StringInfo(x);
            Tuple<string, string> tp = new Tuple<string, string>(null, null);
            if (si.LengthInTextElements != 1) return tp;//(null, null);
            string url = "https://dict.variants.moe.edu.tw/search.jsp?QTP=0&WORD="
                + EncodedStringURL(x)
                 + "#searchL";
            IWebElement iwe;
        retry:
            if (IsDriverInvalid())
                //driver.SwitchTo().Window(driver.WindowHandles.Last());
                driver.SwitchTo().Window(LastValidWindow);
            try
            {
                if (LastValidWindow != driver.CurrentWindowHandle) LastValidWindow = driver.CurrentWindowHandle;
                openNewTabWindow(OpenQA.Selenium.WindowType.Tab);
                driver.Navigate().GoToUrl(url);
                //driver.Navigate().GoToUrl("https://dict.variants.moe.edu.tw/");
                //Clipboard.SetText(x);
                ////輸入「快速搜尋 ariaLabel 」方塊，再按下Enter鍵
                //iwe = waitFindWebElementBySelector_ToBeClickable("#header > div > flex > div:nth-child(3) > div.quick > form > input[type=text]:nth-child(2)");
                //if (iwe == null) return null;
                //iwe.SendKeys(OpenQA.Selenium.Keys.Shift + OpenQA.Selenium.Keys.Insert);
                //iwe.SendKeys(OpenQA.Selenium.Keys.Enter);
                ////網址中還是看得出指令的：https://dict.variants.moe.edu.tw/search.jsp?QTP=0&WORD=%F0%A4%94%AB#searchL 故今改成上式
            }
            catch (Exception ex)
            {
                if (ex.HResult == -2146233088)
                {
                    if (ex.Message.StartsWith("no such window: target window already closed"))
                    {
                        NoSuchWindowErrHandler();
                        //Form1.playSound(Form1.soundLike.error, true);
                        //if (IsWindowHandleValid(driver, LastValidWindow))
                        //    driver.SwitchTo().Window(LastValidWindow);
                        //else
                        //    Form1.ResetLastValidWindow();
                        goto retry;
                    }
                    else
                    {
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                    }
                }
                return tp;
            }
            string urlResult = null;
            try
            {
                //查詢結果：正文 0 字，附收字 0 字
                iwe = WaitFindWebElementBySelector_ToBeClickable("body > main > div > flex > div:nth-child(1)");
                if (iwe != null)
                {
                    if (!iwe.GetAttribute("textContent").EndsWith("查詢結果：正文 0 字，附收字 0 字"))//[ 𪢨 ]， 查詢結果：正文 0 字，附收字 0 字
                                                                                         //[ 襳 ]， 查詢結果：正文 2 字，附收字 0 字
                    {
                        urlResult = driver.Url;
                        Clipboard.SetText(urlResult);
                    }
                }
                else
                {
                    urlResult = driver.Url;
                    Clipboard.SetText(urlResult);
                }
            }
            catch (Exception)
            {
            }
            //return driver.Url;
            return new Tuple<string, string>(url, urlResult);
        }
        /// <summary>
        /// 查找《國語辭典》。最後會將結果網址複製到剪貼簿備用。如果有其他複製項目，可開啟剪貼簿檢視器 Win + v 以選用        
        /// </summary>
        /// <param name="x">要查找的字詞</param>
        /// <returns>傳回查詢字串及結果網址。執行有誤則二者均傳回null</returns>
        //public static (string urlSearch, string urlResult) LookupDictRevised(string x)
        public static Tuple<string, string> LookupDictRevised(string x)
        {// 20240817 creedit with Gemini大菩薩：程式碼評析與改進建議 ： https://g.co/gemini/share/3f1f65fd36e0 (這個建議蠻好的，有空要再仔細看看。感恩感恩　讚歎讚歎　Gemini大菩薩　南無阿彌陀佛）
            StringInfo si = new StringInfo(x); Tuple<string, string> tp = new Tuple<string, string>(null, null);
            if (si.LengthInTextElements < 1) return tp;
            string url = "https://dict.revised.moe.edu.tw/search.jsp?md=1&word="
                + EncodedStringURL(x)
                + "&qMd=0&qCol=1";
            IWebElement iwe;
        retry:
            if (IsDriverInvalid()) driver.SwitchTo().Window(driver.WindowHandles.Last());
            try
            {
                LastValidWindow = driver.CurrentWindowHandle;
                openNewTabWindow(OpenQA.Selenium.WindowType.Tab);
                driver.Navigate().GoToUrl(url);
                //driver.Navigate().GoToUrl("https://dict.variants.moe.edu.tw/");
                //Clipboard.SetText(x);
                ////輸入「快速搜尋 ariaLabel 」方塊，再按下Enter鍵
                //iwe = waitFindWebElementBySelector_ToBeClickable("#header > div > flex > div:nth-child(3) > div.quick > form > input[type=text]:nth-child(2)");
                //if (iwe == null) return null;
                //iwe.SendKeys(OpenQA.Selenium.Keys.Shift + OpenQA.Selenium.Keys.Insert);
                //iwe.SendKeys(OpenQA.Selenium.Keys.Enter);
                ////網址中還是看得出指令的：https://dict.variants.moe.edu.tw/search.jsp?QTP=0&WORD=%F0%A4%94%AB#searchL 故今改成上式
            }
            catch (Exception ex)
            {
                if (ex.HResult == -2146233088)
                {
                    if (ex.Message.StartsWith("no such window: target window already closed"))
                    {
                        //Form1.playSound(Form1.soundLike.error, true);
                        //if (IsWindowHandleValid(driver, LastValidWindow))
                        //    driver.SwitchTo().Window(LastValidWindow);
                        //else
                        //    Form1.ResetLastValidWindow();
                        NoSuchWindowErrHandler();
                        goto retry;
                    }
                    else
                    {
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                    }
                }
                return tp;
            }
            string urlResult = null;
            try
            {
                //查無資料
                iwe = WaitFindWebElementBySelector_ToBeClickable("#searchL > tbody > tr > td");
                if (iwe == null)
                {//if(iwe.GetAttribute("textContent") == "查無資料")
                    urlResult = driver.Url;
                    Clipboard.SetText(urlResult);
                }
            }
            catch (Exception)
            {
            }
            //return driver.Url;
            return new Tuple<string, string>(url, urlResult);
        }

        /// <summary>
        /// 查找《康熙字典網上版》https://www.kangxizidian.com
        /// </summary>
        /// <param name="x">要查找的單字</param>
        /// <returns>執行無誤則傳回true</returns>
        public static bool LookupKangxizidian(string x)
        {
            StringInfo si = new StringInfo(x);
            if (si.LengthInTextElements != 1) return false;
            if (IsDriverInvalid()) driver.SwitchTo().Window(driver.WindowHandles.Last());
            retry:
            try
            {
                LastValidWindow = driver.CurrentWindowHandle;
                openNewTabWindow(OpenQA.Selenium.WindowType.Tab);
                driver.Navigate().GoToUrl("https://www.kangxizidian.com/search/index.php?stype=Word"
                    + "&sword=" + x);// + "&detail=n" );
            }
            catch (Exception ex)
            {
                if (ex.HResult == -2146233088)
                {
                    if (ex.Message.StartsWith("no such window: target window already closed"))
                    {
                        NoSuchWindowErrHandler();
                        goto retry;
                    }
                    else
                    {
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                    }
                }
                return false;
            }
            return true;
        }

        /// <summary>
        /// 0240817Gemini大菩薩：C# 字串網址編碼 https://g.co/gemini/share/e404139f0e17
        /// </summary>
        /// <param name="originalString"></param>
        /// <returns></returns>
        public static string EncodedStringURL(string originalString)
        {
            //string originalString = "這是一個包含中文和特殊符號的字串！&^%";
            //string encodedString = HttpUtility.UrlEncode(originalString);
            return HttpUtility.UrlEncode(originalString);

            //Console.WriteLine("原始字串：{0}", originalString);
            //    Console.WriteLine("編碼後字串：{0}", encodedString);
        }
        public static string DecodedStringURL(string originalString)
        {
            //string originalString = "這是一個包含中文和特殊符號的字串！&^%";
            //string encodedString = HttpUtility.UrlEncode(originalString);
            //return HttpUtility.HtmlDecode(originalString);
            return HttpUtility.UrlDecode(originalString);

            //Console.WriteLine("原始字串：{0}", originalString);
            //    Console.WriteLine("編碼後字串：{0}", encodedString);
        }

        /// <summary>
        /// 取得目前Chrome瀏覽器是否在最大化的狀態
        /// 20240201 Copilot大菩薩：檢查 Chrome 瀏覽器是否已最大化：
        /// 在 Selenium WebDriver 中，並沒有直接的方法可以檢查瀏覽器是否已最大化。但你可以透過比較當前視窗的尺寸和螢幕的解析度來間接判斷。以下是一個可能的 C# 程式碼片段：
        /// 這個 IsBrowserMaximized 函數會回傳一個布林值，表示瀏覽器是否已最大化。請注意，這個方法可能無法在所有情況下正確運作，例如多螢幕設定或者視窗的尺寸與螢幕解析度不完全相同的情況。此外，這個方法需要參考 System.Windows.Forms，所以你需要在你的專案中加入這個參考。如果你的程式是在 .NET Core 或 .NET 5+ 環境下執行，你可能需要改用其他方式來取得螢幕的解析度。希望這個資訊對你有所幫助！
        /// </summary>
        /// <param name="driver"></param>
        /// <returns></returns>
        public static bool IsBrowserMaximized(ChromeDriver driver)
        {
            if (IsDriverInvalid())
            {
                return false;
            }
            try
            {
                Size windowSize = driver.Manage().Window.Size;
                Size workingAreaSize = Screen.PrimaryScreen.WorkingArea.Size;
                return windowSize.Equals(workingAreaSize) || (windowSize.Width >= workingAreaSize.Width && windowSize.Height >= workingAreaSize.Height);

            }
            catch (Exception)
            {
                return false;
            }

        }

        /// <summary>       
        /// 20240430 Copilot大菩薩：下載網頁圖片的錯誤處理：
        /// 以下是一個使用 Selenium 來模擬「另存圖片」的基本範例。請注意，這個範例需要使用到 Actions 類別來模擬鼠標右鍵點擊和選擇「另存圖片」的選項，並且可能需要根據您的瀏覽器和操作系統的具體情況來調整。
        /// 段程式碼會打開圖片的網頁，然後模擬鼠標右鍵點擊圖片，並選擇「另存圖片」的選項。然而，這只是一個基本的範例，並且可能需要根據您的具體情況來調整。例如，處理「另存為」對話框可能需要使用到其他的工具或方法，例如 AutoIt 或 SendKeys。
        /// </summary>
        /// <param name="imageUrl">圖片所在網址</param>
        /// <param name="downloadImgFullName"></param>
        /// <param name="selectedInExplorer"></param>
        /// <returns>成功則傳回true</returns>
        internal static bool DownloadImage(string imageUrl, string downloadImgFullName)
        {
            //var driver = new ChromeDriver();
            openNewTabWindow();
            BringToFront("chrome");
        reGoto:
            try
            {
                driver.Navigate().GoToUrl(imageUrl);
            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case -2146233088://The HTTP request to the remote WebDriver server for URL http://localhost:5908/session/0b71d83809d531eca84ae9d77e0b4888/url timed out after 30.5 seconds.
                        if (ex.Message.EndsWith(" timed out after 30.5 seconds."))
                        {
                            Thread.Sleep(1500);
                            if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("下載書圖的網頁有問題，是否繼續？" +
                                Environment.NewLine + Environment.NewLine + "請確認網頁沒問題再按確定，否則請按取消。感恩感恩　南無阿彌陀佛") == DialogResult.Cancel)
                                return false;
                            goto reGoto;
                        }
                        else
                            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                        break;
                    default:
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                        return false;
                }
            }
            BringToFront("chrome");
            try
            {
                driver.SwitchTo().Window(driver.CurrentWindowHandle);
            }
            catch (Exception)
            {
                return false;
                //throw;
            }
            //IWebElement iw = waitFindWebElementBySelector_ToBeClickable("body > img");
            //Cursor.Position = (Point)iw?.Location;
            ////if (iw != null)  clickCopybutton_GjcoolFastExperience(iw.Location); 

            try
            {
                // 找到圖片元素
                var imageElement = driver.FindElement(By.TagName("img"));

                // 建立 Actions 物件
                var action = new Actions(driver);

                // 模擬鼠標右鍵點擊圖片

                action.ContextClick(imageElement).Perform();

                // 模擬按下「V」鍵，選擇「另存圖片」的選項
                // 注意：這可能需要根據您的瀏覽器和語言設定來調整
                action.SendKeys("v").Perform();
                //SendKeys.Send("{v 2}");
                SendKeys.SendWait("v");

                // TODO: 處理彈出的「另存為」對話框，輸入文件名並點擊「保存」
                // 這可能需要使用到其他的工具或方法，例如 AutoIt 或 SendKeys
                Clipboard.Clear();
                try
                {
                    Clipboard.SetText(downloadImgFullName);
                }
                catch (Exception)
                {
                }
                //Thread.Sleep(1190 + (
                Thread.Sleep(1900 + (
                    800 + Extend_the_wait_time_for_the_Open_Old_File_dialog_box_to_appear_Millisecond < 0 ? 0 : Extend_the_wait_time_for_the_Open_Old_File_dialog_box_to_appear_Millisecond));//最小值（須在重開機後或系統最小負載時）（連「開啟」舊檔之視窗也看不見，即可完成）
                                                                                                                                                                                              //Thread.Sleep(1200);
                                                                                                                                                                                              //Thread.Sleep(500);            


                //輸入：檔案名稱 //SendKeys.Send(downloadImgFullName);
                SendKeys.SendWait("+{Insert}~~");//or "^v"
                                                 //Thread.Sleep(200);
                                                 //SendKeys.Send("{ENTER}");
                                                 //SendKeys.SendWait("%s");
                                                 //Clipboard.Clear();

                //Thread.Sleep(300);
            }
            catch (Exception)
            {
                return false;
            }

            try
            {
                driver.Close();
            }
            catch (Exception)
            {
                driver.SwitchTo().Window(LastValidWindow);//如果沒有切回關閉前的分頁，再打算開新分頁時Selenium就會出錯！20240720
                return false;
            }
            driver.SwitchTo().Window(LastValidWindow);//如果沒有切回關閉前的分頁，再打算開新分頁時Selenium就會出錯！20240720

            ////等待書圖檔下載完成
            //DateTime dt = DateTime.Now;
            //while (!File.Exists(downloadImgFullName))
            //{
            //    if (DateTime.Now.Subtract(dt).TotalSeconds > 28) return false;
            //}
            return true;
        }

        /// <summary>
        /// 在需要連續輸入截圖時 。按下Ctrl並按下滑鼠下一頁鍵時。今因《四庫全書》本《本草綱目》而設 20240510
        /// 須先畫出之截圖區域，然後按下Ctrl並按下滑鼠下一頁鍵時，會自動按下頁面中的[Input picture]連結並再按下 Replace page with this data 按鈕
        /// </summary>
        /// <returns>失敗則傳回false</returns>
        internal static bool Input_picture()
        {
            //按下頁面中的[Input picture]連結

            IWebElement iwe = WaitFindWebElementBySelector_ToBeClickable("#editor > a:nth-child(5)");
            if (iwe != null)
            {
                iwe.Click();
                //再按下 Replace page with this data 按鈕
                iwe = WaitFindWebElementBySelector_ToBeClickable("#pictureinput > input[type=submit]");
                if (iwe != null)
                    iwe.Click();
                else
                {
                    Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("頁面的【Replace page with this data 按鈕】沒找到。");
                    return false;
                }

            }
            else
            {
                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("頁面的【[Input picture]連結元件】沒找到。");
                return false;
            }
            return true;
        }

        /// <summary>
        /// 將指定的程式視窗置於作業系統的最前面、最上端 
        /// Copilot大菩薩 20240704 控制 Chrome 瀏覽器在 C# Windows.Forms 中:
        /// </summary>
        /// <param name="processName">所指定的視窗</param>
        public static void BringToFront(string processName)
        {
            Process[] processes = Process.GetProcessesByName(processName);
            foreach (Process proc in processes)
            {
                // The second parameter 9 means "restore" and "activate" the window.
                //ShowWindow(proc.MainWindowHandle, 1);//Copilot大菩薩：在 ShowWindow 函數中，第二個參數是一個命令，用於指定視窗應該如何顯示。這個參數的值為 9 時，表示 “恢復” 和 “啟動” 視窗。如果視窗最小化或最大化，系統會將其恢復到原來的大小和位置。然後，系統會將該視窗設置為前景視窗。
                //在您的情況下，我們不希望改變視窗的大小和位置，所以我們將該參數改為 1，這表示只 “啟動” 視窗，不改變其大小和位置。
                //在這段程式碼中，我們完全移除了 ShowWindow 函數的調用，只保留了 SetForegroundWindow 函數，這樣就不會改變視窗的大小和位置了。
                SetForegroundWindow(proc.MainWindowHandle);
            }
        }
        /// <summary>
        /// 檢測 driver 是否失效/無效（當前分頁 CurrentWindowHandle）20241008
        /// 若是null則會自行啟用chromedriver 20250215
        /// </summary>
        /// <param name="driver"></param>
        /// <returns></returns>
        internal static bool IsDriverInvalid()
        {
            try
            {
                if (getChromedrivers().Length == 0)
                    RestartChromedriver();

                if (driver == null)
                {
                    Form1.browsrOPMode = Form1.BrowserOPMode.seleniumNew;
                    DriverNew();
                    if (driver == null)
                        RestartDriver();
                }
                try
                {
                    string url;
                    if (driver != null)
                    {
                        if (driver?.WindowHandles.Contains(driver?.CurrentWindowHandle) == false)
                            return true;
                        else
                            url = driver?.Url;
                    }
                }
                catch (Exception)
                {
                    return true;
                }
            }
            catch (Exception)
            {
                return true;
            }
            return false;
        }

        /* GitHub Copilot大菩薩：
         This error happens because the System.TypeInitializationException is thrown when a static constructor throws an exception, or when a static field initializer throws an exception.
            This might be happening because the static field DriverManageTimeoutsPageLoad is being initialized with a value that involves accessing the driver object, which is null at the time of initialization. Specifically, the line internal static readonly TimeSpan DriverManageTimeoutsPageLoad = (TimeSpan)driver?.Manage().Timeouts().PageLoad; tries to access driver?.Manage().Timeouts().PageLoad, but driver is null, leading to a System.InvalidOperationException when trying to access the Value property of a null object.
            To fix this, ensure that driver is properly initialized before it is used in the static field initializer. One way to do this is to initialize DriverManageTimeoutsPageLoad in a static constructor after ensuring driver is not null.
            Example:
        */
        /// <summary>
        /// chromedriver載入頁面的時間上限。預設為5分鐘（原以為是7秒） 20250125
        /// creedit with GitHub Copilot大菩薩。感恩感恩　讚歎讚歎　南無阿彌陀佛　讚美主
        /// </summary>        
        internal static readonly TimeSpan DriverManageTimeoutsPageLoad = driver == null ? TimeSpan.FromMinutes(5) : driver.Manage().Timeouts().PageLoad;
        //internal static readonly TimeSpan DriverManageTimeoutsPageLoad;
        ///// <summary>
        ///// 靜態建構器（constructor），靜態成員之初始化要在此中進行
        ///// </summary>
        //static Browser()
        //{
        //    if (driver != null)
        //    {
        //        DriverManageTimeoutsPageLoad = (TimeSpan)driver.Manage().Timeouts().PageLoad;
        //    }
        //    else
        //    {
        //        // Handle the case where driver is null
        //        //DriverManageTimeoutsPageLoad = TimeSpan.FromSeconds(7); // default value
        //        DriverManageTimeoutsPageLoad = TimeSpan.FromMinutes(5); // default value
        //    }
        //}


        /// <summary>
        /// 進行[《看典古籍·古籍全文檢索》](https://kandianguji.com/search) (d=dian 典) ，成功則傳回true。20241008
        /// </summary>
        /// <param name="searchTxt">要檢索的文字</param>
        /// <returns></returns>
        public static bool KanDianGuJiSearchAll(string searchTxt)
        {
            //防呆：非中文則駁回
            if (Math.Abs(Form1.isChineseChar(searchTxt, true)) != 1) return false;

            void openNewtab(string strUrl)
            {
                openNewTabWindow();
                GoToUrlandActivate(strUrl, true);
            }
            bool exact = false;
            const string url = "https://kandianguji.com/search";
            if (DialogResult.OK == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否要【精確檢索】？")) exact = true;
            TimeSpan ts = new TimeSpan();
            if (!IsDriverInvalid())
            {
                LastValidWindow = driver.CurrentWindowHandle;
                ts = driver.Manage().Timeouts().PageLoad;
                if (ts < DriverManageTimeoutsPageLoad) ts = DriverManageTimeoutsPageLoad;
            }
        reload:
            try
            {
                if (!IsDriverInvalid())
                {
                    if (ts != driver.Manage().Timeouts().PageLoad) ts = driver.Manage().Timeouts().PageLoad;
                    driver.Manage().Timeouts().PageLoad = new TimeSpan(0, 0, 3);
                    if (driver.Url != url)
                    {
                        if (ts != driver.Manage().Timeouts().PageLoad) ts = driver.Manage().Timeouts().PageLoad;
                        driver.Manage().Timeouts().PageLoad = new TimeSpan(0, 0, 3);
                        openNewtab(url);
                    }
                }
                else
                {
                    if (ts != driver.Manage().Timeouts().PageLoad) ts = driver.Manage().Timeouts().PageLoad;
                    driver.Manage().Timeouts().PageLoad = new TimeSpan(0, 0, 3);
                    openNewtab(url);
                }

                driver.SwitchTo().Window(driver.CurrentWindowHandle);
                BringToFront("chrome");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.HResult + ex.Message);
                if (ts != new TimeSpan()) driver.Manage().Timeouts().PageLoad = ts;
                return false;
            }
            if (driver.Url == "about:blank")
            {
                driver.Manage().Timeouts().PageLoad = DriverManageTimeoutsPageLoad;
                goto reload;
            }

            DateTime dt = DateTime.Now;

            // 繁簡同檢
            IWebElement iwe = WaitFindWebElementBySelector_ToBeClickable("#search_select");
            while (iwe == null)
            {
                iwe = WaitFindWebElementBySelector_ToBeClickable("#search_select");
                if (DateTime.Now.Subtract(dt).TotalSeconds > 3 && iwe == null)
                {
                    if (ts != new TimeSpan() && IsDriverInvalid() == false) driver.Manage().Timeouts().PageLoad = ts;
                    return false;
                }
            }
            //iwe.Click();
            //iwe = waitFindWebElementBySelector_ToBeClickable("#search_select > option:nth-child(2)");
            //iwe.Click();
            //iwe.SendKeys(OpenQA.Selenium.Keys.Down);
            //iwe.SendKeys(OpenQA.Selenium.Keys.Down);
            //iwe.Click();
            SetIWebElementValueProperty(iwe, "jianfan");
            //SetIWebElementValueProperty(iwe, string.Empty);
            //SetIWebElementValueProperty(iwe, "jianfan");

            // 設定匹配模式
            iwe = WaitFindWebElementBySelector_ToBeClickable("#search_mode");
            try
            {
                iwe.Click();
            }
            catch (Exception)
            {
                return false;
            }
            if (exact)
            {
                WaitFindWebElementBySelector_ToBeClickable("#search_mode > option:nth-child(2)").Click();
                //SetIWebElementValueProperty(iwe, "accurate");
            }
            else
            {
                WaitFindWebElementBySelector_ToBeClickable("#search_mode > option:nth-child(1)").Click();
                //SetIWebElementValueProperty(iwe, "vague");
            }


            // 檢索詞
            iwe = WaitFindWebElementBySelector_ToBeClickable("#search_input");
            while (iwe == null)
            {
                iwe = WaitFindWebElementBySelector_ToBeClickable("#search_input");
                if (DateTime.Now.Subtract(dt).TotalSeconds > 3 && iwe == null)
                {
                    if (ts != new TimeSpan()) driver.Manage().Timeouts().PageLoad = ts;
                    return false;
                }
            }
            iwe.Clear();
            //iwe.SendKeys(OpenQA.Selenium.Keys.Shift + OpenQA.Selenium.Keys.Insert);
            //SetIWebElementValueProperty(iwe, string.Empty);//設定值無法讓「檢索」按鈕運作，必須用SendKeys才行，故得用到剪貼簿了（因為Selenium還不支援非BMP的字面）
            SetIWebElementValueProperty(iwe, searchTxt);//設定值無法讓「檢索」按鈕運作
                                                        //iwe.SendKeys(OpenQA.Selenium.Keys.Enter);//按下Enter鍵也無效
            iwe.SendKeys(OpenQA.Selenium.Keys.Space);//按下Enter鍵也無效


            // 「檢索」按鈕。好奇怪的檢索按鈕與機制！
            iwe = WaitFindWebElementBySelector_ToBeClickable("#search_button");
            if (iwe == null)
            {
                if (ts != new TimeSpan()) driver.Manage().Timeouts().PageLoad = ts;
                return false;
            }
            //iwe.Click();
            iwe.SendKeys(OpenQA.Selenium.Keys.Space);
            if (ts != new TimeSpan()) driver.Manage().Timeouts().PageLoad = ts;
            return true;
        }
        /// <summary>
        /// chromedriver被誤關時 20241008
        /// </summary>
        internal static void RestartChromedriver()
        {
            killchromedriverFromHere();
            driver = null;
            if (Form1.browsrOPMode != Form1.BrowserOPMode.seleniumNew)
                Form1.browsrOPMode = Form1.BrowserOPMode.seleniumNew;
            DriverNew();
        }
        /// <summary>
        /// 檢索《漢籍全文資料庫》，成功則傳回true。20241008
        /// Alt + h
        /// </summary>
        /// <param name="searchTxt"></param>
        /// <returns></returns>
        public static bool HanchiSearch(string searchTxt)
        {
            if (driver == null) return false;
            bool free = true, inside = false;
            if (!IsDriverInvalid())
            {
                LastValidWindow = driver.CurrentWindowHandle;
                ChromeSetFocus();
                if (!driver.Url.StartsWith("https://hanchi.ihp.sinica.edu.tw/"))
                {
                    if (DialogResult.OK == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否是【授權使用】？"))
                        free = false;
                }
                else
                {
                    if (driver.Url != "https://hanchi.ihp.sinica.edu.tw/ihp/hanji.htm")
                    {
                        inside = true;
                    }
                    else
                    {
                        if (DialogResult.OK == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否是【授權使用】？"))
                            free = false;

                    }


                }

            }
            else
                if (DialogResult.OK == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否是【授權使用】？")) free = false;

            const string url = "https://hanchi.ihp.sinica.edu.tw/ihp/hanji.htm";

            if (!IsDriverInvalid())
            {
                if (!driver.Url.StartsWith("https://hanchi.ihp.sinica.edu.tw/"))
                {
                    try
                    {
                        openNewTabWindow();
                        GoToUrlandActivate(url, true);
                    }
                    catch (Exception)
                    {
                        return false;
                    }
                }
                else
                {
                    if (!inside)
                        driver.Url = url;
                }
            }
            else
            {
                try
                {
                    openNewTabWindow();
                    GoToUrlandActivate(url, true);
                }
                catch (Exception)
                {
                    return false;
                }
            }
            DateTime dt = DateTime.Now;
            IWebElement iwe = null; ReadOnlyCollection<IWebElement> iwes = driver.FindElements(By.TagName("IMG"));
            if (!inside)
            {
                while (iwe == null)
                {
                    if (free)
                    //iwe = waitFindWebElementBySelector_ToBeClickable("body > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td > table > tbody > tr:nth-child(4) > td > a:nth-child(8) > img");
                    //iwe = waitFindWebElementByName_ToBeClickable("免費使用", 2);
                    {
                        foreach (var item in iwes)
                        {
                            if (item.GetAttribute("alt") == "免費使用")
                            {
                                iwe = item; break;
                            }
                        }
                    }
                    else
                    {
                        //iwe = waitFindWebElementBySelector_ToBeClickable("body > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td > table > tbody > tr:nth-child(4) > td > a:nth-child(9) > img");
                        //iwe = waitFindWebElementByName_ToBeClickable("授權使用", 2);
                        foreach (var item in iwes)
                        {
                            if (item.GetAttribute("alt") == "授權使用")
                            {
                                iwe = item; break;
                            }
                        }
                    }
                    if (DateTime.Now.Subtract(dt).TotalSeconds > 5) return false;
                    if (null != WaitFindWebElementBySelector_ToBeClickable("#frmTitle > table > tbody > tr:nth-child(2) > td > table > tbody > tr:nth-child(1) > td > input[type=text]:nth-child(2)")) goto search;
                }
                iwe.Click();
            }
        search:
            dt = DateTime.Now;
            //keyword
            iwe = null;
            while (iwe == null)
            {
                iwe = WaitFindWebElementBySelector_ToBeClickable("#frmTitle > table > tbody > tr:nth-child(2) > td > table > tbody > tr:nth-child(1) > td > input[type=text]:nth-child(2)");
                if (DateTime.Now.Subtract(dt).TotalSeconds > 5) return false;
            }

            SetIWebElementValueProperty(iwe, searchTxt);
            iwe.SendKeys(OpenQA.Selenium.Keys.Enter);

            return true;
        }

        public static bool HanchiTextReadinginPagebyPage()
        {//找不到元件！！沒完成

            //手動開啟【圖】頁面 https://hanchi.ihp.sinica.edu.tw/ihpc/hanji_book?0^0^0529174^DD485BB0290179000100001-1
            bool gotoTab()
            {
                bool found = false;
                for (int i = driver.WindowHandles.Count - 1; i > -1; i--)
                {
                    driver.SwitchTo().Window(driver.WindowHandles[i]);
                    if (driver.Url.StartsWith("https://hanchi.ihp.sinica.edu.tw/ihpc/hanji_book?"))
                    {
                        found = true;
                        break;
                    }
                }
                return found;
            }
            IWebElement iwe = null;

            //檢查頁面
            if (!IsDriverInvalid())
            {
                LastValidWindow = driver.CurrentWindowHandle;

                if (!driver.Url.StartsWith("https://hanchi.ihp.sinica.edu.tw/ihpc/hanji_book?"))
                    if (!gotoTab())
                    {
                        if (!driver.WindowHandles.Contains(LastValidWindow)) return false;
                        driver.SwitchTo().Window(LastValidWindow);
                        return false;
                    }

            }
            else
            {
                RestartChromedriver();
                LastValidWindow = driver.WindowHandles.Last();
                if (!gotoTab())
                {
                    if (!driver.WindowHandles.Contains(LastValidWindow)) return false;
                    driver.SwitchTo().Window(LastValidWindow);
                    return false;
                }
            }

            //第一頁用手動複製？

            //文字框所在元件：
            //iwe = WaitFindWebElementBySelector_ToBeClickable("html");   
            ////iwe = WaitFindWebElementBySelector_ToBeClickable("body > form > div > div:nth-child(3)");
            //Console.WriteLine(iwe.GetAttribute("textContent"));
            //iwe = driver.FindElement(By.XPath("/html/body/form/center/a/img"));
            //iwe = driver.FindElement(By.TagName("IMG"));
            ////iwe = iwe.FindElement(By.XPath("/html/body/form/div"));
            ////iwe = iwe.FindElement(By.XPath("/html/body/form/div"));
            //ReadOnlyCollection<IWebElement> iwes = driver.FindElements(By.TagName("IMG"));
            ////Console.WriteLine(iwe.GetAttribute("textContent"));


            ////iwe = driver.FindElement(By.XPath("/html/body/form/div"));
            ////iwe = driver.FindElement(By.CssSelector("body > form > div"));
            ////Console.WriteLine(iwe.GetAttribute("textContent"));
            //iwe = driver.FindElement(By.ClassName("fulltext"));
            //Console.WriteLine(iwe.GetAttribute("textContent"));
            //if (iwe == null) {
            //    if (!driver.WindowHandles.Contains(LastValidWindow)) return false;
            //    driver.SwitchTo().Window(LastValidWindow);
            //    return false;
            //}


            ////iwe.Click();
            ////// 使用 JavaScript 來全選元素內的文字
            ////OpenQA.Selenium.IJavaScriptExecutor js = (OpenQA.Selenium.IJavaScriptExecutor)driver;
            //////不複製
            ////js.ExecuteScript("var range = document.createRange(); range.selectNodeContents(arguments[0]); var sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(range);", iwe);

            ////取得文本框內容
            //string x = iwe.GetAttribute("textContent");
            //x = iwe.Text;
            //x = iwe.GetAttribute("outerText");
            //if (x.IndexOf("-->" + Environment.NewLine + Environment.NewLine + Environment.NewLine + Environment.NewLine) == -1)
            //{
            //    if (!driver.WindowHandles.Contains(LastValidWindow)) return false;
            //    driver.SwitchTo().Window(LastValidWindow);
            //    return false;
            //}
            //x = x.Substring(x.IndexOf("-->" + Environment.NewLine + Environment.NewLine + Environment.NewLine + Environment.NewLine) +
            //    ("-->" + Environment.NewLine + Environment.NewLine + Environment.NewLine + Environment.NewLine).Length);
            //x = x.Substring(x.IndexOf(Environment.NewLine + Environment.NewLine) + (Environment.NewLine + Environment.NewLine).Length);
            ////取得要擷取的文本部分
            //string content = x.Substring(0, x.IndexOf(Environment.NewLine + Environment.NewLine))
            //        .Replace("．　．　．　．　．　．　．　．　．　．　．　．　．　．　．　．　．　．" + Environment.NewLine, string.Empty);

            ////ActiveForm1.TextBox1_Text = content;

            ////手動翻到下一頁？




            ////「下一頁」按鈕
            //iwe = WaitFindWebElementBySelector_ToBeClickable("body > form > center:nth-child(1) > a.pgdown > img");
            //if (iwe == null)
            //    //無「上一頁」按鈕的「下一頁」按鈕
            //    iwe = WaitFindWebElementBySelector_ToBeClickable("body > form > center > a > img");
            //if (iwe == null) return false;
            //iwe.Click();


            Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("複製完文本後再按下確定！");
            string content = Clipboard.GetText().Replace("．　．　．　．　．　．　．　．　．　．　．　．　．　．　．　．　．　．" + Environment.NewLine, string.Empty);


            //回到CTP圖文對照以輸入文本內容
            if (!driver.WindowHandles.Contains(LastValidWindow)) return false;
            driver.SwitchTo().Window(LastValidWindow);
            在Chrome瀏覽器的Quick_edit文字框中輸入文字(driver, content, driver.Url);



            return true;
        }

        /// <summary>
        /// 翻到[Kanripo](https://www.kanripo.org/)或《國學大師》下一卷在複製其文本後即執行 Word VBA Sub 巨集指令「國學大師_Kanripo_四庫全書本轉來」
        /// Ctrl + Shift + 4 20250207 creedit with GitHub Copilot大菩薩
        /// 自動翻至Kanripo下一卷並複製其文字
        /// </summary>
        /// <param name="url">要翻卷的初始網址</param>
        /// <returns>失敗則傳回fasle</returns>
        internal static bool SikuQuanshu_SKQSContextCopyReading()
        {
            if (IsDriverInvalid())
            {
                if (driver != null)
                    driver.SwitchTo().Window(driver.WindowHandles.Last());
                else
                    return false;
            }
            bool found = false;
            if (driver.WindowHandles.Contains(driver.CurrentWindowHandle))
            {
                if (!Form1.IsValidUrl＿keyDownCtrlAdd(driver.CurrentWindowHandle))
                {

                    for (int i = driver.WindowHandles.Count - 1; i > -1; i--)
                    {
                        driver.SwitchTo().Window(driver.WindowHandles[i]);
                        if (Form1.IsValidUrl＿keyDownCtrlAdd(driver.Url)) { found = true; break; }
                    }
                }

            }
            if (!found)
            {
                openNewTabWindow();
                driver.Navigate().GoToUrl(Form1.InstanceForm1.textBox3Text);
            }
            LastValidWindow = driver.CurrentWindowHandle;

            //複製下一卷或單位的內容
            bool result = CopySKQSNextVolume();

            if (driver.WindowHandles.Contains(LastValidWindow))
            {
                driver.SwitchTo().Window(LastValidWindow);
                if (driver.Url != Form1.InstanceForm1.textBox3Text)
                {
                    //Debugger.Break();
                    bool foundWindowHandle = false;
                    for (int i = driver.WindowHandles.Count - 1; i > -1; i--)
                    {
                        driver.SwitchTo().Window(driver.WindowHandles[i]);
                        if (driver.Url == Form1.InstanceForm1.textBox3Text)
                        {
                            foundWindowHandle = true;
                            LastValidWindow = driver.WindowHandles[i];
                            break;
                        }
                    }
                    if (!foundWindowHandle)
                        driver.SwitchTo().Window(LastValidWindow);
                }
            }
            else
                Debugger.Break();


            //BringToFront("chrome");//●●●●●●●●●●●●●●
            //剪貼簿只能單一執行緒
            //Task.Run(() =>
            //{
            //Form1.InstanceForm1.runWordMacro("中國哲學書電子化計劃.國學大師_Kanripo_四庫全書本轉來");
            //});            
            Form1.InstanceForm1.AvailableInUseBothKeysMouse();
            if (Form1.InstanceForm1.TextBox1_Text == string.Empty)
            {
                BringToFront("chrome");
                Form1.InstanceForm1.AvailableInUseBothKeysMouse();
            }
            return result;
        }
        /// <summary>
        /// 複製下一卷《四庫全書》文本
        /// 佛 20250418 新增非《四庫》之文本
        /// </summary>
        /// <returns></returns>
        internal static bool CopySKQSNextVolume()
        {
            string url = string.Empty; bool result = false;
            string urlPrefixDomain = string.Empty;//= url.Substring(url.IndexOf("//") + "//".Length).Substring(0, url.IndexOf("/"));
            string urlPrefix;// = string.Empty; //url.Substring(0, url.IndexOf("//") + "//".Length);            
                             //http://skqs.guoxuedashi.net/wen_2885i/174671.html#002-1a
                             //https://www.kanripo.org/text/KR4h0141/221
                             //https://github.com/kanripo/KR4h0160/blob/master/KR4h0160_049.txt
            for (int i = driver.WindowHandles.Count - 1; i > -1; i--)
            {
                driver.SwitchTo().Window(driver.WindowHandles[i]);
                url = driver.Url;
                if (!url.Contains("//")) break;
                urlPrefixDomain = url.Substring(url.IndexOf("//") + "//".Length, url.IndexOf("/", url.IndexOf("//") + "//".Length) - (url.IndexOf("//") + "//".Length));
                urlPrefix = url.Substring(0, url.IndexOf("//") + "//".Length);//http://skqs.guoxuedashi.net/wen_2885i/174671.html#002-1a //https://www.kanripo.org/text/KR4h0141/221
                                                                              //if (driver.Url.StartsWith("https://www.kanripo.org/"))
                switch (urlPrefixDomain)
                {
                    case "www.kanripo.org":
                        result = true;
                        goto gotoNext;
                    case "github.com":// https://github.com/kanripo/KR4h0160/blob/master/KR4h0160_049.txt":
                        if (url.IndexOf("kanripo/") > -1)
                        {
                            result = true;
                            goto gotoNext;
                        }
                        break;
                    case "skqs.guoxuedashi.net":
                        result = true;
                        goto gotoNext;
                    case "skqs.39017.com"://其實只是將「guoxuedashi.net」改成「39017.com」爾
                        result = true;
                        goto gotoNext;
                    case "www.inindex.com":
                        result = true;
                        goto gotoNext;
                    case "inindex.com":
                        result = true;
                        goto gotoNext;
                    default:
                        break;
                }
            }
            if (!result) return result;
            gotoNext:
            //取得下一卷的網址
            if (urlPrefixDomain == "skqs.guoxuedashi.net" || urlPrefixDomain == "skqs.39017.com")
            {
                if (url.IndexOf(".html") == -1) return false;
                url = GetNextPageUrl(url.Substring(0, url.IndexOf(".html"))) + ".html";
            }
            else if (urlPrefixDomain == "www.kanripo.org")
                url = GetNextPageUrl(url.IndexOf("#") > -1 ? url.Substring(0, url.IndexOf("#")) : url);
            else if (urlPrefixDomain == "github.com")
            {//url = GetNextPageUrl(url.IndexOf("#") > -1 ? url.Substring(0, url.IndexOf("#")) : url);

                //https://github.com/kanripo/KR4h0160/blob/master/KR4h0160_049.txt
                string prefix = url.Substring(0, url.IndexOf("_") + "_".Length);
                string vol = url.Substring(url.IndexOf("_") + "_".Length, url.IndexOf(".txt") - (url.IndexOf("_") + "_".Length));
                if (!int.TryParse(vol, out int volNum)) return false;
                volNum++;
                //url = prefix + volNum.ToString("D3") + ".txt";
                url = prefix + volNum.ToString("D" + vol.Length.ToString()) + ".txt";

                //url = GetNextPageUrl(url.Substring(0, url.IndexOf(".txt"))) + ".txt";
            }
            else if (urlPrefixDomain == "www.inindex.com" || urlPrefixDomain == "inindex.com")
            {//如果是《元引科技引得數字人文資源平臺·中國歷代文獻》
                //string urlOld = driver.Url;                
                //按下到下一單位的按鈕
                IWebElement iwe = WaitFindWebElementBySelector_ToBeClickable("#printView > div:nth-child(3) > div:nth-child(1) > div");//文本內容框
                if (iwe == null) return false;
                string textContent = iwe.GetAttribute("textContent");//.Substring(0, 100);
                int l = textContent.Length;
                l = l > 50 ? 50 : l;
                textContent = textContent.Replace("　", string.Empty).Substring(0, l);
                iwe = WaitFindWebElementBySelector_ToBeClickable("#root > main > div:nth-child(2) > div:nth-child(1) > div:nth-child(1) > div:nth-child(1) > div.gt");
                if (iwe == null) return false;
                BringToFront("chrome");
                iwe.Click();
                //clickCopybutton_GjcoolFastExperience(iwe.Location);
                //if (Cursor.Position != iwe.Location)
                //    Cursor.Position = iwe.Location;

                iwe = WaitFindWebElementBySelector_ToBeClickable("#printView > div:nth-child(3) > div:nth-child(1) > div");//文本內容框

                DateTime dt = DateTime.Now;
                //while (DateTime.Now.Subtract(dt).TotalSeconds < 2) { }//要有這樣才能複製到正確的卷頁單位
                //while (urlOld == driver.Url)
                string textContent1 = iwe.GetAttribute("textContent");
                if (textContent1.Length > 50) textContent1 = textContent1.Replace("　", string.Empty).Substring(0, l);
                //while (iwe.GetAttribute("textContent").Substring(0,100) == textContent)
                while (textContent1 == textContent)
                {
                    textContent1 = iwe.GetAttribute("textContent");
                    if (textContent1.Length > 50) textContent1 = textContent1.Replace("　", string.Empty).Substring(0, l);
                    if (DateTime.Now.Subtract(dt).TotalSeconds > 10) return false;
                }
                //driver.SwitchTo().Window(driver.CurrentWindowHandle);
                //ChromeSetFocus();
                url = driver.Url;
            }
        retry:
            try
            {
                if (urlPrefixDomain != "www.inindex.com")//已於前翻頁了
                    //翻到下一頁（網頁）即處理下一卷的文本
                    driver.Navigate().GoToUrl(url);
            }
            catch (Exception ex)
            {
                switch (ex.HResult)
                {
                    case -2146233088:
                        if (ex.Message.Contains("timed out after"))
                            if (DialogResult.Cancel == Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("逾時，是否繼續？"))
                                return false;
                        goto retry;
                    default:
                        Console.WriteLine(ex.HResult + ex.Message);
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly(ex.HResult + ex.Message);
                        break;
                }
            }

            string iElementSelector = string.Empty;
            //creedit with GitHub Copilot大菩薩：20250207
            // 找到指定的元素
            switch (urlPrefixDomain)
            {
                case "www.kanripo.org":
                    iElementSelector = "#txtcont > p:nth-child(1)";
                    break;
                case "github.com":
                    iElementSelector = "#read-only-cursor-text-area";
                    break;
                case "skqs.guoxuedashi.net":
                    iElementSelector = "body > div:nth-child(3) > div:nth-child(4) > div.col2";
                    break;
                case "skqs.39017.com":
                    iElementSelector = "body > div:nth-child(3) > div:nth-child(4) > div.col2";
                    break;
                case "www.inindex.com":
                    iElementSelector = "#printView > div:nth-child(3) > div:nth-child(1) > div";
                    break;
                case "inindex.com":
                    iElementSelector = "#printView > div:nth-child(3) > div:nth-child(1) > div";
                    break;
                default:
                    break;
            }
            OpenQA.Selenium.IWebElement element;
            try
            {

                element = driver.FindElement(OpenQA.Selenium.By.CssSelector(iElementSelector));
                if (element == null)
                {
                    result = false; return result;
                }
            }
            catch (Exception)
            {
                result = false; return result;
            }

            element.Click();

            // 使用 JavaScript 來全選元素內的文字
            OpenQA.Selenium.IJavaScriptExecutor js = (OpenQA.Selenium.IJavaScriptExecutor)driver;
            //不複製
            js.ExecuteScript("var range = document.createRange(); range.selectNodeContents(arguments[0]); var sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(range);", element);
            //複製
            //js.ExecuteScript(@"var range = document.createRange(); range.selectNodeContents(arguments[0]); var sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(range); document.execCommand('copy');", element);

            //js.ExecuteScript(@"
            //    var range = document.createRange();
            //    range.selectNodeContents(arguments[0]);
            //    var sel = window.getSelection();
            //    sel.removeAllRanges();
            //    sel.addRange(range);
            //    document.execCommand('copy');
            //", element);
            //以上無效
            //以下只是複製純文字字內容
            //js.ExecuteScript(@"
            //    var range = document.createRange();
            //    range.selectNodeContents(arguments[0]);
            //    var sel = window.getSelection();
            //    sel.removeAllRanges();
            //    sel.addRange(range);
            //    navigator.clipboard.writeText(sel.toString()).then(function() {
            //        console.log('Text copied to clipboard');
            //    }).catch(function(error) {
            //        console.error('Error copying text: ', error);
            //    });
            //", element);

            //// 使用 Actions 來模擬 Ctrl + C 鍵盤操作
            //Actions actions = new Actions(driver);
            ////actions.MoveToElement(element).Click().KeyDown(OpenQA.Selenium.Keys.Control).SendKeys("a").SendKeys("c").KeyUp(OpenQA.Selenium.Keys.Control).Perform();
            //actions.MoveToElement(element).Click().KeyDown(OpenQA.Selenium.Keys.Control).SendKeys("c").KeyUp(OpenQA.Selenium.Keys.Control).Perform();
            //ChromeSetFocus();
            //BringToFront("chrome");
            Clipboard.Clear();
            //element.SendKeys(selm.Keys.Control + "c");//此元件無法與使用者互動
            WaitFindWebElementBySelector_ToBeClickable("body").SendKeys(selm.Keys.Control + "c");//外面的框
            //SendKeys.Send("^c");
            //SendKeys.SendWait("^c");
            //Thread.Sleep(900);
            //Thread.Sleep(150);
            Thread.Sleep(450);
            //DateTime dt = DateTime.Now;
            //try
            //{
            //    while (Clipboard.GetText() == string.Empty)
            //    {
            //        Thread.Sleep(300);
            //        if (DateTime.Now.Subtract(dt).TotalSeconds > 4) return false;
            //    }

            //}
            //catch (Exception)
            //{

            //}
            //try
            //{
            //    if (Clipboard.GetText() == string.Empty) return false;

            //}
            //catch (Exception)
            //{

            //}


            return result;


            string GetNextPageUrl(string currentUrl)
            {
                // 找到最後一個斜杠的位置
                int lastSlashIndex = currentUrl.LastIndexOf('/');

                // 取得頁碼部分
                string pageNumberStr = currentUrl.Substring(lastSlashIndex + 1);

                // 將頁碼轉換為整數並加 1
                int pageNumber = int.Parse(pageNumberStr) + 1;

                // 生成下一頁的 URL
                string nextUrl = currentUrl.Substring(0, lastSlashIndex + 1) + pageNumber.ToString();

                return nextUrl;
            }
        }

        /// <summary>
        /// 讓Chrome瀏覽器取得焦點
        /// </summary>
        /// <returns></returns>
        internal static bool ChromeSetFocus()
        {
            if (driver == null) return false;
            if (!IsDriverInvalid())
            {
                driver.SwitchTo().Window(driver.CurrentWindowHandle);
                //if (ActiveForm1.Active) BringToFront("chrome");//多執行緒時會出現錯誤
                BringToFront("chrome");
                return true;
            }
            else return false;
        }
        /// <summary>
        /// 使用 JavaScriptExecutor 輸入中文。作為Selenium 同名方法的替代性實驗方法
        /// 20240811 Gemini大菩薩：CJKV 字集擴充現況：https://g.co/gemini/share/b6380ac335aa
        /// https://g.co/gemini/share/3fdbf8b43c46
        /// </summary>
        /// <param name="element">要輸入文字的網頁元件</param>
        /// <param name="text">要輸入的文字</param>
        /// <returns></returns>
        public static bool Sendkeys(IWebElement element, string text)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            //string text = "你好，世界！";
            js.ExecuteScript($"arguments[0].value = '{text}';", element);

            return true;
        }
        /// <summary>
        /// 以選取文字檢索Google
        /// </summary>
        /// <param name="x">要檢索的文字</param>
        /// <param name="quote">要加引號""就是true</param>
        /// <returns>成功則傳回true</returns>
        internal static bool GoogleSearch(string x, bool quote = false)
        {
            if (x != "")
            {
                x = x.EndsWith("》") ? x.Substring(0, x.Length - 1) : x;
                x = x.EndsWith(Environment.NewLine) ? x.Substring(0, x.Length - 2) : x;
                x = x.EndsWith("\n") ? x.Substring(0, x.Length - 1) : x;
                if (quote) x = "\"" + x + "\"";
                Clipboard.SetText(x);
                //在Selenium模式下，直接以x搜尋網路
                if (Form1.browsrOPMode != Form1.BrowserOPMode.appActivateByName)
                {
                    if (driver != null)
                    {
                        openNewTabWindow(OpenQA.Selenium.WindowType.Tab);
                        try
                        {
                            driver.Navigate().GoToUrl("https://www.google.com/search?q=" + x);
                        }
                        catch (Exception)
                        {
                            return false;
                        }
                    }
                    else
                    {
                        if (Form1.MessageBoxShowOKCancelExclamationDefaultDesktopOnly("是否執行【網路搜尋_元搜尋-同時搜多個引擎】") == DialogResult.OK)
                            Process.Start(ActiveForm1.dropBoxPathIncldBackSlash + @"VS\VB\網路搜尋_元搜尋-同時搜多個引擎\網路搜尋_元搜尋-同時搜多個引擎\bin\Debug\網路搜尋_元搜尋-同時搜多個引擎.exe");
                    }
                }
                else
                    Process.Start(ActiveForm1.dropBoxPathIncldBackSlash + @"VS\VB\網路搜尋_元搜尋-同時搜多個引擎\網路搜尋_元搜尋-同時搜多個引擎\bin\Debug\網路搜尋_元搜尋-同時搜多個引擎.exe");
            }
            return true;
        }
        /// <summary>
        /// 檢查是否是「Please confirm that you are human! 敬請輸入認證圖案」頁面 網址列：https://ctext.org/wiki.pl 20240929 52生日
        /// <returns></returns>
        internal static bool IsConfirmHumanPage()
        {
            bool result = true; int retryCount = 0;
        retry:
            try
            {
                //result = confirm_that_you_are_human = driver.Url == "https://ctext.org/wiki.pl" || Please_confirm_that_you_are_human_Page != null;
                confirm_that_you_are_human = (driver.Url == "https://ctext.org/wiki.pl" || Please_confirm_that_you_are_human_Page != null);
                result = confirm_that_you_are_human;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.HResult + ex.Message);
                switch (ex.HResult)
                {
                    case -2146233088:
                        if (ex.Message.StartsWith("tab crashed"))
                            RestartChromedriver();
                        break;
                    default:

                        break;
                }
                Console.WriteLine(WebDriverWaitTimeSpan.ToString());
                Console.WriteLine(driver.Manage().Timeouts().PageLoad.ToString());
                //Debugger.Break();
                if (driver.Manage().Timeouts().PageLoad < DriverManageTimeoutsPageLoad)
                    driver.Manage().Timeouts().PageLoad = DriverManageTimeoutsPageLoad;
                Thread.Sleep(1000);
                if (retryCount < 2) { retryCount++; goto retry; }

            }
            return result;
            //return confirm_that_you_are_human;

            //if (driver.Url == "https://ctext.org/wiki.pl" ||Please_confirm_that_you_are_human_Page!=null)
            //{
            //    if (WaitFindWebElementBySelector_ToBeClickable("#content > font")?.GetAttribute("textContent") == "Please confirm that you are human! 敬請輸入認證圖案")
            //    {
            //        confirm_that_you_are_human = true;
            //        return true;
            //    }
            //    else
            //        return false;
            //}
            //else
            //    return false;
        }
        /// <summary>
        /// 打開展開/收起閉合大綱標題（章節頁面）
        /// </summary>
        internal static void OutlineTitlesCloseOpenFoldExpandSwitcher()
        {
            if (driver == null) return;
            ActiveForm1.TopMost = false;
            if (IsDriverInvalid())
            {
                driver.SwitchTo().Window(driver.WindowHandles.LastOrDefault());
            }
            if (!driver.Url.StartsWith("https://ctext.org/wiki.pl?if=gb&res="))
            {
                for (int i = driver.WindowHandles.Count - 1; i > -1; i--)
                {
                    string url = driver.SwitchTo().Window(driver.WindowHandles[i]).Url;
                    //if (driver.SwitchTo().Window(driver.WindowHandles[i]).Url.StartsWith("https://ctext.org/wiki.pl?if=gb&res="))
                    if (url.StartsWith("https://ctext.org/wiki.pl?if=gb&res=") ||
                        url.StartsWith("https://ctext.org/wiki.pl?if=en&res="))
                        break;

                }

            }
            ReadOnlyCollection<IWebElement> iwes = driver.FindElements(By.TagName("DIV"));
            foreach (var item in iwes)
            {
                if (item.GetAttribute("title") == "+")
                    try
                    {
                        item.Click();
                    }
                    catch (Exception)
                    {

                    }
            }
            BringToFront("chrome");
        }


    }

}





