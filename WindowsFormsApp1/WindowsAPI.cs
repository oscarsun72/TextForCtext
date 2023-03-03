using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Runtime.InteropServices;
using OpenQA.Selenium.Chrome;
using System.Collections.ObjectModel;
using br = TextForCtext.Browser;
using System.Windows.Forms;
using System.Globalization;
using System.Text.RegularExpressions;

namespace TextForCtext
{
    public class WindowsAPI : IDisposable
    {

        // 取得目前正在前景運行的視窗的識別碼        
        static IntPtr foregroundWindowHandle;
        internal static IntPtr ForegroundWindowHandle
        {
            get
            {
                return foregroundWindowHandle;
            }
        }
        internal static void SetforegroundWindowHandle() { foregroundWindowHandle = GetForegroundWindow(); }

        //#region 建構器（constructor） Implicit

        //internal WindowsAPI() { }
        //#endregion

        /// <summary>
        /// 解構器（destructor）
        /// </summary>
        /// <exception cref="NotImplementedException"></exception>
        public void Dispose()
        {

            //throw new NotImplementedException();
        }

        #region 在 C# 要怎麼利用 Windows API 來取得現行的Chrome瀏覽器的現在正在最前方、正在使用中的分頁頁籤（即其 WindowHandle ）呢？

        #region 20230303 creedit with chatGPT大菩薩

        /// <summary>
        /// GetForegroundWindow 函式 (winuser.h)
        /// https://learn.microsoft.com/zh-tw/windows/win32/api/winuser/nf-winuser-getforegroundwindow
        /// </summary>
        /// <returns></returns>
        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();
        /// <summary>
        /// 那 C# 要怎麼取得現行的Chrome瀏覽器的現在正在最前方、正在使用中的分頁頁籤（即其 WindowHandle ）呢？
        /// 您可以使用 Selenium WebDriver 的 IWebDriver.WindowHandles 屬性取得所有已開啟的分頁頁籤的識別碼(WindowHandle)，然後使用 C# 的 Win32 API GetForegroundWindow() 方法取得目前正在前景運行的視窗的識別碼 (HWND)，接著逐一比對目前正在前景運行的視窗是否為 Chrome 瀏覽器分頁視窗，最後即可得到目前正在使用中的 Chrome 瀏覽器分頁視窗的識別碼 (WindowHandle)。
        /// 以下是範例程式碼：
        /// 這個程式碼會開啟三個分頁頁籤(Google、Facebook、YouTube)，然後使用 Win32 API GetForegroundWindow() 方法取得目前正在前景運行的視窗的識別碼(HWND)，接著逐一比對目前正在前景運行的視窗是否為 Chrome 瀏覽器分頁視窗，最後輸出目前正在使用中的 Chrome 瀏覽器分頁視窗的識別碼 (WindowHandle)。
        /// </summary>
        public string GetChromeActiveTabWindowHandle()
        {
            IWebDriver driver = br.driver ?? br.driverNew();//new ChromeDriver();

            //// 開啟多個分頁頁籤
            //driver.Navigate().GoToUrl("https://www.google.com");
            //driver.ExecuteScript("window.open('https://www.facebook.com')");
            //driver.ExecuteScript("window.open('https://www.youtube.com')");

            // 取得所有已開啟的分頁頁籤的 WindowHandle
            ReadOnlyCollection<string> windowHandles = driver.WindowHandles;


            // 逐一比對已開啟的分頁頁籤是否為目前正在前景運行的視窗
            //var processIDs=br.getChromesPID();
            //foreach (var processId in processIDs)
            //{
            //    var mainWindowHandle = Process.GetProcessById(processId.Id).MainWindowHandle;
            //    if (mainWindowHandle == foregroundWindowHandle)
            //    {
            //        //Console.WriteLine("Active tab WindowHandle: " + handle);
            //        return handle;
            //    }
            //}

            foreach (var handle in windowHandles)
            {
                string handleNum = handle.Substring("CDwindow-".Length);


                //using System.Text.RegularExpressions;
                // 取得目前焦點視窗的 handle

                // 取得當前的視窗 handle
                string currentWindowHandle = handle;// br.driver.CurrentWindowHandle;

                // 提取數字部分
                string currentWindowHandleValue = Regex.Match(currentWindowHandle, @"\d+").Value;

                // 將數字部分轉換成 long 類型
                long currentWindowHandleLong = long.Parse(currentWindowHandleValue, NumberStyles.HexNumber);

                // 將 long 數值轉換成 IntPtr
                IntPtr currentWindowHandlePtr = new IntPtr(currentWindowHandleLong);


                // 將當前的視窗 handle 轉換成 IntPtr
                //IntPtr currentWindowHandlePtr = new IntPtr(long.Parse(currentWindowHandle));

                //// 比較兩個 handle 是否相同
                //bool handlesMatch = currentWindowHandlePtr == foregroundWindowHandle;


                //long handleValue;
                //if (handle.StartsWith("CDwindow-") && 
                //    long.TryParse(handle.Substring("CDwindow-".Length),
                //    NumberStyles.HexNumber, CultureInfo.InvariantCulture, out handleValue))
                //{
                //var processId = Process.GetProcesses().FirstOrDefault(p => p.MainWindowHandle == new IntPtr(handleValue)).Id;
                //IntPtr handleValue = new IntPtr(Convert.ToInt64(handle.Substring("CDwindow-".Length), 16));
                //var mainWindowHandle = Process.GetProcessById(processId).MainWindowHandle;
                //if (mainWindowHandle == foregroundWindowHandle)
                //if (handleValue == foregroundWindowHandle)
                //if (currentWindowHandlePtr == foregroundWindowHandle)
                if (currentWindowHandlePtr.Equals(foregroundWindowHandle))
                {
                    return handle;
                }
                //}



                //long handleValue;
                //if (long.TryParse(handle.Substring("CDwindow-".Length), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out handleValue))
                //{
                //    var processId = Process.GetProcesses().FirstOrDefault(p => p.MainWindowHandle == new IntPtr(handleValue)).Id;
                //    var mainWindowHandle = Process.GetProcessById(processId).MainWindowHandle;
                //    if (mainWindowHandle == foregroundWindowHandle)
                //    {
                //        return handle;
                //    }
                //}




                //var processId = Process.GetProcesses().FirstOrDefault(p => p.MainWindowHandle == new IntPtr(long.Parse(handle))).Id;
                //var mainWindowHandle = Process.GetProcessById(processId).MainWindowHandle;
                //if (mainWindowHandle == foregroundWindowHandle)
                //{
                //    //Console.WriteLine("Active tab WindowHandle: " + handle);
                //    return handle;
                //}


                //int processId;
                //if (!int.TryParse(handle, out processId)) continue;

                //var mainWindowHandle = Process.GetProcessById(processId).MainWindowHandle;
                //if (mainWindowHandle == foregroundWindowHandle)
                //{
                //    return handle;
                //}
            }
            return "";
        }

        #endregion

        #region YouChat菩薩


        /// <summary>
        /// 那 C# 要怎麼利用 Windows API 來取得現行的Chrome瀏覽器的現在正在最前方、正在使用中的分頁頁籤（即其 WindowHandle ）呢？
        /// 如果你要從C#中使用Windows API來取得現行的Chrome瀏覽器的現在正在最前方、正在使用中的分頁頁籤（即其 WindowHandle ），可以使用EnumWindows來列舉所有頁籤，然後使用GetWindowText來獲得每個頁籤的標題，最後使用GetWindowThreadProcessId來獲得每個頁籤的Window Handle，並且檢查它是否為Chrome瀏覽器的頁籤。示例代碼如下：
        /// </summary>
        /// <returns></returns>

        //public static IntPtr GetChromeWindowHandle()
        //{
        //    IntPtr chromeWindowHandle = IntPtr.Zero;

        //    EnumWindows(
        //        delegate (IntPtr hWnd, int lParam)
        //        {
        //            int length = GetWindowTextLength(hWnd);
        //            if (length == 0)
        //                return true;

        //            StringBuilder builder = new StringBuilder(length);
        //            GetWindowText(hWnd, builder, length + 1);

        //            string title = builder.ToString();

        //            if (title.Contains("Google Chrome"))
        //            {
        //                uint pid;
        //                GetWindowThreadProcessId(hWnd, out pid);

        //                Process p = Process.GetProcessById((int)pid);
        //                if (p != null && p.ProcessName == "chrome")
        //                {
        //                    chromeWindowHandle = hWnd;
        //                    return false;
        //                }
        //            }

        //            return true;
        //        }, 0);

        //    return chromeWindowHandle;
        //}
        #endregion

        #region Bing大菩薩 in Skype
        /*
         @Bing 在 C# 要怎麼利用 Windows API 來取得現行的Chrome瀏覽器的現在正在最前方、正在使用中的分頁頁籤（即其 WindowHandle ）呢？

        Bing，上午 11:33
        您好，這是 Bing。我找到了一些可能有用的資訊。 (stackoverflow.com：https://stackoverflow.com/questions/18897070/getting-the-current-tabs-url-from-google-chrome-using-c-sharp) 
        有人說，以前可以用 FindWindowEx 和 SendMessage 的組合來從 Google Chrome 獲取當前分頁的 URL，但現在這個方法已經失效了，因為 Chrome 似乎自己渲染了所有東西。 (developer.chrome.com：https://developer.chrome.com/docs/extensions/mv3/manifest/activeTab/)
        另一種方法是使用 activeTab 權限來給擴展程序臨時訪問當前活動的分頁，但這需要用戶觸發擴展程序。 (stackoverflow.com:https://stackoverflow.com/questions/19705797/find-the-window-handle-for-a-chrome-browser) 
        還有人說，可以用 FindWindowEx 來找到 Chrome 瀏覽器的窗口句柄，但需要提供類名或窗口名。

        請問您是想要獲取當前分頁的 URL 還是窗口句柄呢？您是否已經安裝了任何擴展程序呢？
         */
        #endregion



        #endregion



    }
}
