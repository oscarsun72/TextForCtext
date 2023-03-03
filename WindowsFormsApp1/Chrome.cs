using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using br = TextForCtext.Browser;
using OpenQA.Selenium;
using WindowsFormsApp1;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Web;

namespace TextForCtext
{
    public class Chrome
    {
        //[DllImport("user32.dll")]
        //static extern IntPtr GetForegroundWindow();

        //[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        //static extern int GetWindowThreadProcessId(IntPtr handle, out int processId);

        //[DllImport("user32.dll", SetLastError = true)]
        //static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        //[DllImport("user32.dll", SetLastError = true)]
        //static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, StringBuilder lParam);


        public static string ActiveTabURL
        {
            get
            {
                string url = getUrl(ControlType.Edit).Trim();
                url = url.StartsWith("https://") ? url : "https://" + url;
                return url;
            }
        }

        //static string getUrl()
        //{
        //    using (Form1 form = new Form1())
        //    {
        //        form.appActivateByName();
        //        //form = null;
        //    }
        //    IntPtr hwnd = GetForegroundWindow();
        //    int pid;
        //    GetWindowThreadProcessId(hwnd, out pid);
        //    Process process = Process.GetProcessById(pid);
        //    if (process.ProcessName == "chrome")
        //    {
        //        //IntPtr hwndAddress = FindWindowEx(hwnd, IntPtr.Zero, "Chrome_OmniboxView", null);
        //        IntPtr hwndAddress = FindWindowEx(hwnd, IntPtr.Zero, "", "input#search-box");
        //        hwndAddress = FindWindowEx(hwnd, IntPtr.Zero, "omnibox", "omnibox");

        //        hwndAddress = FindWindowEx(hwnd, IntPtr.Zero, "Chrome_AutocompleteEditView", null);
        //        hwndAddress = FindWindowEx(hwnd, new IntPtr(0), "omnibox", "");
        //        hwndAddress = FindWindowEx(hwnd, new IntPtr(0), "Chrome_AutocompleteEditView", null);
        //        if (hwndAddress != IntPtr.Zero)
        //        {
        //            StringBuilder urlBuilder = new StringBuilder(1024);
        //            SendMessage(hwndAddress, 0x000D, (IntPtr)urlBuilder.Capacity, urlBuilder);
        //            url = urlBuilder.ToString();

        //            //Console.WriteLine("URL: {0}", url);
        //            return url;
        //        }
        //    }
        //}

        static string browsername = "chrome";
        /// <summary>
        /// 取得Chrome瀏覽器現前網址。結果竟然是我自己之前就實作過的，完全忘了！
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
                Process[] procsBrowser = Process.GetProcessesByName(browsername);
                if (procsBrowser.Length <= 0)
                {
                    //    MessageBox.Show("Chrome is not running");
                    MessageBox.Show(browsername + " " +
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
    }
}
