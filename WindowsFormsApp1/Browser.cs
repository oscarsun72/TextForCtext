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
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.DevTools.V85.ApplicationCache;
using OpenQA.Selenium;
//https://dotblogs.com.tw/supergary/2020/10/29/selenium#images-3

namespace TextForCtext
{
    class Browser
    {
        // 創建Chrome驅動程序對象
        //selm.IWebDriver driver=driverNew();        
        //internal static selm.IWebDriver driver=driverNew();
        internal static ChromeDriver driver = driverNew();
        //static selm.IWebDriver driverNew()



        static Form1 frm;//creedit 
        public Browser(Form1 form)
        {
            frm = form;
        }



        static ChromeDriver driverNew()
        {
            if (driver == null)
            {
                ChromeDriver cDrv;
                try
                {//selenium 如何操作免安裝版的 chrome 瀏覽器 或自訂安裝路徑的 chrome 瀏覽器呢
                    cDrv = new ChromeDriver();
                }
                catch (Exception)
                { //creedit 20230101 : 
                  // 指定 Chrome 瀏覽器的路徑
                     string chrome_path = Form1.getDefaultBrowserEXE();
                    if(chrome_path.IndexOf("chrome") >-1)
                    //# 建立 ChromeDriver 物件
                    cDrv = new ChromeDriver(chrome_path);
                    else
                    throw;
                }
                cDrv.Navigate().GoToUrl(Form1.mainFromTextBox3Text);
                MessageBox.Show("請先登入 Ctext.org 再繼續。按下「確定(OK)」以繼續……");
                return cDrv;
            }
            else
                return driver;
        }
        internal static void 在Chrome瀏覽器的Quick_edit文字框中輸入文字(ChromeDriver driver, string xIuput, string url)//在Chrome瀏覽器的文字框中輸入文字,creedit
        {

            //selm.IWebDriver driver = new ChromeDriver();

            // 使用driver導航到給定的URL
            driver.Navigate().GoToUrl(url);
            //("https://ctext.org/library.pl?if=en&file=79166&page=85&editwiki=297821#editor");//("http://www.example.com");

            // 查找名稱為"textbox"的文字框元素
            selm.IWebElement textbox = driver.FindElement(selm.By.Name("data"));//("textbox"));

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
            if (isAllinBmp(xIuput))
            {
                //textbox.SendKeys(stringtoEscape_sequences_for_Unicode_character_sets(xIuput));//(Keys.Control + "v");            
                textbox.SendKeys(xIuput);
            }
            //若含BMP外的字則用系統貼上的方法
            else
            {
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
            }
            //Task.WaitAll();
            //System.Windows.Forms.Application.DoEvents();
            //送出
            selm.IWebElement submit = driver.FindElement(selm.By.Id("savechangesbutton"));//("textbox"));
            submit.Click();
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
            driver.Navigate().GoToUrl(url);
            //activate and move to most front of desktop
            driver.SwitchTo().Window(driver.CurrentWindowHandle);
            driver.ExecuteScript("window.scrollTo(0, 0)");//chatGPT:您好！如果您使用 C# 和 Selenium 來控制 Chrome 瀏覽器，您可以使用以下的程式碼將網頁捲到最上面：
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
    }
}

