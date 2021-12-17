using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using System.Runtime.InteropServices;
using System.IO;
using System.Drawing.Text;
using System.Media;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        readonly Point textBox4Location; readonly Size textBox4Size;
        readonly string dropBoxPathIncldBackSlash;
        const string CJKBiggestSet = "KaiXinSongB";//"TH-Tshyn-P1";//"HanaMinB";
        Color button2BackColor;
        //bool insertMode = true;
        public Form1()
        {
            InitializeComponent();
            textBox4Location = textBox4.Location;
            textBox4Size = textBox4.Size;
            dropBoxPathIncldBackSlash = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Dropbox\";
            button2BackColor = button2.BackColor;
            var cjk = getHanaminBFontInstalled();
            if (cjk != null)
            {
                if (cjk.Name == "KaiXinSongB")
                {
                    textBox1.Font = new Font(cjk, (float)17);
                }
                else
                {
                    textBox1.Font = new Font(cjk, textBox1.Font.Size);
                }
                textBox2.Font = new Font(cjk, textBox2.Font.Size);
                textBox4.Font = new Font(cjk, textBox4.Font.Size);
            }
        }
        FontFamily getHanaminBFontInstalled()
        { //https://www.cnblogs.com/arxive/p/7795232.html            
            InstalledFontCollection MyFont = new InstalledFontCollection();
            FontFamily[] fontFamilys = MyFont.Families;
            if (fontFamilys == null || fontFamilys.Length < 1)
            {
                return null;
            }
            foreach (FontFamily item in fontFamilys)
            {
                if (item.Name == CJKBiggestSet) return item;
            }
            return null;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            splitLineByFristLen();
        }

        private void splitLineByFristLen()
        {
            //據第一行長度來分行分段
            bool noteFlg = false;
            int selStart = textBox1.SelectionStart;
            string x = "";
            if (selStart == textBox1.Text.Length) selStart = 0;
            if (selStart != 0)
            {
                selToNewline(ref selStart, ref selStart, textBox1.Text, false, textBox1);
            }
            string xPre = textBox1.Text.Substring(0, selStart);
            x = textBox1.Text.Substring(selStart);
            if (x == "") Clipboard.GetText();
            const string omitStr = "{}<p>《》〈〉：，。「」『』　0123456789-‧·\r\n";
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
                //if(mystr=="《"||mystr=="〈")
                //{
                //    break;
                //}

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
            int lineLen = 0;
            if (wordCntr == 0 && noteFlg)//純注文
                lineLen = noteCtr;
            else
                lineLen = wordCntr + noteCtr / 2;//wordCntr+((int)Math.Round(noteCtr/2.0));
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
            textBox1.Text = xPre + resltTxt.Replace("}" + Environment.NewLine + "}", "}}" + Environment.NewLine)
                .Replace(Environment.NewLine + "》", "》" + Environment.NewLine)
                .Replace(Environment.NewLine + "〉", "〉" + Environment.NewLine)
                .Replace(Environment.NewLine + "}}", "}}" + Environment.NewLine)
                .Replace("{{" + Environment.NewLine, Environment.NewLine + "{{");
            textBox1.Focus();
            textBox1.SelectionStart = selStart;
            textBox1.SelectionLength = 0;
            textBox1.ScrollToCaret();
            //Clipboard.SetText(resltTxt);
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                textBox1.Text = Clipboard.GetText();
            }
        }


        private void textBox2_Leave(object sender, EventArgs e)
        {
            string s = textBox2.Text;
            if (s == "") return;
            //如何判斷字串是否代表數值 (c # 程式設計手冊):https://docs.microsoft.com/zh-tw/dotnet/csharp/programming-guide/strings/how-to-determine-whether-a-string-represents-a-numeric-value
            int i = 0;
            bool result = int.TryParse(s, out i); //i now = textBox2.Text
            if (result && (processID == null || processID == ""))
            {
                processID = s;
            }
            string x = textBox1.Text; int xStart = x.IndexOf(s);
            if (xStart > 0)
            {
                textBox1.Focus();
                textBox1.Select(xStart, textBox2.Text.Length);
                textBox1.ScrollToCaret();
                x = x.Substring(0, xStart + textBox2.Text.Length);
                Clipboard.SetText(x);
                //textBox1.Text = textBox1.Text.Substring(xStart + 2);
            }

        }


        private void newTextBox1()
        {
            if (textBox1.Text == "") return;
            saveText();
            //if (textBox1.SelectedText != "")
            //{
            textBox2.Text = "";
            string x = textBox1.Text;
            int s = textBox1.SelectionStart, l = textBox1.SelectionLength;
            string xCopy = x.Substring(0, s + l);
            Clipboard.SetText(xCopy); backupLastPageText(xCopy, false);
            if (xCopy.IndexOf(" ") > -1 || xCopy.IndexOfAny("�".ToCharArray()) > -1)
            {//  「�」甚特別，indexof會失效，明明沒有，而傳回 0 //https://docs.microsoft.com/zh-tw/dotnet/csharp/how-to/compare-strings
             //  //https://docs.microsoft.com/zh-tw/dotnet/api/system.string.compare?view=net-6.0
                SystemSounds.Exclamation.Play();//文本有缺字警告
                Color c = this.BackColor;
                this.BackColor = Color.Yellow;
                Task.Delay(900).Wait();
                this.BackColor = c;
            }
            if (s + l + 2 < textBox1.Text.Length)
            {
                if (x.Substring(s + l, 1) == Environment.NewLine)
                    x = x.Substring(s + l + 2);
                else
                    x = x.Substring(s + l);
            }
            if (s + l >= textBox1.Text.Length)
            {
                x = "";
            }
            textBox1.Text = x;
            //}
            if (textBox1.Text.Length > 1)
            {
                if (textBox1.Text.Substring(0, 2) == Environment.NewLine) textBox1.Text = textBox1.Text.Substring(2);
            }
            textBox1.SelectionStart = 0; textBox1.SelectionLength = 0;
            textBox1.ScrollToCaret();
        }


        string textBox1OriginalText;


        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {

            var m = ModifierKeys;

            if ((m & Keys.Control) == Keys.Control
            && (m & Keys.Shift) == Keys.Shift
            && e.KeyCode == Keys.Up)
            {
                int s = textBox1.SelectionStart, ed = s;
                selToNewline(ref s, ref ed, textBox1.Text, false, textBox1);
                return;
            }
            if ((m & Keys.Control) == Keys.Control
                && (m & Keys.Shift) == Keys.Shift
                && e.KeyCode == Keys.Down)
            {
                int s = textBox1.SelectionStart, ed = s;
                selToNewline(ref s, ref ed, textBox1.Text, true, textBox1);
                return;
            }

            if ((m & Keys.Control) == Keys.Control)
            {//按下Ctrl鍵
                if (e.KeyCode == Keys.F12)
                {
                    string x = textBox1.SelectedText;
                    if (x != "")
                    {
                        Clipboard.SetText(x);
                        Process.Start(dropBoxPathIncldBackSlash + @"VS\VB\查詢國語辭典\查詢國語辭典\bin\Debug\查詢國語辭典.exe");
                    }
                    return;
                }
                if (e.KeyCode == Keys.NumPad5 || e.KeyCode == Keys.Oemplus || e.KeyCode == Keys.Add || e.KeyCode == Keys.Subtract)
                {
                    newTextBox1();
                    pasteToCtext();
                    nextPages(Keys.PageDown);
                    return;
                }
                if (e.KeyCode == Keys.D0 || e.KeyCode == Keys.D9 || e.KeyCode == Keys.D8 || e.KeyCode == Keys.D7)
                {
                    int s = textBox1.SelectionStart, l = textBox1.SelectionLength; string insX = "", x = textBox1.Text;
                    if (textBox1.SelectedText != "")
                        x = x.Substring(0, s) + x.Substring(s + l);
                    if (e.KeyCode == Keys.D0)
                    {
                        insX = Environment.NewLine + "　" + Environment.NewLine +
                            "　" + Environment.NewLine +
                            "　" + Environment.NewLine +
                            "　" + Environment.NewLine;
                    }
                    if (e.KeyCode == Keys.D9)
                    {
                        insX = Environment.NewLine + "　" + Environment.NewLine +
                            "　" + Environment.NewLine;
                    }
                    if (e.KeyCode == Keys.D8)
                    {
                        insX = Environment.NewLine + "　" + Environment.NewLine;
                    }
                    if (e.KeyCode == Keys.D7)
                    {
                        insX = "。}}";
                    }
                    x = x.Substring(0, s) + insX + x.Substring(s);
                    textBox1.Text = x;
                    textBox1.SelectionStart = s + insX.Length;
                    textBox1.ScrollToCaret();
                    return;
                }

                if (e.KeyCode == Keys.H)
                //if ((m & Keys.Control) == Keys.Control && e.KeyCode == Keys.H)
                {
                    //不知為何，就是會將插入點前一個字元給刪除,即使有以下此行也無效
                    e.Handled = true;
                    textBox1OriginalText = textBox1.Text; selLength = textBox1.SelectionLength; selStart = textBox1.SelectionStart;
                    textBox4.Focus();
                    return;
                }

                if (e.KeyCode == Keys.Q)
                {
                    splitLineByFristLen(); return;
                }

                if (e.KeyCode == Keys.OemBackslash || e.KeyCode == Keys.Packet || e.KeyCode == Keys.Oem5)
                {//clear the newline after the caret
                    string x = textBox1.Text;
                    int s = textBox1.SelectionStart;
                    string xNext = x.Substring(s);
                    x = x.Substring(0, textBox1.SelectionStart);
                    xNext = xNext.Replace(Environment.NewLine, "");
                    x = x + xNext;
                    textBox1.Text = x;
                    textBox1.SelectionStart = s; textBox1.SelectionLength = 1;
                    textBox1.ScrollToCaret();
                    return;
                }

            }

            //按下Shift鍵
            if ((m & Keys.Shift) == Keys.Shift)
            {
                if (e.KeyCode == Keys.F3)
                {
                    e.Handled = true;
                    int foundwhere;
                    string findword = textBox1.SelectedText;
                    if (findword == "") findword = textBox2.Text;
                    if (findword != "")
                    {
                        int start = textBox1.SelectionStart - 1; string x = textBox1.Text;
                        foundwhere = x.LastIndexOf(findword, start);
                        if (foundwhere == -1)
                        {
                            MessageBox.Show("not found next!"); return;
                        }
                        textBox1.SelectionStart = foundwhere;
                        textBox1.SelectionLength = findword.Length;
                        textBox1.ScrollToCaret();
                    }
                    return;
                }

            }

            //按下Alt鍵
            if ((m & Keys.Alt) == Keys.Alt)//⇌ if (Control.ModifierKeys == Keys.Alt)
            {
                if (e.KeyCode == Keys.G)
                {
                    string x = textBox1.SelectedText;
                    if (x != "")
                    {
                        Clipboard.SetText(x);
                        Process.Start(dropBoxPathIncldBackSlash + @"VS\VB\網路搜尋_元搜尋-同時搜多個引擎\網路搜尋_元搜尋-同時搜多個引擎\bin\Debug\網路搜尋_元搜尋-同時搜多個引擎.exe");
                    }
                    return;
                }
                if (e.KeyCode == Keys.Q)
                {
                    splitLineByFristLen(); return;
                }


            }

            //按下單一鍵
            //if (e.KeyCode == Keys.Insert)
            //{
            //    if (insertMode) insertMode = false;
            //    else insertMode = true;
            //}
            if (e.KeyCode == Keys.F2)
            {
                keyDownF2(textBox1); return;
            }
            if (e.KeyCode == Keys.F3)
            {
                e.Handled = true;
                int foundwhere;
                string findword = textBox1.SelectedText;
                if (findword == "") findword = textBox2.Text;
                if (findword != "")
                {
                    int start = textBox1.SelectionStart + 1; string x = textBox1.Text;
                    foundwhere = x.IndexOf(findword, start);
                    if (foundwhere == -1)
                    {
                        MessageBox.Show("not found next!"); return;
                    }
                    textBox1.SelectionStart = foundwhere;
                    textBox1.SelectionLength = findword.Length; textBox1.ScrollToCaret();
                }
                return;
            }

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

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_Click(object sender, EventArgs e)
        {
            string x = Clipboard.GetText();
            if (x.Substring(0, 4) == "http")
                if (x.IndexOf("ctext.org") > -1)
                {
                    textBox3.Text = x;
                }

        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            //if ((int)ModifierKeys == (int)Keys.Control+ (int)Keys.Shift&&e.KeyCode==Keys.C)
            //https://bbs.csdn.net/topics/350010591
            //https://zhidao.baidu.com/question/628222381668604284.html
            var m = ModifierKeys; Color C = this.BackColor;
            if ((m & Keys.Control) == Keys.Control
                && (m & Keys.Shift) == Keys.Shift
                && e.KeyCode == Keys.C)
            {
                Clipboard.SetText(textBox1.Text);
            }

            if (Control.ModifierKeys == Keys.Control)
            {//按下Ctrl鍵
                if (e.KeyCode == Keys.F)
                {
                    textBox2.Focus();
                    textBox2.SelectionStart = 0; textBox2.SelectionLength = textBox2.Text.Length; return;
                }
                
                this.BackColor = Color.Green;
                if (e.KeyCode == Keys.D1)
                {
                    runWord("漢籍電子文獻資料庫文本整理_以轉貼到中國哲學書電子化計劃"); return;
                }
                if (e.KeyCode == Keys.D3)
                {
                    runWord("漢籍電子文獻資料庫文本整理_十三經注疏"); return;
                }
                if (e.KeyCode == Keys.D4)
                {
                    runWord("維基文庫四部叢刊本轉來"); return;
                }
                if (e.KeyCode == Keys.S)
                {
                    saveText(); return;
                }
                this.BackColor = C;


                if (e.KeyCode == Keys.PageDown || e.KeyCode == Keys.PageUp)
                {

                    e.Handled = true;//取得或設定值，指出是否處理事件。https://docs.microsoft.com/zh-tw/dotnet/api/system.windows.forms.keyeventargs.handled?view=netframework-4.7.2&f1url=%3FappId%3DDev16IDEF1%26l%3DZH-TW%26k%3Dk(System.Windows.Forms.KeyEventArgs.Handled);k(TargetFrameworkMoniker-.NETFramework,Version%253Dv4.7.2);k(DevLang-csharp)%26rd%3Dtrue
                    nextPages(e.KeyCode);
                    return;
                }
            }
            if (e.KeyCode == Keys.F5)
            {
                loadText();
                return;
            }
            if (e.KeyCode == Keys.F12)
            {
                Color c = this.BackColor;
                this.BackColor = Color.Red;
                e.Handled = true;
                backupLastPageText(Clipboard.GetText(), true);
                Task.Delay(800).Wait();
                this.BackColor = c;
                return;
            }
        }

        const string fName_to_Backup_Txt = "cTextBK.txt";
        void backupLastPageText(string x, bool updateLastBackup)
        {
            //C# 對文字檔案的幾種讀寫方法總結:https://codertw.com/%E7%A8%8B%E5%BC%8F%E8%AA%9E%E8%A8%80/542361/
            string lastPageText = x + "＠"; //"＠" 作為每頁的界號
            if (File.Exists(dropBoxPathIncldBackSlash + fName_to_Backup_Txt))
            {
                if (updateLastBackup)
                {

                    string bk = File.ReadAllText(dropBoxPathIncldBackSlash + fName_to_Backup_Txt);
                    int bkLastEnd = bk.LastIndexOf("＠"), bkLastStart = bk.LastIndexOf("＠", bkLastEnd - 1) + 1;
                    //if (bkLastStart == -1) bkLastStart = 0;
                    bk = bk.Substring(0, bkLastStart) + lastPageText;
                    File.WriteAllText(dropBoxPathIncldBackSlash + fName_to_Backup_Txt, bk, Encoding.UTF8);
                    return;
                }
            }
            File.AppendAllText(dropBoxPathIncldBackSlash + fName_to_Backup_Txt, lastPageText, Encoding.UTF8);

        }

        private void nextPages(Keys eKeyCode)
        {
            string url = textBox3.Text;
            if (url == "") return;
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
            Process.Start(url);
            appActivateByName();
            if (edit > -1)
            {//編輯才執行，瀏覽則省略
                Task.Delay(1900).Wait();
                SendKeys.Send("{Tab}"); //("{Tab 24}");
                Task.Delay(200).Wait();
                SendKeys.Send("^a");
                Task.Delay(500).Wait();
                SendKeys.Send("^{PGUP}");//回上一頁籤檢查文本是否如願貼好
            }
            textBox3.Text = url;
        }

        private void runWord(string runName)
        {
            SystemSounds.Hand.Play();
            Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office
                                    .Interop.Word.Application();
            appWord.Run(runName);
            textBox1.Text = Clipboard.GetText();
            appWord.Quit(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);
        }

        const string fName_to_Save_Txt = "cText.txt";
        private void saveText()
        {
            //C# 對文字檔案的幾種讀寫方法總結:https://codertw.com/%E7%A8%8B%E5%BC%8F%E8%AA%9E%E8%A8%80/542361/
            string str1 = textBox1.Text;
            File.WriteAllText(dropBoxPathIncldBackSlash + fName_to_Save_Txt, str1, Encoding.UTF8);
            // 也可以指定編碼方式 File.WriteAllText(@”c:\temp\test\ascii-2.txt”, str1, Encoding.ASCII);
        }

        private void loadText()
        {
            //C# 對文字檔案的幾種讀寫方法總結:https://codertw.com/%E7%A8%8B%E5%BC%8F%E8%AA%9E%E8%A8%80/542361/
            textBox1.Text = File.ReadAllText(dropBoxPathIncldBackSlash + fName_to_Save_Txt);
        }




        string processID;
        //https://stackoverflow.com/questions/58302052/c-microsoft-visualbasic-interaction-appactivate-no-effect
        [DllImport("user32.dll", SetLastError = true)]
        static extern void SwitchToThisWindow(IntPtr hWnd, bool turnOn);

        void appActivateByName()
        {
            Process[] procsBrowser = Process.GetProcessesByName("chrome");
            if (procsBrowser.Length <= 0)
            {
                MessageBox.Show("Chrome is not running");
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



        private void pasteToCtext()
        {
            appActivateByName();
            Task.Delay(100).Wait();
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
                button2.BackColor = button2BackColor;
            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            textBox1.Height = this.Height - textBox2.Height * 3 - textBox2.Top;
        }

        const int CJK_Crtr_Len_Max = 2;//因為目前CJK最長為2字元

        private void replaceWord(string replacedword, string rplsword)
        {
            if (rplsword == "") return;
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
            if (button2.Text == "選取文")
            {
                replacedword = textBox2.Text;
                if (replacedword == "") return;
                l = textBox1.SelectionLength;
                string xBefore = x.Substring(0, s), xAfter = x.Substring(s + l);
                x = textBox1.SelectedText;
                textBox1.Text = xBefore + x.Replace(replacedword, rplsword) + xAfter;
            }
            else
            {
                l = selWord.LengthInTextElements;
                textBox1.Text = x.Replace(replacedword, rplsword);
            }
            addReplaceWordDefault(replacedword, rplsword);
            textBox1.SelectionStart = s; textBox1.SelectionLength = l;
            textBox1.ScrollToCaret();
            textBox1.Focus();
        }

        List<string> replaceWordList = new List<string>();
        List<string> replacedWordList = new List<string>();

        string getReplaceWordDefault(string replacedWord)
        {
            if (replacedWordList.Count == 0) return "";
            string replsWord = "";
            for (int i = 0; i < replacedWordList.Count; i++)
            {
                if (replacedWord == replacedWordList[i])
                {
                    replsWord += replaceWordList[i];
                }
            }
            return replsWord;

        }
        void addReplaceWordDefault(string replacedWord,
                string replaceWord)
        {
            if (replacedWordList.Contains(replacedWord))
            {
                int i = 0, count = replacedWordList.Count;
                while (i < count)
                {
                    if (replacedWordList.IndexOf(replacedWord, i) == replaceWordList.IndexOf(replaceWord, i))
                    {
                        return;
                    }
                    i++;
                }
            }
            replacedWordList.Add(replacedWord);
            replaceWordList.Add(replaceWord);

        }

        private void textBox4_Leave(object sender, EventArgs e)
        {
            textBox1.Focus();
            saveText();
            replaceWord(textBox1.SelectedText, textBox4.Text);
            textBox4Resize();
            textBox4.Text = "";
        }

        private void textBox4Resize()
        {
            textBox4.Location = textBox4Location;
            textBox4.Size = textBox4Size;
            textBox4.ScrollBars = ScrollBars.None;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //C# 如何取得使用者的螢幕解析度:https://blog.xuite.net/q10814/blog/48070595 https://www.delftstack.com/zh-tw/howto/csharp/screen-size-in-csharp/
            Size Size = SystemInformation.PrimaryMonitorSize;
            int Width = SystemInformation.PrimaryMonitorSize.Width;
            int Height = SystemInformation.PrimaryMonitorSize.Height;
            //MessageBox.Show("你的螢幕解析度是" + Size + "\n Width = " + Width + "\n Height = " + Height);
            //FormStartPosition 列舉:https://docs.microsoft.com/zh-tw/dotnet/api/system.windows.forms.formstartposition?view=netframework-4.7.2
            this.Location = new Point(Width - this.Width, Height - textBox1.Height * 2 + 150);
            //this.PointToScreen();
        }

        private void textBox4_Enter(object sender, EventArgs e)
        {
            if (textBox4.Size == textBox4Size)
                textBox4SizeLarger();
            string rplsdWord = "";
            rplsdWord = textBox1.SelectedText;
            if (rplsdWord != "")
            {
                string rplsWord = getReplaceWordDefault(rplsdWord);
                if (rplsWord != "") textBox4.Text = rplsWord;
            }
        }

        private void textBox4SizeLarger()
        {
            textBox4.Location = new Point(button1.Location.X, textBox4Location.Y);
            textBox4.Size = new Size(textBox2.Size.Width + textBox2.Size.Width +
                                        textBox3.Width + textBox4Size.Width, textBox4Size.Height);
            textBox4.ScrollBars = ScrollBars.Horizontal;
        }

        private void textBox4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F2)
            {
                keyDownF2(textBox4);
            }
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F2)
            {
                keyDownF2(textBox2);
            }
        }

        private void keyDownF2(TextBox textBox)
        {
            if (textBox.Text != "")
            {
                if (textBox.SelectedText != "")
                    textBox.SelectionStart = textBox.Text.Length;
                else
                {
                    textBox.SelectionStart = 0; textBox.SelectionLength = textBox.Text.Length;
                }
            }
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            //textBox2.Text = "";
        }

        private void textBox1_MouseDown(object sender, MouseEventArgs e)
        {
            var m = ModifierKeys;
            if ((m & Keys.Control) == Keys.Control && (m & Keys.Shift) == Keys.Shift)
            {
                runWord("漢籍電子文獻資料庫文本整理_十三經注疏");
            }
            if ((m & Keys.Alt) == Keys.Alt)
            {
                runWord("維基文庫四部叢刊本轉來");
            }
            if (ModifierKeys == Keys.Control && e.Button == MouseButtons.Left)
            {
                richTextBox1.Size = textBox1.Size;
                richTextBox1.Location = textBox1.Location;
                richTextBox1.Show();

            }
        }

        private void richTextBox1_MouseDown(object sender, MouseEventArgs e)
        {
            richTextBox1.Size = textBox1.Size;
            richTextBox1.Location = textBox1.Location;
            richTextBox1.Visible = false;
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
        int selStart = 0; int selLength = 0;
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            undoTextValueChanged(selStart, selLength);
            //if (!insertMode && textBox1.Focused && textBox1.SelectedText=="")
            //{
            //    string x = textBox1.Text;int s = textBox1.SelectionStart;
            //    string xNext = x.Substring(s);
            //    StringInfo xInfo = new StringInfo(xNext);

            //}

        }

        private void undoTextValueChanged(int s, int l)
        {
            if (textBox1OriginalText != "" &&
                                textBox1.Text != textBox1OriginalText)
            {
                textBox1.Text = textBox1OriginalText;
                textBox1OriginalText = "";
                textBox1.SelectionStart = s; textBox1.SelectionLength = l;
            }
        }
    }
}
