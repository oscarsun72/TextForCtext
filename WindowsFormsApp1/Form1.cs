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

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
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
            //string x = textBox1.Text;
            if (selStart == textBox1.Text.Length) selStart = 0;
            string xPre = textBox1.Text.Substring(0, selStart);
            string x = textBox1.Text.Substring(selStart);
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
            textBox1.Text = xPre + resltTxt;
            textBox1.Focus();
            textBox1.SelectionStart = selStart;
            textBox1.SelectionLength = 0;
            textBox1.ScrollToCaret();
            //Clipboard.SetText(resltTxt);
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text=="")
            {
                textBox1.Text = Clipboard.GetText();
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_Leave(object sender, EventArgs e)
        {
            if (textBox2.Text == "") return;
            string x = textBox1.Text; int xStart = x.IndexOf(textBox2.Text);
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

        private void textBox2_Enter(object sender, EventArgs e)
        {
            textBox2.Text = "";
            newTextBox1();
        }

        private void newTextBox1()
        {
            if (textBox1.Text == "") return;
            if (textBox1.SelectedText != "")
            {
                textBox2.Text = "";
                string x = textBox1.Text;
                Clipboard.SetText(x.Substring(0,textBox1.SelectionStart+textBox1.SelectionLength));
                pasteToCtext();
                x = x.Substring(textBox1.SelectionStart + textBox1.SelectionLength + 2);
                textBox1.Text = x;
            }
            if (textBox1.Text.Substring(0, 2) == Environment.NewLine) textBox1.Text = textBox1.Text.Substring(2);
            textBox1.SelectionStart = 0; textBox1.SelectionLength = 0;
            textBox1.ScrollToCaret();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {

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
            var m = ModifierKeys;
            if ((m & Keys.Control) == Keys.Control
                && (m & Keys.Shift) == Keys.Shift
                && e.KeyCode == Keys.C)
            {
                Clipboard.SetText(textBox1.Text);
            }
            if (Control.ModifierKeys == Keys.Control)
            {
                if (e.KeyCode==Keys.S)
                {
                    savetext();
                }
                if (e.KeyCode==Keys.Q)
                {
                    splitLineByFristLen();
                }
                if (e.KeyCode==Keys.OemBackslash||e.KeyCode==Keys.Packet||e.KeyCode==Keys.Oem5)
                {
                    string x = textBox1.Text;
                    int s = textBox1.SelectionStart;
                    string xNext = x.Substring(s);
                    x = x.Substring(0,textBox1.SelectionStart);
                    xNext= xNext.Replace(Environment.NewLine, "");
                    x = x + xNext;
                    textBox1.Text = x;
                    textBox1.SelectionStart = s;textBox1.SelectionLength = 1;
                    textBox1.ScrollToCaret();
                }
                if (e.KeyCode==Keys.Oem7)
                {
                    replaceWord();
                }
                if (e.KeyCode == Keys.NumPad5 || e.KeyCode == Keys.Oemplus || e.KeyCode == Keys.Add)
                {
                    newTextBox1();
                    pasteToCtext();

                }

                if (e.KeyCode == Keys.PageDown || e.KeyCode == Keys.PageUp)
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
                        if (e.KeyCode == Keys.PageDown)
                            url = urlSub + (page + 1).ToString() + url.Substring(url.IndexOf("&editwiki="));
                        if (e.KeyCode == Keys.PageUp)
                            url = urlSub + (page - 1).ToString() + url.Substring(url.IndexOf("&editwiki="));
                        newTextBox1();
                    }
                    else
                    {
                        urlSub = url.Substring(0, url.IndexOf("&page=") + "&page=".Length);
                        page = Int32.Parse(url.Substring(url.IndexOf("&page=") + "&page=".Length));
                        if (e.KeyCode == Keys.PageDown)
                            url = urlSub + (page + 1).ToString();
                        if (e.KeyCode == Keys.PageUp)
                            url = urlSub + (page - 1).ToString();
                    }
                    Process.Start(url);
                    appActivateByID();
                    Task.Delay(1500).Wait();
                    SendKeys.Send("{Tab}"); //("{Tab 24}");
                    Task.Delay(500).Wait();
                    SendKeys.Send("^a");
                    textBox3.Text = url;
                }
            }
        }

        private void savetext()
        {
            string str1 = textBox1.Text;            
            File.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)+@"\Dropbox\cText.txt", str1, Encoding.UTF8); //https://codertw.com/%E7%A8%8B%E5%BC%8F%E8%AA%9E%E8%A8%80/542361/
            // 也可以指定編碼方式 File.WriteAllText(@”c:\temp\test\ascii-2.txt”, str1, Encoding.ASCII);
            //throw new NotImplementedException();
        }

        string processID;
        //https://stackoverflow.com/questions/58302052/c-microsoft-visualbasic-interaction-appactivate-no-effect
        [DllImport("user32.dll", SetLastError = true)]
        static extern void SwitchToThisWindow(IntPtr hWnd, bool turnOn);
        void appActivateByID()
        { //https://docs.microsoft.com/zh-tw/dotnet/csharp/programming-guide/strings/how-to-determine-whether-a-string-represents-a-numeric-value
            int i = 0;
            if (processID == null||processID=="")
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
            appActivateByID();
            SendKeys.Send("^v{tab}~");
            //throw new NotImplementedException();
        }

        
        private void button2_Click(object sender, EventArgs e)
        {
            pasteToCtext();
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            textBox1.Height = this.Height - textBox2.Height*3 - textBox2.Top;
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            replaceWord();
        }

        private void replaceWord()
        {
            if (textBox4.Text == "") return;
                if (!(textBox1.SelectionLength > 2 || textBox1.SelectionLength == 0))
            {
                int s = textBox1.SelectionStart; int l = textBox1.SelectionLength;
                textBox1.Text = textBox1.Text.Replace(textBox1.SelectedText, textBox4.Text);
            }
            //throw new NotImplementedException();
        }
    }
}
