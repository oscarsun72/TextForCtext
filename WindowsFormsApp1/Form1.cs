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
using System.Text.RegularExpressions;
using System.Web;/*c# how to get the focused control on Chrome
                  * *https://docs.microsoft.com/en-us/dotnet/api/system.web.ui.page.setfocus?view=netframework-4.8 */

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        readonly Point textBox4Location; readonly Size textBox4Size;
        readonly string dropBoxPathIncldBackSlash;
        const string CJKBiggestSet = "HanaMinB";//"KaiXinSongB";//"TH-Tshyn-P1";
        Color button2BackColorDefault;
        bool insertMode = true;

        System.Windows.Forms.NotifyIcon nICo;
        int thisHeight, thisWidth, thisLeft, thisTop;
        [DllImport("user32.dll")]
        static extern bool CreateCaret(IntPtr hWnd, IntPtr hBitmap, int nWidth, int nHeight);
        [DllImport("user32.dll")]
        static extern bool ShowCaret(IntPtr hWnd);
        public Form1()
        {
            InitializeComponent();
            textBox4Location = textBox4.Location;
            textBox4Size = textBox4.Size;
            dropBoxPathIncldBackSlash = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Dropbox\";
            button2BackColorDefault = button2.BackColor;
            textBox2BackColorDefault = textBox2.BackColor;
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
            thisHeight = this.Height; thisWidth = this.Width; thisLeft = this.Left; thisTop = this.Top;
            this.nICo = new NotifyIcon();
            this.nICo.Icon = this.Icon;
            this.nICo.MouseClick += new System.Windows.Forms.MouseEventHandler(nICo_MouseClick);
            this.nICo.MouseMove += new System.Windows.Forms.MouseEventHandler(nICo_MouseMove);
            //this.Shown += Form1_Shown;//https://stackoverflow.com/questions/32720207/change-caret-cursor-in-textbox-in-c-sharp
        }

        void Caret_Shown(Control ctl)
        {
            CreateCaret(ctl.Handle, IntPtr.Zero, 4, Convert.ToInt32(ctl.Font.Size * 1.5));
            ShowCaret(ctl.Handle);
        }
        void Caret_Shown_OverwriteMode(Control ctl)
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
        void show_nICo()
        {
            nICo.Visible = false;
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.Height = thisHeight;
            this.Width = thisWidth;
            this.Left = thisLeft;
            this.Top = thisTop;
        }

        private void nICo_MouseClick(object sender, MouseEventArgs e)
        {
            show_nICo();
        }

        private void nICo_MouseMove(object sender, MouseEventArgs e)
        {
            if (Cursor.Position.Y > this.Top + this.Height) show_nICo();
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
            int lineLen = 0;//taked the normal Line and or Para Length
            if (wordCntr == 0 && noteFlg)//純注文
                lineLen = noteCtr;
            else
                lineLen = wordCntr + noteCtr / 2;//wordCntr+((int)Math.Round(noteCtr/2.0));
            normalLineParaLength = lineLen;
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
            textBox1.SelectionStart = s + resltTxt.IndexOf(Environment.NewLine) + Environment.NewLine.Length;//selStart;
            textBox1.SelectionLength = 0;
            textBox1.ScrollToCaret();
            //Clipboard.SetText(resltTxt);
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                textBox1.Text = Clipboard.GetText();
                textBox1.Select(0, 0);
                textBox1.ScrollToCaret();
            }
        }


        private void textBox2_Leave(object sender, EventArgs e)
        {
            string s = textBox2.Text;
            if (s == "" || s == textBox1.SelectedText)
            { textBox2.BackColor = textBox2BackColorDefault; return; }
            //如何判斷字串是否代表數值 (c # 程式設計手冊):https://docs.microsoft.com/zh-tw/dotnet/csharp/programming-guide/strings/how-to-determine-whether-a-string-represents-a-numeric-value
            int i = 0;
            bool result = int.TryParse(s, out i); //i now = textBox2.Text
            if (result && (processID == null || processID == ""))
            {
                processID = s;
            }
            string x = textBox1.Text; int xStart = x.IndexOf(s), nextStart = x.IndexOf(s, xStart + 1);
            Color C = textBox2.BackColor;
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
            string[] replacedChar = { ",", ";", ":" };
            string[] replaceChar = { "，", "；", "：" };
            foreach (var item in replacedChar)
            {
                if (xCopy.IndexOf(item) > -1)
                {
                    if (MessageBox.Show("含半形標點，是否取代為全形？", "", MessageBoxButtons.OKCancel) == DialogResult.OK)
                    {
                        for (int i = 0; i < replaceChar.Length; i++)
                        {
                            xCopy = xCopy.Replace(replacedChar[i], replaceChar[i]);
                        }
                    }
                    break;
                }
            }
            if (xCopy.IndexOf(" ") > -1 || xCopy.IndexOfAny("�".ToCharArray()) > -1 ||
                xCopy.IndexOf("□") > -1)//□為《維基文庫》《四庫全書》的缺字符，" "則是《四部叢刊》的，"�"則是《四部叢刊》的造字符。
            {//  「�」甚特別，indexof會失效，明明沒有，而傳回 0 //https://docs.microsoft.com/zh-tw/dotnet/csharp/how-to/compare-strings
             //  //https://docs.microsoft.com/zh-tw/dotnet/api/system.string.compare?view=net-6.0
                SystemSounds.Hand.Play();//文本有缺字警告
                Color c = this.BackColor;
                this.BackColor = Color.Yellow;
                Task.Delay(900).Wait();
                this.BackColor = c;
                string[] rTxt = { " ", "�", "□" };
                foreach (string rs in rTxt)
                {
                    xCopy = xCopy.Replace(rs, "●");//「●」為《中國哲學書電子化計劃》的缺字符，詳：https://ctext.org/instructions/wiki-formatting/zh
                }
            }
            Clipboard.SetText(xCopy); BackupLastPageText(xCopy, false, false);
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
            //textBox1.Text = x;
            //}
            if (x.Length > 1)
            {
                //清除不需要的部分
                if (x.Substring(0, 2) == Environment.NewLine) x = x.Substring(2);
                string[] rTxt = { "。}}<p>" };
                int lng = rTxt[0].Length;
                string xClr = x.Substring(0, lng);
                int clr = xClr.IndexOf(rTxt[0]);
                while (clr > -1 && x != "")
                {
                    x = x.Substring(clr + lng);
                    xClr = x.Substring(0, lng);
                    clr = xClr.IndexOf(rTxt[0]);
                }
            }
            textBox1.Text = x;
            textBox1.SelectionStart = 0; textBox1.SelectionLength = 0;
            textBox1.ScrollToCaret();
        }


        string textBox1OriginalText = "";


        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {

            var m = ModifierKeys;

            if ((m & Keys.None) == Keys.None && e.KeyCode == Keys.Delete) undoRecord();
            //if ((m & Keys.Control) == Keys.Control && (m & Keys.Alt) == Keys.Alt && e.KeyCode == Keys.G)
            //if((int)Control.ModifierKeys ==
            //    (int)Keys.Control + (int)Keys.Alt && e.KeyCode == Keys.G)
            if ((m & Keys.Shift) == Keys.Shift && e.KeyCode == Keys.Insert) pasteAllOverWrite = true;
            else pasteAllOverWrite = false;
            if ((m & Keys.Control) == Keys.Control
                && (m & Keys.Alt) == Keys.Alt)//https://zhidao.baidu.com/question/628222381668604284.html
            {//https://bbs.csdn.net/topics/350010591
                if (e.KeyCode == Keys.G || e.KeyCode == Keys.Packet)
                { e.Handled = true; return; }
            }
            if ((m & Keys.Control) == Keys.Control
            && (m & Keys.Shift) == Keys.Shift
            && e.KeyCode == Keys.Delete)
            {//Ctrl + Shift + Delete ： 將選取文字於文本中全部清除
                e.Handled = true;
                //int s = textBox1.SelectionStart;
                undoRecord();
                textBox1.Text = textBox1.Text.Replace(textBox1.SelectedText, "");
                textBox1.SelectionStart = selStart;
                textBox1.ScrollToCaret();
                return;
            }
            if ((m & Keys.Control) == Keys.Control
        && (m & Keys.Shift) == Keys.Shift
        && e.KeyCode == Keys.Up)
            {
                int s = textBox1.SelectionStart, ed = s;
                selToNewline(ref s, ref ed, textBox1.Text, false, textBox1); return;
            }
            if ((m & Keys.Control) == Keys.Control
                && (m & Keys.Shift) == Keys.Shift
                && e.KeyCode == Keys.Down)
            {
                int s = textBox1.SelectionStart, ed = s;
                selToNewline(ref s, ref ed, textBox1.Text, true, textBox1); return;
            }

            if ((m & Keys.Control) == Keys.Control)
            {//按下Ctrl鍵
                if (e.KeyCode == Keys.V) pasteAllOverWrite = false;
                else pasteAllOverWrite = false;

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
                if (e.KeyCode == Keys.Add || e.KeyCode == Keys.Oemplus || e.KeyCode == Keys.Subtract || e.KeyCode == Keys.NumPad5)
                {
                    keyDownCtrlAdd();
                    return;
                }
                if (e.KeyCode == Keys.D0 || e.KeyCode == Keys.D9 || e.KeyCode == Keys.D8 || e.KeyCode == Keys.D7 || e.KeyCode == Keys.D6)
                {
                    e.Handled = true;
                    int s = textBox1.SelectionStart, l = textBox1.SelectionLength;
                    string insX = "", x = textBox1.Text;
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
                    insertWords(insX, x);
                    //x = x.Substring(0, s) + insX + x.Substring(s);
                    //textBox1.Text = x;
                    //textBox1.SelectionStart = s + insX.Length;
                    //textBox1.ScrollToCaret();
                    return;
                }
                if (e.KeyCode == Keys.Z)
                {//還原功能
                    e.Handled = true;
                    undoTextBox(textBox1);
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

                if (e.KeyCode == Keys.OemBackslash || e.KeyCode == Keys.Oem5)
                {//clear the newline after the caret
                    string x = textBox1.Text;
                    int s = textBox1.SelectionStart;
                    string xNext = x.Substring(s);
                    x = x.Substring(0, textBox1.SelectionStart);
                    xNext = xNext.Replace(Environment.NewLine, "");
                    normalLineParaLength = 0;
                    x = x + xNext;
                    textBox1.Text = x;
                    textBox1.SelectionStart = s; textBox1.SelectionLength = 1;
                    textBox1.ScrollToCaret();
                    return;
                }
                if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down)
                {/*Ctrl + ↑：從插入點開始向前移至上一段尾
                  * Ctrl + ↓：從插入點開始向後移至這一段末（無分段則不移動）*/
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
                    return;
                }

                if (e.KeyCode == Keys.OemCloseBrackets || e.KeyCode == Keys.OemOpenBrackets)
                {/*Ctrl + [：從插入點開始向前移至{{前
                    Ctrl + ]：從插入點開始向後移至}}後*/
                    int s = textBox1.SelectionStart; string x = textBox1.Text;
                    if (e.KeyCode == Keys.OemCloseBrackets)
                        s = x.IndexOf("}}", s + 1) + 2;
                    else
                        s = x.LastIndexOf("{{", s - 1);
                    if (s > -1)
                        textBox1.SelectionStart = s;
                    else
                        MessageBox.Show("not found!");
                    textBox1.ScrollToCaret();
                    return;
                }

                if (e.KeyCode == Keys.Left || e.KeyCode == Keys.Right)
                {/*Ctrl + →：插入點若在漢字中,從插入點開始向後移至任何非漢字前(即漢字後) 反之亦然
                  * Ctrl + ←：：插入點若在漢字中,從插入點開始向後移至任何非漢字後(即漢字前) 反之亦然*/
                    string x = textBox1.Text;
                    int s = textBox1.SelectionStart;
                    int l;
                    bool isIPCharHanzi = isChineseChar(x.Substring(s, 1));
                    if (e.KeyCode == Keys.Left)
                    {
                        if (isIPCharHanzi) l = findNotChineseCharFarLength(x.Substring(0, s), false);
                        else l = findChineseCharFarLength(x.Substring(0, s), false);
                        if (l != -1)
                        {
                            textBox1.Select(s - l + 1, 0);
                            textBox1.ScrollToCaret();
                            e.Handled = true;
                            return;
                        }
                    }
                    else
                    {
                        if (isIPCharHanzi) l = findNotChineseCharFarLength(x.Substring(s), true);
                        else l = findChineseCharFarLength(x.Substring(s), true);
                        if (l != -1)
                        {
                            textBox1.Select(s + l - 1, 0);
                            textBox1.ScrollToCaret();
                            e.Handled = true;
                            return;
                        }
                    }
                }

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
                if (e.KeyCode == Keys.D9 || e.KeyCode == Keys.D0 || e.KeyCode == Keys.U || e.KeyCode == Keys.Y || e.KeyCode == Keys.I)
                {/* Alt + 9 : 鍵入 「 
                  * Alt + 0 : 鍵入 『 
                  * Alt + u : 鍵入 《 
                  * Alt + y : 鍵入 〈 
                  * Alt + i : 鍵入 》（如 MS Word 自動校正(如在「選項>印刷樣式」中的設定值)，會依前面的符號作結尾號（close），如前是「〈」，則轉為「〉」……）*/
                    e.Handled = true;
                    string insX = "", x = textBox1.Text;
                    if (e.KeyCode == Keys.D9) { insX = "「"; goto insert; }
                    if (e.KeyCode == Keys.D0) { insX = "『"; goto insert; }
                    if (e.KeyCode == Keys.U) { insX = "《"; goto insert; }
                    if (e.KeyCode == Keys.Y) { insX = "〈"; goto insert; }
                    if (e.KeyCode == Keys.I)
                    {
                        int s = textBox1.SelectionStart;
                        if (s > 0)
                        {
                            string xPrevious = x.Substring(0, s);
                            const string symbol = "{（〈《「『』」》〉）";
                            string whatSymbolPrefix = "";
                            string xChk = ""; bool chk = false; bool closeFlag = false;
                            for (int i = xPrevious.Length - 1; i > -1; i--)
                            {
                                whatSymbolPrefix = xPrevious.Substring(i, 1);
                                if (symbol.IndexOf(whatSymbolPrefix) > -1)
                                {
                                    xChk = xPrevious.Substring(0, i + 1); chk = true;
                                    break;
                                }
                            }
                            if (chk)//需要檢查誰沒配對
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
                                        insX = symbolPairChk[sFirst].ToString();
                                        if (symbolPairChkClose.IndexOf(xChk[i]) == -1)
                                        {//如果是open 
                                            if (sPairOpenFirst.Count == 0 ||
                                                !sPairOpenFirst.Contains(xChk[i].ToString()))
                                            {
                                                insX = symbolPairChkClose[sFirst].ToString();
                                                closeFlag = true;
                                                break;
                                            }
                                        }
                                        else
                                        {//如果是close,取得其配對的 open
                                            string sPOF = symbolPairChk[
                                                symbolPairChkClose.IndexOf(insX)].ToString();
                                            if (sPairOpenFirst.Count == 0 || !sPairOpenFirst.Contains(sPOF))
                                            {
                                                sPairOpenFirst.Add(sPOF);
                                            }
                                            continue;
                                        }

                                    }

                                }//end of for loop 
                                if (!closeFlag)
                                {
                                    insX = "》";
                                }
                            }
                            else
                            {
                                insX = "》";

                            }

                        }
                        else
                        {//pick up the close symbol according to the open one
                            switch (insX)
                            {
                                case "{":
                                    insX = "}}";
                                    break;
                                case "（":
                                    insX = "）";
                                    break;
                                case "〈":
                                    insX = "〉";
                                    break;
                                case "《":
                                    insX = "》";
                                    break;
                                case "「":
                                    insX = "」";
                                    break;
                                case "『":
                                    insX = "』";
                                    break;
                                default:
                                    insX = "》";
                                    break;
                            }


                        }
                    }
                    else
                    {
                        insX = "》";
                    }
                insert:
                    insertWords(insX, x);
                    return;
                }

                if (e.KeyCode == Keys.D1)//D1=Menu?
                {//Alt + 1 : 鍵入本站制式留空空格標記「􏿽」：若有選取則取代全形空格「　」為「􏿽」
                    string x = textBox1.Text;
                    int s = textBox1.SelectionStart, l = textBox1.SelectionLength;
                    string sTxt = textBox1.SelectedText;
                    if (sTxt != "")
                    {
                        string sTxtChk = sTxt.Replace("　", "");
                        if (sTxtChk != "") return;
                        for (int i = 0; i < sTxt.Length; i++)
                        {
                            sTxtChk += "􏿽";
                        }
                        x = x.Substring(0, s) + sTxtChk + x.Substring(s + l);
                        textBox1.Text = x;
                        textBox1.SelectionStart = s + sTxtChk.Length;
                    }
                    else
                    {
                        x = x.Substring(0, s) + "􏿽" + x.Substring(s);
                        textBox1.Text = x;
                        textBox1.SelectionStart = s + "􏿽".Length;
                    }
                    textBox1.ScrollToCaret();
                    return;
                }

            }

            if (ModifierKeys == Keys.None)
            {//按下單一鍵
                if (e.KeyCode == Keys.Insert)
                {
                    if (insertMode)
                    {
                        insertMode = false;
                        Caret_Shown_OverwriteMode(textBox1);
                    }
                    else
                    {
                        insertMode = true;
                        Caret_Shown(textBox1);
                    }
                    return;
                }
                if (e.KeyCode == Keys.F1 || e.KeyCode == Keys.Pause)
                {//- 按下 F1 鍵：以找到的字串位置**前**分行分段
                 // -按下 Pause Break 鍵：以找到的字串位置** 後**分行分段
                    e.Handled = true;
                    splitLineParabySeltext(e.KeyCode);
                    return;
                }

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

                if (e.KeyCode == Keys.F4)
                {//F4 ： 重複做最後一次的輸入1次
                    int c = lastKeyPress.Count - 1;
                    if (c < 0) return;
                    string lk = lastKeyPress[c];
                    if (lk != "")
                    {
                        SendKeys.Send(lk);
                    }
                    return;
                }
            }

        }

        private void undoTextBox(TextBox textBox1)
        {
            int s = textBox1.SelectionStart, l = textBox1.SelectionLength;
            if (selStart != s && selStart != 0)
            {
                s = selStart; l = selLength;
            }
            if (undoTextBox1Text.Count - undoTimes > -1)
            {
                textBox1.Text = undoTextBox1Text[undoTextBox1Text.Count - ++undoTimes];
                restoreCaretPosition(textBox1, s, l);
            }
            else
                MessageBox.Show("no more to undo!");

        }

        private static void restoreCaretPosition(TextBox textBox1, int s, int l)
        {
            textBox1.Select(s, l);
            Point caretPosition = textBox1.GetPositionFromCharIndex(s);//c# caret position: https://stackoverflow.com/questions/37782986/how-to-find-the-caret-position-in-a-textbox-using-c
            if (caretPosition.Y > textBox1.Height - textBox1.Top || caretPosition.Y < textBox1.Top)
            {
                textBox1.ScrollToCaret();
                textBox1.AutoScrollOffset = caretPosition;
            }
        }

        private void insertWords(string insX, string x)
        {
            int s = textBox1.SelectionStart, l = textBox1.SelectionLength;
            x = x.Substring(0, s) + insX + x.Substring(s);
            textBox1.Text = x;
            textBox1.SelectionStart = s + insX.Length;
            textBox1.ScrollToCaret();
        }

        List<string> lastKeyPress = new List<string>();
        int findNotChineseCharFarLength(string x, bool forward)
        {
            StringInfo xInfo = new StringInfo(x);
            if (forward)
            {
                for (int i = 0; i < xInfo.LengthInTextElements; i++)
                {

                    if (!isChineseChar(xInfo.SubstringByTextElements(i, 1))) return i + 1;//https://www.jb51.net/article/45556.htm
                }
            }
            else
            {
                for (int i = xInfo.LengthInTextElements - 1; i > -1; i--)
                {

                    if (!isChineseChar(xInfo.SubstringByTextElements(i, 1))) return xInfo.LengthInTextElements - i;
                }

            }
            return -1;
        }

        int findChineseCharFarLength(string x, bool forward)
        {
            StringInfo xInfo = new StringInfo(x);
            if (forward)
            {
                for (int i = 0; i < xInfo.LengthInTextElements; i++)
                {

                    if (isChineseChar(xInfo.SubstringByTextElements(i, 1))) return i + 1;//https://www.jb51.net/article/45556.htm
                }
            }
            else
            {
                for (int i = xInfo.LengthInTextElements - 1; i > -1; i--)
                {

                    if (isChineseChar(xInfo.SubstringByTextElements(i, 1))) return xInfo.LengthInTextElements - i;
                }

            }
            return -1;
        }

        bool isChineseChar(string x)
        {
            if ("　􏿽\r\n<>{}。，；！？、－-《》〈〉「」『』0123456789".IndexOf(x) > -1) return false;

            //https://www.jb51.net/article/45556.htm
            //https://zh.wikipedia.org/wiki/%E4%B8%AD%E6%97%A5%E9%9F%93%E7%B5%B1%E4%B8%80%E8%A1%A8%E6%84%8F%E6%96%87%E5%AD%97
            if (Regex.IsMatch(x, @"[\u4e00-\u9fbb]")) return true;
            if (Regex.IsMatch(x, @"[\u3400-\u4dbf]")) return true;//擴充A區包含有6,592個漢字，位置在U+3400—U+4DBF
                                                                  //以下長度不同,恐怕失效,目前知C即不行,有空再測試            
                                                                  //c# 中文 轉 unicode:
                                                                  //https://www.google.com/search?q=c%23+%E4%B8%AD%E6%96%87+%E8%BD%89+unicode&rlz=1C1GCEU_zh-TWTW823TW823&sxsrf=AOaemvJI_o6pHrTEJVPCsVy0iyVsclLtjQ%3A1640527095825&ei=93TIYbnqMYOmoATnx4rwBg&oq=c%23++%5Cu%E4%B8%AD%E6%96%87%E5%AD%97%E7%A2%BC&gs_lcp=Cgdnd3Mtd2l6EAMYATIFCAAQzQIyBQgAEM0COggIABCwAxDNAjoECCMQJ0oECEEYAUoECEYYAFCzWFjiY2DfcGgCcAB4AIABVYgB1gGSAQEzmAEAoAEByAECwAEB&sclient=gws-wiz
                                                                  //https://www.itread01.com/p/1418585.html

            if (Regex.IsMatch(x, @"[\u20000-\u2A6DD]")) return true;//擴充B區包含有42,717個漢字，位置在U+20000—U+2A6DD
            if (Regex.IsMatch(x, @"[\u2A700-\u2B734]")) return true;//C:位置在U+2A700—U+2B734
            if (Regex.IsMatch(x, @"[\u2B740-\u2B81F]")) return true;//D:範圍為U+2B740–U+2B81F（實際有字元為U+2B740–U+2B81D）
            if (Regex.IsMatch(x, @"[\u2B820-\u2CEAF]")) return true;//E:編碼範圍U+2B820–U+2CEAF
            if (Regex.IsMatch(x, @"[\u2CEB0-\u2EBEF]")) return true;//F:U+2CEB0–U+2EBEF
            if (Regex.IsMatch(x, @"[\u30000-\u3134A]")) return true;//G:U+30000–U+3134A
                                                                    //if (Regex.IsMatch(x, @"[\u-\u]")) return true;//

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
            return false;
        }

        //C#中文字轉換Unicode(\u ):http://trufflepenne.blogspot.com/2013/03/cunicode.html
        private string StringToUnicode(string srcText)
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

        private string UnicodeToString(string srcText)
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


        private void keyDownCtrlAdd()
        {
            string x = textBox1.Text;
            int s = textBox1.SelectionStart, l = textBox1.SelectionLength;
            string xCopy = x.Substring(0, s + l);
            int[] chk = checkAbnormalLinePara(xCopy);
            if (chk.Length > 0)
            {
                if (MessageBox.Show("there is abnormal LinePara Length , check it now?" +
                    Environment.NewLine + Environment.NewLine +
                    "normal= " + chk[2] + "\tabnormal= " + chk[3], "",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button1) == DialogResult.OK)
                {
                    textBox1.Select(chk[0], chk[1]);
                    textBox1.ScrollToCaret();
                    return;
                }
            }
            newTextBox1();
            pasteToCtext();
            nextPages(Keys.PageDown, false);
        }

        const string omitStr = "【】〖〗＝{}<p>（）《》〈〉：；、，。「」『』？！　0123456789-‧·\r\n";
        string clearOmitChar(string x)
        {
            foreach (var item in omitStr)
            {
                x = x.Replace(item.ToString(), "");
            }
            return x;
        }

        int normalLineParaLength = 0;
        private int[] checkAbnormalLinePara(string xChk)
        {
            saveText();
            if (normalLineParaLength < 7) return new int[0];
            string[] xLineParas = xChk.Split(
                Environment.NewLine.ToArray(),
                StringSplitOptions.RemoveEmptyEntries);
            int i = -1, gap = 0, len = 0;
            foreach (string lineParaText in xLineParas)
            {
                i++;
                if (lineParaText.IndexOf("{{{") > -1)//{{{孫守真按：}}}等略去，以人工校對
                {
                    continue;
                }
                int noteTextBlendStart = lineParaText.IndexOf("{"),
                    noteTextBlendEnd = lineParaText.IndexOf("}");
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
                        if (noteTextBlendStart < noteTextBlendEnd)
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
                            len = new StringInfo(text).LengthInTextElements + (int)Math.Ceiling((decimal)new StringInfo(note).LengthInTextElements / 2);
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
                    gap = Math.Abs(len - normalLineParaLength);
                }
                else//only text or note
                {
                    len = new StringInfo(clearOmitChar(lineParaText)).
                        LengthInTextElements;
                    gap = Math.Abs(len - normalLineParaLength);
                }
                if (gap > 3 && !(len < normalLineParaLength
                    && lineParaText.IndexOf("<p>") > -1)
                    && lineParaText != "　" && lineParaText.IndexOf("*")==-1)//&& gap < 8)
                {//select the abnormal one
                    string x = textBox1.Text;
                    int j = -1, lineSeprtEnd = 0, lineSeprtStart = lineSeprtEnd;
                    lineSeprtEnd = x.IndexOf(Environment.NewLine, lineSeprtEnd);
                    while (lineSeprtEnd > -1)
                    {

                        if (++j == i) break;
                        lineSeprtStart = lineSeprtEnd;
                        lineSeprtEnd = x.IndexOf(Environment.NewLine, ++lineSeprtEnd);
                    }
                    return new int[] { lineSeprtStart, lineSeprtEnd - lineSeprtStart ,
                    normalLineParaLength,len};
                }
            }
            return new int[0];
            //throw new NotImplementedException();
        }

        void splitLineParabySeltext(Keys kys)
        {
            if (!(kys == Keys.F1 || kys == Keys.Pause)) return;
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
            textBox1.Text = x;
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
            if (x == "" || x.Length < 4) return;
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
            {//按下Ctrl鍵
                if (e.KeyCode == Keys.F)
                {
                    textBox2.Focus();
                    textBox2.SelectionStart = 0; textBox2.SelectionLength = textBox2.Text.Length; return;
                }

                if (e.KeyCode == Keys.PageDown || e.KeyCode == Keys.PageUp)
                {

                    e.Handled = true;//取得或設定值，指出是否處理事件。https://docs.microsoft.com/zh-tw/dotnet/api/system.windows.forms.keyeventargs.handled?view=netframework-4.7.2&f1url=%3FappId%3DDev16IDEF1%26l%3DZH-TW%26k%3Dk(System.Windows.Forms.KeyEventArgs.Handled);k(TargetFrameworkMoniker-.NETFramework,Version%253Dv4.7.2);k(DevLang-csharp)%26rd%3Dtrue
                    nextPages(e.KeyCode, true);
                    return;
                }
                if (e.KeyCode == Keys.Left || e.KeyCode == Keys.Right)
                {//Ctrl+左右鍵：徵調
                    if (textBox1.Focused) return;
                    Point mouseIP = Cursor.Position;//https://jjnnykimo.pixnet.net/blog/post/27155696
                    if (e.KeyCode == Keys.Left) { this.Left--; mouseIP.X--; }//-= w; }
                    if (e.KeyCode == Keys.Right) { this.Left++; mouseIP.X++; } //+= w; }
                    Cursor.Position = mouseIP;
                    //SetCursorPos(mouseIP.X,mouseIP.Y);
                    return;
                }

                if (e.KeyCode == Keys.D1)
                {
                    runWordMacro("漢籍電子文獻資料庫文本整理_以轉貼到中國哲學書電子化計劃");
                }
                if (e.KeyCode == Keys.D3)
                {
                    runWordMacro("漢籍電子文獻資料庫文本整理_十三經注疏");
                }
                if (e.KeyCode == Keys.D4)
                {
                    runWordMacro("維基文庫四部叢刊本轉來");
                }
                if (e.KeyCode == Keys.S)
                {
                    saveText();
                }
                return;

            }

            //if (((m & Keys.Control) == Keys.Control && (m & Keys.Alt) == Keys.Alt) && 
            //    (e.KeyCode == Keys.Left || e.KeyCode == Keys.Right||e.KeyCode==Keys.Menu))
            //{
            //    const int w= 10;
            //    if (e.KeyCode == Keys.Left) this.Left -= w;
            //    if (e.KeyCode == Keys.Right) this.Left += w;
            //    return;
            //}

            if (Control.ModifierKeys == Keys.Alt)
            {//按下Alt鍵
                if (e.KeyCode == Keys.Left || e.KeyCode == Keys.Right)
                {/*Alt + ←：視窗向左移動30dpi（+ Ctrl：徵調）
                  * Alt + →：視窗向右移動30dpi（+ Ctrl：徵調）*/
                    const int w = 30;
                    //int w = this.Width / 2;
                    if (e.KeyCode == Keys.Left) this.Left -= w;
                    if (e.KeyCode == Keys.Right) this.Left += w;
                    mouseMovein();
                    return;
                }

            }
            if (ModifierKeys == Keys.None)
            {//按下單一鍵            
                if (e.KeyCode == Keys.F5)
                {
                    loadText();
                    return;
                }
                if (e.KeyCode == Keys.F12)
                {
                    e.Handled = true;
                    BackupLastPageText(Clipboard.GetText(), true, true);
                    return;
                }
                if (e.KeyCode == Keys.Escape)
                {
                    hideToNICo();
                    //if (textBox1.Text == "")
                    ////預設為最上層顯示，若textBox1值為空，則按下Esc鍵會隱藏到任務列中；點一下即恢復
                    //{
                    //    hideToNICo();
                    //}
                }
            }
        }

        [DllImport("user32")]
        static extern bool SetCursorPos(int X, int Y);
        private void mouseMovein()
        {//https://lolikitty.pixnet.net/blog/post/164569578
            SetCursorPos(this.Left + 30, this.Top + 100);
        }

        void hideToNICo()
        {
            //https://dotblogs.com.tw/jimmyyu/2009/09/21/10733
            //https://dotblogs.com.tw/chou/2009/02/25/7284 https://yl9111524.pixnet.net/blog/post/49024854
            if (this.WindowState != FormWindowState.Minimized)
            {
                thisHeight = this.Height; thisWidth = this.Width; thisLeft = this.Left; thisTop = this.Top;
            }
            //this.WindowState = FormWindowState.Minimized;
            this.Hide();
            this.nICo.Visible = true;
        }
        const string fName_to_Backup_Txt = "cTextBK.txt";
        void BackupLastPageText(string x, bool updateLastBackup, bool showColorSignal)
        {
            Color C = this.BackColor;
            if (showColorSignal) { this.BackColor = Color.Red; Task.Delay(800).Wait(); }
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
                    if (showColorSignal) this.BackColor = C;
                    return;
                }
            }
            File.AppendAllText(dropBoxPathIncldBackSlash + fName_to_Backup_Txt, lastPageText, Encoding.UTF8);
            if (showColorSignal) this.BackColor = C;
        }

        private void nextPages(Keys eKeyCode, bool stayInHere)
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
            if (stayInHere) this.Activate();
        }

        private void runWordMacro(string runName)
        {
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
            appWord.Run(runName);
            textBox1.Text = Clipboard.GetText();
            textBox1.Select(0, 0);
            textBox1.ScrollToCaret();
            appWord.Quit(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);
            this.BackColor = C;
            show_nICo();
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
                button2.BackColor = button2BackColorDefault;
            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            textBox1Size();
        }

        void textBox1Size()
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
            undoRecord();
            if (button2.Text == "選取文")
            {
                replacedword = textBox2.Text;
                if (replacedword == "") return;
                l = textBox1.SelectionLength;
                string xBefore = x.Substring(0, s), xAfter = x.Substring(s + l);
                x = textBox1.SelectedText;
                if (rplsword == "\"\"") rplsword = "";//要清除所選文字，則選取其字，然後在 textBox4 輸入兩個英文半形雙引號 「""」（即表空字串），則不會取代成「""」，而是清除之。
                textBox1.Text = xBefore + x.Replace(replacedword, rplsword) + xAfter;
            }
            else
            {
                l = selWord.LengthInTextElements;
                if (rplsword == "\"\"") rplsword = "";
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
            this.Location = new Point
                (Width - this.Width, Height - this.Height - (int)(textBox1.Height / 3));
            textBox1Size();
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
                return;
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

            if (e.KeyCode == Keys.F1 || e.KeyCode == Keys.Pause)
            {
                e.Handled = true;
                splitLineParabySeltext(e.KeyCode);
                if (doNotLeaveTextBox2) textBox2.Focus();//方便快速分行分段
                return;
            }
            if (e.KeyCode == Keys.F2)
            {
                keyDownF2(textBox2);
                return;
            }
            if (e.KeyCode == Keys.F3)
            {
                KeyEventArgs ekey = new KeyEventArgs(Keys.F3);
                textBox1_KeyDown(textBox1, ekey);
            }
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
                    textBox.Select(0, textBox.Text.Length);
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
            //if ((m & Keys.Control) == Keys.Control && (m & Keys.Shift) == Keys.Shift)
            //{
            //    runWord("漢籍電子文獻資料庫文本整理_十三經注疏");
            //}
            //if ((m & Keys.Alt) == Keys.Alt)
            //{
            //    runWord("維基文庫四部叢刊本轉來");
            //}
            if (ModifierKeys == Keys.Control && e.Button == MouseButtons.Left)
            {
                richTextBox1.Size = textBox1.Size;
                richTextBox1.Location = textBox1.Location;
                richTextBox1.Show();

            }
            if (m == Keys.Alt && e.Button == MouseButtons.Left)
            {
                BackupLastPageText(Clipboard.GetText(), true, true);
                return;
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
        private Color textBox2BackColorDefault;

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            undoTextValueChanged(selStart, selLength);
            if (textBox1.Text == "" && !pasteAllOverWrite)
            {
                hideToNICo();
            }
        }

        List<string> undoTextBox1Text = new List<string>();
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

        private void Form1_Activated(object sender, EventArgs e)
        {
            if (!this.TopMost) this.TopMost = true;
            if (textBox1.Focused)
            {
                if (insertMode) Caret_Shown(textBox1); else Caret_Shown_OverwriteMode(textBox1);
                if (textBox1.SelectionLength == textBox1.Text.Length)
                    textBox1.Select(selStart, selLength);
            }
            if (textBox2.BackColor == Color.GreenYellow &&
                doNotLeaveTextBox2 && textBox2.Focused) textBox2.SelectAll();
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            if (insertMode) Caret_Shown(textBox1);
            else Caret_Shown_OverwriteMode(textBox1);
        }

        private void textBox2_Enter(object sender, EventArgs e)
        {
            textBox2.BackColor = Color.GreenYellow;
        }

        int surrogate = 0;
        private int undoTimes;

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
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            string x = textBox2.Text, x1 = textBox1.Text;
            if (x == "" || x1 == "") return;
            if (isKeyDownSurrogate(x)) return;//surrogate字在文字方塊輸入時會引發2次keyDown事件            
            var sa = findWord(x, x1);
            if (sa == null) return;
            int s = sa[0], nextS = sa[1];
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

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //https://social.msdn.microsoft.com/Forums/vstudio/en-US/5d021d76-36cd-43e6-b858-5a905c2e86d4/how-to-determine-if-in-insert-mode-or-overwrite-mode?forum=wpf
            //https://stackoverflow.com/questions/1428047/how-to-set-winforms-textbox-to-overwrite-mode/17962132#17962132
            //How can I place a TextBox in overwrite mode instead of insert mode:https://www.syncfusion.com/faq/windowsforms/textbox/how-can-i-place-a-textbox-in-overwrite-mode-instead-of-insert-mode
            if (!checkSurrogatePairsOK(e.KeyChar))
            {
                lastKeyPressElement += e.KeyChar;
                return;
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
            if (!insertMode)
            {//https://stackoverflow.com/questions/1428047/how-to-set-winforms-textbox-to-overwrite-mode/70502655#70502655
                if (textBox1.Text.Length != textBox1.MaxLength && textBox1.SelectedText == ""
                    && textBox1.Text != "" && textBox1.SelectionStart != textBox1.Text.Length)
                {
                    //string x = textBox1.Text; int s = textBox1.SelectionStart;
                    //    string xNext = x.Substring(s);
                    //    StringInfo xInfo = new StringInfo(xNext);                    
                    textBox1.SelectionLength = 1;//對於已經輸入完成的 surrogate C#應該會正確判斷其字長度；實際測試非然也
                    if (char.IsSurrogate(textBox1.SelectedText.ToCharArray()[0]))
                    {
                        textBox1.SelectionLength = 2;
                    }
                }
            }
            if (ModifierKeys == Keys.None)
            {
                undoRecord();
            }
        }

        private void undoRecord()
        {
            selStart = textBox1.SelectionStart; selLength = textBox1.SelectionLength;
            undoTextBox1Text.Add(textBox1.Text);
            if (undoTimes != 0) undoTimes = 0;
            if (undoTextBox1Text.Count > 50)//還原上限定為50個
            {
                undoTextBox1Text.RemoveAt(0);
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
        int indexOfStringInfo(string s, string x)
        {
            //StringInfo sInfo = new StringInfo(s);
            //StringInfo xInfo = new StringInfo(x);
            TextElementEnumerator xTE = StringInfo.GetTextElementEnumerator(x);
            int i = 0;
            while (xTE.MoveNext())
            {
                string sComp = xTE.Current.ToString();
                if (s == sComp)
                {
                    return i++;
                }
            }
            return -1;
        }
        private void Form1_Deactivate(object sender, EventArgs e)
        {//預設表單視窗為最上層顯示，當表單視窗不在作用中時，自動隱藏至系統右下方之系統列/任務列中，當滑鼠滑過任務列中的縮圖ico時，即還原/恢復視窗窗體
            if (!textBox2.Focused && textBox1.Text != "") this.TopMost = false;//hideToNICo();
            selStart = textBox1.SelectionStart; selLength = textBox1.SelectionLength;
        }
    }
}
