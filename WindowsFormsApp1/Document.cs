using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using WindowsFormsApp1;

/* 20250124年假前最後上班日 GitHub Copilot大菩薩：
 * 了解了您的需求，我們可以創建一個 Document 類別來表示 textBox1，並創建一個 Paragraph 類別來表示段落。這些類別將能夠操作 textBox1 中的內容，並確保更改段落的文字後，textBox1 的文字也能對應正確地更新。
    首先，我們創建一個 Document 類別，該類別將包含 textBox1 並提供操作其內容的方法。然後，我們創建一個 Paragraph 類別，該類別將能夠操作 Document 中的段落。
    ……
 * 這樣，我們就創建了 Document 和 Paragraph 類別。Document 類別表示 textBox1，並提供操作其內容的方法。Paragraph 類別表示單個段落，並能夠操作 Document 中的段落。
    您可以在 Form1 中創建 Document 和 Paragraph 對象，並使用它們來操作 textBox1 中的內容。例如：
    private void button1_Click(object sender, EventArgs e)
    {
        var paragraphs = _document.GetParagraphs();
        if (paragraphs.Count > 0)
        {
            paragraphs[0].Text = "新的段落文本";
        }
    }
 * 這樣，當您點擊按鈕時，第一個段落的文本將被更改，並且 textBox1 中的內容也會相應更新。希望這對您有所幫助。感恩感恩，南無阿彌陀佛。
 */
namespace TextForCtext
{
    /// <summary>
    /// 表示文檔（對應於 textBox1）
    /// </summary>
    public class Document : IDisposable
    {
        private TextBox _textBox;
        private int currentParagraphIndex;
        private List<Paragraph> _cachedParagraphs; // 緩存段落集合

        public Document(ref TextBox textBox)
        {
            _textBox = textBox;
            _cachedParagraphs = null;
        }
        public Document(ref string context)
        {
            //if (context.IndexOf("\v") > -1) context.Replace("\v", Environment.NewLine);
            _textBox = new TextBox();
            Text = context;
            _cachedParagraphs = null;
        }

        public string Text
        {
            get => _textBox.Text;
            set
            {
                _textBox.Text = value;
                _cachedParagraphs = null; // 文本改變時清空緩存
                //_textBox.Refresh(); // 更新文本框
            }
        }

        public int SelectionStart
        {
            get => _textBox.SelectionStart;
            set => _textBox.SelectionStart = value;
        }

        public int SelectionLength
        {
            get => _textBox.SelectionLength;
            set => _textBox.SelectionLength = value;
        }
        internal int CurrentParagraphIndex { get => currentParagraphIndex; set => currentParagraphIndex = value; }

        public void InsertText(string text)
        {
            int selectionStart = _textBox.SelectionStart;
            _textBox.Text = _textBox.Text.Insert(selectionStart, text);
            _textBox.SelectionStart = selectionStart + text.Length;
            _cachedParagraphs = null; // 文本改變時清空緩存
        }

        public void ReplaceText(string text)
        {
            int selectionStart = _textBox.SelectionStart;
            _textBox.Text = _textBox.Text.Remove(selectionStart, _textBox.SelectionLength);
            _textBox.Text = _textBox.Text.Insert(selectionStart, text);
            _textBox.SelectionStart = selectionStart + text.Length;
            _cachedParagraphs = null; // 文本改變時清空緩存
        }
        /// <summary>
        /// 取從文件（textBox1內容的段落集合）。20250220補充說明
        /// </summary>
        /// <returns>傳回null則有問題，須檢查！</returns>
        public List<Paragraph> GetParagraphs(bool newOne = false)
        {
            // 檢查並更新緩存
            if (_cachedParagraphs == null || newOne || !IsTextEqualToCachedParagraphs())
            {
                var paragraphs = new List<Paragraph>();
                var lines = _textBox.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                int start = 0;
                for (int i = 0; i < lines.Length; i++)
                {
                    var line = lines[i];
                    start = this.Text.IndexOf(line, start);

                    if (start == -1)
                    {
                        Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("有問題，請檢查！");
                        return null;
                    }

                    paragraphs.Add(new Paragraph(line, this, start, i));

                    start += (line.Length == 0 ? 1 : line.Length);
                    start = start > this.Text.Length ? this.Text.Length : start;
                }

                _cachedParagraphs = paragraphs; // 更新緩存
            }

            return _cachedParagraphs;
        }

        private bool IsTextEqualToCachedParagraphs()
        {
            var lines = _textBox.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            if (lines.Length != _cachedParagraphs.Count)
            {
                return false;
            }

            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i].Length != _cachedParagraphs[i].Text.Length
                    || !lines[i].Equals(_cachedParagraphs[i].Text))
                {
                    return false;
                }
            }

            return true;
        }


        //public List<Paragraph> GetParagraphs(Document document)
        //{
        //    var paragraphs = new List<Paragraph>();
        //    var lines = this.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

        //    foreach (var line in lines)
        //    {
        //        paragraphs.Add(new Paragraph(line, this));
        //    }

        //    return paragraphs;
        //}

        #region 取得插入點所在位置的段落，以及其上個段落和下個段落。
        /// <summary>
        /// 取得插入點所在位置段落
        /// </summary>
        /// <returns></returns>
        public Paragraph GetCurrentParagraph()
        {
            var paragraphs = GetParagraphs();
            int caretPosition = _textBox.SelectionStart;
            int charCount = 0;

            for (int i = 0; i < paragraphs.Count; i++)
            {
                charCount += paragraphs[i].Text.Length + Environment.NewLine.Length;
                if (caretPosition < charCount)
                {
                    return paragraphs[i];
                }
            }

            return null;
        }
        /// <summary>
        /// 取得插入點所在位置後一個段落
        /// </summary>
        /// <returns></returns>
        public Paragraph GetNextParagraph()
        {
            var paragraphs = GetParagraphs();
            int caretPosition = _textBox.SelectionStart;
            int charCount = 0;

            for (int i = 0; i < paragraphs.Count; i++)
            {
                charCount += paragraphs[i].Text.Length + Environment.NewLine.Length;
                if (caretPosition < charCount && i + 1 < paragraphs.Count)
                {
                    return paragraphs[i + 1];
                }
            }

            return null;
        }
        /// <summary>
        /// 取得插入點所在位置前一個段落
        /// </summary>
        /// <returns></returns>
        public Paragraph GetPreviousParagraph()
        {
            var paragraphs = GetParagraphs();
            int caretPosition = _textBox.SelectionStart;
            int charCount = 0;

            for (int i = 0; i < paragraphs.Count; i++)
            {
                charCount += paragraphs[i].Text.Length + Environment.NewLine.Length;
                if (caretPosition < charCount && i - 1 >= 0)
                {
                    return paragraphs[i - 1];
                }
            }

            return null;
        }
        #endregion


        /// <summary>
        /// 
        /// </summary>
        /// <param name="paragraphIndex"></param>
        /// <param name="newChar"></param>
        public void UpdateParagraphFirstCharacter(int paragraphIndex, string newChar)
        {/* 好的，我們可以根據您的需求來編寫測試代碼。以下是實現這些步驟的代碼：
            1.	取得 textBox1 內容的所有段落。
            2.	遍歷各個段落，找出符合條件的段落。
            3.	更改符合條件的段落的下一個段落的第一個字元。
            4.	確保更改反映在 textBox1 的內容中。
            首先，我們需要在 Document 類別中添加一個方法來更新段落的第一個字元：
          */

            var paragraphs = GetParagraphs();
            if (paragraphs.Count > paragraphIndex)
            {
                var paragraph = paragraphs[paragraphIndex];
                if (paragraph.Text.Length > 0)
                {
                    if (char.IsHighSurrogate(paragraph.Text[0]) && paragraph.Text.Length > 1)
                    {
                        paragraph.Text = newChar + paragraph.Text.Substring(2);
                    }
                    else
                    {
                        paragraph.Text = newChar + paragraph.Text.Substring(1);
                    }
                }
            }
        }
        /// <summary>
        /// 將插入點後的2行/段內容，改成夾注語法並接在插入點本行後（類似按下Ctrl + Shift + F1） 20250131大年初三
        /// creedit with GitHub Copilot大菩薩：Alt + w
        /// 加速OCR夾注文本的排版整理
        /// </summary>
        /// <exception cref="InvalidOperationException"></exception>
        public void MergeParagraphsAtCaret()
        {
            var paragraphs = GetParagraphs();
            int caretPosition = _textBox.SelectionStart;
            int currentParagraphIndex = 0;
            int charCount = 0;
            int s = _textBox.SelectionStart;

            // 找到插入點所在的段落
            for (int i = 0; i < paragraphs.Count; i++)
            {
                charCount += paragraphs[i].Text.Length + Environment.NewLine.Length;
                if (caretPosition < charCount)
                {
                    currentParagraphIndex = i;
                    break;
                }
            }

            // 確保有足夠的段落進行操作
            if (currentParagraphIndex + 2 >= paragraphs.Count)
            {
                throw new InvalidOperationException("沒有足夠的段落進行操作。");
            }

            // 取得插入點所在段落及其後的兩個段落
            var currentParagraph = paragraphs[currentParagraphIndex];
            var nextParagraph1 = paragraphs[currentParagraphIndex + 1];
            var nextParagraph2 = paragraphs[currentParagraphIndex + 2];

            // 將這兩個段落前後分別加上「{{」和「}}」，並清除其中的分段符號
            string mergedText = "{{" + nextParagraph1.Text.Replace(Environment.NewLine, "") + nextParagraph2.Text.Replace(Environment.NewLine, "") + "}}";

            // 將這兩個段落的內容併到插入點所在段落的後面
            currentParagraph.Text += mergedText;

            // 刪除原來的兩個段落
            paragraphs.RemoveAt(currentParagraphIndex + 2);
            paragraphs.RemoveAt(currentParagraphIndex + 1);

            // 更新文本框的內容
            _textBox.Text = string.Join(Environment.NewLine, paragraphs.Select(p => p.Text));

            // 更新插入點的位置到合併後的段落的末尾
            //_textBox.SelectionStart += currentParagraph.Text.Length;
            if (_textBox.Text.IndexOf(Environment.NewLine, s) > -1) s = _textBox.Text.IndexOf(Environment.NewLine, s);
            else s = _textBox.TextLength;

            // 合併插入點後的一個段落，但保持插入點在當前位置不變
            if (currentParagraphIndex + 1 < paragraphs.Count)

            {

                var nextParagraph = paragraphs[currentParagraphIndex + 1];

                currentParagraph.Text += nextParagraph.Text.Replace(Environment.NewLine, "");

                paragraphs.RemoveAt(currentParagraphIndex + 1);

                _textBox.Text = string.Join(Environment.NewLine, paragraphs.Select(p => p.Text));

            }
            _textBox.SelectionStart = s;
            _textBox.ScrollToCaret();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <exception cref="InvalidOperationException"></exception>
        public void MergeParagraphsAtCaretWithShift()
        {/* 20250131大年初三：GitHub Copilot大菩薩：
          * 我們可以在 Document 類別中添加一個新的方法 MergeParagraphsAtCaretWithShift 來實現 Alt + Shift + w 的功能。這個方法將插入點所在行視為夾注的第1行，將其後的1行/段合併上來，並在合併後的前後分別加上「{{」和「}}」，然後再將它們後面的那1行也合併上來。最後，插入點將停留在最後合併上來的那行/段文字的起始處。
          * ……
          * 這樣，當您按下 Alt + Shift + w 時，將執行 MergeParagraphsAtCaretWithShift 方法，插入點所在行將視為夾注的第1行，並將其後的1行/段合併上來，前後分別加上「{{」和「}}」，然後再將它們後面的那1行也合併上來。最後，插入點將停留在最後合併上來的那行/段文字的起始處。
          */
            var paragraphs = GetParagraphs();
            int caretPosition = _textBox.SelectionStart;
            int currentParagraphIndex = 0;
            int charCount = 0;
            int s = _textBox.SelectionStart;

            // 找到插入點所在的段落
            for (int i = 0; i < paragraphs.Count; i++)
            {
                charCount += paragraphs[i].Text.Length + Environment.NewLine.Length;
                if (caretPosition < charCount)
                {
                    currentParagraphIndex = i;
                    break;
                }
            }

            // 確保有足夠的段落進行操作
            if (currentParagraphIndex + 2 >= paragraphs.Count)
            {
                throw new InvalidOperationException("沒有足夠的段落進行操作。");
            }

            // 取得插入點所在段落及其後的兩個段落
            var currentParagraph = paragraphs[currentParagraphIndex];
            var nextParagraph1 = paragraphs[currentParagraphIndex + 1];
            var nextParagraph2 = paragraphs[currentParagraphIndex + 2];

            // 將這兩個段落前後分別加上「{{」和「}}」，並清除其中的分段符號
            string mergedText = "{{" + currentParagraph.Text.Replace(Environment.NewLine, "") + nextParagraph1.Text.Replace(Environment.NewLine, "") + "}}";

            // 將這兩個段落的內容併到插入點所在段落的後面
            currentParagraph.Text = mergedText;

            // 刪除原來的兩個段落
            paragraphs.RemoveAt(currentParagraphIndex + 1);
            //paragraphs.RemoveAt(currentParagraphIndex + 1);
            //因為　您給我的程式碼在前面已經清除過分段符號了，再移除一次，就會把後面接的行/段給誤刪了。

            // 更新文本框的內容
            _textBox.Text = string.Join(Environment.NewLine, paragraphs.Select(p => p.Text));

            // 合併插入點後的一個段落，但保持插入點在當前位置不變
            if (currentParagraphIndex + 1 < paragraphs.Count)
            {
                var nextParagraph = paragraphs[currentParagraphIndex + 1];
                currentParagraph.Text += nextParagraph.Text.Replace(Environment.NewLine, "");
                paragraphs.RemoveAt(currentParagraphIndex + 1);
                _textBox.Text = string.Join(Environment.NewLine, paragraphs.Select(p => p.Text));
            }

            // 更新插入點的位置到合併後的段落的起始處
            _textBox.SelectionStart = _textBox.Text.IndexOf("}}", s) + 2;
            _textBox.ScrollToCaret();
        }




        /* 20250203大年初六立春 creedit with GitHub Copilot大菩薩：
                  * 好的，我們可以根據您的需求來實現這個自動標題標記的函式方法。以下是具體實現：
                    1.	檢查每個行/段，根據全形空格數量來判斷是否為標題。
                    2.	根據不同的空格數量來添加標記。
                    3.	處理標題行/段之後的行/段，直到遇到非空格開頭的行/段。
                    以下是具體的實現代碼：
                  */

        /// <summary>
        /// 這個方法會根據指定的全形空格數量來自動標記標題，並在適當的位置添加標題標記的開頭和結束標記。希望這對您有所幫助。感恩感恩，南無阿彌陀佛。
        /// </summary>
        /// <param name="titleIndentationSpaces"></param>
        public void AutoMarkTitles(int titleIndentationSpaces)
        {
            var paragraphs = GetParagraphs();
            var sb = new StringBuilder();
            bool inTitle = false;

            for (int i = 0; i < paragraphs.Count; i++)
            {
                var paragraph = paragraphs[i];
                var text = paragraph.Text;
                int leadingSpaces = text.TakeWhile(c => c == '　').Count();

                if (leadingSpaces == titleIndentationSpaces)
                {
                    // 標題行/段
                    sb.Append(text.Insert(leadingSpaces, "*"));
                    for (int j = leadingSpaces + 1; j < text.Length; j++)
                    {
                        if (text[j] == '　')
                        {
                            sb.Append("􏿽");
                        }
                        else
                        {
                            sb.Append(text[j]);
                        }
                    }
                    inTitle = true;
                }
                else if (leadingSpaces < titleIndentationSpaces)
                {
                    // 非標題行/段，空格數少於指定數量
                    for (int j = 0; j < text.Length; j++)
                    {
                        if (text[j] == '　')
                        {
                            sb.Append("􏿽");
                        }
                        else
                        {
                            sb.Append(text[j]);
                        }
                    }
                    sb.Append("<p>");
                    inTitle = false;
                }
                else
                {
                    // 非標題行/段，空格數多於指定數量
                    sb.Append(text.Insert(leadingSpaces, "**"));
                    for (int j = leadingSpaces + 2; j < text.Length; j++)
                    {
                        if (text[j] == '　')
                        {
                            sb.Append("􏿽");
                        }
                        else
                        {
                            sb.Append(text[j]);
                        }
                    }
                    inTitle = true;
                }

                if (inTitle && (i + 1 >= paragraphs.Count || !paragraphs[i + 1].Text.StartsWith("　")))
                {
                    sb.Append("<p>");
                    inTitle = false;
                }

                sb.Append(Environment.NewLine);
            }

            _textBox.Text = sb.ToString();
        }


        /// <summary>
        /// 取得指定範圍的 Range 物件
        /// </summary>
        /// <param name="start">範圍的起始位置</param>
        /// <param name="end">範圍的結束位置</param>
        /// <returns>Range 物件</returns>
        public Range Range(int start, int end)
        {
            return new Range(this, start, end);
            //return new Range(ref _textBox, start, end);
        }

        public void Dispose()
        {
            _textBox?.Dispose();
        }
        /// <summary>
        /// 取得指定索引前後的段落集合
        /// </summary>
        /// <param name="index">指定的索引</param>
        /// <param name="range">前後段落的範圍</param>
        /// <returns>前後段落的集合</returns>
        public List<Paragraph> GetSurroundingParagraphs(int index, int range = 3)
        {
            // 檢查並更新緩存
            if (_cachedParagraphs == null || _textBox.Text != string.Join(Environment.NewLine, _cachedParagraphs.Select(p => p.Text)))
            {
                _cachedParagraphs = GetParagraphs();
            }

            var paragraphs = _cachedParagraphs;
            if (paragraphs == null || index < 0 || index >= paragraphs.Count)
            {
                return null;
            }

            int start = Math.Max(0, index - range);
            int end = Math.Min(paragraphs.Count - 1, index + range);

            return paragraphs.GetRange(start, end - start + 1);
        }

        /// <summary>
        /// 檢查 textBox1 的內容是否有異，並重新取得段落集合
        /// </summary>
        /// <returns>是否重新取得段落集合</returns>
        public bool CheckAndRefreshParagraphs()
        {
            if (_cachedParagraphs == null || _textBox.Text != string.Join(Environment.NewLine, _cachedParagraphs.Select(p => p.Text)))
            {
                _cachedParagraphs = GetParagraphs();
                return true;
            }
            return false;
        }



    }


    /// <summary>
    /// 表示單個段落
    /// </summary>
    public class Paragraph
    {
        private string _text;
        private readonly string _text_beforeUpdate;
        private Document _document;
        private int _start;
        private readonly int _start_beforeUpdate;
        private int _end;
        private Range _range;
        private int _index; // 新增的屬性

        public Paragraph(string text, Document document, int start, int index)
        {
            _text = text;
            _text_beforeUpdate = _text;
            _document = document;
            _start = start;
            _start_beforeUpdate = _start;
            _end = _start + _text.Length;
            _index = index; // 初始化屬性
        }

        public string Text
        {
            get => _text;
            set
            {
                _text = value;
                _end = _start + _text.Length;
                UpdateDocument();
            }
        }

        public string TextBeforeUpdate => _text_beforeUpdate;


        public int Start
        {
            get => _start;
            set
            {
                _start = value;
                _end = _start + _text.Length;
                UpdateDocument();
            }
        }

        public int StartBeforeUpdate => _start_beforeUpdate;


        public int End
        {
            get => _end;
            set
            {
                _end = value;
                _text = _document.Text.Substring(_start, _end - _start);
                UpdateDocument();
            }
        }

        //public int EndBeforeUpdate
        //{
        //    get => _start_beforeUpdate + _text_beforeUpdate.Length;
        //}
        public int EndBeforeUpdate => _start_beforeUpdate + _text_beforeUpdate.Length;

        public int Index // 新增的屬性
        {
            get => _index;
            set => _index = value;
        }


        private void UpdateDocument()
        {
            var paragraphs = _document.GetParagraphs();
            int index = _index; // 使用 Index 屬性

            if (index < 0 || index >= paragraphs.Count)
            {
                Form1.MessageBoxShowOKExclamationDefaultDesktopOnly("找不到要更新段落"
                    + Environment.NewLine + Environment.NewLine
                    + this.Text + Environment.NewLine + Environment.NewLine
                    + this._text_beforeUpdate);
                return;
            }

            var lines = _document.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            if (index < lines.Length)
            {
                lines[index] = _text;
                _document.Text = string.Join(Environment.NewLine, lines);
            }
            else
            {
                throw new IndexOutOfRangeException("索引超出 lines 陣列的界限。");
            }
        }

        public Range Range
        {
            get
            {
                if (_range == null)
                {
                    _range = new Range(_document, _start, _end);
                    return _range;
                }
                return _range;
            }
        }
    }



    /* 20250204 GitHub　Copilot大菩薩：
      * 這樣，我們就實現了在 Document 類別中取得插入點所在位置的段落，以及其上個段落和下個段落的方法，並且創建了一個 Range 類別來模擬 MS Word VBA 的 Range 物件。希望這對您有所幫助。感恩感恩，南無阿彌陀佛。
      */

    /// <summary>
    /// 創建一個 Range 類別來模擬 MS Word VBA 的 Range 物件。
    /// </summary>
    public class Range
    {
        private Document _document;
        private int _start;
        private int _end;
        List<Paragraph> rangeParagraphs;
        //TextBox _textBox;

        public Range(Document document, int start, int end)
        //public Range(ref TextBox textBox, int start, int end)
        {
            _document = document;
            //_document = new Document(ref textBox);
            _start = start;
            _end = end;
            //Text = _document.Text.Substring(start, end - start);

            rangeParagraphs = Paragraphs;
            //_textBox = textBox;
        }

        public string Text
        {
            get => _document.Text.Substring(_start, _end - _start);
            set
            {
                var text = _document.Text;
                _document.Text = text.Substring(0, _start) + value + text.Substring(_end);
                _end = _start + value.Length;
                UpdateDocument();
                //GitHub　Copilot大菩薩：我們在設置 Text 屬性時調用 UpdateDocument 方法，以確保文檔內容的更新。
            }
        }
        /// <summary>
        /// 在 Range 類別中添加一個 UpdateDocument 方法，並在設置 Text 屬性時調用它。以下是具體實現：
        /// </summary>
        private void UpdateDocument()
        {
            var paragraphs = _document.GetParagraphs();
            int charCount = 0;

            for (int i = 0; i < paragraphs.Count; i++)
            {
                charCount += paragraphs[i].Text.Length + Environment.NewLine.Length;
                if (_start < charCount)
                {
                    if (paragraphs[i].Text != _document.Text.Substring(charCount - paragraphs[i].Text.Length - Environment.NewLine.Length, paragraphs[i].Text.Length))
                        paragraphs[i].Text = _document.Text.Substring(charCount - paragraphs[i].Text.Length - Environment.NewLine.Length, paragraphs[i].Text.Length);//●●●●●●●●●●●●●●20250213
                }
            }
        }
        /*
         GitHub　Copilot大菩薩：新增 Start 和 End 屬性，這些屬性將允許讀取和設置範圍的起始和結束位置。以下是具體實現：20250204
         */
        /// <summary>
        /// 就新增了 Start 和 End 屬性，允許讀取和設置範圍的起始和結束位置。希望這對您有所幫助。感恩感恩，南無阿彌陀佛。
        /// </summary>
        public int Start
        {
            get => _start;
            set => _start = value;
        }
        /// <summary>
        /// 就新增了 Start 和 End 屬性，允許讀取和設置範圍的起始和結束位置。希望這對您有所幫助。感恩感恩，南無阿彌陀佛。
        /// </summary>
        public int End
        {
            get => _end;
            set => _end = value;
        }
        public void Select()
        {
            _document.SelectionStart = _start;
            _document.SelectionLength = _end - _start;
        }

        public Paragraph GetCurrentParagraph()
        {
            var paragraphs = _document.GetParagraphs();
            int charCount = 0;

            for (int i = 0; i < paragraphs.Count; i++)
            {
                charCount += paragraphs[i].Text.Length + Environment.NewLine.Length;
                if (_start < charCount)
                {
                    return paragraphs[i];
                }
            }

            return null;
        }

        public Paragraph GetNextParagraph()
        {
            var paragraphs = _document.GetParagraphs();
            int charCount = 0;

            for (int i = 0; i < paragraphs.Count; i++)
            {
                charCount += paragraphs[i].Text.Length + Environment.NewLine.Length;
                if (_start < charCount && i + 1 < paragraphs.Count)
                {
                    return paragraphs[i + 1];
                }
            }

            return null;
        }

        public Paragraph GetPreviousParagraph()
        {
            var paragraphs = _document.GetParagraphs();
            int charCount = 0;

            for (int i = 0; i < paragraphs.Count; i++)
            {
                charCount += paragraphs[i].Text.Length + Environment.NewLine.Length;
                if (_start < charCount && i - 1 >= 0)
                {
                    return paragraphs[i - 1];
                }
            }

            return null;
        }

        public List<Paragraph> Paragraphs
        {
            get
            {
                if (rangeParagraphs == null || rangeParagraphs.Count == 0)
                {
                    rangeParagraphs = new List<Paragraph>();
                    var paragraphs = _document.GetParagraphs();

                    int charCount = 0;
                    //GitHub　Copilot大菩薩：這樣，Paragraphs 屬性將始終返回最新的段落集合。希望這對您有所幫助。感恩感恩，南無阿彌陀佛。 20250204
                    //故不要改為欄位儲存值，以免造成段落集合不是最新的問題。
                    foreach (var paragraph in paragraphs)
                    {
                        charCount += paragraph.Text.Length + Environment.NewLine.Length;
                        if (charCount > _start && charCount - paragraph.Text.Length - Environment.NewLine.Length < _end)
                        {
                            rangeParagraphs.Add(paragraph);
                        }
                    }
                }
                return rangeParagraphs;
            }
        }
        //public List<Paragraph> Paragraphs
        //        {
        //            get
        //            {
        //                var paragraphs = _document.GetParagraphs();
        //                var rangeParagraphs = new List<Paragraph>();
        //                int charCount = 0;
        //                //GitHub　Copilot大菩薩：這樣，Paragraphs 屬性將始終返回最新的段落集合。希望這對您有所幫助。感恩感恩，南無阿彌陀佛。 20250204
        //                //故不要改為欄位儲存值，以免造成段落集合不是最新的問題。
        //                foreach (var paragraph in paragraphs)
        //                {
        //                    charCount += paragraph.Text.Length + Environment.NewLine.Length;
        //                    if (charCount > _start && charCount - paragraph.Text.Length - Environment.NewLine.Length < _end)
        //                    {
        //                        rangeParagraphs.Add(paragraph);
        //                    }
        //                }

        //                return rangeParagraphs;
        //            }
        //        }
    }
}
